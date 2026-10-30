import hashlib
import json
import logging
import mimetypes
import math
import sqlite3
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

from core import autofix, batch, bl_links, debug_log, fix_images, mark_finder, scan_analysis, link_repair, repair_compare, hidden_repair, pdf_review
from core.common import (
    STATIC_DIR,
    console_print,
    find_backup_sheets,
    sanitize_filename,
    workbook_rel_name,
)
from core.excel_ops import extract_sheet_snapshot, match_excel_row, normalize_workbook_view, open_workbook_in_excel
from core.pdf_detector import (
    extract_page_lines,
    find_clipped_spans_on_page,
    is_region_still_clipped,
    render_crop_base64,
    render_page_overview,
)
from core.reporter import (
    EXECUTION_STATE,
    find_recorded_workbook,
    find_workbook,
    get_execution_report_snapshot,
    render_execution_markdown,
)
from core.report_export import build_report_zip
from core.sheet_preview import preview_image_path, render_workbook_previews, workbook_preview_info
from core.workflow import (
    QUICK_TASK_NAMES,
    STATE,
    STATE_LOCK,
    apply_skip_decisions,
    apply_keep_decisions,
    export_progress,
    fix_selected_workflow,
    get_overview_snapshot,
    get_state_snapshot,
    print_layout_snapshot,
    set_print_layout,
    state_image,
    import_progress,
    list_workbooks,
    log_msg,
    preview_backup_sheets,
    scan_folder_workflow,
    queue_job,
    resync_report_from_state,
    run_pending_workflow,
    set_manual_confirmation,
    update_item_row,
    validate_clean_files,
    validate_clean_plan,
)

__all__ = [
    "STATE",
    "STATE_LOCK",
    "WebHandler",
    "extract_page_lines",
    "extract_sheet_snapshot",
    "find_clipped_spans_on_page",
    "fix_selected_workflow",
    "find_backup_sheets",
    "is_region_still_clipped",
    "log_msg",
    "main",
    "match_excel_row",
    "normalize_workbook_view",
    "render_crop_base64",
    "render_page_overview",
    "sanitize_filename",
    "scan_folder_workflow",
]

MIME_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".pdf": "application/pdf",
}


def _require_folder(payload: dict) -> str:
    folder = payload.get("folder")
    if not isinstance(folder, str) or not folder.strip():
        raise ValueError("请输入 Excel 文件夹路径")
    return folder


class WebHandler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: object) -> None:
        del format, args  # 保持终端清爽

    def _send_bytes(
        self, body: bytes, content_type: str, status: int = 200, filename: str | None = None,
        cacheable: bool = False, etag: str = "",
    ) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        if status != 304:
            self.send_header("Content-Length", str(len(body)))
        # 带版本号的截图地址内容不会变，可以缓存；其余一律不缓存
        self.send_header("Cache-Control", "private, max-age=86400" if cacheable else "no-store")
        if etag:
            self.send_header("ETag", etag)
        if filename is not None:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        if status != 304:
            self.wfile.write(body)

    def _send_json(self, data: dict, status: int = 200) -> None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self._send_bytes(body, "application/json; charset=utf-8", status=status)

    def _serve_static_file(self, rel_path: str) -> None:
        static_root = STATIC_DIR.resolve()
        target = (static_root / rel_path).resolve()
        if not target.is_file() or static_root not in target.parents:
            self.send_error(404)
            return

        content_type = MIME_TYPES.get(
            target.suffix.lower(),
            mimetypes.guess_type(str(target))[0] or "application/octet-stream",
        )
        self._send_bytes(target.read_bytes(), content_type)

    def do_GET(self) -> None:
        try:
            self._handle_get()
        except ConnectionError:
            raise
        except Exception as err:
            # 没预料到的错误：完整 traceback 写进调试日志，网页上给一句说明
            debug_log.exception(err, f"GET {self.path}")
            self._send_json({"ok": False, "error": f"程序出错：{err}（详细信息已写进调试日志）"}, status=500)

    def _handle_get(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path in {"/", "/home.html"}:
            self._serve_static_file("home.html")
        elif parsed.path in {"/clean", "/clean.html"}:
            self._serve_static_file("clean.html")
        elif parsed.path in {"/clean/preview", "/clean_preview.html"}:
            self._serve_static_file("clean_preview.html")
        elif parsed.path.startswith("/api/clean/thumb/"):
            parts = parsed.path.removeprefix("/api/clean/thumb/").split("/")
            image = preview_image_path(*parts) if len(parts) == 2 else None
            if image is None:
                self.send_error(404)
                return
            self._send_bytes(image.read_bytes(), MIME_TYPES[image.suffix])
        elif parsed.path in {"/workbench", "/index.html"}:
            self._serve_static_file("batch.html")
        elif parsed.path == "/repair-links":
            self._serve_static_file("link_repair.html")
        elif parsed.path == "/repair-colors":
            self._serve_static_file("color_repair.html")
        elif parsed.path == "/repair-hidden":
            self._serve_static_file("hidden_repair.html")
        elif parsed.path == "/pdf-review":
            self._serve_static_file("pdf_review.html")
        elif parsed.path == "/api/pdf_review":
            folder = parse_qs(parsed.query).get("folder", [""])[0]
            try:
                self._send_json({"ok": True, **pdf_review.list_pdfs(folder)})
            except ValueError as err:
                self._send_json({"ok": False, "error": str(err)}, status=400)
        elif parsed.path == "/api/pdf_review/rects":
            self._send_json({"ok": True, **pdf_review.rects_snapshot()})
        elif parsed.path == "/api/pdf_review/page":
            query = parse_qs(parsed.query)
            try:
                image = pdf_review.page_png(query.get("pdf", [""])[0], int(query.get("p", ["0"])[0]))
            except ValueError:
                self.send_error(404)
                return
            # 地址带 PDF 的修改时间（?v=），PDF 重新导出后地址就变，可以缓存
            self._send_bytes(image, "image/png", cacheable=True)
        elif parsed.path == "/repair-bl-links":
            self._serve_static_file("bl_links.html")
        elif parsed.path == "/api/bl_links":
            self._send_json(bl_links.status())
        elif parsed.path == "/api/hidden_repair":
            self._send_json(hidden_repair.status())
        elif parsed.path == "/api/link_repair":
            self._send_json(link_repair.status())
        elif parsed.path == "/analysis":
            self._serve_static_file("analysis.html")
        elif parsed.path in {"/api/scan-analysis", "/api/scan-analysis/markdown"}:
            query = parse_qs(parsed.query)
            folder = query.get("folder", [None])[0]
            run_id = query.get("id", [None])[0]
            data = scan_analysis.report(run_id, folder)
            if parsed.path.endswith("/markdown"):
                if data is None:
                    self._send_json({"ok": False, "error": "找不到扫描分析报告"}, status=404)
                else:
                    self._send_bytes(scan_analysis.markdown(data).encode("utf-8"), "text/markdown; charset=utf-8",
                                     filename="scan_analysis.md")
            elif run_id and data is None:
                self._send_json({"ok": False, "error": "找不到扫描分析报告"}, status=404)
            else:
                self._send_json({"ok": True, "reports": scan_analysis.history(folder), "report": data})
        elif parsed.path in {"/records", "/records.html"}:
            self._serve_static_file("records.html")
        elif parsed.path == "/favicon.ico":
            self._serve_static_file("favicon.svg")
        elif parsed.path.startswith("/static/"):
            rel = parsed.path.removeprefix("/static/")
            self._serve_static_file(rel)
        elif parsed.path == "/api/print_layouts":
            folder = parse_qs(parsed.query).get("folder", [None])[0]
            with STATE_LOCK:
                job = STATE.get("job") or {}
                blocked = ("任务进行中，完成后可调整方向。" if STATE["busy"] else
                           "请先结算未完成的写入，再调整方向。" if job.get("pending_adjustment") or job.get("deferred_adjustments") else
                           "请先完成暂停任务，再调整方向。" if job.get("status") == "interrupted" else "")
                result = {"rows": print_layout_snapshot(folder), "busy": STATE["busy"], "blocked": blocked}
            self._send_json(result)
        elif parsed.path == "/api/batch":
            query = parse_qs(parsed.query)
            result = batch.snapshot(set(query.get("file", []))) if query.get("details") == ["lazy"] else batch.snapshot()
            body = json.dumps(result, ensure_ascii=False).encode("utf-8")
            etag = f'"{hashlib.sha256(body).hexdigest()}"'
            validators = {tag.strip().removeprefix("W/") for tag in self.headers.get("If-None-Match", "").split(",")}
            unchanged = etag in validators or "*" in validators
            self._send_bytes(b"" if unchanged else body, "application/json; charset=utf-8",
                             status=304 if unchanged else 200, etag=etag)
        elif parsed.path == "/api/batch/find_marks":
            self._send_json(mark_finder.status())
        elif parsed.path.startswith("/api/batch/image/"):
            name = unquote(parsed.path.removeprefix("/api/batch/image/"))
            path = fix_images.image_path(name)
            if path is None:
                self.send_error(404)
                return
            # 地址带内容版本号（?v=），内容变了地址就变，可以缓存
            self._send_bytes(path.read_bytes(), "image/png", cacheable=True)
        elif parsed.path == "/api/state":
            self._send_json(get_state_snapshot())
        elif parsed.path.startswith("/api/image/"):
            parts = parsed.path.removeprefix("/api/image/").split("/")
            image = state_image(parts[0], unquote(parts[1])) if len(parts) == 2 else None
            if image is None:
                self.send_error(404)
                return
            self._send_bytes(image, "image/png", cacheable=True)
        elif parsed.path == "/api/autofix":
            if parse_qs(parsed.query).get("active") == ["1"]:
                autofix.mark_active()
            self._send_json(autofix.snapshot())
        elif parsed.path.startswith("/api/autofix/image/"):
            key = unquote(parsed.path.removeprefix("/api/autofix/image/"))
            image = autofix.attempt_image(key) if key and "/" not in key else None
            if image is None:
                self.send_error(404)
                return
            # 每张候选图的编号都不同，内容不会变：可以缓存，卡片重画时不用再取一遍
            self._send_bytes(image, "image/png", cacheable=True)
        elif parsed.path == "/api/overview":
            self._send_json(get_overview_snapshot())
        elif parsed.path == "/api/progress/export":
            try:
                progress = export_progress()
            except ValueError as err:
                self._send_json({"ok": False, "error": str(err)}, status=409)
                return
            self._send_bytes(
                json.dumps(progress, ensure_ascii=False).encode("utf-8"),
                "application/json; charset=utf-8",
                filename=f"excel_progress_{datetime.now():%Y%m%d_%H%M%S}.json",
            )
        elif parsed.path == "/api/debug-log/list":
            self._send_json({"tasks": debug_log.tasks()})
        elif parsed.path == "/api/debug-log/copy":
            # 选中的那次任务（默认最近一次），按 〰 分隔线从 log.txt 里整段挑出来
            title, text = debug_log.task_text((parse_qs(parsed.query).get("id") or [""])[0])
            body = debug_log.for_copy(title, text) if text else "（还没有调试日志：跑过一次任务后才会有）"
            self._send_bytes(body.encode("utf-8"), "text/plain; charset=utf-8")
        elif parsed.path == "/api/debug-log/download":
            data = debug_log.LOG_PATH.read_bytes() if debug_log.LOG_PATH.exists() else b""
            self._send_bytes(data, "text/plain; charset=utf-8", filename="log.txt")
        elif parsed.path == "/api/operation-history":
            from core import operation_history
            self._send_json(operation_history.summary())
        elif parsed.path == "/api/records":
            self._send_json(get_execution_report_snapshot())
        elif parsed.path == "/api/records/zip":
            self._send_bytes(build_report_zip(), "application/zip", filename="execution_report.zip")
        elif parsed.path == "/api/records/markdown":
            md_bytes = render_execution_markdown().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/markdown; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="execution_report.md"')
            self.send_header("Content-Length", str(len(md_bytes)))
            self.end_headers()
            self.wfile.write(md_bytes)
        else:
            self.send_error(404)

    def do_POST(self) -> None:
        try:
            length = int(self.headers.get("Content-Length", 0))
            try:
                payload = json.loads(self.rfile.read(length).decode("utf-8")) if length else {}
            except json.JSONDecodeError:
                raise ValueError("请求内容不是有效的 JSON（文件可能已损坏）") from None
            if not isinstance(payload, dict):
                raise ValueError("请求内容必须为 JSON 对象")
            debug_log.event("🖱️ 网页操作", 路径=self.path, 参数=debug_log.summary(payload)[:600])
            self._handle_post(payload)
        except ConnectionError:
            # 浏览器已断开（刷新、关闭页面或离开），结果送不回去，也不必再回错误
            self._log_client_disconnect()
        except (ValueError, TypeError, UnicodeError) as err:
            debug_log.event("🚫 网页操作被拒绝", 路径=self.path, 原因=err)
            if not isinstance(err, ValueError):
                debug_log.exception(err, f"POST {self.path}")
            self._send_json({"ok": False, "error": str(err)}, status=400)
        except sqlite3.Error as err:
            log_msg(f"❌ 无法保存任务: {err}")
            self._send_json({"ok": False, "error": f"无法保存任务: {err}"}, status=500)
        except OSError as err:
            log_msg(f"❌ 操作失败: {err}")
            self._send_json({"ok": False, "error": f"操作失败: {err}"}, status=500)
        except Exception as err:
            debug_log.exception(err, f"POST {self.path}")
            self._send_json({"ok": False, "error": f"程序出错：{err}（详细信息已写进调试日志）"}, status=500)

    def handle_one_request(self) -> None:
        try:
            super().handle_one_request()
        except ConnectionError:
            self._log_client_disconnect()
            self.close_connection = True

    def _log_client_disconnect(self) -> None:
        path = urlparse(self.path).path if getattr(self, "path", None) else "?"
        if path == "/api/clean/preview":
            console_print("ℹ️ 扫描结束前浏览器已断开（页面被刷新或关闭），本次扫描结果未送达，重新扫描即可")
        elif path.startswith("/api/") and getattr(self, "command", "") == "POST":
            console_print(f"ℹ️ 浏览器已断开，{path} 的结果未送达")

    def _handle_post(self, payload: dict) -> None:
        with STATE_LOCK:
            # PDF 人工复检画框只存记录、不碰 Excel：后台任务进行中也照常保存
            if STATE.get("maintenance") and self.path != "/api/pdf_review/save":
                name = ("颜色" if STATE["maintenance"].startswith("color_") else
                        "隐藏状态" if STATE["maintenance"].startswith("hidden_") else
                        "BL 超链接" if STATE["maintenance"].startswith("bl_link_") else
                        "PDF 人工复检标记" if STATE["maintenance"].startswith("pdf_") else "引用错误")
                raise ValueError(f"正在检查或修复{name}，请等当前任务结束")
        if self.path == "/api/batch/start":
            options = {"folder": _require_folder(payload)}
            minutes = payload.get("scan_limit_minutes")
            if minutes is not None:
                # 单个文件扫描的时间上限（分钟），不给就自动估算
                if isinstance(minutes, bool) or minutes not in batch.SCAN_LIMIT_CHOICES:
                    raise ValueError("单个文件的扫描上限只能选自动或 3、5、10、15 分钟")
                options["scan_limit_minutes"] = minutes
            if payload.get("scan_only") is True:
                if payload.get("fix_error_types") is not None:
                    batch.set_fix_error_types(payload.get("fix_error_types"))
                options.update(scan_only=True, highlight=batch.highlight_style(payload.get("highlight")))
            else:
                if payload.get("fix_error_types") is not None:
                    options["fix_error_types"] = batch.set_fix_error_types(
                        payload.get("fix_error_types"), require_non_empty=True)
                else:
                    options["fix_error_types"] = batch.fix_error_types(require_non_empty=True)
                if payload.get("highlight") is not None:
                    options["highlight"] = batch.highlight_style(payload.get("highlight"))
                else:
                    options["highlight_off"] = True
            self._start_job("batch_scan", options)
        elif self.path == "/api/print_layout":
            set_print_layout(payload.get("path"), payload.get("key"), payload.get("direction"))
            if payload.get("direction") == "auto":
                self._start_job("print_layout", {"path": payload["path"], "key": payload["key"]})
            else:
                self._send_json({"ok": True})
        elif self.path == "/api/batch/fix_types":
            self._send_json({"ok": True, "fix_error_types": batch.set_fix_error_types(payload.get("fix_error_types"))})
        elif self.path == "/api/batch/unhighlight":
            self._start_job("batch_unhighlight", {"paths": batch.highlighted_paths()})
        elif self.path == "/api/batch/clear_marks":
            mark_finder.clear_marks(_require_folder(payload))
            self._send_json({"ok": True})
        elif self.path == "/api/pdf_review/save":
            self._send_json({"ok": True, **pdf_review.save(payload.get("pdf") or "", payload.get("rects"))})
        elif self.path == "/api/pdf_review/apply":
            pdf_review.apply(payload.get("highlight"))
            self._send_json({"ok": True, **pdf_review.rects_snapshot()})
        elif self.path == "/api/hidden_repair/run":
            hidden_repair.repair(_require_folder(payload), _require_folder({"folder": payload.get("old_folder")}), payload.get("options"))
            self._send_json({"ok": True})
        elif self.path == "/api/bl_links/scan":
            bl_links.start(_require_folder(payload))
            self._send_json({"ok": True})
        elif self.path == "/api/bl_links/clear":
            bl_links.clear()
            self._send_json({"ok": True})
        elif self.path == "/api/link_repair/run":
            link_repair.repair(_require_folder(payload), _require_folder({"folder": payload.get("old_folder")}))
            self._send_json({"ok": True})
        elif self.path == "/api/batch/find_marks/run":
            mark_finder.repair(_require_folder(payload), _require_folder({"folder": payload.get("old_folder")}), payload.get("colors"))
            self._send_json({"ok": True})
        elif self.path == "/api/link_repair":
            link_repair.start(_require_folder(payload))
            self._send_json({"ok": True})
        elif self.path == "/api/link_repair/preview":
            link_repair.preview(_require_folder(payload))
            self._send_json({"ok": True})
        elif self.path == "/api/link_repair/apply":
            link_repair.apply()
            self._send_json({"ok": True})
        elif self.path == "/api/batch/find_marks":
            mark_finder.start(_require_folder(payload), payload.get("colors"))
            self._send_json({"ok": True})
        elif self.path == "/api/batch/find_marks/preview":
            mark_finder.preview(_require_folder(payload))
            self._send_json({"ok": True})
        elif self.path == "/api/batch/find_marks/apply":
            mark_finder.apply()
            self._send_json({"ok": True})
        elif self.path == "/api/batch/mark":
            batch.mark(payload.get("id"), payload.get("color"))
            self._send_json({"ok": True})
        elif self.path == "/api/batch/mark_many":
            self._send_json({"ok": True, **batch.mark_many(payload.get("card_ids"), payload.get("color"))})
        elif self.path == "/api/batch/pause":
            with STATE_LOCK:
                if not STATE["busy"] or (STATE["job"] or {}).get("kind") != "batch_scan":
                    raise ValueError("没有正在运行的全自动批处理")
                STATE["job"]["pause_requested"] = True
                from core.workflow import _persist_state
                _persist_state()
            self._send_json({"ok": True})
        elif self.path == "/api/batch/to_manual":
            if batch.request_manual():
                self._start_job("resume")  # 暂停中：续跑任务，不再修，直接转人工
            else:
                self._send_json({"ok": True})
        elif self.path == "/api/batch/to_auto":
            batch.check_auto(payload.get("fix_error_types"))
            with STATE_LOCK:
                options = {"folder": STATE["folder"], "back_to_auto": True,
                           "fix_error_types": batch.fix_error_types(require_non_empty=True)}
            if payload.get("highlight") is not None:
                options["highlight"] = batch.highlight_style(payload.get("highlight"))
            else:
                options["highlight_off"] = True
            self._start_job("batch_scan", options)
        elif self.path == "/api/batch/auto_plan":
            # 「🤖 自动调整」先问一下：哪些交给自动、同行连带哪些、哪些跳过和原因，网页确认框里列出来
            self._send_json({"ok": True, **batch.auto_adjust_plan(
                payload.get("card_ids"), fix_types=payload.get("fix_error_types"))})
        elif self.path == "/api/batch/auto_adjust":
            # 自动调整正在跑：排进同一个任务，当前文件做完接着做
            if batch.enqueue_auto_adjust(payload.get("card_ids"), fix_types=payload.get("fix_error_types")):
                self._send_json({"ok": True, "queued": True})
            else:
                self._start_job("batch_scan", batch.start_auto_adjust(
                    payload.get("card_ids"), fix_types=payload.get("fix_error_types")))
        elif self.path == "/api/batch/abandon":
            batch.abandon_pending()
            self._send_json({"ok": True})
        elif self.path == "/api/batch/open_reset":
            self._send_json({"ok": True, "reset": batch.reset_open_history()})
        elif self.path == "/api/batch/open_plan":
            selected_files = payload.get("selected_files") is True
            ids = batch.validate_cards(payload.get("card_ids"), yellow=not selected_files)
            self._send_json({"ok": True, **batch.open_plan(ids, selected_files=selected_files)})
        elif self.path in {"/api/batch/commit", "/api/batch/recheck", "/api/batch/open"}:
            action = self.path.rsplit("/", 1)[1]
            selected_files = action == "open" and payload.get("selected_files") is True
            ids = batch.validate_cards(payload.get("card_ids"), yellow=action != "commit" and not selected_files)
            single = payload.get("single") is True
            color = None
            if action == "commit" and payload.get("color") not in (None, "all"):
                if payload["color"] not in batch.COLOR_NAMES:
                    raise ValueError("请选择要提交的颜色")
                color = payload["color"]
            if action == "recheck" and batch.enqueue_recheck(ids, single=single):
                self._send_json({"ok": True, "queued": True})
            elif action == "open" and STATE["busy"]:
                # 后台在跑：处理完的文件马上打开，还要等的列出来
                self._send_json({"ok": True, **batch.open_ready(ids, selected_files=selected_files)})
            elif action != "open" and batch.enqueue_followup(f"batch_{action}", ids, single=single, color=color):
                # 后台在跑别的任务：排到它后面，结束后接着复检 / 提交
                self._send_json({"ok": True, "queued": True})
            else:
                options = batch.recheck_options(ids, single=single) if action == "recheck" else {"card_ids": ids}
                if action == "open":
                    options = {"card_ids": batch.open_plan(ids, selected_files=selected_files)["card_ids"],
                               "selected_files": selected_files}
                    if not options["card_ids"]:
                        raise ValueError("当前没有可以打开的人工待处理文件")
                if color:
                    options["color"] = color
                self._start_job(f"batch_{action}", options)
        elif self.path == "/api/batch/followups/start":
            # 暂停时留着的排队复检 / 提交：现在就开始
            autofix.release_excel()
            try:
                with STATE_LOCK:
                    started = batch.start_followup({})
            finally:
                autofix.resume_after_release()
            if not started:
                raise ValueError("没有可以开始的排队复检 / 提交（后台可能还在忙）")
            threading.Thread(target=run_pending_workflow, name="任务", daemon=True).start()
            self._send_json({"ok": True})
        elif self.path == "/api/batch/followups/cancel":
            self._send_json({"ok": True, "cancelled": batch.cancel_followups()})
        elif self.path == "/api/batch/slow_files/clear":
            batch.clear_slow_files()
            self._send_json({"ok": True})
        elif self.path == "/api/scan":
            folder = _require_folder(payload)
            self._start_job("scan", {
                "folder": folder,
                "delete_backups": payload.get("delete_backups") is True,
                "full": payload.get("full") is True,
            })

        elif self.path == "/api/fix":
            step_pt = float(payload.get("step_pt", 3.0))
            max_iter = payload.get("max_iter", 6)
            if not math.isfinite(step_pt) or not 0.5 <= step_pt <= 30:
                raise ValueError("每次递增必须在 0.5 至 30 pt 之间")
            if isinstance(max_iter, bool) or not isinstance(max_iter, int) or max_iter < 1:
                raise ValueError("最大循环次数必须为正整数")
            selected_ids = payload.get("selected_ids")
            with STATE_LOCK:
                if (
                    not isinstance(selected_ids, list) or not selected_ids
                    or any(not isinstance(iid, str) or iid not in STATE["items"] for iid in selected_ids)
                ):
                    raise ValueError("请选择有效的待修复项")
            self._start_job("fix", {
                "selected_ids": list(dict.fromkeys(selected_ids)),
                "step_pt": step_pt,
                "max_iter": max_iter,
            })

        elif self.path == "/api/quick_task":
            kind = payload.get("kind")
            if not isinstance(kind, str) or kind not in QUICK_TASK_NAMES:
                raise ValueError("未知的功能")
            folder, workbooks = list_workbooks(_require_folder(payload))
            options: dict[str, object] = {"folder": str(folder)}
            if kind == "clean_sheets" and "plan" in payload:
                options["plan"] = validate_clean_plan(payload["plan"])
            if kind == "clean_sheets" and "files" in payload:
                options["files"] = validate_clean_files(payload["files"])
                existing = {workbook_rel_name(wb, folder) for wb in workbooks}
                if not existing.intersection(options["files"]):
                    raise ValueError("要处理的文件已被移动或改名，请重新扫描")
            if kind == "export_pdf":
                # 用户可选：不比较横竖方向能省一半以上时间（2026-10-30）
                options["choose_orientation"] = payload.get("choose_orientation", True) is not False
            self._start_job(kind, options)

        elif self.path == "/api/clean/preview":
            self._send_json({"ok": True, **preview_backup_sheets(_require_folder(payload))})

        elif self.path == "/api/clean/file":
            file_name = payload.get("file")
            if not isinstance(file_name, str) or not file_name:
                raise ValueError("需要有效的文件名")
            self._send_json({"ok": True, **workbook_preview_info(_require_folder(payload), file_name)})

        elif self.path == "/api/clean/render":
            file_name = payload.get("file")
            if not isinstance(file_name, str) or not file_name:
                raise ValueError("需要有效的文件名")
            try:
                result = render_workbook_previews(_require_folder(payload), file_name)
            except ValueError:
                raise
            except Exception as err:
                self._send_json({"ok": False, "error": f"生成预览图失败: {err}"}, status=500)
                return
            self._send_json({"ok": True, **result})

        elif self.path == "/api/excel/compare":
            values = [payload.get(key, "") for key in ("kind", "folder", "file", "sheet", "cell")]
            if any(not isinstance(value, str) for value in values) or not all(values[:3]):
                raise ValueError("需要有效的修复结果、文件夹和文件名")
            target = repair_compare.comparison_target(*values)
            with STATE_LOCK:
                job = STATE.get("job") or {}
                if STATE["busy"]:
                    for path in (target["current"], target["backup"]):
                        reason = batch.open_block_reason(job, path)
                        if reason:
                            raise ValueError(reason)
            autofix.wait_for_writes(15)
            self._send_json({"ok": True, **repair_compare.open_comparison(target)})

        elif self.path == "/api/excel/open":
            # 网页上任何地方点文件名：在可见 Excel 打开，知道工作表/单元格就跳过去
            file_name, sheet, cell = payload.get("file"), payload.get("sheet") or "", payload.get("cell") or ""
            if not isinstance(file_name, str) or not file_name or not isinstance(sheet, str) or not isinstance(cell, str):
                raise ValueError("需要有效的文件名")
            with STATE_LOCK:
                workbench_folder = STATE.get("folder", "")
            if payload.get("record") is True:
                wb_path = find_recorded_workbook(file_name)
            else:
                folder = payload.get("folder")
                wb_path = find_workbook(file_name, [
                    folder if isinstance(folder, str) else "", workbench_folder, EXECUTION_STATE.get("folder", ""),
                ])
            with STATE_LOCK:
                job = STATE["job"] or {}
                kind = job.get("kind", "")
                # 扫描 / 复检中：已经做完、后面不再碰的文件可以打开，只拦还没轮到和正在处理的
                if STATE["busy"] and (kind in QUICK_TASK_NAMES or kind.startswith("batch_")):
                    reason = batch.open_block_reason(job, wb_path)
                    if reason:
                        raise ValueError(reason)
            try:
                autofix.wait_for_writes(15)  # 后台正在保存这个文件时别同时打开
            except ValueError:
                raise ValueError("后台正在保存 Excel，请几秒后再点") from None
            self._send_json({"ok": True, "note": open_workbook_in_excel(wb_path, sheet, cell)})

        elif self.path == "/api/autofix/choose":
            gid, index = payload.get("id"), payload.get("index")
            if not isinstance(gid, str) or isinstance(index, bool) or not isinstance(index, int):
                raise ValueError("需要有效的分组和图片序号")
            autofix.choose(gid, index)
            self._send_json({"ok": True})

        elif self.path == "/api/autofix/switch":
            gid, index = payload.get("id"), payload.get("index")
            if not isinstance(gid, str) or isinstance(index, bool) or not isinstance(index, int):
                raise ValueError("需要有效的分组和图片序号")
            autofix.switch_group(gid, index)
            self._send_json({"ok": True})

        elif self.path == "/api/autofix/reject":
            gid = payload.get("id")
            if not isinstance(gid, str) or not gid:
                raise ValueError("需要有效的分组")
            autofix.reject_group(gid)
            self._send_json({"ok": True})

        elif self.path == "/api/autofix/confirm":
            gids = payload.get("ids")
            if not isinstance(gids, list) or any(not isinstance(gid, str) or not gid for gid in gids):
                raise ValueError("需要有效的分组")
            self._send_json({"ok": True, "confirmed": autofix.confirm_groups(gids)})

        elif self.path == "/api/autofix/skip":
            gid, item_ids = payload.get("id"), payload.get("item_ids")
            if (
                not isinstance(gid, str) or not gid
                or not isinstance(item_ids, list) or not item_ids
                or any(not isinstance(iid, str) or not iid for iid in item_ids)
            ):
                raise ValueError("需要有效的分组和问题项")
            self._send_json({"ok": True, **autofix.skip_group(gid, item_ids)})

        elif self.path == "/api/review/later":
            gid, item_ids = payload.get("id"), payload.get("item_ids")
            if (
                not isinstance(gid, str) or not gid
                or not isinstance(item_ids, list) or not item_ids
                or any(not isinstance(iid, str) or not iid for iid in item_ids)
            ):
                raise ValueError("需要有效的分组和问题项")
            autofix.defer_group(gid, item_ids)
            self._send_json({"ok": True})

        elif self.path == "/api/review/open":
            self._start_job("review_open")

        elif self.path == "/api/review/recheck":
            self._start_job("review_recheck")

        elif self.path == "/api/autofix/auto":
            on = payload.get("on")
            if not isinstance(on, bool):
                raise ValueError("需要有效的开关状态")
            autofix.set_auto(on)
            self._send_json({"ok": True})

        elif self.path == "/api/progress/import":
            self._send_json({"ok": True, **import_progress(payload)})

        elif self.path == "/api/resume":
            self._start_job("resume")

        elif self.path == "/api/manual/refresh":
            self._start_job("refresh")

        elif self.path == "/api/manual/confirm":
            accepted = payload.get("accepted")
            iid = payload.get("id")
            if not isinstance(accepted, bool) or not isinstance(iid, str):
                raise ValueError("需要有效的问题项及人工确认状态")
            set_manual_confirmation(iid, accepted)
            self._send_json({"ok": True})

        elif self.path == "/api/manual/skip":
            decisions = payload.get("decisions")
            if not isinstance(decisions, dict) or any(not isinstance(v, bool) for v in decisions.values()):
                raise ValueError("需要有效的跳过决定")
            self._send_json({"ok": True, **apply_skip_decisions(decisions)})

        elif self.path == "/api/manual/keep":
            decisions = payload.get("decisions")
            if (not isinstance(decisions, dict) or not decisions
                    or any(not isinstance(iid, str) or not iid or not isinstance(v, bool)
                           for iid, v in decisions.items())):
                raise ValueError("需要有效的保留决定")
            self._send_json({"ok": True, **apply_keep_decisions(decisions)})

        elif self.path in {"/api/manual/open", "/api/manual/recheck"}:
            iid = payload.get("id")
            with STATE_LOCK:
                if not isinstance(iid, str) or iid not in STATE["items"]:
                    raise ValueError("请选择有效的问题项")
            kind = "manual_open" if self.path.endswith("/open") else "recheck"
            self._start_job(kind, {"selected_ids": [iid]})

        elif self.path == "/api/update_row":
            iid = str(payload.get("id", ""))
            try:
                new_row = max(1, int(payload.get("excel_row", 1)))
            except (TypeError, ValueError):
                self._send_json({"ok": False, "error": "invalid excel_row"}, status=400)
                return

            updated = update_item_row(iid, new_row)
            self._send_json(
                {"ok": True} if updated else {
                    "ok": False,
                    "error": "行号未修改：项目不存在，或任务正在执行/等待续跑",
                },
                status=200 if updated else 409,
            )
        else:
            self.send_error(404)

    def _start_job(self, kind: str, options: dict | None = None) -> None:
        if kind == "manual_open":
            # 立即排队并反馈；可见 Excel 使用独立实例，不等待副本导出及 Quit。
            queued = queue_job(kind, options)
            if not queued:
                self._send_json({"ok": False, "error": "已有任务正在执行"}, status=409)
                return
            threading.Thread(target=run_pending_workflow, name="任务", daemon=True).start()
            self._send_json({"ok": True})
            return
        # 先让自动处理的后台关掉它的 Excel，任务排上队后它会一直等到任务结束
        autofix.release_excel()
        try:
            queued = queue_job(kind, options)
        finally:
            autofix.resume_after_release()
        if not queued:
            self._send_json({"ok": False, "error": "已有任务正在执行"}, status=409)
            return
        threading.Thread(target=run_pending_workflow, name="任务", daemon=True).start()
        self._send_json({"ok": True})


def main() -> None:
    from core import operation_history
    operation_history.interrupt_running()
    debug_log.install()
    port = 8765
    try:
        resync_report_from_state()
    except Exception:
        logging.exception("启动时刷新执行记录失败")
    server = ThreadingHTTPServer(("127.0.0.1", port), WebHandler)
    url = f"http://127.0.0.1:{port}"
    console_print(f"🚀 Web 界面已启动: {url}")
    console_print("正在自动打开浏览器... (按 Ctrl+C 可退出程序)")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        console_print("\n已退出。")


if __name__ == "__main__":
    main()
