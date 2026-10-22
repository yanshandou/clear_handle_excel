import json
import mimetypes
import math
import sqlite3
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

from core import autofix
from core.common import (
    STATIC_DIR,
    console_print,
    find_backup_sheets,
    sanitize_filename,
    workbook_rel_name,
)
from core.excel_ops import extract_sheet_snapshot, match_excel_row, normalize_workbook_view
from core.pdf_detector import (
    extract_page_lines,
    find_clipped_spans_on_page,
    is_region_still_clipped,
    render_crop_base64,
    render_page_overview,
)
from core.reporter import get_execution_report_snapshot, open_recorded_workbook, render_execution_markdown
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
    import_progress,
    list_workbooks,
    log_msg,
    preview_backup_sheets,
    scan_folder_workflow,
    queue_job,
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
    ) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if filename is not None:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
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
            self._serve_static_file("index.html")
        elif parsed.path in {"/records", "/records.html"}:
            self._serve_static_file("records.html")
        elif parsed.path == "/favicon.ico":
            self._serve_static_file("favicon.svg")
        elif parsed.path.startswith("/static/"):
            rel = parsed.path.removeprefix("/static/")
            self._serve_static_file(rel)
        elif parsed.path == "/api/state":
            self._send_json(get_state_snapshot())
        elif parsed.path == "/api/autofix":
            if parse_qs(parsed.query).get("active") == ["1"]:
                autofix.mark_active()
            self._send_json(autofix.snapshot())
        elif parsed.path.startswith("/api/autofix/image/"):
            parts = parsed.path.removeprefix("/api/autofix/image/").split("/")
            image = (
                autofix.attempt_image(parts[0], int(parts[1]))
                if len(parts) == 2 and parts[1].isdigit() else None
            )
            if image is None:
                self.send_error(404)
                return
            self._send_bytes(image, "image/png")
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
        elif parsed.path == "/api/records":
            self._send_json(get_execution_report_snapshot())
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
            self._handle_post(payload)
        except ConnectionError:
            # 浏览器已断开（刷新、关闭页面或离开），结果送不回去，也不必再回错误
            self._log_client_disconnect()
        except (ValueError, TypeError, UnicodeError) as err:
            self._send_json({"ok": False, "error": str(err)}, status=400)
        except sqlite3.Error as err:
            log_msg(f"❌ 无法保存任务: {err}")
            self._send_json({"ok": False, "error": f"无法保存任务: {err}"}, status=500)
        except OSError as err:
            log_msg(f"❌ 操作失败: {err}")
            self._send_json({"ok": False, "error": f"操作失败: {err}"}, status=500)

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
        if self.path == "/api/scan":
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

        elif self.path == "/api/records/open":
            file_name = payload.get("file")
            if not isinstance(file_name, str) or not file_name:
                raise ValueError("需要有效的文件名")
            autofix.release_excel()  # 别让系统把文件开到自动处理后台的隐藏 Excel 里
            try:
                open_recorded_workbook(file_name)
            finally:
                autofix.resume_after_release()
            self._send_json({"ok": True})

        elif self.path == "/api/autofix/choose":
            gid, index = payload.get("id"), payload.get("index")
            if not isinstance(gid, str) or isinstance(index, bool) or not isinstance(index, int):
                raise ValueError("需要有效的分组和图片序号")
            autofix.choose(gid, index)
            self._send_json({"ok": True})

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
            threading.Thread(target=run_pending_workflow, daemon=True).start()
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
        threading.Thread(target=run_pending_workflow, daemon=True).start()
        self._send_json({"ok": True})


def main() -> None:
    port = 8765
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
