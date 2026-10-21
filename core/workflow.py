import copy
import re
import shutil
import threading
from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

from core.common import (
    EXCEL_EXTENSIONS,
    EXCLUDE_PDF_SHEETS,
    XL_SHEET_VISIBLE,
    XL_TYPE_PDF,
    as_bbox,
    console_print,
    is_backup_sheet,
    sanitize_filename,
    unique_pdf_path,
)
from core.excel_ops import (
    extract_sheet_snapshot,
    match_excel_row,
    normalize_workbook_view,
    open_workbook_for_manual_review,
    pythoncom,
    scan_workbook_sheets,
    win32,
)
from core.pdf_detector import (
    collect_text_spans,
    find_clipped_spans_on_page,
    fitz,
    is_region_still_clipped,
    locate_text_occurrence,
    refresh_pdf_pages_overview,
    render_crop_base64,
    render_page_overview,
    RegionTrackingError,
)
from core.reporter import (
    finish_execution_batch,
    get_execution_report_snapshot,
    init_execution_batch,
    interrupt_execution_batch,
    replace_execution_report,
    update_file_progress,
)
from core.storage import load_snapshot, save_snapshot

STATE_LOCK = threading.RLock()

# 主页上可单独执行的批处理功能（job kind -> 任务名）
QUICK_TASK_NAMES = {
    "clean_sheets": "精简 Sheet（删除 YYYYMMDD 备份表）",
    "export_pdf": "批量导出 PDF",
    "normalize_view": "视图标准化（A1 + 100% 缩放）",
}

# 全局状态存储（供前端网页读取和操作）
STATE: dict[str, Any] = {
    "folder": "",
    "pdf_folder": "",
    "backup_folder": "",
    "pages": {},  # page_id -> 整页全景图与页面尺寸信息
    "items": {},  # item_id -> 检测出的切字项详情
    "logs": [],
    "busy": False,
    "settings": {"step_pt": 3.0, "max_iter": 6, "delete_backups": True},
    "job": None,
    "checked_files": {},  # wb_path -> 最近一次完整检查未解决项时的文件版本
}

_saved_state = load_snapshot("workflow")
if _saved_state is not None:
    STATE.update(_saved_state)
    STATE["busy"] = False
    if STATE["job"] and STATE["job"]["status"] == "running":
        STATE["job"]["status"] = "interrupted"


def _persist_state() -> None:
    with STATE_LOCK:
        save_snapshot("workflow", STATE)


def queue_job(kind: str, options: dict[str, Any] | None = None) -> bool:
    with STATE_LOCK:
        if STATE["busy"]:
            return False
        _expire_manual_confirmations()
        if kind in {"fix", "manual_open", "recheck"}:
            ids = (options or {}).get("selected_ids")
            if (
                not isinstance(ids, list) or not ids
                or any(not isinstance(iid, str) or iid not in STATE["items"] for iid in ids)
            ):
                raise ValueError("请选择有效的问题项")
        previous = copy.deepcopy(STATE)
        if kind == "resume":
            if not STATE["job"] or STATE["job"]["status"] != "interrupted":
                raise ValueError("没有可继续的中断任务")
        else:
            STATE["job"] = {
                "kind": kind, "options": options or {}, "completed": [],
                "pending_adjustment": None, "status": "running",
            }
            if kind == "fix":
                for iid in STATE["job"]["options"]["selected_ids"]:
                    STATE["items"][iid]["job_iterations"] = 0
            STATE["settings"].update({
                key: value for key, value in (options or {}).items()
                if key in STATE["settings"]
            })
        STATE["job"]["status"] = "running"
        STATE["busy"] = True
        try:
            _persist_state()
        except Exception:
            STATE.clear()
            STATE.update(previous)
            raise
        return True


def run_pending_workflow() -> None:
    try:
        with STATE_LOCK:
            job = STATE["job"]
            options = dict(job["options"])
        if job["kind"] == "scan":
            scan_folder_workflow(options["folder"], options["delete_backups"], resume=True)
            expected = job.get("workbooks", [])
        elif job["kind"] == "manual_open" and job["completed"]:
            expected = options["selected_ids"]
        elif job["kind"] in {"manual_open", "recheck"}:
            manual_review_workflow(options["selected_ids"], open_excel=job["kind"] == "manual_open")
            expected = options["selected_ids"]
        elif job["kind"] in QUICK_TASK_NAMES:
            quick_task_workflow(job["kind"])
            expected = job.get("workbooks", [])
        elif job["kind"] == "refresh":
            if "selected_ids" not in job["options"]:
                ids = _select_changed_problem_items()
                with STATE_LOCK:
                    job["options"]["selected_ids"] = ids
                    _persist_state()
            ids = job["options"]["selected_ids"]
            if ids:
                recheck_items_workflow(ids)
            else:
                message = "刷新复检：没有修改过的未解决问题文件，未打开 Excel 或重新导出 PDF"
                log_msg(message)
                finish_execution_batch(message)
            expected = ids
        else:
            fix_selected_workflow(
                options["selected_ids"], options["step_pt"], options["max_iter"], resume=True
            )
            expected = options["selected_ids"]
        with STATE_LOCK:
            job["status"] = (
                "completed" if set(expected).issubset(job["completed"]) else "interrupted"
            )
        if job["status"] == "interrupted":
            interrupt_execution_batch("部分文件处理失败，可点击继续上次任务重试")
    except Exception as err:
        with STATE_LOCK:
            STATE["job"]["status"] = "interrupted"
        log_msg(f"❌ 任务中断，可继续上次任务: {err}")
        interrupt_execution_batch(f"任务中断: {err}")
    finally:
        with STATE_LOCK:
            STATE["busy"] = False
            _persist_state()


def log_msg(msg: str) -> None:
    timestamp = datetime.now().strftime("%H:%M:%S")
    line = f"[{timestamp}] {msg}"
    with STATE_LOCK:
        STATE["logs"].append(line)
    console_print(line)


def try_acquire_busy() -> bool:
    with STATE_LOCK:
        if STATE["busy"]:
            return False
        STATE["busy"] = True
        return True


def _progress_summary(report_snap: dict[str, Any]) -> dict[str, Any]:
    keys = (
        "task_name", "status", "total_files", "completed_files", "processing_files",
        "pending_files", "error_files", "progress_pct", "current_file", "current_step",
    )
    return {key: report_snap[key] for key in keys}


def get_state_snapshot() -> dict[str, Any]:
    with STATE_LOCK:
        if not STATE["busy"]:
            _expire_manual_confirmations()
    report_snap = get_execution_report_snapshot()
    with STATE_LOCK:
        return {
            "busy": STATE["busy"],
            "folder": STATE["folder"],
            "pdf_folder": STATE["pdf_folder"],
            "logs": list(STATE["logs"][-40:]),
            "pages": copy.deepcopy(STATE["pages"]),
            "items": copy.deepcopy(STATE["items"]),
            "settings": dict(STATE["settings"]),
            "can_resume": bool(STATE["job"] and STATE["job"]["status"] == "interrupted"),
            "job": copy.deepcopy(STATE["job"]),
            "progress": _progress_summary(report_snap),
        }


def get_overview_snapshot() -> dict[str, Any]:
    """主页轮询用的轻量快照：不含整页截图与问题项详情。"""
    with STATE_LOCK:
        if not STATE["busy"]:
            _expire_manual_confirmations()
    report_snap = get_execution_report_snapshot()
    with STATE_LOCK:
        job = STATE["job"]
        unresolved = sum(
            item["status"] not in {"fixed", "accepted"} for item in STATE["items"].values()
        )
        return {
            "busy": STATE["busy"],
            "folder": STATE["folder"],
            "logs": list(STATE["logs"][-40:]),
            "can_resume": bool(job and job["status"] == "interrupted"),
            "job": {
                "kind": job["kind"],
                "status": job["status"],
                **{key: job["options"].get(key, "") for key in ("folder", "pdf_folder", "backup_folder")},
            } if job else None,
            "clips": {"total": len(STATE["items"]), "unresolved": unresolved},
            "progress": _progress_summary(report_snap),
        }


def update_item_row(item_id: str, new_row: int) -> bool:
    with STATE_LOCK:
        if (
            STATE["busy"] or item_id not in STATE["items"]
            or (STATE["job"] and STATE["job"]["status"] == "interrupted")
        ):
            return False
        item = STATE["items"][item_id]
        previous = dict(item)
        item["excel_row"] = new_row
        old_addr = str(item.get("cell_addr", ""))
        item["cell_addr"] = (
            re.sub(r"\d+$", str(new_row), old_addr)
            if re.search(r"\d+$", old_addr)
            else f"Row {new_row}"
        )
        try:
            _persist_state()
        except Exception:
            item.clear()
            item.update(previous)
            raise
        return True


def _detect_sheet_results(
    sheet: Any, wb_path: Path, sheet_name: str, pdf_path: Path, start_seq: int,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """从当前 PDF 构建新快照；不沿用旧文字、坐标或问题状态。"""
    pages: dict[str, Any] = {}
    items: dict[str, Any] = {}
    sheet_snapshot = extract_sheet_snapshot(sheet)
    generation = uuid4().hex
    with fitz.open(pdf_path) as doc:
        if not len(doc):
            raise ValueError(f"导出的 PDF 没有页面: {pdf_path}")
        text_spans = collect_text_spans(doc)
        for page_index in range(len(doc)):
            page = doc[page_index]
            clips = find_clipped_spans_on_page(page)
            if not clips:
                continue
            page_id = f"page_{generation}_{page_index}"
            image, width, height = render_page_overview(page)
            pages[page_id] = {
                "id": page_id, "wb_path": str(wb_path), "wb_name": wb_path.name,
                "sheet_name": sheet_name, "pdf_path": str(pdf_path),
                "page_index": page_index, "total_pages": len(doc),
                "page_width": width, "page_height": height, "image": image,
            }
            for index, clip in enumerate(clips, start=1):
                row, address, height = match_excel_row(
                    sheet_snapshot, clip["text"], page_index, len(doc),
                    clip["bbox"][1] / max(1.0, page.rect.height),
                )
                item_id = f"item_{generation}_{len(items)}"
                occurrence = locate_text_occurrence(text_spans, page_index, clip["text"], clip["bbox"])
                items[item_id] = {
                    "id": item_id, "seq": start_seq + len(items), "page_id": page_id,
                    "index_on_page": index, "wb_path": str(wb_path), "wb_name": wb_path.name,
                    "sheet_name": sheet_name, "pdf_path": str(pdf_path),
                    "page_index": page_index, "text": clip["text"], "bbox": clip["bbox"],
                    "occurrence": list(occurrence) if occurrence else None,
                    "excel_row": row, "cell_addr": address, "row_height": height,
                    "image": render_crop_base64(page, clip["bbox"], highlight=True),
                    "status": "pending", "iterations": 0,
                }
    return pages, items


def _file_version(path: Path) -> dict[str, int]:
    stat = path.stat()
    return {"mtime_ns": stat.st_mtime_ns, "size": stat.st_size}


def _sync_problem_counts(path: str, action: str) -> None:
    items = [item for item in STATE["items"].values() if item["wb_path"] == path]
    fixed = sum(item["status"] == "fixed" for item in items)
    accepted = sum(item["status"] == "accepted" for item in items)
    pending = len(items) - fixed - accepted
    update_file_progress(
        Path(path).name, status="needs_fix" if pending else "done",
        clip_detected=len(items), clip_fixed=fixed, clip_accepted=accepted,
        clip_unfixed=pending, current_action=action, error="",
    )


def _expire_manual_confirmations() -> None:
    changed: set[str] = set()
    versions: dict[str, dict[str, int] | None] = {}
    for item in STATE["items"].values():
        if item["status"] != "accepted":
            continue
        path = item["wb_path"]
        if path not in versions:
            try:
                versions[path] = _file_version(Path(path))
            except FileNotFoundError:
                versions[path] = None
        if versions[path] != item["manual_confirmation"]["file_version"]:
            item["status"] = "needs_review"
            item["review_reason"] = "文件版本已变化，之前的人工确认已失效，请重新检查"
            changed.add(path)
    if changed:
        _persist_state()
        for path in changed:
            _sync_problem_counts(path, "文件已变化，人工确认失效，等待重新检查")
            log_msg(f"人工确认已失效: {Path(path).name}，请重新检查")


def set_manual_confirmation(item_id: str, accepted: bool) -> None:
    with STATE_LOCK:
        if STATE["busy"]:
            raise ValueError("任务正在执行，请结束后再进行人工确认")
        item = STATE["items"].get(item_id)
        if item is None:
            raise ValueError("问题项已更新，请刷新页面后重试")
        previous = copy.deepcopy(item)
        previous_job = copy.deepcopy(STATE["job"])
        if accepted:
            if item["status"] == "fixed":
                raise ValueError("此项已自动修复，无需人工确认")
            version = _file_version(Path(item["wb_path"]))
            if item["status"] == "accepted":
                raise ValueError("此项已人工确认，如需重新判断请先撤销")
            item["manual_confirmation"] = {
                "file_version": version,
                "confirmed_at": datetime.now().isoformat(timespec="seconds"),
                "previous_status": item["status"],
            }
            item["status"] = "accepted"
            item["review_reason"] = "人工确认无需修复（仅当前已保存文件版本有效）"
        else:
            if item["status"] != "accepted":
                raise ValueError("此项没有有效的人工确认")
            item["status"] = item["manual_confirmation"]["previous_status"]
            item["review_reason"] = "人工确认已撤销，重新纳入待处理"
            del item["manual_confirmation"]
        if STATE["job"] and STATE["job"]["status"] == "interrupted":
            STATE["job"]["status"] = "cancelled"
            STATE["job"]["pending_adjustment"] = None
        try:
            _persist_state()
        except Exception:
            item.clear()
            item.update(previous)
            STATE["job"] = previous_job
            raise
        _sync_problem_counts(item["wb_path"], item["review_reason"])
        if previous_job and previous_job["status"] == "interrupted":
            finish_execution_batch("已取消中断任务的续跑，按人工判断更新问题状态")
        log_msg(f"{item['review_reason']}: [{item['wb_name']} / {item['sheet_name']} / {item['cell_addr']}]")
        _persist_state()


def _select_changed_problem_items() -> list[str]:
    with STATE_LOCK:
        _expire_manual_confirmations()
        by_file: dict[str, list[str]] = {}
        for item in STATE["items"].values():
            if item["status"] not in {"fixed", "accepted"}:
                by_file.setdefault(item["wb_path"], []).append(item["id"])
        checked_files = copy.deepcopy(STATE["checked_files"])
    selected: list[str] = []
    changed_files = 0
    for path, ids in by_file.items():
        if _file_version(Path(path)) != checked_files.get(path):
            selected.extend(ids)
            changed_files += 1
            if path not in checked_files:
                log_msg(f"刷新复检：{Path(path).name} 尚无检查时间戳，首次纳入复检")
    log_msg(
        f"刷新复检：{len(by_file)} 个未解决问题文件中，"
        f"筛出 {changed_files} 个已修改文件、{len(selected)} 处问题"
    )
    return selected


def list_workbooks(folder_str: str) -> tuple[Path, list[Path]]:
    """解析文件夹路径并列出其中的 Excel 文件（跳过 ~$ 临时锁文件）"""
    folder = Path(folder_str.strip().strip('"').strip("'")).expanduser().resolve()
    if not folder.is_dir():
        raise ValueError(f"文件夹不存在: {folder}")
    workbooks = sorted(
        p
        for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS and not p.name.startswith("~$")
    )
    if not workbooks:
        raise ValueError("未找到 Excel 文件")
    return folder, workbooks


def _is_exportable_sheet(sheet: Any) -> bool:
    """可见、且不是変更履歴或日期备份表的工作表才导出 PDF"""
    name = sheet.Name
    return (
        name not in EXCLUDE_PDF_SHEETS
        and not is_backup_sheet(name)
        and sheet.Visible == XL_SHEET_VISIBLE
    )


def _export_sheet_pdf(sheet: Any, pdf_path: Path) -> None:
    sheet.ExportAsFixedFormat(
        Type=XL_TYPE_PDF,
        Filename=str(pdf_path.resolve()),
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False,
    )


def _backup_original(wb_path: Path, backup_folder: Path) -> None:
    """修改前备份原文件；续跑时已有备份不覆盖，保留最初版本。"""
    backup_folder.mkdir(parents=True, exist_ok=True)
    if not (backup_folder / wb_path.name).exists():
        shutil.copy2(wb_path, backup_folder / wb_path.name)


def scan_folder_workflow(folder_str: str, delete_backups: bool, *, resume: bool = False) -> None:
    with STATE_LOCK:
        STATE["busy"] = True
        job = STATE["job"] if resume else None
        initialized = bool(job and "workbooks" in job)
        if not initialized:
            STATE["pages"].clear()
            STATE["items"].clear()
            STATE["logs"].clear()
            STATE["checked_files"].clear()

    try:
        folder, workbooks = list_workbooks(folder_str)

        timestamp = f"{datetime.now():%Y%m%d_%H%M%S}"
        pdf_folder = Path(STATE["pdf_folder"]) if initialized else folder / f"_exported_pdfs_{timestamp}"
        backup_folder = Path(STATE["backup_folder"]) if initialized else folder / f"_backup_{timestamp}"
        if initialized:
            assert job is not None
            workbooks = [Path(path) for path in job["workbooks"]]
        pdf_folder.mkdir(parents=True, exist_ok=True)
        used_pdf_paths = {Path(p["pdf_path"]) for p in STATE["pages"].values()}

        with STATE_LOCK:
            STATE["folder"] = str(folder)
            STATE["pdf_folder"] = str(pdf_folder)
            STATE["backup_folder"] = str(backup_folder)

        if not initialized:
            init_execution_batch(
                task_name="Web 导出 PDF 与切字扫描",
                folder=folder,
                workbooks=workbooks,
                pdf_folder=pdf_folder,
                backup_folder=backup_folder,
            )
        with STATE_LOCK:
            if job is not None and not initialized:
                job["workbooks"] = [str(path) for path in workbooks]
            _persist_state()

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        item_seq = max((it["seq"] for it in STATE["items"].values()), default=0) + 1
        try:
            for wb_path in workbooks:
                if job is not None and str(wb_path) in job["completed"]:
                    continue
                # 未完成的文件重新扫描，已完成文件的结果保持不变。
                with STATE_LOCK:
                    for collection in ("pages", "items"):
                        for key in list(STATE[collection]):
                            if STATE[collection][key]["wb_path"] == str(wb_path):
                                del STATE[collection][key]
                log_msg(f"正在处理并导出: {wb_path.name} ...")
                update_file_progress(
                    wb_path.name,
                    status="processing",
                    current_action="正在打开工作簿并检查备份表...",
                )
                wb = None
                wb_exported_pdfs = 0
                wb_detected_clips = 0
                deleted_bkp_list: list[str] = []
                try:
                    wb = excel.Workbooks.Open(str(wb_path.resolve()))
                    all_names = [s.Name for s in wb.Sheets]
                    if initialized:
                        previous = get_execution_report_snapshot()["files"].get(wb_path.name, {})
                        deleted_bkp_list = list(previous.get("deleted_backups", []))
                        all_names = list(dict.fromkeys(previous.get("all_sheets", []) + all_names))
                    bkp_names = [n for n in all_names if is_backup_sheet(n)]
                    current_names = {s.Name for s in wb.Sheets}
                    bkp_names = [n for n in bkp_names if n in current_names]

                    if delete_backups and bkp_names:
                        if len(bkp_names) >= len(all_names):
                            log_msg(f"  ⚠️ 跳过删除备份表（全部工作表均为日期名）: {wb_path.name}")
                        else:
                            _backup_original(wb_path, backup_folder)
                            for bn in bkp_names:
                                wb.Sheets(bn).Delete()
                            deleted_bkp_list = list(dict.fromkeys(deleted_bkp_list + bkp_names))
                            log_msg(f"  已清理 {len(bkp_names)} 个日期备份 Sheet")
                            update_file_progress(
                                wb_path.name, all_sheets=all_names, deleted_backups=deleted_bkp_list
                            )
                            wb.Save()

                    update_file_progress(
                        wb_path.name,
                        all_sheets=all_names,
                        deleted_backups=deleted_bkp_list,
                        current_action="正在导出 PDF 并检测切字...",
                        error="",
                    )

                    scan_version = _file_version(wb_path)
                    file_stem = sanitize_filename(wb_path.stem)
                    for sheet in wb.Sheets:
                        if not _is_exportable_sheet(sheet):
                            continue
                        s_name = sheet.Name

                        pdf_path = unique_pdf_path(
                            pdf_folder, file_stem, sanitize_filename(s_name), used_pdf_paths
                        )
                        try:
                            _export_sheet_pdf(sheet, pdf_path)
                            wb_exported_pdfs += 1
                        except Exception as e:
                            log_msg(f"  导出 PDF 失败 [{s_name}]: {e}")
                            raise

                        pages, items = _detect_sheet_results(sheet, wb_path, s_name, pdf_path, item_seq)
                        wb_detected_clips += len(items)
                        item_seq += len(items)
                        with STATE_LOCK:
                            STATE["pages"].update(pages)
                            STATE["items"].update(items)

                        update_file_progress(
                            wb_path.name,
                            exported_pdfs=wb_exported_pdfs,
                            clip_detected=wb_detected_clips,
                            clip_unfixed=wb_detected_clips,
                        )

                    if _file_version(wb_path) != scan_version:
                        raise ValueError(f"扫描期间文件已修改，请重新检查: {wb_path.name}")
                    if wb_detected_clips > 0:
                        update_file_progress(
                            wb_path.name,
                            status="needs_fix",
                            exported_pdfs=wb_exported_pdfs,
                            clip_detected=wb_detected_clips,
                            clip_unfixed=wb_detected_clips,
                            current_action=f"已导出 {wb_exported_pdfs} 份 PDF，检出 {wb_detected_clips} 处切字待微调",
                        )
                    else:
                        update_file_progress(
                            wb_path.name,
                            status="done",
                            exported_pdfs=wb_exported_pdfs,
                            clip_detected=0,
                            clip_unfixed=0,
                            current_action=f"已导出 {wb_exported_pdfs} 份 PDF，无切字异常",
                        )
                    with STATE_LOCK:
                        STATE["checked_files"][str(wb_path)] = scan_version
                        if job is not None:
                            job["completed"].append(str(wb_path))
                        _persist_state()
                except Exception as err:
                    log_msg(f"❌ 处理文件出错 {wb_path.name}: {err}")
                    update_file_progress(
                        wb_path.name,
                        status="error",
                        error=str(err),
                        current_action=f"处理出错: {err}",
                    )
                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        with STATE_LOCK:
            total_detected = len(STATE["items"])
            total_pages_with_clips = len(STATE["pages"])
        summary_msg = f"✅ 扫描完成！在 {total_pages_with_clips} 页 PDF 中共检测到 {total_detected} 处疑似切字位置。"
        log_msg(summary_msg)
        finish_execution_batch(summary_msg)
    finally:
        with STATE_LOCK:
            if not resume:
                STATE["busy"] = False
                _persist_state()


def _sync_row_height_in_state(wb_path_str: str, sheet_name: str, target_row: int, new_h: float) -> None:
    with STATE_LOCK:
        for other in STATE["items"].values():
            if (
                other["wb_path"] == wb_path_str
                and other["sheet_name"] == sheet_name
                and int(other["excel_row"]) == target_row
            ):
                other["row_height"] = new_h


def _refresh_sheet_pages_in_state(pdf_path: Path) -> None:
    pdf_path_str = str(pdf_path)
    with STATE_LOCK:
        matching_pages = [
            p for p in STATE["pages"].values() if p.get("pdf_path") == pdf_path_str
        ]
    if not matching_pages:
        return

    indices = {int(p["page_index"]) for p in matching_pages}
    refreshed = refresh_pdf_pages_overview(pdf_path, indices)
    if not refreshed:
        return

    with STATE_LOCK:
        for p in matching_pages:
            p_idx = int(p["page_index"])
            if p_idx in refreshed:
                b64, pw, ph = refreshed[p_idx]
                p["image"] = b64
                p["page_width"] = pw
                p["page_height"] = ph


def _update_item_pdf_location(
    item: dict[str, Any],
    image: str,
    bbox: tuple[float, float, float, float],
    page_index: int,
) -> None:
    with STATE_LOCK:
        if page_index != item["page_index"]:
            page_entry = next(
                (page for page in STATE["pages"].values()
                 if page["pdf_path"] == item["pdf_path"] and page["page_index"] == page_index),
                None,
            )
            if page_entry is None:
                with fitz.open(item["pdf_path"]) as doc:
                    overview, width, height = render_page_overview(doc[page_index])
                    page_id = f"relocated_{item['id']}_{page_index}"
                    page_entry = {
                        "id": page_id, "wb_path": item["wb_path"], "wb_name": item["wb_name"],
                        "sheet_name": item["sheet_name"], "pdf_path": item["pdf_path"],
                        "page_index": page_index, "total_pages": len(doc),
                        "page_width": width, "page_height": height, "image": overview,
                    }
                STATE["pages"][page_entry["id"]] = page_entry
            item["page_id"] = page_entry["id"]
        item["page_index"] = page_index
        item["image"] = image
        item["bbox"] = bbox


def _locate_item_for_fix(
    item: dict[str, Any], pdf_path: Path, pdf_unchanged: bool,
) -> tuple[bool, str, tuple[float, float, float, float], int] | None:
    """定位目标文字；无法可靠定位时只将当前项标记为人工核对并返回 None，其余选中项继续处理。"""
    if item.get("occurrence") is None and pdf_unchanged:
        # 旧版扫描结果没有出现序号；本轮尚未重新导出该 PDF 时，可按原坐标精确补录。
        with fitz.open(pdf_path) as doc:
            occurrence = locate_text_occurrence(
                collect_text_spans(doc), item["page_index"], item["text"], as_bbox(item["bbox"])
            )
        if occurrence is not None:
            with STATE_LOCK:
                item["occurrence"] = list(occurrence)
    try:
        located = is_region_still_clipped(
            pdf_path, item["page_index"], item["text"], as_bbox(item["bbox"]),
            occurrence=item.get("occurrence"),
        )
    except RegionTrackingError as err:
        with STATE_LOCK:
            item["status"] = "needs_review"
            item["review_reason"] = str(err)
            _persist_state()
        log_msg(
            f"  ⚠️ [{item['sheet_name']} 行{item['excel_row']}] {err}；已跳过此处，继续处理其余选中项"
        )
        return None
    with STATE_LOCK:
        item.pop("review_reason", None)
    return located


def manual_review_workflow(selected_ids: list[str], *, open_excel: bool) -> None:
    with STATE_LOCK:
        job = STATE["job"]
        if not open_excel and "recheck_plan" in job["options"]:
            items = []
        else:
            items = [STATE["items"][iid] for iid in selected_ids if iid not in job["completed"]]
    if not open_excel and "recheck_plan" in job["options"]:
        recheck_items_workflow(selected_ids)
        return
    if open_excel:
        item = items[0]
        height = open_workbook_for_manual_review(
            Path(item["wb_path"]), item["sheet_name"], item["cell_addr"], int(item["excel_row"])
        )
        with STATE_LOCK:
            _sync_row_height_in_state(item["wb_path"], item["sheet_name"], int(item["excel_row"]), height)
            if item["status"] != "accepted":
                item["status"] = "manual_review"
                item["review_reason"] = "请人工核对自动推测的行号；调整并保存后重新导出复检"
            wb_items = [it for it in STATE["items"].values() if it["wb_path"] == item["wb_path"]]
            fixed = sum(it["status"] == "fixed" for it in wb_items)
            job["completed"].append(item["id"])
            _persist_state()
        message = (
            f"已在 Excel 定位 [{item['wb_name']} / {item['sheet_name']} / {item['cell_addr']}]，"
            "请人工检查并保存，然后点击重新导出复检（不会自动加高）"
        )
        log_msg(message)
        update_file_progress(
            item["wb_name"], status="needs_fix", current_action=message, error="",
            clip_fixed=fixed, clip_unfixed=sum(it["status"] not in {"fixed", "accepted"} for it in wb_items),
        )
        finish_execution_batch("等待人工调整并保存 Excel")
        return

    # 单项强制复检也覆盖同文件的未解决项，避免只检查一项就推进整个文件的时间戳。
    with STATE_LOCK:
        paths = {item["wb_path"] for item in items}
        selected = set(selected_ids)
        ids = [
            item["id"] for item in STATE["items"].values()
            if item["id"] in selected or (item["wb_path"] in paths and item["status"] not in {"fixed", "accepted"})
        ]
        job["options"]["selected_ids"] = ids
        _persist_state()
    recheck_items_workflow(ids)


def recheck_items_workflow(selected_ids: list[str]) -> None:
    with STATE_LOCK:
        job = STATE["job"]
        if "recheck_plan" not in job["options"]:
            plan: dict[str, Any] = {}
            for iid in selected_ids:
                item = STATE["items"][iid]
                entry = plan.setdefault(item["wb_path"], {"item_ids": [], "sheets": {}})
                entry["item_ids"].append(iid)
                existing = entry["sheets"].get(item["sheet_name"])
                if existing is not None and existing != item["pdf_path"]:
                    raise ValueError(f"同一工作表存在多个 PDF 路径，请重新扫描: {item['sheet_name']}")
                entry["sheets"][item["sheet_name"]] = item["pdf_path"]
            job["options"]["recheck_plan"] = plan
            _persist_state()
        by_file = {
            path: entry for path, entry in job["options"]["recheck_plan"].items()
            if not set(entry["item_ids"]).issubset(job["completed"])
        }
    if not by_file:
        finish_execution_batch("复检任务已完成")
        return
    init_execution_batch(
        task_name="人工调整后导出复检（不自动加高）",
        folder=Path(STATE["folder"]),
        workbooks=[Path(path) for path in by_file],
        preserve_existing_files=True,
    )
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        for path, entry in by_file.items():
            workbook = None
            try:
                version = _file_version(Path(path))
                log_msg(f"刷新复检文件: {Path(path).name}（重新检测 {len(entry['sheets'])} 个问题工作表）")
                update_file_progress(
                    Path(path).name, status="processing",
                    current_action="根据新 PDF 重新检测问题工作表（不自动加高）", error="",
                )
                workbook = excel.Workbooks.Open(str(Path(path).resolve()), ReadOnly=True)
                new_pages: dict[str, Any] = {}
                new_items: dict[str, Any] = {}
                with STATE_LOCK:
                    next_seq = max((item["seq"] for item in STATE["items"].values()), default=0) + 1
                for sheet_name, pdf_path_str in entry["sheets"].items():
                    sheet = workbook.Sheets(sheet_name)
                    pdf_path = Path(pdf_path_str)
                    _export_sheet_pdf(sheet, pdf_path)
                    pages, items = _detect_sheet_results(
                        sheet, Path(path), sheet_name, pdf_path, next_seq,
                    )
                    next_seq += len(items)
                    new_pages.update(pages)
                    new_items.update(items)
                    log_msg(f"  [{sheet_name}] 新 PDF 检出 {len(items)} 处疑似压线，将替换旧标记")
                if _file_version(Path(path)) != version:
                    raise ValueError(f"复检期间文件又被修改，本轮不提交，请保存完成后再次刷新: {Path(path).name}")
                with STATE_LOCK:
                    previous = copy.deepcopy(STATE)
                    protected = [
                        item for item in STATE["items"].values()
                        if item["wb_path"] == path and item["status"] == "accepted"
                        and item["manual_confirmation"]["file_version"] == version
                    ]
                    for result in new_items.values():
                        matches = [
                            item for item in protected
                            if item["sheet_name"] == result["sheet_name"]
                            and item["text"] == result["text"]
                            and item["page_index"] == result["page_index"]
                            and all(abs(a - b) < 0.1 for a, b in zip(item["bbox"], result["bbox"]))
                        ]
                        if len(matches) == 1:
                            result["status"] = "accepted"
                            result["manual_confirmation"] = copy.deepcopy(matches[0]["manual_confirmation"])
                            result["review_reason"] = matches[0]["review_reason"]
                    for collection in ("pages", "items"):
                        for key, value in list(STATE[collection].items()):
                            if value["wb_path"] == path and value["sheet_name"] in entry["sheets"]:
                                del STATE[collection][key]
                    STATE["pages"].update(new_pages)
                    STATE["items"].update(new_items)
                    wb_items = [it for it in STATE["items"].values() if it["wb_path"] == path]
                    fixed = sum(it["status"] == "fixed" for it in wb_items)
                    accepted = sum(it["status"] == "accepted" for it in wb_items)
                    # 只有全部未解决项都被检查，才能更新整个文件的版本基线。
                    if all(it["status"] in {"fixed", "accepted"} or it["sheet_name"] in entry["sheets"] for it in wb_items):
                        STATE["checked_files"][path] = version
                    job["completed"].extend(entry["item_ids"])
                    try:
                        _persist_state()
                    except Exception:
                        STATE.clear()
                        STATE.update(previous)
                        raise
                update_file_progress(
                    Path(path).name, status="done" if fixed + accepted == len(wb_items) else "needs_fix",
                    clip_detected=len(wb_items), clip_fixed=fixed, clip_accepted=accepted,
                    clip_unfixed=len(wb_items) - fixed - accepted,
                    current_action=f"已重新检测问题工作表，当前待处理 {len(wb_items) - fixed - accepted} 处，人工确认 {accepted} 处", error="",
                )
            finally:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
        message = f"刷新复检完成：本轮检查 {len(by_file)} 个文件，未自动增加行高"
        log_msg(message)
        finish_execution_batch(message)
    finally:
        try:
            if excel is not None:
                excel.Quit()
        finally:
            pythoncom.CoUninitialize()


def fix_selected_workflow(
    selected_ids: list[str],
    step_pt: float = 3.0,
    max_iter: int = 6,
    *,
    resume: bool = False,
) -> None:
    """
    针对用户勾选的热区/卡片：
    每次将对应 Excel 行高 +step_pt -> 重新导出该 Sheet 的 PDF -> 检测该处是否还切字
    一旦完整显示则立即停止，并同步刷新整页画布与局部 2.5x 放大镜截图！
    """
    with STATE_LOCK:
        STATE["busy"] = True
        job = STATE["job"] if resume else None
        by_wb: dict[str, list[dict[str, Any]]] = {}
        for iid in selected_ids:
            item = STATE["items"].get(iid)
            if item and (job is None or iid not in job["completed"]):
                by_wb.setdefault(item["wb_path"], []).append(item)
                if not resume:
                    item["job_iterations"] = 0
        backup_folder_str = STATE["backup_folder"]

    try:
        if not by_wb:
            if job is not None and set(selected_ids).issubset(job["completed"]):
                finish_execution_batch("上次微调任务已完成")
            else:
                log_msg("⚠️ 没有选中任何需要修复的项。")
            return

        backup_folder = Path(backup_folder_str)
        backup_folder.mkdir(parents=True, exist_ok=True)
        with STATE_LOCK:
            folder_path = Path(STATE["folder"]) if STATE["folder"] else backup_folder.parent

        init_execution_batch(
            task_name="Web 闭环行高微调",
            folder=folder_path,
            workbooks=[Path(p) for p in by_wb.keys()],
            preserve_existing_files=True,
        )

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        total_skipped = 0
        try:
            for wb_path_str, items in by_wb.items():
                if job is not None and job["pending_adjustment"]:
                    pending_id = job["pending_adjustment"]["item_id"]
                    items.sort(key=lambda item: item["id"] != pending_id)
                wb_path = Path(wb_path_str)
                bkp_target = backup_folder / wb_path.name
                if not bkp_target.exists():
                    shutil.copy2(wb_path, bkp_target)

                log_msg(f"🔧 开始微调工作簿: {wb_path.name} (共 {len(items)} 处)...")
                update_file_progress(
                    wb_path.name,
                    status="processing",
                    current_action=f"正在微调 {len(items)} 处选中切字项...",
                    error="",
                )
                wb = None
                touched_pdfs: set[Path] = set()
                attempted_ids: list[str] = []
                skipped_ids: list[str] = []
                try:
                    wb = excel.Workbooks.Open(str(wb_path.resolve()))

                    for item in items:
                        if item["status"] in {"fixed", "accepted"}:
                            continue
                        attempted_ids.append(item["id"])
                        sheet_name = item["sheet_name"]
                        sheet = wb.Sheets(sheet_name)
                        target_row = max(1, int(item["excel_row"]))
                        pdf_path = Path(item["pdf_path"])
                        row_peers = [
                            peer for peer in items
                            if peer["sheet_name"] == sheet_name
                            and int(peer["excel_row"]) == target_row
                            and peer["status"] != "accepted"
                        ]
                        pending = job["pending_adjustment"] if job else None

                        if not pending:
                            located = _locate_item_for_fix(item, pdf_path, pdf_path not in touched_pdfs)
                            if located is None:
                                skipped_ids.append(item["id"])
                                continue
                            still_clip, new_img, new_bbox, resolved_page = located
                            _update_item_pdf_location(item, new_img, new_bbox, resolved_page)
                            curr_h = round(float(sheet.Rows(target_row).RowHeight or 15.0), 1)
                            if not still_clip:
                                with STATE_LOCK:
                                    item["status"] = "fixed"
                                    item["row_height"] = curr_h
                                _sync_row_height_in_state(wb_path_str, sheet_name, target_row, curr_h)
                                log_msg(f"  ✅ [{sheet_name} 行{target_row}] 复检目标文字已完整显示，无需继续加高")
                                _persist_state()
                                continue

                        used_iterations = max(int(peer.get("job_iterations", 0)) for peer in row_peers)
                        for it in range(used_iterations + 1, max_iter + 1):
                            curr_h = float(sheet.Rows(target_row).RowHeight or 15.0)
                            if pending:
                                if pending["item_id"] != item["id"]:
                                    raise ValueError("微调检查点与待处理项不一致")
                                new_h = pending["target_height"]
                            else:
                                new_h = round(curr_h + step_pt, 1)
                                pending = {
                                    "item_id": item["id"],
                                    "target_height": new_h,
                                    "iterations": int(item["iterations"]) + 1,
                                    "peer_iterations": {
                                        peer["id"]: int(peer["iterations"]) + 1 for peer in row_peers
                                    },
                                }
                                if job is not None:
                                    with STATE_LOCK:
                                        job["pending_adjustment"] = pending
                                        _persist_state()

                            # 先写入目标高度，再保存 Excel；重启后重放同一目标，不重复递增。
                            sheet.Rows(target_row).RowHeight = new_h
                            wb.Save()

                            _export_sheet_pdf(sheet, pdf_path)
                            touched_pdfs.add(pdf_path)

                            # 定位失败也要提交本轮已写入的行高，并清除检查点，避免续跑时重复加高。
                            located = _locate_item_for_fix(item, pdf_path, False)
                            with STATE_LOCK:
                                if located is not None:
                                    still_clip, new_img, new_bbox, resolved_page = located
                                    _update_item_pdf_location(item, new_img, new_bbox, resolved_page)
                                    item["status"] = "fixed" if not still_clip else "pending"
                                item["row_height"] = new_h
                                peer_iterations = pending.get("peer_iterations", {
                                    peer["id"]: int(peer["iterations"]) + 1 for peer in row_peers
                                })
                                peer_iterations[item["id"]] = pending["iterations"]
                                for peer in row_peers:
                                    peer["iterations"] = peer_iterations[peer["id"]]
                                    peer["job_iterations"] = it
                                _sync_row_height_in_state(wb_path_str, sheet_name, target_row, new_h)
                                if job is not None:
                                    job["pending_adjustment"] = None
                                _persist_state()
                            pending = None

                            if located is None:
                                skipped_ids.append(item["id"])
                                break
                            if not still_clip:
                                log_msg(
                                    f"  ✅ [{sheet_name} 行{target_row}] 第 {it} 次微调后已完整显示！(高度 -> {new_h}pt)"
                                )
                                break
                            else:
                                log_msg(
                                    f"  ⏳ [{sheet_name} 行{target_row}] 第 {it} 次微调 ({new_h}pt) 目标仍疑似压线"
                                )
                        else:
                            with STATE_LOCK:
                                item["status"] = "max_reached"
                                _persist_state()
                            log_msg(
                                f"  ⚠️ [{sheet_name} 行{target_row}] 已达同行共享微调上限 ({max_iter}次)，"
                                "停止加高，请核对行号或检测结果"
                            )

                    # 刷新该工作簿涉及到的所有整页 PDF 全景图
                    for touched_pdf in touched_pdfs:
                        _refresh_sheet_pages_in_state(touched_pdf)
                    # 中断可能发生在轮次提交之后、全页图刷新之前。
                    for pdf_path in {Path(item["pdf_path"]) for item in items} - touched_pdfs:
                        _refresh_sheet_pages_in_state(pdf_path)

                    # 完成 A1 定位与 100% 缩放保存；若所有处理项都因无法定位而跳过，则不改动 Excel 文件
                    untouched = bool(skipped_ids) and len(skipped_ids) == len(attempted_ids) and not touched_pdfs
                    if untouched:
                        log_msg(f"⏭️ {wb_path.name}: 选中项均无法可靠定位，未修改也未保存 Excel")
                    else:
                        normalize_workbook_view(excel, wb)
                        wb.Save()
                        log_msg(f"💾 已保存并完成 A1/100% 标准化: {wb_path.name}")
                    total_skipped += len(skipped_ids)

                    with STATE_LOCK:
                        wb_all_items = [
                            it for it in STATE["items"].values() if it["wb_path"] == wb_path_str
                        ]
                        wb_total_clips = len(wb_all_items)
                        wb_fixed_clips = sum(1 for it in wb_all_items if it["status"] == "fixed")
                        wb_accepted_clips = sum(it["status"] == "accepted" for it in wb_all_items)
                        wb_unfixed_clips = wb_total_clips - wb_fixed_clips - wb_accepted_clips

                    action = (
                        "未修改 Excel" if untouched
                        else f"已微调保存 (修复 {wb_fixed_clips}/{wb_total_clips} 处，A1/100% 已标准化)"
                    )
                    if skipped_ids:
                        action += f"；{len(skipped_ids)} 处无法可靠定位，已跳过并标记人工核对"
                    update_file_progress(
                        wb_path.name,
                        status="done" if wb_unfixed_clips == 0 else "needs_fix",
                        clip_detected=wb_total_clips,
                        clip_fixed=wb_fixed_clips,
                        clip_accepted=wb_accepted_clips,
                        clip_unfixed=wb_unfixed_clips,
                        normalized_a1=not untouched,
                        current_action=action,
                    )
                    with STATE_LOCK:
                        if job is not None:
                            job["completed"].extend(item["id"] for item in items)
                        _persist_state()
                except Exception as err:
                    log_msg(f"❌ 修复失败 {wb_path.name}: {err}")
                    update_file_progress(
                        wb_path.name,
                        status="error",
                        error=str(err),
                        current_action=f"修复失败: {err}",
                    )
                    raise
                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        done_msg = "🎉 所有选中项微调完毕！画布、放大镜截图与执行记录报告已更新。"
        if total_skipped:
            done_msg += f"其中 {total_skipped} 处无法可靠定位，已跳过并标记为人工核对。"
        log_msg(done_msg)
        finish_execution_batch(done_msg)
    finally:
        with STATE_LOCK:
            if not resume:
                STATE["busy"] = False
                _persist_state()


def preview_backup_sheets(folder_str: str) -> dict[str, Any]:
    """只读列出每个工作簿的全部工作表及其中将被精简的日期备份表；不修改文件，也不占用后台任务。"""
    folder, workbooks = list_workbooks(folder_str)
    files: list[dict[str, Any]] = []
    for wb_path in workbooks:
        entry: dict[str, Any] = {
            "name": wb_path.name, "sheet_count": 0, "sheets": [], "backups": [], "note": "",
        }
        try:
            all_names, backups = scan_workbook_sheets(wb_path)
        except Exception as err:
            entry["note"] = f"无法读取: {err}"
        else:
            entry["sheet_count"] = len(all_names)
            entry["sheets"] = all_names  # 保持工作簿中的原有顺序，便于对照保留与删除
            if backups and len(backups) >= len(all_names):
                entry["note"] = "全部工作表均为日期名，将跳过删除"
            else:
                entry["backups"] = backups
        files.append(entry)
    return {"folder": str(folder), "files": files}


def _clean_backup_sheets(get_excel: Callable[[], Any], wb_path: Path, backup_folder: Path) -> None:
    # 续跑时沿用本批次已记录的删除明细，避免保存后中断导致记录丢失
    previous = get_execution_report_snapshot()["files"].get(wb_path.name, {})
    deleted = list(previous.get("deleted_backups", []))

    def record(all_names: list[str], action: str) -> None:
        update_file_progress(
            wb_path.name, status="done", current_action=action, deleted_backups=deleted,
            all_sheets=list(dict.fromkeys(previous.get("all_sheets", []) + all_names)),
        )

    if wb_path.suffix.lower() != ".xls":
        # 先用 openpyxl 只读预检，没有备份表的文件不必用 Excel 打开
        try:
            all_names, backups = scan_workbook_sheets(wb_path)
        except Exception:
            backups = None  # 预检失败时交给 Excel 判断
        if backups == []:
            record(all_names, f"已删除 {len(deleted)} 个日期备份 Sheet" if deleted else "没有日期备份表，无需精简")
            return

    wb = get_excel().Workbooks.Open(str(wb_path.resolve()))
    try:
        all_names = [sheet.Name for sheet in wb.Sheets]
        backups = [name for name in all_names if is_backup_sheet(name)]
        if not backups:
            action = f"已删除 {len(deleted)} 个日期备份 Sheet" if deleted else "没有日期备份表，无需精简"
        elif len(backups) >= len(all_names):
            action = "全部工作表均为日期名，已跳过删除"
            log_msg(f"  ⚠️ {action}: {wb_path.name}")
        else:
            _backup_original(wb_path, backup_folder)
            for name in backups:
                wb.Sheets(name).Delete()
            deleted = list(dict.fromkeys(deleted + backups))
            # 先记录删除明细再保存，保存后中断也能在续跑时保留记录
            update_file_progress(wb_path.name, deleted_backups=deleted)
            wb.Save()
            action = f"已删除 {len(deleted)} 个日期备份 Sheet"
            log_msg(f"  {action}")
        record(all_names, action)
    finally:
        wb.Close(SaveChanges=False)


def _export_workbook_pdfs(
    get_excel: Callable[[], Any], wb_path: Path, pdf_folder: Path, used_pdf_paths: set[Path],
) -> list[Path]:
    wb = get_excel().Workbooks.Open(str(wb_path.resolve()), ReadOnly=True)
    try:
        file_stem = sanitize_filename(wb_path.stem)
        exported: list[Path] = []
        for sheet in wb.Sheets:
            if not _is_exportable_sheet(sheet):
                continue
            pdf_path = unique_pdf_path(
                pdf_folder, file_stem, sanitize_filename(sheet.Name), used_pdf_paths
            )
            _export_sheet_pdf(sheet, pdf_path)
            exported.append(pdf_path)
            update_file_progress(
                wb_path.name, exported_pdfs=len(exported),
                current_action=f"正在导出 PDF（已导出 {len(exported)} 份）...",
            )
        action = f"已导出 {len(exported)} 份 PDF"
        log_msg(f"  {action}")
        update_file_progress(
            wb_path.name, status="done", exported_pdfs=len(exported), current_action=action,
            all_sheets=[sheet.Name for sheet in wb.Sheets],
        )
        return exported
    finally:
        wb.Close(SaveChanges=False)


def _normalize_workbook_file(get_excel: Callable[[], Any], wb_path: Path, backup_folder: Path) -> None:
    _backup_original(wb_path, backup_folder)
    excel = get_excel()
    wb = excel.Workbooks.Open(str(wb_path.resolve()))
    try:
        normalize_workbook_view(excel, wb, on_error=log_msg)
        wb.Save()
        update_file_progress(
            wb_path.name, status="done", normalized_a1=True,
            current_action="已定位 A1 并设为 100% 缩放",
            all_sheets=[sheet.Name for sheet in wb.Sheets],
        )
    finally:
        wb.Close(SaveChanges=False)


def quick_task_workflow(kind: str) -> None:
    """主页单项批处理：以 Excel 文件为检查点，续跑时跳过已完成的文件。"""
    task_name = QUICK_TASK_NAMES[kind]
    with STATE_LOCK:
        job = STATE["job"]
        options = job["options"]
        initialized = "workbooks" in job
    if initialized:
        folder = Path(options["folder"])
        workbooks = [Path(path) for path in job["workbooks"]]
    else:
        folder, workbooks = list_workbooks(options["folder"])
        timestamp = f"{datetime.now():%Y%m%d_%H%M%S}"
        with STATE_LOCK:
            STATE["logs"].clear()
            options["folder"] = str(folder)
            if kind == "export_pdf":
                options["pdf_folder"] = str(folder / f"_exported_pdfs_{timestamp}")
            else:
                options["backup_folder"] = str(folder / f"_backup_{timestamp}")
            job["workbooks"] = [str(path) for path in workbooks]
            job["exported"] = []
            _persist_state()
        init_execution_batch(
            task_name=task_name,
            folder=folder,
            workbooks=workbooks,
            pdf_folder=Path(options["pdf_folder"]) if kind == "export_pdf" else None,
            backup_folder=None if kind == "export_pdf" else Path(options["backup_folder"]),
        )
    log_msg(f"{'继续' if initialized else '开始'}{task_name}：{folder}（共 {len(workbooks)} 个文件）")
    if kind == "export_pdf":
        Path(options["pdf_folder"]).mkdir(parents=True, exist_ok=True)

    excel = None

    def get_excel() -> Any:
        # 按需启动 Excel：精简 Sheet 时没有备份表的文件无需打开
        nonlocal excel
        if excel is None:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
        return excel

    # 已完成文件的 PDF 名称保持占用，续跑时不会被同名工作簿覆盖
    used_pdf_paths = {Path(path) for path in job["exported"]}
    pythoncom.CoInitialize()
    try:
        for wb_path in workbooks:
            if str(wb_path) in job["completed"]:
                continue
            log_msg(f"正在处理: {wb_path.name} ...")
            update_file_progress(
                wb_path.name, status="processing", current_action=f"正在执行{task_name}...", error="",
            )
            exported: list[Path] = []
            try:
                if kind == "clean_sheets":
                    _clean_backup_sheets(get_excel, wb_path, Path(options["backup_folder"]))
                elif kind == "export_pdf":
                    exported = _export_workbook_pdfs(
                        get_excel, wb_path, Path(options["pdf_folder"]), used_pdf_paths
                    )
                else:
                    _normalize_workbook_file(get_excel, wb_path, Path(options["backup_folder"]))
            except Exception as err:
                log_msg(f"❌ 处理文件出错 {wb_path.name}: {err}")
                update_file_progress(
                    wb_path.name, status="error", error=str(err), current_action=f"处理出错: {err}",
                )
                continue
            with STATE_LOCK:
                job["completed"].append(str(wb_path))
                job["exported"].extend(str(path) for path in exported)
                _persist_state()
    finally:
        try:
            if excel is not None:
                excel.Quit()
        finally:
            pythoncom.CoUninitialize()

    files = list(get_execution_report_snapshot()["files"].values())
    if kind == "clean_sheets":
        detail = f"共删除 {sum(len(f['deleted_backups']) for f in files)} 个日期备份 Sheet"
    elif kind == "export_pdf":
        detail = f"共导出 {sum(int(f['exported_pdfs']) for f in files)} 份 PDF"
    else:
        detail = f"{sum(bool(f['normalized_a1']) for f in files)} 个文件已完成 A1/100% 标准化"
    failed = sum(f["status"] == "error" for f in files)
    message = (
        f"⚠️ {task_name}结束：{detail}，{failed} 个文件失败" if failed
        else f"✅ {task_name}完成：{detail}"
    )
    log_msg(message)
    finish_execution_batch(message)


PROGRESS_FORMAT = "clear_handle_excel.progress"
PROGRESS_VERSION = 1
_WORKFLOW_FIELD_TYPES = {
    "folder": str, "pdf_folder": str, "backup_folder": str, "pages": dict, "items": dict,
    "logs": list, "settings": dict, "checked_files": dict,
}


def _version_or_none(path: str) -> dict[str, int] | None:
    try:
        return _file_version(Path(path))
    except FileNotFoundError:
        return None


def _referenced_workbooks(state: dict[str, Any]) -> set[str]:
    paths = {item["wb_path"] for item in state["items"].values()}
    paths.update(state["checked_files"])
    if state["job"]:
        paths.update(state["job"].get("workbooks", []))
    return paths


def export_progress() -> dict[str, Any]:
    """导出当前进度（工作台状态 + 执行记录），并记录相关 Excel 的文件版本供导入时比对。"""
    with STATE_LOCK:
        if STATE["busy"]:
            raise ValueError("任务执行中，请等任务结束后再导出进度")
        workflow_state = copy.deepcopy(STATE)
        report_state = get_execution_report_snapshot()
    report_state.pop("markdown", None)
    return {
        "format": PROGRESS_FORMAT,
        "version": PROGRESS_VERSION,
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "file_versions": {
            path: _version_or_none(path) for path in sorted(_referenced_workbooks(workflow_state))
        },
        "workflow": workflow_state,
        "report": report_state,
    }


def import_progress(data: Any) -> dict[str, Any]:
    """用导出的进度整体替换当前进度。

    导出后若相关 Excel 已被修改或缺失，取消中断任务的续跑点，避免把旧检查点（如目标行高）
    重放到新版本文件上；扫描结果与人工判断保留，人工确认按文件版本规则自动失效。
    """
    if not isinstance(data, dict) or data.get("format") != PROGRESS_FORMAT:
        raise ValueError("不是本工具导出的进度文件")
    if data.get("version") != PROGRESS_VERSION:
        raise ValueError(f"不支持的进度文件版本: {data.get('version')}")
    workflow_state = data.get("workflow")
    if (
        not isinstance(workflow_state, dict)
        or any(not isinstance(workflow_state.get(key), kind) for key, kind in _WORKFLOW_FIELD_TYPES.items())
        or not isinstance(workflow_state.get("job"), (dict, type(None)))
    ):
        raise ValueError("进度文件损坏：缺少工作台状态")
    file_versions = data.get("file_versions")
    if not isinstance(file_versions, dict):
        raise ValueError("进度文件损坏：缺少文件版本信息")

    imported = copy.deepcopy(workflow_state)
    imported["busy"] = False
    job = imported["job"]
    changed = sorted(
        Path(path).name for path, version in file_versions.items() if _version_or_none(path) != version
    )
    warnings: list[str] = []
    if changed:
        names = "、".join(changed[:5]) + (" 等" if len(changed) > 5 else "")
        warnings.append(f"导出后有 {len(changed)} 个 Excel 文件已修改或缺失：{names}")
    cancelled = False
    if job and job["status"] in {"running", "interrupted"}:
        if changed:
            job["status"] = "cancelled"
            job["pending_adjustment"] = None
            cancelled = True
            warnings.append("已取消上次中断任务的续跑，避免旧检查点覆盖新版本文件；扫描结果和人工判断仍保留")
        else:
            job["status"] = "interrupted"

    with STATE_LOCK:
        if STATE["busy"]:
            raise ValueError("任务执行中，请等任务结束后再导入进度")
        previous_report = get_execution_report_snapshot()
        replace_execution_report(data.get("report"))  # 先校验并替换执行记录，格式不对时什么都不改
        previous = copy.deepcopy(STATE)
        STATE.clear()
        STATE.update(imported)
        try:
            _persist_state()
        except Exception:
            STATE.clear()
            STATE.update(previous)
            replace_execution_report(previous_report)
            raise
        if cancelled:
            finish_execution_batch("已导入进度备份；导出后 Excel 有变化，已取消中断任务的续跑")
        log_msg(f"📥 已导入进度备份（导出于 {data.get('exported_at') or '未知时间'}）")
        for warning in warnings:
            log_msg(f"  ⚠️ {warning}")
        _expire_manual_confirmations()
        _persist_state()
        return {
            "exported_at": data.get("exported_at", ""),
            "items": len(STATE["items"]),
            "files": len(get_execution_report_snapshot()["files"]),
            "can_resume": bool(job and job["status"] == "interrupted"),
            "warnings": warnings,
        }
