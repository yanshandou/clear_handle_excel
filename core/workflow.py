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
    REVIEW_EMOJI,
    XL_SHEET_VISIBLE,
    XL_TYPE_PDF,
    as_bbox,
    col_index_to_letters,
    console_print,
    find_backup_sheets,
    find_clean_targets,
    find_contained_sheets,
    find_new_old_pairs,
    is_uncertain_name,
    review_marked_name,
    sanitize_filename,
    strip_new_prefix,
    strip_review_prefix,
    unique_pdf_path,
    workbook_rel_name,
)
from core.excel_ops import (
    CellProbe,
    close_saved_workbooks,
    extract_sheet_snapshot,
    mark_review_cells,
    open_workbooks_for_review,
    restore_review_cells,
    restore_font_colors,
    set_font_colors,
    match_excel_row,
    normalize_workbook_view,
    open_workbook_for_manual_review,
    pythoncom,
    scan_workbook_sheets,
    snapshot_cells,
    win32,
)
from core.hidden_text import clip_check_supported, collect_text_pieces, find_hidden_cells
from core.row_locator import (
    MAX_TAGGED_CELLS,
    build_anchors,
    candidate_cells,
    colored_span_cells,
    excel_color,
    locate_cell,
    tag_color,
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

MANUAL_CONFIRMED_REASON = "人工确认无需修复"
MANUALLY_FIXED_REASON = "人工修改后重新检测未再检出压线"
RESOLVED_STATUSES = {"accepted", "fixed"}
LATER_REVIEW_REASON = "已挂起，稍后统一人工复检（Excel 里标了黄色）"
LATER_REVIEW_FAILED_REASON = "统一复检仍有问题：请继续修改并保存后再复检，或标记无需修改"

# 主页上可单独执行的批处理功能（job kind -> 任务名）
QUICK_TASK_NAMES = {
    "clean_sheets": "精简 Sheet（删除备份表）",
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
    "settings": {"step_pt": 3.0, "max_iter": 6},
    "job": None,
    "checked_files": {},  # wb_path -> 最近一次完整检查未解决项时的文件版本
    "scan_progress": {},  # 独立于执行报告，避免并行写入覆盖扫描的总进度
    # 稍后人工复检：wb_path -> {"marks": [{sheet, addr, row, item_ids, state, fill, range, error}], "error"}
    # state: pending 等后台标黄 / marked 已标黄 / skipped 没能标黄（仍在队列里）
    "review_queue": {},
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


def scan_job_active() -> bool:
    job = STATE["job"] or {}
    return job.get("kind") == "scan" and job.get("status") in {"running", "interrupted"}


def file_available(path: str) -> bool:
    """调用方持有 STATE_LOCK；扫描尚未提交的文件不能筛选、截图或写回。"""
    return not scan_job_active() or path in STATE["job"]["completed"]


def _check_decision_item(item: dict[str, Any]) -> None:
    if STATE["busy"] and (STATE["job"] or {}).get("kind") != "scan":
        raise ValueError("任务正在执行，请结束后再进行人工判断")
    if not file_available(item["wb_path"]):
        raise ValueError("此文件尚未扫描完成，请等待扫描提交结果")
    from core import autofix
    autofix.check_row_not_writing(item)


def queue_job(kind: str, options: dict[str, Any] | None = None) -> bool:
    with STATE_LOCK:
        if STATE["busy"]:
            return False
        if kind == "review_recheck" and not STATE["review_queue"]:
            raise ValueError("没有挂起的稍后人工复检")
        if kind == "review_open" and not STATE["review_queue"]:
            raise ValueError("没有挂起的稍后人工复检")
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
            if kind == "scan":
                STATE["scan_progress"] = {"total": 0, "completed": 0, "status": "running"}
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
            scan_folder_workflow(
                options["folder"], options["delete_backups"], full=options.get("full", False), resume=True,
            )
            expected = job.get("workbooks", [])
        elif job["kind"] == "manual_open" and job["completed"]:
            expected = options["selected_ids"]
        elif job["kind"] in {"manual_open", "recheck"}:
            manual_review_workflow(options["selected_ids"], open_excel=job["kind"] == "manual_open")
            expected = options["selected_ids"]
        elif job["kind"] == "review_open":
            review_open_workflow()
            expected = []
        elif job["kind"] == "review_recheck":
            review_recheck_workflow()
            expected = job["options"]["selected_ids"]
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
            if STATE["job"] and STATE["job"]["kind"] == "scan":
                STATE["scan_progress"].update(
                    total=len(STATE["job"].get("workbooks", [])),
                    completed=len(STATE["job"]["completed"]), status=STATE["job"]["status"],
                )
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
    report_snap = get_execution_report_snapshot()
    with STATE_LOCK:
        available = {path for path in {it["wb_path"] for it in STATE["items"].values()} if file_available(path)}
        return {
            "busy": STATE["busy"],
            "folder": STATE["folder"],
            "pdf_folder": STATE["pdf_folder"],
            "logs": list(STATE["logs"][-40:]),
            "pages": {key: copy.deepcopy(p) for key, p in STATE["pages"].items() if p["wb_path"] in available},
            "items": {key: copy.deepcopy(it) for key, it in STATE["items"].items() if it["wb_path"] in available},
            "scan_progress": copy.deepcopy(STATE["scan_progress"]),
            "review_queue": copy.deepcopy(STATE["review_queue"]),
            "settings": dict(STATE["settings"]),
            "can_resume": bool(STATE["job"] and STATE["job"]["status"] == "interrupted"),
            "job": copy.deepcopy(STATE["job"]),
            "progress": _progress_summary(report_snap),
        }


def get_overview_snapshot() -> dict[str, Any]:
    """主页轮询用的轻量快照：不含整页截图与问题项详情。"""
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
    *, confirm_locations: bool = True,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """从当前 PDF 构建新快照；不沿用旧文字、坐标或问题状态。"""
    pages: dict[str, Any] = {}
    items: dict[str, Any] = {}
    sheet_snapshot = extract_sheet_snapshot(sheet)
    generation = uuid4().hex
    wb_label = workbook_rel_name(wb_path, STATE.get("folder"))
    with fitz.open(pdf_path) as doc:
        if not len(doc):
            raise ValueError(f"导出的 PDF 没有页面: {pdf_path}")
        text_spans = collect_text_spans(doc)
        probe = CellProbe(sheet)
        cells = snapshot_cells(sheet_snapshot)
        pieces = collect_text_pieces(doc)
        anchors = build_anchors(cells, pieces)

        def page_entry(page_index: int) -> str:
            page_id = f"page_{generation}_{page_index}"
            if page_id not in pages:
                image, width, height = render_page_overview(doc[page_index])
                pages[page_id] = {
                    "id": page_id, "wb_path": str(wb_path), "wb_name": wb_label,
                    "sheet_name": sheet_name, "pdf_path": str(pdf_path),
                    "page_index": page_index, "total_pages": len(doc),
                    "page_width": width, "page_height": height, "image": image,
                }
            return page_id

        candidates: dict[str, list[tuple[int, int]]] = {}
        for page_index in range(len(doc)):
            page = doc[page_index]
            clips = find_clipped_spans_on_page(page)
            if not clips:
                continue
            page_id = page_entry(page_index)
            for index, clip in enumerate(clips, start=1):
                located = locate_cell(cells, anchors, clip["text"], page_index, clip["bbox"], probe.row_height)
                if located is not None:
                    row, col = located
                    address, height = f"{col_index_to_letters(col)}{row}", probe.row_height(row)
                else:
                    # 这一页没有能确定位置的文字时，才按页面位置比例估算
                    row, address, height = match_excel_row(
                        sheet_snapshot, clip["text"], page_index, len(doc),
                        clip["bbox"][1] / max(1.0, page.rect.height),
                    )
                item_id = f"item_{generation}_{len(items)}"
                occurrence = locate_text_occurrence(text_spans, page_index, clip["text"], clip["bbox"])
                candidates[item_id] = candidate_cells(cells, clip["text"], row)
                items[item_id] = {
                    "id": item_id, "seq": start_seq + len(items), "kind": "clip", "page_id": page_id,
                    "index_on_page": index, "wb_path": str(wb_path), "wb_name": wb_label,
                    "sheet_name": sheet_name, "pdf_path": str(pdf_path),
                    "page_index": page_index, "text": clip["text"], "bbox": clip["bbox"],
                    "occurrence": list(occurrence) if occurrence else None,
                    "excel_row": row, "cell_addr": address, "row_height": height,
                    "image": render_crop_base64(page, clip["bbox"], highlight=True),
                    "status": "pending", "iterations": 0,
                }

        _confirm_clip_cells(sheet, pdf_path, items, candidates, probe, allow_export=confirm_locations)

        # 隐藏截断：单元格完整内容在 PDF 里没显示全（外观上看不出来）
        hidden = find_hidden_cells(cells, pieces, probe.text, probe.printed)
        per_page: dict[int | None, int] = {}
        for found in hidden:
            page_index = found["page"]
            per_page[page_index] = per_page.get(page_index, 0) + 1
            address = f"{col_index_to_letters(found['col'])}{found['row']}"
            item_id = f"item_{generation}_{len(items)}"
            items[item_id] = {
                "id": item_id, "seq": start_seq + len(items), "kind": "hidden",
                "page_id": page_entry(page_index) if page_index is not None else None,
                "index_on_page": per_page[page_index], "wb_path": str(wb_path), "wb_name": wb_label,
                "sheet_name": sheet_name, "pdf_path": str(pdf_path),
                "page_index": page_index, "text": found["text"], "bbox": found["bbox"],
                "visible_text": found["visible"], "hidden_reason": found["reason"],
                "occurrence": ["cell", address],
                "excel_row": found["row"], "cell_addr": address, "row_height": probe.row_height(found["row"]),
                "image": (
                    render_crop_base64(doc[page_index], found["bbox"], highlight=True)
                    if page_index is not None else ""
                ),
                "status": "pending", "iterations": 0,
            }
    return pages, items


# 检测规则变化时加一：之前按旧规则查过的文件即使没改过，下次扫描也要整份重新检查
DETECTOR_VERSION = 6


def _confirm_clip_cells(
    sheet: Any, pdf_path: Path, items: dict[str, Any],
    candidates: dict[str, list[tuple[int, int]]], probe: CellProbe,
    *, allow_export: bool = True,
) -> None:
    """切字项逐一确认是哪一格：只有一个候选格子时直接确定；有多个时把候选格子的字体
    改成各不相同的颜色（不保存工作簿）再导出一次 PDF，读切字的颜色就知道是哪一格。
    认不出时（条件格式改了颜色等）保留按位置推算的结果，并标为定位未确认。"""
    pending: list[str] = []
    for item_id, cands in candidates.items():
        item = items[item_id]
        if len(cands) == 1:
            _set_item_cell(item, *cands[0], probe)
            item["location_confirmed"] = True
        else:
            item["location_confirmed"] = False
            pending.append(item_id)
    if not allow_export:
        return  # 单项复检只分析已经导出的 PDF，歧义位置保持“定位未确认”
    tagged = list(dict.fromkeys(rc for item_id in pending for rc in candidates[item_id]))[:MAX_TAGGED_CELLS]
    if not tagged:
        return
    temp_pdf = pdf_path.with_name(f"~定位_{uuid4().hex}.pdf")
    originals = set_font_colors(sheet, {rc: excel_color(tag_color(i)) for i, rc in enumerate(tagged, start=1)})
    try:
        _export_sheet_pdf(sheet, temp_pdf)
        with fitz.open(temp_pdf) as doc:
            found = colored_span_cells(
                doc,
                [(items[i]["page_index"], as_bbox(items[i]["bbox"]), items[i]["text"]) for i in pending],
                tagged,
            )
    except Exception as err:
        log_msg(f"  ⚠️ [{sheet.Name}] 上色确认切字位置失败，沿用按位置推算的行号: {err}")
        return
    finally:
        restore_font_colors(sheet, originals)
        temp_pdf.unlink(missing_ok=True)
    for item_id, cell in zip(pending, found):
        if cell is not None:
            _set_item_cell(items[item_id], *cell, probe)
            items[item_id]["location_confirmed"] = True


def _set_item_cell(item: dict[str, Any], row: int, col: int, probe: CellProbe) -> None:
    item["excel_row"] = row
    item["cell_addr"] = f"{col_index_to_letters(col)}{row}"
    item["row_height"] = probe.row_height(row)


def _file_version(path: Path) -> dict[str, int]:
    stat = path.stat()
    return {"mtime_ns": stat.st_mtime_ns, "size": stat.st_size, "detector": DETECTOR_VERSION}


def _sync_problem_counts(path: str, action: str) -> None:
    items = [item for item in STATE["items"].values() if item["wb_path"] == path]
    fixed = sum(item["status"] == "fixed" for item in items)
    accepted = sum(item["status"] == "accepted" for item in items)
    pending = len(items) - fixed - accepted
    update_file_progress(
        workbook_rel_name(Path(path), STATE.get("folder")),
        status="needs_fix" if pending else "done",
        clip_detected=len(items), clip_fixed=fixed, clip_accepted=accepted,
        clip_unfixed=pending, current_action=action, error="",
    )


def _same_text_problem(a: dict[str, Any], b: dict[str, Any]) -> bool:
    # 旧版记录没有 kind，都是压线切字
    return (a["wb_path"], a["sheet_name"], a.get("kind") or "clip", a["text"]) == (
        b["wb_path"], b["sheet_name"], b.get("kind") or "clip", b["text"]
    )


def _problem_record(item: dict[str, Any]) -> dict[str, Any]:
    """完整重新扫描前记下旧问题项（不含截图等大字段），用于对应同一处问题"""
    return {
        key: copy.deepcopy(item.get(key))
        for key in ("id", "seq", "kind", "wb_path", "sheet_name", "text", "occurrence", "status",
                    "manual_confirmation", "triage_keep", "later_review")
    }


def _match_previous(new_items: dict[str, Any], previous: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """为重新检测出的问题找到对应的旧问题项，返回 {新问题 id: 旧问题项}（一一对应）。

    同一处 = 同一工作簿、工作表、文字，且在同列相同文字中的出现顺序与总数都一致；
    不要求页码和坐标相同（改行高、重新分页后位置会变）。相同文字数量变化等
    无法可靠对应的情况视为新问题，不猜测。
    """
    remaining = list(previous)
    matched: dict[str, dict[str, Any]] = {}
    for result in new_items.values():
        same_text = [old for old in remaining if _same_text_problem(old, result)]
        if not same_text:
            continue
        occurrence = result.get("occurrence")
        matches = [
            old for old in same_text
            if occurrence is not None and old.get("occurrence") is not None
            and list(old["occurrence"]) == list(occurrence)
        ]
        if not matches and (occurrence is None or all(old.get("occurrence") is None for old in same_text)):
            # 旧版扫描结果或无法确定出现序号时，只有该文字在工作表里仅此一处才对应
            peers = [other for other in new_items.values() if _same_text_problem(other, result)]
            if len(same_text) == 1 and len(peers) == 1:
                matches = same_text
        if len(matches) != 1:
            continue
        remaining.remove(matches[0])
        matched[result["id"]] = matches[0]
    return matched


def _merge_redetected(
    new_items: dict[str, Any], previous: list[dict[str, Any]], *, mark_new: bool, keep_unmatched: bool,
) -> dict[str, Any]:
    """合并重新检测结果与旧问题项，返回以问题 id 为键的结果。

    - 对应上的：沿用旧 id 和序号（页面上的筛选记录不丢）；已跳过的仍跳过，其余仍压线则待人工处理
    - 新检出、对应不上的：mark_new 时标为「新发现」
    - 旧问题未再检出（keep_unmatched 时保留记录）：未解决的记为已修复，已跳过 / 已修复的保持不变
    """
    matched = _match_previous(new_items, previous)
    merged: dict[str, Any] = {}
    for item in new_items.values():
        old = matched.get(item["id"])
        item["is_new"] = mark_new and old is None
        if old is not None:
            item["id"], item["seq"] = old["id"], old["seq"]
            item["triage_keep"] = bool(old.get("triage_keep"))
            if old["status"] == "accepted":
                item["status"] = "accepted"
                item["manual_confirmation"] = copy.deepcopy(old["manual_confirmation"])
                item["review_reason"] = MANUAL_CONFIRMED_REASON
            elif old.get("later_review"):
                # 还挂在稍后人工复检里，等统一复检，不回到逐个处理
                item.update(status="manual_review", manual_waiting=True, later_review=True,
                            review_reason=LATER_REVIEW_REASON)
        merged[item["id"]] = item
    if keep_unmatched:
        matched_ids = {old["id"] for old in matched.values()}
        for old in previous:
            if old["id"] in matched_ids:
                continue
            record = copy.deepcopy(old)
            record["page_id"] = None  # 原页面已无此处压线，只在卡片中保留记录
            record["is_new"] = False
            if record["status"] not in RESOLVED_STATUSES:
                record["status"] = "fixed"
                record["review_reason"] = MANUALLY_FIXED_REASON
            merged[record["id"]] = record
    return merged


def _set_item_confirmation(item: dict[str, Any], accepted: bool) -> None:
    if accepted:
        # 不绑定文件版本：保存文件不会让确认失效，重新检测时按同一处文字保留
        item["manual_confirmation"] = {
            "confirmed_at": datetime.now().isoformat(timespec="seconds"),
            "previous_status": item["status"],
        }
        item["status"] = "accepted"
        item["review_reason"] = MANUAL_CONFIRMED_REASON
    else:
        item["status"] = item["manual_confirmation"]["previous_status"]
        item["review_reason"] = "人工确认已撤销，重新纳入待处理"
        del item["manual_confirmation"]


def _commit_confirmations(previous_items: dict[str, dict[str, Any]], action: str) -> None:
    """保存已修改的人工确认（一次写库）；会取消中断任务的续跑点，保存失败时整体回滚。"""
    previous_job = copy.deepcopy(STATE["job"])
    if STATE["job"] and STATE["job"]["status"] == "interrupted" and not scan_job_active():
        STATE["job"]["status"] = "cancelled"
        STATE["job"]["pending_adjustment"] = None
    try:
        _persist_state()
    except Exception:
        for item_id, previous in previous_items.items():
            STATE["items"][item_id].clear()
            STATE["items"][item_id].update(previous)
        STATE["job"] = previous_job
        raise
    for path in dict.fromkeys(STATE["items"][item_id]["wb_path"] for item_id in previous_items):
        _sync_problem_counts(path, action)
    if previous_job and previous_job["status"] == "interrupted" and previous_job.get("kind") != "scan":
        finish_execution_batch("已取消中断任务的续跑，按人工判断更新问题状态")


def set_manual_confirmation(item_id: str, accepted: bool) -> None:
    with STATE_LOCK:
        item = STATE["items"].get(item_id)
        if item is None:
            raise ValueError("问题项已更新，请刷新页面后重试")
        _check_decision_item(item)
        if accepted and item["status"] == "fixed":
            raise ValueError("此项已自动修复，无需人工确认")
        if accepted and item["status"] == "accepted":
            raise ValueError("此项已人工确认，如需重新判断请先撤销")
        if not accepted and item["status"] != "accepted":
            raise ValueError("此项没有有效的人工确认")
        previous = {item_id: copy.deepcopy(item)}
        _set_item_confirmation(item, accepted)
        _commit_confirmations(previous, item["review_reason"])
        log_msg(f"{item['review_reason']}: [{item['wb_name']} / {item['sheet_name']} / {item['cell_addr']}]")


def apply_skip_decisions(decisions: dict[str, bool]) -> dict[str, Any]:
    """批量设置「跳过」（即人工确认无需修复），一次写库。

    与当前状态相同的决定直接忽略，页面可以把连续快速点选的结果合并后一次提交。
    """
    with STATE_LOCK:
        if STATE["busy"] and (STATE["job"] or {}).get("kind") != "scan":
            raise ValueError("任务正在执行，请结束后再标记跳过")
        previous: dict[str, dict[str, Any]] = {}
        errors: list[str] = []
        for iid in decisions:
            if iid in STATE["items"]:
                _check_decision_item(STATE["items"][iid])
        for item_id, skip in decisions.items():
            item = STATE["items"].get(item_id)
            if item is None:
                errors.append("部分问题项已重新检测，请在最新结果上重新标记")
                continue
            if skip == (item["status"] == "accepted"):
                continue
            if skip and item["status"] == "fixed":
                errors.append(f"{item['wb_name']} / {item['cell_addr']} 已自动修复，无需跳过")
                continue
            previous[item_id] = copy.deepcopy(item)
            _set_item_confirmation(item, skip)
        if previous:
            skipped = sum(STATE["items"][item_id]["status"] == "accepted" for item_id in previous)
            action = f"人工筛选：跳过 {skipped} 处，撤销跳过 {len(previous) - skipped} 处"
            _commit_confirmations(previous, action)
            log_msg(action)
        return {"changed": len(previous), "errors": list(dict.fromkeys(errors))}


def apply_keep_decisions(decisions: dict[str, bool]) -> dict[str, Any]:
    """持久化保留／撤销保留；后台只处理同一行所有未解决项都已保留的行。"""
    with STATE_LOCK:
        previous: dict[str, dict[str, Any]] = {}
        for iid in decisions:
            item = STATE["items"].get(iid)
            if item is None:
                raise ValueError("问题项已更新，请刷新页面后重新筛选")
            _check_decision_item(item)
            if item["status"] in RESOLVED_STATUSES:
                raise ValueError("此项已处理，不能更改保留判断")
        for iid, keep in decisions.items():
            item = STATE["items"][iid]
            if bool(item.get("triage_keep")) != keep:
                previous[iid] = copy.deepcopy(item)
                item["triage_keep"] = keep
        if previous:
            _commit_confirmations(previous, "已保存人工筛选，保留的完整行交给后台准备截图")
    from core import autofix
    autofix.sync_groups()
    autofix._WAKE.set()
    return {"changed": len(previous), "errors": []}


def _select_changed_problem_items() -> list[str]:
    with STATE_LOCK:
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
    """解析文件夹路径并递归列出其中及所有子文件夹下的 Excel 文件（跳过 ~$ 临时锁文件）"""
    folder = Path(folder_str.strip().strip('"').strip("'")).expanduser().resolve()
    if not folder.is_dir():
        raise ValueError(f"文件夹不存在: {folder}")
    workbooks = sorted(
        (
            p
            for p in folder.rglob("*")
            if p.is_file()
            and p.suffix.lower() in EXCEL_EXTENSIONS
            and not p.name.startswith("~$")
        ),
        key=lambda p: workbook_rel_name(p, folder),
    )
    if not workbooks:
        raise ValueError("未找到 Excel 文件")
    return folder, workbooks


def _exportable_sheets(wb: Any) -> list[Any]:
    """可见、且不是変更履歴或备份表的工作表才导出 PDF（备份表按整个工作簿判断）"""
    sheets = list(wb.Sheets)
    backup_names = set(find_backup_sheets([sheet.Name for sheet in sheets]))
    return [
        sheet for sheet in sheets
        if sheet.Name not in EXCLUDE_PDF_SHEETS
        and sheet.Name not in backup_names
        and sheet.Visible == XL_SHEET_VISIBLE
    ]


def _export_sheet_pdf(sheet: Any, pdf_path: Path) -> None:
    sheet.ExportAsFixedFormat(
        Type=XL_TYPE_PDF,
        Filename=str(pdf_path.resolve()),
        Quality=0,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False,
    )


def _scan_plan(workbooks: list[Path], *, rescan_all: bool) -> dict[str, dict[str, Any]]:
    """决定每个文件本次怎么检查：full 整个文件 / sheets 只查还有问题的工作表 / skip 沿用上次结果"""
    plan: dict[str, dict[str, Any]] = {}
    for wb_path in workbooks:
        path = str(wb_path)
        items = [item for item in STATE["items"].values() if item["wb_path"] == path]
        scanned = not rescan_all and (path in STATE["checked_files"] or bool(items))
        if not scanned:
            plan[path] = {"mode": "full", "scanned": False}
        elif _file_version(wb_path) == STATE["checked_files"].get(path):
            plan[path] = {"mode": "skip", "scanned": True}
        elif (STATE["checked_files"].get(path) or {}).get("detector") != DETECTOR_VERSION:
            plan[path] = {"mode": "full", "scanned": True}
        else:
            # 改过的文件：还有问题的只查问题工作表，之前已没问题的整个文件重新检查
            sheets = sorted({item["sheet_name"] for item in items if item["status"] not in RESOLVED_STATUSES})
            plan[path] = (
                {"mode": "sheets", "sheets": sheets, "scanned": True} if sheets
                else {"mode": "full", "scanned": True}
            )
    return plan


def _scan_one_workbook(
    excel: Any,
    wb_path: Path,
    entry: dict[str, Any],
    previous: list[dict[str, Any]],
    *,
    job: dict[str, Any] | None,
    resumed: bool,
    rescan_all: bool,
    delete_backups: bool,
    used_pdf_paths: set[Path],
) -> None:
    """检查一个工作簿；所有工作表检测完成后才一次性替换该文件的旧结果，中途失败不影响旧记录。"""
    path = str(wb_path)
    wb_label = workbook_rel_name(wb_path, STATE.get("folder"))
    sheets_only = set(entry["sheets"]) if entry["mode"] == "sheets" else None
    log_msg(
        f"正在检查: {wb_label}"
        + (f"（只查有问题的工作表：{'、'.join(entry['sheets'])}）" if sheets_only else " ...")
    )
    update_file_progress(wb_label, status="processing", current_action="正在打开工作簿...", error="")
    wb = excel.Workbooks.Open(str(wb_path.resolve()))
    try:
        all_names = [s.Name for s in wb.Sheets]
        deleted_bkp_list: list[str] = []
        if resumed:
            previous_report = get_execution_report_snapshot()["files"].get(wb_label, {})
            deleted_bkp_list = list(previous_report.get("deleted_backups", []))
            all_names = list(dict.fromkeys(previous_report.get("all_sheets", []) + all_names))
        current_names = {s.Name for s in wb.Sheets}
        bkp_names = [n for n in find_backup_sheets(all_names) if n in current_names]
        if delete_backups and bkp_names:
            if len(bkp_names) >= len(all_names):
                log_msg(f"  ⚠️ 跳过删除备份表（全部工作表均为备份表）: {wb_label}")
            else:
                for bn in bkp_names:
                    wb.Sheets(bn).Delete()
                deleted_bkp_list = list(dict.fromkeys(deleted_bkp_list + bkp_names))
                log_msg(f"  已清理 {len(bkp_names)} 个备份 Sheet")
                update_file_progress(wb_label, all_sheets=all_names, deleted_backups=deleted_bkp_list)
                wb.Save()
        update_file_progress(
            wb_label, all_sheets=all_names, deleted_backups=deleted_bkp_list,
            current_action="正在导出 PDF 并检测切字...",
        )

        scan_version = _file_version(wb_path)
        file_stem = sanitize_filename(strip_review_prefix(wb_path.stem))
        sheets = [s for s in _exportable_sheets(wb) if sheets_only is None or s.Name in sheets_only]
        rescanning_names = {s.Name for s in sheets}
        with STATE_LOCK:
            for p in STATE["pages"].values():
                if p.get("wb_path") == path and (sheets_only is None or p.get("sheet_name") in rescanning_names):
                    used_pdf_paths.discard(Path(p["pdf_path"]))
            next_seq = max((it["seq"] for it in STATE["items"].values()), default=0) + 1
        file_pages: dict[str, Any] = {}
        file_items: dict[str, Any] = {}
        for sheet in sheets:
            s_name = sheet.Name
            update_file_progress(wb_label, current_action=f"正在导出并检测 [{s_name}]...")
            pdf_path = unique_pdf_path(wb_path.parent, file_stem, sanitize_filename(s_name), used_pdf_paths)
            try:
                _export_sheet_pdf(sheet, pdf_path)
            except Exception as e:
                log_msg(f"  导出 PDF 失败 [{s_name}]: {e}")
                raise
            pages, items = _detect_sheet_results(sheet, wb_path, s_name, pdf_path, next_seq)
            next_seq += len(items)
            file_pages.update(pages)
            file_items.update(items)
        if _file_version(wb_path) != scan_version:
            raise ValueError(f"扫描期间文件已修改，请重新检查: {wb_label}")
    finally:
        wb.Close(SaveChanges=False)

    merged = _merge_redetected(file_items, previous, mark_new=entry["scanned"], keep_unmatched=not rescan_all)
    fixed_now = sum(
        old["status"] not in RESOLVED_STATUSES and merged.get(old["id"], {}).get("status") == "fixed"
        for old in previous
    )
    with STATE_LOCK:
        revision = uuid4().hex
        for item in merged.values():
            item["autofix_revision"] = revision
        for collection in ("pages", "items"):
            for key, value in list(STATE[collection].items()):
                if value["wb_path"] == path and (sheets_only is None or value["sheet_name"] in sheets_only):
                    del STATE[collection][key]
        STATE["pages"].update(file_pages)
        STATE["items"].update(merged)
        STATE["checked_files"][path] = scan_version
        if job is not None:
            job["completed"].append(path)
            STATE["scan_progress"].update(completed=len(job["completed"]))
        else:
            STATE["scan_progress"]["completed"] = STATE["scan_progress"].get("completed", 0) + 1
        _persist_state()
        wb_items = [item for item in STATE["items"].values() if item["wb_path"] == path]
        unresolved = sum(item["status"] not in RESOLVED_STATUSES for item in wb_items)
        new_count = sum(bool(item.get("is_new")) for item in wb_items if item["status"] not in RESOLVED_STATUSES)
        parts = [f"检查 {len(sheets)} 个工作表"]
        parts.append(
            f"待处理 {unresolved} 处" + (f"（新发现 {new_count} 处）" if new_count else "")
            if unresolved else "无待处理切字"
        )
        if fixed_now:
            parts.append(f"本次确认已修复 {fixed_now} 处")
        _sync_problem_counts(path, "，".join(parts))
    update_file_progress(wb_label, exported_pdfs=len(sheets))


def scan_folder_workflow(
    folder_str: str, delete_backups: bool, *, full: bool = False, resume: bool = False,
) -> None:
    """扫描切字。同一文件夹再次扫描时增量检查：新文件完整检查，未修改的文件沿用上次结果，
    修改过的文件只查还有问题的工作表（之前已没问题的整个文件重新检查）。
    full=True 时忽略上次记录完整重新扫描，跳过的位置仍按同一处文字保留。"""
    with STATE_LOCK:
        STATE["busy"] = True
        job = STATE["job"] if resume else None
        initialized = bool(job and "workbooks" in job)

    try:
        if initialized:
            assert job is not None
            folder = Path(STATE["folder"])
            workbooks = [Path(path) for path in job["workbooks"]]
            # 旧版本留下的扫描任务没有检查计划，按完整扫描继续
            plan = job.get("plan") or {str(path): {"mode": "full", "scanned": False} for path in workbooks}
            rescan_all = job.get("rescan_all", "plan" not in job)
            stash = job.get("stash", job.get("confirmed", []))
        else:
            folder, workbooks = list_workbooks(folder_str)
            with STATE_LOCK:
                rescan_all = full or str(folder) != STATE["folder"]
                stash = []
                if rescan_all:
                    # 完整重新扫描会清空旧结果；先记下旧问题（存入任务以便续跑），检测到同一处时沿用编号与跳过状态
                    stash = [_problem_record(item) for item in STATE["items"].values()]
                    STATE["pages"].clear()
                    STATE["items"].clear()
                    STATE["checked_files"].clear()
                STATE["logs"].clear()
                plan = _scan_plan(workbooks, rescan_all=rescan_all)
                STATE.update(folder=str(folder), pdf_folder=str(folder), backup_folder="")
                if job is not None:
                    job.update(
                        workbooks=[str(path) for path in workbooks], plan=plan,
                        rescan_all=rescan_all, stash=stash,
                    )
                STATE["scan_progress"] = {
                    "total": len(workbooks), "completed": 0, "status": "running",
                }
            init_execution_batch(
                task_name="Web 切字扫描（完整）" if rescan_all else "Web 切字扫描（增量）",
                folder=folder,
                workbooks=workbooks,
                pdf_folder=folder,
                backup_folder=None,
            )
            with STATE_LOCK:
                _persist_state()
        used_pdf_paths = {Path(p["pdf_path"]) for p in STATE["pages"].values()}
        skipped = [path for path, entry in plan.items() if entry["mode"] == "skip"]
        if not clip_check_supported():
            log_msg("⚠️ 当前 PyMuPDF 版本较旧，查不出被单元格遮住的文字；请运行 pip install -U pymupdf 后重新扫描")
        log_msg(
            f"{'完整扫描' if rescan_all else '增量扫描'}：共 {len(workbooks)} 个文件，"
            f"{len(workbooks) - len(skipped)} 个需要检查，{len(skipped)} 个未修改、沿用上次结果"
        )

        excel = None
        pythoncom.CoInitialize()
        try:
            for wb_path in workbooks:
                path = str(wb_path)
                if job is not None and path in job["completed"]:
                    continue
                entry = plan[path]
                if entry["mode"] == "skip":
                    with STATE_LOCK:
                        _sync_problem_counts(path, "文件未修改，沿用上次检查结果")
                        if job is not None:
                            job["completed"].append(path)
                            STATE["scan_progress"]["completed"] = len(job["completed"])
                        else:
                            STATE["scan_progress"]["completed"] += 1
                        _persist_state()
                    continue
                if excel is None:
                    # 按需启动 Excel：全部文件都未修改时不必打开
                    excel = win32.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                if rescan_all:
                    previous = [record for record in stash if record["wb_path"] == path]
                else:
                    with STATE_LOCK:
                        previous = [
                            copy.deepcopy(item) for item in STATE["items"].values()
                            if item["wb_path"] == path
                            and (entry["mode"] == "full" or item["sheet_name"] in entry["sheets"])
                        ]
                try:
                    _scan_one_workbook(
                        excel, wb_path, entry, previous,
                        job=job, resumed=initialized, rescan_all=rescan_all,
                        delete_backups=delete_backups, used_pdf_paths=used_pdf_paths,
                    )
                except Exception as err:
                    wb_label = workbook_rel_name(wb_path, folder)
                    log_msg(f"❌ 处理文件出错 {wb_label}: {err}")
                    update_file_progress(
                        wb_label,
                        status="error",
                        error=str(err),
                        current_action=f"处理出错: {err}",
                    )
        finally:
            try:
                if excel is not None:
                    excel.Quit()
            finally:
                pythoncom.CoUninitialize()

        with STATE_LOCK:
            items = list(STATE["items"].values())
            unresolved = [item for item in items if item["status"] not in RESOLVED_STATUSES]
            new_total = sum(bool(item.get("is_new")) for item in unresolved)
            fixed_total = sum(item["status"] == "fixed" for item in items)
            accepted_total = sum(item["status"] == "accepted" for item in items)
        summary_msg = (
            f"✅ 扫描完成：检查 {len(workbooks) - len(skipped)} 个文件"
            + (f"，{len(skipped)} 个未修改沿用上次结果" if skipped else "")
            + f"。当前待处理 {len(unresolved)} 处" + (f"（新发现 {new_total} 处）" if new_total else "")
            + f"，已修复 {fixed_total} 处，已跳过 {accepted_total} 处。"
        )
        log_msg(summary_msg)
        finish_execution_batch(summary_msg)
    finally:
        with STATE_LOCK:
            if not resume:
                STATE["busy"] = False
                progress = STATE["scan_progress"]
                progress["status"] = "completed" if progress.get("completed") == progress.get("total") else "interrupted"
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
        from core import autofix
        autofix.pause_previews()
        autofix.wait_for_writes()
        item = items[0]
        height = open_workbook_for_manual_review(
            Path(item["wb_path"]), item["sheet_name"], item["cell_addr"], int(item["excel_row"])
        )
        autofix.pause_previews()
        with STATE_LOCK:
            _sync_row_height_in_state(item["wb_path"], item["sheet_name"], int(item["excel_row"]), height)
            for peer in STATE["items"].values():
                if (peer["wb_path"], peer["sheet_name"], int(peer["excel_row"])) != (
                    item["wb_path"], item["sheet_name"], int(item["excel_row"])
                ) or peer["status"] in RESOLVED_STATUSES:
                    continue
                peer["status"] = "manual_review"
                peer["manual_waiting"] = True
                peer.pop("manual_recheck_token", None)
                peer["review_reason"] = "请人工核对自动推测的行号；调整并保存后重新导出复检"
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

    # 单项复检只覆盖当前工作表；文件基线仅在所有问题表都检查过时推进。
    with STATE_LOCK:
        sheets = {(item["wb_path"], item["sheet_name"]) for item in items}
        job["options"].setdefault("manual_item_ids", [item["id"] for item in items])
        job["options"].setdefault("manual_rows", [
            [item["wb_path"], item["sheet_name"], int(item["excel_row"])] for item in items
        ])
        selected = set(selected_ids)
        ids = [
            item["id"] for item in STATE["items"].values()
            if item["id"] in selected or (
                (item["wb_path"], item["sheet_name"]) in sheets and item["status"] not in {"fixed", "accepted"}
            )
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
                        confirm_locations=False,
                    )
                    next_seq += len(items)
                    new_pages.update(pages)
                    new_items.update(items)
                    log_msg(f"  [{sheet_name}] 新 PDF 检出 {len(items)} 处疑似压线")
                if _file_version(Path(path)) != version:
                    raise ValueError(f"复检期间文件又被修改，本轮不提交，请保存完成后再次刷新: {Path(path).name}")
                with STATE_LOCK:
                    previous = copy.deepcopy(STATE)
                    previous_items = [
                        item for item in STATE["items"].values()
                        if item["wb_path"] == path and item["sheet_name"] in entry["sheets"]
                    ]
                    merged = _merge_redetected(new_items, previous_items, mark_new=True, keep_unmatched=True)
                    if job["kind"] == "review_recheck":
                        for item in merged.values():
                            if item.get("later_review") and item["status"] not in RESOLVED_STATUSES:
                                item["review_reason"] = LATER_REVIEW_FAILED_REASON
                    if job["kind"] == "recheck":
                        token = uuid4().hex
                        manual_rows = {tuple(row) for row in job["options"].get("manual_rows", [])}
                        manual_ids = set(job["options"].get("manual_item_ids", []))
                        old_by_id = {item["id"]: item for item in previous_items}
                        for item in merged.values():
                            if item["status"] in RESOLVED_STATUSES:
                                continue
                            item["autofix_revision"] = token
                            if (item.get("is_new") or item["id"] in manual_ids
                                    or old_by_id.get(item["id"], {}).get("status") == "manual_review"
                                    or (path, item["sheet_name"], int(item["excel_row"])) in manual_rows):
                                item.update(
                                    status="manual_review", manual_waiting=False,
                                    manual_recheck_token=token,
                                    review_reason="本次复检仍有问题，请继续修改、跳过本行或先看下一行",
                                )
                    for collection in ("pages", "items"):
                        for key, value in list(STATE[collection].items()):
                            if value["wb_path"] == path and value["sheet_name"] in entry["sheets"]:
                                del STATE[collection][key]
                    STATE["pages"].update(new_pages)
                    STATE["items"].update(merged)
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


# ==========================================
# 稍后人工复检：先挂起，最后一次打开全部文件人工修改，再统一复检
# ==========================================
def queue_later_review(item_ids: list[str]) -> None:
    """调用方持有 STATE_LOCK：把这些问题挂到稍后人工复检，后台随后给单元格标黄、给文件名加 🟡。"""
    previous = {iid: copy.deepcopy(STATE["items"][iid]) for iid in item_ids}
    previous_queue = copy.deepcopy(STATE["review_queue"])
    for iid in item_ids:
        item = STATE["items"][iid]
        item.pop("manual_recheck_token", None)
        item.update(status="manual_review", manual_waiting=True, later_review=True, review_reason=LATER_REVIEW_REASON)
        entry = STATE["review_queue"].setdefault(item["wb_path"], {"marks": [], "error": ""})
        addr = str(item.get("cell_addr", ""))
        mark = next((m for m in entry["marks"] if (m["sheet"], m["addr"]) == (item["sheet_name"], addr)), None)
        if mark is None:
            mark = {
                "sheet": item["sheet_name"], "addr": addr, "row": int(item["excel_row"]),
                "item_ids": [], "state": "pending", "error": "",
            }
            entry["marks"].append(mark)
        if iid not in mark["item_ids"]:
            mark["item_ids"].append(iid)
    try:
        _commit_confirmations(previous, "已挂起，稍后统一人工复检")
    except Exception:
        STATE["review_queue"] = previous_queue
        raise


def review_entry_needs_mark(path: str, entry: dict[str, Any]) -> bool:
    """还有没标黄的单元格，或文件名还没加 🟡（后台标黄失败过的不再自动重试）"""
    if entry.get("mark_failed"):
        return False
    return any(m["state"] == "pending" for m in entry["marks"]) or (
        bool(entry["marks"]) and not Path(path).name.startswith(REVIEW_EMOJI)
        and not entry.get("rename_error")
    )


def _rebase_checked_file(path: str, before: dict[str, int]) -> None:
    """只是标黄 / 去黄色时，文件的检查基线跟着前进，免得下次扫描把它当成改过的文件"""
    if STATE["checked_files"].get(path) == before:
        STATE["checked_files"][path] = _file_version(Path(path))


def _replace_job_path(old: str, new: str) -> None:
    job = STATE["job"]
    if not job:
        return
    for key in ("workbooks", "completed"):
        if isinstance(job.get(key), list):
            job[key] = [new if value == old else value for value in job[key]]
    options = job["options"]
    if isinstance(options.get("recheck_plan"), dict) and old in options["recheck_plan"]:
        options["recheck_plan"][new] = options["recheck_plan"].pop(old)
    for row in options.get("manual_rows", []):
        if row[0] == old:
            row[0] = new
    if isinstance(options.get("review_files"), list):
        options["review_files"] = [new if value == old else value for value in options["review_files"]]


def _rename_review_file(path: str, new_name: str) -> str:
    """改文件名（加 / 去 🟡）并同步工作台记录；改不了时记下原因并返回原路径。"""
    old = Path(path)
    new = old.with_name(new_name)
    if new == old:
        return path
    error = ""
    if new.exists():
        error = f"已有同名文件「{new.name}」，没有改文件名"
    else:
        try:
            old.rename(new)
        except OSError as err:
            error = f"改文件名失败: {err}"
    if error:
        with STATE_LOCK:
            if path in STATE["review_queue"]:
                STATE["review_queue"][path]["rename_error"] = error
                _persist_state()
        log_msg(f"⚠️ {old.name}: {error}")
        return path
    from core import autofix
    with STATE_LOCK:
        previous = copy.deepcopy(STATE)
        _sync_renamed_workbook_in_state(old, new, {})
        if path in STATE["review_queue"]:
            STATE["review_queue"][str(new)] = STATE["review_queue"].pop(path)
            STATE["review_queue"][str(new)].pop("rename_error", None)
        _replace_job_path(path, str(new))
        try:
            _persist_state()
        except Exception:
            STATE.clear()
            STATE.update(previous)
            new.rename(old)
            raise
        autofix.rename_groups(path, str(new))
    log_msg(f"文件名 {old.name} → {new.name}")
    return str(new)


def mark_review_file(excel: Any, path: str) -> bool:
    """给待复检单元格标黄并保存，再给文件名加 🟡；文件正被 Excel 打开时返回 False，稍后再试。

    调用方保证此时没有别的任务在用这个文件。
    """
    wb_path = Path(path)
    if wb_path.with_name("~$" + wb_path.name).exists():
        with STATE_LOCK:
            if path in STATE["review_queue"]:
                STATE["review_queue"][path]["error"] = "文件正在 Excel 中打开，关闭后自动标黄"
        return False
    with STATE_LOCK:
        entry = STATE["review_queue"].get(path)
        if entry is None:
            return True
        pending = [copy.deepcopy(m) for m in entry["marks"] if m["state"] == "pending"]
    if pending:
        before = _file_version(wb_path)
        workbook = excel.Workbooks.Open(str(wb_path.resolve()), UpdateLinks=0)
        try:
            mark_review_cells(workbook, pending)
            normalize_workbook_view(excel, workbook)
            workbook.Save()
        finally:
            workbook.Close(SaveChanges=False)
        done = {(m["sheet"], m["addr"]): m for m in pending}
        with STATE_LOCK:
            entry = STATE["review_queue"].get(path, {"marks": []})
            for mark in entry["marks"]:
                result = done.get((mark["sheet"], mark["addr"]))
                if result is not None and mark["state"] == "pending":
                    mark.update({key: result[key] for key in ("state", "error", "range", "fill") if key in result})
            entry["error"] = ""
            _rebase_checked_file(path, before)
            _persist_state()
        marked = sum(m["state"] == "marked" for m in pending)
        skipped = [f"{m['sheet']}!{m['addr']} {m['error']}" for m in pending if m["state"] != "marked"]
        log_msg(
            f"🟡 稍后人工复检：{wb_path.name} 已把 {marked} 处单元格标成黄色"
            + (f"；没标黄 {len(skipped)} 处（{'；'.join(skipped[:3])}）" if skipped else "")
        )
    with STATE_LOCK:
        entry = STATE["review_queue"].get(path)
        rename = entry is not None and bool(entry["marks"])
    if rename:
        _rename_review_file(path, review_marked_name(wb_path.name))
    return True


def _mark_resolved(mark: dict[str, Any]) -> bool:
    # 问题项不在了（例如完整重新扫描后对应不上）也算这处已经没问题
    return all(
        STATE["items"].get(iid, {"status": "fixed"})["status"] in RESOLVED_STATUSES for iid in mark["item_ids"]
    )


def restore_review_file(excel: Any, path: str) -> str:
    """统一复检后：没问题的单元格去掉黄色、还原原来的填充；整个文件都没问题了就去掉 🟡、移出队列。

    返回 clean（已移出队列）/ partial（还有问题）/ locked（文件被打开，没能去黄色）。
    """
    wb_path = Path(path)
    with STATE_LOCK:
        entry = STATE["review_queue"].get(path)
        if entry is None:
            return "clean"
        done = [copy.deepcopy(m) for m in entry["marks"] if _mark_resolved(m)]
    to_restore = [m for m in done if m["state"] == "marked"]
    before = None
    if to_restore:
        if wb_path.with_name("~$" + wb_path.name).exists():
            return "locked"
        before = _file_version(wb_path)
        workbook = excel.Workbooks.Open(str(wb_path.resolve()), UpdateLinks=0)
        try:
            restore_review_cells(workbook, to_restore)
            normalize_workbook_view(excel, workbook)
            workbook.Save()
        finally:
            workbook.Close(SaveChanges=False)
    with STATE_LOCK:
        if before is not None:
            _rebase_checked_file(path, before)
        entry = STATE["review_queue"][path]
        keys = {(m["sheet"], m["addr"]) for m in done}
        entry["marks"] = [m for m in entry["marks"] if (m["sheet"], m["addr"]) not in keys]
        for mark in done:
            for iid in mark["item_ids"]:
                if iid in STATE["items"]:
                    STATE["items"][iid].pop("later_review", None)
                    STATE["items"][iid].pop("manual_waiting", None)
        remaining = len(entry["marks"])
        _persist_state()
    if to_restore:
        log_msg(f"统一复检：{wb_path.name} 已去掉 {len(to_restore)} 处黄色")
    if remaining:
        return "partial"
    new_path = _rename_review_file(path, strip_review_prefix(wb_path.name))
    with STATE_LOCK:
        STATE["review_queue"].pop(new_path, None)
        _persist_state()
    return "clean"


def review_open_workflow() -> None:
    """一次打开全部待复检文件：还没标黄的先标完，再在同一个可见 Excel 里逐个打开并定位。"""
    with STATE_LOCK:
        for entry in STATE["review_queue"].values():
            entry.pop("rename_error", None)  # 打开前再试一次标黄和改名
            entry.pop("mark_failed", None)
        paths = [path for path, entry in STATE["review_queue"].items() if review_entry_needs_mark(path, entry)]
    if paths:
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            for path in paths:
                try:
                    if not mark_review_file(excel, path):
                        log_msg(f"⚠️ {Path(path).name} 正在 Excel 中打开，没能标黄，直接打开")
                except Exception as err:
                    log_msg(f"⚠️ {Path(path).name} 标黄失败: {err}")
        finally:
            try:
                if excel is not None:
                    excel.Quit()
            finally:
                pythoncom.CoUninitialize()
    with STATE_LOCK:
        targets = [
            (Path(path), entry["marks"][0]["sheet"], entry["marks"][0]["addr"], int(entry["marks"][0]["row"]))
            for path, entry in STATE["review_queue"].items() if entry["marks"]
        ]
    errors = open_workbooks_for_review(targets)
    for path, reason in errors.items():
        log_msg(f"⚠️ 没能打开 {path.name}: {reason}")
    message = (
        f"已在 Excel 打开 {len(targets) - len(errors)} 个待复检文件，黄色的单元格就是要改的地方；"
        "改完保存后回到网页点「统一复检」"
    )
    log_msg(message)


def review_recheck_workflow() -> None:
    """统一复检：只重新检查挂起的文件，没问题的去掉黄色和 🟡，还有问题的继续留在队列里。"""
    with STATE_LOCK:
        job = STATE["job"]
        options = job["options"]
        first_run = "review_files" not in options
        if first_run:
            options["review_files"] = list(STATE["review_queue"])
            _persist_state()
        files = list(options["review_files"])
    if first_run:
        unsaved = close_saved_workbooks([Path(path) for path in files])
        skipped: dict[str, str] = {}
        for path in files:
            wb_path = Path(path)
            if wb_path in unsaved:
                skipped[path] = unsaved[wb_path]
            elif not wb_path.exists():
                skipped[path] = "找不到文件，可能被移动或改名"
            elif wb_path.with_name("~$" + wb_path.name).exists():
                skipped[path] = "文件仍在 Excel 中打开，请关闭后再复检"
        with STATE_LOCK:
            ids = [
                iid
                for path, entry in STATE["review_queue"].items() if path in files and path not in skipped
                for mark in entry["marks"] for iid in mark["item_ids"]
                if iid in STATE["items"] and STATE["items"][iid]["status"] not in RESOLVED_STATUSES
            ]
            options["review_skipped"] = skipped
            options["selected_ids"] = list(dict.fromkeys(ids))
            _persist_state()
        for path, reason in skipped.items():
            log_msg(f"⚠️ 统一复检跳过 {Path(path).name}: {reason}")
    with STATE_LOCK:
        ids = list(options["selected_ids"])
        skipped = dict(options["review_skipped"])
    if ids:
        recheck_items_workflow(ids)
    results: dict[str, str] = {}
    pythoncom.CoInitialize()
    excel = None
    try:
        for path in files:
            if path in skipped:
                continue
            if excel is None:
                excel = win32.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
            results[path] = restore_review_file(excel, path)
    finally:
        try:
            if excel is not None:
                excel.Quit()
        finally:
            pythoncom.CoUninitialize()
    with STATE_LOCK:
        left = sum(len(m["item_ids"]) for entry in STATE["review_queue"].values() for m in entry["marks"])
    clean = sum(result == "clean" for result in results.values())
    partial = len(results) - clean
    message = f"统一复检完成：{clean} 个文件已没问题（去掉了黄色和 🟡）"
    if partial:
        message += f"，{partial} 个文件还有问题，继续留在复检队列里（共 {left} 处）"
    if skipped:
        message += f"；{len(skipped)} 个文件没复检（未保存或仍打开）"
    log_msg(message)
    finish_execution_batch(message)


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
            # 隐藏截断项靠加高行不一定能解决（也可能是列宽），只报告，由人工在 Excel 中修改
            if item and item.get("kind") != "hidden" and (job is None or iid not in job["completed"]):
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

        with STATE_LOCK:
            folder_path = (
                Path(STATE["folder"])
                if STATE["folder"]
                else Path(next(iter(by_wb.keys()))).parent
            )

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
    """只读递归列出每个工作簿的全部工作表、默认勾选删除的备份表及已带 ❓ 的待定表；不修改文件，也不占用后台任务。"""
    folder, workbooks = list_workbooks(folder_str)
    files: list[dict[str, Any]] = []
    excel = None

    def get_excel() -> Any:
        # .xls 只能用 Excel 读取：整次预览共用一个 Excel 进程，没有 .xls 时不启动
        nonlocal excel
        if excel is None:
            pythoncom.CoInitialize()
            try:
                excel = win32.DispatchEx("Excel.Application")
            except Exception:
                pythoncom.CoUninitialize()
                raise
            excel.Visible = False
            excel.DisplayAlerts = False
        return excel

    try:
        for wb_path in workbooks:
            files.append(_preview_workbook(wb_path, folder, get_excel))
    finally:
        if excel is not None:
            try:
                excel.Quit()
            finally:
                pythoncom.CoUninitialize()
    return {"folder": str(folder), "files": files}


def _preview_workbook(wb_path: Path, folder: Path, get_excel: Callable[[], Any]) -> dict[str, Any]:
    """读取单个工作簿的工作表，并给出默认标为删除的备份表与已带 ❓ 的待定表"""
    rel_name = workbook_rel_name(wb_path, folder)
    entry: dict[str, Any] = {
        "name": rel_name,
        "sheet_count": 0,
        "sheets": [],
        "backups": [],
        "uncertain": [],
        "new_old_pairs": {},
        "note": "",
    }
    try:
        if wb_path.suffix.lower() == ".xls":
            all_names, _ = scan_workbook_sheets(wb_path, get_excel)
        else:
            all_names, _ = scan_workbook_sheets(wb_path)
    except Exception as err:
        entry["note"] = f"无法读取: {err}"
    else:
        entry["sheet_count"] = len(all_names)
        entry["sheets"] = all_names  # 保持工作簿中的原有顺序，便于对照保留、删除与待定
        uncertain = [s for s in all_names if is_uncertain_name(s)]
        default_delete = [s for s in find_clean_targets(all_names) if s not in set(uncertain)]
        contained = set(find_contained_sheets(all_names))
        entry["uncertain"] = [s for s in all_names if s in contained or s in set(uncertain)]
        entry["new_old_pairs"] = find_new_old_pairs(all_names)  # 旧表全部删除时，新表去掉「新」字
        if default_delete and len(default_delete) >= len(all_names):
            entry["note"] = "全部工作表均被识别为备份表，将跳过删除"
        else:
            entry["backups"] = default_delete
    return entry


def validate_clean_plan(plan: Any) -> dict[str, dict[str, list[str]]]:
    """校验前端提交的精简 Sheet 选择方案：{文件名: {"delete": [...], "uncertain": [...]}}"""
    if not isinstance(plan, dict):
        raise ValueError("精简方案格式无效")
    normalized: dict[str, dict[str, list[str]]] = {}
    for wb_name, item in plan.items():
        if not isinstance(wb_name, str) or not wb_name or not isinstance(item, dict):
            raise ValueError("精简方案格式无效")
        delete_list = item.get("delete", [])
        uncertain_list = item.get("uncertain", [])
        if (
            not isinstance(delete_list, list)
            or not isinstance(uncertain_list, list)
            or any(not isinstance(s, str) or not s for s in delete_list)
            or any(not isinstance(s, str) or not s for s in uncertain_list)
        ):
            raise ValueError("工作表选择列表格式无效")
        del_dedup = list(dict.fromkeys(delete_list))
        unc_dedup = list(dict.fromkeys(uncertain_list))
        if set(del_dedup) & set(unc_dedup):
            raise ValueError(f"{wb_name} 中同一工作表不能同时勾选删除和标记待定")
        normalized[wb_name] = {"delete": del_dedup, "uncertain": unc_dedup}
    return normalized


def validate_clean_files(files: Any) -> list[str]:
    """校验「只处理这些文件」的列表（预览窗口只处理当前文件；主页面跳过已在预览窗口处理过的文件）"""
    if not isinstance(files, list) or not files or any(not isinstance(f, str) or not f for f in files):
        raise ValueError("要处理的文件列表格式无效")
    return list(dict.fromkeys(files))


def _resolve_clean_targets(
    all_names: list[str], file_plan: dict[str, list[str]] | None
) -> tuple[list[str], list[str], list[tuple[str, str]], bool]:
    """根据工作表列表与用户选择，返回 (待删除表, 待定表, [(原表名, 新表名)], 是否试图删空整本工作簿)

    待定表只是本次不删，不再给表名或文件名加 ❓；改名只有删除旧版后去掉「新」字这一种。
    """
    if file_plan is None:
        candidates = find_clean_targets(all_names)
        delete_all = bool(all_names and len(candidates) >= len(all_names))
        to_delete = [] if delete_all else candidates
        uncertain_sel = [s for s in all_names if s not in set(to_delete) and is_uncertain_name(s)]
    else:
        del_set = set(file_plan.get("delete", []))
        unc_set = set(file_plan.get("uncertain", []))
        candidates = [s for s in all_names if s in del_set]
        delete_all = bool(all_names and len(candidates) >= len(all_names))
        to_delete = [] if delete_all else candidates
        uncertain_sel = [s for s in all_names if s not in set(to_delete) and s in unc_set]

    if delete_all:
        return [], [], [], True

    occupied = {s for s in all_names if s not in set(to_delete)}
    renames = _new_prefix_renames(all_names, set(to_delete), set(uncertain_sel), occupied)
    return to_delete, uncertain_sel, renames, False


def _new_prefix_renames(
    all_names: list[str], deleted: set[str], uncertain: set[str], occupied: set[str],
) -> list[tuple[str, str]]:
    """「旧」表全部删除后，对应的「新」表去掉开头的「新」字（Excel 改表名时会自动更新公式与打印区域等引用）。

    「新」表本身被删除或标为待定、还有「旧」表保留着、去掉后与现有表重名（不区分大小写）时都不改名。
    occupied 为执行后剩下的表名，改名后同步更新。
    """
    renames: list[tuple[str, str]] = []
    for new_name, olds in find_new_old_pairs(all_names).items():
        if new_name in deleted or new_name in uncertain or not all(o in deleted for o in olds):
            continue
        target = strip_new_prefix(new_name)
        taken = {n.casefold() for n in occupied if n != new_name}
        if target == new_name or target.casefold() in taken:
            if target != new_name:
                log_msg(f"  ⚠️ 「{new_name}」去掉「新」字后与现有工作表「{target}」重名，保持原名")
            continue
        occupied.discard(new_name)
        occupied.add(target)
        renames.append((new_name, target))
    return renames


def _sync_renamed_workbook_in_state(
    old_path: Path, new_path: Path, sheet_rename_map: dict[str, str], root_folder: Path | None = None
) -> None:
    """工作表改名（去掉「新」字）后，同步更新工作台已有的扫描记录"""
    old_str = str(old_path)
    new_str = str(new_path)
    if old_str == new_str and not sheet_rename_map:
        return
    with STATE_LOCK:
        folder_ref = root_folder or STATE.get("folder")
        new_label = workbook_rel_name(new_path, folder_ref)
        for collection in ("pages", "items"):
            for entry in STATE[collection].values():
                if entry.get("wb_path") == old_str:
                    entry["wb_path"] = new_str
                    entry["wb_name"] = new_label
                    if entry.get("sheet_name") in sheet_rename_map:
                        entry["sheet_name"] = sheet_rename_map[entry["sheet_name"]]
        if old_str != new_str and old_str in STATE["checked_files"]:
            STATE["checked_files"][new_str] = STATE["checked_files"].pop(old_str)


def _format_clean_action(
    deleted_list: list[str], uncertain_list: list[str], new_prefix_removed: int = 0,
) -> str:
    parts: list[str] = []
    if deleted_list:
        parts.append(f"已删除 {len(deleted_list)} 个备份 Sheet")
    if new_prefix_removed:
        parts.append(f"{new_prefix_removed} 个新版 Sheet 已去掉「新」字")
    if uncertain_list:
        parts.append(f"{len(uncertain_list)} 个待定 Sheet 未删除")
    if not parts:
        return "没有备份表，无需精简"
    return "，".join(parts)


def _clean_backup_sheets(
    get_excel: Callable[[], Any],
    wb_path: Path,
    root_folder: Path,
    file_plan: dict[str, list[str]] | None = None,
) -> None:
    wb_label = workbook_rel_name(wb_path, root_folder)
    # 续跑时沿用本批次已记录的删除明细，避免保存后中断导致记录丢失
    previous = get_execution_report_snapshot()["files"].get(wb_label, {})
    deleted = list(previous.get("deleted_backups", []))
    uncertain_recorded = list(previous.get("uncertain_sheets", []))

    def record(all_names: list[str], action: str, *, uncertain: list[str] | None = None) -> None:
        update_file_progress(
            wb_label,
            status="done",
            current_action=action,
            deleted_backups=deleted,
            uncertain_sheets=uncertain if uncertain is not None else uncertain_recorded,
            all_sheets=list(dict.fromkeys(previous.get("all_sheets", []) + all_names)),
        )

    if wb_path.suffix.lower() != ".xls":
        # 先用 openpyxl 只读预检，无需修改工作表的文件不必用 Excel 打开
        try:
            pre_names, _ = scan_workbook_sheets(wb_path)
        except Exception:
            pre_names = None  # 预检失败时交给 Excel 判断
        if pre_names is not None:
            to_delete, uncertain_sel, renames, del_all = _resolve_clean_targets(pre_names, file_plan)
            if not del_all and not to_delete and not renames:
                action = _format_clean_action(deleted, uncertain_sel)
                if uncertain_sel:
                    log_msg(f"  {action}")
                record(pre_names, action, uncertain=uncertain_sel)
                return

    wb = get_excel().Workbooks.Open(str(wb_path.resolve()))
    rename_map: dict[str, str] = {}
    final_uncertain: list[str] = []
    try:
        all_names = [sheet.Name for sheet in wb.Sheets]
        final_names = all_names
        to_delete, uncertain_sel, renames, del_all = _resolve_clean_targets(all_names, file_plan)
        if del_all:
            action = (
                "全部工作表均被识别为备份表，已跳过删除"
                if file_plan is None
                else "全部工作表均被勾选删除，已跳过删除"
            )
            log_msg(f"  ⚠️ {action}: {wb_label}")
        elif not to_delete and not renames:
            final_uncertain = uncertain_sel
            action = _format_clean_action(deleted, final_uncertain)
        else:
            for name in to_delete:
                wb.Sheets(name).Delete()
            deleted = list(dict.fromkeys(deleted + to_delete))
            rename_map = dict(renames)
            for old_name, new_name in renames:
                wb.Sheets(old_name).Name = new_name
                log_msg(f"  已改名: {old_name} → {new_name}")
            final_names = [rename_map.get(s, s) for s in all_names]
            final_uncertain = [rename_map.get(s, s) for s in uncertain_sel]
            # 先记录删除与待定明细再保存，保存后中断也能在续跑时保留记录
            update_file_progress(
                wb_label, deleted_backups=deleted, uncertain_sheets=final_uncertain,
            )
            wb.Save()
            action = _format_clean_action(deleted, final_uncertain, len(renames))
            log_msg(f"  {action}")
    finally:
        wb.Close(SaveChanges=False)

    if rename_map:
        _sync_renamed_workbook_in_state(wb_path, wb_path, rename_map, root_folder)
    record(final_names, action, uncertain=final_uncertain)


def _export_workbook_pdfs(
    get_excel: Callable[[], Any], wb_path: Path, root_folder: Path, used_pdf_paths: set[Path],
) -> list[Path]:
    wb_label = workbook_rel_name(wb_path, root_folder)
    wb = get_excel().Workbooks.Open(str(wb_path.resolve()), ReadOnly=True)
    try:
        file_stem = sanitize_filename(strip_review_prefix(wb_path.stem))
        exported: list[Path] = []
        for sheet in _exportable_sheets(wb):
            pdf_path = unique_pdf_path(
                wb_path.parent, file_stem, sanitize_filename(sheet.Name), used_pdf_paths
            )
            _export_sheet_pdf(sheet, pdf_path)
            exported.append(pdf_path)
            update_file_progress(
                wb_label, exported_pdfs=len(exported),
                current_action=f"正在导出 PDF（已导出 {len(exported)} 份）...",
            )
        action = f"已导出 {len(exported)} 份 PDF"
        log_msg(f"  {action}")
        update_file_progress(
            wb_label, status="done", exported_pdfs=len(exported), current_action=action,
            all_sheets=[sheet.Name for sheet in wb.Sheets],
        )
        return exported
    finally:
        wb.Close(SaveChanges=False)


def _normalize_workbook_file(get_excel: Callable[[], Any], wb_path: Path, root_folder: Path) -> None:
    wb_label = workbook_rel_name(wb_path, root_folder)
    excel = get_excel()
    wb = excel.Workbooks.Open(str(wb_path.resolve()))
    try:
        normalize_workbook_view(excel, wb, on_error=log_msg)
        wb.Save()
        update_file_progress(
            wb_label, status="done", normalized_a1=True,
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
        only_files = options.get("files")
        if kind == "clean_sheets" and isinstance(only_files, list):
            wanted = set(only_files)
            workbooks = [wb for wb in workbooks if workbook_rel_name(wb, folder) in wanted]
        with STATE_LOCK:
            STATE["logs"].clear()
            waiting_review = len(STATE["review_queue"])
            options["folder"] = str(folder)
            if kind == "export_pdf":
                options["pdf_folder"] = str(folder)
            job["workbooks"] = [str(path) for path in workbooks]
            job["exported"] = []
            _persist_state()
        init_execution_batch(
            task_name=task_name,
            folder=folder,
            workbooks=workbooks,
            pdf_folder=folder if kind == "export_pdf" else None,
            backup_folder=None,
        )
        if kind == "export_pdf" and waiting_review:
            log_msg(f"⚠️ 还有 {waiting_review} 个文件在等稍后人工复检（单元格标着黄色），导出的 PDF 里会带黄色底色")
    log_msg(f"{'继续' if initialized else '开始'}{task_name}：{folder}（共 {len(workbooks)} 个文件）")

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
            wb_label = workbook_rel_name(wb_path, folder)
            log_msg(f"正在处理: {wb_label} ...")
            update_file_progress(
                wb_label, status="processing", current_action=f"正在执行{task_name}...", error="",
            )
            exported: list[Path] = []
            try:
                if kind == "clean_sheets":
                    clean_plan = options.get("plan")
                    file_plan = None
                    if isinstance(clean_plan, dict):
                        file_plan = (
                            clean_plan.get(wb_label)
                            if wb_label in clean_plan
                            else clean_plan.get(wb_path.name, {"delete": [], "uncertain": []})
                        )
                    _clean_backup_sheets(
                        get_excel, wb_path, folder, file_plan=file_plan,
                    )
                elif kind == "export_pdf":
                    exported = _export_workbook_pdfs(
                        get_excel, wb_path, folder, used_pdf_paths
                    )
                else:
                    _normalize_workbook_file(get_excel, wb_path, folder)
            except Exception as err:
                log_msg(f"❌ 处理文件出错 {wb_label}: {err}")
                update_file_progress(
                    wb_label, status="error", error=str(err), current_action=f"处理出错: {err}",
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
        deleted_total = sum(len(f["deleted_backups"]) for f in files)
        uncertain_total = sum(len(f.get("uncertain_sheets", [])) for f in files)
        uncertain_files = sum(bool(f.get("uncertain_sheets")) for f in files)
        detail = f"共删除 {deleted_total} 个备份 Sheet"
        if uncertain_total:
            detail += f"，{uncertain_total} 个待定 Sheet 未删除（{uncertain_files} 个文件）"
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
    重放到新版本文件上；扫描结果与人工确认保留。
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
    imported.setdefault("scan_progress", {})
    imported.setdefault("review_queue", {})
    if not isinstance(imported["scan_progress"], dict):
        raise ValueError("进度文件损坏：扫描进度格式无效")
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
        _persist_state()
        return {
            "exported_at": data.get("exported_at", ""),
            "items": len(STATE["items"]),
            "files": len(get_execution_report_snapshot()["files"]),
            "can_resume": bool(job and job["status"] == "interrupted"),
            "warnings": warnings,
        }
