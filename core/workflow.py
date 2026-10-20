import copy
import re
import shutil
import threading
from datetime import datetime
from pathlib import Path
from typing import Any

from core.common import (
    EXCEL_EXTENSIONS,
    EXCLUDE_PDF_SHEETS,
    XL_SHEET_VISIBLE,
    XL_TYPE_PDF,
    append_record,
    as_bbox,
    is_backup_sheet,
    sanitize_filename,
    unique_pdf_path,
)
from core.excel_ops import (
    extract_sheet_snapshot,
    match_excel_row,
    normalize_workbook_view,
    pythoncom,
    win32,
)
from core.pdf_detector import (
    find_clipped_spans_on_page,
    fitz,
    is_region_still_clipped,
    refresh_pdf_pages_overview,
    render_crop_base64,
    render_page_overview,
)
from core.reporter import (
    finish_execution_batch,
    get_execution_report_snapshot,
    init_execution_batch,
    update_file_progress,
)

STATE_LOCK = threading.RLock()

# 全局状态存储（供前端网页读取和操作）
STATE: dict[str, Any] = {
    "folder": "",
    "pdf_folder": "",
    "backup_folder": "",
    "pages": {},  # page_id -> 整页全景图与页面尺寸信息
    "items": {},  # item_id -> 检测出的切字项详情
    "logs": [],
    "busy": False,
}


def log_msg(msg: str) -> None:
    timestamp = datetime.now().strftime("%H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    with STATE_LOCK:
        STATE["logs"].append(line)


def try_acquire_busy() -> bool:
    with STATE_LOCK:
        if STATE["busy"]:
            return False
        STATE["busy"] = True
        return True


def get_state_snapshot() -> dict[str, Any]:
    report_snap = get_execution_report_snapshot()
    with STATE_LOCK:
        return {
            "busy": STATE["busy"],
            "folder": STATE["folder"],
            "pdf_folder": STATE["pdf_folder"],
            "logs": list(STATE["logs"][-40:]),
            "pages": copy.deepcopy(STATE["pages"]),
            "items": copy.deepcopy(STATE["items"]),
            "progress": {
                "task_name": report_snap["task_name"],
                "status": report_snap["status"],
                "total_files": report_snap["total_files"],
                "completed_files": report_snap["completed_files"],
                "processing_files": report_snap["processing_files"],
                "pending_files": report_snap["pending_files"],
                "error_files": report_snap["error_files"],
                "progress_pct": report_snap["progress_pct"],
                "current_file": report_snap["current_file"],
                "current_step": report_snap["current_step"],
            },
        }


def update_item_row(item_id: str, new_row: int) -> bool:
    with STATE_LOCK:
        if item_id not in STATE["items"]:
            return False
        item = STATE["items"][item_id]
        item["excel_row"] = new_row
        old_addr = str(item.get("cell_addr", ""))
        item["cell_addr"] = (
            re.sub(r"\d+$", str(new_row), old_addr)
            if re.search(r"\d+$", old_addr)
            else f"Row {new_row}"
        )
        return True


def _make_page_id(wb_path_str: str, sheet_name: str, page_index: int) -> str:
    safe_key = re.sub(r"[^a-zA-Z0-9_\u4e00-\u9fff\u3040-\u30ff]", "_", f"{Path(wb_path_str).stem}_{sheet_name}")
    return f"page_{safe_key}_{page_index + 1}"


def scan_folder_workflow(folder_str: str, delete_backups: bool) -> None:
    with STATE_LOCK:
        STATE["busy"] = True
        STATE["pages"].clear()
        STATE["items"].clear()
        STATE["logs"].clear()

    try:
        folder = Path(folder_str.strip().strip('"').strip("'")).expanduser().resolve()
        if not folder.is_dir():
            log_msg(f"❌ 文件夹不存在: {folder}")
            return

        workbooks = sorted(
            p
            for p in folder.iterdir()
            if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS and not p.name.startswith("~$")
        )
        if not workbooks:
            log_msg("❌ 未找到 Excel 文件。")
            return

        timestamp = f"{datetime.now():%Y%m%d_%H%M%S}"
        pdf_folder = folder / f"_exported_pdfs_{timestamp}"
        backup_folder = folder / f"_backup_{timestamp}"
        pdf_folder.mkdir(parents=True, exist_ok=True)
        used_pdf_paths: set[Path] = set()

        with STATE_LOCK:
            STATE["folder"] = str(folder)
            STATE["pdf_folder"] = str(pdf_folder)
            STATE["backup_folder"] = str(backup_folder)

        init_execution_batch(
            task_name="Web 导出 PDF 与切字扫描",
            folder=folder,
            workbooks=workbooks,
            pdf_folder=pdf_folder,
            backup_folder=backup_folder,
        )

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        item_seq = 1
        page_seq = 1
        try:
            for wb_path in workbooks:
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
                    bkp_names = [n for n in all_names if is_backup_sheet(n)]

                    if delete_backups and bkp_names:
                        if len(bkp_names) >= len(all_names):
                            log_msg(f"  ⚠️ 跳过删除备份表（全部工作表均为日期名）: {wb_path.name}")
                        else:
                            backup_folder.mkdir(parents=True, exist_ok=True)
                            shutil.copy2(wb_path, backup_folder / wb_path.name)
                            for bn in bkp_names:
                                wb.Sheets(bn).Delete()
                            wb.Save()
                            deleted_bkp_list = list(bkp_names)
                            try:
                                append_record(wb_path.name, all_names, bkp_names)
                            except OSError as err:
                                log_msg(f"  ⚠️ 清理日志写入失败 {wb_path.name}: {err}")
                            log_msg(f"  已清理 {len(bkp_names)} 个日期备份 Sheet")

                    update_file_progress(
                        wb_path.name,
                        all_sheets=all_names,
                        deleted_backups=deleted_bkp_list,
                        current_action="正在导出 PDF 并检测切字...",
                    )

                    file_stem = sanitize_filename(wb_path.stem)
                    for sheet in wb.Sheets:
                        s_name = sheet.Name
                        if s_name in EXCLUDE_PDF_SHEETS or is_backup_sheet(s_name):
                            continue
                        if sheet.Visible != XL_SHEET_VISIBLE:
                            continue

                        pdf_path = unique_pdf_path(
                            pdf_folder, file_stem, sanitize_filename(s_name), used_pdf_paths
                        )
                        try:
                            sheet.ExportAsFixedFormat(
                                Type=XL_TYPE_PDF,
                                Filename=str(pdf_path.resolve()),
                                Quality=0,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False,
                            )
                            wb_exported_pdfs += 1
                        except Exception as e:
                            log_msg(f"  导出 PDF 失败 [{s_name}]: {e}")
                            continue

                        sheet_snapshot = extract_sheet_snapshot(sheet)

                        doc = fitz.open(pdf_path)
                        try:
                            total_pages = len(doc)
                            for p_idx in range(total_pages):
                                page = doc[p_idx]
                                clipped = find_clipped_spans_on_page(page)
                                if not clipped:
                                    continue

                                wb_detected_clips += len(clipped)
                                page_id = f"page_{page_seq}_{_make_page_id(str(wb_path), s_name, p_idx)}"
                                page_seq += 1
                                page_img_b64, page_w, page_h = render_page_overview(page)

                                page_entry = {
                                    "id": page_id,
                                    "wb_path": str(wb_path),
                                    "wb_name": wb_path.name,
                                    "sheet_name": s_name,
                                    "pdf_path": str(pdf_path),
                                    "page_index": p_idx,
                                    "total_pages": total_pages,
                                    "page_width": page_w,
                                    "page_height": page_h,
                                    "image": page_img_b64,
                                }
                                with STATE_LOCK:
                                    STATE["pages"][page_id] = page_entry

                                for clip_idx, clip_info in enumerate(clipped, start=1):
                                    _, by0, _, _ = clip_info["bbox"]
                                    rel_y = by0 / max(1.0, page.rect.height)
                                    excel_r, cell_addr, row_h = match_excel_row(
                                        sheet_snapshot, clip_info["text"], p_idx, total_pages, rel_y
                                    )
                                    img_b64 = render_crop_base64(page, clip_info["bbox"], highlight=True)
                                    item_id = f"item_{item_seq}"
                                    item_seq += 1
                                    new_item = {
                                        "id": item_id,
                                        "seq": item_seq - 1,
                                        "page_id": page_id,
                                        "index_on_page": clip_idx,
                                        "wb_path": str(wb_path),
                                        "wb_name": wb_path.name,
                                        "sheet_name": s_name,
                                        "pdf_path": str(pdf_path),
                                        "page_index": p_idx,
                                        "text": clip_info["text"],
                                        "bbox": clip_info["bbox"],
                                        "excel_row": excel_r,
                                        "cell_addr": cell_addr,
                                        "row_height": row_h,
                                        "image": img_b64,
                                        "status": "pending",  # pending / fixed / max_reached
                                        "iterations": 0,
                                    }
                                    with STATE_LOCK:
                                        STATE["items"][item_id] = new_item
                        finally:
                            doc.close()

                        update_file_progress(
                            wb_path.name,
                            exported_pdfs=wb_exported_pdfs,
                            clip_detected=wb_detected_clips,
                            clip_unfixed=wb_detected_clips,
                        )

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
            STATE["busy"] = False


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


def fix_selected_workflow(
    selected_ids: list[str],
    step_pt: float = 3.0,
    max_iter: int = 6,
) -> None:
    """
    针对用户勾选的热区/卡片：
    每次将对应 Excel 行高 +step_pt -> 重新导出该 Sheet 的 PDF -> 检测该处是否还切字
    一旦完整显示则立即停止，并同步刷新整页画布与局部 2.5x 放大镜截图！
    """
    with STATE_LOCK:
        STATE["busy"] = True
        by_wb: dict[str, list[dict[str, Any]]] = {}
        for iid in selected_ids:
            item = STATE["items"].get(iid)
            if item:
                by_wb.setdefault(item["wb_path"], []).append(item)
        backup_folder_str = STATE["backup_folder"]

    try:
        if not by_wb:
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

        try:
            for wb_path_str, items in by_wb.items():
                wb_path = Path(wb_path_str)
                bkp_target = backup_folder / wb_path.name
                if not bkp_target.exists():
                    shutil.copy2(wb_path, bkp_target)

                log_msg(f"🔧 开始微调工作簿: {wb_path.name} (共 {len(items)} 处)...")
                update_file_progress(
                    wb_path.name,
                    status="processing",
                    current_action=f"正在微调 {len(items)} 处选中切字项...",
                )
                wb = None
                touched_pdfs: set[Path] = set()
                try:
                    wb = excel.Workbooks.Open(str(wb_path.resolve()))
                    adjusted_rows_in_batch: set[tuple[str, int]] = set()

                    for item in items:
                        sheet_name = item["sheet_name"]
                        sheet = wb.Sheets(sheet_name)
                        target_row = max(1, int(item["excel_row"]))
                        pdf_path = Path(item["pdf_path"])
                        row_key = (sheet_name, target_row)

                        # 先检查是否因为同行其他单元格的微调已经顺带修好了
                        if (item["iterations"] > 0 or row_key in adjusted_rows_in_batch) and pdf_path.exists():
                            still_clip, new_img, new_bbox, _ = is_region_still_clipped(
                                pdf_path, item["page_index"], item["text"], as_bbox(item["bbox"])
                            )
                            curr_h = round(float(sheet.Rows(target_row).RowHeight or 15.0), 1)
                            if not still_clip:
                                with STATE_LOCK:
                                    item["image"] = new_img
                                    item["bbox"] = new_bbox
                                    item["status"] = "fixed"
                                    item["row_height"] = curr_h
                                _sync_row_height_in_state(wb_path_str, sheet_name, target_row, curr_h)
                                log_msg(f"  ✅ [{sheet_name} 行{target_row}] 已随同行调整自动恢复完整显示！")
                                continue

                        for it in range(1, max_iter + 1):
                            curr_h = float(sheet.Rows(target_row).RowHeight or 15.0)
                            new_h = round(curr_h + step_pt, 1)
                            sheet.Rows(target_row).RowHeight = new_h
                            adjusted_rows_in_batch.add(row_key)

                            with STATE_LOCK:
                                item["row_height"] = new_h
                                item["iterations"] = int(item.get("iterations", 0)) + 1
                            _sync_row_height_in_state(wb_path_str, sheet_name, target_row, new_h)

                            sheet.ExportAsFixedFormat(
                                Type=XL_TYPE_PDF,
                                Filename=str(pdf_path.resolve()),
                                Quality=0,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False,
                            )
                            touched_pdfs.add(pdf_path)

                            still_clip, new_img, new_bbox, _ = is_region_still_clipped(
                                pdf_path, item["page_index"], item["text"], as_bbox(item["bbox"])
                            )
                            with STATE_LOCK:
                                item["image"] = new_img
                                item["bbox"] = new_bbox

                            if not still_clip:
                                with STATE_LOCK:
                                    item["status"] = "fixed"
                                log_msg(
                                    f"  ✅ [{sheet_name} 行{target_row}] 第 {it} 次微调后已完整显示！(高度 -> {new_h}pt)"
                                )
                                break
                            else:
                                log_msg(
                                    f"  ⏳ [{sheet_name} 行{target_row}] 第 {it} 次微调 ({new_h}pt) 仍略微压线，继续+..."
                                )
                        else:
                            with STATE_LOCK:
                                item["status"] = "max_reached"
                            log_msg(f"  ⚠️ [{sheet_name} 行{target_row}] 已达最大微调次数 ({max_iter}次)")

                    # 刷新该工作簿涉及到的所有整页 PDF 全景图
                    for touched_pdf in touched_pdfs:
                        _refresh_sheet_pages_in_state(touched_pdf)

                    # 完成 A1 定位与 100% 缩放保存
                    normalize_workbook_view(excel, wb)
                    wb.Save()
                    log_msg(f"💾 已保存并完成 A1/100% 标准化: {wb_path.name}")

                    with STATE_LOCK:
                        wb_all_items = [
                            it for it in STATE["items"].values() if it["wb_path"] == wb_path_str
                        ]
                        wb_total_clips = len(wb_all_items)
                        wb_fixed_clips = sum(1 for it in wb_all_items if it["status"] == "fixed")
                        wb_unfixed_clips = wb_total_clips - wb_fixed_clips

                    update_file_progress(
                        wb_path.name,
                        status="done" if wb_unfixed_clips == 0 else "needs_fix",
                        clip_detected=wb_total_clips,
                        clip_fixed=wb_fixed_clips,
                        clip_unfixed=wb_unfixed_clips,
                        normalized_a1=True,
                        current_action=f"已微调保存 (修复 {wb_fixed_clips}/{wb_total_clips} 处，A1/100% 已标准化)",
                    )
                except Exception as err:
                    log_msg(f"❌ 修复失败 {wb_path.name}: {err}")
                    update_file_progress(
                        wb_path.name,
                        status="error",
                        error=str(err),
                        current_action=f"修复失败: {err}",
                    )
                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        done_msg = "🎉 所有选中项微调完毕！画布、放大镜截图与执行记录报告已更新。"
        log_msg(done_msg)
        finish_execution_batch(done_msg)
    finally:
        with STATE_LOCK:
            STATE["busy"] = False
