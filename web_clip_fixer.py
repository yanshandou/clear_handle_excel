import base64
import copy
import json
import re
import shutil
import sys
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

try:
    import pymupdf as fitz  # PyMuPDF
except ImportError:
    try:
        import fitz
    except ImportError:
        print("请先安装 pymupdf: pip install pymupdf")
        sys.exit(1)

try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    print("请先安装 pywin32: pip install pywin32")
    sys.exit(1)

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
EXCLUDE_PDF_SHEETS = {"変更履歴"}
DATE_PATTERN = re.compile(r"(?<!\d)((?:19|20)\d{6})(?!\d)")
RECORD_PATH = Path(__file__).with_name("record.md")

XL_TYPE_PDF = 0
XL_SHEET_VISIBLE = -1
XL_WORKSHEET = -4167

STATE_LOCK = threading.RLock()

# 全局状态存储（供前端网页读取和操作）
STATE = {
    "folder": "",
    "pdf_folder": "",
    "backup_folder": "",
    "items": {},  # id -> 检测出的切字项详情
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


def append_record(
    workbook_name: str,
    original_sheet_names: list[str],
    removed_sheet_names: list[str],
) -> None:
    separator = ""
    if RECORD_PATH.exists() and RECORD_PATH.stat().st_size > 0:
        with RECORD_PATH.open("rb") as record:
            record.seek(-1, 2)
            separator = "\n" if record.read(1) == b"\n" else "\n\n"
    removed_sheets = set(removed_sheet_names)
    with RECORD_PATH.open("a", encoding="utf-8", newline="\n") as record:
        record.write(f"{separator}## {workbook_name}\n\n")
        for sheet_name in original_sheet_names:
            display_name = f"~~{sheet_name}~~" if sheet_name in removed_sheets else sheet_name
            record.write(f"- {display_name}\n")


def is_backup_sheet(sheet_name: str) -> bool:
    for match in DATE_PATTERN.finditer(sheet_name):
        try:
            datetime.strptime(match.group(1), "%Y%m%d")
            return True
        except ValueError:
            continue
    return False


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name).strip()


def extract_page_lines(page: fitz.Page):
    """提取 PDF 页面中的所有水平线与垂直线"""
    h_lines = []  # (y, x0, x1)
    v_lines = []  # (x, y0, y1)
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "l":
                p1, p2 = item[1], item[2]
                if abs(p1.y - p2.y) < 1.2 and abs(p1.x - p2.x) > 5.0:
                    h_lines.append((p1.y, min(p1.x, p2.x), max(p1.x, p2.x)))
                elif abs(p1.x - p2.x) < 1.2 and abs(p1.y - p2.y) > 5.0:
                    v_lines.append((p1.x, min(p1.y, p2.y), max(p1.y, p2.y)))
            elif item[0] == "re":
                rect: fitz.Rect = item[1]
                if rect.height < 2.5 and rect.width > 5.0:
                    h_lines.append(((rect.y0 + rect.y1) / 2, rect.x0, rect.x1))
                elif rect.width < 2.5 and rect.height > 5.0:
                    v_lines.append(((rect.x0 + rect.x1) / 2, rect.y0, rect.y1))
                elif drawing.get("color") is not None:
                    h_lines.append((rect.y0, rect.x0, rect.x1))
                    h_lines.append((rect.y1, rect.x0, rect.x1))
                    v_lines.append((rect.x0, rect.y0, rect.y1))
                    v_lines.append((rect.x1, rect.y0, rect.y1))
    return h_lines, v_lines


def find_clipped_spans_on_page(page: fitz.Page):
    """
    在单页 PDF 上寻找被边框横线切穿/压线的文字块。
    返回列表: [{"text": str, "bbox": (x0,y0,x1,y1), "line_y": float}]
    """
    h_lines, _ = extract_page_lines(page)
    results = []
    text_dict = page.get_text("dict")
    blocks = text_dict.get("blocks", []) if isinstance(text_dict, dict) else []

    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text or text in {"-", "－", "_"}:
                    continue
                b = fitz.Rect(span["bbox"])
                if b.height < 4:
                    continue

                # 字体核心高度区域（去掉上下自带的留白）
                pad = b.height * 0.16
                core_y0 = b.y0 + pad
                core_y1 = b.y1 - pad

                for ly, lx0, lx1 in h_lines:
                    if b.x1 > lx0 + 1.5 and b.x0 < lx1 - 1.5:
                        # 水平线穿过了文字核心区，或紧贴底部压线
                        if core_y0 < ly < core_y1 + 0.6:
                            results.append({
                                "text": text,
                                "bbox": (b.x0, b.y0, b.x1, b.y1),
                                "line_y": ly,
                            })
                            break
    return results


def render_crop_base64(page: fitz.Page, bbox: tuple[float, float, float, float], highlight: bool = True) -> str:
    """将 PDF 指定坐标区域放大截图并转为 Base64 PNG（不污染原页面内容）"""
    x0, y0, x1, y1 = bbox
    # 向四周扩展视野，方便看清上下文单元格
    clip_rect = fitz.Rect(x0 - 80, y0 - 35, x1 + 80, y1 + 25) & page.rect
    annot = None
    try:
        if highlight:
            annot = page.add_rect_annot(fitz.Rect(x0 - 2, y0 - 2, x1 + 2, y1 + 2))
            annot.set_colors(stroke=(1, 0, 0))
            annot.set_border(width=1)
            annot.update()

        pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2), clip=clip_rect, annots=True)
    finally:
        if annot is not None:
            try:
                page.delete_annot(annot)
            except Exception:
                pass

    png_bytes = pix.tobytes("png")
    return "data:image/png;base64," + base64.b64encode(png_bytes).decode("ascii")


def _normalize_used_values(vals) -> list[tuple]:
    if vals is None:
        return []
    if not isinstance(vals, tuple):
        return [(vals,)]
    if vals and not isinstance(vals[0], tuple):
        return [vals]
    return list(vals)


def _col_index_to_letters(col_idx: int) -> str:
    result = []
    n = max(1, col_idx)
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result.append(chr(ord("A") + rem))
    return "".join(reversed(result))


def extract_sheet_snapshot(sheet) -> dict:
    """一次性提取工作表 UsedRange 的文本快照，避免每个切字项重复跨进程读取整表"""
    try:
        used = sheet.UsedRange
        start_r = int(used.Row)
        start_c = int(used.Column)
        rows_vals = _normalize_used_values(used.Value)
    except Exception:
        start_r = 1
        start_c = 1
        rows_vals = []

    return {
        "sheet": sheet,
        "start_r": start_r,
        "start_c": start_c,
        "rows_vals": rows_vals,
        "row_heights": {},
    }


def match_excel_row(
    sheet_snapshot: dict,
    text_snippet: str,
    page_index: int,
    total_pages: int,
    rel_y: float,
) -> tuple[int, str, float]:
    """
    根据 PDF 中被切断的文字片段和跨页纵向位置，反推对应的 Excel 行号与单元格地址。
    返回: (row_index, cell_address, current_row_height)
    """
    sheet = sheet_snapshot["sheet"]
    start_r = sheet_snapshot["start_r"]
    start_c = sheet_snapshot["start_c"]
    rows_vals = sheet_snapshot["rows_vals"]
    row_heights: dict[int, float] = sheet_snapshot["row_heights"]

    candidates: list[tuple[int, int]] = []
    clean_snip = text_snippet.strip()
    compact_snip = re.sub(r"\s+", "", clean_snip)

    total_r = max(1, len(rows_vals))
    for r_idx, row_tuple in enumerate(rows_vals):
        excel_r = start_r + r_idx
        for c_idx, cell_val in enumerate(row_tuple):
            if cell_val is None:
                continue
            s_val = str(cell_val)
            compact_val = re.sub(r"\s+", "", s_val)
            if (
                (clean_snip and clean_snip in s_val)
                or (len(compact_snip) >= 2 and compact_snip in compact_val)
            ):
                excel_c = start_c + c_idx
                candidates.append((excel_r, excel_c))

    # 结合页码计算在整个工作表中的全局纵向比例
    global_rel_y = min(1.0, max(0.0, (page_index + rel_y) / max(1, total_pages)))
    approx_r = start_r + int(total_r * global_rel_y)

    target_c = start_c
    if candidates:
        candidates.sort(key=lambda rc: abs(rc[0] - approx_r))
        target_r, target_c = candidates[0]
    else:
        target_r = max(1, approx_r)

    cell_addr = f"{_col_index_to_letters(target_c)}{target_r}"

    if target_r not in row_heights:
        try:
            row_h = float(sheet.Rows(target_r).RowHeight or 15.0)
        except Exception:
            row_h = 15.0
        row_heights[target_r] = round(row_h, 1)

    return target_r, cell_addr, row_heights[target_r]


def _texts_match(text_a: str, text_b: str) -> bool:
    a = re.sub(r"\s+", "", text_a.strip())
    b = re.sub(r"\s+", "", text_b.strip())
    if not a or not b:
        return False
    if a in b or b in a:
        return True
    # 如果因为重新折行导致首尾字符略有变化，检查是否有至少2个连续公共字符
    if len(a) >= 2 and len(b) >= 2:
        for i in range(len(a) - 1):
            if a[i : i + 2] in b:
                return True
    return False


def _as_bbox(val) -> tuple[float, float, float, float]:
    return (float(val[0]), float(val[1]), float(val[2]), float(val[3]))


def is_region_still_clipped(
    pdf_path: Path,
    page_index: int,
    orig_text: str,
    orig_bbox: tuple[float, float, float, float],
) -> tuple[bool, str, tuple[float, float, float, float]]:
    """
    重新导出 PDF 后，检查原来那个单元格区域内的对应文字是否还压线，并返回最新截图 Base64 与最新坐标
    """
    orig_x0, orig_y0, orig_x1, orig_y1 = orig_bbox
    doc = fitz.open(pdf_path)
    try:
        if len(doc) == 0:
            return False, "", orig_bbox
        if page_index >= len(doc):
            page_index = len(doc) - 1
        page = doc[page_index]
        clipped_list = find_clipped_spans_on_page(page)

        still_clipped = False
        target_bbox: tuple[float, float, float, float] | None = None

        matched_clips: list[tuple[float, tuple[float, float, float, float]]] = []
        for item in clipped_list:
            bx0, by0, bx1, _ = item["bbox"]
            # 1. X 坐标必须在同一列范围内
            if not (bx1 >= orig_x0 - 15 and bx0 <= orig_x1 + 15):
                continue
            y_dist = abs(by0 - orig_y0)
            # 2. 文本内容匹配且纵向距离合理，或纵向几乎在同一位置（处理折行后末行文字微变的情况）
            if (_texts_match(item["text"], orig_text) and y_dist <= 100.0) or y_dist <= 20.0:
                matched_clips.append((y_dist, item["bbox"]))

        if matched_clips:
            matched_clips.sort(key=lambda pair: pair[0])
            still_clipped = True
            target_bbox = matched_clips[0][1]

        # 如果不压线了，找到该文字现在的新位置截图给用户看（含跨到下一页的兜底查找）
        render_page = page
        if target_bbox is None:
            candidate_pages = [page_index]
            if page_index + 1 < len(doc):
                candidate_pages.append(page_index + 1)

            for cand_p_idx in candidate_pages:
                cand_page = doc[cand_p_idx]
                cand_dict = cand_page.get_text("dict")
                cand_blocks = cand_dict.get("blocks", []) if isinstance(cand_dict, dict) else []
                span_matches: list[tuple[float, tuple[float, float, float, float]]] = []
                for block in cand_blocks:
                    if not isinstance(block, dict) or block.get("type") != 0:
                        continue
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            b = fitz.Rect(span["bbox"])
                            if b.x1 >= orig_x0 - 15 and b.x0 <= orig_x1 + 15:
                                span_text = span.get("text", "")
                                y_dist = abs(b.y0 - orig_y0) if cand_p_idx == page_index else b.y0
                                if _texts_match(span_text, orig_text) and (cand_p_idx != page_index or y_dist <= 120.0):
                                    span_matches.append((y_dist, (b.x0, b.y0, b.x1, b.y1)))
                if span_matches:
                    span_matches.sort(key=lambda pair: pair[0])
                    target_bbox = span_matches[0][1]
                    render_page = cand_page
                    break

        if target_bbox is None:
            target_bbox = (orig_x0, orig_y0, orig_x1, orig_y1)

        new_img = render_crop_base64(render_page, target_bbox, highlight=still_clipped)
        return still_clipped, new_img, target_bbox
    finally:
        doc.close()


def normalize_workbook_view(excel, wb) -> None:
    """A1 定位 + 100% 缩放 + 切回第一页"""
    if wb.Windows.Count > 0:
        wb.Windows(1).Visible = True
    first_sheet = None
    for sh in wb.Sheets:
        if sh.Visible != XL_SHEET_VISIBLE:
            continue
        if first_sheet is None:
            first_sheet = sh
        try:
            sh.Activate()
            if sh.Type == XL_WORKSHEET:
                excel.Goto(sh.Range("A1"), Scroll=True)
            excel.ActiveWindow.Zoom = 100
        except Exception:
            pass
    if first_sheet is not None:
        try:
            first_sheet.Activate()
        except Exception:
            pass


def _unique_pdf_path(pdf_folder: Path, file_stem: str, safe_sheet_name: str, used_paths: set[Path]) -> Path:
    candidate = pdf_folder / f"{file_stem}_{safe_sheet_name}.pdf"
    counter = 2
    while candidate in used_paths:
        candidate = pdf_folder / f"{file_stem}_{safe_sheet_name}_{counter}.pdf"
        counter += 1
    used_paths.add(candidate)
    return candidate


def scan_folder_workflow(folder_str: str, delete_backups: bool) -> None:
    with STATE_LOCK:
        STATE["busy"] = True
        STATE["items"].clear()
        STATE["logs"].clear()

    try:
        folder = Path(folder_str.strip().strip('"').strip("'")).expanduser().resolve()
        if not folder.is_dir():
            log_msg(f"❌ 文件夹不存在: {folder}")
            return

        workbooks = sorted(
            p for p in folder.iterdir()
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

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        item_seq = 1
        try:
            for wb_path in workbooks:
                log_msg(f"正在处理并导出: {wb_path.name} ...")
                wb = None
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
                            try:
                                append_record(wb_path.name, all_names, bkp_names)
                            except OSError as err:
                                log_msg(f"  ⚠️ 清理日志写入失败 {wb_path.name}: {err}")
                            log_msg(f"  已清理 {len(bkp_names)} 个日期备份 Sheet")

                    file_stem = sanitize_filename(wb_path.stem)
                    for sheet in wb.Sheets:
                        s_name = sheet.Name
                        if s_name in EXCLUDE_PDF_SHEETS or is_backup_sheet(s_name):
                            continue
                        if sheet.Visible != XL_SHEET_VISIBLE:
                            continue

                        pdf_path = _unique_pdf_path(
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
                        except Exception as e:
                            log_msg(f"  导出 PDF 失败 [{s_name}]: {e}")
                            continue

                        sheet_snapshot = extract_sheet_snapshot(sheet)

                        # 扫描刚导出的 PDF
                        doc = fitz.open(pdf_path)
                        try:
                            total_pages = len(doc)
                            for p_idx in range(total_pages):
                                page = doc[p_idx]
                                clipped = find_clipped_spans_on_page(page)
                                for clip_info in clipped:
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
                except Exception as err:
                    log_msg(f"❌ 处理文件出错 {wb_path.name}: {err}")
                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        with STATE_LOCK:
            total_detected = len(STATE["items"])
        log_msg(f"✅ 扫描完成！共检测到 {total_detected} 处疑似切字位置。")
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


def fix_selected_workflow(selected_ids: list[str], step_pt: float = 3.0, max_iter: int = 6) -> None:
    """
    针对用户勾选的卡片：
    每次将对应 Excel 行高 +step_pt -> 重新导出该 Sheet 的 PDF -> 检测该处是否还切字
    一旦完整显示则立即停止！最后执行 A1 + 100% 缩放保存。
    """
    with STATE_LOCK:
        STATE["busy"] = True
        by_wb: dict[str, list[dict]] = {}
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
                wb = None
                try:
                    wb = excel.Workbooks.Open(str(wb_path.resolve()))
                    adjusted_rows_in_batch: set[tuple[str, int]] = set()

                    for item in items:
                        sheet_name = item["sheet_name"]
                        sheet = wb.Sheets(sheet_name)
                        target_row = max(1, int(item["excel_row"]))
                        pdf_path = Path(item["pdf_path"])
                        row_key = (sheet_name, target_row)

                        # 先检查是否因为同行/同表的微调已经顺带修好了
                        if (item["iterations"] > 0 or row_key in adjusted_rows_in_batch) and pdf_path.exists():
                            still_clip, new_img, new_bbox = is_region_still_clipped(
                                pdf_path, item["page_index"], item["text"], _as_bbox(item["bbox"])
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

                            # 重新导出该 Sheet 的 PDF
                            sheet.ExportAsFixedFormat(
                                Type=XL_TYPE_PDF,
                                Filename=str(pdf_path.resolve()),
                                Quality=0,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False,
                            )

                            still_clip, new_img, new_bbox = is_region_still_clipped(
                                pdf_path, item["page_index"], item["text"], _as_bbox(item["bbox"])
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

                    # 顺便完成 A1 定位与 100% 缩放保存
                    normalize_workbook_view(excel, wb)
                    wb.Save()
                    log_msg(f"💾 已保存并完成 A1/100% 标准化: {wb_path.name}")
                except Exception as err:
                    log_msg(f"❌ 修复失败 {wb_path.name}: {err}")
                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        log_msg("🎉 所有选中项微调完毕！可在网页上查看最新截图效果。")
    finally:
        with STATE_LOCK:
            STATE["busy"] = False


HTML_PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>Excel PDF 切字检测与智能微调工具</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Yu Gothic UI", sans-serif; margin: 0; background: #f4f6f9; color: #1e293b; }
  header { background: #1e293b; color: #fff; padding: 16px 28px; display: flex; align-items: center; justify-content: space-between; }
  header h1 { margin: 0; font-size: 18px; font-weight: 600; }
  .container { max-width: 1400px; margin: 20px auto; padding: 0 20px; }
  .panel { background: #fff; border-radius: 10px; padding: 18px 22px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); margin-bottom: 18px; }
  .controls { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
  input[type="text"], input[type="number"] { padding: 9px 12px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 14px; }
  input[type="text"] { flex: 1; min-width: 320px; }
  button { padding: 9px 16px; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; transition: 0.15s; }
  .btn-primary { background: #2563eb; color: #fff; }
  .btn-primary:hover { background: #1d4ed8; }
  .btn-success { background: #059669; color: #fff; }
  .btn-success:hover { background: #047857; }
  .btn-outline { background: #e2e8f0; color: #334155; }
  button:disabled { opacity: 0.5; cursor: not-allowed; }
  .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(360px, 1fr)); gap: 16px; }
  .card { background: #fff; border: 2px solid #e2e8f0; border-radius: 10px; overflow: hidden; cursor: pointer; transition: 0.15s; position: relative; }
  .card.selected { border-color: #2563eb; box-shadow: 0 0 0 3px rgba(37,99,235,0.2); }
  .card.fixed { border-color: #10b981; }
  .card.max-reached { border-color: #f59e0b; }
  .card-header { padding: 10px 14px; background: #f8fafc; border-bottom: 1px solid #e2e8f0; font-size: 13px; display: flex; justify-content: space-between; align-items: center; }
  .badge { padding: 3px 8px; border-radius: 4px; font-size: 12px; font-weight: 600; }
  .badge-pending { background: #fef3c7; color: #b45309; }
  .badge-fixed { background: #d1fae5; color: #065f46; }
  .badge-warn { background: #fee2e2; color: #991b1b; }
  .card-img-wrap { padding: 12px; text-align: center; background: #fff; min-height: 110px; display: flex; align-items: center; justify-content: center; }
  .card-img-wrap img { max-width: 100%; max-height: 160px; border: 1px solid #e2e8f0; border-radius: 4px; }
  .card-body { padding: 10px 14px; font-size: 13px; color: #475569; border-top: 1px solid #f1f5f9; display: flex; justify-content: space-between; align-items: center; }
  .row-adjust { display: flex; align-items: center; gap: 6px; }
  .row-adjust input { width: 64px; padding: 4px 6px; text-align: center; }
  .log-box { background: #0f172a; color: #e2e8f0; font-family: monospace; font-size: 12px; padding: 12px; border-radius: 8px; max-height: 160px; overflow-y: auto; }
</style>
</head>
<body>
<header>
  <h1>📐 Excel 导出 PDF 切字检测 & 闭环行高微调器</h1>
  <span id="statusText">就绪</span>
</header>
<div class="container">
  <div class="panel">
    <div class="controls">
      <input type="text" id="folderInput" placeholder="粘贴 Excel 文件夹路径，例如 C:\\Users\\...\\Desktop\\excel_files">
      <label style="font-size:13px;"><input type="checkbox" id="delBackupChk" checked> 顺便清理 YYYYMMDD 备份表</label>
      <button class="btn-primary" id="scanBtn" onclick="startScan()">1. 导出 PDF 并扫描切字</button>
    </div>
  </div>

  <div class="panel" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
    <div>
      <button class="btn-outline" onclick="toggleAll(true)">全选待处理</button>
      <button class="btn-outline" onclick="toggleAll(false)">取消全选</button>
      <span style="margin-left:12px; font-size:14px;">已选中 <b id="selCount">0</b> 项</span>
    </div>
    <div style="display:flex; align-items:center; gap:10px;">
      <label style="font-size:13px;">每次递增(pt): <input type="number" id="stepInput" value="3.0" step="0.5" style="width:65px;"></label>
      <button class="btn-success" id="fixBtn" onclick="startFix()">2. 自动迭代微调选中项 (直到完整显示)</button>
    </div>
  </div>

  <div class="panel">
    <div class="log-box" id="logBox">等待操作...</div>
  </div>

  <div class="grid" id="cardGrid"></div>
</div>

<script>
let selectedIds = new Set();
let itemsCache = {};

async function pollState() {
  try {
    const res = await fetch('/api/state');
    const data = await res.json();
    document.getElementById('scanBtn').disabled = data.busy;
    document.getElementById('fixBtn').disabled = data.busy;
    document.getElementById('statusText').innerText = data.busy ? '⏳ 后台处理中...' : '✅ 就绪';

    const logBox = document.getElementById('logBox');
    logBox.innerText = data.logs.join('\\n') || '等待操作...';
    logBox.scrollTop = logBox.scrollHeight;

    renderCards(data.items);
  } catch (e) {}
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function renderCards(items) {
  const grid = document.getElementById('cardGrid');
  const keys = Object.keys(items);
  if (keys.length === 0) {
    grid.innerHTML = '';
    itemsCache = {};
    updateCount();
    return;
  }

  keys.forEach(id => {
    const item = items[id];
    let card = document.getElementById('card_' + id);
    if (!card) {
      card = document.createElement('div');
      card.id = 'card_' + id;
      card.className = 'card';
      card.onclick = (e) => {
        if (e.target.tagName === 'INPUT') return;
        if (selectedIds.has(id)) selectedIds.delete(id);
        else selectedIds.add(id);
        card.classList.toggle('selected', selectedIds.has(id));
        updateCount();
      };
      grid.appendChild(card);
      // 默认自动选中新扫出来的待处理项
      if (item.status === 'pending') {
        selectedIds.add(id);
        card.classList.add('selected');
      }
    }

    if (item.status === 'fixed') {
      selectedIds.delete(id);
      card.classList.remove('selected', 'max-reached');
      card.classList.add('fixed');
    } else if (item.status === 'max_reached') {
      card.classList.remove('fixed');
      card.classList.add('max-reached');
    }

    // 如果用户正在当前卡片的行号输入框内输入，跳过 DOM 重建以防焦点丢失
    const activeEl = document.activeElement;
    if (activeEl && activeEl.tagName === 'INPUT' && card.contains(activeEl)) {
      return;
    }

    const signature = `${item.status}|${item.iterations}|${item.row_height}|${item.excel_row}|${item.cell_addr}|${item.image}`;
    if (itemsCache[id] === signature) {
      return;
    }
    itemsCache[id] = signature;

    let badgeClass = 'badge-pending';
    let badgeText = '疑似压线/切字';
    if (item.status === 'fixed') {
      badgeClass = 'badge-fixed';
      badgeText = `✓ 已修复 (+${item.iterations}次)`;
    } else if (item.status === 'max_reached') {
      badgeClass = 'badge-warn';
      badgeText = `⚠️ 达最大微调次数 (+${item.iterations}次)`;
    }

    card.innerHTML = `
      <div class="card-header">
        <div><b>${escapeHtml(item.wb_name)}</b> <span style="color:#64748b;">/ ${escapeHtml(item.sheet_name)}</span></div>
        <span class="badge ${badgeClass}">${escapeHtml(badgeText)}</span>
      </div>
      <div class="card-img-wrap"><img src="${item.image}" alt="crop"></div>
      <div class="card-body">
        <div>定位: <b>${escapeHtml(item.cell_addr)}</b> | 当前行高: <b>${escapeHtml(item.row_height)}pt</b><br><span style="color:#94a3b8;">文字: ${escapeHtml(item.text)}</span></div>
        <div class="row-adjust" title="如果行号有偏差可手动微调">
          行号: <input type="number" min="1" value="${Number(item.excel_row) || 1}" onchange="updateRow('${escapeHtml(id)}', this.value)">
        </div>
      </div>
    `;
  });
  updateCount();
}

function updateCount() {
  document.getElementById('selCount').innerText = selectedIds.size;
}

function toggleAll(select) {
  document.querySelectorAll('.card').forEach(c => {
    const id = c.id.replace('card_', '');
    if (select && !c.classList.contains('fixed')) {
      selectedIds.add(id);
      c.classList.add('selected');
    } else {
      selectedIds.delete(id);
      c.classList.remove('selected');
    }
  });
  updateCount();
}

async function updateRow(id, newRow) {
  const parsed = parseInt(newRow, 10);
  if (!Number.isFinite(parsed) || parsed < 1) return;
  await fetch('/api/update_row', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({id, excel_row: parsed})
  });
}

async function startScan() {
  const folder = document.getElementById('folderInput').value.trim();
  if (!folder) { alert('请先输入 Excel 文件夹路径！'); return; }
  selectedIds.clear();
  itemsCache = {};
  document.getElementById('cardGrid').innerHTML = '';
  await fetch('/api/scan', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      folder,
      delete_backups: document.getElementById('delBackupChk').checked
    })
  });
}

async function startFix() {
  if (selectedIds.size === 0) { alert('请先点击卡片选中要修复的位置！'); return; }
  const step = parseFloat(document.getElementById('stepInput').value) || 3.0;
  await fetch('/api/fix', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      selected_ids: Array.from(selectedIds),
      step_pt: step
    })
  });
}

setInterval(pollState, 1200);
</script>
</body>
</html>
"""


class WebHandler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: object) -> None:
        del format, args  # 保持终端清爽

    def _send_json(self, data: dict) -> None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = HTML_PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif parsed.path == "/api/state":
            with STATE_LOCK:
                snapshot = {
                    "busy": STATE["busy"],
                    "logs": list(STATE["logs"][-30:]),
                    "items": copy.deepcopy(STATE["items"]),
                }
            self._send_json(snapshot)
        else:
            self.send_error(404)

    def do_POST(self) -> None:
        length = int(self.headers.get("Content-Length", 0))
        payload = json.loads(self.rfile.read(length).decode("utf-8")) if length else {}

        if self.path == "/api/scan":
            if try_acquire_busy():
                threading.Thread(
                    target=scan_folder_workflow,
                    args=(payload.get("folder", ""), bool(payload.get("delete_backups", True))),
                    daemon=True,
                ).start()
            self._send_json({"ok": True})

        elif self.path == "/api/fix":
            if try_acquire_busy():
                try:
                    step_pt = max(0.5, float(payload.get("step_pt", 3.0)))
                except (TypeError, ValueError):
                    step_pt = 3.0
                threading.Thread(
                    target=fix_selected_workflow,
                    args=(payload.get("selected_ids", []), step_pt),
                    daemon=True,
                ).start()
            self._send_json({"ok": True})

        elif self.path == "/api/update_row":
            iid = payload.get("id")
            try:
                new_row = max(1, int(payload.get("excel_row", 1)))
            except (TypeError, ValueError):
                self._send_json({"ok": False, "error": "invalid excel_row"})
                return

            with STATE_LOCK:
                if iid in STATE["items"]:
                    item = STATE["items"][iid]
                    item["excel_row"] = new_row
                    old_addr = str(item.get("cell_addr", ""))
                    item["cell_addr"] = (
                        re.sub(r"\d+$", str(new_row), old_addr)
                        if re.search(r"\d+$", old_addr)
                        else f"Row {new_row}"
                    )
            self._send_json({"ok": True})


def main() -> None:
    port = 8765
    server = ThreadingHTTPServer(("127.0.0.1", port), WebHandler)
    url = f"http://127.0.0.1:{port}"
    print(f"🚀 Web 界面已启动: {url}")
    print("正在自动打开浏览器... (按 Ctrl+C 可退出程序)")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出。")


if __name__ == "__main__":
    main()
