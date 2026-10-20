import re
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

from core.common import (
    EXCLUDE_PDF_SHEETS,
    XL_SHEET_VISIBLE,
    XL_WORKSHEET,
    col_index_to_letters,
    is_backup_sheet,
    normalize_used_values,
)

try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    print("请先安装 pywin32: pip install pywin32")
    sys.exit(1)

try:
    from openpyxl import load_workbook
except ImportError:
    print("请先安装 openpyxl: pip install openpyxl")
    sys.exit(1)


def open_workbook_for_manual_review(
    workbook_path: Path, sheet_name: str, cell_address: str, row: int
) -> float:
    """定位到可见 Excel，交给用户编辑；不保存、不关闭用户的工作簿。"""
    workbook_path = workbook_path.resolve(strict=True)
    pythoncom.CoInitialize()
    try:
        try:
            excel = win32.GetActiveObject("Excel.Application")
        except pythoncom.com_error as err:
            if err.args[0] != -2147221021:  # MK_E_UNAVAILABLE: 没有正在运行的 Excel
                raise
            excel = win32.DispatchEx("Excel.Application")
        workbook = next(
            (book for book in excel.Workbooks
             if Path(book.FullName).resolve() == workbook_path),
            None,
        )
        if workbook is None:
            workbook = excel.Workbooks.Open(str(workbook_path))
        sheet = workbook.Sheets(sheet_name)
        excel.Visible = True
        excel.WindowState = -4143  # xlNormal
        if sheet.Visible != XL_SHEET_VISIBLE:
            raise ValueError("目标工作表已隐藏，请先在 Excel 中取消隐藏后再定位")
        workbook.Activate()
        sheet.Activate()
        if workbook.Windows.Count > 0:
            workbook.Windows(1).Visible = True
            workbook.Windows(1).WindowState = -4143  # xlNormal
        address = cell_address if re.fullmatch(r"[A-Za-z]+[1-9]\d*", cell_address) else f"A{row}"
        excel.Goto(sheet.Range(address), Scroll=True)
        return round(float(sheet.Rows(row).RowHeight or 15.0), 1)
    finally:
        pythoncom.CoUninitialize()


def scan_workbook_sheets(workbook_path: Path) -> tuple[list[str], list[str]]:
    """扫描工作簿中的所有 Sheet 名与包含日期的备份 Sheet 名（支持 .xlsx/.xlsm/.xls）"""
    if workbook_path.suffix.lower() == ".xls":
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = None
        try:
            wb = excel.Workbooks.Open(str(workbook_path.resolve()), ReadOnly=True)
            all_sheets = [sh.Name for sh in wb.Sheets]
            backup_sheets = [name for name in all_sheets if is_backup_sheet(name)]
            return all_sheets, backup_sheets
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()
            pythoncom.CoUninitialize()

    workbook = load_workbook(workbook_path, read_only=True)
    try:
        all_sheets = list(workbook.sheetnames)
        backup_sheets = [name for name in all_sheets if is_backup_sheet(name)]
        return all_sheets, backup_sheets
    finally:
        workbook.close()


def _is_sheet_already_height_adjusted(sheet: Any) -> bool:
    try:
        for prop in sheet.CustomProperties:
            if prop.Name == "_clip_height_adjusted":
                return True
    except Exception:
        pass
    return False


def _mark_sheet_height_adjusted(sheet: Any) -> None:
    try:
        sheet.CustomProperties.Add("_clip_height_adjusted", "1")
    except Exception:
        pass


def adjust_row_heights_against_clipping(wb: Any, extra_points: float = 6.0) -> None:
    """
    智能防切字调整：
    不破坏单行表头或方眼紙格子，只针对【包含非空文本且原本行高 >= 22 磅的多行文本行】
    增加安全余量（默认 +6.0 磅），并通过自定义属性保证幂等性。
    """
    for sheet in wb.Sheets:
        if sheet.Visible != XL_SHEET_VISIBLE or sheet.Type != XL_WORKSHEET:
            continue
        if sheet.Name in EXCLUDE_PDF_SHEETS or is_backup_sheet(sheet.Name):
            continue
        if _is_sheet_already_height_adjusted(sheet):
            continue

        try:
            used_range = sheet.UsedRange
            row_count = used_range.Rows.Count
            if row_count > 1500:
                continue

            rows_values = normalize_used_values(used_range.Value)
            adjusted_any = False

            for r_idx in range(1, row_count + 1):
                if r_idx <= len(rows_values):
                    row_vals = rows_values[r_idx - 1]
                    if not any(cell is not None and str(cell).strip() != "" for cell in row_vals):
                        continue

                row = used_range.Rows(r_idx)
                current_h = row.RowHeight
                if current_h is not None and 22.0 <= current_h < 300.0:
                    row.RowHeight = round(float(current_h) + extra_points, 1)
                    adjusted_any = True

            if adjusted_any:
                _mark_sheet_height_adjusted(sheet)
        except Exception:
            pass


def normalize_workbook_view(
    excel: Any,
    wb: Any,
    on_error: Callable[[str], None] | None = None,
) -> None:
    """将所有可见工作表定位到左上角 A1、缩放设为 100%，并停留在第一个可见 Sheet"""
    if wb.Windows.Count > 0:
        wb.Windows(1).Visible = True

    first_visible_sheet = None

    for sheet in wb.Sheets:
        if sheet.Visible != XL_SHEET_VISIBLE:
            continue
        if first_visible_sheet is None:
            first_visible_sheet = sheet

        try:
            sheet.Activate()
            if sheet.Type == XL_WORKSHEET:
                excel.Goto(sheet.Range("A1"), Scroll=True)
            excel.ActiveWindow.Zoom = 100
        except Exception as err:
            if on_error is not None:
                on_error(f"  视图重置跳过 [{wb.Name} -> {sheet.Name}]: {err}")

    if first_visible_sheet is not None:
        try:
            first_visible_sheet.Activate()
        except Exception:
            pass


def extract_sheet_snapshot(sheet: Any) -> dict[str, Any]:
    """一次性提取工作表 UsedRange 的文本快照，避免每个切字项重复跨进程读取整表"""
    try:
        used = sheet.UsedRange
        start_r = int(used.Row)
        start_c = int(used.Column)
        rows_vals = normalize_used_values(used.Value)
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
    sheet_snapshot: dict[str, Any],
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

    global_rel_y = min(1.0, max(0.0, (page_index + rel_y) / max(1, total_pages)))
    approx_r = start_r + int(total_r * global_rel_y)

    target_c = start_c
    if candidates:
        candidates.sort(key=lambda rc: abs(rc[0] - approx_r))
        target_r, target_c = candidates[0]
    else:
        target_r = max(1, approx_r)

    cell_addr = f"{col_index_to_letters(target_c)}{target_r}"

    if target_r not in row_heights:
        try:
            row_h = float(sheet.Rows(target_r).RowHeight or 15.0)
        except Exception:
            row_h = 15.0
        row_heights[target_r] = round(row_h, 1)

    return target_r, cell_addr, row_heights[target_r]
