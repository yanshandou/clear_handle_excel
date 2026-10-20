import re
from datetime import datetime
from pathlib import Path
from typing import Any

DATE_PATTERN = re.compile(r"(?<!\d)((?:19|20)\d{6})(?!\d)")
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
EXCLUDE_PDF_SHEETS = {"変更履歴"}  # 不导出为 PDF 的 Sheet 名集合
PROJECT_ROOT = Path(__file__).resolve().parent.parent
RECORD_PATH = PROJECT_ROOT / "record.md"
STATIC_DIR = PROJECT_ROOT / "static"

# Excel COM 常量
XL_TYPE_PDF = 0
XL_SHEET_VISIBLE = -1
XL_WORKSHEET = -4167


def is_backup_sheet(sheet_name: str) -> bool:
    """判断工作表名称是否包含有效的 YYYYMMDD 日期格式"""
    for match in DATE_PATTERN.finditer(sheet_name):
        try:
            datetime.strptime(match.group(1), "%Y%m%d")
            return True
        except ValueError:
            continue
    return False


def sanitize_filename(name: str) -> str:
    """将工作簿或工作表名中的非法文件名字符替换为下划线"""
    return re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name).strip()


def unique_pdf_path(
    pdf_folder: Path,
    file_stem: str,
    safe_sheet_name: str,
    used_paths: set[Path],
) -> Path:
    """生成不冲突的 PDF 导出路径，避免同名工作表覆盖"""
    candidate = pdf_folder / f"{file_stem}_{safe_sheet_name}.pdf"
    counter = 2
    while candidate in used_paths:
        candidate = pdf_folder / f"{file_stem}_{safe_sheet_name}_{counter}.pdf"
        counter += 1
    used_paths.add(candidate)
    return candidate


def append_record(
    workbook_name: str,
    original_sheet_names: list[str],
    removed_sheet_names: list[str],
) -> None:
    """将删除备份 Sheet 的记录追加写入 record.md"""
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


def normalize_used_values(vals: Any) -> list[tuple[Any, ...]]:
    """将 Excel UsedRange.Value 统一归一化为二维行元组列表"""
    if vals is None:
        return []
    if not isinstance(vals, tuple):
        return [(vals,)]
    if vals and not isinstance(vals[0], tuple):
        return [vals]
    return list(vals)


def col_index_to_letters(col_idx: int) -> str:
    """将 1-based 列号转换为 Excel 列字母（如 1 -> A, 27 -> AA）"""
    result: list[str] = []
    n = max(1, col_idx)
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result.append(chr(ord("A") + rem))
    return "".join(reversed(result))


def as_bbox(val: Any) -> tuple[float, float, float, float]:
    """将四元序列转换为定长浮点坐标元组 (x0, y0, x1, y1)"""
    return (float(val[0]), float(val[1]), float(val[2]), float(val[3]))
