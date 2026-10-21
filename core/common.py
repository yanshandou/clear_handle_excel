import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

DATE_PATTERN = re.compile(r"(?<!\d)((?:19|20)\d{6})(?!\d)")
# 独立出现的英文备份标记（前后不是字母），如 _bak、Sheet1-BAK、一覧_old；Bakery、Threshold 不算
BACKUP_WORD_PATTERN = re.compile(r"(?<![a-z])(?:bak|backup|old|copy)(?![a-z])", re.IGNORECASE)
# 日文/中文备份标记；「旧」不含「新旧」「復旧」（如 新旧対照表、障害復旧手順 是正式资料）
BACKUP_TEXT_PATTERN = re.compile(r"バックアップ|备份|コピー|副本|(?<![新復])旧")
# Excel 复制工作表时自动追加的「 (2)」后缀；不带空格的「一覧(2)」多为续页，不算
EXCEL_COPY_SUFFIX_PATTERN = re.compile(r"\s\(\d+\)$")
# 去掉备份标记、找原表时额外去掉的写法（「Sheet1のコピー」的「の」、「Copy of 一覧」的「of」）
COPY_PHRASE_PATTERN = re.compile(r"copy\s+of|の?コピー|的?副本", re.IGNORECASE)
# 版本号（如 _v2.4、ver1.0），比对原表时忽略
VERSION_PATTERN = re.compile(r"(?<![a-z])v(?:er)?\.?\d+(?:\.\d+)*(?![a-z])", re.IGNORECASE)
# 比对原表时忽略的空白、分隔符、括号与待定标记
UNCERTAIN_EMOJI = "❓"
EXCEL_SHEET_MAX_LEN = 31
NAME_SEPARATOR_PATTERN = re.compile(r"[\s_\-・.,、。()\[\]【】「」『』❓]+")
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
EXCLUDE_PDF_SHEETS = {"変更履歴"}  # 不导出为 PDF 的 Sheet 名集合
PROJECT_ROOT = Path(__file__).resolve().parent.parent
RECORD_PATH = PROJECT_ROOT / "record.md"
STATIC_DIR = PROJECT_ROOT / "static"

# Excel COM 常量
XL_TYPE_PDF = 0
XL_SHEET_VISIBLE = -1
XL_WORKSHEET = -4167


def console_print(text: str) -> None:
    """打印到终端，不因编码失败抛出异常。

    Python 3.15 之前输出到管道（如 Git Bash）时 stdout 使用系统 ANSI 代码页（cp932/GBK 等），
    无法编码 emoji 和部分中文；此时转义成 \\uXXXX 输出，避免日志打印中断后台任务。
    """
    try:
        print(text)
    except UnicodeEncodeError:
        encoding = getattr(sys.stdout, "encoding", None) or "ascii"
        print(text.encode(encoding, errors="backslashreplace").decode(encoding))


def _is_valid_date(text: str) -> bool:
    try:
        datetime.strptime(text, "%Y%m%d")
        return True
    except ValueError:
        return False


def has_backup_marker(sheet_name: str) -> bool:
    """名称是否带备份标记：有效 YYYYMMDD 日期、bak/backup/old/copy、
    バックアップ/备份/旧/コピー/副本，或 Excel 复制产生的「 (2)」后缀。

    只是候选条件，是否真的是备份表见 find_backup_sheets。
    """
    return bool(
        BACKUP_WORD_PATTERN.search(sheet_name)
        or BACKUP_TEXT_PATTERN.search(sheet_name)
        or EXCEL_COPY_SUFFIX_PATTERN.search(sheet_name)
        or any(_is_valid_date(m.group(1)) for m in DATE_PATTERN.finditer(sheet_name))
    )


def _comparable_name(name: str) -> str:
    """用于比对原表的名称：统一全角半角与大小写，忽略版本号、空白与分隔符"""
    text = unicodedata.normalize("NFKC", name).lower()
    text = VERSION_PATTERN.sub("", text)
    return NAME_SEPARATOR_PATTERN.sub("", text)


def _original_name_of(backup_name: str) -> str:
    """去掉备份标记后的原表名称，如 基本情報_20241003 → 基本情報"""
    text = unicodedata.normalize("NFKC", backup_name)
    text = DATE_PATTERN.sub(lambda m: "" if _is_valid_date(m.group(1)) else m.group(0), text)
    text = EXCEL_COPY_SUFFIX_PATTERN.sub("", text)
    text = COPY_PHRASE_PATTERN.sub("", text)
    text = BACKUP_WORD_PATTERN.sub("", text)
    text = BACKUP_TEXT_PATTERN.sub("", text)
    return _comparable_name(text)


def find_backup_sheets(sheet_names: list[str]) -> list[str]:
    """找出工作簿中的备份表：名称带备份标记，且去掉标记后能在同一工作簿找到对应的原表。

    找不到原表时（如只有 基本情報_20241003 而没有 基本情報），带标记的表就是当前在用的版本，
    保留不删，导出 PDF 时也照常导出。
    """
    candidates = [name for name in sheet_names if has_backup_marker(name)]
    originals = {_comparable_name(name) for name in sheet_names if name not in candidates}
    return [
        name for name in candidates
        if (base := _original_name_of(name)) and base in originals
    ]


def is_uncertain_name(name: str) -> bool:
    """判断文件名或工作表名是否带有「❓ 待定」前缀"""
    return name.startswith(UNCERTAIN_EMOJI)


def strip_uncertain_prefix(name: str) -> str:
    """去掉文件名或工作表名前面的「❓」前缀"""
    cleaned = name
    while cleaned.startswith(UNCERTAIN_EMOJI):
        cleaned = cleaned.removeprefix(UNCERTAIN_EMOJI).lstrip()
    return cleaned or name


def add_uncertain_prefix(name: str) -> str:
    """为文件名或工作表名加上「❓」前缀（避免重复叠加）"""
    return f"{UNCERTAIN_EMOJI}{strip_uncertain_prefix(name)}"


def unique_sheet_name(desired: str, existing_names: set[str]) -> str:
    """生成符合 Excel 31 字符限制且不与现有工作表重名的名称"""
    trimmed = desired[:EXCEL_SHEET_MAX_LEN] or "Sheet"
    if trimmed not in existing_names:
        return trimmed
    counter = 2
    while True:
        suffix = f"_{counter}"
        candidate = f"{trimmed[:EXCEL_SHEET_MAX_LEN - len(suffix)]}{suffix}"
        if candidate not in existing_names:
            return candidate
        counter += 1


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
