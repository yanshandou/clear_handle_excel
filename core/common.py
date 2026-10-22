import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

# 日期：20220929，或用同一种分隔符隔开的 2018-02-26 / 2018.02.26 / 2018/02/26（括号由比对原表时忽略）
DATE_PATTERN = re.compile(r"(?<!\d)((?:19|20)\d{2})([-/.]?)(\d{2})\2(\d{2})(?!\d)")
# 独立出现的英文备份标记（前后不是字母），如 _bak、Sheet1-BAK、一覧_old；Bakery、Threshold 不算
BACKUP_WORD_PATTERN = re.compile(r"(?<![a-z])(?:bak|backup|old|copy)(?![a-z])", re.IGNORECASE)
# 日文/中文备份标记；「旧」不含「新旧」「復旧」（如 新旧対照表、障害復旧手順 是正式资料）
BRACKET_GROUP_PATTERN = re.compile(r"\([^()]*\)|\[[^\[\]]*\]|【[^【】]*】|「[^「」]*」|『[^『』]*』")
NEW_PREFIX = "新"
OLD_PREFIX = "旧"
BACKUP_TEXT_PATTERN = re.compile(r"バックアップ|备份|コピー|副本|(?<![新復])旧")
# 复制工作表产生的「(2)」「(3)」后缀（全角括号统一为半角后判断，有无空格均可），
# 后面还可以再接日期，如「様式１（基本情報）(2)_20221005」；仍要求同工作簿存在原表才算备份
EXCEL_COPY_SUFFIX_PATTERN = re.compile(r"\s*\(\d+\)[\s_\-・.]*$")
# 名称末尾不带年份的 4 位月日（MMDD），如「請求書1)0513」；只作备份标记，不参与「多个日期版本」的新旧比较
MMDD_SUFFIX_PATTERN = re.compile(r"(?<!\d)(\d{2})(\d{2})[\s_\-・.]*$")
# 工具自动生成的工作表（如 自動生成_引数、自動生成_JSON），没有对应原表，精简 Sheet 时默认删除
AUTO_GENERATED_PREFIXES = ("自動生成", "自动生成")
# 去掉备份标记、找原表时额外去掉的写法（「Sheet1のコピー」的「の」、「Copy of 一覧」的「of」）
COPY_PHRASE_PATTERN = re.compile(r"copy\s+of|の?コピー|的?副本", re.IGNORECASE)
# 版本号（如 _v2.4、ver1.0），比对原表时忽略
VERSION_PATTERN = re.compile(r"(?<![a-z])v(?:er)?\.?\d+(?:\.\d+)*(?![a-z])", re.IGNORECASE)
# 比对原表时忽略的空白、分隔符、括号与待定标记
UNCERTAIN_EMOJI = "❓"
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


def _date_of(match: re.Match[str]) -> str | None:
    """DATE_PATTERN 匹配到的有效日期（YYYYMMDD），无效日期返回 None"""
    text = match.group(1) + match.group(3) + match.group(4)
    return text if _is_valid_date(text) else None


def _latest_date_in(name: str) -> str | None:
    dates = [d for m in DATE_PATTERN.finditer(unicodedata.normalize("NFKC", name)) if (d := _date_of(m))]
    return max(dates) if dates else None


def _has_mmdd_suffix(text: str) -> bool:
    """名称末尾是否为有效的 4 位月日（按闰年判断，0229 也算）"""
    m = MMDD_SUFFIX_PATTERN.search(text)
    return bool(m) and _is_valid_date(f"2000{m.group(1)}{m.group(2)}")


def has_backup_marker(sheet_name: str) -> bool:
    """名称是否带备份标记：有效日期（YYYYMMDD 或 YYYY-MM-DD 等，或末尾的 4 位月日 MMDD）、bak/backup/old/copy、
    バックアップ/备份/旧/コピー/副本，或复制产生的「(2)」后缀。

    只是候选条件，是否真的是备份表见 find_backup_sheets。
    """
    return bool(
        BACKUP_WORD_PATTERN.search(sheet_name)
        or BACKUP_TEXT_PATTERN.search(sheet_name)
        or EXCEL_COPY_SUFFIX_PATTERN.search(unicodedata.normalize("NFKC", sheet_name))
        or _latest_date_in(sheet_name) is not None
        or _has_mmdd_suffix(unicodedata.normalize("NFKC", sheet_name))
    )


def _comparable_name(name: str) -> str:
    """用于比对原表的名称：统一全角半角与大小写，忽略版本号、空白与分隔符"""
    text = unicodedata.normalize("NFKC", name).lower()
    text = VERSION_PATTERN.sub("", text)
    return NAME_SEPARATOR_PATTERN.sub("", text)


def _original_name_of(backup_name: str) -> str:
    """去掉备份标记后的原表名称，如 基本情報_20241003 → 基本情報"""
    text = unicodedata.normalize("NFKC", backup_name)
    text = DATE_PATTERN.sub(lambda m: "" if _date_of(m) else m.group(0), text)
    text = EXCEL_COPY_SUFFIX_PATTERN.sub("", text)
    text = COPY_PHRASE_PATTERN.sub("", text)
    text = BACKUP_WORD_PATTERN.sub("", text)
    text = BACKUP_TEXT_PATTERN.sub("", text)
    if _has_mmdd_suffix(text):
        text = MMDD_SUFFIX_PATTERN.sub("", text)
    return _comparable_name(text)


def find_backup_sheets(sheet_names: list[str]) -> list[str]:
    """找出工作簿中的备份表：名称带备份标记，且去掉标记后能在同一工作簿找到对应的原表。

    找不到原表时（如只有 基本情報_20241003 而没有 基本情報），带标记的表就是当前在用的版本，
    保留不删，导出 PDF 时也照常导出。若同一名称有多个日期版本、都找不到原表，
    只保留日期最新的一个（如 20170113要介護度別状況調査表），更早的日期版本视为备份。
    """
    candidates = [name for name in sheet_names if has_backup_marker(name)]
    originals = {_comparable_name(name) for name in sheet_names if name not in candidates}
    backups: set[str] = set()
    dated_without_original: dict[str, list[tuple[str, str]]] = {}
    for name in candidates:
        base = _original_name_of(name)
        if not base:
            continue
        if base in originals:
            backups.add(name)
        elif (date := _latest_date_in(name)) is not None:
            dated_without_original.setdefault(base, []).append((date, name))
    for versions in dated_without_original.values():
        newest = max(date for date, _ in versions)
        backups.update(name for date, name in versions if date < newest)
    return [name for name in candidates if name in backups]


def is_auto_generated_sheet(sheet_name: str) -> bool:
    """名称以「自動生成」开头的工作表（忽略 ❓ 待定前缀与全角半角差异）"""
    text = unicodedata.normalize("NFKC", sheet_name).lstrip(UNCERTAIN_EMOJI).lstrip()
    return text.startswith(AUTO_GENERATED_PREFIXES)


def _new_old_key(sheet_name: str) -> tuple[str, str] | None:
    """「新」「旧」开头的表名拆成 (新/旧, 其余部分的比对名)，其余部分为空时返回 None"""
    text = unicodedata.normalize("NFKC", strip_uncertain_prefix(sheet_name)).lstrip()
    if text[:1] not in (NEW_PREFIX, OLD_PREFIX):
        return None
    rest = _comparable_name(text[1:])
    return (text[0], rest) if rest else None


def find_new_old_pairs(sheet_names: list[str]) -> dict[str, list[str]]:
    """成对出现的新旧版本：{新表名: [对应的旧表名, ...]}

    如 新レイアウト 与 旧レイアウト、新レイアウト(介護給付費請求書1) 与 旧レイアウト (介護給付費請求書1)：
    去掉开头的「新」「旧」后名称相同（忽略全角半角、空白与括号）。只有「新」或只有「旧」的不算。
    """
    news: dict[str, list[str]] = {}
    olds: dict[str, list[str]] = {}
    for name in sheet_names:
        key = _new_old_key(name)
        if key is not None:
            (news if key[0] == NEW_PREFIX else olds).setdefault(key[1], []).append(name)
    return {new: olds[rest] for rest, names in news.items() if rest in olds for new in names}


def strip_new_prefix(sheet_name: str) -> str:
    """去掉表名开头的「新」：新レイアウト(介護給付費請求書1) → レイアウト(介護給付費請求書1)"""
    if not sheet_name.startswith(NEW_PREFIX):
        return sheet_name
    return sheet_name[len(NEW_PREFIX):].lstrip() or sheet_name


def find_clean_targets(sheet_names: list[str]) -> list[str]:
    """精简 Sheet 默认删除的工作表：有原表的备份表、有对应「新」版本的「旧」表，
    加上「自動生成」开头的工作表（保持原顺序）"""
    targets = set(find_backup_sheets(sheet_names))
    for olds in find_new_old_pairs(sheet_names).values():
        targets.update(olds)
    return [name for name in sheet_names if name in targets or is_auto_generated_sheet(name)]


def find_contained_sheets(sheet_names: list[str]) -> list[str]:
    """名称完整包含另一个工作表名称、但多出的部分不是可识别备份标记的工作表（如 一覧 与 一覧詳細），
    可能是备份也可能是正式的明细表，精简 Sheet 时默认标为待定，交给人工确认，不默认删除。

    已默认删除、已带 ❓、「自動生成」开头的表不算；被包含的名称至少 2 个字符，
    且不能把数字从中间切开（Sheet1 与 Sheet12 是不同的表）；多出的部分整段在括号里的也不算；
    自己有备份表的原表也不算（它就是当前在用的版本）。
    """
    targets = set(find_clean_targets(sheet_names))
    targets.update(find_new_old_pairs(sheet_names))  # 新旧成对的「新」表默认保留（删旧表后去掉「新」字）
    # 有备份表的原表就是当前在用的版本（如 精神科基本療養費 与 精神科基本療養費20260402），默认保留
    has_backup = {_original_name_of(name) for name in find_backup_sheets(sheet_names)}
    references = {
        c for name in sheet_names
        if name not in targets and len(c := _comparable_name(name)) >= 2
    }
    contained: list[str] = []
    for name in sheet_names:
        text = _comparable_name(name)
        if name in targets or is_uncertain_name(name) or is_auto_generated_sheet(name) or text in has_backup:
            continue
        # 多出的部分整段在括号里（如 レイアウト 与 レイアウト(介護給付費請求書1)）是用途或页的补充说明，不算
        bare = _comparable_name(BRACKET_GROUP_PATTERN.sub("", unicodedata.normalize("NFKC", name)))
        for ref in references:
            if ref == text or ref == bare:
                continue
            start = text.find(ref)
            while start != -1:
                end = start + len(ref)
                cuts_number = (start > 0 and text[start - 1].isdigit() and ref[0].isdigit()) or (
                    end < len(text) and text[end].isdigit() and ref[-1].isdigit()
                )
                if not cuts_number:
                    break
                start = text.find(ref, start + 1)
            if start != -1:
                contained.append(name)
                break
    return contained


def is_uncertain_name(name: str) -> bool:
    """判断文件名或工作表名是否带有「❓ 待定」前缀"""
    return name.startswith(UNCERTAIN_EMOJI)


def strip_uncertain_prefix(name: str) -> str:
    """去掉文件名或工作表名前面的「❓」前缀"""
    cleaned = name
    while cleaned.startswith(UNCERTAIN_EMOJI):
        cleaned = cleaned.removeprefix(UNCERTAIN_EMOJI).lstrip()
    return cleaned or name


def workbook_rel_name(wb_path: Path, root_folder: Path | str | None = None) -> str:
    """返回工作簿相对于目标根目录的显示名称（顶层文件返回文件名，子文件夹文件返回 sub/file.xlsx）"""
    if root_folder:
        try:
            return wb_path.resolve().relative_to(Path(root_folder).resolve()).as_posix()
        except ValueError:
            pass
    return wb_path.name


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
