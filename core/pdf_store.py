"""导出的 PDF 统一放在所选文件夹下的「PDFS」文件夹里（用户 2026-11-02 定），不再和 Excel 混在一起。

PDFS 里保留和 Excel 一样的子文件夹结构：根目录/01_画面/spec.xlsx 的 PDF 放在 根目录/PDFS/01_画面/spec_表名.pdf，
看 PDF 时按文件夹就能找到是哪份 Excel 的，不同子文件夹里同名的 Excel 也不会撞名。

PDFS 里还有一份「PDF导出记录.json」：每份 Excel（按相对路径）导出时的哈希值，以及导出的每份 PDF
（相对 PDFS 的路径、表名、PDF 的大小和修改时间）。下次导出时 Excel 哈希没变、这些 PDF 也都还在且没被换过，
就直接沿用，不再打开 Excel；哈希变了才删掉这份 Excel 的旧 PDF 重新导出。
记录跟着文件夹走（不放在工具的数据库里），更新工具、换电脑、整个文件夹挪地方都照样认得。
"""
from __future__ import annotations

import json
import os
import threading
from pathlib import Path
from typing import Any

from core import debug_log
from core.common import EXCEL_EXTENSIONS, sanitize_filename, strip_review_prefix, workbook_rel_name

DIR_NAME = "PDFS"
MANIFEST_NAME = "PDF导出记录.json"
VERSION = 1
_LOCK = threading.RLock()


def pdf_root(root: Path | str) -> Path:
    return Path(root) / DIR_NAME


def pdf_dir(wb_path: Path, root: Path | str | None) -> Path:
    """这份 Excel 的 PDF 放在哪：根目录/PDFS/<Excel 所在的子文件夹>。Excel 不在根目录里（不该发生）时放它旁边的 PDFS"""
    if root:
        try:
            rel = wb_path.resolve().parent.relative_to(Path(root).resolve())
            return pdf_root(Path(root).resolve()) / rel
        except ValueError:
            pass
    return wb_path.parent / DIR_NAME


def file_key(wb_path: Path, root: Path | str) -> str:
    """记录里认 Excel 用的名字：相对根目录的路径（去掉「★」临时标记，加不加 ★ 都算同一份）"""
    return workbook_rel_name(wb_path, root)


def export_root_of(pdf: Path) -> Path | None:
    """PDF 在哪个 PDFS 文件夹里：返回那个 PDFS 的上一级（导出时选的根目录）；不在 PDFS 里返回 None"""
    for parent in pdf.parents:
        if parent.name == DIR_NAME:
            return parent.parent
    return None


# ---------- 读写记录 ----------

def _manifest_path(root: Path | str) -> Path:
    return pdf_root(root) / MANIFEST_NAME


def load(root: Path | str) -> dict[str, Any]:
    path = _manifest_path(root)
    with _LOCK:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return {"v": VERSION, "files": {}}
        except Exception as err:  # 记录坏了：当没有记录，全部重新导出
            debug_log.event("⚠️ PDF 导出记录读不了，当作没有", 文件=path, 原因=str(err))
            return {"v": VERSION, "files": {}}
    if not isinstance(data, dict) or not isinstance(data.get("files"), dict):
        return {"v": VERSION, "files": {}}
    return data


def _save(root: Path | str, data: dict[str, Any]) -> None:
    path = _manifest_path(root)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + ".tmp")
    temp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(temp, path)


def _stat(pdf: Path) -> tuple[int, int] | None:
    try:
        st = pdf.stat()
    except OSError:
        return None
    return st.st_size, st.st_mtime_ns


def _entry_pdfs(root: Path | str, entry: dict) -> list[tuple[Path, str, dict]]:
    base = pdf_root(root)
    return [(base / item["pdf"], str(item.get("sheet", "")), item)
            for item in entry.get("pdfs", []) if isinstance(item, dict) and item.get("pdf")]


def reusable(root: Path | str, wb_path: Path, digest: str, *, need_fit: bool = False) -> list[tuple[Path, str]] | None:
    """Excel 哈希和上次导出时一样、那次导出的 PDF 都还在且没被别的东西换过：返回 [(PDF, 表名)]，否则 None。
    need_fit：还要那次导出时每张表都已经设成「宽一页」（调整页面，2026-11-07），旧版本导出的记录不算"""
    with _LOCK:
        entry = load(root)["files"].get(file_key(wb_path, root))
    if not entry or entry.get("hash") != digest:
        return None
    if need_fit and not entry.get("fit"):
        return None
    found: list[tuple[Path, str]] = []
    for pdf, sheet, item in _entry_pdfs(root, entry):
        if _stat(pdf) != (item.get("size"), item.get("mtime_ns")):
            return None
        if _item_hash(entry, item) != digest:
            return None  # 只重新导出过别的表、这张表的 PDF 还是 Excel 改之前的
        if need_fit and "manual" not in (item.get("layout") or {}):
            # 2026-11-08 以前导出的记录没记手动分页：重新导出一次，调整页面才有「清手动分页」按钮
            return None
        found.append((pdf, sheet))
    return found


def _item_hash(entry: dict, item: dict) -> str:
    """这份 PDF 导出时 Excel 的哈希（2026-11-08 起每张表单独记；以前的记录只有整份的）"""
    return str(item.get("hash") or entry.get("hash") or "")


def stale_pdfs(root: Path | str, fingerprint: Any) -> dict[str, dict]:
    """导出以后 Excel 又改过（程序改的、人点路径改完保存的）的 PDF：{PDF 绝对路径: {wb_path, sheet}}。
    fingerprint(Excel 路径) 算现在的哈希；Excel 找不到、读不了的不算"""
    with _LOCK:
        files = load(root)["files"]
    found: dict[str, dict] = {}
    for key, entry in files.items():
        if not isinstance(entry, dict):
            continue
        wb = _find_excel(Path(root), key, entry.get("excel", ""))
        if wb is None:
            continue
        try:
            digest = fingerprint(wb)
        except OSError:
            continue
        for pdf, sheet, item in _entry_pdfs(root, entry):
            if _item_hash(entry, item) != digest and pdf.is_file():
                found[str(pdf)] = {"wb_path": str(wb), "sheet": sheet}
    return found


def sheet_pdfs(root: Path | str, wb_path: Path) -> dict[str, Path]:
    """记录里这份 Excel 每张表的 PDF：{表名: PDF}（只重新生成一张表时，按这个找原来那份 PDF 覆盖）"""
    with _LOCK:
        entry = load(root)["files"].get(file_key(wb_path, root))
    if not isinstance(entry, dict):
        return {}
    return {sheet: pdf for pdf, sheet, _ in _entry_pdfs(root, entry) if sheet}


def record(root: Path | str, wb_path: Path, digest: str, exported: list[tuple[Path, str]],
           notes: dict[str, list[dict]] | None = None, layouts: dict[str, dict] | None = None,
           fit: bool = False, *, merge: bool = False, before: str | None = None) -> None:
    """notes：PDF -> 内容很少的页的提示（page_fill），沿用 PDF 时提示也跟着沿用。
    layouts：PDF -> 这张表的打印设置（方向、页高、末页印了多少），调整页面时网页按这个显示和分类。
    fit：每张表都已经设成「宽一页」了。
    merge：只重新导出了其中几张表（调整页面「重新生成」/「统一修改」，用户 2026-11-08 定按表来）：
    别的表的 PDF 记录照旧留着，只换这几张，哈希记成现在的。
    before：这次导出开始前 Excel 的哈希。别的表的 PDF 那时候是最新的，这次只是程序自己存了「宽一页」等设置，
    跟着记成现在的哈希；那时候就已经旧了的（Excel 先被改过）照旧记旧哈希，打开调整页面时会重新导出"""
    base = pdf_root(root)
    pdfs = []
    for pdf, sheet in exported:
        stat = _stat(pdf)
        if stat is None:
            continue
        item = {"pdf": pdf.relative_to(base).as_posix(), "sheet": sheet, "size": stat[0], "mtime_ns": stat[1],
                "hash": digest}
        if notes and notes.get(str(pdf)):
            item["notes"] = notes[str(pdf)]
        if layouts and layouts.get(str(pdf)):
            item["layout"] = layouts[str(pdf)]
        pdfs.append(item)
    with _LOCK:
        data = load(root)
        data["v"] = VERSION
        key = file_key(wb_path, root)
        old = data["files"].get(key) if merge else None
        if isinstance(old, dict):
            mine = {item["pdf"] for item in pdfs}
            kept = [dict(item) for item in old.get("pdfs", []) if isinstance(item, dict) and item.get("pdf") not in mine]
            for item in kept:
                was = _item_hash(old, item)
                item["hash"] = digest if before is not None and was == before else was
            order = [item.get("pdf") for item in old.get("pdfs", []) if isinstance(item, dict)]
            pdfs = sorted(kept + pdfs, key=lambda item: order.index(item["pdf"]) if item["pdf"] in order else len(order))
            fit = fit and bool(old.get("fit"))
        data["files"][key] = {
            "excel": wb_path.resolve().relative_to(Path(root).resolve()).as_posix()
            if _inside(wb_path, root) else wb_path.name,
            "hash": digest, "pdfs": pdfs,
        }
        if fit:
            data["files"][file_key(wb_path, root)]["fit"] = True
        _save(root, data)


def has_record(root: Path | str, wb_path: Path) -> bool:
    """调整页面给这份 Excel 导出过 PDF（记录里有它）"""
    with _LOCK:
        entry = load(root)["files"].get(file_key(wb_path, root))
    return isinstance(entry, dict) and bool(entry.get("pdfs"))


def follow_save(root: Path | str, wb_path: Path, before: str, after: str) -> bool:
    """程序自己存了一下 Excel、打印出来的样子没变（标色、去标色、存分页预览）：导出记录里按改之前的哈希记着的
    PDF 跟着记成现在的哈希，回到调整页面时不会当成「Excel 改过」整份重新导出（用户 2026-11-09 把切字并进调整页面）。
    改之前就已经旧了的 PDF 照旧是旧的。返回有没有跟进"""
    if not before or not after or before == after:
        return False
    with _LOCK:
        data = load(root)
        entry = data["files"].get(file_key(wb_path, root))
        if not isinstance(entry, dict):
            return False
        moved = False
        for item in entry.get("pdfs", []):
            if isinstance(item, dict) and _item_hash(entry, item) == before:
                item["hash"] = after
                moved = True
        if entry.get("hash") == before:
            entry["hash"] = after
            moved = True
        if moved:
            _save(root, data)
    if moved:
        debug_log.event("♻️ PDF 记录跟上程序自己的保存", 文件=file_key(wb_path, root), 原哈希=before[:12], 新哈希=after[:12])
    return moved


def layout_map(root: Path | str) -> dict[str, dict]:
    """记录里每份 PDF（绝对路径）导出时这张表的打印设置；PDF 被换过的不算"""
    with _LOCK:
        files = load(root)["files"]
    found: dict[str, dict] = {}
    for entry in files.values():
        if not isinstance(entry, dict):
            continue
        for pdf, _, item in _entry_pdfs(root, entry):
            layout = item.get("layout")
            if isinstance(layout, dict) and _stat(pdf) == (item.get("size"), item.get("mtime_ns")):
                found[str(pdf)] = layout
    return found


def notes_map(root: Path | str) -> dict[str, list[dict]]:
    """记录里每份 PDF（绝对路径）的内容很少的页的提示；PDF 被换过（大小、修改时间不对）的不算"""
    with _LOCK:
        files = load(root)["files"]
    found: dict[str, list[dict]] = {}
    for entry in files.values():
        if not isinstance(entry, dict):
            continue
        for pdf, _, item in _entry_pdfs(root, entry):
            notes = item.get("notes")
            if isinstance(notes, list) and notes and _stat(pdf) == (item.get("size"), item.get("mtime_ns")):
                found[str(pdf)] = notes
    return found


def _inside(path: Path, root: Path | str) -> bool:
    try:
        path.resolve().relative_to(Path(root).resolve())
        return True
    except ValueError:
        return False


def owned_by_others(root: Path | str, wb_path: Path) -> set[Path]:
    """别的 Excel 记录下的 PDF：给这份 Excel 起名时避开，别把它们沿用的 PDF 盖掉"""
    key = file_key(wb_path, root)
    with _LOCK:
        files = load(root)["files"]
    return {pdf for k, entry in files.items() if k != key for pdf, _, _ in _entry_pdfs(root, entry)}


def clear_workbook(root: Path | str, wb_path: Path) -> list[Path]:
    """这份 Excel 要重新导出：删掉它记录下的旧 PDF，以及 PDFS 里名字是「Excel 文件名_….pdf」、
    又不属于别的 Excel 的（扫描时导出的、表改名或删掉后留下的）。删不掉的抛 OSError（多半是 PDF 正开着）"""
    key = file_key(wb_path, root)
    with _LOCK:
        entry = load(root)["files"].get(key)
    targets = [pdf for pdf, _, _ in _entry_pdfs(root, entry or {})]
    folder = pdf_dir(wb_path, root)
    if folder.is_dir():
        others = owned_by_others(root, wb_path)
        stem = sanitize_filename(strip_review_prefix(wb_path.stem))
        siblings = _sibling_stems(wb_path)
        for pdf in folder.iterdir():
            if (pdf.is_file() and pdf.suffix.lower() == ".pdf" and pdf not in others and pdf not in targets
                    and _belongs_to(pdf.stem, stem, siblings)):
                targets.append(pdf)
    removed = []
    try:
        for pdf in targets:
            if pdf.exists():
                pdf.unlink()
                removed.append(pdf)
    finally:
        if entry is not None:
            with _LOCK:
                data = load(root)
                data["files"].pop(key, None)
                _save(root, data)
    return removed


def _sibling_stems(wb_path: Path) -> set[str]:
    try:
        return {sanitize_filename(strip_review_prefix(p.stem)) for p in wb_path.parent.iterdir()
                if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS | {".xlsb"} and not p.name.startswith("~$")}
    except OSError:
        return set()


def _belongs_to(pdf_stem: str, stem: str, siblings: set[str]) -> bool:
    """「a_b_表1.pdf」同一文件夹里有 a.xlsx 和 a_b.xlsx 时算 a_b 的：按最长的 Excel 文件名认"""
    if not pdf_stem.startswith(stem + "_"):
        return False
    return not any(len(other) > len(stem) and pdf_stem.startswith(other + "_") for other in siblings)


def prune(root: Path | str, workbooks: list[Path]) -> list[Path]:
    """导出开始时：这次文件夹里已经没有的 Excel（删掉、挪走、改名了），它们的 PDF 和记录都去掉；
    最后把 PDFS 里空了的子文件夹删掉"""
    keep = {file_key(wb, root) for wb in workbooks}
    removed: list[Path] = []
    with _LOCK:
        data = load(root)
        gone = [key for key in data["files"] if key not in keep]
        for key in gone:
            for pdf, _, _ in _entry_pdfs(root, data["files"][key]):
                try:
                    if pdf.exists():
                        pdf.unlink()
                        removed.append(pdf)
                except OSError as err:
                    debug_log.event("⚠️ 删不掉已不在的 Excel 的 PDF", PDF=pdf, 原因=str(err))
            data["files"].pop(key)
        if gone:
            _save(root, data)
    remove_empty_dirs(root)
    return removed


def remove_empty_dirs(root: Path | str) -> None:
    base = pdf_root(root)
    if not base.is_dir():
        return
    for folder in sorted((p for p in base.rglob("*") if p.is_dir()), key=lambda p: len(p.parts), reverse=True):
        try:
            folder.rmdir()  # 只删空的
        except OSError:
            pass


def owner(pdf: Path) -> dict | None:
    """记录里这份 PDF 是哪份 Excel 的哪张表：{wb_path, sheet}；Excel 不在了返回 None"""
    root = export_root_of(pdf)
    if root is None:
        return None
    try:
        rel = pdf.resolve().relative_to(pdf_root(root).resolve()).as_posix()
    except ValueError:
        return None
    with _LOCK:
        files = load(root)["files"]
    for key, entry in files.items():
        for item in entry.get("pdfs", []):
            if isinstance(item, dict) and item.get("pdf") == rel:
                wb = _find_excel(root, key, entry.get("excel", ""))
                return {"wb_path": str(wb), "sheet": str(item.get("sheet", ""))} if wb else None
    return None


def _find_excel(root: Path, key: str, excel: str) -> Path | None:
    for candidate in (root / excel if excel else None, root / key):
        if candidate is not None and candidate.is_file():
            return candidate
    folder = (root / key).parent
    name = Path(key).name
    try:
        for path in folder.iterdir():
            if path.is_file() and strip_review_prefix(path.name) == name:
                return path
    except OSError:
        pass
    return None


def excel_folder_for(pdf: Path) -> Path:
    """PDFS 里的 PDF 对应的 Excel 在哪个文件夹（按文件名猜 Excel 时用）；不在 PDFS 里就是 PDF 自己的文件夹"""
    root = export_root_of(pdf)
    if root is None:
        return pdf.parent
    try:
        return root / pdf.parent.resolve().relative_to(pdf_root(root).resolve())
    except ValueError:
        return pdf.parent


def unchanged(root: Path | str, wb_path: Path, digest: str) -> bool:
    return reusable(root, wb_path, digest) is not None


__all__ = ["DIR_NAME", "MANIFEST_NAME", "pdf_root", "pdf_dir", "export_root_of", "reusable", "record",
           "clear_workbook", "prune", "owner", "layout_map", "notes_map", "owned_by_others", "excel_folder_for", "unchanged",
           "stale_pdfs", "sheet_pdfs", "has_record", "follow_save"]
