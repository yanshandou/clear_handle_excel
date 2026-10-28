"""颜色、引用与隐藏状态恢复共用的备份文件、工作表配对。"""
from pathlib import Path
from typing import Any, Callable
import unicodedata

from core.common import (
    _original_name_of, strip_new_prefix, strip_uncertain_prefix, workbook_rel_name,
)


def norm(name: str) -> str:
    return unicodedata.normalize("NFKC", name).strip().casefold()


def pair_files(old_root: Path, old_paths: list[Path], names: list[str],
               current_root: Path | None = None) -> dict[str, Path | None]:
    by_rel: dict[str, list[Path]] = {}
    by_name: dict[str, list[Path]] = {}
    for path in old_paths:
        by_rel.setdefault(norm(workbook_rel_name(path, old_root)), []).append(path)
        by_name.setdefault(norm(workbook_rel_name(path)), []).append(path)
    counts: dict[str, int] = {}
    for name in names:
        key = norm(Path(name).name)
        counts[key] = counts.get(key, 0) + 1
    pairs = {}
    for name in names:
        # 当前选到 001，备份选到含 001/002/003 的总目录，也按 001 下的相对路径找。
        found = []
        if current_root and norm(current_root.name) != norm(old_root.name):
            found = by_rel.get(norm(f"{current_root.name}/{name}"), [])
        if not found:
            found = by_rel.get(norm(name), [])
        if not found and counts[norm(Path(name).name)] == 1:
            found = by_name.get(norm(Path(name).name), [])
        pairs[name] = found[0] if len(found) == 1 else None
    used: dict[Path, int] = {}
    for path in pairs.values():
        if path is not None:
            used[path] = used.get(path, 0) + 1
    return {name: path if path is None or used[path] == 1 else None for name, path in pairs.items()}


def pair_sheets(current: list[Any], old: list[Any], *, title: Callable,
                code: Callable, words: Callable) -> dict[str, tuple[Any, str]]:
    """先表名、可识别的新旧标记、代码名，最后才用内容；歧义留给人工。"""
    found: dict[str, tuple[Any, str]] = {}
    left, rest = list(old), list(current)

    def assign(key: Callable, how: str) -> None:
        for sheet in list(rest):
            value = key(sheet)
            same = [other for other in left if value and key(other) == value]
            if len(same) == 1 and sum(key(other) == value for other in rest) == 1:
                found[title(sheet)] = (same[0], how)
                left.remove(same[0])
                rest.remove(sheet)

    assign(lambda sheet: norm(title(sheet)), "name")
    assign(lambda sheet: _original_name_of(strip_new_prefix(strip_uncertain_prefix(title(sheet)))), "guess")
    assign(lambda sheet: norm(code(sheet) or ""), "code")
    if rest and left:
        old_words = {id(sheet): words(sheet) for sheet in left}
        for sheet in rest:
            current_words = words(sheet)
            scores = sorted(
                ((len(current_words & old_words[id(other)]) / len(current_words | old_words[id(other)]), i)
                 for i, other in enumerate(left) if current_words | old_words[id(other)]), reverse=True)
            if scores and scores[0][0] >= 0.6 and (len(scores) == 1 or scores[1][0] < scores[0][0]):
                match = left.pop(scores[0][1])
                found[title(sheet)] = (match, "content")
    return found
