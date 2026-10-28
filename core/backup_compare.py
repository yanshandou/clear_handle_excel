"""颜色、引用与隐藏状态恢复共用的备份文件、工作表配对。"""
from pathlib import Path
from typing import Any, Callable
import unicodedata

from core.common import (
    _original_name_of, find_new_prefix_sheets, strip_new_prefix, strip_uncertain_prefix, workbook_rel_name,
)
from core.sheet_history import Projection


def norm(name: str) -> str:
    return unicodedata.normalize("NFKC", name).strip().casefold()


def pair_files(old_root: Path, old_paths: list[Path], names: list[str],
               current_root: Path | None = None, *, inferred: set[str] | None = None) -> dict[str, Path | None]:
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
    exact_candidates: set[Path] = set()
    ambiguous: set[str] = set()
    for name in names:
        # 当前选到 001，备份选到含 001/002/003 的总目录，也按 001 下的相对路径找。
        found = []
        if current_root and norm(current_root.name) != norm(old_root.name):
            found = by_rel.get(norm(f"{current_root.name}/{name}"), [])
        if not found:
            found = by_rel.get(norm(name), [])
        if not found and counts[norm(Path(name).name)] == 1:
            found = by_name.get(norm(Path(name).name), [])
        exact_candidates.update(found)
        if len(found) > 1:
            ambiguous.add(name)
        pairs[name] = found[0] if len(found) == 1 else None
    used: dict[Path, int] = {}
    for path in pairs.values():
        if path is not None:
            used[path] = used.get(path, 0) + 1
    ambiguous.update(name for name, path in pairs.items() if path is not None and used[path] > 1)
    pairs = {name: path if path is None or used[path] == 1 else None for name, path in pairs.items()}
    if inferred is None:
        return pairs  # 调用方明确处理推断标识后，才启用不确定的文件名猜测。

    # 文件名猜测不是清理历史证据，调用方必须将 inferred 中的文件交人工确认，不能自动写回。
    import re
    from datetime import date

    suffix = re.compile(r"(?:[\s_.-]+(?:bak|backup|bk|old|copy|备份|備份|副本|コピー|バックアップ|旧)|\s*\((?:bak|backup|bk|old|copy|备份|備份|副本|コピー|バックアップ|旧|[2-9]\d*)\))$", re.IGNORECASE)
    dated = re.compile(r"(?:[\s_.-]+|\()((?:19|20)\d{2})[-.]?(\d{2})[-.]?(\d{2})\)?$")
    extensions = {".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm"}

    def stem(name: str) -> str:
        return norm(Path(strip_uncertain_prefix(Path(name).name)).stem)

    def base(name: str) -> str:
        value = stem(name)
        while True:
            trimmed = suffix.sub("", value)
            match = dated.search(trimmed)
            if match:
                try:
                    date(*(int(part) for part in match.groups()))
                except ValueError:
                    pass
                else:
                    trimmed = trimmed[:match.start()]
            trimmed = trimmed.rstrip(" _.-")
            if trimmed == value:
                return value
            value = trimmed

    available = [(path, workbook_rel_name(path, old_root)) for path in old_paths if path not in exact_candidates]
    choices: dict[str, list[Path]] = {}
    for name in names:
        if pairs[name] is not None or name in ambiguous or Path(name).suffix.lower() not in extensions:
            continue
        name_base = base(name)
        if not name_base:
            continue
        rel = Path(name)
        ranked: list[tuple[tuple[int, int], Path]] = []
        for path, old_rel in available:
            if path.suffix.lower() not in extensions or base(old_rel) != name_base:
                continue
            old_parent = norm(Path(old_rel).parent.as_posix())
            rel_parent = norm(rel.parent.as_posix())
            nested_parent = norm((Path(current_root.name) / rel.parent).as_posix()) if current_root else ""
            directory_rank = (0 if current_root and norm(current_root.name) != norm(old_root.name)
                              and old_parent == nested_parent else 1 if old_parent == rel_parent else 2)
            ranked.append(((directory_rank, int(stem(name) != stem(old_rel))), path))
        if ranked:
            best = min(rank for rank, _ in ranked)
            choices[name] = [path for rank, path in ranked if rank == best]
    claims: dict[Path, int] = {}
    for candidates in choices.values():
        for path in candidates:
            claims[path] = claims.get(path, 0) + 1
    for name, candidates in choices.items():
        if len(candidates) == 1 and claims[candidates[0]] == 1:
            pairs[name] = candidates[0]
            inferred.add(name)
    return pairs


def pair_sheets(current: list[Any], old: list[Any], *, title: Callable,
                code: Callable, words: Callable, history: Projection | None = None) -> dict[str, tuple[Any, str]]:
    """先按操作历史追溯身份，再按表名、代码名、内容；歧义留给人工。"""
    found: dict[str, tuple[Any, str]] = {}
    left, rest = list(old), list(current)

    origins = history.origins if history else {}
    for sheet in list(rest):
        name = title(sheet)
        if name not in origins:
            continue
        source = origins[name]
        same = [other for other in left if source is not None and norm(title(other)) == norm(source)]
        if len(same) == 1 and list(origins.values()).count(source) == 1:
            found[name] = (same[0], "history")
            left.remove(same[0])
        rest.remove(sheet)  # 有历史但不能唯一追溯时，不回退到同名的另一张表。

    # 旧版清理可能删除同名旧表并把「新X」改为「X」。先检查这类同名冲突，
    # 单靠命名规则无法证明历史：必须有唯一代码名或明显更相似的内容。
    old_new_names = set(find_new_prefix_sheets([title(sheet) for sheet in old]))
    present = {norm(title(sheet)) for sheet in current}
    for sheet in list(rest):
        name = norm(title(sheet))
        candidates = [other for other in left if norm(title(other)) == name or
                      (title(other) in old_new_names and norm(strip_new_prefix(title(other))) == name
                       and norm(title(other)) not in present)]
        alternatives = [other for other in candidates if norm(title(other)) != name]
        if not alternatives or not any(norm(title(other)) == name for other in candidates):
            continue
        code_value = norm(code(sheet) or "")
        coded = [other for other in candidates if code_value and norm(code(other) or "") == code_value]
        code_match = coded[0] if len(coded) == 1 and sum(norm(code(other) or "") == code_value for other in old) == 1 and sum(norm(code(other) or "") == code_value for other in rest) == 1 else None
        current_words = words(sheet)
        scores = []
        for other in candidates:
            old_words = words(other)
            union = current_words | old_words
            overlap = len(current_words & old_words)
            scores.append((overlap / len(union) if union else 0, other, overlap))
        scores.sort(key=lambda pair: pair[0], reverse=True)
        content_match = scores[0][1] if scores[0][2] >= 2 and scores[0][0] >= 0.6 and scores[0][0] - scores[1][0] >= 0.1 else None
        match = None if code_match is not None and content_match is not None and code_match is not content_match else code_match or content_match
        if match is not None:
            found[title(sheet)] = (match, "rule")
            left.remove(match)
        rest.remove(sheet)  # 冲突证据不足时不能再按同名配对。

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


def inferred_renames(matches: dict[str, tuple[Any, str]], title: Callable) -> dict[str, str]:
    return {title(source): current for current, (source, how) in matches.items()
            if how == "rule" and norm(title(source)) != norm(current)
            and norm(strip_new_prefix(title(source))) == norm(current)}


def require_manual_file_check(entry: dict) -> None:
    reason = "备份文件仅按文件名规则推测，请打开新旧文件人工确认；未自动修复"
    entry["error"] = reason
    for item in entry.get("cells", []):
        item.update(action="skip", reason=reason)
