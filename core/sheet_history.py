"""已保存的删表、改名记录，以及备份比较时的工作表身份和内容投影。"""
from __future__ import annotations

import copy
import hashlib
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import quote_sheetname

_records: Callable = lambda: {}


def configure(records: Callable) -> None:
    global _records
    _records = records


def key(path: str | Path) -> str:
    return str(Path(path).resolve()).casefold()


def fingerprint(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def get(path: str | Path) -> list[dict]:
    return copy.deepcopy(_records().get(key(path), []))


def record(path: str | Path, before: list[str], after: list[str], deleted: list[str],
           renamed: dict[str, str], preserved_values: list[dict], before_hash: str) -> dict:
    """调用方在 Excel 保存成功后立即持久化；失败的计划不属于操作历史。"""
    event = {"at": datetime.now().astimezone().isoformat(timespec="seconds"),
             "before": list(before), "after": list(after), "deleted": list(deleted),
             "renamed": dict(renamed), "preserved_values": copy.deepcopy(preserved_values),
             "before_hash": before_hash, "after_hash": fingerprint(path)}
    _records().setdefault(key(path), []).append(event)
    return event


def move(old_path: str | Path, new_path: str | Path) -> None:
    old, new = key(old_path), key(new_path)
    if old != new and old in _records():
        _records()[new] = _records().pop(old)


def validate(records) -> bool:
    if not isinstance(records, dict):
        return False
    for path, events in records.items():
        if not isinstance(path, str) or not isinstance(events, list):
            return False
        for event in events:
            if not isinstance(event, dict):
                return False
            for name in ("before", "after", "deleted"):
                values = event.get(name)
                if not isinstance(values, list) or any(not isinstance(v, str) or not v for v in values):
                    return False
                if len({v.casefold() for v in values}) != len(values):
                    return False
            renamed = event.get("renamed")
            if not isinstance(renamed, dict) or any(not isinstance(k, str) or not k or
                    not isinstance(v, str) or not v for k, v in renamed.items()):
                return False
            if any(not isinstance(event.get(k), str) for k in ("at", "before_hash", "after_hash")):
                return False
            frozen = event.get("preserved_values")
            if not isinstance(frozen, list) or any(not isinstance(item, dict) or
                    any(not isinstance(item.get(k), str) for k in ("sheet", "addr", "formula", "kind")) or
                    "value" not in item for item in frozen):
                return False
            before = set(event["before"])
            deleted = set(event["deleted"])
            if not deleted <= before or not set(renamed) <= before - deleted:
                return False
            if [renamed.get(n, n) for n in event["before"] if n not in deleted] != event["after"]:
                return False
    return True


def formula_value(value, renamed: dict[str, str] | None = None):
    """只改本工作簿的工作表引用；统一引号，保留字符串常量和外部文件引用。"""
    if not isinstance(value, str) or not value.startswith("="):
        return value
    try:
        tokens = Tokenizer(value).items
    except Exception:
        return value
    names = {k.casefold(): v for k, v in (renamed or {}).items()}
    for token in tokens:
        if token.type != "OPERAND" or token.subtype != "RANGE" or "!" not in token.value:
            continue
        head, address = token.value.rsplit("!", 1)
        if head.startswith("'") and head.endswith("'"):
            head = head[1:-1].replace("''", "'")
        if "[" in head or "]" in head:
            continue
        head = ":".join(names.get(part.casefold(), part).casefold() for part in head.split(":"))
        token.value = f"{quote_sheetname(head)}!{address}"
    return "=" + "".join(token.value for token in tokens)


def _references_reused_name(value, names: set[str]) -> bool:
    if not names or not isinstance(value, str) or not value.startswith("="):
        return False
    try:
        tokens = Tokenizer(value).items
    except Exception:
        return True
    for token in tokens:
        if token.type == "FUNC" and token.subtype == "OPEN" and token.value.upper().split(".")[-1] == "INDIRECT(":
            return True  # 字符串拼接可能间接引用已复用的旧表名。
        if token.type != "OPERAND" or token.subtype != "RANGE" or "!" not in token.value:
            continue
        head = token.value.rsplit("!", 1)[0].strip("'").replace("''", "'")
        if "[" not in head and (":" in head or head.casefold() in names):
            return True
    return False


@dataclass
class Projection:
    # 包含 None 的名字禁止继续按同名/内容猜配，避免回退到已删掉的另一张表。
    origins: dict[str, str | None] = field(default_factory=dict)
    operations: list[dict] = field(default_factory=list)
    frozen: list[dict] = field(default_factory=list)

    def __post_init__(self):
        self.frozen = [{(item["sheet"].casefold(), item["addr"].upper()): item for item in op["preserved_values"]
                        if item["kind"] == "cell"} for op in self.operations]

    def value(self, sheet: str, address: str, value):
        for op, frozen in zip(self.operations, self.frozen):
            item = frozen.get((sheet.casefold(), address.upper()))
            if item and formula_value(value) == formula_value(item["formula"]):
                value = item["value"]
            if _references_reused_name(value, set(op.get("reused_names", []))):
                return "[无法确认已复用的原表引用]" + value
            value = formula_value(value, op["renamed"])
            sheet = next((new for old, new in op["renamed"].items() if old.casefold() == sheet.casefold()), sheet)
        return formula_value(value)


def with_inferred_renames(projection: Projection, renamed: dict[str, str]) -> Projection:
    """仅将有代码名或内容证据的名称推断用于公式比较；不猜测删表后的转值。"""
    if not renamed:
        return projection
    return Projection(projection.origins, [*projection.operations, {"renamed": renamed, "preserved_values": [],
                                                                   "reused_names": [name.casefold() for name in renamed.values()]}])


def project(events: list[dict], current: list[str], old: list[str], old_hash: str = "") -> Projection:
    """按备份所在的历史阶段重放身份；重复表名阶段有歧义时拒绝猜测。"""
    if not events:
        return Projection()
    shape = lambda names: {n.casefold() for n in names}
    # 只用连续、且最终 Sheet 集合仍与当前相符的那一段，防止套用到替换过的文件。
    if shape(events[-1]["after"]) != shape(current):
        return Projection()
    start = len(events) - 1
    while start and events[start - 1]["after"] == events[start]["before"]:
        start -= 1
    events = events[start:]
    states = [op["before"] for op in events] + [events[-1]["after"]]
    candidates = [i for i, names in enumerate(states) if shape(names) == shape(old)]
    exact = [i for i in candidates if old_hash and
             ((i < len(events) and events[i]["before_hash"] == old_hash) or
              (i > 0 and events[i - 1]["after_hash"] == old_hash))]
    candidates = exact or candidates
    if not candidates:
        return Projection()
    old_names = {name.casefold(): name for name in old}
    results = []
    for i in candidates:
        origins = {name: old_names[name.casefold()] for name in states[i]}
        for op in events[i:]:
            origins = {op["renamed"].get(n, n): source for n, source in origins.items() if n not in op["deleted"]}
        results.append({name: next((source for n, source in origins.items() if n.casefold() == name.casefold()), None)
                        for name in current})
    origins = {name: results[0][name] if all(r[name] == results[0][name] for r in results) else None
               for name in current}
    # 内容投影也必须有唯一历史阶段；身份相同但阶段不明时仍保留严格内容检查。
    return Projection(origins, events[candidates[0]:] if len(candidates) == 1 else [])


def for_paths(path: str | Path, old_path: str | Path, current: list[str], old: list[str],
              *, old_hash: str = "") -> Projection:
    events = get(path)
    return project(events, current, old, old_hash or fingerprint(old_path)) if events else Projection()
