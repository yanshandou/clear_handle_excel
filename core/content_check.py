"""按内容核对两份 Excel 能不能互相覆盖（用户 2026-11-11 要）。

上传时平常靠 hash（远程现在这份 和 上次拉取 / 上传时记下的版本 比）判断远程有没有被人改过。
但有时候没有 hash 记录，或者手上这份本来就是拿很久以前的版本修的（比如修复图片），
hash 一定对不上、没法用。这时改成直接比「要上传的这份」和「远程现在这份」：

- 全部文字：每张表每个格子显示的值（公式取 Excel 存下的结果）
- 每张表打印区域的格子数（没设打印区域的表按有内容的范围算）

两样都一样才算能覆盖；有一处不一样就不覆盖，提醒先拉取远程最新版、在上面改好再传。
只读文件，不用 Excel（openpyxl）。.xls / .xlsb 这类 openpyxl 读不了的格式返回「没法核对」，由调用方按原来的规则处理。
"""
from __future__ import annotations

import re
import zipfile
from datetime import date, datetime, time
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from core import debug_log

READABLE_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")
MAX_ROW, MAX_COL = 1048576, 16384
_NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
_SHOW_DIFFS = 3  # 每张表最多举几个不一样的格子


class Unreadable(Exception):
    pass


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    return str(value).replace("\r\n", "\n")


def _print_areas(path: Path) -> dict[str, str]:
    """workbook.xml 里每张表的打印区域（_xlnm.Print_Area，localSheetId = 第几张表）"""
    with zipfile.ZipFile(path) as z:
        root = ElementTree.fromstring(z.read("xl/workbook.xml"))
    names = [s.get("name", "") for s in root.findall("m:sheets/m:sheet", _NS)]
    areas = {}
    for item in root.findall("m:definedNames/m:definedName", _NS):
        if item.get("name", "").lower() != "_xlnm.print_area" or item.get("localSheetId") is None:
            continue
        index = int(item.get("localSheetId"))
        if 0 <= index < len(names):
            areas[names[index]] = (item.text or "").strip()
    return areas


def _area_cells(text: str) -> int | None:
    """「'表'!$A$1:$H$50,'表'!$J$1:$K$5」→ 格子总数；看不懂返回 None"""
    from openpyxl.utils.cell import range_boundaries
    total = 0
    parts = re.findall(r"(?:'(?:[^']|'')*'|[^',])+", text)
    if not parts:
        return None
    for part in parts:
        ref = part.rsplit("!", 1)[-1].replace("$", "").strip()
        try:
            left, top, right, bottom = range_boundaries(ref)
        except (ValueError, TypeError):
            return None
        left, right = left or 1, right or MAX_COL
        top, bottom = top or 1, bottom or MAX_ROW
        total += (right - left + 1) * (bottom - top + 1)
    return total


def snapshot(path: Path | str) -> dict[str, dict[str, Any]]:
    """{表名: {"cells": {"B5": "文字"}, "area": "A1:H50", "area_cells": 400}}。读不了抛 Unreadable"""
    path = Path(path)
    if not path.name.lower().endswith(READABLE_SUFFIXES):
        raise Unreadable(f"{path.suffix or '这种'}格式没法按内容核对")
    from openpyxl import load_workbook
    from openpyxl.utils.cell import get_column_letter
    try:
        areas = _print_areas(path)
        book = load_workbook(path, read_only=True, data_only=True)
    except Exception as err:
        raise Unreadable(f"读不了文件内容（{type(err).__name__}: {err}）") from err
    out: dict[str, dict[str, Any]] = {}
    try:
        for sheet in book.worksheets:
            cells: dict[str, str] = {}
            top = left = None
            bottom = right = 0
            for row in sheet.iter_rows():
                for cell in row:
                    value = getattr(cell, "value", None)
                    text = _text(value)
                    if not text:
                        continue
                    r, c = cell.row, cell.column
                    cells[f"{get_column_letter(c)}{r}"] = text
                    top = r if top is None else min(top, r)
                    left = c if left is None else min(left, c)
                    bottom, right = max(bottom, r), max(right, c)
            area = areas.get(sheet.title, "")
            if area:
                count = _area_cells(area)
                label = ",".join(part.rsplit("!", 1)[-1] for part in
                                 re.findall(r"(?:'(?:[^']|'')*'|[^',])+", area)).replace("$", "")
                if count is None:
                    raise Unreadable(f"表「{sheet.title}」的打印区域看不懂：{area}")
            elif top is None:
                count, label = 0, "没设打印区域、表是空的"
            else:
                count = (bottom - top + 1) * (right - left + 1)
                label = (f"没设打印区域，按有内容的范围 {get_column_letter(left)}{top}:"
                         f"{get_column_letter(right)}{bottom}")
            out[sheet.title] = {"cells": cells, "area": label, "area_cells": count}
    except Unreadable:
        raise
    except Exception as err:
        raise Unreadable(f"读不了文件内容（{type(err).__name__}: {err}）") from err
    finally:
        book.close()
    return out


def _short(text: str, limit: int = 20) -> str:
    text = text.replace("\n", " ")
    return text if len(text) <= limit else text[:limit] + "…"


def compare(mine: Path | str, theirs: Path | str, mine_name: str = "本地", theirs_name: str = "远程") -> dict[str, Any]:
    """比两份 Excel 的全部文字和每张表打印区域格子数。
    返回 {"checked": 读得了没有, "same": 都一样没有, "diffs": [不一样的地方], "reason": 读不了的原因}"""
    try:
        a, b = snapshot(mine), snapshot(theirs)
    except Unreadable as err:
        debug_log.event("🔎 按内容核对：读不了", 文件=str(mine), 远程=str(theirs), 原因=str(err))
        return {"checked": False, "same": False, "diffs": [], "reason": str(err)}
    diffs = []
    for name in [n for n in b if n not in a]:
        diffs.append(f"表「{name}」只有{theirs_name}有")
    for name in [n for n in a if n not in b]:
        diffs.append(f"表「{name}」只有{mine_name}有")
    for name in [n for n in a if n in b]:
        ca, cb = a[name]["cells"], b[name]["cells"]
        changed = sorted({k for k in ca.keys() | cb.keys() if ca.get(k, "") != cb.get(k, "")},
                         key=lambda k: (int(re.sub(r"\D", "", k)), len(re.sub(r"\d", "", k)), k))
        if changed:
            examples = "；".join(f"{k} {mine_name}「{_short(ca.get(k, ''))}」{theirs_name}「{_short(cb.get(k, ''))}」"
                                for k in changed[:_SHOW_DIFFS])
            diffs.append(f"表「{name}」文字有 {len(changed)} 处不一样（{examples}{'…' if len(changed) > _SHOW_DIFFS else ''}）")
        if a[name]["area_cells"] != b[name]["area_cells"]:
            diffs.append(f"表「{name}」打印区域格子数不一样：{mine_name} {a[name]['area_cells']}（{a[name]['area']}），"
                         f"{theirs_name} {b[name]['area_cells']}（{b[name]['area']}）")
    debug_log.event("🔎 按内容核对", 文件=str(mine), 远程=str(theirs), 一致=not diffs, 表数=len(a),
                    不一样=len(diffs), 第一处=diffs[0][:80] if diffs else "")
    return {"checked": True, "same": not diffs, "diffs": diffs, "reason": ""}


def describe(result: dict[str, Any], limit: int = 3) -> str:
    diffs = result.get("diffs") or []
    text = "；".join(diffs[:limit])
    return text + (f"；还有 {len(diffs) - limit} 处" if len(diffs) > limit else "")
