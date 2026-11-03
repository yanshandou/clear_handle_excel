"""按分类示范学习修复（用户 2026-11-02 提出）：一次扫描出上千个问题，其实就那么几类。

1. 分类：在 6 种错误类型（一格有几种问题时按这个组合）下，再按「单格 / 合并格、原来有没有开换行、缺几行」细分，同一小类改法基本一样。
2. 示范：扫描完记下每个问题格的原样（行高、列宽、换行、缩小填充、字号、上方分页）。
   你在 Excel 里改好几个、保存，点「学一下」：程序重读这一类的格子，和原样不一样的就是你的示范。
3. 总结规律：对比改前改后，改了哪几样、改多少（固定值、按字号的倍数、按缺的行数、按文字长度），
   选几个示范之间最一致的算法，写成一句话给你看；示范之间对不上的单独列出来。
4. 批量修：你确认后，按规律改这一类剩下没动过的格子，保存，再整份复检。

这里只放纯逻辑和读写单元格属性的小函数；什么时候读、什么时候写由 core/batch.py 的任务决定。
"""
import re
import statistics
import unicodedata
from typing import Any

from core import autofix

XL_PAGEBREAK_MANUAL = -4135
MAX_ROW_HEIGHT = 409.0
MAX_COL_WIDTH = 255.0
HEIGHT_TOLERANCE = .4  # Excel 行高按屏幕像素取整，差这么点算没改（同 batch.ROW_HEIGHT_TOLERANCE）
WIDTH_TOLERANCE = .05
NEAR_TIE = .02  # 两种算法一致程度差不多时，按 MODEL_ORDER 里靠前的
OUTLIER = .25  # 某个示范的参数离中位数超过这么多（相对）就列为对不上

# 每种数值属性能用的算法，靠前的优先（只有一个示范、分不出哪种更一致时用第一个能算的）
MODEL_ORDER = {
    "row": ("lines", "font", "add", "ratio"),
    "col": ("text", "add", "ratio"),
    "font": ("set", "add"),
}

TYPE_ORDER = list(autofix.CELL_ERROR_SPECS)


# ───────────────────────── 分类 ─────────────────────────

def error_types(card: dict) -> list[str]:
    """这一格的问题类型，按 CELL_ERROR_SPECS 的顺序排好、去重"""
    types = [t for t in card.get("error_types") or [] if t in autofix.CELL_ERROR_SPECS]
    if not types:
        types = [autofix.item_error_type(p) for p in card.get("problems") or []] or [autofix.ERROR_TYPE_BORDER_CLIP]
    return sorted(dict.fromkeys(types), key=lambda t: TYPE_ORDER.index(t) if t in TYPE_ORDER else len(TYPE_ORDER))


def missing_lines(items: list[dict]) -> int | None:
    found = [int(it["missing_lines"]) for it in items
             if isinstance(it.get("missing_lines"), (int, float)) and not isinstance(it.get("missing_lines"), bool)
             and it["missing_lines"] > 0]
    return max(found) if found else None


def is_merged(card: dict, baseline: dict | None) -> bool:
    if baseline:
        return len(baseline.get("heights") or []) > 1 or len(baseline.get("widths") or []) > 1
    return ":" in (card.get("cell_addr") or "")


def category(card: dict, items: list[dict], baseline: dict | None) -> tuple[str, str]:
    """返回 (分类键, 给人看的名字)。名字例：「文字截断 · 单格 · 已开换行 · 缺 2 行」。
    一格同时有几种问题（比如压线 + 截断）时，这个组合单独成一类（用户 2026-11-03 定）：
    一次改动没法拆给两个问题，按组合示范、按组合学、按组合修。键里用 + 连起来，名字「边框压线 + 文字截断 · …」"""
    types = error_types(card)
    parts = ["+".join(types)]
    names = [" + ".join(autofix.error_type_label(t) for t in types)]
    if any(t != autofix.ERROR_TYPE_TEXT_MISSING for t in types):
        merged = is_merged(card, baseline)
        parts.append("merged" if merged else "single")
        names.append("合并格" if merged else "单格")
    if any(t in autofix.CUT_ERROR_TYPES for t in types):
        wrap = (baseline or {}).get("wrap")
        parts.append({True: "wrap", False: "nowrap"}.get(wrap, "wrap?"))
        names.append({True: "已开换行", False: "没开换行"}.get(wrap, "换行未知"))
    if autofix.ERROR_TYPE_TEXT_TRUNCATED in types:
        lines = missing_lines(items)
        if lines:
            parts.append(f"lines{min(lines, 3)}")
            names.append(f"缺 {lines} 行" if lines < 3 else "缺 3 行以上")
    return "|".join(parts), " · ".join(names)


def text_width(text: str) -> int:
    """按显示宽度数字数：全角字算 2，半角算 1（同 batch._text_width）。多行取最长一行"""
    lines = str(text or "").splitlines() or [""]
    return max(sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in line) for line in lines)


# ───────────────────────── 读写单元格属性 ─────────────────────────

def _first_cell(addr: str) -> str:
    first = (addr or "").split(":")[0].replace("$", "")
    if not re.fullmatch(r"[A-Za-z]{1,3}[1-9]\d*", first):
        raise ValueError(f"单元格地址无法确认：{addr}")
    return first


def _num(value: Any) -> float | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None  # 合并区域里各格不一样时 Excel 返回 None
    return round(float(value), 2)


def _flag(value: Any) -> bool | None:
    return value if isinstance(value, bool) else None


def read_props(sheet: Any, addr: str) -> dict:
    """读一格（合并单元格按整块）现在的样子"""
    area = sheet.Range(_first_cell(addr)).MergeArea
    top, left = int(area.Row), int(area.Column)
    rows, cols = int(area.Rows.Count), int(area.Columns.Count)
    try:
        page_break = int(sheet.Rows(top).PageBreak) == XL_PAGEBREAK_MANUAL
    except Exception:
        page_break = None
    return {
        "top": top, "left": left,
        "heights": [_num(sheet.Rows(r).RowHeight) for r in range(top, top + rows)],
        "widths": [_num(sheet.Columns(c).ColumnWidth) for c in range(left, left + cols)],
        "wrap": _flag(area.WrapText), "shrink": _flag(area.ShrinkToFit),
        "font": _num(area.Cells(1, 1).Font.Size), "break": page_break,
    }


def _total(values: list) -> float | None:
    if not values or any(v is None for v in values):
        return None
    return round(sum(values), 2)


def same(a: dict, b: dict) -> bool:
    """两次读到的样子算不算一样（行高差一点点不算改过）"""
    return not changes(a, b)


def changes(before: dict, after: dict) -> dict:
    """改了哪几样：{属性: (改前, 改后)}；行高、列宽按整块的合计比"""
    found: dict[str, tuple] = {}
    if len(before.get("heights") or []) != len(after.get("heights") or []) or \
            len(before.get("widths") or []) != len(after.get("widths") or []) or \
            (before.get("top"), before.get("left")) != (after.get("top"), after.get("left")):
        found["merge"] = (before.get("heights"), after.get("heights"))  # 合并 / 取消合并、插删行列：学不了
        return found
    for key, tolerance in (("heights", HEIGHT_TOLERANCE), ("widths", WIDTH_TOLERANCE)):
        a, b = _total(before.get(key) or []), _total(after.get(key) or [])
        if a is not None and b is not None and abs(a - b) > tolerance:
            found["row" if key == "heights" else "col"] = (a, b)
    for key in ("wrap", "shrink", "break"):
        if before.get(key) is not None and after.get(key) is not None and before[key] != after[key]:
            found[key] = (before[key], after[key])
    a, b = before.get("font"), after.get("font")
    if a is not None and b is not None and abs(a - b) > .01:
        found["font"] = (a, b)
    return found


def write_props(sheet: Any, addr: str, current: dict, target: dict) -> None:
    """按 plan() 算出的改法写进去（只改 target 里有的几样）"""
    area = sheet.Range(_first_cell(addr)).MergeArea
    if "wrap" in target:
        area.WrapText = bool(target["wrap"])
    if "shrink" in target:
        area.ShrinkToFit = bool(target["shrink"])
    if "font" in target:
        area.Font.Size = target["font"]
    if "row" in target:
        bottom = current["top"] + len(current["heights"]) - 1  # 合并单元格加在最下面一行（同扫描时的加高）
        sheet.Rows(bottom).RowHeight = target["row"]
    if "col" in target:
        right = current["left"] + len(current["widths"]) - 1
        sheet.Columns(right).ColumnWidth = target["col"]
    if target.get("break"):
        sheet.Rows(current["top"]).PageBreak = XL_PAGEBREAK_MANUAL


XL_PAGEBREAK_NONE = -4142


def restore_props(sheet: Any, addr: str, before: dict, now: dict) -> list[str]:
    """按规律改得不对：把这一格（合并格整块）和改之前不一样的几样改回去，返回改回了哪几样。
    行高、列宽按整块每一行 / 每一列写回；换行、缩小填充、字号、上方分页按记下的值写回"""
    diff = changes(before, now)
    if "merge" in diff:
        raise ValueError("合并范围或行列位置变了，不能按原样改回")
    area = sheet.Range(_first_cell(addr)).MergeArea
    top, left = before["top"], before["left"]
    if "row" in diff:
        for i, height in enumerate(before["heights"]):
            sheet.Rows(top + i).RowHeight = height
    if "col" in diff:
        for i, width in enumerate(before["widths"]):
            sheet.Columns(left + i).ColumnWidth = width
    if "wrap" in diff:
        area.WrapText = bool(before["wrap"])
    if "shrink" in diff:
        area.ShrinkToFit = bool(before["shrink"])
    if "font" in diff:
        area.Font.Size = before["font"]
    if "break" in diff:
        sheet.Rows(top).PageBreak = XL_PAGEBREAK_MANUAL if before["break"] else XL_PAGEBREAK_NONE
    return list(diff)


# ───────────────────────── 总结规律 ─────────────────────────

def _params(prop: str, model: str, demo: dict) -> float | None:
    before, after = demo["change"][prop]
    font, lines, width = demo.get("font"), demo.get("lines"), demo.get("text_width")
    delta = after - before
    if model == "add":
        return delta
    if model == "set":
        return after
    if model == "ratio":
        return after / before if before else None
    if model == "font":
        return delta / font if font else None
    if model == "lines":
        return delta / (lines * font) if lines and font else None
    if model == "text":
        return after / width if width and delta > 0 else None
    return None


def _spread(values: list[float]) -> float:
    middle = statistics.median(values)
    return (max(values) - min(values)) / max(abs(middle), 1e-6)


def _pick_model(prop: str, demos: list[dict]) -> tuple[str, float, list[str]]:
    """选几个示范之间最一致的算法，返回 (算法, 参数中位数, 对不上的示范)"""
    best = None
    for model in MODEL_ORDER[prop]:
        values = [_params(prop, model, d) for d in demos]
        if any(v is None for v in values):
            continue
        spread = _spread(values)
        if best is None or spread < best[0] - NEAR_TIE:
            best = (spread, model, values)
    if best is None:
        values = [_params(prop, "add", d) for d in demos]
        best = (_spread(values), "add", values)
    _, model, values = best
    middle = statistics.median(values)
    off = [d["id"] for d, v in zip(demos, values) if abs(v - middle) > OUTLIER * max(abs(middle), 1e-6)]
    return model, round(middle, 3), off


BOOL_TEXT = {
    "wrap": ("开自动换行", "关掉自动换行"),
    "shrink": ("开「缩小字体填充」", "关掉「缩小字体填充」"),
    "break": ("在上方插手动分页", "去掉上方的手动分页"),
}


def _num_text(x: float) -> str:
    return f"{x:.2f}".rstrip("0").rstrip(".")


def describe(step: dict) -> str:
    prop, model, k = step["prop"], step.get("model"), step.get("value")
    if prop in BOOL_TEXT:
        return BOOL_TEXT[prop][0 if k else 1]
    sign = "加" if (k or 0) >= 0 else "减"
    if prop == "row":
        return {
            "add": f"行高{sign} {_num_text(abs(k))}pt",
            "font": f"行高{sign}「字号 × {_num_text(abs(k))}」",
            "lines": f"行高{sign}「缺的行数 × 字号 × {_num_text(abs(k))}」",
            "ratio": f"行高 × {_num_text(k)}",
        }[model] + "（合并格加在最下面一行）"
    if prop == "col":
        return {
            "add": f"列宽{sign} {_num_text(abs(k))}",
            "text": f"列宽改成「文字长度 × {_num_text(k)}」（只加宽不变窄）",
            "ratio": f"列宽 × {_num_text(k)}",
        }[model] + "（合并格改最右边一列）"
    if prop == "font":
        return f"字号改成 {_num_text(k)}" if model == "set" else f"字号{sign} {_num_text(abs(k))}"
    return prop


def learn(demos: list[dict]) -> dict:
    """demos: [{"id", "label"（表!格）, "change": changes() 的结果, "font", "lines", "text_width"}]
    返回规律：{"ok", "steps": [...], "text", "conflicts": [{"id", "label", "reason"}], "demo_ids"}"""
    usable = [d for d in demos if d["change"] and "merge" not in d["change"]]
    conflicts = [{"id": d["id"], "label": d["label"], "reason": "合并 / 取消合并单元格或插删了行列，学不了这种改法"}
                 for d in demos if "merge" in d["change"]]
    rule = {"ok": False, "steps": [], "text": "", "conflicts": conflicts,
            "demo_ids": [d["id"] for d in demos], "demo_count": len(usable)}
    if not usable:
        rule["text"] = "示范里没看到能学的改动（行高、列宽、换行、缩小填充、字号、分页）"
        return rule
    labels = {d["id"]: d["label"] for d in usable}
    props = [p for p in ("wrap", "shrink", "font", "row", "col", "break") if any(p in d["change"] for d in usable)]
    for prop in props:
        changed = [d for d in usable if prop in d["change"]]
        unchanged = [d for d in usable if prop not in d["change"]]
        if len(changed) < len(unchanged):
            # 少数几个示范改了这一样，多数没改：不算规律，把这几个列出来
            for d in changed:
                conflicts.append({"id": d["id"], "label": labels[d["id"]],
                                  "reason": f"只有少数示范改了：{_change_text(prop, d['change'][prop])}，没算进规律"})
            continue
        for d in unchanged:
            conflicts.append({"id": d["id"], "label": labels[d["id"]], "reason": f"这一格没有{_prop_name(prop)}"})
        if prop in BOOL_TEXT:
            values = [d["change"][prop][1] for d in changed]
            value = statistics.mode(values)
            for d in changed:
                if d["change"][prop][1] != value:
                    conflicts.append({"id": d["id"], "label": labels[d["id"]],
                                      "reason": f"和别的示范相反：{BOOL_TEXT[prop][0 if d['change'][prop][1] else 1]}"})
            step = {"prop": prop, "value": value}
        else:
            model, value, off = _pick_model(prop, changed)
            for cid in off:
                conflicts.append({"id": cid, "label": labels[cid], "reason": f"{_prop_name(prop)}的量和别的示范差得多"})
            step = {"prop": prop, "model": model, "value": value}
        step["text"] = describe(step)
        rule["steps"].append(step)
    rule["ok"] = bool(rule["steps"])
    rule["text"] = "；".join(s["text"] for s in rule["steps"]) if rule["ok"] else "示范之间对不上，没有总结出共同的改法"
    return rule


def _prop_name(prop: str) -> str:
    return {"wrap": "改换行", "shrink": "改缩小填充", "font": "改字号", "row": "改行高", "col": "改列宽",
            "break": "插分页"}.get(prop, prop)


def _change_text(prop: str, change: tuple) -> str:
    before, after = change
    if prop in BOOL_TEXT:
        return BOOL_TEXT[prop][0 if after else 1]
    name = {"font": "字号", "row": "行高", "col": "列宽"}[prop]
    return f"{name} {_num_text(before)} → {_num_text(after)}"


# ───────────────────────── 套用规律 ─────────────────────────

def plan(rule: dict, current: dict, *, lines: int | None, text: str) -> dict:
    """按规律算这一格要改成什么：{属性: 新值}；不用改返回 {}"""
    target: dict[str, Any] = {}
    font = current.get("font") or 11.0
    for step in rule.get("steps") or []:
        prop, model, k = step["prop"], step.get("model"), step.get("value")
        if prop in BOOL_TEXT:
            if current.get(prop) != k:
                target[prop] = k
            continue
        if prop == "font":
            new = k if model == "set" else font + k
            new = round(min(409.0, max(1.0, new)), 1)
            if abs(new - font) > .01:
                target["font"] = new
            continue
        values = current["heights"] if prop == "row" else current["widths"]
        if any(v is None for v in values):
            continue
        total, last = sum(values), values[-1]
        if model == "add":
            delta = k
        elif model == "font":
            delta = k * font
        elif model == "lines":
            delta = k * (lines or 1) * font
        elif model == "ratio":
            delta = total * (k - 1)
        else:  # text：整块宽度改成文字长度 × k，只加宽
            delta = max(0.0, k * text_width(text) - total)
        limit = MAX_ROW_HEIGHT if prop == "row" else MAX_COL_WIDTH
        new = round(min(limit, max(0.0, last + delta)), 2)
        if abs(new - last) > (HEIGHT_TOLERANCE if prop == "row" else WIDTH_TOLERANCE):
            target[prop] = new
    return target
