"""仿 Git 的 Excel 对比（用户 2026-11-11 要）：上传前看「你要上传的这份」和「远程现在这份」到底哪里不一样。

没有 hash 时只比文字容易出问题，最稳妥是让人自己看一眼。这里把两份 Excel 的差异全列出来，网页上左右两栏显示
（左 = 远程现在这份，右 = 你要上传的这份），用户看了自己判断能不能覆盖：

- 表：只有一边有的表（精简注定会删的表单独标出来，不算差异）
- 每个格子：显示的文字、公式、字体（名字 / 字号 / 粗斜体 / 下划线 / 删除线）、字的颜色、填充色、边框、对齐 / 换行、数字格式
- 整张表：行高、列宽、隐藏的行列、合并单元格、打印设置（打印区域 / 方向 / 缩放 / 纸张）、图片图形个数、标签颜色、表是否隐藏、条件格式条数

只读文件，不用 Excel（openpyxl），所以只支持 .xlsx / .xlsm；图片图形只比个数，位置大小要在 Excel 里看。
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from core import content_check, debug_log

CELL_KINDS = ("文字", "公式", "字体", "字的颜色", "填充色", "边框", "对齐", "数字格式")
SHEET_KINDS = ("行高", "列宽", "合并单元格", "打印设置", "图片图形", "标签颜色", "显示隐藏", "条件格式")
MAX_TEXT_ROWS = 400      # 每张表最多列多少个文字 / 公式不一样的格子
MAX_STYLE_GROUPS = 60    # 每张表每类样式最多列多少组
MAX_GROUP_CELLS = 12     # 每组样式举几个格子
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
_BORDER_WORDS = {"thin": "细", "medium": "中", "thick": "粗", "hair": "极细", "dotted": "点", "dashed": "虚",
                 "double": "双", "dashDot": "点划", "mediumDashed": "中虚", "dashDotDot": "双点划",
                 "mediumDashDot": "中点划", "mediumDashDotDot": "中双点划", "slantDashDot": "斜点划"}
_ALIGN_WORDS = {"left": "左", "center": "中", "right": "右", "top": "上", "bottom": "下", "justify": "两端",
                "distributed": "分散", "general": "", "fill": "填满", "centerContinuous": "跨列居中"}


def _color(color: Any) -> str:
    """openpyxl 的颜色 → 「#RRGGBB」/「主题色4+40%」/「自动」；没有 → 空"""
    if color is None:
        return ""
    kind = getattr(color, "type", None)
    if kind == "rgb":
        value = color.rgb
        return f"#{value[-6:].upper()}" if isinstance(value, str) and len(value) >= 6 else ""
    if kind == "theme":
        tint = getattr(color, "tint", 0) or 0
        return f"主题色{color.theme}" + (f"{tint:+.0%}" if abs(tint) > 0.001 else "")
    if kind == "indexed":
        from openpyxl.styles.colors import COLOR_INDEX
        index = color.indexed
        if isinstance(index, int) and 0 <= index < len(COLOR_INDEX):
            return f"#{COLOR_INDEX[index][-6:].upper()}"
        return "自动"
    return "自动" if kind == "auto" else ""


def _font(font: Any) -> tuple[str, str]:
    if font is None:
        return "", ""
    size = font.sz
    size_text = "" if size is None else (str(int(size)) if float(size).is_integer() else str(size))
    flags = "".join(word for on, word in ((font.b, " 粗体"), (font.i, " 斜体"),
                                          (font.u, " 下划线"), (font.strike, " 删除线")) if on)
    return f"{font.name or ''} {size_text}".strip() + flags, _color(font.color)


def _fill(fill: Any) -> str:
    if fill is None or not getattr(fill, "fill_type", None):
        return ""
    if fill.fill_type == "solid":
        return _color(fill.fgColor) or "自动"
    return f"{fill.fill_type} {_color(fill.fgColor)}".strip()


def _border(border: Any) -> str:
    if border is None:
        return ""
    parts = []
    for side, word in (("left", "左"), ("right", "右"), ("top", "上"), ("bottom", "下")):
        style = getattr(getattr(border, side, None), "style", None)
        if style:
            parts.append(word + _BORDER_WORDS.get(style, style))
    return " ".join(parts)


def _align(alignment: Any) -> str:
    if alignment is None:
        return ""
    parts = []
    if alignment.horizontal and _ALIGN_WORDS.get(alignment.horizontal, alignment.horizontal):
        parts.append("横向" + _ALIGN_WORDS.get(alignment.horizontal, alignment.horizontal))
    if alignment.vertical and alignment.vertical != "bottom":
        parts.append("纵向" + _ALIGN_WORDS.get(alignment.vertical, alignment.vertical))
    if alignment.wrap_text:
        parts.append("自动换行")
    if alignment.shrink_to_fit:
        parts.append("缩小字体填充")
    if alignment.text_rotation:
        parts.append(f"旋转{alignment.text_rotation}")
    return " ".join(parts)


def _style_of(cell: Any, cache: dict[tuple, dict[str, str]]) -> dict[str, str]:
    key = tuple(cell._style) if cell.has_style else ()
    found = cache.get(key)
    if found is None:
        font, font_color = _font(cell.font)
        number_format = cell.number_format or ""
        found = cache[key] = {"字体": font, "字的颜色": font_color, "填充色": _fill(cell.fill), "边框": _border(cell.border),
                              "对齐": _align(cell.alignment), "数字格式": "" if number_format == "General" else number_format}
    return found


def _drawing_counts(path: Path) -> dict[str, int]:
    """每张表的图片 / 图形个数（数 drawing XML 里的锚点；openpyxl 不读自选图形）"""
    counts: dict[str, int] = {}
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            book = ElementTree.fromstring(z.read("xl/workbook.xml"))
            ns = book.tag[: book.tag.index("}") + 1] if book.tag.startswith("{") else ""
            rels = ElementTree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            targets = {rel.get("Id"): rel.get("Target") for rel in rels.findall(_PKG_REL)}
            for sheet in book.iter(f"{ns}sheet"):
                target = targets.get(sheet.get(f"{_REL_NS}id")) or ""
                part = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
                rel_part = f"{part.rsplit('/', 1)[0]}/_rels/{part.rsplit('/', 1)[-1]}.rels"
                total = 0
                if rel_part in names:
                    for rel in ElementTree.fromstring(z.read(rel_part)).findall(_PKG_REL):
                        if not (rel.get("Type") or "").endswith("/drawing"):
                            continue
                        drawing = rel.get("Target") or ""
                        full = drawing.lstrip("/") if drawing.startswith("/") else str(
                            Path(part).parent.joinpath(drawing).as_posix())
                        parts, out = full.split("/"), []
                        for piece in parts:  # 处理 ../
                            if piece == "..":
                                out and out.pop()
                            elif piece != ".":
                                out.append(piece)
                        full = "/".join(out)
                        if full in names:
                            total += len(re.findall(rb"<(?:\w+:)?(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)\b", z.read(full)))
                counts[sheet.get("name") or ""] = total
    except Exception as err:
        debug_log.event("⚠️ Excel 对比：读不了图片图形个数", 文件=str(path), 原因=str(err)[:80])
    return counts


def _num(value: Any) -> str:
    if value is None:
        return "默认"
    value = round(float(value), 2)
    return str(int(value)) if value.is_integer() else str(value)


def _page_setup(sheet: Any, area: str) -> dict[str, str]:
    setup, props = sheet.page_setup, sheet.sheet_properties.pageSetUpPr
    fit = bool(props is not None and props.fitToPage)
    if fit:
        zoom = f"宽 {setup.fitToWidth if setup.fitToWidth is not None else 1} 页 × 高 " \
               f"{'自动' if setup.fitToHeight == 0 else (setup.fitToHeight if setup.fitToHeight is not None else 1)} 页"
    else:
        zoom = f"{setup.scale or 100}%"
    return {"打印区域": area, "方向": {"landscape": "横向", "portrait": "纵向"}.get(setup.orientation or "", "默认"),
            "缩放": zoom, "纸张": str(setup.paperSize or "默认")}


def _load(path: Path) -> dict[str, Any]:
    if not path.name.lower().endswith(content_check.READABLE_SUFFIXES):
        raise ValueError(f"{path.suffix or '这种'}格式没法对比，只支持 .xlsx / .xlsm：{path.name}")
    if not path.is_file():
        raise ValueError(f"找不到文件：{path}")
    from openpyxl import load_workbook
    try:
        texts = content_check.snapshot(path)
        book = load_workbook(path)  # 要读样式，不能用只读模式
    except content_check.Unreadable as err:
        raise ValueError(f"{path.name}：{err}") from err
    except Exception as err:
        raise ValueError(f"读不了 {path.name}（{type(err).__name__}: {err}）") from err
    return {"book": book, "texts": texts, "drawings": _drawing_counts(path), "cache": {},
            "doomed": content_check.doomed_sheets(path, [s.title for s in book.worksheets])}


def _sheet_info(data: dict[str, Any], name: str) -> dict[str, Any]:
    sheet = data["book"][name]
    text = data["texts"].get(name, {"cells": {}, "area": ""})
    rows = {r: (_num(d.height) + ("（隐藏）" if d.hidden else "")) for r, d in sheet.row_dimensions.items()
            if d.height is not None or d.hidden}
    cols: dict[str, str] = {}
    from openpyxl.utils.cell import get_column_letter
    for dim in sheet.column_dimensions.values():
        if dim.width is None and not dim.hidden:
            continue
        if not dim.customWidth and not dim.hidden:
            continue
        first, last = dim.min or 0, dim.max or 0
        if not first or last - first > 500:  # 「到最后一列」这种整段设置不展开
            continue
        for index in range(first, last + 1):
            cols[get_column_letter(index)] = _num(dim.width) + ("（隐藏）" if dim.hidden else "")
    tab = sheet.sheet_properties.tabColor
    return {"sheet": sheet, "texts": text["cells"], "rows": rows, "cols": cols,
            "merged": sorted(str(r) for r in sheet.merged_cells.ranges),
            "setup": _page_setup(sheet, text.get("area", "")),
            "drawings": data["drawings"].get(name, 0), "tab": _color(tab),
            "state": {"visible": "显示", "hidden": "隐藏", "veryHidden": "深度隐藏"}.get(sheet.sheet_state, sheet.sheet_state),
            "cf": sum(len(rules.rules) for rules in sheet.conditional_formatting)}


def _addr_key(addr: str) -> tuple[int, int, str]:
    letters = re.sub(r"\d", "", addr)
    return int(re.sub(r"\D", "", addr) or 0), len(letters), letters


def _formula(cell: Any) -> str:
    value = cell.value if cell is not None else None
    if getattr(cell, "data_type", "") == "f":
        return str(getattr(value, "text", value) or "")
    return ""


def _diff_sheet(name: str, old: dict[str, Any], new: dict[str, Any], caches: tuple[dict, dict]) -> dict[str, Any]:
    """old = 远程（左），new = 要上传的（右）"""
    a, b = _sheet_info(old, name), _sheet_info(new, name)
    counts: dict[str, int] = {}
    # ── 文字 ──
    ta, tb = a["texts"], b["texts"]
    changed = sorted((k for k in ta.keys() | tb.keys() if ta.get(k, "") != tb.get(k, "")), key=_addr_key)
    counts["文字"] = len(changed)
    text_rows = [{"cell": k, "old": ta.get(k, ""), "new": tb.get(k, ""),
                  "kind": "add" if k not in ta else "del" if k not in tb else "mod"} for k in changed[:MAX_TEXT_ROWS]]
    # ── 公式、样式 ──
    ca, cb = a["sheet"]._cells, b["sheet"]._cells
    blank = {"字体": None, "字的颜色": None, "填充色": None, "边框": None, "对齐": None, "数字格式": None}
    formulas, groups = [], {kind: {} for kind in CELL_KINDS[2:]}
    from openpyxl.utils.cell import get_column_letter
    for key in sorted(ca.keys() | cb.keys()):
        cell_a, cell_b = ca.get(key), cb.get(key)
        addr = f"{get_column_letter(key[1])}{key[0]}"
        fa, fb = _formula(cell_a), _formula(cell_b)
        if fa != fb:
            formulas.append({"cell": addr, "old": fa, "new": fb, "kind": "add" if not fa else "del" if not fb else "mod"})
        if cell_a is None or cell_b is None:
            # 一边没有这个格子：只有带样式的才算样式不同（没字没样式的空格子不报）
            have = cell_a if cell_a is not None else cell_b
            if not have.has_style:
                continue
            sa = _style_of(cell_a, caches[0]) if cell_a is not None else blank
            sb = _style_of(cell_b, caches[1]) if cell_b is not None else blank
        else:
            if not cell_a.has_style and not cell_b.has_style:
                continue
            sa, sb = _style_of(cell_a, caches[0]), _style_of(cell_b, caches[1])
        for kind in CELL_KINDS[2:]:
            va, vb = sa[kind], sb[kind]
            if va is None or vb is None:  # 一边没有格子：对方是默认样式（空）就不算
                va, vb = va or "", vb or ""
                if kind == "字体":  # 字体名每个工作簿默认不同，没法和「没有格子」比
                    continue
            if va != vb:
                groups[kind].setdefault((va, vb), []).append(addr)
    counts["公式"] = len(formulas)
    styles = []
    for kind, found in groups.items():
        counts[kind] = sum(len(cells) for cells in found.values())
        ordered = sorted(found.items(), key=lambda item: -len(item[1]))
        for (va, vb), cells in ordered[:MAX_STYLE_GROUPS]:
            styles.append({"kind": kind, "old": va, "new": vb, "count": len(cells), "cells": cells[:MAX_GROUP_CELLS]})
        if len(ordered) > MAX_STYLE_GROUPS:
            styles.append({"kind": kind, "old": "", "new": "", "count": sum(len(c) for _, c in ordered[MAX_STYLE_GROUPS:]),
                           "cells": [], "more": len(ordered) - MAX_STYLE_GROUPS})
    # ── 整张表 ──
    whole = []
    for kind, field, label in (("行高", "rows", "第 {} 行"), ("列宽", "cols", "{} 列")):
        da, db = a[field], b[field]
        keys = [k for k in da.keys() | db.keys() if da.get(k, "默认") != db.get(k, "默认")]
        keys.sort(key=(lambda k: k) if field == "rows" else (lambda k: (len(k), k)))
        counts[kind] = len(keys)
        merged: dict[tuple[str, str], list] = {}
        for k in keys:
            merged.setdefault((da.get(k, "默认"), db.get(k, "默认")), []).append(k)
        for (va, vb), items in sorted(merged.items(), key=lambda item: -len(item[1]))[:MAX_STYLE_GROUPS]:
            shown = "、".join(label.format(i) for i in items[:8]) + (f" 等 {len(items)} 处" if len(items) > 8 else "")
            whole.append({"kind": kind, "what": shown, "old": va, "new": vb})
    only_a = [m for m in a["merged"] if m not in set(b["merged"])]
    only_b = [m for m in b["merged"] if m not in set(a["merged"])]
    counts["合并单元格"] = len(only_a) + len(only_b)
    if only_a or only_b:
        short = lambda items: "、".join(items[:10]) + (f" 等 {len(items)} 处" if len(items) > 10 else "")  # noqa: E731
        whole.append({"kind": "合并单元格", "what": "合并的范围", "old": short(only_a), "new": short(only_b)})
    setup = [(k, a["setup"][k], b["setup"][k]) for k in a["setup"] if a["setup"][k] != b["setup"][k]]
    counts["打印设置"] = len(setup)
    whole += [{"kind": "打印设置", "what": k, "old": va, "new": vb} for k, va, vb in setup]
    for kind, field, what, fmt in (("图片图形", "drawings", "图片 / 图形个数", lambda v: f"{v} 个"),
                                   ("标签颜色", "tab", "标签颜色", lambda v: v or "没有"),
                                   ("显示隐藏", "state", "这张表", str),
                                   ("条件格式", "cf", "条件格式规则", lambda v: f"{v} 条")):
        counts[kind] = int(a[field] != b[field])
        if a[field] != b[field]:
            whole.append({"kind": kind, "what": what, "old": fmt(a[field]), "new": fmt(b[field])})
    counts = {k: v for k, v in counts.items() if v}
    return {"name": name, "status": "changed" if counts else "same", "counts": counts,
            "text": text_rows, "text_more": max(0, len(changed) - MAX_TEXT_ROWS),
            "formulas": formulas[:MAX_TEXT_ROWS], "formulas_more": max(0, len(formulas) - MAX_TEXT_ROWS),
            "styles": styles, "whole": whole}


def diff(mine: Path | str, theirs: Path | str) -> dict[str, Any]:
    """mine = 你要上传的这份（右 / 新），theirs = 远程现在这份（左 / 旧）"""
    mine, theirs = Path(mine), Path(theirs)
    with debug_log.side_task("excel_diff", {"mine": str(mine), "theirs": str(theirs)}):
        new, old = _load(mine), _load(theirs)
        try:
            names_new = [s.title for s in new["book"].worksheets]
            names_old = [s.title for s in old["book"].worksheets]
            caches = (old["cache"], new["cache"])
            sheets = []
            for name in names_old + [n for n in names_new if n not in names_old]:
                doomed = old["doomed"].get(name) or new["doomed"].get(name) or ""
                if name not in names_new:
                    sheets.append({"name": name, "status": "only_old", "doomed": doomed, "counts": {}})
                elif name not in names_old:
                    sheets.append({"name": name, "status": "only_new", "doomed": doomed, "counts": {}})
                else:
                    entry = _diff_sheet(name, old, new, caches)
                    entry["doomed"] = doomed
                    sheets.append(entry)
            order_changed = [n for n in names_old if n in names_new] != [n for n in names_new if n in names_old]
        finally:
            new["book"].close()
            old["book"].close()
        real = [s for s in sheets if not (s["doomed"] and s["status"] in ("only_old", "only_new"))]
        totals: dict[str, int] = {}
        for sheet in sheets:
            for kind, count in sheet["counts"].items():
                totals[kind] = totals.get(kind, 0) + count
        summary = {"表": len(sheets), "一样的表": sum(s["status"] == "same" for s in sheets),
                   "有不同的表": sum(s["status"] == "changed" for s in sheets),
                   "只有远程有": sum(s["status"] == "only_old" for s in real),
                   "只有你这份有": sum(s["status"] == "only_new" for s in real),
                   "精简会删不算": len(sheets) - len(real), "表的顺序变了": order_changed, "各类": totals}
        identical = not totals and not summary["只有远程有"] and not summary["只有你这份有"] and not order_changed
        debug_log.event("🔍 Excel 对比", 你这份=str(mine), 远程=str(theirs), 完全一样=identical,
                        **{k: v for k, v in summary.items() if k != "各类"},
                        各类="、".join(f"{k}{v}" for k, v in totals.items()) or "-")
        return {"mine": str(mine), "theirs": str(theirs), "identical": identical, "summary": summary, "sheets": sheets,
                "kinds": {"cell": list(CELL_KINDS), "sheet": list(SHEET_KINDS)}}
