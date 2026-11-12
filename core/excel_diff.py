"""仿 Git 的 Excel 对比（用户 2026-11-11 要）：上传前看「你要上传的这份」和「远程现在这份」到底哪里不一样。

没有 hash 时只比文字容易出问题，最稳妥是让人自己看一眼。这里把两份 Excel 的差异全列出来，网页上左右两栏显示
（左 = 远程现在这份，右 = 你要上传的这份），用户看了自己判断能不能覆盖：

- 表：只有一边有的表（精简注定会删的表单独标出来，不算差异）
- 每个格子：显示的文字、公式、批注、超链接、字体（名字 / 字号 / 粗斜体 / 下划线 / 删除线）、字的颜色、填充色、边框、
  对齐 / 换行、数字格式、保护
- 整张表：行高、列宽、隐藏的行列、合并单元格、冻结窗格、打印设置、图片图形、标签颜色、表是否隐藏、条件格式

只读文件，不用 Excel（openpyxl），支持 .xlsx / .xlsm / .xltx / .xltm。
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from core import content_check, debug_log

CELL_KINDS = ("文字", "公式", "公式缓存", "批注", "超链接", "字体", "字的颜色", "填充色", "边框", "对齐", "数字格式", "保护")
STYLE_KINDS = ("字体", "字的颜色", "填充色", "边框", "对齐", "数字格式", "保护")
SHEET_KINDS = ("行高", "列宽", "合并单元格", "冻结窗格", "打印设置", "图片图形", "标签颜色", "显示隐藏", "条件格式")
# 主题色表缺失时使用槽位名。
_THEME_NAMES = {0: "背景色", 1: "正文色", 2: "辅助背景色", 3: "辅助正文色", 4: "强调色1", 5: "强调色2",
                6: "强调色3", 7: "强调色4", 8: "强调色5", 9: "强调色6"}
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


def _color(color: Any, theme: list[str]) -> str:
    """openpyxl 的颜色 → 「#RRGGBB」/「主题色4+40%」/「自动」；没有 → 空"""
    if color is None:
        return ""
    kind = getattr(color, "type", None)
    if kind == "rgb":
        value = color.rgb
        return f"#{value[-6:].upper()}" if isinstance(value, str) and len(value) >= 6 else ""
    if kind == "theme":
        resolved = _css_color(color, theme)
        tint = getattr(color, "tint", 0) or 0
        name = _THEME_NAMES.get(color.theme, f"主题色{color.theme}")
        label = f"主题{name}" + (f"（{'变浅' if tint > 0 else '变深'} {abs(tint):.0%}）" if abs(tint) > 0.001 else "")
        return f"{resolved} {label}" if resolved else label
    if kind == "indexed":
        from openpyxl.styles.colors import COLOR_INDEX
        index = color.indexed
        if isinstance(index, int) and 0 <= index < len(COLOR_INDEX):
            return f"#{COLOR_INDEX[index][-6:].upper()}"
        return "自动"
    return "自动" if kind == "auto" else ""


def _font(font: Any, theme: list[str]) -> tuple[str, str]:
    if font is None:
        return "", ""
    size = font.sz
    size_text = "" if size is None else (str(int(size)) if float(size).is_integer() else str(size))
    flags = "".join(word for on, word in ((font.b, " 粗体"), (font.i, " 斜体"),
                                          (font.u, " 下划线"), (font.strike, " 删除线")) if on)
    return f"{font.name or ''} {size_text}".strip() + flags, _color(font.color, theme)


def _fill(fill: Any, theme: list[str]) -> str:
    if fill is None or not getattr(fill, "fill_type", None):
        return ""
    if hasattr(fill, "stop"):
        direction = f"{fill.degree:g}°" if fill.type == "linear" else (
            f"左 {fill.left:g} 右 {fill.right:g} 上 {fill.top:g} 下 {fill.bottom:g}"
        )
        stops = "、".join(f"{stop.position:.8g}：{_color(stop.color, theme)}" for stop in fill.stop)
        return f"渐变 {fill.type} {direction} {stops}".strip()
    if fill.fill_type == "solid":
        return _color(fill.fgColor, theme) or "自动"
    return f"{fill.fill_type} 前景 {_color(fill.fgColor, theme)} 背景 {_color(fill.bgColor, theme)}"


def _border(border: Any, theme: list[str]) -> str:
    if border is None:
        return ""
    parts = []
    for side, word in (("left", "左"), ("right", "右"), ("top", "上"), ("bottom", "下"),
                       ("start", "起始"), ("end", "末尾"), ("horizontal", "内横"), ("vertical", "内竖"),
                       ("diagonal", "对角")):
        edge = getattr(border, side, None)
        style = getattr(edge, "style", None)
        if style:
            if side == "diagonal":
                word += ("↗" if border.diagonalUp else "") + ("↘" if border.diagonalDown else "")
            color = _color(edge.color, theme)
            parts.append(word + _BORDER_WORDS.get(style, style) + (f" {color}" if color else ""))
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


def _style_of(cell: Any, data: dict[str, Any]) -> dict[str, str]:
    cache, theme = data["cache"], data["theme"]
    key = tuple(cell._style) if cell.has_style else ()
    found = cache.get(key)
    if found is None:
        font, font_color = _font(cell.font, theme)
        number_format = cell.number_format or ""
        protection = cell.protection
        locked = "" if protection is None else ("" if protection.locked else "不锁定") + (" 隐藏公式" if protection.hidden else "")
        found = cache[key] = {"字体": font, "字的颜色": font_color, "填充色": _fill(cell.fill, theme), "边框": _border(cell.border, theme),
                              "对齐": _align(cell.alignment), "数字格式": "" if number_format == "General" else number_format,
                              "保护": locked.strip()}
    return found


_XDR = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_ANCHORS = ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")
_OBJECT_KINDS = {"pic": "图片", "sp": "图形", "grpSp": "组合", "cxnSp": "连接线", "graphicFrame": "图表"}
MAX_DRAWING_ROWS = 80    # 每张表最多列多少个不一样的图片 / 图形
VIEW_BOX_ROWS = 30       # 效果图上给一个图片 / 图形最多多画多少行


def _zip_join(base: str, target: str) -> str:
    """base 这个零件里写的相对路径 target → 压缩包里的完整名字"""
    full = target.lstrip("/") if target.startswith("/") else f"{base.rsplit('/', 1)[0]}/{target}"
    out: list[str] = []
    for piece in full.split("/"):
        if piece == "..":
            out and out.pop()
        elif piece != ".":
            out.append(piece)
    return "/".join(out)


def _rels(z: zipfile.ZipFile, part: str, names: set[str]) -> list[tuple[str, str, str]]:
    """零件的关系表 [(Id, 类型, 目标的完整名字)]"""
    rel_part = f"{part.rsplit('/', 1)[0]}/_rels/{part.rsplit('/', 1)[-1]}.rels"
    if rel_part not in names:
        return []
    return [(rel.get("Id") or "", rel.get("Type") or "", _zip_join(part, rel.get("Target") or ""))
            for rel in ElementTree.fromstring(z.read(rel_part)).findall(_PKG_REL) if rel.get("TargetMode") != "External"]


def _local(element: Any) -> str:
    return element.tag.rsplit("}", 1)[-1]


def _unwrap(element: Any) -> list[Any]:
    """mc:AlternateContent 包着的取 Choice（没有就 Fallback）里面的东西"""
    out = []
    for child in element:
        if _local(child) == "AlternateContent":
            inner = next((c for c in child if _local(c) == "Choice"), None)
            if inner is None:
                inner = next(iter(child), None)
            out += _unwrap(inner) if inner is not None else []
        else:
            out.append(child)
    return out


def _cell_of(marker: Any) -> tuple[int, int] | None:
    if marker is None:
        return None
    try:
        return int(marker.findtext(f"{_XDR}row")) + 1, int(marker.findtext(f"{_XDR}col")) + 1
    except (TypeError, ValueError):
        return None


def _parse_drawing(xml: bytes, media: Any) -> list[dict[str, Any]]:
    """一张表的 drawing XML → 每个图片 / 图形一条。media(rId) → 这张图片文件的内容（没有返回 None）。
    位置按挂在哪个格子（左上角～右下角），大小按文件里记的宽高（厘米，取 1 位小数）"""
    import hashlib
    out = []
    for anchor in _unwrap(ElementTree.fromstring(xml)):
        if _local(anchor) not in _ANCHORS:
            continue
        children = _unwrap(anchor)
        obj = next((c for c in children if _local(c) in _OBJECT_KINDS), None)
        if obj is None:
            continue
        kind = _OBJECT_KINDS[_local(obj)]
        prop = next((e for e in obj.iter() if _local(e) == "cNvPr"), None)
        if kind == "图形" and any(_local(e) == "cNvSpPr" and e.get("txBox") == "1" for e in obj.iter()):
            kind = "文本框"
        paragraphs = ["".join(t.text or "" for t in para.iter(f"{_A}t")) for para in obj.iter(f"{_A}p")]
        text = "\n".join(p for p in paragraphs if p.strip())
        images, image_bytes = [], 0
        for blip in obj.iter(f"{_A}blip"):
            data = media(blip.get(f"{_REL_NS}embed") or blip.get(f"{_REL_NS}link") or "")
            if data is not None:
                images.append(hashlib.md5(data).hexdigest())
                image_bytes += len(data)
        shapes = sorted(g.get("prst") or "" for g in obj.iter(f"{_A}prstGeom"))
        start = _cell_of(next((c for c in children if _local(c) == "from"), None))
        end = _cell_of(next((c for c in children if _local(c) == "to"), None))
        ext = next((c for c in children if _local(c) == "ext"), None)  # oneCellAnchor / absoluteAnchor 的大小
        if ext is None or _local(anchor) == "twoCellAnchor":
            ext = next((e for e in obj.iter(f"{_A}ext") if e.get("cx") and e.get("cy")), ext)
        size = None
        if ext is not None:
            try:
                size = (round(int(ext.get("cx")) / 360000, 1), round(int(ext.get("cy")) / 360000, 1))
            except (TypeError, ValueError):
                size = None
        out.append({"kind": kind, "name": (prop.get("name") if prop is not None else "") or "",
                    "hidden": bool(prop is not None and prop.get("hidden") in ("1", "true")),
                    "text": text, "images": sorted(images), "image_bytes": image_bytes, "shapes": shapes,
                    "from": start, "to": end, "size": size})
    return out


def _drawings(path: Path) -> dict[str, list[dict[str, Any]]]:
    """每张表的图片 / 图形（读 drawing XML；openpyxl 不读自选图形）"""
    found: dict[str, list[dict[str, Any]]] = {}
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            book = ElementTree.fromstring(z.read("xl/workbook.xml"))
            ns = book.tag[: book.tag.index("}") + 1] if book.tag.startswith("{") else ""
            targets = {rid: target for rid, _, target in _rels(z, "xl/workbook.xml", names)}
            for sheet in book.iter(f"{ns}sheet"):
                part = targets.get(sheet.get(f"{_REL_NS}id") or "") or ""
                items: list[dict[str, Any]] = []
                for _, kind, drawing in _rels(z, part, names) if part else []:
                    if not kind.endswith("/drawing") or drawing not in names:
                        continue
                    media = {rid: target for rid, _, target in _rels(z, drawing, names)}
                    items += _parse_drawing(z.read(drawing), lambda rid: z.read(media[rid]) if media.get(rid) in names else None)
                found[sheet.get("name") or ""] = items
    except Exception as err:
        debug_log.event("⚠️ Excel 对比：读不了图片图形", 文件=str(path), 原因=str(err)[:80])
    return found


def _drawing_where(item: dict[str, Any]) -> str:
    from openpyxl.utils.cell import get_column_letter
    cell = lambda at: f"{get_column_letter(at[1])}{at[0]}"  # noqa: E731
    if not item["from"]:
        return "固定位置"
    return f"{cell(item['from'])}～{cell(item['to'])}" if item["to"] else cell(item["from"])


def _drawing_label(item: dict[str, Any]) -> str:
    return f"{item['kind']}「{item['name']}」" if item["name"] else item["kind"]


def _drawing_content(item: dict[str, Any]) -> tuple:
    return item["kind"], tuple(item["images"]), item["text"], tuple(item["shapes"])


def _drawing_full(item: dict[str, Any]) -> str:
    parts = [f"在 {_drawing_where(item)}"]
    if item["size"]:
        parts.append(f"宽 {_num(item['size'][0])} × 高 {_num(item['size'][1])} 厘米")
    if item["text"]:
        parts.append(f"文字「{item['text']}」")
    if item["hidden"]:
        parts.append("隐藏")
    return "，".join(parts)


def _diff_drawings(old: list[dict[str, Any]], new: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """两边的图片 / 图形一个个配对，列出不一样的。配对顺序：完全一样 → 内容和名字一样（挪了 / 改了大小）
    → 名字和位置一样（内容换了）→ 位置一样 → 内容一样 → 名字一样；剩下的算多了 / 少了"""
    left, right = list(old), list(new)
    pairs: list[tuple[dict[str, Any] | None, dict[str, Any] | None]] = []
    keys = (lambda i: (_drawing_content(i), i["name"], i["from"], i["to"], i["size"], i["hidden"]),
            lambda i: (_drawing_content(i), i["name"]),
            lambda i: (i["kind"], i["name"], i["from"]),
            lambda i: (i["kind"], i["from"], i["to"]),
            lambda i: (_drawing_content(i),),
            lambda i: (i["kind"], i["name"]) if i["name"] else None)
    for level, key in enumerate(keys):
        waiting: dict[Any, list[dict[str, Any]]] = {}
        for item in left:
            if key(item) is not None:
                waiting.setdefault(key(item), []).append(item)
        for item in list(right):
            match = waiting.get(key(item)) if key(item) is not None else None
            if match:
                other = match.pop(0)
                left.remove(other)
                right.remove(item)
                if level:
                    pairs.append((other, item))
    pairs += [(item, None) for item in left] + [(None, item) for item in right]
    rows = []
    for a, b in pairs:
        if a is None or b is None:
            item = a or b
            rows.append({"what": f"{_drawing_label(item)}（{_drawing_where(item)}）", "changes": ["少了" if b is None else "多了"],
                         "old": _drawing_full(a) if a else "", "new": _drawing_full(b) if b else "", "a": a, "b": b})
            continue
        va, vb, changes = [], [], []
        if a["images"] != b["images"]:
            changes.append("图片内容")
            va.append(f"原来的图片（{max(1, a['image_bytes'] // 1024)} KB）")
            vb.append(f"换成了另一张图片（{max(1, b['image_bytes'] // 1024)} KB）")
        if a["text"] != b["text"]:
            changes.append("文字")
            va.append(f"文字「{a['text']}」")
            vb.append(f"文字「{b['text']}」")
        if (a["from"], a["to"]) != (b["from"], b["to"]):
            changes.append("位置")
            va.append(f"在 {_drawing_where(a)}")
            vb.append(f"在 {_drawing_where(b)}")
        if a["size"] != b["size"]:
            changes.append("大小")
            show = lambda size: f"宽 {_num(size[0])} × 高 {_num(size[1])} 厘米" if size else "没记大小"  # noqa: E731
            va.append(show(a["size"]))
            vb.append(show(b["size"]))
        if a["hidden"] != b["hidden"]:
            changes.append("显示隐藏")
            va.append("隐藏" if a["hidden"] else "显示")
            vb.append("隐藏" if b["hidden"] else "显示")
        if a["kind"] != b["kind"] or a["shapes"] != b["shapes"]:
            changes.append("形状")
            va.append(a["kind"] + (f"（{'、'.join(s for s in a['shapes'] if s)}）" if any(a["shapes"]) else ""))
            vb.append(b["kind"] + (f"（{'、'.join(s for s in b['shapes'] if s)}）" if any(b["shapes"]) else ""))
        if a["name"] != b["name"] and not changes:
            continue  # 只是名字不同（Excel 自动起的名），不算
        rows.append({"what": f"{_drawing_label(b)}（{_drawing_where(b)}）", "changes": changes,
                     "old": "，".join(va), "new": "，".join(vb), "a": a, "b": b})
    rows.sort(key=lambda row: (row["b"] or row["a"])["from"] or (0, 0))
    return rows


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
    return {"打印区域": area, "方向": "横向" if setup.orientation == "landscape" else "纵向",  # 没写就是纵向
            "缩放": zoom, "纸张": str(setup.paperSize or "")}


def _load(path: Path) -> dict[str, Any]:
    if not path.name.lower().endswith(content_check.READABLE_SUFFIXES):
        raise ValueError(f"{path.suffix or '这种'}格式没法对比，只支持 .xlsx / .xlsm / .xltx / .xltm：{path.name}")
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
    try:
        return {"book": book, "texts": texts, "drawings": _drawings(path), "cache": {}, "css": {},
                "theme": _theme_colors(path),
                "doomed": content_check.doomed_sheets(path, [s.title for s in book.worksheets])}
    except Exception:
        book.close()
        raise


def _conditional_rules(sheet: Any) -> dict[str, list[tuple[bytes, Any, str]]]:
    from openpyxl.xml.functions import tostring
    result = {}
    for group in sheet.conditional_formatting:
        ranges = " ".join(str(area) for area in sorted(group.sqref.ranges))
        rules = []
        for rule in group.rules:
            node = rule.to_tree()
            node.attrib.pop("dxfId", None)  # 比较规则的实际格式，工作簿里的样式编号可以不同
            operator = {"greaterThan": ">", "lessThan": "<", "greaterThanOrEqual": "≥",
                        "lessThanOrEqual": "≤", "equal": "=", "notEqual": "≠"}.get(rule.operator, rule.operator)
            description = " ".join(part for part in (operator, "、".join(rule.formula or ())) if part) or rule.type
            rules.append((tostring(node), rule.dxf, description[:120]))
        result[ranges] = rules
    return result


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
            "default_row": _num(sheet.sheet_format.defaultRowHeight),
            "default_col": _num(sheet.sheet_format.defaultColWidth),
            "base_col": _num(sheet.sheet_format.baseColWidth),
            "merged": sorted(str(r) for r in sheet.merged_cells.ranges),
            "setup": _page_setup(sheet, text.get("area", "") if sheet.print_area else "未设置"),
            "drawings": data["drawings"].get(name, []), "tab": _color(tab, data["theme"]),
            "state": {"visible": "显示", "hidden": "隐藏", "veryHidden": "深度隐藏"}.get(sheet.sheet_state, sheet.sheet_state),
            "cf": _conditional_rules(sheet)}


# ───────── 效果图（用户 2026-11-11）：左右各画一张表的样子，不一样的格子盖半透明色块；只画有不同的行和上下几行 ─────────
VIEW_CONTEXT = 3       # 不一样的行上下各带几行
VIEW_MAX_ROWS = 300    # 每张表最多画多少行
VIEW_MAX_COLS = 60     # 每张表最多画多少列
_THEME_SLOTS = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")
_DRAWING_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_BORDER_CSS = {"thin": "1px solid", "medium": "2px solid", "thick": "3px solid", "double": "3px double", "hair": "1px dotted",
               "dotted": "1px dotted", "dashed": "1px dashed", "dashDot": "1px dashed", "dashDotDot": "1px dashed",
               "mediumDashed": "2px dashed", "mediumDashDot": "2px dashed", "mediumDashDotDot": "2px dashed",
               "slantDashDot": "2px dashed"}


def _theme_colors(path: Path) -> list[str]:
    """主题色表（按 theme 编号排）：画效果图时把「主题色N」换成真的颜色"""
    try:
        with zipfile.ZipFile(path) as z:
            root = ElementTree.fromstring(z.read("xl/theme/theme1.xml"))
        scheme = root.find(f".//{_DRAWING_NS}clrScheme")
        found = {}
        for slot in (scheme if scheme is not None else []):
            for child in slot:
                value = child.get("lastClr") or child.get("val")
                if value and re.fullmatch(r"[0-9A-Fa-f]{6}", value):
                    found[slot.tag.replace(_DRAWING_NS, "")] = value.upper()
        return [found.get(slot, "") for slot in _THEME_SLOTS]
    except Exception:
        return []


def _css_color(color: Any, theme: list[str]) -> str:
    """openpyxl 颜色 → 「#RRGGBB」（画图用，主题色按主题表换算、明暗按 HLS 近似）；认不出 → 空"""
    if color is None:
        return ""
    kind = getattr(color, "type", None)
    value = ""
    if kind == "rgb" and isinstance(color.rgb, str) and len(color.rgb) >= 6:
        value = color.rgb[-6:]
    elif kind == "theme" and isinstance(color.theme, int) and 0 <= color.theme < len(theme):
        value = theme[color.theme]
    elif kind == "indexed":
        from openpyxl.styles.colors import COLOR_INDEX
        if isinstance(color.indexed, int) and 0 <= color.indexed < len(COLOR_INDEX):
            value = COLOR_INDEX[color.indexed][-6:]
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", value or ""):
        return ""
    tint = float(getattr(color, "tint", 0) or 0)
    if tint:
        import colorsys
        r, g, b = (int(value[i:i + 2], 16) / 255 for i in (0, 2, 4))
        h, l, sat = colorsys.rgb_to_hls(r, g, b)
        l = l * (1 + tint) if tint < 0 else l + (1 - l) * tint
        value = "".join(f"{round(max(0, min(1, x)) * 255):02X}" for x in colorsys.hls_to_rgb(h, l, sat))
    return f"#{value.upper()}"


def _cell_css(cell: Any, data: dict[str, Any]) -> list[Any]:
    """[css, 自动换行, 横向对齐, 纵向对齐]"""
    key = tuple(cell._style) if cell.has_style else ()
    found = data["css"].get(key)
    if found is not None:
        return found
    theme, parts = data["theme"], []
    font = cell.font
    if font is not None:
        if font.name:
            name = re.sub(r"""['"\\\x00-\x1f\x7f]""", lambda m: f"\\{ord(m[0]):x} ", font.name)
            parts.append(f"font-family:'{name}','MS PGothic','Yu Gothic','Meiryo',sans-serif")
        if font.sz:
            parts.append(f"font-size:{float(font.sz):g}pt")
        if font.b:
            parts.append("font-weight:700")
        if font.i:
            parts.append("font-style:italic")
        deco = " ".join(word for on, word in ((font.u, "underline"), (font.strike, "line-through")) if on)
        if deco:
            parts.append(f"text-decoration:{deco}")
        color = _css_color(font.color, theme)
        if color:
            parts.append(f"color:{color}")
    fill = cell.fill
    if fill is not None and fill.fill_type and hasattr(fill, "fgColor"):
        color = _css_color(fill.fgColor, theme)
        if color:
            parts.append(f"background:{color}")
    border = cell.border
    if border is not None:
        for side in ("left", "right", "top", "bottom"):
            edge = getattr(border, side, None)
            style = getattr(edge, "style", None)
            if style:
                parts.append(f"border-{side}:{_BORDER_CSS.get(style, '1px solid')} {_css_color(edge.color, theme) or '#000'}")
    align = cell.alignment
    wrap = bool(align is not None and align.wrap_text)
    horizontal = (align.horizontal if align is not None else None) or ""
    vertical = (align.vertical if align is not None else None) or "bottom"
    found = data["css"][key] = [";".join(parts), int(wrap), horizontal, vertical]
    return found


def _col_px(sheet: Any, last: int) -> list[int]:
    default = sheet.sheet_format.defaultColWidth or 8.43
    widths = [default] * (last + 1)
    for dim in sheet.column_dimensions.values():
        first, end = dim.min or 0, dim.max or 0
        if not first:
            continue
        for index in range(first, min(end, last) + 1):
            widths[index] = 0 if dim.hidden else (dim.width if dim.width is not None else default)
    return [round(w * 7 + 5) if w else 0 for w in widths]


def _row_px(sheet: Any, row: int) -> int:
    if row in sheet.row_dimensions:
        dim = sheet.row_dimensions[row]
        if dim.hidden:
            return 0
        if dim.height is not None:
            return round(dim.height * 96 / 72)
    return round((sheet.sheet_format.defaultRowHeight or 15) * 96 / 72)


def _show_text(text: str) -> str:
    match = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})(?:T00:00:00)?", text)  # 日期：2022-01-05T00:00:00 → 2022/1/5
    return f"{match[1]}/{int(match[2])}/{int(match[3])}" if match else text


def _side_view(data: dict[str, Any], name: str, rows: list[int], last_col: int) -> dict[str, Any]:
    from openpyxl.utils.cell import get_column_letter
    sheet, texts = data["book"][name], data["texts"].get(name, {"cells": {}})["cells"]
    spans, covered = {}, set()
    for rng in sheet.merged_cells.ranges:
        spans[(rng.min_row, rng.min_col)] = (rng.max_row - rng.min_row + 1, rng.max_col - rng.min_col + 1)
        if rng.max_row - rng.min_row < 2000 and rng.max_col - rng.min_col < 200:
            covered.update((r, c) for r in range(rng.min_row, rng.max_row + 1) for c in range(rng.min_col, rng.max_col + 1)
                           if (r, c) != (rng.min_row, rng.min_col))
    styles: list[list[Any]] = []
    index: dict[int, int] = {}
    cells = []
    for row in rows:
        for col in range(1, last_col + 1):
            if (row, col) in covered:
                continue
            cell = sheet._cells.get((row, col))
            text = texts.get(f"{get_column_letter(col)}{row}", "")
            if not text and (cell is None or not cell.has_style):
                continue
            css = _cell_css(cell, data) if cell is not None else ["", 0, "", "bottom"]
            if not text and not css[0]:
                continue
            at = index.get(id(css))
            if at is None:
                at = index[id(css)] = len(styles)
                styles.append(css)
            span = spans.get((row, col), (1, 1))
            numeric = int(cell is not None and cell.data_type == "n" and bool(text))
            cells.append([row, col, _show_text(text), at, span[0], span[1], numeric])
    return {"cells": cells, "styles": styles, "col_px": _col_px(sheet, last_col)[1:], "row_px": [_row_px(sheet, r) for r in rows]}


def _used_bounds(data: dict[str, Any], name: str) -> tuple[int, int]:
    rows = cols = 0
    for addr in data["texts"].get(name, {"cells": {}})["cells"]:
        row, _, letters = _addr_key(addr)
        rows = max(rows, row)
        col = 0
        for ch in letters:
            col = col * 26 + ord(ch) - 64
        cols = max(cols, col)
    return rows, cols


def _view(name: str, old: dict[str, Any] | None, new: dict[str, Any] | None, marks: dict[str, list[str]],
          row_marks: list[int], col_marks: list[str], merge_marks: list[str],
          boxes: dict[str, list] | None = None) -> dict[str, Any] | None:
    """效果图要的数据：只取有不同的行和上下 VIEW_CONTEXT 行；一边没有这张表时画另一边的开头几十行"""
    from openpyxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries
    bounds = [_used_bounds(side, name) for side in (old, new) if side is not None]
    last_row, last_col = max(b[0] for b in bounds), max(b[1] for b in bounds)
    hot: set[int] = set(row_marks)
    for addr in marks:
        row, _, letters = _addr_key(addr)
        hot.add(row)
        last_row, last_col = max(last_row, row), max(last_col, column_index_from_string(letters))
    for rng in merge_marks:
        left, top, right, bottom = range_boundaries(rng)
        hot.update(range(top, min(bottom, top + 50) + 1))
    boxes = boxes or {"old": [], "new": []}
    for top, left, bottom, right, *_ in boxes["old"] + boxes["new"]:  # 图片 / 图形不一样：把它占的那几行也画出来
        hot.update(range(top, min(bottom, top + VIEW_BOX_ROWS) + 1))
        last_row, last_col = max(last_row, min(bottom, top + VIEW_BOX_ROWS)), max(last_col, right)
    if old is None or new is None:
        hot = set(range(1, min(last_row, 40) + 1))
    if not hot and col_marks:
        hot = set(range(1, min(last_row, 10) + 1))  # 只有列宽不同：画开头几行看个大概
    if not hot or not last_row or not last_col:
        return None
    context = 0 if old is None or new is None else VIEW_CONTEXT
    wanted = sorted({r for h in hot for r in range(max(1, h - context), min(max(last_row, h), h + context) + 1)})
    rows, cut = wanted[:VIEW_MAX_ROWS], max(0, len(wanted) - VIEW_MAX_ROWS)
    cols_cut = max(0, last_col - VIEW_MAX_COLS)
    last_col = min(last_col, VIEW_MAX_COLS)
    sides = {key: (_side_view(side, name, rows, last_col) if side is not None else None) for key, side in (("old", old), ("new", new))}
    return {"rows": rows, "cols": [get_column_letter(c) for c in range(1, last_col + 1)], **sides,
            "marks": marks if len(marks) <= 5000 else dict(list(marks.items())[:5000]), "marks_cut": max(0, len(marks) - 5000),
            "row_marks": row_marks, "col_marks": col_marks, "boxes": boxes, "rows_cut": cut, "cols_cut": cols_cut, "last_row": last_row}


def _addr_key(addr: str) -> tuple[int, int, str]:
    letters = re.sub(r"\d", "", addr)
    return int(re.sub(r"\D", "", addr) or 0), len(letters), letters


def _formula(cell: Any) -> str:
    value = cell.value if cell is not None else None
    if getattr(cell, "data_type", "") == "f":
        return str(getattr(value, "text", value) or "")
    return ""


def _comment(cell: Any) -> str:
    comment = getattr(cell, "comment", None) if cell is not None else None
    return str(comment.text or "") if comment is not None else ""


def _hyperlink(cell: Any) -> str:
    link = getattr(cell, "hyperlink", None) if cell is not None else None
    if link is None:
        return ""
    return str(link.target or (f"#{link.location}" if link.location else "") or "")


def _diff_sheet(name: str, old: dict[str, Any], new: dict[str, Any]) -> dict[str, Any]:
    """old = 远程（左），new = 要上传的（右）"""
    a, b = _sheet_info(old, name), _sheet_info(new, name)
    counts: dict[str, int] = {}
    # ── 文字 ──
    ta, tb = a["texts"], b["texts"]
    changed = sorted((k for k in ta.keys() | tb.keys() if ta.get(k, "") != tb.get(k, "")), key=_addr_key)
    # ── 公式、样式 ──
    ca, cb = a["sheet"]._cells, b["sheet"]._cells
    from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter
    text_rows, cache_rows, marks = [], [], {}
    for addr in changed:
        key = coordinate_to_tuple(addr)
        va, vb = ta.get(addr, ""), tb.get(addr, "")
        formula = _formula(ca.get(key))
        if formula and formula == _formula(cb.get(key)):
            cache_rows.append({"cell": addr, "old": va or "缓存结果为空或缺失", "new": vb or "缓存结果为空或缺失"})
            marks[addr] = ["公式缓存"]
        else:
            text_rows.append({"cell": addr, "old": va, "new": vb,
                              "kind": "add" if addr not in ta else "del" if addr not in tb else "mod"})
            marks[addr] = ["文字"]
    counts["文字"], counts["公式缓存"] = len(text_rows), len(cache_rows)
    blank = {kind: None for kind in STYLE_KINDS}
    formulas, groups = [], {kind: {} for kind in STYLE_KINDS}
    notes: dict[str, list] = {"批注": [], "超链接": []}
    for key in sorted(ca.keys() | cb.keys()):
        cell_a, cell_b = ca.get(key), cb.get(key)
        addr = f"{get_column_letter(key[1])}{key[0]}"
        fa, fb = _formula(cell_a), _formula(cell_b)
        if fa != fb:
            formulas.append({"cell": addr, "old": fa, "new": fb, "kind": "add" if not fa else "del" if not fb else "mod"})
            marks.setdefault(addr, []).append("公式")
        for kind, read in (("批注", _comment), ("超链接", _hyperlink)):
            na, nb = read(cell_a), read(cell_b)
            if na != nb:
                notes[kind].append({"cell": addr, "old": na, "new": nb, "kind": kind})
                marks.setdefault(addr, []).append(kind)
        if cell_a is None or cell_b is None:
            # 一边没有这个格子：只有带样式的才算样式不同（没字没样式的空格子不报）
            have = cell_a if cell_a is not None else cell_b
            if not have.has_style:
                continue
            sa = _style_of(cell_a, old) if cell_a is not None else blank
            sb = _style_of(cell_b, new) if cell_b is not None else blank
        else:
            sa, sb = _style_of(cell_a, old), _style_of(cell_b, new)
        for kind in STYLE_KINDS:
            va, vb = sa[kind], sb[kind]
            if va is None or vb is None:  # 一边没有格子：对方是默认样式（空）就不算
                va, vb = va or "", vb or ""
                if kind == "字体":  # 字体名每个工作簿默认不同，没法和「没有格子」比
                    continue
            if va != vb:
                groups[kind].setdefault((va, vb), []).append(addr)
                marks.setdefault(addr, []).append(kind)
    counts["公式"] = len(formulas)
    counts["批注"], counts["超链接"] = len(notes["批注"]), len(notes["超链接"])
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
    size_marks: dict[str, list] = {}
    for kind, field, label in (("行高", "rows", "第 {} 行"), ("列宽", "cols", "{} 列")):
        da, db = a[field], b[field]
        keys = [k for k in da.keys() | db.keys() if da.get(k, "默认") != db.get(k, "默认")]
        keys.sort(key=(lambda k: k) if field == "rows" else (lambda k: (len(k), k)))
        counts[kind] = len(keys)
        size_marks[field] = list(keys)
        merged: dict[tuple[str, str], list] = {}
        for k in keys:
            merged.setdefault((da.get(k, "默认"), db.get(k, "默认")), []).append(k)
        for (va, vb), items in sorted(merged.items(), key=lambda item: -len(item[1]))[:MAX_STYLE_GROUPS]:
            shown = "、".join(label.format(i) for i in items[:8]) + (f" 等 {len(items)} 处" if len(items) > 8 else "")
            whole.append({"kind": kind, "what": shown, "old": va, "new": vb})
    for kind, field, what in (("行高", "default_row", "默认行高"), ("列宽", "default_col", "默认列宽"),
                              ("列宽", "base_col", "基础列宽")):
        if a[field] != b[field]:
            counts[kind] += 1
            whole.append({"kind": kind, "what": what, "old": a[field], "new": b[field]})
    only_a = [m for m in a["merged"] if m not in set(b["merged"])]
    only_b = [m for m in b["merged"] if m not in set(a["merged"])]
    counts["合并单元格"] = len(only_a) + len(only_b)
    if only_a or only_b:
        short = lambda items: "、".join(items[:10]) + (f" 等 {len(items)} 处" if len(items) > 10 else "")  # noqa: E731
        whole.append({"kind": "合并单元格", "what": "合并的范围", "old": short(only_a), "new": short(only_b)})
    freeze_a, freeze_b = a["sheet"].freeze_panes or "", b["sheet"].freeze_panes or ""
    counts["冻结窗格"] = int(freeze_a != freeze_b)
    if freeze_a != freeze_b:
        whole.append({"kind": "冻结窗格", "what": "冻结到", "old": freeze_a, "new": freeze_b})
    setup = [(k, a["setup"][k], b["setup"][k]) for k in a["setup"] if a["setup"][k] != b["setup"][k]
             and not (k == "纸张" and not (a["setup"][k] and b["setup"][k]))]  # 一边没写纸张 = 跟打印机默认，不算不同
    counts["打印设置"] = len(setup)
    whole += [{"kind": "打印设置", "what": k, "old": va, "new": vb} for k, va, vb in setup]
    drawing_rows = _diff_drawings(a["drawings"], b["drawings"])
    counts["图片图形"] = len(drawing_rows)
    boxes: dict[str, list] = {"old": [], "new": []}
    for row in drawing_rows:
        for side, item in (("old", row["a"]), ("new", row["b"])):
            if item and item["from"]:
                end = item["to"] or item["from"]
                boxes[side].append([item["from"][0], item["from"][1], max(end[0], item["from"][0]), max(end[1], item["from"][1]),
                                    f"{_drawing_label(item)}：{'、'.join(row['changes'])}{'不一样' if row['a'] and row['b'] else ''}",
                                    # 只记了左上角的（不随格子变大小）：按它自己的宽高画框
                                    [round(v / 2.54 * 96) for v in item["size"]] if item["size"] and not item["to"] else None])
    whole += [{"kind": "图片图形", "what": row["what"], "old": row["old"], "new": row["new"]} for row in drawing_rows[:MAX_DRAWING_ROWS]]
    if len(drawing_rows) > MAX_DRAWING_ROWS:
        whole.append({"kind": "图片图形", "what": f"还有 {len(drawing_rows) - MAX_DRAWING_ROWS} 个图片 / 图形不一样，太多没列", "old": "", "new": ""})
    for kind, field, what, fmt in (("标签颜色", "tab", "标签颜色", lambda v: v or "没有"),
                                   ("显示隐藏", "state", "这张表", str)):
        counts[kind] = int(a[field] != b[field])
        if a[field] != b[field]:
            whole.append({"kind": kind, "what": what, "old": fmt(a[field]), "new": fmt(b[field])})
    ranges = [r for r in sorted(a["cf"].keys() | b["cf"].keys()) if a["cf"].get(r) != b["cf"].get(r)]
    counts["条件格式"] = len(ranges)

    def describe(rules: list[tuple[bytes, Any, str]]) -> str:
        labels = "；".join(rule[2] for rule in rules[:3])
        return f"{len(rules)} 条" + (f"（{labels}）" if labels else "")

    for region in ranges[:MAX_STYLE_GROUPS]:
        whole.append({"kind": "条件格式", "what": f"{region} 的规则有变化",
                      "old": describe(a["cf"].get(region, [])), "new": describe(b["cf"].get(region, []))})
    counts = {k: v for k, v in counts.items() if v}
    view = _view(name, old, new, marks, size_marks.get("rows", []), size_marks.get("cols", []), only_a + only_b, boxes) if counts else None
    return {"name": name, "status": "changed" if counts else "same", "counts": counts, "view": view,
            "text": text_rows[:MAX_TEXT_ROWS], "text_more": max(0, len(text_rows) - MAX_TEXT_ROWS),
            "cache": cache_rows[:MAX_TEXT_ROWS], "cache_more": max(0, len(cache_rows) - MAX_TEXT_ROWS),
            "formulas": formulas[:MAX_TEXT_ROWS], "formulas_more": max(0, len(formulas) - MAX_TEXT_ROWS),
            "notes": (notes["批注"] + notes["超链接"])[:MAX_TEXT_ROWS],
            "styles": styles, "whole": whole}


def diff(mine: Path | str, theirs: Path | str) -> dict[str, Any]:
    """mine = 你要上传的这份（右 / 新），theirs = 远程现在这份（左 / 旧）"""
    mine, theirs = Path(mine), Path(theirs)
    with debug_log.side_task("excel_diff", {"mine": str(mine), "theirs": str(theirs)}):
        new = _load(mine)
        old = None
        try:
            old = _load(theirs)
            names_new = [s.title for s in new["book"].worksheets]
            names_old = [s.title for s in old["book"].worksheets]
            sheets = []
            for name in names_old + [n for n in names_new if n not in names_old]:
                doomed = old["doomed"].get(name) or new["doomed"].get(name) or ""
                if name not in names_new:
                    sheets.append({"name": name, "status": "only_old", "doomed": doomed, "counts": {},
                                   "view": None if doomed else _view(name, old, None, {}, [], [], [])})
                elif name not in names_old:
                    sheets.append({"name": name, "status": "only_new", "doomed": doomed, "counts": {},
                                   "view": None if doomed else _view(name, None, new, {}, [], [], [])})
                else:
                    entry = _diff_sheet(name, old, new)
                    entry["doomed"] = doomed
                    sheets.append(entry)
            order_changed = [n for n in names_old if n in names_new] != [n for n in names_new if n in names_old]
        finally:
            new["book"].close()
            if old is not None:
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
