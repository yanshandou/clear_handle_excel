"""打印时内容特别少的页（用户 2026-11-03 提）：导出 PDF 后量每页印了多少，只修有把握的，别的标黄让人看。

量法：PDF 每页在上下边距之间（页眉页脚不算）所有文字、表格线、底色、图片的最上到最下，占可打印高度的比例。
设了「每页重复打印标题行」时，第 2 页起顶上的标题行不算（只在缩放是百分比时能算出标题行多高）。

判断：
- 中间页：用了不到 15%，而前一页或后一页用了 60% 以上（别的页都挺满，就这一页突然空了）。
  这一页底下是手动分页 → 把这张表的手动分页全部清掉、改成自动分页（用户 2026-11-08 定：一处处删会留下
  别的手动分页，排得不如全清），重新导出核对：页数没变多、这样的空页变少、没有新切开有字的合并格，
  才算修好，否则全部改回去标黄；是 Excel 自动分页 → 不改，标黄。向下取整不自动做，留给「调整页面」的按钮。
- 最后几页：PDF 里一个字都没有（Excel 只会因为空格子带边框、底色才印出这种页），Excel 里这几行也没有字、
  没有图片图形 → 打印区域收到上一页为止，重新导出核对页数变少才算修好；有图形的标黄。
  最后一页有字但只用了不到 15%、前一页 60% 以上 → 标黄。
只看一块打印区域、没有竖向分页（不按列拆页）、页数和分页对得上的表；设了「缩放到 N 页高」的表 Excel 不管手动分页，不看。
"""
from __future__ import annotations

import math
import statistics
import tempfile
from pathlib import Path
from typing import Any, Callable

from core import debug_log, page_breaks
from core.common import col_index_to_letters
from core.pdf_detector import fitz

FILL_LOW = 0.15   # 一页用了不到这么多算「很少」
FILL_FULL = 0.60  # 相邻页用了这么多以上算「挺满」
SCAN_TAIL_FILL = 0.40  # 扫描时自动分页的末页不到 40%，缩放到少一页（2026-11-07 起扫描不再用）
FLOOR_HINT_FILL = 0.20  # 调整页面：末页印了不到 20% 才标黄建议向下取整（用户 2026-11-08 定，38% 已经够了）
FLOOR_MIN_SCALE = 0.85  # 向下取整：字缩到不到原来的这么多就不改（太小看不清）
MAX_FIXES = 6     # 一张表最多连着修几次（每修一次都要重新导出核对）
SCAN_ROWS_LIMIT = 400  # Excel 里逐行找字最多找这么多行（Find 用不了时）
XL_VALUES, XL_PART, XL_BY_ROWS, XL_PREVIOUS = -4163, 2, 1, 2
XL_A1 = 1


def _pct(fill: float) -> int:
    return max(0, round(fill * 100))


# ---------- 量 PDF ----------

def _page_boxes(page: Any, top: float, bottom: float) -> tuple[list[tuple[float, float]], bool, bool]:
    """这一页上下边距之间的内容：[(上, 下)]，有没有字，有没有图片"""
    boxes: list[tuple[float, float]] = []
    has_text = has_image = False

    def inside(y0: float, y1: float) -> bool:
        mid = (y0 + y1) / 2
        return top - 1 <= mid <= bottom + 1

    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") == 1:
            y0, y1 = block["bbox"][1], block["bbox"][3]
            if inside(y0, y1):
                boxes.append((y0, y1))
                has_image = True
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                if not str(span.get("text", "")).strip():
                    continue
                y0, y1 = span["bbox"][1], span["bbox"][3]
                if inside(y0, y1):
                    boxes.append((y0, y1))
                    has_text = True
    for drawing in page.get_drawings():
        fill, color = drawing.get("fill"), drawing.get("color")
        if color is None and (fill is None or tuple(round(v, 3) for v in fill) == (1, 1, 1)):
            continue  # 白底不算内容
        rect = drawing.get("rect")
        if rect is None:
            continue
        y0, y1 = float(rect.y0), float(rect.y1)
        if inside(y0, y1):
            boxes.append((y0, y1))
    for info in page.get_images(full=True):
        try:
            for rect in page.get_image_rects(info[0]):
                if inside(rect.y0, rect.y1):
                    boxes.append((float(rect.y0), float(rect.y1)))
                    has_image = True
        except Exception:
            pass
    return boxes, has_text, has_image


def measure(pdf_path: Path | str, top_margin: float, bottom_margin: float, title_height: float = 0.0) -> list[dict]:
    """每页：{"fill": 用了可打印高度的几成, "text": 有没有字, "image": 有没有图片}"""
    pages = []
    with fitz.open(str(pdf_path)) as doc:
        for index, page in enumerate(doc):
            height = float(page.rect.height)
            top, bottom = top_margin, height - bottom_margin
            if bottom - top < height * .3:  # 边距读出来不对劲：整页都算
                top, bottom = 0.0, height
            body_top = top + (title_height if index > 0 else 0.0)
            boxes, has_text, has_image = _page_boxes(page, top, bottom)
            if index > 0 and title_height:
                boxes = [(y0, y1) for y0, y1 in boxes if y0 >= body_top - 1]
            fill = 0.0
            if boxes:
                fill = (max(y1 for _, y1 in boxes) - min(y0 for y0, _ in boxes)) / max(1.0, bottom - body_top)
            pages.append({"fill": min(1.0, max(0.0, fill)), "text": has_text, "image": has_image})
    return pages


# ---------- 读 Excel ----------

def _margins(sheet: Any) -> tuple[float, float]:
    try:
        setup = sheet.PageSetup
        return float(setup.TopMargin or 0), float(setup.BottomMargin or 0)
    except Exception:
        return 0.0, 0.0


def _title_height(sheet: Any) -> float:
    """每页重复打印的标题行在 PDF 里多高（点）；缩放不是百分比（缩放到 N 页）时算不出，返回 0"""
    try:
        setup = sheet.PageSetup
        rows = str(setup.PrintTitleRows or "").replace("$", "")
        zoom = setup.Zoom
        if not rows or zoom is False or zoom is None:
            return 0.0
        first, _, last = rows.partition(":")
        first_row, last_row = int(first), int(last or first)
        total = 0.0
        for row in range(first_row, last_row + 1):
            r = sheet.Rows(row)
            if not r.Hidden:
                total += float(r.RowHeight or 0)
        return total * float(zoom) / 100
    except Exception:
        return 0.0


def _area(sheet: Any) -> tuple[str, int, int, int, int] | None:
    """(原来的打印区域文字, 上, 左, 下, 右)；设了几块打印区域返回 None"""
    try:
        text = str(sheet.PageSetup.PrintArea or "")
        if "," in text:
            return None
        rng = sheet.Range(text) if text else sheet.UsedRange
        top, left = int(rng.Row), int(rng.Column)
        return text, top, left, top + int(rng.Rows.Count) - 1, left + int(rng.Columns.Count) - 1
    except Exception:
        return None


def _row_hidden(sheet: Any, row: int) -> bool:
    try:
        r = sheet.Rows(row)
        return bool(r.Hidden) or float(r.RowHeight or 0) == 0
    except Exception:
        return False


def _rows_have_text(sheet: Any, start: int, bottom: int, left: int, right: int) -> bool | None:
    """这几行（打印列内）有没有字；读不了返回 None"""
    try:
        rng = sheet.Range(sheet.Cells(start, left), sheet.Cells(bottom, right))
        found = rng.Find(What="*", After=rng.Cells(1, 1), LookIn=XL_VALUES, LookAt=XL_PART,
                         SearchOrder=XL_BY_ROWS, SearchDirection=XL_PREVIOUS)
        if found is None:
            return False
        if isinstance(getattr(found, "Row", None), int):
            return True
    except Exception:
        pass
    if bottom - start + 1 > SCAN_ROWS_LIMIT:
        return None
    for row in range(start, bottom + 1):
        if not _row_hidden(sheet, row) and page_breaks._row_has_text(sheet, row, left, right):
            return True
    return False


def _shapes_below(sheet: Any, start: int) -> bool | None:
    """第 start 行起有没有图片、图形；读不了返回 None"""
    try:
        shapes = sheet.Shapes
        for i in range(1, int(shapes.Count) + 1):
            if int(shapes.Item(i).BottomRightCell.Row) >= start:
                return True
        return False
    except Exception:
        return None


def _layout(sheet: Any) -> dict | None:
    """打印区域、分页；看不了的表返回 None（原因写调试日志）"""
    name = getattr(sheet, "Name", "")
    if page_breaks.fit_to_pages_tall(sheet):
        debug_log.event("📄 空页检查跳过", 表=name, 原因="缩放到 N 页高")
        return None
    area = _area(sheet)
    if area is None:
        debug_log.event("📄 空页检查跳过", 表=name, 原因="几块打印区域或读不到")
        return None
    text, top, left, bottom, right = area
    breaks = page_breaks.read_breaks(sheet)
    if breaks is None:
        return None
    h = [(row, manual) for row, manual in breaks["h"] if top < row <= bottom]
    v = [col for col in breaks["v"] if left < col <= right]
    if v:
        debug_log.event("📄 空页检查跳过", 表=name, 原因="按列拆页")
        return None
    return {"area": text, "top": top, "left": left, "bottom": bottom, "right": right, "h": h,
            "breaks": breaks}


# Excel XlPaperSize（常用纸型）：毫米和英寸分列，换算后单位为 pt。未知/自定义纸型不猜。
_PAPER_MM = {8: (297, 420), 9: (210, 297), 10: (210, 297), 11: (148, 210),
             12: (250, 354), 13: (182, 257), 66: (420, 594), 70: (210, 297)}
_PAPER_INCH = {1: (8.5, 11), 2: (8.5, 11), 3: (11, 17), 4: (17, 11),
               5: (8.5, 14), 6: (5.5, 8.5), 7: (7.25, 10.5), 17: (11, 17)}


def compact_scan_tail(sheet: Any) -> dict | None:
    """导出前从 Excel 估算末页占用，不生成试算 PDF，不操作分页线。

    只处理单块打印区域、纵向自动分页、未指定页高的表。由纸型/边距/打印缩放和
    末页实际行高估算正文占用；重复标题占用从可用高度中扣除。不足 40% 时只设置
    打印页宽 1、页高 N-1。后续只导出一次 PDF，按这个最终结果检测、标色。
    """
    try:
        layout = _layout(sheet)
        if layout is None or not layout["h"] or any(manual for _, manual in layout["h"]):
            return None
        setup = sheet.PageSetup
        paper, orientation = int(setup.PaperSize), int(setup.Orientation)
        if paper in _PAPER_MM:
            width, height = (v * 72 / 25.4 for v in _PAPER_MM[paper])
        elif paper in _PAPER_INCH:
            width, height = (v * 72 for v in _PAPER_INCH[paper])
        else:
            raise ValueError(f"不支持的纸型 {paper}")
        if orientation == 2:
            width, height = height, width
        elif orientation != 1:
            raise ValueError("读不到纸张方向")
        if setup.PrintHeadings is not False:
            raise ValueError("打印行列标号或读取失败")
        body_height = height - float(setup.TopMargin) - float(setup.BottomMargin)
        body_width = width - float(setup.LeftMargin) - float(setup.RightMargin)
        if body_height <= 0 or body_width <= 0:
            raise ValueError("打印边距无效")
        zoom = setup.Zoom
        if zoom is False:
            if setup.FitToPagesWide != 1:
                raise ValueError("自动缩放没有指定一页宽，无法估算比例")
            area = sheet.Range(sheet.Cells(layout["top"], layout["left"]),
                               sheet.Cells(layout["bottom"], layout["right"]))
            scale = min(1.0, body_width / float(area.Width))
        elif type(zoom) in (int, float) and 10 <= zoom <= 400:
            scale = float(zoom) / 100
        else:
            raise ValueError("读不到打印缩放")
        title_height = 0.0
        title_rows = str(setup.PrintTitleRows or "")
        if title_rows:
            title_height = float(sheet.Range(title_rows).Height)
        available = body_height - title_height * scale
        if available <= 0:
            raise ValueError("重复标题占满了页面")
        last_start = layout["h"][-1][0]
        last = sheet.Range(sheet.Cells(last_start, layout["left"]),
                           sheet.Cells(layout["bottom"], layout["right"]))
        last_height = float(last.Height)
        if not all(math.isfinite(v) and v >= 0 for v in
                   (last_height, scale, available, title_height)):
            raise ValueError("页面高度或缩放无效")
        fill = last_height * scale / available
        if fill >= SCAN_TAIL_FILL:
            return None
        original = {key: getattr(setup, key) for key in ("Zoom", "FitToPagesWide", "FitToPagesTall")}
        count = len(layout["h"]) + 1
        target = {"Zoom": False, "FitToPagesWide": 1, "FitToPagesTall": count - 1}
    except Exception as err:
        debug_log.event("📄 末页缩放跳过", 表=getattr(sheet, "Name", ""), 原因=str(err))
        return None
    # 设置失败必须中止本文件，不能按未保存或部分生效的打印设置标色。
    for key, value in target.items():
        setattr(setup, key, value)
    debug_log.event("📄 Excel 末页估算", 表=sheet.Name, 原页数=count, 目标页数=count - 1,
                    末页占比=round(fill, 4), 缩放比例=round(scale, 4))
    return {"before": original, "after": target, "pages_before": count,
            "pages_after": count - 1, "tail_fill": fill}


# ---------- 检查 ----------

def _quick_suspects(pages: list[dict]) -> bool:
    """只看 PDF 就能排除的不用再问 Excel"""
    n = len(pages)
    if n < 2:
        return False
    if not pages[-1]["text"] or pages[-1]["fill"] < FILL_LOW:
        return True
    return any(pages[i]["fill"] < FILL_LOW and max(pages[j]["fill"] for j in (i - 1, i + 1) if 0 <= j < n) > FILL_FULL
               for i in range(n - 1))


def inspect(sheet: Any, pdf_path: Path | str) -> dict | None:
    """导出好的这张表的 PDF 有没有内容太少的页。没问题、看不了都返回 None。
    返回 {"pages": 页数, "layout": ..., "items": [{"kind": "break" | "tail" | "flag", "page", ...}]}"""
    top_margin, bottom_margin = _margins(sheet)
    try:
        pages = measure(pdf_path, top_margin, bottom_margin, _title_height(sheet))
    except Exception as err:
        debug_log.event("📄 空页检查跳过", 表=getattr(sheet, "Name", ""), 原因=f"PDF 读不了：{err}")
        return None
    if not _quick_suspects(pages):
        return None
    layout = _layout(sheet)
    if layout is None:
        return None
    n, h = len(pages), layout["h"]
    if n != len(h) + 1:
        debug_log.event("📄 空页检查跳过", 表=getattr(sheet, "Name", ""), 原因="页数和分页对不上", 页数=n, 分页=len(h))
        return None
    items: list[dict] = []
    starts = [layout["top"]] + [row for row, _ in h]

    # 最后几页一个字都没有
    blank_from = n
    while blank_from > 1 and not pages[blank_from - 1]["text"] and not pages[blank_from - 1]["image"]:
        blank_from -= 1
    if blank_from < n:
        start, bottom = starts[blank_from], layout["bottom"]
        where = f"第 {blank_from + 1} 页起（第 {start}～{bottom} 行）" if blank_from < n - 1 else f"最后一页（第 {start}～{bottom} 行）"
        has_text = _rows_have_text(sheet, start, bottom, layout["left"], layout["right"])
        shapes = _shapes_below(sheet, start)
        if has_text is False and shapes is False:
            items.append({"kind": "tail", "page": blank_from, "start": start, "bottom": bottom, "where": where})
        elif has_text is False and shapes:
            items.append({"kind": "flag", "page": blank_from,
                          "text": f"{where}没有文字，但有图片或图形；没有自动去掉，请人工判断"})
        elif has_text is None or shapes is None:
            items.append({"kind": "flag", "page": blank_from,
                          "text": f"{where}在 PDF 里没有文字，但 Excel 里读不清这几行；没有自动去掉，请人工判断"})
    elif pages[-1]["fill"] < FILL_LOW and pages[-2]["fill"] > FILL_FULL:
        items.append({"kind": "flag", "page": n - 1,
                      "text": f"最后一页只印了约 {_pct(pages[-1]['fill'])}%（第 {starts[-1]} 行起），内容很少；"
                              "没有自动改，要不要缩一点挤回上一页请人工判断"})

    for i in range(min(n - 1, blank_from)):
        neighbors = [pages[j]["fill"] for j in (i - 1, i + 1) if 0 <= j < blank_from]
        if pages[i]["fill"] >= FILL_LOW or not neighbors or max(neighbors) <= FILL_FULL:
            continue
        row, manual = h[i]
        rows = f"第 {starts[i]}～{row - 1} 行"
        if manual:
            items.append({"kind": "break", "page": i, "row": row, "fill": pages[i]["fill"], "rows": rows})
        else:
            items.append({"kind": "flag", "page": i,
                          "text": f"第 {i + 1} 页只印了约 {_pct(pages[i]['fill'])}%（{rows}），前后页都很满；"
                                  "这里是 Excel 自动分页（多半是下面一块太高放不下），没有自动改，请人工调整"})
    if not items:
        return None
    return {"pages": n, "layout": layout, "items": sorted(items, key=lambda it: it["page"])}


def _sparse_count(found: dict | None) -> int:
    if not found:
        return 0
    return sum(1 for it in found["items"] if it["kind"] == "break"
               or (it["kind"] == "flag" and "前后页都很满" in it["text"]))


def _tail_count(found: dict | None) -> int:
    return sum(1 for it in (found or {}).get("items", []) if it["kind"] == "tail")


def _cuts(sheet: Any, found: dict | None) -> set[str] | None:
    breaks = found["layout"]["breaks"] if found else page_breaks.read_breaks(sheet)
    if breaks is None:
        return None
    return {b["addr"] for b in page_breaks.find_cut_blocks(sheet, breaks)}


# ---------- 修 ----------

def _set_area(sheet: Any, layout: dict, bottom: int) -> None:
    rng = sheet.Range(sheet.Cells(layout["top"], layout["left"]), sheet.Cells(bottom, layout["right"]))
    try:
        style = sheet.Application.ReferenceStyle
    except Exception:
        style = XL_A1
    sheet.PageSetup.PrintArea = rng.Address(True, True, style)


def _area_text(layout: dict, bottom: int) -> str:
    return (f"{col_index_to_letters(layout['left'])}{layout['top']}:"
            f"{col_index_to_letters(layout['right'])}{bottom}")


def _clear_manual_breaks(sheet: Any, pdf_path: Path, export: Callable[[Any, Path], None],
                         current: dict) -> tuple[dict | None, str, str, bool]:
    """有手动分页造成的空页：把打印区域里的手动分页全部清掉，改成自动分页，重新导出核对。
    返回 (清完/改回后的检查结果, 修好的说明 或 "", 没修好的原因 或 "", 有没有改回不去的)"""
    name = getattr(sheet, "Name", "")
    layout, pages = current["layout"], current["pages"]
    manual = [row for row, is_manual in layout["h"] if is_manual]
    sparse = [it for it in current["items"] if it["kind"] == "break"]
    before_cuts = _cuts(sheet, current)
    cleared: list[int] = []
    reason = ""
    dirty = False
    try:
        for row in manual:
            sheet.Rows(row).PageBreak = page_breaks.XL_PAGEBREAK_NONE
            cleared.append(row)
    except Exception as err:
        reason = f"Excel 没让清手动分页（{err}）"
    after = None
    after_pages = pages
    if not reason:
        export(sheet, pdf_path)
        after = inspect(sheet, pdf_path)
        after_pages = after["pages"] if after else _page_count(pdf_path)
        after_cuts = _cuts(sheet, after)
        if after_pages is None:
            reason = "清完 PDF 读不了"
        elif after_pages > pages:
            reason = "清掉全部手动分页后页数反而变多"
        elif _sparse_count(after) >= _sparse_count(current):
            reason = "清掉全部手动分页后还是有内容很少的页"
        elif before_cuts is None or after_cuts is None:
            reason = "读不到分页，没法确认会不会切开合并格"
        elif after_cuts - before_cuts:
            reason = f"清掉全部手动分页会切开合并格 {'、'.join(sorted(after_cuts - before_cuts))}"
    if reason:
        for row in cleared:
            try:
                sheet.Rows(row).PageBreak = page_breaks.XL_PAGEBREAK_MANUAL
            except Exception as err:
                debug_log.event("⚠️ 手动分页改不回去", 表=name, 行=row, 原因=str(err))
                dirty = True  # 改不回去也得保存时知道 Excel 变了
        if cleared:
            export(sheet, pdf_path)
        debug_log.event("📄 清手动分页没用上", 表=name, 分页=len(manual), 原因=reason)
        return inspect(sheet, pdf_path), "", reason, dirty
    where = "；".join(f"第 {it['page'] + 1} 页只印了约 {_pct(it['fill'])}%（{it['rows']}）" for it in sparse)
    rows = "、".join(str(row) for row in manual)
    text = (f"原来{where}：清掉了全部 {len(manual)} 处手动分页（第 {rows} 行上方），改成自动分页，"
            f"{pages} 页 → {after_pages} 页")
    debug_log.event("📄 空页已修", 表=name, 说明=text)
    return after, text, "", False


def fix_sheet(sheet: Any, pdf_path: Path, export: Callable[[Any, Path], None], found: dict) -> tuple[list[dict], bool]:
    """sheet 要是可以保存的。有手动分页造成的空页 → 手动分页全部清掉再核对；最后几页全是空格子 → 一处处收打印区域。
    每次改完都重新导出核对，修不好的改回去。返回 (这张表的提示 [{"page", "fixed", "text"}], 有没有改了 Excel)"""
    name = getattr(sheet, "Name", "")
    fixed_notes: list[dict] = []
    failed: dict[tuple, str] = {}
    changed = False
    current = found
    if any(it["kind"] == "break" for it in current["items"]):
        breaks = [it for it in current["items"] if it["kind"] == "break"]
        current, text, reason, dirty = _clear_manual_breaks(sheet, pdf_path, export, current)
        changed = changed or dirty
        if text:
            changed = True
            # cleared：调整页面上这张表不再显示「清手动分页」
            fixed_notes.append({"page": None, "fixed": True, "text": text, "cleared": True})
        else:
            for it in breaks:
                failed[("break", it["row"])] = reason
    for _ in range(MAX_FIXES):
        todo = [it for it in (current or {}).get("items", []) if it["kind"] == "tail"
                and ("tail", it["start"]) not in failed]
        if not todo:
            break
        item = todo[0]
        key = ("tail", item["start"])
        layout, pages = current["layout"], current["pages"]
        try:
            _set_area(sheet, layout, item["start"] - 1)
        except Exception as err:
            failed[key] = f"Excel 没让改（{err}）"
            continue
        export(sheet, pdf_path)
        after = inspect(sheet, pdf_path)
        after_pages = after["pages"] if after else _page_count(pdf_path)
        reason = ""
        if after_pages is None:
            reason = "改完 PDF 读不了"
        elif after_pages >= pages or _tail_count(after):
            reason = "收了打印区域页数也没变少"
        if reason:
            try:
                sheet.PageSetup.PrintArea = layout["area"]
            except Exception as err:
                debug_log.event("⚠️ 空页修复改不回去", 表=name, 原因=str(err))
                changed = True  # 改不回去也得保存时知道 Excel 变了
            export(sheet, pdf_path)
            failed[key] = reason
            debug_log.event("📄 空页修复没用上", 表=name, 改法="tail", 原因=reason)
            current = inspect(sheet, pdf_path)
            continue
        changed = True
        text = (f"{item['where']}全是没有字的空格子：打印区域收到 {_area_text(layout, item['start'] - 1)}，"
                f"{pages} 页 → {after_pages} 页")
        fixed_notes.append({"page": None, "fixed": True, "text": text})
        debug_log.event("📄 空页已修", 表=name, 说明=text)
        current = after
    notes = fixed_notes + flags(current, failed)
    return notes, changed


def _page_count(pdf_path: Path) -> int | None:
    try:
        with fitz.open(str(pdf_path)) as doc:
            return len(doc)
    except Exception:
        return None


def flags(found: dict | None, failed: dict[tuple, str] | None = None) -> list[dict]:
    """没修（或修不好）的写成要人工看的提示；能修的没修是因为 Excel 打不开改不了"""
    failed = failed or {}
    notes = []
    for it in (found or {}).get("items", []):
        if it["kind"] == "flag":
            text = it["text"]
        elif it["kind"] == "tail":
            why = failed.get(("tail", it["start"]), "没能改 Excel")
            text = f"{it['where']}全是没有字的空格子，{why}，请人工收打印区域"
        else:
            why = failed.get(("break", it["row"]), "没能改 Excel")
            text = (f"第 {it['page'] + 1} 页只印了约 {_pct(it['fill'])}%（{it['rows']}），前后页都很满，"
                    f"原因是第 {it['row']} 行上方的手动分页；{why}，没有删，请人工调整")
        notes.append({"page": it["page"], "fixed": False, "text": text})
    return notes


# ---------- 工具箱：向下取整（用户 2026-11-07 定） ----------
# 复检时看到 PDF 是 4.1 页这种（最后一页只多一点），点按钮标记，一键处理时把这张表设成
# 「1 页宽、N-1 页高」挤回上一页。真导出一次核对：页数少了一页、字没缩到 85% 以下、
# PDF 正文里原来有的每一段字缩完都还在（没缩出 ### 或丢字），三条都满足才保留，否则改回原来的打印设置。

def _texts(pdf_path: Path, top_margin: float, bottom_margin: float) -> set[str]:
    """PDF 正文（上下边距之间，页眉页脚不算：「第几页 / 共几页」缩完会变）里出现过的每一段字"""
    texts: set[str] = set()
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            height = float(page.rect.height)
            top, bottom = top_margin, height - bottom_margin
            if bottom - top < height * .3:  # 边距读出来不对劲：整页都算
                top, bottom = 0.0, height
            for block in page.get_text("dict").get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = str(span.get("text", "")).strip()
                        mid = (span["bbox"][1] + span["bbox"][3]) / 2
                        if text and top - 1 <= mid <= bottom + 1:
                            texts.add(text)
    return texts


def _span_widths(pdf_path: Path) -> dict[str, float]:
    """只出现一次的文字块 -> 宽度（量缩放用）"""
    seen: dict[str, list[float]] = {}
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            for block in page.get_text("dict").get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = str(span.get("text", "")).strip()
                        if len(text) >= 2 and not text.isdigit():
                            seen.setdefault(text, []).append(float(span["bbox"][2] - span["bbox"][0]))
    return {text: widths[0] for text, widths in seen.items() if len(widths) == 1 and widths[0] > 0}


def _scale(before: Path, after: Path) -> float | None:
    """同一段字缩完后宽度是原来的几成（取中位数）；量不出返回 None"""
    old, new = _span_widths(before), _span_widths(after)
    ratios = [new[text] / old[text] for text in old.keys() & new.keys()]
    return statistics.median(ratios) if ratios else None


def _restore_fit(setup: Any, original: dict) -> None:
    if original["Zoom"] is False:
        setup.Zoom = False
        setup.FitToPagesWide = original["FitToPagesWide"]
        setup.FitToPagesTall = original["FitToPagesTall"]
    else:
        setup.Zoom = original["Zoom"]


def floor_pages(sheet: Any, pdf_path: Path, export: Callable[[Any, Path], None]) -> dict:
    """sheet 要是可以保存的；pdf_path 是复检页上看到的那份 PDF（N 页）。
    返回 {"ok", "reason", "before", "after", "scale", "manual"}；ok=False 时 Excel 的打印设置已改回原样"""
    name = getattr(sheet, "Name", "")
    result: dict[str, Any] = {"ok": False, "reason": "", "before": _page_count(pdf_path), "after": None,
                              "scale": None, "manual": False}
    pages = result["before"]
    if not pages:
        result["reason"] = "PDF 读不了"
        return result
    if pages < 2:
        result["reason"] = "只有 1 页，不能再少了"
        return result
    if _area(sheet) is None:
        result["reason"] = "设了几块打印区域，没法挤成一块"
        return result
    breaks = page_breaks.read_breaks(sheet)
    if breaks is None:
        result["reason"] = "读不到 Excel 的分页（可能没装打印机），没有改"
        return result
    if breaks["v"]:
        result["reason"] = "这张表打印时按列拆成了几页宽，向下取整会改成 1 页宽、变化太大，没有改"
        return result
    result["manual"] = any(manual for _, manual in breaks["h"])
    setup = sheet.PageSetup
    try:
        original = {key: getattr(setup, key) for key in ("Zoom", "FitToPagesWide", "FitToPagesTall")}
    except Exception as err:
        result["reason"] = f"读不到打印缩放（{err}）"
        return result
    try:
        setup.Zoom = False
        setup.FitToPagesWide = 1
        setup.FitToPagesTall = pages - 1
    except Exception as err:
        try:
            _restore_fit(setup, original)
        except Exception:
            pass
        result["reason"] = f"Excel 没让改打印缩放（{err}）"
        return result
    with tempfile.TemporaryDirectory(prefix="floor_") as tmp:
        trial = Path(tmp) / "trial.pdf"
        reason = ""
        try:
            export(sheet, trial)
            after = _page_count(trial)
            result["after"] = after
            if after is None:
                reason = "缩完的 PDF 读不了"
            elif after >= pages:
                reason = f"设了 {pages - 1} 页高，导出来还是 {after} 页"
            else:
                scale = _scale(pdf_path, trial)
                result["scale"] = scale
                if scale is None:
                    reason = "量不出字缩了多少（PDF 里没有可比的文字）"
                elif scale < FLOOR_MIN_SCALE:
                    reason = f"要把字缩到约 {_pct(scale)}% 才挤得进 {pages - 1} 页，太小了"
                else:
                    top_margin, bottom_margin = _margins(sheet)
                    lost = _texts(pdf_path, top_margin, bottom_margin) - _texts(trial, top_margin, bottom_margin)
                    if lost:
                        # 用户 2026-11-08 定：照样取整，缩出来的切字留给后面「查看切字」去查、去修，这里只提醒
                        sample = "」「".join(sorted(lost)[:3])
                        result["warn"] = (f"缩小后有 {len(lost)} 处文字可能显示不全（如「{sample}」），"
                                          "后面「查看切字」会查出来")
        except Exception as err:
            reason = f"缩完导出 PDF 出错（{err}）"
    if reason:
        try:
            _restore_fit(setup, original)
        except Exception as err:
            debug_log.event("⚠️ 向下取整改不回去", 表=name, 原因=str(err))
            reason += "；打印设置也没能改回去，这份 Excel 不保存"
            result["dirty"] = True
        result["reason"] = reason
        debug_log.event("📄 向下取整没用上", 表=name, 原页数=pages, 原因=reason, 原设置=original)
        return result
    result["ok"] = True
    debug_log.event("📄 向下取整", 表=name, 原页数=pages, 新页数=result["after"], 字缩到=round(result["scale"], 3),
                    原设置=original, 有手动分页=result["manual"], 可能切字=result.get("warn", ""))
    return result


# ---------- 调整页面：宽一页、高自动（用户 2026-11-07 定） ----------
# 横向不许拆成几页（99 列在第一页、1 列单独印一页是不允许的），所以导出 PDF 前每张表都设成「1 页宽」；
# 纵向先不定死，交给 Excel 自动分页（要 4 页就 4 页）。已经是 1 页宽的表不动（页高是取整定过的 N 页就留着）。
# 导出后末页只印了一点（不到 FLOOR_HINT_FILL）的，网页上自动标黄，建议向下取整。

XL_PORTRAIT, XL_LANDSCAPE = 1, 2


def _int_or_none(value: Any) -> int | None:
    if isinstance(value, bool) or value is None:
        return None
    try:
        number = int(value)
    except (TypeError, ValueError):
        return None
    return number if number > 0 and number == value else None


def one_page_wide(sheet: Any) -> dict | None:
    """sheet 要是可以保存的。不是 1 页宽的设成 Zoom=False、1 页宽、页高自动；改了返回 {"before": 原缩放说明}，
    已经是 1 页宽返回 None。读不到打印设置（例如没装打印机）也返回 None，只记调试日志"""
    name = getattr(sheet, "Name", "")
    try:
        setup = sheet.PageSetup
        zoom, wide = setup.Zoom, setup.FitToPagesWide
    except Exception as err:
        debug_log.event("📐 宽一页跳过", 表=name, 原因=f"读不到打印设置：{err}")
        return None
    if zoom is False and _int_or_none(wide) == 1:
        return None
    if zoom is False:
        tall = _int_or_none(getattr(setup, "FitToPagesTall", None))
        before = f"缩放到 {_int_or_none(wide) or '自动'} 页宽、{tall or '自动'} 页高"
    else:
        before = f"缩放 {zoom}%"
    setup.Zoom = False
    setup.FitToPagesWide = 1
    setup.FitToPagesTall = False
    debug_log.event("📐 设成宽一页、高自动", 表=name, 原来=before)
    return {"before": before}


def sheet_layout(sheet: Any) -> dict:
    """调整页面网页要显示的打印设置：方向（1 竖 / 2 横）、页高（取整定死的 N 页，自动是 None）、上下边距"""
    info: dict[str, Any] = {"orient": None, "tall": None}
    try:
        setup = sheet.PageSetup
        orient = _int_or_none(setup.Orientation)
        info["orient"] = orient if orient in (XL_PORTRAIT, XL_LANDSCAPE) else None
        if setup.Zoom is False:
            info["tall"] = _int_or_none(setup.FitToPagesTall)
    except Exception:
        pass
    # 页高自动时手动分页才起作用：记下有几处，调整页面上给「清手动分页」按钮（用户 2026-11-08 定）
    info["manual"] = manual_break_count(sheet) if info["tall"] is None else 0
    info["margins"] = _margins(sheet)
    return info


def manual_break_count(sheet: Any) -> int:
    """这张表有几处横向手动分页；读不到算 0（不给按钮，不猜）"""
    breaks = page_breaks.read_breaks(sheet)
    return sum(1 for _, manual in (breaks or {}).get("h", []) if manual)


def clear_manual_breaks(sheet: Any) -> dict:
    """调整页面的工具：清掉这张表全部手动分页，改成 Excel 自动分页。
    不在这里核对排得好不好：改完重新生成 PDF，由人看新的 PDF（末页只多一点的照常建议取整）。
    返回 {"ok", "count" 清掉几处, "reason" 没清成的原因}"""
    breaks = page_breaks.read_breaks(sheet)
    if breaks is None:
        return {"ok": False, "count": 0, "reason": "读不到这张表的分页"}
    rows = [row for row, manual in breaks["h"] if manual]
    if not rows:
        return {"ok": False, "count": 0, "reason": "这张表已经没有手动分页了"}
    try:
        sheet.ResetAllPageBreaks()
    except Exception:
        try:
            for row in rows:
                sheet.Rows(row).PageBreak = page_breaks.XL_PAGEBREAK_NONE
        except Exception as err:
            return {"ok": False, "count": 0, "reason": f"Excel 没让清手动分页（{err}）"}
    left = manual_break_count(sheet)
    if left:
        return {"ok": False, "count": len(rows) - left, "dirty": True,
                "reason": f"还剩 {left} 处手动分页清不掉，请在 Excel 里手动处理"}
    return {"ok": True, "count": len(rows), "reason": ""}


def tail_hint(pdf_path: Path | str, margins: tuple[float, float]) -> dict:
    """导出好的 PDF：几页、最后一页印了多少（0~100）、要不要建议向下取整（末页不到 20%）"""
    try:
        pages = measure(pdf_path, *margins)
    except Exception as err:
        debug_log.event("📐 末页量不了", PDF=Path(pdf_path).name, 原因=str(err))
        return {"pages": None, "tail": None, "floor": False}
    if len(pages) < 2:
        return {"pages": len(pages), "tail": None, "floor": False}
    fill = pages[-1]["fill"]
    return {"pages": len(pages), "tail": _pct(fill), "floor": fill < FLOOR_HINT_FILL}
