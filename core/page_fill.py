"""打印时内容特别少的页（用户 2026-11-03 提）：导出 PDF 后量每页印了多少，只修有把握的，别的标黄让人看。

量法：PDF 每页在上下边距之间（页眉页脚不算）所有文字、表格线、底色、图片的最上到最下，占可打印高度的比例。
设了「每页重复打印标题行」时，第 2 页起顶上的标题行不算（只在缩放是百分比时能算出标题行多高）。

判断：
- 中间页：用了不到 15%，而前一页或后一页用了 60% 以上（别的页都挺满，就这一页突然空了）。
  这一页底下是手动分页 → 删掉这条分页，重新导出核对：页数没变多、这样的空页变少、没有新切开有字的合并格，
  才算修好，否则改回去标黄；是 Excel 自动分页 → 不改，标黄。
- 最后几页：PDF 里一个字都没有（Excel 只会因为空格子带边框、底色才印出这种页），Excel 里这几行也没有字、
  没有图片图形 → 打印区域收到上一页为止，重新导出核对页数变少才算修好；有图形的标黄。
  最后一页有字但只用了不到 15%、前一页 60% 以上 → 标黄。
只看一块打印区域、没有竖向分页（不按列拆页）、页数和分页对得上的表；设了「缩放到 N 页高」的表 Excel 不管手动分页，不看。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from core import debug_log, page_breaks
from core.common import col_index_to_letters
from core.pdf_detector import fitz

FILL_LOW = 0.15   # 一页用了不到这么多算「很少」
FILL_FULL = 0.60  # 相邻页用了这么多以上算「挺满」
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


def fix_sheet(sheet: Any, pdf_path: Path, export: Callable[[Any, Path], None], found: dict) -> tuple[list[dict], bool]:
    """sheet 要是可以保存的。一条一条修、每次重新导出核对；修不好的改回去。
    返回 (这张表的提示 [{"page", "fixed", "text"}], 有没有改了 Excel)"""
    name = getattr(sheet, "Name", "")
    fixed_notes: list[dict] = []
    failed: dict[tuple, str] = {}
    changed = False
    current = found
    for _ in range(MAX_FIXES):
        todo = [it for it in (current or {}).get("items", []) if it["kind"] in ("tail", "break")
                and (it["kind"], it.get("row") or it.get("start")) not in failed]
        if not todo:
            break
        item = next((it for it in todo if it["kind"] == "tail"), todo[0])
        key = (item["kind"], item.get("row") or item.get("start"))
        layout, pages = current["layout"], current["pages"]
        before_cuts = _cuts(sheet, current)
        try:
            if item["kind"] == "tail":
                _set_area(sheet, layout, item["start"] - 1)
            else:
                sheet.Rows(item["row"]).PageBreak = page_breaks.XL_PAGEBREAK_NONE
        except Exception as err:
            failed[key] = f"Excel 没让改（{err}）"
            continue
        export(sheet, pdf_path)
        after = inspect(sheet, pdf_path)
        after_pages = after["pages"] if after else _page_count(pdf_path)
        reason = ""
        if after_pages is None:
            reason = "改完 PDF 读不了"
        elif item["kind"] == "tail":
            if after_pages >= pages or _tail_count(after):
                reason = "收了打印区域页数也没变少"
        else:
            after_cuts = _cuts(sheet, after)
            if after_pages > pages:
                reason = "删掉后页数反而变多"
            elif _sparse_count(after) >= _sparse_count(current):
                reason = "删掉后还是有内容很少的页"
            elif before_cuts is None or after_cuts is None:
                reason = "读不到分页，没法确认会不会切开合并格"
            elif after_cuts - before_cuts:
                reason = f"删掉会切开合并格 {'、'.join(sorted(after_cuts - before_cuts))}"
        if reason:
            try:
                if item["kind"] == "tail":
                    sheet.PageSetup.PrintArea = layout["area"]
                else:
                    sheet.Rows(item["row"]).PageBreak = page_breaks.XL_PAGEBREAK_MANUAL
            except Exception as err:
                debug_log.event("⚠️ 空页修复改不回去", 表=name, 原因=str(err))
                changed = True  # 改不回去也得保存时知道 Excel 变了
            export(sheet, pdf_path)
            failed[key] = reason
            debug_log.event("📄 空页修复没用上", 表=name, 改法=item["kind"], 原因=reason)
            current = inspect(sheet, pdf_path)
            continue
        changed = True
        if item["kind"] == "tail":
            text = (f"{item['where']}全是没有字的空格子：打印区域收到 {_area_text(layout, item['start'] - 1)}，"
                    f"{pages} 页 → {after_pages} 页")
        else:
            text = (f"原来第 {item['page'] + 1} 页只印了约 {_pct(item['fill'])}%（{item['rows']}）："
                    f"删掉了第 {item['row']} 行上方的手动分页，{pages} 页 → {after_pages} 页")
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
