"""分页切开合并单元格：Excel 自动分页落在多行合并区域中间，打印出来一块内容被切成两半、分在两页。

修复办法：在合并区域第一行上方插一个手动分页，让整块挪到下一页。只改分页，不改行高、列宽和缩放；
还原时把这个手动分页去掉即可。
"""
import base64
from typing import Any

from core.common import col_index_to_letters
from core.hidden_text import collect_text_pieces, locate_text, normalize_text
from core.pdf_detector import fitz
from core import print_layout

XL_PAGEBREAK_MANUAL, XL_PAGEBREAK_NONE = -4135, -4142
XL_PAGE_BREAK_PREVIEW = 2
XL_OVER_THEN_DOWN = 2
REASON = "合并单元格被分页切开，打印时一半在上一页、一半在下一页"
FIX_REASON = "合并单元格跨页被切开（分页落在合并区域中间）"
FIT_REASON = "工作表设置了「缩放到 N 页高」，Excel 会忽略手动分页；未修改，请人工调整打印设置或标绿忽略"
TALL_REASON = "合并区域比一页还高，挪到下一页也放不下；未修改，请人工调整或标绿忽略"
MANUAL_REASON = "合并区域上方已有手动分页仍被切开；未修改，请人工核对"
STRIP = .35  # 截图取上一页底部、下一页顶部各这么多


def _read_breaks(sheet: Any) -> dict[str, list]:
    h = [(int(b.Location.Row), int(b.Type) == XL_PAGEBREAK_MANUAL)
         for b in (sheet.HPageBreaks(i) for i in range(1, int(sheet.HPageBreaks.Count) + 1))]
    v = [int(sheet.VPageBreaks(i).Location.Column) for i in range(1, int(sheet.VPageBreaks.Count) + 1)]
    return {"h": sorted(h), "v": sorted(v)}


@print_layout.with_layout
def read_breaks(sheet: Any) -> dict[str, list] | None:
    """当前的分页：{"h": [(行, 是否手动)], "v": [列]}，分页落在这一行/列的上方/左边。读不到返回 None。
    后台 Excel 不在「分页预览」时 HPageBreaks 常常没算好、读后面的分页还会报错，先切到分页预览读，读完切回去。"""
    try:
        window = None
        try:
            sheet.Parent.Activate()
            sheet.Activate()
            window = sheet.Parent.Windows(1)
            view = window.View
            window.View = XL_PAGE_BREAK_PREVIEW
        except Exception:
            window = None
        try:
            return _read_breaks(sheet)
        finally:
            if window is not None:
                try:
                    window.View = view
                except Exception:
                    pass
    except Exception:
        return None  # 没装打印机等情况读不到分页：不检测，不猜


def fit_to_pages_tall(sheet: Any) -> bool:
    """设置了「缩放到 N 页高」：Excel 打印时忽略手动分页，插分页修不好"""
    try:
        setup = sheet.PageSetup
        return setup.Zoom is False and setup.FitToPagesTall not in (False, None, 0)
    except Exception:
        return False


def _columns(sheet: Any) -> tuple[int, int] | None:
    """检查哪些列：打印区域（只有一块时）或已用区域"""
    try:
        area = str(sheet.PageSetup.PrintArea or "")
        target = sheet.Range(area) if area else sheet.UsedRange
        if area and int(target.Areas.Count) != 1:
            target = sheet.UsedRange
        first = int(target.Column)
        return first, first + int(target.Columns.Count) - 1
    except Exception:
        return None


def find_cut_blocks(sheet: Any, breaks: dict[str, list], printed: Any = None) -> list[dict]:
    """被自动分页切开、里面有文字的多行合并区域。手动分页是作者自己插的，不当作问题。
    每个自动分页只读那一行：整行都没有合并单元格（MergeCells 为 False）就直接跳过。"""
    columns = _columns(sheet)
    if columns is None:
        return []
    first, last = columns
    blocks: dict[str, dict] = {}
    for row, manual in breaks["h"]:
        if manual or row <= 1:
            continue
        try:
            if sheet.Range(sheet.Cells(row, first), sheet.Cells(row, last)).MergeCells is False:
                continue
        except Exception:
            continue
        col = first
        while col <= last:
            step = 1
            try:
                cell = sheet.Cells(row, col)
                if cell.MergeCells:
                    area = cell.MergeArea
                    top, left = int(area.Row), int(area.Column)
                    step = max(1, left + int(area.Columns.Count) - col)
                    text = str(area.Cells(1, 1).Text or "").strip()
                    if top < row and normalize_text(text) and (printed is None or printed(top, left) is not False):
                        bottom = top + int(area.Rows.Count) - 1
                        right = left + int(area.Columns.Count) - 1
                        addr = f"{col_index_to_letters(left)}{top}:{col_index_to_letters(right)}{bottom}"
                        blocks.setdefault(addr, {"addr": addr, "top": top, "bottom": bottom, "left": left,
                                                 "right": right, "break_row": row, "text": text})
            except Exception:
                pass
            col += step
    return sorted(blocks.values(), key=lambda b: (b["top"], b["left"]))


def still_cut(breaks: dict[str, list], block: dict) -> bool:
    return any(block["top"] < row <= block["bottom"] for row, _ in breaks["h"])


def page_index(sheet: Any, breaks: dict[str, list], row: int, col: int, pages: int) -> int | None:
    """这一格打印在 PDF 第几页（从 0 开始）。按分页和「先列后行 / 先行后列」顺序推算；
    页数对不上（空白页被跳过、多块打印区域等）就返回 None，不猜。"""
    n_h, n_v = len(breaks["h"]) + 1, len(breaks["v"]) + 1
    if n_h * n_v != pages:
        return None
    i = sum(r <= row for r, _ in breaks["h"])
    j = sum(c <= col for c in breaks["v"])
    try:
        across = int(sheet.PageSetup.Order) == XL_OVER_THEN_DOWN
    except Exception:
        across = False
    return i * n_v + j if across else j * n_h + i


def render_break(doc: Any, upper: int, lower: int, text: str = "", pieces: list | None = None,
                 zoom: float = 2.0) -> tuple[str, int | None, list]:
    """上一页底部 + 下一页顶部拼成一张图，中间灰条是两页的分界；能在这两页找到文字就用红框标出。
    返回 (截图, 文字所在页, 文字在那一页的坐标)，找不到文字时后两项为 None / []。"""
    a, b = doc[upper].rect, doc[lower].rect
    clip_a = fitz.Rect(a.x0, a.y1 - a.height * STRIP, a.x1, a.y1)
    clip_b = fitz.Rect(b.x0, b.y0, b.x1, b.y0 + b.height * STRIP)
    gap = 10
    width = max(clip_a.width, clip_b.width)
    out = fitz.open()
    try:
        page = out.new_page(width=width, height=clip_a.height + gap + clip_b.height)
        page.draw_rect(page.rect, color=None, fill=(1, 1, 1))
        page.show_pdf_page(fitz.Rect(0, 0, clip_a.width, clip_a.height), doc, upper, clip=clip_a)
        page.show_pdf_page(fitz.Rect(0, clip_a.height + gap, clip_b.width, clip_a.height + gap + clip_b.height),
                           doc, lower, clip=clip_b)
        page.draw_rect(fitz.Rect(0, clip_a.height, width, clip_a.height + gap), color=None, fill=(.75, .75, .75))
        found_page, found_bbox = None, []
        if len(normalize_text(text)) >= 2:
            found = locate_text(collect_text_pieces(doc) if pieces is None else pieces, text)
            if found is not None and found[0] in (upper, lower):
                found_page, found_bbox = found[0], list(found[1])
                x0, y0, x1, y1 = found_bbox
                shift = -clip_a.y0 if found_page == upper else clip_a.height + gap - clip_b.y0
                box = fitz.Rect(x0 - 2, y0 + shift - 2, x1 + 2, y1 + shift + 2) & page.rect
                if not box.is_empty:
                    page.draw_rect(box, color=(1, 0, 0), width=1.2)
        png = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom)).tobytes("png")
    finally:
        out.close()
    return "data:image/png;base64," + base64.b64encode(png).decode("ascii"), found_page, found_bbox


def shot_around(sheet: Any, breaks: dict[str, list], doc: Any, row: int, col: int, text: str,
                pieces: list | None = None) -> tuple[str, int | None, list]:
    """第 row 行上方那道分页两边的截图（row 是分页后新一页的第一行）"""
    upper = page_index(sheet, breaks, row - 1, col, len(doc))
    lower = page_index(sheet, breaks, row, col, len(doc))
    if upper is None or lower is None or upper == lower:
        return "", None, []
    return render_break(doc, upper, lower, text, pieces)
