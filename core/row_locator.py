"""根据 PDF 里切字的位置，找出它在 Excel 里是哪一格。

只靠文字找格子不够：同样的文字（如「-」「非表示」、字段名）在表里会出现很多次，
以前按「PDF 第几页、页面上多高」按比例估算行号，再挑离估算最近的同文字格子，
行高不一、有打印标题、隐藏行、横向分页时估算会偏很多行，就会定位到别的格子。

这里改用「锚点」：表里只出现一次、PDF 里也只出现一次的文字，能确定它在 PDF 上的位置就是哪一格。
切字上下最近的锚点夹出行号范围，再按 Excel 行高在两个锚点之间换算出最接近的行；
同一范围里有多个同文字格子时，再用锚点的横向位置判断是哪一列。
"""
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from core.hidden_text import Piece, normalize_text
from core.pdf_detector import BBox


@dataclass(frozen=True)
class Anchor:
    row: int
    col: int
    y: float  # 文字在 PDF 页面上的纵向中心
    x: float  # 文字左端


def build_anchors(cells: list[tuple[int, int, Any]], pieces: list[Piece]) -> dict[int, list[Anchor]]:
    """{页码: [锚点]}，按纵向位置排列"""
    cells_by_key: dict[str, list[tuple[int, int]]] = {}
    for row, col, value in cells:
        if isinstance(value, str) and (key := normalize_text(value)):
            cells_by_key.setdefault(key, []).append((row, col))
    pieces_by_key: dict[str, list[Piece]] = {}
    for piece in pieces:
        pieces_by_key.setdefault(piece.key, []).append(piece)

    anchors: dict[int, list[Anchor]] = {}
    for key, found in pieces_by_key.items():
        owners = cells_by_key.get(key, [])
        if len(found) != 1 or len(owners) != 1:
            continue
        piece, (row, col) = found[0], owners[0]
        x0, y0, _, y1 = piece.bbox
        anchors.setdefault(piece.page, []).append(Anchor(row, col, (y0 + y1) / 2, x0))
    for page_anchors in anchors.values():
        page_anchors.sort(key=lambda a: a.y)
    return anchors


def _estimate_row(
    lo: Anchor | None, hi: Anchor | None, y: float, row_height: Callable[[int], float],
) -> int | None:
    """在上下两个锚点之间，按 Excel 行高换算出纵向位置 y 对应的行"""
    if lo is None and hi is None:
        return None
    if lo is None or hi is None or lo.row >= hi.row:
        nearest = lo if hi is None or (lo is not None and y - lo.y <= hi.y - y) else hi
        return nearest.row  # type: ignore[union-attr]
    centers: dict[int, float] = {}
    top = 0.0
    for row in range(lo.row, hi.row + 1):
        height = max(0.0, row_height(row))
        centers[row] = top + height / 2
        top += height
    span = centers[hi.row] - centers[lo.row]
    if span <= 0:
        return lo.row
    scale = (hi.y - lo.y) / span
    return min(centers, key=lambda r: abs(lo.y + (centers[r] - centers[lo.row]) * scale - y))


def locate_cell(
    cells: list[tuple[int, int, Any]],
    anchors: dict[int, list[Anchor]],
    text: str,
    page_index: int,
    bbox: BBox,
    row_height: Callable[[int], float],
) -> tuple[int, int] | None:
    """返回 (行, 列)；这一页没有可用的锚点时返回 None，由调用方按旧方法估算。"""
    page_anchors = anchors.get(page_index, [])
    y = (bbox[1] + bbox[3]) / 2
    margin = max(1.0, (bbox[3] - bbox[1]) * 0.3)
    lo = next((a for a in reversed(page_anchors) if a.y <= y - margin), None)
    hi = next((a for a in page_anchors if a.y >= y + margin), None)
    estimate = _estimate_row(lo, hi, y, row_height)
    if estimate is None:
        return None

    key = normalize_text(text)
    candidates = [
        (row, col) for row, col, value in cells
        if value is not None and key and key in normalize_text(str(value))
    ]
    low = lo.row if lo else float("-inf")
    high = hi.row if hi else float("inf")
    in_range = [rc for rc in candidates if low <= rc[0] <= high]

    # 每列在这一页上的横向位置（取锚点里最靠左的文字起点），用来区分同一行里的同文字格子
    col_x: dict[int, float] = {}
    for anchor in page_anchors:
        col_x[anchor.col] = min(anchor.x, col_x.get(anchor.col, anchor.x))

    def x_distance(col: int) -> float:
        return abs(col_x[col] - bbox[0]) if col in col_x else 0.0

    pool = in_range or candidates
    if pool:
        return min(pool, key=lambda rc: (abs(rc[0] - estimate), x_distance(rc[1])))
    nearest_col = min(col_x, key=x_distance) if col_x else None
    return estimate, nearest_col if nearest_col is not None else min((c for _, c, _ in cells), default=1)
