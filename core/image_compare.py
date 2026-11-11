"""按图片比：原版 Excel / 新版 Excel / 新版 PDF 里同一张截图（连同上面的红框）缩成一样大，看红框偏了多少、整体差多少
（用户 2026-11-11 要：按图形数字对照一份都没找出来，人工却看出很多——多半是只在 PDF 里错开）。

- 只比截图本身那一块，周围格子改了行高、字体都不影响；截图被拉长，缩回同样大小后内容一样，红框位置就偏了。
- 红框：偏红的像素（R 高、G/B 低）的重心，两张差多少，按截图边长的百分比算；一张有红一张没有算 100%。
- 整体差：灰度逐点差的平均（0～1）。
- 不用 numpy / Pillow（程序没装），用 PyMuPDF 读图、纯 Python 采样，48×48 足够。
"""
from __future__ import annotations

from typing import Any

from core.pdf_detector import fitz

SIZE = 48
RED_MIN = 3            # 至少几个红点才算有红框
RED_SHIFT_SUSPECT = 4  # 红框偏 > 4% 算疑似
DIFF_SUSPECT = 0.12    # 整体差 > 0.12 算疑似

Thumb = list[tuple[int, int, int]]


def _sample(pix: Any, x0: float, y0: float, x1: float, y1: float) -> Thumb | None:
    """在像素框 (x0,y0)-(x1,y1) 里按 SIZE×SIZE 网格取点（取格子中心最近的像素）；透明当白色"""
    x0, y0 = max(0.0, x0), max(0.0, y0)
    x1, y1 = min(float(pix.width), x1), min(float(pix.height), y1)
    if x1 - x0 < 4 or y1 - y0 < 4:
        return None
    out: Thumb = []
    alpha = bool(pix.alpha)
    for j in range(SIZE):
        y = min(pix.height - 1, int(y0 + (j + 0.5) * (y1 - y0) / SIZE))
        for i in range(SIZE):
            x = min(pix.width - 1, int(x0 + (i + 0.5) * (x1 - x0) / SIZE))
            value = pix.pixel(x, y)
            if alpha and value[-1] < 128:
                out.append((255, 255, 255))
            else:
                out.append((value[0], value[1], value[2]) if len(value) >= 3 else (value[0],) * 3)
    return out


def thumb_from_png(path: str, area: tuple[float, float, float, float], box: tuple[float, float, float, float]) -> Thumb | None:
    """area：截出来那块区域在表上的 (左, 上, 宽, 高)（磅）；box：截图在表上的 (左, 上, 宽, 高)。按比例换成像素再取样"""
    pix = fitz.Pixmap(path)
    scale_x, scale_y = pix.width / area[2], pix.height / area[3]
    x0, y0 = (box[0] - area[0]) * scale_x, (box[1] - area[1]) * scale_y
    return _sample(pix, x0, y0, x0 + box[2] * scale_x, y0 + box[3] * scale_y)


def thumb_from_page(page: Any, rect: Any) -> Thumb | None:
    zoom = max(1.0, 4 * SIZE / max(1.0, float(rect.width)))
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=rect, annots=False, alpha=False)
    return _sample(pix, 0, 0, pix.width, pix.height)


def _is_red(p: tuple[int, int, int]) -> bool:
    r, g, b = p
    return r > 150 and g < 110 and b < 110 and r - max(g, b) > 60


def _red_center(t: Thumb) -> tuple[float, float, int]:
    xs = ys = n = 0
    for index, p in enumerate(t):
        if _is_red(p):
            xs += index % SIZE
            ys += index // SIZE
            n += 1
    return (xs / n, ys / n, n) if n else (0.0, 0.0, 0)


def compare(a: Thumb | None, b: Thumb | None) -> dict | None:
    """{"diff": 0～1, "red_shift": 百分比或 None（两张都没红框）, "suspect": bool}；有一张没有就返回 None"""
    if not a or not b or len(a) != len(b):
        return None
    diff = sum(abs(sum(p) - sum(q)) for p, q in zip(a, b)) / (3 * 255 * len(a))
    ax, ay, an = _red_center(a)
    bx, by, bn = _red_center(b)
    if an >= RED_MIN and bn >= RED_MIN:
        shift = ((ax - bx) ** 2 + (ay - by) ** 2) ** 0.5 / SIZE * 100
    elif an >= RED_MIN or bn >= RED_MIN:
        shift = 100.0  # 一张有红框一张没有
    else:
        shift = None
    suspect = (shift is not None and shift > RED_SHIFT_SUSPECT) or diff > DIFF_SUSPECT
    return {"diff": round(diff, 3), "red_shift": None if shift is None else round(shift, 1), "suspect": suspect}
