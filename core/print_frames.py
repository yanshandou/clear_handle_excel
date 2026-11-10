"""导出 PDF 时临时藏起「整段内容外面的空心大框」（用户 2026-11-10 要）。

有的设计书用一个没填充色的矩形图形把一大段内容整个框起来。它跨好几页，PDF 里每页只印到一段：
中间页是左右两条粗竖线，最后一页是底边（像「凵」），看着像多出来的黑框。

认定为大框（全部满足才藏）：
- 自选图形里的矩形 / 圆角矩形，没有填充（Fill 不可见），有边线；
- 宽度 ≥ 表格已用区域宽度的 80%，高度 ≥ 已用区域高度的 30%（小的红框、标注框不动）。

导出前 Visible=False，导完 finally 里放回去；只改内存，不保存进 Excel。
认不准、读不到的一律不动，只写调试日志。
"""
from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

from core import debug_log

MSO_AUTO_SHAPE = 1
RECTANGLES = {1, 5}  # msoShapeRectangle、msoShapeRoundedRectangle
MIN_WIDTH_RATIO = 0.8
MIN_HEIGHT_RATIO = 0.3


def _used_size(sheet: Any) -> tuple[float, float] | None:
    try:
        used = sheet.UsedRange
        return float(used.Width), float(used.Height)
    except Exception:
        return None


def _is_frame(shape: Any, used_w: float, used_h: float) -> bool:
    try:
        if int(shape.Type) != MSO_AUTO_SHAPE or int(shape.AutoShapeType) not in RECTANGLES:
            return False
        if not shape.Visible or shape.Fill.Visible or not shape.Line.Visible:
            return False
        return float(shape.Width) >= used_w * MIN_WIDTH_RATIO and float(shape.Height) >= used_h * MIN_HEIGHT_RATIO
    except Exception:
        return False


def find_frames(sheet: Any) -> list[Any]:
    size = _used_size(sheet)
    if not size or size[0] <= 0 or size[1] <= 0:
        return []
    try:
        shapes = sheet.Shapes
        items = [shapes.Item(i) for i in range(1, int(shapes.Count) + 1)]
    except Exception:
        return []
    return [s for s in items if _is_frame(s, *size)]


@contextmanager
def hidden_frames(sheet: Any) -> Iterator[int]:
    """with 块里大框是藏着的；返回藏了几个"""
    hidden: list[Any] = []
    for shape in find_frames(sheet):
        try:
            shape.Visible = False
            hidden.append(shape)
        except Exception as err:
            debug_log.event("⚠️ 空心大框藏不起来", 表=str(getattr(sheet, "Name", "?")), 图形=str(getattr(shape, "Name", "?")), 错误=str(err))
    if hidden:
        debug_log.event("🔲 导出 PDF 临时藏起空心大框", 表=str(getattr(sheet, "Name", "?")),
                        图形="、".join(str(getattr(s, "Name", "?")) for s in hidden))
    try:
        yield len(hidden)
    finally:
        for shape in hidden:
            try:
                shape.Visible = True
            except Exception as err:
                debug_log.event("⚠️ 空心大框放不回去", 表=str(getattr(sheet, "Name", "?")), 图形=str(getattr(shape, "Name", "?")), 错误=str(err))
