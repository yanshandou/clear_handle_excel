"""导出 PDF 时，把「叠在截图上的红框、箭头、标注」临时和截图组合在一起（用户 2026-11-11 问）。

现象：Excel 里红框正好框住截图上的「実行」按钮，导出的 PDF 里截图缩了一点，红框却跑到按钮左上方。
原因：截图和红框是两个独立的图形，各自挂在自己左上角的格子上。打印 / 导出 PDF 时 Excel 按打印机重新量
列宽、行高（标准字体是游ゴシック时差得最明显），格子网格比屏幕上缩了几个百分点，可图形本身按磅数画、
大小不跟着缩。红框的位置 = 它挂的格子 + 偏移，跟着网格往左上缩；按钮在截图里的位置却是截图左上角 + 固定磅数，
两者就错开了，离截图左上角越远错得越多。组合成一个图形后，里面的东西按同一个比例一起画，就不会错开。

做法（全部满足才组合）：
- 截图：图片（嵌入或链接），可见，不在组合里；
- 叠加物：自选图形 / 线条 / 文本框 / 任意多边形 / 标注，可见，不在组合里，整个落在截图范围里（容差 2 磅）；
  落在好几张截图里的，归最小那张；
- 名字重复的不碰（按名字组合，重名会组错）。

导出前 Group()，导完 finally 里 Ungroup()，再把每个图形原来的「对象の位置関係」（Placement）放回去。
只读打开的只改内存；可写打开的随调用方保存，拆开后位置、大小、名字不变。认不准、改不了的一律不动，只写调试日志。
"""
from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

from core import debug_log

MSO_PICTURE = 13
MSO_LINKED_PICTURE = 11
PICTURES = {MSO_PICTURE, MSO_LINKED_PICTURE}
# msoAutoShape、msoCallout、msoFreeform、msoLine、msoTextBox
OVERLAYS = {1, 2, 5, 9, 17}
TOLERANCE = 2.0


def _box(shape: Any) -> tuple[float, float, float, float] | None:
    try:
        left, top = float(shape.Left), float(shape.Top)
        return left, top, left + float(shape.Width), top + float(shape.Height)
    except Exception:
        return None


def _inside(inner: tuple, outer: tuple) -> bool:
    return (inner[0] >= outer[0] - TOLERANCE and inner[1] >= outer[1] - TOLERANCE
            and inner[2] <= outer[2] + TOLERANCE and inner[3] <= outer[3] + TOLERANCE)


def _area(box: tuple) -> float:
    return (box[2] - box[0]) * (box[3] - box[1])


def _visible(shape: Any) -> bool:
    try:
        return bool(shape.Visible)
    except Exception:
        return False


def find_overlay_sets(sheet: Any) -> list[tuple[Any, list[Any]]]:
    """[(截图, [叠在上面的图形])]，只返回至少有一个叠加物的截图"""
    try:
        shapes = sheet.Shapes
        items = [shapes.Item(i) for i in range(1, int(shapes.Count) + 1)]
    except Exception:
        return []
    names: dict[str, int] = {}
    for shape in items:
        name = str(getattr(shape, "Name", "") or "")
        names[name] = names.get(name, 0) + 1
    pictures, overlays = [], []
    for shape in items:
        try:
            kind = int(shape.Type)
        except Exception:
            continue
        if not _visible(shape):
            continue
        box = _box(shape)
        if box is None:
            continue
        if kind in PICTURES:
            pictures.append((shape, box))
        elif kind in OVERLAYS:
            overlays.append((shape, box))
    if not pictures or not overlays:
        return []
    groups: dict[int, list[Any]] = {}
    for shape, box in overlays:
        hosts = [(i, _area(pbox)) for i, (_p, pbox) in enumerate(pictures) if _inside(box, pbox)]
        if hosts:
            groups.setdefault(min(hosts, key=lambda h: h[1])[0], []).append(shape)
    result = []
    for index, members in groups.items():
        picture = pictures[index][0]
        everyone = [picture, *members]
        if any(names.get(str(getattr(s, "Name", "") or ""), 0) != 1 for s in everyone):
            debug_log.event("⚠️ 截图上的标注有重名图形，不组合", 表=str(getattr(sheet, "Name", "?")),
                            截图=str(getattr(picture, "Name", "?")))
            continue
        result.append((picture, members))
    return result


def _placement(shape: Any) -> Any:
    try:
        return shape.Placement
    except Exception:
        return None


@contextmanager
def grouped_overlays(sheet: Any) -> Iterator[int]:
    """with 块里截图和上面的标注是组合着的；返回组合了几组"""
    sheet_name = str(getattr(sheet, "Name", "?"))
    made: list[tuple[Any, dict[str, Any]]] = []  # (组合, {名字: 原 Placement})
    for picture, members in find_overlay_sets(sheet):
        everyone = [picture, *members]
        names = [str(s.Name) for s in everyone]
        placements = {str(s.Name): _placement(s) for s in everyone}
        try:
            group = sheet.Shapes.Range(names).Group()
            made.append((group, placements))
        except Exception as err:
            debug_log.event("⚠️ 截图和标注组合不了，照原样导出", 表=sheet_name, 图形="、".join(names), 错误=str(err))
    if made:
        debug_log.event("📌 导出 PDF 临时把截图和上面的标注组合起来（防止红框错位）", 表=sheet_name,
                        组="；".join("、".join(p) for _g, p in made))
    try:
        yield len(made)
    finally:
        for group, placements in made:
            try:
                group.Ungroup()
            except Exception as err:
                debug_log.event("⚠️ 截图和标注拆不开", 表=sheet_name, 图形="、".join(placements), 错误=str(err))
                continue
            for name, value in placements.items():
                if value is None:
                    continue
                try:
                    shape = sheet.Shapes.Item(name)
                    if shape.Placement != value:
                        shape.Placement = value
                except Exception as err:
                    debug_log.event("⚠️ 图形的位置关係放不回去", 表=sheet_name, 图形=name, 错误=str(err))
