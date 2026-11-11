"""导出 PDF 时，把「叠在截图上的红框、箭头、标注」临时和截图组合在一起（用户 2026-11-11 问）。

现象：Excel 里红框正好框住截图上的「実行」按钮，导出的 PDF 里截图缩了一点，红框却跑到按钮左上方。
原因：截图和红框是两个独立的图形，各自挂在自己左上角的格子上。打印 / 导出 PDF 时 Excel 按打印机重新量
列宽、行高（标准字体是游ゴシック时差得最明显），格子网格比屏幕上缩了几个百分点，可图形本身按磅数画、
大小不跟着缩。红框的位置 = 它挂的格子 + 偏移，跟着网格往左上缩；按钮在截图里的位置却是截图左上角 + 固定磅数，
两者就错开了，离截图左上角越远错得越多。组合成一个图形后，里面的东西按同一个比例一起画，就不会错开。

导出时的做法（grouped_overlays，2026-11-11 扩大）：看得见的图片 / 自选图形 / 线条 / 文本框 / 任意多边形 / 标注 / 已有的组合，
互相重叠或隔 2 磅以内的连成一堆，每堆临时组合成一个图形——没有底图、全靠图形拼出来的画面也一起管。
重名的图形按图形 ID 找序号来组合。

永久组合的做法（group_permanently，会保存进文件，所以保守；全部满足才组合）：
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


MSO_GROUP = 6
CLUSTER_TYPES = PICTURES | OVERLAYS | {MSO_GROUP}
GAP = 2.0  # 两个图形隔这么近（磅）以内就算挨着


def _touch(a: tuple, b: tuple) -> bool:
    return a[0] <= b[2] + GAP and b[0] <= a[2] + GAP and a[1] <= b[3] + GAP and b[1] <= a[3] + GAP


def find_clusters(sheet: Any) -> list[list[Any]]:
    """互相重叠或紧挨着的图形分成一堆一堆（至少两个图形才算一堆），每堆按原来的叠放顺序。
    用户 2026-11-11：用很多图形拼出来的画面（按钮、标注、虚线框，没有底图），导出 PDF 后蓝色按钮掉到绿色标注下面，
    手动组合后就好了——各图形挂各的格子、有的随格子动有的不动，打印时网格一伸缩就互相错开。
    看得见的图片、自选图形、线、文本框、任意多边形、标注、已有的组合都算；图表、控件、批注不碰"""
    try:
        shapes = sheet.Shapes
        items = [shapes.Item(i) for i in range(1, int(shapes.Count) + 1)]
    except Exception:
        return []
    found = []
    for shape in items:
        try:
            kind = int(shape.Type)
        except Exception:
            continue
        if kind not in CLUSTER_TYPES or not _visible(shape):
            continue
        box = _box(shape)
        if box is not None:
            found.append((shape, box))
    parent = list(range(len(found)))

    def root(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i in range(len(found)):
        for j in range(i + 1, len(found)):
            if _touch(found[i][1], found[j][1]):
                parent[root(j)] = root(i)
    clusters: dict[int, list[Any]] = {}
    for i, (shape, _box_) in enumerate(found):
        clusters.setdefault(root(i), []).append(shape)
    return [members for members in clusters.values() if len(members) > 1]


def _shape_id(shape: Any) -> int | None:
    try:
        return int(shape.ID)
    except Exception:
        return None


def _name_counts(sheet: Any) -> dict[str, int]:
    counts: dict[str, int] = {}
    try:
        shapes = sheet.Shapes
        for i in range(1, int(shapes.Count) + 1):
            name = str(getattr(shapes.Item(i), "Name", "") or "")
            counts[name] = counts.get(name, 0) + 1
    except Exception:
        pass
    return counts


def _range_of(sheet: Any, members: list[Any], counts: dict[str, int]) -> Any:
    """这几个图形的 ShapeRange：名字在表里都不重复就按名字拿（不用再读一遍图形，图形多的表省时间）；
    有重名的（复制粘贴出来的图形常常同名）按图形 ID 找现在的序号拿"""
    shapes = sheet.Shapes
    wanted = [str(s.Name) for s in members]
    if all(counts.get(n) == 1 for n in wanted):
        return shapes.Range(wanted)
    ids = {_shape_id(s) for s in members}
    if None in ids or len(ids) != len(members):
        raise RuntimeError("有重名图形，又读不到图形 ID")
    indexes = [i for i in range(1, int(shapes.Count) + 1) if _shape_id(shapes.Item(i)) in ids]
    if len(indexes) != len(members):
        raise RuntimeError("按图形 ID 没找全")
    return shapes.Range(indexes)


def _restore_placements(sheet: Any, parts: Any, by_id: dict[int, Any], by_name: dict[str, Any]) -> None:
    """拆开后把每个图形原来的「对象の位置関係」放回去：先按拆出来的图形的 ID 认，认不到再按（不重名的）名字"""
    done: set[str] = set()
    try:
        members = [parts.Item(i) for i in range(1, int(parts.Count) + 1)] if parts is not None else []
    except Exception:
        members = []
    for shape in members:
        value = by_id.get(_shape_id(shape))
        if value is None:
            continue
        try:
            if shape.Placement != value:
                shape.Placement = value
            done.add(str(shape.Name))
        except Exception as err:
            debug_log.event("⚠️ 图形的位置关係放不回去", 表=str(getattr(sheet, "Name", "?")), 图形=str(getattr(shape, "Name", "?")), 错误=str(err))
    for name, value in by_name.items():
        if value is None or name in done:
            continue
        try:
            shape = sheet.Shapes.Item(name)
            if shape.Placement != value:
                shape.Placement = value
        except Exception as err:
            debug_log.event("⚠️ 图形的位置关係放不回去", 表=str(getattr(sheet, "Name", "?")), 图形=name, 错误=str(err))


@contextmanager
def grouped_overlays(sheet: Any) -> Iterator[int]:
    """with 块里互相重叠 / 挨着的图形是一堆一堆组合着的（截图 + 上面的红框也在内）；返回组合了几组"""
    sheet_name = str(getattr(sheet, "Name", "?"))
    made: list[tuple[Any, dict[int, Any], dict[str, Any], list[str]]] = []  # (组合, {ID: 原 Placement}, {不重名的名字: 原 Placement}, 名字)
    clusters = find_clusters(sheet)
    counts = _name_counts(sheet) if clusters else {}
    for members in clusters:
        names = [str(getattr(s, "Name", "?")) for s in members]
        by_id = {i: _placement(s) for s in members if (i := _shape_id(s)) is not None}
        by_name = {n: _placement(s) for s, n in zip(members, names) if counts.get(n) == 1}
        try:
            group = _range_of(sheet, members, counts).Group()
            made.append((group, by_id, by_name, names))
        except Exception as err:
            debug_log.event("⚠️ 挨着的图形组合不了，照原样导出", 表=sheet_name, 图形="、".join(names[:12]), 个数=len(names), 错误=str(err))
    if made:
        debug_log.event("📌 导出 PDF 临时把挨着的图形组合起来（防止互相错位）", 表=sheet_name, 组数=len(made),
                        组="；".join("、".join(n[:6]) + (f"…共 {len(n)} 个" if len(n) > 6 else "") for *_x, n in made[:20]))
    try:
        yield len(made)
    finally:
        for group, by_id, by_name, names in made:
            try:
                parts = group.Ungroup()
            except Exception as err:
                debug_log.event("⚠️ 临时组合拆不开", 表=sheet_name, 图形="、".join(names[:12]), 错误=str(err))
                continue
            _restore_placements(sheet, parts, by_id, by_name)


XL_MOVE = 2  # 「セルに合わせて移動するがサイズ変更はしない」


def group_permanently(sheet: Any) -> list[str]:
    """把截图和上面的红框、箭头、文字框永久组合（用户 2026-11-11 要，领导演示过），组合设成「随单元格移动、不改变大小」：
    以后再改行高列宽，整组只跟着挪位置，截图不变形、红框不跑偏；打印 / PDF 也不再错开。
    返回组合了哪些（每组一句）；组合不了的只写日志。调用方负责保存"""
    sheet_name = str(getattr(sheet, "Name", "?"))
    done = []
    for picture, members in find_overlay_sets(sheet):
        names = [str(s.Name) for s in [picture, *members]]
        try:
            group = sheet.Shapes.Range(names).Group()
        except Exception as err:
            debug_log.event("⚠️ 截图和标注没能永久组合", 表=sheet_name, 图形="、".join(names), 错误=str(err))
            continue
        try:
            group.Placement = XL_MOVE
        except Exception as err:
            debug_log.event("⚠️ 组合设不了「随单元格移动、不改变大小」", 表=sheet_name, 图形="、".join(names), 错误=str(err))
        done.append(f"截图「{names[0]}」+ {len(members)} 个标注（{'、'.join(names[1:6])}{'…' if len(names) > 6 else ''}）")
    if done:
        debug_log.event("📌 截图和标注永久组合", 表=sheet_name, 组=done)
    return done
