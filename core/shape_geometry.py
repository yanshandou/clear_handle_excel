"""图形的大小 / 位置：读出来、和另一份对照算出怎么放回、放回去（用户 2026-11-11 要）。

两处在用：
- 修复图片（core/image_restore.py）：拿参考文件夹的原文件当底，放回被行高列宽拉坏的截图和红框；
- 改行高列宽保护（snapshot_book / restore_book）：程序修切字改行高、列宽前先记下图形，保存前放回，
  截图不被拉变形、红框不跑偏。

规则：截图 = 图片；整个落在截图里的其他图形算它的标注（落在几张里归最小那张）。截图和普通图形左上角不动、
大小放回；标注大小放回、位置 = 截图现在左上角 + 原来相对截图的偏移。只改差 > 1 磅的；名字重复的不碰。
这个模块不 import workflow / batch，谁都能用。
"""
from __future__ import annotations

from typing import Any

from core import debug_log

PICTURES = {11, 13}  # msoLinkedPicture、msoPicture
# 会被行列拉伸、要放回原样的图形：自选图形、标注、图表、任意多边形、组合、线条、图片、文本框
MOVABLE = {1, 2, 3, 5, 6, 9, 11, 13, 17}
TOL = 1.0         # 判断变没变（磅）
VERIFY_TOL = 1.5  # 改完回读允许的误差（Excel 会把位置凑到像素）
INSIDE_TOL = 2.0


def read_shapes(sheet: Any) -> list[dict]:
    """[{name, type, left, top, width, height}]，只读顶层、可见、会被拉伸的图形"""
    out = []
    try:
        shapes = sheet.Shapes
        count = int(shapes.Count)
    except Exception:
        return out
    for index in range(1, count + 1):
        try:
            shape = shapes.Item(index)
            kind = int(shape.Type)
            if kind not in MOVABLE or not shape.Visible:
                continue
            out.append({"name": str(shape.Name), "type": kind, "left": float(shape.Left), "top": float(shape.Top),
                        "width": float(shape.Width), "height": float(shape.Height)})
        except Exception:
            continue
    return out


def _box(s: dict) -> tuple[float, float, float, float]:
    return s["left"], s["top"], s["left"] + s["width"], s["top"] + s["height"]


def _inside(inner: dict, outer: dict) -> bool:
    a, b = _box(inner), _box(outer)
    return (a[0] >= b[0] - INSIDE_TOL and a[1] >= b[1] - INSIDE_TOL
            and a[2] <= b[2] + INSIDE_TOL and a[3] <= b[3] + INSIDE_TOL)


def _unique(shapes: list[dict]) -> dict[str, dict]:
    seen: dict[str, int] = {}
    for s in shapes:
        seen[s["name"]] = seen.get(s["name"], 0) + 1
    return {s["name"]: s for s in shapes if seen[s["name"]] == 1}


def overlay_hosts(current: list[dict], old: list[dict], every: bool = False) -> dict[str, list[str]]:
    """{截图名: [叠在它上面的图形名]}（按原版位置判断、两边名字类型都对得上的）。
    every=False 时没叠东西的截图不列；every=True 时所有截图、所有组合图形都列（组合里多半是截图+标注）"""
    cur, ref = _unique(current), _unique(old)
    names = [n for n in cur if n in ref and cur[n]["type"] == ref[n]["type"]]
    pictures = [n for n in names if ref[n]["type"] in PICTURES]
    hosts: dict[str, list[str]] = {n: [] for n in names if every and (n in pictures or ref[n]["type"] == 6)}
    for n in names:
        if n in pictures:
            continue
        inside = [p for p in pictures if _inside(ref[n], ref[p])]
        if inside:
            hosts.setdefault(min(inside, key=lambda p: ref[p]["width"] * ref[p]["height"]), []).append(n)
    return hosts


def diagnose(current: list[dict], old: list[dict]) -> dict:
    """调试日志用：两边图形各多少、配上几个、类型分布、最大差多少磅（查「为什么没找出来」）"""
    cur, ref = _unique(current), _unique(old)
    matched = [n for n in cur if n in ref and cur[n]["type"] == ref[n]["type"]]
    kinds: dict[int, int] = {}
    for s in old:
        kinds[s["type"]] = kinds.get(s["type"], 0) + 1
    worst = max((max(abs(cur[n][k] - ref[n][k]) for k in ("width", "height")) for n in matched), default=0.0)
    return {"原版图形": len(old), "新版图形": len(current), "配上": len(matched),
            "没配上": sorted(set(ref) - set(matched))[:8], "类型": kinds, "大小最大差": round(worst, 2),
            "截图": sum(1 for n in matched if ref[n]["type"] in PICTURES),
            "组合": sum(1 for n in matched if ref[n]["type"] == 6), "带标注的截图": len(overlay_hosts(current, old))}


def plan_sheet(current: list[dict], old: list[dict]) -> list[dict]:
    """返回 [{name, kind: picture|overlay|shape, host, before:{…}, after:{…}}]，只列真要改的"""
    cur, ref = _unique(current), _unique(old)
    names = [n for n in cur if n in ref and cur[n]["type"] == ref[n]["type"]]
    pictures = [n for n in names if ref[n]["type"] in PICTURES]
    host_of: dict[str, str] = {}
    for n in names:
        if n in pictures:
            continue
        hosts = [p for p in pictures if _inside(ref[n], ref[p])]
        if hosts:
            host_of[n] = min(hosts, key=lambda p: ref[p]["width"] * ref[p]["height"])
    changes = []

    def add(name: str, kind: str, left: float, top: float, host: str = "") -> None:
        before = {k: cur[name][k] for k in ("left", "top", "width", "height")}
        after = {"left": left, "top": top, "width": ref[name]["width"], "height": ref[name]["height"]}
        if any(abs(before[k] - after[k]) > TOL for k in after):
            item = {"name": name, "kind": kind, "host": host, "before": before, "after": after, "type": cur[name]["type"]}
            if host:  # 相对截图左上角的偏移：放回时按截图改完后的真实位置再算一次
                item["offset"] = (ref[name]["left"] - ref[host]["left"], ref[name]["top"] - ref[host]["top"])
            changes.append(item)

    for n in names:
        if n in host_of:
            continue
        add(n, "picture" if n in pictures else "shape", cur[n]["left"], cur[n]["top"])
    for n, host in host_of.items():
        left = cur[host]["left"] + ref[n]["left"] - ref[host]["left"]
        top = cur[host]["top"] + ref[n]["top"] - ref[host]["top"]
        add(n, "overlay", left, top, host)
    return changes


def describe_change(item: dict) -> str:
    b, a = item["before"], item["after"]
    parts = []
    if abs(b["width"] - a["width"]) > TOL or abs(b["height"] - a["height"]) > TOL:
        parts.append(f"大小 {b['width']:.0f}×{b['height']:.0f} → {a['width']:.0f}×{a['height']:.0f}")
    if abs(b["left"] - a["left"]) > TOL or abs(b["top"] - a["top"]) > TOL:
        parts.append(f"位置挪了 {a['left'] - b['left']:+.0f}, {a['top'] - b['top']:+.0f} 磅")
    label = {"picture": "截图", "overlay": f"截图「{item['host']}」上的标注", "shape": "图形"}[item["kind"]]
    return f"{label}「{item['name']}」：" + "，".join(parts)


# ---------- 改 ----------

def _set_geometry(shape: Any, target: dict) -> None:
    locked = None
    try:
        locked = shape.LockAspectRatio
        if locked:
            shape.LockAspectRatio = 0  # 锁着纵横比时改宽会连带改高
    except Exception:
        locked = None
    try:
        shape.Width = target["width"]
        shape.Height = target["height"]
        shape.Left = target["left"]
        shape.Top = target["top"]
    finally:
        if locked:
            try:
                shape.LockAspectRatio = locked
            except Exception:
                pass


def _matches(shape: Any, target: dict) -> bool:
    return all(abs(float(getattr(shape, key.capitalize())) - value) <= VERIFY_TOL for key, value in target.items())


def apply_sheet(sheet: Any, changes: list[dict]) -> list[str]:
    """标注要跟着截图的新位置走，所以先改截图和普通图形，再改标注；返回没对上的图形名"""
    order = sorted(changes, key=lambda c: c["kind"] == "overlay")
    for item in order:
        shape = sheet.Shapes.Item(item["name"])
        target = dict(item["after"])
        if item["kind"] == "overlay":  # 截图改完后再按它真实的位置算一次
            host = sheet.Shapes.Item(item["host"])
            target["left"] = float(host.Left) + item["offset"][0]
            target["top"] = float(host.Top) + item["offset"][1]
            item["after"] = target
        _set_geometry(shape, target)
    return [item["name"] for item in order if not _matches(sheet.Shapes.Item(item["name"]), item["after"])]




# ---------- 改行高列宽时保护图形 ----------

def snapshot_book(book: Any, sheet_names: Any = None) -> dict[str, list[dict]]:
    """改行高 / 列宽之前调用：记下这些表（不给就是所有表）里图形的大小和位置；没图形的表不记。读不到不报错"""
    snap: dict[str, list[dict]] = {}
    try:
        sheets = ([book.Worksheets(name) for name in dict.fromkeys(sheet_names)] if sheet_names is not None
                  else list(book.Worksheets))
    except Exception as err:
        debug_log.event("⚠️ 记不下图形位置，这次不保护图形", 错误=str(err))
        return snap
    for sheet in sheets:
        try:
            if int(sheet.Shapes.Count) == 0:
                continue
            shapes = read_shapes(sheet)
            if shapes:
                snap[str(sheet.Name)] = shapes
        except Exception:
            continue
    return snap


def restore_book(book: Any, snap: dict[str, list[dict]], label: str = "") -> int:
    """保存之前调用：被行高列宽拉变形的图片改回原大小、跑偏的标注放回截图上原来的位置。
    返回放回了几个图形；出错只写日志，不挡住原来的修改"""
    total = 0
    for name, before in snap.items():
        try:
            sheet = book.Worksheets(name)
            changes = plan_sheet(read_shapes(sheet), before)
            if not changes:
                continue
            details = [describe_change(item) for item in changes]
            bad = apply_sheet(sheet, changes)
            total += len(changes) - len(bad)
            debug_log.event("🛡️ 改行高列宽后把图形放回原样", 来源=label, 表=name, 图形=details, 没对上=bad or "无")
        except Exception as err:
            debug_log.event("⚠️ 图形没能放回原样", 来源=label, 表=name, 错误=str(err))
    return total
