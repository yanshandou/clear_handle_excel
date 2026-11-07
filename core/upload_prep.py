"""上传前每个文件都自动做的预处理（用户 2026-11-07 定，原来主页第 3、8 步不用再单独点）：

1. 清掉本工具加的标记色（条件格式里带 CHE_MARK 标识的规则、标成橙黄的表标签）——和「修复颜色」页的清除一样；
2. 对照参考（原件）恢复残留的配色——和「修复颜色」页「扫描并修复颜色」一样的规则：同一位置内容和参考一样、
   只有颜色不一样的格子改回参考的颜色；内容不一样、合并区域不同、配不上表的不动，记成「改不了」；
3. 视图标准化（分页预览、A1、100%）——和以前第 8 步一样。

三步都成功才上传；任何一步出错，这个文件就不传（distribute._run_task_upload 照旧不复制任何文件）。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from core import debug_log

MAX_LISTED = 50  # 每个文件最多记这么多处「颜色改不了」，太多时只写总数


def _clear_marks(excel: Any, local: Path) -> dict:
    from core import batch
    report: dict[str, Any] = {}
    batch.clear_marks_in_file(excel, str(local), report)
    if report.get("leftover_failed"):
        raise ValueError("有标记删不掉：" + "；".join(report["leftover_failed"][:3]))
    failed_tabs = [t for t in report.get("tabs", []) if t["status"] != "ok"]
    if failed_tabs:
        raise ValueError(failed_tabs[0]["reason"])
    return {"marks": int(report.get("leftovers", 0)), "tabs": sum(1 for t in report.get("tabs", []) if t["status"] == "ok")}


def _restore_colors(excel: Any, local: Path, reference: Path) -> dict:
    from core import mark_finder
    from core.link_repair import fingerprint

    stamp, old_stamp = fingerprint(local), fingerprint(reference)
    plan = mark_finder.compare_file_colors(excel, local, reference)
    if fingerprint(local) != stamp or fingerprint(reference) != old_stamp:
        raise ValueError("文件或参考在检查颜色期间发生变化，请关闭 Excel 后重试")
    cells = plan.get("cells", [])
    ready = [cell for cell in cells if cell.get("action") == "restore"]
    done: list[dict] = []
    if ready:
        done = mark_finder.apply_file(excel, str(local), str(reference), ready, versions=(stamp, old_stamp))
    outcomes = {mark_finder._item_key(cell): cell for cell in done}
    results = [outcomes.get(mark_finder._item_key(cell), cell) for cell in cells]
    fixed = [cell for cell in results if cell.get("status") == "ok"]
    skipped = [cell for cell in results if cell.get("status") != "ok"]
    for cell in results:
        debug_log.event("🎨 上传前恢复颜色", 位置=f"{cell['sheet']}!{cell['addr']}", 结果=cell.get("reason", ""))
    return {"colors": len(fixed), "color_skipped": len(skipped),
            "color_problems": [{"sheet": cell["sheet"], "cell": cell.get("anchor") or cell["addr"],
                                "reason": cell.get("reason", "")} for cell in skipped[:MAX_LISTED]]}


def _normalize(excel: Any, local: Path) -> None:
    from core import print_layout, workflow as wf
    from core.excel_ops import normalize_workbook_view, open_workbook_checked

    book = open_workbook_checked(excel, str(local.resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开（可能在 Excel 里开着），没能标准化视图")
        problems: list[str] = []
        normalize_workbook_view(excel, book, on_error=problems.append)
        for line in problems:
            debug_log.event("⚠️ 视图标准化跳过", 说明=line)
        before = wf._file_version(local)
        book.Save()
        print_layout.followed_save(wf.STATE.setdefault("print_layouts", {}), str(local), before)
    finally:
        book.Close(SaveChanges=False)


def prepare(get_excel: Callable[[], Any], local: Path, reference: Path | None, where: dict | None = None) -> dict:
    """对一个要上传的文件做完三步。出错时原样抛出异常（调用方按 file_problem 的规则决定重试 / 跳过），
    where["step"] 写着是哪一步出的错"""
    where = where if where is not None else {}
    result: dict[str, Any] = {"marks": 0, "tabs": 0, "colors": 0, "color_skipped": 0, "color_problems": [],
                              "reference": bool(reference and reference.is_file()), "normalized": False}
    where["step"] = "清除标记色"
    result.update(_clear_marks(get_excel(), local))
    if result["reference"]:
        where["step"] = "对照参考恢复颜色"
        result.update(_restore_colors(get_excel(), local, reference))
    where["step"] = "视图标准化"
    _normalize(get_excel(), local)
    result["normalized"] = True
    where["step"] = ""
    debug_log.event("🧽 上传前处理完成", 文件=local.name, 清标记=result["marks"], 还原标签=result["tabs"],
                    恢复颜色=result["colors"], 颜色改不了=result["color_skipped"], 有参考=result["reference"])
    return result


def text(prep: dict | None) -> str:
    """一行说明，网页和上传记录共用"""
    if not prep:
        return ""
    parts = [f"清掉标记 {prep['marks']} 条" if prep["marks"] else "没有标记要清"]
    if prep["tabs"]:
        parts.append(f"还原表标签 {prep['tabs']} 张")
    if prep.get("no_reference"):
        parts.append("起点 → 终点上传没有参考，不对照颜色")
    elif not prep["reference"]:
        parts.append("参考里没有这份文件，没对照颜色")
    else:
        parts.append(f"按参考改回颜色 {prep['colors']} 处" if prep["colors"] else "颜色和参考一致")
        if prep["color_skipped"]:
            parts.append(f"颜色改不了 {prep['color_skipped']} 处")
    if prep["normalized"]:
        parts.append("视图已标准化（A1、100%）")
    return "，".join(parts)
