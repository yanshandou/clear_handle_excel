"""上传前每个文件都自动做的预处理（用户 2026-11-07 定，原来主页第 3、8 步不用再单独点）：

1. 清掉本工具加的标记色（条件格式里带 CHE_MARK 标识的规则、标成橙黄的表标签）——和「修复颜色」页的清除一样；
2. 对照参考（原件）恢复残留的配色——和「修复颜色」页「扫描并修复颜色」一样的规则：同一位置内容和参考一样、
   只有颜色不一样的格子改回参考的颜色；内容不一样、合并区域不同、配不上表的不动，记成「改不了」；
3. 视图标准化（分页预览、A1、100%）——和以前第 8 步一样。

三步都成功才上传；任何一步出错，这个文件就不传，同一批的其他文件照传（用户 2026-11-10）。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from core import debug_log

MAX_LISTED = 50  # 每个文件最多记这么多处「颜色改不了」，太多时只写总数
XML_WORKBOOKS = {".xlsx", ".xlsm"}


def _clear_marks(excel: Any, local: Path) -> dict:
    from core import batch
    report: dict[str, Any] = {}
    batch.clear_marks_in_file(excel, str(local), report)
    return _mark_counts(report)


def _mark_counts(report: dict) -> dict:
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
    return _color_counts(cells, done)


def _color_counts(cells: list[dict], done: list[dict]) -> dict:
    from core import mark_finder

    outcomes = {mark_finder._item_key(cell): cell for cell in done}
    results = [outcomes.get(mark_finder._item_key(cell), cell) for cell in cells]
    fixed = [cell for cell in results if cell.get("status") == "ok"]
    skipped = [cell for cell in results if cell.get("status") != "ok"]
    for cell in results:
        debug_log.event("🎨 上传前恢复颜色", 位置=f"{cell['sheet']}!{cell['addr']}", 结果=cell.get("reason", ""))
    return {"colors": len(fixed), "color_skipped": len(skipped),
            "color_problems": [{"sheet": cell["sheet"], "cell": cell.get("anchor") or cell["addr"],
                                "reason": cell.get("reason", "")} for cell in skipped[:MAX_LISTED]]}


def _normalize_view(excel: Any, book: Any) -> None:
    from core.excel_ops import normalize_workbook_view

    problems: list[str] = []
    normalize_workbook_view(excel, book, on_error=problems.append)
    for line in problems:
        debug_log.event("⚠️ 视图标准化跳过", 说明=line)


def _normalize(excel: Any, local: Path) -> None:
    from core import print_layout, workflow as wf
    from core.excel_ops import open_workbook_checked

    book = open_workbook_checked(excel, str(local.resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开（可能在 Excel 里开着），没能标准化视图")
        _normalize_view(excel, book)
        before = wf._file_version(local)
        book.Save()
        print_layout.followed_save(wf.STATE.setdefault("print_layouts", {}), str(local), before)
    finally:
        book.Close(SaveChanges=False)


def _prepare_xml(excel: Any, local: Path, reference: Path | None, where: dict, result: dict) -> None:
    from core import batch, color_compare, mark_finder, print_layout, workflow as wf
    from core.excel_ops import open_workbook_checked
    from core.link_repair import fingerprint

    batch._check_file(str(local))
    cells: list[dict] = []
    old_colors: dict[tuple, dict | str] = {}
    if reference is not None:
        where["step"] = "对照参考恢复颜色"
        stamp, old_stamp = fingerprint(local), fingerprint(reference)
        cells = color_compare.plan_xml(local, reference, clear_tab_color=batch.TAB_HIGHLIGHT)["cells"]
        ready = {mark_finder._item_key(cell): cell for cell in cells if cell["action"] == "restore"}
        if ready:
            old_colors = mark_finder._read_old(excel, str(reference), list(ready.values()))

    def check_versions() -> None:
        if reference is not None and (fingerprint(local) != stamp or fingerprint(reference) != old_stamp):
            raise ValueError("当前文件或备份在检查颜色期间发生变化，请重新检查")

    check_versions()
    where["step"] = "清除标记色"
    book = open_workbook_checked(excel, str(local.resolve()), UpdateLinks=0)
    done: list[dict] = []
    try:
        if book.ReadOnly:
            raise ValueError(f"{local.name} 以只读方式打开，未能清除标记")
        report = batch.clear_workbook_marks(excel, book)
        result.update(_mark_counts(report))
        if reference is not None:
            where["step"] = "对照参考恢复颜色"
            cleared = {tab["sheet"] for tab in report["tabs"]}
            cells = [cell for cell in cells
                     if "tab_cleared" not in cell or cell["tab_cleared"] == (cell["sheet"] in cleared)]
            ready_cells = [cell for cell in cells if cell["action"] == "restore"]
            done = mark_finder.restore_workbook_colors(book, ready_cells, old_colors)
        where["step"] = "视图标准化"
        _normalize_view(excel, book)
        # 旧流程只有最后的纯视图保存会延续打印方向记录的版本。
        view_only = not (result["marks"] or result["tabs"] or any(item["status"] == "ok" for item in done))
        before = wf._file_version(local) if view_only else None
        check_versions()
        book.Save()
        print_layout.followed_save(wf.STATE.setdefault("print_layouts", {}), str(local), before)
    finally:
        book.Close(SaveChanges=False)
    if reference is not None:
        result.update(_color_counts(cells, done))


def prepare(get_excel: Callable[[], Any], local: Path, reference: Path | None, where: dict | None = None) -> dict:
    """对一个要上传的文件做完三步。出错时原样抛出异常（调用方按 file_problem 的规则决定重试 / 跳过），
    where["step"] 写着是哪一步出的错"""
    where = where if where is not None else {}
    reference = reference if reference is not None and reference.is_file() else None
    result: dict[str, Any] = {"marks": 0, "tabs": 0, "colors": 0, "color_skipped": 0, "color_problems": [],
                              "reference": reference is not None, "normalized": False}
    where["step"] = "清除标记色"
    excel = get_excel()
    if local.suffix.lower() in XML_WORKBOOKS and (reference is None or (
            reference.suffix.lower() in XML_WORKBOOKS and reference.resolve() != local.resolve())):
        _prepare_xml(excel, local, reference, where, result)
    else:
        result.update(_clear_marks(excel, local))
        if reference is not None:
            where["step"] = "对照参考恢复颜色"
            result.update(_restore_colors(excel, local, reference))
        where["step"] = "视图标准化"
        _normalize(excel, local)
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
