"""查看切字的单元格工具箱（用户 2026-11-08 要，照「调整页面」的样子）。

网页上按问题分类，每一类（或单个格子）选一个工具，点「🛠 统一修改」后：
改 Excel → 保存 → 整份复检。通过的格子翻篇、去掉标色；没过的留在原分类，写上用过什么工具，
由用户再选别的工具或者打开 Excel 自己改。不自动升一级、不自动试下一个工具。

工具：
- grow：行高 ×1.2 / ×1.5 / ×2。开自动换行，让 Excel 在临时工作簿量出放下要多高（量不出按单词换行估算），
  再加「字号 × 倍数」。合并格加在最下面一行；同一行几个格子取要的最高的；最高 409pt。
  倍数按第一次用工具前的原行高算（data["fix_rows"] 记原行高和程序改成的高度），换倍数不会越叠越高；
  行高被人改过（和程序上次改的不一样）就按现在的行高重新算。
- shrink：缩小字体填充（关自动换行，否则 Excel 不缩）。
- break：在这格上面加一条手动分页线（分页切开合并单元格）。页高定成 N 页（取整）的表分页线不起作用，跳过。
不放「加列宽」：2026-11-04 试过，表是一页宽时整页跟着缩小，切字反而更多（#171）。
"""
from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from core import autofix, debug_log, file_problem, pdf_store, print_layout, shape_geometry, sheet_history, workflow as wf
from core import batch as b
from core.common import workbook_rel_name
from core.excel_ops import close_saved_workbooks, normalize_workbook_view, open_workbook_checked
from core.reporter import update_file_progress

GROW_RATIOS = (1.2, 1.5, 2.0)
CUT_TYPES = set(autofix.CUT_ERROR_TYPES)
# 每种问题能用哪些工具（网页按这个画按钮，后台也按这个核对）
TOOLS_BY_TYPE: dict[str, tuple[str, ...]] = {
    **{t: ("grow", "shrink") for t in autofix.CUT_ERROR_TYPES},
    autofix.ERROR_TYPE_HASH_OVERFLOW: ("shrink",),
    autofix.ERROR_TYPE_PAGE_BREAK_CUT: ("break",),
}
PASSED_MESSAGE = "用单元格工具修改后复检通过"


def tool_label(tool: str, ratio: float | None = None) -> str:
    if tool == "grow":
        return f"行高 ×{ratio:g}"
    return {"shrink": "缩小字体填充", "break": "上面加分页线"}.get(tool, tool)


def failed_message(fix: dict) -> str:
    return f"用了「{fix['label']}」复检还没过：可以换个工具再统一修改，或打开 Excel 自己改"


def allowed_tools(card: dict) -> list[str]:
    tools: list[str] = []
    for kind in b._card_error_types(card):
        for tool in TOOLS_BY_TYPE.get(kind, ()):
            if tool not in tools:
                tools.append(tool)
    return tools


def _fixes() -> dict[str, dict]:
    return b._data().setdefault("cell_fixes", {})


def _fix_rows() -> dict[str, dict]:
    return b._data().setdefault("fix_rows", {})


def card_fix(card_id: str) -> dict | None:
    """调用方持有 STATE_LOCK：这格上次用过的工具（网页显示「用过 ×1.5 没过」）"""
    fix = _fixes().get(card_id)
    return dict(fix) if fix else None


def validate(payload: Any) -> dict[str, dict]:
    """网页发来 {"tools": [{"card_id", "tool", "ratio"?}]}，返回 {卡片 id: {"tool", "ratio"}}"""
    raw = payload.get("tools") if isinstance(payload, dict) else None
    if not isinstance(raw, list) or not raw:
        raise ValueError("请先给格子选工具")
    ids = [entry.get("card_id") if isinstance(entry, dict) else None for entry in raw]
    b.validate_cards(ids, yellow=True)
    plan: dict[str, dict] = {}
    with wf.STATE_LOCK:
        cards = b._data()["cards"]
        for entry in raw:
            cid, tool = entry["card_id"], entry.get("tool")
            card = cards[cid]
            where = f"{card['wb_name']} / {card['sheet_name']}!{card['cell_addr']}"
            if tool not in allowed_tools(card):
                raise ValueError(f"{where} 不能用这个工具")
            choice: dict[str, Any] = {"tool": tool}
            if tool == "grow":
                ratio = entry.get("ratio")
                if isinstance(ratio, bool) or not isinstance(ratio, (int, float)) or float(ratio) not in GROW_RATIOS:
                    raise ValueError("行高倍数只能选 1.2、1.5 或 2")
                choice["ratio"] = float(ratio)
            plan[cid] = choice
    return plan


def job_options(plan: dict[str, dict]) -> dict:
    return {"card_ids": list(plan), "tools": plan}


def _page_count(sheet: Any) -> int | None:
    try:
        return int(sheet.PageSetup.Pages.Count)
    except Exception:
        return None


def _grow_entry(measurer: Any, sheet: Any, area: Any, card: dict, items: list[dict], ratio: float,
                rows: dict, path: str) -> dict:
    """算出这格要把哪一行加到多高，合进 rows（同一行取最高）。返回这格的结果说明"""
    size = b._font_size(area)
    shrink = bool(area.ShrinkToFit)
    if not shrink and b._has_horizontal_space(sheet, area, size):
        # 极窄格子里的横排标题、右边都是空格：开换行会变成一竖条字，不加高，留给人工看
        return {"skip": "右边是空格、文字横着放得下，开换行会变成竖条：没加高，请打开看"}
    if not shrink and area.WrapText is not True:
        area.WrapText = True
        wrapped = True
    else:
        wrapped = False
    last = int(area.Row) + int(area.Rows.Count) - 1  # 合并单元格加在最下面一行
    target_row = sheet.Rows(last)
    current = float(target_row.RowHeight)
    others = max(0.0, float(area.Height) - current)  # 合并格在别的行占的高度
    row_key = b._key(path, sheet.Name, last)
    record = _fix_rows().get(row_key)
    orig = float(record["orig"]) if record and abs(current - float(record["applied"])) < 0.01 else current
    need, by = 0.0, "-"
    if not shrink:
        measured = measurer.height(area)
        if measured > 0:
            need, by = measured, "Excel"
        else:
            _, lines = b._wrap_shortfall(area, size)
            need, by = lines * size * b.LINE_HEIGHT_RATIO, "估算"
        if b._width_only(card, items):
            # Excel 量着放得下、打印时最后几个字母被挤到下一行：至少多给一行
            need, by = max(need, others + orig + size * b.LINE_HEIGHT_RATIO), by + "+一行"
    target = round(min(autofix.MAX_ROW_HEIGHT, max(orig, need - others) + size * ratio), 2)
    entry = rows.setdefault((sheet.Name, last), {"sheet": sheet, "orig": orig, "current": current,
                                                 "target": 0.0, "cards": [], "key": row_key})
    entry["target"] = max(entry["target"], target)
    entry["cards"].append(card["id"])
    debug_log.event("📏 工具箱算行高", 格=f"{sheet.Name}!{card['cell_addr']}", 行=last, 原行高=orig, 现在=current,
                    要=round(need, 2), 量高=by, 字号=size, 倍数=ratio, 目标=target, 开换行="是" if wrapped else "否")
    return {"done": f"行高 ×{ratio:g}", "row": last, "orig": round(orig, 2), "wrap": wrapped}


def _add_break(sheet: Any, area: Any) -> dict:
    row = int(area.Row)
    if row <= 1:
        return {"skip": "这格在第 1 行，上面没法分页"}
    setup = sheet.PageSetup
    if setup.Zoom is False and setup.FitToPagesTall is not False:
        return {"skip": "这张表页高定成了固定页数（取整），分页线不起作用：先在调整页面把页高改回自动"}
    try:
        for i in range(1, int(sheet.HorizontalPageBreaks.Count) + 1):
            if int(sheet.HorizontalPageBreaks(i).Location.Row) == row:
                return {"skip": "这格上面已经有分页线"}
    except Exception as err:
        debug_log.event("⚠️ 读不出分页线", 表=sheet.Name, 原因=str(err)[:120])
    sheet.HorizontalPageBreaks.Add(Before=sheet.Cells(row, 1))
    return {"done": "上面加分页线", "row": row}


def apply_tools(excel: Any, path: str, plan: dict[str, dict]) -> dict:
    """按选好的工具改一份文件、保存一次。返回 {"cells": {卡片 id: 结果}, "pages": [{sheet, before, after}]}。
    文件在上次检测后被人改过时不改（格子位置可能变了），请先复检。"""
    name = Path(path).name
    with wf.STATE_LOCK:
        cards = {cid: dict(b._data()["cards"][cid]) for cid in plan if cid in b._data()["cards"]}
        items = {cid: b._items_at(b._data()["cards"][cid]) for cid in cards}
    b._check_file(path)
    before = wf._file_version(Path(path))
    if wf.STATE["checked_files"].get(path) != before:
        raise ValueError(f"{name} 上次检测后在 Excel 里改过，格子位置可能变了：请先复检这个文件，再用工具")
    debug_log.event("🛠 工具箱开始改", 文件=name, 格数=len(cards))
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0)
    print_layout.bind(book, path)
    results: dict[str, dict] = {}
    sheets: dict[str, Any] = {}
    pages_before: dict[str, int | None] = {}
    rows: dict[tuple[str, int], dict] = {}
    measurer = b._Measurer(excel)
    try:
        if book.ReadOnly:
            raise ValueError(f"{name} 以只读方式打开，没能修改")
        shapes_before = shape_geometry.snapshot_book(book)  # 改行高前记下图形，保存前放回（用户 2026-11-11 要）
        for cid, card in cards.items():
            choice = plan[cid]
            where = f"{card['sheet_name']}!{card['cell_addr']}"
            try:
                sheet = b._find_sheet(book, card["sheet_name"])
                if sheet is None:
                    raise ValueError("找不到这张表")
                if sheet.Name not in sheets:
                    sheets[sheet.Name] = sheet
                    pages_before[sheet.Name] = _page_count(sheet)
                area = sheet.Range(card["cell_addr"]).MergeArea
                if choice["tool"] == "grow":
                    result = _grow_entry(measurer, sheet, area, card, items[cid], choice["ratio"], rows, path)
                elif choice["tool"] == "shrink":
                    area.WrapText = False  # 开着换行 Excel 不缩小字体
                    area.ShrinkToFit = True
                    result = {"done": "缩小字体填充"}
                else:
                    result = _add_break(sheet, area)
            except Exception as err:
                result = {"error": str(err)}
                wf.log_msg(f"⚠️ {name} / {where} 没能修改：{err}")
            result["where"] = where
            results[cid] = result
            debug_log.event("🛠 工具", 格=where, 工具=tool_label(choice["tool"], choice.get("ratio")),
                            结果=result.get("done") or result.get("skip") or result.get("error"))
        for (sheet_name, row), entry in rows.items():
            applied = entry["target"]
            entry["sheet"].Rows(row).RowHeight = applied
            with wf.STATE_LOCK:
                _fix_rows()[entry["key"]] = {"orig": entry["orig"], "applied": applied}
            debug_log.event("📏 工具箱改行高", 表=sheet_name, 行=row, 原行高=entry["orig"], 改前=entry["current"],
                            改后=applied)
            for cid in entry["cards"]:
                results[cid].update(before=entry["current"], applied=applied)
        changed = any("done" in r for r in results.values())
        pages = []
        if changed:
            for sheet_name, sheet in sheets.items():
                after = _page_count(sheet)
                if pages_before.get(sheet_name) is not None and after is not None and after != pages_before[sheet_name]:
                    pages.append({"sheet": sheet_name, "before": pages_before[sheet_name], "after": after})
            shape_geometry.restore_book(book, shapes_before, f"工具箱 {name}")
            normalize_workbook_view(excel, book)
            book.Save()
            b._followed_own_save(path, before)
            wf.log_msg(f"🛠 {name}：{sum('done' in r for r in results.values())} 格已按工具修改，已保存")
            debug_log.event("💾 工具箱已保存", 文件=name, 页数变化=pages or "-")
    finally:
        measurer.close()
        book.Close(SaveChanges=False)
    return {"cells": results, "pages": pages, "changed": changed, "sheets": list(sheets)}


def _remember(path: str, plan: dict[str, dict], cells: dict[str, dict]) -> None:
    """调用方持有 STATE_LOCK：记下每格用过的工具，复检后没过的卡上显示"""
    fixes = _fixes()
    for cid, result in cells.items():
        if "done" in result:
            choice = plan[cid]
            fixes[cid] = {"tool": choice["tool"], "ratio": choice.get("ratio"),
                          "label": tool_label(choice["tool"], choice.get("ratio"))}


def _after_recheck(path: str, plan: dict[str, dict]) -> tuple[int, int]:
    """调用方持有 STATE_LOCK：复检后数一数用了工具的格子过了几个，没过的写上用过什么"""
    passed = left = 0
    cards = b._data()["cards"]
    for cid in plan:
        card, fix = cards.get(cid), _fixes().get(cid)
        if card is None or fix is None:
            continue
        if card.get("dismissed"):
            passed += 1
        else:
            left += 1
            card["message"] = failed_message(fix)
            for item in b._items_at(card):
                item["review_reason"] = card["message"]
    return passed, left


def _fix_file(session: Any, path: str, plan: dict[str, dict]) -> dict:
    name = Path(path).name
    debug_log.begin(f"单元格工具箱 {name}（{len(plan)} 格）")
    unsaved = close_saved_workbooks([Path(path)])
    if Path(path) in unsaved:
        raise ValueError(f"{name}：{unsaved[Path(path)]}，请保存并关闭后再点统一修改")
    session.touch()
    root = wf.STATE.get("folder")
    adjusted = _adjusted_sheets(root, path)
    digest = sheet_history.fingerprint(Path(path)) if adjusted else None
    done = apply_tools(session.app(), path, plan)
    report = {"name": workbook_rel_name(Path(path), wf.STATE["folder"]), "pages": done["pages"],
              "done": sum("done" in r for r in done["cells"].values()),
              "skipped": [{"where": r["where"], "reason": r["skip"]} for r in done["cells"].values() if "skip" in r],
              "errors": [{"where": r["where"], "reason": r["error"]} for r in done["cells"].values() if "error" in r],
              "passed": 0, "left": 0}
    with wf.STATE_LOCK:
        _remember(path, plan, done["cells"])
        wf._persist_state()
    if not done["changed"]:
        return report
    with wf.STATE_LOCK:
        cards = [b._data()["cards"][cid] for cid in plan if cid in b._data()["cards"]]
    changed_sheets = sorted(set(done.get("sheets") or []) & adjusted)
    if changed_sheets:
        session.touch()
        report["pdfs"] = _reexport_adjusted(session, root, path, changed_sheets, digest)
    session.touch()
    b._recheck_file(session.app(), path, cards, passed_message=PASSED_MESSAGE)
    with wf.STATE_LOCK:
        report["passed"], report["left"] = _after_recheck(path, plan)
        wf._persist_state()
    b._sync_highlights_safely(session, path)
    b._forget_opened(path)
    debug_log.event("🛠 工具箱复检完", 文件=name, 通过=report["passed"], 没过=report["left"],
                    跳过=len(report["skipped"]), 出错=len(report["errors"]))
    return report


def _adjusted_sheets(root: Any, path: str) -> set[str]:
    """调整页面给这份 Excel 导出过 PDF 的表（PDFS/PDF导出记录.json）；没导出过返回空"""
    if not root:
        return set()
    try:
        return set(pdf_store.sheet_pdfs(root, Path(path)))
    except Exception as err:
        debug_log.event("⚠️ 读调整页面的 PDF 记录失败", 文件=Path(path).name, 错误=str(err))
        return set()


def _reexport_adjusted(session: Any, root: str, path: str, sheets: list[str], digest: str | None) -> int:
    """改了格子的表在调整页面也有 PDF：复检前先把这几张表的 PDF 重新导出（和「🔄 重新生成」一样，按表来），
    复检直接用这份新的 PDF 检测、不再另导一次，回到调整页面看到的就是改完的样子、不用整份重新导出
    （用户 2026-11-09 把切字并进调整页面）。导不出来不影响复检，调整页面打开时会按「Excel 改过」自动重新导出"""
    started = time.perf_counter()
    try:
        paths = wf._export_workbook_pdfs(session.app, Path(path), Path(root), set(), force=True,
                                         only_sheets=set(sheets), fresh_before=digest)
    except Exception as err:
        debug_log.exception(err, f"单元格工具箱改完重新导出调整页面的 PDF {Path(path).name}")
        wf.log_msg(f"⚠️ {Path(path).name} 改完没能重新导出调整页面的 PDF，打开调整页面时会自动再导：{err}")
        return 0
    debug_log.event("📄 工具箱改完重新导出 PDF", 文件=Path(path).name, 表=sheets, PDF=len(paths),
                    秒=round(time.perf_counter() - started, 2))
    return len(paths)


def run(job: dict, session: Any, watchdog: Any) -> list[str]:
    """batch.run 里 kind == "batch_fix" 时调用：一份一份改、保存、复检。出错重启 Excel 再试一次，
    还不行（文件打不开、被占用等）问用户重试还是跳过；跳过的可「继续上次任务」再做"""
    plan_all: dict[str, dict] = job["options"]["tools"]
    expected = job["workbooks"]
    with wf.STATE_LOCK:
        results = job.setdefault("result", {})
    wf.log_msg(f"🛠 开始统一修改：{len(plan_all)} 格，{len(expected)} 份文件")
    for i, path in enumerate(expected, 1):
        if path in job["completed"]:
            continue
        if job.get("pause_requested"):
            break
        with wf.STATE_LOCK:
            plan = {cid: choice for cid, choice in plan_all.items()
                    if (card := b._data()["cards"].get(cid)) and card["wb_path"] == path and not card.get("dismissed")}
        label = workbook_rel_name(Path(path), wf.STATE["folder"])
        if not plan:
            with wf.STATE_LOCK:
                job["completed"].append(path)
            continue
        wf.log_msg(f"🛠 正在修改（{i} / {len(expected)}）：{Path(path).name}")
        tries = 0
        while True:
            try:
                watchdog.progress()
                report = _fix_file(session, path, plan)
                watchdog.progress()
                with wf.STATE_LOCK:
                    results[path] = report
                    job["completed"].append(path)
                    job.get("file_errors", {}).pop(path, None)
                    wf._persist_state()
                break
            except Exception as err:
                if not isinstance(err, (ValueError, TimeoutError)) and not watchdog.timed_out and tries < 1:
                    tries += 1
                    wf.log_msg(f"⚠️ {Path(path).name} 出错，重启 Excel 后重试：{err}{b._where(err)}")
                    b._restart_excel(session)
                    continue
                session.close()
                message = file_problem.message_for(err, watchdog)
                if file_problem.should_ask(err, message) and not b._stopping(job):
                    wf.log_msg(f"⏸ {Path(path).name} 出错，等你选择重试还是跳过：{message}")
                    update_file_progress(label, status="processing", error=message,
                                         current_action="出错了，等你选择重试还是跳过")
                    choice = file_problem.ask(path, label, message, folder=wf.STATE["folder"],
                                              tick=watchdog.progress, should_stop=lambda: b._stopping(job))
                    watchdog.progress()
                    if choice == "retry":
                        wf.log_msg(f"🔁 重试 {Path(path).name}")
                        tries = 0
                        continue
                wf.log_msg(f"❌ {Path(path).name} 没改成，已跳过，可继续上次任务：{message}")
                with wf.STATE_LOCK:
                    job.setdefault("file_errors", {})[path] = message
                    results[path] = {"name": label, "error": message, "pages": [], "done": 0, "skipped": [],
                                     "errors": [], "passed": 0, "left": 0}
                    wf._persist_state()
                break
    total_passed = sum(r.get("passed", 0) for r in results.values())
    total_left = sum(r.get("left", 0) for r in results.values())
    wf.log_msg(f"🛠 统一修改完成：复检通过 {total_passed} 格，还没过 {total_left} 格")
    return expected
