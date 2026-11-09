"""调整页面里的切字（单元格问题）：用户 2026-11-09 定，把「查看切字」并进调整页面，页面和单元格在一个页面里处理。

- 查切字还是原来的「扫描并标记」（batch_scan）：调整页面导出的 PDF 只要 Excel 没再改过就直接拿来检测（#197），
  所以进页面导完 PDF 接着在后台查切字，不用再导一遍。网页打开时有没查过、查过以后改过的文件就自动开始查。
- 查出来的格子还是工作台的卡片（data["cards"]），这里按「哪张表、第几页、页面上哪一块」整理给网页：
  小图和大图上画红框，每张表下面列出这张表的格子和工具（行高 ×1.2/1.5/2、缩小字体、上面加分页线）。
- 「🛠 统一修改」先改页面（pdf_review.process），再：标成「不用处理」的格子提交成蓝（记下忽略、去掉标色）
  → 页面改过的文件重新查切字 → 按选好的工具改格子、保存、重新导出这几张表的 PDF、复检（cell_toolbox）。
"""
from __future__ import annotations

import os
import threading
from pathlib import Path
from typing import Any

from core import autofix, debug_log, folder_progress, workflow as wf
from core import batch as b
from core.common import workbook_rel_name

GROW_RATIOS = (1.2, 1.5, 2.0)
CELL_TOOLS = ("grow", "shrink", "break")


def _norm(path: Any) -> str:
    try:
        return os.path.normcase(str(Path(str(path)).resolve()))
    except OSError:
        return os.path.normcase(str(path))


def _same_folder(a: str, b_: str) -> bool:
    return bool(a) and bool(b_) and folder_progress.key(a) == folder_progress.key(b_)


def _scan_state() -> dict:
    """后台在不在查切字（调用方持有 STATE_LOCK）"""
    job = wf.STATE.get("job") or {}
    kind = job.get("kind", "")
    running = bool(wf.STATE.get("busy")) and kind == "batch_scan"
    workbooks = job.get("workbooks") or []
    completed = set(job.get("completed") or [])
    current = next((p for p in workbooks if p not in completed), "") if running else ""
    # 网页实时进度框（static/job_panel.js）：每份查完的结果。扫描时只数新查完的文件的卡片，不拖慢扫描
    from core import batch
    feed = batch._job_feed(job, light=True) if kind == "batch_scan" else {"log": [], "step": ""}
    return {"busy": bool(wf.STATE.get("busy")), "kind": kind, "running": running,
            "log": feed["log"], "step": feed.get("step", ""),
            "paused": kind == "batch_scan" and job.get("status") == "interrupted",
            "done": len(completed & set(workbooks)), "total": len(workbooks),
            "current": workbook_rel_name(Path(current), wf.STATE.get("folder") or Path(current).parent) if current else ""}


def _file_status(path: str) -> str:
    """ok 查过、之后没改过 / changed 查过以后改过（要重查）/ new 还没查过（调用方持有 STATE_LOCK）"""
    checked = wf.STATE["checked_files"].get(path)
    if not checked:
        return "new"
    try:
        return "ok" if wf._file_version(Path(path)) == checked else "changed"
    except OSError:
        return "changed"


def _boxes(items: list[dict], positions_ok: bool) -> list[dict]:
    """这格在 PDF 上的位置：[{pdf, page, pages, box: [左, 上, 右, 下]（占这一页宽高的比例）}]（调用方持有 STATE_LOCK）。
    检测以后 Excel 改过就不给位置（PDF 和格子对不上了），网页上只列格子、不画框"""
    if not positions_ok:
        return []
    found = []
    seen = set()
    for item in items:
        bbox = item.get("bbox")
        page = wf.STATE["pages"].get(item.get("page_id")) or {}
        width, height = page.get("page_width"), page.get("page_height")
        if (not isinstance(bbox, (list, tuple)) or len(bbox) != 4 or not item.get("pdf_path")
                or not isinstance(width, (int, float)) or not isinstance(height, (int, float)) or width <= 0 or height <= 0):
            continue
        try:
            box = [max(0.0, min(1.0, float(v) / (width if i % 2 == 0 else height))) for i, v in enumerate(bbox)]
        except (TypeError, ValueError):
            continue
        if box[2] <= box[0] or box[3] <= box[1]:
            continue
        key = (_norm(item["pdf_path"]), item.get("page_index"), tuple(round(v, 4) for v in box))
        if key in seen:
            continue
        seen.add(key)
        found.append({"pdf": key[0], "page": int(item.get("page_index") or 0), "pages": page.get("total_pages"),
                      "box": [round(v, 5) for v in box]})
    return found


def snapshot(folder: str) -> dict:
    """网页轮询：这个文件夹查出来的切字格子 + 每份 Excel 查没查过（files 里没有的就是还没查过）+ 后台查切字的进度。
    工作台现在显示的不是这个文件夹时 active=False"""
    from core import cell_toolbox
    root = str(Path(folder.strip().strip('"').strip("'")).expanduser().resolve())
    with wf.STATE_LOCK:
        scan = _scan_state()
        active = _same_folder(wf.STATE.get("folder", ""), root)
        if not active:
            return {"active": False, "folder": root, "scan": scan, "cards": [], "files": {}}
        statuses = {path: _file_status(path) for path in wf.STATE["checked_files"] if wf.in_current_folder(path)}
        fixes = b._data().get("cell_fixes") or {}
        by_cell: dict[tuple, list[dict]] = {}  # 按格子先分好，不用每张卡都把全部检测项过一遍（扫描时网页也在轮询）
        for item in wf.STATE["items"].values():
            by_cell.setdefault(b._cell(item), []).append(item)
        cards = []
        positions: dict[str, bool] = {}
        for card in sorted(b._data()["cards"].values(), key=lambda c: c["seq"]):
            if card.get("dismissed") or card.get("color") != "yellow" or not card.get("finalized"):
                continue
            path = card["wb_path"]
            if not wf.in_current_folder(path):
                continue
            if path not in positions:
                positions[path] = _file_status(path) == "ok"
            items = by_cell.get(b._cell(card), [])
            view = dict(card)
            view["error_types"] = b._card_error_types(card)
            cat_key, cat_label = b._card_category(view, items)
            fix = fixes.get(card["id"])
            tools = cell_toolbox.allowed_tools(view)
            grows = set(view["error_types"]) & set(b.GROW_ERROR_TYPES)
            cards.append({
                "id": card["id"], "wb_path": path, "sheet": card["sheet_name"], "cell": card.get("cell_addr", ""),
                "text": str(card.get("text") or "")[:80], "type": view["error_types"][0],
                "type_label": autofix.error_type_label(view["error_types"][0]),
                "types": [autofix.error_type_label(t) for t in view["error_types"]],
                "category": cat_label, "category_key": cat_key, "tools": tools,
                "default_tool": "grow:1.5" if not fix and "grow" in tools and grows else "none",
                "cell_fix": fix.get("label") if fix else "", "message": card.get("message", ""),
                "boxes": _boxes(items, positions[path]),
            })
    return {"active": True, "folder": root, "scan": scan, "cards": cards, "files": statuses}


# ---------- 进页面自动查切字 ----------

def _queue(kind: str, options: dict) -> dict:
    """占住任务槽排上一个工作台任务（和网页 /api/batch/... 一样），返回这个任务。占不上抛 ValueError"""
    autofix.release_excel()
    try:
        queued = wf.queue_job(kind, options)
    finally:
        autofix.resume_after_release()
    if not queued:
        raise ValueError("后台有别的任务正在执行，请等它结束")
    with wf.STATE_LOCK:
        return wf.STATE["job"]


def run_job(kind: str, options: dict) -> dict:
    """在当前线程里跑完一个工作台任务（统一修改里一步接一步做），返回这个任务（看 completed、result）"""
    job = _queue(kind, options)
    wf.run_pending_workflow()
    return job


def needs_check(folder: str) -> list[str]:
    """这个文件夹里还没查过、查过以后又改过的 Excel（工作台得是这个文件夹，调用方先 activate）"""
    _root, workbooks = wf.list_workbooks(folder)
    with wf.STATE_LOCK:
        return [str(p) for p in workbooks if _file_status(str(p)) != "ok"]


def start_check(folder: str, highlight: Any = None) -> dict:
    """打开调整页面、PDF 都是最新的以后：有没查过 / 改过的 Excel 就在后台查切字（扫描并标记，增量：没改过的跳过）。
    后台正忙、有暂停的任务、没要查的都不开始，返回原因"""
    root = str(Path(folder.strip().strip('"').strip("'")).expanduser().resolve())
    style = b.highlight_style(highlight)
    with wf.STATE_LOCK:
        if wf.STATE.get("busy") or wf.STATE.get("maintenance"):
            return {"started": False, "busy": "后台有别的任务正在执行，做完再查切字"}
        job = wf.STATE.get("job") or {}
        if job.get("status") == "interrupted" and (job.get("pending_adjustment") or job.get("deferred_adjustments")):
            return {"started": False, "busy": "工作台上还有没写完的修改，请先到「查看切字」继续或放弃"}
    try:
        wf.activate_folder(root)
    except ValueError as err:
        return {"started": False, "busy": str(err)}
    todo = needs_check(root)
    if not todo:
        return {"started": False, "files": 0}
    try:
        _queue("batch_scan", {"folder": root, "highlight": style})
    except ValueError as err:
        return {"started": False, "busy": str(err)}
    debug_log.event("🔍 调整页面：自动查切字", 文件夹=root, 要查=len(todo))
    threading.Thread(target=wf.run_pending_workflow, name="调整页面查切字", daemon=True).start()
    return {"started": True, "files": len(todo)}


# ---------- 统一修改里的单元格部分 ----------

def validate(cells: Any, skips: Any) -> tuple[dict[str, dict], list[str]]:
    """cells：[{card_id, tool, ratio?}]；skips：[卡片 id]（标成「不用处理」）。只看格式，卡片在不在等做的时候再看
    （先改页面、重新查切字以后，同一格还有问题才改）"""
    plan: dict[str, dict] = {}
    for raw in cells or []:
        if not isinstance(raw, dict) or not isinstance(raw.get("card_id"), str) or raw.get("tool") not in CELL_TOOLS:
            raise ValueError("格子的工具不对，请刷新页面再试")
        choice: dict[str, Any] = {"tool": raw["tool"]}
        if raw["tool"] == "grow":
            ratio = raw.get("ratio")
            if isinstance(ratio, bool) or not isinstance(ratio, (int, float)) or float(ratio) not in GROW_RATIOS:
                raise ValueError("行高倍数只能选 1.2、1.5 或 2")
            choice["ratio"] = float(ratio)
        plan[raw["card_id"]] = choice
    skip = []
    for cid in skips or []:
        if not isinstance(cid, str):
            raise ValueError("「不用处理」的格子不对，请刷新页面再试")
        if cid not in skip:
            skip.append(cid)
    return {cid: c for cid, c in plan.items() if cid not in skip}, skip


def _live(ids: list[str]) -> list[str]:
    """还在的、黄色待处理的卡（调用方持有 STATE_LOCK）"""
    cards = b._data()["cards"]
    return [cid for cid in ids if (card := cards.get(cid)) and not card.get("dismissed")
            and card.get("color") == "yellow" and card.get("finalized") and wf.in_current_folder(card["wb_path"])]


def run_steps(root: str, plan: dict[str, dict], skips: list[str], highlight: Any, set_step: Any, log: Any,
              stage: Any) -> dict:
    """pdf_review.process 的后台线程里调用（页面已经改完、任务槽已经放开）。一步接一步跑工作台任务：
    不用处理 → 页面改过的重新查切字 → 改格子。返回 {"skipped": n, "fixed": {...}}，每步结果用 log() 写进进度框"""
    from core import cell_toolbox
    out: dict[str, Any] = {"skipped": 0, "rechecked": 0, "files": []}
    wf.activate_folder(root)
    with wf.STATE_LOCK:
        skip_ids = _live(skips)
    if skip_ids:
        stage("skip")
        set_step(f"把 {len(skip_ids)} 格标成「不用处理」：记下来、去掉 Excel 里的标色")
        b.mark_many(skip_ids, "blue")
        job = run_job("batch_commit", {"card_ids": skip_ids, "color": "blue"})
        with wf.STATE_LOCK:
            done = sum(1 for cid in skip_ids if (b._data()["cards"].get(cid) or {}).get("dismissed"))
        out["skipped"] = done
        log([{"pdf": "", "name": f"{done} 格不用处理", "sheet": "",
              "reason": "记下了，以后不再查出来；Excel 里的标色已去掉" if done == len(skip_ids)
              else f"还有 {len(skip_ids) - done} 格没提交成（{job.get('status')}），可以再点一次"}],
            done == len(skip_ids), "cell")
    # 页面改过（取整、改方向……）、之前查过切字的文件：格子和分页可能都变了，先重新查
    changed = [p for p in needs_check(root)]
    with wf.STATE_LOCK:
        changed = [p for p in changed if p in wf.STATE["checked_files"]]
    if changed:
        stage("scan")
        set_step(f"{len(changed)} 份 Excel 的页面改过：重新查切字（没改过的跳过）")
        run_job("batch_scan", {"folder": root, "highlight": b.highlight_style(highlight)})
        out["rechecked"] = len(changed)
    with wf.STATE_LOCK:
        live = set(_live(list(plan)))
        cards = b._data()["cards"]
        todo = {cid: choice for cid, choice in plan.items()
                if cid in live and choice["tool"] in cell_toolbox.allowed_tools(cards[cid])}
    gone = len(plan) - len(todo)
    if plan and gone:
        log([{"pdf": "", "name": f"{gone} 格没改", "sheet": "",
              "reason": "改完页面重新查过，这几格已经没问题了（或者问题变了、这个工具用不了），没动"}], True, "cell")
    if todo:
        stage("cells")
        set_step(f"按选好的工具改 {len(todo)} 格：改 Excel、保存、重新生成这几张表的 PDF、复检")
        job = run_job("batch_fix", cell_toolbox.job_options(todo))
        for path, report in (job.get("result") or {}).items():
            out["files"].append(path)
            if report.get("error"):
                log([{"pdf": "", "name": report.get("name") or Path(path).name, "sheet": "", "reason": report["error"]}],
                    False, "cell")
                continue
            parts = [f"改了 {report.get('done', 0)} 格，复检通过 {report.get('passed', 0)} 格"]
            if report.get("left"):
                parts.append(f"还没过 {report['left']} 格（换个工具再改，或打开 Excel 自己改）")
            for page in report.get("pages") or []:
                parts.append(f"[{page['sheet']}] {page['before']} 页 → {page['after']} 页")
            for item in (report.get("skipped") or []) + (report.get("errors") or []):
                parts.append(f"{item['where']}：{item['reason']}")
            ok = not report.get("left") and not report.get("errors")
            log([{"pdf": "", "name": report.get("name") or Path(path).name, "sheet": "", "reason": "；".join(parts)}],
                ok, "cell")
    return out
