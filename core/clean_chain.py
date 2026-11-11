"""精简 Sheet 之后自动接着做的例行检查：排查 BL 超链接 → 修复引用错误 →（可选）修复隐藏页 → 统一字体。

用户 2026-11-07 定：主页只留「精简 Sheet」一站，原来第 4、5、6 步不用再单独点。
- 按文件执行，当前工作簿打开一次；各项共用独立工具的检查和写入规则。
- 能安全跳过的检查错误记录后继续；写入失败关闭工作簿并放弃这份尚未保存的修改。
- 结果存进数据库，重开程序也能在精简 Sheet 页看到；原来的页面（主页「更多工具」）还能看每一步的详细结果。
引用错误和隐藏页要对照参考（原始备份）文件夹，没填就跳过这两步并写明原因。
"""
from __future__ import annotations

import copy
import threading
from pathlib import Path
from typing import Any, Callable

from core import bl_links, debug_log, file_problem, font_fix, hidden_repair, link_repair, operation_history
from core import sheet_history, storage, workflow as wf
from core.backup_compare import pair_files, require_manual_file_check
from core.common import workbook_locked, workbook_rel_name
from core.excel_ops import open_workbook_checked, start_background_excel
from core.excel_watchdog import NOT_DONE_REASON, TaskStopped
from core.reference_folders import ReferenceFolders, prepare as prepare_references, unmatched

SNAPSHOT = "clean_chain"
MAX_PROBLEMS = 300  # 每一步最多列这么多处「改不了 / 出错」，太多时只写总数

STEPS = (
    ("bl", "排查 BL 超链接"),
    ("links", "修复引用错误"),
    ("hidden", "修复隐藏页"),
    ("font", "统一字体"),  # 游ゴシック 换成ＭＳ Ｐゴシック（用户 2026-11-10 要），切字误报多半是它
)

_LOCK = threading.Lock()
_STATUS: dict[str, Any] = {}
_LOADED = False


def _load() -> None:
    global _LOADED
    if _LOADED:
        return
    _LOADED = True
    try:
        saved = storage.load_snapshot(SNAPSHOT)
    except Exception as err:
        debug_log.event("⚠️ 读不到上次的自动检查结果", 原因=str(err))
        saved = None
    if isinstance(saved, dict) and saved.get("steps"):
        if saved.get("running"):  # 上次程序在检查途中关掉了
            saved["running"] = False
            saved["interrupted"] = True
            for step in saved["steps"]:
                step.pop("progress", None)
                if step["state"] in {"waiting", "running"}:
                    step.update(state="failed", text="程序关掉了，这一步没做完")
        _STATUS.update(saved)


def _save() -> None:
    try:
        storage.save_snapshot(SNAPSHOT, copy.deepcopy(_STATUS))
    except Exception as err:
        debug_log.event("⚠️ 自动检查结果没能存进数据库", 原因=str(err))


OPTIONAL_STEPS = ("bl", "links", "hidden", "font")  # 网页上能勾掉的几步；hidden 默认不做，其他默认做


def options_from(payload: Any) -> dict | None:
    """网页传来的 chain：{"hidden": bool, "old_folder": str, "bl"/"links"/"font": bool}；没传就不自动检查（旧页面、单项调用）。
    主页工具箱的精简卡上可以勾「顺便做」哪几项（用户 2026-11-10 要）；旧页面没传 bl/links/font 时照旧都做"""
    if payload is None:
        return None
    if not isinstance(payload, dict):
        raise ValueError("自动检查选项格式不正确")
    old_folder = payload.get("old_folder", "")
    flags = {key: payload.get(key, key != "hidden") for key in OPTIONAL_STEPS}
    if not isinstance(old_folder, str) or any(type(value) is not bool for value in flags.values()):
        raise ValueError("自动检查选项格式不正确")
    options = {**flags, "old_folder": old_folder.strip().strip('"').strip("'")}
    if "reference_folders" in payload:
        options["reference_folders"] = payload["reference_folders"]
    return options


def _enabled(chain: dict, key: str) -> bool:
    return bool(chain.get(key, key != "hidden"))


def status() -> dict:
    with _LOCK:
        _load()
        return copy.deepcopy(_STATUS)


def _step(key: str) -> dict:
    return next(step for step in _STATUS["steps"] if step["key"] == key)


def _update(key: str, **fields: Any) -> None:
    with _LOCK:
        _step(key).update(fields)
        _save()


_chain_stop_seq = -1  # 这一串开始时「结束整个任务」点过几次：之后又点了就不再接着做后面几步


def _stopped() -> bool:
    from core import excel_watchdog
    return _chain_stop_seq >= 0 and excel_watchdog.stop_seq() != _chain_stop_seq


def _mark_stopped(key: str) -> None:
    with _LOCK:
        _step(key).pop("progress", None)
    _update(key, state="skipped", text="你点了「结束整个任务」，这一步没做")


def prepare(folder: str, chain: dict) -> None:
    """点了「确认执行」：精简还在跑，先把后面几步排上（页面马上能显示「等精简完成」）"""
    global _chain_stop_seq
    from core import excel_watchdog
    _chain_stop_seq = excel_watchdog.stop_seq()
    with _LOCK:
        _load()
        if _STATUS.get("running") and _STATUS.get("stage") == "chain":
            raise ValueError("上次精简后的自动检查还没做完，请等它结束")
        _STATUS.clear()
        _STATUS.update(running=True, folder=str(folder), old_folder=chain.get("old_folder", ""),
                       hidden=bool(chain.get("hidden")), stage="clean", interrupted=False,
                       steps=[{"key": key, "label": label, "state": "waiting", "text": "等精简 Sheet 做完",
                               "fixed": 0, "problems": [], "problem_total": 0, "error_total": 0}
                              for key, label in STEPS])
        for key in OPTIONAL_STEPS:
            if not _enabled(chain, key):
                _step(key).update(state="off", text="这次没勾选")
        _save()


def drop(folder: str) -> None:
    """精简任务没能开始（已有任务在跑）：刚排上的自动检查撤掉"""
    with _LOCK:
        if _STATUS.get("folder") == str(folder) and _STATUS.get("stage") == "clean":
            _STATUS.clear()
            _save()


def _problem(file: dict, item: dict | None, reason: str, *, error: bool = False) -> dict:
    found = {"file": file.get("name", ""), "path": file.get("path", ""), "reason": reason, "error": error}
    if file.get("old"):  # 配对到的参考文件：点路径时和当前文件一起打开，人工对照
        found["old"] = file["old"]
    if item:
        found.update(sheet=item.get("sheet", ""), cell=item.get("anchor") or item.get("addr", ""),
                     addr=item.get("addr", ""))
        if file.get("old"):
            found["old_sheet"] = item.get("old_sheet", item.get("sheet", ""))
        if item.get("old_sheet") == "" and item.get("source"):
            found["reason"] = f"{reason}。{item['source']}"
    return found


def _finish_step(key: str, fixed: int, problems: list[dict], text: str) -> None:
    errors = sum(1 for p in problems if p["error"])
    state = "failed" if errors else "warn" if problems else "done"
    with _LOCK:
        _step(key).pop("progress", None)
    _update(key, state=state, fixed=fixed, problems=problems[:MAX_PROBLEMS], problem_total=len(problems),
            error_total=errors, text=text)


def _prepare_steps(folder: str, chain: dict) -> tuple[dict, dict, dict]:
    paths_by_step, targets, references = {}, {}, {}
    for key, _ in STEPS:
        if not _enabled(chain, key):
            continue
        if _stopped():
            _mark_stopped(key)
            continue
        try:
            reference_step = key in {"links", "hidden"}
            if reference_step:
                options = (ReferenceFolders.from_payload(chain) if "reference_folders" in chain
                           else chain.get("old_folder", ""))
                if not options:
                    _update(key, state="skipped", text="没填参考（原始备份）文件夹，跳过这一步")
                    continue
            list_books = hidden_repair.list_workbooks if key in {"hidden", "font"} else wf.list_workbooks
            root, paths = list_books(folder)
            if reference_step:
                old_root, old_paths = prepare_references(
                    root, options, list_books, wf.EXCEL_EXTENSIONS | {".xlsb"}, disjoint=key == "hidden")
                inferred: set[str] = set()
                names = [workbook_rel_name(path, root) for path in paths]
                references[key] = old_root, pair_files(old_root, old_paths, names, root, inferred=inferred), inferred
            paths_by_step[key] = paths
            for path in paths:
                targets.setdefault(path, []).append(key)
        except Exception as error:
            debug_log.exception(error, f"精简后自动检查 {key}")
            _finish_step(key, 0, [_problem({}, None, str(error), error=True)], f"没能开始：{error}")
    return paths_by_step, targets, references


def _plan_step(key: str, entry: dict, book: Any, path: Path, entries: dict,
               references: dict, get_reference_excel: Callable[[], Any]) -> list[dict] | dict | None:
    if key == "bl":
        entry["hits"] = bl_links.scan_workbook(book)
        return entry["hits"] or None
    if key == "font":
        cells = font_fix.yu_cells(path)
        if path.suffix.lower() in {".xlsx", ".xlsm"}:
            cells = font_fix.refresh_changed_cells(
                book, cells, entries.get("bl", {}).get("cells", []), entries.get("links", {}).get("cells", []))
        return cells  # 即使没有格子也需要检查图形

    if key == "links":
        entry.update(link_repair.scan_workbook(book))
        if not entry["hits"]:
            return None
        entry["cells"] = [{**hit, "status": "skip", "reason": "未修复"} for hit in entry["hits"]]
    old_root, pairs, inferred = references[key]
    old_path = pairs[entry["name"]]
    if old_path is None:
        raise ValueError(unmatched(old_root, entry["name"], "找不到唯一对应的备份文件，请选到对应子文件夹"))
    entry["old"] = str(old_path)
    if key == "links":
        entry.update(link_repair.preview_file(get_reference_excel(), entry, old_path))
    else:
        entry.update(hidden_repair.plan_file(get_reference_excel(), path, old_path, hidden_repair.DEFAULT_OPTIONS, book=book))
    if entry["name"] in inferred:
        require_manual_file_check(entry)
        return None
    return [cell for cell in entry["cells"] if cell["action"] == "restore"] or None


def _apply_step(key: str, entry: dict, book: Any, ready: list[dict] | dict, guard: file_problem.Guard) -> bool:
    if key == "font":
        entry.update(font_fix.fix_open_workbook(book, ready, lambda *_: guard.progress()))
        return bool(entry["sheets"])
    if key == "bl":
        entry["cells"] = bl_links.clear_workbook(book, ready)
        return bool(entry["cells"])

    if key == "links":
        cells = link_repair.apply_workbook(book, ready)
    else:
        cells = hidden_repair.apply_workbook(book, ready, entry["checks"])
    outcomes = {(cell["sheet"], cell.get("kind"), cell["addr"]): cell for cell in cells}
    entry["cells"] = [outcomes.get((cell["sheet"], cell.get("kind"), cell["addr"]), cell) for cell in entry["cells"]]
    return any(cell["status"] == "ok" for cell in cells)


def _check_stop(guard: file_problem.Guard) -> None:
    if _stopped() or guard.watchdog.skip_requested or guard.watchdog.stop_requested:
        raise TaskStopped(NOT_DONE_REASON)


def _run_workbook(excel: Any, path: Path, entries: dict, references: dict,
                  get_reference_excel: Callable[[], Any], guard: file_problem.Guard,
                  progress: Callable[[str], None], close_excel: Callable[[], None]) -> None:
    stamp = sheet_history.fingerprint(path)
    locked = workbook_locked(path)
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0, ReadOnly=locked)
    try:
        if path.suffix.lower() == ".xls":
            # 旧版 xls 取得写锁时 Excel 会更新文件的 OLE 容器。
            stamp = sheet_history.fingerprint(path)
        modified = []
        for key, entry in entries.items():
            _check_stop(guard)
            progress(key)
            guard.progress()
            try:
                ready = _plan_step(key, entry, book, path, entries, references, get_reference_excel)
                if ready is not None and book.ReadOnly:
                    raise ValueError("文件正在 Excel 中打开或为只读，请关闭后再修复")
            except (OSError, ValueError) as error:
                entry["error"] = str(error)
                continue
            if ready is not None and _apply_step(key, entry, book, ready, guard):
                modified.append(entry)
        _check_stop(guard)
        if stamp != sheet_history.fingerprint(path):
            raise ValueError("当前文件发生变化，请重新检查")
        if modified:
            for entry in modified:
                if entry.get("old_fingerprint") and sheet_history.fingerprint(entry["old"]) != entry["old_fingerprint"]:
                    raise ValueError("备份在自动检查期间发生变化，请重新检查")
            book.Save()
            debug_log.saved_excel("精简后自动检查", path)
    finally:
        try:
            book.Close(SaveChanges=False)
        except Exception as error:
            debug_log.exception(error, f"关闭自动检查工作簿 {path.name}")
            close_excel()


def _process_files(folder: str, paths_by_step: dict, targets: dict, references: dict, results: dict) -> None:
    root = Path(folder)
    guard = file_problem.Guard("精简 Sheet 后自动检查")
    excel = reference_excel = None
    counts = dict.fromkeys(paths_by_step, 0)

    def get_reference_excel() -> Any:
        nonlocal reference_excel
        if reference_excel is None:
            # 当前和参考通常同名，需要两个 Excel 实例。参考模拟可以按候选重开。
            reference_excel = guard.watch(start_background_excel())
        return reference_excel

    def close_excel() -> None:
        nonlocal excel, reference_excel
        file_problem.quit_quietly(reference_excel)
        file_problem.quit_quietly(excel)
        excel = reference_excel = None

    initialized = False
    try:
        wf.pythoncom.CoInitialize()
        initialized = True
        for path, keys in targets.items():
            if _stopped() or guard.watchdog.stop_requested:
                break
            name = workbook_rel_name(path, root)

            def progress(key: str) -> None:
                with _LOCK:
                    for step in _STATUS["steps"]:
                        if step["state"] == "running":
                            step.update(state="waiting", text="等待检查下一份文件")
                    _step(key).update(state="running", text="正在自动检查…",
                                      progress={"phase": "检查", "done": counts[key],
                                                "total": len(paths_by_step[key]), "current": name})

            progress(keys[0])
            while True:
                entries = {key: {"name": name, "path": str(path), "cells": 0 if key == "font" else [], "error": None}
                           for key in keys}
                try:
                    if guard.file(name):
                        close_excel()
                    if excel is None:
                        excel = guard.watch(start_background_excel())
                    _run_workbook(excel, path, entries, references, get_reference_excel, guard, progress, close_excel)
                except Exception as error:
                    debug_log.exception(error, f"精简后自动检查 {name}")
                    retry, reason = guard.problem(error, path, name, root, close=close_excel)
                    if retry:
                        wf.log_msg(f"🔁 重试 {name}")
                        continue
                    for key, entry in entries.items():
                        if not entry["error"]:
                            entry["error"] = reason
                            if key == "font":
                                entry.update(shapes=0, cells=0, sheets=[])
                            else:
                                for cell in entry["cells"]:
                                    if cell.get("status") in {"ok", "resolved"}:
                                        cell.update(status="failed", reason=reason)
                break
            for key, entry in entries.items():
                results[key].append(entry)
                counts[key] += 1
            errors = [entry["error"] for entry in entries.values() if entry["error"]]
            wf.log_msg(f"自动检查 {name}：" + (errors[0] if errors else "已完成"))
    finally:
        close_excel()
        try:
            guard.stop()
        finally:
            if initialized:
                wf.pythoncom.CoUninitialize()


def _finish_results(key: str, files: list[dict]) -> None:
    problems = [_problem(file, None, file["error"], error=True) for file in files if file.get("error")]
    good = [file for file in files if not file.get("error")]
    if key == "font":
        shapes = sum(file.get("shapes", 0) for file in good)
        cells = sum(file.get("cells", 0) for file in good)
        changed = sum(1 for file in good if file.get("shapes") or file.get("cells"))
        parts = [f"{shapes} 个图形" if shapes else "", f"{cells} 个格子" if cells else ""]
        text = (f"{changed} 份 Excel 换成标准字体：{'、'.join(p for p in parts if p)}" if changed
                else "字体都是标准字体，没有要改的")
        _finish_step(key, shapes + cells, problems, text)
        return

    fixed = 0
    for file in good:
        for cell in file.get("cells", []):
            if cell.get("status") in {"ok", "resolved"}:
                fixed += 1
            else:
                problems.append(_problem(file, cell, cell.get("reason") or "未修复"))
    if key == "bl":
        text = f"清除了 {fixed} 处 BL 超链接" if fixed else "没有清除 BL 超链接"
    elif key == "links":
        skipped = sum(1 for problem in problems if not problem["error"])
        text = f"引用错误 {fixed + skipped} 处：修好 {fixed} 处，改不了 {skipped} 处" if fixed or skipped else "没有引用错误"
    else:
        text = f"恢复了 {fixed} 处隐藏表 / 列" if fixed else "没有要恢复的隐藏表 / 列"
    _finish_step(key, fixed, problems, text)


def _publish_details(folder: str, key: str, files: list[dict], references: dict) -> None:
    module = {"bl": bl_links, "links": link_repair, "hidden": hidden_repair, "font": font_fix}[key]
    data = {"folder": folder, "files": files}
    if key in references:
        data["old_folder"] = str(references[key][0])
    if key == "hidden":
        data["options"] = hidden_repair.DEFAULT_OPTIONS
    with module._LOCK:
        module._STATUS.clear()
        module._STATUS.update(running=False, total=len(files), done=len(files), folder=folder, current="")
        if key == "bl":
            module._STATUS.update(phase="scan", result=data)
            hits = [file for file in files if file.get("hits")]
            if hits:
                module._STATUS.update(phase="clear", applied={**data, "files": hits})
        elif key == "font":
            module._STATUS["result"] = data
        else:
            module._STATUS.update(phase="repair", applied=data)


def run(folder: str, chain: dict) -> None:
    """精简完成后逐份检查，同一个当前工作簿处理完所有已选操作才关闭。"""
    current = status()
    if not current.get("running") or current.get("folder") != str(folder):
        prepare(folder, chain)
    with _LOCK:
        _STATUS["stage"] = "chain"
        _save()

    reserved = False
    try:
        with wf.STATE_LOCK:
            if wf.STATE.get("busy"):
                raise ValueError("后台任务正在执行，请等它结束")
            job = wf.STATE.get("job") or {}
            if job.get("pending_adjustment") or job.get("deferred_adjustments"):
                raise ValueError("还有未完成的文件写入，请先继续或放弃上次写入")
            wf.STATE.update(busy=True, maintenance="clean_chain")
            reserved = True
        with debug_log.side_task("clean_chain", {"folder": folder, "old_folder": chain.get("old_folder", ""),
                                                **{key: _enabled(chain, key) for key in OPTIONAL_STEPS}}):
            paths_by_step, targets, references = _prepare_steps(folder, chain)
            results: dict[str, list[dict]] = {key: [] for key in paths_by_step}
            history = {key: operation_history.begin(folder, ["bl_scan" if key == "bl" else key])
                       for key in ("bl", "links", "hidden") if key in paths_by_step}
            try:
                if targets:
                    _process_files(folder, paths_by_step, targets, references, results)
            except Exception as error:
                debug_log.exception(error, "精简后自动检查")
                for files in results.values():
                    files.append({"name": "任务", "path": "", "error": str(error)})
            for key, files in results.items():
                if not files and _stopped():
                    _mark_stopped(key)
                else:
                    _publish_details(folder, key, files, references)
                    _finish_results(key, files)
                    if _stopped() and len(files) < len(paths_by_step[key]):
                        with _LOCK:
                            step = _step(key)
                            step["text"] += f"，已停止，剩余 {len(paths_by_step[key]) - len(files)} 份未检查"
                ids = history.get(key, [])
                if key == "bl" and any(file.get("hits") for file in files):
                    ids += operation_history.begin(folder, ["bl_clear"])
                operation_history.finish(ids, "interrupted" if _stopped() else
                                         "failed" if any(file.get("error") for file in files) else "completed")
    except Exception as error:
        debug_log.exception(error, "精简后自动检查")
        for key, _ in STEPS:
            if _step(key)["state"] in {"waiting", "running"}:
                _finish_step(key, 0, [_problem({}, None, str(error), error=True)], f"没能完成：{error}")
    finally:
        if reserved:
            with wf.STATE_LOCK:
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)
        with _LOCK:
            _STATUS.update(running=False, stage="done")
            _save()
