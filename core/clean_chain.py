"""精简 Sheet 之后自动接着做的例行检查：排查 BL 超链接 → 修复引用错误 →（可选）修复隐藏页。

用户 2026-11-07 定：主页只留「精简 Sheet」一站，原来第 4、5、6 步不用再单独点。
- 每一步都照用原来页面的后台（bl_links / link_repair / hidden_repair），规则一点不改，只是不再等人确认；
- 一步出错不影响后面几步，最后汇总成「改好了 / 改不了 / 出错」，改不了、出错的都列出来；
- 结果存进数据库，重开程序也能在精简 Sheet 页看到；原来的页面（主页「更多工具」）还能看每一步的详细结果。
引用错误和隐藏页要对照参考（原始备份）文件夹，没填就跳过这两步并写明原因。
"""
from __future__ import annotations

import copy
import threading
import time
from typing import Any, Callable

from core import debug_log, storage

SNAPSHOT = "clean_chain"
POLL_SECONDS = 0.4
BUSY_WAIT_SECONDS = 15.0  # 上一步刚结束，后台标志还没放开时最多等这么久
MAX_PROBLEMS = 300  # 每一步最多列这么多处「改不了 / 出错」，太多时只写总数

STEPS = (
    ("bl", "排查 BL 超链接"),
    ("links", "修复引用错误"),
    ("hidden", "修复隐藏页"),
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
                if step["state"] in {"waiting", "running"}:
                    step.update(state="failed", text="程序关掉了，这一步没做完")
        _STATUS.update(saved)


def _save() -> None:
    try:
        storage.save_snapshot(SNAPSHOT, copy.deepcopy(_STATUS))
    except Exception as err:
        debug_log.event("⚠️ 自动检查结果没能存进数据库", 原因=str(err))


def options_from(payload: Any) -> dict | None:
    """网页传来的 chain：{"hidden": bool, "old_folder": str}；没传就不自动检查（旧页面、单项调用）"""
    if payload is None:
        return None
    if not isinstance(payload, dict):
        raise ValueError("自动检查选项格式不正确")
    hidden = payload.get("hidden", False)
    old_folder = payload.get("old_folder", "")
    if type(hidden) is not bool or not isinstance(old_folder, str):
        raise ValueError("自动检查选项格式不正确")
    return {"hidden": hidden, "old_folder": old_folder.strip().strip('"').strip("'")}


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


def prepare(folder: str, chain: dict) -> None:
    """点了「确认执行」：精简还在跑，先把后面几步排上（页面马上能显示「等精简完成」）"""
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
        if not chain.get("hidden"):
            _step("hidden").update(state="off", text="这次没勾选")
        _save()


def drop(folder: str) -> None:
    """精简任务没能开始（已有任务在跑）：刚排上的自动检查撤掉"""
    with _LOCK:
        if _STATUS.get("folder") == str(folder) and _STATUS.get("stage") == "clean":
            _STATUS.clear()
            _save()


def _problem(file: dict, item: dict | None, reason: str, *, error: bool = False) -> dict:
    found = {"file": file.get("name", ""), "path": file.get("path", ""), "reason": reason, "error": error}
    if item:
        found.update(sheet=item.get("sheet", ""), cell=item.get("anchor") or item.get("addr", ""),
                     addr=item.get("addr", ""))
    return found


def _finish_step(key: str, fixed: int, problems: list[dict], text: str) -> None:
    errors = sum(1 for p in problems if p["error"])
    state = "failed" if errors else "warn" if problems else "done"
    _update(key, state=state, fixed=fixed, problems=problems[:MAX_PROBLEMS], problem_total=len(problems),
            error_total=errors, text=text)


def _launch(start: Callable[[], None], key: str) -> bool:
    """开始一步；上一步刚结束、后台标志还没放开时等一会儿再试"""
    deadline = time.monotonic() + BUSY_WAIT_SECONDS
    while True:
        try:
            start()
            return True
        except ValueError as err:
            if "后台任务正在执行" in str(err) and time.monotonic() < deadline:
                time.sleep(POLL_SECONDS)
                continue
            _finish_step(key, 0, [{"file": "", "path": "", "reason": str(err), "error": True}], f"没能开始：{err}")
            return False


def _wait(read_status: Callable[[], dict]) -> dict:
    time.sleep(POLL_SECONDS)
    while read_status().get("running"):
        time.sleep(POLL_SECONDS)
    return read_status()


def _file_errors(files: list[dict]) -> list[dict]:
    return [_problem(file, None, file["error"], error=True) for file in files if file.get("error")]


def _run_bl(folder: str) -> None:
    from core import bl_links
    _update("bl", state="running", text="正在扫描…")
    if not _launch(lambda: bl_links.start(folder), "bl"):
        return
    scanned = _wait(bl_links.status).get("result") or {"files": []}
    problems = _file_errors(scanned["files"])
    hits = sum(len(file.get("hits", [])) for file in scanned["files"] if not file.get("error"))
    if not hits:
        _finish_step("bl", 0, problems, "没有要清除的 BL 超链接")
        return
    _update("bl", text=f"找到 {hits} 处，正在清除…")
    if not _launch(bl_links.clear, "bl"):
        return
    applied = _wait(bl_links.status).get("applied") or {"files": []}
    fixed = sum(1 for file in applied["files"] for cell in file.get("cells", []) if cell.get("status") == "ok")
    problems += _file_errors(applied["files"])
    _finish_step("bl", fixed, problems, f"清除了 {fixed} 处 BL 超链接" if fixed else "没能清除 BL 超链接")


def _need_old_folder(key: str, old_folder: str) -> bool:
    if old_folder:
        return True
    _update(key, state="skipped", text="没填参考（原始备份）文件夹，跳过这一步")
    return False


def _run_links(folder: str, old_folder: str) -> None:
    from core import link_repair
    if not _need_old_folder("links", old_folder):
        return
    _update("links", state="running", text="正在扫描、对照参考、修复…")
    if not _launch(lambda: link_repair.repair(folder, old_folder), "links"):
        return
    applied = _wait(link_repair.status).get("applied") or {"files": []}
    fixed, problems = 0, _file_errors(applied["files"])
    for file in applied["files"]:
        if file.get("error"):
            continue
        for cell in file.get("cells", []):
            if cell.get("status") in {"ok", "resolved"}:
                fixed += 1
            else:
                problems.append(_problem(file, cell, cell.get("reason") or "未修复"))
    total = fixed + sum(1 for p in problems if not p["error"])
    _finish_step("links", fixed, problems,
                 f"引用错误 {total} 处：修好 {fixed} 处" + (f"，改不了 {total - fixed} 处" if total > fixed else "")
                 if total else "没有引用错误")


def _run_hidden(folder: str, old_folder: str) -> None:
    from core import hidden_repair
    if not _need_old_folder("hidden", old_folder):
        return
    _update("hidden", state="running", text="正在对照参考恢复…")
    if not _launch(lambda: hidden_repair.repair(folder, old_folder), "hidden"):
        return
    applied = _wait(hidden_repair.status).get("applied") or {"files": []}
    fixed, problems = 0, _file_errors(applied["files"])
    for file in applied["files"]:
        if file.get("error"):
            continue
        for cell in file.get("cells", []):
            if cell.get("status") == "ok":
                fixed += 1
            elif cell.get("action") == "skip" or cell.get("status"):
                problems.append(_problem(file, cell, cell.get("reason") or "未恢复"))
    _finish_step("hidden", fixed, problems,
                 f"恢复了 {fixed} 处隐藏表 / 列" if fixed else "没有要恢复的隐藏表 / 列")


def run(folder: str, chain: dict) -> None:
    """精简 Sheet 后台任务结束后，在同一个线程里接着做（workflow.run_pending_workflow 调用）"""
    current = status()
    if not current.get("running") or current.get("folder") != str(folder):
        prepare(folder, chain)  # 「继续上次任务」续跑精简时重新排上
    with _LOCK:
        _STATUS["stage"] = "chain"
        _save()
    old_folder = chain.get("old_folder", "")
    with debug_log.side_task("clean_chain", {"folder": folder, "old_folder": old_folder, "hidden": chain.get("hidden")}):
        for key, work in (("bl", lambda: _run_bl(folder)),
                          ("links", lambda: _run_links(folder, old_folder)),
                          ("hidden", lambda: _run_hidden(folder, old_folder) if chain.get("hidden") else None)):
            try:
                work()
            except Exception as err:  # 一步出错不影响后面几步
                debug_log.exception(err, f"精简后自动检查 {key}")
                _finish_step(key, 0, [{"file": "", "path": "", "reason": str(err), "error": True}], f"出错了：{err}")
    with _LOCK:
        _STATUS.update(running=False, stage="done")
        _save()

