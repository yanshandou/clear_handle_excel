"""新任务：只填共享文件夹路径，自动生成本地的「参考」「处理」文件夹（用户 2026-11-03 定）。

一个任务对应共享文件夹里的一个文件夹（连子文件夹一起）。例：

    共享  D:\\Users\\s.go\\OneDrive - 会社\\介舟ファミリー\\20基本設計\\10DB\\12.ビュー定義書
    参考  D:\\Users\\s.go\\Desktop\\handle-excel\\参考\\介舟ファミリー\\20基本設計\\10DB\\12.ビュー定義書
    处理  D:\\Users\\s.go\\Desktop\\handle-excel\\处理\\介舟ファミリー\\20基本設計\\10DB\\12.ビュー定義書

- 「OneDrive - 公司名」后面那一段原样接到工作根目录的「参考」「处理」后面，目录结构和共享一样；
  路径里没有 OneDrive 时只用共享文件夹自己的名字。
- 工作根目录默认是桌面下的 handle-excel（%USERPROFILE%\\Desktop\\handle-excel），可以改；
  不能放进 OneDrive（会被同步给别人），也不能放进共享文件夹里面。
- 参考 = 拿下来的原件，不改，修复颜色 / 引用 / 隐藏页的「原始备份」用它；处理 = 1～8 步改的地方。
- 拿文件、上传都用 core/distribute.py 的「拿最新版」「复制」，覆盖前检查别人改过没有。
  上传不另做备份（用户 2026-11-03 定）：只有共享那份和快照一样时才覆盖，被盖掉的内容就是参考里的原件。
任务列表存在 execution.sqlite3（快照 work_tasks），以前的任务都保留，可以切回去。
"""
from __future__ import annotations

import json
import os
import threading
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from core import debug_log, storage

SNAPSHOT = "work_tasks"
APP_DIR = "handle-excel"
REFERENCE, WORK, FETCH_BACKUP = "参考", "处理", "拿最新版前的备份"
_LOCK = threading.RLock()


def _is_onedrive(part: str) -> bool:
    return part.casefold().startswith("onedrive")


def _clean(value: Any, what: str) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"请填写{what}")
    return Path(value.strip().strip('"').strip("'")).expanduser()


def relative_part(shared: Path) -> list[str]:
    """共享路径里「OneDrive - 公司名」后面的那一段；没有 OneDrive 就只用最后一级文件夹名"""
    parts = list(shared.parts)
    idx = max((i for i, p in enumerate(parts) if _is_onedrive(p)), default=-1)
    rel = parts[idx + 1:] if idx >= 0 else []
    return rel or [shared.name or "共享"]


def default_work_root(shared: Path | None = None) -> Path:
    """默认放在桌面的 handle-excel 下（用户 2026-11-03 定：D:\\Users\\<用户名>\\Desktop\\handle-excel，用户名按当前用户）"""
    return Path.home() / "Desktop" / APP_DIR


def derive(shared_text: Any, work_root_text: Any = "", *, must_exist: bool = True) -> dict[str, Any]:
    """按共享路径算出参考 / 处理 / 备份文件夹，不建任何文件夹"""
    shared = _clean(shared_text, "共享文件夹路径")
    if must_exist and not shared.is_dir():
        raise ValueError(f"共享文件夹不存在：{shared}（OneDrive 的话请确认已经同步到这台电脑）")
    if must_exist:
        shared = shared.resolve()
    root = _clean(work_root_text, "本地工作文件夹") if isinstance(work_root_text, str) and work_root_text.strip() \
        else default_work_root(shared)
    if any(_is_onedrive(p) for p in root.parts):
        raise ValueError(f"本地工作文件夹不能放在 OneDrive 里面（会被同步给别人）：{root}")
    try:
        root_resolved = root.resolve()
    except OSError:
        root_resolved = root
    if root_resolved == shared or shared in root_resolved.parents or root_resolved in shared.parents:
        raise ValueError("本地工作文件夹不能和共享文件夹互相包含，请换一个地方")
    rel = relative_part(shared)
    return {
        "shared": str(shared),
        "work_root": str(root),
        "rel": "/".join(rel),
        "name": rel[-1],
        "reference": str(root.joinpath(REFERENCE, *rel)),
        "work": str(root.joinpath(WORK, *rel)),
        "fetch_backup": str(root.joinpath(FETCH_BACKUP, *rel)),
    }


def _load() -> dict[str, Any]:
    try:
        data = storage.load_snapshot(SNAPSHOT) or {}
    except Exception as err:
        debug_log.event("⚠️ 没能读出任务列表", 原因=str(err))
        data = {}
    tasks = data.get("tasks") if isinstance(data.get("tasks"), list) else []
    return {"tasks": tasks, "current": data.get("current") or "", "single": data.get("single") is True}


def _save(data: dict[str, Any]) -> None:
    storage.save_snapshot(SNAPSHOT, data)


def overview() -> dict[str, Any]:
    data = _load()
    last_root = data["tasks"][0]["work_root"] if data["tasks"] else ""
    return {**data, "last_work_root": last_root}


def get(task_id: str) -> dict[str, Any]:
    for task in _load()["tasks"]:
        if task.get("id") == task_id:
            return task
    raise ValueError("找不到这个任务，可能已经被删除，请重新开一个新任务")


def create(shared_text: Any, work_root_text: Any = "") -> dict[str, Any]:
    """开新任务：建好参考、处理文件夹，设成当前任务。同一个共享文件夹 + 工作根目录的任务已经有了就沿用"""
    paths = derive(shared_text, work_root_text)
    for key in ("reference", "work"):
        Path(paths[key]).mkdir(parents=True, exist_ok=True)
    from core import workflow as wf
    with wf.STATE_LOCK, _LOCK:
        data = _load()
        same = next((t for t in data["tasks"] if t.get("shared") == paths["shared"]
                     and t.get("work_root") == paths["work_root"]), None)
        if same:
            data["tasks"].remove(same)
            task = {**same, **paths, "done": False}  # 同一个文件夹再开一次 = 重新开始做
        else:
            task = {"id": uuid.uuid4().hex[:12], **paths, "created": datetime.now().strftime("%Y-%m-%d %H:%M")}
        task["used"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        data["tasks"].insert(0, task)
        data["current"] = task["id"]
        data["single"] = False  # 开任务 = 回到任务模式
        wf.activate_folder(task["work"], commit=lambda: _save(data))
    debug_log.event("🆕 新任务" if not same else "🔁 沿用以前的任务", 共享=paths["shared"], 参考=paths["reference"],
                    处理=paths["work"])
    return task


def switch(task_id: str) -> dict[str, Any]:
    from core import workflow as wf
    with wf.STATE_LOCK, _LOCK:
        data = _load()
        task = next((t for t in data["tasks"] if t.get("id") == task_id), None)
        if not task:
            raise ValueError("找不到这个任务")
        data["tasks"].remove(task)
        task["used"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        data["tasks"].insert(0, task)
        data["current"] = task_id
        data["single"] = False
        wf.activate_folder(task["work"], commit=lambda: _save(data))
    debug_log.event("🔀 切换任务", 共享=task["shared"], 处理=task["work"])
    return task


def _norm(path: Path) -> str:
    try:
        path = path.resolve()
    except OSError:
        pass
    return os.path.normcase(os.path.normpath(str(path))).casefold()


def _inside(path: Path, base: Path) -> bool:
    a, b = _norm(path), _norm(base)
    return a == b or a.startswith(b.rstrip("\\/") + os.sep)


def set_single(on: Any) -> dict[str, Any]:
    """保留旧版的路径偏好开关：是否优先使用自选路径；两种偏好都允许修改路径。"""
    with _LOCK:
        data = _load()
        data["single"] = on is True
        _save(data)
    debug_log.event("🔧 单项模式" if on is True else "📁 回到任务模式")
    return overview()


def require_folder(folder: Any, role: str = "work") -> str:
    """任务路径只作默认值，各功能可独立填写本地处理 / 参考文件夹。"""
    what = "原始备份文件夹" if role == "reference" else "Excel 文件夹路径"
    path = _clean(folder, what)
    if any(_is_onedrive(p) for p in path.parts):
        raise ValueError(f"这个文件夹在 OneDrive 里，就是共享文件夹，程序不能直接改：{path}"
                         "。请回主页点「🆕 新任务」，在本地的处理文件夹里改，改好再用第 9 步传回去")
    return str(folder)


def plan_for(task: dict[str, Any]) -> dict[str, Any]:
    """把任务换成 distribute 用的清单：来源 = 处理文件夹，目的地 = 共享文件夹"""
    work, shared = Path(task["work"]), Path(task["shared"])
    work.mkdir(parents=True, exist_ok=True)
    plan = {"label": f"任务：{task['name']}", "people": [{"name": task["name"], "folders": [
        {"source": [work.name], "destination": [shared.name]}]}]}
    return {"plan": json.dumps(plan, ensure_ascii=False), "source_root": str(work.parent),
            "dest_root": str(shared.parent), "reference": task["reference"],
            "fetch_backup": task["fetch_backup"]}


# ───────────────────────── 完成任务（主页 FINISH） ─────────────────────────

def finish_check(task_id: str) -> dict[str, Any]:
    """收尾检查：处理文件夹里还有没有没传回共享的修改。只读，不改文件。
    - 没上传：处理里改过，和共享不一样，共享还是拿的时候的版本；
    - 别人改过：和共享不一样，共享也被别人改过了（要先拿最新版）；
    - 共享里没有：处理里有、共享里没有的 Excel（新文件还没传）。"""
    from core import distribute
    task = get(task_id)
    work, shared = Path(task["work"]), Path(task["shared"])
    issues: dict[str, list[str]] = {"待重新处理": [], "没上传": [], "别人改过": [], "共享里没有": [], "读不到": []}
    with debug_log.side_task("work_task_finish_check", {"shared": str(shared), "work": str(work)}):
        if not work.is_dir() or not shared.is_dir():
            raise ValueError("处理文件夹或共享文件夹不存在了，请确认路径")
        baseline = distribute.load_baseline()
        mine = distribute.scan_folder(work)["excel"]
        theirs = {distribute._key(rel): rel for rel in distribute.scan_folder(shared)["excel"]}
        for rel in mine:
            other = theirs.get(distribute._key(rel))
            if not other:
                issues["共享里没有"].append(rel)
                continue
            base = distribute._task_baseline(baseline.get(distribute._base_key(shared / other), {}), task["id"])
            if base.get("review"):
                issues["待重新处理"].append(rel)
                continue
            try:
                a, b = distribute._digest(work / rel), distribute._digest(shared / other)
            except OSError as err:
                issues["读不到"].append(f"{rel}：{err.strerror or err}")
                continue
            if a == b:
                continue
            issues["别人改过" if base and base.get("sha256") != b else "没上传"].append(rel)
        result = {"passed": not any(issues.values()), "issues": issues, "files": len(mine)}
        debug_log.event("🏁 收尾检查", 文件=len(mine), **{k: len(v) for k, v in issues.items()})
    return result


def finish(task_id: str) -> dict[str, Any]:
    """收尾检查通过才标成已完成；没通过返回检查结果，不标"""
    check = finish_check(task_id)
    if not check["passed"]:
        return {**check, "done": False}
    with _LOCK:
        data = _load()
        for task in data["tasks"]:
            if task.get("id") == task_id:
                task["done"] = True
                task["finished"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        _save(data)
    debug_log.event("🏁 任务完成", 任务=task_id)
    return {**check, "done": True}
