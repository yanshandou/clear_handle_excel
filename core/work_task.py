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
            task.pop("focus", None)
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


def upload_dest(task: dict[str, Any]) -> Path:
    """上传的终点：没另外填就是共享文件夹（用户 2026-11-07：几个账号分着做时，先传到团队 OneDrive 的中转文件夹）"""
    return Path(task.get("upload_to") or task["shared"])


def is_relay(task: dict[str, Any]) -> bool:
    return bool(task.get("upload_to")) and Path(task["upload_to"]) != Path(task["shared"])


def set_upload_to(task_id: str, path_text: Any) -> dict[str, Any]:
    """改任务的上传终点。空 = 回到直接传共享文件夹"""
    text = path_text.strip().strip('"').strip("'") if isinstance(path_text, str) else ""
    with _LOCK:
        data = _load()
        task = next((t for t in data["tasks"] if t.get("id") == task_id), None)
        if not task:
            raise ValueError("找不到这个任务")
        if text:
            dest = Path(text).expanduser()
            if not dest.is_dir():
                raise ValueError(f"终点文件夹不存在：{dest}（OneDrive 的话请确认已经同步到这台电脑）")
            dest = dest.resolve()
            for name, own in (("处理", task["work"]), ("参考", task["reference"])):
                mine = Path(own).resolve()
                if dest == mine or mine in dest.parents or dest in mine.parents:
                    raise ValueError(f"终点不能和本任务的{name}文件夹互相包含：{dest}")
            shared = Path(task["shared"]).resolve()
            if dest != shared and (shared in dest.parents or dest in shared.parents):
                raise ValueError("终点不能放在共享文件夹里面，也不能是它的上级文件夹")
            task["upload_to"] = "" if dest == shared else str(dest)
        else:
            task["upload_to"] = ""
        _save(data)
    debug_log.event("🚚 上传终点", 任务=task["name"], 终点=task["upload_to"] or task["shared"])
    return task


def set_single(on: Any) -> dict[str, Any]:
    """保留旧版的路径偏好开关：是否优先使用自选路径；两种偏好都允许修改路径。"""
    with _LOCK:
        data = _load()
        data["single"] = on is True
        _save(data)
    debug_log.event("📂 自选路径优先" if on is True else "📁 使用任务默认路径")
    return overview()


def require_folder(folder: Any, role: str = "work") -> str:
    """任务路径只作默认值，各功能可独立填写本地处理 / 参考文件夹。"""
    what = "原始备份文件夹" if role == "reference" else "Excel 文件夹路径"
    path = _clean(folder, what)
    if any(_is_onedrive(p) for p in path.parts):
        raise ValueError(f"这个文件夹在 OneDrive 里，就是共享文件夹，程序不能直接改：{path}"
                         "。请回主页点「🆕 新任务」，在本地的处理文件夹里改，改好再用「上传 Excel」传回去")
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


# ───────────────────────── 本轮只处理待重新处理文件（用户 2026-11-06 定） ─────────────────────────
# 上传前发现共享有更新，同步下来的文件标成「待重新处理」。上传页一键把这些文件记成本轮清单（task["focus"]），
# 1～8 步列文件时（workflow.list_workbooks）只列清单里的；清单里的文件上传成功就划掉，全划掉自动回到整个文件夹。

def _focus_key(rel: str) -> str:
    from core.common import strip_review_prefix
    parts = rel.replace("\\", "/").split("/")
    parts[-1] = strip_review_prefix(parts[-1])
    return "/".join(parts).casefold()


def focus_files(task: dict[str, Any] | None) -> list[str]:
    focus = (task or {}).get("focus") or {}
    files = focus.get("files") if isinstance(focus, dict) else None
    return [rel for rel in files or [] if isinstance(rel, str) and rel]


def set_focus(task_id: str, files: Any) -> dict[str, Any]:
    """本轮只处理这几个文件（处理文件夹下的相对路径）"""
    if not isinstance(files, list) or not files or not all(isinstance(rel, str) and rel for rel in files):
        raise ValueError("没有待重新处理的文件")
    with _LOCK:
        data = _load()
        task = next((t for t in data["tasks"] if t.get("id") == task_id), None)
        if not task:
            raise ValueError("找不到这个任务")
        work = Path(task["work"])
        keep: dict[str, str] = {}
        for rel in files:
            if "\\" in rel or ":" in rel or any(part in ("", ".", "..") for part in rel.split("/")):
                raise ValueError(f"文件路径无效：{rel}")
            if not (work / rel).is_file():
                raise ValueError(f"处理文件夹里找不到：{rel}，请刷新文件列表")
            keep.setdefault(_focus_key(rel), rel)
        task["focus"] = {"files": list(keep.values()), "since": datetime.now().strftime("%Y-%m-%d %H:%M")}
        _save(data)
    debug_log.event("🎯 本轮只处理待重新处理文件", 任务=task["name"], 文件=len(keep))
    return task


def clear_focus(task_id: str, files: list[str] | None = None, *, why: str = "退出") -> dict[str, Any] | None:
    """files=None：整个清单退出；否则只划掉这几个，划完了就自动退出"""
    with _LOCK:
        data = _load()
        task = next((t for t in data["tasks"] if t.get("id") == task_id), None)
        if not task or not focus_files(task):
            return task
        before = focus_files(task)
        drop = None if files is None else {_focus_key(rel) for rel in files}
        left = [] if drop is None else [rel for rel in before if _focus_key(rel) not in drop]
        if len(left) == len(before):
            return task
        if left:
            task["focus"] = {**task["focus"], "files": left}
        else:
            task.pop("focus", None)
        _save(data)
    debug_log.event("🎯 本轮清单" + ("划掉已上传" if left else "结束，回到整个文件夹"), 任务=task.get("name"),
                    原因=why, 还剩=len(left))
    return task


def _same_or_inside(folder: Path, root_text: str) -> str | None:
    """folder 是 root 本身或它的子文件夹 → 返回相对前缀（本身是 ""）"""
    try:
        root = Path(root_text).expanduser().resolve()
    except (OSError, TypeError, ValueError):
        return None
    a, b = str(folder).replace("\\", "/").casefold(), str(root).replace("\\", "/").casefold()
    if a == b:
        return ""
    if a.startswith(b.rstrip("/") + "/"):
        return a[len(b.rstrip("/")) + 1:]
    return None


def focus_filter(folder: Path) -> tuple[dict[str, Any], set[str], str] | None:
    """当前任务有本轮清单、且 folder 是它的处理 / 参考文件夹（或子文件夹）时，
    返回 (任务, 清单 key 集合, 子文件夹前缀)；否则 None（照常列整个文件夹）"""
    data = _load()
    task = next((t for t in data["tasks"] if t.get("id") == data["current"]), None)
    files = focus_files(task)
    if not task or not files:
        return None
    for field in ("work", "reference"):
        prefix = _same_or_inside(folder, task.get(field) or "")
        if prefix is not None:
            return task, {_focus_key(rel) for rel in files}, prefix
    return None


# ───────────────────────── 完成任务（主页 FINISH） ─────────────────────────

ISSUE_KEYS = ("待重新处理", "没上传", "别人改过", "共享里没有", "读不到")


def _check(task: dict[str, Any], progress=None) -> dict[str, Any]:
    """收尾检查本体：先列两边的文件、查共享里有没有和待重新处理（不读内容，很快），
    再逐个比对内容（要读文件，慢，progress(done, total, rel) 报进度）。只读，不改文件。"""
    from core import distribute
    work, shared = Path(task["work"]), Path(task["shared"])
    issues: dict[str, list[str]] = {key: [] for key in ISSUE_KEYS}
    files: dict[str, dict[str, str]] = {}
    report = progress or (lambda *a, **k: None)
    with debug_log.side_task("work_task_finish_check", {"shared": str(shared), "work": str(work)}):
        if not work.is_dir() or not shared.is_dir():
            raise ValueError("处理文件夹或共享文件夹不存在了，请确认路径")
        report(phase="list")
        baseline = distribute.load_baseline()
        mine = list(distribute.scan_folder(work)["excel"])
        shared_files = distribute.scan_folder(shared)["excel"]
        relay = upload_dest(task) if is_relay(task) else None
        if relay is not None and not relay.is_dir():
            raise ValueError(f"中转文件夹不存在了：{relay}")
        relay_files = {distribute._key(rel): rel for rel in distribute.scan_folder(relay)["excel"]} if relay else {}
        theirs = {distribute._key(rel): rel for rel in shared_files}
        report(phase="listed", work_files=len(mine), shared_files=len(shared_files))
        to_compare = []
        for rel in mine:
            other = theirs.get(distribute._key(rel))
            if not other and relay is None:
                issues["共享里没有"].append(rel)
                files[rel] = {"state": "共享里没有"}
                continue
            base = distribute._task_baseline(baseline.get(distribute._base_key(shared / (other or rel)), {}), task["id"])
            if base.get("review"):
                issues["待重新处理"].append(rel)
                files[rel] = {"state": "待重新处理"}
                continue
            if relay is not None:  # 中转：和传到中转的那份比（公司那份等汇总的人上传）
                other = relay_files.get(distribute._key(rel))
                if not other:
                    issues["没上传"].append(rel)
                    files[rel] = {"state": "没上传"}
                    continue
            to_compare.append((rel, other, base))
            files[rel] = {"state": "wait"}
        report(phase="compare", done=0, total=len(to_compare),
               early={key: list(issues[key]) for key in ("共享里没有", "待重新处理")})
        for index, (rel, other, base) in enumerate(to_compare):
            report(phase="compare", done=index, total=len(to_compare), current=rel)
            try:
                a, b = distribute._digest(work / rel), distribute._digest((relay or shared) / other)
            except OSError as err:
                note = f"{err.strerror or err}"
                issues["读不到"].append(f"{rel}：{note}")
                files[rel] = {"state": "读不到", "note": note}
                continue
            if a == b:
                files[rel] = {"state": "ok"}
                continue
            key = "别人改过" if relay is None and base and base.get("sha256") != b else "没上传"
            issues[key].append(rel)
            files[rel] = {"state": key}
        report(phase="compare", done=len(to_compare), total=len(to_compare))
        result = {"passed": not any(issues.values()), "issues": issues, "files": len(mine),
                  "relay": str(relay) if relay else "",
                  "shared_files": len(shared_files),
                  "file_states": [{"rel": rel, **files[rel]} for rel in mine]}
        debug_log.event("🏁 收尾检查", 文件=len(mine), **{k: len(v) for k, v in issues.items()})
    return result


def finish_check(task_id: str) -> dict[str, Any]:
    """收尾检查：处理文件夹里还有没有没传回共享的修改。只读，不改文件。
    - 没上传：处理里改过，和共享不一样，共享还是拿的时候的版本；
    - 别人改过：和共享不一样，共享也被别人改过了（要先拿最新版）；
    - 共享里没有：处理里有、共享里没有的 Excel（新文件还没传）。"""
    return _check(get(task_id))


def _mark_done(task_id: str) -> None:
    with _LOCK:
        data = _load()
        for task in data["tasks"]:
            if task.get("id") == task_id:
                task["done"] = True
                task["finished"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        _save(data)
    debug_log.event("🏁 任务完成", 任务=task_id)


# 收尾页 /finish：后台跑检查，页面每秒取一次进度（用户 2026-11-06 要：检查哪一项、进度多少、过了打勾）
_FINISH: dict[str, Any] = {"running": False}


def start_finish(task_id: str) -> None:
    task = get(task_id)
    with _LOCK:
        if _FINISH.get("running"):
            raise ValueError("收尾检查正在进行，请等它做完")
        _FINISH.clear()
        _FINISH.update(running=True, task=task, phase="start", done=0, total=0, current="",
                       started=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def progress(**info: Any) -> None:
        with _LOCK:
            _FINISH.update(info)

    def run() -> None:
        try:
            result = _check(task, progress)
            if result["passed"]:
                _mark_done(task["id"])
            with _LOCK:
                _FINISH.update(result=result, done_task=result["passed"])
        except Exception as err:  # 页面上显示原因
            debug_log.event("⚠️ 收尾检查没做完", 原因=str(err))
            with _LOCK:
                _FINISH.update(error=str(err))
        finally:
            with _LOCK:
                _FINISH.update(running=False, current="", finished=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    threading.Thread(target=run, name="work_task_finish", daemon=True).start()


def finish_status() -> dict[str, Any]:
    with _LOCK:
        return json.loads(json.dumps(_FINISH, ensure_ascii=False))


def finish(task_id: str) -> dict[str, Any]:
    """收尾检查通过才标成已完成；没通过返回检查结果，不标"""
    check = finish_check(task_id)
    if not check["passed"]:
        return {**check, "done": False}
    _mark_done(task_id)
    return {**check, "done": True}
