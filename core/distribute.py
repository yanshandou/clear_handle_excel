"""按上传记录 JSON 把 Excel 文件从来源文件夹复制到目的地文件夹（/distribute 页面）。

JSON 格式（用户给的上传记录）::

    {"label": "...", "totalFiles": 153, "folderCount": 8,
     "people": [{"name": "呉", "totalFiles": 77,
                 "folders": [{"source": ["10要件定義", "10画面", "07_SC…"],
                              "destination": ["呉", "10要件定義", "10画面", "07_SC…"],
                              "files": 41, "owner": "呉", "startDate": "2026-10-23"}]}]}

规则（用户 2026-11-01 定）：
- 来源根目录、目的地根目录在页面上选，source / destination 是接在它们后面的相对路径；
- 只复制不移动，整个文件夹连子文件夹一起复制，只复制 Excel（~$ 临时文件不算），其他文件不管；
- 目的地已有同名文件就覆盖，文件名不改，覆盖了哪些在复制记录里写清楚；
- 每个文件夹核对三个数：JSON 写的 files、来源实际找到的 Excel 数、复制后目的地的 Excel 数，
  对不上的写清楚是哪两个数、哪些文件、哪边多了；
- owner 是作者、startDate 只是记录，都不影响复制。

分两步：先 preview() 列出清单给用户确认，再 start() 复制。复制前再查一遍来源，和清单不一样就不复制。
复制记录存进 execution.sqlite3（放在程序文件夹外，更新程序记录不丢），网页上能看、能下载 Markdown。
记录文档里不写日期和时间（要交出去的东西都这样），网页上可以显示。
"""
from __future__ import annotations

import copy
import hashlib
import json
import os
import shutil
import threading
import uuid
from datetime import datetime
from pathlib import Path, PurePath
from typing import Any, Callable

from core import content_check, debug_log, file_problem, pdf_cache, storage, workflow as wf
from core.common import _held_open, workbook_locked
from core.upload_review import Candidate, UploadReview

# Microsoft Excel 的表格文件（加载项 .xla/.xlam、CSV 不算）
EXCEL_SUFFIXES = (".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm")
HISTORY_NAME = "distribute_history"
BASELINE_NAME = "distribute_baseline"
MAX_HISTORY = 200
TEMP_SUFFIX = ".复制中~"

_LOCK = threading.RLock()
_STATUS: dict[str, Any] = {"running": False}
_LIVE: dict[str, Any] = {}  # 正在跑的那批清单（线程里就地改每个文件的 status），只给进度框读


# ───────────────────────── JSON ─────────────────────────

def _segments(value: Any, what: str, where: str) -> list[str]:
    """source / destination：字符串数组；也接受一整个用 / 或 \\ 分隔的字符串"""
    if isinstance(value, str):
        value = value.replace("\\", "/").split("/")
    if not isinstance(value, list):
        raise ValueError(f"{where} 的 {what} 要写成文件夹名数组，例如 [\"10要件定義\", \"10画面\"]")
    parts = []
    for item in value:
        if not isinstance(item, str):
            raise ValueError(f"{where} 的 {what} 里有不是文字的项：{item!r}")
        for piece in item.replace("\\", "/").split("/"):
            piece = piece.strip()
            if not piece or piece == ".":
                continue
            if piece == ".." or ":" in piece:
                raise ValueError(f"{where} 的 {what} 不能有「{piece}」，只能写根目录下面的相对路径")
            parts.append(piece)
    if not parts:
        raise ValueError(f"{where} 的 {what} 是空的")
    return parts


def _count(value: Any) -> int | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, int) and value >= 0:
        return value
    if isinstance(value, str) and value.strip().isdigit():
        return int(value.strip())
    return None


def parse_plan(text: str) -> dict[str, Any]:
    """读上传记录 JSON，拆成一个个文件夹；JSON 里自己的合计数对不上也记下来"""
    if not isinstance(text, str) or not text.strip():
        raise ValueError("请粘贴上传记录 JSON，或选择 .json 文件")
    try:
        data = json.loads(text.lstrip("\ufeff"))
    except json.JSONDecodeError as err:
        raise ValueError(f"JSON 格式不对：第 {err.lineno} 行第 {err.colno} 列，{err.msg}") from None
    if not isinstance(data, dict) or not isinstance(data.get("people"), list) or not data["people"]:
        raise ValueError("JSON 里没有 people（每个人和他的文件夹列表）")
    folders: list[dict[str, Any]] = []
    notes: list[str] = []
    for p_index, person in enumerate(data["people"], 1):
        if not isinstance(person, dict):
            raise ValueError(f"people 第 {p_index} 项格式不对")
        name = str(person.get("name") or f"第 {p_index} 个人").strip()
        items = person.get("folders")
        if not isinstance(items, list):
            raise ValueError(f"{name} 没有 folders 列表")
        person_sum = 0
        for f_index, item in enumerate(items, 1):
            where = f"{name} 的第 {f_index} 个文件夹"
            if not isinstance(item, dict):
                raise ValueError(f"{where} 格式不对")
            files = _count(item.get("files"))
            if item.get("files") is not None and files is None:
                raise ValueError(f"{where} 的 files 要写成整数，现在是 {item.get('files')!r}")
            person_sum += files or 0
            folders.append({
                "id": f"f{len(folders) + 1}",
                "person": name,
                "owner": str(item.get("owner") or "").strip(),
                "start_date": str(item.get("startDate") or "").strip(),
                "source": _segments(item.get("source"), "source", where),
                "destination": _segments(item.get("destination"), "destination", where),
                "json_files": files,
            })
        total = _count(person.get("totalFiles"))
        if total is not None and total != person_sum:
            notes.append(f"JSON 里 {name} 的 totalFiles 写的是 {total}，各文件夹 files 加起来是 {person_sum}")
    total = _count(data.get("totalFiles"))
    folder_sum = sum(f["json_files"] or 0 for f in folders)
    if total is not None and total != folder_sum:
        notes.append(f"JSON 里总的 totalFiles 写的是 {total}，各文件夹 files 加起来是 {folder_sum}")
    count = _count(data.get("folderCount"))
    if count is not None and count != len(folders):
        notes.append(f"JSON 里 folderCount 写的是 {count}，实际列了 {len(folders)} 个文件夹")
    if not folders:
        raise ValueError("JSON 里一个文件夹也没有")
    return {"label": str(data.get("label") or "").strip(), "folders": folders, "notes": notes}


# ───────────────────────── 列文件 ─────────────────────────

def _clean_root(value: Any, what: str) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"请填写{what}")
    root = Path(value.strip().strip('"').strip("'")).expanduser()
    if not root.is_dir():
        raise ValueError(f"{what}不存在：{root}")
    return root.resolve()


def source_open(path: Path) -> bool:
    """来源文件正在 Excel 中打开（旁边有 ~$ 锁文件、文件又被占用）。只看不删：来源文件夹一点都不动"""
    try:
        return path.with_name("~$" + path.name).exists() and _held_open(path)
    except OSError:
        return False


def is_excel(name: str) -> bool:
    return name.lower().endswith(EXCEL_SUFFIXES) and not name.startswith("~$")


def _key(rel: str) -> str:
    """Windows 文件名不分大小写：比较时用这个"""
    return os.path.normcase(rel).replace("\\", "/").casefold()


def scan_folder(folder: Path) -> dict[str, Any]:
    """递归列出文件夹里的 Excel（相对路径、大小、修改时间），另外数一下 ~$ 临时文件和其他文件"""
    excel: dict[str, dict[str, Any]] = {}
    dirs: list[str] = []
    temp: list[str] = []
    other: dict[str, int] = {}
    errors: list[str] = []

    def on_error(err: OSError) -> None:
        errors.append(f"{err.filename}：{err.strerror or err}")

    for dirpath, dirnames, filenames in os.walk(folder, onerror=on_error):
        dirnames.sort()
        base = Path(dirpath)
        dirs.extend((base / name).relative_to(folder).as_posix() for name in dirnames)
        for name in sorted(filenames):
            path = base / name
            rel = path.relative_to(folder).as_posix()
            if name.endswith(TEMP_SUFFIX):
                continue  # 上次复制被打断留下的临时文件
            if name.startswith("~$"):
                temp.append(rel)
                continue
            if not is_excel(name):
                ext = Path(name).suffix.lower() or "（没有扩展名）"
                other[ext] = other.get(ext, 0) + 1
                continue
            try:
                info = path.stat()
            except OSError as err:
                errors.append(f"{rel}：{err.strerror or err}")
                continue
            excel[rel] = {"rel": rel, "size": info.st_size, "mtime": info.st_mtime}
    return {"excel": excel, "dirs": sorted(dirs), "temp": temp, "other": other, "errors": errors}


def _inside(child: Path, parent: Path) -> bool:
    return child == parent or parent in child.parents


def _json_vs_source_hint(json_files: int | None, src: dict[str, Any]) -> str:
    """JSON 只写了数量、没写文件名，对不上时猜一下原因"""
    if json_files is None:
        return ""
    found = len(src["excel"])
    others = sum(src["other"].values())
    temps = len(src["temp"])
    if others and not temps and json_files == found + others:
        return "JSON 的数正好等于来源里 Excel 加上非 Excel 文件的个数，可能是把非 Excel 文件也算进去了"
    if temps and not others and json_files == found + temps:
        return "JSON 的数正好等于来源里 Excel 加上 ~$ 临时文件的个数，可能是把 Excel 打开时留下的临时文件也算进去了"
    if others and temps and json_files == found + others + temps:
        return "JSON 的数正好等于来源里所有文件（包括非 Excel 文件和 ~$ 临时文件）的个数，可能是把它们也算进去了"
    return "JSON 里只有数量没有文件名，没法指出是哪个文件；请对照下面来源文件夹的文件清单"


def _task_args(task_id: Any, plan_text: str, source_root: str, dest_root: str) -> tuple[str, str, str, dict | None]:
    """当前任务（core/work_task）：清单和两个根目录都按任务的处理文件夹、共享文件夹生成"""
    if not task_id:
        return plan_text, source_root, dest_root, None
    from core import work_task
    task = work_task.get(str(task_id))
    args = work_task.plan_for(task)
    info = {"id": task["id"], "name": task["name"], "shared": task["shared"], "work": task["work"],
            "reference": args["reference"], "fetch_backup": args["fetch_backup"], "picks": list(task.get("picks") or [])}
    return args["plan"], args["source_root"], args["dest_root"], info


def preview(plan_text: str, source_root: str, dest_root: str, task_id: Any = None) -> dict[str, Any]:
    """生成复制清单：每个文件夹要复制哪些、哪些会覆盖、三个数对不对得上。不改任何文件"""
    plan_text, source_root, dest_root, task = _task_args(task_id, plan_text, source_root, dest_root)
    plan = parse_plan(plan_text)
    src_root = _clean_root(source_root, "来源根目录")
    dst_root = _clean_root(dest_root, "目的地根目录")
    if src_root == dst_root:
        raise ValueError("来源根目录和目的地根目录是同一个文件夹，请重新选择")
    with debug_log.side_task("distribute_preview", {"source": str(src_root), "dest": str(dst_root),
                                                    "folders": len(plan["folders"]), "label": plan["label"]}):
        folders = []
        baseline = load_baseline()
        for item in plan["folders"]:
            folders.append(_preview_folder(item, src_root, dst_root, baseline))
        _cross_checks(folders)
        result = {
            "id": uuid.uuid4().hex[:12],
            "kind": "copy",
            "task": task,
            "label": plan["label"],
            "source_root": str(src_root),
            "dest_root": str(dst_root),
            "notes": plan["notes"],
            "folders": folders,
            "plan_text": plan_text,
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        result["summary"] = _summary(folders)
        debug_log.event("📋 复制清单", **result["summary"])
    with _LOCK:
        if not _STATUS.get("running"):
            _STATUS["preview"] = result
            _STATUS.pop("applied", None)
    return _public(result)


def _guard(src_file: Path, dest_file: Path, exists: bool, base: dict[str, Any] | None) -> dict[str, Any]:
    """覆盖前检查：目的地这个文件从我们拿走（或上次放进去）之后，有没有被别人改过。
    guard：same 没人动过 / unknown 没有记录、没法检查 / changed 别人改过 / gone 记录里有、现在目的地没有了 /
    identical 别人改过但和我们要放的内容一样 / none 目的地没有、也没有记录（新文件）"""
    if not exists:
        if not base:
            return {"guard": "none", "conflict": ""}
        return {"guard": "gone", "conflict": "拿的时候目的地有这个文件，现在没有了（可能被别人删了或改了名），没有放进去"}
    try:
        now = _digest(dest_file)
    except OSError as err:
        if not base:
            return {"guard": "unknown", "conflict": ""}
        return {"guard": "changed", "conflict": f"读不到目的地现在的文件，没法确认别人改过没有：{_explain(err)}"}
    if base and now == base.get("sha256"):
        return {"guard": "same", "conflict": ""}
    try:
        if _digest(src_file) == now:
            return {"guard": "identical", "conflict": ""}
    except OSError:
        pass
    # 用户 2026-11-11：没有记录、或记录太旧（拿很久以前的版本修的）时，按全部文字 + 打印区域格子数核对
    check = content_check.compare(src_file, dest_file, "来源", "目的地")
    if check["same"]:
        return {"guard": "content_same", "dest_sha": now, "conflict": "",
                "content_check": ("没有拿取记录" if not base else "目的地和拿的时候不一样") + "，按文字和打印区域核对一致，照常覆盖"
                + content_check.skipped_note(check)}
    if check["checked"]:
        return {"guard": "content_diff", "dest_sha": now, "content_diffs": check["diffs"],
                "conflict": ("没有拿取记录，" if not base else "目的地和你拿的时候不一样，")
                + f"按内容核对也不一样（{content_check.describe(check)}），没有覆盖。"
                "请先「⬇ 拿最新版」，在新版上改好再复制"}
    if not base:  # 这种格式按内容核对不了：照旧覆盖但标出
        return {"guard": "unknown", "conflict": ""}
    return {"guard": "changed", "dest_sha": now,
            "conflict": "别人改过了：目的地现在的文件和你拿的时候不一样，没有覆盖。请先「⬇ 拿最新版」，在新版上改好再复制"}


def _preview_folder(item: dict[str, Any], src_root: Path, dst_root: Path,
                    baseline: dict[str, dict[str, Any]] | None = None) -> dict[str, Any]:
    src = src_root.joinpath(*item["source"])
    dst = dst_root.joinpath(*item["destination"])
    entry: dict[str, Any] = {
        **item,
        "source_rel": "/".join(item["source"]),
        "dest_rel": "/".join(item["destination"]),
        "source_path": str(src),
        "dest_path": str(dst),
        "error": "",
        "files": [],
        "dest_extra": [],
        "temp": [],
        "other": {},
        "scan_errors": [],
        "source_count": None,
        "dest_count": None,
        "checks": [],
        "dirs_to_create": [],
    }
    if not src.is_dir():
        entry["error"] = "找不到来源文件夹" if not src.exists() else "来源路径不是文件夹"
        # 目录结构照 JSON 建：来源找不到也先把目的地的空文件夹建好，记录里写明是空的
        entry["dirs_to_create"] = [] if dst.is_dir() else ["."]
        debug_log.event("⚠️ 复制清单：" + entry["error"], 文件夹=str(src))
        entry["checks"] = _checks(entry, None, None)
        return entry
    if _inside(dst.resolve(), src.resolve()) or _inside(src.resolve(), dst.resolve()):
        entry["error"] = "来源文件夹和目的地文件夹互相包含，不能复制"
        debug_log.event("⚠️ 复制清单：" + entry["error"], 来源=str(src), 目的地=str(dst))
        return entry
    found = scan_folder(src)
    have = scan_folder(dst) if dst.is_dir() else {"excel": {}, "errors": []}
    have_keys = {_key(rel): info for rel, info in have["excel"].items()}
    source_keys = {_key(rel) for rel in found["excel"]}
    for rel, info in found["excel"].items():
        old = have_keys.get(_key(rel))
        locked = source_open(src / rel)
        dest_rel = old["rel"] if old else rel
        guard = _guard(src / rel, dst / dest_rel, bool(old), (baseline or {}).get(_base_key(dst / dest_rel)))
        entry["files"].append({
            **info,
            "action": "overwrite" if old else "new",
            "old_size": old["size"] if old else None,
            "old_mtime": old["mtime"] if old else None,
            "dest_rel": dest_rel,
            "source_open": locked,
            **guard,
        })
        if guard["conflict"]:
            debug_log.event("🛑 别人改过，不覆盖", 文件=f"{entry['dest_rel']}/{dest_rel}", 检查=guard["guard"],
                            原因=guard["conflict"][:60])
    entry["dest_extra"] = sorted(info["rel"] for key, info in have_keys.items() if key not in source_keys)
    # 目录结构和来源一样：没有 Excel 的子文件夹（空的、只有 PDF 的）也建出来
    entry["dirs_to_create"] = ([] if dst.is_dir() else ["."]) + [
        rel for rel in found["dirs"] if not (dst / rel).is_dir()]
    entry["temp"] = found["temp"]
    entry["other"] = found["other"]
    entry["scan_errors"] = found["errors"] + have.get("errors", [])
    entry["source_count"] = len(found["excel"])
    entry["dest_before"] = len(have["excel"])
    entry["dest_expected"] = len(found["excel"]) + len(entry["dest_extra"])
    entry["hint"] = _json_vs_source_hint(item["json_files"], found)
    entry["checks"] = _checks(entry, entry["source_count"], None)
    debug_log.event("📋 复制清单文件夹", 来源=entry["source_rel"], 目的地=entry["dest_rel"], JSON=item["json_files"],
                    来源Excel=entry["source_count"], 覆盖=sum(f["action"] == "overwrite" for f in entry["files"]),
                    别人改过=sum(bool(f["conflict"]) for f in entry["files"]),
                    没有版本记录=sum(f["guard"] == "unknown" for f in entry["files"]),
                    新建=sum(f["action"] == "new" for f in entry["files"]), 目的地多出=len(entry["dest_extra"]),
                    读不到=len(entry["scan_errors"]))
    return entry


def _cross_checks(folders: list[dict[str, Any]]) -> None:
    """两个文件夹复制到同一个目的地、或来源互相包含：文件会被复制两次 / 数量重复计算"""
    for i, a in enumerate(folders):
        for b in folders[i + 1:]:
            pairs = (("dest_path", "目的地"), ("source_path", "来源"))
            for field, what in pairs:
                pa, pb = Path(a[field]), Path(b[field])
                if pa == pb:
                    text = f"和「{b['source_rel']}」的{what}文件夹是同一个"
                elif _inside(pa, pb) or _inside(pb, pa):
                    text = f"和「{b['source_rel']}」的{what}文件夹互相包含"
                else:
                    continue
                a.setdefault("overlaps", []).append(text)
                b.setdefault("overlaps", []).append(text.replace(b["source_rel"], a["source_rel"]))


def _checks(entry: dict[str, Any], source: int | None, dest: int | None) -> list[dict[str, Any]]:
    """三个数两两比较：JSON ↔ 来源、来源 ↔ 目的地、JSON ↔ 目的地。没有的数（没写 / 还没复制）不比"""
    numbers = {"JSON": entry.get("json_files"), "来源": source, "目的地": dest}
    out = []
    for a, b in (("JSON", "来源"), ("来源", "目的地"), ("JSON", "目的地")):
        x, y = numbers[a], numbers[b]
        if x is None or y is None:
            continue
        check = {"pair": f"{a} ↔ {b}", "a": a, "b": b, "x": x, "y": y, "ok": x == y}
        if x != y:
            more, diff = (a, x - y) if x > y else (b, y - x)
            check["text"] = f"{a} {x} 个，{b} {y} 个，{more}{' ' if more.isascii() else ''}多 {diff} 个"
        out.append(check)
    return out


def _summary(folders: list[dict[str, Any]]) -> dict[str, int]:
    files = [f for folder in folders for f in folder["files"]]
    ok = [f for f in files if not f.get("conflict")]
    return {
        "文件夹": len(folders),
        "要复制": len(ok),
        "覆盖": sum(f["action"] == "overwrite" for f in ok),
        "新建": sum(f["action"] == "new" for f in ok),
        "别人改过": len(files) - len(ok),
        "没有版本记录": sum(f.get("guard") == "unknown" for f in files),
        "新建文件夹": sum(len(folder["dirs_to_create"]) for folder in folders),
        "数量不一致": sum(any(not c["ok"] for c in folder["checks"]) for folder in folders),
        "目的地多出": sum(bool(folder["dest_extra"]) for folder in folders),
        "有问题": sum(bool(folder["error"] or folder.get("overlaps") or folder["scan_errors"]) for folder in folders),
    }


def _public(result: dict[str, Any]) -> dict[str, Any]:
    return {k: v for k, v in result.items() if k != "plan_text"}


# ───────────────────────── 复制 ─────────────────────────

def _digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for block in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


# ───────────────────────── 版本记录（覆盖前检查别人改过没有） ─────────────────────────
# 目的地（共享文件夹）里每个文件「我们最后知道的版本」：从目的地拿文件时记下拿到的内容的哈希，
# 复制（覆盖）成功后换成我们放进去的内容的哈希。下次覆盖前先比一下目的地现在的哈希：
# 一样 = 没人动过，可以覆盖；不一样 = 别人改过，不覆盖，请先拿最新版。
# 按目的地文件的完整路径记（不分大小写）。同一个共享文件夹用盘符和 \\server 两种写法会被当成两个地方。

_BASE_LOCK = threading.RLock()


def _base_key(path: Path | str) -> str:
    return _key(str(path))


def load_baseline(*, strict: bool = False) -> dict[str, dict[str, Any]]:
    try:
        data = storage.load_snapshot(BASELINE_NAME) or {}
    except Exception as err:
        if strict:
            raise
        debug_log.event("⚠️ 没能读出版本记录", 原因=str(err))
        return {}
    files = data.get("files")
    return dict(files) if isinstance(files, dict) else {}


def _save_baseline(changes: dict[str, dict[str, Any]]) -> None:
    """把这次记下的版本并进记录里（先读再写，别的地方同时记的不丢）"""
    if not changes:
        return
    with _BASE_LOCK:
        files = load_baseline(strict=True)
        for key, change in changes.items():
            old = files.get(key, {})
            if old.get("how") == "复制" and "upload" not in old:
                old["upload"] = {"sha256": old["sha256"], "at": ""}
            tasks = dict(old.get("tasks", {}))
            for task_id, version in change.get("tasks", {}).items():
                tasks[task_id] = {**tasks.get(task_id, {}), **version}
            files[key] = {**old, **change, "tasks": tasks}
        storage.save_snapshot(BASELINE_NAME, {"files": files})
    debug_log.event("📌 记下目的地文件版本", 个数=len(changes))


def _base_entry(path: Path, digest: str, how: str, task_id: str = "") -> dict[str, Any]:
    try:
        size = path.stat().st_size
    except OSError:
        size = None
    entry = {"sha256": digest, "size": size, "path": str(path), "how": how}
    if how == "复制":
        entry["upload"] = {"sha256": digest, "at": datetime.now().isoformat(timespec="seconds")}
    if task_id:
        entry["tasks"] = {task_id: {"sha256": digest}}
    return entry


def _explain(err: BaseException) -> str:
    if isinstance(err, PermissionError) and "Excel" in str(err.strerror or ""):
        return f"{err.strerror}：{err.filename or ''}"
    if isinstance(err, PermissionError):
        return f"文件被占用或没有权限（可能在 Excel 中打开，请保存并关闭后重试）：{err.filename or ''}"
    if isinstance(err, FileNotFoundError):
        return f"找不到文件（可能被移走或网络盘断开）：{err.filename or ''}"
    if isinstance(err, OSError):
        return f"读写失败：{err.strerror or err}（{err.filename or ''}）"
    return str(err) or type(err).__name__


def copy_one(src: Path, dst: Path, *, check_source_open: bool = True,
             before_replace: Callable[[], None] | None = None, expected_digest: str | None = None) -> str:
    """先复制成临时文件再替换：网络断开或中途出错时，目的地原来的文件不会只剩半个。
    返回复制过去的内容的哈希（复制前后核对过，和来源一致）"""
    if check_source_open and source_open(src):
        raise PermissionError(13, "来源文件正在 Excel 中打开（旁边有 ~$ 锁文件），可能有没保存的修改", str(src))
    if dst.exists() and workbook_locked(dst):
        raise PermissionError(13, "目的地的同名文件正在 Excel 中打开，没法覆盖", str(dst))
    dst.parent.mkdir(parents=True, exist_ok=True)
    temp = dst.with_name(f".{uuid.uuid4().hex}{TEMP_SUFFIX}")
    try:
        shutil.copy2(src, temp)
        digest = _digest(temp)
        if temp.stat().st_size != src.stat().st_size or digest != _digest(src):
            raise OSError(5, "复制后内容和来源不一致", str(dst))
        if expected_digest is not None and digest != expected_digest:
            raise ValueError("待上传的文件在检查后发生变化，请重新检查这份文件")
        if before_replace:
            before_replace()
        os.replace(temp, dst)
        return digest
    finally:
        try:
            temp.unlink()
        except FileNotFoundError:
            pass
        except OSError:
            pass


def _progress(done: int, total: int, current: str) -> None:
    with _LOCK:
        _STATUS.update(done=done, total=total, current=current)


def _make_dirs(folder: dict[str, Any]) -> None:
    """先按目录结构把目的地文件夹建好（包括空文件夹），建了哪些、哪些没建成都记下来"""
    folder["created_dirs"], folder["dir_errors"] = [], []
    dst = Path(folder["dest_path"])
    for rel in folder.get("dirs_to_create", []):
        path = dst if rel == "." else dst / rel
        while True:
            try:
                existed = path.is_dir()
                path.mkdir(parents=True, exist_ok=True)
                if not existed:
                    folder["created_dirs"].append(rel)
                    debug_log.event("📁 新建文件夹", 文件夹=str(path))
                break
            except OSError as err:
                message = _explain(err)
                debug_log.exception(err, f"新建文件夹 {path}")
                if file_problem.ask(path, f"{folder['dest_rel']}/{rel}", message, folder=dst.parent) == "retry":
                    continue
                folder["dir_errors"].append(f"{rel}：{message}")
                break


def _stale_sources(preview_data: dict[str, Any]) -> list[str]:
    """复制前再查一遍来源：和确认过的清单不一样（多了、少了、改过）就不复制"""
    changed = []
    for folder in preview_data["folders"]:
        if folder["error"]:
            continue
        src = Path(folder["source_path"])
        if not src.is_dir():
            changed.append(f"{folder['source_rel']}：来源文件夹不见了")
            continue
        now = scan_folder(src)["excel"]
        then = {f["rel"]: f for f in folder["files"]}
        for rel in sorted(set(now) | set(then)):
            if rel not in now:
                changed.append(f"{folder['source_rel']}/{rel}：清单里有，现在没有了")
            elif rel not in then:
                changed.append(f"{folder['source_rel']}/{rel}：清单之后新出现的")
            elif now[rel]["size"] != then[rel]["size"] or abs(now[rel]["mtime"] - then[rel]["mtime"]) > 1:
                changed.append(f"{folder['source_rel']}/{rel}：清单之后被修改过")
    return changed


def _guard_again(src_file: Path, dest_file: Path, item: dict[str, Any], base: dict[str, Any] | None) -> str:
    """复制这一个文件之前再查一次（生成清单之后别人可能刚改过）。返回不能覆盖的原因，能覆盖返回空"""
    exists = dest_file.exists()
    if item.get("guard") == "content_same":  # 清单时按内容核对过：目的地这之后不能再变
        if not exists or _shared_digest(dest_file) != item.get("dest_sha"):
            return "生成清单之后目的地又被改过，没有覆盖。请重新生成清单查看"
        return ""
    if not base:
        if exists and item.get("guard") == "none":
            return "生成清单之后目的地出现了同名文件（可能是别人刚放进去的），没有覆盖。请重新生成清单查看"
        return ""
    return _guard(src_file, dest_file, exists, base)["conflict"]


def _upload_task(task_id: Any) -> dict[str, Any]:
    from core import work_task

    task = work_task.get(str(task_id))
    if work_task.overview()["current"] != task["id"]:
        raise ValueError("当前任务已切换，请重新打开「上传 Excel」")
    return task


def task_files(task_id: Any) -> dict[str, Any]:
    """只列本地文件和已有记录。选定文件后才读取共享端内容。"""
    from core import work_task
    task = _upload_task(task_id)
    work, shared = Path(task["work"]), Path(task["shared"])
    dest_root = work_task.upload_dest(task)
    if not work.is_dir():
        raise ValueError("处理文件夹不存在，请确认路径")
    baseline = load_baseline(strict=True)
    found = scan_folder(work)
    old_upload_dates = {}
    for run in reversed(_load_history()):
        if run.get("kind") != "copy":
            continue
        for folder in run["folders"]:
            for file in folder["files"]:
                if file.get("status") == "ok":
                    key = _base_key(Path(folder["dest_path"]) / file["dest_rel"])
                    old_upload_dates[key] = run.get("finished") or ""
    files = []
    for rel in sorted(found["excel"], key=_key):
        review = baseline.get(_base_key(shared / rel), {}).get("tasks", {}).get(task["id"], {}).get("review")
        key = _base_key(dest_root / rel)  # 上传记录按终点记（中转时是团队文件夹）
        base = baseline.get(key, {})
        upload = base.get("upload")
        if not upload and base.get("how") == "复制":  # 旧版只在版本快照里标记复制
            upload = {"sha256": base["sha256"], "at": ""}
        same = None
        if upload:
            try:
                same = _digest(_task_file(work, rel)) == upload["sha256"]
            except (OSError, ValueError):
                pass
        # 同名 PDF 上次什么时候传的（只看记录，不读共享）：只传 PDF 时按它挑「待上传」
        pdf_upload = baseline.get(_base_key(_pdf_target(dest_root / rel)), {}).get("upload") or {}
        files.append({"rel": rel, "review": review, "uploaded": bool(upload) or key in old_upload_dates,
                      "uploaded_at": (upload or {}).get("at") or old_upload_dates.get(key, ""), "uploaded_current": same,
                      "pdf_uploaded_at": pdf_upload.get("at", "")})
    return {"files": files, "errors": found["errors"], "dest": str(dest_root), "relay": work_task.is_relay(task)}


def _task_baseline(base: dict[str, Any], task_id: str) -> dict[str, Any]:
    versions = base.get("tasks", {})
    return versions.get(task_id, {}) if task_id and versions else base


def _task_file(root: Path, rel: str) -> Path:
    if (not isinstance(rel, str) or "\\" in rel or ":" in rel or "\0" in rel
            or any(part in ("", ".", "..") for part in rel.split("/")) or not is_excel(rel.rsplit("/", 1)[-1])):
        raise ValueError("需要处理文件夹下的有效 Excel 相对路径")
    path = root / rel
    if not _inside(path.resolve(), root.resolve()):
        raise ValueError(f"文件指向任务文件夹之外：{rel}")
    return path


def _guard_mode(value: Any) -> str:
    """默认先核对基准哈希。content 模式逐份做 Excel 对比。"""
    return "content" if value == "content" else "auto"


def start_task_upload(task_id: Any, files: Any, reviewed: Any, pdf: Any = None, guard: Any = None) -> None:
    task = _upload_task(task_id)
    pdf_settings = _pdf_settings(pdf)
    if not isinstance(files, list) or not files:
        raise ValueError("请至少选择一个 Excel 文件")
    if not isinstance(reviewed, dict):
        raise ValueError("复检确认无效，请刷新文件列表")
    work = Path(task["work"])
    items = []
    seen = set()
    for rel in files:
        local = _task_file(work, rel)
        if not local.is_file() or _key(rel) in seen:
            raise ValueError(f"文件不存在或重复选择：{rel}")
        seen.add(_key(rel))
        items.append({"rel": rel, "reviewed": reviewed.get(rel)})
    from core import work_task
    start(upload={"kind": "task_upload", "task": task, "label": f"任务：{task['name']}",
                  "source_root": task["work"], "dest_root": str(work_task.upload_dest(task)),
                  "shared": task["shared"], "relay": work_task.is_relay(task), "files": items, "pdf": pdf_settings,
                  "guard": _guard_mode(guard)})


def _shared_digest(path: Path) -> str | None:
    try:
        return _digest(path)
    except FileNotFoundError:
        return None


class _SharedVersionChanged(ValueError):
    pass


def _verify_upload_candidate(candidate: Candidate) -> None:
    item, local, remote, expected = candidate
    if _shared_digest(remote) != expected:
        raise _SharedVersionChanged("远程文件在检查后发生变化，请重新检查这份文件")
    if _digest(local) != item["src_sha256"]:
        raise ValueError("本地文件在检查后发生变化，请重新检查这份文件")


def _verified_uploads(batch: list[Candidate]) -> list[Candidate]:
    ready = []
    for candidate in batch:
        try:
            _verify_upload_candidate(candidate)
        except (OSError, ValueError) as err:
            candidate[0].update(status="conflict" if isinstance(err, _SharedVersionChanged) else "failed",
                                error=_explain(err))
        else:
            ready.append(candidate)
    return ready


def _upload_review_reason(known: str | None, current: str | None, guard: str) -> str:
    if current is None:
        return "远程原有的同名文件已移走或删除，请确认是否重新创建" if known else ""
    if guard == "content":
        return "已选择逐份做 Excel 对比"
    if not known:
        return "没有这份文件的基准哈希"
    return "远程版本和本地开工时的基准不同" if current != known else ""


def _check_task_file(task: dict[str, Any], item: dict[str, Any], local: Path, shared: Path,
                     baseline: dict[str, Any], guard: str) -> str | None:
    if not Path(task["shared"]).is_dir():
        raise ValueError("共享文件夹不可用，请确认网络连接或同步状态")
    version = baseline.get(_base_key(shared), {}).get("tasks", {}).get(task["id"], {})
    review = version.get("review")
    known = version.get("sha256")
    reference = _task_file(Path(task["reference"]), item["rel"])
    if not known and reference.is_file():
        known = _digest(reference)
    current = _shared_digest(shared)
    reason = _upload_review_reason(known, current, guard)
    if review and (not review["ready"] or current != review["sha256"]):
        reason = reason or "此前的共享版本复检尚未完成"
    if reason:
        item["review_reason"] = reason
    elif review and review["token"] != item["reviewed"]:
        item.update(status="review", error="请确认已在共享最新版上重新处理并复检")
    item["src_sha256"] = _digest(local)
    return current


def _sheet_progress(index: int, total: int, rel: str, guard: Any) -> Callable[[int, int, str], None]:
    def on_sheet(i: int, count: int, sheet: str) -> None:
        guard.progress()
        _progress(index, total, f"导出同名 PDF {rel}：第 {i} / {count} 张表「{sheet}」")
    return on_sheet


def _prepare_task_files(task: dict[str, Any], checked: list[tuple], total: int,
                        pdf_settings: dict[str, Any] | None = None, temp_root: Path | None = None) -> None:
    """上传前处理。只传 PDF（temp_root 给了）时在本机临时副本上做，处理文件夹里的 Excel 一点不动
    （用户 2026-11-10：不然保存后哈希变了，「当前版本已上传」会变成「本地有新版本待上传」）"""
    from core import upload_prep
    from core.excel_ops import quit_background_excel, start_background_excel

    excel = None
    guard = file_problem.Guard("上传前处理（清标记 · 恢复颜色 · 视图标准化）")

    def get_excel() -> Any:
        nonlocal excel
        if excel is None:
            excel = guard.watch(start_background_excel())
        return excel

    wf.pythoncom.CoInitialize()
    try:
        for index, (item, local, _dest, _expected) in enumerate(checked):
            reference = _task_file(Path(task["reference"]), item["rel"]) if task.get("reference") else None
            while True:  # 出错时问用户：「重试」重开 Excel 再做这份，「跳过」这份就不传（其余照传，用户 2026-11-10）
                _progress(index, total, f"上传前处理 {item['rel']}")
                where: dict[str, str] = {}
                try:
                    current = _digest(local)
                    expected_source = item.get("_prepared_sha256", item["src_sha256"]) if temp_root is None else item["src_sha256"]
                    if current != expected_source:
                        item.update(status="failed", error="本地文件在检查后发生变化，请重新检查这份文件")
                        break
                    if guard.file(item["rel"]):  # 进度框能「跳过这一份」
                        excel = None
                    pre = current if pdf_settings else ""
                    cached = pdf_settings and _use_cached_pdf(item, local, _pdf_target(local), pre, pdf_settings)
                    if cached and temp_root is not None:
                        # 只传 PDF、Excel 没变过、PDF 是按这一版导的：不再处理、不再导，直接传这份 PDF（用户 2026-11-10）
                        break
                    target = local
                    if temp_root is not None:
                        target = temp_root / str(index) / local.name
                        target.parent.mkdir(parents=True, exist_ok=True)
                        copy_one(local, target)  # 处理文件夹里那份在 Excel 里开着（有没保存的修改）时这里会报错
                    item["prep"] = upload_prep.prepare(get_excel, target, reference, where)
                    item["_prepared_sha256"] = _digest(target)
                    wf.log_msg(f"🧽 {item['rel']}：{upload_prep.text(item['prep'])}")
                    if cached:  # 表格和 PDF 一起：Excel 照样处理、上传，PDF 沿用；处理保存后的 Excel 也认这份 PDF
                        pdf_cache.add_excel_hash(_pdf_target(local), _digest(local))
                    elif pdf_settings:  # 清完标记、视图标准化之后再导，PDF 里不会有标记色
                        where["step"] = "导出同名 PDF"
                        _export_pdf(get_excel, item, target, pdf_settings, _sheet_progress(index, total, item["rel"], guard))
                        if temp_root is not None:  # 处理文件夹里也留一份 PDF，和「一起传」时一样
                            shutil.copyfile(item["_pdf_local"], _pdf_target(local))
                            item["_pdf_local"] = str(_pdf_target(local))
                        # 记下这份 PDF 是哪版 Excel 导的：处理前那版（只传 PDF 时处理文件夹里就是这版）和处理保存后那版
                        pdf_cache.remember(_pdf_target(local), [pre, _digest(local)], pdf_settings["options"],
                                           item["pdf"], "upload")
                    break
                except Exception as err:
                    debug_log.exception(err, f"上传前处理 {item['rel']}")
                    retry, message = guard.problem(err, local, item["rel"], Path(task["work"]),
                                                   close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {item['rel']}")
                        continue
                    item.update(status="failed", error=f"上传前{where.get('step') or '处理'}没做成：{message}")
                    break
    finally:
        try:
            if excel is not None:
                quit_background_excel(excel)
        finally:
            guard.stop()
            wf.pythoncom.CoUninitialize()


# ───────────────────────── 同名 PDF 一起上传（用户 2026-11-10） ─────────────────────────
# 上传时可以勾「导出同名 PDF 并一起上传」：每份 Excel 上传前处理做完后，用「批量导出同名 PDF」（core/pdf_bundle.py）
# 同样的做法导出 aaa.pdf（所有表合一份、强制宽一页、跳过名单、去页眉页脚），Excel 传成功后 PDF 放到终点 Excel 旁边。
# 终点已有同名 PDF 时默认不覆盖：上传前先列出来（pdf_check），用户逐个或全部选「覆盖 / 跳过」才开始；
# 选了覆盖的，上传那一刻 PDF 和检查时不一样（期间被人换了）也不覆盖。PDF 出问题不影响 Excel 上传，只在结果里写明。

def _pdf_settings(value: Any) -> dict[str, Any] | None:
    """网页传来的 PDF 选项：None / False = 不导出（只传表格）；
    {"options": {...}, "overwrite": {rel: 检查时的哈希}, "only": True = 只传 PDF、Excel 不传}"""
    if value in (None, False):
        return None
    if not isinstance(value, dict):
        raise ValueError("同名 PDF 选项格式不正确")
    from core import pdf_bundle
    options = pdf_bundle.options_for(value.get("options"))
    overwrite = value.get("overwrite") or {}
    if not isinstance(overwrite, dict) or any(not isinstance(k, str) or not isinstance(v, str) for k, v in overwrite.items()):
        raise ValueError("要覆盖的同名 PDF 清单格式不正确，请重新检查")
    only = value.get("only", False)
    if type(only) is not bool:
        raise ValueError("同名 PDF 选项格式不正确")
    return {"options": options, "overwrite": {_key(rel): sha for rel, sha in overwrite.items()}, "only": only}


def _pdf_target(excel_target: Path) -> Path:
    from core import pdf_bundle
    return pdf_bundle.pdf_path_for(excel_target)


def pdf_check(task_id: Any = None, files: Any = None, preview_id: Any = None) -> dict[str, Any]:
    """上传前看终点 Excel 旁边已经有没有同名 PDF（只读）。ours = 上次就是本工具传上去的、之后没人动过"""
    if preview_id is not None:
        with _LOCK:
            data = copy.deepcopy(_STATUS.get("preview"))
        if not data or data.get("kind") != "free_upload" or data.get("id") != preview_id:
            raise ValueError("清单已变化，请重新生成")
        dest_root = Path(data["dest_root"])
        rels = [f["rel"] for f in data["files"] if f["action"] != "conflict"]
    else:
        from core import work_task
        task = _upload_task(task_id)
        if not isinstance(files, list) or not files:
            raise ValueError("请至少选择一个 Excel 文件")
        dest_root = work_task.upload_dest(task)
        rels = [rel for rel in files if _task_file(Path(task["work"]), rel)]
    if not dest_root.is_dir():
        raise ValueError(f"终点文件夹不可用：{dest_root}（OneDrive 的话请确认已经同步到这台电脑）")
    baseline = load_baseline(strict=True)
    found, same_name = [], {}
    for rel in rels:
        target = _pdf_target(dest_root / rel)
        pdf_rel = target.relative_to(dest_root).as_posix()
        if _base_key(target) in same_name:  # aaa.xlsx 和 aaa.xlsm 在一起：只传第一份的 PDF
            found.append({"rel": rel, "pdf": pdf_rel, "exists": False, "clash": same_name[_base_key(target)]})
            continue
        same_name[_base_key(target)] = rel
        entry: dict[str, Any] = {"rel": rel, "pdf": pdf_rel, "exists": False}
        try:
            if target.is_file():
                stat = target.stat()
                digest = _digest(target)
                uploaded = (baseline.get(_base_key(target), {}).get("upload") or {}).get("sha256")
                entry.update(exists=True, sha256=digest, size=stat.st_size, ours=digest == uploaded,
                             mtime=datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M"))
        except OSError as err:
            entry.update(exists=True, sha256="", error=_explain(err))
        found.append(entry)
    existing = [f for f in found if f["exists"]]
    debug_log.event("📕 上传前查终点同名 PDF", 终点=str(dest_root), 文件=len(rels), 已有=len(existing),
                    上次本工具传的=sum(bool(f.get("ours")) for f in existing))
    return {"dest": str(dest_root), "files": found}


def _export_pdf(get_excel: Callable[[], Any], item: dict[str, Any], local: Path, settings: dict[str, Any],
                on_sheet: Callable[[int, int, str], None]) -> None:
    """上传前处理做完后导出同名 PDF（放在 local 旁边），出错原样抛出，调用方按 file_problem 重试 / 跳过"""
    from core import pdf_bundle
    opts = settings["options"]
    info = pdf_bundle.export_workbook(get_excel(), local, opts["include_hidden"], on_sheet, opts["skip_sheets"],
                                      opts["no_header_footer"])
    item["_pdf_local"] = info["pdf"]
    item["pdf"] = {"status": "exported", "name": Path(info["pdf"]).name, "pages": info["pages"],
                   "sheets": len(info["sheets"]), "skipped": info["skipped"]}
    skipped = "、".join(f"{s['sheet']}（{s['reason'].split('，')[0].split('：')[0]}）" for s in info["skipped"])
    wf.log_msg(f"📕 {item['rel']}：导出 {Path(info['pdf']).name}，{len(info['sheets'])} 张表、{info['pages']} 页"
               + (f"；跳过 {skipped}" if skipped else ""))
    debug_log.event("📕 上传前导出同名 PDF", 文件=item["rel"], PDF=info["pdf"], 表=len(info["sheets"]), 页=info["pages"],
                    跳过=[f"{s['sheet']}（{s['reason']}）" for s in info["skipped"]])


def _use_cached_pdf(item: dict[str, Any], excel: Path, pdf: Path, excel_hash: str, settings: dict[str, Any]) -> bool:
    """Excel 的哈希和上次导出 PDF 时记下的一样、导出设置一样、PDF 没被换过 → 直接用这份 PDF，不重新导（core/pdf_cache.py）"""
    hit = pdf_cache.lookup(pdf, excel_hash, settings["options"])
    if not hit:
        return False
    item["_pdf_local"] = str(pdf)
    item["pdf"] = {"status": "exported", "name": pdf.name, "pages": hit.get("pages", 0), "sheets": hit.get("sheets", 0),
                   "skipped": hit.get("skipped") or [], "cached": True}
    wf.log_msg(f"📕 {item['rel']}：Excel 没变，沿用之前导出的 {pdf.name}（{item['pdf']['pages']} 页），不重新导")
    debug_log.event("♻️ 同名 PDF 沿用缓存", 文件=item["rel"], PDF=str(pdf), Excel哈希=excel_hash[:12],
                    来自="批量导出同名 PDF" if hit.get("source") == "bundle" else "上次上传", 导出于=hit.get("at", ""))
    return True


def _upload_pdf(item: dict[str, Any], excel_target: Path, settings: dict[str, Any], claimed: set[str],
                task_id: str = "") -> dict[str, dict[str, Any]]:
    """Excel 传成功后把同名 PDF 放到它旁边。返回要记下的版本（{key: entry}）"""
    pdf = item.get("pdf") or {}
    src = item.pop("_pdf_local", "")
    if not src or pdf.get("status") != "exported":
        return {}
    dest = _pdf_target(excel_target)
    key = _base_key(dest)
    label = f"{item['rel']} → {dest.name}"
    if key in claimed:
        pdf.update(status="skipped", error=f"和同一批里的另一份 Excel 会得到同名的 {dest.name}，这份的 PDF 没传，请改名后再传")
        debug_log.event("⏭ 同名 PDF 撞名没传", 文件=label)
        return {}
    claimed.add(key)
    allowed = settings["overwrite"].get(_key(item["rel"]))
    while True:
        try:
            current = _shared_digest(dest)
            if current is not None and allowed is None:
                pdf.update(status="kept", error=f"终点已有同名 {dest.name}，没有覆盖（上传前没选「覆盖」）")
                debug_log.event("🛑 终点已有同名 PDF，没覆盖", 文件=label)
                return {}
            if current is not None and allowed != current:
                pdf.update(status="kept", error=f"终点的 {dest.name} 在检查之后又被改过，没有覆盖，请重新检查")
                debug_log.event("🛑 同名 PDF 检查之后被改过，没覆盖", 文件=label)
                return {}

            def verify() -> None:
                if _shared_digest(dest) != current:
                    raise _SharedVersionChanged

            digest = copy_one(Path(src), dest, check_source_open=False, before_replace=verify)
            pdf.update(status="ok", action="overwrite" if current is not None else "new", path=str(dest))
            debug_log.event("📕 同名 PDF 传好了", 文件=label, 做法="覆盖" if current is not None else "新建")
            return {key: _base_entry(dest, digest, "复制", task_id)}
        except _SharedVersionChanged:
            pdf.update(status="kept", error=f"上传期间终点出现了另一份 {dest.name}，没有覆盖，请重新检查")
            return {}
        except (OSError, ValueError) as err:
            message = _explain(err)
            if isinstance(err, PermissionError):
                message = f"{dest.name} 正被占用（可能在 PDF 阅读器里开着），关掉后点「重试」：{err.filename or dest}"
            debug_log.exception(err, f"上传同名 PDF {label}")
            if file_problem.ask(dest, label, message, folder=dest.parent) == "retry":
                wf.log_msg(f"🔁 重试 {label}")
                continue
            pdf.update(status="failed", error=f"PDF 没传成：{message}")
            return {}


def _pdf_only_status(item: dict[str, Any]) -> None:
    """只传 PDF 时，这一份的结果就看 PDF：传了或按你的选择没覆盖算做完，没传成算失败"""
    pdf = item.get("pdf") or {}
    item["pdf_only"] = True
    if pdf.get("status") in ("ok", "kept"):
        item.update(status="ok", error="")
    else:
        item.update(status="failed", error=pdf.get("error") or "同名 PDF 没传成")


def _pdf_summary(items: list[dict[str, Any]]) -> dict[str, int]:
    pdfs = [item["pdf"] for item in items if item.get("pdf")]
    if not pdfs:
        return {}
    return {"PDF 已传": sum(p["status"] == "ok" for p in pdfs),
            "PDF 没覆盖": sum(p["status"] == "kept" for p in pdfs),
            "PDF 没传成": sum(p["status"] in ("failed", "skipped") for p in pdfs)}


def _pdf_text(pdf: dict[str, Any] | None) -> str:
    """一行说明，网页和上传记录共用"""
    if not pdf:
        return ""
    if pdf["status"] == "ok":
        return (f"同名 PDF {pdf['name']} 已{'覆盖' if pdf.get('action') == 'overwrite' else '新建'}（{pdf['sheets']} 张表、{pdf['pages']} 页"
                + ("，Excel 没变，沿用之前导出的" if pdf.get("cached") else "") + "）")
    if pdf["status"] == "exported":
        return f"同名 PDF {pdf['name']} 已导出，但这份没传上去，PDF 也没传"
    if pdf["status"] == "stale":
        return f"同名 PDF {pdf['name']} 是按同步前那一版导的，没传，处理文件夹里那份也删掉了"
    return pdf.get("error") or "同名 PDF 没传"


def _run_task_upload(data: dict[str, Any], reviews: UploadReview) -> dict[str, Any]:
    task = data["task"]
    work, shared = Path(task["work"]), Path(task["shared"])
    relay = Path(data["dest_root"]) if data.get("relay") else None
    if relay is not None and not relay.is_dir():
        raise ValueError(f"终点文件夹不可用：{relay}（OneDrive 的话请确认已经同步到这台电脑）")
    items = data["files"]
    baseline = load_baseline(strict=True)
    ready = []
    for index, item in enumerate(items):
        _progress(index, len(items), f"检查 {item['rel']}")
        try:
            local, dest = _task_file(work, item["rel"]), _task_file(shared, item["rel"])
            current = _check_task_file(task, item, local, dest, baseline, data["guard"])
            if not item.get("status"):
                if relay is not None:
                    item["_relay_sha256"] = _shared_digest(_task_file(relay, item["rel"]))
                ready.append((item, local, dest, current))
        except (OSError, ValueError) as err:
            item.update(status="failed", error=_explain(err))

    pdf = data.get("pdf")
    temp_root = None
    if (pdf or {}).get("only"):
        import tempfile
        temp_root = Path(tempfile.mkdtemp(prefix="只传PDF_"))
        data["_temp_root"] = str(temp_root)
    pdf_claimed: set[str] = set()
    relayed: dict[str, dict[str, Any]] = {}
    for batch_index, batch in enumerate(reviews.batches(ready, _verify_upload_candidate)):
        verified = _verified_uploads(batch)
        if not verified:
            continue
        _prepare_task_files(task, verified, len(items), pdf, temp_root / str(batch_index) if temp_root else None)
        for item, local, dest, expected in verified:
            if item.get("status"):
                continue
            _progress(sum(bool(f.get("status")) for f in items), len(items), f"上传 {item['rel']}")
            target = _task_file(relay, item["rel"]) if relay is not None else dest

            def verify_shared() -> None:
                if not shared.is_dir() or _shared_digest(dest) != expected:
                    raise _SharedVersionChanged("上传前共享版本发生变化，请重新检查这份文件")
                if relay is not None and (not relay.is_dir() or _shared_digest(target) != item["_relay_sha256"]):
                    raise _SharedVersionChanged("上传前中转版本发生变化，请重新检查这份文件")

            try:
                if (pdf or {}).get("only"):
                    verify_shared()
                    if _digest(local) != item["src_sha256"]:
                        raise ValueError("本地文件在检查后发生变化，请重新检查这份文件")
                    _save_baseline(_upload_pdf(item, target, pdf, pdf_claimed, task["id"]))
                    _pdf_only_status(item)
                    continue
                digest = copy_one(local, target, before_replace=verify_shared,
                                  expected_digest=item["_prepared_sha256"])
            except (OSError, ValueError) as err:
                item.update(status="failed", error=_explain(err))
                continue
            entry = _base_entry(target, digest, "复制", task["id"])
            item.update(status="ok", uploaded_at=entry["upload"]["at"])
            entry["tasks"][task["id"]]["review"] = None
            changes = {_base_key(target): entry}
            if pdf:
                _progress(sum(bool(f.get("status")) for f in items), len(items), f"上传同名 PDF {item['rel']}")
                changes.update(_upload_pdf(item, target, pdf, pdf_claimed, task["id"]))
            if relay is not None:
                changes[_base_key(dest)] = {"tasks": {task["id"]: {"review": None}}}
                relayed[item["rel"]] = {"shared_sha256": expected, "sha256": digest, "uploaded_at": entry["upload"]["at"]}
            _save_baseline(changes)
        if relayed:
            try:
                manifest = write_manifest(relay, shared, relayed, task["name"])
                data["manifest"] = str(manifest)
                data.pop("manifest_error", None)
                relayed.clear()
            except (OSError, ValueError) as err:
                debug_log.exception(err, "写拿取版本清单")
                data["manifest_error"] = _explain(err)
    done = [item["rel"] for item in items if item.get("status") == "ok" and not item.get("pdf_only")]
    if done:  # 本轮清单里的文件传上去了就划掉，全划掉自动回到整个文件夹
        from core import work_task
        try:
            work_task.clear_focus(task["id"], done, why="已上传")
        except Exception as err:
            debug_log.event("⚠️ 没能更新本轮清单", 原因=str(err))
    summary = {"已复制": sum(item.get("status") == "ok" for item in items),
               "待重新处理": sum(item.get("status") == "review" for item in items),
               "已跳过": sum(item.get("status") == "skipped" for item in items),
               "没复制成功": sum(item.get("status") in ("failed", "conflict") for item in items), **_pdf_summary(items)}
    for item in items:
        item.pop("reviewed", None)
        item.pop("_prepared_sha256", None)
        item.pop("_relay_sha256", None)
        item.pop("_pdf_local", None)
    return {**data, "id": uuid.uuid4().hex[:12], "summary": summary,
            "finished": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}


def _run(preview_data: dict[str, Any]) -> dict[str, Any]:
    folders = copy.deepcopy(preview_data["folders"])
    total = sum(len(f["files"]) for f in folders)
    done = 0
    stop_reason = ""
    baseline = load_baseline()
    stale = _stale_sources(preview_data)
    if stale:
        stop_reason = "来源文件在生成清单之后变了，一个都没有复制，请重新生成清单"
        debug_log.event("⛔ " + stop_reason, 变化=stale[:30])
    for folder in folders:
        if stop_reason:
            break
        _make_dirs(folder)
        if folder["error"]:
            wf.log_msg(f"⚠️ {folder['source_rel']}：{folder['error']}，"
                       + ("已建好空的目的地文件夹" if folder["created_dirs"] else "跳过"))
            continue
        src_dir, dst_dir = Path(folder["source_path"]), Path(folder["dest_path"])
        debug_log.begin(f"复制文件夹 {folder['source_rel']} → {folder['dest_rel']}")
        changes: dict[str, dict[str, Any]] = {}
        try:
            for item in folder["files"]:
                label = f"{folder['source_rel']}/{item['rel']}"
                _progress(done, total, label)
                dest_file = dst_dir / item["dest_rel"]
                conflict = item.get("conflict") or _guard_again(src_dir / item["rel"], dest_file, item,
                                                                baseline.get(_base_key(dest_file)))
                if conflict:
                    item.update(status="conflict", error=conflict)
                    done += 1
                    debug_log.event("🛑 别人改过，没有覆盖", 文件=label, 原因=conflict[:60])
                    continue
                while True:
                    try:
                        digest = copy_one(src_dir / item["rel"], dest_file)
                        item.update(status="ok", error="")
                        if isinstance(digest, str):
                            task_id = (preview_data.get("task") or {}).get("id", "")
                            changes[_base_key(dest_file)] = _base_entry(dest_file, digest, "复制", task_id)
                        break
                    except Exception as err:
                        message = _explain(err)
                        debug_log.exception(err, f"复制 {label}")
                        bad = getattr(err, "filename", None) or str(src_dir / item["rel"])
                        choice = file_problem.ask(bad, label, message, folder=Path(bad).parent)
                        if choice == "retry":
                            wf.log_msg(f"🔁 重试 {label}")
                            continue
                        item.update(status="failed", error=message)
                        break
                done += 1
                action = "覆盖" if item["action"] == "overwrite" else "新建"
                if item["status"] == "ok":
                    debug_log.event(f"📄 复制完成（{action}）", 文件=label, 大小=item["size"],
                                    原大小=item["old_size"] if item["old_size"] is not None else "-")
                else:
                    debug_log.event("❌ 没复制成功", 文件=label, 原因=item["error"])
            after = scan_folder(dst_dir) if dst_dir.is_dir() else {"excel": {}, "errors": []}
            folder["dest_count"] = len(after["excel"])
            after_keys = {_key(rel) for rel in after["excel"]}
            folder["missing_in_dest"] = [f["rel"] for f in folder["files"] if _key(f["dest_rel"]) not in after_keys]
            source_keys = {_key(f["rel"]) for f in folder["files"]}
            folder["dest_extra"] = sorted(rel for rel in after["excel"] if _key(rel) not in source_keys)
            folder["scan_errors"] = folder["scan_errors"] + after["errors"]
            folder["checks"] = _checks(folder, folder["source_count"], folder["dest_count"])
            bad_checks = [c["text"] for c in folder["checks"] if not c["ok"]]
            ok_count = sum(f.get("status") == "ok" for f in folder["files"])
            conflict_count = sum(f.get("status") == "conflict" for f in folder["files"])
            wf.log_msg(f"📁 {folder['source_rel']} → {folder['dest_rel']}：复制 {ok_count}/{len(folder['files'])}"
                       + (f"；🛑 别人改过没覆盖 {conflict_count} 个" if conflict_count else "")
                       + (f"；数量不一致：{'；'.join(bad_checks)}" if bad_checks else "；三个数一致"))
            debug_log.event("🔢 核对数量", 文件夹=folder["dest_rel"], JSON=folder["json_files"],
                            来源=folder["source_count"], 目的地=folder["dest_count"],
                            不一致=bad_checks or "无", 目的地缺=folder["missing_in_dest"][:20],
                            目的地多=folder["dest_extra"][:20])
        finally:
            try:
                _save_baseline(changes)
            except Exception as err:
                debug_log.exception(err, "保存版本记录")
            debug_log.end()
    result = {
        "id": uuid.uuid4().hex[:12],
        "kind": "copy",
        "task": preview_data.get("task"),
        "label": preview_data["label"],
        "source_root": preview_data["source_root"],
        "dest_root": preview_data["dest_root"],
        "notes": preview_data["notes"],
        "folders": folders,
        "stopped": stop_reason,
        "stale": stale,
        "started": preview_data.get("_started", ""),
        "finished": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    for folder in folders:
        folder.setdefault("missing_in_dest", [])
        folder.setdefault("created_dirs", [])
        folder.setdefault("dir_errors", [])
        if folder["error"] or stop_reason:
            folder["checks"] = _checks(folder, folder["source_count"], None)
    result["summary"] = run_summary(result)
    return result


def run_summary(result: dict[str, Any]) -> dict[str, int]:
    files = [f for folder in result["folders"] for f in folder["files"]]
    ok = [f for f in files if f.get("status") == "ok"]
    return {
        "文件夹": len(result["folders"]),
        "已复制": len(ok),
        "新建文件夹": sum(len(folder.get("created_dirs", [])) for folder in result["folders"]),
        "覆盖": sum(f["action"] == "overwrite" for f in ok),
        "新建": sum(f["action"] == "new" for f in ok),
        "别人改过没覆盖": sum(f.get("status") == "conflict" for f in files),
        "没复制成功": sum(f.get("status") == "failed" for f in files),
        "没有复制": sum("status" not in f for f in files),
        "数量不一致": sum(any(not c["ok"] for c in folder["checks"]) for folder in result["folders"]),
        "有问题的文件夹": sum(bool(folder["error"] or folder.get("dir_errors")) for folder in result["folders"]),
    }


def needs_attention(result: dict[str, Any]) -> bool:
    s = result["summary"]
    if result.get("kind") in ("task_upload", "free_upload"):
        return bool(result.get("stopped") or s.get("已复制", 0) != len(result["files"]) or s.get("PDF 没传成"))
    if result.get("kind") == "fetch":
        return bool(result.get("stopped") or s.get("没拿成功") or s.get("没有拿") or s.get("有问题的文件夹"))
    return bool(result.get("stopped") or s["没复制成功"] or s["没有复制"] or s["数量不一致"]
                or s["有问题的文件夹"] or s.get("别人改过没覆盖") or result.get("notes"))


def start(preview_id: str | None = None, *, upload: dict[str, Any] | None = None) -> None:
    """按最后一次确认的清单开始复制"""
    with wf.STATE_LOCK, _LOCK:
        preview_data = upload or copy.deepcopy(_STATUS.get("preview"))
        if not preview_data or preview_id is not None and preview_id != (_STATUS.get("preview") or {}).get("id"):
            raise ValueError("清单已变化，请重新生成")
        from core import bl_links, hidden_repair, link_repair, mark_finder
        if (wf.STATE.get("busy") or _STATUS.get("running") or link_repair.status().get("running")
                or mark_finder.status().get("running") or hidden_repair.status().get("running")
                or bl_links.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束（别的任务可能正在改这些文件）")
        fetching = preview_data.get("kind") == "fetch"
        task_upload = preview_data.get("kind") == "task_upload"
        free_upload = preview_data.get("kind") == "free_upload"
        if fetching and not any(f["files"] for f in preview_data["folders"]):
            raise ValueError("目的地这些文件夹里没有 Excel 文件可拿")
        if not fetching and not task_upload and not free_upload and not any(f["files"] or f.get("dirs_to_create") for f in preview_data["folders"]):
            raise ValueError("清单里没有要复制的 Excel 文件，也没有要新建的文件夹")
        reviews = UploadReview(pdf_only=bool((preview_data.get("pdf") or {}).get("only"))) if task_upload or free_upload else None
        wf.STATE["busy"] = True
        wf.STATE["maintenance"] = "distribute"
        preview_data["_started"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        total = len(preview_data["files"]) if task_upload or free_upload else sum(len(f["files"]) for f in preview_data["folders"])
        _STATUS.update(running=True, done=0, total=total, current="", label=preview_data["label"],
                       kind=preview_data["kind"])
        _LIVE["data"] = preview_data  # 网页实时进度框：跑的时候按每个文件的 status 列出做完的
        if reviews is not None:
            _LIVE["reviews"] = reviews
        if task_upload:
            _STATUS.pop("preview", None)
        _STATUS.pop("applied", None)

    def run() -> None:
        from core import operation_history
        result: dict[str, Any] | None = None
        ids: list[int] = []
        try:
            with debug_log.side_task("distribute_fetch" if fetching else "distribute",
                                     {"source": preview_data["source_root"], "dest": preview_data["dest_root"],
                                      "label": preview_data["label"], "files": _STATUS.get("total")}):
                ids = operation_history.begin(preview_data["source_root" if fetching or task_upload else "dest_root"], ["distribute"])
                if task_upload:
                    result = _run_task_upload(preview_data, reviews)
                elif free_upload:
                    result = _run_free(preview_data, reviews)
                else:
                    with file_problem.task("从目的地拿最新版" if fetching else "按 JSON 复制 Excel"):
                        result = _run_fetch(preview_data) if fetching else _run(preview_data)
                _save_history(result)
                message = "批量上传" if task_upload or free_upload else "拿最新版" if fetching else "按 JSON 复制"
                debug_log.event(f"✅ {message}结束", **result["summary"])
            operation_history.finish(ids, "failed" if needs_attention(result) else "completed")
        except Exception as err:
            debug_log.exception(err, "按 JSON 复制")
            operation_history.finish(ids, "failed")
            result = result or {"id": "", "kind": preview_data["kind"], "task": preview_data.get("task"),
                                "files": [{"rel": f["rel"], "status": f.get("status"), "error": f.get("error")}
                                          for f in preview_data.get("files", [])],
                                "label": preview_data["label"], "folders": [], "notes": [],
                                "source_root": preview_data["source_root"], "dest_root": preview_data["dest_root"],
                                "stopped": f"程序出错：{err}（详细信息已写进调试日志）", "stale": [],
                                "summary": {"已复制": sum(f.get("status") == "ok" for f in preview_data["files"])}
                                if task_upload or free_upload else {}}
        finally:
            if reviews is not None:
                reviews.close()
            temp = preview_data.pop("_temp_root", "")
            if result is not None:
                result.pop("_temp_root", None)
            if temp:
                shutil.rmtree(temp, ignore_errors=True)
            with wf.STATE_LOCK, _LOCK:
                _STATUS.update(running=False, current="")
                _STATUS["log"] = _live_log()  # 做完也留着，刷新页面还能看到这次每个文件的结果
                _LIVE.clear()
                if result is not None:
                    _STATUS["applied"] = result
                    _STATUS.pop("preview", None)
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)

    try:
        threading.Thread(target=run, name="按JSON复制", daemon=True).start()
    except Exception:
        if reviews is not None:
            reviews.close()
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
            _LIVE.clear()
        raise


def _live_log() -> list[dict[str, Any]]:
    """正在复制 / 拿 / 上传的这批里已经有结果的文件（调用方持有 _LOCK）：给网页实时进度框逐条列出"""
    data = _LIVE.get("data") or {}
    if data.get("kind") in ("task_upload", "free_upload"):
        pairs = [(item.get("rel", ""), item) for item in data.get("files", [])]
    else:
        pairs = [(f"{folder.get('dest_rel' if data.get('kind') == 'fetch' else 'source_rel', '')}/{item.get('rel', '')}", item)
                 for folder in data.get("folders", []) for item in folder.get("files", [])]
    words = {"ok": "拿好了" if data.get("kind") == "fetch" else "传好了" if "upload" in str(data.get("kind")) else "复制好了",
             "conflict": "别人改过，没有覆盖", "review": "要重新处理并复检", "failed": "没成功",
             "skipped": "已跳过"}
    log = []
    for name, item in pairs:
        state = item.get("status")
        if state not in words:
            continue
        reason = item.get("error") or ("只传 PDF" if item.get("pdf_only") and state == "ok" else words[state])
        pdf = item.get("pdf") or {}
        if state == "ok" and pdf.get("status") not in (None, "exported"):
            reason += "；" + _pdf_text(pdf)
        bad_pdf = pdf.get("status") in ("failed", "skipped")
        log.append({"name": name, "ok": state == "ok" and not bad_pdf, "reason": reason, "path": "", "sheet": ""})
    return log[-500:]


def status() -> dict[str, Any]:
    with _LOCK:
        data = copy.deepcopy(_STATUS)
        if data.get("running") and _LIVE.get("data"):
            data["log"] = _live_log()
            if (reviews := _LIVE.get("reviews")) is not None:
                data["reviews"], data["comparing"] = reviews.status()
                data["review_task_id"] = (_LIVE["data"].get("task") or {}).get("id")
                data["done"] = sum(bool(f.get("status")) for f in _LIVE["data"]["files"])
    if data.get("preview"):
        data["preview"] = _public(data["preview"])
    return data


def _active_review(token: Any) -> UploadReview:
    if not isinstance(token, str) or not token:
        raise ValueError("需要有效的上传确认")
    with _LOCK:
        reviews = _LIVE.get("reviews")
    if reviews is None:
        raise ValueError("这份上传确认已结束，请回到上传页")
    return reviews


def upload_review(token: Any) -> dict[str, Any]:
    return _active_review(token).result(token)


def decide_upload_review(token: Any, decision: Any) -> None:
    if decision not in ("upload", "skip"):
        raise ValueError("请选择确认上传或跳过")
    _active_review(token).decide(token, decision == "upload")


# ───────────────────────── 拿取版本清单 + 自由上传（起点 → 终点） ─────────────────────────
# 用户 2026-11-07：几个账号分着做，张三把处理好的文件先传到团队 OneDrive（中转），李四再从团队传到公司 OneDrive。
# 张三上传到中转时，在中转文件夹里放一份「拿取版本.json」，记下每份文件当初从公司拿的是哪个版本；
# 李四从中转传到公司时读这份清单：公司那份还是这个版本就覆盖，版本变了就逐份对比并确认。
# 上传前在本机临时副本上清标记 + 视图标准化（保存的人就是最終更新者），起点那份不动。

MANIFEST_NAME = "拿取版本.json"
MANIFEST_NOTE = ("Excel 整理工具箱写的：记着这些文件是从共享文件夹哪个版本拿的。"
                 "从这里上传到共享时用来检查期间有没有人改过，请不要删除或改动。")


def read_manifest(folder: Path) -> dict[str, Any] | None:
    path = folder / MANIFEST_NAME
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as err:
        raise ValueError(f"「{path}」读不出来（可能被改坏了）：{err}") from err
    if not isinstance(data, dict) or not isinstance(data.get("shared"), str) or not isinstance(data.get("files"), dict):
        raise ValueError(f"「{path}」格式不对，请不要手动改它")
    return data


def write_manifest(relay: Path, shared: Path, entries: dict[str, dict[str, Any]], task_name: str = "") -> Path:
    """把这次传到中转的文件并进中转文件夹的拿取版本清单（同名文件以这次为准）"""
    old = read_manifest(relay) or {"shared": str(shared), "files": {}}
    if _base_key(old["shared"]) != _base_key(shared):
        raise ValueError(f"中转文件夹「{relay}」里已经有另一个共享文件夹（{old['shared']}）的拿取版本清单，"
                         "请给这个任务换一个中转文件夹")
    files = {rel: v for rel, v in old["files"].items() if _key(rel) not in {_key(r) for r in entries}}
    for rel, entry in entries.items():
        files[rel] = {**entry, **({"task": task_name} if task_name else {})}
    shared_files = sorted(scan_folder(shared)["excel"]) if shared.is_dir() else old.get("shared_files", [])
    data = {"说明": MANIFEST_NOTE, "version": 1, "shared": str(shared), "files": dict(sorted(files.items())),
            "shared_files": shared_files}
    path = relay / MANIFEST_NAME
    temp = path.with_name(f".{uuid.uuid4().hex}{TEMP_SUFFIX}")
    try:
        temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temp, path)
    finally:
        try:
            temp.unlink()
        except OSError:
            pass
    debug_log.event("🧾 写拿取版本清单", 文件=str(path), 这次=len(entries), 合计=len(files))
    return path


def _find_manifests(src: Path) -> list[tuple[Path, dict[str, Any]]]:
    """起点里（包括子文件夹）和起点往上几层的拿取版本清单，深的在前"""
    found: list[tuple[Path, dict[str, Any]]] = []
    for parent in list(src.parents)[:6]:
        data = read_manifest(parent)
        if data:
            found.append((parent, data))
            break
    for dirpath, _dirs, files in os.walk(src):
        if MANIFEST_NAME in files:
            found.append((Path(dirpath), read_manifest(Path(dirpath))))
    return sorted(found, key=lambda pair: len(pair[0].parts), reverse=True)


def _manifest_for(src_file: Path, manifests: list[tuple[Path, dict[str, Any]]]) -> tuple[Path, dict[str, Any], str] | None:
    for folder, data in manifests:
        if _inside(src_file, folder):
            return folder, data, src_file.relative_to(folder).as_posix()
    return None


_LOCAL_PLACES = {"desktop", "デスクトップ", "桌面", "downloads", "ダウンロード", "下载", "documents", "ドキュメント", "文档"}


def _is_local_shared(shared: Path) -> bool:
    """任务的「共享」其实是本机普通文件夹（在桌面 / 下载 / 文档里，路径里没有 OneDrive）→ 拿取版本清单没意义（用户 2026-11-12）"""
    parts = [part.casefold() for part in shared.parts]
    if any(part.startswith("onedrive") for part in parts):
        return False
    return any(part in _LOCAL_PLACES for part in parts)


def _free_relay_task(src: Path, dst: Path) -> dict[str, Any] | None:
    """起点是某个任务的处理文件夹（或其中的子文件夹）、终点不在那个任务的共享里 → 这次就是「传到中转」，
    要在终点写拿取版本清单（用户 2026-11-12：起点 → 终点传到团队中转也要有 JSON）。返回任务和对应的共享子路径"""
    from core import work_task
    for task in work_task.overview()["tasks"]:
        if not task.get("work") or not task.get("shared"):
            continue
        prefix = work_task._same_or_inside(src, task["work"])
        if prefix is None:
            continue
        shared = Path(task["shared"])
        if work_task._same_or_inside(dst, str(shared)) is not None or work_task._same_or_inside(shared, str(dst)) is not None:
            return None  # 直接传回共享（或终点包着共享），不是中转
        if _is_local_shared(shared):
            debug_log.event("🧾 共享是本机普通文件夹，不写拿取版本清单", 共享=str(shared))
            return None
        return {"id": task["id"], "name": task.get("name") or "", "shared": str(shared.joinpath(*prefix.split("/")) if prefix else shared)}
    return None


def _fetched_version(task: dict[str, Any], rel: str, baseline: dict[str, dict[str, Any]]) -> tuple[bool, str | None]:
    """这份文件当初从共享拿的是哪个版本：(知道, 哈希)。共享里没有、也没记录 = 新文件（None）"""
    shared_file = Path(task["shared"]) / rel
    base = baseline.get(_base_key(shared_file))
    if base:
        digest = _task_baseline(base, task["id"]).get("sha256")
        if digest:
            return True, digest
    if not shared_file.exists():
        return True, None
    return False, None


def preview_free(source_text: Any, dest_text: Any, guard: Any = None) -> dict[str, Any]:
    """自由上传的清单（只读不改）：起点里的每个 Excel 传到终点同样的位置，先看终点那份能不能覆盖"""
    src, dst = _clean_root(source_text, "起点文件夹"), _clean_root(dest_text, "终点文件夹")
    if src == dst or _inside(src, dst) or _inside(dst, src):
        raise ValueError("起点和终点不能是同一个文件夹，也不能互相包含")
    found = scan_folder(src)
    if not found["excel"]:
        raise ValueError(f"起点里没有 Excel 文件：{src}")
    manifests = _find_manifests(src)
    baseline = load_baseline(strict=True)
    files, notes = [], []
    mismatch: set[str] = set()
    listed_shared: dict[str, tuple[Path, set[str]]] = {}  # 用上的清单：共享根目录 → 当时共享里有的文件
    for rel in sorted(found["excel"], key=_key):
        info = found["excel"][rel]
        item = {"rel": rel, "size": info["size"], "mtime": info["mtime"]}
        try:
            src_file, dest_file = _task_file(src, rel), _task_file(dst, rel)
            item.update(src_sha256=_digest(src_file), dest_sha256=_shared_digest(dest_file))
        except (OSError, ValueError) as err:
            files.append({**item, "action": "conflict", "guard": "none", "error": _explain(err)})
            continue
        hit = _manifest_for(src_file, manifests)
        record = None
        if hit:
            folder, data, inner = hit
            company = Path(data["shared"]) / inner
            if _base_key(company) == _base_key(dest_file):
                record = next((v for r, v in data["files"].items() if _key(r) == _key(inner)), None)
                listed_shared.setdefault(_base_key(data["shared"]), (Path(data["shared"]), {_key(r) for r in data.get("shared_files", [])}))
            elif str(folder) not in mismatch:
                mismatch.add(str(folder))
                notes.append(f"「{folder / MANIFEST_NAME}」记的共享文件夹是「{data['shared']}」，和这次的终点对不上，"
                             "里面记的版本没法用（请确认终点有没有填错）")
        upload = (baseline.get(_base_key(dest_file), {}).get("upload") or {})
        if upload.get("source") == _base_key(src_file) and (
            record is None or upload.get("source_sha256") == item["src_sha256"] == record.get("sha256")
        ):
            known = upload.get("sha256")
            item["guard"] = "baseline"
        elif record and item["src_sha256"] == record.get("sha256"):
            known = record.get("shared_sha256")
            item["guard"] = "manifest"
        else:
            known = None
            item["guard"] = "none"
        current = item["dest_sha256"]
        reason = _upload_review_reason(known, current, _guard_mode(guard))
        if reason:
            item.update(action="review", review_reason=reason, note=reason)
        else:
            item["action"] = "new" if current is None else "overwrite"
        files.append(item)
    # 终点里有、起点没有的 Excel：不动，只列出来；有清单的话分出「期间别人新加的」
    after = scan_folder(dst)
    src_keys = {_key(rel) for rel in found["excel"]}
    extra, added = [], []
    for rel in sorted(after["excel"], key=_key):
        if _key(rel) in src_keys:
            continue
        full = dst / rel
        owner = next(((root, names) for root, names in listed_shared.values() if _inside(full, root)), None)
        if owner and _key(full.relative_to(owner[0]).as_posix()) not in owner[1]:
            added.append(rel)
        else:
            extra.append(rel)
    summary = {"要上传": sum(f["action"] != "conflict" for f in files), "新建": sum(f["action"] == "new" for f in files),
               "覆盖": sum(f["action"] == "overwrite" for f in files), "要确认": sum(f["action"] == "review" for f in files),
               "别人改过": sum(f["action"] == "conflict" for f in files), "期间新加": len(added), "终点多出": len(extra),
               "有清单": sum(f["guard"] == "manifest" for f in files)}
    result = {"id": uuid.uuid4().hex[:12], "kind": "free_upload", "label": f"{src.name} → {dst.name}", "guard": _guard_mode(guard),
              "source_root": str(src), "dest_root": str(dst), "files": files, "added": added, "extra": extra,
              "notes": notes, "errors": found["errors"] + after["errors"], "summary": summary,
              "manifests": [str(folder / MANIFEST_NAME) for folder, _ in manifests]}
    relay_task = _free_relay_task(src, dst)
    if relay_task:
        result["relay_task"] = relay_task
    debug_log.event("📋 自由上传清单", 起点=str(src), 终点=str(dst), **summary,
                    写拿取版本清单=relay_task["name"] if relay_task else "否")
    with _LOCK:
        if not _STATUS.get("running"):
            _STATUS["preview"] = result
            _STATUS.pop("applied", None)
    return result


def start_free(preview_id: Any, pdf: Any = None) -> None:
    pdf_settings = _pdf_settings(pdf)
    with _LOCK:
        preview_data = _STATUS.get("preview")
        if not preview_data or preview_data.get("kind") != "free_upload" or preview_data.get("id") != preview_id:
            raise ValueError("清单已变化，请重新生成")
        if not preview_data["summary"]["要上传"]:
            raise ValueError("清单里没有能上传的文件")
        upload = {**copy.deepcopy(preview_data), "pdf": pdf_settings}
    start(preview_id, upload=upload)


def _prepare_free(items: list[dict[str, Any]], temp_root: Path, src: Path,
                  pdf_settings: dict[str, Any] | None = None) -> None:
    """在本机临时副本上清标记 + 视图标准化（没有参考，不恢复颜色）；没做成的那份不传，其余照传"""
    from core import upload_prep
    from core.excel_ops import quit_background_excel, start_background_excel

    excel = None
    guard = file_problem.Guard("上传前处理（清标记 · 视图标准化）")

    def get_excel() -> Any:
        nonlocal excel
        if excel is None:
            excel = guard.watch(start_background_excel())
        return excel

    wf.pythoncom.CoInitialize()
    try:
        for index, item in enumerate(items):
            local = temp_root / str(index) / Path(item["rel"]).name
            item["_temp"] = str(local)
            while True:
                _progress(index, len(items), f"上传前处理 {item['rel']}")
                where: dict[str, str] = {}
                try:
                    if guard.file(item["rel"]):  # 进度框能「跳过这一份」
                        excel = None
                    cached = pdf_settings and _use_cached_pdf(item, src / item["rel"], _pdf_target(src / item["rel"]),
                                                              item["src_sha256"], pdf_settings)
                    if cached and pdf_settings["only"]:  # 只传 PDF、起点的 Excel 没变：不处理、不导，直接传起点旁边那份 PDF
                        if _digest(src / item["rel"]) != item["src_sha256"]:
                            raise ValueError("起点的文件在生成清单之后变了，请重新生成清单")
                        break
                    local.parent.mkdir(parents=True, exist_ok=True)
                    if copy_one(src / item["rel"], local) != item["src_sha256"]:
                        raise ValueError("起点的文件在生成清单之后变了，请重新生成清单")
                    item["prep"] = {**upload_prep.prepare(get_excel, local, None, where), "no_reference": True}
                    item["_prepared_sha256"] = _digest(local)
                    wf.log_msg(f"🧽 {item['rel']}：{upload_prep.text(item['prep'])}")
                    if pdf_settings and not cached:  # 从清过标记的临时副本导出，PDF 放在临时副本旁边
                        where["step"] = "导出同名 PDF"
                        _export_pdf(get_excel, item, local, pdf_settings, _sheet_progress(index, len(items), item["rel"], guard))
                    break
                except Exception as err:
                    debug_log.exception(err, f"上传前处理 {item['rel']}")
                    retry, message = guard.problem(err, src / item["rel"], item["rel"], src,
                                                   close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {item['rel']}")
                        continue
                    item.update(status="failed", error=f"上传前{where.get('step') or '处理'}没做成：{message}")
                    break
    finally:
        try:
            if excel is not None:
                quit_background_excel(excel)
        finally:
            guard.stop()
            wf.pythoncom.CoUninitialize()


def _run_free(data: dict[str, Any], reviews: UploadReview) -> dict[str, Any]:
    import tempfile
    src, dst = Path(data["source_root"]), Path(data["dest_root"])
    if not dst.is_dir():
        raise ValueError(f"终点文件夹不可用：{dst}")
    ready: list[Candidate] = []
    for item in data["files"]:
        if item["action"] == "conflict":
            item["status"] = "conflict"
            continue
        try:
            ready.append((item, _task_file(src, item["rel"]), _task_file(dst, item["rel"]), item["dest_sha256"]))
        except ValueError as err:
            item.update(status="failed", error=_explain(err))
    pdf = data.get("pdf")
    pdf_claimed: set[str] = set()
    relay_task = data.get("relay_task")
    relayed: dict[str, dict[str, Any]] = {}
    unknown: list[str] = []
    baseline = load_baseline() if relay_task else {}
    with tempfile.TemporaryDirectory(prefix="自由上传_") as temp:
        for index, batch in enumerate(reviews.batches(ready, _verify_upload_candidate,
                                                      require_review=not (pdf or {}).get("only"))):
            verified = _verified_uploads(batch)
            if not verified:
                continue
            _prepare_free([item for item, _, _, _ in verified], Path(temp) / str(index), src, pdf)
            for candidate in verified:
                item, source, dest, _ = candidate
                if item.get("status"):
                    continue
                _progress(sum(bool(f.get("status")) for f in data["files"]), len(data["files"]), f"上传 {item['rel']}")

                def verify_dest() -> None:
                    if not dst.is_dir():
                        raise _SharedVersionChanged("终点文件夹不可用，请重新检查")
                    _verify_upload_candidate(candidate)

                try:
                    if (pdf or {}).get("only"):
                        verify_dest()
                        _save_baseline(_upload_pdf(item, dest, pdf, pdf_claimed))
                        _pdf_only_status(item)
                        continue
                    digest = copy_one(Path(item["_temp"]), dest, check_source_open=False, before_replace=verify_dest,
                                      expected_digest=item["_prepared_sha256"])
                except (OSError, ValueError) as err:
                    item.update(status="conflict" if isinstance(err, _SharedVersionChanged) else "failed", error=_explain(err))
                    continue
                entry = _base_entry(dest, digest, "复制")
                entry["upload"]["source"] = _base_key(source)
                entry["upload"]["source_sha256"] = item["src_sha256"]
                item.update(status="ok", uploaded_at=entry["upload"]["at"])
                changes = {_base_key(dest): entry}
                if pdf:
                    changes.update(_upload_pdf(item, dest, pdf, pdf_claimed))
                _save_baseline(changes)
                if relay_task:
                    known, shared_sha = _fetched_version(relay_task, item["rel"], baseline)
                    if known:
                        relayed[item["rel"]] = {"shared_sha256": shared_sha, "sha256": digest,
                                                "uploaded_at": entry["upload"]["at"]}
                    else:
                        unknown.append(item["rel"])
            if relayed:
                try:
                    data["manifest"] = str(write_manifest(dst, Path(relay_task["shared"]), relayed, relay_task["name"]))
                    data.pop("manifest_error", None)
                    relayed.clear()
                except (OSError, ValueError) as err:
                    debug_log.exception(err, "写拿取版本清单")
                    data["manifest_error"] = _explain(err)
    if unknown:
        data["manifest_unknown"] = unknown
        debug_log.event("⚠️ 拿取版本清单没记这几份（找不到当初拿的版本）", 文件=unknown)
    for item in data["files"]:
        item.pop("_temp", None)
        item.pop("_prepared_sha256", None)
        item.pop("_pdf_local", None)
    files = data["files"]
    data["summary"] = {**data["summary"], "已复制": sum(f.get("status") == "ok" for f in files),
                       "别人改过没覆盖": sum(f.get("status") == "conflict" for f in files),
                       "已跳过": sum(f.get("status") == "skipped" for f in files),
                       "没复制成功": sum(f.get("status") == "failed" for f in files), **_pdf_summary(files)}
    return {**data, "id": uuid.uuid4().hex[:12], "finished": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}


def markdown_free(result: dict[str, Any]) -> str:
    s = result.get("summary") or {}
    lines = [f"# 上传记录：{result['label']}", "", f"- 起点：`{result['source_root']}`", f"- 终点：`{result['dest_root']}`",
             "- 每个文件先在本机临时副本上清掉标记色、视图标准化（A1、100%），再复制到终点；起点的文件没有改动。", ""]
    if result.get("stopped"):
        lines += [f"> ⛔ {result['stopped']}", ""]
    for note in result.get("notes", []):
        lines.append(f"> ⚠️ {_md(note)}")
    lines += ["", f"已上传 {s.get('已复制', 0)} / {len(result.get('files', []))} 个；"
              f"别人改过没覆盖 {s.get('别人改过没覆盖', 0)} 个；没上传成功 {s.get('没复制成功', 0)} 个。", "",
              "| 文件 | 结果 |", "| --- | --- |"]
    from core import upload_prep
    for item in result.get("files", []):
        status = {"ok": "已上传", "conflict": "🛑 " + (item.get("error") or "别人改过，没有覆盖")}.get(
            item.get("status"), item.get("error") or "未上传")
        if item.get("pdf_only"):
            status = "只传 PDF，Excel 没传" if item.get("status") == "ok" else status
        elif item.get("status") == "ok":
            status += ("（看过差异后确认覆盖）" if item.get("review_confirmed") and item.get("dest_sha256") else
                       "（强制覆盖）" if item["action"] == "force" else
                       "（覆盖）" if item["action"] in ("overwrite", "confirm") else "（新建）")
        if item.get("prep"):
            status += f"（上传前：{upload_prep.text(item['prep'])}）"
        if item.get("pdf"):
            status += f"；{_pdf_text(item['pdf'])}"
        lines.append(f"| {_md(item['rel'])} | {_md(status)} |")
    if result.get("manifest"):
        lines += ["", f"- 拿取版本清单：`{result['manifest']}`（从这里传回共享时用来检查期间有没有人改过，请不要删除）"]
    if result.get("manifest_error"):
        lines += ["", f"- ⚠️ 拿取版本清单没写成：{_md(result['manifest_error'])}"]
    if result.get("manifest_unknown"):
        lines += ["", "- ⚠️ 这几份找不到当初从共享拿的版本，拿取版本清单里没记：" + "、".join(_md(r) for r in result["manifest_unknown"])]
    if result.get("added"):
        lines += ["", "## 期间别人新加到终点的 Excel（没有动）", ""] + [f"- {_md(rel)}" for rel in result["added"]]
    if result.get("extra"):
        lines += ["", "## 终点里有、起点没有的 Excel（没有动）", ""] + [f"- {_md(rel)}" for rel in result["extra"]]
    return "\n".join(lines) + "\n"


# ───────────────────────── 记录 ─────────────────────────

def _load_history() -> list[dict[str, Any]]:
    try:
        return list((storage.load_snapshot(HISTORY_NAME) or {}).get("runs", []))
    except Exception as err:
        debug_log.event("⚠️ 没能读出复制记录", 原因=str(err))
        return []


def _save_history(result: dict[str, Any]) -> None:
    runs = [result] + [run for run in _load_history() if run.get("id") != result["id"]]
    storage.save_snapshot(HISTORY_NAME, {"runs": runs[:MAX_HISTORY]})


def history() -> list[dict[str, Any]]:
    """网页上的历史列表（只给摘要）"""
    return [{k: run.get(k) for k in ("id", "kind", "label", "source_root", "dest_root", "finished", "summary", "stopped")}
            | {"attention": needs_attention(run) if run.get("summary") else True} for run in _load_history()]


def get_run(run_id: str) -> dict[str, Any]:
    with _LOCK:
        applied = _STATUS.get("applied")
        if applied and applied.get("id") == run_id:
            return copy.deepcopy(applied)
    for run in _load_history():
        if run.get("id") == run_id:
            return run
    raise ValueError("找不到这次复制记录")


def _size(value: Any) -> str:
    if value is None:
        return "-"
    value = float(value)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024
    return "-"


def _md(text: Any) -> str:
    return str(text).replace("|", "\\|").replace("\n", " ")


def markdown(result: dict[str, Any]) -> str:
    """复制记录文档：不写日期和时间"""
    if result.get("kind") == "fetch":
        return markdown_fetch(result)
    if result.get("kind") == "free_upload":
        return markdown_free(result)
    if result.get("kind") == "task_upload":
        lines = [f"# 批量上传记录：{result['label']}", "", f"- 处理文件夹：`{result['source_root']}`"]
        if result.get("relay"):
            lines += [f"- 上传到（中转）：`{result['dest_root']}`", f"- 公司共享文件夹（只核对版本，没有改动）：`{result.get('shared', '')}`"]
            if result.get("manifest"):
                lines.append(f"- 拿取版本清单：`{result['manifest']}`（从中转上传到公司时用来检查期间有没有人改过，请不要删除）")
            if result.get("manifest_error"):
                lines.append(f"- ⚠️ 拿取版本清单没写成：{_md(result['manifest_error'])}")
        else:
            lines.append(f"- 共享文件夹：`{result['dest_root']}`")
        lines += ["", "| 文件 | 结果 |", "| --- | --- |"]
        for item in result["files"]:
            status = (("只传 PDF，Excel 没传" if item.get("pdf_only") else "已上传") if item.get("status") == "ok"
                      else item.get("error") or "未上传")
            if item.get("backup_path"):
                status += f"（旧处理副本备份在 `{item['backup_path']}`）"
            if item.get("ref_backup_path"):
                status += f"（旧参考备份在 `{item['ref_backup_path']}`）"
            if item.get("prep"):
                from core import upload_prep
                status += f"（{'同步前那一版做过的' if item.get('prep_stale') else ''}上传前：{upload_prep.text(item['prep'])}）"
            if item.get("pdf"):
                status += f"；{_pdf_text(item['pdf'])}"
            lines.append(f"| {_md(item['rel'])} | {_md(status)} |")
        return "\n".join(lines) + "\n"
    s = result.get("summary") or {}
    lines = [f"# 复制记录：{result.get('label') or '按 JSON 复制 Excel'}", ""]
    lines += [f"- 来源根目录：`{result.get('source_root', '')}`", f"- 目的地根目录：`{result.get('dest_root', '')}`",
              "- 只复制 Excel 表格文件（包括子文件夹），目的地已有同名文件时覆盖，文件名不改。",
              "- 覆盖前先检查目的地的文件在拿走之后有没有被别人改过，改过的不覆盖。", ""]
    if result.get("stopped"):
        lines += [f"> ⛔ {result['stopped']}", ""]
        lines += [f"- {_md(x)}" for x in result.get("stale", [])[:200]] + [""]
    if s:
        lines += ["## 合计", "", "| 项目 | 数量 |", "| --- | --- |"]
        lines += [f"| {k} | {v} |" for k, v in s.items()] + [""]
    attention = [f for f in result.get("folders", []) if f.get("error") or any(not c["ok"] for c in f.get("checks", []))
                 or any(x.get("status") != "ok" for x in f["files"]) or f.get("dir_errors")]
    conflicts = [(f, x) for f in result.get("folders", []) for x in f["files"] if x.get("status") == "conflict"]
    if conflicts:
        lines += ["## 🛑 别人改过、没有覆盖的文件", "",
                  "这些文件在你拿走之后被别人改过。请先拿最新版，在新版上改好再复制。", ""]
        lines += [f"- `{_md(f['dest_rel'])}/{_md(x['dest_rel'])}`：{_md(x.get('error', ''))}" for f, x in conflicts]
        lines.append("")
    if attention:
        lines += ["## ⚠️ 需要查看的文件夹", ""]
        for f in attention:
            why = [f["error"]] if f.get("error") else []
            why += [c["text"] for c in f.get("checks", []) if not c["ok"]]
            conflict = sum(x.get("status") == "conflict" for x in f["files"])
            failed = sum(x.get("status") != "ok" for x in f["files"]) - conflict
            if conflict:
                why.append(f"{conflict} 个文件别人改过、没有覆盖")
            if failed:
                why.append(f"{failed} 个文件没复制成功")
            if f.get("dir_errors"):
                why.append(f"{len(f['dir_errors'])} 个文件夹没建成")
            lines.append(f"- `{f['source_rel']}` → `{f['dest_rel']}`：{'；'.join(why)}")
        lines.append("")
    if result.get("notes"):
        lines += ["## JSON 本身的合计数", ""] + [f"- {n}" for n in result["notes"]] + [""]
    person = None
    for f in result.get("folders", []):
        if f["person"] != person:
            person = f["person"]
            lines += [f"## {person}", ""]
        lines += [f"### {f['source_rel']} → {f['dest_rel']}", ""]
        meta = [f"上传人：{f['person']}"]
        if f.get("owner"):
            meta.append(f"作者：{f['owner']}")
        lines += ["- " + "，".join(meta), f"- 来源：`{f['source_path']}`", f"- 目的地：`{f['dest_path']}`"]
        if f.get("error"):
            lines += [f"- ⛔ {f['error']}，这个文件夹没有复制" + ("，目的地已建成空文件夹" if "." in f.get("created_dirs", []) else "")]
        created = [d for d in f.get("created_dirs", []) if d != "."]
        if "." in f.get("created_dirs", []) and not f.get("error"):
            lines.append("- 📁 目的地文件夹是这次新建的")
        if created:
            lines.append(f"- 📁 新建子文件夹 {len(created)} 个：" + "、".join(f"`{_md(d)}`" for d in created))
        for err in f.get("dir_errors", []):
            lines.append(f"- ❌ 文件夹没建成：{_md(err)}")
        for text in f.get("overlaps", []):
            lines.append(f"- ⚠️ {text}")
        lines += ["", "| JSON 写的数 | 来源找到的 Excel | 复制后目的地的 Excel |", "| --- | --- | --- |",
                  f"| {f['json_files'] if f['json_files'] is not None else '（没写）'} | "
                  f"{f['source_count'] if f['source_count'] is not None else '-'} | "
                  f"{f['dest_count'] if f.get('dest_count') is not None else '-'} |", ""]
        for c in f.get("checks", []):
            lines.append(f"- {'✅' if c['ok'] else '❌'} {c['pair']}：" + ("一致" if c["ok"] else c["text"]))
        if any(not c["ok"] and c["pair"] == "JSON ↔ 来源" for c in f.get("checks", [])) and f.get("hint"):
            lines.append(f"  - {f['hint']}")
        if f.get("temp") or f.get("other"):
            extra = []
            if f.get("temp"):
                extra.append(f"~$ 临时文件 {len(f['temp'])} 个")
            if f.get("other"):
                extra.append("非 Excel 文件 " + "、".join(f"{ext} {n} 个" for ext, n in sorted(f["other"].items())))
            lines.append(f"- 来源里另有（没有复制、没算进来源数）：{'，'.join(extra)}")
        if f.get("missing_in_dest"):
            lines += ["", "**来源有、目的地没有（来源多）：**", ""] + [f"- {_md(x)}" for x in f["missing_in_dest"]]
        if f.get("dest_extra"):
            lines += ["", "**目的地有、来源没有（目的地多，原来就有，没有动）：**", ""] + [f"- {_md(x)}" for x in f["dest_extra"]]
        for err in f.get("scan_errors", []):
            lines.append(f"- ⚠️ 读不到：{_md(err)}")
        if f["files"]:
            lines += ["", "| 文件 | 操作 | 大小 | 结果 |", "| --- | --- | --- | --- |"]
            for x in f["files"]:
                if x["action"] == "overwrite":
                    action, size = "🔁 覆盖", f"{_size(x['old_size'])} → {_size(x['size'])}"
                else:
                    action, size = "🆕 新建", _size(x["size"])
                state = {"ok": "✅ 已复制", "failed": f"❌ {x.get('error', '')}",
                         "conflict": "🛑 别人改过，没有覆盖"}.get(x.get("status"), "没有复制")
                lines.append(f"| {_md(x['rel'])} | {action} | {size} | {_md(state)} |")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


# ───────────────────────── 拿最新版（目的地 → 来源，记下版本） ─────────────────────────
# 反过来：把目的地（共享文件夹）里的 Excel 拿回来源（本地），同时记下拿到的版本，覆盖前用来检查。
# 本地已有同名文件时：
# - 内容一样：不复制，只记下版本；
# - 不一样、但目的地还是上次记下的版本（别人没动，是你改过）：不拿，保留你改的；
# - 不一样、目的地被改过（或没有记录）：用目的地的替换本地的，你本地那份先备份到来源根目录旁边的备份文件夹。

def _fetch_decide(shared_file: Path, local_file: Path, local_exists: bool, base: dict[str, Any] | None) -> dict[str, Any]:
    if not local_exists:
        return {"action": "copy", "backup": False}
    try:
        shared_sha = _digest(shared_file)
    except OSError:
        shared_sha = None
    try:
        local_sha = _digest(local_file)
    except OSError:
        local_sha = None
    if shared_sha and shared_sha == local_sha:
        return {"action": "same", "backup": False, "sha": shared_sha}
    base_sha = (base or {}).get("sha256")
    if base_sha and shared_sha == base_sha:
        return {"action": "keep", "backup": False}
    # 本地还是上次拿的那份（你没改过）就不用备份
    return {"action": "replace", "backup": not (base_sha and local_sha == base_sha)}


def _fetch_folder(item: dict[str, Any], src_root: Path, dst_root: Path,
                  baseline: dict[str, dict[str, Any]], task_id: str = "", picks: list[str] | None = None) -> dict[str, Any]:
    local = src_root.joinpath(*item["source"])
    shared = dst_root.joinpath(*item["destination"])
    entry: dict[str, Any] = {
        **item,
        "source_rel": "/".join(item["source"]),
        "dest_rel": "/".join(item["destination"]),
        "source_path": str(local),
        "dest_path": str(shared),
        "error": "",
        "files": [],
        "local_extra": [],
        "scan_errors": [],
    }
    if not shared.is_dir():
        entry["error"] = "目的地找不到这个文件夹，没有可拿的" if not shared.exists() else "目的地路径不是文件夹"
        debug_log.event("⚠️ 拿取清单：" + entry["error"], 文件夹=str(shared))
        return entry
    if local.exists() and (_inside(local.resolve(), shared.resolve()) or _inside(shared.resolve(), local.resolve())):
        entry["error"] = "来源文件夹和目的地文件夹互相包含，不能拿"
        return entry
    found = scan_folder(shared)
    have = scan_folder(local) if local.is_dir() else {"excel": {}, "errors": []}
    if picks:  # 新任务填了多个路径：只拿选中的文件夹 / 文件（用户 2026-11-12）
        from core.work_task import in_picks
        found["excel"] = {rel: info for rel, info in found["excel"].items() if in_picks(rel, picks)}
        have["excel"] = {rel: info for rel, info in have["excel"].items() if in_picks(rel, picks)}
        entry["picks"] = list(picks)
        entry["picks_missing"] = [pick for pick in picks if not (shared / pick).exists()]
        for pick in entry["picks_missing"]:
            found["errors"].append(f"{pick}：共享里现在找不到（可能被删了或改了名）")
    have_keys = {_key(rel): info for rel, info in have["excel"].items()}
    shared_keys = {_key(rel) for rel in found["excel"]}
    for rel, info in found["excel"].items():
        old = have_keys.get(_key(rel))
        local_rel = old["rel"] if old else rel
        base = baseline.get(_base_key(shared / rel), {})
        decision = _fetch_decide(shared / rel, local / local_rel, bool(old), _task_baseline(base, task_id))
        entry["files"].append({**info, "local_rel": local_rel, "shared_open": source_open(shared / rel), **decision})
    entry["local_extra"] = sorted(info["rel"] for key, info in have_keys.items() if key not in shared_keys)
    entry["scan_errors"] = found["errors"] + have.get("errors", [])
    counts = {a: sum(f["action"] == a for f in entry["files"]) for a in ("copy", "replace", "keep", "same")}
    debug_log.event("📋 拿取清单文件夹", 目的地=entry["dest_rel"], 来源=entry["source_rel"], 新拿=counts["copy"],
                    替换=counts["replace"], 要备份=sum(f["backup"] for f in entry["files"]),
                    保留你改的=counts["keep"], 已一样=counts["same"], 本地多=len(entry["local_extra"]))
    return entry


def _fetch_summary(folders: list[dict[str, Any]], applied: bool = False) -> dict[str, int]:
    files = [f for folder in folders for f in folder["files"]]
    if applied:
        done = [f for f in files if f.get("status") == "ok"]
        return {
            "文件夹": len(folders),
            "已拿": sum(f["action"] in ("copy", "replace") for f in done),
            "新拿": sum(f["action"] == "copy" for f in done),
            "替换": sum(f["action"] == "replace" for f in done),
            "备份": sum(bool(f.get("backup_path")) for f in done),
            "备份参考": sum(bool(f.get("ref_backup_path")) for f in done),
            "保留你改的": sum(f["action"] == "keep" for f in done),
            "已一样": sum(f["action"] == "same" for f in done),
            "记下版本": sum(bool(f.get("recorded")) for f in done),
            "没拿成功": sum(f.get("status") == "failed" for f in files),
            "没有拿": sum("status" not in f for f in files),
            "有问题的文件夹": sum(bool(folder["error"]) for folder in folders),
        }
    return {
        "文件夹": len(folders),
        "要拿": sum(f["action"] in ("copy", "replace") for f in files),
        "新拿": sum(f["action"] == "copy" for f in files),
        "替换": sum(f["action"] == "replace" for f in files),
        "要备份": sum(bool(f["backup"]) for f in files),
        "保留你改的": sum(f["action"] == "keep" for f in files),
        "已一样": sum(f["action"] == "same" for f in files),
        "有问题": sum(bool(folder["error"] or folder["scan_errors"]) for folder in folders),
    }


def preview_fetch(plan_text: str, source_root: str, dest_root: str, task_id: Any = None) -> dict[str, Any]:
    """生成拿取清单：目的地哪些要拿回来源、哪些替换（先备份）、哪些保留你改的。不改任何文件"""
    plan_text, source_root, dest_root, task = _task_args(task_id, plan_text, source_root, dest_root)
    plan = parse_plan(plan_text)
    src_root = _clean_root(source_root, "来源根目录")
    dst_root = _clean_root(dest_root, "目的地根目录")
    if src_root == dst_root:
        raise ValueError("来源根目录和目的地根目录是同一个文件夹，请重新选择")
    with debug_log.side_task("distribute_fetch_preview", {"source": str(src_root), "dest": str(dst_root),
                                                          "folders": len(plan["folders"]), "label": plan["label"]}):
        baseline = load_baseline()
        folders = [_fetch_folder(item, src_root, dst_root, baseline, (task or {}).get("id", ""), (task or {}).get("picks"))
                   for item in plan["folders"]]
        result = {
            "id": uuid.uuid4().hex[:12],
            "kind": "fetch",
            "task": task,
            "label": plan["label"],
            "source_root": str(src_root),
            "dest_root": str(dst_root),
            "notes": [],
            "folders": folders,
            "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "summary": _fetch_summary(folders),
        }
        debug_log.event("📋 拿取清单", **result["summary"])
    with _LOCK:
        if not _STATUS.get("running"):
            _STATUS["preview"] = result
            _STATUS.pop("applied", None)
    return result


def _backup_root(src_root: Path) -> Path:
    """本地被替换的文件先备份到来源根目录旁边（不放在来源里面，免得被当成要复制的 Excel）"""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if src_root.name and src_root.parent != src_root:
        return src_root.parent / f"{src_root.name}_拿最新版前的备份" / stamp
    return storage.DATABASE_PATH.parent / "拿最新版前的备份" / stamp


def _refresh_reference(fresh_local: Path, ref_file: Path, backup: Path | None = None) -> str:
    """任务的参考文件夹：放一份和共享一样的原件（刚拿下来、和共享一致的本地文件），已经一样就不复制。
    参考里原来那份不一样时，先备份到 backup（用户 2026-11-06 定：处理和参考覆盖前都要备份），返回备份路径。"""
    try:
        if ref_file.exists() and _digest(ref_file) == _digest(fresh_local):
            return ""
    except OSError:
        pass
    saved = ""
    if backup is not None and ref_file.is_file():
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ref_file, backup)
        saved = str(backup)
        debug_log.event("💾 备份参考原件", 文件=str(ref_file), 备份到=saved)
    copy_one(fresh_local, ref_file, check_source_open=False)
    debug_log.event("📚 更新参考原件", 文件=str(ref_file))
    return saved


def _run_fetch(preview_data: dict[str, Any]) -> dict[str, Any]:
    folders = copy.deepcopy(preview_data["folders"])
    total = sum(len(f["files"]) for f in folders)
    done = 0
    baseline = load_baseline()
    task = preview_data.get("task") or {}
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_root = (Path(task["fetch_backup"]) / stamp) if task.get("fetch_backup") \
        else _backup_root(Path(preview_data["source_root"]))
    reference = Path(task["reference"]) if task.get("reference") else None
    for folder in folders:
        if folder["error"]:
            wf.log_msg(f"⚠️ {folder['dest_rel']}：{folder['error']}，跳过")
            continue
        local_dir, shared_dir = Path(folder["source_path"]), Path(folder["dest_path"])
        debug_log.begin(f"拿最新版 {folder['dest_rel']} → {folder['source_rel']}")
        changes: dict[str, dict[str, Any]] = {}
        try:
            for item in folder["files"]:
                label = f"{folder['dest_rel']}/{item['rel']}"
                _progress(done, total, label)
                shared_file, local_file = shared_dir / item["rel"], local_dir / item["local_rel"]
                key = _base_key(shared_file)
                planned = item["action"]
                # 生成清单之后可能又有变化：照现在的情况再判断一次
                base = baseline.get(key, {})
                item.update(_fetch_decide(shared_file, local_file, local_file.exists(), _task_baseline(base, task.get("id", ""))))
                if item["action"] != planned:
                    debug_log.event("ℹ️ 和清单时不一样，按现在的情况处理", 文件=label, 清单=planned, 现在=item["action"])
                while True:
                    try:
                        if item["action"] == "same":
                            digest = item["sha"]
                        elif item["action"] in ("copy", "replace"):
                            if item["backup"] and local_file.exists():
                                backup = (backup_root / "处理" / item["local_rel"]) if reference \
                                    else backup_root / (item["local_rel"] if task else f"{folder['source_rel']}/{item['local_rel']}")
                                backup.parent.mkdir(parents=True, exist_ok=True)
                                shutil.copy2(local_file, backup)
                                item["backup_path"] = str(backup)
                                debug_log.event("💾 备份本地这份", 文件=str(local_file), 备份到=str(backup))
                            digest = copy_one(shared_file, local_file, check_source_open=False)
                        if item["action"] != "keep":
                            if reference:
                                ref_backup = _refresh_reference(local_file, reference / item["rel"], backup_root / "参考" / item["rel"])
                                if ref_backup:
                                    item["ref_backup_path"] = ref_backup
                                item["reference"] = True
                            how = "拿取（本地已一样）" if item["action"] == "same" else "拿取"
                            changes[key] = _base_entry(shared_file, digest, how, task.get("id", ""))
                            item["recorded"] = True
                        item.update(status="ok", error="")
                        break
                    except Exception as err:
                        message = _explain(err)
                        debug_log.exception(err, f"拿最新版 {label}")
                        bad = getattr(err, "filename", None) or str(shared_file)
                        if file_problem.ask(bad, label, message, folder=Path(bad).parent) == "retry":
                            wf.log_msg(f"🔁 重试 {label}")
                            continue
                        item.update(status="failed", error=message)
                        break
                done += 1
                debug_log.event("📄 拿最新版" if item["status"] == "ok" else "❌ 没拿成功", 文件=label,
                                做法=item["action"], 备份=item.get("backup_path", "-"), 原因=item["error"] or "-")
        finally:
            try:
                _save_baseline(changes)
            except Exception as err:
                debug_log.exception(err, "保存版本记录")
            debug_log.end()
    result = {
        "id": uuid.uuid4().hex[:12],
        "kind": "fetch",
        "task": preview_data.get("task"),
        "label": preview_data["label"],
        "source_root": preview_data["source_root"],
        "dest_root": preview_data["dest_root"],
        "backup_root": str(backup_root) if backup_root.exists() else "",
        "notes": [],
        "folders": folders,
        "stopped": "",
        "stale": [],
        "started": preview_data.get("_started", ""),
        "finished": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    result["summary"] = _fetch_summary(folders, applied=True)
    return result


def markdown_fetch(result: dict[str, Any]) -> str:
    """拿取记录文档：不写日期和时间"""
    s = result.get("summary") or {}
    lines = [f"# 拿最新版记录：{result.get('label') or '按 JSON 拿最新版'}", ""]
    lines += [f"- 从目的地根目录：`{result.get('dest_root', '')}`", f"- 拿回来源根目录：`{result.get('source_root', '')}`",
              "- 只拿 Excel 表格文件（包括子文件夹），拿的时候记下版本，复制（覆盖）前用来检查别人改过没有。"]
    if result.get("backup_root"):
        lines.append(f"- 本地被替换的文件备份在：`{result['backup_root']}`"
                     + ("（处理文件夹的在「处理」、参考文件夹的在「参考」子文件夹）" if (result.get("task") or {}).get("reference") else ""))
    lines.append("")
    if s:
        lines += ["## 合计", "", "| 项目 | 数量 |", "| --- | --- |"] + [f"| {k} | {v} |" for k, v in s.items()] + [""]
    names = {"copy": "🆕 新拿", "replace": "🔁 替换本地", "keep": "✋ 保留你改的", "same": "＝ 已一样"}
    for f in result.get("folders", []):
        lines += [f"## {f['dest_rel']} → {f['source_rel']}", ""]
        if f.get("error"):
            lines += [f"- ⛔ {f['error']}", ""]
            continue
        if f.get("local_extra"):
            lines += ["本地有、目的地没有（没有动）：", ""] + [f"- {_md(x)}" for x in f["local_extra"]] + [""]
        for err in f.get("scan_errors", []):
            lines.append(f"- ⚠️ 读不到：{_md(err)}")
        if f["files"]:
            lines += ["| 文件 | 做法 | 结果 |", "| --- | --- | --- |"]
            for x in f["files"]:
                state = {"ok": "✅", "failed": f"❌ {x.get('error', '')}"}.get(x.get("status"), "没有拿")
                if x.get("backup_path"):
                    state += f"（你的那份备份在 `{x['backup_path']}`）"
                if x.get("ref_backup_path"):
                    state += f"（旧参考备份在 `{x['ref_backup_path']}`）"
                lines.append(f"| {_md(x['rel'])} | {names.get(x['action'], x['action'])} | {_md(state)} |")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"
