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
复制记录存进 execution.sqlite3（更新程序时会把它复制过去，记录不丢），网页上能看、能下载 Markdown。
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
from typing import Any

from core import debug_log, file_problem, storage, workflow as wf
from core.common import _held_open, workbook_locked

# Microsoft Excel 的表格文件（加载项 .xla/.xlam、CSV 不算）
EXCEL_SUFFIXES = (".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm")
HISTORY_NAME = "distribute_history"
MAX_HISTORY = 200
TEMP_SUFFIX = ".复制中~"

_LOCK = threading.RLock()
_STATUS: dict[str, Any] = {"running": False}


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


def preview(plan_text: str, source_root: str, dest_root: str) -> dict[str, Any]:
    """生成复制清单：每个文件夹要复制哪些、哪些会覆盖、三个数对不对得上。不改任何文件"""
    plan = parse_plan(plan_text)
    src_root = _clean_root(source_root, "来源根目录")
    dst_root = _clean_root(dest_root, "目的地根目录")
    if src_root == dst_root:
        raise ValueError("来源根目录和目的地根目录是同一个文件夹，请重新选择")
    with debug_log.side_task("distribute_preview", {"source": str(src_root), "dest": str(dst_root),
                                                    "folders": len(plan["folders"]), "label": plan["label"]}):
        folders = []
        for item in plan["folders"]:
            folders.append(_preview_folder(item, src_root, dst_root))
        _cross_checks(folders)
        result = {
            "id": uuid.uuid4().hex[:12],
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


def _preview_folder(item: dict[str, Any], src_root: Path, dst_root: Path) -> dict[str, Any]:
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
        entry["files"].append({
            **info,
            "action": "overwrite" if old else "new",
            "old_size": old["size"] if old else None,
            "old_mtime": old["mtime"] if old else None,
            "dest_rel": old["rel"] if old else rel,
            "source_open": locked,
        })
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
    return {
        "文件夹": len(folders),
        "要复制": len(files),
        "覆盖": sum(f["action"] == "overwrite" for f in files),
        "新建": sum(f["action"] == "new" for f in files),
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


def copy_one(src: Path, dst: Path) -> None:
    """先复制成临时文件再替换：网络断开或中途出错时，目的地原来的文件不会只剩半个"""
    if source_open(src):
        raise PermissionError(13, "来源文件正在 Excel 中打开（旁边有 ~$ 锁文件），可能有没保存的修改", str(src))
    if dst.exists() and workbook_locked(dst):
        raise PermissionError(13, "目的地的同名文件正在 Excel 中打开，没法覆盖", str(dst))
    dst.parent.mkdir(parents=True, exist_ok=True)
    temp = dst.with_name(dst.name + TEMP_SUFFIX)
    try:
        shutil.copy2(src, temp)
        if temp.stat().st_size != src.stat().st_size or _digest(temp) != _digest(src):
            raise OSError(5, "复制后内容和来源不一致", str(dst))
        os.replace(temp, dst)
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


def _run(preview_data: dict[str, Any]) -> dict[str, Any]:
    folders = copy.deepcopy(preview_data["folders"])
    total = sum(len(f["files"]) for f in folders)
    done = 0
    stop_reason = ""
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
        try:
            for item in folder["files"]:
                label = f"{folder['source_rel']}/{item['rel']}"
                _progress(done, total, label)
                while True:
                    try:
                        copy_one(src_dir / item["rel"], dst_dir / item["dest_rel"])
                        item.update(status="ok", error="")
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
            wf.log_msg(f"📁 {folder['source_rel']} → {folder['dest_rel']}：复制 {ok_count}/{len(folder['files'])}"
                       + (f"；数量不一致：{'；'.join(bad_checks)}" if bad_checks else "；三个数一致"))
            debug_log.event("🔢 核对数量", 文件夹=folder["dest_rel"], JSON=folder["json_files"],
                            来源=folder["source_count"], 目的地=folder["dest_count"],
                            不一致=bad_checks or "无", 目的地缺=folder["missing_in_dest"][:20],
                            目的地多=folder["dest_extra"][:20])
        finally:
            debug_log.end()
    result = {
        "id": uuid.uuid4().hex[:12],
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
        "没复制成功": sum(f.get("status") == "failed" for f in files),
        "没有复制": sum("status" not in f for f in files),
        "数量不一致": sum(any(not c["ok"] for c in folder["checks"]) for folder in result["folders"]),
        "有问题的文件夹": sum(bool(folder["error"] or folder.get("dir_errors")) for folder in result["folders"]),
    }


def needs_attention(result: dict[str, Any]) -> bool:
    s = result["summary"]
    return bool(result.get("stopped") or s["没复制成功"] or s["没有复制"] or s["数量不一致"]
                or s["有问题的文件夹"] or result.get("notes"))


def start() -> None:
    """按最后一次确认的清单开始复制"""
    with wf.STATE_LOCK, _LOCK:
        preview_data = copy.deepcopy(_STATUS.get("preview"))
        if not preview_data:
            raise ValueError("请先生成复制清单")
        from core import bl_links, hidden_repair, link_repair, mark_finder
        if (wf.STATE.get("busy") or _STATUS.get("running") or link_repair.status().get("running")
                or mark_finder.status().get("running") or hidden_repair.status().get("running")
                or bl_links.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束（别的任务可能正在改这些文件）")
        if not any(f["files"] or f.get("dirs_to_create") for f in preview_data["folders"]):
            raise ValueError("清单里没有要复制的 Excel 文件，也没有要新建的文件夹")
        wf.STATE["busy"] = True
        wf.STATE["maintenance"] = "distribute"
        preview_data["_started"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        _STATUS.update(running=True, done=0, total=sum(len(f["files"]) for f in preview_data["folders"]),
                       current="", label=preview_data["label"])
        _STATUS.pop("applied", None)

    def run() -> None:
        from core import operation_history
        result: dict[str, Any] | None = None
        ids: list[int] = []
        try:
            with debug_log.side_task("distribute", {"source": preview_data["source_root"],
                                                    "dest": preview_data["dest_root"], "label": preview_data["label"],
                                                    "files": _STATUS.get("total")}):
                ids = operation_history.begin(preview_data["dest_root"], ["distribute"])
                with file_problem.task("按 JSON 复制 Excel"):
                    result = _run(preview_data)
                _save_history(result)
                debug_log.event("✅ 按 JSON 复制结束", **result["summary"])
            operation_history.finish(ids, "failed" if needs_attention(result) else "completed")
        except Exception as err:
            debug_log.exception(err, "按 JSON 复制")
            operation_history.finish(ids, "failed")
            result = result or {"id": "", "label": preview_data["label"], "folders": [], "notes": [],
                                "source_root": preview_data["source_root"], "dest_root": preview_data["dest_root"],
                                "stopped": f"程序出错：{err}（详细信息已写进调试日志）", "stale": [],
                                "summary": {}}
        finally:
            with _LOCK:
                _STATUS.update(running=False, current="")
                if result is not None:
                    _STATUS["applied"] = result
                    _STATUS.pop("preview", None)
            with wf.STATE_LOCK:
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)

    try:
        threading.Thread(target=run, name="按JSON复制", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
        raise


def status() -> dict[str, Any]:
    with _LOCK:
        data = copy.deepcopy(_STATUS)
    if data.get("preview"):
        data["preview"] = _public(data["preview"])
    return data


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
    return [{k: run.get(k) for k in ("id", "label", "source_root", "dest_root", "finished", "summary", "stopped")}
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
    s = result.get("summary") or {}
    lines = [f"# 复制记录：{result.get('label') or '按 JSON 复制 Excel'}", ""]
    lines += [f"- 来源根目录：`{result.get('source_root', '')}`", f"- 目的地根目录：`{result.get('dest_root', '')}`",
              "- 只复制 Excel 表格文件（包括子文件夹），目的地已有同名文件时覆盖，文件名不改。", ""]
    if result.get("stopped"):
        lines += [f"> ⛔ {result['stopped']}", ""]
        lines += [f"- {_md(x)}" for x in result.get("stale", [])[:200]] + [""]
    if s:
        lines += ["## 合计", "", "| 项目 | 数量 |", "| --- | --- |"]
        lines += [f"| {k} | {v} |" for k, v in s.items()] + [""]
    attention = [f for f in result.get("folders", []) if f.get("error") or any(not c["ok"] for c in f.get("checks", []))
                 or any(x.get("status") != "ok" for x in f["files"]) or f.get("dir_errors")]
    if attention:
        lines += ["## ⚠️ 需要查看的文件夹", ""]
        for f in attention:
            why = [f["error"]] if f.get("error") else []
            why += [c["text"] for c in f.get("checks", []) if not c["ok"]]
            failed = sum(x.get("status") != "ok" for x in f["files"])
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
                state = {"ok": "✅ 已复制", "failed": f"❌ {x.get('error', '')}"}.get(x.get("status"), "没有复制")
                lines.append(f"| {_md(x['rel'])} | {action} | {size} | {_md(state)} |")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"
