"""Ticket 目录结构模板库：用户定义多套目录树，出票时按模板建目录、从绑定的来源路径复制文件。

- 模板放在数据文件夹里（桌面\\handle-excel\\程序数据\\目录结构，见 core/data_dir.py），换 ZIP 版本不丢：
    目录结构\\index.json    模板名字与目录树（文件节点带来源路径）
- 树节点：dir 有 children；file 有 source（本机或共享盘上的来源文件，出票时复制过来）。
- 出票前可以排除个别节点、给个别目录/文件改名（overrides），不用改模板本身。
- 来源文件缺失、复制失败只记进问题清单，不影响其余目录文件，也不影响 Ticket.txt 落盘。
- 目标里已存在的文件不覆盖：重复出票、手动放过的文件都原样保留。
"""
from __future__ import annotations

import json
import re
import shutil
import threading
import time
import uuid
from pathlib import Path

from core import data_dir
from core.common import console_print, sanitize_filename

MAX_NAME = 60
MAX_TEMPLATES = 50
MAX_NODES = 500
MAX_DEPTH = 8
MAX_SOURCE = 500

LIB_DIR: Path | None = None   # 测试时改这里；平时用数据文件夹下的「目录结构」
_lock = threading.RLock()


def library_dir() -> Path:
    return LIB_DIR if LIB_DIR is not None else data_dir.DATA_DIR / "目录结构"


def _index_path() -> Path:
    return library_dir() / "index.json"


def _log(message: str, **fields) -> None:
    detail = " ".join(f"{k}={v}" for k, v in fields.items())
    console_print(f"🗂️ [目录结构] {message} {detail}".rstrip())
    try:
        from core import debug_log
        debug_log.event(f"🗂️ 目录结构：{message}", **fields)
    except Exception:  # noqa: BLE001 — 调试日志写不了也不影响功能
        pass


# ---------------------------------------------------------------- 索引

def _load() -> dict:
    path = _index_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        data = {}
    except (OSError, ValueError) as err:
        backup = path.with_name(f"index.broken-{int(time.time())}.json")
        try:
            path.replace(backup)
        except OSError:
            pass
        _log("index.json 读不懂，已备份后重建", 原因=err, 备份=backup)
        data = {}
    templates = [t for t in data.get("templates", []) if isinstance(t, dict) and t.get("id")]
    return {"version": 1, "templates": templates}


def _save(data: dict) -> None:
    path = _index_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)


def _find(data: dict, template_id: str) -> dict:
    for template in data["templates"]:
        if template["id"] == template_id:
            return template
    raise ValueError("没有这个目录结构模板（可能已经删掉了），请刷新页面")


def _check_id(template_id: str) -> str:
    if not re.fullmatch(r"[0-9a-f]{6,32}", template_id or ""):
        raise ValueError("模板编号不对")
    return template_id


def _clean_name(value, label: str) -> str:
    name = sanitize_filename(str(value or ""))[:MAX_NAME].rstrip(" .")
    if not name or name in {".", ".."}:
        raise ValueError(f"{label}请填写为 {MAX_NAME} 字以内的有效名称")
    return name


def _walk(nodes: list, depth: int = 1):
    if depth > MAX_DEPTH:
        raise ValueError(f"目录层级太深（最多 {MAX_DEPTH} 层）")
    for node in nodes:
        yield node, depth
        if node["kind"] == "dir":
            yield from _walk(node["children"], depth + 1)


def _clean_tree(raw, label: str) -> list:
    """整树校验并补 id：管理页本地编辑后整棵保存，已有 id 原样保留（出票时的排除/改名按 id 对应）。"""
    if not isinstance(raw, list):
        raise ValueError(f"{label}的目录树格式不对")
    count = 0

    def clean(nodes, depth):
        nonlocal count
        if depth > MAX_DEPTH:
            raise ValueError(f"目录层级太深（最多 {MAX_DEPTH} 层）")
        if not isinstance(nodes, list):
            raise ValueError(f"{label}的目录树格式不对")
        out = []
        seen = set()
        for item in nodes:
            count += 1
            if count > MAX_NODES:
                raise ValueError(f"目录文件太多（最多 {MAX_NODES} 个）")
            if not isinstance(item, dict):
                raise ValueError(f"{label}的目录树格式不对")
            kind = item.get("kind")
            if kind not in {"dir", "file"}:
                raise ValueError(f"{label}里有不认识的节点类型")
            name = _clean_name(item.get("name"), "目录名" if kind == "dir" else "文件名")
            if name in seen:
                raise ValueError(f"同一层里「{name}」重复了")
            seen.add(name)
            node = {"id": str(item.get("id") or uuid.uuid4().hex[:12]), "kind": kind, "name": name}
            if kind == "dir":
                node["children"] = clean(item.get("children") or [], depth + 1)
            else:
                source = str(item.get("source") or "").strip()
                if len(source) > MAX_SOURCE:
                    raise ValueError(f"「{name}」的来源路径太长")
                if source and not Path(source).is_absolute():
                    raise ValueError(f"「{name}」的来源请填写完整路径（如 D:\\模板\\{name}）")
                node["source"] = source
            out.append(node)
        return out

    return clean(raw, 1)


def _source_warnings(nodes: list) -> list[dict]:
    warnings = []
    for node, _ in _walk(nodes):
        if node["kind"] != "file":
            continue
        source = node.get("source") or ""
        if not source:
            warnings.append({"id": node["id"], "name": node["name"], "reason": "未指定来源文件"})
        elif not Path(source).is_file():
            warnings.append({"id": node["id"], "name": node["name"], "reason": f"来源文件不存在：{source}"})
    return warnings


def _public(template: dict) -> dict:
    return {
        "id": template["id"], "name": template.get("name") or "未命名",
        "created": template.get("created"), "updated": template.get("updated"),
        "tree": template.get("tree") or [],
    }


def get(template_id: str) -> dict:
    template_id = _check_id(template_id)
    with _lock:
        return _public(_find(_load(), template_id))


def listing() -> dict:
    with _lock:
        data = _load()
    templates = [_public(t) for t in data["templates"]]
    for template in templates:
        template["warnings"] = _source_warnings(template["tree"])
    return {"templates": templates, "dir": str(library_dir())}


# ---------------------------------------------------------------- 模板增删改

def create(name: str) -> dict:
    name = _clean_name(name, "模板名")
    with _lock:
        data = _load()
        if len(data["templates"]) >= MAX_TEMPLATES:
            raise ValueError(f"模板太多（最多 {MAX_TEMPLATES} 套），先删掉不用的")
        if any(t.get("name") == name for t in data["templates"]):
            raise ValueError(f"已经有叫「{name}」的模板了")
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        template = {"id": uuid.uuid4().hex[:12], "name": name, "created": now, "updated": now, "tree": []}
        data["templates"].insert(0, template)
        _save(data)
    _log("新建模板", 名字=name)
    return {"template": _public(template)}


def rename(template_id: str, name: str) -> dict:
    template_id = _check_id(template_id)
    name = _clean_name(name, "模板名")
    with _lock:
        data = _load()
        template = _find(data, template_id)
        if any(t.get("name") == name and t["id"] != template_id for t in data["templates"]):
            raise ValueError(f"已经有叫「{name}」的模板了")
        template["name"] = name
        template["updated"] = time.strftime("%Y-%m-%d %H:%M:%S")
        _save(data)
    _log("模板改名", 名字=name)
    return {"template": _public(template)}


def delete(template_id: str) -> dict:
    template_id = _check_id(template_id)
    with _lock:
        data = _load()
        template = _find(data, template_id)
        data["templates"] = [t for t in data["templates"] if t["id"] != template_id]
        _save(data)
    _log("删除模板", 名字=template.get("name"))
    return {"name": template.get("name")}


def save_tree(template_id: str, tree: list) -> dict:
    template_id = _check_id(template_id)
    cleaned = _clean_tree(tree, "模板")
    with _lock:
        data = _load()
        template = _find(data, template_id)
        template["tree"] = cleaned
        template["updated"] = time.strftime("%Y-%m-%d %H:%M:%S")
        _save(data)
    warnings = _source_warnings(cleaned)
    _log("保存目录树", 模板=template.get("name"), 节点数=sum(1 for _ in _walk(cleaned)), 来源警告=len(warnings))
    return {"template": _public(template), "warnings": warnings}


# ---------------------------------------------------------------- 出票时组建

def parse_overrides(payload: dict) -> dict:
    """出票表单里的临时调整：exclude 排除的节点 id；rename 节点 id -> 新名字。"""
    raw = payload.get("structure_overrides") or {}
    if not isinstance(raw, dict):
        raise ValueError("目录结构调整格式不对")
    exclude = raw.get("exclude", [])
    rename = raw.get("rename", {})
    if not isinstance(exclude, list) or len(exclude) > MAX_NODES or not all(isinstance(i, str) for i in exclude):
        raise ValueError("目录结构调整格式不对")
    if not isinstance(rename, dict) or len(rename) > MAX_NODES:
        raise ValueError("目录结构调整格式不对")
    cleaned = {}
    for node_id, name in rename.items():
        if not isinstance(node_id, str):
            raise ValueError("目录结构调整格式不对")
        cleaned[node_id] = _clean_name(name, "名称")
    return {"exclude": set(exclude), "rename": cleaned}


def materialize(template_id: str, folder: Path, overrides: dict | None = None) -> dict:
    """在 Ticket 文件夹下按模板建目录、复制来源文件。问题只记清单，不中断其余部分。"""
    template_id = _check_id(template_id)
    overrides = overrides or {"exclude": set(), "rename": {}}
    with _lock:
        template = _find(_load(), template_id)
    tree = template.get("tree") or []
    exclude, rename = overrides.get("exclude") or set(), overrides.get("rename") or {}
    dirs = files = 0
    problems: list[dict] = []
    items: list[dict] = []

    def build(nodes: list, parent: Path, rel: str):
        nonlocal dirs, files
        for node in nodes:
            if node["id"] in exclude:
                continue
            name = rename.get(node["id"]) or node["name"]
            target = parent / name
            path = f"{rel}/{name}" if rel else name
            if node["kind"] == "dir":
                try:
                    target.mkdir(exist_ok=True)
                except OSError as err:
                    problems.append({"path": path, "reason": f"建目录失败：{err}"})
                    continue
                dirs += 1
                items.append({"kind": "dir", "path": path})
                build(node["children"], target, path)
            else:
                source = node.get("source") or ""
                if not source:
                    problems.append({"path": path, "reason": "模板里未指定来源文件"})
                    continue
                if target.exists():
                    items.append({"kind": "file", "path": path, "exists": True})
                    continue
                try:
                    shutil.copyfile(source, target)
                except OSError as err:
                    problems.append({"path": path, "reason": f"复制失败（{source}）：{err}"})
                    continue
                files += 1
                items.append({"kind": "file", "path": path})

    build(tree, folder, "")
    result = {"template": template.get("name"), "dirs": dirs, "files": files,
              "problems": problems, "items": items}
    _log("按模板组建", 模板=template.get("name"), 目录=dirs, 文件=files, 问题=len(problems), 文件夹=folder)
    return result