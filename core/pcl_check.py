"""PCL 单体测试检查台（static/UIPCL.html，网址 /pcl，用户 2026-11-12 要）。

页面上的 22 项检查清单和 10 条源码静态规则仍在浏览器里跑（static/pcl_rules.js），这里负责：
- 按「检查对象」（一个画面 / 一张 Redmine 票）分开保存每一项的测试结果：结论 OK / NG / 不适用、测试人、
  测试时间、备注、截图，结论每改一次记一条历史；都存在 execution.sqlite3（程序数据文件夹，见 core/data_dir.py）
- 不是每一项都要测：每个对象记下勾选了哪些项目（没存过 = 全选），进度和覆盖率只算勾选的；
  点「测试」时页面只跑选中项目（或单独某一项）的自动规则，每项最近一次的测试结果存在这里
- 保存这个对象的源码卡片；按本地路径读取源码文件（文件夹 = 里面的 .twig / .php / .sql / .js）
- 读取 PHPUnit 生成的 clover.xml，算出每个文件的行 / 方法 / 条件覆盖率和没执行到的行
- 导出 / 导入 JSON：version 2 带结论、截图、源码、覆盖率；也能导入旧页面导出的 version 1（已完成 → OK）
"""
from __future__ import annotations

import base64
import binascii
import json
import os
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Any

from core import storage

EXPORT_FORMAT = "pcl-qa-checklist"
ITEM_RE = re.compile(r"^T\d{2}$")
VERDICTS = ("", "ok", "ng", "na")
FINDING_STATUSES = ("warn", "unknown", "confirmed", "skipped")
RUN_STATUSES = FINDING_STATUSES + ("none",)  # none = 测过了，代码里没有和这一项相关的写法
MAX_RUN_FINDINGS = 40
MAX_RUN_LINES = 30  # 每个文件最多存这么多行摘录，总行数另记在 count 里
SOURCE_TYPES = ("twig", "php", "sql")
SOURCE_EXTS = {".twig", ".php", ".sql", ".js"}
SKIP_DIRS = {".git", ".svn", "node_modules", "vendor", "cache"}
SOURCE_LIMIT = 200_000  # 和页面粘贴框的上限一样
MAX_SOURCE_FILES = 300
MAX_SOURCE_TOTAL = 20_000_000
MAX_COVERAGE_BYTES = 50 * 1024 * 1024
MAX_IMAGE_BYTES = 8 * 1024 * 1024
IMAGE_SIGNATURES = {
    "image/png": (b"\x89PNG\r\n\x1a\n",),
    "image/jpeg": (b"\xff\xd8\xff",),
    "image/gif": (b"GIF87a", b"GIF89a"),
    "image/webp": (b"RIFF",),
}


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _connect():
    db = storage._connect()
    db.executescript(
        "CREATE TABLE IF NOT EXISTS pcl_targets (id INTEGER PRIMARY KEY, name TEXT NOT NULL, note TEXT NOT NULL DEFAULT '',"
        " created_at TEXT NOT NULL, updated_at TEXT NOT NULL);"
        "CREATE TABLE IF NOT EXISTS pcl_results (target_id INTEGER NOT NULL, item TEXT NOT NULL,"
        " verdict TEXT NOT NULL DEFAULT '', tester TEXT NOT NULL DEFAULT '', tested_at TEXT NOT NULL DEFAULT '',"
        " note TEXT NOT NULL DEFAULT '', PRIMARY KEY (target_id, item));"
        "CREATE TABLE IF NOT EXISTS pcl_history (id INTEGER PRIMARY KEY, target_id INTEGER NOT NULL, item TEXT NOT NULL,"
        " verdict TEXT NOT NULL, tester TEXT NOT NULL, at TEXT NOT NULL);"
        "CREATE TABLE IF NOT EXISTS pcl_images (id INTEGER PRIMARY KEY, target_id INTEGER NOT NULL, item TEXT NOT NULL,"
        " name TEXT NOT NULL, mime TEXT NOT NULL, data BLOB NOT NULL, created_at TEXT NOT NULL);"
        "CREATE TABLE IF NOT EXISTS pcl_sources (target_id INTEGER PRIMARY KEY, payload TEXT NOT NULL);"
        "CREATE TABLE IF NOT EXISTS pcl_coverage (target_id INTEGER PRIMARY KEY, payload TEXT NOT NULL);"
        "CREATE TABLE IF NOT EXISTS pcl_selection (target_id INTEGER PRIMARY KEY, payload TEXT NOT NULL);"
        "CREATE TABLE IF NOT EXISTS pcl_runs (target_id INTEGER NOT NULL, item TEXT NOT NULL, payload TEXT NOT NULL,"
        " PRIMARY KEY (target_id, item));"
    )
    return db


def _text(value: Any, label: str, limit: int, *, required: bool = False) -> str:
    if value is None and not required:
        return ""
    if not isinstance(value, str):
        raise ValueError(f"{label}格式不正确")
    value = value.strip() if limit <= 200 else value
    if required and not value.strip():
        raise ValueError(f"请填写{label}")
    if len(value) > limit:
        raise ValueError(f"{label}太长（最多 {limit} 个字）")
    return value


def _item(value: Any) -> str:
    if not isinstance(value, str) or not ITEM_RE.match(value):
        raise ValueError("检查项编号不正确")
    return value


def _target_id(db, value: Any) -> int:
    if isinstance(value, str) and value.isdigit():
        value = int(value)
    if not isinstance(value, int) or isinstance(value, bool):
        raise ValueError("请先选择检查对象")
    if db.execute("SELECT 1 FROM pcl_targets WHERE id=?", (value,)).fetchone() is None:
        raise ValueError("这个检查对象已经不存在了，请刷新页面")
    return value


def _touch(db, target: int) -> None:
    db.execute("UPDATE pcl_targets SET updated_at=? WHERE id=?", (_now(), target))


# ---------- 检查对象 ----------

def list_targets() -> list[dict[str, Any]]:
    db = _connect()
    try:
        rows = db.execute("SELECT id, name, note, created_at, updated_at FROM pcl_targets ORDER BY updated_at DESC, id DESC").fetchall()
        selections = {target: set(json.loads(payload)) for target, payload in db.execute("SELECT target_id, payload FROM pcl_selection")}
        counts: dict[int, dict[str, int]] = {}
        for target, item, verdict in db.execute("SELECT target_id, item, verdict FROM pcl_results WHERE verdict != ''"):
            if target not in selections or item in selections[target]:  # 没勾选的项目不计入进度
                counts.setdefault(target, {}).setdefault(verdict, 0)
                counts[target][verdict] += 1
    finally:
        db.close()
    return [{"id": i, "name": name, "note": note, "created_at": created, "updated_at": updated,
             "selection": sorted(selections[i]) if i in selections else None,
             "counts": {v: counts.get(i, {}).get(v, 0) for v in VERDICTS[1:]}}
            for i, name, note, created, updated in rows]


def create_target(name: Any, note: Any = "") -> dict[str, Any]:
    name = _text(name, "检查对象名称", 120, required=True)
    note = _text(note, "说明", 2000)
    db = _connect()
    try:
        with db:
            now = _now()
            target = db.execute("INSERT INTO pcl_targets(name, note, created_at, updated_at) VALUES (?, ?, ?, ?)",
                                (name, note, now, now)).lastrowid
    finally:
        db.close()
    return next(t for t in list_targets() if t["id"] == target)


def update_target(target: Any, name: Any, note: Any = None) -> dict[str, Any]:
    name = _text(name, "检查对象名称", 120, required=True)
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            if note is None:
                db.execute("UPDATE pcl_targets SET name=?, updated_at=? WHERE id=?", (name, _now(), target))
            else:
                db.execute("UPDATE pcl_targets SET name=?, note=?, updated_at=? WHERE id=?",
                           (name, _text(note, "说明", 2000), _now(), target))
    finally:
        db.close()
    return next(t for t in list_targets() if t["id"] == target)


def _clear(db, target: int, *, keep_target: bool) -> None:
    tables = ["pcl_results", "pcl_history", "pcl_images", "pcl_sources", "pcl_coverage", "pcl_selection", "pcl_runs"]
    for table in tables:
        db.execute(f"DELETE FROM {table} WHERE target_id=?", (target,))
    if not keep_target:
        db.execute("DELETE FROM pcl_targets WHERE id=?", (target,))


def delete_target(target: Any) -> None:
    db = _connect()
    try:
        with db:
            _clear(db, _target_id(db, target), keep_target=False)
    finally:
        db.close()


def reset_results(target: Any) -> None:
    """清空这个对象 22 项的结论、备注、截图、历史和测试结果；勾选了哪些项目、源码卡片和覆盖率留着"""
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            for table in ("pcl_results", "pcl_history", "pcl_images", "pcl_runs"):
                db.execute(f"DELETE FROM {table} WHERE target_id=?", (target,))
            _touch(db, target)
    finally:
        db.close()


def _results(db, target: int, only: str | None = None) -> dict[str, dict[str, Any]]:
    where, args = ("target_id=?", (target,)) if only is None else ("target_id=? AND item=?", (target, only))
    results: dict[str, dict[str, Any]] = {}

    def entry(item: str) -> dict[str, Any]:
        return results.setdefault(item, {"verdict": "", "tester": "", "tested_at": "", "note": "", "images": [], "history": []})

    for item, verdict, tester, tested_at, note in db.execute(
            f"SELECT item, verdict, tester, tested_at, note FROM pcl_results WHERE {where}", args):
        entry(item).update(verdict=verdict, tester=tester, tested_at=tested_at, note=note)
    for image, item, name, mime, size, created in db.execute(
            f"SELECT id, item, name, mime, length(data), created_at FROM pcl_images WHERE {where} ORDER BY id", args):
        entry(item)["images"].append({"id": image, "name": name, "mime": mime, "size": size, "created_at": created})
    for item, verdict, tester, at in db.execute(
            f"SELECT item, verdict, tester, at FROM pcl_history WHERE {where} ORDER BY id", args):
        entry(item)["history"].append({"verdict": verdict, "tester": tester, "at": at})
    return results


def _payload(db, table: str, target: int) -> Any:
    row = db.execute(f"SELECT payload FROM {table} WHERE target_id=?", (target,)).fetchone()
    return json.loads(row[0]) if row else None


def load_target(target: Any) -> dict[str, Any]:
    db = _connect()
    try:
        target = _target_id(db, target)
        info = next(t for t in list_targets() if t["id"] == target)
        runs = {item: json.loads(payload) for item, payload in
                db.execute("SELECT item, payload FROM pcl_runs WHERE target_id=?", (target,))}
        return {"target": info, "results": _results(db, target), "sources": _payload(db, "pcl_sources", target) or [],
                "coverage": _payload(db, "pcl_coverage", target), "runs": runs}
    finally:
        db.close()


# ---------- 选哪些项目测 + 每项最近一次的测试结果 ----------

def _clean_selection(items: Any) -> list[str]:
    if not isinstance(items, list):
        raise ValueError("选中的项目格式不正确")
    return sorted({_item(item) for item in items})


def save_selection(target: Any, items: Any) -> None:
    items = _clean_selection(items)
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            db.execute("INSERT INTO pcl_selection(target_id, payload) VALUES (?, ?) "
                       "ON CONFLICT(target_id) DO UPDATE SET payload=excluded.payload", (target, json.dumps(items)))
            _touch(db, target)
    finally:
        db.close()


def _clean_run(run: Any) -> dict[str, Any]:
    """页面跑完规则送来的一项结果：只留认识的字段，摘录太多的截掉"""
    if not isinstance(run, dict) or run.get("status") not in RUN_STATUSES or not isinstance(run.get("findings", []), list):
        raise ValueError("测试结果格式不正确")
    findings = []
    for finding in run.get("findings", [])[:MAX_RUN_FINDINGS]:
        if not isinstance(finding, dict) or finding.get("status") not in FINDING_STATUSES:
            raise ValueError("测试结果格式不正确")
        files = []
        for file in finding.get("files") or []:
            if not isinstance(file, dict) or not isinstance(file.get("lines"), list):
                raise ValueError("测试结果格式不正确")
            lines = [[_int(line[0]), str(line[1])[:400]] for line in file["lines"][:MAX_RUN_LINES]
                     if isinstance(line, list) and len(line) == 2]
            files.append({"name": _text(file.get("name"), "文件名", 260), "count": max(_int(file.get("count")), len(lines)),
                          "lines": lines})
        findings.append({"status": finding["status"], "message": str(finding.get("message") or "")[:600],
                         "extra": [str(text)[:600] for text in (finding.get("extra") or [])[:10]], "files": files})
    return {"status": run["status"], "sig": _text(run.get("sig"), "代码版本", 120), "findings": findings,
            "tester": _text(run.get("tester"), "测试人", 60), "at": _text(run.get("at"), "测试时间", 32)}


def save_runs(target: Any, runs: Any, tester: Any = "") -> dict[str, Any]:
    """保存这次测的几项（没测的项目不动），返回带上测试人和时间的结果"""
    if not isinstance(runs, dict) or not runs:
        raise ValueError("没有要保存的测试结果")
    tester, now = _text(tester, "测试人", 60), _now()
    clean = {_item(item): {**_clean_run(run), "tester": tester, "at": now} for item, run in runs.items()}
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            _write_runs(db, target, clean)
            _touch(db, target)
    finally:
        db.close()
    return clean


def _write_runs(db, target: int, runs: dict[str, dict[str, Any]]) -> None:
    db.executemany("INSERT INTO pcl_runs(target_id, item, payload) VALUES (?, ?, ?) "
                   "ON CONFLICT(target_id, item) DO UPDATE SET payload=excluded.payload",
                   [(target, item, json.dumps(run, ensure_ascii=False)) for item, run in runs.items()])


# ---------- 每一项的测试结果 ----------

def save_result(target: Any, item: Any, changes: dict[str, Any]) -> dict[str, Any]:
    """只改传进来的字段。结论变了：记一条历史；没另外给时间就用现在（改回待测时清空时间）。"""
    item = _item(item)
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            current = _results(db, target, item).get(item, {"verdict": "", "tester": "", "tested_at": "", "note": ""})
            new = {key: current[key] for key in ("verdict", "tester", "tested_at", "note")}
            if "verdict" in changes:
                if changes["verdict"] not in VERDICTS:
                    raise ValueError("结论只能是 OK / NG / 不适用")
                new["verdict"] = changes["verdict"]
            if "tester" in changes:
                new["tester"] = _text(changes["tester"], "测试人", 60)
            if "note" in changes:
                new["note"] = _text(changes["note"], "备注", 20000)
            if "tested_at" in changes:
                new["tested_at"] = _text(changes["tested_at"], "测试时间", 32)
            elif new["verdict"] != current["verdict"]:
                new["tested_at"] = _now() if new["verdict"] else ""
            db.execute(
                "INSERT INTO pcl_results(target_id, item, verdict, tester, tested_at, note) VALUES (?, ?, ?, ?, ?, ?) "
                "ON CONFLICT(target_id, item) DO UPDATE SET verdict=excluded.verdict, tester=excluded.tester, "
                "tested_at=excluded.tested_at, note=excluded.note",
                (target, item, new["verdict"], new["tester"], new["tested_at"], new["note"]))
            if new["verdict"] != current["verdict"]:
                db.execute("INSERT INTO pcl_history(target_id, item, verdict, tester, at) VALUES (?, ?, ?, ?, ?)",
                           (target, item, new["verdict"], new["tester"], _now()))
            _touch(db, target)
            return _results(db, target, item)[item]
    finally:
        db.close()


def _decode_image(data_url: Any) -> tuple[str, bytes]:
    if not isinstance(data_url, str):
        raise ValueError("截图格式不正确")
    match = re.match(r"^data:(image/[\w.+-]+);base64,", data_url)
    if not match or match.group(1) not in IMAGE_SIGNATURES:
        raise ValueError("只能贴 PNG / JPEG / GIF / WebP 图片")
    mime = match.group(1)
    if (len(data_url) - match.end()) * 3 // 4 > MAX_IMAGE_BYTES:
        raise ValueError(f"图片太大（最多 {MAX_IMAGE_BYTES // 1024 // 1024} MB）")
    try:
        data = base64.b64decode(data_url[match.end():], validate=True)
    except (binascii.Error, ValueError):
        raise ValueError("图片数据已损坏") from None
    if not data.startswith(IMAGE_SIGNATURES[mime]) or (mime == "image/webp" and data[8:12] != b"WEBP"):
        raise ValueError("图片内容和格式对不上")
    return mime, data


def add_image(target: Any, item: Any, data_url: Any, name: Any = "") -> dict[str, Any]:
    item = _item(item)
    mime, data = _decode_image(data_url)
    name = _text(name, "图片名", 120) or f"截图 {_now()}"
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            created = _now()
            image = db.execute("INSERT INTO pcl_images(target_id, item, name, mime, data, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                               (target, item, name, mime, data, created)).lastrowid
            _touch(db, target)
    finally:
        db.close()
    return {"id": image, "name": name, "mime": mime, "size": len(data), "created_at": created}


def get_image(image: Any) -> tuple[str, bytes]:
    if not isinstance(image, str) or not image.isdigit():
        raise ValueError("截图编号不正确")
    db = _connect()
    try:
        row = db.execute("SELECT mime, data FROM pcl_images WHERE id=?", (int(image),)).fetchone()
    finally:
        db.close()
    if row is None:
        raise ValueError("截图已经删掉了")
    return row[0], bytes(row[1])


def delete_image(image: Any) -> None:
    if not isinstance(image, int) or isinstance(image, bool):
        raise ValueError("截图编号不正确")
    db = _connect()
    try:
        with db:
            row = db.execute("SELECT target_id FROM pcl_images WHERE id=?", (image,)).fetchone()
            if row:
                db.execute("DELETE FROM pcl_images WHERE id=?", (image,))
                _touch(db, row[0])
    finally:
        db.close()


# ---------- 源码 ----------

def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp932"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")
    return text.replace("\r\n", "\n").replace("\r", "\n")


def _folder_sources(folder: Path) -> list[tuple[Path, str]]:
    found = []
    for root, dirs, names in os.walk(folder):
        dirs[:] = sorted(d for d in dirs if d.lower() not in SKIP_DIRS and not d.startswith("."))
        for name in sorted(names):
            path = Path(root, name)
            if path.suffix.lower() in SOURCE_EXTS and not name.endswith(".min.js"):
                found.append((path, path.relative_to(folder.parent).as_posix()))
    return found


def read_sources(paths_text: Any) -> dict[str, Any]:
    """每行一个文件或文件夹，读出源码交给页面校验；找不到 / 太大 / 读不了的列出来，其他照常读"""
    if not isinstance(paths_text, str) or not paths_text.strip():
        raise ValueError("请填写源码文件或文件夹的路径，每行一个")
    files: list[dict[str, str]] = []
    missing: list[dict[str, str]] = []
    skipped: list[dict[str, str]] = []
    seen: set[str] = set()
    total = 0
    for line in paths_text.splitlines():
        line = line.strip().strip('"').strip("'")
        if not line:
            continue
        path = Path(line).expanduser()
        if not path.exists():
            missing.append({"path": line, "reason": "找不到这个文件或文件夹"})
            continue
        candidates = [(path, path.name)] if path.is_file() else _folder_sources(path)
        if path.is_dir() and not candidates:
            missing.append({"path": line, "reason": "文件夹里没有 .twig / .php / .sql / .js 文件"})
        for source, name in candidates:
            key = os.path.normcase(str(source.resolve()))
            if key in seen:
                continue
            seen.add(key)
            if len(files) >= MAX_SOURCE_FILES:
                skipped.append({"path": str(source), "reason": f"一次最多读 {MAX_SOURCE_FILES} 个文件"})
                continue
            try:
                code = _read_text(source)
            except OSError as err:
                skipped.append({"path": str(source), "reason": f"读不了：{err.strerror or err}"})
                continue
            if len(code) > SOURCE_LIMIT:
                skipped.append({"path": str(source), "reason": "超过 20 万字符，请拆开或粘贴需要的部分"})
                continue
            if total + len(code) > MAX_SOURCE_TOTAL:
                skipped.append({"path": str(source), "reason": "一次读取的源码太多了"})
                continue
            total += len(code)
            files.append({"name": name, "path": str(source), "code": code})
    if not files and not skipped:
        raise ValueError("没有读到源码：" + "；".join(f"{m['path']}（{m['reason']}）" for m in missing[:3]))
    return {"files": files, "missing": missing, "skipped": skipped}


def _clean_sources(sources: Any) -> list[dict[str, str]]:
    if not isinstance(sources, list):
        raise ValueError("源码卡片格式不正确")
    if len(sources) > MAX_SOURCE_FILES:
        raise ValueError(f"源码卡片最多 {MAX_SOURCE_FILES} 个")
    clean, total = [], 0
    for source in sources:
        if not isinstance(source, dict) or source.get("type") not in SOURCE_TYPES:
            raise ValueError("源码卡片格式不正确")
        code = _text(source.get("code"), "源码", SOURCE_LIMIT, required=True)
        total += len(code)
        clean.append({"name": _text(source.get("name"), "文件名", 260, required=True), "type": source["type"],
                      "code": code, "path": _text(source.get("path"), "路径", 1000)})
    if total > MAX_SOURCE_TOTAL:
        raise ValueError("源码卡片的内容太多了，请删掉一些")
    return clean


def save_sources(target: Any, sources: Any) -> None:
    sources = _clean_sources(sources)
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            db.execute("INSERT INTO pcl_sources(target_id, payload) VALUES (?, ?) "
                       "ON CONFLICT(target_id) DO UPDATE SET payload=excluded.payload",
                       (target, json.dumps(sources, ensure_ascii=False)))
            _touch(db, target)
    finally:
        db.close()


# ---------- 代码覆盖率（PHPUnit clover.xml） ----------

def _int(value: Any) -> int:
    try:
        return max(0, int(value))
    except (TypeError, ValueError):
        return 0


def _ranges(numbers: list[int]) -> list[list[int]]:
    ranges: list[list[int]] = []
    for n in sorted(set(numbers)):
        if ranges and n == ranges[-1][1] + 1:
            ranges[-1][1] = n
        else:
            ranges.append([n, n])
    return ranges


def _common_dir(names: list[str]) -> str:
    parts = [name.replace("\\", "/").split("/")[:-1] for name in names]
    if not parts:
        return ""
    common = parts[0]
    for other in parts[1:]:
        i = 0
        while i < min(len(common), len(other)) and common[i] == other[i]:
            i += 1
        common = common[:i]
    return "/".join(common) + "/" if common else ""


def parse_clover(text: Any) -> dict[str, Any]:
    if not isinstance(text, str) or not text.strip():
        raise ValueError("覆盖率文件是空的")
    if re.search(r"<!(?:DOCTYPE|ENTITY)", text, re.IGNORECASE):
        raise ValueError("覆盖率文件里有 DOCTYPE / ENTITY 声明，PHPUnit 生成的 clover.xml 不会有，不读取")
    try:
        root = ET.fromstring(text)
    except ET.ParseError as err:
        raise ValueError(f"不是有效的 XML：{err}") from None
    project = root.find("project")
    if root.tag != "coverage" or project is None:
        raise ValueError("不是 clover 覆盖率文件（应以 <coverage><project> 开头，用 phpunit --coverage-clover 生成）")
    files = []
    for node in root.iter("file"):
        name = node.get("name") or node.get("path") or ""
        lines = node.findall("line")
        statements = [line for line in lines if line.get("type") in ("stmt", "cond")]
        methods = [line for line in lines if line.get("type") == "method"]
        uncovered = [_int(line.get("num")) for line in statements if _int(line.get("count")) == 0]
        partial = [_int(line.get("num")) for line in lines if line.get("type") == "cond" and _int(line.get("count")) > 0
                   and (line.get("truecount") == "0" or line.get("falsecount") == "0")]
        metrics = node.find("metrics")
        if metrics is not None and metrics.get("statements") is not None:
            stmt = (_int(metrics.get("coveredstatements")), _int(metrics.get("statements")))
            meth = (_int(metrics.get("coveredmethods")), _int(metrics.get("methods")))
            cond = (_int(metrics.get("coveredconditionals")), _int(metrics.get("conditionals")))
        else:
            stmt = (len(statements) - len(uncovered), len(statements))
            meth = (sum(_int(m.get("count")) > 0 for m in methods), len(methods))
            cond = (0, 0)
        files.append({
            "name": name,
            "lines": {"covered": stmt[0], "total": stmt[1]},
            "methods": {"covered": meth[0], "total": meth[1]},
            "conditionals": {"covered": cond[0], "total": cond[1]},
            "uncovered": _ranges(uncovered),
            "partial": _ranges(partial),
            "uncovered_methods": [m.get("name") or f"第 {m.get('num')} 行" for m in methods if _int(m.get("count")) == 0],
        })
    if not files:
        raise ValueError("覆盖率文件里没有任何文件的数据")
    prefix = _common_dir([f["name"] for f in files]) if len(files) > 1 else _common_dir([files[0]["name"]])
    for f in files:
        f["short"] = f["name"].replace("\\", "/")[len(prefix):] or f["name"]
    files.sort(key=lambda f: f["short"].casefold())
    totals = {key: {"covered": sum(f[key]["covered"] for f in files), "total": sum(f[key]["total"] for f in files)}
              for key in ("lines", "methods", "conditionals")}
    stamp = project.get("timestamp") or root.get("generated")
    generated = datetime.fromtimestamp(int(stamp)).strftime("%Y-%m-%d %H:%M") if stamp and stamp.isdigit() else ""
    return {"generated": generated, "prefix": prefix, "files": files, "totals": totals}


def import_coverage(target: Any, path: Any = None, xml: Any = None, name: Any = "") -> dict[str, Any]:
    if isinstance(path, str) and path.strip():
        source = Path(path.strip().strip('"').strip("'")).expanduser()
        if not source.is_file():
            raise ValueError(f"找不到覆盖率文件：{source}")
        if source.stat().st_size > MAX_COVERAGE_BYTES:
            raise ValueError("覆盖率文件太大（最多 50 MB）")
        xml, label = source.read_text(encoding="utf-8", errors="replace"), str(source)
    else:
        if isinstance(xml, str) and len(xml) > MAX_COVERAGE_BYTES:
            raise ValueError("覆盖率文件太大（最多 50 MB）")
        label = _text(name, "文件名", 260) or "clover.xml"
    coverage = parse_clover(xml)
    coverage.update(source=label, imported_at=_now())
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            db.execute("INSERT INTO pcl_coverage(target_id, payload) VALUES (?, ?) "
                       "ON CONFLICT(target_id) DO UPDATE SET payload=excluded.payload",
                       (target, json.dumps(coverage, ensure_ascii=False)))
            _touch(db, target)
    finally:
        db.close()
    return coverage


def clear_coverage(target: Any) -> None:
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            db.execute("DELETE FROM pcl_coverage WHERE target_id=?", (target,))
            _touch(db, target)
    finally:
        db.close()


# ---------- 导出 / 导入 ----------

def export_target(target: Any) -> dict[str, Any]:
    """页面再给每一项补上标题、日文原文、有没有勾选和最近一次的测试结果（autoCheck）后下载"""
    data = load_target(target)
    db = _connect()
    try:
        images: dict[int, str] = {
            image: f"data:{mime};base64,{base64.b64encode(blob).decode('ascii')}"
            for image, mime, blob in db.execute("SELECT id, mime, data FROM pcl_images WHERE target_id=?", (data["target"]["id"],))
        }
    finally:
        db.close()
    items = []
    for item, result in sorted(data["results"].items()):
        items.append({"id": item, "done": bool(result["verdict"]), "verdict": result["verdict"], "tester": result["tester"],
                      "testedAt": result["tested_at"], "note": result["note"], "history": result["history"],
                      "images": [{"name": i["name"], "dataUrl": images[i["id"]]} for i in result["images"] if i["id"] in images]})
    return {"format": EXPORT_FORMAT, "version": 2, "exportedAt": datetime.now().isoformat(timespec="seconds"),
            "target": {"name": data["target"]["name"], "note": data["target"]["note"]},
            "selection": data["target"]["selection"], "runs": data["runs"],
            "items": items, "sources": data["sources"], "coverage": data["coverage"]}


def import_results(target: Any, data: Any) -> dict[str, Any]:
    """覆盖这个对象的结论、备注、截图、历史和测试结果；version 2 带了勾选 / 源码 / 覆盖率时也一起换掉"""
    if not isinstance(data, dict) or data.get("format") != EXPORT_FORMAT or data.get("version") not in (1, 2) \
            or not isinstance(data.get("items"), list):
        raise ValueError("文件格式不正确（需要检查台导出的 JSON）")
    v2 = data["version"] == 2
    rows, images, history = [], [], []
    for entry in data["items"]:
        if not isinstance(entry, dict) or not isinstance(entry.get("id"), str) or not ITEM_RE.match(entry["id"]):
            continue
        item = entry["id"]
        verdict = entry.get("verdict") if v2 and entry.get("verdict") in VERDICTS else ("ok" if entry.get("done") is True else "")
        tester = _text(entry.get("tester"), "测试人", 60) if v2 else ""
        tested_at = _text(entry.get("testedAt"), "测试时间", 32) if v2 else ""
        note = entry.get("note") if isinstance(entry.get("note"), str) else ""
        rows.append((item, verdict, tester, tested_at, _text(note, "备注", 20000)))
        for image in (entry.get("images") or []) if v2 else []:
            if isinstance(image, dict):
                mime, blob = _decode_image(image.get("dataUrl"))
                images.append((item, _text(image.get("name"), "图片名", 120) or "截图", mime, blob))
        for change in (entry.get("history") or []) if v2 else []:
            if isinstance(change, dict) and change.get("verdict") in VERDICTS:
                history.append((item, change["verdict"], _text(change.get("tester"), "测试人", 60), _text(change.get("at"), "时间", 32)))
    sources = _clean_sources(data["sources"]) if v2 and isinstance(data.get("sources"), list) else None
    coverage = data.get("coverage") if v2 and isinstance(data.get("coverage"), dict) else None
    selection = _clean_selection(data["selection"]) if v2 and isinstance(data.get("selection"), list) else None
    runs = {_item(item): _clean_run(run) for item, run in data["runs"].items()} \
        if v2 and isinstance(data.get("runs"), dict) else {}
    db = _connect()
    try:
        with db:
            target = _target_id(db, target)
            for table in ("pcl_results", "pcl_history", "pcl_images", "pcl_runs"):
                db.execute(f"DELETE FROM {table} WHERE target_id=?", (target,))
            _write_runs(db, target, runs)
            if selection is not None:
                db.execute("INSERT INTO pcl_selection(target_id, payload) VALUES (?, ?) "
                           "ON CONFLICT(target_id) DO UPDATE SET payload=excluded.payload", (target, json.dumps(selection)))
            now = _now()
            db.executemany("INSERT INTO pcl_results(target_id, item, verdict, tester, tested_at, note) VALUES (?, ?, ?, ?, ?, ?)",
                           [(target, *row) for row in rows])
            db.executemany("INSERT INTO pcl_images(target_id, item, name, mime, data, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                           [(target, *image, now) for image in images])
            db.executemany("INSERT INTO pcl_history(target_id, item, verdict, tester, at) VALUES (?, ?, ?, ?, ?)",
                           [(target, *change) for change in history])
            if sources is not None:
                db.execute("INSERT INTO pcl_sources(target_id, payload) VALUES (?, ?) "
                           "ON CONFLICT(target_id) DO UPDATE SET payload=excluded.payload",
                           (target, json.dumps(sources, ensure_ascii=False)))
            if coverage is not None:
                db.execute("INSERT INTO pcl_coverage(target_id, payload) VALUES (?, ?) "
                           "ON CONFLICT(target_id) DO UPDATE SET payload=excluded.payload",
                           (target, json.dumps(coverage, ensure_ascii=False)))
            _touch(db, target)
    finally:
        db.close()
    return {"items": len(rows), "images": len(images)}


# ---------- 网页接口 ----------

def handle_post(action: str, payload: dict[str, Any]) -> dict[str, Any]:
    """/api/pcl/<action>；和 Excel 任务互不影响，别的任务在跑也能用"""
    target = payload.get("target")
    if action == "targets/create":
        return {"target": create_target(payload.get("name"), payload.get("note", ""))}
    if action == "targets/update":
        return {"target": update_target(target, payload.get("name"), payload.get("note"))}
    if action == "targets/delete":
        delete_target(target)
        return {}
    if action == "result":
        changes = {key: payload[key] for key in ("verdict", "tester", "tested_at", "note") if key in payload}
        return {"result": save_result(target, payload.get("item"), changes)}
    if action == "reset":
        reset_results(target)
        return {}
    if action == "selection":
        save_selection(target, payload.get("items"))
        return {}
    if action == "runs":
        return {"runs": save_runs(target, payload.get("runs"), payload.get("tester", ""))}
    if action == "image":
        return {"image": add_image(target, payload.get("item"), payload.get("data"), payload.get("name", ""))}
    if action == "image/delete":
        delete_image(payload.get("id"))
        return {}
    if action == "sources/read":
        return read_sources(payload.get("paths"))
    if action == "sources/save":
        save_sources(target, payload.get("sources"))
        return {}
    if action == "coverage":
        return {"coverage": import_coverage(target, payload.get("path"), payload.get("xml"), payload.get("name", ""))}
    if action == "coverage/clear":
        clear_coverage(target)
        return {}
    if action == "import":
        return import_results(target, payload.get("data"))
    raise ValueError(f"不认识的操作：{action}")
