"""扫描阶段的独立计时与报告；详细数据只在分析页面读取，不进入工作台轮询。"""
import json
import sqlite3
import time
from collections import Counter
from contextlib import contextmanager
from contextvars import ContextVar
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

from core import storage

OPERATIONS = {
    "open": "打开工作簿", "layout": "打印方向分析", "export": "PDF 导出",
    "detect": "切字检测与定位", "save_view": "保存分页预览", "highlight": "标色与转人工",
    "restart": "重启 Excel",
}
_current = ContextVar("scan_analysis_file", default=None)


def _now():
    return datetime.now(timezone.utc).isoformat()


@contextmanager
def _connect():
    db = sqlite3.connect(storage.DATABASE_PATH, timeout=30)
    try:
        with db:
            db.execute("CREATE TABLE IF NOT EXISTS scan_analysis_runs "
                       "(id TEXT PRIMARY KEY, started_at TEXT NOT NULL, folder TEXT NOT NULL, payload TEXT NOT NULL)")
            db.execute("CREATE TABLE IF NOT EXISTS scan_analysis_files "
                       "(run_id TEXT NOT NULL, path TEXT NOT NULL, payload TEXT NOT NULL, PRIMARY KEY(run_id, path))")
            yield db
    finally:
        db.close()


def _save_run(run):
    with _connect() as db:
        db.execute("INSERT INTO scan_analysis_runs VALUES (?, ?, ?, ?) ON CONFLICT(id) DO UPDATE SET payload=excluded.payload",
                   (run["id"], run["started_at"], run["folder"], json.dumps(run, ensure_ascii=False)))


def _save_file(run_id, row):
    with _connect() as db:
        db.execute("INSERT INTO scan_analysis_files VALUES (?, ?, ?) ON CONFLICT(run_id, path) DO UPDATE SET payload=excluded.payload",
                   (run_id, row["path"], json.dumps(row, ensure_ascii=False)))


def _file(run_id, path):
    with _connect() as db:
        found = db.execute("SELECT payload FROM scan_analysis_files WHERE run_id=? AND path=?", (run_id, path)).fetchone()
    return json.loads(found[0]) if found else None


def history(folder=None):
    with _connect() as db:
        rows = db.execute("SELECT payload FROM scan_analysis_runs " + ("WHERE folder=? " if folder else "")
                          + "ORDER BY started_at DESC LIMIT 100", (folder,) if folder else ()).fetchall()
    return [json.loads(row[0]) for row in rows]


def begin(job, folder, paths, mode):
    """同一任务续跑使用同一报告，保留失败尝试的耗时。"""
    run_id = job.get("analysis_id")
    run = None
    if run_id:
        with _connect() as db:
            found = db.execute("SELECT payload FROM scan_analysis_runs WHERE id=?", (run_id,)).fetchone()
        run = json.loads(found[0]) if found else None
    if run is None:
        run = {"id": uuid4().hex, "folder": str(folder), "mode": mode, "started_at": _now(), "total": len(paths)}
        with _connect() as db:
            for path in paths:
                path = str(path)
                try:
                    rel = Path(path).relative_to(folder)
                    system = rel.parts[0] if len(rel.parts) > 1 else "根目录文件"
                    name = rel.as_posix()
                except ValueError:
                    system, name = "其他目录", Path(path).name
                row = {"path": path, "name": name, "system": system, "status": "pending", "seconds": 0,
                       "attempts": 0, "operations": {}, "sheets": {}, "issues": {}, "exempt_sheets": [],
                       "skipped_sheets": [], "unconfirmed": 0, "manual_only": 0}
                db.execute("INSERT INTO scan_analysis_files VALUES (?, ?, ?)",
                           (run["id"], path, json.dumps(row, ensure_ascii=False)))
    run.update(status="running", finished_at=None)
    _save_run(run)
    job["analysis_id"] = run["id"]
    return {"id": run["id"], "folder": run["folder"]}


@contextmanager
def attempt(job, path):
    run_id = job.get("analysis_id")
    row = _file(run_id, str(path)) if run_id else None
    if row is None:
        yield
        return
    row.update(status="running", error="", attempts=row["attempts"] + 1)
    # 复试保留累计操作耗时，但覆盖本次扫描的检测结果，避免把同一问题重复计数。
    row.update(issues={}, unconfirmed=0, manual_only=0, exempt_sheets=[], skipped_sheets=[])
    for sheet in row["sheets"].values():
        sheet.pop("checked", None)
        sheet.pop("issues", None)
    try:
        row["size"] = Path(path).stat().st_size
    except OSError:
        row["size"] = None
    _save_file(run_id, row)
    token, started = _current.set(row), time.perf_counter()
    try:
        yield
    except Exception as err:
        row.update(status="error", error=str(err))
        raise
    else:
        row["status"] = "completed"
    finally:
        row["seconds"] += max(0, time.perf_counter() - started)
        _current.reset(token)
        _save_file(run_id, row)


@contextmanager
def operation(key, sheet=""):
    row = _current.get()
    if row is None:
        yield
        return
    started = time.perf_counter()
    try:
        yield
    finally:
        elapsed = max(0, time.perf_counter() - started)
        for target in (row, row["sheets"].setdefault(sheet, {"operations": {}})) if sheet else (row,):
            entry = target["operations"].setdefault(key, {"seconds": 0, "calls": 0})
            entry["seconds"] += elapsed
            entry["calls"] += 1


@contextmanager
def recovery(job, path, *, enabled=True):
    """重启发生在失败尝试之外，也计入该文件，后续自动试修不使用此计时。"""
    run_id = job.get("analysis_id") if enabled else None
    row = _file(run_id, str(path)) if run_id else None
    if row is None:
        yield
        return
    token, started = _current.set(row), time.perf_counter()
    try:
        with operation("restart"):
            yield
    finally:
        row["seconds"] += max(0, time.perf_counter() - started)
        _current.reset(token)
        _save_file(run_id, row)


def note(**values):
    row = _current.get()
    if row is not None:
        row.update(values)


def count(key, amount=1):
    row = _current.get()
    if row is not None:
        row[key] = row.get(key, 0) + amount


def detected(sheet, items):
    row = _current.get()
    if row is None:
        return
    issues = Counter()
    unconfirmed = manual_only = 0
    for item in items.values():
        kind = item.get("kind", "unknown")
        label = {"clip": "边框压线", "hidden": "文字截断或隐藏", "pagebreak": "分页切开合并格"}.get(kind, kind)
        reason = item.get("hidden_reason", "")
        if "###" in reason:
            label = "### 列宽不足"
        if "未匹配" in reason:
            label, manual_only = "PDF 未匹配到文字", manual_only + 1
        unconfirmed += item.get("location_confirmed") is False
        issues[label] += 1
    row["issues"] = dict(Counter(row["issues"]) + issues)
    row["unconfirmed"] += unconfirmed
    row["manual_only"] += manual_only
    row["sheets"].setdefault(sheet, {"operations": {}}).update(issues=dict(issues), checked=True)


def cached(job, path):
    row = _file(job.get("analysis_id"), str(path)) if job.get("analysis_id") else None
    if row:
        row["status"] = "cached"
        _save_file(job["analysis_id"], row)


def timeout(job, path, sheet, limit):
    row = _file(job.get("analysis_id"), str(path)) if job.get("analysis_id") else None
    if row:
        row.update(status="timeout", stuck_sheet=sheet, limit=limit, error=f"扫描超过 {limit / 60:g} 分钟")
        _save_file(job["analysis_id"], row)


def finish(job, *, interrupted=False):
    if not job or not job.get("analysis_id"):
        return
    with _connect() as db:
        found = db.execute("SELECT payload FROM scan_analysis_runs WHERE id=?", (job["analysis_id"],)).fetchone()
    if found:
        run = json.loads(found[0])
        if run["status"] == "running":
            run.update(status="interrupted" if interrupted else "completed", finished_at=_now())
            _save_run(run)


def report(run_id=None, folder=None):
    if not run_id:
        runs = history(folder)
        if not runs:
            return None
        run_id = runs[0]["id"]
    with _connect() as db:
        found = db.execute("SELECT payload FROM scan_analysis_runs WHERE id=?", (run_id,)).fetchone()
        rows = db.execute("SELECT payload FROM scan_analysis_files WHERE run_id=?", (run_id,)).fetchall()
    if not found:
        return None
    run = json.loads(found[0])
    files = [json.loads(row[0]) for row in rows]
    summary = dict(Counter(row["status"] for row in files))
    seconds = sum(row["seconds"] for row in files)
    operations, systems, issues, sheets = {}, {}, Counter(), []
    for row in files:
        group = systems.setdefault(row["system"], {"name": row["system"], "files": 0, "seconds": 0, "issues": 0, "incomplete": 0})
        group["files"] += 1
        group["seconds"] += row["seconds"]
        group["incomplete"] += row["status"] not in {"completed", "cached"}
        # 失败尝试的局部检出不充当完整扫描结果。
        if row["status"] == "completed":
            issues.update(row["issues"])
            group["issues"] += sum(row["issues"].values())
        for key, value in row["operations"].items():
            op = operations.setdefault(key, {"key": key, "name": OPERATIONS.get(key, key), "seconds": 0, "calls": 0})
            op["seconds"] += value["seconds"]
            op["calls"] += value["calls"]
        for name, value in row["sheets"].items():
            sheets.append({"file": row["name"], "sheet": name, "seconds": sum(op["seconds"] for op in value["operations"].values()),
                           "issues": sum(value.get("issues", {}).values()), "complete": row["status"] == "completed" and value.get("checked", False)})
    other = max(0, seconds - sum(op["seconds"] for op in operations.values()))
    operations["other"] = {"key": "other", "name": "其他处理与记录", "seconds": other, "calls": None}
    for op in operations.values():
        op["share"] = op["seconds"] / seconds if seconds else 0
    completed = [row for row in files if row["status"] == "completed"]
    summary.update(total=len(files), seconds=seconds, issues=sum(issues.values()),
                   unconfirmed=sum(row["unconfirmed"] for row in completed),
                   manual_only=sum(row["manual_only"] for row in completed),
                   exempt_sheets=sum(len(row["exempt_sheets"]) for row in files),
                   checked_sheets=sum(bool(s.get("checked")) for row in completed for s in row["sheets"].values()),
                   retries=sum(max(0, row["attempts"] - 1) for row in files))
    findings = []
    incomplete = sum(row["status"] not in {"completed", "cached"} for row in files)
    if incomplete:
        findings.append({"level": "high", "title": f"{incomplete} 个文件未完成完整扫描",
                         "evidence": f"超时 {summary.get('timeout', 0)} 个，失败 {summary.get('error', 0)} 个，待处理或中断 {incomplete - summary.get('timeout', 0) - summary.get('error', 0)} 个。",
                         "action": "先检查下方文件明细中的报错与卡住的 Sheet，修正文件占用、打印区域或 Excel 弹窗后再扫描；这些文件不能视为通过。"})
    if summary["manual_only"] or summary["unconfirmed"]:
        findings.append({"level": "high", "title": "部分问题需要人工核对",
                         "evidence": f"PDF 未匹配文字 {summary['manual_only']} 处，定位未确认 {summary['unconfirmed']} 处（可能重叠）。",
                         "action": "优先核对这些问题的 Excel 原文与打印结果；评估文字定位和匹配规则，避免直接按猜测位置调整。"})
    limited = sum(row.get("scan_scope") == "sheets" for row in completed)
    missing = sum(len(row.get("missing_sheets", [])) for row in files)
    unsupported = sum(row.get("hidden_text_supported") is False for row in files)
    if missing or unsupported:
        findings.append({"level": "high", "title": "部分检查范围存在缺口",
                         "evidence": f"找不到的目标 Sheet {missing} 张；检测库不支持隐藏截断的文件 {unsupported} 个。",
                         "action": "核实工作表改名、删除或隐藏情况；如检测库过旧，升级后再做完整扫描。"})
    if limited or summary.get("cached"):
        findings.append({"level": "info", "title": "本次包含增量检查或缓存复用",
                         "evidence": f"{limited} 个文件仅检查问题 Sheet，{summary.get('cached', 0)} 个文件沿用旧结果。",
                         "action": "这些项目不代表本次重新检查了全部内容；需要评估全量覆盖时，执行完整重新扫描。"})
    if seconds >= 1 and other / seconds >= .2:
        findings.append({"level": "medium", "title": "尚未细分的处理时间占比较高",
                         "evidence": f"已计时操作之外还剩 {other:.1f} 秒，占累计文件处理时间 {other / seconds:.0%}。",
                         "action": "继续拆分结果合并、数据库写入、日志和资源清理的计时；现有数据不足以确定这部分的具体瓶颈。"})
    for op in sorted(operations.values(), key=lambda x: x["seconds"], reverse=True):
        if op["key"] != "other" and op["share"] >= .35 and op["seconds"] >= 1:
            action = {"export": "优先排查慢 Sheet 的打印区域、分页数量和复杂图片，检查能否复用同一版本的 PDF。",
                      "layout": "检查打印方向记录能否复用，优先排查首次方向比较涉及的慢 Sheet。",
                      "detect": "优先分析慢 Sheet 的文字量、合并单元格和定位步骤，评估是否有重复检测。",
                      "open": "优先检查慢文件大小、外部链接、网络目录与 Excel 打开时的交互提示。",
                      "save_view": "优先检查保存视图涉及的文件体积与磁盘耗时，评估重复打开和保存是否可减少。",
                      "highlight": "优先检查问题格数量和条件格式操作，评估能否按范围批量标色。",
                      "restart": "结合重试错误检查 Excel 进程退出、启动和文件占用，优先减少反复重启。"}[op["key"]]
            findings.append({"level": "medium", "title": f"优先排查：{op['name']}",
                             "evidence": f"累计 {op['seconds']:.1f} 秒，占文件处理耗时 {op['share']:.0%}；这是计时观察，尚不能单独确定根因。", "action": action})
    if summary["retries"]:
        findings.append({"level": "medium", "title": "重试增加了扫描成本", "evidence": f"额外尝试 {summary['retries']} 次，耗时已累计到对应文件。",
                         "action": "优先处理反复出错的文件，结合错误信息检查 Excel 稳定性和文件占用。"})
    if not findings:
        findings.append({"level": "info", "title": "当前计时未显示明显集中瓶颈", "evidence": "请结合下方慢文件和操作明细判断。未计时的旧扫描不提供推算数据。",
                         "action": "在同一文件夹再次扫描，对比耗时和覆盖情况，观察是否有稳定重复出现的慢文件。"})
    return {"run": run, "summary": summary, "findings": findings, "issues": dict(issues),
            "operations": sorted(operations.values(), key=lambda x: x["seconds"], reverse=True),
            "systems": sorted(systems.values(), key=lambda x: x["seconds"], reverse=True),
            "files": sorted(files, key=lambda x: x["seconds"], reverse=True),
            "sheets": sorted(sheets, key=lambda x: x["seconds"], reverse=True)}


def markdown(data):
    def cell(value):
        return str(value).replace("|", "\\|").replace("\r", " ").replace("\n", " ")
    run, summary = data["run"], data["summary"]
    lines = ["# 扫描分析报告", "", f"扫描目录：{run['folder']}", f"开始时间：{run['started_at']}",
             f"状态：{run['status']}", "", f"文件 {summary['total']} 个；完成本次检查 {summary.get('completed', 0)} 个；"
             f"沿用缓存 {summary.get('cached', 0)} 个；累计文件处理 {summary['seconds']:.1f} 秒；检出 {summary['issues']} 处。",
             "", "统计扫描阶段，含失败尝试；不含暂停等待和后续自动试修。系统按第一级子文件夹归类。", "", "## 薄弱点与优化建议", ""]
    for finding in data["findings"]:
        lines += [f"### {finding['title']}", finding["evidence"], finding["action"], ""]
    lines += ["## 操作耗时", "", "| 操作 | 秒 | 占比 |", "|---|---:|---:|"]
    lines += [f"| {op['name']} | {op['seconds']:.1f} | {op['share']:.1%} |" for op in data["operations"]]
    lines += ["", "## 系统耗时", "", "| 系统 | 文件 | 秒 | 未完成 |", "|---|---:|---:|---:|"]
    lines += [f"| {cell(s['name'])} | {s['files']} | {s['seconds']:.1f} | {s['incomplete']} |" for s in data["systems"]]
    lines += ["", "## 文件耗时", "", "| 文件 | 状态 | 秒 | 尝试 | 卡住的表 / 错误 |", "|---|---|---:|---:|---|"]
    lines += [f"| {cell(row['name'])} | {row['status']} | {row['seconds']:.1f} | {row['attempts']} | "
              f"{cell(row.get('stuck_sheet', '') + ' ' + row.get('error', ''))} |" for row in data["files"]]
    return "\n".join(lines) + "\n"
