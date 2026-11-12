"""Excel 对比的「真实效果」（用户 2026-11-12）：两份 Excel 用 Excel 真导一次 PDF，网页上两边的页面图逐像素比，
不一样的地方左边盖红、右边盖绿（像素比在浏览器里做，static/excel_diff_real.js）。

- 只导逐格对比说有不同的表（和只有一边有的表），精简注定会删的表、隐藏表不导。
- 印法和上传的同名 PDF 一样（core/pdf_bundle._export_sheet：强制宽一页、藏空心大框、截图和红框临时组合），
  导之前在内存里去掉本工具的标记色；页眉页脚照 Excel 原样。
- 先把 Excel 复制到临时文件夹再只读打开，原文件不碰、不保存。
- 结果按文件内容的 SHA-256 存在程序数据文件夹的 excel_diff_pdf/<哈希>/（每张表一份 PDF + meta.json），
  文件没变再看不用重导；只留最近 MAX_KEEP 份。
- 一个后台线程带一个自己的隐藏 Excel，排队一份一份导；不占任务（别的任务在跑也能看），空闲 IDLE_QUIT 秒退出 Excel。
"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import tempfile
import threading
import time
from pathlib import Path
from typing import Any

from core import debug_log

IDLE_QUIT = 90       # 秒：没活干多久后关掉后台 Excel
MAX_KEEP = 400       # 缓存最多留多少份 Excel 的导出结果
PAGE_ZOOM = 1.5      # 页面图放大倍数（两边一样，像素才对得上）
_SHA = re.compile(r"^[0-9a-f]{64}$")

_LOCK = threading.Condition()
_QUEUE: list[str] = []                      # 排队的哈希
_JOBS: dict[str, dict[str, Any]] = {}       # 哈希 → {path, sheets: set, state, error, current}
_HASHES: dict[tuple[str, int, int], str] = {}  # (路径, 大小, 修改时间) → 哈希，免得每次都重算
_worker: threading.Thread | None = None


def _root() -> Path:
    from core.data_dir import DATA_DIR
    return Path(DATA_DIR) / "excel_diff_pdf"


def _digest(path: Path) -> str:
    stat = path.stat()
    key = (str(path), stat.st_size, int(stat.st_mtime_ns))
    with _LOCK:
        cached = _HASHES.get(key)
    if cached is not None:
        return cached
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for block in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(block)
    digest = h.hexdigest()
    with _LOCK:
        _HASHES[key] = digest
        if len(_HASHES) > 2 * MAX_KEEP:
            _HASHES.pop(next(iter(_HASHES)))
    return digest


def _meta_path(sha: str) -> Path:
    return _root() / sha / "meta.json"


def _load_meta(sha: str) -> dict[str, Any]:
    try:
        data = json.loads(_meta_path(sha).read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save_meta(sha: str, meta: dict[str, Any]) -> None:
    path = _meta_path(sha)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(meta, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


def _prune() -> None:
    root = _root()
    try:
        dirs = sorted((d for d in root.iterdir() if d.is_dir()), key=lambda d: d.stat().st_mtime)
    except OSError:
        return
    for old in dirs[:-MAX_KEEP]:
        shutil.rmtree(old, ignore_errors=True)


def _sheet_view(sha: str, meta: dict[str, Any], names: list[str]) -> dict[str, Any]:
    sheets = meta.get("sheets", {})
    return {name: {k: v for k, v in sheets[name].items() if k != "file"} for name in names if name in sheets}


def request(path_text: Any, sheets: Any) -> dict[str, Any]:
    """要这份 Excel 里这些表的 PDF：已经导过的直接给，没导过的排队。返回 {key, state, sheets, error}"""
    if not isinstance(path_text, str) or not path_text.strip():
        raise ValueError("需要 Excel 的完整路径")
    names = [n for n in dict.fromkeys(sheets or []) if isinstance(n, str) and n]
    path = Path(path_text.strip())
    if not path.is_file():
        raise ValueError(f"找不到文件：{path}")
    sha = _digest(path)
    with _LOCK:
        meta = _load_meta(sha)
        missing = [n for n in names if n not in meta.get("sheets", {})]
        job = _JOBS.get(sha)
        if missing:
            if job is None or job["state"] in ("done", "error"):
                job = _JOBS[sha] = {"path": str(path), "sheets": set(missing), "state": "queued", "error": "", "current": ""}
                _QUEUE.append(sha)
                debug_log.event("📄 对比用 PDF 排队", 文件=path.name, 表=missing)
            else:
                job["sheets"].update(missing)
            _ensure_worker()
            _LOCK.notify_all()
        return _status(sha, names)


def status(keys: Any) -> dict[str, Any]:
    """keys = {哈希: [表名…]}"""
    if not isinstance(keys, dict):
        raise ValueError("keys 格式不对")
    with _LOCK:
        return {sha: _status(sha, names if isinstance(names, list) else []) for sha, names in keys.items() if _SHA.match(str(sha))}


def _status(sha: str, names: list[str]) -> dict[str, Any]:
    meta = _load_meta(sha)
    job = _JOBS.get(sha) or {}
    have = _sheet_view(sha, meta, names)
    state = "done" if len(have) == len(names) else job.get("state")
    error = job.get("error", "") if state == "error" else ""
    if len(have) != len(names) and state in (None, "done"):
        state, error = "error", "部分表的 PDF 已失效，请重新对比"
    position = _QUEUE.index(sha) + 1 if sha in _QUEUE else 0
    return {"key": sha, "state": state, "sheets": have, "error": error,
            "current": job.get("current", ""), "queue": position, "zoom": PAGE_ZOOM}


def page_png(sha: Any, sheet: Any, page: Any) -> bytes:
    """导好的某张表第 page 页（从 0 数）的图"""
    from core.pdf_detector import fitz
    if not isinstance(sha, str) or not _SHA.match(sha):
        raise ValueError("key 不对")
    entry = _load_meta(sha).get("sheets", {}).get(sheet if isinstance(sheet, str) else "")
    if not entry or not entry.get("file"):
        raise ValueError("这张表还没导出 PDF")
    try:
        index = int(page)
    except (TypeError, ValueError):
        raise ValueError("页码不对") from None
    with fitz.open(_root() / sha / entry["file"]) as doc:
        if not 0 <= index < len(doc):
            raise ValueError("页码不对")
        return doc[index].get_pixmap(matrix=fitz.Matrix(PAGE_ZOOM, PAGE_ZOOM), alpha=False).tobytes("png")


# ---------- 后台线程 ----------

def _ensure_worker() -> None:
    global _worker
    if _worker is None or not _worker.is_alive():
        _worker = threading.Thread(target=_run, name="Excel 对比导 PDF", daemon=True)
        _worker.start()


def _trim_jobs() -> None:
    finished = [sha for sha, job in _JOBS.items() if job["state"] in ("done", "error")]
    for sha in finished[:-MAX_KEEP]:
        _JOBS.pop(sha)


def _run() -> None:
    from core import workflow as wf
    from core.excel_ops import quit_background_excel, start_background_excel
    from core.excel_watchdog import ExcelWatchdog
    excel, dog = None, None

    def drop() -> None:  # 关掉后台 Excel 和它的看门狗
        nonlocal excel, dog
        try:
            quit_background_excel(excel)
        except Exception:
            pass
        if dog is not None:
            dog.stop()
        excel, dog = None, None
    wf.pythoncom.CoInitialize()
    try:
        while True:
            with _LOCK:
                if not _QUEUE:
                    _LOCK.wait(IDLE_QUIT)
                idle = not _QUEUE and excel is not None
            if idle:  # 空闲太久：关掉 Excel，下次有活再开（线程留着）
                drop()
                debug_log.event("📄 对比用 PDF：空闲，关掉后台 Excel")
            with _LOCK:
                if not _QUEUE:
                    continue
                sha = _QUEUE.pop(0)
                job = _JOBS[sha]
                job["state"] = "running"
                names = sorted(job["sheets"])
            started = time.perf_counter()
            try:
                if excel is None:
                    # 弹框没人点、或 5 分钟没反应就结束这个 Excel；不进任务进度框（listed=False）
                    dog = ExcelWatchdog(lambda: None, timeout=300)
                    dog.start(listed=False)
                    excel = start_background_excel()
                    dog.register(excel)
                dog.progress()
                _export(excel, sha, Path(job["path"]), names, job, dog)
                with _LOCK:
                    job["sheets"].difference_update(names)
                    if job["sheets"]:
                        job.update(state="queued", current="")
                        _QUEUE.append(sha)
                    else:
                        job.update(state="done", current="")
                    _trim_jobs()
                debug_log.event("📄 对比用 PDF 导好了", 文件=Path(job["path"]).name, 表=names,
                                秒=round(time.perf_counter() - started, 2))
            except Exception as err:
                debug_log.exception(err, f"对比用 PDF {job['path']}")
                reason = dog.explain(err) if dog is not None else str(err)
                with _LOCK:
                    job.update(state="error", error=reason, current="")
                    _trim_jobs()
                drop()  # Excel 可能卡住了，下一份重开
    finally:
        drop()
        wf.pythoncom.CoUninitialize()


def _export(excel: Any, sha: str, path: Path, names: list[str], job: dict[str, Any], dog: Any = None) -> None:
    from core import pdf_bundle
    from core.common import XL_SHEET_VISIBLE
    from core.excel_ops import open_workbook_checked
    folder = _root() / sha
    folder.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(prefix="excel_diff_pdf_"))
    try:
        copy = tmp / ("book" + path.suffix.lower())
        shutil.copyfile(path, copy)  # 复制一份再开：原文件不碰，别人开着也不怕
        wb = open_workbook_checked(excel, str(copy), ReadOnly=True, UpdateLinks=0, IgnoreReadOnlyRecommended=True)
        try:
            pdf_bundle._drop_marks(wb, path)
            by_name = {str(s.Name): s for s in wb.Sheets}
            results: dict[str, dict[str, Any]] = {}
            for index, name in enumerate(names, 1):
                if dog is not None:
                    dog.progress()
                with _LOCK:
                    job["current"] = f"{index} / {len(names)}：{name}"
                sheet = by_name.get(name)
                if sheet is None:
                    results[name] = {"pages": 0, "error": "Excel 里找不到这张表"}
                    continue
                if int(sheet.Visible) != XL_SHEET_VISIBLE:
                    results[name] = {"pages": 0, "error": "隐藏表，不会印出来"}
                    continue
                target = folder / f"{hashlib.md5(name.encode('utf-8')).hexdigest()}.pdf"
                try:
                    pdf_bundle._export_sheet(sheet, target)
                except Exception as err:
                    if not pdf_bundle.nothing_to_print(err):
                        raise
                    results[name] = {"pages": 0, "error": "没有可打印的内容"}
                    continue
                results[name] = {"pages": _page_count(target), "file": target.name}
        finally:
            wb.Close(SaveChanges=False)
        with _LOCK:
            meta = _load_meta(sha)
            meta.setdefault("sheets", {}).update(results)
            meta["source"] = str(path)
            _save_meta(sha, meta)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    _prune()


def _page_count(pdf: Path) -> int:
    from core.pdf_detector import fitz
    with fitz.open(pdf) as doc:
        return doc.page_count
