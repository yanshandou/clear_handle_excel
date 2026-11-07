"""PDF 人工复检：导出的 PDF 一页页（小图）显示在网页上，单击换颜色标记（蓝 / 黄这一页要改 / 红删整张表），
看到问题一点就用 Excel 打开对应的表，跳到 PDF 的那一页（选中这一页打印的范围）。

只精确到「哪张表的第几页」（用户 2026-11-02 定），不再算点的是哪一格：以前为了对回单元格，后台要把每份
Excel 逐行、逐列涂色各导出一次 PDF，很慢，现在不要了。第几页 → Excel 里的范围按打印区域和分页推算
（page_breaks.page_area），页数对不上（PDF 是旧的、Excel 跳过了空白页）时只打开那张表，不硬套。
填了原始备份文件夹时，备份也一起打开（只读）、跳到同一块，左右对照着改。
"""
from __future__ import annotations

import re
import threading
import time
from pathlib import Path
from typing import Any

from core import batch, debug_log, file_problem, page_breaks, page_fill, pdf_store, repair_compare, sheet_history, workflow as wf
from core.backup_compare import pair_files
from core.common import (EXCEL_EXTENSIONS, sanitize_filename, strip_review_prefix,
                         workbook_locked, workbook_rel_name)
from core.excel_ops import open_workbook_at, open_workbook_checked, start_background_excel
from core.formula_links import delete_sheets_preserving_values
from core.pdf_detector import fitz

PAGE_ZOOM = 2.5  # 网页上显示的页面图：A4 竖向约 1490 像素宽，铺满屏幕也看得清字
OLD_LIST_SECONDS = 60  # 备份文件夹里的文件清单记一分钟，连着点不用每次都遍历
_PAGES_CACHE: dict[tuple[str, int, int], list[list[float]]] = {}
_OLD_FILES: dict[str, tuple[float, list[Path]]] = {}


def _exports() -> dict[str, dict]:
    """导出 PDF 时记下的（workflow._export_workbook_pdfs）：PDF 路径 -> {wb_path, sheet}"""
    return wf.STATE.setdefault("pdf_exports", {})


def _in_folder(path: Path) -> bool:
    with wf.STATE_LOCK:
        root = wf.STATE.get("pdf_review", {}).get("folder")
    return bool(root) and path.is_file() and Path(root) in path.parents


def forget_pdfs(paths: list[Path]) -> None:
    """导出前删掉了旧 PDF：它们的导出记录也去掉"""
    gone = {str(p) for p in paths} | {str(Path(p).resolve()) for p in paths}
    with wf.STATE_LOCK:
        for pdf in gone:
            _exports().pop(pdf, None)
        wf._persist_state()


# ---------- 列出 PDF ----------

def _resolve_folder(folder: str) -> Path:
    root = Path(folder.strip().strip('"').strip("'")).expanduser().resolve()
    if not root.is_dir():
        raise ValueError(f"文件夹不存在: {root}")
    return root


def _guess_owner(pdf: Path) -> dict | None:
    """没有导出记录（旧版本导出的、扫描时导出的）：按「Excel 文件名_表名.pdf」找 Excel。
    PDFS 里的 PDF 到对应的 Excel 文件夹找（PDFS/sub/x.pdf → sub/），旧的混放 PDF 在同一个文件夹里找"""
    best: tuple[int, Path, str] | None = None
    folder = pdf_store.excel_folder_for(pdf)
    if not folder.is_dir():
        return None
    for wb in folder.iterdir():
        if not wb.is_file() or wb.suffix.lower() not in EXCEL_EXTENSIONS | {".xlsb"} or wb.name.startswith("~$"):
            continue
        stem = sanitize_filename(strip_review_prefix(wb.stem))
        if pdf.stem.startswith(stem + "_") and len(pdf.stem) > len(stem) + 1:
            if best is None or len(stem) > best[0]:
                best = (len(stem), wb, pdf.stem[len(stem) + 1:])
    return {"wb_path": str(best[1]), "sheet": best[2], "guessed": True} if best else None


def _owner(pdf: Path) -> dict | None:
    record = _exports().get(str(pdf))
    if record and Path(record["wb_path"]).is_file():
        return dict(record)
    return pdf_store.owner(pdf) or _guess_owner(pdf)


def _page_sizes(pdf: Path) -> list[list[float]]:
    stat = pdf.stat()
    key = (str(pdf), stat.st_mtime_ns, stat.st_size)
    if key not in _PAGES_CACHE:
        with fitz.open(pdf) as doc:
            _PAGES_CACHE[key] = [[round(page.rect.width, 2), round(page.rect.height, 2)] for page in doc]
        if len(_PAGES_CACHE) > 2000:
            _PAGES_CACHE.pop(next(iter(_PAGES_CACHE)))
    return _PAGES_CACHE[key]


def _pdf_files(root: Path) -> list[Path]:
    """有 PDFS 文件夹就只看里面的（导出的 PDF 都在这）；没有的话（旧版本导出、和 Excel 混放的）看整个文件夹"""
    base = pdf_store.pdf_root(root)
    where = base if base.is_dir() else root
    return sorted((p for p in where.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"
                   and not p.name.startswith(("~", "."))), key=lambda p: workbook_rel_name(p, where))


def list_pdfs(folder: str) -> dict:
    """文件夹里的 PDF（有 PDFS 文件夹就只看 PDFS 里的），每份带对应的 Excel 和表、各页大小（PDF 点）"""
    root = _resolve_folder(folder)
    pdfs = _pdf_files(root)
    base = pdf_store.pdf_root(root)
    shown_from = base if base.is_dir() else root
    entries = []
    with wf.STATE_LOCK:
        wf.STATE["pdf_review"] = {"folder": str(root)}  # 以前画框标色留下的记录不要了
        for gone in [p for p in _exports() if Path(p).parent == root or root in Path(p).parents]:
            if not Path(gone).is_file():
                _exports().pop(gone)  # PDF 已经删掉的记录不留着
        owners = {str(p): _owner(p) for p in pdfs}
    notes = pdf_store.notes_map(root) if base.is_dir() else {}
    for pdf in pdfs:
        entry: dict[str, Any] = {"pdf": str(pdf), "name": workbook_rel_name(pdf, shown_from), "v": pdf.stat().st_mtime_ns}
        try:
            entry["pages"] = _page_sizes(pdf)
        except Exception as err:
            entry.update(pages=[], error=f"PDF 打不开：{err}")
        owner = owners[str(pdf)]
        if owner:
            entry.update(wb_path=owner["wb_path"], wb_name=workbook_rel_name(Path(owner["wb_path"]), root),
                         sheet=owner["sheet"], guessed=bool(owner.get("guessed")))
        if notes.get(str(pdf)):
            entry["notes"] = notes[str(pdf)]  # 导出时发现的内容很少的页（page_fill）
        entries.append(entry)
    focus = None
    from core import work_task
    try:
        hit = work_task.focus_filter(root)
    except Exception:
        hit = None
    if hit:  # 本轮只处理待重新处理文件：只看这几份 Excel 的 PDF
        _task, keys, prefix = hit
        lead = prefix + "/" if prefix else ""
        before = len(entries)
        entries = [e for e in entries if e.get("wb_name") and work_task._focus_key(lead + e["wb_name"]) in keys]
        focus = {"files": len(keys), "hidden": before - len(entries)}
    debug_log.event("🔎 PDF 复检列出", 文件夹=root, 看的是=shown_from, PDF=len(entries), 找到Excel=sum(1 for e in entries if e.get("wb_path")),
                    本轮清单=bool(focus))
    return {"folder": str(root), "pdf_folder": str(shown_from), "in_pdfs": shown_from == base, "pdfs": entries, "focus": focus}


def page_png(pdf: str, page: int) -> bytes:
    path = Path(pdf).resolve()
    if not _in_folder(path) or path.suffix.lower() != ".pdf":
        raise ValueError("只能查看当前文件夹里的 PDF")
    with fitz.open(path) as doc:
        if not 0 <= page < len(doc):
            raise ValueError("页码不对")
        return doc[page].get_pixmap(matrix=fitz.Matrix(PAGE_ZOOM, PAGE_ZOOM), alpha=False).tobytes("png")


# ---------- 打开 Excel 跳到 PDF 的那一页 ----------

def _find_sheet(book: Any, name: str) -> Any:
    """按表名找；按文件名猜出来的表名是去掉了非法字符的，再按同样规则比一次"""
    sheet = batch._find_sheet(book, name)
    if sheet is not None:
        return sheet
    for candidate in (name, re.sub(r"_\d+$", "", name)):
        for sheet in book.Worksheets:
            if sanitize_filename(str(sheet.Name)) == candidate:
                return sheet
    return None


def _old_files(old_root: Path) -> list[Path]:
    key = str(old_root).casefold()
    cached = _OLD_FILES.get(key)
    if cached and time.monotonic() - cached[0] < OLD_LIST_SECONDS:
        return cached[1]
    files = [p for p in old_root.rglob("*") if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS | {".xlsb"}
             and not p.name.startswith("~$")]
    _OLD_FILES[key] = (time.monotonic(), files)
    return files


def backup_for(wb_path: Path, old_folder: str) -> tuple[Path | None, str]:
    """在原始备份文件夹里找这份 Excel 的老版本：同样的相对路径，或者文件名在两边都只有一个。
    找不到返回 (None, 原因)"""
    old_root = Path(old_folder.strip().strip('"').strip("'")).expanduser().resolve()
    if not old_root.is_dir():
        return None, f"原始备份文件夹不存在：{old_root}"
    with wf.STATE_LOCK:
        root = Path(wf.STATE.get("pdf_review", {}).get("folder") or wb_path.parent)
    name = workbook_rel_name(wb_path, root)
    found = pair_files(old_root, _old_files(old_root), [name], root).get(name)
    if found is None:
        return None, "原始备份文件夹里找不到这份文件（或有几份同名的分不清）"
    if found.resolve() == wb_path.resolve():
        return None, "原始备份和当前文件是同一个文件"
    return found.resolve(), ""


def _check_open_allowed(wb_path: Path) -> None:
    with wf.STATE_LOCK:
        job = wf.STATE.get("job") or {}
        if wf.STATE.get("busy") and (job.get("kind", "") in wf.QUICK_TASK_NAMES or job.get("kind", "").startswith("batch_")):
            reason = batch.open_block_reason(job, wb_path)
            if reason:
                raise ValueError(reason)


def open_page(wb_path: Path, sheet: str, page: int | None, pages: int) -> tuple[str, str, str]:
    """在可见 Excel 打开文件、切到这张表、选中 PDF 第 page 页打印的那一块。返回 (表名, 区域, 提示)"""
    def name_of(book: Any) -> str:
        found = _find_sheet(book, sheet)
        return str(found.Name) if found is not None else sheet

    def locate(target: Any) -> str | None:
        return None if page is None else page_breaks.page_area(target, page, pages)

    name, address, note = open_workbook_at(wb_path, name_of, locate)
    if page is not None and not address and not note:
        note = f"第 {page + 1} 页在 Excel 里的位置没对上（PDF 可能是旧的，可以先「重新生成」），只打开了这张表"
    return name, address, note


def open_at(pdf: str, page: Any, pages: Any = None, old_folder: Any = "") -> dict:
    """在 Excel 打开 PDF 这一页对应的表和范围；有原始备份时备份也打开（只读）跳到同一块"""
    path = Path(pdf).resolve()
    if not _in_folder(path) or path.suffix.lower() != ".pdf":
        raise ValueError("只能查看当前文件夹里的 PDF，请先重新加载")
    if isinstance(page, bool) or not isinstance(page, int) or page < 0:
        raise ValueError("页码不对")
    with wf.STATE_LOCK:
        owner = _owner(path)
    if not owner:
        raise ValueError("找不到这份 PDF 对应的 Excel 文件")
    if isinstance(pages, bool) or not isinstance(pages, int) or pages <= page:
        pages = len(_page_sizes(path))
    wb_path = Path(owner["wb_path"])
    _check_open_allowed(wb_path)
    sheet, address, note = open_page(wb_path, owner["sheet"], page, pages)
    notes = [note] if note else []
    backup, why = backup_for(wb_path, old_folder) if isinstance(old_folder, str) and old_folder.strip() else (None, "")
    if why:
        notes.append(why)
    if backup:
        result = repair_compare.open_comparison({"current": wb_path.resolve(strict=True), "backup": backup,
                                                 "sheet": sheet, "old_sheet": sheet, "cell": address})
        if result.get("note"):
            notes.append(result["note"])
    debug_log.event("🔎 PDF 复检打开", PDF=path.name, 页=page + 1, 共=pages, 表=sheet, 范围=address or "-",
                    备份=backup.name if backup else "-", 说明="；".join(notes) or "-")
    return {"sheet": sheet, "cell": address, "backup": bool(backup), "note": "；".join(notes)}


# ---------- 三色标记：一键处理 ----------
# 网页上每份 PDF（= 一张表）可以标成：蓝（默认，不用管）/ 黄（要人工改）/ 红（这张表不要了）。
# 网页上黄红分开两个按钮（2026-11-02）：「删除标红」只传 red，「打开标黄」只传 yellow；两个都传时先删后开。
# 删除：在后台 Excel 里把标红的表删掉并保存（同一份 Excel 的几张一起删、只开一次），
# 删掉的表对应的 PDF 也删掉；再把标黄的用 Excel 打开（每份 Excel 开一次，跳到标黄时右键的那一格），
# 一次最多开 OPEN_LIMIT 份 Excel（按文件算，同一份 Excel 几张标黄的表只算 1 份），剩下的再点一次接着开。

OPEN_LIMIT = 10
_PROCESS: dict[str, Any] = {"running": False}
_PROCESS_LOCK = threading.Lock()


def process_status() -> dict:
    with _PROCESS_LOCK:
        return {**_PROCESS, "result": _PROCESS.get("result")}


def _process_item(raw: Any, color: str) -> dict:
    if not isinstance(raw, dict) or not isinstance(raw.get("pdf"), str):
        raise ValueError("标记的数据不对，请刷新页面再试")
    path = Path(raw["pdf"]).resolve()
    if not _in_folder(path) or path.suffix.lower() != ".pdf":
        raise ValueError(f"「{path.name}」不在当前文件夹里或已经不在了，请重新加载 PDF")
    with wf.STATE_LOCK:
        owner = _owner(path)
    if not owner:
        raise ValueError(f"「{path.name}」找不到对应的 Excel 文件")
    item = {"pdf": str(path), "name": path.name, "color": color, "wb_path": owner["wb_path"], "sheet": owner["sheet"],
            "spots": []}
    # 标黄的页（只记第几页）
    for spot in raw.get("spots") or []:
        page = spot.get("page") if isinstance(spot, dict) else None
        if isinstance(page, bool) or not isinstance(page, int) or page < 0:
            continue
        item["spots"].append({"page": page})
    item["spots"].sort(key=lambda spot: spot["page"])
    return item


def process(red: Any, yellow: Any, floor: Any = None) -> None:
    """一键处理：后台删掉标红的表，按工具箱标记改打印设置（向下取整），再打开标黄的。
    网页轮询 process_status() 看进度和结果"""
    reds = [_process_item(raw, "red") for raw in (red or [])]
    yellows = [_process_item(raw, "yellow") for raw in (yellow or [])]
    floors = [_process_item(raw, "floor") for raw in (floor or [])]
    if not reds and not yellows and not floors:
        raise ValueError("没有标黄、标红或标了工具的表")
    from core import bl_links, hidden_repair, link_repair, mark_finder
    with wf.STATE_LOCK, _PROCESS_LOCK:
        if (wf.STATE.get("busy") or _PROCESS.get("running") or link_repair.status().get("running")
                or mark_finder.status().get("running") or hidden_repair.status().get("running")
                or bl_links.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束再处理")
        if reds or floors:
            wf.STATE["busy"] = True  # 要改文件：别的任务这时候不能开
            wf.STATE["maintenance"] = "pdf_review_process"
        books = list(dict.fromkeys(item["wb_path"] for item in reds))
        _PROCESS.clear()
        _PROCESS.update(running=True, stage="delete" if reds else "floor" if floors else "open", total=len(books),
                        done=0, current="", result=None)
    root = wf.STATE.get("pdf_review", {}).get("folder") or ""

    def run() -> None:
        result: dict[str, list] = {"deleted": [], "failed": [], "opened": [], "open_failed": [], "later": [],
                                   "floored": [], "floor_failed": []}
        try:
            with debug_log.side_task("pdf_review_process", {"folder": root, "red": len(reds), "yellow": len(yellows),
                                                             "floor": len(floors)}):
                for item in reds + yellows + floors:
                    debug_log.event("🎨 PDF 复检标记",
                                    颜色={"red": "红", "yellow": "黄", "floor": "向下取整"}[item["color"]],
                                    PDF=item["name"], 表=item["sheet"],
                                    页=[spot["page"] + 1 for spot in item["spots"]] or "-")
                if reds:
                    _delete_red(reds, root, result)
                if floors:
                    _floor_marked(floors, root, result)
                if yellows:
                    with _PROCESS_LOCK:
                        _PROCESS.update(stage="open", current="")
                    with wf.STATE_LOCK:
                        wf.STATE["busy"] = False  # 删完了；打开 Excel 不用占着任务槽
                        wf.STATE.pop("maintenance", None)
                    _open_yellow(yellows, root, result)
                debug_log.event("✅ PDF 复检一键处理完成", 删掉=len(result["deleted"]), 没删掉=len(result["failed"]),
                                向下取整=len(result["floored"]), 没取整=len(result["floor_failed"]),
                                打开=len(result["opened"]), 没打开=len(result["open_failed"]), 下次开=len(result["later"]))
        except Exception as err:
            result["failed"].append({"name": "任务", "sheet": "", "reason": f"处理出错：{err}"})
        finally:
            with wf.STATE_LOCK:
                if wf.STATE.get("maintenance") == "pdf_review_process":
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)
            with _PROCESS_LOCK:
                _PROCESS.update(running=False, current="", result=result)

    try:
        threading.Thread(target=run, name="PDF 复检一键处理", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _PROCESS_LOCK:
            if wf.STATE.get("maintenance") == "pdf_review_process":
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)
            _PROCESS["running"] = False
        raise


def _delete_red(reds: list[dict], root: str, result: dict) -> None:
    groups: dict[str, list[dict]] = {}
    for item in reds:
        groups.setdefault(item["wb_path"], []).append(item)
    guard = file_problem.Guard("PDF 复检：删除标红的表")
    excel = None
    wf.pythoncom.CoInitialize()
    try:
        for index, (wb_path, items) in enumerate(groups.items()):
            label = workbook_rel_name(Path(wb_path), Path(root)) if root else Path(wb_path).name
            with _PROCESS_LOCK:
                _PROCESS.update(done=index, current=label)
            while True:
                guard.progress()
                try:
                    if excel is None:
                        excel = guard.watch(start_background_excel())
                    outcome = _delete_in_workbook(excel, Path(wb_path), items, label)
                    break
                except Exception as err:
                    debug_log.exception(err, f"PDF 复检删表 {label}")
                    retry, message = guard.problem(err, wb_path, label, root,
                                                   close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {label}")
                        continue
                    outcome = [{**item, "ok": False, "reason": message} for item in items]
                    break
            for item in outcome:
                entry = {"pdf": item["pdf"], "name": label, "sheet": item["sheet"], "reason": item.get("reason", "")}
                if item["ok"]:
                    entry["pdf_removed"] = item.get("pdf_removed", False)
                result["deleted" if item["ok"] else "failed"].append(entry)
            with _PROCESS_LOCK:
                _PROCESS.update(done=index + 1)
    finally:
        try:
            file_problem.quit_quietly(excel)
        finally:
            guard.stop()
            wf.pythoncom.CoUninitialize()


def _delete_in_workbook(excel: Any, wb_path: Path, items: list[dict], label: str) -> list[dict]:
    """在后台 Excel 里删掉这份文件里标红的表并保存；删表前把别的表引用它的公式转成值（和精简 Sheet 一样）。
    返回每一项：ok、reason；删成功的 PDF 也删掉"""
    if not wb_path.is_file():
        return [{**item, "ok": False, "reason": "Excel 文件已经不在了"} for item in items]
    if workbook_locked(wb_path):
        raise ValueError("文件正在 Excel 中打开，请先在 Excel 里关掉它（改过的先保存），再点「重试」")
    before_hash = sheet_history.fingerprint(wb_path)
    book = open_workbook_checked(excel, str(wb_path.resolve()), UpdateLinks=0)
    outcome: list[dict] = []
    deleting: list[str] = []
    try:
        before = [str(sheet.Name) for sheet in book.Sheets]
        for item in items:
            sheet = _find_sheet(book, item["sheet"])
            if sheet is None:
                outcome.append({**item, "ok": False, "reason": "文件里已经没有这张表（可能改名或删掉了）"})
            elif str(sheet.Name) in deleting:
                outcome.append({**item, "ok": True, "same": True})
            else:
                deleting.append(str(sheet.Name))
                outcome.append({**item, "ok": True, "real": str(sheet.Name)})
        visible = [str(sheet.Name) for sheet in book.Sheets if int(sheet.Visible) == -1]
        if deleting and visible and all(name in deleting for name in visible):
            reason = "这是文件里剩下的唯一一张显示着的表，Excel 不允许删掉（整个文件不要的话请直接删文件）"
            debug_log.event("⚠️ PDF 复检删表跳过", 文件=label, 表=deleting, 原因=reason)
            return [{**item, "ok": False, "reason": reason} if item["ok"] else item for item in outcome]
        if not deleting:
            return outcome
        kept_values = delete_sheets_preserving_values(book, deleting)
        book.Save()
        after = [str(sheet.Name) for sheet in book.Sheets]
    finally:
        book.Close(SaveChanges=False)
    with wf.STATE_LOCK:
        sheet_history.record(wb_path, before, after, deleting, {}, kept_values, before_hash)
        wf._persist_state()
    debug_log.event("🗑️ PDF 复检删表", 文件=label, 删掉=deleting, 保留引用值=len(kept_values))
    wf.log_msg(f"🗑️ {label}：删掉了 {'、'.join(deleting)}" + (f"（另外保留引用值 {len(kept_values)} 处）" if kept_values else ""))
    removed: list[Path] = []
    for item in outcome:
        if not item["ok"]:
            continue
        pdf = Path(item["pdf"])
        try:
            pdf.unlink(missing_ok=True)
            removed.append(pdf)
            item["pdf_removed"] = True
        except OSError as err:
            item["reason"] = f"表删掉了，但 PDF 删不掉（可能正在用别的程序打开）：{err}"
            debug_log.event("⚠️ PDF 复检删 PDF 失败", PDF=pdf.name, 原因=err)
    if removed:
        forget_pdfs(removed)
    if kept_values:
        for item in outcome:
            if item["ok"] and not item.get("reason"):
                item["reason"] = f"别的表引用它的 {len(kept_values)} 处公式已转成值"
    return outcome


def _floor_marked(floors: list[dict], root: str, result: dict) -> None:
    """工具箱「向下取整」：每份 Excel 打开一次，标了的表设成 1 页宽、N-1 页高，真导出核对过才保存；
    保存后把这份 Excel 的 PDF 全部重新导出（和「🔄 重新生成」一样），复检页上看到的就是新的"""
    from core import reporter
    groups: dict[str, list[dict]] = {}
    for item in floors:
        groups.setdefault(item["wb_path"], []).append(item)
    with _PROCESS_LOCK:
        _PROCESS.update(stage="floor", total=len(groups), done=0, current="")
    guard = file_problem.Guard("PDF 复检：向下取整")
    excel = None
    reporter.suspend_updates(True)  # 不是主页的批量任务：不写进执行记录
    wf.pythoncom.CoInitialize()

    def get_excel() -> Any:
        nonlocal excel
        if excel is None:
            excel = guard.watch(start_background_excel())
        return excel

    try:
        for index, (wb_path, items) in enumerate(groups.items()):
            export_root = pdf_store.export_root_of(Path(items[0]["pdf"])) or Path(root)
            label = workbook_rel_name(Path(wb_path), export_root)
            with _PROCESS_LOCK:
                _PROCESS.update(done=index, current=label)
            while True:
                guard.progress()
                try:
                    outcome = _floor_in_workbook(get_excel, Path(wb_path), items, export_root, label)
                    break
                except Exception as err:
                    debug_log.exception(err, f"PDF 复检向下取整 {label}")
                    retry, message = guard.problem(err, wb_path, label, root,
                                                   close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {label}")
                        continue
                    outcome = [{**item, "ok": False, "reason": message} for item in items]
                    break
            for item in outcome:
                entry = {"pdf": item["pdf"], "name": label, "sheet": item["sheet"], "reason": item.get("reason", "")}
                result["floored" if item["ok"] else "floor_failed"].append(entry)
            with _PROCESS_LOCK:
                _PROCESS.update(done=index + 1)
    finally:
        reporter.suspend_updates(False)
        try:
            file_problem.quit_quietly(excel)
        finally:
            guard.stop()
            wf.pythoncom.CoUninitialize()


def _floor_in_workbook(get_excel: Any, wb_path: Path, items: list[dict], export_root: Path, label: str) -> list[dict]:
    if not wb_path.is_file():
        return [{**item, "ok": False, "reason": "Excel 文件已经不在了"} for item in items]
    if workbook_locked(wb_path):
        raise ValueError("文件正在 Excel 中打开，请先在 Excel 里关掉它（改过的先保存），再点「重试」")
    if not pdf_store.unchanged(export_root, wb_path, sheet_history.fingerprint(wb_path)):
        reason = "Excel 在导出这份 PDF 之后改过（或者找不到导出记录），PDF 的页数不一定对；先点「🔄 重新生成」再标"
        return [{**item, "ok": False, "reason": reason} for item in items]
    book = open_workbook_checked(get_excel(), str(wb_path.resolve()), UpdateLinks=0)
    outcome: list[dict] = []
    try:
        done: set[str] = set()
        for item in items:
            sheet = _find_sheet(book, item["sheet"])
            if sheet is None:
                outcome.append({**item, "ok": False, "reason": "文件里已经没有这张表（可能改名或删掉了）"})
                continue
            if str(sheet.Name) in done:
                outcome.append({**item, "ok": False, "reason": "同一张表标了两次，只改一次"})
                continue
            done.add(str(sheet.Name))
            found = page_fill.floor_pages(sheet, Path(item["pdf"]), wf._export_sheet_pdf)
            if found["ok"]:
                reason = f"{found['before']} 页 → {found['after']} 页，字缩到约 {round(found['scale'] * 100)}%"
                if found["manual"]:
                    reason += "；原来的手动分页不再起作用"
            else:
                reason = found["reason"]
            outcome.append({**item, "ok": found["ok"], "reason": reason, "dirty": found.get("dirty", False)})
        if any(item.get("dirty") for item in outcome):
            # 有一张表改不回去：整份不保存，改好的也不算
            for item in outcome:
                if item["ok"]:
                    item.update(ok=False, reason="同一份 Excel 里另一张表的打印设置改不回去，这份没有保存")
            return outcome
        if not any(item["ok"] for item in outcome):
            return outcome
        book.Save()
    finally:
        book.Close(SaveChanges=False)
    floored = [item["sheet"] for item in outcome if item["ok"]]
    debug_log.event("📄 PDF 复检向下取整已保存", 文件=label, 表=floored)
    wf.log_msg(f"📄 {label}：{'、'.join(floored)} 向下取整，少印一页")
    try:
        wf._export_workbook_pdfs(get_excel, wb_path, export_root, set(), force=True)
    except Exception as err:
        debug_log.exception(err, f"向下取整后重新导出 {label}")
        for item in outcome:
            if item["ok"]:
                item["reason"] += f"；Excel 已保存，但重新导出 PDF 出错（{err}），请点「🔄 重新生成」"
    return outcome


def _open_yellow(yellows: list[dict], root: str, result: dict) -> None:
    """每份 Excel 打开一次，跳到这份文件里第一张标黄的页（选中那一页的范围）。
    每一处（表 + 第几页）都写进结果，网页上点一下就在 Excel 跳过去"""
    groups: dict[str, list[dict]] = {}
    for item in yellows:
        groups.setdefault(item["wb_path"], []).append(item)
    for index, (wb_path, items) in enumerate(groups.items()):
        label = workbook_rel_name(Path(wb_path), Path(root)) if root else Path(wb_path).name
        entries = [{"pdf": item["pdf"], "name": label, "sheet": item["sheet"], "page": spot["page"], "cell": ""}
                   for item in items for spot in (item["spots"] or [{"page": None}])]
        if index >= OPEN_LIMIT:
            result["later"].extend(entries)
            continue
        with _PROCESS_LOCK:
            _PROCESS.update(current=label)
        first = entries[0]
        try:
            pages = len(_page_sizes(Path(first["pdf"])))
            sheet, address, note = open_page(Path(wb_path), first["sheet"], first["page"], pages)
        except Exception as err:
            debug_log.exception(err, f"PDF 复检打开标黄 {label}")
            for entry in entries:
                result["open_failed"].append({**entry, "reason": str(err)})
            continue
        first.update(sheet=sheet, cell=address)
        debug_log.event("📂 PDF 复检打开标黄", 文件=label, 表=sheet, 范围=address or "-",
                        标黄处=[f"{e['sheet']}(第{(e['page'] or 0) + 1}页)" for e in entries], 说明=note or "-")
        for entry in entries:
            entry["reason"] = note if entry is first else ""
            entry["first"] = entry is first
            result["opened"].append(entry)


# ---------- 单份 Excel 重新生成 PDF ----------
# 复检时在 Excel 里改了内容、保存后，想马上看改完的效果：只把这一份 Excel 的 PDF 重新导出一次（它的每张表），
# 不用整个文件夹重导。Excel 和上次导出时一样（哈希没变，多半是还没保存）就先问一句，确认了照样重导。

_REGEN: dict[str, Any] = {"running": False}
_REGEN_LOCK = threading.Lock()


def regen_status() -> dict:
    with _REGEN_LOCK:
        return dict(_REGEN)


def regenerate(pdf: Any, force: Any = False) -> dict:
    """返回 {"started": True} 后台开始；{"unchanged": True} Excel 没改过（force=True 时不问）。网页轮询 regen_status()"""
    if not isinstance(pdf, str) or not pdf:
        raise ValueError("要重新生成哪一份？请刷新页面再试")
    path = Path(pdf).resolve()
    if not _in_folder(path) or path.suffix.lower() != ".pdf":
        raise ValueError(f"「{path.name}」不在当前文件夹里或已经不在了，请重新加载 PDF")
    with wf.STATE_LOCK:
        owner = _owner(path)
        review_root = wf.STATE.get("pdf_review", {}).get("folder") or ""
    if not owner or not Path(owner["wb_path"]).is_file():
        raise ValueError(f"「{path.name}」找不到对应的 Excel 文件")
    wb_path = Path(owner["wb_path"])
    # PDF 在哪个 PDFS 里就导回哪个；旧的混放 PDF 导到现在这个文件夹的 PDFS 里
    root = pdf_store.export_root_of(path) or Path(review_root)
    label = workbook_rel_name(wb_path, root)
    digest = sheet_history.fingerprint(wb_path)
    same = pdf_store.unchanged(root, wb_path, digest)
    debug_log.event("🔄 PDF 复检重新生成", 文件=label, PDF=path.name, 哈希=digest[:12], 没变=same, 确认=bool(force))
    if same and force is not True:
        return {"started": False, "unchanged": True, "name": label}
    from core import bl_links, hidden_repair, link_repair, mark_finder
    with wf.STATE_LOCK, _PROCESS_LOCK, _REGEN_LOCK:
        if (wf.STATE.get("busy") or _PROCESS.get("running") or _REGEN.get("running")
                or link_repair.status().get("running") or mark_finder.status().get("running")
                or hidden_repair.status().get("running") or bl_links.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束再重新生成")
        wf.STATE["busy"] = True  # 要写 PDFS：这时候别开导出 PDF 等任务
        wf.STATE["maintenance"] = "pdf_review_regen"
        _REGEN.clear()
        _REGEN.update(running=True, name=label, pdf=str(path), wb_path=str(wb_path), result=None)

    def run() -> None:
        result: dict[str, Any] = {"ok": False, "name": label, "wb_path": str(wb_path), "pdfs": [], "reason": ""}
        try:
            with debug_log.side_task("pdf_review_regen", {"folder": str(root), "file": label}):
                result.update(_regenerate_workbook(wb_path, root, label))
                debug_log.event("✅ PDF 复检重新生成完成" if result["ok"] else "⚠️ PDF 复检重新生成没成功",
                                文件=label, PDF=len(result["pdfs"]), 原因=result["reason"] or "-")
        except Exception as err:
            debug_log.exception(err, f"PDF 复检重新生成 {label}")
            result["reason"] = f"重新生成出错：{err}"
        finally:
            with wf.STATE_LOCK:
                if wf.STATE.get("maintenance") == "pdf_review_regen":
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)
            with _REGEN_LOCK:
                _REGEN.update(running=False, result=result)

    try:
        threading.Thread(target=run, name="PDF 复检重新生成", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _REGEN_LOCK:
            if wf.STATE.get("maintenance") == "pdf_review_regen":
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)
            _REGEN["running"] = False
        raise
    return {"started": True, "name": label}


def _regenerate_workbook(wb_path: Path, root: Path, label: str) -> dict:
    """后台 Excel 只读打开这一份，把它要导出的表都重新导出到 PDFS（和主页导出 PDF 同一套规则），记下新的哈希"""
    from core import reporter
    guard = file_problem.Guard("PDF 复检：重新生成 PDF")
    excel = None
    reporter.suspend_updates(True)  # 不是主页的批量任务：不写进执行记录
    wf.pythoncom.CoInitialize()

    def get_excel() -> Any:
        nonlocal excel
        if excel is None:
            excel = guard.watch(start_background_excel())
        return excel

    try:
        while True:
            guard.progress()
            try:
                paths = wf._export_workbook_pdfs(get_excel, wb_path, root, set(), force=True)
                return {"ok": True, "pdfs": [str(p) for p in paths], "reason": ""}
            except Exception as err:
                debug_log.exception(err, f"PDF 复检重新生成 {label}")
                retry, message = guard.problem(err, wb_path, label, root,
                                               close=lambda: file_problem.quit_quietly(excel))
                if guard.closed:
                    excel = None
                if retry:
                    wf.log_msg(f"🔁 重试 {label}")
                    continue
                return {"ok": False, "pdfs": [], "reason": message}
    finally:
        reporter.suspend_updates(False)
        try:
            file_problem.quit_quietly(excel)
        finally:
            guard.stop()
            wf.pythoncom.CoUninitialize()
