"""批量导出「一份 Excel 一份 PDF」，PDF 放在 Excel 旁边、同名（用户 2026-11-10 要：领导要 Excel 和 PDF 放一起）。

- aaa.xlsx → 同一文件夹的 aaa.pdf（文件名前的临时标记「★」去掉）；已有同名 PDF 直接覆盖。
- 一份 PDF 里按工作表顺序放所有可见表的全部页；勾了「包含隐藏表」时隐藏表也临时显示出来导出（不保存）。
- 打印设置一律按 Excel 现有的（打印区域、缩放、横竖、分页），只读打开、不改不存 Excel。
- 每张表先单独导出成临时 PDF 再用 PyMuPDF 拼起来：整本一次导出时，各表打印质量不同 Excel 会拆成几份 PDF。
- 完全空白的表（没内容、没图形）Excel 没东西可印，跳过并在结果里写明。

和「调整页面」的 PDF（根目录/PDFS、按表一份、哈希沿用）完全分开，互不影响。
"""
from __future__ import annotations

import copy
import os
import shutil
import tempfile
import threading
import time
from pathlib import Path
from typing import Any

import fitz

from core import debug_log, file_problem, progress_log, workflow as wf
from core.common import XL_SHEET_VISIBLE, XL_TYPE_PDF, strip_review_prefix, workbook_rel_name
from core.excel_ops import open_workbook_checked, quit_background_excel, start_background_excel

_LOCK = threading.RLock()
_STATUS: dict[str, Any] = {"running": False}
DEFAULT_OPTIONS = {"include_hidden": False}


def options_for(options: dict | None = None) -> dict:
    if options is not None and not isinstance(options, dict):
        raise ValueError("导出选项格式不正确")
    result = {**DEFAULT_OPTIONS, **(options or {})}
    if set(result) != set(DEFAULT_OPTIONS) or any(type(v) is not bool for v in result.values()):
        raise ValueError("导出选项只能选是否包含隐藏表")
    return result


def pdf_path_for(wb_path: Path) -> Path:
    """aaa.xlsx → 同一文件夹的 aaa.pdf；「★aaa.xlsx」也是 aaa.pdf（★ 只是临时标记）"""
    return wb_path.with_name(strip_review_prefix(wb_path.stem) + ".pdf")


def is_bundle_pdf(pdf: Path) -> bool:
    """旁边有同名 Excel 的 PDF（aaa.pdf 旁边有 aaa.xlsx / ★aaa.xlsx）就是这里导出的整本 PDF"""
    from core.common import EXCEL_EXTENSIONS
    try:
        return any(p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS | {".xlsb"}
                   and strip_review_prefix(p.stem) == pdf.stem for p in pdf.parent.iterdir())
    except OSError:
        return False


def list_workbooks(folder: str) -> tuple[Path, list[Path]]:
    from core import hidden_repair
    return hidden_repair.list_workbooks(folder)


def _sheet_blank(sheet: Any) -> bool:
    """工作表完全空白（没有值、没有图形）：Excel 导出时会报「没有可打印的内容」"""
    try:
        if int(sheet.Shapes.Count):
            return False
        used = sheet.UsedRange
        if int(used.Count) > 1:
            return False
        return str(used.Formula or "") == ""
    except Exception:
        return False  # 图表页等读不到 UsedRange 的照常导出


def _export_sheet(sheet: Any, target: Path) -> None:
    sheet.ExportAsFixedFormat(
        Type=XL_TYPE_PDF, Filename=str(target), Quality=0, IncludeDocProperties=True,
        IgnorePrintAreas=False, OpenAfterPublish=False,
    )


def export_workbook(excel: Any, wb_path: Path, include_hidden: bool, on_sheet=None) -> dict:
    """导出一份 Excel 成一份 PDF。返回 {pdf, sheets, pages, skipped:[{sheet, reason}]}"""
    target = pdf_path_for(wb_path)
    tmp_dir = Path(tempfile.mkdtemp(prefix="pdf_bundle_"))
    skipped: list[dict] = []
    parts: list[tuple[str, Path]] = []
    wb = open_workbook_checked(excel, str(wb_path.resolve()), ReadOnly=True, UpdateLinks=0)
    try:
        sheets = list(wb.Sheets)
        for index, sheet in enumerate(sheets, 1):
            name = str(sheet.Name)
            if on_sheet:
                on_sheet(index, len(sheets), name)
            if int(sheet.Visible) != XL_SHEET_VISIBLE:
                if not include_hidden:
                    skipped.append({"sheet": name, "reason": "隐藏表"})
                    continue
                try:
                    sheet.Visible = XL_SHEET_VISIBLE  # 只读打开，不保存，原文件不变
                except Exception as err:
                    skipped.append({"sheet": name, "reason": f"隐藏表，显示不出来（工作簿可能有保护）：{err}"})
                    continue
            if _sheet_blank(sheet):
                skipped.append({"sheet": name, "reason": "空白表，没有可打印的内容"})
                continue
            part = tmp_dir / f"{index:04d}.pdf"
            _export_sheet(sheet, part)
            if not part.is_file():
                raise RuntimeError(f"工作表「{name}」导出后没有生成 PDF")
            parts.append((name, part))
    finally:
        wb.Close(SaveChanges=False)
    try:
        if not parts:
            raise ValueError("没有可以导出的工作表（都是隐藏表或空白表）")
        merged = fitz.open()
        try:
            toc = []  # PDF 书签：每张表一个，点一下跳到这张表
            for name, part in parts:
                with fitz.open(part) as doc:
                    if doc.page_count:
                        toc.append([1, name, merged.page_count + 1])
                    merged.insert_pdf(doc)
            if toc:
                merged.set_toc(toc)
            pages = merged.page_count
            staged = tmp_dir / "merged.pdf"
            merged.save(str(staged), garbage=3, deflate=True)
        finally:
            merged.close()
        # 先写到 Excel 旁边的临时名，再一次换上，半成品不会留在原名上
        temp_target = target.with_name(f"~{target.stem}.{os.getpid()}.tmp.pdf")
        try:
            shutil.copyfile(staged, temp_target)
            os.replace(temp_target, target)
        except PermissionError as err:
            temp_target.unlink(missing_ok=True)
            raise PermissionError(f"{target.name} 正被别的程序打开（PDF 阅读器？），关掉后点「重试」：{err}") from err
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    return {"pdf": str(target), "sheets": [n for n, _ in parts], "pages": pages, "skipped": skipped}


def _run(root: Path, paths: list[Path], options: dict) -> None:
    files: list[dict] = []
    excel, initialized = None, False
    guard = file_problem.Guard("批量导出同名 PDF")
    claimed: dict[str, str] = {}  # PDF 路径 → 占用它的 Excel（aaa.xlsx 和 aaa.xlsm 在一起时只导第一份）
    try:
        wf.pythoncom.CoInitialize()
        initialized = True
        for index, path in enumerate(paths):
            name = workbook_rel_name(path, root)
            while True:
                entry: dict[str, Any] = {"name": name, "path": str(path), "error": None}
                with _LOCK:
                    _STATUS.update(done=index, current=name)
                    progress_log.step(_STATUS, "正在打开 Excel")
                guard.progress()
                started = time.perf_counter()
                try:
                    target = str(pdf_path_for(path)).lower()
                    if target in claimed:
                        raise ValueError(f"和 {claimed[target]} 会生成同名 PDF，这份没导出，请改名后再导")
                    if excel is None:
                        excel = guard.watch(start_background_excel())

                    def on_sheet(i: int, total: int, sheet: str) -> None:
                        guard.progress()
                        with _LOCK:
                            progress_log.step(_STATUS, f"正在导出第 {i} / {total} 张表：{sheet}")

                    entry.update(export_workbook(excel, path, options["include_hidden"], on_sheet))
                    claimed[target] = name
                except Exception as error:
                    debug_log.exception(error, f"批量导出同名 PDF {name}")
                    retry, entry["error"] = guard.problem(error, path, name, root,
                                                          close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {name}")
                        continue
                break
            files.append(entry)
            seconds = round(time.perf_counter() - started, 2)
            skipped = entry.get("skipped") or []
            debug_log.event("📕 同名 PDF", 文件=name, PDF=entry.get("pdf", ""), 表=len(entry.get("sheets") or []),
                            页=entry.get("pages", 0), 跳过=[f"{s['sheet']}（{s['reason']}）" for s in skipped],
                            错误=entry["error"] or "", 秒=seconds)
            if entry["error"]:
                reason = entry["error"]
            else:
                reason = f"{Path(entry['pdf']).name}：{len(entry['sheets'])} 张表、{entry['pages']} 页"
                if skipped:
                    reason += "；跳过 " + "、".join(f"{s['sheet']}（{s['reason'].split('，')[0].split('：')[0]}）" for s in skipped)
            wf.log_msg(f"📕 {name}：{reason}")
            with _LOCK:
                progress_log.add(_STATUS, name, not entry["error"], reason, path=str(path))
    except Exception as error:
        debug_log.exception(error, "批量导出同名 PDF 任务")
        files.append({"name": "任务", "path": "", "error": str(error)})
    finally:
        try:
            quit_background_excel(excel)
        except Exception:
            pass
        finally:
            guard.stop()
            try:
                if initialized:
                    wf.pythoncom.CoUninitialize()
            finally:
                with _LOCK:
                    _STATUS.update(running=False, done=len(paths), current="", step="",
                                   result={"folder": str(root), "options": options, "files": files})
                with wf.STATE_LOCK:
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)


def start(folder: str, options: dict | None = None) -> None:
    options = options_for(options)
    root, paths = list_workbooks(folder)
    with wf.STATE_LOCK, _LOCK:
        from core import bl_links, hidden_repair, link_repair, mark_finder
        if (wf.STATE.get("busy") or _STATUS.get("running") or link_repair.status().get("running")
                or mark_finder.status().get("running") or hidden_repair.status().get("running")
                or bl_links.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束")
        job = wf.STATE.get("job") or {}
        if job.get("pending_adjustment") or job.get("deferred_adjustments"):
            raise ValueError("还有未完成的文件写入，请先继续或放弃上次写入")
        wf.STATE.update(busy=True, maintenance="pdf_bundle")
        _STATUS.clear()
        _STATUS.update(running=True, folder=str(root), options=options, total=len(paths), done=0, current="")

    def run() -> None:
        with debug_log.side_task("pdf_bundle", {"folder": str(root), "files": len(paths), "options": options}):
            _run(root, paths, options)

    try:
        threading.Thread(target=run, name="批量导出同名 PDF", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
        raise


def status() -> dict:
    with _LOCK:
        return copy.deepcopy(_STATUS)
