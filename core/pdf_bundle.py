"""批量导出「一份 Excel 一份 PDF」，PDF 放在 Excel 旁边、同名（用户 2026-11-10 要：领导要 Excel 和 PDF 放一起）。

- aaa.xlsx → 同一文件夹的 aaa.pdf（文件名前的临时标记「★」去掉）；已有同名 PDF 直接覆盖。
- 一份 PDF 里按工作表顺序放所有可见表的全部页；勾了「包含隐藏表」时隐藏表也临时显示出来导出（不保存）。
- 宽度一律压成一页、高自动（已经是宽一页的不动）；打印区域、横竖照 Excel 现有的。只读打开、不改不存 Excel。
- 每张表先单独导出成临时 PDF 再用 PyMuPDF 拼起来：整本一次导出时，各表打印质量不同 Excel 会拆成几份 PDF。
- 跳过名单里的表（默认 変更履歴，网页上可改，比较时不管全角半角、大小写、前后空格）不放进 PDF。
- 默认去掉页眉页脚（公司名、页码等每页重复的字，用户 2026-11-10 要），网页上可以不去。
- 完全空白的表（没内容、没图形）Excel 没东西可印，跳过并在结果里写明。
- 本工具加的标记色（条件格式里带 CHE_MARK 标识的规则）导出前在内存里去掉（用户 2026-11-10 要），不保存，
  还没清色的文件导出来的 PDF 也是干净的；删不掉的（比如表有保护）在结果里写明。

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

from core.pdf_detector import fitz

from core import debug_log, file_problem, progress_log, workflow as wf
from core.common import XL_SHEET_VISIBLE, XL_TYPE_PDF, strip_review_prefix, workbook_rel_name
from core.excel_ops import open_workbook_checked, quit_background_excel, start_background_excel

_LOCK = threading.RLock()
_STATUS: dict[str, Any] = {"running": False}
DEFAULT_SKIP_SHEETS = ["変更履歴"]  # 不用打印的表（用户 2026-11-10 要），网页上可以改
DEFAULT_OPTIONS = {"include_hidden": False, "skip_sheets": DEFAULT_SKIP_SHEETS, "no_header_footer": True}
MAX_SKIP_NAMES = 200


def options_for(options: dict | None = None) -> dict:
    if options is not None and not isinstance(options, dict):
        raise ValueError("导出选项格式不正确")
    result = {**DEFAULT_OPTIONS, **(options or {})}
    if (set(result) != set(DEFAULT_OPTIONS) or type(result["include_hidden"]) is not bool
            or type(result["no_header_footer"]) is not bool):
        raise ValueError("导出选项只能选是否包含隐藏表、跳过哪些表、去不去页眉页脚")
    names = result["skip_sheets"]
    if not isinstance(names, list) or len(names) > MAX_SKIP_NAMES or any(not isinstance(n, str) for n in names):
        raise ValueError("跳过的表名格式不正确")
    result["skip_sheets"] = list(dict.fromkeys(n.strip() for n in names if n.strip()))
    return result


def sheet_key(name: str) -> str:
    """表名比较时不管前后空格、全角半角、大小写（「変更履歴 」「ＢＫ」也认）"""
    import unicodedata
    return unicodedata.normalize("NFKC", str(name)).strip().casefold()


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


_HEADER_FOOTER = ("LeftHeader", "CenterHeader", "RightHeader", "LeftFooter", "CenterFooter", "RightFooter")


def _clear_header_footer(sheet: Any) -> int:
    """导出前临时清掉页眉页脚（公司名、页码、表名……每页重复的字，用户 2026-11-10 要去掉）和打印的行号列标。
    只读打开、关闭时不保存，Excel 原文件不变。返回清掉了几处；读不到 PageSetup 的（极少见）不管"""
    cleared = 0
    try:
        setup = sheet.PageSetup
    except Exception:
        return 0
    for prop in _HEADER_FOOTER:
        try:
            if str(getattr(setup, prop) or ""):
                setattr(setup, prop, "")
                cleared += 1
        except Exception as err:
            debug_log.event("⚠️ 页眉页脚清不掉", 表=str(getattr(sheet, "Name", "?")), 位置=prop, 错误=str(err))
    try:  # 「打印行号列标」：每页边上印 A B C… / 1 2 3…，对 AI 也是干扰
        if setup.PrintHeadings:
            setup.PrintHeadings = False
            cleared += 1
    except Exception:
        pass
    for page in ("FirstPage", "EvenPage"):  # 「首页不同」「奇偶页不同」时另有一套
        try:
            holder = getattr(setup, page)
        except Exception:
            continue
        for prop in _HEADER_FOOTER:
            try:
                part = getattr(holder, prop)
                if str(part.Text or ""):
                    part.Text = ""
                    cleared += 1
            except Exception:
                pass
    return cleared


def _drop_marks(wb: Any, wb_path: Path) -> tuple[int, list[str]]:
    """只读打开的工作簿里，把本工具加的标记（CHE_MARK 条件格式规则）在内存里删掉，不保存。
    返回 (删掉几条, 删不掉的原因)。上传前已经清过色的文件这里删 0 条"""
    from core import batch
    try:
        _present, dropped, failed = batch._drop_rules(wb, set())
    except Exception as err:  # 读不到条件格式（极少见）：照常导出，结果里写明
        debug_log.exception(err, f"导出同名 PDF 前去标记 {wb_path.name}")
        return 0, [f"读不到条件格式，没能去掉标记：{err}"]
    cleared = sum(dropped.values())
    problems = list(failed.values())[:5]
    if cleared or problems:
        debug_log.event("🧽 导出 PDF 前在内存里去掉标记", 文件=wb_path.name, 条=cleared, 删不掉=problems or "无")
    return cleared, problems


def _export_sheet(sheet: Any, target: Path) -> None:
    from core import page_fill  # 交给用户看的 PDF 一律宽一页（用户 2026-11-10 定）；只读打开，不保存
    from core import print_frames
    page_fill.force_one_page_wide(sheet)
    with print_frames.hidden_frames(sheet):  # 整段外面的空心大框不印（用户 2026-11-10 要）
        sheet.ExportAsFixedFormat(
            Type=XL_TYPE_PDF, Filename=str(target), Quality=0, IncludeDocProperties=True,
            IgnorePrintAreas=False, OpenAfterPublish=False,
        )


def export_workbook(excel: Any, wb_path: Path, include_hidden: bool, on_sheet=None,
                    skip_sheets: list[str] | None = None, no_header_footer: bool = False) -> dict:
    """导出一份 Excel 成一份 PDF。返回 {pdf, sheets, pages, skipped:[{sheet, reason}]}"""
    target = pdf_path_for(wb_path)
    tmp_dir = Path(tempfile.mkdtemp(prefix="pdf_bundle_"))
    skipped: list[dict] = []
    skip_keys = {sheet_key(n) for n in skip_sheets or []}
    headers_cleared = 0
    parts: list[tuple[str, Path]] = []
    wb = open_workbook_checked(excel, str(wb_path.resolve()), ReadOnly=True, UpdateLinks=0)
    try:
        marks_cleared, marks_failed = _drop_marks(wb, wb_path)
        sheets = list(wb.Sheets)
        for index, sheet in enumerate(sheets, 1):
            name = str(sheet.Name)
            if on_sheet:
                on_sheet(index, len(sheets), name)
            if sheet_key(name) in skip_keys:
                skipped.append({"sheet": name, "reason": "在跳过名单里"})
                continue
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
            if no_header_footer:
                cleared = _clear_header_footer(sheet)
                if cleared:
                    headers_cleared += cleared
            part = tmp_dir / f"{index:04d}.pdf"
            _export_sheet(sheet, part)
            if not part.is_file():
                raise RuntimeError(f"工作表「{name}」导出后没有生成 PDF")
            parts.append((name, part))
    finally:
        wb.Close(SaveChanges=False)
    try:
        if not parts:
            raise ValueError("没有可以导出的工作表（都是隐藏表、空白表或在跳过名单里）")
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
    if headers_cleared:
        debug_log.event("🧹 导出时去掉页眉页脚", 文件=wb_path.name, 处=headers_cleared)
    return {"pdf": str(target), "sheets": [n for n, _ in parts], "pages": pages, "skipped": skipped,
            "headers_cleared": headers_cleared, "marks_cleared": marks_cleared, "marks_failed": marks_failed}


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

                    entry.update(export_workbook(excel, path, options["include_hidden"], on_sheet, options["skip_sheets"],
                                                 options["no_header_footer"]))
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
                if entry.get("marks_cleared"):
                    reason += f"；导出时去掉标记 {entry['marks_cleared']} 条（Excel 没改）"
                if entry.get("marks_failed"):
                    reason += "；⚠️ 有标记去不掉，PDF 上可能还有标记色：" + "；".join(entry["marks_failed"][:2])
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
