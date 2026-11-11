"""统一字体：把游ゴシック一家换成标准字体ＭＳ Ｐゴシック（用户 2026-11-10 要）。

游ゴシック比ＭＳ Ｐゴシック宽、行距大，原来放得下的字被挤出去，查切字时报成误报——主要出在图形（对话框、文本框）里。
- 图形：每段文字（Run）各自判断，游ゴシック的换掉，别的字体、字号、颜色都不动；组合图形里的也查。
- 格子：只看有内容的格子（xlsx / xlsm 用 openpyxl 先找出来，再用 Excel 改）；.xls / .xlsb 只改图形。
- 不改 Excel 的「標準」样式（改了整张表的标准列宽、行高都会变）。
- 精简 Sheet 之后自动检查里的第 4 小项（clean_chain），主页「更多工具」也能单独跑（/fix-fonts）。
"""
from __future__ import annotations

import copy
import threading
import time
import unicodedata
from pathlib import Path
from typing import Any

from core import debug_log, file_problem, progress_log, workflow as wf
from core.common import workbook_locked, workbook_rel_name
from core.excel_ops import open_workbook_checked, quit_background_excel, start_background_excel

STANDARD_FONT = "ＭＳ Ｐゴシック"
_LOCK = threading.RLock()
_STATUS: dict[str, Any] = {"running": False}
MAX_ADDRESS_LEN = 240  # Range("A1,B2,…") 的地址最长 255 个字符
MSO_GROUP = 6


def is_yu_gothic(name: Any) -> bool:
    """游ゴシック、游ゴシック 本文、游ゴシック Light、游ｺﾞｼｯｸ、Yu Gothic（UI）……都算"""
    text = unicodedata.normalize("NFKC", str(name or "")).casefold().replace(" ", "")
    return "游ゴシック" in text or "yugothic" in text


def list_workbooks(folder: str) -> tuple[Path, list[Path]]:
    from core import hidden_repair
    return hidden_repair.list_workbooks(folder)


# ---------- 格子 ----------

def yu_cells(path: Path) -> dict[str, list[str]]:
    """{表名: [有内容、字体是游ゴシック的格子地址]}。只读 xlsx / xlsm；别的格式返回空（只改图形）"""
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return {}
    import openpyxl
    found: dict[str, list[str]] = {}
    book = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        for sheet in book.worksheets:
            cells = [cell.coordinate for row in sheet.iter_rows() for cell in row
                     if getattr(cell, "value", None) not in (None, "") and is_yu_gothic(cell.font and cell.font.name)]
            if cells:
                found[sheet.title] = cells
    finally:
        book.close()
    return found


def _address_chunks(cells: list[str]) -> list[str]:
    chunks, current = [], ""
    for addr in cells:
        if current and len(current) + len(addr) + 1 > MAX_ADDRESS_LEN:
            chunks.append(current)
            current = ""
        current = f"{current},{addr}" if current else addr
    if current:
        chunks.append(current)
    return chunks


def fix_cells(sheet: Any, cells: list[str]) -> int:
    for chunk in _address_chunks(cells):
        sheet.Range(chunk).Font.Name = STANDARD_FONT
    return len(cells)


def refresh_changed_cells(book: Any, cells_by_sheet: dict[str, list[str]],
                          cleared: list[dict], restored: list[dict]) -> dict[str, list[str]]:
    """用内存中的修改更新从磁盘读出的字体清单。"""
    updated = {name: dict.fromkeys(cells) for name, cells in cells_by_sheet.items()}
    for edits, has_value in ((cleared, False), (restored, True)):
        for item in edits:
            if item["status"] != "ok":
                continue
            cells = updated.setdefault(item["sheet"], {})
            for cell in book.Worksheets(item["sheet"]).Range(item["addr"]).Cells:
                addr = str(cell.Address).replace("$", "")
                cells.pop(addr, None)
                if has_value and cell.Value2 not in (None, "") and is_yu_gothic(cell.Font.Name):
                    cells[addr] = None
    return {name: list(cells) for name, cells in updated.items() if cells}


# ---------- 图形 ----------

def _shape_items(shapes: Any) -> list[Any]:
    """所有图形，组合图形拆开看里面的每一个"""
    items = []
    for index in range(1, int(shapes.Count) + 1):
        shape = shapes.Item(index)
        try:
            if int(shape.Type) == MSO_GROUP:
                group = shape.GroupItems
                items.extend(group.Item(i) for i in range(1, int(group.Count) + 1))
                continue
        except Exception:
            pass
        items.append(shape)
    return items


def fix_shape(shape: Any) -> int:
    """图形里游ゴシック的那几段换成ＭＳ Ｐゴシック，返回改了几段；没有文字的图形（图片、线）返回 0"""
    try:
        frame = shape.TextFrame2
        if not frame.HasText:
            return 0
        text = frame.TextRange
        runs = text.Runs()
        count = int(runs.Count)
    except Exception:
        return 0
    changed = 0
    for index in range(1, count + 1):
        try:
            font = text.Runs(index).Font
            if is_yu_gothic(font.Name) or is_yu_gothic(font.NameFarEast):
                font.Name = STANDARD_FONT
                font.NameFarEast = STANDARD_FONT
                changed += 1
        except Exception as err:
            debug_log.event("⚠️ 图形字体改不了", 图形=str(getattr(shape, "Name", "?")), 错误=str(err))
    return changed


# ---------- 一份 Excel ----------

def fix_open_workbook(book: Any, cells_by_sheet: dict[str, list[str]], on_sheet=None) -> dict:
    sheets: list[dict] = []
    worksheets = list(book.Worksheets)
    for index, sheet in enumerate(worksheets, 1):
        name = str(sheet.Name)
        if on_sheet:
            on_sheet(index, len(worksheets), name)
        cell_count = fix_cells(sheet, cells_by_sheet.get(name, [])) if cells_by_sheet.get(name) else 0
        shape_names, shape_runs = [], 0
        try:
            items = _shape_items(sheet.Shapes)
        except Exception:
            items = []
        for shape in items:
            runs = fix_shape(shape)
            if runs:
                shape_runs += runs
                shape_names.append(str(getattr(shape, "Name", "")))
        if cell_count or shape_names:
            sheets.append({"sheet": name, "cells": cell_count, "shapes": len(shape_names), "names": shape_names[:20]})
    return {"cells": sum(s["cells"] for s in sheets), "shapes": sum(s["shapes"] for s in sheets), "sheets": sheets}


def fix_workbook(excel: Any, path: Path, on_sheet=None) -> dict:
    """返回 {cells, shapes, sheets:[{sheet, cells, shapes, names:[图形名]}]}；没东西要改就不保存"""
    if workbook_locked(path):
        raise ValueError("文件正在 Excel 中打开，请保存并关闭后再统一字体")
    cells_by_sheet = yu_cells(path)
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开，没有修改")
        result = fix_open_workbook(book, cells_by_sheet, on_sheet)
        if result["sheets"]:
            book.Save()
            debug_log.saved_excel("统一字体", path)
    finally:
        book.Close(SaveChanges=False)
    return result


def describe(entry: dict) -> str:
    if entry.get("error"):
        return entry["error"]
    parts = [f"{entry['shapes']} 个图形" if entry.get("shapes") else "", f"{entry['cells']} 个格子" if entry.get("cells") else ""]
    parts = [p for p in parts if p]
    return f"换成标准字体：{'、'.join(parts)}" if parts else "字体都是对的"


# ---------- 后台任务 ----------

def _run(root: Path, paths: list[Path]) -> None:
    files: list[dict] = []
    excel, initialized = None, False
    guard = file_problem.Guard("统一字体")
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
                started = time.perf_counter()
                try:
                    if guard.file(name):  # 进度框能「跳过这一份」
                        excel = None
                    if excel is None:
                        excel = guard.watch(start_background_excel())

                    def on_sheet(i: int, total: int, sheet: str) -> None:
                        guard.progress()
                        with _LOCK:
                            progress_log.step(_STATUS, f"正在检查第 {i} / {total} 张表：{sheet}")

                    entry.update(fix_workbook(excel, path, on_sheet))
                except Exception as error:
                    debug_log.exception(error, f"统一字体 {name}")
                    retry, entry["error"] = guard.problem(error, path, name, root,
                                                          close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {name}")
                        continue
                break
            files.append(entry)
            debug_log.event("🔤 统一字体", 文件=name, 图形=entry.get("shapes", 0), 格子=entry.get("cells", 0),
                            表=[f"{s['sheet']}（图形 {s['shapes']}、格子 {s['cells']}）" for s in entry.get("sheets", [])],
                            错误=entry["error"] or "", 秒=round(time.perf_counter() - started, 2))
            wf.log_msg(f"🔤 {name}：{describe(entry)}")
            with _LOCK:
                progress_log.add(_STATUS, name, not entry["error"], describe(entry), path=str(path))
    except Exception as error:
        debug_log.exception(error, "统一字体任务")
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
                                   result={"folder": str(root), "files": files})
                with wf.STATE_LOCK:
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)


def start(folder: str) -> None:
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
        wf.STATE.update(busy=True, maintenance="font_fix")
        _STATUS.clear()
        _STATUS.update(running=True, folder=str(root), total=len(paths), done=0, current="")

    def run() -> None:
        with debug_log.side_task("font_fix", {"folder": str(root), "files": len(paths)}):
            _run(root, paths)

    try:
        threading.Thread(target=run, name="统一字体", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
        raise


def status() -> dict:
    with _LOCK:
        return copy.deepcopy(_STATUS)
