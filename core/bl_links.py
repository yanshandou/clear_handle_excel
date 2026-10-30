"""排查并清除「左边单元格含 BL、右边是内容为 @ 的超链接」这种超链接单元格。

规则（两个条件必须同时满足）：
  1. 左边单元格（同一行、紧邻左一列）的内容包含大写 ``BL``，或者有一个以 ``Bl`` / ``bl``
     开头的词（前面不是字母、后面不跟小写字母，如 ``Class:BlKaihoCal``；``Black``、``table`` 不算）；
  2. 右边单元格本身是一个超链接，且显示内容正好是 ``@``。
满足时删除该超链接并清空内容。分两阶段：先扫描列清单，确认后再清除。
"""
from __future__ import annotations

import copy
import hashlib
import re
import threading
from pathlib import Path
from typing import Any

from core import debug_log, workflow as wf
from core.common import workbook_locked, workbook_rel_name
from core.excel_ops import open_workbook_checked, quit_background_excel, start_background_excel

BL_MARK = "BL"
# 驼峰类名里的 Bl（BlKaihoCal、blUtil）：前面不是字母，后面不是小写字母。
BL_WORD = re.compile(r"(?<![A-Za-z])[Bb][Ll](?![a-z])")
AT_MARK = "@"

_LOCK = threading.RLock()
_STATUS: dict[str, Any] = {"running": False}


def fingerprint(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _text(value: Any) -> str:
    return "" if value is None else str(value)


def _clip(value: Any, limit: int = 40) -> str:
    text = _text(value).strip()
    return text if len(text) <= limit else text[:limit] + "…"


def _address(rng: Any) -> str:
    return str(rng.Address).replace("$", "")


def _area(rng: Any) -> Any:
    try:
        return rng.MergeArea if getattr(rng, "MergeCells", False) is True else rng
    except Exception:
        return rng


def _top_cell(rng: Any) -> Any:
    area = _area(rng)
    try:
        return area.Cells(1, 1)
    except Exception:
        return area


def _link_display(link: Any) -> str:
    try:
        return _text(link.TextToDisplay)
    except Exception:
        return ""


def _link_target(link: Any) -> str:
    try:
        addr = _text(getattr(link, "Address", None)).strip()
    except Exception:
        addr = ""
    try:
        sub = _text(getattr(link, "SubAddress", None)).strip()
    except Exception:
        sub = ""
    if addr and sub:
        return f"{addr}#{sub}"
    return addr or (f"#{sub}" if sub else "")


def is_bl_left(value: Any) -> bool:
    """左边单元格是否命中：内容包含大写 BL，或有 Bl/bl 开头的词。"""
    text = _text(value)
    return BL_MARK in text or BL_WORD.search(text) is not None


def is_at_content(value: Any) -> bool:
    """右边单元格内容是否正好是 @（去掉首尾空格）。"""
    return _text(value).strip() == AT_MARK


def _is_at_hyperlink(cell: Any, link: Any) -> bool:
    return is_at_content(_link_display(link)) or is_at_content(getattr(cell, "Value2", None))


def scan_file(excel: Any, path: Path) -> dict:
    stamp = fingerprint(path)
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0, ReadOnly=True)
    try:
        hits: list[dict] = []
        seen: set[tuple[str, str]] = set()
        for sheet in book.Worksheets:
            for index in range(1, int(sheet.Hyperlinks.Count) + 1):
                link = sheet.Hyperlinks(index)
                try:
                    area = _area(link.Range)
                    top = _top_cell(area)
                    row, col = int(top.Row), int(top.Column)
                except Exception:
                    continue  # 个别失效的链接对象读不到范围，跳过
                if col <= 1:
                    continue  # 第一列没有左边邻居
                cell = sheet.Cells(row, col)
                if not _is_at_hyperlink(cell, link):
                    continue
                addr = _address(area)
                if (sheet.Name, addr) in seen:
                    continue
                raw_left = sheet.Cells(row, col - 1)
                left_area = _area(raw_left)
                left_top = _top_cell(left_area)
                left_value = _text(left_top.Value2)
                if not is_bl_left(left_value):
                    continue
                seen.add((sheet.Name, addr))
                hits.append({
                    "sheet": sheet.Name,
                    "addr": addr,
                    "anchor": _address(cell),
                    "left_addr": _address(left_area),
                    "left_value": left_value,
                    "value": _link_display(link).strip() or _text(cell.Value2).strip(),
                    "link_address": _link_target(link),
                })
    finally:
        book.Close(SaveChanges=False)
    if stamp != fingerprint(path):
        raise ValueError("文件在扫描期间发生变化，请重新扫描")
    return {"path": str(path), "fingerprint": stamp, "hits": hits}


def _still_matches(sheet: Any, cell: Any) -> bool:
    try:
        area = _area(cell)
        top = _top_cell(area)
        if int(area.Hyperlinks.Count) == 0:
            return False
        link = area.Hyperlinks(1)
        if _is_at_hyperlink(top, link):
            row, col = int(top.Row), int(top.Column)
            if col <= 1:
                return False
            left_top = _top_cell(sheet.Cells(row, col - 1))
            return is_bl_left(left_top.Value2)
    except Exception:
        return False
    return False


def clear_file(excel: Any, file: dict) -> list[dict]:
    path = Path(file["path"])
    if fingerprint(path) != file["fingerprint"]:
        raise ValueError("当前文件在扫描后发生变化，请重新扫描")
    if workbook_locked(path):
        raise ValueError("文件正在 Excel 中打开，请保存并关闭后再清除")
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0)
    results: list[dict] = []
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开，没有修改")
        # 写入前整体核对，任何一处不再符合规则都不保存本文件。
        for hit in file["hits"]:
            sheet = book.Worksheets(hit["sheet"])
            if not _still_matches(sheet, sheet.Range(hit["addr"])):
                raise ValueError(f"{hit['sheet']}!{hit['anchor']} 已变化，请重新扫描")
        for hit in file["hits"]:
            sheet = book.Worksheets(hit["sheet"])
            area = _area(sheet.Range(hit["addr"]))
            top = _top_cell(area)
            area.Hyperlinks.Delete()
            area.ClearContents()
            if int(area.Hyperlinks.Count) != 0 or _text(top.Value2) not in ("", None):
                raise ValueError(f"{hit['sheet']}!{hit['anchor']} 清除校验失败，本文件没有保存")
            results.append({**hit, "status": "ok", "reason": "已删除超链接并清空内容"})
        book.Save()
    finally:
        book.Close(SaveChanges=False)
    return results


def _progress(done: int, name: str, total: int) -> None:
    with _LOCK:
        _STATUS.update(done=done, current=name, total=total)


def _run(phase: str, folder: str, targets: list) -> None:
    excel = None
    com = False
    results: list[dict] = []
    try:
        wf.pythoncom.CoInitialize()
        com = True
        excel = start_background_excel()
        for index, target in enumerate(targets):
            if phase == "scan":
                name = workbook_rel_name(target, Path(folder))
                entry: dict[str, Any] = {"name": name, "path": str(target), "hits": [], "cells": [], "error": None}
            else:
                entry = {**target, "cells": []}
                name = entry["name"]
            _progress(index, name, len(targets))
            try:
                if phase == "scan":
                    entry.update(scan_file(excel, target))
                else:
                    entry["cells"] = clear_file(excel, entry)
            except Exception as error:
                entry["error"] = str(error)
                debug_log.exception(error, f"BL 超链接 {phase} {name}")
            if phase == "scan":
                for hit in entry["hits"]:
                    debug_log.event("🔎 BL 超链接", 位置=f"{hit['sheet']}!{hit['anchor']}",
                                    左边=_clip(hit["left_value"]), 内容=hit["value"])
                wf.log_msg(f"🔎 {name}：发现 {len(entry['hits'])} 处" + (f"；{entry['error']}" if entry["error"] else ""))
            else:
                for cell in entry["cells"]:
                    debug_log.event("🧽 清除 BL 超链接", 位置=f"{cell['sheet']}!{cell['anchor']}", 结果=cell["reason"])
                count = sum(cell.get("status") == "ok" for cell in entry["cells"])
                wf.log_msg(f"🧽 {name}：已清除 {count} 处" + (f"；{entry['error']}" if entry["error"] else ""))
            results.append(entry)
    except Exception as error:
        debug_log.exception(error, "BL 超链接任务")
        results.append({"name": "任务", "path": "", "hits": [], "cells": [], "error": str(error)})
    finally:
        try:
            quit_background_excel(excel)
        finally:
            if com:
                wf.pythoncom.CoUninitialize()
            with _LOCK:
                key = "result" if phase == "scan" else "applied"
                _STATUS.update(running=False, done=len(targets), current="")
                _STATUS[key] = {"folder": folder, "files": results}
            with wf.STATE_LOCK:
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)


def _launch(phase: str, folder: str, targets: list) -> None:
    with wf.STATE_LOCK, _LOCK:
        from core import hidden_repair, link_repair, mark_finder
        if (wf.STATE.get("busy") or _STATUS.get("running")
                or link_repair.status().get("running") or mark_finder.status().get("running")
                or hidden_repair.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束")
        job = wf.STATE.get("job") or {}
        if job.get("pending_adjustment") or job.get("deferred_adjustments"):
            raise ValueError("还有未完成的文件写入，请先继续或放弃上次写入")
        wf.STATE["busy"] = True
        wf.STATE["maintenance"] = "bl_link_" + phase
        _STATUS.clear()
        _STATUS.update(running=True, phase=phase, folder=folder, total=len(targets), done=0, current="")

    def run() -> None:
        with debug_log.side_task("bl_link_" + phase, {"folder": folder, "files": len(targets)}):
            from core import operation_history
            kind = "bl_scan" if phase == "scan" else "bl_clear"
            ids = operation_history.begin(folder, [kind]) if kind else []
            try:
                _run(phase, folder, targets)
                result = status().get("result" if phase == "scan" else "applied", {})
                operation_history.finish(ids, "failed" if any(f.get("error") for f in result.get("files", [])) else "completed")
            except Exception:
                operation_history.finish(ids, "failed")
                raise

    try:
        threading.Thread(target=run, name="BL超链接" + phase, daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
        raise


def start(folder: str) -> None:
    """扫描一个文件夹，列出所有符合条件的超链接单元格。"""
    root, paths = wf.list_workbooks(folder)
    _launch("scan", str(root), paths)


def clear() -> None:
    """清除上一次扫描结果里符合条件的超链接单元格。"""
    with _LOCK:
        result = copy.deepcopy(_STATUS.get("result"))
    if not result:
        raise ValueError("请先扫描 BL 超链接")
    files = [file for file in result["files"] if file.get("hits") and not file["error"]]
    if not files:
        raise ValueError("没有需要清除的超链接单元格")
    _launch("clear", result["folder"], files)


def status() -> dict:
    with _LOCK:
        return copy.deepcopy(_STATUS)
