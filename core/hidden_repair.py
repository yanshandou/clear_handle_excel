"""对照只读备份，恢复工作表/列隐藏和超窄列宽；不推测作者的隐藏意图。"""
from __future__ import annotations

import copy
import hashlib
import json
import posixpath
import tempfile
import threading
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from core import debug_log, workflow as wf, sheet_history
from core.backup_compare import inferred_renames, pair_files, pair_sheets, require_manual_file_check
from core.common import EXCEL_EXTENSIONS, col_index_to_letters, workbook_locked, workbook_rel_name
from core.excel_ops import open_workbook_checked, quit_background_excel, start_background_excel
from core.link_repair import fingerprint

VISIBLE, HIDDEN, VERY_HIDDEN = -1, 0, 2
NARROW_WIDTH = 1.0  # 字符宽度；超窄列仍可见，恢复列宽而不改成隐藏列。
MAX_CONTENT_CELLS = 1_000_000
_LOCK = threading.RLock()
_STATUS: dict[str, Any] = {"running": False}
DEFAULT_OPTIONS = {"sheets": True, "columns": True, "narrow": True}


def list_workbooks(folder: str) -> tuple[Path, list[Path]]:
    root = Path(folder.strip().strip('"').strip("'")).expanduser().resolve()
    if not root.is_dir():
        raise ValueError(f"文件夹不存在: {root}")
    paths = sorted((path for path in root.rglob("*") if path.is_file()
                    and path.suffix.lower() in EXCEL_EXTENSIONS | {".xlsb"} and not path.name.startswith("~$")),
                   key=lambda path: workbook_rel_name(path, root))
    if not paths:
        raise ValueError("未找到 Excel 文件")
    return root, paths


def options_for(options: dict | None = None) -> dict:
    if options is not None and not isinstance(options, dict):
        raise ValueError("隐藏恢复选项格式不正确")
    result = {**DEFAULT_OPTIONS, **(options or {})}
    if set(result) != set(DEFAULT_OPTIONS) or any(type(value) is not bool for value in result.values()):
        raise ValueError("隐藏恢复选项只能选择工作表、隐藏列和超窄列")
    if not any(result.values()):
        raise ValueError("请至少选择一种要恢复的隐藏状态")
    return result


def column_address(first: int, last: int) -> str:
    return f"{col_index_to_letters(first)}:{col_index_to_letters(last)}"


def column_specs(path: Path) -> dict[str, list[dict]]:
    """从文件的列定义读取分段；Excel 对混合列区域的 COM 属性可能只返回第一列。"""
    with zipfile.ZipFile(path) as archive:
        relations = {node.attrib["Id"]: node.attrib["Target"]
                     for node in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
        result = {}
        for node in ET.fromstring(archive.read("xl/workbook.xml")).findall("{*}sheets/{*}sheet"):
            rel = next(value for key, value in node.attrib.items() if key.endswith("}id"))
            target = relations[rel]
            target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            specs = []
            with archive.open(target) as source:
                for _, col in ET.iterparse(source, events=("start",)):
                    tag = col.tag.rsplit("}", 1)[-1]
                    if tag == "sheetData":
                        break  # 列定义在单元格数据之前，不解析整张表的内容。
                    if tag == "col":
                        first, last = int(col.attrib["min"]), int(col.attrib["max"])
                        hidden = col.attrib.get("hidden", "0").lower() in {"1", "true"} or float(col.attrib.get("width", "1")) == 0
                        specs.append({"first": first, "last": last, "hidden": hidden})
            result[node.attrib["name"]] = sorted(specs, key=lambda item: item["first"])
        return result


def column_runs(sheet: Any, specs: list[dict], limit: int | None = None) -> list[dict]:
    # XML width 包含字体留白，不能直接当作 Excel.ColumnWidth；每段只读一个代表列的精确值。
    last, cursor, result = int(sheet.Columns.Count) if limit is None else limit, 1, []
    # xls 临时转成 xlsx 后多出的列不属于原工作表，不能作为恢复目标。
    specs = [{**spec, "last": min(spec["last"], last)} for spec in specs if spec["first"] <= last]
    for spec in [*specs, {"first": last + 1, "last": last, "hidden": False}]:
        if spec["first"] < cursor or spec["last"] > last:
            raise ValueError("文件列定义重叠或超出范围，不能确认原列位置")
        segments = []
        if cursor < spec["first"]:
            segments.append({"first": cursor, "last": spec["first"] - 1, "hidden": False})
        if spec["first"] <= last:
            segments.append(spec)
        for segment in segments:
            letter = col_index_to_letters(segment["first"])
            width = None if segment["hidden"] else float(sheet.Range(f"{letter}:{letter}").EntireColumn.ColumnWidth)
            result.append({**segment, "width": width})
        cursor = spec["last"] + 1
    return result


def content_info(sheet: Any, history: sheet_history.Projection | None = None) -> dict:
    """按位置核对全部内容和公式；忽略单纯因格式变化而扩大的空白 UsedRange。"""
    used = sheet.UsedRange
    row, col = int(used.Row), int(used.Column)
    rows, cols = int(used.Rows.Count), int(used.Columns.Count)
    if rows * cols > MAX_CONTENT_CELLS:
        raise ValueError("使用区域超过 100 万格，无法完整核对内容，请人工确认隐藏状态")
    digest, words, populated = hashlib.sha256(), set(), 0
    step = max(1, 20_000 // cols)
    for start in range(row, row + rows, step):
        end = min(start + step, row + rows) - 1
        area = sheet.Range(sheet.Cells(start, col), sheet.Cells(end, col + cols - 1))
        values = area.Formula
        matrix = values if isinstance(values, (tuple, list)) else ((values,),)
        # Excel 的 Formula 对数值常量可能返回字符串（42 -> "42"）；Value2 保留数值类型，
        # 才能与删表时记录的公式转值结果比较，也能区分数字和数字文本。
        try:
            constants = area.Value2
        except AttributeError:
            constants = values
        constants = constants if isinstance(constants, (tuple, list)) else ((constants,),)
        for r, items in enumerate(matrix, start):
            for c, value in enumerate(items, col):
                if not isinstance(value, str) or not value.startswith("="):
                    value = constants[r - start][c - col]
                value = (history.value(sheet.Name, f"{col_index_to_letters(c)}{r}", value)
                         if history and history.operations else sheet_history.formula_value(value))
                if value is None or value == "":
                    continue
                populated += 1
                digest.update(json.dumps([r, c, value], ensure_ascii=False, default=str, separators=(",", ":")).encode("utf-8"))
                digest.update(b"\n")
                if r <= 200 and c < col + 256 and not str(value).startswith("="):
                    words.add(str(value))
    return {"signature": digest.hexdigest(), "populated": populated, "words": sorted(words)}


def snapshot(excel: Any, path: Path, options: dict, *, history_path: Path | None = None,
             current_names: list[str] | None = None, old_hash: str = "",
             projection: sheet_history.Projection | None = None) -> dict:
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0, ReadOnly=True)
    temporary = None
    try:
        sheet_names = [s.Name for s in book.Sheets]
        history = projection if projection is not None else (sheet_history.for_paths(
            history_path, path, current_names, sheet_names, old_hash=old_hash) if history_path is not None else None)
        sheets = []
        for sheet in book.Worksheets:
            item = {"name": sheet.Name, "code": sheet.CodeName, "visible": int(sheet.Visible),
                    "column_count": int(sheet.Columns.Count),
                    "columns": [], "words": [], "signature": None, "content_error": "", "columns_error": "",
                    "protected": bool(sheet.ProtectContents)}
            try:
                item.update(content_info(sheet, history) if history else content_info(sheet))
            except Exception as error:
                item["content_error"] = str(error)
            sheets.append(item)
        if options["columns"] or options["narrow"]:
            try:
                source_path = path
                if path.suffix.lower() not in {".xlsx", ".xlsm"}:
                    # 二进制格式仅另存一份临时 xlsx 读取列定义；原件只读打开，从不保存回原路径。
                    temporary = tempfile.TemporaryDirectory(prefix="hidden_columns_")
                    source_path = Path(temporary.name) / "columns.xlsx"
                    book.SaveAs(str(source_path), FileFormat=51)
                specs = column_specs(source_path)
                for item in sheets:
                    try:
                        item["columns"] = column_runs(book.Worksheets(item["name"]), specs[item["name"]], item["column_count"])
                    except Exception as error:
                        item["columns_error"] = str(error)
            except Exception as error:
                for item in sheets:
                    item["columns_error"] = str(error)
        return {"sheets": sheets, "sheet_names": sheet_names, "structure_protected": bool(book.ProtectStructure),
                "visible_count": sum(int(sheet.Visible) == VISIBLE for sheet in book.Sheets)}
    finally:
        try:
            book.Close(SaveChanges=False)
        finally:
            if temporary is not None:
                temporary.cleanup()


def plan_changes(current: dict, old: dict, options: dict, history: sheet_history.Projection | None = None) -> list[dict]:
    matches = pair_sheets(current["sheets"], old["sheets"], title=lambda sheet: sheet["name"],
                          code=lambda sheet: sheet["code"], words=lambda sheet: set(sheet["words"]), history=history)
    result = []
    for sheet in current["sheets"]:
        base = {"sheet": sheet["name"], "addr": "A1", "anchor": "A1", "kind": "sheet", "action": "restore"}
        match = matches.get(sheet["name"])
        if match is None:
            result.append({**base, "action": "skip", "current_state": "待核对", "old_state": "无法配对",
                           "reason": "备份工作表无法唯一配对（可能存在同名旧表和「新」表，证据不足），请人工检查"})
            continue
        source, how = match
        base.update(old_sheet=source["name"], match=how)
        candidates = []
        if options["sheets"] and source["visible"] in (HIDDEN, VERY_HIDDEN) and sheet["visible"] == VISIBLE:
            candidates.append({**base, "before": VISIBLE, "value": source["visible"], "current_state": "工作表可见",
                               "old_state": "工作表深度隐藏" if source["visible"] == VERY_HIDDEN else "工作表隐藏"})
        if options["columns"] or options["narrow"]:
            if sheet["columns_error"] or source["columns_error"]:
                result.append({**base, "kind": "columns", "action": "skip", "current_state": "列状态待核对", "old_state": "读取失败",
                               "reason": sheet["columns_error"] or source["columns_error"]})
            else:
                index = 0
                for original in source["columns"]:
                    hidden = original["hidden"]
                    if not ((hidden and options["columns"]) or (not hidden and options["narrow"] and 0 < original["width"] <= NARROW_WIDTH)):
                        continue
                    while index < len(sheet["columns"]) and sheet["columns"][index]["last"] < original["first"]:
                        index += 1
                    for present in sheet["columns"][index:]:
                        if present["first"] > original["last"]:
                            break
                        if present["hidden"] or (not hidden and present["width"] <= original["width"] + 0.05):
                            continue
                        first, last = max(original["first"], present["first"]), min(original["last"], present["last"])
                        candidates.append({**base, "kind": "columns" if hidden else "narrow", "first": first, "last": last,
                                           "addr": column_address(first, last), "anchor": f"{col_index_to_letters(first)}1",
                                           "before": present["width"], "value": None if hidden else original["width"],
                                           "current_state": f"可见，列宽 {present['width']:g}",
                                           "old_state": "隐藏列" if hidden else f"超窄列，列宽 {original['width']:g}"})
        for item in candidates:
            reason = sheet["content_error"] or source["content_error"]
            if not reason and sheet["signature"] != source["signature"]:
                reason = "新旧表的内容、公式或位置不同，无法确认是否移动过行列，请人工检查"
            if not reason and how not in {"name", "history"} and not sheet.get("populated"):
                reason = "空白工作表名称不同，缺少内容依据，请人工确认配对"
            if not reason and item["kind"] == "sheet" and current["structure_protected"]:
                reason = "工作簿结构受保护，不能自动隐藏工作表"
            if not reason and item["kind"] != "sheet" and sheet["protected"]:
                reason = "工作表受保护，请人工确认列设置"
            item.update(action="skip" if reason else "restore", reason=reason or "与备份内容和位置一致，可以恢复")
            result.append(item)
    sheet_items = [item for item in result if item["kind"] == "sheet" and item["action"] == "restore"]
    if sheet_items and len(sheet_items) >= current["visible_count"]:
        for item in sheet_items:
            item.update(action="skip", reason="恢复后会没有可见工作表，Excel 至少需要一张可见表，请人工确认")
    return result


def plan_file(excel: Any, path: Path, old_path: Path, options: dict) -> dict:
    if path.resolve() == old_path.resolve():
        raise ValueError("当前文件不能作为自己的备份，请选择原始备份")
    stamp, old_stamp = fingerprint(path), fingerprint(old_path)
    # 同名工作簿依次打开并关闭，兼容 Excel 不能同时打开同名文件的限制。
    current = snapshot(excel, path, options)
    names = current.get("sheet_names", [s["name"] for s in current["sheets"]])
    old = snapshot(excel, old_path, options, history_path=path, current_names=names, old_hash=old_stamp)
    history = sheet_history.for_paths(path, old_path, names, old.get("sheet_names", [s["name"] for s in old["sheets"]]), old_hash=old_stamp)
    matches = pair_sheets(current["sheets"], old["sheets"], title=lambda sheet: sheet["name"],
                          code=lambda sheet: sheet["code"], words=lambda sheet: set(sheet["words"]), history=history)
    renames = inferred_renames(matches, lambda sheet: sheet["name"])
    if renames:
        comparison = sheet_history.with_inferred_renames(history, renames)
        old = snapshot(excel, old_path, options, projection=comparison)
    if stamp != fingerprint(path) or old_stamp != fingerprint(old_path):
        raise ValueError("当前文件或备份在对比期间发生变化，请重新扫描")
    return {"path": str(path), "old": str(old_path), "fingerprint": stamp, "old_fingerprint": old_stamp,
            "checks": {sheet["name"]: sheet["signature"] for sheet in current["sheets"]},
            "cells": plan_changes(current, old, options, history)}


def apply_file(excel: Any, file: dict) -> list[dict]:
    path, old_path = Path(file["path"]), Path(file["old"])

    def check_versions(current_stamp: str = file["fingerprint"]) -> None:
        if fingerprint(path) != current_stamp:
            raise ValueError("当前文件发生变化，请重新扫描")
        if fingerprint(old_path) != file["old_fingerprint"]:
            raise ValueError("备份发生变化，请重新扫描")

    check_versions()
    if workbook_locked(path):
        raise ValueError("文件正在 Excel 中打开，请保存并关闭后再修复")
    ready = [item for item in file["cells"] if item["action"] == "restore"]
    if not ready:
        return []
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开，没有修改")
        # 老版 xls 在取得写锁时就会更新 OLE 容器字节；打开前已校验扫描版本，
        # 保存前再与取得写锁后的版本比较，避免把 Excel 自己的更新当作外部编辑。
        opened_stamp = fingerprint(path)
        for name in {item["sheet"] for item in ready}:
            if content_info(book.Worksheets(name))["signature"] != file["checks"][name]:
                raise ValueError(f"{name} 的内容已变，本文件没有保存，请重新扫描")
        for item in ready:
            sheet = book.Worksheets(item["sheet"])
            if item["kind"] == "sheet":
                unchanged = int(sheet.Visible) == item["before"]
            else:
                unchanged = all(sheet.Columns(index).Hidden is False and abs(float(sheet.Columns(index).ColumnWidth) - item["before"]) < 0.001
                                for index in range(item["first"], item["last"] + 1))
            if not unchanged:
                raise ValueError(f"{item['sheet']}!{item['addr']} 的隐藏状态或列宽已变，请重新扫描")
        for item in ready:
            sheet = book.Worksheets(item["sheet"])
            if item["kind"] == "sheet":
                sheet.Visible = item["value"]
                verified = int(sheet.Visible) == item["value"]
            elif item["kind"] == "columns":
                sheet.Range(item["addr"]).EntireColumn.Hidden = True
                verified = all(sheet.Columns(index).Hidden is True for index in range(item["first"], item["last"] + 1))
            else:
                sheet.Range(item["addr"]).EntireColumn.ColumnWidth = item["value"]
                verified = all(sheet.Columns(index).Hidden is False and abs(float(sheet.Columns(index).ColumnWidth) - item["value"]) < 0.05
                               for index in range(item["first"], item["last"] + 1))
            if not verified:
                raise ValueError(f"{item['sheet']}!{item['addr']} 未能恢复，本文件没有保存")
        if not any(int(sheet.Visible) == VISIBLE for sheet in book.Sheets):
            raise ValueError("至少需要保留一张可见工作表，本文件没有保存")
        check_versions(opened_stamp)
        book.Save()
    finally:
        book.Close(SaveChanges=False)
    return [{**item, "status": "ok", "reason": "已恢复备份列宽（仍为可见列）" if item["kind"] == "narrow" else "已恢复备份隐藏状态"}
            for item in ready]


def _run(root: Path, paths: list[Path], old_root: Path, old_paths: list[Path], options: dict) -> None:
    files, excel, initialized = [], None, False
    try:
        wf.pythoncom.CoInitialize()
        initialized = True
        inferred: set[str] = set()
        pairs = pair_files(old_root, old_paths, [workbook_rel_name(path, root) for path in paths], root, inferred=inferred)
        for index, path in enumerate(paths):
            name = workbook_rel_name(path, root)
            entry = {"name": name, "path": str(path), "cells": [], "error": None}
            files.append(entry)
            with _LOCK:
                _STATUS.update(done=index, current=name, stage="preview")
            try:
                old_path = pairs[name]
                if old_path is None:
                    raise ValueError("找不到唯一对应的备份文件，请选到对应子文件夹")
                entry["old"] = str(old_path)
                if excel is None:
                    excel = start_background_excel()
                entry.update(plan_file(excel, path, old_path, options))
                if name in inferred:
                    require_manual_file_check(entry)
                if any(item["action"] == "restore" for item in entry["cells"]):
                    with _LOCK:
                        _STATUS["stage"] = "apply"
                    outcomes = {(item["sheet"], item["kind"], item["addr"]): item for item in apply_file(excel, entry)}
                    entry["cells"] = [outcomes.get((item["sheet"], item["kind"], item["addr"]), item) for item in entry["cells"]]
            except Exception as error:
                entry["error"] = str(error)
                debug_log.exception(error, f"隐藏恢复 {name}")
            for item in entry["cells"]:
                debug_log.event("🙈 隐藏恢复", 位置=f"{item['sheet']}!{item['addr']}", 类型=item["kind"],
                                结果=entry["error"] or item["reason"])
            count = sum(item.get("status") == "ok" for item in entry["cells"])
            wf.log_msg(f"🙈 {name}：已恢复 {count} 处" + (f"；{entry['error']}" if entry["error"] else ""))
    except Exception as error:
        debug_log.exception(error, "隐藏恢复任务")
        files.append({"name": "任务", "path": "", "cells": [], "error": str(error)})
    finally:
        try:
            quit_background_excel(excel)
        finally:
            try:
                if initialized:
                    wf.pythoncom.CoUninitialize()
            finally:
                with _LOCK:
                    _STATUS.update(running=False, done=len(paths), current="", applied={
                        "folder": str(root), "old_folder": str(old_root), "options": options, "files": files})
                with wf.STATE_LOCK:
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)


def repair(folder: str, old_folder: str, options: dict | None = None) -> None:
    options = options_for(options)
    root, paths = list_workbooks(folder)
    old_root, old_paths = list_workbooks(old_folder)
    if root == old_root or root in old_root.parents or old_root in root.parents:
        raise ValueError("当前文件夹与备份文件夹不能相同或互相包含，请分别选择两份文件所在的目录")
    with wf.STATE_LOCK, _LOCK:
        from core import bl_links, link_repair, mark_finder
        if (wf.STATE.get("busy") or _STATUS.get("running") or link_repair.status().get("running")
                or mark_finder.status().get("running") or bl_links.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束")
        job = wf.STATE.get("job") or {}
        if job.get("pending_adjustment") or job.get("deferred_adjustments"):
            raise ValueError("还有未完成的文件写入，请先继续或放弃上次写入")
        wf.STATE.update(busy=True, maintenance="hidden_repair")
        _STATUS.clear()
        _STATUS.update(running=True, phase="repair", stage="preview", folder=str(root), old_folder=str(old_root),
                       options=options, total=len(paths), done=0, current="")

    def run() -> None:
        with debug_log.side_task("hidden_repair", {"folder": str(root), "old_folder": str(old_root), "options": options}):
            _run(root, paths, old_root, old_paths, options)

    try:
        threading.Thread(target=run, name="扫描并修复隐藏页", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
        raise


def status() -> dict:
    with _LOCK:
        return copy.deepcopy(_STATUS)
