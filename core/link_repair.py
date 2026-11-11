"""独立的引用错误扫描、备份对照与修复。备份始终只读。"""
from __future__ import annotations

import copy
import hashlib
import threading
from collections.abc import Callable, Iterator
from contextlib import AbstractContextManager, contextmanager, nullcontext
from itertools import product
from pathlib import Path
from uuid import uuid4
from typing import Any

from core import debug_log, file_problem, formula_links as links, progress_log, workflow as wf, sheet_history
from core import workbook_scope
from core.backup_compare import pair_files, pair_sheets, require_manual_file_check
from core.reference_folders import ReferenceFolders, ReferenceIndex, prepare as prepare_references, unmatched
from core.common import workbook_locked, workbook_rel_name
from core.excel_ops import open_workbook_checked, start_background_excel, quit_background_excel

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


def _sheet_info(sheet: Any) -> dict:
    used = sheet.UsedRange
    first_row, first_col = int(used.Row), int(used.Column)
    last_row = min(first_row + int(used.Rows.Count) - 1, 200)
    last_col = min(first_col + int(used.Columns.Count) - 1, first_col + 255)
    words: set[str] = set()
    if first_row <= last_row:
        values = sheet.Range(sheet.Cells(first_row, first_col), sheet.Cells(last_row, last_col)).Value2
        matrix = values if isinstance(values, tuple) else ((values,),)
        words = {_text(value).strip() for row in matrix for value in row
                 if _text(value).strip() and value not in links.ERROR_VALUES}
    return {"name": sheet.Name, "code": sheet.CodeName, "words": sorted(words)}


def _pair_sheets(current: list[dict], old: list[dict], history=None, *, ambiguous=None) -> dict:
    return pair_sheets(current, old, title=lambda item: item["name"],
                       code=lambda item: item["code"], words=lambda item: set(item["words"]),
                       history=history, ambiguous=ambiguous)


def _anchors(cell: Any) -> list[dict]:
    result = []
    for row in range(max(1, int(cell.Row) - 1), int(cell.Row) + 2):
        for col in range(max(1, int(cell.Column) - 2), int(cell.Column) + 3):
            other = cell.Worksheet.Cells(row, col)
            if not bool(other.HasFormula) and other.Value2 not in (None, "") and not links.is_error(other):
                result.append({"addr": links.address(other), "value": other.Value2})
    return result


def scan_workbook(book: Any) -> dict:
    book.Application.CalculateFull()
    sheets = [_sheet_info(sheet) for sheet in book.Worksheets]
    sheet_names = [sheet.Name for sheet in book.Sheets]
    hits, seen = [], set()
    for hit in links.link_errors(book):
        cell = book.Worksheets(hit["sheet"]).Range(hit["addr"])
        area, kind = links.region(cell)
        addr = links.address(area)
        key = (hit["sheet"], addr)
        if key in seen:
            continue
        seen.add(key)
        hits.append({**hit, "addr": addr, "anchor": links.address(cell), "kind": kind,
                     "formulas": area.Formula, "anchors": _anchors(cell)})
    return {"sheets": sheets, "sheet_names": sheet_names, "hits": hits}


def scan_file(excel: Any, path: Path) -> dict:
    stamp = fingerprint(path)
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0, ReadOnly=True)
    try:
        result = scan_workbook(book)
    finally:
        book.Close(SaveChanges=False)
    if stamp != fingerprint(path):
        raise ValueError("文件在扫描期间发生变化，请重新扫描")
    return {"path": str(path), "fingerprint": stamp, **result}


def _same(first: Any, second: Any) -> bool:
    if isinstance(first, (tuple, list)) and isinstance(second, (tuple, list)):
        return len(first) == len(second) and all(_same(a, b) for a, b in zip(first, second))
    if type(first) is type(second):
        return first == second
    return (isinstance(first, (int, float)) and not isinstance(first, bool)
            and isinstance(second, (int, float)) and not isinstance(second, bool) and first == second)


def _formula_identity(value: Any) -> Any:
    if isinstance(value, (tuple, list)):
        return tuple(_formula_identity(item) for item in value)
    if not isinstance(value, str) or not value.startswith("="):
        return value
    result, previous = [], None
    for item in links.tokens(value):
        # Excel 删表后内存中可能是 #REF!A1，保存重开会折叠为 #REF!。
        # 只折叠相邻的失效地址，不改字符串常量或 #REF! + A1 中仍有效的 A1。
        if not (previous and previous.subtype == "ERROR" and previous.value.upper() == "#REF!"
                and item.type == "OPERAND" and item.subtype == "RANGE"):
            result.append(item.value)
        previous = item
    return "=" + "".join(result)


class _SourceMismatch(ValueError):
    """已读取候选，发现与当前格不符的证据。"""


def _source_check(old: Any, cell: dict, source: str, how: str) -> dict:
    item = {**cell, "old_sheet": source, "match": how, "source": f"{source}!{cell['addr']}"}
    try:
        if not item["formula"]:
            raise _SourceMismatch("当前是固定的错误值，缺少公式线索，不能确认原位置")
        sheet = old.Worksheets(source)
        original_cell = sheet.Range(item["anchor"])
        original_formula = links.formula(original_cell)
        if not original_formula:
            raise _SourceMismatch("备份同一位置没有公式，不能确认对应关系")
        if links.broken_formula(original_formula):
            raise ValueError("备份的原公式本身已有失效引用，不能作为恢复来源")
        for anchor in item["anchors"]:
            other = sheet.Range(anchor["addr"])
            if bool(other.HasFormula) or not _same(other.Value2, anchor["value"]):
                raise _SourceMismatch("附近的内容和备份不同，可能移动过行列，请人工检查")
        original = links.capture(original_cell)
        if original["addr"] != item["addr"] or original["kind"] != item["kind"]:
            raise _SourceMismatch("新旧数组或溢出区域不同，请人工检查，避免覆盖其他格子")
        item.update(check="ready", value=original["value"], rows=original["rows"], cols=original["cols"],
                    source_formula=original["formula"], display=_text(original_cell.Text))
    except _SourceMismatch as error:
        item.update(check="mismatch", reason=str(error))
    except Exception as error:
        item.update(check="unverified", reason=str(error))
    return item


def _sheet_mappings(sheets: list[dict], matched: dict, ambiguous: dict[str, list[str]]) -> tuple[Iterator[dict[str, str]], str]:
    fixed = {name: match[0]["name"] for name, match in matched.items()}
    if not ambiguous:
        return iter((fixed,)), ""
    if fixed.keys() | ambiguous.keys() != {sheet["name"] for sheet in sheets}:
        return iter(()), "还有工作表缺少可核对的来源，无法完整模拟删表"
    used = set(fixed.values())
    options = [[source for source in candidates if source not in used] for candidates in ambiguous.values()]
    return ({**fixed, **dict(zip(ambiguous, sources))} for sources in product(*options)
            if len(set(sources)) == len(sources)), ""


def _simulate_mapping(old: Any, mapping: dict[str, str], ready: list[dict]) -> None:
    keep = set(mapping.values())
    for sheet in list(old.Sheets):
        if sheet.Name not in keep:
            sheet.Delete()
    renames = [(old.Worksheets(source), name) for name, source in mapping.items() if source != name]
    for sheet, _ in renames:
        sheet.Name = "CHE_" + uuid4().hex[:20]
    for sheet, name in renames:
        sheet.Name = name
    old.Application.CalculateFull()
    for item in ready:
        try:
            sheet = old.Worksheets(item["sheet"])
            cell = sheet.Range(item["anchor"])
            if _formula_identity(sheet.Range(item["addr"]).Formula) != _formula_identity(item["formulas"]):
                raise _SourceMismatch("模拟删除旧表后的公式与当前不一致，不能确认原引用")
            if not (links.broken_formula(links.formula(cell)) or links.is_ref_error(cell)):
                raise _SourceMismatch("备份模拟删表后未出现相同引用错误，不能确认错误来源")
            item.update(check="confirmed", action="restore", reason="原公式与删表后的损坏形态一致，恢复备份计算结果为固定值")
        except _SourceMismatch as error:
            item.update(check="mismatch", reason=str(error))
        except Exception as error:
            item.update(check="unverified", reason=str(error))


@contextmanager
def _open_reference(excel: Any, path: str):
    book = open_workbook_checked(excel, path, UpdateLinks=0, ReadOnly=True)
    try:
        yield book
    finally:
        book.Close(SaveChanges=False)


def _check_mapping(open_old: Callable[[], AbstractContextManager[Any]], mapping: dict[str, str],
                   cells: list[dict], checks: list[dict],
                   attempts: list[list[dict]], *, multiple: bool) -> None:
    selected = {i: checks[i][mapping[cell["sheet"]]].copy() for i, cell in enumerate(cells)
                if cell["sheet"] in mapping}
    ready = [item for item in selected.values() if item["check"] == "ready"]
    if ready:
        with open_old() as old:
            try:
                _simulate_mapping(old, mapping, ready)
            except Exception as error:
                if not multiple:
                    raise
                for item in ready:
                    if item["check"] == "ready":
                        item.update(check="unverified", reason=f"无法模拟删表和改名（{error}）")
    for i, item in selected.items():
        attempts[i].append(item)


def _candidate_details(attempts: list[dict]) -> str:
    details = dict.fromkeys(f"{item['source']} {item['source_formula']} → {item['value']!r}"
                            if item["check"] == "confirmed" else f"{item['source']}（{item['reason']}）"
                            for item in attempts)
    return " / ".join(details)


def _candidate_result(cell: dict, attempts: list[dict], ambiguity: str) -> dict:
    if not attempts:
        reason = "备份工作表无法唯一配对，请人工检查"
        return {**cell, "reason": f"{reason}（{ambiguity}）" if ambiguity else reason}
    if len(attempts) == 1:
        return {key: value for key, value in attempts[0].items() if key != "check"}

    confirmed = [item for item in attempts if item["check"] == "confirmed"]
    unverified = any(item["check"] == "unverified" for item in attempts)
    if confirmed and not unverified:
        first = confirmed[0]
        if all(_same(first["value"], item["value"]) and first["rows"] == item["rows"]
               and first["cols"] == item["cols"] for item in confirmed):
            result = {key: value for key, value in first.items() if key != "check"}
            sources = list(dict.fromkeys(item["source"] for item in confirmed))
            if len(sources) > 1:
                formulas = dict.fromkeys(item["source_formula"] for item in confirmed)
                result.update(old_sheet="", source="、".join(sources), source_formula=" / ".join(formulas),
                              reason="通过模拟核对的候选恢复值一致，恢复为固定值")
            elif first["match"] == "candidate":
                result["reason"] = "候选核对后只剩一个来源，恢复备份计算结果为固定值"
            return result
    if unverified:
        reason = "有候选未能完成核对，请人工检查"
    elif confirmed:
        reason = "多个候选通过核对，但恢复值不同，请人工检查"
    else:
        reason = "所有候选均与当前错误不符，请人工检查"
    return {**cell, "old_sheet": "", "source": _candidate_details(attempts),
            "reason": f"备份工作表无法唯一配对（{reason}）"}


def _reconcile_candidates(cells: list[dict], results: list[dict], attempts: list[list[dict]], count: int) -> None:
    """用已通过核对的格子缩小配对范围，同一批恢复必须有共同的方案。"""
    support = {i: {index for index, item in enumerate(choices) if item["check"] == "confirmed"}
               for i, choices in enumerate(attempts) if results[i]["action"] == "restore"}
    if not support:
        return
    compatible = sorted(set.intersection(*support.values()))
    if compatible:
        if len(compatible) == count:
            return
        for i, choices in enumerate(attempts):
            if not choices or results[i]["action"] == "restore":
                continue
            surviving = [choices[index] for index in compatible]
            if all(item["check"] == "confirmed" for item in surviving):
                refined = _candidate_result(cells[i], surviving, "")
                if refined["action"] == "restore":
                    results[i] = {**refined, "match": "candidate", "reason": "结合其他错误格核对配对后，恢复备份计算结果为固定值"}
        return
    for i, options in support.items():
        if len(options) < count:
            results[i] = {**cells[i], "old_sheet": "", "source": _candidate_details(attempts[i]),
                          "reason": "备份工作表无法唯一配对（不同单元格的核对结果指向互相冲突的配对方案）"}


def preview_file(excel: Any, file: dict, old_path: Path) -> dict:
    """只读核对候选原值，再逐个模拟可行的删表与改名方案。"""
    if old_path.resolve() == Path(file["path"]).resolve():
        raise ValueError("配对到的是当前文件本身，请选择原始备份")
    stamp = fingerprint(old_path)
    path = str(old_path.resolve())
    cells = [{**hit, "action": "skip", "reason": "", "source_formula": "", "display": ""} for hit in file["hits"]]
    attempts: list[list[dict]] = [[] for _ in cells]
    with _open_reference(excel, path) as old:
        old.Application.CalculateFull()
        old_sheets = [_sheet_info(sheet) for sheet in old.Worksheets]
        history = sheet_history.for_paths(file["path"], old_path, file.get("sheet_names", [s["name"] for s in file["sheets"]]),
                                          [s.Name for s in old.Sheets], old_hash=stamp)
        ambiguous: dict[str, list[str]] = {}
        matched = _pair_sheets(file["sheets"], old_sheets, history, ambiguous=ambiguous)
        mappings, ambiguity = _sheet_mappings(file["sheets"], matched, ambiguous)
        first_mapping = next(mappings, None)
        if first_mapping is None and not ambiguity:
            ambiguity = "候选工作表之间没有可行的一对一配对"
        reserved = {match[0]["name"] for match in matched.values()}
        checks = []
        for cell in cells:
            name = cell["sheet"]
            sources = ([matched[name][0]["name"]] if name in matched
                       else [source for source in ambiguous.get(name, []) if source not in reserved])
            how = matched[name][1] if name in matched else "candidate"
            checks.append({source: _source_check(old, cell, source, how) for source in sources} if first_mapping is not None else {})
            if first_mapping is None and name in ambiguous:
                cell.update(old_sheet="", source="、".join(ambiguous[name]))
        if first_mapping is not None:
            if any(item["check"] == "ready" for candidates in checks for item in candidates.values()):
                _check_mapping(lambda: nullcontext(old), first_mapping, cells, checks, attempts, multiple=bool(ambiguous))
            else:
                attempts = [list(candidates.values()) for candidates in checks]
                mappings = iter(())
    count = int(first_mapping is not None)
    for mapping in mappings:
        _check_mapping(lambda: _open_reference(excel, path), mapping, cells, checks, attempts, multiple=True)
        count += 1
    results = [_candidate_result(cell, options, ambiguity) for cell, options in zip(cells, attempts)]
    if count > 1:
        _reconcile_candidates(cells, results, attempts, count)
    if stamp != fingerprint(old_path):
        raise ValueError("备份在对比期间发生变化，请重新对比")
    for item in results:
        debug_log.event("🔗 引用对比", 位置=f"{item['sheet']}!{item['addr']}",
                        原公式=item["source_formula"], 动作=item["action"], 原因=item["reason"], 恢复值=item["display"])
    return {**file, "old": str(old_path), "old_fingerprint": stamp, "cells": results}


def apply_workbook(book: Any, cells: list[dict]) -> list[dict]:
    results = []
    book.Application.CalculateFull()
    before = links.error_positions(book)
    # 写入前检查全部目标；任何一处变动都不保存本文件。
    for item in cells:
        sheet = book.Worksheets(item["sheet"])
        if not _same(sheet.Range(item["addr"]).Formula, item["formulas"]):
            raise ValueError(f"{item['sheet']}!{item['addr']} 的公式已变，请重新扫描")
    direct_first = sorted(cells, key=lambda item: not links.broken_formula(item["formula"]))
    for item in direct_first:
        sheet = book.Worksheets(item["sheet"])
        cell = sheet.Range(item["anchor"])
        if not (links.broken_formula(links.formula(cell)) or links.is_ref_error(cell)):
            results.append({**item, "status": "resolved", "reason": "上游引用修复后已正常，保留现有公式"})
            continue
        links.write_values(sheet, item)
        reason = item["reason"] if item.get("match") == "candidate" else "已恢复备份计算结果为固定值"
        results.append({**item, "status": "ok", "reason": reason})
    book.Application.CalculateFull()
    after = links.error_positions(book)
    if after - before:
        raise ValueError("修复后出现新的引用错误，本文件没有保存")
    for item in results:
        if (item["sheet"], item["anchor"]) in after:
            raise ValueError(f"{item['sheet']}!{item['addr']} 修复后仍然报错，本文件没有保存")
    return results


def apply_file(excel: Any, file: dict) -> list[dict]:
    path, old_path = Path(file["path"]), Path(file["old"])
    if fingerprint(path) != file["fingerprint"] or fingerprint(old_path) != file["old_fingerprint"]:
        raise ValueError("当前文件或备份在预览后发生变化，请重新扫描并对比")
    if workbook_locked(path):
        raise ValueError("文件正在 Excel 中打开，请保存并关闭后再修复")
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开，没有修改")
        results = apply_workbook(book, file["cells"])
        if any(item["status"] == "ok" for item in results):
            book.Save()
            debug_log.saved_excel("修复引用错误", path)
    finally:
        book.Close(SaveChanges=False)
    return results


def _progress(done: int, name: str) -> None:
    with _LOCK:
        _STATUS.update(done=done, current=name)


def _run(phase: str, folder: Path, paths: list, old_root: Path | ReferenceIndex | None = None, old_paths: list | None = None) -> None:
    excel = None
    files = []
    com = False
    guard = file_problem.Guard({"scan": "扫描引用错误", "preview": "对比备份", "apply": "修复引用错误",
                                "repair": "扫描并修复引用错误"}[phase])
    try:
        wf.pythoncom.CoInitialize()
        com = True
        excel = guard.watch(start_background_excel())
        names = [workbook_rel_name(source, folder) if isinstance(source, Path) else source["name"] for source in paths]
        inferred: set[str] = set()
        pairs = pair_files(old_root, old_paths, names, folder, inferred=inferred) if phase in {"preview", "repair"} else {}
        for index, source in enumerate(paths):
            while True:  # 出错时问用户：「重试」重开 Excel 再做这份，「跳过」记下错误
                name = workbook_rel_name(source, folder) if isinstance(source, Path) else source["name"]
                _progress(index, name)
                entry = {"name": name, "path": str(source) if isinstance(source, Path) else source["path"], "error": None}
                old_path = pairs.get(name) if phase in {"repair", "preview"} else source.get("old") if isinstance(source, dict) else None
                if old_path:
                    entry["old"] = str(old_path)
                try:
                    if guard.file(name):  # 进度框能「跳过这一份」
                        excel = None
                    if excel is None:
                        excel = guard.watch(start_background_excel())
                    if phase == "repair":
                        with _LOCK:
                            _STATUS["stage"] = "scan"
                        entry.update(scan_file(excel, source))
                        entry["cells"] = [{**hit, "status": "skip", "reason": "未修复"} for hit in entry["hits"]]
                        if entry["hits"]:
                            if not pairs[name]:
                                raise ValueError(unmatched(old_root, name, "找不到唯一对应的备份文件，请选到对应子文件夹"))
                            with _LOCK:
                                _STATUS["stage"] = "preview"
                            entry.update(preview_file(excel, entry, pairs[name]))
                            if name in inferred:
                                require_manual_file_check(entry)
                            ready = [cell for cell in entry["cells"] if cell["action"] == "restore"]
                            if ready:
                                with _LOCK:
                                    _STATUS["stage"] = "apply"
                                outcomes = {(cell["sheet"], cell["addr"]): cell
                                            for cell in apply_file(excel, {**entry, "cells": ready})}
                                entry["cells"] = [outcomes.get((cell["sheet"], cell["addr"]), cell) for cell in entry["cells"]]
                            for item in entry["cells"]:
                                debug_log.event("🔗 引用修复", 位置=f"{item['sheet']}!{item['addr']}", 结果=item["reason"])
                    elif phase == "scan":
                        entry.update(scan_file(excel, source))
                    elif phase == "preview":
                        if not pairs[name]:
                            raise ValueError("找不到唯一对应的备份文件，请选到对应子文件夹")
                        if fingerprint(Path(source["path"])) != source["fingerprint"]:
                            raise ValueError("当前文件在扫描后发生变化，请重新扫描")
                        entry.update(preview_file(excel, source, pairs[name]))
                        if name in inferred:
                            require_manual_file_check(entry)
                    else:
                        entry["cells"] = apply_file(excel, source)
                        for item in entry["cells"]:
                            debug_log.event("🔗 引用修复", 位置=f"{item['sheet']}!{item['addr']}", 结果=item["reason"])
                except Exception as error:
                    debug_log.exception(error, f"引用修复 {phase} {name}")
                    retry, entry["error"] = guard.problem(error, entry["path"], name, folder,
                                                          close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {name}")
                        continue
                break
            files.append(entry)
            repaired = sum(cell.get("status") in {"ok", "resolved"} for cell in entry.get("cells", []))
            message = entry["error"] or (f"已恢复正常 {repaired} 处，未修复 {len(entry.get('cells', [])) - repaired} 处" if phase == "repair" else
                                         f"发现引用错误 {len(entry.get('hits', []))} 处" if phase == "scan" else "对比完成" if phase == "preview" else "修复完成")
            wf.log_msg(f"🔗 {name}：{message}")
            with _LOCK:  # 网页实时进度框：做完一份就列一条结果（没修好的算「!」）
                progress_log.add(_STATUS, name, not entry["error"] and (phase != "repair" or repaired == len(entry.get("cells", []))),
                                 message, path=entry["path"])
    except Exception as error:
        debug_log.exception(error, "引用错误任务")
        files.append({"name": "任务", "path": "", "error": str(error)})
    finally:
        try:
            if excel is not None:
                quit_background_excel(excel)
        finally:
            guard.stop()
            if com:
                wf.pythoncom.CoUninitialize()
            with _LOCK:
                key = {"scan": "result", "preview": "plan", "apply": "applied", "repair": "applied"}[phase]
                _STATUS.update(running=False, done=len(paths), current="")
                _STATUS[key] = {"folder": str(folder), "old_folder": str(old_root or ""), "files": files}
            with wf.STATE_LOCK:
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)


def _launch(phase: str, folder: Path, paths: list, old_root: Path | ReferenceIndex | None = None, old_paths: list | None = None) -> None:
    with wf.STATE_LOCK, _LOCK:
        from core import bl_links, hidden_repair, mark_finder
        if (wf.STATE.get("busy") or _STATUS.get("running") or mark_finder.status().get("running")
                or hidden_repair.status().get("running") or bl_links.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束")
        job = wf.STATE.get("job") or {}
        if job.get("pending_adjustment") or job.get("deferred_adjustments"):
            raise ValueError("还有未完成的文件写入，请先继续或放弃上次写入")
        wf.STATE["busy"] = True
        wf.STATE["maintenance"] = "link_" + phase
        _STATUS.update(running=True, phase=phase, stage="scan", folder=str(folder), old_folder=str(old_root or ""),
                       total=len(paths), done=0, current="", log=[])
        for key in {"scan": ("result", "plan", "applied"), "preview": ("plan", "applied"), "apply": ("plan",),
                    "repair": ("result", "plan", "applied")}[phase]:
            _STATUS.pop(key, None)
    def run() -> None:
        with debug_log.side_task("link_" + phase, {"folder": str(folder), "old_folder": str(old_root or ""), "files": len(paths)}):
            from core import operation_history
            kind = "links" if phase == "repair" else None
            ids = operation_history.begin(folder, [kind]) if kind else []
            try:
                _run(phase, folder, paths, old_root, old_paths)
                result = status().get("applied", {})
                operation_history.finish(ids, "failed" if any(f.get("error") for f in result.get("files", [])) else "completed")
            except Exception:
                operation_history.finish(ids, "failed")
                raise
    try:
        threading.Thread(target=run, name="引用错误" + phase, daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
        raise


def start(folder: str) -> None:
    root, paths = wf.list_workbooks(folder)
    _launch("scan", root, paths)


def repair(folder: str, old_folder: str | ReferenceFolders) -> None:
    """一次后台任务完成扫描、对比和修复，页面关闭后仍继续。"""
    root, paths = wf.list_workbooks(folder)
    old_root, old_paths = prepare_references(root, old_folder, wf.list_workbooks, wf.EXCEL_EXTENSIONS | {".xlsb"})
    _launch("repair", root, paths, old_root, old_paths)


def preview(old_folder: str) -> None:
    with _LOCK:
        result = copy.deepcopy(_STATUS.get("result"))
    if not result:
        raise ValueError("请先扫描引用错误")
    old_root, old_paths = wf.list_workbooks(old_folder)
    root = Path(result["folder"]).resolve()
    if old_root == root:
        raise ValueError("备份文件夹不能与当前文件夹相同")
    if workbook_scope.overlaps_reference(root, old_root):
        raise ValueError("原始备份文件夹不能与批量处理路径重叠")
    files = [file for file in result["files"] if file.get("hits") and not file["error"]]
    if not files:
        raise ValueError("没有可对比的引用错误")
    _launch("preview", Path(result["folder"]), files, old_root, old_paths)


def apply() -> None:
    with _LOCK:
        plan = copy.deepcopy(_STATUS.get("plan"))
    if not plan:
        raise ValueError("请先对比备份并查看修复清单")
    files = [{**file, "cells": [item for item in file.get("cells", []) if item["action"] == "restore"]}
             for file in plan["files"] if not file["error"]]
    files = [file for file in files if file["cells"]]
    if not files:
        raise ValueError("没有可以自动修复的单元格")
    _launch("apply", Path(plan["folder"]), files, Path(plan["old_folder"]))


def status() -> dict:
    with _LOCK:
        return copy.deepcopy(_STATUS)
