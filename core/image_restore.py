"""修复图片变形 / 红框错位：对照参考文件夹里的原文件，把截图和上面的标注放回原来的大小和相对位置（用户 2026-11-11 要）。

起因：以前修切字时加高行、加宽列，Excel 的图形默认「随单元格移动和改变大小」——
截图中间的行加高，截图被拉长变形；红框挂在另一行，只会整体挪动，不跟截图按比例变，于是框不住原来的按钮。

做法（每张表）：
- 处理后文件和参考文件按表名（找不到再按代码名）配对；图形按「名字 + 类型」配对，任一边重名的不碰。
- 截图 = 图片（嵌入 / 链接）；整个落在截图范围里的其他图形（按参考文件里的位置判断）算这张截图的标注，落在几张里归最小那张。
- 截图和不在截图上的图形：左上角不动（行列变了它本来就该跟着走），大小改回参考文件里的。
- 标注：大小改回参考的；位置 = 截图现在的左上角 + 参考文件里它相对截图左上角的偏移。
- 只改真的变了的（大小差 > 1 磅，或标注相对截图偏了 > 1 磅）；行高、列宽、格子内容一律不动。
- 改完回读核对，有一个没对上整份不保存。
"""
from __future__ import annotations

import copy
import threading
import time
from pathlib import Path
from typing import Any

from core import debug_log, file_problem, progress_log, workflow as wf
from core.backup_compare import pair_files, require_manual_file_check
from core.common import workbook_locked, workbook_rel_name
from core.excel_ops import open_workbook_checked, quit_background_excel, start_background_excel
from core.link_repair import fingerprint
from core.shape_geometry import apply_sheet, describe_change, plan_sheet, read_shapes

_LOCK = threading.RLock()
_STATUS: dict[str, Any] = {"running": False}


def list_workbooks(folder: str) -> tuple[Path, list[Path]]:
    from core import hidden_repair
    return hidden_repair.list_workbooks(folder)


def read_book(book: Any) -> list[dict]:
    sheets = []
    for sheet in book.Worksheets:
        try:
            code = str(sheet.CodeName or "")
        except Exception:
            code = ""
        sheets.append({"name": str(sheet.Name), "code": code, "shapes": read_shapes(sheet)})
    return sheets


def pair_book_sheets(current: list[dict], old: list[dict]) -> list[tuple[dict, dict]]:
    """按表名，找不到再按代码名（两边都唯一才算）"""
    by_name = {s["name"]: s for s in old}
    codes: dict[str, list[dict]] = {}
    for s in old:
        if s["code"]:
            codes.setdefault(s["code"], []).append(s)
    pairs = []
    for sheet in current:
        match = by_name.get(sheet["name"])
        if match is None and sheet["code"] and len(codes.get(sheet["code"], [])) == 1:
            match = codes[sheet["code"]][0]
        if match is not None:
            pairs.append((sheet, match))
    return pairs


# ---------- 一份 Excel ----------

def fix_workbook(excel: Any, path: Path, old_path: Path, on_sheet=None) -> dict:
    if path.resolve() == old_path.resolve():
        raise ValueError("处理后的文件和参考文件是同一个，请分别选两个文件夹")
    if workbook_locked(path):
        raise ValueError("文件正在 Excel 中打开，请保存并关闭后再修复图片")
    old_book = open_workbook_checked(excel, str(old_path.resolve()), ReadOnly=True, UpdateLinks=0)
    try:
        old = read_book(old_book)
    finally:
        old_book.Close(SaveChanges=False)
    stamp = fingerprint(path)
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0)
    sheets: list[dict] = []
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开，没有修改")
        current = read_book(book)
        pairs = pair_book_sheets(current, old)
        for index, (cur, ref) in enumerate(pairs, 1):
            if on_sheet:
                on_sheet(index, len(pairs), cur["name"])
            changes = plan_sheet(cur["shapes"], ref["shapes"])
            if not changes:
                continue
            details = [describe_change(item) for item in changes]  # 改之前写，before 是原样
            bad = apply_sheet(book.Worksheets(cur["name"]), changes)
            debug_log.event("🖼️ 修复图片", 文件=path.name, 表=cur["name"], 图形=details, 没对上=bad or "无")
            if bad:
                raise ValueError(f"{cur['name']} 有 {len(bad)} 个图形改完对不上（{'、'.join(bad[:5])}），本文件没有保存")
            sheets.append({"sheet": cur["name"], "count": len(changes), "details": details[:30],
                           "pictures": sum(c["kind"] == "picture" for c in changes),
                           "overlays": sum(c["kind"] == "overlay" for c in changes)})
        if sheets:
            if fingerprint(path) != stamp:
                raise ValueError("修复期间文件被别人改过，没有保存")
            book.Save()
            debug_log.saved_excel("修复图片", path)
    finally:
        book.Close(SaveChanges=False)
    return {"count": sum(s["count"] for s in sheets), "sheets": sheets}


def describe(entry: dict) -> str:
    if entry.get("error"):
        return entry["error"]
    return f"放回 {entry['count']} 个图形（{len(entry['sheets'])} 张表）" if entry.get("count") else "图片都没变"


# ---------- 后台任务 ----------

def _run(root: Path, paths: list[Path], old_root: Path, old_paths: list[Path]) -> None:
    files: list[dict] = []
    excel, initialized = None, False
    guard = file_problem.Guard("修复图片")
    try:
        wf.pythoncom.CoInitialize()
        initialized = True
        inferred: set[str] = set()
        pairs = pair_files(old_root, old_paths, [workbook_rel_name(p, root) for p in paths], root, inferred=inferred)
        for index, path in enumerate(paths):
            name = workbook_rel_name(path, root)
            while True:
                entry: dict[str, Any] = {"name": name, "path": str(path), "error": None, "count": 0, "sheets": []}
                with _LOCK:
                    _STATUS.update(done=index, current=name)
                    progress_log.step(_STATUS, "正在打开 Excel")
                started = time.perf_counter()
                try:
                    if guard.file(name):  # 进度框能「跳过这一份」
                        excel = None
                    old_path = pairs.get(name)
                    if old_path is None:
                        raise ValueError("参考文件夹里找不到唯一对应的原文件")
                    entry["old"] = str(old_path)
                    if name in inferred:
                        require_manual_file_check(entry)
                        break
                    if excel is None:
                        excel = guard.watch(start_background_excel())

                    def on_sheet(i: int, total: int, sheet: str) -> None:
                        guard.progress()
                        with _LOCK:
                            progress_log.step(_STATUS, f"正在对照第 {i} / {total} 张表：{sheet}")

                    entry.update(fix_workbook(excel, path, old_path, on_sheet))
                except Exception as error:
                    debug_log.exception(error, f"修复图片 {name}")
                    retry, entry["error"] = guard.problem(error, path, name, root,
                                                          close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {name}")
                        continue
                break
            files.append(entry)
            debug_log.event("🖼️ 修复图片一份", 文件=name, 图形=entry.get("count", 0), 错误=entry["error"] or "",
                            秒=round(time.perf_counter() - started, 2))
            wf.log_msg(f"🖼️ {name}：{describe(entry)}")
            with _LOCK:
                progress_log.add(_STATUS, name, not entry["error"], describe(entry), path=str(path))
    except Exception as error:
        debug_log.exception(error, "修复图片任务")
        files.append({"name": "任务", "path": "", "error": str(error), "count": 0, "sheets": []})
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
                                   result={"folder": str(root), "old_folder": str(old_root), "files": files})
                with wf.STATE_LOCK:
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)


def start(folder: str, old_folder: str) -> None:
    root, paths = list_workbooks(folder)
    old_root, old_paths = list_workbooks(old_folder)
    if root == old_root or root in old_root.parents or old_root in root.parents:
        raise ValueError("处理后的文件夹和参考文件夹不能相同或互相包含")
    with wf.STATE_LOCK, _LOCK:
        from core import bl_links, font_fix, hidden_repair, link_repair, mark_finder
        if (wf.STATE.get("busy") or _STATUS.get("running") or link_repair.status().get("running")
                or mark_finder.status().get("running") or hidden_repair.status().get("running")
                or bl_links.status().get("running") or font_fix.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束")
        job = wf.STATE.get("job") or {}
        if job.get("pending_adjustment") or job.get("deferred_adjustments"):
            raise ValueError("还有未完成的文件写入，请先继续或放弃上次写入")
        wf.STATE.update(busy=True, maintenance="image_restore")
        _STATUS.clear()
        _STATUS.update(running=True, folder=str(root), old_folder=str(old_root), total=len(paths), done=0, current="")

    def run() -> None:
        with debug_log.side_task("image_restore", {"folder": str(root), "old_folder": str(old_root), "files": len(paths)}):
            _run(root, paths, old_root, old_paths)

    try:
        threading.Thread(target=run, name="修复图片", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
        raise


def status() -> dict:
    with _LOCK:
        return copy.deepcopy(_STATUS)
