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

页面分两步（用户 2026-11-11 要）：先「找出受影响的图片」——两份都只读打开，不改任何文件，每处列出
原版 / 新版的局部截图（把那块区域临时设成打印区域导出 PDF 再转图片）、新版 Excel、对照打开、PDF；
用户看完再点「修复这份」或「全部修复」，也可以自己打开三个文件处理。
"""
from __future__ import annotations

import copy
import hashlib
import os
import tempfile
import threading
import time
from pathlib import Path
from typing import Any

from core import debug_log, file_problem, fix_images, progress_log, workflow as wf
from core.backup_compare import pair_files, require_manual_file_check
from core.common import workbook_locked, workbook_rel_name
from core.excel_ops import open_workbook_checked, quit_background_excel, start_background_excel
from core.link_repair import fingerprint
from core.shape_geometry import _unique, apply_sheet, describe_change, plan_sheet, read_shapes

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


# ---------- 找出受影响的地方（不改文件） ----------

RENDER_WIDTH = 760  # 局部截图的像素宽


def pdf_for(path: Path, root: Path, pdf_root: Path | None) -> Path | None:
    """新版 PDF：填了 PDF 文件夹就按相对路径找同名 .pdf，否则看 Excel 旁边的同名 .pdf"""
    from core import pdf_bundle
    candidates = []
    if pdf_root:
        rel = path.relative_to(root)
        candidates.append(pdf_root / rel.with_name(pdf_bundle.pdf_path_for(path).name))
        candidates.append(pdf_root / pdf_bundle.pdf_path_for(path).name)
    candidates.append(pdf_bundle.pdf_path_for(path))
    return next((c for c in candidates if c.is_file()), None)


def _image_name(path: Path, sheet: str, key: str, side: str) -> str:
    digest = hashlib.sha1(f"{path}\0{sheet}\0{key}".encode()).hexdigest()[:16]
    return f"imgfix_{digest}_{side}.png"


def render_area(sheet: Any, names: list[str], out_name: str) -> str:
    """把这几个图形占的格子区域临时设成打印区域、缩成一页导出 PDF，转成 PNG 存起来；返回文件名，失败返回空。
    只用在只读打开、不保存的工作簿上"""
    from core.pdf_detector import fitz
    try:
        shapes = [sheet.Shapes.Item(n) for n in names]
        r1 = min(int(s.TopLeftCell.Row) for s in shapes)
        c1 = min(int(s.TopLeftCell.Column) for s in shapes)
        r2 = max(int(s.BottomRightCell.Row) for s in shapes)
        c2 = max(int(s.BottomRightCell.Column) for s in shapes)
        area = sheet.Range(sheet.Cells(max(1, r1 - 1), max(1, c1 - 1)), sheet.Cells(r2 + 1, c2 + 1))
        setup = sheet.PageSetup
        setup.PrintArea = area.Address
        setup.Zoom = False
        setup.FitToPagesWide = 1
        setup.FitToPagesTall = 1
        with tempfile.TemporaryDirectory(prefix="imgfix_") as tmp:
            pdf = Path(tmp) / "area.pdf"
            sheet.ExportAsFixedFormat(Type=0, Filename=str(pdf), Quality=0, IncludeDocProperties=False,
                                      IgnorePrintAreas=False, OpenAfterPublish=False)
            doc = fitz.open(str(pdf))
            try:
                page = doc[0]
                zoom = min(3.0, RENDER_WIDTH / max(1.0, float(page.rect.width)))
                fix_images.IMAGES_DIR.mkdir(parents=True, exist_ok=True)
                page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), annots=False).save(str(fix_images.IMAGES_DIR / out_name))
            finally:
                doc.close()
        return out_name
    except Exception as err:
        debug_log.event("⚠️ 局部截图没做出来", 表=str(getattr(sheet, "Name", "?")), 图形="、".join(names), 错误=str(err))
        return ""


def _anchor(sheet: Any, name: str) -> str:
    try:
        return str(sheet.Shapes.Item(name).TopLeftCell.Address(False, False))
    except Exception:
        return ""


def inspect_workbook(excel: Any, path: Path, old_path: Path, on_sheet=None) -> dict:
    """两份都只读打开，返回受影响的地方：{count, sheets:[{sheet, old_sheet, count, groups:[…]}]}"""
    if path.resolve() == old_path.resolve():
        raise ValueError("处理后的文件和参考文件是同一个，请分别选两个文件夹")
    old_book = open_workbook_checked(excel, str(old_path.resolve()), ReadOnly=True, UpdateLinks=0)
    try:
        old = read_book(old_book)
    finally:
        old_book.Close(SaveChanges=False)
    sheets: list[dict] = []
    book = open_workbook_checked(excel, str(path.resolve()), ReadOnly=True, UpdateLinks=0)
    try:
        pairs = pair_book_sheets(read_book(book), old)
        for index, (cur, ref) in enumerate(pairs, 1):
            if on_sheet:
                on_sheet(index, len(pairs), cur["name"])
            changes = plan_sheet(cur["shapes"], ref["shapes"])
            if not changes:
                continue
            sheet = book.Worksheets(cur["name"])
            ref_names = set(_unique(ref["shapes"]))
            groups: dict[str, dict] = {}
            for item in changes:
                key = item["host"] or item["name"]
                group = groups.setdefault(key, {"key": key, "names": [key] if key in ref_names else [], "details": [],
                                                "picture": bool(item["host"]) or item["kind"] == "picture"})
                if item["name"] not in group["names"]:
                    group["names"].append(item["name"])
                group["details"].append(describe_change(item))
            for group in groups.values():
                group["cell"] = _anchor(sheet, group["key"])
                group["new_img"] = render_area(sheet, group["names"], _image_name(path, cur["name"], group["key"], "new"))
            debug_log.event("🔍 找到受影响的图片", 文件=path.name, 表=cur["name"],
                            图形=[d for g in groups.values() for d in g["details"]])
            sheets.append({"sheet": cur["name"], "old_sheet": ref["name"], "count": len(changes), "groups": list(groups.values())})
    finally:
        book.Close(SaveChanges=False)
    if sheets:
        old_book = open_workbook_checked(excel, str(old_path.resolve()), ReadOnly=True, UpdateLinks=0)
        try:
            for entry in sheets:
                old_sheet = old_book.Worksheets(entry["old_sheet"])
                for group in entry["groups"]:
                    group["old_img"] = render_area(old_sheet, group["names"], _image_name(path, entry["sheet"], group["key"], "old"))
        finally:
            old_book.Close(SaveChanges=False)
    return {"count": sum(s["count"] for s in sheets), "sheets": sheets}


def _image_refs() -> list[str]:
    """换文件夹清理截图时，这里列着的局部截图要留着"""
    with _LOCK:
        files = ((_STATUS.get("result") or {}).get("files")) or []
        return [g.get(k) for f in files for s in f.get("sheets", []) for g in s.get("groups", [])
                for k in ("old_img", "new_img") if g.get(k)]


fix_images.register_references(_image_refs)


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


def describe(entry: dict, mode: str = "fix") -> str:
    if entry.get("error"):
        return entry["error"]
    if mode == "scan":
        return f"{entry['count']} 个图形变了（{len(entry['sheets'])} 张表）" if entry.get("count") else "图片都没变"
    return f"放回 {entry['count']} 个图形（{len(entry['sheets'])} 张表）" if entry.get("count") else "图片都没变"


# ---------- 后台任务 ----------

def _run(mode: str, root: Path, paths: list[Path], old_root: Path, old_paths: list[Path], pdf_root: Path | None) -> None:
    """mode=scan：只找不改；mode=fix：改这几份（结果并进上次找出来的清单，标成已修复）"""
    files: list[dict] = []
    excel, initialized = None, False
    guard = file_problem.Guard("找出受影响的图片" if mode == "scan" else "修复图片")
    try:
        wf.pythoncom.CoInitialize()
        initialized = True
        inferred: set[str] = set()
        pairs = pair_files(old_root, old_paths, [workbook_rel_name(p, root) for p in paths], root, inferred=inferred)
        for index, path in enumerate(paths):
            name = workbook_rel_name(path, root)
            while True:
                entry: dict[str, Any] = {"name": name, "path": str(path), "error": None, "count": 0, "sheets": []}
                pdf = pdf_for(path, root, pdf_root)
                entry["pdf"] = str(pdf) if pdf else ""
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

                    if mode == "scan":
                        entry.update(inspect_workbook(excel, path, old_path, on_sheet))
                    else:
                        entry.update(fix_workbook(excel, path, old_path, on_sheet))
                except Exception as error:
                    debug_log.exception(error, f"{'找受影响的图片' if mode == 'scan' else '修复图片'} {name}")
                    retry, entry["error"] = guard.problem(error, path, name, root,
                                                          close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {name}")
                        continue
                break
            files.append(entry)
            debug_log.event("🖼️ 图片一份", 动作=mode, 文件=name, 图形=entry.get("count", 0), 错误=entry["error"] or "",
                            PDF=entry["pdf"] or "-", 秒=round(time.perf_counter() - started, 2))
            wf.log_msg(f"🖼️ {name}：{describe(entry, mode)}")
            with _LOCK:
                progress_log.add(_STATUS, name, not entry["error"], describe(entry, mode), path=str(path))
    except Exception as error:
        debug_log.exception(error, "图片任务")
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
                    _STATUS.update(running=False, done=len(paths), current="", step="")
                    _STATUS["result"] = _merge(_STATUS.get("result"), mode, root, old_root, pdf_root, files)
                with wf.STATE_LOCK:
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)


def _merge(previous: dict | None, mode: str, root: Path, old_root: Path, pdf_root: Path | None, files: list[dict]) -> dict:
    """找的结果直接替换；修的结果按文件名并进上次找到的清单（fixed / fix_error / fixed_count）"""
    base = {"folder": str(root), "old_folder": str(old_root), "pdf_folder": str(pdf_root or ""), "files": files, "mode": mode}
    if mode == "scan" or not previous or previous.get("folder") != str(root) or not previous.get("scanned"):
        return {**base, "scanned": mode == "scan", "stamp": int(time.time())}  # stamp：局部截图地址带上，重新找过浏览器才换图
    by_name = {f["name"]: f for f in files}
    merged = []
    for old in previous.get("files", []):
        done = by_name.pop(old["name"], None)
        if done is not None:
            old = {**old, "fixed": not done["error"], "fix_error": done["error"] or "", "fixed_count": done.get("count", 0)}
        merged.append(old)
    for extra in by_name.values():
        merged.append({**extra, "fixed": not extra["error"], "fix_error": extra["error"] or "", "fixed_count": extra.get("count", 0)})
    return {**previous, "files": merged, "scanned": True, "mode": "fix"}


def start(folder: str, old_folder: str, mode: str = "scan", names: Any = None, pdf_folder: Any = "") -> None:
    """mode=scan 找出受影响的图片（不改文件）；mode=fix 修复。names：只修这几份（文件名列表），不给就是全部"""
    if mode not in {"scan", "fix"}:
        raise ValueError("不认识的操作")
    root, paths = list_workbooks(folder)
    old_root, old_paths = list_workbooks(old_folder)
    if root == old_root or root in old_root.parents or old_root in root.parents:
        raise ValueError("处理后的文件夹和参考文件夹不能相同或互相包含")
    pdf_root = None
    if isinstance(pdf_folder, str) and pdf_folder.strip():
        pdf_root = Path(pdf_folder.strip().strip('"').strip("'")).expanduser().resolve()
        if not pdf_root.is_dir():
            raise ValueError(f"PDF 文件夹不存在：{pdf_root}")
    if names is not None:
        if not isinstance(names, list) or not all(isinstance(n, str) for n in names) or not names:
            raise ValueError("请选择要修复的文件")
        wanted = set(names)
        paths = [p for p in paths if workbook_rel_name(p, root) in wanted]
        if not paths:
            raise ValueError("要修复的文件不在这个文件夹里，请重新找一遍")
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
        result = _STATUS.get("result")
        _STATUS.clear()
        _STATUS.update(running=True, mode=mode, folder=str(root), old_folder=str(old_root), pdf_folder=str(pdf_root or ""),
                       total=len(paths), done=0, current="", result=result)

    def run() -> None:
        with debug_log.side_task("image_restore", {"mode": mode, "folder": str(root), "old_folder": str(old_root), "files": len(paths)}):
            _run(mode, root, paths, old_root, old_paths, pdf_root)

    try:
        threading.Thread(target=run, name="修复图片", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
        raise


def open_pdf(name: str) -> dict:
    """用系统默认程序打开这份 Excel 对应的新版 PDF（只认上次结果里列出来的）"""
    with _LOCK:
        files = ((_STATUS.get("result") or {}).get("files")) or []
        file = next((f for f in files if f.get("name") == name), None)
    if not file or not file.get("pdf"):
        raise ValueError("这份没有找到对应的 PDF")
    pdf = Path(file["pdf"])
    if not pdf.is_file():
        raise ValueError(f"PDF 不在了：{pdf}")
    if not hasattr(os, "startfile"):
        raise ValueError("当前系统不能直接打开 PDF")
    os.startfile(str(pdf))  # noqa: S606 — 只开结果里列出的本地 PDF
    debug_log.event("📄 打开新版 PDF", 文件=name, PDF=str(pdf))
    return {"pdf": str(pdf)}


def compare_target(name: str, sheet: str = "", cell: str = "") -> dict:
    """「对照打开」：左边新版（跳到图形左上角那格），右边原版只读同一格"""
    with _LOCK:
        if _STATUS.get("running"):
            raise ValueError("正在处理，请等任务结束后再打开对比")
        files = ((_STATUS.get("result") or {}).get("files")) or []
        file = next((f for f in files if f.get("name") == name), None)
    if not file or not file.get("old"):
        raise ValueError("这份文件没有对应的原文件，暂时不能同时打开")
    entry = next((s for s in file.get("sheets", []) if s["sheet"] == sheet), None) if sheet else (file.get("sheets") or [None])[0]
    current, backup = Path(file["path"]).resolve(strict=True), Path(file["old"]).resolve(strict=True)
    return {"current": current, "backup": backup, "sheet": entry["sheet"] if entry else "",
            "old_sheet": entry["old_sheet"] if entry else "", "cell": cell if entry else ""}


def status() -> dict:
    with _LOCK:
        return copy.deepcopy(_STATUS)
