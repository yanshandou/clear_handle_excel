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

页面分两步（用户 2026-11-11 要）：先「找出受影响的图片」——两份都只读打开，不改任何文件，
有图片的表三列整页图，每列把这张表的所有页从上到下列全，不按页配对（原版和新版分页位置可能不一样，用户 2026-11-11 定）：
原版 Excel｜新版 Excel 是屏幕原样（按本工具同一套做法导出 PDF 得到页数，再按分页取每页的格子区域「复制为图片」，借剪贴板；
拍不了退回导出效果并标出来）｜新版 PDF（整份 PDF 里按文字找出属于这张表的那一段页），能比对也能定位；用户可以把看过没问题的页标「无需处理」（存在浏览器里）；
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

from core import debug_log, file_problem, fix_images, print_overlays, progress_log, workflow as wf
from core.backup_compare import pair_files, require_manual_file_check
from core.common import workbook_locked, workbook_rel_name
from core.excel_ops import open_workbook_checked, quit_background_excel, start_background_excel
from core.link_repair import fingerprint
from core.shape_geometry import PICTURES, _unique, apply_sheet, diagnose, overlay_hosts, describe_change, plan_sheet, read_shapes

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


PAGE_WIDTH = 1100  # 整页图的像素宽（三列各占约 1/3 屏，点开看大图也够清楚）
PAGE_MATCH_MIN = 0.2  # 两页文字行重合度低于这个，算没找到对应页


def _anchor(sheet: Any, name: str) -> str:
    try:
        return str(sheet.Shapes.Item(name).TopLeftCell.Address(False, False))
    except Exception:
        return ""


def _norm(text: str) -> str:
    import re
    return re.sub(r"\s+", "", text or "")


def export_sheet(sheet: Any, target: Path) -> Path:
    """这张表按本工具导出 PDF 的同一套做法（宽一页、空心大框不印、截图和标注临时组合）导出；
    只用在只读打开、不保存的工作簿上"""
    wf._export_sheet_pdf_raw(sheet, target)
    return target


def _page_infos(doc: Any) -> list[dict]:
    """每页：文字行（去空白，至少 2 个字）、有没有图片"""
    infos = []
    for page in doc:
        lines = {_norm(line) for line in page.get_text().splitlines()}
        try:
            pictures = sum(1 for info in page.get_image_info() if info["bbox"][2] - info["bbox"][0] > 30
                           and info["bbox"][3] - info["bbox"][1] > 30)
        except Exception:
            pictures = 0
        infos.append({"lines": {line for line in lines if len(line) >= 2}, "pictures": pictures})
    return infos


def _best_page(lines: set[str], infos: list[dict]) -> tuple[int, float]:
    """文字行重合度（交集 / 并集）最高的那页；(页号从 0 起, 重合度)，一页都没有返回 (-1, 0)"""
    best, score = -1, 0.0
    for number, info in enumerate(infos):
        union = len(lines | info["lines"])
        value = len(lines & info["lines"]) / union if union else 0.0
        if value > score:
            best, score = number, value
    return best, score


def _save_page(doc: Any, number: int, name: str) -> str:
    from core.pdf_detector import fitz
    page = doc[number]
    zoom = min(3.0, PAGE_WIDTH / max(1.0, float(page.rect.width)))
    fix_images.IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), annots=False, alpha=False).save(str(fix_images.IMAGES_DIR / name))
    return name


def sheet_pages(pdf: Path, path: Path, sheet: str, side: str) -> list[dict]:
    """这张表导出的 PDF 的每一页（先放导出效果的整页图；能拍屏幕原样的后面换掉）。
    用户 2026-11-11 定：不按页配对、不挑有图片的页，一张表的所有页都列出来。
    [{"page", "img", "how": "pdf", "ratio", "_lines"}]"""
    from core.pdf_detector import fitz
    doc = fitz.open(str(pdf))
    try:
        pages = []
        for number, info in enumerate(_page_infos(doc)):
            rect = doc[number].rect
            pages.append({"page": number + 1, "how": "pdf", "_lines": info["lines"],
                          "ratio": round(float(rect.width) / max(1.0, float(rect.height)), 3),  # 图片先按比例占位，加载时页面不跳
                          "img": _save_page(doc, number, _image_name(path, sheet, f"p{number}", side))})
        return pages
    finally:
        doc.close()


def pdf_pages(pdf: Path | None, path: Path, sheet: str, new_pages: list[dict]) -> tuple[list[dict], str]:
    """用户已有的新版 PDF（整份工作簿）里属于这张表的那几页：新版这张表每页按文字行找最像的页，取最前到最后这一段。
    返回 (页列表, 没有时的原因)"""
    from core.pdf_detector import fitz
    if pdf is None:
        return [], "没找到同名 PDF（Excel 旁边和 PDF 文件夹里都没有）"
    try:
        doc = fitz.open(str(pdf))
    except Exception as err:
        debug_log.event("⚠️ 新版 PDF 读不了", PDF=str(pdf), 错误=str(err))
        return [], f"新版 PDF 读不了：{err}"
    try:
        infos = _page_infos(doc)
        found = []
        for page in new_pages:
            number, score = _best_page(page.get("_lines") or set(), infos)
            if number >= 0 and score >= PAGE_MATCH_MIN:
                found.append(number)
        if not found:
            return [], "新版 PDF 里没找到这张表的页（文字对不上）"
        pages = []
        for number in range(min(found), max(found) + 1):
            rect = doc[number].rect
            pages.append({"page": number + 1, "how": "pdf",
                          "ratio": round(float(rect.width) / max(1.0, float(rect.height)), 3),
                          "img": _save_page(doc, number, _image_name(path, sheet, f"p{number}", "pdf"))})
        return pages, ""
    finally:
        doc.close()


# ---------- Excel 屏幕原样的整页图（用户 2026-11-11：「我需要 excel 屏幕原样」） ----------

XL_SCREEN, XL_BITMAP = 1, 2
XL_PAGE_BREAK_PREVIEW = 2
XL_NORMAL_VIEW = 1
SCREEN_PIXELS = 1300  # 屏幕原样整页图大约多少像素宽


def page_ranges(sheet: Any, expected: int) -> list[Any]:
    """这张表每一页对应的格子区域（按导出 PDF 时的分页，宽一页所以一页就是一段行）。
    分页数和导出的 PDF 页数（expected）对不上就返回 []——宁可退回导出效果，也不把别的页当成这一页。
    只用在只读打开、不保存的工作簿上（会切到分页预览，Excel 才肯报全分页）"""
    try:
        sheet.Activate()
        sheet.Application.ActiveWindow.View = XL_PAGE_BREAK_PREVIEW
    except Exception:
        pass
    address = str(sheet.PageSetup.PrintArea or "")
    if "," in address or ";" in address:
        raise RuntimeError("打印区域是好几块")
    area = sheet.Range(address) if address else sheet.UsedRange
    first_row, first_col = int(area.Row), int(area.Column)
    last_row = first_row + int(area.Rows.Count) - 1
    last_col = first_col + int(area.Columns.Count) - 1
    if not address:  # 没设打印区域：Excel 从 A1 印到最后一个有内容的格子，图形伸出去的部分也算
        first_row = first_col = 1
        for index in range(1, int(sheet.Shapes.Count) + 1):
            try:
                cell = sheet.Shapes.Item(index).BottomRightCell
                last_row, last_col = max(last_row, int(cell.Row)), max(last_col, int(cell.Column))
            except Exception:
                continue
    breaks = sheet.HPageBreaks
    cuts = sorted({int(breaks.Item(i).Location.Row) for i in range(1, int(breaks.Count) + 1)})
    cuts = [row for row in cuts if first_row < row <= last_row]
    try:  # 分页读完切回普通视图、不画分页线：拍出来不带「第 N 页」水印和蓝色分页线
        sheet.Application.ActiveWindow.View = XL_NORMAL_VIEW
        sheet.DisplayPageBreaks = False
    except Exception:
        pass
    starts = [first_row, *cuts]
    ends = [*(row - 1 for row in cuts), last_row]
    if len(starts) != expected:
        raise RuntimeError(f"Excel 报的分页是 {len(starts)} 页，导出的 PDF 是 {expected} 页，对不上")
    return [sheet.Range(sheet.Cells(top, first_col), sheet.Cells(bottom, last_col)) for top, bottom in zip(starts, ends)]


def render_screen(sheet: Any, area: Any, out_name: str) -> float:
    """这块格子在 Excel 屏幕上的样子：「复制为图片（如屏幕所示）」贴进临时图表存成 PNG；返回宽高比。
    要借用剪贴板（Ditto 之类会记下这些截图）；剪贴板被别的程序占着时重试几次。临时图表用完删掉"""
    width, height = float(area.Width), float(area.Height)
    if width < 2 or height < 2:
        raise RuntimeError("这一页的格子区域是空的")
    scale = max(0.5, min(2.0, SCREEN_PIXELS / (width * 96 / 72)))
    holder = None
    try:
        last: Exception | None = None
        for attempt in range(6):
            try:
                area.CopyPicture(Appearance=XL_SCREEN, Format=XL_BITMAP)
                break
            except Exception as err:  # 剪贴板正被 Ditto 等占用
                last = err
                time.sleep(0.3 * (attempt + 1))
        else:
            raise RuntimeError(f"复制成图片失败（剪贴板被占用？）：{last}")
        holder = sheet.ChartObjects().Add(0, 0, width * scale, height * scale)
        chart = holder.Chart
        try:
            chart.ChartArea.Format.Line.Visible = 0
            chart.ChartArea.Format.Fill.Visible = 0
        except Exception:
            pass
        try:
            holder.Activate()  # 不激活有时贴进去是空白
        except Exception:
            pass
        chart.Paste()
        pasted = chart.Shapes.Item(chart.Shapes.Count)
        pasted.LockAspectRatio = 0
        pasted.Left, pasted.Top = 0, 0
        pasted.Width, pasted.Height = width * scale, height * scale
        fix_images.IMAGES_DIR.mkdir(parents=True, exist_ok=True)
        target = fix_images.IMAGES_DIR / out_name
        if not chart.Export(str(target), "PNG") or not target.is_file() or target.stat().st_size < 200:
            raise RuntimeError("图表导出的图片是空的")
        return round(width / height, 3)
    finally:
        if holder is not None:
            try:
                holder.Delete()
            except Exception:
                pass


def screen_pages(sheet: Any, pages: list[dict], side: str, path: Path, sheet_name: str) -> None:
    """把这张表的每一页换成 Excel 屏幕原样；哪一页拍不了就留着导出效果的图（页面上会标出来）"""
    if not pages:
        return
    where = {"文件": path.name, "表": sheet_name, "哪边": "新版" if side == "new" else "原版"}
    try:
        ranges = page_ranges(sheet, len(pages))
    except Exception as err:
        debug_log.event("⚠️ 屏幕原样拍不了，这张表用导出效果", **where, 错误=str(err))
        return
    for page, area in zip(pages, ranges):
        name = _image_name(path, sheet_name, f"p{page['page'] - 1}", f"{side}s")
        for attempt in range(3):  # 偶尔贴进图表是空的（实测报「インデックスが境界を超えています」），再拍一次多半就好
            try:
                page["ratio"] = render_screen(sheet, area, name)
                page["img"], page["how"] = name, "screen"
                break
            except Exception as err:
                debug_log.event("⚠️ 屏幕原样拍不了" + ("，再试一次" if attempt < 2 else "，这一页用导出效果"), **where,
                                页=page["page"], 错误=str(err))
                time.sleep(0.4)


def inspect_workbook(excel: Any, path: Path, old_path: Path, on_sheet=None, pdf: Path | None = None) -> dict:
    """两份都只读打开，返回受影响的地方：{count, sheets:[{sheet, old_sheet, count, groups:[…], pages:[…]}]}"""
    if path.resolve() == old_path.resolve():
        raise ValueError("处理后的文件和参考文件是同一个，请分别选两个文件夹")
    old_book = open_workbook_checked(excel, str(old_path.resolve()), ReadOnly=True, UpdateLinks=0)
    try:
        old = read_book(old_book)
    finally:
        old_book.Close(SaveChanges=False)
    sheets: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="imgfix_") as tmp:
        tmp_dir = Path(tmp)
        book = open_workbook_checked(excel, str(path.resolve()), ReadOnly=True, UpdateLinks=0)
        try:
            pairs = pair_book_sheets(read_book(book), old)
            for index, (cur, ref) in enumerate(pairs, 1):
                if on_sheet:
                    on_sheet(index, len(pairs), cur["name"])
                changes = plan_sheet(cur["shapes"], ref["shapes"])
                hosts = overlay_hosts(cur["shapes"], ref["shapes"], every=True)  # 所有图片都列（用户 2026-11-11 要）
                if cur["shapes"] or ref["shapes"]:
                    debug_log.event("🔎 对照一张表", 文件=path.name, 表=cur["name"], 变了=len(changes),
                                    **diagnose(cur["shapes"], ref["shapes"]))
                if not changes and not hosts:
                    continue
                sheet = book.Worksheets(cur["name"])
                ref_names = set(_unique(ref["shapes"]))
                groups: dict[str, dict] = {}
                for item in changes:
                    key = item["host"] or item["name"]
                    group = groups.setdefault(key, {"key": key, "names": [key] if key in ref_names else [], "details": [],
                                                    "picture": bool(item["host"]) or item["kind"] == "picture", "changed": True})
                    if item["name"] not in group["names"]:
                        group["names"].append(item["name"])
                    group["details"].append(describe_change(item))
                for host, overlays in hosts.items():
                    if host not in groups:
                        groups[host] = {"key": host, "names": [host, *overlays], "picture": True, "changed": False, "details": []}
                for group in groups.values():
                    group["cell"] = _anchor(sheet, group["key"])
                    group["sheet"] = cur["name"]
                ordered = sorted(groups.values(), key=lambda g: not g["changed"])
                entry = {"sheet": cur["name"], "old_sheet": ref["name"], "count": len(changes), "groups": ordered,
                         "review": sum(not g["changed"] for g in ordered)}
                new_pdf = tmp_dir / f"new_{index}.pdf"
                if on_sheet:
                    on_sheet(index, len(pairs), f"{cur['name']}（导出整页）")
                try:
                    export_sheet(sheet, new_pdf)
                    entry["_new"] = sheet_pages(new_pdf, path, cur["name"], "new")
                    screen_pages(sheet, entry["_new"], "new", path, cur["name"])
                except Exception as err:
                    debug_log.exception(err, f"整页图 {path.name} {cur['name']}")
                    entry["pages_why"] = f"新版这张表的整页图没做出来：{err}"
                debug_log.event("🔍 找到图片", 文件=path.name, 表=cur["name"], 图形=[d for g in ordered for d in g["details"]])
                sheets.append(entry)
        finally:
            book.Close(SaveChanges=False)
        if any(entry.get("_new") for entry in sheets):
            old_book = open_workbook_checked(excel, str(old_path.resolve()), ReadOnly=True, UpdateLinks=0)
            try:
                for index, entry in enumerate(sheets, 1):
                    if not entry.get("_new"):
                        continue
                    try:
                        old_sheet = old_book.Worksheets(entry["old_sheet"])
                        old_pdf = export_sheet(old_sheet, tmp_dir / f"old_{index}.pdf")
                        entry["_old"] = sheet_pages(old_pdf, path, entry["sheet"], "old")
                    except Exception as err:
                        debug_log.event("⚠️ 原版这张表的整页图没做出来", 文件=path.name, 表=entry["old_sheet"], 错误=str(err))
                        entry["_old_why"] = f"原版这张表的整页图没做出来：{err}"
                        continue
                    screen_pages(old_sheet, entry["_old"], "old", path, entry["sheet"])
            finally:
                old_book.Close(SaveChanges=False)
        for entry in sheets:
            new_pages, old_pages = entry.pop("_new", None), entry.pop("_old", [])
            old_why = entry.pop("_old_why", "")
            if not new_pages:
                continue
            book_pages, pdf_why = pdf_pages(pdf, path, entry["sheet"], new_pages)
            for page in [*new_pages, *old_pages]:
                page.pop("_lines", None)
            # 三列各自从第一页列到最后一页，不按页配对（原版和新版分页位置可能不一样）
            entry["cols"] = {"old": old_pages, "new": new_pages, "pdf": book_pages}
            entry["old_why"], entry["pdf_why"] = old_why, pdf_why
            debug_log.event("📄 整页图", 文件=path.name, 表=entry["sheet"],
                            原版=[f"{p['page']}:{p['how']}" for p in old_pages] or old_why,
                            新版=[f"{p['page']}:{p['how']}" for p in new_pages],
                            PDF=[p["page"] for p in book_pages] or pdf_why)
    return {"count": sum(s["count"] for s in sheets), "sheets": sheets, "review": sum(s["review"] for s in sheets)}


def _image_refs() -> list[str]:
    """换文件夹清理截图时，这里列着的局部截图要留着"""
    with _LOCK:
        files = ((_STATUS.get("result") or {}).get("files")) or []
        return [p["img"] for f in files for s in f.get("sheets", []) for pages in (s.get("cols") or {}).values()
                for p in pages if p.get("img")]


fix_images.register_references(_image_refs)


# ---------- 一份 Excel ----------

def fix_workbook(excel: Any, path: Path, old_path: Path, on_sheet=None, group: bool = True) -> dict:
    """放回原样；group=True 时放回后顺便把每张表的截图和标注永久组合（用户 2026-11-11 要）"""
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
        grouped = _group_book(book) if group else []
        if sheets or grouped:
            if fingerprint(path) != stamp:
                raise ValueError("修复期间文件被别人改过，没有保存")
            book.Save()
            debug_log.saved_excel("修复图片", path)
    finally:
        book.Close(SaveChanges=False)
    return {"count": sum(s["count"] for s in sheets), "sheets": sheets,
            "grouped": sum(len(g["groups"]) for g in grouped), "group_sheets": grouped}


def _group_book(book: Any) -> list[dict]:
    """每张表的截图和标注永久组合，返回 [{sheet, groups:[说明]}]"""
    out = []
    for sheet in book.Worksheets:
        done = print_overlays.group_permanently(sheet)
        if done:
            out.append({"sheet": str(sheet.Name), "groups": done})
    return out


def group_workbook(excel: Any, path: Path, on_sheet=None) -> dict:
    """只组合、不对照原版（「📌 只组合」按钮）。已经被拉坏的图会按坏的样子组合，所以请先修复"""
    if workbook_locked(path):
        raise ValueError("文件正在 Excel 中打开，请保存并关闭后再组合")
    stamp = fingerprint(path)
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开，没有修改")
        grouped = _group_book(book)
        if grouped:
            if fingerprint(path) != stamp:
                raise ValueError("组合期间文件被别人改过，没有保存")
            book.Save()
            debug_log.saved_excel("截图和标注永久组合", path)
    finally:
        book.Close(SaveChanges=False)
    return {"count": 0, "sheets": [], "grouped": sum(len(g["groups"]) for g in grouped), "group_sheets": grouped}


def describe(entry: dict, mode: str = "fix") -> str:
    if entry.get("error"):
        return entry["error"]
    if mode == "scan":
        pages = sum(len((s.get("cols") or {}).get("new", [])) for s in entry.get("sheets", []))
        page_text = f"，有图片的表共 {pages} 页请看三列" if pages else ""
        return (f"{entry['count']} 个图形变了" if entry.get("count") else "程序没量出变化") + page_text
    parts = [f"放回 {entry['count']} 个图形" if entry.get("count") else "",
             f"组合 {entry['grouped']} 组截图和标注" if entry.get("grouped") else ""]
    parts = [p for p in parts if p]
    return "，".join(parts) if parts else ("没有要组合的截图" if mode == "group" else "图片都没变")


# ---------- 后台任务 ----------

def _run(mode: str, root: Path, paths: list[Path], old_root: Path | None, old_paths: list[Path], pdf_root: Path | None,
         group: bool = True) -> None:
    """mode=scan：只找不改；mode=fix：改这几份（结果并进上次找出来的清单，标成已修复）；
    mode=group：只把截图和标注永久组合，不对照原版（old_root 可以是 None）"""
    files: list[dict] = []
    stamp = int(time.time())  # 局部截图地址带上它：查的过程中不变，页面重画不重新下载图
    excel, initialized = None, False
    guard = file_problem.Guard({"scan": "找出受影响的图片", "fix": "修复图片", "group": "组合截图和标注"}[mode])
    try:
        wf.pythoncom.CoInitialize()
        initialized = True
        inferred: set[str] = set()
        pairs = (pair_files(old_root, old_paths, [workbook_rel_name(p, root) for p in paths], root, inferred=inferred)
                 if old_root else {})
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
                    if mode != "group":
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
                        entry.update(inspect_workbook(excel, path, old_path, on_sheet, pdf))
                    elif mode == "group":
                        entry.update(group_workbook(excel, path, on_sheet))
                    else:
                        entry.update(fix_workbook(excel, path, old_path, on_sheet, group))
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
            debug_log.event("🖼️ 图片一份", 动作=mode, 文件=name, 图形=entry.get("count", 0), 组合=entry.get("grouped", 0),
                            错误=entry["error"] or "",
                            PDF=entry["pdf"] or "-", 秒=round(time.perf_counter() - started, 2))
            wf.log_msg(f"🖼️ {name}：{describe(entry, mode)}")
            with _LOCK:
                progress_log.add(_STATUS, name, not entry["error"], describe(entry, mode), path=str(path))
                if mode == "scan":  # 边查边显示：查完一份就能在页面上看三列（用户 2026-11-11 要）
                    _STATUS["result"] = {"folder": str(root), "old_folder": str(old_root or ""), "pdf_folder": str(pdf_root or ""),
                                         "files": list(files), "mode": "scan", "scanned": True, "partial": True, "stamp": stamp}
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
                    _STATUS["result"] = _merge(_STATUS.get("result"), mode, root, old_root, pdf_root, files, stamp)
                with wf.STATE_LOCK:
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)


def _merge(previous: dict | None, mode: str, root: Path, old_root: Path | None, pdf_root: Path | None, files: list[dict],
           stamp: int | None = None) -> dict:
    """找的结果直接替换；修 / 组合的结果按文件名并进上次找到的清单（fixed / fix_error / fixed_count / grouped）"""
    base = {"folder": str(root), "old_folder": str(old_root or ""), "pdf_folder": str(pdf_root or ""), "files": files, "mode": mode}
    if mode == "scan" or not previous or previous.get("folder") != str(root) or not previous.get("scanned"):
        # stamp：局部截图地址带上，重新找过浏览器才换图；沿用查的过程中那个，查完不重新下载图（页面不跳）
        return {**base, "scanned": mode == "scan", "stamp": stamp or int(time.time())}
    def outcome(done: dict) -> dict:
        info = {"grouped": done.get("grouped", 0), "group_sheets": done.get("group_sheets", [])}
        if mode == "fix":
            info.update(fixed=not done["error"], fix_error=done["error"] or "", fixed_count=done.get("count", 0))
        else:
            info.update(group_error=done["error"] or "")
        return info

    by_name = {f["name"]: f for f in files}
    merged = []
    for old in previous.get("files", []):
        done = by_name.pop(old["name"], None)
        merged.append({**old, **outcome(done)} if done is not None else old)
    for extra in by_name.values():
        merged.append({**extra, **outcome(extra)})
    return {**previous, "files": merged, "scanned": True, "mode": mode}


def start(folder: str, old_folder: str, mode: str = "scan", names: Any = None, pdf_folder: Any = "",
          group: Any = True) -> None:
    """mode=scan 找出受影响的图片（不改文件）；mode=fix 修复（group=True 顺便永久组合）；mode=group 只组合（不用原版）。
    names：只做这几份（文件名列表），不给就是全部"""
    if mode not in {"scan", "fix", "group"}:
        raise ValueError("不认识的操作")
    if not isinstance(group, bool):
        raise ValueError("组合选项不对")
    root, paths = list_workbooks(folder)
    old_root, old_paths = None, []
    if mode != "group" or (old_folder or "").strip():
        if not (old_folder or "").strip():
            raise ValueError("请输入原版文件夹")
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
            raise ValueError("请选择要处理的文件")
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
        result = None if mode == "scan" else _STATUS.get("result")  # 重新找：清掉上次的清单，查完一份列一份
        _STATUS.clear()
        _STATUS.update(running=True, mode=mode, folder=str(root), old_folder=str(old_root or ""), pdf_folder=str(pdf_root or ""),
                       total=len(paths), done=0, current="", result=result)

    def run() -> None:
        with debug_log.side_task("image_restore", {"mode": mode, "folder": str(root), "old_folder": str(old_root or ""),
                                                   "files": len(paths), "group": group}):
            _run(mode, root, paths, old_root, old_paths, pdf_root, group)

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


def open_block_reason(name: str) -> str:
    """任务进行中这份能不能打开（空 = 能）。查找只读、查完的不再碰，所以边查边看时查完的可以打开（用户 2026-11-11 要）；
    修复 / 组合要写文件，进行中一律不让开"""
    with _LOCK:
        if not _STATUS.get("running"):
            return ""
        if _STATUS.get("mode") != "scan":
            return "正在修复图片，请等当前任务结束"
        files = ((_STATUS.get("result") or {}).get("files")) or []
        done = any(f.get("name") == name for f in files)
    return "" if done else "这份还没查完，等它出现在结果里再打开"


def compare_target(name: str, sheet: str = "", cell: str = "") -> dict:
    """「对照打开」：左边新版（跳到图形左上角那格），右边原版只读同一格"""
    reason = open_block_reason(name)
    if reason:
        raise ValueError(reason)
    with _LOCK:
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
