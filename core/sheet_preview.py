"""精简 Sheet 的预览确认窗口：用 Excel 只读打开工作簿，按导出 PDF 功能相同的参数把每个工作表导出成 PDF，再用 pymupdf 渲染第 1 页缩略图。

PDF 与缩略图缓存在 _sheet_preview_cache/<键>/ 下，键由文件路径、修改时间与大小决定；文件没改过时再次打开直接读缓存。
"""

import hashlib
import json
import re
import shutil
import threading
import time
from collections import Counter
from pathlib import Path
from typing import Any

from core import debug_log
from core.common import PROJECT_ROOT, XL_SHEET_VISIBLE, normalize_used_values, workbook_rel_name
from core.excel_ops import pythoncom, quit_background_excel, start_background_excel
from core.pdf_detector import fitz
from core.workflow import _export_sheet_pdf, _preview_workbook, list_workbooks

PREVIEW_CACHE_DIR = PROJECT_ROOT / "_sheet_preview_cache"
PREVIEW_TARGET_WIDTH = 1000  # 缩略图宽度（px）：页面上按卡片宽度缩小显示，点开大图时仍然清晰
PREVIEW_MAX_ZOOM = 3.0
PREVIEW_CLIP_MARGIN = 12  # 裁掉页面空白时四周保留的边距（pt）
CACHE_MAX_AGE_SECONDS = 7 * 24 * 3600
MANIFEST_NAME = "manifest.json"
CACHE_KEY_RE = re.compile(r"^[0-9a-f]{16}$")
IMAGE_NAME_RE = re.compile(r"^\d{1,4}\.(?:png|pdf)$")
SIMILARITY_MIN_RATIO = 0.5  # 低于这个相似度不提示；新旧版本之间通常只改了一部分内容
SIMILARITY_MAX_CELLS = 500_000  # 已用区域超过这么多单元格时跳过比对，避免读取过慢
SIMILARITY_MAX_SHAPES = 2000  # 图形太多时不逐个读取，只比较数量
PREVIEW_CACHE_VERSION = 2  # 比对规则变了就加 1，旧缓存里的比对结果作废（2：图片、图形、批注也要相同）

_render_locks: dict[str, threading.Lock] = {}
_render_locks_guard = threading.Lock()


def _find_workbook(folder_str: str, file_name: str) -> tuple[Path, Path]:
    """只接受扫描结果里出现过的文件（所选文件夹内的 Excel），不能借此读取任意路径"""
    folder, workbooks = list_workbooks(folder_str)
    for wb_path in workbooks:
        if workbook_rel_name(wb_path, folder) == file_name:
            return folder, wb_path
    raise ValueError(f"找不到文件「{file_name}」，可能已被移动或改名，请重新扫描")


def workbook_preview_info(folder_str: str, file_name: str) -> dict[str, Any]:
    """单个文件的工作表列表与默认标记（与精简 Sheet 扫描结果一致），预览窗口先用它画出占位卡片"""
    folder, wb_path = _find_workbook(folder_str, file_name)
    holder: dict[str, Any] = {}

    def get_excel() -> Any:
        if "excel" not in holder:
            pythoncom.CoInitialize()
            holder["excel"] = start_background_excel()
        return holder["excel"]

    try:
        entry = _preview_workbook(wb_path, folder, get_excel)
    finally:
        if "excel" in holder:
            try:
                quit_background_excel(holder["excel"])
            finally:
                pythoncom.CoUninitialize()
    return {"folder": str(folder), "file": entry}


def _cache_key(wb_path: Path) -> str:
    stat = wb_path.stat()
    raw = f"{wb_path.resolve()}|{stat.st_mtime_ns}|{stat.st_size}|v{PREVIEW_CACHE_VERSION}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def _read_manifest(cache_dir: Path) -> dict[str, Any] | None:
    try:
        return json.loads((cache_dir / MANIFEST_NAME).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def _prune_old_cache(keep: str) -> None:
    if not PREVIEW_CACHE_DIR.is_dir():
        return
    cutoff = time.time() - CACHE_MAX_AGE_SECONDS
    for child in PREVIEW_CACHE_DIR.iterdir():
        try:
            if child.name != keep and child.is_dir() and child.stat().st_mtime < cutoff:
                shutil.rmtree(child, ignore_errors=True)
        except OSError:
            pass


def _normalize_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_sheet_content(sheet: Any) -> dict[str, Any] | None:
    """读取工作表已用区域的内容，返回 {"cells": {"行,列": 值}, "rows": [每行内容]}；图表页或区域过大时返回 None"""
    try:
        used = sheet.UsedRange
        if used.Rows.Count * used.Columns.Count > SIMILARITY_MAX_CELLS:
            return None
        row0, col0 = used.Row, used.Column
        values = normalize_used_values(used.Value)
    except Exception:
        return None
    cells: dict[str, str] = {}
    rows: list[str] = []
    for r, row in enumerate(values):
        texts = [_normalize_cell(v) for v in row]
        for c, text in enumerate(texts):
            if text:
                cells[f"{row0 + r},{col0 + c}"] = text
        while texts and not texts[-1]:
            texts.pop()
        if texts:
            rows.append("\t".join(texts))
    return {"cells": cells, "rows": rows}


def _shape_text(shape: Any) -> str:
    try:
        if shape.TextFrame2.HasText:
            return " ".join(str(shape.TextFrame2.TextRange.Text).split())
    except Exception:
        pass
    return ""


def read_sheet_objects(sheet: Any) -> list[str] | None:
    """工作表上的图片、图形（红框、箭头、文本框……）和批注，各自变成一行「类型|左|上|宽|高|文字」，排好序返回。
    复制工作表后只换了截图、挪了红框时，单元格完全一样，只有这里不同。读不到时返回 None（不参与比较）"""
    try:
        shapes = sheet.Shapes
        count = shapes.Count
    except Exception:
        return None
    items: list[str] = []
    if count > SIMILARITY_MAX_SHAPES:
        items.append(f"图形数量|{count}")
    else:
        try:
            for shape in shapes:
                box = "|".join(str(round(float(getattr(shape, attr)))) for attr in ("Left", "Top", "Width", "Height"))
                items.append(f"图形|{shape.Type}|{box}|{_shape_text(shape)}")
        except Exception:
            return None
    try:
        for comment in sheet.Comments:
            items.append(f"批注|{comment.Parent.Address}|{' '.join(str(comment.Text()).split())}")
    except Exception:
        pass
    return sorted(items)


def pdf_image_digests(pdf_path: Path) -> list[str] | None:
    """导出的 PDF 里每张图片的像素指纹（全部页）。Excel 读不到图片本身，图形位置一样、只换了截图时靠它分辨"""
    try:
        digests = []
        with fitz.open(pdf_path) as doc:
            for page in doc:
                for img in page.get_images(full=True):
                    pix = fitz.Pixmap(doc, img[0])
                    digests.append(f"{pix.width}x{pix.height}:{hashlib.md5(pix.samples).hexdigest()}")
        return sorted(digests)
    except Exception:
        return None


def _same_extra(a: dict[str, Any], b: dict[str, Any], key: str) -> bool:
    """图形/图片：两边都读到了才比较，读不到的一边不算不同（不让偶尔读取失败把真正的副本判成不同）"""
    if a.get(key) is None or b.get(key) is None:
        return True
    return a[key] == b[key]


def find_similar_sheets(contents: dict[str, dict[str, Any] | None]) -> dict[str, dict[str, Any]]:
    """两两比较工作表内容，给每个表找出最相似的另一个表

    完全相同：每个单元格的位置和值都一样，图片、图形、批注也一样（objects / images 键，可省略）。
    单元格一样、只有图片或图形不同时（复制工作表后换了截图、挪了红框），相似度记 99%，
    并带 same_cells=True、diff=["图形"/"图片"]，页面上提示「文字相同，图片/图形不同」。
    相似度取两种算法的较大值：
    - 按单元格：同一位置值相同的单元格数 ÷ 两表有内容的单元格位置总数（改了几个单元格时准确）
    - 按行：2×内容相同的行数 ÷ 两表行数之和（插入或删除几行、后面整体错位时准确）
    空白表、图表页不参与比较；相似度低于 SIMILARITY_MIN_RATIO 不提示。
    """
    names = [name for name, content in contents.items() if content and content["rows"]]
    counters = {name: Counter(contents[name]["rows"]) for name in names}
    result: dict[str, dict[str, Any]] = {}
    for name in names:
        best: dict[str, Any] | None = None
        for other in names:
            if other == name:
                continue
            a, b = contents[name], contents[other]
            same_cells = a["cells"] == b["cells"]
            diff = [label for key, label in (("objects", "图形"), ("images", "图片")) if not _same_extra(a, b, key)]
            identical = same_cells and not diff
            if identical:
                ratio = 1.0
            elif same_cells:
                ratio = 0.99
            else:
                cells_a, cells_b = contents[name]["cells"], contents[other]["cells"]
                same_count = sum(1 for pos, value in cells_a.items() if cells_b.get(pos) == value)
                cell_ratio = same_count / len(cells_a.keys() | cells_b.keys())
                common = sum((counters[name] & counters[other]).values())
                total = len(contents[name]["rows"]) + len(contents[other]["rows"])
                ratio = min(max(cell_ratio, 2 * common / total), 0.99)
            if ratio < SIMILARITY_MIN_RATIO:
                continue
            if best is None or (identical, ratio) > (best["identical"], best["ratio"]):
                best = {"match": other, "ratio": round(ratio, 2), "identical": identical}
                if same_cells and diff:
                    best.update(same_cells=True, diff=diff)
        if best is not None:
            result[name] = best
    return result


def _log_similar(file_name: str, sheet_name: str, content: dict[str, Any] | None, similar: dict[str, Any] | None) -> None:
    def count(key: str) -> Any:
        return "读不到" if not content or content.get(key) is None else len(content[key])

    if similar is None:
        result = "-"
    elif similar["identical"]:
        result = "完全相同"
    elif similar.get("same_cells"):
        result = f"文字相同但{'、'.join(similar['diff'])}不同"
    else:
        result = f"约{round(similar['ratio'] * 100)}%"
    debug_log.event("🟰 预览比对", 文件=file_name, 表=sheet_name, 最相似=similar["match"] if similar else "-",
                    结果=result, 图形批注数=count("objects"), 图片数=count("images"))


def _content_clip(page: Any) -> Any:
    """只截取页面上有内容的区域：表格只占 A4 左上角一小块时，缩略图里的字也能看清"""
    clip = fitz.Rect()
    for _, rect in page.get_bboxlog():
        r = fitz.Rect(rect)
        if not r.is_empty and not r.is_infinite:
            clip |= r
    if clip.is_empty:
        return page.rect
    m = PREVIEW_CLIP_MARGIN
    clip = fitz.Rect(clip.x0 - m, clip.y0 - m, clip.x1 + m, clip.y1 + m) & page.rect
    return clip if not clip.is_empty else page.rect


def _render_sheet(sheet: Any, pdf_path: Path, png_path: Path) -> dict[str, Any]:
    """导出单个工作表的缩略图；隐藏的表临时取消隐藏（只读打开、不保存，不影响原文件）"""
    hidden = sheet.Visible != XL_SHEET_VISIBLE
    result: dict[str, Any] = {
        "name": sheet.Name, "image": "", "pdf": "", "pages": 0, "hidden": hidden, "note": "",
    }
    if hidden:
        try:
            sheet.Visible = XL_SHEET_VISIBLE
        except Exception:
            result["note"] = "隐藏的工作表，无法取消隐藏来生成预览"
            return result
    try:
        _export_sheet_pdf(sheet, pdf_path)
    except Exception:
        result["note"] = "没有可打印的内容（可能是空白表）"
        return result
    try:
        result["pdf"] = pdf_path.name
        with fitz.open(pdf_path) as doc:
            result["pages"] = doc.page_count
            if doc.page_count:
                page = doc[0]
                clip = _content_clip(page)
                zoom = min(PREVIEW_MAX_ZOOM, max(1.0, PREVIEW_TARGET_WIDTH / clip.width))
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
                pix.save(str(png_path))
                result["image"] = png_path.name
    except Exception as err:
        result["note"] = f"预览图生成失败: {err}"
    if not result["image"] and not result["note"]:
        result["note"] = "没有可打印的内容（可能是空白表）"
    return result


def render_workbook_previews(folder_str: str, file_name: str) -> dict[str, Any]:
    """渲染一个工作簿全部工作表的第 1 页缩略图并比较内容，返回 {key, sheets: [{name, image, pages, hidden, note, similar}]}"""
    _, wb_path = _find_workbook(folder_str, file_name)
    key = _cache_key(wb_path)
    cache_dir = PREVIEW_CACHE_DIR / key
    with _render_locks_guard:
        lock = _render_locks.setdefault(key, threading.Lock())
    with lock:
        manifest = _read_manifest(cache_dir)
        if manifest is not None:
            return {"key": key, "cached": True, **manifest}

        _prune_old_cache(keep=key)
        shutil.rmtree(cache_dir, ignore_errors=True)
        staging = PREVIEW_CACHE_DIR / f".{key}.tmp"
        shutil.rmtree(staging, ignore_errors=True)
        staging.mkdir(parents=True)

        with debug_log.side_task("clean_render", {"file": file_name}):
            sheets: list[dict[str, Any]] = []
            pythoncom.CoInitialize()
            excel = None
            wb = None
            try:
                excel = start_background_excel()
                wb = excel.Workbooks.Open(str(wb_path.resolve()), ReadOnly=True, UpdateLinks=0)
                contents: dict[str, dict[str, Any] | None] = {}
                for idx, sheet in enumerate(list(wb.Sheets)):
                    content = read_sheet_content(sheet)
                    entry = _render_sheet(sheet, staging / f"{idx}.pdf", staging / f"{idx}.png")
                    if content is not None:
                        content["objects"] = read_sheet_objects(sheet)
                        content["images"] = pdf_image_digests(staging / entry["pdf"]) if entry["pdf"] else None
                    contents[sheet.Name] = content
                    sheets.append(entry)
                similar = find_similar_sheets(contents)
                for entry in sheets:
                    entry["similar"] = similar.get(entry["name"])
                    _log_similar(file_name, entry["name"], contents.get(entry["name"]), entry["similar"])
            except Exception:
                shutil.rmtree(staging, ignore_errors=True)
                raise
            finally:
                try:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
                finally:
                    try:
                        if excel is not None:
                            quit_background_excel(excel)
                    finally:
                        pythoncom.CoUninitialize()

        manifest = {"sheets": sheets}
        (staging / MANIFEST_NAME).write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
        staging.rename(cache_dir)
        return {"key": key, "cached": False, **manifest}


def preview_image_path(key: str, image_name: str) -> Path | None:
    """缩略图或导出的 PDF 路径；键或文件名格式不对、文件不存在时返回 None"""
    if not CACHE_KEY_RE.match(key) or not IMAGE_NAME_RE.match(image_name):
        return None
    path = PREVIEW_CACHE_DIR / key / image_name
    return path if path.is_file() else None
