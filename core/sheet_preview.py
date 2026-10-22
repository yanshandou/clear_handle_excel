"""精简 Sheet 的预览确认窗口：用 Excel 只读打开工作簿，把每个工作表导出成临时 PDF，再用 pymupdf 渲染第 1 页缩略图。

缩略图缓存在 _sheet_preview_cache/<键>/ 下，键由文件路径、修改时间与大小决定；文件没改过时再次打开直接读缓存。
"""

import hashlib
import json
import re
import shutil
import tempfile
import threading
import time
from collections import Counter
from pathlib import Path
from typing import Any

from core.common import PROJECT_ROOT, XL_SHEET_VISIBLE, normalize_used_values, workbook_rel_name
from core.excel_ops import pythoncom, win32
from core.pdf_detector import fitz
from core.workflow import _export_sheet_pdf, _preview_workbook, list_workbooks

PREVIEW_CACHE_DIR = PROJECT_ROOT / "_sheet_preview_cache"
PREVIEW_TARGET_WIDTH = 1000  # 缩略图宽度（px）：页面上按卡片宽度缩小显示，点开大图时仍然清晰
PREVIEW_MAX_ZOOM = 3.0
PREVIEW_CLIP_MARGIN = 12  # 裁掉页面空白时四周保留的边距（pt）
CACHE_MAX_AGE_SECONDS = 7 * 24 * 3600
MANIFEST_NAME = "manifest.json"
CACHE_KEY_RE = re.compile(r"^[0-9a-f]{16}$")
IMAGE_NAME_RE = re.compile(r"^\d{1,4}\.png$")
SIMILARITY_MIN_RATIO = 0.8  # 低于这个相似度不提示
SIMILARITY_MAX_CELLS = 500_000  # 已用区域超过这么多单元格时跳过比对，避免读取过慢

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
            holder["excel"] = win32.DispatchEx("Excel.Application")
            holder["excel"].Visible = False
            holder["excel"].DisplayAlerts = False
        return holder["excel"]

    try:
        entry = _preview_workbook(wb_path, folder, get_excel)
    finally:
        if "excel" in holder:
            try:
                holder["excel"].Quit()
            finally:
                pythoncom.CoUninitialize()
    return {"folder": str(folder), "file": entry}


def _cache_key(wb_path: Path) -> str:
    stat = wb_path.stat()
    raw = f"{wb_path.resolve()}|{stat.st_mtime_ns}|{stat.st_size}"
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


def find_similar_sheets(contents: dict[str, dict[str, Any] | None]) -> dict[str, dict[str, Any]]:
    """两两比较工作表内容，给每个表找出最相似的另一个表

    完全相同：每个单元格的位置和值都一样。
    相似度：按「行」比较（行内容相同即算一致，插入或删除几行也不影响其余行），2×相同行数 ÷ 两表行数之和。
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
            identical = contents[name]["cells"] == contents[other]["cells"]
            if identical:
                ratio = 1.0
            else:
                common = sum((counters[name] & counters[other]).values())
                total = len(contents[name]["rows"]) + len(contents[other]["rows"])
                ratio = min(2 * common / total, 0.99)
            if ratio < SIMILARITY_MIN_RATIO:
                continue
            if best is None or (identical, ratio) > (best["identical"], best["ratio"]):
                best = {"match": other, "ratio": round(ratio, 2), "identical": identical}
        if best is not None:
            result[name] = best
    return result


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
    result: dict[str, Any] = {"name": sheet.Name, "image": "", "pages": 0, "hidden": hidden, "note": ""}
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

        sheets: list[dict[str, Any]] = []
        pythoncom.CoInitialize()
        excel = None
        wb = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(str(wb_path.resolve()), ReadOnly=True, UpdateLinks=0)
            contents: dict[str, dict[str, Any] | None] = {}
            with tempfile.TemporaryDirectory() as tmp:
                for idx, sheet in enumerate(list(wb.Sheets)):
                    contents[sheet.Name] = read_sheet_content(sheet)
                    sheets.append(_render_sheet(sheet, Path(tmp) / f"{idx}.pdf", staging / f"{idx}.png"))
            similar = find_similar_sheets(contents)
            for entry in sheets:
                entry["similar"] = similar.get(entry["name"])
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
                        excel.Quit()
                finally:
                    pythoncom.CoUninitialize()

        manifest = {"sheets": sheets}
        (staging / MANIFEST_NAME).write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
        staging.rename(cache_dir)
        return {"key": key, "cached": False, **manifest}


def preview_image_path(key: str, image_name: str) -> Path | None:
    """缩略图文件路径；键或文件名格式不对、文件不存在时返回 None"""
    if not CACHE_KEY_RE.match(key) or not IMAGE_NAME_RE.match(image_name):
        return None
    path = PREVIEW_CACHE_DIR / key / image_name
    return path if path.is_file() else None
