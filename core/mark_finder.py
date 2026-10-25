"""查找旧版本留下的标记颜色：没有记录时，按「底色、字色和标记色完全一致」把格子找出来给人看。

旧版本直接改单元格的底色和字色，记录丢了以后程序认不出、也不知道原来的颜色，所以这里只列出来，不改文件。
新版本的标记是条件格式规则，公式里带标识，点「清除」就能按标识删掉，不用这里找。
用 openpyxl 只读文件内容，不开 Excel（文件在 Excel 里开着也能读）；.xls / .xlsb 读不了，列为需要人工看。
"""

from __future__ import annotations

import threading
from pathlib import Path
from typing import Any

from core import debug_log, workflow as wf
from core.common import workbook_rel_name

READABLE = {".xlsx", ".xlsm", ".xltx", ".xltm"}
MAX_HITS_PER_FILE = 500

_LOCK = threading.Lock()
_STATUS: dict[str, Any] = {"running": False}


def _rgb(color: Any) -> str | None:
    """openpyxl 的颜色 → RRGGBB（只认直接写的 RGB；主题色、索引色不是程序写进去的）"""
    if color is None or getattr(color, "type", None) != "rgb" or not isinstance(color.rgb, str):
        return None
    return color.rgb[-6:].upper()


def _styles(colors: list[dict]) -> list[tuple[str, str]]:
    seen: list[tuple[str, str]] = []
    for item in colors:
        if not isinstance(item, dict):
            continue
        fill, font = str(item.get("fill", "")), str(item.get("font", ""))
        if len(fill) == 7 and fill[0] == "#" and len(font) == 7 and font[0] == "#":
            try:
                int(fill[1:], 16), int(font[1:], 16)
            except ValueError:
                continue
            pair = (fill[1:].upper(), font[1:].upper())
            if pair not in seen:
                seen.append(pair)
    if not seen:
        raise ValueError("没有可以查找的标记颜色")
    return seen


def _recorded(path: str) -> set[tuple[str, str]]:
    """程序里还有旧方式记录的格子（表名, 左上角地址）：点「清除」就能按记录还原"""
    with wf.STATE_LOCK:
        marks = wf.STATE.get("batch_review", {}).get("highlights", {}).values()
        return {(m["sheet"], str(m["addr"]).replace("$", "").split(":")[0])
                for m in marks if m.get("path") == path and "cells" in m}


def scan_file(path: Path, styles: list[tuple[str, str]], recorded: set[tuple[str, str]]) -> list[dict]:
    """一份文件里底色是标记色的格子（合并区域算一格）。match：both 底色字色都对上 / fill 只有底色对上"""
    import openpyxl

    fills = {fill: font for fill, font in styles}
    hits: list[dict] = []
    book = openpyxl.load_workbook(path, data_only=False)
    try:
        for sheet in book.worksheets:
            merged = {}
            for area in sheet.merged_cells.ranges:
                for row in range(area.min_row, area.max_row + 1):
                    for col in range(area.min_col, area.max_col + 1):
                        merged[(row, col)] = area
            seen = set()
            for row in sheet.iter_rows():
                for cell in row:
                    fill = cell.fill
                    if fill is None or fill.patternType != "solid":
                        continue
                    color = _rgb(fill.fgColor)
                    if color not in fills:
                        continue
                    area = merged.get((cell.row, cell.column))
                    key = str(area) if area is not None else cell.coordinate
                    if key in seen:
                        continue
                    seen.add(key)
                    top = key.split(":")[0]
                    font = _rgb(cell.font.color) if cell.font is not None else None
                    hits.append({"sheet": sheet.title, "addr": key, "fill": f"#{color}",
                                 "font": f"#{font}" if font else "",
                                 "match": "both" if font == fills[color] else "fill",
                                 "recorded": (sheet.title, top) in recorded})
                    if len(hits) >= MAX_HITS_PER_FILE:
                        return hits
    finally:
        book.close()
    return hits


def _run(folder: Path, paths: list[Path], styles: list[tuple[str, str]]) -> None:
    files: list[dict] = []
    try:
        for i, path in enumerate(paths, 1):
            name = workbook_rel_name(path, folder)
            with _LOCK:
                _STATUS.update(done=i - 1, current=name)
            entry = {"path": str(path), "name": name, "hits": [], "error": None}
            if path.suffix.lower() not in READABLE:
                entry["error"] = f"{path.suffix} 格式读不了颜色，请人工看一下"
            else:
                try:
                    entry["hits"] = scan_file(path, styles, _recorded(str(path)))
                    entry["truncated"] = len(entry["hits"]) >= MAX_HITS_PER_FILE
                except Exception as err:
                    debug_log.exception(err, f"查找残留标记色 {name}")
                    entry["error"] = f"读取失败：{err}"
            if entry["hits"] or entry["error"]:
                files.append(entry)
        hits = sum(len(f["hits"]) for f in files)
        wf.log_msg(f"🔍 查找残留标记色：{len(paths)} 份文件，{hits} 格底色是标记色"
                   + (f"，{sum(1 for f in files if f['error'])} 份文件读不了" if any(f["error"] for f in files) else ""))
    finally:
        with _LOCK:
            _STATUS.update(running=False, done=len(paths), current="", result={
                "folder": str(folder), "total": len(paths), "files": files,
                "colors": [{"fill": f"#{a}", "font": f"#{b}"} for a, b in styles]})


def start(folder: str, colors: Any) -> None:
    styles = _styles(colors if isinstance(colors, list) else [])
    root, paths = wf.list_workbooks(folder)
    with _LOCK:
        if _STATUS.get("running"):
            raise ValueError("正在查找，请等这次查完")
        _STATUS.clear()
        _STATUS.update(running=True, total=len(paths), done=0, current="", result=None)
    wf.log_msg(f"🔍 开始查找残留标记色：{len(paths)} 份文件（只列出来，不改文件）")
    threading.Thread(target=_run, args=(root, paths, styles), name="查找残留标记色", daemon=True).start()


def status() -> dict:
    with _LOCK:
        return dict(_STATUS)
