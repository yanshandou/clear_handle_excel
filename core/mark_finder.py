"""查找旧版本留下的标记颜色：没有记录时，按「底色、字色和标记色完全一致」把格子找出来给人看。

旧版本直接改单元格的底色和字色，记录丢了以后程序认不出、也不知道原来的颜色，所以这里只列出来，不改文件。
新版本的标记是条件格式规则，公式里带标识，点「清除」就能按标识删掉，不用这里找。
.xlsx 等用 openpyxl 只读文件内容，不开 Excel（文件在 Excel 里开着也能读）；
.xls / .xlsb openpyxl 读不了，改在后台隐藏的 Excel 里只读打开来读颜色（慢一些）。
"""

from __future__ import annotations

import tempfile
import threading
from pathlib import Path
from typing import Any

from core import debug_log, workflow as wf
from core.backup_compare import norm as _norm, pair_files as _pair_files, pair_sheets
from core.common import workbook_rel_name

READABLE = {".xlsx", ".xlsm", ".xltx", ".xltm"}
EXCEL_READABLE = {".xls", ".xlsb"}  # openpyxl 读不了，用 Excel 读
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


XL_SOLID = 1
XL_EXCEL8 = 56  # .xls


def _hex_of(color: Any) -> str | None:
    """Excel 的颜色数（BGR）→ RRGGBB；None 表示这块区域里颜色不一样"""
    if color is None:
        return None
    value = int(color)
    return f"{value & 0xFF:02X}{value >> 8 & 0xFF:02X}{value >> 16 & 0xFF:02X}"


def xls_saved_colors(excel: Any, styles: list[tuple[str, str]]) -> dict[tuple[str, str], tuple[str, str]]:
    """.xls 只能存 56 种颜色：标记色存进 .xls 时会被 Excel 换成最接近的那一种，再打开读到的就不是原来的颜色。
    这里把标记色写进一个临时 .xls、存好再打开，读出 Excel 实际换成的颜色。返回 原来的 → 存进 .xls 后的"""
    with tempfile.TemporaryDirectory() as folder:
        path = str(Path(folder) / "colors.xls")
        book = excel.Workbooks.Add()
        try:
            try:
                book.CheckCompatibility = False  # 不弹「兼容性检查器」
            except Exception:
                pass
            sheet = book.Worksheets(1)
            for row, (fill, font) in enumerate(styles, 1):
                cell = sheet.Cells(row, 1)
                cell.Value = "Aa"
                cell.Interior.Color = int(_excel_color(f"#{fill}"))
                cell.Font.Color = int(_excel_color(f"#{font}"))
            book.SaveAs(path, FileFormat=XL_EXCEL8)
        finally:
            book.Close(SaveChanges=False)
        book = excel.Workbooks.Open(path, ReadOnly=True)
        try:
            sheet = book.Worksheets(1)
            saved = {}
            for row, pair in enumerate(styles, 1):
                cell = sheet.Cells(row, 1)
                saved[pair] = (_hex_of(cell.Interior.Color) or pair[0], _hex_of(cell.Font.Color) or pair[1])
            return saved
        finally:
            book.Close(SaveChanges=False)


def _same_color_blocks(used: Any) -> Any:
    """把已用区域切成一块块颜色一样的区域：整块颜色一样时 Interior.Color 直接给颜色，
    不一样时给 None，就对半切开再看。大部分表只有少数几种底色，比逐格读快得多"""
    stack = [used]
    while stack:
        area = stack.pop()
        color = area.Interior.Color
        if color is not None:
            yield area, _hex_of(color)
            continue
        rows, cols = area.Rows.Count, area.Columns.Count
        if rows > 1:
            half = rows // 2
            stack.append(area.Offset(half, 0).Resize(rows - half, cols))
            stack.append(area.Resize(half, cols))
        elif cols > 1:
            half = cols // 2
            stack.append(area.Offset(0, half).Resize(rows, cols - half))
            stack.append(area.Resize(rows, half))


def scan_file_excel(excel: Any, path: Path, styles: list[tuple[str, str]], recorded: set[tuple[str, str]],
                    saved: dict[tuple[str, str], tuple[str, str]]) -> list[dict]:
    """.xls / .xlsb：在后台 Excel 里只读打开来找，规则和 scan_file 一样（合并区域算一格）。
    .xls 里的标记色可能已经被换成 56 色里最接近的那种（saved），这种也算，标 approx"""
    from core.excel_ops import open_workbook_checked

    fills: dict[str, tuple[str, bool]] = {}
    for fill, font in styles:
        fills.setdefault(fill, (font, False))
        if path.suffix.lower() == ".xls":
            near_fill, near_font = saved.get((fill, font), (fill, font))
            fills.setdefault(near_fill, (near_font, near_fill != fill))
    hits: list[dict] = []
    book = open_workbook_checked(excel, str(path.resolve()), UpdateLinks=0, ReadOnly=True)
    try:
        for sheet in book.Worksheets:
            name, seen = sheet.Name, set()
            for area, color in _same_color_blocks(sheet.UsedRange):
                if color not in fills:
                    continue
                font_want, approx = fills[color]
                for cell in area.Cells:
                    if int(cell.Interior.Pattern) != XL_SOLID:
                        continue
                    # Address 在 pywin32 里是属性（带 $ 的字符串），不能带参数调用
                    key = str((cell.MergeArea if cell.MergeCells else cell).Address).replace("$", "")
                    if key in seen:
                        continue
                    seen.add(key)
                    top = key.split(":")[0]
                    font = _hex_of(sheet.Range(top).Font.Color)
                    hit = {"sheet": name, "addr": key, "fill": f"#{color}", "font": f"#{font}" if font else "",
                           "match": "both" if font == font_want else "fill", "recorded": (name, top) in recorded}
                    if approx:
                        hit["approx"] = True
                    hits.append(hit)
                    if len(hits) >= MAX_HITS_PER_FILE:
                        return hits
    finally:
        book.Close(SaveChanges=False)
    return hits


def _run(folder: Path, paths: list[Path], styles: list[tuple[str, str]]) -> None:
    files: list[dict] = []
    excel, saved, com = None, None, False
    try:
        for i, path in enumerate(paths, 1):
            name = workbook_rel_name(path, folder)
            with _LOCK:
                _STATUS.update(done=i - 1, current=name)
            entry = {"path": str(path), "name": name, "hits": [], "error": None}
            debug_log.begin(f"查找残留 {i} / {len(paths)}：{name}")
            if path.suffix.lower() in EXCEL_READABLE:
                try:
                    if saved is None:  # 第一份 .xls 时才开 Excel；开不了的话后面的也不再试
                        from core.excel_ops import start_background_excel
                        saved = {}
                        wf.pythoncom.CoInitialize()
                        com = True
                        excel = start_background_excel()
                        try:
                            saved = xls_saved_colors(excel, styles)
                        except Exception as err:
                            debug_log.exception(err, "试存 .xls 标记色")
                        debug_log.event("🎨 .xls 里的标记色", 换成={f"#{a}/#{b}": f"#{c}/#{d}"
                                                                    for (a, b), (c, d) in saved.items()} or "没试出来")
                    if excel is None:
                        raise RuntimeError("后台 Excel 没能启动")
                    entry["hits"] = scan_file_excel(excel, path, styles, _recorded(str(path)), saved or {})
                    entry["truncated"] = len(entry["hits"]) >= MAX_HITS_PER_FILE
                except Exception as err:
                    debug_log.exception(err, f"查找残留标记色（用 Excel 读）{name}")
                    entry["error"] = f"用 Excel 读取失败：{err}"
            elif path.suffix.lower() not in READABLE:
                entry["error"] = f"{path.suffix} 格式读不了颜色，请人工看一下"
            else:
                try:
                    entry["hits"] = scan_file(path, styles, _recorded(str(path)))
                    entry["truncated"] = len(entry["hits"]) >= MAX_HITS_PER_FILE
                except Exception as err:
                    debug_log.exception(err, f"查找残留标记色 {name}")
                    entry["error"] = f"读取失败：{err}"
            debug_log.event("📂 查找结果", 文件=name, 找到=len(entry["hits"]), 到上限=entry.get("truncated", False),
                            有记录=sum(1 for h in entry["hits"] if h["recorded"]), 错误=entry["error"] or "-")
            for hit in entry["hits"]:
                debug_log.event("🔍 残留", 位置=f"{hit['sheet']}!{hit['addr']}", 底色=hit["fill"], 字色=hit["font"] or "-",
                                对上=hit["match"], 有记录=hit["recorded"], xls近似色=hit.get("approx", False))
            if entry["hits"] or entry["error"]:
                files.append(entry)
        hits = sum(len(f["hits"]) for f in files)
        wf.log_msg(f"🔍 查找残留标记色：{len(paths)} 份文件，{hits} 格底色是标记色"
                   + (f"，{sum(1 for f in files if f['error'])} 份文件读不了" if any(f["error"] for f in files) else ""))
    finally:
        if excel is not None:
            try:
                from core.excel_ops import quit_background_excel
                quit_background_excel(excel)
            except Exception:
                pass
        if com:
            wf.pythoncom.CoUninitialize()
        with _LOCK:
            _STATUS.update(running=False, done=len(paths), current="", result={
                "folder": str(folder), "total": len(paths), "files": files,
                "colors": [{"fill": f"#{a}", "font": f"#{b}"} for a, b in styles]})


def _logged(kind: str, options: dict, func: Any, *args: Any) -> None:
    """后台线程跑的这几步也在调试日志里算一个任务，网页上能选出来复制"""
    with debug_log.side_task(kind, options):
        func(*args)


def _begin(phase: str, total: int) -> None:
    with wf.STATE_LOCK, _LOCK:
        if wf.STATE.get("busy"):
            raise ValueError("后台任务正在执行，请等它结束")
        if _STATUS.get("running"):
            raise ValueError("上一步还没做完，请稍等")
        _STATUS.update(running=True, phase=phase, total=total, done=0, current="")


def start(folder: str, colors: Any) -> None:
    styles = _styles(colors if isinstance(colors, list) else [])
    root, paths = wf.list_workbooks(folder)
    _begin("find", len(paths))
    with _LOCK:
        for key in ("result", "plan", "applied"):
            _STATUS.pop(key, None)
    wf.log_msg(f"🔍 开始查找残留标记色：{len(paths)} 份文件（只列出来，不改文件）")
    threading.Thread(target=_logged, args=("find_marks", {"folder": str(root), "files": len(paths), "colors": styles},
                                           _run, root, paths, styles), name="查找残留标记色", daemon=True).start()


# ---------- 对比旧版文件夹，把找到的格子改回旧版的颜色 ----------

def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value)


def _row_values(sheet: Any, row: int) -> list[str]:
    values = [_text(c.value) for c in sheet[row]] if row <= sheet.max_row else []
    while values and values[-1] == "":
        values.pop()
    return values


def _old_colors(cell: Any) -> dict:
    """旧版格子的颜色，给页面看：none 无填充；RRGGBB；other 主题色 / 图案（写回时由 Excel 照原样抄）"""
    fill = cell.fill
    if fill is None or fill.patternType is None:
        old_fill = "none"
    elif fill.patternType == "solid" and _rgb(fill.fgColor):
        old_fill = f"#{_rgb(fill.fgColor)}"
    else:
        old_fill = "other"
    font = _rgb(cell.font.color) if cell.font is not None else None
    return {"old_fill": old_fill, "old_font": f"#{font}" if font else ""}


SIMILAR_ROWS = 200
_HOW = {"guess": "去掉「新」/ 日期等后的表名", "code": "代码名", "content": "内容"}


def _sheet_words(ws: Any) -> set[str]:
    """表里前几百行出现过的内容（不管位置），用来认出改过名的表"""
    words: set[str] = set()
    for row in ws.iter_rows(max_row=min(ws.max_row, SIMILAR_ROWS), values_only=True):
        words.update(_text(v).strip() for v in row if _text(v).strip())
    return words


def _match_sheets(current: Any, old: Any) -> dict[str, tuple[Any, str]]:
    return pair_sheets(list(current.worksheets), list(old.worksheets),
                       title=lambda ws: ws.title, code=lambda ws: ws.sheet_properties.codeName,
                       words=_sheet_words)


def plan_file(path: Path, old_path: Path, hits: list[dict]) -> list[dict]:
    """逐格决定改不改：旧版同一张表、同一地址的格子内容一样才改回旧版的颜色。
    空格子内容一样说明不了什么（插了行、地址对不上时也是空的），还要这一整行的内容都一样。"""
    import openpyxl

    current = openpyxl.load_workbook(path, data_only=False)
    old = openpyxl.load_workbook(old_path, data_only=False)
    try:
        matched = _match_sheets(current, old)
        debug_log.event("📄 表对应", 新版=[ws.title for ws in current.worksheets],
                        对上={name: f"{ws.title}（{_HOW.get(how, how)}）" for name, (ws, how) in matched.items()},
                        旧版=[ws.title for ws in old.worksheets])
        plan = []
        for hit in hits:
            item = {"sheet": hit["sheet"], "addr": hit["addr"], "fill": hit["fill"], "font": hit["font"],
                    "action": "skip", "reason": ""}
            plan.append(item)
            top = hit["addr"].split(":")[0]
            sheet = current[hit["sheet"]]
            old_sheet, how = matched.get(hit["sheet"], (None, ""))
            if old_sheet is None:
                names = "、".join(ws.title for ws in old.worksheets)
                item["reason"] = f"旧版里找不到这张表（按表名、猜原名、代码名、内容都对不上；旧版的表：{names}）"
                continue
            item["old_sheet"] = old_sheet.title
            cell, old_cell = sheet[top], old_sheet[top]
            item.update(_old_colors(old_cell))
            if _text(cell.value) != _text(old_cell.value):
                item["reason"] = "和旧版同一位置的内容不一样（可能插删过行列，或改过内容），请人工看一下"
            elif cell.value in (None, "") and _row_values(sheet, cell.row) != _row_values(old_sheet, old_cell.row):
                item["reason"] = "空格子，这一行和旧版对不上（可能插删过行列），请人工看一下"
            elif item["old_fill"] == hit["fill"]:
                item["reason"] = "旧版本来就是这个颜色，不是程序标的，不改"
            else:
                item.update(action="restore", reason="改回旧版的颜色")
            if how != "name":
                item["reason"] += f"（旧版表名「{old_sheet.title}」，按{_HOW[how]}对上）"
        for item in plan:
            debug_log.event("🔍 对比", 位置=f"{item['sheet']}!{item['addr']}", 旧表=item.get("old_sheet", "-"),
                            动作=item["action"], 原因=item["reason"], 标记色=item["fill"],
                            旧底色=item.get("old_fill", "-"), 旧字色=item.get("old_font", "-"))
        return plan
    finally:
        current.close()
        old.close()


def _run_preview(old_root: Path, old_paths: list[Path], files: list[dict], folder: str = "") -> None:
    plan_files: list[dict] = []
    try:
        pairs = _pair_files(old_root, old_paths, [f["name"] for f in files], Path(folder) if folder else None)
        for i, file in enumerate(files, 1):
            with _LOCK:
                _STATUS.update(done=i - 1, current=file["name"])
            old_path = pairs[file["name"]]
            entry = {"path": file["path"], "name": file["name"], "old": str(old_path or ""), "cells": [], "error": None}
            debug_log.begin(f"对比旧版 {i} / {len(files)}：{file['name']}")
            debug_log.event("📂 配对", 文件=file["name"], 旧版=old_path or "没找到", 格子=len(file["hits"]))
            if Path(file["path"]).suffix.lower() not in READABLE:
                entry["error"] = f"{Path(file['path']).suffix} 文件还不能自动对比旧版，请对照上面的位置人工改回"
            elif old_path is None:
                entry["error"] = "旧版文件夹里没找到这份文件（相对路径不同，同名的又不止一个或没有）"
            elif old_path.suffix.lower() not in READABLE:
                entry["error"] = f"旧版文件是 {old_path.suffix}，读不了，请人工对照"
            else:
                try:
                    entry["cells"] = plan_file(Path(file["path"]), old_path, file["hits"])
                except Exception as err:
                    debug_log.exception(err, f"对比旧版 {file['name']}")
                    entry["error"] = f"读取失败：{err}"
            plan_files.append(entry)
        restore = sum(1 for f in plan_files for c in f["cells"] if c["action"] == "restore")
        wf.log_msg(f"🔍 对比旧版：{len(plan_files)} 份文件，{restore} 格可以改回旧版的颜色（还没改，等确认）")
    finally:
        with _LOCK:
            _STATUS.update(running=False, done=len(files), current="",
                           plan={"old_folder": str(old_root), "folder": folder, "files": plan_files})


def preview(old_folder: str) -> None:
    """对比旧版文件夹，算出每个找到的格子要不要改、改成什么颜色。只读，不改文件"""
    with _LOCK:
        result = _STATUS.get("result")
    if not result:
        raise ValueError("请先查找残留标记色")
    old_root, old_paths = wf.list_workbooks(old_folder)
    if old_root == Path(result["folder"]).resolve():
        raise ValueError("旧版文件夹和当前文件夹是同一个，请选原始备份的文件夹")
    files = [f for f in result["files"] if f["hits"]]
    if not files:
        raise ValueError("没有找到要改回的格子")
    _begin("preview", len(files))
    with _LOCK:
        _STATUS.pop("plan", None)
        _STATUS.pop("applied", None)
    threading.Thread(target=_logged, args=("compare_old", {"old_folder": str(old_root), "folder": result["folder"],
                                                           "files": len(files)},
                                           _run_preview, old_root, old_paths, files, result["folder"]),
                     name="对比旧版", daemon=True).start()


XL_PATTERN_NONE, XL_PATTERN_SOLID = -4142, 1


def _excel_color(hex_color: str) -> int:
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (1, 3, 5))
    return r + g * 256 + b * 65536


def _read_old(excel: Any, old_path: str, cells: list[dict]) -> dict[str, dict | str]:
    """先把旧版要用的格子读出来（内容、底色、字色）就关掉：旧版和新版多半同名，
    Excel 不能同时打开两个同名文件（第二个 Open 会返回空对象）。读不到的格子记原因"""
    from core.excel_ops import open_workbook_checked

    found: dict[str, dict | str] = {}
    old = open_workbook_checked(excel, str(Path(old_path).resolve()), UpdateLinks=0, ReadOnly=True)
    try:
        sheets = {_norm(sheet.Name): sheet for sheet in old.Worksheets}
        for item in cells:
            key = f"{item['sheet']}!{item['addr']}"
            try:
                source = sheets[_norm(item.get("old_sheet", item["sheet"]))].Range(item["addr"].split(":")[0])
                interior = source.Interior
                found[key] = {"formula": _text(source.Formula), "pattern": int(interior.Pattern),
                              "color": interior.Color, "tint": interior.TintAndShade,
                              "pattern_color": interior.PatternColor, "font": source.Font.Color}
            except Exception as err:
                found[key] = f"读旧版这一格出错：{err}"
    finally:
        old.Close(SaveChanges=False)
    return found


def apply_file(excel: Any, path: str, old_path: str, cells: list[dict]) -> list[dict]:
    """用 Excel 把旧版格子的底色（含主题色、图案）和字色抄回来，保存。写之前再确认一次：
    内容还和旧版一样、底色还是标记色（期间被人改过就不动）。只改颜色，不动别的格式"""
    from core.excel_ops import normalize_workbook_view, open_workbook_checked

    old = _read_old(excel, old_path, cells)
    done: list[dict] = []
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开（可能在别处开着），没有改")
        changed = False
        for item in cells:
            result = {**item, "status": "failed", "reason": ""}
            done.append(result)
            source = old[f"{item['sheet']}!{item['addr']}"]
            if not isinstance(source, str):
                debug_log.event("📄 旧版这一格", 位置=f"{item['sheet']}!{item['addr']}", 内容=debug_log.clip(source["formula"]),
                                图案=source["pattern"], 底色=source["color"], 深浅=source["tint"], 字色=source["font"])
            if isinstance(source, str):
                result["reason"] = source
                continue
            try:
                target = book.Worksheets(item["sheet"]).Range(item["addr"])
                top = target.Cells(1, 1)
                if _text(top.Formula) != source["formula"]:
                    result["reason"] = "内容和旧版不一样了，没有改"
                    continue
                color = top.Interior.Color
                if color is None or int(color) != _excel_color(item["fill"]):
                    result["reason"] = "底色已经不是标记色（被人改过），没有改"
                    continue
                interior = target.Interior
                if source["pattern"] == XL_PATTERN_NONE:
                    interior.Pattern = XL_PATTERN_NONE
                else:
                    interior.Pattern = source["pattern"]
                    interior.Color = source["color"]
                    interior.TintAndShade = source["tint"]
                    if source["pattern"] != XL_PATTERN_SOLID:
                        interior.PatternColor = source["pattern_color"]
                if source["font"] is not None:
                    target.Font.Color = source["font"]
                changed = True
                result.update(status="ok", reason="已改回旧版的颜色")
            except Exception as err:
                result["reason"] = f"改的时候出错：{err}"
        if changed:
            normalize_workbook_view(excel, book)
            book.Save()
    finally:
        book.Close(SaveChanges=False)
    return done


def _run_apply(files: list[dict], folder: str = "") -> None:
    from core.excel_ops import close_saved_workbooks, quit_background_excel, start_background_excel

    applied: list[dict] = []
    excel = None
    wf.pythoncom.CoInitialize()
    try:
        unsaved = close_saved_workbooks([Path(f["path"]) for f in files])
        excel = start_background_excel()
        for i, file in enumerate(files, 1):
            with _LOCK:
                _STATUS.update(done=i - 1, current=file["name"])
            entry = {"path": file["path"], "name": file["name"], "cells": [], "error": None}
            applied.append(entry)
            debug_log.begin(f"改回旧版颜色 {i} / {len(files)}：{file['name']}")
            if Path(file["path"]) in unsaved:
                entry["error"] = f"{unsaved[Path(file['path'])]}，请保存并关闭后再试"
                continue
            try:
                entry["cells"] = apply_file(excel, file["path"], file["old"], file["cells"])
            except Exception as err:
                debug_log.exception(err, f"改回旧版颜色 {file['name']}")
                entry["error"] = f"没有改：{err}"
            for cell in entry["cells"]:
                debug_log.event("🎨 改回" if cell["status"] == "ok" else "⚠️ 没改", 位置=f"{cell['sheet']}!{cell['addr']}",
                                结果=cell["reason"])
            ok = sum(1 for c in entry["cells"] if c["status"] == "ok")
            wf.log_msg(f"🎨 {file['name']}：改回旧版颜色 {ok} 格" + (f"，{entry['error']}" if entry["error"] else ""))
    except Exception as err:
        debug_log.exception(err, "改回旧版颜色")
        wf.log_msg(f"❌ 改回旧版颜色失败：{err}")
        applied.append({"path": "", "name": "（全部）", "cells": [], "error": str(err)})
    finally:
        try:
            quit_background_excel(excel)
        except Exception:
            pass
        wf.pythoncom.CoUninitialize()
        ok = sum(1 for f in applied for c in f["cells"] if c["status"] == "ok")
        wf.log_msg(f"🎨 改回旧版颜色完成：{ok} 格")
        with _LOCK:
            _STATUS.update(running=False, done=len(files), current="", applied={"folder": folder, "files": applied})


def apply() -> None:
    """按预览清单改文件（清单在服务端，页面只是确认）"""
    with _LOCK:
        plan = _STATUS.get("plan")
    if not plan:
        raise ValueError("请先对比旧版，看过清单再确认")
    with wf.STATE_LOCK:
        if wf.STATE.get("busy"):
            raise ValueError("后台任务正在执行，等它结束再改")
    files = [{**f, "cells": [c for c in f["cells"] if c["action"] == "restore"]} for f in plan["files"] if not f["error"]]
    files = [f for f in files if f["cells"]]
    if not files:
        raise ValueError("清单里没有可以改回的格子")
    _begin("apply", len(files))
    with _LOCK:
        _STATUS.pop("plan", None)  # 一份清单只执行一次
    wf.log_msg(f"🎨 开始改回旧版颜色：{len(files)} 份文件，{sum(len(f['cells']) for f in files)} 格")
    threading.Thread(target=_logged, args=("restore_old", {"folder": plan.get("folder", ""), "files": len(files)},
                                           _run_apply, files, plan.get("folder", "")),
                     name="改回旧版颜色", daemon=True).start()


def status() -> dict:
    with _LOCK:
        return dict(_STATUS)
