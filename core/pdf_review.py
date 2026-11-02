"""PDF 人工复检：导出的 PDF 整页显示在网页上，看到问题点一下，就用 Excel 打开并选中那一格；
填了原始备份文件夹时，备份也一起打开（只读）、跳到同一格，左右对照着改。

点的位置怎么对回单元格：PDF 只是画出来的图，不知道哪块是哪一格。后台只读打开工作簿，在内存里
把打印范围的每一行涂成各不相同的底色导出一次 PDF，再把每一列涂成不同底色导出一次（都不保存）。
读这两份 PDF 里的底色矩形，就知道每页上每一行、每一列占的位置，点的位置落在哪一行、哪一列就是哪一格。
导出用的打印方向和正式导出一样（print_layout），所以页码和位置对得上；页数或页面大小对不上时
说明 Excel 改过、PDF 已经旧了，不硬套。打印标题行、隐藏行列、横向分页都按 Excel 实际打印的样子来。

每份 PDF 只算一次（按 PDF 的修改时间记住；之后在 Excel 里改了内容，PDF 上看到的还是原来的样子，照旧对得上）。
页面一加载，后台就按顺序一份份算；点到还没算的那份时插到最前面先算。
"""
from __future__ import annotations

import re
import tempfile
import threading
import time
from collections import Counter
from pathlib import Path
from typing import Any

from core import batch, debug_log, file_problem, pdf_store, print_layout, repair_compare, sheet_history, workflow as wf
from core.backup_compare import pair_files
from core.common import (EXCEL_EXTENSIONS, col_index_to_letters, sanitize_filename, strip_review_prefix,
                         workbook_locked, workbook_rel_name)
from core.excel_ops import open_workbook_checked, open_workbook_in_excel, quit_background_excel, start_background_excel
from core.excel_watchdog import ExcelWatchdog
from core.formula_links import delete_sheets_preserving_values
from core.pdf_detector import fitz

PAGE_ZOOM = 2.5  # 网页上显示的页面图：A4 竖向约 1490 像素宽，铺满屏幕也看得清字
MAX_PROBE_LINES = 20000  # 打印范围超过这么多行 + 列就不涂色定位了（逐行涂色太慢）
OLD_LIST_SECONDS = 60  # 备份文件夹里的文件清单记一分钟，连着点不用每次都遍历
_PAGES_CACHE: dict[tuple[str, int, int], list[list[float]]] = {}
# 每份 PDF 的行列位置：pdf -> {"v": PDF 修改时间, "state": "ready" / "error", "sheet", "rows", "cols", "error"}
_MAPS: dict[str, dict] = {}
_QUEUE: list[str] = []  # 等着算的 PDF，前面的先算
_COND = threading.Condition()
_WORKING: list[str] = []  # 正在算的那一批（同一份 Excel 的几张表一起算）
_WORKER: threading.Thread | None = None
_OLD_FILES: dict[str, tuple[float, list[Path]]] = {}


def _exports() -> dict[str, dict]:
    """导出 PDF 时记下的（workflow._export_workbook_pdfs）：PDF 路径 -> {wb_path, sheet}"""
    return wf.STATE.setdefault("pdf_exports", {})


def _in_folder(path: Path) -> bool:
    with wf.STATE_LOCK:
        root = wf.STATE.get("pdf_review", {}).get("folder")
    return bool(root) and path.is_file() and Path(root) in path.parents


def forget_pdfs(paths: list[Path]) -> None:
    """导出前删掉了旧 PDF：它们的导出记录和算好的行列位置也去掉"""
    gone = {str(p) for p in paths} | {str(Path(p).resolve()) for p in paths}
    with wf.STATE_LOCK:
        for pdf in gone:
            _exports().pop(pdf, None)
        wf._persist_state()
    with _COND:
        for pdf in gone:
            _MAPS.pop(pdf, None)
            if pdf in _QUEUE:
                _QUEUE.remove(pdf)


# ---------- 列出 PDF ----------

def _resolve_folder(folder: str) -> Path:
    root = Path(folder.strip().strip('"').strip("'")).expanduser().resolve()
    if not root.is_dir():
        raise ValueError(f"文件夹不存在: {root}")
    return root


def _guess_owner(pdf: Path) -> dict | None:
    """没有导出记录（旧版本导出的、扫描时导出的）：按「Excel 文件名_表名.pdf」找 Excel。
    PDFS 里的 PDF 到对应的 Excel 文件夹找（PDFS/sub/x.pdf → sub/），旧的混放 PDF 在同一个文件夹里找"""
    best: tuple[int, Path, str] | None = None
    folder = pdf_store.excel_folder_for(pdf)
    if not folder.is_dir():
        return None
    for wb in folder.iterdir():
        if not wb.is_file() or wb.suffix.lower() not in EXCEL_EXTENSIONS | {".xlsb"} or wb.name.startswith("~$"):
            continue
        stem = sanitize_filename(strip_review_prefix(wb.stem))
        if pdf.stem.startswith(stem + "_") and len(pdf.stem) > len(stem) + 1:
            if best is None or len(stem) > best[0]:
                best = (len(stem), wb, pdf.stem[len(stem) + 1:])
    return {"wb_path": str(best[1]), "sheet": best[2], "guessed": True} if best else None


def _owner(pdf: Path) -> dict | None:
    record = _exports().get(str(pdf))
    if record and Path(record["wb_path"]).is_file():
        return dict(record)
    return pdf_store.owner(pdf) or _guess_owner(pdf)


def _page_sizes(pdf: Path) -> list[list[float]]:
    stat = pdf.stat()
    key = (str(pdf), stat.st_mtime_ns, stat.st_size)
    if key not in _PAGES_CACHE:
        with fitz.open(pdf) as doc:
            _PAGES_CACHE[key] = [[round(page.rect.width, 2), round(page.rect.height, 2)] for page in doc]
        if len(_PAGES_CACHE) > 2000:
            _PAGES_CACHE.pop(next(iter(_PAGES_CACHE)))
    return _PAGES_CACHE[key]


def _pdf_files(root: Path) -> list[Path]:
    """有 PDFS 文件夹就只看里面的（导出的 PDF 都在这）；没有的话（旧版本导出、和 Excel 混放的）看整个文件夹"""
    base = pdf_store.pdf_root(root)
    where = base if base.is_dir() else root
    return sorted((p for p in where.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"
                   and not p.name.startswith(("~", "."))), key=lambda p: workbook_rel_name(p, where))


def list_pdfs(folder: str) -> dict:
    """文件夹里的 PDF（有 PDFS 文件夹就只看 PDFS 里的），每份带对应的 Excel 和表、各页大小（PDF 点）。
    列完就在后台按顺序算每份 PDF 的行列位置，用户点的时候多半已经算好了"""
    root = _resolve_folder(folder)
    pdfs = _pdf_files(root)
    base = pdf_store.pdf_root(root)
    shown_from = base if base.is_dir() else root
    entries = []
    with wf.STATE_LOCK:
        wf.STATE["pdf_review"] = {"folder": str(root)}  # 以前画框标色留下的记录不要了
        for gone in [p for p in _exports() if Path(p).parent == root or root in Path(p).parents]:
            if not Path(gone).is_file():
                _exports().pop(gone)  # PDF 已经删掉的记录不留着
        owners = {str(p): _owner(p) for p in pdfs}
    for pdf in pdfs:
        entry: dict[str, Any] = {"pdf": str(pdf), "name": workbook_rel_name(pdf, shown_from), "v": pdf.stat().st_mtime_ns}
        try:
            entry["pages"] = _page_sizes(pdf)
        except Exception as err:
            entry.update(pages=[], error=f"PDF 打不开：{err}")
        owner = owners[str(pdf)]
        if owner:
            entry.update(wb_path=owner["wb_path"], wb_name=workbook_rel_name(Path(owner["wb_path"]), root),
                         sheet=owner["sheet"], guessed=bool(owner.get("guessed")))
        entries.append(entry)
    debug_log.event("🔎 PDF 复检列出", 文件夹=root, 看的是=shown_from, PDF=len(entries), 找到Excel=sum(1 for e in entries if e.get("wb_path")))
    prepare([e["pdf"] for e in entries if e.get("wb_path") and not e.get("error")])
    return {"folder": str(root), "pdf_folder": str(shown_from), "in_pdfs": shown_from == base,
            "pdfs": entries, "maps": maps_status()}


def page_png(pdf: str, page: int) -> bytes:
    path = Path(pdf).resolve()
    if not _in_folder(path) or path.suffix.lower() != ".pdf":
        raise ValueError("只能查看当前文件夹里的 PDF")
    with fitz.open(path) as doc:
        if not 0 <= page < len(doc):
            raise ValueError("页码不对")
        return doc[page].get_pixmap(matrix=fitz.Matrix(PAGE_ZOOM, PAGE_ZOOM), alpha=False).tobytes("png")


# ---------- 页面上的位置 → 单元格：探测行列位置 ----------

def probe_color(index: int) -> int:
    """第 index 行 / 列（从 1 起）的探测底色，Excel 的 Color（BGR）。每个通道只用 2、6、10…250
    （63 级，不用 254：白色 255 差 1 会被认成探测色），PDF 里读回来差 1 也认得出"""
    r, g, b = (2 + 4 * (index // 63 ** power % 63) for power in (2, 1, 0))
    return r + g * 256 + b * 65536


def probe_index(fill: Any) -> int | None:
    """PDF 里矩形的填充色（0~1 的 RGB）→ 探测底色的行号 / 列号；不是探测色返回 None"""
    if not fill or len(fill) != 3:
        return None
    index = 0
    for value in fill:
        v = round(float(value) * 255)
        digit = round((v - 2) / 4)
        if not 0 <= digit <= 62 or abs(v - (2 + 4 * digit)) > 1:
            return None
        index = index * 63 + digit
    return index or None


def read_bands(pdf_path: Path | str, axis: str) -> list[dict[int, tuple[float, float]]]:
    """探测 PDF 每页上每一行（axis="y"）或每一列（"x"）的位置 {行号 / 列号: (起, 止)}。
    同一行的格子各画一块底色：取出现最多的那个范围；合并单元格的底色会跨几行，次数少、不会被选上。
    一样多时取最窄的"""
    bands: list[dict[int, tuple[float, float]]] = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            seen: dict[int, Counter] = {}
            for drawing in page.get_drawings():
                index = probe_index(drawing.get("fill"))
                if index is None:
                    continue
                rects = [item[1] for item in drawing.get("items", []) if item[0] == "re"] or [drawing["rect"]]
                for rect in rects:
                    lo, hi = (rect.y0, rect.y1) if axis == "y" else (rect.x0, rect.x1)
                    if hi - lo > 0.1:
                        seen.setdefault(index, Counter())[(round(lo, 1), round(hi, 1))] += 1
            bands.append({index: max(counts, key=lambda span: (counts[span], span[0] - span[1]))
                          for index, counts in seen.items()})
    return bands


def block_address(blocks: list[tuple[int, int, int, int]]) -> str:
    """单元格块 → Excel 地址（A1:C3,A7:C9）。Excel 的地址最长 255 个字，太长时用包住全部的一块"""
    def one(r1: int, r2: int, c1: int, c2: int) -> str:
        first = f"{col_index_to_letters(c1)}{r1}"
        return first if (r1, c1) == (r2, c2) else f"{first}:{col_index_to_letters(c2)}{r2}"

    address = ",".join(one(*b) for b in blocks)
    if len(address) > 250:
        address = one(min(b[0] for b in blocks), max(b[1] for b in blocks),
                      min(b[2] for b in blocks), max(b[3] for b in blocks))
    return address


def _area_bounds(rng: Any) -> list[tuple[int, int, int, int]]:
    result = []
    for i in range(1, int(rng.Areas.Count) + 1):
        area = rng.Areas(i)
        row, col = int(area.Row), int(area.Column)
        result.append((row, row + int(area.Rows.Count) - 1, col, col + int(area.Columns.Count) - 1))
    return result


def _probe_blocks(sheet: Any) -> list[tuple[int, int, int, int]]:
    """要涂色的范围：打印区域（没设时是已用区域），加上打印标题行 / 列。只涂这些格子，不会让打印范围变大"""
    setup = sheet.PageSetup
    area = str(setup.PrintArea or "")
    blocks = _area_bounds(sheet.Range(area) if area else sheet.UsedRange)
    extra = []
    for title, is_rows in ((str(setup.PrintTitleRows or ""), True), (str(setup.PrintTitleColumns or ""), False)):
        if not title:
            continue
        for t1, t2, u1, u2 in _area_bounds(sheet.Range(title)):
            for r1, r2, c1, c2 in blocks:
                extra.append((t1, t2, c1, c2) if is_rows else (r1, r2, u1, u2))
    return blocks + extra


def _paint(sheet: Any, blocks: list[tuple[int, int, int, int]], numbers: list[int], by_rows: bool, tick: Any) -> None:
    """每一行（列）涂成自己的探测色：第几个要涂的行（列）就用第几号颜色，行号再大也不会超出颜色的个数"""
    code = {n: k for k, n in enumerate(numbers, start=1)}
    for r1, r2, c1, c2 in blocks:
        for i in (range(r1, r2 + 1) if by_rows else range(c1, c2 + 1)):
            line = (i, i, c1, c2) if by_rows else (r1, r2, i, i)
            sheet.Range(block_address([line])).Interior.Color = probe_color(code[i])
            if code[i] % 200 == 0:
                tick()


def probe_sheet(sheet: Any, tick: Any) -> tuple[list, list, list[list[float]]]:
    """在内存里给这张表涂探测底色、导出两次 PDF，读出每页的行位置和列位置（工作簿不保存）。
    返回 (每页的行位置, 每页的列位置, 每页大小)"""
    blocks = _probe_blocks(sheet)
    rows = sorted({r for r1, r2, _c1, _c2 in blocks for r in range(r1, r2 + 1)})
    cols = sorted({c for _r1, _r2, c1, c2 in blocks for c in range(c1, c2 + 1)})
    if len(rows) + len(cols) > MAX_PROBE_LINES:
        raise ValueError(f"打印范围太大（{len(rows)} 行、{len(cols)} 列），没法逐行定位；请设好打印区域后重新导出 PDF")
    if sheet.ProtectContents:
        try:
            sheet.Unprotect()  # 只在内存里：工作簿是只读打开的，不保存
        except Exception:
            raise ValueError("工作表有密码保护，没法定位框住的格子") from None
    sheet.Cells.FormatConditions.Delete()  # 条件格式的颜色会盖住探测底色
    try:
        if sheet.PageSetup.BlackAndWhite:
            sheet.PageSetup.BlackAndWhite = False  # 黑白打印时底色不导出
    except Exception:
        pass
    for block in blocks:
        sheet.Range(block_address([block])).Interior.Pattern = 1  # xlSolid：带花纹的格子也换成纯色
    result = []
    sizes: list[list[float]] = []
    with tempfile.TemporaryDirectory(prefix="pdf_review_") as tmp:
        for by_rows, numbers in ((True, rows), (False, cols)):
            _paint(sheet, blocks, numbers, by_rows, tick)
            tick()
            out = Path(tmp) / ("rows.pdf" if by_rows else "cols.pdf")
            wf._export_sheet_pdf(sheet, out)
            tick()
            pages = read_bands(out, "y" if by_rows else "x")
            result.append([{numbers[k - 1]: span for k, span in page.items() if k <= len(numbers)} for page in pages])
            if by_rows:
                with fitz.open(out) as doc:
                    sizes = [[round(p.rect.width, 2), round(p.rect.height, 2)] for p in doc]
    return result[0], result[1], sizes


def _find_sheet(book: Any, name: str) -> Any:
    """按表名找；按文件名猜出来的表名是去掉了非法字符的，再按同样规则比一次"""
    sheet = batch._find_sheet(book, name)
    if sheet is not None:
        return sheet
    for candidate in (name, re.sub(r"_\d+$", "", name)):
        for sheet in book.Worksheets:
            if sanitize_filename(str(sheet.Name)) == candidate:
                return sheet
    return None


def _same_pages(a: list[list[float]], b: list[list[float]]) -> bool:
    return len(a) == len(b) and all(abs(x[0] - y[0]) < 2 and abs(x[1] - y[1]) < 2 for x, y in zip(a, b))


# ---------- 后台算行列位置 ----------

def _version(pdf: str) -> int:
    try:
        return Path(pdf).stat().st_mtime_ns
    except OSError:
        return -1


def _fresh(pdf: str) -> dict | None:
    """算好的、PDF 没变过的结果。调用方持有 _COND"""
    found = _MAPS.get(pdf)
    return found if found and found["v"] == _version(pdf) else None


def prepare(pdfs: list[str], *, first: bool = False) -> None:
    """排进后台去算；first=True 插到最前面（用户正等着点开的那份）"""
    global _WORKER
    with _COND:
        for pdf in reversed(pdfs) if first else pdfs:
            if _fresh(pdf) or pdf in _WORKING:
                continue
            if pdf in _QUEUE:
                if not first:
                    continue
                _QUEUE.remove(pdf)
            _QUEUE.insert(0, pdf) if first else _QUEUE.append(pdf)
        if _QUEUE and _WORKER is None:
            _WORKER = threading.Thread(target=_work, name="PDF 复检定位", daemon=True)
            _WORKER.start()
        _COND.notify_all()


def _state(pdf: str) -> str:
    """调用方持有 _COND"""
    found = _fresh(pdf)
    if found:
        return found["state"]
    return "working" if pdf in _WORKING else "pending" if pdf in _QUEUE else "none"


def maps_status() -> dict[str, dict]:
    """每份 PDF 算到哪了：ready / error（带原因）/ working / pending / none"""
    with _COND:
        pdfs = set(_MAPS) | set(_QUEUE) | set(_WORKING)
        return {pdf: {"state": _state(pdf), "error": (_fresh(pdf) or {}).get("error")} for pdf in pdfs}


def _owner_of(pdf: str) -> dict | None:
    if not Path(pdf).is_file():
        return None
    with wf.STATE_LOCK:
        return _owner(Path(pdf))


def _take() -> list[tuple[str, dict]]:
    """取下一批：队首那份，加上同一份 Excel 的其他 PDF（打开一次一起算）。
    查对应的 Excel 要拿 STATE_LOCK，不在 _COND 里面拿，免得两把锁互相等。
    队列空了就在锁里把 _WORKER 清掉，之后再排进来的会另起一个线程，不会排了没人算"""
    global _WORKER
    while True:
        with _COND:
            if not _QUEUE:
                _WORKER = None
                return []
            pdf = _QUEUE.pop(0)
            rest = list(_QUEUE)
        owner = _owner_of(pdf)
        if not owner:
            continue
        peers = [(other, peer) for other in rest
                 if (peer := _owner_of(other)) and peer["wb_path"] == owner["wb_path"]]
        with _COND:
            if _fresh(pdf):
                continue
            group = [(pdf, owner)] + [(other, peer) for other, peer in peers if other in _QUEUE]
            for other, _ in group[1:]:
                _QUEUE.remove(other)
            _WORKING[:] = [item for item, _ in group]
            return group


def _work() -> None:
    global _WORKER
    try:
        _work_loop()
    except Exception as err:  # 不该走到这里；万一出错也要让下次排队能另起线程
        debug_log.exception(err, "PDF 复检定位线程")
        with _COND:
            _WORKER = None
            _WORKING.clear()


def _work_loop() -> None:
    while True:
        group = _take()
        if not group:
            return
        try:
            _probe_workbook(group)
        except Exception as err:
            debug_log.exception(err, "PDF 复检定位")
            with _COND:
                for pdf, _ in group:
                    if not _fresh(pdf):  # 同一份 Excel 里已经算好的几张表留着
                        _MAPS[pdf] = {"v": _version(pdf), "state": "error", "error": f"定位出错：{err}"}
        finally:
            with _COND:
                _WORKING.clear()
                _COND.notify_all()


def _probe_workbook(group: list[tuple[str, dict]]) -> None:
    """只读打开这份 Excel，逐张表涂探测色导出，记下每份 PDF 每页的行列位置（工作簿不保存）"""
    path = group[0][1]["wb_path"]
    watchdog = ExcelWatchdog(lambda: debug_log.event("⚠️ PDF 复检定位：后台 Excel 长时间没反应，已结束它"))
    excel = None
    wf.pythoncom.CoInitialize()
    try:
        watchdog.start()
        excel = start_background_excel()
        watchdog.register(excel)
        book = open_workbook_checked(excel, str(Path(path).resolve()), ReadOnly=True, UpdateLinks=0)
        print_layout.bind(book, path)
        try:
            for pdf, owner in group:
                found = _probe_pdf(book, pdf, owner["sheet"], watchdog.progress)
                with _COND:
                    _MAPS[pdf] = found
                    _COND.notify_all()
        finally:
            book.Close(SaveChanges=False)
    finally:
        try:
            if excel is not None:
                quit_background_excel(excel)
        finally:
            watchdog.stop()
            wf.pythoncom.CoUninitialize()


def _probe_pdf(book: Any, pdf: str, name: str, tick: Any) -> dict:
    version = _version(pdf)
    try:
        sheet = _find_sheet(book, name or "")
        if sheet is None:
            raise ValueError(f"文件里找不到工作表「{name}」（可能改名或删掉了），请重新导出这份 PDF")
        started = time.perf_counter()
        rows, cols, sizes = probe_sheet(sheet, tick)
        expected = _page_sizes(Path(pdf))
        debug_log.event("🔎 探测行列位置", PDF=Path(pdf).name, 表=sheet.Name, 页数=len(sizes), PDF页数=len(expected),
                        行=sum(len(p) for p in rows), 列=sum(len(p) for p in cols),
                        秒=round(time.perf_counter() - started, 1))
        if not _same_pages(sizes, expected):
            raise ValueError(f"Excel 现在打印出来是 {len(sizes)} 页，和这份 PDF（{len(expected)} 页）对不上，"
                             "可能导出 PDF 后改过表格；请重新导出 PDF")
        return {"v": version, "state": "ready", "sheet": str(sheet.Name), "rows": rows, "cols": cols}
    except Exception as err:
        debug_log.event("❌ PDF 复检定位失败", PDF=Path(pdf).name, 表=name, 原因=err)
        return {"v": version, "state": "error", "error": str(err)}


# ---------- 点一下：打开 Excel 跳到那一格 ----------

def _line_at(bands: dict[int, tuple[float, float]], value: float) -> int | None:
    """value 落在哪一行（列）里；落在几行的重叠处取最窄的。都不在（页边空白、格线缝里）时，1 个点以内的也算"""
    for slack in (0.0, 1.0):
        hits = [(b - a, n) for n, (a, b) in bands.items() if a - slack <= value <= b + slack]
        if hits:
            return min(hits)[1]
    return None


def cell_at(rows: dict[int, tuple[float, float]], cols: dict[int, tuple[float, float]], x: float, y: float) -> str | None:
    """PDF 上的点（PDF 点）→ 单元格地址，点在表格外时返回 None"""
    row, col = _line_at(rows, y), _line_at(cols, x)
    return f"{col_index_to_letters(col)}{row}" if row and col else None


def _old_files(old_root: Path) -> list[Path]:
    key = str(old_root).casefold()
    cached = _OLD_FILES.get(key)
    if cached and time.monotonic() - cached[0] < OLD_LIST_SECONDS:
        return cached[1]
    files = [p for p in old_root.rglob("*") if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS | {".xlsb"}
             and not p.name.startswith("~$")]
    _OLD_FILES[key] = (time.monotonic(), files)
    return files


def backup_for(wb_path: Path, old_folder: str) -> tuple[Path | None, str]:
    """在原始备份文件夹里找这份 Excel 的老版本：同样的相对路径，或者文件名在两边都只有一个。
    找不到返回 (None, 原因)"""
    old_root = Path(old_folder.strip().strip('"').strip("'")).expanduser().resolve()
    if not old_root.is_dir():
        return None, f"原始备份文件夹不存在：{old_root}"
    with wf.STATE_LOCK:
        root = Path(wf.STATE.get("pdf_review", {}).get("folder") or wb_path.parent)
    name = workbook_rel_name(wb_path, root)
    found = pair_files(old_root, _old_files(old_root), [name], root).get(name)
    if found is None:
        return None, "原始备份文件夹里找不到这份文件（或有几份同名的分不清）"
    if found.resolve() == wb_path.resolve():
        return None, "原始备份和当前文件是同一个文件"
    return found.resolve(), ""


def open_at(pdf: str, page: Any, x: Any, y: Any, old_folder: Any = "") -> dict:
    """PDF 上点的位置 → 用 Excel 打开那一格；有原始备份时备份也打开（只读）跳到同一格。
    还没算好行列位置时返回 pending，网页等算好再点一次"""
    path = Path(pdf).resolve()
    if not _in_folder(path) or path.suffix.lower() != ".pdf":
        raise ValueError("只能查看当前文件夹里的 PDF，请先重新加载")
    if isinstance(page, bool) or not isinstance(page, int) or any(
            isinstance(v, bool) or not isinstance(v, (int, float)) for v in (x, y)):
        raise ValueError("点的位置不对")
    with wf.STATE_LOCK:
        owner = _owner(path)
    if not owner:
        raise ValueError("找不到这份 PDF 对应的 Excel 文件")
    with _COND:
        found = _fresh(str(path))
    if not found:
        prepare([str(path)], first=True)
        return {"pending": True}
    sheet, cell = owner["sheet"], ""
    if found["state"] == "ready":
        sheet = found["sheet"]
        if 0 <= page < len(found["rows"]):
            cell = cell_at(found["rows"][page], found["cols"][page], float(x), float(y)) or ""
    wb_path = Path(owner["wb_path"])
    with wf.STATE_LOCK:
        job = wf.STATE.get("job") or {}
        if wf.STATE.get("busy") and (job.get("kind", "") in wf.QUICK_TASK_NAMES or job.get("kind", "").startswith("batch_")):
            reason = batch.open_block_reason(job, wb_path)
            if reason:
                raise ValueError(reason)
    notes = [] if found["state"] == "ready" else [f"没能定位到格子（{found['error']}），只打开了工作表"]
    if found["state"] == "ready" and not cell:
        notes.append("点的地方不在表格里，只打开了工作表")
    backup, why = backup_for(wb_path, old_folder) if isinstance(old_folder, str) and old_folder.strip() else (None, "")
    if why:
        notes.append(why)
    if backup:
        result = repair_compare.open_comparison({"current": wb_path.resolve(strict=True), "backup": backup,
                                                 "sheet": sheet, "old_sheet": sheet, "cell": cell})
        if result.get("note"):
            notes.append(result["note"])
    else:
        note = open_workbook_in_excel(wb_path, sheet, cell)
        if note:
            notes.append(note)
    debug_log.event("🔎 PDF 复检打开", PDF=path.name, 页=page + 1, 点=[round(x, 1), round(y, 1)], 表=sheet,
                    格子=cell or "-", 备份=backup.name if backup else "-", 说明="；".join(notes) or "-")
    return {"sheet": sheet, "cell": cell, "backup": bool(backup), "note": "；".join(notes)}


# ---------- 三色标记：一键处理 ----------
# 网页上每份 PDF（= 一张表）可以标成：蓝（默认，不用管）/ 黄（要人工改）/ 红（这张表不要了）。
# 网页上黄红分开两个按钮（2026-11-02）：「删除标红」只传 red，「打开标黄」只传 yellow；两个都传时先删后开。
# 删除：在后台 Excel 里把标红的表删掉并保存（同一份 Excel 的几张一起删、只开一次），
# 删掉的表对应的 PDF 也删掉；再把标黄的用 Excel 打开（每份 Excel 开一次，跳到标黄时右键的那一格），
# 一次最多开 OPEN_LIMIT 份 Excel（按文件算，同一份 Excel 几张标黄的表只算 1 份），剩下的再点一次接着开。

OPEN_LIMIT = 10
_PROCESS: dict[str, Any] = {"running": False}
_PROCESS_LOCK = threading.Lock()


def process_status() -> dict:
    with _PROCESS_LOCK:
        return {**_PROCESS, "result": _PROCESS.get("result")}


def _process_item(raw: Any, color: str) -> dict:
    if not isinstance(raw, dict) or not isinstance(raw.get("pdf"), str):
        raise ValueError("标记的数据不对，请刷新页面再试")
    path = Path(raw["pdf"]).resolve()
    if not _in_folder(path) or path.suffix.lower() != ".pdf":
        raise ValueError(f"「{path.name}」不在当前文件夹里或已经不在了，请重新加载 PDF")
    with wf.STATE_LOCK:
        owner = _owner(path)
    if not owner:
        raise ValueError(f"「{path.name}」找不到对应的 Excel 文件")
    item = {"pdf": str(path), "name": path.name, "color": color, "wb_path": owner["wb_path"], "sheet": owner["sheet"],
            "spots": []}
    # 标黄的页：每页一处，右键标的带点的位置（x, y），按标题按钮标的只有页码
    for spot in raw.get("spots") or []:
        page = spot.get("page") if isinstance(spot, dict) else None
        if isinstance(page, bool) or not isinstance(page, int):
            continue
        x, y = spot.get("x"), spot.get("y")
        has_point = all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in (x, y))
        item["spots"].append({"page": page, **({"x": float(x), "y": float(y)} if has_point else {})})
    item["spots"].sort(key=lambda spot: spot["page"])
    return item


def process(red: Any, yellow: Any) -> None:
    """一键处理：后台删掉标红的表，再打开标黄的。网页轮询 process_status() 看进度和结果"""
    reds = [_process_item(raw, "red") for raw in (red or [])]
    yellows = [_process_item(raw, "yellow") for raw in (yellow or [])]
    if not reds and not yellows:
        raise ValueError("没有标黄或标红的表")
    from core import bl_links, hidden_repair, link_repair, mark_finder
    with wf.STATE_LOCK, _PROCESS_LOCK:
        if (wf.STATE.get("busy") or _PROCESS.get("running") or link_repair.status().get("running")
                or mark_finder.status().get("running") or hidden_repair.status().get("running")
                or bl_links.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束再处理")
        if reds:
            wf.STATE["busy"] = True  # 要改文件：别的任务这时候不能开
            wf.STATE["maintenance"] = "pdf_review_process"
        books = list(dict.fromkeys(item["wb_path"] for item in reds))
        _PROCESS.clear()
        _PROCESS.update(running=True, stage="delete" if reds else "open", total=len(books), done=0, current="",
                        result=None)
    root = wf.STATE.get("pdf_review", {}).get("folder") or ""

    def run() -> None:
        result: dict[str, list] = {"deleted": [], "failed": [], "opened": [], "open_failed": [], "later": []}
        try:
            with debug_log.side_task("pdf_review_process", {"folder": root, "red": len(reds), "yellow": len(yellows)}):
                for item in reds + yellows:
                    debug_log.event("🎨 PDF 复检标记", 颜色="红" if item["color"] == "red" else "黄",
                                    PDF=item["name"], 表=item["sheet"],
                                    页=[spot["page"] + 1 for spot in item["spots"]] or "-")
                if reds:
                    _delete_red(reds, root, result)
                if yellows:
                    with _PROCESS_LOCK:
                        _PROCESS.update(stage="open", current="")
                    with wf.STATE_LOCK:
                        wf.STATE["busy"] = False  # 删完了；打开 Excel 不用占着任务槽
                        wf.STATE.pop("maintenance", None)
                    _open_yellow(yellows, root, result)
                debug_log.event("✅ PDF 复检一键处理完成", 删掉=len(result["deleted"]), 没删掉=len(result["failed"]),
                                打开=len(result["opened"]), 没打开=len(result["open_failed"]), 下次开=len(result["later"]))
        except Exception as err:
            result["failed"].append({"name": "任务", "sheet": "", "reason": f"处理出错：{err}"})
        finally:
            with wf.STATE_LOCK:
                if wf.STATE.get("maintenance") == "pdf_review_process":
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)
            with _PROCESS_LOCK:
                _PROCESS.update(running=False, current="", result=result)

    try:
        threading.Thread(target=run, name="PDF 复检一键处理", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _PROCESS_LOCK:
            if wf.STATE.get("maintenance") == "pdf_review_process":
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)
            _PROCESS["running"] = False
        raise


def _wait_probe_idle(wb_path: str) -> list[str]:
    """后台定位线程正在只读打开这份 Excel 时等它算完（最多 2 分钟）；排队里同一份 Excel 的先拿掉，
    返回拿掉的，删完表再排回去。查对应的 Excel 要拿 STATE_LOCK，不在 _COND 里面拿"""
    target = Path(wb_path).resolve()

    def same(pdf: str) -> bool:
        owner = _owner_of(pdf)
        return bool(owner) and Path(owner["wb_path"]).resolve() == target

    with _COND:
        queued, working = list(_QUEUE), list(_WORKING)
    dropped = [pdf for pdf in queued if same(pdf)]
    busy = any(same(pdf) for pdf in working)
    deadline = time.monotonic() + 120
    with _COND:
        for pdf in dropped:
            if pdf in _QUEUE:
                _QUEUE.remove(pdf)
        while busy and any(pdf in _WORKING for pdf in working) and time.monotonic() < deadline:
            _COND.wait(1)
    return dropped


def _delete_red(reds: list[dict], root: str, result: dict) -> None:
    groups: dict[str, list[dict]] = {}
    for item in reds:
        groups.setdefault(item["wb_path"], []).append(item)
    guard = file_problem.Guard("PDF 复检：删除标红的表")
    excel = None
    wf.pythoncom.CoInitialize()
    try:
        for index, (wb_path, items) in enumerate(groups.items()):
            label = workbook_rel_name(Path(wb_path), Path(root)) if root else Path(wb_path).name
            with _PROCESS_LOCK:
                _PROCESS.update(done=index, current=label)
            dropped = _wait_probe_idle(wb_path)
            while True:
                guard.progress()
                try:
                    if excel is None:
                        excel = guard.watch(start_background_excel())
                    outcome = _delete_in_workbook(excel, Path(wb_path), items, label)
                    break
                except Exception as err:
                    debug_log.exception(err, f"PDF 复检删表 {label}")
                    retry, message = guard.problem(err, wb_path, label, root,
                                                   close=lambda: file_problem.quit_quietly(excel))
                    if guard.closed:
                        excel = None
                    if retry:
                        wf.log_msg(f"🔁 重试 {label}")
                        continue
                    outcome = [{**item, "ok": False, "reason": message} for item in items]
                    break
            for item in outcome:
                entry = {"pdf": item["pdf"], "name": label, "sheet": item["sheet"], "reason": item.get("reason", "")}
                if item["ok"]:
                    entry["pdf_removed"] = item.get("pdf_removed", False)
                result["deleted" if item["ok"] else "failed"].append(entry)
            requeue = [pdf for pdf in dropped if Path(pdf).is_file()]
            if requeue:
                prepare(requeue)  # 同一份 Excel 没删的表接着算行列位置
            with _PROCESS_LOCK:
                _PROCESS.update(done=index + 1)
    finally:
        try:
            file_problem.quit_quietly(excel)
        finally:
            guard.stop()
            wf.pythoncom.CoUninitialize()


def _delete_in_workbook(excel: Any, wb_path: Path, items: list[dict], label: str) -> list[dict]:
    """在后台 Excel 里删掉这份文件里标红的表并保存；删表前把别的表引用它的公式转成值（和精简 Sheet 一样）。
    返回每一项：ok、reason；删成功的 PDF 也删掉"""
    if not wb_path.is_file():
        return [{**item, "ok": False, "reason": "Excel 文件已经不在了"} for item in items]
    if workbook_locked(wb_path):
        raise ValueError("文件正在 Excel 中打开，请先在 Excel 里关掉它（改过的先保存），再点「重试」")
    before_hash = sheet_history.fingerprint(wb_path)
    book = open_workbook_checked(excel, str(wb_path.resolve()), UpdateLinks=0)
    outcome: list[dict] = []
    deleting: list[str] = []
    try:
        before = [str(sheet.Name) for sheet in book.Sheets]
        for item in items:
            sheet = _find_sheet(book, item["sheet"])
            if sheet is None:
                outcome.append({**item, "ok": False, "reason": "文件里已经没有这张表（可能改名或删掉了）"})
            elif str(sheet.Name) in deleting:
                outcome.append({**item, "ok": True, "same": True})
            else:
                deleting.append(str(sheet.Name))
                outcome.append({**item, "ok": True, "real": str(sheet.Name)})
        visible = [str(sheet.Name) for sheet in book.Sheets if int(sheet.Visible) == -1]
        if deleting and visible and all(name in deleting for name in visible):
            reason = "这是文件里剩下的唯一一张显示着的表，Excel 不允许删掉（整个文件不要的话请直接删文件）"
            debug_log.event("⚠️ PDF 复检删表跳过", 文件=label, 表=deleting, 原因=reason)
            return [{**item, "ok": False, "reason": reason} if item["ok"] else item for item in outcome]
        if not deleting:
            return outcome
        kept_values = delete_sheets_preserving_values(book, deleting)
        book.Save()
        after = [str(sheet.Name) for sheet in book.Sheets]
    finally:
        book.Close(SaveChanges=False)
    with wf.STATE_LOCK:
        sheet_history.record(wb_path, before, after, deleting, {}, kept_values, before_hash)
        wf._persist_state()
    debug_log.event("🗑️ PDF 复检删表", 文件=label, 删掉=deleting, 保留引用值=len(kept_values))
    wf.log_msg(f"🗑️ {label}：删掉了 {'、'.join(deleting)}" + (f"（另外保留引用值 {len(kept_values)} 处）" if kept_values else ""))
    removed: list[Path] = []
    for item in outcome:
        if not item["ok"]:
            continue
        pdf = Path(item["pdf"])
        try:
            pdf.unlink(missing_ok=True)
            removed.append(pdf)
            item["pdf_removed"] = True
        except OSError as err:
            item["reason"] = f"表删掉了，但 PDF 删不掉（可能正在用别的程序打开）：{err}"
            debug_log.event("⚠️ PDF 复检删 PDF 失败", PDF=pdf.name, 原因=err)
    if removed:
        forget_pdfs(removed)
    if kept_values:
        for item in outcome:
            if item["ok"] and not item.get("reason"):
                item["reason"] = f"别的表引用它的 {len(kept_values)} 处公式已转成值"
    return outcome


def _spot_cells(item: dict) -> tuple[str, list[dict]]:
    """标黄的每一页 → 那一格（行列位置算好了、右键时点在表格里才有），返回 (工作表名, [{page, cell}])"""
    with _COND:
        found = _fresh(item["pdf"])
    ready = bool(found) and found["state"] == "ready"
    sheet = found["sheet"] if ready else item["sheet"]
    spots = []
    for spot in item["spots"] or [{"page": None}]:
        cell = ""
        page = spot.get("page")
        if ready and "x" in spot and page is not None and 0 <= page < len(found["rows"]):
            cell = cell_at(found["rows"][page], found["cols"][page], spot["x"], spot["y"]) or ""
        spots.append({"page": page, "cell": cell})
    return sheet, spots


def _open_yellow(yellows: list[dict], root: str, result: dict) -> None:
    """每份 Excel 打开一次，跳到这份文件里第一处标黄的格子（右键标黄时记下的；没有就只打开表）。
    每一处都写进结果，网页上点一下跳过去"""
    groups: dict[str, list[dict]] = {}
    for item in yellows:
        groups.setdefault(item["wb_path"], []).append(item)
    for index, (wb_path, items) in enumerate(groups.items()):
        label = workbook_rel_name(Path(wb_path), Path(root)) if root else Path(wb_path).name
        entries = []
        for item in items:
            sheet, spots = _spot_cells(item)
            entries += [{"pdf": item["pdf"], "name": label, "sheet": sheet, **spot} for spot in spots]
        if index >= OPEN_LIMIT:
            result["later"].extend(entries)
            continue
        with _PROCESS_LOCK:
            _PROCESS.update(current=label)
        first = next((entry for entry in entries if entry["cell"]), entries[0])
        try:
            note = open_workbook_in_excel(Path(wb_path), first["sheet"], first["cell"])
        except Exception as err:
            debug_log.exception(err, f"PDF 复检打开标黄 {label}")
            for entry in entries:
                result["open_failed"].append({**entry, "reason": str(err)})
            continue
        debug_log.event("📂 PDF 复检打开标黄", 文件=label, 表=first["sheet"], 格子=first["cell"] or "-",
                        标黄处=[f"{e['sheet']}!{e['cell'] or '-'}(第{(e['page'] or 0) + 1}页)" for e in entries],
                        说明=note or "-")
        for entry in entries:
            entry["reason"] = note if entry is first else ""
            entry["first"] = entry is first
            result["opened"].append(entry)


# ---------- 单份 Excel 重新生成 PDF ----------
# 复检时在 Excel 里改了内容、保存后，想马上看改完的效果：只把这一份 Excel 的 PDF 重新导出一次（它的每张表），
# 不用整个文件夹重导。Excel 和上次导出时一样（哈希没变，多半是还没保存）就先问一句，确认了照样重导。

_REGEN: dict[str, Any] = {"running": False}
_REGEN_LOCK = threading.Lock()


def regen_status() -> dict:
    with _REGEN_LOCK:
        return dict(_REGEN)


def regenerate(pdf: Any, force: Any = False) -> dict:
    """返回 {"started": True} 后台开始；{"unchanged": True} Excel 没改过（force=True 时不问）。网页轮询 regen_status()"""
    if not isinstance(pdf, str) or not pdf:
        raise ValueError("要重新生成哪一份？请刷新页面再试")
    path = Path(pdf).resolve()
    if not _in_folder(path) or path.suffix.lower() != ".pdf":
        raise ValueError(f"「{path.name}」不在当前文件夹里或已经不在了，请重新加载 PDF")
    with wf.STATE_LOCK:
        owner = _owner(path)
        review_root = wf.STATE.get("pdf_review", {}).get("folder") or ""
    if not owner or not Path(owner["wb_path"]).is_file():
        raise ValueError(f"「{path.name}」找不到对应的 Excel 文件")
    wb_path = Path(owner["wb_path"])
    # PDF 在哪个 PDFS 里就导回哪个；旧的混放 PDF 导到现在这个文件夹的 PDFS 里
    root = pdf_store.export_root_of(path) or Path(review_root)
    label = workbook_rel_name(wb_path, root)
    digest = sheet_history.fingerprint(wb_path)
    same = pdf_store.unchanged(root, wb_path, digest)
    debug_log.event("🔄 PDF 复检重新生成", 文件=label, PDF=path.name, 哈希=digest[:12], 没变=same, 确认=bool(force))
    if same and force is not True:
        return {"started": False, "unchanged": True, "name": label}
    from core import bl_links, hidden_repair, link_repair, mark_finder
    with wf.STATE_LOCK, _PROCESS_LOCK, _REGEN_LOCK:
        if (wf.STATE.get("busy") or _PROCESS.get("running") or _REGEN.get("running")
                or link_repair.status().get("running") or mark_finder.status().get("running")
                or hidden_repair.status().get("running") or bl_links.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束再重新生成")
        wf.STATE["busy"] = True  # 要写 PDFS：这时候别开导出 PDF 等任务
        wf.STATE["maintenance"] = "pdf_review_regen"
        _REGEN.clear()
        _REGEN.update(running=True, name=label, pdf=str(path), wb_path=str(wb_path), result=None)

    def run() -> None:
        result: dict[str, Any] = {"ok": False, "name": label, "wb_path": str(wb_path), "pdfs": [], "reason": ""}
        try:
            with debug_log.side_task("pdf_review_regen", {"folder": str(root), "file": label}):
                result.update(_regenerate_workbook(wb_path, root, label))
                debug_log.event("✅ PDF 复检重新生成完成" if result["ok"] else "⚠️ PDF 复检重新生成没成功",
                                文件=label, PDF=len(result["pdfs"]), 原因=result["reason"] or "-")
        except Exception as err:
            debug_log.exception(err, f"PDF 复检重新生成 {label}")
            result["reason"] = f"重新生成出错：{err}"
        finally:
            with wf.STATE_LOCK:
                if wf.STATE.get("maintenance") == "pdf_review_regen":
                    wf.STATE["busy"] = False
                    wf.STATE.pop("maintenance", None)
            with _REGEN_LOCK:
                _REGEN.update(running=False, result=result)

    try:
        threading.Thread(target=run, name="PDF 复检重新生成", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _REGEN_LOCK:
            if wf.STATE.get("maintenance") == "pdf_review_regen":
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)
            _REGEN["running"] = False
        raise
    return {"started": True, "name": label}


def _regenerate_workbook(wb_path: Path, root: Path, label: str) -> dict:
    """后台 Excel 只读打开这一份，把它要导出的表都重新导出到 PDFS（和主页导出 PDF 同一套规则），记下新的哈希"""
    from core import reporter
    _wait_probe_idle(str(wb_path))  # 定位线程正开着这份 Excel 算旧 PDF 的位置：等它算完，旧 PDF 要删掉了
    guard = file_problem.Guard("PDF 复检：重新生成 PDF")
    excel = None
    reporter.suspend_updates(True)  # 不是主页的批量任务：不写进执行记录
    wf.pythoncom.CoInitialize()

    def get_excel() -> Any:
        nonlocal excel
        if excel is None:
            excel = guard.watch(start_background_excel())
        return excel

    try:
        while True:
            guard.progress()
            try:
                paths = wf._export_workbook_pdfs(get_excel, wb_path, root, set(), force=True)
                return {"ok": True, "pdfs": [str(p) for p in paths], "reason": ""}
            except Exception as err:
                debug_log.exception(err, f"PDF 复检重新生成 {label}")
                retry, message = guard.problem(err, wb_path, label, root,
                                               close=lambda: file_problem.quit_quietly(excel))
                if guard.closed:
                    excel = None
                if retry:
                    wf.log_msg(f"🔁 重试 {label}")
                    continue
                return {"ok": False, "pdfs": [], "reason": message}
    finally:
        reporter.suspend_updates(False)
        try:
            file_problem.quit_quietly(excel)
        finally:
            guard.stop()
            wf.pythoncom.CoUninitialize()
