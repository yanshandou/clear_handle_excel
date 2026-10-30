"""PDF 人工复检：导出的 PDF 整页显示在网页上，人工拖框圈出有问题的地方，框住的单元格在原 Excel 里标色。

框怎么对回单元格：PDF 只是画出来的图，不知道哪块是哪一格。写进 Excel 时只读打开工作簿，在内存里
把打印范围的每一行涂成各不相同的底色导出一次 PDF，再把每一列涂成不同底色导出一次（都不保存）。
读这两份 PDF 里的底色矩形，就知道每页上每一行、每一列占的位置；框盖住哪几行、哪几列就是哪些格子。
导出用的打印方向和正式导出一样（print_layout），所以页码和位置对得上；页数或页面大小对不上时
说明 Excel 改过、PDF 已经旧了，不硬套。打印标题行、隐藏行列、横向分页都按 Excel 实际打印的样子来。

标色和工作台共用一套：每个框一条条件格式规则（CHE_MARK_…），记在 batch 的标记记录里（source=pdf_review），
有标记的工作表标签标橙黄色；工作台「清除 Excel 标记颜色」也会一起清掉。
"""
from __future__ import annotations

import copy
import re
import tempfile
import threading
import time
from collections import Counter
from pathlib import Path
from typing import Any

from core import batch, debug_log, print_layout, workflow as wf
from core.common import EXCEL_EXTENSIONS, col_index_to_letters, sanitize_filename, strip_review_prefix, workbook_rel_name
from core.excel_ops import normalize_workbook_view, open_workbook_checked, quit_background_excel, start_background_excel
from core.excel_watchdog import ExcelWatchdog
from core.pdf_detector import fitz

SOURCE = "pdf_review"
PAGE_ZOOM = 2.5  # 网页上显示的页面图：A4 竖向约 1490 像素宽，铺满屏幕也看得清字
MAX_PROBE_LINES = 20000  # 打印范围超过这么多行 + 列就不涂色定位了（逐行涂色太慢）
MAX_RECTS_PER_PDF = 500
RECT_ID = re.compile(r"^[0-9a-f]{8,32}$")
_LOCK = threading.RLock()
_STATUS: dict[str, Any] = {"running": False}
_PAGES_CACHE: dict[tuple[str, int, int], list[list[float]]] = {}


def _store() -> dict[str, Any]:
    """rects: 框 id -> 框；drop: 框删掉了、Excel 里的标记还没去掉的标记 id；trash: 删掉的框（撤销时找回定位结果）"""
    store = wf.STATE.setdefault("pdf_review", {})
    store.setdefault("rects", {})
    store.setdefault("drop", [])
    store.setdefault("trash", {})
    return store


def _exports() -> dict[str, dict]:
    """导出 PDF 时记下的（workflow._export_workbook_pdfs）：PDF 路径 -> {wb_path, sheet}"""
    return wf.STATE.setdefault("pdf_exports", {})


def mark_id(rect_id: str) -> str:
    """框在 batch 标记记录里的 id。标识只能是字母数字（batch.MARK_TAG），不能带下划线"""
    return f"pdf{rect_id}"


# ---------- 列出 PDF ----------

def _resolve_folder(folder: str) -> Path:
    root = Path(folder.strip().strip('"').strip("'")).expanduser().resolve()
    if not root.is_dir():
        raise ValueError(f"文件夹不存在: {root}")
    return root


def _guess_owner(pdf: Path) -> dict | None:
    """没有导出记录（旧版本导出的、扫描时导出的）：按「Excel 文件名_表名.pdf」在同一个文件夹里找 Excel"""
    best: tuple[int, Path, str] | None = None
    for wb in pdf.parent.iterdir():
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
    return _guess_owner(pdf)


def _page_sizes(pdf: Path) -> list[list[float]]:
    stat = pdf.stat()
    key = (str(pdf), stat.st_mtime_ns, stat.st_size)
    if key not in _PAGES_CACHE:
        with fitz.open(pdf) as doc:
            _PAGES_CACHE[key] = [[round(page.rect.width, 2), round(page.rect.height, 2)] for page in doc]
        if len(_PAGES_CACHE) > 2000:
            _PAGES_CACHE.pop(next(iter(_PAGES_CACHE)))
    return _PAGES_CACHE[key]


def list_pdfs(folder: str) -> dict:
    """文件夹（含子文件夹）里的 PDF，每份带对应的 Excel 和表、各页大小（PDF 点），以及画过的框"""
    root = _resolve_folder(folder)
    pdfs = sorted((p for p in root.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf"
                   and not p.name.startswith(("~", "."))), key=lambda p: workbook_rel_name(p, root))
    entries = []
    with wf.STATE_LOCK:
        _store()["folder"] = str(root)
        for gone in [p for p in _exports() if Path(p).parent == root or root in Path(p).parents]:
            if not Path(gone).is_file():
                _exports().pop(gone)  # PDF 已经删掉的记录不留着
        owners = {str(p): _owner(p) for p in pdfs}
    for pdf in pdfs:
        entry: dict[str, Any] = {"pdf": str(pdf), "name": workbook_rel_name(pdf, root), "v": pdf.stat().st_mtime_ns}
        try:
            entry["pages"] = _page_sizes(pdf)
        except Exception as err:
            entry.update(pages=[], error=f"PDF 打不开：{err}")
        owner = owners[str(pdf)]
        if owner:
            entry.update(wb_path=owner["wb_path"], wb_name=workbook_rel_name(Path(owner["wb_path"]), root),
                         sheet=owner["sheet"], guessed=bool(owner.get("guessed")))
        entries.append(entry)
    debug_log.event("🖍️ PDF 复检列出", 文件夹=root, PDF=len(entries), 找到Excel=sum(1 for e in entries if e.get("wb_path")))
    return {"folder": str(root), "pdfs": entries, **rects_snapshot()}


def page_png(pdf: str, page: int) -> bytes:
    path = Path(pdf).resolve()
    with wf.STATE_LOCK:
        root = _store().get("folder")
    if not root or path.suffix.lower() != ".pdf" or not path.is_file() or Path(root) not in path.parents:
        raise ValueError("只能查看当前文件夹里的 PDF")
    with fitz.open(path) as doc:
        if not 0 <= page < len(doc):
            raise ValueError("页码不对")
        return doc[page].get_pixmap(matrix=fitz.Matrix(PAGE_ZOOM, PAGE_ZOOM), alpha=False).tobytes("png")


# ---------- 框：保存、状态 ----------

def _rect_state(rect: dict, marks: dict) -> str:
    if mark_id(rect["id"]) in marks:
        return "marked"
    return "error" if rect.get("error") else "pending"


def rects_snapshot() -> dict:
    """各 PDF 的框（带状态：marked 已标进 Excel / pending 还没写 / error 没标成），还有待去掉的标记数"""
    with wf.STATE_LOCK:
        store, marks = _store(), batch._marks()
        by_pdf: dict[str, list] = {}
        for rect in store["rects"].values():
            item = {k: rect.get(k) for k in ("id", "page", "x0", "y0", "x1", "y1", "addr", "sheet", "error")}
            item["state"] = _rect_state(rect, marks)
            by_pdf.setdefault(rect["pdf"], []).append(item)
        drop = [cid for cid in store["drop"] if cid in marks]
        return {"rects": by_pdf, "drop": len(drop), "job": status()}


def _clean_rect(raw: Any, pages: list[list[float]]) -> dict:
    if not isinstance(raw, dict) or not isinstance(raw.get("id"), str) or not RECT_ID.match(raw["id"]):
        raise ValueError("标记格式不对")
    page = raw.get("page")
    if isinstance(page, bool) or not isinstance(page, int) or not 0 <= page < len(pages):
        raise ValueError("标记的页码不对")
    width, height = pages[page]
    values = []
    for key, limit in (("x0", width), ("y0", height), ("x1", width), ("y1", height)):
        value = raw.get(key)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError("标记的位置不对")
        values.append(min(max(float(value), 0.0), limit))
    x0, y0, x1, y1 = values
    if x1 - x0 < 1 or y1 - y0 < 1:
        raise ValueError("标记太小")
    return {"id": raw["id"], "page": page, "x0": round(x0, 2), "y0": round(y0, 2), "x1": round(x1, 2), "y1": round(y1, 2)}


def save(pdf: str, raw_rects: Any) -> dict:
    """网页上每画一个框、删一个框、撤销一次都把这份 PDF 的全部框发过来：新的记下，没了的删掉
    （已经标进 Excel 的，记进 drop，下次写进 Excel 时去掉标记）；撤销删除时框带着原来的定位结果回来"""
    path = Path(pdf).resolve()
    if not isinstance(raw_rects, list) or len(raw_rects) > MAX_RECTS_PER_PDF:
        raise ValueError(f"每份 PDF 最多 {MAX_RECTS_PER_PDF} 个标记")
    with wf.STATE_LOCK:
        root = _store().get("folder")
    if not root or path.suffix.lower() != ".pdf" or not path.is_file() or Path(root) not in path.parents:
        raise ValueError("只能标记当前文件夹里的 PDF，请先重新加载")
    pages = _page_sizes(path)
    cleaned = [_clean_rect(raw, pages) for raw in raw_rects]
    with wf.STATE_LOCK:
        owner = _owner(path)
        store, marks = _store(), batch._marks()
        ids = {r["id"] for r in cleaned}
        old = {rid: r for rid, r in store["rects"].items() if r["pdf"] == str(path)}
        added = removed = 0
        for rid, rect in old.items():
            if rid in ids:
                continue
            store["rects"].pop(rid)
            store["trash"][rid] = rect
            removed += 1
            if mark_id(rid) in marks and mark_id(rid) not in store["drop"]:
                store["drop"].append(mark_id(rid))
        for rect in cleaned:
            if rect["id"] in old:
                continue
            back = store["trash"].pop(rect["id"], None)
            if back is not None and back["pdf"] == str(path):
                rect = back  # 撤销删除：定位结果和标记都还在
                if mark_id(rect["id"]) in store["drop"]:
                    store["drop"].remove(mark_id(rect["id"]))
            else:
                rect.update(pdf=str(path), wb_path=(owner or {}).get("wb_path"), sheet=(owner or {}).get("sheet"),
                            addr=None, error=None if owner else "找不到这份 PDF 对应的 Excel 文件")
            store["rects"][rect["id"]] = rect
            added += 1
        while len(store["trash"]) > 2000:
            store["trash"].pop(next(iter(store["trash"])))
        wf._persist_state()
    debug_log.event("🖍️ PDF 标记保存", PDF=path.name, 框数=len(cleaned), 新增=added, 删除=removed)
    return rects_snapshot()


# ---------- 框 → 单元格 ----------

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


def pick(bands: dict[int, tuple[float, float]], lo: float, hi: float) -> list[int]:
    """框的范围 [lo, hi] 盖住了哪些行 / 列：盖住这一格三成以上，或框整个落在这一格里（在大格子里画小框）。
    都不够时取重叠最多的那一个"""
    hits, best, best_overlap = [], None, 0.0
    for index, (a, b) in bands.items():
        overlap = min(b, hi) - max(a, lo)
        if overlap <= 0:
            continue
        if overlap >= 0.3 * (b - a) or overlap >= 0.8 * (hi - lo):
            hits.append(index)
        if overlap > best_overlap:
            best, best_overlap = index, overlap
    if not hits and best is not None:
        hits = [best]
    return sorted(hits)


def runs(numbers: list[int]) -> list[tuple[int, int]]:
    """[3, 4, 5, 9, 10] → [(3, 5), (9, 10)]：中间隔着没打印的（隐藏的）行列时分成几段"""
    result: list[tuple[int, int]] = []
    for n in numbers:
        if result and n == result[-1][1] + 1:
            result[-1] = (result[-1][0], n)
        else:
            result.append((n, n))
    return result


def blocks_for(rect: dict, rows: dict[int, tuple[float, float]], cols: dict[int, tuple[float, float]]) -> list[tuple[int, int, int, int]]:
    """框 → 单元格块 [(首行, 末行, 首列, 末列)]；框里没有格子时返回空"""
    row_hits = pick(rows, rect["y0"], rect["y1"])
    col_hits = pick(cols, rect["x0"], rect["x1"])
    if not row_hits or not col_hits:
        return []
    return [(r1, r2, c1, c2) for r1, r2 in runs(row_hits) for c1, c2 in runs(col_hits)]


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


def _with_merged(sheet: Any, block: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    """框边上压到合并单元格的，扩到整个合并区域：合并格只看左上角那一格的格式，不扩的话标不出颜色"""
    r1, r2, c1, c2 = block
    if sheet.Range(block_address([block])).MergeCells is False:
        return block
    for _ in range(5):
        before = (r1, r2, c1, c2)
        edge = {(r, c) for r in range(r1, r2 + 1) for c in (c1, c2)} | {(r, c) for c in range(c1, c2 + 1) for r in (r1, r2)}
        for r, c in sorted(edge)[:2000]:
            cell = sheet.Cells(r, c)
            if cell.MergeCells:
                area = cell.MergeArea
                r1, c1 = min(r1, int(area.Row)), min(c1, int(area.Column))
                r2 = max(r2, int(area.Row) + int(area.Rows.Count) - 1)
                c2 = max(c2, int(area.Column) + int(area.Columns.Count) - 1)
        if (r1, r2, c1, c2) == before:
            break
    return r1, r2, c1, c2


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


# ---------- 写进 Excel ----------

def _plan() -> dict[str, dict]:
    """要处理的文件：wb_path -> {"add": [框], "drop": [标记 id]}。调用方持有 STATE_LOCK"""
    store, marks = _store(), batch._marks()
    plan: dict[str, dict] = {}
    for rect in store["rects"].values():
        if mark_id(rect["id"]) in marks or not rect.get("wb_path"):
            continue
        todo = dict(copy.deepcopy(rect), addr=None, error=None)  # 每次重新定位：Excel 可能改过
        plan.setdefault(rect["wb_path"], {"add": [], "drop": []})["add"].append(todo)
    store["drop"] = [cid for cid in store["drop"] if cid in marks]
    for cid in store["drop"]:
        plan.setdefault(marks[cid]["path"], {"add": [], "drop": []})["drop"].append(cid)
    return plan


def _locate(excel: Any, path: str, rects: list[dict], tick: Any) -> None:
    """只读打开，按表涂探测色定位，框里写上 sheet（实际表名）、addr 或 error"""
    book = open_workbook_checked(excel, str(Path(path).resolve()), ReadOnly=True, UpdateLinks=0)
    print_layout.bind(book, path)
    try:
        by_pdf: dict[str, list[dict]] = {}
        for rect in rects:
            by_pdf.setdefault(rect["pdf"], []).append(rect)
        for pdf, group in by_pdf.items():
            name = group[0]["sheet"] or ""
            try:
                sheet = _find_sheet(book, name)
                if sheet is None:
                    raise ValueError(f"文件里找不到工作表「{name}」（可能改名或删掉了）")
                started = time.perf_counter()
                rows, cols, sizes = probe_sheet(sheet, tick)
                expected = _page_sizes(Path(pdf))
                debug_log.event("🖍️ 探测行列位置", 表=sheet.Name, 页数=len(sizes), PDF页数=len(expected),
                                行=sum(len(p) for p in rows), 列=sum(len(p) for p in cols),
                                秒=round(time.perf_counter() - started, 1))
                if not _same_pages(sizes, expected):
                    raise ValueError(f"Excel 现在打印出来是 {len(sizes)} 页，和这份 PDF（{len(expected)} 页）对不上，"
                                     "可能导出 PDF 后改过表格；请重新导出 PDF 再标记")
            except Exception as err:
                for rect in group:
                    rect["error"] = str(err)
                debug_log.event("❌ PDF 标记定位失败", PDF=Path(pdf).name, 表=name, 原因=err)
                debug_log.exception(err, "PDF 标记定位")
                continue
            for rect in group:
                rect["sheet"] = str(sheet.Name)
                blocks = blocks_for(rect, rows[rect["page"]], cols[rect["page"]])
                if not blocks:
                    rect["error"] = "框里没有表格的格子（框在页边空白处？）"
                else:
                    rect["addr"] = block_address([_with_merged(sheet, b) for b in blocks])
                    rect["error"] = None
                debug_log.event("🖍️ 框→单元格", PDF=Path(pdf).name, 页=rect["page"] + 1,
                                框=[rect["x0"], rect["y0"], rect["x1"], rect["y1"]], 格子=rect["addr"] or "-",
                                出错=rect["error"] or "-")
    finally:
        book.Close(SaveChanges=False)


def _write(excel: Any, path: str, rects: list[dict], drop: list[str], style: dict, tick: Any) -> dict:
    """打开原文件加 / 删条件格式规则，同步工作表标签颜色，保存。返回 {"added": [框 id], "dropped": [标记 id]}"""
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0)
    book_name = Path(path).name
    added: list[str] = []
    dropped: list[str] = []
    try:
        if book.ReadOnly:
            raise ValueError(f"{book_name} 以只读方式打开，未能标色")
        marks = batch._marks()
        with wf.STATE_LOCK:
            keep = {m["tag"] for cid, m in marks.items() if m["path"] == path and cid not in drop and batch._is_rule(m)}
        holders: dict[str, Any] = {}
        _present, _gone, failed = batch._drop_rules(book, keep, tick, holders)
        for cid in drop:
            tag = marks.get(cid, {}).get("tag")
            if tag in failed:
                wf.log_msg(f"⚠️ {book_name}：没能去掉标记，下次再试：{failed[tag]}")
            else:
                dropped.append(cid)
        todo = [r for r in rects if r.get("addr") and not r.get("error")]
        if todo and batch._ungroup_sheets(book):
            wf.log_msg(f"ℹ️ {book_name}：文件里几张工作表是组合状态（同时选中），已解除组合再标色")
        for rect in todo:
            tick()
            cid = mark_id(rect["id"])
            sheet = _find_sheet(book, rect["sheet"])
            if sheet is None:
                rect["error"] = f"文件里找不到工作表「{rect['sheet']}」"
                continue
            mark = {"path": path, "sheet": str(sheet.Name), "addr": rect["addr"], "tag": batch._mark_name(cid),
                    "fill": batch._excel_color(style["fill"]), "font": batch._excel_color(style["font"]),
                    "saved": False, "source": SOURCE}
            with wf.STATE_LOCK:
                marks[cid] = mark  # 先记下再加规则：中途出错时之后还能按标识删掉
                wf._persist_state()
            try:
                batch._add_rule(sheet.Range(rect["addr"]), mark["tag"], mark)
            except Exception as err:
                with wf.STATE_LOCK:
                    marks.pop(cid, None)
                rect["error"] = batch._mark_blocker(book, sheet) or f"未能标色：{err}"
                debug_log.event("❌ PDF 标记标色失败", 位置=f"{sheet.Name}!{rect['addr']}", 原因=err)
                debug_log.exception(err, "PDF 标记标色")
                continue
            holders.setdefault(batch._tab_key(sheet), sheet)
            added.append(rect["id"])
            debug_log.event("🎨 PDF 标记标色", 位置=f"{sheet.Name}!{rect['addr']}", 标识=mark["tag"])
        tabs = batch._sync_tabs(book, path, holders)
        if added or dropped or tabs["marked"] or tabs["restored"]:
            tick()
            normalize_workbook_view(excel, book)
            before = batch._version_or_none(path)
            book.Save()
            batch._followed_own_save(path, before)
        with wf.STATE_LOCK:
            for rid in added:
                marks[mark_id(rid)]["saved"] = True
        parts = [f"标色 {len(added)} 处"] if added else []
        if dropped:
            parts.append(f"去掉标记 {len(dropped)} 处")
        if tabs["marked"]:
            parts.append(f"{tabs['marked']} 张工作表标签标成橙黄色")
        if tabs["restored"]:
            parts.append(f"{tabs['restored']} 张工作表标签改回原来的颜色")
        if parts:
            wf.log_msg(f"🖍️ {book_name}：" + "，".join(parts))
        return {"added": added, "dropped": dropped}
    finally:
        book.Close(SaveChanges=False)


def _finish_file(path: str, rects: list[dict], result: dict | None) -> None:
    """定位结果和标色结果写回框；跑的过程中被删掉的框，标记记进 drop"""
    with wf.STATE_LOCK:
        store, marks = _store(), batch._marks()
        for rect in rects:
            live = store["rects"].get(rect["id"])
            target = live or store["trash"].get(rect["id"])
            if target is not None:
                target.update(sheet=rect["sheet"], addr=rect.get("addr"), error=rect.get("error"))
            if live is None and mark_id(rect["id"]) in marks and mark_id(rect["id"]) not in store["drop"]:
                store["drop"].append(mark_id(rect["id"]))
        for cid in (result or {}).get("dropped", []):
            marks.pop(cid, None)
            if cid in store["drop"]:
                store["drop"].remove(cid)
        wf._persist_state()


def _run(plan: dict[str, dict], style: dict) -> None:
    excel = None
    watchdog = ExcelWatchdog(lambda: wf.log_msg("⚠️ 后台 Excel 长时间没反应，已结束它，这份文件下次再试"))
    initialized = False
    summary = {"marked": 0, "dropped": 0, "failed": 0, "files": 0}
    try:
        wf.pythoncom.CoInitialize()
        initialized = True
        watchdog.start()
        for done, (path, work) in enumerate(plan.items()):
            name = Path(path).name
            with _LOCK:
                _STATUS.update(done=done, current=name)
            debug_log.begin(f"PDF 标记写进 Excel {done + 1} / {len(plan)}：{name}")
            watchdog.progress()
            result = None
            try:
                batch._check_file(path)
                if excel is None:
                    excel = start_background_excel()
                    watchdog.register(excel)
                if work["add"]:
                    _locate(excel, path, work["add"], watchdog.progress)
                result = _write(excel, path, work["add"], work["drop"], style, watchdog.progress)
                summary["files"] += 1
            except Exception as err:
                for rect in work["add"]:
                    rect["error"] = rect.get("error") or str(err)
                wf.log_msg(f"❌ {name}：{err}")
                debug_log.exception(err, f"PDF 标记写进 Excel {name}")
                if excel is not None and not isinstance(err, ValueError):
                    try:
                        quit_background_excel(excel)
                    except Exception:
                        pass
                    excel = None
            _finish_file(path, work["add"], result)
            summary["marked"] += len((result or {}).get("added", []))
            summary["dropped"] += len((result or {}).get("dropped", []))
            summary["failed"] += sum(1 for r in work["add"] if r.get("error"))
        message = f"🖍️ PDF 标记写进 Excel：标色 {summary['marked']} 处"
        if summary["dropped"]:
            message += f"，去掉 {summary['dropped']} 处"
        if summary["failed"]:
            message += f"，{summary['failed']} 处没标成（看框上的说明）"
        wf.log_msg(message)
        summary["message"] = message
    except Exception as err:
        debug_log.exception(err, "PDF 标记写进 Excel")
        summary["message"] = f"❌ 出错：{err}"
    finally:
        try:
            if excel is not None:
                quit_background_excel(excel)
        finally:
            watchdog.stop()
            if initialized:
                wf.pythoncom.CoUninitialize()
            with _LOCK:
                _STATUS.update(running=False, done=len(plan), current="", result=summary)
            with wf.STATE_LOCK:
                wf.STATE["busy"] = False
                wf.STATE.pop("maintenance", None)


def apply(highlight: Any) -> None:
    """把还没标进 Excel 的框标上色，删掉的框去掉标记。和别的后台任务一样占着任务槽，一次只跑一个"""
    style = batch.highlight_style(highlight)
    with wf.STATE_LOCK, _LOCK:
        from core import bl_links, hidden_repair, link_repair, mark_finder
        if (wf.STATE.get("busy") or _STATUS.get("running") or link_repair.status().get("running")
                or mark_finder.status().get("running") or bl_links.status().get("running")
                or hidden_repair.status().get("running")):
            raise ValueError("后台任务正在执行，请等它结束")
        job = wf.STATE.get("job") or {}
        if job.get("pending_adjustment") or job.get("deferred_adjustments"):
            raise ValueError("还有未完成的文件写入，请先在工作台继续或放弃上次写入")
        plan = _plan()
        if not plan:
            raise ValueError("没有要写进 Excel 的标记")
        wf.STATE.update(busy=True, maintenance="pdf_review")
        _STATUS.clear()
        _STATUS.update(running=True, total=len(plan), done=0, current="")
        wf._persist_state()

    def run() -> None:
        with debug_log.side_task("pdf_review", {"files": len(plan), "style": style,
                                                "add": sum(len(w["add"]) for w in plan.values()),
                                                "drop": sum(len(w["drop"]) for w in plan.values())}):
            _run(plan, style)

    try:
        threading.Thread(target=run, name="PDF 标记写进 Excel", daemon=True).start()
    except Exception:
        with wf.STATE_LOCK, _LOCK:
            wf.STATE["busy"] = False
            wf.STATE.pop("maintenance", None)
            _STATUS["running"] = False
        raise


def status() -> dict:
    with _LOCK:
        return copy.deepcopy(_STATUS)
