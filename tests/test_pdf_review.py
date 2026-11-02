"""PDF 人工复检：探测底色编解码、从探测 PDF 读行列位置、点的位置 → 单元格、打开 Excel（和原始备份）。"""
import copy
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from core import pdf_review, workflow as wf
from core.pdf_detector import fitz


def rgb_of(color: int) -> tuple[float, float, float]:
    """Excel 的 Color（BGR）→ PDF 里的填充色（0~1）"""
    return (color & 255) / 255, ((color >> 8) & 255) / 255, ((color >> 16) & 255) / 255


def probe_pdf(path: Path, lines: list[tuple[int, float, float]], axis: str, merged: list | None = None) -> None:
    """模拟 Excel 导出的探测 PDF：每一行（列）的每一格各画一块底色"""
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    for index, lo, hi in lines:
        for k in range(4):  # 一行 4 格
            a, b = 50 + 100 * k, 150 + 100 * k
            rect = fitz.Rect(a, lo, b, hi) if axis == "y" else fitz.Rect(lo, a, hi, b)
            page.draw_rect(rect, color=None, fill=rgb_of(pdf_review.probe_color(index)))
    for index, rect in merged or []:
        page.draw_rect(fitz.Rect(*rect), color=None, fill=rgb_of(pdf_review.probe_color(index)))
    doc.save(path)


class ProbeColorTests(unittest.TestCase):
    def test_round_trip_and_tolerance(self):
        for index in (1, 2, 62, 63, 3968, 3969, 20000, 250046):
            fill = rgb_of(pdf_review.probe_color(index))
            self.assertEqual(pdf_review.probe_index(fill), index)
            # PDF 里读回来差 1 也认得出
            nudged = tuple(min(1.0, v + 1 / 255) for v in fill)
            self.assertEqual(pdf_review.probe_index(nudged), index)

    def test_ordinary_colors_are_not_probe_colors(self):
        for fill in ((1, 1, 1), (0, 0, 0), (1, 1, 0), (0.5, 0.5, 0.5), None, (0.2,)):
            self.assertIsNone(pdf_review.probe_index(fill))


class BandTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def test_rows_use_the_most_common_span_not_merged_cells(self):
        pdf = self.tmp / "rows.pdf"
        # 第 3 行有一个合并格跨到第 5 行：第 3 行的位置仍是它自己那一行
        probe_pdf(pdf, [(1, 40, 60), (3, 60, 80), (4, 80, 100), (5, 100, 130)], "y",
                  merged=[(3, (450, 60, 550, 130))])
        bands = pdf_review.read_bands(pdf, "y")
        self.assertEqual(bands[0], {1: (40.0, 60.0), 3: (60.0, 80.0), 4: (80.0, 100.0), 5: (100.0, 130.0)})

    def test_columns(self):
        pdf = self.tmp / "cols.pdf"
        probe_pdf(pdf, [(2, 50, 120), (3, 120, 200)], "x")
        self.assertEqual(pdf_review.read_bands(pdf, "x")[0], {2: (50.0, 120.0), 3: (120.0, 200.0)})


class CellAtTests(unittest.TestCase):
    rows = {1: (40, 60), 2: (60, 80), 3: (80, 100), 7: (100, 120), 8: (120, 140)}  # 4~6 行隐藏
    cols = {1: (50, 120), 2: (120, 200), 3: (200, 260)}

    def test_point_inside_cells(self):
        self.assertEqual(pdf_review.cell_at(self.rows, self.cols, 60, 45), "A1")
        self.assertEqual(pdf_review.cell_at(self.rows, self.cols, 150, 110), "B7")  # 隐藏行后面接着第 7 行

    def test_point_on_seam_or_margin(self):
        self.assertEqual(pdf_review.cell_at(self.rows, self.cols, 260.5, 139), "C8")  # 格线外 1 个点以内也算
        self.assertIsNone(pdf_review.cell_at(self.rows, self.cols, 300, 50))  # 页边空白
        self.assertIsNone(pdf_review.cell_at(self.rows, self.cols, 60, 10))

    def test_overlap_takes_the_narrowest(self):
        rows = {3: (60, 100), 4: (70, 80)}
        self.assertEqual(pdf_review.cell_at(rows, self.cols, 60, 75), "A4")


class ReviewFolderTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.book = self.tmp / "営業" / "売上.xlsx"
        self.book.parent.mkdir()
        self.book.write_bytes(b"x")
        self.pdf = self.book.parent / "売上_Sheet1.pdf"
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(self.pdf)
        self.state = copy.deepcopy(wf.STATE)
        self.state.update(pdf_review={}, pdf_exports={}, busy=False, job=None)
        for target, name, value in ((wf, "STATE", self.state), (pdf_review, "_MAPS", {}), (pdf_review, "_QUEUE", [])):
            patcher = patch.object(target, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        persist = patch.object(wf, "_persist_state")
        persist.start()
        self.addCleanup(persist.stop)
        prepare = patch.object(pdf_review, "prepare")  # 不真的起后台线程去算
        self.prepare = prepare.start()
        self.addCleanup(prepare.stop)

    def ready(self):
        pdf = str(self.pdf.resolve())
        pdf_review._MAPS[pdf] = {"v": self.pdf.stat().st_mtime_ns, "state": "ready", "sheet": "Sheet1",
                                 "rows": [{5: (100.0, 120.0)}], "cols": [{3: (200.0, 260.0)}]}

    def test_list_guesses_owner_and_queues_all_pdfs(self):
        listed = pdf_review.list_pdfs(str(self.tmp))
        entry = listed["pdfs"][0]
        self.assertEqual((entry["wb_name"], entry["sheet"], entry["guessed"]), ("営業/売上.xlsx", "Sheet1", True))
        self.assertEqual(entry["pages"], [[595.0, 842.0]])
        self.prepare.assert_called_once_with([str(self.pdf)])
        self.assertEqual(self.state["pdf_review"], {"folder": str(self.tmp.resolve())})

    def test_export_record_wins_over_guess(self):
        self.state["pdf_exports"][str(self.pdf)] = {"wb_path": str(self.book), "sheet": "Sheet/1"}
        entry = pdf_review.list_pdfs(str(self.tmp))["pdfs"][0]
        self.assertEqual((entry["sheet"], entry["guessed"]), ("Sheet/1", False))

    def test_pdfs_folder_is_listed_alone_and_manifest_finds_excel(self):
        # 2026-11-02 起 PDF 放在 根目录/PDFS/<子文件夹>；有 PDFS 时只看里面的，旁边混放的旧 PDF 不列
        from core import pdf_store
        inside = self.tmp / "PDFS" / "営業" / "売上_表_1.pdf"
        inside.parent.mkdir(parents=True)
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(inside)
        pdf_store.record(self.tmp, self.book, "h", [(inside, "表/1")])
        listed = pdf_review.list_pdfs(str(self.tmp))
        self.assertEqual([e["name"] for e in listed["pdfs"]], ["営業/売上_表_1.pdf"])
        entry = listed["pdfs"][0]
        self.assertEqual((entry["wb_name"], entry["sheet"], entry["guessed"]), ("営業/売上.xlsx", "表/1", False))
        self.assertTrue(listed["in_pdfs"])

    def test_pdfs_folder_without_record_guesses_excel_in_matching_subfolder(self):
        inside = self.tmp / "PDFS" / "営業" / "売上_Sheet2.pdf"
        inside.parent.mkdir(parents=True)
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(inside)
        entry = pdf_review.list_pdfs(str(self.tmp))["pdfs"][0]
        self.assertEqual((entry["wb_name"], entry["sheet"], entry["guessed"]), ("営業/売上.xlsx", "Sheet2", True))

    def test_map_not_ready_is_queued_first(self):
        pdf_review.list_pdfs(str(self.tmp))
        self.prepare.reset_mock()
        self.assertEqual(pdf_review.open_at(str(self.pdf), 0, 210, 110), {"pending": True})
        self.prepare.assert_called_once_with([str(self.pdf.resolve())], first=True)

    def test_map_is_dropped_when_pdf_changes(self):
        pdf_review.list_pdfs(str(self.tmp))
        self.ready()
        self.assertEqual(pdf_review.maps_status()[str(self.pdf.resolve())]["state"], "ready")
        pdf_review._MAPS[str(self.pdf.resolve())]["v"] -= 1  # PDF 重新导出过
        self.assertEqual(pdf_review.maps_status()[str(self.pdf.resolve())]["state"], "none")
        self.assertEqual(pdf_review.open_at(str(self.pdf), 0, 210, 110), {"pending": True})

    def test_open_current_only(self):
        pdf_review.list_pdfs(str(self.tmp))
        self.ready()
        with patch.object(pdf_review, "open_workbook_in_excel", return_value="") as opened, \
                patch.object(pdf_review.repair_compare, "open_comparison") as compare:
            result = pdf_review.open_at(str(self.pdf), 0, 210, 110)
            outside = pdf_review.open_at(str(self.pdf), 0, 500, 700)
        self.assertEqual(result, {"sheet": "Sheet1", "cell": "C5", "backup": False, "note": ""})
        opened.assert_any_call(self.book, "Sheet1", "C5")
        self.assertEqual(outside["cell"], "")
        self.assertIn("不在表格里", outside["note"])
        compare.assert_not_called()

    def test_open_with_backup_jumps_to_same_cell(self):
        old = Path(tempfile.mkdtemp())
        (old / "営業").mkdir()
        backup = old / "営業" / "売上.xlsx"
        backup.write_bytes(b"old")
        pdf_review.list_pdfs(str(self.tmp))
        self.ready()
        with patch.object(pdf_review.repair_compare, "open_comparison", return_value={"opened": ["current", "backup"], "note": ""}) as compare, \
                patch.object(pdf_review, "open_workbook_in_excel") as opened:
            result = pdf_review.open_at(str(self.pdf), 0, 210, 110, str(old))
        self.assertTrue(result["backup"])
        target = compare.call_args[0][0]
        self.assertEqual((target["current"], target["backup"], target["sheet"], target["old_sheet"], target["cell"]),
                         (self.book.resolve(), backup.resolve(), "Sheet1", "Sheet1", "C5"))
        opened.assert_not_called()

    def test_backup_not_found_still_opens_current(self):
        old = Path(tempfile.mkdtemp())
        pdf_review.list_pdfs(str(self.tmp))
        self.ready()
        with patch.object(pdf_review, "open_workbook_in_excel", return_value="") as opened:
            result = pdf_review.open_at(str(self.pdf), 0, 210, 110, str(old))
            missing = pdf_review.open_at(str(self.pdf), 0, 210, 110, str(old / "nope"))
        self.assertFalse(result["backup"])
        self.assertIn("找不到这份文件", result["note"])
        self.assertIn("不存在", missing["note"])
        self.assertEqual(opened.call_count, 2)

    def test_open_is_blocked_while_scan_still_needs_the_file(self):
        pdf_review.list_pdfs(str(self.tmp))
        self.ready()
        self.state.update(busy=True, job={"kind": "batch_commit"})
        with patch.object(pdf_review, "open_workbook_in_excel") as opened, self.assertRaises(ValueError):
            pdf_review.open_at(str(self.pdf), 0, 210, 110)
        opened.assert_not_called()

    def test_rejects_pdf_outside_folder(self):
        pdf_review.list_pdfs(str(self.tmp))
        other = Path(tempfile.mkdtemp()) / "x.pdf"
        other.write_bytes(self.pdf.read_bytes())
        with self.assertRaises(ValueError):
            pdf_review.open_at(str(other), 0, 1, 1)
        with self.assertRaises(ValueError):
            pdf_review.open_at(str(self.pdf), "0", 1, 1)


class WorkQueueTests(unittest.TestCase):
    def test_prepare_puts_clicked_pdf_first_and_skips_fresh(self):
        tmp = Path(tempfile.mkdtemp())
        pdfs = []
        for name in ("a.pdf", "b.pdf", "c.pdf"):
            (tmp / name).write_bytes(b"%PDF")
            pdfs.append(str(tmp / name))
        maps = {pdfs[2]: {"v": (tmp / "c.pdf").stat().st_mtime_ns, "state": "ready"}}
        with patch.object(pdf_review, "_MAPS", maps), patch.object(pdf_review, "_QUEUE", []) as queue, \
                patch.object(pdf_review, "_WORKER", object()):  # 当作已有线程在跑，不真的起线程
            pdf_review.prepare(pdfs)
            self.assertEqual(queue, pdfs[:2])
            pdf_review.prepare([pdfs[1]], first=True)
            self.assertEqual(queue, [pdfs[1], pdfs[0]])
            self.assertEqual({p: s["state"] for p, s in pdf_review.maps_status().items()},
                             {pdfs[0]: "pending", pdfs[1]: "pending", pdfs[2]: "ready"})

    def test_forget_pdfs_drops_maps(self):
        tmp = Path(tempfile.mkdtemp())
        pdf = str(tmp / "a.pdf")
        state = copy.deepcopy(wf.STATE)
        state["pdf_exports"] = {pdf: {"wb_path": "x", "sheet": "S"}}
        with patch.object(wf, "STATE", state), patch.object(wf, "_persist_state"), \
                patch.object(pdf_review, "_MAPS", {pdf: {"v": 1}}) as maps, patch.object(pdf_review, "_QUEUE", [pdf]) as queue:
            pdf_review.forget_pdfs([Path(pdf)])
            self.assertEqual((maps, queue, state["pdf_exports"]), ({}, [], {}))


def cells_of(address: str) -> list[tuple[int, int]]:
    from openpyxl.utils.cell import range_boundaries
    result = []
    for part in address.replace("$", "").split(","):
        c1, r1, c2, r2 = range_boundaries(part if ":" in part else f"{part}:{part}")
        result += [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
    return result


class FakeArea:
    def __init__(self, sheet, address):
        self.sheet, self.address = sheet, address
        cells = cells_of(address)
        rows, cols = [r for r, _ in cells], [c for _, c in cells]
        self.Row, self.Column = min(rows), min(cols)
        self.Rows = type("N", (), {"Count": max(rows) - min(rows) + 1})()
        self.Columns = type("N", (), {"Count": max(cols) - min(cols) + 1})()
        self.Areas = type("A", (), {"Count": 1, "__call__": lambda _s, i: self})()
        area = self

        class Interior:
            Pattern = 1

            @property
            def Color(self):
                return None

            @Color.setter
            def Color(self, value):
                for cell in cells_of(area.address):
                    area.sheet.fills[cell] = value
        self.Interior = Interior()

    @property
    def MergeCells(self):
        merged = [m for m in self.sheet.merged if set(cells_of(m)) & set(cells_of(self.address))]
        return bool(merged) if not merged or len(cells_of(self.address)) == 1 else None

    @property
    def MergeArea(self):
        for m in self.sheet.merged:
            if set(cells_of(self.address)) <= set(cells_of(m)):
                return FakeArea(self.sheet, m)
        return self


class FakeProbeSheet:
    """10 行 × 5 列，第 6 行隐藏（不打印）；每行 20pt 高、每列 60pt 宽，第 1~5 行在第 1 页，其余在第 2 页"""
    def __init__(self, name="Sheet1"):
        self.Name = name
        self.fills = {}
        self.merged = ["B8:C9"]
        self.ProtectContents = False
        self.PageSetup = type("P", (), {"PrintArea": "", "PrintTitleRows": "", "PrintTitleColumns": "",
                                        "BlackAndWhite": False})()
        self.UsedRange = FakeArea(self, "A1:E10")
        self.Cells = type("C", (), {"FormatConditions": type("F", (), {"Delete": lambda _s: None})()})()
        self.exports = 0

    def Range(self, address):
        return FakeArea(self, address)

    def CellsAt(self, r, c):
        return FakeArea(self, f"{chr(64 + c)}{r}")

    def export(self, out):
        self.exports += 1
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.new_page(width=595, height=842)
        for (r, c), color in self.fills.items():
            if r == 6:
                continue
            page, top = (doc[0], 50 + 20 * (r - 1)) if r <= 5 else (doc[1], 50 + 20 * (r - 7))
            rect = fitz.Rect(40 + 60 * (c - 1), top, 100 + 60 * (c - 1), top + 20)
            if any((r, c) in cells_of(m) and (r, c) != cells_of(m)[0] for m in self.merged):
                continue  # 合并格只画左上角那一块（盖住整个合并区域）
            for m in self.merged:
                if (r, c) == cells_of(m)[0]:
                    rows = {x for x, _ in cells_of(m)}
                    cols = {y for _, y in cells_of(m)}
                    rect = fitz.Rect(rect.x0, rect.y0, rect.x0 + 60 * len(cols), rect.y0 + 20 * len(rows))
            page.draw_rect(rect, color=None, fill=rgb_of(color))
        doc.save(out)


class FakeProbeBook:
    def __init__(self, sheet):
        self.sheet = sheet
        self.Worksheets = [sheet]
        self.closed = False

    @property
    def Sheets(self):
        sheet = self.sheet

        class Sheets(list):  # 像 COM 一样能按名字取，也能遍历
            def __call__(self, name):
                if name != sheet.Name:
                    raise KeyError(name)
                return sheet
        return Sheets([sheet])

    def Close(self, SaveChanges=False):
        self.closed = True


class ProbeInExcelTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.sheet = FakeProbeSheet()
        self.sheet.Cells = type("C", (), {
            "FormatConditions": type("F", (), {"Delete": lambda _s: None})(),
            "__call__": lambda _s, r, c: self.sheet.CellsAt(r, c)})()
        self.book = FakeProbeBook(self.sheet)
        self.pdf = self.tmp / "売上_Sheet1.pdf"
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.new_page(width=595, height=842)
        doc.save(self.pdf)

    def probe(self):
        with patch.object(wf, "_export_sheet_pdf", side_effect=lambda sheet, out, *a: sheet.export(out)):
            return pdf_review._probe_pdf(self.book, str(self.pdf), "Sheet1", lambda: None)

    def test_points_map_to_cells_across_pages_and_hidden_rows(self):
        found = self.probe()
        self.assertEqual((found["state"], found["sheet"], self.sheet.exports), ("ready", "Sheet1", 2))
        at = lambda page, x, y: pdf_review.cell_at(found["rows"][page], found["cols"][page], x, y)
        self.assertEqual(at(0, 110, 75), "B2")     # 第 1 页第 2 行第 2 列
        self.assertEqual(at(1, 60, 60), "A7")      # 第 2 页第 1 行 = 第 7 行（第 6 行隐藏）
        self.assertEqual(at(1, 230, 95), "D9")
        self.assertIsNone(at(0, 500, 700))         # 页边空白

    def test_page_count_mismatch_is_reported(self):
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(self.pdf)
        found = self.probe()
        self.assertEqual(found["state"], "error")
        self.assertIn("对不上", found["error"])

    def test_missing_sheet(self):
        found = pdf_review._probe_pdf(self.book, str(self.pdf), "別の表", lambda: None)
        self.assertIn("找不到工作表", found["error"])


class ClearOldPdfTests(unittest.TestCase):
    def test_deletes_only_this_workbooks_old_pdfs(self):
        tmp = Path(tempfile.mkdtemp())
        book = tmp / "★売上.xlsx"
        book.write_bytes(b"x")
        old = [tmp / "売上_Sheet1.pdf", tmp / "売上_削除した表_2.PDF"]
        keep = [tmp / "取扱説明書.pdf", tmp / "売上.pdf", tmp / "その他_売上_Sheet1.pdf"]
        for path in old + keep:
            path.write_bytes(b"%PDF")
        state = copy.deepcopy(wf.STATE)
        state.update(pdf_exports={str(old[0]): {"wb_path": str(book), "sheet": "Sheet1"}})
        with patch.object(wf, "STATE", state), patch.object(wf, "_persist_state"), patch.object(wf, "log_msg"):
            wf._clear_old_pdfs([book])
        self.assertTrue(all(not p.exists() for p in old))
        self.assertTrue(all(p.exists() for p in keep))
        self.assertEqual(state["pdf_exports"], {})


class FakeSheet:
    def __init__(self, name, visible=-1):
        self.Name, self.Visible = name, visible


class FakeBook:
    def __init__(self, names, hidden=()):
        self.sheets = [FakeSheet(n, 0 if n in hidden else -1) for n in names]
        self.saved = 0
        self.closed = False
        book = self

        class Sheets(list):
            def __call__(self, key):
                return next(s for s in book.sheets if s.Name == key)
        self.Sheets = Sheets(self.sheets)
        self.Worksheets = self.Sheets

    def Save(self):
        self.saved += 1

    def Close(self, SaveChanges=False):
        self.closed = True


class ProcessTests(unittest.TestCase):
    """一键处理：标红的表在后台删掉并保存、对应 PDF 删掉；标黄的打开 Excel"""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.book_path = self.tmp / "売上.xlsx"
        self.book_path.write_bytes(b"x")
        self.pdfs = {}
        for sheet in ("一覧", "JSONサンプル", "明細"):
            pdf = self.tmp / f"売上_{sheet}.pdf"
            doc = fitz.open()
            doc.new_page(width=595, height=842)
            doc.save(pdf)
            self.pdfs[sheet] = pdf
        self.state = copy.deepcopy(wf.STATE)
        self.state.update(pdf_review={}, pdf_exports={}, busy=False, job=None)
        for target, name, value in ((wf, "STATE", self.state), (pdf_review, "_MAPS", {}), (pdf_review, "_QUEUE", []),
                                    (pdf_review, "_PROCESS", {"running": False})):
            patcher = patch.object(target, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        for target, name in ((wf, "_persist_state"), (wf, "log_msg"), (pdf_review, "prepare"),
                             (pdf_review.file_problem, "Guard"), (pdf_review, "start_background_excel"),
                             (pdf_review.sheet_history, "record")):
            patcher = patch.object(target, name)
            patcher.start()
            self.addCleanup(patcher.stop)
        pdf_review.file_problem.Guard.return_value.problem.return_value = (False, "打不开")
        pdf_review.file_problem.Guard.return_value.closed = False
        self.fp = patch.object(pdf_review.sheet_history, "fingerprint", return_value="h")
        self.fp.start()
        self.addCleanup(self.fp.stop)
        pdf_review.list_pdfs(str(self.tmp))

    def run_process(self, red=(), yellow=(), book=None, deleted=None):
        book = book or FakeBook(["一覧", "JSONサンプル", "明細"])
        calls = []

        def fake_delete(b, names):
            calls.append(list(names))
            b.sheets[:] = [s for s in b.sheets if s.Name not in names]
            return []
        with patch.object(pdf_review, "open_workbook_checked", return_value=book), \
                patch.object(pdf_review, "delete_sheets_preserving_values", side_effect=fake_delete), \
                patch.object(pdf_review, "open_workbook_in_excel", return_value="") as opened, \
                patch.object(pdf_review, "workbook_locked", return_value=False):
            pdf_review.process([{"pdf": str(self.pdfs[s])} for s in red],
                               [{"pdf": str(self.pdfs[s]), **extra} for s, extra in yellow])
            for _ in range(200):
                if not pdf_review.process_status()["running"]:
                    break
                time.sleep(0.02)
        status = pdf_review.process_status()
        self.assertFalse(status["running"])
        return status["result"], calls, book, opened

    def test_red_sheets_are_deleted_and_saved_with_pdf_removed(self):
        result, calls, book, _ = self.run_process(red=["JSONサンプル"])
        self.assertEqual(calls, [["JSONサンプル"]])
        self.assertEqual(book.saved, 1)
        self.assertTrue(book.closed)
        self.assertEqual([(e["sheet"], e["pdf_removed"]) for e in result["deleted"]], [("JSONサンプル", True)])
        self.assertFalse(self.pdfs["JSONサンプル"].exists())
        self.assertTrue(self.pdfs["一覧"].exists())
        self.assertFalse(self.state["busy"])
        pdf_review.sheet_history.record.assert_called_once()

    def test_two_red_sheets_in_one_file_open_it_once(self):
        result, calls, book, _ = self.run_process(red=["JSONサンプル", "明細"])
        self.assertEqual(calls, [["JSONサンプル", "明細"]])
        self.assertEqual(book.saved, 1)
        self.assertEqual(len(result["deleted"]), 2)

    def test_last_visible_sheet_is_not_deleted(self):
        book = FakeBook(["一覧", "JSONサンプル", "明細"], hidden=("一覧", "明細"))
        result, calls, book, _ = self.run_process(red=["JSONサンプル"], book=book)
        self.assertEqual(calls, [])
        self.assertEqual(book.saved, 0)
        self.assertIn("唯一", result["failed"][0]["reason"])
        self.assertTrue(self.pdfs["JSONサンプル"].exists())

    def test_yellow_opens_each_file_once_at_the_first_marked_cell(self):
        pdf = str(self.pdfs["一覧"].resolve())
        pdf_review._MAPS[pdf] = {"v": self.pdfs["一覧"].stat().st_mtime_ns, "state": "ready", "sheet": "一覧",
                                 "rows": [{5: (100.0, 120.0)}, {9: (100.0, 120.0)}],
                                 "cols": [{3: (200.0, 260.0)}, {2: (200.0, 260.0)}]}
        result, calls, _, opened = self.run_process(
            yellow=[("一覧", {"spots": [{"page": 1, "x": 210, "y": 110}, {"page": 0, "x": 210, "y": 110}]}),
                    ("明細", {"spots": [{"page": 0}]})])
        opened.assert_called_once_with(self.book_path, "一覧", "C5")  # 第 1 页那处先
        self.assertEqual([(e["sheet"], e["page"], e["cell"]) for e in result["opened"]],
                         [("一覧", 0, "C5"), ("一覧", 1, "B9"), ("明細", 0, "")])
        self.assertEqual([e["first"] for e in result["opened"]], [True, False, False])
        self.assertEqual(calls, [])

    def test_red_then_yellow_in_same_run(self):
        result, calls, _, opened = self.run_process(red=["JSONサンプル"], yellow=[("明細", {})])
        self.assertEqual(calls, [["JSONサンプル"]])
        opened.assert_called_once_with(self.book_path, "明細", "")

    def test_more_than_limit_files_wait_for_next_press(self):
        with patch.object(pdf_review, "OPEN_LIMIT", 0):
            result, _, _, opened = self.run_process(yellow=[("明細", {})])
        opened.assert_not_called()
        self.assertEqual(len(result["later"]), 1)

    def test_locked_file_asks_and_records_failure(self):
        with patch.object(pdf_review, "open_workbook_checked") as open_book, \
                patch.object(pdf_review, "workbook_locked", return_value=True):
            pdf_review.process([{"pdf": str(self.pdfs["明細"])}], [])
            for _ in range(200):
                if not pdf_review.process_status()["running"]:
                    break
                time.sleep(0.02)
        open_book.assert_not_called()
        guard = pdf_review.file_problem.Guard.return_value
        self.assertIn("Excel 中打开", str(guard.problem.call_args[0][0]))
        self.assertEqual(pdf_review.process_status()["result"]["failed"][0]["reason"], "打不开")
        self.assertTrue(self.pdfs["明細"].exists())

    def test_rejects_busy_or_outside_pdf(self):
        self.state["busy"] = True
        with self.assertRaises(ValueError):
            pdf_review.process([{"pdf": str(self.pdfs["明細"])}], [])
        self.state["busy"] = False
        other = Path(tempfile.mkdtemp()) / "x.pdf"
        other.write_bytes(self.pdfs["明細"].read_bytes())
        with self.assertRaises(ValueError):
            pdf_review.process([{"pdf": str(other)}], [])
        with self.assertRaises(ValueError):
            pdf_review.process([], [])


if __name__ == "__main__":
    unittest.main()


class RegenerateTests(unittest.TestCase):
    """复检时只重新生成一份 Excel 的 PDF"""

    def setUp(self):
        from core import pdf_store
        self.store = pdf_store
        self.tmp = Path(tempfile.mkdtemp())
        self.book = self.tmp / "売上.xlsx"
        self.book.write_bytes(b"x")
        self.pdf = self.tmp / "PDFS" / "売上_一覧.pdf"
        self.pdf.parent.mkdir()
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(self.pdf)
        pdf_store.record(self.tmp, self.book, pdf_review.sheet_history.fingerprint(self.book), [(self.pdf, "一覧")])
        self.state = copy.deepcopy(wf.STATE)
        self.state.update(pdf_review={}, pdf_exports={}, busy=False, job=None)
        for target, name, value in ((wf, "STATE", self.state), (pdf_review, "_MAPS", {}), (pdf_review, "_QUEUE", []),
                                    (pdf_review, "_REGEN", {"running": False})):
            patcher = patch.object(target, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        for target, name in ((wf, "_persist_state"), (wf, "log_msg"), (pdf_review, "prepare"),
                             (pdf_review.file_problem, "Guard"), (pdf_review, "start_background_excel")):
            patcher = patch.object(target, name)
            patcher.start()
            self.addCleanup(patcher.stop)
        pdf_review.list_pdfs(str(self.tmp))

    def wait(self):
        for _ in range(200):
            if not pdf_review.regen_status()["running"]:
                break
            time.sleep(0.02)
        status = pdf_review.regen_status()
        self.assertFalse(status["running"])
        return status["result"]

    def test_unchanged_excel_asks_first(self):
        with patch.object(wf, "_export_workbook_pdfs") as export:
            self.assertEqual(pdf_review.regenerate(str(self.pdf)),
                             {"started": False, "unchanged": True, "name": "売上.xlsx"})
        export.assert_not_called()

    def test_changed_or_confirmed_exports_only_this_workbook(self):
        self.book.write_bytes(b"edited and saved")
        with patch.object(wf, "_export_workbook_pdfs", return_value=[self.pdf]) as export:
            self.assertEqual(pdf_review.regenerate(str(self.pdf)), {"started": True, "name": "売上.xlsx"})
            result = self.wait()
        self.assertEqual(result["ok"], True)
        self.assertEqual(result["pdfs"], [str(self.pdf)])
        args, kwargs = export.call_args
        self.assertEqual((args[1], args[2], kwargs), (self.book, self.tmp.resolve(), {"force": True}))
        self.assertFalse(self.state["busy"])

    def test_failure_is_reported_and_busy_released(self):
        pdf_review.file_problem.Guard.return_value.problem.return_value = (False, "文件打不开")
        pdf_review.file_problem.Guard.return_value.closed = False
        with patch.object(wf, "_export_workbook_pdfs", side_effect=RuntimeError("boom")):
            pdf_review.regenerate(str(self.pdf), True)
            result = self.wait()
        self.assertEqual((result["ok"], result["reason"]), (False, "文件打不开"))
        self.assertFalse(self.state["busy"])

    def test_busy_is_rejected(self):
        self.state["busy"] = True
        with self.assertRaises(ValueError):
            pdf_review.regenerate(str(self.pdf), True)
