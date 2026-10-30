"""PDF 人工复检：探测底色编解码、从探测 PDF 读行列位置、点的位置 → 单元格、打开 Excel（和原始备份）。"""
import copy
import tempfile
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

    def test_map_not_ready_is_queued_first(self):
        pdf_review.list_pdfs(str(self.tmp))
        self.prepare.reset_mock()
        self.assertEqual(pdf_review.open_at(str(self.pdf), 0, 210, 110), {"pending": True})
        self.prepare.assert_called_once_with([str(self.pdf.resolve())], first=True)
        self.assertEqual(pdf_review.map_for(str(self.pdf))["state"], "none")

    def test_map_is_dropped_when_pdf_changes(self):
        pdf_review.list_pdfs(str(self.tmp))
        self.ready()
        self.assertEqual(pdf_review.map_for(str(self.pdf))["pages"], [{"rows": [[5, 100.0, 120.0]], "cols": [[3, 200.0, 260.0]]}])
        pdf_review._MAPS[str(self.pdf.resolve())]["v"] -= 1  # PDF 重新导出过
        self.assertNotEqual(pdf_review.map_for(str(self.pdf))["state"], "ready")

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


if __name__ == "__main__":
    unittest.main()
