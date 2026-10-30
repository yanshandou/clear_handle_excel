"""PDF 人工复检：探测底色编解码、从探测 PDF 读行列位置、框 → 单元格，以及框的保存 / 撤销。"""
import copy
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import batch, pdf_review, workflow as wf
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


class LocateTests(unittest.TestCase):
    rows = {1: (40, 60), 2: (60, 80), 3: (80, 100), 7: (100, 120), 8: (120, 140)}  # 4~6 行隐藏
    cols = {1: (50, 120), 2: (120, 200), 3: (200, 260)}

    def blocks(self, x0, y0, x1, y1):
        return pdf_review.blocks_for({"x0": x0, "y0": y0, "x1": x1, "y1": y1}, self.rows, self.cols)

    def test_rect_over_cells(self):
        self.assertEqual(self.blocks(55, 62, 190, 98), [(2, 3, 1, 2)])

    def test_slightly_touching_neighbours_are_left_out(self):
        # 框边只压到旁边格子一点点（不到三成）不算
        self.assertEqual(self.blocks(115, 57, 205, 83), [(2, 2, 2, 2)])

    def test_small_rect_inside_one_cell(self):
        self.assertEqual(self.blocks(130, 85, 140, 90), [(3, 3, 2, 2)])

    def test_hidden_rows_split_the_block(self):
        self.assertEqual(self.blocks(55, 82, 115, 138), [(3, 3, 1, 1), (7, 8, 1, 1)])
        self.assertEqual(pdf_review.block_address([(3, 3, 1, 1), (7, 8, 1, 1)]), "A3,A7:A8")

    def test_margin_only(self):
        self.assertEqual(self.blocks(300, 300, 400, 400), [])

    def test_long_address_falls_back_to_one_block(self):
        blocks = [(r, r, 1, 3) for r in range(1, 200, 2)]
        self.assertEqual(pdf_review.block_address(blocks), "A1:C199")


class SaveTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.book = self.tmp / "売上.xlsx"
        self.book.write_bytes(b"x")
        self.pdf = self.tmp / "売上_Sheet1.pdf"
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(self.pdf)
        self.state = copy.deepcopy(wf.STATE)
        self.state.update(pdf_review={}, pdf_exports={}, batch_review={"cards": {}, "rows": {}, "ignored": []})
        patcher = patch.object(wf, "STATE", self.state)
        patcher.start()
        self.addCleanup(patcher.stop)
        persist = patch.object(wf, "_persist_state")
        persist.start()
        self.addCleanup(persist.stop)

    def rect(self, rid, **extra):
        return {"id": rid, "page": 0, "x0": 10, "y0": 10, "x1": 100, "y1": 60, **extra}

    def test_list_guesses_owner_by_file_name(self):
        listed = pdf_review.list_pdfs(str(self.tmp))
        self.assertEqual(len(listed["pdfs"]), 1)
        entry = listed["pdfs"][0]
        self.assertEqual((entry["wb_name"], entry["sheet"], entry["guessed"]), ("売上.xlsx", "Sheet1", True))
        self.assertEqual(entry["pages"], [[595.0, 842.0]])

    def test_export_record_wins_over_guess(self):
        self.state["pdf_exports"][str(self.pdf)] = {"wb_path": str(self.book), "sheet": "Sheet/1"}
        entry = pdf_review.list_pdfs(str(self.tmp))["pdfs"][0]
        self.assertEqual((entry["sheet"], entry["guessed"]), ("Sheet/1", False))

    def test_delete_marked_rect_then_undo(self):
        pdf_review.list_pdfs(str(self.tmp))
        pdf_review.save(str(self.pdf), [self.rect("aaaaaaaa"), self.rect("bbbbbbbb")])
        rects = self.state["pdf_review"]["rects"]
        self.assertEqual(rects["aaaaaaaa"]["sheet"], "Sheet1")
        self.assertEqual(rects["aaaaaaaa"]["wb_path"], str(self.book))
        # 第一个框已经标进 Excel
        rects["aaaaaaaa"]["addr"] = "B2:C3"
        batch._marks()["pdfaaaaaaaa"] = {"path": str(self.book), "sheet": "Sheet1", "addr": "B2:C3",
                                         "tag": "CHE_MARK_pdfaaaaaaaa", "source": "pdf_review", "saved": True}
        snap = pdf_review.save(str(self.pdf), [self.rect("bbbbbbbb")])
        self.assertEqual(snap["drop"], 1)
        self.assertEqual([r["id"] for r in snap["rects"][str(self.pdf.resolve())]], ["bbbbbbbb"])
        # 撤销删除：框带着定位结果回来，不用再去掉标记
        snap = pdf_review.save(str(self.pdf), [self.rect("bbbbbbbb"), self.rect("aaaaaaaa")])
        self.assertEqual(snap["drop"], 0)
        states = {r["id"]: (r["state"], r["addr"]) for r in snap["rects"][str(self.pdf.resolve())]}
        self.assertEqual(states, {"aaaaaaaa": ("marked", "B2:C3"), "bbbbbbbb": ("pending", None)})

    def test_rejects_pdf_outside_folder_and_bad_rects(self):
        pdf_review.list_pdfs(str(self.tmp))
        with self.assertRaises(ValueError):
            pdf_review.save(str(self.pdf), [self.rect("zz")])
        with self.assertRaises(ValueError):
            pdf_review.save(str(self.pdf), [self.rect("aaaaaaaa", page=3)])
        other = Path(tempfile.mkdtemp()) / "x.pdf"
        other.write_bytes(self.pdf.read_bytes())
        with self.assertRaises(ValueError):
            pdf_review.save(str(other), [])

    def test_plan_groups_new_rects_and_drops_by_file(self):
        pdf_review.list_pdfs(str(self.tmp))
        pdf_review.save(str(self.pdf), [self.rect("aaaaaaaa")])
        batch._marks()["pdfcccccccc"] = {"path": str(self.book), "tag": "CHE_MARK_pdfcccccccc", "source": "pdf_review"}
        self.state["pdf_review"]["drop"] = ["pdfcccccccc", "pdfgone0000"]
        with wf.STATE_LOCK:
            plan = pdf_review._plan()
        self.assertEqual(list(plan), [str(self.book)])
        self.assertEqual([r["id"] for r in plan[str(self.book)]["add"]], ["aaaaaaaa"])
        self.assertEqual(plan[str(self.book)]["drop"], ["pdfcccccccc"])


class WorkbenchKeepsDrawnMarksTests(unittest.TestCase):
    def test_wanted_marks_keep_pdf_review_marks(self):
        state = copy.deepcopy(wf.STATE)
        state["batch_review"] = {"cards": {}, "rows": {}, "ignored": [], "highlights": {
            "pdfaaaaaaaa": {"path": "a.xlsx", "source": "pdf_review"},
            "card1": {"path": "a.xlsx"},
        }}
        with patch.object(wf, "STATE", state):
            self.assertEqual(batch._wanted_marks("a.xlsx"), {"pdfaaaaaaaa"})
            # 标识里只能有字母数字：PDF 标记的 id 整个都在标识里
            self.assertEqual(batch.MARK_TAG.search('=N("' + batch._mark_name("pdfaaaaaaaa") + '")=0').group(0),
                             "CHE_MARK_pdfaaaaaaaa")


if __name__ == "__main__":
    unittest.main()


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

    def Sheets(self, name):
        if name != self.sheet.Name:
            raise KeyError(name)
        return self.sheet

    def Close(self, SaveChanges=False):
        self.closed = True


class LocateInExcelTests(unittest.TestCase):
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

    def locate(self, rects):
        with patch.object(pdf_review, "open_workbook_checked", return_value=self.book), \
                patch.object(pdf_review.print_layout, "bind"), \
                patch.object(wf, "_export_sheet_pdf", side_effect=lambda sheet, out, *a: sheet.export(out)):
            pdf_review._locate(object(), "売上.xlsx", rects, lambda: None)

    def rect(self, rid, page, x0, y0, x1, y1):
        return {"id": rid, "pdf": str(self.pdf), "page": page, "sheet": "Sheet1",
                "x0": x0, "y0": y0, "x1": x1, "y1": y1, "addr": None, "error": None}

    def test_rects_map_to_cells_across_pages_hidden_rows_and_merged_cells(self):
        rects = [
            self.rect("a1", 0, 105, 72, 215, 108),   # 第 1 页 B2:C3
            self.rect("a2", 1, 45, 52, 95, 68),      # 第 2 页第 1 行 = 第 7 行（第 6 行隐藏）A7
            self.rect("a3", 1, 165, 72, 215, 88),    # 压到合并格 B8:C9 的右上角 → 扩到整个合并区域
            self.rect("a4", 0, 500, 700, 580, 800),  # 页边空白
        ]
        self.locate(rects)
        self.assertTrue(self.book.closed)
        self.assertEqual(self.sheet.exports, 2)
        result = {r["id"]: (r["addr"], r["error"]) for r in rects}
        self.assertEqual(result["a1"], ("B2:C3", None))
        self.assertEqual(result["a2"], ("A7", None))
        self.assertEqual(result["a3"], ("B8:C9", None))
        self.assertIsNone(result["a4"][0])
        self.assertIn("没有表格的格子", result["a4"][1])

    def test_page_count_mismatch_is_reported(self):
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(self.pdf)
        rects = [self.rect("a1", 0, 105, 72, 215, 108)]
        self.locate(rects)
        self.assertIn("对不上", rects[0]["error"])
        self.assertIsNone(rects[0]["addr"])


class WriteTests(unittest.TestCase):
    def setUp(self):
        self.state = copy.deepcopy(wf.STATE)
        self.state.update(pdf_review={}, batch_review={"cards": {}, "rows": {}, "ignored": [], "highlights": {
            "pdfold00000": {"path": "a.xlsx", "sheet": "Sheet1", "addr": "A1", "tag": "CHE_MARK_pdfold00000",
                            "source": "pdf_review", "saved": True},
            "card1": {"path": "a.xlsx", "sheet": "Sheet1", "addr": "D4", "tag": "CHE_MARK_card1", "saved": True},
        }})
        for target, value in ((wf, "STATE"), (wf, "_persist_state"), (wf, "log_msg")):
            p = patch.object(target, value, self.state) if value == "STATE" else patch.object(target, value)
            p.start()
            self.addCleanup(p.stop)

    def test_adds_rules_drops_removed_marks_keeps_workbench_marks(self):
        sheet = type("S", (), {"Name": "Sheet1", "Range": lambda _s, a: ("range", a)})()
        book = type("B", (), {"ReadOnly": False, "Worksheets": [sheet], "Save": lambda _s: None,
                              "Close": lambda _s, SaveChanges=False: None,
                              "Sheets": lambda _s, n: sheet})()
        dropped_keep = {}

        def drop_rules(_book, keep, tick=None, holders=None):
            dropped_keep["keep"] = set(keep)
            return set(keep), {"CHE_MARK_pdfold00000": 1}, {}

        added = []
        with patch.object(pdf_review, "open_workbook_checked", return_value=book), \
                patch.object(batch, "_drop_rules", side_effect=drop_rules), \
                patch.object(batch, "_ungroup_sheets", return_value=False), \
                patch.object(batch, "_add_rule", side_effect=lambda area, tag, mark: added.append((area, tag))), \
                patch.object(batch, "_sync_tabs", return_value={"marked": 1, "restored": 0, "failed": 0}), \
                patch.object(batch, "_followed_own_save"), patch.object(pdf_review, "normalize_workbook_view"):
            rects = [{"id": "abcdef12", "sheet": "Sheet1", "addr": "B2:C3", "error": None},
                     {"id": "abcdef13", "sheet": "Sheet1", "addr": None, "error": "框里没有表格的格子"}]
            result = pdf_review._write(object(), "a.xlsx", rects, ["pdfold00000"],
                                       {"fill": "#C0007A", "font": "#FFFFFF"}, lambda: None)
        self.assertEqual(dropped_keep["keep"], {"CHE_MARK_card1"})  # 工作台的标记留着
        self.assertEqual(result, {"added": ["abcdef12"], "dropped": ["pdfold00000"]})
        self.assertEqual(added, [(("range", "B2:C3"), "CHE_MARK_pdfabcdef12")])
        mark = batch._marks()["pdfabcdef12"]
        self.assertEqual((mark["addr"], mark["source"], mark["saved"]), ("B2:C3", "pdf_review", True))
        pdf_review._finish_file("a.xlsx", rects, result)
        self.assertNotIn("pdfold00000", batch._marks())


class ClearOldPdfTests(unittest.TestCase):
    def test_deletes_only_this_workbooks_old_pdfs_and_their_boxes(self):
        tmp = Path(tempfile.mkdtemp())
        book = tmp / "★売上.xlsx"
        book.write_bytes(b"x")
        old = [tmp / "売上_Sheet1.pdf", tmp / "売上_削除した表_2.PDF"]
        keep = [tmp / "取扱説明書.pdf", tmp / "売上.pdf", tmp / "その他_売上_Sheet1.pdf"]
        for path in old + keep:
            path.write_bytes(b"%PDF")
        state = copy.deepcopy(wf.STATE)
        state.update(pdf_exports={str(old[0]): {"wb_path": str(book), "sheet": "Sheet1"}},
                     pdf_review={"rects": {"aaaaaaaa": {"id": "aaaaaaaa", "pdf": str(old[0])},
                                           "bbbbbbbb": {"id": "bbbbbbbb", "pdf": str(keep[0])}},
                                 "drop": [], "trash": {}})
        with patch.object(wf, "STATE", state), patch.object(wf, "_persist_state"), patch.object(wf, "log_msg"):
            wf._clear_old_pdfs([book])
        self.assertTrue(all(not p.exists() for p in old))
        self.assertTrue(all(p.exists() for p in keep))
        self.assertEqual(state["pdf_exports"], {})
        self.assertEqual(list(state["pdf_review"]["rects"]), ["bbbbbbbb"])
