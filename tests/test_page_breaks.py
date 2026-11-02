import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, fix_images, page_breaks, reporter, storage, workflow
from core.pdf_detector import fitz

MANUAL, NONE, AUTO = page_breaks.XL_PAGEBREAK_MANUAL, page_breaks.XL_PAGEBREAK_NONE, -4105


class FakeArea:
    def __init__(self, sheet, top, left, bottom, right):
        self.sheet, self.Row, self.Column = sheet, top, left
        self.Rows = Mock(Count=bottom - top + 1)
        self.Columns = Mock(Count=right - left + 1)

    def Cells(self, r, c):
        return self.sheet.Cells(self.Row + r - 1, self.Column + c - 1)


class FakeCell:
    def __init__(self, sheet, row, col):
        self.sheet, self.row, self.col = sheet, row, col

    @property
    def Text(self):
        return self.sheet.texts.get((self.row, self.col), "")

    def _area(self):
        for top, left, bottom, right in self.sheet.merges:
            if top <= self.row <= bottom and left <= self.col <= right:
                return top, left, bottom, right
        return None

    @property
    def MergeCells(self):
        return self._area() is not None

    @property
    def MergeArea(self):
        area = self._area() or (self.row, self.col, self.row, self.col)
        return FakeArea(self.sheet, *area)


class FakeRow:
    def __init__(self, sheet, row):
        self.sheet, self.row = sheet, row

    @property
    def PageBreak(self):
        return MANUAL if self.row in self.sheet.manual else AUTO if self.row in self.sheet.auto_rows() else NONE

    @PageBreak.setter
    def PageBreak(self, value):
        if value == MANUAL:
            self.sheet.manual.add(self.row)
        else:
            self.sheet.manual.discard(self.row)


class FakeSheet:
    """每行 10pt、每页 page 行高；手动分页另起一页。merges: [(top, left, bottom, right)]"""

    def __init__(self, rows=60, page=100.0, merges=(), texts=None):
        self.rows, self.page = rows, page
        self.merges = list(merges)
        self.texts = dict(texts or {})
        self.manual: set[int] = set()
        self.PageSetup = Mock(PrintArea="", Zoom=100, FitToPagesTall=False, Order=1)
        self.UsedRange = Mock(Row=1, Column=1, Columns=Mock(Count=8))
        self.Parent = Mock()

    def breaks(self):
        result, used = [], 0.0
        for row in range(1, self.rows + 1):
            if row in self.manual and row > 1:
                result.append((row, True))
                used = 0.0
            elif used + 10 > self.page:
                result.append((row, False))
                used = 0.0
            used += 10
        return result

    def auto_rows(self):
        return {row for row, manual in self.breaks() if not manual}

    @property
    def HPageBreaks(self):
        found = self.breaks()
        pb = Mock(side_effect=lambda i: Mock(Location=Mock(Row=found[i - 1][0]),
                                              Type=MANUAL if found[i - 1][1] else AUTO))
        pb.Count = len(found)
        return pb

    @property
    def VPageBreaks(self):
        return Mock(Count=0)

    def Activate(self):
        pass

    def Cells(self, row, col):
        return FakeCell(self, row, col)

    def Rows(self, row):
        return FakeRow(self, row)

    def Range(self, a, b=None):
        cells = [(r, c) for r in range(a.row, b.row + 1) for c in range(a.col, b.col + 1)]
        merged = [FakeCell(self, r, c).MergeCells for r, c in cells]
        return Mock(MergeCells=all(merged) if all(merged) else (None if any(merged) else False))


class DetectTests(unittest.TestCase):
    def test_finds_text_block_cut_by_automatic_break_only(self):
        # 每页 10 行：自动分页在第 11 行上方；B9:F12 有文字被切开，H9:H12 空的不算
        sheet = FakeSheet(merges=[(9, 2, 12, 6), (9, 8, 12, 8), (3, 1, 4, 1)], texts={(9, 2): "見出し"})
        breaks = page_breaks.read_breaks(sheet)
        self.assertEqual(breaks["h"][0], (11, False))
        blocks = page_breaks.find_cut_blocks(sheet, breaks)
        self.assertEqual([b["addr"] for b in blocks], ["B9:F12"])
        self.assertEqual((blocks[0]["top"], blocks[0]["bottom"], blocks[0]["break_row"]), (9, 12, 11))

    def test_manual_break_inside_block_is_authors_choice(self):
        sheet = FakeSheet(merges=[(5, 1, 7, 3)], texts={(5, 1): "表題"})
        sheet.manual.add(6)
        self.assertEqual(page_breaks.find_cut_blocks(sheet, page_breaks.read_breaks(sheet)), [])

    def test_hidden_block_is_not_reported(self):
        sheet = FakeSheet(merges=[(9, 2, 12, 6)], texts={(9, 2): "見出し"})
        breaks = page_breaks.read_breaks(sheet)
        self.assertEqual(page_breaks.find_cut_blocks(sheet, breaks, lambda r, c: False), [])

    def test_unreadable_breaks_skip_detection(self):
        sheet = Mock()
        sheet.HPageBreaks.Count = Mock(side_effect=RuntimeError)  # int(Mock) 失败
        self.assertIsNone(page_breaks.read_breaks(sheet))

    def test_page_index_follows_print_order_and_refuses_mismatched_page_counts(self):
        sheet = FakeSheet()
        breaks = {"h": [(11, False), (21, False)], "v": [5]}
        self.assertEqual(page_breaks.page_index(sheet, breaks, 15, 6, 6), 4)  # 先列后行：第 2 列第 2 段
        sheet.PageSetup.Order = page_breaks.XL_OVER_THEN_DOWN
        self.assertEqual(page_breaks.page_index(sheet, breaks, 15, 6, 6), 3)
        self.assertIsNone(page_breaks.page_index(sheet, breaks, 15, 6, 5))  # 空白页被跳过时不猜

    def test_render_break_joins_two_pages_and_finds_text(self):
        doc = fitz.open()
        for text in ("上一页", "見出しテキスト"):
            page = doc.new_page(width=300, height=400)
            page.insert_text((40, 360 if text == "上一页" else 40), text, fontname="japan", fontsize=12)
        image, found_page, bbox = page_breaks.render_break(doc, 0, 1, "見出しテキスト")
        self.assertTrue(image.startswith("data:image/png;base64,"))
        self.assertEqual(found_page, 1)
        self.assertEqual(len(bbox), 4)


class PrepareTests(unittest.TestCase):
    def session_for(self, sheet):
        session = Mock()
        session.copy_of.return_value.Sheets.return_value = sheet
        session.chosen_pdf.return_value = Path(tempfile.mkdtemp()) / "chosen.pdf"
        return session

    def item(self, top, bottom, left=2):
        return {"sheet_name": "S1", "excel_row": top, "text": "見出し",
                "block": {"top": top, "bottom": bottom, "left": left, "right": 6, "break_row": 11}}

    def test_searches_upward_across_multiple_text_rows_or_vertical_merges(self):
        # 第 3 行空白，第 4、5 行都有字（或纵向合并跨过 4..6 行），内容在 6..11 行 → 分页插在第 4 行上方
        sheet = FakeSheet(merges=[(6, 2, 11, 6)], texts={(2, 2): "前段", (4, 2): "大標題", (5, 2): "小標題", (6, 2): "本文"})
        self.assertEqual(page_breaks.find_break_row(sheet, {"top": 6, "bottom": 11}), 4)
        # A4:A6 为纵向合并单元格，B4 虽无字但 A4:A6 跨过 4-5 分界线，不能在 5 上方切开 A4:A6，应再往上到 4 上方
        merged_side = FakeSheet(merges=[(4, 1, 6, 1), (6, 2, 11, 6)],
                                texts={(2, 2): "前段", (4, 1): "分類", (5, 2): "項目", (6, 2): "本文"})
        self.assertEqual(page_breaks.find_break_row(merged_side, {"top": 6, "bottom": 11}), 4)

class BatchBreakTests(unittest.TestCase):
    def setUp(self):
        # 自动修复现在只修切字；旧的自动插分页代码还留着，这里放开类型限制继续测它
        legacy = patch.object(batch.autofix, "AUTO_FIX_ERROR_TYPES", batch.autofix.FIXABLE_ERROR_TYPES)
        legacy.start()
        self.addCleanup(legacy.stop)
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for obj, name, value in (
            (storage, "DATABASE_PATH", self.root / "state.sqlite3"),
            (reporter, "RECORD_PATH", self.root / "record.md"),
            (fix_images, "IMAGES_DIR", self.root / "images"),
        ):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.wf = importlib.reload(workflow)
        self.path = str(self.root / "book.xlsx")
        Path(self.path).write_bytes(b"workbook")
        self.wf.STATE.update(folder=str(self.root), items={}, pages={},
                             batch_review={"cards": {}, "rows": {}, "ignored": []},
                             job={"kind": "batch_scan", "options": {}, "completed": [], "pending_adjustment": None})
        self.original = FakeSheet(merges=[(9, 2, 12, 6)], texts={(9, 2): "見出し"})
        self.book = Mock(ReadOnly=False)
        self.book.Sheets.return_value = self.original
        self.excel = Mock()
        self.excel.Workbooks.Open.return_value = self.book
        patch.object(batch, "normalize_workbook_view").start()
        self.refresh = patch.object(batch, "_refresh").start()
        self.addCleanup(patch.stopall)
        self.item = {"id": "b1", "seq": 1, "kind": "pagebreak", "text": "見出し", "wb_path": self.path,
                     "wb_name": "book.xlsx", "sheet_name": "S1", "cell_addr": "B9:F12", "excel_row": 9,
                     "block": {"top": 9, "bottom": 12, "left": 2, "right": 6, "break_row": 11},
                     "location_confirmed": True, "status": "pending", "image": "", "page_index": 1,
                     "bbox": [10, 10, 100, 25]}
        self.wf.STATE["items"]["b1"] = self.item

    def repair(self):
        session = Mock()
        with patch.object(batch, "_prepare_break", return_value={"after": "data:image/png;base64,YWZ0ZXI=", "pdf": None}):
            batch._repair_file(self.excel, session, self.path)
        return session

if __name__ == "__main__":
    unittest.main()
