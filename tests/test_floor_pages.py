"""PDF 工具箱「向下取整」：4.1 页挤成 4 页，真导出核对页数、字缩了多少、字有没有少，不行就改回原设置"""
import math
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import page_breaks, page_fill
from core.pdf_detector import fitz

ROW, PER_PAGE, MARGIN = 10.0, 10, 10.0


class FitSheet:
    """每行 10pt 写一行字；Zoom=100 时一页 10 行，设了 N 页高按比例缩小（不放大）"""

    def __init__(self, rows, hashes_below=None):
        self.Name = "S1"
        self.rows = rows
        self.hashes_below = hashes_below  # 缩放低于这个值时数字列变成 ###（模拟缩出来的切字）
        self.PageSetup = Mock(PrintArea="", Zoom=100, FitToPagesWide=False, FitToPagesTall=False,
                              TopMargin=MARGIN, BottomMargin=MARGIN)
        self.UsedRange = Mock(Row=1, Column=1, Rows=Mock(Count=rows), Columns=Mock(Count=8))

    def scale(self):
        setup = self.PageSetup
        if setup.Zoom is not False:
            return setup.Zoom / 100
        return min(1.0, PER_PAGE / math.ceil(self.rows / setup.FitToPagesTall))

    def export(self, _sheet, pdf_path):
        scale = self.scale()
        fit = self.PageSetup.Zoom is False and scale < 1
        per_page = math.ceil(self.rows / self.PageSetup.FitToPagesTall) if fit else max(1, int(PER_PAGE / scale + 1e-9))
        doc = fitz.open()
        for start in range(0, self.rows, per_page):
            page = doc.new_page(width=300, height=ROW * PER_PAGE + 2 * MARGIN)
            for k, row in enumerate(range(start, min(self.rows, start + per_page))):
                y = MARGIN + (k + .8) * ROW * scale
                page.insert_text((5, y), f"項目row{row}", fontsize=7 * scale, fontname="japan")
                amount = "###" if self.hashes_below and scale < self.hashes_below else f"{row * 1000}"
                page.insert_text((150, y), amount, fontsize=7 * scale)
            page.insert_text((140, ROW * PER_PAGE + 2 * MARGIN - 2), f"{len(doc)}/{math.ceil(self.rows / per_page)}",
                             fontsize=6)
        doc.save(str(pdf_path))
        doc.close()


class FloorPagesTests(unittest.TestCase):
    def setUp(self):
        self.pdf = Path(tempfile.mkdtemp()) / "S1.pdf"
        for target, name, value in ((page_fill.debug_log, "event", Mock()),
                                    (page_breaks, "read_breaks", Mock(return_value={"h": [(11, False)], "v": []}))):
            patcher = patch.object(target, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)

    def floor(self, sheet):
        sheet.export(sheet, self.pdf)
        return page_fill.floor_pages(sheet, self.pdf, sheet.export)

    def test_last_page_with_a_little_is_squeezed_into_previous(self):
        sheet = FitSheet(rows=42)  # 4.2 页
        found = self.floor(sheet)
        self.assertTrue(found["ok"], found["reason"])
        self.assertEqual((found["before"], found["after"]), (5, 4))
        self.assertAlmostEqual(found["scale"], 10 / 11, places=2)
        setup = sheet.PageSetup
        self.assertEqual((setup.Zoom, setup.FitToPagesWide, setup.FitToPagesTall), (False, 1, 4))

    def test_too_much_shrink_is_put_back(self):
        sheet = FitSheet(rows=48)  # 4.8 页：一页要放 12 行，缩到 83%
        found = self.floor(sheet)
        self.assertFalse(found["ok"])
        self.assertIn("太小", found["reason"])
        self.assertEqual(sheet.PageSetup.Zoom, 100)

    def test_lost_text_is_put_back(self):
        sheet = FitSheet(rows=42, hashes_below=0.99)
        found = self.floor(sheet)
        self.assertFalse(found["ok"])
        self.assertIn("对不上", found["reason"])
        self.assertEqual(sheet.PageSetup.Zoom, 100)

    def test_fit_settings_are_restored_when_already_fit(self):
        sheet = FitSheet(rows=48)
        sheet.PageSetup.Zoom, sheet.PageSetup.FitToPagesWide, sheet.PageSetup.FitToPagesTall = False, 1, 5
        found = self.floor(sheet)
        self.assertFalse(found["ok"])
        setup = sheet.PageSetup
        self.assertEqual((setup.Zoom, setup.FitToPagesWide, setup.FitToPagesTall), (False, 1, 5))

    def test_pages_split_by_columns_or_single_page_are_not_touched(self):
        page_breaks.read_breaks.return_value = {"h": [(11, False)], "v": [5]}
        sheet = FitSheet(rows=42)
        self.assertIn("按列", self.floor(sheet)["reason"])
        self.assertEqual(sheet.PageSetup.Zoom, 100)
        self.assertIn("1 页", self.floor(FitSheet(rows=8))["reason"])

    def test_manual_break_is_reported(self):
        page_breaks.read_breaks.return_value = {"h": [(11, True)], "v": []}
        found = self.floor(FitSheet(rows=42))
        self.assertTrue(found["ok"])
        self.assertTrue(found["manual"])


if __name__ == "__main__":
    unittest.main()
