"""PDF 工具箱「向下取整」：4.1 页挤成 4 页，真导出核对页数、字缩了多少、字有没有少，不行就改回原设置"""
import math
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
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

    def test_big_shrink_still_floors_with_warning(self):
        # 用户 2026-11-08 定：点了取整就一定取整，字缩到 85% 以下也照改，只提醒
        sheet = FitSheet(rows=48)  # 4.8 页：一页要放 12 行，缩到 83%
        found = self.floor(sheet)
        self.assertTrue(found["ok"], found["reason"])
        self.assertIn("字缩到约 83%", found["warn"])
        self.assertEqual(sheet.PageSetup.FitToPagesTall, 4)

    def test_lost_text_still_floors_with_warning(self):
        # 用户 2026-11-08 定：缩出切字也照样取整，只提醒，后面「查看切字」会查
        sheet = FitSheet(rows=42, hashes_below=0.99)
        found = self.floor(sheet)
        self.assertTrue(found["ok"])
        self.assertIn("显示不全", found["warn"])
        self.assertEqual(sheet.PageSetup.FitToPagesTall, 4)

    def test_fit_settings_are_restored_when_already_fit(self):
        sheet = FitSheet(rows=48)
        sheet.PageSetup.Zoom, sheet.PageSetup.FitToPagesWide, sheet.PageSetup.FitToPagesTall = False, 1, 5
        sheet.export(sheet, self.pdf)

        def broken(_sheet, _pdf):  # 缩完导出出错：改回原来的设置
            raise RuntimeError("printer")
        found = page_fill.floor_pages(sheet, self.pdf, broken)
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


class OnePageWideTests(unittest.TestCase):
    """调整页面（用户 2026-11-07）：导出前设成宽一页、高自动；末页只多一点的建议取整"""

    def sheet(self, **setup):
        values = dict(Zoom=100, FitToPagesWide=False, FitToPagesTall=False, Orientation=1, TopMargin=0, BottomMargin=0)
        values.update(setup)
        return SimpleNamespace(Name="表", PageSetup=SimpleNamespace(**values))

    def test_percent_zoom_becomes_one_page_wide_auto_tall(self):
        sheet = self.sheet(Zoom=85)
        self.assertEqual(page_fill.one_page_wide(sheet), {"before": "缩放 85%"})
        self.assertEqual((sheet.PageSetup.Zoom, sheet.PageSetup.FitToPagesWide, sheet.PageSetup.FitToPagesTall),
                         (False, 1, False))

    def test_fit_two_pages_wide_becomes_one(self):
        sheet = self.sheet(Zoom=False, FitToPagesWide=2, FitToPagesTall=5)
        self.assertEqual(page_fill.one_page_wide(sheet), {"before": "缩放到 2 页宽、5 页高"})
        self.assertEqual(sheet.PageSetup.FitToPagesTall, False)

    def test_already_one_page_wide_keeps_floored_tall(self):
        sheet = self.sheet(Zoom=False, FitToPagesWide=1, FitToPagesTall=3)
        self.assertIsNone(page_fill.one_page_wide(sheet))
        self.assertEqual(sheet.PageSetup.FitToPagesTall, 3)

    def test_unreadable_setup_is_skipped(self):
        self.assertIsNone(page_fill.one_page_wide(SimpleNamespace(Name="表")))

    def test_layout_reads_orientation_and_fixed_tall(self):
        self.assertEqual(page_fill.sheet_layout(self.sheet(Zoom=False, FitToPagesWide=1, FitToPagesTall=3, Orientation=2)),
                         {"orient": 2, "tall": 3, "manual": 0, "margins": (0.0, 0.0)})
        self.assertEqual(page_fill.sheet_layout(self.sheet())["tall"], None)

    def test_tail_hint_suggests_floor_when_last_page_is_short(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "a.pdf"
            doc = fitz.open()
            for bottom in (800, 800, 150):
                page = doc.new_page(width=595, height=842)
                page.insert_text((50, 50), "上")
                page.insert_text((50, bottom), "下")
            doc.save(pdf)
            hint = page_fill.tail_hint(pdf, (0.0, 0.0))
            self.assertEqual(hint["pages"], 3)
            self.assertTrue(hint["floor"])
            self.assertLess(hint["tail"], 40)
            # 末页印了 38% 已经够了，不建议取整（用户 2026-11-08：不到 20% 才标黄）
            mid = Path(tmp) / "c.pdf"
            doc = fitz.open()
            for bottom in (800, 800, 330):
                page = doc.new_page(width=595, height=842)
                page.insert_text((50, 50), "上")
                page.insert_text((50, bottom), "下")
            doc.save(mid)
            hint = page_fill.tail_hint(mid, (0.0, 0.0))
            self.assertGreater(hint["tail"], 30)
            self.assertFalse(hint["floor"])
            one = Path(tmp) / "b.pdf"
            doc = fitz.open()
            doc.new_page(width=595, height=842)
            doc.save(one)
            self.assertEqual(page_fill.tail_hint(one, (0.0, 0.0)), {"pages": 1, "tail": None, "floor": False})
