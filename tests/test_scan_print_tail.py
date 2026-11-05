"""只读 Excel 分页估算；扫描只允许打印页数设置和颜色标记。"""
import copy
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from core import page_fill, workflow


class ExcelTailTests(unittest.TestCase):
    def sheet(self, pages=3, fill=.30, *, zoom=100, titles=0):
        body = 297 * 72 / 25.4 - 72
        setup = SimpleNamespace(Zoom=zoom, FitToPagesWide=1, FitToPagesTall=False,
                                PaperSize=9, Orientation=1, PrintHeadings=False,
                                TopMargin=36, BottomMargin=36, LeftMargin=36, RightMargin=36,
                                PrintTitleRows="$1:$2" if titles else "")
        sheet = Mock(Name="S1", PageSetup=setup)
        sheet.Cells.side_effect = lambda row, col: (row, col)
        sheet.Range.side_effect = lambda first, last=None: SimpleNamespace(
            Height=titles if isinstance(first, str) else (body - titles) * fill,
            Width=(210 * 72 / 25.4 - 72) * 2)
        layout = {"top": 1, "left": 1, "bottom": pages * 10, "right": 8,
                  "h": [(10 * i + 1, False) for i in range(1, pages)]}
        return sheet, layout

    def compact(self, sheet, layout):
        with patch.object(page_fill, "_layout", return_value=layout):
            return page_fill.compact_scan_tail(sheet)

    def test_three_and_six_pages_reduce_by_one_without_touching_cells(self):
        for count, fill in ((3, .3), (6, .2), (2, .1)):
            with self.subTest(count=count):
                sheet, layout = self.sheet(count, fill)
                change = self.compact(sheet, layout)
                self.assertAlmostEqual(change["tail_fill"], fill)
                self.assertEqual(change["after"], {"Zoom": False, "FitToPagesWide": 1,
                                                    "FitToPagesTall": count - 1})
                self.assertEqual(sheet.PageSetup.FitToPagesTall, count - 1)
                sheet.ExportAsFixedFormat.assert_not_called()
                sheet.Rows.assert_not_called()
                sheet.Columns.assert_not_called()

    def test_forty_percent_is_not_reduced_and_single_page_is_kept(self):
        for pages, fill in ((3, .40), (6, .8), (1, .2)):
            sheet, layout = self.sheet(pages, fill)
            before = copy.copy(sheet.PageSetup)
            self.assertIsNone(self.compact(sheet, layout))
            self.assertEqual(vars(before), vars(sheet.PageSetup))

    def test_percent_zoom_and_fit_width_scale_row_height(self):
        for zoom in (50, False):
            sheet, layout = self.sheet(fill=.6, zoom=zoom)
            change = self.compact(sheet, layout)
            self.assertAlmostEqual(change["tail_fill"], .3)

    def test_repeated_titles_are_removed_from_available_body_height(self):
        sheet, layout = self.sheet(fill=.35, titles=60)
        change = self.compact(sheet, layout)
        self.assertAlmostEqual(change["tail_fill"], .35)

    def test_manual_break_and_unknown_layout_are_unchanged(self):
        sheet, layout = self.sheet()
        layout["h"][0] = (11, True)
        self.assertIsNone(self.compact(sheet, layout))
        self.assertIsNone(self.compact(sheet, None))
        self.assertEqual(sheet.PageSetup.Zoom, 100)

    def test_unknown_paper_and_invalid_scale_are_unchanged(self):
        for key, value in (("PaperSize", 256), ("Zoom", None), ("PrintHeadings", True)):
            sheet, layout = self.sheet()
            setattr(sheet.PageSetup, key, value)
            before = copy.copy(sheet.PageSetup)
            self.assertIsNone(self.compact(sheet, layout))
            self.assertEqual(vars(before), vars(sheet.PageSetup))


class SavePrintSettingsTests(unittest.TestCase):
    def test_only_print_settings_are_written_and_save_failure_does_not_succeed(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "book.xlsx"
            path.write_bytes(b"original")
            before = {"Zoom": 100, "FitToPagesWide": 1, "FitToPagesTall": False}
            after = {"Zoom": False, "FitToPagesWide": 1, "FitToPagesTall": 2}
            setup = SimpleNamespace(**before)
            sheet = Mock(PageSetup=setup)
            book = Mock(ReadOnly=False)
            book.Sheets.return_value = sheet
            with patch.object(workflow, "open_workbook_checked", return_value=book):
                workflow._save_scan_print_settings(Mock(), path, workflow._file_version(path),
                                                   {"S1": {"before": before, "after": after}})
            self.assertEqual(vars(setup), after)
            book.Save.assert_called_once()
            book.Close.assert_called_once_with(SaveChanges=False)
            sheet.Rows.assert_not_called()
            sheet.Columns.assert_not_called()
            sheet.Activate.assert_not_called()

    def test_changed_workbook_is_not_opened_for_writing(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "book.xlsx"
            path.write_bytes(b"original")
            before = workflow._file_version(path)
            path.write_bytes(b"changed workbook")
            with patch.object(workflow, "open_workbook_checked") as opened:
                with self.assertRaisesRegex(ValueError, "文件已修改"):
                    workflow._save_scan_print_settings(Mock(), path, before, {"S1": {}})
            opened.assert_not_called()
