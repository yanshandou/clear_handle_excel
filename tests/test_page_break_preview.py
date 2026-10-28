import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from core import excel_ops, workflow
from core.common import XL_PAGE_BREAK_PREVIEW, XL_WORKSHEET


class Sheets(list):
    def __call__(self, index):
        return self[index - 1]

    @property
    def Count(self):
        return len(self)


class Sheet:
    def __init__(self, excel, name, visible=-1, kind=XL_WORKSHEET):
        self.excel, self.Name, self.Visible, self.Type = excel, name, visible, kind

    def Activate(self):
        self.excel.ActiveWindow.ActiveSheet = self

    def Range(self, addr):
        return (self.Name, addr)


class Window:
    def __init__(self, excel, sheet, views, visible=True):
        self.excel, self.ActiveSheet, self.Visible = excel, sheet, visible
        self.views = dict(views)
        self.zooms = {name: 85 for name in views}

    def Activate(self):
        if not self.Visible:
            raise RuntimeError("隐藏窗口不能激活")
        self.excel.ActiveWindow = self

    @property
    def View(self):
        return self.views[self.ActiveSheet.Name]

    @View.setter
    def View(self, value):
        self.views[self.ActiveSheet.Name] = value
        self.Zoom = 60  # Excel 切换视图时可能采用另一套缩放。

    @property
    def Zoom(self):
        return self.zooms[self.ActiveSheet.Name]

    @Zoom.setter
    def Zoom(self, value):
        self.zooms[self.ActiveSheet.Name] = value


class PageBreakPreviewTests(unittest.TestCase):
    def setup_book(self):
        excel = SimpleNamespace(Goto=Mock())
        sheets = Sheets([Sheet(excel, "First"), Sheet(excel, "Second"),
                         Sheet(excel, "Hidden", visible=0), Sheet(excel, "Chart", kind=-4109)])
        windows = Sheets([Window(excel, sheets[1], {s.Name: 1 for s in sheets}),
                          Window(excel, sheets[0], {s.Name: 3 for s in sheets})])
        excel.ActiveWindow = windows[0]
        book = SimpleNamespace(Windows=windows, Worksheets=sheets[:3], Sheets=sheets)
        return excel, book

    def test_all_windows_keep_preview_and_restore_active_sheet_and_zoom(self):
        excel, book = self.setup_book()
        self.assertTrue(excel_ops.enable_page_break_preview(excel, book))
        for window, original_view in zip(book.Windows, (1, 3)):
            self.assertEqual(window.views, {"First": 2, "Second": 2, "Hidden": original_view,
                                           "Chart": original_view})
            self.assertEqual(set(window.zooms.values()), {85})
        self.assertIs(excel.ActiveWindow, book.Windows[0])
        self.assertIs(book.Windows[0].ActiveSheet, book.Sheets[1])
        self.assertIs(book.Windows[1].ActiveSheet, book.Sheets[0])
        excel.Goto.assert_not_called()
        self.assertFalse(excel_ops.enable_page_break_preview(excel, book))

    def test_hidden_workbook_window_is_set_and_hidden_again(self):
        excel, book = self.setup_book()
        book.Windows[1].Visible = False
        self.assertTrue(excel_ops.enable_page_break_preview(excel, book))
        self.assertEqual(book.Windows[1].views["First"], XL_PAGE_BREAK_PREVIEW)
        self.assertFalse(book.Windows[1].Visible)
        self.assertIs(excel.ActiveWindow, book.Windows[0])

    def test_normalization_retains_preview_with_a1_and_100_percent(self):
        excel, book = self.setup_book()
        excel_ops.normalize_workbook_view(excel, book)
        self.assertEqual(book.Windows[0].views["First"], XL_PAGE_BREAK_PREVIEW)
        self.assertEqual(book.Windows[0].views["Second"], XL_PAGE_BREAK_PREVIEW)
        self.assertEqual(book.Windows[0].views["Hidden"], 1)
        self.assertEqual(book.Windows[0].views["Chart"], 1)
        self.assertEqual(book.Windows[0].zooms["First"], 100)
        self.assertEqual(book.Windows[0].zooms["Second"], 100)
        self.assertIs(book.Windows[0].ActiveSheet, book.Sheets[0])
        self.assertEqual(excel.Goto.call_count, 2)

    def test_missing_window_is_reported(self):
        with self.assertRaisesRegex(ValueError, "没有可用窗口"):
            excel_ops.enable_page_break_preview(Mock(), SimpleNamespace(Windows=Sheets()))


class PreviewSaveTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "book.xlsx"
        self.path.write_bytes(b"original workbook")
        self.before = workflow._file_version(self.path)
        self.book = Mock(ReadOnly=False)
        self.excel = Mock()
        self.excel.Workbooks.Open.return_value = self.book
        self.preview = patch.object(workflow, "enable_page_break_preview", return_value=True).start()
        self.followed = patch.object(workflow.print_layout, "followed_save").start()
        self.addCleanup(patch.stopall)

    def test_opens_fresh_original_and_saves_view(self):
        self.book.Save.side_effect = lambda: self.path.write_bytes(b"workbook with saved view")
        after = workflow._save_page_break_preview(self.excel, self.path, self.before)
        self.assertNotEqual(after, self.before)
        self.excel.Workbooks.Open.assert_called_once_with(str(self.path.resolve()), ReadOnly=False, UpdateLinks=0)
        self.preview.assert_called_once_with(self.excel, self.book)
        self.book.Save.assert_called_once()
        self.book.Close.assert_called_once_with(SaveChanges=False)
        self.followed.assert_called_once()

    def test_already_enabled_view_does_not_save_again(self):
        self.preview.return_value = False
        self.assertEqual(workflow._save_page_break_preview(self.excel, self.path, self.before), self.before)
        self.book.Save.assert_not_called()

    def test_read_only_original_is_not_silently_completed(self):
        self.book.ReadOnly = True
        with self.assertRaisesRegex(ValueError, "只读"):
            workflow._save_page_break_preview(self.excel, self.path, self.before)
        self.book.Save.assert_not_called()
        self.book.Close.assert_called_once_with(SaveChanges=False)

    def test_changed_file_is_not_overwritten(self):
        self.path.write_bytes(b"user changed workbook")
        with self.assertRaisesRegex(ValueError, "文件已修改"):
            workflow._save_page_break_preview(self.excel, self.path, self.before)
        self.excel.Workbooks.Open.assert_not_called()

    def test_change_while_setting_view_is_not_overwritten(self):
        self.preview.side_effect = lambda *_: self.path.write_bytes(b"user edited during view setup")
        with self.assertRaisesRegex(ValueError, "文件已修改"):
            workflow._save_page_break_preview(self.excel, self.path, self.before)
        self.book.Save.assert_not_called()
        self.book.Close.assert_called_once_with(SaveChanges=False)

    def test_xls_open_metadata_changes_are_accepted(self):
        path = self.path.with_suffix(".XLS")
        path.write_bytes(b"binary workbook before open")
        before = workflow._file_version(path)
        def opened(*_args, **_kwargs):
            path.write_bytes(b"binary workbook with Excel open metadata")
            return self.book
        self.excel.Workbooks.Open.side_effect = opened
        self.preview.return_value = True
        workflow._save_page_break_preview(self.excel, path, before)
        self.book.Save.assert_called_once()

    def test_xlsx_changes_while_opening_are_rejected(self):
        def opened(*_args, **_kwargs):
            self.path.write_bytes(b"user edited while opening")
            return self.book
        self.excel.Workbooks.Open.side_effect = opened
        with self.assertRaisesRegex(ValueError, "文件已修改"):
            workflow._save_page_break_preview(self.excel, self.path, self.before)
        self.book.Save.assert_not_called()

    def test_setting_failure_closes_without_save(self):
        self.preview.side_effect = RuntimeError("printer unavailable")
        with self.assertRaisesRegex(RuntimeError, "printer unavailable"):
            workflow._save_page_break_preview(self.excel, self.path, self.before)
        self.book.Save.assert_not_called()
        self.book.Close.assert_called_once_with(SaveChanges=False)

    def test_changed_file_is_rejected_even_if_preview_was_already_enabled(self):
        def changed(*_):
            self.path.write_bytes(b"user edited while checking view")
            return False
        self.preview.side_effect = changed
        with self.assertRaisesRegex(ValueError, "文件已修改"):
            workflow._save_page_break_preview(self.excel, self.path, self.before)
        self.book.Save.assert_not_called()


class PreviewCacheTests(unittest.TestCase):
    def test_view_only_save_reuses_current_pdf_and_copy_but_does_not_update_stale_cache(self):
        from core.autofix import _ExcelSession
        session = _ExcelSession()
        before, after, stale = {"size": 10}, {"size": 20}, {"size": 5}
        session.pdfs = {("book", "S1"): (before, "s1.pdf", (1, 2)),
                        ("book", "S2"): (stale, "s2.pdf", (3, 4)),
                        ("other", "S1"): (before, "other.pdf", (1, 2))}
        copy = Mock()
        session.copies["book"] = (before, copy)
        session.view_saved("book", before, after)
        self.assertEqual(session.pdfs[("book", "S1")], (after, "s1.pdf", (1, 2)))
        self.assertEqual(session.pdfs[("book", "S2")], (stale, "s2.pdf", (3, 4)))
        self.assertEqual(session.pdfs[("other", "S1")][0], before)
        self.assertEqual(session.copies["book"], (after, copy))
        copy.Close.assert_not_called()


if __name__ == "__main__":
    unittest.main()
