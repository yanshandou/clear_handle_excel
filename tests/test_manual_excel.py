import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock, patch

from core import excel_ops


class ManualExcelTests(unittest.TestCase):
    def test_reuses_open_workbook_and_goes_to_target_without_saving_or_quitting(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "book.xlsx"
            path.touch()
            sheet = Mock(Visible=-1)
            sheet.Rows.return_value.RowHeight = 25.5
            workbook = Mock(FullName=str(path))
            workbook.Sheets.return_value = sheet
            workbook.Windows.Count = 1
            excel = Mock(Workbooks=[workbook])
            com = Mock()
            client = Mock(GetActiveObject=Mock(return_value=excel))
            with patch.object(excel_ops, "pythoncom", com), patch.object(excel_ops, "win32", client):
                height = excel_ops.open_workbook_for_manual_review(path, "Sheet1", "C61", 61)
            self.assertEqual(height, 25.5)
            sheet.Range.assert_called_once_with("C61")
            excel.Goto.assert_called_once_with(sheet.Range.return_value, Scroll=True)
            self.assertTrue(excel.Visible)
            workbook.Save.assert_not_called()
            workbook.Close.assert_not_called()
            excel.Quit.assert_not_called()
            client.DispatchEx.assert_not_called()
            com.CoUninitialize.assert_called_once()

    def test_starts_excel_only_when_no_active_instance_exists(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "book.xlsx"
            path.touch()
            sheet = Mock(Visible=-1)
            sheet.Rows.return_value.RowHeight = 25.5
            workbook = Mock()
            workbook.Sheets.return_value = sheet
            workbook.Windows.Count = 1
            books = MagicMock()
            books.__iter__.return_value = iter([])
            books.Open.return_value = workbook
            excel = Mock(Workbooks=books)
            client = Mock(
                GetActiveObject=Mock(side_effect=excel_ops.pythoncom.com_error(
                    -2147221021, "no active Excel", None, None
                )),
                DispatchEx=Mock(return_value=excel),
            )
            com = SimpleNamespace(
                CoInitialize=Mock(), CoUninitialize=Mock(), com_error=excel_ops.pythoncom.com_error
            )
            with patch.object(excel_ops, "pythoncom", com), patch.object(excel_ops, "win32", client):
                excel_ops.open_workbook_for_manual_review(path, "Sheet1", "Row 61", 61)
            books.Open.assert_called_once_with(str(path), UpdateLinks=0, IgnoreReadOnlyRecommended=True)
            sheet.Range.assert_called_once_with("A61")
            client.DispatchEx.assert_called_once_with("Excel.Application")
            excel.Quit.assert_not_called()

    def test_hidden_background_excel_is_not_shown_or_reused(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "book.xlsx"
            path.touch()
            hidden = Mock(Visible=False)
            visible = Mock(Visible=True)
            visible.Workbooks = MagicMock()
            visible.Workbooks.__iter__.return_value = iter([])
            workbook = visible.Workbooks.Open.return_value
            workbook.Windows.Count = 1
            workbook.Sheets.return_value.Visible = -1
            workbook.Sheets.return_value.Rows.return_value.RowHeight = 25.5
            client = Mock(GetActiveObject=Mock(return_value=hidden), DispatchEx=Mock(return_value=visible))
            with patch.object(excel_ops, "pythoncom", Mock()), patch.object(excel_ops, "win32", client):
                excel_ops.open_workbook_for_manual_review(path, "Sheet1", "C61", 61)
            self.assertFalse(hidden.Visible)
            hidden.Workbooks.Open.assert_not_called()
            hidden.Quit.assert_not_called()
            client.DispatchEx.assert_called_once_with("Excel.Application")
            visible.Workbooks.Open.assert_called_once_with(str(path), UpdateLinks=0, IgnoreReadOnlyRecommended=True)
            visible.Goto.assert_called_once()

    def _open_from_link(self, sheet_visible=-1, sheet_error=None, **kwargs):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "book.xlsx"
            path.touch()
            hidden = Mock(Visible=False)
            visible = Mock(Visible=True)
            visible.Workbooks = MagicMock()
            visible.Workbooks.__iter__.return_value = iter([])
            workbook = visible.Workbooks.Open.return_value
            workbook.Windows.Count = 1
            if sheet_error:
                workbook.Sheets.side_effect = sheet_error
            workbook.Sheets.return_value.Visible = sheet_visible
            client = Mock(GetActiveObject=Mock(return_value=hidden), DispatchEx=Mock(return_value=visible))
            with patch.object(excel_ops, "pythoncom", Mock()), patch.object(excel_ops, "win32", client):
                note = excel_ops.open_workbook_in_excel(path, **kwargs)
            hidden.Workbooks.Open.assert_not_called()
            visible.Workbooks.Open.assert_called_once_with(str(path), UpdateLinks=0, IgnoreReadOnlyRecommended=True)
            workbook.Save.assert_not_called()
            workbook.Close.assert_not_called()
            visible.Quit.assert_not_called()
            return note, visible, workbook

    def test_link_opens_file_in_separate_visible_excel_and_jumps_to_cell(self):
        note, excel, workbook = self._open_from_link(sheet_name="Sheet1", cell_address="AG72")
        self.assertEqual(note, "")
        self.assertTrue(excel.Visible)
        workbook.Sheets.return_value.Activate.assert_called_once()
        excel.Goto.assert_called_once_with(workbook.Sheets.return_value.Range.return_value, Scroll=True)
        workbook.Sheets.return_value.Range.assert_called_once_with("AG72")

    def test_link_without_sheet_only_opens_file(self):
        note, excel, workbook = self._open_from_link()
        self.assertEqual(note, "")
        workbook.Sheets.assert_not_called()
        excel.Goto.assert_not_called()

    def test_link_to_hidden_or_missing_sheet_still_opens_file(self):
        note, excel, _ = self._open_from_link(sheet_visible=0, sheet_name="Old", cell_address="A1")
        self.assertIn("隐藏", note)
        excel.Goto.assert_not_called()
        note, excel, _ = self._open_from_link(sheet_error=Exception("no sheet"), sheet_name="Gone")
        self.assertIn("找不到工作表", note)
        excel.Goto.assert_not_called()

    def test_other_com_errors_are_not_silently_replaced_by_new_excel(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "book.xlsx"
            path.touch()
            error_type = excel_ops.pythoncom.com_error
            client = Mock(GetActiveObject=Mock(side_effect=error_type(
                -2147024891, "access denied", None, None
            )))
            com = SimpleNamespace(CoInitialize=Mock(), CoUninitialize=Mock(), com_error=error_type)
            with patch.object(excel_ops, "pythoncom", com), patch.object(excel_ops, "win32", client):
                with self.assertRaises(error_type):
                    excel_ops.open_workbook_for_manual_review(path, "Sheet1", "A1", 1)
            client.DispatchEx.assert_not_called()
            com.CoUninitialize.assert_called_once()


if __name__ == "__main__":
    unittest.main()
