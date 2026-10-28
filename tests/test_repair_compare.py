import queue
import tempfile
import threading
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock, patch

from core import excel_ops, link_repair, mark_finder, hidden_repair, repair_compare as compare


class RepairComparisonTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.current, self.backup = self.root / "current" / "book.xlsx", self.root / "backup" / "book.xlsx"
        for path in (self.current, self.backup):
            path.parent.mkdir()
            path.touch()
        self.target = {"current":self.current, "backup":self.backup, "sheet":"NewMain", "old_sheet":"Main", "cell":"B7"}
        self.data = {"folder":str(self.current.parent), "files":[{"name":"book.xlsx", "path":str(self.current), "old":str(self.backup),
            "cells":[{"sheet":"NewMain", "old_sheet":"Main", "addr":"B7:D7", "anchor":"B7"}]}]}

    def test_server_pair_resolves_renamed_sheets_and_merged_anchors_for_both_features(self):
        for kind, module in (("links", link_repair), ("colors", mark_finder), ("hidden", hidden_repair)):
            with patch.object(module, "status", return_value={"applied":self.data}):
                for sheet, cell in (("", ""), ("NewMain", "B7"), ("NewMain", "$B$7:$D$7")):
                    self.assertEqual(compare.comparison_target(kind, str(self.current.parent), "book.xlsx", sheet, cell), self.target)

    def test_only_current_server_results_and_listed_cells_can_be_opened(self):
        with patch.object(link_repair, "status", return_value={"applied":self.data}):
            for args in (("unknown", str(self.current.parent), "book.xlsx"),
                         ("links", str(self.backup.parent), "book.xlsx"),
                         ("links", str(self.current.parent), "outside.xlsx"),
                         ("links", str(self.current.parent), "book.xlsx", "Main", "A1")):
                with self.subTest(args=args), self.assertRaises(ValueError):
                    compare.comparison_target(*args)
        with patch.object(link_repair, "status", return_value={"running":True, "applied":self.data}):
            with self.assertRaisesRegex(ValueError, "结束"):
                compare.comparison_target("links", str(self.current.parent), "book.xlsx")

    def excel(self, path=None, readonly=False):
        books = []
        collection = MagicMock()
        collection.__iter__.side_effect = lambda: iter(books)
        excel = Mock(Workbooks=collection, Visible=True, DisplayAlerts=True)

        def open_book(name, **options):
            workbook = Mock(FullName=name, ReadOnly=options.get("ReadOnly", False), Saved=True)
            workbook.Windows.Count = 1
            workbook.Sheets.return_value.Visible = -1
            books.append(workbook)
            return workbook

        collection.Open.side_effect = open_book
        if path:
            open_book(str(path), ReadOnly=readonly)
            collection.Open.reset_mock()
        return excel, books

    def test_same_name_files_use_distinct_instances_readonly_backup_and_reuse_on_next_click(self):
        current, current_books = self.excel(self.current)
        backup, backup_books = self.excel()
        viewer = compare.ComparisonWindows()
        with patch.object(excel_ops, "_visible_excel", return_value=current), \
                patch.object(excel_ops.win32, "DispatchEx", return_value=backup) as dispatch, \
                patch.object(excel_ops, "focus_workbook_in_excel", return_value="") as focus, \
                patch.object(excel_ops, "_goto_fit_width"), patch.object(excel_ops, "_bring_to_front"), \
                patch.object(compare, "_tile_windows", return_value="") as tile:
            first = viewer.open(self.target)
            second = viewer.open({**self.target, "cell":"B8"})
        self.assertEqual(first["opened"], ["current", "backup"])
        self.assertEqual(second["note"], "")
        dispatch.assert_called_once_with("Excel.Application")
        current.Workbooks.Open.assert_not_called()
        backup.Workbooks.Open.assert_called_once_with(str(self.backup), UpdateLinks=0, ReadOnly=True,
                                                      IgnoreReadOnlyRecommended=True, AddToMru=False)
        self.assertIn("只读", backup_books[0].Windows(1).Caption)
        self.assertEqual(focus.call_args_list[0].args[2:], ("NewMain", "B7"))
        self.assertEqual(focus.call_args_list[1].args[2:], ("Main", "B7"))
        self.assertEqual(tile.call_count, 2)
        for excel, book in ((current, current_books[0]), (backup, backup_books[0])):
            book.Save.assert_not_called()
            book.Close.assert_not_called()
            excel.Quit.assert_not_called()

    def test_conflicting_current_filename_gets_its_own_instance_without_closing_user_file(self):
        other, books = self.excel(self.backup)
        dedicated, _ = self.excel()
        viewer = compare.ComparisonWindows()
        with patch.object(excel_ops, "_visible_excel", return_value=other), \
                patch.object(excel_ops.win32, "DispatchEx", return_value=dedicated):
            excel, book = viewer._open(self.current, False)
        self.assertIs(excel, dedicated)
        other.Workbooks.Open.assert_not_called()
        other.Quit.assert_not_called()
        books[0].Close.assert_not_called()

    def test_hidden_sheets_and_partial_open_failures_are_reported_without_saving(self):
        viewer = compare.ComparisonWindows()
        excel, books = self.excel(self.current)
        with patch.object(viewer, "_open", side_effect=[(excel, books[0]), RuntimeError("无法读取备份")]), \
                patch.object(excel_ops, "focus_workbook_in_excel", return_value="工作表已隐藏"), \
                patch.object(excel_ops, "_bring_to_front"):
            result = viewer.open(self.target)
        self.assertEqual(result["opened"], ["current"])
        self.assertIn("隐藏", result["note"])
        self.assertIn("无法读取备份", result["note"])
        books[0].Save.assert_not_called()

    def test_worker_reuses_com_apartment_across_http_requests(self):
        task_queue = queue.Queue()
        pending = {}
        com = Mock()
        viewer = Mock()
        calls = []

        def opened(target):
            calls.append(threading.get_ident())
            return {"opened":["current", "backup"], "note":""}

        viewer.open.side_effect = opened
        with patch.object(compare, "_QUEUE", task_queue), patch.object(compare, "_THREAD", None), \
                patch.object(compare, "_PENDING", pending), patch.object(compare, "ComparisonWindows", return_value=viewer), \
                patch.object(excel_ops, "pythoncom", com):
            result = compare.open_comparison(self.target)
            self.assertEqual(result["opened"], ["current", "backup"])
            compare.open_comparison(self.target)
            task_queue.put(None)
            compare._THREAD.join(2)
        self.assertEqual(len(set(calls)), 1)
        self.assertNotEqual(calls[0], threading.get_ident())
        com.CoInitialize.assert_called_once()
        com.CoUninitialize.assert_called_once()

    def test_http_comparison_uses_server_targets_and_obeys_write_locks(self):
        import web_clip_fixer as web

        handler = object.__new__(web.WebHandler)
        handler.path = "/api/excel/compare"
        handler._send_json = Mock()
        state = {"busy":False, "job":None}
        payload = {"kind":"links", "folder":str(self.current.parent), "file":"book.xlsx", "sheet":"NewMain", "cell":"B7"}
        with patch.object(web, "STATE", state), patch.object(compare, "comparison_target", return_value=self.target) as target, \
                patch.object(compare, "open_comparison", return_value={"opened":["current","backup"], "note":""}) as opened, \
                patch.object(web.autofix, "wait_for_writes"):
            handler._handle_post({**payload, "backup":"untrusted.xlsx"})
            target.assert_called_once_with("links", str(self.current.parent), "book.xlsx", "NewMain", "B7")
            opened.assert_called_once_with(self.target)
            self.assertTrue(handler._send_json.call_args.args[0]["ok"])
            opened.reset_mock()
            with self.assertRaises(ValueError):
                handler._handle_post({"kind":[], "file":"book.xlsx"})
            state.update(busy=True, maintenance="color_repair")
            with self.assertRaisesRegex(ValueError, "颜色"):
                handler._handle_post(payload)
            opened.assert_not_called()


if __name__ == "__main__":
    unittest.main()
