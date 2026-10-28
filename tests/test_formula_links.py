import copy
from pathlib import Path
from types import SimpleNamespace as NS
import unittest
from unittest.mock import Mock, patch

from core.backup_compare import pair_files
from core import formula_links as links, link_repair


class DependencyTests(unittest.TestCase):
    def setUp(self):
        sheets = [NS(Name=name, ListObjects=[]) for name in ("Main", "Old", "Middle", "Last")]
        sheets[1].ListObjects = [NS(Name="DataTable")]
        book = NS(Name="book.xlsx", Sheets=sheets, Worksheets=sheets, Names=[
            NS(Name="Original", RefersTo="=Old!$A$1"),
            NS(Name="Alias", RefersTo="=Original*2"),
            NS(Name="'Main'!Local", RefersTo="=Old!B2"),
            NS(Name="Circular", RefersTo="=Circular"),
            NS(Name="Dynamic", RefersTo='=INDIRECT("Old!A1")'),
        ])
        self.index = links.Dependencies(book, ["Old"])

    def test_direct_names_tables_dynamic_and_three_dimensional_references(self):
        for expression in ("=Old!A1*2", "=SUM('Old:Last'!A1)", "=Alias", "=Local", "=Dynamic",
                           "=SUM(DataTable[Amount])", '=INDIRECT(A1&"!B2")', '=[book.xlsx]Old!A1'):
            with self.subTest(expression=expression):
                self.assertTrue(self.index.depends(expression, "Main"))
        middle = copy.copy(self.index)
        middle.deleting = {"middle"}
        self.assertTrue(middle.depends("=SUM(Old:Last!A1)", "Main"))

    def test_strings_other_workbooks_and_circular_names_are_not_dependencies(self):
        for expression in ('="Old!A1"', "=Main!A1+OriginalText", "='[other.xlsx]Old'!A1", "=Circular", "=Local"):
            with self.subTest(expression=expression):
                scope = "Last" if expression == "=Local" else "Main"
                self.assertFalse(self.index.depends(expression, scope))

    def test_qualified_token_does_not_change_scope_of_later_local_name(self):
        self.assertTrue(self.index.depends("=Last!A1+Local", "Main"))

    def test_ref_string_is_not_an_error_and_saved_damage_is_normalized(self):
        self.assertFalse(links.broken_formula('="#REF!A1"'))
        self.assertTrue(links.broken_formula('=IFERROR(#REF!A1,0)'))
        self.assertEqual(link_repair._formula_identity('=IFERROR(#REF!A1,0)'), '=IFERROR(#REF!,0)')
        self.assertEqual(link_repair._formula_identity('=#REF!+A1'), '=#REF!+A1')
        self.assertEqual(link_repair._formula_identity('="#REF!A1"'), '="#REF!A1"')


class BackupPairingTests(unittest.TestCase):
    def test_current_001_matches_001_in_larger_backup_even_with_duplicate_names(self):
        root = Path("backup")
        paths = [root / name / "book.xlsx" for name in ("001", "002", "003")]
        self.assertEqual(pair_files(root, paths, ["book.xlsx"], Path("current/001")), {"book.xlsx": paths[0]})

    def test_ambiguous_fallback_is_not_used_and_exact_path_wins(self):
        root = Path("backup")
        paths = [root / "001" / "book.xlsx", root / "002" / "book.xlsx"]
        self.assertIsNone(pair_files(root, paths, ["book.xlsx"])["book.xlsx"])
        self.assertEqual(pair_files(root, paths, ["001/book.xlsx"])["001/book.xlsx"], paths[0])

    def test_matching_001_branch_precedes_unrelated_root_file(self):
        root = Path("backup")
        paths = [root / "book.xlsx", root / "001" / "book.xlsx"]
        self.assertEqual(pair_files(root, paths, ["book.xlsx"], Path("current/001"))["book.xlsx"], paths[1])


class RepairTaskTests(unittest.TestCase):
    def setUp(self):
        self.state = {"busy": False, "job": None}
        patcher = patch.object(link_repair.wf, "STATE", self.state)
        patcher.start()
        self.addCleanup(patcher.stop)
        patcher = patch.object(link_repair, "_STATUS", {"running": False})
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_busy_and_unsettled_writes_prevent_new_repair_task(self):
        self.state["busy"] = True
        with self.assertRaises(ValueError):
            link_repair._launch("scan", Path("current"), [])
        self.state.update(busy=False, job={"pending_adjustment": {"path": "book.xlsx"}})
        with self.assertRaises(ValueError):
            link_repair._launch("apply", Path("current"), [])
        self.assertFalse(self.state["busy"])

    def test_thread_start_failure_releases_reservation(self):
        with patch("core.mark_finder.status", return_value={"running": False}), \
                patch.object(link_repair.threading, "Thread") as thread:
            thread.return_value.start.side_effect = RuntimeError("no thread")
            with self.assertRaises(RuntimeError):
                link_repair._launch("scan", Path("current"), [])
        self.assertFalse(self.state["busy"])
        self.assertNotIn("maintenance", self.state)
        self.assertFalse(link_repair.status()["running"])

    def test_stale_preview_and_open_target_are_never_written(self):
        file = {"path": "current.xlsx", "old": "old.xlsx", "fingerprint": "a", "old_fingerprint": "b"}
        excel = Mock()
        with patch.object(link_repair, "fingerprint", return_value="changed"):
            with self.assertRaises(ValueError):
                link_repair.apply_file(excel, file)
        with patch.object(link_repair, "fingerprint", side_effect=["a", "b"]), \
                patch.object(link_repair, "workbook_locked", return_value=True):
            with self.assertRaises(ValueError):
                link_repair.apply_file(excel, file)
        excel.Workbooks.Open.assert_not_called()

    def test_numeric_equivalence_does_not_confuse_booleans(self):
        self.assertTrue(link_repair._same(10, 10.0))
        self.assertFalse(link_repair._same(True, 1))
        self.assertTrue(link_repair._same(((10,),), [[10.0]]))

    def test_current_file_cannot_be_used_as_its_own_backup(self):
        excel = Mock()
        with self.assertRaises(ValueError):
            link_repair.preview_file(excel, {"path": "book.xlsx"}, Path("book.xlsx"))
        excel.Workbooks.Open.assert_not_called()

    def test_second_write_failure_closes_without_saving_any_change(self):
        file = {"path": "current.xlsx", "old": "old.xlsx", "fingerprint": "a", "old_fingerprint": "b",
                "cells": [{"sheet": "Main", "addr": addr, "anchor": addr, "formula": "=#REF!",
                           "formulas": "=#REF!", "kind": "cell", "value": 10} for addr in ("A1", "A2")]}
        book = Mock(ReadOnly=False)
        book.Worksheets.return_value.Range.return_value.Formula = "=#REF!"
        with patch.object(link_repair, "fingerprint", side_effect=["a", "b"]), \
                patch.object(link_repair, "workbook_locked", return_value=False), \
                patch.object(link_repair, "open_workbook_checked", return_value=book), \
                patch.object(links, "error_positions", return_value=set()), \
                patch.object(links, "write_values", side_effect=[None, RuntimeError("write failed")]):
            with self.assertRaisesRegex(RuntimeError, "write failed"):
                link_repair.apply_file(Mock(), file)
        book.Save.assert_not_called()
        book.Close.assert_called_once_with(SaveChanges=False)


if __name__ == "__main__":
    unittest.main()
