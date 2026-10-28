"""隐藏恢复：备份依据、列位置、文件版本、整文件保存和后台互斥。"""
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from core import hidden_repair as repair, workflow as wf


def run(first, last, hidden=False, width=8.43):
    return {"first": first, "last": last, "hidden": hidden, "width": None if hidden else width}


def sheet(name="Main", **changes):
    return {"name": name, "code": name, "visible": -1, "words": ["same"], "signature": "same",
            "populated": 1, "content_error": "", "columns_error": "", "protected": False,
            "columns": [run(1, 10)], **changes}


def book(*sheets, **changes):
    return {"sheets": list(sheets), "visible_count": sum(item["visible"] == -1 for item in sheets),
            "structure_protected": False, **changes}


class ColumnRange:
    def __init__(self, sheet, first, last):
        self.sheet, self.first, self.last = sheet, first, last

    @property
    def EntireColumn(self):
        return self

    @property
    def Hidden(self):
        values = {self.sheet.hidden.get(index, False) for index in range(self.first, self.last + 1)}
        return next(iter(values)) if len(values) == 1 else None

    @Hidden.setter
    def Hidden(self, value):
        for index in range(self.first, self.last + 1):
            self.sheet.hidden[index] = value

    @property
    def ColumnWidth(self):
        values = {0 if self.sheet.hidden.get(index) else self.sheet.widths.get(index, 8.43)
                  for index in range(self.first, self.last + 1)}
        return next(iter(values)) if len(values) == 1 else None

    @ColumnWidth.setter
    def ColumnWidth(self, value):
        for index in range(self.first, self.last + 1):
            self.sheet.widths[index] = value


class ColumnSheet:
    def __init__(self, name="Main"):
        self.Name, self.Visible = name, -1
        class Columns:
            Count = 16384

            def __call__(columns, index):
                return ColumnRange(self, index, index)

        self.Columns = Columns()
        self.hidden, self.widths, self.calls = {}, {}, 0

    def Range(self, address):
        self.calls += 1

        def index(letters):
            value = 0
            for letter in letters:
                value = value * 26 + ord(letter) - 64
            return value

        first, last = map(index, address.split(":"))
        return ColumnRange(self, first, last)


class HiddenPlanTests(unittest.TestCase):
    def plan(self, current, old, **options):
        return repair.plan_changes(current, old, repair.options_for(options))

    def test_hidden_and_very_hidden_sheets_columns_and_narrow_widths_are_distinct(self):
        old = book(sheet(), sheet("Aux", visible=0), sheet("Deep", visible=2))
        old["sheets"][0]["columns"] = [run(1, 1), run(2, 3, True), run(4, 4, width=0.25), run(5, 10, width=12)]
        current = book(sheet(columns=[run(1, 10, width=20)]), sheet("Aux"), sheet("Deep"))
        plan = self.plan(current, old)
        self.assertEqual([(item["sheet"], item["kind"], item["addr"], item["value"]) for item in plan],
                         [("Main", "columns", "B:C", None), ("Main", "narrow", "D:D", 0.25),
                          ("Aux", "sheet", "A1", 0), ("Deep", "sheet", "A1", 2)])
        self.assertTrue(all(item["action"] == "restore" for item in plan))

    def test_no_unhiding_or_widening_existing_narrow_columns_and_normal_widths_ignored(self):
        current = book(sheet(columns=[run(1, 2, True), run(3, 3, width=0.1), run(4, 10, width=30)]))
        old = book(sheet(columns=[run(1, 1, width=0.25), run(2, 2, True), run(3, 3, width=0.5), run(4, 10, width=2)]))
        self.assertEqual(self.plan(current, old), [])

    def test_content_formula_position_changes_or_failed_reads_are_never_hidden(self):
        for changes in ({"signature": "shifted"}, {"content_error": "使用区域超过 100 万格"}):
            with self.subTest(changes=changes):
                plan = self.plan(book(sheet(**changes)), book(sheet(columns=[run(1, 10, True)])))
                self.assertEqual([item["action"] for item in plan], ["skip"])
                self.assertTrue(plan[0]["reason"])

    def test_options_last_visible_sheet_and_protection(self):
        current = book(sheet())
        old = book(sheet(visible=2, columns=[run(1, 10, True)]))
        plan = self.plan(current, old)
        self.assertIn("没有可见工作表", plan[0]["reason"])
        self.assertEqual(plan[1]["action"], "restore")
        self.assertEqual(len(self.plan(current, old, sheets=False, narrow=False)), 1)
        current = book(sheet(), sheet("Keep"), structure_protected=True)
        current["sheets"][0]["protected"] = True
        plan = self.plan(current, book(old["sheets"][0], sheet("Keep")))
        self.assertTrue(all(item["action"] == "skip" for item in plan))

    def test_renamed_sheets_use_content_checks_and_ambiguous_matches_are_skipped(self):
        plan = self.plan(book(sheet("新Main", code="Code", columns=[run(1, 10)])),
                         book(sheet("Main", code="Code", columns=[run(1, 10, True)])))
        self.assertEqual(plan[0]["old_sheet"], "Main")
        self.assertEqual(plan[0]["action"], "restore")
        plan = self.plan(book(sheet("Unknown", code="Other", words=[])),
                         book(sheet("A", words=[]), sheet("B", words=[])))
        self.assertEqual(plan[0]["action"], "skip")

    def test_column_scan_uses_serialized_spans_and_reads_exact_width_once_per_span(self):
        grid = ColumnSheet()
        grid.hidden = {3: True, 4: True, 16384: True}
        grid.widths = {5: 0.25}
        specs = [{"first": 3, "last": 4, "hidden": True}, {"first": 5, "last": 5, "hidden": False},
                 {"first": 16384, "last": 16384, "hidden": True}]
        runs = repair.column_runs(grid, specs)
        self.assertIn(run(3, 4, True), runs)
        self.assertIn(run(5, 5, width=0.25), runs)
        self.assertIn(run(16384, 16384, True), runs)
        self.assertLess(grid.calls, 6)
        self.assertEqual(repair.column_runs(grid, specs, limit=256)[-1]["last"], 256)

    def test_serialized_hidden_and_zero_width_columns_are_read_without_used_range_limits(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "metadata.xlsx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("xl/workbook.xml", '<workbook xmlns="urn:sheet" xmlns:r="urn:rel"><sheets><sheet name="Main" r:id="rId1"/></sheets></workbook>')
                archive.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="rId1" Target="/xl/worksheets/sheet1.xml"/></Relationships>')
                archive.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="urn:sheet"><cols><col min="2" max="3" hidden="1"/><col min="16384" max="16384" width="0"/></cols></worksheet>')
            self.assertEqual(repair.column_specs(path)["Main"], [{"first": 2, "last": 3, "hidden": True},
                                                               {"first": 16384, "last": 16384, "hidden": True}])

    def test_signature_ignores_expanded_blank_formatting_but_tracks_coordinates_and_formulas(self):
        def content(row, values):
            sample = SimpleNamespace(UsedRange=SimpleNamespace(Row=row, Column=1, Rows=SimpleNamespace(Count=len(values)),
                                                              Columns=SimpleNamespace(Count=len(values[0]))),
                                     Cells=lambda r, c: (r, c), Range=lambda *args: SimpleNamespace(Formula=values))
            return repair.content_info(sample)["signature"]

        original = content(1, (("title", "=A2"), (None, 3)))
        self.assertEqual(original, content(1, (("title", "=A2", None), (None, 3, None), (None, None, None))))
        self.assertNotEqual(original, content(2, (("title", "=A2"), (None, 3))))
        self.assertNotEqual(original, content(1, (("title", "=B2"), (None, 3))))


class HiddenWriteTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.path, self.backup = self.root / "current.xlsx", self.root / "backup.xlsx"
        self.path.write_bytes(b"current")
        self.backup.write_bytes(b"backup")
        self.sheet = ColumnSheet()
        self.workbook = SimpleNamespace(Worksheets=lambda name: self.sheet, Sheets=[self.sheet], ReadOnly=False,
                                        Save=Mock(), Close=Mock())
        self.file = {"path": str(self.path), "old": str(self.backup), "fingerprint": repair.fingerprint(self.path),
                     "old_fingerprint": repair.fingerprint(self.backup), "checks": {"Main": "same"},
                     "cells": [{"sheet": "Main", "kind": "columns", "addr": "B:C", "first": 2, "last": 3, "before": 8.43,
                                "value": None, "action": "restore"},
                               {"sheet": "Main", "kind": "narrow", "addr": "D:D", "first": 4, "last": 4, "before": 8.43,
                                "value": 0.25, "action": "restore"}]}
        for name, value in (("open_workbook_checked", Mock(return_value=self.workbook)),
                            ("content_info", Mock(return_value={"signature": "same"})),
                            ("workbook_locked", Mock(return_value=False))):
            patcher = patch.object(repair, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)

    def test_only_visibility_and_target_widths_change_and_success_is_reported_after_save(self):
        result = repair.apply_file(Mock(), self.file)
        self.assertEqual([item["status"] for item in result], ["ok", "ok"])
        self.assertTrue(self.sheet.hidden[2] and self.sheet.hidden[3])
        self.assertEqual(self.sheet.widths, {4: 0.25})
        self.assertEqual(self.backup.read_bytes(), b"backup")
        self.workbook.Save.assert_called_once()
        self.workbook.Close.assert_called_once_with(SaveChanges=False)

    def test_changed_backup_open_target_readonly_content_or_column_drift_are_not_saved(self):
        for failure in ("backup", "locked", "readonly", "content", "width"):
            with self.subTest(failure=failure):
                self.backup.write_bytes(b"backup")
                self.workbook.ReadOnly = False
                self.sheet.widths = {}
                repair.workbook_locked.return_value = failure == "locked"
                repair.content_info.return_value = {"signature": "changed" if failure == "content" else "same"}
                if failure == "backup":
                    self.backup.write_bytes(b"changed")
                elif failure == "readonly":
                    self.workbook.ReadOnly = True
                elif failure == "width":
                    self.sheet.widths[2] = 20
                with self.assertRaises(ValueError):
                    repair.apply_file(Mock(), self.file)
                self.workbook.Save.assert_not_called()
                self.assertFalse(self.sheet.hidden)

    def test_save_failure_closes_without_saving_and_does_not_return_success(self):
        self.workbook.Save.side_effect = RuntimeError("保存失败")
        with self.assertRaisesRegex(RuntimeError, "保存失败"):
            repair.apply_file(Mock(), self.file)
        self.workbook.Close.assert_called_once_with(SaveChanges=False)

    def test_excel_legacy_container_update_on_open_is_allowed_but_later_external_edits_are_rejected(self):
        def open_legacy(*args, **kwargs):
            self.path.write_bytes(b"Excel OLE container update")
            return self.workbook

        repair.open_workbook_checked.side_effect = open_legacy
        repair.apply_file(Mock(), self.file)
        self.workbook.Save.assert_called_once()
        self.workbook.Save.reset_mock()
        self.path.write_bytes(b"current")
        self.sheet.hidden, self.sheet.widths = {}, {}

        def changed_after_open(*args):
            self.path.write_bytes(b"external edit after Excel open")
            return {"signature": "same"}

        repair.content_info.side_effect = changed_after_open
        with self.assertRaisesRegex(ValueError, "当前文件发生变化"):
            repair.apply_file(Mock(), self.file)
        self.workbook.Save.assert_not_called()

    def test_background_run_keeps_skips_and_releases_lock_when_a_file_fails(self):
        state = {"busy": True, "maintenance": "hidden_repair"}
        planned = {**self.file, "cells": [*self.file["cells"], {"sheet": "Main", "kind": "sheet", "addr": "A1",
                     "action": "skip", "reason": "位置不同"}]}
        with patch.object(wf, "STATE", state), patch.object(wf, "pythoncom", Mock()), patch.object(wf, "log_msg"), \
                patch.object(repair, "_STATUS", {}), patch.object(repair, "start_background_excel", return_value=Mock()), \
                patch.object(repair, "quit_background_excel"), patch.object(repair, "plan_file", return_value=planned), \
                patch.object(repair, "apply_file", side_effect=RuntimeError("保存失败")), \
                patch.object(repair.debug_log, "event"), patch.object(repair.debug_log, "exception"):
            repair._run(self.root, [self.path], self.root, [self.path], repair.DEFAULT_OPTIONS)
            file = repair.status()["applied"]["files"][0]
            self.assertIn("保存失败", file["error"])
            self.assertEqual(len(file["cells"]), 3)
            self.assertFalse(any(item.get("status") == "ok" for item in file["cells"]))
            self.assertFalse(state["busy"])
            self.assertNotIn("maintenance", state)

    def test_start_validates_scope_folders_busy_and_unsettled_writes_and_thread_failure(self):
        current, old = self.root / "current", self.root / "old"
        current.mkdir()
        old.mkdir()
        (current / "book.xlsx").write_bytes(b"current")
        (old / "book.xlsx").write_bytes(b"old")
        (current / "binary.xlsb").write_bytes(b"binary")
        (current / "~$binary.xlsb").write_bytes(b"lock")
        self.assertEqual(len(repair.list_workbooks(str(current))[1]), 2)
        state = {"busy": False, "job": None}
        with patch.object(wf, "STATE", state), patch.object(repair, "_STATUS", {"running": False}), \
                patch.object(repair.threading, "Thread") as thread:
            for options in ({"sheets": False, "columns": False, "narrow": False}, {"columns": "yes"}, []):
                with self.assertRaises(ValueError):
                    repair.repair(str(current), str(old), options)
            with self.assertRaises(ValueError):
                repair.repair(str(current), str(self.root))
            state["job"] = {"pending_adjustment": {"path": str(self.path)}}
            with self.assertRaises(ValueError):
                repair.repair(str(current), str(old))
            state["job"] = None
            thread.return_value.start.side_effect = RuntimeError("线程失败")
            with self.assertRaises(RuntimeError):
                repair.repair(str(current), str(old))
            self.assertFalse(state["busy"])
            self.assertNotIn("maintenance", state)
            self.assertFalse(repair.status()["running"])
            thread.return_value.start.side_effect = None
            repair.repair(str(current), str(old))
            self.assertTrue(state["busy"] and repair.status()["running"])
            with self.assertRaises(ValueError):
                repair.repair(str(current), str(old))


if __name__ == "__main__":
    unittest.main()
