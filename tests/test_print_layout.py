import copy
import http.client
import importlib
import json
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from core import print_layout as layout, storage, page_breaks
from core.pdf_detector import fitz


def metrics(**changes):
    result = {"pages": 2, "font_pt": 11.0, "across": 1, "visible": list("Content"),
              "hidden": [], "borders": [], "cuts": [], "images": 0}
    result.update(changes)
    return result


class Sheet:
    def __init__(self, path, name="Wide", code="Sheet1", orientation=1):
        self.Name, self.CodeName = name, code
        self.Parent = SimpleNamespace(FullName=str(path))
        props = dict.fromkeys(layout.SETTINGS)
        props.update(Orientation=orientation, PaperSize=9, PrintArea="", Zoom=False,
                     FitToPagesWide=1, FitToPagesTall=False, TopMargin=36., BottomMargin=36.)
        self.PageSetup = SimpleNamespace(**props)
        self.UsedRange = SimpleNamespace(Row=1, Column=1, Value=(("Content",),),
                                         Columns=SimpleNamespace(Count=12))
        self.Cells = lambda row, col: SimpleNamespace(Text="Content", MergeCells=False)
        self.Rows = lambda row: SimpleNamespace(Hidden=False, RowHeight=15.)
        self.Columns = lambda col: SimpleNamespace(Hidden=False, ColumnWidth=10.)
        self.exports = []
        self.ExportAsFixedFormat = lambda **kwargs: self.export(self, Path(kwargs["Filename"]))

    def export(self, sheet, pdf):
        target = self.PageSetup.Orientation
        self.exports.append(target)
        with fitz.open() as doc:
            for _ in range(2 if target == 1 else 1):
                page = doc.new_page(width=595 if target == 1 else 842, height=842 if target == 1 else 595)
                page.insert_text((60, 100), "Content", fontsize=11)
            doc.save(pdf)


class DecisionTests(unittest.TestCase):
    def choose(self, original=1, base=None, other=None):
        return layout.choose(original, {original: base or metrics(), 3-original: other or metrics(pages=1, across=0)})[0]

    def test_reduce_horizontal_split(self):
        self.assertEqual(self.choose(), 2)

    def test_portrait_is_also_a_candidate(self):
        self.assertEqual(self.choose(original=2), 1)

    def test_tie_keeps_original(self):
        self.assertEqual(self.choose(other=metrics()), 1)

    def test_more_readable_without_new_split(self):
        self.assertEqual(self.choose(other=metrics(font_pt=13)), 2)

    def test_fewer_pages_cannot_hide_tiny_text(self):
        self.assertEqual(self.choose(other=metrics(pages=1, across=0, font_pt=5)), 1)

    def test_added_problem_in_another_cell_rejects_candidate(self):
        base = metrics(hidden=[[1, 1, "missing"], [2, 1, "missing"]])
        other = metrics(hidden=[[3, 1, "missing"]], pages=1, across=0)
        self.assertEqual(self.choose(base=base, other=other), 1)

    def test_content_loss_border_or_merge_cut_rejects_candidate(self):
        for change in ({"visible": list("Cont")}, {"borders": ["Content"]}, {"cuts": ["A1:B3"]}):
            with self.subTest(change=change):
                self.assertEqual(self.choose(other=metrics(pages=1, across=0, **change)), 1)

    def test_unknown_pagination_or_image_only_keeps_original(self):
        self.assertEqual(self.choose(other=metrics(across=None)), 1)
        self.assertEqual(self.choose(base=metrics(visible=[])), 1)


class LayoutTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.path = self.root / "book.xlsx"
        self.path.write_bytes(b"source workbook")
        self.records = {}
        layout.configure(lambda: self.records)
        self.sheet = Sheet(self.path)
        layout.bind(self.sheet.Parent, str(self.path))
        patcher = patch.object(page_breaks, "read_breaks", side_effect=lambda sheet: {
            "h": [], "v": [7] if sheet.PageSetup.Orientation == 1 else []})
        patcher.start()
        self.addCleanup(patcher.stop)

    def prepare(self, sheet=None, output=None):
        sheet = sheet or self.sheet
        return layout.prepare(sheet, str(self.path), self.records, sheet.export, output)

    def test_real_pdf_comparison_reuses_selected_pdf_and_restores_source(self):
        before = self.path.read_bytes()
        pdf = self.root / "chosen.pdf"
        record, generated = self.prepare(output=pdf)
        self.assertEqual(record["selected"], 2)
        self.assertTrue(generated)
        self.assertEqual(self.sheet.exports, [1, 2])
        self.assertEqual(self.sheet.PageSetup.Orientation, 1)
        self.assertEqual(self.path.read_bytes(), before)
        with fitz.open(pdf) as doc:
            self.assertEqual(len(doc), 1)
            self.assertGreater(doc[0].rect.width, doc[0].rect.height)

    def test_valid_record_does_not_reexport(self):
        record, _ = self.prepare()
        other, generated = self.prepare()
        self.assertEqual(record["revision"], other["revision"])
        self.assertFalse(generated)
        self.assertEqual(len(self.sheet.exports), 2)

    def test_external_change_and_print_settings_change_reanalyze(self):
        first, _ = self.prepare()
        self.path.write_bytes(b"new workbook contents")
        second, _ = self.prepare()
        self.assertNotEqual(first["revision"], second["revision"])
        self.sheet.PageSetup.TopMargin = 50.
        third, _ = self.prepare()
        self.assertNotEqual(second["revision"], third["revision"])
        self.assertEqual(len(self.sheet.exports), 6)

    def test_known_program_save_keeps_direction_fixed(self):
        first, _ = self.prepare()
        before = layout.version(str(self.path))
        self.path.write_bytes(b"after row adjustment")
        layout.followed_save(self.records, str(self.path), before)
        second, _ = self.prepare()
        self.assertEqual(first["revision"], second["revision"])
        self.assertEqual(len(self.sheet.exports), 2)

    def test_failed_analysis_falls_back_without_persisting_orientation(self):
        self.sheet.export = Mock(side_effect=RuntimeError("export failed"))
        record, generated = self.prepare()
        self.assertEqual(record["selected"], 1)
        self.assertTrue(record["fallback"])
        self.assertFalse(generated)
        self.assertEqual(self.sheet.PageSetup.Orientation, 1)

    def test_com_orientation_can_be_returned_as_a_float(self):
        self.sheet.PageSetup.Orientation = 1.0
        record, _ = self.prepare()
        self.assertEqual(record["selected"], 2)
        self.assertEqual(type(record["original"]), int)

    def test_footer_page_numbers_do_not_imply_body_content_loss(self):
        def export(sheet, pdf):
            with fitz.open() as doc:
                target = sheet.PageSetup.Orientation
                for index in range(4 if target == 1 else 1):
                    page = doc.new_page(width=595 if target == 1 else 842, height=842 if target == 1 else 595)
                    page.insert_text((60, 100), "Content", fontsize=11)
                    page.insert_text((60, page.rect.height-12), str(index+1), fontsize=6)
                doc.save(pdf)
        self.sheet.export = export
        record, _ = self.prepare()
        self.assertEqual(record["selected"], 2)
        self.assertEqual(record["metrics"]["1"]["font_pt"], 11.)

    def test_nested_application_and_exceptions_always_restore(self):
        self.prepare()
        with self.assertRaisesRegex(RuntimeError, "inside"):
            with layout.applied(self.sheet):
                self.assertEqual(self.sheet.PageSetup.Orientation, 2)
                with layout.applied(self.sheet):
                    self.assertEqual(self.sheet.PageSetup.Orientation, 2)
                raise RuntimeError("inside")
        self.assertEqual(self.sheet.PageSetup.Orientation, 1)

    def test_failure_to_restore_stops_instead_of_publishing_a_plan(self):
        class Stuck(SimpleNamespace):
            def __setattr__(self, name, value):
                if name == "Orientation" and value == 1 and getattr(self, name, None) == 2:
                    raise RuntimeError("Excel refused restoration")
                super().__setattr__(name, value)
        self.sheet.PageSetup = Stuck(**vars(self.sheet.PageSetup))
        with self.assertRaisesRegex(RuntimeError, "未恢复"):
            self.prepare()
        self.assertEqual(self.records, {})
        self.assertEqual(self.path.read_bytes(), b"source workbook")

    def test_stale_record_is_not_applied_before_revalidation(self):
        self.prepare()
        self.path.write_bytes(b"external changes")
        with layout.applied(self.sheet):
            self.assertEqual(self.sheet.PageSetup.Orientation, 1)

    def test_copy_uses_source_identity(self):
        self.prepare()
        copy_sheet = Sheet(self.root / "copy.xlsx")
        layout.bind(copy_sheet.Parent, str(self.path))
        with layout.applied(copy_sheet):
            self.assertEqual(copy_sheet.PageSetup.Orientation, 2)
        self.assertEqual(copy_sheet.PageSetup.Orientation, 1)

    def test_each_sheet_is_independent_and_sheet_rename_uses_code(self):
        first, _ = self.prepare()
        second_sheet = Sheet(self.path, name="Tall", code="Sheet2", orientation=1)
        second_sheet.export = self.sheet.export
        with patch.object(layout, "inspect_pdf", side_effect=[metrics(), metrics()]):
            second, _ = self.prepare(second_sheet)
        self.assertEqual(first["selected"], 2)
        self.assertEqual(second["selected"], 1)
        self.sheet.Name = "Renamed"
        again, _ = self.prepare()
        self.assertEqual(again["revision"], first["revision"])
        self.assertEqual(again["sheet"], "Renamed")
        self.assertEqual(len(self.records[str(self.path)]), 2)

    def test_manual_override_survives_content_changes(self):
        record, _ = self.prepare()
        record.update(manual=1, selected=1)
        self.path.write_bytes(b"external content changes")
        record, generated = self.prepare()
        self.assertEqual(record["selected"], 1)
        self.assertFalse(generated)
        self.assertEqual(len(self.sheet.exports), 2)

    def test_file_change_during_analysis_does_not_publish_record(self):
        original_export = self.sheet.export
        def export(sheet, pdf):
            original_export(sheet, pdf)
            self.path.write_bytes(self.path.read_bytes() + b"changed")
        self.sheet.export = export
        with self.assertRaisesRegex(ValueError, "文件已修改"):
            self.prepare()
        self.assertEqual(self.records, {})

    def test_owned_file_and_name_only_sheet_rename(self):
        record, _ = self.prepare()
        self.records[str(self.path)] = {"name:Wide": {**record, "identity": "name:Wide"}}
        new = self.root / "renamed.xlsx"
        self.path.rename(new)
        layout.rename(self.records, str(self.path), str(new), {"Wide": "New"})
        self.assertNotIn(str(self.path), self.records)
        self.assertEqual(self.records[str(new)]["name:New"]["sheet"], "New")

    def test_ooxml_identity_recognizes_renamed_sheet_without_codename(self):
        from openpyxl import Workbook
        book = Workbook()
        book.active.title = "Wide"
        book.save(self.path)
        self.sheet.CodeName = ""
        key = layout.identity(self.sheet)[0]
        self.assertTrue(key.startswith("xml:"))
        record, _ = self.prepare()
        record.update(manual=1, selected=1)
        book.active.title = "Renamed"
        book.save(self.path)
        self.sheet.Name = "Renamed"
        self.assertEqual(layout.identity(self.sheet)[0], key)
        record, _ = self.prepare()
        self.assertEqual(record["manual"], 1)
        self.assertEqual(record["sheet"], "Renamed")

    def test_deleted_or_reassigned_sheet_ids_are_pruned(self):
        record, _ = self.prepare()
        layout.prune(self.records, str(self.path), [("code:Sheet2", "Different")])
        self.assertEqual(self.records[str(self.path)], {})


class PersistenceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        patcher = patch.object(storage, "DATABASE_PATH", self.root / "test.sqlite3")
        patcher.start()
        self.addCleanup(patcher.stop)
        from core import workflow, reporter
        for name, value in (("RECORD_PATH", self.root / "record.md"),
                            ("REPORT_STATE_PATH", self.root / "legacy.json")):
            patcher = patch.object(reporter, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        self.workflow = importlib.reload(workflow)
        self.addCleanup(reporter.flush_execution_report)
        self.path = self.root / "book.xlsx"
        self.path.write_bytes(b"workbook")
        self.workflow.STATE["folder"] = str(self.root)
        self.sheet = Sheet(self.path)
        with patch.object(layout, "inspect_pdf", side_effect=[metrics(), metrics(pages=1, across=0)]):
            self.workflow.prepare_print_layout(self.sheet, str(self.path))
        self.assertEqual(self.workflow.STATE["print_layouts"][str(self.path)]["code:Sheet1"]["selected"], 2)

    def test_restart_and_progress_transfer_preserve_plans(self):
        exported = self.workflow.export_progress()
        self.assertIn(str(self.path), exported["file_versions"])
        restored = importlib.reload(self.workflow)
        self.assertEqual(restored.STATE["print_layouts"], exported["workflow"]["print_layouts"])
        restored.STATE["print_layouts"] = {}
        restored.import_progress(exported)
        self.assertEqual(restored.STATE["print_layouts"], exported["workflow"]["print_layouts"])

    def test_manual_choice_invalidates_old_scan_and_survives_restart(self):
        self.workflow.STATE["checked_files"][str(self.path)] = self.workflow._file_version(self.path)
        self.workflow.set_print_layout(str(self.path), "code:Sheet1", "portrait")
        self.assertNotIn(str(self.path), self.workflow.STATE["checked_files"])
        restored = importlib.reload(self.workflow)
        self.assertEqual(restored.STATE["print_layouts"][str(self.path)]["code:Sheet1"]["selected"], 1)

    def test_manual_change_invalidates_old_unapplied_fix(self):
        from core import batch
        card = {"wb_path": str(self.path), "sheet_name": "Wide", "color": "blue", "verified": True,
                "base_height": 15., "attempt_height": 30., "attempt_wrap": ["A1"], "after_kind": "attempt"}
        self.workflow.STATE["batch_review"]["cards"] = {"card": card}
        self.workflow.set_print_layout(str(self.path), "code:Sheet1", "portrait")
        self.assertTrue(card["layout_stale"])
        self.assertFalse(card["verified"])
        self.assertFalse(batch._adoptable(card))
        self.assertNotIn("attempt_height", card)

    def test_busy_unsettled_and_paused_jobs_block_direction_changes(self):
        for change in ({"busy": True}, {"job": {"pending_adjustment": {"x": 1}}},
                       {"job": {"status": "interrupted"}}):
            with self.subTest(change=change), patch.dict(self.workflow.STATE, change):
                with self.assertRaises(ValueError):
                    self.workflow.set_print_layout(str(self.path), "code:Sheet1", "portrait")

    def test_pdf_cache_is_invalidated_by_direction_revision(self):
        from core.autofix import _ExcelSession
        pdf = self.root / "cached.pdf"
        self.sheet.export(self.sheet, pdf)
        session = _ExcelSession()
        version = self.workflow._file_version(self.path)
        session.remember_pdf(str(self.path), "Wide", version, pdf)
        self.assertEqual(session.pdf_for(str(self.path), "Wide", version), pdf)
        self.workflow.set_print_layout(str(self.path), "code:Sheet1", "portrait")
        self.assertIsNone(session.pdf_for(str(self.path), "Wide", version))

    def test_batch_export_uses_saved_direction_without_saving_source(self):
        source = self.path.read_bytes()
        book = Mock(FullName=str(self.path))
        book.Sheets = [self.sheet]
        excel = Mock()
        excel.Workbooks.Open.return_value = book
        with patch.object(self.workflow, "_exportable_sheets", return_value=[self.sheet]), \
                patch.object(self.workflow, "update_file_progress"):
            exported = self.workflow._export_workbook_pdfs(lambda: excel, self.path, self.root, set())
        book.Save.assert_not_called()
        book.Close.assert_called_once_with(SaveChanges=False)
        self.assertTrue(excel.Workbooks.Open.call_args.kwargs["ReadOnly"])
        self.assertEqual(self.path.read_bytes(), source)
        self.assertEqual(self.sheet.PageSetup.Orientation, 1)
        with fitz.open(exported[0]) as doc:
            self.assertGreater(doc[0].rect.width, doc[0].rect.height)

    def test_http_lists_and_updates_plans_and_queues_reanalysis(self):
        import web_clip_fixer
        handler = importlib.reload(web_clip_fixer).WebHandler
        server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
        worker = threading.Thread(target=server.serve_forever, daemon=True)
        worker.start()
        try:
            def request(method, url, payload=None):
                connection = http.client.HTTPConnection("127.0.0.1", server.server_address[1], timeout=5)
                try:
                    connection.request(method, url, None if payload is None else json.dumps(payload),
                                       {"Content-Type": "application/json"})
                    response = connection.getresponse()
                    return response.status, json.loads(response.read())
                finally:
                    connection.close()
            status, data = request("GET", "/api/print_layouts")
            self.assertEqual(status, 200)
            self.assertEqual(data["rows"][0]["selected"], 2)
            payload = {"path": str(self.path), "key": "code:Sheet1", "direction": "portrait"}
            self.assertEqual(request("POST", "/api/print_layout", payload)[0], 200)
            payload["direction"] = "invalid"
            self.assertEqual(request("POST", "/api/print_layout", payload)[0], 400)
            payload["direction"] = "auto"
            with patch.object(handler, "_start_job", autospec=True) as start:
                start.side_effect = lambda instance, *args: instance._send_json({"ok": True})
                self.assertEqual(request("POST", "/api/print_layout", payload)[0], 200)
                self.assertEqual(start.call_args.args[1], "print_layout")
                self.assertEqual(start.call_args.args[2]["key"], "code:Sheet1")
        finally:
            server.shutdown()
            server.server_close()
            worker.join(timeout=5)


if __name__ == "__main__":
    unittest.main()
