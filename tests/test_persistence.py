import copy
import http.client
import importlib
import json
import os
import subprocess
import sys
import tempfile
import threading
import unittest
from contextlib import contextmanager
from http.server import ThreadingHTTPServer
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from core import storage


class PersistenceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.database = self.folder / "test.sqlite3"
        self.patch_value(storage, "DATABASE_PATH", self.database)
        from core import reporter, workflow

        self.reporter = importlib.reload(reporter)
        self.workflow = importlib.reload(workflow)
        self.patch_value(reporter, "RECORD_PATH", self.folder / "record.md")
        self.patch_value(reporter, "REPORT_STATE_PATH", self.folder / "legacy.json")
        self.wb_path = self.folder / "book.xlsx"
        self.wb_path.write_bytes(b"test workbook")
        self.pdf_path = self.folder / "sheet.pdf"
        with workflow.fitz.open() as doc:
            doc.new_page()
            doc.save(self.pdf_path)
        workflow.STATE.update(
            folder=str(self.folder),
            backup_folder=str(self.folder / "backup"),
            pdf_folder=str(self.folder),
        )
        workflow.STATE["items"]["item_1"] = {
            "id": "item_1", "seq": 1, "wb_path": str(self.wb_path),
            "wb_name": self.wb_path.name, "sheet_name": "Sheet1",
            "excel_row": 2, "cell_addr": "A2", "row_height": 15.0,
            "pdf_path": str(self.pdf_path), "page_index": 0,
            "text": "sample", "bbox": [0, 0, 10, 10], "image": "old",
            "status": "pending", "iterations": 0,
        }
        self.row = SimpleNamespace(RowHeight=15.0)
        self.saved_heights = []
        self.sheet = Mock()
        self.sheet.Rows.return_value = self.row
        self.wb = Mock()
        self.wb.Sheets.return_value = self.sheet
        self.wb.Save.side_effect = lambda: self.saved_heights.append(self.row.RowHeight)
        self.excel = Mock()
        self.excel.Workbooks.Open.return_value = self.wb
        self.patch_value(workflow, "pythoncom", Mock())
        self.patch_value(workflow, "win32", Mock(DispatchEx=Mock(return_value=self.excel)))
        self.patch_value(workflow, "normalize_workbook_view", Mock())
        self.patch_value(workflow, "_refresh_sheet_pages_in_state", Mock())
        self.detect = Mock(return_value=(True, "new", [0, 0, 10, 10], 0))
        self.patch_value(workflow, "is_region_still_clipped", self.detect)
        self.rescan = Mock(return_value=[])
        self.patch_value(workflow, "find_clipped_spans_on_page", self.rescan)
        self.patch_value(workflow, "extract_sheet_snapshot", Mock(return_value={}))
        self.patch_value(workflow, "match_excel_row", Mock(return_value=(2, "A2", 37.0)))

    def patch_value(self, target, name, value):
        patcher = patch.object(target, name, value)
        patcher.start()
        self.addCleanup(patcher.stop)

    def start_fix(self, max_iter=2):
        self.assertTrue(self.workflow.queue_job("fix", {
            "selected_ids": ["item_1"], "step_pt": 3.0, "max_iter": max_iter,
        }))
        self.workflow.run_pending_workflow()

    def snapshot(self, name):
        state = storage.load_snapshot(name)
        assert state is not None
        return state

    def restart(self):
        loaded = self.snapshot("workflow")
        self.workflow.STATE.clear()
        self.workflow.STATE.update(copy.deepcopy(loaded))
        self.workflow.STATE["busy"] = False
        if self.workflow.STATE["job"] and self.workflow.STATE["job"]["status"] == "running":
            self.workflow.STATE["job"]["status"] = "interrupted"

    def test_custom_loop_limit_is_used_and_saved(self):
        self.start_fix(max_iter=8)
        item = self.snapshot("workflow")["items"]["item_1"]
        self.assertEqual(item["iterations"], 8)
        self.assertEqual(item["row_height"], 39.0)
        self.assertEqual(item["status"], "max_reached")
        self.assertEqual(self.sheet.ExportAsFixedFormat.call_count, 8)
        self.assertEqual(self.snapshot("workflow")["settings"]["max_iter"], 8)

    def test_stop_as_soon_as_fixed(self):
        self.detect.side_effect = [
            (True, "clipped", [0, 0, 10, 10], 0),
            (False, "fixed", [0, 0, 10, 10], 0),
        ]
        self.start_fix(max_iter=8)
        self.assertEqual(self.sheet.ExportAsFixedFormat.call_count, 1)
        self.assertEqual(self.workflow.STATE["items"]["item_1"]["status"], "fixed")
        self.assertEqual(self.snapshot("report")["files"]["book.xlsx"]["clip_fixed"], 1)

    def test_same_row_items_share_one_iteration_budget(self):
        second = copy.deepcopy(self.workflow.STATE["items"]["item_1"])
        second.update(id="item_2", cell_addr="B2", text="another text")
        self.workflow.STATE["items"]["item_2"] = second
        self.workflow.queue_job("fix", {
            "selected_ids": ["item_1", "item_2"], "step_pt": 10.0, "max_iter": 6,
        })
        self.workflow.run_pending_workflow()
        self.assertEqual(self.sheet.ExportAsFixedFormat.call_count, 6)
        self.assertEqual(self.row.RowHeight, 75.0)
        for item in self.workflow.STATE["items"].values():
            self.assertEqual(item["job_iterations"], 6)
            self.assertEqual(item["status"], "max_reached")

    def test_tracking_failure_stops_before_modifying_excel(self):
        self.detect.side_effect = self.workflow.RegionTrackingError("ambiguous target")
        self.start_fix()
        self.wb.Save.assert_not_called()
        self.sheet.ExportAsFixedFormat.assert_not_called()
        item = self.snapshot("workflow")["items"]["item_1"]
        self.assertEqual(item["status"], "needs_review")
        self.assertEqual(item["review_reason"], "ambiguous target")

    def test_workflow_updates_item_page_and_canvas_after_pagination_changes(self):
        with self.workflow.fitz.open() as doc:
            doc.new_page()
            doc.new_page()
            doc.save(self.pdf_path)
        self.detect.side_effect = [
            (True, "old", (0, 0, 10, 10), 0),
            (False, "new", (0, 200, 10, 210), 1),
        ]
        self.start_fix()
        item = self.workflow.STATE["items"]["item_1"]
        self.assertEqual(item["page_index"], 1)
        page = self.workflow.STATE["pages"][item["page_id"]]
        self.assertEqual(page["page_index"], 1)
        self.assertEqual(page["total_pages"], 2)

    def test_manual_open_locates_item_and_persists_review_status(self):
        opener = Mock(return_value=33.0)
        self.patch_value(self.workflow, "open_workbook_for_manual_review", opener)
        self.workflow.queue_job("manual_open", {"selected_ids": ["item_1"]})
        self.workflow.run_pending_workflow()
        opener.assert_called_once_with(self.wb_path, "Sheet1", "A2", 2)
        self.assertEqual(self.snapshot("workflow")["items"]["item_1"]["status"], "manual_review")
        self.assertEqual(self.row.RowHeight, 15.0)
        self.wb.Save.assert_not_called()

    def test_manual_open_cannot_run_during_automatic_adjustment(self):
        self.workflow.STATE["busy"] = True
        self.assertFalse(self.workflow.queue_job("manual_open", {"selected_ids": ["item_1"]}))

    def test_manual_confirmation_is_persistent_separate_from_fixed_and_reversible(self):
        self.workflow.set_manual_confirmation("item_1", True)
        saved = self.snapshot("workflow")["items"]["item_1"]
        self.assertEqual(saved["status"], "accepted")
        self.assertEqual(saved["manual_confirmation"]["file_version"], self.workflow._file_version(self.wb_path))
        report = self.snapshot("report")["files"]["book.xlsx"]
        self.assertEqual((report["clip_fixed"], report["clip_accepted"], report["clip_unfixed"]), (0, 1, 0))
        self.assertEqual(report["status"], "done")
        self.assertIn("人工确认", self.reporter.render_execution_markdown())
        self.restart()
        self.workflow.get_state_snapshot()
        self.assertEqual(self.workflow.STATE["items"]["item_1"]["status"], "accepted")
        self.assertEqual(self.workflow._select_changed_problem_items(), [])
        self.workflow.set_manual_confirmation("item_1", False)
        report = self.snapshot("report")["files"]["book.xlsx"]
        self.assertEqual((report["clip_accepted"], report["clip_unfixed"]), (0, 1))
        self.assertEqual(self.workflow.STATE["items"]["item_1"]["status"], "pending")
        self.assertEqual(self.row.RowHeight, 15.0)
        self.wb.Save.assert_not_called()

    def test_modified_file_expires_confirmation_and_reenters_incremental_refresh(self):
        self.workflow.STATE["checked_files"][str(self.wb_path)] = self.workflow._file_version(self.wb_path)
        self.workflow.set_manual_confirmation("item_1", True)
        self.touch_modified(self.wb_path)
        state = self.workflow.get_state_snapshot()
        self.assertEqual(state["items"]["item_1"]["status"], "needs_review")
        self.assertEqual(self.workflow._select_changed_problem_items(), ["item_1"])
        self.assertEqual(self.snapshot("report")["files"]["book.xlsx"]["clip_accepted"], 0)

    def test_confirmed_item_cannot_be_automatically_adjusted_even_with_direct_request(self):
        self.workflow.set_manual_confirmation("item_1", True)
        self.start_fix()
        self.sheet.ExportAsFixedFormat.assert_not_called()
        self.detect.assert_not_called()
        self.assertEqual(self.row.RowHeight, 15.0)
        self.assertEqual(self.workflow.STATE["items"]["item_1"]["status"], "accepted")

    def test_confirmation_rejects_busy_and_rolls_back_on_database_error(self):
        self.workflow.STATE["busy"] = True
        with self.assertRaisesRegex(ValueError, "正在执行"):
            self.workflow.set_manual_confirmation("item_1", True)
        self.workflow.STATE["busy"] = False
        with patch.object(self.workflow, "save_snapshot", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                self.workflow.set_manual_confirmation("item_1", True)
        self.assertEqual(self.workflow.STATE["items"]["item_1"]["status"], "pending")
        self.assertNotIn("manual_confirmation", self.workflow.STATE["items"]["item_1"])

    def test_confirmation_cancels_interrupted_adjustment_checkpoint(self):
        self.sheet.ExportAsFixedFormat.side_effect = RuntimeError("export interrupted")
        self.start_fix()
        self.workflow.set_manual_confirmation("item_1", True)
        saved = self.snapshot("workflow")
        self.assertEqual(saved["job"]["status"], "cancelled")
        self.assertIsNone(saved["job"]["pending_adjustment"])
        self.assertFalse(self.workflow.get_state_snapshot()["can_resume"])

    def test_forced_rescan_preserves_current_version_manual_confirmation_at_same_location(self):
        self.workflow.STATE["items"]["item_1"]["bbox"] = [60, 90, 100, 105]
        self.rescan.return_value = [{"text": "sample", "bbox": (60, 90, 100, 105), "line_y": 100}]
        self.workflow.set_manual_confirmation("item_1", True)
        self.workflow.queue_job("recheck", {"selected_ids": ["item_1"]})
        self.workflow.run_pending_workflow()
        self.assertEqual(len(self.workflow.STATE["items"]), 1)
        item = next(iter(self.workflow.STATE["items"].values()))
        self.assertEqual(item["status"], "accepted")
        self.assertEqual(self.workflow._select_changed_problem_items(), [])

    def test_manual_recheck_exports_without_saving_or_increasing_height(self):
        self.detect.return_value = (False, "fixed", [0, 0, 10, 10], 0)
        self.row.RowHeight = 37.0
        self.workflow.queue_job("recheck", {"selected_ids": ["item_1"]})
        self.workflow.run_pending_workflow()
        self.excel.Workbooks.Open.assert_called_once_with(str(self.wb_path), ReadOnly=True)
        self.wb.Save.assert_not_called()
        self.sheet.ExportAsFixedFormat.assert_called_once()
        self.assertEqual(self.row.RowHeight, 37.0)
        self.assertNotIn("item_1", self.snapshot("workflow")["items"])
        self.detect.assert_not_called()
        self.rescan.assert_called_once()
        self.assertEqual(self.snapshot("report")["files"]["book.xlsx"]["clip_detected"], 0)

    def test_manual_recheck_discards_old_ambiguous_target_when_new_pdf_has_no_clips(self):
        self.detect.side_effect = self.workflow.RegionTrackingError("ambiguous target")
        self.workflow.queue_job("recheck", {"selected_ids": ["item_1"]})
        self.workflow.run_pending_workflow()
        self.assertNotIn("item_1", self.snapshot("workflow")["items"])
        self.detect.assert_not_called()
        self.assertEqual(self.row.RowHeight, 15.0)
        self.wb.Save.assert_not_called()

    def install_old_page(self):
        item = self.workflow.STATE["items"]["item_1"]
        item.update(
            text="er", status="needs_review", page_id="old_page",
            review_reason="multiple old text candidates", bbox=[60, 500, 80, 520],
        )
        self.workflow.STATE["pages"]["old_page"] = {
            "id": "old_page", "wb_path": str(self.wb_path), "wb_name": self.wb_path.name,
            "sheet_name": "Sheet1", "pdf_path": str(self.pdf_path),
            "page_index": 0, "total_pages": 1,
            "page_width": 595, "page_height": 842, "image": "old image",
        }

    def test_real_pdf_repeated_text_without_lines_removes_stale_boxes_and_pages(self):
        from core.pdf_detector import find_clipped_spans_on_page

        self.install_old_page()
        with self.workflow.fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((60, 100), "er")
            page.insert_text((60, 200), "er")
            doc.save(self.pdf_path)
        self.rescan.side_effect = find_clipped_spans_on_page
        self.workflow.queue_job("recheck", {"selected_ids": ["item_1"]})
        self.workflow.run_pending_workflow()
        state = self.snapshot("workflow")
        self.assertEqual(state["items"], {})
        self.assertEqual(state["pages"], {})
        self.assertEqual(state["job"]["status"], "completed")
        self.detect.assert_not_called()
        report = self.snapshot("report")["files"]["book.xlsx"]
        self.assertEqual(report["clip_detected"], 0)
        self.assertEqual(report["clip_unfixed"], 0)
        self.assertEqual(report["status"], "done")

    def test_real_pdf_fresh_scan_finds_new_clip_at_current_position_not_old_bbox(self):
        from core.pdf_detector import find_clipped_spans_on_page

        self.install_old_page()
        with self.workflow.fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((60, 100), "er")
            page.insert_text((60, 200), "er")
            spans = [
                span for block in page.get_text("dict")["blocks"]
                for line in block.get("lines", []) for span in line["spans"]
            ]
            bbox = tuple(spans[1]["bbox"])
            line_y = (bbox[1] + bbox[3]) / 2
            page.draw_line((55, line_y), (150, line_y), width=0.5)
            doc.save(self.pdf_path)
        self.rescan.side_effect = find_clipped_spans_on_page
        self.workflow.queue_job("recheck", {"selected_ids": ["item_1"]})
        self.workflow.run_pending_workflow()
        state = self.snapshot("workflow")
        self.assertNotIn("item_1", state["items"])
        self.assertNotIn("old_page", state["pages"])
        self.assertEqual(len(state["items"]), 1)
        current = next(iter(state["items"].values()))
        self.assertEqual(tuple(current["bbox"]), bbox)
        self.assertEqual(current["status"], "pending")
        self.assertNotIn("review_reason", current)
        self.assertEqual(current["cell_addr"], "A2")
        self.assertIn(current["page_id"], state["pages"])
        self.workflow.match_excel_row.assert_called_once()
        self.detect.assert_not_called()

    def test_rescan_second_sheet_failure_keeps_all_old_items_and_pages(self):
        self.install_old_page()
        second = copy.deepcopy(self.workflow.STATE["items"]["item_1"])
        second.update(id="item_2", sheet_name="Sheet2", pdf_path=str(self.folder / "second.pdf"))
        self.workflow.STATE["items"]["item_2"] = second
        before_items = copy.deepcopy(self.workflow.STATE["items"])
        before_pages = copy.deepcopy(self.workflow.STATE["pages"])
        self.sheet.ExportAsFixedFormat.side_effect = [None, RuntimeError("second sheet failed")]
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        saved = self.snapshot("workflow")
        self.assertEqual(saved["items"], before_items)
        self.assertEqual(saved["pages"], before_pages)
        self.assertEqual(saved["job"]["completed"], [])
        self.assertEqual(saved["job"]["status"], "interrupted")

    def add_problem_file(self, index, status="pending"):
        path = self.folder / f"book_{index}.xlsx"
        path.write_bytes(b"test workbook")
        item = copy.deepcopy(self.workflow.STATE["items"]["item_1"])
        item.update(
            id=f"extra_{index}", wb_path=str(path), wb_name=path.name,
            pdf_path=str(self.folder / f"sheet_{index}.pdf"), status=status,
        )
        self.workflow.STATE["items"][item["id"]] = item
        Path(item["pdf_path"]).write_bytes(self.pdf_path.read_bytes())
        self.workflow.STATE["checked_files"][str(path)] = self.workflow._file_version(path)
        return path, item

    def touch_modified(self, path):
        version = path.stat()
        os.utime(path, ns=(version.st_atime_ns, version.st_mtime_ns + 2_000_000_000))

    def test_refresh_filters_100_files_to_only_two_changed_problem_files(self):
        self.workflow.STATE["checked_files"][str(self.wb_path)] = self.workflow._file_version(self.wb_path)
        paths = [self.wb_path]
        for index in range(1, 100):
            status = "needs_review" if index == 1 else ("pending" if index < 20 else "fixed")
            path, _ = self.add_problem_file(index, status)
            paths.append(path)
        for path in (paths[0], paths[1], paths[50]):
            self.touch_modified(path)
        unaffected = {
            iid: copy.deepcopy(item)
            for iid, item in self.workflow.STATE["items"].items()
            if item["wb_path"] not in {str(paths[0]), str(paths[1])}
        }
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        self.assertEqual(
            [call.args[0] for call in self.excel.Workbooks.Open.call_args_list],
            [str(paths[0]), str(paths[1])],
        )
        self.assertEqual(self.sheet.ExportAsFixedFormat.call_count, 2)
        self.assertEqual(self.rescan.call_count, 2)
        self.wb.Save.assert_not_called()
        self.assertEqual(self.row.RowHeight, 15.0)
        for iid, item in unaffected.items():
            self.assertEqual(self.workflow.STATE["items"][iid], item)
        saved = self.snapshot("workflow")
        for path in paths[:2]:
            self.assertEqual(saved["checked_files"][str(path)], self.workflow._file_version(path))

        self.restart()
        self.excel.Workbooks.Open.reset_mock()
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        self.excel.Workbooks.Open.assert_not_called()
        self.assertIn("没有修改过", self.workflow.STATE["logs"][-1])

    def test_refresh_opens_each_workbook_and_exports_each_affected_sheet_once(self):
        item = self.workflow.STATE["items"]["item_1"]
        second = copy.deepcopy(item)
        second.update(id="item_2", text="second", excel_row=3)
        third = copy.deepcopy(item)
        third.update(id="item_3", text="third", sheet_name="Sheet2", pdf_path=str(self.folder / "second.pdf"))
        Path(third["pdf_path"]).write_bytes(self.pdf_path.read_bytes())
        self.workflow.STATE["items"].update(item_2=second, item_3=third)
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        self.excel.Workbooks.Open.assert_called_once()
        self.assertEqual(self.sheet.ExportAsFixedFormat.call_count, 2)
        self.assertEqual(self.rescan.call_count, 2)
        self.detect.assert_not_called()
        self.assertEqual(self.workflow.STATE["job"]["status"], "completed")

    def test_unchanged_file_can_be_forced_and_checks_all_its_unresolved_items(self):
        self.workflow.STATE["checked_files"][str(self.wb_path)] = self.workflow._file_version(self.wb_path)
        second = copy.deepcopy(self.workflow.STATE["items"]["item_1"])
        second.update(id="item_2", text="second")
        self.workflow.STATE["items"]["item_2"] = second
        self.workflow.queue_job("recheck", {"selected_ids": ["item_1"]})
        self.workflow.run_pending_workflow()
        self.rescan.assert_called_once()
        self.detect.assert_not_called()
        self.sheet.ExportAsFixedFormat.assert_called_once()
        self.excel.Workbooks.Open.assert_called_once()
        self.assertEqual(self.workflow.STATE["job"]["options"]["selected_ids"], ["item_1", "item_2"])

    def test_refresh_empty_selection_does_not_start_excel_or_replace_baseline(self):
        baseline = self.workflow._file_version(self.wb_path)
        self.workflow.STATE["checked_files"][str(self.wb_path)] = baseline
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        self.workflow.win32.DispatchEx.assert_not_called()
        self.assertEqual(self.snapshot("workflow")["checked_files"][str(self.wb_path)], baseline)
        self.assertEqual(self.workflow.STATE["job"]["status"], "completed")

    def test_legacy_problem_without_timestamp_is_checked_once(self):
        self.assertEqual(self.workflow.STATE["checked_files"], {})
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        self.sheet.ExportAsFixedFormat.assert_called_once()
        self.assertTrue(any("尚无检查时间戳" in line for line in self.workflow.STATE["logs"]))
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        self.sheet.ExportAsFixedFormat.assert_called_once()

    def test_file_size_change_is_detected_even_when_modification_time_is_preserved(self):
        stat = self.wb_path.stat()
        self.workflow.STATE["checked_files"][str(self.wb_path)] = self.workflow._file_version(self.wb_path)
        self.wb_path.write_bytes(b"modified workbook with a different size")
        os.utime(self.wb_path, ns=(stat.st_atime_ns, stat.st_mtime_ns))
        self.assertEqual(self.wb_path.stat().st_mtime_ns, stat.st_mtime_ns)
        self.assertEqual(self.workflow._select_changed_problem_items(), ["item_1"])

    def test_refresh_failure_does_not_advance_baseline_and_can_retry(self):
        baseline = self.workflow._file_version(self.wb_path)
        self.workflow.STATE["checked_files"][str(self.wb_path)] = baseline
        self.touch_modified(self.wb_path)
        self.sheet.ExportAsFixedFormat.side_effect = RuntimeError("export failed")
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        self.assertEqual(self.snapshot("workflow")["checked_files"][str(self.wb_path)], baseline)
        self.assertEqual(self.workflow.STATE["job"]["status"], "interrupted")
        self.sheet.ExportAsFixedFormat.side_effect = None
        self.restart()
        self.workflow.queue_job("resume")
        self.workflow.run_pending_workflow()
        self.assertEqual(
            self.snapshot("workflow")["checked_files"][str(self.wb_path)],
            self.workflow._file_version(self.wb_path),
        )

    def test_file_changed_during_refresh_does_not_commit_results_or_timestamp(self):
        baseline = self.workflow._file_version(self.wb_path)
        self.workflow.STATE["checked_files"][str(self.wb_path)] = baseline
        self.touch_modified(self.wb_path)
        def save_during_export(**kwargs):
            self.assertEqual(kwargs["Filename"], str(self.pdf_path))
            self.touch_modified(self.wb_path)

        self.sheet.ExportAsFixedFormat.side_effect = save_during_export
        self.detect.return_value = (False, "fixed", [0, 0, 10, 10], 0)
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        saved = self.snapshot("workflow")
        self.assertEqual(saved["checked_files"][str(self.wb_path)], baseline)
        self.assertEqual(saved["items"]["item_1"]["status"], "pending")
        self.assertEqual(saved["job"]["completed"], [])
        self.assertEqual(saved["job"]["status"], "interrupted")

    def test_refresh_resumes_without_reopening_already_checked_files(self):
        second_path, _ = self.add_problem_file(2)
        self.touch_modified(second_path)
        self.sheet.ExportAsFixedFormat.side_effect = [None, RuntimeError("export failed")]
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        self.assertEqual(self.workflow.STATE["job"]["completed"], ["item_1"])
        self.restart()
        self.excel.Workbooks.Open.reset_mock()
        self.sheet.ExportAsFixedFormat.side_effect = None
        self.workflow.queue_job("resume")
        self.workflow.run_pending_workflow()
        self.excel.Workbooks.Open.assert_called_once_with(str(second_path), ReadOnly=True)
        self.assertEqual(self.workflow.STATE["job"]["status"], "completed")

    def test_missing_problem_file_reports_error_instead_of_skipping(self):
        self.wb_path.unlink()
        self.workflow.queue_job("refresh")
        self.workflow.run_pending_workflow()
        self.assertEqual(self.workflow.STATE["job"]["status"], "interrupted")
        self.assertTrue(any("book.xlsx" in line for line in self.workflow.STATE["logs"]))
        self.workflow.win32.DispatchEx.assert_not_called()

    def test_manual_open_replaces_unsafe_old_resume_checkpoint(self):
        opener = Mock(return_value=18.0)
        self.patch_value(self.workflow, "open_workbook_for_manual_review", opener)
        self.sheet.ExportAsFixedFormat.side_effect = RuntimeError("export interrupted")
        self.start_fix()
        self.assertIsNotNone(self.workflow.STATE["job"]["pending_adjustment"])
        self.workflow.queue_job("manual_open", {"selected_ids": ["item_1"]})
        self.workflow.run_pending_workflow()
        self.assertFalse(self.workflow.get_state_snapshot()["can_resume"])
        self.assertIsNone(self.snapshot("workflow")["job"]["pending_adjustment"])
        self.assertTrue(self.workflow.update_item_row("item_1", 3))

    def test_resume_after_excel_save_does_not_add_height_twice(self):
        self.sheet.ExportAsFixedFormat.side_effect = [None, RuntimeError("power loss")]
        self.start_fix()
        interrupted = self.snapshot("workflow")
        self.assertEqual(interrupted["job"]["status"], "interrupted")
        self.assertEqual(interrupted["job"]["pending_adjustment"]["target_height"], 21.0)
        self.assertEqual(interrupted["items"]["item_1"]["job_iterations"], 1)
        self.assertEqual(self.saved_heights, [18.0, 21.0])

        self.restart()
        self.sheet.ExportAsFixedFormat.side_effect = None
        self.workflow.queue_job("resume")
        self.workflow.run_pending_workflow()

        item = self.workflow.STATE["items"]["item_1"]
        self.assertEqual(item["iterations"], 2)
        self.assertEqual(item["job_iterations"], 2)
        self.assertEqual(item["row_height"], 21.0)
        self.assertEqual(self.saved_heights, [18.0, 21.0, 21.0, 21.0])
        self.assertEqual(self.workflow.STATE["job"]["status"], "completed")
        self.assertIsNone(self.workflow.STATE["job"]["pending_adjustment"])
        self.assertEqual(self.snapshot("report")["files"]["book.xlsx"]["error"], "")

    def test_resume_after_last_iteration_commit_does_not_repeat_loop(self):
        normalize = self.workflow.normalize_workbook_view
        normalize.side_effect = RuntimeError("interrupted before final save")
        self.start_fix()
        self.assertEqual(self.sheet.ExportAsFixedFormat.call_count, 2)
        self.restart()
        normalize.side_effect = None
        self.workflow.queue_job("resume")
        self.workflow.run_pending_workflow()
        self.assertEqual(self.sheet.ExportAsFixedFormat.call_count, 2)
        self.assertEqual(self.workflow.STATE["items"]["item_1"]["iterations"], 2)
        self.assertEqual(self.workflow.STATE["job"]["status"], "completed")

    def test_resume_before_excel_save_replays_recorded_target_height(self):
        self.wb.Save.side_effect = RuntimeError("save interrupted")
        self.start_fix(max_iter=1)
        self.assertEqual(self.snapshot("workflow")["job"]["pending_adjustment"]["target_height"], 18.0)
        self.restart()
        self.row.RowHeight = 15.0
        self.wb.Save.side_effect = None
        self.workflow.queue_job("resume")
        self.workflow.run_pending_workflow()
        self.assertEqual(self.row.RowHeight, 18.0)
        self.assertEqual(self.workflow.STATE["items"]["item_1"]["iterations"], 1)
        self.assertEqual(self.sheet.ExportAsFixedFormat.call_count, 1)

    def test_restart_in_another_process_restores_results_and_interrupted_job(self):
        self.workflow.queue_job("fix", {
            "selected_ids": ["item_1"], "step_pt": 2.5, "max_iter": 9,
        })
        code = (
            "import json; from pathlib import Path; from core import storage; "
            f"storage.DATABASE_PATH = Path({str(self.database)!r}); "
            "from core.workflow import get_state_snapshot; "
            "s = get_state_snapshot(); "
            "print(json.dumps({k: s[k] for k in "
            "['busy', 'can_resume', 'items', 'settings']}))"
        )
        result = subprocess.run(
            [sys.executable, "-c", code], capture_output=True, text=True, check=True,
            cwd=Path(__file__).resolve().parent.parent,
        )
        restored = json.loads(result.stdout)
        self.assertFalse(restored["busy"])
        self.assertTrue(restored["can_resume"])
        self.assertEqual(restored["items"]["item_1"]["cell_addr"], "A2")
        self.assertEqual(restored["settings"]["max_iter"], 9)

    def test_row_correction_is_persistent_and_rejected_while_busy(self):
        self.assertTrue(self.workflow.update_item_row("item_1", 17))
        self.assertEqual(self.snapshot("workflow")["items"]["item_1"]["cell_addr"], "A17")
        self.workflow.STATE["busy"] = True
        self.assertFalse(self.workflow.update_item_row("item_1", 20))

    def test_busy_request_does_not_replace_job(self):
        self.workflow.queue_job("fix", {
            "selected_ids": ["item_1"], "step_pt": 3.0, "max_iter": 6,
        })
        before = self.snapshot("workflow")
        self.assertFalse(self.workflow.queue_job("scan", {"folder": "other"}))
        self.assertEqual(self.snapshot("workflow"), before)

    def test_report_legacy_json_is_migrated_to_database(self):
        legacy = copy.deepcopy(self.reporter.EXECUTION_STATE)
        legacy["task_name"] = "legacy task"
        self.reporter.REPORT_STATE_PATH.write_text(json.dumps(legacy), encoding="utf-8")
        self.reporter._restore_report()
        self.assertEqual(self.snapshot("report")["task_name"], "legacy task")

    def test_report_restart_does_not_show_stale_processing_status(self):
        self.reporter.init_execution_batch("scan", self.folder, [self.wb_path])
        self.reporter.update_file_progress("book.xlsx", status="processing")
        self.reporter._restore_report()
        report = self.snapshot("report")
        self.assertEqual(report["status"], "interrupted")
        self.assertEqual(report["files"]["book.xlsx"]["status"], "pending")
        self.assertEqual(report["processing_files"], 0)
        self.assertEqual(report["pending_files"], 1)

    def test_database_failure_rejects_job_without_acquiring_busy(self):
        with patch.object(self.workflow, "save_snapshot", side_effect=OSError("disk full")):
            with self.assertRaisesRegex(OSError, "disk full"):
                self.workflow.queue_job("fix", {
                    "selected_ids": ["item_1"], "step_pt": 3.0, "max_iter": 6,
                })
        self.assertFalse(self.workflow.STATE["busy"])
        self.assertIsNone(self.workflow.STATE["job"])

    def test_database_failure_rolls_back_row_correction(self):
        with patch.object(self.workflow, "save_snapshot", side_effect=OSError("disk full")):
            with self.assertRaisesRegex(OSError, "disk full"):
                self.workflow.update_item_row("item_1", 20)
        self.assertEqual(self.workflow.STATE["items"]["item_1"]["excel_row"], 2)
        self.assertEqual(self.workflow.STATE["items"]["item_1"]["cell_addr"], "A2")

    @contextmanager
    def serve_api(self):
        import web_clip_fixer

        importlib.reload(web_clip_fixer)
        with patch.object(web_clip_fixer, "run_pending_workflow"):
            server = ThreadingHTTPServer(("127.0.0.1", 0), web_clip_fixer.WebHandler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                yield server.server_port
            finally:
                server.shutdown()
                server.server_close()
                thread.join()

    def post(self, port, path, payload):
        connection = http.client.HTTPConnection("127.0.0.1", port, timeout=5)
        try:
            connection.request(
                "POST", path, json.dumps(payload), {"Content-Type": "application/json"}
            )
            response = connection.getresponse()
            return response.status, json.loads(response.read())
        finally:
            connection.close()

    def test_api_accepts_loop_limit_and_rejects_concurrent_job(self):
        with self.serve_api() as port:
            status, data = self.post(port, "/api/fix", {
                "selected_ids": ["item_1"], "step_pt": 2.5, "max_iter": 11,
            })
            self.assertEqual((status, data["ok"]), (200, True))
            self.assertEqual(self.snapshot("workflow")["job"]["options"]["max_iter"], 11)
            status, data = self.post(port, "/api/resume", {})
            self.assertEqual((status, data["ok"]), (409, False))

    def test_api_manual_actions_accept_known_ids_only(self):
        with self.serve_api() as port:
            status, data = self.post(port, "/api/manual/open", {"id": "unknown"})
            self.assertEqual((status, data["ok"]), (400, False))
            status, data = self.post(port, "/api/manual/open", {"id": "item_1"})
            self.assertEqual((status, data["ok"]), (200, True))
            self.assertEqual(self.snapshot("workflow")["job"]["kind"], "manual_open")
            self.workflow.STATE["busy"] = False
            status, data = self.post(port, "/api/manual/recheck", {"id": "item_1"})
            self.assertEqual((status, data["ok"]), (200, True))
            self.assertEqual(self.snapshot("workflow")["job"]["kind"], "recheck")
            self.workflow.STATE["busy"] = False
            status, data = self.post(port, "/api/manual/refresh", {})
            self.assertEqual((status, data["ok"]), (200, True))
            self.assertEqual(self.snapshot("workflow")["job"]["kind"], "refresh")

    def test_api_manual_confirmation_and_revoke(self):
        with self.serve_api() as port:
            for invalid in ({"id": "unknown", "accepted": True}, {"id": "item_1", "accepted": "true"}):
                status, data = self.post(port, "/api/manual/confirm", invalid)
                self.assertEqual((status, data["ok"]), (400, False))
            status, data = self.post(port, "/api/manual/confirm", {"id": "item_1", "accepted": True})
            self.assertEqual((status, data["ok"]), (200, True))
            self.assertEqual(self.snapshot("workflow")["items"]["item_1"]["status"], "accepted")
            status, data = self.post(port, "/api/manual/confirm", {"id": "item_1", "accepted": False})
            self.assertEqual((status, data["ok"]), (200, True))
            self.assertEqual(self.snapshot("workflow")["items"]["item_1"]["status"], "pending")

    def test_api_rejects_invalid_loop_limits_and_steps(self):
        with self.serve_api() as port:
            for value in (0, -1, 1.5, True, "6", None):
                with self.subTest(max_iter=value):
                    status, data = self.post(port, "/api/fix", {
                        "selected_ids": ["item_1"], "max_iter": value,
                    })
                    self.assertEqual((status, data["ok"]), (400, False))
            for value in (0, -1, 31, "NaN", "Infinity"):
                with self.subTest(step_pt=value):
                    status, data = self.post(port, "/api/fix", {
                        "selected_ids": ["item_1"], "step_pt": value,
                    })
                    self.assertEqual((status, data["ok"]), (400, False))
            self.assertFalse(self.workflow.STATE["busy"])

    def test_scan_resume_skips_completed_workbooks(self):
        second = self.folder / "second.xlsx"
        second.write_bytes(b"second workbook")
        self.workflow.STATE["pages"] = {}
        self.workflow.STATE["items"] = {}
        self.workflow.queue_job("scan", {
            "folder": str(self.folder), "delete_backups": False,
        })
        job = self.workflow.STATE["job"]
        job["workbooks"] = [str(self.wb_path), str(second)]
        job["completed"] = [str(self.wb_path)]
        self.reporter.init_execution_batch(
            "scan", self.folder, [self.wb_path, second],
        )
        self.wb.Sheets = []
        self.workflow.run_pending_workflow()
        self.excel.Workbooks.Open.assert_called_once_with(str(second.resolve()))
        self.assertEqual(set(job["completed"]), {str(self.wb_path), str(second)})
        self.assertEqual(job["status"], "completed")
        self.assertEqual(
            self.snapshot("workflow")["checked_files"][str(second)],
            self.workflow._file_version(second),
        )


if __name__ == "__main__":
    unittest.main()
