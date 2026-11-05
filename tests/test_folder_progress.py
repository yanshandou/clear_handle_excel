import copy
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

from core import batch, fix_images, folder_progress, reporter, storage, workflow as wf, work_task


class FolderProgressTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for obj, field, value in ((storage, "DATABASE_PATH", self.root / "test.sqlite3"),
                                  (reporter, "RECORD_PATH", self.root / "record.md"),
                                  (fix_images, "IMAGES_DIR", self.root / "images")):
            p = patch.object(obj, field, value)
            p.start()
            self.addCleanup(p.stop)
        self.previous = copy.deepcopy(wf.STATE)
        self.previous_report = copy.deepcopy(reporter.EXECUTION_STATE)
        self.addCleanup(self.restore_globals)
        wf.STATE.clear()
        wf.STATE.update(copy.deepcopy(wf._EMPTY_STATE))
        reporter.EXECUTION_STATE.clear()
        reporter.EXECUTION_STATE.update(reporter.empty_execution_report(""))
        self.a, self.b = self.root / "A", self.root / "B"
        self.a.mkdir()
        self.b.mkdir()
        self.path = str(self.a / "book.xlsx")
        Path(self.path).write_bytes(b"unchanged")
        wf.activate_folder(str(self.a))

    def restore_globals(self):
        wf.STATE.clear()
        wf.STATE.update(self.previous)
        reporter.EXECUTION_STATE.clear()
        reporter.EXECUTION_STATE.update(self.previous_report)

    def populate(self, status="completed"):
        state = wf.STATE
        state["items"] = {"a": {"id": "a", "wb_path": self.path, "image": "data:image/png;base64,YQ==",
                                  "status": "pending", "fix_images": {"before": "before.png"}}}
        state["pages"] = {"page_a": {"wb_path": self.path, "image": "page image", "batch_context_image": "page.png"}}
        state["checked_files"] = {self.path: wf._file_version(Path(self.path))}
        state["job"] = {"kind": "batch_scan", "status": status, "options": {"folder": str(self.a)},
                        "completed": [self.path], "scanned": [self.path], "workbooks": [self.path, str(self.a / 'later.xlsx')],
                        "scan_versions": {self.path: state["checked_files"][self.path]},
                        "pending_adjustment": None, "pause_requested": status == "interrupted"}
        state["batch_review"] = {
            "cards": {"a": {"id": "a", "wb_path": self.path, "color": "green", "finalized": True,
                              "context": {"image_file": "context.png"}}},
            "rows": {"row": {"path": self.path, "base": 18, "saved": True}},
            "ignored": [{"cell": [self.path, "Sheet1", "A1"], "signatures": []}],
            "highlights": {"a": {"path": self.path, "original": {"fill": 123, "font": 456}}},
            "tab_marks": {self.path: {"Sheet1": {"color": 77}}},
            "sheet_checks": {self.path: {"*": state["checked_files"][self.path]}},
            "open_history": {folder_progress.key(str(self.a)): [self.path]},
            "slow_files": {self.path: {"sheet": "Sheet1"}}, "followups": []}
        state.update(logs=["A log"], scan_progress={"total": 2, "completed": 1},
                     latest_scan_analysis={"id": "A-analysis", "folder": str(self.a)},
                     sheet_history={self.path: ["history"]}, print_layouts={self.path: {"sheet": "layout"}})
        report = reporter.empty_execution_report(str(self.a))
        report.update(task_name="A 扫描", status="completed", batch=["book.xlsx"],
                      files={"book.xlsx": {"name": "book.xlsx", "status": "needs_fix"}})
        reporter.replace_execution_report(report)
        wf._persist_state()
        return copy.deepcopy(state)

    def test_round_trip_keeps_all_state_without_scanning_or_touching_excel(self):
        expected = self.populate()
        before = Path(self.path).stat().st_mtime_ns
        with patch.object(wf, "_scan_one_workbook", side_effect=AssertionError("must not scan")), \
                patch.object(wf, "start_background_excel", side_effect=AssertionError("must not open Excel")):
            wf.activate_folder(str(self.b))
            self.assertEqual(wf.STATE["items"], {})
            self.assertEqual(wf.STATE["batch_review"]["cards"], {})
            self.assertEqual(reporter.EXECUTION_STATE["files"], {})
            self.assertNotIn("batch", reporter.EXECUTION_STATE)
            wf.STATE["logs"] = ["B log"]
            wf._persist_state()
            wf.activate_folder(str(self.a))
        for field in ("items", "pages", "batch_review", "checked_files", "job", "logs", "scan_progress",
                      "latest_scan_analysis", "sheet_history", "print_layouts"):
            self.assertEqual(wf.STATE[field], expected[field], field)
        self.assertEqual(reporter.EXECUTION_STATE["task_name"], "A 扫描")
        self.assertEqual(Path(self.path).stat().st_mtime_ns, before)
        self.assertTrue(batch._unchanged_since_check(self.path))
        wf.activate_folder(str(self.b))
        self.assertEqual(wf.STATE["logs"], ["B log"])

    def test_paused_scan_survives_switch_and_restart(self):
        self.populate("interrupted")
        wf.activate_folder(str(self.b))
        # 清空内存后从活动快照恢复，模拟重启，不靠进程内缓存。
        active = storage.load_snapshot("workflow")
        wf.STATE.clear()
        wf.STATE.update(active)
        wf.activate_folder(str(self.a))
        self.assertEqual(wf.STATE["job"]["status"], "interrupted")
        self.assertEqual(wf.STATE["job"]["scanned"], [self.path])
        self.assertTrue(wf.queue_job("resume"))
        self.assertNotIn("pause_requested", wf.STATE["job"])
        self.assertEqual(wf.STATE["job"]["completed"], [self.path])

    def test_busy_or_maintenance_or_opening_cannot_change_folder(self):
        for fields in ({"busy": True}, {"maintenance": "color_scan"}):
            wf.STATE.update(fields)
            with self.assertRaisesRegex(ValueError, "暂停"):
                wf.activate_folder(str(self.b))
            self.assertEqual(wf.STATE["folder"], str(self.a))
            wf.STATE.update(busy=False)
            wf.STATE.pop("maintenance", None)
        with batch._opening_files():
            with self.assertRaisesRegex(ValueError, "打开"):
                wf.activate_folder(str(self.b))

    def test_archived_screenshots_survive_cleanup(self):
        self.populate()
        fix_images.IMAGES_DIR.mkdir()
        for name in ("before.png", "context.png", "page.png", "unused.png"):
            (fix_images.IMAGES_DIR / name).write_bytes(b"png")
        wf.activate_folder(str(self.b))
        fix_images.clear()
        self.assertEqual({p.name for p in fix_images.IMAGES_DIR.iterdir()}, {"before.png", "context.png", "page.png"})

    def test_legacy_mismatched_report_is_never_shown_under_wrong_folder(self):
        self.populate()
        report = reporter.empty_execution_report(str(self.b))
        report["task_name"] = "B 的报告"
        reporter.replace_execution_report(report)
        self.assertNotEqual(wf.folder_history(str(self.a))["report"]["task_name"], "B 的报告")
        wf.activate_folder(str(self.b))
        self.assertNotEqual(wf.folder_history(str(self.a))["report"]["task_name"], "B 的报告")

    def test_legacy_records_are_split_for_other_folders(self):
        old = str(self.b / "old.xlsx")
        wf.STATE.pop("folder_progress_version", None)
        # 模拟真正升级前：去掉 setUp 创建的空迁移快照。
        with closing(storage._connect()) as db, db:
            db.execute("DELETE FROM snapshots WHERE name=?", (folder_progress.LEGACY,))
        self.populate()
        wf.STATE.pop("folder_progress_version", None)
        wf.STATE["batch_review"]["cards"]["old"] = {"wb_path": old, "color": "blue"}
        wf.STATE["batch_review"]["highlights"]["old"] = {"path": old, "original": 789}
        wf.STATE["batch_review"]["ignored"].append({"cell": [old, "S", "A1"]})
        wf.activate_folder(str(self.b))
        self.assertEqual(set(wf.STATE["batch_review"]["cards"]), {"old"})
        self.assertEqual(wf.STATE["batch_review"]["highlights"]["old"]["original"], 789)
        self.assertIsNone(wf.STATE["job"])
        self.assertEqual(wf.STATE["batch_review"]["ignored"], [{"cell": [old, "S", "A1"]}])
        wf.activate_folder(str(self.a))
        self.assertEqual(set(wf.STATE["batch_review"]["cards"]), {"a"})

    def test_reading_history_never_switches_or_exports_pdf(self):
        self.populate()
        wf.activate_folder(str(self.b))
        before = copy.deepcopy(wf.STATE)
        history = wf.folder_history()
        self.assertEqual({r["folder"] for r in history["folders"]}, {str(self.a), str(self.b)})
        detail = wf.folder_history(str(self.a))
        self.assertEqual(detail["logs"], ["A log"])
        self.assertEqual(detail["report"]["task_name"], "A 扫描")
        self.assertNotIn("markdown", detail["report"])
        self.assertEqual(wf.STATE, before)

    def test_failed_restore_retains_source_and_both_archives(self):
        self.populate()
        wf.activate_folder(str(self.b))
        wf.STATE["logs"] = ["B log"]
        actual = wf._persist_state
        def fail_destination(*args, **kwargs):
            if wf.STATE["folder"] == str(self.a):
                raise OSError("disk error")
            return actual(*args, **kwargs)
        with patch.object(wf, "_persist_state", side_effect=fail_destination):
            with self.assertRaisesRegex(OSError, "disk error"):
                wf.activate_folder(str(self.a))
        self.assertEqual(wf.STATE["folder"], str(self.b))
        self.assertEqual(wf.STATE["logs"], ["B log"])
        self.assertEqual(storage.load_snapshot("workflow")["folder"], str(self.b))
        wf.activate_folder(str(self.a))
        self.assertEqual(wf.STATE["logs"], ["A log"])

    def test_task_switch_restores_progress_and_busy_failure_keeps_current_task(self):
        shared = self.root / "shared"
        shared.mkdir()
        task = work_task.create(str(shared), str(self.root / "work"))
        wf.STATE["logs"] = ["task one"]
        shared2 = self.root / "shared2"
        shared2.mkdir()
        other = work_task.create(str(shared2), str(self.root / "work"))
        wf.STATE["busy"] = True
        with self.assertRaises(ValueError):
            work_task.switch(task["id"])
        self.assertEqual(work_task.overview()["current"], other["id"])
        wf.STATE["busy"] = False
        work_task.switch(task["id"])
        self.assertEqual(wf.STATE["logs"], ["task one"])

    def test_folder_spelling_and_trailing_separator_use_same_archive(self):
        self.populate()
        wf.activate_folder(str(self.b))
        wf.activate_folder(str(self.a).upper() + "/")
        self.assertEqual(wf.STATE["logs"], ["A log"])

    def test_task_metadata_failure_rolls_back_workbench(self):
        self.populate()
        shared = self.root / "shared"
        shared.mkdir()
        with patch.object(work_task, "_save", side_effect=OSError("metadata disk error")):
            with self.assertRaisesRegex(OSError, "metadata disk error"):
                work_task.create(str(shared), str(self.root / "work"))
        self.assertEqual(wf.STATE["folder"], str(self.a))
        self.assertEqual(wf.STATE["logs"], ["A log"])
        self.assertEqual(reporter.EXECUTION_STATE["task_name"], "A 扫描")

    def test_history_routes_are_read_only_and_restore_route_does_not_queue_scan(self):
        import web_clip_fixer as server
        self.populate()
        wf.activate_folder(str(self.b))
        handler = object.__new__(server.WebHandler)
        handler._send_json = lambda value: setattr(handler, "response", value)
        handler._serve_static_file = lambda name: setattr(handler, "page", name)
        handler.path = "/history"
        handler._handle_get()
        self.assertEqual(handler.page, "history.html")
        handler.path = "/api/folders"
        handler._handle_get()
        self.assertEqual(len(handler.response["folders"]), 2)
        from urllib.parse import quote
        handler.path = "/api/folders/detail?folder=" + quote(str(self.a))
        handler._handle_get()
        self.assertEqual(handler.response["report"]["task_name"], "A 扫描")
        self.assertEqual(wf.STATE["folder"], str(self.b))
        handler.path = "/api/folders/activate"
        with patch.object(server, "_require_folder", return_value=str(self.a)), \
                patch.object(wf, "queue_job", side_effect=AssertionError("must not queue scan")):
            handler._handle_post({"folder": str(self.a)})
        self.assertEqual(wf.STATE["logs"], ["A log"])


if __name__ == "__main__":
    unittest.main()
