import copy
import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, storage, workflow


class OpenFilesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        patcher = patch.object(storage, "DATABASE_PATH", self.root / "state.sqlite3")
        patcher.start()
        self.addCleanup(patcher.stop)
        self.wf = importlib.reload(workflow)
        self.wf.STATE.update(folder=str(self.root), items={}, pages={}, busy=False, job=None,
                             batch_review={"cards": {}, "rows": {}, "ignored": []})
        self.cards = []
        for i in range(80):
            card = {"id": str(i), "seq": i, "wb_path": str(self.root / f"book-{i}.xlsx"),
                    "wb_name": f"book-{i}.xlsx", "sheet_name": "Sheet1", "cell_addr": "A1", "excel_row": 1,
                    "color": "yellow", "finalized": True}
            self.cards.append(card)
        batch._data()["cards"] = {c["id"]: c for c in self.cards}
        self.ids = [c["id"] for c in self.cards]

    def test_default_open_is_fixed_to_five_unique_files(self):
        duplicate = {**self.cards[0], "id": "duplicate", "seq": 100, "cell_addr": "A2"}
        batch._data()["cards"][duplicate["id"]] = duplicate
        plan = batch.open_plan(self.ids + [duplicate["id"]])
        self.assertEqual((plan["total"], len(plan["files"]), plan["remaining"]), (80, 5, 75))
        self.assertEqual(plan["card_ids"], self.ids[:5])

    def test_explicitly_selected_files_include_all_colors_and_are_not_truncated(self):
        self.cards[0]["color"] = "blue"
        self.cards[1]["color"] = "green"
        ids = batch.validate_cards(self.ids[:9], yellow=False)
        self.assertEqual(len(batch.open_plan(ids, selected_files=True)["files"]), 9)

    def test_waiting_files_do_not_consume_the_five_available_slots(self):
        self.wf.STATE.update(busy=True, job={"kind": "batch_scan"})
        with patch.object(batch, "open_block_reason", side_effect=lambda _job, path:
                          "还在处理" if path == Path(self.cards[0]["wb_path"]) else ""):
            plan = batch.open_plan(self.ids)
        self.assertEqual(plan["card_ids"], self.ids[1:6])
        self.assertEqual(len(plan["waiting"]), 1)
        self.assertEqual(plan["remaining"], 75)

    def test_busy_open_opens_only_five_and_selected_open_uses_the_selection(self):
        with patch.object(batch, "open_workbooks_for_review", return_value={}) as opener, \
                patch.object(batch.autofix, "wait_for_writes"):
            result = batch.open_ready(self.ids)
            self.assertEqual(len(opener.call_args.args[0]), 5)
            self.assertEqual(len(result["opened"]), 5)
            result = batch.open_ready(self.ids[:8], selected_files=True)
            self.assertEqual(len(opener.call_args.args[0]), 8)
            self.assertEqual(len(result["opened"]), 8)

    def test_worker_enforces_fixed_limit_even_when_old_options_request_eighty(self):
        for selected, expected in ((False, 5), (True, 8)):
            with self.subTest(selected=selected):
                ids = self.ids[:8] if selected else self.ids
                self.wf.STATE["job"] = {"kind": "batch_open", "completed": [], "pending_adjustment": None,
                                         "options": {"card_ids": ids, "selected_files": selected, "open_limit": 80}}
                with patch.object(batch.autofix, "_ExcelSession", return_value=Mock()), \
                        patch.object(batch, "ExcelWatchdog", return_value=Mock()), \
                        patch.object(batch, "open_workbooks_for_review", return_value={}) as opener, \
                        patch.object(batch, "finish_execution_batch"):
                    result = batch.run()
                self.assertEqual(len(opener.call_args.args[0]), expected)
                self.assertEqual(len(result), expected)
                self.assertEqual(len(self.wf.STATE["job"]["workbooks"]), expected)

    def test_open_plan_does_not_modify_review_colors_or_selection_data(self):
        before = copy.deepcopy(batch._data()["cards"])
        batch.open_plan(self.ids, selected_files=True)
        self.assertEqual(batch._data()["cards"], before)

    def handler(self, path):
        import web_clip_fixer
        web_clip_fixer = importlib.reload(web_clip_fixer)
        handler = object.__new__(web_clip_fixer.WebHandler)
        handler.path = path
        handler._send_json = Mock()
        handler._start_job = Mock()
        return handler

    def test_http_default_plan_and_open_ignore_legacy_configurable_limit(self):
        plan = self.handler('/api/batch/open_plan')
        plan._handle_post({"card_ids": self.ids, "open_limit": 80})
        self.assertEqual(len(plan._send_json.call_args.args[0]["files"]), 5)
        opened = self.handler('/api/batch/open')
        opened._handle_post({"card_ids": self.ids, "open_limit": 80})
        self.assertEqual(opened._start_job.call_args.args[0], "batch_open")
        self.assertEqual(len(opened._start_job.call_args.args[1]["card_ids"]), 5)
        batch._remember_opened([c["wb_path"] for c in self.cards[:5]])
        opened._handle_post({"card_ids": self.ids})
        self.assertEqual(opened._start_job.call_args.args[1]["card_ids"], self.ids[5:10])

    def test_http_selected_open_accepts_mixed_colors_and_keeps_all_chosen_files(self):
        batch._remember_opened([c["wb_path"] for c in self.cards[:5]])
        self.cards[0]["color"] = "blue"
        self.cards[1]["color"] = "green"
        handler = self.handler('/api/batch/open')
        handler._handle_post({"card_ids": self.ids[:8], "selected_files": True})
        self.assertTrue(handler._start_job.call_args.args[1]["selected_files"])
        self.assertEqual(len(handler._start_job.call_args.args[1]["card_ids"]), 8)

    def test_sixteen_clicks_open_eighty_different_files_without_wrapping(self):
        batches = []
        with patch.object(batch, "open_workbooks_for_review", side_effect=lambda targets: batches.append(targets) or {}), \
                patch.object(batch.autofix, "wait_for_writes"):
            for _ in range(16):
                self.assertEqual(len(batch.open_ready(self.ids)["opened"]), 5)
            result = batch.open_ready(self.ids)
        paths = [str(path) for targets in batches for path, *_ in targets]
        self.assertEqual(len(paths), 80)
        self.assertEqual(len(set(paths)), 80)
        self.assertEqual(result["opened"], [])
        self.assertEqual(len(batches), 16)
        self.assertTrue(batch.open_plan(self.ids)["exhausted"])
        self.assertEqual(batch.open_plan(self.ids)["remaining"], 0)

    def test_failed_files_remain_eligible_while_successes_are_skipped(self):
        errors = {Path(self.cards[0]["wb_path"]): "打不开", Path(self.cards[3]["wb_path"]): "被占用"}
        with patch.object(batch, "open_workbooks_for_review", return_value=errors), \
                patch.object(batch.autofix, "wait_for_writes"):
            result = batch.open_ready(self.ids)
        self.assertEqual(len(batch._opened_paths()), 3)
        self.assertEqual(result["remaining"], 77)
        plan = batch.open_plan(self.ids)
        self.assertEqual(plan["card_ids"], ["0", "3", "5", "6", "7"])
        self.assertEqual((plan["already_opened"], plan["unopened"]), (3, 77))

    def test_plan_is_read_only_and_service_restart_keeps_progress(self):
        self.assertEqual(batch.open_plan(self.ids)["card_ids"], self.ids[:5])
        self.assertEqual(batch._opened_paths(), set())
        batch._remember_opened([c["wb_path"] for c in self.cards[:5]])
        importlib.reload(workflow)
        self.assertEqual(batch.open_plan(self.ids)["card_ids"], self.ids[5:10])
        saved = storage.load_snapshot("workflow")
        self.assertEqual(len(saved["batch_review"]["open_history"][batch._open_folder_key()]), 5)

    def test_selected_files_can_reopen_seen_files_and_update_the_shared_history(self):
        batch._remember_opened([c["wb_path"] for c in self.cards[:5]])
        with patch.object(batch, "open_workbooks_for_review", return_value={}) as opener, \
                patch.object(batch.autofix, "wait_for_writes"):
            batch.open_ready(self.ids[:8], selected_files=True)
        self.assertEqual(len(opener.call_args.args[0]), 8)
        self.assertEqual(batch.open_plan(self.ids)["card_ids"], self.ids[8:13])

    def test_reset_is_scoped_to_the_current_folder(self):
        batch._remember_opened([c["wb_path"] for c in self.cards[:5]])
        self.wf.STATE["folder"] = str(self.root / "other")
        batch._remember_opened([str(self.root / "other" / "new.xlsx")])
        self.assertEqual(batch.reset_open_history(), 1)
        self.assertFalse(batch._opened_paths())
        self.wf.STATE["folder"] = str(self.root)
        self.assertEqual(len(batch._opened_paths()), 5)
        self.assertEqual(batch.reset_open_history(), 5)
        self.assertEqual(batch.open_plan(self.ids)["card_ids"], self.ids[:5])

    def test_failed_history_save_and_failed_reset_roll_back_in_memory(self):
        first = [c["wb_path"] for c in self.cards[:5]]
        batch._remember_opened(first)
        with patch.object(self.wf, "_persist_state", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                batch._remember_opened([self.cards[5]["wb_path"]])
            self.assertEqual(batch._opened_paths(), set(first))
            with self.assertRaises(OSError):
                batch.reset_open_history()
            self.assertEqual(batch._opened_paths(), set(first))

    def test_simultaneous_open_or_reset_is_rejected_without_changing_progress(self):
        with batch._opening_files(), patch.object(batch, "open_workbooks_for_review") as opener:
            with self.assertRaisesRegex(ValueError, "正在打开文件"):
                batch.open_ready(self.ids)
            with self.assertRaisesRegex(ValueError, "正在打开文件"):
                batch.reset_open_history()
            opener.assert_not_called()

    def test_new_scan_round_does_not_wait_for_or_fail_on_an_ongoing_open(self):
        batch._remember_opened([c["wb_path"] for c in self.cards[:5]])
        with batch._opening_files():
            batch.start_new_open_round()
        self.assertFalse(batch._opened_paths())

    def test_open_finishing_after_a_new_scan_round_is_not_recorded_in_the_new_round(self):
        def opener(targets):
            batch.start_new_open_round()  # 打开途中开始了新扫描
            return {}
        with patch.object(batch, "open_workbooks_for_review", side_effect=opener), \
                patch.object(batch.autofix, "wait_for_writes"):
            self.assertEqual(len(batch.open_ready(self.ids)["opened"]), 5)
        self.assertFalse(batch._opened_paths())
        self.assertEqual(batch.open_plan(self.ids)["card_ids"], self.ids[:5])

    def test_worker_limit_skips_files_already_opened_elsewhere(self):
        batch._remember_opened([c["wb_path"] for c in self.cards[1:3]])
        self.wf.STATE["job"] = {"kind": "batch_open", "completed": [], "pending_adjustment": None,
                                 "options": {"card_ids": self.ids[:10]}}
        with patch.object(batch.autofix, "_ExcelSession", return_value=Mock()), \
                patch.object(batch, "ExcelWatchdog", return_value=Mock()), \
                patch.object(batch, "open_workbooks_for_review", return_value={}) as opener, \
                patch.object(batch, "finish_execution_batch"):
            batch.run()
        opened = [str(t[0]) for t in opener.call_args.args[0]]
        self.assertEqual(opened, [self.cards[i]["wb_path"] for i in (0, 3, 4, 5, 6)])

    def test_worker_resume_retries_only_the_failed_file(self):
        self.wf.STATE["job"] = {"kind": "batch_open", "completed": [], "pending_adjustment": None,
                                 "options": {"card_ids": self.ids[:5]}}
        failed = Path(self.cards[3]["wb_path"])
        with patch.object(batch.autofix, "_ExcelSession", return_value=Mock()), \
                patch.object(batch, "ExcelWatchdog", return_value=Mock()), \
                patch.object(batch, "open_workbooks_for_review", side_effect=[{failed: "暂时失败"}, {}]) as opener, \
                patch.object(batch, "finish_execution_batch"):
            batch.run()
            batch.run()
        self.assertEqual([t[0] for t in opener.call_args.args[0]], [failed])
        self.assertEqual(len(batch._opened_paths()), 5)
        self.assertEqual(batch.open_plan(self.ids)["card_ids"], self.ids[5:10])

    def test_new_scan_resets_the_round_but_resume_preserves_it(self):
        batch._remember_opened([c["wb_path"] for c in self.cards[:5]])
        self.wf.STATE["job"] = {"kind": "batch_scan", "completed": [], "pending_adjustment": None,
                                 "options": {"folder": str(self.root), "scan_only": True}}
        with patch.object(batch.autofix, "_ExcelSession", return_value=Mock()), \
                patch.object(batch, "ExcelWatchdog", return_value=Mock()), \
                patch.object(self.wf, "list_workbooks", return_value=(self.root, [])), \
                patch.object(batch, "init_execution_batch"), patch.object(batch, "finish_execution_batch"):
            batch.run()
            self.assertFalse(batch._opened_paths())
            batch._remember_opened([c["wb_path"] for c in self.cards[:5]])
            batch.run()
            self.assertEqual(len(batch._opened_paths()), 5)

    def test_http_reset_does_not_change_card_colors(self):
        batch._remember_opened([c["wb_path"] for c in self.cards[:5]])
        before = copy.deepcopy(batch._data()["cards"])
        handler = self.handler('/api/batch/open_reset')
        handler._handle_post({})
        self.assertEqual(handler._send_json.call_args.args[0], {"ok": True, "reset": 5})
        self.assertEqual(batch._data()["cards"], before)


if __name__ == "__main__":
    unittest.main()
