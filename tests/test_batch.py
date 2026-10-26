import copy
import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, fix_images, reporter, storage, workflow


class BatchTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for obj, name, value in (
            (storage, "DATABASE_PATH", self.root / "state.sqlite3"),
            (reporter, "RECORD_PATH", self.root / "record.md"),
            (fix_images, "IMAGES_DIR", self.root / "images"),
        ):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.wf = importlib.reload(workflow)
        self.path = str(self.root / "book.xlsx")
        Path(self.path).write_bytes(b"workbook")
        self.wf.STATE.update(folder=str(self.root), items={}, pages={},
                             batch_review={"cards": {}, "rows": {}, "ignored": []},
                             job={"kind": "batch_scan", "options": {}, "completed": [], "pending_adjustment": None})
        self.heights = {5: 24.0, 6: 27.0}
        self.book = Mock(ReadOnly=False)
        def row_at(row):
            class Row:
                @property
                def RowHeight(_):
                    return self.heights[row]
                @RowHeight.setter
                def RowHeight(_, value):
                    self.heights[row] = value
            return Row()
        self.book.Sheets.return_value.Rows.side_effect = row_at
        self.excel = Mock()
        self.excel.Workbooks.Open.return_value = self.book
        self.normalize = patch.object(batch, "normalize_workbook_view").start()
        self.real_refresh = batch._refresh
        self.refresh = patch.object(batch, "_refresh").start()
        self.addCleanup(patch.stopall)

    def test_empty_open_keeps_write_journal_and_original_height(self):
        card, _ = self.add_card()
        pending = {"action": "fix", "path": self.path, "version": self.wf._file_version(Path(self.path)),
                   "cards": [card["id"]], "rows": [
                       {"id": batch._key(self.path, "S1", 5), "sheet": "S1", "row": 5,
                        "expected": 24, "target": 30}]}
        self.wf.STATE["job"]["pending_adjustment"] = pending
        self.excel.Workbooks.Open.return_value = None
        with self.assertRaisesRegex(RuntimeError, "Open 返回空对象"):
            batch._finish_pending(self.excel)
        self.assertIs(self.wf.STATE["job"]["pending_adjustment"], pending)
        self.assertEqual(self.heights[5], 24)
        self.book.Save.assert_not_called()

    def test_same_named_trial_and_original_use_separate_excel_instances(self):
        from core import autofix
        trial_app, trial_book = Mock(), Mock()
        original_open = self.excel.Workbooks.Open
        trial_app.Workbooks.Open.return_value = trial_book
        # 模拟 Excel：一个实例内不能同时打开不同目录的同名工作簿。
        def open_original(path, **options):
            if path != self.path:
                return None
            return self.book
        original_open.side_effect = open_original
        session = autofix._ExcelSession()
        with patch.object(autofix, "start_background_excel", side_effect=[self.excel, trial_app]), \
                patch.object(autofix, "quit_background_excel") as quit_excel:
            try:
                self.assertIs(session.copy_of(self.path), trial_book)
                self.assertEqual(Path(trial_app.Workbooks.Open.call_args.args[0]).name, Path(self.path).name)
                batch._write_rows(session.app(), {
                    "path": self.path, "version": self.wf._file_version(Path(self.path)),
                    "rows": [{"sheet": "S1", "row": 5, "expected": 24, "target": 30}]})
                self.assertEqual(self.heights[5], 30)
                self.book.Save.assert_called_once()
                trial_book.Close.assert_not_called()
                self.assertIs(session.copy_of(self.path), trial_book)
            finally:
                session.close()
            self.assertEqual(quit_excel.call_count, 2)
            self.assertIsNone(session.copy_excel)
            self.assertIsNone(session.excel)

    def test_failed_trial_open_is_not_cached(self):
        from core import autofix
        trial_app = Mock()
        trial_app.Workbooks.Open.return_value = None
        session = autofix._ExcelSession()
        with patch.object(autofix, "start_background_excel", side_effect=[self.excel, trial_app]), \
                patch.object(autofix, "quit_background_excel"):
            try:
                with self.assertRaisesRegex(RuntimeError, "Open 返回空对象"):
                    session.copy_of(self.path)
                self.assertNotIn(self.path, session.copies)
            finally:
                session.close()

    def add_card(self, address="A5", color="blue", row=5):
        iid = f"i{address}"
        item = {"id": iid, "seq": len(self.wf.STATE["items"]) + 1, "kind": "hidden", "text": address,
                "wb_path": self.path, "wb_name": "book.xlsx", "sheet_name": "S1", "cell_addr": address,
                "excel_row": row, "row_height": self.heights[row], "status": "fixed", "image": ""}
        item.update(page_index=0, bbox=[10, 10, 100, 25])
        self.wf.STATE["items"][iid] = item
        card = batch._card_for([item])
        card.update(color=color, verified=True, waiting=False)
        rid = batch._key(self.path, "S1", row)
        batch._data()["rows"][rid] = {"id": rid, "path": self.path, "sheet": "S1", "row": row,
                                       "base": 15.0, "applied": self.heights[row], "saved": True}
        return card, item

    def commit(self, cards):
        self.wf.STATE["job"].update(kind="batch_commit", options={"card_ids": [c["id"] for c in cards]})
        batch._commit_review(self.excel)

    def test_commit_by_color_keeps_other_colors_and_mixed_rows(self):
        """筛选黄色只提交黄卡：别的颜色不动；同行混着蓝卡（共用已写入的行高）的整行先不交，不连带转黄"""
        mixed_yellow, _ = self.add_card("A5", color="yellow")
        mixed_blue, _ = self.add_card("B5", color="blue")
        alone, _ = self.add_card("A6", color="yellow", row=6)
        self.heights[7] = 30.0
        other, _ = self.add_card("A7", color="green", row=7)
        self.wf.STATE["job"].update(kind="batch_commit", options={
            "card_ids": [mixed_yellow["id"], alone["id"]], "color": "yellow"})
        with patch.object(batch, "_finish_pending"), patch.object(batch, "_refresh"):
            batch._commit_review(self.excel)
        self.assertEqual(self.wf.STATE["job"]["options"]["commit_card_ids"], [alone["id"]])
        self.assertTrue(alone["finalized"] and alone["manual"])
        for card, color in ((mixed_yellow, "yellow"), (mixed_blue, "blue"), (other, "green")):
            self.assertEqual((card["color"], card.get("finalized", False), card.get("conflict", False)),
                             (color, False, False))

    def test_commit_selection_by_color_on_untouched_rows(self):
        """行高没写入过的行，各格互不影响：只挑出这种颜色的卡"""
        yellow, _ = self.add_card("A5", color="yellow")
        blue, _ = self.add_card("B5", color="blue")
        batch._data()["rows"][batch._key(self.path, "S1", 5)]["saved"] = False
        self.assertEqual(batch.commit_selection([yellow["id"], blue["id"]], "yellow"), ([yellow["id"]], []))
        self.assertEqual(batch.commit_selection([yellow["id"]], "blue"), ([blue["id"]], []))
        self.assertEqual(sorted(batch.commit_selection([yellow["id"]])[0]), sorted([yellow["id"], blue["id"]]))

    def test_commit_skips_cards_not_yet_processed(self):
        """暂停后提交：还没轮到的卡不能被转成人工项，否则以后全自动不会再修它们"""
        done, _ = self.add_card("A5", color="blue")
        _, waiting_item = self.add_card("A6", row=6)
        batch._data()["rows"].pop(batch._key(self.path, "S1", 6))
        waiting_item["status"] = "pending"
        waiting = batch._data()["cards"][batch._key(self.path, "S1", "A6")]
        waiting.update(color="yellow", verified=False, waiting=True, message=batch.WAITING_MESSAGE)
        self.commit([done, waiting])
        self.assertTrue(done["dismissed"])
        self.assertEqual((waiting["color"], waiting["finalized"], waiting["manual"]), ("yellow", False, False))
        self.assertEqual(waiting_item["status"], "pending")
        self.wf.STATE["job"].update(kind="batch_scan", attempted_rows=[])
        prepared = []
        def prepare(_session, _path, items, **_):
            prepared.append(items[0]["cell_addr"])
            return 27, None, {}
        with patch.object(batch, "_prepare_row", side_effect=prepare):
            batch._repair_file(self.excel, Mock(), self.path)
        self.assertEqual(prepared, ["A6"])
        self.assertFalse(waiting["waiting"])

    def test_old_waiting_cards_without_flag_are_recognised(self):
        card, _ = self.add_card()
        card.pop("waiting")
        card["message"] = batch.WAITING_MESSAGE
        self.assertTrue(batch._waiting(card))

    def test_paused_scan_survives_review_jobs(self):
        scan = {"kind": "batch_scan", "options": {"folder": str(self.root)}, "completed": ["x"],
                "pending_adjustment": None, "status": "interrupted", "pause_requested": True}
        self.wf.STATE.update(busy=False, job=scan)
        card, _ = self.add_card(color="blue")
        self.assertTrue(self.wf.queue_job("batch_commit", {"card_ids": [card["id"]]}))
        self.assertEqual(self.wf.STATE["job"]["paused_scan"]["completed"], ["x"])
        with patch.object(batch, "run", side_effect=lambda: []):
            self.wf.run_pending_workflow()
        job = self.wf.STATE["job"]
        self.assertEqual((job["kind"], job["status"], job["completed"]), ("batch_scan", "interrupted", ["x"]))
        self.assertTrue(self.wf.queue_job("resume"))
        self.assertNotIn("pause_requested", self.wf.STATE["job"])

    def test_column_width_problem_goes_manual_without_trial_exports(self):
        from core import autofix
        from core.hidden_text import REASON_HASH
        hashed, hash_item = self.add_card("A5")
        other, other_item = self.add_card("B5")
        batch._data()["rows"].clear()
        hash_item.update(status="pending", hidden_reason=REASON_HASH)
        other_item["status"] = "pending"
        prepared = []
        def prepare(_session, _path, items, **_):
            prepared.append([it["cell_addr"] for it in items])
            return 15, None, {}
        with patch.object(batch, "_prepare_row", side_effect=prepare):
            batch._repair_file(self.excel, Mock(), self.path)
        self.assertEqual(prepared, [["B5"]])  # 同行其他格照常试
        self.assertEqual((hashed["color"], hashed["manual"]), ("yellow", True))
        self.assertEqual(hashed["message"], autofix.COLUMN_ONLY_REASON)

    def test_only_column_width_problem_in_row_skips_copy(self):
        from core.hidden_text import REASON_HASH
        card, item = self.add_card("A5")
        batch._data()["rows"].clear()
        item.update(status="pending", hidden_reason=REASON_HASH)
        session = Mock()
        with patch.object(batch, "_prepare_row") as prepare:
            batch._repair_file(self.excel, session, self.path)
        prepare.assert_not_called()
        session.copy_of.assert_not_called()
        self.assertTrue(card["manual"])

    def test_refresh_keeps_users_green_mark_when_cell_breaks_again(self):
        card, item = self.add_card()
        batch.mark(card["id"], "green")
        item["status"] = "pending"
        with patch.object(self.wf, "_scan_one_workbook"):
            self.real_refresh(self.excel, self.path, ["S1"])
        self.assertEqual((card["color"], card["verified"]), ("green", False))
        self.assertIn("保留你的标记", card["message"])

    def test_recreated_card_does_not_overwrite_earlier_fix_images(self):
        card, item = self.add_card()
        batch._save_shot(card, "before", "data:image/png;base64,b2xk")
        item["fix_images"] = dict(card["fix_images"])
        card["dismissed"] = True
        again = batch._card_for([{**item, "id": "new", "image": "data:image/png;base64,bmV3"}])
        self.assertIsNot(again, card)
        self.assertNotEqual(again["fix_images"]["before"], item["fix_images"]["before"])
        self.assertEqual(fix_images.image_path(item["fix_images"]["before"]).read_bytes(), b"old")

    def test_report_reason_does_not_mention_manual_process(self):
        from core.hidden_text import REASON_MISSING
        reason = self.wf._fix_record({"kind": "hidden", "hidden_reason": REASON_MISSING, "sheet_name": "S1",
                                      "cell_addr": "A5", "excel_row": 5})["reason"]
        for word in ("人工", "核对", "复检", "PDF"):
            self.assertNotIn(word, reason)

    def test_mark_only_persists_color_and_never_opens_excel(self):
        card, _ = self.add_card()
        batch.mark(card["id"], "green")
        self.assertEqual(self.heights[5], 24)
        self.excel.Workbooks.Open.assert_not_called()
        self.assertEqual(storage.load_snapshot("workflow")["batch_review"]["cards"][card["id"]]["color"], "green")

    def test_green_restores_before_ignoring_and_survives_redetection(self):
        card, item = self.add_card(color="green")
        self.commit([card])
        self.assertEqual(self.heights[5], 15)
        self.assertTrue(card["dismissed"])
        self.assertEqual(item["status"], "accepted")
        redetected = copy.deepcopy(item)
        redetected.update(status="pending", id="new")
        batch.carry_item_policies({"new": redetected})
        self.assertEqual(redetected["status"], "accepted")
        for change in ({"cell_addr": "B5"}, {"text": "changed"}, {"sheet_name": "S2"}):
            other = {**redetected, **change, "status": "pending"}
            batch.carry_item_policies({"other": other})
            self.assertEqual(other["status"], "pending")

    def test_yellow_restores_and_remains_in_manual_queue(self):
        card, item = self.add_card(color="yellow")
        self.commit([card])
        self.assertEqual(self.heights[5], 15)
        self.assertFalse(card["dismissed"])
        self.assertTrue(card["manual"])
        self.assertEqual(item["status"], "manual_review")

    def test_blue_is_passed_without_saving_excel(self):
        card, item = self.add_card()
        self.commit([card])
        self.assertEqual(self.heights[5], 24)
        self.assertTrue(card["dismissed"])
        self.assertEqual(item["status"], "fixed")
        self.book.Save.assert_not_called()

    def test_same_row_conflict_restores_whole_row_and_keeps_both_yellow(self):
        a, _ = self.add_card("A5", "blue")
        b, _ = self.add_card("B5", "green")
        self.commit([a, b])
        self.assertEqual(self.heights[5], 15)
        self.assertEqual((a["color"], b["color"]), ("yellow", "yellow"))
        self.assertFalse(a["dismissed"] or b["dismissed"])
        self.assertEqual(batch._data()["ignored"], [])
        self.book.Save.assert_called_once()

    def test_partial_commit_cannot_restore_without_considering_row_peers(self):
        a, _ = self.add_card("A5", "blue")
        b, _ = self.add_card("B5", "green")
        self.commit([b])
        self.assertEqual((a["color"], b["color"]), ("yellow", "yellow"))

    def test_user_modified_height_is_not_overwritten(self):
        card, _ = self.add_card(color="yellow")
        self.heights[5] = 36
        self.commit([card])
        self.assertEqual(self.heights[5], 36)
        job = self.wf.STATE["job"]
        self.assertIn(self.path, job["deferred_adjustments"])
        self.assertIn("行高已被", card["message"])
        self.assertNotIn(self.path, job["completed"])
        self.assertEqual(card["color"], "yellow")
        self.assertFalse(card["finalized"])

    def test_resume_after_save_does_not_apply_adjustment_twice(self):
        card, _ = self.add_card(color="yellow")
        self.wf.STATE["job"]["pending_adjustment"] = {
            "path": self.path, "action": "restore", "version": self.wf._file_version(Path(self.path)),
            "cards": [card["id"]], "rows": [{"id": batch._key(self.path, "S1", 5), "sheet": "S1",
                                            "row": 5, "expected": 24, "target": 15}],
        }
        self.heights[5] = 15  # 相当于 Save 成功、数据库提交前退出
        batch._finish_pending(self.excel)
        self.book.Save.assert_not_called()
        self.assertIsNone(self.wf.STATE["job"]["pending_adjustment"])
        self.assertFalse(next(iter(batch._data()["rows"].values()))["saved"])

    def test_excel_pixel_rounding_is_not_treated_as_manual_change(self):
        """日文 Excel 默认 18.75pt：旧版按 0.1 取整记成 18.8，写入核对时误报「行高已被其他操作修改」"""
        card, _ = self.add_card(color="blue")
        self.heights[5] = 18.75
        self.wf.STATE["job"]["pending_adjustment"] = {
            "path": self.path, "action": "fix", "version": self.wf._file_version(Path(self.path)),
            "cards": [card["id"]], "rows": [{"id": batch._key(self.path, "S1", 5), "sheet": "S1",
                                            "row": 5, "expected": 18.8, "target": 21.75}],
        }
        batch._finish_pending(self.excel)
        self.assertEqual(self.heights[5], 21.75)
        self.book.Save.assert_called_once()
        self.assertIsNone(self.wf.STATE["job"]["pending_adjustment"])

    def test_all_heights_validated_before_any_change(self):
        a, _ = self.add_card("A5", "yellow")
        b, _ = self.add_card("A6", "yellow", row=6)
        self.heights[6] = 90
        self.commit([a, b])
        self.assertIn(self.path, self.wf.STATE["job"]["deferred_adjustments"])
        self.assertEqual(self.heights, {5: 24.0, 6: 90})
        self.book.Save.assert_not_called()

    def test_unsubmitted_yellow_cannot_be_rechecked(self):
        card, _ = self.add_card(color="yellow")
        with self.assertRaisesRegex(ValueError, "先提交"):
            batch.validate_cards([card["id"]], yellow=True)

    def test_mark_many_persists_once_and_allows_unverified_blue(self):
        a, _ = self.add_card("A5", "yellow")
        b, _ = self.add_card("B5", "green")
        a["verified"] = False
        with patch.object(self.wf, "_persist_state", wraps=self.wf._persist_state) as persist:
            result = batch.mark_many([a["id"], b["id"]], "blue")
        persist.assert_called_once()
        self.assertEqual(result, {"changed": 2, "skipped": 0})
        self.assertEqual((a["color"], b["color"]), ("blue", "blue"))
        self.assertTrue(b["marked"])
        batch.mark_many([a["id"], b["id"]], "green")
        self.assertEqual((a["color"], b["color"]), ("green", "green"))
        self.excel.Workbooks.Open.assert_not_called()

    def test_mark_many_rejects_stale_cards_without_changing_any(self):
        a, _ = self.add_card("A5", "yellow")
        with self.assertRaisesRegex(ValueError, "失效"):
            batch.mark_many([a["id"], "missing"], "green")
        self.assertEqual(a["color"], "yellow")

    def test_user_can_cycle_all_colors_without_verification_or_images(self):
        card, _ = self.add_card(color="yellow")
        card["verified"] = False
        for color in ["blue", "green", "yellow", "blue", "yellow", "green"]:
            batch.mark(card["id"], color)
            self.assertEqual(card["color"], color)
            self.assertFalse(card["verified"])
        self.excel.Workbooks.Open.assert_not_called()

    def test_unverified_blue_without_trial_passes_without_writing_or_reappearing(self):
        card, item = self.add_card(color="yellow")
        batch._data()["rows"].clear()
        card["verified"] = False
        item["status"] = "pending"
        batch.mark(card["id"], "blue")
        self.commit([card])
        self.assertTrue(card["dismissed"])
        self.assertEqual(item["status"], "accepted")
        redetected = dict(item, status="pending")
        batch.carry_item_policies({"new": redetected})
        self.assertEqual(redetected["status"], "accepted")
        self.excel.Workbooks.Open.assert_not_called()

    def attempt_card(self, address="A5", color="yellow"):
        """自动加高没通过检查、留下「最后一次尝试」截图、原文件没改的卡"""
        card, item = self.add_card(address, color)
        batch._data()["rows"].pop(batch._key(self.path, "S1", 5), None)
        item["status"] = "pending"
        card.update(verified=False, after_kind="attempt", base_height=24.0, attempt_height=33.0,
                    fix_images={"before": "b.png", "after": "a.png"})
        return card, item

    def test_blue_on_failed_attempt_writes_attempt_height_and_counts_as_fixed(self):
        """检测器没通过但人看图已修好：可以标蓝，提交时按最后一次尝试的行高写入，报告里算修复"""
        card, item = self.attempt_card()
        self.assertTrue(batch.snapshot()["cards"][0]["adoptable"])
        batch.mark(card["id"], "blue")
        self.commit([card])
        self.assertEqual(self.heights[5], 33.0)
        self.assertTrue(card["dismissed"])
        self.assertEqual((item["status"], item["row_height"]), ("fixed", 33.0))
        self.assertFalse(batch._data()["rows"][batch._key(self.path, "S1", 5)]["saved"])
        # 重新检测仍会报这一格（误报）：按规则记为已修复，不再出卡
        redetected = dict(item, id="new", status="pending", fix_images={})
        batch.carry_item_policies({"new": redetected})
        self.assertEqual((redetected["status"], redetected["row_height"]), ("fixed", 33.0))
        self.assertEqual(redetected["fix_images"]["after"], "a.png")

    def test_all_blue_commit_still_starts_excel_to_write_an_adopted_attempt(self):
        """全部标蓝提交、其中有「最后一次尝试」：以前不开 Excel，写行高时报 'NoneType' object has no attribute 'Workbooks'"""
        card, _ = self.attempt_card()
        batch.mark(card["id"], "blue")
        self.wf.STATE.update(busy=True)
        self.wf.STATE["job"] = {"kind": "batch_commit", "options": {"card_ids": [card["id"]]},
                                "completed": [], "pending_adjustment": None, "status": "running"}
        session = Mock()
        session.app.return_value = self.excel
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "_sync_highlights_safely"), \
                patch.object(batch, "finish_execution_batch"):
            batch.run()
        self.assertEqual(self.heights[5], 33.0)
        self.assertIsNone(self.wf.STATE["job"]["pending_adjustment"])
        self.assertTrue(card["dismissed"])

    def test_blue_attempt_with_disagreeing_row_peer_writes_nothing(self):
        card, _ = self.attempt_card("A5", "blue")
        peer, _ = self.attempt_card("B5", "green")
        self.commit([card, peer])
        self.assertEqual(self.heights[5], 24.0)
        self.assertEqual((card["color"], peer["color"]), ("yellow", "yellow"))
        self.excel.Workbooks.Open.assert_not_called()

    def test_waiting_card_can_be_marked_blue(self):
        card, _ = self.attempt_card()
        card.update(waiting=True, message=batch.WAITING_MESSAGE)
        self.assertEqual(batch.mark_many([card["id"]], "blue"), {"changed": 1, "skipped": 0})

    def test_recheck_passes_one_cell_and_keeps_ignored_cell_out(self):
        a, item_a = self.add_card("A5", "yellow")
        b, item_b = self.add_card("B5", "green")
        # 两处都已还原，不再有自动行高冲突。
        next(iter(batch._data()["rows"].values()))["saved"] = False
        self.commit([a, b])
        def refresh(*_, **__):
            item_a["status"] = "fixed"
            item_b["status"] = "pending"
            batch.carry_item_policies(self.wf.STATE["items"])
        with patch.object(batch, "_refresh", side_effect=refresh):
            batch._recheck_file(self.excel, self.path, [a])
        self.assertEqual(a["color"], "blue")
        self.assertTrue(a["verified"])
        self.assertTrue(b["dismissed"])
        self.assertEqual(item_b["status"], "accepted")

    def test_row_saved_then_sheet_refreshed_before_next_row(self):
        a, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        self.wf.STATE["job"].update(kind="batch_scan")
        events = []
        cid = a["id"]
        def prepare(*_, **__):
            events.append("prepare")
            return 15, {"height": 24, "shots": {cid: ""}}, {cid: ""}
        def finish(*_):
            events.append("save")
            item["status"] = "fixed"
            self.wf.STATE["job"]["pending_adjustment"] = None
        def refresh(_excel, _path, sheets, **kwargs):
            source = "copy" if kwargs.get("workbook") is not None else "original"
            events.append(f"fresh-{source}-{'reuse' if kwargs.get('reuse', True) else 'export'}-{','.join(sheets)}")
        session = Mock()
        with patch.object(batch, "_prepare_row", side_effect=prepare), \
                patch.object(batch, "_finish_pending", side_effect=finish), \
                patch.object(batch, "_refresh", side_effect=refresh):
            batch._repair_file(self.excel, session, self.path)
        # 写入后用改好的副本重新检测（不重开原文件）；文件处理完再用原文件重新导出复查一遍
        self.assertEqual(events, ["prepare", "save", "fresh-copy-reuse-S1", "fresh-original-export-S1"])
        session.written.assert_called_once()
        self.assertEqual(session.written.call_args.args[2:6], ("S1", 5, 24, []))

    def test_new_finding_after_repagination_is_processed_from_new_snapshot(self):
        a, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        prepared = []
        def prepare(_session, _path, items, **_):
            prepared.append(items[0]["cell_addr"])
            return 15, {"height": 24, "shots": {}}, {}
        def finish(*_):
            for iid in list(self.wf.STATE["items"]):
                self.wf.STATE["items"][iid]["status"] = "fixed"
            self.wf.STATE["job"]["pending_adjustment"] = None
        def refresh(*_, **__):
            if len(prepared) == 1:
                _, new_item = self.add_card("A6", row=6)
                batch._data()["rows"].pop(batch._key(self.path, "S1", 6))
                new_item.update(status="pending", page_index=3, bbox=[0, 100, 20, 120])
        with patch.object(batch, "_prepare_row", side_effect=prepare), \
                patch.object(batch, "_finish_pending", side_effect=finish), \
                patch.object(batch, "_refresh", side_effect=refresh):
            batch._repair_file(self.excel, Mock(), self.path)
        self.assertEqual(prepared, ["A5", "A6"])

    def test_same_cell_two_detector_findings_produce_one_card(self):
        _, item = self.add_card()
        other = {**item, "id": "clip", "kind": "clip"}
        groups = batch._group_cells([item, other])
        self.assertEqual(len(groups), 1)
        card = batch._card_for(groups[0])
        self.assertEqual(card["item_ids"], [item["id"], "clip"])

    def test_pause_keeps_remaining_files_for_resume(self):
        self.wf.STATE["job"]["pause_requested"] = True
        with patch.object(batch, "_prepare_row") as prepare:
            batch._repair_file(self.excel, Mock(), self.path)
        prepare.assert_not_called()

    def test_unlocated_text_is_manual_without_creating_copy_or_saving(self):
        card, item = self.add_card()
        batch._data()["rows"].clear()
        item.update(status="pending", page_index=None, bbox=None)
        session = Mock()
        with patch.object(batch, "_prepare_row") as prepare:
            batch._repair_file(self.excel, session, self.path)
        prepare.assert_not_called()
        session.copy_of.assert_not_called()
        self.book.Save.assert_not_called()
        self.assertEqual(card["color"], "yellow")
        self.assertTrue(card["manual"])
        self.assertFalse(card["verified"])
        self.assertIn("不能确认作者", card["message"])
        self.assertIsNone(self.wf.STATE["job"]["pending_adjustment"])

    def test_unlocated_cell_blocks_automatic_row_height_changes_for_its_peers(self):
        a, first = self.add_card("A5")
        b, other = self.add_card("B5")
        first["status"] = other["status"] = "pending"
        other.update(page_index=None, bbox=None)
        batch._data()["rows"].clear()
        with patch.object(batch, "_prepare_row") as prepare:
            batch._repair_file(self.excel, Mock(), self.path)
        prepare.assert_not_called()
        self.assertTrue(a["manual"] and b["manual"])

    def test_missing_cached_image_does_not_make_located_text_low_confidence(self):
        from core import autofix
        _, item = self.add_card()
        self.assertEqual(autofix.review_only_reason([item]), "")

    def test_commit_resume_finishes_refresh_even_when_all_green_cards_are_hidden(self):
        card, _ = self.add_card(color="green")
        self.refresh.side_effect = RuntimeError("导出暂时失败")
        with self.assertRaisesRegex(RuntimeError, "导出暂时失败"):
            self.commit([card])
        self.assertTrue(card["dismissed"])
        self.assertEqual(self.heights[5], 15)
        self.assertEqual(self.wf.STATE["job"]["completed"], [])
        self.refresh.side_effect = None
        batch._commit_review(self.excel)
        self.assertEqual(self.wf.STATE["job"]["completed"], [self.path])
        self.assertEqual(self.refresh.call_count, 2)
        self.book.Save.assert_called_once()

    def test_a_pending_write_log_cannot_be_replaced_by_another_task(self):
        self.wf.STATE["job"]["pending_adjustment"] = {"path": self.path}
        with self.assertRaisesRegex(ValueError, "写入尚未结算"):
            self.wf.queue_job("batch_scan", {"folder": str(self.root)})

    def test_unsettled_files_do_not_stop_later_files_and_resume_keeps_each_journal(self):
        paths = [self.path, str(self.root / "second.xlsx"), str(self.root / "third.xlsx")]
        for path in paths[1:]:
            Path(path).write_bytes(b"workbook")
        job = self.wf.STATE["job"]
        job.update(workbooks=paths, options={"folder": str(self.root)})
        session = Mock()
        session.app.return_value = self.excel
        visited = []
        def repair(excel, session, path):
            visited.append(path)
            if path != paths[2]:
                job["pending_adjustment"] = {"path": path, "cards": [], "rows": [], "action": "fix"}
                raise RuntimeError("保存后连接断开，状态未知")
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "_repair_file", side_effect=repair), \
                patch.object(batch, "finish_execution_batch"):
            self.assertEqual(batch.run(), paths)
        self.assertEqual(visited, paths)
        self.assertEqual(job["completed"], [paths[2]])
        self.assertEqual(job["pending_adjustment"]["path"], paths[0])
        self.assertEqual(set(job["deferred_adjustments"]), {paths[1]})
        self.assertEqual(set(job["file_errors"]), set(paths[:2]))
        saved = storage.load_snapshot("workflow")["job"]
        self.assertEqual(saved["pending_adjustment"]["path"], paths[0])
        self.assertIn(paths[1], saved["deferred_adjustments"])
        # 恢复时第一份依旧失败，第二份可结算；两份都不会阻塞别的文件。
        settled, repaired = [], []
        def finish(excel):
            pending = job["pending_adjustment"]
            settled.append(pending["path"])
            if pending["path"] == paths[0]:
                raise ValueError("文件被占用")
            job["pending_adjustment"] = None
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "_finish_pending", side_effect=finish), \
                patch.object(batch, "_repair_file", side_effect=lambda excel, session, path: repaired.append(path)), \
                patch.object(batch, "finish_execution_batch"):
            batch.run()
        self.assertEqual(settled, paths[:2])
        self.assertEqual(repaired, [paths[1]])
        self.assertEqual(set(job["completed"]), {paths[1], paths[2]})
        self.assertEqual(job["pending_adjustment"]["path"], paths[0])

    def test_scan_retries_errors_then_skips_file_without_marking_it_complete(self):
        second = str(self.root / "next.xlsx")
        Path(second).write_bytes(b"workbook")
        job = self.wf.STATE["job"]
        job.update(workbooks=[self.path, second], options={"folder": str(self.root)})
        session = Mock()
        attempts = []
        def refresh(excel, path, **kwargs):
            attempts.append(path)
            if path == self.path:
                raise RuntimeError("Excel 不可用")
        self.refresh.side_effect = refresh
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "_repair_file"), patch.object(batch, "finish_execution_batch"):
            batch.run()
        self.assertEqual(attempts, [self.path] * 3 + [second])
        self.assertEqual(job["completed"], [second])
        self.assertIn(self.path, job["file_errors"])

    def test_abandon_promotes_next_deferred_write_without_losing_it(self):
        job = self.wf.STATE["job"]
        job["pending_adjustment"] = {"path": self.path, "rows": [], "cards": []}
        remaining = {"path": "second.xlsx", "rows": [], "cards": []}
        job["deferred_adjustments"] = {"second.xlsx": remaining}
        self.wf.STATE["busy"] = False
        batch.abandon_pending()
        self.assertIs(job["pending_adjustment"], remaining)
        with self.assertRaisesRegex(ValueError, "写入尚未结算"):
            self.wf.queue_job("batch_scan", {"folder": str(self.root)})

    def test_deferred_writes_also_prevent_replacing_job(self):
        self.wf.STATE["job"]["deferred_adjustments"] = {self.path: {"path": self.path}}
        with self.assertRaisesRegex(ValueError, "写入尚未结算"):
            self.wf.queue_job("batch_scan", {"folder": str(self.root)})

    def test_timeout_skips_file_without_another_long_retry(self):
        next_path = str(self.root / "next.xlsx")
        Path(next_path).write_bytes(b"workbook")
        job = self.wf.STATE["job"]
        job.update(workbooks=[self.path, next_path], options={"folder": str(self.root)})
        session = Mock()
        def refresh(excel, path, **kwargs):
            if path == self.path:
                raise TimeoutError("Excel 超时")
        self.refresh.side_effect = refresh
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "_repair_file"), patch.object(batch, "finish_execution_batch"):
            batch.run()
        self.assertEqual(self.refresh.call_count, 2)
        self.assertEqual(job["completed"], [next_path])

    def test_timeout_during_trial_escapes_row_retry_loop(self):
        _, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        session = Mock()
        session.watchdog.timed_out = True
        with patch.object(batch, "_prepare_row", side_effect=RuntimeError("Excel 进程被结束")) as prepare:
            with self.assertRaises(TimeoutError):
                batch._repair_file(self.excel, session, self.path)
        prepare.assert_called_once()

    def test_formally_accepted_blue_height_is_no_longer_pending_rollback(self):
        card, _ = self.add_card()
        self.commit([card])
        self.assertFalse(next(iter(batch._data()["rows"].values()))["saved"])

    def test_preparation_uses_one_extra_step_and_falls_back_if_extra_fails(self):
        from test_autofix import AutofixTests
        fixture = AutofixTests("runTest")
        fixture.setUp()
        self.addCleanup(fixture.doCleanups)
        items = list(workflow.STATE["items"].values())
        with patch.object(batch, "extract_sheet_snapshot", fixture.autofix.extract_sheet_snapshot):
            base, chosen, before = batch._prepare_row(fixture.session, str(fixture.wb_path), items)
        self.assertEqual(base, 15)
        self.assertEqual(chosen["height"], 24)  # 21 首次通过，再加 3pt
        self.assertTrue(all(before.values()))
        cid = batch._key(*batch._cell(items[0]))
        self.assertTrue(chosen["shots"][cid])
        original_evaluate = fixture.autofix._evaluate
        def evaluate(*args, **kwargs):
            result = original_evaluate(*args, **kwargs)
            if fixture.session.copy_of(str(fixture.wb_path)).heights[5] == 24:
                return False, result[1], 0
            return result
        with patch.object(batch, "extract_sheet_snapshot", fixture.autofix.extract_sheet_snapshot), \
                patch.object(fixture.autofix, "_evaluate", side_effect=evaluate):
            _, chosen, _ = batch._prepare_row(fixture.session, str(fixture.wb_path), items)
        self.assertEqual(chosen["height"], 21)

    def test_preparation_reuses_checked_pdf_and_keeps_the_written_state(self):
        """原样那张用刚检测导出的 PDF；最后一次导出就是写入的状态时留下给重新检测用"""
        from test_autofix import AutofixTests
        fixture = AutofixTests("runTest")
        fixture.setUp()
        self.addCleanup(fixture.doCleanups)
        items = list(workflow.STATE["items"].values())
        session, path = fixture.session, str(fixture.wb_path)
        book = session.copy_of(path)
        self.assertEqual(book.path.name, fixture.wb_path.name)  # 副本和原文件同名，页眉里的文件名一样
        version = workflow._file_version(fixture.wb_path)
        checked = fixture.root / "checked.pdf"
        book.Sheets("S1").ExportAsFixedFormat(Filename=str(checked))
        session.remember_pdf(path, "S1", version, checked)
        other = fixture.root / "other.pdf"
        other.write_bytes(b"%PDF other sheet")
        session.remember_pdf(path, "S2", version, other)
        timing = {}
        with patch.object(batch, "extract_sheet_snapshot", fixture.autofix.extract_sheet_snapshot):
            base, chosen, before = batch._prepare_row(session, path, items, timing=timing)
        self.assertEqual((base, chosen["height"]), (15, 24))
        self.assertTrue(all(before.values()))
        self.assertEqual(timing["exports"], 3)  # 18、21、24 三次；原样 15 那次没有再导出
        self.assertEqual(chosen["pdf"], session.chosen_pdf())
        self.assertTrue(chosen["pdf"].is_file())
        # 原文件写入保存后：副本做同样修改继续用，这张表用通过那次的 PDF，别的表沿用
        fixture.wb_path.write_bytes(b"workbook saved")
        copy = session.written(path, version, "S1", 5, 24, [], chosen["pdf"])
        self.assertIs(copy, book)
        self.assertEqual(book.heights[5], 24)
        self.assertIs(session.copy_of(path), book)
        self.assertEqual(session.pdf_for(path, "S1"), chosen["pdf"])
        self.assertEqual(session.pdf_for(path, "S2"), other)
        other.write_bytes(b"%PDF overwritten by another export")
        self.assertIsNone(session.pdf_for(path, "S2"))  # 被别的导出覆盖了就不能再用
        # 通过那次不是最后一次（多加的那次没过）：不留 PDF，重新检测时照常导出
        original_evaluate = fixture.autofix._evaluate
        def evaluate(*args, **kwargs):
            result = original_evaluate(*args, **kwargs)
            return (False, result[1], 0) if session.copy_of(path).heights[5] == 27 else result
        with patch.object(batch, "extract_sheet_snapshot", fixture.autofix.extract_sheet_snapshot), \
                patch.object(fixture.autofix, "_evaluate", side_effect=evaluate):
            _, chosen, _ = batch._prepare_row(session, path, items)
        self.assertNotIn("pdf", chosen)

    def test_recheck_on_copy_uses_exported_pdf_and_leaves_copy_open(self):
        from core import autofix
        sheet = Mock(Name="S1")
        copy = Mock()
        copy.Sheets = [sheet]
        original = Mock(Sheets=[sheet])
        self.excel.Workbooks.Open.return_value = original
        session = autofix._ExcelSession()
        exported = []
        def export(_sheet, pdf):
            exported.append(pdf)
            pdf.write_bytes(b"%PDF fresh")
        def scan(**kwargs):
            self.wf._scan_one_workbook(
                self.excel, Path(self.path), {"mode": "sheets", "sheets": ["S1"], "scanned": True}, [],
                job=None, resumed=False, rescan_all=False, delete_backups=False, used_pdf_paths=set(),
                read_only=True, pdfs=session, **kwargs)
        with patch.object(self.wf, "_exportable_sheets", return_value=[sheet]), \
                patch.object(self.wf, "_export_sheet_pdf", side_effect=export), \
                patch.object(self.wf, "_detect_sheet_results", return_value=({}, {})), \
                patch.object(self.wf, "update_file_progress"), patch.object(self.wf, "_sync_problem_counts"):
            chosen = self.root / "chosen.pdf"
            chosen.write_bytes(b"%PDF chosen")
            session.remember_pdf(self.path, "S1", self.wf._file_version(Path(self.path)), chosen)
            timing = {}
            scan(workbook=copy, timing=timing)
            self.assertEqual(exported, [])  # 同一状态已导出过：复制那份，不再导出
            self.excel.Workbooks.Open.assert_not_called()
            copy.Close.assert_not_called()
            target = session.pdf_for(self.path, "S1", self.wf._file_version(Path(self.path)))
            self.assertEqual(target.read_bytes(), b"%PDF chosen")
            self.assertEqual(timing.get("exports", 0), 0)
            scan(reuse_pdfs=False)  # 用原文件复查：一律重新导出
            self.assertEqual(len(exported), 1)
            self.excel.Workbooks.Open.assert_called_once()
            original.Close.assert_called_once()

    def wrap_sheet(self, *, wrap=False, merged=True, neighbor="", align=1):
        cells = {}
        def cell(row, col):
            return cells.setdefault((row, col), Mock(MergeCells=False, Formula=neighbor if col != 16 else "x"))
        area = Mock(WrapText=wrap, ShrinkToFit=False, Row=146, Column=16, HorizontalAlignment=align)
        area.Cells.Count = 3 if merged else 1
        sheet = Mock()
        sheet.Range.return_value.MergeArea = area
        sheet.Cells.side_effect = cell
        return sheet, area

    def test_unwrapped_text_cut_off_sideways_is_wrapped(self):
        """文字不换行、横向被截掉（如 概要!P146）：加高没用，要先开自动换行"""
        item = {"kind": "hidden", "hidden_reason": batch.REASON_TRUNCATED, "cell_addr": "P146"}
        sheet, _ = self.wrap_sheet()
        self.assertEqual(batch._wrap_targets(sheet, [item]), ["P146"])
        # 已经换行、压线问题、没合并且右边是空格（文字本来会延伸显示）的都不动
        self.assertEqual(batch._wrap_targets(self.wrap_sheet(wrap=True)[0], [item]), [])
        self.assertEqual(batch._wrap_targets(sheet, [{**item, "kind": "clip"}]), [])
        self.assertEqual(batch._wrap_targets(self.wrap_sheet(merged=False)[0], [item]), [])
        self.assertEqual(batch._wrap_targets(self.wrap_sheet(merged=False, neighbor="備考")[0], [item]), ["P146"])
        # 右对齐看左边
        right = self.wrap_sheet(merged=False, neighbor="x", align=batch.XL_HALIGN_RIGHT)[0]
        self.assertEqual(batch._wrap_targets(right, [item]), ["P146"])

    def test_preparation_turns_on_wrap_from_first_raise_and_restores_copy(self):
        from test_autofix import AutofixTests, FakeSheet
        fixture = AutofixTests("runTest")
        fixture.setUp()
        self.addCleanup(fixture.doCleanups)
        items = list(workflow.STATE["items"].values())
        area = Mock(WrapText=False)
        states = []
        frames = []
        def record(index, height, shots, wrapped):
            states.append(area.WrapText)
            frames.append((index, list(wrapped)))
        with patch.object(batch, "extract_sheet_snapshot", fixture.autofix.extract_sheet_snapshot), \
                patch.object(batch, "_wrap_targets", return_value=["B5"]), \
                patch.object(FakeSheet, "Range", create=True, new=lambda self, addr: Mock(MergeArea=area)):
            base, chosen, _ = batch._prepare_row(fixture.session, str(fixture.wb_path), items, on_attempt=record)
        self.assertEqual(frames[0], (0, []))  # 修改前就是原样
        self.assertEqual(states[:2], [False, True])
        self.assertTrue(all(wrapped == ["B5"] for _, wrapped in frames[1:]))
        self.assertEqual(chosen["wrap"], ["B5"])
        self.assertIs(area.WrapText, False)  # 副本也还原
        self.assertFalse(any(book.saved for book in fixture.books))

    def test_wrapped_fix_is_saved_restored_and_reported(self):
        card, item = self.add_card(address="A5", color="yellow")
        item["status"] = "pending"
        batch._data()["rows"].clear()
        area = self.book.Sheets.return_value.Range.return_value.MergeArea
        area.WrapText = False
        def prepare(*_, on_attempt, **__):
            on_attempt(0, 24, {card["id"]: "data:image/png;base64,YmVmb3Jl"}, [])
            on_attempt(1, 33, {card["id"]: "data:image/png;base64,bGFzdA=="}, ["A5"])
            return 24, {"height": 36, "shots": {}, "wrap": ["A5"]}, {}
        with patch.object(batch, "_prepare_row", side_effect=prepare):
            batch._repair_file(self.excel, Mock(), self.path)
        self.assertIs(area.WrapText, True)
        self.assertEqual(self.heights[5], 36)
        self.assertTrue(item["wrap_text"])
        self.assertEqual(self.wf._fix_record(item)["reason"], self.wf.WRAP_FIX_REASON)
        self.assertTrue(batch.snapshot()["cards"][0]["row_height"]["wrap"])
        # 续跑时换行和行高都到位：不重复保存
        self.book.Save.reset_mock()
        rid = batch._key(self.path, "S1", 5)
        batch._write_rows(self.excel, {"path": self.path, "version": "old", "rows": [
            {"id": rid, "sheet": "S1", "row": 5, "expected": 24, "target": 36, "wrap": {"A5": True}}]})
        self.book.Save.assert_not_called()
        # 标绿：行高和自动换行一起还原
        batch.mark(card["id"], "green")
        self.commit([card])
        self.assertEqual(self.heights[5], 24)
        self.assertIs(area.WrapText, False)

    def test_blue_on_failed_wrapped_attempt_writes_wrap_too(self):
        card, item = self.attempt_card()
        card["attempt_wrap"] = ["A5"]
        area = self.book.Sheets.return_value.Range.return_value.MergeArea
        area.WrapText = False
        batch.mark(card["id"], "blue")
        self.commit([card])
        self.assertEqual(self.heights[5], 33.0)
        self.assertIs(area.WrapText, True)
        self.assertTrue(item["wrap_text"])

    def test_snapshot_restores_cards_but_never_includes_inline_image_data(self):
        card, _ = self.add_card()
        batch._save_shot(card, "before", "data:image/png;base64,YQ==")
        self.wf._persist_state()
        restored = storage.load_snapshot("workflow")
        self.assertIn(card["id"], restored["batch_review"]["cards"])
        snap = batch.snapshot()
        self.assertTrue(snap["cards"][0]["images"]["before"].startswith("/api/batch/image/"))
        self.assertNotIn("base64", str(snap))

    def context_page(self, item, page_id="page_original"):
        item["page_id"] = page_id
        self.wf.STATE["pages"][page_id] = {"id": page_id, "wb_path": self.path, "sheet_name": "S1",
            "page_index": 1, "page_width": 600, "page_height": 800, "image": "data:image/png;base64,cGFnZQ=="}

    def test_page_context_and_box_remain_paired_after_repagination_and_restart(self):
        card, item = self.add_card()
        self.context_page(item)
        batch._card_for([item])
        original = copy.deepcopy(card["context"])
        self.wf.STATE["pages"].clear()
        self.context_page(item, "page_repaginated")
        item["bbox"] = [200, 300, 300, 350]
        batch._card_for([item])
        self.assertEqual(card["context"], original)
        self.wf._persist_state()
        self.wf.STATE.update(storage.load_snapshot("workflow"))
        snap = batch.snapshot()["cards"][0]["context"]
        self.assertEqual(snap["bbox"], [10, 10, 100, 25])
        self.assertEqual(snap["page_index"], 1)
        self.assertEqual(fix_images.image_path(snap["image_file"]).read_bytes(), b"page")
        self.assertTrue(snap["image"].startswith("/api/batch/image/"))
        self.assertNotIn("base64", str(snap))

    def test_cells_on_same_page_share_one_saved_context_image(self):
        a, first = self.add_card("A5")
        b, second = self.add_card("B5")
        self.context_page(first)
        second["page_id"] = first["page_id"]
        second["bbox"] = [200, 10, 300, 25]
        with patch.object(fix_images, "save", wraps=fix_images.save) as save:
            batch._card_for([first])
            batch._card_for([second])
        self.assertEqual(save.call_count, 1)
        self.assertEqual(a["context"]["image_file"], b["context"]["image_file"])
        self.assertNotEqual(a["context"]["bbox"], b["context"]["bbox"])

    def test_context_is_not_guessed_without_page_or_box(self):
        card, item = self.add_card()
        self.context_page(item)
        item["bbox"] = None
        batch._card_for([item])
        self.assertNotIn("context", card)

    def test_missing_context_image_is_reported_without_using_new_pdf_coordinates(self):
        card, item = self.add_card()
        self.context_page(item)
        batch._card_for([item])
        fix_images.image_path(card["context"]["image_file"]).unlink()
        self.assertEqual(batch.snapshot()["cards"][0]["context"]["image"], "")

    def test_failed_trials_persist_last_unapplied_picture_without_saving_original(self):
        card, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        original = Path(self.path).read_bytes()
        before = "data:image/png;base64,YmVmb3Jl"
        last = "data:image/png;base64,bGFzdA=="
        def prepare(*_, on_attempt, **__):
            on_attempt(0, 24, {card["id"]: before})
            on_attempt(1, 27, {card["id"]: "data:image/png;base64,Zmlyc3Q="})
            on_attempt(8, 48, {card["id"]: last})
            return 24, None, {card["id"]: before}
        with patch.object(batch, "_prepare_row", side_effect=prepare):
            batch._repair_file(self.excel, Mock(), self.path)
        self.book.Save.assert_not_called()
        self.excel.Workbooks.Open.assert_not_called()
        self.assertEqual(Path(self.path).read_bytes(), original)
        self.assertEqual(self.heights[5], 24)
        self.assertEqual(card["color"], "yellow")
        self.assertTrue(card["manual"])
        self.assertEqual(card["after_kind"], "attempt")
        self.assertEqual(card["attempt_height"], 48)
        self.assertIsNone(self.wf.STATE["job"]["pending_adjustment"])
        self.assertNotIn("after", item.get("fix_images", {}))
        saved = storage.load_snapshot("workflow")
        self.wf.STATE.update(saved)  # 模拟从 SQLite 恢复，不依赖临时 Excel 副本。
        restored = batch.snapshot()["cards"][0]
        self.assertEqual(restored["after_kind"], "attempt")
        self.assertEqual(fix_images.image_path(restored["fix_images"]["before"]).read_bytes(), b"before")
        self.assertEqual(fix_images.image_path(restored["fix_images"]["after"]).read_bytes(), b"last")
        self.assertEqual(set(restored["images"]), {"before", "after"})

    def test_export_error_keeps_last_completed_trial_picture(self):
        card, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        def prepare(*_, on_attempt, **__):
            on_attempt(0, 24, {card["id"]: "data:image/png;base64,YmVmb3Jl"})
            on_attempt(1, 27, {card["id"]: "data:image/png;base64,bGFzdA=="})
            raise RuntimeError("下一张 PDF 导出失败")
        with patch.object(batch, "_prepare_row", side_effect=prepare):
            batch._repair_file(self.excel, Mock(), self.path)
        self.assertEqual(fix_images.image_path(card["fix_images"]["after"]).read_bytes(), b"last")
        self.assertEqual(card["after_kind"], "attempt")
        self.assertIn("PDF 导出失败", card["message"])
        self.book.Save.assert_not_called()

    def test_failure_before_first_adjustment_keeps_the_actual_original_preview(self):
        card, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        def prepare(*_, on_attempt, **__):
            on_attempt(0, 24, {card["id"]: "data:image/png;base64,YmVmb3Jl"})
            raise RuntimeError("第一张调整图未能导出")
        with patch.object(batch, "_prepare_row", side_effect=prepare):
            batch._repair_file(self.excel, Mock(), self.path)
        self.assertEqual(fix_images.image_path(card["fix_images"]["after"]).read_bytes(), b"before")
        self.assertEqual(card["after_kind"], "attempt")
        self.assertEqual(card["attempt_height"], 24)
        self.book.Save.assert_not_called()

    def test_failed_preview_restores_copy_height_after_all_eight_attempts(self):
        from test_autofix import AutofixTests
        fixture = AutofixTests("runTest")
        fixture.setUp()
        self.addCleanup(fixture.doCleanups)
        fixture.needed = 999  # 行高上限内始终压线
        items = list(workflow.STATE["items"].values())
        frames = []
        with patch.object(batch, "extract_sheet_snapshot", fixture.autofix.extract_sheet_snapshot):
            base, chosen, before = batch._prepare_row(
                fixture.session, str(fixture.wb_path), items,
                on_attempt=lambda index, height, shots, wrapped: frames.append((index, height, shots)))
        self.assertIsNone(chosen)
        self.assertEqual((base, frames[-1][0], frames[-1][1]), (15, 8, 39))
        self.assertTrue(all(before.values()) and all(frames[-1][2].values()))
        self.assertEqual(fixture.real_heights[5], 15)
        self.assertEqual(fixture.session.copy_of(str(fixture.wb_path)).heights[5], 15)
        self.assertFalse(any(book.saved for book in fixture.books))

    def test_successful_write_clears_unapplied_trial_label(self):
        card, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        def prepare(*_, on_attempt, **__):
            on_attempt(0, 24, {card["id"]: "data:image/png;base64,YmVmb3Jl"})
            on_attempt(1, 27, {card["id"]: "data:image/png;base64,bGFzdA=="})
            return 24, {"height": 27, "shots": {card["id"]: "data:image/png;base64,c2VsZWN0ZWQ="}}, {}
        with patch.object(batch, "_prepare_row", side_effect=prepare):
            batch._repair_file(self.excel, Mock(), self.path)
        self.assertEqual((card["color"], card["after_kind"]), ("blue", "fixed"))
        self.assertEqual(fix_images.image_path(card["fix_images"]["after"]).read_bytes(), b"selected")
        self.assertEqual(self.heights[5], 27)
        self.book.Save.assert_called_once()

    def test_manual_recheck_replaces_unapplied_trial_with_actual_after_image(self):
        card, item = self.add_card(color="yellow")
        card.update(manual=True, after_kind="attempt")
        batch._save_shot(card, "after", "data:image/png;base64,dHJpYWw=")
        item.update(status="fixed")
        with patch.object(batch, "_latest_shot", return_value="data:image/png;base64,YWN0dWFs"):
            batch._recheck_file(self.excel, self.path, [card])
        self.assertEqual((card["color"], card["after_kind"]), ("blue", "fixed"))
        self.assertEqual(fix_images.image_path(card["fix_images"]["after"]).read_bytes(), b"actual")
        self.assertEqual(item["fix_images"]["after"], card["fix_images"]["after"])  # 报告用同一张

    def manual_card(self, address, row=5, sheet="S1", status="manual_review"):
        card, item = self.add_card(address, color="yellow", row=row)
        batch._data()["rows"].pop(batch._key(self.path, "S1", row), None)
        if sheet != "S1":
            del batch._data()["cards"][card["id"]]
            item["sheet_name"] = sheet
            card = batch._card_for([item])
        card.update(color="yellow", finalized=True, manual=True, verified=False, waiting=False)
        item["status"] = status
        return card, item

    def test_recheck_checks_whole_file_and_settles_its_other_cards(self):
        """改一格可能影响别处：整份文件全部重查，文件里其他已提交的黄卡一起出结果"""
        a, item_a = self.manual_card("A5")
        b, item_b = self.manual_card("B5")
        other, item_o = self.manual_card("C5", sheet="S2")
        def refresh(excel, path, sheets=None, **_):
            self.assertIsNone(sheets)  # 整份
            item_a["status"] = item_b["status"] = "fixed"
        self.refresh.side_effect = refresh
        with patch.object(batch, "_latest_shot", return_value=""):
            batch._recheck_file(self.excel, self.path, [a])
        self.assertTrue(a["dismissed"] and b["dismissed"])  # 翻篇，移出列表
        self.assertEqual((other["color"], other["finalized"], other["dismissed"]), ("yellow", True, False))

    def test_failed_recheck_keeps_card_ready_to_recheck_again(self):
        """已提交的人工项复检没过：仍是已提交的人工项，改好后直接再复检，不必先重新提交"""
        card, item = self.manual_card("A5")
        def refresh(*_, **__):
            item["status"] = "pending"  # 重新检测出来的新记录
        self.refresh.side_effect = refresh
        batch._recheck_file(self.excel, self.path, [card])
        self.assertEqual((card["color"], card["finalized"], card["manual"]), ("yellow", True, True))
        self.assertEqual(item["status"], "manual_review")
        self.assertEqual(batch.validate_cards([card["id"]], yellow=True), [card["id"]])

    def test_problems_found_by_recheck_are_shown_ready_for_manual_fixing(self):
        """原来的卡翻篇，这次查出的问题直接是人工待处理项：改好保存后直接复检，不用先提交"""
        card, item = self.manual_card("A5")
        new_item = dict(item, id="new", cell_addr="D5", status="pending", seq=99)
        def refresh(*_, **__):
            item["status"] = "fixed"
            self.wf.STATE["items"]["new"] = new_item
        self.refresh.side_effect = refresh
        with patch.object(batch, "_latest_shot", return_value=""):
            batch._recheck_file(self.excel, self.path, [card])
        new_card = batch._data()["cards"][batch._key(self.path, "S1", "D5")]
        self.assertEqual((new_card["color"], new_card["finalized"], new_card["manual"]), ("yellow", True, True))
        self.assertEqual(new_item["status"], "manual_review")
        self.assertTrue(card["dismissed"])
        self.assertEqual(batch.validate_cards([new_card["id"]], yellow=True), [new_card["id"]])

    def single_recheck(self, card, found=(), span=True):
        """单项复检：这一格在第 2 页（第 4 行上方有分页），只导出这一页"""
        sheet = Mock(Name="S1")
        sheet.PageSetup.Pages.Count = 2 if span else 5
        exported, detected = [], []
        def detect(*_, **kwargs):
            detected.append(kwargs)
            return {}, {f"n{i}": item for i, item in enumerate(found)}
        with patch.object(batch, "_current_area", return_value=(sheet, (5, 1, 5, 1))), \
                patch.object(batch.page_breaks, "read_breaks", return_value={"h": [(4, False)], "v": []}), \
                patch.object(self.wf, "_export_sheet_pdf", side_effect=lambda s, pdf, pages=None: exported.append(pages)), \
                patch.object(self.wf, "_detect_sheet_results", side_effect=detect), \
                patch.object(batch, "_latest_shot", return_value=""):
            batch._recheck_cell(self.excel, self.path, card)
        return exported, detected

    def test_single_recheck_only_exports_the_cells_page_and_leaves_other_cards(self):
        a, item_a = self.manual_card("A5")
        b, item_b = self.manual_card("B5")
        exported, detected = self.single_recheck(a)
        self.assertEqual(exported, [(2, 2)])
        self.assertEqual((detected[0]["page_span"], detected[0]["area"]), ((2, 2), (4, (1 << 20) - 1, 1, (1 << 14) - 1)))
        self.assertTrue(a["dismissed"])
        self.assertEqual(item_a["status"], "fixed")
        self.refresh.assert_not_called()  # 不整份重查
        self.assertEqual((b["color"], b["dismissed"], item_b["status"]), ("yellow", False, "manual_review"))
        self.book.Close.assert_called_once_with(SaveChanges=False)

    def test_single_recheck_only_counts_problems_in_this_cell(self):
        a, item_a = self.manual_card("A5")
        self.single_recheck(a, found=[{"kind": "hidden", "cell_addr": "B7", "text": "别处"}])
        self.assertTrue(a["dismissed"])  # 别处的问题不管
        b, _ = self.manual_card("B5")
        self.single_recheck(b, found=[{"kind": "clip", "cell_addr": "D9", "text": "别的字", "location_confirmed": False}])
        self.assertTrue(b["dismissed"])
        c, item_c = self.manual_card("C5")
        self.single_recheck(c, found=[{"kind": "clip", "cell_addr": "D9", "text": "C", "location_confirmed": False}])
        self.assertFalse(c["dismissed"])  # 认不准是哪一格、文字对得上的切字算这一格的
        self.assertEqual((c["color"], c["finalized"], item_c["status"]), ("yellow", True, "manual_review"))
        self.assertTrue(c["message"].startswith(batch.SINGLE_FAIL_MESSAGE))
        self.assertEqual(batch.validate_cards([c["id"]], yellow=True), [c["id"]])

    def test_single_recheck_exports_the_whole_sheet_when_page_is_unclear(self):
        a, _ = self.manual_card("A5")
        exported, detected = self.single_recheck(a, span=False)  # 页数和分页对不上
        self.assertEqual(exported, [None])
        self.assertIsNone(detected[0]["page_span"])

    def test_single_recheck_of_unchanged_file_keeps_the_card(self):
        a, _ = self.manual_card("A5")
        batch._sheet_checks()[self.path] = {"*": self.wf._file_version(Path(self.path))}
        batch._recheck_cell(self.excel, self.path, a)
        self.excel.Workbooks.Open.assert_not_called()
        self.assertEqual((a["color"], a["dismissed"]), ("yellow", False))

    def test_single_recheck_follows_the_mark_rule_to_the_cells_new_place(self):
        a, _ = self.manual_card("A5")
        batch._marks()[a["id"]] = {"path": self.path, "sheet": "S1", "addr": "A5", "tag": batch._mark_name(a["id"])}
        sheet = Mock(Name="S1")
        area = Mock(Row=8, Column=1, Worksheet=sheet)  # 上面插了 3 行
        area.Rows.Count = area.Columns.Count = 1
        rule = Mock(Formula1=f'=N("{batch._mark_name(a["id"])}")=0')
        rule.AppliesTo.Areas.return_value.Cells.return_value.MergeArea = area
        sheet.Cells.FormatConditions.Count = 1
        sheet.Cells.FormatConditions.Item.return_value = rule
        book = Mock(Worksheets=[sheet])
        book.Sheets.return_value = sheet
        self.assertEqual(batch._current_area(book, self.path, a), (sheet, (8, 1, 8, 1)))

    def test_page_span_covers_the_rows_between_breaks(self):
        sheet = Mock()
        sheet.PageSetup.Pages.Count = 3
        breaks = {"h": [(10, False), (20, True)], "v": []}
        self.assertEqual(batch._page_span(sheet, breaks, (12, 2, 12, 2)), ((2, 2), (10, 19, 1, (1 << 14) - 1)))
        self.assertEqual(batch._page_span(sheet, breaks, (9, 1, 11, 1))[0], (1, 2))  # 合并格跨两页
        self.assertIsNone(batch._page_span(sheet, None, (12, 2, 12, 2)))
        sheet.PageSetup.Pages.Count = 4  # 有空白页被跳过等：页码算不准
        self.assertIsNone(batch._page_span(sheet, breaks, (12, 2, 12, 2)))

    def test_refresh_remembers_which_file_version_each_sheet_was_checked_at(self):
        version = self.wf._file_version(Path(self.path))
        def scan(*_, **kwargs):
            self.wf.STATE["checked_files"][self.path] = version
        with patch.object(self.wf, "_scan_one_workbook", side_effect=scan):
            self.real_refresh(self.excel, self.path, ["S1"])
            self.assertFalse(batch._unchanged_since_check(self.path))  # 只查了一张表不算整份
            # 副本上的检测不算原文件的检测
            self.real_refresh(self.excel, self.path, workbook=Mock())
            self.assertFalse(batch._unchanged_since_check(self.path))
            self.real_refresh(self.excel, self.path)
            self.assertTrue(batch._unchanged_since_check(self.path))
        Path(self.path).write_bytes(b"edited workbook")
        self.assertFalse(batch._unchanged_since_check(self.path))

    def test_recheck_of_unchanged_file_reuses_last_detection_without_exporting(self):
        card, item = self.manual_card("A5")
        item["status"] = "pending"
        batch._sheet_checks()[self.path] = {"*": self.wf._file_version(Path(self.path))}
        batch._recheck_file(self.excel, self.path, [card])
        self.refresh.assert_not_called()
        self.assertEqual((card["color"], card["finalized"]), ("yellow", True))
        # 文件改过就重新导出检测
        Path(self.path).write_bytes(b"edited workbook")
        batch._recheck_file(self.excel, self.path, [card])
        self.refresh.assert_called_once()

    def test_cards_on_a_sheet_no_longer_in_the_file_pass_with_a_clear_message(self):
        """表被删除 / 隐藏 / 改名认不出：整份重查已包含现在所有的表，原来的问题算解决"""
        card, item = self.manual_card("A5")
        def refresh(*_, **__):
            item["status"] = "fixed"
            return ["S1"]
        self.refresh.side_effect = refresh
        batch._recheck_file(self.excel, self.path, [card])
        self.assertTrue(card["dismissed"])

    def test_renamed_sheet_moves_cards_rows_rules_and_highlights(self):
        card, item = self.add_card("A5", color="yellow")
        old_id, old_rid = card["id"], batch._key(self.path, "S1", 5)
        batch._marks()[old_id] = {"path": self.path, "sheet": "S1", "addr": "A5"}
        batch._data()["ignored"].append({"cell": [self.path, "S1", "B5"], "signatures": []})
        batch._sheet_checks()[self.path] = {"S1": {"v": 1}}
        self.wf.STATE["job"].update(kind="batch_recheck", options={"card_ids": [old_id]})
        item["sheet_name"] = "新"
        batch.rename_sheet(self.path, "S1", "新")
        new_id = batch._key(self.path, "新", "A5")
        self.assertIs(batch._data()["cards"][new_id], card)
        self.assertNotIn(old_id, batch._data()["cards"])
        self.assertEqual((card["id"], card["sheet_name"]), (new_id, "新"))
        self.assertIs(batch._card_for([item]), card)  # 重新检测出来的同一格还是这张卡
        self.assertEqual(batch._data()["rows"][batch._key(self.path, "新", 5)]["sheet"], "新")
        self.assertNotIn(old_rid, batch._data()["rows"])
        self.assertEqual(batch._marks()[new_id]["sheet"], "新")
        self.assertEqual(batch._data()["ignored"][-1]["cell"][1], "新")
        self.assertEqual(batch._sheet_checks()[self.path], {"新": {"v": 1}})
        self.assertEqual(self.wf.STATE["job"]["options"]["card_ids"], [new_id])

    def test_recheck_without_marked_colors_checks_the_file_directly(self):
        card, _ = self.manual_card("A5")
        batch._recheck_file(self.excel, self.path, [card])
        self.assertIsNone(self.refresh.call_args.kwargs["workbook"])
        self.excel.Workbooks.Open.assert_not_called()

    def test_unfinished_auto_cards_are_left_alone_by_recheck(self):
        card, _ = self.manual_card("A5")
        waiting, item_w = self.add_card("B5", color="yellow")
        waiting.update(waiting=True, message=batch.WAITING_MESSAGE)
        item_w["status"] = "pending"
        batch._recheck_file(self.excel, self.path, [card])
        self.assertEqual((waiting["finalized"], waiting.get("manual")), (False, False))

    def test_one_huge_log_line_is_cut_short(self):
        self.wf.log_msg("x" * 100000)
        self.assertLess(len(self.wf.STATE["logs"][-1]), self.wf.MAX_LOG_CHARS + 100)

    def recheck_job(self, cards, paths):
        self.wf.STATE.update(busy=True)
        self.wf.STATE["job"] = {"kind": "batch_recheck", "options": {"card_ids": [c["id"] for c in cards]},
                                "completed": [], "pending_adjustment": None, "status": "running",
                                "workbooks": list(paths)}
        return self.wf.STATE["job"]

    def test_recheck_clicked_while_rechecking_joins_the_running_task(self):
        second = str(self.root / "second.xlsx")
        Path(second).write_bytes(b"workbook")
        a, _ = self.manual_card("A5")
        b, item_b = self.manual_card("B5")
        item_b["wb_path"] = second
        del batch._data()["cards"][b["id"]]
        b = batch._card_for([item_b])
        b.update(color="yellow", finalized=True, manual=True)
        job = self.recheck_job([a], [self.path])
        session = Mock()
        session.app.return_value = self.excel
        checked = []
        def recheck(excel, path, cards):
            checked.append((path, [c["id"] for c in cards]))
            if len(checked) == 1:
                self.assertTrue(batch.enqueue_recheck([b["id"]]))    # 另一个文件：排在后面
                self.assertTrue(batch.enqueue_recheck([a["id"]]))    # 正在查的文件又点了：查完再查一遍
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "close_saved_workbooks", return_value={}), \
                patch.object(batch, "_recheck_file", side_effect=recheck), \
                patch.object(batch, "_sync_highlights_safely"), \
                patch.object(batch, "finish_execution_batch"):
            self.assertEqual(batch.run(), [second, self.path])
        self.assertEqual(checked, [(self.path, [a["id"]]), (second, [b["id"]]), (self.path, [a["id"]])])
        self.assertEqual(set(job["completed"]), {self.path, second})
        self.assertTrue(job["closing"])
        self.assertFalse(batch.enqueue_recheck([a["id"]]))  # 已在收尾：另起任务

    def test_card_recheck_runs_single_cell_check_and_file_recheck_runs_whole_file(self):
        a, _ = self.manual_card("A5")
        b, _ = self.manual_card("B5")
        job = self.recheck_job([a], [self.path])
        job["options"] = batch.recheck_options([a["id"]], single=True)
        self.assertEqual(job["options"]["full_paths"], [])
        session = Mock()
        session.app.return_value = self.excel
        calls = []
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "close_saved_workbooks", return_value={}), \
                patch.object(batch, "_recheck_file", side_effect=lambda e, p, cards: calls.append(("file", len(cards)))), \
                patch.object(batch, "_recheck_cell", side_effect=lambda e, p, card: calls.append(("cell", card["id"]))), \
                patch.object(batch, "_sync_highlights_safely"), \
                patch.object(batch, "finish_execution_batch"):
            batch.run()
            self.assertEqual(calls, [("cell", a["id"])])
            calls.clear()
            job = self.recheck_job([b], [self.path])
            job["options"] = batch.recheck_options([b["id"]])
            batch.run()
            self.assertEqual(calls, [("file", 1)])

    def test_single_recheck_clicked_after_file_recheck_only_checks_that_cell(self):
        a, _ = self.manual_card("A5")
        job = self.recheck_job([a], [self.path])
        job["options"] = batch.recheck_options([a["id"]])
        job.update(completed=[self.path], skipped=[], again=[], current=None)
        self.assertTrue(batch.enqueue_recheck([a["id"]], single=True))
        self.assertFalse(batch._full_recheck(job["options"], self.path))
        self.assertTrue(batch.enqueue_recheck([a["id"]]))  # 再点整份
        self.assertTrue(batch._full_recheck(job["options"], self.path))
        self.assertTrue(batch.enqueue_recheck([a["id"]], single=True))  # 整份还在排队：整份已包含这一格
        self.assertTrue(batch._full_recheck(job["options"], self.path))
        self.assertTrue(batch._full_recheck({"card_ids": []}, self.path))  # 旧任务都是整份

    def run_recheck(self, recheck, session):
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "close_saved_workbooks", return_value={}), \
                patch.object(batch, "_recheck_file", side_effect=recheck), \
                patch.object(batch, "_sync_highlights_safely"), \
                patch.object(batch, "_restart_excel") as restart, \
                patch.object(batch, "finish_execution_batch"), \
                patch.object(self.wf, "log_msg") as log:
            batch.run()
        return restart, [c.args[0] for c in log.call_args_list]

    def test_recheck_restarts_excel_when_it_crashes_and_tries_again(self):
        """Excel 挂了（RPC サーバーを利用できません）：换新的 Excel 再查，不让后面的文件跟着全失败"""
        a, _ = self.manual_card("A5")
        job = self.recheck_job([a], [self.path])
        calls = []
        def recheck(excel, path, cards):
            calls.append(path)
            if len(calls) == 1:
                raise RuntimeError("RPC サーバーを利用できません。")
        restart, messages = self.run_recheck(recheck, Mock())
        restart.assert_called_once()
        self.assertEqual(calls, [self.path, self.path])
        self.assertIn(self.path, job["completed"])
        self.assertTrue(any("重启 Excel 后重试" in m and "出错位置" in m for m in messages), messages)

    def test_recheck_gives_up_after_one_retry_and_explains_where(self):
        a, _ = self.manual_card("A5")
        job = self.recheck_job([a], [self.path])
        def recheck(excel, path, cards):
            raise RuntimeError("例外が発生しました。")
        restart, messages = self.run_recheck(recheck, Mock())
        restart.assert_called_once()
        self.assertNotIn(self.path, job["completed"])
        self.assertTrue(any("复检失败" in m and "出错位置：run 第" in m for m in messages), messages)

    def test_recheck_does_not_retry_check_failures(self):
        a, _ = self.manual_card("A5")
        self.recheck_job([a], [self.path])
        restart, _ = self.run_recheck(Mock(side_effect=ValueError("扫描期间文件已修改")), Mock())
        restart.assert_not_called()

    def test_highlight_update_retries_once_with_a_new_excel(self):
        session = Mock()
        with patch.object(batch, "_sync_highlights", side_effect=[AttributeError("Open.Close"), None]) as sync, \
                patch.object(batch, "_restart_excel") as restart:
            batch._sync_highlights_safely(session, self.path)
        self.assertEqual(sync.call_count, 2)
        restart.assert_called_once_with(session)

    def test_recheck_cannot_join_other_tasks(self):
        a, _ = self.manual_card("A5")
        self.wf.STATE.update(busy=True)
        self.wf.STATE["job"].update(kind="batch_commit", options={"card_ids": []})
        self.assertFalse(batch.enqueue_recheck([a["id"]]))
        self.wf.STATE.update(busy=False)
        self.recheck_job([a], [self.path])["status"] = "interrupted"
        self.wf.STATE.update(busy=False)
        self.assertFalse(batch.enqueue_recheck([a["id"]]))

    def test_manual_recheck_does_not_reuse_crop_from_old_coordinates(self):
        """复检通过时按旧坐标裁的图可能是别处：最新 PDF 找不到这格文字就不放图"""
        card, item = self.add_card(color="yellow")
        card.update(manual=True, after_kind="attempt")
        batch._save_shot(card, "after", "data:image/png;base64,dHJpYWw=")
        stale = {"id": "stale", "wb_path": self.path}
        fix_images.save(stale, "after", "data:image/png;base64,c3RhbGU=")
        item.update(status="fixed", fix_images=stale["fix_images"])
        with patch.object(batch, "_latest_shot", return_value=""):
            batch._recheck_file(self.excel, self.path, [card])
        self.assertEqual(card["color"], "blue")
        self.assertNotIn("after", card["fix_images"])
        self.assertNotIn("after", item["fix_images"])

    def test_abandon_unblocks_tool_after_user_changed_height(self):
        card, item = self.add_card(color="yellow")
        self.heights[5] = 36
        self.commit([card])
        self.assertIn("放弃这次写入", card["message"])
        self.wf.STATE.update(busy=False)
        self.wf.STATE["job"]["status"] = "interrupted"
        batch._show_deferred_pending()
        batch.abandon_pending()
        self.assertIsNone(self.wf.STATE["job"]["pending_adjustment"])
        self.assertEqual(self.heights[5], 36)
        self.assertEqual((card["color"], card["manual"], card["finalized"]), ("yellow", True, True))
        self.assertEqual(item["status"], "manual_review")
        self.assertFalse(next(iter(batch._data()["rows"].values()))["saved"])
        self.assertTrue(self.wf.queue_job("batch_scan", {"folder": str(self.root)}))

    def test_unsettled_file_does_not_block_committing_other_files(self):
        """一份文件行高被人工改过、还原核对不通过：只跳过这份，别的文件的绿卡照常提交"""
        stuck, _ = self.add_card(color="green")
        self.heights[5] = 36
        other = str(self.root / "other.xlsx")
        Path(other).write_bytes(b"workbook")
        green, item = self.add_card("A6", color="green", row=6)
        green["wb_path"] = item["wb_path"] = other
        batch._data()["rows"].pop(batch._key(self.path, "S1", 6))
        self.commit([stuck, green])
        job = self.wf.STATE["job"]
        self.assertEqual(job["completed"], [other])
        self.assertTrue(green["dismissed"] and green["finalized"])
        self.assertEqual((stuck["color"], stuck["finalized"]), ("green", False))
        self.assertEqual(self.heights[5], 36)
        self.assertIsNone(job["pending_adjustment"])
        self.assertEqual(set(job["deferred_adjustments"]), {self.path})

    def test_review_tasks_start_while_another_file_has_unsettled_write(self):
        pending = {"path": self.path, "action": "restore", "rows": [], "cards": []}
        job = self.wf.STATE["job"]
        job.update(kind="batch_commit", status="interrupted", pending_adjustment=pending)
        self.wf.STATE["busy"] = False
        with self.assertRaisesRegex(ValueError, "写入尚未结算"):
            self.wf.queue_job("batch_scan", {"folder": str(self.root)})
        self.assertTrue(self.wf.queue_job("batch_commit", {"card_ids": []}))
        new = self.wf.STATE["job"]
        self.assertIsNot(new, job)
        self.assertIsNone(new["pending_adjustment"])
        self.assertEqual(new["deferred_adjustments"], {self.path: pending})

    def test_resume_with_failing_write_still_commits_other_files(self):
        pending = {"path": self.path, "action": "restore", "rows": [], "cards": []}
        job = self.wf.STATE["job"]
        job.update(kind="batch_commit", options={"card_ids": []}, pending_adjustment=pending)
        session = Mock()
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "_finish_pending", side_effect=ValueError("行高已被其他操作修改")), \
                patch.object(batch, "_commit_review") as commit, patch.object(batch, "finish_execution_batch"):
            batch.run()
        commit.assert_called_once()
        self.assertIs(job["pending_adjustment"], pending)  # 结束后照旧显示「继续 / 放弃这次写入」
        self.assertEqual(job.get("deferred_adjustments"), {})

    def test_abandon_requires_a_pending_write(self):
        self.wf.STATE["busy"] = False
        with self.assertRaisesRegex(ValueError, "没有需要放弃"):
            batch.abandon_pending()

    def test_starting_another_folder_keeps_review_data(self):
        card, _ = self.add_card(color="green")
        batch._data()["ignored"].append({"cell": [self.path, "S1", "Z1"], "signatures": [["clip", "x"]]})
        other = self.root / "other"
        other.mkdir()
        self.wf.STATE["job"] = {"kind": "batch_scan", "options": {"folder": str(other)}, "completed": [],
                                "pending_adjustment": None, "status": "running"}
        session = Mock()
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "init_execution_batch"), patch.object(batch, "finish_execution_batch"), \
                patch.object(self.wf, "list_workbooks", return_value=(other, [])):
            batch.run()
        self.assertEqual(self.wf.STATE["folder"], str(other))
        self.assertIn(card["id"], batch._data()["cards"])
        self.assertEqual(len(batch._data()["ignored"]), 1)
        self.assertEqual(self.wf.STATE["items"], {})

    def test_excel_error_restarts_excel_and_retries_row_before_manual(self):
        card, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        calls = []
        def prepare(*_, **__):
            calls.append(1)
            if len(calls) == 1:
                raise RuntimeError("RPC 服务器不可用")
            return 15, {"height": 24, "shots": {}}, {}
        session = Mock()
        session.app.return_value = self.excel
        with patch.object(batch, "_prepare_row", side_effect=prepare):
            batch._repair_file(self.excel, session, self.path)
        self.assertEqual(len(calls), 2)
        session.close.assert_called_once()
        self.assertEqual(card["color"], "blue")
        self.assertFalse(card.get("manual"))

    def test_repeated_excel_errors_end_as_manual(self):
        card, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        session = Mock()
        with patch.object(batch, "_prepare_row", side_effect=RuntimeError("导出失败")) as prepare:
            batch._repair_file(self.excel, session, self.path)
        self.assertEqual(prepare.call_count, 3)
        self.assertEqual(session.close.call_count, 2)
        self.assertTrue(card["manual"])
        self.assertEqual(card["color"], "yellow")

    def test_detection_reason_is_not_retried(self):
        card, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        session = Mock()
        with patch.object(batch, "_prepare_row", side_effect=ValueError("单元格已隐藏")) as prepare:
            batch._repair_file(self.excel, session, self.path)
        self.assertEqual(prepare.call_count, 1)
        session.close.assert_not_called()
        self.assertTrue(card["manual"])

    def test_original_already_passing_writes_nothing_and_is_not_a_fix(self):
        card, item = self.add_card(color="yellow")
        item["status"] = "pending"
        batch._data()["rows"].clear()
        with patch.object(batch, "_prepare_row", return_value=(24.0, {"height": 24.0, "shots": {}}, {})):
            batch._repair_file(self.excel, Mock(), self.path)
        self.excel.Workbooks.Open.assert_not_called()
        self.assertEqual((card["color"], card["after_kind"]), ("blue", "unchanged"))
        self.assertEqual(item["status"], "accepted")
        self.assertEqual(batch._data()["rows"], {})
        self.assertIsNone(self.wf._problem_counts(self.path)["fixes"][item["id"]])

    def test_refresh_downgrades_blue_card_whose_problem_came_back(self):
        card, item = self.add_card()
        item["status"] = "pending"
        def scan(*_args, **_kwargs):
            pass
        with patch.object(self.wf, "_scan_one_workbook", side_effect=scan):
            self.real_refresh(self.excel, self.path, ["S1"])
        self.assertEqual(card["color"], "yellow")
        self.assertFalse(card["verified"])

    def test_cell_passing_without_its_row_written_replaces_stale_trial_picture(self):
        # 试算第一张（原样）同时存成了「修改后」；之后别的行调整让这格不再检出，不能沿用那张一样的图
        card, item = self.add_card()
        batch._data()["rows"].clear()
        batch._save_shot(card, "before", "data:image/png;base64,YmVmb3Jl")
        batch._save_shot(card, "after", "data:image/png;base64,YmVmb3Jl")
        card.update(color="yellow", verified=False, after_kind="attempt")
        with patch.object(self.wf, "_scan_one_workbook"), \
                patch.object(batch, "_latest_shot", return_value="data:image/png;base64,bGF0ZXN0"):
            self.real_refresh(self.excel, self.path, ["S1"])
        self.assertEqual((card["color"], card["after_kind"]), ("blue", "rechecked"))
        self.assertIn("本行行高未修改", card["message"])
        self.assertEqual(fix_images.image_path(card["fix_images"]["after"]).read_bytes(), b"latest")

    def test_cell_passing_without_picture_shows_none_instead_of_old_one(self):
        card, item = self.add_card()
        batch._data()["rows"].clear()
        batch._save_shot(card, "after", "data:image/png;base64,YmVmb3Jl")
        with patch.object(self.wf, "_scan_one_workbook"), patch.object(batch, "_latest_shot", return_value=""):
            self.real_refresh(self.excel, self.path, ["S1"])
        self.assertNotIn("after", card["fix_images"])

    def test_snapshot_explains_problem_and_row_height_change(self):
        card, item = self.add_card()
        self.assertEqual(card["problems"][0]["label"], "隐藏截断")
        item.update(text="上記以外、\n空", visible_text="上記以外、", hidden_reason="后面的文字没显示")
        card.pop("problems")
        result = batch.snapshot()["cards"][0]
        self.assertEqual(result["problems"], [{"kind": "hidden", "label": "隐藏截断",
                                               "reason": "后面的文字没显示", "missing": "空"}])
        self.assertEqual(result["row_height"], {"base": 15.0, "applied": 24.0, "saved": True, "wrap": False})
        clip = {**item, "id": "clip", "kind": "clip", "cell_addr": "B5"}
        self.assertEqual(batch._problem(clip)["label"], "压线")
        self.assertEqual(batch._problem(clip)["missing"], "")

    def test_snapshot_image_urls_change_with_content(self):
        card, _ = self.add_card()
        batch._save_shot(card, "before", "data:image/png;base64,YQ==")
        first = batch.snapshot()["cards"][0]["images"]["before"]
        batch._save_shot(card, "before", "data:image/png;base64,YWJj")
        second = batch.snapshot()["cards"][0]["images"]["before"]
        self.assertNotEqual(first, second)
        self.assertEqual(first.split("?")[0], second.split("?")[0])

    def test_folder_switch_keeps_pictures_still_used_by_cards(self):
        card, _ = self.add_card()
        batch._save_shot(card, "before", "data:image/png;base64,YQ==")
        orphan = fix_images.IMAGES_DIR / "orphan_after.png"
        orphan.write_bytes(b"x")
        fix_images.clear()
        self.assertIsNotNone(fix_images.image_path(card["fix_images"]["before"]))
        self.assertFalse(orphan.exists())

    def test_review_tasks_do_not_write_another_folders_report(self):
        reporter.EXECUTION_STATE.update(folder=str(self.root / "elsewhere"), files={})
        self.wf.STATE.update(busy=True)
        self.wf.STATE["job"].update(kind="batch_recheck", options={"card_ids": []}, status="running")
        def run():
            reporter.update_file_progress("book.xlsx", status="done")
            reporter.finish_execution_batch("x")
            return []
        with patch.object(batch, "run", side_effect=run):
            self.wf.run_pending_workflow()
        self.assertEqual(reporter.EXECUTION_STATE["files"], {})
        reporter.update_file_progress("book.xlsx", status="done")  # 任务结束后恢复写入
        self.assertIn("book.xlsx", reporter.EXECUTION_STATE["files"])


if __name__ == "__main__":
    unittest.main()
