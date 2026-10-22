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
        self.refresh = patch.object(batch, "_refresh").start()
        self.addCleanup(patch.stopall)

    def add_card(self, address="A5", color="blue", row=5):
        iid = f"i{address}"
        item = {"id": iid, "seq": len(self.wf.STATE["items"]) + 1, "kind": "hidden", "text": address,
                "wb_path": self.path, "wb_name": "book.xlsx", "sheet_name": "S1", "cell_addr": address,
                "excel_row": row, "row_height": self.heights[row], "status": "fixed", "image": ""}
        item.update(page_index=0, bbox=[10, 10, 100, 25])
        self.wf.STATE["items"][iid] = item
        card = batch._card_for([item])
        card.update(color=color, verified=True)
        rid = batch._key(self.path, "S1", row)
        batch._data()["rows"][rid] = {"id": rid, "path": self.path, "sheet": "S1", "row": row,
                                       "base": 15.0, "applied": self.heights[row], "saved": True}
        return card, item

    def commit(self, cards):
        self.wf.STATE["job"].update(kind="batch_commit", options={"card_ids": [c["id"] for c in cards]})
        batch._commit_review(self.excel)

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
        with self.assertRaisesRegex(ValueError, "行高已被"):
            self.commit([card])
        self.assertEqual(self.heights[5], 36)
        self.assertIsNotNone(self.wf.STATE["job"]["pending_adjustment"])
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

    def test_all_heights_validated_before_any_change(self):
        a, _ = self.add_card("A5", "yellow")
        b, _ = self.add_card("A6", "yellow", row=6)
        self.heights[6] = 90
        with self.assertRaises(ValueError):
            self.commit([a, b])
        self.assertEqual(self.heights, {5: 24.0, 6: 90})
        self.book.Save.assert_not_called()

    def test_unsubmitted_yellow_cannot_be_rechecked(self):
        card, _ = self.add_card(color="yellow")
        with self.assertRaisesRegex(ValueError, "先提交"):
            batch.validate_cards([card["id"]], yellow=True)

    def test_no_automatic_blue_for_unverified_manual_item(self):
        card, _ = self.add_card(color="yellow")
        card["verified"] = False
        with self.assertRaisesRegex(ValueError, "尚未通过"):
            batch.mark(card["id"], "blue")

    def test_recheck_passes_one_cell_and_keeps_ignored_cell_out(self):
        a, item_a = self.add_card("A5", "yellow")
        b, item_b = self.add_card("B5", "green")
        # 两处都已还原，不再有自动行高冲突。
        next(iter(batch._data()["rows"].values()))["saved"] = False
        self.commit([a, b])
        def refresh(*_):
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
        def prepare(*_):
            events.append("prepare")
            return 15, {"height": 24, "shots": {cid: ""}}, {cid: ""}
        def finish(*_):
            events.append("save")
            item["status"] = "fixed"
            self.wf.STATE["job"]["pending_adjustment"] = None
        def refresh(*_):
            events.append("fresh-sheet")
        with patch.object(batch, "_prepare_row", side_effect=prepare), \
                patch.object(batch, "_finish_pending", side_effect=finish), \
                patch.object(batch, "_refresh", side_effect=refresh):
            batch._repair_file(self.excel, Mock(), self.path)
        self.assertEqual(events, ["prepare", "save", "fresh-sheet"])

    def test_new_finding_after_repagination_is_processed_from_new_snapshot(self):
        a, item = self.add_card()
        item["status"] = "pending"
        batch._data()["rows"].clear()
        prepared = []
        def prepare(_session, _path, items):
            prepared.append(items[0]["cell_addr"])
            return 15, {"height": 24, "shots": {}}, {}
        def finish(*_):
            for iid in list(self.wf.STATE["items"]):
                self.wf.STATE["items"][iid]["status"] = "fixed"
            self.wf.STATE["job"]["pending_adjustment"] = None
        def refresh(*_):
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

    def test_snapshot_restores_cards_but_never_includes_inline_image_data(self):
        card, _ = self.add_card()
        batch._save_shot(card, "before", "data:image/png;base64,YQ==")
        self.wf._persist_state()
        restored = storage.load_snapshot("workflow")
        self.assertIn(card["id"], restored["batch_review"]["cards"])
        snap = batch.snapshot()
        self.assertTrue(snap["cards"][0]["images"]["before"].startswith("/api/batch/image/"))
        self.assertNotIn("base64", str(snap))


if __name__ == "__main__":
    unittest.main()
