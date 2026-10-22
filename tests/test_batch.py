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
        def prepare(*_, **__):
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
        def prepare(_session, _path, items, **_):
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
        def prepare(*_, on_attempt):
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
        def prepare(*_, on_attempt):
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
        def prepare(*_, on_attempt):
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
                on_attempt=lambda index, height, shots: frames.append((index, height, shots)))
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
        def prepare(*_, on_attempt):
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
        actual = {"id": "actual", "wb_path": self.path}
        fix_images.save(actual, "after", "data:image/png;base64,YWN0dWFs")
        item.update(status="fixed", fix_images=actual["fix_images"])
        batch._recheck_file(self.excel, self.path, [card])
        self.assertEqual((card["color"], card["after_kind"]), ("blue", "fixed"))
        self.assertEqual(fix_images.image_path(card["fix_images"]["after"]).read_bytes(), b"actual")


if __name__ == "__main__":
    unittest.main()
