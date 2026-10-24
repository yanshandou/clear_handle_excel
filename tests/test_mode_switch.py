import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, fix_images, reporter, storage, workflow


class ModeSwitchTests(unittest.TestCase):
    """全自动先扫完全部文件再修；扫完后可以「转人工」，人工后可以「转回自动」，来回切换"""

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
        self.paths = [str(self.root / "a.xlsx"), str(self.root / "b.xlsx")]
        for path in self.paths:
            Path(path).write_bytes(b"workbook")
        self.wf.STATE.update(folder=str(self.root), items={}, pages={}, busy=True,
                             batch_review={"cards": {}, "rows": {}, "ignored": []})
        self.job = self.wf.STATE["job"] = {
            "kind": "batch_scan", "options": {"folder": str(self.root), "highlight": dict(batch.DEFAULT_HIGHLIGHT)},
            "completed": [], "pending_adjustment": None, "status": "running", "workbooks": list(self.paths)}
        self.session = Mock()
        patch.object(batch.autofix, "_ExcelSession", return_value=self.session).start()
        patch.object(batch, "finish_execution_batch").start()
        patch.object(batch, "init_execution_batch").start()
        patch.object(self.wf, "_sync_problem_counts").start()
        self.sync = patch.object(batch, "_sync_highlights").start()
        self.addCleanup(patch.stopall)

    def add(self, path, address, row, *, color="yellow", waiting=True, status="pending", height=18.0, **card):
        iid = f"{Path(path).stem}{address}"
        item = {"id": iid, "seq": len(self.wf.STATE["items"]) + 1, "kind": "hidden", "text": address,
                "wb_path": path, "wb_name": Path(path).name, "sheet_name": "S1", "cell_addr": address,
                "excel_row": row, "row_height": height, "status": status, "image": ""}
        self.wf.STATE["items"][iid] = item
        result = batch._card_for([item])
        result.update(color=color, waiting=waiting, **card)
        return result, item

    def test_all_files_are_scanned_before_any_repair(self):
        order = []
        with patch.object(batch, "_refresh", side_effect=lambda excel, path, *a, **k: order.append(("scan", path))), \
                patch.object(batch, "_repair_file", side_effect=lambda excel, session, path: order.append(("fix", path))):
            batch.run()
        self.assertEqual(order, [("scan", self.paths[0]), ("scan", self.paths[1]),
                                 ("fix", self.paths[0]), ("fix", self.paths[1])])
        self.assertEqual(self.job["phase"], "repair")
        self.assertEqual(self.job["completed"], self.paths)

    def test_cannot_switch_to_manual_while_still_scanning(self):
        with self.assertRaisesRegex(ValueError, "扫描完成后"):
            batch.request_manual()
        self.job["phase"] = "repair"
        self.assertFalse(batch.request_manual())  # 正在跑：做完当前这一行再转
        self.assertTrue(self.job["to_manual"])
        self.job.pop("to_manual")
        self.wf.STATE["busy"] = False
        self.job["status"] = "interrupted"
        self.assertTrue(batch.request_manual())  # 已暂停：由调用方续跑任务去转

    def test_switch_to_manual_after_current_row(self):
        fixed, fixed_item = self.add(self.paths[0], "A5", 5, color="blue", waiting=False, status="fixed",
                                     verified=True)
        failed, failed_item = self.add(self.paths[0], "B6", 6, waiting=False, manual=True,
                                       message="自动试修未通过，原文件行高未修改")
        later, later_item = self.add(self.paths[0], "C7", 7)
        other, other_item = self.add(self.paths[1], "A3", 3, height=30.0)
        repaired = []

        def repair(excel, session, path):
            repaired.append(path)
            self.job["to_manual"] = True  # 修这份文件时点了「转人工」：做完当前行返回

        self.job.update(phase="repair", scanned=list(self.paths))
        self.job["scan_versions"] = {p: self.wf._file_version(Path(p)) for p in self.paths}
        with patch.object(batch, "_repair_file", side_effect=repair), patch.object(batch, "_refresh"):
            self.assertEqual(batch.run(), self.paths)
        self.assertEqual(repaired, [self.paths[0]])
        # 已修好的蓝卡原样等审核
        self.assertEqual((fixed["color"], fixed.get("finalized"), fixed_item["status"]), ("blue", False, "fixed"))
        # 自动试过没通过的：交给人工，以后也不交回自动
        self.assertEqual((failed["finalized"], failed.get("handed_over")), (True, False))
        self.assertEqual(failed["message"], "自动试修未通过，原文件行高未修改")
        # 还没轮到的（包括没修到的第二份文件）全部转黄，可以直接人工处理
        for card, item in ((later, later_item), (other, other_item)):
            self.assertEqual((card["color"], card["manual"], card["finalized"], card["handed_over"]),
                             ("yellow", True, True, True))
            self.assertEqual(item["status"], "manual_review")
        self.assertEqual(self.job["completed"], self.paths)
        self.assertNotIn("to_manual", self.job)
        handover = batch._data()["handover"][self.paths[1]]
        self.assertEqual(handover["rows"], {batch._row_key("S1", 3): 30.0})
        self.assertEqual(handover["version"], self.wf._file_version(Path(self.paths[1])))
        self.assertEqual(batch._data()["highlight_files"][self.paths[1]], batch.DEFAULT_HIGHLIGHT)
        self.assertEqual(self.sync.call_count, 2)

    def handed_over(self):
        self.job.update(phase="repair", scanned=list(self.paths), status="completed")
        self.wf.STATE["busy"] = False
        card, item = self.add(self.paths[0], "A5", 5)
        failed, _ = self.add(self.paths[0], "B6", 6, waiting=False, manual=True, message="自动试修未通过")
        batch._hand_over(self.paths[0], switch=True)
        batch._remember_handover(self.paths[0])
        return card, item, failed

    def test_back_to_auto_on_untouched_file_reuses_detection(self):
        card, item, failed = self.handed_over()
        self.assertTrue(batch._mode()["to_auto"])
        with patch.object(batch, "close_saved_workbooks", return_value={}), \
                patch.object(batch, "_recheck_file") as recheck:
            batch._back_to_auto(Mock(), self.session, self.paths[0])
        recheck.assert_not_called()
        self.assertEqual((card["manual"], card["finalized"], card["waiting"], card["handed_over"]),
                         (False, False, True, False))
        self.assertEqual(item["status"], "pending")
        self.assertEqual((failed["manual"], failed["finalized"]), (True, True))
        # 标记颜色全部还原后再自动处理
        self.assertTrue(self.sync.call_args.kwargs["clear"])
        self.assertFalse(batch._mode()["to_auto"])

    def test_back_to_auto_rescans_changed_file_and_leaves_touched_rows_manual(self):
        card, item, _ = self.handed_over()
        other, other_item = self.add(self.paths[0], "C7", 7)
        batch._hand_over(self.paths[0], switch=True)
        batch._remember_handover(self.paths[0])
        Path(self.paths[0]).write_bytes(b"edited by hand")

        def recheck(excel, path, cards):
            item["row_height"] = 30.0  # 人把第 5 行加高了，这一格还是有问题；第 7 行没动

        with patch.object(batch, "close_saved_workbooks", return_value={}), \
                patch.object(batch, "_recheck_file", side_effect=recheck) as rechecked:
            batch._back_to_auto(Mock(), self.session, self.paths[0])
        rechecked.assert_called_once()
        self.assertEqual((card["manual"], card["finalized"], card["message"]), (True, True, batch.TOUCHED_MESSAGE))
        self.assertEqual(item["status"], "manual_review")
        self.assertEqual((other["manual"], other["waiting"]), (False, True))
        self.assertEqual(other_item["status"], "pending")

    def test_open_unsaved_workbook_is_not_taken_back(self):
        self.handed_over()
        with patch.object(batch, "close_saved_workbooks", return_value={Path(self.paths[0]): "还没保存"}):
            with self.assertRaisesRegex(ValueError, "请保存后继续"):
                batch._back_to_auto(Mock(), self.session, self.paths[0])

    def test_back_to_auto_job_only_visits_handed_over_files_then_repairs(self):
        self.handed_over()
        self.wf.STATE["busy"] = True
        self.wf.STATE["job"] = self.job = {
            "kind": "batch_scan", "options": {"folder": str(self.root), "back_to_auto": True},
            "completed": [], "pending_adjustment": None, "status": "running"}
        order = []
        with patch.object(batch, "_back_to_auto", side_effect=lambda excel, session, path: order.append(("back", path))), \
                patch.object(batch, "_refresh") as refresh, \
                patch.object(batch, "_repair_file", side_effect=lambda excel, session, path: order.append(("fix", path))):
            self.assertEqual(batch.run(), [self.paths[0]])
        refresh.assert_not_called()  # 文件在交回自动之后没变：不用再检测一遍
        self.assertEqual(order, [("back", self.paths[0]), ("fix", self.paths[0])])

    def test_recheck_marks_newly_found_cells_as_not_yet_tried_by_auto(self):
        old, _ = self.add(self.paths[0], "A5", 5, waiting=False, manual=True, finalized=True)
        new_item = {"id": "new", "seq": 9, "kind": "hidden", "text": "B9", "wb_path": self.paths[0],
                    "wb_name": "a.xlsx", "sheet_name": "S1", "cell_addr": "B9", "excel_row": 9, "status": "pending"}

        def refresh(excel, path, **kwargs):
            self.wf.STATE["items"]["new"] = new_item
            with self.wf.STATE_LOCK:
                batch._card_for([new_item])

        with patch.object(batch, "_refresh", side_effect=refresh), patch.object(batch, "_open_without_marks", return_value=None):
            batch._recheck_file(Mock(), self.paths[0], [old])
        new = batch._data()["cards"][batch._key(self.paths[0], "S1", "B9")]
        self.assertTrue(new["handed_over"])
        self.assertFalse(old.get("handed_over"))

    def test_resuming_in_the_middle_of_a_file_detects_it_again_first(self):
        self.job.update(phase="repair", scanned=list(self.paths), repairing=self.paths[0])
        self.job["scan_versions"] = {p: self.wf._file_version(Path(p)) for p in self.paths}
        refreshed = []
        with patch.object(batch, "_refresh", side_effect=lambda excel, path, *a, **k: refreshed.append(path)), \
                patch.object(batch, "_repair_file"):
            batch.run()
        self.assertEqual(refreshed, [self.paths[0]])

    def test_renamed_sheet_keeps_handover_row_heights(self):
        batch._data()["handover"] = {self.paths[0]: {"version": None, "rows": {batch._row_key("S1", 5): 18.0,
                                                                               batch._row_key("S10", 5): 20.0}}}
        batch.rename_sheet(self.paths[0], "S1", "新")
        self.assertEqual(batch._data()["handover"][self.paths[0]]["rows"],
                         {batch._row_key("新", 5): 18.0, batch._row_key("S10", 5): 20.0})



if __name__ == "__main__":
    unittest.main()
