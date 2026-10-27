import http.client
import importlib
import json
import tempfile
import threading
import time
import unittest
from http.server import ThreadingHTTPServer
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


    def scanned_only(self):
        """只扫描后的状态：A5、B5 同一行，C7 自动试过没通过，D9 是 ###，E11 你标了黄还没提交"""
        self.job.update(phase="repair", scanned=list(self.paths), status="completed")
        self.wf.STATE["busy"] = False
        cards = {}
        for addr, row in (("A5", 5), ("B5", 5), ("C7", 7), ("E11", 11)):
            cards[addr], _ = self.add(self.paths[0], addr, row)
        cards["D9"], item = self.add(self.paths[0], "D9", 9)
        item["hidden_reason"] = batch.autofix.REASON_HASH
        for card in cards.values():
            self.wf.STATE["items"][card["item_ids"][0]].update(page_index=0, bbox=[0, 0, 1, 1])
        batch._hand_over(self.paths[0])
        batch._remember_handover(self.paths[0])
        cards["C7"].update(handed_over=False, message="自动试修未通过")
        cards["E11"].update(finalized=False, marked=True)
        batch._data()["auto_versions"] = {self.paths[0]: self.wf._file_version(Path(self.paths[0]))}
        return cards

    def test_auto_adjust_plan_takes_whole_row_and_explains_skips(self):
        cards = self.scanned_only()
        plan = batch.auto_adjust_plan([cards[a]["id"] for a in ("A5", "C7", "D9", "E11")])
        self.assertEqual(plan["run"], [cards["A5"]["id"]])
        self.assertEqual(plan["extra"], [cards["B5"]["id"]])  # 同一行共用行高，一起调整
        reasons = {s["id"]: s["reason"] for s in plan["skip"]}
        self.assertIn("文件之后没改过", reasons[cards["C7"]["id"]])
        self.assertIn("###", reasons[cards["D9"]["id"]])
        self.assertIn("先提交", reasons[cards["E11"]["id"]])

    def test_auto_failed_cell_is_retried_only_after_file_changes(self):
        cards = self.scanned_only()
        self.assertEqual(batch.auto_adjust_plan([cards["C7"]["id"]])["run"], [])
        # 程序自己保存（标色、写行高）不算你改过
        before = self.wf._file_version(Path(self.paths[0]))
        Path(self.paths[0]).write_bytes(b"saved by program")
        batch._followed_own_save(self.paths[0], before)
        self.assertEqual(batch.auto_adjust_plan([cards["C7"]["id"]])["run"], [])
        Path(self.paths[0]).write_bytes(b"edited by hand, longer")
        self.assertEqual(batch.auto_adjust_plan([cards["C7"]["id"]])["run"], [cards["C7"]["id"]])

    def test_snapshot_hides_auto_button_for_cells_auto_already_failed(self):
        cards = self.scanned_only()
        shown = {c["id"]: c for c in batch.snapshot()["cards"]}
        self.assertEqual(shown[cards["A5"]["id"]]["auto_skip"], "")
        self.assertIn("文件之后没改过", shown[cards["C7"]["id"]]["auto_skip"])
        self.assertIn("###", shown[cards["D9"]["id"]]["auto_skip"])
        self.assertNotIn("auto_skip", shown[cards["E11"]["id"]])  # 没提交的卡本来就没有自动调整按钮
        Path(self.paths[0]).write_bytes(b"edited by you, a longer file")
        shown = {c["id"]: c for c in batch.snapshot()["cards"]}
        self.assertEqual(shown[cards["C7"]["id"]]["auto_skip"], "")  # 你改过文件：按钮回来

    def test_auto_adjust_refuses_when_nothing_can_run(self):
        cards = self.scanned_only()
        with self.assertRaisesRegex(ValueError, "都不能自动调整"):
            batch.start_auto_adjust([cards["D9"]["id"]])
        self.wf.STATE["busy"] = True
        with self.assertRaisesRegex(ValueError, "后台任务执行中"):
            batch.auto_adjust_plan([cards["A5"]["id"]])

    def test_picked_back_to_auto_only_takes_chosen_cells_and_keeps_other_marks(self):
        cards = self.scanned_only()
        options = batch.start_auto_adjust([cards["A5"]["id"]])
        self.assertEqual(options["card_ids"], [cards["A5"]["id"], cards["B5"]["id"]])
        self.wf.STATE["job"] = {"kind": "batch_scan", "options": options, "completed": [], "status": "running"}
        with patch.object(batch, "close_saved_workbooks", return_value={}), patch.object(batch, "_recheck_file"):
            batch._back_to_auto(Mock(), self.session, self.paths[0], picked=True)
        for addr in ("A5", "B5"):
            self.assertEqual((cards[addr]["waiting"], cards[addr]["manual"], cards[addr]["message"]),
                             (True, False, batch.PICKED_MESSAGE))
            self.assertNotIn("auto_pick", cards[addr])
        for addr in ("C7", "D9"):
            self.assertEqual((cards[addr]["manual"], cards[addr]["finalized"]), (True, True))
        self.assertTrue(cards["D9"]["handed_over"])  # 没选的照旧可以以后再转回自动
        self.assertFalse(self.sync.call_args.kwargs["clear"])  # 只去掉交给自动的格子的标色
        self.assertEqual(self.wf.STATE["job"]["picked_back"], [cards["A5"]["id"], cards["B5"]["id"]])

    def test_auto_adjust_job_runs_only_picked_files_and_finalizes_failures(self):
        cards = self.scanned_only()
        other, _ = self.add(self.paths[1], "A3", 3)
        batch._hand_over(self.paths[1])
        options = batch.start_auto_adjust([cards["A5"]["id"]])
        self.wf.STATE["busy"] = True
        self.wf.STATE["job"] = self.job = {"kind": "batch_scan", "options": options, "completed": [],
                                           "pending_adjustment": None, "status": "running"}
        order = []

        def back(excel, session, path, picked=False):
            order.append(("back", path, picked))
            for addr in ("A5", "B5"):
                cards[addr].update(waiting=True, manual=False, finalized=False)
            self.job["picked_back"] = [cards["A5"]["id"], cards["B5"]["id"]]

        def repair(excel, session, path):
            order.append(("fix", path))
            cards["A5"].update(color="blue", verified=True, waiting=False)
            cards["B5"].update(color="yellow", manual=True, waiting=False, message="自动试修未通过")

        with patch.object(batch, "_back_to_auto", side_effect=back), patch.object(batch, "_refresh"), \
                patch.object(batch, "_repair_file", side_effect=repair), \
                patch.object(batch, "_sync_highlights_safely") as sync:
            self.assertEqual(batch.run(), [self.paths[0]])
        self.assertEqual(order, [("back", self.paths[0], True), ("fix", self.paths[0])])
        self.assertEqual((cards["A5"]["color"], cards["A5"].get("finalized")), ("blue", False))
        # 没修好、没写过文件：直接算已提交，可以马上人工处理、复检
        self.assertEqual((cards["B5"]["color"], cards["B5"]["finalized"]), ("yellow", True))
        self.assertEqual(other["finalized"], True)
        sync.assert_called_once_with(self.session, self.paths[0])
        self.assertEqual(batch._data()["auto_versions"][self.paths[0]],
                         self.wf._file_version(Path(self.paths[0])))

    def test_colors_can_change_while_auto_adjust_runs_except_files_it_still_needs(self):
        cards = self.scanned_only()
        other, _ = self.add(self.paths[1], "A3", 3)
        batch._hand_over(self.paths[1])
        self.wf.STATE["busy"] = True
        self.wf.STATE["job"] = {"kind": "batch_scan", "options": {"picked": True, "back_to_auto": True, "card_ids": []},
                                "completed": [], "scanned": [], "pending_adjustment": None, "status": "running",
                                "workbooks": [self.paths[0]]}
        batch.mark(other["id"], "green")  # 不在这次任务里的文件：照常改
        self.assertEqual(other["color"], "green")
        with self.assertRaisesRegex(ValueError, "才能改颜色"):
            batch.mark(cards["A5"]["id"], "green")
        with self.assertRaisesRegex(ValueError, "a.xlsx"):
            batch.mark_many([other["id"], cards["A5"]["id"]], "yellow")
        self.assertEqual(other["color"], "green")  # 一张不行就全部不改
        self.wf.STATE["job"]["completed"] = [self.paths[0]]
        batch.mark(cards["A5"]["id"], "green")  # 做完的文件也能改了
        # 提交等其他任务照旧等它结束
        self.wf.STATE["job"] = {"kind": "batch_commit", "options": {"card_ids": []}, "completed": [],
                                "workbooks": [self.paths[0]]}
        with self.assertRaisesRegex(ValueError, "结束或暂停后再审核"):
            batch.mark(other["id"], "yellow")

    def test_auto_adjust_clicked_during_a_running_one_is_queued_into_it(self):
        cards = self.scanned_only()
        other, item = self.add(self.paths[1], "A3", 3)
        item.update(page_index=0, bbox=[0, 0, 1, 1])
        batch._hand_over(self.paths[1])
        self.assertFalse(batch.enqueue_auto_adjust([other["id"]]))  # 没有在跑的：另起任务
        options = batch.start_auto_adjust([cards["A5"]["id"]])
        self.wf.STATE["busy"] = True
        self.wf.STATE["job"] = self.job = {"kind": "batch_scan", "options": options, "completed": [],
                                           "pending_adjustment": None, "status": "running"}
        order = []

        def back(excel, session, path, picked=False):
            order.append(("back", path))
            if path == self.paths[0]:
                # 第一个文件还在跑：同一文件不能再排，另一个文件排进来
                with self.assertRaisesRegex(ValueError, "再点自动调整"):
                    batch.enqueue_auto_adjust([cards["C7"]["id"]])
                self.assertTrue(batch.enqueue_auto_adjust([other["id"]]))
                self.assertTrue(other["auto_pick"])

        with patch.object(batch, "_back_to_auto", side_effect=back), patch.object(batch, "_refresh"), \
                patch.object(batch, "_repair_file", side_effect=lambda e, s, path: order.append(("fix", path))), \
                patch.object(batch, "_sync_highlights_safely"), patch.object(batch, "add_to_execution_batch") as add:
            self.assertEqual(batch.run(), self.paths)
        self.assertEqual(order, [("back", self.paths[0]), ("fix", self.paths[0]),
                                 ("back", self.paths[1]), ("fix", self.paths[1])])
        add.assert_called_once_with(Path(self.wf.STATE["folder"]), [Path(self.paths[1])])
        self.assertEqual(self.job["completed"], self.paths)
        self.assertIn(other["id"], self.job["options"]["card_ids"])
        # 收尾后再点就另起任务
        self.assertTrue(self.job["closing"])
        self.assertFalse(batch.enqueue_auto_adjust([other["id"]]))

    def test_finished_file_can_be_queued_again_and_failed_file_is_not_retried_forever(self):
        cards = self.scanned_only()
        options = batch.start_auto_adjust([cards["A5"]["id"]])
        self.wf.STATE["busy"] = True
        self.wf.STATE["job"] = self.job = {"kind": "batch_scan", "options": options, "completed": [],
                                           "pending_adjustment": None, "status": "running"}
        order = []
        other, item = self.add(self.paths[1], "A3", 3)
        item.update(page_index=0, bbox=[0, 0, 1, 1])
        batch._hand_over(self.paths[1])

        def back(excel, session, path, picked=False):
            order.append(("back", path))
            if path == self.paths[1]:
                raise ValueError("文件被占用")

        def repair(excel, session, path):
            order.append(("fix", path))
            if order.count(("fix", self.paths[0])) == 1:
                # 第一个文件修完前把第二个文件排进来；第二个出错跳过，不会一直重试
                self.assertTrue(batch.enqueue_auto_adjust([other["id"]]))

        with patch.object(batch, "_back_to_auto", side_effect=back), patch.object(batch, "_refresh"), \
                patch.object(batch, "_repair_file", side_effect=repair), patch.object(batch, "_finish_picked"), \
                patch.object(batch, "_sync_highlights_safely"), patch.object(batch, "add_to_execution_batch"):
            batch.run()
        self.assertEqual(order, [("back", self.paths[0]), ("fix", self.paths[0]), ("back", self.paths[1])])
        self.assertEqual(self.job["completed"], [self.paths[0]])
        # 做完的文件可以再排：从做完里拿出来重新做
        self.job.update(closing=False)
        with patch.object(batch, "add_to_execution_batch"):
            self.assertTrue(batch.enqueue_auto_adjust([cards["A5"]["id"]]))
        self.assertNotIn(self.paths[0], self.job["completed"])
        self.assertNotIn(self.paths[0], self.job["scanned"])
        self.assertEqual(self.job["rounds"][self.paths[0]], 1)

    def scan_only_job(self):
        """「只扫描」正在跑：两个文件各一处问题，扫到时转成黄色人工项"""
        self.job["options"] = {"folder": str(self.root), "scan_only": True, "highlight": dict(batch.DEFAULT_HIGHLIGHT)}
        cards = []
        for path, addr in ((self.paths[0], "A5"), (self.paths[1], "A3")):
            card, item = self.add(path, addr, int(addr[1:]))
            item.update(page_index=0, bbox=[0, 0, 1, 1])
            cards.append(card)
        return cards

    def test_auto_adjust_clicked_during_scan_only_runs_after_all_files_are_scanned(self):
        first, second = self.scan_only_job()
        order = []

        def refresh(excel, path, *a, **k):
            order.append(("scan", path))
            if path == self.paths[1]:
                # 扫第二个文件时：扫完的第一个文件可以排队，正在扫的不行
                with self.assertRaisesRegex(ValueError, "再点自动调整"):
                    batch.enqueue_auto_adjust([second["id"]])
                self.assertTrue(batch.enqueue_auto_adjust([first["id"]]))
                self.assertTrue(first["auto_pick"])
                # 排队的文件后台还要读、还要改：先不能打开、不能改颜色
                self.assertIn("排队等自动调整", batch.open_block_reason(self.job, Path(self.paths[0])))
                with self.assertRaisesRegex(ValueError, "才能改颜色"):
                    batch.mark(first["id"], "green")

        with patch.object(batch, "_refresh", side_effect=refresh), \
                patch.object(batch, "_back_to_auto", side_effect=lambda e, s, path, picked=False:
                             order.append(("back", path, picked))), \
                patch.object(batch, "_repair_file", side_effect=lambda e, s, path: order.append(("fix", path))), \
                patch.object(batch, "_finish_picked"), patch.object(batch, "_sync_highlights_safely"):
            self.assertEqual(batch.run(), [self.paths[0]])
        # 两个文件都扫完，才接着自动调整排队的文件
        self.assertEqual(order, [("scan", self.paths[0]), ("scan", self.paths[1]),
                                 ("back", self.paths[0], True), ("fix", self.paths[0])])
        self.assertEqual(second["color"], "yellow")  # 没排队的照旧留给人工
        # 同一个任务就地改成自动调整；全部扫完了，没有要放回的扫描任务
        self.assertEqual(self.job["options"], {"folder": str(self.root), "back_to_auto": True, "picked": True,
                                               "card_ids": [first["id"]]})
        self.assertEqual((self.job["workbooks"], self.job["completed"]), ([self.paths[0]], [self.paths[0]]))
        self.assertNotIn("auto_queue", self.job)
        self.assertNotIn("paused_scan", self.job)
        self.assertTrue(self.job["closing"])

    def test_scan_only_without_queue_finishes_as_before(self):
        self.scan_only_job()
        with patch.object(batch, "_refresh"), patch.object(batch, "_back_to_auto") as back:
            self.assertEqual(batch.run(), self.paths)
        back.assert_not_called()
        self.assertTrue(self.job["options"]["scan_only"])
        self.assertEqual(self.job["completed"], self.paths)
        self.assertTrue(self.job["closing"])  # 收尾时再点自动调整另起任务
        self.assertFalse(batch.enqueue_auto_adjust([]))

    def test_unscanned_files_can_still_be_resumed_after_queued_auto_adjust(self):
        first, _ = self.scan_only_job()

        def refresh(excel, path, *a, **k):
            if path == self.paths[1]:
                self.assertTrue(batch.enqueue_auto_adjust([first["id"]]))
                raise ValueError("文件被占用")

        with patch.object(batch, "_refresh", side_effect=refresh), patch.object(batch, "_back_to_auto"), \
                patch.object(batch, "_repair_file"), patch.object(batch, "_finish_picked"), \
                patch.object(batch, "_sync_highlights_safely"), \
                patch.object(self.wf, "interrupt_execution_batch") as interrupt:
            self.wf.run_pending_workflow()
        # 自动调整做完：没扫完的「只扫描」放回来，可以「继续上次任务」接着扫出错的文件
        job = self.wf.STATE["job"]
        self.assertEqual((job["status"], job["options"]["scan_only"]), ("interrupted", True))
        self.assertEqual(job["completed"], [self.paths[0]])
        self.assertIn(self.paths[1], job["file_errors"])
        self.assertNotIn("auto_queue", job)
        self.assertIn("继续上次任务", interrupt.call_args.args[0])
        self.assertEqual(batch._data()["auto_versions"][self.paths[0]], self.wf._file_version(Path(self.paths[0])))

    def test_pause_during_scan_only_keeps_the_queue_for_resume(self):
        first, _ = self.scan_only_job()

        def refresh(excel, path, *a, **k):
            if path == self.paths[1]:
                self.assertTrue(batch.enqueue_auto_adjust([first["id"]]))
                self.job["pause_requested"] = True

        with patch.object(batch, "_refresh", side_effect=refresh), patch.object(batch, "_back_to_auto") as back:
            batch.run()
        back.assert_not_called()
        self.assertEqual(self.job["auto_queue"], [first["id"]])
        self.assertTrue(self.job["options"]["scan_only"])

    def handed_over_cards(self):
        """两个文件各一张已提交的黄卡（只扫描转人工后的样子）"""
        cards = self.scan_only_job()
        for path in self.paths:
            batch._hand_over(path)
        return cards

    def test_recheck_and_commit_clicked_during_another_task_are_queued_behind_it(self):
        first, second = self.handed_over_cards()
        self.wf.STATE["busy"] = False
        self.assertFalse(batch.enqueue_followup("batch_recheck", [first["id"]]))  # 后台空着：调用方直接开始
        self.wf.STATE["busy"] = True
        self.job["completed"] = [self.paths[0]]  # 第一个文件扫完，第二个还在扫
        with self.assertRaisesRegex(ValueError, "b.xlsx：.*再点复检"):
            batch.enqueue_followup("batch_recheck", [second["id"]])
        self.assertTrue(batch.enqueue_followup("batch_recheck", [first["id"]]))
        self.assertTrue(batch.enqueue_followup("batch_recheck", [first["id"]], single=True))  # 和上一个复检合并
        self.assertTrue(batch.enqueue_followup("batch_commit", [first["id"]], color="yellow"))
        self.assertTrue(batch.enqueue_followup("batch_recheck", [first["id"]], single=True))  # 隔着提交：另排
        self.assertEqual(batch._data()["followups"], [
            {"kind": "batch_recheck", "card_ids": [first["id"]], "full_paths": [self.paths[0]]},
            {"kind": "batch_commit", "card_ids": [first["id"]], "color": "yellow"},
            {"kind": "batch_recheck", "card_ids": [first["id"]], "full_paths": []},
        ])
        self.assertEqual(len(batch.snapshot()["followups"]), 3)
        # 别的页面的任务（精简 Sheet 等）不排
        self.job["kind"] = "clean_sheets"
        with self.assertRaisesRegex(ValueError, "后台正在精简 Sheet.*结束后才能提交"):
            batch.enqueue_followup("batch_commit", [first["id"]])
        self.assertEqual(batch.cancel_followups(), 3)
        self.assertEqual(batch.snapshot()["followups"], [])

    def run_with_followups(self, followups, *, paused=False):
        """跑一次任务（batch.run 换成只记下任务种类），结束后看排队的复检 / 提交怎么接着开始"""
        batch._data()["followups"] = followups
        ran = []

        def fake_run():
            job = self.wf.STATE["job"]
            ran.append((job["kind"], dict(job["options"])))
            if paused and len(ran) == 1:
                job["pause_requested"] = True
                return list(job["workbooks"])  # 暂停：文件没做完，任务停在「可继续」
            return []

        with patch.object(batch, "run", side_effect=fake_run), \
                patch.object(self.wf, "interrupt_execution_batch"), patch.object(self.wf, "suspend_report_updates"):
            self.wf.run_pending_workflow()
        return ran

    def test_queued_followups_start_in_order_when_the_task_ends(self):
        first, second = self.handed_over_cards()
        second.update(color="blue", finalized=False, marked=True)  # 排队后又改成蓝：复检不再查它
        ran = self.run_with_followups([
            {"kind": "batch_recheck", "card_ids": [second["id"]], "full_paths": [self.paths[1]]},
            {"kind": "batch_recheck", "card_ids": [first["id"]], "full_paths": [self.paths[0], self.paths[1]]},
            {"kind": "batch_commit", "card_ids": [second["id"]], "color": None},
        ])
        self.assertEqual([kind for kind, _ in ran], ["batch_scan", "batch_recheck", "batch_commit"])
        self.assertEqual(ran[1][1], {"card_ids": [first["id"]], "full_paths": [self.paths[0]]})
        self.assertEqual(ran[2][1], {"card_ids": [second["id"]]})  # 提交按到时的颜色
        self.assertEqual(batch._data()["followups"], [])
        self.assertFalse(self.wf.STATE["busy"])
        self.assertTrue(any("已经没有要复检的卡" in line for line in self.wf.STATE["logs"]))

    def test_followups_wait_when_you_pause_and_can_start_now(self):
        first, _ = self.handed_over_cards()
        queued = [{"kind": "batch_recheck", "card_ids": [first["id"]], "full_paths": [self.paths[0]]}]
        ran = self.run_with_followups(list(queued), paused=True)
        self.assertEqual([kind for kind, _ in ran], ["batch_scan"])
        self.assertEqual(batch._data()["followups"], queued)  # 暂停了：留着，页面上可以现在开始或取消
        with self.wf.STATE_LOCK:
            self.assertTrue(batch.start_followup({}))  # 「现在开始」
        self.assertEqual((self.wf.STATE["job"]["kind"], self.wf.STATE["busy"]), ("batch_recheck", True))
        self.assertEqual(self.wf.STATE["job"]["paused_scan"]["kind"], "batch_scan")  # 暂停的扫描收起来，之后还能继续

    def test_open_all_during_a_task_opens_finished_files_and_lists_the_rest(self):
        first, second = self.handed_over_cards()
        self.job["completed"] = [self.paths[0]]
        with patch.object(batch, "open_workbooks_for_review", return_value={}) as opener, \
                patch.object(batch.autofix, "wait_for_writes"):
            result = batch.open_ready([first["id"], second["id"]])
        opener.assert_called_once_with([(Path(self.paths[0]), "S1", "A5", 5)])
        self.assertEqual(result["opened"], ["a.xlsx"])
        self.assertEqual([w["name"] for w in result["waiting"]], ["b.xlsx"])
        self.assertIn("才能打开", result["waiting"][0]["reason"])
        self.assertEqual(result["errors"], [])

    def test_file_over_the_scan_limit_is_skipped_for_manual_check(self):
        first, second = self.scan_only_job()
        watchdog = Mock(deadline_hit=False, timed_out=False)
        limits = []

        def refresh(excel, path, *a, **k):
            limits.append(self.job["scan_limit"]["seconds"])
            if path == self.paths[0]:
                # 卡在第二张表：看门狗到点结束 Excel，COM 调用报错
                self.wf.SCAN_NOW.update(path=path, sheet="レイアウト", index=2, total=5, started=time.time() - 600)
                watchdog.deadline_hit = True
                raise RuntimeError("RPC サーバーを利用できません。")
            watchdog.deadline_hit = False

        with patch.object(batch, "ExcelWatchdog", return_value=watchdog), \
                patch.object(batch, "_refresh", side_effect=refresh), patch.object(batch, "_hand_over") as hand:
            self.assertEqual(batch.run(), self.paths)
        self.assertEqual(limits, [batch.SCAN_LIMIT_FIRST, batch.SCAN_LIMIT_FIRST])  # 还没扫完 2 份：先给 10 分钟
        watchdog.deadline.assert_any_call(batch.SCAN_LIMIT_FIRST)
        # 超时的文件不重试、不转黄卡，算这次做完（继续上次任务也不再卡在它上面）
        hand.assert_called_once_with(self.paths[1])
        self.assertEqual(self.job["completed"], self.paths)
        self.assertNotIn(self.paths[0], self.job.get("file_errors", {}))
        self.assertEqual(first["color"], "yellow")
        slow = batch.snapshot()["slow_files"]
        self.assertEqual([(s["name"], s["sheet"], s["limit"]) for s in slow], [("a.xlsx", "レイアウト", 600)])
        self.assertTrue(any("扫描超过 10 分钟，卡在表「レイアウト」" in line for line in self.wf.STATE["logs"]))
        batch.clear_slow_files()
        self.assertEqual(batch.snapshot()["slow_files"], [])

    def test_successful_scan_clears_an_old_slow_file_notice(self):
        self.scan_only_job()
        batch._data()["slow_files"] = {self.paths[0]: {"sheet": "S1", "limit": 180, "at": ""}}
        with patch.object(batch, "_refresh"):
            batch.run()
        self.assertEqual(batch._data()["slow_files"], {})
        self.assertEqual(set(self.job["scan_seconds"]), set(self.paths))

    def test_scan_limit_is_estimated_from_files_already_scanned(self):
        job = {"options": {}, "scan_seconds": {self.paths[0]: 20.0}}
        self.assertEqual(batch._scan_limit(job, self.paths[1]), batch.SCAN_LIMIT_FIRST)  # 只扫完 1 份：估不准
        job["scan_seconds"][self.paths[1]] = 40.0  # 两份一样大，平均 30 秒
        self.assertEqual(batch._scan_limit(job, self.paths[0]), batch.SCAN_LIMIT_FLOOR)  # 5 倍才 150 秒：最少 3 分钟
        job["scan_seconds"] = {self.paths[0]: 100.0, self.paths[1]: 140.0}
        self.assertEqual(batch._scan_limit(job, self.paths[0]), 600)  # 平均 120 秒 × 5
        job["scan_seconds"] = {self.paths[0]: 400.0, self.paths[1]: 400.0}
        self.assertEqual(batch._scan_limit(job, self.paths[0]), 900)  # 最多和看门狗的无进展上限一样
        job["options"]["scan_limit_minutes"] = 5
        self.assertEqual(batch._scan_limit(job, self.paths[0]), 300)  # 网页上选了固定分钟数

    def test_scan_eta_uses_finished_files_and_the_current_one(self):
        self.job.update(options={"folder": str(self.root), "scan_only": True}, scan_seconds={self.paths[0]: 10.0},
                        scanned=[self.paths[0]], completed=[self.paths[0]],
                        scan_limit={"path": self.paths[1], "seconds": 180})
        self.wf.SCAN_NOW.clear()
        self.wf.SCAN_NOW.update(path=self.paths[1], sheet="S2", index=3, total=8, started=time.time() - 4)
        eta = batch.snapshot()["scan_eta"]
        self.assertEqual((eta["elapsed"], eta["remaining"]), (14, 6))  # 两份一样大：这份估计 10 秒，已用 4 秒
        self.assertEqual(eta["current"], {"name": "b.xlsx", "sheet": "S2", "index": 3, "total": 8, "seconds": 4,
                                          "limit": 180})
        self.job["scan_seconds"] = {}
        self.assertIsNone(batch.snapshot()["scan_eta"]["remaining"])  # 一份都没扫完：不估
        self.job["phase"] = "repair"
        self.assertIsNone(batch.snapshot()["scan_eta"])  # 全自动的修复阶段不算扫描

    def test_web_queues_recheck_and_commit_and_opens_ready_files_during_a_task(self):
        import web_clip_fixer

        first, second = self.handed_over_cards()
        self.job["completed"] = [self.paths[0]]
        importlib.reload(web_clip_fixer)  # 换成本测试的 STATE
        server = ThreadingHTTPServer(("127.0.0.1", 0), web_clip_fixer.WebHandler)
        threading.Thread(target=server.serve_forever, daemon=True).start()
        self.addCleanup(server.server_close)
        self.addCleanup(server.shutdown)

        def post(path, payload):
            connection = http.client.HTTPConnection("127.0.0.1", server.server_port, timeout=5)
            try:
                connection.request("POST", path, json.dumps(payload), {"Content-Type": "application/json"})
                return json.loads(connection.getresponse().read())
            finally:
                connection.close()

        self.assertEqual(post("/api/batch/recheck", {"card_ids": [first["id"]]}), {"ok": True, "queued": True})
        self.assertEqual(post("/api/batch/commit", {"card_ids": [first["id"]], "color": "yellow"}),
                         {"ok": True, "queued": True})
        self.assertIn("再点复检", post("/api/batch/recheck", {"card_ids": [second["id"]]})["error"])
        with patch.object(batch, "open_workbooks_for_review", return_value={}), \
                patch.object(batch.autofix, "wait_for_writes"):
            opened = post("/api/batch/open", {"card_ids": [first["id"], second["id"]]})
        self.assertEqual((opened["opened"], [w["name"] for w in opened["waiting"]]), (["a.xlsx"], ["b.xlsx"]))
        self.assertIn("后台可能还在忙", post("/api/batch/followups/start", {})["error"])
        self.assertEqual(post("/api/batch/followups/cancel", {}), {"ok": True, "cancelled": 2})
        self.assertIn("扫描上限", post("/api/batch/start", {"folder": str(self.root), "scan_limit_minutes": 7})["error"])

    def test_fix_error_types_filters_autofix_and_allows_reenabling_later(self):
        from core import autofix
        from core.hidden_text import REASON_CLIPPED, REASON_HASH, REASON_MISSING, REASON_TRUNCATED

        self.assertEqual(autofix.item_error_type({"kind": "clip"}), "border_clip")
        self.assertEqual(autofix.item_error_type({"kind": "hidden", "hidden_reason": REASON_TRUNCATED}), "text_truncated")
        self.assertEqual(autofix.item_error_type({"kind": "hidden", "hidden_reason": REASON_CLIPPED}), "box_overflow")
        self.assertEqual(autofix.item_error_type({"kind": "hidden", "hidden_reason": REASON_HASH}), "hash_overflow")
        self.assertEqual(autofix.item_error_type({"kind": "hidden", "hidden_reason": REASON_MISSING}), "text_missing")
        self.assertEqual(autofix.item_error_type({"kind": "pagebreak"}), "page_break_cut")

        # 只勾选「边框压线」，不勾选「文字截断」和「分页切开合并单元格」
        batch.set_fix_error_types(["border_clip"])
        self.job["options"]["fix_error_types"] = ["border_clip"]
        clip_card, clip_item = self.add(self.paths[0], "A5", 5)
        clip_item.update(kind="clip", error_type="border_clip", page_index=0, bbox=[0, 0, 10, 10])
        batch._remember_problems(clip_card, [clip_item])
        trunc_card, trunc_item = self.add(self.paths[0], "B5", 5)
        trunc_item.update(kind="hidden", hidden_reason=REASON_TRUNCATED, error_type="text_truncated",
                          page_index=0, bbox=[0, 0, 10, 10])
        batch._remember_problems(trunc_card, [trunc_item])
        brk_card, brk_item = self.add(self.paths[0], "C8:D10", 8)
        brk_item.update(kind="pagebreak", error_type="page_break_cut", block={"top": 8, "bottom": 10, "left": 3})
        brk_card["kind"] = "pagebreak"
        batch._remember_problems(brk_card, [brk_item])

        prepared = []
        def fake_prepare(_session, _path, items, **_):
            prepared.append([it["cell_addr"] for it in items])
            return 18.0, {"height": 18.0, "shots": {}}, {}

        with patch.object(batch, "_prepare_row", side_effect=fake_prepare), \
                patch.object(batch, "_repair_break") as repair_break, \
                patch.object(batch, "_refresh"):
            batch._repair_file(Mock(), self.session, self.paths[0])

        # 同一行的 A5（边框压线）照常自动修；B5（文字截断）和 C8:D10（分页切开）因未勾选而转黄留给人工
        self.assertEqual(prepared, [["A5"]])
        repair_break.assert_not_called()
        self.assertEqual((clip_card["color"], clip_card["verified"]), ("blue", True))
        self.assertEqual((trunc_card["color"], trunc_card["manual"], trunc_card["handed_over"]), ("yellow", True, True))
        self.assertIn("未勾选「文字截断」", trunc_card["message"])
        self.assertEqual((brk_card["color"], brk_card["manual"], brk_card["handed_over"]), ("yellow", True, True))
        self.assertIn("未勾选「分页切开合并单元格」", brk_card["message"])

        # 提交标记后（蓝卡移出列表，黄卡 finalized），未勾选时点自动调整或转回自动都会被拦截并说明原因
        clip_card["dismissed"] = True
        trunc_card["finalized"] = True
        brk_card["finalized"] = True
        self.wf.STATE["busy"] = False
        self.job["status"] = "completed"
        plan = batch.auto_adjust_plan([trunc_card["id"], brk_card["id"]])
        self.assertEqual(plan["run"], [])
        self.assertEqual(len(plan["skip"]), 2)
        with self.assertRaisesRegex(ValueError, "未勾选的错误类型"):
            batch.check_auto()

        # 重新勾上「文字截断」后，立刻可以自动调整或转回自动
        batch.set_fix_error_types(["border_clip", "text_truncated"])
        self.assertTrue(batch._mode()["to_auto"])
        plan2 = batch.auto_adjust_plan([trunc_card["id"], brk_card["id"]])
        self.assertEqual(plan2["run"], [trunc_card["id"]])
        self.assertEqual([s["id"] for s in plan2["skip"]], [brk_card["id"]])

        with self.assertRaisesRegex(ValueError, "至少勾选一种"):
            batch.set_fix_error_types([], require_non_empty=True)


if __name__ == "__main__":
    unittest.main()
