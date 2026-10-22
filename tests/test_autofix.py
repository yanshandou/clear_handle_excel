import importlib
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import storage


class FakeRow:
    def __init__(self, heights, row):
        self.heights, self.row = heights, row

    @property
    def RowHeight(self):
        return self.heights.get(self.row, 15.0)

    @RowHeight.setter
    def RowHeight(self, value):
        self.heights[self.row] = value


class FakeSheet:
    def __init__(self, book, needed):
        self.book, self.needed = book, needed
        self.Name = "S1"
        self.Visible = -1

    def Rows(self, row):
        return FakeRow(self.book.heights, row)

    def ExportAsFixedFormat(self, **kwargs):
        from core.pdf_detector import fitz
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((60, 100), "Target", fontsize=12)
            if self.book.heights.get(5, 15.0) < self.needed:
                page.draw_line((55, 96), (250, 96), width=0.5)  # 行高不够，下边框压到字上
            doc.save(kwargs["Filename"])


class FakeBook:
    def __init__(self, test, path, heights):
        self.test, self.path, self.heights = test, path, heights
        self.saved = False

    def Sheets(self, name):
        return FakeSheet(self, self.test.needed)

    def Save(self):
        self.saved = True

    def Close(self, SaveChanges):
        pass


class AutofixTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        patcher = patch.object(storage, "DATABASE_PATH", self.root / "test.sqlite3")
        patcher.start()
        self.addCleanup(patcher.stop)
        from core import autofix, reporter, workflow
        self.workflow = importlib.reload(workflow)
        reporter_patch = patch.object(reporter, "RECORD_PATH", self.root / "record.md")
        reporter_patch.start()
        self.addCleanup(reporter_patch.stop)
        self.autofix = importlib.reload(autofix)
        self.autofix._worker = Mock(is_alive=Mock(return_value=True))  # 测试里不启动后台线程
        self.autofix.AUTOFIX["active_until"] = float("inf")  # 相当于页面一直开着「② 自动处理」
        self.autofix.AUTOFIX["auto"] = False  # 下面多数用例测逐张点选；默认全自动的流程见 test_default_auto_*

        self.wb_path = self.root / "book.xlsx"
        self.wb_path.write_bytes(b"workbook")
        self.needed = 21.0
        self.real_heights = {5: 15.0}
        self.books = []

        def open_book(path, **_):
            heights = self.real_heights if Path(path) == self.wb_path.resolve() else dict(self.real_heights)
            book = FakeBook(self, Path(path), heights)
            self.books.append(book)
            return book
        excel = Mock()
        excel.Workbooks.Open.side_effect = open_book
        for name, value in {
            "start_background_excel": Mock(return_value=excel),
            "extract_sheet_snapshot": Mock(return_value={"start_r": 5, "start_c": 2, "rows_vals": [("Target",)]}),
            "normalize_workbook_view": Mock(),
        }.items():
            p = patch.object(self.autofix, name, value)
            p.start()
            self.addCleanup(p.stop)

        from core.pdf_detector import fitz
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((60, 100), "Target", fontsize=12)
            span = page.get_text("dict")["blocks"][0]["lines"][0]["spans"][0]
            bbox = tuple(span["bbox"])
        self.workflow.STATE.update(busy=False, folder=str(self.root), items={"i1": {
            "id": "i1", "seq": 1, "kind": "clip", "wb_path": str(self.wb_path), "wb_name": "book.xlsx",
            "sheet_name": "S1", "excel_row": 5, "cell_addr": "B5", "row_height": 15.0, "page_index": 0,
            "text": "Target", "bbox": bbox, "occurrence": None, "status": "pending", "iterations": 0, "triage_keep": True,
        }})
        self.session = self.autofix._ExcelSession()
        self.addCleanup(self.session.close)

    def group(self):
        self.autofix.sync_groups()
        (group,) = self.autofix.AUTOFIX["groups"].values()
        return group

    def test_preview_stops_one_step_after_success_and_recommends_it(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        group = self.group()
        self.assertEqual(
            [(a["height"], a["ok"]) for a in group["attempts"]],
            [(15.0, False), (18.0, False), (21.0, True), (24.0, True)],
        )
        self.assertEqual((group["status"], group["recommended"]), ("ready", 3))
        self.assertTrue(all(a["image"] for a in group["attempts"]))
        self.assertIsNotNone(self.autofix.attempt_image(gid, 0))
        self.assertEqual(self.real_heights, {5: 15.0})  # 预览只动副本
        self.assertNotIn("image", str(self.autofix.snapshot()["groups"][0]["attempts"]))

        with self.assertRaisesRegex(ValueError, "已显示全"):
            self.autofix.choose(gid, 1)
        with self.assertRaisesRegex(ValueError, "已显示全"):
            self.autofix.choose(gid, 0)
        self.autofix.choose(gid, 2)
        self.assertTrue(self.autofix._apply(self.session, str(self.wb_path), [gid]))
        self.assertEqual(self.real_heights, {5: 21.0})
        self.assertTrue(self.books[-1].saved)
        item = self.workflow.STATE["items"]["i1"]
        self.assertEqual((item["status"], item["row_height"]), ("fixed", 21.0))
        self.assertEqual(self.autofix.AUTOFIX["groups"][gid]["status"], "applied")

    def test_gives_up_after_nine_pictures(self):
        self.needed = 999.0
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        group = self.group()
        self.assertEqual(len(group["attempts"]), 9)
        self.assertEqual((group["status"], group["recommended"]), ("failed", None))
        self.assertIn("人工处理", group["message"])

    def test_skip_failed_row_saves_all_problems_without_changing_excel(self):
        self.needed = 999.0
        self.workflow.STATE["items"]["i2"] = dict(self.workflow.STATE["items"]["i1"], id="i2", seq=2)
        group = self.group()
        gid = group["id"]
        self.autofix._preview(self.session, gid)

        result = self.autofix.skip_group(gid, ["i1", "i2"])
        self.assertEqual(result, {"changed": 2, "errors": []})
        saved = storage.load_snapshot("workflow")
        self.assertEqual([saved["items"][iid]["status"] for iid in ("i1", "i2")], ["accepted", "accepted"])
        self.assertNotIn(gid, self.autofix.AUTOFIX["groups"])
        self.assertIsNone(self.autofix._next_task({}))
        self.assertEqual(self.real_heights, {5: 15.0})
        self.assertTrue(all(not book.saved for book in self.books))

    def test_skip_during_preview_stops_it_without_writing(self):
        gid = self.group()["id"]
        evaluate = self.autofix._evaluate

        def skip_after_export(*args):
            result = evaluate(*args)
            self.autofix.skip_group(gid, ["i1"])
            return result

        with patch.object(self.autofix, "_evaluate", side_effect=skip_after_export):
            with self.assertRaises(self.autofix._Paused):
                self.autofix._preview(self.session, gid)
        self.assertNotIn(gid, self.autofix.AUTOFIX["groups"])
        self.assertEqual(self.real_heights, {5: 15.0})
        self.assertTrue(all(not book.saved for book in self.books))
        with self.assertRaises(self.autofix._Paused):
            self.autofix._preview(self.session, gid)  # 已取出的旧预览任务也安静停止

    def test_skip_rejects_row_already_queued_for_writing(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.autofix.choose(gid, 3)
        with self.assertRaisesRegex(ValueError, "开始写入"):
            self.autofix.skip_group(gid, ["i1"])
        self.assertEqual(self.workflow.STATE["items"]["i1"]["status"], "pending")
        self.assertEqual(self.autofix.AUTOFIX["groups"][gid]["status"], "applying")

    def test_skip_rejects_stale_items_and_busy_tasks(self):
        gid = self.group()["id"]
        with self.assertRaisesRegex(ValueError, "已更新"):
            self.autofix.skip_group(gid, ["old"])
        self.workflow.STATE["busy"] = True
        with self.assertRaisesRegex(ValueError, "任务正在执行"):
            self.autofix.skip_group(gid, ["i1"])
        self.assertEqual(self.workflow.STATE["items"]["i1"]["status"], "pending")

    def test_skip_database_failure_keeps_row_in_queue(self):
        gid = self.group()["id"]
        with patch.object(self.workflow, "_persist_state", side_effect=RuntimeError("database unavailable")):
            with self.assertRaisesRegex(RuntimeError, "database unavailable"):
                self.autofix.skip_group(gid, ["i1"])
        self.assertEqual(self.workflow.STATE["items"]["i1"]["status"], "pending")
        self.assertIn(gid, self.autofix.AUTOFIX["groups"])

    def test_group_sync_cannot_reintroduce_a_skipped_row_from_old_state(self):
        gid = self.group()["id"]
        reading_done = threading.Event()
        continue_sync = threading.Event()
        original_lock = self.autofix._LOCK
        errors = []

        class PausedGroupLock:
            def __enter__(self):
                if threading.current_thread().name == "test-group-sync":
                    reading_done.set()
                    if not continue_sync.wait(5):
                        raise AssertionError("sync did not resume")
                return original_lock.__enter__()

            def __exit__(self, *args):
                return original_lock.__exit__(*args)

        def sync():
            try:
                self.autofix.sync_groups()
            except Exception as error:
                errors.append(error)

        with patch.object(self.autofix, "_LOCK", PausedGroupLock()):
            thread = threading.Thread(target=sync, name="test-group-sync")
            thread.start()
            try:
                self.assertTrue(reading_done.wait(5))
                acquired = self.workflow.STATE_LOCK.acquire(blocking=False)
                if acquired:
                    self.workflow.STATE_LOCK.release()
                self.assertFalse(acquired, "同步分组期间必须锁住问题状态，防止跳过判断被旧快照覆盖")
            finally:
                continue_sync.set()
                thread.join(5)
        self.assertFalse(thread.is_alive())
        self.assertEqual(errors, [])
        self.autofix.skip_group(gid, ["i1"])
        self.autofix.sync_groups()
        self.assertNotIn(gid, self.autofix.AUTOFIX["groups"])

    def test_full_auto_applies_recommendation(self):
        gid = self.group()["id"]
        self.autofix.set_auto(True)
        self.autofix._preview(self.session, gid)
        group = self.group()
        self.assertEqual((group["status"], group["chosen"]), ("applying", 3))
        self.assertTrue(self.autofix._apply(self.session, str(self.wb_path), [gid]))
        self.assertEqual(self.real_heights, {5: 24.0})

    def written(self, gid):
        self.assertTrue(self.autofix._apply(self.session, str(self.wb_path), [gid]))
        return self.autofix.AUTOFIX["groups"][gid]

    def test_default_auto_writes_then_waits_for_confirmation(self):
        self.autofix.AUTOFIX["auto"] = True
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        group = self.written(gid)
        self.assertEqual((group["status"], group["confirmed"], self.real_heights), ("applied", False, {5: 24.0}))
        item = self.workflow.STATE["items"]["i1"]
        self.assertEqual((item["status"], item["autofix_unconfirmed"], item["autofix_base_height"]), ("fixed", True, 15.0))
        self.assertEqual(self.autofix.confirm_groups([gid]), 1)
        self.assertTrue(group["confirmed"])
        self.assertNotIn("autofix_unconfirmed", self.workflow.STATE["items"]["i1"])
        self.assertNotIn("autofix_unconfirmed", str(storage.load_snapshot("workflow")["items"]["i1"]))

    def test_switch_to_other_candidate_rewrites_height_before_confirming(self):
        self.autofix.AUTOFIX["auto"] = True
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.written(gid)
        with self.assertRaisesRegex(ValueError, "已显示全"):
            self.autofix.switch_group(gid, 1)
        self.autofix.switch_group(gid, 2)
        self.autofix.sync_groups()  # 问题项已解决，改图期间分组不能被清掉
        self.assertEqual(self.autofix.AUTOFIX["groups"][gid]["status"], "applying")
        self.assertEqual(self.autofix._next_task({}), ("apply", (str(self.wb_path), [gid])))
        group = self.written(gid)
        self.assertEqual((group["status"], group["chosen"], group["confirmed"]), ("applied", 2, False))
        self.assertEqual(self.real_heights, {5: 21.0})
        self.assertEqual(self.workflow.STATE["items"]["i1"]["row_height"], 21.0)

    def test_reject_restores_original_height_and_queues_later_review(self):
        self.autofix.AUTOFIX["auto"] = True
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.written(gid)
        self.autofix.reject_group(gid)
        with self.assertRaisesRegex(ValueError, "正在写入"):
            self.autofix.reject_group(gid)
        group = self.written(gid)
        self.assertEqual(self.real_heights, {5: 15.0})
        item = self.workflow.STATE["items"]["i1"]
        self.assertEqual((item["status"], item["later_review"], item["row_height"]), ("manual_review", True, 15.0))
        self.assertNotIn("autofix_unconfirmed", item)
        self.assertIn(str(self.wb_path), self.workflow.STATE["review_queue"])
        self.autofix.sync_groups()
        self.assertEqual(group["status"], "reverted")
        self.assertTrue(group["manual"])  # 不再做图
        self.autofix.confirm_groups([gid])  # 批量确认时把已退回的收起来
        self.assertTrue(group["dismissed"])

    def test_default_auto_sends_rows_it_cannot_fix_to_later_review(self):
        self.autofix.AUTOFIX["auto"] = True
        self.needed = 999.0
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.autofix._auto_defer(gid)
        item = self.workflow.STATE["items"]["i1"]
        self.assertEqual((item["status"], item.get("later_review")), ("manual_review", True))
        self.assertEqual(self.real_heights, {5: 15.0})

    def test_unconfirmed_rows_survive_restart_without_pictures(self):
        self.autofix.AUTOFIX["auto"] = True
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.written(gid)
        self.autofix.AUTOFIX["groups"].clear()  # 相当于重启：内存里的图片没了
        group = self.group()
        self.assertEqual((group["status"], group["confirmed"], group["base_height"], group["attempts"]),
                         ("applied", False, 15.0, []))
        self.autofix.reject_group(gid)
        self.written(gid)
        self.assertEqual(self.real_heights, {5: 15.0})

    def test_hidden_preview_waits_until_border_no_longer_crosses_text(self):
        self.workflow.STATE["items"]["i1"]["kind"] = "hidden"
        self.needed = 27.0
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        group = self.group()
        self.assertEqual([(a["height"], a["ok"]) for a in group["attempts"]],
                         [(15.0, False), (18.0, False), (21.0, False), (24.0, False),
                          (27.0, True), (30.0, True)])
        self.assertEqual(group["recommended"], 5)

    def test_missing_whole_lines_are_added_at_once(self):
        # 9 行的格子缺 2 行：3pt 一次加不完 8 次，先按缺的行数补上，剩下再按 3pt 加
        self.workflow.STATE["items"]["i1"]["kind"] = "hidden"

        def evaluate(*args):
            height = self.books[-1].heights[5]
            return height >= 42.0, "data:image/png;base64,", 24.0 if height == 15.0 else 0.0

        gid = self.group()["id"]
        with patch.object(self.autofix, "_evaluate", side_effect=evaluate):
            self.autofix._preview(self.session, gid)
        group = self.group()
        self.assertEqual([(a["height"], a["ok"]) for a in group["attempts"]],
                         [(15.0, False), (39.0, False), (42.0, True), (45.0, True)])
        self.assertEqual(group["recommended"], 3)

    def test_full_auto_does_not_queue_hidden_text_that_is_still_crossed(self):
        self.workflow.STATE["items"]["i1"]["kind"] = "hidden"
        self.needed = 999.0
        gid = self.group()["id"]
        self.autofix.set_auto(True)
        self.autofix._preview(self.session, gid)
        group = self.group()
        self.assertEqual((group["status"], group["recommended"], group["chosen"]), ("failed", None, None))
        self.assertTrue(all(not a["ok"] for a in group["attempts"]))
        self.assertIsNone(self.autofix._next_task({}))
        self.assertEqual(self.real_heights, {5: 15.0})

    def test_open_file_waits_before_writing(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.autofix.choose(gid, 3)
        lock = self.wb_path.with_name("~$book.xlsx")
        lock.write_bytes(b"")
        self.assertFalse(self.autofix._apply(self.session, str(self.wb_path), [gid]))
        self.assertEqual(self.real_heights, {5: 15.0})
        self.assertIn("关闭后自动写入", self.group()["message"])
        lock.unlink()
        self.assertTrue(self.autofix._apply(self.session, str(self.wb_path), [gid]))

    def test_background_task_or_manual_pauses_preview(self):
        gid = self.group()["id"]
        self.workflow.STATE["busy"] = True
        with self.assertRaises(self.autofix._Paused):
            self.autofix._preview(self.session, gid)
        self.assertEqual((self.group()["status"], self.group()["attempts"]), ("queued", []))
        self.workflow.STATE["busy"] = False
        self.workflow.STATE["items"]["i1"]["status"] = "manual_review"
        self.autofix.sync_groups()
        self.assertIsNone(self.autofix._next_task({}))

    def test_rescan_with_different_items_resets_group(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.workflow.STATE["items"]["i2"] = dict(self.workflow.STATE["items"]["i1"], id="i2", seq=2)
        group = self.group()
        self.assertEqual((group["status"], group["attempts"], sorted(group["item_ids"])), ("queued", [], ["i1", "i2"]))


    def test_recheck_still_clipped_after_manual_prepares_again(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.workflow.STATE["items"]["i1"]["status"] = "manual_review"
        self.assertTrue(self.group()["manual"])
        self.assertEqual(len(self.group()["attempts"]), 4)  # 转人工本身不丢图片
        self.workflow.STATE["items"]["i1"]["status"] = "pending"  # 复检后仍检出
        group = self.group()
        self.assertEqual((group["manual"], group["status"], group["attempts"]), (False, "queued", []))
        self.assertIn("复检", group["message"])


    def test_previews_only_while_fix_page_open(self):
        gid = self.group()["id"]
        self.autofix.AUTOFIX["active_until"] = 0.0  # 用户还在筛选
        self.assertIsNone(self.autofix._next_task({}))
        self.autofix.mark_active()
        self.assertEqual(self.autofix._next_task({}), ("preview", gid))
        self.autofix.AUTOFIX["active_until"] = 0.0
        self.autofix.set_auto(True)  # 全自动也要开着 ② 页面才做图
        self.assertIsNone(self.autofix._next_task({}))

    def test_failed_manual_recheck_reuses_one_fresh_image_without_preparing_again(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        item = self.workflow.STATE["items"]["i1"]
        item.update(status="manual_review", manual_waiting=False,
                    manual_recheck_token="check1", image="data:image/png;base64,YQ==", row_height=33.0)
        group = self.group()
        self.assertEqual((group["status"], group["recommended"]), ("failed", None))
        self.assertEqual(self.autofix.snapshot()["groups"][0]["attempts"], [{"height": 33.0, "ok": False}])
        self.assertEqual(self.autofix.attempt_image(gid, 0), b"a")
        self.assertNotIn("recheck_image", group)
        self.assertIsNone(self.autofix._next_task({}))
        self.autofix.sync_groups()
        self.assertEqual(len(self.group()["attempts"]), 1)
        # 重启后没有旧分组，也必须直接显示复检结果。
        self.autofix.AUTOFIX["groups"].clear()
        self.assertEqual((self.group()["status"], len(self.group()["attempts"])), ("failed", 1))

    def test_manual_open_waits_only_for_original_file_writes(self):
        self.autofix._IDLE.clear()  # 副本 Excel 尚未关闭
        self.autofix.wait_for_writes(timeout=0)
        self.autofix._WRITES_IDLE.clear()
        with self.assertRaisesRegex(ValueError, "保存"):
            self.autofix.wait_for_writes(timeout=0)

    def test_new_sheet_detection_invalidates_old_recommendation_and_queued_write(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.autofix.choose(gid, 3)
        self.workflow.STATE["items"]["i1"]["autofix_revision"] = "new-sheet-pdf"
        group = self.group()
        self.assertEqual((group["status"], group["attempts"], group["chosen"]), ("queued", [], None))
        self.assertEqual(self.autofix._next_task({}), ("preview", gid))

    def test_chosen_picture_is_still_written_after_leaving_fix_page(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.autofix.choose(gid, 3)
        self.autofix.AUTOFIX["active_until"] = 0.0
        self.assertEqual(self.autofix._next_task({}), ("apply", (str(self.wb_path), [gid])))

    def test_row_waits_for_every_unresolved_cell_to_be_screened(self):
        self.workflow.STATE["items"]["i2"] = dict(self.workflow.STATE["items"]["i1"], id="i2", triage_keep=False)
        self.autofix.sync_groups()
        self.assertEqual(self.autofix.AUTOFIX["groups"], {})
        self.workflow.apply_keep_decisions({"i2": True})
        self.assertEqual(self.group()["item_ids"], ["i1", "i2"])
        self.workflow.apply_keep_decisions({"i2": False})
        self.assertEqual(self.autofix.AUTOFIX["groups"], {})
        self.workflow.apply_skip_decisions({"i2": True})
        self.assertEqual(self.group()["item_ids"], ["i1"])

    def test_scan_and_autofix_overlap_only_for_completed_files(self):
        future = self.root / "future.xlsx"
        future.write_bytes(b"future workbook")
        self.workflow.STATE["items"]["future"] = dict(
            self.workflow.STATE["items"]["i1"], id="future", wb_path=str(future), wb_name=future.name,
        )
        self.workflow.STATE.update(busy=True, job={
            "kind": "scan", "status": "running", "completed": [str(self.wb_path)],
            "workbooks": [str(self.wb_path), str(future)], "options": {}, "pending_adjustment": None,
        })
        self.workflow.STATE["scan_progress"] = {"total": 2, "completed": 1, "status": "running"}
        gid = self.group()["id"]
        self.assertFalse(self.autofix._busy())
        self.autofix._preview(self.session, gid)
        self.autofix.choose(gid, 3)
        self.assertTrue(self.autofix._apply(self.session, str(self.wb_path), [gid]))
        snapshot = self.workflow.get_state_snapshot()
        self.assertEqual(list(snapshot["items"]), ["i1"])
        self.assertEqual(snapshot["scan_progress"]["completed"], 1)
        self.assertEqual(snapshot["items"]["i1"]["status"], "fixed")
        self.assertTrue(snapshot["busy"])
        self.assertEqual(self.workflow.STATE["items"]["future"]["status"], "pending")

    def test_uncompleted_file_cannot_be_written_or_decided_even_with_stale_group(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.autofix.choose(gid, 3)
        self.workflow.STATE.update(busy=True, job={"kind": "scan", "status": "running", "completed": []})
        self.assertFalse(self.autofix._apply(self.session, str(self.wb_path), [gid]))
        self.assertTrue(all(not book.saved for book in self.books))
        with self.assertRaisesRegex(ValueError, "尚未扫描完成"):
            self.workflow.apply_keep_decisions({"i1": False})
        with self.assertRaisesRegex(ValueError, "尚未扫描完成"):
            self.workflow.apply_skip_decisions({"i1": True})
        with self.assertRaises(ValueError):
            self.autofix.choose(gid, 3)

    def test_keep_decision_persists_and_rolls_back_if_database_write_fails(self):
        self.workflow.STATE["items"]["i1"]["triage_keep"] = False
        self.workflow.apply_keep_decisions({"i1": True})
        self.assertTrue(storage.load_snapshot("workflow")["items"]["i1"]["triage_keep"])
        with patch.object(self.workflow, "_persist_state", side_effect=OSError("disk full")):
            with self.assertRaisesRegex(OSError, "disk full"):
                self.workflow.apply_keep_decisions({"i1": False})
        self.assertTrue(self.workflow.STATE["items"]["i1"]["triage_keep"])

    def test_undo_keep_stops_preview_but_rejects_a_row_already_writing(self):
        gid = self.group()["id"]
        evaluate = self.autofix._evaluate

        def undo_after_export(*args):
            result = evaluate(*args)
            self.workflow.apply_keep_decisions({"i1": False})
            return result

        with patch.object(self.autofix, "_evaluate", side_effect=undo_after_export):
            with self.assertRaises(self.autofix._Paused):
                self.autofix._preview(self.session, gid)
        self.assertEqual(self.autofix.AUTOFIX["groups"], {})
        self.workflow.apply_keep_decisions({"i1": True})
        self.autofix._preview(self.session, gid)
        self.autofix.choose(gid, 3)
        with self.assertRaisesRegex(ValueError, "开始写入"):
            self.workflow.apply_keep_decisions({"i1": False})
        self.assertTrue(self.workflow.STATE["items"]["i1"]["triage_keep"])


class AutofixEvaluationTests(unittest.TestCase):
    def setUp(self):
        from core import autofix
        self.autofix = autofix
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.pdf = Path(self.temp.name) / "attempt.pdf"

    def evaluate(self, text, lines, *, display=None, border=None, cells=None, kind="hidden", item_text=None):
        from core.pdf_detector import fitz
        with fitz.open() as doc:
            page = doc.new_page()
            for x, y, line in lines:
                page.insert_text((x, y), line, fontsize=10, fontname="japan")
            if border:
                page.draw_line(border[0], border[1], width=0.5)
            doc.save(self.pdf)
        item = {"kind": kind, "text": item_text or text, "excel_row": 67,
                "cell_addr": "R67", "page_index": 0, "bbox": (60, 90, 150, 102)}
        return self.autofix._evaluate(
            self.pdf, [item], cells if cells is not None else [(67, 18, text)],
            lambda r, c: display if display is not None else text,
            lambda r, c: True,
        )[:2]

    def test_complete_hidden_text_crossed_by_border_is_not_success(self):
        ok, image = self.evaluate("ryouyouhi_houmon_name", [
            (60, 100, "ryouyouhi_houmon_n"), (60, 112, "ame"),
        ], border=((55, 109), (300, 109)))
        self.assertFalse(ok)
        self.assertTrue(image)

    def test_shortened_display_text_cannot_hide_missing_suffix(self):
        ok, _ = self.evaluate("文字列(100)", [(60, 100, "文字列")], display="文字列")
        self.assertFalse(ok)

    def test_hidden_problem_is_matched_by_cell_even_if_text_representation_changed(self):
        ok, _ = self.evaluate("文字列(100)", [(60, 100, "文字列")], item_text="文字列（100）")
        self.assertFalse(ok)

    def test_same_row_problem_in_other_column_does_not_replace_target(self):
        ok, _ = self.evaluate("Same", [(60, 100, "Same")], cells=[(67, 18, "Same"), (67, 19, "Same")])
        self.assertTrue(ok)

    def test_missing_target_evidence_is_not_success(self):
        ok, _ = self.evaluate("Target", [(60, 100, "Target")], cells=[])
        self.assertFalse(ok)

    def test_clip_problem_also_checks_hidden_suffix(self):
        ok, _ = self.evaluate("TargetMore", [(60, 100, "Target")], kind="clip", item_text="Target")
        self.assertFalse(ok)

    def test_complete_target_without_crossing_border_passes(self):
        ok, image = self.evaluate("文字列(100)", [(60, 100, "文字列(100)")])
        self.assertTrue(ok)
        self.assertTrue(image)

    def test_truncated_target_reports_how_much_height_is_missing(self):
        from core.pdf_detector import fitz
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((60, 100), "桁数④が1以上の場合、", fontsize=10, fontname="japan")
            page.insert_text((60, 112), "4", fontsize=10, fontname="japan")
            doc.save(self.pdf)
        text = "桁数④が1以上の場合、4\n桁数③が1以上の場合、3\n上記以外の場合、0"
        item = {"kind": "hidden", "text": text, "excel_row": 67, "cell_addr": "R67"}
        ok, _, need = self.autofix._evaluate(
            self.pdf, [item], [(67, 18, text)], lambda r, c: text, lambda r, c: True,
        )
        self.assertFalse(ok)
        self.assertAlmostEqual(need, 3 * 12.0, places=1)

    def test_border_crossing_other_column_does_not_fail_target(self):
        ok, _ = self.evaluate("Target", [(60, 100, "Target"), (250, 100, "Other")],
                              border=((245, 96), (400, 96)))
        self.assertTrue(ok)


if __name__ == "__main__":
    unittest.main()
