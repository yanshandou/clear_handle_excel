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

        self.wb_path = self.root / "book.xlsx"
        self.wb_path.write_bytes(b"workbook")
        self.needed = 21.0
        self.real_heights = {5: 15.0}
        self.books = []

        def open_book(path):
            heights = self.real_heights if Path(path) == self.wb_path.resolve() else dict(self.real_heights)
            book = FakeBook(self, Path(path), heights)
            self.books.append(book)
            return book
        excel = Mock()
        excel.Workbooks.Open.side_effect = open_book
        for name, value in {
            "win32": Mock(DispatchEx=Mock(return_value=excel)),
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
            "text": "Target", "bbox": bbox, "occurrence": None, "status": "pending", "iterations": 0,
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

    def test_chosen_picture_is_still_written_after_leaving_fix_page(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.autofix.choose(gid, 3)
        self.autofix.AUTOFIX["active_until"] = 0.0
        self.assertEqual(self.autofix._next_task({}), ("apply", (str(self.wb_path), [gid])))


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
        )

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

    def test_border_crossing_other_column_does_not_fail_target(self):
        ok, _ = self.evaluate("Target", [(60, 100, "Target"), (250, 100, "Other")],
                              border=((245, 96), (400, 96)))
        self.assertTrue(ok)


if __name__ == "__main__":
    unittest.main()
