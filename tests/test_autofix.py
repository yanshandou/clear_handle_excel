import importlib
import tempfile
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
            "extract_sheet_snapshot": Mock(return_value={}),
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

    def test_full_auto_applies_recommendation(self):
        gid = self.group()["id"]
        self.autofix.set_auto(True)
        self.autofix._preview(self.session, gid)
        group = self.group()
        self.assertEqual((group["status"], group["chosen"]), ("applying", 3))
        self.assertTrue(self.autofix._apply(self.session, str(self.wb_path), [gid]))
        self.assertEqual(self.real_heights, {5: 24.0})

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


    def test_previews_only_while_fix_page_open_or_full_auto(self):
        gid = self.group()["id"]
        self.autofix.AUTOFIX["active_until"] = 0.0  # 用户还在筛选
        self.assertIsNone(self.autofix._next_task({}))

        self.autofix.mark_active()
        self.assertEqual(self.autofix._next_task({}), ("preview", gid))
        self.autofix.AUTOFIX["active_until"] = 0.0
        self.autofix.set_auto(True)
        self.assertEqual(self.autofix._next_task({}), ("preview", gid))

    def test_chosen_picture_is_still_written_after_leaving_fix_page(self):
        gid = self.group()["id"]
        self.autofix._preview(self.session, gid)
        self.autofix.choose(gid, 3)
        self.autofix.AUTOFIX["active_until"] = 0.0
        self.assertEqual(self.autofix._next_task({}), ("apply", (str(self.wb_path), [gid])))


if __name__ == "__main__":
    unittest.main()
