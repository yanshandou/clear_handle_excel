import importlib
import os
import re
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import storage


class FakeSheet:
    def __init__(self, test, book, name):
        self.test = test
        self.book = book
        self.Name = name
        self.Visible = -1

    def ExportAsFixedFormat(self, **kwargs):
        if self.test.fail_export.get(self.book) == self.Name:
            raise RuntimeError("导出失败")
        self.test.exports.append((self.book, self.Name))
        shutil.copyfile(self.test.blank_pdf, kwargs["Filename"])


class FakeWorkbook:
    def __init__(self, test, name, sheet_names):
        self.Sheets = [FakeSheet(test, name, sheet) for sheet in sheet_names]

    def Save(self):
        raise AssertionError("扫描不应保存工作簿")

    def Close(self, SaveChanges):
        pass


class IncrementalScanTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.folder = self.root / "excel"
        self.folder.mkdir()
        self.patch_value(storage, "DATABASE_PATH", self.root / "test.sqlite3")
        from core import reporter, workflow

        self.reporter = importlib.reload(reporter)
        self.workflow = importlib.reload(workflow)
        self.patch_value(reporter, "RECORD_PATH", self.root / "record.md")
        self.patch_value(reporter, "REPORT_STATE_PATH", self.root / "legacy.json")

        self.blank_pdf = self.root / "blank.pdf"
        with workflow.fitz.open() as doc:
            doc.new_page()
            doc.save(self.blank_pdf)
        self.books = {}
        self.clips = {}  # (工作簿名, 工作表名) -> 检出的文字
        self.exports = []
        self.fail_export = {}
        self.excel = Mock()
        self.excel.Workbooks.Open.side_effect = lambda path: self.books[Path(path).stem]
        self.dispatch = Mock(return_value=self.excel)
        self.patch_value(workflow, "pythoncom", Mock())
        self.patch_value(workflow, "win32", Mock(DispatchEx=self.dispatch))
        self.patch_value(workflow, "extract_sheet_snapshot", Mock(return_value={}))
        self.patch_value(workflow, "match_excel_row", Mock(return_value=(2, "A2", 15.0)))
        self.patch_value(workflow, "find_clipped_spans_on_page", self.fake_clips)

    def patch_value(self, target, name, value):
        patcher = patch.object(target, name, value)
        patcher.start()
        self.addCleanup(patcher.stop)

    def fake_clips(self, page):
        # 同一秒内两次扫描会共用 PDF 文件夹，导出文件名会带「_2」后缀
        book, sheet = re.sub(r"_\d+$", "", Path(page.parent.name).stem).split("_", 1)
        return [
            {"text": text, "bbox": (60, 100 + 40 * i, 100, 115 + 40 * i), "line_y": 110 + 40 * i}
            for i, text in enumerate(self.clips.get((book, sheet), []))
        ]

    def add_book(self, name, sheets):
        path = self.folder / f"{name}.xlsx"
        path.write_bytes(f"workbook {name}".encode())
        self.books[name] = FakeWorkbook(self, name, sheets)
        return path

    def modify(self, path):
        stat = path.stat()
        os.utime(path, ns=(stat.st_atime_ns, stat.st_mtime_ns + 2_000_000_000))

    def scan(self, full=False):
        self.exports.clear()
        self.workflow.STATE["busy"] = False
        self.assertTrue(self.workflow.queue_job("scan", {
            "folder": str(self.folder), "delete_backups": False, "full": full,
        }))
        self.workflow.run_pending_workflow()
        return self.workflow.STATE["job"]

    def items_by_text(self):
        return {item["text"]: item for item in self.workflow.STATE["items"].values()}

    def test_unchanged_files_are_not_opened_again(self):
        self.add_book("book", ["S1"])
        self.clips[("book", "S1")] = ["k"]
        self.scan()
        before = self.items_by_text()["k"]["id"]
        self.dispatch.reset_mock()

        job = self.scan()
        self.assertEqual(job["status"], "completed")
        self.dispatch.assert_not_called()
        self.assertEqual(self.exports, [])
        self.assertEqual(self.items_by_text()["k"]["id"], before)
        report = self.reporter.get_execution_report_snapshot()["files"]["book.xlsx"]
        self.assertEqual((report["status"], report["clip_unfixed"]), ("needs_fix", 1))
        self.assertIn("沿用", report["current_action"])

    def test_modified_problem_file_only_rechecks_problem_sheets(self):
        path = self.add_book("book", ["S1", "S2"])
        self.clips[("book", "S1")] = ["k", "i"]
        self.clips[("book", "S2")] = ["x"]
        self.scan()
        first = self.items_by_text()
        self.workflow.apply_skip_decisions({first["x"]["id"]: True})  # S2 只剩已跳过的问题

        # 人工修改后：S1 的 i 已改好，k 仍压线，又挤出一处 new
        self.clips[("book", "S1")] = ["k", "new"]
        self.modify(path)
        job = self.scan()

        self.assertEqual(job["status"], "completed")
        self.assertEqual(self.exports, [("book", "S1")])
        items = self.items_by_text()
        self.assertEqual((items["k"]["id"], items["k"]["status"]), (first["k"]["id"], "pending"))
        self.assertFalse(items["k"]["is_new"])
        self.assertEqual((items["i"]["id"], items["i"]["status"]), (first["i"]["id"], "fixed"))
        self.assertTrue(items["new"]["is_new"])
        self.assertEqual(items["x"]["status"], "accepted")
        report = self.reporter.get_execution_report_snapshot()["files"]["book.xlsx"]
        self.assertEqual(
            (report["clip_unfixed"], report["clip_fixed"], report["clip_accepted"]), (2, 1, 1),
        )
        self.assertIn("本次确认已修复 1 处", report["current_action"])

    def test_modified_file_without_open_problems_is_fully_rechecked(self):
        path = self.add_book("book", ["S1", "S2"])
        self.scan()
        self.assertEqual(self.workflow.STATE["items"], {})

        self.clips[("book", "S2")] = ["late"]
        self.modify(path)
        self.scan()
        self.assertEqual(sorted(self.exports), [("book", "S1"), ("book", "S2")])
        self.assertTrue(self.items_by_text()["late"]["is_new"])

    def test_files_checked_by_older_rules_are_fully_rechecked(self):
        self.add_book("book", ["S1", "S2"])
        self.clips[("book", "S1")] = ["k"]
        self.scan()
        for version in self.workflow.STATE["checked_files"].values():
            del version["detector"]  # 加入隐藏截断检查之前的扫描记录

        self.scan()
        self.assertEqual(self.exports, [("book", "S1"), ("book", "S2")])
        self.assertFalse(self.items_by_text()["k"]["is_new"])

        self.scan()
        self.assertEqual(self.exports, [])

    def test_hidden_wrapped_text_is_reported_with_cell(self):
        path = self.add_book("book", ["S1"])
        with self.workflow.fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((60, 100), "DbExKhplg0", fontsize=10)
            doc.save(self.root / "hidden.pdf")
        self.blank_pdf = self.root / "hidden.pdf"
        self.patch_value(self.workflow, "extract_sheet_snapshot", Mock(return_value={
            "start_r": 23, "start_c": 8, "rows_vals": [("DbExKhplg0000Mapper",)],
        }))
        sheet = self.books["book"].Sheets[0]
        sheet.Rows = Mock(return_value=Mock(RowHeight=13.5))
        self.scan()

        (item,) = self.workflow.STATE["items"].values()
        self.assertEqual(item["kind"], "hidden")
        self.assertEqual((item["cell_addr"], item["excel_row"], item["row_height"]), ("H23", 23, 13.5))
        self.assertEqual((item["text"], item["visible_text"]), ("DbExKhplg0000Mapper", "DbExKhplg0"))
        self.assertEqual(item["page_index"], 0)
        self.assertIn(item["page_id"], self.workflow.STATE["pages"])
        report = self.reporter.get_execution_report_snapshot()["files"][path.name]
        self.assertEqual(report["clip_unfixed"], 1)

        # 改好后重新检查不再检出，记为已修复
        with self.workflow.fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((60, 100), "DbExKhplg0000Mapper", fontsize=10)
            doc.save(self.root / "hidden.pdf")
        self.modify(path)
        self.scan()
        (item,) = self.workflow.STATE["items"].values()
        self.assertEqual(item["status"], "fixed")

    def test_new_file_is_fully_scanned_and_not_marked_new(self):
        self.add_book("book", ["S1"])
        self.scan()
        self.add_book("other", ["A", "B"])
        self.clips[("other", "B")] = ["k"]
        self.scan()
        self.assertEqual(sorted(self.exports), [("other", "A"), ("other", "B")])
        self.assertFalse(self.items_by_text()["k"]["is_new"])

    def test_full_rescan_rechecks_everything_and_keeps_skips(self):
        self.add_book("book", ["S1"])
        self.clips[("book", "S1")] = ["k", "i"]
        self.scan()
        first = self.items_by_text()
        self.workflow.apply_skip_decisions({first["i"]["id"]: True})

        self.scan(full=True)
        self.assertEqual(self.exports, [("book", "S1")])
        items = self.items_by_text()
        self.assertEqual(items["i"]["status"], "accepted")
        self.assertEqual({item["id"] for item in items.values()}, {first["k"]["id"], first["i"]["id"]})
        self.assertFalse(any(item["is_new"] for item in items.values()))

    def test_failed_file_keeps_previous_records_and_can_resume(self):
        path = self.add_book("book", ["S1"])
        self.clips[("book", "S1")] = ["k"]
        self.scan()
        first_id = self.items_by_text()["k"]["id"]

        self.modify(path)
        self.fail_export["book"] = "S1"
        job = self.scan()
        self.assertEqual(job["status"], "interrupted")
        self.assertEqual(self.items_by_text()["k"]["id"], first_id)
        self.assertEqual(self.items_by_text()["k"]["status"], "pending")

        self.fail_export.clear()
        self.clips[("book", "S1")] = []
        self.workflow.STATE["busy"] = False
        self.assertTrue(self.workflow.queue_job("resume"))
        self.workflow.run_pending_workflow()
        self.assertEqual(self.workflow.STATE["job"]["status"], "completed")
        self.assertEqual(self.items_by_text()["k"]["status"], "fixed")


if __name__ == "__main__":
    unittest.main()
