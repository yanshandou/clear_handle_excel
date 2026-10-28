import importlib
import os
import re
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import excel_ops, fix_images, storage


class FakeSheet:
    def __init__(self, test, book, name):
        self.test = test
        self.book = book
        self.Name = name
        self.CodeName = f"code_{name}"
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
        self.patch_value(fix_images, "IMAGES_DIR", self.root / "_report_images")
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
        self.excel.Workbooks.Open.side_effect = lambda path, **_: self.books[Path(path).stem]
        self.dispatch = Mock(return_value=self.excel)
        self.patch_value(workflow, "pythoncom", Mock())
        self.patch_value(workflow, "win32", Mock(DispatchEx=self.dispatch))
        self.patch_value(excel_ops, "win32", workflow.win32)  # 后台 Excel 由 excel_ops 启动
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

    def test_first_file_is_reviewable_and_queues_work_before_second_file_opens(self):
        from core import autofix
        autofix = importlib.reload(autofix)
        autofix._worker = Mock(is_alive=Mock(return_value=True))
        autofix.AUTOFIX["active_until"] = float("inf")
        first = self.add_book("first", ["S1"])
        self.add_book("second", ["S1"])
        self.clips[("first", "S1")] = ["one", "two"]
        self.books["first"].Close = Mock()
        observed = []

        def open_book(path, **_):
            if Path(path).stem == "second":
                self.books["first"].Close.assert_called_once_with(SaveChanges=False)
                state = self.workflow.get_state_snapshot()
                self.assertTrue(state["busy"])
                self.assertEqual(state["scan_progress"], {"total": 2, "completed": 1, "status": "running"})
                one, two = sorted(state["items"].values(), key=lambda it: it["text"])
                self.workflow.apply_keep_decisions({one["id"]: True})
                self.assertEqual(autofix.snapshot()["groups"], [])  # 同行还有一处没筛
                self.workflow.apply_skip_decisions({two["id"]: True})
                group = autofix.snapshot()["groups"][0]
                self.assertEqual(autofix._next_task({}), ("preview", group["id"]))
                self.assertEqual(group["wb_path"], str(first))
                observed.append(True)
            return self.books[Path(path).stem]

        self.excel.Workbooks.Open.side_effect = open_book
        job = self.scan()
        self.assertEqual(job["status"], "completed")
        self.assertEqual(observed, [True])
        self.assertTrue(self.items_by_text()["one"]["triage_keep"])
        self.assertEqual(self.items_by_text()["two"]["status"], "accepted")

    def test_failed_incremental_file_is_not_exposed_as_newly_scanned(self):
        first = self.add_book("first", ["S1"])
        second = self.add_book("second", ["S1"])
        self.clips[("first", "S1")] = ["one"]
        self.clips[("second", "S1")] = ["two"]
        self.scan()
        self.modify(second)
        self.fail_export["second"] = "S1"
        self.assertEqual(self.scan()["status"], "interrupted")
        state = self.workflow.get_state_snapshot()
        self.assertEqual({it["wb_path"] for it in state["items"].values()}, {str(first)})
        old = self.items_by_text()["two"]
        with self.assertRaisesRegex(ValueError, "尚未扫描完成"):
            self.workflow.apply_skip_decisions({old["id"]: True})
        current = self.items_by_text()["one"]
        self.workflow.apply_skip_decisions({current["id"]: True})
        self.assertEqual(self.workflow.STATE["job"]["status"], "interrupted")
        self.assertEqual(self.workflow.STATE["scan_progress"]["completed"], 1)

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

    def test_renamed_sheet_is_followed_by_code_name(self):
        """人工处理时改了表名：按 Excel 代码名认出新表，照常检测，不能当作问题都修好了"""
        path = self.add_book("book", ["S1", "S2"])
        self.clips[("book", "S1")] = ["k", "i"]
        self.scan()
        first = self.items_by_text()
        self.books["book"].Sheets[0].Name = "S1b"
        self.clips[("book", "S1b")] = ["k"]
        self.modify(path)
        self.scan()
        self.assertEqual(self.exports, [("book", "S1b")])
        items = self.items_by_text()
        self.assertEqual((items["k"]["id"], items["k"]["status"], items["k"]["sheet_name"]),
                         (first["k"]["id"], "pending", "S1b"))
        self.assertEqual((items["i"]["status"], items["i"]["sheet_name"]), ("fixed", "S1b"))

    def test_missing_sheet_keeps_its_problems_open(self):
        for case in ("deleted", "hidden", "no code name"):
            with self.subTest(case):
                self.workflow.STATE.update(items={}, pages={}, checked_files={}, folder="")
                path = self.add_book("book", ["S1", "S2"])
                self.clips[("book", "S1")] = ["k"]
                if case == "no code name":
                    self.books["book"].Sheets[0].CodeName = ""
                self.scan()
                sheet = self.books["book"].Sheets[0]
                if case == "deleted":
                    self.books["book"].Sheets.remove(sheet)
                elif case == "hidden":
                    sheet.Visible = 0
                else:
                    sheet.Name = "S1b"
                    self.clips[("book", "S1b")] = ["k"]
                self.modify(path)
                self.scan()
                self.assertNotIn(("book", "S1"), self.exports)
                item = self.items_by_text()["k"]
                self.assertEqual((item["status"], item["sheet_name"]), ("pending", "S1"))

    def test_full_check_counts_problems_on_a_vanished_sheet_as_solved(self):
        """整份检查已包含现在所有的表：表不在了，上面的旧问题算解决（改名认不出时问题按新表另出）"""
        path = self.add_book("book", ["S1", "S2"])
        self.clips[("book", "S1")] = ["k"]
        self.books["book"].Sheets[0].CodeName = ""
        self.scan()
        self.books["book"].Sheets[0].Name = "S1b"
        self.clips[("book", "S1b")] = ["k"]
        self.modify(path)
        previous = [dict(it) for it in self.workflow.STATE["items"].values()]
        entry = {"mode": "full", "sheets": None, "scanned": True}
        self.workflow._scan_one_workbook(
            self.excel, path, entry, previous, job=None, resumed=True, rescan_all=False,
            delete_backups=False, used_pdf_paths=set(), read_only=True)
        self.assertEqual(entry["missing"], ["S1"])
        by_sheet = {(it["sheet_name"], it["status"]) for it in self.workflow.STATE["items"].values()}
        self.assertEqual(by_sheet, {("S1", "fixed"), ("S1b", "pending")})

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

    def test_previous_detector_version_rechecks_unchanged_workbooks(self):
        self.add_book("book", ["S1", "S2"])
        self.scan()
        for version in self.workflow.STATE["checked_files"].values():
            version["detector"] = self.workflow.DETECTOR_VERSION - 1

        self.scan()
        self.assertEqual(self.exports, [("book", "S1"), ("book", "S2")])
        self.scan()
        self.assertEqual(self.exports, [])

    def test_visible_spaced_paragraphs_do_not_create_hidden_items(self):
        path = self.add_book("book", ["S1"])
        lines = ["●登録前チェック", "BIIryouChk::chkErrorRyouyouhiMeisai()",
                 "●請求エラーチェック", "BIIryouKokuhoError::getKokuhoKindEntry()"]
        text = "\n".join(lines[:2]) + "\n\n" + "\n".join(lines[2:])
        with self.workflow.fitz.open() as doc:
            page = doc.new_page()
            for y, line in zip((100, 112, 148, 160), lines):
                page.insert_text((60, y), line, fontsize=10, fontname="japan")
            doc.save(self.root / "paragraphs.pdf")
        self.blank_pdf = self.root / "paragraphs.pdf"
        self.patch_value(self.workflow, "extract_sheet_snapshot", Mock(return_value={
            "start_r": 47, "start_c": 33, "rows_vals": [(text,)],
        }))

        job = self.scan()
        self.assertEqual(job["status"], "completed")
        self.assertEqual(self.workflow.STATE["items"], {})
        report = self.reporter.get_execution_report_snapshot()["files"][path.name]
        self.assertEqual((report["status"], report["clip_unfixed"]), ("done", 0))

    def test_halfwidth_merged_table_headers_are_not_missing(self):
        path = self.add_book("book", ["S1"])
        with self.workflow.fitz.open() as doc:
            page = doc.new_page()
            for y in (100, 150):
                page.draw_rect(self.workflow.fitz.Rect(50, y - 20, 150, y + 20))
                page.insert_text((70, y + 5), "ﾃｰﾌﾞﾙ", fontsize=10, fontname="japan")
            doc.save(self.root / "headers.pdf")
        self.blank_pdf = self.root / "headers.pdf"
        # 两个两行合并表头，各自仅在左上角单元格有值。
        rows = [("ﾃｰﾌﾞﾙ",)] + [(None,)] * 9 + [("ﾃｰﾌﾞﾙ",), (None,)]
        self.patch_value(self.workflow, "extract_sheet_snapshot", Mock(return_value={
            "start_r": 46, "start_c": 26, "rows_vals": rows,
        }))

        job = self.scan()
        self.assertEqual(job["status"], "completed")
        self.assertEqual(self.workflow.STATE["items"], {})
        report = self.reporter.get_execution_report_snapshot()["files"][path.name]
        self.assertEqual((report["status"], report["clip_unfixed"]), ("done", 0))

    def test_strikethrough_does_not_create_a_problem_card(self):
        path = self.add_book("book", ["S1"])
        text = "DbExCmok06Mapper"
        with self.workflow.fitz.open() as doc:
            page = doc.new_page()
            page.draw_rect(self.workflow.fitz.Rect(50, 75, 250, 120))
            page.insert_text((60, 100), text, fontsize=12)
            span = page.get_text("dict")["blocks"][0]["lines"][0]["spans"][0]
            x0, _, x1, _ = span["bbox"]
            page.draw_line((x0, 96.4), (x1, 96.4), width=0.5)
            doc.save(self.root / "strikethrough.pdf")
        self.blank_pdf = self.root / "strikethrough.pdf"
        self.patch_value(self.workflow, "find_clipped_spans_on_page",
                         self.workflow.pdf_detector.find_clipped_spans_on_page)
        self.patch_value(self.workflow, "extract_sheet_snapshot", Mock(return_value={
            "start_r": 23, "start_c": 8, "rows_vals": [(text,)],
        }))

        self.assertEqual(self.scan()["status"], "completed")
        self.assertEqual(self.workflow.STATE["items"], {})
        report = self.reporter.get_execution_report_snapshot()["files"][path.name]
        self.assertEqual((report["status"], report["clip_unfixed"]), ("done", 0))

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
        sheet.Rows = Mock(return_value=Mock(RowHeight=13.5, Hidden=False))
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

    def test_clip_among_same_texts_is_confirmed_by_font_color(self):
        self.add_book("book", ["S1"])
        sheet = self.books["book"].Sheets[0]
        colors = {}  # (行, 列) -> Excel 字体颜色（BGR）
        rows = {2: "非表示", 3: "非表示", 4: "非表示"}

        class Font:
            def __init__(self, rc):
                self.rc = rc

            @property
            def Color(self):
                return colors.get(self.rc, 0)

            @Color.setter
            def Color(self, value):
                colors[self.rc] = value

        sheet.Cells = lambda r, c: Mock(Font=Font((r, c)))
        sheet.Rows = Mock(return_value=Mock(RowHeight=15.0, Hidden=False))
        workflow = self.workflow

        def export(**kwargs):
            # 按格子当前的字体颜色画出三行同样的文字
            with workflow.fitz.open() as doc:
                page = doc.new_page()
                for row, text in rows.items():
                    bgr = colors.get((row, 1), 0)
                    rgb = (bgr & 0xFF, bgr >> 8 & 0xFF, bgr >> 16 & 0xFF)
                    page.insert_text((60, 40 * row), text, fontsize=10, fontname="japan",
                                     color=tuple(v / 255 for v in rgb))
                doc.save(kwargs["Filename"])
        sheet.ExportAsFixedFormat = export
        self.patch_value(workflow, "extract_sheet_snapshot", Mock(return_value={
            "start_r": 2, "start_c": 1, "rows_vals": [(t,) for t in rows.values()],
        }))
        # 推算的行号故意给错，确认以颜色为准
        self.patch_value(workflow, "locate_cell", Mock(return_value=(2, 1)))

        def clips(page):
            spans = [
                span for block in page.get_text("dict")["blocks"]
                for line in block.get("lines", []) for span in line["spans"]
            ]
            span = next(sp for sp in spans if abs(sp["origin"][1] - 120) < 1)  # 第 3 行
            return [{"text": span["text"], "bbox": tuple(span["bbox"]), "line_y": 0}]
        self.patch_value(workflow, "find_clipped_spans_on_page", clips)
        self.scan()

        clip_items = [it for it in workflow.STATE["items"].values() if it["kind"] == "clip"]
        self.assertEqual(len(clip_items), 1)
        self.assertEqual((clip_items[0]["cell_addr"], clip_items[0]["location_confirmed"]), ("A3", True))
        self.assertEqual(colors, {(2, 1): 0, (3, 1): 0, (4, 1): 0})  # 颜色已还原
        self.assertFalse(any(p.name.startswith("~定位_") for p in self.folder.iterdir()))

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
