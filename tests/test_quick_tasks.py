import copy
import http.client
import importlib
import io
import json
import shutil
import tempfile
import threading
import unittest
from contextlib import contextmanager
from http.server import ThreadingHTTPServer
from pathlib import Path
from unittest.mock import Mock, patch

from core import storage


class FakeSheet:
    def __init__(self, book, name, visible=-1):
        self.book = book
        self.Name = name
        self.Visible = visible
        self.ExportAsFixedFormat = Mock(side_effect=self._export)

    def _export(self, **kwargs):
        if self.book.fail_export:
            raise RuntimeError("导出失败")
        Path(kwargs["Filename"]).write_bytes(f"pdf of {self.book.name}".encode())

    def Delete(self):
        self.book.events.append(("delete", self.Name))
        self.book.sheets.remove(self)


class FakeSheets:
    def __init__(self, book):
        self.book = book

    def __iter__(self):
        return iter(list(self.book.sheets))

    def __call__(self, name):
        return next(sheet for sheet in self.book.sheets if sheet.Name == name)


class FakeWorkbook:
    def __init__(self, name, sheet_names, hidden, events):
        self.name = name
        self.events = events
        self.sheets = [FakeSheet(self, n, 0 if n in hidden else -1) for n in sheet_names]
        self.Sheets = FakeSheets(self)
        self.saved = []
        self.fail_export = False

    def Save(self):
        self.events.append(("save", self.name))
        self.saved.append([sheet.Name for sheet in self.sheets])

    def Close(self, SaveChanges):
        self.events.append(("close", self.name))


class FakeExcel:
    def __init__(self, books):
        self.books = books
        self.opened = []
        self.Workbooks = Mock()
        self.Workbooks.Open.side_effect = self.open
        self.Quit = Mock()

    def open(self, path, ReadOnly=False):
        name = Path(path).name
        self.opened.append((name, ReadOnly))
        return self.books[name]


class QuickTaskTests(unittest.TestCase):
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
        self.events = []
        self.books = {}
        self.excel = FakeExcel(self.books)
        self.dispatch = Mock(return_value=self.excel)
        self.normalize = Mock()
        self.patch_value(workflow, "pythoncom", Mock())
        self.patch_value(workflow, "win32", Mock(DispatchEx=self.dispatch))
        self.patch_value(workflow, "normalize_workbook_view", self.normalize)
        self.patch_value(workflow, "scan_workbook_sheets", self.read_sheet_names)

        real_copy2 = shutil.copy2

        def recording_copy2(src, dst):
            self.events.append(("backup", Path(src).name))
            return real_copy2(src, dst)

        self.patch_value(workflow.shutil, "copy2", recording_copy2)

    def patch_value(self, target, name, value):
        patcher = patch.object(target, name, value)
        patcher.start()
        self.addCleanup(patcher.stop)

    def add_book(self, name, sheets, hidden=()):
        path = self.folder / name
        path.write_bytes(f"original {name}".encode())
        self.books[name] = FakeWorkbook(name, sheets, hidden, self.events)
        return path

    def read_sheet_names(self, path):
        """模拟 openpyxl 只读读取工作表名"""
        book = self.books.get(Path(path).name)
        if book is None:
            raise ValueError("文件已损坏")
        names = [sheet.Name for sheet in book.sheets]
        return names, [name for name in names if self.workflow.is_backup_sheet(name)]

    def run_task(self, kind):
        self.assertTrue(self.workflow.queue_job(kind, {"folder": str(self.folder)}))
        self.workflow.run_pending_workflow()
        return self.workflow.STATE["job"]

    def report_files(self):
        return self.reporter.get_execution_report_snapshot()["files"]

    def restart(self):
        loaded = storage.load_snapshot("workflow")
        self.workflow.STATE.clear()
        self.workflow.STATE.update(copy.deepcopy(loaded))
        self.workflow.STATE["busy"] = False

    def test_clean_sheets_backs_up_before_deleting_only_dated_sheets(self):
        self.add_book("a.xlsx", ["機能設計", "機能設計_20250315_bak", "Sheet1_20261020"])
        self.add_book("b.xlsx", ["Sheet1"])
        job = self.run_task("clean_sheets")

        self.assertEqual(job["status"], "completed")
        self.assertEqual(self.books["a.xlsx"].saved, [["機能設計"]])
        self.assertEqual(
            self.events[:4],
            [("backup", "a.xlsx"), ("delete", "機能設計_20250315_bak"),
             ("delete", "Sheet1_20261020"), ("save", "a.xlsx")],
        )
        backup_folder = Path(job["options"]["backup_folder"])
        self.assertEqual((backup_folder / "a.xlsx").read_bytes(), b"original a.xlsx")
        self.assertFalse((backup_folder / "b.xlsx").exists())
        # 预检没有备份表的文件不会用 Excel 打开
        self.assertEqual(self.excel.opened, [("a.xlsx", False)])
        files = self.report_files()
        self.assertEqual(files["a.xlsx"]["deleted_backups"], ["機能設計_20250315_bak", "Sheet1_20261020"])
        self.assertEqual(files["b.xlsx"]["status"], "done")

    def test_clean_sheets_never_deletes_every_sheet(self):
        self.add_book("dated.xlsx", ["S_20250101", "S_20250102"])
        job = self.run_task("clean_sheets")

        self.assertEqual(job["status"], "completed")
        self.assertEqual(self.books["dated.xlsx"].saved, [])
        self.assertEqual(len(self.books["dated.xlsx"].sheets), 2)
        self.assertIn("跳过", self.report_files()["dated.xlsx"]["current_action"])

    def test_clean_sheets_skips_excel_entirely_when_no_backups(self):
        self.add_book("a.xlsx", ["Sheet1"])
        job = self.run_task("clean_sheets")

        self.assertEqual(job["status"], "completed")
        self.dispatch.assert_not_called()

    def test_export_pdf_is_read_only_and_skips_hidden_history_and_backup_sheets(self):
        self.reporter.init_execution_batch(
            "old", self.folder, [], backup_folder=self.root / "old_backup",
        )
        self.add_book("spec.xlsx", ["表紙", "変更履歴", "機能_20250101", "隠し"], hidden=["隠し"])
        job = self.run_task("export_pdf")

        self.assertEqual(job["status"], "completed")
        self.assertEqual(self.excel.opened, [("spec.xlsx", True)])
        book = self.books["spec.xlsx"]
        self.assertEqual(book.saved, [])
        self.assertEqual([s.Name for s in book.sheets if s.ExportAsFixedFormat.called], ["表紙"])
        pdf_folder = Path(job["options"]["pdf_folder"])
        self.assertEqual([p.name for p in pdf_folder.iterdir()], ["spec_表紙.pdf"])
        self.assertNotIn("backup_folder", job["options"])
        report = self.reporter.get_execution_report_snapshot()
        self.assertEqual(report["files"]["spec.xlsx"]["exported_pdfs"], 1)
        self.assertEqual(report["pdf_folder"], str(pdf_folder))
        # 新批次不沿用上一批的备份目录
        self.assertEqual(report["backup_folder"], "")

    def test_export_resume_retries_failed_workbook_without_overwriting_finished_pdf(self):
        self.add_book("a.xlsm", ["Sheet1"])
        self.add_book("a.xlsx", ["Sheet1"])
        self.books["a.xlsx"].fail_export = True
        job = self.run_task("export_pdf")

        self.assertEqual(job["status"], "interrupted")
        self.assertTrue(self.workflow.get_overview_snapshot()["can_resume"])
        self.assertEqual(self.report_files()["a.xlsx"]["status"], "error")

        self.restart()
        self.books["a.xlsx"].fail_export = False
        self.excel.opened.clear()
        self.assertTrue(self.workflow.queue_job("resume"))
        self.workflow.run_pending_workflow()

        job = self.workflow.STATE["job"]
        self.assertEqual(job["status"], "completed")
        self.assertEqual(self.excel.opened, [("a.xlsx", True)])
        pdf_folder = Path(job["options"]["pdf_folder"])
        self.assertEqual((pdf_folder / "a_Sheet1.pdf").read_bytes(), b"pdf of a.xlsm")
        self.assertEqual((pdf_folder / "a_Sheet1_2.pdf").read_bytes(), b"pdf of a.xlsx")

    def test_normalize_view_backs_up_then_saves_each_workbook(self):
        self.add_book("a.xlsx", ["Sheet1", "Sheet2"])
        job = self.run_task("normalize_view")

        self.assertEqual(job["status"], "completed")
        self.assertIs(self.normalize.call_args.args[1], self.books["a.xlsx"])
        self.assertEqual(self.events[0], ("backup", "a.xlsx"))
        self.assertEqual(len(self.books["a.xlsx"].saved), 1)
        self.assertTrue((Path(job["options"]["backup_folder"]) / "a.xlsx").exists())
        self.assertTrue(self.report_files()["a.xlsx"]["normalized_a1"])

    def test_quick_task_keeps_workbench_scan_results_and_folders(self):
        self.add_book("a.xlsx", ["Sheet1"])
        self.workflow.STATE.update(folder="D:/workbench", backup_folder="D:/workbench/_backup_x")
        self.workflow.STATE["items"]["item_1"] = {
            "id": "item_1", "wb_path": "D:/workbench/x.xlsx", "status": "pending",
        }
        self.run_task("export_pdf")

        self.assertEqual(self.workflow.STATE["folder"], "D:/workbench")
        self.assertEqual(self.workflow.STATE["backup_folder"], "D:/workbench/_backup_x")
        self.assertIn("item_1", self.workflow.STATE["items"])
        overview = self.workflow.get_overview_snapshot()
        self.assertEqual(overview["clips"], {"total": 1, "unresolved": 1})
        self.assertEqual(overview["job"]["kind"], "export_pdf")
        self.assertNotIn("pages", overview)

    def test_non_utf8_console_does_not_interrupt_background_task(self):
        # 管道输出时 stdout 使用系统 ANSI 代码页（如 cp932），无法编码 emoji 和部分中文
        self.add_book("機能設計書.xlsx", ["表紙"])
        console = io.TextIOWrapper(io.BytesIO(), encoding="cp932")
        with patch("sys.stdout", console):
            job = self.run_task("export_pdf")
            console.flush()
            output = console.buffer.getvalue()

        self.assertEqual(job["status"], "completed")
        self.assertTrue(any(line.endswith("✅ 批量导出 PDF完成：共导出 1 份 PDF") for line in self.workflow.STATE["logs"]))
        self.assertIn(b"\\u2705", output)

    def test_preview_lists_backups_without_opening_excel(self):
        self.add_book("a.xlsx", ["Main", "Main_20250101"])
        self.add_book("b.xlsx", ["X_20250101"])
        self.add_book("c.xlsx", ["Main"])
        (self.folder / "d.xlsx").write_bytes(b"broken")
        (self.folder / "~$a.xlsx").write_bytes(b"lock file")
        data = self.workflow.preview_backup_sheets(str(self.folder))

        files = {f["name"]: f for f in data["files"]}
        self.assertEqual(sorted(files), ["a.xlsx", "b.xlsx", "c.xlsx", "d.xlsx"])
        self.assertEqual(files["a.xlsx"]["backups"], ["Main_20250101"])
        self.assertEqual(files["b.xlsx"]["backups"], [])
        self.assertIn("全部工作表", files["b.xlsx"]["note"])
        self.assertEqual(files["c.xlsx"]["backups"], [])
        self.assertTrue(files["d.xlsx"]["note"].startswith("无法读取"))
        self.dispatch.assert_not_called()

    @contextmanager
    def serve_api(self):
        import web_clip_fixer

        importlib.reload(web_clip_fixer)
        with patch.object(web_clip_fixer, "run_pending_workflow"):
            server = ThreadingHTTPServer(("127.0.0.1", 0), web_clip_fixer.WebHandler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                yield server.server_port
            finally:
                server.shutdown()
                server.server_close()
                thread.join()

    def request(self, port, method, path, payload=None):
        connection = http.client.HTTPConnection("127.0.0.1", port, timeout=5)
        try:
            body = json.dumps(payload) if payload is not None else None
            connection.request(method, path, body, {"Content-Type": "application/json"})
            response = connection.getresponse()
            return response.status, response.read().decode("utf-8")
        finally:
            connection.close()

    def test_home_page_and_workbench_routes(self):
        with self.serve_api() as port:
            status, body = self.request(port, "GET", "/")
            self.assertEqual(status, 200)
            self.assertIn("常用功能", body)
            for path in ("/workbench", "/index.html"):
                status, body = self.request(port, "GET", path)
                self.assertEqual(status, 200)
                self.assertIn("STEP 01", body)
            status, body = self.request(port, "GET", "/api/overview")
            self.assertEqual(status, 200)
            overview = json.loads(body)
            self.assertFalse(overview["busy"])
            self.assertNotIn("items", overview)

    def test_api_quick_task_validates_kind_and_folder(self):
        self.add_book("a.xlsx", ["Sheet1"])
        empty = self.root / "empty"
        empty.mkdir()
        with self.serve_api() as port:
            for payload in (
                {"kind": "delete_everything", "folder": str(self.folder)},
                {"kind": ["export_pdf"], "folder": str(self.folder)},
                {"kind": "export_pdf"},
                {"kind": "export_pdf", "folder": str(self.root / "missing")},
                {"kind": "export_pdf", "folder": str(empty)},
            ):
                with self.subTest(payload=payload):
                    status, body = self.request(port, "POST", "/api/quick_task", payload)
                    self.assertEqual((status, json.loads(body)["ok"]), (400, False))
            self.assertIsNone(self.workflow.STATE["job"])

            status, body = self.request(port, "POST", "/api/quick_task", {
                "kind": "export_pdf", "folder": f'"{self.folder}"',
            })
            self.assertEqual((status, json.loads(body)["ok"]), (200, True))
            job = storage.load_snapshot("workflow")["job"]
            self.assertEqual(job["kind"], "export_pdf")
            self.assertEqual(job["options"]["folder"], str(self.folder.resolve()))

            status, body = self.request(port, "POST", "/api/quick_task", {
                "kind": "clean_sheets", "folder": str(self.folder),
            })
            self.assertEqual((status, json.loads(body)["ok"]), (409, False))

    def test_api_clean_preview_is_read_only(self):
        self.add_book("a.xlsx", ["Main", "Main_20250101"])
        with self.serve_api() as port:
            status, body = self.request(port, "POST", "/api/clean/preview", {"folder": str(self.folder)})
            data = json.loads(body)
            self.assertEqual((status, data["ok"]), (200, True))
            self.assertEqual(data["files"][0]["backups"], ["Main_20250101"])
            status, body = self.request(port, "POST", "/api/clean/preview", {"folder": ""})
            self.assertEqual(status, 400)
        self.assertIsNone(self.workflow.STATE["job"])
        self.assertEqual(self.books["a.xlsx"].saved, [])


if __name__ == "__main__":
    unittest.main()
