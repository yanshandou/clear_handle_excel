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
from types import SimpleNamespace
from http.server import ThreadingHTTPServer
from pathlib import Path
from unittest.mock import Mock, patch

from core import excel_ops, fix_images, storage


class FakeSheet:
    def __init__(self, book, name, visible=-1):
        self.book = book
        self.Name = name
        self.Visible = visible
        self.tab_color = None  # 标签颜色 RRGGBB；None = 没设
        self.ExportAsFixedFormat = Mock(side_effect=self._export)

    @property
    def Tab(self):
        if self.tab_color is None:
            return SimpleNamespace(ColorIndex=-4142, Color=16777215)
        r, g, b = (int(self.tab_color[i:i + 2], 16) for i in (0, 2, 4))
        return SimpleNamespace(ColorIndex=1, Color=r | (g << 8) | (b << 16))

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
    def __init__(self, books, root_folder):
        self.books = books
        self.root_folder = root_folder
        self.opened = []
        self.Workbooks = Mock()
        self.Workbooks.Open.side_effect = self.open
        self.Quit = Mock()

    def open(self, path, ReadOnly=False, UpdateLinks=3):
        p = Path(path)
        try:
            rel = p.resolve().relative_to(self.root_folder.resolve()).as_posix()
        except ValueError:
            rel = p.name
        key = rel if rel in self.books else p.name
        self.opened.append((key, ReadOnly))
        return self.books[key]


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
        self.patch_value(fix_images, "IMAGES_DIR", self.root / "_report_images")
        self.patch_value(reporter, "REPORT_STATE_PATH", self.root / "legacy.json")
        self.events = []
        self.books = {}
        self.excel = FakeExcel(self.books, self.folder)
        self.dispatch = Mock(return_value=self.excel)
        self.normalize = Mock()
        self.patch_value(workflow, "pythoncom", Mock())
        self.patch_value(workflow, "win32", Mock(DispatchEx=self.dispatch))
        self.patch_value(excel_ops, "win32", workflow.win32)  # 后台 Excel 由 excel_ops 启动
        self.patch_value(workflow, "normalize_workbook_view", self.normalize)
        self.patch_value(workflow, "ExcelWatchdog", Mock())
        self.patch_value(workflow, "scan_workbook_sheets", self.read_sheet_names)

    def patch_value(self, target, name, value):
        patcher = patch.object(target, name, value)
        patcher.start()
        self.addCleanup(patcher.stop)

    def add_book(self, name, sheets, hidden=()):
        path = self.folder / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(f"original {name}".encode())
        self.books[name] = FakeWorkbook(name, sheets, hidden, self.events)
        return path

    def read_sheet_names(self, path, get_excel=None, tab_colors=None):
        """模拟 openpyxl 只读读取工作表名；.xls 通过传入的共享 Excel 打开"""
        if get_excel is not None:
            book = get_excel().Workbooks.Open(str(path), ReadOnly=True)
            names = [sheet.Name for sheet in book.sheets]
            book.Close(SaveChanges=False)
            if tab_colors is not None:
                tab_colors.update({s.Name: s.tab_color for s in book.sheets if s.tab_color})
            return names, self.workflow.find_backup_sheets(names)
        p = Path(path)
        try:
            rel = p.resolve().relative_to(self.folder.resolve()).as_posix()
        except ValueError:
            rel = p.name
        book = self.books.get(rel) or self.books.get(p.name)
        if book is None:
            raise ValueError("文件已损坏")
        names = [sheet.Name for sheet in book.sheets]
        if tab_colors is not None:
            tab_colors.update({s.Name: s.tab_color for s in book.sheets if s.tab_color})
        return names, self.workflow.find_backup_sheets(names)

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

    def test_clean_empty_open_reports_open_failure_and_keeps_file_uncompleted(self):
        path = self.add_book("a.xlsx", ["Sheet1", "Sheet1_bak"])
        self.excel.Workbooks.Open.side_effect = lambda *args, **kwargs: None
        job = self.run_task("clean_sheets")
        self.assertNotIn(str(path), job["completed"])
        error = self.report_files()["a.xlsx"]["error"]
        self.assertIn("Open 返回空对象", error)
        self.assertNotIn("Close", error)
        self.assertEqual(self.books["a.xlsx"].saved, [])

    def test_clean_sheets_backs_up_before_deleting_only_backups_with_original(self):
        self.add_book("a.xlsx", ["機能設計", "機能設計_20250315_bak", "基本情報_20241003", "Sheet1 (2)"])
        self.add_book("b.xlsx", ["Sheet1"])
        job = self.run_task("clean_sheets")

        self.assertEqual(job["status"], "completed")
        # 没有 基本情報、Sheet1 原表，带标记的这两张是在用的表，保留
        self.assertEqual(self.books["a.xlsx"].saved, [["機能設計", "基本情報_20241003", "Sheet1 (2)"]])
        self.assertEqual(
            self.events[:2],
            [("delete", "機能設計_20250315_bak"), ("save", "a.xlsx")],
        )
        self.assertNotIn("backup_folder", job["options"])
        # 预检没有备份表的文件不会用 Excel 打开
        self.assertEqual(self.excel.opened, [("a.xlsx", False)])
        files = self.report_files()
        self.assertEqual(files["a.xlsx"]["deleted_backups"], ["機能設計_20250315_bak"])
        self.assertEqual(files["b.xlsx"]["status"], "done")

    def test_dated_sheets_without_original_are_not_deleted(self):
        # 只有 基本情報_20241003、没有 基本情報 时，带日期的就是在用的表
        self.add_book("SSPLG1100_個別支援計画書情報.xlsx", [
            "変更履歴", "個別支援計画書情報", "基本情報_20241003", "具体的な支援目標及び支援計画_20241003",
        ])
        job = self.run_task("clean_sheets")

        self.assertEqual(job["status"], "completed")
        self.dispatch.assert_not_called()
        self.assertEqual(len(self.books["SSPLG1100_個別支援計画書情報.xlsx"].sheets), 4)
        self.assertEqual(
            self.report_files()["SSPLG1100_個別支援計画書情報.xlsx"]["current_action"], "没有备份表，无需精简",
        )

    def test_clean_sheets_skips_excel_entirely_when_no_backups(self):
        self.add_book("a.xlsx", ["Sheet1"])
        job = self.run_task("clean_sheets")

        self.assertEqual(job["status"], "completed")
        self.dispatch.assert_not_called()

    def test_preview_reuses_one_excel_for_all_xls_files(self):
        self.add_book("a.xls", ["Main", "Main_20250101"])
        self.add_book("sub/b.xls", ["X"])
        self.add_book("c.xlsx", ["Y"])
        data = self.workflow.preview_backup_sheets(str(self.folder))

        files = {f["name"]: f for f in data["files"]}
        self.assertEqual(files["a.xls"]["backups"], ["Main_20250101"])
        self.assertEqual(files["sub/b.xls"]["sheets"], ["X"])
        self.assertEqual(files["c.xlsx"]["sheets"], ["Y"])
        self.dispatch.assert_called_once()
        self.excel.Quit.assert_called_once()
        self.assertEqual([key for key, _ in self.excel.opened], ["a.xls", "sub/b.xls"])

    def test_export_pdf_is_read_only_and_skips_hidden_history_and_backup_sheets(self):
        self.reporter.init_execution_batch(
            "old", self.folder, [], backup_folder=self.root / "old_backup",
        )
        self.add_book(
            "spec.xlsx", ["表紙", "変更履歴", "表紙_20250101", "基本情報_20241003", "隠し"], hidden=["隠し"],
        )
        job = self.run_task("export_pdf")

        self.assertEqual(job["status"], "completed")
        self.assertEqual(self.excel.opened, [("spec.xlsx", True)])
        book = self.books["spec.xlsx"]
        self.assertEqual(book.saved, [])
        # 表紙_20250101 有原表，是备份表；基本情報_20241003 没有原表，照常导出，且直接保存在 Excel 所在目录下
        self.assertEqual(
            [s.Name for s in book.sheets if s.ExportAsFixedFormat.called], ["表紙", "基本情報_20241003"],
        )
        pdf_folder = Path(job["options"]["pdf_folder"])
        self.assertEqual(pdf_folder, self.folder.resolve())
        self.assertEqual(
            sorted(p.name for p in self.folder.glob("*.pdf")),
            ["spec_基本情報_20241003.pdf", "spec_表紙.pdf"],
        )
        self.assertNotIn("backup_folder", job["options"])
        report = self.reporter.get_execution_report_snapshot()
        self.assertEqual(report["files"]["spec.xlsx"]["exported_pdfs"], 2)
        self.assertEqual(report["pdf_folder"], str(pdf_folder))
        # 新批次不沿用上一批的备份目录
        self.assertEqual(report["backup_folder"], "")

    def test_standalone_pdf_export_keeps_inspection_exempt_layout_sheets(self):
        self.add_book("book.xlsx", ["レイアウト", "レイアウト(複数ページ)", "雛形", "JSONサンプル"])
        job = self.run_task("export_pdf")
        self.assertEqual(job["status"], "completed")
        self.assertEqual([s.Name for s in self.books["book.xlsx"].sheets if s.ExportAsFixedFormat.called],
                         ["レイアウト", "レイアウト(複数ページ)", "雛形", "JSONサンプル"])
        self.assertEqual(self.books["book.xlsx"].saved, [])

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
        self.assertEqual((self.folder / "a_Sheet1.pdf").read_bytes(), b"pdf of a.xlsm")
        self.assertEqual((self.folder / "a_Sheet1_2.pdf").read_bytes(), b"pdf of a.xlsx")

    def test_normalize_view_backs_up_then_saves_each_workbook(self):
        self.add_book("a.xlsx", ["Sheet1", "Sheet2"])
        job = self.run_task("normalize_view")

        self.assertEqual(job["status"], "completed")
        self.assertIs(self.normalize.call_args.args[1], self.books["a.xlsx"])
        self.assertEqual(self.events[0], ("save", "a.xlsx"))
        self.assertEqual(len(self.books["a.xlsx"].saved), 1)
        self.assertNotIn("backup_folder", job["options"])
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
        self.add_book("a.xlsx", ["表紙", "Main_20250101", "Main"])
        self.add_book("b.xlsx", ["X_20250101"])
        self.add_book("c.xlsx", ["Main"])
        (self.folder / "d.xlsx").write_bytes(b"broken")
        (self.folder / "~$a.xlsx").write_bytes(b"lock file")
        data = self.workflow.preview_backup_sheets(str(self.folder))

        files = {f["name"]: f for f in data["files"]}
        self.assertEqual(sorted(files), ["a.xlsx", "b.xlsx", "c.xlsx", "d.xlsx"])
        self.assertEqual(files["a.xlsx"]["backups"], ["Main_20250101"])
        # 全部工作表按原顺序返回，页面据此同时列出保留与删除的工作表
        self.assertEqual(files["a.xlsx"]["sheets"], ["表紙", "Main_20250101", "Main"])
        # 没有原表 X，X_20250101 不是备份表
        self.assertEqual(files["b.xlsx"]["backups"], [])
        self.assertEqual(files["b.xlsx"]["sheets"], ["X_20250101"])
        self.assertEqual(files["b.xlsx"]["note"], "")
        self.assertEqual((files["c.xlsx"]["backups"], files["c.xlsx"]["sheets"]), ([], ["Main"]))
        self.assertTrue(files["d.xlsx"]["note"].startswith("无法读取"))
        self.assertEqual(files["d.xlsx"]["sheets"], [])
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
                self.assertIn("开始全自动处理", body)
                self.assertIn("确认并处理全部标记", body)
            status, body = self.request(port, "GET", "/api/overview")
            self.assertEqual(status, 200)
            overview = json.loads(body)
            self.assertFalse(overview["busy"])
            self.assertNotIn("items", overview)

    def test_clean_page_route(self):
        with self.serve_api() as port:
            for path in ("/clean", "/clean.html"):
                status, body = self.request(port, "GET", path)
                self.assertEqual(status, 200)
                self.assertIn("精简 Sheet", body)

    def test_browser_disconnect_during_preview_is_logged_quietly(self):
        import web_clip_fixer

        self.add_book("a.xlsx", ["Main"])
        writes = []

        def broken_send(handler, body, content_type, status=200, filename=None):
            writes.append(status)
            raise ConnectionAbortedError(10053, "connection aborted by host")

        with self.serve_api() as port, \
                patch.object(web_clip_fixer.WebHandler, "_send_bytes", broken_send), \
                patch.object(web_clip_fixer, "console_print") as console:
            with self.assertRaises(http.client.RemoteDisconnected):
                self.request(port, "POST", "/api/clean/preview", {"folder": str(self.folder)})
        # 只尝试发送一次结果，不再追加 500 错误响应，也不写入任务日志
        self.assertEqual(writes, [200])
        self.assertFalse(any("❌" in line for line in self.workflow.STATE["logs"]))
        console.assert_called_once()
        self.assertIn("浏览器已断开", console.call_args.args[0])

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

    def test_api_scan_deletes_backups_only_when_explicitly_requested(self):
        self.add_book("a.xlsx", ["Sheet1"])
        with self.serve_api() as port:
            for payload, expected in (
                ({"folder": str(self.folder)}, False),
                ({"folder": str(self.folder), "delete_backups": "true"}, False),
                ({"folder": str(self.folder), "delete_backups": 1}, False),
                ({"folder": str(self.folder), "delete_backups": True}, True),
            ):
                with self.subTest(payload=payload):
                    self.workflow.STATE["busy"] = False
                    status, body = self.request(port, "POST", "/api/scan", payload)
                    self.assertEqual((status, json.loads(body)["ok"]), (200, True))
                    job = storage.load_snapshot("workflow")["job"]
                    self.assertIs(job["options"]["delete_backups"], expected)
        # 不再作为「上次提交的参数」保存，页面每次都从不勾选开始
        self.assertNotIn("delete_backups", self.workflow.STATE["settings"])

    def test_api_clean_preview_is_read_only(self):
        self.add_book("a.xlsx", ["Main", "Main_20250101", "❓Draft"])
        with self.serve_api() as port:
            status, body = self.request(port, "POST", "/api/clean/preview", {"folder": str(self.folder)})
            data = json.loads(body)
            self.assertEqual((status, data["ok"]), (200, True))
            self.assertEqual(data["files"][0]["backups"], ["Main_20250101"])
            self.assertEqual(data["files"][0]["uncertain"], ["❓Draft"])
            status, body = self.request(port, "POST", "/api/clean/preview", {"folder": ""})
            self.assertEqual(status, 400)
        self.assertIsNone(self.workflow.STATE["job"])
        self.assertEqual(self.books["a.xlsx"].saved, [])

    def test_clean_sheets_with_custom_plan_deletes_checked_and_keeps_uncertain(self):
        # 用户自定义勾选：取消默认勾选的 Main_20250101（改为待定），手动勾选 SubSheet 删除
        self.add_book("spec.xlsx", ["Main", "Main_20250101", "SubSheet"])
        self.add_book("plain.xlsx", ["OnlySheet"])

        plan = {
            "spec.xlsx": {
                "delete": ["SubSheet"],
                "uncertain": ["Main_20250101"],
            },
            "plain.xlsx": {
                "delete": [],
                "uncertain": [],
            },
        }
        self.assertTrue(
            self.workflow.queue_job("clean_sheets", {"folder": str(self.folder), "plan": plan})
        )
        self.workflow.run_pending_workflow()
        job = self.workflow.STATE["job"]
        self.assertEqual(job["status"], "completed")

        # SubSheet 被删除，Main_20250101 标为待定只是不删，表名不加 ❓
        self.assertEqual(self.books["spec.xlsx"].saved, [["Main", "Main_20250101"]])
        # 工作簿文件名不变，不生成 _backup_ 目录
        self.assertNotIn("backup_folder", job["options"])
        self.assertTrue((self.folder / "spec.xlsx").exists())
        # plain.xlsx 无需修改，未用 Excel 打开
        self.assertEqual(self.excel.opened, [("spec.xlsx", False)])

        files = self.report_files()
        self.assertEqual(files["spec.xlsx"]["deleted_backups"], ["SubSheet"])
        self.assertEqual(files["spec.xlsx"]["uncertain_sheets"], ["Main_20250101"])
        self.assertEqual(files["spec.xlsx"]["renamed_to"], "")

    def test_clean_sheets_deletes_old_versions_and_strips_new_prefix(self):
        self.add_book("layout.xlsx", [
            "新レイアウト", "旧レイアウト", "新レイアウト(介護給付費請求書1)", "旧レイアウト (介護給付費請求書1)",
            "新画面", "旧画面", "画面",
        ])
        self.run_task("clean_sheets")

        # 新画面 去掉「新」后与现有的 画面 重名，保持原名
        self.assertEqual(self.books["layout.xlsx"].saved, [
            ["レイアウト", "レイアウト(介護給付費請求書1)", "新画面", "画面"],
        ])
        files = self.report_files()
        self.assertEqual(files["layout.xlsx"]["deleted_backups"], [
            "旧レイアウト", "旧レイアウト (介護給付費請求書1)", "旧画面",
        ])
        self.assertIn("2 个新版 Sheet 已去掉「新」字", files["layout.xlsx"]["current_action"])

    def test_clean_sheets_previews_and_deletes_unpaired_old_layout(self):
        name = "SSPLE0002_個別支援計画書（様式２）.xlsx"
        sheets = [
            "変更履歴", "概要", "帳票_出力内容", "レイアウト(1ページ)",
            "レイアウト(複数ページ)", "旧レイアウト", "JSONサンプル",
        ]
        self.add_book(name, sheets)

        preview = self.workflow.preview_backup_sheets(str(self.folder))["files"][0]
        self.assertEqual(preview["backups"], ["旧レイアウト"])
        self.assertEqual(preview["uncertain"], [])
        self.assertEqual(preview["new_old_pairs"], {})
        self.assertEqual(self.books[name].saved, [])

        job = self.run_task("clean_sheets")
        self.assertEqual(job["status"], "completed")
        self.assertEqual(self.books[name].saved, [[s for s in sheets if s != "旧レイアウト"]])
        self.assertEqual(self.report_files()[name]["deleted_backups"], ["旧レイアウト"])

    def test_unpaired_old_sheet_can_be_manually_kept(self):
        self.add_book("layout.xlsx", ["レイアウト(1ページ)", "旧レイアウト"])
        plan = {"layout.xlsx": {"delete": [], "uncertain": []}}
        self.assertTrue(self.workflow.queue_job("clean_sheets", {"folder": str(self.folder), "plan": plan}))
        self.workflow.run_pending_workflow()

        self.assertEqual(self.workflow.STATE["job"]["status"], "completed")
        self.assertEqual([s.Name for s in self.books["layout.xlsx"].sheets], ["レイアウト(1ページ)", "旧レイアウト"])
        self.assertEqual(self.report_files()["layout.xlsx"]["deleted_backups"], [])

    def test_clean_sheets_strips_new_prefix_without_old_version(self):
        # 没有旧表时「新」字也没有意义；只改名、不删表也要用 Excel 打开保存
        self.add_book("layout.xlsx", ["レイアウト", "新レイアウト（２ページ以降）", "新規登録"])
        self.run_task("clean_sheets")

        self.assertEqual(self.books["layout.xlsx"].saved, [
            ["レイアウト", "レイアウト（２ページ以降）", "新規登録"],
        ])
        files = self.report_files()
        self.assertEqual(files["layout.xlsx"]["deleted_backups"], [])
        self.assertIn("1 个新版 Sheet 已去掉「新」字", files["layout.xlsx"]["current_action"])

    def test_new_sheet_keeps_prefix_when_old_version_is_kept(self):
        self.add_book("layout.xlsx", ["新レイアウト", "旧レイアウト", "新帳票", "旧帳票"])
        plan = {"layout.xlsx": {"delete": ["旧帳票"], "uncertain": []}}
        self.assertTrue(
            self.workflow.queue_job("clean_sheets", {"folder": str(self.folder), "plan": plan})
        )
        self.workflow.run_pending_workflow()
        self.assertEqual(self.books["layout.xlsx"].saved, [["新レイアウト", "旧レイアウト", "帳票"]])

    def test_clean_sheets_only_listed_files_leaves_other_files_untouched(self):
        # 预览窗口只处理当前文件：其他文件即使带 ❓ 也不会被恢复原名
        self.add_book("spec.xlsx", ["Main", "Main_20250101"])
        self.add_book("❓other.xlsx", ["Sheet1", "❓Draft"])

        plan = {"spec.xlsx": {"delete": ["Main_20250101"], "uncertain": []}}
        self.assertTrue(self.workflow.queue_job(
            "clean_sheets", {"folder": str(self.folder), "plan": plan, "files": ["spec.xlsx"]},
        ))
        self.workflow.run_pending_workflow()

        self.assertEqual(self.workflow.STATE["job"]["status"], "completed")
        self.assertEqual(self.books["spec.xlsx"].saved, [["Main"]])
        self.assertTrue((self.folder / "❓other.xlsx").exists())
        self.assertEqual(self.books["❓other.xlsx"].saved, [])
        self.assertEqual(sorted(self.report_files()), ["spec.xlsx"])

    def test_validate_clean_files_rejects_invalid_lists(self):
        self.assertEqual(self.workflow.validate_clean_files(["a.xlsx", "a.xlsx", "b.xlsx"]), ["a.xlsx", "b.xlsx"])
        for bad in ([], "a.xlsx", [""], [1]):
            with self.assertRaises(ValueError):
                self.workflow.validate_clean_files(bad)

    def test_recursive_subfolders_scan_clean_and_export_pdf_in_place(self):
        # 不同子文件夹下存在同名 Excel 文件，递归扫描、精简与同目录导出 PDF 互不冲突
        self.add_book("01_画面/spec.xlsx", ["Main", "Main_20250101"])
        self.add_book("02_帳票/spec.xlsx", ["Report", "Draft"])

        preview = self.workflow.preview_backup_sheets(str(self.folder))
        preview_map = {f["name"]: f for f in preview["files"]}
        self.assertEqual(sorted(preview_map), ["01_画面/spec.xlsx", "02_帳票/spec.xlsx"])
        self.assertEqual(preview_map["01_画面/spec.xlsx"]["backups"], ["Main_20250101"])

        plan = {
            "01_画面/spec.xlsx": {"delete": ["Main_20250101"], "uncertain": []},
            "02_帳票/spec.xlsx": {"delete": [], "uncertain": ["Draft"]},
        }
        self.assertTrue(
            self.workflow.queue_job("clean_sheets", {"folder": str(self.folder), "plan": plan})
        )
        self.workflow.run_pending_workflow()
        self.assertEqual(self.books["01_画面/spec.xlsx"].saved, [["Main"]])
        # 待定的 Draft 只是不删，文件与表名都不改，也不用 Excel 打开
        self.assertEqual(self.books["02_帳票/spec.xlsx"].saved, [])
        self.assertTrue((self.folder / "02_帳票" / "spec.xlsx").exists())
        self.assertEqual(self.report_files()["02_帳票/spec.xlsx"]["uncertain_sheets"], ["Draft"])

        # 导出 PDF 时，PDF 直接生成在各自的子文件夹中
        self.run_task("export_pdf")
        self.assertTrue((self.folder / "01_画面" / "spec_Main.pdf").exists())
        self.assertTrue((self.folder / "02_帳票" / "spec_Report.pdf").exists())
        self.assertTrue((self.folder / "02_帳票" / "spec_Draft.pdf").exists())

    def test_gray_tab_sheets_default_to_delete_and_skip_pdf(self):
        # 灰色标签的表不看名字：精简 Sheet 默认删除，导出 PDF 也跳过（用户 2026-10-26）；橙黄问题标签不算
        book = self.add_book("spec.xlsx", ["Main", "Main詳細", "20161101サービスコード選択(削除)", "Report"])
        sheets = {s.Name: s for s in self.books["spec.xlsx"].sheets}
        sheets["Main詳細"].tab_color = "A6A6A6"
        sheets["20161101サービスコード選択(削除)"].tab_color = "808080"
        sheets["Report"].tab_color = "FFC000"
        files = {f["name"]: f for f in self.workflow.preview_backup_sheets(str(self.folder))["files"]}
        self.assertEqual(files["spec.xlsx"]["backups"], ["Main詳細", "20161101サービスコード選択(削除)"])
        self.assertEqual(files["spec.xlsx"]["gray_tabs"], ["Main詳細", "20161101サービスコード選択(削除)"])
        self.assertEqual(files["spec.xlsx"]["uncertain"], [])  # 灰色优先于「名称包含 → 待定」

        self.run_task("export_pdf")
        self.assertTrue((self.folder / "spec_Main.pdf").exists())
        self.assertTrue((self.folder / "spec_Report.pdf").exists())
        self.assertFalse((self.folder / "spec_Main詳細.pdf").exists())
        self.assertEqual(sorted(p.name for p in self.folder.glob("*.pdf")), ["spec_Main.pdf", "spec_Report.pdf"])
        self.assertTrue(book.exists())

    def test_clean_sheets_leaves_existing_question_marks_untouched(self):
        # 以前加过 ❓ 的文件与表：不再自动去掉，也不再改名
        self.add_book("❓spec.xlsx", ["Main", "❓Draft"])
        plan = {"❓spec.xlsx": {"delete": [], "uncertain": []}}
        self.assertTrue(
            self.workflow.queue_job("clean_sheets", {"folder": str(self.folder), "plan": plan})
        )
        self.workflow.run_pending_workflow()

        self.assertEqual(self.books["❓spec.xlsx"].saved, [])
        self.assertTrue((self.folder / "❓spec.xlsx").exists())
        self.assertEqual(self.excel.opened, [])

    def test_api_quick_task_validates_clean_plan(self):
        self.add_book("a.xlsx", ["Main", "Sub"])
        with self.serve_api() as port:
            # 同一工作表不能同时在 delete 和 uncertain 中
            status, body = self.request(port, "POST", "/api/quick_task", {
                "kind": "clean_sheets",
                "folder": str(self.folder),
                "plan": {"a.xlsx": {"delete": ["Sub"], "uncertain": ["Sub"]}},
            })
            self.assertEqual((status, json.loads(body)["ok"]), (400, False))

            # 合法的自定义精简方案
            status, body = self.request(port, "POST", "/api/quick_task", {
                "kind": "clean_sheets",
                "folder": str(self.folder),
                "plan": {"a.xlsx": {"delete": [], "uncertain": ["Sub"]}},
            })
            self.assertEqual((status, json.loads(body)["ok"]), (200, True))

    def run_plan(self, plan):
        self.assertTrue(
            self.workflow.queue_job("clean_sheets", {"folder": str(self.folder), "plan": plan})
        )
        self.workflow.run_pending_workflow()

    def test_clean_sheets_applies_custom_sheet_renames(self):
        self.add_book("a.xlsx", ["振込データ提出先選択", "Main", "Main_20250101", "Sub"])
        self.workflow.STATE["pages"]["p1"] = {
            "wb_path": str(self.folder / "a.xlsx"), "wb_name": "a.xlsx", "sheet_name": "振込データ提出先選択",
        }
        self.run_plan({"a.xlsx": {
            "delete": ["Main_20250101"], "uncertain": [],
            "rename": {"振込データ提出先選択": "提出先選択", "Main_20250101": "删掉的表不改名"},
        }})
        self.assertEqual(self.books["a.xlsx"].saved, [["提出先選択", "Main", "Sub"]])
        files = self.report_files()
        self.assertIn("1 个 Sheet 已改名", files["a.xlsx"]["current_action"])
        # 工作台记录跟着新名字
        self.assertEqual(self.workflow.STATE["pages"]["p1"]["sheet_name"], "提出先選択")

    def test_clean_sheets_rename_only_opens_excel_and_skips_duplicates(self):
        self.add_book("a.xlsx", ["Main", "Sub", "Other"])
        self.run_plan({"a.xlsx": {
            "delete": [], "uncertain": ["Other"],
            # 待定的表也能改名；改成现有表名（不区分大小写）的不改
            "rename": {"Other": "其他", "Sub": "main"},
        }})
        self.assertEqual(self.books["a.xlsx"].saved, [["Main", "Sub", "其他"]])

    def test_custom_rename_replaces_new_prefix_rename(self):
        self.add_book("a.xlsx", ["新レイアウト", "旧レイアウト"])
        self.run_plan({"a.xlsx": {
            "delete": ["旧レイアウト"], "uncertain": [], "rename": {"新レイアウト": "レイアウト v2"},
        }})
        self.assertEqual(self.books["a.xlsx"].saved, [["レイアウト v2"]])
        action = self.report_files()["a.xlsx"]["current_action"]
        self.assertIn("1 个 Sheet 已改名", action)
        self.assertNotIn("去掉「新」字", action)

    def test_validate_clean_plan_checks_excel_sheet_name_rules(self):
        for bad in ["", "a/b", "x" * 32, "'abc", "History", "[1]"]:
            with self.assertRaises(ValueError, msg=bad):
                self.workflow.validate_clean_plan({"a.xlsx": {"delete": [], "uncertain": [], "rename": {"S": bad}}})
        plan = self.workflow.validate_clean_plan(
            {"a.xlsx": {"delete": [], "uncertain": [], "rename": {"S": "新しい名前（２）", "T": "T"}}})
        self.assertEqual(plan["a.xlsx"]["rename"], {"S": "新しい名前（２）"})


if __name__ == "__main__":
    unittest.main()
