import http.client
import importlib
import json
import tempfile
import threading
import unittest
from contextlib import contextmanager
from http.server import ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

from core import fix_images, storage


class ProgressTransferTests(unittest.TestCase):
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
        self.record_path = self.root / "record.md"
        self.patch_value(reporter, "RECORD_PATH", self.record_path)
        self.patch_value(fix_images, "IMAGES_DIR", self.root / "_report_images")
        self.patch_value(reporter, "REPORT_STATE_PATH", self.root / "legacy.json")

        self.wb_path = self.folder / "book.xlsx"
        self.wb_path.write_bytes(b"workbook v1")
        version = self.workflow._file_version(self.wb_path)
        self.workflow.STATE.update(folder=str(self.folder), backup_folder=str(self.folder / "_backup_x"))
        self.workflow.STATE["checked_files"][str(self.wb_path)] = version
        self.workflow.STATE["items"] = {
            "item_1": self.item("item_1", status="accepted", manual_confirmation={
                "file_version": version, "confirmed_at": "2026-10-21T10:00:00", "previous_status": "pending",
            }),
            "item_2": self.item("item_2"),
        }
        self.workflow.STATE["job"] = {
            "kind": "fix", "status": "interrupted", "completed": [],
            "options": {"selected_ids": ["item_2"], "step_pt": 3.0, "max_iter": 6},
            "pending_adjustment": {"item_id": "item_2", "target_height": 21.0, "iterations": 1},
        }
        self.workflow._persist_state()
        self.reporter.init_execution_batch("Web 闭环行高微调", self.folder, [self.wb_path])
        self.reporter.update_file_progress(self.wb_path.name, status="needs_fix", clip_detected=2)

    def patch_value(self, target, name, value):
        patcher = patch.object(target, name, value)
        patcher.start()
        self.addCleanup(patcher.stop)

    def item(self, iid, status="pending", **extra):
        return {
            "id": iid, "seq": 1, "wb_path": str(self.wb_path), "wb_name": self.wb_path.name,
            "sheet_name": "Sheet1", "excel_row": 2, "cell_addr": "A2", "row_height": 18.0,
            "pdf_path": str(self.folder / "book_Sheet1.pdf"), "page_index": 0, "text": iid,
            "bbox": [0, 0, 10, 10], "image": "img", "status": status, "iterations": 1, **extra,
        }

    def exported(self):
        # 经过 JSON 序列化，与真实下载的文件一致
        return json.loads(json.dumps(self.workflow.export_progress(), ensure_ascii=False))

    def lose_progress(self):
        """模拟误操作：开始了新的扫描，旧结果与执行记录被替换"""
        self.workflow.STATE["items"].clear()
        self.workflow.STATE["job"] = None
        self.workflow._persist_state()
        self.reporter.init_execution_batch("另一批任务", self.root, [])

    def test_restore_brings_back_results_confirmations_and_report(self):
        data = self.exported()
        self.lose_progress()
        result = self.workflow.import_progress(data)

        self.assertEqual(result["warnings"], [])
        self.assertEqual(result["items"], 2)
        saved = storage.load_snapshot("workflow")
        self.assertEqual(saved["items"]["item_1"]["status"], "accepted")
        self.assertEqual(saved["checked_files"], data["workflow"]["checked_files"])
        self.assertFalse(saved["busy"])
        report = self.reporter.get_execution_report_snapshot()
        self.assertEqual(report["task_name"], "Web 闭环行高微调")
        self.assertEqual(report["files"]["book.xlsx"]["clip_detected"], 2)
        self.assertIn("已中断", self.record_path.read_text(encoding="utf-8"))

    def test_unchanged_files_keep_interrupted_job_resumable(self):
        data = self.exported()
        self.lose_progress()
        result = self.workflow.import_progress(data)

        self.assertTrue(result["can_resume"])
        job = storage.load_snapshot("workflow")["job"]
        self.assertEqual(job["status"], "interrupted")
        self.assertEqual(job["pending_adjustment"]["target_height"], 21.0)
        self.assertTrue(self.workflow.get_overview_snapshot()["can_resume"])

    def test_file_changed_after_export_cancels_resume_but_keeps_confirmation(self):
        data = self.exported()
        self.wb_path.write_bytes(b"workbook v2 edited later")
        self.lose_progress()
        result = self.workflow.import_progress(data)

        self.assertFalse(result["can_resume"])
        self.assertIn("book.xlsx", result["warnings"][0])
        saved = storage.load_snapshot("workflow")
        self.assertEqual(saved["job"]["status"], "cancelled")
        self.assertIsNone(saved["job"]["pending_adjustment"])
        # 扫描结果和人工确认都保留；保存文件不会让人工确认失效
        self.assertEqual(set(saved["items"]), {"item_1", "item_2"})
        self.assertEqual(saved["items"]["item_1"]["status"], "accepted")
        self.assertEqual(self.reporter.get_execution_report_snapshot()["status"], "completed")

    def test_missing_file_also_cancels_resume(self):
        data = self.exported()
        self.wb_path.unlink()
        result = self.workflow.import_progress(data)

        self.assertFalse(result["can_resume"])
        self.assertEqual(storage.load_snapshot("workflow")["job"]["status"], "cancelled")

    def test_busy_task_blocks_export_and_import(self):
        data = self.exported()
        self.workflow.STATE["busy"] = True
        with self.assertRaisesRegex(ValueError, "任务执行中"):
            self.workflow.export_progress()
        with self.assertRaisesRegex(ValueError, "任务执行中"):
            self.workflow.import_progress(data)
        self.assertTrue(self.workflow.STATE["busy"])

    def test_invalid_files_are_rejected_without_changing_progress(self):
        data = self.exported()
        self.lose_progress()
        before_workflow = storage.load_snapshot("workflow")
        before_report = storage.load_snapshot("report")
        broken_workflow = json.loads(json.dumps(data))
        del broken_workflow["workflow"]["items"]
        broken_report = json.loads(json.dumps(data))
        broken_report["report"] = {"task_name": "x"}
        for invalid in (
            {}, {"format": "other"}, {**data, "version": 2}, broken_workflow,
            {**data, "file_versions": None}, broken_report,
        ):
            with self.subTest(invalid=str(invalid)[:60]):
                with self.assertRaises(ValueError):
                    self.workflow.import_progress(invalid)
        self.assertEqual(storage.load_snapshot("workflow"), before_workflow)
        self.assertEqual(storage.load_snapshot("report"), before_report)
        self.assertEqual(self.workflow.STATE["items"], {})

    @contextmanager
    def serve_api(self):
        import web_clip_fixer

        importlib.reload(web_clip_fixer)
        server = ThreadingHTTPServer(("127.0.0.1", 0), web_clip_fixer.WebHandler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            yield server.server_port
        finally:
            server.shutdown()
            server.server_close()
            thread.join()

    def request(self, port, method, path, body=None):
        connection = http.client.HTTPConnection("127.0.0.1", port, timeout=5)
        try:
            connection.request(method, path, body, {"Content-Type": "application/json"})
            response = connection.getresponse()
            return response.status, response.getheader("Content-Disposition"), response.read()
        finally:
            connection.close()

    def test_api_export_download_and_import(self):
        with self.serve_api() as port:
            status, disposition, body = self.request(port, "GET", "/api/progress/export")
            self.assertEqual(status, 200)
            self.assertRegex(disposition, r'attachment; filename="excel_progress_\d{8}_\d{6}\.json"')
            self.assertEqual(json.loads(body)["format"], "clear_handle_excel.progress")

            self.lose_progress()
            status, _, response = self.request(port, "POST", "/api/progress/import", body)
            result = json.loads(response)
            self.assertEqual((status, result["ok"], result["items"]), (200, True, 2))

            status, _, response = self.request(port, "POST", "/api/progress/import", b"{not json")
            self.assertEqual(status, 400)
            self.assertIn("不是有效的 JSON", json.loads(response)["error"])

            self.workflow.STATE["busy"] = True
            status, _, response = self.request(port, "GET", "/api/progress/export")
            self.assertEqual(status, 409)
            self.assertIn("任务执行中", json.loads(response)["error"])


if __name__ == "__main__":
    unittest.main()
