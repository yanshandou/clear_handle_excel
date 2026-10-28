import importlib
import io
import json
import tempfile
import unittest
from email.message import Message
from pathlib import Path
from unittest.mock import patch

from core import scan_analysis as analysis, storage


class ScanAnalysisTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        patcher = patch.object(storage, "DATABASE_PATH", self.root / "test.sqlite3")
        patcher.start()
        self.addCleanup(patcher.stop)
        self.paths = [self.root / "SystemA" / "one.xlsx", self.root / "two.xlsx"]
        for path in self.paths:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(b"workbook")
        self.job = {}

    def begin(self):
        return analysis.begin(self.job, self.root, self.paths, "只扫描")

    def test_success_and_cache_are_distinct_and_operation_totals_are_consistent(self):
        self.begin()
        with patch.object(analysis.time, "perf_counter", side_effect=[0, 1, 7, 10]):
            with analysis.attempt(self.job, self.paths[0]):
                with analysis.operation("export", "Sheet1"):
                    analysis.count("pdf_exports")
                analysis.note(exempt_sheets=["雛形"])
                analysis.detected("Sheet1", {"a": {"kind": "clip", "location_confirmed": False}})
        analysis.cached(self.job, self.paths[1])
        analysis.finish(self.job)
        report = analysis.report(self.job["analysis_id"])
        self.assertEqual(report["summary"]["seconds"], 10)
        self.assertEqual(report["summary"]["completed"], 1)
        self.assertEqual(report["summary"]["cached"], 1)
        self.assertEqual(report["summary"]["unconfirmed"], 1)
        self.assertEqual(report["summary"]["exempt_sheets"], 1)
        self.assertAlmostEqual(sum(op["share"] for op in report["operations"]), 1)
        self.assertEqual(report["systems"][0]["name"], "SystemA")
        self.assertEqual(report["sheets"][0]["sheet"], "Sheet1")
        self.assertTrue(report["sheets"][0]["complete"])
        self.assertIn("优先排查：PDF 导出", [f["title"] for f in report["findings"]])

    def test_timeout_and_pending_files_are_not_reported_as_clean(self):
        self.begin()
        with self.assertRaisesRegex(RuntimeError, "Excel 卡住"):
            with analysis.attempt(self.job, self.paths[0]):
                analysis.detected("S1", {"a": {"kind": "clip"}})
                raise RuntimeError("Excel 卡住")
        analysis.timeout(self.job, self.paths[0], "S2", 180)
        analysis.finish(self.job)
        report = analysis.report(self.job["analysis_id"])
        self.assertEqual(report["summary"]["timeout"], 1)
        self.assertEqual(report["summary"]["pending"], 1)
        self.assertEqual(report["summary"]["issues"], 0)
        self.assertEqual(report["files"][0]["stuck_sheet"], "S2")
        self.assertTrue(any(f["level"] == "high" for f in report["findings"]))

    def test_resume_retains_failed_attempt_time_without_counting_stale_issues(self):
        first = self.begin()
        with patch.object(analysis.time, "perf_counter", side_effect=[0, 4]):
            with self.assertRaises(ValueError):
                with analysis.attempt(self.job, self.paths[0]):
                    analysis.detected("old", {"a": {"kind": "clip"}})
                    raise ValueError("first failure")
        analysis.finish(self.job, interrupted=True)
        self.assertEqual(self.begin()["id"], first["id"])
        with patch.object(analysis.time, "perf_counter", side_effect=[5, 11]):
            with analysis.attempt(self.job, self.paths[0]):
                analysis.detected("new", {})
        analysis.finish(self.job)
        report = analysis.report(first["id"])
        self.assertEqual(report["summary"]["seconds"], 10)
        self.assertEqual(report["summary"]["retries"], 1)
        self.assertEqual(report["summary"]["issues"], 0)
        self.assertEqual(report["summary"]["checked_sheets"], 1)
        self.assertFalse(next(s for s in report["sheets"] if s["sheet"] == "old")["complete"])

    def test_history_is_persistent_and_can_be_filtered_by_folder(self):
        first = self.begin()["id"]
        analysis.finish(self.job)
        other = {}
        analysis.begin(other, self.root / "SystemA", [self.paths[0]], "完整扫描")
        importlib.reload(analysis)
        self.assertEqual(analysis.history(str(self.root))[0]["id"], first)
        self.assertEqual(analysis.report(folder=str(self.root))["run"]["id"], first)
        self.assertIsNone(analysis.report("unknown"))

    def test_excel_restart_is_counted_without_overwriting_failure_status(self):
        self.begin()
        with patch.object(analysis.time, "perf_counter", side_effect=[0, 2]):
            with self.assertRaises(ValueError):
                with analysis.attempt(self.job, self.paths[0]):
                    raise ValueError("first failure")
        with patch.object(analysis.time, "perf_counter", side_effect=[2, 3, 8, 9]):
            with analysis.recovery(self.job, self.paths[0]):
                pass
        report = analysis.report(self.job["analysis_id"])
        self.assertEqual(report["summary"]["seconds"], 9)
        self.assertEqual(report["summary"]["error"], 1)
        self.assertEqual(report["files"][0]["operations"]["restart"]["seconds"], 5)

    def test_markdown_includes_evidence_and_escapes_table_cells(self):
        self.begin()
        with analysis.attempt(self.job, self.paths[0]):
            analysis.note(missing_sheets=["gone"], hidden_text_supported=False)
        data = analysis.report(self.job["analysis_id"])
        data["files"][0]["name"] = "name|next\nrow.xlsx"
        text = analysis.markdown(data)
        self.assertIn("部分检查范围存在缺口", text)
        self.assertIn("name\\|next row.xlsx", text)
        self.assertIn("系统耗时", text)

    def test_analysis_api_and_download_work_without_changing_scan_state(self):
        from web_clip_fixer import WebHandler
        self.begin()
        for url, content_type in (("/api/scan-analysis", "application/json"),
                                  ("/api/scan-analysis/markdown", "text/markdown"),
                                  ("/analysis", "text/html")):
            with self.subTest(url=url):
                handler = object.__new__(WebHandler)
                handler.path, handler.headers, handler.wfile = url, Message(), io.BytesIO()
                statuses, headers = [], {}
                handler.send_response = statuses.append
                handler.send_header = headers.__setitem__
                handler.end_headers = lambda: None
                handler._handle_get()
                self.assertEqual(statuses, [200])
                self.assertTrue(headers["Content-Type"].startswith(content_type))
                if url == "/api/scan-analysis":
                    body = json.loads(handler.wfile.getvalue())
                    self.assertEqual(body["report"]["run"]["id"], self.job["analysis_id"])


if __name__ == "__main__":
    unittest.main()
