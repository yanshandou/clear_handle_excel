import copy
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import reporter, storage


class ReportAccumulateTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.folder = self.root / "excel"
        self.folder.mkdir()
        for name, value in [(storage, ("DATABASE_PATH", self.root / "t.sqlite3")),
                            (reporter, ("RECORD_PATH", self.root / "record.md"))]:
            p = patch.object(name, *value)
            p.start()
            self.addCleanup(p.stop)
        saved = copy.deepcopy(reporter.EXECUTION_STATE)
        self.addCleanup(lambda: (reporter.EXECUTION_STATE.clear(), reporter.EXECUTION_STATE.update(saved)))
        reporter.EXECUTION_STATE.update(folder="", files={}, steps=[], batch=[], started_at="")

    def test_same_folder_keeps_earlier_steps_and_other_folder_starts_over(self):
        a, b = self.folder / "a.xlsx", self.folder / "b.xlsx"
        reporter.init_execution_batch("精简 Sheet（删除备份表）", self.folder, [a, b])
        reporter.update_file_progress("a.xlsx", status="done", all_sheets=["S1", "S1 (2)"], deleted_backups=["S1 (2)"])
        reporter.update_file_progress("b.xlsx", status="done")
        reporter.finish_execution_batch()

        reporter.init_execution_batch("批量导出 PDF", self.folder, [a], pdf_folder=self.folder)
        reporter.update_file_progress("a.xlsx", status="done", exported_pdfs=1)
        reporter.finish_execution_batch()
        state = reporter.get_execution_report_snapshot()
        self.assertEqual(state["files"]["a.xlsx"]["deleted_backups"], ["S1 (2)"])
        self.assertEqual(state["files"]["a.xlsx"]["exported_pdfs"], 1)
        self.assertIn("b.xlsx", state["files"])
        self.assertEqual((state["completed_files"], state["total_files"]), (1, 1))  # 进度只算这一步的文件
        self.assertIn("精简 Sheet（删除备份表） → 批量导出 PDF", state["markdown"])
        self.assertIn("已作为备份表清理", (self.folder / "_执行记录报告.md").read_text(encoding="utf-8"))

        other = self.root / "other"
        other.mkdir()
        reporter.init_execution_batch("批量导出 PDF", other, [other / "c.xlsx"])
        state = reporter.get_execution_report_snapshot()
        self.assertEqual(list(state["files"]), ["c.xlsx"])
        self.assertEqual(state["steps"], ["批量导出 PDF"])

    def test_progress_writes_are_batched_but_flushed_on_finish(self):
        a = self.folder / "a.xlsx"
        reporter.init_execution_batch("视图标准化（A1 + 100% 缩放）", self.folder, [a])
        with patch.object(reporter, "save_snapshot") as save:
            for _ in range(20):
                reporter.update_file_progress("a.xlsx", current_action="处理中")
            self.assertLessEqual(save.call_count, 1)
            reporter.finish_execution_batch()
            self.assertGreaterEqual(save.call_count, 1)
            self.assertIsNone(reporter._write_timer)


if __name__ == "__main__":
    unittest.main()
