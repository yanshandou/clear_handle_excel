import re
import unittest
from unittest.mock import patch

from core import reporter


class ReportMarkdownTests(unittest.TestCase):
    def test_markdown_has_no_dates_or_times(self):
        state = {
            "task_name": "Web 切字扫描（完整）",
            "folder": r"D:\work\excel",
            "status": "completed",
            "progress_pct": 100.0,
            "completed_files": 1,
            "total_files": 1,
            "started_at": "2026-10-22 11:05:50",
            "updated_at": "2026-10-22 11:10:52",
            "files": {
                "a.xlsx": {
                    "index": 1, "name": "a.xlsx", "status": "done",
                    "all_sheets": ["Sheet1"], "current_action": "无待处理切字",
                    "updated_at": "11:06:25",
                },
            },
        }
        with patch.dict(reporter.EXECUTION_STATE, state):
            md = reporter.render_execution_markdown()
        self.assertIn("a.xlsx", md)
        self.assertNotIn("开始时间", md)
        self.assertNotIn("最后更新", md)
        self.assertIsNone(re.search(r"\d{1,2}:\d{2}|\d{4}-\d{2}-\d{2}", md))

    def test_strike_and_bold_markers_hug_sheet_name(self):
        state = {
            "task_name": "清理 Sheet",
            "folder": r"D:\work\excel",
            "status": "completed",
            "progress_pct": 100.0,
            "completed_files": 1,
            "total_files": 1,
            "files": {
                "a.xlsx": {
                    "index": 1, "name": "a.xlsx", "status": "done",
                    "all_sheets": ["基本情報", "基本情報20210308 ", "\u3000一覧 "],
                    "deleted_backups": ["基本情報20210308 "],
                    "uncertain_sheets": ["\u3000一覧 "],
                },
            },
        }
        with patch.dict(reporter.EXECUTION_STATE, state):
            md = reporter.render_execution_markdown()
        self.assertIn("- ~~基本情報20210308~~ *(", md)
        self.assertIn("- **一覧** *(", md)
        self.assertNotIn("~~基本情報20210308 ~~", md)
        self.assertNotIn("**\u3000一覧 **", md)

    def test_append_record_strips_spaces_inside_strike(self):
        import tempfile
        from pathlib import Path
        from core import common
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "record.md"
            with patch.object(common, "RECORD_PATH", path):
                common.append_record("a.xlsx", ["基本情報", " 基本情報20210308\u3000"], [" 基本情報20210308\u3000"])
            text = path.read_text(encoding="utf-8")
        self.assertIn("- ~~基本情報20210308~~\n", text)
        self.assertIn("- 基本情報\n", text)

    def test_md_wrap_keeps_blank_name(self):
        from core.common import md_wrap
        self.assertEqual(md_wrap("  ", "~~"), "  ")
        self.assertEqual(md_wrap("\u3000表 ", "~~"), "~~表~~")


if __name__ == "__main__":
    unittest.main()
