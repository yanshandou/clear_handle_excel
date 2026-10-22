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


if __name__ == "__main__":
    unittest.main()
