import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import workflow as wf
from core.pdf_detector import fitz


class DetectTimingTests(unittest.TestCase):
    """检测一张表时记下各步耗时和问 Excel 显示文字的次数（用户 2026-11-07：一张表检测 108 秒，要找出慢在哪）"""

    def test_detect_logs_step_seconds_and_display_text_reads(self):
        with tempfile.TemporaryDirectory() as temp:
            pdf = Path(temp) / "s.pdf"
            doc = fitz.open()
            doc.new_page().insert_text((72, 72), "abc")
            doc.save(pdf)
            doc.close()
            sheet = Mock(Name="S1")
            sheet.UsedRange = Mock(Row=1, Column=1, Value=((12345, "abc"),))
            sheet.Rows.return_value = Mock(Hidden=False, RowHeight=15.0, MergeCells=False)
            sheet.Columns.return_value = Mock(Hidden=False, ColumnWidth=8.0)
            sheet.Cells.return_value = Mock(Text="12345", MergeCells=False)
            sheet.PageSetup.PrintArea = ""
            with patch.object(wf.page_breaks, "read_breaks", return_value=None), \
                    patch.object(wf.debug_log, "event") as event:
                wf._detect_sheet_results(sheet, Path(temp) / "book.xlsx", "S1", pdf, 1, confirm_locations=False)
        timing = [call.kwargs for call in event.call_args_list if call.args and call.args[0] == "⏱ 检测分步"]
        self.assertEqual(len(timing), 1)
        fields = timing[0]
        self.assertEqual(fields["表"], "S1")
        self.assertEqual(fields["格子数"], 2)
        self.assertGreaterEqual(fields["问显示文字次数"], 1)  # 数字格要问 Excel 显示成什么
        for key in ("读格子值", "合并格", "打印范围", "读PDF", "压线", "没显示全", "问显示文字秒", "合计秒"):
            self.assertIn(key, fields)


if __name__ == "__main__":
    unittest.main()
