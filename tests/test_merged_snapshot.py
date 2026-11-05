import unittest
from types import SimpleNamespace
from unittest.mock import Mock

from core.excel_ops import snapshot_cells
from tests.test_hidden_text import detect


class MergedSnapshotTests(unittest.TestCase):
    def sheet(self, merges):
        sheet = Mock(Name="S1")
        sheet.Rows.side_effect = lambda row: SimpleNamespace(
            MergeCells=None if any(top <= row <= bottom for top, _, bottom, _ in merges) else False)
        def cell(row, col):
            for top, left, bottom, right in merges:
                if top <= row <= bottom and left <= col <= right:
                    return SimpleNamespace(MergeCells=True, MergeArea=SimpleNamespace(
                        Row=top, Column=left, Rows=SimpleNamespace(Count=bottom-top+1),
                        Columns=SimpleNamespace(Count=right-left+1)))
            return SimpleNamespace(MergeCells=False)
        sheet.Cells.side_effect = cell
        return sheet

    def test_pdf_text_in_c97_g97_is_not_reported_as_missing_four_more_times(self):
        text = "会議実施日"
        sheet = self.sheet([(97, 3, 97, 7)])
        snapshot = {"sheet": sheet, "start_r": 97, "start_c": 3, "rows_vals": [(text,) * 5]}
        def pdf(page):
            page.insert_text((50, 100), text, fontname="japan", fontsize=11)
        # 原来把 C～G 的同一个合并格当成五份文字，PDF 只有一份，额外四份全部误报。
        raw = [(97, col, text) for col in range(3, 8)]
        self.assertEqual(len(detect(raw, pdf)), 4)
        cells = snapshot_cells(snapshot)
        self.assertEqual(cells, [(97, 3, text)])
        self.assertEqual(detect(cells, pdf), [])
        sheet.Cells.assert_called_once_with(97, 3)

    def test_equal_text_in_separate_cells_still_requires_separate_pdf_occurrences(self):
        sheet = self.sheet([])
        cells = snapshot_cells({"sheet": sheet, "start_r": 1, "start_c": 1,
                                "rows_vals": [("meeting_date", "meeting_date")]})
        self.assertEqual(len(cells), 2)
        self.assertEqual(len(detect(cells, lambda page: page.insert_text((50, 100), "meeting_date"))), 1)
        sheet.Cells.assert_not_called()  # 整行没有合并格时不逐格查 COM。

    def test_multi_row_merge_has_only_one_anchor(self):
        sheet = self.sheet([(1, 1, 2, 2)])
        self.assertEqual(snapshot_cells({"sheet": sheet, "rows_vals": [("x", "x"), ("x", "x")]}), [(1, 1, "x")])
        sheet.Cells.assert_called_once_with(1, 1)

    def test_non_anchor_residual_content_is_not_independent_printed_text(self):
        sheet = self.sheet([(1, 1, 1, 2)])
        self.assertEqual(snapshot_cells({"sheet": sheet, "rows_vals": [("visible", "old hidden value")]}),
                         [(1, 1, "visible")])

    def test_unreadable_merge_properties_do_not_drop_real_text(self):
        sheet = self.sheet([(1, 1, 1, 2)])
        sheet.Cells.side_effect = RuntimeError("COM unavailable")
        self.assertEqual(snapshot_cells({"sheet": sheet, "rows_vals": [("one", "two")]}),
                         [(1, 1, "one"), (1, 2, "two")])
