import unittest
from types import SimpleNamespace
from unittest.mock import Mock

from core.common import is_inspection_exempt_sheet
from core.excel_ops import adjust_row_heights_against_clipping


class SheetExemptionTests(unittest.TestCase):
    def test_layout_substring_including_halfwidth_katakana_is_exempt(self):
        for name in ("レイアウト", "レイアウト(複数ページ)", "レイアウト（複数ページ）",
                     "新レイアウト", "帳票レイアウト確認用", "ﾚｲｱｳﾄ", "雛形", "帳票雛形", "雛形(複数ページ)"):
            with self.subTest(name=name):
                self.assertTrue(is_inspection_exempt_sheet(name))
        for name in ("概要", "帳票_出力内容", "JSONサンプル", "複数ページ出力", "Layout", ""):
            with self.subTest(name=name):
                self.assertFalse(is_inspection_exempt_sheet(name))

    def test_bulk_row_height_adjustment_skips_layout_sheets(self):
        sheets, rows = [], []
        for name in ("レイアウト", "レイアウト(複数ページ)", "雛形", "帳票雛形", "通常表"):
            row = SimpleNamespace(RowHeight=24.0)
            used_rows = Mock(return_value=row, Count=1)
            sheet = SimpleNamespace(Name=name, Visible=-1, Type=-4167, CustomProperties=[],
                                    UsedRange=SimpleNamespace(Rows=used_rows, Value=(("text",),)))
            sheets.append(sheet)
            rows.append(row)
        adjust_row_heights_against_clipping(SimpleNamespace(Sheets=sheets))
        self.assertEqual([row.RowHeight for row in rows], [24.0, 24.0, 24.0, 24.0, 30.0])


if __name__ == "__main__":
    unittest.main()
