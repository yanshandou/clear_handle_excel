import unittest
from types import SimpleNamespace
from unittest.mock import Mock

from core.autofix import validate_print_targets
from core.excel_ops import CellProbe


class PrintabilityTests(unittest.TestCase):
    def sheet(self):
        sheet = Mock()
        sheet.Rows.return_value = SimpleNamespace(Hidden=False, RowHeight=15)
        sheet.Columns.return_value = SimpleNamespace(Hidden=False, ColumnWidth=10)
        sheet.PageSetup.PrintArea = ""
        return sheet

    def test_hidden_and_zero_sized_rows_and_columns_are_excluded(self):
        for axis, field, value in (("Rows", "Hidden", True), ("Rows", "RowHeight", 0),
                                    ("Columns", "Hidden", True), ("Columns", "ColumnWidth", 0)):
            with self.subTest(axis=axis, field=field):
                sheet = self.sheet()
                setattr(getattr(sheet, axis).return_value, field, value)
                self.assertIs(CellProbe(sheet).printed(5, 2), False)

    def test_print_area_includes_only_the_selected_cells(self):
        sheet = self.sheet()
        sheet.PageSetup.PrintArea = "$B$2:$D$10"
        sheet.Range.return_value.Areas = [SimpleNamespace(Row=2, Column=2,
            Rows=SimpleNamespace(Count=9), Columns=SimpleNamespace(Count=3))]
        probe = CellProbe(sheet)
        self.assertIs(probe.printed(5, 2), True)
        self.assertIs(probe.printed(11, 2), False)
        self.assertIs(probe.printed(5, 1), False)

    def test_property_read_failure_is_unknown_instead_of_printable(self):
        for axis in ("Rows", "Columns"):
            sheet = self.sheet()
            getattr(sheet, axis).side_effect = RuntimeError("COM unavailable")
            self.assertIsNone(CellProbe(sheet).printed(5, 2))

    def test_print_area_read_failure_is_unknown(self):
        sheet = self.sheet()
        sheet.PageSetup.PrintArea = "$B$2:$D$10"
        sheet.Range.side_effect = RuntimeError("COM unavailable")
        probe = CellProbe(sheet)
        self.assertIsNone(probe.printed(5, 2))
        self.assertIsNone(probe.printed(6, 2))

    def test_explicit_hidden_row_is_excluded_even_when_column_cannot_be_read(self):
        sheet = self.sheet()
        sheet.Rows.return_value.Hidden = True
        sheet.Columns.side_effect = RuntimeError("COM unavailable")
        self.assertIs(CellProbe(sheet).printed(5, 2), False)

    def test_automatic_preview_refuses_hidden_or_unknown_print_targets(self):
        for unknown in (False, True):
            sheet = self.sheet()
            if unknown:
                sheet.Rows.side_effect = RuntimeError("COM unavailable")
            else:
                sheet.Rows.return_value.Hidden = True
            with self.subTest(unknown=unknown), self.assertRaisesRegex(ValueError, "未调整行高"):
                validate_print_targets(sheet, [{"cell_addr": "AC108"}])

    def test_automatic_preview_accepts_an_explicitly_visible_print_target(self):
        validate_print_targets(self.sheet(), [{"cell_addr": "AC108"}])


class CellProbeTextTests(unittest.TestCase):
    def test_display_text_is_read_once_per_cell(self):
        sheet = Mock()
        sheet.Cells.return_value = SimpleNamespace(Text="#REF!")
        probe = CellProbe(sheet)
        self.assertEqual([probe.text(2, 3) for _ in range(3)], ["#REF!"] * 3)
        self.assertEqual(sheet.Cells.call_count, 1)


if __name__ == "__main__":
    unittest.main()
