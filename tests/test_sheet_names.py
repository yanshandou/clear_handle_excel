import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from core import excel_ops


class XlsxSheetNameTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def test_reads_sheet_names_in_tab_order_without_openpyxl(self):
        wb = Workbook()
        wb.active.title = "表紙"
        wb.create_sheet("Main_20250101")
        wb.create_sheet("A&B <old>")
        wb.create_chartsheet("グラフ")
        wb.create_sheet("Main")
        wb["Main"].sheet_state = "hidden"
        path = self.folder / "book.xlsx"
        wb.save(path)

        names, backups = excel_ops.scan_workbook_sheets(path)
        self.assertEqual(names, ["表紙", "Main_20250101", "A&B <old>", "グラフ", "Main"])
        self.assertEqual(backups, ["Main_20250101"])

    def test_broken_file_still_raises(self):
        path = self.folder / "broken.xlsx"
        path.write_bytes(b"not a zip")
        with self.assertRaises(Exception):
            excel_ops.scan_workbook_sheets(path)


if __name__ == "__main__":
    unittest.main()
