import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from core import bl_links, workflow as wf


class FakeCollection:
    def __init__(self, items, delete=None):
        self.items = list(items)
        self.Delete = delete or Mock(side_effect=self.items.clear)

    @property
    def Count(self):
        return len(self.items)

    def __iter__(self):
        return iter(self.items)

    def __call__(self, index_or_name):
        if isinstance(index_or_name, int):
            return self.items[index_or_name - 1]
        return next(item for item in self.items if getattr(item, "Name", None) == index_or_name)


class FakeCell:
    def __init__(self, sheet, row, col, value="", merge_area=None):
        self.sheet = sheet
        self.Row = row
        self.Column = col
        self.Value2 = value
        self.Address = f"${chr(64 + col)}${row}"
        self.MergeCells = merge_area is not None
        self.MergeArea = merge_area or self
        self.Hyperlinks = FakeCollection([])
        self.ClearContents = Mock(side_effect=self._clear)

    def _clear(self):
        self.Value2 = ""

    def Cells(self, r, c):
        return self


class FakeMergeArea:
    def __init__(self, top_cell, address):
        self.top_cell = top_cell
        self.Address = address
        self.Row = top_cell.Row
        self.Column = top_cell.Column
        self.MergeCells = True
        self.MergeArea = self
        self.Hyperlinks = top_cell.Hyperlinks
        self.ClearContents = top_cell.ClearContents

    def Cells(self, r, c):
        return self.top_cell


class MatchingRuleTests(unittest.TestCase):
    """纯匹配规则：左边含大写 BL 或 Bl 开头的词、右边内容正好是 @，两个条件同时满足才命中。"""

    def test_left_contains_uppercase_bl(self):
        for value in ("BL", "BL001", "逻辑BL", "BL(論理)", "BL ", "BLK",
                      "Class:BlKaihoCal\nMethod:calcSeiykuushuuseiTouroku", "Bl", "bl", "blUtil", "Bl_01", "●Bl処理"):
            with self.subTest(value=value):
                self.assertTrue(bl_links.is_bl_left(value))
        for value in ("B L", "Black", "black", "Table", "public", "Class:TableUtil", "", None, "LABEL"):
            with self.subTest(value=value):
                self.assertFalse(bl_links.is_bl_left(value))

    def test_right_content_is_at(self):
        for value in ("@", " @ ", "@"):
            with self.subTest(value=value):
                self.assertTrue(bl_links.is_at_content(value))
        for value in ("", "@@", "xx@yy", None, "＠"):
            with self.subTest(value=value):
                self.assertFalse(bl_links.is_at_content(value))

    def test_scan_and_clear_support_merged_cells_and_internal_subaddress(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "book.xlsx"
            path.write_bytes(b"excel")

            cells = {}
            b2 = FakeCell(None, 2, 2, "BL001")
            b2_d2 = FakeMergeArea(b2, "$B$2:$D$2")
            for c in (2, 3, 4):
                cell = b2 if c == 2 else FakeCell(None, 2, c, None, merge_area=b2_d2)
                cell.MergeCells = True
                cell.MergeArea = b2_d2
                cells[(2, c)] = cell

            e2 = FakeCell(None, 2, 5, "@")
            e2_f2 = FakeMergeArea(e2, "$E$2:$F$2")
            e2.MergeCells = True
            e2.MergeArea = e2_f2
            link = SimpleNamespace(Range=e2, TextToDisplay="@", Address="", SubAddress="'BL001'!A1")
            e2.Hyperlinks = FakeCollection([link])
            e2_f2.Hyperlinks = e2.Hyperlinks
            cells[(2, 5)] = e2

            by_addr = {"E2:F2": e2_f2, "E2": e2}
            sheet = SimpleNamespace(
                Name="Main",
                Hyperlinks=FakeCollection([link]),
                Cells=lambda r, c: cells.get((r, c), FakeCell(None, r, c, "")),
                Range=lambda addr: by_addr[addr],
            )
            book = SimpleNamespace(
                Worksheets=FakeCollection([sheet]),
                ReadOnly=False,
                Save=Mock(),
                Close=Mock(),
            )
            with patch.object(bl_links, "open_workbook_checked", return_value=book), \
                    patch.object(bl_links, "workbook_locked", return_value=False):
                scanned = bl_links.scan_file(Mock(), path)
                self.assertEqual(len(scanned["hits"]), 1)
                hit = scanned["hits"][0]
                self.assertEqual(hit["addr"], "E2:F2")
                self.assertEqual(hit["anchor"], "E2")
                self.assertEqual(hit["left_addr"], "B2:D2")
                self.assertEqual(hit["left_value"], "BL001")
                self.assertEqual(hit["link_address"], "#'BL001'!A1")

                cleared = bl_links.clear_file(Mock(), {"path": str(path), "fingerprint": scanned["fingerprint"], "hits": scanned["hits"]})
                self.assertEqual(len(cleared), 1)
                self.assertEqual(cleared[0]["status"], "ok")
                e2.Hyperlinks.Delete.assert_called_once()
                e2.ClearContents.assert_called_once()
                book.Save.assert_called_once()


if __name__ == "__main__":
    unittest.main()
