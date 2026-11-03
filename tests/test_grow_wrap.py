"""扫描后的简单处理：切字的格子开自动换行、所在行加高「字号 × 倍数」，保存一次（用户 2026-11-02 定，取代自动修复）"""
import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, fix_images, reporter, storage, workflow


class FakeArea:
    def __init__(self, row, rows=1, size=11.0, wrap=False, shrink=False, value="文字", width=60.0, height=None):
        self.Row, self.Rows, self.Columns = row, Mock(Count=rows), Mock(Count=1)
        self.WrapText, self.ShrinkToFit = wrap, shrink
        self.font = Mock(Size=size)
        self.value, self.Width, self.Height = value, width, 18.75 * rows if height is None else height

    def Cells(self, _r, _c):
        return Mock(Font=self.font, Value=self.value)


class FakeRow:
    def __init__(self, height):
        self.RowHeight = height


class FakeSheet:
    def __init__(self, name, areas, heights):
        self.Name, self.areas = name, areas
        self.rows = {r: FakeRow(h) for r, h in heights.items()}

    def Range(self, addr):
        return Mock(MergeArea=self.areas[addr])

    def Rows(self, row):
        return self.rows[row]


class GrowWrapTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for obj, name, value in (
            (storage, "DATABASE_PATH", self.root / "state.sqlite3"),
            (reporter, "RECORD_PATH", self.root / "record.md"),
            (fix_images, "IMAGES_DIR", self.root / "images"),
        ):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.wf = importlib.reload(workflow)
        self.path = str(self.root / "book.xlsx")
        Path(self.path).write_bytes(b"workbook")
        self.wf.STATE.update(folder=str(self.root), items={}, pages={},
                             batch_review={"cards": {}, "rows": {}, "ignored": []},
                             checked_files={self.path: self.wf._file_version(Path(self.path))},
                             job={"kind": "batch_scan", "options": {}, "completed": [], "pending_adjustment": None})
        self.sheet = FakeSheet("S1", {
            "A5": FakeArea(5, size=11.0),
            "C5": FakeArea(5, size=14.0, wrap=True),
            "D7": FakeArea(7, rows=2, size=10.0, shrink=True),  # D7:D8 合并、缩小字体填充
            "E9": FakeArea(9),
        }, {5: 18.75, 8: 13.5, 9: 18.75})
        self.book = Mock(ReadOnly=False)
        self.book.Sheets.return_value = self.sheet
        patch.object(batch, "open_workbook_checked", return_value=self.book).start()
        patch.object(batch, "normalize_workbook_view").start()
        self.addCleanup(patch.stopall)

    def add(self, addr, row, **extra):
        item = {"id": f"i{addr}", "seq": len(self.wf.STATE["items"]) + 1, "kind": "clip", "wb_path": self.path,
                "wb_name": "book.xlsx", "sheet_name": "S1", "cell_addr": addr, "excel_row": row, "text": "文字",
                "status": "pending", "location_confirmed": True, "page_index": 0, "bbox": [0, 0, 10, 10], **extra}
        self.wf.STATE["items"][item["id"]] = item
        return item

    def card(self, addr):
        return batch._data()["cards"][batch._key(self.path, "S1", addr)]

    def test_wraps_cells_and_raises_each_row_once_by_largest_font(self):
        self.add("A5", 5)
        self.add("C5", 5, kind="hidden", hidden_reason="truncated")
        self.add("D7", 7)
        done = batch._grow_and_wrap(Mock(), self.path, 1.5)
        self.assertTrue(self.sheet.areas["A5"].WrapText)
        # 同一行两格只加一次：按大的字号 14 × 1.5 = 21
        self.assertEqual(self.sheet.rows[5].RowHeight, 39.75)
        # 合并单元格加在最下面一行；缩小字体填充的不开换行
        self.assertEqual(self.sheet.rows[8].RowHeight, 28.5)
        self.assertFalse(self.sheet.areas["D7"].WrapText)
        self.book.Save.assert_called_once()
        self.book.Close.assert_called_once_with(SaveChanges=False)
        self.assertEqual(done[batch._key(self.path, "S1", "A5")],
                         {"wrap": True, "row": 5, "font": 11.0, "merged": False, "sheet": "S1", "base": 18.75, "applied": 39.75, "ratio": 1.5})
        self.assertFalse(done[batch._key(self.path, "S1", "C5")]["wrap"])  # 本来就开着

    def test_long_text_gets_enough_lines_before_the_extra(self):
        """很长的文字（例如 SQL）：按格子宽度估算换行后要几行，先补够高度，再加「字号 × 倍数」"""
        sql = "SELECT a.col1, b.col2 FROM table_a a JOIN table_b b ON a.id = b.id WHERE a.flag = 1"  # 84 个半角
        self.sheet.areas["E9"] = FakeArea(9, value=sql, width=100.0, height=18.75)
        self.add("E9", 9)
        done = batch._grow_and_wrap(Mock(), self.path, 1.5)
        # 84 × 11 × 0.55 = 508.2pt，可用宽 96pt → 6 行 × 11 × 1.25 = 82.5pt；还差 63.75，再加 16.5
        self.assertEqual(self.sheet.rows[9].RowHeight, 99.0)
        self.assertEqual(done[batch._key(self.path, "S1", "E9")]["lines"], 6)

    def test_short_text_that_already_fits_only_gets_the_extra(self):
        self.sheet.areas["E9"] = FakeArea(9, value="給与", width=100.0)
        self.add("E9", 9)
        batch._grow_and_wrap(Mock(), self.path, 1.5)
        self.assertEqual(self.sheet.rows[9].RowHeight, 35.25)

    def test_line_breaks_and_merged_height_count(self):
        # 合并两行（共 37.5pt），文字有 4 段：4 × 13.75 = 55，还差 17.5，加在最下面一行
        self.sheet.areas["E9"] = FakeArea(8, rows=2, value="一\n二\n三\n四", width=100.0, height=37.5)
        self.add("E9", 8)
        batch._grow_and_wrap(Mock(), self.path, 1.5)
        self.assertEqual(self.sheet.rows[9].RowHeight, 18.75 + 17.5 + 16.5)

    def test_ratio_one_point_two_and_row_height_cap(self):
        self.add("E9", 9)
        self.sheet.rows[9].RowHeight = 400.0
        batch._grow_and_wrap(Mock(), self.path, 1.2)
        self.assertEqual(self.sheet.rows[9].RowHeight, 409.0)

    def test_other_problems_and_unlocated_cells_are_only_marked(self):
        self.add("A5", 5, kind="pagebreak")
        self.add("E9", 9, location_confirmed=False)
        self.assertEqual(batch._grow_and_wrap(Mock(), self.path, 1.5), {})
        batch.open_workbook_checked.assert_not_called()
        self.assertEqual(self.sheet.rows[5].RowHeight, 18.75)

    def test_file_changed_after_scan_is_not_touched(self):
        self.add("A5", 5)
        Path(self.path).write_bytes(b"edited by someone")
        with self.assertRaisesRegex(ValueError, "扫描后文件已修改"):
            batch._grow_and_wrap(Mock(), self.path, 1.5)
        batch.open_workbook_checked.assert_not_called()

    def test_hand_over_writes_what_was_done_on_the_card(self):
        self.add("A5", 5)
        self.add("E9", 9, kind="pagebreak")
        batch._hand_over(self.path, batch._grow_and_wrap(Mock(), self.path, 1.5))
        grown, other = self.card("A5"), self.card("E9")
        self.assertEqual((grown["color"], grown["finalized"], grown["message"]), ("yellow", True, batch.GROWN_MESSAGE))
        self.assertEqual(grown["grown"]["applied"], 35.25)
        self.assertEqual(other["message"], batch.SCAN_ONLY_MESSAGE)
        self.assertNotIn("grown", other)
        # 下次扫描没再加高的，旧的记录去掉
        batch._hand_over(self.path)
        self.assertNotIn("grown", self.card("A5"))

    def test_category_groups_by_type_merge_grow_and_missing_lines(self):
        self.add("A5", 5, kind="hidden", hidden_reason="truncated", missing_lines=1)
        self.add("E9", 9, kind="pagebreak")
        done = batch._grow_and_wrap(Mock(), self.path, 1.5)
        batch._hand_over(self.path, done)
        snap = {c["cell_addr"]: c for c in batch.snapshot()["cards"]}
        key, label = snap["A5"]["category"]
        self.assertEqual(label, "文字截断 · 单格 · 已加高 · 缺 1 行")
        self.assertEqual(key, "text_truncated|single|grown|lines1")
        self.assertEqual(snap["A5"]["cell_text"], "文字")
        self.assertNotIn("已加高", snap["E9"]["category"][1])
        merged = batch._card_category({"error_types": ["border_clip"], "cell_addr": "B2", "grown": {"merged": True}}, [])
        self.assertEqual(merged[1], "边框压线 · 合并格 · 已加高")
        many = batch._card_category({"error_types": ["text_truncated"], "cell_addr": "B2:C3"}, [{"missing_lines": 4}])
        self.assertEqual(many[1], "文字截断 · 合并格 · 没加高 · 缺 3 行以上")

    def test_grow_ratio_only_accepts_the_two_choices(self):
        self.assertEqual(batch.grow_ratio(None), 1.5)
        self.assertEqual(batch.grow_ratio(1.2), 1.2)
        self.assertEqual(batch.grow_ratio("1.5"), 1.5)
        for bad in (2, True, "abc", 0):
            with self.subTest(bad=bad), self.assertRaisesRegex(ValueError, "1.5 或 1.2"):
                batch.grow_ratio(bad)


if __name__ == "__main__":
    unittest.main()
