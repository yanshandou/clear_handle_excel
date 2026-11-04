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
        # Excel 量不出（这里没有真 Excel）就估算：按单词换行，可用宽 96pt → 7 行 × 11 × 1.25 = 96.25pt；
        # 还差 77.5，再加 16.5
        self.assertEqual(self.sheet.rows[9].RowHeight, 112.75)
        self.assertEqual(done[batch._key(self.path, "S1", "E9")]["lines"], 7)
        self.assertEqual(done[batch._key(self.path, "S1", "E9")]["measured_by"], "估算")

    def test_excel_measures_merged_cell_on_a_temp_cell_of_the_same_width(self):
        """合并格 Excel 不会自动调整：搬到临时工作簿一个同宽的格子上量，量出来不够的补够再加倍数"""
        self.sheet.areas["E9"] = FakeArea(9, value="SELECT a FROM b", width=220.0, height=18.75)
        self.add("E9", 9)
        class Column:
            ColumnWidth = 10.0
            @property
            def Width(self):
                return self.ColumnWidth * 10.0  # 每个字宽 10pt
        col, row = Column(), Mock(RowHeight=60.0)
        tmp = Mock()
        tmp.Columns.return_value, tmp.Rows.return_value = col, row
        temp_book = Mock()
        temp_book.Worksheets.return_value = tmp
        excel = Mock()
        excel.Workbooks.Add.return_value = temp_book
        done = batch._grow_and_wrap(excel, self.path, 1.5)
        self.assertAlmostEqual(col.Width, 220.0)
        row.AutoFit.assert_called_once()
        # 量出 60pt，现在 18.75，补 41.25，再加 16.5
        self.assertEqual(self.sheet.rows[9].RowHeight, 76.5)
        self.assertEqual(done[batch._key(self.path, "S1", "E9")]["measured_by"], "Excel")
        temp_book.Close.assert_called_once_with(SaveChanges=False)

    def test_word_wrap_counts_lines_like_excel(self):
        self.assertEqual(batch._wrapped_lines("abc", 11, 100), 1)
        self.assertEqual(batch._wrapped_lines("給与" * 30, 11, 100), 7)
        # 按字数 3 行，按单词放不下的整个挪到下一行就成了 4 行
        sql = "T1.okyakusama_id = T3.okyakusama_id AND T2.eigyousho_id = T3.eigyousho_id AND T1.month_id >= T3.start_month"
        self.assertEqual(batch._wrapped_lines(sql, 11, 216), 4)

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

    def test_opened_file_waits_for_recheck_then_comes_back(self):
        self.add("A5", 5)
        batch._hand_over(self.path)
        batch.note_opened(Path(self.path))
        self.assertEqual(batch.snapshot()["opened_files"], [self.path])
        batch.note_opened(self.root / "other.xlsx")  # 没有问题的文件不记
        self.assertEqual(batch._opened_paths(), {self.path})
        batch._forget_opened(self.path)
        self.assertEqual(batch._opened_paths(), set())

    def test_grow_ratio_only_accepts_the_two_choices(self):
        self.assertEqual(batch.grow_ratio(None), 1.5)
        self.assertEqual(batch.grow_ratio(1.2), 1.2)
        self.assertEqual(batch.grow_ratio("1.5"), 1.5)
        for bad in (2, True, "abc", 0):
            with self.subTest(bad=bad), self.assertRaisesRegex(ValueError, "1.5 或 1.2"):
                batch.grow_ratio(bad)

    def columns(self, pages=(1, 1)):
        """给假表加上列（ColumnWidth 按 1 字 = 6pt）和打印页数（加宽前、后各读一次）"""
        class Column:
            def __init__(self):
                self.ColumnWidth = 10.0

            @property
            def Width(self):
                return self.ColumnWidth * 6.0

        cols = {}
        self.sheet.Columns = lambda col: cols.setdefault(col, Column())
        self.sheet.Shapes = []
        self.sheet.PageSetup = Mock()
        type(self.sheet.PageSetup).Pages = property(lambda _s, it=iter(pages): Mock(Count=next(it)))
        for area in self.sheet.areas.values():
            area.Column = 1
        return cols

    def test_one_english_word_wider_than_the_column_widens_the_column(self):
        """shikuchoson 只打印出 shikuchoso：Excel 量着一行放得下，加高没用，改成把列加宽 1.5 个字母（用户 2026-11-04 定）"""
        cols = self.columns()
        self.add("A5", 5, kind="hidden", hidden_reason="truncated", text="shikuchoson", visible_text="shikuchoso")
        self.add("E9", 9, kind="hidden", hidden_reason="truncated", text="長い説明文", visible_text="長い")
        done = batch._grow_and_wrap(Mock(), self.path, 1.5)
        a5 = batch._key(self.path, "S1", "A5")
        self.assertTrue(done[a5]["width_only"])
        self.assertNotIn("error", done[a5])
        self.assertEqual(self.sheet.rows[5].RowHeight, 18.75)  # 不加高
        self.assertGreater(self.sheet.rows[9].RowHeight, 18.75)  # 日文照样加高
        # 字号 11：1.5 × 11 × 0.55 = 9.075pt，原来 60pt
        self.assertAlmostEqual(cols[1].Width, 69.075)
        self.assertAlmostEqual(done[a5]["applied_width"], 69.07, places=1)
        self.book.Save.assert_called_once()
        batch._hand_over(self.path, done)
        card = self.card("A5")
        self.assertEqual(card["message"], batch.WIDENED_MESSAGE)
        self.assertTrue(card["width_only"])
        self.assertIn("列宽不够 · 已加宽", batch._card_category(card, batch._items_at(card))[1])
        # 只是一个单词但显示的不是开头一段（例如被边框压住）、或有空格：照样加高
        self.assertFalse(batch._width_only({"text": "shikuchoson"}, [{"kind": "clip", "visible_text": "shikuchoso"}]))
        self.assertFalse(batch._width_only({"text": "two words"}, [{"kind": "hidden", "visible_text": "two"}]))

    def test_widening_that_adds_a_page_goes_back(self):
        cols = self.columns(pages=(1, 2))
        self.add("A5", 5, kind="hidden", hidden_reason="truncated", text="shikuchoson", visible_text="shikucho")
        done = batch._grow_and_wrap(Mock(), self.path, 1.5)
        a5 = batch._key(self.path, "S1", "A5")
        self.assertEqual(cols[1].ColumnWidth, 10.0)  # 退回
        # 改成加一行行高：Excel 量着放得下（还差 0），也至少多给一行 11 × 1.25，再加 11 × 1.5
        self.assertEqual(done[a5]["widen_failed"], "加宽会多出一页，没加")
        self.assertNotIn("error", done[a5])
        self.assertEqual(self.sheet.rows[5].RowHeight, 18.75 + 13.75 + 16.5)
        batch._hand_over(self.path, done)
        self.assertEqual(self.card("A5")["message"], batch.GROWN_MESSAGE)
        self.assertNotIn("width_only", self.card("A5"))

    # 复检又查出切字（例如人工改完后行高不够）：和扫描一样开换行加高，再复检一遍（用户 2026-11-04 要）
    def session(self):
        return Mock(app=Mock(return_value=Mock()))

    def test_new_cut_cells_found_by_recheck_are_grown_then_rechecked(self):
        self.add("A5", 5)
        self.add("E9", 9)
        batch._hand_over(self.path)
        batch._data()["grow_ratio"] = 1.2
        tried = batch._grown_before_recheck(self.path)
        self.assertEqual(tried, set())
        with patch.object(batch, "_recheck_file") as recheck:
            batch._grow_after_recheck(self.session(), self.path, tried)
        self.assertEqual(self.card("A5")["grown"]["ratio"], 1.2)
        self.assertEqual(self.card("A5")["message"], batch.GROWN_MESSAGE)
        recheck.assert_called_once()
        _, kwargs = recheck.call_args
        self.assertEqual(kwargs["passed_message"], batch.RECHECK_GROWN_PASSED_MESSAGE)
        self.assertEqual(kwargs["found_message"], batch.GROWN_FAILED_MESSAGE)
        self.assertEqual({c["cell_addr"] for c in recheck.call_args[0][2]}, {"A5", "E9"})

    def test_cells_already_grown_before_recheck_are_left_to_people(self):
        self.add("A5", 5)
        batch._hand_over(self.path, batch._grow_and_wrap(Mock(), self.path, 1.5))
        height = self.sheet.rows[5].RowHeight
        self.book.Save.reset_mock()
        tried = batch._grown_before_recheck(self.path)
        self.assertEqual(tried, {batch._key(self.path, "S1", "A5")})
        with patch.object(batch, "_recheck_file") as recheck:
            batch._grow_after_recheck(self.session(), self.path, tried)
        self.assertEqual(self.sheet.rows[5].RowHeight, height)  # 不再加
        self.book.Save.assert_not_called()
        recheck.assert_not_called()

    def test_file_changed_meanwhile_is_not_touched(self):
        self.add("A5", 5)
        batch._hand_over(self.path)
        Path(self.path).write_bytes(b"edited again")
        with patch.object(batch, "_recheck_file") as recheck, patch.object(batch, "_restart_excel") as restart:
            batch._grow_after_recheck(self.session(), self.path, set())
        recheck.assert_not_called()
        restart.assert_not_called()
        self.assertNotIn("grown", self.card("A5"))


if __name__ == "__main__":
    unittest.main()
