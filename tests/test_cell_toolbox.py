"""查看切字的单元格工具箱：按选好的工具改格子、保存（用户 2026-11-08 要）"""
import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, cell_toolbox, fix_images, reporter, storage, workflow


class FakeArea:
    def __init__(self, row, rows=1, size=11.0, wrap=False, shrink=False, value="文字", width=60.0, sheet=None):
        self.Row, self.Rows, self.Columns = row, Mock(Count=rows), Mock(Count=1)
        self.WrapText, self.ShrinkToFit = wrap, shrink
        self.font = Mock(Size=size)
        self.value, self.Width, self.rows, self.sheet = value, width, rows, sheet

    @property
    def Height(self):
        return sum(self.sheet.rows[r].RowHeight for r in range(self.Row, self.Row + self.rows))

    def Cells(self, _r, _c):
        return Mock(Font=self.font, Value=self.value)


class FakeRow:
    def __init__(self, height):
        self.RowHeight = height


class FakeBreaks:
    def __init__(self):
        self.added = []
        self.Count = 0

    def __call__(self, i):
        return Mock(Location=Mock(Row=self.added[i - 1]))

    def Add(self, Before):
        self.added.append(Before.row)
        self.Count = len(self.added)


class FakeSheet:
    def __init__(self, name, heights):
        self.Name, self.areas = name, {}
        self.rows = {r: FakeRow(h) for r, h in heights.items()}
        self.HorizontalPageBreaks = FakeBreaks()
        self.PageSetup = Mock(Zoom=False, FitToPagesTall=False)
        self.PageSetup.Pages.Count = 2

    def area(self, addr, row, **kw):
        self.areas[addr] = FakeArea(row, sheet=self, **kw)
        return self.areas[addr]

    def Range(self, addr):
        return Mock(MergeArea=self.areas[addr])

    def Rows(self, row):
        return self.rows[row]

    def Cells(self, row, _col):
        return Mock(row=row)


class CellToolboxTests(unittest.TestCase):
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
                             job={"kind": "batch_fix", "options": {}, "completed": [], "pending_adjustment": None})
        self.sheet = FakeSheet("S1", {5: 18.75, 6: 13.5, 7: 13.5, 9: 18.75})
        self.book = Mock(ReadOnly=False)
        self.book.Sheets.return_value = self.sheet
        patch.object(cell_toolbox, "open_workbook_checked", return_value=self.book).start()
        patch.object(cell_toolbox, "normalize_workbook_view").start()
        patch.object(batch, "_find_sheet", return_value=self.sheet).start()
        self.addCleanup(patch.stopall)

    def add(self, addr, row, **extra):
        item = {"id": f"i{addr}", "seq": len(self.wf.STATE["items"]) + 1, "kind": "clip", "wb_path": self.path,
                "wb_name": "book.xlsx", "sheet_name": "S1", "cell_addr": addr, "excel_row": row, "text": "文字",
                "status": "pending", "location_confirmed": True, "page_index": 0, "bbox": [0, 0, 10, 10], **extra}
        self.wf.STATE["items"][item["id"]] = item
        return batch._key(self.path, "S1", addr)

    def hand_over(self):
        batch._hand_over(self.path)

    def test_grow_uses_the_original_height_so_ratios_do_not_stack(self):
        self.sheet.area("A5", 5)
        cid = self.add("A5", 5)
        self.hand_over()
        done = cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "grow", "ratio": 1.5}})
        # 短字一行放得下：原行高 18.75 + 11 × 1.5
        self.assertEqual(self.sheet.rows[5].RowHeight, 35.25)
        self.assertTrue(self.sheet.areas["A5"].WrapText)
        self.book.Save.assert_called_once()
        self.assertEqual(done["cells"][cid]["applied"], 35.25)
        # 没过，换成 ×2：还是按原来的 18.75 算，不在 35.25 上再加
        self.wf.STATE["checked_files"][self.path] = self.wf._file_version(Path(self.path))
        cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "grow", "ratio": 2.0}})
        self.assertEqual(self.sheet.rows[5].RowHeight, 40.75)

    def test_row_changed_by_hand_counts_from_the_new_height(self):
        self.sheet.area("A5", 5)
        cid = self.add("A5", 5)
        self.hand_over()
        cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "grow", "ratio": 1.5}})
        self.sheet.rows[5].RowHeight = 50.0  # 人在 Excel 里改过
        cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "grow", "ratio": 1.2}})
        self.assertEqual(self.sheet.rows[5].RowHeight, 63.2)

    def test_long_text_gets_enough_lines_and_merged_cell_grows_bottom_row(self):
        sql = "SELECT a.col1, b.col2 FROM table_a a JOIN table_b b ON a.id = b.id WHERE a.flag = 1"
        self.sheet.area("C6", 6, rows=2, value=sql, width=100.0)  # C6:C7 合并，共 27pt
        cid = self.add("C6", 6)
        self.hand_over()
        cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "grow", "ratio": 1.5}})
        # 估算 7 行 × 11 × 1.25 = 96.25pt，第 6 行占 13.5 → 第 7 行要 82.75，再加 16.5
        self.assertEqual(self.sheet.rows[7].RowHeight, 99.25)
        self.assertEqual(self.sheet.rows[6].RowHeight, 13.5)

    def test_same_row_takes_the_largest(self):
        self.sheet.area("A5", 5)
        self.sheet.area("B5", 5, size=14.0)
        a, c = self.add("A5", 5), self.add("B5", 5)
        self.hand_over()
        cell_toolbox.apply_tools(Mock(), self.path, {a: {"tool": "grow", "ratio": 1.2}, c: {"tool": "grow", "ratio": 1.5}})
        self.assertEqual(self.sheet.rows[5].RowHeight, 39.75)

    def test_shrink_turns_off_wrap(self):
        self.sheet.area("A5", 5, wrap=True)
        cid = self.add("A5", 5, kind="hidden", hidden_reason="hash")
        self.hand_over()
        cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "shrink"}})
        self.assertFalse(self.sheet.areas["A5"].WrapText)
        self.assertTrue(self.sheet.areas["A5"].ShrinkToFit)
        self.assertEqual(self.sheet.rows[5].RowHeight, 18.75)

    def test_page_break_above_the_cell_and_skips_fixed_page_height(self):
        self.sheet.area("A9", 9)
        cid = self.add("A9", 9, kind="pagebreak")
        self.hand_over()
        done = cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "break"}})
        self.assertEqual(self.sheet.HorizontalPageBreaks.added, [9])
        self.assertEqual(done["cells"][cid]["done"], "上面加分页线")
        self.sheet.PageSetup.FitToPagesTall = 2
        self.wf.STATE["checked_files"][self.path] = self.wf._file_version(Path(self.path))
        done = cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "break"}})
        self.assertIn("固定页数", done["cells"][cid]["skip"])
        self.assertFalse(done["changed"])

    def test_page_count_change_is_reported(self):
        self.sheet.area("A5", 5)
        cid = self.add("A5", 5)
        self.hand_over()
        counts = iter([2, 3])
        type(self.sheet.PageSetup.Pages).Count = property(lambda _self: next(counts))
        done = cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "grow", "ratio": 1.5}})
        self.assertEqual(done["pages"], [{"sheet": "S1", "before": 2, "after": 3}])

    def test_file_changed_after_check_is_refused(self):
        self.sheet.area("A5", 5)
        cid = self.add("A5", 5)
        self.hand_over()
        self.wf.STATE["checked_files"][self.path] = {"size": 1, "mtime_ns": 1}
        with self.assertRaisesRegex(ValueError, "请先复检"):
            cell_toolbox.apply_tools(Mock(), self.path, {cid: {"tool": "grow", "ratio": 1.5}})
        self.book.Save.assert_not_called()

    def test_validate_only_allows_tools_for_the_problem(self):
        self.sheet.area("A5", 5)
        cid = self.add("A5", 5)
        self.hand_over()
        self.assertEqual(cell_toolbox.validate({"tools": [{"card_id": cid, "tool": "grow", "ratio": 2}]}),
                         {cid: {"tool": "grow", "ratio": 2.0}})
        with self.assertRaisesRegex(ValueError, "不能用这个工具"):
            cell_toolbox.validate({"tools": [{"card_id": cid, "tool": "break"}]})
        with self.assertRaisesRegex(ValueError, "倍数"):
            cell_toolbox.validate({"tools": [{"card_id": cid, "tool": "grow", "ratio": 3}]})

    def test_failed_cells_say_which_tool_was_used(self):
        self.sheet.area("A5", 5)
        cid = self.add("A5", 5)
        self.hand_over()
        plan = {cid: {"tool": "grow", "ratio": 1.5}}
        with self.wf.STATE_LOCK:
            cell_toolbox._remember(self.path, plan, {cid: {"done": "行高 ×1.5"}})
            self.assertEqual(cell_toolbox._after_recheck(self.path, plan), (0, 1))
        card = batch._data()["cards"][cid]
        self.assertIn("行高 ×1.5", card["message"])
        snap = next(c for c in batch.snapshot()["cards"] if c["id"] == cid)
        self.assertEqual(snap["cell_fix"]["label"], "行高 ×1.5")


if __name__ == "__main__":
    unittest.main()
