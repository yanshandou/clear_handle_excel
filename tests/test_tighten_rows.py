"""收紧行高（用户 2026-11-03 要）：人工先给问题行一个很大的行高，程序量出要多高、收回多余的；
只收不加，收完整份复检，没过的行退回原来的行高"""
import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, fix_images, reporter, storage, workflow


class Area:
    def __init__(self, row, col, value, rows=1, size=11.0, wrap=True, height=None, need=0.0, orientation=-4128):
        self.Row, self.Column = row, col
        self.Rows, self.Columns = Mock(Count=rows), Mock(Count=1)
        self.value, self.WrapText, self.ShrinkToFit = value, wrap, False
        self.font, self.orientation, self.need = Mock(Size=size), orientation, need
        self.fixed_height = height

    def Cells(self, _r, _c):
        return Mock(Font=self.font, Value=self.value, Orientation=self.orientation)


class Row:
    def __init__(self, height):
        self.RowHeight, self.Hidden = height, False


class Sheet:
    def __init__(self, heights, areas, shapes=()):
        self.Name = "S1"
        self.rows = {r: Row(h) for r, h in heights.items()}
        self.areas = areas  # (行, 列) -> Area
        self.UsedRange = Mock(Column=1, Columns=Mock(Count=3))
        self.Shapes = list(shapes)

    def Rows(self, row):
        return self.rows[row]

    def Cells(self, row, col):
        area = self.areas.get((row, col)) or Area(row, col, None)
        if area.fixed_height is None:  # 合并格的高度 = 它占的各行加起来
            area.Height = sum(self.rows[r].RowHeight for r in range(area.Row, area.Row + area.Rows.Count))
        else:
            area.Height = area.fixed_height
        return Mock(MergeArea=area)

    def Range(self, addr):
        row, col = batch._cell_rc(addr)
        return self.Cells(row, col)


class TightenRowsTests(unittest.TestCase):
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
                             batch_review={"cards": {}, "rows": {}, "ignored": []}, checked_files={},
                             job={"kind": "batch_tighten", "options": {}, "completed": [], "pending_adjustment": None})
        self.book = Mock(ReadOnly=False)
        patch.object(batch, "open_workbook_checked", return_value=self.book).start()
        patch.object(batch, "normalize_workbook_view").start()
        # 没有真 Excel：量高度直接用格子上写好的 need
        patch.object(batch._Measurer, "height", lambda _self, area, wrap=True: area.need).start()
        self.addCleanup(patch.stopall)

    def use(self, sheet):
        self.sheet = sheet
        self.book.Sheets.return_value = sheet

    def card(self, addr, types=("border_clip",), **extra):
        cid = batch._key(self.path, "S1", addr)
        batch._data()["cards"][cid] = {"id": cid, "wb_path": self.path, "wb_name": "book.xlsx", "sheet_name": "S1",
                                       "cell_addr": addr, "excel_row": batch._cell_rc(addr)[0], "color": "yellow",
                                       "finalized": True, "dismissed": False, "error_types": list(types), **extra}
        return cid

    def test_shrinks_to_tallest_cell_in_the_row_plus_margin(self):
        # 第 5 行你给了 200pt：A5 要 40，同一行 B5（不是问题格）要 60、字号 14 → 收到 60 + 14 × 1.5 = 81
        self.use(Sheet({5: 200.0}, {(5, 1): Area(5, 1, "长文字", need=40.0),
                                     (5, 2): Area(5, 2, "别的格子", size=14.0, need=60.0)}))
        self.card("A5")
        done = batch._tighten_rows(Mock(), self.path, 1.5)
        self.assertEqual(self.sheet.rows[5].RowHeight, 81.0)
        self.assertEqual(done, {("S1", 5): {"before": 200.0, "after": 81.0, "by": "Excel"}})
        self.book.Save.assert_called_once()

    def test_only_shrinks_never_grows(self):
        self.use(Sheet({5: 30.0}, {(5, 1): Area(5, 1, "长文字", need=40.0)}))
        self.card("A5")
        self.assertEqual(batch._tighten_rows(Mock(), self.path, 1.5), {})
        self.assertEqual(self.sheet.rows[5].RowHeight, 30.0)
        self.book.Save.assert_not_called()

    def test_merged_cell_counts_height_of_its_other_rows(self):
        # A7:A8 合并、要 100pt：第 7 行你给了 200，第 8 行 20 → 第 7 行要 100 − 20 + 16.5 = 96.5，第 8 行不再收
        merged = Area(7, 1, "SQL", rows=2, need=100.0)
        self.use(Sheet({7: 200.0, 8: 20.0}, {(7, 1): merged, (8, 1): merged}))
        self.card("A7:A8")
        done = batch._tighten_rows(Mock(), self.path, 1.5)
        self.assertEqual(self.sheet.rows[7].RowHeight, 96.5)
        self.assertEqual(self.sheet.rows[8].RowHeight, 20.0)
        self.assertEqual(list(done), [("S1", 7)])

    def test_rows_with_pictures_or_rotated_text_and_other_problem_types_stay(self):
        shape = Mock(Type=13, TopLeftCell=Mock(Row=5), BottomRightCell=Mock(Row=6))
        comment = Mock(Type=4, TopLeftCell=Mock(Row=9), BottomRightCell=Mock(Row=9))
        self.use(Sheet({5: 200.0, 9: 200.0, 11: 200.0, 13: 200.0},
                       {(5, 1): Area(5, 1, "字", need=15.0), (9, 1): Area(9, 1, "字", need=15.0, orientation=90),
                        (11, 1): Area(11, 1, "字", need=15.0), (13, 1): Area(13, 1, "字", need=15.0)},
                       shapes=[shape, comment]))
        self.card("A5")
        self.card("A9")
        self.card("A11", types=("page_break_cut",))  # 分页切开不收
        self.card("A13", types=("border_clip",), marked=True, color="blue", finalized=False)  # 你标了蓝的不动
        self.assertEqual(batch._tighten_rows(Mock(), self.path, 1.5), {})
        self.assertEqual({r: self.sheet.rows[r].RowHeight for r in (5, 9, 11, 13)}, {5: 200.0, 9: 200.0, 11: 200.0, 13: 200.0})

    def test_unmeasurable_cell_is_estimated(self):
        # Excel 量不出（返回 0）：按单词换行估算，这里一行字 11 × 1.25 = 13.75，再加 16.5
        self.use(Sheet({5: 200.0}, {(5, 1): Area(5, 1, "給与", need=0.0, height=200.0)}))
        self.sheet.areas[(5, 1)].Width = 100.0
        self.card("A5")
        done = batch._tighten_rows(Mock(), self.path, 1.5)
        self.assertEqual(self.sheet.rows[5].RowHeight, 30.25)
        self.assertEqual(done[("S1", 5)]["by"], "估算")

    def test_rows_that_fail_recheck_go_back_to_your_height(self):
        self.use(Sheet({5: 200.0, 9: 150.0}, {(5, 1): Area(5, 1, "字", need=15.0), (9, 1): Area(9, 1, "字", need=15.0)}))
        self.card("A5")
        self.card("A9")
        session = Mock()
        rechecks = []

        def recheck(_excel, path, cards, passed_message="人工修改后复检通过", found_message=""):
            rechecks.append(passed_message)
            # 第一次复检：第 9 行收紧后还截断；退回后第二次复检通过
            self.wf.STATE["items"] = {} if len(rechecks) > 1 else {"x": {
                "id": "x", "wb_path": path, "sheet_name": "S1", "cell_addr": "A9", "excel_row": 9, "status": "manual_review"}}

        with patch.object(batch, "_recheck_file", side_effect=recheck), \
                patch.object(batch, "close_saved_workbooks", return_value={}):
            result = batch._tighten_file(session, self.path, 1.5)
        self.assertEqual(self.sheet.rows[5].RowHeight, 31.5)
        self.assertEqual(self.sheet.rows[9].RowHeight, 150.0)
        self.assertEqual(result, {"tightened": 1, "restored": 1, "saved_pt": 168.5})
        self.assertEqual(rechecks, [batch.TIGHTEN_PASSED_MESSAGE, "人工修改后复检通过"])
        self.assertEqual(self.book.Save.call_count, 2)

    def test_file_still_open_and_unsaved_asks_to_save(self):
        self.card("A5")
        with patch.object(batch, "close_saved_workbooks", return_value={Path(self.path): "文件还有没保存的修改"}), \
                self.assertRaisesRegex(ValueError, "保存并关闭后重试"):
            batch._tighten_file(Mock(), self.path, 1.5)


if __name__ == "__main__":
    unittest.main()
