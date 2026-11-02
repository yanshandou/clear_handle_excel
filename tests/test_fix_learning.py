"""按分类示范学习修复：分类、对比改前改后、总结规律、按规律算新值"""
import unittest
from unittest.mock import Mock

from core import fix_learning as fl


def props(heights=(18.75,), widths=(8.38,), wrap=True, shrink=False, font=11.0, top=5, left=2, brk=False):
    return {"top": top, "left": left, "heights": list(heights), "widths": list(widths),
            "wrap": wrap, "shrink": shrink, "font": font, "break": brk}


def demo(cid, before, after, lines=None, text="文字"):
    return {"id": cid, "label": f"S!{cid}", "change": fl.changes(before, after), "font": before["font"],
            "lines": lines, "text_width": fl.text_width(text)}


class CategoryTests(unittest.TestCase):
    def test_splits_by_type_merge_wrap_and_missing_lines(self):
        card = {"error_types": ["text_truncated"], "cell_addr": "B5"}
        key, label = fl.category(card, [{"missing_lines": 2}], props(wrap=False))
        self.assertEqual(key, "text_truncated|single|nowrap|lines2")
        self.assertEqual(label, "文字截断 · 单格 · 没开换行 · 缺 2 行")
        key, label = fl.category(card, [{"missing_lines": 5}], props(heights=(15, 15)))
        self.assertEqual(label, "文字截断 · 合并格 · 已开换行 · 缺 3 行以上")

    def test_picks_first_type_and_skips_wrap_for_hash(self):
        card = {"error_types": ["hash_overflow", "border_clip"], "cell_addr": "B5"}
        self.assertEqual(fl.category(card, [], None)[0], "border_clip|single|wrap?")
        self.assertEqual(fl.category({"error_types": ["hash_overflow"], "cell_addr": "C3"}, [], None)[1],
                         "### 列宽不足 · 单格")


class ChangesTests(unittest.TestCase):
    def test_small_row_rounding_is_not_a_change(self):
        self.assertTrue(fl.same(props(heights=(18.75,)), props(heights=(18.8,))))
        self.assertEqual(fl.changes(props(), props(heights=(30,), wrap=False)),
                         {"row": (18.75, 30), "wrap": (True, False)})

    def test_merge_change_is_reported(self):
        self.assertIn("merge", fl.changes(props(), props(heights=(15, 15))))


class LearnTests(unittest.TestCase):
    def test_row_height_by_font_size_when_fonts_differ(self):
        demos = [demo("A1", props(font=10), props(heights=(18.75 + 15,), font=10)),
                 demo("A2", props(font=12), props(heights=(18.75 + 18,), font=12)),
                 demo("A3", props(font=14), props(heights=(18.75 + 21,), font=14))]
        rule = fl.learn(demos)
        self.assertTrue(rule["ok"])
        self.assertEqual(rule["steps"], [{"prop": "row", "model": "font", "value": 1.5,
                                          "text": "行高加「字号 × 1.5」（合并格加在最下面一行）"}])
        self.assertEqual(rule["conflicts"], [])

    def test_row_height_by_missing_lines(self):
        demos = [demo("A1", props(), props(heights=(18.75 + 13.2,)), lines=1),
                 demo("A2", props(), props(heights=(18.75 + 39.6,)), lines=3)]
        rule = fl.learn(demos)
        self.assertEqual((rule["steps"][0]["model"], rule["steps"][0]["value"]), ("lines", 1.2))

    def test_fixed_add_when_constant(self):
        demos = [demo("A1", props(font=10), props(heights=(28.75,), font=10)),
                 demo("A2", props(font=14), props(heights=(28.75,), font=14))]
        rule = fl.learn(demos)
        self.assertEqual((rule["steps"][0]["model"], rule["steps"][0]["value"]), ("add", 10))

    def test_minority_change_and_outlier_are_listed(self):
        demos = [demo("A1", props(), props(heights=(35,), wrap=False)),
                 demo("A2", props(), props(heights=(35,))),
                 demo("A3", props(), props(heights=(35,))),
                 demo("A4", props(), props(heights=(80,)))]
        rule = fl.learn(demos)
        self.assertEqual([s["prop"] for s in rule["steps"]], ["row"])
        reasons = {c["id"]: c["reason"] for c in rule["conflicts"]}
        self.assertIn("只有少数示范改了", reasons["A1"])
        self.assertIn("差得多", reasons["A4"])

    def test_column_width_by_text_length(self):
        demos = [demo("A1", props(widths=(8,)), props(widths=(12,)), text="1234567890"),
                 demo("A2", props(widths=(8,)), props(widths=(24,)), text="12345678901234567890")]
        rule = fl.learn(demos)
        self.assertEqual((rule["steps"][0]["model"], rule["steps"][0]["value"]), ("text", 1.2))
        target = fl.plan(rule, props(widths=(5,)), lines=None, text="12345")
        self.assertEqual(target, {"col": 6.0})
        self.assertEqual(fl.plan(rule, props(widths=(30,)), lines=None, text="12345"), {})

    def test_nothing_changed(self):
        rule = fl.learn([])
        self.assertFalse(rule["ok"])


class PlanTests(unittest.TestCase):
    def test_applies_wrap_and_rows_to_bottom_row(self):
        rule = {"steps": [{"prop": "wrap", "value": True}, {"prop": "row", "model": "lines", "value": 1.2}]}
        target = fl.plan(rule, props(heights=(15, 15), wrap=False, font=10), lines=2, text="x")
        self.assertEqual(target, {"wrap": True, "row": 39.0})

    def test_caps_row_height(self):
        rule = {"steps": [{"prop": "row", "model": "add", "value": 500}]}
        self.assertEqual(fl.plan(rule, props(), lines=None, text=""), {"row": 409.0})


class ComTests(unittest.TestCase):
    def test_read_and_write_merged_area(self):
        rows = {5: Mock(RowHeight=15.0, PageBreak=-4142), 6: Mock(RowHeight=15.0)}
        cols = {2: Mock(ColumnWidth=8.38), 3: Mock(ColumnWidth=8.38)}
        area = Mock(Row=5, Column=2, Rows=Mock(Count=2), Columns=Mock(Count=2), WrapText=False, ShrinkToFit=False)
        area.Cells.return_value = Mock(Font=Mock(Size=11.0))
        sheet = Mock()
        sheet.Range.return_value = Mock(MergeArea=area)
        sheet.Rows.side_effect = rows.__getitem__
        sheet.Columns.side_effect = cols.__getitem__
        now = fl.read_props(sheet, "B5:C6")
        sheet.Range.assert_called_with("B5")
        self.assertEqual(now, props(heights=(15, 15), widths=(8.38, 8.38), wrap=False, top=5, left=2))
        fl.write_props(sheet, "B5", now, {"wrap": True, "row": 30.0, "col": 10.0})
        self.assertTrue(area.WrapText)
        self.assertEqual(rows[6].RowHeight, 30.0)
        self.assertEqual(rows[5].RowHeight, 15.0)
        self.assertEqual(cols[3].ColumnWidth, 10.0)


if __name__ == "__main__":
    unittest.main()


class FakeCell:
    """一格（不合并）：行高、列宽存在共用的表上"""
    def __init__(self, sheet, row, col):
        self.sheet, self.Row, self.Column = sheet, row, col
        self.Rows, self.Columns = Mock(Count=1), Mock(Count=1)
        self.WrapText, self.ShrinkToFit = True, False
        self.font = Mock(Size=11.0)
        self.Font = self.font

    @property
    def MergeArea(self):
        return self

    def Cells(self, _r, _c):
        return Mock(Font=self.font)


class FakeSheet:
    def __init__(self, name):
        self.Name = name
        self.rows, self.cols, self.cells = {}, {}, {}

    def Range(self, addr):
        import re
        m = re.fullmatch(r"([A-Z]+)(\d+)", addr)
        col = ord(m[1]) - 64
        row = int(m[2])
        self.rows.setdefault(row, Mock(RowHeight=18.75, PageBreak=-4142))
        self.cols.setdefault(col, Mock(ColumnWidth=8.38))
        return self.cells.setdefault(addr, FakeCell(self, row, col))

    def Rows(self, r):
        return self.rows.setdefault(r, Mock(RowHeight=18.75, PageBreak=-4142))

    def Columns(self, c):
        return self.cols.setdefault(c, Mock(ColumnWidth=8.38))


class LearnJobTests(unittest.TestCase):
    """学一下：和原样不一样的格子当示范；按规律修：只改没动过的格子，保存"""

    def setUp(self):
        import importlib, tempfile
        from pathlib import Path
        from unittest.mock import patch
        from core import batch, fix_images, reporter, storage, workflow
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        for obj, name, value in ((storage, "DATABASE_PATH", root / "s.sqlite3"), (reporter, "RECORD_PATH", root / "r.md"),
                                 (fix_images, "IMAGES_DIR", root / "img")):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.wf = importlib.reload(workflow)
        self.batch = batch
        self.path = str(root / "book.xlsx")
        Path(self.path).write_bytes(b"x")
        self.wf.STATE.update(folder=str(root), items={}, pages={},
                             batch_review={"cards": {}, "rows": {}, "ignored": []}, checked_files={},
                             job={"kind": "batch_learn", "options": {}, "completed": []})
        self.sheet = FakeSheet("S1")
        self.book = Mock(ReadOnly=False)
        self.book.Sheets.return_value = self.sheet
        for name in ("open_workbook_checked",):
            p = patch.object(batch, name, return_value=self.book)
            p.start()
            self.addCleanup(p.stop)
        p = patch.object(batch, "normalize_workbook_view")
        p.start()
        self.addCleanup(p.stop)
        self.cards = []
        for i, (addr, size) in enumerate((("A1", 10), ("A2", 12), ("A3", 14), ("A4", 11), ("A5", 16)), 1):
            item = {"id": f"i{i}", "seq": i, "kind": "clip", "wb_path": self.path, "wb_name": "book.xlsx",
                    "sheet_name": "S1", "cell_addr": addr, "excel_row": i, "text": "文字", "status": "manual_review",
                    "error_type": "border_clip"}
            self.wf.STATE["items"][item["id"]] = item
            card = batch._card_for([item])
            card.update(color="yellow", finalized=True)
            self.sheet.Range(addr).font.Size = size
            self.cards.append(card)
        batch._capture_baselines(Mock(), self.path, only_missing=False)

    def test_learns_from_changed_cells_then_applies_to_the_rest(self):
        # 你示范了三格：行高加「字号 × 1.5」
        for card in self.cards[:3]:
            row = int(card["cell_addr"][1:])
            self.sheet.rows[row].RowHeight += self.sheet.Range(card["cell_addr"]).font.Size * 1.5
        job = self.wf.STATE["job"]
        key = self.batch._card_category(self.cards[0], [])[0]
        job["options"] = {"category": key, "card_ids": [c["id"] for c in self.cards]}
        self.batch._learn_file(Mock(), self.path, self.cards, job)
        self.batch._finish_learn(job)
        rule = self.batch._rules()[key]
        self.assertTrue(rule["ok"])
        self.assertEqual(rule["text"], "行高加「字号 × 1.5」（合并格加在最下面一行）")
        self.assertEqual(sorted(rule["demo_ids"]), sorted(c["id"] for c in self.cards[:3]))
        changed = self.batch._apply_file(Mock(), self.path, self.cards, rule)
        self.assertEqual(changed, 2)
        self.assertEqual(self.sheet.rows[4].RowHeight, 18.75 + 16.5)
        self.assertEqual(self.sheet.rows[5].RowHeight, 18.75 + 24)
        self.assertEqual(self.sheet.rows[1].RowHeight, 18.75 + 15)  # 示范的不动
        self.book.Save.assert_called_once()
        snap = self.batch.snapshot()
        cat = next(c for c in snap["categories"] if c["key"] == key)
        states = {c["cell_addr"]: c["state"] for c in cat["cards"]}
        self.assertEqual(states, {"A1": "demo", "A2": "demo", "A3": "demo", "A4": "applied", "A5": "applied"})
        self.assertEqual(cat["rule"]["text"], rule["text"])
