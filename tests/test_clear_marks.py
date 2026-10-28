"""clear_marks_in_file：不依赖 STATE 记录，全局清除程序加的条件格式标记（CHE_MARK 规则）并还原橙黄标签。"""
import unittest
from unittest.mock import Mock, patch

from core import batch


class FakeTab:
    def __init__(self):
        self.ColorIndex = -4142  # 没颜色
        self.Color = 16777215


class FakeRule:
    def __init__(self, rules, formula, protected=False):
        self.rules = rules
        self.Formula1 = formula
        self.protected = protected

    def Delete(self):
        if self.protected:
            raise RuntimeError("例外が発生しました。")
        self.rules.remove(self)


class FakeFormatConditions:
    def __init__(self, rules):
        self.rules = rules

    @property
    def Count(self):
        return len(self.rules)

    def Item(self, index):
        return self.rules[index - 1]


class FakeSheet:
    def __init__(self, name="S1"):
        self.Name = name
        self.Tab = FakeTab()
        self.rules = []
        self.Cells = Mock()
        self.Cells.FormatConditions = FakeFormatConditions(self.rules)

    def add_rule(self, formula, protected=False):
        rule = FakeRule(self.rules, formula, protected)
        self.rules.append(rule)
        return rule


class FakeBook:
    def __init__(self, sheets, readonly=False):
        self.Worksheets = sheets
        self.ReadOnly = readonly
        self.saved = False

    def Save(self):
        self.saved = True

    def Close(self, SaveChanges=False):
        self.closed = True


class ClearMarksInFileTests(unittest.TestCase):
    def setUp(self):
        self.sheet = FakeSheet("S1")
        self.book = FakeBook([self.sheet])
        self.excel = Mock()

    def clear(self, report):
        with patch.object(batch, "_check_file"), patch.object(
                batch, "open_workbook_checked", return_value=self.book):
            batch.clear_marks_in_file(self.excel, "book.xlsx", report)
        return report

    def test_clears_che_mark_rules_and_restores_tab(self):
        self.sheet.add_rule('=N("CHE_MARK_deadbeef")=0')
        user = self.sheet.add_rule("=$C$1>0")
        self.sheet.Tab.ColorIndex = 46
        self.sheet.Tab.Color = batch._excel_color(batch.TAB_HIGHLIGHT)
        report = {"leftovers": 0, "leftover_failed": [], "tabs": []}
        self.clear(report)
        self.assertEqual(self.sheet.rules, [user])
        self.assertEqual(report["leftovers"], 1)
        self.assertEqual(report["tabs"], [{"sheet": "S1", "status": "ok", "reason": "已还原为无颜色"}])
        self.assertEqual(self.sheet.Tab.ColorIndex, batch.XL_COLOR_INDEX_NONE)
        self.assertTrue(self.book.saved)

    def test_no_marks_does_not_save(self):
        report = {"leftovers": 0, "leftover_failed": [], "tabs": []}
        self.clear(report)
        self.assertEqual(report["leftovers"], 0)
        self.assertEqual(report["tabs"], [])
        self.assertFalse(self.book.saved)

    def test_protected_rule_is_reported_but_does_not_save(self):
        self.sheet.add_rule('=N("CHE_MARK_aaaa")=0', protected=True)
        report = {"leftovers": 0, "leftover_failed": [], "tabs": []}
        self.clear(report)
        self.assertEqual(report["leftovers"], 0)
        self.assertTrue(report["leftover_failed"])
        self.assertFalse(self.book.saved)  # 一条都没删掉、也没还原标签：不保存

    def test_readonly_raises(self):
        self.book.ReadOnly = True
        with self.assertRaises(ValueError):
            self.clear({"leftovers": 0, "leftover_failed": [], "tabs": []})


if __name__ == "__main__":
    unittest.main()
