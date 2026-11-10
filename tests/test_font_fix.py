"""统一字体：游ゴシック一家换成ＭＳ Ｐゴシック（图形每段文字、有内容的格子），别的字体不动、没改不保存。"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
from openpyxl.styles import Font

from core import font_fix


class FontObj:
    def __init__(self, name, far=None):
        self.Name, self.NameFarEast = name, far or name


class Runs:
    def __init__(self, fonts):
        self.fonts = fonts
        self.Count = len(fonts)


class TextRange:
    def __init__(self, fonts):
        self.fonts = fonts

    def Runs(self, index=None):
        if index is None:
            return Runs(self.fonts)
        return type("R", (), {"Font": self.fonts[index - 1]})()


class Shape:
    def __init__(self, name, fonts, kind=1, children=None):
        self.Name, self.Type = name, kind
        self.TextFrame2 = type("F", (), {"HasText": bool(fonts), "TextRange": TextRange(fonts)})()
        if children is not None:
            self.GroupItems = Shapes(children)


class Shapes:
    def __init__(self, items):
        self.items = items
        self.Count = len(items)

    def Item(self, index):
        return self.items[index - 1]


class Range:
    def __init__(self, sheet, addr):
        self.sheet, self.addr = sheet, addr
        self.Font = self

    def __setattr__(self, key, value):
        if key == "Name":
            self.sheet.set_fonts.append((self.addr, value))
        object.__setattr__(self, key, value)


class Sheet:
    def __init__(self, name, shapes=()):
        self.Name = name
        self.Shapes = Shapes(list(shapes))
        self.set_fonts = []

    def Range(self, addr):
        return Range(self, addr)


class Book:
    def __init__(self, sheets):
        self.Worksheets = sheets
        self.ReadOnly = False
        self.saved = 0

    def Save(self):
        self.saved += 1

    def Close(self, SaveChanges):
        pass


class FontFixTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def test_yu_family_detected(self):
        for name in ["游ゴシック", "游ゴシック 本文", "游ｺﾞｼｯｸ Light", "Yu Gothic UI", "YuGothic"]:
            self.assertTrue(font_fix.is_yu_gothic(name), name)
        for name in ["ＭＳ Ｐゴシック", "ＭＳ ゴシック", "メイリオ", "Arial", "游明朝", None]:
            self.assertFalse(font_fix.is_yu_gothic(name), name)

    def test_shape_runs_only_yu_changed(self):
        yu, ms, arial = FontObj("游ゴシック 本文"), FontObj("ＭＳ Ｐゴシック"), FontObj("Arial", "游ゴシック")
        shape = Shape("吹き出し 3", [yu, ms, arial])
        self.assertEqual(font_fix.fix_shape(shape), 2)
        self.assertEqual((yu.Name, yu.NameFarEast), (font_fix.STANDARD_FONT,) * 2)
        self.assertEqual(ms.Name, "ＭＳ Ｐゴシック")
        self.assertEqual(arial.NameFarEast, font_fix.STANDARD_FONT)
        self.assertEqual(font_fix.fix_shape(Shape("図 1", [])), 0)  # 图片没有文字

    def test_yu_cells_only_non_empty(self):
        path = self.root / "a.xlsx"
        book = openpyxl.Workbook()
        ws = book.active
        ws.title = "画面"
        ws["A1"] = "見出し"; ws["A1"].font = Font(name="游ゴシック")
        ws["B1"] = "正常"; ws["B1"].font = Font(name="ＭＳ Ｐゴシック")
        ws["C1"].font = Font(name="游ゴシック")  # 空格子不管
        book.save(path)
        self.assertEqual(font_fix.yu_cells(path), {"画面": ["A1"]})
        self.assertEqual(font_fix.yu_cells(self.root / "b.xls"), {})

    def test_address_chunks_under_limit(self):
        cells = [f"A{i}" for i in range(1, 200)]
        chunks = font_fix._address_chunks(cells)
        self.assertTrue(all(len(c) <= font_fix.MAX_ADDRESS_LEN for c in chunks))
        self.assertEqual(",".join(chunks).split(","), cells)

    def test_fix_workbook_groups_and_save(self):
        path = self.root / "c.xlsx"
        openpyxl.Workbook().save(path)
        inner = Shape("テキスト 2", [FontObj("游ゴシック")])
        group = Shape("グループ 1", [], kind=font_fix.MSO_GROUP, children=[inner])
        sheet = Sheet("画面", [group, Shape("吹き出し", [FontObj("游ゴシック 本文")])])
        book = Book([sheet])
        with patch.object(font_fix, "open_workbook_checked", return_value=book), \
                patch.object(font_fix, "yu_cells", return_value={"画面": ["A1", "B2"]}), \
                patch.object(font_fix, "workbook_locked", return_value=False):
            result = font_fix.fix_workbook(object(), path)
        self.assertEqual((result["shapes"], result["cells"], book.saved), (2, 2, 1))
        self.assertEqual(sheet.set_fonts, [("A1,B2", font_fix.STANDARD_FONT)])
        self.assertIn("2 个图形", font_fix.describe(result))

    def test_nothing_to_change_not_saved(self):
        path = self.root / "d.xlsx"
        openpyxl.Workbook().save(path)
        book = Book([Sheet("A", [Shape("x", [FontObj("ＭＳ Ｐゴシック")])])])
        with patch.object(font_fix, "open_workbook_checked", return_value=book), \
                patch.object(font_fix, "workbook_locked", return_value=False):
            result = font_fix.fix_workbook(object(), path)
        self.assertEqual((result["shapes"], book.saved), (0, 0))
        self.assertEqual(font_fix.describe(result), "字体都是对的")

    def test_locked_file_refused(self):
        with patch.object(font_fix, "workbook_locked", return_value=True), self.assertRaises(ValueError):
            font_fix.fix_workbook(object(), self.root / "e.xlsx")


if __name__ == "__main__":
    unittest.main()
