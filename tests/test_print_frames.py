import unittest

from core import print_frames


class Obj:
    def __init__(self, **kw):
        self.__dict__.update(kw)


def shape(name, w, h, fill=0, line=-1, kind=1, auto=1, visible=-1):
    return Obj(Name=name, Type=kind, AutoShapeType=auto, Width=w, Height=h, Visible=visible,
               Fill=Obj(Visible=fill), Line=Obj(Visible=line))


class Shapes:
    def __init__(self, items):
        self.items = items
        self.Count = len(items)

    def Item(self, i):
        return self.items[i - 1]


class Sheet:
    def __init__(self, items, w=500, h=2000):
        self.Name = "詳細"
        self.Shapes = Shapes(items)
        self.UsedRange = Obj(Width=w, Height=h)
        self.seen = None

    def ExportAsFixedFormat(self):
        self.seen = [s.Visible for s in self.Shapes.items]


class PrintFramesTest(unittest.TestCase):
    def test_only_big_hollow_rectangle_hidden_and_restored(self):
        frame = shape("枠", 520, 1500)
        small = shape("赤枠", 120, 60)          # 标注小框
        filled = shape("塗り", 520, 1500, fill=-1)
        line = shape("線", 520, 1500, kind=9)    # 线条不算
        sheet = Sheet([frame, small, filled, line])
        with print_frames.hidden_frames(sheet) as n:
            sheet.ExportAsFixedFormat()
        self.assertEqual(n, 1)
        self.assertEqual(sheet.seen, [False, -1, -1, -1])
        self.assertEqual(frame.Visible, True)

    def test_restored_when_export_fails(self):
        frame = shape("枠", 500, 800, auto=5)
        sheet = Sheet([frame])
        with self.assertRaises(RuntimeError):
            with print_frames.hidden_frames(sheet):
                self.assertFalse(frame.Visible)
                raise RuntimeError("COM")
        self.assertTrue(frame.Visible)

    def test_unreadable_sheet_does_nothing(self):
        sheet = Obj(Name="グラフ")
        with print_frames.hidden_frames(sheet) as n:
            self.assertEqual(n, 0)

    def test_real_export_function_runs(self):
        """workflow 真正的导出函数要能跑到 ExportAsFixedFormat（#217 漏了 import，所有导出都报 NameError）"""
        from pathlib import Path
        from unittest.mock import patch
        from core import workflow
        frame = shape("枠", 500, 800)
        calls = []

        class ExportSheet(Sheet):
            def ExportAsFixedFormat(self, **kw):
                calls.append((kw["Filename"], frame.Visible))

        sheet = ExportSheet([frame])
        with patch.object(workflow.page_fill, "force_one_page_wide"):
            workflow._export_sheet_pdf_raw(sheet, Path("out.pdf"))
        self.assertEqual(len(calls), 1)
        self.assertFalse(calls[0][1])
        self.assertTrue(frame.Visible)


if __name__ == "__main__":
    unittest.main()
