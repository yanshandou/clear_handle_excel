import unittest

from core import image_restore as ir, shape_geometry as sg


def s(name, kind, left, top, w, h):
    return {"name": name, "type": kind, "left": left, "top": top, "width": w, "height": h}


class Shape:
    def __init__(self, d, lock=-1):
        self.Name, self.Type = d["name"], d["type"]
        self.Left, self.Top, self.Width, self.Height = d["left"], d["top"], d["width"], d["height"]
        self.LockAspectRatio, self.Visible = lock, True


class Shapes:
    def __init__(self, items):
        self.order = list(items)
        self.items = {i.Name: i for i in items}
        self.Count = len(items)

    def Item(self, key):
        return self.order[key - 1] if isinstance(key, int) else self.items[key]


class Sheet:
    def __init__(self, items):
        self.Shapes = Shapes(items)


class PlanTest(unittest.TestCase):
    ref = [s("図 1", 13, 30, 40, 330, 310), s("正方形 2", 1, 230, 300, 50, 33), s("矢印 3", 1, 500, 40, 50, 20)]

    def test_stretched_picture_and_shifted_box(self):
        # 截图中间的行加高 60 磅：截图被拉长，红框只往下挪了 80（挂在下面的格子上）
        cur = [s("図 1", 13, 30, 100, 330, 370), s("正方形 2", 1, 230, 380, 50, 33), s("矢印 3", 1, 500, 100, 50, 20)]
        plan = {c["name"]: c for c in sg.plan_sheet(cur, self.ref)}
        self.assertEqual(set(plan), {"図 1", "正方形 2"})  # 只挪了位置的箭头不动
        self.assertEqual(plan["図 1"]["after"], {"left": 30, "top": 100, "width": 330, "height": 310})
        self.assertEqual(plan["正方形 2"]["kind"], "overlay")
        self.assertEqual(plan["正方形 2"]["after"], {"left": 230, "top": 360, "width": 50, "height": 33})

    def test_unchanged_sheet_has_no_plan(self):
        cur = [dict(x, top=x["top"] + 20) for x in self.ref]  # 上面的行加高，全体一起往下挪
        self.assertEqual(sg.plan_sheet(cur, self.ref), [])

    def test_duplicate_or_type_mismatch_skipped(self):
        cur = [s("図 1", 13, 30, 40, 330, 400), s("図 1", 13, 0, 0, 10, 10), s("正方形 2", 9, 0, 0, 5, 5)]
        self.assertEqual(sg.plan_sheet(cur, self.ref), [])

    def test_apply_places_overlay_after_host_and_unlocks_aspect(self):
        cur = [s("図 1", 13, 30, 100, 330, 370), s("正方形 2", 1, 230, 380, 50, 33)]
        changes = sg.plan_sheet(cur, self.ref)
        sheet = Sheet([Shape(cur[0]), Shape(cur[1], lock=0)])
        self.assertEqual(sg.apply_sheet(sheet, changes), [])
        pic, box = sheet.Shapes.Item("図 1"), sheet.Shapes.Item("正方形 2")
        self.assertEqual((pic.Width, pic.Height, pic.LockAspectRatio), (330, 310, -1))
        self.assertEqual((box.Left, box.Top), (230, 360))

    def test_pair_sheets_by_name_then_code(self):
        cur = [{"name": "画面", "code": "Sheet1"}, {"name": "新名", "code": "Sheet2"}, {"name": "无", "code": ""}]
        old = [{"name": "画面", "code": "Sheet9"}, {"name": "旧名", "code": "Sheet2"}]
        pairs = [(a["name"], b["name"]) for a, b in ir.pair_book_sheets(cur, old)]
        self.assertEqual(pairs, [("画面", "画面"), ("新名", "旧名")])


class Worksheets:
    """像 COM 的集合：能遍历，也能 Worksheets("名字")"""
    def __init__(self, sheets):
        self.sheets = sheets

    def __call__(self, name):
        return self.sheets[name]

    def __iter__(self):
        return iter(self.sheets.values())


class Book:
    def __init__(self, sheets):
        self.Worksheets = Worksheets(sheets)


class GuardSheet(Sheet):
    def __init__(self, name, items):
        super().__init__(items)
        self.Name = name


class GuardTest(unittest.TestCase):
    def test_row_grow_inside_picture_restored_before_save(self):
        pic, box = Shape(s("図 1", 13, 30, 40, 330, 310)), Shape(s("正方形 2", 1, 230, 300, 50, 33), lock=0)
        sheet = GuardSheet("手順", [pic, box])
        empty = GuardSheet("表紙", [])
        book = Book({"手順": sheet, "表紙": empty})
        snap = sg.snapshot_book(book)
        self.assertEqual(list(snap), ["手順"])
        pic.Height, box.Top = 370, 380  # 截图中间的行加高 60：截图被拉长，红框挪了 80
        self.assertEqual(sg.restore_book(book, snap, "测试"), 2)
        self.assertEqual((pic.Top, pic.Height, box.Top), (40, 310, 300))

    def test_unreadable_book_is_ignored(self):
        self.assertEqual(sg.snapshot_book(object()), {})
        self.assertEqual(sg.restore_book(object(), {"手順": []}), 0)


if __name__ == "__main__":
    unittest.main()
