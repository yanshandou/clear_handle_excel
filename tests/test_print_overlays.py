import unittest

from core import print_overlays


class Shape:
    def __init__(self, name, kind, left, top, w, h, visible=True, placement=1):
        self.Name, self.Type = name, kind
        self.Left, self.Top, self.Width, self.Height = left, top, w, h
        self.Visible, self.Placement = visible, placement
        Shape.count = getattr(Shape, "count", 0) + 1
        self.ID = Shape.count


class Group:
    def __init__(self, shapes, members):
        self.shapes, self.members = shapes, members
        self.Type = 6

    def Ungroup(self):
        for s in self.members:
            s.Placement = 2  # Excel 拆开后 Placement 可能被改掉
        self.shapes.items.remove(self)
        self.shapes.items.extend(self.members)
        members = self.members
        return type("Parts", (), {"Count": len(members), "Item": staticmethod(lambda i: members[i - 1])})


class Range:
    def __init__(self, shapes, names):
        self.shapes, self.names = shapes, names

    def Group(self):
        if self.shapes.fail:
            raise RuntimeError("保護されたシート")
        members = [self.shapes.Item(n) for n in self.names]
        self.shapes.keys.append(list(self.names))
        for m in members:
            self.shapes.items.remove(m)
        group = Group(self.shapes, members)
        self.shapes.items.append(group)
        self.shapes.groups.append(members)
        return group


class Shapes:
    def __init__(self, items, fail=False):
        self.items, self.fail, self.groups, self.keys = list(items), fail, [], []

    @property
    def Count(self):
        return len(self.items)

    def Item(self, key):
        if isinstance(key, int):
            return self.items[key - 1]
        return next(s for s in self.items if s.Name == key)

    def Range(self, names):
        return Range(self, list(names))


class Sheet:
    def __init__(self, items, fail=False):
        self.Name = "手順"
        self.Shapes = Shapes(items, fail)


class PrintOverlaysTest(unittest.TestCase):
    def test_red_box_on_picture_grouped_then_restored(self):
        pic = Shape("図 1", 13, 30, 40, 330, 310)
        box = Shape("正方形 2", 1, 220, 300, 50, 33)
        far = Shape("矢印 3", 1, 500, 40, 50, 20)      # 不在截图上
        sheet = Sheet([pic, box, far])
        with print_overlays.grouped_overlays(sheet) as n:
            self.assertEqual(n, 1)
            self.assertEqual(sheet.Shapes.groups, [[pic, box]])
            self.assertNotIn(box, sheet.Shapes.items)
        self.assertIn(box, sheet.Shapes.items)
        self.assertEqual((pic.Placement, box.Placement, far.Placement), (1, 1, 1))

    def test_smallest_picture_wins_and_hidden_ignored(self):
        big = Shape("図 1", 13, 0, 0, 800, 800)
        small = Shape("図 2", 11, 100, 100, 200, 200)
        box = Shape("正方形 3", 1, 150, 150, 20, 20)
        hidden = Shape("正方形 4", 1, 10, 10, 20, 20, visible=False)
        sets = print_overlays.find_overlay_sets(Sheet([big, small, box, hidden]))
        self.assertEqual([(p.Name, [m.Name for m in ms]) for p, ms in sets], [("図 2", ["正方形 3"])])

    def test_duplicate_names_skipped(self):
        pic = Shape("図 1", 13, 0, 0, 300, 300)
        a = Shape("正方形 2", 1, 10, 10, 20, 20)
        b = Shape("正方形 2", 1, 500, 500, 20, 20)
        self.assertEqual(print_overlays.find_overlay_sets(Sheet([pic, a, b])), [])

    def test_group_failure_exports_as_is(self):
        pic = Shape("図 1", 13, 0, 0, 300, 300)
        box = Shape("正方形 2", 1, 10, 10, 20, 20)
        sheet = Sheet([pic, box], fail=True)
        with print_overlays.grouped_overlays(sheet) as n:
            self.assertEqual(n, 0)
        self.assertEqual(sheet.Shapes.items, [pic, box])

    def test_unreadable_sheet_does_nothing(self):
        class Bad:
            Name = "グラフ"
        with print_overlays.grouped_overlays(Bad()) as n:
            self.assertEqual(n, 0)


class ClusterTest(unittest.TestCase):
    """没有底图、全靠图形拼出来的画面（用户 2026-11-11：PDF 里蓝色按钮掉到绿色标注下面）"""
    def test_touching_shapes_cluster_and_far_ones_stay(self):
        frame = Shape("枠 1", 1, 0, 0, 600, 300)
        button = Shape("ボタン 2", 6, 20, 280, 120, 40)        # 已有的组合，伸出框外
        callout = Shape("吹き出し 3", 2, 100, 322, 400, 30)     # 只和按钮挨着（隔 2 磅）
        alone = Shape("メモ 4", 17, 900, 900, 50, 20)
        chart = Shape("グラフ 5", 3, 10, 10, 50, 50)            # 图表不碰
        sheet = Sheet([frame, button, callout, alone, chart])
        self.assertEqual(print_overlays.find_clusters(sheet), [[frame, button, callout]])
        with print_overlays.grouped_overlays(sheet) as n:
            self.assertEqual(n, 1)
            self.assertEqual(sheet.Shapes.groups, [[frame, button, callout]])
        self.assertEqual({s.Placement for s in (frame, button, callout)}, {1})

    def test_duplicate_names_grouped_by_index_and_placement_restored_by_id(self):
        a = Shape("正方形 2", 1, 0, 0, 50, 50, placement=3)
        other = Shape("線 9", 9, 500, 500, 0, 40)
        b = Shape("正方形 2", 1, 40, 40, 50, 50, placement=1)   # 复制出来的同名图形
        sheet = Sheet([a, other, b])
        with print_overlays.grouped_overlays(sheet) as n:
            self.assertEqual(n, 1)
            self.assertEqual(sheet.Shapes.keys, [[1, 3]])
            self.assertEqual(sheet.Shapes.groups, [[a, b]])
        self.assertEqual((a.Placement, b.Placement), (3, 1))


class GroupPermanentlyTest(unittest.TestCase):
    def test_grouped_and_set_move_dont_size(self):
        pic = Shape("図 1", 13, 30, 40, 330, 310)
        box = Shape("正方形 2", 1, 220, 300, 50, 33)
        sheet = Sheet([pic, box])
        done = print_overlays.group_permanently(sheet)
        self.assertEqual(len(done), 1)
        self.assertEqual(sheet.Shapes.groups, [[pic, box]])
        group = next(s for s in sheet.Shapes.items if isinstance(s, Group))
        self.assertEqual(group.Placement, print_overlays.XL_MOVE)

    def test_nothing_to_group(self):
        self.assertEqual(print_overlays.group_permanently(Sheet([Shape("図 1", 13, 0, 0, 100, 100)])), [])


if __name__ == "__main__":
    unittest.main()
