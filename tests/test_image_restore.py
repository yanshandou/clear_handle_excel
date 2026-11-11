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


def _pdf(target, pages):
    """pages：[(文字行列表, 有没有图片)]"""
    from core.pdf_detector import fitz
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 40, 30), 0)
    pix.clear_with(200)
    doc = fitz.open()
    for lines, picture in pages:
        page = doc.new_page()
        for i, line in enumerate(lines):
            page.insert_text((50, 60 + 20 * i), line, fontname="japan")
        if picture:
            page.insert_image(fitz.Rect(50, 300, 450, 600), pixmap=pix)
    doc.save(target)
    return target


class PageRowsTest(unittest.TestCase):
    def test_picture_pages_matched_by_text(self):
        import tempfile
        from pathlib import Path
        from unittest.mock import patch
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            steps = ["手順1 ログイン画面を開く", "実行ボタンを押下する", "結果一覧が表示される"]
            new = _pdf(tmp / "new.pdf", [(["表紙 システム概要"], False), (steps, True), (["補足 注意事項"], False)])
            old = _pdf(tmp / "old.pdf", [(["表紙 システム概要"], False), (["補足 注意事項"], False), (steps, True)])
            book = _pdf(tmp / "book.pdf", [(["別シート 変更履歴"], False), (["表紙 システム概要"], False), (steps, True)])
            with patch.object(ir.fix_images, "IMAGES_DIR", tmp):
                rows = ir.page_rows(tmp / "a.xlsx", "手順", new, old, book)
                self.assertEqual(len(rows), 1)
                row = rows[0]
                self.assertEqual((row["page"], row["old_page"], row["pdf_page"]), (2, 3, 3))
                for key in ("new_img", "old_img", "pdf_img"):
                    self.assertTrue((tmp / row[key]).is_file())
                none = ir.page_rows(tmp / "a.xlsx", "手順", new, None, None)
                self.assertIn("没导出", none[0]["old_why"])
                self.assertIn("没找到同名 PDF", none[0]["pdf_why"])
                other = _pdf(tmp / "other.pdf", [(["まったく別の内容"], False)])
                miss = ir.page_rows(tmp / "a.xlsx", "手順", new, other, other)
                self.assertIn("没找到文字对得上的页", miss[0]["old_why"])

    def test_no_bitmap_lists_first_pages(self):
        import tempfile
        from pathlib import Path
        from unittest.mock import patch
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            new = _pdf(tmp / "new.pdf", [([f"ページ{i} の内容です"], False) for i in range(8)])
            with patch.object(ir.fix_images, "IMAGES_DIR", tmp):
                rows = ir.page_rows(tmp / "a.xlsx", "手順", new, new, None)
            self.assertEqual(len(rows), ir.SHOW_ALL_MAX)
            self.assertTrue(rows[0]["note"])
            self.assertEqual(rows[3]["old_page"], 4)


class _Cell:
    def __init__(self, row, col):
        self.Row, self.Column = row, col


class _Range:
    def __init__(self, r1, c1, r2, c2):
        self.Row, self.Column, self.r2, self.c2 = r1, c1, r2, c2
        self.Rows = type("R", (), {"Count": r2 - r1 + 1})
        self.Columns = type("C", (), {"Count": c2 - c1 + 1})


class _Breaks:
    def __init__(self, rows):
        self.rows, self.Count = rows, len(rows)

    def Item(self, i):
        return type("B", (), {"Location": _Cell(self.rows[i - 1], 1)})


class PageSheet:
    """打印区域 A1:K120，分页在第 41、81 行"""
    def __init__(self, breaks=(41, 81), area="$A$1:$K$120"):
        self.PageSetup = type("P", (), {"PrintArea": area})
        self.HPageBreaks = _Breaks(list(breaks))
        self.UsedRange = _Range(3, 2, 100, 8)
        self.Shapes = type("S", (), {"Count": 1, "Item": staticmethod(lambda i: type("Sh", (), {"BottomRightCell": _Cell(130, 12)}))})

    def Activate(self):
        pass

    def Cells(self, row, col):
        return _Cell(row, col)

    def Range(self, a, b=None):
        return _Range(1, 1, 120, 11) if b is None else _Range(a.Row, a.Column, b.Row, b.Column)


class PageRangesTest(unittest.TestCase):
    def test_pages_are_row_bands_between_breaks(self):
        bands = ir.page_ranges(PageSheet(), 3)
        self.assertEqual([(b.Row, b.r2, b.c2) for b in bands], [(1, 40, 11), (41, 80, 11), (81, 120, 11)])

    def test_no_print_area_starts_at_a1_and_covers_shapes(self):
        bands = ir.page_ranges(PageSheet(area=""), 3)
        self.assertEqual([(b.Row, b.Column, b.r2, b.c2) for b in bands], [(1, 1, 40, 12), (41, 1, 80, 12), (81, 1, 130, 12)])

    def test_page_count_mismatch_refused(self):
        with self.assertRaises(RuntimeError):
            ir.page_ranges(PageSheet(), 4)

    def test_screen_pages_keeps_export_image_when_capture_fails(self):
        from pathlib import Path
        from unittest.mock import patch
        rows = [{"page": 2, "new_img": "pdf.png", "new_how": "pdf", "old_page": 3, "old_img": "o.png", "old_how": "pdf"}]
        with patch.object(ir, "render_screen", return_value=1.5) as shot:
            ir.screen_pages(PageSheet(), rows, "new", 3, Path("a.xlsx"), "手順")
            self.assertEqual((rows[0]["new_how"], rows[0]["new_ratio"], shot.call_args[0][1].Row), ("screen", 1.5, 41))
            ir.screen_pages(PageSheet(), rows, "old", 3, Path("a.xlsx"), "手順")
            self.assertEqual((rows[0]["old_how"], shot.call_args[0][1].Row), ("screen", 81))
        rows = [{"page": 2, "new_img": "pdf.png", "new_how": "pdf"}]
        with patch.object(ir, "render_screen", side_effect=RuntimeError("剪贴板被占用")):
            ir.screen_pages(PageSheet(), rows, "new", 3, Path("a.xlsx"), "手順")
        self.assertEqual((rows[0]["new_img"], rows[0]["new_how"]), ("pdf.png", "pdf"))
        ir.screen_pages(PageSheet(), rows, "new", 5, Path("a.xlsx"), "手順")  # 分页数对不上：整张表不拍
        self.assertEqual(rows[0]["new_how"], "pdf")


class MergeTest(unittest.TestCase):
    def test_group_result_merged_into_scan(self):
        from pathlib import Path
        root = Path("D:/x")
        scan = ir._merge(None, "scan", root, Path("D:/y"), None,
                         [{"name": "a.xlsx", "error": None, "count": 2, "sheets": []}])
        merged = ir._merge(scan, "group", root, None, None,
                           [{"name": "a.xlsx", "error": None, "count": 0, "grouped": 3, "group_sheets": []}])
        self.assertEqual(merged["files"][0]["count"], 2)
        self.assertEqual(merged["files"][0]["grouped"], 3)
        self.assertTrue(merged["scanned"])
        only = ir._merge(None, "group", root, None, None, [{"name": "b.xlsx", "error": None, "count": 0, "grouped": 1}])
        self.assertFalse(only["scanned"])
        self.assertEqual(only["old_folder"], "")

    def test_scan_keeps_stamp_so_images_do_not_reload(self):
        from pathlib import Path
        done = ir._merge(None, "scan", Path("D:/x"), Path("D:/y"), None, [], 12345)
        self.assertEqual(done["stamp"], 12345)


class OverlayHostsTest(unittest.TestCase):
    def test_every_lists_all_pictures_and_groups(self):
        ref = [s("図 1", 13, 0, 0, 300, 300), s("正方形 2", 1, 10, 10, 20, 20), s("図 3", 13, 400, 0, 100, 100),
               s("グループ 4", 6, 600, 0, 100, 100)]
        self.assertEqual(sg.overlay_hosts(ref, ref), {"図 1": ["正方形 2"]})
        self.assertEqual(sg.overlay_hosts(ref, ref, every=True), {"図 1": ["正方形 2"], "図 3": [], "グループ 4": []})
        info = sg.diagnose(ref, ref)
        self.assertEqual((info["配上"], info["截图"], info["组合"], info["大小最大差"]), (4, 2, 1, 0))


if __name__ == "__main__":
    unittest.main()
