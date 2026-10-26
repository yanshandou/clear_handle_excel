import importlib
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import openpyxl
from openpyxl.styles import Font, PatternFill

from core import fix_images, mark_finder, reporter, storage, workflow

MAGENTA = [{"fill": "#C0007A", "font": "#FFFFFF"}, {"fill": "#000000", "font": "#FFFF00"}]


def paint(cell, fill, font=None):
    cell.fill = PatternFill("solid", fgColor="FF" + fill)
    if font:
        cell.font = Font(color="FF" + font)


class MarkFinderTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        for obj, name, value in ((storage, "DATABASE_PATH", self.root / "state.sqlite3"),
                                 (reporter, "RECORD_PATH", self.root / "record.md"),
                                 (fix_images, "IMAGES_DIR", self.root / "images")):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.wf = importlib.reload(workflow)
        mark_finder.wf = self.wf
        self.folder = self.root / "成品"
        self.folder.mkdir()
        self.path = self.folder / "book.xlsx"
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "明細"
        paint(sheet["A5"], "C0007A", "FFFFFF")       # 底色字色都对上
        paint(sheet["B7"], "C0007A", "123456")       # 只有底色对上（一格几种字色时旧版不改字色）
        paint(sheet["C9"], "C0007B", "FFFFFF")       # 差一点也不算
        paint(sheet["D3"], "FFFF00")                 # 别的颜色
        sheet.merge_cells("E2:G2")
        for col in "EFG":
            paint(sheet[f"{col}2"], "000000", "FFFF00")  # 合并区域整块算一格
        book.save(self.path)
        (self.folder / "old.xls").write_bytes(b"x")

    def test_exact_marker_colors_are_listed_merged_area_once(self):
        hits = mark_finder.scan_file(self.path, mark_finder._styles(MAGENTA), set())
        found = {h["addr"]: h["match"] for h in hits}
        self.assertEqual(found, {"A5": "both", "B7": "fill", "E2:G2": "both"})
        self.assertEqual(next(h for h in hits if h["addr"] == "B7")["font"], "#123456")

    def test_cells_the_program_still_has_records_for_are_flagged(self):
        self.wf.STATE.setdefault("batch_review", {})["highlights"] = {
            "x": {"path": str(self.path), "sheet": "明細", "addr": "$A$5", "cells": []},
            "y": {"path": str(self.path), "sheet": "明細", "addr": "$B$7", "tag": "CHE_MARK_y"}}  # 新方式的不算
        hits = mark_finder.scan_file(self.path, mark_finder._styles(MAGENTA), mark_finder._recorded(str(self.path)))
        self.assertEqual({h["addr"] for h in hits if h["recorded"]}, {"A5"})

    def test_background_run_reports_each_file_and_unreadable_formats(self):
        with patch.object(self.wf, "log_msg") as log, \
                patch("core.excel_ops.start_background_excel", side_effect=RuntimeError("没有 Excel")):
            mark_finder.start(str(self.folder), MAGENTA)
            for _ in range(100):
                if not mark_finder.status()["running"]:
                    break
                time.sleep(0.05)
        result = mark_finder.status()["result"]
        self.assertEqual(result["total"], 2)
        files = {f["name"]: f for f in result["files"]}
        self.assertEqual(len(files["book.xlsx"]["hits"]), 3)
        self.assertIn("用 Excel 读取失败", files["old.xls"]["error"])
        self.assertTrue(any("3 格底色是标记色" in c.args[0] for c in log.call_args_list))

    def old_copy(self, edit=None):
        """旧版文件：同样的内容，原来的颜色"""
        old_root = self.root / "原始备份"
        old_root.mkdir(exist_ok=True)
        book = openpyxl.load_workbook(self.path)
        sheet = book["明細"]
        for addr in ("A5", "B7", "E2", "F2", "G2"):
            sheet[addr].fill = PatternFill(fill_type=None)
            sheet[addr].font = Font(color="FF000000")
        paint(sheet["B7"], "DDEEFF", "123456")
        if edit:
            edit(sheet)
        book.save(old_root / "book.xlsx")
        return old_root

    def plan(self, old_root):
        hits = mark_finder.scan_file(self.path, mark_finder._styles(MAGENTA), set())
        return {c["addr"]: c for c in mark_finder.plan_file(self.path, old_root / "book.xlsx", hits)}

    def test_plan_restores_cells_whose_content_matches_the_old_version(self):
        book = openpyxl.load_workbook(self.path)
        book["明細"]["A5"] = "合計"
        book["明細"]["B7"] = 12
        book.save(self.path)
        plan = self.plan(self.old_copy())
        self.assertEqual((plan["A5"]["action"], plan["A5"]["old_fill"], plan["A5"]["old_font"]), ("restore", "none", "#000000"))
        self.assertEqual((plan["B7"]["action"], plan["B7"]["old_fill"]), ("restore", "#DDEEFF"))
        self.assertEqual(plan["E2:G2"]["action"], "restore")

    def test_plan_skips_moved_or_changed_cells_and_original_marker_colors(self):
        book = openpyxl.load_workbook(self.path)
        book["明細"]["A5"] = "合計"
        book["明細"]["B7"] = "x"
        book.save(self.path)

        def edit(sheet):
            sheet["A5"] = "小計"                 # 内容不同（多半是插了行）
            sheet["B7"] = "x"
            sheet["B8"] = "多了一格"             # 不影响 B7：B7 有内容，只比这一格
            sheet["H2"] = "这一行多了内容"       # E2 是空格子：这一行和旧版对不上
            paint(sheet["B7"], "C0007A", "123456")  # 旧版本来就是这个颜色
        plan = self.plan(self.old_copy(edit))
        self.assertIn("内容不一样", plan["A5"]["reason"])
        self.assertIn("本来就是这个颜色", plan["B7"]["reason"])
        self.assertIn("这一行和旧版对不上", plan["E2:G2"]["reason"])
        self.assertTrue(all(c["action"] == "skip" for c in plan.values()))

    def test_renamed_sheets_are_matched_by_guessed_name_code_name_then_content(self):
        book = openpyxl.load_workbook(self.path)
        book["明細"]["A5"] = "合計"
        other = book.create_sheet("レイアウト")
        other.sheet_properties.codeName = "Sheet9"
        other["A1"], other["A2"] = "訪問看護計画書", "氏名"
        paint(other["A2"], "C0007A", "FFFFFF")
        bill = book.create_sheet("請求書")
        bill["C3"] = "金額"
        paint(bill["C3"], "C0007A", "FFFFFF")
        book["明細"].sheet_properties.codeName = "Sheet1"
        book.save(self.path)

        def edit(sheet):
            sheet.title = "明細新"                 # 精简 Sheet 去掉了「新」，代码名没变
            sheet.sheet_properties.codeName = "Sheet1"
            lay = sheet.parent["レイアウト"]
            lay.title = "レイアウト(訪問看護計画書)"
            lay.sheet_properties.codeName = "Sheet3"   # 代码名也对不上，只能按内容
            lay["A2"].fill = PatternFill(fill_type=None)
            sheet.parent.create_sheet("全然違う")["A1"] = "別の表"
            bill = sheet.parent["請求書"]
            bill.title = "新請求書"                  # 「新」字被去掉了：加回去就能对上
            bill["C3"].fill = PatternFill(fill_type=None)
        plan = self.plan(self.old_copy(edit))
        self.assertEqual((plan["A5"]["action"], plan["A5"]["old_sheet"]), ("restore", "明細新"))
        self.assertIn("按代码名对上", plan["A5"]["reason"])
        self.assertEqual((plan["A2"]["action"], plan["A2"]["old_sheet"]), ("restore", "レイアウト(訪問看護計画書)"))
        self.assertIn("按内容对上", plan["A2"]["reason"])
        self.assertEqual((plan["C3"]["action"], plan["C3"]["old_sheet"]), ("restore", "新請求書"))
        self.assertIn("去掉「新」", plan["C3"]["reason"])

    def test_sheet_without_any_match_lists_old_sheet_names(self):
        def edit(sheet):
            sheet.title = "別の名前"
            sheet.sheet_properties.codeName = "Sheet7"
            for addr in ("A1", "A5", "B7"):
                sheet[addr] = "まったく違う内容"
        plan = self.plan(self.old_copy(edit))
        self.assertEqual(plan["A5"]["action"], "skip")
        self.assertIn("別の名前", plan["A5"]["reason"])

    def test_files_are_paired_by_path_then_by_unique_name(self):
        old = self.root / "old"
        paths = [old / "a.xlsx", old / "x" / "b.xlsx", old / "y" / "c.xlsx", old / "z" / "c.xlsx"]
        pairs = mark_finder._pair_files(old, paths, ["a.xlsx", "sub/b.xlsx", "c.xlsx", "d.xlsx"])
        self.assertEqual(pairs, {"a.xlsx": paths[0], "sub/b.xlsx": paths[1], "c.xlsx": None, "d.xlsx": None})

    def test_apply_copies_old_colors_only_when_still_marked_and_same_content(self):
        marker = mark_finder._excel_color("#C0007A")

        def cell(formula, pattern, color, font, tint=0):
            return Mock(Formula=formula, Interior=Mock(Pattern=pattern, Color=color, TintAndShade=tint), Font=Mock(Color=font))
        cur = {"A5": cell("合計", 1, marker, 0xFFFFFF), "B7": cell("x", 1, 0x123456, 0), "C9": cell("新", 1, marker, 0)}
        old = {"A5": cell("合計", -4142, 0xFFFFFF, 0), "B7": cell("x", -4142, 0xFFFFFF, 0),
               "C9": cell("旧", -4142, 0xFFFFFF, 0)}
        for c in cur.values():
            c.Cells.side_effect = lambda r, col, c=c: c
        old_sheet, sheet = Mock(Name="明細"), Mock()
        old_sheet.Range.side_effect = lambda a: old[a]
        sheet.Range.side_effect = lambda a: cur[a]
        old_book, book = Mock(Worksheets=[old_sheet]), Mock(ReadOnly=False)
        book.Worksheets.return_value = sheet
        excel = Mock()
        opened = []

        def open_book(path, **_):
            # 旧版和新版同名：Excel 里已经开着同名文件时 Open 返回空对象，所以旧版要先读完关掉
            if opened and not opened[-1].Close.called:
                return None
            opened.append(old_book if not opened else book)
            return opened[-1]
        excel.Workbooks.Open.side_effect = open_book
        old_sheet.Name = "明細新"  # 旧版表名不同：按清单里对上的旧表名去读
        cells = [{"sheet": "明細", "old_sheet": "明細新", "addr": a, "fill": "#C0007A", "font": "#FFFFFF", "action": "restore"} for a in cur]
        with patch("core.excel_ops.normalize_workbook_view"):
            done = {c["addr"]: c for c in mark_finder.apply_file(excel, str(self.path), str(self.path), cells)}
        self.assertEqual(done["A5"]["status"], "ok")
        self.assertEqual((cur["A5"].Interior.Pattern, cur["A5"].Font.Color), (-4142, 0))
        self.assertIn("被人改过", done["B7"]["reason"])
        self.assertIn("内容和旧版不一样", done["C9"]["reason"])
        self.assertEqual(cur["C9"].Interior.Color, marker)
        book.Save.assert_called_once()

    def test_bad_colors_are_refused(self):
        with self.assertRaises(ValueError):
            mark_finder._styles([{"fill": "red", "font": "#FFFFFF"}])


def bgr(hex_color):
    return mark_finder._excel_color("#" + hex_color)


def col_name(col):
    return "ABCDEFGHIJ"[col - 1]


class FakeRange:
    """Excel Range 替身：整块颜色一样时 Interior.Color 给颜色，不一样给 None（和真 Excel 一样）"""
    def __init__(self, sheet, row, col, rows=1, cols=1):
        self.sheet, self.row, self.col, self.n_rows, self.n_cols = sheet, row, col, rows, cols
        self.Rows, self.Columns = Mock(Count=rows), Mock(Count=cols)
        sheet.calls += 1

    def _cells(self):
        return [(r, c) for r in range(self.row, self.row + self.n_rows) for c in range(self.col, self.col + self.n_cols)]

    @property
    def Interior(self):
        colors = {self.sheet.fill.get(rc, (bgr("FFFFFF"), -4142))[0] for rc in self._cells()}
        pattern = self.sheet.fill.get((self.row, self.col), (0, -4142))[1]
        return Mock(Color=colors.pop() if len(colors) == 1 else None, Pattern=pattern)

    @property
    def Font(self):
        return Mock(Color=self.sheet.font.get((self.row, self.col), 0))

    def Offset(self, rows, cols):
        return FakeRange(self.sheet, self.row + rows, self.col + cols, self.n_rows, self.n_cols)

    def Resize(self, rows, cols):
        return FakeRange(self.sheet, self.row, self.col, rows, cols)

    @property
    def Cells(self):
        return [FakeRange(self.sheet, r, c) for r, c in self._cells()]

    @property
    def MergeCells(self):
        return any((self.row, self.col) in area for area in self.sheet.merged)

    @property
    def MergeArea(self):
        area = next(a for a in self.sheet.merged if (self.row, self.col) in a)
        first, last = min(area), max(area)
        return Mock(Address=f"${col_name(first[1])}${first[0]}:${col_name(last[1])}${last[0]}")

    @property
    def Address(self):  # pywin32 里 Address 是属性，返回带 $ 的字符串
        return f"${col_name(self.col)}${self.row}"


class FakeSheet:
    def __init__(self, name, rows, cols):
        self.Name, self.rows, self.cols = name, rows, cols
        self.fill, self.font, self.merged, self.calls = {}, {}, [], 0

    def paint(self, row, col, fill, font=None):
        self.fill[(row, col)] = (bgr(fill), 1)
        if font:
            self.font[(row, col)] = bgr(font)

    @property
    def UsedRange(self):
        return FakeRange(self, 1, 1, self.rows, self.cols)

    def Range(self, addr):
        return FakeRange(self, int(addr[1:]), "ABCDEFGHIJ".index(addr[0]) + 1)


class ExcelReadTests(unittest.TestCase):
    """.xls / .xlsb 用后台 Excel 读颜色"""
    def scan(self, sheet, suffix=".xls", saved=None, recorded=()):
        book = Mock(Worksheets=[sheet])
        excel = Mock()
        excel.Workbooks.Open.return_value = book
        hits = mark_finder.scan_file_excel(excel, Path("C:/x/book" + suffix), mark_finder._styles(MAGENTA),
                                           set(recorded), saved or {})
        self.assertEqual(excel.Workbooks.Open.call_args.kwargs["ReadOnly"], True)
        book.Close.assert_called_once_with(SaveChanges=False)
        return {h["addr"]: h for h in hits}

    def test_marker_cells_are_found_merged_area_once(self):
        sheet = FakeSheet("明細", 40, 8)
        sheet.paint(5, 1, "C0007A", "FFFFFF")      # A5 底色字色都对上
        sheet.paint(7, 2, "C0007A", "123456")      # B7 只有底色
        sheet.paint(9, 3, "C0007B", "FFFFFF")      # 差一点不算
        sheet.paint(3, 4, "FFFF00")                # 别的颜色
        sheet.merged.append({(2, c) for c in (5, 6, 7)})
        for c in (5, 6, 7):
            sheet.paint(2, c, "000000", "FFFF00")  # E2:G2 合并区域算一格
        hits = self.scan(sheet, recorded=[("明細", "A5")])
        self.assertEqual({a: h["match"] for a, h in hits.items()}, {"A5": "both", "B7": "fill", "E2:G2": "both"})
        self.assertEqual(hits["B7"]["font"], "#123456")
        self.assertTrue(hits["A5"]["recorded"])
        self.assertLess(sheet.calls, 40 * 8)  # 对半切开找，不是逐格读

    def test_xls_color_replaced_by_palette_color_still_counts(self):
        sheet = FakeSheet("表", 10, 3)
        sheet.paint(4, 2, "993366", "FFFFFF")      # 存进 .xls 后标记色被换成了 56 色里的这一种
        saved = {("C0007A", "FFFFFF"): ("993366", "FFFFFF")}
        hits = self.scan(sheet, saved=saved)
        self.assertEqual(hits["B4"]["match"], "both")
        self.assertTrue(hits["B4"]["approx"])
        self.assertEqual(self.scan(sheet, ".xlsb", saved=saved), {})  # .xlsb 不会换色

    def test_blank_sheet_is_one_call(self):
        sheet = FakeSheet("空", 500, 10)
        self.assertEqual(self.scan(sheet), {})
        self.assertLessEqual(sheet.calls, 2)


if __name__ == "__main__":
    unittest.main()
