import importlib
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

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
        with patch.object(self.wf, "log_msg") as log:
            mark_finder.start(str(self.folder), MAGENTA)
            for _ in range(100):
                if not mark_finder.status()["running"]:
                    break
                time.sleep(0.05)
        result = mark_finder.status()["result"]
        self.assertEqual(result["total"], 2)
        files = {f["name"]: f for f in result["files"]}
        self.assertEqual(len(files["book.xlsx"]["hits"]), 3)
        self.assertIn("读不了", files["old.xls"]["error"])
        self.assertTrue(any("3 格底色是标记色" in c.args[0] for c in log.call_args_list))

    def test_bad_colors_are_refused(self):
        with self.assertRaises(ValueError):
            mark_finder._styles([{"fill": "red", "font": "#FFFFFF"}])


if __name__ == "__main__":
    unittest.main()
