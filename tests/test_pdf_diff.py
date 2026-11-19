"""PDF 左右对比（用户 2026-11-19）：A + B 拖到左边、C 拖到右边，接起来逐行比文字"""
import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from core import pdf_diff
from core.pdf_detector import fitz


def make_pdf(pages: list[list[str]], title: str = "画面設計書") -> bytes:
    """每页：页眉标题 + 每行一段字 + 页脚页码"""
    doc = fitz.open()
    for number, rows in enumerate(pages, 1):
        page = doc.new_page(width=595, height=842)
        page.insert_text((50, 40), title, fontsize=12)
        for i, row in enumerate(rows):
            page.insert_text((50, 70 + i * 18), row, fontsize=10)
        page.insert_text((280, 810), f"{number} / {len(pages)}", fontsize=9)
    body = doc.tobytes()
    doc.close()
    return body


def rows(start: int, count: int) -> list[str]:
    return [f"No.{i}  item {i}  check {i % 7}" for i in range(start, start + count)]


class PdfDiffTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        patcher = mock.patch.object(pdf_diff, "DATA_DIR", Path(self.tmp.name))
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(self.tmp.cleanup)
        pdf_diff._text_cache.clear()

    def store(self, name, pages):
        return pdf_diff.store(name, make_pdf(pages))["id"]

    def test_a_plus_b_equals_c_ignores_page_breaks_headers_and_page_numbers(self):
        a = self.store("A.pdf", [rows(1, 20), rows(21, 20)])
        b = self.store("B.pdf", [rows(41, 15)])
        c = self.store("C.pdf", [rows(1, 30), rows(31, 25)])  # 分页位置不一样、页码不一样
        result = pdf_diff.diff([a, b], [c])
        self.assertTrue(result["stats"]["same"], result["ops"])
        self.assertEqual(len(result["left"]), 55)
        self.assertEqual(result["files"]["left"], [{"pages": 2, "lines": 40}, {"pages": 1, "lines": 15}])
        self.assertGreater(result["dropped"]["left"], 0)

    def test_headers_compared_when_option_off(self):
        a = self.store("A.pdf", [rows(1, 20), rows(21, 20)])
        c = self.store("C.pdf", [rows(1, 30), rows(31, 10)])
        result = pdf_diff.diff([a], [c], {"ignore_header": False})
        self.assertFalse(result["stats"]["same"])
        self.assertEqual(result["dropped"], {"left": 0, "right": 0})

    def test_changed_removed_added_lines(self):
        left = rows(1, 10)
        right = list(left)
        right[2] = right[2].replace("check 3", "check 3 (length)")
        del right[5]
        right.insert(8, "No.9-2  new item")
        result = pdf_diff.diff([self.store("L.pdf", [left])], [self.store("R.pdf", [right])])
        stats = result["stats"]
        self.assertEqual((stats["hunks"], stats["changed"], stats["removed"], stats["added"]), (3, 1, 1, 1))
        changed = [op for op in result["ops"] if op[0] == "r"][0]
        left_ranges, right_ranges = result["chars"][f"{changed[1]}:{changed[3]}"]
        text = result["right"][changed[3]]["s"]
        self.assertIn("(length)", "".join(text[a:b] for a, b in right_ranges))
        self.assertEqual(left_ranges, [])
        # 每行记着在哪份文件第几页
        self.assertEqual({(x["f"], x["p"]) for x in result["right"]}, {(0, 1)})

    def test_ignore_space(self):
        left = self.store("L.pdf", [["ABC DEF"]])
        right = self.store("R.pdf", [["ABCDEF"]])
        self.assertTrue(pdf_diff.diff([left], [right])["stats"]["same"])
        self.assertFalse(pdf_diff.diff([left], [right], {"ignore_space": False})["stats"]["same"])

    def test_page_number_only_dropped_at_page_edge(self):
        self.assertTrue(pdf_diff._is_page_no("3 / 12"))
        self.assertTrue(pdf_diff._is_page_no("- 3 -"))
        self.assertTrue(pdf_diff._is_page_no("Page 3 of 9"))
        self.assertTrue(pdf_diff._is_page_no("3ページ"))
        self.assertFalse(pdf_diff._is_page_no("No.3"))
        # 表中间只有数字的一行照样比
        left = self.store("L.pdf", [["head", "x", "5", "y", "z", "tail"]])
        right = self.store("R.pdf", [["head", "x", "6", "y", "z", "tail"]])
        self.assertFalse(pdf_diff.diff([left], [right])["stats"]["same"])

    def test_rejects_bad_input(self):
        with self.assertRaises(ValueError):
            pdf_diff.store("x.pdf", b"not a pdf")
        with self.assertRaises(ValueError):
            pdf_diff.store("x.pdf", b"")
        with self.assertRaises(ValueError):
            pdf_diff.diff([], ["0" * 32])
        with self.assertRaises(ValueError):
            pdf_diff.diff(["../../etc"], ["0" * 32])
        with self.assertRaises(ValueError):  # 过期被清掉的
            pdf_diff.diff(["0" * 32], ["1" * 32])

    def test_old_uploads_cleaned(self):
        old = self.store("old.pdf", [["a"]])
        path = Path(self.tmp.name) / "pdf_diff" / f"{old}.pdf"
        os.utime(path, (1, 1))
        self.store("new.pdf", [["b"]])
        self.assertFalse(path.exists())


if __name__ == "__main__":
    unittest.main()
