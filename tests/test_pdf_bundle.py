"""批量导出同名 PDF：文件名、合并所有表、跳过隐藏表 / 空白表、只读不保存、同名冲突、调整页面不受影响。"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

from core import pdf_bundle, workflow as wf


def write_pdf(path, pages):
    doc = fitz.open()
    for _ in range(pages):
        doc.new_page()
    doc.save(str(path))
    doc.close()


class Shapes:
    def __init__(self, count):
        self.Count = count


class Used:
    def __init__(self, count, formula):
        self.Count, self.Formula = count, formula


class Sheet:
    def __init__(self, name, pages=1, visible=-1, blank=False, protect_unhide=False):
        self.Name, self.pages, self._visible, self.protect_unhide = name, pages, visible, protect_unhide
        self.Shapes = Shapes(0)
        self.UsedRange = Used(1, "") if blank else Used(20, None)
        self.exported = []

    @property
    def Visible(self):
        return self._visible

    @Visible.setter
    def Visible(self, value):
        if self.protect_unhide:
            raise RuntimeError("ブックが保護されています")
        self._visible = value

    def ExportAsFixedFormat(self, **kw):
        self.exported.append(kw)
        write_pdf(kw["Filename"], self.pages)


class Book:
    def __init__(self, sheets):
        self.Sheets = sheets
        self.closed_with = None

    def Close(self, SaveChanges):
        self.closed_with = SaveChanges


class BundleTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)

    def export(self, name, sheets, include_hidden=False):
        path = self.root / name
        path.write_bytes(b"x")
        book = Book(sheets)
        with patch.object(pdf_bundle, "open_workbook_checked", return_value=book) as opened:
            result = pdf_bundle.export_workbook(object(), path, include_hidden)
        self.assertEqual(opened.call_args.kwargs["ReadOnly"], True)
        self.assertIs(book.closed_with, False)
        return path, result

    def test_same_name_pdf_next_to_excel_with_all_sheets(self):
        path, result = self.export("aaa.xlsx", [Sheet("表紙", 1), Sheet("一覧", 3), Sheet("詳細", 2)])
        pdf = self.root / "aaa.pdf"
        self.assertEqual(result["pdf"], str(pdf))
        self.assertEqual(result["sheets"], ["表紙", "一覧", "詳細"])
        with fitz.open(pdf) as doc:
            self.assertEqual(doc.page_count, 6)
            self.assertEqual([t[1:] for t in doc.get_toc()], [["表紙", 1], ["一覧", 2], ["詳細", 5]])
        self.assertEqual(result["pages"], 6)
        self.assertEqual([p.name for p in self.root.iterdir() if p.suffix == ".pdf"], ["aaa.pdf"])

    def test_review_star_prefix_dropped(self):
        self.assertEqual(pdf_bundle.pdf_path_for(Path("/x/★aaa.xlsm")).name, "aaa.pdf")

    def test_hidden_and_blank_sheets_skipped(self):
        hidden = Sheet("隠し", 1, visible=0)
        _, result = self.export("b.xlsx", [Sheet("A", 2), hidden, Sheet("空", blank=True)])
        self.assertEqual(result["sheets"], ["A"])
        self.assertEqual([s["sheet"] for s in result["skipped"]], ["隠し", "空"])
        self.assertEqual(hidden.exported, [])

    def test_include_hidden_unhides_in_memory(self):
        hidden = Sheet("隠し", 1, visible=2)
        _, result = self.export("c.xlsx", [Sheet("A", 1), hidden, Sheet("保護", visible=0, protect_unhide=True)], True)
        self.assertEqual(result["sheets"], ["A", "隠し"])
        self.assertEqual(result["pages"], 2)
        self.assertIn("保護", result["skipped"][0]["sheet"])

    def test_nothing_exportable_is_error(self):
        with self.assertRaises(ValueError):
            self.export("d.xlsx", [Sheet("空", blank=True)])
        self.assertFalse((self.root / "d.pdf").exists())

    def test_overwrites_existing_pdf(self):
        write_pdf(self.root / "e.pdf", 9)
        self.export("e.xlsx", [Sheet("A", 1)])
        with fitz.open(self.root / "e.pdf") as doc:
            self.assertEqual(doc.page_count, 1)

    def test_bundle_pdf_recognized_and_not_cleared_as_old(self):
        (self.root / "aaa.xlsx").write_bytes(b"x")
        (self.root / "aaa_b.xlsx").write_bytes(b"x")
        for name in ["aaa.pdf", "aaa_b.pdf", "aaa_表1.pdf"]:
            write_pdf(self.root / name, 1)
        self.assertTrue(pdf_bundle.is_bundle_pdf(self.root / "aaa.pdf"))
        self.assertFalse(pdf_bundle.is_bundle_pdf(self.root / "aaa_表1.pdf"))
        wf._clear_old_pdfs([self.root / "aaa.xlsx", self.root / "aaa_b.xlsx"])
        left = sorted(p.name for p in self.root.iterdir() if p.suffix == ".pdf")
        self.assertEqual(left, ["aaa.pdf", "aaa_b.pdf"])  # 只删以前按表混放的旧 PDF

    def test_options_validated(self):
        self.assertEqual(pdf_bundle.options_for(None), {"include_hidden": False})
        with self.assertRaises(ValueError):
            pdf_bundle.options_for({"include_hidden": "yes"})
        with self.assertRaises(ValueError):
            pdf_bundle.options_for({"other": True})


if __name__ == "__main__":
    unittest.main()
