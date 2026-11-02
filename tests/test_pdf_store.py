"""PDFS 文件夹和导出记录（按 Excel 哈希沿用 PDF）"""
import tempfile
import unittest
from pathlib import Path

from core import pdf_store


class PdfStoreTests(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp()).resolve()

    def pdf(self, rel, data=b"%PDF"):
        path = self.root / "PDFS" / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        return path

    def test_pdf_dir_keeps_subfolders(self):
        wb = self.root / "01_画面" / "spec.xlsx"
        self.assertEqual(pdf_store.pdf_dir(wb, self.root), self.root / "PDFS" / "01_画面")
        self.assertEqual(pdf_store.export_root_of(self.root / "PDFS" / "01_画面" / "spec_A.pdf"), self.root)
        self.assertIsNone(pdf_store.export_root_of(self.root / "01_画面" / "spec_A.pdf"))

    def test_star_prefix_is_same_workbook(self):
        wb = self.root / "spec.xlsx"
        wb.write_bytes(b"x")
        pdf = self.pdf("spec_A.pdf")
        pdf_store.record(self.root, wb, "h1", [(pdf, "A")])
        starred = wb.rename(self.root / "★spec.xlsx")
        self.assertEqual(pdf_store.reusable(self.root, starred, "h1"), [(pdf, "A")])
        self.assertIsNone(pdf_store.reusable(self.root, starred, "h2"))
        self.assertEqual(pdf_store.owner(pdf), {"wb_path": str(starred), "sheet": "A"})

    def test_clear_keeps_pdfs_of_longer_named_sibling(self):
        a = self.root / "a.xlsx"
        ab = self.root / "a_b.xlsx"
        for wb in (a, ab):
            wb.write_bytes(b"x")
        mine, leftover, theirs = self.pdf("a_S1.pdf"), self.pdf("a_old.pdf"), self.pdf("a_b_S1.pdf")
        unrelated = self.pdf("readme.pdf")
        pdf_store.record(self.root, a, "h", [(mine, "S1")])
        removed = pdf_store.clear_workbook(self.root, a)
        self.assertEqual(sorted(p.name for p in removed), ["a_S1.pdf", "a_old.pdf"])
        self.assertTrue(theirs.exists() and unrelated.exists())
        self.assertNotIn("a.xlsx", pdf_store.load(self.root)["files"])
        self.assertFalse(leftover.exists())

    def test_prune_removes_gone_workbooks_and_empty_folders(self):
        wb = self.root / "sub" / "x.xlsx"
        wb.parent.mkdir()
        wb.write_bytes(b"x")
        pdf = self.pdf("sub/x_A.pdf")
        pdf_store.record(self.root, wb, "h", [(pdf, "A")])
        self.assertEqual(pdf_store.prune(self.root, []), [pdf])
        self.assertFalse((self.root / "PDFS" / "sub").exists())
        self.assertEqual(pdf_store.load(self.root)["files"], {})

    def test_broken_manifest_means_export_again(self):
        (self.root / "PDFS").mkdir()
        (self.root / "PDFS" / pdf_store.MANIFEST_NAME).write_text("{oops", encoding="utf-8")
        self.assertIsNone(pdf_store.reusable(self.root, self.root / "x.xlsx", "h"))


if __name__ == "__main__":
    unittest.main()
