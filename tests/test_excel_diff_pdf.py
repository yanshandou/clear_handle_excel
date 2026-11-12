"""对比的真实效果（core/excel_diff_pdf.py）：按文件哈希排队导 PDF、只导要的表、导过的不再导、页面图取得到"""
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from core import excel_diff_pdf, excel_ops, excel_watchdog, pdf_bundle
from core.pdf_detector import fitz


class _Sheet:
    def __init__(self, name, visible=-1):
        self.Name, self.Visible = name, visible


class _Book:
    def __init__(self):
        self.Sheets = [_Sheet("表1"), _Sheet("表2"), _Sheet("藏起来", 0)]
        self.closed = False

    def Close(self, SaveChanges=False):
        self.closed = True


class ExcelDiffPdfTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        self.book = root / "a.xlsx"
        self.book.write_bytes(b"fake excel")
        self.exported = []

        def export(sheet, target):
            self.exported.append(sheet.Name)
            doc = fitz.open()
            for _ in range(2):
                doc.new_page(width=200, height=300)
            doc.save(target)
        self.patches = [
            patch.object(excel_diff_pdf, "_root", lambda: root / "cache"),
            patch.object(excel_ops, "start_background_excel", lambda: object()),
            patch.object(excel_ops, "quit_background_excel", lambda excel: None),
            patch.object(excel_ops, "open_workbook_checked", lambda excel, path, **k: _Book()),
            patch.object(pdf_bundle, "_export_sheet", export),
            patch.object(pdf_bundle, "_drop_marks", lambda wb, path: (0, [])),
            patch.object(excel_watchdog, "_owned_process", lambda excel: (lambda: None, lambda: None)),
        ]
        for p in self.patches:
            p.start()

    def tearDown(self):
        for p in self.patches:
            p.stop()
        self.tmp.cleanup()

    def wait_done(self, key, names):
        for _ in range(200):
            state = excel_diff_pdf.status({key: names})[key]
            if state["state"] in ("done", "error"):
                return state
            time.sleep(0.02)
        self.fail("后台没导完")

    def test_exports_only_requested_and_reuses(self):
        first = excel_diff_pdf.request(str(self.book), ["表1", "藏起来"])
        done = self.wait_done(first["key"], ["表1", "藏起来"])
        self.assertEqual(done["state"], "done")
        self.assertEqual(done["sheets"]["表1"]["pages"], 2)
        self.assertIn("隐藏表", done["sheets"]["藏起来"]["error"])
        self.assertEqual(self.exported, ["表1"])
        png = excel_diff_pdf.page_png(first["key"], "表1", 1)
        self.assertTrue(png.startswith(b"\x89PNG"))
        again = excel_diff_pdf.request(str(self.book), ["表1"])
        self.assertEqual(again["state"], "done")  # 导过的直接给
        self.assertEqual(self.exported, ["表1"])

    def test_bad_page_requests(self):
        with self.assertRaises(ValueError):
            excel_diff_pdf.page_png("../x", "表1", 0)
        with self.assertRaises(ValueError):
            excel_diff_pdf.page_png("0" * 64, "表1", 0)

    def test_request_for_additional_sheet_while_exporting(self):
        entered, release = threading.Event(), threading.Event()
        export = pdf_bundle._export_sheet

        def pause_first(sheet, target):
            if sheet.Name == "表1":
                entered.set()
                self.assertTrue(release.wait(4))
            export(sheet, target)

        with patch.object(pdf_bundle, "_export_sheet", pause_first):
            first = excel_diff_pdf.request(str(self.book), ["表1"])
            try:
                self.assertTrue(entered.wait(4))
                second = excel_diff_pdf.request(str(self.book), ["表1", "表2"])
                self.assertEqual(first["key"], second["key"])
            finally:
                release.set()
            done = self.wait_done(first["key"], ["表1", "表2"])
        self.assertEqual(done["state"], "done")
        self.assertEqual(set(done["sheets"]), {"表1", "表2"})
        self.assertEqual(self.exported, ["表1", "表2"])


if __name__ == "__main__":
    unittest.main()
