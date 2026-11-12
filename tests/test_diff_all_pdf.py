"""一页对比（/excel-diff/all）并排看同名 PDF：只给任务两边已有的 aaa.pdf，路径不能跳出任务文件夹"""
import os
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from core import distribute, work_task


class DiffAllPdfTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        self.work, self.shared = root / "work", root / "shared"
        for folder in (self.work / "sub", self.shared / "sub"):
            folder.mkdir(parents=True)
        self.task = {"id": "t1", "work": str(self.work), "shared": str(self.shared)}
        self.get = patch.object(work_task, "get", lambda task_id: self.task if task_id == "t1" else (_ for _ in ()).throw(ValueError("找不到")))
        self.get.start()

    def tearDown(self):
        self.get.stop()
        self.tmp.cleanup()

    def test_paths_and_stale(self):
        (self.shared / "sub" / "a.xlsx").write_bytes(b"x")
        (self.shared / "sub" / "a.pdf").write_bytes(b"%PDF")
        (self.work / "sub" / "a.xlsx").write_bytes(b"x")
        old = time.time() - 100
        os.utime(self.shared / "sub" / "a.pdf", (old, old))
        self.assertEqual(distribute.diff_pdf("t1", "sub/a.xlsx", "theirs"), self.shared / "sub" / "a.pdf")
        self.assertEqual(distribute.diff_pdf("t1", "sub/a.xlsx", "mine"), self.work / "sub" / "a.pdf")
        info = distribute.diff_pdfs("t1", "sub/a.xlsx")
        self.assertTrue(info["theirs"]["exists"])
        self.assertTrue(info["theirs"]["stale"])  # PDF 比 Excel 旧
        self.assertFalse(info["mine"]["exists"])

    def test_rejects_bad_input(self):
        for rel in ("../x.xlsx", "sub/a.pdf", "C:/a.xlsx", ""):
            with self.assertRaises(ValueError):
                distribute.diff_pdf("t1", rel, "mine")
        with self.assertRaises(ValueError):
            distribute.diff_pdf("t1", "sub/a.xlsx", "other")


if __name__ == "__main__":
    unittest.main()
