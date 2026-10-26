import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import batch, reporter


class FindRecordedWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.folder = Path(self.tmp.name)
        (self.folder / "sub").mkdir()
        (self.folder / "sub" / "a.xlsx").write_bytes(b"")
        (self.folder / "★b.xlsx").write_bytes(b"")
        (self.folder / "secret.txt").write_text("x")
        state = {
            "folder": str(self.folder),
            "files": {
                "sub/a.xlsx": {"name": "sub/a.xlsx"},
                "b.xlsx": {"name": "b.xlsx"},
                "../outside.xlsx": {"name": "../outside.xlsx"},
                "secret.txt": {"name": "secret.txt"},
                "gone.xlsx": {"name": "gone.xlsx"},
            },
        }
        patcher = patch.dict(reporter.EXECUTION_STATE, state)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_finds_file_inside_target_folder(self):
        path = reporter.find_recorded_workbook("sub/a.xlsx")
        self.assertEqual(path, (self.folder / "sub" / "a.xlsx").resolve())

    def test_finds_file_renamed_for_later_review(self):
        self.assertEqual(reporter.find_recorded_workbook("b.xlsx"), (self.folder / "★b.xlsx").resolve())

    def test_rejects_unknown_outside_or_non_excel_files(self):
        for name in ["other.xlsx", "../outside.xlsx", "secret.txt", "gone.xlsx"]:
            with self.subTest(name=name), self.assertRaises(ValueError):
                reporter.find_recorded_workbook(name)


class FindWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.folder = Path(self.tmp.name) / "target"
        (self.folder / "sub").mkdir(parents=True)
        (self.folder / "sub" / "a.xlsx").write_bytes(b"")
        (self.folder / "plain.xls").write_bytes(b"")
        (Path(self.tmp.name) / "outside.xlsx").write_bytes(b"")

    def test_relative_and_absolute_names(self):
        a = (self.folder / "sub" / "a.xlsx").resolve()
        self.assertEqual(reporter.find_workbook("sub/a.xlsx", [str(self.folder)]), a)
        self.assertEqual(reporter.find_workbook(str(a), ["", str(self.folder)]), a)

    def test_follows_star_prefix_both_ways(self):
        a = self.folder / "sub" / "a.xlsx"
        a.rename(a.with_name("★a.xlsx"))
        # 页面上还是旧名字，磁盘上已经加了 ★
        self.assertEqual(reporter.find_workbook(str(a), [str(self.folder)]), a.with_name("★a.xlsx").resolve())
        # 页面上带 ★（或旧版 🟡），磁盘上已经去掉
        self.assertEqual(reporter.find_workbook("🟡plain.xls", [str(self.folder)]), (self.folder / "plain.xls").resolve())

    def test_rejects_files_outside_folders(self):
        outside = Path(self.tmp.name) / "outside.xlsx"
        for name in [str(outside), "../outside.xlsx", "missing.xlsx"]:
            with self.subTest(name=name), self.assertRaises(ValueError):
                reporter.find_workbook(name, [str(self.folder)])
        with self.assertRaises(ValueError):
            reporter.find_workbook("sub/a.xlsx", [])


class StillNeedsTests(unittest.TestCase):
    """扫描 / 复检进行中点文件名：做完的文件放行，没轮到和正在处理的拦下"""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.a, self.b, self.c = (str(Path(self.tmp.name) / n) for n in ("a.xlsx", "b.xlsx", "c.xlsx"))

    def test_scan_lets_finished_and_unrelated_files_open(self):
        job = {"kind": "batch_scan", "workbooks": [self.a, self.b], "completed": [self.a]}
        self.assertFalse(batch.still_needs(job, Path(self.a)))
        self.assertTrue(batch.still_needs(job, Path(self.b)))
        self.assertFalse(batch.still_needs(job, Path(self.c)))

    def test_scan_blocks_when_finishing_touches_finished_files(self):
        job = {"kind": "batch_scan", "workbooks": [self.a], "completed": [self.a], "to_manual": True}
        self.assertTrue(batch.still_needs(job, Path(self.a)))
        job = {"kind": "batch_scan", "workbooks": [self.a], "completed": [self.a],
               "deferred_adjustments": {self.a: {}}}
        self.assertTrue(batch.still_needs(job, Path(self.a)))

    def test_recheck_blocks_current_and_queued_again(self):
        job = {"kind": "batch_recheck", "workbooks": [self.a, self.b], "completed": [self.a], "current": self.b}
        self.assertFalse(batch.still_needs(job, Path(self.a)))
        self.assertTrue(batch.still_needs(job, Path(self.b)))
        job["again"] = [self.a]
        self.assertTrue(batch.still_needs(job, Path(self.a)))

    def test_other_tasks_still_block(self):
        self.assertTrue(batch.still_needs({"kind": "batch_commit", "completed": []}, Path(self.a)))


if __name__ == "__main__":
    unittest.main()
