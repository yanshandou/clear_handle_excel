import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import reporter


class OpenRecordedWorkbookTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self.tmp.name)
        (self.folder / "sub").mkdir()
        (self.folder / "sub" / "a.xlsx").write_bytes(b"")
        (self.folder / "secret.txt").write_text("x")
        state = {
            "folder": str(self.folder),
            "files": {
                "sub/a.xlsx": {"name": "sub/a.xlsx"},
                "../outside.xlsx": {"name": "../outside.xlsx"},
                "secret.txt": {"name": "secret.txt"},
                "gone.xlsx": {"name": "gone.xlsx"},
            },
        }
        patcher = patch.dict(reporter.EXECUTION_STATE, state)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(self.tmp.cleanup)
        startfile = patch.object(reporter.os, "startfile", create=True)
        self.startfile = startfile.start()
        self.addCleanup(startfile.stop)

    def test_opens_file_inside_target_folder(self):
        path = reporter.open_recorded_workbook("sub/a.xlsx")
        self.assertEqual(path, (self.folder / "sub" / "a.xlsx").resolve())
        self.startfile.assert_called_once_with(path)

    def test_rejects_unknown_outside_or_non_excel_files(self):
        for name in ["other.xlsx", "../outside.xlsx", "secret.txt", "gone.xlsx"]:
            with self.subTest(name=name), self.assertRaises(ValueError):
                reporter.open_recorded_workbook(name)
        self.startfile.assert_not_called()


if __name__ == "__main__":
    unittest.main()
