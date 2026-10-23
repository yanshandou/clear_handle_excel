import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import common


class WorkbookLockTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "a.xlsx"
        self.path.write_bytes(b"book")
        self.lock = self.path.with_name("~$a.xlsx")

    def tearDown(self):
        self.tmp.cleanup()

    def test_no_lock_file_is_not_locked(self):
        self.assertFalse(common.workbook_locked(self.path))

    def test_lock_file_with_file_held_open_is_locked(self):
        self.lock.write_bytes(b"owner")
        with patch.object(common, "_held_open", return_value=True):
            self.assertTrue(common.workbook_locked(self.path))
        self.assertTrue(self.lock.exists())

    def test_stale_lock_file_is_removed(self):
        self.lock.write_bytes(b"owner")
        with patch.object(common, "_held_open", return_value=False):
            self.assertFalse(common.workbook_locked(self.path))
        self.assertFalse(self.lock.exists())

    def test_held_open_on_windows_uses_sharing_violation(self):
        with patch.object(common.sys, "platform", "win32"):
            self.assertFalse(common._held_open(self.path))
            with patch.object(common.os, "open", side_effect=PermissionError):
                self.assertTrue(common._held_open(self.path))


if __name__ == "__main__":
    unittest.main()
