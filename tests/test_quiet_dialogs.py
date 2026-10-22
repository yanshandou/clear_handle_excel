import unittest
from unittest.mock import Mock

from core.quiet_dialogs import hide_progress_dialogs, is_progress_title


class QuietDialogTests(unittest.TestCase):
    def test_only_export_progress_titles_are_hidden(self):
        for title in ("発行中...", "Publishing...", "正在发布..."):
            self.assertTrue(is_progress_title(title))
        for title in ("Microsoft Excel", "", "保存しますか"):
            self.assertFalse(is_progress_title(title))

    def test_does_nothing_outside_windows(self):
        self.assertIsNone(hide_progress_dialogs(Mock(Hwnd=1)))


if __name__ == "__main__":
    unittest.main()
