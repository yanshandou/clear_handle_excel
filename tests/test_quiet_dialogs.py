import unittest
from unittest.mock import Mock

from core.quiet_dialogs import corner_rect, hide_progress_dialogs, is_progress_title, release_excel


class QuietDialogTests(unittest.TestCase):
    def test_only_export_progress_titles_are_hidden(self):
        for title in ("発行中...", "Publishing...", "正在发布..."):
            self.assertTrue(is_progress_title(title))
        for title in ("Microsoft Excel", "", "保存しますか"):
            self.assertFalse(is_progress_title(title))

    def test_does_nothing_outside_windows(self):
        self.assertIsNone(hide_progress_dialogs(Mock(Hwnd=1)))
        release_excel(Mock(Hwnd=1))

    def test_parent_window_goes_to_bottom_right_of_work_area(self):
        # 1920x1080 屏幕、任务栏高 40
        self.assertEqual(corner_rect((0, 0, 1920, 1040)), (1600, 920, 320, 120))
        # 副屏在左侧时坐标为负也照样贴右下角
        self.assertEqual(corner_rect((-1280, 0, 0, 1024)), (-320, 904, 320, 120))


if __name__ == "__main__":
    unittest.main()
