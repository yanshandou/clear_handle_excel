import threading
import unittest
from unittest.mock import Mock, patch

from core.excel_watchdog import ExcelWatchdog, idle_timeout


class WatchdogTests(unittest.TestCase):
    def test_timeout_terminates_only_registered_process_and_releases_handle(self):
        killed = threading.Event()
        close, notify = Mock(), threading.Event()
        watchdog = ExcelWatchdog(notify.set, timeout=.04)
        with patch("core.excel_watchdog._owned_process", return_value=(killed.set, close)) as owned:
            app = object()
            watchdog.register(app)
            watchdog.start()
            try:
                self.assertTrue(killed.wait(2))
                self.assertTrue(notify.wait(2))
                owned.assert_called_once_with(app)
            finally:
                watchdog.stop()
        close.assert_called_once()

    def test_reported_progress_postpones_deadline(self):
        killed = Mock()
        watchdog = ExcelWatchdog(Mock(), timeout=100)
        with patch("core.excel_watchdog._owned_process", return_value=(killed, Mock())):
            watchdog.register(object())
            watchdog._progress = 0
            with patch("core.excel_watchdog.time.monotonic", return_value=50):
                watchdog.progress()
            with patch("core.excel_watchdog.time.monotonic", return_value=120), \
                    patch.object(watchdog._stop, "wait", side_effect=[False, True]):
                watchdog._run()
        killed.assert_not_called()

    def test_new_progress_rearms_timeout_for_next_excel_instance(self):
        first, second = threading.Event(), threading.Event()
        watchdog = ExcelWatchdog(Mock(), timeout=.04)
        with patch("core.excel_watchdog._owned_process", side_effect=[(first.set, Mock()), (second.set, Mock())]):
            watchdog.register(object())
            watchdog.start()
            try:
                self.assertTrue(first.wait(2))
                watchdog.progress()
                watchdog.register(object())
                self.assertTrue(second.wait(2))
            finally:
                watchdog.stop()

    def test_file_deadline_ends_excel_even_while_progress_is_reported(self):
        killed = Mock()
        watchdog = ExcelWatchdog(Mock(), timeout=100)
        with patch("core.excel_watchdog._owned_process", return_value=(killed, Mock())):
            watchdog.register(object())
            with patch("core.excel_watchdog.time.monotonic", return_value=0):
                watchdog.deadline(30)
            with patch("core.excel_watchdog.time.monotonic", return_value=20):
                watchdog.progress()  # 有进展也不顶用：一份文件总共只给 30 秒
            with patch("core.excel_watchdog.time.monotonic", return_value=31), \
                    patch.object(watchdog._stop, "wait", side_effect=[False, True]):
                watchdog._run()
        killed.assert_called_once()
        watchdog.deadline(None)
        self.assertTrue(watchdog.deadline_hit)  # 取消上限后，异常处理里还看得到是超时
        watchdog.progress()
        self.assertFalse(watchdog.deadline_hit)  # 开始下一件事就清掉

    def test_cancelled_deadline_does_not_end_excel(self):
        killed = Mock()
        watchdog = ExcelWatchdog(Mock(), timeout=100)
        with patch("core.excel_watchdog._owned_process", return_value=(killed, Mock())):
            watchdog.register(object())
            with patch("core.excel_watchdog.time.monotonic", return_value=0):
                watchdog.deadline(30)
                watchdog.progress()
            watchdog.deadline(None)
            with patch("core.excel_watchdog.time.monotonic", return_value=31), \
                    patch.object(watchdog._stop, "wait", side_effect=[False, True]):
                watchdog._run()
        killed.assert_not_called()
        self.assertFalse(watchdog.deadline_hit)

    def test_timeout_configuration_has_safe_default_and_bounds(self):
        for value in ["bad", "nan", "inf", "0", "-1"]:
            with self.subTest(value=value), patch.dict("os.environ", CLIP_EXCEL_IDLE_TIMEOUT_SECONDS=value):
                self.assertEqual(idle_timeout(), 900)
        with patch.dict("os.environ", CLIP_EXCEL_IDLE_TIMEOUT_SECONDS="1800"):
            self.assertEqual(idle_timeout(), 1800)
