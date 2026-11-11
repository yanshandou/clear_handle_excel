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


class SkipThisFileTests(unittest.TestCase):
    """进度框「跳过这一份」：结束正在做这份的后台 Excel，原因只给这一份用一次"""

    def make(self):
        from core import excel_watchdog
        killed = Mock()
        watchdog = ExcelWatchdog(Mock(), timeout=3600)
        patcher = patch("core.excel_watchdog._owned_process", return_value=(killed, Mock()))
        patcher.start()
        self.addCleanup(patcher.stop)
        watchdog.register(object())
        with excel_watchdog._ACTIVE_LOCK:
            excel_watchdog._ACTIVE.append(watchdog)
        self.addCleanup(lambda: excel_watchdog._ACTIVE.remove(watchdog) if watchdog in excel_watchdog._ACTIVE else None)
        return excel_watchdog, watchdog, killed

    def test_skip_kills_current_excel_and_explains_once(self):
        mod, watchdog, killed = self.make()
        self.assertFalse(watchdog.file("a.xlsx"))
        self.assertEqual(mod.current()["label"], "a.xlsx")
        result = mod.request_skip("a.xlsx")
        self.assertTrue(result["ok"])
        killed.assert_called_once()
        self.assertIn("跳过这一份", watchdog.explain(RuntimeError("RPC")))
        self.assertIn("跳过这一份", watchdog.take_skip())
        self.assertEqual(watchdog.take_skip(), "")

    def test_skip_of_finished_file_is_refused(self):
        mod, watchdog, killed = self.make()
        watchdog.file("b.xlsx")
        result = mod.request_skip("a.xlsx")
        self.assertFalse(result["ok"])
        self.assertIn("已经做完", result["reason"])
        killed.assert_not_called()

    def test_not_skippable_while_saving(self):
        mod, watchdog, killed = self.make()
        watchdog.file("a.xlsx")
        watchdog.file(None)
        self.assertFalse(mod.current()["skippable"])
        self.assertFalse(mod.request_skip("a.xlsx")["ok"])
        killed.assert_not_called()

    def test_skip_while_not_in_excel_tells_next_file_to_restart_excel(self):
        mod, watchdog, _killed = self.make()
        watchdog.file("a.xlsx")
        mod.request_skip("")
        self.assertTrue(watchdog.file("b.xlsx"))  # a 没报错就做完了：b 要换新的 Excel
        self.assertFalse(watchdog.file("c.xlsx"))

    def test_file_limit_skips_slow_file_automatically(self):
        _mod, watchdog, killed = self.make()
        watchdog.file_limit = lambda: 60
        with patch("core.excel_watchdog.time.monotonic", return_value=0):
            watchdog.file("slow.xlsx")
        with patch("core.excel_watchdog.time.monotonic", return_value=30):
            watchdog._check_file_limit()
        killed.assert_not_called()
        with patch("core.excel_watchdog.time.monotonic", return_value=61):
            watchdog._check_file_limit()
        killed.assert_called_once()
        self.assertIn("每份最多 1 分钟", watchdog.take_skip())

    def test_zero_limit_never_skips(self):
        _mod, watchdog, killed = self.make()
        watchdog.file_limit = lambda: 0
        with patch("core.excel_watchdog.time.monotonic", return_value=0):
            watchdog.file("slow.xlsx")
        with patch("core.excel_watchdog.time.monotonic", return_value=99999):
            watchdog._check_file_limit()
        killed.assert_not_called()

    def test_stop_ends_current_file_and_refuses_next(self):
        from core.excel_watchdog import TaskStopped
        mod, watchdog, killed = self.make()
        watchdog.file("a.xlsx")
        before = mod.stop_seq()
        result = mod.request_stop()
        self.assertEqual(result["label"], "a.xlsx")
        killed.assert_called_once()
        self.assertEqual(mod.stop_seq(), before + 1)
        self.assertTrue(watchdog.stop_requested)
        self.assertIn("结束整个任务", watchdog.take_skip())
        with self.assertRaises(TaskStopped):
            watchdog.file("b.xlsx")
        self.assertFalse(watchdog.file(None))  # 不能跳过的一步照常
