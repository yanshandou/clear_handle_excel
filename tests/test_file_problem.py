import threading
import time
import unittest
from unittest.mock import Mock, patch

from core import file_problem
from core.excel_watchdog import ExcelWatchdog, is_dialog


class FileProblemTests(unittest.TestCase):
    def setUp(self):
        file_problem.AUTO_ANSWER = ""  # 真的等用户选择
        patcher = patch.object(file_problem, "_settings_cache", {"timeout_minutes": 0})
        patcher.start()
        self.addCleanup(patcher.stop)
        file_problem.begin("视图标准化")

    def tearDown(self):
        file_problem.finish()
        file_problem.AUTO_ANSWER = None

    def ask_in_thread(self, **kwargs):
        result = {}
        thread = threading.Thread(target=lambda: result.update(choice=file_problem.ask(
            "/data/a.xlsx", "a.xlsx", "要密码", folder="/data", **kwargs)))
        thread.start()
        for _ in range(100):
            if file_problem.pending():
                break
            time.sleep(0.01)
        return thread, result

    def test_waits_until_user_answers_retry(self):
        thread, result = self.ask_in_thread()
        problem = file_problem.pending()
        self.assertEqual(problem["task"], "视图标准化")
        self.assertEqual(problem["name"], "a.xlsx")
        self.assertTrue(thread.is_alive())
        file_problem.answer(problem["id"], "retry")
        thread.join(3)
        self.assertEqual(result["choice"], "retry")
        self.assertIsNone(file_problem.pending())

    def test_skip_all_applies_to_rest_of_this_task_only(self):
        thread, result = self.ask_in_thread()
        file_problem.answer(file_problem.pending()["id"], "skip", auto_skip=True)
        thread.join(3)
        self.assertEqual(result["choice"], "skip")
        self.assertEqual(file_problem.ask("b.xlsx", "b.xlsx", "坏了"), "skip")  # 不再等
        file_problem.begin("导出 PDF")
        thread, result = self.ask_in_thread()
        self.assertIsNotNone(file_problem.pending())  # 新任务又会问
        file_problem.answer(file_problem.pending()["id"], "skip")
        thread.join(3)

    def test_stale_answer_is_rejected(self):
        thread, _ = self.ask_in_thread()
        with self.assertRaises(ValueError):
            file_problem.answer("old", "retry")
        with self.assertRaises(ValueError):
            file_problem.answer(file_problem.pending()["id"], "delete")
        file_problem.answer(file_problem.pending()["id"], "skip")
        thread.join(3)

    def test_stop_request_counts_as_skip_and_tick_keeps_watchdog_alive(self):
        tick, stop = Mock(), threading.Event()
        thread, result = self.ask_in_thread(tick=tick, should_stop=stop.is_set)
        time.sleep(1.2)
        stop.set()
        thread.join(3)
        self.assertEqual(result["choice"], "skip")
        self.assertTrue(tick.called)

    def test_nobody_answers_within_timeout_skips(self):
        file_problem._settings_cache["timeout_minutes"] = 1
        clock = iter(range(0, 10000, 30))
        with patch.object(file_problem.time, "monotonic", side_effect=lambda: next(clock)):
            self.assertEqual(file_problem.ask("/data/a.xlsx", "a.xlsx", "要密码"), "skip")
        self.assertIsNone(file_problem.pending())

    def test_pending_reports_countdown_and_settings_are_validated(self):
        file_problem._settings_cache["timeout_minutes"] = 30
        thread, _ = self.ask_in_thread()
        problem = file_problem.pending()
        self.assertEqual(problem["timeout_minutes"], 30)
        self.assertGreater(problem["auto_skip_in"], 1790)
        for bad in (-1, 1441, 1.5, True, "5"):
            with self.assertRaises(ValueError):
                file_problem.set_settings(bad)
        file_problem.answer(problem["id"], "skip")
        thread.join(3)

    def test_settings_saved_and_reloaded(self):
        import tempfile
        from pathlib import Path
        from core import storage
        with tempfile.TemporaryDirectory() as folder, \
                patch.object(storage, "DATABASE_PATH", Path(folder) / "t.sqlite3"):
            file_problem.set_settings(45)
            file_problem._settings_cache = None
            self.assertEqual(file_problem.get_settings(), {"timeout_minutes": 45})

    def test_should_ask_only_for_file_problems(self):
        self.assertTrue(file_problem.should_ask(RuntimeError("COM 错误")))
        self.assertTrue(file_problem.should_ask(ValueError("文件仍在 Excel 中打开，请保存并关闭后继续")))
        self.assertFalse(file_problem.should_ask(ValueError("找不到唯一对应的备份文件")))

    def test_under_tests_without_override_skips_immediately(self):
        file_problem.AUTO_ANSWER = None
        self.assertEqual(file_problem.ask("a.xlsx", "a.xlsx", "坏了"), "skip")


class DialogWatchTests(unittest.TestCase):
    def test_only_dialog_windows_that_are_not_progress_count(self):
        self.assertTrue(is_dialog("bosa_sdm_XL9", "パスワード"))
        self.assertTrue(is_dialog("#32770", "Microsoft Excel"))
        self.assertFalse(is_dialog("#32770", "発行中"))
        self.assertFalse(is_dialog("XLMAIN", "Excel"))

    def test_dialog_that_stays_ends_excel_and_explains_why(self):
        killed, notify = threading.Event(), threading.Event()
        watchdog = ExcelWatchdog(notify.set, timeout=3600)
        with patch("core.excel_watchdog._owned_process", return_value=(killed.set, Mock(), 42, 7)), \
                patch("core.excel_watchdog.DIALOG_SECONDS", 0.05), \
                patch("core.excel_watchdog.find_dialogs", return_value=[(9, "パスワード：パスワードを入力してください")]) as find, \
                patch.object(watchdog._stop, "wait", side_effect=lambda t: time.sleep(0.03) or killed.is_set()):
            watchdog.register(object())
            watchdog._progress -= 1
            watchdog._run()
        self.assertTrue(notify.is_set())
        self.assertEqual(find.call_args[0], ({42}, {7}))
        message = watchdog.explain(RuntimeError("(-2147023170, 'RPC')"))
        self.assertIn("弹出了提示框", message)
        self.assertIn("パスワード", message)
        self.assertEqual(watchdog.explain(RuntimeError("x")), "后台 Excel 60 分钟没有反应，已结束后台 Excel（x）")

    def test_dialog_ignored_while_task_makes_progress(self):
        killed = Mock()
        watchdog = ExcelWatchdog(Mock(), timeout=3600)
        with patch("core.excel_watchdog._owned_process", return_value=(killed, Mock(), 42, 7)), \
                patch("core.excel_watchdog.find_dialogs", return_value=[(9, "x")]) as find:
            watchdog.register(object())
            self.assertEqual(watchdog._check_dialog(), "")
        find.assert_not_called()
        killed.assert_not_called()


class GuardSkipTests(unittest.TestCase):
    def test_user_skip_does_not_ask_and_drops_excel(self):
        from core import excel_watchdog
        guard = file_problem.Guard("批量导出同名 PDF")
        try:
            with patch("core.excel_watchdog._owned_process", return_value=(Mock(), Mock())):
                guard.watch(object())
            self.assertFalse(guard.file("a.xlsx"))
            self.assertTrue(excel_watchdog.request_skip("a.xlsx")["ok"])
            close = Mock()
            with patch.object(file_problem, "ask") as ask:
                retry, message = guard.problem(RuntimeError("RPC サーバーを利用できません"), "/d/a.xlsx", "a.xlsx", "/d",
                                               close=close)
            ask.assert_not_called()
            close.assert_called_once()
            self.assertFalse(retry)
            self.assertTrue(guard.closed)
            self.assertIn("跳过这一份", message)
        finally:
            guard.stop()
        self.assertFalse(excel_watchdog.current()["running"])

    def test_stopped_task_records_rest_as_not_done_without_asking(self):
        from core import excel_watchdog
        guard = file_problem.Guard("统一字体")
        try:
            excel_watchdog.request_stop()
            with self.assertRaises(excel_watchdog.TaskStopped) as caught:
                guard.file("b.xlsx")
            with patch.object(file_problem, "ask") as ask:
                retry, message = guard.problem(caught.exception, "/d/b.xlsx", "b.xlsx", "/d")
            ask.assert_not_called()
            self.assertFalse(retry)
            self.assertEqual(message, excel_watchdog.NOT_DONE_REASON)
        finally:
            guard.stop()

    def test_clean_chain_skips_later_steps_after_stop(self):
        from core import clean_chain, excel_watchdog
        with patch.object(clean_chain, "_chain_stop_seq", excel_watchdog.stop_seq()):
            self.assertFalse(clean_chain._stopped())
            excel_watchdog.request_stop()
            self.assertTrue(clean_chain._stopped())

    def test_slow_settings_validated(self):
        for bad in (-1, 1441, "5", True, 1.5):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                file_problem.set_slow_settings(bad)
        with patch("core.storage.save_snapshot") as save:
            self.assertEqual(file_problem.set_slow_settings(30), {"file_limit_minutes": 30})
        save.assert_called_once()
        self.assertEqual(file_problem.file_limit_seconds(), 1800)
        file_problem._slow_cache = None


if __name__ == "__main__":
    unittest.main()
