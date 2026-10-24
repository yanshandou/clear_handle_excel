import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import debug_log


class com_error(Exception):  # 和 pywintypes.com_error 同名、同样的 args 结构
    pass


class DebugLogTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        for name, value in (("DEBUG_DIR", root), ("APP_LOG", root / "app.log"), ("_under_test", lambda: False)):
            p = patch.object(debug_log, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(debug_log.end_task, "测试结束")

    def text(self) -> str:
        return debug_log.latest_log().read_text(encoding="utf-8")

    def test_task_file_has_version_options_and_end(self):
        debug_log.start_task("batch_scan", {"folder": "C:/x", "card_ids": [str(i) for i in range(50)]})
        debug_log.event("检出", 位置="Sheet1!B5", 行高=13.5)
        debug_log.end_task("completed")
        text = self.text()
        self.assertIn("程序版本:", text)
        self.assertIn("…共 50 项", text)  # 很长的列表只记数量
        self.assertIn("位置=Sheet1!B5 行高=13.50", text)
        self.assertIn("任务结束：completed", text)

    def test_clip_keeps_only_a_few_characters(self):
        self.assertEqual(debug_log.clip("短"), "'短'")
        clipped = debug_log.clip("株式会社テスト商事 東京都千代田区丸の内")
        self.assertIn("(共 20 字)", clipped)
        self.assertNotIn("丸の内", clipped)

    def test_com_error_code_and_traceback_written_once(self):
        debug_log.start_task("batch_commit")
        try:
            try:
                raise com_error(-2147352567, "例外が発生しました。", (0, "Microsoft Excel", "保存できません", None, 0, -2146827284), None)
            except com_error as inner:
                raise ValueError("保存失败") from inner
        except ValueError:
            debug_log.mirror("❌ book.xlsx 保存失败")
            debug_log.mirror("❌ 任务中断")  # 同一个异常第二次记录：不再重复写 traceback
        text = self.text()
        self.assertIn("HRESULT=0x80020009", text)
        self.assertIn("scode=0x800A03EC", text)
        self.assertIn("说明=保存できません", text)
        self.assertEqual(text.count("!! 异常"), 1)
        self.assertIn("The above exception was the direct cause", text)  # 内层的 COM 错误也在

    def test_info_lines_do_not_dump_traceback(self):
        debug_log.start_task("batch_scan")
        try:
            raise KeyError("x")
        except KeyError:
            debug_log.mirror("ℹ️ 沿用上次检测结果")
        self.assertNotIn("Traceback", self.text())

    def test_timed_records_seconds_and_failure(self):
        debug_log.start_task("batch_scan")
        with debug_log.timed("导出", 表="Sheet1") as extra:
            extra["页"] = 3
        with self.assertRaises(RuntimeError):
            with debug_log.timed("标色"):
                raise RuntimeError("Excel 崩溃")
        text = self.text()
        self.assertRegex(text, r"◀ 导出 秒=\d+\.\d\d 表=Sheet1 页=3")
        self.assertIn("✖ 标色 失败", text)
        self.assertIn("RuntimeError: Excel 崩溃", text)

    def test_thread_dump_shows_where_each_thread_is(self):
        debug_log.start_task("batch_unhighlight")
        debug_log.dump_threads("测试")
        text = self.text()
        self.assertIn("线程现场：测试", text)
        self.assertIn("test_thread_dump_shows_where_each_thread_is", text)

    def test_copy_keeps_head_errors_and_tail_of_a_big_log(self):
        debug_log.start_task("batch_scan")
        for i in range(6000):
            debug_log.event("检出", 序号=i, 填充="x" * 40)
            if i == 3000:
                try:
                    raise OSError("磁盘满了")
                except OSError as err:
                    debug_log.exception(err, "保存")
        debug_log.end_task("completed")
        copied = debug_log.for_copy(debug_log.latest_log())
        self.assertLess(len(copied.encode("utf-8")), 220_000)
        self.assertIn("程序版本:", copied)
        self.assertIn("OSError: 磁盘满了", copied)
        self.assertIn("序号=5999", copied)
        self.assertNotIn("序号=1500 ", copied)

    def test_only_recent_task_logs_are_kept(self):
        for i in range(debug_log.KEEP_TASK_LOGS + 5):
            (debug_log.DEBUG_DIR / f"task_20260101_0000{i:02d}_scan.log").write_text("x", encoding="utf-8")
        debug_log.start_task("batch_scan")
        self.assertEqual(len(debug_log.task_logs()), debug_log.KEEP_TASK_LOGS)

    def test_without_task_lines_go_to_app_log(self):
        debug_log.event("网页操作", 路径="/api/clean/preview")
        self.assertIn("/api/clean/preview", debug_log.APP_LOG.read_text(encoding="utf-8"))

    def test_logging_never_breaks_the_caller(self):
        class Broken:
            def __str__(self):
                raise RuntimeError("不能转成文字")
        debug_log.start_task("batch_scan")
        debug_log.event("坏数据", 值=Broken())  # 不抛异常就行


if __name__ == "__main__":
    unittest.main()
