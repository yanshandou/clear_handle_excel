import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import patch

from core import debug_log


class com_error(Exception):  # 和 pywintypes.com_error 同名、同样的 args 结构
    pass


class DebugLogTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "log.txt"
        for name, value in (("LOG_PATH", self.path), ("_under_test", lambda: False), ("_fh", None)):
            p = patch.object(debug_log, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.addCleanup(self.close)

    def close(self):
        debug_log.end_task("测试结束")
        if debug_log._fh is not None:
            debug_log._fh.close()
            debug_log._fh = None

    def lines(self) -> list[str]:
        return self.path.read_text(encoding="utf-8").splitlines()

    def test_one_line_each_emoji_first_and_wave_blocks(self):
        debug_log.start_task("batch_scan", {"folder": "C:/x", "card_ids": [str(i) for i in range(50)]})
        debug_log.begin("扫描 1 / 2：a.xlsx")
        debug_log.event("🔍 检出", 位置="Sheet1!B5", 行高=13.5)
        debug_log.begin("行 Sheet1 第 5 行", 2)
        debug_log.event("🧪 试算", 次=1)
        debug_log.begin("扫描 2 / 2：b.xlsx")  # 同级：前面两个块先自动收尾
        debug_log.end_task("completed")
        lines = self.lines()
        waves = [line for line in lines if line.startswith("〰")]
        self.assertRegex(waves[0], r"^〰{10} ⬇任务 batch_scan #\d{8}_\d{6} 开始⬇ \d{4}-\d\d-\d\d \d\d:\d\d:\d\d\.\d{3} 〰{10}$")
        self.assertEqual([w.split(" ")[1] for w in waves],
                         [w.split(" ")[1] for w in waves])  # 每条分隔线都是一行
        order = [w.split(" ")[1][1:] for w in waves]
        self.assertEqual(order[1:], ["扫描", "行", "行", "扫描", "扫描", "扫描", "任务"])
        self.assertTrue(waves[-1].startswith("〰〰〰〰〰〰〰〰〰〰 ⬆任务 batch_scan"))
        for line in lines:
            if not line.startswith("〰"):
                self.assertRegex(line, r"^\S+ \[\d{4}-\d\d-\d\d \d\d:\d\d:\d\d\.\d{3}\] \[")
        text = "\n".join(lines)
        self.assertIn("🧭 [", text)
        self.assertIn("…共 50 项", text)  # 很长的列表只记数量
        self.assertIn("检出 位置=Sheet1!B5 行高=13.50", text)
        self.assertIn("任务结束：completed", text)

    def test_clip_keeps_only_a_few_characters(self):
        self.assertEqual(debug_log.clip("短"), "'短'")
        clipped = debug_log.clip("株式会社テスト商事 東京都千代田区丸の内")
        self.assertIn("(共 20 字)", clipped)
        self.assertNotIn("丸の内", clipped)

    def test_com_error_with_cause_on_one_line_written_once(self):
        debug_log.start_task("batch_commit")
        try:
            try:
                raise com_error(-2147352567, "例外が発生しました。",
                                (0, "Microsoft Excel", "保存できません", None, 0, -2146827284), None)
            except com_error as inner:
                raise ValueError("保存失败") from inner
        except ValueError:
            debug_log.mirror("❌ book.xlsx 保存失败")
            debug_log.mirror("❌ 任务中断")  # 同一个异常第二次记录：不再重复写
        errors = [line for line in self.lines() if line.startswith("❌")]
        self.assertEqual(len(errors), 1)
        self.assertIn("HRESULT=0x80020009", errors[0])
        self.assertIn("scode=0x800A03EC", errors[0])
        self.assertIn("说明=保存できません", errors[0])
        self.assertIn("起因: com_error", errors[0])
        self.assertIn("test_debug_log.py:", errors[0])

    def test_info_lines_do_not_dump_stack(self):
        debug_log.start_task("batch_scan")
        try:
            raise KeyError("x")
        except KeyError:
            debug_log.mirror("ℹ️ 沿用上次检测结果")
        self.assertFalse(any(line.startswith("❌") for line in self.lines()))
        self.assertTrue(any(line.startswith("💬") and "沿用上次" in line for line in self.lines()))

    def test_timed_records_seconds_and_failure(self):
        debug_log.start_task("batch_scan")
        with debug_log.timed("导出", 表="Sheet1") as extra:
            extra["页"] = 3
        with self.assertRaises(RuntimeError):
            with debug_log.timed("标色"):
                raise RuntimeError("Excel 崩溃")
        text = self.path.read_text(encoding="utf-8")
        self.assertRegex(text, r"⏱ \[.*\] 导出 秒=\d+\.\d\d 表=Sheet1 页=3")
        self.assertIn("标色 失败", text)
        self.assertIn("RuntimeError: Excel 崩溃", text)
        self.assertIn("⬆标色 结束⬆", text)

    def test_thread_dump_one_line_per_thread(self):
        debug_log.start_task("batch_unhighlight")
        debug_log.dump_threads("测试")
        lines = [line for line in self.lines() if line.startswith("🧵")]
        self.assertTrue(any("test_thread_dump_one_line_per_thread" in line for line in lines))

    def test_old_lines_pruned_but_running_task_kept(self):
        old = (datetime.now() - timedelta(hours=13)).strftime("%Y-%m-%d %H:%M:%S.000")
        recent = (datetime.now() - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S.000")
        self.path.write_text(f"💬 [{old}] [任务] 很早\n💬 [{recent}] [任务] 不久前\n", encoding="utf-8")
        debug_log.prune()
        self.assertEqual(len(self.lines()), 1)
        self.assertIn("不久前", self.lines()[0])
        debug_log.start_task("batch_scan")
        debug_log.write("任务里")
        debug_log.prune(datetime.now() + timedelta(hours=20))  # 任务跑了很久：开始以来的都留着
        self.assertIn("⬇任务 batch_scan", self.lines()[0])
        self.assertIn("任务里", self.lines()[-1])

    def test_copy_picks_one_task_and_trims_big_ones(self):
        debug_log.start_task("clean_sheets")
        debug_log.write("第一次任务")
        debug_log.end_task("completed")
        with patch.object(debug_log, "datetime", wraps=datetime) as fake:
            fake.now.return_value = datetime.now() + timedelta(seconds=2)  # 任务编号到秒，不能撞
            debug_log.start_task("batch_scan")
        for i in range(6000):
            debug_log.event("🔍 检出", 序号=i, 填充="x" * 40)
            if i == 3000:
                try:
                    raise OSError("磁盘满了")
                except OSError as err:
                    debug_log.exception(err, "保存")
        debug_log.end_task("completed")
        tasks = debug_log.tasks()
        self.assertEqual([t["kind"] for t in tasks], ["batch_scan", "clean_sheets"])
        title, text = debug_log.task_text(tasks[1]["id"])
        self.assertIn("第一次任务", text)
        self.assertNotIn("序号=", text)
        copied = debug_log.for_copy(*debug_log.task_text())
        self.assertLess(len(copied.encode("utf-8")), 220_000)
        self.assertIn("程序版本:", copied)
        self.assertIn("OSError: 磁盘满了", copied)
        self.assertIn("序号=5999", copied)
        self.assertNotIn("序号=1500 ", copied)
        self.assertNotIn("第一次任务", copied)

    def test_side_task_is_its_own_task_or_nests_in_running_one(self):
        with debug_log.side_task("find_marks", {"folder": "C:/x"}):
            debug_log.begin("查找残留 1 / 1：a.xlsx")
            debug_log.event("🔍 残留", 位置="Sheet1!B5")
        self.assertEqual([t["kind"] for t in debug_log.tasks()], ["find_marks"])
        self.assertIn("⬆查找残留 1 / 1：a.xlsx 结束⬆", self.path.read_text(encoding="utf-8"))

        import threading
        debug_log.start_task("batch_scan")
        debug_log.begin("扫描 1 / 2：b.xlsx")

        def other():
            with debug_log.side_task("compare_old"):
                debug_log.begin("对比旧版 1 / 1：c.xlsx")  # 另一个线程：不能把任务线程的块收掉
        thread = threading.Thread(target=other)
        thread.start()
        thread.join()
        debug_log.event("🔍 检出", 位置="Sheet1!A1")
        debug_log.end_task("completed")
        _, text = debug_log.task_text()
        self.assertIn("compare_old（和正在跑的任务同时进行） 开始", text)
        self.assertIn("📂 [", text)
        lines = text.splitlines()
        close = next(i for i, line in enumerate(lines) if "⬆扫描 1 / 2：b.xlsx" in line)
        self.assertLess(next(i for i, line in enumerate(lines) if "检出 位置=Sheet1!A1" in line), close)
        self.assertEqual([t["kind"] for t in debug_log.tasks()], ["batch_scan", "find_marks"])

    def test_without_task_lines_still_go_to_log(self):
        debug_log.event("🖱️ 网页操作", 路径="/api/clean/preview")
        self.assertIn("/api/clean/preview", self.path.read_text(encoding="utf-8"))

    def test_logging_never_breaks_the_caller(self):
        class Broken:
            def __str__(self):
                raise RuntimeError("不能转成文字")
        debug_log.start_task("batch_scan")
        debug_log.event("坏数据", 值=Broken())  # 不抛异常就行


if __name__ == "__main__":
    unittest.main()
