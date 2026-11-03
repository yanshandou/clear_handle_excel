import threading
import time
import unittest
from unittest import mock

from core import lock_timing


class TimedRLockTests(unittest.TestCase):
    def test_reentrant_and_nonblocking_acquire_still_work(self):
        lock = lock_timing.TimedRLock("测试")
        with lock:
            with lock:
                self.assertTrue(lock.acquire(blocking=False))
                lock.release()
        got = []
        worker = threading.Thread(target=lambda: got.append(lock.acquire(blocking=False)))
        with lock:
            worker.start()
            worker.join()
        self.assertEqual(got, [False])
        self.assertTrue(lock.acquire(blocking=False))
        lock.release()

    def test_logs_who_waited_and_who_held_the_lock(self):
        lock = lock_timing.TimedRLock("测试")
        notes = []
        held = threading.Event()

        def holder():
            with lock:
                held.set()
                time.sleep(0.15)

        with mock.patch.object(lock_timing, "WAIT_NOTE_SECONDS", 0.05), \
                mock.patch.object(lock_timing, "HOLD_NOTE_SECONDS", 0.1), \
                mock.patch.object(lock_timing.TimedRLock, "_note", staticmethod(lambda tag, **f: notes.append((tag, f)))):
            worker = threading.Thread(target=holder, name="占锁线程")
            worker.start()
            held.wait(2)
            with lock:
                pass
            worker.join()
        tags = [tag for tag, _ in notes]
        self.assertIn("⏳ 等锁", tags)
        self.assertIn("🔒 占锁久", tags)
        wait = dict(notes)["⏳ 等锁"]
        self.assertEqual(wait["占着的线程"], "占锁线程")
        self.assertIn("holder", wait["占着的位置"])

    def test_quick_locks_log_nothing(self):
        lock = lock_timing.TimedRLock("测试")
        with mock.patch.object(lock_timing.TimedRLock, "_note") as note:
            for _ in range(100):
                with lock:
                    pass
        note.assert_not_called()


class DeferredPersistTests(unittest.TestCase):
    def test_saves_once_per_file_but_forced_saves_go_out_now(self):
        from core import workflow
        saves = []
        with mock.patch.object(workflow, "save_snapshot", lambda *a, **k: saves.append(1)):
            with workflow.deferred_persist():
                workflow._persist_state()
                workflow._persist_state()
                self.assertEqual(saves, [])
                workflow._persist_state(force=True)
                self.assertEqual(saves, [1])
                other = threading.Thread(target=workflow._persist_state)
                other.start()
                other.join()
                self.assertEqual(len(saves), 2)  # 别的线程（网页）照常马上存
            self.assertEqual(len(saves), 3)  # 文件做完存一次
            with workflow.deferred_persist():
                pass
            self.assertEqual(len(saves), 3)  # 没改动不存

    def test_saves_even_when_the_file_fails(self):
        from core import workflow
        saves = []
        with mock.patch.object(workflow, "save_snapshot", lambda *a, **k: saves.append(1)):
            with self.assertRaises(RuntimeError):
                with workflow.deferred_persist():
                    workflow._persist_state()
                    raise RuntimeError("Excel 崩溃")
        self.assertEqual(saves, [1])


if __name__ == "__main__":
    unittest.main()
