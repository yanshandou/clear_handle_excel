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


if __name__ == "__main__":
    unittest.main()
