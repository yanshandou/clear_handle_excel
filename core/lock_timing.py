"""带计时的可重入锁：等锁或占锁太久时写一行调试日志，排查「扫描中途莫名空等几秒」用（用户 2026-11-03 要求详细时间日志）。

日志里写清楚：谁在等、等了几秒；当时谁占着锁、在哪个函数里拿的锁、已经占了多久。
只在最外层拿锁 / 放锁时计时（RLock 同一线程重入不算）。"""

from __future__ import annotations

import sys
import threading
import time
from typing import Any

WAIT_NOTE_SECONDS = 0.5  # 等锁超过这么久记一行
HOLD_NOTE_SECONDS = 1.0  # 一次占锁超过这么久记一行


def _caller(depth: int) -> str:
    """拿锁的位置：跳过 with / acquire 本身，记「文件名:行号 函数名」，往上再带一层调用方"""
    try:
        frame = sys._getframe(depth)
    except ValueError:
        return "?"
    parts = []
    while frame is not None and len(parts) < 2:
        name = frame.f_code.co_filename.replace("\\", "/").rsplit("/", 1)[-1]
        if name not in {"lock_timing.py", "threading.py", "contextlib.py"}:
            parts.append(f"{name}:{frame.f_lineno} {frame.f_code.co_name}")
        frame = frame.f_back
    return " ← ".join(parts) or "?"


def _thread_name() -> str:
    name = threading.current_thread().name
    return "网页" if name.endswith("(process_request_thread)") else name


class TimedRLock:
    def __init__(self, label: str) -> None:
        self.label = label
        self._lock = threading.RLock()
        self._depth: dict[int, int] = {}
        self._holder: tuple[str, str, float] | None = None  # (线程, 位置, 开始占锁的时间)

    def acquire(self, blocking: bool = True, timeout: float = -1) -> bool:
        me = threading.get_ident()
        if self._depth.get(me):
            got = self._lock.acquire(blocking, timeout)
            if got:
                self._depth[me] += 1
            return got
        holder = self._holder
        started = time.perf_counter()
        got = self._lock.acquire(blocking, timeout)
        if not got:
            return False
        waited = time.perf_counter() - started
        where = _caller(2)
        self._depth[me] = 1
        self._holder = (_thread_name(), where, time.perf_counter())
        if waited >= WAIT_NOTE_SECONDS:
            self._note("⏳ 等锁", 锁=self.label, 秒=round(waited, 2), 等的线程=_thread_name(), 等的位置=where,
                       占着的线程=holder[0] if holder else "?", 占着的位置=holder[1] if holder else "?")
        return True

    def release(self) -> None:
        me = threading.get_ident()
        depth = self._depth.get(me, 0)
        if depth <= 1:
            self._depth.pop(me, None)
            holder, self._holder = self._holder, None
            self._lock.release()
            if holder is not None:
                held = time.perf_counter() - holder[2]
                if held >= HOLD_NOTE_SECONDS:
                    self._note("🔒 占锁久", 锁=self.label, 秒=round(held, 2), 线程=holder[0], 位置=holder[1])
            return
        self._depth[me] = depth - 1
        self._lock.release()

    def __enter__(self) -> bool:
        return self.acquire()

    def __exit__(self, *_exc: Any) -> None:
        self.release()

    @staticmethod
    def _note(tag: str, **fields: Any) -> None:
        try:
            from core import debug_log
            debug_log.event(tag, **fields)
        except Exception:
            pass  # 记日志出错不能影响拿锁放锁
