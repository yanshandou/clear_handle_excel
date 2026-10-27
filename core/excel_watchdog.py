"""结束无进展的专用后台 Excel，让阻塞的 COM 调用返回异常。

持有进程句柄，而非稍后按 PID 查找，避免进程退出后 PID 被其他程序复用。
只允许注册调用方通过 DispatchEx 新建的专用 Excel 实例。
"""
import os
import threading
import time
from typing import Any, Callable


def _owned_process(excel: Any) -> tuple[Callable[[], None], Callable[[], None]]:
    import win32api
    import win32process

    _, pid = win32process.GetWindowThreadProcessId(int(excel.Hwnd))
    handle = win32api.OpenProcess(0x0001 | 0x00100000, False, pid)  # TERMINATE | SYNCHRONIZE
    return lambda: win32api.TerminateProcess(handle, 1), handle.Close


def idle_timeout() -> float:
    try:
        seconds = float(os.environ.get("CLIP_EXCEL_IDLE_TIMEOUT_SECONDS", "900"))
        return seconds if 60 <= seconds <= 86400 else 900.0
    except ValueError:
        return 900.0


class ExcelWatchdog:
    def __init__(self, on_timeout: Callable[[], None], timeout: float | None = None):
        self.timeout = idle_timeout() if timeout is None else timeout
        self.on_timeout = on_timeout
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._progress = time.monotonic()
        self._expired = False
        self._deadline: float | None = None
        self._deadline_hit = False
        self._processes: list[tuple[Callable[[], None], Callable[[], None]]] = []
        self._thread = threading.Thread(target=self._run, name="excel-watchdog", daemon=True)

    def start(self) -> None:
        self._thread.start()

    def progress(self) -> None:
        with self._lock:
            self._progress = time.monotonic()
            self._expired = False
            self._deadline_hit = False  # 开始下一件事了：上一份文件超时已经处理过

    def deadline(self, seconds: float | None) -> None:
        """一份文件扫描的时间上限：从现在起超过 seconds 秒，不管有没有进展都结束 Excel（deadline_hit 为真）。
        None 取消上限，但保留 deadline_hit，调用方在异常处理里还要看是不是超时"""
        with self._lock:
            if seconds is None:
                self._deadline = None
            else:
                self._deadline = time.monotonic() + seconds
                self._deadline_hit = False

    @property
    def deadline_hit(self) -> bool:
        with self._lock:
            return self._deadline_hit

    @property
    def timed_out(self) -> bool:
        with self._lock:
            return self._expired

    def register(self, excel: Any) -> None:
        process = _owned_process(excel)
        with self._lock:
            self._processes.append(process)
            self._progress = time.monotonic()

    def _run(self) -> None:
        while not self._stop.wait(min(1.0, self.timeout / 4)):
            expired = False
            with self._lock:
                now = time.monotonic()
                overdue = self._deadline is not None and now >= self._deadline
                if not self._expired and (overdue or now - self._progress >= self.timeout):
                    self._expired = True
                    if overdue:
                        self._deadline, self._deadline_hit = None, True
                    expired = bool(self._processes)
                    if expired:
                        # 结束 Excel 之前记下任务线程停在哪一行：卡在哪个调用上一看就知道
                        from core import debug_log
                        debug_log.dump_threads("看门狗：这份文件扫描超过时间上限，即将结束后台 Excel" if overdue
                                               else f"看门狗：后台 Excel {self.timeout:.0f} 秒没有进展，即将结束它")
                    for terminate, _ in self._processes:
                        try:
                            terminate()
                        except Exception:
                            pass  # 已正常退出的进程不需要再次结束
            if expired:
                try:
                    self.on_timeout()
                except Exception:
                    pass

    def stop(self) -> None:
        self._stop.set()
        self._thread.join(timeout=2)
        with self._lock:
            for _, close in self._processes:
                close()
            self._processes.clear()
