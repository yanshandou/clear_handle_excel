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
        self._processes: list[tuple[Callable[[], None], Callable[[], None]]] = []
        self._thread = threading.Thread(target=self._run, name="excel-watchdog", daemon=True)

    def start(self) -> None:
        self._thread.start()

    def progress(self) -> None:
        with self._lock:
            self._progress = time.monotonic()
            self._expired = False

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
                if not self._expired and time.monotonic() - self._progress >= self.timeout:
                    self._expired = True
                    expired = bool(self._processes)
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
