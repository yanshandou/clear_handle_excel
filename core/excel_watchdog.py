"""结束无进展的专用后台 Excel，让阻塞的 COM 调用返回异常。

持有进程句柄，而非稍后按 PID 查找，避免进程退出后 PID 被其他程序复用。
只允许注册调用方通过 DispatchEx 新建的专用 Excel 实例。

另外盯着后台 Excel 有没有弹出提示框（要密码、文件损坏要不要修复、文件被占用等）：
隐藏的 Excel 弹框没人能点，COM 调用会一直等下去。提示框连续 DIALOG_SECONDS 秒还在、
任务也没有进展，就记下框里的字、结束这个 Excel，任务马上能拿到异常（不用等 15 分钟），
再由 core/file_problem 弹框问用户重试还是跳过。
"""
import os
import threading
import time
from typing import Any, Callable


DIALOG_SECONDS = 20  # 提示框连续出现这么久、任务又没进展，就当作卡在提示框上
# 只认这些窗口类：消息框（#32770）、Excel 自己的对话框（bosa_sdm_XL9 等，要密码就是它）、新式对话框（NUIDialog）
DIALOG_CLASSES = ("#32770", "bosa_sdm_", "NUIDialog")
# 正常进度框，不算卡住（导出 PDF 的「発行中」、打印机驱动的「印刷中」）
BUSY_TITLES = ("発行中", "Publishing", "正在发布", "正在發佈", "正在發布", "印刷", "Printing", "打印", "列印")


def _owned_process(excel: Any) -> tuple[Callable[[], None], Callable[[], None], int, int]:
    import win32api
    import win32process

    hwnd = int(excel.Hwnd)
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    handle = win32api.OpenProcess(0x0001 | 0x00100000, False, pid)  # TERMINATE | SYNCHRONIZE
    return lambda: win32api.TerminateProcess(handle, 1), handle.Close, pid, hwnd


def is_dialog(class_name: str, title: str) -> bool:
    if not any(class_name.startswith(prefix) for prefix in DIALOG_CLASSES):
        return False
    return not any(word in title for word in BUSY_TITLES)


def find_dialogs(pids: set[int], skip: set[int]) -> list[tuple[int, str]]:
    """这些 Excel 进程里正显示着的提示框：[(窗口句柄, 「标题：框里的字」)]。非 Windows 返回空"""
    import sys
    if sys.platform != "win32" or not pids:
        return []
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.EnumWindows.argtypes = [proc, wintypes.LPARAM]
    user32.EnumChildWindows.argtypes = [wintypes.HWND, proc, wintypes.LPARAM]
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.IsWindowVisible.argtypes = [wintypes.HWND]
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

    def text(hwnd: int) -> str:
        buf = ctypes.create_unicode_buffer(512)
        user32.GetWindowTextW(hwnd, buf, 512)
        return buf.value.strip()

    def cls(hwnd: int) -> str:
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, buf, 256)
        return buf.value

    found: list[tuple[int, str]] = []

    def each(hwnd: int, _: int) -> bool:
        try:
            if not hwnd or hwnd in skip or not user32.IsWindowVisible(hwnd):
                return True
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value not in pids:
                return True
            title = text(hwnd)
            if not is_dialog(cls(hwnd), title):
                return True
            parts: list[str] = []

            def child(c: int, _: int) -> bool:
                words = text(c)
                if words and words not in parts and len(parts) < 12:
                    parts.append(words)
                return True

            user32.EnumChildWindows(hwnd, proc(child), 0)
            found.append((int(hwnd), "：".join(filter(None, [title, " / ".join(parts)]))[:300] or "（框里没有可读的文字）"))
        except Exception:
            pass
        return True

    user32.EnumWindows(proc(each), 0)
    return found


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
        self._pids: set[int] = set()
        self._main_windows: set[int] = set()
        self._dialog_since: float | None = None
        self._dialog_text = ""
        self._dialog = ""  # 因为提示框结束了 Excel 时，框里的字
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

    @property
    def dialog(self) -> str:
        with self._lock:
            return self._dialog

    def explain(self, err: BaseException) -> str:
        """给用户看的出错原因：Excel 弹了框 / 太久没反应时说清楚，后面附上原来的报错"""
        raw = str(err) or type(err).__name__
        with self._lock:
            dialog, expired, deadline = self._dialog, self._expired, self._deadline_hit
            self._dialog = ""
        if dialog:
            return (f"后台 Excel 弹出了提示框，没人能点，已结束后台 Excel。框里写着：「{dialog}」"
                    f"（常见原因：文件要密码、文件损坏要修复、文件被别人占用）")
        if expired and not deadline:
            return f"后台 Excel {self.timeout / 60:.0f} 分钟没有反应，已结束后台 Excel（{raw}）"
        return raw

    def register(self, excel: Any) -> None:
        terminate, close, *ids = _owned_process(excel)
        with self._lock:
            self._processes.append((terminate, close))
            if len(ids) == 2:
                self._pids.add(ids[0])
                self._main_windows.add(ids[1])
            self._progress = time.monotonic()

    def _check_dialog(self) -> str:
        """提示框连续出现 DIALOG_SECONDS 秒、任务也没进展时返回框里的字，否则返回空"""
        with self._lock:
            if self._expired or not self._pids or time.monotonic() - self._progress < DIALOG_SECONDS:
                self._dialog_since = None
                return ""
            pids, skip = set(self._pids), set(self._main_windows)
        try:
            dialogs = find_dialogs(pids, skip)
        except Exception:
            dialogs = []
        now = time.monotonic()
        with self._lock:
            if not dialogs:
                self._dialog_since = None
                return ""
            if self._dialog_since is None:
                self._dialog_since = now
                self._dialog_text = dialogs[0][1]
                from core import debug_log
                debug_log.event("🪧 后台 Excel 弹出了提示框，再等一会儿看它会不会自己消失", 内容=self._dialog_text)
                return ""
            if now - self._dialog_since < DIALOG_SECONDS:
                return ""
            return self._dialog_text

    def _run(self) -> None:
        while not self._stop.wait(min(1.0, self.timeout / 4)):
            expired = False
            dialog = self._check_dialog()
            with self._lock:
                now = time.monotonic()
                overdue = self._deadline is not None and now >= self._deadline
                if not self._expired and (dialog or overdue or now - self._progress >= self.timeout):
                    self._expired = True
                    if dialog:
                        self._dialog, self._dialog_since = dialog, None
                    elif overdue:
                        self._deadline, self._deadline_hit = None, True
                    expired = bool(self._processes)
                    if expired:
                        # 结束 Excel 之前记下任务线程停在哪一行：卡在哪个调用上一看就知道
                        from core import debug_log
                        debug_log.dump_threads(f"看门狗：后台 Excel 弹出提示框「{dialog}」，没人能点，即将结束它" if dialog
                                               else "看门狗：这份文件扫描超过时间上限，即将结束后台 Excel" if overdue
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
            self._pids.clear()
            self._main_windows.clear()
