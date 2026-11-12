"""结束无进展的专用后台 Excel，让阻塞的 COM 调用返回异常。

持有进程句柄，而非稍后按 PID 查找，避免进程退出后 PID 被其他程序复用。
只允许注册调用方通过 DispatchEx 新建的专用 Excel 实例。

另外盯着后台 Excel 有没有弹出提示框（要密码、文件损坏要不要修复、文件被占用等）：
隐藏的 Excel 弹框没人能点，COM 调用会一直等下去。提示框连续 DIALOG_SECONDS 秒还在、
任务也没有进展，就记下框里的字、结束这个 Excel，任务马上能拿到异常（不用等 15 分钟），
再由 core/file_problem 弹框问用户重试还是跳过。

用户嫌某一份太慢时可以在进度框点「跳过这一份」（用户 2026-11-11 要）：request_skip() 结束正在做这份的后台 Excel，
卡着的 COM 调用马上报错返回，任务把这份记成「手动跳过」接着做下一份（不弹重试/跳过框）。
任务每份开始时调 file(label) 告诉看门狗现在做哪份；label 为 None 表示这一步不能跳过（比如正在写文件）。
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


_ACTIVE_LOCK = threading.Lock()
_ACTIVE: list["ExcelWatchdog"] = []  # 正在跑的看门狗（一次只有一个后台任务，偶尔嵌套）
SKIP_REASON = "你点了「跳过这一份」：已结束后台 Excel，这份没做完，之后可以单独再做"
STOP_REASON = "你点了「结束整个任务」：已结束后台 Excel，这份没做完"
NOT_DONE_REASON = "你点了「结束整个任务」：这份没做"
_stop_seq = 0  # 每点一次「结束整个任务」加一：精简后自动检查这种一连串的任务看它决定后面几步还做不做


class TaskStopped(RuntimeError):
    """用户点了「结束整个任务」：后面的文件不做了（file() 抛出，调用方记成「没做」接着很快走完）"""


def stop_seq() -> int:
    return _stop_seq


def request_stop() -> dict[str, Any]:
    """进度框点了「结束整个任务」：正在做的这份马上结束（同跳过），后面的都不做"""
    global _stop_seq
    with _ACTIVE_LOCK:
        _stop_seq += 1
        dogs = list(_ACTIVE)
    label = ""
    for dog in dogs:
        with dog._lock:
            dog._stop_task = True
        result = dog.skip("", STOP_REASON)
        if result and result.get("ok"):
            label = result["label"]
    from core import debug_log
    debug_log.event("⏹ 用户点了结束整个任务", 正在做=label or "（没有正在用 Excel 的文件）", 任务数=len(dogs))
    return {"ok": True, "label": label, "running": bool(dogs)}


def current() -> dict[str, Any]:
    """进度框问：现在能不能跳过、正在做哪份"""
    with _ACTIVE_LOCK:
        dogs = list(_ACTIVE)
    for dog in reversed(dogs):
        with dog._lock:
            if dog.label:
                return {"skippable": True, "label": dog.label, "seconds": round(time.monotonic() - dog._file_started)}
    return {"skippable": False, "label": "", "running": bool(dogs)}


def request_skip(label: str = "") -> dict[str, Any]:
    """进度框点了「跳过这一份」。label 是用户确认时看到的文件，和现在做的不一样（那份刚好做完了）就不跳"""
    with _ACTIVE_LOCK:
        dogs = list(_ACTIVE)
    for dog in reversed(dogs):
        result = dog.skip(label)
        if result is not None:
            return result
    from core import debug_log
    debug_log.event("⏭ 点了跳过这一份，但现在没有能跳过的文件", 文件=label or "（没给）", 任务数=len(dogs))
    return {"ok": False, "reason": "现在这一步不能跳过（可能正在保存文件或做收尾），等它做完这一步再点"
            if dogs else "现在没有正在跑的任务"}


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
        self.label: str | None = None  # 现在做的是哪份（file() 设），None = 现在不能跳过
        self._file_started = time.monotonic()
        self._skip = ""  # 这份是用户点「跳过这一份」结束的：原因（file() / take_skip() 清掉）
        # 每份最多做几秒，超过自动跳过（进度框里用户自己设，0 = 不限）；None = 这个任务不用（扫描有自己的上限）
        self.file_limit: Callable[[], float] | None = None
        self._stop_task = False  # 用户点了「结束整个任务」
        self._thread = threading.Thread(target=self._run, name="excel-watchdog", daemon=True)

    def start(self, *, listed: bool = True) -> None:
        """listed=False：不占任务槽的旁路活（Excel 对比导 PDF）用，进度框的「跳过这一份 / 结束整个任务」管不到它"""
        if listed:
            with _ACTIVE_LOCK:
                _ACTIVE.append(self)
        self._thread.start()

    def file(self, label: str | None) -> bool:
        """开始做一份文件（label = 显示给用户的名字），或者进入不能跳过的一步（None）。
        返回真：上一份点了跳过、Excel 已经被结束了，那份却没报错就做完了（比如正在拼 PDF，没用到 Excel），
        调用方要把手里的 Excel 扔掉、下次重开"""
        with self._lock:
            if self._stop_task and label is not None:
                raise TaskStopped(NOT_DONE_REASON)
            self.label = label
            if label is None:  # 进入不能跳过的一步：点过的跳过留着给调用方的异常处理看
                return False
            stale = bool(self._skip)
            self._skip = ""
            self._file_started = time.monotonic()  # 重试同一份也从头算时间
        self.progress()
        if stale:
            from core import debug_log
            debug_log.event("♻️ 上一份点了跳过但已经做完，后台 Excel 已结束，下一份换新的", 下一份=label or "")
        return stale

    def skip(self, label: str = "", reason: str = "") -> dict[str, Any] | None:
        """结束现在这份的后台 Excel。返回 None 表示这个看门狗现在没在做能跳过的文件"""
        from core import debug_log
        with self._lock:
            if not self.label:
                return None
            if label and label != self.label:
                return {"ok": False, "reason": f"「{label}」已经做完了，现在在做「{self.label}」；要跳过这份请再点一次"}
            current_label, seconds = self.label, round(time.monotonic() - self._file_started)
            self._skip = reason or SKIP_REASON
            self._expired = True
            processes = list(self._processes)
        who = "超过每份时间上限，自动跳过" if reason else "用户点了跳过这一份"
        debug_log.dump_threads(f"{who}：{current_label}（已做 {seconds} 秒），即将结束后台 Excel")
        for terminate, _ in processes:
            try:
                terminate()
            except Exception:
                pass
        debug_log.event(f"⏭ {who}，已结束后台 Excel", 文件=current_label, 已做秒=seconds, Excel进程=len(processes))
        return {"ok": True, "label": current_label, "seconds": seconds}

    def _check_file_limit(self) -> None:
        if self.file_limit is None:
            return
        try:
            limit = float(self.file_limit() or 0)
        except Exception:
            return
        with self._lock:
            if not limit or not self.label or self._skip or not self._processes:
                return
            label, spent = self.label, time.monotonic() - self._file_started
        if spent < limit:
            return
        self.skip(label, f"这份做了 {spent / 60:.0f} 分钟还没做完，超过你设的「每份最多 {limit / 60:.0f} 分钟」，"
                         f"已自动跳过（结束了后台 Excel），之后可以单独再做")

    @property
    def stop_requested(self) -> bool:
        with self._lock:
            return self._stop_task

    @property
    def skip_requested(self) -> bool:
        with self._lock:
            return bool(self._skip)

    def take_skip(self) -> str:
        """这份是不是用户点跳过结束的：是的话返回原因并清掉（只算一次）"""
        with self._lock:
            reason, self._skip = self._skip, ""
            return reason

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
            dialog, expired, deadline, skip = self._dialog, self._expired, self._deadline_hit, self._skip
            self._dialog = ""
        if skip:
            return skip
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
            self._check_file_limit()
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
        with _ACTIVE_LOCK:
            if self in _ACTIVE:
                _ACTIVE.remove(self)
        self._stop.set()
        self._thread.join(timeout=2)
        with self._lock:
            for _, close in self._processes:
                close()
            self._processes.clear()
            self._pids.clear()
            self._main_windows.clear()
