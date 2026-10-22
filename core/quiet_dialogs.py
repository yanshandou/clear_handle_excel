"""让后台隐藏 Excel 导出 PDF 时弹出的「発行中...」进度框尽量不打扰用户。

Excel.Visible = False 也挡不住这个进度框，后台连续导出时它会一闪一闪。这里做两件事：

1. 把隐藏的 Excel 主窗口（进度框的父窗口）挪到屏幕右下角。进度框按父窗口居中弹出，
   所以就算来不及藏，也只是在右下角闪一下，不会挡在屏幕中间。主窗口本身仍然隐藏，
   退出前 release() 会挪回原位，免得 Excel 把这个位置记下来、影响用户自己打开的 Excel。
2. 用 WinEvent 钩子盯住这个 Excel 进程：进度框一出现（创建、显示、改标题、被挪动），
   就立刻挪到屏幕外（不关闭、不点取消，导出照常进行）。钩子是异步的，
   所以第一帧可能还是会露出来，这就是第 1 步兜底的原因。

只认进度框的标题，别的提示框照常显示，免得藏起来没人点导致 Excel 卡住。
Excel 进程退出后钩子线程自己结束。非 Windows 环境什么也不做。
"""
import sys
import threading
from typing import Any

# 各语言版 Excel 导出进度框的标题
PROGRESS_TITLES = ("発行中", "Publishing", "正在发布", "正在發佈", "正在發布")

_EVENT_OBJECT_CREATE = 0x8000
_EVENT_OBJECT_NAMECHANGE = 0x800C  # 中间包括 SHOW(0x8002) 和 LOCATIONCHANGE(0x800B)
_OBJID_WINDOW = 0
_WINEVENT_OUTOFCONTEXT = 0x0000
_WINEVENT_SKIPOWNPROCESS = 0x0002
_SWP_NOSIZE = 0x0001
_SWP_NOZORDER = 0x0004
_SWP_NOACTIVATE = 0x0010
_GA_ROOT = 2
_WM_QUIT = 0x0012
_PM_REMOVE = 0x0001
_QS_ALLINPUT = 0x04FF
_INFINITE = 0xFFFFFFFF
_WAIT_OBJECT_0 = 0
_SYNCHRONIZE = 0x00100000
_SPI_GETWORKAREA = 0x0030
OFFSCREEN = -32000
# 挪到右下角的隐藏主窗口大小（进度框按它居中，超出屏幕的部分系统会推回屏幕内）
PARENT_WIDTH = 320
PARENT_HEIGHT = 120

_HIDERS: dict[int, "ProgressDialogHider"] = {}
_HIDERS_LOCK = threading.Lock()


def is_progress_title(title: str) -> bool:
    return any(key in title for key in PROGRESS_TITLES)


def corner_rect(work_area: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    """工作区（左, 上, 右, 下）右下角的 (x, y, 宽, 高)"""
    left, top, right, bottom = work_area
    width = min(PARENT_WIDTH, right - left)
    height = min(PARENT_HEIGHT, bottom - top)
    return right - width, bottom - height, width, height


def _user32() -> Any:
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.GetAncestor.restype = wintypes.HWND
    user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    user32.SetWindowPos.argtypes = [
        wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.UINT,
    ]
    return user32


class ProgressDialogHider:
    """盯住一个 Excel 进程（pid），Excel 退出后自动结束，也可以用 stop() 提前结束"""

    def __init__(self, pid: int, excel_hwnd: int = 0) -> None:
        self.pid = pid
        self.excel_hwnd = excel_hwnd
        self._original_rect: tuple[int, int, int, int] | None = None
        self._thread_id: int | None = None
        self._ready = threading.Event()
        if excel_hwnd:
            self._move_parent_to_corner()
        self._thread = threading.Thread(target=self._loop, name="quiet-dialogs", daemon=True)
        self._thread.start()
        self._ready.wait(5)

    def _move_parent_to_corner(self) -> None:
        import ctypes
        from ctypes import wintypes

        user32 = _user32()
        rect = wintypes.RECT()
        if not user32.GetWindowRect(self.excel_hwnd, ctypes.byref(rect)):
            return
        area = wintypes.RECT()
        if not user32.SystemParametersInfoW(_SPI_GETWORKAREA, 0, ctypes.byref(area), 0):
            return
        self._original_rect = (rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)
        x, y, width, height = corner_rect((area.left, area.top, area.right, area.bottom))
        # 不带 SWP_SHOWWINDOW，主窗口仍然隐藏
        user32.SetWindowPos(self.excel_hwnd, None, x, y, width, height, _SWP_NOZORDER | _SWP_NOACTIVATE)

    def _restore_parent(self) -> None:
        if self._original_rect is None:
            return
        try:
            x, y, width, height = self._original_rect
            _user32().SetWindowPos(self.excel_hwnd, None, x, y, width, height, _SWP_NOZORDER | _SWP_NOACTIVATE)
        except Exception:
            pass
        self._original_rect = None

    def _loop(self) -> None:
        import ctypes
        from ctypes import wintypes

        user32 = _user32()
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        proc_type = ctypes.WINFUNCTYPE(
            None, wintypes.HANDLE, wintypes.DWORD, wintypes.HWND,
            wintypes.LONG, wintypes.LONG, wintypes.DWORD, wintypes.DWORD,
        )
        user32.SetWinEventHook.restype = wintypes.HANDLE
        user32.SetWinEventHook.argtypes = [
            wintypes.DWORD, wintypes.DWORD, wintypes.HMODULE, proc_type,
            wintypes.DWORD, wintypes.DWORD, wintypes.DWORD,
        ]
        user32.UnhookWinEvent.argtypes = [wintypes.HANDLE]
        user32.MsgWaitForMultipleObjects.argtypes = [
            wintypes.DWORD, ctypes.POINTER(wintypes.HANDLE), wintypes.BOOL, wintypes.DWORD, wintypes.DWORD,
        ]
        user32.MsgWaitForMultipleObjects.restype = wintypes.DWORD
        kernel32.OpenProcess.restype = wintypes.HANDLE
        kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

        def callback(_hook, _event, hwnd, id_object, _id_child, _thread, _time):
            if not hwnd or id_object != _OBJID_WINDOW or user32.GetAncestor(hwnd, _GA_ROOT) != hwnd:
                return
            buffer = ctypes.create_unicode_buffer(256)
            user32.GetWindowTextW(hwnd, buffer, 256)
            if not is_progress_title(buffer.value):
                return
            rect = wintypes.RECT()
            # 已经在屏幕外就不再挪，避免自己挪动触发的位置事件来回循环
            if user32.GetWindowRect(hwnd, ctypes.byref(rect)) and rect.left == OFFSCREEN:
                return
            user32.SetWindowPos(
                hwnd, None, OFFSCREEN, OFFSCREEN, 0, 0, _SWP_NOSIZE | _SWP_NOZORDER | _SWP_NOACTIVATE,
            )

        self._callback = proc_type(callback)  # 保持引用，避免被回收
        self._thread_id = kernel32.GetCurrentThreadId()
        hook = user32.SetWinEventHook(
            _EVENT_OBJECT_CREATE, _EVENT_OBJECT_NAMECHANGE, None, self._callback, self.pid, 0,
            _WINEVENT_OUTOFCONTEXT | _WINEVENT_SKIPOWNPROCESS,
        )
        process = kernel32.OpenProcess(_SYNCHRONIZE, False, self.pid)
        self._ready.set()
        if not hook:
            if process:
                kernel32.CloseHandle(process)
            return
        try:
            handles = (wintypes.HANDLE * 1)(process)
            msg = wintypes.MSG()
            while True:
                if process:
                    result = user32.MsgWaitForMultipleObjects(1, handles, False, _INFINITE, _QS_ALLINPUT)
                    if result == _WAIT_OBJECT_0:
                        break  # Excel 已退出
                elif not user32.WaitMessage():
                    break
                stop = False
                while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, _PM_REMOVE):
                    if msg.message == _WM_QUIT:
                        stop = True
                        break
                    user32.TranslateMessage(ctypes.byref(msg))
                    user32.DispatchMessageW(ctypes.byref(msg))
                if stop:
                    break
        finally:
            user32.UnhookWinEvent(hook)
            if process:
                kernel32.CloseHandle(process)
            with _HIDERS_LOCK:
                if _HIDERS.get(self.pid) is self:
                    del _HIDERS[self.pid]

    def stop(self) -> None:
        self._restore_parent()
        if self._thread_id is not None:
            import ctypes
            ctypes.WinDLL("user32").PostThreadMessageW(self._thread_id, _WM_QUIT, 0, 0)
        self._thread.join(2)


def _excel_window(excel: Any) -> tuple[int, int]:
    import ctypes
    from ctypes import wintypes

    hwnd = int(excel.Hwnd)
    pid = wintypes.DWORD()
    ctypes.WinDLL("user32").GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(pid))
    return hwnd, pid.value


def hide_progress_dialogs(excel: Any) -> ProgressDialogHider | None:
    """给 DispatchEx 出来的隐藏 Excel 装上钩子；拿不到进程或不是 Windows 时返回 None"""
    if sys.platform != "win32":
        return None
    try:
        hwnd, pid = _excel_window(excel)
        if not pid:
            return None
        hider = ProgressDialogHider(pid, hwnd)
        with _HIDERS_LOCK:
            _HIDERS[pid] = hider
        return hider
    except Exception:
        return None


def release_excel(excel: Any) -> None:
    """Excel.Quit() 之前调用：把主窗口挪回原位并卸下钩子；没装过钩子时什么也不做"""
    if sys.platform != "win32" or excel is None:
        return
    try:
        _, pid = _excel_window(excel)
    except Exception:
        return
    with _HIDERS_LOCK:
        hider = _HIDERS.pop(pid, None)
    if hider is not None:
        hider.stop()
