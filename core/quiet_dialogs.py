"""把后台隐藏 Excel 导出 PDF 时弹出的「発行中...」进度框挪到屏幕外。

Excel.Visible = False 也挡不住这个进度框，后台连续导出时它会一闪一闪。
这里用 WinEvent 钩子盯住指定的 Excel 进程：一有窗口创建或显示，标题是导出进度框的，
就立刻挪到屏幕外（不关闭、不点取消，导出照常进行）。只认进度框的标题，
别的提示框照常显示，免得藏起来没人点导致 Excel 卡住。非 Windows 环境什么也不做。
"""
import sys
import threading
from typing import Any

# 各语言版 Excel 导出进度框的标题
PROGRESS_TITLES = ("発行中", "Publishing", "正在发布", "正在發佈", "正在發布")

_EVENT_OBJECT_CREATE = 0x8000
_EVENT_OBJECT_SHOW = 0x8002
_OBJID_WINDOW = 0
_WINEVENT_OUTOFCONTEXT = 0x0000
_WINEVENT_SKIPOWNPROCESS = 0x0002
_SWP_NOSIZE = 0x0001
_SWP_NOZORDER = 0x0004
_SWP_NOACTIVATE = 0x0010
_GA_ROOT = 2
_WM_QUIT = 0x0012
OFFSCREEN = -32000


def is_progress_title(title: str) -> bool:
    return any(key in title for key in PROGRESS_TITLES)


class ProgressDialogHider:
    """盯住一个 Excel 进程（pid），用 stop() 结束"""

    def __init__(self, pid: int) -> None:
        self.pid = pid
        self._thread_id: int | None = None
        self._ready = threading.Event()
        self._thread = threading.Thread(target=self._loop, name="quiet-dialogs", daemon=True)
        self._thread.start()
        self._ready.wait(5)

    def _loop(self) -> None:
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.WinDLL("user32", use_last_error=True)
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
        user32.GetAncestor.restype = wintypes.HWND
        user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
        user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
        user32.SetWindowPos.argtypes = [
            wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.UINT,
        ]
        user32.UnhookWinEvent.argtypes = [wintypes.HANDLE]

        def callback(_hook, _event, hwnd, id_object, _id_child, _thread, _time):
            if not hwnd or id_object != _OBJID_WINDOW or user32.GetAncestor(hwnd, _GA_ROOT) != hwnd:
                return
            buffer = ctypes.create_unicode_buffer(256)
            user32.GetWindowTextW(hwnd, buffer, 256)
            if is_progress_title(buffer.value):
                user32.SetWindowPos(
                    hwnd, None, OFFSCREEN, OFFSCREEN, 0, 0, _SWP_NOSIZE | _SWP_NOZORDER | _SWP_NOACTIVATE,
                )

        self._callback = proc_type(callback)  # 保持引用，避免被回收
        self._thread_id = kernel32.GetCurrentThreadId()
        hook = user32.SetWinEventHook(
            _EVENT_OBJECT_CREATE, _EVENT_OBJECT_SHOW, None, self._callback, self.pid, 0,
            _WINEVENT_OUTOFCONTEXT | _WINEVENT_SKIPOWNPROCESS,
        )
        self._ready.set()
        if not hook:
            return
        try:
            msg = wintypes.MSG()
            while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        finally:
            user32.UnhookWinEvent(hook)

    def stop(self) -> None:
        if self._thread_id is not None:
            import ctypes
            ctypes.WinDLL("user32").PostThreadMessageW(self._thread_id, _WM_QUIT, 0, 0)
        self._thread.join(2)


def hide_progress_dialogs(excel: Any) -> ProgressDialogHider | None:
    """给 DispatchEx 出来的隐藏 Excel 装上钩子；拿不到进程或不是 Windows 时返回 None"""
    if sys.platform != "win32":
        return None
    try:
        import ctypes
        from ctypes import wintypes
        pid = wintypes.DWORD()
        ctypes.WinDLL("user32").GetWindowThreadProcessId(wintypes.HWND(int(excel.Hwnd)), ctypes.byref(pid))
        return ProgressDialogHider(pid.value) if pid.value else None
    except Exception:
        return None
