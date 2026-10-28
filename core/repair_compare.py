"""修复结果一键打开新旧文件；独立 Excel 实例避开同名冲突，备份只读。

可见 Excel 的 COM 对象留在专用线程中，后续请求复用，避免跨线程使用 COM 或重复开窗口。
"""
from __future__ import annotations

import queue
import threading
from concurrent.futures import Future, TimeoutError
from pathlib import Path
from typing import Any

from core import excel_ops as ops


def comparison_target(kind: str, folder: str, name: str, sheet: str = "", cell: str = "") -> dict:
    """只接受服务端结果中已经配对的两个文件，旧表定位也由服务端清单决定。"""
    from core import link_repair, mark_finder

    if kind not in {"links", "colors"}:
        raise ValueError("未知的修复结果类型")
    status = (link_repair if kind == "links" else mark_finder).status()
    if status.get("running"):
        raise ValueError("正在扫描或修复，请等任务结束后再打开对比")
    data = status.get("applied") or status.get("plan")
    if not data or Path(data["folder"]).resolve() != Path(folder).resolve():
        raise ValueError("修复结果已更新，请刷新页面后再打开")
    file = next((file for file in data["files"] if file["name"] == name), None)
    if not file or not file.get("old"):
        raise ValueError("这份文件没有唯一对应的备份，暂时不能同时打开")
    cells = file.get("cells") or file.get("hits") or []
    if sheet:
        item = next((item for item in cells if item["sheet"] == sheet
                     and cell.replace("$", "") in {item["addr"].replace("$", ""), item.get("anchor", item["addr"].split(":")[0]).replace("$", "")}), None)
        if item is None:
            raise ValueError("这个位置不在当前修复结果中，请刷新页面")
    else:
        item = cells[0] if cells else {}
    current, backup = Path(file["path"]).resolve(strict=True), Path(file["old"]).resolve(strict=True)
    if current == backup:
        raise ValueError("当前文件和备份是同一个文件，不能用于对比")
    address = item.get("anchor") or item.get("addr", "").split(":")[0]
    return {"current": current, "backup": backup, "sheet": item.get("sheet", ""),
            "old_sheet": item.get("old_sheet", item.get("sheet", "")), "cell": address}


def _tile_windows(current: tuple, backup: tuple) -> str:
    """把两份工作簿放到当前窗口所在显示器的左右两半；失败时仍保留已打开文件。"""
    if ops.sys.platform != "win32":
        return "当前系统无法自动左右排列，请手动排列窗口"
    try:
        import win32api
        import win32con
        import win32gui

        handles = [int(book.Windows(1).Hwnd) for _, book in (current, backup)]
        monitor = win32api.MonitorFromWindow(handles[0], win32con.MONITOR_DEFAULTTONEAREST)
        left, top, right, bottom = win32api.GetMonitorInfo(monitor)["Work"]
        middle = left + (right - left) // 2
        for hwnd, x, width in ((handles[0], left, middle - left), (handles[1], middle, right - middle)):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, top, width, bottom - top, win32con.SWP_NOACTIVATE)
        return ""
    except Exception:
        return "已打开两个文件，自动左右排列失败，请手动排列窗口"


class ComparisonWindows:
    """只在创建它的 COM 线程中使用，不保存、关闭用户文件或退出可见 Excel。"""
    def __init__(self):
        self.sessions: dict[tuple[Path, bool], Any] = {}

    def _open(self, path: Path, readonly: bool) -> tuple:
        key = (path, readonly)
        owned = False
        excel = self.sessions.get(key)
        if excel is not None:
            try:
                list(excel.Workbooks)  # 用户可能已经关闭了整个 Excel。
            except Exception:
                self.sessions.pop(key, None)
                excel = None
        if excel is None:
            excel = ops.win32.DispatchEx("Excel.Application") if readonly else ops._visible_excel()
            owned = readonly
        books = list(excel.Workbooks)
        workbook = next((book for book in books if Path(book.FullName).resolve() == path), None)
        if workbook is None:
            # 当前文件也可能与用户已打开的另一目录文件同名，不能塞进同一个实例。
            if any(Path(book.FullName).name.casefold() == path.name.casefold() for book in books):
                excel = ops.win32.DispatchEx("Excel.Application")
                owned = True
            if readonly:
                alerts = excel.DisplayAlerts
                excel.AutomationSecurity = ops.MSO_AUTOMATION_SECURITY_FORCE_DISABLE
                excel.EnableEvents = False
                excel.DisplayAlerts = False
            options = {"UpdateLinks": 0, "IgnoreReadOnlyRecommended": True}
            if readonly:
                options["ReadOnly"] = True
                options["AddToMru"] = False
            try:
                workbook = ops.open_workbook_checked(excel, str(path), **options)
            except Exception:
                if owned:
                    try:
                        if not excel.Workbooks.Count:
                            excel.Quit()  # 只退出新建而且尚未打开任何文件的实例。
                    except Exception:
                        pass
                raise
            finally:
                if readonly:
                    try:
                        excel.DisplayAlerts = alerts
                    except Exception:
                        pass
        excel.Visible = True
        excel.UserControl = True
        self.sessions[key] = excel
        return excel, workbook

    def open(self, target: dict) -> dict:
        opened, notes, windows = [], [], {}
        for role, readonly in (("current", False), ("backup", True)):
            label = "当前文件" if role == "current" else "原始备份"
            try:
                excel, book = self._open(target[role], readonly)
                windows[role] = (excel, book)
                opened.append(role)
                # 文件名相同也能一眼区分两边；标题不写回磁盘。
                if book.Windows.Count:
                    try:
                        book.Windows(1).Caption = f"{label}{'（只读）' if readonly else ''} · {target[role].name}"
                    except Exception:
                        pass
                note = ops.focus_workbook_in_excel(excel, book, target["old_sheet"] if readonly else target["sheet"], target["cell"])
                if note:
                    notes.append(f"{label}：{note}")
                if not readonly and book.ReadOnly:
                    notes.append("当前文件以只读方式打开，请关闭其他占用它的窗口后再编辑")
            except Exception as error:
                notes.append(f"{label}打开或定位失败：{error}")
        if not opened:
            raise ValueError("；".join(notes))
        if len(windows) == 2:
            note = _tile_windows(windows["current"], windows["backup"])
            if note:
                notes.append(note)
            # 半屏后重新按实际可用宽度缩放；隐藏表仍只提示，不取消隐藏。
            for role, (excel, book) in windows.items():
                try:
                    sheet = book.Sheets(target["old_sheet"] if role == "backup" else target["sheet"])
                    if target["cell"] and sheet.Visible == ops.XL_SHEET_VISIBLE:
                        ops._goto_fit_width(excel, book, sheet, target["cell"])
                except Exception:
                    pass
        for role in ("backup", "current"):
            if role in windows:
                ops._bring_to_front(windows[role][0])
        return {"opened": opened, "note": "；".join(notes)}


_LOCK = threading.Lock()
_QUEUE: queue.Queue = queue.Queue()
_THREAD: threading.Thread | None = None
_PENDING: dict[tuple, Future] = {}


def _serve() -> None:
    windows = None
    initialized = False
    try:
        while True:
            request = _QUEUE.get()
            if request is None:
                return
            key, target, future = request
            try:
                if not initialized:
                    ops.pythoncom.CoInitialize()
                    initialized = True
                    windows = ComparisonWindows()
                future.set_result(windows.open(target))
            except Exception as error:
                future.set_exception(error)
            finally:
                with _LOCK:
                    if _PENDING.get(key) is future:
                        _PENDING.pop(key, None)
    finally:
        if initialized:
            ops.pythoncom.CoUninitialize()


def open_comparison(target: dict, timeout: float = 45) -> dict:
    """HTTP 线程等结果，窗口对象在同一个专用线程中复用；超时不重复排同一请求。"""
    global _THREAD
    key = tuple(target[name] for name in ("current", "backup", "sheet", "old_sheet", "cell"))
    with _LOCK:
        if _THREAD is None or not _THREAD.is_alive():
            _THREAD = threading.Thread(target=_serve, name="Excel 新旧对比窗口", daemon=True)
            _THREAD.start()
        future = _PENDING.get(key)
        if future is None or future.done():
            future = _PENDING[key] = Future()
            _QUEUE.put((key, target, future))
    try:
        return future.result(timeout)
    except TimeoutError:
        raise ValueError("Excel 打开较慢，仍在后台继续；请先关闭 Excel 对话框，稍后再点此链接查看") from None
