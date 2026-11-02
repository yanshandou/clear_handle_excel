"""批量任务遇到有问题的文件：停下来问用户「重试」还是「跳过」，而不是卡死或者悄悄跳过。

后台任务线程调用 ask()，它会一直等到网页上有人选（每个打开的页面都会弹框，见 static/file_problem.js）：
- 重试：用户在 Excel 里把文件修好、保存并关闭后点，调用方重启后台 Excel 再做这份文件
- 跳过：照以前那样记下错误，接着做下一份

弹框上可以勾「这次任务后面再出错都直接跳过」，夜里挂机时不用守着。
也可以设「没人选时多少分钟后自动跳过」（主页和弹框上都能改，0 = 一直等），存在数据库里，重启也在。
什么错误要问见 should_ask()：Excel 打不开、卡住、弹框、崩溃、文件被占用这类「文件本身的问题」才问；
程序自己的检查结论（找不到备份、扫描后文件变了等）不问，照旧记下错误。

一次只有一个后台任务（busy 标志），这里全局只放一个等待中的问题。
"""
import re
import threading
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Callable, Iterator

from core import debug_log

_COND = threading.Condition()
_pending: dict[str, Any] | None = None
_answer: dict[str, Any] | None = None
_run: dict[str, Any] = {"task": "", "auto_skip": False, "asked": 0}
_seq = 0
SETTINGS_NAME = "file_problem_settings"
MAX_TIMEOUT_MINUTES = 24 * 60
_settings_cache: dict[str, Any] | None = None

# 跑回归测试时没人点弹框：默认直接跳过（和以前一样）。测试要真等用户选择时设成 ""
AUTO_ANSWER: str | None = None

# 这些字眼说明是文件本身出了问题（打不开 / 被占用），用户修好或关掉文件后重试可能就好了
_FILE_HINTS = ("Excel 中打开", "被占用", "打不开", "未能打开", "没能打开", "无法打开", "密码", "パスワード",
               "提示框", "没有反应", "无进展", "Excel 弹出")


def should_ask(err: BaseException, message: str = "") -> bool:
    """程序自己判断出来的结论（ValueError，比如找不到备份）不打扰用户；Excel / 系统报的错才问"""
    text = message or str(err)
    if any(hint in text for hint in _FILE_HINTS):
        return True
    return not isinstance(err, ValueError)


def get_settings() -> dict[str, Any]:
    """timeout_minutes：等用户选择最多等多少分钟，到时间自动跳过；0 = 一直等"""
    global _settings_cache
    if _settings_cache is None:
        saved: dict[str, Any] = {}
        try:
            from core import storage
            saved = storage.load_snapshot(SETTINGS_NAME) or {}
        except Exception as err:
            debug_log.event("⚠️ 没能读出出错弹框设置，按一直等处理", 原因=str(err))
        minutes = saved.get("timeout_minutes", 0)
        _settings_cache = {"timeout_minutes": minutes if isinstance(minutes, int) and not isinstance(minutes, bool)
                           and 0 <= minutes <= MAX_TIMEOUT_MINUTES else 0}
    return dict(_settings_cache)


def set_settings(timeout_minutes: Any) -> dict[str, Any]:
    global _settings_cache
    if (isinstance(timeout_minutes, bool) or not isinstance(timeout_minutes, int)
            or not 0 <= timeout_minutes <= MAX_TIMEOUT_MINUTES):
        raise ValueError(f"自动跳过的等待时间要填 0–{MAX_TIMEOUT_MINUTES} 之间的整数分钟（0 = 一直等）")
    from core import storage
    settings = {"timeout_minutes": timeout_minutes}
    storage.save_snapshot(SETTINGS_NAME, settings)
    with _COND:
        _settings_cache = dict(settings)
        _COND.notify_all()  # 正在等的那份文件马上按新时间算
    debug_log.event("⚙️ 出错弹框设置", 没人选时自动跳过分钟=timeout_minutes or "一直等")
    return dict(settings)


def begin(name: str) -> None:
    """一个批量任务开始：任务名显示在弹框上，「后面都跳过」只在这次任务里有效"""
    with _COND:
        _run.update(task=name, auto_skip=False, asked=0)
        _clear_locked()


def finish() -> None:
    with _COND:
        _run.update(task="", auto_skip=False)
        _clear_locked()


@contextmanager
def task(name: str) -> Iterator[None]:
    begin(name)
    try:
        yield
    finally:
        finish()


def _clear_locked() -> None:
    global _pending, _answer
    _pending, _answer = None, None
    _COND.notify_all()


def pending() -> dict[str, Any] | None:
    """网页轮询：现在有没有在等用户选择的出错文件"""
    with _COND:
        if _pending is None:
            return None
        problem = dict(_pending)
    waited = time.time() - problem["since"]
    minutes = get_settings()["timeout_minutes"]
    left = max(0, round(minutes * 60 - waited)) if minutes else None
    return {**problem, "waited": round(waited), "timeout_minutes": minutes, "auto_skip_in": left}


def answer(problem_id: str, choice: str, auto_skip: bool = False) -> None:
    global _answer
    if choice not in {"retry", "skip"}:
        raise ValueError("只能选择重试或跳过")
    with _COND:
        if _pending is None or _pending["id"] != problem_id:
            raise ValueError("这个问题已经处理过了（可能在别的页面选过），页面会自动刷新")
        _answer = {"id": problem_id, "choice": choice, "auto_skip": bool(auto_skip) and choice == "skip"}
        _COND.notify_all()


def ask(path: str | Path, label: str, error: str, *, folder: str | Path = "",
        tick: Callable[[], None] | None = None, should_stop: Callable[[], bool] | None = None) -> str:
    """停下来等用户选择，返回 "retry" 或 "skip"。

    调用前先关掉后台 Excel（别让它占着文件）。tick 每秒调一次（让看门狗知道不是卡住）；
    should_stop 返回真（用户点了暂停 / 停止）时当作跳过，让任务按原来的方式停下来。"""
    global _pending, _answer, _seq
    auto = AUTO_ANSWER if AUTO_ANSWER is not None else ("skip" if debug_log._under_test() else "")
    if auto:
        return auto
    with _COND:
        if _run["auto_skip"]:
            debug_log.event("⏭ 文件出错，按用户选择自动跳过", 文件=label, 原因=error)
            return "skip"
        _seq += 1
        _run["asked"] += 1
        problem_id = f"p{_seq}_{int(time.time())}"
        _pending = {"id": problem_id, "task": _run["task"], "path": str(path), "label": label,
                    "name": re.split(r"[\\/]", str(path))[-1], "folder": str(folder or Path(path).parent),
                    "error": error, "since": time.time(), "count": _run["asked"]}
        _answer = None
    debug_log.event("⏸ 文件出错，暂停等用户选择重试或跳过", 任务=_run["task"], 文件=label, 原因=error)
    started = time.monotonic()
    choice = "skip"
    try:
        while True:
            minutes = get_settings()["timeout_minutes"]
            if minutes and time.monotonic() - started >= minutes * 60:
                debug_log.event(f"⏰ {minutes} 分钟没人选择，自动跳过这份文件", 文件=label)
                break
            with _COND:
                if _answer is not None and _answer["id"] == problem_id:
                    choice = _answer["choice"]
                    if _answer["auto_skip"]:
                        _run["auto_skip"] = True
                    break
                if _pending is None or _pending["id"] != problem_id:
                    break  # 任务被收尾清掉了
                _COND.wait(1.0)
            if tick is not None:
                try:
                    tick()
                except Exception:
                    pass
            if should_stop is not None:
                try:
                    if should_stop():
                        debug_log.event("⏹ 等待中用户要求停止任务，这份文件按跳过处理", 文件=label)
                        break
                except Exception:
                    pass
    finally:
        with _COND:
            if _pending is not None and _pending["id"] == problem_id:
                _pending, _answer = None, None
    debug_log.event("▶️ 用户选择" + ("重试" if choice == "retry" else "跳过"), 文件=label,
                    等了秒=round(time.monotonic() - started), 后面都跳过=_run["auto_skip"])
    return choice


def message_for(err: BaseException, watchdog: Any = None) -> str:
    """出错原因：看门狗发现 Excel 弹了框 / 没反应时，说清楚是怎么回事，不只给一串 COM 错误码"""
    if watchdog is not None:
        try:
            text = watchdog.explain(err)
            if isinstance(text, str) and text:
                return text
        except Exception:
            pass
    return str(err) or type(err).__name__


class Guard:
    """修复类页面（修复颜色 / 引用错误 / 隐藏页 / BL 超链接）的后台任务用：

    - 盯着后台 Excel 有没有弹出没人能点的提示框（不设空闲超时：这些任务以前就没有，大文件慢慢做不打断）
    - 某份文件出错时关掉后台 Excel、弹框问用户重试还是跳过

    用法：guard = Guard("扫描并修复颜色")；开 Excel 后 guard.watch(excel)；每份文件开始 guard.progress()；
    出错时 retry, message = guard.problem(err, path, label, folder, close=...)；最后 guard.stop()。"""

    def __init__(self, title: str) -> None:
        from core.excel_watchdog import ExcelWatchdog

        begin(title)
        self.title = title
        self.closed = False
        self.watchdog = ExcelWatchdog(
            lambda: debug_log.event(f"⚠️ {title}：后台 Excel 弹出了没人能点的提示框，已结束它", 内容=self.watchdog.dialog),
            timeout=86400)
        self.watchdog.start()

    def watch(self, excel: Any) -> Any:
        if excel is not None:
            try:
                self.watchdog.register(excel)
            except Exception as err:  # 看不住也照常做，只是卡住时没人救
                debug_log.event("⚠️ 没能盯住后台 Excel 的提示框", 原因=str(err))
        return excel

    def progress(self) -> None:
        self.watchdog.progress()

    def problem(self, err: BaseException, path: str | Path, label: str, folder: str | Path,
                close: Callable[[], None] | None = None) -> tuple[bool, str]:
        """返回 (要不要重试, 出错原因)。close 关掉后台 Excel：要问用户时先关（别占着文件），
        Excel 报的错也关（可能已经坏了，下一份换新的）；程序自己的检查结论（ValueError）不关。
        关了的话 self.closed 为真，调用方把手里的 Excel 置空、下次重开"""
        message = message_for(err, self.watchdog)
        asking = should_ask(err, message)
        self.closed = asking or not isinstance(err, ValueError)
        if self.closed and close is not None:
            try:
                close()
            except Exception:
                pass
        if not asking:
            return False, message
        choice = ask(path, label, message, folder=folder, tick=self.watchdog.progress)
        self.watchdog.progress()
        return choice == "retry", message

    def stop(self) -> None:
        try:
            self.watchdog.stop()
        finally:
            finish()


def quit_quietly(excel: Any) -> None:
    """后台 Excel 可能已经被看门狗结束了：退出时出错不管"""
    if excel is None:
        return
    from core.excel_ops import quit_background_excel
    try:
        quit_background_excel(excel)
    except Exception:
        pass
