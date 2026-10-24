"""调试日志：公司的 Excel 不能发出去，出问题时靠这份日志远程排查。

- 每个任务（全自动、只扫描、提交、复检、清除颜色、主页单独任务……）一个文件，放在程序目录的 _debug_logs/，
  最近 KEEP_TASK_LOGS 个保留，更早的自动删除。没有任务时的记录（网页请求出错等）写到 _debug_logs/app.log。
- 比网页实时日志详细：每步耗时、完整 traceback、Excel / COM 错误码、每个检出问题的原因、每次试算的行高、
  写入前后的行高、标色 / 还原的每一格、程序版本和环境。网页上的实时日志照常只留最近几百行，这里不进 /api 快照。
- 单元格内容只记前 TEXT_PREVIEW_CHARS 个字和总长度（clip），文件名、表名、单元格地址照常记。
- 默认开启；环境变量 CLIP_DEBUG_LOG=0 可关闭。跑回归测试时自动不写。
"""
import functools
import json
import os
import platform
import shutil
import subprocess
import sys
import threading
import time
import traceback
from collections import deque
from collections.abc import Iterator
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any

from core import build_info
from core.common import PROJECT_ROOT

DEBUG_DIR = PROJECT_ROOT / "_debug_logs"
APP_LOG = DEBUG_DIR / "app.log"
KEEP_TASK_LOGS = 30
MAX_TASK_LOG_BYTES = 50 * 1024 * 1024  # 一个任务的日志写到这么大后只再记报错，防止把磁盘写满
MAX_APP_LOG_BYTES = 5 * 1024 * 1024
MAX_LINE_CHARS = 4000
TEXT_PREVIEW_CHARS = 12  # 单元格内容只记这么多字，不把公司数据整段写进日志
QUIET_DUMP_SECONDS = 180  # 任务进行中这么久没有任何新日志：记下每个线程停在哪一行（卡住时最有用）
COPY_HEAD_BYTES = 8_000
COPY_ERRORS_BYTES = 60_000
COPY_TAIL_BYTES = 120_000
ERROR_PREFIX = "!! "  # 报错块的第一行，复制时按它把报错单独挑出来

_lock = threading.RLock()
_task: dict[str, Any] | None = None
_last_write = time.monotonic()
_quiet_dumped = False
# 已经写过 traceback 的异常，同一个异常只写一次。不留异常对象本身（会连带留住 Excel 的 COM 对象），
# 只记 id 加类型和内容，id 被新异常重用时也认得出不是同一个
_dumped: deque[tuple] = deque(maxlen=64)
_monitor: threading.Thread | None = None


def _never_raise(func: Any) -> Any:
    """调试日志自己出什么错都不能影响任务本身"""
    @functools.wraps(func)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return func(*args, **kwargs)
        except Exception:
            return None
    return wrapper


def _under_test() -> bool:
    """跑回归测试时不写：免得测试里的假任务混进你真正的调试日志"""
    return "pytest" in sys.modules or "unittest" in (sys.argv[0] if sys.argv else "")


def enabled() -> bool:
    return (os.environ.get("CLIP_DEBUG_LOG", "1").strip() not in {"0", "off", "false", "no"}
            and not _under_test())


def clip(value: Any, limit: int = TEXT_PREVIEW_CHARS) -> str:
    """单元格内容只留开头几个字和总长度，例如 '株式会社テス…'(共 25 字)"""
    if value is None:
        return "None"
    text = str(value)
    if len(text) <= limit:
        return repr(text)
    return f"{text[:limit]!r}…(共 {len(text)} 字)"


def _value(value: Any) -> str:
    if isinstance(value, float):
        return f"{value:.2f}"
    if isinstance(value, (list, tuple, set)) and len(value) > 12:
        items = list(value)
        return f"[{', '.join(map(str, items[:5]))} …共 {len(items)} 项]"
    return str(value)


def _write_line(text: str, *, force: bool = False) -> None:
    global _last_write, _quiet_dumped
    if not enabled():
        return
    stamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
    thread = threading.current_thread().name
    if thread.endswith("(process_request_thread)"):
        thread = "网页"  # 处理网页请求的线程，名字太长
    if len(text) > MAX_LINE_CHARS:
        text = text[:MAX_LINE_CHARS] + f"……（这一行共 {len(text)} 字，后面省略）"
    line = f"[{stamp}] [{thread}] {text}\n"
    with _lock:
        try:
            if _task is not None:
                if _task["bytes"] >= MAX_TASK_LOG_BYTES and not force:
                    if not _task["full"]:
                        _task["full"] = True
                        _task["fh"].write(f"[{stamp}] 调试日志已到 {MAX_TASK_LOG_BYTES // 1024 // 1024} MB，之后只记报错\n")
                    return
                _task["fh"].write(line)
                _task["bytes"] += len(line.encode("utf-8"))
            else:
                DEBUG_DIR.mkdir(exist_ok=True)
                if APP_LOG.exists() and APP_LOG.stat().st_size > MAX_APP_LOG_BYTES:
                    APP_LOG.replace(APP_LOG.with_name("app.old.log"))
                with APP_LOG.open("a", encoding="utf-8") as fh:
                    fh.write(line)
            _last_write = time.monotonic()
            _quiet_dumped = False
        except OSError:
            pass  # 写不了调试日志（磁盘满、被占用）不能影响任务本身


@_never_raise
def write(message: str) -> None:
    for i, part in enumerate(str(message).splitlines() or [""]):
        _write_line(part if i == 0 else "    " + part)


@_never_raise
def event(tag: str, **fields: Any) -> None:
    """一行结构化记录：tag key=value key=value"""
    _write_line(tag + "".join(f" {key}={_value(value)}" for key, value in fields.items()))


@_never_raise
def section(title: str) -> None:
    _write_line("")
    _write_line(f"==================== {title} ====================")


def com_details(err: BaseException) -> str:
    """Excel / COM 错误的 HRESULT 和 Excel 自己给的说明（pywintypes.com_error），沿着异常链都找一遍"""
    parts = []
    seen: set[int] = set()
    current: BaseException | None = err
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        args = getattr(current, "args", ())
        if len(args) >= 2 and isinstance(args[0], int) and type(current).__name__ == "com_error":
            text = f"HRESULT=0x{args[0] & 0xFFFFFFFF:08X}({args[0]}) {args[1]}"
            info = args[2] if len(args) > 2 else None
            if isinstance(info, tuple) and len(info) >= 6:
                scode = info[5]
                text += (f" | 来源={info[1]} 说明={info[2]}"
                         + (f" scode=0x{scode & 0xFFFFFFFF:08X}({scode})" if isinstance(scode, int) else ""))
            if len(args) > 3 and args[3] is not None:
                text += f" 参数位置={args[3]}"
            parts.append(text)
        elif getattr(current, "winerror", None):
            parts.append(f"WinError={current.winerror}")
        current = current.__cause__ or current.__context__
    return "；".join(parts)


def _key(err: BaseException) -> tuple:
    return id(err), type(err).__name__, str(err)[:200], id(err.__traceback__)


@_never_raise
def exception(err: BaseException, context: str = "") -> None:
    """完整 traceback + COM 错误码。同一个异常被几层 except 记录时只写一次"""
    if _key(err) in _dumped:
        return
    _dumped.append(_key(err))
    lines = [f"{ERROR_PREFIX}异常{('：' + context) if context else ''} {type(err).__name__}: {err}"]
    details = com_details(err)
    if details:
        lines.append(f"   COM: {details}")
    # 前一个异常已经写过（except 里又出错）：只写这次的，不再重复前面那段
    earlier = err.__cause__ or err.__context__
    chain = earlier is None or _key(earlier) not in _dumped
    if not chain:
        lines.append(f"   （发生在处理上面那个 {type(earlier).__name__} 的时候）")
    root = str(PROJECT_ROOT)
    for chunk in traceback.format_exception(type(err), err, err.__traceback__, chain=chain):
        for line in chunk.rstrip().splitlines():
            if line.strip() and not line.strip().strip("~^"):
                continue  # Python 3.11 起标出错位置的 ~~~^^^ 行，贴过来只占地方
            lines.append("   " + line.replace(root, "<程序目录>"))
    for line in lines:
        _write_line(line, force=True)


@_never_raise
def mirror(line: str) -> None:
    """网页实时日志的每一行也写进调试日志；在 except 里记的报错顺带写出完整 traceback"""
    _write_line("[日志] " + line)
    err = sys.exc_info()[1]
    if err is not None and any(mark in line for mark in ("❌", "⚠️", "失败", "出错", "中断")):
        exception(err, "上一行日志")


@contextmanager
def timed(label: str, **fields: Any) -> Iterator[dict[str, Any]]:
    """记下一步的开始、结束和耗时；出错时记下 traceback。yield 的 dict 里放的内容会写进结束那一行"""
    extra: dict[str, Any] = {}
    event(f"▶ {label}", **fields)
    started = time.perf_counter()
    try:
        yield extra
    except BaseException as err:
        event(f"✖ {label} 失败", 秒=time.perf_counter() - started, **fields, **extra)
        if isinstance(err, Exception):
            exception(err, label)
        raise
    event(f"◀ {label}", 秒=time.perf_counter() - started, **fields, **extra)


def _git(*args: str) -> str:
    try:
        result = subprocess.run(["git", *args], cwd=PROJECT_ROOT, capture_output=True, text=True, timeout=5,
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return result.stdout.strip() if result.returncode == 0 else ""
    except (OSError, subprocess.SubprocessError):
        return ""


def version() -> str:
    if not build_info.COMMIT.startswith("$"):
        return f"{build_info.COMMIT}（{build_info.COMMIT_DATE}，ZIP 下载）"
    commit = _git("rev-parse", "--short", "HEAD")
    if not commit:
        return "未知（不是从 GitHub ZIP 或 git 取得的副本）"
    dirty = "，有未提交的修改" if _git("status", "--porcelain", "--untracked-files=no") else ""
    return f"{commit}（{_git('log', '-1', '--format=%ci')}，git{dirty}）"


def _package(name: str) -> str:
    try:
        from importlib.metadata import version as package_version
        return package_version(name)
    except Exception:
        return "未安装"


def _memory() -> str:
    if sys.platform != "win32":
        return ""
    try:
        import ctypes

        class _Status(ctypes.Structure):
            _fields_ = [("length", ctypes.c_ulong), ("load", ctypes.c_ulong)] + [
                (name, ctypes.c_ulonglong) for name in
                ("total", "avail", "total_page", "avail_page", "total_virtual", "avail_virtual", "ext")]

        status = _Status(length=ctypes.sizeof(_Status))
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status))
        return f"物理内存 {status.total / 2**30:.1f} GB，可用 {status.avail / 2**30:.1f} GB（占用 {status.load}%）"
    except Exception:
        return ""


def environment() -> list[str]:
    lines = [
        f"程序版本: {version()}",
        f"Python: {sys.version.split()[0]} {platform.architecture()[0]} ({sys.executable})",
        f"系统: {platform.platform()} {' '.join(platform.win32_ver()[:2]) if sys.platform == 'win32' else ''}".rstrip(),
        f"编码: 首选 {__import__('locale').getpreferredencoding(False)}，文件系统 {sys.getfilesystemencoding()}，"
        f"终端 {getattr(sys.stdout, 'encoding', None)}",
        f"依赖: pywin32 {_package('pywin32')}，PyMuPDF {_package('PyMuPDF')}，openpyxl {_package('openpyxl')}",
        f"程序目录: {PROJECT_ROOT}",
    ]
    try:
        usage = shutil.disk_usage(PROJECT_ROOT)
        lines.append(f"磁盘: 剩余 {usage.free / 2**30:.1f} GB / 共 {usage.total / 2**30:.1f} GB")
    except OSError:
        pass
    memory = _memory()
    if memory:
        lines.append(f"内存: {memory}")
    timeout = os.environ.get("CLIP_EXCEL_IDLE_TIMEOUT_SECONDS")
    if timeout:
        lines.append(f"CLIP_EXCEL_IDLE_TIMEOUT_SECONDS={timeout}")
    return lines


@_never_raise
def excel_started(app: Any, role: str = "后台") -> None:
    """记下 Excel 版本：同一个问题在不同版本的 Excel 上表现可能不同"""
    info = {}
    for name in ("Version", "Build", "OperatingSystem", "DecimalSeparator"):
        try:
            info[name] = getattr(app, name)
        except Exception:
            pass
    try:
        info["Hwnd"] = int(app.Hwnd)
    except Exception:
        pass
    event(f"Excel 已启动（{role}）", **info)


def summary(options: dict[str, Any]) -> str:
    def shrink(value: Any) -> Any:
        if isinstance(value, list) and len(value) > 5:
            return value[:3] + [f"…共 {len(value)} 项"]
        if isinstance(value, dict):
            return {key: shrink(item) for key, item in value.items()}
        return value
    try:
        return json.dumps(shrink(options), ensure_ascii=False, default=str)
    except Exception:
        return str(options)[:1000]


def _prune() -> None:
    logs = sorted(DEBUG_DIR.glob("task_*.log"))
    for old in logs[:-KEEP_TASK_LOGS]:
        try:
            old.unlink()
        except OSError:
            pass


@_never_raise
def start_task(kind: str, options: dict[str, Any] | None = None, extra: dict[str, Any] | None = None,
               resumed: bool = False) -> None:
    global _task
    if not enabled():
        return
    with _lock:
        if _task is not None:
            end_task("被下一个任务替换")
        try:
            DEBUG_DIR.mkdir(exist_ok=True)
            path = DEBUG_DIR / f"task_{datetime.now():%Y%m%d_%H%M%S}_{kind}.log"
            fh = path.open("a", encoding="utf-8", buffering=1)  # 按行写入：程序崩溃时日志也在
        except OSError:
            return
        _task = {"path": path, "fh": fh, "kind": kind, "started": time.perf_counter(), "bytes": 0, "full": False}
        _prune()
    section(f"任务 {kind}{'（继续上次任务）' if resumed else ''}")
    for line in environment():
        write(line)
    for key, value in (extra or {}).items():
        write(f"{key}: {value}")
    write(f"任务参数: {summary(options or {})}")
    _ensure_monitor()


@_never_raise
def end_task(status: str) -> None:
    global _task
    with _lock:
        if _task is None:
            return
        _write_line(f"任务结束：{status}，共 {time.perf_counter() - _task['started']:.1f} 秒", force=True)
        try:
            _task["fh"].close()
        except OSError:
            pass
        _task = None


@_never_raise
def dump_threads(reason: str) -> None:
    """每个线程正停在哪一行：任务卡住、看门狗触发时，这能直接看出卡在哪个调用上"""
    names = {thread.ident: thread.name for thread in threading.enumerate()}
    _write_line(f"{ERROR_PREFIX}线程现场：{reason}", force=True)
    for ident, frame in sys._current_frames().items():
        if names.get(ident) == "debug-log-monitor":
            continue
        _write_line(f"   --- 线程 {names.get(ident, ident)} ---", force=True)
        for line in "".join(traceback.format_stack(frame)).rstrip().splitlines():
            _write_line("   " + line, force=True)


def _watch_quiet() -> None:
    global _quiet_dumped
    while True:
        time.sleep(15)
        with _lock:
            quiet = time.monotonic() - _last_write
            due = _task is not None and not _quiet_dumped and quiet >= QUIET_DUMP_SECONDS
        if due:
            dump_threads(f"已经 {quiet:.0f} 秒没有新日志")
            with _lock:
                _quiet_dumped = True  # 有新日志后才会再记


def _ensure_monitor() -> None:
    global _monitor
    with _lock:
        if _monitor is None:
            _monitor = threading.Thread(target=_watch_quiet, name="debug-log-monitor", daemon=True)
            _monitor.start()


def install() -> None:
    """程序启动时调用：没被 except 接住的异常、logging 记的报错也写进调试日志"""
    import logging

    class _Handler(logging.Handler):
        def emit(self, record: logging.LogRecord) -> None:
            _write_line(f"[logging {record.levelname}] {record.getMessage()}", force=True)
            if record.exc_info and record.exc_info[1] is not None:
                exception(record.exc_info[1], "logging")

    logging.getLogger().addHandler(_Handler(logging.WARNING))
    previous = threading.excepthook

    def thread_hook(args: Any) -> None:
        if isinstance(args.exc_value, Exception):
            exception(args.exc_value, f"线程 {getattr(args.thread, 'name', '?')} 未处理的异常")
        previous(args)

    threading.excepthook = thread_hook
    _write_line(f"程序启动，版本 {version()}")


def task_logs() -> list[Path]:
    return sorted(DEBUG_DIR.glob("task_*.log"), reverse=True) if DEBUG_DIR.exists() else []


def latest_log() -> Path | None:
    logs = task_logs()
    return logs[0] if logs else None


def for_copy(path: Path) -> str:
    """复制到剪贴板用：开头的版本环境、全部报错块、最后一段。日志不大时原样全部给出"""
    data = path.read_bytes()
    budget = COPY_HEAD_BYTES + COPY_ERRORS_BYTES + COPY_TAIL_BYTES
    text = data.decode("utf-8", errors="replace")
    if len(data) <= budget:
        return f"### 调试日志 {path.name}（完整）\n{text}"
    lines = text.splitlines()
    head = data[:COPY_HEAD_BYTES].decode("utf-8", errors="ignore").rsplit("\n", 1)[0]
    errors: list[str] = []
    size = 0
    block = False
    for line in lines:
        is_error = ERROR_PREFIX in line[:60]
        if is_error or (block and line.startswith("   ")):
            block = True
            size += len(line.encode("utf-8")) + 1
            if size > COPY_ERRORS_BYTES:
                errors.append("……（报错太多，后面省略；完整内容请点「下载调试日志」）")
                break
            errors.append(line)
        else:
            block = False
    tail = data[-COPY_TAIL_BYTES:].decode("utf-8", errors="ignore").split("\n", 1)[-1]
    return "\n".join([
        f"### 调试日志 {path.name}（共 {len(data) // 1024} KB，太长，只复制开头、全部报错和最后一段；完整版请下载）",
        head, "", "### 报错汇总", *(errors or ["（没有报错）"]), "", "### 最后一段", tail,
    ])
