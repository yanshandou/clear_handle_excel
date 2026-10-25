"""调试日志：公司的 Excel 不能发出去，出问题时靠这份日志远程排查。

- 全部写进程序目录的 log.txt，每条一行（不换行），开头用 Emoji 标类型：
  🧭 版本环境 ⚙️ Excel 📂 打开文件 📄 工作表 🔍 检出 🧪 试算 ✍️ 写入 💾 保存 🎨 标色 🧽 还原 🔎 兜底查找
  🖱️ 网页操作 🚫 操作被拒绝 💬 网页实时日志 ⏱ 耗时 ✅ 完成 🪟 你的 Excel ❌ 异常（含调用栈和 COM 错误码）🧵 线程现场
- 一件事牵涉多条日志时用分隔线包起来，可以嵌套（任务 > 文件 > 行）：
  〰〰〰〰〰〰〰〰〰〰 ⬇任务 batch_scan #20261024_183012 开始⬇ 〰〰〰〰〰〰〰〰〰〰
  〰〰〰〰〰〰〰〰〰〰 ⬆任务 batch_scan #20261024_183012 结束⬆ 〰〰〰〰〰〰〰〰〰〰
  网页上「复制调试日志」按任务的分隔线把那一次任务整段挑出来。
- 每行都带完整日期时间，一直往同一个 log.txt 追加；程序启动、每个任务开始和之后每小时，删掉 KEEP_HOURS 小时以前的行
  （正在跑的任务从开始起都留着）。
- 单元格内容只记前 TEXT_PREVIEW_CHARS 个字和总长度（clip），文件名、表名、单元格地址照常记。
- 默认开启；环境变量 CLIP_DEBUG_LOG=0 可关闭。跑回归测试时自动不写。
"""
import functools
import json
import os
import platform
import re
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

LOG_PATH = PROJECT_ROOT / "log.txt"
KEEP_HOURS = 12
PRUNE_EVERY_SECONDS = 3600
MAX_TASK_BYTES = 50 * 1024 * 1024  # 一个任务写到这么大后只再记报错，防止把磁盘写满
MAX_LINE_CHARS = 8000
TEXT_PREVIEW_CHARS = 12  # 单元格内容只记这么多字，不把公司数据整段写进日志
QUIET_DUMP_SECONDS = 180  # 任务进行中这么久没有任何新日志：记下每个线程停在哪一行（卡住时最有用）
COPY_HEAD_BYTES = 8_000
COPY_ERRORS_BYTES = 60_000
COPY_TAIL_BYTES = 120_000
WAVE = "〰" * 10
ERROR_MARKS = ("❌", "🧵", "⚠️")  # 复制时这些开头的行单独挑出来

_lock = threading.RLock()
_fh: Any = None
_task: dict[str, Any] | None = None
_open: list[str] = []  # 还没写结束线的分隔块（任务不算在内），外层在前
_last_write = time.monotonic()
_last_prune = 0.0
_quiet_dumped = False
# 已经写过调用栈的异常，同一个异常只写一次。不留异常对象本身（会连带留住 Excel 的 COM 对象），
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


def _one_line(text: str) -> str:
    return " ⏎ ".join(part.strip() for part in str(text).splitlines() if part.strip()) or ""


def _emit(raw: str, *, force: bool = False) -> None:
    """写一行。raw 已经排好（Emoji 开头或分隔线）"""
    global _fh, _last_write, _quiet_dumped
    if not enabled():
        return
    if len(raw) > MAX_LINE_CHARS:
        raw = raw[:MAX_LINE_CHARS] + f"……（这一行共 {len(raw)} 字，后面省略）"
    line = raw + "\n"
    with _lock:
        try:
            if _task is not None:
                if _task["bytes"] >= MAX_TASK_BYTES and not force:
                    if _task["full"]:
                        return
                    _task["full"] = True
                    line = f"⚠️ 这个任务的日志已到 {MAX_TASK_BYTES // 2**20} MB，之后只记报错\n"
                _task["bytes"] += len(line.encode("utf-8"))
            if _fh is None:
                _fh = LOG_PATH.open("a", encoding="utf-8", buffering=1)  # 按行写入：程序崩溃时日志也在
            _fh.write(line)
            _last_write = time.monotonic()
            _quiet_dumped = False
        except OSError:
            pass  # 写不了调试日志（磁盘满、被占用）不能影响任务本身


def _thread() -> str:
    name = threading.current_thread().name
    return "网页" if name.endswith("(process_request_thread)") else name


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]


def _line(emoji: str, text: str, *, force: bool = False) -> None:
    stamp = _now()
    _emit(f"{emoji} [{stamp}] [{_thread()}] {_one_line(text)}", force=force)


def _wave(arrow: str, title: str, *, force: bool = True) -> None:
    stamp = _now()
    word = "开始" if arrow == "⬇" else "结束"
    _emit(f"{WAVE} {arrow}{title} {word}{arrow} {stamp} {WAVE}", force=force)


@_never_raise
def write(message: str, emoji: str = "ℹ️") -> None:
    _line(emoji, message)


@_never_raise
def event(tag: str, **fields: Any) -> None:
    """一行结构化记录。tag 以 Emoji 开头：「🔍 检出」→ 🔍 [时间] [线程] 检出 key=value key=value"""
    emoji, _, rest = tag.strip().partition(" ")
    if not rest or emoji[0].isalnum():
        emoji, rest = "ℹ️", tag.strip()
    _line(emoji, rest + "".join(f" {key}={_value(value)}" for key, value in fields.items()))


@_never_raise
def begin(title: str, level: int = 1) -> None:
    """开一个分隔块（1 = 文件，2 = 文件里的一行）。同级或更深的块还开着时先写结束线，
    中途 continue、出错跳过也不会少结束线；任务结束时全部收尾"""
    with _lock:
        if _task is not None and _task.get("thread") != threading.get_ident():
            # 别的线程（和后台任务同时进行的查找、预览等）：分隔块是任务线程的，这里只记一行，免得把它的块收掉
            _line("📂", title)
            return
        while len(_open) >= level:
            _wave("⬆", _open.pop())
        _open.append(title)
        _wave("⬇", title)


@_never_raise
def end(level: int = 1) -> None:
    with _lock:
        if _task is not None and _task.get("thread") != threading.get_ident():
            return
        while len(_open) >= level:
            _wave("⬆", _open.pop())


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


def _frames(frames: list[traceback.FrameSummary]) -> str:
    """调用栈写成一行：程序自己的文件只写相对路径，外层在前、出错处在后"""
    root = str(PROJECT_ROOT)
    parts = []
    for frame in frames:
        name = frame.filename
        name = name[len(root) + 1:] if name.startswith(root) else Path(name).name
        code = f" `{frame.line.strip()}`" if frame.line else ""
        parts.append(f"{name}:{frame.lineno} {frame.name}{code}")
    return " → ".join(parts)


def _key(err: BaseException) -> tuple:
    return id(err), type(err).__name__, str(err)[:200], id(err.__traceback__)


@_never_raise
def exception(err: BaseException, context: str = "") -> None:
    """一行写完：异常类型和内容、COM 错误码、调用栈，以及起因（except 里又出错时的前一个异常）。
    同一个异常被几层 except 记录时只写一次"""
    if _key(err) in _dumped:
        return
    _dumped.append(_key(err))
    parts = [f"异常{('（' + context + '）') if context else ''} {type(err).__name__}: {err}"]
    details = com_details(err)
    if details:
        parts.append(f"COM: {details}")
    parts.append("调用栈: " + _frames(traceback.extract_tb(err.__traceback__)))
    earlier = err.__cause__ or err.__context__
    while earlier is not None:
        if _key(earlier) in _dumped:
            parts.append(f"起因: 上面已记的 {type(earlier).__name__}")
            break
        parts.append(f"起因: {type(earlier).__name__}: {earlier} | 调用栈: "
                     + _frames(traceback.extract_tb(earlier.__traceback__)))
        earlier = earlier.__cause__ or earlier.__context__
    _line("❌", " ｜ ".join(parts), force=True)


@_never_raise
def mirror(line: str) -> None:
    """网页实时日志的每一行也写进调试日志；在 except 里记的报错顺带写出调用栈"""
    _line("💬", line)
    err = sys.exc_info()[1]
    if err is not None and any(mark in line for mark in ("❌", "⚠️", "失败", "出错", "中断")):
        exception(err, "上一条日志")


@contextmanager
def timed(label: str, **fields: Any) -> Iterator[dict[str, Any]]:
    """记下一步的开始、结束和耗时；出错时记下调用栈。yield 的 dict 里放的内容会写进结束那一行"""
    extra: dict[str, Any] = {}
    _wave("⬇", label, force=False)
    started = time.perf_counter()
    try:
        yield extra
    except BaseException as err:
        event(f"⏱ {label} 失败", 秒=time.perf_counter() - started, **fields, **extra)
        if isinstance(err, Exception):
            exception(err, label)
        _wave("⬆", label)
        raise
    event(f"⏱ {label}", 秒=time.perf_counter() - started, **fields, **extra)
    _wave("⬆", label, force=False)


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
    event(f"⚙️ Excel 已启动（{role}）", **info)


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


_STAMP = re.compile(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})")


@_never_raise
def prune(now: datetime | None = None) -> None:
    """删掉 KEEP_HOURS 小时以前的行；正在跑的任务从开始那一行起都留着。文件按时间顺序追加，
    找到第一行不早于界线的，前面的整段去掉"""
    global _fh, _last_prune
    now = now or datetime.now()
    cutoff = now.timestamp() - KEEP_HOURS * 3600
    with _lock:
        _last_prune = time.monotonic()
        if _task is not None:
            cutoff = min(cutoff, _task["since"])
        if not LOG_PATH.exists():
            return
        lines = LOG_PATH.read_bytes().decode("utf-8", errors="replace").splitlines(keepends=True)
        keep_from = len(lines)
        for i, line in enumerate(lines):
            match = _STAMP.search(line[:120])
            if match and datetime.strptime(match.group(1), "%Y-%m-%d %H:%M:%S").timestamp() >= cutoff:
                keep_from = i
                break
        if keep_from == 0:
            return
        if _fh is not None:
            _fh.close()
            _fh = None
        temp = LOG_PATH.with_name("log.txt.tmp")
        temp.write_text("".join(lines[keep_from:]), encoding="utf-8")
        temp.replace(LOG_PATH)


@_never_raise
def start_task(kind: str, options: dict[str, Any] | None = None, extra: dict[str, Any] | None = None,
               resumed: bool = False) -> None:
    global _task
    if not enabled():
        return
    with _lock:
        if _task is not None:
            end_task("被下一个任务替换")
        prune()
        title = f"任务 {kind} #{datetime.now():%Y%m%d_%H%M%S}"
        _task = {"title": title, "started": time.perf_counter(), "since": int(time.time()), "thread": threading.get_ident(), "bytes": 0, "full": False}
        _wave("⬇", title)
    write(f"任务 {kind}{'（继续上次任务）' if resumed else ''}", "🧭")
    for line in environment():
        write(line, "🧭")
    for key, value in (extra or {}).items():
        write(f"{key}: {value}", "🧭")
    write(f"任务参数: {summary(options or {})}", "🧭")
    _ensure_monitor()


@_never_raise
def end_task(status: str) -> None:
    global _task
    with _lock:
        if _task is None:
            return
        while _open:  # 不管哪个线程来结束，任务里还开着的块都收尾
            _wave("⬆", _open.pop())
        _line("✅" if status.startswith("completed") else "⏹", f"任务结束：{status}，共 "
              f"{time.perf_counter() - _task['started']:.1f} 秒", force=True)
        _wave("⬆", _task["title"])
        _task = None


@contextmanager
def side_task(kind: str, options: dict[str, Any] | None = None) -> Iterator[None]:
    """不走后台任务槽的操作（查找残留标记色、对比旧版、精简 Sheet 预览……）也当一个任务记，网页上能选、能复制。
    后台任务正在跑时不另起（同一时间只开一个任务），在它里面用分隔线包起来"""
    with _lock:
        own = _task is None and enabled()
    if own:
        start_task(kind, options)
    else:
        _wave("⬇", f"{kind}（和正在跑的任务同时进行）")
        event(f"🧭 {kind} 参数", 参数=summary(options or {}))
    mine = _task
    status = "completed"
    try:
        yield
    except BaseException as err:
        status = f"出错 {type(err).__name__}: {err}"
        if isinstance(err, Exception):
            exception(err, kind)
        raise
    finally:
        if not own:
            _wave("⬆", f"{kind}（和正在跑的任务同时进行）")
        elif _task is mine:  # 期间开始了后台任务时，自己这段已被它收尾，不能把它的任务结束掉
            end_task(status)


@_never_raise
def dump_threads(reason: str) -> None:
    """每个线程正停在哪一行：任务卡住、看门狗触发时，这能直接看出卡在哪个调用上。一个线程一行"""
    names = {thread.ident: thread.name for thread in threading.enumerate()}
    title = f"线程现场：{reason}"
    _wave("⬇", title)
    for ident, frame in sys._current_frames().items():
        name = names.get(ident, str(ident))
        if name == "debug-log-monitor":
            continue
        _line("🧵", f"线程 {name}: " + _frames(traceback.extract_stack(frame)), force=True)
    _wave("⬆", title)


def _watch_quiet() -> None:
    global _quiet_dumped
    while True:
        time.sleep(15)
        with _lock:
            quiet = time.monotonic() - _last_write
            due = _task is not None and not _quiet_dumped and quiet >= QUIET_DUMP_SECONDS
            stale = time.monotonic() - _last_prune >= PRUNE_EVERY_SECONDS
        if stale:
            prune()
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
            _line("⚠️", f"[logging {record.levelname}] {record.getMessage()}", force=True)
            if record.exc_info and record.exc_info[1] is not None:
                exception(record.exc_info[1], "logging")

    logging.getLogger().addHandler(_Handler(logging.WARNING))
    previous = threading.excepthook

    def thread_hook(args: Any) -> None:
        if isinstance(args.exc_value, Exception):
            exception(args.exc_value, f"线程 {getattr(args.thread, 'name', '?')} 未处理的异常")
        previous(args)

    threading.excepthook = thread_hook
    prune()
    write(f"程序启动，版本 {version()}", "🚀")
    _ensure_monitor()


# ---- 网页上选任务、复制、下载 ----

_TASK_START = re.compile(rf"^{WAVE} ⬇任务 (\S+) #(\d{{8}}_\d{{6}}) 开始⬇", re.M)


def _read() -> str:
    return LOG_PATH.read_bytes().decode("utf-8", errors="replace") if LOG_PATH.exists() else ""


def tasks(limit: int = 30) -> list[dict[str, Any]]:
    """log.txt 里最近的任务，新的在前"""
    text = _read()
    found = list(_TASK_START.finditer(text))
    result = []
    for i, match in enumerate(reversed(found)):
        if i >= limit:
            break
        stop = found[len(found) - i].start() if i else len(text)
        result.append({"id": match.group(2), "kind": match.group(1),
                       "kb": max(1, len(text[match.start():stop].encode("utf-8")) // 1024)})
    return result


def task_text(task_id: str | None = None) -> tuple[str, str]:
    """(标题, 内容)：某次任务从开始线到结束线（还没结束就到文件末尾）。没给 id 就是最近一次"""
    text = _read()
    found = list(_TASK_START.finditer(text))
    if task_id:
        found = [m for m in found if m.group(2) == task_id] or found
    if not found:
        return "", ""
    start = found[-1]
    end_mark = f"{WAVE} ⬆任务 {start.group(1)} #{start.group(2)} 结束⬆"
    stop = text.find(end_mark, start.end())
    stop = len(text) if stop < 0 else text.find("\n", stop) + 1 or len(text)
    return f"任务 {start.group(1)} #{start.group(2)}", text[start.start():stop]


def for_copy(title: str, text: str) -> str:
    """复制到剪贴板用：不长就全部；太长时只要开头的版本环境、全部报错和最后一段"""
    data = text.encode("utf-8")
    if len(data) <= COPY_HEAD_BYTES + COPY_ERRORS_BYTES + COPY_TAIL_BYTES:
        return f"### 调试日志 {title}（完整）\n{text}"
    head = data[:COPY_HEAD_BYTES].decode("utf-8", errors="ignore").rsplit("\n", 1)[0]
    errors: list[str] = []
    size = 0
    for line in text.splitlines():
        if line.startswith(ERROR_MARKS):
            size += len(line.encode("utf-8")) + 1
            if size > COPY_ERRORS_BYTES:
                errors.append("……（报错太多，后面省略；完整内容请点「下载调试日志」）")
                break
            errors.append(line)
    tail = data[-COPY_TAIL_BYTES:].decode("utf-8", errors="ignore").split("\n", 1)[-1]
    return "\n".join([
        f"### 调试日志 {title}（共 {len(data) // 1024} KB，太长，只复制开头、全部报错和最后一段；完整版请下载）",
        head, "", "### 报错汇总", *(errors or ["（没有报错）"]), "", "### 最后一段", tail,
    ])
