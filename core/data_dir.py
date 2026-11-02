"""数据文件夹：execution.sqlite3 和 _report_images/ 放在程序文件夹外面，换新版本（重新下载 ZIP）不用再搬。

位置：
- 设了环境变量 CLEAR_HANDLE_EXCEL_DATA 就用它；
- Windows：%LOCALAPPDATA%\\clear_handle_excel（一般是 C:\\Users\\<用户名>\\AppData\\Local\\clear_handle_excel）；
- 其他系统：~/.clear_handle_excel。

第一次用到数据库时（数据文件夹里还没有 execution.sqlite3），自动从旧位置复制一份过来：
先找本程序文件夹，再找旁边解压出来的其他版本（同一个上级目录下、或 ZIP 套了一层文件夹的），
挑最近改过的那份。只复制不删除，旧文件原样留着。
"""
import logging
import os
import shutil
import sqlite3
import sys
import threading
from pathlib import Path

from core.common import PROJECT_ROOT, console_print

DATABASE_NAME = "execution.sqlite3"
IMAGES_NAME = "_report_images"


def _default_dir() -> Path:
    custom = os.environ.get("CLEAR_HANDLE_EXCEL_DATA", "").strip()
    if custom:
        return Path(custom).expanduser()
    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA", "").strip()
        if base:
            return Path(base) / "clear_handle_excel"
    return Path.home() / ".clear_handle_excel"


DATA_DIR = _default_dir()

_lock = threading.Lock()
_checked: set[str] = set()


def _is_program_dir(folder: Path) -> bool:
    return (folder / "core" / "storage.py").is_file() and (folder / "web_clip_fixer.py").is_file()


def _modified(database: Path) -> float:
    times = []
    for path in (database, database.with_name(database.name + "-wal")):
        try:
            times.append(path.stat().st_mtime)
        except OSError:
            pass
    return max(times, default=0.0)


def legacy_candidates(program_root: Path = PROJECT_ROOT) -> list[Path]:
    """可能存着旧数据库的地方：本程序文件夹优先，再按修改时间从新到旧排其他版本的程序文件夹"""
    own = program_root / DATABASE_NAME
    found: dict[str, Path] = {}
    parent = program_root.parent
    # Downloads\\clear_handle_excel-main (1)\\ 和 Downloads\\clear_handle_excel-main (1)\\clear_handle_excel-main\\ 两种解压法
    patterns = [(parent, f"*/{DATABASE_NAME}"), (parent.parent, f"*/*/{DATABASE_NAME}")]
    for base, pattern in patterns:
        if base == base.parent and pattern.startswith("*/*/"):
            continue
        try:
            for path in base.glob(pattern):
                if path != own and path.is_file() and _is_program_dir(path.parent):
                    found[str(path.resolve())] = path
        except OSError:
            continue
    others = sorted(found.values(), key=_modified, reverse=True)
    return ([own] if own.is_file() else []) + others


def _copy_database(source: Path, target: Path) -> None:
    """用 SQLite 自己的备份复制，旧库还有没合并的 -wal 也不会丢"""
    temp = target.with_name(target.name + ".copying")
    temp.unlink(missing_ok=True)
    src = sqlite3.connect(f"file:{source.as_posix()}?mode=ro", uri=True, timeout=30)
    try:
        dst = sqlite3.connect(temp)
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()
    os.replace(temp, target)


def ensure_migrated(database: Path, images: Path, program_root: Path = PROJECT_ROOT) -> None:
    """数据文件夹里还没有数据库时，从旧位置复制数据库和截图过来。每个进程每个位置只检查一次。"""
    key = str(database)
    if key in _checked:
        return
    with _lock:
        if key in _checked:
            return
        _checked.add(key)
        try:
            database.parent.mkdir(parents=True, exist_ok=True)
        except OSError:
            logging.exception("无法创建数据文件夹: %s", database.parent)
            return
        if database.exists():
            if (program_root / DATABASE_NAME).is_file():
                console_print(f"ℹ️ 数据库在 {database}，程序文件夹里的 {DATABASE_NAME} 已不再使用，可以删掉。")
            return
        for source in legacy_candidates(program_root):
            try:
                _copy_database(source, database)
            except (OSError, sqlite3.Error):
                logging.exception("复制旧数据库失败: %s", source)
                continue
            old_images = source.parent / IMAGES_NAME
            if old_images.is_dir():
                try:
                    shutil.copytree(old_images, images, dirs_exist_ok=True)
                except OSError:
                    logging.exception("复制旧截图失败: %s", old_images)
            console_print(f"📦 已把旧数据库复制到 {database}（来自 {source}），以后更新程序不用再搬。")
            return
