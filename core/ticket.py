"""将确认后的 Ticket 表单保存到 Windows 桌面。"""

from __future__ import annotations

import re
import threading
from contextlib import suppress
from itertools import count
from pathlib import Path

from core.common import sanitize_filename

DETAILS = {
    "title": "任务标题",
    "project": "所属项目",
    "parent_project": "上级项目名称",
    "parent": "关联案件编号",
    "parent_category": "关联案件分类",
    "parent_title": "关联案件标题",
    "status": "状态",
    "priority": "优先级",
    "assignee": "负责人",
    "assignee_short": "负责人简称",
    "start_date": "开始日期",
    "due_date": "截止日期",
    "progress": "进度",
    "version": "目标版本",
    "screen_id": "画面 / 报表 ID",
    "estimated_hours": "预计工数（人日）",
    "actual_hours": "实际工数（人日）",
}
INFO_FILE = "Ticket.txt"
_CREATE_LOCK = threading.Lock()


def _desktop() -> Path:
    from win32com.shell import shell, shellcon

    return Path(shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOPDIRECTORY, 0, 0)).resolve()


def _field(payload: dict, name: str, label: str) -> str:
    value = payload.get(name, "")
    if not isinstance(value, str) or len(value) > 500 or any(ord(char) < 32 for char in value):
        raise ValueError(f"{label}请填写为 500 字以内的单行文本")
    return value.strip()


def _same_ticket(folder: Path, info: Path, content: bytes) -> bool:
    # 文件夹已存在时，只复用桌面上的普通文件夹和信息文件。
    try:
        if folder.resolve() == folder and info.resolve() == info and info.is_file():
            with info.open("rb") as stream:
                return stream.read(len(content) + 1) == content
    except (OSError, RuntimeError):
        pass
    return False


def create(payload: dict) -> dict[str, str]:
    details = {name: _field(payload, name, label) for name, label in DETAILS.items()}
    if not details["title"]:
        raise ValueError("请填写任务标题")
    number = _field(payload, "ticket_id", "任务编号").translate(str.maketrans("＃０１２３４５６７８９", "#0123456789"))
    if number and not re.fullmatch(r"#?\s*[0-9]{1,20}", number):
        raise ValueError("任务编号请填写为 1 到 20 位数字")
    number = number.removeprefix("#").strip()

    heading = f"Ticket #{number}" if number else "Ticket"
    lines = [heading, ""]
    lines.extend(f"{label}：{details[name]}" for name, label in DETAILS.items() if details[name])
    content = ("\r\n".join(lines) + "\r\n").encode("utf-8-sig")
    title = sanitize_filename(details["title"])[:70].rstrip(" .")
    name = f"{heading} - {title}" if title else heading
    desktop = _desktop()

    with _CREATE_LOCK:
        for index in count(1):
            folder = desktop / (name if index == 1 else f"{name} ({index})")
            info = folder / INFO_FILE
            try:
                folder.mkdir()
            except FileExistsError:
                if _same_ticket(folder, info, content):
                    break
                continue
            created_info = False
            try:
                with info.open("xb") as stream:
                    created_info = True
                    stream.write(content)
            except OSError:
                if created_info:
                    with suppress(OSError):
                        info.unlink(missing_ok=True)
                with suppress(OSError):
                    folder.rmdir()
                raise
            break

    return {"folder": str(folder), "filename": INFO_FILE}


def open_folder(payload: dict) -> dict:
    import os

    target = payload.get("folder", "")
    if not isinstance(target, str) or not target:
        raise ValueError("没有可以打开的文件夹")
    folder = Path(target).resolve()
    desktop = _desktop()
    if folder != desktop and desktop not in folder.parents:
        raise ValueError("只能打开桌面上创建的 Ticket 文件夹")
    if not folder.is_dir():
        raise ValueError("文件夹不存在，可能已被移动或删除")
    os.startfile(folder)
    return {}
