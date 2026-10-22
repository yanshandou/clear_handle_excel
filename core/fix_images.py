"""报告用的修复前 / 修复后截图：存成 PNG 文件，问题项里只记文件名（item["fix_images"]）。

截图会被后面的复检、重新扫描覆盖，所以在覆盖前把第一次的问题截图存成「修复前」，
修好时把当时的截图存成「修复后」。下载报告时打进压缩包。
"""
import base64
import hashlib
import logging
import shutil
from pathlib import Path
from typing import Any

from core.common import PROJECT_ROOT

IMAGES_DIR = PROJECT_ROOT / "_report_images"


def save(item: dict[str, Any], which: str, data_url: str, *, overwrite: bool = True, key: str = "") -> None:
    """which：before | after。data_url 是 data:image/png;base64,...；空图不存。
    key 相同的问题项共用同一个文件（同一行的几处问题）"""
    names = item.setdefault("fix_images", {})
    if not data_url or "," not in data_url or (not overwrite and names.get(which)):
        return
    digest = hashlib.sha1((key or f"{item.get('wb_path', '')}\0{item['id']}").encode()).hexdigest()[:16]
    name = f"{digest}_{which}.png"
    try:
        IMAGES_DIR.mkdir(exist_ok=True)
        (IMAGES_DIR / name).write_bytes(base64.b64decode(data_url.split(",", 1)[1]))
    except (OSError, ValueError):
        logging.exception("无法保存报告截图: %s", name)
        return
    names[which] = name


def remember_before(item: dict[str, Any]) -> None:
    """截图要被覆盖前调用：还没存过修复前截图、且这处还没解决时，把现在的截图存下来"""
    if item.get("status") not in {"fixed", "accepted"}:
        save(item, "before", item.get("image") or "", overwrite=False)


def forget(item: dict[str, Any], which: str) -> None:
    (item.get("fix_images") or {}).pop(which, None)


def image_path(name: str) -> Path | None:
    if not name or Path(name).name != name:
        return None
    path = IMAGES_DIR / name
    return path if path.is_file() else None


def clear() -> None:
    """换了文件夹：旧文件夹的截图不会再用到"""
    shutil.rmtree(IMAGES_DIR, ignore_errors=True)
