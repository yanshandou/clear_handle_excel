"""同名 PDF 缓存（用户 2026-11-10）：Excel 没变、导出设置没变，就直接用上次导出的 PDF，不再重新处理、重新导。

怎么认「这份 PDF 就是这份 Excel 导的」：导出时把 Excel 的哈希（SHA-256）记成这份 PDF 的身份，
连同导出设置（跳过哪些表、隐藏表、页眉页脚）和 PDF 自己的哈希一起记下。下次要用时三样都对上才算：
- Excel 现在的哈希 = 记下的（Excel 一改，哪怕只改一个字，哈希就不一样，PDF 重导）；
- 导出设置和记下的一样；
- PDF 还在、内容没被换过（PDF 的哈希也对得上）。

上传时「表格和 PDF 一起」会先在 Excel 上做上传前处理、保存，保存后文件的哈希会变，
所以一份 PDF 可以记几个 Excel 哈希（处理前、处理后），处理文件夹里现在是哪一版都认得出来。
记录放在程序数据文件夹的 pdf_cache.json，按 PDF 路径存，只留最近 MAX_ENTRIES 条。
"""
from __future__ import annotations

import hashlib
import json
import os
import threading
from datetime import datetime
from pathlib import Path
from typing import Any

from core import debug_log

_LOCK = threading.RLock()
MAX_ENTRIES = 20000
MAX_HASHES = 6


def _store_path() -> Path:
    from core import storage
    return Path(storage.DATABASE_PATH).parent / "pdf_cache.json"


def _key(pdf: Path) -> str:
    return os.path.normcase(os.path.abspath(str(pdf)))


def options_key(options: dict[str, Any]) -> str:
    return json.dumps(options, ensure_ascii=False, sort_keys=True)


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as fh:
        for block in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def _load() -> dict[str, Any]:
    try:
        data = json.loads(_store_path().read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def _save(data: dict[str, Any]) -> None:
    path = _store_path()
    if len(data) > MAX_ENTRIES:
        data = dict(sorted(data.items(), key=lambda kv: kv[1].get("at", ""))[-MAX_ENTRIES:])
    tmp = path.with_name(path.name + f".{os.getpid()}.tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, path)
    except OSError as err:
        tmp.unlink(missing_ok=True)
        debug_log.exception(err, "保存同名 PDF 缓存记录")


def remember(pdf: Path, excel_hashes: list[str], options: dict[str, Any], info: dict[str, Any], source: str) -> None:
    """记下 pdf 是由（哈希为 excel_hashes 之一的）Excel 按 options 导出的。info 里放页数、表（名单或张数）、跳过的表"""
    hashes = [h for h in dict.fromkeys(excel_hashes) if h]
    if not hashes or not Path(pdf).is_file():
        return
    try:
        pdf_sha = digest(Path(pdf))
    except OSError:
        return
    sheets = info.get("sheets") or 0
    sheets = len(sheets) if isinstance(sheets, list) else int(sheets)  # 批量导出给的是表名单，上传给的是张数
    with _LOCK:
        data = _load()
        old = data.get(_key(pdf)) or {}
        if old.get("pdf_sha256") == pdf_sha and old.get("options") == options_key(options):
            hashes = list(dict.fromkeys(hashes + old.get("excel_sha256", [])))[:MAX_HASHES]  # 同一份 PDF 再认一个 Excel 版本
        data[_key(pdf)] = {"excel_sha256": hashes, "options": options_key(options), "pdf_sha256": pdf_sha,
                           "pages": info.get("pages", 0), "sheets": sheets,
                           "skipped": list(info.get("skipped") or []), "source": source,
                           "at": datetime.now().isoformat(timespec="seconds")}
        _save(data)


def add_excel_hash(pdf: Path, excel_hash: str) -> None:
    """同一份 PDF 再认一个 Excel 版本（上传前处理保存后 Excel 哈希变了，但 PDF 就是按这版内容导的）"""
    with _LOCK:
        data = _load()
        entry = data.get(_key(pdf))
        if not entry or not excel_hash or excel_hash in entry["excel_sha256"]:
            return
        entry["excel_sha256"] = [excel_hash, *entry["excel_sha256"]][:MAX_HASHES]
        _save(data)


def lookup(pdf: Path, excel_hash: str, options: dict[str, Any]) -> dict[str, Any] | None:
    """PDF 还在、内容没换、是这版 Excel 按同样设置导的 → 返回记录；否则 None（要重新导）"""
    with _LOCK:
        entry = _load().get(_key(pdf))
    if not entry or excel_hash not in entry.get("excel_sha256", []) or entry.get("options") != options_key(options):
        return None
    try:
        if not Path(pdf).is_file() or digest(Path(pdf)) != entry.get("pdf_sha256"):
            return None
    except OSError:
        return None
    return entry


def forget(pdf: Path) -> None:
    with _LOCK:
        data = _load()
        if data.pop(_key(pdf), None) is not None:
            _save(data)
