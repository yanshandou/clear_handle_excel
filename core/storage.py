import json
import sqlite3
import time
from typing import Any

from core import data_dir

# 放在程序文件夹外（见 core/data_dir.py），重新下载 ZIP 更新程序不用搬
DATABASE_PATH = data_dir.DATA_DIR / data_dir.DATABASE_NAME

# 每个数据库里已写进 blobs 表的大字段：(数据库, 快照名) -> {键: 写入时的字符串对象}
# 字符串不可变，同一个对象就说明没改过，不用每次都重写几 MB 的截图
_WRITTEN_BLOBS: dict[tuple[str, str], dict[str, str]] = {}


def _connect() -> sqlite3.Connection:
    if DATABASE_PATH == data_dir.DATA_DIR / data_dir.DATABASE_NAME:
        data_dir.ensure_migrated(DATABASE_PATH, data_dir.DATA_DIR / data_dir.IMAGES_NAME)
    connection = sqlite3.connect(DATABASE_PATH, timeout=30)
    connection.execute("PRAGMA synchronous = FULL")
    connection.execute(
        "CREATE TABLE IF NOT EXISTS snapshots "
        "(name TEXT PRIMARY KEY, payload TEXT NOT NULL, updated_at TEXT NOT NULL)"
    )
    connection.execute(
        "CREATE TABLE IF NOT EXISTS blobs "
        "(name TEXT NOT NULL, key TEXT NOT NULL, data TEXT NOT NULL, PRIMARY KEY (name, key))"
    )
    return connection


def save_snapshot(
    name: str, state: dict[str, Any], *, blob_field: str | None = None, blob_collections: tuple[str, ...] = (),
) -> None:
    """保存快照。给了 blob_field 时，state[集合][id][blob_field]（截图）单独存，只写有变化的。"""
    started = time.perf_counter()
    blobs: dict[str, str] = {}
    if blob_field:
        slim = dict(state)
        for collection in blob_collections:
            slim[collection] = {}
            for key, entry in state.get(collection, {}).items():
                if isinstance(entry, dict) and entry.get(blob_field):
                    blobs[f"{collection}\t{key}\t{blob_field}"] = entry[blob_field]
                    entry = {k: v for k, v in entry.items() if k != blob_field}
                slim[collection][key] = entry
        state = slim
    payload = json.dumps(state, ensure_ascii=False, allow_nan=False)
    cache_key = (str(DATABASE_PATH), name)
    written = _WRITTEN_BLOBS.get(cache_key, {})
    changed = [(name, key, value) for key, value in blobs.items() if written.get(key) is not value]
    removed = [(name, key) for key in written if key not in blobs]
    connection = _connect()
    try:
        with connection:
            if blob_field and cache_key not in _WRITTEN_BLOBS:
                # 这个进程第一次写：不知道库里原来有什么，整个换掉
                connection.execute("DELETE FROM blobs WHERE name = ?", (name,))
            connection.executemany(
                "INSERT INTO blobs(name, key, data) VALUES (?, ?, ?) "
                "ON CONFLICT(name, key) DO UPDATE SET data=excluded.data",
                changed,
            )
            connection.executemany("DELETE FROM blobs WHERE name = ? AND key = ?", removed)
            connection.execute(
                "INSERT INTO snapshots(name, payload, updated_at) "
                "VALUES (?, ?, datetime('now')) "
                "ON CONFLICT(name) DO UPDATE SET "
                "payload=excluded.payload, updated_at=excluded.updated_at",
                (name, payload),
            )
    finally:
        connection.close()
    _note_slow_save(name, started, len(payload), len(changed))
    if blob_field:
        _WRITTEN_BLOBS[cache_key] = blobs


def load_snapshot(name: str, *, include_blobs: bool = True) -> dict[str, Any] | None:
    connection = _connect()
    try:
        row = connection.execute(
            "SELECT payload FROM snapshots WHERE name = ?", (name,)
        ).fetchone()
        blob_rows = connection.execute("SELECT key, data FROM blobs WHERE name = ?", (name,)).fetchall() if include_blobs else []
    finally:
        connection.close()
    if row is None:
        return None
    state = json.loads(row[0])
    if not isinstance(state, dict):
        raise ValueError(f"Invalid database snapshot: {name}")
    for key, data in blob_rows:
        collection, item_id, field = key.split("\t")
        entry = state.get(collection, {}).get(item_id)
        if isinstance(entry, dict):
            entry.setdefault(field, data)
    return state


SLOW_SAVE_SECONDS = 0.2  # 存一次记录超过这么久写调试日志（排查扫描中途空等，用户 2026-11-03 要求）


def _note_slow_save(name: str, started: float, size: int, blobs: int) -> None:
    seconds = time.perf_counter() - started
    if seconds < SLOW_SAVE_SECONDS:
        return
    try:
        from core import debug_log
        debug_log.event("💾 存记录慢", 名=name, 秒=round(seconds, 2), 大小KB=size // 1024, 新截图=blobs)
    except Exception:
        pass
