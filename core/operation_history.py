"""Persistent history of major user operations, independent of the current report."""
import logging
import os
from datetime import datetime
from core import storage

LABELS = {
    "clean": "精简 Sheet", "scan": "扫描切字", "autofix": "自动修复切字",
    "pdf": "导出 PDF", "view": "视图标准化", "links": "修复引用错误",
    "colors": "修复颜色", "hidden": "修复隐藏页", "bl_scan": "排查 BL 超链接",
    "bl_clear": "清除 BL 超链接",
}

def _connect():
    db = storage._connect()
    db.execute("CREATE TABLE IF NOT EXISTS operation_history (id INTEGER PRIMARY KEY, folder TEXT NOT NULL, folder_key TEXT NOT NULL, kind TEXT NOT NULL, started_at TEXT NOT NULL, status TEXT NOT NULL)")
    return db

def begin(folder, kinds):
    try:
        folder = os.path.abspath(str(folder))
        db = _connect()
        try:
            with db:
                return [db.execute("INSERT INTO operation_history(folder, folder_key, kind, started_at, status) VALUES (?, ?, ?, ?, 'running')", (folder, os.path.normcase(folder), kind, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))).lastrowid for kind in kinds]
        finally:
            db.close()
    except Exception:
        logging.exception("无法保存用户操作记录")
        return []

def finish(ids, status):
    if not ids:
        return
    try:
        db = _connect()
        try:
            with db:
                db.executemany("UPDATE operation_history SET status=? WHERE id=?", [(status, i) for i in ids])
        finally:
            db.close()
    except Exception:
        logging.exception("无法更新用户操作记录")

def interrupt_running():
    db = _connect()
    try:
        with db:
            db.execute("UPDATE operation_history SET status='interrupted' WHERE status='running'")
    finally:
        db.close()

def summary():
    db = _connect()
    try:
        rows = db.execute("SELECT folder, folder_key, kind, started_at, status FROM operation_history ORDER BY id DESC").fetchall()
    finally:
        db.close()
    folders = {}
    for folder, key, kind, started, status in rows:
        group = folders.setdefault(key, {"folder": folder, "operations": {}})
        entry = group["operations"].setdefault(kind, {"count": 0, "last": started, "status": status})
        entry["count"] += 1
    return {"labels": LABELS, "folders": list(folders.values())}
