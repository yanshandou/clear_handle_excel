import json
import sqlite3
from typing import Any

from core.common import PROJECT_ROOT

DATABASE_PATH = PROJECT_ROOT / "execution.sqlite3"


def _connect() -> sqlite3.Connection:
    connection = sqlite3.connect(DATABASE_PATH, timeout=30)
    connection.execute("PRAGMA synchronous = FULL")
    connection.execute(
        "CREATE TABLE IF NOT EXISTS snapshots "
        "(name TEXT PRIMARY KEY, payload TEXT NOT NULL, updated_at TEXT NOT NULL)"
    )
    return connection


def save_snapshot(name: str, state: dict[str, Any]) -> None:
    payload = json.dumps(state, ensure_ascii=False, allow_nan=False)
    connection = _connect()
    try:
        with connection:
            connection.execute(
                "INSERT INTO snapshots(name, payload, updated_at) "
                "VALUES (?, ?, datetime('now')) "
                "ON CONFLICT(name) DO UPDATE SET "
                "payload=excluded.payload, updated_at=excluded.updated_at",
                (name, payload),
            )
    finally:
        connection.close()


def load_snapshot(name: str) -> dict[str, Any] | None:
    connection = _connect()
    try:
        row = connection.execute(
            "SELECT payload FROM snapshots WHERE name = ?", (name,)
        ).fetchone()
    finally:
        connection.close()
    if row is None:
        return None
    state = json.loads(row[0])
    if not isinstance(state, dict):
        raise ValueError(f"Invalid database snapshot: {name}")
    return state
