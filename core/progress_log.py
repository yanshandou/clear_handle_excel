"""后台批量任务的「每份做完的结果」清单，给网页上的实时进度框用（用户 2026-11-09 要：所有批量操作都要实时进度框）。

各个模块在自己的 _STATUS 里放一个 log 列表，做完一份 Excel 就调 add() 追加一条：
{"name": 显示的文件名, "ok": 成功没, "reason": 结果或原因, "path": 完整路径（能点开 Excel）, "sheet": 表名}。
网页的进度框（static/job_panel.js）把它倒过来列，最新的在最上面。调用方持有自己模块的锁。
"""
from __future__ import annotations

from typing import Any

LIMIT = 500  # 一次任务最多留这么多条，几千份文件时网页每秒取的数据也不会太大


def add(status: dict[str, Any], name: str, ok: bool, reason: str = "", *, path: str = "", sheet: str = "") -> None:
    log = status.setdefault("log", [])
    log.append({"name": str(name or ""), "ok": bool(ok), "reason": str(reason or ""),
                "path": str(path or ""), "sheet": str(sheet or "")})
    if len(log) > LIMIT:
        del log[:-LIMIT]


def step(status: dict[str, Any], text: str) -> None:
    """这份 Excel 现在做到哪一步（进度框第二行）"""
    status["step"] = str(text or "")
