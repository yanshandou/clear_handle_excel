"""挑文件复制（用户 2026-11-12 要）：贴一串路径 + 一个目标文件夹，保留目录结构，只复制这些文件，
方便把目标文件夹打包发给同事。

- 每行一个文件或文件夹（文件夹 = 里面所有 Excel，含子文件夹）；重复的行只算一次；可贴带引号的路径。
- 目录结构从这些路径共同的上级文件夹算起，连这个上级文件夹自己的名字一起保留：
  …\\doublecheck\\10画面\\01_CM\\a.xlsx + …\\10画面\\03_KS\\b.xlsx → 目标\\10画面\\01_CM\\a.xlsx、目标\\10画面\\03_KS\\b.xlsx
- 只复制，不改来源、不动任务、不记版本；找不到的行列出来，其他照常复制。
- 目标里已有同名文件：先列出来，页面确认后才覆盖。
"""
from __future__ import annotations

import os
import shutil
from pathlib import Path
from typing import Any

from core import debug_log


def _key(path: Path | str) -> str:
    return str(path).replace("\\", "/").casefold()


def _clean(line: str) -> Path:
    return Path(line.strip().strip('"').strip("'")).expanduser()


def _plan(paths_text: Any, dest_text: Any) -> dict[str, Any]:
    from core.distribute import is_excel, scan_folder
    if not isinstance(paths_text, str) or not paths_text.strip():
        raise ValueError("请贴上要复制的文件路径，每行一个")
    if not isinstance(dest_text, str) or not dest_text.strip():
        raise ValueError("请填写目标文件夹")
    dest = _clean(dest_text)
    if dest.exists() and not dest.is_dir():
        raise ValueError(f"目标不是文件夹：{dest}")
    lines = [line for line in paths_text.splitlines() if line.strip()]
    seen: set[str] = set()
    picked: list[Path] = []
    missing: list[dict[str, str]] = []
    for line in lines:
        path = _clean(line)
        if _key(path) in seen:
            continue
        seen.add(_key(path))
        if not path.exists():
            hint = "找不到这个文件"
            if "/" in line.strip().strip('"') and "\\" in line:
                hint += "（路径里有半角「/」，文件名里原本可能是全角「／」）"
            missing.append({"path": str(path), "reason": hint})
        elif path.is_file() and path.name.startswith("~$"):
            missing.append({"path": str(path), "reason": "这是 Excel 打开时留下的临时文件，不复制"})
        else:
            picked.append(path.resolve())
    if not picked:
        raise ValueError("这些路径一个也找不到：" + "；".join(m["path"] for m in missing[:3]))
    try:
        root = Path(os.path.commonpath([str(p.parent) for p in picked]))
    except ValueError:
        raise ValueError("这些路径不在同一个盘上，请分几次复制") from None
    top = root.name  # 共同上级自己的名字也保留；在盘的根目录时没有名字，就直接放进目标
    files: dict[str, dict[str, Any]] = {}
    for path in picked:
        sources = [path] if path.is_file() else [path / rel for rel in scan_folder(path)["excel"]]
        if path.is_dir() and not sources:
            missing.append({"path": str(path), "reason": "这个文件夹里没有 Excel"})
        for src in sources:
            rel = Path(top, src.relative_to(root)) if top else src.relative_to(root)
            target = dest / rel
            files.setdefault(_key(src), {"source": str(src), "rel": rel.as_posix(), "target": str(target),
                                         "exists": target.exists(), "size": src.stat().st_size})
    for item in files.values():
        if os.path.normcase(os.path.abspath(item["target"])) == os.path.normcase(os.path.abspath(item["source"])):
            raise ValueError("目标文件夹正好是这些文件现在所在的地方，复制过去等于盖住自己，请换一个目标文件夹")
    items = sorted(files.values(), key=lambda item: _key(item["rel"]))
    return {"dest": str(dest), "root": str(root), "top": top, "lines": len(lines), "duplicates": len(lines) - len(seen),
            "files": items, "missing": missing, "exists": sum(item["exists"] for item in items)}


def preview(paths_text: Any, dest_text: Any) -> dict[str, Any]:
    """只列出会复制什么，不改任何东西"""
    return _plan(paths_text, dest_text)


def run(paths_text: Any, dest_text: Any, overwrite: bool = False) -> dict[str, Any]:
    plan = _plan(paths_text, dest_text)
    if plan["exists"] and not overwrite:
        raise ValueError(f"目标里已经有 {plan['exists']} 个同名文件，请确认覆盖后再复制")
    with debug_log.side_task("pick_copy", {"dest": plan["dest"], "files": len(plan["files"]), "root": plan["root"]}):
        for item in plan["files"]:
            target = Path(item["target"])
            try:
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(item["source"], target)
                item["status"] = "ok"
            except OSError as err:
                item.update(status="failed", error=f"{err.strerror or err}")
                debug_log.event("❌ 挑文件复制没成功", 文件=item["source"], 原因=item["error"])
        plan["copied"] = sum(item["status"] == "ok" for item in plan["files"])
        plan["failed"] = len(plan["files"]) - plan["copied"]
        debug_log.event("📦 挑文件复制", 目标=plan["dest"], 复制=plan["copied"], 没成功=plan["failed"],
                        找不到=len(plan["missing"]), 重复行=plan["duplicates"])
    return plan


def open_dest(dest_text: Any) -> None:
    dest = _clean(dest_text) if isinstance(dest_text, str) and dest_text.strip() else None
    if dest is None or not dest.is_dir():
        raise ValueError("目标文件夹不存在")
    os.startfile(str(dest))  # type: ignore[attr-defined]
