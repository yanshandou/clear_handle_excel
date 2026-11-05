"""文件夹工作台存档。当前工作台仍即时保存；切换时归档，恢复时不启动扫描。"""
import copy
import hashlib
from pathlib import Path
from typing import Any

from core.storage import load_snapshot, save_snapshot

INDEX = "workflow_folders"
LEGACY = "workflow_before_folder_progress"


def key(folder: str) -> str:
    return str(Path(folder).resolve()).casefold()


def _name(folder: str) -> str:
    return "workflow_folder:" + hashlib.sha256(key(folder).encode()).hexdigest()


def images(state: dict) -> set[str]:
    names = set()
    for collection in (state.get("items", {}), state.get("pages", {}),
                       state.get("batch_review", {}).get("cards", {})):
        for entry in collection.values():
            names.update((entry.get("fix_images") or {}).values())
            names.add((entry.get("context") or {}).get("image_file", ""))
            names.add(entry.get("batch_context_image", ""))
    return names - {""}


def archived_images() -> set[str]:
    index = load_snapshot(INDEX) or {}
    return {name for entry in index.values() for name in entry.get("images", [])}


def summary(state: dict, report: dict) -> dict:
    job = state.get("job") or {}
    cards = state.get("batch_review", {}).get("cards", {}).values()
    return {"folder": state.get("folder", ""), "saved_at": state.get("progress_saved_at", ""),
            "status": job.get("status", report.get("status", "idle")),
            "task_name": report.get("task_name", "尚未开始任务"),
            "files": len(state.get("checked_files", {})),
            "issues": sum(not c.get("dismissed") for c in cards),
            "completed": len(job.get("completed", [])), "total": len(job.get("workbooks", [])),
            "can_resume": job.get("status") == "interrupted"}


def save(state: dict, report: dict, *, legacy: bool = False) -> None:
    folder = state.get("folder", "")
    if not folder and not legacy:
        return
    name = LEGACY if legacy else _name(folder)
    payload = dict(state, folder_report=report)
    save_snapshot(name, payload, blob_field="image", blob_collections=("pages", "items"))
    index = load_snapshot(INDEX) or {}
    references = images(state)
    for file in report.get("files", {}).values():
        for fix in (file.get("fixes") or {}).values():
            references.update(fix.get(part) for part in ("before", "after") if fix.get(part))
    index[name] = {**summary(state, report), "images": sorted(references)}
    save_snapshot(INDEX, index)


def load(folder: str, *, include_images: bool = True) -> dict | None:
    return load_snapshot(_name(folder), include_blobs=include_images)


def scoped_records(old: dict, folder: str) -> dict:
    """按路径拆开旧版混合记录；保留原色、审核规则和文件历史。"""
    result: dict[str, Any] = {"batch_review": dict(old.get("batch_review", {}))}
    root = key(folder).rstrip("/\\")

    def inside(path: str) -> bool:
        if not path:
            return False
        normalized = key(path)
        return normalized == root or normalized.startswith(root + "/") or normalized.startswith(root + "\\")

    for field in ("checked_files", "print_layouts", "sheet_history", "sheet_codes", "review_queue"):
        result[field] = {p: copy.deepcopy(v) for p, v in old.get(field, {}).items() if inside(p)}
    for field in ("items", "pages", "pdf_exports"):
        result[field] = {p: copy.deepcopy(v) for p, v in old.get(field, {}).items() if inside(v.get("wb_path", ""))}
    batch = old.get("batch_review", {})
    for field in ("cards", "rows", "columns", "highlights"):
        result["batch_review"][field] = {p: copy.deepcopy(v) for p, v in batch.get(field, {}).items()
                                          if inside(v.get("wb_path") or v.get("path", ""))}
    for field in ("tab_marks", "sheet_checks", "slow_files", "highlight_files", "open_history"):
        result["batch_review"][field] = {p: copy.deepcopy(v) for p, v in batch.get(field, {}).items() if inside(p)}
    result["batch_review"]["ignored"] = [copy.deepcopy(v) for v in batch.get("ignored", [])
                                           if v.get("cell") and inside(v["cell"][0])]
    cards = result["batch_review"]["cards"]
    followups = []
    for entry in batch.get("followups", []):
        selected = [cid for cid in entry.get("card_ids", []) if cid in cards]
        if selected:
            followups.append({**entry, "card_ids": selected,
                              "full_paths": [p for p in entry.get("full_paths", []) if inside(p)]})
    result["batch_review"]["followups"] = followups
    return result


def legacy_state(folder: str, defaults: dict[str, Any]) -> dict:
    """旧版其他文件夹只恢复仍存在的记录，不推测已经丢失的扫描进度。"""
    old = load_snapshot(LEGACY) or {}
    result = copy.deepcopy(defaults)
    result.update(folder=str(Path(folder).resolve()), pdf_folder=str(Path(folder).resolve()))
    if not old:
        return result
    result.update(scoped_records(old, folder))
    report = old.get("folder_report", {})
    if report.get("folder") and key(report["folder"]) == key(folder):
        result["folder_report"] = copy.deepcopy(report)
    return result
