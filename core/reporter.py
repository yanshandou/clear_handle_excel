import atexit
import copy
import json
import logging
import threading
import time
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

from core import fix_images, operation_history
from core.common import (
    EXCEL_EXTENSIONS, PROJECT_ROOT, RECORD_PATH, has_review_prefix, md_wrap, review_name_variants, strip_review_prefix,
    workbook_rel_name,
)
from core.storage import load_snapshot, save_snapshot

REPORT_LOCK = threading.RLock()
REPORT_STATE_PATH = PROJECT_ROOT / "record_state.json"
WRITE_INTERVAL = 1.0  # 进度更新很频繁：数据库和 Markdown 最多每秒写一次，开始、结束、中断时立即写
_history_context = threading.local()
_last_write = 0.0
_write_timer: threading.Timer | None = None

EXECUTION_STATE: dict[str, Any] = {
    "task_name": "尚未开始任务",
    "folder": "",
    "pdf_folder": "",
    "backup_folder": "",
    "status": "idle",  # idle | running | completed
    "started_at": "",
    "updated_at": "",
    "total_files": 0,
    "completed_files": 0,
    "processing_files": 0,
    "pending_files": 0,
    "error_files": 0,
    "progress_pct": 0.0,
    "current_file": "",
    "current_step": "等待开始...",
    "files": {},  # wb_name -> file record dict
}


def _restore_report() -> None:
    loaded = load_snapshot("report")
    if loaded is None and REPORT_STATE_PATH.is_file():
        loaded = json.loads(REPORT_STATE_PATH.read_text(encoding="utf-8"))
        if not isinstance(loaded, dict) or "files" not in loaded:
            raise ValueError("Invalid legacy execution report")
        save_snapshot("report", loaded)
    if loaded is not None:
        EXECUTION_STATE.update(loaded)
        changed = _merge_review_duplicates_unlocked()
        if _mark_running_as_interrupted_unlocked() or changed:
            save_snapshot("report", EXECUTION_STATE)


def _base_name(name: str) -> str:
    """记录名里的文件名部分，去掉「★」"""
    return strip_review_prefix(Path(name).name)


def _loose_name(name: str) -> str:
    """比较文件名用：去掉「★」、全角半角统一（Ｗ/W）、去空格、不分大小写"""
    return "".join(unicodedata.normalize("NFKC", _base_name(name)).split()).casefold()


def _fold_duplicate(target: dict[str, Any], dup: dict[str, Any]) -> None:
    """旧版本把后面步骤的结果记在了带前缀的那一条上：并进原文件那一条，不丢数据"""
    target.setdefault("fixes", {}).update(dup.get("fixes") or {})
    target["deleted_backups"] = list(dict.fromkeys(list(target.get("deleted_backups", [])) + list(dup.get("deleted_backups", []))))
    if not target.get("all_sheets") and dup.get("all_sheets"):
        target["all_sheets"] = list(dup["all_sheets"])
    target["normalized_a1"] = bool(target.get("normalized_a1") or dup.get("normalized_a1"))
    target["exported_pdfs"] = max(int(target.get("exported_pdfs", 0) or 0), int(dup.get("exported_pdfs", 0) or 0))
    if target.get("status") in {"pending", "processing"} and dup.get("status") not in {None, "pending", "processing"}:
        for field in ("status", "current_action", "error", "updated_at", "clip_detected", "clip_fixed",
                      "clip_accepted", "clip_unfixed"):
            if field in dup:
                target[field] = dup[field]


def _merge_review_duplicates_unlocked() -> bool:
    """旧版本按「🟡」改名后的文件名另记了一条：并回原文件那一条（数字之后由工作台重新同步）"""
    files = EXECUTION_STATE["files"]
    changed = False
    for key in [k for k in files if has_review_prefix(Path(k).name)]:
        dup = files.pop(key)
        changed = True
        targets = [k for k in files if _base_name(k) == _base_name(key)]
        if len(targets) != 1:
            targets = [k for k in files if _loose_name(k) == _loose_name(key)]
        if len(targets) == 1:
            _fold_duplicate(files[targets[0]], dup)
        elif dup.get("all_sheets") or dup.get("fixes") or dup.get("deleted_backups"):
            plain = str(Path(key).with_name(_base_name(key))).replace("\\", "/")
            if plain in files:
                _fold_duplicate(files[plain], dup)
            else:
                dup["name"] = plain
                files[plain] = dup
        # 其余（没有工作表、没有修复记录的空壳）直接去掉，切字数字由工作台重新同步
    batch = EXECUTION_STATE.get("batch")
    if changed and batch:
        EXECUTION_STATE["batch"] = [k for k in dict.fromkeys(batch) if k in files]
    if changed:
        _recompute_counters_unlocked()
    return changed


def _mark_running_as_interrupted_unlocked() -> bool:
    """恢复的记录若停在执行中，说明任务已中断；返回是否有修改。"""
    if EXECUTION_STATE["status"] != "running":
        return False
    EXECUTION_STATE["status"] = "interrupted"
    EXECUTION_STATE["current_step"] = "任务已中断，可在主页或工作台继续上次任务"
    EXECUTION_STATE["current_file"] = ""
    for entry in EXECUTION_STATE["files"].values():
        if entry["status"] == "processing":
            entry["status"] = "pending"
            entry["current_action"] = "任务中断，等待继续"
    _recompute_counters_unlocked()
    return True


def replace_execution_report(state: dict[str, Any]) -> None:
    """导入进度时整体替换执行记录，并同步写入数据库与 Markdown 报告。"""
    if (
        not isinstance(state, dict)
        or not isinstance(state.get("files"), dict)
        or not isinstance(state.get("status"), str)
    ):
        raise ValueError("进度文件损坏：缺少执行记录")
    with REPORT_LOCK:
        previous = copy.deepcopy(EXECUTION_STATE)
        EXECUTION_STATE.update(copy.deepcopy(state))
        EXECUTION_STATE.pop("markdown", None)
        _merge_review_duplicates_unlocked()
        _mark_running_as_interrupted_unlocked()
        try:
            _persist_and_write_markdown_unlocked()
        except Exception:
            EXECUTION_STATE.clear()
            EXECUTION_STATE.update(previous)
            raise


def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _time_str() -> str:
    return datetime.now().strftime("%H:%M:%S")


def _recompute_counters_unlocked() -> None:
    batch = EXECUTION_STATE.get("batch") or []
    files = (
        [EXECUTION_STATE["files"][name] for name in batch if name in EXECUTION_STATE["files"]]
        if batch else list(EXECUTION_STATE["files"].values())
    )
    total = len(files)
    done_cnt = sum(1 for f in files if f["status"] in {"done", "needs_fix"})
    proc_cnt = sum(1 for f in files if f["status"] == "processing")
    pend_cnt = sum(1 for f in files if f["status"] == "pending")
    err_cnt = sum(1 for f in files if f["status"] == "error")

    finished_cnt = done_cnt + err_cnt
    pct = round((finished_cnt / total) * 100.0, 1) if total > 0 else 0.0

    EXECUTION_STATE["total_files"] = total
    EXECUTION_STATE["completed_files"] = done_cnt
    EXECUTION_STATE["processing_files"] = proc_cnt
    EXECUTION_STATE["pending_files"] = pend_cnt
    EXECUTION_STATE["error_files"] = err_cnt
    EXECUTION_STATE["progress_pct"] = pct
    EXECUTION_STATE["updated_at"] = _now_str()


def report_folder() -> str:
    with REPORT_LOCK:
        return str(EXECUTION_STATE.get("folder", ""))


def same_folder(a: str, b: Path | str) -> bool:
    return _same_folder(a, Path(b)) if b else False


def _same_folder(a: str, b: Path) -> bool:
    if not a:
        return False
    try:
        return Path(a).resolve() == b.resolve()
    except OSError:
        return Path(a) == b


_SUSPENDED = threading.local()


def suspend_updates(suspended: bool) -> None:
    """当前线程暂不写执行记录：工作台审核任务运行时，执行记录属于别的文件夹（主页对别的文件夹跑过任务）"""
    _SUSPENDED.on = suspended


def _updates_suspended() -> bool:
    return getattr(_SUSPENDED, "on", False)


def init_execution_batch(
    task_name: str,
    folder: Path,
    workbooks: list[Path],
    pdf_folder: Path | None = None,
    backup_folder: Path | None = None,
    preserve_existing_files: bool = False,
) -> None:
    """初始化一批待处理 Excel 文件的执行记录队列。

    同一个文件夹接着做下一步（精简 Sheet → 导出 PDF → 切字扫描……）时累加：
    每个文件的记录合并，前面步骤的结果不丢；换了文件夹才清空重来。"""
    kinds = ({"精简 Sheet（删除备份表）": ["clean"], "批量导出 PDF": ["pdf"],
              "视图标准化（分页预览 + A1 + 100% 缩放）": ["view"],
              "只扫描（全部转人工）": ["scan"], "全自动扫描与修复": ["scan", "autofix"],
              "自动调整选中的格子": ["autofix"], "Web 切字扫描（完整）": ["scan"],
              "Web 切字扫描（增量）": ["scan"]}).get(task_name, [])
    with REPORT_LOCK:
        _history_context.ids = operation_history.begin(folder, kinds) if kinds else []
        EXECUTION_STATE["operation_history_ids"] = _history_context.ids
        now = _now_str()
        old_folder = EXECUTION_STATE.get("folder", "")
        # 换了文件夹一定清空：否则上一个文件夹的文件会出现在这个文件夹的报告里，也打不开
        keep = _same_folder(old_folder, folder) or (preserve_existing_files and not old_folder)
        EXECUTION_STATE["task_name"] = task_name
        EXECUTION_STATE["folder"] = str(folder)
        # 换文件夹不沿用上一批的输出目录，避免报告里出现与本次任务无关的路径
        if pdf_folder is not None or not keep:
            EXECUTION_STATE["pdf_folder"] = str(pdf_folder) if pdf_folder is not None else ""
        if backup_folder is not None or not preserve_existing_files:
            EXECUTION_STATE["backup_folder"] = str(backup_folder) if backup_folder is not None else ""
        EXECUTION_STATE["status"] = "running"
        if not keep or not EXECUTION_STATE["started_at"]:
            EXECUTION_STATE["started_at"] = now
        EXECUTION_STATE["updated_at"] = now
        EXECUTION_STATE["current_file"] = ""
        EXECUTION_STATE["current_step"] = f"准备执行：{task_name}"

        if not keep:
            EXECUTION_STATE["files"].clear()
            EXECUTION_STATE["steps"] = []
            if old_folder:
                fix_images.clear()  # 上一个文件夹的修复截图用不到了
        steps = EXECUTION_STATE.setdefault("steps", [])
        if task_name not in steps:
            steps.append(task_name)

        # 进度只按这一步要处理的文件算；报告里仍列出这个文件夹做过的全部文件
        EXECUTION_STATE["batch"] = _queue_files_unlocked(folder, workbooks)

        _recompute_counters_unlocked()
        _persist_and_write_markdown_unlocked()


def add_to_execution_batch(folder: Path, workbooks: list[Path]) -> None:
    """正在跑的任务又排进来几份文件（「🤖 自动调整」排队）：加进本批次，进度和报告跟着算"""
    if _updates_suspended():
        return
    with REPORT_LOCK:
        batch = EXECUTION_STATE.setdefault("batch", [])
        batch.extend(name for name in _queue_files_unlocked(folder, workbooks) if name not in batch)
        EXECUTION_STATE["status"] = "running"
        _recompute_counters_unlocked()
        _persist_and_write_markdown_unlocked()


def _queue_files_unlocked(folder: Path, workbooks: list[Path]) -> list[str]:
    """把这些文件记成排队等待处理（已有记录的沿用），返回它们在报告里的名字。调用方持有 REPORT_LOCK"""
    next_index = max((int(f.get("index", 0)) for f in EXECUTION_STATE["files"].values()), default=0) + 1
    batch: list[str] = []
    for wb_path in workbooks:
        wb_name = workbook_rel_name(wb_path, folder)
        batch.append(wb_name)
        existing = EXECUTION_STATE["files"].get(wb_name)
        if existing is not None:
            existing["status"] = "pending"
            existing["path"] = str(wb_path)
            existing["current_action"] = "排队等待处理..."
            existing["error"] = ""
            existing["updated_at"] = _time_str()
        else:
            EXECUTION_STATE["files"][wb_name] = {
                "index": next_index,
                "name": wb_name,
                "path": str(wb_path),
                "status": "pending",  # pending | processing | needs_fix | done | error
                "current_action": "排队等待处理...",
                "all_sheets": [],
                "deleted_backups": [],
                "uncertain_sheets": [],
                "renamed_to": "",
                "exported_pdfs": 0,
                "clip_detected": 0,
                "clip_fixed": 0,
                "clip_unfixed": 0,
                "normalized_a1": False,
                "updated_at": _time_str(),
                "error": "",
            }
            next_index += 1
    return batch


def update_file_progress(
    wb_name: str,
    *,
    status: str | None = None,
    current_action: str | None = None,
    all_sheets: list[str] | None = None,
    deleted_backups: list[str] | None = None,
    sheet_change: dict | None = None,
    uncertain_sheets: list[str] | None = None,
    renamed_to: str | None = None,
    exported_pdfs: int | None = None,
    clip_detected: int | None = None,
    clip_fixed: int | None = None,
    clip_accepted: int | None = None,
    clip_unfixed: int | None = None,
    normalized_a1: bool | None = None,
    fixes: dict[str, dict[str, Any] | None] | None = None,
    error: str | None = None,
) -> None:
    """实时更新单个 Excel 文件的处理进度与统计指标

    fixes：{问题 id: 修复记录}，值为 None 表示这处现在不算已修复；没提到的问题沿用以前的记录
    （重新扫描后旧问题项可能不在工作台了，但文件里确实修过）。"""
    if _updates_suspended():
        return
    with REPORT_LOCK:
        entry = EXECUTION_STATE["files"].get(wb_name)
        if entry is None and EXECUTION_STATE.get("folder"):
            rel_key = workbook_rel_name(Path(wb_name), EXECUTION_STATE["folder"])
            entry = EXECUTION_STATE["files"].get(rel_key)
            if entry is not None:
                wb_name = rel_key
        if entry is None:
            # 只给了文件名（可能还带着「★」）：按去掉「★」后的文件名找同一个文件
            base = _base_name(wb_name)
            matches = [
                (k, v)
                for k, v in EXECUTION_STATE["files"].items()
                if base in {_base_name(k), _base_name(v.get("path", k))}
            ]
            if len(matches) == 1:
                wb_name, entry = matches[0]
            else:
                wb_name = str(Path(wb_name).with_name(base)).replace("\\", "/")
                entry = EXECUTION_STATE["files"].get(wb_name)
        if entry is None:
            entry = {
                "index": max((int(f.get("index", 0)) for f in EXECUTION_STATE["files"].values()), default=0) + 1,
                "name": wb_name,
                "path": wb_name,
                "status": "pending",
                "current_action": "",
                "all_sheets": [],
                "deleted_backups": [],
                "uncertain_sheets": [],
                "renamed_to": "",
                "exported_pdfs": 0,
                "clip_detected": 0,
                "clip_fixed": 0,
                "clip_unfixed": 0,
                "normalized_a1": False,
                "updated_at": _time_str(),
                "error": "",
            }
            EXECUTION_STATE["files"][wb_name] = entry
            if EXECUTION_STATE.get("batch"):
                EXECUTION_STATE["batch"].append(wb_name)

        if status is not None:
            entry["status"] = status
        if current_action is not None:
            entry["current_action"] = current_action
            EXECUTION_STATE["current_step"] = f"[{wb_name}] {current_action}"
        if status == "processing":
            EXECUTION_STATE["current_file"] = wb_name
        if deleted_backups is not None:
            # 删掉的表只会越来越多：后面的步骤（切字扫描等）不知道前面删过什么，不能把记录清掉
            entry["deleted_backups"] = list(dict.fromkeys(list(entry.get("deleted_backups", [])) + list(deleted_backups)))
        if sheet_change is not None:
            entry.setdefault("sheet_changes", []).append(copy.deepcopy(sheet_change))
        if all_sheets is not None:
            entry["all_sheets"] = _merge_sheet_list(entry.get("all_sheets", []), list(all_sheets), entry.get("deleted_backups", []))
        if fixes is not None:
            known = entry.setdefault("fixes", {})
            for item_id, record in fixes.items():
                if record is None:
                    known.pop(item_id, None)
                else:
                    known[item_id] = dict(record)
        if uncertain_sheets is not None:
            entry["uncertain_sheets"] = list(uncertain_sheets)
        if renamed_to is not None:
            entry["renamed_to"] = str(renamed_to)
        if exported_pdfs is not None:
            entry["exported_pdfs"] = int(exported_pdfs)
        if clip_detected is not None:
            entry["clip_detected"] = int(clip_detected)
        if clip_fixed is not None:
            entry["clip_fixed"] = int(clip_fixed)
        if clip_accepted is not None:
            entry["clip_accepted"] = int(clip_accepted)
        if clip_unfixed is not None:
            entry["clip_unfixed"] = int(clip_unfixed)
        if normalized_a1 is not None:
            entry["normalized_a1"] = bool(normalized_a1)
        if error is not None:
            entry["error"] = str(error)
        if entry["status"] == "done" and int(entry.get("clip_unfixed", 0) or 0) > 0:
            # 导出 PDF / 视图标准化等步骤做完了，但还有切字没修，不能算完成
            entry["status"] = "needs_fix"

        entry["updated_at"] = _time_str()
        _recompute_counters_unlocked()
        _write_soon_unlocked()


def _merge_sheet_list(old: list[str], new: list[str], deleted: list[str]) -> list[str]:
    """新读到的工作表列表里已经没有删掉的表了：把以前删掉的表按原来的位置插回去"""
    result = list(new)
    gone = set(deleted) - set(new)
    for pos, name in enumerate(old):
        if name not in gone or name in result:
            continue
        before = next((old[i] for i in range(pos - 1, -1, -1) if old[i] in result), None)
        result.insert(result.index(before) + 1 if before is not None else 0, name)
    return result


def resume_operation_history(folder: str, task_names: set[str]) -> None:
    """继续同一批任务时更新原记录，不重复计数。"""
    with REPORT_LOCK:
        if (EXECUTION_STATE.get("status") != "interrupted"
                or EXECUTION_STATE.get("task_name") not in task_names
                or not _same_folder(EXECUTION_STATE.get("folder", ""), folder)):
            return
        ids = EXECUTION_STATE.get("operation_history_ids", [])
        _history_context.ids = ids
        operation_history.finish(ids, "running")


def finish_execution_batch(summary_step: str = "全部任务处理完毕") -> None:
    """标记当前批次任务完成，并持久化最新 Markdown 与 JSON 状态"""
    if _updates_suspended():
        return
    with REPORT_LOCK:
        operation_history.finish(getattr(_history_context, "ids", []), "failed" if EXECUTION_STATE.get("error_files") or "未完成" in summary_step else "completed")
        EXECUTION_STATE["status"] = "completed"
        EXECUTION_STATE["current_file"] = ""
        EXECUTION_STATE["current_step"] = summary_step
        _recompute_counters_unlocked()
        _persist_and_write_markdown_unlocked()


def interrupt_execution_batch(message: str) -> None:
    if _updates_suspended():
        return
    with REPORT_LOCK:
        operation_history.finish(getattr(_history_context, "ids", []), "interrupted")
        EXECUTION_STATE["status"] = "interrupted"
        EXECUTION_STATE["current_file"] = ""
        EXECUTION_STATE["current_step"] = message
        for entry in EXECUTION_STATE["files"].values():
            if entry["status"] == "processing":
                entry["status"] = "pending"
                entry["current_action"] = "任务中断，等待继续"
        _recompute_counters_unlocked()
        _persist_and_write_markdown_unlocked()


def _fix_sort_key(record: dict[str, Any], sheets: list[str]) -> tuple[Any, ...]:
    sheet = record.get("sheet", "")
    return (sheets.index(sheet) if sheet in sheets else len(sheets), sheet,
            int(record.get("row") or 0), str(record.get("cell", "")))


def _md_cell(text: Any) -> str:
    return str(text).replace("|", "\\|").replace("\n", " ").strip() or "-"


def _unfinished(f: dict[str, Any]) -> bool:
    return f.get("status") in {"pending", "processing", "error"} or int(f.get("clip_unfixed", 0) or 0) > 0


def _report_state() -> dict[str, Any]:
    with REPORT_LOCK:
        return copy.deepcopy(EXECUTION_STATE)


def _fix_place(fix: dict[str, Any]) -> str:
    place = fix.get("cell") or ""
    if fix.get("row"):
        place = f"{place}（第 {fix['row']} 行）" if place else f"第 {fix['row']} 行"
    return place


def _fix_shots(fixes: list[dict[str, Any]]) -> list[tuple[list[int], str, str]]:
    """有截图的修复：[(用这组截图的序号们, 修复前文件名, 修复后文件名)]；同一行共用的截图只放一次"""
    shots: dict[tuple[str, str], list[int]] = {}
    for i, fix in enumerate(fixes, start=1):
        before = fix.get("before") if fix_images.image_path(fix.get("before") or "") else ""
        after = fix.get("after") if fix_images.image_path(fix.get("after") or "") else ""
        if before or after:
            shots.setdefault((before, after), []).append(i)
    return [(numbers, before, after) for (before, after), numbers in shots.items()]


def _status_text(state: dict[str, Any], files: list[dict[str, Any]]) -> str:
    unfinished = [f for f in files if _unfinished(f)]
    if state["status"] == "running":
        return "🔄 正在处理"
    if state["status"] == "interrupted":
        return "⏸ 已中断，还没处理完"
    if unfinished:
        return f"⚠️ 还有 {len(unfinished)} 个文件没处理完"
    return "✅ 已全部处理完成"


def render_execution_markdown(image_dir: str | None = None) -> str:
    """根据当前所有文件的处理结果生成 Markdown 报告（要交给领导）

    每个文件一节：① 删除了哪些备份表（删除线），② 修复了几处、在哪、原因。
    只写结果：不写日期时间，不写人工确认 / 自动处理 / 跳过 / 复检等过程，判断为不用改的地方不列出。
    image_dir：下载压缩包时给图片文件夹名，每处修复下面附修复前 / 修复后截图。
    """
    state = _report_state()

    files = sorted(state["files"].values(), key=lambda f: int(f.get("index", 0)))
    total_deleted = sum(len(f.get("deleted_backups", [])) for f in files)
    total_fixed = sum(len(f.get("fixes") or {}) for f in files)

    status_text = _status_text(state, files)

    lines: list[str] = [
        "# 📊 Excel 整理与切字修复记录",
        "",
        f"- **目标文件夹**：`{state['folder'] or '-'}`",
        f"- **处理状态**：{status_text}",
        f"- **汇总**：共 {len(files)} 个文件 ｜ 删除备份 Sheet **{total_deleted}** 个 ｜ 修复切字 **{total_fixed}** 处",
        "",
        "---",
        "",
    ]

    for number, f in enumerate(files, start=1):
        title = f["name"]
        if f.get("renamed_to") and f["renamed_to"] != f["name"]:
            title += f" → {f['renamed_to']}"
        title += "（还没处理完）" if _unfinished(f) and state["status"] != "running" else ""
        lines += [f"## {number}. {title}", ""]

        all_sheets = list(f.get("all_sheets", []))
        deleted = list(f.get("deleted_backups", []))
        deleted_set = set(deleted)
        lines.append(f"**① 删除备份 Sheet：{len(deleted)} 个**" if deleted else "**① 删除备份 Sheet：无**")
        lines.append("")
        for sh in all_sheets:
            lines.append(f"- {md_wrap(sh, '~~')}" if sh in deleted_set else f"- {sh}")
        for sh in deleted:
            if sh not in all_sheets:
                lines.append(f"- {md_wrap(sh, '~~')}")
        if all_sheets or deleted:
            lines.append("")

        for change in f.get("sheet_changes", []):
            for old, new in change["renamed"].items():
                lines.append(f"- Sheet 改名：{old} → {new}")
        if f.get("sheet_changes"):
            lines.append("")
        fixes = sorted((f.get("fixes") or {}).values(), key=lambda r: _fix_sort_key(r, all_sheets))
        if fixes:
            lines += [f"**② 修复切字：{len(fixes)} 处**", "", "| # | Sheet | 位置 | 原因 |", "|---|---|---|---|"]
            for i, fix in enumerate(fixes, start=1):
                lines.append(
                    f"| {i} | {_md_cell(fix.get('sheet', ''))} | {_md_cell(_fix_place(fix))} | {_md_cell(fix.get('reason', ''))} |"
                )
            if image_dir is not None:
                for numbers, before, after in _fix_shots(fixes):
                    first = fixes[numbers[0] - 1]
                    label = "、".join(f"#{n}" for n in numbers)
                    lines += ["", f"**{label}　{_md_cell(first.get('sheet', ''))} {_md_cell(_fix_place(first))}**", "",
                              "| 修复前 | 修复后 |", "|---|---|",
                              f"| {_md_image(image_dir, before, '修复前')} | {_md_image(image_dir, after, '修复后')} |"]
        else:
            lines.append("**② 修复切字：无**")
        lines.append("")
        if f.get("status") == "error" and f.get("error"):
            lines += [f"> ❌ 处理出错：{str(f['error']).replace(chr(10), ' ')}", ""]

    return "\n".join(lines)


def _md_image(image_dir: str, name: str, alt: str) -> str:
    return f"![{alt}]({image_dir}/{name})" if name else "-"


def _write_soon_unlocked() -> None:
    global _write_timer
    wait = _last_write + WRITE_INTERVAL - time.monotonic()
    if wait <= 0:
        _persist_and_write_markdown_unlocked()
    elif _write_timer is None:
        _write_timer = threading.Timer(wait, _flush_pending_write)
        _write_timer.daemon = True
        _write_timer.start()


def _flush_pending_write() -> None:
    global _write_timer
    with REPORT_LOCK:
        if _write_timer is None:
            return  # 已经立即写过了
        _write_timer = None
        try:
            _persist_and_write_markdown_unlocked()
        except Exception:
            logging.exception("无法保存执行记录")


def flush_execution_report() -> None:
    """把还没写的进度立即写进数据库和 Markdown（程序退出时也会调用）"""
    with REPORT_LOCK:
        if _write_timer is not None:
            _persist_and_write_markdown_unlocked()


atexit.register(flush_execution_report)


def _persist_and_write_markdown_unlocked() -> None:
    global _last_write, _write_timer
    if _write_timer is not None:
        _write_timer.cancel()
        _write_timer = None
    _last_write = time.monotonic()
    save_snapshot("report", EXECUTION_STATE)

    md_text = render_execution_markdown()
    try:
        RECORD_PATH.write_text(md_text, encoding="utf-8", newline="\n")
    except OSError:
        logging.exception("无法写入执行记录 Markdown: %s", RECORD_PATH)

    folder_str = EXECUTION_STATE.get("folder", "")
    if folder_str:
        folder_path = Path(folder_str)
        if folder_path.is_dir():
            try:
                (folder_path / "_执行记录报告.md").write_text(
                    md_text, encoding="utf-8", newline="\n"
                )
            except OSError:
                logging.exception("无法写入目标目录的执行记录 Markdown: %s", folder_path)


def get_execution_report_snapshot() -> dict[str, Any]:
    """获取已从数据库恢复的执行记录快照。"""
    with REPORT_LOCK:
        snapshot = copy.deepcopy(EXECUTION_STATE)
        snapshot["markdown"] = render_execution_markdown()
        return snapshot


def find_workbook(file_name: str, roots: list[str]) -> Path:
    """网页上显示的文件名（相对目标文件夹，或工作台里的完整路径）对应到磁盘上现在的文件。

    文件名前加了 / 去了「★」（旧版本是「🟡」）也能找到；只认这些文件夹内的 Excel 文件。
    """
    folders = [Path(root).resolve() for root in dict.fromkeys(filter(None, roots))]
    if not folders:
        raise ValueError("还没有选择目标文件夹")
    raw = Path(file_name)
    bases = [raw] if raw.is_absolute() else [folder / raw for folder in folders]
    for base in bases:
        names = dict.fromkeys([base.name, strip_review_prefix(base.name), *review_name_variants(base.name)])
        for name in names:
            wb_path = base.with_name(name).resolve()
            if (
                any(wb_path.is_relative_to(folder) for folder in folders)
                and wb_path.suffix.lower() in EXCEL_EXTENSIONS
                and not wb_path.name.startswith("~$")
                and wb_path.is_file()
            ):
                return wb_path
    raise ValueError(f"找不到文件「{raw.name}」，可能已被移动或改名")


def find_recorded_workbook(file_name: str) -> Path:
    """看板里的文件现在在磁盘上的路径；只接受本次目标文件夹内的 Excel 文件"""
    with REPORT_LOCK:
        folder_str = EXECUTION_STATE.get("folder", "")
        entry = EXECUTION_STATE["files"].get(file_name)
    if not folder_str or entry is None:
        raise ValueError("看板里没有这个文件，请刷新页面")
    for name in filter(None, [entry.get("renamed_to"), file_name]):
        try:
            return find_workbook(name, [folder_str])
        except ValueError:
            continue
    raise ValueError(f"找不到文件「{file_name}」，可能已被移动或改名")


_restore_report()
