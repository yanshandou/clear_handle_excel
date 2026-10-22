import atexit
import copy
import json
import logging
import os
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from core.common import EXCEL_EXTENSIONS, PROJECT_ROOT, RECORD_PATH, review_marked_name, workbook_rel_name
from core.storage import load_snapshot, save_snapshot

REPORT_LOCK = threading.RLock()
REPORT_STATE_PATH = PROJECT_ROOT / "record_state.json"
WRITE_INTERVAL = 1.0  # 进度更新很频繁：数据库和 Markdown 最多每秒写一次，开始、结束、中断时立即写
_last_write = 0.0
_write_timer: threading.Timer | None = None

STATUS_LABELS = {
    "pending": "⏳ 待处理",
    "processing": "🔄 正在处理",
    "needs_fix": "🔧 已扫描（待人工修改）",
    "done": "✅ 已完成",
    "error": "❌ 处理失败",
}

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
        if _mark_running_as_interrupted_unlocked():
            save_snapshot("report", EXECUTION_STATE)


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


def _same_folder(a: str, b: Path) -> bool:
    if not a:
        return False
    try:
        return Path(a).resolve() == b.resolve()
    except OSError:
        return Path(a) == b


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
    with REPORT_LOCK:
        now = _now_str()
        keep = preserve_existing_files or _same_folder(EXECUTION_STATE.get("folder", ""), folder)
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
        steps = EXECUTION_STATE.setdefault("steps", [])
        if task_name not in steps:
            steps.append(task_name)

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
        # 进度只按这一步要处理的文件算；报告里仍列出这个文件夹做过的全部文件
        EXECUTION_STATE["batch"] = batch

        _recompute_counters_unlocked()
        _persist_and_write_markdown_unlocked()


def update_file_progress(
    wb_name: str,
    *,
    status: str | None = None,
    current_action: str | None = None,
    all_sheets: list[str] | None = None,
    deleted_backups: list[str] | None = None,
    uncertain_sheets: list[str] | None = None,
    renamed_to: str | None = None,
    exported_pdfs: int | None = None,
    clip_detected: int | None = None,
    clip_fixed: int | None = None,
    clip_accepted: int | None = None,
    clip_unfixed: int | None = None,
    normalized_a1: bool | None = None,
    error: str | None = None,
) -> None:
    """实时更新单个 Excel 文件的处理进度与统计指标"""
    with REPORT_LOCK:
        entry = EXECUTION_STATE["files"].get(wb_name)
        if entry is None and EXECUTION_STATE.get("folder"):
            rel_key = workbook_rel_name(Path(wb_name), EXECUTION_STATE["folder"])
            entry = EXECUTION_STATE["files"].get(rel_key)
            if entry is not None:
                wb_name = rel_key
        if entry is None:
            matches = [
                (k, v)
                for k, v in EXECUTION_STATE["files"].items()
                if Path(v.get("path", k)).name == Path(wb_name).name
            ]
            if len(matches) == 1:
                wb_name, entry = matches[0]
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
        if all_sheets is not None:
            entry["all_sheets"] = list(all_sheets)
        if deleted_backups is not None:
            entry["deleted_backups"] = list(deleted_backups)
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

        entry["updated_at"] = _time_str()
        _recompute_counters_unlocked()
        _write_soon_unlocked()


def finish_execution_batch(summary_step: str = "全部任务处理完毕") -> None:
    """标记当前批次任务完成，并持久化最新 Markdown 与 JSON 状态"""
    with REPORT_LOCK:
        EXECUTION_STATE["status"] = "completed"
        EXECUTION_STATE["current_file"] = ""
        EXECUTION_STATE["current_step"] = summary_step
        _recompute_counters_unlocked()
        _persist_and_write_markdown_unlocked()


def interrupt_execution_batch(message: str) -> None:
    with REPORT_LOCK:
        EXECUTION_STATE["status"] = "interrupted"
        EXECUTION_STATE["current_file"] = ""
        EXECUTION_STATE["current_step"] = message
        for entry in EXECUTION_STATE["files"].values():
            if entry["status"] == "processing":
                entry["status"] = "pending"
                entry["current_action"] = "任务中断，等待继续"
        _recompute_counters_unlocked()
        _persist_and_write_markdown_unlocked()


def render_execution_markdown() -> str:
    """根据当前所有文件的处理状态生成结构化 Markdown 报告

    报告会提交给别人，不写任何日期和时间（开始时间、最后更新、各文件处理时间只在网页看板显示）。
    """
    with REPORT_LOCK:
        state = copy.deepcopy(EXECUTION_STATE)

    files = list(state["files"].values())
    total_clips = sum(int(f.get("clip_detected", 0)) for f in files)
    fixed_clips = sum(int(f.get("clip_fixed", 0)) for f in files)
    accepted_clips = sum(int(f.get("clip_accepted", 0)) for f in files)
    total_pdfs = sum(int(f.get("exported_pdfs", 0)) for f in files)
    total_bkp_del = sum(len(f.get("deleted_backups", [])) for f in files)
    total_uncertain = sum(len(f.get("uncertain_sheets", [])) for f in files)

    status_badge = {
        "idle": "⚪ 空闲",
        "running": "🔄 正在执行中",
        "completed": "✅ 已全部完成",
        "interrupted": "⏸ 已中断，等待继续",
    }.get(state["status"], state["status"])

    summary_parts = [
        f"共清理备份 Sheet **{total_bkp_del}** 个",
        *(
            [f"待定存疑 Sheet **{total_uncertain}** 个"]
            if total_uncertain
            else []
        ),
        f"导出 PDF **{total_pdfs}** 份",
        f"检出切字 **{total_clips}** 处（已修复 **{fixed_clips}** 处，人工确认无需修复 **{accepted_clips}** 处）",
    ]

    lines: list[str] = [
        "# 📊 Excel 批处理与切字检查执行记录报告",
        "",
        f"- **当前任务**：{state['task_name']}",
        *([f"- **已执行步骤**：{' → '.join(state['steps'])}"] if len(state.get("steps") or []) > 1 else []),
        f"- **任务状态**：{status_badge}（当前进度：**{state['progress_pct']}%**，`{state['completed_files']}/{state['total_files']}` 个文件）",
        f"- **目标文件夹**：`{state['folder'] or '-'}`",
    ]
    if state.get("pdf_folder"):
        lines.append(f"- **PDF 导出目录**：`{state['pdf_folder']}`")
    if state.get("backup_folder"):
        lines.append(f"- **原文件备份目录**：`{state['backup_folder']}`")
    lines.extend([
        f"- **统计汇总**：{' ｜ '.join(summary_parts)}",
        "",
        "---",
        "",
        "## 各工作簿处理与 Sheet 明细",
        "",
    ])

    for f in files:
        st_label = STATUS_LABELS.get(f["status"], f["status"])
        title_name = f["name"]
        if f.get("renamed_to") and f["renamed_to"] != f["name"]:
            title_name += f" → {f['renamed_to']}"
        lines.append(f"### {f['index']}. {title_name} （{st_label}）")
        lines.append("")

        all_sheets = f.get("all_sheets", [])
        deleted_set = set(f.get("deleted_backups", []))
        uncertain_set = set(f.get("uncertain_sheets", []))
        sheet_cnt = len(all_sheets)
        bkp_cnt = len(deleted_set)
        unc_cnt = len(uncertain_set)
        pdf_cnt = int(f.get("exported_pdfs", 0))
        det_cnt = int(f.get("clip_detected", 0))
        fix_cnt = int(f.get("clip_fixed", 0))
        note = str(f.get("error") or f.get("current_action") or "-").replace("\n", " ")

        sheet_meta = f"共 {sheet_cnt} 个 Sheet"
        sheet_sub: list[str] = []
        if bkp_cnt:
            sheet_sub.append(f"已删 {bkp_cnt}")
        if unc_cnt:
            sheet_sub.append(f"待定 {unc_cnt}")
        if sheet_sub:
            sheet_meta += f"（{' · '.join(sheet_sub)}）"

        meta_parts: list[str] = [note, sheet_meta]
        if pdf_cnt > 0:
            meta_parts.append(f"导出 PDF {pdf_cnt} 份")
        if det_cnt > 0:
            clip_str = f"切字 {fix_cnt}/{det_cnt}"
            if f.get("clip_accepted", 0):
                clip_str += f"（人工确认 {f['clip_accepted']}）"
            meta_parts.append(clip_str)
        if f.get("normalized_a1"):
            meta_parts.append("A1+100% 已标准化")

        lines.append(f"> {' ｜ '.join(meta_parts)}")
        lines.append("")

        if all_sheets:
            for sh in all_sheets:
                if sh in deleted_set:
                    lines.append(f"- ~~{sh}~~ *(已作为备份表清理)*")
                elif sh in uncertain_set:
                    lines.append(f"- **{sh}** *(待定存疑，本次未删除)*")
                else:
                    lines.append(f"- {sh}")
        else:
            lines.append("- *(暂未读取工作表列表)*")
        lines.append("")

    return "\n".join(lines)


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


def open_recorded_workbook(file_name: str) -> Path:
    """用默认程序（Excel）打开看板里的文件；只接受本次目标文件夹内的 Excel 文件"""
    with REPORT_LOCK:
        folder_str = EXECUTION_STATE.get("folder", "")
        entry = EXECUTION_STATE["files"].get(file_name)
    if not folder_str or entry is None:
        raise ValueError("看板里没有这个文件，请刷新页面")
    folder = Path(folder_str).resolve()
    candidates = [entry.get("renamed_to"), file_name]
    # 等待稍后人工复检的文件名前临时加了「🟡」
    candidates += [str(Path(name).with_name(review_marked_name(Path(name).name))) for name in filter(None, candidates)]
    for name in filter(None, candidates):
        wb_path = (folder / name).resolve()
        try:
            wb_path.relative_to(folder)
        except ValueError:
            raise ValueError("只能打开目标文件夹内的文件") from None
        if (
            wb_path.suffix.lower() in EXCEL_EXTENSIONS
            and not wb_path.name.startswith("~$")
            and wb_path.is_file()
        ):
            os.startfile(wb_path)  # type: ignore[attr-defined]  # 仅 Windows
            return wb_path
    raise ValueError(f"找不到文件「{file_name}」，可能已被移动或改名")


_restore_report()
