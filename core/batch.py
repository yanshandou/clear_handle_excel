"""无人值守扫描修复，以及按单元格进行的三色结果审核。

PDF 坐标只用于当前一次导出。每保存一行，重新检测整张表再选下一行。
同一状态不重复导出：试算的原样那张用刚检测过的 PDF；写入后副本做同样修改继续用，
不再重开原文件重新检测；每个文件最后再用原文件完整复查改过的表。
审核颜色只写数据库；原文件的还原由独立、可续跑的审核任务统一执行。
"""
import copy
import hashlib
import math
import re
import shutil
import tempfile
import time
import traceback
import unicodedata
from collections.abc import Callable
from pathlib import Path
from typing import Any
from uuid import uuid4

from core import autofix, debug_log, fix_images, page_breaks, workflow as wf
from core.common import STALE_LOCK_HINT, workbook_locked, workbook_rel_name
from core.excel_ops import (
    CellProbe, close_saved_workbooks, extract_sheet_snapshot, normalize_workbook_view,
    open_workbooks_for_review, snapshot_cells, open_workbook_checked,
)
from core.excel_watchdog import ExcelWatchdog, idle_timeout
from core.hidden_text import (
    REASON_CLIPPED, REASON_TRUNCATED, collect_text_pieces, is_excel_error, locate_text, normalize_text,
)
from core.pdf_detector import fitz, render_crop_base64
from core.reporter import add_to_execution_batch, finish_execution_batch, init_execution_batch, update_file_progress

COLORS = ("blue", "green", "yellow")
KIND_LABELS = {"hidden": "隐藏截断", "clip": "压线", "pagebreak": "分页切开"}
CLIP_REASON = "文字被边框横线切穿，打印出来被切掉一部分（行高不够）"
WAITING_MESSAGE = "等待自动处理"
# Excel 行高按屏幕像素取整（100% 缩放 0.75pt、125% 0.6pt），写入后读回会差一点；
# 比这个差得少就算同一个行高（日文 Excel 默认 18.75pt，按 0.1 取整成 18.8 会被误判成「已被修改」）
ROW_HEIGHT_TOLERANCE = .4
XL_HALIGN_RIGHT, XL_HALIGN_CENTER = -4152, -4108
XL_PATTERN_NONE, XL_PATTERN_SOLID = -4142, 1
# 只扫描模式在原 Excel 里给问题单元格标的颜色（可在页面上改）。
# 默认洋红底白字：业务表格几乎不用这个颜色，一眼就能看到，白字对比度约 6:1；黄色已有别人在用
DEFAULT_HIGHLIGHT = {"fill": "#C0007A", "font": "#FFFFFF"}
SCAN_ONLY_MESSAGE = "只扫描：已在 Excel 中标色，请人工修改并保存后复检"
HANDOVER_MESSAGE = "转人工：自动还没处理到这一格，请人工修改并保存后复检；也可以点「转回自动」交还自动处理"
TOUCHED_MESSAGE = "转人工后这一行的行高被改过，仍有问题：留给人工处理，自动不再改这一行"
PICKED_MESSAGE = "等待自动调整（你选的格子）"


def _data() -> dict[str, Any]:
    return wf.STATE.setdefault("batch_review", {"cards": {}, "rows": {}, "ignored": []})


def fix_error_types(value: Any = None, *, require_non_empty: bool = False) -> list[str]:
    """当前自动修复要处理的错误类型；PDF 未匹配文字固定转人工。"""
    if value is None:
        saved = _data().get("fix_error_types")
        selected = (list(autofix.FIXABLE_ERROR_TYPES) if not isinstance(saved, list)
                    else [t for t in autofix.FIXABLE_ERROR_TYPES if t in saved])
    else:
        if not isinstance(value, list) or any(not isinstance(t, str) or t not in autofix.CELL_ERROR_SPECS for t in value):
            raise ValueError("请选择有效的单元格错误类型")
        selected = [t for t in autofix.FIXABLE_ERROR_TYPES if t in value]
    if require_non_empty and not selected:
        raise ValueError("请至少勾选一种要自动修复的错误类型，或使用「只扫描，全部转人工」")
    return selected


def set_fix_error_types(value: Any, *, require_non_empty: bool = False) -> list[str]:
    with wf.STATE_LOCK:
        selected = fix_error_types(value, require_non_empty=require_non_empty)
        _data()["fix_error_types"] = list(selected)
        job = wf.STATE.get("job") or {}
        if job.get("kind") == "batch_scan" and isinstance(job.get("options"), dict):
            job["options"]["fix_error_types"] = list(selected)
        wf._persist_state()
        return selected


def _active_fix_types(job: dict | None = None) -> set[str]:
    if job is not None:
        raw = (job.get("options") or {}).get("fix_error_types")
        if isinstance(raw, list):
            return set(fix_error_types(raw))
    return set(fix_error_types())


def _marks() -> dict[str, dict]:
    """已在原 Excel 里标色的单元格（按卡片 id），记着原来的底色和字色，人工处理完再还原"""
    return _data().setdefault("highlights", {})


def _tabs() -> dict[str, dict[str, dict]]:
    """标成橙黄色的工作表标签：文件 -> {代码名（没有时用 "name:" + 表名）: 原来的标签颜色}，表里的标记全去掉后还原"""
    return _data().setdefault("tab_marks", {})


def highlight_style(value: Any) -> dict[str, str]:
    """页面传来的 {"fill": "#RRGGBB", "font": "#RRGGBB"}；缺了或格式不对用默认颜色"""
    style = dict(DEFAULT_HIGHLIGHT)
    if isinstance(value, dict):
        for key in style:
            color = value.get(key)
            if isinstance(color, str) and len(color) == 7 and color[0] == "#":
                try:
                    int(color[1:], 16)
                except ValueError:
                    continue
                style[key] = color.upper()
    if style["fill"] == style["font"]:
        raise ValueError("标记的底色和字色不能相同")
    return style


def _excel_color(hex_color: str) -> int:
    """#RRGGBB → Excel 的 Color（BGR 整数）"""
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (1, 3, 5))
    return r + g * 256 + b * 65536


def _sheet_checks() -> dict[str, dict[str, Any]]:
    """文件 -> {工作表: 检测时原文件的版本}；"*" 表示整份文件按那个版本检测过"""
    return _data().setdefault("sheet_checks", {})


def _unchanged_since_check(path: str, current: dict[str, int] | None = None) -> bool:
    """上次整份检测之后原文件没再改过（检测规则也没变）：重新导出检测的结果和上次一样"""
    checks = _sheet_checks().get(path) or {}
    if current is None:
        try:
            current = wf._file_version(Path(path))
        except OSError:
            return False
    return checks.get("*") == current


def rename_sheet(path: str, old: str, new: str) -> None:
    """工作表改名时迁移审核卡片、修改记录和未完成的写入日志。调用方持有 STATE_LOCK。"""
    _upgrade_column_records()
    data = _data()
    moved: dict[str, str] = {}
    moved_rows: dict[str, str] = {}
    for cid, card in list(data["cards"].items()):
        if card["wb_path"] == path and card["sheet_name"] == old:
            card["sheet_name"] = new
            card["id"] = _key(*_cell(card))
            moved[cid] = card["id"]
            data["cards"][card["id"]] = data["cards"].pop(cid)
    for rid, row in list(data["rows"].items()):
        if row["path"] == path and row["sheet"] == old:
            row["sheet"] = new
            row["id"] = (_col_key(path, new, row["col"]) if row.get("kind") == "col" else
                         _key(path, new, int(row["row"]), "break") if row.get("kind") == "break" else
                         _key(path, new, int(row["row"])))
            moved_rows[rid] = row["id"]
            data["rows"][row["id"]] = data["rows"].pop(rid)
    for card in data["cards"].values():
        if "col_ids" in card:
            card["col_ids"] = [moved_rows.get(rid, rid) for rid in card["col_ids"]]
    for record in data["rows"].values():
        if "cards" in record:
            record["cards"] = [moved.get(cid, cid) for cid in record["cards"]]
    for rule in data.get("ignored", []):
        if rule["cell"][0] == path and rule["cell"][1] == old:
            rule["cell"][1] = new
    marks = _marks()
    for cid, mark in list(marks.items()):
        if mark["path"] == path and mark["sheet"] == old:
            mark["sheet"] = new
            marks[moved.get(cid, cid)] = marks.pop(cid)
    for tab in _tabs().get(path, {}).values():
        if tab["sheet"] == old:
            tab["sheet"] = new
    checks = _sheet_checks().get(path)
    if checks and old in checks:
        checks[new] = checks.pop(old)
    hand = _data().get("handover", {}).get(path)
    if hand:
        prefix = _row_key(old, 0)[:-1]
        hand["rows"] = {(_row_key(new, 0)[:-1] + key[len(prefix):] if key.startswith(prefix) else key): height
                        for key, height in hand["rows"].items()}
    job = wf.STATE.get("job") or {}
    for field in ("card_ids", "commit_card_ids"):
        if isinstance(job.get("options", {}).get(field), list):
            job["options"][field] = [moved.get(cid, cid) for cid in job["options"][field]]
    if "attempted_rows" in job:
        job["attempted_rows"] = [moved_rows.get(rid, rid) for rid in job["attempted_rows"]]
    for journal in [job.get("pending_adjustment"), *job.get("deferred_adjustments", {}).values()]:
        if not journal or journal["path"] != path:
            continue
        journal["cards"] = [moved.get(cid, cid) for cid in journal["cards"]]
        for row in journal["rows"]:
            row["id"] = moved_rows.get(row["id"], row["id"])
            if row["sheet"] == old:
                row["sheet"] = new
    refresh = job.get("options", {}).get("refresh_sheets", {})
    if path in refresh:
        refresh[path] = [new if name == old else name for name in refresh[path]]


def _key(*parts: Any) -> str:
    return hashlib.sha256("\0".join(map(str, parts)).encode()).hexdigest()[:24]


def _cell(item: dict) -> tuple[str, str, str]:
    return item["wb_path"], item["sheet_name"], item.get("cell_addr", "")


def _is_break(obj: dict) -> bool:
    """分页切开合并区域的问题项 / 卡片：修法是插手动分页，不是加高"""
    return obj.get("kind") == "pagebreak"


def _rid(obj: dict) -> str:
    """一处可写入、可还原的修改：同一行的行高共用一个；插分页另算一个，和行高互不影响"""
    parts = (obj["wb_path"], obj["sheet_name"], int(obj["excel_row"]))
    return _key(*parts, "break") if _is_break(obj) else _key(*parts)


def _signature(item: dict) -> list[str]:
    return [item.get("kind") or "clip", item.get("text", "")]


def carry_item_policies(items: dict[str, dict]) -> None:
    """忽略必须同时匹配路径、表、单元格和问题内容；改内容后重新报告。"""
    data = wf.STATE.get("batch_review")
    if not data:
        return
    paths = {c["wb_path"] for c in data["cards"].values()}
    paths.update(rule["cell"][0] for rule in data.get("ignored", []))
    for item in items.values():
        if item["wb_path"] not in paths:
            continue
        item["batch_managed"] = True
        item["triage_keep"] = False
        for ignored in data.get("ignored", []):
            if _cell(item) == tuple(ignored["cell"]) and _signature(item) in ignored["signatures"]:
                if ignored.get("fixed"):
                    # 人工认可的加高：检测仍会报（误报），按已修复记账，报告里照常列出
                    item.update(status="fixed", **copy.deepcopy(ignored["fixed"]))
                else:
                    item.update(status="accepted", manual_confirmation={"accepted": True},
                                review_reason=wf.MANUAL_CONFIRMED_REASON)
                break


def _problem(item: dict) -> dict:
    """卡片上说明检测到了什么：问题类型、原因，只打印出开头一段时还有没显示的部分"""
    kind = item.get("kind") or "clip"
    error_type = autofix.item_error_type(item)
    text, visible = normalize_text(item.get("text", "")), normalize_text(item.get("visible_text") or "")
    missing = text[len(visible):] if kind == "hidden" and visible and text.startswith(visible) else ""
    reason = {"hidden": item.get("hidden_reason"), "pagebreak": page_breaks.REASON}.get(kind) or CLIP_REASON
    return {"kind": kind, "error_type": error_type, "label": KIND_LABELS.get(kind, kind),
            "reason": reason, "missing": missing}


def _remember_problems(card: dict, items: list[dict]) -> None:
    problems = {p["kind"]: p for p in card.get("problems", [])}
    problems.update((p["kind"], p) for p in map(_problem, items))  # 同一类型以最新一次检测为准
    card["problems"] = list(problems.values())
    card["error_types"] = list(dict.fromkeys(
        [*(p.get("error_type") for p in card["problems"] if p.get("error_type")),
         *(autofix.item_error_type(it) for it in items)]
    ))


def _card_fixable_types(card: dict, items: list[dict] | None = None) -> list[str]:
    items = _items_at(card) if items is None else items
    raw = [autofix.item_error_type(it) for it in items] if items else list(card.get("error_types") or [])
    if not raw:
        for p in card.get("problems") or ():
            if p.get("error_type"):
                raw.append(p["error_type"])
            else:
                raw.append(autofix.item_error_type({"kind": p.get("kind"), "hidden_reason": p.get("reason")}))
    if not raw and _is_break(card):
        raw = [autofix.ERROR_TYPE_PAGE_BREAK_CUT]
    return [t for t in dict.fromkeys(raw) if t in autofix.FIXABLE_ERROR_TYPES]


def _unchecked_fix_reason(card: dict, items: list[dict] | None = None, enabled: set[str] | None = None) -> str:
    fixable = _card_fixable_types(card, items)
    if not fixable:
        return ""
    active = _active_fix_types() if enabled is None else enabled
    if any(t in active for t in fixable):
        return ""
    labels = "、".join(autofix.error_type_label(t) for t in fixable)
    action = ("未修改" if autofix.ERROR_TYPE_PAGE_BREAK_CUT in fixable
              else "未调整列宽" if fixable == [autofix.ERROR_TYPE_HASH_OVERFLOW]
              else "未调整行高")
    return f"未勾选「{labels}」自动修复，{action}，请人工处理或标绿忽略"


def _col_rid(card: dict) -> str:
    """旧版按单元格记录的列宽编号，仅用于迁移。"""
    return _key(card["wb_path"], card["sheet_name"], card.get("cell_addr", ""), "col")


def _col_key(path: str, sheet: str, col: int) -> str:
    return _key(path, sheet, int(col), "col")


def _upgrade_column_records() -> None:
    """把旧版单元格列宽记录拆成共享的单列记录，同时迁移未完成的写入日志。持有 STATE_LOCK。"""
    data = _data()
    legacy = [r for r in data["rows"].values() if r.get("kind") == "col" and "col" not in r]
    if not legacy:
        return
    legacy_ids = {r["id"] for r in legacy}
    records = {rid: copy.deepcopy(r) for rid, r in data["rows"].items() if rid not in legacy_ids}
    job = wf.STATE.get("job") or {}
    journals = [job.get("pending_adjustment"), *job.get("deferred_adjustments", {}).values()]
    pending_fixes = {r["id"] for journal in journals if journal and journal["action"] == "fix"
                     for r in journal["rows"]}
    moved: dict[str, list[str]] = {}
    links: dict[str, list[str]] = {}
    for old in legacy:
        cid = _key(old["path"], old["sheet"], old.get("cell_addr", ""))
        card = data["cards"].get(cid) or {}
        saved = old["saved"] and not (card.get("dismissed") and card.get("color") == "blue")
        ids = moved[old["id"]] = []
        for col, base in old["base_cols"].items():
            rid = _col_key(old["path"], old["sheet"], int(col))
            applied = old["applied_cols"][col]
            previous = records.get(rid)
            if previous and previous["saved"] and (saved or old["id"] in pending_fixes):
                if abs(previous["applied"] - base) < COL_WIDTH_TOLERANCE:
                    previous["applied"] = applied
                elif abs(applied - previous["base"]) < COL_WIDTH_TOLERANCE:
                    previous["base"] = base
                else:
                    raise ValueError("旧列宽记录存在不连续的修改，请先核对原文件；恢复日志已保留")
                previous["cards"] = list(dict.fromkeys([*previous["cards"], cid]))
            elif not previous or not previous["saved"]:
                records[rid] = {"id": rid, "kind": "col", "path": old["path"], "sheet": old["sheet"],
                                "row": old["row"], "col": int(col), "base": base, "applied": applied,
                                "saved": saved, "cards": [cid], "wrap": []}
            ids.append(rid)
        links[cid] = ids
    migrated = []
    for journal in journals:
        if not journal:
            continue
        rows = []
        restored = set()
        for row in journal["rows"]:
            if row["id"] not in moved:
                rows.append(row)
                continue
            for rid in moved[row["id"]]:
                col = str(records[rid]["col"])
                if journal["action"] == "restore":
                    if rid in restored:
                        continue
                    restored.add(rid)
                    expected, target = records[rid]["applied"], records[rid]["base"]
                else:
                    expected, target = row["cols_expected"][col], row["cols_target"][col]
                rows.append({**row, "id": rid, "expected": expected, "target": target,
                             "cols_expected": {col: expected}, "cols_target": {col: target}})
        migrated.append((journal, rows))
    data["rows"] = records
    for cid, ids in links.items():
        if cid in data["cards"]:
            data["cards"][cid]["col_ids"] = ids
    for journal, rows in migrated:
        journal["rows"] = rows
    if "attempted_rows" in job:
        job["attempted_rows"] = [new for rid in job["attempted_rows"] for new in moved.get(rid, [rid])]
    wf._persist_state()


def _column_records(card: dict) -> list[dict]:
    rows = _data()["rows"]
    ids = card.get("col_ids") or [_col_rid(card)]
    return [rows[rid] for rid in ids if rid in rows and rows[rid].get("kind") == "col"
            and ("col" not in rows[rid] or card["id"] in rows[rid].get("cards", []))]


def _column_card(card: dict) -> bool:
    types = card.get("error_types")
    columns = _column_records(card)
    return (any(r["saved"] for r in columns) or autofix.ERROR_TYPE_HASH_OVERFLOW in (types or [])
            or bool(columns and not types))


def _saved_column_ids(card: dict) -> list[str]:
    ids = []
    for record in _column_records(card):
        if not record["saved"]:
            continue
        if "col" in record:
            ids.append(record["id"])
        else:
            ids.extend(_col_key(record["path"], record["sheet"], int(col)) for col in record["applied_cols"])
    return ids


def _row_record(card: dict) -> dict | None:
    columns = _column_records(card) if _column_card(card) else []
    if columns:
        return {"kind": "col", "base": sum(r["base"] for r in columns),
                "applied": sum(r["applied"] for r in columns), "saved": any(r["saved"] for r in columns),
                "applied_cols": {str(r["col"]): r["applied"] for r in columns if "col" in r}
                or columns[0].get("applied_cols", {})}
    return None if _column_card(card) else _data()["rows"].get(_rid(card))


def _latest_shot(card: dict, items: list[dict]) -> str:
    """本行没改、因别处调整而通过的格子：从刚导出的 PDF 重新截图，不沿用之前尝试或检测时的旧图。
    文字太短（容易对到别的格子）或找不到时不截图，宁可显示「暂无截图」也不放错的图。"""
    if len(normalize_text(card.get("text", ""))) < 4:
        return ""
    for pdf_path in dict.fromkeys(it.get("pdf_path") for it in items if it.get("pdf_path")):
        try:
            with fitz.open(pdf_path) as doc:
                found = locate_text(collect_text_pieces(doc), card["text"])
                if found is not None:
                    return render_crop_base64(doc[found[0]], found[1], highlight=False, zoom=4.0)
        except Exception:
            continue
    return ""


def _save_shot(card: dict, which: str, image: str) -> None:
    # 卡片 id 按单元格固定：同一格以后再出问题会新建同 id 的卡，带上代次才不会覆盖旧修复留在报告里的截图
    gen = card.get("gen")
    fix_images.save(card, which, image, key=f"batch:{card['id']}" + (f":{gen}" if gen else ""))


def _stopping(job: dict) -> bool:
    """点了「完成当前行后暂停」或「转人工」：做完当前这一行就停"""
    return bool(job.get("pause_requested") or job.get("to_manual"))


def open_block_reason(job: dict, path: Path) -> str:
    """后台任务在跑时，网页上点这份文件能不能在 Excel 打开：能打开返回空字符串，不能就返回原因（页面上直接显示）。

    还没轮到、正在处理的文件要等：用户在 Excel 里开着，后台轮到它时会因为文件被占用跳过。
    已经做完、后面不再碰的文件，和不在这次任务里的文件，可以打开。调用方持有 STATE_LOCK。"""
    kind = job.get("kind")
    if kind not in {"batch_scan", "batch_recheck"} or "workbooks" not in job:
        # 提交、清除标色、精简 Sheet、导出 PDF 等任务：每份文件做什么不好判断，照旧等它结束
        return "后台任务正在处理文件，结束后才能打开"
    target = path.resolve()
    mine = [p for p in job["workbooks"] if Path(p).resolve() == target]
    if not mine:
        return ""
    completed = job.get("completed", [])
    if kind == "batch_recheck":
        if any(p == job.get("current") for p in mine):
            return "正在复检这份文件，查完才能打开"
        if any(p in job.get("again", []) for p in mine):
            return "这份文件查完还要再复检一遍，查完才能打开"
        return "" if all(p in completed for p in mine) else "这份文件在排队等复检，查完才能打开"
    if job.get("to_manual"):
        return "正在转人工，要给这份文件标色，标完才能打开"
    if any(p in job.get("deferred_adjustments", {}) for p in mine):
        return "这份文件还有没写完的修改，后台还要回来处理"
    if job.get("options", {}).get("scan_only") and any(Path(p).resolve() == target for p in _auto_queue_paths(job)):
        return "排队等自动调整（全部文件扫完后开始），调整完才能打开"
    if all(p in completed for p in mine):
        return ""
    scanned = job.get("scanned", [])
    repairing = job.get("phase") == "repair" and not job.get("options", {}).get("scan_only")
    pending = [p for p in job["workbooks"] if p not in (completed if repairing else scanned)]
    current = (job.get("repairing") if repairing else None) or (pending[0] if pending else None)
    if current in mine:
        return "后台正在处理这份文件，处理完才能打开"
    if all(p in scanned for p in mine):
        return "已扫描，还要等自动修复，修完才能打开"
    return "还没轮到这份文件，处理完才能打开"


def _open_blocked(paths: list[str]) -> dict[str, str]:
    """页面上各文件现在能不能点开 Excel：只列不能打开的，值是原因。调用方持有 STATE_LOCK。"""
    job = wf.STATE.get("job") or {}
    kind = job.get("kind") or ""
    if not wf.STATE.get("busy") or not (kind.startswith("batch_") or kind in wf.QUICK_TASK_NAMES):
        return {}
    return {path: reason for path in dict.fromkeys(paths) if (reason := open_block_reason(job, Path(path)))}


def _waiting(card: dict) -> bool:
    """还没轮到自动处理的卡（旧数据没有 waiting 字段，按说明文字判断）"""
    return bool(card.get("waiting", card.get("message") == WAITING_MESSAGE))


def _adoptable(card: dict) -> bool:
    """自动检查没通过、但留下了「最后一次尝试」截图的卡：用户看图确认已显示完整时可以标蓝，
    提交时按那次尝试的行高写入原文件（检测器可能误报，人眼判断为准）。"""
    heights = (card.get("base_height"), card.get("attempt_height"))
    return (not card.get("verified") and card.get("after_kind") == "attempt" and not _waiting(card)
            and all(isinstance(h, (int, float)) and math.isfinite(h) for h in heights)
            and bool(card.get("fix_images", {}).get("after")))


def _adopting(card: dict) -> bool:
    return card["color"] == "blue" and _adoptable(card)


def _items_at(card: dict) -> list[dict]:
    return [it for it in wf.STATE["items"].values() if _cell(it) == _cell(card)]


def _remember_context(card: dict, items: list[dict]) -> None:
    """整页图和框的位置一起固定，后续加高、重新分页后不套用到新页面。"""
    if card.get("context"):
        return
    for item in items:
        page_id = item.get("page_id")
        page = wf.STATE["pages"].get(page_id)
        bbox = item.get("bbox")
        if (not page or page.get("wb_path") != item["wb_path"] or page.get("sheet_name") != item["sheet_name"]
                or not isinstance(bbox, (list, tuple)) or len(bbox) != 4):
            continue
        width, height = page.get("page_width"), page.get("page_height")
        if not all(isinstance(n, (int, float)) and math.isfinite(n) for n in [width, height, *bbox]):
            continue
        if width <= 0 or height <= 0 or bbox[2] <= bbox[0] or bbox[3] <= bbox[1]:
            continue
        name = page.get("batch_context_image", "")
        if not fix_images.image_path(name):
            # 同一页的多个单元格共用一张图，不给每张卡片重复保存整页 PNG。
            holder = {"id": page_id}
            fix_images.save(holder, "context", page.get("image", ""), key=f"batch-page:{page_id}")
            name = holder.get("fix_images", {}).get("context", "")
            if not name:
                continue
            page["batch_context_image"] = name
        card["context"] = {"image_file": name, "page_index": page.get("page_index"),
                           "page_width": width, "page_height": height, "bbox": list(bbox)}
        return


def _card_for(items: list[dict]) -> dict:
    first = items[0]
    data = _data()
    cid = _key(*_cell(first))
    card = data["cards"].get(cid)
    signatures = [_signature(it) for it in items]
    if card is None or card.get("dismissed"):
        card = data["cards"][cid] = {
            "id": cid, "wb_path": first["wb_path"], "wb_name": first["wb_name"],
            "sheet_name": first["sheet_name"], "cell_addr": first.get("cell_addr", ""),
            "excel_row": int(first["excel_row"]), "seq": first["seq"],
            "text": first["text"], "color": "yellow", "dismissed": False,
            "finalized": False, "verified": False, "manual": False, "message": WAITING_MESSAGE,
            "waiting": True, "gen": uuid4().hex[:8], "signatures": signatures,
        }
        if _is_break(first):
            card["kind"] = "pagebreak"
        _save_shot(card, "before", first.get("image", ""))
    card["item_ids"] = [it["id"] for it in items]
    _remember_problems(card, items)
    for signature in signatures:
        if signature not in card["signatures"]:
            card["signatures"].append(signature)
    _remember_context(card, items)
    return card


def _image_url(name: str) -> str:
    # 同名文件会被后面的尝试覆盖：带版本号，内容变了地址才变，浏览器可以放心缓存
    return f"/api/batch/image/{name}?v={fix_images.version(name)}"


def snapshot() -> dict:
    with wf.STATE_LOCK:
        cards = []
        by_cell: dict[tuple[str, str, str], list[dict]] = {}
        for item in wf.STATE["items"].values():
            by_cell.setdefault(_cell(item), []).append(item)
        file_versions: dict[str, Any] = {}
        image_urls: dict[str, str] = {}
        folder_paths: dict[str, bool] = {}

        def in_folder(path: str) -> bool:
            if path not in folder_paths:
                folder_paths[path] = wf.in_current_folder(path)
            return folder_paths[path]

        def image_url(name: str) -> str:
            if name not in image_urls:
                image_urls[name] = _image_url(name) if fix_images.image_path(name) else ""
            return image_urls[name]

        for card in sorted(_data()["cards"].values(), key=lambda c: c["seq"]):
            if card.get("dismissed") or not in_folder(card["wb_path"]):
                continue
            items = by_cell.get(_cell(card), [])
            # 释放 STATE_LOCK 后才序列化 JSON，返回值仍需与后台线程的状态隔离。
            result = copy.deepcopy({k: v for k, v in card.items() if k not in {"signatures", "item_ids"}})
            if not result.get("problems"):  # 旧版本留下的卡片没存问题说明，按单元格现有的问题项补上
                result["problems"] = list({p["kind"]: p for p in map(_problem, items)}.values())
            if not result.get("error_types"):
                result["error_types"] = list(dict.fromkeys(
                    p.get("error_type") or autofix.item_error_type({"kind": p.get("kind"), "hidden_reason": p.get("reason")})
                    for p in result["problems"]
                ))
            row = _row_record(card)
            if row is not None and _is_break(card):
                result["page_break"] = {"row": row["row"], "saved": row["saved"], "applied": row["applied"] >= .5}
            elif row is not None and row.get("kind") == "col":
                result["col_width"] = {"base": row["base"], "applied": row["applied"], "saved": row["saved"],
                                       "cols": list((row.get("applied_cols") or {}).keys()),
                                       "ids": _saved_column_ids(card)}
            elif row is not None:
                result["row_height"] = {"base": row["base"], "applied": row["applied"], "saved": row["saved"],
                                        "wrap": card.get("cell_addr") in row.get("wrap", [])}
            result["adoptable"] = _adoptable(card)
            if card["color"] == "yellow" and card.get("finalized"):
                # 自动试过没修好（文件没改过）等不能交给自动的黄卡：页面上不给「🤖 自动调整」按钮，写明原因
                result["auto_skip"] = _auto_skip_reason(card, items, file_versions=file_versions)
            result["highlighted"] = _shows_mark(_marks().get(card["id"]))
            result["images"] = {
                which: url for which, name in card.get("fix_images", {}).items()
                if (url := image_url(name))
            }
            if result.get("context"):
                context = result["context"]
                name = context["image_file"]
                context["image"] = image_url(name)
            cards.append(result)
        highlighted = sum(1 for m in _marks().values() if in_folder(m["path"]) and _shows_mark(m))
        tabs = sum(len(t) for p, t in _tabs().items() if in_folder(p))
        return {"cards": cards, "busy": wf.STATE["busy"], "folder": wf.STATE["folder"], "highlighted": highlighted,
                "tab_marks": tabs,
                "error_types": list(autofix.CELL_ERROR_SPECS.values()),
                "fix_error_types": fix_error_types(),
                "mode": _mode(by_cell, in_folder), "open_blocked": _open_blocked([c["wb_path"] for c in cards]),
                "followups": copy.deepcopy(_data().get("followups") or []),
                "scan_eta": _scan_eta(wf.STATE["job"]) if wf.STATE["job"] else None,
                "slow_files": [{"path": p, "name": Path(p).name, **info} for p, info in _data().get("slow_files", {}).items()
                               if in_folder(p)],
                "job": copy.deepcopy(wf.STATE["job"]), "progress": dict(wf.STATE["scan_progress"]),
                "logs": wf.STATE["logs"][-100:]}


def _mark_block_reason(path: str) -> str:
    """后台任务在跑时这份文件的卡能不能改颜色：能改返回空字符串。调用方持有 STATE_LOCK。
    改颜色只写记录、不碰 Excel：自动处理 / 复检时，不在这次任务里的文件、已经做完的文件都能改；
    还没轮到和正在处理的文件要等（后台还要读它们的卡）。提交、清除标色等任务照旧等它结束"""
    if not wf.STATE["busy"]:
        return ""
    job = wf.STATE.get("job") or {}
    if job.get("kind") not in {"batch_scan", "batch_recheck"} or "workbooks" not in job:
        return "任务执行中，请结束或暂停后再审核"
    return open_block_reason(job, Path(path)).replace("才能打开", "才能改颜色")


def mark(cid: str, color: str) -> None:
    with wf.STATE_LOCK:
        card = _data()["cards"].get(cid) if isinstance(cid, str) else None
        if color not in COLORS or card is None or card.get("dismissed") or not wf.in_current_folder(card["wb_path"]):
            raise ValueError("请选择当前文件夹的有效卡片及颜色")
        if reason := _mark_block_reason(card["wb_path"]):
            raise ValueError(reason)
        previous = copy.deepcopy(card)
        card.update(color=color, finalized=False, marked=True, waiting=False)
        try:
            wf._persist_state()
        except Exception:
            card.clear()
            card.update(previous)
            raise


def mark_many(ids: Any, color: str) -> dict:
    """批量改色，只写一次数据库；三色均由用户自由选择。"""
    with wf.STATE_LOCK:
        if color not in COLORS or not isinstance(ids, list) or not ids:
            raise ValueError("请选择卡片及颜色")
        cards = _data()["cards"]
        targets = []
        for cid in dict.fromkeys(ids):
            card = cards.get(cid) if isinstance(cid, str) else None
            if card is None or card.get("dismissed") or not wf.in_current_folder(card["wb_path"]):
                raise ValueError("选中的卡片已失效，请刷新页面后重新选择")
            if reason := _mark_block_reason(card["wb_path"]):
                raise ValueError(f"{card['wb_name']}：{reason}（请先取消选中这个文件的卡）")
            targets.append(card)
        changed = targets
        previous = [copy.deepcopy(c) for c in changed]
        for card in changed:
            card.update(color=color, finalized=False, marked=True, waiting=False)
        try:
            wf._persist_state()
        except Exception:
            for card, old in zip(changed, previous):
                card.clear()
                card.update(old)
            raise
        return {"changed": len(changed), "skipped": 0}


def _check_file(path: str) -> None:
    p = Path(path)
    if not p.is_file():
        raise ValueError(f"找不到文件：{p.name}")
    if workbook_locked(p):
        raise ValueError(f"文件仍在 Excel 中打开，请保存并关闭后继续：{p.name}" + STALE_LOCK_HINT)


def _version_or_none(path: str) -> Any:
    try:
        return wf._file_version(Path(path))
    except OSError:
        return None


def _followed_own_save(path: str, before: Any) -> None:
    """程序自己保存了文件（写行高、标色）：自动处理完时记下的版本跟着更新，不算你改过文件，
    自动没修好的格子也就不会因此被「自动调整」重试"""
    with wf.STATE_LOCK:
        versions = _data().get("auto_versions", {})
        if before is not None and versions.get(path) == before:
            versions[path] = _version_or_none(path)


STEP_COL_WIDTH = 2.0
MAX_COL_WIDTH = 255.0
COL_WIDTH_TOLERANCE = 0.4


def _row_value(book: Any, row: dict) -> float:
    """写入日志里一行现在的值：行高；插分页的行是 1（上方有手动分页）/ 0（没有）；列宽是各列宽度之和"""
    sheet = book.Sheets(row["sheet"])
    if row.get("kind") == "col":
        cols = row.get("cols_target") or row.get("cols_expected") or {}
        return round(sum(float(sheet.Columns(int(c)).ColumnWidth) for c in cols), 2)
    target = sheet.Rows(row["row"])
    if row.get("kind") == "break":
        return 1.0 if int(target.PageBreak) == page_breaks.XL_PAGEBREAK_MANUAL else 0.0
    return float(target.RowHeight)


def _cols_match(book: Any, row: dict, field: str) -> bool:
    if row.get("kind") != "col":
        return True
    sheet = book.Sheets(row["sheet"])
    return all(
        abs(float(sheet.Columns(int(c)).ColumnWidth) - float(w)) < COL_WIDTH_TOLERANCE
        for c, w in (row.get(field) or {}).items()
    )


def _value_matches(book: Any, value: float, row: dict, field: str) -> bool:
    if row.get("kind") == "col":
        return _cols_match(book, row, "cols_" + field)
    return abs(value - row[field]) < ROW_HEIGHT_TOLERANCE


def _adjustment_where(row: dict) -> str:
    if row.get("kind") == "col":
        cols = row.get("cols_target") or row.get("cols_expected") or {str(row["col"]): None}
        return f"{row['sheet']} 第 {'、'.join(cols)} 列"
    return f"{row['sheet']} 第 {row['row']} 行"


def _write_rows(excel: Any, journal: dict) -> None:
    """先核对所有修改，再写；续跑时目标值已保存则只补数据库提交。"""
    path = journal["path"]
    where = f"{Path(path).name}（" + "、".join(map(_adjustment_where, journal["rows"])) + "）"
    _check_file(path)
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError(f"{where} 文件以只读方式打开，未写入")
        actual = [_row_value(book, r) for r in journal["rows"]]
        for value, r in zip(actual, journal["rows"]):
            debug_log.event("✍️ 写入行", 动作=journal.get("action"), 位置=_adjustment_where(r),
                            类型=r.get("kind") or "行高", 现在=value, 预期原值=r["expected"], 目标=r["target"],
                            换行=r.get("wrap") or "-")
        # 自动换行和行高一起保存；续跑时两样都已到位才算写过
        wraps = [(r["sheet"], addr, on) for r in journal["rows"] for addr, on in (r.get("wrap") or {}).items()]
        at_target = (all(_value_matches(book, h, r, "target")
                         for h, r in zip(actual, journal["rows"]))
                     and all(bool(book.Sheets(sh).Range(addr).MergeArea.WrapText) == on for sh, addr, on in wraps))
        if wf._file_version(Path(path)) != journal["version"] and not at_target:
            raise ValueError(f"{where} 文件在准备后被修改，已停止写入；核对后可继续任务，或点「放弃这次写入」")
        if not at_target:
            changed = [r for h, r in zip(actual, journal["rows"]) if not _value_matches(book, h, r, "expected")]
            if changed:
                kind = "列宽" if changed[0].get("kind") == "col" else "分页" if changed[0].get("kind") == "break" else "行高"
                raise ValueError(f"{where} {kind}已被其他操作修改，未覆盖人工修改；可点「放弃这次写入」转人工处理")
            for sheet, addr, on in wraps:
                book.Sheets(sheet).Range(addr).MergeArea.WrapText = on
            for row in journal["rows"]:
                sheet = book.Sheets(row["sheet"])
                if row.get("kind") == "col":
                    for col_str, w in (row.get("cols_target") or {}).items():
                        sheet.Columns(int(col_str)).ColumnWidth = float(w)
                elif row.get("kind") == "break":
                    sheet.Rows(row["row"]).PageBreak = (
                        page_breaks.XL_PAGEBREAK_MANUAL if row["target"] else page_breaks.XL_PAGEBREAK_NONE
                    )
                else:
                    sheet.Rows(row["row"]).RowHeight = row["target"]
            normalize_workbook_view(excel, book)
            before = _version_or_none(path)
            book.Save()
            _followed_own_save(path, before)
            debug_log.event("💾 已保存", 文件=Path(path).name,
                            写后=[round(_row_value(book, r), 2) for r in journal["rows"]])
        else:
            debug_log.event("💾 目标值已在文件里（续跑），不再写", 文件=Path(path).name)
    finally:
        book.Close(SaveChanges=False)


def _wrapped(pending: dict, card: dict) -> list[str]:
    """这次写入在卡片所在行开启自动换行的单元格"""
    return [addr for r in pending["rows"] if r["row"] == card["excel_row"] and r["sheet"] == card["sheet_name"]
            for addr, on in (r.get("wrap") or {}).items() if on]


def _finish_pending(excel: Any) -> None:
    with wf.STATE_LOCK:
        _upgrade_column_records()
    job = wf.STATE["job"]
    pending = job.get("pending_adjustment")
    if not pending:
        return
    _write_rows(excel, pending)
    with wf.STATE_LOCK:
        data = _data()
        for row in pending["rows"]:
            record = data["rows"][row["id"]]
            # 人工认可的加高等同正式接受，不再属于可还原的待审核修改
            record["saved"] = pending["action"] == "fix"
            record["applied"] = row["target"]
            if pending["action"] != "restore":
                record["wrap"] = [addr for addr, on in (row.get("wrap") or {}).items() if on]
            if pending["action"] == "fix" and row["id"] not in job.setdefault("attempted_rows", []):
                job["attempted_rows"].append(row["id"])
        for cid in pending["cards"]:
            card = data["cards"][cid]
            if pending["action"] == "fix" and _is_break(card):
                card.update(color="blue", verified=True, after_kind="fixed",
                            message="已在合并区域上方插入分页并保存，等待审核")
                for item in _items_at(card):
                    item.update(status="fixed", batch_managed=True, triage_keep=False, page_break=True,
                                fix_images=copy.deepcopy(card.get("fix_images", {})),
                                review_reason="全自动批处理：已修复并保存")
            elif pending["action"] == "fix" and any(r.get("kind") == "col" for r in pending["rows"]):
                card.update(color="blue", verified=True, after_kind="fixed",
                            message="已自动加宽列并保存，等待审核")
                for item in _items_at(card):
                    item.update(status="fixed", batch_managed=True, triage_keep=False, col_widened=True,
                                fix_images=copy.deepcopy(card.get("fix_images", {})),
                                review_reason="全自动批处理：已自动加宽列并保存")
            elif pending["action"] == "fix":
                card.update(color="blue", verified=True, after_kind="fixed", message="已自动修复并保存，等待审核")
                for item in _items_at(card):
                    item.update(status="fixed", batch_managed=True, triage_keep=False,
                                wrap_text=card.get("cell_addr") in _wrapped(pending, card),
                                row_height=next(r["target"] for r in pending["rows"] if r["row"] == card["excel_row"]
                                                and r["sheet"] == card["sheet_name"]),
                                fix_images=copy.deepcopy(card.get("fix_images", {})),
                                review_reason="全自动批处理：已修复并保存")
            elif pending["action"] == "adopt":
                card.update(verified=True, adopted=True, after_kind="fixed",
                            message="已按最后一次尝试" + ("开启自动换行、" if _wrapped(pending, card) else "")
                            + "加高并保存",
                            adopted_height=next(r["target"] for r in pending["rows"] if r["row"] == card["excel_row"]
                                                and r["sheet"] == card["sheet_name"]))
            else:
                card["message"] = "已还原自动修改，等待人工处理" if card["color"] == "yellow" else "已还原，待提交忽略"
        job["pending_adjustment"] = None
        try:
            wf._persist_state()
        except Exception:
            job["pending_adjustment"] = pending
            raise
        wf._sync_problem_counts(pending["path"], {"fix": "自动处理已保存", "adopt": "审核加高已保存"}.get(
            pending["action"], "审核还原已保存"))


def _park_pending(error: Exception) -> None:
    """隔离本文件的未结算写入；保存日志后允许继续其他文件，不猜测是否已保存。"""
    with wf.STATE_LOCK:
        job = wf.STATE["job"]
        pending = job.get("pending_adjustment")
        if not pending:
            return
        job.setdefault("deferred_adjustments", {})[pending["path"]] = pending
        job["pending_adjustment"] = None
        for cid in pending["cards"]:
            card = _data()["cards"].get(cid)
            if card and pending["action"] == "fix":
                card.update(color="yellow", verified=False, waiting=False,
                            message=f"文件写入状态待核对，已跳过并保留恢复日志：{error}")
            elif card:
                # 提交时的还原 / 采用：颜色是你标的，不替你改，等你继续或放弃这次写入
                card["message"] = f"文件写入待核对，这份文件先不处理：{error}"
        wf._persist_state()


def _unsettled(job: dict, path: str) -> bool:
    """这份文件还有没写完的修改（写入核对没通过）：提交 / 复检 / 清除颜色先跳过它，不挡别的文件"""
    if path not in job.get("deferred_adjustments", {}):
        return False
    wf.log_msg(f"⏸ {Path(path).name} 上次写入还没完成，这次先跳过；点「继续上次任务」重新核对写入，"
               "或点「放弃这次写入」转人工后再处理")
    return True


def _show_deferred_pending() -> None:
    """任务结束或暂停后沿用现有 UI，逐份展示可续跑 / 放弃的写入。"""
    job = wf.STATE["job"]
    deferred = job.get("deferred_adjustments", {})
    if not job.get("pending_adjustment") and deferred:
        path = next(iter(deferred))
        job["pending_adjustment"] = deferred.pop(path)


def abandon_pending() -> None:
    """写入日志结算不了（行高被人工改过、文件被移走等）时放弃这次写入：不动文件，相关卡片转黄色人工处理。
    否则除了「继续上次任务」什么都不能开始，而继续又会报同样的错。"""
    with wf.STATE_LOCK:
        job = wf.STATE["job"] or {}
        pending = job.get("pending_adjustment")
        if wf.STATE["busy"] or not pending or not job.get("kind", "").startswith("batch_"):
            raise ValueError("没有需要放弃的写入")
        data = _data()
        for row in pending["rows"]:
            record = data["rows"].get(row["id"])
            if record is not None:
                record["saved"] = False  # 不确定文件里现在是哪个行高，不能再自动还原
        for cid in pending["cards"]:
            card = data["cards"].get(cid)
            if card is None:
                continue
            target = "相关列的列宽" if any(r.get("kind") == "col" for r in pending["rows"]) else "这一行的行高"
            card.update(color="yellow", verified=False, manual=True, finalized=True, dismissed=False, waiting=False,
                        message=f"已放弃未完成的自动写入，请在 Excel 中核对{target}，保存后复检")
            for item in _items_at(card):
                item.update(status="manual_review", batch_managed=True, triage_keep=False,
                            review_reason=card["message"])
        job["pending_adjustment"] = None  # 任务本身仍可「继续上次任务」，这几格已转人工，不会再自动写
        _show_deferred_pending()
        wf._persist_state()
    wf.log_msg(f"⚠️ 已放弃 {Path(pending['path']).name} 未完成的写入，相关单元格转为人工处理")
    wf._sync_problem_counts(pending["path"], "已放弃未完成的写入，等待人工处理")


def _error_value_card(card: dict) -> bool:
    """旧版把 #REF! 等公式错误（COM 读成 -2146826265 这类整数）当成隐藏截断出了卡片，现在不再检测"""
    try:
        return is_excel_error(int(str(card.get("text", "")).strip()))
    except ValueError:
        return False


def _refresh(
    excel: Any, path: str, sheets: list[str] | None = None, *,
    session: Any = None, workbook: Any = None, reuse: bool = True, timing: dict[str, float] | None = None,
) -> list[str]:
    """重新检测（默认整个文件）。session 给了就沿用同一状态已导出的 PDF（reuse=False 时一律重新导出）；
    workbook 给了就直接用它检测（全自动的副本，或复检时在内存里还原过标记色的原文件），不再打开原文件。
    返回这次找不到、没有检测的工作表（被删除、隐藏，或改名后认不出来）。"""
    with wf.STATE_LOCK:
        cards = _data()["cards"]
        for cid in [cid for cid, card in cards.items()
                    if card["wb_path"] == path and (sheets is None or card["sheet_name"] in sheets)
                    and _error_value_card(card) and not card.get("finalized")]:
            del cards[cid]
        # 升级前人工跳过的记录也转成明确的单元格忽略规则，后续复检仍有效。
        for item in wf.STATE["items"].values():
            if (item["wb_path"] == path and item["status"] == "accepted"
                    and not item.get("batch_managed") and item.get("cell_addr")
                    and item.get("location_confirmed") is not False):
                rule = {"cell": list(_cell(item)), "signatures": [_signature(item)]}
                if rule not in _data()["ignored"]:
                    _data()["ignored"].append(rule)
        previous = [copy.deepcopy(it) for it in wf.STATE["items"].values()
                    if it["wb_path"] == path and (sheets is None or it["sheet_name"] in sheets)]
    entry = {"mode": "sheets" if sheets else "full", "sheets": sheets, "scanned": bool(previous)}
    wf._scan_one_workbook(
        excel, Path(path), entry,
        previous, job=None, resumed=True, rescan_all=False, delete_backups=False,
        used_pdf_paths={Path(p["pdf_path"]) for p in wf.STATE["pages"].values()},
        read_only=True, workbook=workbook, pdfs=session, reuse_pdfs=reuse, timing=timing,
    )
    with wf.STATE_LOCK:
        missing = list(entry.get("missing", []))
        if workbook is None:
            # 记下这几张表是按原文件哪个版本检测的：文件没再改过时复检结果必然一样，不用重新导出
            checked = wf.STATE["checked_files"].get(path)
            if sheets is None:
                _sheet_checks()[path] = {"*": checked}
            else:
                checks = _sheet_checks().setdefault(path, {})
                for name in missing:
                    checks.pop(name, None)
                checks.update(dict.fromkeys(entry["sheets"], checked))
        for item in wf.STATE["items"].values():
            if item["wb_path"] == path:
                item.update(batch_managed=True, triage_keep=False)
        carry_item_policies(wf.STATE["items"])
        # 一开始检出的每个单元格都留卡；前面一行的修改顺带解决的其他格子也应可审核。
        for group in _group_cells([it for it in wf.STATE["items"].values() if it["wb_path"] == path
                                  and it["status"] != "accepted"]):
            cid = _key(*_cell(group[0]))
            unresolved = any(it["status"] not in wf.RESOLVED_STATUSES for it in group)
            existing = _data()["cards"].get(cid)
            if unresolved:
                card = _card_for(group)
                if card.get("color") == "blue" or card.get("verified"):
                    # 后面的行加高、还原改变了分页，之前通过的格子又出问题：不能再当作已修好。
                    # 用户手动选择的三色均保留；只有程序自动标蓝的结果才自动转黄。
                    keep = bool(card.get("marked"))
                    card.update(color=card["color"] if keep else "yellow", verified=False,
                                message="最新导出仍有问题" + ("，已保留你的标记" if keep else "，需重新处理"))
            elif existing and not existing.get("dismissed") and not existing.get("manual") and not existing.get("marked"):
                row = _row_record(existing)
                changed = (any(abs(r["applied"] - r["base"]) >= COL_WIDTH_TOLERANCE
                               for r in _column_records(existing)) if row and row.get("kind") == "col" else
                           bool(row and abs(row["applied"] - row["base"]) >= ROW_HEIGHT_TOLERANCE))
                if row and row["saved"] and changed:
                    existing.update(color="blue", verified=True, after_kind="fixed", waiting=False,
                                    message=("已插入分页" if _is_break(existing) else
                                             "已加宽列" if row.get("kind") == "col" else "已加高")
                                    + "并保存，最新导出已通过检查，等待审核")
                    after = next(((it.get("fix_images") or {}).get("after") for it in group
                                  if (it.get("fix_images") or {}).get("after")), None)
                    if after:
                        existing.setdefault("fix_images", {})["after"] = after
                else:
                    # 这一行没写过：之前留下的「修改后」可能是试算第一张（和修改前一模一样），必须换成最新导出的图
                    fix_images.forget(existing, "after")
                    _save_shot(existing, "after", _latest_shot(existing, group))
                    existing.update(color="blue", verified=True, after_kind="rechecked", waiting=False,
                                    message="未插入分页：其他行调整后重新导出，已不再被分页切开，等待审核"
                                    if _is_break(existing) else
                                    "本格未单独调整列宽：其他单元格调整后重新导出，已完整显示，等待审核"
                                    if _column_card(existing) else
                                    "本行行高未修改：其他行调整后重新导出，这一格已完整显示，等待审核")
        wf._persist_state()
    return missing


def _group_cells(items: list[dict]) -> list[list[dict]]:
    cells: dict[tuple, list[dict]] = {}
    for item in items:
        cells.setdefault(_cell(item), []).append(item)
    return list(cells.values())


def _wrap_targets(sheet: Any, items: list[dict]) -> list[str]:
    """没开自动换行、整行文字横向放不下被截掉的格子：怎么加高都没用，要先开自动换行再加高。
    只改这一格（合并单元格按整个合并区域），不动列宽，整页排版不变。
    没合并的格子旁边是空格时，文字本来会延伸到旁边显示，截断多半不是横向的，不开换行，免得改变原来的显示方式。"""
    targets = []
    for item in items:
        if item.get("kind") != "hidden" or item.get("hidden_reason") not in (REASON_TRUNCATED, REASON_CLIPPED):
            continue
        addr = item.get("cell_addr") or ""
        try:
            area = sheet.Range(addr).MergeArea
            if area.WrapText is not False or area.ShrinkToFit:
                continue  # 已经换行（或合并区域里设置不一致），缩小字体填充的也不动
            row, col = int(area.Row), int(area.Column)
            if int(area.Cells.Count) == 1:
                align = area.HorizontalAlignment
                sides = ([col - 1] if align == XL_HALIGN_RIGHT else [col - 1, col + 1] if align == XL_HALIGN_CENTER
                         else [col + 1])
                if not any(c > 0 and (sheet.Cells(row, c).MergeCells or str(sheet.Cells(row, c).Formula) != "")
                           for c in sides):
                    continue
        except Exception:
            continue  # 读不到格式就按原来的方式只加高
        if addr not in targets:
            targets.append(addr)
    return targets


def _prepare_row(
    session: Any, path: str, items: list[dict], *,
    on_attempt: Callable[[int, float, dict, list[str]], None] | None = None,
    timing: dict[str, float] | None = None,
) -> tuple[float, dict | None, dict]:
    """只在副本上做图；逐次交出截图，即使随后导出失败也能保留最后已生成的效果。
    原样没通过、有不换行被横向截断的格子时，从第一次加高起先给这些格子开自动换行。
    原样那张直接用刚才检测导出的 PDF（同一版本）；通过的方案正好是最后一次导出时，
    chosen["pdf"] 留下那份 PDF，写入后重新检测可直接用。"""
    timing = {} if timing is None else timing
    sheet = session.copy_of(path).Sheets(items[0]["sheet_name"])
    exported = session.pdf_for(path, items[0]["sheet_name"])
    autofix.validate_print_targets(sheet, items)
    wrap = _wrap_targets(sheet, items)
    wrapped: list[str] = []
    row = int(items[0]["excel_row"])
    base = round(float(sheet.Rows(row).RowHeight), 2)
    cells = snapshot_cells(extract_sheet_snapshot(sheet))
    probe = CellProbe(sheet)
    groups = _group_cells(items)
    before: dict[str, str] = {}
    chosen = None
    extra = 0.0
    scale = None
    success = False
    pdf = session.temp_pdf()
    try:
        for attempt in range(autofix.MAX_ATTEMPTS + 1):
            height = min(autofix.MAX_ROW_HEIGHT, round(base + extra, 2))
            if attempt and height <= last_height:
                break
            if attempt == 1 and wrap:
                for addr in wrap:
                    sheet.Range(addr).MergeArea.WrapText = True
                wrapped = list(wrap)
            sheet.Rows(row).RowHeight = height
            started = time.perf_counter()
            source = exported if attempt == 0 and exported is not None else pdf
            if source == pdf:
                wf._export_sheet_pdf(sheet, pdf)
                timing["exports"] = timing.get("exports", 0) + 1
            timing["export"] = timing.get("export", 0.0) + time.perf_counter() - started
            started = time.perf_counter()
            # 每次行高改变后重建 probe，避免沿用 Range.Text 缓存。
            probe = CellProbe(sheet)
            ok, _, need = autofix._evaluate(source, items, cells, probe.text, probe.printed)
            shots = {_key(*_cell(group[0])): autofix._evaluate(source, group, cells, probe.text, probe.printed)[1]
                     for group in groups}
            timing["evaluate"] = timing.get("evaluate", 0.0) + time.perf_counter() - started
            debug_log.event("🧪 试算", 次=attempt, 行高=height, 原行高=base, 通过=ok, 缺口pt=need,
                            打印缩放=scale, 换行=wrapped or "-", 用检测时的PDF=source != pdf,
                            比对秒=time.perf_counter() - started)
            if on_attempt is not None:
                on_attempt(attempt, height, shots, wrapped)
            if attempt == 0:
                before = shots
            if ok:
                chosen = {"height": height, "shots": shots, "wrap": wrapped}
            if (attempt == 0 and ok) or success:
                if attempt and ok:
                    # 最后这次导出就是要写入的状态
                    shutil.copyfile(pdf, session.chosen_pdf())
                    chosen["pdf"] = session.chosen_pdf()
                break
            success = ok
            last_height = height
            if need:
                # 缺口是在 PDF 上量的，打印缩放 70% 时要除以 0.7 才是实际要加的行高
                if scale is None:
                    scale = autofix.print_scale(sheet)
                need /= scale
            extra = max(extra + autofix.STEP_PT,
                        math.ceil(round((extra + need) / autofix.STEP_PT, 3)) * autofix.STEP_PT)
    finally:
        sheet.Rows(row).RowHeight = base
        for addr in wrapped:
            sheet.Range(addr).MergeArea.WrapText = False
        pdf.unlink(missing_ok=True)
    return base, chosen, before


def _cell_columns(sheet: Any, addr: str) -> tuple[int, list[int]]:
    match = re.fullmatch(r"([A-Za-z]+)([1-9]\d*)", (addr or "").split(":")[0])
    if match is None:
        raise ValueError("单元格地址无法确认，未调整列宽，请人工核对")
    row = int(match[2])
    first_col = 0
    for ch in match[1].upper():
        first_col = first_col * 26 + ord(ch) - ord("A") + 1
    col_count = 1
    try:
        area = sheet.Range(addr).MergeArea
        c0, cnt = int(area.Column), int(area.Columns.Count)
        if c0 >= 1 and cnt >= 1:
            first_col, col_count = c0, cnt
    except Exception:
        pass
    return row, list(range(first_col, first_col + col_count))


def _prepare_column(session: Any, path: str, items: list[dict], timing: dict[str, float]) -> dict:
    """### 列宽不足：单列直接加宽该列，跨多列合并时平均微调合并区各列，直到 ### 消失并留余量"""
    item = items[0]
    sheet = session.copy_of(path).Sheets(item["sheet_name"])
    row, cols = _cell_columns(sheet, item.get("cell_addr") or "")
    try:
        base_cols = {}
        for c in cols:
            w = sheet.Columns(c).ColumnWidth
            if isinstance(w, bool) or not isinstance(w, (int, float)) or w <= 0:
                raise TypeError("invalid ColumnWidth")
            base_cols[str(c)] = round(float(w), 2)
    except Exception:
        raise ValueError(autofix.COLUMN_ONLY_REASON) from None
    autofix.validate_print_targets(sheet, items, action="未调整列宽")
    step = max(0.5, round(STEP_COL_WIDTH / len(cols), 2))
    margin = max(0.3, round(1.0 / len(cols), 2))
    pdf = session.temp_pdf()
    try:
        for attempt in range(1, autofix.MAX_ATTEMPTS + 1):
            target_cols = {str(c): min(MAX_COL_WIDTH, round(base_cols[str(c)] + step * attempt, 2)) for c in cols}
            for c in cols:
                sheet.Columns(c).ColumnWidth = target_cols[str(c)]
            try:
                disp = str(sheet.Range(item["cell_addr"]).MergeArea.Cells(1, 1).Text or "").strip()
            except Exception:
                disp = ""
            if len(disp) >= 3 and set(disp) == {"#"} and attempt < autofix.MAX_ATTEMPTS:
                continue  # 仍是 ###，继续加宽，不用每次都导出 PDF
            target_cols = {str(c): min(MAX_COL_WIDTH, round(target_cols[str(c)] + margin, 2)) for c in cols}
            for c in cols:
                sheet.Columns(c).ColumnWidth = target_cols[str(c)]
            started = time.perf_counter()
            wf._export_sheet_pdf(sheet, pdf)
            timing["exports"] = timing.get("exports", 0) + 1
            timing["export"] = timing.get("export", 0.0) + time.perf_counter() - started
            started = time.perf_counter()
            cells = snapshot_cells(extract_sheet_snapshot(sheet))
            probe = CellProbe(sheet)
            ok, shot, _ = autofix._evaluate(pdf, items, cells, probe.text, probe.printed)
            timing["evaluate"] = timing.get("evaluate", 0.0) + time.perf_counter() - started
            if ok:
                chosen_pdf = session.chosen_pdf()
                shutil.copyfile(pdf, chosen_pdf)
                return {"base_cols": base_cols, "target_cols": target_cols, "after": shot, "pdf": chosen_pdf}
    finally:
        for c in cols:
            sheet.Columns(c).ColumnWidth = base_cols[str(c)]
        pdf.unlink(missing_ok=True)
    raise ValueError(autofix.COLUMN_ONLY_REASON)


def _repair_column(
    excel: Any, session: Any, path: str, items: list[dict], card: dict,
    version: dict, timing: dict[str, float],
) -> Any:
    """### 列宽不足的单元格：在副本上试算加宽列，通过后走写入日志保存到原文件（可续跑、提交绿/黄时还原原列宽）"""
    started = time.perf_counter()
    chosen = _prepare_column(session, path, items, timing)
    timing["prepare"] = timing.get("prepare", 0.0) + time.perf_counter() - started
    item = items[0]
    with wf.STATE_LOCK:
        _upgrade_column_records()
        for col, base in chosen["base_cols"].items():
            record = _data()["rows"].get(_col_key(path, item["sheet_name"], int(col)))
            if record and record["saved"] and abs(record["applied"] - base) >= COL_WIDTH_TOLERANCE:
                raise ValueError("列宽已被其他操作修改，未覆盖人工修改，请核对后复检")
            if record and record["saved"] and any(
                (peer.get("marked") or peer.get("manual")) and not peer.get("dismissed")
                for cid in record["cards"] if (peer := _data()["cards"].get(cid))
            ):
                raise ValueError("该列已有人工审核意见，未再次调整列宽，请先提交审核")
        fix_images.forget(card, "after")
        _save_shot(card, "after", chosen["after"])
        rows = []
        col_ids = []
        for col, base in chosen["base_cols"].items():
            rid = _col_key(path, item["sheet_name"], int(col))
            record = _data()["rows"].get(rid)
            if not record or not record["saved"]:
                record = _data()["rows"][rid] = {
                    "id": rid, "kind": "col", "path": path, "sheet": item["sheet_name"],
                    "row": int(item["excel_row"]), "col": int(col), "base": base,
                    "applied": base, "saved": False, "cards": [], "wrap": [],
                }
            if card["id"] not in record["cards"]:
                record["cards"].append(card["id"])
            col_ids.append(rid)
            rows.append({"id": rid, "kind": "col", "sheet": item["sheet_name"], "row": int(item["excel_row"]),
                         "expected": base, "target": chosen["target_cols"][col],
                         "cols_expected": {col: base}, "cols_target": {col: chosen["target_cols"][col]}})
        card["col_ids"] = col_ids
        wf.STATE["job"]["pending_adjustment"] = {
            "action": "fix", "path": path, "version": version, "cards": [card["id"]],
            "rows": rows,
        }
        wf._persist_state()
    started = time.perf_counter()
    _finish_pending(excel)
    timing["write"] = timing.get("write", 0.0) + time.perf_counter() - started

    def widen(sheet: Any) -> None:
        for col_str, w in chosen["target_cols"].items():
            sheet.Columns(int(col_str)).ColumnWidth = float(w)
    return session.written_with(path, version, item["sheet_name"], widen, chosen["pdf"])


def _prepare_break(session: Any, path: str, item: dict, timing: dict[str, float]) -> dict:
    """在副本上合并区域第一行上方插手动分页，核对整块不再被切开，导出一份截图后把副本改回去。
    缩放到 N 页高（Excel 忽略手动分页）、合并区域比一页还高、上方已有手动分页的都不改，转人工。"""
    block = item.get("block") or {}
    top, bottom = int(block.get("top") or 0), int(block.get("bottom") or 0)
    if not top or bottom <= top:
        raise ValueError("合并区域信息缺失，请重新扫描")
    sheet = session.copy_of(path).Sheets(item["sheet_name"])
    if page_breaks.fit_to_pages_tall(sheet):
        raise ValueError(page_breaks.FIT_REASON)
    row = sheet.Rows(top)
    if int(row.PageBreak) == page_breaks.XL_PAGEBREAK_MANUAL:
        raise ValueError(page_breaks.MANUAL_REASON)
    pdf = session.chosen_pdf()
    row.PageBreak = page_breaks.XL_PAGEBREAK_MANUAL
    try:
        breaks = page_breaks.read_breaks(sheet)
        if breaks is None:
            raise ValueError("读不到 Excel 的分页信息（可能没有设置打印机），未修改，请人工处理")
        if not any(r == top and manual for r, manual in breaks["h"]):
            raise ValueError("插入的分页没有生效（打印设置不支持手动分页），未修改，请人工处理")
        if page_breaks.still_cut(breaks, block):
            raise ValueError(page_breaks.TALL_REASON)
        started = time.perf_counter()
        wf._export_sheet_pdf(sheet, pdf)
        timing["exports"] = timing.get("exports", 0) + 1
        timing["export"] = timing.get("export", 0.0) + time.perf_counter() - started
        with fitz.open(pdf) as doc:
            after = page_breaks.shot_around(sheet, breaks, doc, top, int(block.get("left") or 1), item["text"])[0]
    finally:
        row.PageBreak = page_breaks.XL_PAGEBREAK_NONE
    return {"after": after, "pdf": pdf}


def _repair_break(
    excel: Any, session: Any, path: str, item: dict, cards: list[dict], rid: str,
    version: dict, timing: dict[str, float],
) -> Any:
    """分页切开的合并区域：插分页，走写入日志保存到原文件（可续跑、提交绿 / 黄时可还原）。返回和原文件一致的副本"""
    started = time.perf_counter()
    chosen = _prepare_break(session, path, item, timing)
    timing["prepare"] = time.perf_counter() - started
    top = int(item["block"]["top"])
    with wf.STATE_LOCK:
        for card in cards:
            fix_images.forget(card, "after")
            _save_shot(card, "after", chosen["after"])
        _data()["rows"][rid] = {"id": rid, "kind": "break", "path": path, "sheet": item["sheet_name"], "row": top,
                                "base": 0.0, "applied": 1.0, "saved": False, "wrap": []}
        wf.STATE["job"]["pending_adjustment"] = {
            "action": "fix", "path": path, "version": version, "cards": [c["id"] for c in cards],
            "rows": [{"id": rid, "kind": "break", "sheet": item["sheet_name"], "row": top,
                      "expected": 0.0, "target": 1.0}],
        }
        wf._persist_state()
    started = time.perf_counter()
    _finish_pending(excel)
    timing["write"] = time.perf_counter() - started

    def insert(sheet: Any) -> None:
        sheet.Rows(top).PageBreak = page_breaks.XL_PAGEBREAK_MANUAL
    return session.written_with(path, version, item["sheet_name"], insert, chosen["pdf"])


class _RowDone(Exception):
    """这一行没有需要自动加高的格子了（如只有 ### 列宽问题），直接算处理过"""


def _restart_excel(session: Any) -> Any:
    """后台 Excel 可能已经崩溃或卡住：关掉整个会话（含副本），重新开一个"""
    session.close()
    return session.app()


def _repair_file(excel: Any, session: Any, path: str) -> None:
    job = wf.STATE["job"]
    attempted = job.setdefault("attempted_rows", [])
    copy_checked: set[str] = set()  # 用副本重新检测过的表：最后要用原文件完整复查
    while True:
        if _stopping(job):
            return
        with wf.STATE_LOCK:
            pending = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                       and it["status"] not in wf.RESOLVED_STATUSES]
            manual_rows = {_rid(c) for c in _data()["cards"].values()
                           if c["wb_path"] == path and (c.get("manual") or c.get("marked"))
                           and not c.get("dismissed") and not _column_card(c)}
            manual_cells = {_cell(c) for c in _data()["cards"].values()
                            if c["wb_path"] == path and (c.get("manual") or c.get("marked"))
                            and not c.get("dismissed") and _column_card(c)}
            manual_rows.update(r["id"] for r in _data()["rows"].values()
                               if r["path"] == path and r["saved"]
                               and any(not c.get("dismissed") and c["wb_path"] == path and _rid(c) == r["id"]
                                       for c in _data()["cards"].values()))
            candidates = [it for it in pending if _rid(it) not in attempted and _rid(it) not in manual_rows
                          and _cell(it) not in manual_cells]
            sheets = sorted(copy_checked) if not candidates else []
            copy_checked.difference_update(sheets)
        if sheets:
            # 副本和原文件应完全一致；兜底用原文件重新导出检测一遍，复查结果为准
            session.drop_copy(path)
            started = time.perf_counter()
            _refresh(excel, path, sheets, session=session, reuse=False)
            wf.log_msg(f"⏱ {Path(path).name} 用原文件复查 {len(sheets)} 张表：{time.perf_counter() - started:.1f} 秒")
            continue
        with wf.STATE_LOCK:
            if not candidates:
                # 之前修过的格子因后续分页变化又出现问题，也交给人工，不能循环自动加高。
                for group in _group_cells(pending):
                    card = _card_for(group)
                    if not card.get("dismissed") and not card.get("marked") and not card.get("manual"):
                        card.update(color="yellow", verified=False, waiting=False, message="当前导出仍有问题，需人工处理")
                wf._persist_state()
                return
            # 从上往下处理；同一行先插分页（合并区域整块挪到下一页后，里面的格子常常就不缺字了）
            first = min(candidates, key=lambda it: (it["sheet_name"], int(it["excel_row"]), not _is_break(it), it["seq"]))
            rid = _rid(first)
            items = [it for it in pending if _rid(it) == rid and _cell(it) not in manual_cells]
            cards = [_card_for(group) for group in _group_cells(items)]
            wf._persist_state()
        _check_file(path)
        version = wf._file_version(Path(path))
        debug_log.begin(f"行 {first['sheet_name']} 第 {first['excel_row']} 行", 2)
        debug_log.event("🛠 处理行", 位置=f"{first['sheet_name']} 第 {first['excel_row']} 行",
                        格子=[f"{it.get('cell_addr')}:{it['kind']}/{it.get('hidden_reason', '')}" for it in items],
                        已试行数=len(attempted), 剩余候选=len(candidates))

        def remember_attempt(index: int, height: float, shots: dict, wrapped: list[str] = ()) -> None:
            session.touch()
            with wf.STATE_LOCK:
                for card in cards:
                    if index == 0:
                        _save_shot(card, "before", shots.get(card["id"], ""))
                        card["base_height"] = height
                    # 尚未产生调整图就出错时，原样也是最后一次实际生成的预览。
                    # 只给卡片留参考图，不写入问题项的修复报告，也不写原 Excel。
                    fix_images.forget(card, "after")
                    _save_shot(card, "after", shots.get(card["id"], ""))
                    card.update(after_kind="attempt", attempt_height=height,
                                attempt_wrap=[a for a in wrapped if a == card.get("cell_addr")])
                wf._persist_state()

        timing: dict[str, float] = {}
        row_started = time.perf_counter()
        workbook = None  # 和原文件一致的副本：重新检测时直接用，不再打开原文件
        try:
            enabled_types = _active_fix_types(job)
            if _is_break(first):
                unchecked = _unchecked_fix_reason(cards[0], items, enabled_types)
                if unchecked:
                    with wf.STATE_LOCK:
                        for card in cards:
                            card.update(color="yellow", verified=False, manual=True, waiting=False,
                                        handed_over=True, after_kind="handover", message=unchecked)
                        wf.log_msg(f"⏭️ {first['wb_name']} / {first['sheet_name']} {first.get('cell_addr', '')}：{unchecked}")
                        wf._persist_state()
                    _remember_handover(path)
                    raise _RowDone
                workbook = _repair_break(excel, session, path, items[0], cards, rid, version, timing)
                raise _RowDone
            column_only = [it for it in items if autofix.column_width_only(it)]
            if column_only:
                for group in _group_cells(column_only):
                    card = _card_for(group)
                    unchecked = _unchecked_fix_reason(card, group, enabled_types)
                    if unchecked:
                        with wf.STATE_LOCK:
                            card.update(color="yellow", verified=False, manual=True, waiting=False,
                                        handed_over=True, after_kind="handover", message=unchecked)
                            wf.log_msg(f"⏭️ {first['wb_name']} / {first['sheet_name']} 行{first['excel_row']}：{unchecked}")
                            wf._persist_state()
                        _remember_handover(path)
                        continue
                    try:
                        workbook = _repair_column(excel, session, path, group, card, version, timing)
                        version = wf._file_version(Path(path))
                    except ValueError as err:
                        if job.get("pending_adjustment"):
                            raise
                        with wf.STATE_LOCK:
                            card.update(color="yellow", verified=False, manual=True, waiting=False, message=str(err))
                            wf.log_msg(f"⚠️ {first['wb_name']} / {first['sheet_name']} 行{first['excel_row']}：{err}")
                            wf._persist_state()
                with wf.STATE_LOCK:
                    done = {_cell(it) for it in column_only}
                    items = [it for it in items if _cell(it) not in done]
                    cards = [c for c in cards if _cell(c) not in done]
                    wf._persist_state()
            if not items:
                raise _RowDone
            reason = autofix.review_only_reason(items)
            if reason:
                for card in cards:
                    card["manual"] = True  # 同行共用行高；该行也不能因另一格而自动加高。
                raise ValueError(reason)
            unchecked_cells: set[tuple[str, str, str]] = set()
            with wf.STATE_LOCK:
                for group in _group_cells(items):
                    card = _card_for(group)
                    unchecked = _unchecked_fix_reason(card, group, enabled_types)
                    if unchecked:
                        card.update(color="yellow", verified=False, manual=True, waiting=False,
                                    handed_over=True, after_kind="handover", message=unchecked)
                        wf.log_msg(f"⏭️ {first['wb_name']} / {first['sheet_name']} 行{first['excel_row']}：{unchecked}")
                        unchecked_cells.add(_cell(group[0]))
                if unchecked_cells:
                    items = [it for it in items if _cell(it) not in unchecked_cells]
                    cards = [c for c in cards if _cell(c) not in unchecked_cells]
                    wf._persist_state()
            if unchecked_cells:
                _remember_handover(path)
            if not items:
                raise _RowDone
            items = [it for it in items if autofix.item_error_type(it) in enabled_types]
            if any(it.get("location_confirmed") is False or not it.get("cell_addr") for it in items):
                raise ValueError("单元格定位未确认，请人工核对")
            base, chosen, before = _prepare_row(session, path, items, on_attempt=remember_attempt, timing=timing)
            timing["prepare"] = time.perf_counter() - row_started
            with wf.STATE_LOCK:
                for card in cards:
                    _save_shot(card, "before", before.get(card["id"], ""))
                if chosen is None:
                    raise ValueError("自动试修未通过，原文件行高未修改；最后一次尝试未采用，请人工处理或标绿忽略")
                for card in cards:
                    _save_shot(card, "after", chosen["shots"].get(card["id"], ""))
                unchanged = abs(chosen["height"] - base) < ROW_HEIGHT_TOLERANCE
                if unchanged:
                    # 原样导出就通过：不写文件，也不算一处修复（报告里不列）
                    for card in cards:
                        card.update(color="blue", verified=True, after_kind="unchanged",
                                    message="原样导出已通过检查，未修改行高，等待审核")
                        for item in _items_at(card):
                            item.update(status="accepted", batch_managed=True, triage_keep=False,
                                        review_reason="原样导出已通过检查，未修改")
                else:
                    _data()["rows"][rid] = {"id": rid, "path": path, "sheet": first["sheet_name"],
                                            "row": int(first["excel_row"]), "base": base,
                                            "applied": chosen["height"], "saved": False, "wrap": []}
                    job["pending_adjustment"] = {
                        "action": "fix", "path": path, "version": version, "cards": [c["id"] for c in cards],
                        "rows": [{"id": rid, "sheet": first["sheet_name"], "row": int(first["excel_row"]),
                                  "expected": base, "target": chosen["height"],
                                  "wrap": {addr: True for addr in chosen.get("wrap", [])}}],
                    }
                wf._persist_state()
            if unchanged:
                workbook = session.copy_of(path)
            else:
                started = time.perf_counter()
                _finish_pending(excel)
                timing["write"] = time.perf_counter() - started
                workbook = session.written(path, version, first["sheet_name"], int(first["excel_row"]),
                                           chosen["height"], chosen.get("wrap", []), chosen.get("pdf"))
        except _RowDone:
            pass
        except Exception as err:
            # 有写入日志时不能吞掉异常：可能原文件已经保存，必须按日志续跑。
            if job.get("pending_adjustment"):
                raise
            if getattr(getattr(session, "watchdog", None), "timed_out", False) is True:
                raise TimeoutError("后台 Excel 无进展超时，已跳过该文件，待后续重试") from err
            if not isinstance(err, ValueError):
                # 程序出错（Excel 崩溃、一时忙不过来等），不是检测不过：重启 Excel 重试这一行，几次都不行才转人工
                with wf.STATE_LOCK:
                    errors = job.setdefault("row_errors", {})
                    errors[rid] = errors.get(rid, 0) + 1
                    retry = errors[rid] <= autofix.MAX_ERROR_RETRIES
                    wf._persist_state()
                if retry:
                    wf.log_msg(f"⚠️ {first['wb_name']} / {first['sheet_name']} 行{first['excel_row']} 出错，"
                               f"重启 Excel 后重试（第 {errors[rid]} 次）：{err}")
                    excel = _restart_excel(session)
                    continue
            with wf.STATE_LOCK:
                for card in cards:
                    card.update(color="yellow", verified=False, manual=True, waiting=False, message=str(err))
                wf.log_msg(f"⚠️ {first['wb_name']} / {first['sheet_name']} 行{first['excel_row']}：{err}")
        with wf.STATE_LOCK:
            if rid not in attempted:
                attempted.append(rid)
            for card in cards:
                card["waiting"] = False
            wf._persist_state()
        if workbook is None:
            session.drop_copy(path)
        elif "write" in timing:
            copy_checked.add(first["sheet_name"])
        # 包括成功写入后原来没被检出的新问题，始终以最新整表 PDF 为准。
        started = time.perf_counter()
        checking: dict[str, float] = {}
        _refresh(excel, path, [first["sheet_name"]], session=session, workbook=workbook, timing=checking)
        if "prepare" in timing:
            wf.log_msg(_timing_line(first, time.perf_counter() - row_started, timing,
                                    time.perf_counter() - started, checking))


def _timing_line(first: dict, total: float, timing: dict, check: float, checking: dict) -> str:
    """每行耗时，方便实测时看出慢在哪一步"""
    parts = [f"试算导出 {int(timing.get('exports', 0))} 次 {timing.get('export', 0):.1f} 秒",
             f"比对 {timing.get('evaluate', 0):.1f} 秒"]
    if "write" in timing:
        parts.append(f"写入保存 {timing['write']:.1f} 秒")
    detail = [f"导出 {int(checking.get('exports', 0))} 次 {checking.get('export', 0):.1f} 秒",
              f"检测 {checking.get('detect', 0):.1f} 秒"]
    if checking.get("open", 0) >= .05:
        detail.insert(0, f"打开 {checking['open']:.1f} 秒")
    parts.append(f"重新检测 {check:.1f} 秒（{'，'.join(detail)}）")
    return f"⏱ {first['wb_name']} / {first['sheet_name']} 行{first['excel_row']}：共 {total:.1f} 秒；" + "，".join(parts)


COLOR_NAMES = {"blue": "蓝", "green": "绿", "yellow": "黄"}


def _linked_row(rid: str, peers: list[dict]) -> bool:
    """同一行的卡共用一次行高修改：已写入原文件、或有人工认可要写入时，整行只能一起保留或一起还原"""
    row = _data()["rows"].get(rid)
    return bool((row and row["saved"]) or any(_adopting(c) for c in peers))


def _review_groups(cards: list[dict]) -> list[tuple[list[dict], str]]:
    """审核联动范围：行高按行、列宽按共享列；重叠合并区域通过共同的卡片连成一组。"""
    parents = {c["id"]: c["id"] for c in cards}

    def root(cid: str) -> str:
        while parents[cid] != cid:
            parents[cid] = parents[parents[cid]]
            cid = parents[cid]
        return cid

    resources: dict[tuple[str, str], list[dict]] = {}
    for card in cards:
        if _column_card(card):
            for row in _column_records(card):
                if row["saved"]:
                    resources.setdefault(("col", row["id"]), []).append(card)
        else:
            resources.setdefault(("row", _rid(card)), []).append(card)
    for peers in resources.values():
        first = root(peers[0]["id"])
        for card in peers[1:]:
            parents[root(card["id"])] = first
    groups: dict[str, list[dict]] = {}
    scopes: dict[str, str] = {}
    for card in cards:
        groups.setdefault(root(card["id"]), []).append(card)
    for (kind, rid), peers in resources.items():
        if kind == "col" or _linked_row(rid, peers):
            scopes[root(peers[0]["id"])] = "列宽" if kind == "col" else "行高"
    return [(peers, scopes.get(cid, "")) for cid, peers in groups.items()]


def commit_selection(ids: list[str], color: str | None = None) -> tuple[list[str], list[str]]:
    """按共享行高 / 列宽补齐提交范围；按颜色提交时跳过有不同意见的联动组。"""
    _upgrade_column_records()
    data = _data()
    selected = set(ids)
    ready = [c for c in data["cards"].values() if not c.get("dismissed") and not _waiting(c)]
    chosen: list[str] = []
    skipped: list[str] = []
    for peers, scope in _review_groups(ready):
        if not any(c["id"] in selected for c in peers):
            continue
        if not color:
            chosen.extend(c["id"] for c in peers)
        elif scope and any(c["color"] != color for c in peers):
            skipped.extend(c["id"] for c in peers if c["color"] == color)
        else:
            chosen.extend(c["id"] for c in peers if c["color"] == color)
    return chosen, skipped


def _commit_review(excel: Any) -> None:
    job = wf.STATE["job"]
    data = _data()
    if "commit_card_ids" not in job["options"]:
        job["options"]["commit_card_ids"], skipped = commit_selection(job["options"]["card_ids"],
                                                                      job["options"].get("color"))
        if skipped:
            wf.log_msg(f"⏭ {len(skipped)} 张{COLOR_NAMES[job['options']['color']]}卡与其他颜色共用行高或列宽，这次不提交")
    # 提交计划固定下来；即使保存还原后、复检前退出，已收起的绿/蓝卡也仍会完成该文件的复检。
    job["options"]["card_ids"] = list(job["options"]["commit_card_ids"])
    cards = [data["cards"][cid] for cid in job["options"]["commit_card_ids"]]
    by_row: dict[str, list[dict]] = {}
    for card in cards:
        if not _column_card(card):
            by_row.setdefault(_rid(card), []).append(card)
    # 只在仍有自动修改需要还原时检查同行冲突；已经忽略的格子不因后来复检重新入队。
    # 人工认可的加高也改整行行高：同行颜色不一致同样整行转黄
    for peers, scope in _review_groups(cards):
        if (scope and len({c["color"] for c in peers}) > 1
                and peers[0]["wb_path"] not in job.get("deferred_adjustments", {})):
            for card in peers:
                card.update(color="yellow", conflict=True, column_conflict=scope == "列宽",
                            message=("同列存在不同处理意见，相关列还原后交给人工" if scope == "列宽" else
                                     "同行存在不同处理意见，整行还原后交给人工"))
    wf._persist_state()
    debug_log.event("📨 提交审核", 卡片=len(cards), **{color: sum(c["color"] == color for c in cards)
                                                      for color in ("blue", "green", "yellow")})
    for path in dict.fromkeys(c["wb_path"] for c in cards):
        if path in job["completed"] or _unsettled(job, path):
            continue
        debug_log.begin(f"提交 {Path(path).name}")
        rows = []
        reverting = []
        columns = {r["id"]: r for c in cards for r in _column_records(c) if r["path"] == path}
        for rid, record in columns.items():
            peers = [c for c in cards if rid in c.get("col_ids", [])]
            if not record["saved"] or all(c["color"] == "blue" for c in peers):
                continue
            col = str(record["col"])
            rows.append({"id": rid, "kind": "col", "sheet": record["sheet"], "row": record["row"],
                         "expected": record["applied"], "target": record["base"],
                         "cols_expected": {col: record["applied"]}, "cols_target": {col: record["base"]}})
            reverting.extend(c["id"] for c in peers)
        for rid, peers in by_row.items():
            row = data["rows"].get(rid)
            if not row or row["path"] != path or not row["saved"] or all(c["color"] == "blue" for c in peers):
                continue
            rows.append({"id": rid, "kind": row.get("kind"), "sheet": row["sheet"], "row": row["row"],
                         "expected": row["applied"], "target": row["base"],
                         "wrap": {addr: False for addr in row.get("wrap", [])}})
            reverting.extend(c["id"] for c in peers)
        try:
            if rows:
                job["options"].setdefault("refresh_sheets", {})[path] = sorted({r["sheet"] for r in rows})
                job["pending_adjustment"] = {"action": "restore", "path": path,
                                             "version": wf._file_version(Path(path)), "rows": rows,
                                             "cards": list(dict.fromkeys(reverting))}
                wf._persist_state()
                _finish_pending(excel)
            _adopt_rows(excel, path, by_row)
        except Exception as err:
            if not job.get("pending_adjustment"):
                raise
            # 这份文件的行高核对不通过（多半在 Excel 里改过）：留着写入日志，先提交其他文件
            _park_pending(err)
            wf.log_msg(f"❌ {Path(path).name} 写入未结算，已跳过，继续提交其他文件：{err}")
            continue
        with wf.STATE_LOCK:
            for card in [c for c in cards if c["wb_path"] == path]:
                card["finalized"] = True
                # 标蓝但检测器仍报的格子（人工认可）也要记规则，否则重新检测又出卡
                accepted_by_eye = card["color"] == "blue" and (card.get("adopted") or not card.get("verified"))
                if card["color"] == "green" or accepted_by_eye:
                    rule = {"cell": list(_cell(card)), "signatures": card["signatures"]}
                    if card.get("adopted"):
                        rule["fixed"] = {"row_height": card["adopted_height"], "batch_managed": True,
                                         "wrap_text": bool(card.get("attempt_wrap")),
                                         "fix_images": copy.deepcopy(card.get("fix_images", {})),
                                         "review_reason": "全自动批处理：已修复并保存"}
                    if rule not in data["ignored"]:
                        data["ignored"].append(rule)
                card["dismissed"] = card["color"] != "yellow"
                card["manual"] = card["color"] == "yellow"
                if card["manual"]:
                    conflict = ("同列存在不同处理意见，相关列已还原；" if card.get("column_conflict") else
                                "同行存在不同处理意见，整行已还原；") if card.get("conflict") else ""
                    card.update(verified=False, message=conflict
                                + "自动修改已还原，请在 Excel 中修改并保存后复检")
                for item in _items_at(card):
                    item.update(batch_managed=True, triage_keep=False)
                    if accepted_by_eye:
                        fixed = next((r["fixed"] for r in data["ignored"] if r["cell"] == list(_cell(card))
                                      and r.get("fixed")), None) if card.get("adopted") else None
                        if fixed:
                            item.update(status="fixed", **copy.deepcopy(fixed))
                        else:
                            item.update(status="accepted", manual_confirmation={"accepted": True},
                                        review_reason=wf.MANUAL_CONFIRMED_REASON)
                    elif card["color"] == "green":
                        item.update(status="accepted", manual_confirmation={"accepted": True},
                                    review_reason=wf.MANUAL_CONFIRMED_REASON)
                    elif card["color"] == "yellow":
                        item.update(status="manual_review", review_reason=card["message"])
            wf._persist_state()
            wf._sync_problem_counts(path, "已提交审核标记，黄色等待人工处理")
        refresh_sheets = job["options"].get("refresh_sheets", {}).get(path)
        if refresh_sheets:
            # 还原也会改变分页；不能把因此重新出现问题的蓝色卡片直接放行。
            _refresh(excel, path, refresh_sheets)
            with wf.STATE_LOCK:
                unresolved = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                              and it["sheet_name"] in refresh_sheets
                              and it["status"] not in wf.RESOLVED_STATUSES]
                for group in _group_cells(unresolved):
                    card = _card_for(group)
                    if card["color"] == "blue" or not card.get("finalized"):
                        card.update(color="yellow", verified=False, dismissed=False, manual=True, waiting=False,
                                    finalized=False, message="还原后重新检测仍有问题，请提交黄色标记后人工处理")
                wf._persist_state()
        with wf.STATE_LOCK:
            for rid, peers in by_row.items():
                row = data["rows"].get(rid)
                if row and row["path"] == path and all(c.get("dismissed") and c["color"] == "blue" for c in peers):
                    row["saved"] = False  # 已正式接受的历史修改不再属于可还原的待审核修改。
            for rid, record in columns.items():
                peers = [data["cards"][cid] for cid in record["cards"] if cid in data["cards"]
                         and rid in data["cards"][cid].get("col_ids", [])]
                if peers and all(c.get("dismissed") and c["color"] == "blue" for c in peers):
                    record["saved"] = False
            job["completed"].append(path)
            wf._persist_state()


def _adopt_rows(excel: Any, path: str, by_row: dict[str, list[dict]]) -> None:
    """标蓝的「最后一次尝试」：按那次尝试的行高写入原文件（同行都标蓝才写）。走写入日志，可续跑。"""
    job = wf.STATE["job"]
    data = _data()
    rows, adopting = [], []
    for rid, peers in by_row.items():
        mine = [c for c in peers if c["wb_path"] == path and _adopting(c)]
        if not mine or any(c["color"] != "blue" for c in peers):
            continue
        first = mine[0]
        target = max(c["attempt_height"] for c in mine)
        if abs(target - first["base_height"]) < ROW_HEIGHT_TOLERANCE:
            continue  # 原样就是这样：不用写，提交时按忽略处理
        data["rows"].setdefault(rid, {"id": rid, "path": path, "sheet": first["sheet_name"],
                                      "row": int(first["excel_row"]), "base": first["base_height"],
                                      "applied": first["base_height"], "saved": False})
        rows.append({"id": rid, "sheet": first["sheet_name"], "row": int(first["excel_row"]),
                     "expected": first["base_height"], "target": target,
                     "wrap": {addr: True for c in mine for addr in c.get("attempt_wrap", [])}})
        adopting.extend(c["id"] for c in mine)
    if not rows:
        return
    refresh = job["options"].setdefault("refresh_sheets", {})
    refresh[path] = sorted(set(refresh.get(path, [])) | {r["sheet"] for r in rows})
    job["pending_adjustment"] = {"action": "adopt", "path": path, "version": wf._file_version(Path(path)),
                                 "rows": rows, "cards": adopting}
    wf._persist_state()
    _finish_pending(excel)


def _recheck_file(excel: Any, path: str, cards: list[dict]) -> None:
    """人工修改后复检：不管改了哪里，整份文件所有要打印的表全部重新导出检测，不去找原来那一格。
    原来的黄卡全部翻篇：这次没再查出问题的移出列表；这次查出的问题（同一格的就是原来那张卡）
    直接作为人工待处理项显示，改好保存后再复检，直到查不出问题为止。"""
    _check_file(path)
    with wf.STATE_LOCK:
        # 整份本来就要重新检测：文件里其他已提交的黄色卡一起出结果，不用再一张张单独复检
        listed = {c["id"] for c in cards}
        cards = cards + [c for c in _data()["cards"].values()
                         if c["id"] not in listed and c["wb_path"] == path
                         and c["color"] == "yellow" and c.get("finalized") and not c.get("dismissed")]
        # 全自动还没轮到的卡（按对象记：工作表改名时卡片编号会跟着换）
        waiting = {id(c) for c in _data()["cards"].values()
                   if c["wb_path"] == path and not c.get("dismissed") and _waiting(c)}
    started = time.perf_counter()
    if _unchanged_since_check(path):
        # 上次整份检测之后文件没改过，重新导出结果必然一样：直接用上次的检测结果，准确度不变
        wf.log_msg(f"ℹ️ {Path(path).name}：上次检测后文件没有修改，沿用上次检测结果")
        with wf.STATE_LOCK:
            carry_item_policies(wf.STATE["items"])
    else:
        checking: dict[str, float] = {}
        book = _open_without_marks(excel, path)
        try:
            _refresh(excel, path, workbook=book, timing=checking)
        finally:
            if book is not None:
                book.Close(SaveChanges=False)  # 只在内存里还原过颜色，不保存；之后标色时统一打开保存一次
        wf.log_msg(f"⏱ {Path(path).name} 整份复检：{time.perf_counter() - started:.1f} 秒"
                   f"（打开 {checking.get('open', 0):.1f} 秒，导出 {int(checking.get('exports', 0))} 次 "
                   f"{checking.get('export', 0):.1f} 秒，检测 {checking.get('detect', 0):.1f} 秒）")
    with wf.STATE_LOCK:
        unresolved = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                      and it["status"] not in wf.RESOLVED_STATUSES]
        problem_cells = {_cell(it) for it in unresolved}
        passed = 0
        for card in cards:
            if _cell(card) in problem_cells:
                continue
            passed += 1
            # 这一格这次没查出问题：翻篇，移出列表。报告里的修复后截图用刚导出的 PDF 按文字重新截，截不到就不放
            prior = _items_at(card)
            sources = [p for p in wf.STATE["pages"].values()
                       if p.get("wb_path") == path and p.get("sheet_name") == card["sheet_name"]] + prior
            fix_images.forget(card, "after")
            _save_shot(card, "after", _latest_shot(card, sources))
            after = card.get("fix_images", {}).get("after")
            for item in prior:
                if after:
                    item.setdefault("fix_images", {})["after"] = after
                else:
                    fix_images.forget(item, "after")
            card.update(color="blue", verified=True, finalized=True, dismissed=True, waiting=False,
                        after_kind="fixed", message="人工修改后复检通过")
        found = 0
        for group in _group_cells(unresolved):
            card = _card_for(group)
            if id(card) in waiting or (card.get("marked") and card["color"] != "yellow" and not card.get("finalized")):
                continue  # 全自动还没轮到的，或你已标蓝 / 绿还没提交的，保留原样
            found += 1
            message = "复检查出的问题：请修改并保存后再复检"
            if _waiting(card):
                card["handed_over"] = True  # 这次才查出来的格子，自动还没试过：「转回自动」时可以交给自动
            card.update(color="yellow", verified=False, manual=True, finalized=True, waiting=False,
                        dismissed=False, message=message)
            for item in group:
                item.update(status="manual_review", batch_managed=True, triage_keep=False, review_reason=message)
        wf._persist_state()
        wf._sync_problem_counts(path, "人工复检完成" + (f"，还有 {found} 处问题" if found else "，没有查出问题"))
    wf.log_msg(f"🔁 {Path(path).name}：原来 {len(cards)} 处翻篇（{passed} 处已不再有问题），"
               + (f"这次查出 {found} 处问题，改好后再复检" if found else "这次没有查出问题，全部通过"))


SINGLE_PASS_MESSAGE = "人工修改后单项复检通过（只查了这一格）"
SINGLE_FAIL_MESSAGE = "单项复检仍有问题：请修改并保存后再复检"


def _sheet_now(book: Any, path: str, name: str) -> Any:
    """按表名找工作表；改过名的按 Excel 代码名认出来。找不到返回 None"""
    sheet = _find_sheet(book, name)
    if sheet is not None:
        return sheet
    with wf.STATE_LOCK:
        code = (wf.STATE.get("sheet_codes", {}).get(path) or {}).get(name)
    for candidate in book.Worksheets if code else []:
        try:
            if candidate.CodeName == code:
                return candidate
        except Exception:
            continue
    return None


def _current_area(book: Any, path: str, card: dict) -> tuple[Any, tuple[int, int, int, int]]:
    """单项复检：这一格现在在哪。标色规则跟着格子走（插删行列后也找得到），旧方式按隐藏名称，
    都没有就按原来的表名和地址。返回 (工作表, (首行, 首列, 末行, 末列))，合并格按整块。
    要在去掉标记之前调用：去掉后规则就找不到了。"""
    with wf.STATE_LOCK:
        mark = copy.deepcopy(_marks().get(card["id"]))
    target = None
    if mark and _is_rule(mark):
        home = _sheet_now(book, path, card["sheet_name"])
        for sheet in [home] + [s for s in book.Worksheets if home is None or s.Name != home.Name]:
            if sheet is None:
                continue
            rules = sheet.Cells.FormatConditions
            for i in range(1, int(rules.Count) + 1):
                rule = rules.Item(i)
                if _rule_tag(rule) == mark["tag"]:
                    target = rule.AppliesTo.Areas(1).Cells(1, 1)
                    break
            if target is not None:
                break
    elif mark:
        found = _locate_mark(book, mark)
        if found is None:
            raise ValueError("这一格已被删除，请点「复检这个文件」整份重查")
        if found[0] is not None:
            target = _target(found[0], mark["addr"], found[1])
    if target is None:
        sheet = _sheet_now(book, path, card["sheet_name"])
        if sheet is None:
            raise ValueError(f"找不到工作表「{card['sheet_name']}」（可能被删除、隐藏或改名），请点「复检这个文件」整份重查")
        target = sheet.Range(card["cell_addr"])
    area = target.MergeArea
    top, left = int(area.Row), int(area.Column)
    return area.Worksheet, (top, left, top + int(area.Rows.Count) - 1, left + int(area.Columns.Count) - 1)


def _page_span(sheet: Any, breaks: dict | None, box: tuple[int, int, int, int]
               ) -> tuple[tuple[int, int], tuple[int, int, int, int]] | None:
    """这一格打印在哪几页：((第一页, 最后一页)（从 1 起）, 这几页上的格子 (首行, 末行, 首列, 末列))。
    分页读不到、页数和分页对不上（空白页被跳过、多块打印区域等）返回 None，由调用方导出整张表"""
    if breaks is None:
        return None
    try:
        pages = int(sheet.PageSetup.Pages.Count)
    except Exception:
        return None
    top, left, bottom, right = box
    first = page_breaks.page_index(sheet, breaks, top, left, pages)
    last = page_breaks.page_index(sheet, breaks, bottom, right, pages)
    if first is None or last is None:
        return None
    rows, cols = [r for r, _ in breaks["h"]], list(breaks["v"])
    area = (max([r for r in rows if r <= top], default=1), min([r for r in rows if r > bottom], default=1 << 20) - 1,
            max([c for c in cols if c <= left], default=1), min([c for c in cols if c > right], default=1 << 14) - 1)
    return (min(first, last) + 1, max(first, last) + 1), area


def _cut_by_break(sheet: Any, breaks: dict | None, box: tuple[int, int, int, int]) -> bool:
    """多行合并格被自动分页切开（和整份检测的 find_cut_blocks 同一标准：手动分页、没字、不打印的不算）"""
    top, left, bottom, _ = box
    if breaks is None or bottom <= top or not any(top < r <= bottom and not manual for r, manual in breaks["h"]):
        return False
    try:
        text = str(sheet.Cells(top, left).MergeArea.Cells(1, 1).Text or "").strip()
    except Exception:
        return False
    return bool(normalize_text(text)) and CellProbe(sheet).printed(top, left) is not False


def _recheck_cell(excel: Any, path: str, card: dict) -> None:
    """单项复检：只看这一格修好没有，同文件其他黄卡、别处有没有出新问题都不管（整份重查用「复检这个文件」）。
    分页和导出整张表时一样，只导出这一格所在的那一页（算不准第几页就导出这一张表），只拿这一格的检测结果：
    通过就翻篇（变蓝移出列表、去掉标记色），没过照旧是黄色人工项，可以接着改、接着复检。"""
    _check_file(path)
    label = f"{card['sheet_name']}!{card['cell_addr']}"
    if _unchanged_since_check(path):
        # 上次整份检测之后文件没改过：这一格的结果就是卡片现在的样子
        wf.log_msg(f"ℹ️ {Path(path).name} {label}：上次检测后文件没有修改，结果不变，仍需人工处理")
        return
    with wf.STATE_LOCK:
        marks = [copy.deepcopy(m) for m in _marks().values() if m["path"] == path]
    started = time.perf_counter()
    timing: dict[str, Any] = {}
    pdf = Path(tempfile.gettempdir()) / f"~单项复检_{uuid4().hex}.pdf"
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True)
    try:
        timing["open"] = time.perf_counter() - started
        sheet, box = _current_area(book, path, card)
        if marks:
            _strip_marks(book, path, marks)  # 检测的是最后交出去的样子，不保存
        top, left, bottom, right = box
        now = f"{sheet.Name}!{_col_letters(left)}{top}"
        breaks = page_breaks.read_breaks(sheet)
        span = _page_span(sheet, breaks, box)
        clock = time.perf_counter()
        wf._export_sheet_pdf(sheet, pdf, span[0] if span else None)
        timing["export"] = time.perf_counter() - clock
        clock = time.perf_counter()
        _, found = wf._detect_sheet_results(sheet, Path(path), sheet.Name, pdf, 0,
                                            page_span=span[0] if span else None, area=span[1] if span else None)
        text = normalize_text(card.get("text", ""))

        def mine(item: dict) -> bool:
            row, col = _cell_rc(item.get("cell_addr") or "A1")
            if top <= row <= bottom and left <= col <= right:
                return True
            # 认不准是哪一格的切字，文字对得上就算这一格的，宁可多改一次也不放过
            clip = normalize_text(item.get("text", ""))
            return (item.get("kind") == "clip" and item.get("location_confirmed") is False
                    and bool(clip) and clip in text)

        problems = [it for it in found.values() if mine(it)]
        if span and _cut_by_break(sheet, breaks, box):
            problems.append({"kind": "pagebreak", "text": card.get("text", "")})
        timing["detect"] = time.perf_counter() - clock
        shot = "" if problems else _latest_shot(card, [{"pdf_path": str(pdf)}])
    finally:
        book.Close(SaveChanges=False)
        pdf.unlink(missing_ok=True)
    pages = f"第 {span[0][0]}" + (f"–{span[0][1]}" if span[0][1] != span[0][0] else "") + " 页" if span else "整张表"
    wf.log_msg(f"⏱ {Path(path).name} {label} 单项复检：{time.perf_counter() - started:.1f} 秒"
               f"（打开 {timing['open']:.1f} 秒，导出{pages} {timing['export']:.1f} 秒，检测 {timing['detect']:.1f} 秒）")
    debug_log.event("🎯 单项复检", 位置=label, 现在=now, 导出=pages, 问题=[p["kind"] for p in problems] or "-")
    moved = "" if now == label else f"（这一格现在在 {now}）"
    with wf.STATE_LOCK:
        if problems:
            card.update(color="yellow", verified=False, manual=True, finalized=True, waiting=False,
                        dismissed=False, message=SINGLE_FAIL_MESSAGE + moved)
            _remember_problems(card, problems)
            for item in _items_at(card):
                item.update(status="manual_review", batch_managed=True, triage_keep=False,
                            review_reason=card["message"])
        else:
            fix_images.forget(card, "after")
            _save_shot(card, "after", shot)
            after = card.get("fix_images", {}).get("after")
            for item in _items_at(card):
                item.update(status="fixed", batch_managed=True, triage_keep=False)
                if after:
                    item.setdefault("fix_images", {})["after"] = after
                else:
                    fix_images.forget(item, "after")
            card.update(color="blue", verified=True, finalized=True, dismissed=True, waiting=False,
                        after_kind="fixed", message=SINGLE_PASS_MESSAGE)
        wf._persist_state()
    wf._sync_problem_counts(path, "单项复检完成")
    wf.log_msg(f"🎯 {Path(path).name} {label}{moved}："
               + ("仍有问题，改好后再复检" if problems else "这一格已没有问题，通过"))


def _open_without_marks(excel: Any, path: str) -> Any:
    """复检用：只读打开原文件，在内存里把标记全部去掉（删掉我们加的条件格式规则；旧版本改过的底色还原），
    不保存，检测的就是最后交出去的样子。没标过返回 None，照常由检测自己打开。"""
    with wf.STATE_LOCK:
        marks = [copy.deepcopy(m) for m in _marks().values() if m["path"] == path]
    if not marks:
        return None
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True)
    try:
        _strip_marks(book, path, marks)
    except BaseException:
        book.Close(SaveChanges=False)
        raise
    return book


def _strip_marks(book: Any, path: str, marks: list[dict]) -> None:
    """在内存里去掉这份文件的标记（不保存）：删我们加的条件格式规则，旧版本改过的底色还原"""
    failed: list[str] = []
    if any(_is_rule(m) for m in marks):
        try:
            failed += list(_drop_rules(book, set())[2].values())
        except Exception as err:
            failed.append(str(err))
    for mark in marks:
        if _is_rule(mark):
            continue
        try:
            found = _locate_mark(book, mark)
            if found is not None and found[0] is not None:
                _restore_cells(found[0], mark, found[1])
        except Exception as err:
            # 一格还原不了（例如工作表被保护）不该让整份复检失败：照常检测，日志里说一声
            failed.append(f"{mark['sheet']}!{mark['addr']}（{err}）")
    if failed:
        wf.log_msg(f"⚠️ {Path(path).name}：{len(failed)} 处标记没能在检测前去掉，带着标记色检测："
                   + "；".join(failed[:5]) + ("……" if len(failed) > 5 else ""))


def _hand_over(path: str, *, switch: bool = False) -> None:
    """本文件检出的问题转人工（黄色、已提交），不试修、不改行高。
    只扫描时全部转；全自动中途「转人工」（switch）时只转还没轮到自动的格子，
    自动已经试过、没通过的照原样交给人工，以后「转回自动」也不再试。
    用户已经手动标成蓝 / 绿、还没提交的卡保留用户的选择。"""
    with wf.STATE_LOCK:
        pending = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                   and it["status"] not in wf.RESOLVED_STATUSES]
        for group in _group_cells(pending):
            card = _card_for(group)
            if card.get("marked") and card["color"] != "yellow" and not card.get("finalized"):
                card["waiting"] = False
                continue
            if switch and card.get("handed_over") and card.get("finalized"):
                continue  # 上次转人工时就交给人工、还没处理完的，原样留着
            if switch and not _waiting(card):
                row = _row_record(card)
                if card["color"] == "yellow" and not (row and row["saved"]):
                    # 没写过原文件，不用等提交还原：直接可以人工处理、复检
                    card.update(manual=True, finalized=True, handed_over=False)
                    for item in _items_at(card):
                        item.update(status="manual_review", batch_managed=True, triage_keep=False,
                                    review_reason=card["message"])
                continue
            if any(autofix.column_width_only(it) for it in group):
                message = autofix.COLUMN_ONLY_REASON
            elif not _can_highlight(card):
                message = "单元格定位未确认，未在 Excel 中标色，请按截图查找后人工处理"
            else:
                message = HANDOVER_MESSAGE if switch else SCAN_ONLY_MESSAGE
            card.update(color="yellow", verified=False, manual=True, finalized=True, waiting=False,
                        dismissed=False, after_kind="handover" if switch else "scan_only", message=message,
                        handed_over=True)
            for item in _items_at(card):
                item.update(status="manual_review", batch_managed=True, triage_keep=False, review_reason=message)
        wf._persist_state()
    wf._sync_problem_counts(path, "已转人工处理" if switch else "只扫描完成，全部转人工处理")


def _row_key(sheet: str, row: Any) -> str:
    return f"{sheet}\0{int(row)}"


def _remember_handover(path: str) -> None:
    """转人工那一刻记下文件版本和每个问题行的行高：「转回自动」时据此判断人改过文件、动过哪些行"""
    with wf.STATE_LOCK:
        rows = {_row_key(it["sheet_name"], it["excel_row"]): float(it["row_height"])
                for it in wf.STATE["items"].values()
                if it["wb_path"] == path and isinstance(it.get("row_height"), (int, float))
                and it.get("excel_row")}
        try:
            version = wf._file_version(Path(path))
        except OSError:
            version = None
        _data().setdefault("handover", {})[path] = {"version": version, "rows": rows}
        wf._persist_state()


def _touched(item: dict, rows: dict[str, float]) -> bool:
    """这一行的行高和转人工时不一样：人动过。转人工时没有记录的行（插行后的新位置等）算没动过"""
    before = rows.get(_row_key(item["sheet_name"], item["excel_row"])) if item.get("excel_row") else None
    now = item.get("row_height")
    return before is not None and isinstance(now, (int, float)) and abs(now - before) >= ROW_HEIGHT_TOLERANCE


def _back_to_auto(excel: Any, session: Any, path: str, *, picked: bool = False) -> None:
    """「转回自动」时每份文件先做的事：转人工后文件改过就整份重扫（规则同复检，原来的黄卡翻篇）；
    还没被自动试过、所在行也没被人动过的问题交回自动队列，其余留给人工；最后还原 Excel 里的标记颜色，
    自动处理的截图和检测都按原来的颜色来。
    picked：「🤖 自动调整」只交回你选的格子（卡片上 auto_pick）和同一行的格子，其余黄卡和它们的标色都不动。"""
    unsaved = close_saved_workbooks([Path(path)])
    if Path(path) in unsaved:
        raise ValueError(f"{Path(path).name}：{unsaved[Path(path)]}，请保存后继续")
    _check_file(path)
    with wf.STATE_LOCK:
        hand = _data().get("handover", {}).get(path) or {}
        unchanged = hand.get("version") is not None and hand["version"] == wf._file_version(Path(path))
        cards = [c for c in _data()["cards"].values() if c["wb_path"] == path and not c.get("dismissed")
                 and c["color"] == "yellow" and c.get("finalized")]
    if unchanged:
        wf.log_msg(f"ℹ️ {Path(path).name}：转人工后没有修改，沿用上次检测结果")
    else:
        _recheck_file(excel, path, cards)
    with wf.STATE_LOCK:
        rows = hand.get("rows", {})
        enabled_types = _active_fix_types(wf.STATE.get("job"))
        unresolved = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                      and it["status"] not in wf.RESOLVED_STATUSES]
        # 复检可能查出同一行新的格子：一行共用行高，要一起交给自动，不然这一行会被当成人工行跳过
        picked_rows = {_rid(c) for c in _data()["cards"].values() if c["wb_path"] == path and c.get("auto_pick")}
        back = kept = 0
        for group in _group_cells(unresolved):
            card = _card_for(group)
            chosen = (card.get("auto_pick") or _rid(card) in picked_rows) if picked else card.get("handed_over")
            if not chosen or card.get("dismissed") or (
                    card.get("marked") and card["color"] != "yellow" and not card.get("finalized")):
                continue  # 自动试过没通过的、你已标蓝 / 绿还没提交的、这次没选的，保持原样
            if _unchecked_fix_reason(card, group, enabled_types):
                continue  # 未勾选的错误类型继续留给人工
            if not unchanged and any(_touched(it, rows) for it in group):
                card.update(handed_over=False, message=TOUCHED_MESSAGE)
                kept += 1
                continue
            card.update(color="yellow", verified=False, manual=False, finalized=False, marked=False,
                        waiting=True, handed_over=False, message=PICKED_MESSAGE if picked else WAITING_MESSAGE)
            card.pop("after_kind", None)
            for item in group:
                item.update(status="pending", batch_managed=True, triage_keep=False)
            back += 1
            if picked:
                given = wf.STATE["job"].setdefault("picked_back", [])
                if card["id"] not in given:
                    given.append(card["id"])
            debug_log.event("🤖 交给自动", 位置=f"{card['sheet_name']}!{card.get('cell_addr')}",
                            行=card.get("excel_row"), 你选的=bool(card.get("auto_pick")))
        for card in _data()["cards"].values():
            if card["wb_path"] == path:
                card.pop("auto_pick", None)
        wf._persist_state()
    # 自动处理时 Excel 里不留标记色；还留给人工的格子，提交或复检后再按设置标色。
    # 只调整选中的格子时，只去掉交给自动的这些格子的标色，没选的黄卡照旧标着
    _sync_highlights(session.app, path, clear=not picked)
    wf.log_msg(f"🤖 {Path(path).name}：{back} 处{'交给自动调整' if picked else '交回自动处理'}"
               + (f"，{kept} 处所在行转人工后被改过，留给人工" if kept else ""))


def _handed_over_paths(by_cell: dict[tuple[str, str, str], list[dict]] | None = None,
                       in_folder: Callable[[str], bool] | None = None) -> list[str]:
    """调用方持有 STATE_LOCK：当前文件夹里还有可交回自动的问题的文件"""
    enabled = _active_fix_types()
    in_folder = wf.in_current_folder if in_folder is None else in_folder
    cards = sorted(_data()["cards"].values(), key=lambda c: c["seq"])
    return list(dict.fromkeys(c["wb_path"] for c in cards if c.get("handed_over") and not c.get("dismissed")
                              and in_folder(c["wb_path"])
                              and not _unchecked_fix_reason(c, None if by_cell is None else by_cell.get(_cell(c), []), enabled)))


def _mode(by_cell: dict[tuple[str, str, str], list[dict]] | None = None,
          in_folder: Callable[[str], bool] | None = None) -> dict[str, bool]:
    """调用方持有 STATE_LOCK：现在能不能「转人工」/「转回自动」"""
    job = wf.STATE["job"] or {}
    auto = job.get("kind") == "batch_scan" and not job.get("options", {}).get("scan_only")
    active = wf.STATE["busy"] or job.get("status") == "interrupted"
    return {
        # 全部文件扫描完才能切；暂停中也可以直接转
        "to_manual": auto and active and job.get("phase") == "repair" and not job.get("to_manual"),
        "to_auto": not wf.STATE["busy"] and not (auto and job.get("status") == "interrupted")
                   and bool(_handed_over_paths(by_cell, in_folder)),
    }


def request_manual() -> bool:
    """「转人工」：任务在跑就做完当前这一行后转；已暂停返回 True，由调用方续跑任务去做转人工"""
    with wf.STATE_LOCK:
        if not _mode()["to_manual"]:
            job = wf.STATE["job"] or {}
            if job.get("kind") == "batch_scan" and job.get("phase") != "repair" and (
                    wf.STATE["busy"] or job.get("status") == "interrupted"):
                raise ValueError("全部文件扫描完成后才能转人工，请稍候")
            raise ValueError("没有正在进行的全自动处理")
        job = wf.STATE["job"]
        job["to_manual"] = True
        wf._persist_state()
        return not wf.STATE["busy"]


def check_auto(fix_types: Any = None) -> None:
    with wf.STATE_LOCK:
        if fix_types is not None:
            set_fix_error_types(fix_types, require_non_empty=True)
        if not _mode()["to_auto"]:
            cards = [c for c in _data()["cards"].values() if c.get("handed_over") and not c.get("dismissed")
                     and wf.in_current_folder(c["wb_path"])]
            if cards and all(_unchecked_fix_reason(c) for c in cards):
                raise ValueError("待交回自动的问题均属于未勾选的错误类型，请先勾选对应错误类型后再点「转回自动」")
            raise ValueError("没有可以交回自动处理的问题")


def _auto_retry_blocked(card: dict, file_versions: dict[str, Any] | None = None) -> bool:
    """自动试过没通过的格子：文件之后没改过就不再试（同样的算法，结果也一样）。调用方持有 STATE_LOCK"""
    path = card["wb_path"]
    if file_versions is not None and path in file_versions:
        current = file_versions[path]
    else:
        try:
            current = wf._file_version(Path(path))
        except OSError:
            current = None
        if file_versions is not None:
            file_versions[path] = current
    if current is None:
        return False
    done = _data().get("auto_versions", {}).get(path)
    # 旧版本没记自动处理完时的文件版本：按上次整份检测之后有没有改过判断
    return done == current if done is not None else _unchanged_since_check(path, current)


def _auto_skip_reason(card: dict, items: list[dict] | None = None, enabled: set[str] | None = None,
                      file_versions: dict[str, Any] | None = None) -> str:
    """「🤖 自动调整」时这一格为什么不能交给自动；能交给自动返回空字符串。调用方持有 STATE_LOCK"""
    if card["color"] != "yellow" or not card.get("finalized"):
        return "颜色标记还没提交，请先提交黄色标记"
    items = _items_at(card) if items is None else items
    if card.get("message") == TOUCHED_MESSAGE:
        return "这一行的行高被你改过，自动不再改这一行"
    if not _can_highlight(card, items):
        return "单元格位置没确认，不能自动调整"
    if not _is_break(card) and items and autofix.review_only_reason(items):
        return "PDF 里找不到这段文字，不能确认要不要自动调整"
    if unchecked := _unchecked_fix_reason(card, items, enabled):
        return unchecked
    if not card.get("handed_over") and _auto_retry_blocked(card, file_versions):
        return "自动试过没通过，文件之后没改过，再试结果也一样"
    return ""


def auto_adjust_plan(ids: Any, fix_types: Any = None) -> dict:
    """「🤖 自动调整」选中的黄卡：run 交给自动；extra 是同一行连带一起调整的黄卡（一行共用行高）；
    skip 是跳过的格子和原因。同一行有不能自动的格子时整行跳过（自动加高会连带改到它）。"""
    with wf.STATE_LOCK:
        if wf.STATE["busy"] and not _auto_queue_open():
            raise ValueError("后台任务执行中，请结束或暂停后再自动调整")
        if fix_types is not None:
            set_fix_error_types(fix_types)
        enabled = _active_fix_types()
        cards = {c["id"]: c for c in _data()["cards"].values()
                 if not c.get("dismissed") and wf.in_current_folder(c["wb_path"])}
        if not isinstance(ids, list) or not ids or any(not isinstance(cid, str) or cid not in cards for cid in ids):
            raise ValueError("请选择当前文件夹的有效审核卡片")
        ids = list(dict.fromkeys(ids))
        if wf.STATE["busy"]:
            # 自动调整在跑：排进去的文件不能是还没做完的（它的卡后台还要读、还要改）
            for path in dict.fromkeys(cards[cid]["wb_path"] for cid in ids):
                if reason := open_block_reason(wf.STATE["job"], Path(path)):
                    raise ValueError(f"{Path(path).name}：{reason.replace('才能打开', '再点自动调整')}")
        by_row: dict[str, list[dict]] = {}
        for card in cards.values():
            by_row.setdefault(_rid(card), []).append(card)
        run: list[str] = []
        extra: list[str] = []
        skip: dict[str, str] = {}
        for cid in ids:
            if cid in run or cid in extra or cid in skip:
                continue
            card = cards[cid]
            reason = _auto_skip_reason(card, enabled=enabled)
            peers = [c for c in by_row[_rid(card)] if c["id"] != cid]
            if not reason:
                for peer in peers:
                    if peer["color"] != "yellow" or not peer.get("finalized"):
                        reason = "同一行还有标了颜色、没提交的格子，请先提交"
                    elif autofix.review_only_reason(_items_at(peer)) and not _is_break(peer):
                        reason = "同一行有 PDF 里找不到的格子，整行不能自动加高"
                    elif peer.get("message") == TOUCHED_MESSAGE:
                        reason = "这一行的行高被你改过，自动不再改这一行"
                    if reason:
                        break
            if reason:
                skip[cid] = reason
                continue
            run.append(cid)
            for peer in peers:
                if peer["id"] in ids:
                    run.append(peer["id"])
                elif (not any(autofix.column_width_only(it) for it in _items_at(peer))
                      and not _unchecked_fix_reason(peer, _items_at(peer), enabled)):
                    extra.append(peer["id"])
        order = {cid: cards[cid]["seq"] for cid in cards}
        return {"run": sorted(dict.fromkeys(run), key=order.get),
                "extra": sorted(dict.fromkeys(extra), key=order.get),
                "skip": [{"id": cid, "reason": why} for cid, why in skip.items()]}


def _finish_picked(path: str) -> None:
    """「🤖 自动调整」做完一份文件：这次交给自动、没修好、也没写过原文件的黄卡直接算已提交，可以马上人工处理、复检，
    不用再提交一次黄色（修好的蓝卡照旧等你审核提交）"""
    with wf.STATE_LOCK:
        given = set(wf.STATE["job"].get("picked_back", []))
        left = 0
        for card in _data()["cards"].values():
            if (card["id"] not in given or card["wb_path"] != path or card.get("dismissed") or card["color"] != "yellow"
                    or card.get("finalized") or card.get("marked") or _waiting(card)):
                continue
            row = _row_record(card)
            if row and row["saved"]:
                continue
            card.update(manual=True, finalized=True)
            for item in _items_at(card):
                item.update(status="manual_review", batch_managed=True, triage_keep=False,
                            review_reason=card["message"])
            left += 1
            debug_log.event("🤖 自动没修好", 位置=f"{card['sheet_name']}!{card.get('cell_addr')}", 原因=card["message"])
        wf._persist_state()
    if left:
        wf.log_msg(f"🙋 {Path(path).name}：{left} 处自动没修好，仍是黄色人工项，可以直接人工处理、复检")


def _auto_queue_open() -> bool:
    """调用方持有 STATE_LOCK：正在跑「🤖 自动调整」，还能把别的文件排进去接着做；
    「只扫描」还没扫完时也能排，全部扫完后本任务接着自动调整（_queued_auto_to_picked）"""
    job = wf.STATE.get("job") or {}
    options = job.get("options", {})
    return bool(wf.STATE["busy"] and job.get("kind") == "batch_scan" and (options.get("picked") or options.get("scan_only"))
                and "workbooks" in job and not job.get("closing") and not job.get("pause_requested")
                and not job.get("to_manual"))


SCAN_LIMIT_FLOOR = 180  # 自动上限最少 3 分钟
SCAN_LIMIT_FIRST = 600  # 扫完的文件不到 2 份、还估不准时先给 10 分钟
SCAN_LIMIT_MARGIN = 5  # 自动上限 = 估计时间的 5 倍
SCAN_LIMIT_CHOICES = (3, 5, 10, 15)  # 网页上可选的固定分钟数


def _scan_size(job: dict, path: str) -> int:
    """文件大小（估算扫描时间用），每份只查一次，记在任务里"""
    sizes = job.setdefault("sizes", {})
    if path not in sizes:
        try:
            sizes[path] = Path(path).stat().st_size
        except OSError:
            sizes[path] = 0
    return sizes[path]


def _scan_model(job: dict) -> tuple[float, float] | None:
    """按这次已扫完文件的耗时估算扫描速度：(平均每份秒数, 每字节秒数)；一份都没扫完时返回 None"""
    done = job.get("scan_seconds") or {}
    if not done:
        return None
    total = sum(done.values())
    sizes = sum(_scan_size(job, p) for p in done)
    return total / len(done), (total / sizes if sizes else 0.0)


def _expected_scan_seconds(job: dict, path: str, model: tuple[float, float]) -> float:
    """这份文件大概要扫多久：一半按平均每份的时间，一半按文件大小折算（大文件慢）"""
    average, rate = model
    return (average + (rate * _scan_size(job, path) if rate else average)) / 2


def _scan_limit(job: dict, path: str) -> float:
    """这份文件最多扫多久（秒），超过就跳过、转人工。网页上选了分钟数就按它；自动时按估计时间的 5 倍，
    最少 3 分钟，最多不超过看门狗的无进展上限（默认 15 分钟）；扫完的不到 2 份时先给 10 分钟"""
    minutes = job["options"].get("scan_limit_minutes")
    if minutes:
        return minutes * 60.0
    model = _scan_model(job) if len(job.get("scan_seconds") or {}) >= 2 else None
    limit = (SCAN_LIMIT_FIRST if model is None
             else max(SCAN_LIMIT_FLOOR, SCAN_LIMIT_MARGIN * _expected_scan_seconds(job, path, model)))
    return min(limit, idle_timeout())


def _minutes(seconds: float) -> str:
    return f"{round(seconds / 60, 1):g} 分钟"


def _skip_slow_file(job: dict, path: str, limit: float) -> None:
    """扫描超过时间上限：这份文件跳过、这次任务里不再处理（继续上次任务也不重试），
    记下卡在哪张表，页面上提示人工检查；之后这份文件扫描通过就不再提示"""
    with wf.STATE_LOCK:
        now = wf.SCAN_NOW if wf.SCAN_NOW.get("path") == path else {}
        sheet = now.get("sheet") or ""
        job.setdefault("slow_seconds", {})[path] = round(time.time() - now["started"] if now else limit)
        _data().setdefault("slow_files", {})[path] = {"sheet": sheet, "limit": round(limit),
                                                      "at": time.strftime("%Y-%m-%d %H:%M")}
        for done in ("scanned", "completed"):
            if path not in job.setdefault(done, []):
                job[done].append(path)
        job.get("file_errors", {}).pop(path, None)
        wf.STATE["scan_progress"] = {"total": len(job["workbooks"]), "completed": len(job["completed"]),
                                     "status": "running"}
        wf._persist_state()
    message = f"扫描超过 {_minutes(limit)}" + (f"，卡在表「{sheet}」" if sheet else "") + "，已跳过，请人工检查"
    wf.log_msg(f"⏱ {Path(path).name}：{message}")
    debug_log.event("⏱ 扫描超时跳过", 文件=Path(path).name, 表=sheet, 上限秒=round(limit))
    update_file_progress(workbook_rel_name(Path(path), wf.STATE["folder"]), status="error", error=message,
                         current_action="扫描超时，已跳过")


def clear_slow_files() -> None:
    """页面上「知道了」：不再提示当前文件夹扫描超时跳过的文件"""
    with wf.STATE_LOCK:
        slow = _data().get("slow_files", {})
        for path in [p for p in slow if wf.in_current_folder(p)]:
            del slow[path]
        wf._persist_state()


def _scan_eta(job: dict) -> dict | None:
    """调用方持有 STATE_LOCK：扫描时的已用 / 预计剩余时间（秒）和当前文件的进度。剩余按这次已扫完文件的耗时、
    文件大小估算，一份都没扫完时不估；全自动的修复阶段、转回自动、自动调整不算"""
    if (not wf.STATE["busy"] or job.get("kind") != "batch_scan" or job["options"].get("back_to_auto")
            or job.get("phase") == "repair" or "workbooks" not in job):
        return None
    finished = set(job.get("scanned", [])) | set(job.get("completed", []))
    pending = [p for p in job["workbooks"] if p not in finished]
    now = wf.SCAN_NOW if wf.SCAN_NOW.get("path") in pending else {}
    running = max(0.0, time.time() - now["started"]) if now else 0.0
    model = _scan_model(job)
    limit = job.get("scan_limit") or {}
    return {
        "elapsed": round(sum((job.get("scan_seconds") or {}).values())
                         + sum((job.get("slow_seconds") or {}).values()) + running),
        "remaining": None if model is None
        else round(max(0.0, sum(_expected_scan_seconds(job, p, model) for p in pending) - running)),
        "current": {"name": Path(now["path"]).name, "sheet": now.get("sheet", ""),
                    "index": now.get("index", 0), "total": now.get("total", 0), "seconds": round(running),
                    "limit": limit.get("seconds") if limit.get("path") == now["path"] else None} if now else None,
    }


def _auto_queue_paths(job: dict) -> set[str]:
    """调用方持有 STATE_LOCK：「只扫描」时排队等自动调整的文件"""
    cards = _data()["cards"]
    return {cards[cid]["wb_path"] for cid in job.get("auto_queue", []) if cid in cards}


def _queued_auto_to_picked(job: dict) -> bool:
    """调用方持有 STATE_LOCK：「只扫描」全部扫完，扫描中排队的「🤖 自动调整」本任务接着做：
    任务就地改成自动调整选中的格子（和点「🤖 自动调整」另起的任务一样）。
    有文件扫描出错没扫完时，扫描任务先收起来，自动调整做完放回，仍可「继续上次任务」接着扫"""
    cards = _data()["cards"]
    ids = [cid for cid in job.pop("auto_queue", []) if cid in cards and not cards[cid].get("dismissed")]
    if not ids:
        return False
    if not set(job["workbooks"]).issubset(job["completed"]):
        job["paused_scan"] = copy.deepcopy({k: v for k, v in job.items() if k != "paused_scan"})
    paths = list(dict.fromkeys(cards[cid]["wb_path"] for cid in sorted(ids, key=lambda cid: cards[cid]["seq"])))
    options = job["options"]
    folder = options["folder"]
    fix_types = options.get("fix_error_types")
    options.clear()  # 原地改：run 里的 options 还指着它
    options.update(folder=folder, back_to_auto=True, picked=True, card_ids=ids,
                   **({"fix_error_types": fix_types} if fix_types is not None else {}))
    for key in ("phase", "repairing", "picked_back", "closing"):
        job.pop(key, None)
    job.update(workbooks=paths, completed=[], scanned=[], scan_versions={}, rounds={}, file_errors={})
    wf.STATE["scan_progress"] = {"total": len(paths), "completed": 0, "status": "running"}
    wf._persist_state()
    debug_log.event("🤖 扫完接着自动调整", 文件=len(paths), 格子=len(ids))
    wf.log_msg(f"🤖 全部文件扫描完成，接着自动调整排队的 {len(paths)} 个文件（{len(ids)} 格）")
    return True


def enqueue_auto_adjust(ids: Any, fix_types: Any = None) -> bool:
    """「🤖 自动调整」正在跑时又点了别的文件的自动调整：排进同一个任务，当前文件做完接着做，不用等它结束。
    没有在跑的自动调整（或它已经在收尾）返回 False，由调用方另起任务"""
    with wf.STATE_LOCK:
        if not _auto_queue_open():
            return False
    plan = auto_adjust_plan(ids, fix_types=fix_types)
    picked = plan["run"] + plan["extra"]
    if not picked:
        raise ValueError("选中的格子都不能自动调整：" + "；".join(dict.fromkeys(s["reason"] for s in plan["skip"])))
    with wf.STATE_LOCK:
        if not _auto_queue_open():
            raise ValueError("自动调整正在收尾，请过几秒再点")
        job = wf.STATE["job"]
        cards = _data()["cards"]
        paths = list(dict.fromkeys(cards[cid]["wb_path"] for cid in sorted(picked, key=lambda cid: cards[cid]["seq"])))
        for path in paths:
            if reason := open_block_reason(job, Path(path)):
                raise ValueError(f"{Path(path).name}：{reason.replace('才能打开', '再点自动调整')}")
        for cid in picked:
            cards[cid]["auto_pick"] = True
        if job["options"].get("scan_only"):
            # 「只扫描」还没扫完：先记下，全部扫完后本任务接着自动调整
            job["auto_queue"] = list(dict.fromkeys(job.get("auto_queue", []) + picked))
            wf._persist_state()
            debug_log.event("🖱️ 自动调整排队（扫完再做）", 选中=len(ids), 交给自动=len(plan["run"]),
                            同行连带=len(plan["extra"]), 跳过=[f"{s['id'][:6]}:{s['reason']}" for s in plan["skip"]])
            wf.log_msg(f"🤖 已排队自动调整：{'、'.join(Path(p).name for p in paths)}（{len(picked)} 格），全部文件扫完后接着做")
            return True
        options = job["options"]
        options["card_ids"] = list(dict.fromkeys(options["card_ids"] + picked))
        rounds = job.setdefault("rounds", {})
        for path in paths:
            # 这次任务里已经做完的文件：重新排队，再交一次你新选的格子
            for done in ("scanned", "completed"):
                if path in job.get(done, []):
                    job[done].remove(path)
            if path not in job["workbooks"]:
                job["workbooks"].append(path)
            job.get("file_errors", {}).pop(path, None)  # 出错跳过的文件重新排队：网页上不再显示出错
            rounds[path] = rounds.get(path, 0) + 1
        wf._persist_state()
        folder = Path(wf.STATE["folder"])
    add_to_execution_batch(folder, [Path(p) for p in paths])
    debug_log.event("🖱️ 自动调整排队", 选中=len(ids), 交给自动=len(plan["run"]), 同行连带=len(plan["extra"]),
                    跳过=[f"{s['id'][:6]}:{s['reason']}" for s in plan["skip"]])
    wf.log_msg(f"🤖 已排队自动调整：{'、'.join(Path(p).name for p in paths)}（{len(picked)} 格），前面的文件做完接着做")
    return True


def start_auto_adjust(ids: Any, fix_types: Any = None) -> dict:
    """「🤖 自动调整」：给要交给自动的卡打上 auto_pick，返回 batch_scan（转回自动）任务的参数"""
    plan = auto_adjust_plan(ids, fix_types=fix_types)
    picked = plan["run"] + plan["extra"]
    if not picked:
        raise ValueError("选中的格子都不能自动调整：" + "；".join(dict.fromkeys(s["reason"] for s in plan["skip"])))
    with wf.STATE_LOCK:
        for card in _data()["cards"].values():
            card.pop("auto_pick", None)  # 上次没跑起来留下的
        for cid in picked:
            _data()["cards"][cid]["auto_pick"] = True
        wf._persist_state()
        folder = wf.STATE["folder"]
        selected_types = fix_error_types()
    debug_log.event("🖱️ 自动调整", 选中=len(ids), 交给自动=len(plan["run"]), 同行连带=len(plan["extra"]),
                    跳过=[f"{s['id'][:6]}:{s['reason']}" for s in plan["skip"]])
    return {"folder": folder, "back_to_auto": True, "picked": True, "card_ids": picked,
            "fix_error_types": selected_types}


def _can_highlight(card: dict, items: list[dict] | None = None) -> bool:
    return bool(card.get("cell_addr")) and all(it.get("location_confirmed") is not False
                                             for it in (_items_at(card) if items is None else items))


def _sheet_key(name: str) -> str:
    return unicodedata.normalize("NFKC", name).strip().casefold()


def _find_sheet(book: Any, name: str) -> Any:
    """按名字找工作表；找不到时忽略首尾空格（含全角）、全半角和大小写再找一次。
    还是没有（被删除或改名了）返回 None。"""
    try:
        return book.Sheets(name)
    except Exception:  # 名字不存在时 COM 报 DISP_E_BADINDEX（-2147352565）
        pass
    want = _sheet_key(name)
    for sheet in book.Sheets:
        if _sheet_key(sheet.Name) == want:
            return sheet
    return None


def _mark_name(cid: str) -> str:
    """标记的标识：新方式写在条件格式规则的公式里，旧方式是隐藏名称"""
    return f"CHE_MARK_{cid[:20]}"


def _sheet_ref(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def _delete_name(book: Any, name: str | None) -> None:
    if not name:
        return
    try:
        book.Names(name).Delete()
    except Exception:
        pass


def _cell_rc(addr: str) -> tuple[int, int]:
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", addr.split(":")[0])
    if match is None:
        raise ValueError(f"无法识别的单元格地址：{addr}")
    col = 0
    for ch in match.group(1).upper():
        col = col * 26 + ord(ch) - 64
    return int(match.group(2)), col


def _locate_mark(book: Any, mark: dict) -> tuple[Any, tuple[int, int]] | None:
    """标过色的格子现在在哪：先按标色时建的隐藏名称找（插删行列、改表名后 Excel 会自动更新），
    返回 (工作表, (下移行数, 右移列数))；名称还在但格子已被删掉返回 None。
    旧记录没有名称、或名称被删了，按原来的表名和地址找（工作表找不到时工作表为 None）。"""
    name = mark.get("name")
    if name:
        try:
            defined = book.Names(name)
        except Exception:
            defined = None
        if defined is not None:
            try:
                area = defined.RefersToRange
                row, col = _cell_rc(mark["addr"])
                shift = (int(area.Row) - row, int(area.Column) - col)
                return area.Worksheet, shift
            except Exception:
                # 指向 #REF!：整张表被删了，或者格子被删了
                return (None, (0, 0)) if _find_sheet(book, mark["sheet"]) is None else None
    return _find_sheet(book, mark["sheet"]), (0, 0)


def _target(sheet: Any, addr: str, shift: tuple[int, int] = (0, 0)) -> Any:
    """记录里的一格或一整块合并区域（$A$1:$C$2），按插删行列挪动的行数、列数找到现在的位置"""
    if shift == (0, 0):
        return sheet.Range(addr)
    corners = [sheet.Cells(row + shift[0], col + shift[1]) for row, col in map(_cell_rc, addr.split(":"))]
    return corners[0] if len(corners) == 1 else sheet.Range(corners[0], corners[1])


def _restore_range(target: Any, cell: dict, mark: dict) -> bool:
    fill, pattern = target.Interior.Color, target.Interior.Pattern
    font = target.Font.Color if cell["font"] is not None else 0
    if ":" in cell["addr"] and (fill is None or pattern is None or font is None):
        # 整块记下的合并区域，现在颜色不一致了（人工改过其中几格）：一格一格看，只还原还是标记色的格子
        return any([_restore_range(one, {**cell, "addr": one.Address}, mark) for one in target.Cells])
    changed = False
    if fill is not None and pattern is not None and int(fill) == mark["fill"] and int(pattern) == XL_PATTERN_SOLID:
        changed = True
        if cell["pattern"] == XL_PATTERN_NONE:
            target.Interior.Pattern = XL_PATTERN_NONE
        else:
            target.Interior.Pattern = cell["pattern"]
            target.Interior.Color = cell["color"]
            target.Interior.TintAndShade = cell["tint"]
    if cell["font"] is not None and font is not None and int(font) == mark["font"]:
        target.Font.Color = cell["font"]
    return changed


def _restore_cells(sheet: Any, mark: dict, shift: tuple[int, int] = (0, 0)) -> bool:
    """只还原仍是标记颜色的格子；人工处理时自己改过颜色的不动。shift 是格子被插删行列挪动的行数、列数。
    返回这里原来是不是标记色（不是的话，标记色可能被挪到了别处）。"""
    return any([_restore_range(_target(sheet, cell["addr"], shift), cell, mark) for cell in mark["cells"]])


def _still_marked(book: Any, mark: dict) -> bool:
    """上次标色有没有真的存进文件：中途被打断（看门狗结束 Excel、程序被关）时记录在、颜色没保存"""
    found = _locate_mark(book, mark)
    if found is None or found[0] is None or not mark["cells"]:
        return False
    fill = _target(found[0], mark["cells"][0]["addr"], found[1]).Interior.Color
    return fill is None or int(fill) == mark["fill"]  # 合并区域颜色不一致：至少部分还是标记色，按已标过算


def _shows_mark(mark: dict | None) -> bool:
    return mark is not None and mark.get("saved") is not False


XL_FORMULAS, XL_WHOLE = -4123, 1
UNHIGHLIGHT_STALL_SECONDS = 300


def _marked_matches(sheet: Any, value: str, fill: int) -> list[tuple[int, int]]:
    """在工作表里按内容找（整格完全一致、区分大小写），只留下底色还是标记色的格子（合并区域按左上角）。
    不用「按格式查找」：实测在用户机器上不生效，会把每一格都当成匹配。"""
    if not value or len(value) > 255:
        return []  # Excel 查找最多 255 个字
    what = value.replace("~", "~~").replace("*", "~*").replace("?", "~?")
    cells = sheet.UsedRange
    cell = cells.Find(What=what, LookIn=XL_FORMULAS, LookAt=XL_WHOLE, MatchCase=True)
    found: list[tuple[int, int]] = []
    first, steps = None, 0
    while cell is not None and steps < 200:
        if cell.Address == first:
            break
        first = first or cell.Address
        area = cell.MergeArea
        top = (int(area.Row), int(area.Column))
        color = area.Interior.Color
        if (top not in found and color is not None and int(color) == fill
                and int(area.Interior.Pattern) == XL_PATTERN_SOLID):
            found.append(top)
        cell = cells.FindNext(cell)
        steps += 1
    return found


def _sweep_leftover_marks(book: Any, path: str, missed: list[dict], kept: list[dict]) -> list[dict]:
    """兜底：原位置没找到标记色的记录（多半是旧记录没有隐藏名称时被插删行列挪了位置），
    按标色时记下的格子内容在同一张表里找，找到的格子还得是标记底色、而且只有一个，才按这条记录还原；
    其他情况不乱改。只在有这种记录时才找，平常不多花时间。"""
    keep: set[tuple[str, int, int]] = set()
    for mark in kept:
        located = _locate_mark(book, mark)
        if located and located[0] is not None:
            sheet, (dr, dc) = located
            for cell in mark["cells"]:
                row, col = _cell_rc(cell["addr"])
                keep.add((sheet.Name, row + dr, col + dc))
    fixed, left, done = 0, [], []
    for mark in missed:
        sheet = _find_sheet(book, mark["sheet"])
        if sheet is None or not mark.get("value"):
            continue  # 表不在了（前面已说明）或旧记录没记内容：没法找
        try:
            spots = [rc for rc in _marked_matches(sheet, mark["value"], mark["fill"])
                     if (sheet.Name, *rc) not in keep]
            debug_log.event("🔎 兜底查找", 位置=f"{mark['sheet']}!{mark['addr']}", 内容=debug_log.clip(mark["value"]),
                            找到=[f"{_col_letters(c)}{r}" for r, c in spots[:5]])
            if len(spots) == 1:
                r0, c0 = _cell_rc(mark["addr"])
                if _restore_cells(sheet, mark, (spots[0][0] - r0, spots[0][1] - c0)):
                    fixed += 1
                    done.append(mark)
                    continue
            if spots:
                left.append(f"{sheet.Name}!" + "、".join(f"{_col_letters(c)}{r}" for r, c in spots[:3]))
        except Exception as err:
            debug_log.exception(err, f"兜底查找 {mark['sheet']}!{mark['addr']}")
            continue
    if fixed:
        wf.log_msg(f"🧹 {Path(path).name}：{fixed} 格标记颜色挪了位置，按格子内容找到并还原了")
    if left:
        wf.log_msg(f"⚠️ {Path(path).name}：有 {len(left)} 处标记颜色挪了位置、对上的格子不止一个，没有自动还原，"
                   "请在 Excel 里看一下：" + "；".join(left[:20]) + ("……" if len(left) > 20 else ""))
    return done


def _col_letters(col: int) -> str:
    letters = ""
    while col:
        col, rest = divmod(col - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


XL_EXPRESSION = 2
MARK_TAG = re.compile(r"CHE_MARK_[0-9A-Za-z]+")


def _is_rule(mark: dict) -> bool:
    """新方式的记录：问题格上加一条条件格式规则，单元格本身的颜色不动。旧记录（改单元格底色）带 cells"""
    return "tag" in mark


def _add_rule(area: Any, tag: str, mark: dict) -> None:
    """给问题格（合并格整块）加一条永远成立的条件格式规则，公式里带这张卡的标识：
    =N("CHE_MARK_…")=0，N(文字) 永远是 0。调到最前面并勾「如果为真则停止」，表里原有的条件格式也盖得住。
    规则跟着格子走（插删行列、剪切移动），清除时按标识找，不用记原色、不用找格子。"""
    rule = _new_rule(area.FormatConditions, f'=N("{tag}")=0')
    rule.Interior.Color = mark["fill"]
    rule.Font.Color = mark["font"]
    rule.StopIfTrue = True
    rule.SetFirstPriority()  # 最后再调顺序：先调的话，有的 Excel 版本手里的规则对象会指到别的规则上


def _mark_blocker(book: Any, sheet: Any) -> str | None:
    """标色出错时查一下 Excel 为什么不让加条件格式：工作表受保护、共享工作簿、几张表被同时选中（组合）。
    这几种情况 Excel 只回一个空的「例外が発生しました」，看不出原因。都不是时返回 None"""
    def flag(obj: Any, name: str) -> bool:
        try:
            return getattr(obj, name) is True
        except Exception:
            return False

    if sheet is not None and flag(sheet, "ProtectContents"):
        return "工作表受保护（校閲 → シート保護の解除 后再试）"
    if flag(book, "MultiUserEditing"):
        return "这是共享工作簿（ブックの共有），Excel 不允许加条件格式；取消共享后再试"
    try:
        if int(book.Windows(1).SelectedSheets.Count) > 1:
            return "几张工作表被同时选中（组合 / グループ），Excel 不允许加条件格式"
    except Exception:
        pass
    return None


def _ungroup_sheets(book: Any) -> bool:
    """文件保存时几张表是组合状态（同时选中）的话，打开后加条件格式会失败：只选中第一张可见的表，解除组合。
    返回是否解除了"""
    try:
        if int(book.Windows(1).SelectedSheets.Count) <= 1:
            return False
        for sheet in book.Worksheets:
            if sheet.Visible == -1:  # xlSheetVisible
                sheet.Select()
                return True
    except Exception as err:
        debug_log.exception(err, "解除工作表组合")
    return False


DISP_E_PARAMNOTOPTIONAL = -2147352561
_raw_add = False  # 这台机器上用普通写法加规则会报「参数不能省略」，之后直接用底层调用


def _param_not_optional(err: Exception) -> bool:
    info = getattr(err, "excepinfo", None) or (err.args[2] if len(getattr(err, "args", ())) > 2 else None)
    return bool(info) and len(info) > 5 and info[5] == DISP_E_PARAMNOTOPTIONAL


def _new_rule(rules: Any, formula: str) -> Any:
    """加一条公式型条件格式规则。
    有的 Excel 上 pywin32 按类型库帮忙补上没写的参数（Formula2、String、ScopeType…），Excel 不认，
    回「例外が発生しました」（scode 0x8002000F 参数不能省略）。这时改用底层 Invoke，只传 Type、Operator、Formula1 三个参数，
    和 VBA 里 .Add Type:=xlExpression, Formula1:=… 一样。Operator 对公式型规则没用，传 xlBetween(1) 占位"""
    global _raw_add
    ole = getattr(rules, "_oleobj_", None)
    if not _raw_add or ole is None:
        try:
            return rules.Add(Type=XL_EXPRESSION, Formula1=formula)
        except Exception as err:
            if ole is None or not _param_not_optional(err):
                raise
            debug_log.event("🔧 加规则改用底层调用", 原因=err)
            _raw_add = True
    import win32com.client
    dispid = ole.GetIDsOfNames("Add")
    return win32com.client.Dispatch(ole.Invoke(dispid, 0, 1, True, XL_EXPRESSION, 1, formula))  # 1 = DISPATCH_METHOD


def _rule_tag(rule: Any) -> str | None:
    try:
        formula = str(rule.Formula1)  # 色阶、数据条这类规则没有公式，读了会报错
    except Exception:
        return None
    found = MARK_TAG.search(formula)
    return found.group(0) if found else None


def _drop_rules(book: Any, keep: set[str], tick: Callable[[], None] | None = None,
                holders: dict[str, Any] | None = None) -> tuple[set[str], dict[str, int], dict[str, str]]:
    """整份文件每张表的条件格式过一遍：公式里带我们标识、又不在 keep 里的规则删掉，原有的规则和顺序都不动。
    复制粘贴带到别处、拆成几条的，标识一样，一起删；表改了名也照样找到。
    holders：传入时记下还留着标记的工作表（标签键 -> 工作表），用来给工作表标签标色。
    返回 (还留着的标识, 删掉的标识 -> 条数, 删不掉的标识 -> 原因)。"""
    present: set[str] = set()
    dropped: dict[str, int] = {}
    failed: dict[str, str] = {}
    for sheet in book.Worksheets:
        rules = sheet.Cells.FormatConditions
        for i in range(int(rules.Count), 0, -1):  # 从后往前删，前面的序号不变
            rule = rules.Item(i)
            tag = _rule_tag(rule)
            if tag is None:
                continue
            if tag in keep:
                present.add(tag)
                if holders is not None:
                    holders[_tab_key(sheet)] = sheet
                continue
            try:
                rule.Delete()
                dropped[tag] = dropped.get(tag, 0) + 1
            except Exception as err:  # 例如工作表受保护
                failed[tag] = f"{sheet.Name}：{err}"
        if tick is not None:
            tick()
    return present, dropped, failed


# 有问题格的工作表，标签标成橙黄色（Excel 标准色「オレンジ」#FFC000），多张表时一眼看出哪张要改
TAB_HIGHLIGHT = "#FFC000"
XL_COLOR_INDEX_NONE = -4142


def _tab_key(sheet: Any) -> str:
    """认工作表用的键：优先用代码名（改表名也不变），读不到代码名时用表名"""
    try:
        code = str(sheet.CodeName or "")
    except Exception:
        code = ""
    return code or f"name:{sheet.Name}"


def _sheet_by_tab_key(book: Any, key: str, name: str) -> Any:
    if not key.startswith("name:"):
        for sheet in book.Worksheets:
            try:
                if str(sheet.CodeName or "") == key:
                    return sheet
            except Exception:
                continue
    return _find_sheet(book, name)


def _read_tab(sheet: Any) -> dict:
    """记下工作表标签原来的颜色：没颜色 / 主题色（ThemeColor + TintAndShade）/ 普通颜色"""
    tab = sheet.Tab
    index = int(tab.ColorIndex)
    record: dict[str, Any] = {"index": index}
    if index == XL_COLOR_INDEX_NONE:
        return record
    record["color"] = int(tab.Color)
    try:
        theme = int(tab.ThemeColor)  # 不是主题色时有的版本会报错
    except Exception:
        theme = 0
    record["theme"] = theme if theme > 0 else None
    try:
        record["tint"] = float(tab.TintAndShade)
    except Exception:
        record["tint"] = 0.0
    return record


def _tab_is_ours(sheet: Any) -> bool:
    tab = sheet.Tab
    return int(tab.ColorIndex) != XL_COLOR_INDEX_NONE and int(tab.Color) == _excel_color(TAB_HIGHLIGHT)


def _restore_tab(sheet: Any, record: dict) -> None:
    tab = sheet.Tab
    if record["index"] == XL_COLOR_INDEX_NONE:
        tab.ColorIndex = XL_COLOR_INDEX_NONE
    elif record.get("theme"):
        tab.ThemeColor = record["theme"]
        tab.TintAndShade = record.get("tint", 0.0)
    else:
        tab.Color = record["color"]
        if record.get("tint"):
            tab.TintAndShade = record["tint"]


def _tabs_stale(path: str, sheets: set[str]) -> bool:
    """标过色的工作表标签和这次还要标的工作表（按标记记录里的表名）对不上：要打开文件改标签。调用方持有 STATE_LOCK"""
    return {t["sheet"] for t in _tabs().get(path, {}).values()} != sheets


def _sync_tabs(book: Any, path: str, holders: dict[str, Any], report: dict | None = None) -> dict[str, int]:
    """工作表标签和标记保持一致：表里还有标记的标成橙黄色（先记下原来的颜色），标记全去掉的改回原来的颜色。
    标签已经不是橙黄色（有人改过，或上次标完没保存）的不覆盖，只删记录。改不了的（工作簿结构受保护等）记日志，
    要还原的记录留着下次再试。返回 {"marked": 新标的张数, "restored": 改回的张数, "failed": 出错的张数}"""
    book_name = Path(path).name
    count = {"marked": 0, "restored": 0, "failed": 0}
    with wf.STATE_LOCK:
        records = _tabs().setdefault(path, {})
        for key, sheet in holders.items():
            if key in records:
                records[key]["sheet"] = str(sheet.Name)
    for key, sheet in holders.items():
        name = str(sheet.Name)
        if key not in records:
            try:
                if _tab_is_ours(sheet):
                    # 已经是橙黄色却没有记录（例如旧状态文件丢了）：原色不知道，还原时当作原来没颜色
                    original = {"index": XL_COLOR_INDEX_NONE, "guessed": True}
                else:
                    original = _read_tab(sheet)
            except Exception as err:
                # 读不了标签颜色就不改：也记一笔，下次不用为它再打开文件，还原时直接跳过
                original = {"index": XL_COLOR_INDEX_NONE, "skip": True}
                count["failed"] += 1
                wf.log_msg(f"⚠️ {book_name} / 工作表「{name}」读不了标签颜色，不标橙黄色：{err}")
                debug_log.exception(err, "读工作表标签颜色")
            with wf.STATE_LOCK:
                records[key] = {"sheet": name, **original}
                wf._persist_state()  # 先记下原色再改：中途崩溃时下次还能改回去
        if records[key].get("skip"):
            continue
        try:
            if not _tab_is_ours(sheet):
                sheet.Tab.Color = _excel_color(TAB_HIGHLIGHT)
                count["marked"] += 1
                debug_log.event("🏷️ 工作表标签标橙黄", 文件=book_name, 工作表=name, 原色=records[key])
        except Exception as err:
            count["failed"] += 1
            wf.log_msg(f"⚠️ {book_name} / 工作表「{name}」的标签未能标成橙黄色：{err}")
            debug_log.exception(err, "工作表标签标色")
    for key in [k for k in records if k not in holders]:
        record = records[key]
        status = "ok"
        try:
            sheet = None if record.get("skip") else _sheet_by_tab_key(book, key, record["sheet"])
            if record.get("skip"):
                reason = "当时没能标色，不用还原"
            elif sheet is None:
                reason = "工作表已不在文件里，不用还原"
            elif _tab_is_ours(sheet):
                _restore_tab(sheet, record)
                count["restored"] += 1
                reason = "已改回原来的标签颜色"
            else:
                reason = "标签已经不是橙黄色（有人改过，或上次标完没保存），不覆盖"
            debug_log.event("🏷️ 还原工作表标签", 文件=book_name, 工作表=record["sheet"], 结果=reason, 原色=record)
        except Exception as err:
            count["failed"] += 1
            status, reason = "failed", f"还原标签颜色出错，下次再试：{err}"
            wf.log_msg(f"⚠️ {book_name} / 工作表「{record['sheet']}」的标签{reason}")
            debug_log.exception(err, "还原工作表标签")
        if report is not None and not record.get("skip"):
            report.setdefault("tabs", []).append({"sheet": record["sheet"], "status": status, "reason": reason})
        if status == "ok":
            with wf.STATE_LOCK:
                records.pop(key, None)
    with wf.STATE_LOCK:
        if not records:
            _tabs().pop(path, None)
        wf._persist_state()
    return count


def _sync_highlights(excel_for: Callable[[], Any], path: str, *, clear: bool = False,
                     tick: Callable[[], None] | None = None, report: dict | None = None) -> None:
    """让原 Excel 里的标记和卡片一致：黄色人工待处理的格子加一条条件格式规则；
    复检通过（变蓝）、提交为蓝 / 绿、或要求清除时把规则删掉。打开一次、保存一次。
    旧版本改过单元格底色的记录照旧还原原来的颜色。
    tick：每处理一格告诉看门狗还在动，问题格很多的大文件要好几分钟，不会被当成 Excel 卡死。
    report：清除标记时传入，逐格记下去掉了没有、没去掉的原因（cells），以及清掉的残留规则条数，页面上显示结果。"""
    with wf.STATE_LOCK:
        data, marks = _data(), _marks()
        style = data.get("highlight_files", {}).get(path)
        cards = [c for c in data["cards"].values() if c["wb_path"] == path]
        want = set() if clear else _wanted_marks(path)
        have = {cid for cid, m in marks.items() if m["path"] == path}
        # 没确认存进文件的记录（上次标色被打断，或旧版本没记这一项）：打开文件看一眼标记还在不在
        unsure = sorted(cid for cid in have if not marks[cid].get("saved"))
        add = [c for c in cards if c["id"] in want - have]
        remove = sorted(have - want)
        tabs_stale = _tabs_stale(path, {marks[cid]["sheet"] for cid in have if cid not in remove}
                                 | {c["sheet_name"] for c in add})
    if not add and not remove and not unsure and not tabs_stale:
        return
    _check_file(path)
    debug_log.event("🎨 同步标色", 文件=Path(path).name, 清除=clear, 要标=len(add), 要去掉=len(remove),
                    待核对=len(unsure), 标签要核对=tabs_stale, 样式=style or "-")
    app = excel_for()
    book = open_workbook_checked(app, str(Path(path).resolve()), UpdateLinks=0)
    added: dict[str, dict] = {}
    failed: dict[str, str] = {}
    deleted = restored = unmarked = 0
    leftovers: dict[str, int] = {}
    tabs = {"marked": 0, "restored": 0, "failed": 0}
    stuck: list[str] = []  # 去掉 / 还原时出错的记录，留着下次再试
    book_name, last_log = Path(path).name, time.monotonic()

    def note(mark: dict, status: str, reason: str) -> None:
        """status：ok 已去掉 / check 需要人看一下 / failed 没去掉"""
        if report is not None:
            report.setdefault("cells", []).append(
                {"sheet": mark["sheet"], "addr": str(mark["addr"]).replace("$", ""), "status": status, "reason": reason})

    def step(what: str, done: int, total: int) -> None:
        nonlocal last_log
        if tick is not None:
            tick()
        if total > 50 and time.monotonic() - last_log >= 30:
            last_log = time.monotonic()
            wf.log_msg(f"🎨 {book_name}：正在{what} {done} / {total} 格…")

    try:
        try:
            app.ScreenUpdating = False  # 隐藏的后台 Excel 也不用每改一格就重画
        except Exception:
            pass
        if book.ReadOnly:
            raise ValueError(f"{Path(path).name} 以只读方式打开，未能标色")
        # 条件格式规则：留下还要标的，其余带标识的全部删掉（包括崩溃留下的、复制粘贴带出去的）
        keep = {marks[cid]["tag"] for cid in have if cid not in remove and _is_rule(marks[cid])}
        holders: dict[str, Any] = {}  # 还留着标记的工作表，标签要标橙黄色
        present, dropped, drop_failed = _drop_rules(book, keep, tick, holders)
        ours = {marks[cid]["tag"]: cid for cid in remove if _is_rule(marks[cid])}
        for tag, cid in ours.items():
            if tag in drop_failed:
                stuck.append(cid)
                wf.log_msg(f"⚠️ {book_name} / {marks[cid]['sheet']}!{marks[cid]['addr']} 未能去掉标记，下次再试："
                           f"{drop_failed[tag]}")
                note(marks[cid], "failed", f"删不掉标记规则（{drop_failed[tag]}），工作表可能受保护")
            else:
                unmarked += 1  # 规则被人手动删掉的也算：本来就没东西要清
                note(marks[cid], "ok", "已去掉标记" if tag in dropped else "标记规则已被删掉（或格子已删），不用清")
            debug_log.event("🧽 去掉标记", 位置=f"{marks[cid]['sheet']}!{marks[cid]['addr']}", 标识=tag,
                            条数=dropped.get(tag, 0), 出错=drop_failed.get(tag, "-"))
        for tag, count in dropped.items():
            if tag not in ours:
                leftovers[tag] = count
        if leftovers:
            debug_log.event("🧹 残留标记", 标识=list(leftovers)[:10], 条数=sum(leftovers.values()))
        left_failed = [reason for tag, reason in drop_failed.items() if tag not in ours]
        if report is not None:
            report["leftovers"] = sum(leftovers.values())
            report["leftover_failed"] = left_failed[:20]
        if left_failed:
            wf.log_msg(f"⚠️ {book_name}：{len(left_failed)} 条残留标记没能去掉：" + "；".join(left_failed[:5]))
        lost = []
        for i, cid in enumerate(unsure, 1):
            if cid in remove and _is_rule(marks[cid]):
                continue  # 本来就要去掉
            still = marks[cid]["tag"] in present if _is_rule(marks[cid]) else _still_marked(book, marks[cid])
            if still:
                marks[cid]["saved"] = True
            else:
                lost.append(cid)
            step("核对上次的标色", i, len(unsure))
        if lost:
            with wf.STATE_LOCK:
                for cid in lost:
                    _delete_name(book, marks.pop(cid).get("name"))
                remove = [cid for cid in remove if cid not in lost]
                add += [c for c in cards if c["id"] in want and c["id"] in lost]
            wf.log_msg(f"ℹ️ {book_name}：{len(lost)} 格上次标色没有存进文件（多半是中途被打断），这次重新标")
        sheets: dict[str, Any] = {}

        def sheet_of(name: str) -> Any:
            if name not in sheets:
                sheets[name] = _find_sheet(book, name)
            return sheets[name]

        areas: dict[str, Any] = {}
        if add and _ungroup_sheets(book):
            wf.log_msg(f"ℹ️ {book_name}：文件里几张工作表是组合状态（同时选中），已解除组合再标色")
        blockers: dict[str, str | None] = {}

        def why(name: str, err: Exception) -> str:
            """出错原因：能查出来的（受保护、共享工作簿、组合）直接说，查不出来的照抄 Excel 的报错"""
            if name not in blockers:
                blockers[name] = _mark_blocker(book, sheet_of(name))
                debug_log.event("🔎 标色出错原因", 工作表=name, 原因=blockers[name] or "查不出，见上面的异常")
            return blockers[name] or str(err)

        for i, card in enumerate(add, 1):
            step("找问题格", i, len(add))
            sheet = sheet_of(card["sheet_name"])
            if sheet is None:
                failed[card["id"]] = f"工作表「{card['sheet_name']}」已不在文件里，可能已被删除或改名"
                continue
            try:
                # 分页切开的卡片是整个合并区域（A10:H12），按左上角找合并区域
                areas[card["id"]] = area = sheet.Range(card["cell_addr"].split(":")[0]).MergeArea
                added[card["id"]] = {"path": card["wb_path"], "sheet": card["sheet_name"], "addr": str(area.Address),
                                     "tag": _mark_name(card["id"]), "fill": _excel_color(style["fill"]),
                                     "font": _excel_color(style["font"])}
            except Exception as err:
                failed[card["id"]] = why(card["sheet_name"], err)
                debug_log.event("❌ 标色失败", 位置=f"{card['sheet_name']}!{card.get('cell_addr')}", 原因=err)
                debug_log.exception(err, "标色")
        with wf.STATE_LOCK:
            # 先记下再加规则、再保存：中途出错时记录还在，之后按标识还能删掉
            for mark in added.values():
                mark["saved"] = False  # 保存成功后才算标上了
            marks.update(added)
            wf._persist_state()
        for i, (cid, mark) in enumerate(list(added.items()), 1):
            step("标色", i, len(added))
            try:
                _add_rule(areas[cid], mark["tag"], mark)
                debug_log.event("🎨 标色", 位置=f"{mark['sheet']}!{mark['addr']}", 标识=mark["tag"])
            except Exception as err:  # 例如工作表受保护
                failed[cid] = why(mark["sheet"], err)
                with wf.STATE_LOCK:
                    marks.pop(cid, None)
                del added[cid]
                debug_log.event("❌ 标色失败", 位置=f"{mark['sheet']}!{mark['addr']}", 原因=err)
                debug_log.exception(err, "标色")
        # 旧版本改过单元格底色的记录：照旧还原原来的颜色
        gone: dict[str, int] = {}
        missed: list[dict] = []  # 没在原位置找到标记色的记录，留给下面的兜底检查
        old = [cid for cid in remove if not _is_rule(marks[cid])]
        for i, cid in enumerate(old, 1):
            step("还原颜色", i, len(old))
            mark = marks[cid]
            try:
                found = _locate_mark(book, mark)
                debug_log.event("🧽 还原", 位置=f"{mark['sheet']}!{mark['addr']}", 名称=mark.get("name"),
                                找到=found is not None and found[0] is not None,
                                挪位=found[1] if found and found[0] is not None else "-",
                                表在=not (found is not None and found[0] is None))
                if found is None:
                    deleted += 1  # 标色的格子已被人工删掉，没有颜色要还原
                    note(mark, "ok", "格子已被删掉，不用还原")
                elif found[0] is None:
                    # 工作表已被删除（例如精简 Sheet 或人工处理时删掉）就没有颜色要还原了，按表汇总说一次
                    gone[mark["sheet"]] = gone.get(mark["sheet"], 0) + 1
                    missed.append(mark)
                elif _restore_cells(found[0], mark, found[1]):
                    restored += 1
                    note(mark, "ok", "已还原成原来的颜色")
                else:
                    debug_log.event("🔎 原位置不是标记色，留给兜底查找", 位置=f"{mark['sheet']}!{mark['addr']}")
                    missed.append(mark)
            except Exception as err:
                # 还原出错：记录留着（按钮上的数字还在），下次提交 / 复检 / 清除时再试，不当成已还原
                stuck.append(cid)
                wf.log_msg(f"⚠️ {Path(path).name} / {mark['sheet']}!{mark['addr']} 未能还原颜色，下次再试：{err}")
                note(mark, "failed", f"还原颜色出错：{err}")
                continue
            _delete_name(book, mark.get("name"))
        if missed:
            kept = [m for cid, m in marks.items() if m["path"] == path and cid not in remove and not _is_rule(m)]
            found_later = _sweep_leftover_marks(book, path, missed, kept)
            for mark in missed:
                if any(mark is m for m in found_later):
                    note(mark, "ok", "格子挪了位置，按内容找到并还原了")
                elif mark["sheet"] in gone:
                    note(mark, "check", "工作表已不在文件里（被删除或改名）；如果是改名，请在新表里看一下有没有标记色")
                else:
                    note(mark, "check", "原位置已经不是标记色（可能人工改过颜色，或格子挪了位置），请看一下")
        for name, count in gone.items():
            wf.log_msg(f"ℹ️ {Path(path).name}：工作表「{name}」已不在文件里（被删除或改名），"
                       f"跳过还原 {count} 格标记颜色；如果是改名，请在新表里手动去掉标记色")
        for cid, mark in marks.items():
            # 这次新标的，和旧版本改底色、还留着的标记（新方式还留着的在 _drop_rules 里已经找到）
            if mark["path"] == path and cid not in remove and (cid in added or not _is_rule(mark)):
                sheet = sheet_of(mark["sheet"])
                if sheet is not None:
                    holders.setdefault(_tab_key(sheet), sheet)
        tabs = _sync_tabs(book, path, holders, report)
        if added or remove or dropped or tabs["marked"] or tabs["restored"]:  # 只核对、没有要改的：不保存
            if tick is not None:
                tick()
            normalize_workbook_view(app, book)
            before = _version_or_none(path)
            book.Save()
            _followed_own_save(path, before)
        for mark in added.values():
            mark["saved"] = True
    finally:
        try:
            app.ScreenUpdating = True
        except Exception:
            pass
        book.Close(SaveChanges=False)
    with wf.STATE_LOCK:
        for cid in remove:
            if cid not in stuck:
                marks.pop(cid, None)
        for cid, error in failed.items():
            card = data["cards"].get(cid)
            if card:
                card["message"] = f"未能在 Excel 中标色（{error}），请按截图查找后人工处理"
        wf._persist_state()
    parts = [f"标色 {len(added)} 格"] if added else []
    if unmarked:
        parts.append(f"去掉标记 {unmarked} 格")
    if leftovers:
        parts.append(f"清掉残留标记 {sum(leftovers.values())} 条")
    if tabs["marked"]:
        parts.append(f"{tabs['marked']} 张工作表标签标成橙黄色")
    if tabs["restored"]:
        parts.append(f"{tabs['restored']} 张工作表标签改回原来的颜色")
    if tabs["failed"]:
        parts.append(f"{tabs['failed']} 张工作表标签没改成（见上面的提示）")
    if restored:
        parts.append(f"还原颜色 {restored} 格")  # 只算真的改回去的格子（旧版本改过底色的）
    if stuck:
        parts.append(f"{len(stuck)} 格去掉标记出错，记录留着下次再试")
    if deleted:
        parts.append(f"{deleted} 格标色的格子已被删掉，不用还原")
    if failed:
        reasons = sorted(set(failed.values()))
        parts.append(f"{len(failed)} 格未能标色（{'；'.join(reasons[:3])}{'……' if len(reasons) > 3 else ''}）")
    if parts:
        wf.log_msg(f"🎨 {Path(path).name}：" + "，".join(parts))


def _wanted_marks(path: str) -> set[str]:
    """这份文件里现在应该标色的卡片：黄色人工待处理的。调用方持有 STATE_LOCK"""
    data = _data()
    if not data.get("highlight_files", {}).get(path):
        return set()
    return {c["id"] for c in data["cards"].values()
            if c["wb_path"] == path and not c.get("dismissed") and c["color"] == "yellow"
            and c.get("manual") and _can_highlight(c)}


def _leftover_mark_paths(skip: list[str]) -> list[str]:
    """上次没还原成的标记颜色（例如 Excel 中途崩溃）所在的文件：卡片已经不用标了、记录还在。
    提交 / 复检时顺带把这些文件也还原，不用等这份文件下次再复检"""
    with wf.STATE_LOCK:
        paths = dict.fromkeys(m["path"] for m in _marks().values()
                              if m["path"] not in skip and wf.in_current_folder(m["path"])
                              and Path(m["path"]).exists())
        found = [p for p in paths if {cid for cid, m in _marks().items() if m["path"] == p} - _wanted_marks(p)]
        # 标记都去掉了、工作表标签还没改回去的（上次改标签时出错或中途崩溃）
        found += [p for p in _tabs() if p not in skip and p not in found and wf.in_current_folder(p)
                  and Path(p).exists() and not any(m["path"] == p for m in _marks().values())]
        return found


def _where(err: BaseException) -> str:
    """出错在哪个函数哪一行（只看本程序的代码），日志里带上，下次出问题能直接定位"""
    frames = [f for f in traceback.extract_tb(err.__traceback__) if "core" in Path(f.filename).parts]
    return f"（出错位置：{frames[-1].name} 第 {frames[-1].lineno} 行）" if frames else ""


def _unhighlight_summary(results: dict[str, dict]) -> str:
    cells = [c for r in results.values() for c in r["cells"]]
    count = {status: sum(1 for c in cells if c["status"] == status) for status in ("ok", "check", "failed")}
    parts = [f"{len(results)} 份文件", f"已清除 {count['ok']} 格"]
    if count["check"]:
        parts.append(f"需要看一下 {count['check']} 格")
    if count["failed"]:
        parts.append(f"没清掉 {count['failed']} 格")
    errors = sum(1 for r in results.values() if r.get("error"))
    if errors:
        parts.append(f"{errors} 份文件没能处理")
    tabs = sum(1 for r in results.values() for t in r.get("tabs", []) if t["status"] == "ok")
    if tabs:
        parts.append(f"工作表标签已处理 {tabs} 张")
    leftovers = sum(r.get("leftovers", 0) for r in results.values())
    if leftovers:
        parts.append(f"另清掉残留标记 {leftovers} 条")
    return "🧽 清除完成：" + "，".join(parts)


def _sync_highlights_safely(session: Any, path: str) -> None:
    """提交、复检之后顺带更新标色；失败只记日志，不影响审核结果。
    Excel 出毛病（崩溃、连不上）时换一个新的 Excel 再试一次；还是不行，记录留着，下次提交 / 复检时再还原"""
    if path in (wf.STATE["job"] or {}).get("deferred_adjustments", {}):
        return  # 上次写入还没结算：保存标色会改动文件，之后就核对不了，等继续或放弃后再标
    for tries in range(2):
        try:
            _sync_highlights(session.app, path, tick=session.touch)
            return
        except Exception as err:
            if isinstance(err, ValueError) or tries:
                wf.log_msg(f"⚠️ {Path(path).name} 未能更新 Excel 里的标记颜色：{err}{_where(err)}")
                return
            wf.log_msg(f"⚠️ {Path(path).name} 更新标记颜色出错，重启 Excel 后重试：{err}{_where(err)}")
            try:
                _restart_excel(session)
            except Exception as restart_err:
                wf.log_msg(f"❌ 无法重新启动 Excel：{restart_err}")
                return


def highlighted_paths() -> list[str]:
    with wf.STATE_LOCK:
        paths = list(dict.fromkeys([m["path"] for m in _marks().values() if wf.in_current_folder(m["path"])]
                                   + [p for p in _tabs() if wf.in_current_folder(p)]))
        if not paths:
            raise ValueError("当前文件夹没有标过颜色的单元格")
        return paths


def recheck_options(ids: list[str], *, single: bool = False) -> dict:
    """复检任务的参数。full_paths 里的文件整份重查；卡片上的「单项复检」（single）只查点的那几格"""
    with wf.STATE_LOCK:
        paths = [] if single else list(dict.fromkeys(_data()["cards"][cid]["wb_path"] for cid in ids))
    return {"card_ids": ids, "full_paths": paths}


def _full_recheck(options: dict, path: str) -> bool:
    # 旧版本留下的任务没有 full_paths：那时的复检都是整份重查
    return "full_paths" not in options or path in options["full_paths"]


def enqueue_recheck(ids: list[str], *, single: bool = False) -> bool:
    """复检任务正在跑时把新点的复检排进同一个任务（同一个后台 Excel 接着查），不用等它结束再点。
    没有在跑的复检任务、或它已经查完准备收尾时返回 False，由调用方另起任务。"""
    with wf.STATE_LOCK:
        job = wf.STATE.get("job") or {}
        if not wf.STATE.get("busy") or job.get("kind") != "batch_recheck" or job.get("closing"):
            return False
        cards = _data()["cards"]
        options = job["options"]
        options["card_ids"] = list(dict.fromkeys(options["card_ids"] + ids))
        for path in dict.fromkeys(cards[cid]["wb_path"] for cid in ids):
            if "full_paths" in options:
                queued = (path in job["workbooks"] and path != job.get("current")
                          and path not in job.get("completed", []) and path not in job.get("skipped", []))
                if not single and path not in options["full_paths"]:
                    options["full_paths"].append(path)
                elif single and not queued and path in options["full_paths"]:
                    # 这份文件整份查过或正在查，现在只要再查这一格；还在排队等整份查的，整份查已包含这一格
                    options["full_paths"].remove(path)
            if path == job.get("current"):
                # 正在查的文件又点了复检（多半是刚又改过）：查完再查一遍，没改过的话第二遍不会重新导出
                if path not in job.setdefault("again", []):
                    job["again"].append(path)
                continue
            # 已查完或跳过的文件重新排队
            for done in ("completed", "skipped"):
                if path in job.get(done, []):
                    job[done].remove(path)
            if path not in job["workbooks"]:
                job["workbooks"].append(path)
        wf._persist_state()
    wf.log_msg(f"➕ 已加入复检队列：{len(ids)} 格，当前文件查完后接着查")
    return True


FOLLOWUP_NAMES = {"batch_recheck": "复检", "batch_commit": "提交"}


def enqueue_followup(kind: str, ids: list[str], *, single: bool = False, color: str | None = None) -> bool:
    """后台在跑别的批处理任务时点了复检 / 提交：排到它后面，当前任务结束后接着做（start_followup）。
    后台空着返回 False，由调用方直接开始。自动处理 / 复检还要处理的文件不能排：它的卡后台还要读、还要改"""
    verb = FOLLOWUP_NAMES[kind]
    with wf.STATE_LOCK:
        if not wf.STATE["busy"]:
            return False
        job = wf.STATE.get("job") or {}
        if not job.get("kind", "").startswith("batch_"):
            raise ValueError(f"后台正在{wf.QUICK_TASK_NAMES.get(job.get('kind'), '处理别的任务')}，结束后才能{verb}")
        cards = _data()["cards"]
        paths = list(dict.fromkeys(cards[cid]["wb_path"] for cid in ids))
        if job["kind"] in {"batch_scan", "batch_recheck"} and "workbooks" in job:
            for path in paths:
                if reason := open_block_reason(job, Path(path)):
                    raise ValueError(f"{Path(path).name}：{reason.replace('才能打开', f'再点{verb}')}")
        queue = _data().setdefault("followups", [])
        if queue and queue[-1]["kind"] == kind and queue[-1].get("color") == color:
            entry = queue[-1]  # 和排在最后的同一种合并；中间隔着别的就另排，先后顺序不变
        else:
            entry = ({"kind": kind, "card_ids": [], "full_paths": []} if kind == "batch_recheck"
                     else {"kind": kind, "card_ids": [], "color": color})
            queue.append(entry)
        entry["card_ids"] = list(dict.fromkeys(entry["card_ids"] + ids))
        if kind == "batch_recheck" and not single:
            entry["full_paths"] = list(dict.fromkeys(entry["full_paths"] + paths))
        wf._persist_state()
    debug_log.event(f"🖱️ {verb}排队", 格子=len(ids), 文件=len(paths))
    wf.log_msg(f"⏳ 已排队{verb}：{'、'.join(Path(p).name for p in paths)}（{len(ids)} 格），当前任务结束后接着{verb}")
    return True


def start_followup(ran: dict) -> bool:
    """调用方持有 STATE_LOCK，后台刚空下来：排在后面的复检 / 提交接着开始，卡和颜色按现在的算。
    ran 是刚结束的任务：你点了暂停的先不做，页面上可以「现在开始」或取消；网页「现在开始」时传空字典"""
    queue = _data().get("followups") or []
    if wf.STATE["busy"] or ran.get("pause_requested") or not queue:
        return False
    started = False
    while queue and not started:
        entry = queue.pop(0)
        verb = FOLLOWUP_NAMES[entry["kind"]]
        cards = _data()["cards"]
        # 排队后又提交 / 复检过、换了文件夹的卡不算；复检只查还是已提交黄卡的
        ids = [cid for cid in entry["card_ids"] if (card := cards.get(cid)) and not card.get("dismissed")
               and wf.in_current_folder(card["wb_path"])
               and (entry["kind"] != "batch_recheck" or (card["color"] == "yellow" and card.get("finalized")))]
        if not ids:
            wf.log_msg(f"ℹ️ 排队的{verb}已经没有要{verb}的卡（可能已经提交或复检过），跳过")
            continue
        if entry["kind"] == "batch_recheck":
            mine = {cards[cid]["wb_path"] for cid in ids}
            options = {"card_ids": ids, "full_paths": [p for p in entry["full_paths"] if p in mine]}
        else:
            options = {"card_ids": ids, **({"color": entry["color"]} if entry.get("color") else {})}
        try:
            started = wf.queue_job(entry["kind"], options)
        except ValueError as err:
            wf.log_msg(f"⚠️ 排队的{verb}没能开始，已跳过：{err}")
            continue
        if started:
            wf.log_msg(f"▶️ 开始排队的{verb}：{len({cards[cid]['wb_path'] for cid in ids})} 个文件（{len(ids)} 格）")
    wf._persist_state()
    return started


def cancel_followups() -> int:
    """取消还没开始的排队复检 / 提交，返回取消了几项"""
    with wf.STATE_LOCK:
        queue = _data().pop("followups", None) or []
        wf._persist_state()
    if queue:
        wf.log_msg(f"🚫 已取消 {len(queue)} 项排队的复检 / 提交")
    return len(queue)


def open_ready(ids: list[str]) -> dict:
    """后台在跑时「打开全部人工待处理文件」：后台不再碰的文件马上在可见 Excel 打开、跳到第一处黄格，
    还要等的文件列出原因，不排队（等它的话还不如做完一份开一份）"""
    with wf.STATE_LOCK:
        cards = sorted((_data()["cards"][cid] for cid in ids), key=lambda c: c["seq"])
        first = {}
        for card in cards:
            first.setdefault(card["wb_path"], card)
        job = wf.STATE.get("job") or {}
        busy = wf.STATE["busy"]
        waiting = {path: reason for path in first if busy and (reason := open_block_reason(job, Path(path)))}
    targets = [(Path(path), c["sheet_name"], c["cell_addr"], c["excel_row"]) for path, c in first.items()
               if path not in waiting]
    if targets:
        try:
            autofix.wait_for_writes(15)  # 后台正在保存的文件别同时打开
        except ValueError:
            raise ValueError("后台正在保存 Excel，请几秒后再点") from None
    errors = open_workbooks_for_review(targets) if targets else {}
    for path, error in errors.items():
        wf.log_msg(f"⚠️ {path.name}：{error}")
    return {"opened": [Path(p).name for p, *_ in targets if p not in errors],
            "waiting": [{"name": Path(p).name, "reason": why} for p, why in waiting.items()],
            "errors": [{"name": p.name, "reason": str(e)} for p, e in errors.items()]}


def validate_cards(ids: Any = None, *, yellow: bool = False) -> list[str]:
    with wf.STATE_LOCK:
        available = {c["id"]: c for c in _data()["cards"].values()
                     if not c.get("dismissed") and wf.in_current_folder(c["wb_path"])}
        if ids is None:
            ids = list(available)
        if not isinstance(ids, list) or not ids or any(not isinstance(cid, str) or cid not in available for cid in ids):
            raise ValueError("请选择当前文件夹的有效审核卡片")
        if yellow and any(available[cid]["color"] != "yellow" or not available[cid].get("finalized") for cid in ids):
            raise ValueError("请先提交黄色标记，完成还原，再人工修改和复检")
        return list(dict.fromkeys(ids))


def _switch_to_manual(session: Any, job: dict, paths: list[str]) -> None:
    """全自动中途「转人工」：当前这一行已经做完。已修好的蓝卡原样留着审核；
    还没轮到的格子全部转黄、在 Excel 里标色，记下转人工时的行高，之后「转回自动」时用"""
    options = job["options"]
    left = 0
    for path in paths:
        with wf.STATE_LOCK:
            if isinstance(options.get("highlight"), dict):
                _data().setdefault("highlight_files", {})[path] = highlight_style(options["highlight"])
            left += sum(1 for c in _data()["cards"].values() if c["wb_path"] == path and _waiting(c)
                        and not c.get("dismissed"))
        _hand_over(path, switch=True)
        _sync_highlights_safely(session, path)
        _remember_handover(path)
    with wf.STATE_LOCK:
        job["completed"] = list(dict.fromkeys(job["completed"] + paths))
        job.pop("to_manual", None)
        job.pop("repairing", None)
        wf._persist_state()
    wf.log_msg(f"🙋 已转人工：{left} 处还没自动处理的问题转为黄色人工项；已修好的蓝卡保留，可随时审核")


def run() -> list[str]:
    """由 workflow 的单一任务槽执行；不依赖浏览器心跳。"""
    job = wf.STATE["job"]
    kind, options = job["kind"], job["options"]
    scan_only = kind == "batch_scan" and bool(options.get("scan_only"))
    session = autofix._ExcelSession()
    # 每种任务都看着后台 Excel：它在看不见的地方卡住（弹了没人能点的框等）时结束它，不让任务永远卡着。
    # 清除颜色每改一格都报进度，卡 5 分钟就算卡住；检测类一张大表导出可能要很久，按默认时长
    watchdog = ExcelWatchdog(lambda: wf.log_msg(
        "⏱ 这份文件扫描超过时间上限，已结束本任务专用 Excel，跳过这份文件" if watchdog.deadline_hit else
        "⚠️ 后台 Excel 长时间无进展，已结束本任务专用 Excel，当前文件将记录错误并重试或跳过"),
        timeout=UNHIGHLIGHT_STALL_SECONDS if kind == "batch_unhighlight" else None)
    session.watchdog = watchdog
    if watchdog is not None:
        watchdog.start()
    wf.pythoncom.CoInitialize()
    try:
        excel = session.app() if job.get("pending_adjustment") and kind != "batch_scan" else None
        if excel is not None:
            try:
                _finish_pending(excel)
            except Exception as err:
                # 核对还是不通过：只把这份文件放一边，其他文件照常做
                path = job["pending_adjustment"]["path"]
                _park_pending(err)
                wf.log_msg(f"❌ {Path(path).name} 写入未结算，已跳过，继续其他文件：{err}")
                session.close()
        if kind != "batch_scan" and "workbooks" not in job:
            # 审核提交 / 复检 / 打开 / 清除颜色也记下本次涉及的文件，网页按本任务显示进度
            with wf.STATE_LOCK:
                job["workbooks"] = (list(dict.fromkeys(options["paths"])) if kind == "batch_unhighlight" else
                                    list(dict.fromkeys(_data()["cards"][cid]["wb_path"] for cid in options["card_ids"])))
                wf._persist_state()
        if kind == "batch_scan":
            back = bool(options.get("back_to_auto"))
            picked = back and bool(options.get("picked"))
            if "workbooks" not in job:
                if back:
                    # 「转回自动」：只处理还有可交回自动的问题的文件，审核数据都在，不清空。
                    # 「🤖 自动调整」只处理选中的格子所在的文件
                    with wf.STATE_LOCK:
                        folder = Path(wf.STATE["folder"])
                        cards = _data()["cards"]
                        job["workbooks"] = (list(dict.fromkeys(
                            cards[cid]["wb_path"] for cid in sorted(
                                (cid for cid in options["card_ids"] if cid in cards), key=lambda cid: cards[cid]["seq"])))
                            if picked else _handed_over_paths())
                        wf._persist_state()
                    paths = [Path(p) for p in job["workbooks"]]
                else:
                    folder, paths = wf.list_workbooks(options["folder"])
                    with wf.STATE_LOCK:
                        if str(folder) != wf.STATE["folder"]:
                            # 审核卡片、忽略规则、行高记录按文件路径区分，各文件夹分开保留：切回来还能审核、还原
                            wf.STATE.update(items={}, pages={}, checked_files={})
                        wf.STATE.update(folder=str(folder), pdf_folder=str(folder))
                        job["workbooks"] = [str(p) for p in paths]
                        wf._persist_state()
                init_execution_batch(task_name="只扫描（全部转人工）" if scan_only else
                                     "自动调整选中的格子" if picked else "全自动扫描与修复",
                                     folder=folder, workbooks=paths, pdf_folder=folder)
            expected = job["workbooks"]
            with wf.STATE_LOCK:
                if job.get("pending_adjustment"):
                    pending = job["pending_adjustment"]
                    job.setdefault("deferred_adjustments", {})[pending["path"]] = pending
                    job["pending_adjustment"] = None
                # 先把全部文件扫完，再逐行修；旧版本的任务没分两步，做完的文件当作扫过
                scanned = job.setdefault("scanned", list(job["completed"]))
                versions = job.setdefault("scan_versions", {})
                job["closing"] = False  # 暂停后继续：又可以往里排自动调整
                wf._persist_state()

            def each_file(path: str, work: Callable[[str], None]) -> bool:
                """处理一份文件；程序出错（多半是 Excel 崩溃或卡住）重启 Excel 再试，还不行就跳过、可续跑"""
                for tries in range(autofix.MAX_ERROR_RETRIES + 1):
                    try:
                        watchdog.progress()
                        work(path)
                        watchdog.progress()
                        return True
                    except Exception as err:
                        if job.get("pending_adjustment"):
                            _park_pending(err)
                            wf.log_msg(f"❌ {Path(path).name} 写入未结算，已跳过，继续其他文件：{err}")
                            with wf.STATE_LOCK:
                                job.setdefault("file_errors", {})[path] = str(err)
                                wf._persist_state()
                            update_file_progress(workbook_rel_name(Path(path), wf.STATE["folder"]),
                                                 status="error", error=str(err), current_action="写入待核对，已跳过，可续跑")
                            session.close()
                            return False
                        if watchdog.deadline_hit:
                            # 扫描超过时间上限：不重试，这份文件跳过、转人工，下一份用新的 Excel
                            _skip_slow_file(job, path, (job.get("scan_limit") or {}).get("seconds") or 0)
                            session.close()
                            return False
                        # 检测类原因（ValueError）不重试
                        if (not isinstance(err, (ValueError, TimeoutError)) and not watchdog.timed_out
                                and tries < autofix.MAX_ERROR_RETRIES):
                            wf.log_msg(f"⚠️ {Path(path).name} 出错，重启 Excel 后重试：{err}")
                            try:
                                _restart_excel(session)
                            except Exception as restart_err:
                                wf.log_msg(f"❌ 无法重新启动 Excel：{restart_err}")
                                err = restart_err
                            else:
                                continue
                        wf.log_msg(f"❌ {Path(path).name} 已跳过，继续其他文件，可续跑：{err}")
                        with wf.STATE_LOCK:
                            job.setdefault("file_errors", {})[path] = str(err)
                            wf._persist_state()
                        update_file_progress(workbook_rel_name(Path(path), wf.STATE["folder"]),
                                             status="error", error=str(err), current_action="文件未完成，可续跑")
                        # 下一个文件用新的 Excel，避免一次崩溃拖垮后面所有文件
                        session.close()
                        return False
                    finally:
                        session.drop_copy(path)
                return False

            def scan(path: str) -> None:
                debug_log.begin(f"{'自动调整' if picked else '转回自动' if back else '扫描'} "
                                  f"{expected.index(path) + 1} / {len(expected)}："
                                  f"{workbook_rel_name(Path(path), wf.STATE['folder'])}")
                _check_file(path)
                excel = session.app()
                if picked:
                    _back_to_auto(excel, session, path, picked=True)
                elif back:
                    _back_to_auto(excel, session, path)
                else:
                    # 只给检测（只读打开原文件）设上限：超时结束 Excel 不会弄坏文件；之后标色要写文件，不在上限内
                    with wf.STATE_LOCK:
                        limit = _scan_limit(job, path)
                        job["scan_limit"] = {"path": path, "seconds": round(limit)}
                    started, checking = time.perf_counter(), {}
                    watchdog.deadline(limit)
                    try:
                        _refresh(excel, path, session=session, timing=checking)
                    finally:
                        watchdog.deadline(None)
                    with wf.STATE_LOCK:
                        job.setdefault("scan_seconds", {})[path] = round(time.perf_counter() - started, 1)
                        _data().get("slow_files", {}).pop(path, None)  # 以前超时跳过的，这次扫完了
                    wf.log_msg(f"⏱ {Path(path).name} 整份检查：{time.perf_counter() - started:.1f} 秒"
                               f"（导出 {int(checking.get('exports', 0))} 次 {checking.get('export', 0):.1f} 秒，"
                               f"检测 {checking.get('detect', 0):.1f} 秒）")
                if scan_only:
                    # 不试修、不改行高：检出的全部转人工，在原 Excel 里标色方便定位
                    with wf.STATE_LOCK:
                        _data().setdefault("highlight_files", {})[path] = highlight_style(options.get("highlight"))
                    _hand_over(path)
                    _sync_highlights(session.app, path, tick=session.touch)
                    _remember_handover(path)

            # 「🤖 自动调整」运行中还能排进新的文件（enqueue_auto_adjust），做完一轮再看有没有新排进来的。
            # 每份文件每轮（排一次队算一轮）只试一次，出错跳过的不在这次任务里反复重试
            tried: set[tuple[str, str, int]] = set()

            def first_try(step: str, path: str) -> bool:
                key = (step, path, job.get("rounds", {}).get(path, 0))
                if key in tried:
                    return False
                tried.add(key)
                return True

            def scan_all() -> None:
                for path in list(expected):
                    if path in scanned:
                        continue
                    if job.get("pause_requested"):
                        break
                    if not first_try("scan", path) or not each_file(path, scan):
                        continue
                    with wf.STATE_LOCK:
                        scanned.append(path)
                        versions[path] = wf._file_version(Path(path))
                        job.get("file_errors", {}).pop(path, None)
                        if scan_only:
                            job["completed"].append(path)
                            wf.STATE["scan_progress"] = {"total": len(expected), "completed": len(job["completed"]),
                                                         "status": "running"}
                        wf._persist_state()
                    if not scan_only and not back:
                        wf._sync_problem_counts(path, "扫描完成，等待自动处理")
                if not scan_only and not job.get("pause_requested"):
                    with wf.STATE_LOCK:
                        job["phase"] = "repair"  # 全部扫完：从这里起可以「转人工」
                        wf._persist_state()

            def repair(path: str) -> None:
                debug_log.begin(f"修复 {expected.index(path) + 1} / {len(expected)}："
                                  f"{workbook_rel_name(Path(path), wf.STATE['folder'])}")
                _check_file(path)
                excel = session.app()
                if job.get("pending_adjustment"):
                    _finish_pending(excel)
                with wf.STATE_LOCK:
                    # 全自动处理时运行中不改颜色（免得和写入行高冲突）；提交后黄色人工项再按这个颜色标
                    if isinstance(options.get("highlight"), dict):
                        _data().setdefault("highlight_files", {})[path] = highlight_style(options["highlight"])
                    elif options.get("highlight_off"):
                        _data().get("highlight_files", {}).pop(path, None)
                    stale = job.get("repairing") == path or versions.get(path) != wf._file_version(Path(path))
                    job["repairing"] = path
                    wf._persist_state()
                if stale:
                    # 上次修到一半停下（或扫描后文件又写过）：先按现在的文件重新检测，再接着修
                    _refresh(excel, path, session=session)
                _repair_file(excel, session, path)

            def repair_all() -> None:
                for path in list(expected) if job.get("phase") == "repair" else []:
                    if _stopping(job):
                        break
                    if path not in scanned or (path in job["completed"] and path not in job.get("deferred_adjustments", {})):
                        continue
                    if not first_try("repair", path):
                        continue
                    with wf.STATE_LOCK:
                        pending = job.get("deferred_adjustments", {}).pop(path, None)
                        if pending:
                            job["pending_adjustment"] = pending
                            wf._persist_state()
                    if not each_file(path, repair) or _stopping(job):
                        continue
                    with wf.STATE_LOCK:
                        if path not in job["completed"]:
                            job["completed"].append(path)
                        job.pop("repairing", None)
                        job.get("file_errors", {}).pop(path, None)
                        # 记下自动处理完这份文件时的版本：之后文件没改过，自动没修好的格子再点「自动调整」也不重试
                        _data().setdefault("auto_versions", {})[path] = _version_or_none(path)
                        wf.STATE["scan_progress"] = {"total": len(expected), "completed": len(job["completed"]), "status": "running"}
                        wf._persist_state()
                    if picked:
                        _finish_picked(path)
                        # 自动没修好的格子马上在 Excel 里标回颜色（没选的黄卡本来就标着）
                        _sync_highlights_safely(session, path)
                    wf._sync_problem_counts(path, "自动调整完成，等待结果审核" if picked else "全自动处理完成，等待结果审核")

            while True:
                scan_all()
                repair_all()
                with wf.STATE_LOCK:
                    more = picked and not _stopping(job) and any(
                        p not in job["completed"]
                        and ("repair" if p in scanned else "scan", p, job.get("rounds", {}).get(p, 0)) not in tried
                        for p in expected)
                    if scan_only and not job.get("pause_requested") and _queued_auto_to_picked(job):
                        # 只扫描时排队的「🤖 自动调整」：全部扫完，本任务接着按自动调整做
                        scan_only, back, picked = False, True, True
                        expected, scanned, versions = job["workbooks"], job["scanned"], job["scan_versions"]
                        tried.clear()
                        more = True
                    if not more:
                        job["closing"] = picked or scan_only  # 之后再点自动调整另起任务，不会漏掉
                        wf._persist_state()
                        break
            if job.get("to_manual") and not job.get("pause_requested"):
                _switch_to_manual(session, job, [p for p in expected if p in scanned])
        elif kind == "batch_commit":
            # 全是蓝卡、原文件不用改时不开 Excel；但标蓝的「最后一次尝试」要把那次的行高写进原文件，也要开
            if options.get("refresh_sheets") or any((c["color"] != "blue" and not c.get("dismissed")) or _adopting(c)
                                                   for c in _data()["cards"].values()):
                excel = session.app()
            _commit_review(excel)
            expected = list(dict.fromkeys(_data()["cards"][cid]["wb_path"] for cid in options["card_ids"]))
            for path in expected:
                _sync_highlights_safely(session, path)
            for path in _leftover_mark_paths(expected):
                if not workbook_locked(Path(path)):  # 你正在 Excel 里改的文件不动，下次再还原
                    _sync_highlights_safely(session, path)
        elif kind in {"batch_recheck", "batch_open"}:
            cards = [_data()["cards"][cid] for cid in options["card_ids"]]
            expected = list(dict.fromkeys(c["wb_path"] for c in cards))
            if kind == "batch_open":
                targets = [(Path(path), c["sheet_name"], c["cell_addr"], c["excel_row"])
                           for path in expected for c in [next(c for c in cards if c["wb_path"] == path)]]
                errors = open_workbooks_for_review(targets)
                job["completed"].extend(path for path in expected if Path(path) not in errors)
                for path, error in errors.items():
                    wf.log_msg(f"⚠️ {path.name}：{error}")
            else:
                excel = session.app()
                with wf.STATE_LOCK:
                    job.update(skipped=[], again=[], closing=False, current=None)
                while True:
                    with wf.STATE_LOCK:
                        # 复检进行中还能继续点复检（enqueue_recheck 追加到本任务），这里每次取下一个没查过的文件
                        path = next((p for p in job["workbooks"]
                                     if p not in job["completed"] and p not in job["skipped"]), None)
                        if path is None:
                            job["closing"], job["current"] = True, None  # 之后再点复检另起任务，不会漏掉
                            break
                        job["skipped"].append(path)
                        if _unsettled(job, path):
                            continue
                        job["current"] = path
                        cards = [card for cid in dict.fromkeys(options["card_ids"])
                                 for card in [_data()["cards"].get(cid)]
                                 if card and card["wb_path"] == path and not card.get("dismissed")]
                    debug_log.begin(f"复检 {Path(path).name}（{len(cards)} 张卡）")
                    try:
                        unsaved = close_saved_workbooks([Path(path)])
                        if Path(path) in unsaved:
                            wf.log_msg(f"⚠️ {Path(path).name}：{unsaved[Path(path)]}，请保存后继续")
                            continue
                        if cards:
                            session.touch()
                            full = _full_recheck(options, path)
                            for tries in range(2):
                                try:
                                    if full:
                                        _recheck_file(session.app(), path, cards)  # Excel 重启过就用新的
                                    else:
                                        for card in cards:
                                            if not card.get("dismissed"):
                                                _recheck_cell(session.app(), path, card)
                                    break
                                except Exception as err:
                                    if isinstance(err, ValueError) or tries:
                                        raise
                                    # 多半是后台 Excel 崩溃或出了毛病：换一个新的 Excel 再查一次，
                                    # 不然 Excel 一挂，后面每份文件都报「RPC サーバーを利用できません」
                                    wf.log_msg(f"⚠️ {Path(path).name} 复检出错，重启 Excel 后重试：{err}{_where(err)}")
                                    _restart_excel(session)
                            _sync_highlights_safely(session, path)
                        with wf.STATE_LOCK:
                            job["skipped"].remove(path)
                            if path not in job["again"]:
                                job["completed"].append(path)
                            wf._persist_state()
                    except Exception as err:
                        wf.log_msg(f"❌ {Path(path).name} 复检失败，可继续任务：{err}{_where(err)}")
                    finally:
                        with wf.STATE_LOCK:
                            job["current"] = None
                            if path in job["again"]:
                                # 查的时候又点了这个文件的复检：排到队尾再查一遍
                                job["again"].remove(path)
                                if path in job["skipped"]:
                                    job["skipped"].remove(path)
                                job["workbooks"].remove(path)
                                job["workbooks"].append(path)
                            wf._persist_state()
                expected = list(job["workbooks"])
                for path in _leftover_mark_paths(expected):
                    if not workbook_locked(Path(path)):  # 你正在 Excel 里改的文件不动，下次再还原
                        _sync_highlights_safely(session, path)
        elif kind == "batch_unhighlight":
            expected = list(options["paths"])
            wf.log_msg(f"🧽 开始清除标记颜色：{len(expected)} 份文件")
            unsaved = close_saved_workbooks([Path(path) for path in expected])
            with wf.STATE_LOCK:
                # 每份文件的清除结果，结束后页面弹窗显示；继续上次任务时已完成的文件结果留着
                results = job.setdefault("result", {})
            for i, path in enumerate(expected, 1):
                if path in job["completed"]:
                    continue
                report = {"name": workbook_rel_name(Path(path), wf.STATE["folder"]), "cells": [], "leftovers": 0,
                          "leftover_failed": [], "error": None}
                if _unsettled(job, path):
                    with wf.STATE_LOCK:
                        results[path] = {**report, "error": "上次写入还没完成，请先继续上次任务或放弃这次写入"}
                    continue
                if Path(path) in unsaved:
                    wf.log_msg(f"⚠️ {Path(path).name}：{unsaved[Path(path)]}，请保存后继续")
                    with wf.STATE_LOCK:
                        results[path] = {**report, "error": f"{unsaved[Path(path)]}，请保存并关闭后再点清除"}
                    continue
                debug_log.begin(f"清除颜色 {i} / {len(expected)}：{Path(path).name}")
                wf.log_msg(f"🧽 正在清除（{i} / {len(expected)}）：{Path(path).name}")
                for tries in range(2):
                    report.update(cells=[], tabs=[], leftovers=0, leftover_failed=[], error=None)
                    try:
                        watchdog.progress()
                        _sync_highlights(session.app, path, clear=True, tick=session.touch, report=report)
                        with wf.STATE_LOCK:
                            results[path] = report
                            # 清除后不再自动标色；以后再「只扫描」才重新标
                            _data().get("highlight_files", {}).pop(path, None)
                            job["completed"].append(path)
                            wf._persist_state()
                        break
                    except Exception as err:
                        if isinstance(err, ValueError) or tries:
                            wf.log_msg(f"❌ {Path(path).name} 清除标记颜色失败，可继续任务：{err}{_where(err)}")
                            with wf.STATE_LOCK:
                                results[path] = {**report, "error": f"清除失败：{err}"}
                            break
                        # 后台 Excel 卡住被结束、或崩溃了：换一个新的再试，后面的文件也用新的
                        wf.log_msg(f"⚠️ {Path(path).name} 清除标记颜色出错，重启 Excel 后重试：{err}{_where(err)}")
                        _restart_excel(session)
            wf.log_msg(_unhighlight_summary(results))
        else:
            raise ValueError("未知批处理任务")
        finish_execution_batch("处理完成，请审核结果" if set(expected).issubset(job["completed"]) else "部分文件尚未完成，可继续任务")
        return expected
    finally:
        try:
            with wf.STATE_LOCK:
                _show_deferred_pending()
                wf._persist_state()
        finally:
            try:
                session.close()
            finally:
                if watchdog is not None:
                    watchdog.stop()
                wf.pythoncom.CoUninitialize()
