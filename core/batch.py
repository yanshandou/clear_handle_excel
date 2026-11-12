"""无人值守扫描、标记，以及按单元格进行的三色结果审核。

扫描先将自动分页且末页不足 40% 的表缩放到少一页，再检测 PDF、在原 Excel 里标色。
不自动加高、开换行或调整原有样式（用户 2026-11-05 定）。
审核颜色只写数据库；旧版本留下的自动修改由独立、可续跑的审核任务统一还原。
"""
import base64
import copy
import hashlib
import math
import re
import tempfile
import threading
import time
import traceback
import unicodedata
from collections.abc import Callable
from contextlib import contextmanager
from pathlib import Path
from typing import Any
from uuid import uuid4

from core import (autofix, debug_log, file_problem, fix_images, page_breaks, pdf_store, print_layout, scan_analysis, shape_geometry,
                  sheet_history, workflow as wf)
from core.common import STALE_LOCK_HINT, is_inspection_exempt_sheet, workbook_locked, workbook_rel_name
from core.excel_ops import (
    CellProbe, close_saved_workbooks, enable_scan_view, normalize_workbook_view,
    open_workbooks_for_review, open_workbook_checked,
)
from core.excel_watchdog import ExcelWatchdog, idle_timeout
from core.hidden_text import (
    collect_text_pieces, is_excel_error, locate_text, normalize_text,
)
from core.pdf_detector import crossing_horizontal_line, fitz, render_crop_base64
from core.reporter import finish_execution_batch, init_execution_batch, update_file_progress

COLORS = ("blue", "green", "yellow")
KIND_LABELS = {"hidden": "隐藏截断", "clip": "压线", "pagebreak": "分页切开"}
CLIP_REASON = "文字被边框横线切穿，打印出来被切掉一部分（行高不够）"
WAITING_MESSAGE = "等待自动处理"
# Excel 行高按屏幕像素取整（100% 缩放 0.75pt、125% 0.6pt），写入后读回会差一点；
# 比这个差得少就算同一个行高（日文 Excel 默认 18.75pt，按 0.1 取整成 18.8 会被误判成「已被修改」）
ROW_HEIGHT_TOLERANCE = .4
XL_HALIGN_RIGHT, XL_HALIGN_CENTER = -4152, -4108
XL_HALIGN_GENERAL, XL_HALIGN_LEFT, XL_HALIGN_CENTER_ACROSS = 1, -4131, 7
XL_PATTERN_NONE, XL_PATTERN_SOLID = -4142, 1
# 只扫描模式在原 Excel 里给问题单元格标的颜色（可在页面上改）。
# 默认洋红底白字：业务表格几乎不用这个颜色，一眼就能看到，白字对比度约 6:1；黄色已有别人在用
DEFAULT_HIGHLIGHT = {"fill": "#C0007A", "font": "#FFFFFF"}
# 每种问题一种颜色（用户 2026-11-03 定），打开 Excel 就知道是哪种问题；一格有几种问题用「多种问题」的颜色。
# 都是页面原来的预设色，白字对比度都在 4.5:1 以上；「修复颜色」「查找残留标记色」也认得这些颜色
MULTI_TYPE = "multi"
DEFAULT_TYPE_HIGHLIGHT = {
    autofix.ERROR_TYPE_BORDER_CLIP: {"fill": "#C0007A", "font": "#FFFFFF"},  # 洋红
    autofix.ERROR_TYPE_TEXT_TRUNCATED: {"fill": "#1649CC", "font": "#FFFFFF"},  # 宝蓝
    autofix.ERROR_TYPE_BOX_OVERFLOW: {"fill": "#006633", "font": "#FFFFFF"},  # 墨绿
    autofix.ERROR_TYPE_PAGE_BREAK_CUT: {"fill": "#C2410C", "font": "#FFFFFF"},  # 橙红
    autofix.ERROR_TYPE_HASH_OVERFLOW: {"fill": "#6D28D9", "font": "#FFFFFF"},  # 深紫
    autofix.ERROR_TYPE_TEXT_MISSING: {"fill": "#333333", "font": "#FFFFFF"},  # 深灰
    MULTI_TYPE: {"fill": "#000000", "font": "#FFFF00"},  # 纯黑底黄字
}
SCAN_ONLY_MESSAGE = "已在 Excel 中标色，请人工修改并保存后复检"


def _data() -> dict[str, Any]:
    return wf.STATE.setdefault("batch_review", {"cards": {}, "rows": {}, "ignored": []})


def _marks() -> dict[str, dict]:
    """已在原 Excel 里标色的单元格（按卡片 id），记着原来的底色和字色，人工处理完再还原"""
    return _data().setdefault("highlights", {})


def _tabs() -> dict[str, dict[str, dict]]:
    """标成橙黄色的工作表标签：文件 -> {代码名（没有时用 "name:" + 表名）: 原来的标签颜色}，表里的标记全去掉后还原"""
    return _data().setdefault("tab_marks", {})


def _hex_color(value: Any) -> str | None:
    if isinstance(value, str) and len(value) == 7 and value[0] == "#":
        try:
            int(value[1:], 16)
        except ValueError:
            return None
        return value.upper()
    return None


def highlight_style(value: Any) -> dict[str, Any]:
    """页面传来的 {"types": {问题类型: {"fill": "#RRGGBB", "font": "#RRGGBB"}}}（旧版只有 fill / font 一种颜色）。
    缺了或格式不对用默认颜色。返回 {"fill", "font", "types"}：fill / font 是旧字段，留着兼容，按「边框压线」的颜色"""
    given = value.get("types") if isinstance(value, dict) and isinstance(value.get("types"), dict) else {}
    types: dict[str, dict[str, str]] = {}
    for kind, default in DEFAULT_TYPE_HIGHLIGHT.items():
        color = dict(default)
        if isinstance(given.get(kind), dict):
            for part in color:
                color[part] = _hex_color(given[kind].get(part)) or color[part]
        if color["fill"] == color["font"]:
            raise ValueError(f"「{type_color_label(kind)}」的标记底色和字色不能相同")
        types[kind] = color
    style: dict[str, Any] = dict(DEFAULT_HIGHLIGHT)
    if isinstance(value, dict) and not given:
        # 旧版页面只传一种颜色：按问题分色用默认颜色，fill / font 照旧记下
        for part in style:
            style[part] = _hex_color(value.get(part)) or style[part]
        if style["fill"] == style["font"]:
            raise ValueError("标记的底色和字色不能相同")
    else:
        style.update(types[autofix.ERROR_TYPE_BORDER_CLIP])
    style["types"] = types
    return style


def type_color_label(kind: str) -> str:
    return "多种问题" if kind == MULTI_TYPE else autofix.error_type_label(kind)


def _card_error_types(card: dict) -> list[str]:
    """这一格的问题类型，按 CELL_ERROR_SPECS 的顺序排好、去重"""
    order = list(autofix.CELL_ERROR_SPECS)
    types = [t for t in card.get("error_types") or [] if t in autofix.CELL_ERROR_SPECS]
    if not types:
        types = [autofix.item_error_type(p) for p in card.get("problems") or []] or [autofix.ERROR_TYPE_BORDER_CLIP]
    return sorted(dict.fromkeys(types), key=lambda t: order.index(t) if t in order else len(order))


def _card_severity(card: dict, items: list[dict] | None = None) -> str:
    """这一格按「字丢没丢」算哪一级（red / gray / yellow）。按这格现在还没解决的检测结果算，
    卡片上记的问题类型会累积历史（以前截断、改好后只剩压线），只在没有检测结果时才用它。调用方持有 STATE_LOCK"""
    items = _items_at(card) if items is None else items
    live = [autofix.item_error_type(it) for it in items if it.get("status") not in wf.RESOLVED_STATUSES]
    return autofix.error_severity(live or _card_error_types(card))


def _items_by_cell(paths: set[str] | None = None) -> dict[tuple[str, str, str], list[dict]]:
    """检测结果按单元格分好（只取这些文件的），一次整理、多张卡共用，免得每张卡都把全部结果过一遍"""
    by_cell: dict[tuple[str, str, str], list[dict]] = {}
    for item in wf.STATE["items"].values():
        if paths is None or item["wb_path"] in paths:
            by_cell.setdefault(_cell(item), []).append(item)
    return by_cell


def _hidden_low_risk(card: dict, items: list[dict] | None = None) -> bool:
    """黄色（只是被裁，字都在）且没打开「也列出来」：不列卡片、不标色、不算待处理。调用方持有 STATE_LOCK"""
    return autofix.low_risk_hidden(wf.STATE) and _card_severity(card, items) == autofix.SEVERITY_YELLOW


def set_show_low_risk(show: Any) -> bool:
    """查看切字页的开关：黄色（被裁）也列出来。Excel 里的标色在下次复检 / 打开时跟着更新"""
    if not isinstance(show, bool):
        raise ValueError("请给出开或关")
    with wf.STATE_LOCK:
        _data()["show_low_risk"] = show
        for path in {it["wb_path"] for it in wf.STATE["items"].values()}:
            wf._sync_problem_counts(path, "按新的分级重新统计")
        wf._persist_state(force=True)
    debug_log.event("🟡 黄色也列出来", 开=show)
    return show


def _mark_type(card: dict) -> str:
    """这一格按哪种问题上色：只有一种问题就是那种，几种问题是 multi"""
    types = _card_error_types(card)
    return types[0] if len(types) == 1 else MULTI_TYPE


def _mark_color(style: dict | None, card: dict) -> dict[str, int]:
    """这一格的标记颜色（Excel 的 BGR 整数）。旧版记下的样式没有 types，也按默认的分色"""
    kind = _mark_type(card)
    types = (style or {}).get("types") or DEFAULT_TYPE_HIGHLIGHT
    color = types.get(kind) or DEFAULT_TYPE_HIGHLIGHT[kind]
    return {"fill": _excel_color(color["fill"]), "font": _excel_color(color["font"])}


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
                         _key(path, new, int(row.get("excel_row") or row["row"]), "break") if row.get("kind") == "break" else
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


def _latest_shot(card: dict, items: list[dict], cache: dict[str, Any] | None = None) -> str:
    """本行没改、因别处调整而通过的格子：从刚导出的 PDF 重新截图，不沿用之前尝试或检测时的旧图。
    文字太短（容易对到别的格子）或找不到时不截图，宁可显示「暂无截图」也不放错的图。
    cache：同一份 PDF 的文字只读一次（{pdf 路径: 文字块}），一份文件复检通过几十格时不用每格重读整份 PDF。"""
    if len(normalize_text(card.get("text", ""))) < 4:
        return ""
    for pdf_path in dict.fromkeys(it.get("pdf_path") for it in items if it.get("pdf_path")):
        try:
            with fitz.open(pdf_path) as doc:
                if cache is None:
                    pieces = collect_text_pieces(doc)
                else:
                    if pdf_path not in cache:
                        cache[pdf_path] = collect_text_pieces(doc)
                    pieces = cache[pdf_path]
                found = locate_text(pieces, card["text"])
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
    """点了「完成当前文件后暂停」：做完当前这份文件就停"""
    return bool(job.get("pause_requested"))


# 主页单项任务里「做完一份就不再碰」的：做完的文件可以先打开、先在调整页面改（精简 Sheet 做完还要自动检查，先不放）
STREAM_QUICK_TASKS = {"export_pdf", "normalize_view"}


def open_block_reason(job: dict, path: Path) -> str:
    """后台任务在跑时，网页上点这份文件能不能在 Excel 打开：能打开返回空字符串，不能就返回原因（页面上直接显示）。

    还没轮到、正在处理的文件要等：用户在 Excel 里开着，后台轮到它时会因为文件被占用跳过。
    已经做完、后面不再碰的文件，和不在这次任务里的文件，可以打开。调用方持有 STATE_LOCK。"""
    kind = job.get("kind")
    if kind in STREAM_QUICK_TASKS and "workbooks" in job:
        # 导出 PDF / 视图标准化：一份份往后做，做完（或出错跳过）的这次不会再碰，可以先打开（边导边看，用户 2026-11-12 定）
        target = path.resolve()
        mine = [p for p in job["workbooks"] if Path(p).resolve() == target]
        done = set(job.get("completed", [])) | set(job.get("failed", []))
        if not mine or all(p in done for p in mine):
            return ""
        pending = [p for p in job["workbooks"] if p not in done]
        if pending and pending[0] in mine:
            return "后台正在处理这份文件，处理完才能打开"
        return "还没轮到这份文件，处理完才能打开"
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
    if any(p in job.get("deferred_adjustments", {}) for p in mine):
        return "这份文件还有没写完的修改，后台还要回来处理"
    if all(p in completed for p in mine):
        return ""
    scanned = job.get("scanned", [])
    pending = [p for p in job["workbooks"] if p not in scanned]
    if pending and pending[0] in mine:
        return "后台正在处理这份文件，处理完才能打开"
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
    return (not card.get("layout_stale") and not card.get("verified") and card.get("after_kind") == "attempt" and not _waiting(card)
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


CARD_DETAIL_FIELDS = {"text", "problems", "fix_images", "context", "images", "message"}


# ───────────────────── 按分类人工处理（用户 2026-11-03：留排版、去掉学一下） ─────────────────────

CATEGORY_TYPE_ORDER = list(autofix.CELL_ERROR_SPECS)


def _missing_lines(items: list[dict]) -> int | None:
    found = [int(it["missing_lines"]) for it in items
             if isinstance(it.get("missing_lines"), (int, float)) and not isinstance(it.get("missing_lines"), bool)
             and it["missing_lines"] > 0]
    return max(found) if found else None


def _card_category(card: dict, items: list[dict]) -> tuple[str, str]:
    """返回 (分类键, 给人看的名字)，例：「文字截断 · 合并格 · 缺 1 行」。
    一格有几种问题时按 CELL_ERROR_SPECS 的顺序取第一种。只用来分组显示，不改格子。"""
    types = [t for t in card.get("error_types") or [] if t in autofix.CELL_ERROR_SPECS]
    if not types:
        types = [autofix.item_error_type(it) for it in items] or [autofix.ERROR_TYPE_BORDER_CLIP]
    kind = min(types, key=lambda t: CATEGORY_TYPE_ORDER.index(t) if t in CATEGORY_TYPE_ORDER else len(CATEGORY_TYPE_ORDER))
    parts, names = [kind], [autofix.error_type_label(kind)]
    if kind != autofix.ERROR_TYPE_TEXT_MISSING:
        grown = card.get("grown") or {}
        merged = grown["merged"] if "merged" in grown else ":" in (card.get("cell_addr") or "")
        parts.append("merged" if merged else "single")
        names.append("合并格" if merged else "单格")
    if kind == autofix.ERROR_TYPE_TEXT_TRUNCATED:
        lines = _missing_lines(items)
        if lines:
            parts.append(f"lines{min(lines, 3)}")
            names.append(f"缺 {lines} 行" if lines < 3 else "缺 3 行以上")
    return "|".join(parts), " · ".join(names)


def snapshot(detail_paths: set[str] | None = None) -> dict:
    from core import cell_toolbox  # 它也引用本模块，放这里免得循环引用
    with wf.STATE_LOCK:
        cards = []
        by_cell: dict[tuple[str, str, str], list[dict]] = {}
        for item in wf.STATE["items"].values():
            by_cell.setdefault(_cell(item), []).append(item)
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

        hide_low_risk, low_risk_hidden = autofix.low_risk_hidden(wf.STATE), 0
        for card in sorted(_data()["cards"].values(), key=lambda c: c["seq"]):
            if card.get("dismissed") or not in_folder(card["wb_path"]):
                continue
            items = by_cell.get(_cell(card), [])
            severity = _card_severity(card, items)
            if severity == autofix.SEVERITY_YELLOW and hide_low_risk:
                low_risk_hidden += 1
                continue
            # 释放 STATE_LOCK 后才序列化 JSON，返回值仍需与后台线程的状态隔离。
            detailed = detail_paths is None or card["wb_path"] in detail_paths
            omitted = {"signatures", "item_ids"} | (set() if detailed else CARD_DETAIL_FIELDS)
            result = copy.deepcopy({k: v for k, v in card.items() if k not in omitted})
            if detailed and not result.get("problems"):  # 旧版本留下的卡片没存问题说明，按单元格现有的问题项补上
                result["problems"] = list({p["kind"]: p for p in map(_problem, items)}.values())
            if not result.get("error_types"):
                result["error_types"] = list(dict.fromkeys(
                    p.get("error_type") or autofix.item_error_type({"kind": p.get("kind"), "hidden_reason": p.get("reason")})
                    for p in (result.get("problems") or card.get("problems") or map(_problem, items))
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
            result["severity"] = severity
            result["category"] = list(_card_category(result, items))
            result["cell_text"] = str(card.get("text") or "")[:60]  # 分类清单用，文件框收起时 text 不发
            result["adoptable"] = _adoptable(card)
            fix = (_data().get("cell_fixes") or {}).get(card["id"])
            if fix:
                result["cell_fix"] = dict(fix)  # 用过单元格工具箱的什么工具（没过时网页上写出来）
            # 单元格工具箱：这格能用哪些工具；切字的默认「行高 ×1.5」，用过工具没过的不默认（不自动升一级）
            result["tools"] = cell_toolbox.allowed_tools(result)
            result["default_tool"] = ("grow:1.5" if not fix and "grow" in result["tools"]
                                      and set(result["error_types"]) & set(GROW_ERROR_TYPES) else "none")
            result["highlighted"] = _shows_mark(_marks().get(card["id"]))
            if detailed:
                result["images"] = {
                    which: url for which, name in card.get("fix_images", {}).items()
                    if (url := image_url(name))
                }
                if result.get("context"):
                    context = result["context"]
                    name = context["image_file"]
                    context["image"] = image_url(name)
            else:
                result["details_pending"] = True
                result["waiting"] = _waiting(card)
            cards.append(result)
        highlighted = sum(1 for m in _marks().values() if in_folder(m["path"]) and _shows_mark(m))
        tabs = sum(len(t) for p, t in _tabs().items() if in_folder(p))
        return {"cards": cards, "busy": wf.STATE["busy"], "folder": wf.STATE["folder"], "highlighted": highlighted,
                "scan_analysis": copy.deepcopy(wf.STATE.get("latest_scan_analysis")),
                "tab_marks": tabs,
                "error_types": list(autofix.CELL_ERROR_SPECS.values()),
                "show_low_risk": not hide_low_risk, "low_risk_hidden": low_risk_hidden,
                "open_blocked": _open_blocked([c["wb_path"] for c in cards]),
                "followups": copy.deepcopy(_data().get("followups") or []),
                "opened_files": sorted(_opened_paths()),
                "recheck_ready": _recheck_ready([c["wb_path"] for c in cards]),
                "scan_eta": _scan_eta(wf.STATE["job"]) if wf.STATE["job"] else None,
                "slow_files": [{"path": p, "name": Path(p).name, **info} for p, info in _data().get("slow_files", {}).items()
                               if in_folder(p)],
                "job": copy.deepcopy(wf.STATE["job"]), "progress": dict(wf.STATE["scan_progress"]),
                "feed": _job_feed(wf.STATE["job"]),
                "logs": wf.STATE["logs"][-100:]}


PAGE_VIEW_ZOOM = 2.0  # 大图占满屏宽也清楚（A4 横向约 1680 像素宽）
PAGE_VIEW_BOX_LIMIT = 300


def page_view(card_id: Any) -> dict[str, Any]:
    """查看切字「看大图」（用户 2026-11-11 要）：这一格所在的 PDF 整页，打开时才渲染（扫描不画，不拖慢扫描），
    加上这一页所有还没解决的问题格位置，网页按红 / 灰 / 黄画框、写原因，一眼看出为什么标色、问题出在哪。"""
    if not isinstance(card_id, str):
        raise ValueError("请给出格子")
    hide = None
    with wf.STATE_LOCK:
        card = _data()["cards"].get(card_id)
        if card is None or card.get("dismissed"):
            raise ValueError("这一格已经不在列表里了，刷新一下再看")
        mine = _items_at(card)
        live = [it for it in mine if it.get("status") not in wf.RESOLVED_STATUSES] or mine
        item = next((it for it in live if it.get("pdf_path") and it.get("page_index") is not None
                     and isinstance(it.get("bbox"), (list, tuple)) and len(it["bbox"]) == 4), None)
        if item is None:
            raise ValueError("这一格在 PDF 里没定位到（PDF 未匹配到文字），没有大图，请打开 Excel 看")
        pdf_path, page_index = item["pdf_path"], int(item["page_index"])
        hide = autofix.low_risk_hidden(wf.STATE)
        boxes = []
        for it in wf.STATE["items"].values():
            if (it.get("pdf_path") != pdf_path or it.get("page_index") != page_index
                    or not isinstance(it.get("bbox"), (list, tuple)) or len(it["bbox"]) != 4):
                continue
            current = _cell(it) == _cell(card)
            if it.get("status") in wf.RESOLVED_STATUSES and not current:
                continue
            kind = autofix.item_error_type(it)
            severity = autofix.error_severity([kind])
            if hide and severity == autofix.SEVERITY_YELLOW and not current:
                continue  # 黄色默认藏起来的，大图上也不画，免得满页都是框
            problem = _problem(it)
            boxes.append({"bbox": [float(v) for v in it["bbox"]], "severity": severity, "current": current,
                          "addr": it.get("cell_addr") or "", "label": autofix.error_type_label(kind),
                          "reason": problem["reason"] or "", "missing": problem["missing"],
                          "text": str(it.get("text") or "")[:200],
                          "visible": str(it.get("visible_text") or "")[:200] if it.get("kind") == "hidden" else ""})
        boxes.sort(key=lambda b: (not b["current"], b["bbox"][1], b["bbox"][0]))
        boxes = boxes[:PAGE_VIEW_BOX_LIMIT]
        title = f"{card.get('wb_name') or Path(card['wb_path']).name} / {card['sheet_name']}!{card.get('cell_addr', '')}"
    if not Path(pdf_path).is_file():
        raise ValueError("这份 PDF 已经不在了（可能复检后重新导出过），复检一次再看")
    started = time.perf_counter()
    with fitz.open(pdf_path) as doc:
        if not 0 <= page_index < len(doc):
            raise ValueError("PDF 页数变了，复检一次再看")
        page = doc[page_index]
        png = page.get_pixmap(matrix=fitz.Matrix(PAGE_VIEW_ZOOM, PAGE_VIEW_ZOOM), annots=False).tobytes("png")
        width, height, pages = float(page.rect.width), float(page.rect.height), len(doc)
        x0, y0 = float(page.rect.x0), float(page.rect.y0)
    for box in boxes:  # 换成相对整页的比例，网页按百分比画框，图多大都对得上
        b = box["bbox"]
        box["rect"] = [round((b[0] - x0) / width, 5), round((b[1] - y0) / height, 5),
                       round((b[2] - x0) / width, 5), round((b[3] - y0) / height, 5)]
        del box["bbox"]
    debug_log.event("🖼 看大图", 格=title, 页=page_index + 1, 框=len(boxes), 秒=round(time.perf_counter() - started, 2))
    return {"title": title, "page": page_index + 1, "pages": pages, "width": width, "height": height,
            "image": "data:image/png;base64," + base64.b64encode(png).decode("ascii"), "boxes": boxes,
            "low_risk_hidden": bool(hide)}


CARD_IMAGES_LIMIT = 200


def card_images(ids: Any) -> dict[str, dict[str, str]]:
    """单元格工具箱按需取截图：{卡片 id: {"before": 检测时的 PDF 截图, "after": 最近一次截图}}"""
    if not isinstance(ids, list) or any(not isinstance(cid, str) for cid in ids):
        raise ValueError("请给出卡片")
    with wf.STATE_LOCK:
        cards = _data()["cards"]
        names = {cid: dict(cards[cid].get("fix_images") or {}) for cid in ids[:CARD_IMAGES_LIMIT] if cid in cards}
    return {cid: {which: _image_url(name) for which, name in shots.items() if fix_images.image_path(name)}
            for cid, shots in names.items()}


_FEED_COUNTS: dict[str, int] = {}  # 扫描时只取进度：做完的文件查出几处，每份只数一次（不每次把所有卡片过一遍）


def _cards_left(paths: list[str], light: bool) -> dict[str, int]:
    """每份文件还剩几处要人工看。light（扫描中只取进度）时只数新做完的、记住结果"""
    cards = _data().get("cards") or {}
    if not light:
        left: dict[str, int] = {}
        by_cell = _items_by_cell()
        for card in cards.values():
            if not card.get("dismissed") and not _hidden_low_risk(card, by_cell.get(_cell(card), [])):
                left[card["wb_path"]] = left.get(card["wb_path"], 0) + 1
        return left
    new = {p for p in paths if p not in _FEED_COUNTS}
    if new:
        for p in new:
            _FEED_COUNTS[p] = 0
        by_cell = _items_by_cell(new)
        for card in cards.values():
            if card.get("wb_path") in new and not card.get("dismissed") and not _hidden_low_risk(card, by_cell.get(_cell(card), [])):
                _FEED_COUNTS[card["wb_path"]] += 1
    return _FEED_COUNTS


def _job_feed(job: dict | None, light: bool = False) -> dict:
    """调用方持有 STATE_LOCK。查看切字页的实时进度框（用户 2026-11-09 要：所有批量操作都要实时进度框）：
    正在处理哪份、做到哪一步、每份做完的结果（✓ 做好 / ! 出错或还要再看和原因），最新的放最后"""
    if not job or not isinstance(job.get("workbooks"), list):
        return {"current": "", "step": "", "log": []}
    kind, folder = job.get("kind", ""), wf.STATE.get("folder") or ""

    def label(path: str) -> str:
        try:
            return workbook_rel_name(Path(path), folder) if folder else Path(path).name
        except Exception:
            return Path(path).name

    results, errors = job.get("result") or {}, job.get("file_errors") or {}
    completed = list(dict.fromkeys(job.get("completed") or []))
    if not light:
        _FEED_COUNTS.clear()
    left = _cards_left(completed, light) if kind in {"batch_scan", "batch_recheck"} else {}
    log: list[dict] = []
    for path in completed:
        r = results.get(path) or {}
        ok, reason = True, "做完了"
        if kind == "batch_fix":
            pages = len(r.get("pages") or [])
            ok = not r.get("error") and not r.get("errors") and not r.get("left")
            reason = r.get("error") or (f"改了 {r.get('done', 0)} 格，复检通过 {r.get('passed', 0)} 格"
                                        + (f"，还没过 {r['left']} 格" if r.get("left") else "")
                                        + (f"，{len(r['errors'])} 处出错" if r.get("errors") else "")
                                        + (f"，{pages} 张表页数变了" if pages else ""))
        elif kind == "batch_unhighlight":
            ok = not r.get("error") and not r.get("leftover_failed")
            reason = r.get("error") or f"清掉 {len(r.get('cells') or [])} 格标色" + ("，有删不掉的标记" if r.get("leftover_failed") else "")
        elif kind == "batch_scan":
            reason = f"查出 {left[path]} 处要人工看" if left.get(path) else "没查出问题"
        elif kind == "batch_recheck":
            reason = f"复检完，还剩 {left[path]} 处" if left.get(path) else "复检全部通过"
        elif kind == "batch_open":
            reason = "已在 Excel 打开"
        log.append({"name": label(path), "path": path, "ok": ok, "reason": reason, "sheet": ""})
    for path, message in errors.items():
        if path not in completed:
            log.append({"name": label(path), "path": path, "ok": False, "reason": str(message), "sheet": ""})
    current = job.get("current") or ""
    if kind == "batch_recheck":
        for path in job.get("skipped") or []:
            if path != current and path not in completed and path not in errors:
                log.append({"name": label(path), "path": path, "ok": False,
                            "reason": "没复检成功（多半是文件没保存关闭），详情看下面的日志", "sheet": ""})
    step = ""
    if wf.STATE["busy"]:
        if kind == "batch_scan" and wf.SCAN_NOW.get("path") in job["workbooks"]:
            current = wf.SCAN_NOW["path"]
            if wf.SCAN_NOW.get("sheet"):
                step = f"检查第 {wf.SCAN_NOW.get('index', 0)} / {wf.SCAN_NOW.get('total', 0)} 张表：{wf.SCAN_NOW['sheet']}"
        if not current:
            current = next((p for p in job["workbooks"] if p not in completed and p not in errors), "")
        if not step and wf.STATE["logs"]:
            step = str(wf.STATE["logs"][-1])
    return {"current": label(current) if current else "", "step": step, "log": log[-500:]}


def progress_snapshot() -> dict:
    """扫描期间网页每 1.8 秒只取这些：进度条、当前文件、预计时间和实时日志。
    卡片、标色数等完整内容扫描时不变或用户不看，点「刷新」或扫描结束时再取完整的（用户 2026-11-03 定），
    免得每次刷新都把几百个文件的卡片整理一遍、拖慢扫描。"""
    with wf.STATE_LOCK:
        job = wf.STATE["job"]
        blocked = {}
        if wf.STATE["busy"] and job and job.get("kind") == "batch_scan":
            # 任务里都是规范化的完整路径，直接用集合更新按钮状态，不遍历卡片或逐次解析文件路径。
            completed, scanned = set(job.get("completed", [])), set(job.get("scanned", []))
            current = next((p for p in job.get("workbooks", []) if p not in scanned), None)
            for path in job.get("workbooks", []):
                if path in job.get("deferred_adjustments", {}):
                    blocked[path] = "这份文件还有没写完的修改，后台还要回来处理"
                elif path not in completed:
                    blocked[path] = ("后台正在处理这份文件，处理完才能打开" if path == current
                                     else "还没轮到这份文件，处理完才能打开")
        return {"light": True, "busy": wf.STATE["busy"], "folder": wf.STATE["folder"],
                "job": copy.deepcopy(job), "progress": dict(wf.STATE["scan_progress"]),
                "open_blocked": blocked,
                "scan_eta": _scan_eta(job) if job else None,
                "slow_files": [{"path": p, "name": Path(p).name, **info} for p, info in _data().get("slow_files", {}).items()],
                "feed": _job_feed(job, light=True),
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
    """程序自己保存了文件（写行高、标色）：打印方向记录跟着更新，不算你改过文件"""
    with wf.STATE_LOCK:
        print_layout.followed_save(wf.STATE.setdefault("print_layouts", {}), path, before)


def _adjust_pdf_digest(path: str) -> str | None:
    """调整页面给这份 Excel 导出过 PDF 时，保存前的哈希（保存后跟进导出记录用）；没导出过返回 None，不用算"""
    root = wf.STATE.get("folder")
    try:
        if root and pdf_store.has_record(root, Path(path)):
            return sheet_history.fingerprint(Path(path))
    except Exception as err:
        debug_log.event("⚠️ 读调整页面的 PDF 记录失败", 文件=Path(path).name, 错误=str(err))
    return None


def _follow_marks_save(path: str, before: Any, pdf_digest: str | None) -> None:
    """标色 / 去标色存进 Excel 后：只是条件格式变了，格子位置、行高、分页都没变。
    - 检测基线跟着前进（以前保守按改过处理：下次扫描整份重查、单元格工具箱也要先复检才能用，
      调整页面里自动查切字时每次都会把标过色的文件全部重扫一遍；用户 2026-11-09 把切字并进调整页面后改）。
      保存前文件就已经不是检测时的版本（人改过）的不跟。
    - 调整页面导出的 PDF 记录也跟着记成新的哈希，回到调整页面不会当成 Excel 改过、整份重新导出"""
    try:
        after = wf._file_version(Path(path))
    except OSError:
        return
    with wf.STATE_LOCK:
        if before is not None and wf.STATE["checked_files"].get(path) == before:
            wf.STATE["checked_files"][path] = after
            checks = _sheet_checks().get(path, {})
            for name, checked in list(checks.items()):
                if checked == before:
                    checks[name] = after
    if pdf_digest:
        try:
            pdf_store.follow_save(wf.STATE.get("folder"), Path(path), pdf_digest,
                                  sheet_history.fingerprint(Path(path)))
        except Exception as err:
            debug_log.event("⚠️ PDF 记录没能跟上标色保存", 文件=Path(path).name, 错误=str(err))


def _save_scan_preview(excel: Any, path: str, session: Any) -> None:
    """扫描的只读工作簿已关闭，另开原文件保存视图，并跟进只涉及视图的版本变化。"""
    _check_file(path)
    before = wf._file_version(Path(path))
    if wf.STATE["checked_files"].get(path) != before:
        raise ValueError(f"扫描后文件已修改，请重新扫描：{Path(path).name}")
    after = wf._save_scan_view(excel, Path(path), before)
    with wf.STATE_LOCK:
        wf.STATE["checked_files"][path] = after
        checks = _sheet_checks().get(path, {})
        for name, checked in list(checks.items()):
            if checked == before:
                checks[name] = after
        _followed_own_save(path, before)
        session.view_saved(path, before, after)
        wf._persist_state()


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
        shapes_before = shape_geometry.snapshot_book(book)  # 改行高列宽前记下图形，保存前放回（用户 2026-11-11 要）
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
            shape_geometry.restore_book(book, shapes_before, f"写入行高 {Path(path).name}")
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
                break_row = next((r["row"] for r in pending["rows"] if r.get("kind") == "break"), card.get("excel_row"))
                card.update(color="blue", verified=True, after_kind="fixed",
                            message=(f"已在第 {break_row} 行上方插入分页并保存，等待审核"
                                     if break_row and break_row != card.get("excel_row")
                                     else "已在合并区域上方插入分页并保存，等待审核"))
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
    view: dict | None = None, compact_print_tail: bool = False, before_print_save: Callable | None = None,
) -> list[str]:
    """重新检测（默认整个文件）。session 给了就沿用同一状态已导出的 PDF（reuse=False 时一律重新导出）；
    workbook 给了就直接用它检测（全自动的副本，或复检时在内存里还原过标记色的原文件），不再打开原文件。
    view 给了就顺便查原文件视图，view["needed"] 表示是否还要另存分页预览。
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
        scan_view="check_scan_view" if view is not None else None,
        compact_print_tail=compact_print_tail, before_print_save=before_print_save,
    )
    if view is not None:
        view["needed"] = entry.get("scan_view_needed", True)
    with wf.STATE_LOCK:
        missing = list(entry.get("missing", []))
        for card in _data()["cards"].values():
            if card["wb_path"] == path and is_inspection_exempt_sheet(card["sheet_name"]):
                card.update(dismissed=True, waiting=False, verified=False,
                            message="命中表名免检规则，免检")
                continue
            if card["wb_path"] == path and card["sheet_name"] not in missing and (
                    sheets is None or card["sheet_name"] in entry["sheets"]):
                card.pop("layout_stale", None)
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


def _text_width(text: str) -> int:
    """按显示宽度数字数：全角字（日文、中文）算 2，半角字母数字算 1"""
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)


DIRECT_LINE_PT = 12.0  # 读不到字号时的单行字高（约 10pt 字体 × 1.2）
LINE_HEIGHT_RATIO = 1.2  # 一行字高 = 字号 × 1.2（Excel 默认行距；只加 1 倍字号时上下仍常被切）


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


def _restart_excel(session: Any) -> Any:
    """后台 Excel 可能已经崩溃或卡住：关掉整个会话（含副本），重新开一个"""
    session.close()
    return session.app()


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


def _recheck_file(excel: Any, path: str, cards: list[dict], *, passed_message: str = "人工修改后复检通过",
                  found_message: str = "复检查出的问题：请修改并保存后再复检") -> None:
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
    def shot_sources(card: dict) -> list[dict]:
        # 调用方持有 STATE_LOCK
        return [p for p in wf.STATE["pages"].values()
                if p.get("wb_path") == path and p.get("sheet_name") == card["sheet_name"]] + _items_at(card)

    # 通过的格子要从刚导出的 PDF 截修复后的图：截图慢（几十格要十几秒），放在锁外做，网页取进度不用等
    with wf.STATE_LOCK:
        problem_cells = {_cell(it) for it in wf.STATE["items"].values() if it["wb_path"] == path
                         and it["status"] not in wf.RESOLVED_STATUSES}
        todo = [(dict(card), [{"pdf_path": p.get("pdf_path")} for p in shot_sources(card)]) for card in cards
                if not is_inspection_exempt_sheet(card["sheet_name"]) and _cell(card) not in problem_cells]
    shot_started, cache = time.perf_counter(), {}
    shots = {card["id"]: _latest_shot(card, sources, cache) for card, sources in todo}
    if todo:
        debug_log.event("📸 复检通过截图", 文件=Path(path).name, 格数=len(todo),
                        秒=round(time.perf_counter() - shot_started, 2))
    with wf.STATE_LOCK:
        unresolved = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                      and it["status"] not in wf.RESOLVED_STATUSES]
        problem_cells = {_cell(it) for it in unresolved}
        passed = 0
        for card in cards:
            if is_inspection_exempt_sheet(card["sheet_name"]):
                continue
            if _cell(card) in problem_cells:
                continue
            passed += 1
            # 这一格这次没查出问题：翻篇，移出列表。报告里的修复后截图用刚导出的 PDF 按文字重新截，截不到就不放
            prior = _items_at(card)
            fix_images.forget(card, "after")
            image = shots[card["id"]] if card["id"] in shots else _latest_shot(card, shot_sources(card), cache)
            _save_shot(card, "after", image)
            after = card.get("fix_images", {}).get("after")
            for item in prior:
                if after:
                    item.setdefault("fix_images", {})["after"] = after
                else:
                    fix_images.forget(item, "after")
            message = card.pop("horizontal_fix", None) or passed_message
            card.update(color="blue", verified=True, finalized=True, dismissed=True, waiting=False,
                        after_kind="fixed", message=message)
        found = 0
        for group in _group_cells(unresolved):
            card = _card_for(group)
            if id(card) in waiting or (card.get("marked") and card["color"] != "yellow" and not card.get("finalized")):
                continue  # 全自动还没轮到的，或你已标蓝 / 绿还没提交的，保留原样
            found += 1
            card.pop("horizontal_fix", None)
            # 旧扫描留下的「疑似跨格横排」记号：扫描已不自动加高，复检查出的问题一律当普通人工项、照常标色
            # （2026-11-06 用户实测：带这个记号的格子在「打开下一批」里却没标色，标签也被改回原色）
            card.pop("horizontal_layout", None)
            message = found_message if card.get("grown") else "复检查出的问题：请修改并保存后再复检"
            card.update(color="yellow", verified=False, manual=True, finalized=True, waiting=False,
                        dismissed=False, message=message)
            for item in group:
                item.update(status="manual_review", batch_managed=True, triage_keep=False, review_reason=message)
            if _hidden_low_risk(card, group):
                found -= 1  # 只是被裁（字都在）：默认不列，不算查出的问题
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


@print_layout.with_layout
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
    if is_inspection_exempt_sheet(card["sheet_name"]):
        _refresh(excel, path, [card["sheet_name"]])
        wf.log_msg(f"ℹ️ {Path(path).name} {label}：命中表名免检规则，免检")
        return
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
    print_layout.bind(book, path)
    try:
        timing["open"] = time.perf_counter() - started
        sheet, box = _current_area(book, path, card)
        wf.prepare_print_layout(sheet, path)
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


# 扫描后的简单处理（用户 2026-11-02 定，2026-11-03 去掉按分类学习后恢复）：切字的格子（边框压线、文字截断、
# 超出裁剪框）开自动换行，按文字长度和格子宽度估算换行后要几行，行高不够先补到够，再加「字号 × 倍数」，保存。
# 保存后整份复检一次：通过的格子翻篇、不标色；没通过的不再改，标色交给人工。
# 分页切开、### 列宽不足、PDF 未匹配到文字只标色。
GROW_ERROR_TYPES = autofix.CUT_ERROR_TYPES
GROW_RATIOS = (1.5, 1.2)
DEFAULT_GROW_RATIO = 1.5
DEFAULT_FONT_PT = 11.0  # 读不到字号时按日文 Excel 默认的 11pt
LINE_HEIGHT_RATIO = 1.25  # 估算换行后的高度：每行按「字号 × 1.25」pt（MS Pゴシック 11pt 一行 13.5pt）
HALF_WIDTH_RATIO = 0.55  # 半角字符（英数、SQL）宽度约为字号的 0.55 倍，全角按 1 倍
CELL_PADDING_PT = 4.0  # 格子左右留白
HASH_MESSAGE = "### 列宽不足：没有调整列宽，请人工核对数字格式、日期或列宽，也可标绿忽略"
HORIZONTAL_LAYOUT_MESSAGE = "疑似跨格横排：右侧有连续空格，已保留原排版。请核对书式的应用范围，保存后复检"
HORIZONTAL_FORMAT_PASSED = "已向右延伸条件格式，横排文字复检通过"
HORIZONTAL_MARK_PASSED = "已清除截断横排文字的临时标记，复检通过"
GROWN_MESSAGE = "已开自动换行、加高并保存，效果请人工看：好了标蓝，不好就改了保存再复检"
GROWN_PASSED_MESSAGE = "扫描后已开自动换行、加高，复检通过"
GROWN_FAILED_MESSAGE = "已开自动换行、加高并复检，还有问题：请人工修改并保存后再复检"


def grow_ratio(value: Any = None) -> float:
    """加高倍数：每个切字的格子所在行加「字号 × 倍数」pt。没给按 1.5，只能选 1.5 或 1.2"""
    if value is None:
        return DEFAULT_GROW_RATIO
    try:
        ratio = float(value)
    except (TypeError, ValueError):
        ratio = 0.0
    if isinstance(value, bool) or ratio not in GROW_RATIOS:
        raise ValueError("加高倍数只能选 1.5 或 1.2")
    return ratio


def _keeps_user_mark(card: dict) -> bool:
    """你已经手动标成蓝 / 绿、还没提交的卡：保留你的选择"""
    return bool(card.get("marked") and card["color"] != "yellow" and not card.get("finalized"))


def _grow_targets(path: str) -> list[dict]:
    """调用方持有 STATE_LOCK：本文件要开换行、加高的卡（切字、定位已确认）"""
    pending = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
               and it["status"] not in wf.RESOLVED_STATUSES]
    targets = []
    for group in _group_cells(pending):
        card = _card_for(group)
        if card.get("dismissed") or _keeps_user_mark(card) or not _can_highlight(card, group):
            continue
        if any(autofix.item_error_type(it) in GROW_ERROR_TYPES for it in group):
            targets.append(card)
    return targets


def _font_size(area: Any) -> float:
    try:
        size = area.Cells(1, 1).Font.Size
    except Exception:
        return DEFAULT_FONT_PT
    if isinstance(size, bool) or not isinstance(size, (int, float)) or not 4.0 <= float(size) <= 72.0:
        return DEFAULT_FONT_PT  # 合并区域里字号不一样时 Excel 返回 None
    return float(size)


def _text_width_pt(text: str, size: float) -> float:
    """一段文字不换行时大约多宽（pt）：全角按字号，半角按字号 × 0.55"""
    return sum(size if unicodedata.east_asian_width(ch) in "WFA" else size * HALF_WIDTH_RATIO for ch in text)


def _has_horizontal_space(sheet: Any, area: Any, size: float) -> bool:
    """识别极窄单格中的横排单行文字，且右侧确实有足够连续的空白宽度。"""
    try:
        if (area.WrapText is not False or area.Rows.Count != 1 or area.Columns.Count != 1
                or area.Orientation not in (0, XL_HORIZONTAL)
                or area.HorizontalAlignment not in (XL_HALIGN_GENERAL, XL_HALIGN_LEFT, XL_HALIGN_CENTER_ACROSS)):
            return False
        width = float(area.Width)
        if not 0 < width <= 3 * size + CELL_PADDING_PT or float(area.Height) < size:
            return False
        text = area.Cells(1, 1).Value
        if not isinstance(text, str) or not text.strip() or "\n" in text or "\r" in text:
            return False
        needed = _text_width_pt(text.rstrip(), size) + CELL_PADDING_PT
        if needed < 3 * width:
            return False
        row, col = int(area.Row), int(area.Column)
        # 只检查标题附近。值、公式、合并格和隐藏列都会终止这段可用宽度。
        last = min(int(sheet.Columns.Count), col + 64)
        for c in range(col + 1, last + 1):
            neighbor = sheet.Cells(row, c)
            if neighbor.Value2 is not None or neighbor.HasFormula or neighbor.MergeCells:
                break
            extra = float(neighbor.Width)
            if extra <= 0:
                break
            width += extra
            if width >= needed:
                return True
    except Exception:
        return False
    return False


def _visible_horizontal_prefix(doc: Any, text: str, anchor: list[float]) -> int:
    """只比较原位置同一行的可见文字，不把别处的同文或裁剪框外的字算进去。"""
    x, y, _, bottom = anchor
    tolerance = max(1.0, (bottom - y) * 0.2)
    matches = [piece for piece in collect_text_pieces(doc)
               if abs(piece.bbox[0] - x) <= tolerance and abs(piece.bbox[1] - y) <= tolerance]
    if len(matches) != 1:
        return 0
    piece = matches[0]
    visible = []
    for char, box, shown in piece.chars:
        if not shown or crossing_horizontal_line(box, piece.rules) is not None:
            break
        visible.append(char)
    prefix = normalize_text("".join(visible))
    for i, (expected, actual) in enumerate(zip(text, prefix)):
        if expected != actual:
            return i
    return min(len(text), len(prefix))


class _FormatRestoreError(RuntimeError):
    """试改的规则恢复失败，整份工作簿必须关闭且不保存。"""


@print_layout.with_layout
def _try_extend_horizontal_format(sheet: Any, area: Any, items: list[dict]) -> str | None:
    """在当前页验证横排全文，逐列延伸原有条件格式或清除挡字的临时标记。"""
    rules = area.FormatConditions
    if int(rules.Count) != 1:
        return None
    rule = rules.Item(1)
    if int(rule.Type) not in (XL_CELL_VALUE, XL_EXPRESSION):
        return None
    tag = _rule_tag(rule)
    original = rule.AppliesTo
    row, col = int(area.Row), int(area.Column)
    if (int(original.Areas.Count) != 1 or int(original.Rows.Count) != 1
            or int(original.Row) != row or int(original.Column) != col):
        return None

    text = normalize_text(area.Cells(1, 1).Value)
    evidence = next((it for it in items if it.get("kind") == "hidden" and it.get("pdf_path")
                     and it.get("bbox") and isinstance(it.get("page_index"), int)
                     and normalize_text(it["text"]) == text), None)
    if evidence is None:
        return None
    index, anchor = evidence["page_index"], evidence["bbox"]
    with fitz.open(evidence["pdf_path"]) as doc:
        page_address = page_breaks.page_area(sheet, index, len(doc))
        if not page_address:
            return None
        visible = _visible_horizontal_prefix([doc[index]], text, anchor)
    if not 0 < visible < len(text):
        return None
    top, left = _cell_rc(page_address.split(":")[0])
    bottom, right = _cell_rc(page_address.split(":")[-1])
    current = col + int(original.Columns.Count) - 1
    if not (top <= row <= bottom and left <= col <= current < right):
        return None

    pdf = Path(tempfile.gettempdir()) / f"~横排复检_{uuid4().hex}.pdf"
    attempted = False
    kept = False
    try:
        # 工具自己的标记最终要清除。先挪到空格，验证标题在清除标记后能否显示完整。
        last = current + 1 if tag else min(right, col + 32)
        for end in range(current + 1, last + 1):
            neighbor = sheet.Cells(row, end)
            if (neighbor.Value2 is not None or neighbor.HasFormula or neighbor.MergeCells
                    or float(neighbor.Width) <= 0 or int(neighbor.FormatConditions.Count)):
                break
            address = f"{_col_letters(col)}{row}:{_col_letters(end)}{row}"
            attempted = True
            rule.ModifyAppliesToRange(neighbor if tag else sheet.Range(address))
            pdf.unlink(missing_ok=True)
            wf._export_sheet_pdf(sheet, pdf, (index + 1, index + 1))
            with fitz.open(pdf) as page:
                shown = _visible_horizontal_prefix(page, text, anchor) if len(page) == 1 else 0
            if shown == len(text):
                if tag:
                    rule.Delete()
                kept = True
                debug_log.event("↔️ 横排全文可见", 表=sheet.Name, 应用范围="已清除临时标记" if tag else address)
                return HORIZONTAL_MARK_PASSED if tag else HORIZONTAL_FORMAT_PASSED
            if shown <= visible:
                break
            visible = shown
    finally:
        try:
            if attempted and not kept:
                rule.ModifyAppliesToRange(original)
        except Exception as err:
            raise _FormatRestoreError("未能恢复条件格式的应用范围") from err
        finally:
            try:
                pdf.unlink(missing_ok=True)
            except OSError as err:
                debug_log.exception(err, "清理横排试改 PDF")
    return None


def _wrapped_lines(text: str, size: float, usable: float) -> int:
    """像 Excel 一样按单词换行数行数：英数单词放不下整个挪到下一行，比一行还长的单词才拆开；全角字可以随处断"""
    lines, used = 1, 0.0
    for token in re.findall(r"[^\s\u3000-\u9fff\uff00-\uffef]+\s*|\s+|.", text):
        width, full = _text_width_pt(token.rstrip(), size), _text_width_pt(token, size)
        if used and used + width > usable:
            lines, used = lines + 1, 0.0
        while width > usable:  # 一个单词比一整行还长：拆开
            lines, width, full = lines + 1, width - usable, full - usable
        used += full
    return lines


def _wrap_shortfall(area: Any, size: float) -> tuple[float, int]:
    """开换行后估算要几行，返回（现在的高度还差多少 pt，估算行数）。读不到文字或宽度时返回 (0, 0)。
    合并单元格按整个区域的宽、高算。"""
    try:
        value = area.Cells(1, 1).Value
        width, height = area.Width, area.Height
    except Exception:
        return 0.0, 0
    if isinstance(value, bool) or not isinstance(value, (str, int, float)):
        return 0.0, 0
    if isinstance(width, bool) or not isinstance(width, (int, float)) \
            or isinstance(height, bool) or not isinstance(height, (int, float)):
        return 0.0, 0
    usable = max(float(width) - CELL_PADDING_PT, size)
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    lines = sum(_wrapped_lines(part, size, usable) for part in text.split("\n"))
    need = lines * size * LINE_HEIGHT_RATIO
    return max(0.0, need - float(height)), lines


def _measure_wrapped_height(tmp: Any, area: Any, wrap: bool = True) -> float:
    """让 Excel 自己量这格开换行后要多高（pt）：Excel 的「自动调整行高」会跳过合并单元格，
    所以把文字和字体搬到临时工作簿的一个不合并的格子上，宽度调成和整个合并区域一样，开换行后自动调整行高，
    读出高度。Excel 按单词换行（SQL 里长单词放不下会整个挪到下一行），比按字数估算准。量不出返回 0。
    wrap=False：不换行的格子（收紧行高时量同一行别的格子），只量一行字要多高，不用对宽度。"""
    first = area.Cells(1, 1)
    value = first.Value
    text = value if isinstance(value, str) else str(first.Text or "")
    if not text:
        return 0.0
    cell = tmp.Cells(1, 1)
    cell.NumberFormat = "@"  # SQL 等以 = 开头的文字也按文字放，不当公式
    cell.Value = text
    font, src = cell.Font, first.Font
    font.Name, font.Size, font.Bold, font.Italic = src.Name, src.Size, src.Bold, src.Italic
    cell.WrapText = wrap
    if not wrap:
        tmp.Rows(1).AutoFit()
        height = float(tmp.Rows(1).RowHeight)
        cell.Clear()
        return height
    width = float(area.Width)
    column = tmp.Columns(1)
    for _ in range(4):  # 列宽单位是字数、和 pt 不完全成比例，调几次贴近
        now = float(column.Width)
        if now <= 0:
            break
        column.ColumnWidth = min(255.0, max(0.5, float(column.ColumnWidth) * width / now))
    if abs(float(column.Width) - width) > 3:
        raise ValueError(f"临时格子宽度对不上（要 {width}pt，只能 {column.Width}pt）")
    tmp.Rows(1).AutoFit()
    height = float(tmp.Rows(1).RowHeight)
    cell.Clear()
    return height


class _Measurer:
    """一份文件开换行加高时共用一个临时工作簿量高度（不保存、用完关掉，不碰原文件）"""

    def __init__(self, excel: Any) -> None:
        self.excel, self.book, self.failed = excel, None, False

    def height(self, area: Any, wrap: bool = True) -> float:
        if self.failed:
            return 0.0
        try:
            if self.book is None:
                self.book = self.excel.Workbooks.Add()
            return _measure_wrapped_height(self.book.Worksheets(1), area, wrap)
        except Exception as err:
            debug_log.event("📏 Excel 量不出行高", 原因=str(err)[:120])
            if self.book is None:
                self.failed = True  # 临时工作簿都开不了：这份文件后面都用估算
            return 0.0

    def close(self) -> None:
        if self.book is not None:
            try:
                self.book.Close(SaveChanges=False)
            except Exception as err:
                debug_log.event("⚠️ 临时工作簿关不掉", 原因=str(err)[:120])
            self.book = None


_ONE_WORD = re.compile(r"[!-~]+")  # 只有半角字母数字符号、中间没有空格


def _width_only(card: dict, items: list[dict]) -> bool:
    """调用方持有 STATE_LOCK：只是一个英文单词放不下（例如 shikuchoson 只显示出 shikuchoso）。
    PDF 里的字比 Excel 屏幕上宽一点，Excel 量着放得下，打印时最后一两个字母被挤到下一行：至少多给一行行高。
    （2026-11-04 试过加列宽：表是「缩放到一页宽」时整页跟着缩小，下划线被切掉，一份多出 44 处问题，已去掉）"""
    text = str(card.get("text") or "").strip()
    if not _ONE_WORD.fullmatch(text) or not items:
        return False
    for it in items:
        if autofix.item_error_type(it) != autofix.ERROR_TYPE_TEXT_TRUNCATED:
            return False
        shown = str(it.get("visible_text") or "").strip()
        if not shown or shown == text or not text.startswith(shown):
            return False
    return True


def _grow_and_wrap(excel: Any, path: str, ratio: float, skip: set[str] | None = None) -> dict[str, dict]:
    """切字的格子开自动换行，估算换行后要几行，行高不够先补够，再加「字号 × 倍数」，保存一次。
    返回 {卡片 id: 做了什么}。同一行有几个格子只加一次，按其中要加最多的算；最高到 Excel 上限 409pt。
    skip：这些卡不再加（复检时已经加高过、还没过的，交给人工）。"""
    with wf.STATE_LOCK:
        cards = [c for c in _grow_targets(path) if c["id"] not in (skip or set())]
        items_by_card = {c["id"]: _items_at(c) for c in cards}
        # 一个英文单词放不下：最后几个字母被挤到下一行，至少多给一行
        extra_line = {c["id"] for c in cards if _width_only(c, items_by_card[c["id"]])}
    if not cards:
        return {}
    _check_file(path)
    before = wf._file_version(Path(path))
    if wf.STATE["checked_files"].get(path) != before:
        raise ValueError(f"扫描后文件已修改，请重新扫描：{Path(path).name}")
    name = Path(path).name
    debug_log.event("📐 开换行加高", 文件=name, 格数=len(cards), 倍数=ratio)
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0)
    print_layout.bind(book, path)
    done: dict[str, dict] = {}
    measurer = _Measurer(excel)
    try:
        if book.ReadOnly:
            raise ValueError(f"{name} 以只读方式打开，没能开换行、加高")
        shapes_before = shape_geometry.snapshot_book(book)  # 加高前记下图形，保存前放回（用户 2026-11-11 要）
        rows: dict[tuple[str, int], dict] = {}
        for card in cards:
            where = f"{card['sheet_name']}!{card['cell_addr']}"
            try:
                sheet = _find_sheet(book, card["sheet_name"])
                if sheet is None:
                    raise ValueError("找不到这张表")
                area = sheet.Range(card["cell_addr"]).MergeArea
                size = _font_size(area)
                shrink = bool(area.ShrinkToFit)
                if not shrink and _has_horizontal_space(sheet, area, size):
                    result = done[card["id"]] = {"horizontal_layout": True}
                    try:
                        fixed = _try_extend_horizontal_format(sheet, area, items_by_card[card["id"]])
                    except _FormatRestoreError:
                        raise
                    except Exception as err:
                        debug_log.exception(err, f"试延伸条件格式 {where}")
                    else:
                        if fixed:
                            result["horizontal_fix"] = fixed
                    debug_log.event("↔️ 保留横排", 格=where, 结果=result.get("horizontal_fix", "保留原范围"))
                    continue
                if shrink:
                    wrap = "缩小字体填充，不开换行"
                elif area.WrapText is True:
                    wrap = "本来就开着"
                else:
                    area.WrapText = True
                    wrap = "已开"
                # 文字很长（例如一整句 SQL）时，只加一行不够：先让 Excel 在临时格子上量出换行后要多高
                # （合并单元格按整个区域的宽度），量不出才按字数估算；不够的先补够，再加「字号 × 倍数」
                shortfall, lines, by = 0.0, 0, "-"
                if not shrink:
                    measured = measurer.height(area)
                    if measured > 0:
                        shortfall, by = max(0.0, measured - float(area.Height)), "Excel"
                    else:
                        shortfall, lines = _wrap_shortfall(area, size)
                        by = "估算"
                    if card["id"] in extra_line:
                        # Excel 量着放得下、打印时最后几个字母被挤到下一行：至少多给一行
                        shortfall, by = max(shortfall, size * LINE_HEIGHT_RATIO), by + "+一行"
                row = int(area.Row) + int(area.Rows.Count) - 1  # 合并单元格加在最下面一行
                entry = rows.setdefault((sheet.Name, row), {"sheet": sheet, "size": 0.0, "extra": 0.0, "cards": []})
                entry["size"] = max(entry["size"], size)
                entry["extra"] = max(entry["extra"], shortfall + size * ratio)
                entry["cards"].append(card["id"])
                done[card["id"]] = {"wrap": wrap == "已开", "row": row, "font": size,
                                    "merged": int(area.Rows.Count) > 1 or int(area.Columns.Count) > 1}
                if card["id"] in extra_line:
                    done[card["id"]]["extra_line"] = True
                if shortfall:
                    done[card["id"]].update(shortfall=round(shortfall, 2), measured_by=by)
                    if lines:
                        done[card["id"]]["lines"] = lines
                debug_log.event("↪️ 自动换行", 格=where, 结果=wrap, 字号=size, 量高=by, 估算行数=lines or "-",
                                还差=round(shortfall, 2))
            except _FormatRestoreError:
                raise
            except Exception as err:
                done[card["id"]] = {"error": str(err)}
                wf.log_msg(f"⚠️ {name} / {where} 没能开换行、加高：{err}")
                debug_log.event("⚠️ 开换行加高失败", 格=where, 原因=str(err))
        for (sheet_name, row), entry in rows.items():
            target = entry["sheet"].Rows(row)
            base = float(target.RowHeight)
            applied = round(min(autofix.MAX_ROW_HEIGHT, base + entry["extra"]), 2)
            target.RowHeight = applied
            debug_log.event("📏 加高", 表=sheet_name, 行=row, 原行高=base, 新行高=applied,
                            字号=entry["size"], 倍数=ratio, 估算补高=round(entry["extra"] - entry["size"] * ratio, 2))
            for cid in entry["cards"]:
                done[cid].update(sheet=sheet_name, base=round(base, 2), applied=applied, ratio=ratio)
        horizontal_count = sum(bool(d.get("horizontal_fix")) for d in done.values())
        if rows or horizontal_count:
            shape_geometry.restore_book(book, shapes_before, f"开换行加高 {name}")
            normalize_workbook_view(excel, book)
            book.Save()
            _followed_own_save(path, before)
            changed = sum("applied" in d for d in done.values())
            if rows:
                wf.log_msg(f"📐 {name}：{changed} 格开自动换行、{len(rows)} 行加高（字号 × {ratio}），已保存")
            if horizontal_count:
                wf.log_msg(f"↔️ {name}：{horizontal_count} 处横排文字调整条件格式后显示完整，已保存")
            debug_log.event("💾 已保存", 文件=name, 行数=len(rows), 条件格式=horizontal_count)
    finally:
        measurer.close()
        book.Close(SaveChanges=False)
    return done


def _recheck_after_grow(session: Any, path: str) -> None:
    """开完换行、加完高保存后整份复检一次（用户 2026-11-03 定）：通过的格子翻篇、不标色；
    没通过的不再改，留黄色交给人工。复检出错不影响扫描结果，所有问题照旧标色转人工。"""
    with wf.STATE_LOCK:
        cards = [c for c in _data()["cards"].values() if c["wb_path"] == path
                 and c["color"] == "yellow" and c.get("finalized") and not c.get("dismissed")]
    if not cards:
        return
    debug_log.event("🔁 加高后复检", 文件=Path(path).name, 卡数=len(cards))
    session.touch()
    try:
        _recheck_file(session.app(), path, cards, passed_message=GROWN_PASSED_MESSAGE,
                      found_message=GROWN_FAILED_MESSAGE)
    except Exception as err:
        wf.log_msg(f"⚠️ {Path(path).name} 加高后复检没做完，问题照旧标色转人工，可稍后点复检：{err}{_where(err)}")
        debug_log.event("⚠️ 加高后复检失败", 文件=Path(path).name, 原因=str(err))
        _restart_excel(session)


RECHECK_GROWN_PASSED_MESSAGE = "复检查出切字，已开自动换行、加高，再复检通过"


def _grown_before_recheck(path: str) -> set[str]:
    """复检前还在列表里、已经开过换行加高的卡：复检又查出来就不再加，交给人工（不会每复检一次就加高一次）"""
    with wf.STATE_LOCK:
        return {c["id"] for c in _data()["cards"].values() if c["wb_path"] == path and c.get("grown")
                and c["color"] == "yellow" and c.get("finalized") and not c.get("dismissed")}


def _grow_after_recheck(session: Any, path: str, tried: set[str], ratio: float | None = None) -> None:
    """整份复检又查出新的切字（例如人工改完后行高不够）：和扫描一样开换行、加高、保存，再整份复检一次
    （用户 2026-11-04 要）。只加不减，复检前已经加高过的格子不再加；加完还没过的留黄色交给人工。
    出错只记日志，复检结果照旧。"""
    name = Path(path).name
    if ratio is None:
        with wf.STATE_LOCK:
            ratio = grow_ratio(_data().get("grow_ratio"))
    session.touch()
    try:
        grown = _grow_and_wrap(session.app(), path, ratio, skip=tried)
    except Exception as err:
        wf.log_msg(f"⚠️ {name} 复检后没能开换行、加高，请人工处理：{err}{_where(err)}")
        debug_log.event("⚠️ 复检后加高失败", 文件=name, 原因=str(err))
        if not isinstance(err, ValueError):
            _restart_excel(session)
        return
    if not grown:
        return
    with wf.STATE_LOCK:
        cards = _data()["cards"]
        for cid, d in grown.items():
            card = cards.get(cid)
            if card is None:
                continue
            if d.get("horizontal_layout"):
                card.update(horizontal_layout=True, message=HORIZONTAL_LAYOUT_MESSAGE)
                if d.get("horizontal_fix"):
                    card["horizontal_fix"] = d["horizontal_fix"]
                for item in _items_at(card):
                    item["review_reason"] = HORIZONTAL_LAYOUT_MESSAGE
            else:
                was_horizontal = card.pop("horizontal_layout", None)
                if "applied" in d:
                    card.update(grown=d, message=GROWN_MESSAGE)
                elif was_horizontal:
                    card["message"] = "复检查出的问题：请修改并保存后再复检"
        wf._persist_state()
    changed = sum("applied" in d or bool(d.get("horizontal_fix")) for d in grown.values())
    if not changed:
        return
    wf.log_msg(f"📐 {name}：已处理 {changed} 处切字，再复检一遍")
    debug_log.event("🔁 复检加高后再复检", 文件=name, 格数=changed)
    with wf.STATE_LOCK:
        cards = [c for c in _data()["cards"].values() if c["wb_path"] == path
                 and c["color"] == "yellow" and c.get("finalized") and not c.get("dismissed")]
    session.touch()
    try:
        _recheck_file(session.app(), path, cards, passed_message=RECHECK_GROWN_PASSED_MESSAGE,
                      found_message=GROWN_FAILED_MESSAGE)
    except Exception as err:
        wf.log_msg(f"⚠️ {name} 加高后复检没做完，可稍后再点复检：{err}{_where(err)}")
        debug_log.event("⚠️ 复检加高后再复检失败", 文件=name, 原因=str(err))
        _restart_excel(session)


# 收紧行高（用户 2026-11-03 要）：人工修切字时先给问题行一个很大的行高（例如 200），再让程序把多余的收回来。
# 不一点点试：每试一次都要导出 PDF 检测，太慢。改成一步量出要多高：这一行每个有字的格子都搬到临时工作簿里
# 同宽、同字体的格子上让 Excel 自动调整行高（量不出才按单词换行估算），取要的最高的，再加「字号 × 倍数」余量。
# 只收不加；收完保存、整份复检，复检没过的行退回原来的行高，保存后再复检一次。
TIGHTEN_SLACK_PT = 0.75  # 只差一个像素以内的不收
TIGHTEN_PASSED_MESSAGE = "收紧行高后复检通过"
XL_HORIZONTAL = -4128  # Orientation：横排


def _addr_rows(addr: str) -> tuple[int, int]:
    """单元格地址占哪几行（合并格 G5:I7 → 5–7）"""
    parts = [_cell_rc(p)[0] for p in addr.split(":") if p]
    return min(parts), max(parts)


def _tighten_targets(path: str) -> list[dict]:
    """调用方持有 STATE_LOCK：这份文件里出过切字问题的格子（还在列表里的、复检通过翻篇的都算），按它们所在的行收紧"""
    return [c for c in _data()["cards"].values() if c["wb_path"] == path and c.get("cell_addr")
            and not _keeps_user_mark(c) and set(_card_error_types(c)) & set(GROW_ERROR_TYPES)]


def _card_rows(cards: list[dict]) -> set[tuple[str, int]]:
    """这些卡的格子压着哪些行（表名不分大小写）"""
    rows: set[tuple[str, int]] = set()
    for card in cards:
        try:
            low, high = _addr_rows(card.get("cell_addr") or "")
        except ValueError:
            low = high = int(card.get("excel_row") or 0)
        rows.update((_sheet_key(card["sheet_name"]), r) for r in range(low, high + 1))
    return rows


def _shape_rows(sheet: Any) -> set[int]:
    """图片、图形压着的行不收：收了图会被压扁或挡住字，复检也查不出来（批注不算）"""
    rows: set[int] = set()
    try:
        shapes = list(sheet.Shapes)
    except Exception:
        return rows
    for shape in shapes:
        try:
            if int(shape.Type) == 4:  # msoComment
                continue
            rows.update(range(int(shape.TopLeftCell.Row), int(shape.BottomRightCell.Row) + 1))
        except Exception:
            continue
    return rows


def _measure_area(measurer: "_Measurer", area: Any, size: float) -> tuple[float, str]:
    """这一格的字全部显示要多高（pt）：先让 Excel 在临时格子上量，量不出按单词换行估算"""
    wrap = area.WrapText is True and not bool(area.ShrinkToFit)
    measured = measurer.height(area, wrap=wrap)
    if measured > 0:
        return measured, "Excel"
    if not wrap:
        return size * LINE_HEIGHT_RATIO, "估算"
    _, lines = _wrap_shortfall(area, size)
    if not lines:
        return size * LINE_HEIGHT_RATIO, "估算"
    return lines * size * LINE_HEIGHT_RATIO, "估算"


def _row_target(measurer: "_Measurer", sheet: Any, row: int, height: float, ratio: float) -> tuple[float | None, str]:
    """这一行收到多高：每个有字的格子（合并格减去它在别的行占的高度）要的高度加「字号 × 倍数」，取最高的。
    有竖排、斜排文字的行量不准，返回 (None, 原因)"""
    used = sheet.UsedRange
    first = int(used.Column)
    last = first + int(used.Columns.Count) - 1
    target, seen, why = 0.0, set(), "-"
    for col in range(first, last + 1):
        area = sheet.Cells(row, col).MergeArea
        top, left = int(area.Row), int(area.Column)
        if (top, left) in seen:
            continue
        seen.add((top, left))
        head = area.Cells(1, 1)
        value = head.Value
        if value is None or (isinstance(value, str) and not value.strip()):
            continue
        try:
            orientation = head.Orientation
        except Exception:
            orientation = XL_HORIZONTAL
        if isinstance(orientation, (int, float)) and not isinstance(orientation, bool) and int(orientation) not in (0, XL_HORIZONTAL):
            return None, "有竖排或斜排的字，量不准"
        size = _font_size(area)
        need, by = _measure_area(measurer, area, size)
        need -= float(area.Height) - height  # 合并格在别的行已经占了的高度
        if need + size * ratio > target:
            target, why = need + size * ratio, by
    if target <= 0:
        return None, "这一行没有字"
    return round(min(autofix.MAX_ROW_HEIGHT, target), 2), why


def _tighten_rows(excel: Any, path: str, ratio: float) -> dict[tuple[str, int], dict]:
    """把这份文件里出过切字问题的行收到「量出的高度 + 字号 × 倍数」，只收不加，保存一次。
    返回 {(表名, 行): {"before": 原行高, "after": 新行高, "by": 怎么量的}}，只含真的收了的行"""
    with wf.STATE_LOCK:
        cards = _tighten_targets(path)
    if not cards:
        return {}
    _check_file(path)
    before = wf._file_version(Path(path))
    name = Path(path).name
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0)
    measurer = _Measurer(excel)
    done: dict[tuple[str, int], dict] = {}
    try:
        if book.ReadOnly:
            raise ValueError(f"{name} 以只读方式打开，没能收紧行高")
        shapes_before = shape_geometry.snapshot_book(book)  # 收行高前记下图形，保存前放回（用户 2026-11-11 要）
        rows: dict[str, tuple[Any, set[int]]] = {}
        for card in cards:
            sheet = _find_sheet(book, card["sheet_name"])
            if sheet is None:
                debug_log.event("⚠️ 收紧行高：找不到这张表", 文件=name, 表=card["sheet_name"])
                continue
            try:
                area = sheet.Range(card["cell_addr"]).MergeArea
                top = int(area.Row)
                span = range(top, top + int(area.Rows.Count))
            except Exception as err:
                debug_log.event("⚠️ 收紧行高：找不到这一格", 格=f"{card['sheet_name']}!{card['cell_addr']}", 原因=str(err)[:120])
                continue
            rows.setdefault(sheet.Name, (sheet, set()))[1].update(span)
        for sheet_name, (sheet, wanted) in rows.items():
            busy_rows = _shape_rows(sheet)
            for row in sorted(wanted):
                where = f"{sheet_name}!{row} 行"
                try:
                    target_row = sheet.Rows(row)
                    if target_row.Hidden is True:
                        continue
                    height = float(target_row.RowHeight)
                    if row in busy_rows:
                        debug_log.event("↕️ 不收", 行=where, 原因="有图片或图形", 行高=height)
                        continue
                    target, by = _row_target(measurer, sheet, row, height, ratio)
                    if target is None or target >= height - TIGHTEN_SLACK_PT:
                        debug_log.event("↕️ 不收", 行=where, 行高=height, 要=target, 原因=by if target is None else "已经不高")
                        continue
                    target_row.RowHeight = target
                    done[(sheet_name, row)] = {"before": round(height, 2), "after": target, "by": by}
                    debug_log.event("↕️ 收紧", 行=where, 原行高=height, 新行高=target, 量高=by, 倍数=ratio)
                except Exception as err:
                    wf.log_msg(f"⚠️ {name} / {where} 没能收紧：{err}")
                    debug_log.event("⚠️ 收紧行高失败", 行=where, 原因=str(err)[:200])
        if done:
            shape_geometry.restore_book(book, shapes_before, f"收紧行高 {name}")
            normalize_workbook_view(excel, book)
            book.Save()
            _followed_own_save(path, before)
            debug_log.event("💾 已保存", 文件=name, 收紧行数=len(done))
    finally:
        measurer.close()
        book.Close(SaveChanges=False)
    saved = sum(v["before"] - v["after"] for v in done.values())
    wf.log_msg(f"↕️ {name}：{len(done)} 行收紧行高（共收回 {saved:.0f}pt）" if done else f"ℹ️ {name}：没有可以收紧的行")
    return done


def _failed_rows(path: str, tightened: dict[tuple[str, int], dict]) -> list[tuple[str, int]]:
    """复检后还有问题的格子压着的、这次收紧过的行"""
    with wf.STATE_LOCK:
        unresolved = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                      and it["status"] not in wf.RESOLVED_STATUSES]
    failed = []
    for it in unresolved:
        try:
            low, high = _addr_rows(it.get("cell_addr") or "")
        except ValueError:
            low = high = int(it.get("excel_row") or 0)
        failed += [key for key in tightened if _sheet_key(key[0]) == _sheet_key(it["sheet_name"])
                   and low <= key[1] <= high and key not in failed]
    return failed


def _restore_rows(excel: Any, path: str, rows: dict[tuple[str, int], float]) -> None:
    """复检没过的行退回原来的行高，保存"""
    _check_file(path)
    before = wf._file_version(Path(path))
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0)
    try:
        shapes_before = shape_geometry.snapshot_book(book)  # 退回行高前记下图形，保存前放回（用户 2026-11-11 要）
        for (sheet_name, row), height in rows.items():
            sheet = _find_sheet(book, sheet_name)
            if sheet is not None:
                sheet.Rows(row).RowHeight = height
                debug_log.event("↩️ 退回原行高", 行=f"{sheet_name}!{row} 行", 行高=height)
        shape_geometry.restore_book(book, shapes_before, f"退回原行高 {Path(path).name}")
        normalize_workbook_view(excel, book)
        book.Save()
        _followed_own_save(path, before)
    finally:
        book.Close(SaveChanges=False)


def _tighten_file(session: Any, path: str, ratio: float) -> dict:
    """一份文件：收紧 → 整份复检 → 没过的行退回原行高 → 再复检一次（只有退回过才要）。返回这份的结果"""
    name = Path(path).name
    unsaved = close_saved_workbooks([Path(path)])
    if Path(path) in unsaved:
        raise ValueError(f"{unsaved[Path(path)]}，请在 Excel 里保存并关闭后重试")
    with wf.STATE_LOCK:
        handled = _card_rows(_tighten_targets(path))  # 以前出过切字的行：收紧管的，之后不自动加高
    session.touch()
    tightened = _tighten_rows(session.app(), path, ratio)

    def recheck(passed: str) -> None:
        with wf.STATE_LOCK:
            cards = [c for c in _data()["cards"].values() if c["wb_path"] == path
                     and c["color"] == "yellow" and c.get("finalized") and not c.get("dismissed")]
        session.touch()
        _recheck_file(session.app(), path, cards, passed_message=passed)

    recheck(TIGHTEN_PASSED_MESSAGE if tightened else "人工修改后复检通过")
    failed = _failed_rows(path, tightened)
    if failed:
        wf.log_msg(f"↩️ {name}：{len(failed)} 行收紧后复检没过，退回原来的行高再复检")
        _restore_rows(session.app(), path, {key: tightened[key]["before"] for key in failed})
        recheck("人工修改后复检通过")
    # 复检在别的行查出新的切字（以前没出过问题、没收过也没退回过）：和普通复检一样开换行、加高再复检（用户 2026-11-04 要）
    with wf.STATE_LOCK:
        skip = {c["id"] for c in _data()["cards"].values() if c["wb_path"] == path and _card_rows([c]) & handled}
    _grow_after_recheck(session, path, skip | _grown_before_recheck(path), ratio)
    result = {"tightened": len(tightened) - len(failed), "restored": len(failed),
              "saved_pt": round(sum(v["before"] - v["after"] for k, v in tightened.items() if k not in failed), 1)}
    debug_log.event("↕️ 收紧行高完成", 文件=name, **{"收紧行数": result["tightened"], "退回行数": result["restored"],
                                                 "收回pt": result["saved_pt"]})
    return result


def _hand_over(path: str, grown: dict[str, dict] | None = None) -> None:
    """本文件检出的问题全部转人工（黄色、已提交）。grown：这次开过换行、加高的格子，卡片上写明做了什么。
    你已经手动标成蓝 / 绿、还没提交的卡保留你的选择。"""
    grown = grown or {}
    with wf.STATE_LOCK:
        pending = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                   and it["status"] not in wf.RESOLVED_STATUSES]
        for group in _group_cells(pending):
            card = _card_for(group)
            if _keeps_user_mark(card):
                card["waiting"] = False
                continue
            done = grown.get(card["id"], {})
            if done.get("horizontal_layout"):
                card["horizontal_layout"] = True
                message = HORIZONTAL_LAYOUT_MESSAGE
            elif any(autofix.column_width_only(it) for it in group):
                message = HASH_MESSAGE
            elif not _can_highlight(card):
                message = "单元格定位未确认，未在 Excel 中标色，请按截图查找后人工处理"
            elif done and "error" in done:
                message = f"没能开换行、加高（{done['error']}），请人工处理并保存后复检"
            elif done:
                message = GROWN_MESSAGE
            else:
                message = SCAN_ONLY_MESSAGE
            card.update(color="yellow", verified=False, manual=True, finalized=True, waiting=False,
                        dismissed=False, after_kind="scan_only", message=message)
            card.pop("handed_over", None)
            card.pop("width_only", None)  # 2026-11-04 试过加列宽的旧记录
            if not done.get("horizontal_layout"):
                card.pop("horizontal_layout", None)
            if done.get("horizontal_fix"):
                card["horizontal_fix"] = done["horizontal_fix"]
            else:
                card.pop("horizontal_fix", None)
            if "applied" in done:
                card["grown"] = done
            else:
                card.pop("grown", None)
            for item in _items_at(card):
                item.update(status="manual_review", batch_managed=True, triage_keep=False, review_reason=message)
        wf._persist_state()
    wf._sync_problem_counts(path, "扫描完成，已标记，转人工处理")


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


def _scan_model(job: dict, key: str = "scan_seconds") -> tuple[float, float] | None:
    """按这次已扫完文件的耗时估算扫描速度：(平均每份秒数, 每字节秒数)；一份都没扫完时返回 None。
    key="scan_seconds" 只算检测（定超时上限用）；"file_seconds" 是整份做完（含加高、复检、标色，估剩余时间用）"""
    done = job.get(key) or {}
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


def _skip_slow_file(job: dict, path: str, limit: float, manual: bool = False) -> None:
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
    message = (("你点了「跳过这一份」" if manual else f"扫描超过 {_minutes(limit)}")
               + (f"，卡在表「{sheet}」" if sheet else "") + "，已跳过，请人工检查")
    scan_analysis.timeout(job, path, sheet, limit)
    wf.log_msg(f"⏱ {Path(path).name}：{message}")
    debug_log.event("⏭ 扫描手动跳过" if manual else "⏱ 扫描超时跳过", 文件=Path(path).name, 表=sheet, 上限秒=round(limit))
    update_file_progress(workbook_rel_name(Path(path), wf.STATE["folder"]), status="error", error=message,
                         current_action="手动跳过" if manual else "扫描超时，已跳过")


def clear_slow_files() -> None:
    """页面上「知道了」：不再提示当前文件夹扫描超时跳过的文件"""
    with wf.STATE_LOCK:
        slow = _data().get("slow_files", {})
        for path in [p for p in slow if wf.in_current_folder(p)]:
            del slow[path]
        wf._persist_state()


def _scan_eta(job: dict) -> dict | None:
    """调用方持有 STATE_LOCK：扫描时的已用 / 预计剩余时间（秒）和当前文件的进度。剩余按这次已扫完文件的耗时、
    文件大小估算，一份都没扫完时不估"""
    if not wf.STATE["busy"] or job.get("kind") != "batch_scan" or "workbooks" not in job:
        return None
    finished = set(job.get("scanned", [])) | set(job.get("completed", []))
    pending = [p for p in job["workbooks"] if p not in finished]
    now = wf.SCAN_NOW if wf.SCAN_NOW.get("path") in pending else {}
    running = max(0.0, time.time() - now["started"]) if now else 0.0
    # 剩余时间按整份做完的耗时估（加高后还要复检一次，只按检测估会少一半）
    model = _scan_model(job, "file_seconds") or _scan_model(job)
    limit = job.get("scan_limit") or {}
    return {
        "elapsed": round(sum({**(job.get("scan_seconds") or {}), **(job.get("file_seconds") or {})}.values())
                         + sum((job.get("slow_seconds") or {}).values()) + running),
        "remaining": None if model is None
        else round(max(0.0, sum(_expected_scan_seconds(job, p, model) for p in pending) - running)),
        "current": {"name": Path(now["path"]).name, "sheet": now.get("sheet", ""),
                    "index": now.get("index", 0), "total": now.get("total", 0), "seconds": round(running),
                    "limit": limit.get("seconds") if limit.get("path") == now["path"] else None} if now else None,
    }


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


XL_CELL_VALUE, XL_EXPRESSION = 1, 2
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
                wf._persist_state(force=True)  # 先记下原色再改：中途崩溃时下次还能改回去
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
                     tick: Callable[[], None] | None = None, report: dict | None = None,
                     scan_view: Any = None) -> bool:
    """让原 Excel 里的标记和卡片一致：黄色人工待处理的格子加一条条件格式规则；
    复检通过（变蓝）、提交为蓝 / 绿、或要求清除时把规则删掉。打开一次、保存一次。
    旧版本改过单元格底色的记录照旧还原原来的颜色。
    tick：每处理一格告诉看门狗还在动，问题格很多的大文件要好几分钟，不会被当成 Excel 卡死。
    report：清除标记时传入，逐格记下去掉了没有、没去掉的原因（cells），以及清掉的残留规则条数，页面上显示结果。
    scan_view：扫描时传入后台会话（autofix._ExcelSession）：标色这次打开文件时顺便存分页预览，
    和标色一起只保存一次（以前另外再打开、保存一次，用户 2026-11-07 同意合并）。
    返回是否已经顺便存了视图；没东西要标、根本没打开文件时返回 False，调用方照旧单独存视图。"""
    with wf.STATE_LOCK:
        data, marks = _data(), _marks()
        style = data.get("highlight_files", {}).get(path)
        cards = [c for c in data["cards"].values() if c["wb_path"] == path]
        want = set() if clear else _wanted_marks(path)
        have = {cid for cid, m in marks.items() if m["path"] == path}
        # 已经标了、颜色和这格现在该用的不一样（问题类型变了、换了配色、旧版只有一种颜色）：去掉重标
        by_id = {c["id"]: c for c in cards}
        recolor = {cid for cid in want & have if _is_rule(marks[cid])
                   and {k: marks[cid].get(k) for k in ("fill", "font")} != _mark_color(style, by_id[cid])}
        # 没确认存进文件的记录（上次标色被打断，或旧版本没记这一项）：打开文件看一眼标记还在不在
        unsure = sorted(cid for cid in have if not marks[cid].get("saved"))
        add = [c for c in cards if c["id"] in (want - have) | recolor]
        remove = sorted((have - want) | recolor)
        tabs_stale = _tabs_stale(path, {marks[cid]["sheet"] for cid in have if cid not in remove}
                                 | {c["sheet_name"] for c in add})
    if not add and not remove and not unsure and not tabs_stale:
        return False
    _check_file(path)
    view_before = None
    if scan_view is not None:
        # 和单独存视图时一样：检测之后文件被别人改过就不存，要重新扫描
        view_before = wf._file_version(Path(path))
        with wf.STATE_LOCK:
            checked = wf.STATE["checked_files"].get(path)
        if checked != view_before:
            raise ValueError(f"扫描后文件已修改，请重新扫描：{Path(path).name}")
    debug_log.event("🎨 同步标色", 文件=Path(path).name, 清除=clear, 要标=len(add), 要去掉=len(remove),
                    待核对=len(unsure), 换色=len(recolor), 标签要核对=tabs_stale, 样式=style or "-")
    app = excel_for()
    pre_open = _version_or_none(path)  # 打开前的版本（旧版 .xls 可写打开就会改文件时间，按打开前的认）
    pdf_digest = _adjust_pdf_digest(path)  # 调整页面的 PDF 记录按哈希认，Excel 打开前先算好
    book = open_workbook_checked(app, str(Path(path).resolve()), UpdateLinks=0)
    added: dict[str, dict] = {}
    failed: dict[str, str] = {}
    deleted = restored = unmarked = 0
    leftovers: dict[str, int] = {}
    tabs = {"marked": 0, "restored": 0, "failed": 0}
    stuck: list[str] = []  # 去掉 / 还原时出错的记录，留着下次再试
    book_name, last_log = Path(path).name, time.monotonic()
    marks_changed = view_changed = False
    saved_from: Any = None  # 保存前的文件版本（真的保存了才有）

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
        if (view_before is not None and Path(path).suffix.lower() != ".xls"
                and wf._file_version(Path(path)) != view_before):
            # 旧版 .xls 可写打开时 Excel 会改文件时间，不算别人改过
            raise ValueError(f"扫描后文件已修改，请重新扫描：{Path(path).name}")
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
            elif cid not in recolor:
                unmarked += 1  # 规则被人手动删掉的也算：本来就没东西要清
                note(marks[cid], "ok", "已去掉标记" if tag in dropped else "标记规则已被删掉（或格子已删），不用清")
            debug_log.event("🧽 去掉标记", 位置=f"{marks[cid]['sheet']}!{marks[cid]['addr']}", 标识=tag,
                            条数=dropped.get(tag, 0), 出错=drop_failed.get(tag, "-"))
        if recolor & set(stuck):  # 旧颜色的规则删不掉：先不重标，免得一格两条规则
            add = [c for c in add if c["id"] not in stuck]
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
                                     "tag": _mark_name(card["id"]), "type": _mark_type(card),
                                     **_mark_color(style, card)}
            except Exception as err:
                failed[card["id"]] = why(card["sheet_name"], err)
                debug_log.event("❌ 标色失败", 位置=f"{card['sheet_name']}!{card.get('cell_addr')}", 原因=err)
                debug_log.exception(err, "标色")
        with wf.STATE_LOCK:
            # 先记下再加规则、再保存：中途出错时记录还在，之后按标识还能删掉
            for mark in added.values():
                mark["saved"] = False  # 保存成功后才算标上了
            marks.update(added)
            wf._persist_state(force=True)
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
        marks_changed = bool(added or remove or dropped or tabs["marked"] or tabs["restored"])
        if view_before is not None:
            try:
                app.ScreenUpdating = True  # 切视图要真的切到各张表上，先恢复刷新
            except Exception:
                pass
            view_changed = enable_scan_view(app, book)
        if marks_changed or view_changed:  # 只核对、没有要改的：不保存
            if tick is not None:
                tick()
            before = _version_or_none(path)
            book.Save()
            _followed_own_save(path, before)
            saved_from = pre_open
        for mark in added.values():
            mark["saved"] = True
    finally:
        try:
            app.ScreenUpdating = True
        except Exception:
            pass
        book.Close(SaveChanges=False)
    if saved_from is not None:
        _follow_marks_save(path, saved_from, pdf_digest)
    with wf.STATE_LOCK:
        for cid in remove:
            if cid not in stuck and cid not in added:  # 换色的格子去掉旧规则后又标上了新的
                marks.pop(cid, None)
        for cid, error in failed.items():
            card = data["cards"].get(cid)
            if card:
                card["message"] = f"未能在 Excel 中标色（{error}），请按截图查找后人工处理"
        if view_before is not None and not marks_changed:
            # 只改了视图：打印内容没变，和单独存视图一样，文件仍算「检测过的版本」，导出的 PDF 照样能用。
            # 标色也改了文件时不跟进（和以前标色另存一次一样，保守按改过处理）
            after = wf._file_version(Path(path))
            wf.STATE["checked_files"][path] = after
            checks = _sheet_checks().get(path, {})
            for name, checked in list(checks.items()):
                if checked == view_before:
                    checks[name] = after
            scan_view.view_saved(path, view_before, after)
        wf._persist_state()
    parts = [f"标色 {len(added)} 格"] if added else []
    recolored = len(recolor & set(added))
    if recolored:
        parts[0] += f"（其中 {recolored} 格换成这种问题的颜色）"
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
    if view_changed:
        parts.append("同时存了分页预览")
    if parts:
        wf.log_msg(f"🎨 {Path(path).name}：" + "，".join(parts))
    return view_before is not None


def _wanted_marks(path: str) -> set[str]:
    """这份文件里现在应该标色的卡片：黄色人工待处理的。调用方持有 STATE_LOCK"""
    data = _data()
    if not data.get("highlight_files", {}).get(path):
        return set()
    by_cell = _items_by_cell({path})
    return {c["id"] for c in data["cards"].values()
            if c["wb_path"] == path and not c.get("dismissed") and c["color"] == "yellow"
            and c.get("manual") and _can_highlight(c) and not _hidden_low_risk(c, by_cell.get(_cell(c), []))}


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


def clear_workbook_marks(excel: Any, book: Any) -> dict:
    """清理已打开工作簿中的工具标记，由调用方决定是否保存。"""
    try:
        try:
            excel.ScreenUpdating = False
        except Exception:
            pass
        _present, dropped, failed = _drop_rules(book, set())
        tabs = []
        for sheet in book.Worksheets:
            name = str(sheet.Name)
            try:
                if _tab_is_ours(sheet):
                    sheet.Tab.ColorIndex = XL_COLOR_INDEX_NONE
                    tabs.append({"sheet": name, "status": "ok", "reason": "已还原为无颜色"})
            except Exception as err:
                tabs.append({"sheet": name, "status": "failed", "reason": f"还原标签失败：{err}"})
        return {"leftovers": sum(dropped.values()), "leftover_failed": list(failed.values())[:20], "tabs": tabs}
    finally:
        try:
            excel.ScreenUpdating = True
        except Exception:
            pass


def clear_marks_in_file(excel: Any, path: str, report: dict | None = None) -> None:
    """不依赖 STATE 记录，全局清除一份文件里程序加的条件格式标记（CHE_MARK 规则），
    并还原被标成橙黄色的工作表标签；单元格直接设置的底色 / 字色不动。
    report：传入时填 leftovers（删掉的规则条数）、leftover_failed（删不掉的规则原因）、tabs（标签还原明细）。"""
    _check_file(path)
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError(f"{Path(path).name} 以只读方式打开，未能清除标记")
        cleared = clear_workbook_marks(excel, book)
        if report is not None:
            report.update(cleared)
        if cleared["leftovers"] or any(t["status"] == "ok" for t in cleared["tabs"]):
            book.Save()
    finally:
        book.Close(SaveChanges=False)


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


DEFAULT_OPEN_LIMIT = 10
_OPEN_LOCK = threading.Lock()
# 每次重置打开记录加一；打开途中轮次变了，这批结果属于旧一轮，不再记入新一轮。
_open_round = 0


@contextmanager
def _opening_files():
    if not _OPEN_LOCK.acquire(blocking=False):
        raise ValueError("正在打开文件，请等这批打开完成后再试")
    try:
        yield
    finally:
        _OPEN_LOCK.release()


def _open_folder_key() -> str:
    return str(Path(wf.STATE.get("folder") or ".").resolve()).casefold()


def _recheck_ready(card_paths: list[str]) -> list[str]:
    """打开过、在 Excel 里改过并保存、已经关掉的文件：网页上可以自动复检（用户 2026-11-10 要改起来顺手）。
    只看打开记录里的、还有卡片的文件；Excel 还开着（有锁文件）的不算，免得复检把你正在改的文件关掉。调用方持有 STATE_LOCK"""
    ready = []
    for path in sorted(_opened_paths() & set(card_paths)):
        checked = wf.STATE["checked_files"].get(path)
        try:
            if not checked or wf._file_version(Path(path)) == checked or workbook_locked(Path(path)):
                continue
        except OSError:
            continue
        ready.append(path)
    return ready


def _opened_paths() -> set[str]:
    """按当前文件夹保存本轮成功打开的路径；关闭 Excel 或重新启动服务都不清空。"""
    return set(_data().get("open_history", {}).get(_open_folder_key(), []))


def _remember_opened(paths: list[str], *, folder_key: str | None = None, round_: int | None = None) -> None:
    if not paths:
        return
    with wf.STATE_LOCK:
        if round_ is not None and round_ != _open_round:
            return
        history = _data().setdefault("open_history", {})
        key = folder_key or _open_folder_key()
        previous = history.get(key)
        history[key] = list(dict.fromkeys([*(previous or []), *paths]))
        try:
            wf._persist_state()
        except Exception:
            if previous is None:
                history.pop(key, None)
            else:
                history[key] = previous
            raise


def note_opened(path: Path | str) -> None:
    """网页上点文件名打开了一份还有问题的文件：也记成「已打开、待复检」，人工处理里各分类先不再列它"""
    want = str(Path(path).resolve()).casefold()
    with wf.STATE_LOCK:
        paths = {c["wb_path"] for c in _data()["cards"].values() if not c.get("dismissed")}
    key = next((p for p in paths if str(Path(p).resolve()).casefold() == want), None)
    if key:
        _remember_opened([key])


def _forget_opened(path: str) -> None:
    """复检完一份文件：不再算「待复检」，还有问题就回到各分类里"""
    with wf.STATE_LOCK:
        history = _data().get("open_history", {})
        key = _open_folder_key()
        if path in (history.get(key) or []):
            history[key] = [p for p in history[key] if p != path]
            wf._persist_state()


def reset_open_history() -> int:
    """「重置打开记录」：只重置当前文件夹的批量打开进度，不关闭 Excel，也不改变审核结果。"""
    global _open_round
    with _opening_files(), wf.STATE_LOCK:
        history = _data().get("open_history", {})
        key = _open_folder_key()
        previous = history.pop(key, None)
        try:
            wf._persist_state()
        except Exception:
            if previous is not None:
                history[key] = previous
            raise
        _open_round += 1
        return len(previous or [])


def start_new_open_round() -> None:
    """新扫描开始新一轮。调用方持有 STATE_LOCK 并负责持久化；不等正在进行的打开，
    也不会失败，避免扫描在清空旧结果后因打开操作抛错。"""
    global _open_round
    with wf.STATE_LOCK:
        _data().get("open_history", {}).pop(_open_folder_key(), None)
        _open_round += 1


def _open_candidates(ids: list[str]) -> list[dict]:
    """按传进来的卡片顺序列出文件（人工处理面板按左边列表的顺序传），同一文件的多张黄卡只算一次。调用方持有 STATE_LOCK。"""
    first = {}
    for card in (_data()["cards"][cid] for cid in dict.fromkeys(ids) if cid in _data()["cards"]):
        first.setdefault(card["wb_path"], card)
    return list(first.values())


def open_plan(ids: list[str], *, selected_files: bool = False) -> dict:
    """默认一次最多五份；明确勾选的文件按选择打开。后台仍要处理的文件暂不打开。"""
    with wf.STATE_LOCK:
        candidates = _open_candidates(ids)
        opened = _opened_paths()
        unopened = [c for c in candidates if c["wb_path"] not in opened]
        job = wf.STATE.get("job") or {}
        waiting, available = [], []
        for card in candidates if selected_files else unopened:
            reason = open_block_reason(job, Path(card["wb_path"])) if wf.STATE["busy"] else ""
            if reason:
                waiting.append({"name": card["wb_name"], "reason": reason})
            else:
                available.append(card)
        selected = available if selected_files else available[:DEFAULT_OPEN_LIMIT]
        return {"total": len(candidates), "available": len(available),
                "files": [{"name": c["wb_name"], "path": c["wb_path"]} for c in selected],
                "card_ids": [c["id"] for c in selected], "waiting": waiting,
                "already_opened": len(candidates) - len(unopened), "unopened": len(unopened),
                "round_opened": len(opened),
                "exhausted": not selected_files and bool(candidates) and not unopened,
                "remaining": len(candidates if selected_files else unopened) - len(selected)}


def open_ready(ids: list[str], *, selected_files: bool = False) -> dict:
    """后台在跑时批量打开：后台不再碰的文件马上在可见 Excel 打开、跳到目标单元格，
    还要等的文件列出原因，不排队（等它的话还不如做完一份开一份）"""
    with _opening_files():
        with wf.STATE_LOCK:
            plan = open_plan(ids, selected_files=selected_files)
            folder_key, round_ = _open_folder_key(), _open_round
            cards = [_data()["cards"][cid] for cid in plan["card_ids"]]
            targets = [(Path(c["wb_path"]), c["sheet_name"], c["cell_addr"], c["excel_row"]) for c in cards]
        if targets:
            try:
                autofix.wait_for_writes(15)  # 后台正在保存的文件别同时打开
            except ValueError:
                raise ValueError("后台正在保存 Excel，请几秒后再点") from None
        errors = open_workbooks_for_review(targets) if targets else {}
        _remember_opened([str(p) for p, *_ in targets if p not in errors], folder_key=folder_key, round_=round_)
    for path, error in errors.items():
        wf.log_msg(f"⚠️ {path.name}：{error}")
    return {"opened": [Path(p).name for p, *_ in targets if p not in errors],
            "waiting": plan["waiting"], "remaining": plan["remaining"] + len(errors),
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


def run() -> list[str]:
    """由 workflow 的单一任务槽执行；不依赖浏览器心跳。"""
    job = wf.STATE["job"]
    kind, options = job["kind"], job["options"]
    if kind == "batch_tighten":
        raise ValueError("自动调整行高已停用，请在 Excel 中人工修改并保存后复检")
    session = autofix._ExcelSession()
    # 每种任务都看着后台 Excel：它在看不见的地方卡住（弹了没人能点的框等）时结束它，不让任务永远卡着。
    # 清除颜色每改一格都报进度，卡 5 分钟就算卡住；检测类一张大表导出可能要很久，按默认时长
    watchdog = ExcelWatchdog(lambda: wf.log_msg(
        "⚠️ 后台 Excel 弹出了没人能点的提示框，已结束本任务专用 Excel" if watchdog.dialog else
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
            # 扫描只调整稀疏末页的打印缩放，再检测、标色；不自动修复单元格。
            # 旧版本留下的全自动 / 转回自动 / 自动调整任务续跑时也按扫描做
            if "workbooks" not in job:
                retry_path = options.get("retry_path")
                if retry_path:
                    folder, paths = Path(options["folder"]).resolve(), [Path(retry_path)]
                else:
                    folder, paths = wf.list_workbooks(options["folder"])
                with wf.STATE_LOCK:
                    if str(folder) != wf.STATE["folder"]:
                        # 审核卡片、忽略规则、行高记录按文件路径区分，各文件夹分开保留：切回来还能审核、还原
                        wf.STATE.update(items={}, pages={}, checked_files={})
                    wf.STATE.update(folder=str(folder), pdf_folder=str(folder))
                    if not retry_path:
                        start_new_open_round()  # 只有扫描整个文件夹才重置批量打开进度。
                    job["workbooks"] = [str(p) for p in paths]
                    wf._persist_state()
                init_execution_batch(task_name="扫描并标记", folder=folder, workbooks=paths,
                                     pdf_folder=None if retry_path else folder,
                                     preserve_existing_files=bool(retry_path))
            expected = job["workbooks"]
            with wf.STATE_LOCK:
                if job.get("pending_adjustment"):
                    pending = job["pending_adjustment"]
                    job.setdefault("deferred_adjustments", {})[pending["path"]] = pending
                    job["pending_adjustment"] = None
                scanned = job.setdefault("scanned", list(job["completed"]))
                versions = job.setdefault("scan_versions", {})
                if not job.get("analysis_id") or any(p not in scanned for p in expected):
                    wf.STATE["latest_scan_analysis"] = scan_analysis.begin(job, wf.STATE["folder"], expected, "扫描")
                job["closing"] = False
                wf._persist_state()

            def each_file(path: str, work: Callable[[str], None]) -> bool:
                """处理一份文件；程序出错（多半是 Excel 崩溃或卡住）重启 Excel 再试，还不行就跳过、可续跑"""
                tries = -1
                while True:
                    tries += 1
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
                        manual = watchdog.take_skip()
                        if manual and _stopping(job):  # 点的是「结束整个任务」：不算扫描超时，下次继续时重扫
                            err = RuntimeError(manual)
                            manual = ""
                        if watchdog.deadline_hit or manual:
                            # 扫描超过时间上限 / 用户点了「跳过这一份」：不重试，这份文件跳过、转人工，下一份用新的 Excel
                            _skip_slow_file(job, path, (job.get("scan_limit") or {}).get("seconds") or 0, manual=bool(manual))
                            session.close()
                            return False
                        # 检测类原因（ValueError）不重试
                        if (not isinstance(err, (ValueError, TimeoutError)) and not watchdog.timed_out
                                and tries < autofix.MAX_ERROR_RETRIES):
                            wf.log_msg(f"⚠️ {Path(path).name} 出错，重启 Excel 后重试：{err}")
                            try:
                                with scan_analysis.recovery(job, path, enabled=work is scan):
                                    _restart_excel(session)
                            except Exception as restart_err:
                                wf.log_msg(f"❌ 无法重新启动 Excel：{restart_err}")
                                err = restart_err
                            else:
                                continue
                        # 下一个文件用新的 Excel，避免一次崩溃拖垮后面所有文件；也别让它占着文件（用户要打开修）
                        session.close()
                        message = file_problem.message_for(err, watchdog)
                        label = workbook_rel_name(Path(path), wf.STATE["folder"])
                        if file_problem.should_ask(err, message) and not _stopping(job):
                            # 文件本身有问题（打不开、要密码、Excel 弹框卡住、被占用等）：停下来问用户
                            wf.log_msg(f"⏸ {Path(path).name} 出错，等你选择重试还是跳过：{message}")
                            update_file_progress(label, status="processing", error=message,
                                                 current_action="出错了，等你选择重试还是跳过")
                            choice = file_problem.ask(path, label, message, folder=wf.STATE["folder"],
                                                      tick=watchdog.progress, should_stop=lambda: _stopping(job))
                            watchdog.progress()
                            if choice == "retry":
                                wf.log_msg(f"🔁 重试 {Path(path).name}")
                                update_file_progress(label, status="processing", error="",
                                                     current_action="正在重试...")
                                tries = -1
                                continue
                        wf.log_msg(f"❌ {Path(path).name} 已跳过，继续其他文件，可续跑：{message}")
                        with wf.STATE_LOCK:
                            job.setdefault("file_errors", {})[path] = message
                            wf._persist_state()
                        update_file_progress(label, status="error", error=message, current_action="文件未完成，可续跑")
                        return False
                    finally:
                        session.drop_copy(path)

            def scan_body(path: str) -> None:
                debug_log.begin(f"扫描 {expected.index(path) + 1} / {len(expected)}："
                                  f"{workbook_rel_name(Path(path), wf.STATE['folder'])}")
                steps: dict[str, float] = {}
                mark = file_started = time.perf_counter()

                def step(name: str) -> None:
                    # 每步用了几秒，文件做完写一行「⏱ 分步」（排查扫描中途空等，用户 2026-11-03 要求）
                    nonlocal mark
                    now = time.perf_counter()
                    steps[name] = round(steps.get(name, 0.0) + now - mark, 2)
                    mark = now

                _check_file(path)
                # 检测这一步进度框能「跳过这一份」（用户 2026-11-11 要）；上一份跳过时 Excel 已结束，换新的
                if watchdog.file(workbook_rel_name(Path(path), wf.STATE["folder"])):
                    session.close()
                excel = session.app()
                step("查文件和Excel")
                # 只给检测（只读打开原文件）设上限：超时结束 Excel 不会弄坏文件；之后标色要写文件，不在上限内
                with wf.STATE_LOCK:
                    limit = _scan_limit(job, path)
                    job["scan_limit"] = {"path": path, "seconds": round(limit)}
                started, checking, view = time.perf_counter(), {}, {}
                watchdog.deadline(limit)
                step("定上限")
                try:
                    # 末页要不要取整在「调整页面」由用户定（2026-11-07），扫描不再自动缩
                    _refresh(excel, path, session=session, timing=checking, compact_print_tail=False, view=view,
                             before_print_save=lambda: (watchdog.deadline(None), watchdog.file(None)))
                finally:
                    watchdog.deadline(None)
                    watchdog.file(None)  # 后面标色要写文件，不能中途结束 Excel
                step("检测")
                with wf.STATE_LOCK:
                    job.setdefault("scan_seconds", {})[path] = round(time.perf_counter() - started, 1)
                    _data().get("slow_files", {}).pop(path, None)  # 以前超时跳过的，这次扫完了
                wf.log_msg(f"⏱ {Path(path).name} 整份检查：{time.perf_counter() - started:.1f} 秒"
                           f"（导出 {int(checking.get('exports', 0))} 次 {checking.get('export', 0):.1f} 秒，"
                           f"检测 {checking.get('detect', 0):.1f} 秒）")
                # 检出的全部转人工，在原 Excel 里标色方便定位
                with wf.STATE_LOCK:
                    _data().setdefault("highlight_files", {})[path] = highlight_style(options.get("highlight"))
                _hand_over(path)
                step("转人工")
                # 分页预览（只改视图，不改选区或缩放）和标色一起打开、一起保存，少开存一次文件（用户 2026-11-07 同意）；
                # 没有要标色的文件不会打开，照旧单独存视图
                view_needed = view.get("needed", True)
                with scan_analysis.operation("highlight"):
                    viewed = _sync_highlights(session.app, path, tick=session.touch,
                                              scan_view=session if view_needed else None)
                step("标色")
                if view_needed and not viewed:
                    with scan_analysis.operation("save_view"):
                        _save_scan_preview(excel, path, session)
                step("存分页预览")
                with wf.STATE_LOCK:
                    job.setdefault("file_seconds", {})[path] = round(time.perf_counter() - started, 1)
                debug_log.event("⏱ 分步", 文件=Path(path).name, 合计秒=round(time.perf_counter() - file_started, 2),
                                **{k: v for k, v in steps.items()})

            def scan(path: str) -> None:
                with scan_analysis.attempt(job, path):
                    scan_body(path)

            def scan_and_record(path: str) -> None:
                if not each_file(path, scan):
                    return
                with wf.STATE_LOCK:
                    scanned.append(path)
                    versions[path] = wf._file_version(Path(path))
                    job.get("file_errors", {}).pop(path, None)
                    job["completed"].append(path)
                    wf.STATE["scan_progress"] = {"total": len(expected), "completed": len(job["completed"]),
                                                 "status": "running"}
                    wf._persist_state()

            for path in list(expected):
                if path in scanned:
                    continue
                if job.get("pause_requested") or watchdog.stop_requested is True:
                    job["pause_requested"] = True  # 进度框「结束整个任务」= 暂停，之后可以继续
                    break
                # 一个文件做完才存一次记录（以前每步都存，整份几 MB，一个文件要多花四五秒；用户 2026-11-03 同意）。
                # 出错、跳过时离开这里也会存
                with wf.deferred_persist():
                    scan_and_record(path)
            scan_analysis.finish(job, interrupted=bool(job.get("pause_requested")))
            with wf.STATE_LOCK:
                job["closing"] = True
                wf._persist_state()
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
                # 后端也执行上限，旧任务续跑或直接调用接口不会绕过每次打开的数量。
                # 先去掉别处已打开的，再取前十份，避免名额被已打开的文件占掉。
                selected = options.get("selected_files") is True
                with _opening_files():
                    with wf.STATE_LOCK:
                        opened = set() if selected else _opened_paths()
                        folder_key, round_ = _open_folder_key(), _open_round
                        if not selected:
                            expected = [path for path in expected
                                        if path in job["completed"] or path not in opened][:DEFAULT_OPEN_LIMIT]
                        job["workbooks"] = expected
                        wf._persist_state()
                    targets = [(Path(path), c["sheet_name"], c["cell_addr"], c["excel_row"])
                               for path in expected if path not in job["completed"] and path not in opened
                               for c in [next(c for c in cards if c["wb_path"] == path)]]
                    errors = open_workbooks_for_review(targets) if targets else {}
                    _remember_opened([str(p) for p, *_ in targets if p not in errors],
                                     folder_key=folder_key, round_=round_)
                job["completed"] = list(dict.fromkeys(job["completed"] + [path for path in expected if Path(path) not in errors]))
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
                        if cards:
                            _forget_opened(path)
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
        elif kind == "batch_fix":
            # 单元格工具箱：按网页上选好的工具改格子、保存、整份复检（用户 2026-11-08 要）
            from core import cell_toolbox
            expected = cell_toolbox.run(job, session, watchdog)
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
        # 收尾每步记时间：2026-11-04 用户扫完 327 / 327 后页面一直停在「正在扫描」，日志里看不出卡在哪一步
        debug_log.event("🏁 收尾：写执行记录")
        ending = time.perf_counter()
        finish_execution_batch("处理完成，请审核结果" if set(expected).issubset(job["completed"]) else "部分文件尚未完成，可继续任务")
        debug_log.event("🏁 收尾：执行记录写完", 秒=round(time.perf_counter() - ending, 2))
        return expected
    finally:
        scan_analysis.finish(job, interrupted=True)
        try:
            with wf.STATE_LOCK:
                _show_deferred_pending()
                wf._persist_state()
        finally:
            try:
                debug_log.event("🏁 收尾：关后台 Excel")
                closing = time.perf_counter()
                session.close()
                debug_log.event("🏁 收尾：后台 Excel 已关", 秒=round(time.perf_counter() - closing, 2))
            finally:
                if watchdog is not None:
                    watchdog.stop()
                wf.pythoncom.CoUninitialize()
