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
from core.excel_watchdog import ExcelWatchdog
from core.hidden_text import (
    REASON_CLIPPED, REASON_TRUNCATED, collect_text_pieces, is_excel_error, locate_text, normalize_text,
)
from core.pdf_detector import fitz, render_crop_base64
from core.reporter import finish_execution_batch, init_execution_batch, update_file_progress

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


def _data() -> dict[str, Any]:
    return wf.STATE.setdefault("batch_review", {"cards": {}, "rows": {}, "ignored": []})


def _marks() -> dict[str, dict]:
    """已在原 Excel 里标色的单元格（按卡片 id），记着原来的底色和字色，人工处理完再还原"""
    return _data().setdefault("highlights", {})


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


def _unchanged_since_check(path: str) -> bool:
    """上次整份检测之后原文件没再改过（检测规则也没变）：重新导出检测的结果和上次一样"""
    checks = _sheet_checks().get(path) or {}
    try:
        current = wf._file_version(Path(path))
    except OSError:
        return False
    return checks.get("*") == current


def rename_sheet(path: str, old: str, new: str) -> None:
    """工作表被人改了名（复检时按 Excel 代码名认出来）：审核卡片、行高记录、忽略规则、标色记录都换成新表名。
    卡片和行高记录的编号由表名算出，一起换成新编号。调用方持有 STATE_LOCK。"""
    data = _data()
    moved: dict[str, str] = {}
    for cid, card in list(data["cards"].items()):
        if card["wb_path"] == path and card["sheet_name"] == old:
            card["sheet_name"] = new
            card["id"] = _key(*_cell(card))
            moved[cid] = card["id"]
            data["cards"][card["id"]] = data["cards"].pop(cid)
    for rid, row in list(data["rows"].items()):
        if row["path"] == path and row["sheet"] == old:
            row["sheet"] = new
            row["id"] = (_key(path, new, int(row["row"]), "break") if row.get("kind") == "break"
                         else _key(path, new, int(row["row"])))
            data["rows"][row["id"]] = data["rows"].pop(rid)
    for rule in data.get("ignored", []):
        if rule["cell"][0] == path and rule["cell"][1] == old:
            rule["cell"][1] = new
    marks = _marks()
    for cid, mark in list(marks.items()):
        if mark["path"] == path and mark["sheet"] == old:
            mark["sheet"] = new
            marks[moved.get(cid, cid)] = marks.pop(cid)
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
    text, visible = normalize_text(item.get("text", "")), normalize_text(item.get("visible_text") or "")
    missing = text[len(visible):] if kind == "hidden" and visible and text.startswith(visible) else ""
    reason = {"hidden": item.get("hidden_reason"), "pagebreak": page_breaks.REASON}.get(kind) or CLIP_REASON
    return {"kind": kind, "label": KIND_LABELS.get(kind, kind), "reason": reason, "missing": missing}


def _remember_problems(card: dict, items: list[dict]) -> None:
    problems = {p["kind"]: p for p in card.get("problems", [])}
    problems.update((p["kind"], p) for p in map(_problem, items))  # 同一类型以最新一次检测为准
    card["problems"] = list(problems.values())


def _row_record(card: dict) -> dict | None:
    return _data()["rows"].get(_rid(card))


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
        for card in sorted(_data()["cards"].values(), key=lambda c: c["seq"]):
            if card.get("dismissed") or not wf.in_current_folder(card["wb_path"]):
                continue
            result = copy.deepcopy(card)
            if not result.get("problems"):  # 旧版本留下的卡片没存问题说明，按单元格现有的问题项补上
                result["problems"] = list({p["kind"]: p for p in map(_problem, _items_at(card))}.values())
            row = _row_record(card)
            if row is not None and _is_break(card):
                result["page_break"] = {"row": row["row"], "saved": row["saved"], "applied": row["applied"] >= .5}
            elif row is not None:
                result["row_height"] = {"base": row["base"], "applied": row["applied"], "saved": row["saved"],
                                        "wrap": card.get("cell_addr") in row.get("wrap", [])}
            result["adoptable"] = _adoptable(card)
            result["highlighted"] = _shows_mark(_marks().get(card["id"]))
            result["images"] = {
                which: _image_url(name) for which, name in card.get("fix_images", {}).items()
                if fix_images.image_path(name)
            }
            if result.get("context"):
                context = result["context"]
                name = context["image_file"]
                context["image"] = _image_url(name) if fix_images.image_path(name) else ""
            cards.append(result)
        highlighted = sum(1 for m in _marks().values() if wf.in_current_folder(m["path"]) and _shows_mark(m))
        return {"cards": cards, "busy": wf.STATE["busy"], "folder": wf.STATE["folder"], "highlighted": highlighted,
                "mode": _mode(),
                "job": copy.deepcopy(wf.STATE["job"]), "progress": dict(wf.STATE["scan_progress"]),
                "logs": wf.STATE["logs"][-100:]}


def mark(cid: str, color: str) -> None:
    with wf.STATE_LOCK:
        if wf.STATE["busy"]:
            raise ValueError("任务执行中，请结束或暂停后再审核")
        card = _data()["cards"].get(cid) if isinstance(cid, str) else None
        if color not in COLORS or card is None or card.get("dismissed") or not wf.in_current_folder(card["wb_path"]):
            raise ValueError("请选择当前文件夹的有效卡片及颜色")
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
        if wf.STATE["busy"]:
            raise ValueError("任务执行中，请结束或暂停后再审核")
        if color not in COLORS or not isinstance(ids, list) or not ids:
            raise ValueError("请选择卡片及颜色")
        cards = _data()["cards"]
        targets = []
        for cid in dict.fromkeys(ids):
            card = cards.get(cid) if isinstance(cid, str) else None
            if card is None or card.get("dismissed") or not wf.in_current_folder(card["wb_path"]):
                raise ValueError("选中的卡片已失效，请刷新页面后重新选择")
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


def _row_value(book: Any, row: dict) -> float:
    """写入日志里一行现在的值：行高；插分页的行是 1（上方有手动分页）/ 0（没有）"""
    target = book.Sheets(row["sheet"]).Rows(row["row"])
    if row.get("kind") == "break":
        return 1.0 if int(target.PageBreak) == page_breaks.XL_PAGEBREAK_MANUAL else 0.0
    return float(target.RowHeight)


def _write_rows(excel: Any, journal: dict) -> None:
    """先核对所有行，再写；续跑时目标行高已保存则只补数据库提交。"""
    path = journal["path"]
    where = f"{Path(path).name}（" + "、".join(f"{r['sheet']} 第 {r['row']} 行" for r in journal["rows"]) + "）"
    _check_file(path)
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError(f"{where} 文件以只读方式打开，未写入")
        actual = [_row_value(book, r) for r in journal["rows"]]
        for value, r in zip(actual, journal["rows"]):
            debug_log.event("  写入行", 动作=journal.get("action"), 位置=f"{r['sheet']} 第 {r['row']} 行",
                            类型=r.get("kind") or "行高", 现在=value, 预期原值=r["expected"], 目标=r["target"],
                            换行=r.get("wrap") or "-")
        # 自动换行和行高一起保存；续跑时两样都已到位才算写过
        wraps = [(r["sheet"], addr, on) for r in journal["rows"] for addr, on in (r.get("wrap") or {}).items()]
        at_target = (all(abs(h - r["target"]) < ROW_HEIGHT_TOLERANCE for h, r in zip(actual, journal["rows"]))
                     and all(bool(book.Sheets(sh).Range(addr).MergeArea.WrapText) == on for sh, addr, on in wraps))
        if wf._file_version(Path(path)) != journal["version"] and not at_target:
            raise ValueError(f"{where} 文件在准备后被修改，已停止写入；核对后可继续任务，或点「放弃这次写入」")
        if not at_target:
            if any(abs(h - r["expected"]) >= ROW_HEIGHT_TOLERANCE for h, r in zip(actual, journal["rows"])):
                raise ValueError(f"{where} 行高已被其他操作修改，未覆盖人工修改；可点「放弃这次写入」转人工处理")
            for sheet, addr, on in wraps:
                book.Sheets(sheet).Range(addr).MergeArea.WrapText = on
            for row in journal["rows"]:
                target = book.Sheets(row["sheet"]).Rows(row["row"])
                if row.get("kind") == "break":
                    target.PageBreak = page_breaks.XL_PAGEBREAK_MANUAL if row["target"] else page_breaks.XL_PAGEBREAK_NONE
                else:
                    target.RowHeight = row["target"]
            normalize_workbook_view(excel, book)
            book.Save()
            debug_log.event("  已保存", 文件=Path(path).name,
                            写后=[round(_row_value(book, r), 2) for r in journal["rows"]])
        else:
            debug_log.event("  目标值已在文件里（续跑），不再写", 文件=Path(path).name)
    finally:
        book.Close(SaveChanges=False)


def _wrapped(pending: dict, card: dict) -> list[str]:
    """这次写入在卡片所在行开启自动换行的单元格"""
    return [addr for r in pending["rows"] if r["row"] == card["excel_row"] and r["sheet"] == card["sheet_name"]
            for addr, on in (r.get("wrap") or {}).items() if on]


def _finish_pending(excel: Any) -> None:
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
            if card:
                card.update(color="yellow", verified=False, waiting=False,
                            message=f"文件写入状态待核对，已跳过并保留恢复日志：{error}")
        wf._persist_state()


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
            card.update(color="yellow", verified=False, manual=True, finalized=True, dismissed=False, waiting=False,
                        message="已放弃未完成的自动写入，请在 Excel 中核对这一行的行高，保存后复检")
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
                if row and row["saved"] and abs(row["applied"] - row["base"]) >= ROW_HEIGHT_TOLERANCE:
                    existing.update(color="blue", verified=True, after_kind="fixed", waiting=False,
                                    message=("已插入分页" if _is_break(existing) else "已加高")
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
            debug_log.event("    试算", 次=attempt, 行高=height, 原行高=base, 通过=ok, 缺口pt=need,
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
                           if c["wb_path"] == path and (c.get("manual") or c.get("marked")) and not c.get("dismissed")}
            manual_rows.update(r["id"] for r in _data()["rows"].values()
                               if r["path"] == path and r["saved"]
                               and any(not c.get("dismissed") and c["wb_path"] == path and _rid(c) == r["id"]
                                       for c in _data()["cards"].values()))
            candidates = [it for it in pending if _rid(it) not in attempted and _rid(it) not in manual_rows]
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
            items = [it for it in pending if _rid(it) == rid]
            cards = [_card_for(group) for group in _group_cells(items)]
            wf._persist_state()
        _check_file(path)
        version = wf._file_version(Path(path))
        debug_log.event("  处理行", 位置=f"{first['sheet_name']} 第 {first['excel_row']} 行",
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
            if _is_break(first):
                workbook = _repair_break(excel, session, path, items[0], cards, rid, version, timing)
                raise _RowDone
            column_only = [it for it in items if autofix.column_width_only(it)]
            if column_only:
                # ### 交给人工改列宽；同行其他格子照常自动加高
                with wf.STATE_LOCK:
                    for group in _group_cells(column_only):
                        card = _card_for(group)
                        card.update(color="yellow", verified=False, manual=True, waiting=False,
                                    message=autofix.COLUMN_ONLY_REASON)
                    wf.log_msg(f"⚠️ {first['wb_name']} / {first['sheet_name']} 行{first['excel_row']}："
                               f"{autofix.COLUMN_ONLY_REASON}")
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


def _commit_review(excel: Any) -> None:
    job = wf.STATE["job"]
    data = _data()
    if "commit_card_ids" not in job["options"]:
        # 还没轮到自动处理的卡不提交：提交会把它们转成人工项，之后再跑全自动也不会修了
        cards = [data["cards"][cid] for cid in job["options"]["card_ids"] if not _waiting(data["cards"][cid])]
        selected_rows = {_rid(c) for c in cards}
        job["options"]["commit_card_ids"] = [c["id"] for c in data["cards"].values()
            if not c.get("dismissed") and not _waiting(c) and _rid(c) in selected_rows]
    # 提交计划固定下来；即使保存还原后、复检前退出，已收起的绿/蓝卡也仍会完成该文件的复检。
    job["options"]["card_ids"] = list(job["options"]["commit_card_ids"])
    cards = [data["cards"][cid] for cid in job["options"]["commit_card_ids"]]
    by_row: dict[str, list[dict]] = {}
    for card in cards:
        by_row.setdefault(_rid(card), []).append(card)
    # 只在仍有自动修改需要还原时检查同行冲突；已经忽略的格子不因后来复检重新入队。
    # 人工认可的加高也改整行行高：同行颜色不一致同样整行转黄
    for rid, peers in by_row.items():
        row = data["rows"].get(rid)
        if ((row and row["saved"]) or any(_adopting(c) for c in peers)) and len({c["color"] for c in peers}) > 1:
            for card in peers:
                card.update(color="yellow", conflict=True, message="同行存在不同处理意见，整行还原后交给人工")
    wf._persist_state()
    debug_log.event("提交审核", 卡片=len(cards), **{color: sum(c["color"] == color for c in cards)
                                                      for color in ("blue", "green", "yellow")})
    for path in dict.fromkeys(c["wb_path"] for c in cards):
        if path in job["completed"]:
            continue
        debug_log.section(f"提交 {Path(path).name}")
        rows = []
        reverting = []
        for rid, peers in by_row.items():
            row = data["rows"].get(rid)
            if not row or row["path"] != path or not row["saved"] or all(c["color"] == "blue" for c in peers):
                continue
            rows.append({"id": rid, "kind": row.get("kind"), "sheet": row["sheet"], "row": row["row"],
                         "expected": row["applied"], "target": row["base"],
                         "wrap": {addr: False for addr in row.get("wrap", [])}})
            reverting.extend(c["id"] for c in peers)
        if rows:
            job["options"].setdefault("refresh_sheets", {})[path] = sorted({r["sheet"] for r in rows})
            job["pending_adjustment"] = {"action": "restore", "path": path,
                                         "version": wf._file_version(Path(path)), "rows": rows, "cards": reverting}
            wf._persist_state()
            _finish_pending(excel)
        _adopt_rows(excel, path, by_row)
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
                    card.update(verified=False, message=("同行存在不同处理意见，整行已还原；" if card.get("conflict") else "")
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


def _open_without_marks(excel: Any, path: str) -> Any:
    """复检用：只读打开原文件，在内存里把标过的颜色全部还原（不保存），检测的就是最后交出去的样子。
    没标过颜色返回 None，照常由检测自己打开。"""
    with wf.STATE_LOCK:
        marks = [copy.deepcopy(m) for m in _marks().values() if m["path"] == path]
    if not marks:
        return None
    book = open_workbook_checked(excel, str(Path(path).resolve()), UpdateLinks=0, ReadOnly=True)
    failed: list[str] = []
    try:
        for mark in marks:
            try:
                found = _locate_mark(book, mark)
                if found is not None and found[0] is not None:
                    _restore_cells(found[0], mark, found[1])
            except Exception as err:
                # 一格还原不了（例如工作表被保护）不该让整份复检失败：照常检测，日志里说一声
                failed.append(f"{mark['sheet']}!{mark['addr']}（{err}）")
    except BaseException:
        book.Close(SaveChanges=False)
        raise
    if failed:
        wf.log_msg(f"⚠️ {Path(path).name}：{len(failed)} 格标记颜色没能在检测前去掉，带着标记色检测："
                   + "；".join(failed[:5]) + ("……" if len(failed) > 5 else ""))
    return book


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


def _back_to_auto(excel: Any, session: Any, path: str) -> None:
    """「转回自动」时每份文件先做的事：转人工后文件改过就整份重扫（规则同复检，原来的黄卡翻篇）；
    还没被自动试过、所在行也没被人动过的问题交回自动队列，其余留给人工；最后还原 Excel 里的标记颜色，
    自动处理的截图和检测都按原来的颜色来。"""
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
        unresolved = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                      and it["status"] not in wf.RESOLVED_STATUSES]
        back = kept = 0
        for group in _group_cells(unresolved):
            card = _card_for(group)
            if not card.get("handed_over") or card.get("dismissed") or (
                    card.get("marked") and card["color"] != "yellow" and not card.get("finalized")):
                continue  # 自动试过没通过的、你已标蓝 / 绿还没提交的，保持原样
            if not unchanged and any(_touched(it, rows) for it in group):
                card.update(handed_over=False, message=TOUCHED_MESSAGE)
                kept += 1
                continue
            card.update(color="yellow", verified=False, manual=False, finalized=False, marked=False,
                        waiting=True, handed_over=False, message=WAITING_MESSAGE)
            card.pop("after_kind", None)
            for item in group:
                item.update(status="pending", batch_managed=True, triage_keep=False)
            back += 1
        wf._persist_state()
    # 自动处理时 Excel 里不留标记色；还留给人工的格子，提交或复检后再按设置标色
    _sync_highlights(session.app, path, clear=True)
    wf.log_msg(f"🤖 {Path(path).name}：{back} 处交回自动处理"
               + (f"，{kept} 处所在行转人工后被改过，留给人工" if kept else ""))


def _handed_over_paths() -> list[str]:
    """调用方持有 STATE_LOCK：当前文件夹里还有可交回自动的问题的文件"""
    cards = sorted(_data()["cards"].values(), key=lambda c: c["seq"])
    return list(dict.fromkeys(c["wb_path"] for c in cards if c.get("handed_over") and not c.get("dismissed")
                              and wf.in_current_folder(c["wb_path"])))


def _mode() -> dict[str, bool]:
    """调用方持有 STATE_LOCK：现在能不能「转人工」/「转回自动」"""
    job = wf.STATE["job"] or {}
    auto = job.get("kind") == "batch_scan" and not job.get("options", {}).get("scan_only")
    active = wf.STATE["busy"] or job.get("status") == "interrupted"
    return {
        # 全部文件扫描完才能切；暂停中也可以直接转
        "to_manual": auto and active and job.get("phase") == "repair" and not job.get("to_manual"),
        "to_auto": not wf.STATE["busy"] and not (auto and job.get("status") == "interrupted")
                   and bool(_handed_over_paths()),
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


def check_auto() -> None:
    with wf.STATE_LOCK:
        if not _mode()["to_auto"]:
            raise ValueError("没有可以交回自动处理的问题")


def _can_highlight(card: dict) -> bool:
    return bool(card.get("cell_addr")) and all(it.get("location_confirmed") is not False for it in _items_at(card))


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


def _sweep_leftover_marks(book: Any, path: str, missed: list[dict], kept: list[dict]) -> None:
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
    fixed, left = 0, []
    for mark in missed:
        sheet = _find_sheet(book, mark["sheet"])
        if sheet is None or not mark.get("value"):
            continue  # 表不在了（前面已说明）或旧记录没记内容：没法找
        try:
            spots = [rc for rc in _marked_matches(sheet, mark["value"], mark["fill"])
                     if (sheet.Name, *rc) not in keep]
            debug_log.event("      兜底查找", 位置=f"{mark['sheet']}!{mark['addr']}", 内容=debug_log.clip(mark["value"]),
                            找到=[f"{_col_letters(c)}{r}" for r, c in spots[:5]])
            if len(spots) == 1:
                r0, c0 = _cell_rc(mark["addr"])
                if _restore_cells(sheet, mark, (spots[0][0] - r0, spots[0][1] - c0)):
                    fixed += 1
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


def _col_letters(col: int) -> str:
    letters = ""
    while col:
        col, rest = divmod(col - 1, 26)
        letters = chr(65 + rest) + letters
    return letters


def _whole_area(area: Any) -> dict | None:
    """合并区域里每格底色、字色都一样（绝大多数情况）时整块读一次、记一条，
    不用一格一格跨进程读（方眼紙的合并区域常有几十格，问题格多的文件会慢很多）。颜色不一致返回 None。"""
    try:
        if ":" not in str(area.Address):
            return None  # 单格照常读
        interior = area.Interior
        values = [interior.Pattern, interior.Color, interior.TintAndShade, area.Font.Color]
        if any(v is None for v in values):
            return None
        return {"addr": str(area.Address), "pattern": int(values[0]), "color": int(values[1]),
                "tint": float(values[2] or 0), "font": int(values[3])}
    except Exception:
        return None


def _apply_cells(sheet: Any, card: dict, style: dict) -> dict:
    """记下合并区域里每一格原来的底色、字色，再标色。只改颜色，不改字体粗细大小，不影响排版和分页。
    字色混合（同一格里有几种颜色的文字）时不改字色，免得还原时丢掉原来的分段颜色。"""
    addr = card["cell_addr"].split(":")[0]  # 分页切开的卡片是整个合并区域（A10:H12），按左上角找合并区域
    area = sheet.Range(addr).MergeArea
    mark = {"path": card["wb_path"], "sheet": card["sheet_name"], "addr": addr,
            "fill": _excel_color(style["fill"]), "font": _excel_color(style["font"]), "cells": []}
    try:
        mark["value"] = str(sheet.Range(addr).Formula)  # 兜底：格子挪了位置又找不到时按内容认
    except Exception:
        pass
    whole = _whole_area(area)
    if whole is not None:
        mark["cells"].append(whole)
        return mark
    for cell in area.Cells:
        interior, font = cell.Interior, cell.Font.Color
        mark["cells"].append({"addr": cell.Address, "pattern": int(interior.Pattern), "color": int(interior.Color),
                              "tint": float(interior.TintAndShade or 0),
                              "font": None if font is None else int(font)})
    return mark


def _sync_highlights(excel_for: Callable[[], Any], path: str, *, clear: bool = False,
                     tick: Callable[[], None] | None = None) -> None:
    """让原 Excel 里的标色和卡片一致：只扫描过的文件里，黄色人工待处理的格子标色；
    复检通过（变蓝）、提交为蓝 / 绿、或要求清除时还原原来的颜色。打开一次、保存一次。
    tick：每处理一格告诉看门狗还在动，问题格很多的大文件标色要好几分钟，不会被当成 Excel 卡死。"""
    with wf.STATE_LOCK:
        data, marks = _data(), _marks()
        style = data.get("highlight_files", {}).get(path)
        cards = [c for c in data["cards"].values() if c["wb_path"] == path]
        want = set() if clear else _wanted_marks(path)
        have = {cid for cid, m in marks.items() if m["path"] == path}
        # 没确认存进文件的记录（上次标色被打断，或旧版本没记这一项）：打开文件看一眼颜色还在不在
        unsure = sorted(cid for cid in have if not marks[cid].get("saved"))
        add = [c for c in cards if c["id"] in want - have]
        remove = sorted(have - want)
    if not add and not remove and not unsure:
        return
    _check_file(path)
    debug_log.event("🎨 同步标色", 文件=Path(path).name, 清除=clear, 要标=len(add), 要还原=len(remove),
                    待核对=len(unsure), 样式=style or "-")
    app = excel_for()
    book = open_workbook_checked(app, str(Path(path).resolve()), UpdateLinks=0)
    added: dict[str, dict] = {}
    failed: dict[str, str] = {}
    deleted = 0
    book_name, last_log = Path(path).name, time.monotonic()

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
        lost = []
        for i, cid in enumerate(unsure, 1):
            if _still_marked(book, marks[cid]):
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

        for i, card in enumerate(add, 1):
            step("记下原来的颜色", i, len(add))
            sheet = sheet_of(card["sheet_name"])
            if sheet is None:
                failed[card["id"]] = f"工作表「{card['sheet_name']}」已不在文件里，可能已被删除或改名"
                continue
            try:
                added[card["id"]] = _apply_cells(sheet, card, style)
                added[card["id"]]["name"] = _mark_name(card["id"])
                debug_log.event("    标色", 位置=f"{card['sheet_name']}!{added[card['id']]['addr']}",
                                原色=[(c["addr"], c.get("pattern"), c.get("color"), c.get("font"))
                                      for c in added[card["id"]]["cells"][:4]],
                                格数=len(added[card["id"]]["cells"]))
            except Exception as err:
                failed[card["id"]] = str(err)
                debug_log.event("    标色失败", 位置=f"{card['sheet_name']}!{card.get('cell_addr')}", 原因=err)
                debug_log.exception(err, "标色")
        with wf.STATE_LOCK:
            # 先记下原来的颜色再改、再保存：中途出错时记录还在，之后还能还原（不是标记颜色的格子不会被改）
            for mark in added.values():
                mark["saved"] = False  # 保存成功后才算标上了
            marks.update(added)
            wf._persist_state()
        for i, mark in enumerate(added.values(), 1):
            step("标色", i, len(added))
            sheet = sheet_of(mark["sheet"])
            area = sheet.Range(mark["addr"]).MergeArea
            try:
                # 隐藏名称跟着格子走：人工插删行列、改表名后，Excel 会自动更新它指向的位置，还原时按它找到格子
                _delete_name(book, mark["name"])
                book.Names.Add(Name=mark["name"], RefersTo=f"={_sheet_ref(sheet.Name)}!{area.Address}", Visible=False)
            except Exception:
                mark.pop("name", None)  # 建不了名称就只能按原地址还原
            area.Interior.Pattern = XL_PATTERN_SOLID
            area.Interior.Color = mark["fill"]
            for cell in mark["cells"]:
                if cell["font"] is not None:
                    sheet.Range(cell["addr"]).Font.Color = mark["font"]
        gone: dict[str, int] = {}
        missed: list[dict] = []  # 没在原位置找到标记色的记录，留给下面的兜底检查
        for i, cid in enumerate(remove, 1):
            step("还原颜色", i, len(remove))
            mark = marks[cid]
            found = _locate_mark(book, mark)
            debug_log.event("    还原", 位置=f"{mark['sheet']}!{mark['addr']}", 名称=mark.get("name"),
                            找到=found is not None and found[0] is not None,
                            挪位=found[1] if found and found[0] is not None else "-",
                            表在=not (found is not None and found[0] is None))
            _delete_name(book, mark.get("name"))
            if found is None:
                deleted += 1  # 标色的格子已被人工删掉，没有颜色要还原
                continue
            sheet, shift = found
            if sheet is None:
                # 工作表已被删除（例如精简 Sheet 或人工处理时删掉）就没有颜色要还原了，按表汇总说一次
                gone[mark["sheet"]] = gone.get(mark["sheet"], 0) + 1
                missed.append(mark)
                continue
            try:
                if not _restore_cells(sheet, mark, shift):
                    debug_log.event("      原位置不是标记色，留给兜底查找", 位置=f"{mark['sheet']}!{mark['addr']}")
                    missed.append(mark)
            except Exception as err:
                wf.log_msg(f"⚠️ {Path(path).name} / {mark['sheet']}!{mark['addr']} 未能还原颜色：{err}")
        if missed:
            kept = [m for cid, m in marks.items() if m["path"] == path and cid not in remove]
            _sweep_leftover_marks(book, path, missed, kept)
        for name, count in gone.items():
            wf.log_msg(f"ℹ️ {Path(path).name}：工作表「{name}」已不在文件里（被删除或改名），"
                       f"跳过还原 {count} 格标记颜色；如果是改名，请在新表里手动去掉标记色")
        if added or remove:  # 只核对了上次的标色、没有要改的：不保存，文件保持原样
            if tick is not None:
                tick()
            normalize_workbook_view(app, book)
            book.Save()
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
            marks.pop(cid, None)
        for cid, error in failed.items():
            card = data["cards"].get(cid)
            if card:
                card["message"] = f"未能在 Excel 中标色（{error}），请按截图查找后人工处理"
        wf._persist_state()
    parts = [f"标色 {len(added)} 格"] if added else []
    if len(remove) > sum(gone.values()) + deleted:
        parts.append(f"还原颜色 {len(remove) - sum(gone.values()) - deleted} 格")
    if deleted:
        parts.append(f"{deleted} 格标色的格子已被删掉，不用还原")
    if failed:
        parts.append(f"{len(failed)} 格未能标色（工作表受保护等）")
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
        return [p for p in paths if {cid for cid, m in _marks().items() if m["path"] == p} - _wanted_marks(p)]


def _where(err: BaseException) -> str:
    """出错在哪个函数哪一行（只看本程序的代码），日志里带上，下次出问题能直接定位"""
    frames = [f for f in traceback.extract_tb(err.__traceback__) if "core" in Path(f.filename).parts]
    return f"（出错位置：{frames[-1].name} 第 {frames[-1].lineno} 行）" if frames else ""


def _sync_highlights_safely(session: Any, path: str) -> None:
    """提交、复检之后顺带更新标色；失败只记日志，不影响审核结果。
    Excel 出毛病（崩溃、连不上）时换一个新的 Excel 再试一次；还是不行，记录留着，下次提交 / 复检时再还原"""
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
        paths = list(dict.fromkeys(m["path"] for m in _marks().values() if wf.in_current_folder(m["path"])))
        if not paths:
            raise ValueError("当前文件夹没有标过颜色的单元格")
        return paths


def enqueue_recheck(ids: list[str]) -> bool:
    """复检任务正在跑时把新点的复检排进同一个任务（同一个后台 Excel 接着查），不用等它结束再点。
    没有在跑的复检任务、或它已经查完准备收尾时返回 False，由调用方另起任务。"""
    with wf.STATE_LOCK:
        job = wf.STATE.get("job") or {}
        if not wf.STATE.get("busy") or job.get("kind") != "batch_recheck" or job.get("closing"):
            return False
        cards = _data()["cards"]
        job["options"]["card_ids"] = list(dict.fromkeys(job["options"]["card_ids"] + ids))
        for path in dict.fromkeys(cards[cid]["wb_path"] for cid in ids):
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
        "⚠️ 后台 Excel 长时间无进展，已结束本任务专用 Excel，当前文件将记录错误并重试或跳过"),
        timeout=UNHIGHLIGHT_STALL_SECONDS if kind == "batch_unhighlight" else None)
    session.watchdog = watchdog
    if watchdog is not None:
        watchdog.start()
    wf.pythoncom.CoInitialize()
    try:
        excel = session.app() if job.get("pending_adjustment") and kind != "batch_scan" else None
        if excel is not None:
            _finish_pending(excel)
        if kind != "batch_scan" and "workbooks" not in job:
            # 审核提交 / 复检 / 打开 / 清除颜色也记下本次涉及的文件，网页按本任务显示进度
            with wf.STATE_LOCK:
                job["workbooks"] = (list(dict.fromkeys(options["paths"])) if kind == "batch_unhighlight" else
                                    list(dict.fromkeys(_data()["cards"][cid]["wb_path"] for cid in options["card_ids"])))
                wf._persist_state()
        if kind == "batch_scan":
            back = bool(options.get("back_to_auto"))
            if "workbooks" not in job:
                if back:
                    # 「转回自动」：只处理还有可交回自动的问题的文件，审核数据都在，不清空
                    with wf.STATE_LOCK:
                        folder = Path(wf.STATE["folder"])
                        job["workbooks"] = _handed_over_paths()
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
                init_execution_batch(task_name="只扫描（全部转人工）" if scan_only else "全自动扫描与修复",
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
                debug_log.section(f"{'转回自动' if back else '扫描'} {expected.index(path) + 1} / {len(expected)}："
                                  f"{workbook_rel_name(Path(path), wf.STATE['folder'])}")
                _check_file(path)
                excel = session.app()
                if back:
                    _back_to_auto(excel, session, path)
                else:
                    started, checking = time.perf_counter(), {}
                    _refresh(excel, path, session=session, timing=checking)
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

            for path in expected:
                if path in scanned:
                    continue
                if job.get("pause_requested"):
                    break
                if not each_file(path, scan):
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
                debug_log.section(f"修复 {expected.index(path) + 1} / {len(expected)}："
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

            for path in expected if job.get("phase") == "repair" else []:
                if _stopping(job):
                    break
                if path not in scanned or (path in job["completed"] and path not in job.get("deferred_adjustments", {})):
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
                    wf.STATE["scan_progress"] = {"total": len(expected), "completed": len(job["completed"]), "status": "running"}
                    wf._persist_state()
                wf._sync_problem_counts(path, "全自动处理完成，等待结果审核")
            if job.get("to_manual") and not job.get("pause_requested"):
                _switch_to_manual(session, job, [p for p in expected if p in scanned])
        elif kind == "batch_commit":
            if options.get("refresh_sheets") or any(c["color"] != "blue" and not c.get("dismissed")
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
                        job["current"] = path
                        cards = [card for cid in dict.fromkeys(options["card_ids"])
                                 for card in [_data()["cards"].get(cid)]
                                 if card and card["wb_path"] == path and not card.get("dismissed")]
                    debug_log.section(f"复检 {Path(path).name}（{len(cards)} 张卡）")
                    try:
                        unsaved = close_saved_workbooks([Path(path)])
                        if Path(path) in unsaved:
                            wf.log_msg(f"⚠️ {Path(path).name}：{unsaved[Path(path)]}，请保存后继续")
                            continue
                        if cards:
                            session.touch()
                            for tries in range(2):
                                try:
                                    _recheck_file(session.app(), path, cards)  # Excel 重启过就用新的
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
            for i, path in enumerate(expected, 1):
                if path in job["completed"]:
                    continue
                if Path(path) in unsaved:
                    wf.log_msg(f"⚠️ {Path(path).name}：{unsaved[Path(path)]}，请保存后继续")
                    continue
                debug_log.section(f"清除颜色 {i} / {len(expected)}：{Path(path).name}")
                wf.log_msg(f"🧽 正在还原（{i} / {len(expected)}）：{Path(path).name}")
                for tries in range(2):
                    try:
                        watchdog.progress()
                        _sync_highlights(session.app, path, clear=True, tick=session.touch)
                        with wf.STATE_LOCK:
                            # 清除后不再自动标色；以后再「只扫描」才重新标
                            _data().get("highlight_files", {}).pop(path, None)
                            job["completed"].append(path)
                            wf._persist_state()
                        break
                    except Exception as err:
                        if isinstance(err, ValueError) or tries:
                            wf.log_msg(f"❌ {Path(path).name} 清除标记颜色失败，可继续任务：{err}{_where(err)}")
                            break
                        # 后台 Excel 卡住被结束、或崩溃了：换一个新的再试，后面的文件也用新的
                        wf.log_msg(f"⚠️ {Path(path).name} 清除标记颜色出错，重启 Excel 后重试：{err}{_where(err)}")
                        _restart_excel(session)
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
