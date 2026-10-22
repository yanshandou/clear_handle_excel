"""无人值守扫描修复，以及按单元格进行的三色结果审核。

PDF 坐标只用于当前一次导出。每保存一行，重新检测整张表再选下一行。
审核颜色只写数据库；原文件的还原由独立、可续跑的审核任务统一执行。
"""
import copy
import hashlib
import math
from collections.abc import Callable
from pathlib import Path
from typing import Any
from uuid import uuid4

from core import autofix, fix_images, workflow as wf
from core.common import workbook_rel_name
from core.excel_ops import (
    CellProbe, close_saved_workbooks, extract_sheet_snapshot, normalize_workbook_view,
    open_workbooks_for_review, snapshot_cells,
)
from core.reporter import finish_execution_batch, init_execution_batch, update_file_progress

COLORS = ("blue", "green", "yellow")


def _data() -> dict[str, Any]:
    return wf.STATE.setdefault("batch_review", {"cards": {}, "rows": {}, "ignored": []})


def _key(*parts: Any) -> str:
    return hashlib.sha256("\0".join(map(str, parts)).encode()).hexdigest()[:24]


def _cell(item: dict) -> tuple[str, str, str]:
    return item["wb_path"], item["sheet_name"], item.get("cell_addr", "")


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
                item.update(status="accepted", manual_confirmation={"accepted": True},
                            review_reason=wf.MANUAL_CONFIRMED_REASON)
                break


def _save_shot(card: dict, which: str, image: str) -> None:
    fix_images.save(card, which, image, key=f"batch:{card['id']}")


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
            "finalized": False, "verified": False, "manual": False, "message": "等待自动处理",
            "signatures": signatures,
        }
        _save_shot(card, "before", first.get("image", ""))
    card["item_ids"] = [it["id"] for it in items]
    for signature in signatures:
        if signature not in card["signatures"]:
            card["signatures"].append(signature)
    _remember_context(card, items)
    return card


def snapshot() -> dict:
    with wf.STATE_LOCK:
        cards = []
        for card in sorted(_data()["cards"].values(), key=lambda c: c["seq"]):
            if card.get("dismissed") or not wf.in_current_folder(card["wb_path"]):
                continue
            result = copy.deepcopy(card)
            result["images"] = {
                which: f"/api/batch/image/{name}" for which, name in card.get("fix_images", {}).items()
                if fix_images.image_path(name)
            }
            if result.get("context"):
                context = result["context"]
                name = context["image_file"]
                context["image"] = f"/api/batch/image/{name}" if fix_images.image_path(name) else ""
            cards.append(result)
        return {"cards": cards, "busy": wf.STATE["busy"], "folder": wf.STATE["folder"],
                "job": copy.deepcopy(wf.STATE["job"]), "progress": dict(wf.STATE["scan_progress"]),
                "logs": wf.STATE["logs"][-100:]}


def mark(cid: str, color: str) -> None:
    with wf.STATE_LOCK:
        if wf.STATE["busy"]:
            raise ValueError("任务执行中，请结束或暂停后再审核")
        card = _data()["cards"].get(cid) if isinstance(cid, str) else None
        if color not in COLORS or card is None or card.get("dismissed") or not wf.in_current_folder(card["wb_path"]):
            raise ValueError("请选择当前文件夹的有效卡片及颜色")
        if color == "blue" and not card.get("verified"):
            raise ValueError("此处尚未通过检查，请修改 Excel 后复检，或标为绿色忽略")
        previous = copy.deepcopy(card)
        card.update(color=color, finalized=False, marked=True)
        try:
            wf._persist_state()
        except Exception:
            card.clear()
            card.update(previous)
            raise


def _check_file(path: str) -> None:
    p = Path(path)
    if not p.is_file():
        raise ValueError(f"找不到文件：{p.name}")
    if p.with_name("~$" + p.name).exists():
        raise ValueError(f"文件仍在 Excel 中打开，请保存并关闭后继续：{p.name}")


def _write_rows(excel: Any, journal: dict) -> None:
    """先核对所有行，再写；续跑时目标行高已保存则只补数据库提交。"""
    path = journal["path"]
    _check_file(path)
    book = excel.Workbooks.Open(str(Path(path).resolve()), UpdateLinks=0)
    try:
        if book.ReadOnly:
            raise ValueError("文件以只读方式打开，未写入")
        actual = [float(book.Sheets(r["sheet"]).Rows(r["row"]).RowHeight) for r in journal["rows"]]
        at_target = all(abs(h - r["target"]) < .05 for h, r in zip(actual, journal["rows"]))
        if wf._file_version(Path(path)) != journal["version"] and not at_target:
            raise ValueError("文件在准备后被修改，已停止写入；请核对文件后继续")
        if not at_target:
            if any(abs(h - r["expected"]) >= .05 for h, r in zip(actual, journal["rows"])):
                raise ValueError("行高已被其他操作修改，未覆盖人工修改")
            for row in journal["rows"]:
                book.Sheets(row["sheet"]).Rows(row["row"]).RowHeight = row["target"]
            normalize_workbook_view(excel, book)
            book.Save()
    finally:
        book.Close(SaveChanges=False)


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
            record["saved"] = pending["action"] == "fix"
            record["applied"] = row["target"]
            if pending["action"] == "fix" and row["id"] not in job.setdefault("attempted_rows", []):
                job["attempted_rows"].append(row["id"])
        for cid in pending["cards"]:
            card = data["cards"][cid]
            if pending["action"] == "fix":
                card.update(color="blue", verified=True, after_kind="fixed", message="已自动修复并保存，等待审核")
                for item in _items_at(card):
                    item.update(status="fixed", batch_managed=True, triage_keep=False,
                                row_height=next(r["target"] for r in pending["rows"] if r["row"] == card["excel_row"]
                                                and r["sheet"] == card["sheet_name"]),
                                fix_images=copy.deepcopy(card.get("fix_images", {})),
                                review_reason="全自动批处理：已修复并保存")
            else:
                card["message"] = "已还原自动修改，等待人工处理" if card["color"] == "yellow" else "已还原，待提交忽略"
        job["pending_adjustment"] = None
        wf._persist_state()
        wf._sync_problem_counts(pending["path"], "自动处理已保存" if pending["action"] == "fix" else "审核还原已保存")


def _refresh(excel: Any, path: str, sheets: list[str] | None = None) -> None:
    with wf.STATE_LOCK:
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
    wf._scan_one_workbook(
        excel, Path(path), {"mode": "sheets" if sheets else "full", "sheets": sheets, "scanned": bool(previous)},
        previous, job=None, resumed=True, rescan_all=False, delete_backups=False,
        used_pdf_paths={Path(p["pdf_path"]) for p in wf.STATE["pages"].values()},
        read_only=True,
    )
    with wf.STATE_LOCK:
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
                _card_for(group)
            elif existing and not existing.get("dismissed") and not existing.get("manual") and not existing.get("marked"):
                existing.update(color="blue", verified=True, after_kind="fixed", message="最新导出已通过检查，等待审核")
                after = next(((it.get("fix_images") or {}).get("after") for it in group
                              if (it.get("fix_images") or {}).get("after")), None)
                if after:
                    existing.setdefault("fix_images", {})["after"] = after
        wf._persist_state()


def _group_cells(items: list[dict]) -> list[list[dict]]:
    cells: dict[tuple, list[dict]] = {}
    for item in items:
        cells.setdefault(_cell(item), []).append(item)
    return list(cells.values())


def _prepare_row(
    session: Any, path: str, items: list[dict], *,
    on_attempt: Callable[[int, float, dict], None] | None = None,
) -> tuple[float, dict | None, dict]:
    """只在副本上做图；逐次交出截图，即使随后导出失败也能保留最后已生成的效果。"""
    sheet = session.copy_of(path).Sheets(items[0]["sheet_name"])
    autofix.validate_print_targets(sheet, items)
    row = int(items[0]["excel_row"])
    base = round(float(sheet.Rows(row).RowHeight), 1)
    cells = snapshot_cells(extract_sheet_snapshot(sheet))
    probe = CellProbe(sheet)
    groups = _group_cells(items)
    before: dict[str, str] = {}
    chosen = None
    extra = 0.0
    success = False
    pdf = session.temp_pdf()
    try:
        for attempt in range(autofix.MAX_ATTEMPTS + 1):
            height = min(autofix.MAX_ROW_HEIGHT, round(base + extra, 1))
            if attempt and height <= last_height:
                break
            sheet.Rows(row).RowHeight = height
            wf._export_sheet_pdf(sheet, pdf)
            # 每次行高改变后重建 probe，避免沿用 Range.Text 缓存。
            probe = CellProbe(sheet)
            ok, _, need = autofix._evaluate(pdf, items, cells, probe.text, probe.printed)
            shots = {_key(*_cell(group[0])): autofix._evaluate(pdf, group, cells, probe.text, probe.printed)[1]
                     for group in groups}
            if on_attempt is not None:
                on_attempt(attempt, height, shots)
            if attempt == 0:
                before = shots
            if ok:
                chosen = {"height": height, "shots": shots}
            if (attempt == 0 and ok) or success:
                break
            success = ok
            last_height = height
            extra = max(extra + autofix.STEP_PT,
                        math.ceil(round((extra + need) / autofix.STEP_PT, 3)) * autofix.STEP_PT)
    finally:
        sheet.Rows(row).RowHeight = base
        pdf.unlink(missing_ok=True)
    return base, chosen, before


def _repair_file(excel: Any, session: Any, path: str) -> None:
    job = wf.STATE["job"]
    attempted = job.setdefault("attempted_rows", [])
    while True:
        if job.get("pause_requested"):
            return
        with wf.STATE_LOCK:
            pending = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                       and it["status"] not in wf.RESOLVED_STATUSES]
            manual_rows = {(c["sheet_name"], c["excel_row"]) for c in _data()["cards"].values()
                           if c["wb_path"] == path and (c.get("manual") or c.get("marked")) and not c.get("dismissed")}
            manual_rows.update((r["sheet"], r["row"]) for r in _data()["rows"].values()
                               if r["path"] == path and r["saved"]
                               and any(not c.get("dismissed") and c["wb_path"] == path
                                       and c["sheet_name"] == r["sheet"] and c["excel_row"] == r["row"]
                                       for c in _data()["cards"].values()))
            candidates = [it for it in pending if _key(path, it["sheet_name"], it["excel_row"]) not in attempted
                          and (it["sheet_name"], it["excel_row"]) not in manual_rows]
            if not candidates:
                # 之前修过的格子因后续分页变化又出现问题，也交给人工，不能循环自动加高。
                for group in _group_cells(pending):
                    card = _card_for(group)
                    if not card.get("dismissed") and not card.get("marked") and not card.get("manual"):
                        card.update(color="yellow", verified=False, message="当前导出仍有问题，需人工处理")
                wf._persist_state()
                return
            first = min(candidates, key=lambda it: (it["sheet_name"], int(it["excel_row"]), it["seq"]))
            items = [it for it in pending if it["sheet_name"] == first["sheet_name"]
                     and it["excel_row"] == first["excel_row"]]
            cards = [_card_for(group) for group in _group_cells(items)]
            rid = _key(path, first["sheet_name"], first["excel_row"])
            wf._persist_state()
        _check_file(path)
        version = wf._file_version(Path(path))

        def remember_attempt(index: int, height: float, shots: dict) -> None:
            with wf.STATE_LOCK:
                for card in cards:
                    if index == 0:
                        _save_shot(card, "before", shots.get(card["id"], ""))
                    # 尚未产生调整图就出错时，原样也是最后一次实际生成的预览。
                    # 只给卡片留参考图，不写入问题项的修复报告，也不写原 Excel。
                    fix_images.forget(card, "after")
                    _save_shot(card, "after", shots.get(card["id"], ""))
                    card.update(after_kind="attempt", attempt_height=height)
                wf._persist_state()

        try:
            reason = autofix.review_only_reason(items)
            if reason:
                for card in cards:
                    card["manual"] = True  # 同行共用行高；该行也不能因另一格而自动加高。
                raise ValueError(reason)
            if any(it.get("location_confirmed") is False or not it.get("cell_addr") for it in items):
                raise ValueError("单元格定位未确认，请人工核对")
            base, chosen, before = _prepare_row(session, path, items, on_attempt=remember_attempt)
            with wf.STATE_LOCK:
                for card in cards:
                    _save_shot(card, "before", before.get(card["id"], ""))
                if chosen is None:
                    raise ValueError("自动试修未通过，原文件行高未修改；最后一次尝试未采用，请人工处理或标绿忽略")
                for card in cards:
                    _save_shot(card, "after", chosen["shots"].get(card["id"], ""))
                _data()["rows"][rid] = {"id": rid, "path": path, "sheet": first["sheet_name"],
                                        "row": int(first["excel_row"]), "base": base,
                                        "applied": chosen["height"], "saved": False}
                job["pending_adjustment"] = {
                    "action": "fix", "path": path, "version": version, "cards": [c["id"] for c in cards],
                    "rows": [{"id": rid, "sheet": first["sheet_name"], "row": int(first["excel_row"]),
                              "expected": base, "target": chosen["height"]}],
                }
                wf._persist_state()
            _finish_pending(excel)
        except Exception as err:
            # 有写入日志时不能吞掉异常：可能原文件已经保存，必须按日志续跑。
            if job.get("pending_adjustment"):
                raise
            with wf.STATE_LOCK:
                for card in cards:
                    card.update(color="yellow", verified=False, manual=True, message=str(err))
                wf.log_msg(f"⚠️ {first['wb_name']} / {first['sheet_name']} 行{first['excel_row']}：{err}")
        with wf.STATE_LOCK:
            if rid not in attempted:
                attempted.append(rid)
            wf._persist_state()
        session.drop_copy(path)
        # 包括成功写入后原来没被检出的新问题，始终以最新整表 PDF 为准。
        _refresh(excel, path, [first["sheet_name"]])


def _commit_review(excel: Any) -> None:
    job = wf.STATE["job"]
    data = _data()
    if "commit_card_ids" not in job["options"]:
        cards = [data["cards"][cid] for cid in job["options"]["card_ids"]]
        selected_rows = {_key(c["wb_path"], c["sheet_name"], c["excel_row"]) for c in cards}
        job["options"]["commit_card_ids"] = [c["id"] for c in data["cards"].values()
            if not c.get("dismissed") and _key(c["wb_path"], c["sheet_name"], c["excel_row"]) in selected_rows]
    # 提交计划固定下来；即使保存还原后、复检前退出，已收起的绿/蓝卡也仍会完成该文件的复检。
    job["options"]["card_ids"] = list(job["options"]["commit_card_ids"])
    cards = [data["cards"][cid] for cid in job["options"]["commit_card_ids"]]
    by_row: dict[str, list[dict]] = {}
    for card in cards:
        by_row.setdefault(_key(card["wb_path"], card["sheet_name"], card["excel_row"]), []).append(card)
    # 只在仍有自动修改需要还原时检查同行冲突；已经忽略的格子不因后来复检重新入队。
    for rid, peers in by_row.items():
        row = data["rows"].get(rid)
        if row and row["saved"] and len({c["color"] for c in peers}) > 1:
            for card in peers:
                card.update(color="yellow", conflict=True, message="同行存在不同处理意见，整行还原后交给人工")
    wf._persist_state()
    for path in dict.fromkeys(c["wb_path"] for c in cards):
        if path in job["completed"]:
            continue
        rows = []
        reverting = []
        for rid, peers in by_row.items():
            row = data["rows"].get(rid)
            if not row or row["path"] != path or not row["saved"] or all(c["color"] == "blue" for c in peers):
                continue
            rows.append({"id": rid, "sheet": row["sheet"], "row": row["row"],
                         "expected": row["applied"], "target": row["base"]})
            reverting.extend(c["id"] for c in peers)
        if rows:
            job["options"].setdefault("refresh_sheets", {})[path] = sorted({r["sheet"] for r in rows})
            job["pending_adjustment"] = {"action": "restore", "path": path,
                                         "version": wf._file_version(Path(path)), "rows": rows, "cards": reverting}
            wf._persist_state()
            _finish_pending(excel)
        with wf.STATE_LOCK:
            for card in [c for c in cards if c["wb_path"] == path]:
                card["finalized"] = True
                if card["color"] == "green":
                    rule = {"cell": list(_cell(card)), "signatures": card["signatures"]}
                    if rule not in data["ignored"]:
                        data["ignored"].append(rule)
                card["dismissed"] = card["color"] != "yellow"
                card["manual"] = card["color"] == "yellow"
                if card["manual"]:
                    card.update(verified=False, message=("同行存在不同处理意见，整行已还原；" if card.get("conflict") else "")
                                + "自动修改已还原，请在 Excel 中修改并保存后复检")
                for item in _items_at(card):
                    item.update(batch_managed=True, triage_keep=False)
                    if card["color"] == "green":
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
                        card.update(color="yellow", verified=False, dismissed=False, manual=True,
                                    finalized=False, message="还原后重新检测仍有问题，请提交黄色标记后人工处理")
                wf._persist_state()
        with wf.STATE_LOCK:
            for rid, peers in by_row.items():
                row = data["rows"].get(rid)
                if row and row["path"] == path and all(c.get("dismissed") and c["color"] == "blue" for c in peers):
                    row["saved"] = False  # 已正式接受的历史修改不再属于可还原的待审核修改。
            job["completed"].append(path)
            wf._persist_state()


def _recheck_file(excel: Any, path: str, cards: list[dict]) -> None:
    _check_file(path)
    sheets = sorted({c["sheet_name"] for c in cards})
    _refresh(excel, path, sheets)
    with wf.STATE_LOCK:
        unresolved = [it for it in wf.STATE["items"].values() if it["wb_path"] == path
                      and it["sheet_name"] in sheets and it["status"] not in wf.RESOLVED_STATUSES]
        for card in cards:
            still = [it for it in unresolved if _cell(it) == _cell(card)]
            # 定位不确定的同表问题不能据此断言原格已修复。
            uncertain = any(it.get("location_confirmed") is False and it["sheet_name"] == card["sheet_name"]
                            for it in unresolved)
            card.update(color="yellow" if still or uncertain else "blue", verified=not (still or uncertain),
                        finalized=False, message="复检仍有问题" if still or uncertain else "人工修改复检通过，等待确认")
            if not still and not uncertain:
                prior = _items_at(card)
                after = next(((it.get("fix_images") or {}).get("after") for it in prior
                              if (it.get("fix_images") or {}).get("after")), None)
                if after:
                    card.setdefault("fix_images", {})["after"] = after
                else:
                    fix_images.forget(card, "after")
                card["after_kind"] = "fixed"
        for group in _group_cells(unresolved):
            card = _card_for(group)
            if not card.get("dismissed"):
                card.update(color="yellow", verified=False, manual=True, finalized=False,
                            message="复检仍有问题或发现新问题，请人工处理")
        wf._persist_state()
        wf._sync_problem_counts(path, "人工复检完成")


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
    session = autofix._ExcelSession()
    wf.pythoncom.CoInitialize()
    try:
        excel = session.app() if job.get("pending_adjustment") else None
        if excel is not None:
            _finish_pending(excel)
        if kind == "batch_scan":
            if "workbooks" not in job:
                folder, paths = wf.list_workbooks(options["folder"])
                with wf.STATE_LOCK:
                    if str(folder) != wf.STATE["folder"]:
                        wf.STATE.update(items={}, pages={}, checked_files={},
                                        batch_review={"cards": {}, "rows": {}, "ignored": []})
                    wf.STATE.update(folder=str(folder), pdf_folder=str(folder))
                    job["workbooks"] = [str(p) for p in paths]
                    wf._persist_state()
                init_execution_batch(task_name="全自动扫描与修复", folder=folder, workbooks=paths, pdf_folder=folder)
            expected = job["workbooks"]
            for path in expected:
                if path in job["completed"]:
                    continue
                if job.get("pause_requested"):
                    break
                try:
                    _check_file(path)
                    excel = session.app()
                    _refresh(excel, path)
                    _repair_file(excel, session, path)
                    if job.get("pause_requested"):
                        break
                    with wf.STATE_LOCK:
                        job["completed"].append(path)
                        wf.STATE["scan_progress"] = {"total": len(expected), "completed": len(job["completed"]), "status": "running"}
                        wf._persist_state()
                    wf._sync_problem_counts(path, "全自动处理完成，等待结果审核")
                except Exception as err:
                    wf.log_msg(f"❌ {Path(path).name} 未完成，可续跑：{err}")
                    update_file_progress(workbook_rel_name(Path(path), wf.STATE["folder"]),
                                         status="error", error=str(err), current_action="文件未完成，可续跑")
                    if job.get("pending_adjustment"):
                        break  # 写入日志未结算前不能处理另一文件
                finally:
                    session.drop_copy(path)
        elif kind == "batch_commit":
            if options.get("refresh_sheets") or any(c["color"] != "blue" and not c.get("dismissed")
                                                   for c in _data()["cards"].values()):
                excel = session.app()
            _commit_review(excel)
            expected = list(dict.fromkeys(_data()["cards"][cid]["wb_path"] for cid in options["card_ids"]))
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
                unsaved = close_saved_workbooks([Path(path) for path in expected])
                for path in expected:
                    if path in job["completed"]:
                        continue
                    if Path(path) in unsaved:
                        wf.log_msg(f"⚠️ {Path(path).name}：{unsaved[Path(path)]}，请保存后继续")
                        continue
                    try:
                        _recheck_file(excel, path, [c for c in cards if c["wb_path"] == path])
                        job["completed"].append(path)
                        wf._persist_state()
                    except Exception as err:
                        wf.log_msg(f"❌ {Path(path).name} 复检失败，可继续任务：{err}")
        else:
            raise ValueError("未知批处理任务")
        finish_execution_batch("处理完成，请审核结果" if set(expected).issubset(job["completed"]) else "部分文件尚未完成，可继续任务")
        return expected
    finally:
        session.close()
        wf.pythoncom.CoUninitialize()
