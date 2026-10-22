"""自动处理切字：在工作簿副本上一次加一点行高，每次导出 PDF 截图并判断是否已显示全，
做出最多 9 张图（原样 + 8 次加高）给用户挑；用户点哪张，就把那张的行高写进原文件。

- 每次加 3pt；隐藏截断还缺整行没打印时，先按缺的行数一次补上，剩下不到一行再按 3pt 加

- 同一行的几处问题合成一组处理（加的是整行高度）
- 预览只在副本上做，不锁、不改原文件；用户挑选时后台继续准备后面的组
- 显示全之后再多加一次就停，推荐多加的那张（留一点余量）
- 默认全自动：准备好就按推荐写进原文件并保存，再列出来等用户批量确认；
  确认前可以改用别的候选图（重新写入），或点「不对」改回原行高、转稍后人工复检
- 全自动时加高 8 次仍不行（或出错）的行，直接转稍后人工复检
- 扫描期间处理已经提交的文件；复检、人工定位等任务时暂停
- 工作台打开时可为已筛完的行做图；尚未筛完的行不入队
- 「稍后人工复检」的行不再做图；后台在没有任务时给这些单元格标黄、给文件名加 ★
"""
import base64
import hashlib
import math
import shutil
import tempfile
import threading
import time
from pathlib import Path
from typing import Any

from core.common import STALE_LOCK_HINT, XL_SHEET_VISIBLE, as_bbox, col_index_to_letters
from core.excel_ops import (
    CellProbe, extract_sheet_snapshot, normalize_workbook_view, quit_background_excel, snapshot_cells,
    start_background_excel,
)
from core.hidden_text import collect_text_pieces, find_hidden_cells, locate_text, normalize_text
from core.pdf_detector import (
    RegionTrackingError, crossing_horizontal_line, extract_page_lines, fitz,
    is_region_still_clipped, render_crop_base64,
)
from core.workflow import (
    RESOLVED_STATUSES,
    STATE,
    STATE_LOCK,
    apply_skip_decisions,
    file_available,
    _export_sheet_pdf,
    _file_version,
    _persist_state,
    _sync_problem_counts,
    _check_decision_item,
    log_msg,
    mark_review_file,
    queue_later_review,
    restore_review_file,
    review_entry_needs_mark,
    review_entry_needs_restore,
    pythoncom,
)

STEP_PT = 3.0
MAX_ATTEMPTS = 8  # 加上原样共 9 张，正好九宫格
MAX_ROW_HEIGHT = 409.0  # Excel 行高上限
MANUAL_STATUSES = {"manual_review"}
WRITE_STATUSES = {"applying", "reverting"}  # 等后台写进原文件：按选中的图 / 改回原行高
LOCKED_RETRY_SECONDS = 5.0
MAX_ERROR_RETRIES = 2  # 偶然出错（Excel 正忙、临时拒绝调用等）先重试这么多次，再转人工
ACTIVE_GRACE_SECONDS = 8.0  # 页面每 1.5 秒报一次「正开着 ② 自动处理」，超过这么久没报就停下

_LOCK = threading.RLock()
_WAKE = threading.Event()
_IDLE = threading.Event()  # 后台没有开着的 Excel
_IDLE.set()
_WRITES_IDLE = threading.Event()  # 人工打开只需等正在进行的原文件写入，不等副本 Excel 退出
_WRITES_IDLE.set()
AUTOFIX: dict[str, Any] = {"auto": True, "groups": {}, "hold": 0, "active_until": 0.0}
_worker: threading.Thread | None = None


class _Paused(Exception):
    """后台有别的任务、或这组已不需要处理，中止当前预览，稍后重来"""


def _group_id(wb_path: str, sheet_name: str, row: int) -> str:
    return hashlib.sha1(f"{wb_path}\0{sheet_name}\0{row}".encode()).hexdigest()[:16]


def _busy() -> bool:
    with _LOCK:
        if AUTOFIX["hold"]:
            return True
    with STATE_LOCK:
        return bool(STATE["busy"] and (STATE["job"] or {}).get("kind") != "scan")


def check_row_not_writing(item: dict[str, Any]) -> None:
    """调用方已持有 STATE_LOCK，与 choose 的 STATE → 分组锁顺序一致。"""
    with _LOCK:
        group = AUTOFIX["groups"].get(_group_id(item["wb_path"], item["sheet_name"], int(item["excel_row"])))
        if group and group["status"] in {"applying", "reverting"}:
            raise ValueError("这一行已开始写入，不能撤销或重新筛选")


def mark_active() -> None:
    """工作台处于前台：可以为已筛选的行预先做图。"""
    with _LOCK:
        was_active = _preview_allowed()
        AUTOFIX["active_until"] = time.monotonic() + ACTIVE_GRACE_SECONDS
    if not was_active:
        _WAKE.set()


def _preview_allowed() -> bool:
    with _LOCK:
        return time.monotonic() < AUTOFIX["active_until"]


def pause_previews() -> None:
    with _LOCK:
        AUTOFIX["active_until"] = 0.0
    _WAKE.set()


def wait_for_writes(timeout: float = 120.0) -> None:
    if not _WRITES_IDLE.wait(timeout):
        raise ValueError("后台仍在保存 Excel，请保存结束后再打开人工处理")


def release_excel(timeout: float = 120.0) -> None:
    """扫描、复检等任务前让后台关闭 Excel；人工定位只需 wait_for_writes。
    任务排上队后调用 resume_after_release。"""
    with _LOCK:
        AUTOFIX["hold"] += 1
    _WAKE.set()
    if _worker is not None and _worker.is_alive():
        _IDLE.wait(timeout)


def resume_after_release() -> None:
    with _LOCK:
        AUTOFIX["hold"] = max(0, AUTOFIX["hold"] - 1)
    _WAKE.set()


def sync_groups() -> None:
    """按当前未解决的问题项整理分组；组内问题项变了（重新扫描等）就重新准备"""
    with STATE_LOCK:
        # 读取问题状态和更新分组必须处于同一把状态锁内，
        # 避免较早的同步在跳过保存后又按旧快照把该行加回队列。
        _sync_groups_locked()


def _sync_groups_locked() -> None:
    rows: dict[tuple[str, str, int], list[dict[str, Any]]] = {}
    for item in STATE["items"].values():
        if item["status"] not in RESOLVED_STATUSES and file_available(item["wb_path"]):
            rows.setdefault((item["wb_path"], item["sheet_name"], int(item["excel_row"])), []).append(item)
    eligible = {item["id"] for peers in rows.values() if all(it.get("triage_keep") for it in peers) for item in peers}
    items = [
        (item["id"], item["wb_path"], item["wb_name"], item["sheet_name"], int(item["excel_row"]),
         int(item.get("seq") or 0), item["status"],
         item.get("manual_waiting", item["status"] == "manual_review"),
         item.get("manual_recheck_token"), item.get("image", ""), item.get("row_height"),
         item.get("autofix_revision"))
        for item in STATE["items"].values() if item["id"] in eligible
    ]
    wanted: dict[str, dict[str, Any]] = {}
    for item_id, wb_path, wb_name, sheet_name, row, seq, status, waiting, token, image, height, revision in items:
        gid = _group_id(wb_path, sheet_name, row)
        entry = wanted.setdefault(gid, {
            "wb_path": wb_path, "wb_name": wb_name, "sheet_name": sheet_name, "row": row,
            "item_ids": [], "seq": seq, "manual": False, "manual_waiting": False,
            "recheck_token": None, "recheck_image": "", "recheck_height": None,
            "revision": revision,
        })
        entry["item_ids"].append(item_id)
        entry["seq"] = min(entry["seq"], seq)
        entry["manual"] = entry["manual"] or status in MANUAL_STATUSES
        entry["manual_waiting"] = entry["manual_waiting"] or waiting
        if token and not entry["recheck_token"]:
            entry.update(recheck_token=token, recheck_image=image, recheck_height=height)
    with _LOCK:
        groups = AUTOFIX["groups"]
        for gid in list(groups):
            # 已写入的组问题项都已解决，保留记录供页面显示；确认前改图、退回时问题项也已解决
            if gid not in wanted and not _kept_after_apply(groups[gid]):
                del groups[gid]
        _restore_unconfirmed_locked(groups, wanted)
        for gid, entry in wanted.items():
            group = groups.get(gid)
            # 写入过的行又检出了问题（重新扫描）也重新准备；
            # 转人工后复检仍没显示全的，按用户改过的行高重新准备，旧图片已经不对了
            rechecked = group is not None and group["manual"] and not entry["manual"]
            if (group is None or group["status"] == "applied" or rechecked
                    or group.get("revision") != entry["revision"]
                    or sorted(group["item_ids"]) != sorted(entry["item_ids"])):
                group = groups[gid] = {
                    "id": gid, **{k: v for k, v in entry.items() if not k.startswith("recheck_")},
                    "status": "queued", "attempts": [], "recommended": None,
                    "chosen": None, "base_height": None, "confirmed": False,
                    "message": "复检后仍未显示全，已按你改过的行高重新准备" if rechecked else "",
                }
            group["manual"] = entry["manual"]
            group["manual_waiting"] = entry["manual_waiting"]
            group["seq"] = entry["seq"]
            if entry["manual_waiting"]:
                group.pop("recheck_token", None)
            if entry["recheck_token"] and group.get("recheck_token") != entry["recheck_token"]:
                group.update(
                    status="failed", recheck_token=entry["recheck_token"],
                    base_height=entry["recheck_height"], recommended=None, chosen=None,
                    attempts=[{"height": entry["recheck_height"], "ok": False, "image": entry["recheck_image"]}],
                    message="本次复检仍有问题，请继续修改、跳过本行或先看下一行",
                )


def _kept_after_apply(group: dict[str, Any]) -> bool:
    status = group["status"]
    return (status == "applied" or status == "reverting" or (status == "applying" and group.get("reapply"))
            or (status == "reverted" and not group.get("dismissed")))


def _restore_unconfirmed_locked(groups: dict[str, Any], wanted: dict[str, Any]) -> None:
    """重启后图片没了：已自动写入、还没确认的行仍列出来，可以确认或退回（改回原行高）"""
    rows: dict[str, list[dict[str, Any]]] = {}
    for item in STATE["items"].values():
        if item["status"] == "fixed" and item.get("autofix_unconfirmed"):
            gid = _group_id(item["wb_path"], item["sheet_name"], int(item["excel_row"]))
            if gid not in groups and gid not in wanted:
                rows.setdefault(gid, []).append(item)
    for gid, items in rows.items():
        first = items[0]
        groups[gid] = {
            "id": gid, "wb_path": first["wb_path"], "wb_name": first["wb_name"],
            "sheet_name": first["sheet_name"], "row": int(first["excel_row"]),
            "item_ids": [it["id"] for it in items], "seq": min(int(it.get("seq") or 0) for it in items),
            "manual": False, "manual_waiting": False, "revision": first.get("autofix_revision"),
            "status": "applied", "attempts": [], "recommended": None, "chosen": None,
            "base_height": first.get("autofix_base_height"), "applied_height": first.get("row_height"),
            "confirmed": False, "restored": True,
            "message": "程序重启过，候选图已清除；行高已写入，可直接确认或点「不对」改回原行高",
        }


def snapshot() -> dict[str, Any]:
    """页面轮询用：不含图片（图片按需单独取）"""
    ensure_worker()
    sync_groups()
    with _LOCK:
        groups = sorted(AUTOFIX["groups"].values(), key=lambda g: g["seq"])
        return {
            "auto": AUTOFIX["auto"],
            "step_pt": STEP_PT,
            "groups": [
                {
                    **{k: v for k, v in g.items() if k != "attempts"},
                    "attempts": [{"height": a["height"], "ok": a["ok"]} for a in g["attempts"]],
                }
                for g in groups
            ],
        }


def attempt_image(gid: str, index: int) -> bytes | None:
    with _LOCK:
        group = AUTOFIX["groups"].get(gid)
        if group is None or not 0 <= index < len(group["attempts"]):
            return None
        image = group["attempts"][index]["image"]
    if not image:
        return None
    return base64.b64decode(image.split(",", 1)[1])


def choose(gid: str, index: int) -> None:
    with STATE_LOCK:
        if STATE["busy"] and (STATE["job"] or {}).get("kind") != "scan":
            raise ValueError("任务正在执行，请结束后再选图")
        _sync_groups_locked()
        _choose_locked(gid, index)
    _WAKE.set()


def _choose_locked(gid: str, index: int) -> None:
    with _LOCK:
        group = AUTOFIX["groups"].get(gid)
        if group is None:
            raise ValueError("这一组已重新扫描或已处理，请刷新页面")
        if group["status"] not in {"ready", "failed"}:
            raise ValueError("这一组还没准备好或已在写入")
        if not 1 <= index < len(group["attempts"]) or not group["attempts"][index]["ok"]:
            raise ValueError("只能选择已显示全且无压线的调整结果")
        group["chosen"] = index
        group["status"] = "applying"
        group["confirmed"] = True  # 用户亲手选的，不用再确认
        group["message"] = "正在写入 Excel..."


def skip_group(gid: str, item_ids: list[str]) -> dict[str, Any]:
    """把当前行全部问题持久化为无需修改；与选图／全自动排入写入队列互斥。"""
    with STATE_LOCK:
        if STATE["busy"] and (STATE["job"] or {}).get("kind") != "scan":
            raise ValueError("任务正在执行，请结束后再标记跳过")
        sync_groups()
        with _LOCK:
            group = AUTOFIX["groups"].get(gid)
            if group is None or sorted(group["item_ids"]) != sorted(item_ids):
                raise ValueError("这一行的问题已更新，请刷新后重新判断")
            if group["status"] in {"applying", "applied"}:
                raise ValueError("这一行已开始写入或已保存，不能再跳过")
            result = apply_skip_decisions({iid: True for iid in group["item_ids"]})
            sync_groups()
    _WAKE.set()
    return result


def defer_group(gid: str, item_ids: list[str]) -> None:
    """把这一行挂到稍后人工复检：不再做图，等最后统一打开 Excel 修改。"""
    with STATE_LOCK:
        if STATE["busy"] and (STATE["job"] or {}).get("kind") != "scan":
            raise ValueError("任务正在执行，请结束后再操作")
        sync_groups()
        with _LOCK:
            group = AUTOFIX["groups"].get(gid)
            if group is None or sorted(group["item_ids"]) != sorted(item_ids):
                raise ValueError("这一行的问题已更新，请刷新后重新判断")
            if group["status"] in {"applying", "applied"}:
                raise ValueError("这一行已开始写入或已保存，不能再挂起")
            for iid in group["item_ids"]:
                _check_decision_item(STATE["items"][iid])
            queue_later_review(list(group["item_ids"]))
            sync_groups()
    _WAKE.set()


def _check_not_busy() -> None:
    if STATE["busy"] and (STATE["job"] or {}).get("kind") != "scan":
        raise ValueError("任务正在执行，请结束后再操作")


def _unconfirmed(gid: str) -> dict[str, Any]:
    group = AUTOFIX["groups"].get(gid)
    if group is None or group["confirmed"]:
        raise ValueError("这一行已确认或已重新扫描，请刷新页面")
    if group["status"] != "applied":
        raise ValueError("这一行正在写入，请稍候再操作")
    return group


def confirm_groups(gids: list[str]) -> int:
    """批量确认已自动写入的行；已退回的行从确认列表里收起"""
    with STATE_LOCK:
        _check_not_busy()
        _sync_groups_locked()
        confirmed: list[str] = []
        with _LOCK:
            for gid in gids:
                group = AUTOFIX["groups"].get(gid)
                if group is None or group["confirmed"]:
                    continue
                if group["status"] == "applied":
                    group["confirmed"] = True
                    confirmed.extend(group["item_ids"])
                elif group["status"] == "reverted":
                    group["dismissed"] = True
        for iid in confirmed:
            if iid in STATE["items"]:
                STATE["items"][iid].pop("autofix_unconfirmed", None)
        if confirmed:
            _persist_state()
            log_msg(f"✅ 已确认自动处理结果 {len(confirmed)} 处")
    return len(confirmed)


def switch_group(gid: str, index: int) -> None:
    """确认前改用另一张候选图：按它的行高重新写入原文件"""
    with STATE_LOCK:
        _check_not_busy()
        _sync_groups_locked()
        with _LOCK:
            group = _unconfirmed(gid)
            if not 1 <= index < len(group["attempts"]) or not group["attempts"][index]["ok"]:
                raise ValueError("只能选择已显示全且无压线的调整结果")
            if index == group["chosen"]:
                return
            group.update(chosen=index, status="applying", reapply=True, message="正在改用这张，写入 Excel...")
    _WAKE.set()


def reject_group(gid: str) -> None:
    """自动处理的结果不对：改回原行高，然后转稍后人工复检"""
    with STATE_LOCK:
        _check_not_busy()
        _sync_groups_locked()
        with _LOCK:
            group = _unconfirmed(gid)
            if group["base_height"] is None:
                raise ValueError("不知道原来的行高，无法自动改回，请在 Excel 里手动修改")
            group.update(status="reverting", message="正在改回原行高...")
    _WAKE.set()


def _auto_defer(gid: str) -> None:
    """全自动时处理不了的行直接转稍后人工复检；任务进行中等情况先留在页面上"""
    with STATE_LOCK:
        _sync_groups_locked()
        with _LOCK:
            group = AUTOFIX["groups"].get(gid)
            if (group is None or not AUTOFIX["auto"] or group["status"] not in {"failed", "error"}
                    or group["manual"] or group.get("recheck_token")):
                return
            item_ids = list(group["item_ids"])
            label = f"{group['wb_name']} / {group['sheet_name']} 行{group['row']}"
            message = group["message"]
        try:
            for iid in item_ids:
                _check_decision_item(STATE["items"][iid])
            queue_later_review(item_ids)
        except Exception:
            return
        _sync_groups_locked()
        log_msg(f"★ [{label}] 自动处理不了，已转稍后人工复检：{message}")
    _WAKE.set()


def rename_groups(old_path: str, new_path: str) -> None:
    """文件改名（加 / 去 ★）后分组跟着改，已点选等待写入的行不丢。调用方持有 STATE_LOCK。"""
    with _LOCK:
        groups = AUTOFIX["groups"]
        for gid in [gid for gid, g in groups.items() if g["wb_path"] == old_path]:
            group = groups.pop(gid)
            group["wb_path"] = new_path
            group["id"] = _group_id(new_path, group["sheet_name"], group["row"])
            groups[group["id"]] = group


def set_auto(on: bool) -> None:
    with STATE_LOCK, _LOCK:
        _sync_groups_locked()
        AUTOFIX["auto"] = on
        if on:
            for group in AUTOFIX["groups"].values():
                if group["status"] == "ready" and not group["manual"] and group["recommended"] is not None:
                    group["chosen"] = group["recommended"]
                    group["status"] = "applying"
                    group["confirmed"] = False
                    group["message"] = "全自动：正在写入 Excel..."
    _WAKE.set()


def ensure_worker() -> None:
    global _worker
    with _LOCK:
        if _worker is not None and _worker.is_alive():
            return
        _worker = threading.Thread(target=_run, name="autofix", daemon=True)
        _worker.start()


class _ExcelSession:
    """后台专用的隐藏 Excel；预览用工作簿副本，写入时才短暂打开原文件"""

    def __init__(self) -> None:
        self.excel: Any = None
        self.temp_dir: Path | None = None
        self.copies: dict[str, tuple[dict[str, int], Any]] = {}
        self.copy_count = 0  # 副本文件名编号只增不减，不同子文件夹里的同名文件也不会撞名

    def app(self) -> Any:
        if self.excel is None:
            _IDLE.clear()
            self.excel = start_background_excel()
            self.temp_dir = Path(tempfile.mkdtemp(prefix="clip_autofix_"))
        return self.excel

    def copy_of(self, wb_path: str) -> Any:
        """原文件的副本（原文件改过就重新复制）"""
        version = _file_version(Path(wb_path))
        cached = self.copies.get(wb_path)
        if cached is not None and cached[0] == version:
            return cached[1]
        if cached is not None:
            cached[1].Close(SaveChanges=False)
        app = self.app()
        assert self.temp_dir is not None
        self.copy_count += 1
        target = self.temp_dir / f"{self.copy_count}_{Path(wb_path).name}"
        shutil.copy2(wb_path, target)
        workbook = app.Workbooks.Open(str(target), UpdateLinks=0)
        self.copies[wb_path] = (version, workbook)
        return workbook

    def drop_copy(self, wb_path: str) -> None:
        cached = self.copies.pop(wb_path, None)
        if cached is not None:
            try:
                cached[1].Close(SaveChanges=False)
            except Exception:
                pass

    def temp_pdf(self) -> Path:
        self.app()
        assert self.temp_dir is not None
        return self.temp_dir / "attempt.pdf"

    def close(self) -> None:
        for _, workbook in self.copies.values():
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        self.copies = {}
        if self.excel is not None:
            try:
                quit_background_excel(self.excel)
            except Exception:
                pass
            self.excel = None
        if self.temp_dir is not None:
            shutil.rmtree(self.temp_dir, ignore_errors=True)
            self.temp_dir = None
        _IDLE.set()


def _run() -> None:
    pythoncom.CoInitialize()
    session = _ExcelSession()
    blocked: dict[str, float] = {}  # 原文件被占用，暂缓写入到这个时间
    try:
        while True:
            if _busy():
                session.close()  # 让出 Excel，避免和扫描、人工定位抢同一个文件
                _WAKE.wait(2)
                _WAKE.clear()
                continue
            sync_groups()
            with _LOCK:
                failed = [gid for gid, g in AUTOFIX["groups"].items() if g["status"] in {"failed", "error"}]
            for gid in failed:
                _auto_defer(gid)  # 任务进行中转不了的，下一轮再试
            task = _next_task(blocked)
            if task is None:
                session.close()
                _WAKE.wait(3)
                _WAKE.clear()
                continue
            kind, payload = task
            try:
                if kind == "apply":
                    with STATE_LOCK:
                        if (STATE["busy"] and (STATE["job"] or {}).get("kind") != "scan") or not file_available(payload[0]):
                            continue
                        _WRITES_IDLE.clear()
                    try:
                        if not _apply(session, *payload):
                            blocked[payload[0]] = time.monotonic() + LOCKED_RETRY_SECONDS
                    finally:
                        _WRITES_IDLE.set()
                elif kind == "mark":
                    with STATE_LOCK:
                        if STATE["busy"]:
                            continue
                        _WRITES_IDLE.clear()
                    try:
                        session.drop_copy(payload)
                        if not mark_review_file(session.app(), payload):
                            blocked[payload] = time.monotonic() + LOCKED_RETRY_SECONDS
                    except Exception as err:
                        with STATE_LOCK:
                            entry = STATE["review_queue"].get(payload)
                            if entry is not None:
                                # 不再自动重试；点「一次打开全部」时会再试一次
                                entry["mark_failed"] = True
                                entry["error"] = f"标黄失败: {err}"
                        raise
                    finally:
                        _WRITES_IDLE.set()
                elif kind == "restore":
                    with STATE_LOCK:
                        if STATE["busy"]:
                            continue
                        _WRITES_IDLE.clear()
                    try:
                        session.drop_copy(payload)
                        if restore_review_file(session.app(), payload) == "locked":
                            blocked[payload] = time.monotonic() + LOCKED_RETRY_SECONDS
                    except Exception as err:
                        with STATE_LOCK:
                            entry = STATE["review_queue"].get(payload)
                            if entry is not None:
                                # 不再自动重试；点「一次打开全部」或「统一复检」时会再处理
                                entry["restore_failed"] = True
                                entry["error"] = f"去掉黄色失败: {err}"
                        raise
                    finally:
                        _WRITES_IDLE.set()
                else:
                    _preview(session, payload)
            except _Paused:
                continue
            except Exception as err:
                _task_failed(kind, payload, err, blocked)
                session.close()
    finally:
        session.close()
        pythoncom.CoUninitialize()


def _task_failed(kind: str, payload: Any, err: Exception, blocked: dict[str, float]) -> None:
    """做图或写入出错：先重试几次（多半是 Excel 一时忙不过来），还不行再标成出错、转人工"""
    gids = payload[1] if kind == "apply" else [payload] if kind == "preview" else []
    retrying = False
    with _LOCK:
        for gid in gids:
            group = AUTOFIX["groups"].get(gid)
            if group is None:
                continue
            group["error_count"] = group.get("error_count", 0) + 1
            if group["error_count"] <= MAX_ERROR_RETRIES:
                retrying = True
                group["message"] = f"出错了，稍后自动重试（第 {group['error_count']} 次）: {err}"
                if kind == "preview":
                    group.update(status="queued", attempts=[])
                # 写入的保持 applying / reverting，过一会儿再写
            else:
                group["status"] = "error"
                group["message"] = f"自动处理失败，请人工处理: {err}"
    if retrying and kind == "apply":
        blocked[payload[0]] = time.monotonic() + LOCKED_RETRY_SECONDS
    log_msg(f"⚠️ 自动处理{'出错，稍后重试' if retrying else '失败'}: {err}")


def _next_task(blocked: dict[str, float]) -> tuple[str, Any] | None:
    now = time.monotonic()
    with STATE_LOCK:
        # 标黄、改文件名只在没有任何任务（包括扫描）时做，免得和扫描抢文件、改掉扫描清单里的路径
        to_mark = [] if STATE["busy"] else [
            path for path, entry in STATE["review_queue"].items()
            if review_entry_needs_mark(path, entry) and blocked.get(path, 0) <= now
        ]
        to_restore = [] if STATE["busy"] else [
            path for path, entry in STATE["review_queue"].items()
            if review_entry_needs_restore(path, entry) and blocked.get(path, 0) <= now
        ]
    with _LOCK:
        groups = sorted(AUTOFIX["groups"].values(), key=lambda g: g["seq"])
        for group in groups:
            if group["status"] in WRITE_STATUSES and blocked.get(group["wb_path"], 0) <= now:
                same_file = [
                    g["id"] for g in groups if g["status"] in WRITE_STATUSES and g["wb_path"] == group["wb_path"]
                ]
                return "apply", (group["wb_path"], same_file)
    if to_mark:
        return "mark", to_mark[0]
    if to_restore:
        return "restore", to_restore[0]
    with _LOCK:
        groups = sorted(AUTOFIX["groups"].values(), key=lambda g: g["seq"])
        for group in groups:
            if group["status"] == "queued" and not group["manual"] and _preview_allowed():
                return "preview", group["id"]
    return None


def _items(item_ids: list[str]) -> list[dict[str, Any]]:
    with STATE_LOCK:
        return [dict(STATE["items"][i]) for i in item_ids if i in STATE["items"]]


def _check_still_wanted(gid: str) -> None:
    with STATE_LOCK:
        with _LOCK:
            group = AUTOFIX["groups"].get(gid)
            if group is None or not file_available(group["wb_path"]):
                raise _Paused
    with _LOCK:
        group = AUTOFIX["groups"].get(gid)
        if group is None or group["status"] != "running" or group["manual"]:
            raise _Paused
    if _busy() or not _preview_allowed():
        raise _Paused


def _preview(session: _ExcelSession, gid: str) -> None:
    with _LOCK:
        group = AUTOFIX["groups"].get(gid)
        if group is None:
            raise _Paused
        group.update(status="running", attempts=[], recommended=None, chosen=None, message="")
        wb_path, sheet_name, row = group["wb_path"], group["sheet_name"], group["row"]
        item_ids = list(group["item_ids"])
    items = _items(item_ids)
    sheet = session.copy_of(wb_path).Sheets(sheet_name)
    if sheet.Visible != XL_SHEET_VISIBLE:
        raise ValueError("工作表已隐藏")
    cells = snapshot_cells(extract_sheet_snapshot(sheet))
    probe = CellProbe(sheet)
    texts: dict[tuple[int, int], str | None] = {}

    def display_text(r: int, c: int) -> str | None:
        if (r, c) not in texts:
            texts[(r, c)] = probe.text(r, c)
        return texts[(r, c)]

    base = round(float(sheet.Rows(row).RowHeight or 15.0), 1)
    with _LOCK:
        group["base_height"] = base
    pdf_path = session.temp_pdf()
    success_at = None
    extra = 0.0
    try:
        for k in range(MAX_ATTEMPTS + 1):
            _check_still_wanted(gid)
            height = min(MAX_ROW_HEIGHT, round(base + extra, 1))
            with _LOCK:
                if k and height == group["attempts"][-1]["height"]:
                    break
            sheet.Rows(row).RowHeight = height
            texts.clear()  # 行高变化后 Range.Text 也可能变化，不能沿用上一张候选图的显示值
            _export_sheet_pdf(sheet, pdf_path)
            ok, image, need = _evaluate(pdf_path, items, cells, display_text, probe.printed)
            with _LOCK:
                group["attempts"].append({"height": height, "ok": ok, "image": image})
            if success_at is not None:
                break  # 显示全之后再多加一次就停
            if k and ok:
                success_at = k
            # 多行文字缺好几行时 3pt 一次加不完 8 次，按缺的行数直接补上（仍按 3pt 取整）
            extra = max(extra + STEP_PT, math.ceil(round((extra + need) / STEP_PT, 3)) * STEP_PT)
    except _Paused:
        with _LOCK:
            if gid in AUTOFIX["groups"] and AUTOFIX["groups"][gid]["status"] == "running":
                AUTOFIX["groups"][gid].update(status="queued", attempts=[])
        raise
    finally:
        sheet.Rows(row).RowHeight = base
        pdf_path.unlink(missing_ok=True)

    with _LOCK:
        group.pop("error_count", None)
        oks = [i for i, a in enumerate(group["attempts"]) if i and a["ok"]]
        group["recommended"] = oks[-1] if oks else None
        if not oks:
            group["status"] = "failed"
            group["message"] = f"加高 {MAX_ATTEMPTS} 次仍未通过完整显示与压线检查，可能定位不对或是列宽不够，请人工处理"
        elif AUTOFIX["auto"]:
            group.update(status="applying", chosen=group["recommended"], confirmed=False,
                         message="全自动：正在写入 Excel...")
            _WAKE.set()
        else:
            group["status"] = "ready"


def _evaluate(
    pdf_path: Path, items: list[dict[str, Any]], cells, display_text, printed,
) -> tuple[bool, str, float]:
    """要求每个目标单元格取得完整匹配、字符可见且无压线；没有证据不能判通过。
    第三个返回值：还有整行没打印出来时，大约还要加多高（pt），否则为 0。"""
    all_ok = bool(items)
    need = 0.0
    boxes: list[tuple[int, tuple[float, float, float, float]]] = []
    with fitz.open(pdf_path) as doc:
        pieces = collect_text_pieces(doc)
        matches: dict[tuple[int, int], dict[str, Any]] = {}
        hidden_results = find_hidden_cells(cells, pieces, display_text, printed, matches=matches)
        by_address = {
            f"{col_index_to_letters(col)}{row}": evidence for (row, col), evidence in matches.items()
        }
        hidden_by_address = {
            f"{col_index_to_letters(r['col'])}{r['row']}": r for r in hidden_results
        }
        page_lines = {}
        for item in items:
            address = str(item.get("cell_addr", "")).upper()
            evidence = by_address.get(address)
            found = hidden_by_address.get(address)
            ok = evidence is not None and found is None
            if found is not None and found.get("missing_lines"):
                need = max(need, found["missing_lines"] * found["line_pitch"])
            if item.get("kind") == "hidden":
                ok = ok and normalize_text(evidence["text"]) == normalize_text(item["text"])
                location = (
                    (found["page"], found["bbox"]) if found and found["bbox"] is not None
                    else (evidence["page"], evidence["bbox"]) if evidence is not None
                    else locate_text(pieces, item["text"])
                )
            else:
                try:
                    still, _, bbox, page = is_region_still_clipped(
                        pdf_path, int(item["page_index"]), item["text"], as_bbox(item["bbox"]),
                        occurrence=item.get("occurrence"),
                    )
                    ok, location = ok and not still, (page, bbox)
                except RegionTrackingError:
                    ok, location = False, None
            if evidence is not None:
                page_index = evidence["page"]
                if page_index not in page_lines:
                    page_lines[page_index] = extract_page_lines(doc[page_index])[0]
                if any(
                    not shown or crossing_horizontal_line(bbox, page_lines[page_index]) is not None
                    for char, bbox, shown in evidence["chars"] if normalize_text(char)
                ):
                    ok = False
            all_ok = all_ok and ok
            if location is not None:
                boxes.append(location)
        image = ""
        if boxes:
            page = boxes[0][0]
            same_page = [b for p, b in boxes if p == page]
            union = (
                min(b[0] for b in same_page), min(b[1] for b in same_page),
                max(b[2] for b in same_page), max(b[3] for b in same_page),
            )
            image = render_crop_base64(doc[page], union, highlight=not all_ok)
    return all_ok, image, need


def _apply(session: _ExcelSession, wb_path: str, gids: list[str]) -> bool:
    """把选中的行高（退回的改回原行高）写进原文件并保存；原文件正被 Excel 打开时返回 False，稍后再试"""
    with STATE_LOCK:
        if not file_available(wb_path):
            return False
    path = Path(wb_path)
    if path.with_name("~$" + path.name).exists():
        with _LOCK:
            for gid in gids:
                AUTOFIX["groups"][gid]["message"] = "文件正在 Excel 中打开，关闭后自动写入" + STALE_LOCK_HINT
        return False
    with _LOCK:
        plan = [
            (gid, g["sheet_name"], g["row"], g["base_height"],
             g["base_height"] if g["status"] == "reverting" else g["attempts"][g["chosen"]]["height"],
             g["chosen"], list(g["item_ids"]), g["status"] == "reverting", g["confirmed"])
            for gid in gids if (g := AUTOFIX["groups"].get(gid)) and g["status"] in WRITE_STATUSES
        ]
    if not plan:
        return True
    excel = session.app()
    workbook = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0)
    try:
        for _, sheet_name, row, _, height, *_ in plan:
            workbook.Sheets(sheet_name).Rows(row).RowHeight = height
        normalize_workbook_view(excel, workbook)
        workbook.Save()
    finally:
        workbook.Close(SaveChanges=False)

    with STATE_LOCK:
        reverted: list[str] = []
        for _, sheet_name, row, base, height, chosen, item_ids, revert, confirmed in plan:
            for item_id in item_ids:
                item = STATE["items"].get(item_id)
                if item is None:
                    continue
                item["row_height"] = height
                if revert:
                    item.pop("autofix_unconfirmed", None)
                    item.pop("autofix_base_height", None)
                    reverted.append(item_id)
                    continue
                item["status"] = "fixed"
                item["iterations"] = chosen
                item["review_reason"] = f"自动处理：行高 {base}pt → {height}pt"
                item["autofix_base_height"] = base
                if confirmed:
                    item.pop("autofix_unconfirmed", None)
                else:
                    item["autofix_unconfirmed"] = True
            if revert:
                log_msg(f"↩️ 自动处理退回 [{path.name} / {sheet_name} 行{row}] 行高已改回 {base}pt，转稍后人工复检")
            else:
                log_msg(f"✅ 自动处理 [{path.name} / {sheet_name} 行{row}] 行高 {base}pt → {height}pt，已保存")
        with _LOCK:
            for gid, *_, revert, _ in plan:
                if gid in AUTOFIX["groups"]:
                    AUTOFIX["groups"][gid].pop("error_count", None)
                    AUTOFIX["groups"][gid].update(
                        status="reverted" if revert else "applied", reapply=False,
                        message="已改回原行高，转稍后人工复检" if revert else "已写入")
        if reverted:
            queue_later_review(reverted)  # 一并保存状态
        else:
            _persist_state()
        _sync_problem_counts(wb_path, f"自动处理写入 {len(plan)} 行")
    return True
