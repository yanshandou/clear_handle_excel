"""Excel 对比报告（用户 2026-11-17 要）：任选两份 Excel，把所有不同找出来，生成一个 HTML 文件。

- 怎么比：core/excel_diff.diff(full=True)，和上传前的对比页同一套检查，只是清单一条不截、效果图画得更多、
  不管精简会不会删那张表；另外多比了工作簿级的定义名称 / 工作簿保护。
- 报告是单个 .html：样式 static/excel_report.css、画法 static/excel_report_view.js、对比数据（JSON）都塞在文件里，
  发给别人双击就能看，不用连本程序、不用联网。报告里不写日期时间。
- 生成的报告另存一份在 程序数据/excel_reports/，网页上可以直接下载。
"""
from __future__ import annotations

import html
import json
import re
import threading
import time
from pathlib import Path
from typing import Any

from core import debug_log, excel_diff

STATIC_DIR = Path(__file__).resolve().parent.parent / "static"
_lock = threading.Lock()
_state: dict[str, Any] = {"running": False, "stage": "", "done": 0, "total": 0, "started": 0.0}


def reports_dir() -> Path:
    from core.data_dir import DATA_DIR
    return Path(DATA_DIR) / "excel_reports"


def status() -> dict[str, Any]:
    with _lock:
        state = dict(_state)
    state["elapsed"] = round(time.monotonic() - state["started"], 1) if state["running"] else 0
    return state


def _progress(stage: str, done: int, total: int) -> None:
    with _lock:
        _state.update(stage=stage, done=done, total=total)
    debug_log.event("📊 对比报告进度", 阶段=stage, 做完=done, 共=total)


def _safe_name(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", text).strip(" .")[:60] or "Excel"


def file_name(name_a: str, name_b: str) -> str:
    return f"Excel对比报告_{_safe_name(Path(name_a).stem)}_vs_{_safe_name(Path(name_b).stem)}.html"


def render_html(result: dict[str, Any]) -> str:
    """把 excel_diff.diff 的结果做成单个自带样式和脚本的 HTML"""
    css = (STATIC_DIR / "excel_report.css").read_text(encoding="utf-8")
    script = (STATIC_DIR / "excel_report_view.js").read_text(encoding="utf-8")
    # JSON 放在 <script type="application/json"> 里：把 < > & 转成 \\u 写法，文字里出现 </script> 也不会提前结束
    data = json.dumps(result, ensure_ascii=False, separators=(",", ":"))
    data = data.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    title = f"Excel 对比报告 - {Path(result['theirs']).name} vs {Path(result['mine']).name}"
    return ("<!DOCTYPE html>\n<html lang=\"zh-CN\">\n<head>\n<meta charset=\"UTF-8\">\n"
            "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n"
            f"<title>{html.escape(title)}</title>\n<style>\n{css}\n</style>\n</head>\n<body>\n"
            "<div id=\"xr-root\"><p>正在显示报告…（如果一直停在这里，请用 Chrome / Edge 打开）</p></div>\n"
            f"<script id=\"xr-data\" type=\"application/json\">{data}</script>\n"
            "<script>\n" + script.replace("</script", "<\\/script") + "\n</script>\n</body>\n</html>\n")


def build(path_a: Path | str, path_b: Path | str, label_a: str = "", label_b: str = "",
          save: bool = True) -> dict[str, Any]:
    """path_a = 文件 A（左 / 旧），path_b = 文件 B（右 / 新）。label_* = 报告上显示的路径（上传的文件写原来的文件名）"""
    path_a, path_b = Path(path_a), Path(path_b)
    with _lock:
        if _state["running"]:
            raise ValueError("已经在生成一份对比报告了，请等它做完再生成下一份")
        _state.update(running=True, stage="开始", done=0, total=0, started=time.monotonic())
    started = time.monotonic()
    try:
        with debug_log.side_task("excel_report", {"A": label_a or str(path_a), "B": label_b or str(path_b)}):
            result = excel_diff.diff(path_b, path_a, full=True, progress=_progress)
            result["label_a"], result["label_b"] = label_a or str(path_a), label_b or str(path_b)
            _progress("正在生成报告", len(result["sheets"]), len(result["sheets"]))
            page = render_html(result)
            name = file_name(Path(result["label_a"]).name, Path(result["label_b"]).name)
            saved = ""
            if save:
                folder = reports_dir()
                folder.mkdir(parents=True, exist_ok=True)
                target = folder / name
                target.write_text(page, encoding="utf-8")
                saved = str(target)
            totals = result["summary"]["各类"]
            debug_log.event("📊 对比报告生成", A=result["label_a"], B=result["label_b"], 完全一样=result["identical"],
                            表=len(result["sheets"]), 各类="、".join(f"{k}{v}" for k, v in totals.items()) or "-",
                            大小KB=round(len(page.encode("utf-8")) / 1024), 存到=saved or "-",
                            用时秒=round(time.monotonic() - started, 1))
            sheets = result["sheets"]
            return {"html": page, "filename": name, "saved": saved, "identical": result["identical"],
                    "totals": totals, "order_changed": result["summary"]["表的顺序变了"], "book": len(result.get("book", [])),
                    "sheets": {key: sum(s["status"] == key for s in sheets) for key in ("changed", "only_old", "only_new", "same")}}
    finally:
        with _lock:
            _state.update(running=False, stage="", done=0, total=0)


MAX_UPLOAD_MB = 200


def saved_report(name: str) -> Path:
    """下载用：只给 excel_reports 文件夹里直接放着的 .html"""
    folder = reports_dir().resolve()
    target = (folder / Path(name).name).resolve()
    if not name or target.parent != folder or target.suffix.lower() != ".html" or not target.is_file():
        raise ValueError(f"找不到这份报告：{name}")
    return target


def build_from_request(side_a: Any, side_b: Any) -> dict[str, Any]:
    """网页来的请求：每边要么是 {path: 完整路径}，要么是 {name: 文件名, data: base64}（在网页上选 / 拖进来的文件）"""
    import base64
    import shutil
    import tempfile
    temp = None
    try:
        paths, labels = [], []
        for key, side in (("A", side_a), ("B", side_b)):
            if not isinstance(side, dict):
                raise ValueError(f"请先选文件 {key}")
            path = str(side.get("path") or "").strip().strip('"')
            if path:
                paths.append(Path(path))
                labels.append(path)
                continue
            name, data = str(side.get("name") or "").strip(), side.get("data")
            if not name or not isinstance(data, str) or not data:
                raise ValueError(f"请先选文件 {key}（填完整路径，或点「选择文件」）")
            if len(data) > MAX_UPLOAD_MB * 1024 * 1024 * 4 // 3:
                raise ValueError(f"文件 {key} 太大（超过 {MAX_UPLOAD_MB} MB），请改填完整路径")
            if temp is None:
                temp = Path(tempfile.mkdtemp(prefix="excel_report_"))
            target = temp / key / Path(name).name  # 两边同名也不会互相盖住
            target.parent.mkdir(parents=True, exist_ok=True)
            try:
                target.write_bytes(base64.b64decode(data, validate=True))
            except ValueError as err:
                raise ValueError(f"文件 {key} 传过来坏了，请重新选一次（{err}）") from err
            paths.append(target)
            labels.append(f"{name}（在网页上选的文件）")
        return build(paths[0], paths[1], labels[0], labels[1])
    finally:
        if temp is not None:
            shutil.rmtree(temp, ignore_errors=True)
