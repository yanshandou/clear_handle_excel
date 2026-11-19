"""PDF 左右对比（用户 2026-11-19）：像 Git 一样左右两栏看两边的文字哪里不同。

- 网页上把 PDF 拖进左边或右边，每份先传上来存一份（store），每边可以放好几份，按放进去的顺序接起来再比。
  例：A + B 合起来应该等于 C，就把 A、B 拖到左边，C 拖到右边。
- 比的是文字：每页按位置把同一高度的字连成一行（Excel 导出的 PDF 一行就是表里的一行），
  左右两边所有行接起来以后用 difflib 逐行比，改了的行再逐字标出哪几个字不同。
- 页眉页脚默认不比：A+B 和 C 的分页、页码本来就不一样。页码 = 只有页码的行（「1 / 5」「- 3 -」「Page 2」「2ページ」）；
  页眉页脚 = 左右两边所有的页里，一半以上的页都在页面最上面 / 最下面两行出现的同一行字（每页重复的标题、公司名、打印标题行）。
- 只读：传上来的 PDF 放在数据文件夹的 pdf_diff 子文件夹，过一天自动清掉，原文件不动。
"""
from __future__ import annotations

import difflib
import hashlib
import re
import threading
import time
import uuid
from pathlib import Path
from typing import Any

from core import debug_log
from core.data_dir import DATA_DIR
from core.pdf_detector import fitz

MAX_BYTES = 300 * 1024 * 1024  # 一份最多 300 MB
KEEP_SECONDS = 24 * 3600        # 传上来的 PDF 留一天
CHAR_DIFF_LIMIT = 2000          # 一行太长就不再逐字比，整行标色
EDGE_LINES = 2                  # 页面最上面 / 最下面几行算页眉页脚的位置

_lock = threading.Lock()
_text_cache: dict[str, list[dict[str, Any]]] = {}  # 文件 id → 每页的行（同一份换了选项再比不用重读）


def _folder() -> Path:
    folder = Path(DATA_DIR) / "pdf_diff"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _cleanup(folder: Path) -> None:
    now = time.time()
    for path in folder.glob("*.pdf"):
        try:
            if now - path.stat().st_mtime > KEEP_SECONDS:
                path.unlink()
                _text_cache.pop(path.stem, None)
        except OSError:
            pass


def _path(file_id: Any) -> Path:
    if not isinstance(file_id, str) or not re.fullmatch(r"[0-9a-f]{32}", file_id):
        raise ValueError("文件编号不对，请重新拖入这份 PDF")
    path = _folder() / f"{file_id}.pdf"
    if not path.is_file():
        raise ValueError("这份 PDF 已经过期被清掉了，请重新拖入")
    return path


def store(name: str, body: bytes) -> dict[str, Any]:
    """存下网页传上来的一份 PDF，返回编号、页数。"""
    name = (name or "").strip() or "未命名.pdf"
    if not body:
        raise ValueError(f"{name} 是空文件")
    if len(body) > MAX_BYTES:
        raise ValueError(f"{name} 超过 300 MB，太大了")
    if not body.lstrip()[:5].startswith(b"%PDF"):
        raise ValueError(f"{name} 不是 PDF 文件")
    try:
        with fitz.open(stream=body, filetype="pdf") as doc:
            if doc.needs_pass:
                raise ValueError(f"{name} 有密码，打不开")
            pages = doc.page_count
    except ValueError:
        raise
    except Exception as err:  # PyMuPDF 打不开的坏文件
        raise ValueError(f"{name} 打不开：{err}") from None
    folder = _folder()
    with _lock:
        _cleanup(folder)
        file_id = uuid.uuid4().hex
        (folder / f"{file_id}.pdf").write_bytes(body)
    debug_log.event("📄 PDF 对比收到文件", 文件=name, 页数=pages, 大小=len(body), 编号=file_id)
    return {"id": file_id, "name": name, "pages": pages, "size": len(body),
            "sha256": hashlib.sha256(body).hexdigest()[:12]}


_path_ids: dict[tuple, dict[str, Any]] = {}  # (路径, 修改时间, 大小) → store 的结果：同一份在批量里出现好几次只读一次


def store_path(text: Any) -> dict[str, Any]:
    """批量对比：按本机路径读一份 PDF（贴进来的路径可带引号），文件没变过就用上次存的那份。"""
    if not isinstance(text, str) or not text.strip():
        raise ValueError("路径是空的")
    raw = text.strip().strip('"').strip("'").strip()
    path = Path(raw).expanduser()
    if path.suffix.lower() != ".pdf":
        raise ValueError(f"不是 PDF：{raw}")
    try:
        info = path.stat()
    except OSError:
        raise ValueError(f"找不到这份 PDF：{raw}") from None
    if not path.is_file():
        raise ValueError(f"不是文件：{raw}")
    key = (str(path.resolve()), info.st_mtime_ns, info.st_size)
    with _lock:
        cached = _path_ids.get(key)
    if cached and (_folder() / f"{cached['id']}.pdf").is_file():
        return cached
    try:
        body = path.read_bytes()
    except OSError as err:
        raise ValueError(f"读不了 {raw}：{err}") from None
    stored = {**store(path.name, body), "path": str(path)}
    with _lock:
        _path_ids[key] = stored
    return stored


PAGE_ZOOM = 1.5  # 页面图放大倍数（约 108 dpi）：网页上逐像素比，小字也看得清


ZOOM_SHARP = 3.0  # 鼠标移到框上弹出的放大图用更清楚的图


def page_png(file_id: Any, page: Any, sharp: bool = False) -> bytes:
    """传上来的那份 PDF 第 page 页（从 1 数）的图，网页上两边逐像素比、画红绿框；sharp 是放大图用的高清图"""
    try:
        number = int(page)
    except (TypeError, ValueError):
        raise ValueError("页码不对") from None
    with fitz.open(_path(file_id)) as doc:
        if not 1 <= number <= doc.page_count:
            raise ValueError("页码不对")
        return doc[number - 1].get_pixmap(matrix=fitz.Matrix(*(2 * [ZOOM_SHARP if sharp else PAGE_ZOOM])), alpha=False).tobytes("png")


def file_bytes(file_id: Any) -> bytes:
    return _path(file_id).read_bytes()


# ---------- 抽文字 ----------

def _page_lines(page: Any) -> list[str]:
    """把一页的字按高度连成行：同一高度（上下重叠一半以上）的词算一行，从左到右接起来；隔得远的中间放两个空格。"""
    words = page.get_text("words")
    if not words:
        return []
    words = sorted(words, key=lambda w: ((w[1] + w[3]) / 2, w[0]))
    rows: list[list[tuple]] = []
    row_mid = row_height = 0.0
    for w in words:
        mid, height = (w[1] + w[3]) / 2, max(w[3] - w[1], 1.0)
        if rows and abs(mid - row_mid) <= max(height, row_height) * 0.5:
            rows[-1].append(w)
            continue
        rows.append([w])
        row_mid, row_height = mid, height
    lines = []
    for row in rows:
        row.sort(key=lambda w: w[0])
        parts, right = [], None
        for w in row:
            if right is not None:
                gap = w[0] - right
                parts.append("  " if gap > (w[3] - w[1]) * 1.2 else " ")
            parts.append(w[4])
            right = w[2]
        text = "".join(parts).replace("\xad", "-").strip()
        if text:
            lines.append(text)
    return lines


def _read(file_id: str) -> list[dict[str, Any]]:
    with _lock:
        cached = _text_cache.get(file_id)
    if cached is not None:
        return cached
    pages = []
    with fitz.open(_path(file_id)) as doc:
        for index, page in enumerate(doc):
            pages.append({"page": index + 1, "lines": _page_lines(page)})
    with _lock:
        _text_cache[file_id] = pages
    return pages


_PAGE_NO = re.compile(
    r"^(?:[-－—–‐ー〜~\s]*\d{1,4}[-－—–‐ー〜~\s]*"            # 3 / - 3 -
    r"|\d{1,4}\s*[/／]\s*\d{1,4}"                               # 3 / 12
    r"|(?:page|p\.?)\s*\d{1,4}(?:\s*(?:of|[/／])\s*\d{1,4})?"   # Page 3 / Page 3 of 12
    r"|第?\s*\d{1,4}\s*(?:ページ|頁|页)(?:\s*[/／]\s*\d{1,4}\s*(?:ページ|頁|页)?)?"  # 3ページ / 第3页
    r")$", re.IGNORECASE)


def _is_page_no(text: str) -> bool:
    return bool(_PAGE_NO.match(text.strip()))


def _key(text: str, ignore_space: bool) -> str:
    text = text.replace("　", " ")
    return re.sub(r"\s+", "", text) if ignore_space else re.sub(r"\s+", " ", text).strip()


def _repeated(pages: list[dict[str, Any]]) -> set[str]:
    """每页都重复的页眉页脚：一半以上的页在最上面 / 最下面 EDGE_LINES 行里出现的同一行字（至少 3 页，免得两页对两页时把正好在边上的正文当成页眉）。
    同一页里出现两次只算一次。"""
    if len(pages) < 2:
        return set()
    counts: dict[str, int] = {}
    for page in pages:
        lines = page["lines"]
        edge = set(lines[:EDGE_LINES] + lines[-EDGE_LINES:])
        for text in edge:
            key = _key(text, True)
            counts[key] = counts.get(key, 0) + 1
    need = max(3, (len(pages) + 1) // 2)
    return {key for key, n in counts.items() if n >= need}


def _check_ids(ids: Any, label: str) -> list[str]:
    if not isinstance(ids, list) or not ids:
        raise ValueError(f"{label}还没有放 PDF")
    for file_id in ids:
        _path(file_id)
    return ids


def _side(ids: list[str], repeated: set[str], ignore_header: bool) -> tuple[list[dict], list[dict], int]:
    """一边的所有 PDF 按顺序接起来：返回 行、每份文件的信息、没比的页眉页脚行数。"""
    lines, files, dropped = [], [], 0
    for f_index, file_id in enumerate(ids):
        pages = _read(file_id)
        count = 0
        for page in pages:
            last = len(page["lines"]) - 1
            for index, text in enumerate(page["lines"]):
                at_edge = index < EDGE_LINES or index > last - EDGE_LINES
                if ignore_header and at_edge and (_is_page_no(text) or _key(text, True) in repeated):
                    dropped += 1
                    continue
                lines.append({"s": text, "f": f_index, "p": page["page"]})
                count += 1
        files.append({"pages": len(pages), "lines": count})
    return lines, files, dropped


def _char_diff(a: str, b: str) -> tuple[list, list]:
    """改了的一行：逐字标出哪几段不同。返回 [[开始, 结束], …]（左、右各自的位置）。"""
    if len(a) > CHAR_DIFF_LIMIT or len(b) > CHAR_DIFF_LIMIT:
        return [[0, len(a)]], [[0, len(b)]]
    left, right = [], []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag == "equal":
            continue
        if i2 > i1:
            left.append([i1, i2])
        if j2 > j1:
            right.append([j1, j2])
    return left, right


def diff(left_ids: Any, right_ids: Any, options: Any = None) -> dict[str, Any]:
    options = options if isinstance(options, dict) else {}
    ignore_space = options.get("ignore_space", True) is not False
    ignore_header = options.get("ignore_header", True) is not False
    started = time.perf_counter()
    left_ids, right_ids = _check_ids(left_ids, "左边"), _check_ids(right_ids, "右边")
    # 页眉页脚按左右两边所有的页一起认：只有一页的文件（例如 B）也能认出它的页眉
    repeated = _repeated([page for file_id in left_ids + right_ids for page in _read(file_id)]) if ignore_header else set()
    left, left_files, left_dropped = _side(left_ids, repeated, ignore_header)
    right, right_files, right_dropped = _side(right_ids, repeated, ignore_header)
    read_seconds = time.perf_counter() - started

    a = [_key(x["s"], ignore_space) for x in left]
    b = [_key(x["s"], ignore_space) for x in right]
    matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)  # 1.5 万行各改几十处实测 1 秒内
    ops, chars = [], {}
    removed = added = changed = 0
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        ops.append([tag[0], i1, i2, j1, j2])  # e / r / d / i
        if tag == "replace":
            pairs = min(i2 - i1, j2 - j1)
            changed += pairs
            removed += (i2 - i1) - pairs
            added += (j2 - j1) - pairs
            for k in range(pairs):
                chars[f"{i1 + k}:{j1 + k}"] = _char_diff(left[i1 + k]["s"], right[j1 + k]["s"])
        elif tag == "delete":
            removed += i2 - i1
        elif tag == "insert":
            added += j2 - j1
    hunks = sum(1 for op in ops if op[0] != "e")
    seconds = time.perf_counter() - started
    debug_log.event("📄 PDF 对比完成", 左边=len(left_ids), 右边=len(right_ids), 左行数=len(left), 右行数=len(right),
                    不同处=hunks, 改=changed, 删=removed, 增=added, 读文字秒=round(read_seconds, 2), 总秒=round(seconds, 2))
    return {
        "left": left, "right": right, "ops": ops, "chars": chars,
        "files": {"left": left_files, "right": right_files},
        "dropped": {"left": left_dropped, "right": right_dropped},
        "stats": {"hunks": hunks, "changed": changed, "removed": removed, "added": added,
                  "same": not hunks, "seconds": round(seconds, 2)},
        "options": {"ignore_space": ignore_space, "ignore_header": ignore_header},
    }
