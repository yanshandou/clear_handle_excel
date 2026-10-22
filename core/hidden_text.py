"""隐藏截断检测：拿 Excel 单元格的完整内容和导出的 PDF 对比，找出打印时没显示全的文字。

现有的切字检查只能发现「字被边框横线切穿」；行高不够时 Excel 会整行不画（自动换行的后几行直接消失），
列宽不够时超出部分被裁剪区域遮住，这两种从 PDF 外观上都看不出来。这里的判断依据：

1. PDF 里找不到完整内容、只能找到开头一段：后面的行没打印出来（行高不够）
2. 内容都在 PDF 里，但有字落在裁剪区域外：打印出来看不到（列宽 / 行高不够被裁掉）
3. 数字显示成 ###：列宽不够
4. PDF 里完全找不到：只要这格会被打印（不在隐藏行列、在打印区域内），也算没显示

对不上的一律报告（宁可多报不漏报）。
拼接多行时保留原文空行的信息，避免把空行后仍完整显示的下一段误报为截断。
"""
import re
import unicodedata
from collections.abc import Callable
from dataclasses import dataclass, field
from functools import cache
from typing import Any

from core.pdf_detector import BBox, fitz

# 不插入 MuPDF 推测的空格：同一行相邻两格的文字靠下面的间距判断分开
_BASE_FLAGS = fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_INHIBIT_SPACES
# 只取裁剪区域内的字（PyMuPDF 1.25 起，旧版只按页面边界裁剪）
_CLIP_FLAG = getattr(fitz, "TEXT_CLIP", fitz.TEXT_MEDIABOX_CLIP)
_SPACE_PATTERN = re.compile(r"\s+")

REASON_TRUNCATED = "PDF 里只打印出开头一部分，后面的文字没显示（行高或列宽不够）"
REASON_CLIPPED = "有文字超出单元格被裁掉，打印出来看不到（列宽或行高不够）"
REASON_HASH = "列宽不够，数字显示成了 ###"
REASON_MISSING = "PDF 里找不到这格的文字，打印时没有显示（行高或列宽太小）"


def normalize_text(text: str) -> str:
    """比对用：统一全角半角，去掉所有空白（换行位置在 Excel 和 PDF 里不同）"""
    return _SPACE_PATTERN.sub("", unicodedata.normalize("NFKC", text))


def _blank_line_offsets(text: str) -> dict[int, int]:
    """原文每个段落起点前的空行数，位置使用去空白后的字符偏移。"""
    offsets: dict[int, int] = {}
    pos = 0
    blanks = 0
    for line in text.splitlines():
        key = normalize_text(line)
        if not key:
            if pos:
                blanks += 1
            continue
        if blanks:
            offsets[pos] = blanks
        pos += len(key)
        blanks = 0
    return offsets


@dataclass
class Piece:
    """PDF 同一行里连续的一段文字，通常就是一个单元格的一行"""
    page: int
    chars: list[tuple[str, BBox, bool]]  # (字, 坐标, 是否在裁剪区域内可见)
    key: str = ""
    owner: list[tuple[int, ...]] = field(default_factory=list)  # 合成字可能对应多个原始字符
    used: list[bool] = field(default_factory=list)

    def __post_init__(self) -> None:
        cluster = ""
        indices: set[int] = set()

        def flush() -> None:
            nonlocal cluster
            key = normalize_text(cluster)
            self.key += key
            self.owner.extend([tuple(sorted(indices))] * len(key))
            cluster = ""
            indices.clear()

        for index, (char, _, _) in enumerate(self.chars):
            # 先兼容分解，再把基字与组合符一起合成；不能逐字 NFKC，
            # 否则半角「ﾌ」「ﾞ」会留下「フ + 组合浊点」，对不上原文的「ブ」。
            for c in unicodedata.normalize("NFKD", char):
                if cluster and not unicodedata.combining(c):
                    before = unicodedata.normalize("NFC", cluster)
                    if unicodedata.normalize("NFC", cluster + c) == before + c:
                        flush()
                cluster += c
                indices.add(index)
        flush()
        self.used = [False] * len(self.key)
        xs0, ys0, xs1, ys1 = zip(*(b for _, b, _ in self.chars))
        self.bbox: BBox = (min(xs0), min(ys0), max(xs1), max(ys1))

    @property
    def height(self) -> float:
        return max(1.0, self.bbox[3] - self.bbox[1])

    def free(self, start: int, end: int) -> bool:
        return not any(self.used[start:end])


Chunk = tuple[Piece, int, int]  # piece.key[start:end]


@cache
def clip_check_supported() -> bool:
    """当前 PyMuPDF 能否分辨被裁剪区域遮住的字；不能时「被裁掉」这一类查不出来，需要升级"""
    with fitz.open() as doc:
        page = doc.new_page()
        page.insert_text((0, 0), " ", fontsize=1, fontname="helv")
        xref = page.get_contents()[-1]
        doc.update_stream(
            xref, doc.xref_stream(xref) + b"\nq 50 700 20 14 re W n BT /helv 10 Tf 52 705 Td (ABCDEFGH) Tj ET Q"
        )
        data = doc.tobytes()
    with fitz.open("pdf", data) as doc:
        return "H" not in doc[0].get_text("text", flags=_BASE_FLAGS | _CLIP_FLAG)


def _char_key(c: str, bbox: Any) -> tuple[str, int, int]:
    return (c, round(bbox[0] * 4), round(bbox[1] * 4))


def collect_text_pieces(doc: fitz.Document) -> list[Piece]:
    """按行提取 PDF 文字，同一行里间距明显的地方切开；同时标出每个字是否落在裁剪区域外。"""
    pieces: list[Piece] = []
    for page_index in range(len(doc)):
        page = doc[page_index]
        visible: dict[tuple[str, int, int], int] = {}
        for block in page.get_text("rawdict", flags=_BASE_FLAGS | _CLIP_FLAG)["blocks"]:
            for line in block.get("lines", []):
                for span in line["spans"]:
                    for ch in span["chars"]:
                        key = _char_key(ch["c"], ch["bbox"])
                        visible[key] = visible.get(key, 0) + 1
        for block in page.get_text("rawdict", flags=_BASE_FLAGS)["blocks"]:
            for line in block.get("lines", []):
                current: list[tuple[str, BBox, bool]] = []
                prev_x1 = None
                for span in line["spans"]:
                    gap_limit = max(1.5, float(span.get("size", 10)) * 0.15)
                    for ch in span["chars"]:
                        bbox = tuple(float(v) for v in ch["bbox"])
                        if prev_x1 is not None and bbox[0] - prev_x1 > gap_limit and current:
                            pieces.append(Piece(page_index, current))
                            current = []
                        key = _char_key(ch["c"], ch["bbox"])
                        shown = visible.get(key, 0) > 0
                        if shown:
                            visible[key] -= 1
                        current.append((ch["c"], bbox, shown))  # type: ignore[arg-type]
                        prev_x1 = bbox[2]
                if current:
                    pieces.append(Piece(page_index, current))
    return [p for p in pieces if p.key]


class _Matcher:
    def __init__(self, pieces: list[Piece]):
        self.pieces = pieces
        self.by_char: dict[str, list[Piece]] = {}
        for piece in pieces:
            for c in set(piece.key):
                self.by_char.setdefault(c, []).append(piece)
        self.by_page: dict[int, list[Piece]] = {}
        for piece in pieces:
            self.by_page.setdefault(piece.page, []).append(piece)

    def _below(self, prev: Piece, blank_lines: int = 0) -> list[Piece]:
        """同页、横向重叠的下一行；仅在原文确有空行处扩大纵向间距。"""
        x0, y0, x1, y1 = prev.bbox
        max_y = y1 + prev.height * (1 + blank_lines)
        if blank_lines:
            max_y += 1.0  # 给空行排版和 PDF 坐标取整留少量余量
        return [
            p for p in self.by_page[prev.page]
            if y0 + prev.height * 0.3 < p.bbox[1] < max_y
            and p.bbox[0] < x1 + 2 and p.bbox[2] > x0 - 2
        ]

    def _chunks(
        self, text: str, pos: int, prev: Piece | None, loose: bool, blank_lines: int = 0,
    ) -> list[Chunk]:
        """text[pos:] 开头能对上的片段。exact 只认整段；loose 允许首段是某段的后半、末段是某段的前半。"""
        candidates = self.by_char.get(text[pos], []) if prev is None else self._below(prev, blank_lines)
        found: list[Chunk] = []
        for piece in candidates:
            starts = [0] if prev is not None or not loose else [
                i for i, c in enumerate(piece.key) if c == text[pos]
            ]
            for start in starts:
                length = 0
                while (
                    start + length < len(piece.key) and pos + length < len(text)
                    and piece.key[start + length] == text[pos + length]
                ):
                    length += 1
                if not length or not piece.free(start, start + length):
                    continue
                end = start + length
                done = pos + length == len(text)
                if not loose and (start or end != len(piece.key)):
                    continue
                # 没结束就要换行，这一段必须到行尾
                if not done and end != len(piece.key):
                    continue
                found.append((piece, start, end))
        return found

    def find(self, text: str, loose: bool) -> list[Chunk] | None:
        """找一组能拼出完整 text 的片段（自上而下的换行），找不到返回 None"""
        blank_lines = _blank_line_offsets(text)
        text = normalize_text(text)

        def search(pos: int, prev: Piece | None, chain: list[Chunk]) -> list[Chunk] | None:
            for chunk in self._chunks(text, pos, prev, loose, blank_lines.get(pos, 0)):
                piece, start, end = chunk
                nxt = pos + end - start
                if nxt == len(text):
                    return chain + [chunk]
                result = search(nxt, piece, chain + [chunk])
                if result:
                    return result
            return None
        return search(0, None, []) if text else None

    def find_prefix(self, text: str) -> list[Chunk] | None:
        """只能拼出 text 开头一段时，返回能拼出的最长一段（后面的没打印出来）"""
        blank_lines = _blank_line_offsets(text)
        text = normalize_text(text)
        best: list[Chunk] = []
        best_len = 0

        def search(pos: int, prev: Piece | None, chain: list[Chunk]) -> None:
            nonlocal best, best_len
            for chunk in self._chunks(text, pos, prev, prev is None, blank_lines.get(pos, 0)):
                piece, start, end = chunk
                if end != len(piece.key):
                    continue
                nxt = pos + end - start
                if nxt >= len(text):
                    continue
                if nxt > best_len:
                    best, best_len = chain + [chunk], nxt
                search(nxt, piece, chain + [chunk])
        if text:
            search(0, None, [])
        return best or None


def _consume(chain: list[Chunk]) -> None:
    for piece, start, end in chain:
        for i in range(start, end):
            piece.used[i] = True


def _chain_chars(chain: list[Chunk]) -> list[tuple[str, BBox, bool]]:
    chars: list[tuple[str, BBox, bool]] = []
    for piece, start, end in chain:
        indices = sorted({index for source_indices in piece.owner[start:end] for index in source_indices})
        chars.extend(piece.chars[i] for i in indices)
    return chars


def _chain_location(chain: list[Chunk]) -> tuple[int, BBox]:
    boxes = [b for _, b, _ in _chain_chars(chain)]
    return chain[0][0].page, (
        min(b[0] for b in boxes), min(b[1] for b in boxes),
        max(b[2] for b in boxes), max(b[3] for b in boxes),
    )


def _guess_display(value: Any) -> str:
    """不跨进程读取时先猜显示文字；猜不中再读 Excel 的 Range.Text"""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def find_hidden_cells(
    cells: list[tuple[int, int, Any]],
    pieces: list[Piece],
    display_text: Callable[[int, int], str | None],
    is_printed: Callable[[int, int], bool],
) -> list[dict[str, Any]]:
    """
    cells: [(行, 列, 单元格值)]；display_text 读取 Excel 显示的文字（Range.Text）；
    is_printed 判断这格会不会被打印（隐藏行列、打印区域外的不算）。
    返回: [{"row", "col", "text", "visible", "reason", "page", "bbox"}]，page/bbox 在 PDF 里完全找不到时为 None
    """
    matcher = _Matcher(pieces)
    results: list[dict[str, Any]] = []
    pending: list[dict[str, Any]] = []
    hash_cells: list[tuple[dict[str, Any], str]] = []
    for row, col, value in cells:
        if value is None or (isinstance(value, str) and not normalize_text(value)):
            continue
        cell = {"row": row, "col": col, "value": value, "text": value}
        if not isinstance(value, str):
            # 数字、日期的显示文字取决于格式，必须读 Excel；猜的话「1」这类常见文字会被别处的「1」顶替
            shown = display_text(row, col)
            if shown is None:
                shown = _guess_display(value)
            if not normalize_text(shown):
                continue
            if set(shown.strip()) == {"#"}:
                hash_cells.append((cell, shown))
                continue
            cell["text"] = shown
        pending.append(cell)

    def match_full(group: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """先整段严格匹配、再放宽，长的先配，避免短文字占掉长文字的位置"""
        for loose in (False, True):
            left = []
            for cell in sorted(group, key=lambda c: -len(normalize_text(c["text"]))):
                chain = matcher.find(cell["text"], loose)
                if chain is None:
                    left.append(cell)
                    continue
                _consume(chain)
                hidden = [c for c, _, shown in _chain_chars(chain) if not shown]
                if hidden:
                    page, bbox = _chain_location(chain)
                    shown_text = "".join(c for c, _, shown in _chain_chars(chain) if shown)
                    results.append(_result(cell, shown_text, REASON_CLIPPED, page, bbox))
            group = left
        return group

    for cell, shown in hash_cells:
        chain = matcher.find(normalize_text(shown), True)
        page, bbox = _chain_location(chain) if chain else (None, None)
        if chain:
            _consume(chain)
        results.append(_result(cell, shown, REASON_HASH, page, bbox, text=_guess_display(cell["value"])))

    left = match_full(pending)

    # 文字对不上的再读一次 Excel 实际显示的文字（如「"No."@」这类带格式的文字）
    unchanged, refetched = [], []
    for cell in left:
        shown = display_text(cell["row"], cell["col"]) if isinstance(cell["value"], str) else None
        if shown is None or shown == cell["text"]:
            unchanged.append(cell)
        elif normalize_text(shown):
            cell["text"] = shown
            refetched.append(cell)
        # 显示为空：格式设成不显示（如 ;;;），不算问题
    left = unchanged + match_full(refetched)

    for cell in sorted(left, key=lambda c: -len(normalize_text(c["text"]))):
        chain = matcher.find_prefix(cell["text"])
        if chain is not None:
            _consume(chain)
            page, bbox = _chain_location(chain)
            shown_text = "".join(c for c, _, _ in _chain_chars(chain))
            results.append(_result(cell, shown_text, REASON_TRUNCATED, page, bbox))
        elif is_printed(cell["row"], cell["col"]):
            results.append(_result(cell, "", REASON_MISSING, None, None))
    results.sort(key=lambda r: (r["row"], r["col"]))
    return results


def _result(
    cell: dict[str, Any], visible: str, reason: str, page: int | None, bbox: BBox | None,
    *, text: str | None = None,
) -> dict[str, Any]:
    return {
        "row": cell["row"], "col": cell["col"], "text": text if text is not None else cell["text"],
        "visible": visible, "reason": reason, "page": page, "bbox": bbox,
    }


def locate_text(pieces: list[Piece], text: str) -> tuple[int, BBox] | None:
    """PDF 里这段文字（完整或开头一段）所在的页码和范围，找不到返回 None；只读，不占用片段"""
    matcher = _Matcher(pieces)
    chain = matcher.find(text, True) or matcher.find_prefix(text)
    return _chain_location(chain) if chain else None
