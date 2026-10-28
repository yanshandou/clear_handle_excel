"""Sheet 级打印方向：比较实际 PDF，保存证据，临时应用后恢复原设置。"""
from contextlib import contextmanager
from functools import lru_cache, wraps
from pathlib import Path
import math
import shutil
import tempfile
import threading
import zipfile
from xml.etree import ElementTree
from typing import Any
from uuid import uuid4

from core.pdf_detector import fitz, find_clipped_spans_on_page
from core.hidden_text import collect_text_pieces, find_hidden_cells, normalize_text
from core.excel_ops import CellProbe, extract_sheet_snapshot, snapshot_cells

PORTRAIT, LANDSCAPE = 1, 2
RULE_VERSION = 1
MIN_FONT_PT = 8.0
SETTINGS = (
    "PaperSize", "LeftMargin", "RightMargin", "TopMargin", "BottomMargin",
    "HeaderMargin", "FooterMargin", "PrintArea", "PrintTitleRows", "PrintTitleColumns",
    "Zoom", "FitToPagesWide", "FitToPagesTall", "Order", "CenterHorizontally", "CenterVertically",
    "LeftHeader", "CenterHeader", "RightHeader", "LeftFooter", "CenterFooter", "RightFooter",
)
# 后台副本与源文件的关联。只保存路径；方案始终从最新 STATE 读取。
_sources: dict[str, str] = {}
_records = lambda: {}
_analysis = threading.local()


def configure(records) -> None:
    global _records
    _records = records
    _sources.clear()
    _xml_sheet_ids.cache_clear()


def label(value: int) -> str:
    return "横向" if value == LANDSCAPE else "竖向"


def version(path: str) -> dict:
    stat = Path(path).stat()
    return {"mtime_ns": stat.st_mtime_ns, "size": stat.st_size}


def _path(path: str) -> str:
    return str(Path(path).resolve())


def bind(book: Any, source: str) -> None:
    try:
        full = book.FullName
        if isinstance(full, str) and full:
            _sources[_path(full)] = _path(source)
    except Exception:
        pass


def source(sheet: Any) -> str | None:
    try:
        full = sheet.Parent.FullName
        if isinstance(full, str) and full:
            full = _path(full)
            return _sources.get(full, full)
    except Exception:
        pass
    return None


@lru_cache(maxsize=128)
def _xml_sheet_ids(path: str, mtime_ns: int, size: int) -> dict[str, str]:
    try:
        with zipfile.ZipFile(path) as archive:
            root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        return {sheet.attrib["name"]: sheet.attrib["sheetId"]
                for sheet in root.iter() if sheet.tag.rsplit("}", 1)[-1] == "sheet"
                and "name" in sheet.attrib and "sheetId" in sheet.attrib}
    except (OSError, KeyError, zipfile.BadZipFile, ElementTree.ParseError):
        return {}


def identity(sheet: Any) -> tuple[str, str] | None:
    try:
        name = sheet.Name
        if not isinstance(name, str) or not name:
            return None
        try:
            code = sheet.CodeName
        except Exception:
            code = None
        if isinstance(code, str) and code:
            return "code:" + code, name
        path = source(sheet)
        if path:
            try:
                file_version = version(path)
                sheet_id = _xml_sheet_ids(path, file_version["mtime_ns"], file_version["size"]).get(name)
                if sheet_id:
                    return "xml:" + sheet_id, name
            except OSError:
                pass
        return "name:" + name, name
    except Exception:
        return None


def settings(sheet: Any) -> dict | None:
    try:
        setup = sheet.PageSetup
        orientation = setup.Orientation
        if type(orientation) not in (int, float) or orientation not in (PORTRAIT, LANDSCAPE):
            return None
        result = {"Orientation": int(orientation)}
        for prop in SETTINGS:
            value = getattr(setup, prop)
            if value is not None and type(value) not in (str, int, float, bool):
                return None
            if isinstance(value, float) and not math.isfinite(value):
                return None
            result[prop] = round(value, 4) if type(value) is float else value
        return result
    except Exception:
        return None


@contextmanager
def orientation(sheet: Any, target: int):
    setup = sheet.PageSetup
    original = setup.Orientation
    try:
        if original != target:
            setup.Orientation = target
        if setup.Orientation != target:
            raise ValueError("Excel 未应用指定的打印方向")
        yield
    finally:
        if setup.Orientation != original:
            setup.Orientation = original
        if setup.Orientation != original:
            raise RuntimeError("无法恢复原打印方向，已停止处理，避免保存临时设置")


def record_for(sheet: Any, path: str | None = None) -> dict | None:
    path = path or source(sheet)
    key = identity(sheet)
    return _records().get(_path(path), {}).get(key[0]) if path and key else None


@contextmanager
def applied(sheet: Any):
    record = None if getattr(_analysis, "active", False) else record_for(sheet)
    if record and (record.get("stale") or record.get("rule_version") != RULE_VERSION):
        record = None
    if record:
        try:
            if record.get("version") != version(source(sheet)):
                record = None
        except (OSError, TypeError):
            record = None
    if record and record.get("selected") in (PORTRAIT, LANDSCAPE):
        with orientation(sheet, record["selected"]):
            yield
    else:
        yield


def with_layout(function):
    @wraps(function)
    def run(sheet, *args, **kwargs):
        with applied(sheet):
            return function(sheet, *args, **kwargs)
    return run


def token(path: str, name: str) -> str:
    return "|".join(sorted(r.get("revision", "") for r in _records().get(_path(path), {}).values()
                           if r.get("sheet") == name))


def _percentile(fonts: list[tuple[float, int]]) -> float:
    total = sum(weight for _, weight in fonts)
    accumulated = 0
    for size, weight in sorted(fonts):
        accumulated += weight
        if accumulated >= total * .1:
            return round(size, 2)
    return 0.0


def inspect_pdf(sheet: Any, pdf: Path, cells: list, target: int) -> dict:
    # 延后导入，read_breaks 自身也使用 applied；分析时使用未绑定的副本。
    from core import page_breaks
    fonts = []
    visible = set()
    borders = set()
    shapes = 0
    top = float(sheet.PageSetup.TopMargin or 0)
    bottom = float(sheet.PageSetup.BottomMargin or 0)
    with fitz.open(pdf) as doc:
        if not len(doc):
            raise ValueError("导出的 PDF 没有页面")
        for page in doc:
            if abs(page.rect.width - page.rect.height) > 1 and (page.rect.width > page.rect.height) != (target == LANDSCAPE):
                raise ValueError("导出 PDF 的页面方向与候选方向不一致")
            for block in page.get_text("dict").get("blocks", []):
                if block.get("type") == 1:
                    shapes += 1
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        if span["bbox"][1] < top - 2 or span["bbox"][3] > page.rect.height - bottom + 2:
                            continue  # 页眉、页脚的字号与页码不作为正文比较证据。
                        text = normalize_text(span.get("text", ""))
                        if text:
                            fonts.append((float(span["size"]), len(text)))
            borders.update(normalize_text(it["text"]) for it in find_clipped_spans_on_page(page))
        pieces = collect_text_pieces(doc)
        for piece in pieces:
            if piece.bbox[1] < top - 2 or piece.bbox[3] > doc[piece.page].rect.height - bottom + 2:
                continue
            visible.update(normalize_text("".join(char for char, _, shown in piece.chars if shown)))
        probe = CellProbe(sheet)
        hidden = {(it["row"], it["col"], it["reason"])
                  for it in find_hidden_cells(cells, pieces, probe.text, probe.printed)}
        breaks = page_breaks.read_breaks(sheet)
        cuts = page_breaks.find_cut_blocks(sheet, breaks, probe.printed) if breaks is not None else []
        columns = page_breaks._columns(sheet)
        across = (sum(columns[0] < c <= columns[1] for c in breaks["v"])
                  if breaks is not None and columns is not None else None)
        return {"pages": len(doc), "font_pt": _percentile(fonts), "visible": sorted(visible),
                "hidden": sorted([list(key) for key in hidden]), "borders": sorted(borders),
                "cuts": sorted(block["addr"] for block in cuts), "across": across, "images": shapes}


def choose(original: int, metrics: dict[int, dict]) -> tuple[int, str]:
    other = LANDSCAPE if original == PORTRAIT else PORTRAIT
    base, candidate = metrics[original], metrics[other]
    if not base["visible"] or not candidate["visible"]:
        return original, "可提取文字不足，保留原方向"
    # 逐项防止损失；不能用“错误总数变少”掩盖另一格新增的问题。
    if (not set(base["visible"]).issubset(candidate["visible"])
            or (base["images"] and not candidate["images"])
            or not {tuple(x) for x in candidate["hidden"]}.issubset(tuple(x) for x in base["hidden"])
            or not set(candidate["borders"]).issubset(base["borders"])
            or not set(candidate["cuts"]).issubset(base["cuts"])):
        return original, "另一方向出现内容减少或新增切字、压线、合并区域跨页，保留原方向"
    font, previous_font = candidate["font_pt"], base["font_pt"]
    if font < min(MIN_FONT_PT, previous_font) - .1 or font < previous_font - max(.5, previous_font * .05):
        return original, "另一方向实际字大小变小，保留原方向"
    if candidate["across"] is None or base["across"] is None:
        return original, "无法可靠读取分页，保留原方向"
    if len(candidate["hidden"]) < len(base["hidden"]) or len(candidate["borders"]) < len(base["borders"]):
        return other, "文字完整性或压线情况改善，且未发现新增问题"
    if font >= previous_font + max(.75, previous_font * .1) and candidate["across"] <= base["across"]:
        return other, "实际字大小改善，且横向拆页未增加"
    if candidate["across"] < base["across"]:
        return other, "减少横向拆页，且文字可读性未明显降低"
    if candidate["across"] > base["across"]:
        return original, "另一方向增加横向拆页，保留原方向"
    if len(candidate["cuts"]) < len(base["cuts"]):
        return other, "减少合并区域跨页"
    if candidate["pages"] < base["pages"] and len(candidate["cuts"]) <= len(base["cuts"]):
        return other, "内容与可读性相近，总页数更少"
    return original, "两种方向没有明显改善，保留原方向"


def prepare(sheet: Any, path: str, records: dict, export, output: Path | None = None) -> tuple[dict | None, bool]:
    """在未修改内容的源工作表或副本上分析；返回 (记录, 是否已生成选定 PDF)。"""
    key, initial = identity(sheet), settings(sheet)
    if key is None or initial is None:
        return None, False
    path = _path(path)
    before = version(path)
    existing = records.get(path, {}).get(key[0])
    if (existing and existing.get("version") == before and existing.get("settings") == initial
            and existing.get("rule_version") == RULE_VERSION and not existing.get("stale")):
        existing["sheet"] = key[1]
        return existing, False
    # 人工指定在内容修改后重新验证打印设置和文件版本，不覆盖用户选择。
    manual = existing.get("manual") if existing else None
    if existing and key[0].startswith("name:") and existing.get("version") != before:
        manual = None  # 旧格式无稳定 Sheet 身份时，不把人工设置套到可能被替换的同名表。
    record = {"sheet": key[1], "identity": key[0], "original": initial["Orientation"],
              "selected": manual or initial["Orientation"], "settings": initial, "version": before,
              "rule_version": RULE_VERSION, "revision": uuid4().hex, "manual": manual,
              "reason": "人工指定打印方向" if manual else "", "metrics": {}, "stale": False}
    made_pdf = False
    if not manual:
        # 分页读取不能把旧方案套回正在比较的候选方向。
        analyzing = getattr(_analysis, "active", False)
        _analysis.active = True
        try:
            with tempfile.TemporaryDirectory(prefix="sheet_orientation_") as temporary:
                cells = snapshot_cells(extract_sheet_snapshot(sheet))
                pdfs, metrics = {}, {}
                for target in (initial["Orientation"], LANDSCAPE if initial["Orientation"] == PORTRAIT else PORTRAIT):
                    pdf = Path(temporary) / f"{target}.pdf"
                    with orientation(sheet, target):
                        export(sheet, pdf)
                        metrics[target] = inspect_pdf(sheet, pdf, cells, target)
                    pdfs[target] = pdf
                record["selected"], record["reason"] = choose(initial["Orientation"], metrics)
                record["metrics"] = {str(k): {field: value for field, value in m.items() if field != "visible"}
                                     for k, m in metrics.items()}
                if output is not None:
                    shutil.copyfile(pdfs[record["selected"]], output)
                    made_pdf = True
        except Exception as error:
            # 方向恢复失败不能降级继续，否则后续保存可能带入临时方向。
            if settings(sheet) != initial:
                raise RuntimeError("试算后打印设置未恢复，已停止处理") from error
            record["reason"] = f"方向分析未完成，沿用原方向：{error}"
            record["fallback"] = True
        finally:
            _analysis.active = analyzing
    if version(path) != before:
        raise ValueError("方向分析期间文件已修改，请保存完成后重新分析")
    records.setdefault(path, {})[key[0]] = record
    return record, made_pdf


def followed_save(records: dict, path: str, before: dict | None) -> None:
    if before is None:
        return
    entries = records.get(_path(path), {})
    if not entries:
        return
    old_version = {k: before[k] for k in ("mtime_ns", "size")}
    after = version(path)
    for record in entries.values():
        if record.get("version") == old_version:
            record["version"] = after


def rename(records: dict, old: str, new: str, names: dict[str, str]) -> None:
    old, new = _path(old), _path(new)
    entries = records.pop(old, {}) if old != new else records.get(old, {})
    if not entries:
        return
    for key, record in list(entries.items()):
        name = record["sheet"]
        if name in names:
            record["sheet"] = names[name]
            if key.startswith("name:"):
                del entries[key]
                record["identity"] = "name:" + names[name]
                entries[record["identity"]] = record
        # 打印区域中可能包含表名，改名后交给 prepare 重新验证。
    records[new] = entries


def prune(records: dict, path: str, keys: list) -> None:
    """删除已不存在的 Sheet 记录；隐藏或被过滤的表仍保留记录。"""
    if not all(keys):
        return  # 身份读取不完整时不猜测删除。
    valid = dict(keys)
    names = set(valid.values())
    entries = records.get(_path(path), {})
    for key in list(entries):
        if key not in valid or (entries[key]["sheet"] != valid[key] and entries[key]["sheet"] in names):
            del entries[key]


def validate_records(records: Any) -> bool:
    if not isinstance(records, dict):
        return False
    for path, entries in records.items():
        if not isinstance(path, str) or not isinstance(entries, dict):
            return False
        for key, record in entries.items():
            if (not isinstance(key, str) or not isinstance(record, dict)
                    or not isinstance(record.get("sheet"), str)
                    or type(record.get("selected")) is not int or record["selected"] not in (PORTRAIT, LANDSCAPE)
                    or type(record.get("original")) is not int or record["original"] not in (PORTRAIT, LANDSCAPE)
                    or not isinstance(record.get("settings"), dict)
                    or not isinstance(record.get("version"), dict)
                    or (record.get("manual") is not None and
                        (type(record["manual"]) is not int or record["manual"] not in (PORTRAIT, LANDSCAPE)))
                    or type(record.get("rule_version")) is not int
                    or not isinstance(record.get("revision"), str)
                    or not isinstance(record.get("identity"), str)
                    or any(type(record["version"].get(field)) is not int or record["version"][field] < 0
                           for field in ("mtime_ns", "size"))
                    or not isinstance(record.get("metrics", {}), dict)):
                return False
            for direction, m in record.get("metrics", {}).items():
                if (direction not in ("1", "2") or not isinstance(m, dict)
                        or type(m.get("pages")) is not int or m["pages"] < 1
                        or type(m.get("font_pt")) not in (int, float) or not math.isfinite(m["font_pt"])
                        or m["font_pt"] < 0
                        or (m.get("across") is not None and (type(m["across"]) is not int or m["across"] < 0))
                        or any(not isinstance(m.get(field), list) for field in ("hidden", "borders", "cuts"))):
                    return False
                if any(not isinstance(x, str) for field in ("borders", "cuts") for x in m[field]):
                    return False
                if any(not isinstance(x, list) or len(x) != 3 or type(x[0]) is not int
                       or type(x[1]) is not int or not isinstance(x[2], str) for x in m["hidden"]):
                    return False
    return True
