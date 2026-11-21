"""设置图标（用户 2026-11-20 要）：拖进透明底 PNG → 生成 Windows 图标（.ico），存进图标库，可以套到 Windows 文件夹上。

- 图标库放在数据文件夹里（桌面\\handle-excel\\程序数据\\图标库，见 core/data_dir.py），换 ZIP 版本不丢：
    图标库\\index.json        名字、分类、套到了哪些文件夹、工具箱里哪个工具用哪个图标
    图标库\\icons\\<id>.ico    多尺寸图标（16/20/24/32/40/48/64/96/128/256）
    图标库\\icons\\<id>.png    256×256 预览（网页显示用）
    图标库\\icons\\<id>_orig.png  拖进来的原图，原样留着
- 生成 .ico 只用 PyMuPDF（本来就要装），不另外装 Pillow：每个尺寸从原图直接缩，居中放进正方形；
  256 存 PNG，其余存 32 位 BMP（老一点的资源管理器也认）。
- 套到文件夹：把 .ico 复制进文件夹（隐藏 + 系统文件，文件名带版本号，换图后资源管理器不会一直显示旧图），
  写 desktop.ini 的 IconResource，desktop.ini 设隐藏 + 系统，文件夹设只读（Windows 认 desktop.ini 的条件），
  再通知资源管理器刷新。文件夹里的是复制件：以后删掉图标库里的图标，文件夹的图标也不会坏。
- 替换图标图片时，已经套过这个图标的文件夹一起换成新图。
- 系统图标（用户 2026-11-21 定）：网站自带的图标（static/images/harbor/sprite 和 tools）自动放进图标库，
  分类固定叫「系统图标」，只能用、不能改名 / 改分类 / 替换 / 删除；「系统图标」这个分类名用户不能用，
  用户上传的图标分类名自己定。
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import struct
import sys
import threading
import time
import uuid
from pathlib import Path

from core import data_dir
from core.common import console_print

logger = logging.getLogger(__name__)

SIZES = (256, 128, 96, 64, 48, 40, 32, 24, 20, 16)
MAX_BYTES = 30 * 1024 * 1024
ICON_PREFIX = "_folder_icon_"   # 复制进文件夹的图标文件名前缀（后面跟 图标id_版本.ico）
MAX_NAME = 60
MAX_CATEGORY = 30
_PIXEL_FINGERPRINT = "pixels_v1"

LIB_DIR: Path | None = None   # 测试时改这里；平时用数据文件夹下的「图标库」
_lock = threading.RLock()
_fingerprint_cache: dict[Path, dict] = {}


SYSTEM_CATEGORY = "系统图标"
_HARBOR = Path(__file__).resolve().parent.parent / "static" / "images" / "harbor"
# 网站自带的图标：(文件, 名字)。主页精灵图按格子切成 sprite/<名字>.png，工具专属图在 tools/
SYSTEM_ICONS = (
    ("sprite/home.png", "主页"), ("sprite/grid.png", "表格"), ("sprite/search.png", "放大镜"),
    ("sprite/chart.png", "柱状图"), ("sprite/folder.png", "文件夹"), ("sprite/layers.png", "文件叠"),
    ("sprite/download.png", "下载"), ("sprite/upload.png", "上传"), ("sprite/settings.png", "齿轮"),
    ("sprite/link.png", "链条"), ("sprite/eye.png", "眼睛"), ("sprite/palette.png", "调色板"),
    ("sprite/target.png", "靶心"), ("sprite/document.png", "文档"), ("sprite/type.png", "字体"),
    ("sprite/image.png", "图片"), ("sprite/check.png", "勾选清单"), ("sprite/clean.png", "清理表格"),
    ("sprite/adjust.png", "尺子"), ("sprite/cells.png", "查格子"), ("sprite/crate.png", "上传货箱"),
    ("sprite/new.png", "宝箱"), ("sprite/finish.png", "奖杯"), ("sprite/map.png", "地图"),
    ("sprite/compass.png", "罗盘"),
    ("tools/bl.png", "断链警示"), ("tools/reflink.png", "链条扳手"), ("tools/widths.png", "列宽"),
    ("tools/pdfstack.png", "PDF 一叠"), ("tools/pdfdiff.png", "PDF 对比"), ("tools/report.png", "对比报告"),
    ("tools/json.png", "JSON 表格"), ("tools/pick.png", "挑文件"), ("tools/view.png", "视图"),
    ("tools/appicon.png", "应用图标"), ("tools/imgfix.png", "修复图片"), ("tools/ticket.png", "船票"),
)


class DuplicateIconError(ValueError):
    """图片内容与图标库中的条目重复。"""


def _log(message: str, **fields) -> None:
    """调试日志：控制台 + log.txt 都能看到，统一以「[设置图标]」开头"""
    detail = " ".join(f"{k}={v}" for k, v in fields.items())
    logger.info("[设置图标] %s %s", message, detail)
    try:
        from core import debug_log
        debug_log.event(f"🖼️ 设置图标：{message}", **fields)
    except Exception:  # noqa: BLE001 — 调试日志写不了也不影响功能
        pass


def library_dir() -> Path:
    return LIB_DIR if LIB_DIR is not None else data_dir.DATA_DIR / "图标库"


def _icons_dir() -> Path:
    return library_dir() / "icons"


def _index_path() -> Path:
    return library_dir() / "index.json"


# ---------------------------------------------------------------- 索引

def _load() -> dict:
    path = _index_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        data = {}
    except (OSError, ValueError) as err:
        # 读不懂就先备份一份再从空的开始，别把用户的图标文件弄丢（图标文件本身还在 icons 里）
        backup = path.with_name(f"index.broken-{int(time.time())}.json")
        try:
            path.replace(backup)
        except OSError:
            pass
        _log("index.json 读不懂，已备份后重建", 原因=err, 备份=backup)
        data = {}
    icons = [i for i in data.get("icons", []) if isinstance(i, dict) and i.get("id")]
    tools = {str(k): str(v) for k, v in (data.get("tools") or {}).items() if v}
    return {"version": 1, "icons": icons, "tools": tools}


def _save(data: dict) -> None:
    path = _index_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)


def _find(data: dict, icon_id: str) -> dict:
    for icon in data["icons"]:
        if icon["id"] == icon_id:
            return icon
    raise ValueError("图标库里没有这个图标（可能已经删掉了），请刷新页面")


def _clean_text(value, limit: int) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text[:limit]


def _public(icon: dict) -> dict:
    folders = icon.get("folders") or []
    return {
        "id": icon["id"], "name": icon.get("name") or "未命名", "category": icon.get("category") or "",
        "width": icon.get("width"), "height": icon.get("height"), "transparent": icon.get("transparent", True),
        "render_width": icon.get("render_width"), "render_height": icon.get("render_height"),
        "created": icon.get("created"), "updated": icon.get("updated"), "version": icon.get("version", 1),
        "source": icon.get("source") or "", "folders": folders, "system": bool(icon.get("system")),
    }


def listing() -> dict:
    with _lock:
        data = _load()
        if _ensure_system(data):
            _save(data)
        groups, unchecked = _duplicate_groups(data)
    icons = [_public(i) for i in data["icons"]]
    categories: dict[str, int] = {}
    for icon in icons:
        categories[icon["category"]] = categories.get(icon["category"], 0) + 1
    return {
        "icons": icons,
        "categories": [{"name": k, "count": v} for k, v in sorted(categories.items(), key=lambda kv: (kv[0] != SYSTEM_CATEGORY, kv[0] == "", kv[0]))],
        "system_category": SYSTEM_CATEGORY,
        "tools": data["tools"],
        "dir": str(library_dir()),
        "windows": sys.platform == "win32",
        "duplicates": groups,
        "duplicate_unchecked": unchecked,
    }


# ---------------------------------------------------------------- 系统图标

def _system_id(rel: str) -> str:
    return hashlib.sha1(f"system:{rel}".encode()).hexdigest()[:12]


def _ensure_system(data: dict) -> bool:
    """网站自带的图标补进图标库（缺的才生成 .ico，图片换了按文件大小/时间重出）；返回有没有改动。"""
    wanted = {_system_id(rel): (rel, name) for rel, name in SYSTEM_ICONS}
    have = {i["id"]: i for i in data["icons"] if i.get("system")}
    changed = False
    for icon_id in set(have) - set(wanted):   # 新版本里不再自带的，从库里拿掉（文件留着不碍事）
        data["icons"] = [i for i in data["icons"] if i["id"] != icon_id]
        changed = True
    system_icons = []
    for icon_id, (rel, name) in wanted.items():
        path = _HARBOR / rel
        icon = have.get(icon_id)
        try:
            stat = path.stat()
            stamp = [stat.st_size, int(stat.st_mtime)]
            ready = icon and icon.get("stamp") == stamp and (_icons_dir() / f"{icon_id}.ico").is_file()
            if not ready:
                body = path.read_bytes()
                info = _render(icon_id, body, _load_png(body), trim=True)
                now = time.strftime("%Y-%m-%d %H:%M:%S")
                icon = {**(icon or {"created": now, "folders": [], "version": 0}),
                        "id": icon_id, "width": info["width"], "height": info["height"], "transparent": info["transparent"],
                        "render_width": info["render_width"], "render_height": info["render_height"],
                        "updated": now, "source": rel, "stamp": stamp}
                icon["version"] = int(icon.get("version", 0)) + 1
                icon.pop("fingerprints", None)
                changed = True
        except (OSError, ValueError) as err:
            _log("系统图标生成不了", 文件=rel, 原因=err)
            if not icon:
                continue
        if icon.get("name") != name or icon.get("category") != SYSTEM_CATEGORY or not icon.get("system"):
            icon.update(name=name, category=SYSTEM_CATEGORY, system=True)
            changed = True
        system_icons.append(icon)
    if changed:
        users = [i for i in data["icons"] if not i.get("system")]
        data["icons"] = users + system_icons   # 系统图标排在用户图标后面
        _log("系统图标已更新", 数量=len(system_icons))
    return changed


def _check_user_icon(icon: dict, action: str) -> None:
    if icon.get("system"):
        raise ValueError(f"「{icon.get('name')}」是系统图标，只能用，不能{action}")


def _check_category(category: str) -> str:
    if category == SYSTEM_CATEGORY:
        raise ValueError(f"「{SYSTEM_CATEGORY}」是网站自带图标的分类，请给上传的图标起别的分类名")
    return category


# ---------------------------------------------------------------- PNG → ICO

def _fitz():
    from core.pdf_detector import fitz  # 统一从这里导入 PyMuPDF（别直接 import fitz）
    return fitz


def _validate_image_bytes(body: bytes) -> None:
    if not body:
        raise ValueError("这是空文件")
    if len(body) > MAX_BYTES:
        raise ValueError("图片超过 30 MB，太大了")


def _load_png(body: bytes):
    """读 PNG（JPG/BMP 等 PyMuPDF 认的也行），统一转成 RGBA"""
    _validate_image_bytes(body)
    fitz = _fitz()
    try:
        pix = fitz.Pixmap(body)
    except Exception as err:  # noqa: BLE001 — PyMuPDF 各版本报的错不一样
        raise ValueError(f"读不了这张图片，请拖 PNG 图片（{err}）") from None
    if pix.width < 1 or pix.height < 1:
        raise ValueError("图片是空的")
    if pix.width * pix.height > 40_000_000:
        raise ValueError(f"图片太大（{pix.width}×{pix.height}），请先缩小到 4000×4000 以内")
    if pix.colorspace is None or pix.colorspace.n != 3:
        pix = fitz.Pixmap(fitz.csRGB, pix)
    if not pix.alpha:
        pix = fitz.Pixmap(pix, 1)  # 加不透明的 alpha
    return pix


def _alpha(pix) -> bytes:
    return bytes(pix.samples)[3::4]


def _content_box(pix) -> tuple[int, int, int, int]:
    """按可见像素找范围；忽略近乎透明的杂点，不改范围内的半透明像素。"""
    w, h = pix.width, pix.height
    alpha = _alpha(pix)
    # 低于 8/255 的零星像素肉眼几乎不可见，却可能延伸到原图边缘。
    # 只用于确定裁剪范围；如果整张图都很淡，则回退到非零 alpha，避免裁掉内容。
    visible = alpha.translate(bytes(0 if i < 8 else 1 for i in range(256)))
    if not visible.strip(b"\x00"):
        visible = alpha
    x0, y0, x1, y1 = w, h, -1, -1
    for y in range(h):
        row = visible[y * w:(y + 1) * w]
        left = w - len(row.lstrip(b"\x00"))
        if left == w:
            continue
        right = len(row.rstrip(b"\x00"))
        y0 = min(y0, y)
        y1 = y
        x0 = min(x0, left)
        x1 = max(x1, right - 1)
    if y1 < 0:
        return 0, 0, w - 1, h - 1
    return x0, y0, x1, y1


def _crop(pix, box):
    fitz = _fitz()
    x0, y0, x1, y1 = box
    if (x0, y0, x1, y1) == (0, 0, pix.width - 1, pix.height - 1):
        return pix
    out = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, x1 - x0 + 1, y1 - y0 + 1), 1)
    out.clear_with()  # 不带参数 = 全透明（clear_with(0) 是不透明黑）
    src = fitz.Pixmap(pix)  # 复制一份再挪原点，别改到调用方的
    src.set_origin(-x0, -y0)
    out.copy(src, out.irect)
    return out


def _square(pix, size: int):
    """缩到 size×size 的正方形里（长边贴满，居中，四周透明）"""
    fitz = _fitz()
    w, h = pix.width, pix.height
    k = size / max(w, h)
    sw, sh = max(1, round(w * k)), max(1, round(h * k))
    scaled = pix if (sw, sh) == (w, h) else fitz.Pixmap(pix, sw, sh, None)
    canvas = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, size, size), 1)
    canvas.clear_with()  # 不带参数 = 全透明
    scaled = fitz.Pixmap(scaled)
    scaled.set_origin((size - sw) // 2, (size - sh) // 2)
    canvas.copy(scaled, scaled.irect)
    return canvas


def _fingerprint(pix) -> str:
    """比较裁掉透明边后的真实像素，不受文件名、PNG 元数据影响。"""
    pix = _crop(pix, _content_box(pix))
    digest = hashlib.sha256(struct.pack(">II", pix.width, pix.height))
    digest.update(pix.samples)
    return digest.hexdigest()


def _original_path(icon_id: str) -> Path:
    return _icons_dir() / f"{icon_id}_orig.png"


def _original_stamp(path: Path) -> list[int]:
    stat = path.stat()
    return [stat.st_mtime_ns, stat.st_size]


def _record_fingerprints(icon: dict, original: str, pixels: str) -> None:
    icon["fingerprints"] = {
        "file": _original_stamp(_original_path(icon["id"])),
        "original": original,
        _PIXEL_FINGERPRINT: pixels,
    }


def _icon_fingerprint(icon: dict, kind: str) -> str:
    """旧条目按需补算指纹。原文件变动时重新计算。"""
    path = _original_path(icon["id"])
    stamp = _original_stamp(path)
    fingerprints = icon.get("fingerprints")
    if not isinstance(fingerprints, dict) or fingerprints.get("file") != stamp:
        fingerprints = _fingerprint_cache.get(path, {})
    if fingerprints.get("file") != stamp:
        fingerprints = {"file": stamp}
    if not fingerprints.get(kind):
        body = path.read_bytes()
        fingerprints["original"] = hashlib.sha256(body).hexdigest()
        if kind == _PIXEL_FINGERPRINT:
            fingerprints[kind] = _fingerprint(_load_png(body))
        _fingerprint_cache[path] = fingerprints
    icon["fingerprints"] = fingerprints
    return fingerprints[kind]


def _duplicate_groups(data: dict) -> tuple[list[list[str]], list[str]]:
    hashes: dict[str, list[str]] = {}
    unchecked = []
    for icon in data["icons"]:
        try:
            hashes.setdefault(_icon_fingerprint(icon, _PIXEL_FINGERPRINT), []).append(icon["id"])
        except (OSError, ValueError):
            unchecked.append(icon["id"])
    return [ids for ids in hashes.values() if len(ids) > 1], unchecked


def _check_duplicate(data: dict, fingerprint: str, kind: str, exclude_id: str = "") -> None:
    for icon in data["icons"]:
        if icon["id"] == exclude_id:
            continue
        try:
            existing = _icon_fingerprint(icon, kind)
        except (OSError, ValueError):
            continue
        if existing == fingerprint:
            raise DuplicateIconError(f"与图标库里的「{icon.get('name') or '未命名'}」（{icon.get('category') or '未分类'}）重复")


def _bmp_entry(pix) -> bytes:
    """32 位 BMP（ICO 里的 DIB：高度写两倍、自下而上、BGRA，后面跟 1 位 AND 遮罩）"""
    w, h = pix.width, pix.height
    samples = bytes(pix.samples)  # PyMuPDF 的 RGBA 是预乘过 alpha 的，这里换回直通 alpha
    rows = []
    for y in range(h - 1, -1, -1):
        row = bytearray(w * 4)
        base = y * w * 4
        for x in range(w):
            r, g, b, a = samples[base + x * 4: base + x * 4 + 4]
            if a and a < 255:
                r, g, b = min(255, r * 255 // a), min(255, g * 255 // a), min(255, b * 255 // a)
            elif not a:
                r = g = b = 0
            row[x * 4: x * 4 + 4] = bytes((b, g, r, a))
        rows.append(bytes(row))
    mask_row = ((w + 31) // 32) * 4
    mask = b"\x00" * (mask_row * h)  # 透明度看 alpha，遮罩全 0
    header = struct.pack("<IiiHHIIiiII", 40, w, h * 2, 1, 32, 0, len(rows) * w * 4 + len(mask), 0, 0, 0, 0)
    return header + b"".join(rows) + mask


def build_ico(pix) -> bytes:
    images = []
    for size in SIZES:
        square = _square(pix, size)
        data = square.tobytes("png") if size >= 256 else _bmp_entry(square)
        images.append((size, data))
    head = struct.pack("<HHH", 0, 1, len(images))
    offset = 6 + 16 * len(images)
    entries, blobs = [], []
    for size, data in images:
        dim = 0 if size >= 256 else size
        entries.append(struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, len(data), offset))
        blobs.append(data)
        offset += len(data)
    return head + b"".join(entries) + b"".join(blobs)


def _render(icon_id: str, body: bytes, pix, trim: bool) -> dict:
    alpha = _alpha(pix)
    transparent = alpha.count(255) != len(alpha)
    width, height = pix.width, pix.height
    if trim and transparent:
        pix = _crop(pix, _content_box(pix))
    folder = _icons_dir()
    folder.mkdir(parents=True, exist_ok=True)
    ico = build_ico(pix)
    (folder / f"{icon_id}.ico").write_bytes(ico)
    (folder / f"{icon_id}.png").write_bytes(_square(pix, 256).tobytes("png"))
    _original_path(icon_id).write_bytes(body)
    return {
        "width": width, "height": height, "transparent": transparent,
        "render_width": pix.width, "render_height": pix.height,
        "trimmed": (pix.width, pix.height) != (width, height),
        "small": max(pix.width, pix.height) < 128, "ico_bytes": len(ico),
    }


def _name_from_file(filename: str) -> str:
    stem = Path(filename or "").stem
    return _clean_text(stem, MAX_NAME) or "未命名"


def add(filename: str, body: bytes, category: str = "", trim: bool = True) -> dict:
    _validate_image_bytes(body)
    category = _check_category(_clean_text(category, MAX_CATEGORY))
    original = hashlib.sha256(body).hexdigest()
    with _lock:
        data = _load()
        _check_duplicate(data, original, "original")
        pix = _load_png(body)
        pixels = _fingerprint(pix)
        _check_duplicate(data, pixels, _PIXEL_FINGERPRINT)
        icon_id = uuid.uuid4().hex[:12]
        info = _render(icon_id, body, pix, trim)
        now = time.strftime("%Y-%m-%d %H:%M:%S")
        icon = {"id": icon_id, "name": _name_from_file(filename), "category": category,
                "width": info["width"], "height": info["height"], "transparent": info["transparent"],
                "render_width": info["render_width"], "render_height": info["render_height"],
                "created": now, "updated": now, "version": 1, "source": Path(filename or "").name, "folders": []}
        _record_fingerprints(icon, original, pixels)
        data["icons"].insert(0, icon)
        _save(data)
    _log("加进图标库", 名字=icon["name"], 分类=icon["category"] or "未分类", 尺寸=f"{info['width']}x{info['height']}",
         透明=info["transparent"], 裁边=info["trimmed"], ico大小=info["ico_bytes"])
    return {"icon": _public(icon), "info": info}


def replace(icon_id: str, filename: str, body: bytes, trim: bool = True) -> dict:
    with _lock:
        data = _load()
        icon = _find(data, icon_id)
        _check_user_icon(icon, "替换图片")
        pix = _load_png(body)
        pixels = _fingerprint(pix)
        try:
            unchanged = pixels == _icon_fingerprint(icon, _PIXEL_FINGERPRINT)
        except (OSError, ValueError):
            unchanged = False
        if not unchanged:
            _check_duplicate(data, pixels, _PIXEL_FINGERPRINT, exclude_id=icon_id)
        info = _render(icon_id, body, pix, trim)
        icon.update(width=info["width"], height=info["height"], transparent=info["transparent"],
                    render_width=info["render_width"], render_height=info["render_height"],
                    updated=time.strftime("%Y-%m-%d %H:%M:%S"), version=int(icon.get("version", 1)) + 1,
                    source=Path(filename or "").name)
        _record_fingerprints(icon, hashlib.sha256(body).hexdigest(), pixels)
        _save(data)
        folders = list(icon.get("folders") or [])
    _log("替换图片", 名字=icon["name"], 新图=filename, 版本=icon["version"], 已套文件夹=len(folders))
    # 已经套过这个图标的文件夹一起换成新图（文件夹里的 desktop.ini 已不是本工具的就不动）
    refreshed = []
    for folder in folders:
        try:
            if _folder_icon_id(Path(folder)) == icon_id:
                refreshed.append(_apply_one(icon, Path(folder)))
            else:
                refreshed.append({"folder": folder, "ok": False, "skipped": True, "error": "这个文件夹已经换成别的图标，没有动"})
        except (OSError, ValueError) as err:
            refreshed.append({"folder": folder, "ok": False, "error": str(err)})
    if refreshed:
        _forget_stale(icon_id, refreshed)
    return {"icon": _public(icon), "info": info, "folders": refreshed}


def retrim(icon_id: str) -> dict:
    """用保存的原图重新裁边，并沿用替换图片时的文件夹刷新规则。"""
    with _lock:
        icon = _find(_load(), icon_id)
        _check_user_icon(icon, "重新裁剪")
        try:
            body = _original_path(icon_id).read_bytes()
        except OSError as err:
            raise ValueError(f"找不到保存的原图，无法重新裁剪（{err}）") from None
        return replace(icon_id, icon.get("source") or "image.png", body, trim=True)


def update(icon_id: str, name=None, category=None) -> dict:
    with _lock:
        data = _load()
        icon = _find(data, icon_id)
        _check_user_icon(icon, "改名字和分类")
        old = (icon.get("name"), icon.get("category"))
        if name is not None:
            icon["name"] = _clean_text(name, MAX_NAME) or icon.get("name") or "未命名"
        if category is not None:
            icon["category"] = _check_category(_clean_text(category, MAX_CATEGORY))
        _save(data)
    _log("改名字 / 分类", 原来=f"{old[0]} / {old[1] or '未分类'}", 现在=f"{icon['name']} / {icon['category'] or '未分类'}")
    return {"icon": _public(icon)}


def rename_category(old: str, new: str) -> dict:
    old, new = _clean_text(old, MAX_CATEGORY), _clean_text(new, MAX_CATEGORY)
    if old == SYSTEM_CATEGORY:
        raise ValueError(f"「{SYSTEM_CATEGORY}」是固定分类，不能改名")
    _check_category(new)
    with _lock:
        data = _load()
        count = 0
        for icon in data["icons"]:
            if not icon.get("system") and (icon.get("category") or "") == old:
                icon["category"] = new
                count += 1
        _save(data)
    _log("分类改名", 原来=old or "未分类", 现在=new or "未分类", 图标数=count)
    return {"count": count}


def delete(icon_id: str) -> dict:
    with _lock:
        data = _load()
        icon = _find(data, icon_id)
        _check_user_icon(icon, "删除")
        data["icons"] = [i for i in data["icons"] if i["id"] != icon_id]
        tools = [k for k, v in data["tools"].items() if v == icon_id]
        for key in tools:
            data["tools"].pop(key, None)
        _save(data)
    for suffix in (".ico", ".png", "_orig.png"):
        try:
            (_icons_dir() / f"{icon_id}{suffix}").unlink()
        except FileNotFoundError:
            pass
    _log("删除图标", 名字=icon.get("name"), 已套文件夹=len(icon.get("folders") or []), 工具恢复默认=",".join(tools) or "无")
    return {"folders": len(icon.get("folders") or []), "tools": tools}


def icon_file(icon_id: str, kind: str) -> tuple[bytes, str]:
    if not re.fullmatch(r"[0-9a-f]{6,32}", icon_id or ""):
        raise ValueError("图标编号不对")
    suffix = {"png": ".png", "ico": ".ico", "orig": "_orig.png"}.get(kind)
    if not suffix:
        raise ValueError("不认识的文件种类")
    with _lock:
        icon = _find(_load(), icon_id)
    body = (_icons_dir() / f"{icon_id}{suffix}").read_bytes()
    return body, icon.get("name") or "icon"


def set_tool_icon(tool_id: str, icon_id) -> dict:
    tool_id = _clean_text(tool_id, 40)
    if not re.fullmatch(r"[A-Za-z0-9_-]{1,40}", tool_id):
        raise ValueError("工具编号不对")
    with _lock:
        data = _load()
        if icon_id:
            _find(data, str(icon_id))
            data["tools"][tool_id] = str(icon_id)
        else:
            data["tools"].pop(tool_id, None)
        _save(data)
    _log("工具箱图标", 工具=tool_id, 图标=icon_id or "恢复默认")
    return {"tools": data["tools"]}


def open_library() -> dict:
    folder = library_dir()
    folder.mkdir(parents=True, exist_ok=True)
    if hasattr(os, "startfile"):
        os.startfile(str(folder))  # type: ignore[attr-defined]  # noqa: S606 — 只开图标库文件夹
    return {"dir": str(folder)}


# ---------------------------------------------------------------- 套到 Windows 文件夹
# 只给用到图标的地方调用（目录结构模板 / 船票建文件夹时），图标页本身不再套文件夹。

def parse_folders(value) -> list[str]:
    if isinstance(value, list):
        lines = [str(v) for v in value]
    else:
        lines = str(value or "").splitlines()
    seen, out = set(), []
    for line in lines:
        text = line.strip().strip('"').strip("'").strip()
        if not text:
            continue
        key = os.path.normcase(os.path.normpath(text))
        if key in seen:
            continue
        seen.add(key)
        out.append(text)
    if len(out) > 200:
        raise ValueError("一次最多 200 个文件夹")
    return out


def _attrs(path: Path, hidden_system: bool) -> None:
    """Windows：设 / 清 隐藏 + 系统属性（别的系统什么都不做）"""
    if sys.platform != "win32":
        return
    import ctypes
    FILE_ATTRIBUTE_NORMAL, HIDDEN, SYSTEM = 0x80, 0x2, 0x4
    ctypes.windll.kernel32.SetFileAttributesW(str(path), (HIDDEN | SYSTEM) if hidden_system else FILE_ATTRIBUTE_NORMAL)


def _make_system_folder(folder: Path, on: bool) -> None:
    if sys.platform != "win32":
        return
    import ctypes
    # PathMakeSystemFolderW / PathUnmakeSystemFolderW 是 shlwapi 的导出，shell32 里没有
    shlwapi = ctypes.windll.shlwapi
    (shlwapi.PathMakeSystemFolderW if on else shlwapi.PathUnmakeSystemFolderW)(str(folder))


def _notify(folder: Path) -> None:
    """叫资源管理器刷新这个文件夹的图标"""
    if sys.platform != "win32":
        return
    import ctypes
    SHCNE_UPDATEITEM, SHCNE_UPDATEDIR, SHCNF_PATHW = 0x2000, 0x1000, 0x0005
    shell32 = ctypes.windll.shell32
    shell32.SHChangeNotify(SHCNE_UPDATEITEM, SHCNF_PATHW, ctypes.c_wchar_p(str(folder)), None)
    shell32.SHChangeNotify(SHCNE_UPDATEDIR, SHCNF_PATHW, ctypes.c_wchar_p(str(folder.parent)), None)


def _read_ini(path: Path) -> tuple[list[str], str]:
    try:
        raw = path.read_bytes()
    except FileNotFoundError:
        return [], "ascii"
    if raw.startswith(b"\xff\xfe"):
        encoding, text = "utf-16", raw.decode("utf-16")
    elif raw.startswith(b"\xef\xbb\xbf"):
        encoding, text = "utf-8-sig", raw.decode("utf-8-sig")
    else:
        try:
            encoding, text = "utf-8", raw.decode("utf-8")
        except UnicodeDecodeError:
            encoding, text = "cp932", raw.decode("cp932", errors="replace")
    return text.splitlines(), encoding


def _write_ini(path: Path, lines: list[str], encoding: str) -> None:
    text = "\r\n".join(lines).rstrip("\r\n") + "\r\n"
    if encoding == "ascii" and not text.isascii():
        encoding = "utf-16"
    if path.exists():
        _attrs(path, False)  # 隐藏 + 系统的文件直接覆盖会被 Windows 拒绝，先清掉属性
    path.write_bytes(text.encode(encoding))
    _attrs(path, True)


_ICON_KEYS = re.compile(r"^\s*(IconResource|IconFile|IconIndex)\s*=", re.I)
_SECTION = re.compile(r"^\s*\[(.+?)\]\s*$")


def _set_icon_lines(lines: list[str], value: str | None) -> list[str]:
    """改 [.ShellClassInfo] 里的图标设置；value=None 就是去掉。别的内容原样留着"""
    out, in_shell, found = [], False, False
    for line in lines:
        section = _SECTION.match(line)
        if section:
            in_shell = section.group(1).strip().lower() == ".shellclassinfo"
            out.append(line)
            if in_shell and not found:
                found = True
                if value:
                    out.append(f"IconResource={value}")
            continue
        if in_shell and _ICON_KEYS.match(line):
            continue
        out.append(line)
    if not found and value:
        out = ["[.ShellClassInfo]", f"IconResource={value}"] + ([""] if out else []) + out
    return out


def _ini_is_empty(lines: list[str]) -> bool:
    for line in lines:
        text = line.strip()
        if not text or text.startswith(";") or _SECTION.match(text):
            continue
        return False
    return True


def _folder_icon_id(folder: Path) -> str:
    """这个文件夹现在用的是不是本工具套的图标，是的话返回图标 id"""
    lines, _ = _read_ini(folder / "desktop.ini")
    for line in lines:
        m = re.match(r"^\s*IconResource\s*=\s*" + re.escape(ICON_PREFIX) + r"([0-9a-f]+)_\d+\.ico", line, re.I)
        if m:
            return m.group(1)
    return ""


def _check_folder(text: str) -> Path:
    folder = Path(text).expanduser()
    if not folder.is_absolute():
        raise ValueError("请填完整路径（例如 D:\\Users\\…\\文件夹）")
    if not folder.exists():
        raise ValueError("找不到这个文件夹")
    if not folder.is_dir():
        raise ValueError("这是文件，不是文件夹")
    return folder


def _remove_our_icons(folder: Path, keep: str = "") -> None:
    for old in folder.glob(f"{ICON_PREFIX}*.ico"):
        if old.name == keep:
            continue
        _attrs(old, False)
        try:
            old.unlink()
        except OSError as err:
            _log("旧图标文件删不掉", 文件=old, 原因=err)


def _apply_one(icon: dict, folder: Path) -> dict:
    name = f"{ICON_PREFIX}{icon['id']}_{int(icon.get('version', 1))}.ico"
    target = folder / name
    body = (_icons_dir() / f"{icon['id']}.ico").read_bytes()
    if target.exists():
        _attrs(target, False)
    target.write_bytes(body)
    _attrs(target, True)
    ini = folder / "desktop.ini"
    lines, encoding = _read_ini(ini)
    _write_ini(ini, _set_icon_lines(lines, f"{name},0"), encoding)
    _remove_our_icons(folder, keep=name)
    _make_system_folder(folder, True)
    _notify(folder)
    return {"folder": str(folder), "ok": True}


def _forget_stale(icon_id: str, results: list[dict]) -> None:
    """替换后已经不用这个图标（或找不到）的文件夹，从记录里拿掉"""
    gone = {r["folder"] for r in results if not r.get("ok")}
    if not gone:
        return
    with _lock:
        data = _load()
        for icon in data["icons"]:
            if icon["id"] == icon_id:
                icon["folders"] = [f for f in icon.get("folders") or [] if f not in gone]
        _save(data)


def _record(folder: str, icon_id: str | None) -> None:
    """记下文件夹现在用哪个图标（一个文件夹只记在一个图标下）"""
    key = os.path.normcase(os.path.normpath(folder))
    with _lock:
        data = _load()
        for icon in data["icons"]:
            folders = [f for f in icon.get("folders") or [] if os.path.normcase(os.path.normpath(f)) != key]
            if icon["id"] == icon_id:
                folders.append(folder)
            icon["folders"] = folders
        _save(data)


def apply(icon_id: str, folders) -> dict:
    paths = parse_folders(folders)
    if not paths:
        raise ValueError("请填要换图标的文件夹路径（每行一个）")
    with _lock:
        icon = _find(_load(), icon_id)
    results = []
    for text in paths:
        try:
            folder = _check_folder(text)
            result = _apply_one(icon, folder)
            _record(str(folder), icon_id)
        except PermissionError as err:
            result = {"folder": text, "ok": False, "error": f"没有权限改这个文件夹（{err.strerror or err}）"}
        except (OSError, ValueError) as err:
            result = {"folder": text, "ok": False, "error": str(err)}
        results.append(result)
        _log("套到文件夹" if result["ok"] else "套到文件夹失败", 图标=icon.get("name"), 文件夹=text, 原因=result.get("error", ""))
    ok = sum(1 for r in results if r["ok"])
    if ok:
        console_print(f"🖼️ 已把图标「{icon.get('name')}」套到 {ok} 个文件夹")
    return {"results": results, "ok_count": ok, "windows": sys.platform == "win32"}


def restore(folders) -> dict:
    """文件夹恢复默认图标：只去掉本工具加的（desktop.ini 里别的设置留着）"""
    paths = parse_folders(folders)
    if not paths:
        raise ValueError("请填要恢复默认图标的文件夹路径（每行一个）")
    results = []
    for text in paths:
        try:
            folder = _check_folder(text)
            ini = folder / "desktop.ini"
            lines, encoding = _read_ini(ini)
            ours = bool(_folder_icon_id(folder))
            other = any(_ICON_KEYS.match(line) for line in lines) and not ours
            if other:
                raise ValueError("这个文件夹的图标不是本工具设的，没有动")
            if lines:
                left = _set_icon_lines(lines, None)
                if _ini_is_empty(left):
                    _attrs(ini, False)
                    ini.unlink()
                    _make_system_folder(folder, False)
                else:
                    _write_ini(ini, left, encoding)
            _remove_our_icons(folder)
            _record(str(folder), None)
            _notify(folder)
            results.append({"folder": str(folder), "ok": True, "had_icon": ours})
        except PermissionError as err:
            results.append({"folder": text, "ok": False, "error": f"没有权限改这个文件夹（{err.strerror or err}）"})
        except (OSError, ValueError) as err:
            results.append({"folder": text, "ok": False, "error": str(err)})
        _log("恢复默认图标", 文件夹=text, 结果="成功" if results[-1]["ok"] else results[-1]["error"])
    return {"results": results, "ok_count": sum(1 for r in results if r["ok"])}
