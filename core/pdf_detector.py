import base64
import re
import sys
from pathlib import Path
from typing import Any

try:
    import pymupdf as fitz  # PyMuPDF
except ImportError:
    try:
        import fitz
    except ImportError:
        print("请先安装 pymupdf: pip install pymupdf")
        sys.exit(1)


def extract_page_lines(
    page: fitz.Page,
) -> tuple[list[tuple[float, float, float]], list[tuple[float, float, float]]]:
    """提取 PDF 页面中的所有水平线与垂直线"""
    h_lines: list[tuple[float, float, float]] = []  # (y, x0, x1)
    v_lines: list[tuple[float, float, float]] = []  # (x, y0, y1)
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] == "l":
                p1, p2 = item[1], item[2]
                if abs(p1.y - p2.y) < 1.2 and abs(p1.x - p2.x) > 5.0:
                    h_lines.append((p1.y, min(p1.x, p2.x), max(p1.x, p2.x)))
                elif abs(p1.x - p2.x) < 1.2 and abs(p1.y - p2.y) > 5.0:
                    v_lines.append((p1.x, min(p1.y, p2.y), max(p1.y, p2.y)))
            elif item[0] == "re":
                rect: fitz.Rect = item[1]
                if rect.height < 2.5 and rect.width > 5.0:
                    h_lines.append(((rect.y0 + rect.y1) / 2, rect.x0, rect.x1))
                elif rect.width < 2.5 and rect.height > 5.0:
                    v_lines.append(((rect.x0 + rect.x1) / 2, rect.y0, rect.y1))
                elif drawing.get("color") is not None:
                    h_lines.append((rect.y0, rect.x0, rect.x1))
                    h_lines.append((rect.y1, rect.x0, rect.x1))
                    v_lines.append((rect.x0, rect.y0, rect.y1))
                    v_lines.append((rect.x1, rect.y0, rect.y1))
    return h_lines, v_lines


def find_clipped_spans_on_page(page: fitz.Page) -> list[dict[str, Any]]:
    """
    在单页 PDF 上寻找被边框横线切穿/压线的文字块。
    返回列表: [{"text": str, "bbox": (x0, y0, x1, y1), "line_y": float}]
    """
    h_lines, _ = extract_page_lines(page)
    results: list[dict[str, Any]] = []
    text_dict = page.get_text("dict")
    blocks = text_dict.get("blocks", []) if isinstance(text_dict, dict) else []

    for block in blocks:
        if not isinstance(block, dict) or block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text or text in {"-", "－", "_"}:
                    continue
                b = fitz.Rect(span["bbox"])
                if b.height < 4:
                    continue

                pad = b.height * 0.16
                core_y0 = b.y0 + pad
                core_y1 = b.y1 - pad

                for ly, lx0, lx1 in h_lines:
                    if b.x1 > lx0 + 1.5 and b.x0 < lx1 - 1.5:
                        if core_y0 < ly < core_y1 + 0.6:
                            results.append({
                                "text": text,
                                "bbox": (b.x0, b.y0, b.x1, b.y1),
                                "line_y": ly,
                            })
                            break
    return results


def render_crop_base64(
    page: fitz.Page,
    bbox: tuple[float, float, float, float],
    highlight: bool = True,
    zoom: float = 2.5,
) -> str:
    """将 PDF 指定坐标区域放大截图（默认 2.5 倍高清放大镜）并转为 Base64 PNG，不污染原页面"""
    x0, y0, x1, y1 = bbox
    clip_rect = fitz.Rect(x0 - 85, y0 - 36, x1 + 85, y1 + 28) & page.rect
    annot = None
    try:
        if highlight:
            annot = page.add_rect_annot(fitz.Rect(x0 - 2, y0 - 2, x1 + 2, y1 + 2))
            annot.set_colors(stroke=(1, 0, 0))
            annot.set_border(width=1)
            annot.update()

        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip_rect, annots=True)
    finally:
        if annot is not None:
            try:
                page.delete_annot(annot)
            except Exception:
                pass

    png_bytes = pix.tobytes("png")
    return "data:image/png;base64," + base64.b64encode(png_bytes).decode("ascii")


def render_page_overview(page: fitz.Page, zoom: float = 1.6) -> tuple[str, float, float]:
    """将整页 PDF 渲染为高清全景图 Base64 PNG，并返回 (base64_str, page_width_pt, page_height_pt)"""
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), annots=False)
    png_bytes = pix.tobytes("png")
    b64 = "data:image/png;base64," + base64.b64encode(png_bytes).decode("ascii")
    return b64, round(float(page.rect.width), 2), round(float(page.rect.height), 2)


def texts_match(text_a: str, text_b: str) -> bool:
    """判断复检前后的两段文字是否对应同一单元格内容"""
    a = re.sub(r"\s+", "", text_a.strip())
    b = re.sub(r"\s+", "", text_b.strip())
    if not a or not b:
        return False
    return bool(a and b and a == b)


class RegionTrackingError(ValueError):
    """无法可靠定位原文字时中止调整，不能据此判定修复成功或继续加高。"""


def is_region_still_clipped(
    pdf_path: Path,
    page_index: int,
    orig_text: str,
    orig_bbox: tuple[float, float, float, float],
) -> tuple[bool, str, tuple[float, float, float, float], int]:
    """
    重新导出 PDF 后，检查原来那个单元格区域内的对应文字是否还压线，
    返回: (still_clipped, new_crop_base64, updated_bbox, resolved_page_index)
    """
    orig_x0, _, orig_x1, _ = orig_bbox
    doc = fitz.open(pdf_path)
    try:
        matches: list[tuple[int, tuple[float, float, float, float]]] = []
        # 行高调整可能改变缩放和分页；不能先找任意压线文字，再猜它是原目标。
        for candidate_index in range(len(doc)):
            page = doc[candidate_index]
            text_dict = page.get_text("dict")
            blocks = text_dict.get("blocks", []) if isinstance(text_dict, dict) else []
            for block in blocks:
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        b = fitz.Rect(span["bbox"])
                        if (
                            texts_match(span.get("text", ""), orig_text)
                            and b.x1 >= orig_x0 - 15 and b.x0 <= orig_x1 + 15
                        ):
                            matches.append((candidate_index, (b.x0, b.y0, b.x1, b.y1)))

        if not matches:
            raise RegionTrackingError(
                f"无法定位原文字「{orig_text}」（原第 {page_index + 1} 页），"
                "已停止自动加高，请检查 PDF、Excel 行号或重新扫描"
            )
        if len(matches) != 1:
            raise RegionTrackingError(
                f"原文字「{orig_text}」有 {len(matches)} 个候选位置，"
                "无法唯一定位，已停止自动加高，请人工核对"
            )
        resolved_page_index, target_bbox = matches[0]
        render_page = doc[resolved_page_index]
        still_clipped = any(
            texts_match(clip["text"], orig_text)
            and all(abs(a - b) < 0.1 for a, b in zip(clip["bbox"], target_bbox))
            for clip in find_clipped_spans_on_page(render_page)
        )
        new_img = render_crop_base64(render_page, target_bbox, highlight=still_clipped)
        return still_clipped, new_img, target_bbox, resolved_page_index
    finally:
        doc.close()


def refresh_pdf_pages_overview(
    pdf_path: Path,
    page_indices: set[int],
) -> dict[int, tuple[str, float, float]]:
    """重新渲染指定页码集合的整页全景图，供微调后前端画布实时刷新"""
    updated: dict[int, tuple[str, float, float]] = {}
    if not pdf_path.exists():
        return updated
    doc = fitz.open(pdf_path)
    try:
        for p_idx in page_indices:
            if 0 <= p_idx < len(doc):
                updated[p_idx] = render_page_overview(doc[p_idx])
    finally:
        doc.close()
    return updated
