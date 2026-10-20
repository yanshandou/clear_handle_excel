import base64
import json
import re
import shutil
import sys
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

try:
    import fitz  # PyMuPDF
except ImportError:
    import pymupdf as fitz

try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    print("请先安装 pywin32: pip install pywin32")
    sys.exit(1)

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
EXCLUDE_PDF_SHEETS = {"変更履歴"}
DATE_PATTERN = re.compile(r"(?<!\d)((?:19|20)\d{6})(?!\d)")

XL_TYPE_PDF = 0
XL_SHEET_VISIBLE = -1
XL_WORKSHEET = -4167

# 全局状态存储（供前端网页读取和操作）
STATE = {
    "folder": "",
    "pdf_folder": "",
    "backup_folder": "",
    "items": {},  # id -> 检测出的切字项详情
    "logs": [],
    "busy": False,
}


def log_msg(msg: str) -> None:
    timestamp = datetime.now().strftime("%H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    STATE["logs"].append(line)


def is_backup_sheet(sheet_name: str) -> bool:
    for match in DATE_PATTERN.finditer(sheet_name):
        try:
            datetime.strptime(match.group(1), "%Y%m%d")
            return True
        except ValueError:
            continue
    return False


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name).strip()


def extract_page_lines(page: fitz.Page):
    """提取 PDF 页面中的所有水平线与垂直线"""
    h_lines = []  # (y, x0, x1)
    v_lines = []  # (x, y0, y1)
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


def find_clipped_spans_on_page(page: fitz.Page):
    """
    在单页 PDF 上寻找被边框横线切穿/压线的文字块。
    返回列表: [{"text": str, "bbox": (x0,y0,x1,y1), "line_y": float}]
    """
    h_lines, _ = extract_page_lines(page)
    results = []
    text_dict = page.get_text("dict")

    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text or text in {"-", "－", "_"}:
                    continue
                b = fitz.Rect(span["bbox"])
                if b.height < 4:
                    continue

                # 字体核心高度区域（去掉上下自带的留白）
                pad = b.height * 0.16
                core_y0 = b.y0 + pad
                core_y1 = b.y1 - pad

                for ly, lx0, lx1 in h_lines:
                    if b.x1 > lx0 + 1.5 and b.x0 < lx1 - 1.5:
                        # 水平线穿过了文字核心区，或紧贴底部压线
                        if core_y0 < ly < core_y1 + 0.6:
                            results.append({
                                "text": text,
                                "bbox": (b.x0, b.y0, b.x1, b.y1),
                                "line_y": ly,
                            })
                            break
    return results


def render_crop_base64(page: fitz.Page, bbox: tuple[float, float, float, float], highlight: bool = True) -> str:
    """将 PDF 指定坐标区域放大截图并转为 Base64 PNG"""
    x0, y0, x1, y1 = bbox
    # 向四周扩展视野，方便看清上下文单元格
    clip_rect = fitz.Rect(x0 - 80, y0 - 35, x1 + 80, y1 + 25) & page.rect
    # 临时在页面上画红框截图后再清理，或使用 Pixmap 后处理
    shape = None
    if highlight:
        shape = page.new_shape()
        shape.draw_rect(fitz.Rect(x0 - 2, y0 - 2, x1 + 2, y1 + 2))
        shape.finish(color=(1, 0, 0), width=1.2)
        shape.commit()

    pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2), clip=clip_rect)
    png_bytes = pix.tobytes("png")
    return "data:image/png;base64," + base64.b64encode(png_bytes).decode("ascii")


def match_excel_row(sheet, text_snippet: str, page_index: int, rel_y: float) -> tuple[int, str, float]:
    """
    根据 PDF 中被切断的文字片段和纵向位置，反推对应的 Excel 行号与单元格地址。
    返回: (row_index, cell_address, current_row_height)
    """
    used = sheet.UsedRange
    start_r = used.Row
    start_c = used.Column
    vals = used.Value

    candidates = []
    clean_snip = text_snippet.strip()

    if isinstance(vals, tuple):
        total_r = len(vals)
        for r_idx, row_tuple in enumerate(vals):
            excel_r = start_r + r_idx
            for c_idx, cell_val in enumerate(row_tuple):
                if cell_val is not None:
                    s_val = str(cell_val)
                    if clean_snip in s_val or (len(clean_snip) >= 2 and s_val.endswith(clean_snip)):
                        excel_c = start_c + c_idx
                        candidates.append((excel_r, excel_c))
    else:
        total_r = 1

    # 如果通过文本匹配到了候选单元格，挑选纵向相对位置最接近的一行
    target_r = None
    target_c = 1
    if candidates:
        approx_r = start_r + int(total_r * rel_y)
        candidates.sort(key=lambda rc: abs(rc[0] - approx_r))
        target_r, target_c = candidates[0]
    else:
        # 兜底：按比例估算行号
        target_r = max(1, start_r + int(total_r * rel_y))

    try:
        cell_addr = sheet.Cells(target_r, target_c).Address.replace("$", "")
    except Exception:
        cell_addr = f"Row {target_r}"

    try:
        row_h = float(sheet.Rows(target_r).RowHeight or 15.0)
    except Exception:
        row_h = 15.0

    return target_r, cell_addr, round(row_h, 1)


def is_region_still_clipped(pdf_path: Path, page_index: int, orig_text: str, orig_x0: float, orig_x1: float) -> tuple[bool, str]:
    """
    重新导出 PDF 后，检查原来那个列区间内的对应文字是否还压线，并返回最新截图 Base64
    """
    doc = fitz.open(pdf_path)
    try:
        if page_index >= len(doc):
            page_index = len(doc) - 1
        page = doc[page_index]
        clipped_list = find_clipped_spans_on_page(page)

        still_clipped = False
        target_bbox = None

        for item in clipped_list:
            bx0, by0, bx1, by1 = item["bbox"]
            # 在同一列范围内（X坐标重叠）
            if bx1 >= orig_x0 - 15 and bx0 <= orig_x1 + 15:
                still_clipped = True
                target_bbox = item["bbox"]
                break

        # 如果不压线了，找到该文字现在的新位置截图给用户看
        if target_bbox is None:
            for block in page.get_text("dict").get("blocks", []):
                if block.get("type") != 0:
                    continue
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        b = fitz.Rect(span["bbox"])
                        if b.x1 >= orig_x0 - 15 and b.x0 <= orig_x1 + 15:
                            if orig_text in span.get("text", "") or span.get("text", "").strip() in orig_text:
                                target_bbox = (b.x0, b.y0, b.x1, b.y1)
                                break

        if target_bbox is None:
            target_bbox = (orig_x0, 100, orig_x1, 130)

        new_img = render_crop_base64(page, target_bbox, highlight=still_clipped)
        return still_clipped, new_img
    finally:
        doc.close()


def normalize_workbook_view(excel, wb) -> None:
    """A1 定位 + 100% 缩放 + 切回第一页"""
    if wb.Windows.Count > 0:
        wb.Windows(1).Visible = True
    first_sheet = None
    for sh in wb.Sheets:
        if sh.Visible != XL_SHEET_VISIBLE:
            continue
        if first_sheet is None:
            first_sheet = sh
        try:
            sh.Activate()
            if sh.Type == XL_WORKSHEET:
                excel.Goto(sh.Range("A1"), Scroll=True)
            excel.ActiveWindow.Zoom = 100
        except Exception:
            pass
    if first_sheet is not None:
        try:
            first_sheet.Activate()
        except Exception:
            pass


def scan_folder_workflow(folder_str: str, delete_backups: bool) -> None:
    STATE["busy"] = True
    STATE["items"].clear()
    STATE["logs"].clear()

    try:
        folder = Path(folder_str.strip().strip('"').strip("'")).expanduser().resolve()
        if not folder.is_dir():
            log_msg(f"❌ 文件夹不存在: {folder}")
            return

        workbooks = sorted(
            p for p in folder.iterdir()
            if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS and not p.name.startswith("~$")
        )
        if not workbooks:
            log_msg("❌ 未找到 Excel 文件。")
            return

        timestamp = f"{datetime.now():%Y%m%d_%H%M%S}"
        pdf_folder = folder / f"_exported_pdfs_{timestamp}"
        backup_folder = folder / f"_backup_{timestamp}"
        pdf_folder.mkdir(parents=True, exist_ok=True)

        STATE["folder"] = str(folder)
        STATE["pdf_folder"] = str(pdf_folder)
        STATE["backup_folder"] = str(backup_folder)

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        item_seq = 1
        try:
            for wb_path in workbooks:
                log_msg(f"正在处理并导出: {wb_path.name} ...")
                wb = None
                try:
                    wb = excel.Workbooks.Open(str(wb_path.resolve()))
                    all_names = [s.Name for s in wb.Sheets]
                    bkp_names = [n for n in all_names if is_backup_sheet(n)]

                    if delete_backups and bkp_names and len(bkp_names) < len(all_names):
                        backup_folder.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(wb_path, backup_folder / wb_path.name)
                        for bn in bkp_names:
                            wb.Sheets(bn).Delete()
                        wb.Save()
                        log_msg(f"  已清理 {len(bkp_names)} 个日期备份 Sheet")

                    file_stem = sanitize_filename(wb_path.stem)
                    for sheet in wb.Sheets:
                        s_name = sheet.Name
                        if s_name in EXCLUDE_PDF_SHEETS or is_backup_sheet(s_name):
                            continue
                        if sheet.Visible != XL_SHEET_VISIBLE:
                            continue

                        pdf_path = pdf_folder / f"{file_stem}_{sanitize_filename(s_name)}.pdf"
                        try:
                            sheet.ExportAsFixedFormat(
                                Type=XL_TYPE_PDF,
                                Filename=str(pdf_path.resolve()),
                                Quality=0,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False,
                            )
                        except Exception as e:
                            log_msg(f"  导出 PDF 失败 [{s_name}]: {e}")
                            continue

                        # 扫描刚导出的 PDF
                        doc = fitz.open(pdf_path)
                        try:
                            for p_idx in range(len(doc)):
                                page = doc[p_idx]
                                clipped = find_clipped_spans_on_page(page)
                                for clip_info in clipped:
                                    bx0, by0, bx1, by1 = clip_info["bbox"]
                                    rel_y = by0 / max(1.0, page.rect.height)
                                    excel_r, cell_addr, row_h = match_excel_row(
                                        sheet, clip_info["text"], p_idx, rel_y
                                    )
                                    img_b64 = render_crop_base64(page, clip_info["bbox"], highlight=True)
                                    item_id = f"item_{item_seq}"
                                    item_seq += 1
                                    STATE["items"][item_id] = {
                                        "id": item_id,
                                        "wb_path": str(wb_path),
                                        "wb_name": wb_path.name,
                                        "sheet_name": s_name,
                                        "pdf_path": str(pdf_path),
                                        "page_index": p_idx,
                                        "text": clip_info["text"],
                                        "bbox": clip_info["bbox"],
                                        "excel_row": excel_r,
                                        "cell_addr": cell_addr,
                                        "row_height": row_h,
                                        "image": img_b64,
                                        "status": "pending",  # pending / fixed / max_reached
                                        "iterations": 0,
                                    }
                        finally:
                            doc.close()
                except Exception as err:
                    log_msg(f"❌ 处理文件出错 {wb_path.name}: {err}")
                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        log_msg(f"✅ 扫描完成！共检测到 {len(STATE['items'])} 处疑似切字位置。")
    finally:
        STATE["busy"] = False


def fix_selected_workflow(selected_ids: list[str], step_pt: float = 3.0, max_iter: int = 6) -> None:
    """
    针对用户勾选的卡片：
    每次将对应 Excel 行高 +step_pt -> 重新导出该 Sheet 的 PDF -> 检测该处是否还切字
    一旦完整显示则立即停止！最后执行 A1 + 100% 缩放保存。
    """
    STATE["busy"] = True
    try:
        # 按工作簿分组处理
        by_wb: dict[str, list[dict]] = {}
        for iid in selected_ids:
            item = STATE["items"].get(iid)
            if item:
                by_wb.setdefault(item["wb_path"], []).append(item)

        if not by_wb:
            log_msg("⚠️ 没有选中任何需要修复的项。")
            return

        backup_folder = Path(STATE["backup_folder"])
        backup_folder.mkdir(parents=True, exist_ok=True)

        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            for wb_path_str, items in by_wb.items():
                wb_path = Path(wb_path_str)
                bkp_target = backup_folder / wb_path.name
                if not bkp_target.exists():
                    shutil.copy2(wb_path, bkp_target)

                log_msg(f"🔧 开始微调工作簿: {wb_path.name} (共 {len(items)} 处)...")
                wb = None
                try:
                    wb = excel.Workbooks.Open(str(wb_path.resolve()))
                    for item in items:
                        sheet = wb.Sheets(item["sheet_name"])
                        target_row = int(item["excel_row"])
                        pdf_path = Path(item["pdf_path"])
                        bx0, _, bx1, _ = item["bbox"]

                        # 先检查是否因为同行其他单元格的微调已经顺带修好了
                        if item["iterations"] == 0 and pdf_path.exists():
                            still_clip, new_img = is_region_still_clipped(
                                pdf_path, item["page_index"], item["text"], bx0, bx1
                            )
                            if not still_clip:
                                item["image"] = new_img
                                item["status"] = "fixed"
                                item["row_height"] = round(float(sheet.Rows(target_row).RowHeight or 15.0), 1)
                                log_msg(f"  ✅ [{item['sheet_name']} 行{target_row}] 已随同行调整自动恢复完整显示！")
                                continue

                        for it in range(1, max_iter + 1):
                            curr_h = float(sheet.Rows(target_row).RowHeight or 15.0)
                            new_h = round(curr_h + step_pt, 1)
                            sheet.Rows(target_row).RowHeight = new_h
                            item["row_height"] = new_h
                            item["iterations"] = it

                            # 重新导出该 Sheet 的 PDF
                            sheet.ExportAsFixedFormat(
                                Type=XL_TYPE_PDF,
                                Filename=str(pdf_path.resolve()),
                                Quality=0,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False,
                            )

                            still_clip, new_img = is_region_still_clipped(
                                pdf_path, item["page_index"], item["text"], bx0, bx1
                            )
                            item["image"] = new_img

                            if not still_clip:
                                item["status"] = "fixed"
                                log_msg(
                                    f"  ✅ [{item['sheet_name']} 行{target_row}] 第 {it} 次微调后已完整显示！(高度 -> {new_h}pt)"
                                )
                                break
                            else:
                                log_msg(
                                    f"  ⏳ [{item['sheet_name']} 行{target_row}] 第 {it} 次微调 ({new_h}pt) 仍略微压线，继续+..."
                                )
                        else:
                            item["status"] = "max_reached"
                            log_msg(f"  ⚠️ [{item['sheet_name']} 行{target_row}] 已达最大微调次数 ({max_iter}次)")                    # 顺便完成 A1 定位与 100% 缩放保存
                    normalize_workbook_view(excel, wb)
                    wb.Save()
                    log_msg(f"💾 已保存并完成 A1/100% 标准化: {wb_path.name}")
                except Exception as err:
                    log_msg(f"❌ 修复失败 {wb_path.name}: {err}")
                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        log_msg("🎉 所有选中项微调完毕！可在网页上查看最新截图效果。")
    finally:
        STATE["busy"] = False


HTML_PAGE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>Excel PDF 切字检测与智能微调工具</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Yu Gothic UI", sans-serif; margin: 0; background: #f4f6f9; color: #1e293b; }
  header { background: #1e293b; color: #fff; padding: 16px 28px; display: flex; align-items: center; justify-content: space-between; }
  header h1 { margin: 0; font-size: 18px; font-weight: 600; }
  .container { max-width: 1400px; margin: 20px auto; padding: 0 20px; }
  .panel { background: #fff; border-radius: 10px; padding: 18px 22px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); margin-bottom: 18px; }
  .controls { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
  input[type="text"], input[type="number"] { padding: 9px 12px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 14px; }
  input[type="text"] { flex: 1; min-width: 320px; }
  button { padding: 9px 16px; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; transition: 0.15s; }
  .btn-primary { background: #2563eb; color: #fff; }
  .btn-primary:hover { background: #1d4ed8; }
  .btn-success { background: #059669; color: #fff; }
  .btn-success:hover { background: #047857; }
  .btn-outline { background: #e2e8f0; color: #334155; }
  button:disabled { opacity: 0.5; cursor: not-allowed; }
  .grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(360px, 1fr)); gap: 16px; }
  .card { background: #fff; border: 2px solid #e2e8f0; border-radius: 10px; overflow: hidden; cursor: pointer; transition: 0.15s; position: relative; }
  .card.selected { border-color: #2563eb; box-shadow: 0 0 0 3px rgba(37,99,235,0.2); }
  .card.fixed { border-color: #10b981; }
  .card-header { padding: 10px 14px; background: #f8fafc; border-bottom: 1px solid #e2e8f0; font-size: 13px; display: flex; justify-content: space-between; align-items: center; }
  .badge { padding: 3px 8px; border-radius: 4px; font-size: 12px; font-weight: 600; }
  .badge-pending { background: #fef3c7; color: #b45309; }
  .badge-fixed { background: #d1fae5; color: #065f46; }
  .card-img-wrap { padding: 12px; text-align: center; background: #fff; min-height: 110px; display: flex; align-items: center; justify-content: center; }
  .card-img-wrap img { max-width: 100%; max-height: 160px; border: 1px solid #e2e8f0; border-radius: 4px; }
  .card-body { padding: 10px 14px; font-size: 13px; color: #475569; border-top: 1px solid #f1f5f9; display: flex; justify-content: space-between; align-items: center; }
  .row-adjust { display: flex; align-items: center; gap: 6px; }
  .row-adjust input { width: 64px; padding: 4px 6px; text-align: center; }
  .log-box { background: #0f172a; color: #e2e8f0; font-family: monospace; font-size: 12px; padding: 12px; border-radius: 8px; max-height: 160px; overflow-y: auto; }
</style>
</head>
<body>
<header>
  <h1>📐 Excel 导出 PDF 切字检测 & 闭环行高微调器</h1>
  <span id="statusText">就绪</span>
</header>
<div class="container">
  <div class="panel">
    <div class="controls">
      <input type="text" id="folderInput" placeholder="粘贴 Excel 文件夹路径，例如 C:\\Users\\...\\Desktop\\excel_files">
      <label style="font-size:13px;"><input type="checkbox" id="delBackupChk" checked> 顺便清理 YYYYMMDD 备份表</label>
      <button class="btn-primary" id="scanBtn" onclick="startScan()">1. 导出 PDF 并扫描切字</button>
    </div>
  </div>

  <div class="panel" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
    <div>
      <button class="btn-outline" onclick="toggleAll(true)">全选待处理</button>
      <button class="btn-outline" onclick="toggleAll(false)">取消全选</button>
      <span style="margin-left:12px; font-size:14px;">已选中 <b id="selCount">0</b> 项</span>
    </div>
    <div style="display:flex; align-items:center; gap:10px;">
      <label style="font-size:13px;">每次递增(pt): <input type="number" id="stepInput" value="3.0" step="0.5" style="width:65px;"></label>
      <button class="btn-success" id="fixBtn" onclick="startFix()">2. 自动迭代微调选中项 (直到完整显示)</button>
    </div>
  </div>

  <div class="panel">
    <div class="log-box" id="logBox">等待操作...</div>
  </div>

  <div class="grid" id="cardGrid"></div>
</div>

<script>
let selectedIds = new Set();
let itemsCache = {};

async function pollState() {
  try {
    const res = await fetch('/api/state');
    const data = await res.json();
    document.getElementById('scanBtn').disabled = data.busy;
    document.getElementById('fixBtn').disabled = data.busy;
    document.getElementById('statusText').innerText = data.busy ? '⏳ 后台处理中...' : '✅ 就绪';

    const logBox = document.getElementById('logBox');
    logBox.innerText = data.logs.join('\\n') || '等待操作...';
    logBox.scrollTop = logBox.scrollHeight;

    renderCards(data.items);
  } catch (e) {}
}

function renderCards(items) {
  const grid = document.getElementById('cardGrid');
  const keys = Object.keys(items);
  if (keys.length === 0) {
    grid.innerHTML = '';
    updateCount();
    return;
  }

  keys.forEach(id => {
    const item = items[id];
    let card = document.getElementById('card_' + id);
    if (!card) {
      card = document.createElement('div');
      card.id = 'card_' + id;
      card.className = 'card';
      card.onclick = (e) => {
        if (e.target.tagName === 'INPUT') return;
        if (selectedIds.has(id)) selectedIds.delete(id);
        else selectedIds.add(id);
        card.classList.toggle('selected', selectedIds.has(id));
        updateCount();
      };
      grid.appendChild(card);
      // 默认自动选中新扫出来的待处理项
      if (item.status === 'pending') {
        selectedIds.add(id);
        card.classList.add('selected');
      }
    }

    if (item.status === 'fixed') {
      selectedIds.delete(id);
      card.classList.remove('selected');
      card.classList.add('fixed');
    }

    const badgeClass = item.status === 'fixed' ? 'badge-fixed' : 'badge-pending';
    const badgeText = item.status === 'fixed' ? `✓ 已修复 (+${item.iterations}次)` : '疑似压线/切字';

    card.innerHTML = `
      <div class="card-header">
        <div><b>${item.wb_name}</b> <span style="color:#64748b;">/ ${item.sheet_name}</span></div>
        <span class="badge ${badgeClass}">${badgeText}</span>
      </div>
      <div class="card-img-wrap"><img src="${item.image}" alt="crop"></div>
      <div class="card-body">
        <div>定位: <b>${item.cell_addr}</b> | 当前行高: <b>${item.row_height}pt</b><br><span style="color:#94a3b8;">文字: ${item.text}</span></div>
        <div class="row-adjust" title="如果行号有偏差可手动微调">
          行号: <input type="number" value="${item.excel_row}" onchange="updateRow('${id}', this.value)">
        </div>
      </div>
    `;
  });
  updateCount();
}

function updateCount() {
  document.getElementById('selCount').innerText = selectedIds.size;
}

function toggleAll(select) {
  document.querySelectorAll('.card').forEach(c => {
    const id = c.id.replace('card_', '');
    if (select && !c.classList.contains('fixed')) {
      selectedIds.add(id);
      c.classList.add('selected');
    } else {
      selectedIds.delete(id);
      c.classList.remove('selected');
    }
  });
  updateCount();
}

async function updateRow(id, newRow) {
  await fetch('/api/update_row', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({id, excel_row: parseInt(newRow, 10)})
  });
}

async function startScan() {
  const folder = document.getElementById('folderInput').value.trim();
  if (!folder) { alert('请先输入 Excel 文件夹路径！'); return; }
  selectedIds.clear();
  document.getElementById('cardGrid').innerHTML = '';
  await fetch('/api/scan', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      folder,
      delete_backups: document.getElementById('delBackupChk').checked
    })
  });
}

async function startFix() {
  if (selectedIds.size === 0) { alert('请先点击卡片选中要修复的位置！'); return; }
  const step = parseFloat(document.getElementById('stepInput').value) || 3.0;
  await fetch('/api/fix', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      selected_ids: Array.from(selectedIds),
      step_pt: step
    })
  });
}

setInterval(pollState, 1200);
</script>
</body>
</html>
"""


class WebHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # 保持终端清爽

    def _send_json(self, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = HTML_PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif parsed.path == "/api/state":
            self._send_json({
                "busy": STATE["busy"],
                "logs": STATE["logs"][-30:],
                "items": STATE["items"],
            })
        else:
            self.send_error(404)

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        payload = json.loads(self.rfile.read(length).decode("utf-8")) if length else {}

        if self.path == "/api/scan":
            if not STATE["busy"]:
                threading.Thread(
                    target=scan_folder_workflow,
                    args=(payload.get("folder", ""), bool(payload.get("delete_backups", True))),
                    daemon=True,
                ).start()
            self._send_json({"ok": True})

        elif self.path == "/api/fix":
            if not STATE["busy"]:
                threading.Thread(
                    target=fix_selected_workflow,
                    args=(payload.get("selected_ids", []), float(payload.get("step_pt", 3.0))),
                    daemon=True,
                ).start()
            self._send_json({"ok": True})

        elif self.path == "/api/update_row":
            iid = payload.get("id")
            if iid in STATE["items"]:
                STATE["items"][iid]["excel_row"] = int(payload.get("excel_row", 1))
            self._send_json({"ok": True})


def main():
    port = 8765
    server = HTTPServer(("127.0.0.1", port), WebHandler)
    url = f"http://127.0.0.1:{port}"
    print(f"🚀 Web 界面已启动: {url}")
    print("正在自动打开浏览器... (按 Ctrl+C 可退出程序)")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出。")


if __name__ == "__main__":
    main()
