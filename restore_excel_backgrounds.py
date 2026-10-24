#!/usr/bin/env python3
"""按原始备份恢复 Excel 中原本有填充色的单元格。Python 3.10+，无第三方依赖。

用法：修改下面「用户配置」中的路径，然后直接运行本文件，无需命令行参数。
  python restore_excel_backgrounds.py

仅处理 .xlsx/.xlsm/.xltx/.xltm；旧版 .xls/.xlsb 会在报告中说明。
按相对路径优先匹配；路径不同则仅匹配两边都唯一的同名文件。
按工作表名称和单元格地址比较，不猜测插入/删除行列后的对应关系。
检查单元格直接设置的填充（包括显式白色），不处理条件格式、表格样式、
整行/整列样式产生的背景。原本没有直接填充的单元格完全不动。
直接修改 OOXML 压缩包中的样式引用，不通过 Excel/openpyxl 重存整个文件。
报告的 SVG 缩略图是单元格颜色示意，并非 Excel 实际截图。
"""

from __future__ import annotations

from collections import Counter, defaultdict
from copy import deepcopy
from datetime import datetime
import html
import os
from pathlib import Path
import posixpath
import re
import shutil
import sys
import tempfile
from urllib.parse import unquote
import xml.etree.ElementTree as ET
import zipfile

# ==================== 用户配置：以后只需要修改这里 ====================
# 旧文件：保留原始背景色的备份文件夹。
BACKUP_DIR = r"D:\原始备份"
# 新文件：需要检查和修复的成品文件夹，会在这里直接修复。
FINISHED_DIR = r"D:\成品"
# 报告保存位置：留空表示保存在本 Python 文件旁边，也可以填 r"D:\检查报告"。
# 每次运行会在该位置创建一个带时间的新文件夹，里面包含报告、图片和修复前副本。
REPORT_ROOT = r""
# True = 只检查并生成报告；False = 直接修复并生成报告。
DRY_RUN = False
# True = 同时恢复字体颜色；False = 只恢复背景色。
# 两种模式均只处理备份中原本有背景色的单元格。
RESTORE_FONT_COLOR = False
# ========================== 用户配置结束 ==========================

# 以下 http 地址是 Excel XML 格式的命名空间标识，不会访问网站或联网。
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SUPPORTED = {".xlsx", ".xlsm", ".xltx", ".xltm"}
EXCEL = SUPPORTED | {".xls", ".xlsb"}
ET.register_namespace("", NS)


def tag(name):
    return f"{{{NS}}}{name}"


def canonical(node):
    if node is None:
        return None
    return (node.tag, tuple(sorted(node.attrib.items())),
            (node.text or "").strip(), tuple(canonical(c) for c in node))


def escaped(value):
    return str(value).replace("\\", "\\\\").replace("|", "\\|").replace("[", "\\[").replace("]", "\\]").replace("<", "&lt;").replace(">", "&gt;").replace("\r", " ").replace("\n", " ")


def link(path, label=None):
    return f"[{escaped(label or path)}](<{Path(path).resolve().as_uri()}>)"


class Book:
    def __init__(self, path):
        self.path = path
        self.zip = zipfile.ZipFile(path)
        try:
            if any(n.startswith("_xmlsignatures/") for n in self.zip.namelist()):
                raise ValueError("文件含数字签名，修改会使签名失效，已跳过")
            self.styles_raw = self.zip.read("xl/styles.xml")
            root = ET.fromstring(self.styles_raw)
            self.fills = list(root.find(tag("fills")))
            self.fonts = list(root.find(tag("fonts")))
            self.xfs = list(root.find(tag("cellXfs")))
            self.theme = []
            if "xl/theme/theme1.xml" in self.zip.namelist():
                theme_root = ET.fromstring(self.zip.read("xl/theme/theme1.xml"))
                scheme = theme_root.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme")
                colors = {c.tag.rsplit("}", 1)[-1]: list(c)[0].get("val")
                          if list(c)[0].tag.endswith("}srgbClr") else list(c)[0].get("lastClr")
                          for c in scheme}
                self.theme = [colors.get(n) for n in
                              ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                               "accent4", "accent5", "accent6", "hlink", "folHlink")]
            rels = ET.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
            targets = {}
            for r in rels:
                if r.get("TargetMode") == "External":
                    continue
                target = unquote(r.get("Target", ""))
                targets[r.get("Id")] = (target.lstrip("/") if target.startswith("/") else
                                         posixpath.normpath(posixpath.join("xl", target)))
            workbook = ET.fromstring(self.zip.read("xl/workbook.xml"))
            self.sheets = {s.get("name"): targets[s.get(f"{{{REL}}}id")]
                           for s in workbook.find(tag("sheets"))
                           if s.get(f"{{{REL}}}id") in targets and
                           "/worksheets/" in targets[s.get(f"{{{REL}}}id") ]}
            self.original_counts = (len(self.fills), len(self.fonts), len(self.xfs))
            self.xf_cache = {}
        except Exception:
            self.zip.close()
            raise

    def close(self):
        self.zip.close()

    def cells(self, sheet):
        root = ET.fromstring(self.zip.read(self.sheets[sheet]))
        return {c.get("r"): c for c in root.findall(f"{tag('sheetData')}/{tag('row')}/{tag('c')}")}

    def style(self, cell):
        return self.xfs[int(cell.get("s", "0"))]

    def fill(self, cell):
        return self.fills[int(self.style(cell).get("fillId", "0"))]

    def colored(self, cell):
        xf = self.style(cell)
        if xf.get("applyFill") == "0":
            return False
        fill = self.fill(cell)
        pattern = fill.find(tag("patternFill"))
        return fill.find(tag("gradientFill")) is not None or (
            pattern is not None and pattern.get("patternType", "none") not in {"none", "gray125"})

    def normalized(self, node):
        """将可解析的主题色转换为实际 RGB，避免不同文件的主题色编号造成误判。"""
        result = deepcopy(node)
        if result is None:
            return None
        for c in result.iter():
            if "theme" in c.attrib:
                index = int(c.get("theme"))
                rgb = self.theme[index] if index < len(self.theme) else None
                if not rgb:
                    raise ValueError(f"主题色 {index} 无法解析")
                tint = float(c.get("tint", "0"))
                # 无 tint 的主题色可准确转为 RGB；带 tint 保留原引用并检查主题一致。
                if tint == 0:
                    c.attrib.clear()
                    c.set("rgb", "FF" + rgb[-6:].upper())
            if "rgb" in c.attrib:
                c.set("rgb", "FF" + c.get("rgb")[-6:].upper())
        return result

    def portable(self, node, destination):
        result = self.normalized(node)
        if result is None:
            return None
        if any("theme" in c.attrib for c in result.iter()) and self.theme != destination.theme:
            raise ValueError("带 tint 的主题色在两文件主题不同时无法安全恢复，已跳过该文件")
        if any("indexed" in c.attrib for c in result.iter()):
            a = ET.fromstring(self.styles_raw).find(tag("colors"))
            b = ET.fromstring(destination.styles_raw).find(tag("colors"))
            if canonical(a) != canonical(b):
                raise ValueError("两文件的索引色调色板不同，已跳过该文件")
        return result

    def font_color(self, cell):
        return self.fonts[int(self.style(cell).get("fontId", "0"))].find(tag("color"))

    def fill_key(self, cell):
        if not self.colored(cell):
            return None
        return self.color_key(self.fill(cell))

    def color_key(self, node):
        normalized = self.normalized(node)
        if normalized is None:
            return None
        theme = tuple(self.theme) if any("theme" in c.attrib for c in normalized.iter()) else None
        palette = (canonical(ET.fromstring(self.styles_raw).find(tag("colors")))
                   if any("indexed" in c.attrib for c in normalized.iter()) else None)
        return canonical(normalized), theme, palette

    def restore_style(self, original, source, current, restore_font):
        fill = original.portable(original.fill(source), self)
        font_color = original.portable(original.font_color(source), self) if restore_font else None
        cache_key = (int(current.get("s", "0")), canonical(fill), restore_font, canonical(font_color))
        if cache_key in self.xf_cache:
            return self.xf_cache[cache_key]
        xf = deepcopy(self.style(current))
        xf.set("fillId", str(intern(self.fills, fill)))
        xf.set("applyFill", "1")
        if restore_font:
            font = deepcopy(self.fonts[int(xf.get("fontId", "0"))])
            for old in list(font.findall(tag("color"))):
                font.remove(old)
            if font_color is not None:
                font.append(font_color)
            xf.set("fontId", str(intern(self.fonts, font)))
            xf.set("applyFont", "1")
        result = intern(self.xfs, xf)
        self.xf_cache[cache_key] = result
        return result

    def styles_output(self):
        # 只追加样式，保留原 XML 的命名空间、扩展、以及既有样式编号。
        raw = self.styles_raw
        for name, nodes, count in zip(("fills", "fonts", "cellXfs"),
                                      (self.fills, self.fonts, self.xfs), self.original_counts):
            if len(nodes) == count:
                continue
            pattern = rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)" + name.encode() + rb"\b[^>]*>.*?</(?P=prefix)" + name.encode() + rb"\s*>"
            match = re.search(pattern, raw, re.S)
            if not match:
                raise ValueError(f"无法安全追加样式：{name}")
            block = match.group()
            opening_end = block.index(b">") + 1
            opening = re.sub(rb'\s+count\s*=\s*[\"\'][^\"\']*[\"\']', b"", block[:opening_end])
            opening = opening[:-1] + f' count="{len(nodes)}">'.encode()
            closing_start = block.rfind(b"</")
            additions = b"".join(ET.tostring(n, encoding="utf-8") for n in nodes[count:])
            replacement = opening + block[opening_end:closing_start] + additions + block[closing_start:]
            raw = raw[:match.start()] + replacement + raw[match.end():]
        return raw


def intern(nodes, node):
    key = canonical(node)
    for index, existing in enumerate(nodes):
        if canonical(existing) == key:
            return index
    nodes.append(node)
    return len(nodes) - 1


def patch_cells(raw, changes):
    remaining = dict(changes)
    def replace(match):
        opening = match.group()
        address = re.search(rb'\br\s*=\s*([\"\'])(.*?)\1', opening)
        if not address or address[2].decode() not in remaining:
            return opening
        style = remaining.pop(address[2].decode())
        opening = re.sub(rb'\s+s\s*=\s*([\"\']).*?\1', b"", opening)
        end = b"/>" if opening.endswith(b"/>") else b">"
        return opening[:-len(end)] + f' s="{style}"'.encode() + end
    result = re.sub(rb"<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*>", replace, raw)
    if remaining:
        # 空白单元格被清除填充后，Excel 可能将整个 <c> 删除；补回样式，绝不复制值。
        by_row = defaultdict(dict)
        for address, style in remaining.items():
            by_row[int(re.search(r"\d+$", address).group())][address] = style
        data_pattern = rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)sheetData\b[^>]*?(?:/>|>.*?</(?P=prefix)sheetData\s*>)"
        data_match = re.search(data_pattern, result, re.S)
        if not data_match:
            raise ValueError("找不到 sheetData，无法恢复已删除的空白单元格")
        data = data_match.group()
        if data.endswith(b"/>"):
            data = data[:-2] + b"></" + data_match['prefix'] + b"sheetData>"

        def cell_xml(address, style):
            return f'<c xmlns="{NS}" r="{address}" s="{style}"/>'.encode()

        def col_number(address):
            value = 0
            for letter in re.match(r"[A-Z]+", address).group():
                value = value * 26 + ord(letter) - 64
            return value

        def insert_into_row(match):
            row = match.group()
            number_match = re.search(rb'\br\s*=\s*([\"\'])(\d+)\1', row.split(b">", 1)[0])
            if not number_match:
                return row
            cells = by_row.pop(int(number_match[2]), None)
            if not cells:
                return row
            if row.endswith(b"/>"):
                row = row[:-2] + b"></" + match['prefix'] + b"row>"
            for address, style in sorted(cells.items(), key=lambda p: col_number(p[0])):
                position = row.rfind(b"</")
                for other in re.finditer(rb"<(?:[A-Za-z_][\w.-]*:)?c\b[^>]*>", row):
                    r = re.search(rb'\br\s*=\s*([\"\'])(.*?)\1', other.group())
                    if r and col_number(r[2].decode()) > col_number(address):
                        position = other.start()
                        break
                row = row[:position] + cell_xml(address, style) + row[position:]
            return row

        row_pattern = rb"<(?P<prefix>(?:[A-Za-z_][\w.-]*:)?)row\b[^>]*?(?:/>|>.*?</(?P=prefix)row\s*>)"
        data = re.sub(row_pattern, insert_into_row, data, flags=re.S)
        for number, cells in sorted(by_row.items()):
            new_row = f'<row xmlns="{NS}" r="{number}">'.encode() + b"".join(
                cell_xml(a, s) for a, s in sorted(cells.items(), key=lambda p: col_number(p[0]))) + b"</row>"
            position = data.rfind(b"</")
            for row in re.finditer(rb"<(?:[A-Za-z_][\w.-]*:)?row\b[^>]*>", data):
                r = re.search(rb'\br\s*=\s*([\"\'])(\d+)\1', row.group())
                if r and int(r[2]) > number:
                    position = row.start()
                    break
            data = data[:position] + new_row + data[position:]
        result = result[:data_match.start()] + data + result[data_match.end():]
        # dimension 是范围缓存；删除缓存让 Excel 根据真实单元格重新计算范围。
        result = re.sub(rb"<(?:[A-Za-z_][\w.-]*:)?dimension\b[^>]*/>", b"", result)
    return result


def color_description(book, cell):
    if not book.colored(cell):
        return "无直接填充", "#FFFFFF"
    fill = book.normalized(book.fill(cell))
    pattern = fill.find(tag("patternFill"))
    kind = pattern.get("patternType", "none") if pattern is not None else "gradient"
    color = pattern.find(tag("fgColor")) if pattern is not None else None
    if color is None:
        color = fill.find(f".//{tag('color')}")
    if color is None:
        return kind, "#EEEEEE"
    text = ", ".join(f"{k}={v}" for k, v in color.attrib.items())
    rgb = color.get("rgb")
    return f"{kind}: {text}", "#" + rgb[-6:] if rgb else "#EEEEEE"


def thumbnail(path, sheet, address, before, after):
    esc = html.escape
    svg = ['<svg xmlns="http://www.w3.org/2000/svg" width="620" height="150" viewBox="0 0 620 150">',
           '<rect width="620" height="150" fill="white"/>',
           f'<text x="14" y="23" font-size="15">{esc(sheet + "!" + address)} — 颜色示意（非截图）</text>']
    for x, title, (description, color) in ((14, "修复前", before), (324, "恢复后", after)):
        svg += [f'<text x="{x}" y="49" font-size="14">{title}</text>',
                f'<rect x="{x}" y="58" width="280" height="46" fill="{color}" stroke="#64748b"/>',
                f'<text x="{x}" y="128" font-size="11">{esc(description[:46])}</text>']
    path.write_text("\n".join(svg + ["</svg>"]), encoding="utf-8")


def save_book(book, updates, safety_copy):
    before_stat = book.path.stat()
    safety_copy.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(book.path, safety_copy)
    handle, temp_name = tempfile.mkstemp(prefix=".restore_", suffix=book.path.suffix, dir=book.path.parent)
    os.close(handle)
    try:
        with zipfile.ZipFile(temp_name, "w") as dest:
            dest.comment = book.zip.comment
            for info in book.zip.infolist():
                dest.writestr(info, updates.get(info.filename, book.zip.read(info.filename)))
        with zipfile.ZipFile(temp_name) as check:
            if check.testzip() is not None:
                raise ValueError("输出文件校验失败，未替换原文件")
            for name in updates:
                ET.fromstring(check.read(name))
        book.close()
        current_stat = book.path.stat()
        if (before_stat.st_size, before_stat.st_mtime_ns) != (current_stat.st_size, current_stat.st_mtime_ns):
            raise ValueError("文件在处理过程中被其他程序修改，未替换")
        os.replace(temp_name, book.path)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)


def compare(source, target, output, relative, dry_run=False, restore_font_color=False):
    original = current = None
    try:
        original = Book(source)
        current = Book(target)
        details, notes, updates = [], [], {}
        for sheet in original.sheets:
            if sheet not in current.sheets:
                notes.append(f"缺少工作表：{sheet}")
                continue
            left, right = original.cells(sheet), current.cells(sheet)
            changes = {}
            for address, source_cell in left.items():
                if not original.colored(source_cell):
                    continue
                target_cell = right.get(address)
                if target_cell is None:
                    target_cell = ET.Element(tag("c"), {"r": address})
                fill_diff = original.fill_key(source_cell) != current.fill_key(target_cell)
                font_diff = restore_font_color and original.color_key(original.font_color(source_cell)) != current.color_key(current.font_color(target_cell))
                if not fill_diff and not font_diff:
                    continue
                # 可解析但不同的主题/索引色上下文在恢复前明确检查。
                new_style = current.restore_style(original, source_cell, target_cell, restore_font_color)
                changes[address] = new_style
                before = color_description(current, target_cell)
                after = color_description(original, source_cell)
                details.append((sheet, address, before, after, bool(font_diff)))
            if changes:
                part = current.sheets[sheet]
                updates[part] = patch_cells(current.zip.read(part), changes)
        safety = None
        if details and not dry_run:
            updates["xl/styles.xml"] = current.styles_output()
            safety = output / "before_restore" / relative
            save_book(current, updates, safety)
        return details, notes, safety
    finally:
        if original is not None:
            original.close()
        if current is not None:
            current.close()


def progress(done, total, message):
    width = 26
    filled = width * done // max(total, 1)
    print(f"\r[{'#' * filled}{'-' * (width - filled)}] {done}/{total} {message[:55]:55}", end="", flush=True)


def within(path, parent):
    return path == parent or parent in path.parents


def collect(root, excluded):
    files = []
    for folder, dirs, names in os.walk(root, followlinks=False):
        dirs[:] = sorted(d for d in dirs if not within((Path(folder) / d).resolve(), excluded)
                         and not (Path(folder) / d).is_symlink()
                         and not (Path(folder) / d / ".excel_background_restore_report").exists())
        for name in sorted(names):
            p = Path(folder) / name
            if p.suffix.lower() in EXCEL and not name.startswith(("~$", ".restore_")) and not p.is_symlink():
                files.append(p)
    return files


def main():
    if not str(BACKUP_DIR).strip() or not str(FINISHED_DIR).strip():
        raise ValueError("请先在文件顶部填写 BACKUP_DIR（旧文件夹）和 FINISHED_DIR（新文件夹）")
    backup, finished = Path(BACKUP_DIR).resolve(), Path(FINISHED_DIR).resolve()
    if not backup.is_dir():
        raise ValueError(f"原始备份文件夹不存在，请修改顶部 BACKUP_DIR：{backup}")
    if not finished.is_dir():
        raise ValueError(f"成品文件夹不存在，请修改顶部 FINISHED_DIR：{finished}")
    if within(backup, finished) or within(finished, backup):
        raise ValueError("备份与成品文件夹必须互相独立，不能相同或互相包含")
    report_root = Path(REPORT_ROOT).resolve() if str(REPORT_ROOT).strip() else Path(__file__).resolve().parent
    output = report_root / datetime.now().strftime("background_report_%Y%m%d_%H%M%S_%f")
    if within(output, backup):
        raise ValueError("报告目录不能放在原始备份文件夹内，请修改顶部 REPORT_ROOT")
    if output.exists():
        raise ValueError("报告目录已存在，请重新运行以创建新的记录")
    output.mkdir(parents=True)
    (output / ".excel_background_restore_report").touch()
    images = output / "images"
    images.mkdir()
    print("正在递归查找 Excel 文件……", flush=True)
    sources, targets = collect(backup, output), collect(finished, output)
    by_relative = defaultdict(list)
    by_name = defaultdict(list)
    source_counts = Counter(p.name.casefold() for p in sources)
    for p in targets:
        by_relative[p.relative_to(finished).as_posix().casefold()].append(p)
        by_name[p.name.casefold()].append(p)
    reserved = {p for source in sources for p in
                by_relative.get(source.relative_to(backup).as_posix().casefold(), [])}
    report = ["# Excel 背景色恢复报告", "", f"- 原始备份：{link(backup)}", f"- 成品文件夹：{link(finished)}",
              f"- 模式：{'仅检查，未修改文件' if DRY_RUN else '修复成品，修复前副本存放于 before_restore'}",
              f"- 字体颜色：{'同时恢复（仅原本有背景色的单元格）' if RESTORE_FONT_COLOR else '不修改'}", "",
              "按工作表名称和单元格地址对应；原本没有直接填充的单元格不处理。显式白色属于有填充。",
              "条件格式、表格样式、行列默认样式不在检查范围。插入/删除/重排过行列的文件请先核对位置。",
              "缩略图仅展示填充色示意，不是 Excel 截图；图案、渐变、索引色及带 tint 的主题色可能只显示占位色。",
              "部分 Markdown 查看器会禁止打开本地 file 链接，可复制报告里的完整路径打开。", ""]
    changed_files = changed_cells = failures = unmatched = 0
    used = set()
    for index, source in enumerate(sources, 1):
        relative = source.relative_to(backup)
        progress(index - 1, len(sources), str(relative))
        report += [f"## {escaped(relative)}", "", f"原始：{link(source)}", ""]
        candidates = by_relative.get(relative.as_posix().casefold(), [])
        if not candidates and source_counts[source.name.casefold()] == 1:
            candidates = [p for p in by_name.get(source.name.casefold(), []) if p not in reserved]
        if len(candidates) != 1:
            unmatched += 1
            report += ["未处理：找不到对应文件，或同名文件存在歧义。", ""]
            continue
        target = candidates[0]
        if target in used:
            unmatched += 1
            report += ["未处理：成品已匹配其他备份文件，避免重复覆盖。", ""]
            continue
        used.add(target)
        report += [f"成品：{link(target)}", ""]
        if source.suffix.lower() not in SUPPORTED:
            failures += 1
            report += ["未处理：旧版 .xls/.xlsb 不支持，请使用 Excel 转为 .xlsx 后比较。", ""]
            continue
        try:
            details, notes, safety = compare(source, target, output, target.relative_to(finished), DRY_RUN, RESTORE_FONT_COLOR)
            changed_files += bool(details)
            changed_cells += len(details)
            report += [f"{'发现待恢复' if DRY_RUN else '已恢复'} {len(details)} 个单元格。", ""]
            if safety:
                report += [f"修复前副本：{link(safety)}", ""]
            if notes:
                report += [f"- {escaped(note)}" for note in notes] + [""]
            if details:
                report += ["| 工作表 | 单元格 | 修复前填充 | 原始 / 恢复后填充 | 字体颜色 | 前后示意 |",
                           "| --- | --- | --- | --- | --- | --- |"]
                for number, (sheet, address, before, after, font_changed) in enumerate(details, 1):
                    image_name = f"{index:05d}_{number:06d}.svg"
                    try:
                        thumbnail(images / image_name, sheet, address, before, after)
                        preview = f"![颜色示意](images/{image_name})"
                    except OSError as error:
                        preview = escaped(f"缩略图生成失败：{error}")
                    report.append(f"| {escaped(sheet)} | {address} | {escaped(before[0])} | {escaped(after[0])} | {'待恢复' if DRY_RUN and font_changed else '已恢复' if font_changed else '不变'} | {preview} |")
                report.append("")
        except Exception as error:
            failures += 1
            report += [f"处理失败（未确认修复）：{escaped(type(error).__name__ + ': ' + str(error))}", ""]
        # 每个文件完成后落盘，中途中断也保留已经完成的记录。
        (output / "report.md").write_text("\n".join(report), encoding="utf-8-sig")
    progress(len(sources), len(sources), "完成")
    print()
    extras = [p for p in targets if p not in used]
    report += ["## 汇总", "", f"- 备份 Excel 文件：{len(sources)}", f"- {'待恢复' if DRY_RUN else '已恢复'}：{changed_files} 个文件，{changed_cells} 个单元格",
               f"- 未匹配 / 歧义：{unmatched}；失败 / 不支持：{failures}",
               f"- 成品中未匹配文件：{len(extras)}", ""]
    report += [f"- {link(p)}" for p in extras]
    report_path = output / "report.md"
    report_path.write_text("\n".join(report), encoding="utf-8-sig")
    print(f"{'待恢复' if DRY_RUN else '已恢复'} {changed_files} 个文件、{changed_cells} 个单元格。")
    print(f"报告：{report_path}")
    return 1 if failures or unmatched else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n已中断；已完成的文件及报告保留。", file=sys.stderr)
        sys.exit(130)
    except (ValueError, OSError) as error:
        print(f"运行失败：{error}", file=sys.stderr)
        sys.exit(1)
