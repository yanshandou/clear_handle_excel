"""下载用的执行记录报告压缩包：网页版（单个 HTML，截图内嵌）+ Markdown 版（图片放 images 文件夹）。

内容规则同 Markdown 报告：只写结果，不写日期时间和处理过程。
"""
import base64
import html
import io
import zipfile
from typing import Any

from core import fix_images
from core.reporter import (
    _fix_place, _fix_shots, _fix_sort_key, _report_state, _status_text, _unfinished, render_execution_markdown,
)

REPORT_NAME = "execution_report"  # 压缩包里只用英文文件名：日文 Windows 自带解压对中文文件名可能显示乱码
IMAGE_DIR = "images"

_CSS = """
:root { color-scheme: light; }
* { box-sizing: border-box; }
body { margin: 0; background: #f4f6fa; color: #1f2937; font: 14px/1.6 "Yu Gothic UI", "Meiryo", "Microsoft YaHei", sans-serif; }
main { max-width: 1100px; margin: 0 auto; padding: 32px 20px 48px; }
h1 { font-size: 24px; margin: 0 0 12px; }
.meta { background: #fff; border: 1px solid #e5e7eb; border-radius: 10px; padding: 14px 18px; margin-bottom: 24px; }
.meta div { margin: 2px 0; }
.meta code { background: #f3f4f6; padding: 1px 6px; border-radius: 4px; }
.num { font-weight: 700; color: #1d4ed8; }
section.file { background: #fff; border: 1px solid #e5e7eb; border-radius: 10px; padding: 18px 20px; margin-bottom: 20px; }
section.file h2 { font-size: 17px; margin: 0 0 12px; word-break: break-all; }
h3 { font-size: 14.5px; margin: 14px 0 6px; }
ul.sheets { margin: 0; padding-left: 22px; }
ul.sheets del { color: #9ca3af; }
table.fixes { border-collapse: collapse; width: 100%; margin: 4px 0 8px; }
table.fixes th, table.fixes td { border: 1px solid #e5e7eb; padding: 6px 10px; text-align: left; vertical-align: top; }
table.fixes th { background: #f9fafb; }
.shot { border: 1px solid #e5e7eb; border-radius: 8px; padding: 10px 12px; margin: 12px 0; break-inside: avoid; }
.shot-title { font-weight: 700; margin-bottom: 8px; }
.pair { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.pair figure { margin: 0; }
.pair figcaption { font-size: 12.5px; font-weight: 700; margin-bottom: 4px; }
.pair .before figcaption { color: #b91c1c; }
.pair .after figcaption { color: #047857; }
.pair img { width: 100%; height: auto; border: 1px solid #e5e7eb; border-radius: 4px; background: #fff; }
.none { color: #9ca3af; font-size: 12.5px; }
.error { color: #b91c1c; }
@media (max-width: 700px) { .pair { grid-template-columns: 1fr; } }
@media print { body { background: #fff; } main { padding: 0; } section.file, .meta { border: none; padding: 0; } }
"""


def _esc(text: Any) -> str:
    return html.escape(str(text))


def _img_tag(name: str, alt: str) -> str:
    path = fix_images.image_path(name)
    if path is None:
        return '<div class="none">（无截图）</div>'
    data = base64.b64encode(path.read_bytes()).decode("ascii")
    return f'<img src="data:image/png;base64,{data}" alt="{alt}">'


def render_execution_html() -> str:
    state = _report_state()
    files = sorted(state["files"].values(), key=lambda f: int(f.get("index", 0)))
    total_deleted = sum(len(f.get("deleted_backups", [])) for f in files)
    total_fixed = sum(len(f.get("fixes") or {}) for f in files)

    out = [
        "<!doctype html>", '<html lang="zh-CN"><head><meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        "<title>Excel 整理与切字修复记录</title>", f"<style>{_CSS}</style></head><body><main>",
        "<h1>📊 Excel 整理与切字修复记录</h1>",
        '<div class="meta">',
        f"<div><b>目标文件夹</b>：<code>{_esc(state.get('folder') or '-')}</code></div>",
        f"<div><b>处理状态</b>：{_esc(_status_text(state, files))}</div>",
        f'<div><b>汇总</b>：共 {len(files)} 个文件 ｜ 删除备份 Sheet <span class="num">{total_deleted}</span> 个'
        f' ｜ 修复切字 <span class="num">{total_fixed}</span> 处</div>',
        "</div>",
    ]
    for number, f in enumerate(files, start=1):
        title = f["name"]
        if f.get("renamed_to") and f["renamed_to"] != f["name"]:
            title += f" → {f['renamed_to']}"
        title += "（还没处理完）" if _unfinished(f) and state["status"] != "running" else ""
        out += ['<section class="file">', f"<h2>{number}. {_esc(title)}</h2>"]

        all_sheets = list(f.get("all_sheets", []))
        deleted = list(f.get("deleted_backups", []))
        out.append(f"<h3>① 删除备份 Sheet：{len(deleted)} 个</h3>" if deleted else "<h3>① 删除备份 Sheet：无</h3>")
        names = all_sheets + [sh for sh in deleted if sh not in all_sheets]
        if names:
            out.append('<ul class="sheets">')
            out += [f"<li><del>{_esc(sh.strip())}</del></li>" if sh in deleted else f"<li>{_esc(sh)}</li>" for sh in names]
            out.append("</ul>")

        for change in f.get("sheet_changes", []):
            for old, new in change["renamed"].items():
                out.append(f"<div>Sheet 改名：{_esc(old)} → {_esc(new)}</div>")
        fixes = sorted((f.get("fixes") or {}).values(), key=lambda r: _fix_sort_key(r, all_sheets))
        if fixes:
            out += [f"<h3>② 修复切字：{len(fixes)} 处</h3>", '<table class="fixes"><tr><th>#</th><th>Sheet</th><th>位置</th><th>原因</th></tr>']
            out += [
                f"<tr><td>{i}</td><td>{_esc(fix.get('sheet', ''))}</td><td>{_esc(_fix_place(fix) or '-')}</td>"
                f"<td>{_esc(fix.get('reason', '') or '-')}</td></tr>"
                for i, fix in enumerate(fixes, start=1)
            ]
            out.append("</table>")
            for numbers, before, after in _fix_shots(fixes):
                first = fixes[numbers[0] - 1]
                label = "、".join(f"#{n}" for n in numbers)
                out += [
                    '<div class="shot">',
                    f'<div class="shot-title">{label}　{_esc(first.get("sheet", ""))} {_esc(_fix_place(first))}</div>',
                    '<div class="pair">',
                    f'<figure class="before"><figcaption>修复前</figcaption>{_img_tag(before, "修复前")}</figure>',
                    f'<figure class="after"><figcaption>修复后</figcaption>{_img_tag(after, "修复后")}</figure>',
                    "</div></div>",
                ]
        else:
            out.append("<h3>② 修复切字：无</h3>")
        if f.get("status") == "error" and f.get("error"):
            out.append(f'<p class="error">❌ 处理出错：{_esc(f["error"])}</p>')
        out.append("</section>")
    out.append("</main></body></html>")
    return "\n".join(out)


def build_report_zip() -> bytes:
    """压缩包：execution_report.html（双击即可看，图片已内嵌）、execution_report.md + images/"""
    markdown = render_execution_markdown(image_dir=IMAGE_DIR)
    state = _report_state()
    names: set[str] = set()
    for f in state["files"].values():
        for fix in (f.get("fixes") or {}).values():
            names.update(n for n in (fix.get("before"), fix.get("after")) if n and fix_images.image_path(n))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"{REPORT_NAME}.html", render_execution_html())
        archive.writestr(f"{REPORT_NAME}.md", markdown)
        for name in sorted(names):
            path = fix_images.image_path(name)
            if path is not None:
                archive.write(path, f"{IMAGE_DIR}/{name}")
    return buffer.getvalue()
