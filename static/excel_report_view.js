// Excel 对比报告的画法（用户 2026-11-17）：core/excel_report.py 把这个文件原样塞进报告 HTML，报告自己就能显示，
// 不需要连本程序、不需要网络。数据在报告里的 <script id="xr-data" type="application/json">（core/excel_diff.diff 的结果）。
// 左 = 文件 A（旧，红色 −），右 = 文件 B（新，绿色 +）。效果图画法和 static/excel_diff.js 的按格子示意图一样。
(function () {
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const fileName = path => String(path || '').split(/[\\/]/).pop();
  const off = new Set();

  function value(text, kind) {
    if (text === '' || text == null) return '<span>（空）</span>';
    const color = (kind === '字的颜色' || kind === '填充色' || kind === '标签颜色') && /^#[0-9A-F]{6}(?=$| )/.exec(text);
    return (color ? `<span class="xd-swatch" style="background:${color[0]}"></span>` : '') + esc(text);
  }
  function row(kind, whereHtml, oldText, newText, jump = '') {
    const oldEmpty = oldText === '' || oldText == null, newEmpty = newText === '' || newText == null;
    return `<tr data-kind="${esc(kind)}"${jump ? ` data-jump="${esc(jump)}" title="点一下，效果图跳到这里"` : ''}><td>${whereHtml}</td><td class="xd-kind">${esc(kind)}</td>`
      + `<td class="${oldEmpty ? 'xd-none' : 'xd-old'}">${value(oldText, kind)}</td><td class="${newEmpty ? 'xd-none' : 'xd-new'}">${value(newText, kind)}</td></tr>`;
  }
  function table(rows) {
    return `<table class="xd-table"><colgroup><col class="c-where"><col class="c-kind"><col><col></colgroup>`
      + `<thead><tr><th>位置</th><th>类型</th><th>− 文件 A（旧）</th><th>+ 文件 B（新）</th></tr></thead><tbody>${rows.join('')}</tbody></table>`;
  }

  // ───── 效果图：左右各画一张表的样子，不一样的格子左边盖红色、右边盖绿色半透明块 ─────
  const HEAD_W = 46, HEAD_H = 22, GAP_H = 26;
  const colIndex = letters => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  function paneHtml(view, sideKey) {
    const side = view[sideKey], other = view[sideKey === 'old' ? 'new' : 'old'];
    if (!side) return `<div class="xv-pane xv-empty"><p>文件 ${sideKey === 'old' ? 'A' : 'B'} 没有这张表</p></div>`;
    const tone = sideKey;
    const colW = view.cols.map((_, i) => Math.max(side.col_px[i] || 0, other ? other.col_px[i] || 0 : 0));
    const rowH = view.rows.map((_, i) => Math.max(side.row_px[i] || 0, other ? other.row_px[i] || 0 : 0));
    const colX = [], rowY = [];
    let x = 0, y = 0;
    colW.forEach(w => { colX.push(x); x += w; });
    const gaps = [];
    view.rows.forEach((r, i) => {
      const skipped = i === 0 ? r - 1 : r - view.rows[i - 1] - 1;
      if (skipped > 0) { gaps.push({ y, count: skipped }); y += GAP_H; }
      rowY.push(y); y += rowH[i];
    });
    const tail = view.last_row - view.rows[view.rows.length - 1];
    if (tail > 0 || view.rows_cut) { gaps.push({ y, count: tail, cut: view.rows_cut }); y += GAP_H; }
    const width = x, height = y;
    const rowAt = new Map(view.rows.map((r, i) => [r, i]));
    const rowMarks = new Set(view.row_marks), colMarks = new Set(view.col_marks);
    const parts = [];
    colX.forEach((cx, i) => parts.push(`<i class="xv-vl" style="left:${cx + colW[i] - 1}px;height:${height}px"></i>`));
    view.rows.forEach((_, i) => parts.push(`<i class="xv-hl" style="top:${rowY[i] + rowH[i] - 1}px;width:${width}px"></i>`));
    view.rows.forEach((_, i) => { const own = side.row_px[i] || 0; if (own < rowH[i]) parts.push(`<i class="xv-pad" style="left:0;top:${rowY[i] + own}px;width:${width}px;height:${rowH[i] - own}px"></i>`); });
    view.cols.forEach((_, i) => { const own = side.col_px[i] || 0; if (own < colW[i]) parts.push(`<i class="xv-pad" style="left:${colX[i] + own}px;top:0;width:${colW[i] - own}px;height:${height}px"></i>`); });
    for (const gap of gaps) {
      parts.push(`<div class="xv-gap" style="top:${gap.y}px;width:${width}px;height:${GAP_H}px">⋯ ${gap.count > 0 ? `${gap.count} 行一样，没画` : ''}${gap.cut ? `${gap.count > 0 ? '；' : ''}后面还有 ${gap.cut} 行有不同，请看下面的清单` : ''} ⋯</div>`);
    }
    const filled = new Set(side.cells.filter(c => c[2]).map(c => c[0] + ':' + c[1]));
    const boxes = new Map();
    for (const [r, c, text, styleIndex, rs, cs, numeric] of side.cells) {
      const ri = rowAt.get(r);
      if (ri === undefined || c > view.cols.length) continue;
      const [css, wrap, horizontal, vertical] = side.styles[styleIndex];
      let w = 0, h = 0;
      for (let k = 0; k < cs && c - 1 + k < colW.length; k++) w += k === cs - 1 ? (side.col_px[c - 1 + k] || 0) : colW[c - 1 + k];
      for (let k = 0; k < rs; k++) { const at = rowAt.get(r + k); if (at === undefined) break; h += (k === rs - 1 || rowAt.get(r + k + 1) === undefined) ? (side.row_px[at] || 0) : rowH[at]; }
      if (!w || !h) continue;
      boxes.set(r + ':' + c, { x: colX[c - 1], y: rowY[ri], w, h });
      const align = horizontal && horizontal !== 'general' ? horizontal : numeric ? 'right' : 'left';
      const spill = !wrap && align === 'left' && cs === 1 && !filled.has(r + ':' + (c + 1));
      const v = vertical === 'top' ? 'flex-start' : vertical === 'center' ? 'center' : 'flex-end';
      const style = `left:${colX[c - 1]}px;top:${rowY[ri]}px;width:${w}px;height:${h}px;justify-content:${v};text-align:${align === 'centerContinuous' ? 'center' : align};${css}`;
      parts.push(`<div class="xv-cell${wrap ? ' is-wrap' : ''}${spill ? ' is-spill' : ''}" style="${esc(style)}"><span>${esc(text)}</span></div>`);
    }
    for (const [addr, kinds] of Object.entries(view.marks || {})) {
      const m = /^([A-Z]+)(\d+)$/.exec(addr);
      if (!m) continue;
      const c = colIndex(m[1]), r = Number(m[2]), ri = rowAt.get(r);
      if (ri === undefined || c > view.cols.length) continue;
      const box = boxes.get(r + ':' + c) || { x: colX[c - 1], y: rowY[ri], w: side.col_px[c - 1] || colW[c - 1], h: side.row_px[ri] || rowH[ri] };
      parts.push(`<b class="xv-mark is-${tone}" data-addr="${esc(addr)}" style="left:${box.x}px;top:${box.y}px;width:${Math.max(box.w, 6)}px;height:${Math.max(box.h, 6)}px" title="${esc(addr)}：${esc(kinds.join('、'))}不一样"></b>`);
    }
    for (const [top, left, bottom, right, label, own] of (view.boxes || {})[sideKey] || []) {
      const shown = view.rows.map((r, i) => (r >= top && r <= bottom ? i : -1)).filter(i => i >= 0);
      if (!shown.length || left > view.cols.length) continue;
      const first = shown[0], last = shown[shown.length - 1], lastCol = Math.min(right, view.cols.length) - 1;
      const bx = colX[left - 1], by = rowY[first];
      const bw = Math.max(own ? Math.min(own[0], width - bx) : colX[lastCol] + colW[lastCol] - bx, 24);
      const bh = Math.max(own ? Math.min(own[1], height - by) : rowY[last] + rowH[last] - by, 18);
      parts.push(`<b class="xv-shape is-${tone}" style="left:${bx}px;top:${by}px;width:${bw}px;height:${bh}px" title="${esc(label)}"><span>🖼 ${esc(label)}</span></b>`);
    }
    const colHead = view.cols.map((name, i) => colW[i] ? `<span data-c="${name}" class="${colMarks.has(name) ? 'is-' + tone : ''}" style="left:${colX[i]}px;width:${colW[i]}px"${colMarks.has(name) ? ` title="${name} 列的列宽不一样"` : ''}>${name}</span>` : '').join('');
    const rowHead = view.rows.map((r, i) => rowH[i] ? `<span data-r="${r}" class="${rowMarks.has(r) ? 'is-' + tone : ''}" style="top:${rowY[i]}px;height:${rowH[i]}px"${rowMarks.has(r) ? ` title="第 ${r} 行的行高不一样"` : ''}>${r}</span>` : '').join('');
    return `<div class="xv-pane is-${tone}"><div class="xv-fit" data-w="${width + HEAD_W}">`
      + `<div class="xv-colhead" style="width:${width + HEAD_W}px;height:${HEAD_H}px"><em style="width:${HEAD_W}px"></em><div style="left:${HEAD_W}px">${colHead}</div></div>`
      + `<div class="xv-row" style="width:${width + HEAD_W}px"><div class="xv-rowhead" style="width:${HEAD_W}px;height:${height}px">${rowHead}</div>`
      + `<div class="xv-canvas" style="width:${width}px;height:${height}px">${parts.join('')}</div></div></div></div>`;
  }
  function viewHtml(sheet) {
    const view = sheet.view;
    if (!view) return '';
    const count = Object.keys(view.marks || {}).length + (view.marks_cut || 0);
    const shapes = ((view.boxes || {}).old || []).length + ((view.boxes || {}).new || []).length;
    const note = [count ? `色块 = 不一样的格子（${count} 格，鼠标移上去看是哪一类）` : '', shapes ? '虚线框 = 不一样的图片 / 图形占的位置' : '',
      view.row_marks.length ? '行号带颜色 = 行高不一样' : '', view.col_marks.length ? '列标带颜色 = 列宽不一样' : '',
      '斜纹 = 这边的行 / 列比对面小', view.cols_cut ? `右边还有 ${view.cols_cut} 列没画` : '', view.marks_cut ? '不一样的格子太多，图上只标了一部分' : '']
      .filter(Boolean).join('；');
    const scope = view.whole ? `整张表 ${view.rows.length} 行${view.cols_cut ? `、前 ${view.cols.length} 列` : ''}` : '表太大，只画有不同的行和上下 3 行';
    const jumps = `<div class="xr-bar">${count ? `<span class="xr-jump"><button type="button" data-step="-1">⬆ 上一处不同</button><button type="button" data-step="1">⬇ 下一处不同</button><em></em></span>` : ''}`
      + `<span class="xr-zoombar" role="group" aria-label="缩放"><button type="button" data-zoom="out" title="缩小">－</button><b class="xr-zoom-pct">100%</b>`
      + `<button type="button" data-zoom="in" title="放大（放不下就出现左右滚动）">＋</button><button type="button" data-zoom="fit" title="缩到正好一栏宽，没有左右滚动">适合宽度</button>`
      + `<button type="button" data-zoom="100" title="原样大小">100%</button></span></div>`;
    return `<details class="xv-sketch" open><summary>🧩 效果图（按格子画的样子，${scope}）</summary>${jumps}`
      + `<div class="xv-titles"><strong class="is-old">− 文件 A（旧）</strong><strong class="is-new">+ 文件 B（新）</strong></div>`
      + `<div class="xv-wrap">${paneHtml(view, 'old')}${paneHtml(view, 'new')}</div>`
      + `<p class="xr-hint">${note}。图片图形本身画不出来，只标位置。两边一起滚动；点下面清单里的一行，图上会跳到那个位置。</p></details>`;
  }

  function sheetRows(sheet) {
    const rows = [];
    const more = (kind, count, text) => count ? rows.push(`<tr data-kind="${esc(kind)}"><td colspan="4" class="xd-more">${esc(text)}</td></tr>`) : 0;
    for (const item of sheet.text) rows.push(row('文字', esc(item.cell), item.old, item.new, item.cell));
    more('文字', sheet.text_more, `文字还有 ${sheet.text_more} 格不一样，没有全列出来`);
    for (const item of sheet.formulas) rows.push(row('公式', esc(item.cell), item.old, item.new, item.cell));
    more('公式', sheet.formulas_more, `公式还有 ${sheet.formulas_more} 格不一样，没有全列出来`);
    for (const item of sheet.cache || []) rows.push(row('公式缓存', esc(item.cell), item.old, item.new, item.cell));
    more('公式缓存', sheet.cache_more, `还有 ${sheet.cache_more} 格的公式结果不同，没有全列出来`);
    for (const item of sheet.notes || []) rows.push(row(item.kind, esc(item.cell), item.old, item.new, item.cell));
    for (const group of sheet.styles) {
      if (group.more) { more(group.kind, 1, `${group.kind}还有 ${group.more} 种改法（共 ${group.count} 格），没有全列出来`); continue; }
      const cells = group.cells.map(esc).join('、') + (group.count > group.cells.length ? ` 等 ${group.count} 格` : group.count > 1 ? `（共 ${group.count} 格）` : '');
      rows.push(row(group.kind, cells, group.old, group.new, group.cells[0] || ''));
    }
    for (const item of sheet.whole) {
      const m = /第 (\d+) 行/.exec(item.what) || /^([A-Z]+) 列/.exec(item.what) || /^([A-Z]+\d+)(?::|$)/.exec(item.what);
      rows.push(row(item.kind, esc(item.what), item.old, item.new, m ? (/^\d+$/.test(m[1]) ? `R${m[1]}` : /\d/.test(m[1]) ? m[1] : `C${m[1]}`) : ''));
    }
    return rows;
  }

  const STATUS = {
    changed: ['±', '有不同', 'is-changed'], only_old: ['−', '只有文件 A 有', 'is-old'],
    only_new: ['+', '只有文件 B 有', 'is-new'], same: ['＝', '没有不同', 'is-same'],
  };
  function sheetHtml(sheet, index) {
    const [sign, label, cls] = STATUS[sheet.status];
    const id = `xr-sheet-${index}`;
    if (sheet.status !== 'changed') {
      return `<details class="xd-sheet ${cls}" id="${id}"${sheet.view ? ' open' : ''}><summary>${sign} ${esc(sheet.name)}<span class="xd-badge">${label}</span></summary>`
        + (sheet.view ? `<div class="xd-body">${viewHtml(sheet)}</div>` : '') + '</details>';
    }
    const badges = Object.entries(sheet.counts).map(([kind, count]) => `<span class="xd-badge">${esc(kind)} ${count}</span>`).join('');
    return `<details class="xd-sheet ${cls}" id="${id}" open><summary>${sign} ${esc(sheet.name)}${badges}</summary>`
      + `<div class="xd-body">${viewHtml(sheet)}<details class="xd-list" open><summary>📋 差异清单（逐条列出改了什么）</summary>${table(sheetRows(sheet))}</details></div></details>`;
  }

  function render(data, root) {
    const s = data.summary, kinds = Object.entries(s['各类']);
    const order = { changed: 0, only_old: 1, only_new: 1, same: 3 };
    const sheets = data.sheets.map((sheet, index) => ({ sheet, index }));
    const sorted = [...sheets].sort((a, b) => (order[a.sheet.status] - order[b.sheet.status]) || a.index - b.index);
    const onlyOld = data.sheets.filter(x => x.status === 'only_old').length, onlyNew = data.sheets.filter(x => x.status === 'only_new').length;
    const changed = data.sheets.filter(x => x.status === 'changed').length, same = data.sheets.filter(x => x.status === 'same').length;
    const verdict = data.identical
      ? '<div class="xd-verdict is-same">✓ 两份文件在检查的项目里没有不同</div>'
      : `<div class="xd-verdict is-diff">有不同：${[changed ? `${changed} 张表里有改动` : '', onlyOld ? `${onlyOld} 张表只有文件 A 有` : '',
        onlyNew ? `${onlyNew} 张表只有文件 B 有` : '', s['表的顺序变了'] ? '表的顺序变了' : '', (data.book || []).length ? `工作簿设置有 ${(data.book || []).length} 处不同` : '']
        .filter(Boolean).join('，')}</div>`;
    const chips = kinds.map(([kind, count]) => `<button type="button" class="xd-chip" data-kind="${esc(kind)}" title="点一下不看这一类，再点恢复">${esc(kind)}<b>${count}</b></button>`).join('')
      + (onlyOld ? `<span class="xd-chip is-static is-old">只有文件 A 有的表<b>${onlyOld}</b></span>` : '')
      + (onlyNew ? `<span class="xd-chip is-static is-new">只有文件 B 有的表<b>${onlyNew}</b></span>` : '')
      + `<span class="xd-chip is-static is-gray">没有不同的表<b>${same}</b></span>`;
    const toc = sorted.map(({ sheet, index }) => {
      const [sign, , cls] = STATUS[sheet.status];
      const total = Object.values(sheet.counts || {}).reduce((a, b) => a + b, 0);
      return `<li class="${cls}"><a href="#xr-sheet-${index}">${sign} ${esc(sheet.name)}</a>${total ? `<b>${total}</b>` : ''}</li>`;
    }).join('');
    const names = data.names || { old: [], new: [] };
    const orderRow = s['表的顺序变了'] ? [row('表的顺序', '所有表', names.old.join(' → '), names.new.join(' → '))] : [];
    const bookRows = orderRow.concat((data.book || []).map(item => row(item.kind, esc(item.what), item.old, item.new)));
    root.innerHTML = `<header class="xr-head"><h1>Excel 对比报告</h1>`
      + `<div class="xd-files"><div class="xd-file is-old"><strong>− 文件 A（旧）</strong>${esc(fileName(data.theirs))}<small>${esc(data.label_a || data.theirs)}</small></div>`
      + `<div class="xd-file is-new"><strong>+ 文件 B（新）</strong>${esc(fileName(data.mine))}<small>${esc(data.label_b || data.mine)}</small></div></div></header>`
      + verdict
      + `<div class="xr-layout"><nav class="xr-toc" aria-label="表"><h2>表（${data.sheets.length}）</h2><ul>${toc}</ul>`
      + `<p class="xr-tools"><button type="button" data-act="open">全部展开</button><button type="button" data-act="close">全部收起</button><button type="button" data-act="same">藏起没有不同的表</button></p></nav>`
      + `<main class="xr-main">`
      + (kinds.length ? `<p class="xr-hint">数字是不一样的格子 / 处数。点类型可以先不看这一类。</p><div class="xd-chips">${chips}</div>`
        + `<p class="xd-search"><input id="xrSearch" type="search" spellcheck="false" autocomplete="off" placeholder="🔍 在差异里搜：格子（B5）、文字、字体名、颜色…" aria-label="在差异里搜索"><span id="xrSearchCount"></span></p>` : `<div class="xd-chips">${chips}</div>`)
      + (bookRows.length ? `<details class="xd-sheet is-changed" open><summary>📘 整个工作簿</summary><div class="xd-body">${table(bookRows)}</div></details>` : '')
      + sorted.map(({ sheet, index }) => sheetHtml(sheet, index)).join('')
      + `<p class="xr-hint xr-scope">检查的项目：表的增删和顺序、格子的文字 / 公式 / 公式结果 / 批注 / 超链接 / 字体 / 字的颜色 / 填充色 / 边框 / 对齐 / 数字格式 / 保护，`
      + `行高、列宽、隐藏行列、合并单元格、冻结窗格、打印设置（打印区域、方向、缩放、纸张、页边距、打印标题）、页眉页脚、图片图形、标签颜色、表的显示隐藏、条件格式、数据验证、自动筛选、工作表保护、定义的名称、工作簿保护。</p>`
      + `</main></div>`;
    bind(root);
    fit(root);
  }

  function bind(root) {
    root.addEventListener('click', event => {
      const chip = event.target.closest('button.xd-chip');
      if (chip) {
        const kind = chip.dataset.kind;
        if (off.has(kind)) off.delete(kind); else off.add(kind);
        chip.classList.toggle('is-off', off.has(kind));
        root.querySelectorAll('tr[data-kind]').forEach(tr => { if (tr.dataset.kind === kind) tr.classList.toggle('xd-off', off.has(kind)); });
        return;
      }
      const step = event.target.closest('button[data-step]');
      if (step) { stepMark(step.closest('.xv-sketch'), Number(step.dataset.step)); return; }
      const line = event.target.closest('tr[data-jump]');
      if (line) { jumpTo(line.closest('.xd-sheet'), line.dataset.jump); return; }
      const zoomBtn = event.target.closest('button[data-zoom]');
      if (zoomBtn) { setZoom(zoomBtn.closest('.xv-sketch'), zoomBtn.dataset.zoom); return; }
      const tool = event.target.closest('button[data-act]');
      if (!tool) return;
      const act = tool.dataset.act;
      if (act === 'open' || act === 'close') root.querySelectorAll('details').forEach(d => { d.open = act === 'open'; });
      if (act === 'same') {
        const hide = root.classList.toggle('xr-hide-same');
        tool.textContent = hide ? '显示没有不同的表' : '藏起没有不同的表';
      }
    });
    let syncing = false;
    root.addEventListener('scroll', event => {
      const pane = event.target;
      if (syncing || !pane.classList || !pane.classList.contains('xv-pane')) return;
      const mate = [...pane.parentElement.children].find(el => el !== pane && el.classList.contains('xv-pane') && !el.classList.contains('xv-empty'));
      if (!mate) return;
      syncing = true;
      mate.scrollTop = pane.scrollTop; mate.scrollLeft = pane.scrollLeft;
      requestAnimationFrame(() => { syncing = false; });
    }, true);
    root.addEventListener('input', event => {
      if (event.target.id !== 'xrSearch') return;
      const word = event.target.value.trim().toLowerCase();
      let shown = 0, total = 0;
      root.querySelectorAll('tbody tr').forEach(tr => {
        const hit = !word || tr.textContent.toLowerCase().includes(word);
        tr.classList.toggle('xd-miss', !hit);
        total += 1; shown += hit ? 1 : 0;
      });
      root.querySelector('#xrSearchCount').textContent = word ? `匹配 ${shown} / ${total} 行` : '';
    });
  }

  // ───── 定位：清单点一行 / 上一处下一处 → 两边效果图滚到那里并闪一下 ─────
  function place(pane, jump) {
    let m = /^R(\d+)$/.exec(jump);
    const rowSpan = r => pane.querySelector(`.xv-rowhead span[data-r="${r}"]`);
    const colSpan = c => pane.querySelector(`.xv-colhead span[data-c="${c}"]`);
    if (m) { const el = rowSpan(m[1]); return el && { top: el.offsetTop, height: el.offsetHeight, left: 0, width: 0 }; }
    m = /^C([A-Z]+)$/.exec(jump);
    if (m) { const el = colSpan(m[1]); return el && { top: 0, height: 0, left: el.offsetLeft, width: el.offsetWidth }; }
    m = /^([A-Z]+)(\d+)$/.exec(jump);
    if (!m) return null;
    const mark = pane.querySelector(`.xv-mark[data-addr="${jump}"]`);
    if (mark) return { top: mark.offsetTop, height: mark.offsetHeight, left: mark.offsetLeft, width: mark.offsetWidth };
    const r = rowSpan(m[2]), c = colSpan(m[1]);
    return r && c ? { top: r.offsetTop, height: r.offsetHeight, left: c.offsetLeft, width: c.offsetWidth } : null;
  }
  function jumpTo(sheetBox, jump) {
    const sketch = sheetBox && sheetBox.querySelector('.xv-sketch');
    if (!sketch) return false;
    sketch.open = true;
    const panes = [...sketch.querySelectorAll('.xv-pane:not(.xv-empty)')];
    let found = false;
    panes.forEach((pane, i) => {
      const at = place(pane, jump);
      if (!at) return;
      found = true;
      if (i === 0 || !panes[0] || !place(panes[0], jump)) {
        const zoom = Number(pane.dataset.zoom || 1);
        pane.scrollTo({ top: Math.max(0, (at.top + 22) * zoom - pane.clientHeight / 3), left: Math.max(0, (at.left + 46 - 120) * zoom) });
      }
      const canvas = pane.querySelector('.xv-canvas');
      const flash = document.createElement('i');
      flash.className = 'xv-flash';
      flash.style.cssText = at.width && at.height ? `left:${at.left}px;top:${at.top}px;width:${at.width}px;height:${at.height}px`
        : at.width ? `left:${at.left}px;top:0;width:${at.width}px;height:100%` : `left:0;top:${at.top}px;width:100%;height:${at.height}px`;
      canvas.appendChild(flash);
      setTimeout(() => flash.remove(), 2200);
    });
    if (found) sketch.scrollIntoView({ block: 'nearest' });
    return found;
  }
  function stepMark(sketch, dir) {
    const pane = sketch.querySelector('.xv-pane:not(.xv-empty)');
    if (!pane) return;
    const seen = new Set(), list = [];
    sketch.querySelectorAll('.xv-mark[data-addr]').forEach(el => {
      if (seen.has(el.dataset.addr)) return;
      seen.add(el.dataset.addr);
      list.push({ addr: el.dataset.addr, top: el.offsetTop, left: el.offsetLeft });
    });
    list.sort((a, b) => a.top - b.top || a.left - b.left);
    if (!list.length) return;
    let at = Number(sketch.dataset.at ?? -1) + dir;
    at = (at + list.length) % list.length;
    sketch.dataset.at = at;
    sketch.querySelector('.xr-jump em').textContent = `第 ${at + 1} / ${list.length} 处：${list[at].addr}`;
    jumpTo(sketch.closest('.xd-sheet'), list[at].addr);
  }

  // 不要水平滚动（用户 2026-11-17）：默认整张表等比例缩到正好一栏宽；可以用 ＋ － 自己调，放大后出现左右滚动（仍左右联动）
  const STEPS = [0.1, 0.15, 0.2, 0.25, 0.33, 0.4, 0.5, 0.6, 0.75, 0.9, 1, 1.25, 1.5, 2];
  function fitZoom(pane) {
    const inner = pane.querySelector('.xv-fit');
    return Math.min(1, (pane.clientWidth - 2) / Number(inner.dataset.w || 1));
  }
  function applyZoom(sketch) {
    const panes = [...sketch.querySelectorAll('.xv-pane:not(.xv-empty)')];
    if (!panes.length || !panes[0].clientWidth) return;
    const mode = sketch.dataset.zoomMode || 'fit';
    const zoom = mode === 'fit' ? fitZoom(panes[0]) : Number(mode);
    panes.forEach(pane => {
      pane.querySelector('.xv-fit').style.zoom = zoom;
      pane.dataset.zoom = zoom;
      pane.classList.toggle('is-wide', zoom > fitZoom(pane) + 0.001);
    });
    const pct = sketch.querySelector('.xr-zoom-pct');
    if (pct) pct.textContent = `${Math.round(zoom * 100)}%`;
    sketch.querySelectorAll('button[data-zoom]').forEach(b => b.classList.toggle('is-on', b.dataset.zoom === mode || (b.dataset.zoom === '100' && mode === '1')));
  }
  function setZoom(sketch, how) {
    const pane = sketch.querySelector('.xv-pane:not(.xv-empty)');
    if (!pane) return;
    const now = Number(pane.dataset.zoom || 1);
    let mode = how === 'fit' ? 'fit' : how === '100' ? '1' : null;
    if (!mode) {
      const next = how === 'in' ? STEPS.find(z => z > now + 0.001) : [...STEPS].reverse().find(z => z < now - 0.001);
      mode = String(next ?? now);
    }
    sketch.dataset.zoomMode = mode;
    applyZoom(sketch);
  }
  function fit(root) {
    const list = root.matches && root.matches('.xv-sketch') ? [root] : root.querySelectorAll('.xv-sketch');
    list.forEach(applyZoom);
  }
  let fitTimer = 0;
  window.addEventListener('resize', () => { clearTimeout(fitTimer); fitTimer = setTimeout(() => fit(document), 150); });
  document.addEventListener('toggle', event => { if (event.target.matches && event.target.matches('details')) fit(event.target); }, true);

  window.ExcelReport = { render, fit };
  const holder = document.getElementById('xr-data');
  if (holder) {
    try {
      render(JSON.parse(holder.textContent), document.getElementById('xr-root'));
    } catch (err) {
      document.getElementById('xr-root').textContent = `报告显示不出来：${err}`;
    }
  }
})();
