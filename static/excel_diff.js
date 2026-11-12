// 仿 Git 的 Excel 对比（用户 2026-11-11）：左 = 远程现在这份（旧），右 = 你要上传的这份（新）
(function () {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const params = new URLSearchParams(location.search);
  $('xdMine').value = params.get('mine') || '';
  $('xdTheirs').value = params.get('theirs') || '';
  let running = false;
  const off = new Set();  // 点掉不看的类型

  function split(path) {
    const at = Math.max(path.lastIndexOf('\\'), path.lastIndexOf('/'));
    return at < 0 ? { folder: '', file: path } : { folder: path.slice(0, at), file: path.slice(at + 1) };
  }
  function link(label, path, sheet = '', cell = '') {
    const { folder, file } = split(path);
    return window.excelLinkHtml(esc(label), { folder, file, sheet, cell });
  }
  // 值：颜色带色块；空的写「（空）」
  function value(text, kind) {
    if (text === '' || text == null) return '<span>（空）</span>';
    const m = /^#[0-9A-F]{6}$/.test(text) && (kind === '字的颜色' || kind === '填充色' || kind === '标签颜色');
    return (m ? `<span class="xd-swatch" style="background:${text}"></span>` : '') + esc(text);
  }
  function row(kind, whereHtml, oldText, newText) {
    const oldEmpty = oldText === '' || oldText == null, newEmpty = newText === '' || newText == null;
    return `<tr data-kind="${esc(kind)}"${off.has(kind) ? ' class="xd-off"' : ''}><td>${whereHtml}</td><td class="xd-kind">${esc(kind)}</td>`
      + `<td class="${oldEmpty ? 'xd-none' : 'xd-old'}">${value(oldText, kind)}</td><td class="${newEmpty ? 'xd-none' : 'xd-new'}">${value(newText, kind)}</td></tr>`;
  }

  // ───── 效果图（用户 2026-11-11）：左右各画一张表的样子，不一样的格子左边盖红色、右边盖绿色半透明块 ─────
  // 只画有不同的行和上下几行；左右两边每行每列占同样大的位置（取两边较大的），所以始终对得齐
  const HEAD_W = 46, HEAD_H = 22, GAP_H = 26;
  const colIndex = letters => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  function paneHtml(view, sideKey, sheetName) {
    const side = view[sideKey], other = view[sideKey === 'old' ? 'new' : 'old'];
    if (!side) return `<div class="xv-pane xv-empty"><p>${sideKey === 'old' ? '远程' : '你这份'}没有这张表</p></div>`;
    const tone = sideKey === 'old' ? 'old' : 'new';
    // 每列 / 每行的位置：格子大小取两边较大的，自己这边实际多大就画多大
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
    // 网格线
    colX.forEach((cx, i) => parts.push(`<i class="xv-vl" style="left:${cx + colW[i] - 1}px;height:${height}px"></i>`));
    view.rows.forEach((_, i) => parts.push(`<i class="xv-hl" style="top:${rowY[i] + rowH[i] - 1}px;width:${width}px"></i>`));
    // 这边比对面矮 / 窄的行列：多出来的位置画斜纹，表示「这边没这么大」
    view.rows.forEach((_, i) => { const own = side.row_px[i] || 0; if (own < rowH[i]) parts.push(`<i class="xv-pad" style="left:0;top:${rowY[i] + own}px;width:${width}px;height:${rowH[i] - own}px"></i>`); });
    view.cols.forEach((_, i) => { const own = side.col_px[i] || 0; if (own < colW[i]) parts.push(`<i class="xv-pad" style="left:${colX[i] + own}px;top:0;width:${colW[i] - own}px;height:${height}px"></i>`); });
    for (const gap of gaps) {
      parts.push(`<div class="xv-gap" style="top:${gap.y}px;width:${width}px;height:${GAP_H}px">⋯ ${gap.count > 0 ? `${gap.count} 行一样，没画` : ''}${gap.cut ? `${gap.count > 0 ? '；' : ''}后面还有 ${gap.cut} 行有不同，太多没画，请看下面的清单` : ''} ⋯</div>`);
    }
    const filled = new Set(side.cells.filter(c => c[2]).map(c => c[0] + ':' + c[1]));
    const boxes = new Map();  // "行:列" → 这个格子画在哪（合并格按整块）
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
      const spill = !wrap && align === 'left' && cs === 1 && !filled.has(r + ':' + (c + 1));  // 右边空着：字可以像 Excel 那样伸过去
      const v = vertical === 'top' ? 'flex-start' : vertical === 'center' ? 'center' : 'flex-end';
      parts.push(`<div class="xv-cell${wrap ? ' is-wrap' : ''}${spill ? ' is-spill' : ''}" style="left:${colX[c - 1]}px;top:${rowY[ri]}px;width:${w}px;height:${h}px;justify-content:${v};text-align:${align === 'centerContinuous' ? 'center' : align};${css}"><span>${esc(text)}</span></div>`);
    }
    // 不一样的格子：半透明色块
    for (const [addr, kinds] of Object.entries(view.marks || {})) {
      const m = /^([A-Z]+)(\d+)$/.exec(addr);
      if (!m) continue;
      const c = colIndex(m[1]), r = Number(m[2]), ri = rowAt.get(r);
      if (ri === undefined || c > view.cols.length) continue;
      const box = boxes.get(r + ':' + c) || { x: colX[c - 1], y: rowY[ri], w: side.col_px[c - 1] || colW[c - 1], h: side.row_px[ri] || rowH[ri] };
      parts.push(`<b class="xv-mark is-${tone}" style="left:${box.x}px;top:${box.y}px;width:${Math.max(box.w, 6)}px;height:${Math.max(box.h, 6)}px" title="${esc(addr)}：${esc(kinds.join('、'))}不一样"></b>`);
    }
    // 不一样的图片 / 图形：在它占的格子上画虚线框（图片本身画不出来，只标位置）
    for (const [top, left, bottom, right, label, own] of (view.boxes || {})[sideKey] || []) {
      const shown = view.rows.map((r, i) => (r >= top && r <= bottom ? i : -1)).filter(i => i >= 0);
      if (!shown.length || left > view.cols.length) continue;
      const first = shown[0], last = shown[shown.length - 1], lastCol = Math.min(right, view.cols.length) - 1;
      const bx = colX[left - 1], by = rowY[first];
      const bw = Math.max(own ? Math.min(own[0], width - bx) : colX[lastCol] + colW[lastCol] - bx, 24);
      const bh = Math.max(own ? Math.min(own[1], height - by) : rowY[last] + rowH[last] - by, 18);
      parts.push(`<b class="xv-shape is-${tone}" style="left:${bx}px;top:${by}px;width:${bw}px;height:${bh}px" title="${esc(label)}"><span>🖼 ${esc(label)}</span></b>`);
    }
    const colHead = view.cols.map((name, i) => colW[i] ? `<span class="${colMarks.has(name) ? 'is-' + tone : ''}" style="left:${colX[i]}px;width:${colW[i]}px"${colMarks.has(name) ? ` title="${name} 列的列宽不一样"` : ''}>${name}</span>` : '').join('');
    const rowHead = view.rows.map((r, i) => rowH[i] ? `<span class="${rowMarks.has(r) ? 'is-' + tone : ''}" style="top:${rowY[i]}px;height:${rowH[i]}px"${rowMarks.has(r) ? ` title="第 ${r} 行的行高不一样"` : ''}>${r}</span>` : '').join('');
    return `<div class="xv-pane is-${tone}" data-sheet="${esc(sheetName)}">`
      + `<div class="xv-colhead" style="width:${width + HEAD_W}px;height:${HEAD_H}px"><em style="width:${HEAD_W}px"></em><div style="left:${HEAD_W}px">${colHead}</div></div>`
      + `<div class="xv-row" style="width:${width + HEAD_W}px"><div class="xv-rowhead" style="width:${HEAD_W}px;height:${height}px">${rowHead}</div>`
      + `<div class="xv-canvas" style="width:${width}px;height:${height}px">${parts.join('')}</div></div></div>`;
  }
  function viewHtml(sheet) {
    const view = sheet.view;
    if (!view) return '';
    const count = Object.keys(view.marks || {}).length + (view.marks_cut || 0);
    const shapes = ((view.boxes || {}).old || []).length + ((view.boxes || {}).new || []).length;
    const note = [shapes ? '虚线框 = 不一样的图片 / 图形占的位置（图片本身画不出来）' : '', count ? `色块 = 不一样的格子（${count} 格，鼠标移上去看是哪一类）` : '', view.row_marks.length ? '行号带颜色 = 行高不一样' : '',
      view.col_marks.length ? '列标带颜色 = 列宽不一样' : '', '斜纹 = 这边的行 / 列比对面小', view.cols_cut ? `右边还有 ${view.cols_cut} 列没画` : '',
      view.marks_cut ? `不一样的格子太多，只标了前 5000 格` : ''].filter(Boolean).join('；');
    // 真实效果在上面（static/excel_diff_real.js），这张按格子画的示意图收起来，要看格子位置再展开（用户 2026-11-12）
    return `<details class="xv-sketch"><summary>🧩 按格子画的示意图（不是真实效果，看格子位置用）</summary><div class="xv-titles"><strong class="is-old">− 远程现在这份（旧）</strong><strong class="is-new">+ 你要上传的这份（新）</strong></div>`
      + `<div class="xv-wrap">${paneHtml(view, 'old', sheet.name)}${paneHtml(view, 'new', sheet.name)}</div>`
      + `<p class="toolbar-hint xv-note">按格子画出来的样子（字体、颜色、边框、行高列宽、合并格），图片图形本身画不出来；只画有不同的行和上下 3 行。${note}。两边一起滚动。</p></details>`;
  }

  function sheetHtml(sheet, data) {
    const doomed = sheet.doomed ? `<span class="xd-badge">精简会删（${esc(sheet.doomed)}）</span>` : '';
    if (sheet.status === 'only_old' || sheet.status === 'only_new') {
      const old = sheet.status === 'only_old';
      const cls = sheet.doomed ? 'is-doomed' : old ? 'is-old' : 'is-new';
      const text = old ? '只有远程有，你这份没有' + (sheet.doomed ? '（精简注定会删，不算差异）' : '（上传后这张表就没了）')
        : '只有你这份有，远程没有' + (sheet.doomed ? '（精简注定会删，不算差异）' : '（上传后会多出这张表）');
      return `<details class="xd-sheet ${cls}" data-sheet="${esc(sheet.name)}"${sheet.view || !sheet.doomed ? ' open' : ''}><summary>${old ? '−' : '+'} ${link(sheet.name, old ? data.theirs : data.mine, sheet.name)}`
        + `<span class="xd-badge">${text}</span>${doomed}</summary>${sheet.view ? `<div class="xd-body">${viewHtml(sheet)}</div>` : ''}</details>`;
    }
    if (sheet.status === 'same') {
      return `<details class="xd-sheet is-same"><summary>＝ ${esc(sheet.name)}<span class="xd-badge">完全一样</span>${doomed}</summary></details>`;
    }
    const where = (label, cell) => `${link(label, data.mine, sheet.name, cell)}`;
    const rows = [];
    for (const item of sheet.text) rows.push(row('文字', where(item.cell, item.cell), item.old, item.new));
    if (sheet.text_more) rows.push(`<tr data-kind="文字"><td colspan="4" class="xd-more">文字还有 ${sheet.text_more} 格不一样，没有全列出来</td></tr>`);
    for (const item of sheet.formulas) rows.push(row('公式', where(item.cell, item.cell), item.old, item.new));
    if (sheet.formulas_more) rows.push(`<tr data-kind="公式"><td colspan="4" class="xd-more">公式还有 ${sheet.formulas_more} 格不一样，没有全列出来</td></tr>`);
    for (const item of sheet.notes || []) rows.push(row(item.kind, where(item.cell, item.cell), item.old, item.new));
    for (const group of sheet.styles) {
      if (group.more) { rows.push(`<tr data-kind="${esc(group.kind)}"><td colspan="4" class="xd-more">${esc(group.kind)}还有 ${group.more} 种改法（共 ${group.count} 格），没有全列出来</td></tr>`); continue; }
      const cells = group.cells.map(cell => where(cell, cell)).join('、') + (group.count > group.cells.length ? ` 等 ${group.count} 格` : '');
      rows.push(row(group.kind, cells, group.old, group.new));
    }
    for (const item of sheet.whole) rows.push(row(item.kind, esc(item.what), item.old, item.new));
    const badges = Object.entries(sheet.counts).map(([kind, count]) => `<span class="xd-badge">${esc(kind)} ${count}</span>`).join('');
    return `<details class="xd-sheet is-changed" data-sheet="${esc(sheet.name)}" open><summary>± ${link(sheet.name, data.mine, sheet.name)}${badges}${doomed}</summary>`
      + `<div class="xd-body">${viewHtml(sheet)}<details class="xd-list"${sheet.view ? '' : ' open'}><summary>📋 差异清单（逐条列出改了什么）</summary><table class="xd-table"><colgroup><col class="c-where"><col class="c-kind"><col><col></colgroup>`
      + `<thead><tr><th>位置（点了在你这份里打开）</th><th>类型</th><th>− 远程现在这份（旧）</th><th>+ 你要上传的这份（新）</th></tr></thead>`
      + `<tbody>${rows.join('')}</tbody></table></details></div></details>`;
  }

  function render(data) {
    const s = data.summary;
    const kinds = Object.entries(s['各类']);
    const verdict = data.identical
      ? `<div class="xd-verdict is-same">✓ 两份完全一样（文字、公式、样式、行高列宽、打印设置都没有不同${s['精简会删不算'] ? '；精简会删的表不算' : ''}）</div>`
      : `<div class="xd-verdict is-diff">有不同：${[
        s['有不同的表'] ? `${s['有不同的表']} 张表里有改动` : '', s['只有远程有'] ? `${s['只有远程有']} 张表只有远程有` : '',
        s['只有你这份有'] ? `${s['只有你这份有']} 张表只有你这份有` : '', s['表的顺序变了'] ? '表的顺序变了' : '',
      ].filter(Boolean).join('，')}。请看下面是不是都是你自己改的。</div>`;
    const chips = kinds.map(([kind, count]) => `<button type="button" class="xd-chip${off.has(kind) ? ' is-off' : ''}" data-kind="${esc(kind)}" title="点一下不看这一类，再点恢复">${esc(kind)}<b>${count}</b></button>`).join('')
      + (s['只有远程有'] ? `<span class="xd-chip is-static is-old">只有远程有的表<b>${s['只有远程有']}</b></span>` : '')
      + (s['只有你这份有'] ? `<span class="xd-chip is-static is-new">只有你这份有的表<b>${s['只有你这份有']}</b></span>` : '')
      + (s['精简会删不算'] ? `<span class="xd-chip is-static is-gray">精简会删、不算<b>${s['精简会删不算']}</b></span>` : '')
      + `<span class="xd-chip is-static is-gray">完全一样的表<b>${s['一样的表']}</b></span>`;
    const order = { changed: 0, only_old: 1, only_new: 1, same: 3 };
    const sheets = data.sheets.map((sheet, index) => ({ sheet, index }))
      .sort((a, b) => ((a.sheet.doomed && a.sheet.status !== 'changed' ? 2 : order[a.sheet.status]) - (b.sheet.doomed && b.sheet.status !== 'changed' ? 2 : order[b.sheet.status])) || a.index - b.index)
      .map(({ sheet }) => sheetHtml(sheet, data)).join('');
    $('xdResult').innerHTML = `<div class="xd-files">`
      + `<div class="xd-file is-old"><strong>− 远程现在这份（旧）</strong>${link(split(data.theirs).file, data.theirs)}<br><small>${esc(data.theirs)}</small></div>`
      + `<div class="xd-file is-new"><strong>+ 你要上传的这份（新）</strong>${link(split(data.mine).file, data.mine)}<br><small>${esc(data.mine)}</small></div></div>`
      + verdict + (kinds.length ? `<p class="toolbar-hint">下面的数字是不一样的格子 / 处数。点类型可以先不看这一类（比如只想看文字）。</p>` : '')
      + `<div class="xd-chips">${chips}</div>`
      + (kinds.length ? `<p class="xd-search"><input id="xdSearch" type="search" spellcheck="false" autocomplete="off" placeholder="🔍 在差异里搜：格子（B5）、文字、字体名、颜色…" aria-label="在差异里搜索"><span id="xdSearchCount"></span></p>` : '')
      + sheets;
  }

  $('xdResult').addEventListener('click', event => {
    const chip = event.target.closest('button.xd-chip');
    if (!chip) return;
    const kind = chip.dataset.kind;
    if (off.has(kind)) off.delete(kind); else off.add(kind);
    chip.classList.toggle('is-off', off.has(kind));
    for (const tr of $('xdResult').querySelectorAll('tr[data-kind]')) {
      if (tr.dataset.kind === kind) tr.classList.toggle('xd-off', off.has(kind));
    }
    console.info('🔍 Excel 对比：不看的类型', [...off]);
  });

  // 左右两张效果图一起滚动
  let syncing = false;
  $('xdResult').addEventListener('scroll', event => {
    const pane = event.target;
    if (syncing || !pane.classList || !pane.classList.contains('xv-pane')) return;
    const mate = [...pane.parentElement.children].find(el => el !== pane && el.classList.contains('xv-pane') && !el.classList.contains('xv-empty'));
    if (!mate) return;
    syncing = true;
    mate.scrollTop = pane.scrollTop; mate.scrollLeft = pane.scrollLeft;
    requestAnimationFrame(() => { syncing = false; });
  }, true);

  // 搜索：只藏不匹配的行，不重画（页面不动）
  $('xdResult').addEventListener('input', event => {
    if (event.target.id !== 'xdSearch') return;
    const word = event.target.value.trim().toLowerCase();
    let shown = 0, total = 0;
    for (const tr of $('xdResult').querySelectorAll('tbody tr')) {
      const hit = !word || tr.textContent.toLowerCase().includes(word);
      tr.classList.toggle('xd-miss', !hit);
      total += 1; shown += hit ? 1 : 0;
    }
    $('xdSearchCount').textContent = word ? `匹配 ${shown} / ${total} 行` : '';
  });

  async function run() {
    const mine = $('xdMine').value.trim(), theirs = $('xdTheirs').value.trim();
    $('xdError').textContent = '';
    if (!mine || !theirs) { $('xdError').textContent = '请填写两份 Excel 的完整路径'; return; }
    if (running) return;
    running = true;
    $('xdRun').disabled = true;
    $('xdStatus').textContent = '正在读两份文件并对比（大文件要等一会儿）…';
    const started = performance.now();
    try {
      const response = await fetch('/api/excel_diff', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ mine, theirs }) });
      const data = await response.json();
      if (!response.ok || data.ok === false) throw new Error(data.error || '对比失败');
      console.info('🔍 Excel 对比', { 用时秒: ((performance.now() - started) / 1000).toFixed(1), 完全一样: data.identical, ...data.summary });
      off.clear();
      render(data);
      offerUpload(data);
      window.dispatchEvent(new CustomEvent('xd:done', { detail: data }));  // 一页对比里嵌着时告诉外面（static/excel_diff_embed.js）
      $('xdStatus').textContent = `✓ 对比完成，用了 ${((performance.now() - started) / 1000).toFixed(1)} 秒。`;
      history.replaceState(null, '', `?theirs=${encodeURIComponent(theirs)}&mine=${encodeURIComponent(mine)}`);
    } catch (error) {
      $('xdStatus').textContent = '';
      $('xdError').textContent = error.message;
    } finally {
      running = false;
      $('xdRun').disabled = false;
    }
  }
  // 从上传页的清单进来（当前任务带 task，起点 → 终点带 free = 清单 id，都带 rel）：对比完可以直接上传这一份（用户 2026-11-12）
  const uploadTask = params.get('task') || '', uploadFree = params.get('free') || '', uploadRel = params.get('rel') || '';
  let seen = null, uploading = false;
  const load = key => { try { return localStorage.getItem(key); } catch (_) { return null; } };
  // 传什么按上传页选好的来（同一份 localStorage）：excel 只传表格 / both 表格和 PDF 一起 / pdf 只传 PDF
  const pdfMode = () => { const mode = load('clear_handle_excel.uploadWithPdf'); return mode === 'both' || mode === 'pdf' ? mode : 'excel'; };
  const MODE_TEXT = { excel: '只传表格（按上传页选的）。', both: '表格和同名 PDF 一起传（按上传页选的）。', pdf: '只传同名 PDF，Excel 不传（按上传页选的）。' };
  function pdfOptions() {
    const skip = load('clear_handle_excel.pdfBundleSkip') ?? '変更履歴';
    return { include_hidden: load('clear_handle_excel.pdfBundleHidden') === '1',
      skip_sheets: [...new Set(skip.split(/[\n,，、;；]+/).map(v => v.trim()).filter(Boolean))],
      no_header_footer: load('clear_handle_excel.pdfBundleNoHeader') !== '0' };
  }
  function showMode() {
    const mode = pdfMode();
    $('xdUploadMode').textContent = MODE_TEXT[mode];
    $('xdUpload').textContent = mode === 'pdf' ? '⬆ 没问题，上传这一份的 PDF' : mode === 'both' ? '⬆ 没问题，上传这一份（带 PDF）' : '⬆ 没问题，上传这一份';
  }
  window.addEventListener('focus', () => { if (!uploading) showMode(); });  // 在上传页改了选项，回来就更新
  function offerUpload(data) {
    seen = null;
    if ((!uploadTask && !uploadFree) || !uploadRel || !data.sha256 || !data.sha256.mine) return;
    if ($('xdMine').value.trim() !== (params.get('mine') || '') || $('xdTheirs').value.trim() !== (params.get('theirs') || '')) {
      $('xdUploadPanel').hidden = true;  // 路径被手动改过：不是清单里那一份了
      return;
    }
    seen = data.sha256;
    $('xdUploadName').textContent = '上传这一份：' + uploadRel;
    $('xdUploadPanel').hidden = false;
    if (!uploading) { $('xdUpload').disabled = false; $('xdUploadStatus').textContent = ''; showMode(); }
  }
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  async function api(path, body) {
    const response = await fetch(path, body ? { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) } : { cache: 'no-store' });
    const data = await response.json();
    if (!response.ok || data.ok === false) throw new Error(data.error || '请求失败');
    return data;
  }
  // 要带 PDF 时：终点已经有同名 PDF 就问一次要不要覆盖。返回发给后台的 pdf 选项；null = 不带 PDF；false = 用户取消
  async function pdfPlan() {
    const mode = pdfMode();
    if (mode === 'excel') return null;
    const check = await api('/api/distribute/pdf_check', uploadFree ? { preview_id: uploadFree, files: [uploadRel] } : { task_id: uploadTask, files: [uploadRel] });
    const found = (check.files || []).find(file => file.rel === uploadRel);
    const overwrite = {};
    if (found && found.exists) {
      if (found.error) throw new Error('终点的同名 PDF 读不了：' + found.error);
      const yes = await window.appConfirm(`终点已经有同名 PDF：${found.pdf}（${found.mtime || ''}${found.ours ? '，上次就是本工具传的' : '，不是本工具传的'}）。\n\n点「覆盖」用这次导出的换掉它；点「取消」这次先不传。`,
        { title: '同名 PDF 已存在', okText: '覆盖', tone: 'danger' });
      if (!yes) return false;
      overwrite[uploadRel] = found.sha256;
    }
    return { options: pdfOptions(), overwrite, only: mode === 'pdf' };
  }
  function pdfLine(pdf) {
    if (!pdf) return '';
    return pdf.status === 'ok' ? ` 同名 PDF ${pdf.name || ''} 也传了。` : ` 但同名 PDF 没传：${pdf.error || pdf.status || ''}`;
  }
  async function uploadThis() {
    if (!seen || uploading) return;
    uploading = true;
    $('xdUpload').disabled = true;
    $('xdUploadError').textContent = '';
    try {
      $('xdUploadStatus').textContent = pdfMode() === 'excel' ? '正在开始上传…' : '正在看终点有没有同名 PDF…';
      const pdf = await pdfPlan();
      if (pdf === false) { $('xdUploadStatus').textContent = ''; $('xdUpload').disabled = false; return; }
      console.info('📤 对比页：上传这一份', { 文件: uploadRel, 任务: uploadTask, 清单: uploadFree, 传什么: pdfMode() });
      $('xdUploadStatus').textContent = '正在开始上传…';
      if (uploadFree) await api('/api/distribute/free_start', { preview_id: uploadFree, pdf, only: uploadRel, confirmed: seen });
      else await api('/api/distribute/task_upload', { task_id: uploadTask, files: [uploadRel], reviewed: {}, pdf,
        guard: 'auto', confirmed: { [uploadRel]: seen } });
      let status;
      for (;;) {
        await wait(1000);
        status = await api('/api/distribute/status');
        if (!status.running) break;
        $('xdUploadStatus').textContent = (status.reviews || []).length
          ? '这份文件在对比之后又变了，没有直接传。请回上传页，在「等待确认」里重新看差异再确认。'
          : '正在上传：' + (status.current || '准备中…');
      }
      const done = uploadFree ? (status.single && status.single.rel === uploadRel ? status.single : null)
        : ((status.applied || {}).files || []).find(item => item.rel === uploadRel);
      if (done && done.status === 'ok') {
        $('xdUploadStatus').textContent = (done.pdf_only ? '✓ 已上传这一份的 PDF（Excel 没传）。' : '✓ 已上传这一份。' + pdfLine(done.pdf));
        console.info('📤 对比页：上传完成', { 文件: uploadRel, PDF: done.pdf && done.pdf.status });
        return;
      }
      $('xdUploadStatus').textContent = '';
      $('xdUploadError').textContent = '! 这一份没有传成：' + ((done && done.error) || (status.applied || {}).stopped || '请回上传页看原因') + pdfLine(done && done.pdf && done.pdf.status !== 'ok' ? done.pdf : null);
      console.warn('📤 对比页：没传成', { 文件: uploadRel, 结果: done });
    } catch (error) {
      $('xdUploadStatus').textContent = '';
      $('xdUploadError').textContent = error.message;
      $('xdUpload').disabled = false;
    } finally {
      uploading = false;
    }
  }
  $('xdUpload').addEventListener('click', uploadThis);

  const token = params.get('review');
  async function reviewRequest(action, fields = {}) {
    const response = await fetch('/api/distribute/upload_' + action, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ token, ...fields }),
    });
    const data = await response.json();
    if (!response.ok || data.ok === false) throw new Error(data.error || '上传确认失败');
    return data;
  }

  async function loadReview() {
    $('xdInputPanel').hidden = true;
    $('xdReviewPanel').hidden = false;
    $('xdResultPanel').hidden = true;
    $('xdReviewStatus').textContent = '正在读取后台已完成的对比…';
    try {
      const data = await reviewRequest('review');
      $('xdReviewName').textContent = data.rel;
      $('xdReviewStatus').textContent = data.reason;
      if (data.diff) {
        render(data.diff);
        $('xdResultPanel').hidden = false;
      } else {
        $('xdApprove').textContent = '确认重新创建并上传';
      }
      if (data.pdf_only) $('xdApprove').textContent = '确认上传这份 PDF';
      $('xdReviewActions').hidden = false;
    } catch (error) {
      $('xdReviewStatus').textContent = '';
      $('xdReviewError').textContent = error.message;
    }
  }

  async function decideReview(decision) {
    const buttons = [$('xdApprove'), $('xdSkip')];
    buttons.forEach(button => { button.disabled = true; });
    $('xdReviewError').textContent = '';
    try {
      await reviewRequest('decision', { decision });
      $('xdReviewActions').hidden = true;
      $('xdReviewStatus').textContent = decision === 'upload' ? '已进入上传队列。请回到上传页查看结果。' : '已跳过这份。';
    } catch (error) {
      $('xdReviewError').textContent = error.message;
      buttons.forEach(button => { button.disabled = false; });
    }
  }

  if (token) {
    $('xdApprove').addEventListener('click', () => decideReview('upload'));
    $('xdSkip').addEventListener('click', () => decideReview('skip'));
    loadReview();
    return;
  }
  $('xdRun').addEventListener('click', run);
  for (const id of ['xdMine', 'xdTheirs']) $(id).addEventListener('keydown', event => { if (event.key === 'Enter') run(); });
  if ($('xdMine').value && $('xdTheirs').value) run();
})();
