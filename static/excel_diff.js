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

  function sheetHtml(sheet, data) {
    const doomed = sheet.doomed ? `<span class="xd-badge">精简会删（${esc(sheet.doomed)}）</span>` : '';
    if (sheet.status === 'only_old' || sheet.status === 'only_new') {
      const old = sheet.status === 'only_old';
      const cls = sheet.doomed ? 'is-doomed' : old ? 'is-old' : 'is-new';
      const text = old ? '只有远程有，你这份没有' + (sheet.doomed ? '（精简注定会删，不算差异）' : '（上传后这张表就没了）')
        : '只有你这份有，远程没有' + (sheet.doomed ? '（精简注定会删，不算差异）' : '（上传后会多出这张表）');
      return `<details class="xd-sheet ${cls}"><summary>${old ? '−' : '+'} ${link(sheet.name, old ? data.theirs : data.mine, sheet.name)}`
        + `<span class="xd-badge">${text}</span>${doomed}</summary></details>`;
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
    return `<details class="xd-sheet is-changed" open><summary>± ${link(sheet.name, data.mine, sheet.name)}${badges}${doomed}</summary>`
      + `<div class="xd-body"><table class="xd-table"><colgroup><col class="c-where"><col class="c-kind"><col><col></colgroup>`
      + `<thead><tr><th>位置（点了在你这份里打开）</th><th>类型</th><th>− 远程现在这份（旧）</th><th>+ 你要上传的这份（新）</th></tr></thead>`
      + `<tbody>${rows.join('')}</tbody></table></div></details>`;
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
  $('xdRun').addEventListener('click', run);
  for (const id of ['xdMine', 'xdTheirs']) $(id).addEventListener('keydown', event => { if (event.key === 'Enter') run(); });
  if ($('xdMine').value && $('xdTheirs').value) run();
})();
