(() => {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
  const api = '/api/bl_links';
  const FOLDER = 'clear_handle_excel.lastFolder';
  let latest = {}, submitting = false, loading = true, timer, initial = true, pollError = '';
  let resultFilter = 'all';

  try { $('blFolder').value = localStorage.getItem(FOLDER) || ''; } catch (_) {}

  function hitsOf(files) {
    return (files || []).filter(f => !f.error).flatMap(f => f.hits || f.cells || []);
  }

  function controls() {
    const busy = loading || submitting || latest.running;
    $('blFolder').disabled = !!busy;
    $('blScan').disabled = !!busy;
    const data = latest.applied || latest.result;
    const canClear = !busy && !latest.applied && latest.result && hitsOf(latest.result.files).length > 0;
    $('blClear').disabled = !canClear;
    $('blScan').textContent = latest.running && latest.phase === 'scan' ? '正在扫描…' : '🔎 扫描 BL 超链接';
    $('blClear').textContent = latest.running && latest.phase === 'clear' ? '正在清除…' : '🧽 清除这些单元格';
  }

  function fileTable(file, data, applied) {
    const cells = applied ? (file.cells || []) : (file.hits || []);
    const target = { folder: data.folder, file: file.name };
    const rows = cells.map(cell => {
      const location = window.excelLinkHtml(esc(`${cell.sheet}!${cell.addr}`), { ...target, sheet: cell.sheet, cell: cell.anchor || cell.addr.split(':')[0] });
      const left = window.excelLinkHtml(esc(`${cell.left_addr} · ${cell.left_value || ''}`), { ...target, sheet: cell.sheet, cell: cell.left_addr });
      if (applied) {
        const ok = !file.error && cell.status === 'ok';
        const reason = file.error ? `未清除：${file.error}` : (cell.reason || '未清除');
        return `<tr><td>${location}</td><td>${left}</td><td class="${ok ? 'link-ok' : 'link-skip'}">${esc(reason)}</td></tr>`;
      }
      return `<tr><td>${location}</td><td>${left}</td><td><code>${esc(cell.value || '')}</code></td><td><code>${esc(cell.link_address || '')}</code></td></tr>`;
    });
    const columns = applied ? '<th>位置</th><th>左边单元格（含 BL）</th><th>处理结果</th>'
      : '<th>位置</th><th>左边单元格（含 BL）</th><th>内容</th><th>链接地址</th>';
    const fileLink = window.excelLinkHtml(esc(file.name), target);
    return `<details class="link-file"${cells.length || file.error ? ' open' : ''}><summary>${fileLink} · ${cells.length} 处${file.error ? ' · 处理失败' : ''}</summary>`
      + (file.error ? `<p class="link-error">${esc(file.error)}</p>` : '')
      + (rows.length ? `<div class="link-table-wrap"><table class="link-table"><thead><tr>${columns}</tr></thead><tbody>${rows.join('')}</tbody></table></div>` : '<p>没有需要处理的问题。</p>')
      + '</details>';
  }

  function renderSummary(files, applied) {
    const failed = files.filter(f => f.error);
    const ok = files.filter(f => !f.error);
    const chips = [];
    if (applied) {
      const cleared = ok.flatMap(f => f.cells || []).filter(c => c.status === 'ok').length;
      chips.push({ cls: 'ok', text: `已清除 ${cleared} 处` });
    } else {
      const hitFiles = ok.filter(f => f.hits && f.hits.length);
      const total = hitFiles.flatMap(f => f.hits).length;
      chips.push({ cls: 'hit', text: `命中 ${total} 处` });
      chips.push({ cls: 'hit', text: `${hitFiles.length} 份文件` });
    }
    if (failed.length) chips.push({ cls: 'err', text: `处理失败 ${failed.length} 份` });
    const empty = ok.filter(f => !(applied ? (f.cells || []).length : (f.hits || []).length)).length;
    if (empty) chips.push({ cls: 'empty', text: `无问题 ${empty} 份` });
    return `<div class="bl-summary">${chips.map(c => `<span class="bl-chip bl-chip-${c.cls}">${esc(c.text)}</span>`).join('')}</div>`;
  }

  function renderFiles(files, data, applied) {
    if (!files.length) return '<p>没有待处理的文件。</p>';
    const failed = files.filter(f => f.error);
    const ok = files.filter(f => !f.error);
    const has = f => (applied ? (f.cells || []).length : (f.hits || []).length) > 0;
    const groups = [
      [applied ? '已处理的文件' : '发现问题的文件', ok.filter(has)],
      ['无问题的文件', ok.filter(f => !has(f))],
      ['处理失败的文件', failed],
    ];
    const html = [];
    for (const [title, group] of groups) {
      if (!group.length) continue;
      html.push(`<h3 class="bl-group-title">${esc(title)} · ${group.length} 份</h3>`);
      html.push(group.map(f => fileTable(f, data, applied)).join(''));
    }
    return html.join('');
  }

  function render() {
    controls();
    if (latest.running) {
      const phase = latest.phase === 'clear' ? '清除' : '扫描';
      $('blProgress').textContent = `正在${phase} ${Math.min((latest.done || 0) + 1, latest.total || 0)} / ${latest.total || 0}：${latest.current || '准备中…'}`;
      $('blProgressFill').style.width = `${latest.total ? latest.done / latest.total * 100 : 0}%`;
      $('blResults').textContent = `正在${phase}，完成后显示明细。`;
      return;
    }
    const applied = latest.applied;
    const data = applied || latest.result;
    if (!data) return;
    const files = data.files;
    const total = hitsOf(files).length;
    const cleared = files.filter(f => !f.error).flatMap(f => f.cells || []).filter(c => c.status === 'ok').length;
    const errors = files.filter(f => f.error).length;
    $('blResultTitle').textContent = applied ? '清除结果' : '扫描结果';
    $('blResultFolder').textContent = `当前文件夹：${data.folder}`;
    $('blProgress').textContent = applied
      ? `已清除 ${cleared} 处；${errors} 份文件处理失败。`
      : `发现 ${total} 处待清除。核对无误后点「🧽 清除这些单元格」。`;
    $('blProgressFill').style.width = '100%';
    const has = file => (applied ? (file.cells || []) : (file.hits || [])).length > 0;
    const groups = [['all', '全部', files], ['hit', applied ? '已处理' : '发现问题', files.filter(f => !f.error && has(f))],
      ['empty', '无问题', files.filter(f => !f.error && !has(f))], ['failed', '处理失败', files.filter(f => f.error)]];
    const shown = groups.find(([key]) => key === resultFilter)[2];
    const filters = '<div class="repair-filter-buttons" role="group" aria-label="BL 超链接结果分类">'
      + groups.map(([key, label, group]) => `<button type="button" class="btn ${key === resultFilter ? 'btn-primary' : 'btn-outline'} btn-sm" data-bl-filter="${key}" aria-pressed="${key === resultFilter}">${label}（${group.length} 份）</button>`).join('')
      + '</div><p class="toolbar-hint">筛选只影响显示，清除范围仍为全部扫描结果。</p>';
    $('blResults').innerHTML = renderSummary(files, applied) + filters + renderFiles(shown, data, applied);
  }

  async function refresh() {
    clearTimeout(timer);
    let failed = false;
    try {
      const response = await fetch(api, { cache: 'no-store' });
      if (!response.ok) throw Error('读取任务进度失败，正在重试');
      latest = await response.json();
      if (pollError && $('blError').textContent === pollError) $('blError').textContent = '';
      pollError = '';
      const data = latest.applied || latest.result;
      if (latest.running || initial) {
        $('blFolder').value = latest.folder || data?.folder || $('blFolder').value;
      }
      initial = false;
      loading = false;
      render();
    } catch (error) { failed = true; pollError = error.message; $('blError').textContent = pollError; }
    if (latest.running || failed) timer = setTimeout(refresh, failed ? 1800 : 900);
  }

  async function post(path, payload) {
    const response = await fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload || {}) });
    const result = await response.json();
    if (!response.ok || !result.ok) throw Error(result.error || '操作失败');
    return result;
  }

  async function scan() {
    if (loading || submitting || latest.running) return;
    const folder = $('blFolder').value.trim();
    $('blError').textContent = '';
    if (!folder) { $('blError').textContent = '请填写 Excel 文件夹路径'; return; }
    submitting = true; controls();
    try {
      await post(api + '/scan', { folder });
      resultFilter = 'all';
      try { localStorage.setItem(FOLDER, folder); } catch (_) {}
      await refresh();
    } catch (error) {
      $('blError').textContent = error.message;
      await refresh();
    } finally { submitting = false; controls(); }
  }

  async function clear() {
    if (loading || submitting || latest.running) return;
    const files = latest.result?.files || [];
    const total = hitsOf(files).length;
    if (!total) return;
    const fileCount = files.filter(f => !f.error && f.hits && f.hits.length).length;
    const ok = await window.appConfirm(
      `将删除 ${total} 处超链接并清空内容（共 ${fileCount} 份文件）。\n这是不可恢复操作，请确认已在扫描结果里核对无误。`,
      { title: '清除 BL 超链接', okText: '确认清除', tone: 'danger' });
    if (!ok) return;
    submitting = true; controls();
    try {
      await post(api + '/clear', {});
      resultFilter = 'all';
      await refresh();
    } catch (error) {
      $('blError').textContent = error.message;
      await refresh();
    } finally { submitting = false; controls(); }
  }

  $('blScan').addEventListener('click', scan);
  $('blResults').addEventListener('click', event => {
    const button = event.target.closest('[data-bl-filter]');
    if (!button) return;
    resultFilter = button.dataset.blFilter;
    render();
    $('blResults').querySelector('[aria-pressed="true"]')?.focus();
  });
  $('blClear').addEventListener('click', clear);
  $('blFolder').addEventListener('keydown', e => { if (e.key === 'Enter') scan(); });
  controls();
  refresh();
})();
