// 首页和工作台共用的 Sheet 打印方案；折叠时不额外轮询。
(() => {
  const panel = document.getElementById('printLayoutsDetails');
  if (!panel) return;
  const list = document.getElementById('printLayoutsList');
  const status = document.getElementById('printLayoutsStatus');
  const escape = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[c]));
  const label = value => value === 2 ? '横向' : '竖向';
  let rows = [], inFlight = false, updating = false, disabled = '';
  const search = document.getElementById('printLayoutsSearch');
  const filter = document.getElementById('printLayoutsFilter');
  const folder = () => (document.getElementById('folderInput') || document.getElementById('taskFolderInput'))?.value.trim()
    || localStorage.getItem('clear_handle_excel.lastFolder') || '';
  const metrics = row => ['1', '2'].map(direction => {
    const m = row.metrics?.[direction];
    return m ? escape(`${label(Number(direction))}：${m.pages} 页 · 字约 ${m.font_pt}pt · 横向分页 ${m.across ?? '未知'} 处 · 合并跨页 ${m.cuts} 处`) : '';
  }).filter(Boolean).join('<br>');
  function renderRows() {
      const term = search.value.trim().toLowerCase();
      const selected = filter.value;
      const shown = rows.map((row, index) => ({row, index})).filter(({row}) =>
        (!term || `${row.file} ${row.sheet} ${row.path}`.toLowerCase().includes(term)) &&
        (selected === 'all' || selected === 'stale' && row.stale || selected === 'review' && row.review_needed
          || selected === 'manual' && row.manual || selected === 'auto' && !row.manual));
      const html = shown.length ? `<table class="print-layout-table"><thead><tr><th>文件 / Sheet</th><th>原方向</th><th>采用方向</th><th>比较与原因</th><th>调整</th></tr></thead><tbody>${shown.map(({row, index}) => `<tr>
        <td title="${escape(row.path)}">${escape(row.file)}<br><strong>${escape(row.sheet)}</strong></td>
        <td>${label(row.original)}</td><td>${label(row.selected)}${row.manual ? '（人工）' : ''}${row.stale ? '<br><span class="print-layout-stale">待重新验证</span>' : row.review_needed ? '<br><span class="print-layout-stale">出现新的合并跨页，建议重新分析</span>' : ''}</td>
        <td>${metrics(row)}<br>${escape(row.reason)}</td><td><select data-layout-index="${index}" aria-label="${escape(row.sheet)} 的打印方向" ${disabled}>
          <option value="recommended" ${!row.manual ? 'selected' : ''}>自动推荐（${label(row.selected)}）</option>
          <option value="portrait" ${row.manual === 1 ? 'selected' : ''}>指定竖向</option>
          <option value="landscape" ${row.manual === 2 ? 'selected' : ''}>指定横向</option></select>
          <button type="button" class="btn btn-outline btn-sm" data-layout-reanalyze="${index}" ${disabled}>重新分析</button></td></tr>`).join('')}</tbody></table>`
        : rows.length ? '<p>没有符合筛选条件的记录。</p>' : '<p>这个文件夹还没有打印方向记录。首次扫描或导出 PDF 时会自动分析并记录。</p>';
      if (list.innerHTML !== html) list.innerHTML = html;
  }
  async function refresh() {
    if (!panel.open || inFlight || updating || list.contains(document.activeElement)) return;
    inFlight = true;
    try {
      const response = await fetch(`/api/print_layouts?folder=${encodeURIComponent(folder())}`, {cache:'no-store'});
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || '读取打印方向失败');
      rows = result.rows;
      disabled = result.busy || result.blocked ? 'disabled' : '';
      renderRows();
      status.textContent = result.blocked || (result.busy ? '任务进行中，完成后可调整方向。' : '方向记录供扫描、复检和批量导出共用；调整后请复检以更新定位。');
    } catch (error) { status.textContent = error.message; }
    finally { inFlight = false; }
  }
  async function update(index, direction) {
    const row = rows[index];
    if (!row || updating) return;
    updating = true;
    status.textContent = direction === 'auto' ? '正在启动方向分析…' : '正在保存打印方向…';
    try {
      const response = await fetch('/api/print_layout', {method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({path:row.path, key:row.key, direction})});
      const result = await response.json();
      if (!response.ok || !result.ok) throw new Error(result.error || '调整失败');
      document.activeElement?.blur();
    } catch (error) { status.textContent = error.message; return; }
    finally { updating = false; }
    await refresh();
  }
  list.addEventListener('change', event => {
    const select = event.target.closest('select[data-layout-index]');
    if (select) update(Number(select.dataset.layoutIndex), select.value === 'recommended' ? 'auto' : select.value);
  });
  search.addEventListener('input', renderRows);
  filter.addEventListener('change', renderRows);
  list.addEventListener('click', event => {
    const button = event.target.closest('[data-layout-reanalyze]');
    if (button) update(Number(button.dataset.layoutReanalyze), 'auto');
  });
  document.getElementById('printLayoutsRefresh').addEventListener('click', () => { document.activeElement?.blur(); refresh(); });
  panel.addEventListener('toggle', refresh);
  setInterval(refresh, 4000);
})();
