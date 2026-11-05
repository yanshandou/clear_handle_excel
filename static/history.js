(() => {
  const $ = id => document.getElementById(id);
  const esc = text => String(text ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const statuses = { idle:'尚未开始', running:'执行中', interrupted:'已暂停 / 中断', completed:'已完成', failed:'失败', error:'错误', cancelled:'已取消', done:'完成', needs_fix:'待处理', pending:'待执行', processing:'处理中' };
  const status = value => esc(statuses[value] || value || '尚未开始');
  const norm = path => String(path || '').replace(/\\/g, '/').replace(/\/+$/, '').toLowerCase();
  let data = { folders:[] }, tasks = { tasks:[] }, selected = '', sequence = 0, restoring = false;
  async function json(url, payload) {
    const response = await fetch(url, payload ? { method:'POST', headers:{ 'Content-Type':'application/json' }, body:JSON.stringify(payload) } : { cache:'no-store' });
    const result = await response.json();
    if (!response.ok || result.ok === false) throw Error(result.error || `请求失败：${response.status}`);
    return result;
  }
  function renderFolders() {
    const query = $('filter').value.trim().toLowerCase();
    const rows = data.folders.filter(row => row.folder.toLowerCase().includes(query));
    $('folders').innerHTML = rows.length ? rows.map(row => `<button class="history-folder" data-folder="${esc(row.folder)}" aria-pressed="${row.folder === selected}"><b>${esc(row.folder)}</b><small>${status(row.status)} · 已检查 ${row.files || 0} 个文件 · 问题卡片 ${row.issues || 0}${norm(row.folder) === norm(data.current) ? ' · 当前工作台' : ''}</small><small>最近保存：${esc(row.saved_at || '—')}</small></button>`).join('') : '<p>没有匹配的已保存进度。</p>';
  }
  function table(head, rows) {
    return `<div class="history-scroll"><table class="records-table"><thead><tr>${head.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.length ? rows.join('') : `<tr><td colspan="${head.length}">暂无记录</td></tr>`}</tbody></table></div>`;
  }
  function fileDetails(file) {
    const notes = [];
    if (file.deleted_backups?.length) notes.push(`已删除表：${esc(file.deleted_backups.join('、'))}`);
    for (const change of file.sheet_changes || []) {
      for (const [from, to] of Object.entries(change.renamed || {})) notes.push(`表改名：${esc(from)} → ${esc(to)}`);
    }
    if (file.exported_pdfs) notes.push(`已导出 ${Number(file.exported_pdfs)} 份 PDF`);
    if (file.normalized_a1) notes.push('已做视图标准化');
    if (file.clip_detected) notes.push(`检出 ${Number(file.clip_detected)} 处 · 已修 ${Number(file.clip_fixed || 0)} 处 · 人工确认无需修改 ${Number(file.clip_accepted || 0)} 处`);
    const fixes = Object.values(file.fixes || {}).map(fix => {
      const links = ['before', 'after'].filter(part => fix[part]).map(part => `<a href="/api/batch/image/${encodeURIComponent(fix[part])}" target="_blank" rel="noopener">${part === 'before' ? '修改前' : '修改后'}</a>`).join(' · ');
      return `<li>${esc(fix.sheet)}!${esc(fix.cell || fix.row)}：${esc(fix.reason)} ${links}</li>`;
    });
    return notes.map(note => `<div>${note}</div>`).join('') + (fixes.length ? `<details><summary>修复位置与截图（${fixes.length} 处）</summary><ul>${fixes.join('')}</ul></details>` : '');
  }
  async function show(folder) {
    selected = folder;
    renderFolders();
    const version = ++sequence;
    $('detail').innerHTML = '<p>正在读取记录…</p>';
    try {
      const value = await json(`/api/folders/detail?folder=${encodeURIComponent(folder)}`);
      if (version !== sequence) return;
      const report = value.report || {}, summary = value.summary || {};
      const matching = tasks.tasks.find(task => norm(task.work) === norm(folder));
      const allowed = matching || tasks.single || norm(data.current) === norm(folder);
      const files = Object.values(report.files || {});
      const operations = (value.operations || []).map(row => `<tr><td>${esc(row.started_at)}</td><td>${esc(row.name)}</td><td>${status(row.status)}</td></tr>`);
      const scans = (value.scans || []).map(row => `<tr><td>${esc(row.started_at)}</td><td>${esc(row.mode)}</td><td>${status(row.status)}</td><td><a href="/analysis?id=${encodeURIComponent(row.id)}&folder=${encodeURIComponent(folder)}">查看逐文件扫描分析</a></td></tr>`);
      const fileRows = files.map(file => `<tr><td>${esc(file.name)}</td><td>${status(file.status)}</td><td>${esc(file.current_action)}${fileDetails(file)}</td><td>${esc(file.error || '')}</td></tr>`);
      $('detail').innerHTML = `<h2 class="history-path">${esc(folder)}</h2><p>${status(summary.status)} · ${esc(summary.task_name)}${summary.total ? ` · 进度 ${summary.completed} / ${summary.total} 个文件` : ''}</p><button id="restore" class="btn btn-primary" ${data.busy || !allowed ? 'disabled' : ''}>恢复并继续处理</button><p>${data.busy ? '当前后台任务正在执行，完成或暂停后可恢复。' : !allowed ? '这是单项模式的文件夹，请先到主页打开单项模式，再恢复。' : '恢复仅切换工作台；继续暂停的任务时，请在工作台点「继续上次任务」。'}</p><h3>最近一次执行的文件明细</h3>${table(['文件','状态','执行信息','错误'], fileRows)}<h3>扫描 / 复检批次</h3>${table(['开始时间','类型','状态','明细'], scans)}<h3>操作时间线（最近 200 次）</h3>${table(['开始时间','操作','状态'], operations)}<details><summary>工作台日志（最近 100 条）</summary><pre class="history-log">${esc((value.logs || []).join('\n'))}</pre></details>`;
      $('restore').addEventListener('click', async () => {
        if (restoring) return;
        restoring = true;
        $('restore').disabled = true;
        try {
          if (matching) {
            const result = await json('/api/tasks/switch', { id:matching.id });
            try { localStorage.setItem('clear_handle_excel.oldFolder', result.task.reference); } catch (_) {}
          } else await json('/api/folders/activate', { folder });
          try { localStorage.setItem('clear_handle_excel.lastFolder', folder); } catch (_) {}
          window.location.href = '/workbench';
        } catch (error) { $('error').textContent = error.message; restoring = false; $('restore').disabled = false; }
      });
    } catch (error) { if (version === sequence) $('detail').textContent = error.message; }
  }
  async function load() {
    try {
      [data, tasks] = await Promise.all([json('/api/folders'), json('/api/tasks')]);
      $('error').textContent = '';
      renderFolders();
      if (selected) await show(selected);
      else if (data.folders.length) await show(data.current || data.folders[0].folder);
    } catch (error) { $('error').textContent = error.message; }
  }
  $('folders').addEventListener('click', event => { const button = event.target.closest('[data-folder]'); if (button) show(button.dataset.folder); });
  $('filter').addEventListener('input', renderFolders);
  $('refresh').addEventListener('click', load);
  load();
})();
