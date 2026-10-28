(() => {
  const $ = id => document.getElementById(id);
  const escape = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[char]));
  const FOLDER = 'clear_handle_excel.lastFolder', OLD = 'clear_handle_excel.oldFolder';
  let latest = {}, submitting = false, timer;
  try { $('linkFolder').value = localStorage.getItem(FOLDER) || ''; $('linkOldFolder').value = localStorage.getItem(OLD) || ''; } catch (_) {}

  function controls() {
    const busy = submitting || latest.running;
    for (const id of ['linkFolder', 'linkOldFolder', 'linkScan']) $(id).disabled = !!busy;
    const hits = latest.result?.files.some(file => !file.error && file.hits?.length);
    const count = latest.plan?.files.filter(file => !file.error).flatMap(file => file.cells || []).filter(cell => cell.action === 'restore').length || 0;
    $('linkPreview').disabled = !!busy || !hits;
    $('linkApply').disabled = !!busy || !count;
    $('linkApply').textContent = count ? `③ 确认修复 ${count} 处` : '③ 确认修复';
  }

  function table(file, data, phase) {
    const cells = phase === 'scan' ? file.hits || [] : file.cells || [];
    const rows = cells.map(cell => {
      const location = window.excelLinkHtml(escape(`${cell.sheet}!${cell.addr}`), {folder: data.folder, file: file.name, sheet: cell.sheet, cell: cell.anchor || cell.addr.split(':')[0]});
      const ok = cell.action === 'restore' || ['ok', 'resolved'].includes(cell.status);
      return `<tr><td>${location}</td><td><code>${escape(cell.formula || '固定错误值')}</code></td>`
        + (phase === 'scan' ? `<td>${escape(cell.error)}</td>` : `<td><code>${escape(cell.source_formula)}</code><br>${escape(cell.source)}</td><td>${escape(cell.display === '' ? JSON.stringify(cell.value ?? '') : cell.display)}</td><td class="${ok ? 'link-ok' : 'link-skip'}">${escape(cell.reason)}</td>`) + '</tr>';
    });
    const columns = phase === 'scan' ? '<th>位置</th><th>当前公式</th><th>问题</th>' : '<th>位置</th><th>修复前公式</th><th>备份公式与位置</th><th>恢复值</th><th>处理结果</th>';
    const fileLink = window.excelLinkHtml(escape(file.name), {folder: data.folder, file: file.name});
    return `<details class="link-file"${cells.length || file.error ? ' open' : ''}><summary>${fileLink} · ${cells.length} 处${file.error ? ' · 处理失败' : ''}</summary>`
      + (file.old ? `<p>备份：${escape(file.old)}</p>` : '')
      + (file.error ? `<p class="link-error">${escape(file.error)}</p>` : '')
      + (rows.length ? `<div class="link-table-wrap"><table class="link-table"><thead><tr>${columns}</tr></thead><tbody>${rows.join('')}</tbody></table></div>` : '<p>没有需要处理的引用错误。</p>') + '</details>';
  }

  function render() {
    controls();
    if (latest.running) {
      $('linkProgress').textContent = `${{scan:'正在扫描', preview:'正在对比', apply:'正在修复'}[latest.phase]} ${Math.min((latest.done || 0) + 1, latest.total)} / ${latest.total}：${latest.current || '准备中…'}`;
      return;
    }
    const phase = latest.applied ? 'apply' : latest.plan ? 'preview' : 'scan';
    const data = latest.applied || latest.plan || latest.result;
    if (!data) return;
    $('linkFolder').value = data.folder;
    const hits = data.files.flatMap(file => phase === 'scan' ? file.hits || [] : file.cells || []);
    const errors = data.files.filter(file => file.error).length;
    const count = phase === 'scan' ? hits.length : hits.filter(cell => phase === 'preview' ? cell.action === 'restore' : ['ok','resolved'].includes(cell.status)).length;
    $('linkResultTitle').textContent = {scan:'扫描结果', preview:'备份对比：确认后才修改', apply:'引用错误修复结果'}[phase];
    $('linkProgress').textContent = `${{scan:'发现引用错误', preview:'可以自动修复', apply:'已恢复正常'}[phase]} ${count} 处；${errors} 份文件处理失败。`;
    $('linkResults').innerHTML = data.files.map(file => table(file, data, phase)).join('') || '<p>没有待处理的文件。</p>';
  }

  async function refresh() {
    clearTimeout(timer);
    try {
      const response = await fetch('/api/link_repair', {cache: 'no-store'});
      if (!response.ok) throw Error('读取任务进度失败');
      latest = await response.json();
      if (latest.result && !$('linkFolder').value) $('linkFolder').value = latest.result.folder;
      render();
    } catch (error) { $('linkError').textContent = error.message; }
    if (latest.running) timer = setTimeout(refresh, 900);
  }

  async function action(phase) {
    if (submitting || latest.running) return;
    $('linkError').textContent = '';
    const folder = $(phase === 'preview' ? 'linkOldFolder' : 'linkFolder').value.trim();
    if (phase !== 'apply' && !folder) { $('linkError').textContent = '请填写文件夹路径'; return; }
    submitting = true; controls();
    try {
      const response = await fetch('/api/link_repair' + ({scan:'', preview:'/preview', apply:'/apply'}[phase]), {
        method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(phase === 'apply' ? {} : {folder})
      });
      const result = await response.json();
      if (!response.ok || !result.ok) throw Error(result.error || '操作失败');
      try { if (phase === 'scan') localStorage.setItem(FOLDER, folder); if (phase === 'preview') localStorage.setItem(OLD, folder); } catch (_) {}
      await refresh();
    } catch (error) { $('linkError').textContent = error.message; }
    finally { submitting = false; controls(); }
  }
  $('linkScan').addEventListener('click', () => action('scan'));
  $('linkPreview').addEventListener('click', () => action('preview'));
  $('linkApply').addEventListener('click', () => action('apply'));
  // 清单绑定的是扫描时的路径；更改路径后须重新扫描，不能将旧清单误用于新文件夹。
  $('linkFolder').addEventListener('input', () => { latest = {}; controls(); $('linkResults').textContent = '路径已改变，请重新扫描。'; });
  $('linkOldFolder').addEventListener('input', () => { delete latest.plan; delete latest.applied; controls(); });
  refresh();
})();
