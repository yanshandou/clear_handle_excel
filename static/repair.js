(() => {
  const $ = id => document.getElementById(id);
  const escape = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[char]));
  const colors = document.body.dataset.repairKind === 'colors';
  const hidden = document.body.dataset.repairKind === 'hidden';
  const kind = hidden ? 'hidden' : colors ? 'colors' : 'links';
  const api = hidden ? '/api/hidden_repair' : colors ? '/api/batch/find_marks' : '/api/link_repair';
  const FOLDER = 'clear_handle_excel.lastFolder', OLD = 'clear_handle_excel.oldFolder';
  const label = hidden ? '🙈 扫描并修复隐藏页' : colors ? '🎨 扫描并修复颜色' : '🔗 扫描并修复引用错误';
  const optionIds = {sheets:'repairSheets', columns:'repairColumns', narrow:'repairNarrow'};
  let latest = {}, submitting = false, loading = true, timer, initial = true, pollError = '';
  const successful = cell => ['ok', 'resolved'].includes(cell.status);
  try { $('repairFolder').value = localStorage.getItem(FOLDER) || ''; $('repairOldFolder').value = localStorage.getItem(OLD) || ''; } catch (_) {}

  function palette() {
    const palette = savedHighlightColors();
    if (colors) {
      $('repairColorsSummary').textContent = `本次查找 ${palette.length} 组标记配色`;
      $('repairColors').innerHTML = palette.map(color => `<span class="repair-color" style="background:${color.fill};color:${color.font}">底 ${color.fill} / 字 ${color.font}</span>`).join('');
    }
    return palette;
  }
  palette();

  function controls() {
    const busy = loading || submitting || latest.running;
    for (const id of ['repairFolder', 'repairOldFolder', 'repairRun']) $(id).disabled = !!busy;
    if (hidden) for (const id of Object.values(optionIds)) $(id).disabled = !!busy;
    $('repairRun').textContent = submitting || latest.running ? '正在扫描并修复…' : label;
  }

  function swatch(fill, font) {
    if (fill === 'none') return '无填充';
    if (fill === 'other') return '主题色 / 图案';
    const color = normalizeHighlight({fill, font: font || '#000000'});
    return color ? `<span class="repair-color" style="background:${color.fill};color:${color.font}">Aa</span> ${escape(fill)}${font ? ` / ${escape(font)}` : ''}` : escape(fill || '—');
  }

  function table(file, data, phase) {
    const cells = file.cells || file.hits || [];
    const target = {folder: data.folder, file: file.name, ...(file.old ? {compare: kind} : {})};
    const rows = cells.map(cell => {
      const location = window.excelLinkHtml(escape(`${cell.sheet}!${cell.addr}`), {...target, sheet: cell.sheet, cell: cell.anchor || cell.addr.split(':')[0]});
      const ok = phase === 'apply' && !file.error && successful(cell);
      const reason = file.error ? `未修复：${file.error}` : cell.reason || (phase === 'scan' ? '已发现，等待重新扫描并修复' : '未修复');
      return `<tr><td>${location}</td>` + (hidden
        ? `<td>${escape({sheet:'工作表', columns:'隐藏列', narrow:'超窄列'}[cell.kind] || '待核对')}</td><td>${escape(cell.current_state)}</td><td>${escape(cell.old_state)}${cell.old_sheet ? `<br>${escape(cell.old_sheet)}!${escape(cell.addr)}` : ''}</td>`
        : colors
        ? `<td>${swatch(cell.fill, cell.font)}</td><td>${swatch(cell.old_fill, cell.old_font)}${cell.old_sheet ? `<br>${escape(cell.old_sheet)}!${escape(cell.addr)}` : ''}</td>`
        : `<td><code>${escape(cell.formula || '固定错误值')}</code></td><td><code>${escape(cell.source_formula)}</code><br>${escape(cell.source)}</td><td>${escape(cell.display === '' ? JSON.stringify(cell.value ?? '') : cell.display)}</td>`)
        + `<td class="${ok ? 'link-ok' : 'link-skip'}">${escape(reason)}</td></tr>`;
    });
    const columns = hidden ? '<th>位置</th><th>恢复类型</th><th>修复前状态</th><th>备份状态与位置</th><th>处理结果</th>'
      : colors ? '<th>位置</th><th>修复前配色</th><th>备份配色与位置</th><th>处理结果</th>'
      : '<th>位置</th><th>修复前公式</th><th>备份公式与位置</th><th>恢复值</th><th>处理结果</th>';
    const fileLink = window.excelLinkHtml(escape(file.name), target);
    return `<details class="link-file"${cells.length || file.error ? ' open' : ''}><summary>${fileLink} · ${cells.length} 处${file.error ? ' · 处理失败' : ''}</summary>`
      + (file.old ? `<p class="repair-pair-action">${window.excelLinkHtml('↔ 同时打开当前与备份', target)}<span>左侧当前文件 · 右侧备份（只读）</span></p><p class="repair-backup-path">备份：${window.excelLinkHtml(escape(file.old), target)}</p>` : '')
      + (file.error ? `<p class="link-error">${escape(file.error)}</p>` : '')
      + (file.truncated ? '<p class="link-skip">本文件只检查前 500 处标记；处理完成后可再次扫描，仍未修复的格子请人工检查。</p>' : '')
      + (rows.length ? `<div class="link-table-wrap"><table class="link-table"><thead><tr>${columns}</tr></thead><tbody>${rows.join('')}</tbody></table></div>` : '<p>没有需要处理的问题。</p>') + '</details>';
  }

  function render() {
    controls();
    if (latest.running) {
      const stage = latest.phase === 'repair' ? latest.stage : latest.phase;
      $('repairProgress').textContent = `${{scan:'正在扫描', find:'正在扫描', preview:'正在对比备份', apply:'正在修复'}[stage] || '正在处理'} ${Math.min((latest.done || 0) + 1, latest.total || 0)} / ${latest.total || 0}：${latest.current || '准备中…'}`;
      $('repairProgressFill').style.width = `${latest.total ? latest.done / latest.total * 100 : 0}%`;
      $('repairResults').textContent = '正在扫描并修复，完成后显示明细。';
      return;
    }
    const phase = latest.applied ? 'apply' : latest.plan ? 'preview' : 'scan';
    const data = latest.applied || latest.plan || latest.result;
    if (!data) return;
    const cells = data.files.flatMap(file => file.cells || file.hits || []);
    const errors = data.files.filter(file => file.error).length;
    const count = data.files.filter(file => !file.error).flatMap(file => file.cells || []).filter(successful).length;
    $('repairResultTitle').textContent = hidden ? '隐藏恢复结果' : colors ? '颜色修复结果' : '引用错误修复结果';
    $('repairResultFolder').textContent = `当前文件夹：${data.folder}${data.old_folder ? `；原始备份：${data.old_folder}` : ''}`;
    $('repairProgress').textContent = phase === 'apply'
      ? `已${hidden ? '恢复隐藏设置' : colors ? '恢复颜色' : '恢复正常'} ${count} 处；未修复 ${cells.length - count} 处；${errors} 份文件处理失败。${hidden ? '连续列按列段计数。' : ''}`
      : '上次检查尚未修复，点击“扫描并修复”即可重新检查并处理。';
    $('repairProgressFill').style.width = '100%';
    $('repairResults').innerHTML = data.files.map(file => table(file, data, phase)).join('') || '<p>没有待处理的文件。</p>';
  }

  async function refresh() {
    clearTimeout(timer);
    let failed = false;
    try {
      const response = await fetch(api, {cache: 'no-store'});
      if (!response.ok) throw Error('读取任务进度失败，正在重试');
      latest = await response.json();
      if (pollError && $('repairError').textContent === pollError) $('repairError').textContent = '';
      pollError = '';
      const data = latest.applied || latest.plan || latest.result;
      if (latest.running || initial) {
        $('repairFolder').value = latest.folder || data?.folder || $('repairFolder').value;
        $('repairOldFolder').value = latest.old_folder || data?.old_folder || $('repairOldFolder').value;
        if (hidden) {
          const options = latest.options || data?.options;
          if (options) for (const [key, id] of Object.entries(optionIds)) $(id).checked = options[key] !== false;
        }
      }
      initial = false;
      loading = false;
      render();
    } catch (error) { failed = true; pollError = error.message; $('repairError').textContent = pollError; }
    if (latest.running || failed) timer = setTimeout(refresh, failed ? 1800 : 900);
  }

  async function run() {
    if (loading || submitting || latest.running) return;
    const folder = $('repairFolder').value.trim(), old_folder = $('repairOldFolder').value.trim();
    $('repairError').textContent = '';
    if (!folder || !old_folder) { $('repairError').textContent = '请填写当前 Excel 文件夹和原始备份文件夹'; return; }
    submitting = true; controls();
    try {
      const payload = {folder, old_folder};
      if (colors) payload.colors = palette();
      if (hidden) {
        payload.options = Object.fromEntries(Object.entries(optionIds).map(([key, id]) => [key, $(id).checked]));
        if (!Object.values(payload.options).some(Boolean)) throw Error('请至少选择一种要恢复的隐藏状态');
      }
      const response = await fetch(api + '/run', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)});
      const result = await response.json();
      if (!response.ok || !result.ok) throw Error(result.error || '操作失败');
      try { localStorage.setItem(FOLDER, folder); localStorage.setItem(OLD, old_folder); } catch (_) {}
      await refresh();
    } catch (error) {
      $('repairError').textContent = error.message;
      await refresh();
    } finally { submitting = false; controls(); }
  }
  $('repairRun').addEventListener('click', run);
  controls();
  refresh();
})();
