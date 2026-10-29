(() => {
  const $ = id => document.getElementById(id);
  const escape = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[char]));
  const colors = document.body.dataset.repairKind === 'colors';
  const hidden = document.body.dataset.repairKind === 'hidden';
  const kind = hidden ? 'hidden' : colors ? 'colors' : 'links';
  const api = hidden ? '/api/hidden_repair' : colors ? '/api/batch/find_marks' : '/api/link_repair';
  const FOLDER = 'clear_handle_excel.lastFolder', OLD = 'clear_handle_excel.oldFolder';
  const label = hidden ? '🙈 扫描并修复隐藏页' : colors ? '🎨 清除标记并修复颜色' : '🔗 扫描并修复引用错误';
  const optionIds = {sheets:'repairSheets', columns:'repairColumns', narrow:'repairNarrow'};
  const displayLimits = new Map();
  const CATEGORIES = [['no_difference','配色无差异'], ['restored_cell','单元格配色已恢复'], ['restored_tab','Sheet 标签颜色已恢复'], ['cleared_tab','Sheet 标签标记色已清除'], ['backup_not_unique','备份文件无法唯一匹配'], ['self_backup','当前文件被用作备份'], ['sheet_unmatched','工作表无法唯一匹配'], ['merged_range_mismatch','合并区域不一致'], ['content_mismatch','单元格内容或公式不一致'], ['empty_row_mismatch','空白格所在行内容不一致'], ['dimension_default','整行整列默认配色需人工处理'], ['gradient','渐变填充需人工处理'], ['template','模板格式不支持自动写回'], ['file_changed','文件在处理中发生变化'], ['file_locked','文件被占用'], ['file_readonly','文件只读'], ['backup_cell_read_failed','备份单元格读取失败'], ['backup_color_unreadable','备份颜色无法读取'], ['cell_verify_failed','单元格颜色恢复校验失败'], ['tab_verify_failed','标签颜色恢复校验失败'], ['other','其他执行异常']];
  const CATEGORY_LABELS = new Map(CATEGORIES);
  const FILE_LEVEL_CODES = new Set(['no_difference', 'backup_not_unique', 'self_backup', 'file_changed', 'file_locked', 'file_readonly', 'tab_verify_failed']);
  let resultFilter = colors ? 'unresolved' : 'all';
  let latest = {}, submitting = false, loading = true, timer, initial = true, pollError = '', clearFilter = 'all';
  const successful = cell => ['ok', 'resolved'].includes(cell.status);
  try { $('repairFolder').value = localStorage.getItem(FOLDER) || ''; $('repairOldFolder').value = localStorage.getItem(OLD) || ''; } catch (_) {}

  function controls() {
    const busy = loading || submitting || latest.running;
    for (const id of ['repairFolder', 'repairOldFolder', 'repairRun']) $(id).disabled = !!busy;
    if (hidden) for (const id of Object.values(optionIds)) $(id).disabled = !!busy;
    $('repairRun').textContent = submitting || latest.running ? (colors ? '正在处理…' : '正在扫描并修复…') : label;
  }

  function swatch(fill, font, tab = false) {
    const text = fill === 'none' ? (tab ? '无标签颜色' : '无填充')
      : fill === 'other' ? (tab ? '主题色' : '主题色 / 图案') : escape(fill || '—');
    const color = normalizeHighlight({fill, font: tab ? fill : font});
    const preview = color ? `<span class="repair-color" style="background:${color.fill}${tab ? '' : `;color:${color.font}`}">${tab ? '&nbsp;' : 'Aa'}</span> ` : '';
    return `${preview}${text}${!tab && font ? ` / 字色 ${escape(font)}` : ''}`;
  }

  function cellCode(cell, phase) {
    if (hidden) {
      if (phase === 'apply' && successful(cell)) return {sheet:'hidden_sheet', columns:'hidden_columns', narrow:'hidden_narrow'}[cell.kind] || 'hidden_restored';
      const reason = cell.reason || '';
      if (reason.includes('无法唯一配对') || reason.includes('缺少内容依据')) return 'hidden_pair';
      if (reason.includes('内容、公式或位置不同')) return 'hidden_content';
      if (reason.includes('受保护')) return 'hidden_protected';
      if (reason.includes('可见工作表')) return 'hidden_visible';
      return 'hidden_other';
    }
    if (!colors) {
      if (phase === 'apply' && successful(cell)) return 'link_restored';
      const reason = cell.reason || '';
      if (reason.includes('工作表无法唯一配对')) return 'link_sheet';
      if (reason.includes('固定的错误值')) return 'link_fixed_error';
      if (reason.includes('附近的内容') || reason.includes('数组或溢出区域')) return 'link_content';
      if (reason.includes('备份同一位置没有公式') || reason.includes('备份的原公式本身')) return 'link_backup_formula';
      if (reason.includes('模拟删除') || reason.includes('模拟删表')) return 'link_mismatch';
      return 'link_other';
    }
    if (phase === 'apply' && successful(cell)) return cell.code || (cell.kind === 'tab' ? 'restored_tab' : 'restored_cell');
    return cell.code || 'other';
  }

  function analyzeFile(file, phase) {
    const cells = file.cells || file.hits || [];
    const coded = cells.map(cell => ({cell, code: !colors && file.error ? 'link_file_error' : cellCode(cell, phase)}));
    const applied = phase === 'apply' && !file.error;
    const pending = coded.filter(item => !(applied && successful(item.cell)));
    const codes = new Set(coded.map(item => item.code));
    const errorCode = file.error ? (colors ? file.code || 'other' : 'link_file_error') : null;
    if (errorCode) codes.add(errorCode);
    const noDifference = !cells.length && !errorCode && !file.truncated;
    if (noDifference) codes.add('no_difference');
    return {file, coded, pending, codes, errorCode, noDifference,
      truncated: !!file.truncated, needsAttention: !!errorCode || !!file.truncated || (phase === 'apply' ? pending.length > 0 : cells.length > 0)};
  }

  function categoryCount(analyses, code) {
    return analyses.reduce((sum, analysis) => sum
      + (code === 'no_difference' ? (analysis.noDifference ? 1 : 0)
        : analysis.coded.filter(item => item.code === code).length + (analysis.errorCode === code ? 1 : 0)), 0);
  }

  function viewEntries(analyses, view) {
    return analyses
      .filter(analysis => view === 'all' ? true : view === 'unresolved' ? analysis.needsAttention : analysis.codes.has(view))
      .map(analysis => ({
        name: analysis.file.name, old: analysis.file.old,
        cells: (view === 'all' ? analysis.coded : view === 'unresolved' ? analysis.pending : analysis.coded.filter(item => item.code === view)).map(item => item.cell),
        error: view === 'all' || view === 'unresolved' ? analysis.file.error : analysis.errorCode === view ? analysis.file.error : null,
        truncated: view === 'all' || view === 'unresolved' ? analysis.truncated : false,
      }));
  }

  function table(file, data, phase, view = 'all') {
    const cells = file.cells || file.hits || [];
    const shown = colors ? cells.slice(0, displayLimits.get(file.name) || 200) : cells;
    const target = {folder: data.folder, file: file.name, ...(file.old ? {compare: kind} : {})};
    const rows = shown.map(cell => {
      const tab = colors && cell.kind === 'tab';
      const location = window.excelLinkHtml(escape(tab ? `${cell.sheet} · Sheet 标签` : `${cell.sheet}!${cell.addr}`), {...target, sheet: cell.sheet, cell: cell.anchor || cell.addr.split(':')[0]});
      const ok = phase === 'apply' && !file.error && successful(cell);
      const reason = file.error ? `未修复：${file.error}` : cell.reason || (phase === 'scan' ? '已发现，等待重新扫描并修复' : '未修复');
      return `<tr><td>${location}</td>` + (hidden
        ? `<td>${escape({sheet:'工作表', columns:'隐藏列', narrow:'超窄列'}[cell.kind] || '待核对')}</td><td>${escape(cell.current_state)}</td><td>${escape(cell.old_state)}${cell.old_sheet ? `<br>${escape(cell.old_sheet)}!${escape(cell.addr)}` : ''}</td>`
        : colors
        ? `<td>${tab ? '标签：' : ''}${swatch(cell.fill, cell.font, tab)}</td><td>${swatch(cell.old_fill, cell.old_font, tab)}${cell.old_sheet ? `<br>${escape(cell.old_sheet)}${tab ? ' · Sheet 标签' : `!${escape(cell.addr)}`}` : ''}</td>`
        : `<td><code>${escape(cell.formula || '固定错误值')}</code></td><td><code>${escape(cell.source_formula)}</code><br>${escape(cell.source)}</td><td>${escape(cell.display === '' ? JSON.stringify(cell.value ?? '') : cell.display)}</td>`)
        + `<td class="${ok ? 'link-ok' : 'link-skip'}">${escape(reason)}</td></tr>`;
    });
    const columns = hidden ? '<th>位置</th><th>恢复类型</th><th>修复前状态</th><th>备份状态与位置</th><th>处理结果</th>'
      : colors ? '<th>位置</th><th>修复前配色</th><th>备份配色与位置</th><th>处理结果</th>'
      : '<th>位置</th><th>修复前公式</th><th>备份公式与位置</th><th>恢复值</th><th>处理结果</th>';
    const fileLink = window.excelLinkHtml(escape(file.name), target);
    const countLabel = !colors ? ` · ${cells.length} 处${file.error ? ' · 处理失败' : ''}`
      : file.error ? `${cells.length ? ` · ${cells.length} 处` : ''} · 处理失败`
      : !cells.length ? ' · 无需处理'
      : view === 'unresolved' ? ` · 未修复 ${cells.length} 处`
      : ` · ${cells.length} 处`;
    return `<details class="link-file"${cells.length || file.error || file.truncated ? ' open' : ''}><summary>${fileLink}${countLabel}</summary>`
      + (file.old ? `<p class="repair-pair-action">${window.excelLinkHtml('↔ 同时打开当前与备份', target)}<span>左侧当前文件 · 右侧备份（只读）</span></p><p class="repair-backup-path">备份：${window.excelLinkHtml(escape(file.old), target)}</p>` : '')
      + (file.error ? `<p class="link-error">${escape(file.error)}</p>` : '')
      + (file.truncated ? '<p class="link-skip">本文件只检查前 500 处标记；处理完成后可再次扫描，仍未修复的格子请人工检查。</p>' : '')
      + (rows.length ? `<div class="link-table-wrap"><table class="link-table"><thead><tr>${columns}</tr></thead><tbody>${rows.join('')}</tbody></table></div>` : colors ? '' : '<p>没有需要处理的问题。</p>')
      + (shown.length < cells.length ? `<p>当前显示${view === 'unresolved' ? '未修复项' : ''} ${shown.length} / ${cells.length} 处。 <button class="btn btn-outline btn-sm" type="button" data-repair-more="${escape(file.name)}">查看更多差异</button></p>` : '') + '</details>';
  }

  function renderClear(data, combined = false) {
    const filters = combined ? 'clearFilters' : 'repairFilters';
    const results = combined ? 'clearResults' : 'repairResults';
    const files = data.files || [];
    const totalRules = files.reduce((sum, file) => sum + (file.leftovers || 0), 0);
    const totalTabs = files.reduce((sum, file) => sum + (file.tabs || []).filter(t => t.status === 'ok').length, 0);
    const failed = file => !!file.error || (file.leftover_failed || []).length > 0 || (file.tabs || []).some(tab => tab.status !== 'ok');
    const errors = files.filter(failed).length;
    const summary = `已清除 ${totalRules} 条标记，还原标签 ${totalTabs} 张；${errors} 份文件存在清除失败。`;
    if (combined) { $('clearResultPanel').hidden = false; $('clearResultSummary').textContent = summary; }
    else { $('repairResultTitle').textContent = '清除标记颜色结果'; $('repairResultFolder').textContent = `当前文件夹：${data.folder}`; $('repairProgress').textContent = summary; $('repairProgressFill').style.width = '100%'; }
    const hasWork = file => !!file.leftovers || (file.tabs || []).some(tab => tab.status === 'ok');
    const categories = [['all', '全部', files], ['clear_done', '已清除', files.filter(hasWork)],
      ['clear_empty', '无需处理', files.filter(file => !hasWork(file) && !failed(file))], ['clear_failed', '处理失败', files.filter(failed)]];
    let selected = combined ? clearFilter : resultFilter;
    if (!categories.some(([key]) => key === selected)) selected = 'all';
    if (combined) clearFilter = selected; else resultFilter = selected;
    $(filters).hidden = false;
    $(filters).innerHTML = '<div class="repair-filter-buttons" role="group" aria-label="清除标记结果分类">'
      + categories.map(([key, label, group]) => `<button type="button" class="btn ${key === selected ? 'btn-primary' : 'btn-outline'} btn-sm" data-repair-filter="${key}" aria-pressed="${key === selected}" aria-controls="${results}">${label}（${group.length} 份）</button>`).join('')
      + '</div><p class="toolbar-hint">分类只影响显示，同一文件可能出现在多个分类中。</p>';
    const shown = categories.find(([key]) => key === selected)[2];
    $(results).innerHTML = shown.map(file => {
      const target = {folder: data.folder, file: file.name};
      const fileLink = window.excelLinkHtml(escape(file.name), target);
      const tabs = file.tabs || [];
      const tabOk = tabs.filter(t => t.status === 'ok').length;
      const tabFailed = tabs.filter(t => t.status !== 'ok');
      const leftoverFailed = file.leftover_failed || [];
      const hasWork = !!(file.error || file.leftovers || tabOk || tabFailed.length || leftoverFailed.length);
      const summary = hasWork
        ? `<p>清掉标记规则 ${file.leftovers || 0} 条${tabOk ? `，还原标签 ${tabOk} 张` : ''}${file.error ? `；${escape(file.error)}` : ''}</p>`
        : '<p>没有程序加的标记，无需处理。</p>';
      const detail = summary
        + (leftoverFailed.length ? `<p class="link-skip">删不掉的残留标记：${escape(leftoverFailed.join('；'))}</p>` : '')
        + (tabFailed.length ? `<p class="link-skip">标签还原失败：${tabFailed.map(t => `${escape(t.sheet)}（${escape(t.reason)}）`).join('；')}</p>` : '');
      return `<details class="link-file"${hasWork ? ' open' : ''}><summary>${fileLink}${file.error ? ' · 处理失败' : (file.leftovers || tabOk) ? ` · 清掉 ${file.leftovers || 0} 条规则${tabOk ? `、${tabOk} 张标签` : ''}` : ' · 无需处理'}</summary>${detail}</details>`;
    }).join('') || '<p>该分类没有结果。</p>';
  }

  function render() {
    controls();
    const phase = latest.applied ? 'apply' : latest.plan ? 'preview' : 'scan';
    const data = latest.applied || latest.plan || latest.result;
    $('repairFilters').hidden = !!latest.running || phase !== 'apply' || !data;
    if (latest.running) {
      const clearing = latest.phase === 'clear' || (colors && latest.stage === 'clear');
      const stage = latest.phase === 'repair' ? latest.stage : latest.phase;
      if (colors) $('clearResultPanel').hidden = true;
      const progressId = 'repairProgress';
      $(progressId).textContent = `${{scan:'正在扫描', find:'正在扫描', preview:'正在对比备份', apply:'正在修复', clear:'正在清除标记'}[stage] || '正在处理'} ${Math.min((latest.done || 0) + 1, latest.total || 0)} / ${latest.total || 0}：${latest.current || '准备中…'}`;
      const fraction = latest.total ? latest.done / latest.total : 0;
      $(progressId + 'Fill').style.width = `${colors && latest.phase === 'repair' ? (clearing ? fraction * 50 : 50 + fraction * 50) : fraction * 100}%`;
      $('repairResults').textContent = clearing ? '正在清除标记，完成后显示明细。' : '正在扫描并修复，完成后显示明细。';
      return;
    }
    if (!data) return;
    if (colors && latest.phase === 'clear') { renderClear(data); return; }
    const cells = data.files.flatMap(file => file.cells || file.hits || []);
    const errors = data.files.filter(file => file.error).length;
    const count = data.files.filter(file => !file.error).flatMap(file => file.cells || []).filter(successful).length;
    const restoredTabs = colors ? data.files.filter(file => !file.error).flatMap(file => file.cells || []).filter(cell => cell.kind === 'tab' && successful(cell)).length : 0;
    $('repairResultTitle').textContent = hidden ? '隐藏恢复结果' : colors ? '颜色修复结果' : '引用错误修复结果';
    $('repairResultFolder').textContent = `当前文件夹：${data.folder}${data.old_folder ? `；原始备份：${data.old_folder}` : ''}`;
    $('repairProgress').textContent = phase === 'apply'
      ? `已${hidden ? '恢复隐藏设置' : colors ? '恢复颜色' : '恢复正常'} ${count} 处；未修复 ${cells.length - count} 处；${errors} 份文件处理失败。${colors && restoredTabs ? `其中 Sheet 标签 ${restoredTabs} 张。` : ''}${hidden ? '连续列按列段计数。' : ''}`
      : '上次检查尚未修复，点击“扫描并修复”即可重新检查并处理。';
    $('repairProgressFill').style.width = '100%';
    let entries, view = 'all', emptyMessage = '没有待处理的文件。';
    if (colors) {
      const analyses = data.files.map(file => analyzeFile(file, phase));
      if (phase === 'apply') {
        if (resultFilter !== 'unresolved' && resultFilter !== 'all' && !analyses.some(analysis => analysis.codes.has(resultFilter))) resultFilter = 'unresolved';
        view = resultFilter;
        const chips = [['unresolved', `未处理好（${analyses.filter(analysis => analysis.needsAttention).length} 份）`], ['all', `全部（${analyses.length} 份）`]]
          .concat(CATEGORIES.filter(([code]) => analyses.some(analysis => analysis.codes.has(code)))
            .map(([code, label]) => [code, `${label}（${categoryCount(analyses, code)} ${FILE_LEVEL_CODES.has(code) ? '份' : '处'}）`]));
        $('repairFilters').innerHTML = '<div class="repair-filter-buttons" role="group" aria-label="颜色修复结果分类">'
          + chips.map(([key, label]) => `<button type="button" class="btn ${key === view ? 'btn-primary' : 'btn-outline'} btn-sm" data-repair-filter="${key}" aria-pressed="${key === view}" aria-controls="repairResults">${label}</button>`).join('')
          + '</div><p class="toolbar-hint">具体分类按本次结果实际出现的原因列出；同一文件可能出现在多个分类中，明细仅显示当前分类的项目。</p>';
      }
      entries = viewEntries(analyses, view);
      emptyMessage = view === 'unresolved' ? '没有需要人工处理的配色问题。' : view === 'all' ? '没有检查结果。' : `没有「${CATEGORY_LABELS.get(view)}」的项目。`;
    } else if (phase === 'apply') {
      const analyses = data.files.map(file => analyzeFile(file, phase));
      const categories = hidden ? [['hidden_sheet', '工作表已恢复'], ['hidden_columns', '隐藏列已恢复'],
        ['hidden_narrow', '窄列宽已恢复'], ['hidden_restored', '其他已恢复'],
        ['hidden_pair', '工作表无法配对'], ['hidden_content', '工作表内容或位置不同'],
        ['hidden_protected', '工作簿或工作表受保护'], ['hidden_visible', '需保留可见工作表'],
        ['hidden_other', '其他未恢复'], ['link_file_error', '文件处理失败'], ['no_difference', '无须恢复']]
        : [['link_restored', '已恢复'], ['link_sheet', '工作表无法配对'],
        ['link_fixed_error', '固定错误值'], ['link_content', '附近内容或区域不同'],
        ['link_backup_formula', '备份公式无法使用'], ['link_mismatch', '引用来源无法确认'],
        ['link_other', '其他未修复'], ['link_file_error', '文件处理失败'], ['no_difference', '无引用错误']];
      if (resultFilter !== 'all' && resultFilter !== 'unresolved' && !analyses.some(analysis => analysis.codes.has(resultFilter))) resultFilter = 'all';
      view = resultFilter;
      const chips = [['all', `全部（${analyses.length} 份）`], ['unresolved', `未修复（${analyses.filter(a => a.needsAttention).length} 份）`]]
        .concat(categories.filter(([code]) => analyses.some(a => a.codes.has(code))).map(([code, label]) => {
          const fileLevel = code === 'link_file_error' || code === 'no_difference';
          const count = fileLevel ? analyses.filter(a => a.codes.has(code)).length : categoryCount(analyses, code);
          return [code, `${label}（${count} ${fileLevel ? '份' : '处'}）`];
        }));
      $('repairFilters').innerHTML = `<div class="repair-filter-buttons" role="group" aria-label="${hidden ? '隐藏恢复' : '引用错误修复'}结果分类">`
        + chips.map(([key, label]) => `<button type="button" class="btn ${key === view ? 'btn-primary' : 'btn-outline'} btn-sm" data-repair-filter="${key}" aria-pressed="${key === view}" aria-controls="repairResults">${label}</button>`).join('')
        + '</div><p class="toolbar-hint">同一文件可能出现在多个分类中；明细只显示当前分类的项目。</p>';
      entries = viewEntries(analyses, view);
      emptyMessage = view === 'unresolved' ? (hidden ? '没有未恢复的项目或处理失败的文件。' : '没有未修复的引用错误或处理失败的文件。') : '该分类没有结果。';
    } else entries = data.files;
    $('repairResults').innerHTML = entries.map(file => table(file, data, phase, view)).join('') || `<p>${emptyMessage}</p>`;
    if (colors) {
      $('clearResultPanel').hidden = !data.clear_files;
      if (data.clear_files) renderClear({folder:data.folder, files:data.clear_files}, true);
    }
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
      if (hidden) {
        payload.options = Object.fromEntries(Object.entries(optionIds).map(([key, id]) => [key, $(id).checked]));
        if (!Object.values(payload.options).some(Boolean)) throw Error('请至少选择一种要恢复的隐藏状态');
      }
      const response = await fetch(api + '/run', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)});
      const result = await response.json();
      if (!response.ok || !result.ok) throw Error(result.error || '操作失败');
      displayLimits.clear();
      clearFilter = 'all';
      resultFilter = colors ? 'unresolved' : 'all';
      try { localStorage.setItem(FOLDER, folder); localStorage.setItem(OLD, old_folder); } catch (_) {}
      await refresh();
    } catch (error) {
      $('repairError').textContent = error.message;
      await refresh();
    } finally { submitting = false; controls(); }
  }
  $('repairRun').addEventListener('click', run);
  if (colors) $('clearFilters').addEventListener('click', event => {
    const button = event.target.closest('[data-repair-filter]');
    if (!button) return;
    clearFilter = button.dataset.repairFilter;
    render();
    $('clearFilters').querySelector('[aria-pressed="true"]').focus();
  });
  $('repairFilters').addEventListener('click', event => {
    const button = event.target.closest('[data-repair-filter]');
    if (!button) return;
    resultFilter = button.dataset.repairFilter;
    displayLimits.clear();
    render();
    $('repairFilters').querySelector('[aria-pressed="true"]').focus();
  });
  $('repairResults').addEventListener('click', event => {
    const button = event.target.closest('[data-repair-more]');
    if (!button) return;
    const name = button.dataset.repairMore;
    displayLimits.set(name, (displayLimits.get(name) || 200) + 200);
    render();
  });
  controls();
  refresh();
})();
