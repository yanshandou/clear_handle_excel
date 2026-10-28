'use strict';
const el = id => document.getElementById(id);
const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
const statuses = {completed:'已完成', interrupted:'已中断，可续跑', running:'进行中', pending:'未处理', cached:'沿用缓存', error:'处理失败', timeout:'扫描超时'};
const duration = value => value >= 60 ? `${Math.floor(Math.round(value) / 60)} 分 ${Math.round(value) % 60} 秒` : `${Number(value || 0).toFixed(1)} 秒`;
const issueCount = row => Object.values(row.issues || {}).reduce((sum, count) => sum + count, 0);
const params = new URLSearchParams(location.search);
let currentReport = null, page = 0, loading = false;

function renderFiles() {
  const query = el('fileSearch').value.trim().toLocaleLowerCase(), status = el('statusFilter').value;
  const rows = currentReport.files.filter(row => (!status || row.status === status)
    && `${row.name} ${row.system}`.toLocaleLowerCase().includes(query));
  const sort = el('sortFiles').value;
  rows.sort(sort === 'name' ? (a,b) => a.name.localeCompare(b.name) : sort === 'issues'
    ? (a,b) => issueCount(b) - issueCount(a) || b.seconds - a.seconds : (a,b) => b.seconds - a.seconds);
  const pages = Math.max(1, Math.ceil(rows.length / 50));
  page = Math.min(page, pages - 1);
  el('fileCount').textContent = `${rows.length} / ${currentReport.files.length} 个文件，展开文件名可查看详细数据。失败文件的检出数属于局部结果。`;
  el('files').innerHTML = rows.slice(page * 50, (page + 1) * 50).map(row => {
    const ops = Object.entries(row.operations).map(([key, op]) => {
      const label = currentReport.operations.find(item => item.key === key)?.name || key;
      return `${escapeHtml(label)} ${duration(op.seconds)}`;
    }).join(' · ');
    const sheets = Object.entries(row.sheets).map(([name, sheet]) => `${escapeHtml(name)}（${sheet.checked ? '已检测' : '未完成'}）`).join('、');
    return `<tr><td><details><summary>${escapeHtml(row.name)}</summary><div class="analysis-detail"><p>系统：${escapeHtml(row.system)} · 文件大小：${row.size == null ? '未知' : `${(row.size / 1048576).toFixed(2)} MB`}</p><p>${ops || '没有已采集的操作计时'}</p><p>检测 Sheet：${sheets || '无'}</p><p>免检：${row.exempt_sheets.map(escapeHtml).join('、') || '无'}；隐藏、备份等过滤表：${row.skipped_sheets.map(escapeHtml).join('、') || '无'}</p><p>PDF 导出尝试 ${row.pdf_exports || 0} 次，复用 ${row.pdf_reused || 0} 次，检测页面 ${row.pdf_pages || 0} 页（包含重试）。</p>${row.stuck_sheet ? `<p>卡住的 Sheet：${escapeHtml(row.stuck_sheet)}</p>` : ''}${row.error ? `<p class="analysis-error">${escapeHtml(row.error)}</p>` : ''}</div></details></td><td><span class="analysis-status ${escapeHtml(row.status)}">${statuses[row.status] || escapeHtml(row.status)}</span></td><td>${duration(row.seconds)}<br><small>${row.attempts} 次尝试</small></td><td>${issueCount(row)}${row.status !== 'completed' && issueCount(row) ? '（局部）' : ''}</td></tr>`;
  }).join('') || '<tr><td colspan="4">没有符合条件的文件。</td></tr>';
  el('pageLabel').textContent = `${page + 1} / ${pages} 页`;
  el('prevPage').disabled = page === 0;
  el('nextPage').disabled = page >= pages - 1;
}

function renderReport(data) {
  currentReport = data;
  const {run, summary:s} = data;
  el('runMeta').textContent = `${run.folder} · ${run.mode} · ${new Date(run.started_at).toLocaleString('zh-CN')} · ${statuses[run.status] || run.status}`;
  const incomplete = s.total - (s.completed || 0) - (s.cached || 0);
  el('runNotice').textContent = run.status === 'running' ? '扫描进行中：已结束文件的计时已保存，可点击“刷新数据”查看最新结果。'
    : incomplete ? `本次扫描有 ${incomplete} 个文件未完成检查，${s.issues} 处检出仅统计已完成文件，不能将其余文件视为通过。`
    : `本次扫描已结束：${s.completed || 0} 个文件完成检查，${s.cached || 0} 个沿用缓存。报告保留扫描时的数据，后续修复不会覆盖。`;
  el('runNotice').classList.toggle('warning', incomplete > 0);
  const metrics = [['累计文件处理时间',duration(s.seconds),'info'],['完成 / 文件总数',`${s.completed || 0} / ${s.total}`,''],['本次检出',`${s.issues} 处`,''],['未完成文件',`${incomplete} 个`,incomplete ? 'alert' : ''],['已检测 Sheet',`${s.checked_sheets} 张`,''],['免检 Sheet',`${s.exempt_sheets} 张`,'']];
  el('metrics').innerHTML = metrics.map(([label, value, cls]) => `<div class="analysis-metric ${cls}"><strong>${value}</strong><span>${label}</span></div>`).join('');
  el('findings').innerHTML = data.findings.map(item => `<article class="analysis-finding ${item.level}"><h3>${escapeHtml(item.title)}</h3><p>${escapeHtml(item.evidence)}</p><p><strong>建议：</strong>${escapeHtml(item.action)}</p></article>`).join('');
  el('issueTypes').innerHTML = Object.entries(data.issues).map(([name,count]) => `<span class="analysis-chip">${escapeHtml(name)} · ${count}</span>`).join('') || '<span class="analysis-muted">已完成文件中没有检出问题；请同时确认上方覆盖情况。</span>';
  el('operations').innerHTML = data.operations.map(op => `<div class="analysis-bar-row"><div class="analysis-bar-label"><span>${escapeHtml(op.name)}</span><span>${duration(op.seconds)} · ${(op.share * 100).toFixed(1)}%</span></div><div class="analysis-bar-track" role="img" aria-label="${escapeHtml(op.name)}占比 ${(op.share * 100).toFixed(1)}%"><span class="analysis-bar-fill" style="width:${Math.min(100,Math.max(0,op.share*100))}%"></span></div></div>`).join('');
  el('systems').innerHTML = data.systems.map(row => `<tr><td>${escapeHtml(row.name)}</td><td>${row.files}</td><td>${duration(row.seconds)}</td><td>${row.incomplete}</td></tr>`).join('');
  el('sheets').innerHTML = data.sheets.slice(0,20).map(row => `<tr><td>${escapeHtml(row.file)}</td><td>${escapeHtml(row.sheet)}</td><td>${duration(row.seconds)}</td><td>${row.complete ? `${row.issues} 处检出` : '局部计时，未完成检查'}</td></tr>`).join('') || '<tr><td colspan="4">本次没有 Sheet 级检测计时（可能沿用缓存、全部免检或尚未进入检测）。</td></tr>';
  el('downloadBtn').href = `/api/scan-analysis/markdown?id=${encodeURIComponent(run.id)}`;
  el('downloadBtn').classList.remove('hidden');
  renderFiles();
}

async function loadReport(id = params.get('id')) {
  if (loading) return;
  loading = true;
  el('refreshBtn').disabled = true;
  el('analysisError').classList.add('hidden');
  try {
    const query = new URLSearchParams();
    if (params.get('folder')) query.set('folder', params.get('folder'));
    if (id) query.set('id', id);
    const response = await fetch(`/api/scan-analysis?${query}`);
    const body = await response.json();
    if (!response.ok || !body.ok) throw new Error(body.error || '无法读取扫描报告');
    el('reportSelect').innerHTML = body.reports.map(run => `<option value="${escapeHtml(run.id)}">${escapeHtml(new Date(run.started_at).toLocaleString('zh-CN'))} · ${escapeHtml(run.mode)} · ${run.total} 个文件 · ${escapeHtml(statuses[run.status] || run.status)}</option>`).join('') || '<option value="">暂无扫描记录</option>';
    const data = body.report;
    el('emptyState').classList.toggle('hidden', !!data);
    el('reportBody').classList.toggle('hidden', !data);
    el('downloadBtn').classList.toggle('hidden', !data);
    if (data) {
      if (!body.reports.some(run => run.id === data.run.id)) el('reportSelect').add(new Option('当前扫描报告', data.run.id));
      el('reportSelect').value = data.run.id;
      params.set('id', data.run.id);
      history.replaceState(null, '', `?${params}`);
      page = 0;
      renderReport(data);
    }
  } catch (error) {
    el('analysisError').textContent = error.message;
    el('analysisError').classList.remove('hidden');
  } finally {
    loading = false;
    el('refreshBtn').disabled = false;
  }
}
el('refreshBtn').addEventListener('click', () => loadReport(el('reportSelect').value || null));
el('reportSelect').addEventListener('change', event => loadReport(event.target.value));
for (const id of ['fileSearch','statusFilter','sortFiles']) el(id).addEventListener(id === 'fileSearch' ? 'input' : 'change', () => {page = 0; if(currentReport) renderFiles();});
el('prevPage').addEventListener('click', () => {page--; renderFiles();});
el('nextPage').addEventListener('click', () => {page++; renderFiles();});
loadReport();
