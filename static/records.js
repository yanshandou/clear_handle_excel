let currentFilter = 'all';
let latestReportData = null;
let lastRenderKey = '';

// 看板按记录的键（相对目标文件夹的路径）排序显示；打开文件也用这个键
function filesOf(data) {
  return Object.entries(data.files || {})
    .map(([key, f]) => ({ ...f, key }))
    .sort((a, b) => Number(a.index || 0) - Number(b.index || 0) || a.key.localeCompare(b.key));
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function setStatusFilter(filter) {
  currentFilter = filter;
  document.querySelectorAll('#filterBar .filter-pill').forEach(btn => {
    btn.classList.toggle('active', btn.getAttribute('data-filter') === filter);
  });
  if (latestReportData) {
    lastRenderKey = '';
    renderFilesTable(filesOf(latestReportData));
  }
}

function getFileStatusBadge(status) {
  switch (status) {
    case 'processing':
      return '<span class="badge badge-selected">🔄 正在处理</span>';
    case 'pending':
      return '<span class="badge badge-ignored">⏳ 待处理</span>';
    case 'needs_fix':
      return '<span class="badge badge-pending">🔧 待人工修改</span>';
    case 'done':
      return '<span class="badge badge-fixed">✅ 已完成</span>';
    case 'error':
      return '<span class="badge badge-warn">❌ 处理失败</span>';
    default:
      return `<span class="badge badge-ignored">${escapeHtml(status)}</span>`;
  }
}

function renderFilesTable(filesList) {
  // 内容没变就不重建表格：不打断选中文字，也不让刚点的链接被换掉
  const renderKey = currentFilter + JSON.stringify(filesList);
  if (renderKey === lastRenderKey) return;
  lastRenderKey = renderKey;
  const tbody = document.getElementById('recordsTableBody');

  const cntAll = filesList.length;
  const cntProc = filesList.filter(f => f.status === 'processing').length;
  const cntPend = filesList.filter(f => f.status === 'pending').length;
  const cntNeeds = filesList.filter(f => f.status === 'needs_fix').length;
  const cntDone = filesList.filter(f => f.status === 'done').length;
  const cntErr = filesList.filter(f => f.status === 'error').length;

  document.getElementById('cntAll').innerText = cntAll;
  document.getElementById('cntProcessing').innerText = cntProc;
  document.getElementById('cntPending').innerText = cntPend;
  document.getElementById('cntNeedsFix').innerText = cntNeeds;
  document.getElementById('cntDone').innerText = cntDone;
  document.getElementById('cntError').innerText = cntErr;

  const filtered = filesList.filter(f => {
    if (currentFilter === 'all') return true;
    return f.status === currentFilter;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="9" class="empty-cell">当前筛选条件下没有匹配的 Excel 文件记录。</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = filtered
    .map(f => {
      const allSheets = f.all_sheets || [];
      const delSet = new Set(f.deleted_backups || []);
      const uncSet = new Set(f.uncertain_sheets || []);
      let sheetsHtml = `<span style="color:#475569;">共 ${allSheets.length} 个 Sheet</span>`;
      if (delSet.size > 0) {
        const delTags = Array.from(delSet)
          .map(s => `<span class="deleted-sheet-tag">${escapeHtml(s)}</span>`)
          .join(' ');
        sheetsHtml += `<div style="margin-top:4px;">已删备份: ${delTags}</div>`;
      }
      if (uncSet.size > 0) {
        const uncTags = Array.from(uncSet)
          .map(s => `<span class="uncertain-sheet-tag">${escapeHtml(s)}</span>`)
          .join(' ');
        sheetsHtml += `<div style="margin-top:4px;">待定存疑: ${uncTags}</div>`;
      }

      for (const change of f.sheet_changes || []) {
        for (const [oldName, newName] of Object.entries(change.renamed || {})) {
          sheetsHtml += `<div style="margin-top:4px;" title="${escapeHtml(change.at || '')}">改名: ${escapeHtml(oldName)} → ${escapeHtml(newName)}</div>`;
        }
      }

      const fileDisplayName = f.renamed_to && f.renamed_to !== f.name
        ? `${escapeHtml(f.name)} <span style="color:#b45309; font-size:12px;">→ ${escapeHtml(f.renamed_to)}</span>`
        : escapeHtml(f.name);
      const fileLink = excelLinkHtml(`📄 ${fileDisplayName}`, { file: f.key, record: true });

      const det = Number(f.clip_detected || 0);
      const fix = Number(f.clip_fixed || 0);
      const accepted = Number(f.clip_accepted || 0);
      // 还没轮到 / 正在处理的文件不能显示「无切字」，看起来像已经查过了
      let clipHtml = ['pending', 'processing'].includes(f.status) && det === 0
        ? '<span style="color:#64748b;">-</span>'
        : '<span style="color:#047857; font-weight:600;">无切字</span>';
      if (det > 0) {
        // 「已处理 / 检出」：修好的和确认无需修改的都算处理完
        const clipColor = fix + accepted >= det ? '#047857' : '#b45309';
        const detail = [fix ? `已修 ${fix}` : '', accepted ? `无需修改 ${accepted}` : ''].filter(Boolean).join(' · ');
        clipHtml = `<b style="color:${clipColor};">${fix + accepted} / ${det}</b>` +
          (detail ? `<div>${detail}</div>` : '');
      }

      const a1Html = f.normalized_a1
        ? '<span style="color:#047857; font-weight:700;">✓ 已标准化</span>'
        : '<span style="color:#64748b;">-</span>';

      const rowClass = f.status === 'processing' ? 'row-processing' : '';
      const noteText = f.error ? `❌ ${f.error}` : (f.current_action || '-');

      return `
        <tr class="${rowClass}">
          <td><b>${Number(f.index || 1)}</b></td>
          <td><b>${fileLink}</b></td>
          <td>${getFileStatusBadge(f.status)}</td>
          <td>${sheetsHtml}</td>
          <td style="text-align:center;"><b>${Number(f.exported_pdfs || 0)}</b></td>
          <td style="text-align:center;">${clipHtml}</td>
          <td style="text-align:center;">${a1Html}</td>
          <td style="color:${f.error ? '#b91c1c' : '#334155'};">${escapeHtml(noteText)}</td>
          <td style="text-align:center; color:#475569; font-family:monospace;">${escapeHtml(f.updated_at || '-')}</td>
        </tr>
      `;
    })
    .join('');
}

async function pollRecords() {
  try {
    const res = await fetch('/api/records');
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    // 看着看板时任务跑完：撒花庆祝
    if (latestReportData && latestReportData.status === 'running' && data.status === 'completed') {
      window.gameFx?.celebrate('任务完成！', data.task_name || '');
    }
    latestReportData = data;

    const isRunning = data.status === 'running';
    document.getElementById('statusText').innerText = isRunning ? '⏳ 任务执行中...'
      : (data.status === 'interrupted' ? '任务已中断，请回主页或工作台继续' : '✅ 就绪');

    const taskBadge = document.getElementById('taskBadge');
    if (data.status === 'running') {
      taskBadge.className = 'badge badge-selected';
      taskBadge.textContent = '🔄 正在执行';
    } else if (data.status === 'completed') {
      taskBadge.className = 'badge badge-fixed';
      taskBadge.textContent = '✅ 已完成';
    } else if (data.status === 'interrupted') {
      taskBadge.className = 'badge badge-warn';
      taskBadge.textContent = '已中断，等待继续';
    } else {
      taskBadge.className = 'badge badge-ignored';
      taskBadge.textContent = '⚪ 空闲';
    }

    document.getElementById('taskNameText').textContent = data.task_name || '尚未开始任务';
    document.getElementById('folderMetaText').textContent =
      `目标文件夹：${data.folder || '暂无'} ｜ 开始时间：${data.started_at || '-'} ｜ 最后更新：${data.updated_at || '-'}`;

    const pct = Number(data.progress_pct || 0);
    const total = Number(data.total_files || 0);
    const done = Number(data.completed_files || 0);

    document.getElementById('currentStepText').textContent = data.current_step || '等待开始...';
    document.getElementById('progressPctText').textContent = `${pct}% (${done} / ${total} 文件)`;
    document.getElementById('progressBarFill').style.width = `${Math.min(100, Math.max(0, pct))}%`;

    const filesList = filesOf(data);
    const totalClips = filesList.reduce((acc, f) => acc + Number(f.clip_detected || 0), 0);
    const fixedClips = filesList.reduce((acc, f) => acc + Number(f.clip_fixed || 0), 0);
    const acceptedClips = filesList.reduce((acc, f) => acc + Number(f.clip_accepted || 0), 0);

    document.getElementById('kpiTotal').innerText = total;
    document.getElementById('kpiDone').innerText = done;
    document.getElementById('kpiProcessing').innerText = Number(data.processing_files || 0);
    document.getElementById('kpiPending').innerText = Number(data.pending_files || 0);
    document.getElementById('kpiClips').innerText = `${fixedClips + acceptedClips} / ${totalClips}` +
      (acceptedClips ? `（已修 ${fixedClips}，无需修改 ${acceptedClips}）` : '');

    renderFilesTable(filesList);
  } catch (e) {
    // 工具窗口关了或出错时要说清楚，不然看起来像面板一直不更新
    document.getElementById('statusText').innerText = '⚠️ 连不上工具（命令行窗口是否关了？），下面是之前的数据';
  }
}

async function copyMarkdownReport() {
  const md = latestReportData?.markdown || '';
  if (!md) return;
  try {
    await navigator.clipboard.writeText(md);
    const toast = document.getElementById('copyToast');
    toast.textContent = '✓ 已复制 Markdown 到剪贴板！';
    setTimeout(() => {
      toast.textContent = '';
    }, 2500);
  } catch (e) {
    appAlert('复制失败，请直接点击「只下载 Markdown (.md)」按钮。', { tone: 'error' });
  }
}

setInterval(pollRecords, 1200);
pollRecords();

let operationHistory = null;
let historyKey = '';
function renderOperationHistory(data) {
  const select = document.getElementById('historyFolder');
  const selected = select.value;
  const folders = data.folders || [];
  select.innerHTML = folders.length ? folders.map(f => `<option value="${escapeHtml(f.folder)}">${escapeHtml(f.folder)}</option>`).join('') : '<option value="">暂无操作记录</option>';
  if (folders.some(f => f.folder === selected)) select.value = selected;
  const group = folders.find(f => f.folder === select.value);
  const statuses = {running: '进行中', completed: '已完成', failed: '部分失败或失败', interrupted: '已中断'};
  document.getElementById('historyBody').innerHTML = Object.entries(data.labels).map(([kind, label]) => {
    const item = group?.operations?.[kind];
    return `<tr><td>${escapeHtml(label)}</td><td>${item ? item.count + ' 次' : '0 次'}</td><td>${escapeHtml(item?.last || '—')}</td><td>${escapeHtml(item ? (statuses[item.status] || item.status) : '未操作')}</td></tr>`;
  }).join('');
}
document.getElementById('historyFolder').addEventListener('change', () => { if (operationHistory) renderOperationHistory(operationHistory); });
async function pollOperationHistory() {
  try {
    const response = await fetch('/api/operation-history');
    if (!response.ok) throw new Error('history unavailable');
    const data = await response.json();
    const key = JSON.stringify(data);
    if (key !== historyKey) { historyKey = key; operationHistory = data; renderOperationHistory(data); }
  } catch (e) { document.getElementById('historyBody').innerHTML = '<tr><td colspan="4">无法加载操作记录，请稍后重试。</td></tr>'; historyKey = ''; }
}
pollOperationHistory();
setInterval(pollOperationHistory, 3000);
