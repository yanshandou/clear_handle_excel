let currentFilter = 'all';
let latestReportData = null;
let lastRenderKey = '';
const openingFiles = new Set();  // 正在打开的文件：表格每次刷新都会重建，不能靠元素上的 class 防连点

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
      let sheetsHtml = `<span style="color:#64748b;">共 ${allSheets.length} 个 Sheet</span>`;
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

      const fileDisplayName = f.renamed_to && f.renamed_to !== f.name
        ? `${escapeHtml(f.name)} <span style="color:#b45309; font-size:12px;">→ ${escapeHtml(f.renamed_to)}</span>`
        : escapeHtml(f.name);
      const fileLink = `<a href="#" class="file-open-link" data-file="${escapeHtml(f.key)}" title="点击在 Excel 中打开">📄 ${fileDisplayName}</a>`;

      const det = Number(f.clip_detected || 0);
      const fix = Number(f.clip_fixed || 0);
      const accepted = Number(f.clip_accepted || 0);
      // 还没轮到 / 正在处理的文件不能显示「无切字」，看起来像已经查过了
      let clipHtml = ['pending', 'processing'].includes(f.status) && det === 0
        ? '<span style="color:#94a3b8;">-</span>'
        : '<span style="color:#059669; font-weight:600;">无切字</span>';
      if (det > 0) {
        // 「已处理 / 检出」：修好的和确认无需修改的都算处理完
        const clipColor = fix + accepted >= det ? '#059669' : '#d97706';
        const detail = [fix ? `已修 ${fix}` : '', accepted ? `无需修改 ${accepted}` : ''].filter(Boolean).join(' · ');
        clipHtml = `<b style="color:${clipColor};">${fix + accepted} / ${det}</b>` +
          (detail ? `<div>${detail}</div>` : '');
      }

      const a1Html = f.normalized_a1
        ? '<span style="color:#059669; font-weight:700;">✓ 已标准化</span>'
        : '<span style="color:#94a3b8;">-</span>';

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
          <td style="color:${f.error ? '#dc2626' : '#334155'};">${escapeHtml(noteText)}</td>
          <td style="text-align:center; color:#64748b; font-family:monospace;">${escapeHtml(f.updated_at || '-')}</td>
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

async function openRecordFile(link) {
  const file = link.dataset.file;
  if (openingFiles.has(file)) return;
  openingFiles.add(file);
  try {
    const res = await fetch('/api/records/open', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file }),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || '打开失败');
  } catch (e) {
    alert(`无法打开文件：${e.message}`);
  } finally {
    setTimeout(() => openingFiles.delete(file), 1500);
  }
}

document.getElementById('recordsTableBody').addEventListener('click', e => {
  const link = e.target.closest('.file-open-link');
  if (!link) return;
  e.preventDefault();
  openRecordFile(link);
});

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
    alert('复制失败，请直接点击「下载 Markdown 报告 (.md)」按钮。');
  }
}

setInterval(pollRecords, 1200);
pollRecords();
