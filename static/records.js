let currentFilter = 'all';
let latestReportData = null;

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
    renderFilesTable(Object.values(latestReportData.files || {}));
  }
}

function getFileStatusBadge(status) {
  switch (status) {
    case 'processing':
      return '<span class="badge badge-selected">🔄 正在处理</span>';
    case 'pending':
      return '<span class="badge badge-ignored">⏳ 待处理</span>';
    case 'needs_fix':
      return '<span class="badge badge-pending">🔧 待微调切字</span>';
    case 'done':
      return '<span class="badge badge-fixed">✅ 已完成</span>';
    case 'error':
      return '<span class="badge badge-warn">❌ 处理失败</span>';
    default:
      return `<span class="badge badge-ignored">${escapeHtml(status)}</span>`;
  }
}

function renderFilesTable(filesList) {
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
      let sheetsHtml = `<span style="color:#64748b;">共 ${allSheets.length} 个 Sheet</span>`;
      if (delSet.size > 0) {
        const delTags = Array.from(delSet)
          .map(s => `<span class="deleted-sheet-tag">~~${escapeHtml(s)}~~</span>`)
          .join(' ');
        sheetsHtml += `<div style="margin-top:4px;">已删备份: ${delTags}</div>`;
      }

      const det = Number(f.clip_detected || 0);
      const fix = Number(f.clip_fixed || 0);
      let clipHtml = '<span style="color:#059669; font-weight:600;">无切字</span>';
      if (det > 0) {
        const clipColor = fix >= det ? '#059669' : '#d97706';
        clipHtml = `<b style="color:${clipColor};">${fix} / ${det}</b>`;
      }

      const a1Html = f.normalized_a1
        ? '<span style="color:#059669; font-weight:700;">✓ 已标准化</span>'
        : '<span style="color:#94a3b8;">-</span>';

      const rowClass = f.status === 'processing' ? 'row-processing' : '';
      const noteText = f.error ? `❌ ${f.error}` : (f.current_action || '-');

      return `
        <tr class="${rowClass}">
          <td><b>${Number(f.index || 1)}</b></td>
          <td><b>📄 ${escapeHtml(f.name)}</b></td>
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
    const data = await res.json();
    latestReportData = data;

    const isRunning = data.status === 'running';
    document.getElementById('statusText').innerText = isRunning ? '⏳ 任务执行中...' : '✅ 就绪';

    const taskBadge = document.getElementById('taskBadge');
    if (data.status === 'running') {
      taskBadge.className = 'badge badge-selected';
      taskBadge.textContent = '🔄 正在执行';
    } else if (data.status === 'completed') {
      taskBadge.className = 'badge badge-fixed';
      taskBadge.textContent = '✅ 已完成';
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

    const filesList = Object.values(data.files || {});
    const totalClips = filesList.reduce((acc, f) => acc + Number(f.clip_detected || 0), 0);
    const fixedClips = filesList.reduce((acc, f) => acc + Number(f.clip_fixed || 0), 0);

    document.getElementById('kpiTotal').innerText = total;
    document.getElementById('kpiDone').innerText = done;
    document.getElementById('kpiProcessing').innerText = Number(data.processing_files || 0);
    document.getElementById('kpiPending').innerText = Number(data.pending_files || 0);
    document.getElementById('kpiClips').innerText = `${fixedClips} / ${totalClips}`;

    renderFilesTable(filesList);

    const mdPreview = document.getElementById('markdownPreview');
    if (data.markdown && mdPreview.textContent !== data.markdown) {
      mdPreview.textContent = data.markdown;
    }
  } catch (e) {}
}

async function copyMarkdownReport() {
  const md = latestReportData?.markdown || document.getElementById('markdownPreview').textContent || '';
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
