const TASKS = {
  clean_sheets: {
    title: '精简 Sheet',
    icon: '🧹',
    theme: 'feature-violet',
    sub: '删除名称中带 YYYYMMDD 日期的备份工作表',
    notes: [
      '先列出将被删除的备份表，确认后才会修改 Excel',
      '删除前自动备份原文件到 _backup_时间戳 文件夹',
      '全部工作表都是日期名的工作簿会跳过，不会被删空',
    ],
    preview: true,
  },
  export_pdf: {
    title: '导出 PDF',
    icon: '📄',
    theme: 'feature-sky',
    sub: '把每个工作簿的可见工作表分别导出为 PDF',
    notes: [
      '以只读方式打开 Excel，不修改原文件',
      '跳过隐藏表、変更履歴和日期备份表',
      'PDF 输出到所选文件夹下的 _exported_pdfs_时间戳 文件夹',
    ],
    runLabel: '📄 开始导出',
  },
  normalize_view: {
    title: '视图标准化',
    icon: '🎯',
    theme: 'feature-mint',
    sub: '所有可见工作表定位到 A1、缩放设为 100% 后保存',
    notes: [
      '保存时停留在第一个可见工作表',
      '修改前自动备份原文件到 _backup_时间戳 文件夹',
    ],
    runLabel: '🎯 开始标准化',
  },
};

// 切字工作台发起的任务类型
const CLIP_JOB_KINDS = new Set(['scan', 'fix', 'refresh', 'recheck', 'manual_open']);
const FOLDER_STORAGE_KEY = 'clear_handle_excel.lastFolder';
const STATUS_BADGES = {
  running: ['badge-selected', '🔄 正在执行'],
  completed: ['badge-fixed', '✅ 已完成'],
  interrupted: ['badge-warn', '⏸ 已中断'],
};

let overview = null;
let activeKind = null;
let phase = 'form'; // form | preview | running | done
let previewData = null;
let pollSeq = 0;
let minAcceptedSeq = 0; // 发起操作前已在途的轮询结果可能过期，直接丢弃

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function readLastFolder() {
  try {
    return localStorage.getItem(FOLDER_STORAGE_KEY) || '';
  } catch (e) {
    return '';
  }
}

function saveLastFolder(folder) {
  try {
    localStorage.setItem(FOLDER_STORAGE_KEY, folder);
  } catch (e) {}
}

function isTaskRunning(kind) {
  return Boolean(overview && overview.busy && overview.job && overview.job.kind === kind);
}

function canResumeActiveTask() {
  return Boolean(overview && overview.can_resume && overview.job && overview.job.kind === activeKind);
}

function countBackups(data) {
  return data ? data.files.reduce((acc, f) => acc + f.backups.length, 0) : 0;
}

/* ==========================================
   主页：顶部状态、最近任务条与功能卡片
   ========================================== */
function renderHeader() {
  const p = overview.progress;
  const pct = Number(p.progress_pct || 0);
  document.getElementById('navProgressBadge').textContent =
    p.total_files > 0 ? `${p.completed_files}/${p.total_files} (${pct}%)` : '0%';
  document.getElementById('statusText').innerText = overview.busy
    ? '⏳ 后台处理中...' : (overview.can_resume ? '任务已中断，可继续上次任务' : '✅ 就绪');
}

function renderStrip() {
  const p = overview.progress;
  const pct = Number(p.progress_pct || 0);
  const status = overview.busy ? 'running' : p.status;
  const [badgeClass, badgeText] = STATUS_BADGES[status] || ['badge-ignored', '⚪ 空闲'];

  document.getElementById('taskStrip').classList.toggle('hidden', status === 'idle' && !overview.can_resume);
  const badge = document.getElementById('stripBadge');
  badge.className = `badge ${badgeClass}`;
  badge.textContent = badgeText;
  document.getElementById('stripTitle').textContent = p.task_name || '尚未开始任务';
  document.getElementById('stripStep').textContent = p.current_step || '';
  document.getElementById('stripPct').textContent = `${pct}% · ${p.completed_files}/${p.total_files} 个文件`;
  document.getElementById('stripFill').style.width = `${Math.min(100, Math.max(0, pct))}%`;

  const resumeBtn = document.getElementById('stripResumeBtn');
  resumeBtn.classList.toggle('hidden', !overview.can_resume);
  resumeBtn.disabled = overview.busy;
}

function renderCards() {
  const pct = Number(overview.progress.progress_pct || 0);
  Object.keys(TASKS).forEach(kind => {
    const meta = document.getElementById('meta_' + kind);
    if (!meta.dataset.idleText) meta.dataset.idleText = meta.textContent;
    meta.textContent = isTaskRunning(kind) ? `⏳ 进行中 ${pct}%` : meta.dataset.idleText;
  });

  const clipsMeta = document.getElementById('meta_clips');
  const clips = overview.clips;
  if (overview.busy && overview.job && CLIP_JOB_KINDS.has(overview.job.kind)) {
    clipsMeta.textContent = `⏳ 进行中 ${pct}%`;
  } else if (!clips.total) {
    clipsMeta.textContent = '暂无切字记录';
  } else if (clips.unresolved) {
    clipsMeta.textContent = `待处理 ${clips.unresolved} 处 / 共 ${clips.total} 处`;
  } else {
    clipsMeta.textContent = `${clips.total} 处已全部处理`;
  }
}

/* ==========================================
   单项功能对话框
   ========================================== */
function openTask(kind) {
  const task = TASKS[kind];
  activeKind = kind;
  previewData = null;
  phase = isTaskRunning(kind) ? 'running' : 'form';

  document.getElementById('taskDialogHead').className = `task-dialog-head ${task.theme}`;
  document.getElementById('taskDialogIcon').textContent = task.icon;
  document.getElementById('taskDialogTitle').textContent = task.title;
  document.getElementById('taskDialogSub').textContent = task.sub;
  document.getElementById('taskNotes').innerHTML =
    task.notes.map(note => `<li>${escapeHtml(note)}</li>`).join('');

  const input = document.getElementById('taskFolderInput');
  input.value = (phase === 'running' && overview.job.folder) || readLastFolder() || (overview && overview.folder) || '';

  renderDialog();
  document.getElementById('taskDialog').showModal();
  if (phase === 'form') input.focus();
}

function closeTask() {
  document.getElementById('taskDialog').close();
}

function renderDialog() {
  if (!activeKind) return;
  const task = TASKS[activeKind];
  const busyOther = Boolean(overview && overview.busy && !isTaskRunning(activeKind));
  const btn = document.getElementById('taskPrimaryBtn');
  const hint = document.getElementById('taskHint');

  document.getElementById('taskFolderInput').disabled = phase === 'running';
  document.getElementById('previewBox').classList.toggle('hidden', phase !== 'preview');
  document.getElementById('taskRunBox').classList.toggle('hidden', phase !== 'running' && phase !== 'done');
  hint.textContent = '';

  if (phase === 'running') {
    btn.disabled = true;
    btn.textContent = '⏳ 正在执行...';
    hint.textContent = '可以关闭此窗口，任务会在后台继续';
  } else if (phase === 'done') {
    btn.disabled = false;
    btn.textContent = canResumeActiveTask() ? '🔁 继续上次任务' : '再执行一次';
  } else if (phase === 'preview') {
    const total = countBackups(previewData);
    btn.disabled = total === 0 || busyOther;
    btn.textContent = total ? `🧹 删除 ${total} 个备份表` : '没有可删除的备份表';
  } else {
    // 预览只读取文件，后台忙时也允许
    btn.disabled = busyOther && !task.preview;
    btn.textContent = task.preview ? '🔎 扫描备份表' : task.runLabel;
  }
  if (busyOther && (phase === 'form' || phase === 'preview')) {
    hint.textContent = `后台正在执行「${overview.progress.task_name}」，完成后才能开始`;
  }
  if (phase === 'running' || phase === 'done') renderRunBox();
}

function renderPreview() {
  const files = previewData.files;
  const withBackups = files.filter(f => f.backups.length);
  const withNotes = files.filter(f => f.note);
  const plainCount = files.length - withBackups.length - withNotes.length;

  document.getElementById('previewBox').innerHTML = `
    <div class="preview-summary">
      共 <b>${files.length}</b> 个 Excel 文件，其中 <b>${withBackups.length}</b> 个包含
      <b>${countBackups(previewData)}</b> 个日期备份表
    </div>
    <ul class="preview-list">
      ${withBackups.map(f => `
        <li>
          <b>📄 ${escapeHtml(f.name)}</b>
          <span class="preview-tags">
            ${f.backups.map(s => `<span class="deleted-sheet-tag preview-sheet-tag">${escapeHtml(s)}</span>`).join('')}
          </span>
        </li>`).join('')}
      ${withNotes.map(f => `
        <li class="is-warn">
          <b>📄 ${escapeHtml(f.name)}</b>
          <span>⚠️ ${escapeHtml(f.note)}</span>
        </li>`).join('')}
    </ul>
    ${plainCount ? `<div class="preview-foot">其余 ${plainCount} 个文件没有备份表，不会被修改。</div>` : ''}
  `;
}

function renderRunBox() {
  const p = overview.progress;
  const pct = Number(p.progress_pct || 0);
  document.getElementById('taskStepText').textContent = p.current_step || '准备中...';
  // 进度百分比把失败文件也计为已结束，这里分开显示完成与失败数
  document.getElementById('taskStatsText').textContent =
    `完成 ${p.completed_files}/${p.total_files}` +
    (p.error_files ? ` · 失败 ${p.error_files}` : '') + ` · ${pct}%`;
  document.getElementById('taskProgressFill').style.width = `${Math.min(100, Math.max(0, pct))}%`;

  const logBox = document.getElementById('taskLogBox');
  const logText = overview.logs.length ? overview.logs.join('\n') : '等待操作...';
  if (logBox.innerText !== logText) {
    logBox.innerText = logText;
    logBox.scrollTop = logBox.scrollHeight;
  }

  const result = document.getElementById('taskResult');
  if (phase !== 'done') {
    result.classList.add('hidden');
    return;
  }
  const job = overview.job || {};
  const failed = canResumeActiveTask() || p.status === 'interrupted';
  const isExport = activeKind === 'export_pdf';
  const outDir = isExport ? job.pdf_folder : job.backup_folder;
  const html = `
    <b>${failed ? '⚠️ 部分文件未处理成功' : '✅ 已完成'}</b>
    <span>${escapeHtml(failed ? '失败原因见下方日志，可点击「继续上次任务」只重试未完成的文件。' : p.current_step)}</span>
    ${outDir ? `<span>${isExport ? 'PDF 输出目录' : '原文件备份目录（有修改时才会创建）'}：<code>${escapeHtml(outDir)}</code></span>` : ''}
    <a href="/records">查看执行记录 ↗</a>
  `;
  // 轮询时内容不变就不重建，避免打断用户选中复制路径
  if (result.dataset.html !== html) {
    result.dataset.html = html;
    result.innerHTML = html;
  }
  result.className = `task-result ${failed ? 'is-warn' : 'is-ok'}`;
}

function onPrimaryAction() {
  if (!activeKind || document.getElementById('taskPrimaryBtn').disabled) return;
  if (phase === 'done' && canResumeActiveTask()) {
    resumeTask();
  } else if (phase === 'done') {
    phase = 'form';
    previewData = null;
    renderDialog();
  } else if (phase === 'form' && TASKS[activeKind].preview) {
    runPreview();
  } else {
    startTask();
  }
}

function readFolderInput() {
  const folder = document.getElementById('taskFolderInput').value.trim();
  if (!folder) alert('请先输入 Excel 文件夹路径！');
  return folder;
}

async function runPreview() {
  const folder = readFolderInput();
  if (!folder) return;
  const btn = document.getElementById('taskPrimaryBtn');
  btn.disabled = true;
  btn.textContent = '🔎 正在读取工作表...';
  try {
    previewData = await postAction('/api/clean/preview', { folder });
    saveLastFolder(folder);
    phase = 'preview';
    renderPreview();
  } catch (e) {
    alert(e.message);
  }
  renderDialog();
}

async function startTask() {
  const folder = readFolderInput();
  if (!folder) return;
  if (overview && overview.can_resume && !confirm('开始新任务会替换上次中断任务的续跑进度，确定吗？')) return;
  try {
    await postAction('/api/quick_task', { kind: activeKind, folder });
  } catch (e) {
    alert(e.message);
    return;
  }
  saveLastFolder(folder);
  phase = 'running';
  minAcceptedSeq = pollSeq + 1;
  await pollState();
}

async function resumeTask() {
  try {
    await postAction('/api/resume', {});
  } catch (e) {
    alert(e.message);
    return;
  }
  if (canResumeActiveTask()) phase = 'running';
  minAcceptedSeq = pollSeq + 1;
  await pollState();
}

/* ==========================================
   后端状态轮询与接口调用
   ========================================== */
async function pollState() {
  const seq = ++pollSeq;
  try {
    const res = await fetch('/api/overview');
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '读取状态失败');
    if (seq < minAcceptedSeq) return;
    overview = data;

    renderHeader();
    renderStrip();
    renderCards();
    if (activeKind && phase === 'running' && !isTaskRunning(activeKind)) {
      phase = overview.job && overview.job.kind === activeKind ? 'done' : 'form';
    }
    renderDialog();
  } catch (e) {
    document.getElementById('statusText').innerText = `连接或读取失败: ${e.message}`;
  }
}

async function postAction(path, payload) {
  const res = await fetch(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const data = await res.json();
  if (!res.ok || !data.ok) throw new Error(data.error || '操作未成功，任务可能正在执行');
  return data;
}

const taskDialog = document.getElementById('taskDialog');
taskDialog.addEventListener('close', () => {
  activeKind = null;
});
// 点击遮罩关闭；对话框内部由子元素铺满，不会误触
taskDialog.addEventListener('click', e => {
  if (e.target === taskDialog) closeTask();
});

const taskFolderInput = document.getElementById('taskFolderInput');
taskFolderInput.addEventListener('input', () => {
  if (phase === 'preview') {
    phase = 'form';
    previewData = null;
    renderDialog();
  }
});
taskFolderInput.addEventListener('keydown', e => {
  if (e.key === 'Enter') onPrimaryAction();
});

setInterval(pollState, 1200);
pollState();
