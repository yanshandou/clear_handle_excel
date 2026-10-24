const TASKS = {
  export_pdf: {
    title: '导出 PDF',
    icon: '📄',
    theme: 'feature-sky',
    sub: '把每个工作簿的可见工作表分别导出为 PDF',
    notes: [
      '自动递归扫描所选文件夹及其所有子文件夹，以只读方式打开 Excel，不修改原文件',
      '跳过隐藏表、変更履歴和备份表',
      '导出的 PDF 直接保存在对应 Excel 文件所在的同一目录下',
    ],
    runLabel: '📄 开始导出',
  },
  normalize_view: {
    title: '视图标准化',
    icon: '🎯',
    theme: 'feature-mint',
    sub: '所有可见工作表定位到 A1、缩放设为 100% 后保存',
    notes: [
      '自动递归处理所选文件夹及其所有子文件夹下的 Excel 文件',
      '保存时停留在第一个可见工作表',
    ],
    runLabel: '🎯 开始标准化',
  },
};

// 切字工作台发起的任务类型
const CLIP_JOB_KINDS = new Set(['scan', 'fix', 'refresh', 'recheck', 'manual_open', 'batch_scan', 'batch_commit', 'batch_recheck', 'batch_open', 'batch_unhighlight']);
const FOLDER_STORAGE_KEY = 'clear_handle_excel.lastFolder';
const STATUS_BADGES = {
  running: ['badge-selected', '🔄 正在执行'],
  completed: ['badge-fixed', '✅ 已完成'],
  interrupted: ['badge-warn', '⏸ 已中断'],
};

let overview = null;
let activeKind = null;
let phase = 'form'; // form | running | done
let starting = false; // 已提交、后台尚未确认开始时保持按钮为执行中
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
  // 任务执行中进度还在变化，不允许导出或导入
  document.getElementById('exportProgressBtn').disabled = overview.busy;
  document.getElementById('importProgressBtn').disabled = overview.busy;
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
  [...Object.keys(TASKS), 'clean_sheets'].forEach(kind => {
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

  document.getElementById('taskFolderInput').disabled = starting || phase === 'running';
  document.getElementById('taskRunBox').classList.toggle('hidden', phase !== 'running' && phase !== 'done');
  hint.textContent = '';

  if (starting || phase === 'running') {
    btn.disabled = true;
    btn.textContent = '⏳ 正在执行...';
    hint.textContent = '可以关闭此窗口，任务会在后台继续';
  } else if (phase === 'done') {
    btn.disabled = false;
    btn.textContent = canResumeActiveTask() ? '🔁 继续上次任务' : '再执行一次';
  } else {
    btn.disabled = busyOther;
    btn.textContent = task.runLabel;
  }
  if (busyOther && phase === 'form' && !starting) {
    hint.textContent = `后台正在执行「${overview.progress.task_name}」，完成后才能开始`;
  }
  if (phase === 'running' || phase === 'done') renderRunBox();
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
  const outDir = job.pdf_folder || job.folder;
  const html = `
    <b>${failed ? '⚠️ 部分文件未处理成功' : '✅ 已完成'}</b>
    <span>${escapeHtml(failed ? '失败原因见下方日志，可点击「继续上次任务」只重试未完成的文件。' : p.current_step)}</span>
    ${isExport && outDir ? `<span>PDF 输出位置：与各 Excel 文件在同一目录下（<code>${escapeHtml(outDir)}</code>）</span>` : ''}
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
  if (!activeKind || starting || document.getElementById('taskPrimaryBtn').disabled) return;
  if (phase === 'done' && canResumeActiveTask()) {
    resumeTask();
  } else if (phase === 'done') {
    phase = 'form';
    renderDialog();
  } else {
    startTask();
  }
}

function readFolderInput() {
  const folder = document.getElementById('taskFolderInput').value.trim();
  if (!folder) appAlert('请先输入 Excel 文件夹路径！', { tone: 'warn' });
  return folder;
}

async function startTask() {
  const folder = readFolderInput();
  if (!folder) return;
  if (overview && overview.can_resume && !(await appConfirm('开始新任务会替换上次中断任务的续跑进度，确定吗？', { okText: '开始新任务' }))) return;
  // 后端先递归列出文件再返回，期间轮询不能把按钮还原成可点击
  starting = true;
  renderDialog();
  try {
    await postAction('/api/quick_task', { kind: activeKind, folder });
    saveLastFolder(folder);
    phase = 'running';
    minAcceptedSeq = pollSeq + 1;
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
  }
  starting = false;
  renderDialog();
  await pollState();
}

async function resumeTask() {
  try {
    await postAction('/api/resume', {});
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
    return;
  }
  if (canResumeActiveTask()) phase = 'running';
  minAcceptedSeq = pollSeq + 1;
  await pollState();
}

/* ==========================================
   进度备份：导出 / 导入
   ========================================== */
async function exportProgress() {
  const btn = document.getElementById('exportProgressBtn');
  btn.disabled = true;
  try {
    const res = await fetch('/api/progress/export');
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || '导出进度失败');
    }
    const blob = await res.blob();
    const match = /filename="([^"]+)"/.exec(res.headers.get('Content-Disposition') || '');
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = match ? match[1] : 'excel_progress.json';
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
  } finally {
    btn.disabled = Boolean(overview && overview.busy);
  }
}

function pickProgressFile() {
  const input = document.getElementById('progressFileInput');
  input.value = '';
  input.click();
}

async function importProgressFile(file) {
  const message = `导入「${file.name}」会替换当前全部进度（扫描结果、人工判断、续跑点和执行记录），无法撤销。\n`
    + '建议先点「导出进度」备份当前进度。确定导入吗？';
  if (!(await appConfirm(message, { title: '导入进度', okText: '导入', tone: 'danger' }))) return;
  try {
    const res = await fetch('/api/progress/import', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: await file.text(),
    });
    const data = await res.json();
    if (!res.ok || !data.ok) throw new Error(data.error || '操作未成功');
    const lines = [
      `已导入进度（导出于 ${data.exported_at || '未知时间'}）`,
      `切字问题 ${data.items} 处 · 执行记录 ${data.files} 个文件`,
    ];
    if (data.can_resume) lines.push('包含中断的任务，可点击「继续上次任务」继续。');
    data.warnings.forEach(warning => lines.push(`⚠️ ${warning}`));
    appAlert(lines.join('\n'), { title: '导入完成', icon: '✅' });
  } catch (e) {
    appAlert(`导入失败：${e.message}`, { tone: 'error' });
  }
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
    const wasBusy = Boolean(overview && overview.busy);
    overview = data;
    if (wasBusy && !data.busy) celebrateFinished();

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

// 后台任务刚跑完且全部成功：撒花庆祝（失败或中断不庆祝，结果框里会提示）
function celebrateFinished() {
  const p = overview.progress || {};
  if (overview.can_resume || p.status === 'interrupted' || p.error_files) return;
  const files = p.total_files ? ` · ${p.total_files} 个文件` : '';
  window.gameFx?.celebrate('任务完成！', `${p.task_name || '后台任务'}${files}`);
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
taskFolderInput.addEventListener('keydown', e => {
  if (e.key === 'Enter') onPrimaryAction();
});

document.getElementById('progressFileInput').addEventListener('change', e => {
  const file = e.target.files[0];
  if (file) importProgressFile(file);
});

setInterval(pollState, 1200);
pollState();
