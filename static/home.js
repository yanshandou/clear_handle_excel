const TASKS = {
  export_pdf: {
    title: '调整页面',
    icon: 'adjust',
    theme: 'feature-sky',
    sub: '先把每张表设成宽一页、高自动并导出 PDF，再在网页上点按钮告诉程序怎么改（横竖、向下取整），程序改完重新生成 PDF',
    // 说明分组显示在对话框里的「📖 程序会怎么做」（默认折叠）；要和 core/workflow.py _export_workbook_pdfs 的实际做法一致
    tip: '每张表设成宽一页、高自动并导出 PDF，导完进调整页面点按钮让程序改',
    notes: [
      { title: '📂 导哪些', items: [
        '所选的 Excel，和所选文件夹里（含子文件夹）的 Excel',
        '每份里可见的表；隐藏表、変更履歴、备份表、灰色标签的表跳过',
        'Excel 没改过（哈希和上次导出时一样）、PDF 也都还在的，直接沿用上次的 PDF，不打开 Excel',
      ] },
      { title: '📐 导出时改 Excel 的地方', items: [
        '没设成宽一页的表改成「宽一页、高自动」并保存进 Excel：横向不会拆成几页，纵向几页由 Excel 定；已经是宽一页的不动，向下取整定死的页高也保留',
        '能保存的文件顺便把视图存成分页预览（缩放、选区不变）',
        '导出时发现内容很少的页：手动分页造成的空页，清掉手动分页再核对；最后几页全是空格子的也会修；修好了保存，没把握的在调整页面上标黄',
        '正在 Excel 里打开的文件照样按宽一页导出 PDF，只是这些改动存不进 Excel；关掉文件后在调整页面点「🔄 重新生成」',
      ] },
      { title: '📁 PDF 放在哪', items: [
        '统一放在所选文件夹下的「PDFS」文件夹（保留子文件夹），一张表一份',
        '开始前删掉 PDFS 里对应 Excel 已经不在了的 PDF，以及以前和 Excel 放在一起的「Excel 名_表名.pdf」旧 PDF；「批量导出同名 PDF」生成的 aaa.pdf 不删',
      ] },
      { title: '➡️ 边导边看', items: [
        '不用等全部导完：导出一开始按钮就变成「📐 先看导好的」，导好的那几份先在调整页面看、先改，后面的导好会接在最下面',
        '导好的 Excel 可以先打开；还没轮到的、正在导的那份要等导完',
        '在调整页面点「统一修改 / 删除标红 / 重新生成」：导出这边做完手上这一份就停一下，先改你点的，改完接着导（同一时间只有一处在用 Excel）',
      ] },
      { title: '➡️ 导完以后', items: [
        '按钮变成「📐 调整页面」，点了进调整页面',
        '末页只多一点（两页以上、最后一页印了不到 20%）的表在调整页面上自动标黄，建议向下取整；横竖、取整都在网页上点，程序去改 Excel',
      ] },
    ],
    runLabel: '📄 导出 PDF',
  },
  normalize_view: {
    title: '视图标准化',
    icon: 'target',
    theme: 'feature-mint',
    sub: '所有可见工作表开启分页预览、定位到 A1、缩放设为 100% 后保存',
    tip: '每张可见表切到分页预览、选中 A1、缩放 100%，保存时停在第一张可见表',
    notes: [
      { title: '🎯 做什么', items: [
        '所选的 Excel，和所选文件夹里（含子文件夹）的 Excel',
        '每张可见的表：切到分页预览、跳到 A1、缩放设成 100%；隐藏表不动',
        '保存时停在第一张可见表，别人打开时看到的就是这样',
      ] },
      { title: 'ℹ️ 注意', items: [
        '会直接保存 Excel，请先关掉要处理的文件',
        '上传页每份文件上传前也会自动做这一步，一般不用单独跑',
      ] },
    ],
    runLabel: '🎯 开始标准化',
  },
};

// 切字工作台发起的任务类型
const CLIP_JOB_KINDS = new Set(['scan', 'fix', 'refresh', 'recheck', 'manual_open', 'batch_scan', 'batch_commit', 'batch_recheck', 'batch_tighten', 'batch_fix', 'batch_open', 'batch_unhighlight']);
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

// 导出完点「调整页面」：调整页面读这个文件夹，先记下输入框里的路径再过去
function openPdfReview() {
  const folder = document.getElementById('taskFolderInput').value.trim();
  if (folder) saveLastFolder(overview?.job?.folder || folder);
  window.location.href = '/pdf-review';
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
  const stripDone = status === 'completed';
  document.getElementById('stripPct').textContent = stripDone ? `✅ 全部完成 · ${p.completed_files}/${p.total_files} 个文件` : `${pct}% · ${p.completed_files}/${p.total_files} 个文件`;
  document.getElementById('stripFill').style.width = `${Math.min(100, Math.max(0, pct))}%`;
  const stripTrack = document.getElementById('stripFill').parentElement;
  stripTrack.classList.toggle('is-done', stripDone);
  stripTrack.classList.toggle('is-paused', status !== 'running');
  window.ProgressEta?.update('stripFill', 'task', pct / 100, status === 'running');

  const resumeBtn = document.getElementById('stripResumeBtn');
  resumeBtn.classList.toggle('hidden', !overview.can_resume);
  resumeBtn.disabled = overview.busy;
}

function renderCards() {
  const pct = Number(overview.progress.progress_pct || 0);
  [...Object.keys(TASKS), 'clean_sheets'].forEach(kind => {
    const meta = document.getElementById('meta_' + kind);
    if (!meta) return; // 视图标准化已并进上传，只在「更多工具」里有入口，没有状态文字
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
  document.getElementById('taskDialogIcon').className = `task-dialog-icon harbor-icon icon-${task.icon}`;
  document.getElementById('taskDialogTitle').textContent = task.title;
  document.getElementById('taskDialogSub').textContent = task.sub;
  document.getElementById('taskNotesTip').textContent = task.tip || '';
  document.getElementById('taskNotesBox').open = false; // 每次打开对话框都先折叠
  document.getElementById('taskNotes').innerHTML = task.notes.map(group =>
    `<section class="how-notes-group"><h3>${escapeHtml(group.title)}</h3><ul>`
    + group.items.map(note => `<li>${escapeHtml(note)}</li>`).join('') + '</ul></section>').join('');

  const input = document.getElementById('taskFolderInput');
  const folder = (phase === 'running' && overview.job.folder) || readLastFolder() || (overview && overview.folder) || '';
  if (window.WorkbookPaths) window.WorkbookPaths.show(input, folder);
  else input.value = folder;

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

  if (phase === 'running' && activeKind === 'export_pdf' && !starting) {
    // 边导边看（用户 2026-11-12 定）：导好的先进调整页面看、先改，后面的接着在后台导
    btn.disabled = false;
    btn.textContent = '📐 先看导好的';
    hint.textContent = '导好的 PDF 已经可以在调整页面看、改了；后面的在后台接着导，导好会接在最下面';
  } else if (starting || phase === 'running') {
    btn.disabled = true;
    btn.textContent = '⏳ 正在执行...';
    hint.textContent = '可以关闭此窗口，任务会在后台继续';
  } else if (phase === 'done') {
    btn.disabled = false;
    // 调整页面：导出完一个按钮变成「调整页面」；没导完的先继续导
    btn.textContent = canResumeActiveTask() ? '🔁 继续上次任务' : activeKind === 'export_pdf' ? '📐 调整页面' : '再执行一次';
  } else {
    btn.disabled = busyOther;
    btn.textContent = task.runLabel;
  }
  // 调整页面窗口：导出 PDF 是蓝色，导完的「调整页面」是绿色，一眼分得出是哪一步
  const viewing = activeKind === 'export_pdf' && ((phase === 'done' && !canResumeActiveTask()) || (phase === 'running' && !starting));
  btn.classList.toggle('btn-success', activeKind === 'export_pdf' && !viewing);
  btn.classList.toggle('btn-primary', activeKind !== 'export_pdf' || viewing);
  if (busyOther && phase === 'form' && !starting) {
    hint.textContent = `后台正在执行「${overview.progress.task_name}」，完成后才能开始`;
  }
  if (phase === 'running' || phase === 'done') renderRunBox();
}

function renderRunBox() {
  const p = overview.progress;
  const pct = Number(p.progress_pct || 0);
  const stepText = document.getElementById('taskStepText');
  stepText.textContent = p.current_step || '准备中...';
  stepText.title = stepText.textContent;
  // 进度百分比把失败文件也计为已结束，这里分开显示完成与失败数
  document.getElementById('taskStatsText').textContent =
    `完成 ${p.completed_files}/${p.total_files}` +
    (p.error_files ? ` · 失败 ${p.error_files}` : '') + ` · ${pct}%`;
  document.getElementById('taskProgressFill').style.width = `${Math.min(100, Math.max(0, pct))}%`;
  // 结束后下面的结果框已写明完成 / 未完成，进度条收起
  document.getElementById('taskProgressFill').closest('.mini-progress-box').classList.toggle('hidden', phase === 'done');
  window.ProgressEta?.update('taskProgressFill', 'task', pct / 100, phase === 'running' && overview.busy);

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
    ${isExport && outDir ? `<span>PDF 输出位置：<code>${escapeHtml(outDir)}</code></span>` : ''}
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
  } else if ((phase === 'done' || phase === 'running') && activeKind === 'export_pdf') {
    openPdfReview();
  } else if (phase === 'done') {
    phase = 'form';
    renderDialog();
  } else {
    startTask();
  }
}

function readFolderInput() {
  const folder = document.getElementById('taskFolderInput').value.trim();
  if (!folder) appAlert('请先输入 Excel 文件或文件夹路径', { tone: 'warn' });
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
    const payload = { kind: activeKind, folder };
    const result = await postAction('/api/quick_task', payload);
    saveLastFolder(result.folder || folder);
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
    if (wasBusy && !data.busy) {
      celebrateFinished();
      refreshDoneMarks();
    }

    renderHeader();
    renderStrip();
    renderCards();
    window.JobPanel?.overview(overview); // 实时进度框：每份做完的结果
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
// 关掉对话框（调整页面 / 视图标准化）也算「回到首页」：那张卡片边框闪几下
const TASK_TILES = { export_pdf: 'adjust', normalize_view: 'view' };
taskDialog.addEventListener('close', () => {
  if (TASK_TILES[activeKind]) window.HomeTiles?.flash(TASK_TILES[activeKind]);
  activeKind = null;
});
// 点击遮罩关闭；对话框内部由子元素铺满，不会误触
taskDialog.addEventListener('click', e => {
  if (e.target === taskDialog) closeTask();
});

const taskFolderInput = document.getElementById('taskFolderInput');
taskFolderInput.addEventListener('keydown', e => {
  if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) onPrimaryAction();
});

document.getElementById('progressFileInput').addEventListener('change', e => {
  const file = e.target.files[0];
  if (file) importProgressFile(file);
});

setInterval(pollState, 1200);

/* ==========================================
   路线图卡片右上角的「做过了」标记
   按上次用的文件夹（各页面共用 lastFolder）查操作历史：在这个文件夹或它的上级文件夹做完过就打 ✓；
   只有带出错文件的那次，打黄色 !。主页打开、任务跑完、切回这个页面时刷新。
   ========================================== */
function lastFolder() {
  try {
    return (localStorage.getItem(FOLDER_STORAGE_KEY) || '').trim();
  } catch (e) {
    return '';
  }
}

async function refreshDoneMarks() {
  const folder = lastFolder();
  let done = {};
  if (folder) {
    try {
      const res = await fetch('/api/operation-history?folder=' + encodeURIComponent(folder));
      if (res.ok) done = (await res.json()).done || {};
    } catch (e) {
      return; // 读不到就保持原样，不影响使用
    }
  }
  document.querySelectorAll('.tile[data-ops]').forEach(stop => {
    const entries = stop.dataset.ops.split(' ').map(k => done[k]).filter(Boolean);
    const mark = stop.querySelector('.tile-done');
    if (!mark) return;
    if (!entries.length) {
      mark.classList.add('hidden');
      stop.classList.remove('is-done');
      return;
    }
    const completed = entries.some(e => e.completed);
    const count = entries.reduce((n, e) => n + e.count, 0);
    const last = entries.map(e => e.last).sort().pop();
    const label = completed
      ? `已做过 ${count} 次，最近一次 ${last}`
      : `做过 ${count} 次，但有出错的文件，最近一次 ${last}`;
    mark.textContent = completed ? '✓' : '!';
    mark.classList.toggle('is-warn', !completed);
    mark.classList.remove('hidden');
    mark.title = label;
    mark.setAttribute('aria-label', label);
    stop.classList.add('is-done');
  });
  const note = document.getElementById('routeDoneNote');
  note.classList.toggle('hidden', !folder);
  note.textContent = folder ? `✓ 表示在「${folder}」里已经做过` : '';
  note.title = folder;
}

refreshDoneMarks();
window.addEventListener('focus', refreshDoneMarks);
document.addEventListener('visibilitychange', () => {
  if (!document.hidden) refreshDoneMarks();
});
pollState();
const startTaskKind = new URLSearchParams(location.search).get('task');
if (TASKS[startTaskKind]) openTask(startTaskKind); // 带 ?task= 打开主页：直接打开对应对话框
