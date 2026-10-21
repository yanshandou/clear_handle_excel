const TASKS = {
  clean_sheets: {
    title: '精简 Sheet',
    icon: '🧹',
    theme: 'feature-violet',
    sub: '扫描全部工作表，自由勾选要删除的表，或将拿不准的表标记为 ❓ 待定',
    notes: [
      '自动初筛：名称含 YYYYMMDD 日期、bak / backup、old / 旧（「新旧」「復旧」除外）、バックアップ / 备份、コピー / copy / 副本或「 (2)」，且同工作簿存在原表的，默认勾选删除',
      '自由调整：直接勾选或取消复选框即可决定删除哪些工作表（每个工作簿至少保留 1 个工作表）',
      '存疑待定：拿不准的表可点「❓」标为待定——本次保留不删，并在该 Excel 文件名与该 Sheet 名前加上「❓」图标，按名称排序即可集中人工确认',
      '安全备份：修改或重命名前自动按原文件名备份到 _backup_时间戳 文件夹',
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
      '跳过隐藏表、変更履歴和备份表',
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
let sheetDecisions = {}; // { [fileName]: { [sheetName]: 'keep' | 'delete' | 'uncertain' } }
let plainDetailsOpen = false;
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

function stripUncertainEmoji(name) {
  return String(name ?? '').replace(/^❓+\s*/, '');
}

function addUncertainEmoji(name) {
  const base = stripUncertainEmoji(name);
  return base ? `❓${base}` : '❓';
}

function initSheetDecisions() {
  sheetDecisions = {};
  plainDetailsOpen = false;
  if (!previewData || !Array.isArray(previewData.files)) return;
  previewData.files.forEach(file => {
    const delSet = new Set(file.backups || []);
    const uncSet = new Set(file.uncertain || []);
    const map = {};
    (file.sheets || []).forEach(name => {
      if (delSet.has(name)) map[name] = 'delete';
      else if (uncSet.has(name)) map[name] = 'uncertain';
      else map[name] = 'keep';
    });
    sheetDecisions[file.name] = map;
  });
}

function getSheetState(fileName, sheetName) {
  return (sheetDecisions[fileName] && sheetDecisions[fileName][sheetName]) || 'keep';
}

function toggleSheetDelete(fileIndex, sheetIndex, checked) {
  const file = previewData?.files?.[fileIndex];
  if (!file) return;
  const sheetName = file.sheets[sheetIndex];
  if (sheetName === undefined) return;
  sheetDecisions[file.name] = sheetDecisions[file.name] || {};
  sheetDecisions[file.name][sheetName] = checked ? 'delete' : 'keep';
  renderPreview();
  renderDialog();
}

function toggleSheetUncertain(fileIndex, sheetIndex) {
  const file = previewData?.files?.[fileIndex];
  if (!file) return;
  const sheetName = file.sheets[sheetIndex];
  if (sheetName === undefined) return;
  const current = getSheetState(file.name, sheetName);
  sheetDecisions[file.name] = sheetDecisions[file.name] || {};
  sheetDecisions[file.name][sheetName] = current === 'uncertain' ? 'keep' : 'uncertain';
  renderPreview();
  renderDialog();
}

function resetSheetDecisions() {
  initSheetDecisions();
  renderPreview();
  renderDialog();
}

function summarizeDecisions() {
  const summary = {
    deletedTotal: 0,
    uncertainTotal: 0,
    keptTotal: 0,
    uncertainFiles: 0,
    unmarkedCount: 0,
    emptyFiles: [],
  };
  if (!previewData || !Array.isArray(previewData.files)) return summary;
  previewData.files.forEach(file => {
    let fileDel = 0;
    let fileUnc = 0;
    (file.sheets || []).forEach(name => {
      const st = getSheetState(file.name, name);
      if (st === 'delete') fileDel += 1;
      else if (st === 'uncertain') fileUnc += 1;
      else {
        summary.keptTotal += 1;
        if (String(name).startsWith('❓')) summary.unmarkedCount += 1;
      }
    });
    summary.deletedTotal += fileDel;
    summary.uncertainTotal += fileUnc;
    if (fileUnc > 0) summary.uncertainFiles += 1;
    else if (String(file.name).startsWith('❓') && file.sheets.length > 0) summary.unmarkedCount += 1;
    if (file.sheets.length > 0 && fileDel >= file.sheets.length) {
      summary.emptyFiles.push(file.name);
    }
  });
  return summary;
}

function buildCleanPlan() {
  const plan = {};
  if (!previewData || !Array.isArray(previewData.files)) return plan;
  previewData.files.forEach(file => {
    const del = [];
    const unc = [];
    (file.sheets || []).forEach(name => {
      const st = getSheetState(file.name, name);
      if (st === 'delete') del.push(name);
      else if (st === 'uncertain') unc.push(name);
    });
    plan[file.name] = { delete: del, uncertain: unc };
  });
  return plan;
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
    const stats = summarizeDecisions();
    const hasAction = stats.deletedTotal > 0 || stats.uncertainTotal > 0 || stats.unmarkedCount > 0;
    const hasError = stats.emptyFiles.length > 0;
    btn.disabled = !hasAction || hasError || busyOther;
    if (hasError) {
      btn.textContent = '⚠️ 请至少保留 1 个工作表';
      hint.textContent = `「${stats.emptyFiles[0]}」不能将所有工作表全部勾选删除`;
    } else if (stats.deletedTotal > 0 && stats.uncertainTotal > 0) {
      btn.textContent = `🧹 确认执行（删 ${stats.deletedTotal} 表 · 待定 ${stats.uncertainTotal} 表）`;
    } else if (stats.deletedTotal > 0) {
      btn.textContent = `🧹 删除 ${stats.deletedTotal} 个工作表`;
    } else if (stats.uncertainTotal > 0) {
      btn.textContent = `❓ 标记 ${stats.uncertainTotal} 个待定工作表`;
    } else if (stats.unmarkedCount > 0) {
      btn.textContent = `✨ 清除 ${stats.unmarkedCount} 处 ❓ 待定标记`;
    } else {
      btn.textContent = '没有勾选删除或待定的工作表';
    }
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

// 按工作簿中的原有顺序列出全部工作表，支持勾选删除与标记 ❓ 待定
function previewFileItem(file, fileIndex, extraClass = '') {
  let delCount = 0;
  let uncCount = 0;
  file.sheets.forEach(name => {
    const st = getSheetState(file.name, name);
    if (st === 'delete') delCount += 1;
    else if (st === 'uncertain') uncCount += 1;
  });
  const kept = file.sheets.length - delCount - uncCount;
  const allDeleted = file.sheets.length > 0 && delCount >= file.sheets.length;

  const tags = file.sheets.map((name, sheetIndex) => {
    const st = getSheetState(file.name, name);
    const isDeleted = st === 'delete';
    const isUncertain = st === 'uncertain';
    const displayName = isUncertain
      ? addUncertainEmoji(name)
      : (name.startsWith('❓') ? `${stripUncertainEmoji(name)} (去❓)` : name);
    return `
      <span class="preview-sheet-tag is-${st}">
        <label class="sheet-check-label" title="${isDeleted ? '已勾选删除（点击取消勾选）' : '勾选后删除此工作表'}">
          <input type="checkbox" class="sheet-check-input" ${isDeleted ? 'checked' : ''}
            onchange="toggleSheetDelete(${fileIndex}, ${sheetIndex}, this.checked)">
          <span class="sheet-tag-name">${escapeHtml(displayName)}</span>
        </label>
        <button type="button" class="sheet-uncertain-btn ${isUncertain ? 'active' : ''}"
          onclick="toggleSheetUncertain(${fileIndex}, ${sheetIndex})"
          title="${isUncertain ? '已标记为待定（再次点击改为普通保留）' : '拿不准要不要删？点击标为「❓ 待定」：先不删，给文件名与该 Sheet 名加 ❓ 图标'}">
          ${isUncertain ? '❓待定' : '❓'}
        </button>
      </span>`;
  }).join('');

  let renameBadge = '';
  if (uncCount > 0) {
    const targetFileName = addUncertainEmoji(file.name);
    renameBadge = `<span class="preview-rename-badge" title="含待定 Sheet，文件名将带 ❓ 前缀以便按名称排序集中查看">🏷️ ${escapeHtml(targetFileName)}</span>`;
  } else if (file.name.startsWith('❓') && file.sheets.length > 0) {
    renameBadge = `<span class="preview-rename-badge is-restore" title="已无待定 Sheet，执行后将恢复原文件名">🏷️ 恢复为 ${escapeHtml(stripUncertainEmoji(file.name))}</span>`;
  }

  const countParts = [`保留 ${kept}`];
  if (delCount) countParts.push(`<em>删除 ${delCount}</em>`);
  if (uncCount) countParts.push(`<strong class="is-uncertain-count">❓待定 ${uncCount}</strong>`);

  return `
    <li class="${extraClass} ${allDeleted ? 'is-warn' : ''}">
      <div class="preview-file-head">
        <div class="preview-file-title">
          <b>📄 ${escapeHtml(file.name)}</b>
          ${renameBadge}
        </div>
        ${file.sheets.length ? `<span class="preview-counts">${countParts.join(' · ')}</span>` : ''}
      </div>
      ${file.note ? `<span class="preview-note">⚠️ ${escapeHtml(file.note)}</span>` : ''}
      ${allDeleted ? `<span class="preview-note">⚠️ 不能删除全部工作表，请至少取消勾选 1 个或标为 ❓ 待定</span>` : ''}
      ${tags ? `<span class="preview-tags">${tags}</span>` : ''}
    </li>`;
}

function renderPreview() {
  const detailsEl = document.querySelector('#previewBox details.preview-plain');
  if (detailsEl) plainDetailsOpen = detailsEl.open;

  const files = previewData.files;
  const indexed = files.map((file, idx) => ({ file, idx }));
  const primary = indexed.filter(({ file }) =>
    file.backups.length || (file.uncertain && file.uncertain.length) || file.note ||
    file.sheets.some(s => getSheetState(file.name, s) !== 'keep')
  );
  const primarySet = new Set(primary.map(x => x.idx));
  const plain = indexed.filter(({ idx }) => !primarySet.has(idx));
  const stats = summarizeDecisions();

  document.getElementById('previewBox').innerHTML = `
    <div class="preview-summary-row">
      <div class="preview-summary">
        共 <b>${files.length}</b> 个文件：
        勾选删除 <b class="is-deleted">${stats.deletedTotal}</b> 个工作表 ·
        标记待定 <b class="is-uncertain">${stats.uncertainTotal}</b> 个${stats.uncertainFiles ? `（<b>${stats.uncertainFiles}</b> 个文件将加 ❓ 前缀）` : ''} ·
        保留 <b>${stats.keptTotal}</b> 个
      </div>
      <button type="button" class="btn btn-outline btn-sm preview-reset-btn" onclick="resetSheetDecisions()" title="恢复为系统自动识别的初始勾选状态">
        ↺ 恢复默认勾选
      </button>
    </div>
    <div class="preview-legend">
      <span class="preview-sheet-tag is-deleted"><span class="sheet-tag-name">☑️ 勾选 = 删除</span></span>
      <span class="preview-sheet-tag is-kept"><span class="sheet-tag-name">⬜ 未勾 = 保留</span></span>
      <span class="preview-sheet-tag is-uncertain"><span class="sheet-tag-name">❓ 点击问号 = 待定（不删，文件名与 Sheet 加 ❓）</span></span>
    </div>
    <ul class="preview-list">
      ${primary.map(({ file, idx }) => previewFileItem(file, idx, file.note ? 'is-warn' : '')).join('')}
    </ul>
    ${plain.length ? `
      <details class="preview-plain" ${plainDetailsOpen ? 'open' : ''} ontoggle="plainDetailsOpen = this.open">
        <summary>其余 ${plain.length} 个文件未检出备份表（点击展开可手动勾选删除或标记 ❓ 待定）</summary>
        <ul class="preview-list">${plain.map(({ file, idx }) => previewFileItem(file, idx)).join('')}</ul>
      </details>` : ''}
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
    initSheetDecisions();
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
  const payload = { kind: activeKind, folder };
  if (activeKind === 'clean_sheets' && previewData) {
    payload.plan = buildCleanPlan();
  }
  try {
    await postAction('/api/quick_task', payload);
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
    alert(e.message);
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
  if (!confirm(message)) return;
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
    alert(lines.join('\n'));
  } catch (e) {
    alert(`导入失败：${e.message}`);
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

document.getElementById('progressFileInput').addEventListener('change', e => {
  const file = e.target.files[0];
  if (file) importProgressFile(file);
});

setInterval(pollState, 1200);
pollState();
