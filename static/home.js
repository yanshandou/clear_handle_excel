const TASKS = {
  clean_sheets: {
    title: '精简 Sheet',
    icon: '🧹',
    theme: 'feature-violet',
    sub: '扫描全部工作表（含子文件夹），单击标签即可在「保留 / 删除 / 待定」三种状态间切换',
    notes: [
      '自动递归扫描所选文件夹及其所有子文件夹下的 Excel 文件',
      '自动初筛：名称含 YYYYMMDD 日期、bak / backup、old / 旧（「新旧」「復旧」除外）、バックアップ / 备份、コピー / copy / 副本或「 (2)」，且同工作簿存在原表的，默认标为「删除」，并用马克笔下划线标出原表与备份表的同名部分',
      '单击切换：直接单击任意工作表标签，即可在「保留 → 删除 → 待定 → 保留」之间循环切换，仅靠颜色区分状态（每个工作簿至少保留 1 个工作表）',
      '存疑待定：标为「待定」的工作表本次保留不删，并在执行后为该 Excel 文件名与该 Sheet 名前加上「❓」图标，按名称排序即可集中人工确认',
    ],
    preview: true,
  },
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

function splitDirAndBase(name) {
  const str = String(name ?? '');
  const idx = Math.max(str.lastIndexOf('/'), str.lastIndexOf('\\'));
  return idx >= 0 ? [str.slice(0, idx + 1), str.slice(idx + 1)] : ['', str];
}

function hasUncertainEmoji(name) {
  const [, base] = splitDirAndBase(name);
  return base.startsWith('❓');
}

function stripUncertainEmoji(name) {
  const [dir, base] = splitDirAndBase(name);
  return dir + base.replace(/^❓+\s*/, '');
}

function addUncertainEmoji(name) {
  const [dir, base] = splitDirAndBase(name);
  const clean = base.replace(/^❓+\s*/, '');
  return dir + (clean ? `❓${clean}` : '❓');
}

/* ==========================================
   同名原表与备份表的马克笔下划线定位
   ========================================== */
const MARK_PALETTE_SIZE = 5;

function isValidYmdDate(yyyymmdd) {
  if (!/^(?:19|20)\d{6}$/.test(yyyymmdd)) return false;
  const y = Number(yyyymmdd.slice(0, 4));
  const m = Number(yyyymmdd.slice(4, 6));
  const d = Number(yyyymmdd.slice(6, 8));
  if (m < 1 || m > 12 || d < 1 || d > 31) return false;
  const dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d;
}

function hasBackupMarker(sheetName) {
  const text = String(sheetName ?? '').normalize('NFKC');
  if (/(?<![a-z])(?:bak|backup|old|copy)(?![a-z])/i.test(text)) return true;
  if (/バックアップ|备份|コピー|副本|(?<![新復])旧/.test(text)) return true;
  if (/\s\(\d+\)$/.test(text)) return true;
  const dateRe = /(?<!\d)((?:19|20)\d{6})(?!\d)/g;
  let m;
  while ((m = dateRe.exec(text)) !== null) {
    if (isValidYmdDate(m[1])) return true;
  }
  return false;
}

function buildMaskFromCoreIndices(len, coreIndices, markerRemovedOrig) {
  const mask = new Array(len).fill(false);
  if (!coreIndices.length) return mask;
  const sorted = Array.from(new Set(coreIndices)).sort((a, b) => a - b);
  sorted.forEach(idx => {
    if (idx >= 0 && idx < len) mask[idx] = true;
  });
  // 若两个核心字符之间只有普通空白或连接符（不含被剔除的日期/备份词/版本号），连成一整段下划线
  for (let i = 0; i < sorted.length - 1; i++) {
    const a = sorted[i];
    const b = sorted[i + 1];
    if (b > a + 1) {
      let hasRemovedMarker = false;
      for (let p = a + 1; p < b; p++) {
        if (markerRemovedOrig.has(p)) {
          hasRemovedMarker = true;
          break;
        }
      }
      if (!hasRemovedMarker) {
        for (let p = a + 1; p < b; p++) mask[p] = true;
      }
    }
  }
  return mask;
}

function analyzeSheetName(displayName, { stripBackup = false, stripVersion = true } = {}) {
  const str = String(displayName ?? '');
  let items = [];
  for (let i = 0; i < str.length; i++) {
    const normPiece = str[i].normalize('NFKC');
    for (let j = 0; j < normPiece.length; j++) {
      items.push({ ch: normPiece[j], origIdx: i });
    }
  }
  const markerRemovedOrig = new Set();

  const applyStrip = (regex, predicate = null, isMarker = true) => {
    const currentStr = items.map(x => x.ch).join('');
    const removePos = new Set();
    const re = new RegExp(regex.source, regex.flags.includes('g') ? regex.flags : regex.flags + 'g');
    let m;
    while ((m = re.exec(currentStr)) !== null) {
      if (!predicate || predicate(m)) {
        for (let k = m.index; k < m.index + m[0].length; k++) {
          removePos.add(k);
          if (isMarker) markerRemovedOrig.add(items[k].origIdx);
        }
      }
    }
    if (removePos.size > 0) {
      items = items.filter((_, idx) => !removePos.has(idx));
    }
  };

  if (stripBackup) {
    applyStrip(/(?<!\d)((?:19|20)\d{6})(?!\d)/g, m => isValidYmdDate(m[1]), true);
    applyStrip(/\s\(\d+\)$/g, null, true);
    applyStrip(/copy\s+of|の?コピー|的?副本/gi, null, true);
    applyStrip(/(?<![a-z])(?:bak|backup|old|copy)(?![a-z])/gi, null, true);
    applyStrip(/バックアップ|备份|コピー|副本|(?<![新復])旧/g, null, true);
  }
  if (stripVersion) {
    applyStrip(/(?<![a-z])v(?:er)?\.?\d+(?:\.\d+)*(?![a-z])/gi, null, true);
  }
  applyStrip(/[\s_\-・.,、。()\[\]【】「」『』❓]+/g, null, false);

  const key = items.map(x => x.ch.toLowerCase()).join('');
  const coreOrigIndices = items.map(x => x.origIdx);
  const mask = buildMaskFromCoreIndices(str.length, coreOrigIndices, markerRemovedOrig);
  return { key, items, markerRemovedOrig, mask };
}

function computeFileSheetHighlights(file) {
  const sheets = file.sheets || [];
  const entries = sheets.map((rawName, idx) => {
    const cleanName = stripUncertainEmoji(rawName);
    const state = getSheetState(file.name, rawName);
    const isCandidate = hasBackupMarker(cleanName);
    return {
      idx,
      rawName,
      cleanName,
      state,
      isCandidate,
      strict: analyzeSheetName(cleanName, { stripBackup: isCandidate, stripVersion: false }),
      loose: analyzeSheetName(cleanName, { stripBackup: isCandidate, stripVersion: true }),
      baseLoose: analyzeSheetName(cleanName, { stripBackup: true, stripVersion: true }),
    };
  });

  const highlights = new Array(sheets.length).fill(null);
  let nextColorIdx = 0;

  // 1. 按原表与备份表规则配对：共享同一个去掉备份标记后的基础名称
  const byBaseKey = new Map();
  entries.forEach(entry => {
    const key = entry.baseLoose.key;
    if (!key) return;
    if (!byBaseKey.has(key)) byBaseKey.set(key, []);
    byBaseKey.get(key).push(entry);
  });

  byBaseKey.forEach(group => {
    if (group.length < 2) return;
    const hasOrigAndCandidate = group.some(e => !e.isCandidate && e.loose.key === e.baseLoose.key) &&
      group.some(e => e.isCandidate);
    const hasKeepAndAction = group.some(e => e.state === 'keep') &&
      group.some(e => e.state !== 'keep');
    if (!hasOrigAndCandidate && !hasKeepAndAction) return;

    const colorIdx = nextColorIdx % MARK_PALETTE_SIZE;
    nextColorIdx += 1;
    const firstStrictKey = group[0].strict.key;
    const allSameStrict = Boolean(firstStrictKey) && group.every(e => e.strict.key === firstStrictKey);

    group.forEach(entry => {
      const chosen = allSameStrict
        ? entry.strict
        : (entry.loose.key === entry.baseLoose.key ? entry.loose : entry.baseLoose);
      if (chosen.mask.some(Boolean)) {
        highlights[entry.idx] = { mask: chosen.mask, colorIdx };
      }
    });
  });

  // 2. 用户手动切换为删除/待定的工作表：若与某张保留表存在包含关系的共有名称（>=2 字），也标出共有部分
  entries.forEach(entry => {
    if (highlights[entry.idx] || entry.state === 'keep') return;
    const targetKey = entry.baseLoose.key;
    if (targetKey.length < 2) return;

    let bestKeep = null;
    let bestLen = 0;
    let keepInTargetIdx = -1;
    let targetInKeepIdx = -1;

    entries.forEach(other => {
      if (other.idx === entry.idx || other.state !== 'keep') return;
      const otherKey = other.baseLoose.key;
      if (otherKey.length < 2) return;
      if (targetKey.includes(otherKey) && otherKey.length > bestLen) {
        bestKeep = other;
        bestLen = otherKey.length;
        keepInTargetIdx = targetKey.indexOf(otherKey);
        targetInKeepIdx = 0;
      } else if (otherKey.includes(targetKey) && targetKey.length > bestLen) {
        bestKeep = other;
        bestLen = targetKey.length;
        keepInTargetIdx = 0;
        targetInKeepIdx = otherKey.indexOf(targetKey);
      }
    });

    if (bestKeep && bestLen >= 2) {
      const colorIdx = highlights[bestKeep.idx]
        ? highlights[bestKeep.idx].colorIdx
        : (nextColorIdx++ % MARK_PALETTE_SIZE);
      const entrySlice = entry.baseLoose.items
        .slice(keepInTargetIdx, keepInTargetIdx + bestLen)
        .map(x => x.origIdx);
      const entryMask = buildMaskFromCoreIndices(
        entry.cleanName.length,
        entrySlice,
        entry.baseLoose.markerRemovedOrig,
      );
      if (entryMask.some(Boolean)) {
        highlights[entry.idx] = { mask: entryMask, colorIdx };
      }
      if (!highlights[bestKeep.idx]) {
        const keepSlice = bestKeep.baseLoose.items
          .slice(targetInKeepIdx, targetInKeepIdx + bestLen)
          .map(x => x.origIdx);
        const keepMask = buildMaskFromCoreIndices(
          bestKeep.cleanName.length,
          keepSlice,
          bestKeep.baseLoose.markerRemovedOrig,
        );
        if (keepMask.some(Boolean)) {
          highlights[bestKeep.idx] = { mask: keepMask, colorIdx };
        }
      }
    }
  });

  return highlights;
}

function renderSheetLabelHtml(displayName, highlight) {
  if (!highlight || !highlight.mask || !highlight.mask.some(Boolean)) {
    return escapeHtml(displayName);
  }
  const { mask, colorIdx } = highlight;
  let html = '';
  let i = 0;
  while (i < displayName.length) {
    const marked = Boolean(mask[i]);
    let j = i + 1;
    while (j < displayName.length && Boolean(mask[j]) === marked) {
      j += 1;
    }
    const textChunk = escapeHtml(displayName.slice(i, j));
    if (marked) {
      html += `<span class="sheet-shared-mark mark-${colorIdx}">${textChunk}</span>`;
    } else {
      html += textChunk;
    }
    i = j;
  }
  return html;
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

const NEXT_SHEET_STATE = {
  keep: 'delete',
  delete: 'uncertain',
  uncertain: 'keep',
};

function getSheetState(fileName, sheetName) {
  return (sheetDecisions[fileName] && sheetDecisions[fileName][sheetName]) || 'keep';
}

function cycleSheetState(fileIndex, sheetIndex) {
  const file = previewData?.files?.[fileIndex];
  if (!file) return;
  const sheetName = file.sheets[sheetIndex];
  if (sheetName === undefined) return;
  const current = getSheetState(file.name, sheetName);
  sheetDecisions[file.name] = sheetDecisions[file.name] || {};
  sheetDecisions[file.name][sheetName] = NEXT_SHEET_STATE[current] || 'keep';
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
    else if (hasUncertainEmoji(file.name) && file.sheets.length > 0) summary.unmarkedCount += 1;
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
      hint.textContent = `「${stats.emptyFiles[0]}」不能将所有工作表全部标为删除`;
    } else if (stats.deletedTotal > 0 && stats.uncertainTotal > 0) {
      btn.textContent = `🧹 确认执行（删 ${stats.deletedTotal} 表 · 待定 ${stats.uncertainTotal} 表）`;
    } else if (stats.deletedTotal > 0) {
      btn.textContent = `🧹 删除 ${stats.deletedTotal} 个工作表`;
    } else if (stats.uncertainTotal > 0) {
      btn.textContent = `❓ 标记 ${stats.uncertainTotal} 个待定工作表`;
    } else if (stats.unmarkedCount > 0) {
      btn.textContent = `✨ 清除 ${stats.unmarkedCount} 处 ❓ 待定标记`;
    } else {
      btn.textContent = '没有标记删除或待定的工作表';
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

// 按工作簿中的原有顺序列出全部工作表，单击标签循环切换「保留 / 删除 / 待定」
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
  const highlights = computeFileSheetHighlights(file);

  const tags = file.sheets.map((name, sheetIndex) => {
    const st = getSheetState(file.name, name);
    const isDeleted = st === 'delete';
    const isUncertain = st === 'uncertain';
    const displayName = stripUncertainEmoji(name);
    const labelHtml = renderSheetLabelHtml(displayName, highlights[sheetIndex]);
    const tip = isDeleted
      ? '当前：删除（单击切换为「待定」，再点切换为「保留」）'
      : (isUncertain
        ? '当前：待定，本次不删并在执行后给文件名与 Sheet 加 ❓（单击切换为「保留」）'
        : '当前：保留（单击切换为「删除」）');
    return `<button type="button" class="preview-sheet-tag is-${st}" onclick="cycleSheetState(${fileIndex}, ${sheetIndex})" title="${tip}">${labelHtml}</button>`;
  }).join('');

  let renameBadge = '';
  if (uncCount > 0) {
    const targetFileName = addUncertainEmoji(file.name);
    renameBadge = `<span class="preview-rename-badge" title="含待定 Sheet，文件名将带 ❓ 前缀以便按名称排序集中查看">🏷️ ${escapeHtml(targetFileName)}</span>`;
  } else if (hasUncertainEmoji(file.name) && file.sheets.length > 0) {
    renameBadge = `<span class="preview-rename-badge is-restore" title="已无待定 Sheet，执行后将恢复原文件名">🏷️ 恢复为 ${escapeHtml(stripUncertainEmoji(file.name))}</span>`;
  }

  const countParts = [`保留 ${kept}`];
  if (delCount) countParts.push(`<em>删除 ${delCount}</em>`);
  if (uncCount) countParts.push(`<strong class="is-uncertain-count">待定 ${uncCount}</strong>`);

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
      ${allDeleted ? `<span class="preview-note">⚠️ 不能删除全部工作表，请单击标签至少将 1 个改为「保留」或「待定」</span>` : ''}
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
        删除 <b class="is-deleted">${stats.deletedTotal}</b> 个 ·
        待定 <b class="is-uncertain">${stats.uncertainTotal}</b> 个${stats.uncertainFiles ? `（<b>${stats.uncertainFiles}</b> 个文件加 ❓）` : ''} ·
        保留 <b>${stats.keptTotal}</b> 个
      </div>
      <button type="button" class="btn btn-outline btn-sm preview-reset-btn" onclick="resetSheetDecisions()" title="恢复为系统自动识别的初始标记状态">
        ↺ 恢复默认
      </button>
    </div>
    <div class="preview-legend">
      <span class="preview-sheet-tag is-keep">保留</span>
      <span>→</span>
      <span class="preview-sheet-tag is-delete">删除</span>
      <span>→</span>
      <span class="preview-sheet-tag is-uncertain">待定</span>
      <span class="preview-legend-sep">｜</span>
      <span><span class="sheet-shared-mark mark-0">同名部分</span> 为原表与备份表共有的名称（单击标签循环切换状态）</span>
    </div>
    <ul class="preview-list">
      ${primary.map(({ file, idx }) => previewFileItem(file, idx, file.note ? 'is-warn' : '')).join('')}
    </ul>
    ${plain.length ? `
      <details class="preview-plain" ${plainDetailsOpen ? 'open' : ''} ontoggle="plainDetailsOpen = this.open">
        <summary>其余 ${plain.length} 个文件未检出备份表（点击展开，单击标签可改为删除或待定）</summary>
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
