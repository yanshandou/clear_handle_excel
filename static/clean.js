// 精简 Sheet 独立页面：填写文件夹 → 扫描预览 → 单击标签调整 → 执行 → 查看结果
const CLEAN_KIND = 'clean_sheets';
const FOLDER_STORAGE_KEY = 'clear_handle_excel.lastFolder';

let overview = null;
let phase = 'form'; // form | preview | running | done
let pending = null; // scan | start：请求已发出、后台尚未确认完成，按钮保持「扫描中 / 整理中」
let pendingSince = 0;
let previewData = null;
let sheetDecisions = {}; // { [fileName]: { [sheetName]: 'keep' | 'delete' | 'uncertain' } }
let processedFiles = new Set(); // 已在预览窗口单独执行过的文件：卡片原地锁定，执行时跳过，重新扫描后才更新
let previewFromFile = null; // 最近一次打开预览窗口的文件：卡片边框发亮，从预览窗口返回时滚过去
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

function splitDirAndBase(name) {
  const str = String(name ?? '');
  const idx = Math.max(str.lastIndexOf('/'), str.lastIndexOf('\\'));
  return idx >= 0 ? [str.slice(0, idx + 1), str.slice(idx + 1)] : ['', str];
}

function stripUncertainEmoji(name) {
  const [dir, base] = splitDirAndBase(name);
  return dir + base.replace(/^❓+\s*/, '');
}

/* ==========================================
   同名原表与备份表的马克笔下划线定位
   ========================================== */
const MARK_PALETTE_SIZE = 5;

// 20220929，或用同一种分隔符隔开的 2018-02-26 / 2018.02.26 / 2018/02/26
const DATE_RE = /(?<!\d)((?:19|20)\d{2})([-/.]?)(\d{2})\2(\d{2})(?!\d)/g;

// 名称末尾不带年份的 4 位月日（MMDD），如「請求書1)0513」
const MMDD_SUFFIX_RE = /(?<!\d)(\d{2})(\d{2})[\s_\-・.]*$/;

function hasMmddSuffix(text) {
  const m = MMDD_SUFFIX_RE.exec(text);
  return Boolean(m) && isValidYmdDate(`2000${m[1]}${m[2]}`);
}

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
  if (/\s*\(\d+\)[\s_\-・.]*$/.test(text)) return true;
  if (hasMmddSuffix(text)) return true;
  const dateRe = new RegExp(DATE_RE.source, 'g');
  let m;
  while ((m = dateRe.exec(text)) !== null) {
    if (isValidYmdDate(m[1] + m[3] + m[4])) return true;
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
    applyStrip(DATE_RE, m => isValidYmdDate(m[1] + m[3] + m[4]), true);
    applyStrip(/\s*\(\d+\)[\s_\-・.]*$/g, null, true);
    applyStrip(/copy\s+of|の?コピー|的?副本/gi, null, true);
    applyStrip(/(?<![a-z])(?:bak|backup|old|copy)(?![a-z])/gi, null, true);
    applyStrip(/バックアップ|备份|コピー|副本|(?<![新復])旧/g, null, true);
    applyStrip(/(?<!\d)(\d{2})(\d{2})[\s_\-・.]*$/g, m => isValidYmdDate(`2000${m[1]}${m[2]}`), true);
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

// 「旧」表全部标为删除（或本来就没有「旧」表，olds 为空）、「新」表保留时，执行后「新」表去掉开头的「新」字
// （与后端 _new_prefix_renames 一致）；去掉后与剩下的表重名时不改名
function findNewPrefixRenames(sheets, stateOf, pairs) {
  const renamed = new Set();
  const remaining = sheets.filter(n => stateOf(n) !== 'delete');
  Object.entries(pairs || {}).forEach(([newName, olds]) => {
    if (stateOf(newName) !== 'keep' || !olds.every(o => stateOf(o) === 'delete')) return;
    const target = newName.replace(/^新\s*/, '');
    if (!target || target === newName) return;
    const taken = remaining.some(n => n !== newName && n.toLowerCase() === target.toLowerCase());
    if (!taken) renamed.add(newName);
  });
  return renamed;
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
  processedFiles = new Set();
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
  broadcastStates(file.name);
  renderPreview();
  renderPage();
}

function resetSheetDecisions() {
  const processed = processedFiles;
  initSheetDecisions();
  processedFiles = processed;
  syncChannel?.postMessage({ type: 'rescanned', folder: previewData?.folder });
  renderPreview();
  renderPage();
}

/* ==========================================
   与 Sheet 预览确认窗口同步（同一浏览器内的其他标签页）
   ========================================== */
const syncChannel = 'BroadcastChannel' in window ? new BroadcastChannel('clear_handle_excel.clean') : null;

function broadcastStates(fileName) {
  if (!syncChannel || !previewData) return;
  syncChannel.postMessage({
    type: 'states', folder: previewData.folder, file: fileName, states: sheetDecisions[fileName] || {},
  });
}

function openSheetPreview(fileIndex) {
  const file = previewData?.files?.[fileIndex];
  if (!file) return;
  const query = new URLSearchParams({ folder: previewData.folder, file: file.name });
  window.open(`/clean/preview?${query}`, '_blank');
  previewFromFile = file.name;
  renderPreview(); // 只加发亮边框（box-shadow），卡片不移动
}

// 从预览窗口返回：来源卡片不在视野内时滚到中间，并让边框闪几下
function revealReturnedCard(fileName) {
  previewFromFile = fileName;
  renderPreview();
  const idx = previewData.files.findIndex(f => f.name === fileName);
  const card = document.querySelector(`#previewBox li[data-file-index="${idx}"]`);
  if (!card) return;
  const rect = card.getBoundingClientRect();
  const bar = document.getElementById('cleanActionBar');
  const bottom = bar && !bar.classList.contains('hidden') ? bar.getBoundingClientRect().top : window.innerHeight;
  if (rect.top < 80 || rect.bottom > bottom) card.scrollIntoView({ block: 'center', behavior: 'smooth' });
  card.classList.add('is-return-pulse');
}

if (syncChannel) {
  syncChannel.onmessage = event => {
    const msg = event.data || {};
    if (phase !== 'preview' || !previewData || msg.folder !== previewData.folder) return;
    const file = previewData.files.find(f => f.name === msg.file);
    if (!file || processedFiles.has(file.name)) return;
    if (msg.type === 'hello') {
      broadcastStates(file.name);
    } else if (msg.type === 'states' && msg.states) {
      const map = sheetDecisions[file.name] = sheetDecisions[file.name] || {};
      file.sheets.forEach(name => {
        if (['keep', 'delete', 'uncertain'].includes(msg.states[name])) map[name] = msg.states[name];
      });
      renderPreview();
      renderPage();
    } else if (msg.type === 'return') {
      window.focus();
      revealReturnedCard(file.name);
    } else if (msg.type === 'done') {
      processedFiles.add(file.name);
      renderPreview();
      renderPage();
    }
  };
}

function summarizeDecisions() {
  const summary = {
    deletedTotal: 0,
    uncertainTotal: 0,
    keptTotal: 0,
    emptyFiles: [],
  };
  if (!previewData || !Array.isArray(previewData.files)) return summary;
  previewData.files.forEach(file => {
    if (processedFiles.has(file.name)) return;
    let fileDel = 0;
    let fileUnc = 0;
    (file.sheets || []).forEach(name => {
      const st = getSheetState(file.name, name);
      if (st === 'delete') fileDel += 1;
      else if (st === 'uncertain') fileUnc += 1;
      else summary.keptTotal += 1;
    });
    summary.deletedTotal += fileDel;
    summary.uncertainTotal += fileUnc;
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
    if (processedFiles.has(file.name)) return;
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

function isCleanRunning() {
  return Boolean(overview && overview.busy && overview.job && overview.job.kind === CLEAN_KIND);
}

function isBusyOther() {
  return Boolean(overview && overview.busy && !isCleanRunning());
}

function canResumeClean() {
  return Boolean(overview && overview.can_resume && overview.job && overview.job.kind === CLEAN_KIND);
}

function elapsedText() {
  const secs = Math.floor((Date.now() - pendingSince) / 1000);
  return secs >= 1 ? ` ${secs}s` : '';
}

function setPending(kind) {
  pending = kind;
  pendingSince = Date.now();
  renderPage();
}

/* ==========================================
   页面渲染
   ========================================== */
function renderHeader() {
  const p = overview.progress;
  const pct = Number(p.progress_pct || 0);
  document.getElementById('navProgressBadge').textContent =
    p.total_files > 0 ? `${p.completed_files}/${p.total_files} (${pct}%)` : '0%';
  document.getElementById('statusText').innerText = overview.busy
    ? '⏳ 后台处理中...' : (overview.can_resume ? '任务已中断，可继续上次任务' : '✅ 就绪');
}

function renderPage() {
  const working = pending !== null || phase === 'running';
  const scanBtn = document.getElementById('scanBtn');
  const execBtn = document.getElementById('execBtn');
  const hint = document.getElementById('cleanHint');

  document.getElementById('cleanFolderInput').disabled = working;
  document.getElementById('cleanActionBar').classList.toggle('hidden', phase !== 'preview');
  document.getElementById('cleanLegend').classList.toggle('hidden', phase !== 'preview');
  document.getElementById('previewBox').classList.toggle('hidden', phase !== 'preview');
  document.getElementById('previewBox').classList.toggle('is-busy', pending !== null);
  document.getElementById('cleanRunBox').classList.toggle('hidden', phase !== 'running' && phase !== 'done');
  document.getElementById('cleanEmpty').classList.toggle('hidden', phase !== 'form' || pending === 'scan');
  document.getElementById('cleanScanning').classList.toggle('hidden', pending !== 'scan');

  if (pending === 'scan') {
    scanBtn.disabled = true;
    scanBtn.textContent = `⏳ 扫描中…${elapsedText()}`;
    document.getElementById('cleanScanningText').textContent =
      `正在递归读取所有 Excel 的工作表，文件较多时需要一些时间…${elapsedText()}`;
  } else {
    // 预览只读取文件，后台忙于其他任务时也允许
    scanBtn.disabled = working;
    scanBtn.textContent = phase === 'form' ? '🔎 扫描备份表' : '🔄 重新扫描';
  }

  hint.textContent = '';
  if (phase === 'preview') {
    const stats = summarizeDecisions();
    renderSummary(stats);
    const hasAction = stats.deletedTotal > 0; // 待定只是本次不删，不改文件
    const hasError = stats.emptyFiles.length > 0;
    execBtn.disabled = working || !hasAction || hasError || isBusyOther() || isCleanRunning();
    if (pending === 'start') {
      execBtn.textContent = `⏳ 整理中…${elapsedText()}`;
    } else if (hasError) {
      execBtn.textContent = '⚠️ 请至少保留 1 个工作表';
      hint.textContent = `「${stats.emptyFiles[0]}」不能将所有工作表全部标为删除`;
    } else if (stats.deletedTotal > 0) {
      execBtn.textContent = `🧹 删除 ${stats.deletedTotal} 个工作表`;
    } else {
      execBtn.textContent = '没有标记删除的工作表';
    }
    if (isBusyOther()) {
      hint.textContent = `后台正在执行「${overview.progress.task_name}」，完成后才能开始`;
    } else if (isCleanRunning()) {
      hint.textContent = '预览窗口正在处理文件，完成后才能开始';
    }
  }
  if (phase === 'running' || phase === 'done') renderRunBox();
}

function renderSummary(stats) {
  const files = previewData.files;
  document.getElementById('cleanSummary').innerHTML = `
    共 <b>${files.length}</b> 个文件：
    删除 <b class="is-deleted">${stats.deletedTotal}</b> 个 ·
    待定 <b class="is-uncertain">${stats.uncertainTotal}</b> 个（不删） ·
    保留 <b>${stats.keptTotal}</b> 个
  `;
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
  const processed = processedFiles.has(file.name);
  const allDeleted = !processed && file.sheets.length > 0 && delCount >= file.sheets.length;
  const highlights = computeFileSheetHighlights(file);
  const newRenames = findNewPrefixRenames(file.sheets, n => getSheetState(file.name, n), file.new_old_pairs);

  const tags = file.sheets.map((name, sheetIndex) => {
    const st = getSheetState(file.name, name);
    const isDeleted = st === 'delete';
    const isUncertain = st === 'uncertain';
    const displayName = stripUncertainEmoji(name);
    const highlight = highlights[sheetIndex];
    // 删除线只划掉「新」字，标签宽度不变，切换旧表状态时卡片不会跳动
    const stripNew = newRenames.has(name);
    const labelHtml = stripNew
      ? '<s class="strip-new-prefix">新</s>' + renderSheetLabelHtml(
        displayName.slice(1), highlight && { ...highlight, mask: highlight.mask.slice(1) })
      : renderSheetLabelHtml(displayName, highlight);
    const renameTip = stripNew ? `；执行后去掉「新」字，改名为「${displayName.replace(/^新\s*/, '')}」` : '';
    const tip = (isDeleted
      ? '当前：删除（单击切换为「待定」，再点切换为「保留」）'
      : (isUncertain
        ? '当前：待定，本次不删，等确认后再处理（单击切换为「保留」）'
        : '当前：保留（单击切换为「删除」）')) + renameTip;
    if (processed) {
      return `<span class="preview-sheet-tag is-${st} is-processed" title="已在预览窗口处理，重新扫描后更新">${labelHtml}</span>`;
    }
    return `<button type="button" class="preview-sheet-tag is-${st}" onclick="cycleSheetState(${fileIndex}, ${sheetIndex})" title="${tip}">${labelHtml}</button>`;
  }).join('');

  const countParts = [`保留 ${kept}`];
  if (delCount) countParts.push(`<em>删除 ${delCount}</em>`);
  if (uncCount) countParts.push(`<strong class="is-uncertain-count">待定 ${uncCount}</strong>`);

  return `
    <li class="${extraClass} ${allDeleted ? 'is-warn' : ''} ${file.name === previewFromFile ? 'is-preview-from' : ''}" data-file-index="${fileIndex}">
      <div class="preview-file-head">
        <div class="preview-file-title">
          <b title="${escapeHtml(file.name)}">${excelLinkHtml(`📄 ${escapeHtml(splitDirAndBase(file.name)[1])}`, { file: file.name, folder: previewData.folder })}</b>
        </div>
        ${file.sheets.length ? `<span class="preview-file-actions">
          ${processed
            ? '<span class="preview-counts is-processed">✅ 已在预览窗口处理</span>'
            : `<span class="preview-counts">${countParts.join(' · ')}</span>
               <button type="button" class="btn btn-outline btn-sm preview-open-btn" onclick="openSheetPreview(${fileIndex})" title="在新标签页里看每个工作表的预览图并标记，标记完成后返回这里统一执行">🖼 预览</button>`}
        </span>` : ''}
      </div>
      ${file.note ? `<span class="preview-note">⚠️ ${escapeHtml(file.note)}</span>` : ''}
      ${allDeleted ? `<span class="preview-note">⚠️ 不能删除全部工作表，请单击标签至少将 1 个改为「保留」或「待定」</span>` : ''}
      ${tags ? `<span class="preview-tags">${tags}</span>` : ''}
    </li>`;
}

// 按所在文件夹分组：文件夹路径只在组标题显示一次，卡片里只显示文件名
function renderFolderGroups(items, extraClassOf = () => '') {
  const groups = new Map();
  items.forEach(item => {
    const [dir] = splitDirAndBase(item.file.name);
    if (!groups.has(dir)) groups.set(dir, []);
    groups.get(dir).push(item);
  });
  return Array.from(groups, ([dir, group]) => `
    <div class="clean-folder-group">
      <div class="clean-folder-title" title="${escapeHtml(dir || '所选文件夹')}">
        📁 ${escapeHtml(dir ? dir.replace(/[\\/]$/, '') : '（所选文件夹）')}
        <span>${group.length} 个文件</span>
      </div>
      <ul class="preview-list clean-grid">
        ${group.map(({ file, idx }) => previewFileItem(file, idx, extraClassOf(file))).join('')}
      </ul>
    </div>`).join('');
}

// 瀑布流：每张卡片按自身高度占用若干 1px 的网格行，矮卡片下方不再留空，且保持从左到右的阅读顺序
const MASONRY_GAP = 8;

// 先量完所有卡片高度再统一写入：卡片是顶端对齐，高度只取决于内容，不需要先清掉 span。
// 以前是逐张「清 span → 量高度」，量的时候整片网格塌成 1px 一行，页面瞬间变矮，
// 浏览器把滚动位置夹回顶部附近（页面滚得越靠下越容易发生）。
function layoutMasonry() {
  const grids = Array.from(document.querySelectorAll('#previewBox .clean-grid'))
    .filter(grid => grid.offsetParent); // 折叠中的列表展开时再排
  const spans = grids.map(grid => Array.from(grid.querySelectorAll(':scope > li'), li =>
    [li, Math.ceil(li.getBoundingClientRect().height + MASONRY_GAP)]));
  grids.forEach((grid, i) => {
    grid.classList.add('is-masonry');
    spans[i].forEach(([li, span]) => { li.style.gridRowEnd = `span ${span}`; });
  });
}

// 全部文件按文件夹放在一起，按扫描顺序排列；没有检出备份表的文件所有标签默认为保留，单击标签不会改变卡片位置
function renderPreview() {
  const indexed = previewData.files.map((file, idx) => ({ file, idx }));
  const box = document.getElementById('previewBox');
  const scrollY = window.scrollY;
  // 重建期间先撑住原高度，页面不会因为暂时变矮而被拉回顶部
  box.style.minHeight = `${box.offsetHeight}px`;
  box.innerHTML = renderFolderGroups(indexed, file => (file.note ? 'is-warn' : ''));
  layoutMasonry();
  box.style.minHeight = '';
  if (window.scrollY !== scrollY) window.scrollTo(window.scrollX, scrollY);
}

function renderRunBox() {
  const p = overview.progress;
  const pct = Number(p.progress_pct || 0);
  document.getElementById('cleanStepText').textContent = p.current_step || '准备中...';
  // 进度百分比把失败文件也计为已结束，这里分开显示完成与失败数
  document.getElementById('cleanStatsText').textContent =
    `完成 ${p.completed_files}/${p.total_files}` +
    (p.error_files ? ` · 失败 ${p.error_files}` : '') + ` · ${pct}%`;
  document.getElementById('cleanProgressFill').style.width = `${Math.min(100, Math.max(0, pct))}%`;

  const runBtn = document.getElementById('runStateBtn');
  runBtn.classList.toggle('hidden', phase === 'done' && !canResumeClean());
  if (phase === 'running') {
    runBtn.disabled = true;
    runBtn.textContent = `⏳ 整理中… ${p.completed_files}/${p.total_files}`;
  } else {
    runBtn.disabled = Boolean(overview && overview.busy);
    runBtn.textContent = '🔁 继续上次任务';
  }

  const logBox = document.getElementById('cleanLogBox');
  const logText = overview.logs.length ? overview.logs.join('\n') : '等待操作...';
  if (logBox.innerText !== logText) {
    logBox.innerText = logText;
    logBox.scrollTop = logBox.scrollHeight;
  }

  const result = document.getElementById('cleanResult');
  if (phase !== 'done') {
    result.classList.add('hidden');
    return;
  }
  const failed = canResumeClean() || p.status === 'interrupted';
  const html = `
    <b>${failed ? '⚠️ 部分文件未处理成功' : '✅ 已完成'}</b>
    <span>${escapeHtml(failed ? '失败原因见下方日志，可点击「继续上次任务」只重试未完成的文件。' : p.current_step)}</span>
    <span>如需再次检查，点击上方「重新扫描」。</span>
    <a href="/records">查看执行记录 ↗</a>
  `;
  // 轮询时内容不变就不重建，避免打断用户选中复制
  if (result.dataset.html !== html) {
    result.dataset.html = html;
    result.innerHTML = html;
  }
  result.className = `task-result ${failed ? 'is-warn' : 'is-ok'}`;
}

/* ==========================================
   操作
   ========================================== */
function readFolderInput() {
  const folder = document.getElementById('cleanFolderInput').value.trim();
  if (!folder) appAlert('请先输入 Excel 文件夹路径！', { tone: 'warn' });
  return folder;
}

async function runPreview() {
  if (pending || phase === 'running') return;
  const folder = readFolderInput();
  if (!folder) return;
  setPending('scan');
  try {
    const data = await postAction('/api/clean/preview', { folder });
    previewData = data;
    saveLastFolder(folder);
    initSheetDecisions();
    previewFromFile = null;
    phase = 'preview';
    renderPreview();
    syncChannel?.postMessage({ type: 'rescanned', folder: previewData.folder });
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
  } finally {
    pending = null;
    renderPage();
  }
}

async function startTask() {
  if (pending || phase !== 'preview' || document.getElementById('execBtn').disabled) return;
  const folder = readFolderInput();
  if (!folder) return;
  if (overview && overview.can_resume && !(await appConfirm('开始新任务会替换上次中断任务的续跑进度，确定吗？', { okText: '开始新任务' }))) return;
  setPending('start');
  try {
    const payload = { kind: CLEAN_KIND, folder, plan: buildCleanPlan() };
    if (processedFiles.size) payload.files = Object.keys(payload.plan); // 跳过已在预览窗口处理过的文件
    await postAction('/api/quick_task', payload);
    saveLastFolder(folder);
    phase = 'running';
    minAcceptedSeq = pollSeq + 1;
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
  } finally {
    pending = null;
  }
  renderPage();
  await pollState();
}

async function resumeTask() {
  try {
    await postAction('/api/resume', {});
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
    return;
  }
  phase = 'running';
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
    if (!pending) {
      if (phase === 'running' && !isCleanRunning()) {
        phase = overview.job && overview.job.kind === CLEAN_KIND ? 'done' : 'form';
      } else if ((phase === 'form' || phase === 'done') && isCleanRunning()) {
        // 打开页面时精简任务已在后台执行，直接显示进度
        phase = 'running';
        if (overview.job.folder) document.getElementById('cleanFolderInput').value = overview.job.folder;
      }
    }
    renderPage();
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

const cleanFolderInput = document.getElementById('cleanFolderInput');
cleanFolderInput.value = readLastFolder();
cleanFolderInput.addEventListener('input', () => {
  if (phase === 'preview' || phase === 'done') {
    phase = 'form';
    previewData = null;
    renderPage();
  }
});
cleanFolderInput.addEventListener('keydown', e => {
  if (e.key === 'Enter') runPreview();
});

// 预览区宽度变化时重新排列：包括窗口缩放，以及首次渲染后出现滚动条导致的变窄
let masonryWidth = 0;
new ResizeObserver(entries => {
  const width = Math.round(entries[0].contentRect.width);
  if (width === masonryWidth) return;
  masonryWidth = width;
  layoutMasonry();
}).observe(document.getElementById('previewBox'));

setInterval(pollState, 1200);
// 扫描或提交期间每秒刷新已用时间，让用户看到仍在处理中
setInterval(() => {
  if (pending) renderPage();
}, 1000);
renderPage();
pollState();
