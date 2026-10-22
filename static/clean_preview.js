// Sheet 预览确认窗口：从精简 Sheet 页面的卡片打开，每个工作表显示第 1 页预览图，
// 单击切换「保留 / 删除 / 待定」并与精简 Sheet 页面同步，「确认执行」只处理当前这一个文件
const CLEAN_KIND = 'clean_sheets';
const SYNC_CHANNEL = 'clear_handle_excel.clean';
const SIZE_STORAGE_KEY = 'clear_handle_excel.previewSize';

const params = new URLSearchParams(location.search);
const folder = params.get('folder') || '';
const fileName = params.get('file') || '';

let fileInfo = null; // { name, sheets, backups, uncertain, note }
let defaults = {};
let states = {}; // { [sheetName]: 'keep' | 'delete' | 'uncertain' }
let renders = null; // { key, sheets: [{ name, image, pages, hidden, note }] }
let renderError = '';
let phase = 'loading'; // loading | ready | running | done
let runResult = null;
let pollTimer = null;

const channel = 'BroadcastChannel' in window ? new BroadcastChannel(SYNC_CHANNEL) : null;

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function splitDirAndBase(name) {
  const str = String(name ?? '');
  const idx = Math.max(str.lastIndexOf('/'), str.lastIndexOf('\\'));
  return idx >= 0 ? [str.slice(0, idx + 1), str.slice(idx + 1)] : ['', str];
}

function stripUncertainEmoji(name) {
  return String(name ?? '').replace(/^❓+\s*/, '');
}

const NEXT_SHEET_STATE = { keep: 'delete', delete: 'uncertain', uncertain: 'keep' };
const STATE_LABEL = { keep: '保留', delete: '删除', uncertain: '待定' };

function computeDefaults(file) {
  const delSet = new Set(file.backups || []);
  const uncSet = new Set(file.uncertain || []);
  const map = {};
  (file.sheets || []).forEach(name => {
    if (delSet.has(name)) map[name] = 'delete';
    else if (uncSet.has(name)) map[name] = 'uncertain';
    else map[name] = 'keep';
  });
  return map;
}

function broadcast(message) {
  if (channel) channel.postMessage({ folder, file: fileName, ...message });
}

function locked() {
  return phase !== 'ready';
}

function cycleState(sheetIndex) {
  if (locked()) return;
  const name = fileInfo.sheets[sheetIndex];
  if (name === undefined) return;
  states[name] = NEXT_SHEET_STATE[states[name]] || 'keep';
  broadcast({ type: 'states', states });
  fileInfo.sheets.forEach((_, idx) => updateCard(idx)); // 旧表状态会影响新表是否去掉「新」字
  renderActionBar();
}

function resetStates() {
  if (locked()) return;
  states = { ...defaults };
  broadcast({ type: 'states', states });
  fileInfo.sheets.forEach((_, idx) => updateCard(idx));
  renderActionBar();
}

function counts() {
  const result = { keep: 0, delete: 0, uncertain: 0 };
  fileInfo.sheets.forEach(name => {
    result[states[name] || 'keep'] += 1;
  });
  return result;
}

/* ==========================================
   页面渲染：卡片尺寸固定，预览图加载前后、切换状态时卡片都不移动
   ========================================== */
// 「旧」表全部标为删除、「新」表保留时，执行后「新」表去掉开头的「新」字（与后端 _new_prefix_renames 一致）；
// 去掉后与剩下的表重名时不改名
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

function renderInfoForSheet(name) {
  if (!renders) return null;
  return renders.sheets.find(s => s.name === name) || null;
}

function cardNameHtml(name) {
  const display = stripUncertainEmoji(name);
  const renamed = findNewPrefixRenames(fileInfo.sheets, n => states[n] || 'keep', fileInfo.new_old_pairs);
  if (!renamed.has(name)) return escapeHtml(display);
  return `<s class="strip-new-prefix" title="旧版删除后去掉「新」字">新</s>${escapeHtml(display.slice(1))}`;
}

function cardBodyHtml(name, sheetIndex) {
  const info = renderInfoForSheet(name);
  let preview;
  if (info && info.image) {
    const src = `/api/clean/thumb/${encodeURIComponent(renders.key)}/${encodeURIComponent(info.image)}`;
    const pdf = info.pdf
      ? `/api/clean/thumb/${encodeURIComponent(renders.key)}/${encodeURIComponent(info.pdf)}`
      : src;
    preview = `<img src="${src}" alt="" loading="lazy" draggable="false">
      <a class="pv-zoom" href="${pdf}" target="_blank" title="在新标签页打开导出的 PDF（与「导出 PDF」功能的结果一致，含全部页）" onclick="event.stopPropagation()">🔍 PDF</a>`;
  } else if (info) {
    preview = `<span class="pv-thumb-note">${escapeHtml(info.note || '无法生成预览')}</span>`;
  } else if (renderError) {
    preview = `<span class="pv-thumb-note">预览图生成失败：${escapeHtml(renderError)}</span>`;
  } else {
    preview = '<span class="pv-thumb-note"><span class="clean-spinner" aria-hidden="true"></span>正在生成预览…</span>';
  }
  let similar = '';
  if (info && info.similar) {
    const other = escapeHtml(stripUncertainEmoji(info.similar.match));
    similar = info.similar.identical
      ? `🟰 与「${other}」内容完全相同`
      : `≈ 与「${other}」约 ${Math.round(info.similar.ratio * 100)}% 相同，可能是新旧版本`;
  }
  const st = states[name] || 'keep';
  const meta = [];
  if (info && info.hidden) meta.push('隐藏');
  if (info && info.pages > 1) meta.push(`共 ${info.pages} 页`);
  return `
    <span class="pv-thumb">${preview}</span>
    <span class="pv-card-foot">
      <span class="pv-card-name" title="${escapeHtml(name)}">${cardNameHtml(name)}</span>
      <span class="pv-card-meta">${escapeHtml(meta.join(' · '))}</span>
      <span class="preview-sheet-tag is-${st}">${STATE_LABEL[st]}</span>
    </span>
    <span class="pv-card-similar${info && info.similar && info.similar.identical ? ' is-identical' : ''}" title="${similar}">${similar}</span>`;
}

function renderGrid() {
  const grid = document.getElementById('pvGrid');
  grid.innerHTML = fileInfo.sheets.map((name, idx) => `
    <div class="pv-card is-${states[name] || 'keep'}" id="pvCard${idx}" role="button" tabindex="0"
      onclick="cycleState(${idx})" onkeydown="if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); cycleState(${idx}); }">
      ${cardBodyHtml(name, idx)}
    </div>`).join('');
  grid.classList.toggle('is-locked', locked());
  grid.classList.remove('hidden');
}

function updateCard(idx) {
  const card = document.getElementById(`pvCard${idx}`);
  if (!card) return;
  const name = fileInfo.sheets[idx];
  card.className = `pv-card is-${states[name] || 'keep'}`;
  card.querySelector('.pv-card-name').innerHTML = cardNameHtml(name);
  card.querySelector('.pv-card-foot .preview-sheet-tag').outerHTML =
    `<span class="preview-sheet-tag is-${states[name] || 'keep'}">${STATE_LABEL[states[name] || 'keep']}</span>`;
}

function renderActionBar() {
  const bar = document.getElementById('pvActionBar');
  bar.classList.toggle('hidden', !fileInfo || !fileInfo.sheets.length);
  if (!fileInfo) return;
  const c = counts();
  const execBtn = document.getElementById('pvExecBtn');
  const hint = document.getElementById('pvHint');
  document.getElementById('pvResetBtn').disabled = locked();
  document.getElementById('pvGrid').classList.toggle('is-locked', locked());

  document.getElementById('pvSummary').innerHTML = `
    共 <b>${fileInfo.sheets.length}</b> 个工作表：
    删除 <b class="is-deleted">${c.delete}</b> 个 ·
    待定 <b class="is-uncertain">${c.uncertain}</b> 个 ·
    保留 <b>${c.keep}</b> 个`;

  hint.textContent = '';
  if (phase === 'running') {
    execBtn.disabled = true;
    execBtn.textContent = '⏳ 整理中…';
    return;
  }
  if (phase === 'done') {
    execBtn.disabled = true;
    execBtn.textContent = runResult && runResult.ok ? '✅ 已处理完成' : '⚠️ 处理失败';
    hint.innerHTML = runResult && runResult.ok
      ? '这个文件已处理完成，可以关闭此窗口'
      : `${escapeHtml(runResult ? runResult.message : '')} <a href="/records" target="_blank">查看执行记录 ↗</a>`;
    return;
  }
  const allDeleted = c.delete >= fileInfo.sheets.length;
  const hasAction = c.delete > 0; // 待定只是本次不删，不改文件
  execBtn.disabled = phase !== 'ready' || allDeleted || !hasAction;
  if (allDeleted) {
    execBtn.textContent = '⚠️ 请至少保留 1 个工作表';
    hint.textContent = '不能将所有工作表全部标为删除';
  } else if (!hasAction) {
    execBtn.textContent = '没有标记删除的工作表';
  } else {
    execBtn.textContent = `🧹 确认执行（删除 ${c.delete} 个工作表）`;
  }
}

function showMessage(text, { spinning = false, error = false } = {}) {
  const box = document.getElementById('pvMessage');
  box.classList.toggle('hidden', !text);
  box.classList.toggle('is-scanning', spinning);
  box.classList.toggle('is-error', error);
  box.querySelector('.clean-spinner')?.classList.toggle('hidden', !spinning);
  document.getElementById('pvMessageText').textContent = text;
}

function setCardSize(size) {
  const grid = document.getElementById('pvGrid');
  grid.classList.remove('pv-size-s', 'pv-size-m', 'pv-size-l');
  grid.classList.add(`pv-size-${size}`);
  document.querySelectorAll('.pv-size-switch .btn').forEach(btn => {
    btn.classList.toggle('is-active', btn.dataset.size === size);
  });
  try {
    localStorage.setItem(SIZE_STORAGE_KEY, size);
  } catch (e) {}
}

/* ==========================================
   与精简 Sheet 页面同步：任一边切换状态都广播整份标记
   ========================================== */
if (channel) {
  channel.onmessage = event => {
    const msg = event.data || {};
    if (msg.folder !== folder || (msg.file !== undefined && msg.file !== fileName) || !fileInfo) return;
    if (msg.type === 'states' && msg.states && phase === 'ready') {
      fileInfo.sheets.forEach(name => {
        if (['keep', 'delete', 'uncertain'].includes(msg.states[name])) states[name] = msg.states[name];
      });
      fileInfo.sheets.forEach((_, idx) => updateCard(idx));
      renderActionBar();
    } else if (msg.type === 'rescanned') {
      // 精简 Sheet 页面重新扫描或恢复默认后，按那边的最新标记重新同步
      broadcast({ type: 'hello' });
    }
  };
}

/* ==========================================
   执行：只处理当前文件
   ========================================== */
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

async function fetchOverview() {
  const res = await fetch('/api/overview');
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || '读取状态失败');
  return data;
}

async function executeFile() {
  if (phase !== 'ready' || document.getElementById('pvExecBtn').disabled) return;
  let overview;
  try {
    overview = await fetchOverview();
  } catch (e) {
    alert(e.message);
    return;
  }
  if (overview.busy) {
    alert(`后台正在执行「${overview.progress.task_name}」，完成后才能开始`);
    return;
  }
  if (overview.can_resume && !confirm('开始新任务会替换上次中断任务的续跑进度，确定吗？')) return;

  const del = fileInfo.sheets.filter(name => states[name] === 'delete');
  const unc = fileInfo.sheets.filter(name => states[name] === 'uncertain');
  phase = 'running';
  renderActionBar();
  try {
    await postAction('/api/quick_task', {
      kind: CLEAN_KIND,
      folder,
      plan: { [fileName]: { delete: del, uncertain: unc } },
      files: [fileName],
    });
  } catch (e) {
    phase = 'ready';
    renderActionBar();
    alert(e.message);
    return;
  }
  pollTimer = setInterval(pollRun, 1000);
}

async function pollRun() {
  let overview;
  try {
    overview = await fetchOverview();
  } catch (e) {
    return;
  }
  document.getElementById('statusText').innerText = overview.busy ? '⏳ 后台处理中...' : '✅ 就绪';
  if (overview.busy) return;
  clearInterval(pollTimer);
  const p = overview.progress;
  const ok = !p.error_files && overview.job && overview.job.kind === CLEAN_KIND && !overview.can_resume;
  runResult = { ok, message: ok ? p.current_step : '处理失败，原因见执行记录' };
  phase = 'done';
  if (ok) broadcast({ type: 'done' });
  renderActionBar();
}

/* ==========================================
   初始化：先列出工作表（可以马上开始标记），再用 Excel 生成预览图
   ========================================== */
async function init() {
  const [dir, base] = splitDirAndBase(fileName);
  document.title = `预览 - ${base}`;
  document.getElementById('pvFileName').textContent = `🖼️ ${base || 'Sheet 预览确认'}`;
  document.getElementById('pvFolder').textContent = `${folder}${dir ? ` › ${dir.replace(/[\\/]$/, '')}` : ''}`;
  let size = 'm';
  try {
    size = localStorage.getItem(SIZE_STORAGE_KEY) || 'm';
  } catch (e) {}
  setCardSize(['s', 'm', 'l'].includes(size) ? size : 'm');

  if (!folder || !fileName) {
    showMessage('缺少文件信息，请从精简 Sheet 页面的卡片上点击「🖼 预览」打开', { error: true });
    return;
  }
  try {
    const data = await postAction('/api/clean/file', { folder, file: fileName });
    fileInfo = data.file;
  } catch (e) {
    showMessage(e.message, { error: true });
    return;
  }
  if (!fileInfo.sheets.length) {
    showMessage(fileInfo.note || '没有读取到工作表', { error: true });
    return;
  }
  defaults = computeDefaults(fileInfo);
  states = { ...defaults };
  phase = 'ready';
  showMessage('');
  renderGrid();
  renderActionBar();
  broadcast({ type: 'hello' }); // 精简 Sheet 页面若已打开，会回传当前标记

  // 进度只显示在顶栏状态里：页面上方不插入提示块，预览图生成完毕时卡片不会移动
  const status = document.getElementById('statusText');
  const started = Date.now();
  const tick = () => {
    const secs = Math.floor((Date.now() - started) / 1000);
    status.innerText = `⏳ 正在用 Excel 生成预览图…${secs >= 1 ? ` ${secs}s` : ''}`;
  };
  tick();
  const timer = setInterval(tick, 1000);
  try {
    renders = await postAction('/api/clean/render', { folder, file: fileName });
    status.innerText = renders.cached ? '✅ 预览图已生成（文件未改动，使用上次的预览图）' : '✅ 预览图已生成';
  } catch (e) {
    renderError = e.message;
    status.innerText = '⚠️ 预览图生成失败';
    status.title = e.message;
  } finally {
    clearInterval(timer);
  }
  fileInfo.sheets.forEach((name, idx) => {
    const card = document.getElementById(`pvCard${idx}`);
    if (card) card.innerHTML = cardBodyHtml(name, idx);
  });
}

init();
