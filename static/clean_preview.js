// Sheet 预览确认窗口：从精简 Sheet 页面的卡片打开，每个工作表显示第 1 页预览图，
// 单击切换「保留 / 删除 / 待定」、单击表名改名，并与精简 Sheet 页面同步；标记完成后返回精简 Sheet 页面统一执行
const CLEAN_KIND = 'clean_sheets';
const SYNC_CHANNEL = 'clear_handle_excel.clean';
const SIZE_STORAGE_KEY = 'clear_handle_excel.previewSize';

const params = new URLSearchParams(location.search);
const folder = params.get('folder') || '';
const fileName = params.get('file') || '';

let fileInfo = null; // { name, sheets, backups, uncertain, note }
let defaults = {};
let states = {}; // { [sheetName]: 'keep' | 'delete' | 'uncertain' }
let renames = {}; // { [原表名]: 新表名 }：执行时和删除一起改名
let renders = null; // { key, sheets: [{ name, image, pages, hidden, note }] }
let renderError = '';
let phase = 'loading'; // loading | ready

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

function broadcastStates() {
  broadcast({ type: 'states', states, renames });
}

function locked() {
  return phase !== 'ready';
}

function cycleState(sheetIndex) {
  if (locked()) return;
  // 正在改名时点了别的卡片：这一下只用来结束输入（blur 先确认改名），不切换那张卡片
  if (editingIndex >= 0 || performance.now() - renameBlurAt < 400) return;
  delete document.getElementById('pvHint').dataset.sticky;
  const name = fileInfo.sheets[sheetIndex];
  if (name === undefined) return;
  states[name] = NEXT_SHEET_STATE[states[name]] || 'keep';
  broadcastStates();
  fileInfo.sheets.forEach((_, idx) => updateCard(idx)); // 旧表状态会影响新表是否去掉「新」字
  renderActionBar();
}

function resetStates() {
  if (locked()) return;
  if (editingIndex >= 0) finishRename(false);
  delete document.getElementById('pvHint').dataset.sticky;
  states = { ...defaults };
  renames = {};
  broadcastStates();
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
// 「旧」表全部标为删除（或本来就没有「旧」表，olds 为空）、「新」表保留时，执行后「新」表去掉开头的「新」字
// （与后端 _new_prefix_renames 一致）；去掉后与剩下的表重名时不改名；手动改过名的表不再自动去「新」
function findNewPrefixRenames(sheets, stateOf, pairs, custom = {}) {
  const renamed = new Set();
  const remaining = sheets.filter(n => stateOf(n) !== 'delete').flatMap(n => (custom[n] ? [n, custom[n]] : [n]));
  Object.entries(pairs || {}).forEach(([newName, olds]) => {
    if (custom[newName] || stateOf(newName) !== 'keep' || !olds.every(o => stateOf(o) === 'delete')) return;
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
  let html;
  if (renames[name]) {
    html = `<span class="sheet-renamed-mark" aria-hidden="true">✏️</span>${escapeHtml(renames[name])}`;
  } else {
    const renamed = findNewPrefixRenames(fileInfo.sheets, n => states[n] || 'keep', fileInfo.new_old_pairs, renames);
    html = renamed.has(name)
      ? `<s class="strip-new-prefix">新</s>${escapeHtml(display.slice(1))}`
      : escapeHtml(display);
  }
  return `<span class="pv-name-text" onclick="startRename(event, ${fileInfo.sheets.indexOf(name)})">${html}</span>`;
}

function cardNameTitle(name) {
  if (renames[name]) {
    return `原名「${name}」，确认执行时改名为「${renames[name]}」（单击表名修改；清空或改回原名就是撤销）`;
  }
  return `${name}（单击表名可以改名，确认执行时一起改）`;
}

/* ==========================================
   单击表名就地改名：只换掉名字那一格，卡片尺寸不变；回车/点别处确认，Esc 取消
   ========================================== */
const SHEET_NAME_MAX = 31;
const SHEET_NAME_BAD_CHARS = [':', '\\', '/', '?', '*', '[', ']'];

// 与后端 sheet_name_error 一致；另外检查和本文件执行后剩下的表重名（不区分大小写）
function sheetNameError(name, sheetName) {
  if (!name.trim()) return '名字不能为空';
  if (name.length > SHEET_NAME_MAX) return `最多 ${SHEET_NAME_MAX} 个字（现在 ${name.length} 个）`;
  const bad = SHEET_NAME_BAD_CHARS.filter(c => name.includes(c));
  if (bad.length) return `不能包含 ${bad.join(' ')}`;
  if (name.startsWith("'") || name.endsWith("'")) return "开头和结尾不能是 '";
  if (name.toLowerCase() === 'history') return 'History 是 Excel 保留的名字';
  const lower = name.toLowerCase();
  const clash = fileInfo.sheets.find(other => other !== sheetName && (states[other] || 'keep') !== 'delete' &&
    (other.toLowerCase() === lower || (renames[other] || '').toLowerCase() === lower));
  if (clash) return `和工作表「${renames[clash] && renames[clash].toLowerCase() === lower ? renames[clash] : clash}」重名`;
  return '';
}

let editingIndex = -1;
let renameBlurAt = -Infinity;

function startRename(event, sheetIndex) {
  event.stopPropagation(); // 单击表名只改名，不切换保留/删除/待定
  if (locked() || editingIndex >= 0) return;
  const name = fileInfo.sheets[sheetIndex];
  const box = document.querySelector(`#pvCard${sheetIndex} .pv-card-name`);
  if (name === undefined || !box) return;
  editingIndex = sheetIndex;
  box.classList.add('is-editing');
  box.innerHTML = `<input class="pv-name-input" type="text" maxlength="${SHEET_NAME_MAX}" spellcheck="false"
    value="${escapeHtml(renames[name] || stripUncertainEmoji(name))}" aria-label="新的工作表名">`;
  const input = box.querySelector('input');
  input.addEventListener('click', e => e.stopPropagation());
  input.addEventListener('keydown', e => {
    e.stopPropagation(); // 空格、回车不能冒泡到卡片去切换状态
    if (e.key === 'Enter') { e.preventDefault(); finishRename(true); }
    else if (e.key === 'Escape') { e.preventDefault(); finishRename(false); }
  });
  input.addEventListener('input', () => {
    const error = sheetNameError(input.value.trim(), name);
    input.classList.toggle('is-invalid', Boolean(error) && input.value.trim() !== '');
    setRenameHint(error && input.value.trim() !== '' ? `⚠️ ${error}` : '回车确认，Esc 取消；清空或改回原名就是撤销改名');
  });
  input.addEventListener('blur', () => {
    if (editingIndex === sheetIndex) renameBlurAt = performance.now();
    finishRename(true);
  });
  input.focus({ preventScroll: true });
  input.select();
  setRenameHint('回车确认，Esc 取消；清空或改回原名就是撤销改名');
}

function finishRename(commit) {
  if (editingIndex < 0) return;
  const idx = editingIndex;
  const name = fileInfo.sheets[idx];
  const input = document.querySelector(`#pvCard${idx} .pv-name-input`);
  const value = input ? input.value.trim() : '';
  if (commit) {
    if (!value || value === stripUncertainEmoji(name) || value === name) {
      delete renames[name];
    } else {
      const error = sheetNameError(value, name);
      if (error) {
        // 点别处时不合规的名字直接放弃，回车时留在输入框里让用户改
        if (document.activeElement === input) {
          input.classList.add('is-invalid');
          setRenameHint(`⚠️ ${error}`);
          return;
        }
        setRenameHint(`⚠️ 没有改名：${error}`, true);
        commit = false;
      } else {
        renames[name] = value;
      }
    }
  }
  editingIndex = -1;
  fileInfo.sheets.forEach((_, i) => updateCard(i)); // 改名会影响别的表能否去掉「新」字
  if (commit) {
    broadcastStates();
    setRenameHint('');
  } else if (!document.getElementById('pvHint').dataset.sticky) {
    setRenameHint('');
  }
  renderActionBar();
}

function setRenameHint(text, sticky = false) {
  const hint = document.getElementById('pvHint');
  hint.textContent = text;
  if (sticky) hint.dataset.sticky = '1';
  else delete hint.dataset.sticky;
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
    if (info.similar.identical) {
      similar = `🟰 与「${other}」内容完全相同`;
    } else if (info.similar.same_cells) {
      // 复制工作表后只换了截图、挪了红框：文字一样，图片/图形不一样
      similar = `≈ 与「${other}」文字相同，但${escapeHtml((info.similar.diff || ['图片']).join('、'))}不同，可能是新旧版本`;
    } else {
      similar = `≈ 与「${other}」约 ${Math.round(info.similar.ratio * 100)}% 相同，可能是新旧版本`;
    }
  }
  const st = states[name] || 'keep';
  const meta = [];
  if (info && info.hidden) meta.push('隐藏');
  if (info && info.pages > 1) meta.push(`共 ${info.pages} 页`);
  const excel = excelLinkHtml('📗 Excel', { file: fileName, folder, sheet: name }, 'pv-excel');
  return `
    <span class="pv-thumb">${preview}${excel}</span>
    <span class="pv-card-foot">
      <span class="pv-card-name" title="${escapeHtml(cardNameTitle(name))}">${cardNameHtml(name)}</span>
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
  if (idx === editingIndex) return; // 正在输入新名字，不打断
  const nameBox = card.querySelector('.pv-card-name');
  nameBox.classList.remove('is-editing');
  nameBox.innerHTML = cardNameHtml(name);
  nameBox.title = cardNameTitle(name);
  card.querySelector('.pv-card-foot .preview-sheet-tag').outerHTML =
    `<span class="preview-sheet-tag is-${states[name] || 'keep'}">${STATE_LABEL[states[name] || 'keep']}</span>`;
}

function renderActionBar() {
  const bar = document.getElementById('pvActionBar');
  bar.classList.toggle('hidden', !fileInfo || !fileInfo.sheets.length);
  if (!fileInfo) return;
  const c = counts();
  const hint = document.getElementById('pvHint');
  document.getElementById('pvResetBtn').disabled = locked();
  document.getElementById('pvGrid').classList.toggle('is-locked', locked());

  const renameCount = Object.keys(renames).filter(n => (states[n] || 'keep') !== 'delete').length;
  document.getElementById('pvSummary').innerHTML = `
    共 <b>${fileInfo.sheets.length}</b> 个工作表：
    删除 <b class="is-deleted">${c.delete}</b> 个 ·
    待定 <b class="is-uncertain">${c.uncertain}</b> 个 ·
    保留 <b>${c.keep}</b> 个` + (renameCount ? ` · 改名 <b>${renameCount}</b> 个` : '');

  if (editingIndex >= 0 || hint.dataset.sticky) return; // 改名的提示还在显示
  const allDeleted = c.delete >= fileInfo.sheets.length;
  // 全部标为删除时也可以返回，精简 Sheet 页面同样会提示并禁止执行
  hint.textContent = allDeleted ? '⚠️ 不能将所有工作表全部标为删除，请至少保留 1 个' : '';
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
      if (msg.renames && typeof msg.renames === 'object') {
        renames = {};
        fileInfo.sheets.forEach(name => {
          if (typeof msg.renames[name] === 'string' && msg.renames[name]) renames[name] = msg.renames[name];
        });
      }
      fileInfo.sheets.forEach((_, idx) => updateCard(idx));
      renderActionBar();
    } else if (msg.type === 'rescanned') {
      // 精简 Sheet 页面重新扫描或恢复默认后，按那边的最新标记重新同步
      broadcast({ type: 'hello' });
    }
  };
}

/* ==========================================
   标记完成后返回精简 Sheet 页面：这里不执行，最后在精简 Sheet 页面统一「确认执行」
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

function returnToClean() {
  if (editingIndex >= 0) finishRename(true);
  if (fileInfo && phase === 'ready') broadcastStates();
  broadcast({ type: 'return' }); // 精简 Sheet 页面滚到来源卡片并让边框发亮
  try {
    if (window.opener && !window.opener.closed) window.opener.focus();
  } catch (e) {}
  window.close();
  // 不是从精简 Sheet 页面打开的窗口关不掉，直接跳过去
  setTimeout(() => {
    location.href = '/clean';
  }, 300);
}

/* ==========================================
   初始化：先列出工作表（可以马上开始标记），再用 Excel 生成预览图
   ========================================== */
async function init() {
  const [dir, base] = splitDirAndBase(fileName);
  document.title = `预览 - ${base}`;
  document.getElementById('pvFileName').innerHTML = base
    ? `🖼️ ${excelLinkHtml(escapeHtml(base), { file: fileName, folder })}`
    : '🖼️ Sheet 预览确认';
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
