'use strict';
const REVIEW_COLORS = ['blue', 'green', 'yellow'];
const REVIEW_LABELS = { blue: '蓝 · 已修复', green: '绿 · 还原并忽略', yellow: '黄 · 人工处理' };
let batchState = { cards: [], busy: false, job: null, progress: {}, logs: [] };
let selectedFilter = 'all';
let largeMode = false;
let activeCardId = null;
let requestBusy = false;
let pollVersion = 0;
let renderKey = '';
let lastWheelAt = 0;
const $ = id => document.getElementById(id);
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const nextColor = (color, direction = 1) => REVIEW_COLORS[(REVIEW_COLORS.indexOf(color) + direction + 3) % 3];
const shownCards = () => batchState.cards.filter(c => selectedFilter === 'all' || c.color === selectedFilter);
const manualCards = () => batchState.cards.filter(c => c.color === 'yellow' && c.finalized);

async function post(url, payload = {}) {
  const response = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
  const result = await response.json();
  if (!response.ok || !result.ok) throw new Error(result.error || `HTTP ${response.status}`);
  return result;
}

async function poll() {
  const version = ++pollVersion;
  try {
    const response = await fetch('/api/batch', { cache: 'no-store' });
    if (!response.ok) throw new Error(`无法读取进度：HTTP ${response.status}`);
    const value = await response.json();
    if (version !== pollVersion || requestBusy) return;
    batchState = value;
    if (!document.activeElement || document.activeElement.id !== 'folderInput') {
      if (value.folder && !$('folderInput').value) $('folderInput').value = value.folder;
      if (value.busy && value.folder) $('folderInput').value = value.folder;
    }
    render();
  } catch (error) { $('errorText').textContent = error.message; }
}

async function action(work) {
  if (requestBusy) return;
  requestBusy = true;
  ++pollVersion;
  $('errorText').textContent = '';
  renderControls();
  try { await work(); } catch (error) { $('errorText').textContent = error.message; }
  finally { requestBusy = false; await poll(); }
}

function renderControls() {
  const busy = batchState.busy || requestBusy;
  for (const id of ['startBtn', 'commitBtn', 'folderInput']) $(id).disabled = busy;
  $('commitBtn').disabled = busy || !batchState.cards.length;
  $('largeBtn').disabled = !batchState.cards.length;
  $('openAllBtn').disabled = busy || !manualCards().length;
  $('recheckAllBtn').disabled = busy || !manualCards().length;
  $('pauseBtn').classList.toggle('hidden', !batchState.busy || batchState.job?.kind !== 'batch_scan');
  $('pauseBtn').disabled = requestBusy || !!batchState.job?.pause_requested;
  $('resumeBtn').classList.toggle('hidden', batchState.busy || batchState.job?.status !== 'interrupted');
  $('resumeBtn').disabled = requestBusy;
  $('statusText').textContent = batchState.busy ? '⏳ 后台处理中' : batchState.job?.status === 'interrupted' ? '⏸ 可继续任务' : '就绪';
}

function cardHtml(card, large = false) {
  const link = excelLinkHtml(esc(card.wb_name), { file: card.wb_name, folder: batchState.folder, sheet: card.sheet_name, cell: card.cell_addr });
  const labels = { before: '修改前', after: card.after_kind === 'attempt' ? '最后一次尝试（未采用）' : '修改后' };
  const images = ['before', 'after'].map(which => `<figure><figcaption>${labels[which]}</figcaption>${card.images?.[which]
    ? `<img ${large ? '' : 'loading="lazy"'} src="${esc(card.images[which])}" alt="${labels[which]}截图" data-preview="1" onerror="this.outerHTML='<span class=\'batch-missing\'>截图暂不可用</span>'">`
    : '<span class="batch-missing">暂无截图</span>'}</figure>`).join('');
  return `<article class="batch-card" tabindex="0" role="button" aria-label="${esc(`${card.wb_name} ${card.sheet_name}!${card.cell_addr} ${REVIEW_LABELS[card.color]}，单击切换状态`)}" data-id="${esc(card.id)}" data-color="${card.color}">
    <div class="batch-card-head"><div>${link}<p><b>${esc(card.sheet_name)}!${esc(card.cell_addr || '定位待确认')}</b></p></div><span class="batch-color-label">${REVIEW_LABELS[card.color]}</span></div>
    <div class="batch-shots">${images}</div><p>${esc(card.text)}</p><p>${esc(card.message)}</p>
    <p>${card.finalized ? '已提交 · 可人工处理' : '颜色标记待提交'}${batchState.busy ? ' · 后台任务结束后可标记' : ''}</p>
    <div class="batch-card-footer"><button class="btn btn-outline" data-action="large">▣ 大图审核</button>${card.color === 'yellow' && card.finalized ? '<button class="btn btn-outline" data-action="recheck">🔁 单项复检</button>' : ''}</div></article>`;
}

function render() {
  renderControls();
  const p = batchState.progress || {};
  const job = batchState.job;
  const total = job?.workbooks?.length || p.total || 0;
  const completed = job?.kind === 'batch_scan' ? job.completed.length : p.completed || 0;
  $('progressText').textContent = `${batchState.busy ? '正在处理' : job?.status === 'interrupted' ? '已暂停或部分文件未完成' : '当前进度'}：${completed} / ${total} 个文件 · 待审核 ${batchState.cards.length} 个单元格`;
  $('progressFill').style.width = `${total ? Math.round(completed * 100 / total) : 0}%`;
  $('logBox').textContent = batchState.logs.join('\n') || '等待操作…';
  const counts = Object.fromEntries(REVIEW_COLORS.map(c => [c, batchState.cards.filter(card => card.color === c).length]));
  const filterHtml = ['all', ...REVIEW_COLORS].map(color => `<button class="btn btn-outline" data-filter="${color}" aria-pressed="${selectedFilter === color}">${color === 'all' ? `全部 ${batchState.cards.length}` : `${REVIEW_LABELS[color]} ${counts[color]}`}</button>`).join('');
  if ($('filters').innerHTML !== filterHtml) $('filters').innerHTML = filterHtml;
  const key = JSON.stringify([batchState.cards, selectedFilter, largeMode, activeCardId, batchState.busy]);
  if (renderKey === key) return;
  renderKey = key;
  $('cardList').classList.toggle('hidden', largeMode);
  $('largeView').classList.toggle('hidden', !largeMode);
  const shown = shownCards();
  if (largeMode) {
    if (!shown.some(c => c.id === activeCardId)) activeCardId = shown[0]?.id;
    const index = shown.findIndex(c => c.id === activeCardId);
    $('largeIndex').textContent = shown.length ? `${index + 1} / ${shown.length}` : '没有当前颜色的卡片';
    $('largeCard').innerHTML = index >= 0 ? cardHtml(shown[index], true) : '<p class="batch-empty">当前筛选下没有卡片</p>';
    $('prevBtn').disabled = index <= 0;
    $('nextBtn').disabled = index >= shown.length - 1;
  } else {
    const columnCount = Math.max(1, Math.min(3, Math.floor((($('cardList').clientWidth || 1200) + 18) / 328)));
    $('cardList').innerHTML = shown.length ? ['yellow', 'green', 'blue'].map(color => {
      const cards = shown.filter(c => c.color === color);
      const columns = Array.from({ length: columnCount }, () => []);
      cards.forEach((card, index) => columns[index % columnCount].push(card));
      return cards.length ? `<section><h2 class="batch-group-title">${REVIEW_LABELS[color]} · ${cards.length}</h2><div class="batch-waterfall">${columns.map(column => `<div class="batch-column">${column.map(c => cardHtml(c)).join('')}</div>`).join('')}</div></section>` : '';
    }).join('') : `<section class="panel batch-empty"><h2>${batchState.busy ? '正在自动扫描与修复…' : '当前没有待审核卡片'}</h2><p>${batchState.busy ? '文件处理后会在这里出现修改前后对比，无需逐项放行。' : '开始全自动处理，或切换颜色筛选查看结果。'}</p></section>`;
  }
}

function openLarge(id) {
  largeMode = true;
  activeCardId = id || shownCards()[0]?.id;
  render();
  $('largeView').scrollIntoView({ block: 'start', behavior: 'smooth' });
}

function moveCard(delta) {
  const cards = shownCards();
  const index = cards.findIndex(c => c.id === activeCardId);
  activeCardId = cards[Math.max(0, Math.min(cards.length - 1, index + delta))]?.id;
  render();
}

async function changeColor(id, direction = 1) {
  if (batchState.busy || requestBusy) return;
  const card = batchState.cards.find(c => c.id === id);
  if (!card) return;
  let color = nextColor(card.color, direction);
  // 自动失败的黄色项也能用鼠标标绿；复检通过前跳过不可用的蓝色状态。
  if (color === 'blue' && !card.verified) color = nextColor(color, direction);
  await action(() => post('/api/batch/mark', { id, color }));
}

async function startJob(url, body = {}) {
  if (batchState.job?.status === 'interrupted' && url !== '/api/resume'
      && !window.confirm('这会替换上次任务的续跑点。已保存的修复和审核标记会保留，继续吗？')) return;
  await action(() => post(url, body));
}

$('startBtn').addEventListener('click', () => {
  const folder = $('folderInput').value.trim();
  if (!folder) { $('errorText').textContent = '请输入 Excel 文件夹路径'; return; }
  startJob('/api/batch/start', { folder }).then(() => {
    if (batchState.busy) { try { localStorage.setItem('clear_handle_excel.lastFolder', folder); } catch (_) {} }
  });
});
$('resumeBtn').addEventListener('click', () => startJob('/api/resume'));
$('pauseBtn').addEventListener('click', () => action(() => post('/api/batch/pause')));
$('commitBtn').addEventListener('click', () => startJob('/api/batch/commit', { card_ids: batchState.cards.map(c => c.id) }));
$('openAllBtn').addEventListener('click', () => startJob('/api/batch/open', { card_ids: manualCards().map(c => c.id) }));
$('recheckAllBtn').addEventListener('click', () => startJob('/api/batch/recheck', { card_ids: manualCards().map(c => c.id) }));
$('largeBtn').addEventListener('click', () => openLarge());
$('closeLargeBtn').addEventListener('click', () => { largeMode = false; render(); });
$('prevBtn').addEventListener('click', () => moveCard(-1));
$('nextBtn').addEventListener('click', () => moveCard(1));
$('filters').addEventListener('click', event => {
  const button = event.target.closest('[data-filter]');
  if (button) { selectedFilter = button.dataset.filter; render(); }
});

for (const id of ['cardList', 'largeCard']) $(id).addEventListener('click', event => {
  const card = event.target.closest('[data-id]');
  if (!card || event.target.closest('a')) return;
  const button = event.target.closest('[data-action]');
  if (button?.dataset.action === 'large') openLarge(card.dataset.id);
  else if (button?.dataset.action === 'recheck') startJob('/api/batch/recheck', { card_ids: [card.dataset.id] });
  else changeColor(card.dataset.id);
});

document.addEventListener('keydown', event => {
  if (event.target.closest('input, textarea, select') || event.ctrlKey || event.metaKey || event.altKey) return;
  if (largeMode) {
    if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Escape'].includes(event.key)) event.preventDefault();
    if (event.key === 'ArrowUp') moveCard(-1);
    if (event.key === 'ArrowDown') moveCard(1);
    if (event.key === 'ArrowLeft') changeColor(activeCardId, -1);
    if (event.key === 'ArrowRight') changeColor(activeCardId, 1);
    if (event.key === 'Escape') { largeMode = false; render(); }
  } else if (['Enter', ' '].includes(event.key) && event.target.matches('.batch-card')) {
    event.preventDefault(); changeColor(event.target.dataset.id);
  }
});
$('largeView').addEventListener('wheel', event => {
  if (!largeMode || Math.abs(event.deltaY) < 8 || event.ctrlKey) return;
  event.preventDefault();
  if (Date.now() - lastWheelAt < 280) return;
  lastWheelAt = Date.now();
  moveCard(event.deltaY > 0 ? 1 : -1);
}, { passive: false });

document.addEventListener('pointermove', event => {
  const image = event.target.closest('#cardList img[data-preview]');
  const preview = $('imagePreview');
  preview.classList.toggle('hidden', !image || largeMode);
  if (!image || largeMode) return;
  preview.querySelector('img').src = image.src;
  preview.style.left = `${Math.max(8, Math.min(event.clientX + 20, innerWidth - preview.offsetWidth - 12))}px`;
  preview.style.top = `${Math.max(8, Math.min(event.clientY + 20, innerHeight - preview.offsetHeight - 12))}px`;
});
document.addEventListener('scroll', () => $('imagePreview').classList.add('hidden'), true);
window.addEventListener('resize', () => { renderKey = ''; render(); });
try { $('folderInput').value = localStorage.getItem('clear_handle_excel.lastFolder') || ''; } catch (_) {}
poll();
setInterval(poll, 1800);
