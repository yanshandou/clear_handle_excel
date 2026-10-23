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
let layoutKey = '';
let cardHtmlCache = new Map();
let lastActiveIndex = 0;
let lastWheelAt = 0;
let selectedIds = new Set();
let selectNotice = '';
// 筛选后改了颜色的卡片继续留在当前筛选里，不立刻消失；重新点筛选按钮才按新颜色过滤
let filterKeep = new Set();
const $ = id => document.getElementById(id);
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const nextColor = (color, direction = 1) => REVIEW_COLORS[(REVIEW_COLORS.indexOf(color) + direction + 3) % 3];
const shownCards = () => batchState.cards.filter(c => {
  if (selectedFilter === 'all') return true;
  if (c.color === selectedFilter) filterKeep.add(c.id);
  return filterKeep.has(c.id);
});
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
  for (const id of ['startBtn', 'scanOnlyBtn', 'commitBtn', 'folderInput']) $(id).disabled = busy;
  $('unhighlightBtn').disabled = busy || !batchState.highlighted;
  $('unhighlightBtn').textContent = `🧽 清除 Excel 标记颜色${batchState.highlighted ? `（${batchState.highlighted}）` : ''}`;
  $('commitBtn').disabled = busy || !batchState.cards.length;
  $('largeBtn').disabled = !batchState.cards.length;
  $('largeBtn').title = batchState.cards.length ? '逐张放大查看卡片' : '有待审核卡片后可用';
  $('openAllBtn').disabled = busy || !manualCards().length;
  $('recheckAllBtn').disabled = busy || !manualCards().length;
  $('pauseBtn').classList.toggle('hidden', !batchState.busy || batchState.job?.kind !== 'batch_scan');
  $('pauseBtn').disabled = requestBusy || !!batchState.job?.pause_requested;
  $('resumeBtn').classList.toggle('hidden', batchState.busy || batchState.job?.status !== 'interrupted');
  $('resumeBtn').disabled = requestBusy;
  $('abandonBtn').classList.toggle('hidden', batchState.busy || !batchState.job?.pending_adjustment);
  $('abandonBtn').disabled = requestBusy;
  renderPendingNotice();
  renderSelectBar();
  $('statusText').textContent = batchState.busy ? '⏳ 后台处理中' : batchState.job?.status === 'interrupted' ? '⏸ 可继续任务' : '就绪';
}

// 复制报错 / 日志：页面跑在 localhost，一般能用 clipboard；不行就退回 execCommand
const errorLogLines = () => batchState.logs.filter(line => /❌|⚠️/.test(line));

async function copyText(text, button) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const area = Object.assign(document.createElement('textarea'), { value: text });
    area.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.append(area);
    area.select();
    const ok = document.execCommand('copy');
    area.remove();
    if (!ok) { $('errorText').textContent = '复制失败，请手动选中文字复制'; return; }
  }
  const label = button.dataset.label || (button.dataset.label = button.textContent);
  button.textContent = '已复制 ✓';
  clearTimeout(button.copyTimer);
  button.copyTimer = setTimeout(() => { button.textContent = label; }, 1500);
}

function renderCopyErrorBtn() {
  const hasError = $('errorText').textContent.trim() || !$('pendingNotice').classList.contains('hidden');
  $('copyErrorBtn').classList.toggle('hidden', !hasError);
}

// 写入核对没通过时任务停住、什么都不能开始：写清楚卡在哪、两个按钮分别做什么
function renderPendingNotice() {
  const pending = !batchState.busy && batchState.job?.pending_adjustment;
  $('pendingNotice').classList.toggle('hidden', !pending);
  if (!pending) return;
  const name = String(pending.path || '').split(/[\\/]/).pop();
  const rows = (pending.rows || []).map(r => `${r.sheet} 第 ${r.row} 行`).join('、');
  const resume = batchState.job.status === 'interrupted' ? '点「继续上次任务」重新核对并写入；' : '';
  $('pendingNotice').textContent = `上次写入没有完成：${name}（${rows}）。${resume}如果你在 Excel 里改过这一行，或不想再自动改它，点「放弃这次写入」：文件不再改动，这几格转为黄色人工处理，之后可以继续处理其他文件。`;
}

// 图片取不到时换成文字占位（整页图连同红框一起换掉）
const missingHandler = (text, wholeParent = false) => `${wholeParent ? 'this.parentElement.replaceChildren' : 'this.replaceWith'}(Object.assign(document.createElement('span'),{className:'batch-missing',textContent:'${text}'}))`;

function contextHtml(card) {
  const context = card.context;
  if (!context?.image) return '<aside class="batch-context"><h3>问题所在页面</h3><div class="batch-context-missing">暂无整页定位图<br>未保存整页截图或未匹配到文字位置，可点击文件名在 Excel 中核对。</div></aside>';
  const width = Number(context.page_width), height = Number(context.page_height);
  const box = context.bbox;
  if (!(Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0)
      || !Array.isArray(box) || box.length !== 4 || !box.every(Number.isFinite)) {
    return '<aside class="batch-context"><h3>问题所在页面</h3><div class="batch-context-missing">暂无可靠的页面定位</div></aside>';
  }
  const clamp = value => Math.max(0, Math.min(100, value));
  const left = clamp(box[0] / width * 100), top = clamp(box[1] / height * 100);
  const right = clamp(box[2] / width * 100), bottom = clamp(box[3] / height * 100);
  const page = Number.isInteger(context.page_index) ? ` · 第 ${context.page_index + 1} 页` : '';
  return `<aside class="batch-context"><h3>问题所在页面${page}</h3><p class="batch-context-note">检测时的整页截图 · 红框为问题位置</p>
    <div class="batch-context-page" style="--page-ratio:${width / height}">
      <img src="${esc(context.image)}" alt="问题所在整页，红框标出原检测位置" onerror="${missingHandler('整页截图暂不可用', true)}">
      <span class="batch-context-marker" aria-label="问题位置" style="left:${left}%;top:${top}%;width:${Math.max(0, right - left)}%;height:${Math.max(0, bottom - top)}%"></span>
    </div></aside>`;
}

const pt = value => `${Number(value).toFixed(1).replace(/\.0$/, '')}pt`;

// 检测到什么、怎么处理的：让人不用对比两张图也知道为什么要修、改了多少
function problemHtml(card) {
  const problems = (card.problems || []).map(p => `<b>${esc(p.label)}</b>：${esc(p.reason)}${p.missing ? `（没显示出来的：「${esc(p.missing)}」）` : ''}`);
  return problems.length ? `<p class="batch-problem">检测到 ${problems.join('；')}</p>` : '';
}

function handlingText(card) {
  const row = card.row_height;
  if (card.after_kind === 'attempt') {
    const base = Number.isFinite(card.base_height) ? `${pt(card.base_height)} → ` : '';
    const added = Number.isFinite(card.base_height) && Number.isFinite(card.attempt_height) ? `（+${pt(card.attempt_height - card.base_height)}）` : '';
    const wrap = card.attempt_wrap?.length ? `${card.cell_addr} 开启自动换行，` : '';
    return Number.isFinite(card.attempt_height) ? `处理：${wrap}第 ${card.excel_row} 行行高试到 ${base}${pt(card.attempt_height)}${added}仍未通过，原文件未修改` : '处理：原文件未修改';
  }
  const brk = card.page_break;
  if (brk?.saved && brk.applied) return `处理：在第 ${brk.row} 行上方插入分页，整块移到下一页，已保存`;
  if (brk && !brk.applied && card.finalized) return '处理：已去掉插入的分页';
  if (card.kind === 'pagebreak' && card.after_kind === 'rechecked') return '处理：未插入分页（其他行调整后重新导出已不再被切开）';
  if (card.after_kind === 'scan_only') return `处理：只扫描，未自动修改${card.highlighted ? '；已在 Excel 中给这一格标色' : ''}`;
  if (card.after_kind === 'unchanged') return '处理：原样导出已完整显示，未修改行高';
  if (row?.saved && Math.abs(row.applied - row.base) >= 0.05) return `处理：${row.wrap ? `${card.cell_addr} 开启自动换行，` : ''}第 ${card.excel_row} 行行高 ${pt(row.base)} → ${pt(row.applied)}（+${pt(row.applied - row.base)}），已保存`;
  if (row && !row.saved && Math.abs(row.applied - row.base) < 0.05 && card.finalized) return `处理：已还原为原行高 ${pt(row.base)}${row.wrap ? '，并取消自动换行' : ''}`;
  if (card.after_kind === 'rechecked') return '处理：本行行高未修改（其他行调整后重新导出已完整显示）';
  return '';
}

function blueNote(card) {
  if (card.adoptable) {
    const height = Number.isFinite(card.attempt_height) ? ` ${pt(card.attempt_height)}` : '';
    const wrap = card.attempt_wrap?.length ? `给 ${card.cell_addr} 开启自动换行，` : '';
    return card.color === 'blue'
      ? `已按你的判断标蓝：提交时${wrap}把第 ${card.excel_row} 行行高改为${height}（最后一次尝试）并保存，记为已修复。`
      : `自动检查没通过，但如果「最后一次尝试」的图已经显示完整，可以标蓝：提交时${wrap}按${height}写入原文件。`;
  }
  return '三种颜色均可自由选择。标蓝表示你认可当前结果；没有可采用的试修方案时，提交只按人工确认放行，不额外修改 Excel。';
}

function cardHtml(card, large = false) {
  const link = excelLinkHtml(esc(card.wb_name), { file: card.wb_name, folder: batchState.folder, sheet: card.sheet_name, cell: card.cell_addr });
  const labels = { before: '修改前', after: card.after_kind === 'attempt' ? '最后一次尝试（未采用）' : card.after_kind === 'unchanged' ? '原样导出（未修改）' : card.after_kind === 'rechecked' ? '最新导出（本行未修改）' : card.after_kind === 'scan_only' ? '人工修改后（复检通过后显示）' : '修改后' };
  const handling = handlingText(card);
  const images = ['before', 'after'].map(which => `<figure><figcaption>${labels[which]}</figcaption>${card.images?.[which]
    ? `<img ${large ? '' : 'loading="lazy"'} src="${esc(card.images[which])}" alt="${labels[which]}截图" data-preview="1" onerror="${missingHandler('截图暂不可用')}">`
    : '<span class="batch-missing">暂无截图</span>'}</figure>`).join('');
  const comparison = `<div class="batch-shots">${images}</div>`;
  const visuals = large ? `<div class="batch-large-layout">${comparison}${contextHtml(card)}</div>` : comparison;
  return `<article class="batch-card" tabindex="0" role="button" aria-label="${esc(`${card.wb_name} ${card.sheet_name}!${card.cell_addr} ${REVIEW_LABELS[card.color]}，单击切换状态`)}" data-id="${esc(card.id)}" data-color="${card.color}">
    <div class="batch-card-head"><button type="button" class="batch-select" data-action="select" aria-pressed="false" title="选中 / 取消选中">✓</button><div class="batch-card-title">${link}<p class="batch-cell">${esc(card.sheet_name)}!${esc(card.cell_addr || '定位待确认')}</p></div><span class="batch-color-label">${REVIEW_LABELS[card.color]}</span></div>
    ${visuals}${problemHtml(card)}${handling ? `<p class="batch-handling">${esc(handling)}</p>` : ''}${card.text ? `<p class="batch-text">${esc(card.text)}</p>` : ''}${card.message ? `<p class="batch-message">${esc(card.message)}</p>` : ''}
    <p class="batch-state">${card.finalized ? '已提交 · 可人工处理' : '颜色标记待提交'}${card.highlighted && card.after_kind !== 'scan_only' ? ' · 已在 Excel 中标色' : ''}${batchState.busy ? ' · 后台任务结束后可标记' : ''}</p>
    ${card.verified ? '' : `<p class="batch-no-blue">${blueNote(card)}</p>`}
    <div class="batch-card-footer"><button class="btn btn-outline" data-action="large"><span class="batch-icon-enlarge" aria-hidden="true"></span>大图审核</button>${card.color === 'yellow' && card.finalized ? '<button class="btn btn-outline" data-action="recheck">🔁 单项复检</button>' : ''}</div></article>`;
}

// 选中状态不进卡片 HTML：勾选只改类名，不重画卡片、不重新加载图片
function applySelection() {
  for (const element of document.querySelectorAll('.batch-card[data-id]')) {
    const selected = selectedIds.has(element.dataset.id);
    element.classList.toggle('is-selected', selected);
    element.querySelector('.batch-select')?.setAttribute('aria-pressed', String(selected));
  }
  renderSelectBar();
}

function renderSelectBar() {
  const busy = batchState.busy || requestBusy;
  $('selectBar').classList.toggle('hidden', largeMode || !batchState.cards.length);
  $('selectCount').textContent = (selectedIds.size ? `已选 ${selectedIds.size} 张` : '未选择卡片') + (selectNotice ? ` · ${selectNotice}` : '');
  for (const button of $('selectBar').querySelectorAll('[data-set-color]')) button.disabled = busy || !selectedIds.size;
}

function toggleSelected(id) {
  if (selectedIds.has(id)) selectedIds.delete(id); else selectedIds.add(id);
  selectNotice = '';
  applySelection();
}

function selectShown(mode) {
  const ids = shownCards().map(c => c.id);
  if (mode === 'all') ids.forEach(id => selectedIds.add(id));
  else if (mode === 'invert') ids.forEach(id => selectedIds.has(id) ? selectedIds.delete(id) : selectedIds.add(id));
  else selectedIds.clear();
  selectNotice = '';
  applySelection();
}

async function changeSelectedColor(color) {
  if (batchState.busy || requestBusy || !selectedIds.size) return;
  const ids = batchState.cards.filter(c => selectedIds.has(c.id)).map(c => c.id);
  await action(async () => {
    await post('/api/batch/mark_many', { card_ids: ids, color });
    selectNotice = '';
  });
}

// 整页重画前记下视口里一张颜色没变的卡片的位置，重画后滚回同一位置，页面不跳
function scrollAnchor() {
  const colors = new Map(batchState.cards.map(c => [c.id, c.color]));
  let fallback = null;
  for (const element of $('cardList').querySelectorAll('.batch-card[data-id]')) {
    const rect = element.getBoundingClientRect();
    if (rect.bottom < 90 || rect.top > innerHeight) continue;
    const anchor = { id: element.dataset.id, top: rect.top };
    if (colors.get(element.dataset.id) === element.dataset.color) return anchor;
    fallback = fallback || anchor;
  }
  return fallback;
}

function restoreAnchor(anchor) {
  const element = anchor && $('cardList').querySelector(`[data-id="${CSS.escape(anchor.id)}"]`);
  if (element) window.scrollBy(0, element.getBoundingClientRect().top - anchor.top);
}

function render() {
  renderControls();
  // 只保留当前筛选下看得到的卡片的选中，批量改色不会改到看不见的卡片
  const visible = new Set(shownCards().map(c => c.id));
  for (const id of selectedIds) if (!visible.has(id)) selectedIds.delete(id);
  const p = batchState.progress || {};
  const job = batchState.job;
  const total = job?.workbooks?.length || p.total || 0;
  const completed = job?.kind === 'batch_scan' ? job.completed.length : p.completed || 0;
  $('progressText').textContent = `${batchState.busy ? '正在处理' : job?.status === 'interrupted' ? '已暂停或部分文件未完成' : '当前进度'}：${completed} / ${total} 个文件 · 待审核 ${batchState.cards.length} 个单元格`;
  $('progressFill').style.width = `${total ? Math.round(completed * 100 / total) : 0}%`;
  $('logBox').textContent = batchState.logs.join('\n') || '等待操作…';
  const counts = Object.fromEntries(REVIEW_COLORS.map(c => [c, batchState.cards.filter(card => card.color === c).length]));
  // 筛选按钮用对应颜色的边框、底色和圆点；选中时整块填色
  const filterHtml = ['all', ...REVIEW_COLORS].map(color => `<button type="button" class="batch-filter" data-filter="${color}" aria-pressed="${selectedFilter === color}"><span class="batch-dot" aria-hidden="true"></span>${color === 'all' ? '全部' : REVIEW_LABELS[color]}<span class="batch-count">${color === 'all' ? batchState.cards.length : counts[color]}</span></button>`).join('');
  if ($('filters').innerHTML !== filterHtml) $('filters').innerHTML = filterHtml;
  const key = JSON.stringify([batchState.cards, selectedFilter, largeMode, activeCardId, batchState.busy]);
  if (renderKey === key) { applySelection(); return; }
  renderKey = key;
  $('cardList').classList.toggle('hidden', largeMode);
  $('largeView').classList.toggle('hidden', !largeMode);
  const shown = shownCards();
  if (largeMode) {
    // 当前卡片被提交移出列表：停在原位置（即下一张），不跳回第一张
    if (!shown.some(c => c.id === activeCardId)) activeCardId = shown[Math.min(lastActiveIndex, shown.length - 1)]?.id;
    const index = shown.findIndex(c => c.id === activeCardId);
    if (index >= 0) lastActiveIndex = index;
    $('largeIndex').textContent = shown.length ? `${index + 1} / ${shown.length}` : '没有当前颜色的卡片';
    $('largeCard').innerHTML = index >= 0 ? cardHtml(shown[index], true) : '<p class="batch-empty">当前筛选下没有卡片</p>';
    $('prevBtn').disabled = index <= 0;
    $('nextBtn').disabled = index >= shown.length - 1;
  } else {
    const columnCount = Math.max(1, Math.min(3, Math.floor((($('cardList').clientWidth || 1200) + 18) / 328)));
    // 不按颜色分组：卡片按处理顺序固定排列，改色只换颜色，位置不变
    const columns = Array.from({ length: columnCount }, () => []);
    shown.forEach((card, index) => columns[index % columnCount].push(card));
    const html = new Map(shown.map(c => [c.id, cardHtml(c)]));
    const layout = JSON.stringify([columns.map(column => column.map(c => c.id)), batchState.busy]);
    if (layout === layoutKey && shown.length) {
      // 卡片位置没变：只换内容有变化的卡片，不整页重画（图片不重新加载、页面不跳）
      for (const [id, value] of html) {
        if (cardHtmlCache.get(id) === value) continue;
        const element = $('cardList').querySelector(`[data-id="${CSS.escape(id)}"]`);
        if (element) element.outerHTML = value;
      }
    } else {
      const anchor = scrollAnchor();
      $('cardList').innerHTML = shown.length
        ? `<div class="batch-waterfall">${columns.map(column => `<div class="batch-column">${column.map(c => html.get(c.id)).join('')}</div>`).join('')}</div>`
        : `<section class="panel batch-empty"><h2>${batchState.busy ? '正在自动扫描与修复…' : '当前没有待审核卡片'}</h2><p>${batchState.busy ? '文件处理后会在这里出现修改前后对比，无需逐项放行。' : '开始全自动处理，或切换颜色筛选查看结果。'}</p></section>`;
      restoreAnchor(anchor);
    }
    layoutKey = layout;
    cardHtmlCache = html;
  }
  applySelection();
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
  const color = nextColor(card.color, direction);
  await action(() => post('/api/batch/mark', { id, color }));
}

// 只扫描时 Excel 里的标记颜色：预设 + 自定义，记在本机浏览器
const HIGHLIGHT_KEY = 'clear_handle_excel.highlightColors';
const HIGHLIGHT_PRESETS = [
  { label: '洋红底白字（推荐）', fill: '#C0007A', font: '#FFFFFF' },
  { label: '宝蓝底白字', fill: '#1649CC', font: '#FFFFFF' },
  { label: '墨绿底白字', fill: '#006633', font: '#FFFFFF' },
  { label: '深青底白字', fill: '#006A71', font: '#FFFFFF' },
  { label: '深紫底白字', fill: '#6D28D9', font: '#FFFFFF' },
  { label: '橙红底白字', fill: '#C2410C', font: '#FFFFFF' },
  { label: '深灰底白字', fill: '#333333', font: '#FFFFFF' },
  { label: '纯黑底黄字', fill: '#000000', font: '#FFFF00' },
];
const HEX = /^#[0-9a-f]{6}$/i;
let highlight = { fill: HIGHLIGHT_PRESETS[0].fill, font: HIGHLIGHT_PRESETS[0].font };
try {
  const saved = JSON.parse(localStorage.getItem(HIGHLIGHT_KEY) || 'null');
  if (HEX.test(saved?.fill) && HEX.test(saved?.font)) highlight = { fill: saved.fill.toUpperCase(), font: saved.font.toUpperCase() };
} catch (_) {}

function luminance(hex) {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255)
    .map(c => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
const contrastRatio = (a, b) => { const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x); return (hi + 0.05) / (lo + 0.05); };

function renderHighlight() {
  $('highlightFill').value = highlight.fill.toLowerCase();
  $('highlightFont').value = highlight.font.toLowerCase();
  $('highlightSample').style.background = highlight.fill;
  $('highlightSample').style.color = highlight.font;
  $('highlightPresets').innerHTML = HIGHLIGHT_PRESETS.map((p, i) => `<button type="button" class="batch-highlight-preset" data-preset="${i}" aria-pressed="${p.fill === highlight.fill && p.font === highlight.font}" style="background:${p.fill};color:${p.font}">${esc(p.label)}</button>`).join('');
  const ratio = contrastRatio(highlight.fill, highlight.font);
  $('highlightContrast').textContent = `对比度 ${ratio.toFixed(1)}:1${ratio < 4.5 ? '，偏低，字可能看不清' : ''}`;
  $('highlightContrast').classList.toggle('is-low', ratio < 4.5);
}

function setHighlight(next) {
  highlight = { fill: next.fill.toUpperCase(), font: next.font.toUpperCase() };
  try { localStorage.setItem(HIGHLIGHT_KEY, JSON.stringify(highlight)); } catch (_) {}
  renderHighlight();
}
$('highlightPresets').addEventListener('click', event => {
  const button = event.target.closest('[data-preset]');
  if (button) setHighlight(HIGHLIGHT_PRESETS[Number(button.dataset.preset)]);
});
const HIGHLIGHT_AUTO_KEY = 'clear_handle_excel.highlightAuto';
try { $('highlightAuto').checked = localStorage.getItem(HIGHLIGHT_AUTO_KEY) !== '0'; } catch (_) { $('highlightAuto').checked = true; }
$('highlightAuto').addEventListener('change', event => {
  try { localStorage.setItem(HIGHLIGHT_AUTO_KEY, event.target.checked ? '1' : '0'); } catch (_) {}
});
$('highlightFill').addEventListener('input', event => setHighlight({ ...highlight, fill: event.target.value }));
$('highlightFont').addEventListener('input', event => setHighlight({ ...highlight, font: event.target.value }));

// 暂停的全自动处理在提交、复检、打开文件后会放回来，不算替换续跑点
const KEEPS_PAUSED_SCAN = new Set(['/api/batch/commit', '/api/batch/recheck', '/api/batch/open', '/api/batch/unhighlight']);
const waitingCard = card => card.waiting ?? card.message === '等待自动处理';

async function startJob(url, body = {}) {
  const keeps = batchState.job?.kind === 'batch_scan' && KEEPS_PAUSED_SCAN.has(url);
  if (batchState.job?.status === 'interrupted' && url !== '/api/resume' && !keeps
      && !(await appConfirm('这会替换上次任务的续跑点。已保存的修复和审核标记会保留，继续吗？', { okText: '继续' }))) return;
  await action(() => post(url, body));
}

async function startBatch(extra = {}) {
  const folder = $('folderInput').value.trim();
  if (!folder) { $('errorText').textContent = '请输入 Excel 文件夹路径'; return; }
  const unsent = batchState.cards.filter(c => !c.finalized).length;
  if (unsent && batchState.folder && folder !== batchState.folder
      && !(await appConfirm(`当前文件夹还有 ${unsent} 张卡片没有提交审核。\n审核记录会保留，以后对原文件夹重新开始全自动处理即可继续审核。\n\n确定换到新文件夹吗？`, { title: '换文件夹', okText: '换到新文件夹' }))) return;
  await startJob('/api/batch/start', { folder, ...extra }).then(() => {
    if (batchState.busy) { try { localStorage.setItem('clear_handle_excel.lastFolder', folder); } catch (_) {} }
  });
}
$('startBtn').addEventListener('click', () => startBatch($('highlightAuto').checked ? { highlight: { ...highlight } } : {}));
$('scanOnlyBtn').addEventListener('click', async () => {
  if (!$('folderInput').value.trim()) { $('errorText').textContent = '请输入 Excel 文件夹路径'; return; }
  if (!(await appConfirm('不自动加高，检出的问题全部转为黄色人工项。\n会在原 Excel 里给问题单元格换上标记底色和字色并保存（复检通过或提交蓝 / 绿后自动还原）。\n\n确定开始吗？', { title: '只扫描，全部转人工', okText: '开始扫描' }))) return;
  await startBatch({ scan_only: true, highlight: { ...highlight } });
});
$('unhighlightBtn').addEventListener('click', async () => {
  if (await appConfirm(`把当前文件夹 ${batchState.highlighted || 0} 个标过颜色的单元格还原成原来的颜色？\n需要先在 Excel 里保存并关闭这些文件；你自己改过颜色的格子不动。之后提交、复检也不再自动标色，重新开始处理时才按设置标色。`, { title: '清除 Excel 标记颜色', okText: '还原颜色' })) await startJob('/api/batch/unhighlight');
});
$('resumeBtn').addEventListener('click', () => startJob('/api/resume'));
$('pauseBtn').addEventListener('click', () => action(() => post('/api/batch/pause')));
$('commitBtn').addEventListener('click', async () => {
  const waiting = batchState.cards.filter(waitingCard).length;
  const cards = batchState.cards.filter(c => !waitingCard(c));
  if (!cards.length) { $('errorText').textContent = '卡片都还没轮到自动处理，暂时没有可提交的标记'; return; }
  const count = color => cards.filter(c => c.color === color).length;
  const unmarked = cards.filter(c => !c.marked && !c.finalized).length;
  const message = `将提交 ${cards.length} 张卡片：\n蓝（保留修复）${count('blue')}　绿（还原并忽略）${count('green')}　黄（还原后人工）${count('yellow')}`
    + (cards.some(c => c.color === 'blue' && c.adoptable) ? `\n\n其中 ${cards.filter(c => c.color === 'blue' && c.adoptable).length} 张蓝色是自动检查没通过、由你看图认可的：会按最后一次尝试的行高写入原文件。` : '')
    + (unmarked ? `\n\n其中 ${unmarked} 张没有手动改过颜色，将按当前颜色处理。` : '')
    + (waiting ? `\n\n另有 ${waiting} 张还没轮到自动处理，这次不提交，继续全自动处理后再审核。` : '') + '\n\n确定提交吗？';
  if (await appConfirm(message, { title: '确认并处理标记', okText: '提交', tone: 'info' })) await startJob('/api/batch/commit', { card_ids: cards.map(c => c.id) });
});
$('abandonBtn').addEventListener('click', async () => {
  if (await appConfirm('放弃上次未完成的写入？\n不会再改动这个文件，相关单元格转为黄色人工处理，请在 Excel 中核对这一行的行高。', { title: '放弃这次写入', okText: '放弃写入', tone: 'danger' })) action(() => post('/api/batch/abandon'));
});
$('copyErrorBtn').addEventListener('click', event => {
  const parts = [$('pendingNotice').classList.contains('hidden') ? '' : $('pendingNotice').textContent,
                 $('errorText').textContent.trim()].filter(Boolean);
  const recent = errorLogLines().slice(-10);
  if (recent.length) parts.push('最近的报错日志：\n' + recent.join('\n'));
  copyText(parts.join('\n\n'), event.currentTarget);
});
$('copyLogBtn').addEventListener('click', event => copyText(batchState.logs.join('\n') || '（暂无日志）', event.currentTarget));
$('copyLogErrorsBtn').addEventListener('click', event => copyText(errorLogLines().join('\n') || '（日志里没有报错）', event.currentTarget));
// 报错文字在很多地方被设置/清空，统一盯着它显示或隐藏复制按钮
new MutationObserver(renderCopyErrorBtn).observe($('errorText'), { childList: true, characterData: true, subtree: true });
new MutationObserver(renderCopyErrorBtn).observe($('pendingNotice'), { attributes: true, attributeFilter: ['class'] });
$('openAllBtn').addEventListener('click', () => startJob('/api/batch/open', { card_ids: manualCards().map(c => c.id) }));
$('recheckAllBtn').addEventListener('click', () => startJob('/api/batch/recheck', { card_ids: manualCards().map(c => c.id) }));
$('largeBtn').addEventListener('click', () => openLarge());
$('closeLargeBtn').addEventListener('click', () => { largeMode = false; layoutKey = ''; render(); });
$('prevBtn').addEventListener('click', () => moveCard(-1));
$('nextBtn').addEventListener('click', () => moveCard(1));
$('selectBar').addEventListener('click', event => {
  const button = event.target.closest('button');
  if (button?.dataset.select) selectShown(button.dataset.select);
  else if (button?.dataset.setColor) changeSelectedColor(button.dataset.setColor);
});
$('filters').addEventListener('click', event => {
  const button = event.target.closest('[data-filter]');
  if (button) { selectedFilter = button.dataset.filter; filterKeep = new Set(); renderKey = ''; render(); }
});

for (const id of ['cardList', 'largeCard']) $(id).addEventListener('click', event => {
  const card = event.target.closest('[data-id]');
  if (!card || event.target.closest('a')) return;
  const button = event.target.closest('[data-action]');
  if (button?.dataset.action === 'select') toggleSelected(card.dataset.id);
  else if (button?.dataset.action === 'large') openLarge(card.dataset.id);
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
    if (event.key === 'Escape') { largeMode = false; layoutKey = ''; render(); }
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
window.addEventListener('resize', () => { renderKey = ''; layoutKey = ''; render(); });
try { $('folderInput').value = localStorage.getItem('clear_handle_excel.lastFolder') || ''; } catch (_) {}
renderHighlight();
poll();
setInterval(poll, 1800);
