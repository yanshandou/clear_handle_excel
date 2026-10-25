'use strict';
const REVIEW_COLORS = ['blue', 'green', 'yellow'];
const REVIEW_LABELS = { blue: '蓝 · 已修复', green: '绿 · 还原并忽略', yellow: '黄 · 人工处理' };
const REVIEW_JOB_NAMES = { batch_commit: '提交审核', batch_recheck: '复检', batch_open: '打开文件', batch_unhighlight: '清除标记颜色' };
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
    celebrateFinished(batchState, value);
    if (value.busy && value.job?.kind === 'batch_unhighlight') clearPending = true;
    const cleared = clearPending && !value.busy && value.job?.kind === 'batch_unhighlight';
    if (cleared) clearPending = false;
    batchState = value;
    if (!document.activeElement || document.activeElement.id !== 'folderInput') {
      if (value.folder && !$('folderInput').value) $('folderInput').value = value.folder;
      if (value.busy && value.folder) $('folderInput').value = value.folder;
    }
    render();
    if (cleared && unhighlightResult()) showUnhighlightResult();
  } catch (error) { $('errorText').textContent = error.message; }
}

// 清除标记的结果：任务结束后弹窗，之后可点「上次清除结果」再看。
// 清除任务很快时轮询可能一次都没看到「进行中」，所以点了清除（或继续清除任务）就先记下来
let clearPending = false;
function unhighlightResult() {
  const job = batchState.job;
  return job?.kind === 'batch_unhighlight' && job.result && Object.keys(job.result).length ? job.result : null;
}

const CLEAR_STATUS = { ok: '✅', check: '⚠️', failed: '❌' };
function showUnhighlightResult() {
  const files = Object.values(unhighlightResult() || {});
  const count = { ok: 0, check: 0, failed: 0 };
  for (const file of files) for (const cell of file.cells) count[cell.status] += 1;
  const errors = files.filter(f => f.error).length;
  const leftovers = files.reduce((sum, f) => sum + (f.leftovers || 0), 0);
  const trouble = f => !!f.error || (f.leftover_failed || []).length > 0 || f.cells.some(c => c.status !== 'ok');
  files.sort((a, b) => trouble(b) - trouble(a) || a.name.localeCompare(b.name));
  const chips = [`<span class="unhighlight-chip is-ok">✅ 已清除 ${count.ok} 格</span>`];
  if (count.check) chips.push(`<span class="unhighlight-chip is-check">⚠️ 需要看一下 ${count.check} 格</span>`);
  if (count.failed) chips.push(`<span class="unhighlight-chip is-failed">❌ 没清掉 ${count.failed} 格</span>`);
  if (errors) chips.push(`<span class="unhighlight-chip is-failed">❌ ${errors} 份文件没能处理</span>`);
  if (leftovers) chips.push(`<span class="unhighlight-chip">🧹 另清掉残留标记 ${leftovers} 条</span>`);
  const fileHtml = file => {
    const own = { ok: 0, check: 0, failed: 0 };
    for (const cell of file.cells) own[cell.status] += 1;
    const tally = file.error ? `<span class="unhighlight-file-error">${esc(file.error)}</span>`
      : file.cells.length ? `✅ ${own.ok}${own.check ? ` · ⚠️ ${own.check}` : ''}${own.failed ? ` · ❌ ${own.failed}` : ''}` : '没有要清的标记';
    const rows = file.cells.map(cell => {
      const where = `${cell.sheet}!${cell.addr}`;
      const link = excelLinkHtml(esc(where), { file: file.name, folder: batchState.folder, sheet: cell.sheet, cell: cell.addr.split(':')[0] });
      return `<li class="is-${cell.status}"><span aria-hidden="true">${CLEAR_STATUS[cell.status] || ''}</span>${link}<span class="unhighlight-reason">${esc(cell.reason)}</span></li>`;
    });
    for (const reason of file.leftover_failed || []) {
      rows.push(`<li class="is-failed"><span aria-hidden="true">❌</span><span>残留标记</span><span class="unhighlight-reason">删不掉：${esc(reason)}</span></li>`);
    }
    return `<details class="unhighlight-file"${trouble(file) ? ' open' : ''}><summary><b>${esc(file.name)}</b><span class="unhighlight-tally">${tally}</span></summary>`
      + (rows.length ? `<ul>${rows.join('')}</ul>` : '') + '</details>';
  };
  const tone = count.failed || errors ? 'error' : count.check ? 'warn' : 'info';
  $('unhighlightDialog').dataset.tone = tone;
  $('unhighlightTitle').textContent = tone === 'info' ? '标记已全部清除' : '清除结果：有格子要看一下';
  $('unhighlightBody').innerHTML = `<p class="unhighlight-summary">${files.length} 份文件　${chips.join('')}</p>`
    + `<p class="unhighlight-hint">点格子位置可以用 Excel 打开并跳到那一格。⚠️ 和 ❌ 的请人工看一下。</p>`
    + `<div class="unhighlight-files">${files.map(fileHtml).join('')}</div>`;
  if (!$('unhighlightDialog').open) $('unhighlightDialog').showModal();
  $('unhighlightBody').scrollTop = 0;
  $('unhighlightCloseBtn').focus(); // 不然浏览器把焦点放在第一个文件上，显示一圈焦点框
}

// 后台任务刚结束且成功：撒花。提交 / 复检时若有暂停的全自动处理会被放回来，任务种类变了也算提交完成
const FINISH_TITLES = { batch_scan: '全自动处理完成！', batch_commit: '审核已提交！', batch_recheck: '复检完成！' };
function celebrateFinished(before, after) {
  const kind = before.job?.kind;
  if (!before.busy || after.busy || !FINISH_TITLES[kind]) return;
  if (after.job?.kind === kind && after.job?.status !== 'completed') return;
  const left = after.cards.length;
  const title = kind === 'batch_scan' && before.job.options?.scan_only ? '扫描完成！'
    : kind === 'batch_scan' && before.job.to_manual ? '已转人工！' : FINISH_TITLES[kind];
  window.gameFx?.celebrate(title, left ? `还有 ${left} 张卡片待审核` : '全部卡片都处理完了');
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
  $('unhighlightResultBtn').classList.toggle('hidden', batchState.busy || !unhighlightResult());
  $('commitBtn').disabled = busy || !batchState.cards.length;
  $('largeBtn').disabled = !batchState.cards.length;
  $('largeBtn').title = batchState.cards.length ? '逐张放大查看卡片' : '有待审核卡片后可用';
  $('openAllBtn').disabled = busy || !manualCards().length;
  // 复检进行中还能接着点复检：排进同一个任务，不用等它结束
  const rechecking = batchState.busy && batchState.job?.kind === 'batch_recheck';
  $('recheckAllBtn').disabled = (busy && !rechecking) || requestBusy || !manualCards().length;
  $('pauseBtn').classList.toggle('hidden', !batchState.busy || batchState.job?.kind !== 'batch_scan');
  $('pauseBtn').disabled = requestBusy || !!batchState.job?.pause_requested;
  $('toManualBtn').classList.toggle('hidden', !batchState.mode?.to_manual);
  $('toManualBtn').disabled = requestBusy;
  $('toAutoBtn').classList.toggle('hidden', !batchState.mode?.to_auto);
  $('toAutoBtn').disabled = busy;
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
  if (card.after_kind === 'handover') return '处理：转人工时自动还没处理到这一格，原文件未修改';
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
  const labels = { before: '修改前', after: card.after_kind === 'attempt' ? '最后一次尝试（未采用）' : card.after_kind === 'unchanged' ? '原样导出（未修改）' : card.after_kind === 'rechecked' ? '最新导出（本行未修改）' : ['scan_only', 'handover'].includes(card.after_kind) ? '人工修改后（复检通过后显示）' : '修改后' };
  const handling = handlingText(card);
  const images = ['before', 'after'].map(which => `<figure><figcaption>${labels[which]}</figcaption>${card.images?.[which]
    ? `<img ${large ? '' : 'loading="lazy"'} src="${esc(card.images[which])}" alt="${labels[which]}截图" data-preview="1" onerror="${missingHandler('截图暂不可用')}">`
    : '<span class="batch-missing">暂无截图</span>'}</figure>`);
  // 左右两块斜切外框中间放 VS 徽章；截图本身保持方正，不倾斜不变形
  const comparison = `<div class="batch-shots">${images[0]}<span class="batch-vs" aria-hidden="true">VS</span>${images[1]}</div>`;
  const visuals = large ? `<div class="batch-large-layout">${comparison}${contextHtml(card)}</div>` : comparison;
  return `<article class="batch-card" tabindex="0" role="button" aria-label="${esc(`${card.wb_name} ${card.sheet_name}!${card.cell_addr} ${REVIEW_LABELS[card.color]}，单击切换状态`)}" data-id="${esc(card.id)}" data-color="${card.color}">
    <div class="batch-card-head"><button type="button" class="batch-select" data-action="select" aria-pressed="false" title="选中 / 取消选中">✓</button><div class="batch-card-title">${link}<p class="batch-cell">${esc(card.sheet_name)}!${esc(card.cell_addr || '定位待确认')}</p></div><span class="batch-color-label">${REVIEW_LABELS[card.color]}</span></div>
    ${visuals}${problemHtml(card)}${handling ? `<p class="batch-handling">${esc(handling)}</p>` : ''}${card.text ? `<p class="batch-text">${esc(card.text)}</p>` : ''}${card.message ? `<p class="batch-message">${esc(card.message)}</p>` : ''}
    <p class="batch-state">${card.finalized ? '已提交 · 可人工处理' : '颜色标记待提交'}${card.highlighted && card.after_kind !== 'scan_only' ? ' · 已在 Excel 中标色' : ''}${card.handed_over ? ' · 自动还没试过，可转回自动' : ''}${batchState.busy ? ' · 后台任务结束后可标记' : ''}</p>
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

// 瀑布流分列：每张卡片放进当时最矮的一列（按实际高度算），放好后就记住，不再挪动。
// 以前按序号轮流分到各列，截图和文字高度不一，卡片一多各列高度越差越大，后面只剩一列
const CARD_GAP = 18;
let columnOf = new Map();
let columnMeta = '';
const cardHeights = new Map();

function cardHeight(id) {
  if (cardHeights.has(id)) return cardHeights.get(id);
  let sum = 0;
  for (const value of cardHeights.values()) sum += value;
  return cardHeights.size ? sum / cardHeights.size : 520;
}

function shortestColumn(heights) {
  return heights.reduce((best, h, i) => (h < heights[best] - 1 ? i : best), 0);
}

// 已放好的卡片保持原列；列数或筛选变了、或有卡片被移出（提交后）时整体重排。返回这次新放的卡片
function placeCards(shown, count) {
  const meta = JSON.stringify([count, selectedFilter]);
  const ids = new Set(shown.map(c => c.id));
  if (meta !== columnMeta || [...columnOf.keys()].some(id => !ids.has(id))) {
    columnOf = new Map();
    columnMeta = meta;
  }
  const heights = Array(count).fill(0);
  for (const card of shown) if (columnOf.has(card.id)) heights[columnOf.get(card.id)] += cardHeight(card.id);
  const fresh = new Set();
  for (const card of shown) {
    if (columnOf.has(card.id)) continue;
    const column = shortestColumn(heights);
    columnOf.set(card.id, column);
    heights[column] += cardHeight(card.id);
    fresh.add(card.id);
  }
  return fresh;
}

function cardColumns(shown, count) {
  const columns = Array.from({ length: count }, () => []);
  for (const card of shown) columns[columnOf.get(card.id)].push(card);
  return columns;
}

// 量出卡片真实高度，按真实高度重新给新卡片选列；有变动就把卡片节点挪过去（图片不重新加载）
function rebalanceFresh(shown, count, fresh) {
  const columnElements = $('cardList').querySelectorAll('.batch-waterfall > .batch-column');
  if (columnElements.length !== count) return false;
  const elements = new Map();
  for (const element of $('cardList').querySelectorAll('.batch-card[data-id]')) {
    elements.set(element.dataset.id, element);
    cardHeights.set(element.dataset.id, element.getBoundingClientRect().height + CARD_GAP);
  }
  if (!fresh.size) return false;
  const heights = Array(count).fill(0);
  for (const card of shown) if (!fresh.has(card.id)) heights[columnOf.get(card.id)] += cardHeight(card.id);
  let moved = false;
  for (const card of shown) {
    if (!fresh.has(card.id)) continue;
    const column = shortestColumn(heights);
    if (columnOf.get(card.id) !== column) moved = true;
    columnOf.set(card.id, column);
    heights[column] += cardHeight(card.id);
  }
  if (!moved) return false;
  cardColumns(shown, count).forEach((column, i) => columnElements[i].append(...column.map(c => elements.get(c.id)).filter(Boolean)));
  return true;
}

function render() {
  renderControls();
  // 只保留当前筛选下看得到的卡片的选中，批量改色不会改到看不见的卡片
  const visible = new Set(shownCards().map(c => c.id));
  for (const id of selectedIds) if (!visible.has(id)) selectedIds.delete(id);
  const p = batchState.progress || {};
  const job = batchState.job;
  // 全自动处理看本任务的文件；审核提交 / 复检等运行中看它自己涉及的文件；其余时候显示上次扫描的进度
  const own = job && Array.isArray(job.workbooks) && (job.kind === 'batch_scan' || (job.kind?.startsWith('batch_') && job.status !== 'completed'));
  // 全自动处理先把全部文件扫完：扫描阶段按扫过的文件算
  const scanning = own && job.kind === 'batch_scan' && !job.options?.scan_only && job.phase !== 'repair';
  const done = own ? new Set((scanning ? job.scanned : job.completed) || []) : null;
  const total = own ? job.workbooks.length : p.total || 0;
  // 已完成只算本任务列表里的文件（去重），不会超过总数
  const completed = own ? job.workbooks.filter(path => done.has(path)).length : Math.min(p.completed || 0, total);
  const label = batchState.busy ? (own && job.kind !== 'batch_scan' ? `正在${REVIEW_JOB_NAMES[job.kind] || '处理'}`
    : job?.to_manual ? '做完当前这一行后转人工' : scanning && job.options?.back_to_auto ? '正在准备转回自动（改过的文件重新检测）'
    : scanning ? '正在扫描（全部扫完后开始修，之后可以转人工）'
    : job?.options?.scan_only ? '正在扫描' : '正在自动处理')
    : job?.status === 'interrupted' ? '已暂停或部分文件未完成' : '当前进度';
  $('progressText').textContent = `${label}：${completed} / ${total} 个文件 · 待审核 ${batchState.cards.length} 个单元格`;
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
    const fresh = placeCards(shown, columnCount);
    const columns = cardColumns(shown, columnCount);
    const html = new Map(shown.map(c => [c.id, cardHtml(c)]));
    let layout = JSON.stringify([columns.map(column => column.map(c => c.id)), batchState.busy]);
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
      // 新卡片先按估计高度放下，量出真实高度后在浏览器画出来之前挪到最矮的列，用户看不到挪动
      if (rebalanceFresh(shown, columnCount, fresh)) layout = JSON.stringify([cardColumns(shown, columnCount).map(column => column.map(c => c.id)), batchState.busy]);
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

// 查找残留标记色：旧版本直接改了单元格颜色、记录又丢了的，按标记色完全一致找出来，只列出、不改文件
let findingMarks = false;
async function findMarks() {
  if (findingMarks) return;
  const folder = $('folderInput').value.trim() || batchState.folder;
  const colors = [...HIGHLIGHT_PRESETS.map(p => ({ fill: p.fill, font: p.font })), { ...highlight }];
  $('errorText').textContent = '';
  try { await post('/api/batch/find_marks', { folder, colors }); } catch (error) { $('errorText').textContent = error.message; return; }
  await waitFindMarks('正在查找残留标记色…', '正在查找', status => showFoundMarks(status.result));
}

// 查找 / 对比旧版 / 改回颜色都在后台线程里做：弹窗显示进度，做完换成结果
async function waitFindMarks(title, doing, done) {
  findingMarks = true;
  $('findMarksBtn').disabled = true;
  $('findMarksTitle').textContent = title;
  $('findMarksBody').innerHTML = '<p class="unhighlight-summary" id="findMarksProgress">准备中…</p>';
  $('findMarksApplyBtn').classList.add('hidden');
  $('findMarksCloseBtn').textContent = '在后台继续';
  $('findMarksCloseBtn').className = 'btn btn-primary';
  if (!$('findMarksDialog').open) $('findMarksDialog').showModal();
  try {
    for (;;) {
      const response = await fetch('/api/batch/find_marks', { cache: 'no-store' });
      const status = await response.json();
      if (!status.running) { done(status); break; }
      const progress = $('findMarksProgress');
      if (progress) progress.textContent = `${doing} ${Math.min(status.done + 1, status.total)} / ${status.total}：${status.current || ''}`;
      await new Promise(resolve => setTimeout(resolve, 800));
    }
  } catch (error) { $('errorText').textContent = `${doing}失败：${error.message}`; }
  finally { findingMarks = false; $('findMarksBtn').disabled = false; }
}

const OLD_FOLDER_KEY = 'clear_handle_excel.oldFolder';
function savedOldFolder() {
  try { return localStorage.getItem(OLD_FOLDER_KEY) || ''; } catch (_) { return ''; }
}

async function previewRestore() {
  const old = $('oldFolderInput')?.value.trim() || '';
  if (!old) { $('oldFolderError').textContent = '请输入旧版（原始备份）文件夹路径'; return; }
  try { localStorage.setItem(OLD_FOLDER_KEY, old); } catch (_) {}
  try { await post('/api/batch/find_marks/preview', { folder: old }); } catch (error) { $('oldFolderError').textContent = error.message; return; }
  await waitFindMarks('正在对比旧版…', '正在对比', status => showRestorePlan(status.plan));
}

function swatchHtml(fill, font) {
  if (fill === 'none') return '<span class="find-marks-swatch is-none" title="无填充">无</span>';
  if (fill === 'other') return '<span class="find-marks-swatch is-none" title="主题色或图案，改回时照旧版原样抄">原</span>';
  return `<span class="find-marks-swatch" style="background:${esc(fill)};color:${esc(font || '#000000')}">Aa</span>`;
}

function cellLink(file, folder, cell) {
  return excelLinkHtml(esc(`${cell.sheet}!${cell.addr}`), { file: file.name, folder, sheet: cell.sheet, cell: cell.addr.split(':')[0] });
}

function showRestorePlan(plan) {
  if (!plan) return;
  const cells = plan.files.flatMap(f => f.cells);
  const restore = cells.filter(c => c.action === 'restore').length;
  const missing = plan.files.filter(f => f.error).length;
  const chips = [`<span class="unhighlight-chip is-ok">✅ 会改回 ${restore} 格</span>`];
  if (cells.length - restore) chips.push(`<span class="unhighlight-chip is-check">⚠️ 不改、请人工看 ${cells.length - restore} 格</span>`);
  if (missing) chips.push(`<span class="unhighlight-chip is-failed">❌ ${missing} 份文件没法对比</span>`);
  const fileHtml = file => {
    const rows = file.cells.map(c => `<li class="is-${c.action === 'restore' ? 'ok' : 'check'}">`
      + (c.action === 'restore' ? swatchHtml(c.old_fill, c.old_font) : '<span aria-hidden="true">⚠️</span>')
      + `${cellLink(file, plan.folder || batchState.folder, c)}<span class="unhighlight-reason">${c.action === 'restore'
        ? `改回旧版：底色 ${c.old_fill === 'none' ? '无填充' : c.old_fill === 'other' ? '主题色 / 图案（照旧版抄）' : esc(c.old_fill)}${c.old_font ? `，字色 ${esc(c.old_font)}` : ''}${c.old_sheet && c.old_sheet !== c.sheet ? `（旧版表名「${esc(c.old_sheet)}」）` : ''}`
        : esc(c.reason)}</span></li>`);
    const own = file.cells.filter(c => c.action === 'restore').length;
    const tally = file.error ? `<span class="unhighlight-file-error">${esc(file.error)}</span>` : `✅ ${own}${file.cells.length - own ? ` · ⚠️ ${file.cells.length - own}` : ''}`;
    return `<details class="unhighlight-file"${plan.files.length <= 5 || file.error || own < file.cells.length ? ' open' : ''}><summary><b>${esc(file.name)}</b><span class="unhighlight-tally">${tally}</span></summary>`
      + (rows.length ? `<ul>${rows.join('')}</ul>` : '') + '</details>';
  };
  $('findMarksTitle').textContent = '对比旧版：确认后才改文件';
  $('findMarksDialog').dataset.tone = 'warn';
  $('findMarksBody').innerHTML = `<p class="unhighlight-summary">${chips.join('')}</p>`
    + `<p class="unhighlight-hint">旧版文件夹：${esc(plan.old_folder)}。只改同一位置内容和旧版一样的格子，只改底色和字色。请先在 Excel 里保存并关闭这些文件。</p>`
    + `<div class="unhighlight-files">${plan.files.map(fileHtml).join('')}</div>`;
  $('findMarksApplyBtn').textContent = `确认改回 ${restore} 格`;
  $('findMarksApplyBtn').classList.toggle('hidden', !restore);
  $('findMarksCloseBtn').textContent = '取消';
  $('findMarksCloseBtn').className = 'btn btn-outline';
  $('findMarksBody').scrollTop = 0;
}

async function applyRestore() {
  try { await post('/api/batch/find_marks/apply'); } catch (error) { $('errorText').textContent = error.message; return; }
  await waitFindMarks('正在改回旧版颜色…', '正在改', status => showRestoreResult(status.applied));
}

function showRestoreResult(applied) {
  if (!applied) return;
  const cells = applied.files.flatMap(f => f.cells);
  const ok = cells.filter(c => c.status === 'ok').length;
  const errors = applied.files.filter(f => f.error).length;
  const chips = [`<span class="unhighlight-chip is-ok">✅ 已改回 ${ok} 格</span>`];
  if (cells.length - ok) chips.push(`<span class="unhighlight-chip is-failed">❌ 没改 ${cells.length - ok} 格</span>`);
  if (errors) chips.push(`<span class="unhighlight-chip is-failed">❌ ${errors} 份文件没能处理</span>`);
  const fileHtml = file => {
    const rows = file.cells.map(c => `<li class="is-${c.status === 'ok' ? 'ok' : 'failed'}"><span aria-hidden="true">${c.status === 'ok' ? '✅' : '❌'}</span>`
      + `${cellLink(file, applied.folder || batchState.folder, c)}<span class="unhighlight-reason">${esc(c.reason)}</span></li>`);
    const own = file.cells.filter(c => c.status === 'ok').length;
    const tally = file.error ? `<span class="unhighlight-file-error">${esc(file.error)}</span>` : `✅ ${own}${file.cells.length - own ? ` · ❌ ${file.cells.length - own}` : ''}`;
    return `<details class="unhighlight-file"${file.error || own < file.cells.length ? ' open' : ''}><summary><b>${esc(file.name)}</b><span class="unhighlight-tally">${tally}</span></summary>`
      + (rows.length ? `<ul>${rows.join('')}</ul>` : '') + '</details>';
  };
  $('findMarksTitle').textContent = cells.length - ok || errors ? '改回结果：有格子没改' : '已全部改回旧版的颜色';
  $('findMarksDialog').dataset.tone = cells.length - ok || errors ? 'warn' : 'info';
  $('findMarksBody').innerHTML = `<p class="unhighlight-summary">${chips.join('')}</p>`
    + `<div class="unhighlight-files">${applied.files.map(fileHtml).join('')}</div>`;
  $('findMarksApplyBtn').classList.add('hidden');
  $('findMarksCloseBtn').textContent = '知道了';
  $('findMarksCloseBtn').className = 'btn btn-primary';
  $('findMarksCloseBtn').focus();
}

function showFoundMarks(result) {
  if (!result) return;
  const files = result.files;
  const hits = files.flatMap(f => f.hits);
  const both = hits.filter(h => h.match === 'both').length;
  const recorded = hits.filter(h => h.recorded).length;
  const unread = files.filter(f => f.error).length;
  const chips = [`<span class="unhighlight-chip">${result.total} 份文件</span>`];
  chips.push(`<span class="unhighlight-chip ${hits.length ? 'is-check' : 'is-ok'}">🎯 底色字色都对上 ${both} 格</span>`);
  if (hits.length - both) chips.push(`<span class="unhighlight-chip is-check">🎨 只有底色对上 ${hits.length - both} 格</span>`);
  if (recorded) chips.push(`<span class="unhighlight-chip is-ok">📝 程序有记录 ${recorded} 格</span>`);
  if (unread) chips.push(`<span class="unhighlight-chip is-failed">❌ ${unread} 份文件读不了</span>`);
  const fileHtml = file => {
    const rows = file.hits.map(hit => {
      const link = excelLinkHtml(esc(`${hit.sheet}!${hit.addr}`), { file: file.name, folder: result.folder, sheet: hit.sheet, cell: hit.addr.split(':')[0] });
      const swatch = `<span class="find-marks-swatch" style="background:${esc(hit.fill)};color:${esc(hit.font || '#000000')}">Aa</span>`;
      const note = (hit.match === 'both' ? '底色、字色都和标记色一致' : `只有底色一致（字色 ${esc(hit.font || '不是直接设置的颜色')}）`)
        + (hit.recorded ? '；程序有记录，点「清除 Excel 标记颜色」会还原' : '');
      return `<li class="is-${hit.recorded ? 'ok' : 'check'}">${swatch}${link}<span class="unhighlight-reason">${note}</span></li>`;
    });
    const tally = file.error ? `<span class="unhighlight-file-error">${esc(file.error)}</span>`
      : `${file.hits.length} 格${file.truncated ? '（只列前 500 格）' : ''}`;
    return `<details class="unhighlight-file"${files.length <= 5 || file.error ? ' open' : ''}><summary><b>${esc(file.name)}</b><span class="unhighlight-tally">${tally}</span></summary>`
      + (rows.length ? `<ul>${rows.join('')}</ul>` : '') + '</details>';
  };
  $('findMarksTitle').textContent = hits.length ? '找到底色是标记色的格子' : '没有找到残留标记色';
  $('findMarksDialog').dataset.tone = hits.length || unread ? 'warn' : 'info';
  $('findMarksBody').innerHTML = `<p class="unhighlight-summary">${chips.join('')}</p>`
    + '<p class="unhighlight-hint">查找只列出来，不改文件：程序不知道这些格子原来是什么颜色，也可能本来就是这个颜色。点位置用 Excel 打开看一下；有原始备份的话，可以在下面对比旧版，确认后自动改回去。</p>'
    + (hits.length ? '<div class="find-marks-old"><label for="oldFolderInput">旧版（原始备份）文件夹</label>'
      + `<input id="oldFolderInput" class="input" type="text" placeholder="例如 D:\\原始备份" value="${esc(savedOldFolder())}">`
      + '<button type="button" id="previewRestoreBtn" class="btn btn-outline">对比旧版，看看要改回什么</button><span id="oldFolderError" class="batch-error"></span></div>' : '')
    + `<div class="unhighlight-files">${files.map(fileHtml).join('')}</div>`;
  $('previewRestoreBtn')?.addEventListener('click', previewRestore);
  $('findMarksCloseBtn').textContent = '知道了';
  if (!$('findMarksDialog').open) $('findMarksDialog').showModal();
  $('findMarksBody').scrollTop = 0;
  $('findMarksCloseBtn').focus();
}

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
  const clearing = url === '/api/batch/unhighlight' || (url === '/api/resume' && batchState.job?.kind === 'batch_unhighlight');
  await action(async () => {
    await post(url, body);
    if (clearing) clearPending = true;
  });
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
  if (await appConfirm(`去掉当前文件夹 ${batchState.highlighted || 0} 个单元格的标记颜色？\n标记是程序加的条件格式，只删这些规则；单元格本身的颜色和表里原有的条件格式都不动。\n需要先在 Excel 里保存并关闭这些文件。之后提交、复检也不再自动标色，重新开始处理时才按设置标色。`, { title: '清除 Excel 标记颜色', okText: '去掉标记' })) await startJob('/api/batch/unhighlight');
});
$('unhighlightResultBtn').addEventListener('click', showUnhighlightResult);
$('findMarksBtn').addEventListener('click', findMarks);
$('findMarksCloseBtn').addEventListener('click', () => $('findMarksDialog').close());
$('findMarksApplyBtn').addEventListener('click', applyRestore);
$('unhighlightCloseBtn').addEventListener('click', () => $('unhighlightDialog').close());
$('resumeBtn').addEventListener('click', () => startJob('/api/resume'));
$('pauseBtn').addEventListener('click', () => action(() => post('/api/batch/pause')));
$('toManualBtn').addEventListener('click', async () => {
  const marking = batchState.job?.options?.highlight ? '，并在 Excel 里给这些格子标色' : '';
  if (!(await appConfirm(`做完当前这一行就停，还没轮到自动处理的问题全部转为黄色人工项${marking}。\n已经修好的蓝卡保留在原文件里，可以随时审核。\n人工处理累了可以点「转回自动」接着修。\n\n确定转人工吗？`, { title: '转人工', okText: '转人工' }))) return;
  await action(() => post('/api/batch/to_manual'));
});
$('toAutoBtn').addEventListener('click', async () => {
  const count = batchState.cards.filter(c => c.handed_over).length;
  if (!(await appConfirm(`把 ${count} 处自动还没试过的黄色问题交回自动处理。\n请先在 Excel 里保存并关闭这些文件。\n转人工后改过的文件整份重新检测；你改过行高的行、自动试过没通过的问题仍留给人工。\nExcel 里的标记颜色会先还原。\n\n确定转回自动吗？`, { title: '转回自动', okText: '转回自动' }))) return;
  await startJob('/api/batch/to_auto', $('highlightAuto').checked ? { highlight: { ...highlight } } : {});
});
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
