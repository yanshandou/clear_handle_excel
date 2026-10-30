'use strict';
// PDF 人工复检：导出的 PDF 每页铺满页面宽度，按住拖动画圆角框圈出问题，双击框删除，Ctrl+Z 撤销。
// 框的位置按 PDF 的点（pt）记，页面缩放后照样对得上；每改一次就把这份 PDF 的全部框存到后台。
// 「写进 Excel」由后台找出框住的单元格，在原 Excel 里标色。

const FOLDER_KEY = 'clear_handle_excel.lastFolder';
const MIN_DRAG_PX = 6; // 拖得比这还小当成点了一下，不算一个框
const STATE_LABEL = { marked: '已标进 Excel', pending: '待写入', error: '没标成' };

let folder = '';
let pdfs = [];          // [{pdf, name, v, pages: [[宽, 高]], wb_path, wb_name, sheet, guessed, error}]
let rects = {};         // pdf -> [框]
let dropCount = 0;      // 删掉了、Excel 里的标记还没去掉的框
let job = { running: false };
let undoStack = [];     // [{pdf, op: 'add' | 'del', rect}]
let drawing = null;
let polling = null;
const saving = {};      // pdf -> 正在进行的保存（同一份 PDF 按顺序存）
let pageEls = [];       // 所有页面元素，从上到下
let whereDoc = -1;      // 底部条上现在显示的是哪一份（换了才重建链接）
let whereTick = 0;
let opened = new Set(); // 这一轮已经用「打开下一批」打开过的 Excel 路径
let opening = false;

const $ = id => document.getElementById(id);

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

function newId() {
  const bytes = new Uint8Array(8);
  crypto.getRandomValues(bytes);
  return [...bytes].map(b => b.toString(16).padStart(2, '0')).join('');
}

function highlight() {
  let saved = null;
  try { saved = normalizeHighlight(JSON.parse(localStorage.getItem(HIGHLIGHT_KEY) || 'null')); } catch (_) {}
  return saved || { fill: HIGHLIGHT_PRESETS[0].fill, font: HIGHLIGHT_PRESETS[0].font };
}

async function api(url, body) {
  const res = await fetch(url, body === undefined ? {} : {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.ok === false) throw new Error(data.error || `HTTP ${res.status}`);
  return data;
}

function excelTarget(doc, cell = '') {
  return { file: doc.wb_name, sheet: doc.sheet || '', cell, folder };
}

/* ---------- 列出 PDF ---------- */

async function load() {
  const value = $('prFolder').value.trim();
  if (!value) {
    appAlert('请先输入 Excel 文件夹路径！', { tone: 'warn' });
    return;
  }
  $('prLoad').disabled = true;
  $('prStatus').textContent = '正在读取 PDF…';
  try {
    const data = await api(`/api/pdf_review?folder=${encodeURIComponent(value)}`);
    folder = data.folder;
    try { localStorage.setItem(FOLDER_KEY, value); } catch (_) {}
    pdfs = data.pdfs;
    undoStack = [];
    takeSnapshot(data);
    renderDocs();
  } catch (e) {
    $('prStatus').textContent = `读取失败：${e.message}`;
    appAlert(e.message, { tone: 'error' });
  } finally {
    $('prLoad').disabled = false;
  }
}

function takeSnapshot(data) {
  rects = {};
  for (const doc of pdfs) rects[doc.pdf] = (data.rects[doc.pdf] || []).slice();
  dropCount = data.drop || 0;
  job = data.job || { running: false };
  opened = new Set(data.opened || []);
}

function renderDocs() {
  const box = $('prDocs');
  if (!pdfs.length) {
    box.innerHTML = '<p class="pr-empty">这个文件夹里没有 PDF。先在主页用「导出 PDF」导出，再回来复检。</p>';
    pageEls = [];
    $('prBar').classList.add('hidden');
    updateStatus();
    return;
  }
  box.innerHTML = pdfs.map((doc, d) => {
    const excel = doc.wb_name
      ? `<span class="pr-doc-excel">📗 ${excelLinkHtml(`${escapeHtml(doc.wb_name)} › ${escapeHtml(doc.sheet)}`, excelTarget(doc))}`
        + `${doc.guessed ? '<span class="pr-doc-note" title="导出记录里没有这份 PDF，按「Excel 文件名_表名.pdf」找到的">按文件名对应</span>' : ''}</span>`
      : '<span class="pr-doc-warn">⚠ 找不到对应的 Excel，可以标记但不能写进 Excel</span>';
    const pages = doc.pages.map(([w, h], p) => `
      <div class="pr-page is-loading" data-doc="${d}" data-page="${p}" style="aspect-ratio:${w} / ${h}">
        <img class="pr-img" loading="lazy" decoding="async" draggable="false" alt="第 ${p + 1} 页"
          src="/api/pdf_review/page?pdf=${encodeURIComponent(doc.pdf)}&p=${p}&v=${doc.v}">
        <div class="pr-layer"></div>
        ${doc.wb_name ? `<div class="pr-chip">📗 ${excelLinkHtml(`${escapeHtml(doc.wb_name)} › ${escapeHtml(doc.sheet)}`, excelTarget(doc))}</div>` : ''}
        <span class="pr-pageno">${p + 1} / ${doc.pages.length}</span>
      </div>`).join('');
    return `<article class="pr-doc" id="prDoc${d}">
      <header class="pr-doc-head" title="${escapeHtml(doc.pdf)}">
        <span class="pr-doc-name">📄 ${escapeHtml(doc.name)}</span>
        <span class="pr-doc-pages">${doc.pages.length} 页</span>${excel}
        <span class="pr-doc-count" data-count="${d}"></span>
      </header>
      ${doc.error ? `<p class="pr-doc-warn">${escapeHtml(doc.error)}</p>` : pages}
    </article>`;
  }).join('');
  pdfs.forEach((_, d) => renderRects(d));
  $('prBar').classList.remove('hidden');
  pageEls = [...box.querySelectorAll('.pr-page')];
  renderTrack();
  updateStatus();
}

/* ---------- 看到哪里（底部条） ---------- */

const READ_LINE = 0.4; // 窗口上方 40% 处那条线落在哪一页，就算正在看哪一页

function renderTrack() {
  const total = pageEls.length;
  let before = 0;
  const ticks = pdfs.slice(0, -1).map(doc => {
    before += doc.error ? 0 : doc.pages.length;
    return total ? `<span class="pr-tick" style="left:${before / total * 100}%"></span>` : '';
  }).join('');
  $('prTrack').innerHTML = `<div id="prFill" class="pr-fill"></div>${ticks}`;
  whereDoc = -1;
  updateWhere();
}

// 读线所在的页（序号）和读线在这一页里的位置（0~1）；页面从上到下排，二分查找
function readingPosition() {
  // 滚到页面最底下时读线到不了最后一页的末尾，直接算看完了
  if (window.innerHeight + window.scrollY >= document.documentElement.scrollHeight - 2) {
    const last = pageEls[pageEls.length - 1].getBoundingClientRect();
    if (last.top < window.innerHeight) return { index: pageEls.length - 1, within: 1 };
  }
  const line = window.innerHeight * READ_LINE;
  let lo = 0, hi = pageEls.length - 1, found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (pageEls[mid].getBoundingClientRect().top <= line) { found = mid; lo = mid + 1; } else hi = mid - 1;
  }
  if (found < 0) return { index: 0, within: 0 };
  const box = pageEls[found].getBoundingClientRect();
  return { index: found, within: Math.min(1, Math.max(0, (line - box.top) / (box.height || 1))) };
}

function updateWhere() {
  whereTick = 0;
  if (!pageEls.length) {
    $('prWhereText').innerHTML = '<span class="pr-where-pos">没有可看的页面</span>';
    $('prFill').style.width = '0';
    whereDoc = -1;
    return;
  }
  const { index, within } = readingPosition();
  const page = pageEls[index];
  const d = Number(page.dataset.doc);
  const p = Number(page.dataset.page);
  const doc = pdfs[d];
  $('prFill').style.width = `${(index + within) / pageEls.length * 100}%`;
  const pos = `第 ${d + 1} / ${pdfs.length} 份 · 第 ${p + 1} / ${doc.pages.length} 页`;
  if (whereDoc !== d) {
    whereDoc = d;
    const excel = doc.wb_name
      ? `<span class="pr-where-excel" title="${escapeHtml(doc.wb_name)} › ${escapeHtml(doc.sheet)}">📗 ${excelLinkHtml(`${escapeHtml(doc.wb_name)} › ${escapeHtml(doc.sheet)}`, excelTarget(doc))}</span>`
      : `<span class="pr-doc-warn" title="${escapeHtml(doc.pdf)}">⚠ ${escapeHtml(doc.name)} 找不到对应的 Excel</span>`;
    $('prWhereText').innerHTML = `<span class="pr-where-pos"></span>${excel}`;
  }
  $('prWhereText').querySelector('.pr-where-pos').textContent = pos;
}

function scheduleWhere() {
  if (!whereTick) whereTick = requestAnimationFrame(updateWhere);
}

// 点进度条：按点的位置换算成第几页的哪里，把那里滚到读线上
function jumpTo(e) {
  if (!pageEls.length) return;
  const box = $('prTrack').getBoundingClientRect();
  const at = Math.min(0.9999, Math.max(0, (e.clientX - box.left) / box.width)) * pageEls.length;
  const page = pageEls[Math.floor(at)].getBoundingClientRect();
  window.scrollBy({ top: page.top + page.height * (at % 1) - window.innerHeight * READ_LINE, behavior: 'smooth' });
}

/* ---------- 框 ---------- */

function rectBadge(doc, rect) {
  if (rect.state === 'marked' && rect.addr && doc.wb_name) {
    const first = rect.addr.split(/[,:]/)[0];
    return `<span class="pr-badge is-marked">${excelLinkHtml(escapeHtml(rect.addr), excelTarget({ ...doc, sheet: rect.sheet || doc.sheet }, first))}</span>`;
  }
  if (rect.state === 'error') return `<span class="pr-badge is-error" title="${escapeHtml(rect.error)}">⚠ 没标成</span>`;
  return '<span class="pr-badge">待写入</span>';
}

function renderRects(d) {
  const doc = pdfs[d];
  const byPage = {};
  for (const rect of rects[doc.pdf] || []) (byPage[rect.page] = byPage[rect.page] || []).push(rect);
  document.querySelectorAll(`.pr-page[data-doc="${d}"]`).forEach(page => {
    const p = Number(page.dataset.page);
    const [w, h] = doc.pages[p];
    page.querySelector('.pr-layer').innerHTML = (byPage[p] || []).map(rect => {
      const title = rect.state === 'error' ? `没标成：${rect.error}` : `${STATE_LABEL[rect.state]}${rect.addr ? `：${rect.sheet || doc.sheet}!${rect.addr}` : ''}`;
      const inside = rect.y0 / h < 0.04 ? ' badge-inside' : ''; // 贴着页顶的框：标签放进框里，免得被页面边缘切掉
      return `<div class="pr-rect is-${rect.state}${inside}" data-id="${rect.id}" title="${escapeHtml(title)}（双击删除）"
        style="left:${rect.x0 / w * 100}%;top:${rect.y0 / h * 100}%;width:${(rect.x1 - rect.x0) / w * 100}%;height:${(rect.y1 - rect.y0) / h * 100}%">
        ${rectBadge(doc, rect)}</div>`;
    }).join('');
  });
  const list = rects[doc.pdf] || [];
  const count = document.querySelector(`[data-count="${d}"]`);
  if (count) count.textContent = list.length ? `标记 ${list.length} 处` : '';
}

function updateStatus() {
  const all = Object.values(rects).flat();
  const pending = all.filter(r => r.state !== 'marked').length;
  const marked = all.length - pending;
  const pages = pdfs.reduce((n, doc) => n + doc.pages.length, 0);
  $('prStatus').textContent = pdfs.length
    ? `共 ${pdfs.length} 份 PDF、${pages} 页 · 标记 ${all.length} 处（已标进 Excel ${marked} 处）`
    : $('prStatus').textContent;
  const parts = [`待写入 <b>${pending}</b> 处`];
  if (dropCount) parts.push(`待去掉 <b>${dropCount}</b> 处`);
  if (job.running) parts.push(`正在写进 Excel：${job.done || 0} / ${job.total || 0} 个文件${job.current ? `（${escapeHtml(job.current)}）` : ''}`);
  $('prSummary').innerHTML = parts.join(' · ');
  $('prApply').disabled = job.running || (!pending && !dropCount);
  const files = markedFiles();
  const unopened = files.filter(path => !opened.has(path)).length;
  $('prOpen').textContent = opening ? '⏳ 正在打开…' : `📂 打开下一批（未开 ${unopened} / 共 ${files.length} 个）`;
  $('prOpen').disabled = opening || job.running || !unopened;
  $('prOpenReset').disabled = opening || job.running || !files.some(path => opened.has(path));
  $('prApply').textContent = job.running ? '⏳ 正在写进 Excel…' : '🖍️ 写进 Excel';
}

function docIndex(pdf) {
  return pdfs.findIndex(doc => doc.pdf === pdf);
}

// 同一份 PDF 的保存排队进行，后发的一定后到
function save(pdf) {
  const d = docIndex(pdf);
  const body = { pdf, rects: rects[pdf].map(({ id, page, x0, y0, x1, y1 }) => ({ id, page, x0, y0, x1, y1 })) };
  saving[pdf] = (saving[pdf] || Promise.resolve()).then(async () => {
    try {
      const data = await api('/api/pdf_review/save', body);
      mergeStates(data);
    } catch (e) {
      appAlert(`标记没保存上：${e.message}`, { tone: 'error' });
    }
    renderRects(d);
    updateStatus();
  });
}

// 后台返回的状态（已标 / 待写入 / 没标成、单元格地址）合进来；框的增删以网页上为准（后面可能还有没存完的）
function mergeStates(data) {
  for (const doc of pdfs) {
    const server = new Map((data.rects[doc.pdf] || []).map(r => [r.id, r]));
    for (const rect of rects[doc.pdf] || []) {
      const s = server.get(rect.id);
      if (s) Object.assign(rect, { state: s.state, addr: s.addr, sheet: s.sheet, error: s.error });
    }
  }
  dropCount = data.drop || 0;
  job = data.job || job;
  if (data.opened) opened = new Set(data.opened);
}

function addRect(pdf, rect, record = true) {
  rects[pdf].push(rect);
  if (record) undoStack.push({ pdf, op: 'add', rect });
  renderRects(docIndex(pdf));
  updateStatus();
  save(pdf);
}

function removeRect(pdf, id, record = true) {
  const list = rects[pdf];
  const i = list.findIndex(r => r.id === id);
  if (i < 0) return;
  const [rect] = list.splice(i, 1);
  if (record) undoStack.push({ pdf, op: 'del', rect });
  renderRects(docIndex(pdf));
  updateStatus();
  save(pdf);
}

function undo() {
  const step = undoStack.pop();
  if (!step) return;
  if (step.op === 'add') removeRect(step.pdf, step.rect.id, false);
  else addRect(step.pdf, step.rect, false);
}

/* ---------- 拖动画框 ---------- */

function pointIn(page, e) {
  const box = page.getBoundingClientRect();
  return {
    x: Math.min(Math.max(e.clientX - box.left, 0), box.width),
    y: Math.min(Math.max(e.clientY - box.top, 0), box.height),
    w: box.width, h: box.height,
  };
}

function onPointerDown(e) {
  const layer = e.target.closest('.pr-layer');
  if (!layer || e.button !== 0 || e.target.closest('.pr-badge') || layer.closest('.is-loading, .is-failed')) return;
  const page = layer.parentElement;
  const start = pointIn(page, e);
  const ghost = document.createElement('div');
  ghost.className = 'pr-rect is-drawing';
  layer.appendChild(ghost);
  // 不用 setPointerCapture：抓住指针后双击的目标会变成整页，双击框就删不掉了；移动和松开在 document 上听
  drawing = { page, layer, start, ghost, pointer: e.pointerId };
  e.preventDefault();
}

function drawnBox(e) {
  const end = pointIn(drawing.page, e);
  const { start } = drawing;
  return {
    left: Math.min(start.x, end.x), top: Math.min(start.y, end.y),
    width: Math.abs(end.x - start.x), height: Math.abs(end.y - start.y), w: end.w, h: end.h,
  };
}

function onPointerMove(e) {
  if (!drawing || e.pointerId !== drawing.pointer) return;
  const b = drawnBox(e);
  Object.assign(drawing.ghost.style, {
    left: `${b.left}px`, top: `${b.top}px`, width: `${b.width}px`, height: `${b.height}px`,
  });
}

function onPointerUp(e) {
  if (!drawing || e.pointerId !== drawing.pointer) return;
  const b = drawnBox(e);
  const { page, ghost } = drawing;
  ghost.remove();
  drawing = null;
  if (b.width < MIN_DRAG_PX || b.height < MIN_DRAG_PX) return;
  const doc = pdfs[Number(page.dataset.doc)];
  const p = Number(page.dataset.page);
  const [w, h] = doc.pages[p];
  const k = w / b.w; // 网页像素 → PDF 点
  const round = v => Math.round(v * 100) / 100;
  addRect(doc.pdf, {
    id: newId(), page: p, state: 'pending', addr: null, sheet: doc.sheet, error: null,
    x0: round(b.left * k), y0: round(b.top * k),
    x1: round(Math.min(w, (b.left + b.width) * k)), y1: round(Math.min(h, (b.top + b.height) * k)),
  });
}

function cancelDrawing() {
  if (!drawing) return;
  drawing.ghost.remove();
  drawing = null;
}

/* ---------- 写进 Excel ---------- */

async function applyToExcel() {
  const all = Object.values(rects).flat();
  const files = new Set(pdfs.filter(doc => (rects[doc.pdf] || []).some(r => r.state !== 'marked'))
    .map(doc => doc.wb_name).filter(Boolean));
  const pending = all.filter(r => r.state !== 'marked').length;
  const message = `把 ${pending} 处标记写进 Excel${dropCount ? `，并去掉 ${dropCount} 处删掉的标记` : ''}。\n`
    + `后台会打开${files.size ? ` ${files.size} 个` : ''} Excel 文件，找出框住的单元格后标色并保存。\n请先保存并关闭这些 Excel 文件。`;
  if (!(await appConfirm(message, { title: '写进 Excel', okText: '开始' }))) return;
  try {
    const data = await api('/api/pdf_review/apply', { highlight: highlight() });
    mergeStates(data);
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
    return;
  }
  updateStatus();
  startPolling();
}

function startPolling() {
  clearTimeout(polling);
  polling = setTimeout(async () => {
    let data;
    try {
      data = await api('/api/pdf_review/rects');
    } catch (_) {
      startPolling();
      return;
    }
    mergeStates(data);
    pdfs.forEach((_, d) => renderRects(d));
    updateStatus();
    if (job.running) {
      startPolling();
    } else if (job.result) {
      const failed = job.result.failed || 0;
      appAlert(`${job.result.message || '已完成'}${failed ? '\n没标成的框上写着原因（鼠标移上去看）。' : ''}`,
        { title: '写进 Excel', tone: failed ? 'warn' : undefined, icon: failed ? undefined : '✅' });
    }
  }, 1000);
}

/* ---------- 批量打开有标记的 Excel ---------- */

// 网页上画了框的 Excel（写没写进 Excel 都算），按 PDF 顺序，每份只算一次
function markedFiles() {
  return [...new Set(pdfs.filter(doc => doc.wb_path && (rects[doc.pdf] || []).length).map(doc => doc.wb_path))];
}

async function openNext() {
  if (opening) return;
  opening = true;
  updateStatus();
  try {
    const plan = await api('/api/pdf_review/open_plan', {});
    const waiting = plan.waiting.length
      ? `\n\n还要等后台处理的 ${plan.waiting.length} 个，暂不打开：\n${plan.waiting.map(w => `· ${w.name}：${w.reason}`).join('\n')}` : '';
    if (!plan.files.length) {
      await appAlert(plan.exhausted ? '画了框的 Excel 这一轮都打开过了，不会自动重复打开。需要重开时，点「重置打开记录」。'
        : `共 ${plan.total} 个画了框的 Excel，目前没有可以打开的文件。${waiting}`, { title: '打开下一批' });
      return;
    }
    const message = `网页上画了框的 Excel 共 ${plan.total} 个，这一轮还没打开的 ${plan.unopened} 个。\n`
      + `本次最多打开 5 个，实际打开 ${plan.files.length} 个；打开后还剩 ${plan.remaining} 个。\n`
      + '已经写进 Excel 的跳到第一个框的单元格，还没写的只切到对应的工作表。'
      + `\n\n本次文件：\n${plan.files.map(f => `· ${f.name}`).join('\n')}${waiting}`;
    if (!(await appConfirm(message, { title: '确认打开文件', okText: `打开 ${plan.files.length} 个文件`, tone: 'info' }))) return;
    const result = await api('/api/pdf_review/open', {});
    const fresh = result.opened.length ? await api('/api/pdf_review/rects').catch(() => null) : null;
    if (fresh) mergeStates(fresh); // 记下这批已打开，按钮上的「未开」跟着变
    if (result.errors.length || result.waiting.length) {
      const list = items => items.map(item => `· ${item.name}：${item.reason}`).join('\n');
      const parts = [result.opened.length ? `已打开 ${result.opened.length} 个文件。` : '这次没有打开文件。'];
      if (result.waiting.length) parts.push(`还要等后台处理的 ${result.waiting.length} 个：\n${list(result.waiting)}`);
      if (result.errors.length) parts.push(`打不开的 ${result.errors.length} 个：\n${list(result.errors)}`);
      await appAlert(parts.join('\n\n'), { title: '打开下一批', tone: 'warn' });
    }
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
  } finally {
    opening = false;
    updateStatus();
  }
}

async function resetOpened() {
  if (!(await appConfirm('清掉这个文件夹的打开记录，下次「打开下一批」从头开始。不会关闭 Excel，标记也不动。', { title: '重置打开记录', okText: '重置' }))) return;
  try {
    mergeStates(await api('/api/pdf_review/open_reset', {}));
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
  }
  updateStatus();
}

/* ---------- 启动 ---------- */

// 页面图片加载完才能画框：没加载出来时画在空白上，看不出框住了什么。load / error 不冒泡，在捕获阶段听
document.addEventListener('load', e => {
  if (e.target.classList && e.target.classList.contains('pr-img')) e.target.parentElement.classList.remove('is-loading');
}, true);
document.addEventListener('error', e => {
  if (e.target.classList && e.target.classList.contains('pr-img')) {
    e.target.parentElement.classList.remove('is-loading');
    e.target.parentElement.classList.add('is-failed');
  }
}, true);
document.addEventListener('pointerdown', onPointerDown);
document.addEventListener('pointermove', onPointerMove);
document.addEventListener('pointerup', onPointerUp);
document.addEventListener('pointercancel', cancelDrawing);
document.addEventListener('dblclick', e => {
  const el = e.target.closest('.pr-rect:not(.is-drawing)');
  if (!el || e.target.closest('.excel-link')) return;
  const page = el.closest('.pr-page');
  removeRect(pdfs[Number(page.dataset.doc)].pdf, el.dataset.id);
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') cancelDrawing();
  const typing = e.target.closest && e.target.closest('input, textarea, select, [contenteditable]');
  if (!typing && (e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 'z' && !document.querySelector('dialog[open]')) {
    e.preventDefault();
    undo();
  }
});
$('prLoad').addEventListener('click', load);
$('prFolder').addEventListener('keydown', e => { if (e.key === 'Enter') load(); });
$('prApply').addEventListener('click', applyToExcel);
$('prTrack').addEventListener('click', jumpTo);
$('prOpen').addEventListener('click', openNext);
$('prOpenReset').addEventListener('click', resetOpened);
window.addEventListener('scroll', scheduleWhere, { passive: true });
window.addEventListener('resize', scheduleWhere);
{
  const color = highlight();
  document.documentElement.style.setProperty('--pr-mark', color.fill);
  Object.assign($('prSample').style, { background: color.fill, color: color.font });
}
try { $('prFolder').value = localStorage.getItem(FOLDER_KEY) || ''; } catch (_) {}
if ($('prFolder').value) load().then(() => { if (job.running) startPolling(); });
