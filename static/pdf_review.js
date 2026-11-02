'use strict';
// PDF 人工复检：导出的 PDF 整页显示，看到问题点一下，就用 Excel 打开并选中那一格；
// 填了原始备份文件夹时，备份也一起打开（只读）、跳到同一格，左右对照着改。
// 每份 PDF 的行列位置由后台算（只算一次），点的格子也由后台按它算。
// 三色标记（借鉴精简 Sheet）：蓝 = 不用管（默认）、黄 = 这一页要人工改（按页标）、红 = 整张表不要了。
// 单击页面换颜色（蓝 → 黄 → 红 → 蓝，标黄记住点的那一格）；用 Excel 打开那一格改成 Ctrl + 单击或右键菜单。
// 标题行右边的三个按钮一次改整张表。
// 看完后点底部「一键处理」：后台删掉标红的表并保存，再用 Excel 打开标黄的。

const FOLDER_KEY = 'clear_handle_excel.lastFolder';
const OLD_FOLDER_KEY = 'clear_handle_excel.oldFolder'; // 和修复颜色 / 引用 / 隐藏页共用
const VIEW_KEY = 'clear_handle_excel.pdfReviewView';
const MARKS_KEY = 'clear_handle_excel.pdfReviewMarks'; // {文件夹: {pdf: {v, red, pages: {页码: {x, y}}, opened}}}
const MARK_TEXT = { blue: '不用管', yellow: '要人工改', red: '删除这张表' };
const MARK_TITLE = { blue: '整张表不用管（清掉这张表的所有标记）', yellow: '整张表每页都标黄；只标某一页就单击那一页', red: '删除整张表（所有页）' };
const VIEWS = ['large', 'medium', 'small'];
const MIN_DRAG_PX = 6;  // 按下到松开挪动超过这么多就是在拖（选字、滚动），不算点
const POLL_MS = 1500;

let folder = '';
let pdfs = [];          // [{pdf, name, v, pages: [[宽, 高]], wb_path, wb_name, sheet, guessed, error}]
let maps = {};          // pdf -> {state: ready / error / working / pending / none, error}
let pageEls = [];       // 所有页面元素，从上到下
let whereDoc = -1;      // 底部条上现在显示的是哪一份（换了才重建链接）
let whereTick = 0;
let polling = 0;
let waiting = null;     // 点了还没算好的那份：{pdf, page, x, y, el}；算好就接着打开
let opening = false;
let press = null;       // 按下的位置，松开时判断是不是点了一下
let marks = {};         // pdf -> {v, red, pages: {页码: {x, y}}, opened}；全蓝（默认）不记
let deleted = new Set(); // 这次一键处理删掉的表（PDF 也删了，页面上灰掉，刷新后就不在了）
let processing = false;

const $ = id => document.getElementById(id);

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
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

function toast(text, isError = false) {
  if (window.excelToast) window.excelToast(text, isError);
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
    maps = {};
    deleted = new Set();
    loadMarks();
    mergeMaps(data.maps);
    renderDocs();
    startPolling();
  } catch (e) {
    $('prStatus').textContent = `读取失败：${e.message}`;
    appAlert(e.message, { tone: 'error' });
  } finally {
    $('prLoad').disabled = false;
  }
}

function renderDocs() {
  const box = $('prDocs');
  if (!pdfs.length) {
    box.innerHTML = '<p class="pr-empty">这个文件夹里没有 PDF。先在主页「人工复检」里点「导出 PDF」，导完点「查看 PDF」回来。</p>';
    pageEls = [];
    $('prBar').classList.add('hidden');
    updateStatus();
    return;
  }
  box.innerHTML = pdfs.map((doc, d) => {
    const excel = doc.wb_name
      ? `<span class="pr-doc-excel">📗 ${excelLinkHtml(`${escapeHtml(doc.wb_name)} › ${escapeHtml(doc.sheet)}`, excelTarget(doc))}`
        + `${doc.guessed ? '<span class="pr-doc-note" title="导出记录里没有这份 PDF，按「Excel 文件名_表名.pdf」找到的">按文件名对应</span>' : ''}</span>`
      : '<span class="pr-doc-warn">⚠ 找不到对应的 Excel，点了也打不开</span>';
    const pages = doc.pages.map(([w, h], p) => `
      <div class="pr-page is-loading" data-doc="${d}" data-page="${p}" style="aspect-ratio:${w} / ${h};--ar:${w / h}">
        <img class="pr-img" loading="lazy" decoding="async" draggable="false" alt="第 ${p + 1} 页"
          src="/api/pdf_review/page?pdf=${encodeURIComponent(doc.pdf)}&p=${p}&v=${doc.v}">
        <div class="pr-layer"></div>
        <span class="pr-pageno">${p + 1} / ${doc.pages.length}</span>
      </div>`).join('');
    const markable = doc.wb_path && !doc.error;
    const buttons = markable ? `<span class="pr-mark" role="group" aria-label="标记">${['blue', 'yellow', 'red'].map(c =>
      `<button type="button" class="pr-mark-btn" data-mark-btn="${c}" data-doc="${d}" title="${MARK_TITLE[c]}">${MARK_TEXT[c]}</button>`).join('')}</span>` : '';
    return `<article class="pr-doc" id="prDoc${d}" data-doc="${d}" data-mark="blue">
      <header class="pr-doc-head" title="${escapeHtml(doc.pdf)}">
        <span class="pr-doc-name">📄 ${escapeHtml(doc.name)}</span>
        <span class="pr-doc-pages">${doc.pages.length} 页</span>${excel}
        <span class="pr-doc-state" data-state="${d}"></span>${buttons}
      </header>
      ${doc.error ? `<p class="pr-doc-warn">${escapeHtml(doc.error)}</p>` : `<div class="pr-pages">${pages}</div>`}
    </article>`;
  }).join('');
  $('prBar').classList.remove('hidden');
  pageEls = [...box.querySelectorAll('.pr-page')];
  pdfs.forEach((_, d) => paintMark(d));
  renderTrack();
  updateStatus();
}

/* ---------- 行列位置（后台算） ---------- */

const STATE_TEXT = {
  ready: '✓ 可以点格子', error: '⚠ 定位不了', working: '⏳ 正在定位…', pending: '⏳ 排队定位', none: '',
};

function mergeMaps(status) {
  for (const [pdf, s] of Object.entries(status || {})) maps[pdf] = { state: s.state, error: s.error };
}

function updateStatus() {
  const docs = pdfs.filter(doc => doc.wb_path && !doc.error);
  const ready = docs.filter(doc => maps[doc.pdf]?.state === 'ready').length;
  const failed = docs.filter(doc => maps[doc.pdf]?.state === 'error').length;
  const pages = pdfs.reduce((n, doc) => n + doc.pages.length, 0);
  if (pdfs.length) $('prStatus').textContent = `共 ${pdfs.length} 份 PDF、${pages} 页`;
  const parts = [`已定位 <b>${ready}</b> / ${docs.length} 份`];
  if (failed) parts.push(`<span class="pr-doc-warn">${failed} 份定位不了</span>`);
  $('prSummary').innerHTML = parts.join(' · ');
  pdfs.forEach((doc, d) => {
    const el = document.querySelector(`[data-state="${d}"]`);
    if (!el) return;
    const m = maps[doc.pdf] || { state: 'none' };
    el.textContent = doc.wb_path ? STATE_TEXT[m.state] || '' : '';
    el.dataset.kind = m.state;
    el.title = m.state === 'error' ? m.error || '' : '';
  });
}

function startPolling() {
  clearTimeout(polling);
  polling = setTimeout(async () => {
    try {
      mergeMaps((await api('/api/pdf_review/maps')).maps);
      updateStatus();
      if (waiting && ['ready', 'error'].includes(maps[waiting.pdf]?.state)) {
        const next = waiting;
        waiting = null;
        await openAt(next);
      }
    } catch (_) { /* 下一轮再试 */ }
    if (Object.values(maps).some(m => m.state === 'pending' || m.state === 'working') || waiting) startPolling();
  }, POLL_MS);
}

// 页面上的点 → {pdf, page, x, y}（PDF 点）
function pointOn(pageEl, e) {
  const doc = pdfs[Number(pageEl.dataset.doc)];
  const p = Number(pageEl.dataset.page);
  const box = pageEl.getBoundingClientRect();
  const [w, h] = doc.pages[p];
  return { doc, pdf: doc.pdf, page: p, x: (e.clientX - box.left) / box.width * w, y: (e.clientY - box.top) / box.height * h };
}

/* ---------- 点一下打开 ---------- */

function flashAt(pageEl, e) {
  const box = pageEl.getBoundingClientRect();
  const dot = Object.assign(document.createElement('span'), { className: 'pr-click' });
  dot.style.left = `${e.clientX - box.left}px`;
  dot.style.top = `${e.clientY - box.top}px`;
  pageEl.querySelector('.pr-layer').appendChild(dot);
  setTimeout(() => dot.remove(), 700);
}

async function openAt(pt) {
  if (opening) return;
  const oldFolder = $('prOldFolder').value.trim();
  opening = true;
  toast(oldFolder ? '正在打开当前文件和备份…' : '正在用 Excel 打开…');
  try {
    const data = await api('/api/pdf_review/open', { pdf: pt.pdf, page: pt.page, x: pt.x, y: pt.y, old_folder: oldFolder });
    if (data.pending) {
      waiting = pt;
      toast('这张表还在定位，算好就自动打开…');
      startPolling();
      return;
    }
    try { if (oldFolder) localStorage.setItem(OLD_FOLDER_KEY, oldFolder); } catch (_) {}
    const where = data.cell ? `${data.sheet}!${data.cell}` : data.sheet;
    toast(data.note || (data.backup ? `已打开 ${where}：左边当前文件，右边原始备份（只读）` : `已在 Excel 中打开 ${where}`), !!data.note);
  } catch (e) {
    toast(`无法打开：${e.message}`, true);
  } finally {
    opening = false;
  }
}

function onPointerDown(e) {
  if (!$('prMenu').hidden && !e.target.closest('#prMenu')) closeMenu();
  const pageEl = e.target.closest('.pr-page');
  if (!pageEl || e.button !== 0 || pageEl.matches('.is-loading, .is-failed') || e.target.closest('.excel-link')) return;
  press = { pageEl, x: e.clientX, y: e.clientY };
}

function onPointerUp(e) {
  const start = press;
  press = null;
  if (!start || Math.hypot(e.clientX - start.x, e.clientY - start.y) > MIN_DRAG_PX) return;
  const pt = pointOn(start.pageEl, e);
  if (!pt.doc.wb_path) {
    toast('这份 PDF 找不到对应的 Excel', true);
    return;
  }
  flashAt(start.pageEl, e);
  // 单击换颜色（蓝 → 黄 → 红 → 蓝）；按住 Ctrl 单击、或右键菜单里才是用 Excel 打开这一格
  if (e.ctrlKey || e.metaKey) openAt(pt);
  else cyclePage(Number(start.pageEl.dataset.doc), pt.page, pt);
}

/* ---------- 三色标记 ----------
   黄色按页标：每页单独「要人工改」，单击标黄会记住点的那一格；红色是整张表（Excel 只能删整张表），
   标了红所有页都算红。标题颜色：有红 → 红；有一页黄 → 黄；所有页都是蓝 → 蓝。
   marks[pdf] = {v, red, pages: {页码: {x, y} 或 {}}, opened} */

function loadMarks() {
  marks = {};
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem(MARKS_KEY) || '{}')[folder] || {}; } catch (_) {}
  for (const doc of pdfs) {
    const m = saved[doc.pdf];
    // PDF 重新导出过（修改时间变了）就不算了：看到的已经不是标记时那一版
    if (!m || m.v !== doc.v || !doc.wb_path || doc.error || typeof m.pages !== 'object' || !m.pages) continue;
    const pages = {};
    for (const [p, spot] of Object.entries(m.pages)) if (Number(p) < doc.pages.length) pages[p] = spot || {};
    if (m.red || Object.keys(pages).length) marks[doc.pdf] = { v: m.v, red: !!m.red, pages, opened: !!m.opened };
  }
}

function saveMarks() {
  try {
    const all = JSON.parse(localStorage.getItem(MARKS_KEY) || '{}');
    if (Object.keys(marks).length) all[folder] = marks; else delete all[folder];
    localStorage.setItem(MARKS_KEY, JSON.stringify(all));
  } catch (_) {}
}

function colorOf(doc) {
  if (deleted.has(doc.pdf)) return 'deleted';
  const m = marks[doc.pdf];
  return m?.red ? 'red' : m && Object.keys(m.pages).length ? 'yellow' : 'blue';
}

function pageColor(doc, p) {
  const color = colorOf(doc);
  if (color === 'deleted' || color === 'red') return color;
  return marks[doc.pdf]?.pages[p] ? 'yellow' : 'blue';
}

// 只改颜色、角标和按钮状态；不改大小、不滚动
function paintMark(d) {
  const doc = pdfs[d];
  const article = $(`prDoc${d}`);
  if (!article) return;
  const color = colorOf(doc);
  const m = marks[doc.pdf];
  article.dataset.mark = color;
  article.classList.toggle('is-opened', !!m?.opened);
  article.querySelectorAll('[data-mark-btn]').forEach(b => {
    const on = b.dataset.markBtn === color;
    b.classList.toggle('is-on', on);
    b.setAttribute('aria-pressed', on);
    b.disabled = color === 'deleted' || processing;
  });
  article.querySelectorAll('.pr-page').forEach(pageEl => {
    const p = Number(pageEl.dataset.page);
    pageEl.dataset.mark = pageColor(doc, p);
    // 右键标黄时记下的位置：页面上放一个黄色小圈（浮层，不占位置）
    pageEl.querySelectorAll('.pr-pin').forEach(el => el.remove());
    const spot = color === 'yellow' ? m.pages[p] : null;
    if (spot && spot.x !== undefined) {
      const [w, h] = doc.pages[p];
      const pin = Object.assign(document.createElement('span'), { className: 'pr-pin', title: '标黄时点的位置，一键处理时 Excel 跳到这一格' });
      pin.style.left = `${spot.x / w * 100}%`;
      pin.style.top = `${spot.y / h * 100}%`;
      pageEl.querySelector('.pr-layer').appendChild(pin);
    }
  });
}

function editable(doc) {
  return doc && doc.wb_path && !doc.error && !deleted.has(doc.pdf) && !processing;
}

function commit(d) {
  const doc = pdfs[d];
  const m = marks[doc.pdf];
  if (m && !m.red && !Object.keys(m.pages).length) delete marks[doc.pdf];
  saveMarks();
  paintMark(d);
  updateProcessButton();
}

function entry(doc) {
  if (!marks[doc.pdf]) marks[doc.pdf] = { v: doc.v, red: false, pages: {}, opened: false };
  return marks[doc.pdf];
}

// 标题上的按钮：整张表一起改。蓝 = 全部清掉；黄 = 每页都标黄（右键记过的格子留着）；红 = 删除整张表
function setDocMark(d, color) {
  const doc = pdfs[d];
  if (!editable(doc)) return;
  if (color === 'blue') delete marks[doc.pdf];
  else {
    const m = entry(doc);
    m.red = color === 'red';
    if (color === 'yellow') doc.pages.forEach((_, p) => { m.pages[p] = m.pages[p] || {}; });
    m.opened = false;
  }
  commit(d);
}

// 单击 / 右键某一页：这一页标黄（记住右键的格子）或改回蓝；标过红的表这时就不删了
function setPageMark(d, page, color, point = null) {
  const doc = pdfs[d];
  if (!editable(doc)) return;
  const m = entry(doc);
  m.red = false;
  m.opened = false;
  if (color === 'yellow') m.pages[page] = point ? { x: point.x, y: point.y } : (m.pages[page] || {});
  else delete m.pages[page];
  commit(d);
}

// 单击这一页：蓝 → 黄（记住点的格子）→ 红（整张表）→ 这一页改回蓝（别的页标的黄还在）
function cyclePage(d, page, point) {
  const doc = pdfs[d];
  if (!editable(doc)) {
    if (processing) toast('正在一键处理，处理完再改颜色', true);
    return;
  }
  const now = pageColor(doc, page);
  if (now === 'blue') setPageMark(d, page, 'yellow', point);
  else if (now === 'yellow') setDocMark(d, 'red');
  else setPageMark(d, page, 'blue');
}

function marked(color) {
  return pdfs.filter(doc => colorOf(doc) === color);
}

function updateProcessButton() {
  const yellow = marked('yellow');
  const red = marked('red');
  const button = $('prProcess');
  $('prCountYellow').textContent = yellow.length;
  $('prCountRed').textContent = red.length;
  if (!processing) {
    $('prProcessText').textContent = '⚡ 一键处理';
    button.disabled = !yellow.length && !red.length;
  }
}

/* ---------- 右键菜单 ---------- */

let menuFor = null; // {d, page, point, pageEl, x, y}

function closeMenu() {
  $('prMenu').hidden = true;
  menuFor = null;
}

function onContextMenu(e) {
  const article = e.target.closest('.pr-doc');
  if (!article || e.target.closest('.excel-link, input')) return;
  const d = Number(article.dataset.doc);
  const doc = pdfs[d];
  if (!doc || !doc.wb_path || doc.error) return;
  e.preventDefault();
  const pageEl = e.target.closest('.pr-page');
  const ready = pageEl && !pageEl.matches('.is-loading, .is-failed');
  const point = ready ? pointOn(pageEl, e) : null;
  const page = pageEl ? Number(pageEl.dataset.page) : null;
  menuFor = { d, page, point, pageEl, x: e.clientX, y: e.clientY };
  const menu = $('prMenu');
  const now = page === null ? colorOf(doc) : pageColor(doc, page);
  const where = page === null ? '整张表' : `第 ${page + 1} 页`;
  menu.querySelector('.pr-menu-title').textContent = `${doc.wb_name ? doc.wb_name.split('/').pop() : doc.name} › ${doc.sheet || ''} · ${where}`;
  menu.querySelector('[data-menu="blue"] b').textContent = page === null ? '整张表不用管' : '这一页不用管';
  menu.querySelector('[data-menu="yellow"] b').textContent = page === null ? '整张表要人工改' : '这一页要人工改';
  menu.querySelector('[data-menu="yellow"] small').textContent = point ? '记住这一格，处理时 Excel 跳过来' : '处理时打开这张表';
  menu.querySelectorAll('[data-menu]').forEach(b => {
    const kind = b.dataset.menu;
    b.classList.toggle('is-on', kind === now);
    b.disabled = kind === 'open' ? !point : !editable(doc);
  });
  menu.hidden = false;
  // 放在鼠标旁边，靠窗口边缘时往里收（菜单是 fixed 浮层，不挪页面）
  const box = menu.getBoundingClientRect();
  menu.style.left = `${Math.min(e.clientX, window.innerWidth - box.width - 8)}px`;
  menu.style.top = `${Math.min(e.clientY, window.innerHeight - box.height - 8)}px`;
}

function onMenuClick(e) {
  const item = e.target.closest('[data-menu]');
  if (!item || !menuFor) return;
  const { d, page, point, pageEl, x, y } = menuFor;
  const kind = item.dataset.menu;
  closeMenu();
  if (kind === 'open') {
    if (point) {
      flashAt(pageEl, { clientX: x, clientY: y });
      openAt(point);
    }
  } else if (kind === 'red' || page === null) setDocMark(d, kind);
  else setPageMark(d, page, kind, point);
}

/* ---------- 一键处理 ---------- */

function listLines(docs, limit = 12) {
  const lines = docs.slice(0, limit).map(doc => {
    const pages = colorOf(doc) === 'yellow' ? Object.keys(marks[doc.pdf].pages).map(p => Number(p) + 1) : [];
    const which = pages.length && pages.length < doc.pages.length ? `（第 ${pages.join('、')} 页）` : '';
    return `  · ${doc.wb_name} › ${doc.sheet}${which}`;
  });
  if (docs.length > limit) lines.push(`  · ……另外 ${docs.length - limit} 张`);
  return lines.join('\n');
}

async function runProcess() {
  if (processing) return;
  const red = marked('red');
  let yellow = marked('yellow').filter(doc => !marks[doc.pdf].opened);
  if (!yellow.length && !red.length) yellow = marked('yellow'); // 都打开过了：再点一次就全部重新打开
  if (!red.length && !yellow.length) return;
  const files = docs => new Set(docs.map(doc => doc.wb_path)).size;
  const parts = [];
  if (red.length) {
    parts.push(`🔴 删除 ${red.length} 张表（${files(red)} 份 Excel），删完保存，对应的 PDF 也删掉：\n${listLines(red)}`);
  }
  if (yellow.length) {
    const n = files(yellow);
    parts.push(`🟡 用 Excel 打开 ${n} 份文件（${yellow.length} 张有标黄的表${n > 10 ? '，一次最多开 10 份，剩下的再点一次' : ''}）：\n${listLines(yellow)}`);
  }
  if (red.length) parts.push('删表前会先把别的表引用它的公式转成值；删表后出现新的引用错误时这份文件不保存。');
  const ok = await appConfirm(parts.join('\n\n'), {
    title: '一键处理', tone: red.length ? 'danger' : 'info', okText: red.length ? '删除并处理' : '打开',
  });
  if (!ok) return;
  const spots = doc => Object.entries(marks[doc.pdf].pages).map(([p, s]) => ({ page: Number(p), ...s }));
  try {
    await api('/api/pdf_review/process', {
      red: red.map(doc => ({ pdf: doc.pdf })), yellow: yellow.map(doc => ({ pdf: doc.pdf, spots: spots(doc) })),
    });
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
    return;
  }
  setProcessing(true);
  pollProcess();
}

function setProcessing(on) {
  processing = on;
  $('prProcess').disabled = on;
  $('prProcess').classList.toggle('is-running', on);
  pdfs.forEach((_, d) => paintMark(d));
  if (!on) updateProcessButton();
}

async function pollProcess() {
  let status;
  try {
    status = await api('/api/pdf_review/process');
  } catch (_) {
    setTimeout(pollProcess, 1500);
    return;
  }
  if (status.running) {
    $('prProcessText').textContent = status.stage === 'delete'
      ? `⏳ 删除中 ${status.done} / ${status.total}` : '⏳ 正在打开…';
    setTimeout(pollProcess, 1000);
    return;
  }
  setProcessing(false);
  if (status.result) finishProcess(status.result);
}

// 结果框：每一处都是链接，点一下 Excel 跳过去
function resultLine(item, { link = true } = {}) {
  const doc = pdfs.find(x => x.pdf === item.pdf);
  const page = item.page !== undefined && item.page !== null ? ` · 第 ${item.page + 1} 页` : '';
  const label = `${escapeHtml(item.name)} › ${escapeHtml(item.sheet)}${item.cell ? ` ${escapeHtml(item.cell)}` : ''}${page}`;
  const body = link && doc && doc.wb_name ? excelLinkHtml(label, excelTarget({ ...doc, sheet: item.sheet }, item.cell || '')) : label;
  const note = item.reason ? `<span class="pr-result-note">${escapeHtml(item.reason)}</span>` : '';
  return `<li>${body}${item.first ? ' <span class="pr-result-tag">已跳到这里</span>' : ''}${note}</li>`;
}

function finishProcess(result) {
  for (const item of result.deleted || []) {
    deleted.add(item.pdf);
    delete marks[item.pdf];
  }
  for (const item of result.opened || []) {
    if (marks[item.pdf]) marks[item.pdf].opened = true;
  }
  saveMarks();
  pdfs.forEach((_, d) => paintMark(d));
  updateProcessButton();
  const files = items => new Set(items.map(i => i.name)).size;
  const section = (title, items, opts) => items?.length
    ? `<h3>${title}</h3><ul>${items.map(i => resultLine(i, opts)).join('')}</ul>` : '';
  const html = [
    section(`🗑 删掉了 ${result.deleted?.length} 张表（已保存，PDF 也删了）`, result.deleted, { link: false }),
    section(`⚠ 没删掉 ${result.failed?.length} 张（还是红色，可以再点一次）`, result.failed),
    section(`📗 已用 Excel 打开 ${files(result.opened || [])} 份，标黄的地方（点一下跳过去）`, result.opened),
    section(`⚠ 打不开 ${result.open_failed?.length} 处`, result.open_failed),
    section(`⏸ 还有 ${files(result.later || [])} 份没打开（一次最多 10 份，改完关掉后再点「一键处理」）`, result.later),
  ].join('');
  const bad = result.failed?.length || result.open_failed?.length;
  toast(bad ? '一键处理完成，有没处理成功的，请看结果' : '一键处理完成', !!bad);
  $('prResultBody').innerHTML = html || '<p>没有要处理的</p>';
  $('prResult').dataset.tone = bad ? 'warn' : 'info';
  $('prResult').showModal();
}

/* ---------- 看到哪里（底部条） ---------- */

const READ_LINE = 0.4; // 窗口上方 40% 处那条线落在哪一页，就算正在看哪一页

function atBottom() {
  return window.innerHeight + window.scrollY >= document.documentElement.scrollHeight - 2;
}

// 进度条上每份 PDF 之间画一条分界线；份数多了线挤成一片竖条纹，看不清，
// 所以平均间距小于 TICK_MIN_GAP 像素就不画分界线，只用深色小段标出正在看的这一份
const TICK_MIN_GAP = 14;
let docSpans = [];       // 每份 PDF 在进度条上的 [起点, 终点]（0~1）

function renderTrack() {
  const total = pageEls.length;
  let before = 0;
  docSpans = pdfs.map(doc => {
    const start = before;
    before += doc.error ? 0 : doc.pages.length;
    return total ? [start / total, before / total] : [0, 0];
  });
  const ticks = docSpans.slice(0, -1)
    .map(([, end]) => `<span class="pr-tick" style="left:${end * 100}%"></span>`).join('');
  $('prTrack').innerHTML = `<div id="prFill" class="pr-fill"></div><div id="prCur" class="pr-cur"></div>${ticks}`;
  fitTicks();
  whereDoc = -1;
  updateWhere();
}

// 按进度条现在的宽度决定画不画分界线（窗口大小变了再判断一次）
function fitTicks() {
  const width = $('prTrack').getBoundingClientRect().width;
  $('prTrack').classList.toggle('is-dense', pdfs.length > 1 && width / pdfs.length < TICK_MIN_GAP);
}

// 读线所在的页（序号）和读线在这一页里的位置（0~1）；页面从上到下排，二分查找
function readingPosition() {
  // 滚到页面最底下时读线到不了最后一页的末尾，直接算看完了
  if (atBottom()) {
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
    if ($('prCur')) $('prCur').style.width = '0';
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
    const [start, end] = docSpans[d] || [0, 0];
    $('prCur').style.left = `${start * 100}%`;
    $('prCur').style.width = `${(end - start) * 100}%`;
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

/* ---------- 大图 / 中图 / 小图 ---------- */

// 换大小后停在原来看的那一页的同一位置
function setView(view, keep = true) {
  if (!VIEWS.includes(view)) view = 'large';
  const at = keep && pageEls.length ? readingPosition() : null;
  const box = $('prDocs');
  VIEWS.forEach(v => box.classList.toggle(`view-${v}`, v === view));
  document.querySelectorAll('#prView [data-view]').forEach(b => {
    b.classList.toggle('is-active', b.dataset.view === view);
    b.setAttribute('aria-pressed', b.dataset.view === view);
  });
  try { localStorage.setItem(VIEW_KEY, view); } catch (_) {}
  if (at) {
    const page = pageEls[at.index].getBoundingClientRect();
    window.scrollBy({ top: page.top + page.height * at.within - window.innerHeight * READ_LINE });
  }
  scheduleWhere();
}

/* ---------- 顶栏和底部条：自动隐藏 ---------- */
// 看 PDF 时顶栏和底部条会挡住内容：不碰它们 3 秒就藏起来（只往外滑，页面不挪）。
// 单击页面是打开 Excel，所以条只靠鼠标移到窗口最上 / 最下边缘叫出来。

const BARS_HIDE_MS = 3000;
const EDGE_TOP = 50;
const EDGE_BOTTOM = 90;
let barsTimer = 0;
let barsHover = false;

function barsHidden() {
  return document.body.classList.contains('pr-bars-hidden');
}

function setBars(show) {
  document.body.classList.toggle('pr-bars-hidden', !show);
  clearTimeout(barsTimer);
  if (show) barsTimer = setTimeout(autoHideBars, BARS_HIDE_MS);
}

function autoHideBars() {
  // 鼠标在条上、开着对话框、还在页面顶上（顶栏没挡住 PDF）时先不藏，过一会儿再看
  if (barsHover || document.querySelector('dialog[open]') || !pdfs.length || window.scrollY < 120) {
    barsTimer = setTimeout(autoHideBars, BARS_HIDE_MS);
    return;
  }
  setBars(false);
}

function watchBars() {
  for (const el of [document.querySelector('.app-header'), $('prBar')]) {
    el.addEventListener('mouseenter', () => { barsHover = true; if (barsHidden()) setBars(true); });
    el.addEventListener('mouseleave', () => { barsHover = false; setBars(true); });
  }
  document.addEventListener('mousemove', e => {
    if (!barsHidden()) return;
    if (e.clientY < EDGE_TOP || e.clientY > window.innerHeight - EDGE_BOTTOM) setBars(true);
  }, { passive: true });
  setBars(true);
}

/* ---------- 启动 ---------- */

// 页面图片加载完才能点：没加载出来时看不到点的是什么。load / error 不冒泡，在捕获阶段听
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
document.addEventListener('pointerup', onPointerUp);
document.addEventListener('pointercancel', () => { press = null; });
$('prLoad').addEventListener('click', load);
document.addEventListener('contextmenu', onContextMenu);
$('prMenu').addEventListener('click', onMenuClick);
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('prMenu').hidden) closeMenu(); });
window.addEventListener('scroll', () => { if (!$('prMenu').hidden) closeMenu(); }, { passive: true });
$('prDocs').addEventListener('click', e => {
  const button = e.target.closest('[data-mark-btn]');
  if (button) setDocMark(Number(button.dataset.doc), button.dataset.markBtn);
});
$('prProcess').addEventListener('click', runProcess);
$('prFolder').addEventListener('keydown', e => { if (e.key === 'Enter') load(); });
$('prTrack').addEventListener('click', jumpTo);
$('prView').addEventListener('click', e => {
  const button = e.target.closest('[data-view]');
  if (button) setView(button.dataset.view);
});
window.addEventListener('scroll', scheduleWhere, { passive: true });
window.addEventListener('resize', () => { if (pageEls.length) fitTicks(); scheduleWhere(); });
watchBars();
{
  let view = 'large';
  try { view = localStorage.getItem(VIEW_KEY) || view; } catch (_) {}
  setView(view, false);
}
try {
  $('prFolder').value = localStorage.getItem(FOLDER_KEY) || '';
  $('prOldFolder').value = localStorage.getItem(OLD_FOLDER_KEY) || '';
} catch (_) {}
if ($('prFolder').value) load();
// 刷新页面时一键处理还在后台跑：接着看进度
api('/api/pdf_review/process').then(status => { if (status.running) { setProcessing(true); pollProcess(); } }).catch(() => {});
