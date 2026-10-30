'use strict';
// PDF 人工复检：导出的 PDF 整页显示，看到问题点一下，就用 Excel 打开并选中那一格；
// 填了原始备份文件夹时，备份也一起打开（只读）、跳到同一格，左右对照着改。
// 每份 PDF 的行列位置由后台算（只算一次），点的格子也由后台按它算。

const FOLDER_KEY = 'clear_handle_excel.lastFolder';
const OLD_FOLDER_KEY = 'clear_handle_excel.oldFolder'; // 和修复颜色 / 引用 / 隐藏页共用
const VIEW_KEY = 'clear_handle_excel.pdfReviewView';
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
      : '<span class="pr-doc-warn">⚠ 找不到对应的 Excel，点了也打不开</span>';
    const pages = doc.pages.map(([w, h], p) => `
      <div class="pr-page is-loading" data-doc="${d}" data-page="${p}" style="aspect-ratio:${w} / ${h};--ar:${w / h}">
        <img class="pr-img" loading="lazy" decoding="async" draggable="false" alt="第 ${p + 1} 页"
          src="/api/pdf_review/page?pdf=${encodeURIComponent(doc.pdf)}&p=${p}&v=${doc.v}">
        <div class="pr-layer"></div>
        <span class="pr-pageno">${p + 1} / ${doc.pages.length}</span>
      </div>`).join('');
    return `<article class="pr-doc" id="prDoc${d}">
      <header class="pr-doc-head" title="${escapeHtml(doc.pdf)}">
        <span class="pr-doc-name">📄 ${escapeHtml(doc.name)}</span>
        <span class="pr-doc-pages">${doc.pages.length} 页</span>${excel}
        <span class="pr-doc-state" data-state="${d}"></span>
      </header>
      ${doc.error ? `<p class="pr-doc-warn">${escapeHtml(doc.error)}</p>` : `<div class="pr-pages">${pages}</div>`}
    </article>`;
  }).join('');
  $('prBar').classList.remove('hidden');
  pageEls = [...box.querySelectorAll('.pr-page')];
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
  openAt(pt);
}

/* ---------- 看到哪里（底部条） ---------- */

const READ_LINE = 0.4; // 窗口上方 40% 处那条线落在哪一页，就算正在看哪一页

function atBottom() {
  return window.innerHeight + window.scrollY >= document.documentElement.scrollHeight - 2;
}

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
$('prFolder').addEventListener('keydown', e => { if (e.key === 'Enter') load(); });
$('prTrack').addEventListener('click', jumpTo);
$('prView').addEventListener('click', e => {
  const button = e.target.closest('[data-view]');
  if (button) setView(button.dataset.view);
});
window.addEventListener('scroll', scheduleWhere, { passive: true });
window.addEventListener('resize', scheduleWhere);
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
