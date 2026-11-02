'use strict';
// PDF 人工复检：导出的 PDF 一页页用小图排开，快速过一遍（用户 2026-11-02 定：只用小图，不再切大 / 中图）。
// 三色标记（借鉴精简 Sheet）：蓝 = 不用管（默认）、黄 = 这一页要人工改（按页标）、红 = 整张表不要了。
// 单击页面换颜色（蓝 → 黄 → 红 → 蓝）。只记到「哪张表的第几页」，不记点的是哪一格（用户 2026-11-02 定）。
// 用 Excel 打开：Ctrl + 单击页面（或右键菜单）→ Excel 打开这张表、选中 PDF 这一页打印的范围；
// 填了原始备份文件夹时，备份也一起打开（只读）、跳到同一块，左右对照着改。
// 标题行右边的三个按钮一次改整张表。
// 看完后点底部按钮，黄红分开：「打开标黄」用 Excel 打开标黄的（一次最多 10 份 Excel），「删除标红」在后台删表并保存。
// 改完 Excel 想看效果：标题上「🔄 重新生成」只把这一份 Excel 的 PDF 重新导出一次。

const FOLDER_KEY = 'clear_handle_excel.lastFolder';
const OLD_FOLDER_KEY = 'clear_handle_excel.oldFolder'; // 和修复颜色 / 引用 / 隐藏页共用
const MARKS_KEY = 'clear_handle_excel.pdfReviewMarks'; // {文件夹: {pdf: {v, red, pages: {页码: {}}, opened}}}
const MARK_TEXT = { blue: '不用管', yellow: '要人工改', red: '删除这张表' };
const MARK_TITLE = { blue: '整张表不用管（清掉这张表的所有标记）', yellow: '整张表每页都标黄；只标某一页就单击那一页', red: '删除整张表（所有页）' };
const MIN_DRAG_PX = 6;  // 按下到松开挪动超过这么多就是在拖（选字、滚动），不算点
const OPEN_LIMIT = 10;   // 标黄的一次最多打开几份 Excel（和后台 pdf_review.OPEN_LIMIT 一样）

let folder = '';
let pdfs = [];          // [{pdf, name, v, pages: [[宽, 高]], wb_path, wb_name, sheet, guessed, error}]
let pageEls = [];       // 所有页面元素，从上到下
let whereDoc = -1;      // 底部条上现在显示的是哪一份（换了才重建链接）
let whereTick = 0;
let opening = false;
let press = null;       // 按下的位置，松开时判断是不是点了一下
let marks = {};         // pdf -> {v, red, pages: {页码: {}}, opened}；全蓝（默认）不记
let deleted = new Set(); // 这次一键处理删掉的表（PDF 也删了，页面上灰掉，刷新后就不在了）
let processing = false;
let regenWb = '';       // 正在重新生成 PDF 的那份 Excel（wb_path）

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
    deleted = new Set();
    loadMarks();
    renderDocs();
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
    const regen = doc.wb_path ? `<button type="button" class="pr-regen-btn" data-regen="${d}"
      title="改完 Excel 并保存后点这里：只把这一份 Excel 的 PDF 重新导出一次，马上看改完的效果">🔄 重新生成</button>` : '';
    return `<article class="pr-doc" id="prDoc${d}" data-doc="${d}" data-mark="blue">
      <header class="pr-doc-head" title="${escapeHtml(doc.pdf)}">
        <span class="pr-doc-name">📄 ${escapeHtml(doc.name)}</span>
        <span class="pr-doc-pages">${doc.pages.length} 页</span>${excel}
        <span class="pr-doc-gap"></span>${buttons}${regen}
      </header>
      ${doc.error ? `<p class="pr-doc-warn">${escapeHtml(doc.error)}</p>` : `<div class="pr-pages">${pages}</div>`}
    </article>`;
  }).join('');
  $('prBar').classList.remove('hidden');
  pageEls = [...box.querySelectorAll('.pr-page')];
  pdfs.forEach((_, d) => paintMark(d));
  paintRegen();
  renderTrack();
  updateStatus();
}

/* ---------- 份数 ---------- */

function updateStatus() {
  const pages = pdfs.reduce((n, doc) => n + doc.pages.length, 0);
  if (pdfs.length) $('prStatus').textContent = `共 ${pdfs.length} 份 PDF、${pages} 页`;
  $('prSummary').innerHTML = `共 <b>${pdfs.length}</b> 份 · ${pages} 页`;
}

/* ---------- 用 Excel 打开这一页 ---------- */

function flashPage(pageEl) {
  pageEl.classList.remove('is-flash');
  void pageEl.offsetWidth; // 重新触发动画
  pageEl.classList.add('is-flash');
}

async function openPage(d, page, pageEl = null) {
  const doc = pdfs[d];
  if (!doc || !doc.wb_path) {
    toast('这份 PDF 找不到对应的 Excel', true);
    return;
  }
  if (opening) return;
  if (pageEl) flashPage(pageEl);
  const oldFolder = $('prOldFolder').value.trim();
  opening = true;
  toast(oldFolder ? '正在打开当前文件和备份…' : `正在用 Excel 打开第 ${page + 1} 页…`);
  try {
    const data = await api('/api/pdf_review/open', { pdf: doc.pdf, page, pages: doc.pages.length, old_folder: oldFolder });
    try { if (oldFolder) localStorage.setItem(OLD_FOLDER_KEY, oldFolder); } catch (_) {}
    const where = `${data.sheet} 第 ${page + 1} 页${data.cell ? `（${data.cell}）` : ''}`;
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
  if (!pageEl || e.button !== 0 || pageEl.matches('.is-loading, .is-failed')
      || e.target.closest('.excel-link')) return;
  press = { pageEl, x: e.clientX, y: e.clientY };
}

function onPointerUp(e) {
  const start = press;
  press = null;
  if (!start || Math.hypot(e.clientX - start.x, e.clientY - start.y) > MIN_DRAG_PX) return;
  const d = Number(start.pageEl.dataset.doc);
  const page = Number(start.pageEl.dataset.page);
  if (!pdfs[d].wb_path) {
    toast('这份 PDF 找不到对应的 Excel', true);
    return;
  }
  // 单击换颜色（蓝 → 黄 → 红 → 蓝）；按住 Ctrl 单击是用 Excel 打开这一页
  if (e.ctrlKey || e.metaKey) openPage(d, page, start.pageEl);
  else cyclePage(d, page);
}

/* ---------- 三色标记 ----------
   黄色按页标：每页单独「要人工改」（只记第几页）；红色是整张表（Excel 只能删整张表），
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

// 单击 / 右键某一页：这一页标黄或改回蓝；标过红的表这时就不删了
function setPageMark(d, page, color) {
  const doc = pdfs[d];
  if (!editable(doc)) return;
  const m = entry(doc);
  m.red = false;
  m.opened = false;
  if (color === 'yellow') m.pages[page] = {};
  else delete m.pages[page];
  commit(d);
}

// 单击这一页：蓝 → 黄 → 红（整张表）→ 这一页改回蓝（别的页标的黄还在）
function cyclePage(d, page) {
  const doc = pdfs[d];
  if (!editable(doc)) {
    if (processing) toast('正在处理，处理完再改颜色', true);
    return;
  }
  const now = pageColor(doc, page);
  if (now === 'blue') setPageMark(d, page, 'yellow');
  else if (now === 'yellow') setDocMark(d, 'red');
  else setPageMark(d, page, 'blue');
}

function marked(color) {
  return pdfs.filter(doc => colorOf(doc) === color);
}

const PROCESS_TEXT = { yellow: '📗 打开标黄', red: '🗑 删除标红' };
const PROCESS_BUTTON = { yellow: 'prOpenYellow', red: 'prDeleteRed' };
let processKind = '';  // 正在处理的是哪个按钮：yellow / red

// 标黄的 Excel 文件：一共几个、这一轮已经打开了几个（同一份 Excel 几张标黄的表算 1 个）
function yellowFiles() {
  const all = new Map();
  for (const doc of marked('yellow')) {
    if (!all.has(doc.wb_path)) all.set(doc.wb_path, false);
    if (marks[doc.pdf].opened) all.set(doc.wb_path, true);
  }
  const opened = [...all.values()].filter(Boolean).length;
  return { total: all.size, opened };
}

function updateProcessButton() {
  const counts = { yellow: marked('yellow').length, red: marked('red').length };
  const files = yellowFiles();
  // 和切字工作台「打开下一批」一样写明进度：已开几个 / 共几个
  $('prCountYellow').textContent = files.total ? `${files.opened} / ${files.total}` : '0';
  $('prCountYellow').parentElement.title = files.total
    ? `标黄的 Excel 共 ${files.total} 个，这一轮已经打开 ${files.opened} 个（一次最多开 ${OPEN_LIMIT} 个）`
    : '没有标黄的表';
  $('prCountRed').textContent = counts.red;
  if (processing) return;
  const yellowText = files.total && files.opened >= files.total ? '📗 再开一遍标黄'
    : files.opened ? '📗 打开下一批标黄' : PROCESS_TEXT.yellow;
  document.querySelector('[data-process-text="yellow"]').textContent = yellowText;
  document.querySelector('[data-process-text="red"]').textContent = PROCESS_TEXT.red;
  $('prOpenYellow').title = files.total
    ? `标黄的 Excel 共 ${files.total} 个，已经打开 ${files.opened} 个，还剩 ${files.total - files.opened} 个；一次最多开 ${OPEN_LIMIT} 个`
    : '用 Excel 打开标黄的文件，跳到标黄的那一页（一次最多 10 个）';
  for (const kind of ['yellow', 'red']) $(PROCESS_BUTTON[kind]).disabled = !counts[kind] || !!regenWb;
}

/* ---------- 右键菜单 ---------- */

let menuFor = null; // {d, page, pageEl}

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
  const page = pageEl ? Number(pageEl.dataset.page) : null;
  menuFor = { d, page, pageEl };
  const menu = $('prMenu');
  const now = page === null ? colorOf(doc) : pageColor(doc, page);
  const where = page === null ? '整张表' : `第 ${page + 1} 页`;
  menu.querySelector('.pr-menu-title').textContent = `${doc.wb_name ? doc.wb_name.split('/').pop() : doc.name} › ${doc.sheet || ''} · ${where}`;
  menu.querySelector('[data-menu="blue"] b').textContent = page === null ? '整张表不用管' : '这一页不用管';
  menu.querySelector('[data-menu="yellow"] b').textContent = page === null ? '整张表要人工改' : '这一页要人工改';
  menu.querySelector('[data-menu="yellow"] small').textContent = page === null ? '处理时打开这张表' : '「打开标黄」时 Excel 跳到这一页';
  menu.querySelector('[data-menu="open"] b').textContent = page === null ? '在 Excel 打开这张表' : `在 Excel 打开第 ${page + 1} 页`;
  menu.querySelectorAll('[data-menu]').forEach(b => {
    const kind = b.dataset.menu;
    b.classList.toggle('is-on', kind === now);
    b.disabled = kind === 'open' ? false : !editable(doc);
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
  const { d, page, pageEl } = menuFor;
  const kind = item.dataset.menu;
  closeMenu();
  if (kind === 'open') openPage(d, page ?? 0, pageEl);
  else if (kind === 'red' || page === null) setDocMark(d, kind);
  else setPageMark(d, page, kind);
}

/* ---------- 打开标黄 / 删除标红 ---------- */

function listLines(docs, limit = 12) {
  const lines = docs.slice(0, limit).map(doc => {
    const pages = colorOf(doc) === 'yellow' ? Object.keys(marks[doc.pdf].pages).map(p => Number(p) + 1) : [];
    const which = pages.length && pages.length < doc.pages.length ? `（第 ${pages.join('、')} 页）` : '';
    return `  · ${doc.wb_name} › ${doc.sheet}${which}`;
  });
  if (docs.length > limit) lines.push(`  · ……另外 ${docs.length - limit} 张`);
  return lines.join('\n');
}

// 黄红分开（用户 2026-11-02 定）：kind = 'yellow' 只打开标黄的，'red' 只删标红的
async function runProcess(kind) {
  if (processing || regenWb) return;
  const files = docs => new Set(docs.map(doc => doc.wb_path)).size;
  let red = [];
  let yellow = [];
  let text;
  if (kind === 'red') {
    red = marked('red');
    if (!red.length) return;
    text = `🔴 删除 ${red.length} 张表（${files(red)} 份 Excel），删完保存，对应的 PDF 也删掉：\n${listLines(red)}`
      + '\n\n删表前会先把别的表引用它的公式转成值；删表后出现新的引用错误时这份文件不保存。';
  } else {
    yellow = marked('yellow').filter(doc => !marks[doc.pdf].opened);
    if (!yellow.length) yellow = marked('yellow'); // 都打开过了：再点一次就全部重新打开
    if (!yellow.length) return;
    const n = files(yellow);
    const { total, opened } = yellowFiles();
    // 按 Excel 文件算：同一份 Excel 有几张标黄的表也只打开一次
    const now = Math.min(n, OPEN_LIMIT);
    const before = opened >= total ? 0 : opened; // 全开过了再点：从头再开一遍
    text = `🟡 这次用 Excel 打开 ${now} 个文件（标黄的共 ${total} 个，打开后已开 ${before + now} / ${total} 个`
      + `${before + now < total ? `，还剩 ${total - before - now} 个，改完关掉后再点一次` : ''}）。\n`
      + `这次的 ${now} 个文件：\n${listLines(yellow.filter(doc => [...new Set(yellow.map(x => x.wb_path))].slice(0, OPEN_LIMIT).includes(doc.wb_path)))}`;
  }
  const ok = await appConfirm(text, kind === 'red'
    ? { title: '删除标红的表', tone: 'danger', okText: '删除' }
    : { title: '打开标黄的文件', tone: 'info', okText: '打开' });
  if (!ok) return;
  const spots = doc => Object.keys(marks[doc.pdf].pages).map(p => ({ page: Number(p) }));
  try {
    await api('/api/pdf_review/process', {
      red: red.map(doc => ({ pdf: doc.pdf })), yellow: yellow.map(doc => ({ pdf: doc.pdf, spots: spots(doc) })),
    });
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
    return;
  }
  processKind = kind;
  setProcessing(true);
  pollProcess();
}

function setProcessing(on) {
  processing = on;
  for (const kind of ['yellow', 'red']) {
    $(PROCESS_BUTTON[kind]).disabled = on;
    $(PROCESS_BUTTON[kind]).classList.toggle('is-running', on && processKind === kind);
  }
  pdfs.forEach((_, d) => paintMark(d));
  paintRegen();
  if (!on) {
    processKind = '';
    updateProcessButton();
  }
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
    if (!processKind) processKind = status.stage === 'delete' ? 'red' : 'yellow';
    const label = document.querySelector(`[data-process-text="${processKind}"]`);
    label.textContent = status.stage === 'delete' ? `⏳ 删除中 ${status.done} / ${status.total}` : '⏳ 正在打开…';
    $(PROCESS_BUTTON[processKind]).classList.add('is-running');
    setTimeout(pollProcess, 1000);
    return;
  }
  setProcessing(false);
  if (status.result) finishProcess(status.result);
}

// 结果框：每一处都是链接，点一下 Excel 跳过去
function resultLine(item, { link = true } = {}) {
  const d = pdfs.findIndex(x => x.pdf === item.pdf);
  const hasPage = item.page !== undefined && item.page !== null;
  const label = `${escapeHtml(item.name)} › ${escapeHtml(item.sheet)}${hasPage ? ` · 第 ${item.page + 1} 页` : ''}`;
  const body = link && d >= 0 && pdfs[d].wb_path
    ? `<button type="button" class="pr-result-open" data-result-doc="${d}" data-result-page="${hasPage ? item.page : 0}" title="在 Excel 打开这一页">📗 ${label}</button>`
    : label;
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
    section(`📗 这次打开了 ${files(result.opened || [])} 个文件（标黄的共 ${yellowFiles().total} 个，已开 ${yellowFiles().opened} 个）。标黄的地方点一下跳过去`, result.opened),
    section(`⚠ 打不开 ${result.open_failed?.length} 处`, result.open_failed),
    section(`⏸ 还剩 ${files(result.later || [])} 个没打开（一次最多 ${OPEN_LIMIT} 个，改完关掉后再点「📗 打开下一批标黄」）`, result.later),
  ].join('');
  const bad = result.failed?.length || result.open_failed?.length;
  const yf = yellowFiles();
  const openedNote = result.opened?.length
    ? `已打开 ${files(result.opened)} 个，标黄共 ${yf.total} 个，已开 ${yf.opened} 个${yf.total > yf.opened ? `，还剩 ${yf.total - yf.opened} 个` : '，全部开过了'}`
    : '';
  toast(bad ? '处理完成，有没处理成功的，请看结果' : openedNote || '处理完成', !!bad);
  $('prResultBody').innerHTML = html || '<p>没有要处理的</p>';
  $('prResult').dataset.tone = bad ? 'warn' : 'info';
  $('prResult').showModal();
}

/* ---------- 单份 Excel 重新生成 PDF ---------- */
// 复检时在 Excel 里改了、保存了，想马上看改完的效果：只把这一份 Excel 的 PDF 重新导出一次。
// 做完重新读一遍列表；页面停在原来那份上（按标题位置对齐），不跳走。

function paintRegen() {
  document.querySelectorAll('[data-regen]').forEach(b => {
    const doc = pdfs[Number(b.dataset.regen)];
    const mine = regenWb && doc && doc.wb_path === regenWb;
    b.textContent = mine ? '⏳ 生成中…' : '🔄 重新生成';
    b.classList.toggle('is-running', !!mine);
    b.disabled = !!regenWb || processing || deleted.has(doc?.pdf);
  });
  updateProcessButton(); // 重新生成时两个处理按钮也先不能点
}

async function regenerate(d) {
  const doc = pdfs[d];
  if (!doc || !doc.wb_path || regenWb || processing) return;
  const siblings = pdfs.filter(x => x.wb_path === doc.wb_path);
  if (siblings.some(x => marks[x.pdf])) {
    const ok = await appConfirm(
      `重新生成「${doc.wb_name}」的 PDF（${siblings.length} 张表）？\n\n这份 Excel 在 PDF 上标的黄色 / 红色会清掉：新的 PDF 是改过的一版，标记对不上了。`,
      { title: '重新生成 PDF', tone: 'warn', okText: '重新生成' });
    if (!ok) return;
  }
  let data;
  try {
    data = await api('/api/pdf_review/regenerate', { pdf: doc.pdf });
    if (data.unchanged) {
      const again = await appConfirm(
        `「${data.name}」和上次导出 PDF 时一模一样（哈希值没变），PDF 已经是最新的。\n\n刚在 Excel 里改过的话，多半还没保存：先在 Excel 里保存，再点「重新生成」。\n\n还是要照样重新生成一次吗？`,
        { title: 'Excel 没有改过', tone: 'warn', okText: '照样重新生成' });
      if (!again) return;
      data = await api('/api/pdf_review/regenerate', { pdf: doc.pdf, force: true });
    }
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
    return;
  }
  regenWb = doc.wb_path;
  paintRegen();
  toast(`正在重新生成「${doc.wb_name}」的 PDF…`);
  pollRegen(doc.pdf);
}

async function pollRegen(anchorPdf) {
  let status;
  try {
    status = await api('/api/pdf_review/regenerate');
  } catch (_) {
    setTimeout(() => pollRegen(anchorPdf), 1500);
    return;
  }
  if (status.running) {
    regenWb = status.wb_path || regenWb;
    paintRegen();
    setTimeout(() => pollRegen(anchorPdf), 1000);
    return;
  }
  const result = status.result;
  const wb = regenWb;
  regenWb = '';
  if (!result) { paintRegen(); return; }
  if (!result.ok) {
    paintRegen();
    toast(`「${result.name}」没能重新生成`, true);
    appAlert(`「${result.name}」没能重新生成 PDF：\n${result.reason || '未知原因'}\n\n原来的 PDF 可能已经删掉了，修好后再点一次「重新生成」，或者回主页重新导出。`, { tone: 'error' });
  }
  await reloadKeeping(anchorPdf || status.pdf, wb || result.wb_path);
  if (result.ok) toast(`已重新生成「${result.name}」的 ${result.pdfs.length} 份 PDF`);
}

// 重新读列表，页面停在原来那份 Excel 的标题上（同一个 PDF 还在就对齐它，不在了对齐同一份 Excel 的第一张）
async function reloadKeeping(pdf, wbPath) {
  const before = pdfs.findIndex(x => x.pdf === pdf);
  const old = before >= 0 ? $(`prDoc${before}`) : null;
  const top = old ? old.getBoundingClientRect().top : null;
  let data;
  try {
    data = await api(`/api/pdf_review?folder=${encodeURIComponent(folder)}`);
  } catch (e) {
    appAlert(`重新读取 PDF 列表失败：${e.message}`, { tone: 'error' });
    return;
  }
  pdfs = data.pdfs;
  deleted = new Set([...deleted].filter(p => pdfs.some(x => x.pdf === p)));
  loadMarks(); // 重新生成的 PDF 修改时间变了，它的标记自动不算了
  saveMarks();
  renderDocs();
  updateProcessButton();
  let after = pdfs.findIndex(x => x.pdf === pdf);
  if (after < 0) after = pdfs.findIndex(x => x.wb_path === wbPath);
  const now = after >= 0 ? $(`prDoc${after}`) : null;
  if (now && top !== null) window.scrollBy(0, now.getBoundingClientRect().top - top);
  scheduleWhere();
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
// 只用小图（用户 2026-11-02 定）：一行放好几页，快速过一遍
function setView() {
  $('prDocs').classList.add('view-small');
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
  const regen = e.target.closest('[data-regen]');
  if (regen) regenerate(Number(regen.dataset.regen));

});
$('prOpenYellow').addEventListener('click', () => runProcess('yellow'));
$('prDeleteRed').addEventListener('click', () => runProcess('red'));
$('prFolder').addEventListener('keydown', e => { if (e.key === 'Enter') load(); });
$('prTrack').addEventListener('click', jumpTo);
$('prResultBody').addEventListener('click', e => {
  const button = e.target.closest('[data-result-doc]');
  if (button) openPage(Number(button.dataset.resultDoc), Number(button.dataset.resultPage));
});
window.addEventListener('scroll', scheduleWhere, { passive: true });
window.addEventListener('resize', () => { if (pageEls.length) fitTicks(); scheduleWhere(); });
watchBars();
setView();
try {
  $('prFolder').value = localStorage.getItem(FOLDER_KEY) || '';
  $('prOldFolder').value = localStorage.getItem(OLD_FOLDER_KEY) || '';
} catch (_) {}
if ($('prFolder').value) load();
// 刷新页面时一键处理还在后台跑：接着看进度
api('/api/pdf_review/process').then(status => { if (status.running) { setProcessing(true); pollProcess(); } }).catch(() => {});
// 重新生成还在后台跑：接着看，做完刷新列表
api('/api/pdf_review/regenerate').then(status => {
  if (status.running) { regenWb = status.wb_path || ''; paintRegen(); pollRegen(status.pdf); }
}).catch(() => {});
