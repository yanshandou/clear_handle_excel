'use strict';
// 调整页面（用户 2026-11-07 定，原「PDF 人工复检」页改成工具箱）：人不再手动改 Excel 的打印设置，
// 在网页上点按钮告诉程序怎么改，底部「🛠 统一修改」后台改 Excel、重新导出 PDF，看新的 PDF 再决定，
// 可以来回好几轮，页面都没问题了再去下一步（查看切字：扫描标色、调整单元格）。
// - 导出 PDF 时每张表已经设成「宽一页、高自动」（横向不许拆页，纵向交给 Excel 自动分页）。
// - 标题上每张表三组按钮：方向（竖向 / 横向）、页高（高自动 / ⤓ 取整 N→N-1）、删除这张表。
// - 末页只印了一点（不到 20%，用户 2026-11-08 定）的表自动标黄 = 建议向下取整；不想取整再点一下取消。
// - 改了方向后要不要取整就不一定了：先改方向，看新的 PDF 再决定（同一张表方向和取整不能一起改）。
// - 按问题分类：上面的分类按钮只看某一类（建议取整 / 竖向 / 横向 / 页高定死的），并能一次改这一类。
// 要自己在 Excel 里改的：点标题上的 Excel 路径，或 Ctrl + 单击页面，Excel 打开并跳到那一页；
// 改完保存后点「🔄 重新生成」。

const FOLDER_KEY = 'clear_handle_excel.lastFolder';
const OLD_FOLDER_KEY = 'clear_handle_excel.oldFolder';
const VIEW_KEY = 'clear_handle_excel.pdfReviewView';
const VIEWS = ['large', 'medium', 'small']; // 和修复颜色 / 引用 / 隐藏页共用
const MARKS_KEY = 'clear_handle_excel.pdfReviewMarks'; // {文件夹: {pdf: {v, red}}}：删除这张表
const TOOLS_KEY = 'clear_handle_excel.pdfReviewTools'; // {文件夹: {pdf: {v, floor | auto | orient | breaks}}}：工具箱标记
const AUTO_KEY = 'clear_handle_excel.pdfReviewAutoFloor'; // {文件夹: {pdf: v}}：这一版 PDF 已经自动标过「建议取整」
const CAT_KEY = 'clear_handle_excel.pdfReviewCategory';
const MIN_DRAG_PX = 6;  // 按下到松开挪动超过这么多就是在拖（选字、滚动），不算点
const ORIENT_TEXT = { 1: '竖向', 2: '横向' };

let folder = '';
let allPdfs = [];       // 文件夹里所有的 PDF：[{pdf, name, v, pages: [[宽, 高]], wb_path, wb_name, sheet, layout, ...}]
let pdfs = [];          // 当前分类里的（页面上显示的）
let category = 'all';
let pageEls = [];       // 所有页面元素，从上到下
let whereDoc = -1;      // 底部条上现在显示的是哪一份（换了才重建链接）
let whereTick = 0;
let opening = false;
let press = null;       // 按下的位置，松开时判断是不是点了一下
let marks = {};         // pdf -> {v, red: true}：删除这张表
let tools = {};         // pdf -> {v, floor: true} | {v, auto: true} | {v, orient: 1|2} | {v, breaks: true}，可以 orient + auto/breaks
let deleted = new Set(); // 这次一键处理删掉的表（PDF 也删了，页面上灰掉，刷新后就不在了）
let processing = false;
let regenWb = '';       // 正在重新生成 PDF 的那份 Excel（wb_path）
let busyWb = '';        // 点了「重新生成」、还在等的那份 Excel（含先改标记再生成）：页面上显示「正在重新生成」
let busySince = 0;      // 从什么时候开始等，按钮上显示已经等了几秒
let busyTimer = 0;

const $ = id => document.getElementById(id);

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, ch => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

async function api(url, body) {
  let res;
  try {
    res = await fetch(url, body === undefined ? {} : {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
    });
  } catch (_) {
    // 浏览器只说「Failed to fetch」：其实是连不上后台程序（黑色窗口关了、程序重启中或崩了）
    throw new Error('连不上后台程序：请看一下运行本工具的黑色窗口还开着没有（关了就重新双击启动），窗口里有红字报错的话截图发过来。程序开着的话，刷新网页再试一次。');
  }
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

/* ---------- 每张表的打印设置 ---------- */

// 方向：导出记录里有就用记录的，没有（旧版本导出的）看第一页是横的还是竖的
function orientOf(doc) {
  if (doc.layout?.orient === 1 || doc.layout?.orient === 2) return doc.layout.orient;
  const [w, h] = doc.pages[0] || [1, 2];
  return w > h ? 2 : 1;
}

function tallOf(doc) {
  return Number.isInteger(doc.layout?.tall) ? doc.layout.tall : null;
}

// 有几处手动分页：页高定死（取整过）时 Excel 不管手动分页，不算
function manualOf(doc) {
  const n = doc.layout?.manual;
  return Number.isInteger(n) && n > 0 && tallOf(doc) === null ? n : 0;
}

const FLOOR_HINT_PCT = 20; // 末页印了不到 20% 才建议取整（用户 2026-11-08 定；和后台 page_fill.FLOOR_HINT_FILL 一样）

// 按记下的末页占比判断，以前按 40% 导出的记录也按新标准算，不用重新导出
function suggestsFloor(doc) {
  if (doc.pages.length < 2) return false;
  const tail = doc.layout?.tail;
  return typeof tail === 'number' ? tail < FLOOR_HINT_PCT : !!doc.layout?.floor;
}

/* ---------- 分类 ---------- */

const CATEGORIES = [
  { key: 'all', label: '全部', test: () => true },
  { key: 'floor', label: '🟡 末页只多一点', hint: '末页只印了不到 20%，建议向下取整（已自动标黄）', test: suggestsFloor },
  { key: 'portrait', label: '竖向', test: doc => orientOf(doc) === 1 },
  { key: 'landscape', label: '横向', test: doc => orientOf(doc) === 2 },
  { key: 'fixed', label: '页高定死的', hint: '取整过、或者原来就设了「N 页高」的表', test: doc => tallOf(doc) !== null },
  { key: 'manual', label: '有手动分页', hint: '页高自动、但插了手动分页的表：清掉以后由 Excel 自动分页，常常能排得更满', test: doc => manualOf(doc) > 0 },
];

function catOf(key) {
  return CATEGORIES.find(c => c.key === key) || CATEGORIES[0];
}

function renderCats() {
  const usable = allPdfs.filter(doc => doc.wb_path && !doc.error);
  $('prCats').hidden = !allPdfs.length;
  $('prCats').innerHTML = CATEGORIES.map(c => {
    const n = c.key === 'all' ? allPdfs.length : usable.filter(c.test).length;
    return `<button type="button" class="pr-cat${c.key === category ? ' is-on' : ''}" data-cat="${c.key}"
      ${n || c.key === 'all' ? '' : 'disabled'} title="${escapeHtml(c.hint || `只看${c.label}的表`)}" aria-pressed="${c.key === category}">
      ${c.label}<span class="pr-cat-n">${n}</span></button>`;
  }).join('') + '<span class="pr-cat-gap"></span><button type="button" id="prBulk" class="btn btn-outline btn-sm pr-bulk" hidden></button>';
  paintBulk();
}

// 这一类一起改：按钮上写着这次会做什么
function bulkPlan() {
  const docs = pdfs.filter(doc => editable(doc) && colorOf(doc) !== 'red');
  if (category === 'floor') {
    const on = docs.filter(doc => tools[doc.pdf]?.floor);
    return on.length === docs.length
      ? { text: `这一类全部不取整（${docs.length}）`, run: () => docs.forEach(doc => setTool(doc, 'floor', false)) }
      : { text: `这一类全部取整（${docs.length}）`, run: () => docs.forEach(doc => setTool(doc, 'floor', true)) };
  }
  if (category === 'portrait' || category === 'landscape') {
    const target = category === 'portrait' ? 2 : 1;
    return { text: `这一类全部改成${ORIENT_TEXT[target]}（${docs.length}）`, run: () => docs.forEach(doc => setOrient(doc, target)) };
  }
  if (category === 'fixed') {
    return { text: `这一类页高全部改回自动（${docs.length}）`, run: () => docs.forEach(doc => setTool(doc, 'auto', true)) };
  }
  if (category === 'manual') {
    const on = docs.filter(doc => tools[doc.pdf]?.breaks);
    return on.length === docs.length
      ? { text: `这一类全部不清手动分页（${docs.length}）`, run: () => docs.forEach(doc => setTool(doc, 'breaks', false)) }
      : { text: `这一类全部清手动分页（${docs.length}）`, run: () => docs.forEach(doc => setTool(doc, 'breaks', true)) };
  }
  return null;
}

function paintBulk() {
  const button = $('prBulk');
  if (!button) return;
  const plan = category === 'all' ? null : bulkPlan();
  button.hidden = !plan;
  if (plan) {
    button.textContent = plan.text;
    button.disabled = processing || !!regenWb;
  }
}

function runBulk() {
  const plan = bulkPlan();
  if (!plan) return;
  plan.run();
  saveTools();
  pdfs.forEach((_, d) => paintMark(d));
  updateProcessButton();
}

function setCategory(key) {
  category = catOf(key).key;
  try { localStorage.setItem(CAT_KEY, category); } catch (_) {}
  applyCategory();
  window.scrollTo({ top: 0 });
}

function applyCategory() {
  const cat = catOf(category);
  pdfs = category === 'all' ? allPdfs : allPdfs.filter(doc => doc.wb_path && !doc.error && cat.test(doc));
  renderCats();
  renderDocs();
  updateProcessButton();
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
    allPdfs = data.pdfs;
    deleted = new Set();
    loadMarks();
    applyCategory();
  } catch (e) {
    $('prStatus').textContent = `读取失败：${e.message}`;
    appAlert(e.message, { tone: 'error' });
  } finally {
    $('prLoad').disabled = false;
  }
}

function headButtons(doc, d) {
  const orient = orientOf(doc);
  const tall = tallOf(doc);
  const n = doc.pages.length;
  const orientGroup = `<span class="pr-mark pr-set" role="group" aria-label="打印方向">${[1, 2].map(o =>
    `<button type="button" class="pr-mark-btn" data-orient="${o}" data-doc="${d}"
      title="${o === orient ? `现在是${ORIENT_TEXT[o]}` : `改成${ORIENT_TEXT[o]}（点「统一修改」后改 Excel、重新生成 PDF；再点一次取消）`}">${ORIENT_TEXT[o]}</button>`).join('')}</span>`;
  const autoTitle = tall === null ? '现在是宽一页、高自动（Excel 自己定几页）'
    : `现在定死了 ${tall} 页高；点这里改回自动（Excel 自己定几页）`;
  const floorTitle = n > 1
    ? `最后一页只多一点时点这里：「统一修改」时挤成 ${n - 1} 页（宽一页、${n - 1} 页高；字缩到不到 85%、缩出切字就不改）`
    : '只有 1 页，不能再少';
  const tallGroup = `<span class="pr-mark pr-set" role="group" aria-label="页高">
    <button type="button" class="pr-mark-btn" data-tall="auto" data-doc="${d}" title="${autoTitle}">${tall === null ? '高自动' : `定死 ${tall} 页 → 自动`}</button>
    <button type="button" class="pr-mark-btn" data-tall="floor" data-doc="${d}" ${n > 1 ? '' : 'disabled'} title="${floorTitle}">⤓ 取整 ${n}→${Math.max(1, n - 1)} 页</button></span>`;
  const manual = manualOf(doc);
  const breaks = manual ? `<span class="pr-mark pr-set" role="group" aria-label="手动分页">
    <button type="button" class="pr-mark-btn" data-breaks="${d}"
      title="这张表有 ${manual} 处手动分页：点这里，「统一修改」时全部清掉、改成 Excel 自动分页，保存后重新生成 PDF 给你看（末页只多一点的照常建议取整）；再点一次取消">✂ 清手动分页（${manual} 处）</button></span>` : '';
  const red = `<button type="button" class="pr-red-btn" data-red="${d}" title="整张表不要了：点「🗑 删除标红」时在后台删掉并保存，PDF 也删掉；再点一次取消">删除这张表</button>`;
  return orientGroup + tallGroup + breaks + red;
}

function renderDocs() {
  const box = $('prDocs');
  if (!pdfs.length) {
    box.innerHTML = allPdfs.length
      ? '<p class="pr-empty">这一类没有表了。点上面的「全部」看所有的表。</p>'
      : '<p class="pr-empty">这个文件夹里没有 PDF。先在主页「调整页面」里点「导出 PDF」，导完点「调整页面」回来。</p>';
    pageEls = [];
    $('prBar').classList.toggle('hidden', !allPdfs.length);
    renderTrack();
    updateStatus();
    return;
  }
  box.innerHTML = pdfs.map((doc, d) => {
    const excel = doc.wb_name
      ? `<span class="pr-doc-excel">📗 ${excelLinkHtml(`${escapeHtml(doc.wb_name)} › ${escapeHtml(doc.sheet)}`, excelTarget(doc))}`
        + `${doc.guessed ? '<span class="pr-doc-note" title="导出记录里没有这份 PDF，按「Excel 文件名_表名.pdf」找到的">按文件名对应</span>' : ''}</span>`
      : '<span class="pr-doc-warn">⚠ 找不到对应的 Excel，点了也打不开</span>';
    const notes = doc.notes || [];
    const pageNote = p => notes.filter(n => !n.fixed && n.page === p).map(n => n.text).join('\n');
    const pages = doc.pages.map(([w, h], p) => `
      <div class="pr-page is-loading" data-doc="${d}" data-page="${p}" style="aspect-ratio:${w} / ${h};--ar:${w / h}">
        <img class="pr-img" loading="lazy" decoding="async" draggable="false" alt="第 ${p + 1} 页"
          src="/api/pdf_review/page?pdf=${encodeURIComponent(doc.pdf)}&p=${p}&v=${doc.v}">
        <div class="pr-layer"></div>
        <span class="pr-pageno">${p + 1} / ${doc.pages.length}</span>
        ${pageNote(p) ? `<span class="pr-page-note" title="${escapeHtml(pageNote(p))}">⚠ 内容很少</span>` : ''}
      </div>`).join('');
    const fixedNotes = notes.filter(n => n.fixed);
    const openNotes = notes.filter(n => !n.fixed);
    const noteRows = fixedNotes.map(n => `<li class="pr-fill-fixed">✓ 自动修好：${escapeHtml(n.text)}</li>`).join('')
      + openNotes.map(n => `<li class="pr-fill-open">⚠ ${escapeHtml(n.text)}</li>`).join('');
    const noteBox = noteRows ? `<ul class="pr-fill-notes" aria-label="内容很少的页">${noteRows}</ul>` : '';
    const markable = doc.wb_path && !doc.error;
    const tail = suggestsFloor(doc) ? `<span class="pr-doc-tail" title="最后一页只印了约 ${doc.layout.tail}%，建议向下取整">末页 ${doc.layout.tail}%</span>` : '';
    const regen = doc.wb_path ? `<button type="button" class="pr-regen-btn" data-regen="${d}"
      title="只把这一份 Excel 的 PDF 重新导出一次，马上看效果。这份 Excel 上点了改方向、取整等还没改的，会先改进 Excel 再导出">🔄 重新生成</button>` : '';
    return `<article class="pr-doc" id="prDoc${d}" data-doc="${d}" data-mark="blue">
      <header class="pr-doc-head" title="${escapeHtml(doc.pdf)}">
        <span class="pr-doc-name">📄 ${escapeHtml(doc.name)}</span>
        <span class="pr-doc-pages">${doc.pages.length} 页</span>${tail}${excel}
        <span class="pr-doc-gap"></span>${markable ? headButtons(doc, d) : ''}${regen}
      </header>
      ${doc.error ? `<p class="pr-doc-warn">${escapeHtml(doc.error)}</p>` : `${noteBox}<div class="pr-pages">${pages}</div>`}
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
  const which = category === 'all' ? '' : `「${catOf(category).label}」`;
  if (allPdfs.length) $('prStatus').textContent = `${which}共 ${pdfs.length} 份 PDF、${pages} 页${which ? `（全部 ${allPdfs.length} 份）` : ''}`;
  $('prSummary').innerHTML = `${escapeHtml(which)}共 <b>${pdfs.length}</b> 份 · ${pages} 页`;
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
  // 只有 Ctrl + 单击是用 Excel 打开这一页；单击不做事（改什么都用标题上的按钮）
  if (!(e.ctrlKey || e.metaKey)) return;
  const d = Number(start.pageEl.dataset.doc);
  if (!pdfs[d].wb_path) {
    toast('这份 PDF 找不到对应的 Excel', true);
    return;
  }
  openPage(d, Number(start.pageEl.dataset.page), start.pageEl);
}

/* ---------- 标记 ----------
   红 = 删除这张表（marks）；工具（tools）：floor 向下取整（黄色）、auto 页高改回自动、orient 改方向、breaks 清手动分页。
   PDF 重新导出过（修改时间变了）标记就不算了：看到的已经不是标记时那一版 */

function loadMarks() {
  marks = {};
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem(MARKS_KEY) || '{}')[folder] || {}; } catch (_) {}
  for (const doc of allPdfs) {
    const m = saved[doc.pdf];
    if (m && m.v === doc.v && m.red && doc.wb_path && !doc.error) marks[doc.pdf] = { v: m.v, red: true };
  }
  saveMarks();
  loadTools();
}

function saveMarks() {
  try {
    const all = JSON.parse(localStorage.getItem(MARKS_KEY) || '{}');
    if (Object.keys(marks).length) all[folder] = marks; else delete all[folder];
    localStorage.setItem(MARKS_KEY, JSON.stringify(all));
  } catch (_) {}
}

function loadTools() {
  tools = {};
  let saved = {};
  let autoSeen = {};
  try { saved = JSON.parse(localStorage.getItem(TOOLS_KEY) || '{}')[folder] || {}; } catch (_) {}
  try { autoSeen = JSON.parse(localStorage.getItem(AUTO_KEY) || '{}')[folder] || {}; } catch (_) {}
  for (const doc of allPdfs) {
    const t = saved[doc.pdf];
    if (!t || t.v !== doc.v || !doc.wb_path || doc.error) continue;
    const kept = { v: t.v };
    // 程序自动标的黄（byAuto，或旧版按 40% 自动标过这一版）：按现在的标准不该标了就去掉，自己点的保留
    const autoMarked = t.byAuto || autoSeen[doc.pdf] === doc.v;
    if (t.floor && doc.pages.length > 1 && !(autoMarked && !suggestsFloor(doc))) {
      kept.floor = true;
      if (t.byAuto) kept.byAuto = true;
    }
    if (t.auto) kept.auto = true;
    if (t.orient === 1 || t.orient === 2) kept.orient = t.orient;
    if (t.breaks && manualOf(doc)) kept.breaks = true;
    if (Object.keys(kept).length > 1) tools[doc.pdf] = kept;
  }
  autoFloor();
  saveTools();
}

// 末页只多一点的表：每一版 PDF 自动标一次黄（建议取整）；你取消以后这一版不再标
function autoFloor() {
  let all = {};
  try { all = JSON.parse(localStorage.getItem(AUTO_KEY) || '{}'); } catch (_) {}
  const seen = all[folder] || {};
  const now = {};
  for (const doc of allPdfs) {
    if (!suggestsFloor(doc) || !doc.wb_path || doc.error) continue;
    now[doc.pdf] = doc.v;
    if (seen[doc.pdf] === doc.v || marks[doc.pdf]?.red) continue;
    const t = tools[doc.pdf];
    if (!t?.orient && !t?.auto && !t?.breaks) tools[doc.pdf] = { ...(t || {}), v: doc.v, floor: true, byAuto: true };
  }
  try {
    if (Object.keys(now).length) all[folder] = now; else delete all[folder];
    localStorage.setItem(AUTO_KEY, JSON.stringify(all));
  } catch (_) {}
}

function saveTools() {
  try {
    const all = JSON.parse(localStorage.getItem(TOOLS_KEY) || '{}');
    if (Object.keys(tools).length) all[folder] = tools; else delete all[folder];
    localStorage.setItem(TOOLS_KEY, JSON.stringify(all));
  } catch (_) {}
}

function editable(doc) {
  return doc && doc.wb_path && !doc.error && !deleted.has(doc.pdf) && !processing;
}

// floor 和 auto 是页高的两种改法，只能选一个；改方向、清手动分页时不取整（页数全变了，看新 PDF 再决定）
function setTool(doc, tool, on) {
  if (!editable(doc) || marks[doc.pdf]?.red) return;
  const t = { ...(tools[doc.pdf] || {}), v: doc.v };
  delete t.byAuto; // 自己点过就算自己定的
  if (tool === 'breaks') {
    if (on && !manualOf(doc)) return;
    if (on) { t.breaks = true; delete t.floor; } else delete t.breaks;
  } else if (tool === 'floor') {
    if (on && doc.pages.length < 2) return;
    delete t.auto;
    if (on) { t.floor = true; delete t.orient; delete t.breaks; } else delete t.floor;
  } else if (tool === 'auto') {
    if (tallOf(doc) === null) on = false; // 本来就是自动
    delete t.floor;
    if (on) t.auto = true; else delete t.auto;
  }
  if (Object.keys(t).length > 1) tools[doc.pdf] = t; else delete tools[doc.pdf];
}

function setOrient(doc, target) {
  if (!editable(doc) || marks[doc.pdf]?.red) return;
  const t = { ...(tools[doc.pdf] || {}), v: doc.v };
  delete t.byAuto;
  if (target === orientOf(doc)) delete t.orient;
  else { t.orient = target; delete t.floor; }
  if (Object.keys(t).length > 1) tools[doc.pdf] = t; else delete tools[doc.pdf];
}

function colorOf(doc) {
  if (deleted.has(doc.pdf)) return 'deleted';
  if (marks[doc.pdf]?.red) return 'red';
  return tools[doc.pdf]?.floor ? 'yellow' : 'blue';
}

function pageColor(doc, p) {
  const color = colorOf(doc);
  if (color === 'yellow') return p === doc.pages.length - 1 ? 'yellow' : 'blue'; // 取整只标最后一页
  return color;
}

// 只改颜色、角标和按钮状态；不改大小、不滚动
function paintMark(d) {
  const doc = pdfs[d];
  const article = $(`prDoc${d}`);
  if (!article) return;
  const color = colorOf(doc);
  const t = color === 'red' || color === 'deleted' ? {} : (tools[doc.pdf] || {});
  const locked = color === 'deleted' || processing || !!regenWb;
  article.dataset.mark = color;
  article.dataset.tool = t.orient || t.auto || t.breaks ? 'set' : '';
  const orient = orientOf(doc);
  article.querySelectorAll('[data-orient]').forEach(b => {
    const o = Number(b.dataset.orient);
    const pending = t.orient === o;
    const current = o === orient;
    b.classList.toggle('is-on', current && !t.orient);
    b.classList.toggle('is-was', current && !!t.orient);
    b.classList.toggle('is-pending', pending);
    b.textContent = pending ? `→ ${ORIENT_TEXT[o]}` : ORIENT_TEXT[o];
    b.setAttribute('aria-pressed', pending || (current && !t.orient));
    b.disabled = locked || color === 'red';
  });
  article.querySelectorAll('[data-tall]').forEach(b => {
    const floor = b.dataset.tall === 'floor';
    const auto = !floor && tallOf(doc) === null;
    b.classList.toggle('is-on', auto && !t.floor);
    b.classList.toggle('is-yellow', floor && !!t.floor);
    b.classList.toggle('is-pending', !floor && !!t.auto);
    b.setAttribute('aria-pressed', floor ? !!t.floor : auto || !!t.auto);
    b.disabled = locked || color === 'red' || (floor && doc.pages.length < 2);
  });
  article.querySelectorAll('[data-breaks]').forEach(b => {
    b.classList.toggle('is-pending', !!t.breaks);
    b.setAttribute('aria-pressed', !!t.breaks);
    b.textContent = t.breaks ? `→ 清掉 ${manualOf(doc)} 处手动分页` : `✂ 清手动分页（${manualOf(doc)} 处）`;
    b.disabled = locked || color === 'red';
  });
  article.querySelectorAll('[data-red]').forEach(b => {
    b.classList.toggle('is-on', color === 'red');
    b.setAttribute('aria-pressed', color === 'red');
    b.disabled = locked;
  });
  article.querySelectorAll('.pr-page').forEach(pageEl => {
    const p = Number(pageEl.dataset.page);
    pageEl.dataset.mark = pageColor(doc, p);
    pageEl.dataset.badge = color === 'yellow' && p === doc.pages.length - 1
      ? `⤓ 向下取整${doc.layout?.tail != null ? `（末页只印了 ${doc.layout.tail}%）` : ''}` : '';
  });
}

function changed(d) {
  saveMarks();
  saveTools();
  paintMark(d);
  updateProcessButton();
}

function toggleRed(d) {
  const doc = pdfs[d];
  if (!editable(doc)) return;
  if (marks[doc.pdf]?.red) delete marks[doc.pdf];
  else marks[doc.pdf] = { v: doc.v, red: true };
  changed(d);
}

function onOrient(d, target) {
  setOrient(pdfs[d], target);
  changed(d);
}

function onTall(d, which) {
  const doc = pdfs[d];
  if (which === 'floor') setTool(doc, 'floor', !tools[doc.pdf]?.floor);
  else if (tools[doc.pdf]?.floor) setTool(doc, 'floor', false); // 取消取整：回到原来的页高
  else setTool(doc, 'auto', !tools[doc.pdf]?.auto);
  changed(d);
}

function onBreaks(d) {
  const doc = pdfs[d];
  setTool(doc, 'breaks', !tools[doc.pdf]?.breaks);
  changed(d);
}

function redDocs() {
  return allPdfs.filter(doc => colorOf(doc) === 'red');
}

// 要统一修改的：[{doc, tool, value}]
function toolItems() {
  const items = [];
  for (const doc of allPdfs) {
    const t = tools[doc.pdf];
    if (!t || colorOf(doc) === 'red' || colorOf(doc) === 'deleted' || !doc.wb_path) continue;
    if (t.orient) items.push({ doc, tool: 'orient', value: t.orient });
    if (t.floor) items.push({ doc, tool: 'floor' });
    if (t.auto) items.push({ doc, tool: 'auto' });
    if (t.breaks) items.push({ doc, tool: 'breaks' });
  }
  return items;
}

const PROCESS_TEXT = { red: '🗑 删除标红', tools: '🛠 统一修改' };
const PROCESS_BUTTON = { red: 'prDeleteRed', tools: 'prApply' };
const PROCESS_KINDS = ['red', 'tools'];
let processKind = '';  // 正在处理的是哪个按钮：red / tools

function updateProcessButton() {
  const items = toolItems();
  const counts = { red: redDocs().length, tools: items.length };
  $('prCountRed').textContent = counts.red;
  $('prCountApply').textContent = counts.tools;
  const by = tool => items.filter(i => i.tool === tool).length;
  $('prApply').title = items.length
    ? `取整 ${by('floor')} 张、改方向 ${by('orient')} 张、页高改回自动 ${by('auto')} 张、清手动分页 ${by('breaks')} 张：后台改 Excel 并保存，再重新生成 PDF`
    : '在标题上点方向、页高的按钮做标记，再点这里统一修改';
  paintBulk();
  if (processing) return;
  for (const kind of PROCESS_KINDS) {
    document.querySelector(`[data-process-text="${kind}"]`).textContent = PROCESS_TEXT[kind];
    $(PROCESS_BUTTON[kind]).disabled = !counts[kind] || !!regenWb;
  }
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
  const where = page === null ? '整张表' : `第 ${page + 1} 页`;
  menu.querySelector('.pr-menu-title').textContent = `${doc.wb_name ? doc.wb_name.split('/').pop() : doc.name} › ${doc.sheet || ''} · ${where}`;
  menu.querySelector('[data-menu="open"] b').textContent = page === null ? '在 Excel 打开这张表' : `在 Excel 打开第 ${page + 1} 页`;
  const red = colorOf(doc) === 'red';
  menu.querySelector('[data-menu="red"] b').textContent = red ? '不删这张表了' : '删除这张表';
  menu.querySelector('[data-menu="red"]').classList.toggle('is-on', red);
  menu.querySelector('[data-menu="red"]').disabled = !editable(doc);
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
  else if (kind === 'red') toggleRed(d);
}

/* ---------- 删除标红 / 统一修改 ---------- */

function listLines(lines, limit = 12) {
  const shown = lines.slice(0, limit).map(line => `  · ${line}`);
  if (lines.length > limit) shown.push(`  · ……另外 ${lines.length - limit} 张`);
  return shown.join('\n');
}

function toolLine({ doc, tool, value }) {
  const n = doc.pages.length;
  const what = tool === 'floor' ? `向下取整 ${n} 页 → ${n - 1} 页`
    : tool === 'auto' ? `页高 定死 ${tallOf(doc)} 页 → 自动`
    : tool === 'breaks' ? `清掉 ${manualOf(doc)} 处手动分页，改成自动分页`
      : `${ORIENT_TEXT[orientOf(doc)]} → ${ORIENT_TEXT[value]}`;
  return `${doc.wb_name} › ${doc.sheet}：${what}`;
}

// onlyWb：只改这一份 Excel 上的标记（点这份的「🔄 重新生成」时）
async function runProcess(kind, onlyWb = '') {
  if (processing || regenWb) return;
  const files = docs => new Set(docs.map(doc => doc.wb_path)).size;
  let red = [];
  let items = [];
  let text;
  if (kind === 'tools') {
    items = toolItems().filter(i => !onlyWb || i.doc.wb_path === onlyWb);
    if (!items.length) return;
    text = (onlyWb ? `「${items[0].doc.wb_name}」上点的标记还没改进 Excel，先改好、保存，再重新生成这份 Excel 的 PDF。\n\n` : '')
      + `🛠 统一修改 ${items.length} 处（${files(items.map(i => i.doc))} 份 Excel）：\n${listLines(items.map(toolLine))}`
      + (items.some(i => i.tool === 'floor') ? '\n\n向下取整会先真导出一次核对：页数少了一页、字没缩到 85% 以下、PDF 里原来的字一个不少，才算数，不行的改回原样。' : '\n')
      + '\n改好的 Excel 保存后重新生成 PDF，末页只多一点的会重新判断、重新标黄（改了方向、清了手动分页的要看新的 PDF 再决定取不取整）。';
  } else {
    red = redDocs();
    if (!red.length) return;
    text = `🔴 删除 ${red.length} 张表（${files(red)} 份 Excel），删完保存，对应的 PDF 也删掉：\n${listLines(red.map(doc => `${doc.wb_name} › ${doc.sheet}`))}`
      + '\n\n删表前会先把别的表引用它的公式转成值；删表后出现新的引用错误时这份文件不保存。';
  }
  const ok = await appConfirm(text, kind === 'red'
    ? { title: '删除标红的表', tone: 'danger', okText: '删除' }
    : { title: onlyWb ? '改好并重新生成' : '统一修改页面设置', tone: 'warn', okText: onlyWb ? '改好并重新生成' : '开始修改' });
  if (!ok) return;
  if (onlyWb) setBusy(onlyWb);
  try {
    await api('/api/pdf_review/process', {
      red: red.map(doc => ({ pdf: doc.pdf })), yellow: [],
      tools: items.map(({ doc, tool, value }) => ({ pdf: doc.pdf, tool, ...(value ? { value } : {}) })),
    });
  } catch (e) {
    setBusy('');
    appAlert(e.message, { tone: 'error' });
    return;
  }
  processKind = kind;
  setProcessing(true);
  pollProcess();
}

function setProcessing(on) {
  processing = on;
  for (const kind of PROCESS_KINDS) {
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
    if (!processKind) processKind = status.stage === 'delete' ? 'red' : 'tools';
    const label = document.querySelector(`[data-process-text="${processKind}"]`);
    label.textContent = status.stage === 'delete' ? `⏳ 删除中 ${status.done} / ${status.total}`
      : `⏳ 修改中 ${status.done} / ${status.total}`;
    $(PROCESS_BUTTON[processKind]).classList.add('is-running');
    setTimeout(pollProcess, 1000);
    return;
  }
  setProcessing(false);
  setBusy('');
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
  return `<li>${body}${note}</li>`;
}

async function finishProcess(result) {
  // 改好的：Excel 已保存、PDF 已重新生成，新 PDF 修改时间变了，旧标记自动不算
  for (const item of result.deleted || []) {
    deleted.add(item.pdf);
    delete marks[item.pdf];
  }
  saveMarks();
  pdfs.forEach((_, d) => paintMark(d));
  updateProcessButton();
  const section = (title, items, opts) => items?.length
    ? `<h3>${title}</h3><ul>${items.map(i => resultLine(i, opts)).join('')}</ul>` : '';
  const html = [
    section(`🗑 删掉了 ${result.deleted?.length} 张表（已保存，PDF 也删了）`, result.deleted, { link: false }),
    section(`⚠ 没删掉 ${result.failed?.length} 张（还是红色，可以再点一次）`, result.failed),
    section(`🛠 改好了 ${result.changed?.length} 处（Excel 已保存，PDF 已重新生成，请再看一遍）`, result.changed),
    section(`⚠ 没改成 ${result.change_failed?.length} 处（打印设置没动，标记还在）`, result.change_failed),
  ].join('');
  const bad = result.failed?.length || result.change_failed?.length;
  toast(bad ? '处理完成，有没处理成功的，请看结果' : '处理完成', !!bad);
  $('prResultBody').innerHTML = html || '<p>没有要处理的</p>';
  $('prResult').dataset.tone = bad ? 'warn' : 'info';
  // 改了页面的 PDF 换成了新的一版：重新读列表，页面停在原来那份上
  const done = result.changed || [];
  if (done.length) {
    const doc = allPdfs.find(x => x.pdf === done[0].pdf);
    await reloadKeeping(done[0].pdf, doc?.wb_path || '');
  }
  $('prResult').showModal();
}

/* ---------- 单份 Excel 重新生成 PDF ---------- */
// 复检时在 Excel 里改了、保存了，想马上看改完的效果：只把这一份 Excel 的 PDF 重新导出一次。
// 做完重新读一遍列表；页面停在原来那份上（按标题位置对齐），不跳走。

// 重新生成大约要 10 秒：一点下去，按钮马上变「⏳ 生成中 N 秒」，这份 Excel 的页面盖一层「正在重新生成」，
// 做完才收起（用户 2026-11-08 提：以前没提示，以为卡住了或者没点上）
function setBusy(wb) {
  busyWb = wb || '';
  clearInterval(busyTimer);
  if (busyWb) {
    busySince = Date.now();
    busyTimer = setInterval(paintRegen, 1000);
  }
  paintRegen();
}

function paintRegen() {
  const wb = regenWb || busyWb;
  const secs = Math.max(0, Math.round((Date.now() - busySince) / 1000));
  document.querySelectorAll('[data-regen]').forEach(b => {
    const doc = pdfs[Number(b.dataset.regen)];
    const mine = wb && doc && doc.wb_path === wb;
    b.textContent = mine ? (busyWb ? `⏳ 生成中 ${secs} 秒` : '⏳ 生成中…') : '🔄 重新生成';
    b.classList.toggle('is-running', !!mine);
    b.disabled = !!wb || processing || deleted.has(doc?.pdf);
  });
  document.querySelectorAll('.pr-doc').forEach(article => {
    const doc = pdfs[Number(article.dataset.doc)];
    article.classList.toggle('is-regen', !!(wb && doc && doc.wb_path === wb));
  });
  updateProcessButton(); // 重新生成时两个处理按钮也先不能点
}

async function regenerate(d) {
  const doc = pdfs[d];
  if (!doc || !doc.wb_path || regenWb || processing) return;
  // 这份 Excel 上点了还没改的（改方向、取整、页高改回自动）：先改进 Excel 再重新生成，
  // 不然 Excel 没动、哈希一样，用户会以为没反应（用户 2026-11-08 提）
  if (toolItems().some(i => i.doc.wb_path === doc.wb_path)) {
    runProcess('tools', doc.wb_path);
    return;
  }
  try {
    await regenerateNow(doc);
  } finally {
    if (!regenWb) setBusy(''); // 取消了、出错了：收起提示；正常开始的等做完再收
  }
}

async function regenerateNow(doc) {
  if (allPdfs.some(x => x.wb_path === doc.wb_path && marks[x.pdf])) {
    const ok = await appConfirm(
      `重新生成「${doc.wb_name}」的 PDF（${allPdfs.filter(x => x.wb_path === doc.wb_path).length} 张表）？\n\n这份 Excel 上标的「删除这张表」会清掉：新的 PDF 是另一版，标记对不上了。`,
      { title: '重新生成 PDF', tone: 'warn', okText: '重新生成' });
    if (!ok) return;
  }
  let data;
  try {
    setBusy(doc.wb_path);
    data = await api('/api/pdf_review/regenerate', { pdf: doc.pdf });
    if (data.unchanged) {
      setBusy('');
      const again = await appConfirm(
        `「${data.name}」和上次导出 PDF 时一模一样（哈希值没变），PDF 已经是最新的。\n\n刚在 Excel 里改过的话，多半还没保存：先在 Excel 里保存，再点「重新生成」。\n\n还是要照样重新生成一次吗？`,
        { title: 'Excel 没有改过', tone: 'warn', okText: '照样重新生成' });
      if (!again) return;
      setBusy(doc.wb_path);
      data = await api('/api/pdf_review/regenerate', { pdf: doc.pdf, force: true });
    }
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
    return;
  }
  regenWb = doc.wb_path;
  paintRegen();
  toast(`正在重新生成「${doc.wb_name}」的 PDF，大约要 10 秒…`);
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
  setBusy('');
  if (!result) return;
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
  allPdfs = data.pdfs;
  deleted = new Set([...deleted].filter(p => allPdfs.some(x => x.pdf === p)));
  loadMarks(); // 重新生成的 PDF 修改时间变了，它的标记自动不算了（末页只多一点的重新自动标黄）
  applyCategory();
  let after = pdfs.findIndex(x => x.pdf === pdf);
  if (after < 0) after = pdfs.findIndex(x => x.wb_path === wbPath);
  const now = after >= 0 ? $(`prDoc${after}`) : null;
  if (now && top !== null) window.scrollBy(0, now.getBoundingClientRect().top - top);
  scheduleWhere();
}

/* ---------- 看到哪里（底部条） ---------- */
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

// 大图 = 宽度占满屏幕宽，中图 = 高度占满屏幕高，小图 = 一行放好几页（用户 2026-11-03 定）。
// 换大小后停在原来看的那一页的同一位置；选择记在 localStorage
function setView(view, keep = true) {
  if (!VIEWS.includes(view)) view = 'small';
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
  const orient = e.target.closest('[data-orient]');
  if (orient) onOrient(Number(orient.dataset.doc), Number(orient.dataset.orient));
  const tall = e.target.closest('[data-tall]');
  if (tall) onTall(Number(tall.dataset.doc), tall.dataset.tall);
  const breaks = e.target.closest('[data-breaks]');
  if (breaks) onBreaks(Number(breaks.dataset.breaks));
  const red = e.target.closest('[data-red]');
  if (red) toggleRed(Number(red.dataset.red));
  const regen = e.target.closest('[data-regen]');
  if (regen) regenerate(Number(regen.dataset.regen));
});
$('prCats').addEventListener('click', e => {
  const cat = e.target.closest('[data-cat]');
  if (cat) setCategory(cat.dataset.cat);
  if (e.target.closest('#prBulk')) runBulk();
});
$('prDeleteRed').addEventListener('click', () => runProcess('red'));
$('prApply').addEventListener('click', () => runProcess('tools'));
$('prFolder').addEventListener('keydown', e => { if (e.key === 'Enter') load(); });
$('prTrack').addEventListener('click', jumpTo);
$('prView').addEventListener('click', e => {
  const button = e.target.closest('[data-view]');
  if (button) setView(button.dataset.view);
});
$('prResultBody').addEventListener('click', e => {
  const button = e.target.closest('[data-result-doc]');
  if (button) openPage(Number(button.dataset.resultDoc), Number(button.dataset.resultPage));
});
window.addEventListener('scroll', scheduleWhere, { passive: true });
window.addEventListener('resize', () => { if (pageEls.length) fitTicks(); scheduleWhere(); });
watchBars();
{
  let view = 'small';
  try { view = localStorage.getItem(VIEW_KEY) || view; } catch (_) {}
  setView(view, false);
}
try {
  $('prFolder').value = localStorage.getItem(FOLDER_KEY) || '';
  $('prOldFolder').value = localStorage.getItem(OLD_FOLDER_KEY) || '';
} catch (_) {}
try { category = catOf(localStorage.getItem(CAT_KEY) || 'all').key; } catch (_) {}
if ($('prFolder').value) load();
// 刷新页面时一键处理还在后台跑：接着看进度
api('/api/pdf_review/process').then(status => { if (status.running) { setProcessing(true); pollProcess(); } }).catch(() => {});
// 重新生成还在后台跑：接着看，做完刷新列表
api('/api/pdf_review/regenerate').then(status => {
  if (status.running) { regenWb = status.wb_path || ''; paintRegen(); pollRegen(status.pdf); }
}).catch(() => {});
