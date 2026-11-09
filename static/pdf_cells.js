'use strict';
// 调整页面里的切字（用户 2026-11-09 定：把「查看切字」并进调整页面，页面和单元格在一个页面里处理）。
// - 进页面 PDF 都是最新的以后，后台自动查切字（就是原来的「扫描并标记」，直接用这里导出的 PDF 检测，不再导一遍）。
// - 分类多一行「单元格」：文字截断、边框压线……点了只看有这一类格子的表，并能一次给这一类选工具。
// - 小图上用红框画出查出来的格子，角上的数字 = 这一页有几格；每张表下面固定一条，列出这张表的格子和工具。
// - 单击小图看大图：大图上单击格子选修法、Ctrl + 单击多选、拖一个框把碰到的格子都选上；
//   格子右上角 ✕ = 不用处理（灰色虚线），再点 ↺ 恢复。
// - 选好的工具和查看切字的单元格工具箱共用（localStorage clear_handle_excel.cellTools），底部「🛠 统一修改」一起改：
//   先改页面 → 页面改过的重新查切字 → 按工具改格子、保存、重新生成 PDF、复检（core/adjust_cells.py）。
// 点按钮只换颜色和文字，不挪位置；查切字时新出来的格子会把下面的内容往下推，这时按看着的那张表把位置对回来。
// 这个文件在 pdf_review.js 之前加载：里面用到的 pdf_review.js 的变量和函数都只在函数里用。

const PC_TOOLS_KEY = 'clear_handle_excel.cellTools'; // {工作台文件夹: {卡片 id: 'grow:1.5' | 'shrink' | 'break' | 'none' | 'skip'}}
const PC_TOOL_DEFS = [
  { value: 'grow:1.2', tool: 'grow', ratio: 1.2, label: '行高 ×1.2', tip: '开自动换行，先补够放下文字要的高度，再加「字号 × 1.2」' },
  { value: 'grow:1.5', tool: 'grow', ratio: 1.5, label: '行高 ×1.5', tip: '开自动换行，先补够放下文字要的高度，再加「字号 × 1.5」' },
  { value: 'grow:2', tool: 'grow', ratio: 2, label: '行高 ×2', tip: '开自动换行，先补够放下文字要的高度，再加「字号 × 2」' },
  { value: 'shrink', tool: 'shrink', label: '缩小字体', tip: '缩小字体填充：字缩到放得下（会关掉自动换行）' },
  { value: 'break', tool: 'break', label: '上面加分页线', tip: '在这格上面加一条手动分页线，整格挪到下一页' },
  { value: 'none', tool: 'none', label: '不动', tip: '这次不改这格，留着标色，你自己打开 Excel 改' },
  { value: 'skip', tool: 'skip', label: '✕ 不用处理', tip: '这格其实没问题：统一修改时记下来，以后不再查出来，Excel 里的标色去掉' },
];
const PC_TOOL_BY_VALUE = Object.fromEntries(PC_TOOL_DEFS.map(t => [t.value, t]));
const PC_STRIP_LIMIT = 12;  // 每张表下面先列这么多格，多的点「全部列出」
const PC_POLL_MS = 2500;

let pcData = { active: false, folder: '', cards: [], files: {}, scan: {} };
let pcByDoc = new Map();      // Excel + 表 → 这张表的格子
let pcFiles = new Map();      // Excel（统一写法）→ ok / changed
let pcChoices = {};
let pcSig = '';               // 格子有没有变（变了才重画）
let pcTimer = 0;
let pcAsked = '';             // 这个文件夹这次打开页面已经自动叫过「查切字」了
let pcExpanded = new Set();   // 点了「全部列出」的表
let pcStarting = false;
try { pcChoices = JSON.parse(localStorage.getItem(PC_TOOLS_KEY) || '{}') || {}; } catch (_) {}

const pcNorm = p => String(p || '').replace(/\//g, '\\').toLowerCase();
const pcDocKey = (wb, sheet) => `${pcNorm(wb)}\u0000${sheet || ''}`;
const pcEsc = value => String(value ?? '').replace(/[&<>"']/g, ch => (
  { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const pcById = id => pcData.cards.find(c => c.id === id);

/* ---------- 每格选什么工具 ---------- */

function pcFolderChoices() {
  const key = pcData.folder || '';
  if (!pcChoices[key] || typeof pcChoices[key] !== 'object') pcChoices[key] = {};
  return pcChoices[key];
}

function pcSave() {
  try { localStorage.setItem(PC_TOOLS_KEY, JSON.stringify(pcChoices)); } catch (_) {}
}

function pcOptions(card) {
  return PC_TOOL_DEFS.filter(t => t.tool === 'none' || t.tool === 'skip' || (card.tools || []).includes(t.tool));
}

function pcToolOf(card) {
  const value = pcFolderChoices()[card.id] ?? card.default_tool ?? 'none';
  const def = PC_TOOL_BY_VALUE[value];
  return def && pcOptions(card).includes(def) ? value : 'none';
}

function pcSetTool(cards, value) {
  const choices = pcFolderChoices();
  const def = PC_TOOL_BY_VALUE[value];
  for (const card of cards) {
    if (def && pcOptions(card).includes(def)) choices[card.id] = value;
  }
  pcSave();
}

// ✕ 不用处理 ↔ ↺ 恢复（回到默认的选法）
function pcToggleSkip(card) {
  const choices = pcFolderChoices();
  if (pcToolOf(card) === 'skip') delete choices[card.id];
  else choices[card.id] = 'skip';
  pcSave();
}

// 统一修改要做的：cells 按工具改，skips 标成不用处理
function pcPlan() {
  const cells = [];
  const skips = [];
  if (!pcData.active) return { cells, skips };
  for (const card of pcData.cards) {
    if (!pcVisibleCard(card)) continue;
    const value = pcToolOf(card);
    const def = PC_TOOL_BY_VALUE[value];
    if (value === 'skip') skips.push(card);
    else if (def && def.tool !== 'none') cells.push({ card, def });
  }
  return { cells, skips };
}

// 网页上列出来的格子才算（被删掉、找不到 PDF 的表上的格子不算）
function pcVisibleCard(card) {
  return allPdfs.some(doc => !doc.error && !deleted.has(doc.pdf) && pcDocKey(doc.wb_path, doc.sheet) === pcDocKey(card.wb_path, card.sheet));
}

function pcCount() {
  const { cells, skips } = pcPlan();
  return cells.length + skips.length;
}

function pcPayload() {
  const { cells, skips } = pcPlan();
  return {
    cells: cells.map(({ card, def }) => ({ card_id: card.id, tool: def.tool, ...(def.ratio ? { ratio: def.ratio } : {}) })),
    skips: skips.map(card => card.id),
  };
}

// 确认框里的几行
function pcPlanLines() {
  const { cells, skips } = pcPlan();
  const lines = [];
  const by = new Map();
  for (const { def } of cells) by.set(def.label, (by.get(def.label) || 0) + 1);
  for (const [label, n] of by) lines.push(`${label}：${n} 格`);
  if (skips.length) lines.push(`✕ 不用处理：${skips.length} 格（记下来，以后不再查出来，去掉标色）`);
  return lines;
}

/* ---------- 读格子 ---------- */

function pcCardsOf(doc) {
  return doc && doc.wb_path ? pcByDoc.get(pcDocKey(doc.wb_path, doc.sheet)) || [] : [];
}

// 这格在这一页上的框（PDF 是旧的、页数对不上就不画）
function pcBoxesOn(doc, card, p) {
  if (!doc || doc.stale) return [];
  const pdf = pcNorm(doc.pdf);
  return (card.boxes || []).filter(b => pcNorm(b.pdf) === pdf && b.page === p && (!b.pages || b.pages === doc.pages.length));
}

function pcPagesOf(doc, card) {
  if (!doc || doc.stale) return [];
  const pdf = pcNorm(doc.pdf);
  return [...new Set((card.boxes || []).filter(b => pcNorm(b.pdf) === pdf && (!b.pages || b.pages === doc.pages.length)).map(b => b.page))].sort((a, b) => a - b);
}

function pcFileState(doc) {
  if (!pcData.active || !doc.wb_path) return 'new';
  return pcFiles.get(pcNorm(doc.wb_path)) || 'new';
}

async function pcLoad({ quiet = false } = {}) {
  if (!folder) return;
  let data;
  try {
    data = await api(`/api/pdf_review/cells?folder=${encodeURIComponent(folder)}`);
  } catch (e) {
    if (!quiet) toast(`切字结果没读到：${e.message}`, true);
    pcSchedule(true);
    return;
  }
  const sig = JSON.stringify([data.active, data.cards.map(c => [c.id, c.boxes, c.cell_fix, c.default_tool, c.tools]), data.files]);
  pcData = data;
  pcByDoc = new Map();
  for (const card of data.cards) {
    const key = pcDocKey(card.wb_path, card.sheet);
    if (!pcByDoc.has(key)) pcByDoc.set(key, []);
    pcByDoc.get(key).push(card);
  }
  pcFiles = new Map(Object.entries(data.files || {}).map(([path, state]) => [pcNorm(path), state]));
  if (sig !== pcSig) {
    pcSig = sig;
    pcRestoreCategory();
    // 只看单元格某一类时，哪些表算这一类也变了：整个重排；平时只把框和格子条换掉
    if (category.startsWith('cell:')) applyCategory(); else pcRedraw();
  } else pcPaintScan();
  pcPanel();
  pcSchedule(data.scan?.running);
}

// 后台查切字的实时进度框（static/job_panel.js，用户 2026-11-09 要：所有批量操作都要有）。
// 统一修改里重新查切字时用统一修改自己的进度框，这里不出
function pcPanel() {
  const scan = pcData.scan || {};
  if (!window.JobPanel || processing || scan.kind !== 'batch_scan') return;
  const log = scan.log || [];
  if (scan.running) {
    JobPanel.run('pdf_review:scan', {
      title: '正在后台查切字', count: JobPanel.count(scan.done, scan.total, '份 Excel'),
      now: scan.current ? `📗 ${scan.current}` : '正在启动后台 Excel…', step: scan.step || '',
      fraction: scan.total ? scan.done / scan.total : 0, log, folder,
    });
  } else {
    const bad = log.filter(i => !i.ok).length;
    JobPanel.finish('pdf_review:scan', {
      title: scan.paused ? '查切字暂停了，可以点「▶ 接着查」' : `切字查完了${bad ? `，${bad} 份没查成` : ''}`,
      warn: scan.paused || bad > 0, log, folder,
      step: '查出的格子画在 PDF 上了，单击页面看大图、选修法，再点底部「🛠 统一修改」。',
    });
  }
}

function pcSchedule(fast) {
  clearTimeout(pcTimer);
  // 查切字时 2.5 秒看一次；平时 15 秒看一次（工作台那边复检、提交完这里也跟着变）
  pcTimer = setTimeout(() => pcLoad({ quiet: true }), fast ? PC_POLL_MS : 15000);
}

// 分类记的是单元格的某一类：格子读到了再换过去（打开页面时还没读到）
function pcRestoreCategory() {
  let saved = '';
  try { saved = localStorage.getItem(CAT_KEY) || ''; } catch (_) {}
  if (saved.startsWith('cell:') && category !== saved && pcCategories().some(c => c.key === saved)) {
    category = saved;
    applyCategory();
  } else if (category.startsWith('cell:') && !pcCategories().some(c => c.key === category)) {
    category = 'all';
    applyCategory();
  }
}

/* ---------- 进页面自动查切字 ---------- */

// 打开页面、Excel 改过的 PDF 都换成最新的以后调用：有没查过 / 改过的 Excel 就在后台查
async function pcAutoCheck(force = false) {
  if (!folder || processing || regenWb || pcStarting) return;
  if (!force && pcAsked === folder) return;
  pcAsked = folder;
  pcStarting = true;
  let data;
  try {
    data = await api('/api/pdf_review/check', { folder, highlight: { types: pcTypeColors() } });
  } catch (e) {
    pcStarting = false;
    toast(`没能开始查切字：${e.message}`, true);
    return;
  }
  pcStarting = false;
  if (data.started) toast(`正在后台查切字（${data.files} 份 Excel）…页面照样能看、能点`);
  else if (data.busy && force) appAlert(data.busy, { tone: 'warn' });
  pcLoad({ quiet: true });
}

// 和查看切字页面共用的标色颜色（问题类型 → Excel 里那种颜色）
function pcTypeColors() {
  const colors = {};
  try {
    const saved = JSON.parse(localStorage.getItem('clear_handle_excel.typeHighlightColors') || 'null') || {};
    for (const [type, color] of Object.entries(saved)) {
      if (/^#[0-9a-f]{6}$/i.test(color?.fill || '') && /^#[0-9a-f]{6}$/i.test(color?.font || '')) colors[type] = color;
    }
  } catch (_) {}
  return colors;
}

async function pcPauseScan() {
  try {
    await api('/api/batch/pause', {});
    toast('查切字会在这份 Excel 做完后暂停');
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
  }
}

async function pcResumeScan() {
  try {
    await api('/api/resume', {});
    toast('接着查切字…');
  } catch (e) {
    appAlert(e.message, { tone: 'error' });
  }
  pcLoad({ quiet: true });
}

/* ---------- 分类：单元格那一行 ---------- */

const PC_TYPE_ORDER = ['text_truncated', 'border_clip', 'box_overflow', 'page_break_cut', 'hash_overflow', 'text_missing'];

function pcCategories() {
  if (!pcData.active || !pcData.cards.length) return [];
  const types = new Map();
  for (const card of pcData.cards) {
    if (!types.has(card.type)) types.set(card.type, { label: card.type_label, n: 0 });
    types.get(card.type).n += 1;
  }
  const cats = [...types].sort(([a], [b]) => PC_TYPE_ORDER.indexOf(a) - PC_TYPE_ORDER.indexOf(b)).map(([type, info]) => ({
    key: `cell:${type}`, label: info.label, n: info.n, unit: '格', type,
    hint: `只看有「${info.label}」格子的表，还能一次给这一类选工具`,
    test: doc => pcCardsOf(doc).some(c => c.type === type),
  }));
  cats.unshift({ key: 'cell:any', label: '有切字的表', n: allPdfs.filter(doc => pcCardsOf(doc).length).length, unit: '张',
    hint: '只看查出切字的表', test: doc => pcCardsOf(doc).length > 0 });
  return cats;
}

function pcCategoryCards() {
  const type = category.startsWith('cell:') ? category.slice(5) : '';
  if (!type) return [];
  return pcData.cards.filter(c => pcVisibleCard(c) && (type === 'any' || c.type === type));
}

function pcCatsRowHtml() {
  if (!allPdfs.length) return '';
  const chips = pcCategories().map(c => `<button type="button" class="pr-cat pc-cat${c.key === category ? ' is-on' : ''}" data-cat="${c.key}"
      title="${pcEsc(c.hint)}" aria-pressed="${c.key === category}">${pcEsc(c.label)}<span class="pr-cat-n">${c.n}${c.key === 'cell:any' ? '' : ' 格'}</span></button>`).join('');
  return `<div class="pr-cat-row pc-cat-row"><span class="pr-cat-title">单元格</span>${chips || '<span class="pc-cat-empty">—</span>'}`
    + `<span class="pr-cat-gap"></span><span id="pcScan" class="pc-scan" role="status" aria-live="polite">${pcScanHtml()}</span></div>`
    + pcBulkHtml();
}

// 选了单元格的某一类：这一类一起用什么工具
function pcBulkHtml() {
  const cards = pcCategoryCards();
  if (!cards.length) return '';
  const defs = PC_TOOL_DEFS.filter(t => t.tool === 'none' || t.tool === 'skip' || cards.some(c => (c.tools || []).includes(t.tool)));
  return `<div class="pc-bulk" data-pc-bulk><span class="pc-bulk-label">这一类 ${cards.length} 格都用：</span>`
    + defs.map(t => `<button type="button" class="pc-chip pc-chip-lg${t.tool === 'skip' ? ' is-skip' : ''}" data-pc-bulk-tool="${t.value}" title="${pcEsc(t.tip)}"
        aria-pressed="${cards.every(c => pcToolOf(c) === t.value || !pcOptions(c).includes(t))}">${pcEsc(t.label)}</button>`).join('')
    + '<span class="pc-bulk-hint">单个格子在表下面或大图里另选；点底部「🛠 统一修改」才改 Excel</span></div>';
}

function pcScanHtml() {
  const scan = pcData.scan || {};
  if (scan.running) {
    return `<span class="pc-scan-text">🔍 正在查切字 第 ${Math.min(scan.done + 1, scan.total || 1)} / ${scan.total || '?'} 份${scan.current ? ` · ${pcEsc(scan.current)}` : ''}</span>`
      + '<button type="button" class="btn btn-outline btn-sm" data-pc-pause>⏸ 暂停</button>';
  }
  if (scan.busy) return '<span class="pc-scan-text">后台在忙，做完再查切字</span>';
  if (scan.paused && pcData.active) {
    return '<span class="pc-scan-text">⏸ 查切字暂停了</span><button type="button" class="btn btn-outline btn-sm" data-pc-resume>▶ 接着查</button>';
  }
  const docs = allPdfs.filter(doc => doc.wb_path && !doc.error);
  const todo = docs.filter(doc => pcFileState(doc) !== 'ok').length;
  if (!pcData.active || todo) {
    return `<span class="pc-scan-text">${pcData.active ? `${todo} 张表还没查切字` : '还没查切字'}</span>`
      + '<button type="button" class="btn btn-outline btn-sm" data-pc-check title="没查过、查过以后又改过的 Excel 在后台查切字（没改过的跳过）">🔍 查切字</button>';
  }
  const { cells, skips } = pcPlan();
  return `<span class="pc-scan-text">✓ 切字查完了：${pcData.cards.length} 格待处理${cells.length || skips.length ? `，选了修法 ${cells.length} 格、不用处理 ${skips.length} 格` : ''}</span>`;
}

function pcPaintScan() {
  const box = $('pcScan');
  if (box) box.innerHTML = pcScanHtml();
}

/* ---------- 小图上的框、表下面的格子条 ---------- */

function pcBoxStyle(b) {
  const [x0, y0, x1, y1] = b.box;
  const pad = 0.003;
  return `left:${((x0 - pad) * 100).toFixed(3)}%;top:${((y0 - pad) * 100).toFixed(3)}%;`
    + `width:${((x1 - x0 + pad * 2) * 100).toFixed(3)}%;height:${((y1 - y0 + pad * 2) * 100).toFixed(3)}%`;
}

function pcLayerHtml(doc, p) {
  const cards = pcCardsOf(doc);
  const boxes = cards.flatMap(card => pcBoxesOn(doc, card, p).map(b =>
    `<span class="pc-mini" data-pc-card="${pcEsc(card.id)}" data-state="${pcStateOf(card)}" style="${pcBoxStyle(b)}"></span>`)).join('');
  const cellsOnPage = cards.filter(card => pcBoxesOn(doc, card, p).length).length;
  return boxes + (cellsOnPage ? `<span class="pc-badge" title="这一页有 ${cellsOnPage} 格切字，单击看大图、选修法">${cellsOnPage}</span>` : '');
}

function pcStateOf(card) {
  const value = pcToolOf(card);
  return value === 'skip' ? 'skip' : value === 'none' ? 'none' : 'tool';
}

function pcStripHtml(doc, d) {
  if (!doc.wb_path || doc.error) return '';
  const cards = pcCardsOf(doc);
  const state = pcFileState(doc);
  let head;
  if (!cards.length) {
    const scanning = pcData.scan?.running;
    const text = state === 'ok' ? '✓ 没查出切字'
      : state === 'changed' ? (scanning ? '⏳ Excel 改过，正在 / 等着重新查切字' : 'Excel 改过，还没重新查切字')
        : scanning ? '⏳ 等着查切字' : '还没查切字';
    return `<div class="pc-strip is-empty" data-pc-doc="${d}"><span class="pc-strip-title">单元格</span><span class="pc-strip-note">${text}</span></div>`;
  }
  const type = category.startsWith('cell:') && category !== 'cell:any' ? category.slice(5) : '';
  const limit = pcExpanded.has(doc.pdf) ? cards.length : PC_STRIP_LIMIT;
  const shown = cards.slice(0, limit);
  const more = cards.length - shown.length;
  head = `<div class="pc-strip-head"><span class="pc-strip-title">这张表的单元格问题 <b>${cards.length} 格</b></span>`
    + `<span class="pc-strip-note">${state === 'changed' ? '⚠ 查完以后 Excel 改过，框的位置可能不准，统一修改时会先重新查一遍 · ' : ''}`
    + '默认按分类的选法，这里可以单格另选 · 单击格子名用 Excel 打开 · 单击上面的页面看大图</span></div>';
  return `<div class="pc-strip" data-pc-doc="${d}">${head}<div class="pc-cells">`
    + shown.map(card => pcCellHtml(doc, card, type)).join('') + '</div>'
    + (more > 0 ? `<button type="button" class="btn btn-outline btn-sm pc-more" data-pc-more="${d}">全部列出（还有 ${more} 格）</button>` : '')
    + '</div>';
}

function pcCellHtml(doc, card, type) {
  const where = `${card.sheet}!${card.cell}`;
  const pages = pcPagesOf(doc, card);
  const pageText = pages.length ? `第 ${pages.map(p => p + 1).join('、')} 页` : '';
  const jump = { file: doc.wb_name, sheet: card.sheet, cell: (card.cell || '').split(':')[0], folder };
  const value = pcToolOf(card);
  const chips = pcOptions(card).map(t => `<button type="button" class="pc-chip${t.tool === 'skip' ? ' is-skip' : ''}" data-pc-tool="${t.value}" data-pc-card="${pcEsc(card.id)}"
      aria-pressed="${t.value === value}" title="${pcEsc(t.tip)}">${pcEsc(t.label)}</button>`).join('');
  const fix = card.cell_fix ? `<p class="pc-note">用过「${pcEsc(card.cell_fix)}」复检还没过：换个工具，或打开 Excel 自己改</p>` : '';
  return `<div class="pc-cell${type && card.type !== type ? ' is-dim' : ''}" data-pc-card="${pcEsc(card.id)}" data-state="${pcStateOf(card)}">`
    + `<div class="pc-cell-text" title="${pcEsc(card.text)}">${pcEsc(card.text) || '（没有文字）'}</div>`
    + `<div class="pc-cell-info">${excelLinkHtml(`<b>${pcEsc(where)}</b>`, jump)}<span>${pcEsc(card.types.join(' + '))}${pageText ? ` · ${pageText}` : ''}</span></div>`
    + fix + `<div class="pc-tools" role="group" aria-label="${pcEsc(where)} 用什么修法">${chips}</div></div>`;
}

// 格子变了（查切字新出来、统一修改后）：重画框和格子条。表下面的格子条变高会把下面的表往下推，
// 按现在看着的那张表的标题把位置对回来，看着的东西不跳
function pcRedraw() {
  const box = $('prDocs');
  if (!box) return;
  renderCats();
  const articles = [...box.querySelectorAll('.pr-doc')];
  if (!articles.length) { pcPaintScan(); return; }
  const anchor = articles.find(a => a.getBoundingClientRect().bottom > 80) || articles[0];
  const before = anchor.getBoundingClientRect().top;
  for (const article of articles) {
    const d = Number(article.dataset.doc);
    const doc = pdfs[d];
    if (!doc) continue;
    article.querySelectorAll('.pr-page').forEach(pageEl => {
      const layer = pageEl.querySelector('.pr-layer');
      if (layer) layer.innerHTML = pcLayerHtml(doc, Number(pageEl.dataset.page));
    });
    const strip = article.querySelector('.pc-strip');
    const html = pcStripHtml(doc, d);
    if (strip) strip.outerHTML = html || '<div class="pc-strip" hidden></div>';
    else if (html) article.insertAdjacentHTML('beforeend', html);
  }
  const shift = anchor.getBoundingClientRect().top - before;
  if (Math.abs(shift) > 1 && window.scrollY > 0) window.scrollBy(0, shift);
  updateProcessButton();
  if (typeof scheduleWhere === 'function') scheduleWhere();
}

// 只换颜色、按下状态和文字（点工具时），不改大小、不挪位置
function pcPaint() {
  document.querySelectorAll('[data-pc-card]').forEach(el => {
    const card = pcById(el.dataset.pcCard);
    if (!card) return;
    const value = pcToolOf(card);
    if (el.matches('.pc-cell, .pc-mini, .pc-box')) el.dataset.state = pcStateOf(card);
    if (el.matches('[data-pc-tool]')) el.setAttribute('aria-pressed', String(el.dataset.pcTool === value));
  });
  document.querySelectorAll('[data-pc-bulk-tool]').forEach(el => {
    const cards = pcCategoryCards();
    const def = PC_TOOL_BY_VALUE[el.dataset.pcBulkTool];
    el.setAttribute('aria-pressed', String(cards.length > 0 && cards.every(c => pcToolOf(c) === def.value || !pcOptions(c).includes(def))));
  });
  pcPaintBig();
  pcPaintScan();
  updateProcessButton();
}

/* ---------- 大图：单击小图打开 ---------- */
// 宽度占满屏幕、上下滚动看；格子上的框：单击选中（变蓝）并在旁边浮出修法框，Ctrl + 单击加选 / 去掉，
// 从空白处拖一个框，碰到的格子都选上；框右上角 ✕ = 不用处理（灰色虚线），再点 ↺ 恢复。
// ← → 翻页（翻到头接着上一张 / 下一张表），Esc 先取消选择、再关掉。右键格子用 Excel 打开那一格。

let pcBig = null;            // {d, p}
let pcSel = new Set();       // 选中的格子
let pcDrag = null;           // 拖框：{x, y, ctrl, moved}
let pcPress = null;          // 按在格子上：{id, x, y}

function pcBuildBig() {
  if ($('pcBig')) return;
  document.body.insertAdjacentHTML('beforeend', `
    <dialog id="pcBig" class="pc-big" aria-labelledby="pcBigTitle">
      <header class="pc-big-head">
        <span id="pcBigTitle" class="pc-big-title"></span>
        <span class="pc-big-hint">单击格子选修法 · Ctrl + 单击多选 · 拖框多选 · ✕ 不用处理 · 右键用 Excel 打开 · ← → 翻页 · Esc 关</span>
        <span id="pcBigCount" class="pc-big-count"></span>
        <button type="button" id="pcBigPrev" class="pc-big-nav" title="上一页（←）">‹</button>
        <button type="button" id="pcBigNext" class="pc-big-nav" title="下一页（→）">›</button>
        <button type="button" id="pcBigClose" class="pc-big-close" title="关掉大图（Esc）" aria-label="关掉大图">✕</button>
      </header>
      <div id="pcBigBody" class="pc-big-body">
        <div id="pcBigPage" class="pc-big-page">
          <img id="pcBigImg" class="pc-big-img" alt="" draggable="false">
          <div id="pcBigLayer" class="pc-big-layer"></div>
          <div id="pcDragBox" class="pc-drag" hidden></div>
        </div>
      </div>
      <div id="pcPop" class="pc-pop" role="dialog" aria-label="选修法" hidden></div>
      <div id="pcBigLinks" hidden></div>
    </dialog>`);
  const dlg = $('pcBig');
  dlg.addEventListener('cancel', e => { e.preventDefault(); pcEscape(); });
  dlg.addEventListener('keydown', e => {
    if (e.key === 'ArrowLeft') { e.preventDefault(); pcStep(-1); }
    if (e.key === 'ArrowRight') { e.preventDefault(); pcStep(1); }
  });
  $('pcBigClose').addEventListener('click', pcCloseBig);
  $('pcBigPrev').addEventListener('click', () => pcStep(-1));
  $('pcBigNext').addEventListener('click', () => pcStep(1));
  const page = $('pcBigPage');
  page.addEventListener('pointerdown', pcBigDown);
  dlg.addEventListener('pointermove', pcBigMove);
  dlg.addEventListener('pointerup', pcBigUp);
  dlg.addEventListener('pointercancel', () => { pcDrag = null; pcPress = null; $('pcDragBox').hidden = true; });
  page.addEventListener('contextmenu', pcBigContext);
  // 点大图外面的空白（深色底）不关，免得点歪了丢掉选择；只有 ✕ 和 Esc 关
  $('pcPop').addEventListener('click', pcPopClick);
  $('pcBigLayer').addEventListener('click', e => {
    const x = e.target.closest('[data-pc-x]');
    if (!x) return;
    e.stopPropagation();
    const card = pcById(x.dataset.pcX);
    if (card) { pcToggleSkip(card); pcPaint(); }
  });
}

function pcOpenBig(d, p) {
  const doc = pdfs[d];
  if (!doc || doc.error) return;
  pcBuildBig();
  pcSel = new Set();
  pcShowPage(d, p);
  const dlg = $('pcBig');
  if (!dlg.open) dlg.showModal();
  $('pcBigBody').focus?.();
}

function pcShowPage(d, p) {
  const doc = pdfs[d];
  pcBig = { d, p };
  pcSel = new Set([...pcSel].filter(id => pcCardsOf(doc).some(c => c.id === id)));
  const [w, h] = doc.pages[p] || [1, 1.414];
  const page = $('pcBigPage');
  page.style.aspectRatio = `${w} / ${h}`;
  $('pcBigImg').src = `/api/pdf_review/page?pdf=${encodeURIComponent(doc.pdf)}&p=${p}&v=${doc.v}`;
  $('pcBigImg').alt = `${doc.name} 第 ${p + 1} 页`;
  $('pcBigTitle').textContent = `🔍 ${doc.wb_name || doc.name}${doc.sheet ? ` › ${doc.sheet}` : ''} · 第 ${p + 1} / ${doc.pages.length} 页`;
  const boxes = pcCardsOf(doc).flatMap(card => pcBoxesOn(doc, card, p).map(b => `
    <div class="pc-box" data-pc-card="${pcEsc(card.id)}" style="${pcBoxStyle(b)}" title="${pcEsc(`${card.sheet}!${card.cell} · ${card.types.join(' + ')}`)}">
      <button type="button" class="pc-x" data-pc-x="${pcEsc(card.id)}"></button><span class="pc-tag"></span></div>`)).join('');
  $('pcBigLayer').innerHTML = boxes;
  $('pcBigBody').scrollTop = 0;
  pcHidePop();
  pcPaintBig();
  $('pcBigPrev').disabled = !pcNeighbor(-1);
  $('pcBigNext').disabled = !pcNeighbor(1);
}

function pcPaintBig() {
  if (!pcBig || !$('pcBig')?.open) return;
  const doc = pdfs[pcBig.d];
  if (!doc) return;
  let tools = 0, skips = 0, total = 0;
  for (const card of pcCardsOf(doc)) {
    if (!pcBoxesOn(doc, card, pcBig.p).length) continue;
    total += 1;
    const value = pcToolOf(card);
    if (value === 'skip') skips += 1;
    else if (value !== 'none') tools += 1;
  }
  $('pcBigCount').textContent = total ? `本页 ${total} 格：选了修法 ${tools} · 不用处理 ${skips}` : '这一页没有查出切字';
  $('pcBigLayer').querySelectorAll('.pc-box').forEach(el => {
    const card = pcById(el.dataset.pcCard);
    if (!card) return;
    const value = pcToolOf(card);
    el.dataset.state = pcStateOf(card);
    el.classList.toggle('is-sel', pcSel.has(card.id));
    el.querySelector('.pc-tag').textContent = value === 'skip' ? '不用处理' : PC_TOOL_BY_VALUE[value]?.label || '不动';
    const x = el.querySelector('.pc-x');
    x.textContent = value === 'skip' ? '↺' : '✕';
    x.title = value === 'skip' ? '恢复：这格还要处理' : '不用处理：这格其实没问题';
    x.setAttribute('aria-label', x.title);
  });
  if (!$('pcPop').hidden) pcFillPop();
}

function pcCloseBig() {
  pcHidePop();
  pcSel = new Set();
  pcBig = null;
  if ($('pcBig')?.open) $('pcBig').close();
}

function pcEscape() {
  if (pcSel.size) { pcSel = new Set(); pcHidePop(); pcPaintBig(); return; }
  pcCloseBig();
}

// 前后一页：翻到这张表的头 / 尾接着上一张 / 下一张表
function pcNeighbor(step) {
  if (!pcBig) return null;
  let { d, p } = pcBig;
  p += step;
  while (d >= 0 && d < pdfs.length) {
    const doc = pdfs[d];
    if (doc && !doc.error && p >= 0 && p < doc.pages.length) return { d, p };
    d += step;
    const next = pdfs[d];
    if (!next) return null;
    p = step > 0 ? 0 : next.pages.length - 1;
  }
  return null;
}

function pcStep(step) {
  const next = pcNeighbor(step);
  if (next) pcShowPage(next.d, next.p);
}

function pcBigDown(e) {
  if (e.button !== 0 || e.target.closest('[data-pc-x]')) return;
  const box = e.target.closest('.pc-box');
  const ctrl = e.ctrlKey || e.metaKey;
  if (box) {
    pcPress = { id: box.dataset.pcCard, x: e.clientX, y: e.clientY, ctrl };
    return;
  }
  e.preventDefault();
  pcDrag = { x: e.clientX, y: e.clientY, ctrl, moved: false };
}

function pcBigMove(e) {
  if (!pcDrag) return;
  if (!pcDrag.moved && Math.hypot(e.clientX - pcDrag.x, e.clientY - pcDrag.y) < MIN_DRAG_PX) return;
  pcDrag.moved = true;
  const page = $('pcBigPage').getBoundingClientRect();
  const left = Math.min(pcDrag.x, e.clientX) - page.left;
  const top = Math.min(pcDrag.y, e.clientY) - page.top;
  const el = $('pcDragBox');
  el.hidden = false;
  el.style.left = `${left}px`;
  el.style.top = `${top}px`;
  el.style.width = `${Math.abs(e.clientX - pcDrag.x)}px`;
  el.style.height = `${Math.abs(e.clientY - pcDrag.y)}px`;
}

function pcBigUp(e) {
  const press = pcPress;
  const drag = pcDrag;
  pcPress = null;
  pcDrag = null;
  if (press) {
    if (Math.hypot(e.clientX - press.x, e.clientY - press.y) > MIN_DRAG_PX) return;
    if (press.ctrl) {
      if (pcSel.has(press.id)) pcSel.delete(press.id); else pcSel.add(press.id);
    } else pcSel = new Set([press.id]);
    pcAfterSelect();
    return;
  }
  if (!drag) return;
  $('pcDragBox').hidden = true;
  if (!drag.moved) {
    // 点在空白处：取消选择
    if (!drag.ctrl && pcSel.size) { pcSel = new Set(); pcHidePop(); pcPaintBig(); }
    return;
  }
  const x0 = Math.min(drag.x, e.clientX), x1 = Math.max(drag.x, e.clientX);
  const y0 = Math.min(drag.y, e.clientY), y1 = Math.max(drag.y, e.clientY);
  const hit = [...$('pcBigLayer').querySelectorAll('.pc-box')].filter(el => {
    const r = el.getBoundingClientRect();
    return r.right >= x0 && r.left <= x1 && r.bottom >= y0 && r.top <= y1;
  }).map(el => el.dataset.pcCard);
  if (!drag.ctrl) pcSel = new Set();
  hit.forEach(id => pcSel.add(id));
  pcAfterSelect();
}

function pcAfterSelect() {
  pcPaintBig();
  if (pcSel.size) pcFillPop(true); else pcHidePop();
}

function pcHidePop() {
  const pop = $('pcPop');
  if (pop) pop.hidden = true;
}

// 修法框：只放选中的格子都能用的工具；浮在选中的格子旁边（fixed，大图不动）
function pcFillPop(place = false) {
  const pop = $('pcPop');
  const cards = [...pcSel].map(pcById).filter(Boolean);
  if (!cards.length) { pcHidePop(); return; }
  const common = PC_TOOL_DEFS.filter(t => cards.every(c => pcOptions(c).includes(t)));
  const kinds = new Map();
  for (const c of cards) kinds.set(c.type_label, (kinds.get(c.type_label) || 0) + 1);
  const values = new Set(cards.map(pcToolOf));
  const one = cards.length === 1 ? cards[0] : null;
  const doc = pdfs[pcBig?.d];
  const open = one && doc ? excelLinkHtml('📗 在 Excel 打开', { file: doc.wb_name, sheet: one.sheet, cell: (one.cell || '').split(':')[0], folder }, 'pc-chip pc-open') : '';
  const kindText = one ? one.types.join(' + ') : [...kinds].map(([k, n]) => `${k} ${n}`).join(' · ');
  pop.innerHTML = `<p class="pc-pop-title">${one ? `${pcEsc(one.sheet)}!${pcEsc(one.cell)}` : `已选 ${cards.length} 格`}`
    + `<small>${pcEsc(kindText)}</small></p>`
    + (one?.text ? `<p class="pc-pop-text" title="${pcEsc(one.text)}">${pcEsc(one.text)}</p>` : '')
    + (one?.cell_fix ? `<p class="pc-note">用过「${pcEsc(one.cell_fix)}」复检还没过</p>` : '')
    + `<div class="pc-pop-tools">${common.filter(t => t.tool !== 'skip' && t.tool !== 'none').map(t => `<button type="button" class="pc-chip" data-pc-pop="${t.value}" title="${pcEsc(t.tip)}"
        aria-pressed="${values.size === 1 && values.has(t.value)}">${pcEsc(t.label)}</button>`).join('') || '<span class="pc-pop-none">这几格没有共同能用的工具</span>'}</div>`
    + `<div class="pc-pop-tools"><button type="button" class="pc-chip is-skip" data-pc-pop="skip" aria-pressed="${values.size === 1 && values.has('skip')}" title="${pcEsc(PC_TOOL_BY_VALUE.skip.tip)}">✕ 不用处理</button>`
    + `<button type="button" class="pc-chip" data-pc-pop="none" aria-pressed="${values.size === 1 && values.has('none')}" title="${pcEsc(PC_TOOL_BY_VALUE.none.tip)}">不动</button>${open}`
    + '<button type="button" class="pc-chip pc-chip-plain" data-pc-pop="clear">取消选择</button></div>';
  pop.hidden = false;
  if (place) pcPlacePop();
}

function pcPlacePop() {
  const pop = $('pcPop');
  const rects = [...$('pcBigLayer').querySelectorAll('.pc-box.is-sel')].map(el => el.getBoundingClientRect());
  if (!rects.length) return;
  const right = Math.max(...rects.map(r => r.right));
  const left = Math.min(...rects.map(r => r.left));
  const top = Math.min(...rects.map(r => r.top));
  const box = pop.getBoundingClientRect();
  const gap = 14;
  let x = right + gap;
  if (x + box.width > window.innerWidth - 8) x = left - gap - box.width;
  if (x < 8) x = Math.max(8, window.innerWidth - box.width - 8);
  const y = Math.min(Math.max(70, top), window.innerHeight - box.height - 8);
  pop.style.left = `${x}px`;
  pop.style.top = `${y}px`;
}

function pcPopClick(e) {
  const button = e.target.closest('[data-pc-pop]');
  if (!button) return;
  const value = button.dataset.pcPop;
  if (value === 'clear') { pcSel = new Set(); pcHidePop(); pcPaintBig(); return; }
  const cards = [...pcSel].map(pcById).filter(Boolean);
  pcSetTool(cards, value);
  pcPaint();
}

// 右键格子：用 Excel 打开那一格（和格子条上点格子名一样）
function pcBigContext(e) {
  const box = e.target.closest('.pc-box');
  if (!box) return;
  e.preventDefault();
  const card = pcById(box.dataset.pcCard);
  const doc = pdfs[pcBig?.d];
  if (!card || !doc) return;
  const holder = $('pcBigLinks');
  holder.innerHTML = excelLinkHtml(pcEsc(card.cell), { file: doc.wb_name, sheet: card.sheet, cell: (card.cell || '').split(':')[0], folder });
  holder.querySelector('a').click();
}

/* ---------- 页面上的点击 ---------- */

document.addEventListener('click', e => {
  const tool = e.target.closest('[data-pc-tool]');
  if (tool && !tool.closest('#pcPop')) {
    const card = pcById(tool.dataset.pcCard);
    if (card) {
      if (tool.dataset.pcTool === 'skip' && pcToolOf(card) === 'skip') pcToggleSkip(card);
      else pcSetTool([card], tool.dataset.pcTool);
      pcPaint();
    }
    return;
  }
  const bulk = e.target.closest('[data-pc-bulk-tool]');
  if (bulk) {
    pcSetTool(pcCategoryCards(), bulk.dataset.pcBulkTool);
    pcPaint();
    return;
  }
  const more = e.target.closest('[data-pc-more]');
  if (more) {
    const doc = pdfs[Number(more.dataset.pcMore)];
    if (doc) { pcExpanded.add(doc.pdf); pcRedraw(); }
    return;
  }
  if (e.target.closest('[data-pc-check]')) { pcAutoCheck(true); return; }
  if (e.target.closest('[data-pc-pause]')) { pcPauseScan(); return; }
  if (e.target.closest('[data-pc-resume]')) pcResumeScan();
});
