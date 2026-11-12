// 按分类人工处理（用户 2026-11-03：留这个排版，去掉「学一下 / 按规律修」）。
// 左边按问题分类，右边这一类的工具箱（cell_toolbox.js：选工具 → 底部统一修改）、打开 / 复检按钮和格子截图。点分类只换颜色，不挪位置、不自动滚动。
// 左上角可切换「按文件」：一个文件打开一次，把里面所有问题改完再复检，免得同一个文件在几个分类里反复打开。
// 依赖 batch.js 里的 batchState、$、esc、startJob、openReady、requestBusy、busyJobName、openBlocked；open_excel.js 的 excelLinkHtml。
let categorySelected = null;
let categoryRenderKey = '';
const CATEGORY_SELECTED_KEY = 'clear_handle_excel.learnCategory';
const CATEGORY_MODE_KEY = 'clear_handle_excel.categoryMode';
let categoryMode = 'category'; // 有了工具箱（2026-11-08）默认按分类；切到按文件会记住
let fileSelected = null;
try {
  categorySelected = localStorage.getItem(CATEGORY_SELECTED_KEY);
  if (localStorage.getItem(CATEGORY_MODE_KEY) === 'file') categoryMode = 'file';
} catch (_) {}

const shortLabel = label => label.split(' · ')[0];

// 按「字丢没丢」分级（用户 2026-11-10）：红 = 字丢了必须改；灰 = PDF 里没对上文字，扫一眼；黄 = 被裁，字都在，可以不管（压线 2026-11-12 起算红）
const SEVERITY_RANK = { red: 0, gray: 1, yellow: 2 };
const SEVERITY_HEADS = {
  red: ['🔴 必须改：PDF 上看得出的毛病', '压线（半截字）、被裁、只印出前半段、###、被分页切开'],
  gray: ['⚪ 扫一眼：PDF 里没对上文字', '多半是符号、格式造成的误报'],
  yellow: ['🟡 可以不管：被裁', '字超出格子被挡住（多半是隐藏了列），字还在 PDF 里'],
};
const sevRank = card => SEVERITY_RANK[card.severity] ?? 0;
const groupSeverity = cards => cards.reduce((best, c) => (sevRank(c) < SEVERITY_RANK[best] ? c.severity : best), 'yellow');

function fileList() {
  const groups = new Map();
  for (const card of pendingCards()) {
    if (!groups.has(card.wb_path)) groups.set(card.wb_path, { key: card.wb_path, label: card.wb_name, type: 'file', cards: [] });
    groups.get(card.wb_path).cards.push(card);
  }
  return [...groups.values()].map(g => {
    const kinds = new Map();
    for (const c of g.cards) { const k = shortLabel((c.category || ['', '其他'])[1]); kinds.set(k, (kinds.get(k) || 0) + 1); }
    const red = g.cards.filter(c => (c.severity || 'red') === 'red').length;
    return { ...g, count: g.cards.length, files: 1, red,
      summary: [...kinds].sort((a, b) => b[1] - a[1]).map(([k, n]) => `${k} ${n}`).join(' · ') };
  }).sort((a, b) => b.red - a.red || b.count - a.count || a.label.localeCompare(b.label));
}

function groupList() {
  const opened = openedGroup();
  return [...(opened ? [opened] : []), ...(categoryMode === 'file' ? fileList() : categoryList())];
}
function groupCurrent(groups) {
  const want = categoryMode === 'file' ? fileSelected : categorySelected;
  return groups.find(g => g.key === want) || groups.find(g => g.key !== OPENED_KEY) || groups[0] || null;
}

// 已打开、还没复检的文件：改完之前别的分类里先不列它，免得同一个文件反复打开；复检完回到各分类
const OPENED_KEY = '__opened__';
const openedSet = () => new Set(batchState.opened_files || []);
const pendingCards = () => { const opened = openedSet(); return (batchState.cards || []).filter(c => !opened.has(c.wb_path)); };

function openedGroup() {
  const opened = openedSet();
  const cards = (batchState.cards || []).filter(c => opened.has(c.wb_path));
  if (!cards.length) return null;
  return { key: OPENED_KEY, label: '已打开、待复检', type: OPENED_KEY, cards, count: cards.length,
    files: new Set(cards.map(c => c.wb_path)).size };
}

function categoryList() {
  const order = (batchState.error_types || []).map(t => t.code);
  const groups = new Map();
  for (const card of pendingCards()) {
    const [key, label] = card.category || ['other', '其他'];
    if (!groups.has(key)) groups.set(key, { key, label, type: key.split('|')[0], cards: [] });
    groups.get(key).cards.push(card);
  }
  const rank = type => (order.indexOf(type) < 0 ? order.length : order.indexOf(type));
  return [...groups.values()].map(g => ({ ...g, count: g.cards.length, files: new Set(g.cards.map(c => c.wb_path)).size, severity: groupSeverity(g.cards) }))
    .sort((a, b) => SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity] || rank(a.type) - rank(b.type) || b.count - a.count || a.label.localeCompare(b.label));
}

function categoryCurrent(cats) {
  return cats.find(c => c.key === categorySelected) || cats[0] || null;
}

function categoryBusyReason() {
  if (requestBusy) return '正在提交请求';
  return batchState.busy ? `后台正在${busyJobName()}，结束后再点` : '';
}

function categoryOpenBlockReason(cards) {
  if (requestBusy) return '正在提交请求';
  if (openChoosing || (batchState.busy && batchState.job?.kind === 'batch_open')) return '正在打开文件';
  if (batchState.busy && cards.length && cards.every(c => openBlocked(c))) return openBlocked(cards[0]);
  return '';
}

function categoryRecheckBlockReason(cards) {
  if (requestBusy) return '正在提交请求';
  const ready = recheckAllCards(cards);
  return recheckBlockReason(ready.length ? ready : cards);
}

const yellowCards = cat => cat.cards.filter(c => c.color === 'yellow' && c.finalized);

function fileListHtml(files, current) {
  return files.map(f => `<li><button type="button" class="learn-cat learn-file" data-learn-cat="${esc(f.key)}" aria-pressed="${f === current}" title="${esc(f.key)}">`
    + `<span class="learn-cat-name">${esc(f.label)}<small class="learn-file-kinds">${f.red ? `<b class="learn-file-red">必须改 ${f.red}</b> · ` : ''}${esc(f.summary)}</small></span>`
    + `<span class="learn-count">${f.count}</span></button></li>`).join('');
}

function openedItemHtml(group, current) {
  return `<li><button type="button" class="learn-cat learn-opened" data-learn-cat="${OPENED_KEY}" aria-pressed="${group === current}" title="打开过、还没复检的文件，改完保存后复检">`
    + `<span class="learn-cat-name">📂 已打开、待复检<small class="learn-file-kinds">${group.files} 个文件，复检后还有问题的回到下面</small></span>`
    + `<span class="learn-count">${group.files}</span></button></li>`;
}

function categoryListHtml(all, current) {
  if (!all.length) return '<p class="learn-empty">没有要人工处理的问题。先在上面扫描一个文件夹。</p>';
  const opened = all[0]?.key === OPENED_KEY ? openedItemHtml(all[0], current) : '';
  const cats = opened ? all.slice(1) : all;
  return opened + (!cats.length ? '<p class="learn-empty">其余的文件都已经打开过了，改完复检。</p>'
    : categoryMode === 'file' ? fileListHtml(cats, current) : categoryListOnlyHtml(cats, current));
}

function categoryListOnlyHtml(cats, current) {
  let lastType = '', lastSeverity = '';
  return cats.map(cat => {
    const typeLabel = cat.label.split(' · ')[0];
    const color = typeColorOf(cat.type);
    let sev = '';
    if (cat.severity !== lastSeverity) {
      const [title, note] = SEVERITY_HEADS[cat.severity] || SEVERITY_HEADS.red;
      sev = `<li class="learn-sev learn-sev-${esc(cat.severity)}" aria-hidden="true">${esc(title)}<small>${esc(note)}</small></li>`;
      lastSeverity = cat.severity;
      lastType = '';
    }
    const head = sev + (cat.type !== lastType ? `<li class="learn-type" aria-hidden="true"><span class="learn-type-chip" style="background:${color.fill};color:${color.font}" title="Excel 里这种问题的标记颜色">${esc(typeLabel)}</span></li>` : '');
    lastType = cat.type;
    const rest = cat.label.split(' · ').slice(1).join(' · ') || typeLabel;
    return `${head}<li><button type="button" class="learn-cat" data-learn-cat="${esc(cat.key)}" aria-pressed="${cat === current}">`
      + `<span class="learn-cat-name">${esc(rest)}</span><span class="learn-count">${cat.count}</span></button></li>`;
  }).join('');
}

// 打开 / 复检两个按钮放一行：主要靠上面的工具，工具解决不了的再打开 Excel 自己改
function categoryActionsHtml(cat, cards) {
  const disabled = reason => reason ? ` disabled title="${esc(reason)}"` : '';
  const yellow = cards.length;
  const openBusy = categoryOpenBlockReason(cards);
  const recheckBusy = categoryRecheckBlockReason(cards);
  const opened = cat.key === OPENED_KEY;
  const openText = categoryMode === 'file' ? '📂 打开这个文件' : '📂 打开这一类的文件（每次最多 10 个）';
  const recheckText = opened ? `🔁 复检这 ${cat.files} 个文件` : categoryMode === 'file' ? '🔁 复检这个文件' : `🔁 复检这一类（${cat.files} 个文件）`;
  return '<div class="tb-actions">'
    + (opened ? '' : `<button type="button" class="btn btn-outline btn-sm" id="categoryOpenBtn"${disabled(openBusy || (!yellow && '没有可打开的文件'))}>${openText}</button>`)
    + `<button type="button" class="btn btn-outline btn-sm" id="categoryRecheckBtn"${disabled(recheckBusy || (!yellow && '没有要复检的格子'))}>${recheckText}</button>`
    + `<button type="button" class="btn btn-outline btn-sm tb-next-btn" id="categoryNextBtn"${disabled(!cat.cards.length && '没有格子')} title="在 Excel 里跳到清单里的下一个问题格（文件没开会先打开），不用自己找标色">➡ 下一处</button>`
    + `<label class="tb-auto-recheck" title="打开过的文件在 Excel 里改完、保存、关掉后，自动复检这一份（Excel 还开着的不动）"><input type="checkbox" id="autoRecheckToggle"${autoRecheck ? ' checked' : ''}> 改完关掉文件后自动复检</label>`
    + `<span class="tb-actions-hint">${opened ? '这些文件打开过：在 Excel 里改完保存，再复检；复检后还有问题的回到各分类。' : '工具解决不了的，打开 Excel 自己改，保存后复检（复检只检测、不改格子）。单击清单里的格子名直接跳过去。'}</span></div>`;
}

function categoryDetailHtml(cat) {
  if (!cat) return '';
  const cards = yellowCards(cat);
  const fixed = cat.cards.filter(c => c.cell_fix).length;
  const title = cat.key === OPENED_KEY ? '已打开、待复检' : cat.label;
  const meta = cat.key === OPENED_KEY ? `${cat.files} 个文件 · ${cat.count} 格`
    : categoryMode === 'file' ? `${cat.count} 格 · ${esc(cat.summary)}` : `${cat.count} 格 · ${cat.files} 个文件`;
  return `<div class="learn-head"><h3>${esc(title)}</h3><span class="learn-meta">${meta}`
    + (fixed ? ` · ${fixed} 格用过工具还没过` : '') + '</span></div>'
    + toolboxGroupBarHtml(cat) + categoryActionsHtml(cat, cards)
    + toolboxCellsHtml(cat, categoryMode === 'category' || cat.key === OPENED_KEY);
}

function renderCategories() {
  const panel = $('learnPanel');
  if (!panel) return;
  const cats = groupList();
  const current = groupCurrent(cats);
  renderLowRisk();
  const key = JSON.stringify([categoryMode, batchState.opened_files, cats.map(c => [c.key, c.cards.map(x => [x.id, x.color, x.severity, x.finalized, x.cell_fix?.label, x.default_tool, (x.tools || []).join()])]), current?.key,
    categoryBusyReason(), openChoosing, batchState.job?.kind, batchState.job?.closing,
    batchState.folder, excelLinkHtml.revision, batchState.open_blocked]);
  renderOpenNext();
  maybeAutoRecheck();
  if (key === categoryRenderKey) { toolboxAfterRender(); return; }
  categoryRenderKey = key;
  // 重画时保留清单框的滚动位置
  const scroll = $('learnDetail').querySelector('.tb-grid')?.scrollTop || 0;
  const rest = cats.filter(c => c.key !== OPENED_KEY);
  $('learnSummary').textContent = cats.length ? `共 ${cats.reduce((n, c) => n + c.count, 0)} 格，`
    + (categoryMode === 'file' ? `${rest.length} 个文件` : `分成 ${rest.length} 类`)
    + (rest.length < cats.length ? `，${cats[0].files} 个文件待复检` : '') : '';
  for (const b of document.querySelectorAll('[data-learn-mode]')) b.setAttribute('aria-pressed', String(b.dataset.learnMode === categoryMode));
  $('learnList').innerHTML = categoryListHtml(cats, current);
  $('learnDetail').innerHTML = categoryDetailHtml(current);
  const wrap = $('learnDetail').querySelector('.tb-grid');
  if (wrap && `${categoryMode}:${current?.key}` === renderCategories.lastKey) wrap.scrollTop = scroll;
  renderCategories.lastKey = `${categoryMode}:${current?.key}`;
  toolboxAfterRender();
}

// 「也列出被裁」：黄色默认不列、不标色、不算待处理；打开后马上列出来，Excel 标色下次打开 / 复检时跟上
function renderLowRisk() {
  const box = $('lowRiskToggle');
  if (!box) return;
  const show = !!batchState.show_low_risk;
  if (document.activeElement !== box || !box.dataset.saving) box.checked = show;
  const hidden = batchState.low_risk_hidden || 0;
  $('lowRiskCount').textContent = show ? '（已列出）' : hidden ? `（藏了 ${hidden} 格）` : '';
}

$('lowRiskToggle')?.addEventListener('change', async event => {
  const box = event.target;
  box.dataset.saving = '1';
  try {
    const res = await fetch('/api/batch/show_low_risk', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ show: box.checked }) });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || '没改成');
    batchState.show_low_risk = data.show;
    lastEtag = ''; await poll();
  } catch (err) {
    box.checked = !box.checked;
    appAlert(`没改成：${err.message || err}`);
  } finally {
    delete box.dataset.saving;
    renderLowRisk();
  }
});

// 「➡ 下一处」：按清单顺序点下一个格子链接（open_excel.js 打开并跳过去、给链接加框），从上次点过的那个往下数，到底再从头。
// 不滚动页面（用户讨厌页面自己动），清单里加框的就是现在跳到的那格
function jumpNextCell() {
  const links = [...$('learnDetail').querySelectorAll('.tb-grid .excel-link:not(.is-blocked)')];
  if (!links.length) return appAlert('这一组没有能跳过去的格子（文件可能还在后台处理）');
  const last = links.findIndex(a => a.classList.contains('is-last'));
  links[(last + 1) % links.length].click();
}

// 改完关掉文件后自动复检（用户 2026-11-10：人工改的过程要丝滑）：后台发现打开过的文件保存过、Excel 已经关掉，
// 就自动排一次复检，不用回网页点。每份文件这次触发一次；后台在忙、任务待续跑时先不动，空了再看
const AUTO_RECHECK_KEY = 'clear_handle_excel.autoRecheck';
let autoRecheck = true;
try { if (localStorage.getItem(AUTO_RECHECK_KEY) === '0') autoRecheck = false; } catch (_) {}
const autoRechecked = new Set();
let autoRecheckBusy = false;
async function maybeAutoRecheck() {
  const ready = new Set(batchState.recheck_ready || []);
  for (const path of [...autoRechecked]) if (!ready.has(path)) autoRechecked.delete(path);  // 复检完或又打开了：下次保存关掉再触发
  if (!autoRecheck || autoRecheckBusy || requestBusy || openChoosing || batchState.busy
      || batchState.job?.status === 'interrupted') return;
  const paths = [...ready].filter(p => !autoRechecked.has(p));
  if (!paths.length) return;
  const want = new Set(paths);
  const cards = recheckAllCards(manualCards().filter(c => want.has(c.wb_path)));
  if (!cards.length) return;
  paths.forEach(p => autoRechecked.add(p));
  autoRecheckBusy = true;
  try { await startJob('/api/batch/recheck', { card_ids: cards.map(c => c.id) }); }
  finally { autoRecheckBusy = false; }
}

// 顶上的「打开下一批」：按左边列表的顺序（按文件时问题多的在前），打开还没打开过的文件，每次最多 10 个
function nextBatchCards() {
  return groupList().filter(g => g.key !== OPENED_KEY).flatMap(g => yellowCards(g));
}

function renderOpenNext() {
  const button = $('categoryOpenNextBtn');
  if (!button) return;
  const cards = nextBatchCards();
  const files = new Set(cards.map(c => c.wb_path)).size;
  const available = new Set(cards.filter(c => !batchState.busy || !openBlocked(c)).map(c => c.wb_path)).size;
  const busy = categoryOpenBlockReason(cards);
  button.textContent = files ? `📂 打开下一批（还有 ${files} 个文件，当前可打开 ${available} 个）` : '📂 文件都打开过了';
  button.disabled = !!busy || !files;
  button.title = busy || '按左边列表的顺序，打开还没打开过的文件，每次最多 10 个；打开的放进「已打开、待复检」';
}

async function categoryOpen() {
  const cat = groupCurrent(groupList());
  if (cat) await openReady(false, yellowCards(cat).map(c => c.id));
}

async function categoryRecheck() {
  const cat = groupCurrent(groupList());
  if (!cat) return;
  const all = yellowCards(cat);
  const cards = recheckAllCards(all);
  if (!cards.length) return appAlert(categoryRecheckBlockReason(all) || '没有要复检的格子');
  const waiting = new Set(all.filter(c => !cards.includes(c)).map(c => c.wb_path)).size;
  const what = cat.key === OPENED_KEY ? `已打开的 ${cat.files} 个文件` : categoryMode === 'file' ? `「${cat.label}」` : `「${cat.label}」涉及的 ${cat.files} 个文件`;
  if (!(await appConfirm(`复检${what}：整份重新检测，通过的移出列表、去掉标色。${queueNote('复检')}\n请先在 Excel 里保存修改。${waiting ? `\n还有 ${waiting} 个文件待后台处理，本次不加入复检。` : ''}`,
    { title: cat.key === OPENED_KEY || categoryMode === 'file' ? '复检文件' : '复检这一类', okText: '🔁 复检', tone: 'info' }))) return;
  await startJob('/api/batch/recheck', { card_ids: cards.map(c => c.id) });
}

$('categoryOpenNextBtn')?.addEventListener('click', () => {
  const ids = nextBatchCards().map(c => c.id);
  if (ids.length) openReady(false, ids);
});
$('learnList').addEventListener('click', event => {
  const button = event.target.closest('[data-learn-cat]');
  if (!button) return;
  if (button.dataset.learnCat === OPENED_KEY) { fileSelected = categorySelected = OPENED_KEY; }
  else if (categoryMode === 'file') fileSelected = button.dataset.learnCat;
  else {
    categorySelected = button.dataset.learnCat;
    try { localStorage.setItem(CATEGORY_SELECTED_KEY, categorySelected); } catch (_) {}
  }
  renderCategories();
});
for (const button of document.querySelectorAll('[data-learn-mode]')) button.addEventListener('click', () => {
  categoryMode = button.dataset.learnMode;
  try { localStorage.setItem(CATEGORY_MODE_KEY, categoryMode); } catch (_) {}
  renderCategories();
});
$('learnDetail').addEventListener('click', event => {
  if (toolboxHandleClick(event)) return;
  const id = event.target.closest('button')?.id;
  if (id === 'categoryOpenBtn') categoryOpen();
  else if (id === 'categoryRecheckBtn') categoryRecheck();
  else if (id === 'categoryNextBtn') jumpNextCell();
});
$('learnDetail').addEventListener('change', event => {
  if (event.target.id !== 'autoRecheckToggle') return;
  autoRecheck = event.target.checked;
  try { localStorage.setItem(AUTO_RECHECK_KEY, autoRecheck ? '1' : '0'); } catch (_) {}
  maybeAutoRecheck();
});
renderCategories();
