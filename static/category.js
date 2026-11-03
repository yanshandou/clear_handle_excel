// 按分类人工处理（用户 2026-11-03：留这个排版，去掉「学一下 / 按规律修」）。
// 左边按问题分类，右边这一类的两步（打开文件改 → 复检这一类）和格子清单。点分类只换颜色，不挪位置、不自动滚动。
// 左上角可切换「按文件」：一个文件打开一次，把里面所有问题改完再复检，免得同一个文件在几个分类里反复打开。
// 依赖 batch.js 里的 batchState、$、esc、startJob、openReady、requestBusy、busyJobName、openBlocked；open_excel.js 的 excelLinkHtml。
let categorySelected = null;
let categoryRenderKey = '';
const CATEGORY_CELL_LIMIT = 300; // 清单最多画这么多行，太多只是拖慢页面
const CATEGORY_SELECTED_KEY = 'clear_handle_excel.learnCategory';
const CATEGORY_MODE_KEY = 'clear_handle_excel.categoryMode';
let categoryMode = 'file'; // 平时按文件，一个文件只开一次
let fileSelected = null;
try {
  categorySelected = localStorage.getItem(CATEGORY_SELECTED_KEY);
  if (localStorage.getItem(CATEGORY_MODE_KEY) === 'category') categoryMode = 'category';
} catch (_) {}

const shortLabel = label => label.split(' · ')[0];

function fileList() {
  const groups = new Map();
  for (const card of pendingCards()) {
    if (!groups.has(card.wb_path)) groups.set(card.wb_path, { key: card.wb_path, label: card.wb_name, type: 'file', cards: [] });
    groups.get(card.wb_path).cards.push(card);
  }
  return [...groups.values()].map(g => {
    const kinds = new Map();
    for (const c of g.cards) { const k = shortLabel((c.category || ['', '其他'])[1]); kinds.set(k, (kinds.get(k) || 0) + 1); }
    return { ...g, count: g.cards.length, files: 1,
      summary: [...kinds].sort((a, b) => b[1] - a[1]).map(([k, n]) => `${k} ${n}`).join(' · ') };
  }).sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
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
  return [...groups.values()].map(g => ({ ...g, count: g.cards.length, files: new Set(g.cards.map(c => c.wb_path)).size }))
    .sort((a, b) => rank(a.type) - rank(b.type) || b.count - a.count || a.label.localeCompare(b.label));
}

function categoryCurrent(cats) {
  return cats.find(c => c.key === categorySelected) || cats[0] || null;
}

function categoryBusyReason() {
  if (requestBusy) return '正在提交请求';
  return batchState.busy ? `后台正在${busyJobName()}，结束后再点` : '';
}

const yellowCards = cat => cat.cards.filter(c => c.color === 'yellow' && c.finalized);

function fileListHtml(files, current) {
  return files.map(f => `<li><button type="button" class="learn-cat learn-file" data-learn-cat="${esc(f.key)}" aria-pressed="${f === current}" title="${esc(f.key)}">`
    + `<span class="learn-cat-name">${esc(f.label)}<small class="learn-file-kinds">${esc(f.summary)}</small></span>`
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
  let lastType = '';
  return cats.map(cat => {
    const typeLabel = cat.label.split(' · ')[0];
    const color = typeColorOf(cat.type);
    const head = cat.type !== lastType ? `<li class="learn-type" aria-hidden="true"><span class="learn-type-chip" style="background:${color.fill};color:${color.font}" title="Excel 里这种问题的标记颜色">${esc(typeLabel)}</span></li>` : '';
    lastType = cat.type;
    const rest = cat.label.split(' · ').slice(1).join(' · ') || typeLabel;
    return `${head}<li><button type="button" class="learn-cat" data-learn-cat="${esc(cat.key)}" aria-pressed="${cat === current}">`
      + `<span class="learn-cat-name">${esc(rest)}</span><span class="learn-count">${cat.count}</span></button></li>`;
  }).join('');
}

function categoryState(card) {
  if (card.color !== 'yellow') return ['marked', card.color === 'blue' ? '已标蓝' : '已标绿'];
  if (card.grown && !card.grown.error) return ['grown', '加高后没过'];
  return ['todo', '待改'];
}

function categoryCellsHtml(cat) {
  const byFile = categoryMode === 'file' && cat.key !== OPENED_KEY ? null : new Map();
  if (byFile) for (const c of pendingCards()) byFile.set(c.wb_path, (byFile.get(c.wb_path) || 0) + 1);
  const cards = byFile ? cat.cards : [...cat.cards].sort((a, b) => (a.category?.[1] || '').localeCompare(b.category?.[1] || '') || a.seq - b.seq);
  const rows = cards.slice(0, CATEGORY_CELL_LIMIT).map(c => {
    const [state, label] = categoryState(c);
    const jump = { file: c.wb_name, folder: batchState.folder, sheet: c.sheet_name, cell: (c.cell_addr || '').split(':')[0], blocked: openBlocked(c) };
    const where = `${esc(c.sheet_name)}!${esc(c.cell_addr)}`;
    let second;
    if (byFile) {
      // 这个文件在别的分类里还有几处：打开时顺手一起改，复检是整份查的，会一起出结果
      const others = cat.key === OPENED_KEY ? 0 : byFile.get(c.wb_path) - cat.cards.filter(x => x.wb_path === c.wb_path).length;
      second = excelLinkHtml(esc(c.wb_name), jump) + (others > 0 ? ` <span class="learn-others" title="Excel 里这些格子也标了色，打开时可以顺手一起改">还有 ${others} 处别的问题</span>` : '');
    } else second = esc((c.category || ['', ''])[1]);
    return `<tr data-state="${state}"><td><span class="learn-state">${label}</span></td>`
      + `<td>${second}</td><td>${byFile ? where : excelLinkHtml(where, jump)}</td><td class="learn-text">${esc(c.cell_text ?? c.text)}</td></tr>`;
  }).join('');
  const more = cards.length > CATEGORY_CELL_LIMIT ? `<p class="toolbar-hint">只列出前 ${CATEGORY_CELL_LIMIT} 格，还有 ${cards.length - CATEGORY_CELL_LIMIT} 格。</p>` : '';
  return `<div class="learn-table-wrap"><table class="learn-table"><thead><tr><th>状态</th><th>${byFile ? '文件' : '分类'}</th><th>格子</th><th>内容</th></tr></thead><tbody>${rows}</tbody></table></div>${more}`;
}

// 收紧行高：改的时候先给问题行一个很大的行高，保存后让程序量出要多高、收回多余的，再整份复检
function tightenButtonHtml(disabledAttr) {
  return `<button type="button" class="btn btn-outline" id="categoryTightenBtn"${disabledAttr} title="改的时候可以先给问题行一个很大的行高（例如 200），保存后点这里：程序量出每行要多高，收到刚好够再留「字号 × 倍数」的余量（只收不加），然后整份复检；复检没过的行退回你原来给的行高">↕️ 收紧行高并复检</button>`;
}

function categoryDetailHtml(cat) {
  if (!cat) return '';
  const busy = categoryBusyReason();
  const yellow = yellowCards(cat).length;
  const grown = cat.cards.filter(c => categoryState(c)[0] === 'grown').length;
  const disabled = reason => reason ? ` disabled title="${esc(reason)}"` : '';
  if (cat.key === OPENED_KEY) {
    return `<div class="learn-head"><h3>已打开、待复检</h3><span class="learn-meta">${cat.files} 个文件 · ${cat.count} 格</span></div>`
      + '<ol class="learn-steps">'
      + `<li><span class="learn-step-no">✓</span><div><strong>这些文件已经打开过</strong><p>在 Excel 里改完保存，再复检。复检前别的分类里不再列这些文件，免得重复打开；复检后还有问题的会回到各分类。想重新打开，点清单里的文件名。行高不好估的，可以先给一个很大的行高（例如 200），保存后点「收紧行高并复检」，多余的会收回来。</p>`
      + `<button type="button" class="btn btn-primary" id="categoryRecheckBtn"${disabled(busy || (!yellow && '没有要复检的格子'))}>🔁 复检这 ${cat.files} 个文件</button> ${tightenButtonHtml(disabled(busy || (!yellow && '没有要复检的格子')))}</div></li>`
      + '</ol>' + categoryCellsHtml(cat);
  }
  if (categoryMode === 'file') {
    return `<div class="learn-head"><h3>${esc(cat.label)}</h3><span class="learn-meta">${cat.count} 格 · ${esc(cat.summary)}</span></div>`
      + '<ol class="learn-steps">'
      + `<li><span class="learn-step-no">1</span><div><strong>打开改</strong><p>打开这个文件，把下面的问题一次改完（Excel 里都标了色），保存。单击清单里的格子可以直接跳过去。</p>`
      + `<button type="button" class="btn btn-outline" id="categoryOpenBtn"${disabled(busy || (!yellow && '这个文件没有要改的格子'))}>📂 打开这个文件</button></div></li>`
      + `<li><span class="learn-step-no">2</span><div><strong>复检</strong><p>改完保存后复检这个文件：整份重新检测，通过的移出列表、去掉标色；还有问题的留在这里继续改。行高不好估的，可以先给一个很大的行高（例如 200），保存后点「收紧行高并复检」，多余的会收回来。</p>`
      + `<button type="button" class="btn btn-primary" id="categoryRecheckBtn"${disabled(busy || (!yellow && '这个文件没有要复检的格子'))}>🔁 复检这个文件</button> ${tightenButtonHtml(disabled(busy || (!yellow && '这个文件没有要复检的格子')))}</div></li>`
      + '</ol>' + categoryCellsHtml(cat);
  }
  return `<div class="learn-head"><h3>${esc(cat.label)}</h3><span class="learn-meta">${cat.count} 格 · ${cat.files} 个文件`
    + (grown ? ` · ${grown} 格扫描后已开换行加高、复检没过` : '') + '</span></div>'
    + '<ol class="learn-steps">'
    + `<li><span class="learn-step-no">1</span><div><strong>打开改</strong><p>打开这一类的文件，在 Excel 里改好这些格子（行高、列宽、换行、缩小填充、字号、插分页都行），保存。单击清单里的文件名可以直接跳到那一格。</p>`
    + `<button type="button" class="btn btn-outline" id="categoryOpenBtn"${disabled(busy || (!yellow && '这一类没有可打开的文件'))}>📂 打开这一类的文件（每次最多 10 个）</button></div></li>`
    + `<li><span class="learn-step-no">2</span><div><strong>复检</strong><p>改完保存后复检这一类涉及的文件：整份重新检测，通过的移出列表、去掉标色；还有问题的留在这里继续改。同一份文件里别的分类的格子也会一起出结果。</p>`
    + `<button type="button" class="btn btn-primary" id="categoryRecheckBtn"${disabled(busy || (!yellow && '这一类没有要复检的格子'))}>🔁 复检这一类（${cat.files} 个文件）</button> ${tightenButtonHtml(disabled(busy || (!yellow && '这一类没有要复检的格子')))}</div></li>`
    + '</ol>' + categoryCellsHtml(cat);
}

function renderCategories() {
  const panel = $('learnPanel');
  if (!panel) return;
  const cats = groupList();
  const current = groupCurrent(cats);
  const key = JSON.stringify([categoryMode, batchState.opened_files, cats.map(c => [c.key, c.cards.map(x => [x.id, x.color, x.finalized, !!x.grown])]), current?.key,
    categoryBusyReason(), batchState.folder, excelLinkHtml.revision, batchState.open_blocked]);
  if (key === categoryRenderKey) return;
  categoryRenderKey = key;
  // 重画时保留清单框的滚动位置
  const scroll = $('learnDetail').querySelector('.learn-table-wrap')?.scrollTop || 0;
  const rest = cats.filter(c => c.key !== OPENED_KEY);
  $('learnSummary').textContent = cats.length ? `共 ${cats.reduce((n, c) => n + c.count, 0)} 格，`
    + (categoryMode === 'file' ? `${rest.length} 个文件` : `分成 ${rest.length} 类`)
    + (rest.length < cats.length ? `，${cats[0].files} 个文件待复检` : '') : '';
  for (const b of document.querySelectorAll('[data-learn-mode]')) b.setAttribute('aria-pressed', String(b.dataset.learnMode === categoryMode));
  $('learnList').innerHTML = categoryListHtml(cats, current);
  $('learnDetail').innerHTML = categoryDetailHtml(current);
  const wrap = $('learnDetail').querySelector('.learn-table-wrap');
  if (wrap && `${categoryMode}:${current?.key}` === renderCategories.lastKey) wrap.scrollTop = scroll;
  renderCategories.lastKey = `${categoryMode}:${current?.key}`;
}

async function categoryOpen() {
  const cat = groupCurrent(groupList());
  if (cat) await openReady(false, yellowCards(cat).map(c => c.id));
}

async function categoryRecheck() {
  const cat = groupCurrent(groupList());
  if (!cat) return;
  const what = cat.key === OPENED_KEY ? `已打开的 ${cat.files} 个文件` : categoryMode === 'file' ? `「${cat.label}」` : `「${cat.label}」涉及的 ${cat.files} 个文件`;
  if (!(await appConfirm(`复检${what}：整份重新检测，通过的移出列表、去掉标色。\n请先在 Excel 里保存修改。`,
    { title: cat.key === OPENED_KEY || categoryMode === 'file' ? '复检文件' : '复检这一类', okText: '🔁 复检', tone: 'info' }))) return;
  await startJob('/api/batch/recheck', { card_ids: yellowCards(cat).map(c => c.id) });
}

async function categoryTighten() {
  const cat = groupCurrent(groupList());
  if (!cat) return;
  const what = cat.key === OPENED_KEY ? `已打开的 ${cat.files} 个文件` : categoryMode === 'file' ? `「${cat.label}」` : `「${cat.label}」涉及的 ${cat.files} 个文件`;
  const ratio = document.getElementById('growRatio')?.value || '1.5';
  if (!(await appConfirm(`收紧${what}里出过切字问题的行：\n· 量出每行的字全部显示要多高，再留「字号 × ${ratio}」的余量，比这个高的收回来，只收不加\n· 收完整份复检，复检没过的行退回原来的行高\n· 有图片、竖排字的行不动\n请先在 Excel 里保存修改。`,
    { title: '收紧行高并复检', okText: '↕️ 收紧并复检', tone: 'info' }))) return;
  await startJob('/api/batch/tighten', { card_ids: yellowCards(cat).map(c => c.id), grow_ratio: Number(ratio) });
}

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
  const id = event.target.closest('button')?.id;
  if (id === 'categoryOpenBtn') categoryOpen();
  else if (id === 'categoryRecheckBtn') categoryRecheck();
  else if (id === 'categoryTightenBtn') categoryTighten();
});
renderCategories();
