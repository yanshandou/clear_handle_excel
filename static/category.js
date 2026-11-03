// 按分类人工处理（用户 2026-11-03：留这个排版，去掉「学一下 / 按规律修」）。
// 左边按问题分类，右边这一类的两步（打开文件改 → 复检这一类）和格子清单。点分类只换颜色，不挪位置、不自动滚动。
// 依赖 batch.js 里的 batchState、$、esc、startJob、openReady、requestBusy、busyJobName、openBlocked；open_excel.js 的 excelLinkHtml。
let categorySelected = null;
let categoryRenderKey = '';
const CATEGORY_CELL_LIMIT = 300; // 清单最多画这么多行，太多只是拖慢页面
const CATEGORY_SELECTED_KEY = 'clear_handle_excel.learnCategory';
try { categorySelected = localStorage.getItem(CATEGORY_SELECTED_KEY); } catch (_) {}

function categoryList() {
  const order = (batchState.error_types || []).map(t => t.code);
  const groups = new Map();
  for (const card of batchState.cards || []) {
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

function categoryListHtml(cats, current) {
  if (!cats.length) return '<p class="learn-empty">没有要人工处理的问题。先在上面扫描一个文件夹。</p>';
  let lastType = '';
  return cats.map(cat => {
    const typeLabel = cat.label.split(' · ')[0];
    const head = cat.type !== lastType ? `<li class="learn-type" aria-hidden="true">${esc(typeLabel)}</li>` : '';
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
  const rows = cat.cards.slice(0, CATEGORY_CELL_LIMIT).map(c => {
    const [state, label] = categoryState(c);
    const link = excelLinkHtml(esc(c.wb_name), { file: c.wb_name, folder: batchState.folder, sheet: c.sheet_name,
      cell: (c.cell_addr || '').split(':')[0], blocked: openBlocked(c) });
    return `<tr data-state="${state}"><td><span class="learn-state">${label}</span></td>`
      + `<td>${link}</td><td>${esc(c.sheet_name)}!${esc(c.cell_addr)}</td><td class="learn-text">${esc(c.cell_text ?? c.text)}</td></tr>`;
  }).join('');
  const more = cat.cards.length > CATEGORY_CELL_LIMIT ? `<p class="toolbar-hint">只列出前 ${CATEGORY_CELL_LIMIT} 格，还有 ${cat.cards.length - CATEGORY_CELL_LIMIT} 格。</p>` : '';
  return `<div class="learn-table-wrap"><table class="learn-table"><thead><tr><th>状态</th><th>文件</th><th>格子</th><th>内容</th></tr></thead><tbody>${rows}</tbody></table></div>${more}`;
}

function categoryDetailHtml(cat) {
  if (!cat) return '';
  const busy = categoryBusyReason();
  const yellow = yellowCards(cat).length;
  const grown = cat.cards.filter(c => categoryState(c)[0] === 'grown').length;
  const disabled = reason => reason ? ` disabled title="${esc(reason)}"` : '';
  return `<div class="learn-head"><h3>${esc(cat.label)}</h3><span class="learn-meta">${cat.count} 格 · ${cat.files} 个文件`
    + (grown ? ` · ${grown} 格扫描后已开换行加高、复检没过` : '') + '</span></div>'
    + '<ol class="learn-steps">'
    + `<li><span class="learn-step-no">1</span><div><strong>打开改</strong><p>打开这一类的文件，在 Excel 里改好这些格子（行高、列宽、换行、缩小填充、字号、插分页都行），保存。单击清单里的文件名可以直接跳到那一格。</p>`
    + `<button type="button" class="btn btn-outline" id="categoryOpenBtn"${disabled(busy || (!yellow && '这一类没有可打开的文件'))}>📂 打开这一类的文件（每次最多 10 个）</button></div></li>`
    + `<li><span class="learn-step-no">2</span><div><strong>复检</strong><p>改完保存后复检这一类涉及的文件：整份重新检测，通过的移出列表、去掉标色；还有问题的留在这里继续改。同一份文件里别的分类的格子也会一起出结果。</p>`
    + `<button type="button" class="btn btn-primary" id="categoryRecheckBtn"${disabled(busy || (!yellow && '这一类没有要复检的格子'))}>🔁 复检这一类（${cat.files} 个文件）</button></div></li>`
    + '</ol>' + categoryCellsHtml(cat);
}

function renderCategories() {
  const panel = $('learnPanel');
  if (!panel) return;
  const cats = categoryList();
  const current = categoryCurrent(cats);
  const key = JSON.stringify([cats.map(c => [c.key, c.cards.map(x => [x.id, x.color, x.finalized, !!x.grown])]), current?.key,
    categoryBusyReason(), batchState.folder, excelLinkHtml.revision, batchState.open_blocked]);
  if (key === categoryRenderKey) return;
  categoryRenderKey = key;
  // 重画时保留清单框的滚动位置
  const scroll = $('learnDetail').querySelector('.learn-table-wrap')?.scrollTop || 0;
  $('learnSummary').textContent = cats.length ? `共 ${cats.reduce((n, c) => n + c.count, 0)} 格，分成 ${cats.length} 类` : '';
  $('learnList').innerHTML = categoryListHtml(cats, current);
  $('learnDetail').innerHTML = categoryDetailHtml(current);
  const wrap = $('learnDetail').querySelector('.learn-table-wrap');
  if (wrap && current?.key === renderCategories.lastKey) wrap.scrollTop = scroll;
  renderCategories.lastKey = current?.key;
}

async function categoryOpen() {
  const cat = categoryCurrent(categoryList());
  if (cat) await openReady(false, yellowCards(cat).map(c => c.id));
}

async function categoryRecheck() {
  const cat = categoryCurrent(categoryList());
  if (!cat) return;
  if (!(await appConfirm(`复检「${cat.label}」涉及的 ${cat.files} 个文件：整份重新检测，通过的移出列表、去掉标色。\n请先在 Excel 里保存修改。`,
    { title: '复检这一类', okText: '🔁 复检', tone: 'info' }))) return;
  await startJob('/api/batch/recheck', { card_ids: yellowCards(cat).map(c => c.id) });
}

$('learnList').addEventListener('click', event => {
  const button = event.target.closest('[data-learn-cat]');
  if (!button) return;
  categorySelected = button.dataset.learnCat;
  try { localStorage.setItem(CATEGORY_SELECTED_KEY, categorySelected); } catch (_) {}
  renderCategories();
});
$('learnDetail').addEventListener('click', event => {
  const id = event.target.closest('button')?.id;
  if (id === 'categoryOpenBtn') categoryOpen();
  else if (id === 'categoryRecheckBtn') categoryRecheck();
});
renderCategories();
