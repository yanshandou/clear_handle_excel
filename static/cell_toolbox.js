// 查看切字的单元格工具箱（用户 2026-11-08 要，照「调整页面」的样子）。
// 「人工处理」面板里每个分类（或文件）上面一排工具按钮，点一下这一类都用这个工具；每个格子下面也能另选。
// 最底下「🛠 统一修改」：后台按选好的工具改 Excel、保存，再整份复检。通过的移出列表、去掉标色；
// 没过的留在原分类，写明用过什么工具，默认不再选（不自动升一级）。
// 点按钮只换颜色，不挪位置、不自动滚动。依赖 batch.js 的 batchState、$、esc、startJob、requestBusy；
// category.js 调用这里的函数画清单。
const TOOLBOX_KEY = 'clear_handle_excel.cellTools';
const TOOLBOX_PAGE = 60; // 每类先画这么多格，多的点「再显示」接着往下画
const TOOL_DEFS = [
  { value: 'grow:1.2', tool: 'grow', ratio: 1.2, label: '行高 ×1.2', tip: '开自动换行，先补够放下文字要的高度，再加「字号 × 1.2」' },
  { value: 'grow:1.5', tool: 'grow', ratio: 1.5, label: '行高 ×1.5', tip: '开自动换行，先补够放下文字要的高度，再加「字号 × 1.5」' },
  { value: 'grow:2', tool: 'grow', ratio: 2, label: '行高 ×2', tip: '开自动换行，先补够放下文字要的高度，再加「字号 × 2」' },
  { value: 'shrink', tool: 'shrink', label: '缩小字体', tip: '缩小字体填充：字缩到放得下（会关掉自动换行）' },
  { value: 'break', tool: 'break', label: '上面加分页线', tip: '在这格上面加一条手动分页线，整格挪到下一页' },
  { value: 'none', tool: 'none', label: '不动', tip: '这格不用工具，留给你打开 Excel 自己改' },
];
const TOOL_BY_VALUE = Object.fromEntries(TOOL_DEFS.map(t => [t.value, t]));
const toolboxImages = new Map(); // 卡片 id → 截图地址（'' = 没有截图）
const toolboxShown = new Map(); // 分组 → 已经画了几格
let toolboxChoices = {};
let toolboxWatching = false; // 这个页面点过统一修改：做完弹一次结果
try { toolboxChoices = JSON.parse(localStorage.getItem(TOOLBOX_KEY) || '{}') || {}; } catch (_) {}

function toolboxFolderChoices() {
  const key = batchState.folder || '';
  if (!toolboxChoices[key] || typeof toolboxChoices[key] !== 'object') toolboxChoices[key] = {};
  return toolboxChoices[key];
}
function toolboxSave() {
  try { localStorage.setItem(TOOLBOX_KEY, JSON.stringify(toolboxChoices)); } catch (_) {}
}
const toolboxEligible = card => card.color === 'yellow' && card.finalized && !card.dismissed;
const toolOptions = card => TOOL_DEFS.filter(t => t.tool === 'none' || (card.tools || []).includes(t.tool));
function toolOf(card) {
  if (!toolboxEligible(card)) return 'none';
  const chosen = toolboxFolderChoices()[card.id];
  const value = chosen ?? card.default_tool ?? 'none';
  const def = TOOL_BY_VALUE[value];
  return def && (def.tool === 'none' || (card.tools || []).includes(def.tool)) ? value : 'none';
}
function setTool(cards, value) {
  const choices = toolboxFolderChoices();
  const def = TOOL_BY_VALUE[value];
  for (const card of cards) {
    if (!toolboxEligible(card)) continue;
    if (def.tool !== 'none' && !(card.tools || []).includes(def.tool)) continue;
    choices[card.id] = value;
  }
  toolboxSave();
}

function toolChipsHtml(card) {
  if (!toolboxEligible(card)) return '<span class="tb-none">已标色，不用工具</span>';
  const current = toolOf(card);
  return toolOptions(card).map(t => `<button type="button" class="tb-chip" data-tb-tool="${t.value}" data-tb-kind="${t.tool}" aria-pressed="${t.value === current}" title="${esc(t.tip)}">${esc(t.label)}</button>`).join('');
}

// 分组上面那一排：点一下这一组能用这个工具的格子都选上
function toolboxGroupBarHtml(group) {
  const cards = group.cards.filter(toolboxEligible);
  if (!cards.length) return '';
  const usable = TOOL_DEFS.filter(t => t.tool === 'none' || cards.some(c => (c.tools || []).includes(t.tool)));
  if (usable.length <= 1) {
    return '<div class="tb-group-bar"><span class="tb-group-label">这一类没有工具</span><span class="tb-group-hint">PDF 未匹配到文字等只能打开 Excel 自己看，改完复检</span></div>';
  }
  const who = group.type === 'file' ? '这个文件' : group.key === OPENED_KEY ? '这些文件' : '这一类';
  return `<div class="tb-group-bar" data-tb-group><span class="tb-group-label">${who}都用：</span>`
    + usable.map(t => `<button type="button" class="tb-chip tb-chip-lg" data-tb-group-tool="${t.value}" data-tb-kind="${t.tool}" aria-pressed="${cards.every(c => toolOf(c) === t.value || !toolOptions(c).some(o => o.value === t.value))}" title="${esc(t.tip)}">${esc(t.label)}</button>`).join('')
    // 只修这一类（用户 2026-11-09 要）：只把这一组里选了工具的格子交给后台，别的分类不动
    + `<button type="button" class="btn btn-primary btn-sm tb-group-apply" data-tb-group-apply><span>🛠 只修${who}</span><span class="tb-count" data-tb-group-count>0</span></button>`
    + `<span class="tb-group-hint">单个格子在下面另选。点「🛠 只修${who}」只改${who}；最下面的「🛠 统一修改」一次改所有分类里选了工具的格子</span></div>`;
}

function toolboxCellHtml(card, showFile) {
  const jump = { file: card.wb_name, folder: batchState.folder, sheet: card.sheet_name, cell: (card.cell_addr || '').split(':')[0], blocked: openBlocked(card) };
  const where = `${esc(card.sheet_name)}!${esc(card.cell_addr)}`;
  const image = toolboxImages.get(card.id);
  const shot = image ? `<img src="${esc(image)}" alt="${where} 的 PDF 截图" loading="lazy">`
    : `<span class="tb-shot-empty">${image === '' ? '没有截图' : '截图加载中…'}</span>`;
  const fix = card.cell_fix ? `<p class="tb-note" title="${esc(card.message || '')}">用过「${esc(card.cell_fix.label)}」还没过</p>` : '';
  const kinds = (card.category || ['', ''])[1];
  return `<div class="tb-cell" data-tb-card="${esc(card.id)}">`
    + `<button type="button" class="tb-shot" data-tb-shot title="单击看大图；Ctrl + 单击用 Excel 打开这一格">${shot}</button>`
    + `<div class="tb-cell-info"><span class="tb-where">${excelLinkHtml(where, jump)}</span>`
    + (showFile ? `<span class="tb-file" title="${esc(card.wb_name)}">${esc(card.wb_name)}</span>` : `<span class="tb-file" title="${esc(kinds)}">${esc(kinds)}</span>`)
    + `<span class="tb-text" title="${esc(card.cell_text ?? card.text ?? '')}">${esc(card.cell_text ?? card.text ?? '')}</span></div>`
    + fix + `<div class="tb-tools" role="group" aria-label="这格用什么工具">${toolChipsHtml(card)}</div></div>`;
}

function toolboxCellsHtml(group, showFile) {
  const cards = showFile ? group.cards : [...group.cards].sort((a, b) => (a.category?.[1] || '').localeCompare(b.category?.[1] || '') || a.seq - b.seq);
  const limit = toolboxShown.get(group.key) || TOOLBOX_PAGE;
  const more = cards.length - limit;
  return `<div class="tb-grid">${cards.slice(0, limit).map(c => toolboxCellHtml(c, showFile)).join('')}</div>`
    + (more > 0 ? `<p class="tb-more"><button type="button" class="btn btn-outline btn-sm" data-tb-more="${esc(group.key)}">再显示 ${Math.min(more, TOOLBOX_PAGE)} 格（还有 ${more} 格）</button></p>` : '');
}

// 只改按钮颜色，不重画整块（截图、滚动位置都不动）
function toolboxRefreshPressed() {
  const detail = $('learnDetail');
  const cards = new Map((batchState.cards || []).map(c => [c.id, c]));
  for (const el of detail.querySelectorAll('[data-tb-card]')) {
    const card = cards.get(el.dataset.tbCard);
    if (!card) continue;
    const current = toolOf(card);
    for (const chip of el.querySelectorAll('[data-tb-tool]')) chip.setAttribute('aria-pressed', String(chip.dataset.tbTool === current));
  }
  const group = groupCurrent(groupList());
  for (const chip of detail.querySelectorAll('[data-tb-group-tool]')) {
    const value = chip.dataset.tbGroupTool;
    const eligible = (group?.cards || []).filter(toolboxEligible);
    chip.setAttribute('aria-pressed', String(eligible.length > 0 && eligible.every(c => toolOf(c) === value || !toolOptions(c).some(o => o.value === value))));
  }
  toolboxRenderBar();
}

async function toolboxLoadImages() {
  const ids = [...$('learnDetail').querySelectorAll('[data-tb-card]')].map(el => el.dataset.tbCard).filter(id => !toolboxImages.has(id));
  if (!ids.length) return;
  for (const id of ids) toolboxImages.set(id, undefined);
  try {
    const response = await fetch('/api/batch/card_images', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ card_ids: ids }) });
    const result = await response.json();
    if (!response.ok || !result.ok) throw new Error(result.error || `HTTP ${response.status}`);
    for (const id of ids) {
      const shots = result.images?.[id] || {};
      toolboxImages.set(id, shots.before || shots.after || '');
    }
  } catch (error) {
    for (const id of ids) toolboxImages.delete(id);
    console.warn('工具箱截图没取到', error);
    return;
  }
  for (const el of $('learnDetail').querySelectorAll('[data-tb-card]')) {
    const url = toolboxImages.get(el.dataset.tbCard);
    const box = el.querySelector('[data-tb-shot]');
    if (url === undefined || !box || box.querySelector('img')) continue;
    box.innerHTML = url ? `<img src="${esc(url)}" alt="PDF 截图" loading="lazy">` : '<span class="tb-shot-empty">没有截图</span>';
  }
}

// 底部悬浮条：这个文件夹里所有选了工具的格子
function toolboxPlan(cards = batchState.cards || []) {
  return cards.filter(toolboxEligible).map(c => [c, toolOf(c)]).filter(([, v]) => v !== 'none');
}
function toolboxBusyReason() {
  if (requestBusy) return '正在提交请求';
  return batchState.busy ? `后台正在${busyJobName()}，结束后再点` : '';
}
function toolboxRenderBar() {
  const bar = $('tbBar');
  if (!bar) return;
  const plan = toolboxPlan();
  const counts = new Map();
  for (const [, v] of plan) counts.set(v, (counts.get(v) || 0) + 1);
  const files = new Set(plan.map(([c]) => c.wb_path)).size;
  $('tbBarText').textContent = plan.length
    ? `${plan.length} 格 · ${files} 个文件：` + TOOL_DEFS.filter(t => counts.has(t.value)).map(t => `${t.label} ${counts.get(t.value)} 格`).join(' · ')
    : '还没有选工具的格子：在上面的分类或格子上点工具';
  $('tbCount').textContent = plan.length;
  const busy = toolboxBusyReason();
  const button = $('tbApply');
  button.disabled = !!busy || !plan.length;
  button.title = busy || (plan.length ? '后台按选好的工具改 Excel、保存，再整份复检' : '先选工具');
  // 分类上的「只修这一类」：数字、能不能点跟着这一组选了几格变
  const groupButton = document.querySelector('[data-tb-group-apply]');
  if (groupButton) {
    const mine = toolboxPlan(groupCurrent(groupList())?.cards || []);
    groupButton.querySelector('[data-tb-group-count]').textContent = mine.length;
    groupButton.disabled = !!busy || !mine.length;
    groupButton.title = busy || (mine.length ? `只把这里选了工具的 ${mine.length} 格交给后台改、保存，再整份复检；别的分类不动` : '这里还没有选工具的格子');
  }
}

// group：只修这一类（这个文件）时传当前分组；不传 = 底部统一修改，所有分类一起改
async function toolboxApply(group = null) {
  const plan = toolboxPlan(group ? group.cards : undefined);
  if (!plan.length) return;
  const files = new Set(plan.map(([c]) => c.wb_path)).size;
  const counts = new Map();
  for (const [, v] of plan) counts.set(v, (counts.get(v) || 0) + 1);
  const lines = TOOL_DEFS.filter(t => counts.has(t.value)).map(t => `· ${t.label}：${counts.get(t.value)} 格`).join('\n');
  const scope = !group ? '' : group.key === OPENED_KEY ? '「已打开、待复检」' : `「${group.label}」`;
  const others = group ? toolboxPlan().length - plan.length : 0;
  console.info('🛠 单元格工具箱提交', { 范围: scope || '全部分类', 格: plan.length, 文件: files, 别的分类没动: others });
  if (!(await appConfirm(`${group ? `只修${scope}` : '统一修改'} ${plan.length} 格（${files} 个文件）：\n${lines}${others > 0 ? `\n\n别的分类里选了工具的 ${others} 格这次不动。` : ''}\n\n后台改 Excel、保存，再整份复检：通过的移出列表、去掉标色；没过的留在原来的分类，写明用过什么工具，可以换个工具再改。\n加了行高可能多出一页，做完会告诉你哪些表页数变了。\n请先在 Excel 里保存并关闭这些文件。`,
    { title: group ? `🛠 只修${scope}` : '🛠 统一修改', okText: '🛠 开始修改', tone: 'info' }))) return;
  const tools = plan.map(([card, value]) => {
    const def = TOOL_BY_VALUE[value];
    return def.tool === 'grow' ? { card_id: card.id, tool: 'grow', ratio: def.ratio } : { card_id: card.id, tool: def.tool };
  });
  if (batchState.job?.status === 'interrupted'
      && !(await appConfirm('这会替换上次任务的续跑点。已保存的修复和审核标记会保留，继续吗？', { okText: '继续' }))) return;
  let sent = false;
  await action(async () => { await post('/api/batch/fix', { tools }); sent = true; });
  if (sent) {
    // 提交了：这些格子的选择清掉，复检后没过的默认「不动」，不会再按原工具改一遍
    const choices = toolboxFolderChoices();
    for (const t of tools) choices[t.card_id] = 'none';
    toolboxSave();
    toolboxWatching = true;
  }
}

function toolboxResultText(job) {
  const results = Object.values(job?.result || {});
  const sum = key => results.reduce((n, r) => n + (r[key] || 0), 0);
  const lines = [`改了 ${sum('done')} 格，复检通过 ${sum('passed')} 格，还没过 ${sum('left')} 格。`];
  const pages = results.flatMap(r => (r.pages || []).map(p => `· ${r.name} / ${p.sheet}：${p.before} 页 → ${p.after} 页`));
  if (pages.length) lines.push('', '页数变了（可以去「调整页面」看要不要取整）：', ...pages.slice(0, 15), ...(pages.length > 15 ? [`……还有 ${pages.length - 15} 张表`] : []));
  const skipped = results.flatMap(r => (r.skipped || []).map(s => `· ${r.name} / ${s.where}：${s.reason}`));
  if (skipped.length) lines.push('', '没改的格子：', ...skipped.slice(0, 10), ...(skipped.length > 10 ? [`……还有 ${skipped.length - 10} 格`] : []));
  const errors = results.flatMap(r => r.error ? [`· ${r.name}：${r.error}`] : (r.errors || []).map(e => `· ${r.name} / ${e.where}：${e.reason}`));
  if (errors.length) lines.push('', '出错的：', ...errors.slice(0, 10), ...(errors.length > 10 ? [`……还有 ${errors.length - 10} 处`] : []));
  return lines.join('\n');
}

function toolboxAfterRender() {
  toolboxRenderBar();
  toolboxLoadImages();
  if (toolboxWatching && !batchState.busy && batchState.job?.kind === 'batch_fix') {
    // 结果写在右下角的实时进度框里（每份改了几格、过了几格、页数变了没），不再另外弹框（用户 2026-11-09）
    toolboxWatching = false;
    console.info('🛠 统一修改完成', toolboxResultText(batchState.job));
  }
}

function toolboxShowLarge(card) {
  const url = toolboxImages.get(card.id);
  if (!url) return;
  const dialog = $('tbLarge');
  $('tbLargeTitle').textContent = `${card.wb_name} / ${card.sheet_name}!${card.cell_addr}`;
  $('tbLargeImg').src = url;
  dialog.showModal();
}

function toolboxHandleClick(event) {
  const chip = event.target.closest('[data-tb-tool]');
  const cards = batchState.cards || [];
  if (chip) {
    const card = cards.find(c => c.id === chip.closest('[data-tb-card]').dataset.tbCard);
    if (card) { setTool([card], chip.dataset.tbTool); toolboxRefreshPressed(); }
    return true;
  }
  if (event.target.closest('[data-tb-group-apply]')) {
    const group = groupCurrent(groupList());
    if (group) toolboxApply(group);
    return true;
  }
  const groupChip = event.target.closest('[data-tb-group-tool]');
  if (groupChip) {
    const group = groupCurrent(groupList());
    if (group) { setTool(group.cards, groupChip.dataset.tbGroupTool); toolboxRefreshPressed(); }
    return true;
  }
  const more = event.target.closest('[data-tb-more]');
  if (more) {
    const key = more.dataset.tbMore;
    toolboxShown.set(key, (toolboxShown.get(key) || TOOLBOX_PAGE) + TOOLBOX_PAGE);
    categoryRenderKey = '';
    renderCategories();
    return true;
  }
  const shot = event.target.closest('[data-tb-shot]');
  if (shot) {
    const el = shot.closest('[data-tb-card]');
    const card = cards.find(c => c.id === el.dataset.tbCard);
    if (!card) return true;
    if (event.ctrlKey || event.metaKey) el.querySelector('.tb-where .excel-link')?.click();
    else toolboxShowLarge(card);
    return true;
  }
  return false;
}

$('tbApply')?.addEventListener('click', () => toolboxApply());
$('tbLargeClose')?.addEventListener('click', () => $('tbLarge').close());
$('tbLarge')?.addEventListener('click', event => { if (event.target === $('tbLarge')) $('tbLarge').close(); });
