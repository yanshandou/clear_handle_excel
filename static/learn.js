// 按分类示范学习修复：左边分类列表，右边这一类的三步（示范 → 学一下 → 按规律修）和格子清单。
// 依赖 batch.js 里的 batchState、$、esc、post、startJob、openReady、requestBusy、busyJobName。
let learnSelected = null;
let learnRenderKey = '';
const LEARN_CELL_LIMIT = 300; // 清单最多画这么多行，太多只是拖慢页面
const LEARN_STATE = {
  demo: ['示范', '你改过的格子，学的时候用来总结规律'],
  applied: ['已按规律改', '按规律改过、复检还没过：可以在 Excel 里再改，当作新的示范'],
  todo: ['待处理', ''],
  new: ['没记原样', '以前扫描的格子没记下原样：点「学一下」时先记下现在的样子，下一次才能当示范'],
};
const LEARN_SELECTED_KEY = 'clear_handle_excel.learnCategory';
try { learnSelected = localStorage.getItem(LEARN_SELECTED_KEY); } catch (_) {}

function learnCategories() { return batchState.categories || []; }

function learnCurrent() {
  const cats = learnCategories();
  return cats.find(c => c.key === learnSelected) || cats[0] || null;
}

function learnBusyReason() {
  if (requestBusy) return '正在提交请求';
  return batchState.busy ? `后台正在${busyJobName()}，结束后再点` : '';
}

function learnListHtml(cats, current) {
  if (!cats.length) return '<p class="learn-empty">还没有要处理的问题。先在上面扫描一个文件夹。</p>';
  let lastType = '';
  return cats.map(cat => {
    const typeLabel = cat.label.split(' · ')[0];
    const color = typeColorOf(cat.type);
    const many = String(cat.type).includes('+') ? '多种问题：' : '';
    const head = cat.type !== lastType ? `<li class="learn-type" aria-hidden="true"><span class="learn-type-chip" style="background:${color.fill};color:${color.font}" title="Excel 里这种问题的标记颜色">${esc(many + typeLabel)}</span></li>` : '';
    lastType = cat.type;
    const rest = cat.label.split(' · ').slice(1).join(' · ') || typeLabel;
    const done = learnDone(cat);
    const badge = cat.rule?.ok && !done ? '<span class="learn-badge" title="已经学到规律">🎓</span>' : '';
    const mark = done ? `<span class="learn-done" title="${esc(done)}" aria-label="已改完">✓</span>` : '';
    return `${head}<li><button type="button" class="learn-cat${done ? ' is-done' : ''}" data-learn-cat="${esc(cat.key)}" aria-pressed="${cat === current}">`
      + `<span class="learn-cat-name">${esc(rest)}</span>${badge}<span class="learn-count">${cat.count}</span>${mark}</button></li>`;
  }).join('');
}

// 这一类改完了没有：学到了规律，而且剩下的格子都是示范或已按规律改过（没有待处理、没记原样的）。
// 改完的格子复检通过后移出列表；整类都通过，这一类就不再显示。返回悬停说明，没改完返回空字符串
function learnDone(cat) {
  if (!cat.rule?.ok || !cat.cards.length) return '';
  if (cat.cards.some(c => c.state !== 'demo' && c.state !== 'applied')) return '';
  const demo = cat.cards.filter(c => c.state === 'demo').length;
  return `这一类已经改完：示范 ${demo} 格，按规律改 ${cat.cards.length - demo} 格。复检通过的会移出列表`;
}

function learnRuleHtml(rule) {
  if (!rule) return '<p class="learn-rule learn-rule-none">还没学过。改好几个示范、保存后点「🎓 学一下」。</p>';
  const demos = (rule.demos || []).map(d => `<li>${esc(d.label)}：${esc((d.change || []).join('，'))}</li>`).join('');
  const conflicts = (rule.conflicts || []).map(c => `<li>${esc(c.label)}：${esc(c.reason)}</li>`).join('');
  return `<div class="learn-rule ${rule.ok ? 'is-ok' : 'is-bad'}" role="status">`
    + `<p class="learn-rule-text">${rule.ok ? `学到的规律：<strong>${esc(rule.text)}</strong>（看了 ${rule.demo_count} 个示范）` : esc(rule.text)}</p>`
    + (demos ? `<details><summary>示范了哪些格（${rule.demos.length}）</summary><ul>${demos}</ul></details>` : '')
    + (conflicts ? `<details open><summary>⚠️ 对不上的示范（${rule.conflicts.length}）</summary><ul>${conflicts}</ul></details>` : '')
    + '</div>';
}

function learnCellsHtml(cat) {
  const rows = cat.cards.slice(0, LEARN_CELL_LIMIT).map(c => {
    const [label, tip] = LEARN_STATE[c.state] || LEARN_STATE.todo;
    const link = excelLinkHtml(esc(c.wb_name), { file: c.wb_name, folder: batchState.folder, sheet: c.sheet_name,
      cell: (c.cell_addr || '').split(':')[0], blocked: openBlocked(c) });
    return `<tr data-state="${esc(c.state)}"><td><span class="learn-state" title="${esc(tip)}">${label}</span></td>`
      + `<td>${link}</td><td>${esc(c.sheet_name)}!${esc(c.cell_addr)}</td><td class="learn-text">${esc(c.text)}</td></tr>`;
  }).join('');
  const more = cat.cards.length > LEARN_CELL_LIMIT ? `<p class="toolbar-hint">只列出前 ${LEARN_CELL_LIMIT} 格，还有 ${cat.cards.length - LEARN_CELL_LIMIT} 格。</p>` : '';
  return `<div class="learn-table-wrap"><table class="learn-table"><thead><tr><th>状态</th><th>文件</th><th>格子</th><th>内容</th></tr></thead><tbody>${rows}</tbody></table></div>${more}`;
}

function learnDetailHtml(cat) {
  if (!cat) return '';
  const count = state => cat.cards.filter(c => c.state === state).length;
  const busy = learnBusyReason();
  const todo = count('todo') + count('applied');
  const ok = cat.rule?.ok;
  const openIds = cat.cards.filter(c => c.color === 'yellow' && c.finalized).length;
  const disabled = reason => reason ? ` disabled title="${esc(reason)}"` : '';
  const fresh = count('new');
  return `<div class="learn-head"><h3>${esc(cat.label)}</h3><span class="learn-meta">${cat.count} 格 · ${cat.files} 个文件`
    + (count('demo') ? ` · 示范 ${count('demo')} 格` : '') + '</span></div>'
    + '<ol class="learn-steps">'
    + `<li><span class="learn-step-no">1</span><div><strong>示范</strong><p>在 Excel 里把这一类改好几格（建议 3～10 格，挑字号、长短不一样的），保存。改什么都行：行高、列宽、换行、缩小填充、字号、插分页。</p>`
    + `<button type="button" class="btn btn-outline" id="learnOpenBtn"${disabled(busy || (!openIds && '这一类没有可打开的文件'))}>📂 打开这一类的文件（每次最多 10 个）</button></div></li>`
    + `<li><span class="learn-step-no">2</span><div><strong>学一下</strong><p>程序对比你改的格子和扫描时的原样，看你改了什么、改了多少，写成一句话。</p>`
    + `<button type="button" class="btn btn-primary" id="learnBtn"${disabled(busy)}>🎓 学一下</button>`
    + (fresh ? `<p class="toolbar-hint">${fresh} 格是以前扫描的，没记下原样：这次先记下现在的样子，改好后再点一次才算示范。</p>` : '')
    + `${learnRuleHtml(cat.rule)}</div></li>`
    + `<li><span class="learn-step-no">3</span><div><strong>按规律修</strong><p>按学到的规律改这一类剩下的格子（示范的不动），保存后整份复检。通过的移出列表；没通过的留下，可以再示范几个、再学一次。</p>`
    + `<button type="button" class="btn btn-success" id="learnApplyBtn"${disabled(busy || (!ok && '还没学到规律') || (!todo && '没有剩下要改的格子'))}>⚡ 按规律修剩下的 ${todo} 格</button></div></li>`
    + '</ol>' + learnCellsHtml(cat);
}

function renderLearn() {
  const panel = $('learnPanel');
  if (!panel) return;
  const cats = learnCategories();
  const current = learnCurrent();
  const key = JSON.stringify([cats, current?.key, learnBusyReason(), batchState.folder, excelLinkHtml.revision, batchState.open_blocked, typeColors]);
  if (key === learnRenderKey) return;
  learnRenderKey = key;
  $('learnSummary').textContent = cats.length
    ? `共 ${cats.reduce((n, c) => n + c.count, 0)} 格，分成 ${cats.length} 类` : '';
  $('learnList').innerHTML = learnListHtml(cats, current);
  $('learnDetail').innerHTML = learnDetailHtml(current);
}

async function learnOpen() {
  const cat = learnCurrent();
  if (!cat) return;
  await openReady(false, cat.cards.filter(c => c.color === 'yellow' && c.finalized).map(c => c.id));
}

async function learnRun() {
  const cat = learnCurrent();
  if (!cat) return;
  if (!(await appConfirm(`看「${cat.label}」这一类的 ${cat.count} 格（${cat.files} 个文件），找出你改过的格子，总结改法。\n只读文件、不改文件。请先在 Excel 里保存你的修改。`,
    { title: '学一下', okText: '🎓 学一下', tone: 'info' }))) return;
  await startJob('/api/batch/learn', { category: cat.key });
}

async function learnApply() {
  const cat = learnCurrent();
  if (!cat?.rule?.ok) return;
  const todo = cat.cards.filter(c => c.state === 'todo' || c.state === 'applied').length;
  const demos = cat.cards.filter(c => c.state === 'demo').length;
  const conflicts = cat.rule.conflicts?.length ? `\n\n注意：有 ${cat.rule.conflicts.length} 个示范和规律对不上，建议先看一眼。` : '';
  if (!(await appConfirm(`按这条规律改「${cat.label}」剩下的 ${todo} 格：\n${cat.rule.text}\n\n示范的 ${demos} 格不动；你改过但没当成示范的格子也不动。改完保存原 Excel，再整份复检。请先保留原件，并在 Excel 里保存这些文件。${conflicts}`,
    { title: '按规律修', okText: `⚡ 修改 ${todo} 格`, tone: 'info' }))) return;
  await startJob('/api/batch/apply', { category: cat.key });
}

$('learnList').addEventListener('click', event => {
  const button = event.target.closest('[data-learn-cat]');
  if (!button) return;
  learnSelected = button.dataset.learnCat;
  try { localStorage.setItem(LEARN_SELECTED_KEY, learnSelected); } catch (_) {}
  renderLearn();
});
$('learnDetail').addEventListener('click', event => {
  const id = event.target.closest('button')?.id;
  if (id === 'learnOpenBtn') learnOpen();
  else if (id === 'learnBtn') learnRun();
  else if (id === 'learnApplyBtn') learnApply();
});
renderLearn();
