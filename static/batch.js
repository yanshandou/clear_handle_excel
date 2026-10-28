'use strict';
const REVIEW_COLORS = ['blue', 'green', 'yellow'];
const REVIEW_LABELS = { blue: '蓝 · 已修复', green: '绿 · 还原并忽略', yellow: '黄 · 人工处理' };
const REVIEW_JOB_NAMES = { batch_commit: '提交审核', batch_recheck: '复检', batch_open: '打开文件', batch_unhighlight: '清除标记颜色' };
let batchState = { cards: [], busy: false, job: null, progress: {}, logs: [] };
let selectedFilter = 'all';
let largeMode = false;
let activeCardId = null;
let cardPointer = null;
let previewSourceImage = null;
let leftAltHandled = false;
let requestBusy = false;
let openChoosing = false;
let pollVersion = 0;
let lastEtag = '';
let pollInFlight = 0;
let renderKey = '';
const cardHtmlCache = new Map();
const headHtmlCache = new Map();
const cardMemo = new Map();
const headMemo = new Map();
const cardElements = new Map();
const fileElements = new Map();
let largeCardKey = '';
let lastActiveIndex = 0;
let selectedIds = new Set();
const selectedFiles = new Set();
const expandedFiles = new Set();
const fileMountLimits = new Map();
const fileMountFrames = new Map();
const FILE_MOUNT_BATCH = 18;
const CARD_DETAIL_FIELDS = ['text', 'problems', 'fix_images', 'context', 'images', 'message'];
let lastSnapshotUrl = '';
let fileFolder = null;
let selectNotice = '';
// 筛选后改了颜色的卡片继续留在当前筛选里，不立刻消失；重新点筛选按钮才按新颜色过滤
let filterKeep = new Set();
const $ = id => document.getElementById(id);
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const nextColor = (color, direction = 1) => REVIEW_COLORS[(REVIEW_COLORS.indexOf(color) + direction + 3) % 3];
// 页面按文件分组：一个文件一个框，文件按第一次出现（处理顺序）排，框里的卡片按处理顺序
const fileKey = card => card.wb_path || card.wb_name || '';
// 后台任务在跑时这份文件能不能点开 Excel：不能时是原因（后端按任务进度算），能打开是空字符串
const openBlocked = card => (batchState.open_blocked || {})[card.wb_path] || '';
// 自动处理 / 复检在跑时：不在这次任务里的文件、已经做完的文件可以改颜色（只写记录，不碰 Excel）
const markWhileBusy = () => ['batch_scan', 'batch_recheck'].includes(batchState.job?.kind) && Array.isArray(batchState.job?.workbooks);
const canMark = card => !batchState.busy || (markWhileBusy() && !openBlocked(card));
// 「🤖 自动调整」在跑时还能把别的文件排进去，做完当前文件接着做；
// 「只扫描」还没扫完时也能排，全部文件扫完后同一个任务接着自动调整
const autoQueueOpen = () => {
  const job = batchState.job;
  return !!(batchState.busy && job?.kind === 'batch_scan' && (job.options?.picked || job.options?.scan_only)
    && Array.isArray(job.workbooks) && !job.closing && !job.pause_requested && !job.to_manual);
};
// 后台在跑的任务叫什么：按钮不可点时写进提示
function busyJobName() {
  const job = batchState.job || {};
  if (job.kind !== 'batch_scan') return REVIEW_JOB_NAMES[job.kind] || '处理别的任务';
  return job.to_manual ? '转人工' : job.options?.picked ? '自动调整' : job.options?.scan_only ? '只扫描'
    : job.options?.back_to_auto ? '转回自动' : '全自动处理';
}
// 「🤖 自动调整」现在为什么点不了：能点返回空字符串，不能点时按钮提示写明原因
function autoBlockReason(cards) {
  if (!batchState.busy) return '';
  const job = batchState.job || {};
  if (!autoQueueOpen()) {
    return job.pause_requested ? '正在暂停，停下后才能自动调整'
      : job.closing ? `${busyJobName()}正在收尾，结束后才能自动调整`
      : `后台正在${busyJobName()}，结束后才能自动调整`;
  }
  const blocked = cards.find(c => openBlocked(c));
  return blocked ? `${blocked.wb_name}：${openBlocked(blocked).replace('才能打开', '再点自动调整')}` : '';
}
// 后台在跑时点了会怎样：写进可点的自动调整按钮提示
const autoQueueNote = () => !batchState.busy ? ''
  : batchState.job?.options?.scan_only ? '；现在正在扫描：先排队，全部文件扫完后接着做' : '；正在自动调整，会排到后面接着做';
// 复检 / 提交：后台在跑别的批处理任务时排到它后面，结束后接着做；自动处理 / 复检还要处理的文件不能排
// （和后端 enqueue_followup 一致）。能点返回空字符串，不能点时按钮提示写明原因
function queueBlockReason(cards, verb) {
  if (!batchState.busy) return '';
  if (!batchState.job?.kind?.startsWith('batch_')) return `后台正在${busyJobName()}，结束后才能${verb}`;
  const blocked = markWhileBusy() && cards.find(c => openBlocked(c));
  return blocked ? `${blocked.wb_name}：${openBlocked(blocked).replace('才能打开', `再点${verb}`)}` : '';
}
// 复检进行中再点复检：并进正在跑的复检，不用排到后面
const recheckJoins = () => batchState.busy && batchState.job?.kind === 'batch_recheck' && !batchState.job.closing;
const recheckBlockReason = cards => recheckJoins() ? '' : queueBlockReason(cards, '复检');
// 可点的复检 / 提交按钮提示里写明会排队
function queueNote(verb) {
  if (!batchState.busy) return '';
  if (verb === '复检' && recheckJoins()) return '；复检进行中：排进去，当前文件查完接着查';
  return `；后台正在${busyJobName()}：先排队，结束后接着${verb}`;
}
// 「统一复检黄色项」：后台还要处理的文件这次不排，其余照常
const recheckAllCards = (manual = manualCards()) => recheckJoins() ? manual : manual.filter(c => !queueBlockReason([c], '复检'));
const FOLLOWUP_NAMES = { batch_recheck: '复检', batch_commit: '提交' };
// 文件框上：这份文件有排在当前任务后面的复检 / 提交
function followupBadge(path, fileCards = batchState.cards.filter(c => fileKey(c) === path), context) {
  const kinds = context ? context.followups.get(path) || [] : (batchState.followups || []).filter(entry => {
    const ids = new Set(entry.card_ids);
    return fileCards.some(c => ids.has(c.id));
  }).map(entry => FOLLOWUP_NAMES[entry.kind]);
  if (!kinds.length) return '';
  return `<span class="batch-file-queue" data-queue="waiting">⏳ 排队等${[...new Set(kinds)].join('、')} · ${batchState.busy ? '当前任务结束后开始' : '上一个任务暂停了，还没开始'}</span>`;
}
// 「🤖 自动调整」任务里这份文件到哪一步了：文件框上显示，排队的写明前面还有几个文件
function autoQueueBadge(path, fileCards = batchState.cards.filter(c => fileKey(c) === path), context = fileRenderContext()) {
  const job = batchState.job;
  if (batchState.busy && job?.kind === 'batch_scan' && job.options?.scan_only) {
    return fileCards.some(c => context.queued.has(c.id))
      ? '<span class="batch-file-queue" data-queue="waiting">⏳ 排队等自动调整 · 全部文件扫完后开始</span>' : '';
  }
  if (!batchState.busy || job?.kind !== 'batch_scan' || !job.options?.picked || !context.workbooks.has(path)) return '';
  const { done, failed, current, ahead } = context;
  if (done.has(path)) return '<span class="batch-file-queue" data-queue="done">✅ 自动调整完成</span>';
  if (failed[path]) return `<span class="batch-file-queue" data-queue="error" title="${esc(failed[path])}">⚠️ 自动调整出错，已跳过</span>`;
  if (path === current) return '<span class="batch-file-queue" data-queue="running">🤖 正在自动调整</span>';
  return `<span class="batch-file-queue" data-queue="waiting">⏳ 排队等自动调整 · 前面还有 ${ahead.get(path) || 0} 个文件</span>`;
}

function fileRenderContext() {
  const job = batchState.job || {};
  const done = new Set(job.completed || []), failed = job.file_errors || {};
  const waiting = (job.workbooks || []).filter(path => !done.has(path) && !failed[path]);
  const current = waiting.includes(job.repairing) ? job.repairing : waiting[0];
  const ahead = new Map();
  waiting.filter(path => path !== current).forEach((path, i) => { if (!ahead.has(path)) ahead.set(path, i + 1); });
  const byId = new Map(batchState.cards.map(card => [card.id, fileKey(card)])), followups = new Map();
  for (const entry of batchState.followups || []) {
    for (const path of new Set(entry.card_ids.map(id => byId.get(id)).filter(path => path !== undefined))) {
      if (!followups.has(path)) followups.set(path, []);
      followups.get(path).push(FOLLOWUP_NAMES[entry.kind]);
    }
  }
  return { done, failed, current, ahead, followups, workbooks: new Set(job.workbooks || []), queued: new Set(job.auto_queue || []),
    opened: new Set(batchState.opened_files || []) };
}
const canAuto = cards => !autoBlockReason(cards);
function fileGroups(cards) {
  const groups = new Map();
  for (const card of cards) {
    const key = fileKey(card);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(card);
  }
  return [...groups];
}
// 顺序和页面上一致（先按文件、再按处理顺序），大图审核的上一处 / 下一处也按这个走
const shownCards = () => fileGroups(batchState.cards.filter(c => {
  if (selectedFilter === 'all') return true;
  if (c.color === selectedFilter) filterKeep.add(c.id);
  return filterKeep.has(c.id);
})).flatMap(([, cards]) => cards);
const manualCards = () => batchState.cards.filter(c => c.color === 'yellow' && c.finalized);

function snapshotUrl() {
  const paths = new Set(expandedFiles);
  const active = largeMode && batchState.cards.find(c => c.id === activeCardId);
  if (active) paths.add(fileKey(active));
  return '/api/batch?details=lazy' + [...paths].sort().map(path => `&file=${encodeURIComponent(path)}`).join('');
}

function requestFileDetails() {
  lastEtag = '';
  return poll();
}

function toggleFileExpanded(key) {
  const top = fileElements.get(key)?.element.getBoundingClientRect().top;
  if (expandedFiles.has(key)) {
    expandedFiles.delete(key);
    fileMountLimits.delete(key);
    // 收起就释放重内容和卡片节点，下次展开按最新结果重新加载。
    batchState.cards = batchState.cards.map(card => {
      if (fileKey(card) !== key) return card;
      const light = Object.fromEntries(Object.entries(card).filter(([field]) => !CARD_DETAIL_FIELDS.includes(field)));
      return { ...light, waiting: card.waiting ?? card.message === '等待自动处理', details_pending: true };
    });
  } else {
    expandedFiles.add(key);
    fileMountLimits.set(key, FILE_MOUNT_BATCH);
  }
  renderKey = '';
  render();
  if (top !== undefined) window.scrollBy(0, fileElements.get(key).element.getBoundingClientRect().top - top);
  // 先绘制展开动画和加载提示，再开始请求、解析与分批创建卡片。
  requestAnimationFrame(() => requestFileDetails());
}

function toggleFileChecked(key) {
  selectedFiles.has(key) ? selectedFiles.delete(key) : selectedFiles.add(key);
  applySelection();
  renderControls();
}

// 每个选中文件取一张卡作打开后的跳转目标：优先第一张人工待处理的黄卡。
function selectedFileCards() {
  const picked = new Map();
  for (const card of batchState.cards) {
    const key = fileKey(card);
    if (!selectedFiles.has(key)) continue;
    const manual = card.color === 'yellow' && card.finalized;
    if (!picked.has(key) || (manual && !picked.get(key).manual)) picked.set(key, { card, manual });
  }
  return [...picked.values()].map(entry => entry.card);
}

function mountedFileCards(key, cards) {
  if (!expandedFiles.has(key)) return [];
  return cards.filter(c => !c.details_pending).slice(0, fileMountLimits.get(key) ?? FILE_MOUNT_BATCH);
}

async function post(url, payload = {}) {
  lastEtag = '';
  ++pollVersion;
  const response = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
  const result = await response.json();
  if (!response.ok || !result.ok) throw new Error(result.error || `HTTP ${response.status}`);
  return result;
}

async function poll(background = false) {
  if (background && pollInFlight) return;
  ++pollInFlight;
  const version = ++pollVersion;
  try {
    const url = snapshotUrl();
    if (lastSnapshotUrl !== url) { lastEtag = ''; lastSnapshotUrl = url; }
    const response = await fetch(url, { cache: 'no-store', headers: lastEtag ? { 'If-None-Match': lastEtag } : {} });
    if (response.status === 304) {
      if (version === pollVersion && !requestBusy && !colorSaving) renderControls();
      return;
    }
    if (!response.ok) throw new Error(`无法读取进度：HTTP ${response.status}`);
    const value = await response.json();
    if (version !== pollVersion || requestBusy || colorSaving) return;
    celebrateFinished(batchState, value);
    if (value.busy && value.job?.kind === 'batch_unhighlight') clearPending = true;
    const cleared = clearPending && !value.busy && value.job?.kind === 'batch_unhighlight';
    if (cleared) clearPending = false;
    if (fileFolder !== null && fileFolder !== value.folder) {
      expandedFiles.clear(); selectedFiles.clear(); fileMountLimits.clear();
      // 切换文件夹的这一帧也不保留意外随旧展开请求返回的重内容。
      value.cards = value.cards.map(card => ({ ...Object.fromEntries(Object.entries(card)
        .filter(([field]) => !CARD_DETAIL_FIELDS.includes(field))), details_pending: true }));
    }
    fileFolder = value.folder;
    batchState = value;
    lastEtag = response.headers.get('ETag') || '';
    if (Array.isArray(value.fix_error_types) && !hasSavedFixTypes) {
      fixErrorTypes = new Set(FIXABLE_CODES.filter(code => value.fix_error_types.includes(code)));
      renderFixErrorTypes();
    }
    if (!document.activeElement || document.activeElement.id !== 'folderInput') {
      if (value.folder && !$('folderInput').value) $('folderInput').value = value.folder;
      if (value.busy && value.folder) $('folderInput').value = value.folder;
    }
    render();
    if (cleared && unhighlightResult()) showUnhighlightResult();
  } catch (error) {
    if (version === pollVersion && !requestBusy && !colorSaving) $('errorText').textContent = error.message;
  } finally {
    --pollInFlight;
  }
}

// 清除标记的结果：任务结束后弹窗，之后可点「上次清除结果」再看。
// 清除任务很快时轮询可能一次都没看到「进行中」，所以点了清除（或继续清除任务）就先记下来
let clearPending = false;
function unhighlightResult() {
  const job = batchState.job;
  return job?.kind === 'batch_unhighlight' && job.result && Object.keys(job.result).length ? job.result : null;
}

const CLEAR_STATUS = { ok: '✅', check: '⚠️', failed: '❌' };
function showUnhighlightResult() {
  const files = Object.values(unhighlightResult() || {});
  const count = { ok: 0, check: 0, failed: 0 };
  for (const file of files) for (const cell of file.cells) count[cell.status] += 1;
  const errors = files.filter(f => f.error).length;
  const leftovers = files.reduce((sum, f) => sum + (f.leftovers || 0), 0);
  const tabs = f => f.tabs || [];
  const trouble = f => !!f.error || (f.leftover_failed || []).length > 0 || f.cells.some(c => c.status !== 'ok')
    || tabs(f).some(t => t.status !== 'ok');
  files.sort((a, b) => trouble(b) - trouble(a) || a.name.localeCompare(b.name));
  const chips = [`<span class="unhighlight-chip is-ok">✅ 已清除 ${count.ok} 格</span>`];
  if (count.check) chips.push(`<span class="unhighlight-chip is-check">⚠️ 需要看一下 ${count.check} 格</span>`);
  if (count.failed) chips.push(`<span class="unhighlight-chip is-failed">❌ 没清掉 ${count.failed} 格</span>`);
  if (errors) chips.push(`<span class="unhighlight-chip is-failed">❌ ${errors} 份文件没能处理</span>`);
  if (leftovers) chips.push(`<span class="unhighlight-chip">🧹 另清掉残留标记 ${leftovers} 条</span>`);
  const tabCount = files.reduce((sum, f) => sum + tabs(f).length, 0);
  if (tabCount) chips.push(`<span class="unhighlight-chip">🏷️ 工作表标签 ${tabCount} 张</span>`);
  const fileHtml = file => {
    const own = { ok: 0, check: 0, failed: 0 };
    for (const cell of file.cells) own[cell.status] += 1;
    const tally = file.error ? `<span class="unhighlight-file-error">${esc(file.error)}</span>`
      : file.cells.length || tabs(file).length ? `✅ ${own.ok}${own.check ? ` · ⚠️ ${own.check}` : ''}${own.failed ? ` · ❌ ${own.failed}` : ''}` : '没有要清的标记';
    const rows = file.cells.map(cell => {
      const where = `${cell.sheet}!${cell.addr}`;
      const link = excelLinkHtml(esc(where), { file: file.name, folder: batchState.folder, sheet: cell.sheet, cell: cell.addr.split(':')[0] });
      return `<li class="is-${cell.status}"><span aria-hidden="true">${CLEAR_STATUS[cell.status] || ''}</span>${link}<span class="unhighlight-reason">${esc(cell.reason)}</span></li>`;
    });
    for (const tab of tabs(file)) {
      const link = excelLinkHtml(esc(`🏷️ ${tab.sheet}`), { file: file.name, folder: batchState.folder, sheet: tab.sheet });
      rows.push(`<li class="is-${tab.status}"><span aria-hidden="true">${CLEAR_STATUS[tab.status] || ''}</span>${link}<span class="unhighlight-reason">工作表标签：${esc(tab.reason)}</span></li>`);
    }
    for (const reason of file.leftover_failed || []) {
      rows.push(`<li class="is-failed"><span aria-hidden="true">❌</span><span>残留标记</span><span class="unhighlight-reason">删不掉：${esc(reason)}</span></li>`);
    }
    return `<details class="unhighlight-file"${trouble(file) ? ' open' : ''}><summary><b>${fileNameLink(file, batchState.folder)}</b><span class="unhighlight-tally">${tally}</span></summary>`
      + (rows.length ? `<ul>${rows.join('')}</ul>` : '') + '</details>';
  };
  const tone = count.failed || errors ? 'error' : count.check ? 'warn' : 'info';
  $('unhighlightDialog').dataset.tone = tone;
  $('unhighlightTitle').textContent = tone === 'info' ? '标记已全部清除' : '清除结果：有格子要看一下';
  $('unhighlightBody').innerHTML = `<p class="unhighlight-summary">${files.length} 份文件　${chips.join('')}</p>`
    + `<p class="unhighlight-hint">点格子位置可以用 Excel 打开并跳到那一格。⚠️ 和 ❌ 的请人工看一下。</p>`
    + `<div class="unhighlight-files">${files.map(fileHtml).join('')}</div>`;
  if (!$('unhighlightDialog').open) $('unhighlightDialog').showModal();
  $('unhighlightBody').scrollTop = 0;
  $('unhighlightCloseBtn').focus(); // 不然浏览器把焦点放在第一个文件上，显示一圈焦点框
}

// 后台任务刚结束且成功：撒花。提交 / 复检时若有暂停的全自动处理会被放回来，任务种类变了也算提交完成
const FINISH_TITLES = { batch_scan: '全自动处理完成！', batch_commit: '审核已提交！', batch_recheck: '复检完成！' };
function celebrateFinished(before, after) {
  const kind = before.job?.kind;
  if (!before.busy || after.busy || !FINISH_TITLES[kind]) return;
  if (after.job?.kind === kind && after.job?.status !== 'completed') return;
  const left = after.cards.length;
  const title = kind === 'batch_scan' && before.job.options?.scan_only ? '扫描完成！'
    : kind === 'batch_scan' && before.job.to_manual ? '已转人工！'
    : kind === 'batch_scan' && before.job.options?.picked ? '自动调整完成！' : FINISH_TITLES[kind];
  window.gameFx?.celebrate(title, left ? `还有 ${left} 张卡片待审核` : '全部卡片都处理完了');
}

async function action(work) {
  if (requestBusy) return;
  requestBusy = true;
  // 刚点的颜色还在保存：等保存完再提交、复检，按你最后选的颜色来
  while (colorSaving) await new Promise(resolve => setTimeout(resolve, 50));
  ++pollVersion;
  $('errorText').textContent = '';
  renderControls();
  try { await work(); } catch (error) { $('errorText').textContent = error.message; }
  finally { requestBusy = false; lastEtag = ''; await poll(); }
}

function renderControls() {
  const busy = batchState.busy || requestBusy;
  const analysis = batchState.scan_analysis;
  $('scanReportBtn').classList.toggle('hidden', !analysis || analysis.folder !== batchState.folder);
  if (analysis) $('scanReportBtn').href = `/analysis?id=${encodeURIComponent(analysis.id)}&folder=${encodeURIComponent(analysis.folder)}`;
  const manual = manualCards(), recheckAll = recheckAllCards(manual), recheckSet = new Set(recheckAll);
  for (const id of ['startBtn', 'scanOnlyBtn', 'commitBtn', 'folderInput', 'scanLimit']) $(id).disabled = busy;
  $('unhighlightBtn').disabled = busy || !(batchState.highlighted || batchState.tab_marks);
  $('unhighlightBtn').textContent = `🧽 清除 Excel 标记颜色${batchState.highlighted ? `（${batchState.highlighted}）` : ''}`;
  $('unhighlightResultBtn').classList.toggle('hidden', batchState.busy || !unhighlightResult());
  const plan = commitPlan();
  const offered = plan.cards.length + plan.skipped.length;
  // 后台在跑别的批处理任务时提交先排队；后台还要处理的文件里的卡这次不交
  const commitBlock = queueBlockReason([], '提交');
  $('commitBtn').disabled = requestBusy || !offered || !!commitBlock;
  $('commitBtn').title = commitBlock ? `现在不能点：${commitBlock}`
    : batchState.busy ? `后台正在${busyJobName()}：先排队，结束后接着提交，卡片按那时的颜色处理`
      + (plan.held.length ? `；后台还要处理的文件里的 ${plan.held.length} 张这次不交` : '') : '';
  $('commitBtn').textContent = selectedFilter === 'all' ? `✓ 确认并处理全部标记${offered ? `（${offered}）` : ''}`
    : `✓ 只处理${COLOR_NAMES[selectedFilter]}色标记（${offered}）`;
  const openFiles = new Map(manual.map(c => [fileKey(c), c]));
  const opened = new Set(batchState.opened_files || []);
  const unopened = [...openFiles.values()].filter(c => !opened.has(fileKey(c)));
  const openAvailable = unopened.filter(c => !batchState.busy || !openBlocked(c)).length;
  const opening = batchState.busy && batchState.job?.kind === 'batch_open';
  const pickedFiles = selectedFileCards();
  $('openAllBtn').textContent = openFiles.size && !unopened.length
    ? `📂 本轮已全部打开（共 ${openFiles.size} 个）`
    : `📂 打开下一批（未开 ${unopened.length} / 共 ${openFiles.size} 个）`;
  $('openAllBtn').disabled = requestBusy || openChoosing || opening || !unopened.length;
  $('openAllBtn').title = '每次最多打开 5 个本轮尚未成功打开的文件；关闭 Excel 后仍会记住进度';
  $('resetOpenHistoryBtn').disabled = requestBusy || openChoosing || opening || !opened.size;
  $('openSelectedFilesBtn').textContent = `📂 打开选中文件（${pickedFiles.length} 个）`;
  $('openSelectedFilesBtn').disabled = requestBusy || openChoosing || opening || !pickedFiles.length;
  $('clearSelectedFilesBtn').disabled = openChoosing || !selectedFiles.size;
  $('openFilesHint').textContent = `本轮累计已打开 ${opened.size} 个文件；当前有 ${openFiles.size} 个人工待处理文件，其中 ${unopened.length} 个本轮未打开（当前可打开 ${openAvailable} 个）。每次最多打开 5 个新文件，全部打开后不自动循环；重置打开记录可从头开始，勾选文件也可明确重新打开。`;
  // 复检进行中还能接着点复检：排进同一个任务；后台在跑别的批处理任务时排到它后面
  const held = new Set(manual.filter(c => !recheckSet.has(c)).map(fileKey)).size;
  $('recheckAllBtn').disabled = requestBusy || !recheckAll.length;
  $('recheckAllBtn').title = !manual.length ? '还没有已提交的黄色人工项'
    : !recheckAll.length ? `现在不能点：${recheckBlockReason(manual)}`
    : `整份重新检测有黄卡的 ${new Set(recheckAll.map(fileKey)).size} 个文件`
      + (held ? `（后台还要处理的 ${held} 个文件这次不排）` : '') + queueNote('复检');
  $('pauseBtn').classList.toggle('hidden', !batchState.busy || batchState.job?.kind !== 'batch_scan');
  $('pauseBtn').disabled = requestBusy || !!batchState.job?.pause_requested;
  $('toManualBtn').classList.toggle('hidden', !batchState.mode?.to_manual);
  $('toManualBtn').disabled = requestBusy;
  $('toAutoBtn').classList.toggle('hidden', !batchState.mode?.to_auto);
  $('toAutoBtn').disabled = busy;
  $('resumeBtn').classList.toggle('hidden', batchState.busy || batchState.job?.status !== 'interrupted');
  $('resumeBtn').disabled = requestBusy;
  $('abandonBtn').classList.toggle('hidden', batchState.busy || !batchState.job?.pending_adjustment);
  $('abandonBtn').disabled = requestBusy;
  renderPendingNotice();
  renderFollowups();
  renderSelectBar();
  $('statusText').textContent = batchState.busy ? '⏳ 后台处理中' : batchState.job?.status === 'interrupted' ? '⏸ 可继续任务' : '就绪';
}

// 复制报错 / 日志：页面跑在 localhost，一般能用 clipboard；不行就退回 execCommand
const errorLogLines = () => batchState.logs.filter(line => /❌|⚠️/.test(line));

async function copyText(text, button) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const area = Object.assign(document.createElement('textarea'), { value: text });
    area.style.cssText = 'position:fixed;left:-9999px;top:0';
    document.body.append(area);
    area.select();
    const ok = document.execCommand('copy');
    area.remove();
    if (!ok) { $('errorText').textContent = '复制失败，请手动选中文字复制'; return; }
  }
  const label = button.dataset.label || (button.dataset.label = button.textContent);
  button.textContent = '已复制 ✓';
  clearTimeout(button.copyTimer);
  button.copyTimer = setTimeout(() => { button.textContent = label; }, 1500);
}

function renderCopyErrorBtn() {
  const hasError = $('errorText').textContent.trim() || !$('pendingNotice').classList.contains('hidden');
  $('copyErrorBtn').classList.toggle('hidden', !hasError);
}

// 写入核对没通过时任务停住、什么都不能开始：写清楚卡在哪、两个按钮分别做什么
function renderPendingNotice() {
  const pending = !batchState.busy && batchState.job?.pending_adjustment;
  $('pendingNotice').classList.toggle('hidden', !pending);
  if (!pending) return;
  const name = String(pending.path || '').split(/[\\/]/).pop();
  const rows = (pending.rows || []).map(r => `${r.sheet} 第 ${r.row} 行`).join('、');
  const resume = batchState.job.status === 'interrupted' ? '点「继续上次任务」重新核对并写入；' : '';
  $('pendingNotice').textContent = `上次写入没有完成：${name}（${rows}）。${resume}如果你在 Excel 里改过这一行，或不想再自动改它，点「放弃这次写入」：文件不再改动，这几格转为黄色人工处理。其他文件不受影响，可以照常提交、复检。`;
}

// 秒数写成「3 分 20 秒」「1 小时 5 分」
function duration(seconds) {
  const s = Math.max(0, Math.round(seconds));
  if (s < 60) return `${s} 秒`;
  if (s < 3600) return `${Math.floor(s / 60)} 分${s % 60 ? ` ${s % 60} 秒` : ''}`;
  return `${Math.floor(s / 3600)} 小时${Math.floor(s % 3600 / 60) ? ` ${Math.floor(s % 3600 / 60)} 分` : ''}`;
}
// 扫描时进度后面接上已用和预计剩余时间（按这次已扫完文件的耗时和文件大小估算）
function etaText(eta) {
  if (!eta) return '';
  if (eta.remaining == null) return ` · 已用 ${duration(eta.elapsed)} · 扫完第一个文件后估算剩余时间`;
  const done = new Date(Date.now() + eta.remaining * 1000);
  const clock = `${String(done.getHours()).padStart(2, '0')}:${String(done.getMinutes()).padStart(2, '0')}`;
  return ` · 已用 ${duration(eta.elapsed)} · 预计还要约 ${duration(eta.remaining)}（${clock} 左右扫完）`;
}
// 正在检查的文件：第几张表、这份用了多久、超过多久会跳过
function renderScanNow(current) {
  $('scanNow').classList.toggle('hidden', !current);
  if (!current) return;
  const where = current.total ? `第 ${current.index} / ${current.total} 张表「${current.sheet}」` : '正在打开';
  $('scanNow').textContent = `正在检查 ${current.name}：${where} · 这份已用 ${duration(current.seconds)}`
    + (current.limit ? `（上限 ${duration(current.limit)}，超过就跳过、留给人工）` : '');
}
// 扫描超时跳过的文件：写明卡在哪张表，点文件名用 Excel 打开人工检查
function renderSlowFiles() {
  const slow = batchState.slow_files || [];
  $('slowNotice').classList.toggle('hidden', !slow.length);
  if (!slow.length) return;
  const files = slow.map(f => `${excelLinkHtml(esc(f.name), { file: f.name, folder: batchState.folder, blocked: openBlocked({ wb_path: f.path }) })}`
    + `（超过 ${duration(f.limit)}${f.sheet ? `，卡在表「${esc(f.sheet)}」` : ''}）`).join('、');
  const html = `⏱ ${slow.length} 个文件扫描超时，已跳过，请人工检查：${files}`;
  if ($('slowText').innerHTML !== html) $('slowText').innerHTML = html;
  $('slowClearBtn').disabled = requestBusy;
}

// 排在当前任务后面的复检 / 提交：写明排了什么、什么时候做；暂停时留着的可以现在开始，都可以取消
function renderFollowups() {
  const queue = batchState.followups || [];
  $('followupNotice').classList.toggle('hidden', !queue.length);
  if (!queue.length) return;
  const byId = new Map(batchState.cards.map(c => [c.id, fileKey(c)]));
  const files = ids => new Set(ids.filter(id => byId.has(id)).map(id => byId.get(id))).size;
  const parts = queue.map(entry => entry.kind === 'batch_recheck'
    ? `复检 ${files(entry.card_ids)} 个文件（${entry.card_ids.length} 格）`
    : `提交${entry.color ? `${COLOR_NAMES[entry.color]}色` : ''} ${entry.card_ids.length} 张卡`);
  $('followupText').textContent = `⏳ 排队中：${parts.join(' → ')}。`
    + (batchState.busy ? `后台正在${busyJobName()}，结束后按顺序接着做。` : '上一个任务暂停了，这些还没做：可以现在开始，或取消排队。');
  $('followupStartBtn').classList.toggle('hidden', batchState.busy);
  $('followupStartBtn').disabled = requestBusy;
  $('followupCancelBtn').disabled = requestBusy;
}

// 图片取不到时换成文字占位（整页图连同红框一起换掉）
const missingHandler = (text, wholeParent = false) => `${wholeParent ? 'this.parentElement.replaceChildren' : 'this.replaceWith'}(Object.assign(document.createElement('span'),{className:'batch-missing',textContent:'${text}'}))`;

function contextHtml(card) {
  const context = card.context;
  if (!context?.image) return '<aside class="batch-context"><h3>问题所在页面</h3><div class="batch-context-missing">暂无整页定位图<br>未保存整页截图或未匹配到文字位置，可点击文件名在 Excel 中核对。</div></aside>';
  const width = Number(context.page_width), height = Number(context.page_height);
  const box = context.bbox;
  if (!(Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0)
      || !Array.isArray(box) || box.length !== 4 || !box.every(Number.isFinite)) {
    return '<aside class="batch-context"><h3>问题所在页面</h3><div class="batch-context-missing">暂无可靠的页面定位</div></aside>';
  }
  const clamp = value => Math.max(0, Math.min(100, value));
  const left = clamp(box[0] / width * 100), top = clamp(box[1] / height * 100);
  const right = clamp(box[2] / width * 100), bottom = clamp(box[3] / height * 100);
  const page = Number.isInteger(context.page_index) ? ` · 第 ${context.page_index + 1} 页` : '';
  return `<aside class="batch-context"><h3>问题所在页面${page}</h3><p class="batch-context-note">检测时的整页截图 · 红框为问题位置</p>
    <div class="batch-context-page" style="--page-ratio:${width / height}">
      <img src="${esc(context.image)}" alt="问题所在整页，红框标出原检测位置" onerror="${missingHandler('整页截图暂不可用', true)}">
      <span class="batch-context-marker" aria-label="问题位置" style="left:${left}%;top:${top}%;width:${Math.max(0, right - left)}%;height:${Math.max(0, bottom - top)}%"></span>
    </div></aside>`;
}

const pt = value => `${Number(value).toFixed(1).replace(/\.0$/, '')}pt`;

// 检测到什么、怎么处理的：让人不用对比两张图也知道为什么要修、改了多少
function problemHtml(card) {
  const problems = (card.problems || []).map(p => `<b>${esc(p.label)}</b>：${esc(p.reason)}${p.missing ? `（没显示出来的：「${esc(p.missing)}」）` : ''}`);
  return problems.length ? `<p class="batch-problem">检测到 ${problems.join('；')}</p>` : '';
}

function handlingText(card) {
  const row = card.row_height;
  if (card.after_kind === 'attempt') {
    const base = Number.isFinite(card.base_height) ? `${pt(card.base_height)} → ` : '';
    const added = Number.isFinite(card.base_height) && Number.isFinite(card.attempt_height) ? `（+${pt(card.attempt_height - card.base_height)}）` : '';
    const wrap = card.attempt_wrap?.length ? `${card.cell_addr} 开启自动换行，` : '';
    return Number.isFinite(card.attempt_height) ? `处理：${wrap}第 ${card.excel_row} 行行高试到 ${base}${pt(card.attempt_height)}${added}仍未通过，原文件未修改` : '处理：原文件未修改';
  }
  const brk = card.page_break;
  if (brk?.saved && brk.applied) return `处理：在第 ${brk.row} 行上方插入分页，整块移到下一页，已保存`;
  if (brk && !brk.applied && card.finalized) return '处理：已去掉插入的分页';
  const col = card.col_width;
  if (col?.saved && Math.abs(col.applied - col.base) >= 0.05) return `处理：${card.cell_addr} ${col.cols?.length > 1 ? '涉及列宽合计' : '所在列宽'} ${pt(col.base).replace('pt', '')} → ${pt(col.applied).replace('pt', '')}（+${pt(col.applied - col.base).replace('pt', '')}），已保存`;
  if (col && !col.saved && Math.abs(col.applied - col.base) < 0.05 && card.finalized) return `处理：已还原为原列宽 ${pt(col.base).replace('pt', '')}`;
  if (card.kind === 'pagebreak' && card.after_kind === 'rechecked') return '处理：未插入分页（其他行调整后重新导出已不再被切开）';
  if (card.after_kind === 'handover') return '处理：转人工时自动还没处理到这一格，原文件未修改';
  if (card.after_kind === 'scan_only') return `处理：只扫描，未自动修改${card.highlighted ? '；已在 Excel 中给这一格标色' : ''}`;
  if (card.after_kind === 'unchanged') return '处理：原样导出已完整显示，未修改行高';
  if (row?.saved && Math.abs(row.applied - row.base) >= 0.05) return `处理：${row.wrap ? `${card.cell_addr} 开启自动换行，` : ''}第 ${card.excel_row} 行行高 ${pt(row.base)} → ${pt(row.applied)}（+${pt(row.applied - row.base)}），已保存`;
  if (row && !row.saved && Math.abs(row.applied - row.base) < 0.05 && card.finalized) return `处理：已还原为原行高 ${pt(row.base)}${row.wrap ? '，并取消自动换行' : ''}`;
  if (card.after_kind === 'rechecked') return card.error_types?.includes('hash_overflow')
    ? '处理：本格未单独调整列宽（其他单元格调整后重新导出已完整显示）'
    : '处理：本行行高未修改（其他行调整后重新导出已完整显示）';
  return '';
}

function blueNote(card) {
  if (card.adoptable) {
    const height = Number.isFinite(card.attempt_height) ? ` ${pt(card.attempt_height)}` : '';
    const wrap = card.attempt_wrap?.length ? `给 ${card.cell_addr} 开启自动换行，` : '';
    return card.color === 'blue'
      ? `已按你的判断标蓝：提交时${wrap}把第 ${card.excel_row} 行行高改为${height}（最后一次尝试）并保存，记为已修复。`
      : `自动检查没通过，但如果「最后一次尝试」的图已经显示完整，可以标蓝：提交时${wrap}按${height}写入原文件。`;
  }
  return '三种颜色均可自由选择。标蓝表示你认可当前结果；没有可采用的试修方案时，提交只按人工确认放行，不额外修改 Excel。';
}

// 没有修改后图时，在单图下面说明原因
function afterNote(card) {
  if (['scan_only', 'handover'].includes(card.after_kind)) return '还没有修改后的图：人工修改后复检通过再显示';
  if (waitingCard(card)) return '还没有修改后的图：自动处理到这一格后显示';
  if (card.after_kind === 'attempt') return '最后一次尝试没有截到图，原文件未修改';
  return card.after_kind ? '没有截到修改后的图，可点击表名在 Excel 中核对' : '还没有修改后的图';
}

// 卡片上的「🎯 单项复检」：后台在跑时排队，不可点时提示写明原因
function recheckButtonHtml(card) {
  const blocked = recheckBlockReason([card]);
  const title = blocked ? `现在不能点：${blocked}`
    : `只重新检测这一格（只导出它所在的那一页），同文件其他黄卡不动；整份重查请用文件框上的「复检这个文件」${queueNote('复检')}`;
  return `<button class="btn btn-outline" data-action="recheck"${blocked ? ' disabled' : ''} title="${esc(title)}">🎯 单项复检</button>`;
}

// 卡片上的「🤖 自动调整」：不可点时提示写明原因
function autoButtonHtml(card) {
  const blocked = autoBlockReason([card]);
  const title = blocked ? `现在不能点：${blocked}`
    : `让程序自动加高这一格所在的行（同一行的黄卡一起），同文件其他黄卡不动${autoQueueNote()}`;
  return `<button class="btn btn-outline" data-action="auto"${blocked ? ' disabled' : ''} title="${esc(title)}">🤖 自动调整</button>`;
}

function cardHtml(card, large = false) {
  const where = `${card.sheet_name}!${card.cell_addr || '定位待确认'}`;
  const target = { file: card.wb_name, folder: batchState.folder, sheet: card.sheet_name, cell: card.cell_addr, blocked: openBlocked(card) };
  // 文件名已经写在文件框上：小卡片只写「表!格」（点了照样打开 Excel 跳到这一格）；大图审核没有文件框，照旧写文件名
  const title = large ? `${excelLinkHtml(esc(card.wb_name), target)}<p class="batch-cell">${esc(where)}</p>`
    : excelLinkHtml(esc(where), target, 'batch-cell-link');
  const labels = { before: '修改前', after: card.after_kind === 'attempt' ? '最后一次尝试（未采用）' : card.after_kind === 'unchanged' ? '原样导出（未修改）' : card.after_kind === 'rechecked' ? '最新导出（本行未修改）' : ['scan_only', 'handover'].includes(card.after_kind) ? '人工修改后（复检通过后显示）' : '修改后' };
  const handling = handlingText(card);
  const shot = which => `<figure><figcaption>${labels[which]}</figcaption>${card.images?.[which]
    ? `<img ${large ? '' : 'loading="lazy"'} src="${esc(card.images[which])}" alt="${labels[which]}截图" data-preview="1" onerror="${missingHandler('截图暂不可用')}">`
    : '<span class="batch-missing">暂无截图</span>'}</figure>`;
  // 还没有修改后的图（只扫描、转人工、等自动处理、截图失败）：不放空的 VS 右半边，只放一张大一点的修改前图，并写明修改后图什么时候会有
  const comparison = card.images?.after
    // 左右两块斜切外框中间放 VS 徽章；截图本身保持方正，不倾斜不变形
    ? `<div class="batch-shots">${shot('before')}<span class="batch-vs" aria-hidden="true">VS</span>${shot('after')}</div>`
    : `<div class="batch-shots is-single">${shot('before')}<p class="batch-after-note">${esc(afterNote(card))}</p></div>`;
  const visuals = large ? `<div class="batch-large-layout">${comparison}${contextHtml(card)}</div>` : comparison;
  return `<article class="batch-card" tabindex="0" role="button" aria-label="${esc(`${card.wb_name} ${card.sheet_name}!${card.cell_addr} ${REVIEW_LABELS[card.color]}，单击切换状态`)}" data-id="${esc(card.id)}" data-color="${card.color}">
    <div class="batch-card-head"><button type="button" class="batch-select" data-action="select" aria-pressed="false" title="选中 / 取消选中">✓</button><div class="batch-card-title">${title}</div><span class="batch-color-label">${REVIEW_LABELS[card.color]}</span></div>
    ${visuals}${problemHtml(card)}${handling ? `<p class="batch-handling">${esc(handling)}</p>` : ''}${card.text?.trim() ? `<p class="batch-text">${esc(card.text)}</p>` : ''}${card.message ? `<p class="batch-message">${esc(card.message)}</p>` : ''}
    <p class="batch-state">${card.finalized ? '已提交 · 可人工处理' : '颜色标记待提交'}${card.highlighted && card.after_kind !== 'scan_only' ? ' · 已在 Excel 中标色' : ''}${card.handed_over ? ' · 自动还没试过，可转回自动' : ''}${canMark(card) ? '' : ' · 后台处理完这个文件后可标记'}</p>
    ${card.verified ? '' : `<p class="batch-no-blue">${blueNote(card)}</p>`}
    ${card.auto_skip ? `<p class="batch-auto-off">🙅 不能自动调整：${esc(card.auto_skip)}</p>` : ''}
    ${card.color === 'yellow' && card.finalized ? `<div class="batch-card-footer">${recheckButtonHtml(card)}${card.auto_skip ? '' : autoButtonHtml(card)}</div>` : ''}</article>`;
}

// 选中状态不进卡片 HTML：勾选只改类名，不重画卡片、不重新加载图片
let domRevision = 0;
let lastSelectionKey = '';
const selectionOnElement = new WeakMap();
function applySelection() {
  renderSelectBar();
  const key = JSON.stringify([[...selectedIds], [...selectedFiles]]) + '|' + domRevision;
  if (key === lastSelectionKey) return;
  lastSelectionKey = key;
  const elements = [...cardElements.values()];
  if (largeMode) {
    const large = $('largeCard').querySelector('.batch-card[data-id]');
    if (large) elements.push(large);
  }
  for (const element of elements) {
    const selected = selectedIds.has(element.dataset.id);
    if (selectionOnElement.get(element) === selected) continue;
    selectionOnElement.set(element, selected);
    element.classList.toggle('is-selected', selected);
    element.querySelector('.batch-select')?.setAttribute('aria-pressed', String(selected));
  }
  const byFile = new Map(fileGroups(shownCards()));
  for (const [path, file] of fileElements) {
    const checkbox = file.head.querySelector('[data-file-check]');
    if (checkbox) checkbox.checked = selectedFiles.has(path);
    file.element.classList.toggle('is-file-selected', selectedFiles.has(path));
    const button = file.head.querySelector('[data-file-action="select"]');
    if (!button) continue;
    const ids = (byFile.get(path) || []).map(c => c.id);
    const all = ids.length > 0 && ids.every(id => selectedIds.has(id));
    button.setAttribute('aria-pressed', String(all));
    button.textContent = all ? '取消选中本文件' : '选中本文件的卡';
  }
}

function renderSelectBar() {
  const busy = batchState.busy || requestBusy;
  const files = selectedFileCards().length;
  const cardTools = cardElements.size > 0 || selectedIds.size > 0;
  $('selectBar').classList.toggle('hidden', largeMode || (!files && !cardTools));
  $('fileSelectActions').classList.toggle('hidden', !files);
  $('cardSelectActions').classList.toggle('hidden', !cardTools);
  $('pickedFileCount').textContent = `已选 ${files} 个文件`;
  $('openPickedFilesBtn').disabled = requestBusy || openChoosing || (batchState.busy && batchState.job?.kind === 'batch_open') || !files;
  $('clearPickedFilesBtn').disabled = openChoosing || !files;
  $('selectCount').textContent = (selectedIds.size ? `已选 ${selectedIds.size} 张` : '未选择卡片') + (selectNotice ? ` · ${selectNotice}` : '');
  const chosen = batchState.cards.filter(c => selectedIds.has(c.id));
  for (const button of $('selectBar').querySelectorAll('[data-set-color]')) {
    button.disabled = requestBusy || !selectedIds.size || !chosen.every(canMark);
  }
  const pickedCards = manualCards().filter(c => selectedIds.has(c.id));
  const picked = pickedCards.length;
  const autoBlock = picked ? autoBlockReason(pickedCards) : '';
  $('autoSelectedBtn').disabled = requestBusy || !picked || !!autoBlock;
  $('autoSelectedBtn').textContent = `🤖 自动调整选中${picked ? `（${picked}）` : ''}`;
  $('autoSelectedBtn').title = !picked ? '先选中已提交的黄卡（同一行的黄卡一起调整），确认框里会列出跳过的和原因'
    : autoBlock ? `现在不能点：${autoBlock}`
    : `让程序自动加高选中的 ${picked} 张黄卡（同一行的黄卡一起），确认框里会列出跳过的和原因${autoQueueNote()}`;
}

function toggleSelected(id) {
  if (selectedIds.has(id)) selectedIds.delete(id); else selectedIds.add(id);
  selectNotice = '';
  applySelection();
}

// 文件框上的「选中本文件的卡」：已经全选中就全部取消，否则把这个文件当前看得到的卡全部选上
function toggleFileSelected(key) {
  const ids = shownCards().filter(c => fileKey(c) === key).map(c => c.id);
  const all = ids.every(id => selectedIds.has(id));
  ids.forEach(id => all ? selectedIds.delete(id) : selectedIds.add(id));
  selectNotice = '';
  applySelection();
}

function selectShown(mode) {
  const ids = shownCards().map(c => c.id);
  if (mode === 'all') ids.forEach(id => selectedIds.add(id));
  else if (mode === 'invert') ids.forEach(id => selectedIds.has(id) ? selectedIds.delete(id) : selectedIds.add(id));
  else selectedIds.clear();
  selectNotice = '';
  applySelection();
}

async function changeSelectedColor(color) {
  const chosen = batchState.cards.filter(c => selectedIds.has(c.id));
  if (requestBusy || !chosen.length || !chosen.every(canMark)) return;
  const ids = chosen.map(c => c.id);
  await action(async () => {
    await post('/api/batch/mark_many', { card_ids: ids, color });
    selectNotice = '';
  });
}

// 更新前记下视口里一张颜色没变的卡片的位置，更新后滚回同一位置。
function scrollAnchor() {
  const colors = new Map(batchState.cards.map(c => [c.id, c.color]));
  let fallback = null;
  for (const file of fileElements.values()) {
    if (!file.element.isConnected) continue;
    const fileRect = file.element.getBoundingClientRect();
    if (fileRect.bottom < 90 || fileRect.top > innerHeight) continue;
    for (const element of file.element.querySelectorAll('.batch-card[data-id]')) {
      const rect = element.getBoundingClientRect();
      if (rect.bottom < 90 || rect.top > innerHeight) continue;
      const anchor = { id: element.dataset.id, top: rect.top };
      if (colors.get(element.dataset.id) === element.dataset.color) return anchor;
      fallback = fallback || anchor;
    }
  }
  return fallback;
}

function restoreAnchor(anchor) {
  const element = anchor && cardElements.get(anchor.id);
  if (element?.isConnected) window.scrollBy(0, element.getBoundingClientRect().top - anchor.top);
}

// 每个文件框里是一个瀑布流：卡片放进这个框里当时最矮的一列（按实际高度算），放好后就记住，不再挪动。
// 以前按序号轮流分到各列，截图和文字高度不一，卡片一多各列高度越差越大，后面只剩一列
const CARD_GAP = 18;
let columnOf = new Map();
let columnMeta = '';
const cardHeights = new Map();
let cardHeightSum = 0;
const dirtyHeights = new Set();
let measuredWidth = 0;
let heightFrame = null;

function rememberCardHeight(id, height) {
  cardHeightSum += height - (cardHeights.get(id) || 0);
  cardHeights.set(id, height);
}

function forgetCardHeight(id) {
  cardHeightSum -= cardHeights.get(id) || 0;
  cardHeights.delete(id);
  dirtyHeights.delete(id);
}

function measureDirtyCards() {
  if (largeMode) return;
  // 先批量读取高度，之后再移动节点。
  for (const id of dirtyHeights) {
    const element = cardElements.get(id);
    if (!element?.isConnected) continue;
    const height = element.getBoundingClientRect().height;
    if (height <= 0) continue;
    rememberCardHeight(id, height + CARD_GAP);
    dirtyHeights.delete(id);
  }
}

function scheduleCardMeasurement(event) {
  const element = event.target.closest?.('.batch-card[data-id]');
  if (!element) return;
  dirtyHeights.add(element.dataset.id);
  if (heightFrame !== null) return;
  heightFrame = requestAnimationFrame(() => { heightFrame = null; measureDirtyCards(); });
}

function cardHeight(id) {
  if (cardHeights.has(id)) return cardHeights.get(id);
  return cardHeights.size ? cardHeightSum / cardHeights.size : 520;
}

function shortestColumn(heights) {
  return heights.reduce((best, h, i) => (h < heights[best] - 1 ? i : best), 0);
}

// 已放好的卡片保持原列；列数或筛选变了、或有卡片被移出（提交后）时整体重排。返回这次新放的卡片
function placeCards(groups, count) {
  const meta = JSON.stringify([count, selectedFilter]);
  const ids = new Set(groups.flatMap(([, cards]) => cards.map(c => c.id)));
  if (meta !== columnMeta || [...columnOf.keys()].some(id => !ids.has(id))) {
    columnOf = new Map();
    columnMeta = meta;
  }
  const fresh = new Set();
  for (const [, cards] of groups) {
    const heights = Array(count).fill(0);
    for (const card of cards) if (columnOf.has(card.id)) heights[columnOf.get(card.id)] += cardHeight(card.id);
    for (const card of cards) {
      if (columnOf.has(card.id)) continue;
      const column = shortestColumn(heights);
      columnOf.set(card.id, column);
      heights[column] += cardHeight(card.id);
      fresh.add(card.id);
    }
  }
  return fresh;
}

function cardColumns(cards, count) {
  const columns = Array.from({ length: count }, () => []);
  for (const card of cards) columns[columnOf.get(card.id)].push(card);
  return columns;
}

// 量出卡片真实高度，按真实高度重新给新卡片选列；有变动就把卡片节点挪过去（图片不重新加载）
function rebalanceFresh(groups, count, fresh) {
  if (!fresh.size) return false;
  let moved = false;
  groups.forEach(([key, cards]) => {
    const columns = fileElements.get(key).columns;
    if (!cards.some(c => fresh.has(c.id))) return;
    const heights = Array(count).fill(0);
    for (const card of cards) if (!fresh.has(card.id)) heights[columnOf.get(card.id)] += cardHeight(card.id);
    let changed = false;
    for (const card of cards) {
      if (!fresh.has(card.id)) continue;
      const column = shortestColumn(heights);
      if (columnOf.get(card.id) !== column) changed = true;
      columnOf.set(card.id, column);
      heights[column] += cardHeight(card.id);
    }
    if (!changed) return;
    moved = true;
    cardColumns(cards, count).forEach((column, i) => reconcileChildren(columns[i], column.map(c => cardElements.get(c.id))));
  });
  return moved;
}

// 文件框的标题栏：文件名（点开 Excel）、三色数量、选中本文件、打开、复检这个文件。
// 数量按整个文件算（不随筛选变），复检只对已提交的黄卡有效：复检本来就是整份文件重新检测
function fileHeadHtml(key, cards = batchState.cards.filter(c => fileKey(c) === key), context) {
  const first = cards[0] || {};
  const manual = cards.filter(c => c.color === 'yellow' && c.finalized);
  const jump = manual[0] || first;
  const folder = String(batchState.folder || '').replace(/[\\/]+$/, '');
  const where = String(first.wb_path || '');
  const sub = folder && where.toLowerCase().startsWith(folder.toLowerCase())
    ? where.slice(folder.length).replace(/^[\\/]+/, '').split(/[\\/]/).slice(0, -1).join(' / ') : '';
  const blocked = openBlocked(first);
  const target = { file: first.wb_name, folder: batchState.folder, blocked };
  // 后台任务在跑时，标题栏写明这份文件现在能不能打开（不能时文件名和「打开文件」变灰）
  const openState = !batchState.busy ? ''
    : blocked ? `<span class="batch-file-open" data-open="blocked" title="${esc(blocked)}">🔒 暂不能打开 · ${esc(blocked)}</span>`
    : '<span class="batch-file-open" data-open="ready" title="后台不会再处理这份文件，现在可以在 Excel 中打开">✅ 已处理完，可以打开</span>';
  const openedState = (context?.opened || new Set(batchState.opened_files || [])).has(key)
    ? '<span class="batch-file-open" data-open="opened" title="本轮成功打开过；不代表已修复或复检通过">📂 本轮已打开</span>' : '';
  const counts = REVIEW_COLORS.map(color => `<span class="batch-file-count" data-color="${color}"><span class="batch-dot" aria-hidden="true"></span>${COLOR_NAMES[color]} ${cards.filter(c => c.color === color).length}</span>`).join('');
  const recheckBlock = manual.length ? recheckBlockReason(manual) : '';
  const recheckTitle = recheckBlock ? `现在不能点：${recheckBlock}`
    : `整份文件所有表重新检测，这个文件的 ${manual.length} 张黄卡一起出结果${queueNote('复检')}`;
  const recheck = manual.length
    ? `<button type="button" class="btn btn-outline" data-file-action="recheck" data-file="${esc(key)}"${recheckBlock ? ' disabled' : ''} title="${esc(recheckTitle)}">🔁 复检这个文件</button>`
    : '<button type="button" class="btn btn-outline" disabled title="这个文件还没有已提交的黄色人工项">🔁 复检这个文件</button>';
  // 自动试过没修好、文件没改过的黄卡不算：一张都不能自动时按钮不可点
  const autoable = manual.filter(c => !c.auto_skip);
  const autoBlock = autoable.length ? autoBlockReason(autoable) : '';
  const autoTitle = !manual.length ? '这个文件还没有已提交的黄色人工项'
    : !autoable.length ? `这个文件的 ${manual.length} 张黄卡都不能交给自动（卡片上写了原因），请人工处理`
    : autoBlock ? `现在不能点：${autoBlock}`
    : `让程序自动加高这个文件的 ${autoable.length} 张黄卡（能自动的才调整，确认框里会列出跳过的和原因）${autoQueueNote()}`;
  const auto = `<button type="button" class="btn btn-outline" data-file-action="auto" data-file="${esc(key)}"${!autoable.length || autoBlock ? ' disabled' : ''} title="${esc(autoTitle)}">🤖 自动调整这个文件${autoable.length && autoable.length < manual.length ? `（${autoable.length} 张）` : ''}</button>`;
  const badges = autoQueueBadge(key, cards, context) + followupBadge(key, cards, context);
  const queue = badges.match(/data-queue="(\w+)"/)?.[1];
  return `<header class="batch-file-head"${batchState.busy ? ` data-open="${blocked ? 'blocked' : 'ready'}"` : ''}${queue ? ` data-queue="${queue}"` : ''}>
    <div class="batch-file-title"><input type="checkbox" class="batch-file-checkbox" data-file-check="${esc(key)}" aria-label="选择文件 ${esc(first.wb_name || key)}"><span class="batch-file-icon" aria-hidden="true">📄</span><div class="batch-file-name">${excelLinkHtml(esc(first.wb_name || key), target)}${sub ? `<p class="batch-file-path">${esc(sub)}</p>` : ''}</div></div>${badges}${openState}${openedState}
    <div class="batch-file-counts" aria-label="这个文件的卡片数">${counts}<span class="batch-file-total">共 ${cards.length} 格</span></div>
    <div class="batch-file-actions"><button type="button" class="btn btn-outline batch-file-toggle" data-file-action="toggle" data-file="${esc(key)}" aria-expanded="${expandedFiles.has(key)}">${expandedFiles.has(key) ? '▾ 收起卡片' : `▸ 展开 ${cards.length} 张卡片`}</button><button type="button" class="btn btn-outline" data-file-action="select" data-file="${esc(key)}" aria-pressed="false">选中本文件的卡</button>${excelLinkHtml('📂 打开文件', { ...target, sheet: jump.sheet_name, cell: jump.cell_addr }, 'btn btn-outline')}${recheck}${auto}</div>
  </header>`;
}

// 公共依赖每轮只算一次，文件阻塞原因单独进入该文件的卡片签名。
function cardChromeSig() {
  const job = batchState.job || {}, options = job.options || {};
  return JSON.stringify([batchState.folder, batchState.busy, markWhileBusy(), autoQueueOpen(),
    job.kind, job.closing, job.pause_requested, job.to_manual, options.picked, options.scan_only, options.back_to_auto,
    excelLinkHtml.revision || 0]);
}

function memoHtml(cache, id, signature, build) {
  const previous = cache.get(id);
  if (previous?.signature === signature) return previous.html;
  const html = build();
  cache.set(id, { signature, html });
  return html;
}

function htmlElement(html) {
  const holder = document.createElement('div');
  holder.innerHTML = html;
  return holder.firstElementChild;
}

function reconcileChildren(parent, children) {
  let next = parent.firstElementChild;
  for (const element of children) {
    if (element === next) next = next.nextElementSibling;
    else parent.insertBefore(element, next);
  }
}

let filesContainer = null;
function syncCardList(groups, count, html, heads, fresh) {
  const list = $('cardList');
  const anchor = scrollAnchor();
  if (!filesContainer) {
    filesContainer = document.createElement('div');
    filesContainer.className = 'batch-files';
  }
  const mounted = new Map(groups.map(([key, cards]) => [key, mountedFileCards(key, cards)]));
  const visible = new Set([...mounted.values()].flatMap(cards => cards.map(c => c.id)));
  const owners = new Map(batchState.cards.map(c => [c.id, fileKey(c)]));
  const visibleFiles = new Set(groups.map(([key]) => key));
  for (const [id, element] of cardElements) if (!visible.has(id)) {
    element.remove();
    if (!expandedFiles.has(owners.get(id))) {
      cardElements.delete(id); cardHtmlCache.delete(id); cardMemo.delete(id); forgetCardHeight(id);
    }
  }
  for (const [key, file] of fileElements) if (!visibleFiles.has(key)) file.element.remove();
  for (const [key, allCards] of groups) {
    const cards = mounted.get(key);
    let file = fileElements.get(key);
    if (!file) {
      const element = document.createElement('section');
      element.className = 'batch-file';
      element.dataset.file = key;
      element.setAttribute('aria-label', allCards[0].wb_name || key);
      const head = htmlElement(heads.get(key));
      element.append(head);
      file = { element, head, waterfall: null, body: null, status: null, columns: [] };
      fileElements.set(key, file);
      headHtmlCache.set(key, heads.get(key));
      ++domRevision;
    } else if (headHtmlCache.get(key) !== heads.get(key)) {
      const head = htmlElement(heads.get(key));
      file.head.replaceWith(head);
      file.head = head;
      headHtmlCache.set(key, heads.get(key));
      file.element.setAttribute('aria-label', allCards[0].wb_name || key);
      ++domRevision;
    }
    const expanded = expandedFiles.has(key);
    file.element.classList.toggle('is-collapsed', !expanded);
    if (!expanded) {
      file.body?.remove(); file.body = null; file.waterfall = null; file.status = null; file.columns = [];
      continue;
    }
    if (!file.body) {
      file.body = document.createElement('div'); file.body.className = 'batch-file-body is-opening';
      const body = file.body;
      body.addEventListener('animationend', event => {
        if (event.target === body) body.classList.remove('is-opening');
      });
      file.waterfall = document.createElement('div'); file.waterfall.className = 'batch-waterfall';
      file.status = document.createElement('p'); file.status.className = 'batch-file-loading';
      file.status.setAttribute('role', 'status');
      file.body.append(file.status, file.waterfall); file.element.append(file.body);
    }
    const pending = allCards.some(c => c.details_pending);
    const more = !pending && cards.length < allCards.length;
    file.status.hidden = !pending && !more;
    file.status.textContent = pending ? '正在加载此文件的卡片…' : more ? `正在展开卡片… ${cards.length} / ${allCards.length}` : '';
    file.body.setAttribute('aria-busy', String(pending || more));
    if (more && !fileMountFrames.has(key)) {
      fileMountFrames.set(key, requestAnimationFrame(() => {
        fileMountFrames.delete(key);
        if (!expandedFiles.has(key)) return;
        fileMountLimits.set(key, (fileMountLimits.get(key) ?? FILE_MOUNT_BATCH) + FILE_MOUNT_BATCH);
        renderKey = ''; render();
      }));
    }
    if (file.columns.length !== count) {
      file.columns = Array.from({ length: count }, () => {
        const column = document.createElement('div');
        column.className = 'batch-column';
        return column;
      });
      file.waterfall.replaceChildren(...file.columns);
      cards.forEach(c => dirtyHeights.add(c.id));
      ++domRevision;
    }
    for (const card of cards) {
      if (!cardElements.get(card.id)?.isConnected) dirtyHeights.add(card.id);
      if (!cardElements.has(card.id) || cardHtmlCache.get(card.id) !== html.get(card.id)) {
        const element = htmlElement(html.get(card.id));
        cardElements.get(card.id)?.replaceWith(element);
        cardElements.set(card.id, element);
        cardHtmlCache.set(card.id, html.get(card.id));
        dirtyHeights.add(card.id);
        ++domRevision;
      }
    }
    cardColumns(cards, count).forEach((column, i) => reconcileChildren(file.columns[i], column.map(c => cardElements.get(c.id))));
  }
  reconcileChildren(filesContainer, groups.map(([key]) => fileElements.get(key).element));
  if (groups.length) {
    if (list.firstElementChild !== filesContainer) { list.replaceChildren(filesContainer); ++domRevision; }
    measureDirtyCards();
    rebalanceFresh([...mounted], count, fresh);
  } else {
    const empty = `<section class="panel batch-empty"><h2>${batchState.busy ? '正在自动扫描与修复…' : '当前没有待审核卡片'}</h2><p>${batchState.busy ? '文件处理后会在这里出现修改前后对比，无需逐项放行。' : '开始全自动处理，或切换颜色筛选查看结果。'}</p></section>`;
    if (list.innerHTML !== empty) { list.innerHTML = empty; ++domRevision; }
  }
  restoreAnchor(anchor);
  ++domRevision; // 移动节点和切换筛选也会改变文件全选按钮的状态。
}

function pruneRenderCaches(allByFile) {
  const ids = new Set(batchState.cards.map(c => c.id));
  for (const id of cardMemo.keys()) if (!ids.has(id)) {
    cardMemo.delete(id);
    cardHtmlCache.delete(id);
    cardElements.get(id)?.remove();
    cardElements.delete(id);
    forgetCardHeight(id);
  }
  for (const key of headMemo.keys()) if (!allByFile.has(key)) {
    headMemo.delete(key);
    headHtmlCache.delete(key);
    fileElements.get(key)?.element.remove();
    fileElements.delete(key);
    expandedFiles.delete(key); selectedFiles.delete(key); fileMountLimits.delete(key);
  }
}

function render() {
  renderControls();
  // 只保留当前筛选下看得到的卡片的选中，批量改色不会改到看不见的卡片
  const shown = shownCards();
  const visible = new Set(shown.map(c => c.id));
  for (const id of selectedIds) if (!visible.has(id)) selectedIds.delete(id);
  const p = batchState.progress || {};
  const job = batchState.job;
  // 全自动处理看本任务的文件；审核提交 / 复检等运行中看它自己涉及的文件；其余时候显示上次扫描的进度
  const own = job && Array.isArray(job.workbooks) && (job.kind === 'batch_scan' || (job.kind?.startsWith('batch_') && job.status !== 'completed'));
  // 全自动处理先把全部文件扫完：扫描阶段按扫过的文件算
  const scanning = own && job.kind === 'batch_scan' && !job.options?.scan_only && job.phase !== 'repair';
  const done = own ? new Set((scanning ? job.scanned : job.completed) || []) : null;
  const total = own ? job.workbooks.length : p.total || 0;
  // 已完成只算本任务列表里的文件（去重），不会超过总数
  const completed = own ? job.workbooks.filter(path => done.has(path)).length : Math.min(p.completed || 0, total);
  const label = batchState.busy ? (own && job.kind !== 'batch_scan' ? `正在${REVIEW_JOB_NAMES[job.kind] || '处理'}`
    : job?.to_manual ? '做完当前这一行后转人工' : scanning && job.options?.picked ? '正在准备自动调整（改过的文件重新检测）'
    : job?.options?.picked ? '正在自动调整你选的格子'
    : scanning && job.options?.back_to_auto ? '正在准备转回自动（改过的文件重新检测）'
    : scanning ? '正在扫描（全部扫完后开始修，之后可以转人工）'
    : job?.options?.scan_only ? (job.auto_queue?.length ? '正在扫描（扫完后接着自动调整排队的文件）' : '正在扫描') : '正在自动处理')
    : job?.status === 'interrupted' ? '已暂停或部分文件未完成' : '当前进度';
  $('progressText').textContent = `${label}：${completed} / ${total} 个文件 · 待审核 ${batchState.cards.length} 个单元格${etaText(batchState.scan_eta)}`;
  renderScanNow(batchState.scan_eta?.current);
  renderSlowFiles();
  $('progressFill').style.width = `${total ? Math.round(completed * 100 / total) : 0}%`;
  $('logBox').textContent = batchState.logs.join('\n') || '等待操作…';
  const counts = Object.fromEntries(REVIEW_COLORS.map(c => [c, batchState.cards.filter(card => card.color === c).length]));
  // 筛选按钮用对应颜色的边框、底色和圆点；选中时整块填色
  const filterHtml = ['all', ...REVIEW_COLORS].map(color => `<button type="button" class="batch-filter" data-filter="${color}" aria-pressed="${selectedFilter === color}"><span class="batch-dot" aria-hidden="true"></span>${color === 'all' ? '全部' : REVIEW_LABELS[color]}<span class="batch-count">${color === 'all' ? batchState.cards.length : counts[color]}</span></button>`).join('');
  if ($('filters').innerHTML !== filterHtml) $('filters').innerHTML = filterHtml;
  // 大图是原列表上的弹层，保留列表节点与滚动位置。
  const width = $('cardList').clientWidth || measuredWidth || 1200;
  const columnCount = Math.max(1, Math.min(3, Math.floor((width - 44 + 18) / 328)));
  const allByFile = new Map(fileGroups(batchState.cards)), context = fileRenderContext();
  const chrome = cardChromeSig();
  const groups = fileGroups(shown);
  const mountedGroups = groups.map(([key, cards]) => [key, mountedFileCards(key, cards)]);
  const rendering = mountedGroups.flatMap(([, cards]) => cards);
  const active = largeMode && shown.find(c => c.id === activeCardId);
  if (active && !rendering.includes(active)) rendering.push(active);
  const cardSigs = new Map(rendering.map(card => [card.id, JSON.stringify([card, chrome, openBlocked(card)])]));
  const headSigs = new Map([...allByFile].map(([key, cards]) => [key, JSON.stringify([
    cards.map(c => [c.id, c.wb_path, c.wb_name, c.sheet_name, c.cell_addr, c.color, c.finalized, c.auto_skip]),
    chrome, expandedFiles.has(key), context.opened.has(key), cards.some(c => c.details_pending), openBlocked(cards[0]), autoQueueBadge(key, cards, context), followupBadge(key, cards, context),
  ])]));
  const key = JSON.stringify([chrome, [...cardSigs], [...headSigs], selectedFilter, largeMode, activeCardId, width, [...fileMountLimits]]);
  if (renderKey === key) { applySelection(); return; }
  renderKey = key;
  if (largeMode) {
    // 当前卡片被提交移出列表：停在原位置（即下一张），不跳回第一张
    if (!shown.some(c => c.id === activeCardId)) activeCardId = shown[Math.min(lastActiveIndex, shown.length - 1)]?.id;
    const index = shown.findIndex(c => c.id === activeCardId);
    if (index >= 0) lastActiveIndex = index;
    $('largeIndex').textContent = shown.length ? `${index + 1} / ${shown.length}` : '没有当前颜色的卡片';
    const signature = index >= 0 ? cardSigs.get(shown[index].id) : 'empty';
    if (largeCardKey !== signature) {
      hideImagePreview();
      $('largeCard').innerHTML = index >= 0 ? (shown[index].details_pending
        ? '<p class="batch-file-loading" role="status">正在加载此文件的卡片…</p>' : cardHtml(shown[index], true))
        : '<p class="batch-empty">当前筛选下没有卡片</p>';
      largeCardKey = signature;
      ++domRevision;
    }
    $('prevBtn').disabled = index <= 0;
    $('nextBtn').disabled = index >= shown.length - 1;
  } else {
    // 按文件分组，框里按处理顺序固定排列；改色只换颜色，卡片和文件框都不挪位置。筛选时没有卡片的文件框不显示
    const fresh = placeCards(mountedGroups, columnCount);
    const html = new Map(rendering.map(c => [c.id, memoHtml(cardMemo, c.id, cardSigs.get(c.id), () => cardHtml(c))]));
    const heads = new Map(groups.map(([key]) => [key, memoHtml(headMemo, key, headSigs.get(key), () => fileHeadHtml(key, allByFile.get(key), context))]));
    if (width !== measuredWidth) rendering.forEach(c => dirtyHeights.add(c.id));
    measuredWidth = width;
    syncCardList(groups, columnCount, html, heads, fresh);
  }
  pruneRenderCaches(allByFile);
  applySelection();
}

function openLarge(id) {
  if (!shownCards().some(card => card.id === id)) return;
  largeMode = true;
  activeCardId = id;
  hideImagePreview();
  render();
  if (!$('largeView').open) $('largeView').showModal();
  if (batchState.cards.find(c => c.id === id)?.details_pending) requestFileDetails();
}

function closeLarge() {
  if (!largeMode) return;
  largeMode = false;
  hideImagePreview();
  $('largeView').close();
  render();
}

function hideImagePreview() {
  previewSourceImage = null;
  $('imagePreview').classList.add('hidden');
  $('imagePreview').setAttribute('aria-hidden', 'true');
}

function sizeImagePreview() {
  const image = previewSourceImage;
  if (!image?.naturalWidth || !image.naturalHeight) return;
  const preview = $('imagePreview');
  const enlarged = preview.querySelector('img');
  const rect = image.getBoundingClientRect(), style = getComputedStyle(image);
  // object-fit: contain 可能在小图框内留白，按实际显示的图片内容放大 1.5 倍。
  const displayedScale = Math.min(
    (image.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight)) / image.naturalWidth,
    (image.clientHeight - parseFloat(style.paddingTop) - parseFloat(style.paddingBottom)) / image.naturalHeight);
  const scale = Math.min(displayedScale * 1.5, Math.max(1, innerWidth - 64) / image.naturalWidth,
    Math.max(1, innerHeight - 64) / image.naturalHeight);
  enlarged.style.width = `${image.naturalWidth * scale}px`;
  enlarged.style.height = `${image.naturalHeight * scale}px`;
  // 以小图中心展开，不跟随鼠标；只有超出屏幕边缘时才平移。
  preview.style.left = `${Math.max(8, Math.min(rect.left + rect.width / 2 - preview.offsetWidth / 2, innerWidth - preview.offsetWidth - 12))}px`;
  preview.style.top = `${Math.max(8, Math.min(rect.top + rect.height / 2 - preview.offsetHeight / 2, innerHeight - preview.offsetHeight - 12))}px`;
}

function moveCard(delta) {
  const cards = shownCards();
  const index = cards.findIndex(c => c.id === activeCardId);
  activeCardId = cards[Math.max(0, Math.min(cards.length - 1, index + delta))]?.id;
  render();
  if (batchState.cards.find(c => c.id === activeCardId)?.details_pending) requestFileDetails();
}

// 单击换色马上在页面上变，后台按顺序保存最后选的颜色：连点两下（双击）也能 黄→蓝→绿，不会因为上一次还没保存完被吞掉
const colorQueue = new Map();
let colorSaving = false;
function changeColor(id, direction = 1) {
  const card = batchState.cards.find(c => c.id === id);
  if (requestBusy || !card || card.details_pending || !canMark(card)) return;
  card.color = nextColor(card.color, direction);
  colorQueue.set(id, card.color);
  render();
  saveColors();
}
async function saveColors() {
  if (colorSaving) return;
  colorSaving = true;
  ++pollVersion;
  $('errorText').textContent = '';
  try {
    while (colorQueue.size) {
      const [id, color] = colorQueue.entries().next().value;
      colorQueue.delete(id);
      try { await post('/api/batch/mark', { id, color }); }
      catch (error) { $('errorText').textContent = error.message; }
    }
  } finally {
    colorSaving = false;
    lastEtag = '';
    await poll();
  }
}

// Excel 里的标记颜色：当前配色和历史自定义配色都记在本机浏览器
let highlight = { fill: HIGHLIGHT_PRESETS[0].fill, font: HIGHLIGHT_PRESETS[0].font };
try {
  highlight = normalizeHighlight(JSON.parse(localStorage.getItem(HIGHLIGHT_KEY) || 'null')) || highlight;
} catch (_) {}
let highlightHistory = [];
try {
  const saved = JSON.parse(localStorage.getItem(HIGHLIGHT_HISTORY_KEY) || '[]');
  if (Array.isArray(saved)) highlightHistory = uniqueHighlights(saved)
    .filter(color => !HIGHLIGHT_PRESETS.some(p => sameHighlight(p, color)));
} catch (_) {}
function rememberHighlight(color) {
  if (HIGHLIGHT_PRESETS.some(p => sameHighlight(p, color))) return;
  highlightHistory = uniqueHighlights([color, ...highlightHistory]);
  try { localStorage.setItem(HIGHLIGHT_HISTORY_KEY, JSON.stringify(highlightHistory)); } catch (_) {}
}
// 升级时保留浏览器里已经保存的当前自定义配色。
rememberHighlight(highlight);

// 清除标记结果中的文件名可直接打开 Excel。
function fileNameLink(file, folder) {
  return excelLinkHtml(esc(file.name), { file: file.name, folder });
}

function luminance(hex) {
  const [r, g, b] = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255)
    .map(c => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
const contrastRatio = (a, b) => { const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x); return (hi + 0.05) / (lo + 0.05); };

function renderHighlight() {
  $('highlightFill').value = highlight.fill.toLowerCase();
  $('highlightFont').value = highlight.font.toLowerCase();
  $('highlightSample').style.background = highlight.fill;
  $('highlightSample').style.color = highlight.font;
  $('highlightPresets').innerHTML = HIGHLIGHT_PRESETS.map((p, i) => `<button type="button" class="batch-highlight-preset" data-preset="${i}" aria-pressed="${p.fill === highlight.fill && p.font === highlight.font}" style="background:${p.fill};color:${p.font}">${esc(p.label)}</button>`).join('');
  $('highlightHistoryRow').classList.toggle('hidden', !highlightHistory.length);
  $('highlightHistory').innerHTML = highlightHistory.map((color, i) => `<button type="button" class="batch-highlight-preset" data-history="${i}" aria-pressed="${sameHighlight(color, highlight)}" title="底色 ${color.fill}，字色 ${color.font}；点击使用" style="background:${color.fill};color:${color.font}">底 ${color.fill} / 字 ${color.font}</button>`).join('');
  const ratio = contrastRatio(highlight.fill, highlight.font);
  $('highlightContrast').textContent = `对比度 ${ratio.toFixed(1)}:1${ratio < 4.5 ? '，偏低，字可能看不清' : ''}`;
  $('highlightContrast').classList.toggle('is-low', ratio < 4.5);
}

function setHighlight(next, remember = true) {
  const color = normalizeHighlight(next);
  if (!color) return;
  highlight = color;
  try { localStorage.setItem(HIGHLIGHT_KEY, JSON.stringify(highlight)); } catch (_) {}
  if (remember) rememberHighlight(highlight);
  renderHighlight();
}
$('highlightPresets').addEventListener('click', event => {
  const button = event.target.closest('[data-preset]');
  if (button) setHighlight(HIGHLIGHT_PRESETS[Number(button.dataset.preset)]);
});
$('highlightHistory').addEventListener('click', event => {
  const button = event.target.closest('[data-history]');
  if (button) setHighlight(highlightHistory[Number(button.dataset.history)]);
});
const HIGHLIGHT_AUTO_KEY = 'clear_handle_excel.highlightAuto';
try { $('highlightAuto').checked = localStorage.getItem(HIGHLIGHT_AUTO_KEY) !== '0'; } catch (_) { $('highlightAuto').checked = true; }
$('highlightAuto').addEventListener('change', event => {
  try { localStorage.setItem(HIGHLIGHT_AUTO_KEY, event.target.checked ? '1' : '0'); } catch (_) {}
});
// 拖动取色器时只预览，确认选择后再记历史，避免把每个过渡颜色都存下来。
for (const [id, part] of [['highlightFill', 'fill'], ['highlightFont', 'font']]) {
  $(id).addEventListener('input', event => setHighlight({ ...highlight, [part]: event.target.value }, false));
  $(id).addEventListener('change', event => setHighlight({ ...highlight, [part]: event.target.value }));
}

const FIX_ERROR_TYPES_KEY = 'clear_handle_excel.fixErrorTypes';
const CELL_ERROR_TYPES = [
  { code: 'border_clip', label: '边框压线', fixable: true, hint: '文字被边框横线切穿：自动加高行高' },
  { code: 'text_truncated', label: '文字截断', fixable: true, hint: 'PDF 只印出开头一部分、后半段未显示：自动开启换行并加高行高' },
  { code: 'box_overflow', label: '超出裁剪框', fixable: true, hint: '文字超出单元格裁剪框看不见：自动开启换行并加高行高' },
  { code: 'page_break_cut', label: '分页切开合并单元格', fixable: true, hint: '自动分页落在多行合并区域中间：自动在上方插入手动分页' },
  { code: 'hash_overflow', label: '### 列宽不足', fixable: true, hint: '列宽不够导致数字或日期显示成 ###：自动加宽该列（跨多列合并时平均加宽各列）' },
  { code: 'text_missing', label: 'PDF 未匹配到文字（固定转人工）', fixable: false, hint: 'PDF 中未匹配到该格文字，无法确认是否需要显示，固定转人工' },
];
const FIXABLE_CODES = CELL_ERROR_TYPES.filter(t => t.fixable).map(t => t.code);
let fixErrorTypes = new Set(FIXABLE_CODES);
let hasSavedFixTypes = false;
try {
  const savedTypes = JSON.parse(localStorage.getItem(FIX_ERROR_TYPES_KEY) || 'null');
  if (Array.isArray(savedTypes)) {
    fixErrorTypes = new Set(FIXABLE_CODES.filter(code => savedTypes.includes(code)));
    hasSavedFixTypes = true;
  }
} catch (_) {}

function renderFixErrorTypes() {
  const el = $('fixErrorTypes');
  if (!el) return;
  const html = CELL_ERROR_TYPES.map(t => t.fixable
    ? `<label class="batch-error-type-chip" title="${esc(t.hint)}"><input type="checkbox" data-error-type="${t.code}"${fixErrorTypes.has(t.code) ? ' checked' : ''}><span>${esc(t.label)}</span></label>`
    : `<label class="batch-error-type-chip is-manual-only" title="${esc(t.hint)}"><input type="checkbox" disabled><span>${esc(t.label)}</span></label>`
  ).join('');
  if (el.innerHTML !== html) el.innerHTML = html;
}

$('fixErrorTypes').addEventListener('change', async event => {
  const input = event.target.closest?.('[data-error-type]');
  if (!input) return;
  const code = input.dataset.errorType;
  if (input.checked) fixErrorTypes.add(code); else fixErrorTypes.delete(code);
  hasSavedFixTypes = true;
  try { localStorage.setItem(FIX_ERROR_TYPES_KEY, JSON.stringify([...fixErrorTypes])); } catch (_) {}
  renderFixErrorTypes();
  try {
    await post('/api/batch/fix_types', { fix_error_types: [...fixErrorTypes] });
    await poll();
  } catch (error) { $('errorText').textContent = error.message; }
});

// 暂停的全自动处理在提交、复检、打开文件后会放回来，不算替换续跑点
const KEEPS_PAUSED_SCAN = new Set(['/api/batch/commit', '/api/batch/recheck', '/api/batch/open', '/api/batch/unhighlight']);
const waitingCard = card => card.waiting ?? card.message === '等待自动处理';
const COLOR_NAMES = { blue: '蓝', green: '绿', yellow: '黄' };
const rowKey = card => [card.wb_path, card.sheet_name, card.excel_row, card.kind === 'pagebreak'].join('\0');
// 和后端 commit_selection 一致：共用行高或列宽的卡一起审核，按颜色提交时跳过混色组。
// 后台在跑时排队提交：后台还要处理的文件里的卡（held）这次不交
function commitPlan(color = selectedFilter) {
  const unsettled = batchState.cards.filter(c => !waitingCard(c));
  const held = batchState.busy && markWhileBusy()
    ? unsettled.filter(c => openBlocked(c) && (color === 'all' || c.color === color)) : [];
  const ready = unsettled.filter(c => !held.includes(c));
  if (color === 'all') return { cards: ready, skipped: [], held };
  const parents = new Map(ready.map(c => [c.id, c.id]));
  const root = id => {
    while (parents.get(id) !== id) {
      parents.set(id, parents.get(parents.get(id)));
      id = parents.get(id);
    }
    return id;
  };
  const resources = new Map();
  for (const card of ready) {
    const keys = card.col_width || card.error_types?.includes('hash_overflow')
      ? (card.col_width?.ids || []).map(id => `col:${id}`) : [`row:${rowKey(card)}`];
    for (const key of keys) {
      if (resources.has(key)) parents.set(root(card.id), root(resources.get(key)));
      else resources.set(key, card.id);
    }
  }
  const groups = new Map();
  for (const card of ready) {
    const key = root(card.id);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(card);
  }
  const cards = [], skipped = [];
  for (const peers of groups.values()) {
    const own = peers.filter(c => c.color === color);
    const linked = peers.some(c => c.col_width?.ids?.length || c.row_height?.saved
      || c.page_break?.saved || (c.color === 'blue' && c.adoptable));
    (linked && own.length < peers.length ? skipped : cards).push(...own);
  }
  return { cards, skipped, held };
}

async function startJob(url, body = {}) {
  const keeps = batchState.job?.kind === 'batch_scan' && KEEPS_PAUSED_SCAN.has(url);
  if (batchState.job?.status === 'interrupted' && url !== '/api/resume' && !keeps
      && !(await appConfirm('这会替换上次任务的续跑点。已保存的修复和审核标记会保留，继续吗？', { okText: '继续' }))) return;
  const clearing = url === '/api/batch/unhighlight' || (url === '/api/resume' && batchState.job?.kind === 'batch_unhighlight');
  await action(async () => {
    await post(url, body);
    if (clearing) clearPending = true;
  });
}

// 「🤖 自动调整」：卡片、文件框、多选三个入口都走这里。先问后端哪些能自动、同行连带哪些、哪些跳过，确认后再开始
const cardWhere = id => {
  const card = batchState.cards.find(c => c.id === id);
  return card ? `${card.wb_name} ${card.sheet_name}!${card.cell_addr || '?'}` : id;
};
// 普通打开固定最多五份；文件复选框是明确选择，按勾选的文件打开。
async function openReady(selectedOnly = false) {
  if (requestBusy || openChoosing) return;
  const ids = (selectedOnly ? selectedFileCards() : manualCards()).map(c => c.id);
  if (!ids.length) return;
  openChoosing = true;
  renderControls();
  try {
    const plan = await post('/api/batch/open_plan', { card_ids: ids, selected_files: selectedOnly });
    const waiting = plan.waiting.length ? `\n其中 ${plan.waiting.length} 个还要等后台处理，暂不打开。` : '';
    if (!plan.files.length) {
      await appAlert(plan.exhausted ? '本轮所有人工待处理文件都已打开过，不会自动重复打开。需要重开时，可重置打开记录，或勾选文件后打开。'
        : `共 ${plan.total} 个${selectedOnly ? '选中文件' : '人工待处理文件'}，目前没有可以打开的文件。${waiting}`, { title: '批量打开文件' });
      return;
    }
    const message = (selectedOnly ? `已选择 ${plan.total} 个文件。\n本次将打开 ${plan.files.length} 个选中文件。`
      : `当前有 ${plan.total} 个人工待处理文件，本轮累计已打开 ${plan.round_opened ?? plan.already_opened ?? 0} 个。\n本次最多打开 5 个新文件，实际将打开 ${plan.files.length} 个；完成后还剩 ${plan.remaining} 个未打开，其余不会自动打开。`)
      + waiting + '\n\n本次文件：\n' + plan.files.map(file => `· ${file.name}`).join('\n')
      + (selectedOnly ? '\n\n同时打开较多文件会增加内存占用，可取消后减少勾选。' : '');
    if (!(await appConfirm(message, { title: '确认打开文件', okText: `打开 ${plan.files.length} 个文件`, tone: 'info' }))) return;
    const body = { card_ids: plan.card_ids, selected_files: selectedOnly };
    if (!batchState.busy) { await startJob('/api/batch/open', body); return; }
    let result;
    await action(async () => { result = await post('/api/batch/open', body); });
    if (!result?.opened || (!result.waiting.length && !result.errors.length)) return;
    const list = items => items.map(item => `· ${item.name}：${item.reason}`).join('\n');
    const parts = [result.opened.length ? `已打开 ${result.opened.length} 个文件：${result.opened.join('、')}` : '现在还没有可以打开的文件。'];
    if (result.waiting.length) parts.push(`还要等后台处理的 ${result.waiting.length} 个文件，处理完再打开：\n${list(result.waiting)}`);
    if (result.errors.length) parts.push(`打不开的 ${result.errors.length} 个文件：\n${list(result.errors)}`);
    await appAlert(parts.join('\n\n'), { title: '打开人工待处理文件' });
  } catch (error) { $('errorText').textContent = error.message; }
  finally { openChoosing = false; renderControls(); }
}

async function autoAdjust(ids, title) {
  const queue = batchState.busy;
  const afterScan = queue && batchState.job?.options?.scan_only;
  if (requestBusy || !ids.length || !canAuto(batchState.cards.filter(c => ids.includes(c.id)))) return;
  if (!fixErrorTypes.size) { $('errorText').textContent = '请至少勾选一种要自动修复的错误类型'; return; }
  const selectedTypes = [...fixErrorTypes];
  let plan;
  try { plan = await post('/api/batch/auto_plan', { card_ids: ids, fix_error_types: selectedTypes }); }
  catch (error) { $('errorText').textContent = error.message; return; }
  const total = plan.run.length + plan.extra.length;
  const skipped = plan.skip.length
    ? `\n\n跳过 ${plan.skip.length} 格：\n` + plan.skip.slice(0, 8).map(s => `· ${cardWhere(s.id)}：${s.reason}`).join('\n')
      + (plan.skip.length > 8 ? `\n· ……还有 ${plan.skip.length - 8} 格` : '')
    : '';
  if (!total) { await appAlert(`选中的格子都不能自动调整。${skipped}`, { title }); return; }
  const files = new Set([...plan.run, ...plan.extra].map(id => batchState.cards.find(c => c.id === id)?.wb_path)).size;
  const message = `让程序自动加高 ${total} 格（${files} 个文件）。`
    + (plan.extra.length ? `\n其中 ${plan.extra.length} 格是和你选的格子在同一行的黄卡：一行的行高共用，一起调整。` : '')
    + skipped
    + '\n\n请先在 Excel 里保存并关闭这些文件。扫描后改过的文件先整份重新检测。'
    + '\n这些格子的标记颜色会先去掉，没选的黄卡照旧标着。修好的变蓝等你审核；没修好的仍是黄色，重新标色。'
    + (afterScan ? '\n\n现在正在扫描：这些先排队，全部文件扫完后接着自动调整，排队的文件调整完才能打开。'
      : queue ? '\n\n现在有自动调整正在跑：这些会排在后面，前面的文件做完接着做。' : '')
    + `\n\n确定${queue ? '排队' : '开始'}自动调整吗？`;
  if (!(await appConfirm(message, { title, okText: queue ? '排队自动调整' : '开始自动调整', tone: 'info' }))) return;
  await startJob('/api/batch/auto_adjust', { card_ids: ids, fix_error_types: selectedTypes });
}

async function startBatch(extra = {}) {
  const folder = $('folderInput').value.trim();
  if (!folder) { $('errorText').textContent = '请输入 Excel 文件夹路径'; return; }
  if (!extra.scan_only && !fixErrorTypes.size) {
    $('errorText').textContent = '请至少勾选一种要自动修复的错误类型，或使用「只扫描，全部转人工」';
    return;
  }
  const unsent = batchState.cards.filter(c => !c.finalized).length;
  if (unsent && batchState.folder && folder !== batchState.folder
      && !(await appConfirm(`当前文件夹还有 ${unsent} 张卡片没有提交审核。\n审核记录会保留，以后对原文件夹重新开始全自动处理即可继续审核。\n\n确定换到新文件夹吗？`, { title: '换文件夹', okText: '换到新文件夹' }))) return;
  const limit = Number($('scanLimit').value);  // 空 = 自动估算
  await startJob('/api/batch/start', { folder, fix_error_types: [...fixErrorTypes], ...extra, ...(limit ? { scan_limit_minutes: limit } : {}) }).then(() => {
    if (batchState.busy) { try { localStorage.setItem('clear_handle_excel.lastFolder', folder); } catch (_) {} }
  });
}
const SCAN_LIMIT_KEY = 'clear_handle_excel.scanLimit';
try { $('scanLimit').value = localStorage.getItem(SCAN_LIMIT_KEY) || ''; } catch (_) {}
$('scanLimit').addEventListener('change', event => {
  try { localStorage.setItem(SCAN_LIMIT_KEY, event.target.value); } catch (_) {}
});
$('slowClearBtn').addEventListener('click', () => action(() => post('/api/batch/slow_files/clear')));
$('startBtn').addEventListener('click', () => startBatch($('highlightAuto').checked ? { highlight: { ...highlight } } : {}));
$('scanOnlyBtn').addEventListener('click', async () => {
  if (!$('folderInput').value.trim()) { $('errorText').textContent = '请输入 Excel 文件夹路径'; return; }
  if (!(await appConfirm('不自动加高，检出的问题全部转为黄色人工项。\n会在原 Excel 里给问题单元格换上标记底色和字色并保存（复检通过或提交蓝 / 绿后自动还原）。\n\n确定开始吗？', { title: '只扫描，全部转人工', okText: '开始扫描' }))) return;
  await startBatch({ scan_only: true, highlight: { ...highlight } });
});
$('unhighlightBtn').addEventListener('click', async () => {
  if (await appConfirm(`去掉当前文件夹 ${batchState.highlighted || 0} 个单元格的标记颜色？\n标记是程序加的条件格式，只删这些规则；单元格本身的颜色和表里原有的条件格式都不动。标成橙黄色的工作表标签改回原来的颜色。\n需要先在 Excel 里保存并关闭这些文件。之后提交、复检也不再自动标色，重新开始处理时才按设置标色。`, { title: '清除 Excel 标记颜色', okText: '去掉标记' })) await startJob('/api/batch/unhighlight');
});
$('unhighlightResultBtn').addEventListener('click', showUnhighlightResult);
$('unhighlightCloseBtn').addEventListener('click', () => $('unhighlightDialog').close());
$('resumeBtn').addEventListener('click', () => startJob('/api/resume'));
$('pauseBtn').addEventListener('click', () => action(() => post('/api/batch/pause')));
$('toManualBtn').addEventListener('click', async () => {
  const marking = batchState.job?.options?.highlight ? '，并在 Excel 里给这些格子标色' : '';
  if (!(await appConfirm(`做完当前这一行就停，还没轮到自动处理的问题全部转为黄色人工项${marking}。\n已经修好的蓝卡保留在原文件里，可以随时审核。\n人工处理累了可以点「转回自动」接着修。\n\n确定转人工吗？`, { title: '转人工', okText: '转人工' }))) return;
  await action(() => post('/api/batch/to_manual'));
});
$('toAutoBtn').addEventListener('click', async () => {
  if (!fixErrorTypes.size) { $('errorText').textContent = '请至少勾选一种要自动修复的错误类型'; return; }
  const count = batchState.cards.filter(c => c.handed_over && !c.auto_skip).length || batchState.cards.filter(c => c.handed_over).length;
  if (!(await appConfirm(`把 ${count} 处自动还没试过的黄色问题交回自动处理。\n请先在 Excel 里保存并关闭这些文件。\n转人工后改过的文件整份重新检测；你改过行高的行、自动试过没通过的问题和未勾选的错误类型仍留给人工。\nExcel 里的标记颜色会先还原。\n\n确定转回自动吗？`, { title: '转回自动', okText: '转回自动' }))) return;
  await startJob('/api/batch/to_auto', { fix_error_types: [...fixErrorTypes], ...($('highlightAuto').checked ? { highlight: { ...highlight } } : {}) });
});
$('commitBtn').addEventListener('click', async () => {
  const color = selectedFilter;
  const only = color === 'all' ? '' : COLOR_NAMES[color];
  const waiting = batchState.cards.filter(c => waitingCard(c) && (!only || c.color === color)).length;
  const { cards, skipped, held } = commitPlan(color);
  const queue = batchState.busy;
  const skippedText = skipped.length ? `\n\n另有 ${skipped.length} 张${only}卡和别的颜色的卡共用行高或列宽，这次先不提交。把相关卡片改成同一种颜色，或切到「全部」一起提交（意见不一致时还原相关修改并转黄）。` : '';
  const heldText = held.length ? `\n\n另有 ${held.length} 张在后台还要处理的文件里，这次不提交，后台处理完再提交。` : '';
  if (!cards.length) {
    $('errorText').textContent = skipped.length ? skippedText.trim() : held.length ? heldText.trim()
      : waiting ? '卡片都还没轮到自动处理，暂时没有可提交的标记' : `没有${only}色卡片可提交`;
    return;
  }
  const count = c => cards.filter(card => card.color === c).length;
  const unmarked = cards.filter(c => !c.marked && !c.finalized).length;
  const message = (only ? `只提交 ${cards.length} 张${only}色卡片（${REVIEW_LABELS[color].split(' · ')[1]}），其他颜色这次不动。`
                       : `将提交 ${cards.length} 张卡片：\n蓝（保留修复）${count('blue')}　绿（还原并忽略）${count('green')}　黄（还原后人工）${count('yellow')}`)
    + (cards.some(c => c.color === 'blue' && c.adoptable) ? `\n\n其中 ${cards.filter(c => c.color === 'blue' && c.adoptable).length} 张蓝色是自动检查没通过、由你看图认可的：会按最后一次尝试的行高写入原文件。` : '')
    + (unmarked ? `\n\n其中 ${unmarked} 张没有手动改过颜色，将按当前颜色处理。` : '')
    + skippedText + heldText
    + (waiting ? `\n\n另有 ${waiting} 张还没轮到自动处理，这次不提交，继续全自动处理后再审核。` : '')
    + (queue ? `\n\n后台正在${busyJobName()}：先排队，结束后接着提交。到时这些卡按那时的颜色处理，排队期间还可以改颜色。` : '')
    + `\n\n确定${queue ? '排队' : ''}提交吗？`;
  if (await appConfirm(message, { title: only ? `只处理${only}色标记` : '确认并处理全部标记', okText: queue ? '排队提交' : '提交', tone: 'info' })) {
    await startJob('/api/batch/commit', { card_ids: cards.map(c => c.id), color });
  }
});
$('abandonBtn').addEventListener('click', async () => {
  if (await appConfirm('放弃上次未完成的写入？\n不会再改动这个文件，相关单元格转为黄色人工处理，请在 Excel 中核对这一行的行高。', { title: '放弃这次写入', okText: '放弃写入', tone: 'danger' })) action(() => post('/api/batch/abandon'));
});
$('copyErrorBtn').addEventListener('click', event => {
  const parts = [$('pendingNotice').classList.contains('hidden') ? '' : $('pendingNotice').textContent,
                 $('errorText').textContent.trim()].filter(Boolean);
  const recent = errorLogLines().slice(-10);
  if (recent.length) parts.push('最近的报错日志：\n' + recent.join('\n'));
  copyText(parts.join('\n\n'), event.currentTarget);
});
$('copyLogBtn').addEventListener('click', event => copyText(batchState.logs.join('\n') || '（暂无日志）', event.currentTarget));
$('copyLogErrorsBtn').addEventListener('click', event => copyText(errorLogLines().join('\n') || '（日志里没有报错）', event.currentTarget));
// 报错文字在很多地方被设置/清空，统一盯着它显示或隐藏复制按钮
new MutationObserver(renderCopyErrorBtn).observe($('errorText'), { childList: true, characterData: true, subtree: true });
new MutationObserver(renderCopyErrorBtn).observe($('pendingNotice'), { attributes: true, attributeFilter: ['class'] });
$('openAllBtn').addEventListener('click', () => openReady(false));
$('resetOpenHistoryBtn').addEventListener('click', async () => {
  if (requestBusy || openChoosing) return;
  if (await appConfirm('重置当前文件夹的打开记录？下一次将从第一个待处理文件开始。\n已打开的 Excel 不会关闭，审核颜色和修复结果不变。', { title: '从头开始打开', okText: '重置记录' })) {
    await action(() => post('/api/batch/open_reset'));
  }
});
$('openSelectedFilesBtn').addEventListener('click', () => openReady(true));
$('clearSelectedFilesBtn').addEventListener('click', () => { selectedFiles.clear(); applySelection(); renderControls(); });
$('openPickedFilesBtn').addEventListener('click', () => openReady(true));
$('clearPickedFilesBtn').addEventListener('click', () => { selectedFiles.clear(); applySelection(); renderControls(); });
$('recheckAllBtn').addEventListener('click', () => startJob('/api/batch/recheck', { card_ids: recheckAllCards().map(c => c.id) }));
$('followupStartBtn').addEventListener('click', () => action(() => post('/api/batch/followups/start')));
$('followupCancelBtn').addEventListener('click', async () => {
  if (await appConfirm('取消还没开始的排队复检 / 提交？卡片的颜色和标记都不变，之后可以再点。', { title: '取消排队', okText: '取消排队' })) {
    await action(() => post('/api/batch/followups/cancel'));
  }
});
$('closeLargeBtn').addEventListener('click', closeLarge);
$('largeView').addEventListener('cancel', event => { event.preventDefault(); closeLarge(); });
$('prevBtn').addEventListener('click', () => moveCard(-1));
$('nextBtn').addEventListener('click', () => moveCard(1));
$('selectBar').addEventListener('click', event => {
  const button = event.target.closest('button');
  if (button?.dataset.select) selectShown(button.dataset.select);
  else if (button?.dataset.setColor) changeSelectedColor(button.dataset.setColor);
  else if (button?.id === 'autoSelectedBtn') autoAdjust(manualCards().filter(c => selectedIds.has(c.id)).map(c => c.id), '自动调整选中的格子');
});
$('filters').addEventListener('click', event => {
  const button = event.target.closest('[data-filter]');
  if (button) { selectedFilter = button.dataset.filter; filterKeep = new Set(); renderKey = ''; render(); }
});

for (const id of ['cardList', 'largeCard']) $(id).addEventListener('click', event => {
  const fileButton = event.target.closest('[data-file-action]');
  if (fileButton) {
    const key = fileButton.dataset.file;
    if (fileButton.dataset.fileAction === 'toggle') toggleFileExpanded(key);
    else if (fileButton.dataset.fileAction === 'select') toggleFileSelected(key);
    else if (fileButton.dataset.fileAction === 'recheck') {
      startJob('/api/batch/recheck', { card_ids: manualCards().filter(c => fileKey(c) === key).map(c => c.id) });
    } else if (fileButton.dataset.fileAction === 'auto') {
      autoAdjust(manualCards().filter(c => fileKey(c) === key && !c.auto_skip).map(c => c.id), '自动调整这个文件');
    }
    return;
  }
  const head = event.target.closest('.batch-file-head');
  if (head && !event.target.closest('a, button, input, label')) {
    toggleFileExpanded(head.closest('.batch-file').dataset.file);
    return;
  }
  const card = event.target.closest('[data-id]');
  if (!card || event.target.closest('a')) return;
  const button = event.target.closest('[data-action]');
  if (button?.dataset.action === 'select') toggleSelected(card.dataset.id);
  else if (button?.dataset.action === 'recheck') startJob('/api/batch/recheck', { card_ids: [card.dataset.id], single: true });
  else if (button?.dataset.action === 'auto') autoAdjust([card.dataset.id], '自动调整这一格');
  else changeColor(card.dataset.id);
});

$('cardList').addEventListener('change', event => {
  const checkbox = event.target.closest('[data-file-check]');
  if (checkbox) toggleFileChecked(checkbox.dataset.fileCheck);
});

document.addEventListener('keydown', event => {
  // Alt 必须先于通用修饰键过滤处理，且只认空格左边的 Alt，避免右 Alt / AltGr 和组合快捷键误触。
  if (event.isComposing || event.target.closest('input, textarea, select, [contenteditable]:not([contenteditable="false"])')
    || event.ctrlKey || event.metaKey || event.shiftKey || document.querySelector('dialog[open]:not(#largeView)')) return;
  if (event.code === 'AltLeft') {
    // 按键时重新命中鼠标下的卡片：滚动、筛选或后台更新后也不会打开旧位置的卡片。
    const card = !largeMode && cardPointer
      ? document.elementFromPoint(cardPointer.x, cardPointer.y)?.closest('#cardList .batch-card[data-id]') : null;
    if (!largeMode && !card) return;
    event.preventDefault();
    leftAltHandled = true;
    if (!event.repeat) largeMode ? closeLarge() : openLarge(card.dataset.id);
    return;
  }
  if (event.altKey) return;
  if (largeMode) {
    if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Escape'].includes(event.key)) event.preventDefault();
    if (event.key === 'ArrowUp') moveCard(-1);
    if (event.key === 'ArrowDown') moveCard(1);
    if (event.key === 'ArrowLeft') changeColor(activeCardId, -1);
    if (event.key === 'ArrowRight') changeColor(activeCardId, 1);
    if (event.key === 'Escape') closeLarge();
  } else if (['Enter', ' '].includes(event.key) && event.target.matches('.batch-card')) {
    event.preventDefault(); changeColor(event.target.dataset.id);
  }
});
document.addEventListener('keyup', event => {
  if (event.code === 'AltLeft' && leftAltHandled) { event.preventDefault(); leftAltHandled = false; }
});

document.addEventListener('pointermove', event => {
  cardPointer = { x: event.clientX, y: event.clientY };
  const image = largeMode && event.target.closest('#largeCard .batch-shots img[data-preview]');
  const preview = $('imagePreview');
  if (!image) { hideImagePreview(); return; }
  previewSourceImage = image;
  const enlarged = preview.querySelector('img');
  if (enlarged.src !== image.src) enlarged.src = image.src;
  preview.classList.remove('hidden');
  preview.setAttribute('aria-hidden', 'false');
  sizeImagePreview();
});
$('imagePreview').querySelector('img').addEventListener('load', () => {
  if (!$('imagePreview').classList.contains('hidden')) sizeImagePreview();
});
document.addEventListener('pointerout', event => {
  if (!event.relatedTarget) { cardPointer = null; hideImagePreview(); }
});
window.addEventListener('blur', () => { cardPointer = null; leftAltHandled = false; hideImagePreview(); });
document.addEventListener('scroll', hideImagePreview, true);
$('cardList').addEventListener('load', scheduleCardMeasurement, true);
$('cardList').addEventListener('error', scheduleCardMeasurement, true);
window.addEventListener('resize', () => { hideImagePreview(); renderKey = ''; render(); });
try { $('folderInput').value = localStorage.getItem('clear_handle_excel.lastFolder') || ''; } catch (_) {}
renderHighlight();
renderFixErrorTypes();
poll();
setInterval(() => poll(true), 1800);
