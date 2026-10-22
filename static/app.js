let latestPages = {};
let serverItems = {}; // 服务器返回的问题项
let stateLoaded = false; // 至少拿到过一次服务器状态
let latestItems = {}; // 叠加尚未保存的跳过决定后的问题项，页面以此渲染
let workspaceRestored = false;
let canResume = false;
let resumeScan = false;
let isBusy = false;
let scanRunning = false;
let scanProgress = {};
let pollSeq = 0;
let minAcceptedSeq = 0; // 发起操作前已在途的轮询结果可能过期，直接丢弃
let lastAppliedPollSeq = 0;

// 与首页、清理 Sheet 共用同一个键：任一处填过的文件夹路径，其他地方自动带出
const FOLDER_STORAGE_KEY = 'clear_handle_excel.lastFolder';

function readLastFolder() {
  try {
    return localStorage.getItem(FOLDER_STORAGE_KEY) || '';
  } catch (e) {
    return '';
  }
}

function saveLastFolder(folder) {
  try {
    localStorage.setItem(FOLDER_STORAGE_KEY, folder);
  } catch (e) {}
}

// 点击即跳过：先在页面上生效，再把这期间的多次点选合并成一次提交（每次写库较慢）
let pendingSkips = new Map(); // item id -> 是否跳过（尚未提交）
let inFlightSkips = new Map(); // 正在提交的决定
let skipRequest = null;
let skipCancelsResumeConfirmed = false;

// 保留判断存入服务器；尚未保存的决定在本地叠加，完整的一行才能交给自动处理。
let triageKept = new Set();
let pendingKeeps = new Map();
let inFlightKeeps = new Map();
let keepRequest = null;
let legacyKeeps = new Set();
let triageHistory = []; // [{ id, action: 'keep' | 'skip' }]，用于撤销
let triageCurrentId = null;
let triageDrag = null;
let triageAnimationTimer = null;
const TRIAGE_SWIPE_THRESHOLD = 110;

// 第 2 步自动处理：后台为每一行准备的加高截图（分组），以及当前显示的那一组
let stepTab = 'triage'; // 'triage' | 'fix'
let fixData = { auto: false, groups: [] };
let fixCurrentId = null;
let fixDeferred = new Set(); // 点了「下一个」暂时跳过的组
let fixChosen = new Set(); // 已点选、等待写入的组
let fixSkipping = new Set(); // 正在保存“本行无需修改”判断的组
let manualAction = null; // 打开或复检的任务，等待后端完成后才解除按钮锁定
// 稍后人工复检：挂起的文件（服务器 review_queue）和一次打开 / 统一复检的任务
let reviewQueue = {};
let reviewAction = null; // { kind: 'review_open' | 'review_recheck', submitted }
let fixDeferring = new Set(); // 正在挂起到稍后人工复检的组
// 默认全自动：已写入、等你确认的行。卡片按第一次出现的顺序排，之后不再挪动
let confirmOrder = [];
let confirmBusy = new Set(); // 正在提交改图 / 退回的组
let confirmSubmitting = false;

const TRIAGE_KEPT_STORAGE_KEY = 'clear_handle_excel.triageKept';
const RECHECK_TITLE = '请先在 Excel 中保存手动修改。当前工作表只导出一次 PDF 复检，不会自动增加行高';

function isResolved(item) {
  return item.status === 'fixed' || item.status === 'accepted';
}

function processingBusy() {
  return isBusy && !scanRunning;
}

// 隐藏截断：PDF 外观看不出来、实际没显示全的文字
function isHidden(item) {
  return item.kind === 'hidden';
}

function itemTextLabel(item) {
  if (!isHidden(item)) return `文字: ${item.text}`;
  const shown = item.visible_text ? `PDF 中只显示「${item.visible_text}」` : 'PDF 中没有显示';
  return `【隐藏截断】完整内容「${item.text}」，${shown}`;
}

function pageLabel(item) {
  return item.page_index === null || item.page_index === undefined ? '?' : Number(item.page_index) + 1;
}

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getRowGroupKey(item) {
  return `${item.wb_path}::${item.sheet_name}::${item.excel_row}`;
}

function getPeerItemsInSameRow(targetItem) {
  const key = getRowGroupKey(targetItem);
  return Object.values(latestItems).filter(it => getRowGroupKey(it) === key);
}

/* ==========================================
   人工筛选：点击即跳过 / 撤销
   ========================================== */
function applyPendingSkips() {
  latestItems = {};
  for (const [id, item] of Object.entries(serverItems)) {
    const decisions = pendingSkips.has(id) ? pendingSkips : inFlightSkips;
    if (!decisions.has(id)) {
      latestItems[id] = item;
      continue;
    }
    const skip = decisions.get(id);
    const unskipped = item.status === 'accepted'
      ? (item.manual_confirmation?.previous_status || 'pending')
      : item.status;
    latestItems[id] = { ...item, status: skip ? 'accepted' : unskipped };
  }
  triageKept = new Set(Object.values(latestItems).filter(item => {
    const keeps = pendingKeeps.has(item.id) ? pendingKeeps : inFlightKeeps;
    return keeps.has(item.id) ? keeps.get(item.id) : item.triage_keep;
  }).map(item => item.id));
}

// 返回是否已生效（任务执行中或用户取消时返回 false）
function setSkip(id, skip) {
  const item = latestItems[id];
  if (!item || item.status === 'fixed') return false;
  if (processingBusy()) {
    alert('任务执行中，请结束后再标记跳过。');
    return false;
  }
  if (canResume && !resumeScan && !skipCancelsResumeConfirmed) {
    // 跳过本身可随时撤销，只有会取消中断任务续跑点（不可撤销）时才提示
    if (!confirm('标记跳过会取消上次中断任务的续跑，避免旧检查点覆盖人工判断。确定吗？')) return false;
    skipCancelsResumeConfirmed = true;
  }
  const baseline = inFlightSkips.has(id) ? inFlightSkips.get(id) : serverItems[id]?.status === 'accepted';
  if (skip === baseline) {
    pendingSkips.delete(id);
  } else {
    pendingSkips.set(id, skip);
  }
  applyPendingSkips();
  renderItems();
  flushSkips();
  return true;
}

function flushSkips() {
  if (skipRequest || pendingSkips.size === 0) return skipRequest;
  inFlightSkips = pendingSkips;
  pendingSkips = new Map();
  skipRequest = (async () => {
    try {
      const data = await postAction('/api/manual/skip', { decisions: Object.fromEntries(inFlightSkips) });
      if (data.errors.length) alert(data.errors.join('\n'));
    } catch (e) {
      alert(`标记跳过失败：${e.message}`);
    }
    // 取得服务器最新状态后再去掉叠加，避免界面闪回
    minAcceptedSeq = pollSeq + 1;
    await pollState();
    inFlightSkips = new Map();
    skipRequest = null;
    applyPendingSkips();
    renderItems();
    flushSkips();
  })();
  return skipRequest;
}

// 发起后台任务前先保存所有跳过决定，确保复检和重新扫描看到的是最新筛选结果
async function syncSkips() {
  while (skipRequest || pendingSkips.size) {
    await (skipRequest || flushSkips());
  }
  while (keepRequest || pendingKeeps.size) await (keepRequest || flushKeeps());
}

function setKeep(id, keep) {
  const item = latestItems[id];
  if (!item || isResolved(item) || processingBusy() || manualAction) return false;
  if (canResume && !resumeScan && !skipCancelsResumeConfirmed) {
    if (!confirm('保留判断会取消上次中断任务的续跑，确定吗？')) return false;
    skipCancelsResumeConfirmed = true;
  }
  const group = fixData.groups.find(g => g.item_ids.includes(id));
  if (group?.status === 'applying' || fixChosen.has(group?.id)) {
    alert('这一行已开始写入，不能撤销筛选。');
    return false;
  }
  pendingKeeps.set(id, keep);
  if (keep) triageKept.add(id); else triageKept.delete(id);
  saveTriageKept();
  flushKeeps();
  return true;
}

function flushKeeps() {
  if (keepRequest || !pendingKeeps.size) return keepRequest;
  inFlightKeeps = pendingKeeps;
  pendingKeeps = new Map();
  keepRequest = (async () => {
    try {
      await postAction('/api/manual/keep', {decisions: Object.fromEntries(inFlightKeeps)});
    } catch (e) {
      alert(`保存保留判断失败：${e.message}`);
    }
    minAcceptedSeq = pollSeq + 1;
    await pollState();
    inFlightKeeps = new Map();
    keepRequest = null;
    applyPendingSkips();
    saveTriageKept();
    renderItems();
    pollAutofix();
    flushKeeps();
  })();
  return keepRequest;
}

function getItemVisualState(item) {
  const tagNo = item.index_on_page || item.seq;
  if (item.status === 'accepted') {
    return {
      boxClass: 'is-ignored is-resolved', cardClass: 'skipped is-resolved', badgeClass: 'badge-ignored',
      badgeText: '已跳过（无需修改）', shortTag: `跳过#${tagNo}`,
    };
  }
  if (item.status === 'fixed') {
    return {
      boxClass: 'is-fixed is-resolved',
      cardClass: 'fixed is-resolved',
      badgeClass: 'badge-fixed',
      badgeText: item.iterations ? `✓ 已修复 (+${item.iterations}次)` : '✓ 已修复（复检通过）',
      shortTag: `✓#${tagNo}`,
    };
  }
  const newTag = item.is_new ? '新' : '';
  if (item.status === 'needs_review' || item.status === 'manual_review') {
    return {
      boxClass: 'is-max-reached',
      cardClass: 'max-reached',
      badgeClass: 'badge-warn',
      badgeText: item.status === 'manual_review' ? '已在 Excel 中定位，修改并保存后重新扫描' : '需人工复检',
      shortTag: `${newTag}?#${tagNo}`,
    };
  }
  if (item.status === 'max_reached') {
    return {
      boxClass: 'is-max-reached',
      cardClass: 'max-reached',
      badgeClass: 'badge-warn',
      badgeText: `⚠️ 达到上限，建议人工复检 (${item.iterations}次)`,
      shortTag: `⚠#${tagNo}`,
    };
  }
  return {
    boxClass: 'is-selected',
    cardClass: 'selected',
    badgeClass: item.is_new ? 'badge-pending' : 'badge-selected',
    badgeText: item.is_new ? '🆕 新发现 · 待人工修改' : '待人工修改（点击可跳过）',
    shortTag: `${newTag}#${tagNo}`,
  };
}

// open = 还需人工处理（含已在 Excel 中定位等状态）；isNew = 其中本次复检新发现的
function countByState(items) {
  const counts = { open: 0, isNew: 0, skipped: 0, fixed: 0 };
  items.forEach(item => {
    if (item.status === 'accepted') counts.skipped += 1;
    else if (item.status === 'fixed') counts.fixed += 1;
    else {
      counts.open += 1;
      if (item.is_new) counts.isNew += 1;
    }
  });
  return counts;
}

function refreshCounts() {
  const totals = countByState(Object.values(latestItems));
  document.getElementById('cntPendingFix').innerText = totals.open;
  document.getElementById('cntNew').innerText = totals.isNew;
  document.getElementById('cntSkipped').innerText = totals.skipped;
  document.getElementById('cntFixed').innerText = totals.fixed;
}

function renderItems() {
  refreshCounts();
  renderTriage();
  renderStepTabs();
  renderFix();
  renderPipelineProgress();
}

function pipelineMetrics() {
  const items = Object.values(latestItems);
  const screened = items.filter(it => isResolved(it) || triageKept.has(it.id)).length;
  const rows = new Map();
  for (const item of items) {
    const key = getRowGroupKey(item);
    if (!rows.has(key)) rows.set(key, []);
    rows.get(key).push(item);
  }
  const handed = [...rows.values()].filter(peers => peers.some(it => triageKept.has(it.id))
    && peers.every(it => isResolved(it) || triageKept.has(it.id)));
  // 挂到稍后人工复检的行，② 这一步已经处理完
  const done = handed.filter(peers => peers.every(it => isResolved(it) || it.later_review)).length;
  const ready = fixQueue().filter(g => ['ready', 'failed', 'error'].includes(g.status)).length
    + confirmGroups().filter(g => g.status !== 'reverted').length;
  return {total: items.length, screened, available: items.length - screened, handed: handed.length, done, ready};
}

function renderPipelineProgress() {
  const m = pipelineMetrics();
  const total = Number(scanProgress.total || 0);
  const completed = Number(scanProgress.completed || 0);
  document.getElementById('scanPipelineText').textContent = total
    ? `扫描：已完成 ${completed} / ${total} 个文件 · 剩余 ${Math.max(0, total - completed)} 个${scanProgress.status === 'interrupted' ? '（有文件失败，可继续上次任务）' : ''}`
    : (scanRunning ? '扫描正在准备文件列表…' : '扫描：等待开始');
  document.getElementById('scanPipelineFill').style.width = `${total ? Math.min(100, completed / total * 100) : 0}%`;
  document.getElementById('supplyProgressText').textContent =
    `当前已产出 ${m.total} 处 · 已筛选 ${m.screened} 处 · 可筛选 ${m.available} 处`;
  document.getElementById('supplyProgressFill').style.width = `${m.total ? m.screened / m.total * 100 : 0}%`;
  document.getElementById('supplyProgressHint').textContent = scanRunning
    ? '扫描还在继续；这里只统计已完成文件，后续产出会自动加入。'
    : (scanProgress.status === 'interrupted' ? '已完成文件可继续处理；其余文件等待重试扫描。' : '当前结果可随时筛选，已筛完的行会在后台准备截图。');
  document.getElementById('fixSupplyText').textContent =
    `已交接 ${m.handed} 行 · 待你确认 ${m.ready} 行 · 已完成 ${m.done} 行`;
  document.getElementById('fixSupplyFill').style.width = `${m.handed ? m.done / m.handed * 100 : 0}%`;
}

/* ==========================================
   逐个筛选（← 保留 / 跳过 →）
   ========================================== */
function loadTriageKept() {
  try {
    return new Set(JSON.parse(localStorage.getItem(TRIAGE_KEPT_STORAGE_KEY) || '[]'));
  } catch (e) {
    return new Set();
  }
}

function saveTriageKept() {
  try {
    localStorage.setItem(TRIAGE_KEPT_STORAGE_KEY, JSON.stringify([...triageKept]));
  } catch (e) {}
}

// 尚未判断的问题，按扫描顺序排列；跳过的已解决，保留的记在 triageKept
function triageQueue() {
  return Object.values(latestItems)
    .filter(item => !isResolved(item) && !triageKept.has(item.id))
    .sort((a, b) => (a.seq || 0) - (b.seq || 0));
}

function renderTriage() {
  const reviewable = Object.values(latestItems);
  const skipped = reviewable.filter(item => item.status === 'accepted').length;
  const kept = reviewable.filter(item => !isResolved(item) && triageKept.has(item.id)).length;
  const queue = triageQueue();
  const decided = reviewable.length - queue.length;
  document.getElementById('triageProgressText').textContent =
    `当前已产出 ${reviewable.length} 处 · 已筛选 ${decided} 处 · 可筛选 ${queue.length} 处（保留 ${kept} · 跳过 ${skipped}）`;
  document.getElementById('triageProgressFill').style.width =
    `${reviewable.length ? Math.round((decided / reviewable.length) * 100) : 0}%`;
  document.getElementById('triageUndoBtn').disabled = triageHistory.length === 0;

  // 拖动或飞出动画进行中不换卡
  if (triageDrag || triageAnimationTimer) return;
  const stage = document.getElementById('triageStage');
  const next = queue[0];
  if (!next) {
    triageCurrentId = null;
    const signature = `done|${kept}|${skipped}|${reviewable.length}|${scanRunning}|${scanProgress.status}`;
    if (stage.dataset.signature === signature) return;
    stage.dataset.signature = signature;
    stage.innerHTML = reviewable.length ? `
      <div class="triage-done">
        <b>${scanRunning ? '当前结果已筛选完，等待扫描产生新结果…' : '当前结果已筛选完'}</b>
        <span>保留 ${kept} 处、跳过 ${skipped} 处。保留的到第 ② 步自动处理。</span>
        ${kept ? '<button type="button" class="btn btn-primary btn-sm" onclick="setStepTab(\'fix\')">进入 ② 自动处理 →</button>' : ''}
        ${kept ? '<button type="button" class="btn btn-outline btn-sm" onclick="triageRestartKept()">重新筛选已保留的项</button>' : ''}
      </div>` : `<div class="triage-done"><span>${scanRunning ? '扫描进行中，等待第一个文件的检测结果…' : '暂无可筛选的问题。'}</span></div>`;
    return;
  }

  const page = latestPages[next.page_id];
  const signature = `${next.id}|${next.status}|${next.image}|${page ? page.image.length : 0}|${isBusy}|${!!manualAction}`;
  if (stage.dataset.signature === signature) return;
  stage.dataset.signature = signature;
  triageCurrentId = next.id;

  const st = getItemVisualState(next);
  const peers = getPeerItemsInSameRow(next).length;
  let pageHtml = '';
  if (page) {
    const pw = Math.max(1, Number(page.page_width) || 595);
    const ph = Math.max(1, Number(page.page_height) || 842);
    const [x0, y0, x1, y1] = next.bbox;
    const box = [
      `left:${Math.max(0, ((x0 - 4) / pw) * 100).toFixed(3)}%`,
      `top:${Math.max(0, ((y0 - 3) / ph) * 100).toFixed(3)}%`,
      `width:${Math.min(100, ((x1 - x0 + 8) / pw) * 100).toFixed(3)}%`,
      `height:${Math.min(100, ((y1 - y0 + 6) / ph) * 100).toFixed(3)}%`,
    ].join(';');
    pageHtml = `
      <div class="triage-page" title="这处问题在整页中的位置">
        <img src="${page.image}" alt="整页位置" draggable="false">
        <div class="triage-page-box" style="${box}"></div>
      </div>`;
  }
  stage.innerHTML = `
    <div class="triage-card ${next.status === 'pending' ? '' : 'is-review'}" id="triageCard">
      <div class="triage-stamp triage-stamp-keep">保留</div>
      <div class="triage-stamp triage-stamp-skip">跳过</div>
      <div class="triage-card-head">
        <div>
          <b>#${Number(next.seq || 1)} ${excelLinkHtml(escapeHtml(next.wb_name), { file: next.wb_path, sheet: next.sheet_name, cell: next.cell_addr })}</b>
          <span style="color:#64748b;">/ ${escapeHtml(next.sheet_name)}（第 ${pageLabel(next)} 页）</span>
        </div>
        <span class="badge ${st.badgeClass}">${escapeHtml(next.status === 'pending' ? (next.is_new ? '🆕 新发现 · 待筛选' : '待筛选') : st.badgeText)}</span>
      </div>
      <div class="triage-card-body">
        ${next.image ? `<div class="triage-crop"><img src="${next.image}" alt="2.5x 局部特写" draggable="false"></div>` : ''}
        ${pageHtml}
      </div>
      <div class="triage-card-foot">
        定位 <b>${escapeHtml(next.cell_addr)}</b>${next.location_confirmed === false ? '<span style="color:#b45309;">（定位未确认，请核对行号）</span>' : ''}（Excel 第 ${Number(next.excel_row) || 1} 行 · 行高 ${escapeHtml(next.row_height)}pt${peers > 1 ? ` · 同行共 ${peers} 处切字` : ''}）<br>
        ${escapeHtml(itemTextLabel(next))}
        ${isHidden(next) ? `<br><span style="color:#b45309;">${escapeHtml(next.hidden_reason || '')}</span>` : ''}
        ${next.status !== 'pending' && next.review_reason ? `<br><span style="color:#b45309;">${escapeHtml(next.review_reason)}</span>` : ''}
      </div>
      <div class="triage-card-tools" onpointerdown="event.stopPropagation();">
        <label title="如果行号有偏差可手动修正">
          行号
          <input type="number" min="1" ${isBusy ? 'disabled' : ''} value="${Number(next.excel_row) || 1}" onchange="updateRow('${escapeHtml(next.id)}', this.value)">
        </label>
        <button type="button" class="btn btn-outline btn-sm" ${isBusy || manualAction ? 'disabled' : ''} title="${scanRunning ? '人工打开需等待扫描结束；筛选和选图可继续' : ''}" onclick="manualOpen('${escapeHtml(next.id)}')">在 Excel 中定位</button>
        <button type="button" class="btn btn-outline btn-sm" ${isBusy || manualAction ? 'disabled' : ''} title="${RECHECK_TITLE}" onclick="manualRecheck('${escapeHtml(next.id)}')">重新导出复检</button>
      </div>
    </div>
  `;
  attachTriageSwipe(document.getElementById('triageCard'));
}

function resetTriageCard(card) {
  if (!card) return;
  card.style.transform = '';
  card.style.setProperty('--keep-opacity', 0);
  card.style.setProperty('--skip-opacity', 0);
}

function finishTriageAnimation() {
  if (!triageAnimationTimer) return;
  clearTimeout(triageAnimationTimer);
  triageAnimationTimer = null;
  renderTriage();
}

function triageDecide(action) {
  finishTriageAnimation();
  const item = latestItems[triageCurrentId];
  if (!item || isResolved(item)) return;
  const card = document.getElementById('triageCard');
  // 先占住动画计时，避免下面标记跳过触发的重新渲染立刻换卡
  triageAnimationTimer = setTimeout(finishTriageAnimation, 180);
  if (action === 'skip') {
    if (!setSkip(item.id, true)) {
      clearTimeout(triageAnimationTimer);
      triageAnimationTimer = null;
      resetTriageCard(card);
      return;
    }
  } else {
    if (!setKeep(item.id, true)) {
      clearTimeout(triageAnimationTimer);
      triageAnimationTimer = null;
      resetTriageCard(card);
      return;
    }
  }
  triageHistory.push({ id: item.id, action });
  if (card) {
    card.style.transform = '';
    card.style.setProperty(action === 'skip' ? '--skip-opacity' : '--keep-opacity', 1);
    card.classList.add(action === 'skip' ? 'fly-right' : 'fly-left');
  }
  renderTriage();
  renderStepTabs();
  renderPipelineProgress();
}

function triageUndo() {
  finishTriageAnimation();
  const last = triageHistory.pop();
  if (!last) return;
  if (last.action === 'skip' && latestItems[last.id]) {
    if (!setSkip(last.id, false)) {
      triageHistory.push(last);
      return;
    }
  } else {
    if (!setKeep(last.id, false)) {
      triageHistory.push(last);
      return;
    }
  }
  renderTriage();
}

function triageRestartKept() {
  for (const id of [...triageKept]) if (latestItems[id] && !isResolved(latestItems[id])) setKeep(id, false);
  saveTriageKept();
  triageHistory = [];
  renderTriage();
}

function attachTriageSwipe(card) {
  card.addEventListener('pointerdown', e => {
    if (e.button !== 0 || triageAnimationTimer) return;
    triageDrag = { startX: e.clientX, dx: 0 };
    card.setPointerCapture(e.pointerId);
    card.classList.add('dragging');
  });
  card.addEventListener('pointermove', e => {
    if (!triageDrag) return;
    const dx = e.clientX - triageDrag.startX;
    triageDrag.dx = dx;
    card.style.transform = `translateX(${dx}px) rotate(${dx / 25}deg)`;
    card.style.setProperty('--keep-opacity', Math.min(1, Math.max(0, -dx / TRIAGE_SWIPE_THRESHOLD)));
    card.style.setProperty('--skip-opacity', Math.min(1, Math.max(0, dx / TRIAGE_SWIPE_THRESHOLD)));
  });
  const end = () => {
    if (!triageDrag) return;
    const { dx } = triageDrag;
    triageDrag = null;
    card.classList.remove('dragging');
    if (Math.abs(dx) >= TRIAGE_SWIPE_THRESHOLD) {
      triageDecide(dx < 0 ? 'keep' : 'skip');
    } else {
      resetTriageCard(card);
      renderTriage(); // 拖动期间若有数据更新，补一次渲染
    }
  };
  card.addEventListener('pointerup', end);
  card.addEventListener('pointercancel', end);
}

document.addEventListener('keydown', e => {
  if (stepTab !== 'triage' || e.repeat || e.altKey) return;
  if (e.target.closest && e.target.closest('input, textarea, select')) return;
  if (e.key === 'ArrowLeft') {
    e.preventDefault();
    triageDecide('keep');
  } else if (e.key === 'ArrowRight') {
    e.preventDefault();
    triageDecide('skip');
  } else if (e.key === 'Backspace' || ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z')) {
    e.preventDefault();
    triageUndo();
  }
});

/* ==========================================
   第 2 步：自动处理（九宫格选图）
   ========================================== */
// 三阶段流水线：任意时候都能查看已交接的行。
function triageComplete() {
  return stateLoaded && triageQueue().length === 0;
}

function fixAvailable() {
  return stateLoaded;
}

function renderStepTabs() {
  const btn = document.getElementById('tabFixBtn');
  const left = stateLoaded ? pipelineMetrics().ready : 0;
  const locked = !fixAvailable();
  btn.classList.toggle('is-locked', locked);
  btn.textContent = `② 自动处理${left ? `（待确认 ${left} 行）` : ''}`;
  btn.title = '已筛完的行可立即处理，无需等待整批扫描或筛选完成';
}

function setStepTab(tab) {
  if (tab === 'fix' && !fixAvailable()) return;
  stepTab = tab === 'fix' ? 'fix' : 'triage';
  document.getElementById('tabTriageBtn').classList.toggle('active', stepTab === 'triage');
  document.getElementById('tabFixBtn').classList.toggle('active', stepTab === 'fix');
  document.getElementById('triageView').classList.toggle('hidden', stepTab !== 'triage');
  document.getElementById('fixView').classList.toggle('hidden', stepTab !== 'fix');
  if (stepTab === 'fix') pollAutofix();
}

const FIX_OPEN_STATUSES = new Set(['queued', 'running', 'ready', 'failed', 'error']);

// 还需要用户逐行处理的组：没写入、没点选；跳过（无需修改）的问题项已不在分组里。
// 全自动时排队、做图中的行不单独显示；处理不了的由后台直接转稍后人工复检
function needsUser(g) {
  if (!fixData.auto || g.manual || g.recheck_token) return true;
  return g.status === 'ready';
}

function fixQueue() {
  return fixData.groups.filter(g => FIX_OPEN_STATUSES.has(g.status) && needsUser(g) && !fixChosen.has(g.id)
    && !fixDeferring.has(g.id) && !g.item_ids.some(id => latestItems[id]?.later_review)
    && g.item_ids.some(id => latestItems[id] && !isResolved(latestItems[id]))
    && g.item_ids.every(id => latestItems[id] && (isResolved(latestItems[id]) || triageKept.has(id))));
}

function pickFixGroup() {
  const queue = fixQueue();
  if (queue.some(g => g.id === fixCurrentId)) return queue.find(g => g.id === fixCurrentId);
  const next = queue.find(g => !fixDeferred.has(g.id)) || queue[0];
  if (next && fixDeferred.size && queue.every(g => fixDeferred.has(g.id))) fixDeferred.clear();
  fixCurrentId = next ? next.id : null;
  return next || null;
}

function attemptLabel(group, index) {
  const attempt = group.attempts[index];
  if (index === 0) return `${group.recheck_token ? '本次复检' : '原样'} ${attempt.height}pt`;
  return `+${(attempt.height - group.attempts[0].height).toFixed(1).replace(/\.0$/, '')}pt → ${attempt.height}pt`;
}

// 候选图：按每张图自己的编号取（内容不变、可缓存）；没有图或多次取不到时显示占位文字，不显示裂图
function fixTileImage(group, index, lazy = false) {
  const attempt = group.attempts[index];
  if (!attempt.key) return '<span class="fix-tile-missing">暂无截图</span>';
  return `<img ${lazy ? 'loading="lazy" ' : ''}src="/api/autofix/image/${encodeURIComponent(attempt.key)}" alt="${escapeHtml(attemptLabel(group, index))}" onerror="fixImageFailed(this)">`;
}

function fixImageFailed(img) {
  const tries = Number(img.dataset.tries || 0);
  if (tries < 2) {
    // 偶尔一次没取到（程序正忙等）：过一会儿再试
    img.dataset.tries = tries + 1;
    setTimeout(() => {
      if (img.isConnected) img.src = `${img.src.split('?')[0]}?retry=${tries + 1}`;
    }, 1500 * (tries + 1));
    return;
  }
  const missing = document.createElement('span');
  missing.className = 'fix-tile-missing';
  missing.textContent = '截图暂不可用';
  img.replaceWith(missing);
}

function renderFix() {
  renderPipelineProgress();
  const groups = fixData.groups;
  const queue = fixQueue();
  const applied = groups.filter(g => g.status === 'applied').length;
  const writing = groups.filter(g => g.status === 'applying' || fixChosen.has(g.id) && g.status !== 'applied').length;
  document.getElementById('fixProgressText').textContent = autoProgressText(groups);
  renderConfirm();

  const stage = document.getElementById('fixStage');
  if (!fixAvailable()) {
    const signature = 'loading';
    if (stage.dataset.signature === signature) return;
    stage.dataset.signature = signature;
    stage.innerHTML = `<div class="triage-done">
      <span>正在读取已交接的处理队列…</span>
    </div>`;
    return;
  }
  const untriaged = triageQueue().length;
  const notice = untriaged ? `<div class="fix-note">当前还有 ${untriaged} 处可筛选，已交接的行可继续处理。
    <button type="button" class="btn btn-outline btn-sm" onclick="setStepTab('triage')">到 ① 筛选确认</button></div>` : '';
  const group = pickFixGroup();
  if (!group) {
    const metrics = pipelineMetrics();
    const syncing = metrics.handed > metrics.done;
    const waiting = scanRunning || untriaged > 0 || writing > 0 || pendingKeeps.size > 0 || !!keepRequest || syncing;
    const reviewFiles = Object.keys(reviewQueue).length;
    const working = fixData.auto ? groups.filter(g => ['queued', 'running'].includes(g.status) && !g.manual
      && g.item_ids.every(id => latestItems[id])).length : 0;
    const toConfirm = confirmGroups().length;
    const signature = `done|${applied}|${groups.length}|${untriaged}|${waiting}|${syncing}|${reviewFiles}|${working}|${toConfirm > 0}`;
    if (stage.dataset.signature === signature) return;
    stage.dataset.signature = signature;
    const reviewHint = reviewFiles ? `<span>还有 ${reviewFiles} 个文件挂在下方「稍后人工复检」，最后一次打开统一修改。</span>` : '';
    const confirmHint = toConfirm ? '<span>处理好的行在下方等你确认。</span>' : '';
    stage.innerHTML = `${notice}<div class="triage-done">${working
      ? `<b>后台正在全自动处理 ${working} 行…</b>${confirmHint}${reviewHint}`
      : waiting
      ? `<b>${writing ? '正在保存…' : syncing ? '已交接的行正在同步处理队列…' : '当前队列已处理完，等待扫描或筛选交接新行…'}</b>${confirmHint}${reviewHint}`
      : groups.length || reviewFiles ? `<b>${toConfirm ? '全自动处理完了，请在下方确认' : '当前队列已处理完'}</b>${toConfirm ? '' : `<span>已写入 ${applied} 行。</span>`}${reviewHint}`
      : '<span>没有需要处理的切字。先在第 ① 步筛选，保留的会出现在这里。</span>'}</div>`;
    return;
  }

  const items = group.item_ids.map(id => latestItems[id]).filter(Boolean);
  const skipping = fixSkipping.has(group.id);
  const blocked = skipping || processingBusy() || !!manualAction || !!reviewAction;
  const signature = JSON.stringify([
    group.id, group.status, group.attempts, group.recommended, group.manual, group.message,
    skipping, isBusy, manualAction, untriaged, group.recheck_token,
    items.map(item => [item.id, item.status, item.cell_addr]),
  ]);
  if (stage.dataset.signature === signature) return;
  stage.dataset.signature = signature;

  const choosable = ['ready', 'failed'].includes(group.status) && !group.manual && !blocked;
  const tiles = group.attempts.map((attempt, index) => {
    const canPick = choosable && index > 0 && attempt.ok;
    const classes = ['fix-tile', attempt.ok ? 'is-ok' : 'is-bad'];
    if (index === group.recommended) classes.push('is-recommended');
    if (canPick) classes.push('is-pickable');
    if (index === 0) classes.push('is-original');
    return `
      <button type="button" class="${classes.join(' ')}" ${canPick ? `onclick="chooseAttempt('${escapeHtml(group.id)}', ${index})"` : 'disabled'}>
        ${index === group.recommended ? '<span class="fix-tile-badge">推荐</span>' : ''}
        ${fixTileImage(group, index)}
        <span class="fix-tile-caption">${escapeHtml(attemptLabel(group, index))} · ${index === 0 ? (attempt.ok ? '完整且无压线' : '未通过检查') : (attempt.ok ? '✓ 完整且无压线' : '✗ 未通过完整显示/压线检查')}</span>
      </button>`;
  });
  if (!group.manual && ['queued', 'running'].includes(group.status) && tiles.length < 9) {
    const ahead = fixQueue().filter(g => g.status === 'queued' && g.id !== group.id && g.seq < group.seq).length;
    tiles.push(`<div class="fix-tile is-pending"><span>${group.status === 'running'
      ? `后台正在生成第 ${tiles.length + 1} 张...` : `排队中${ahead ? `，前面还有 ${ahead} 行` : ''}...`}</span></div>`);
  }
  let note = '';
  if (manualAction) note = manualAction.kind === 'manual_open' ? '正在打开 Excel 并定位，请稍候…' : '正在复检当前工作表，完成后自动更新结果…';
  else if (group.recheck_token) note = group.message;
  else if (group.manual) note = '已转人工：在 Excel 里改好并保存后，点「复检」；也可以挂到稍后统一人工复检。';
  else if (group.message) note = group.message;
  const first = items[0];
  stage.innerHTML = `
    ${notice}
    <div class="fix-card">
      <div class="triage-card-head">
        <div>
          <b>${groupExcelLink(group, items)}</b>
          <span style="color:#64748b;">/ ${escapeHtml(group.sheet_name)} · 第 ${Number(group.row)} 行${group.base_height ? `（当前行高 ${group.base_height}pt）` : ''}</span>
        </div>
        <span class="badge ${group.manual ? 'badge-warn' : 'badge-selected'}">${group.manual ? '人工处理中' : `本行 ${items.length} 处`}</span>
      </div>
      <div class="fix-items">${items.map(item => `<div>${escapeHtml(item.cell_addr)}：${escapeHtml(itemTextLabel(item))}</div>`).join('')}</div>
      <div class="fix-grid">${tiles.join('')}</div>
      ${note ? `<div class="fix-note">${escapeHtml(note)}</div>` : ''}
      <div class="triage-card-tools">
        <button type="button" class="btn btn-primary btn-sm" ${blocked ? 'disabled' : ''}
          title="将本行所有问题保存为人工确认无需修改，不修改 Excel"
          onclick="skipFixGroup('${escapeHtml(group.id)}')">${skipping ? '正在保存跳过...' : '跳过本行（无需修改）→'}</button>
        <button type="button" class="btn btn-outline btn-sm" ${blocked ? 'disabled' : ''}
          title="先记下来，不现在打开 Excel：程序把要改的单元格标成黄色、文件名前加 ★，最后一次打开全部文件统一修改"
          onclick="deferReview('${escapeHtml(group.id)}')"><span class="review-dot" aria-hidden="true"></span>稍后人工复检</button>
        ${group.manual && first ? `<button type="button" class="btn btn-primary btn-sm" ${blocked || isBusy ? 'disabled' : ''} title="${RECHECK_TITLE}" onclick="manualRecheck('${escapeHtml(first.id)}')">${manualAction?.kind === 'recheck' ? '正在复检…' : '复检'}</button>` : ''}
        <button type="button" class="btn btn-outline btn-sm" ${blocked ? 'disabled' : ''} title="暂时换到下一行，不保存跳过判断，稍后还会显示本行" onclick="deferFixGroup('${escapeHtml(group.id)}')">先看下一个 →</button>
      </div>
    </div>`;
}

/* ==========================================
   全自动处理结果：批量确认 · 改用别的图 · 不对就退回
   ========================================== */
function autoProgressText(groups) {
  const live = groups.filter(g => g.item_ids.every(id => latestItems[id]));
  const queued = live.filter(g => g.status === 'queued' && !g.manual).length;
  const running = live.find(g => g.status === 'running');
  const writing = live.filter(g => g.status === 'applying' || fixChosen.has(g.id) && g.status !== 'applied').length;
  const waiting = confirmGroups().filter(g => g.status === 'applied').length;
  const parts = [];
  if (running) parts.push(`正在处理 ${running.wb_name} / ${running.sheet_name} 第 ${running.row} 行（第 ${running.attempts.length + 1} 张）`);
  if (queued) parts.push(`排队 ${queued} 行`);
  if (writing) parts.push(`写入中 ${writing} 行`);
  parts.push(`已写入待确认 ${waiting} 行`);
  return `${fixData.auto ? '全自动处理：' : ''}${parts.join(' · ')}`;
}

// 全自动写进 Excel、还没确认的行（含正在改图、正在退回、刚退回的）
function confirmGroups() {
  const listed = fixData.groups.filter(g => !g.confirmed && (
    g.status === 'applied' || g.status === 'applying' || g.status === 'reverting'
    || (g.status === 'reverted' && !g.dismissed)));
  const key = confirmKey;
  for (const g of listed) if (!confirmOrder.includes(key(g))) confirmOrder.push(key(g));
  const rank = new Map(confirmOrder.map((k, i) => [k, i]));
  return listed.sort((a, b) => rank.get(key(a)) - rank.get(key(b)));
}

// 文件名前加 / 去 ★ 时分组 id 会变，排序用去掉 ★ 的路径
function confirmKey(g) {
  return `${g.wb_path.replace(REVIEW_PREFIX_PATTERN, '$1')}\0${g.sheet_name}\0${g.row}`;
}

// 分组标题里的文件名：点开 Excel 跳到这一行第一处问题
function groupExcelLink(group, items) {
  return excelLinkHtml(escapeHtml(group.wb_name), {
    file: group.wb_path, sheet: group.sheet_name, cell: items[0]?.cell_addr || `A${Number(group.row) || 1}`,
  });
}

function confirmCardHtml(group) {
  const items = group.item_ids.map(id => latestItems[id] || serverItems[id]).filter(Boolean);
  const reverted = group.status === 'reverted';
  const busy = confirmBusy.has(group.id) || confirmSubmitting || processingBusy() || !!manualAction || !!reviewAction;
  const writing = group.status === 'applying' || group.status === 'reverting';
  const canAct = group.status === 'applied' && !busy;
  const chosen = group.chosen;
  const tiles = group.attempts.map((attempt, index) => {
    const isChosen = index === chosen;
    const canPick = canAct && index > 0 && attempt.ok && !isChosen;
    const classes = ['fix-tile', attempt.ok ? 'is-ok' : 'is-bad'];
    if (isChosen && !reverted) classes.push('is-recommended');
    if (canPick) classes.push('is-pickable');
    if (index === 0) classes.push('is-original');
    const badge = isChosen && !reverted ? (group.status === 'applying' ? '写入中' : '已写入')
      : index === group.recommended ? '推荐' : '';
    return `
      <button type="button" class="${classes.join(' ')}" ${canPick ? `title="改用这张：按它的行高重新写入" onclick="switchAttempt('${escapeHtml(group.id)}', ${index})"` : 'disabled'}>
        ${badge ? `<span class="fix-tile-badge">${badge}</span>` : ''}
        ${fixTileImage(group, index, true)}
        <span class="fix-tile-caption">${escapeHtml(attemptLabel(group, index))} · ${index === 0 ? (attempt.ok ? '完整且无压线' : '未通过检查') : (attempt.ok ? '✓ 完整且无压线' : '✗ 未通过')}</span>
      </button>`;
  });
  const height = group.chosen != null && group.attempts[group.chosen] ? group.attempts[group.chosen].height
    : items[0]?.row_height;
  const state = reverted ? `已改回原行高 ${group.base_height}pt，转到下方「稍后人工复检」`
    : group.status === 'reverting' ? '正在改回原行高…'
    : group.status === 'applying' ? '正在写入 Excel…'
    : `行高 ${group.base_height}pt → ${height}pt，已保存`;
  return `
    <div class="triage-card-head">
      <div>
        <b>${groupExcelLink(group, items)}</b>
        <span style="color:#64748b;">/ ${escapeHtml(group.sheet_name)} · 第 ${Number(group.row)} 行</span>
      </div>
      <span class="badge ${reverted ? 'badge-warn' : 'badge-selected'}">${escapeHtml(state)}</span>
    </div>
    <div class="fix-items">${items.map(item => `<div>${escapeHtml(item.cell_addr)}：${escapeHtml(itemTextLabel(item))}</div>`).join('')}</div>
    ${reverted ? '' : `<div class="fix-grid fix-grid-compact">${tiles.join('')}</div>`}
    ${group.restored && !reverted ? `<div class="fix-note">${escapeHtml(group.message)}</div>` : ''}
    ${reverted ? '' : `<div class="triage-card-tools">
      <button type="button" class="btn btn-outline btn-sm" ${canAct && group.base_height != null ? '' : 'disabled'}
        title="改回原来的行高，并转到稍后人工复检（单元格标黄、文件名加 ★）"
        onclick="rejectConfirm('${escapeHtml(group.id)}')">${group.status === 'reverting' || confirmBusy.has(group.id) && !writing ? '正在提交…' : '✗ 不对，改回原样转人工'}</button>
    </div>`}`;
}

function renderConfirm() {
  const panel = document.getElementById('confirmPanel');
  if (!panel) return;
  const groups = confirmGroups();
  panel.classList.toggle('hidden', groups.length === 0);
  const confirmable = groups.filter(g => g.status === 'applied');
  const head = document.getElementById('confirmHead');
  const busy = confirmSubmitting || processingBusy() || !!manualAction || !!reviewAction;
  const headSig = JSON.stringify([confirmable.length, groups.length, busy, confirmSubmitting]);
  if (head.dataset.signature !== headSig) {
    head.dataset.signature = headSig;
    const reverted = groups.filter(g => g.status === 'reverted').length;
    head.innerHTML = `
      <div>
        <b>✅ 已自动处理，待你确认：${confirmable.length} 行</b>
        <span>每行已按带框的那张写进 Excel 并保存。要换就点另一张通过的图；不对就点「改回原样转人工」。${reverted ? `已退回 ${reverted} 行。` : ''}</span>
      </div>`;
    // 按钮放在列表底部的悬浮条里（sticky），往下翻卡片时一直在窗口底部，不用回到顶部找
    document.getElementById('confirmBar').innerHTML = `
      <span>待确认 ${confirmable.length} 行${reverted ? ` · 已退回 ${reverted} 行` : ''}</span>
      <button type="button" class="btn btn-primary btn-sm" ${busy || !(confirmable.length || reverted) ? 'disabled' : ''}
        onclick="confirmAll()">${confirmSubmitting ? '正在确认…' : `✓ 全部确认（${confirmable.length} 行）`}</button>`;
  }
  const list = document.getElementById('confirmList');
  const wanted = new Set(groups.map(confirmKey));
  for (const el of [...list.children]) if (!wanted.has(el.dataset.key)) el.remove();
  let previous = null;
  for (const group of groups) {
    const key = confirmKey(group);
    let el = [...list.children].find(c => c.dataset.key === key);
    if (!el) {
      el = document.createElement('div');
      el.dataset.key = key;
      if (previous) previous.after(el); else list.prepend(el);
    }
    const sig = JSON.stringify([group.id, group.status, group.chosen, group.attempts, group.base_height,
      confirmBusy.has(group.id), confirmSubmitting, processingBusy(), !!manualAction, !!reviewAction,
      group.item_ids.map(id => latestItems[id]?.cell_addr)]);
    if (el.dataset.signature !== sig) {
      el.dataset.signature = sig;
      el.className = `fix-card fix-confirm-card${group.status === 'reverted' ? ' is-reverted' : ''}`;
      el.innerHTML = confirmCardHtml(group);
    }
    previous = el;
  }
}

async function confirmAction(groupId, path, payload) {
  if (confirmBusy.has(groupId) || processingBusy() || manualAction || reviewAction) return;
  confirmBusy.add(groupId);
  renderConfirm();
  try {
    await postAction(path, payload);
  } catch (e) {
    alert(e.message);
  } finally {
    await pollAutofix();
    confirmBusy.delete(groupId);
    renderConfirm();
  }
}

function switchAttempt(groupId, index) {
  return confirmAction(groupId, '/api/autofix/switch', { id: groupId, index });
}

function rejectConfirm(groupId) {
  return confirmAction(groupId, '/api/autofix/reject', { id: groupId });
}

async function confirmAll() {
  const groups = confirmGroups().filter(g => g.status === 'applied' || g.status === 'reverted');
  if (!groups.length || confirmSubmitting || processingBusy() || manualAction || reviewAction) return;
  confirmSubmitting = true;
  renderConfirm();
  try {
    await postAction('/api/autofix/confirm', { ids: groups.map(g => g.id) });
    const done = new Set(groups.map(confirmKey));
    confirmOrder = confirmOrder.filter(k => !done.has(k));
  } catch (e) {
    alert(e.message);
  } finally {
    confirmSubmitting = false;
    await pollAutofix();
  }
}

async function pollAutofix() {
  try {
    // 筛选时也为已交接的行准备截图；当前人工编辑卡片保持暂停。
    const current = pickFixGroup();
    const active = stateLoaded && !processingBusy() && !(stepTab === 'fix' && current?.manual)
      && !manualAction && !document.hidden ? '?active=1' : '';
    const res = await fetch(`/api/autofix${active}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '读取自动处理进度失败');
    fixData = data;
    const ids = new Set(data.groups.map(g => g.id));
    fixChosen = new Set([...fixChosen].filter(id => ids.has(id) && data.groups.find(g => g.id === id).status !== 'queued'));
    fixDeferred = new Set([...fixDeferred].filter(id => ids.has(id)));
    renderFix();
    renderStepTabs();
    renderPipelineProgress();
  } catch (e) {
    document.getElementById('fixProgressText').textContent = `读取自动处理进度失败: ${e.message}`;
  }
}

async function chooseAttempt(groupId, index) {
  if (processingBusy() || manualAction) return;
  fixChosen.add(groupId);
  fixCurrentId = null;
  renderFix(); // 点完立刻换下一行，写入在后台进行
  try {
    await postAction('/api/autofix/choose', { id: groupId, index });
  } catch (e) {
    fixChosen.delete(groupId);
    alert(e.message);
  }
  pollAutofix();
}

async function deferReview(groupId) {
  const group = fixData.groups.find(g => g.id === groupId);
  if (!group || fixDeferring.has(groupId) || processingBusy() || manualAction || reviewAction) return;
  if (canResume && !resumeScan && !confirm('挂起会取消上次中断任务的续跑，避免旧检查点覆盖人工判断。确定吗？')) return;
  fixDeferring.add(groupId);
  fixCurrentId = null;
  renderFix(); // 立刻换下一行，标黄在后台进行
  try {
    await syncSkips();
    await postAction('/api/review/later', { id: groupId, item_ids: group.item_ids });
    await pollState();
  } catch (e) {
    alert(`挂起失败：${e.message}`);
  } finally {
    fixDeferring.delete(groupId);
    await pollAutofix();
  }
}

/* ==========================================
   稍后人工复检：一次打开全部 · 统一复检
   ========================================== */
// 文件名前缀：★，旧版本是 🟡（日文 Windows 10 的字体没有 🟡，会显示成方框）
const REVIEW_PREFIXES = ['★', '🟡'];
const REVIEW_PREFIX_PATTERN = /(^|[\\/])(?:★|🟡)/;

function hasReviewPrefix(name) {
  return REVIEW_PREFIXES.some(p => name.startsWith(p));
}

function stripReviewPrefix(name) {
  const prefix = REVIEW_PREFIXES.find(p => name.startsWith(p));
  return prefix ? name.slice(prefix.length) : name;
}

// 复检列表里每处显示「内容」和原因；内容只有「-」这类短符号时也看得出是单元格的文字
function reviewItemText(item) {
  const reason = isHidden(item) ? (item.hidden_reason || '文字没显示全') : '文字被边框线压到';
  return `「${item.text}」${reason}`;
}

function reviewFileLabel(path, entry) {
  const firstId = entry.marks.flatMap(m => m.item_ids).find(id => latestItems[id]);
  return firstId ? latestItems[firstId].wb_name : stripReviewPrefix(path.split(/[\\/]/).pop());
}

function firstOpenMark(entry) {
  return entry.marks.find(m => !markResolved(m)) || entry.marks[0];
}

function markResolved(mark) {
  return mark.item_ids.every(id => !serverItems[id] || isResolved(serverItems[id]));
}

function renderReview() {
  const panel = document.getElementById('reviewPanel');
  const files = Object.entries(reviewQueue).sort((a, b) => a[0].localeCompare(b[0]));
  panel.classList.toggle('hidden', files.length === 0);
  const busy = isBusy || !!reviewAction || !!manualAction;
  const signature = JSON.stringify([files, busy, scanRunning, reviewAction,
    files.flatMap(([, e]) => e.marks.flatMap(m => m.item_ids.map(id => serverItems[id]?.status)))]);
  if (panel.dataset.signature === signature) return;
  panel.dataset.signature = signature;
  if (!files.length) {
    panel.innerHTML = '';
    return;
  }
  const marks = files.flatMap(([, e]) => e.marks);
  const open = marks.filter(m => !markResolved(m)).length;
  const unmarked = files.filter(([path, e]) => !e.rename_error
    && (e.marks.some(m => m.state === 'pending') || !hasReviewPrefix(path.split(/[\\/]/).pop())));
  const failed = unmarked.filter(([, e]) => e.mark_failed).length;
  const pending = unmarked.length - failed;
  const status = [
    pending ? (isBusy ? `${pending} 个文件等当前任务结束后标黄…` : `后台正在给 ${pending} 个文件标黄、加 ★…`) : '',
    failed ? `${failed} 个文件标黄失败，点「一次打开全部」时会再试一次。` : '',
  ].filter(Boolean).join(' ') || '已全部标黄，文件名前加了 ★。';
  const rows = files.map(([path, entry]) => `
    <div class="review-file">
      <div class="review-file-head">
        <b><span class="review-dot" aria-hidden="true"></span>${excelLinkHtml(escapeHtml(reviewFileLabel(path, entry)), {
          file: path, sheet: firstOpenMark(entry)?.sheet, cell: firstOpenMark(entry)?.addr })}</b>
        <span>${entry.marks.length} 处</span>
      </div>
      ${entry.error ? `<div class="review-file-error">${escapeHtml(entry.error)}</div>` : ''}
      ${entry.rename_error ? `<div class="review-file-error">${escapeHtml(entry.rename_error)}</div>` : ''}
      ${entry.marks.map(mark => {
        const resolved = markResolved(mark);
        const text = mark.item_ids.map(id => serverItems[id]).filter(Boolean).map(reviewItemText).join(' / ');
        const state = resolved ? escapeHtml('✓ 已解决，复检时去掉黄色')
          : mark.state === 'marked' ? '<span class="review-swatch" aria-hidden="true"></span>已标黄'
          : escapeHtml(mark.state === 'skipped' ? `⚠️ 没标黄：${mark.error || ''}` : '⏳ 等待标黄');
        const preview = mark.item_ids.map(id => serverItems[id]?.image).find(Boolean) || '';
        const zoomCaption = `${mark.sheet}!${mark.addr} ${text}`;
        return `<div class="review-mark ${resolved ? 'is-resolved' : ''}" data-zoom="${escapeHtml(preview)}" data-zoom-caption="${escapeHtml(zoomCaption)}">
          <span class="review-mark-cell">${excelLinkHtml(`${escapeHtml(mark.sheet)}!${escapeHtml(mark.addr)}`, { file: path, sheet: mark.sheet, cell: mark.addr })}</span>
          <span class="review-mark-text" title="${escapeHtml(text)}">${escapeHtml(text)}</span>
          <span class="review-mark-state">${state}</span>
          ${resolved ? '' : `<button type="button" class="btn btn-outline btn-sm" ${busy ? 'disabled' : ''}
            title="确认这处其实不用改：记为跳过，统一复检时去掉黄色" onclick="reviewSkip('${escapeHtml(mark.item_ids.join(','))}')">无需修改</button>`}
        </div>`;
      }).join('')}
    </div>`).join('');
  const opening = reviewAction?.kind === 'review_open';
  const rechecking = reviewAction?.kind === 'review_recheck';
  panel.innerHTML = `
    <div class="review-head">
      <div>
        <b><span class="review-dot" aria-hidden="true"></span>稍后人工复检：${files.length} 个文件 · 待修改 ${open} 处</b>
        <span>${escapeHtml(status)}</span>
      </div>
      <div class="review-actions">
        <button type="button" class="btn btn-outline btn-sm" ${busy ? 'disabled' : ''} onclick="reviewOpenAll()"
          title="在 Excel 里一次打开全部文件，各自跳到第一处黄色单元格">${opening ? '正在打开 Excel…' : `📂 一次打开全部（${files.length} 个）`}</button>
        <button type="button" class="btn btn-primary btn-sm" ${busy ? 'disabled' : ''} onclick="reviewRecheck()"
          title="只重新检查这些文件；没问题的去掉黄色和 ★，还有问题的继续留在这里">${rechecking ? '正在统一复检…' : '🔁 统一复检'}</button>
      </div>
    </div>
    <div class="review-list">${rows}</div>
    <p class="triage-hint">改完记得在 Excel 里保存。统一复检时，已保存的文件会自动关闭再检查；还有没保存修改的文件会跳过。没问题的单元格会还原成原来的底色，整个文件都没问题了就去掉文件名前的 ★。</p>`;
}

async function submitReviewAction(kind, path) {
  if (isBusy || reviewAction || manualAction) return;
  if (canResume && !resumeScan && !confirm('此操作会替换上次中断任务的续跑检查点，确定吗？')) return;
  reviewAction = {kind, submitted: false};
  renderReview();
  renderFix();
  try {
    await syncSkips();
    await postAction(path, {});
    minAcceptedSeq = pollSeq + 1;
    reviewAction.submitted = true;
    await pollState();
  } catch (e) {
    reviewAction = null;
    renderReview();
    renderFix();
    alert(e.message);
  }
}

function reviewOpenAll() {
  submitReviewAction('review_open', '/api/review/open');
}

function reviewRecheck() {
  submitReviewAction('review_recheck', '/api/review/recheck');
}

async function reviewSkip(ids) {
  if (isBusy || reviewAction) return;
  try {
    await syncSkips();
    await postAction('/api/manual/skip', { decisions: Object.fromEntries(ids.split(',').map(id => [id, true])) });
    await pollState();
  } catch (e) {
    alert(e.message);
  }
}

function finishReviewAction(data) {
  if (!reviewAction?.submitted || data.busy || data.job?.kind !== reviewAction.kind
      || !['completed', 'interrupted'].includes(data.job.status)) return false;
  reviewAction = null;
  return true;
}

function deferFixGroup(groupId) {
  fixDeferred.add(groupId);
  fixCurrentId = null;
  renderFix();
}

async function skipFixGroup(groupId) {
  const group = fixData.groups.find(g => g.id === groupId);
  if (!group || !FIX_OPEN_STATUSES.has(group.status) || fixChosen.has(groupId) || fixSkipping.has(groupId)) return;
  if (processingBusy() || manualAction) {
    alert('任务执行中，请结束后再标记跳过。');
    return;
  }
  if (canResume && !resumeScan && !confirm('跳过本行会取消上次中断任务的续跑，避免旧检查点覆盖人工判断。确定吗？')) return;
  fixSkipping.add(groupId);
  renderFix();
  try {
    await syncSkips();
    await postAction('/api/autofix/skip', { id: groupId, item_ids: group.item_ids });
    fixData.groups = fixData.groups.filter(g => g.id !== groupId);
    fixDeferred.delete(groupId);
    fixCurrentId = null;
    await pollState();
    await pollAutofix();
  } catch (e) {
    alert(`跳过本行失败：${e.message}`);
  } finally {
    fixSkipping.delete(groupId);
    renderFix();
  }
}

/* ==========================================
   后端状态轮询与接口调用
   ========================================== */
async function pollState() {
  const seq = ++pollSeq;
  try {
    const res = await fetch('/api/state');
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '读取任务失败');
    if (seq < minAcceptedSeq || seq < lastAppliedPollSeq) return;
    lastAppliedPollSeq = seq;

    canResume = data.can_resume;
    resumeScan = !!(canResume && data.job?.kind === 'scan');
    isBusy = data.busy;
    scanRunning = !!(data.busy && data.job?.kind === 'scan');
    scanProgress = data.scan_progress || {};
    if (!canResume) skipCancelsResumeConfirmed = false;
    document.getElementById('resumeBtn').classList.toggle('hidden', !canResume);
    document.getElementById('resumeBtn').disabled = data.busy;
    if (!workspaceRestored) {
      // 进行中或可续跑的任务以服务器为准，否则优先用上次填过的路径
      const taskFolder = (data.busy || canResume) ? data.folder : '';
      document.getElementById('folderInput').value = taskFolder || readLastFolder() || data.folder || '';
    }

    document.getElementById('scanBtn').disabled = data.busy;
    document.getElementById('statusText').innerText = data.busy
      ? '⏳ 后台处理中...' : (canResume ? '任务已中断，可继续上次任务' : '✅ 就绪');

    const logBox = document.getElementById('logBox');
    logBox.innerText = (data.logs && data.logs.length > 0) ? data.logs.join('\n') : '等待操作...';
    logBox.scrollTop = logBox.scrollHeight;

    if (data.progress) {
      const p = data.progress;
      const pct = Number(p.progress_pct || 0);
      const navBadge = document.getElementById('navProgressBadge');
      if (navBadge) {
        navBadge.textContent = p.total_files > 0 ? `${p.completed_files}/${p.total_files} (${pct}%)` : '0%';
      }
      const miniStep = document.getElementById('miniStepText');
      if (miniStep) {
        miniStep.textContent = scanRunning ? '扫描进行中：已完成文件可以先筛选，已筛完的行在后台准备截图。' : (p.current_step || '等待开始任务...');
      }
      const miniStats = document.getElementById('miniStatsText');
      if (miniStats) {
        miniStats.textContent =
          `已完成 ${p.completed_files} · 处理中 ${p.processing_files} · 待处理 ${p.pending_files}（总进度 ${pct}%）`;
      }
      const miniFill = document.getElementById('miniProgressFill');
      if (miniFill) {
        miniFill.style.width = `${Math.min(100, Math.max(0, pct))}%`;
      }
    }

    latestPages = data.pages || {};
    serverItems = data.items || {};
    stateLoaded = true;
    for (const id of [...legacyKeeps]) {
      if (!serverItems[id]) continue;
      if (canResume && !resumeScan) continue;
      if (!isResolved(serverItems[id]) && !serverItems[id].triage_keep && !processingBusy()) pendingKeeps.set(id, true);
      legacyKeeps.delete(id);
    }
    // 已被重新检测替换掉的问题项，不再保留未提交的决定
    for (const id of [...pendingSkips.keys()]) {
      if (!serverItems[id]) pendingSkips.delete(id);
    }
    applyPendingSkips();
    flushKeeps();
    const manualCompleted = finishManualAction(data);
    reviewQueue = data.review_queue || {};
    const reviewCompleted = finishReviewAction(data);

    const hasResults = Object.keys(latestPages).length > 0 || Object.keys(latestItems).length > 0;
    const emptyStage = document.getElementById('emptyStagePlaceholder');
    if (emptyStage) {
      emptyStage.classList.toggle('hidden', hasResults);
    }

    workspaceRestored = true;
    renderItems();
    renderReview();
    if (manualCompleted || reviewCompleted) pollAutofix();
  } catch (e) {
    document.getElementById('statusText').innerText = `连接或读取失败: ${e.message}`;
  }
}

async function postAction(path, payload) {
  const res = await fetch(path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const data = await res.json();
  if (!res.ok || !data.ok) throw new Error(data.error || '操作未成功，任务可能正在执行');
  return data;
}

async function updateRow(id, newRow) {
  const parsed = parseInt(newRow, 10);
  if (!Number.isFinite(parsed) || parsed < 1) return;
  try {
    await postAction('/api/update_row', { id, excel_row: parsed });
  } catch (e) { alert(e.message); }
}

async function startScan() {
  const folder = document.getElementById('folderInput').value.trim();
  if (!folder) {
    alert('请先输入 Excel 文件夹路径！');
    return;
  }
  if (canResume && !confirm('开始扫描会替换上次中断任务的续跑进度，确定不继续上次任务吗？')) return;
  await syncSkips();
  try {
    await postAction('/api/scan', { folder, delete_backups: false, full: false });
    saveLastFolder(folder);
  } catch (e) {
    alert(e.message);
    return;
  }
  await pollState();
}

async function resumeTask() {
  await syncSkips();
  try {
    await postAction('/api/resume', {});
    await pollState();
  } catch (e) { alert(e.message); }
}

async function manualOpen(id) {
  if (isBusy || manualAction) return;
  if (canResume && !confirm('转人工将替换中断任务，旧检查点不再自动续跑。剩余项可之后重新处理。确定吗？')) return;
  await submitManualAction(id, 'manual_open', '/api/manual/open');
}

async function manualRecheck(id) {
  if (isBusy || manualAction) return;
  // 复检只读 Excel、不加高，平时不再弹窗；只有会替换中断任务续跑点（不可撤销）时才提示
  const message = '此操作会替换上次中断任务的续跑检查点，确定吗？';
  if (canResume && !confirm(message)) return;
  await submitManualAction(id, 'recheck', '/api/manual/recheck');
}

async function submitManualAction(id, kind, path) {
  manualAction = {id, kind, submitted: false};
  renderFix();
  try {
    await syncSkips();
    await postAction(path, { id });
    minAcceptedSeq = pollSeq + 1;
    manualAction.submitted = true;
    await pollState();
  } catch (e) {
    manualAction = null;
    renderFix();
    alert(e.message);
  }
}

function finishManualAction(data) {
  if (!manualAction?.submitted || data.busy || data.job?.kind !== manualAction.kind
      || !data.job.options?.selected_ids?.includes(manualAction.id)
      || !['completed', 'interrupted'].includes(data.job.status)) return false;
  if (data.job.status === 'completed' && manualAction.kind === 'recheck'
      && latestItems[manualAction.id] && isResolved(latestItems[manualAction.id])) {
    fixCurrentId = null;
  }
  manualAction = null;
  return true;
}

// 离开页面前若还有未保存的跳过决定，提示一下
window.addEventListener('beforeunload', e => {
  if (pendingSkips.size || skipRequest || pendingKeeps.size || keepRequest) {
    e.preventDefault();
    e.returnValue = '';
  }
});

triageKept = loadTriageKept();
legacyKeeps = new Set(triageKept);

// 定时轮询：上一次还没回来就不再发新的，免得请求越积越多（操作后主动刷新不受影响）
function whenIdle(poll) {
  let inFlight = false;
  return async () => {
    if (inFlight) return;
    inFlight = true;
    try {
      await poll();
    } finally {
      inFlight = false;
    }
  };
}

setInterval(whenIdle(pollState), 1200);
setInterval(whenIdle(pollAutofix), 1500);
pollAutofix();
pollState();

// 悬停放大（② 的候选图、稍后人工复检的每一处）：固定定位的浮层挂在 body 上，跟着鼠标走并限制在窗口内，不改变页面布局
const fixZoom = (() => {
  const box = document.createElement('div');
  box.className = 'fix-zoom';
  box.innerHTML = '<img alt=""><span class="fix-zoom-empty">暂无截图预览</span><span class="fix-zoom-caption"></span>';
  document.body.appendChild(box);
  const img = box.querySelector('img');
  const empty = box.querySelector('.fix-zoom-empty');
  const caption = box.querySelector('.fix-zoom-caption');
  const GAP = 18;
  const MARGIN = 12;
  let current = null;
  let lastX = 0;
  let lastY = 0;

  function place() {
    if (!current || (!img.hidden && !img.naturalWidth)) return;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    if (!img.hidden) {
      // 候选图按 4 倍渲染（扫描时的特写是 2.5 倍），最多显示原图的 3/4，且不超出窗口
      const maxW = Math.min(img.naturalWidth * 0.75, vw * 0.7);
      const maxH = vh - MARGIN * 2 - 40;
      const scale = Math.min(maxW / img.naturalWidth, maxH / img.naturalHeight, 1);
      img.style.width = `${Math.round(img.naturalWidth * scale)}px`;
      img.style.height = `${Math.round(img.naturalHeight * scale)}px`;
    }
    const w = box.offsetWidth;
    const h = box.offsetHeight;
    let left = lastX + GAP;
    if (left + w > vw - MARGIN) left = lastX - GAP - w;
    left = Math.max(MARGIN, Math.min(left, vw - MARGIN - w));
    let top = lastY - h / 2;
    top = Math.max(MARGIN, Math.min(top, vh - MARGIN - h));
    box.style.left = `${Math.round(left)}px`;
    box.style.top = `${Math.round(top)}px`;
  }

  function hide() {
    current = null;
    box.classList.remove('is-visible');
  }

  img.addEventListener('load', place);

  document.addEventListener('mousemove', event => {
    lastX = event.clientX;
    lastY = event.clientY;
    const target = event.target.closest && event.target.closest('.fix-tile, [data-zoom]');
    let src = '';
    let text = '';
    if (target && target.classList.contains('fix-tile')) {
      src = target.querySelector('img')?.src || '';
      text = target.querySelector('.fix-tile-caption')?.textContent || '';
      if (!src) {
        hide();
        return;
      }
    } else if (target) {
      src = target.dataset.zoom ? new URL(target.dataset.zoom, location.href).href : '';
      text = target.dataset.zoomCaption || '';
    } else {
      hide();
      return;
    }
    const key = `${src}|${text}`;
    if (key !== current) {
      current = key;
      caption.textContent = text;
      img.hidden = !src;
      empty.hidden = !!src;
      if (src && img.src !== src) img.src = src;
      box.classList.add('is-visible');
    }
    place();
  });
  document.addEventListener('mouseleave', hide);
  window.addEventListener('scroll', hide, true);
  window.addEventListener('blur', hide);
  return { hide };
})();
