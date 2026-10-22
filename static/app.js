let latestPages = {};
let serverItems = {}; // 服务器返回的问题项
let latestItems = {}; // 叠加尚未保存的跳过决定后的问题项，页面以此渲染
let workspaceRestored = false;
let canResume = false;
let isBusy = false;
let pollSeq = 0;
let minAcceptedSeq = 0; // 发起操作前已在途的轮询结果可能过期，直接丢弃

// 点击即跳过：先在页面上生效，再把这期间的多次点选合并成一次提交（每次写库较慢）
let pendingSkips = new Map(); // item id -> 是否跳过（尚未提交）
let inFlightSkips = new Map(); // 正在提交的决定
let skipRequest = null;
let skipCancelsResumeConfirmed = false;

// 逐个筛选：「保留」只表示已看过、需人工修改（问题项状态不变），记在本浏览器里用于定位下一处
let triageKept = new Set();
let triageHistory = []; // [{ id, action: 'keep' | 'skip' }]，用于撤销
let triageCurrentId = null;
let triageDrag = null;
let triageAnimationTimer = null;
const TRIAGE_SWIPE_THRESHOLD = 110;

const TRIAGE_KEPT_STORAGE_KEY = 'clear_handle_excel.triageKept';
const RECHECK_TITLE = '请先在 Excel 中保存手动修改。根据新 PDF 重新检测该文件的问题工作表并替换旧标记，不会自动增加行高';

function isResolved(item) {
  return item.status === 'fixed' || item.status === 'accepted';
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
}

// 返回是否已生效（任务执行中或用户取消时返回 false）
function setSkip(id, skip) {
  const item = latestItems[id];
  if (!item || item.status === 'fixed') return false;
  if (isBusy) {
    alert('任务执行中，请结束后再标记跳过。');
    return false;
  }
  if (canResume && !skipCancelsResumeConfirmed) {
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
  const reviewable = Object.values(latestItems).filter(item => item.status !== 'fixed');
  const skipped = reviewable.filter(item => item.status === 'accepted').length;
  const kept = reviewable.filter(item => !isResolved(item) && triageKept.has(item.id)).length;
  const queue = triageQueue();
  const decided = skipped + kept;
  document.getElementById('triageProgressText').textContent =
    `已筛选 ${decided} / ${reviewable.length}（保留 ${kept} · 跳过 ${skipped} · 剩余 ${queue.length}）`;
  document.getElementById('triageProgressFill').style.width =
    `${reviewable.length ? Math.round((decided / reviewable.length) * 100) : 0}%`;
  document.getElementById('triageUndoBtn').disabled = triageHistory.length === 0;

  // 拖动或飞出动画进行中不换卡
  if (triageDrag || triageAnimationTimer) return;
  const stage = document.getElementById('triageStage');
  const next = queue[0];
  if (!next) {
    triageCurrentId = null;
    const signature = `done|${kept}|${skipped}|${reviewable.length}`;
    if (stage.dataset.signature === signature) return;
    stage.dataset.signature = signature;
    stage.innerHTML = reviewable.length ? `
      <div class="triage-done">
        <b>🎉 已全部筛选完</b>
        <span>保留 ${kept} 处、跳过 ${skipped} 处。保留的请在 Excel 中修改并保存，再点「扫描切字」复检。</span>
        ${kept ? '<button type="button" class="btn btn-outline btn-sm" onclick="triageRestartKept()">重新筛选已保留的项</button>' : ''}
      </div>` : '<div class="triage-done"><span>暂无切字问题，请先扫描。</span></div>';
    return;
  }

  const page = latestPages[next.page_id];
  const signature = `${next.id}|${next.status}|${next.image}|${page ? page.image.length : 0}`;
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
          <b>#${Number(next.seq || 1)} ${escapeHtml(next.wb_name)}</b>
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
          <input type="number" min="1" value="${Number(next.excel_row) || 1}" onchange="updateRow('${escapeHtml(next.id)}', this.value)">
        </label>
        <button type="button" class="btn btn-outline btn-sm" onclick="manualOpen('${escapeHtml(next.id)}')">在 Excel 中定位</button>
        <button type="button" class="btn btn-outline btn-sm" title="${RECHECK_TITLE}" onclick="manualRecheck('${escapeHtml(next.id)}')">重新导出复检</button>
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
    triageKept.add(item.id);
    saveTriageKept();
  }
  triageHistory.push({ id: item.id, action });
  if (card) {
    card.style.transform = '';
    card.style.setProperty(action === 'skip' ? '--skip-opacity' : '--keep-opacity', 1);
    card.classList.add(action === 'skip' ? 'fly-right' : 'fly-left');
  }
  renderTriage();
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
    triageKept.delete(last.id);
    saveTriageKept();
  }
  renderTriage();
}

function triageRestartKept() {
  triageKept.clear();
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
  if (e.repeat || e.altKey) return;
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
   后端状态轮询与接口调用
   ========================================== */
async function pollState() {
  const seq = ++pollSeq;
  try {
    const res = await fetch('/api/state');
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '读取任务失败');
    if (seq < minAcceptedSeq) return;

    canResume = data.can_resume;
    isBusy = data.busy;
    if (!canResume) skipCancelsResumeConfirmed = false;
    document.getElementById('resumeBtn').classList.toggle('hidden', !canResume);
    document.getElementById('resumeBtn').disabled = data.busy;
    if (!workspaceRestored) {
      document.getElementById('folderInput').value = data.folder || '';
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
        miniStep.textContent = p.current_step || '等待开始任务...';
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
    // 已被重新检测替换掉的问题项，不再保留未提交的决定
    for (const id of [...pendingSkips.keys()]) {
      if (!serverItems[id]) pendingSkips.delete(id);
    }
    const keptBefore = triageKept.size;
    triageKept = new Set([...triageKept].filter(id => serverItems[id]));
    if (triageKept.size !== keptBefore) saveTriageKept();
    applyPendingSkips();

    const hasResults = Object.keys(latestPages).length > 0 || Object.keys(latestItems).length > 0;
    const emptyStage = document.getElementById('emptyStagePlaceholder');
    if (emptyStage) {
      emptyStage.classList.toggle('hidden', hasResults);
    }

    workspaceRestored = true;
    renderItems();
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
  if (canResume && !confirm('转人工将替换中断任务，旧检查点不再自动续跑。剩余项可之后重新处理。确定吗？')) return;
  await syncSkips();
  try {
    await postAction('/api/manual/open', { id });
    await pollState();
  } catch (e) { alert(e.message); }
}

async function manualRecheck(id) {
  // 复检只读 Excel、不加高，平时不再弹窗；只有会替换中断任务续跑点（不可撤销）时才提示
  const message = '此操作会替换上次中断任务的续跑检查点，确定吗？';
  if (canResume && !confirm(message)) return;
  await syncSkips();
  try {
    await postAction('/api/manual/recheck', { id });
    await pollState();
  } catch (e) { alert(e.message); }
}

// 离开页面前若还有未保存的跳过决定，提示一下
window.addEventListener('beforeunload', e => {
  if (pendingSkips.size || skipRequest) {
    e.preventDefault();
    e.returnValue = '';
  }
});

triageKept = loadTriageKept();

setInterval(pollState, 1200);
pollState();
