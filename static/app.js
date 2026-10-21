let currentViewMode = 'canvas'; // 'canvas' | 'cards'
let latestPages = {};
let serverItems = {}; // 服务器返回的问题项
let latestItems = {}; // 叠加尚未保存的跳过决定后的问题项，页面以此渲染
let pagesSignatureCache = {};
let cardsSignatureCache = {};
let activeHoverItemId = null;
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

const ONLY_PENDING_STORAGE_KEY = 'clear_handle_excel.onlyPending';
const RECHECK_TITLE = '请先在 Excel 中保存手动修改。根据新 PDF 重新检测该文件的问题工作表并替换旧标记，不会自动增加行高';

function isResolved(item) {
  return item.status === 'fixed' || item.status === 'accepted';
}

function skipButton(item) {
  if (item.status === 'fixed') return '';
  return `<button class="btn btn-outline btn-sm skip-btn" data-item-id="${escapeHtml(item.id)}" onclick="toggleSkip('${escapeHtml(item.id)}')"></button>`;
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

function setViewMode(mode) {
  currentViewMode = mode === 'cards' ? 'cards' : 'canvas';
  document.getElementById('viewCanvasBtn').classList.toggle('active', currentViewMode === 'canvas');
  document.getElementById('viewCardsBtn').classList.toggle('active', currentViewMode === 'cards');
  document.getElementById('canvasView').classList.toggle('hidden', currentViewMode !== 'canvas');
  document.getElementById('cardGrid').classList.toggle('hidden', currentViewMode !== 'cards');
  hideMagnifier();
}

function setOnlyPending(onlyPending) {
  document.body.classList.toggle('only-pending', onlyPending);
  document.getElementById('onlyPendingChk').checked = onlyPending;
  try {
    localStorage.setItem(ONLY_PENDING_STORAGE_KEY, onlyPending ? '1' : '0');
  } catch (e) {}
  hideMagnifier();
  refreshItemClasses();
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

function toggleSkip(id) {
  const item = latestItems[id];
  if (!item || item.status === 'fixed') return;
  if (isBusy) {
    alert('任务执行中，请结束后再标记跳过。');
    return;
  }
  if (canResume && !skipCancelsResumeConfirmed) {
    // 跳过本身可随时撤销，只有会取消中断任务续跑点（不可撤销）时才提示
    if (!confirm('标记跳过会取消上次中断任务的续跑，避免旧检查点覆盖人工判断。确定吗？')) return;
    skipCancelsResumeConfirmed = true;
  }
  const skip = item.status !== 'accepted';
  const baseline = inFlightSkips.has(id) ? inFlightSkips.get(id) : serverItems[id]?.status === 'accepted';
  if (skip === baseline) {
    pendingSkips.delete(id);
  } else {
    pendingSkips.set(id, skip);
  }
  applyPendingSkips();
  renderItems();
  flushSkips();
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

// 发起后台任务前先保存所有跳过决定，确保自动微调、复检和重新扫描看到的是最新筛选结果
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
      badgeText: '已跳过（无需修复）', shortTag: `跳过#${tagNo}`,
    };
  }
  if (item.status === 'fixed') {
    return {
      boxClass: 'is-fixed is-resolved',
      cardClass: 'fixed is-resolved',
      badgeClass: 'badge-fixed',
      badgeText: `✓ 已修复 (+${item.iterations}次)`,
      shortTag: `✓#${tagNo}`,
    };
  }
  if (item.status === 'needs_review' || item.status === 'manual_review') {
    return {
      boxClass: 'is-max-reached',
      cardClass: 'max-reached',
      badgeClass: 'badge-warn',
      badgeText: item.status === 'manual_review' ? '等待人工调整并保存' : '需人工复检',
      shortTag: `?#${tagNo}`,
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
    badgeClass: 'badge-selected',
    badgeText: '待修复（点击可跳过）',
    shortTag: `#${tagNo}`,
  };
}

function countByState(items) {
  const counts = { pending: 0, review: 0, skipped: 0, fixed: 0 };
  items.forEach(item => {
    if (item.status === 'pending') counts.pending += 1;
    else if (item.status === 'accepted') counts.skipped += 1;
    else if (item.status === 'fixed') counts.fixed += 1;
    else counts.review += 1;
  });
  return counts;
}

function refreshItemClasses() {
  Object.values(latestItems).forEach(item => {
    const st = getItemVisualState(item);

    const hotspot = document.getElementById('hotspot_' + item.id);
    if (hotspot) {
      hotspot.className = `hotspot-box ${st.boxClass}`;
      const tag = hotspot.querySelector('.hotspot-tag');
      if (tag) tag.textContent = `${st.shortTag} ${item.cell_addr}`;
    }

    const capsule = document.getElementById('capsule_' + item.id);
    if (capsule) {
      capsule.className = `item-capsule ${st.boxClass}`;
    }

    const card = document.getElementById('card_' + item.id);
    if (card) {
      card.className = `card ${st.cardClass}`;
      const badge = card.querySelector('.badge');
      if (badge) {
        badge.className = `badge ${st.badgeClass}`;
        badge.textContent = st.badgeText;
      }
    }

    const skipped = item.status === 'accepted';
    document.querySelectorAll(`.skip-btn[data-item-id="${CSS.escape(item.id)}"]`).forEach(btn => {
      btn.textContent = skipped ? '撤销跳过' : '跳过（无需修复）';
      btn.title = skipped
        ? '撤销跳过，重新纳入处理'
        : '跳过此处（人工确认无需修复）；保存文件或重新检测时，同一处文字仍保持跳过';
    });
  });

  const itemsByPage = {};
  Object.values(latestItems).forEach(item => {
    (itemsByPage[item.page_id] = itemsByPage[item.page_id] || []).push(item);
  });
  Object.keys(latestPages).forEach(pageId => {
    const counts = countByState(itemsByPage[pageId] || []);
    const stage = document.getElementById('stage_' + pageId);
    if (stage) stage.classList.toggle('page-resolved', counts.pending + counts.review === 0);
    const label = document.getElementById('stagecounts_' + pageId);
    if (label) {
      label.innerHTML = `（待修复 <b>${counts.pending}</b> · 需人工 <b>${counts.review}</b> · `
        + `已跳过 ${counts.skipped} · 已修复 <b style="color:#059669;">${counts.fixed}</b>）`;
    }
  });

  const totals = countByState(Object.values(latestItems));
  document.getElementById('cntPendingFix').innerText = totals.pending;
  document.getElementById('cntReview').innerText = totals.review;
  document.getElementById('cntSkipped').innerText = totals.skipped;
  document.getElementById('cntFixed').innerText = totals.fixed;
  const hasItems = Object.keys(latestItems).length > 0;
  document.getElementById('allDoneHint').classList.toggle(
    'hidden', !hasItems || totals.pending + totals.review > 0,
  );
}

function renderItems() {
  renderCanvasStages(latestPages, latestItems);
  renderCards(latestItems);
  refreshItemClasses();
  if (activeHoverItemId && latestItems[activeHoverItemId]) {
    updateMagnifierContent(latestItems[activeHoverItemId]);
  }
}

/* ==========================================
   同行高亮联动 & 2.5x 悬停放大镜气泡
   ========================================== */
function setRowPeerHighlight(item, active) {
  getPeerItemsInSameRow(item).forEach(peer => {
    const h = document.getElementById('hotspot_' + peer.id);
    if (h) {
      h.classList.toggle('is-peer-row', active && peer.id !== item.id);
      h.classList.toggle('is-active-hover', active && peer.id === item.id);
    }
    const c = document.getElementById('capsule_' + peer.id);
    if (c) {
      c.classList.toggle('is-peer-row', active);
    }
    const card = document.getElementById('card_' + peer.id);
    if (card) {
      card.classList.toggle('is-peer-row', active && peer.id !== item.id);
    }
  });
}

function updateMagnifierContent(item) {
  const st = getItemVisualState(item);
  const peers = getPeerItemsInSameRow(item);
  const peerTip = peers.length > 1 ? ` · 同行共 ${peers.length} 处切字` : '';

  document.getElementById('magTitle').textContent =
    `#${item.index_on_page || item.seq} 定位: ${item.cell_addr} (Excel 第 ${item.excel_row} 行 · 当前行高 ${item.row_height}pt${peerTip})`;
  const badgeEl = document.getElementById('magBadge');
  badgeEl.className = `badge ${st.badgeClass}`;
  badgeEl.textContent = st.badgeText;

  document.getElementById('magImg').src = item.image;
  document.getElementById('magText').textContent = `文字内容: ${item.text}`;
  let hint = '💡 点击框即可跳过此处（无需修复），再点一次撤销；可在画布下方胶囊修正行号';
  if (item.status === 'accepted') hint = '已跳过（人工确认无需修复），再点一次可撤销';
  else if (item.status === 'fixed') hint = '✅ 该位置已完整显示，无需继续调整';
  else if (item.status !== 'pending') hint = '需人工处理：可在 Excel 中定位或重新导出复检；确认无需修复时点击框跳过';
  document.getElementById('magHint').textContent = hint;
}

function positionMagnifierNearRect(anchorRect) {
  const pop = document.getElementById('magnifierPopover');
  const popW = 440;
  const popH = pop.offsetHeight || 280;
  const margin = 14;

  let left = anchorRect.right + margin;
  if (left + popW > window.innerWidth - 12) {
    left = anchorRect.left - popW - margin;
  }
  if (left < 12) {
    left = Math.max(12, Math.min(window.innerWidth - popW - 12, anchorRect.left));
  }

  let top = anchorRect.top - 10;
  if (top + popH > window.innerHeight - 12) {
    top = window.innerHeight - popH - 12;
  }
  if (top < 12) {
    top = 12;
  }

  pop.style.left = `${Math.round(left)}px`;
  pop.style.top = `${Math.round(top)}px`;
}

function showMagnifierForElement(item, anchorEl) {
  activeHoverItemId = item.id;
  setRowPeerHighlight(item, true);
  updateMagnifierContent(item);

  const pop = document.getElementById('magnifierPopover');
  pop.classList.remove('hidden');
  positionMagnifierNearRect(anchorEl.getBoundingClientRect());
}

function hideMagnifier(item) {
  activeHoverItemId = null;
  if (item) {
    setRowPeerHighlight(item, false);
  }
  const pop = document.getElementById('magnifierPopover');
  if (pop) {
    pop.classList.add('hidden');
  }
}

/* ==========================================
   视图一渲染：整页画布 + 热区 + 底部胶囊条
   ========================================== */
function renderCanvasStages(pages, items) {
  const container = document.getElementById('canvasView');
  const pageIds = Object.keys(pages);
  for (const pageId of Object.keys(pagesSignatureCache)) {
    if (!pages[pageId]) {
      document.getElementById('stage_' + pageId)?.remove();
      delete pagesSignatureCache[pageId];
    }
  }

  if (pageIds.length === 0) {
    container.innerHTML = '';
    pagesSignatureCache = {};
    return;
  }

  // 若用户在某个胶囊输入框内正在输入，跳过本次 DOM 重建以保留输入焦点
  const activeEl = document.activeElement;
  if (activeEl && activeEl.tagName === 'INPUT' && container.contains(activeEl)) {
    return;
  }

  const itemsByPage = {};
  Object.values(items).forEach(item => {
    if (!itemsByPage[item.page_id]) {
      itemsByPage[item.page_id] = [];
    }
    itemsByPage[item.page_id].push(item);
  });

  pageIds.forEach(pageId => {
    const page = pages[pageId];
    const pageItems = itemsByPage[pageId] || [];

    // 状态变化（跳过 / 撤销）只更新样式，不重建整页，避免点选时整页图片闪烁
    const sigParts = [
      page.image,
      ...pageItems.map(
        it => `${it.id}:${it.iterations}:${it.row_height}:${it.excel_row}:${it.cell_addr}:${it.bbox.join(',')}:${it.image}`
      ),
    ];
    const signature = sigParts.join('|');

    let stageCard = document.getElementById('stage_' + pageId);
    if (stageCard && pagesSignatureCache[pageId] === signature) {
      return;
    }
    pagesSignatureCache[pageId] = signature;

    if (!stageCard) {
      stageCard = document.createElement('div');
      stageCard.id = 'stage_' + pageId;
      stageCard.className = 'page-stage-card';
      container.appendChild(stageCard);
    }

    const pw = Math.max(1, Number(page.page_width) || 595);
    const ph = Math.max(1, Number(page.page_height) || 842);

    const hotspotsHtml = pageItems
      .map(item => {
        const [x0, y0, x1, y1] = item.bbox;
        const padX = 4;
        const padY = 3;
        const leftPct = Math.max(0, ((x0 - padX) / pw) * 100).toFixed(3);
        const topPct = Math.max(0, ((y0 - padY) / ph) * 100).toFixed(3);
        const widthPct = Math.min(100, ((x1 - x0 + padX * 2) / pw) * 100).toFixed(3);
        const heightPct = Math.min(100, ((y1 - y0 + padY * 2) / ph) * 100).toFixed(3);
        const st = getItemVisualState(item);

        return `
          <div
            id="hotspot_${escapeHtml(item.id)}"
            class="hotspot-box ${st.boxClass}"
            style="left:${leftPct}%; top:${topPct}%; width:${widthPct}%; height:${heightPct}%;"
            data-item-id="${escapeHtml(item.id)}"
          >
            <span class="hotspot-tag">${escapeHtml(st.shortTag)} ${escapeHtml(item.cell_addr)}</span>
          </div>
        `;
      })
      .join('');

    const capsulesHtml = pageItems
      .map(item => {
        const st = getItemVisualState(item);
        return `
          <div
            id="capsule_${escapeHtml(item.id)}"
            class="item-capsule ${st.boxClass}"
            data-item-id="${escapeHtml(item.id)}"
          >
            <span class="capsule-seq">#${Number(item.index_on_page || item.seq)}</span>
            <span class="capsule-addr">${escapeHtml(item.cell_addr)} (${escapeHtml(item.row_height)}pt)</span>
            <span class="capsule-text" title="${escapeHtml(item.text)}">${escapeHtml(item.text)}</span>
            <label onclick="event.stopPropagation();" style="display:inline-flex; align-items:center; gap:3px;">
              行:
              <input
                type="number"
                min="1"
                class="capsule-row-input"
                value="${Number(item.excel_row) || 1}"
                onchange="updateRow('${escapeHtml(item.id)}', this.value)"
              >
            </label>
            <span onclick="event.stopPropagation();">
              <button class="btn btn-outline btn-sm" onclick="manualOpen('${escapeHtml(item.id)}')">在 Excel 中定位</button>
              <button class="btn btn-outline btn-sm" title="${RECHECK_TITLE}" onclick="manualRecheck('${escapeHtml(item.id)}')">重新导出复检</button>
              ${skipButton(item)}
            </span>
          </div>
        `;
      })
      .join('');

    stageCard.innerHTML = `
      <div class="page-stage-header">
        <div class="page-stage-title">
          <b>📄 ${escapeHtml(page.wb_name)}</b>
          <span style="color:#64748b;">/ ${escapeHtml(page.sheet_name)}</span>
          <span class="page-tag">第 ${Number(page.page_index) + 1} / ${Number(page.total_pages)} 页</span>
          <span id="stagecounts_${escapeHtml(pageId)}" style="font-size:12.5px; color:#475569;"></span>
        </div>
      </div>

      <div class="page-canvas-scroll">
        <div class="page-canvas-wrap">
          <img class="page-canvas-img" src="${page.image}" alt="${escapeHtml(page.sheet_name)} 整页预览">
          <div class="hotspot-layer">
            ${hotspotsHtml}
          </div>
        </div>
      </div>

      <div class="capsule-bar">
        <span class="capsule-bar-label">🎯 本页切字定位 &amp; 行号微调：</span>
        ${capsulesHtml}
      </div>
    `;

    // 绑定热区与胶囊交互事件
    stageCard.querySelectorAll('.hotspot-box, .item-capsule').forEach(el => {
      const itemId = el.getAttribute('data-item-id');
      el.addEventListener('click', e => {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'LABEL') return;
        toggleSkip(itemId);
      });
      el.addEventListener('mouseenter', () => {
        const currentItem = latestItems[itemId];
        if (!currentItem) return;
        const hotspotEl = document.getElementById('hotspot_' + itemId) || el;
        showMagnifierForElement(currentItem, hotspotEl);
      });
      el.addEventListener('mouseleave', () => {
        const currentItem = latestItems[itemId];
        hideMagnifier(currentItem);
      });
    });
  });
}

/* ==========================================
   视图二渲染：局部 2.5x 特写卡片模式
   ========================================== */
function renderCards(items) {
  const grid = document.getElementById('cardGrid');
  const keys = Object.keys(items);
  for (const id of Object.keys(cardsSignatureCache)) {
    if (!items[id]) {
      document.getElementById('card_' + id)?.remove();
      delete cardsSignatureCache[id];
    }
  }
  if (keys.length === 0) {
    grid.innerHTML = '';
    cardsSignatureCache = {};
    return;
  }

  keys.forEach(id => {
    const item = items[id];
    let card = document.getElementById('card_' + id);
    if (!card) {
      card = document.createElement('div');
      card.id = 'card_' + id;
      card.className = 'card';
      card.onclick = e => {
        if (e.target.tagName === 'INPUT') return;
        toggleSkip(id);
      };
      card.onmouseenter = () => setRowPeerHighlight(latestItems[id] || item, true);
      card.onmouseleave = () => setRowPeerHighlight(latestItems[id] || item, false);
      grid.appendChild(card);
    }

    const activeEl = document.activeElement;
    if (activeEl && activeEl.tagName === 'INPUT' && card.contains(activeEl)) {
      return;
    }

    // 状态变化只更新样式（见 refreshItemClasses），不重建卡片
    const st = getItemVisualState(item);
    const signature = `${item.iterations}|${item.row_height}|${item.excel_row}|${item.cell_addr}|${item.image}|${item.review_reason || ''}|${item.status === 'fixed'}`;
    if (cardsSignatureCache[id] === signature) {
      return;
    }
    cardsSignatureCache[id] = signature;

    card.className = `card ${st.cardClass}`;
    card.innerHTML = `
      <div class="card-header">
        <div>
          <b>#${Number(item.seq || 1)} ${escapeHtml(item.wb_name)}</b>
          <span style="color:#64748b;">/ ${escapeHtml(item.sheet_name)} (P${Number(item.page_index) + 1})</span>
        </div>
        <span class="badge ${st.badgeClass}">${escapeHtml(st.badgeText)}</span>
      </div>
      <div class="card-img-wrap"><img src="${item.image}" alt="crop"></div>
      <div class="card-body">
        <div>
          定位: <b>${escapeHtml(item.cell_addr)}</b> | 当前行高: <b>${escapeHtml(item.row_height)}pt</b><br>
          <span style="color:#64748b;">文字: ${escapeHtml(item.text)}</span>
        </div>
        <div class="row-adjust" title="如果行号有偏差可手动微调">
          行号:
          <input
            type="number"
            min="1"
            value="${Number(item.excel_row) || 1}"
            onchange="updateRow('${escapeHtml(id)}', this.value)"
          >
        </div>
      </div>
      <div onclick="event.stopPropagation();">
        <button class="btn btn-outline btn-sm" onclick="manualOpen('${escapeHtml(id)}')">在 Excel 中定位</button>
        <button class="btn btn-outline btn-sm" title="${RECHECK_TITLE}" onclick="manualRecheck('${escapeHtml(id)}')">重新导出复检（不加高）</button>
        ${skipButton(item)}
        <p>${escapeHtml(item.review_reason || '自动调整无效时，请转人工核对行号、合并单元格及边框。')}</p>
      </div>
    `;
  });
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
    if (seq < minAcceptedSeq) return;

    canResume = data.can_resume;
    isBusy = data.busy;
    if (!canResume) skipCancelsResumeConfirmed = false;
    document.getElementById('resumeBtn').classList.toggle('hidden', !canResume);
    document.getElementById('resumeBtn').disabled = data.busy;
    if (!workspaceRestored) {
      document.getElementById('folderInput').value = data.folder || '';
      document.getElementById('stepInput').value = data.settings.step_pt;
      document.getElementById('maxIterInput').value = data.settings.max_iter;
      // 「顺便清理备份表」会直接删除且不预览，不沿用上次的勾选，每次扫描都需手动勾选
    }

    document.getElementById('scanBtn').disabled = data.busy;
    document.getElementById('fixBtn').disabled = data.busy;
    document.getElementById('refreshBtn').disabled = data.busy;
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
    applyPendingSkips();
    if (activeHoverItemId && !latestItems[activeHoverItemId]) {
      hideMagnifier();
    }

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
  if (canResume && !confirm('开始新扫描会替换上次任务及其扫描结果，确定不继续上次任务吗？')) return;
  await syncSkips();
  try {
    await postAction('/api/scan', {
      folder,
      delete_backups: document.getElementById('delBackupChk').checked,
    });
  } catch (e) {
    alert(e.message);
    return;
  }
  pagesSignatureCache = {};
  cardsSignatureCache = {};
  hideMagnifier();
  document.getElementById('canvasView').innerHTML = '';
  document.getElementById('cardGrid').innerHTML = '';

  await pollState();
}

async function startFix() {
  await syncSkips();
  const ids = Object.values(latestItems).filter(item => item.status === 'pending').map(item => item.id);
  if (ids.length === 0) {
    alert('没有待修复的项。需人工复检的项可先「重新导出复检」，确认无需修复的直接点击跳过。');
    return;
  }
  const step = Number(document.getElementById('stepInput').value);
  const maxIter = Number(document.getElementById('maxIterInput').value);
  if (!Number.isFinite(step) || step < 0.5 || step > 30 ||
      !Number.isSafeInteger(maxIter) || maxIter < 1) {
    alert('每次递增必须为 0.5 至 30 pt，最大循环次数必须为正整数。');
    return;
  }
  if (canResume && !confirm('新微调会替换上次任务的续跑进度，确定吗？')) return;
  hideMagnifier();
  try {
    await postAction('/api/fix', {
      selected_ids: ids,
      step_pt: step,
      max_iter: maxIter,
    });
    await pollState();
  } catch (e) { alert(e.message); }
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
    hideMagnifier();
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
    hideMagnifier();
    await postAction('/api/manual/recheck', { id });
    await pollState();
  } catch (e) { alert(e.message); }
}

async function refreshProblems() {
  if (canResume && !confirm('刷新复检会替换上次中断任务的续跑检查点。请确认已保存人工调整，是否继续？')) return;
  await syncSkips();
  try {
    hideMagnifier();
    await postAction('/api/manual/refresh', {});
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

let savedOnlyPending = false;
try {
  savedOnlyPending = localStorage.getItem(ONLY_PENDING_STORAGE_KEY) === '1';
} catch (e) {}
setOnlyPending(savedOnlyPending);

setInterval(pollState, 1200);
pollState();
