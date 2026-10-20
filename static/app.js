let selectedIds = new Set();
let seenItemIds = new Set();
let currentViewMode = 'canvas'; // 'canvas' | 'cards'
let latestPages = {};
let latestItems = {};
let pagesSignatureCache = {};
let cardsSignatureCache = {};
let activeHoverItemId = null;
let workspaceRestored = false;
let canResume = false;

function isResolved(item) {
  return item.status === 'fixed' || item.status === 'accepted';
}

function confirmationButton(item) {
  const accepted = item.status === 'accepted';
  if (item.status === 'fixed') return '';
  return `<button class="btn btn-outline btn-sm" onclick="confirmManual('${escapeHtml(item.id)}', ${!accepted})">${accepted ? '撤销人工确认' : '人工确认：无需修复'}</button>`;
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

function isLinkRowEnabled() {
  const chk = document.getElementById('linkRowChk');
  return chk ? chk.checked : true;
}

function setViewMode(mode) {
  currentViewMode = mode === 'cards' ? 'cards' : 'canvas';
  document.getElementById('viewCanvasBtn').classList.toggle('active', currentViewMode === 'canvas');
  document.getElementById('viewCardsBtn').classList.toggle('active', currentViewMode === 'cards');
  document.getElementById('canvasView').classList.toggle('hidden', currentViewMode !== 'canvas');
  document.getElementById('cardGrid').classList.toggle('hidden', currentViewMode !== 'cards');
  hideMagnifier();
}

function syncNewItemsDefaultSelection(items) {
  for (const id of seenItemIds) {
    if (!items[id]) {
      seenItemIds.delete(id);
      selectedIds.delete(id);
    }
  }
  Object.keys(items).forEach(id => {
    const item = items[id];
    if (!seenItemIds.has(id)) {
      seenItemIds.add(id);
      if (item.status === 'pending') {
        selectedIds.add(id);
      }
    }
    if (isResolved(item)) {
      selectedIds.delete(id);
    }
  });
}

function toggleItemSelection(id) {
  const item = latestItems[id];
  if (!item || isResolved(item)) return;

  const shouldSelect = !selectedIds.has(id);
  const targets = isLinkRowEnabled() ? getPeerItemsInSameRow(item) : [item];

  targets.forEach(t => {
    if (isResolved(t)) return;
    if (shouldSelect) {
      selectedIds.add(t.id);
    } else {
      selectedIds.delete(t.id);
    }
  });

  refreshSelectionClasses();
  if (activeHoverItemId === id) {
    updateMagnifierContent(item);
  }
}

function togglePageItems(pageId, select) {
  Object.values(latestItems).forEach(item => {
    if (item.page_id !== pageId || isResolved(item)) return;
    if (select) {
      selectedIds.add(item.id);
    } else {
      selectedIds.delete(item.id);
    }
  });
  refreshSelectionClasses();
}

function toggleAll(select) {
  Object.values(latestItems).forEach(item => {
    if (isResolved(item)) {
      selectedIds.delete(item.id);
      return;
    }
    if (select) {
      selectedIds.add(item.id);
    } else {
      selectedIds.delete(item.id);
    }
  });
  refreshSelectionClasses();
}

function updateCount() {
  document.getElementById('selCount').innerText = selectedIds.size;
  const rowKeys = new Set();
  selectedIds.forEach(id => {
    const item = latestItems[id];
    if (item) {
      rowKeys.add(getRowGroupKey(item));
    }
  });
  document.getElementById('selRowCount').innerText = rowKeys.size;
}

function getItemVisualState(item) {
  if (item.status === 'accepted') {
    return {
      boxClass: 'is-fixed', cardClass: 'fixed', badgeClass: 'badge-fixed',
      badgeText: '人工确认无需修复', shortTag: `人工✓#${item.index_on_page || item.seq}`,
    };
  }
  if (item.status === 'fixed') {
    return {
      boxClass: 'is-fixed',
      cardClass: 'fixed',
      badgeClass: 'badge-fixed',
      badgeText: `✓ 已修复 (+${item.iterations}次)`,
      shortTag: `✓#${item.index_on_page || item.seq}`,
    };
  }
  if (item.status === 'needs_review' || item.status === 'manual_review') {
    return {
      boxClass: 'is-max-reached',
      cardClass: 'max-reached',
      badgeClass: 'badge-warn',
      badgeText: item.status === 'manual_review' ? '等待人工调整并保存' : '需人工复检',
      shortTag: `?#${item.index_on_page || item.seq}`,
    };
  }
  if (item.status === 'max_reached') {
    const isSel = selectedIds.has(item.id);
    return {
      boxClass: isSel ? 'is-selected' : 'is-max-reached',
      cardClass: (isSel ? 'selected ' : '') + 'max-reached',
      badgeClass: 'badge-warn',
      badgeText: `⚠️ 达到上限，建议人工复检 (${item.iterations}次)`,
      shortTag: `⚠#${item.index_on_page || item.seq}`,
    };
  }
  const isSel = selectedIds.has(item.id);
  return {
    boxClass: isSel ? 'is-selected' : 'is-ignored',
    cardClass: isSel ? 'selected' : '',
    badgeClass: isSel ? 'badge-selected' : 'badge-ignored',
    badgeText: isSel ? '待修复 (已选中)' : '已忽略 (点击可选中)',
    shortTag: `#${item.index_on_page || item.seq}`,
  };
}

function refreshSelectionClasses() {
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
  });
  updateCount();
}

/* ==========================================
   同行高亮联动 & 2.5x 悬停放大镜气泡
   ========================================== */
function setRowPeerHighlight(item, active) {
  const peers = isLinkRowEnabled() ? getPeerItemsInSameRow(item) : [item];
  peers.forEach(peer => {
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
  document.getElementById('magHint').textContent =
    isResolved(item)
      ? (item.status === 'accepted' ? '人工确认无需修复（当前已保存版本有效，可撤销）' : '✅ 该位置已完整显示，无需继续调整')
      : '💡 点击热区即可切换「选中修复 / 忽略」；可在画布下方胶囊直接微调行号';
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

    const sigParts = [
      page.image,
      ...pageItems.map(
        it => `${it.id}:${it.status}:${it.iterations}:${it.row_height}:${it.excel_row}:${it.cell_addr}:${it.bbox.join(',')}:${it.image}`
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

    const fixedCount = pageItems.filter(it => it.status === 'fixed').length;
    const acceptedCount = pageItems.filter(it => it.status === 'accepted').length;
    const pendingCount = pageItems.length - fixedCount - acceptedCount;

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
              <button class="btn btn-outline btn-sm" onclick="manualRecheck('${escapeHtml(item.id)}')">重新导出复检</button>
              ${confirmationButton(item)}
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
          <span style="font-size:12.5px; color:#475569;">
            （待处理 <b>${pendingCount}</b> 处 · 已修复 <b style="color:#059669;">${fixedCount}</b> 处 · 人工确认 ${acceptedCount} 处）
          </span>
        </div>
        <div class="page-stage-actions">
          <button class="btn btn-outline btn-sm" onclick="togglePageItems('${escapeHtml(pageId)}', true)">本页全选</button>
          <button class="btn btn-outline btn-sm" onclick="togglePageItems('${escapeHtml(pageId)}', false)">本页取消</button>
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
        toggleItemSelection(itemId);
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
        toggleItemSelection(id);
      };
      card.onmouseenter = () => setRowPeerHighlight(latestItems[id] || item, true);
      card.onmouseleave = () => setRowPeerHighlight(latestItems[id] || item, false);
      grid.appendChild(card);
    }

    const activeEl = document.activeElement;
    if (activeEl && activeEl.tagName === 'INPUT' && card.contains(activeEl)) {
      return;
    }

    const st = getItemVisualState(item);
    const signature = `${item.status}|${item.iterations}|${item.row_height}|${item.excel_row}|${item.cell_addr}|${item.image}|${item.review_reason || ''}|${selectedIds.has(id)}`;
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
        <button class="btn btn-outline btn-sm" onclick="manualRecheck('${escapeHtml(id)}')">重新导出复检（不加高）</button>
        ${confirmationButton(item)}
        <p>${escapeHtml(item.review_reason || '自动调整无效时，请转人工核对行号、合并单元格及边框。')}</p>
      </div>
    `;
  });
}

/* ==========================================
   后端状态轮询与接口调用
   ========================================== */
async function pollState() {
  try {
    const res = await fetch('/api/state');
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || '读取任务失败');

    canResume = data.can_resume;
    document.getElementById('resumeBtn').classList.toggle('hidden', !canResume);
    document.getElementById('resumeBtn').disabled = data.busy;
    if (!workspaceRestored) {
      document.getElementById('folderInput').value = data.folder || '';
      document.getElementById('stepInput').value = data.settings.step_pt;
      document.getElementById('maxIterInput').value = data.settings.max_iter;
      document.getElementById('delBackupChk').checked = data.settings.delete_backups;
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
    latestItems = data.items || {};
    if (activeHoverItemId && !latestItems[activeHoverItemId]) {
      hideMagnifier();
    }

    const hasResults = Object.keys(latestPages).length > 0 || Object.keys(latestItems).length > 0;
    const emptyStage = document.getElementById('emptyStagePlaceholder');
    if (emptyStage) {
      emptyStage.classList.toggle('hidden', hasResults);
    }

    syncNewItemsDefaultSelection(latestItems);
    if (!workspaceRestored && data.job && data.job.kind === 'fix') {
      selectedIds = new Set(data.job.options.selected_ids.filter(
        id => latestItems[id] && !isResolved(latestItems[id])
      ));
    }
    workspaceRestored = true;
    renderCanvasStages(latestPages, latestItems);
    renderCards(latestItems);
    refreshSelectionClasses();
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
  try {
    await postAction('/api/scan', {
      folder,
      delete_backups: document.getElementById('delBackupChk').checked,
    });
  } catch (e) {
    alert(e.message);
    return;
  }
  selectedIds.clear();
  seenItemIds.clear();
  pagesSignatureCache = {};
  cardsSignatureCache = {};
  hideMagnifier();
  document.getElementById('canvasView').innerHTML = '';
  document.getElementById('cardGrid').innerHTML = '';
  updateCount();

  await pollState();
}

async function startFix() {
  if (selectedIds.size === 0) {
    alert('请先在画布或卡片上选中要修复的位置！');
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
      selected_ids: Array.from(selectedIds),
      step_pt: step,
      max_iter: maxIter,
    });
    await pollState();
  } catch (e) { alert(e.message); }
}

async function resumeTask() {
  try {
    await postAction('/api/resume', {});
    await pollState();
  } catch (e) { alert(e.message); }
}

async function manualOpen(id) {
  if (canResume && !confirm('转人工将替换中断任务，旧检查点不再自动续跑。剩余项可重新选中处理。确定吗？')) return;
  try {
    hideMagnifier();
    await postAction('/api/manual/open', { id });
    await pollState();
  } catch (e) { alert(e.message); }
}

async function manualRecheck(id) {
  const message = '请先在 Excel 中保存手动修改。本次根据新 PDF 重新检测该文件的问题工作表并替换旧标记，不会自动增加行高。'
    + (canResume ? '\n此操作会替换上次中断任务的续跑检查点。' : '');
  if (!confirm(message)) return;
  try {
    hideMagnifier();
    await postAction('/api/manual/recheck', { id });
    await pollState();
  } catch (e) { alert(e.message); }
}

async function refreshProblems() {
  if (canResume && !confirm('刷新复检会替换上次中断任务的续跑检查点。请确认已保存人工调整，是否继续？')) return;
  try {
    hideMagnifier();
    await postAction('/api/manual/refresh', {});
    await pollState();
  } catch (e) { alert(e.message); }
}

async function confirmManual(id, accepted) {
  const message = '确认此处无需修复？请先保存 Excel。人工结论仅对当前已保存版本有效，文件再次修改后需重新检查。'
    + (canResume ? '\n此操作会取消中断任务的续跑，避免旧检查点覆盖人工结论。其他未解决项仍可重新选择处理。' : '');
  if (accepted && !confirm(message)) return;
  try {
    hideMagnifier();
    await postAction('/api/manual/confirm', { id, accepted });
    if (!accepted) selectedIds.add(id);
    await pollState();
  } catch (e) { alert(e.message); }
}

setInterval(pollState, 1200);
pollState();
