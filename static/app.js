let selectedIds = new Set();
let seenItemIds = new Set();
let currentViewMode = 'canvas'; // 'canvas' | 'cards'
let latestPages = {};
let latestItems = {};
let pagesSignatureCache = {};
let cardsSignatureCache = {};
let activeHoverItemId = null;

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
  Object.keys(items).forEach(id => {
    const item = items[id];
    if (!seenItemIds.has(id)) {
      seenItemIds.add(id);
      if (item.status === 'pending') {
        selectedIds.add(id);
      }
    }
    if (item.status === 'fixed') {
      selectedIds.delete(id);
    }
  });
}

function toggleItemSelection(id) {
  const item = latestItems[id];
  if (!item || item.status === 'fixed') return;

  const shouldSelect = !selectedIds.has(id);
  const targets = isLinkRowEnabled() ? getPeerItemsInSameRow(item) : [item];

  targets.forEach(t => {
    if (t.status === 'fixed') return;
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
    if (item.page_id !== pageId || item.status === 'fixed') return;
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
    if (item.status === 'fixed') {
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
  if (item.status === 'fixed') {
    return {
      boxClass: 'is-fixed',
      cardClass: 'fixed',
      badgeClass: 'badge-fixed',
      badgeText: `✓ 已修复 (+${item.iterations}次)`,
      shortTag: `✓#${item.index_on_page || item.seq}`,
    };
  }
  if (item.status === 'max_reached') {
    const isSel = selectedIds.has(item.id);
    return {
      boxClass: isSel ? 'is-selected' : 'is-max-reached',
      cardClass: (isSel ? 'selected ' : '') + 'max-reached',
      badgeClass: 'badge-warn',
      badgeText: `⚠️ 达最大微调次数 (+${item.iterations}次)`,
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
    item.status === 'fixed'
      ? '✅ 该位置已完整显示，无需继续调整'
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
    const pendingCount = pageItems.length - fixedCount;

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
            （待处理 <b>${pendingCount}</b> 处 · 已修复 <b style="color:#059669;">${fixedCount}</b> 处）
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
    const signature = `${item.status}|${item.iterations}|${item.row_height}|${item.excel_row}|${item.cell_addr}|${item.image}|${selectedIds.has(id)}`;
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

    document.getElementById('scanBtn').disabled = data.busy;
    document.getElementById('fixBtn').disabled = data.busy;
    document.getElementById('statusText').innerText = data.busy ? '⏳ 后台处理中...' : '✅ 就绪';

    const logBox = document.getElementById('logBox');
    logBox.innerText = (data.logs && data.logs.length > 0) ? data.logs.join('\n') : '等待操作...';
    logBox.scrollTop = logBox.scrollHeight;

    latestPages = data.pages || {};
    latestItems = data.items || {};

    syncNewItemsDefaultSelection(latestItems);
    renderCanvasStages(latestPages, latestItems);
    renderCards(latestItems);
    refreshSelectionClasses();
  } catch (e) {}
}

async function updateRow(id, newRow) {
  const parsed = parseInt(newRow, 10);
  if (!Number.isFinite(parsed) || parsed < 1) return;
  await fetch('/api/update_row', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ id, excel_row: parsed }),
  });
}

async function startScan() {
  const folder = document.getElementById('folderInput').value.trim();
  if (!folder) {
    alert('请先输入 Excel 文件夹路径！');
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

  await fetch('/api/scan', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      folder,
      delete_backups: document.getElementById('delBackupChk').checked,
    }),
  });
}

async function startFix() {
  if (selectedIds.size === 0) {
    alert('请先在画布或卡片上选中要修复的位置！');
    return;
  }
  const step = parseFloat(document.getElementById('stepInput').value) || 3.0;
  hideMagnifier();
  await fetch('/api/fix', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      selected_ids: Array.from(selectedIds),
      step_pt: step,
    }),
  });
}

setInterval(pollState, 1200);
pollState();
