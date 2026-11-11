// 主页工具箱（用户 2026-11-10 要）：所有工具卡都放主页，按住卡片拖进上面的「本次流程」轨道排成第 1 步、第 2 步……
// - 第一行固定 精简 → 调整页面 → 查看切字 → 上传；下面的卡按用过的次数排，只在打开主页时排一次，操作时卡片不挪位置
// - 拖的时候只出一条落点竖线，轨道里的步骤和卡片都不动；松手才改流程
// - 精简卡下面勾「顺便做」哪几项（和精简页的勾选共用），上传卡写明上传前每次都会做的三项
// 工具清单、流程存取在 static/flow.js
(function () {
  const F = window.Flow;
  if (!F) return;
  const esc = F.esc;
  const road = document.getElementById('flowRoad');
  const grid = document.getElementById('toolGrid');
  const moreGrid = document.getElementById('toolGridMore');
  if (!road || !grid) return;

  let flow = F.load();
  let dragging = null; // { id, from: 'card' | 'step' }

  function cardHtml(t) {
    const vectors = { bundle: 'document', json: 'layers', colors: 'palette', images: 'image', font: 'type', links: 'link', bl: 'link', hidden: 'eye', view: 'target' };
    const icon = t.main
      ? `<span class="feature-icon garden-sprite sprite-${t.id}"></span>`
      : `<span class="tool-vector"><svg class="ui-icon" aria-hidden="true"><use href="/static/icons/garden-ui.svg#${vectors[t.id]}"></use></svg></span>`;
    const attrs = t.action ? `type="button" onclick="${t.action}"` : `href="${esc(t.url)}"${t.linkId ? ` id="${t.linkId}"` : ''}`;
    const tag = t.action ? 'button' : 'a';
    return `<li class="route-stop tool-stop" data-tool="${t.id}"${t.ops ? ` data-ops="${t.ops}"` : ''}>
      <span class="route-done hidden" role="img"></span>
      <${tag} ${attrs} class="feature-card tone-${t.tone}" draggable="true" data-tool="${t.id}" title="单击打开；按住拖到上面的「本次流程」里，排进流程">
        <span class="feature-text">
          <span class="feature-title">${esc(t.title)}</span>
          <span class="feature-sub">${esc(t.sub)}</span>
          <span class="feature-meta"${t.metaId ? ` id="${t.metaId}"` : ''}>${esc(t.meta)}</span>
        </span>
        <span class="feature-orb" aria-hidden="true">${icon}</span>
      </${tag}>
      <span class="tool-inc hidden"></span>
      ${extrasHtml(t)}
    </li>`;
  }

  function extrasHtml(t) {
    if (t.id === 'clean') {
      const on = F.chainOptions();
      return `<div class="tool-extras" aria-label="精简后顺便做">顺便做：${F.CHAIN.map(c =>
        `<label><input type="checkbox" data-chain="${c.key}"${on[c.key] ? ' checked' : ''}> ${esc(c.label)}</label>`).join('')}</div>`;
    }
    if (t.id === 'upload') {
      return '<div class="tool-extras is-fixed" title="上传前每份都会做，不用单独点">上传前都会做：<span>✓ 清标记</span><span>✓ 恢复颜色</span><span>✓ 标准化</span></div>';
    }
    return '';
  }

  // 用过的次数：所有文件夹的执行记录加起来（读不到就按默认顺序）
  async function usageCounts() {
    try {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 1500);
      const res = await fetch('/api/operation-history', { signal: ctrl.signal });
      clearTimeout(timer);
      if (!res.ok) return {};
      const counts = {};
      ((await res.json()).folders || []).forEach(group => {
        Object.entries(group.operations || {}).forEach(([kind, entry]) => { counts[kind] = (counts[kind] || 0) + (entry.count || 0); });
      });
      return counts;
    } catch (_) {
      return {};
    }
  }

  function toolUses(t, counts) {
    return (t.ops || '').split(' ').filter(Boolean).reduce((n, k) => n + (counts[k] || 0), 0);
  }

  // 第一行四张大卡先放，页面马上能用（home.js 要找卡上的状态文字）
  grid.innerHTML = F.TOOLS.filter(t => t.main).map(cardHtml).join('');

  async function renderMore() {
    const counts = await usageCounts();
    const rest = F.TOOLS.filter(t => !t.main);
    const order = new Map(rest.map((t, i) => [t.id, i]));
    // 独立的工具（不被大卡顺便做的）放前面，再按用过的次数，一样多就按默认顺序
    rest.sort((a, b) => (Boolean(a.inc) - Boolean(b.inc)) || (toolUses(b, counts) - toolUses(a, counts)) || (order.get(a.id) - order.get(b.id)));
    moreGrid.innerHTML = rest.map(cardHtml).join('');
    paintCards();
    window.refreshDoneMarks?.();
  }

  // ---------- 卡片上的「第 N 步」「已含在…里」 ----------
  function paintCards() {
    const chain = F.chainOptions();
    document.querySelectorAll('.tool-stop').forEach(li => {
      const t = F.byId[li.dataset.tool];
      const pos = flow.indexOf(t.id);
      li.classList.toggle('in-flow', pos >= 0);
      li.querySelector('.feature-card').style.setProperty('--tile-cat', `'${pos >= 0 ? `第 ${pos + 1} 步` : (t.inc ? '单独用' : '工具')}'`);
      const inc = li.querySelector('.tool-inc');
      const included = pos < 0 && t.inc && flow.includes(t.inc) && (!t.incKey || chain[t.incKey]);
      inc.textContent = included ? `已含在「${F.byId[t.inc].title}」里` : '';
      inc.classList.toggle('hidden', !included);
    });
  }

  // ---------- 轨道 ----------
  function renderRoad() {
    const doneIds = new Set([...document.querySelectorAll('.tool-stop.is-done')]
      .filter(li => !li.querySelector('.route-done.is-warn')).map(li => li.dataset.tool));
    const current = flow.find(id => !doneIds.has(id));
    const steps = flow.map((id, i) => {
      const t = F.byId[id];
      const st = doneIds.has(id) ? 'done' : (id === current ? 'cur' : '');
      return `<span class="flow-link" aria-hidden="true"></span>
        <span class="flow-step ${st}" draggable="true" data-index="${i}" title="按住拖动调顺序">
          <span class="flow-no">${st === 'done' ? '✓' : i + 1}</span>
          <a class="flow-name" href="${esc(t.url)}" draggable="false">${esc(t.title)}</a>
          ${st === 'cur' ? '<span class="flow-now">现在</span>' : ''}
          <button type="button" class="flow-del" data-index="${i}" title="从流程里去掉" aria-label="从流程里去掉 ${esc(t.title)}">✕</button>
        </span>`;
    }).join('');
    const empty = flow.length ? '' : '<span class="flow-link" aria-hidden="true"></span><span class="flow-empty">把下面的工具卡拖到这里</span>';
    road.innerHTML = `<a class="flow-sign flow-start" href="/distribute/start" id="workTaskNew" title="打开「新任务 · 拿文件」页：填共享文件夹路径，自动生成参考、处理文件夹，再从共享拿文件"><span class="garden-sprite sprite-new" aria-hidden="true"></span>新任务</a>
      ${steps}${empty}<span class="flow-link" aria-hidden="true"></span>
      <a class="flow-sign flow-finish" href="/finish?start=1" id="workTaskFinish" title="打开「完成任务 · 收尾检查」页：逐项检查处理好的文件都传回共享了没有，通过就标记任务完成"><span class="garden-sprite sprite-finish" aria-hidden="true"></span>完成任务</a>
      <i class="flow-marker hidden" aria-hidden="true"></i>`;
  }

  function setFlow(ids) {
    flow = ids;
    F.save(ids);
    renderRoad();
    paintCards();
    window.refreshDoneMarks?.();
  }

  // 落点：鼠标在哪一步的左半边就插在它前面；轨道折行时按同一行算
  function dropIndex(x, y) {
    const chips = [...road.querySelectorAll('.flow-step')];
    let index = chips.length;
    for (let i = 0; i < chips.length; i++) {
      const r = chips[i].getBoundingClientRect();
      if (y < r.top - 6) { index = i; break; }
      if (y <= r.bottom + 6 && x < r.left + r.width / 2) { index = i; break; }
    }
    return index;
  }

  function showMarker(index) {
    const marker = road.querySelector('.flow-marker');
    const chips = [...road.querySelectorAll('.flow-step')];
    const base = road.getBoundingClientRect();
    let ref;
    let left;
    if (index < chips.length) {
      ref = chips[index].getBoundingClientRect();
      left = ref.left - base.left - 11;
    } else {
      ref = (chips[chips.length - 1] || road.querySelector('.flow-empty') || road.querySelector('.flow-start')).getBoundingClientRect();
      left = ref.right - base.left + 9;
    }
    marker.style.left = `${left}px`;
    marker.style.top = `${ref.top - base.top - 4}px`;
    marker.style.height = `${ref.height + 8}px`;
    marker.classList.remove('hidden');
  }

  function hideMarker() {
    road.querySelector('.flow-marker')?.classList.add('hidden');
    road.closest('.flow-bar').classList.remove('dropzone');
  }

  document.addEventListener('dragstart', e => {
    const card = e.target.closest?.('.tool-stop .feature-card');
    const step = e.target.closest?.('.flow-step');
    if (card) dragging = { id: card.dataset.tool, from: 'card' };
    else if (step) dragging = { id: flow[Number(step.dataset.index)], from: 'step' };
    else return;
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', F.byId[dragging.id].title);
    road.closest('.flow-bar').classList.add('dropzone');
  });

  document.addEventListener('dragend', () => {
    dragging = null;
    hideMarker();
  });

  const bar = road.closest('.flow-bar');
  bar.addEventListener('dragover', e => {
    if (!dragging) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    showMarker(dropIndex(e.clientX, e.clientY));
  });

  bar.addEventListener('drop', e => {
    if (!dragging) return;
    e.preventDefault();
    let index = dropIndex(e.clientX, e.clientY);
    const ids = flow.slice();
    const old = ids.indexOf(dragging.id);
    if (old >= 0) {
      ids.splice(old, 1); // 已经在流程里：挪到新位置，不会重复
      if (old < index) index -= 1;
    }
    ids.splice(index, 0, dragging.id);
    dragging = null;
    hideMarker();
    setFlow(ids);
  });

  road.addEventListener('click', e => {
    const del = e.target.closest('.flow-del');
    if (!del) return;
    const ids = flow.slice();
    ids.splice(Number(del.dataset.index), 1);
    setFlow(ids);
  });

  document.getElementById('flowResetBtn')?.addEventListener('click', () => setFlow(F.DEFAULT_FLOW.slice()));
  document.getElementById('flowClearBtn')?.addEventListener('click', async () => {
    if (!flow.length) return;
    if (window.appConfirm && !(await window.appConfirm('把本次流程里的步骤全部去掉吗？工具卡都还在，可以重新拖进来。', { okText: '清空' }))) return;
    setFlow([]);
  });

  grid.addEventListener('change', e => {
    const box = e.target.closest('input[data-chain]');
    if (!box) return;
    F.setChainOption(box.dataset.chain, box.checked);
    paintCards();
  });

  // 切换任务（处理文件夹变了）：换成那个任务自己的流程
  window.addEventListener('storage', e => { if (e.key === 'clear_handle_excel.lastFolder') { flow = F.load(); renderRoad(); paintCards(); } });
  window.HomeToolbox = {
    reload() { flow = F.load(); renderRoad(); paintCards(); },
    renderRoad,
    flow: () => flow.slice(),
  };

  renderRoad();
  paintCards();
  renderMore();
})();
