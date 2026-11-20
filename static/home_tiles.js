// 主页桌面（用户 2026-11-20 要：把项目做成自己的编程工具箱）
// - 所有工具都是一样大的小图标（像手机桌面 APP），收在「工具箱」文件夹里，点开展开
// - 主页上可以有好几条「流程」：从工具箱把图标拖进流程，流程里拖动调顺序，✕ 去掉，流程名字点一下就能改
// - 流程存在浏览器 localStorage（clear_handle_excel.flows），换 ZIP 版本不丢
// - 拖的时候只出一条发光的落点竖线，别的图标不挪；不自动滚动
// - 从某一页点「返回首页」回来（/?from=工具）：流程里这个工具的图标闪几下；流程里没有就闪工具箱
// 工具清单在 static/tools.js；调试日志都以「[工具箱]」开头打在浏览器控制台
(function () {
  const T = window.Tools;
  const desk = document.getElementById('homeDesk');
  if (!T || !desk) return;
  const esc = T.esc;
  const log = (...args) => console.info('[工具箱]', ...args);

  const FLOWS_KEY = 'clear_handle_excel.flows';
  const DEFAULT_FLOWS = [
    { name: 'Excel 交付流程', steps: ['new', 'clean', 'adjust', 'cells', 'upload', 'finish'] },
  ];

  // ---------- 流程存取 ----------
  function newId() { return 'f' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

  function loadFlows() {
    let raw = null;
    try { raw = localStorage.getItem(FLOWS_KEY); } catch (_) {}
    if (raw) {
      try {
        const data = JSON.parse(raw);
        const flows = (Array.isArray(data.flows) ? data.flows : []).map(f => ({
          id: String(f.id || newId()),
          name: String(f.name || '未命名流程').slice(0, 40),
          steps: (Array.isArray(f.steps) ? f.steps : []).filter(id => T.byId[id]),
        }));
        log('读到流程', flows.length, '条');
        return flows;
      } catch (e) {
        console.warn('[工具箱] 流程记录读不懂，改用默认流程', e);
      }
    }
    log('没有流程记录，用默认流程');
    return DEFAULT_FLOWS.map(f => ({ id: newId(), name: f.name, steps: f.steps.slice() }));
  }

  let flows = loadFlows();

  function saveFlows(reason) {
    try { localStorage.setItem(FLOWS_KEY, JSON.stringify({ version: 1, flows })); } catch (e) { console.warn('[工具箱] 存不了流程', e); }
    log('已保存：' + reason, flows.map(f => `${f.name}(${f.steps.join('→')})`).join(' | '));
  }

  const flowById = id => flows.find(f => f.id === id);

  // ---------- 图标 ----------
  function appHtml(t, extra = '') {
    const inner = `<span class="app-icon" aria-hidden="true"><span class="harbor-icon icon-${t.icon}"></span></span>
      <span class="app-name">${esc(t.title)}</span><span class="app-done hidden" role="img"></span>`;
    const common = `class="app tone-${t.tone}" data-tool="${t.id}"${t.ops ? ` data-ops="${t.ops}"` : ''} title="${esc(t.title)}：${esc(t.sub)}" draggable="false" ${extra}`;
    return t.task
      ? `<button type="button" ${common} data-task="${t.task}">${inner}</button>`
      : `<a href="${esc(t.url)}" ${common}>${inner}</a>`;
  }

  // ---------- 换过的工具图标（设置图标页，存在程序数据\图标库） ----------
  // 用一段样式盖掉那个工具的精灵图，流程重画、工具箱展开都不用再管；图片和原来一样大，不会挤动别的图标
  fetch('/api/icons', { cache: 'no-store' }).then(r => r.json()).then(data => {
    if (!data || !data.ok) return;
    const icons = Object.fromEntries((data.icons || []).map(i => [i.id, i]));
    const rules = Object.entries(data.tools || {}).filter(([tool, id]) => T.byId[tool] && icons[id]).map(([tool, id]) =>
      `[data-tool="${tool}"] .harbor-icon, .folder-mini[data-tool="${tool}"] .harbor-icon { background: url('/api/icons/file?id=${id}&kind=png&v=${icons[id].version}') center / contain no-repeat; }`);
    if (!rules.length) return;
    const style = document.createElement('style');
    style.id = 'customToolIcons';
    style.textContent = rules.join('\n');
    document.head.appendChild(style);
    log('换过图标的工具', Object.keys(data.tools).join(', '));
  }).catch(e => console.warn('[工具箱] 读不了图标库，工具图标用默认的', e));

  // ---------- 工具箱文件夹 ----------
  const folderBtn = document.getElementById('toolboxFolder');
  const pop = document.getElementById('toolboxPop');
  const popBody = document.getElementById('toolboxBody');

  folderBtn.querySelector('.folder-preview').innerHTML = T.TOOLS.slice(0, 9)
    .map(t => `<span class="folder-mini tone-${t.tone}" data-tool="${t.id}"><span class="harbor-icon icon-${t.icon}"></span></span>`).join('');
  folderBtn.querySelector('.folder-count').textContent = `${T.TOOLS.length} 个工具`;

  popBody.innerHTML = T.GROUPS.map(g => {
    const tools = T.TOOLS.filter(t => t.group === g.id);
    if (!tools.length) return '';
    return `<section class="toolbox-group"><h3>${esc(g.title)}</h3><div class="app-grid">${tools.map(t => appHtml(t)).join('')}</div></section>`;
  }).join('');

  function openToolbox() {
    pop.hidden = false;
    folderBtn.setAttribute('aria-expanded', 'true');
    requestAnimationFrame(() => pop.classList.add('is-open'));
    log('打开工具箱');
  }
  function closeToolbox() {
    if (pop.hidden) return;
    pop.classList.remove('is-open', 'is-peek');
    folderBtn.setAttribute('aria-expanded', 'false');
    pop.hidden = true;
    log('收起工具箱');
  }
  folderBtn.addEventListener('click', openToolbox);
  document.getElementById('toolboxClose').addEventListener('click', closeToolbox);
  pop.addEventListener('click', e => { if (e.target === pop) closeToolbox(); });
  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    if (drag) { cancelDrag(); return; }
    if (!pop.hidden && !document.querySelector('dialog[open]')) closeToolbox();
  });

  // ---------- 流程 ----------
  const flowList = document.getElementById('flowList');

  function flowHtml(f) {
    const steps = f.steps.map((id, i) => `${i ? '<span class="flow-arrow" aria-hidden="true">›</span>' : ''}
      <div class="flow-step" data-index="${i}"><span class="flow-num">${i + 1}</span>${appHtml(T.byId[id])}
        <button type="button" class="flow-remove" data-remove="${i}" title="从这个流程去掉「${esc(T.byId[id].title)}」" aria-label="去掉">✕</button></div>`).join('');
    return `<article class="flow-card" data-flow="${f.id}">
      <header class="flow-head">
        <span class="flow-badge" aria-hidden="true">⚑</span>
        <h3 class="flow-name" title="点一下改名字">${esc(f.name)}</h3>
        <button type="button" class="flow-icon-btn" data-rename title="改名字" aria-label="改名字">✎</button>
        <span class="flow-count">${f.steps.length} 步</span>
        <button type="button" class="flow-icon-btn flow-delete" data-delete title="删除这个流程" aria-label="删除这个流程">🗑</button>
      </header>
      <div class="flow-rail${f.steps.length ? '' : ' is-empty'}">
        ${steps}
        <div class="flow-slot" aria-hidden="true">${f.steps.length ? '＋ 拖到这里' : '打开工具箱，把图标拖到这里'}</div>
        <i class="flow-caret" aria-hidden="true"></i>
      </div>
    </article>`;
  }

  function renderFlows() {
    flowList.innerHTML = flows.length ? flows.map(flowHtml).join('')
      : '<p class="flow-none">还没有流程。点「＋ 新建流程」，再从工具箱把图标拖进来。</p>';
    window.paintDoneMarks?.();
  }

  function startRename(card) {
    const f = flowById(card.dataset.flow);
    const h = card.querySelector('.flow-name');
    if (!f || h.isContentEditable) return;
    h.contentEditable = 'plaintext-only';
    if (h.contentEditable !== 'plaintext-only') h.contentEditable = 'true';
    h.classList.add('is-editing');
    h.focus({ preventScroll: true });
    const range = document.createRange();
    range.selectNodeContents(h);
    const sel = getSelection(); sel.removeAllRanges(); sel.addRange(range);
    let done = false;
    const finish = save => {
      if (done) return;
      done = true;
      h.removeAttribute('contenteditable');
      h.classList.remove('is-editing');
      const name = h.textContent.replace(/\s+/g, ' ').trim().slice(0, 40);
      if (save && name && name !== f.name) {
        log('流程改名', f.name, '→', name);
        f.name = name;
        saveFlows('改名');
      }
      h.textContent = f.name;
    };
    h.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); finish(true); h.blur(); }
      else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); finish(false); h.blur(); }
    });
    h.addEventListener('blur', () => finish(true), { once: true });
  }

  flowList.addEventListener('click', async e => {
    const card = e.target.closest('.flow-card');
    if (!card) return;
    const f = flowById(card.dataset.flow);
    if (!f) return;
    const rm = e.target.closest('[data-remove]');
    if (rm) {
      const i = Number(rm.dataset.remove);
      log('从流程去掉', f.name, '第', i + 1, '步', f.steps[i]);
      f.steps.splice(i, 1);
      saveFlows('去掉一步');
      renderFlows();
      return;
    }
    if (e.target.closest('[data-rename]') || e.target.closest('.flow-name')) { startRename(card); return; }
    if (e.target.closest('[data-delete]')) {
      const ok = await appConfirm(`删除流程「${f.name}」？\n只删这条流程，工具箱里的工具不受影响。`,
        { tone: 'danger', title: '删除流程', okText: '删除' });
      if (!ok) return;
      log('删除流程', f.name);
      flows = flows.filter(x => x !== f);
      saveFlows('删除流程');
      renderFlows();
    }
  });

  document.getElementById('flowAdd').addEventListener('click', () => {
    let n = flows.length + 1;
    while (flows.some(f => f.name === `新流程 ${n}`)) n++;
    const f = { id: newId(), name: `新流程 ${n}`, steps: [] };
    flows.push(f);
    saveFlows('新建流程');
    renderFlows();
    startRename(flowList.querySelector(`.flow-card[data-flow="${f.id}"]`));
  });

  // ---------- 点图标：打开工具 ----------
  let suppressClick = false;
  const inScope = el => el && (desk.contains(el) || pop.contains(el));
  document.addEventListener('click', e => {
    const app = e.target.closest('.app');
    if (suppressClick && !app) { suppressClick = false; return; }
    if (!inScope(app)) return;
    if (suppressClick) { e.preventDefault(); e.stopPropagation(); suppressClick = false; return; }
    if (app.dataset.task) {
      e.preventDefault();
      log('打开', app.dataset.tool);
      if (typeof window.openTask === 'function') window.openTask(app.dataset.task);
    } else {
      log('打开', app.dataset.tool, app.getAttribute('href'));
    }
  }, true);

  // ---------- 拖放（自己用指针事件做，好控制落点和「别的图标不挪」） ----------
  let press = null; // 按下还没拖动
  let drag = null;  // 正在拖

  document.addEventListener('pointerdown', e => {
    if (e.button !== 0) return;
    const app = e.target.closest('.app');
    if (!inScope(app)) return;
    const step = app.closest('.flow-step');
    press = { app, x: e.clientX, y: e.clientY, pointerId: e.pointerId,
      fromFlow: step ? step.closest('.flow-card').dataset.flow : null, fromIndex: step ? Number(step.dataset.index) : -1 };
  });

  window.addEventListener('pointermove', e => {
    if (press && !drag && e.pointerId === press.pointerId && Math.hypot(e.clientX - press.x, e.clientY - press.y) > 6) beginDrag(e);
    if (drag) moveDrag(e);
  });
  window.addEventListener('pointerup', e => {
    if (drag) endDrag(e);
    press = null;
  });
  window.addEventListener('pointercancel', () => { if (drag) cancelDrag(); press = null; });
  window.addEventListener('blur', () => { if (drag) cancelDrag(); });

  function beginDrag(e) {
    const { app } = press;
    const r = app.getBoundingClientRect();
    const ghost = app.cloneNode(true);
    ghost.classList.add('app-ghost');
    ghost.removeAttribute('href');
    ghost.style.width = r.width + 'px';
    document.body.appendChild(ghost);
    drag = { ...press, tool: app.dataset.tool, ghost, dx: press.x - r.left, dy: press.y - r.top, target: null };
    app.classList.add('is-drag-source');
    document.body.classList.add('is-app-dragging');
    if (!pop.hidden && pop.contains(app)) pop.classList.add('is-peek'); // 工具箱先变透明，露出后面的流程
    log('开始拖', drag.tool, drag.fromFlow ? `（从流程第 ${drag.fromIndex + 1} 步）` : '（从工具箱）');
    moveDrag(e);
  }

  function dropTarget(x, y) {
    const el = document.elementFromPoint(x, y);
    const card = el && el.closest('.flow-card');
    if (!card || !flowList.contains(card)) return null;
    const rail = card.querySelector('.flow-rail');
    const steps = [...rail.querySelectorAll('.flow-step')];
    let index = steps.length;
    let best = null;
    steps.forEach((s, i) => {
      const r = s.getBoundingClientRect();
      if (y < r.top - 14 || y > r.bottom + 14) return; // 只看同一行
      const d = Math.abs(x - (r.left + r.width / 2));
      if (!best || d < best.d) best = { d, i, r };
    });
    if (best) index = x < best.r.left + best.r.width / 2 ? best.i : best.i + 1;
    return { card, rail, steps, index };
  }

  function showCaret(t) {
    document.querySelectorAll('.flow-rail.is-target').forEach(r => { if (!t || r !== t.rail) r.classList.remove('is-target'); });
    if (!t) return;
    t.rail.classList.add('is-target');
    const caret = t.rail.querySelector('.flow-caret');
    const base = t.rail.getBoundingClientRect();
    let left, top, height;
    const ref = t.steps[t.index] || null;
    if (ref) {
      const r = ref.getBoundingClientRect();
      left = r.left - base.left - 12; top = r.top - base.top; height = r.height;
    } else {
      const slot = t.rail.querySelector('.flow-slot').getBoundingClientRect();
      left = slot.left - base.left - 6; top = slot.top - base.top; height = slot.height;
    }
    caret.style.cssText = `left:${left}px;top:${top}px;height:${height}px`;
    t.rail.querySelector('.flow-slot').textContent = t.steps.length ? `松开放到第 ${t.index + 1} 步` : '松开放到第 1 步';
  }

  function moveDrag(e) {
    drag.ghost.style.transform = `translate(${e.clientX - drag.dx}px, ${e.clientY - drag.dy}px) scale(1.08)`;
    const t = dropTarget(e.clientX, e.clientY);
    drag.target = t;
    showCaret(t);
  }

  function clearDrag() {
    if (!drag) return;
    drag.ghost.remove();
    drag.app.classList.remove('is-drag-source');
    document.body.classList.remove('is-app-dragging');
    pop.classList.remove('is-peek');
    showCaret(null);
    document.querySelectorAll('.flow-slot').forEach(s => {
      s.textContent = s.closest('.flow-rail').classList.contains('is-empty') ? '打开工具箱，把图标拖到这里' : '＋ 拖到这里';
    });
    drag = null;
  }

  function cancelDrag() {
    log('取消拖放');
    clearDrag();
    suppressClick = true;
    setTimeout(() => { suppressClick = false; }, 0);
  }

  function endDrag(e) {
    const d = drag;
    const t = dropTarget(e.clientX, e.clientY);
    clearDrag();
    suppressClick = true; // 拖完松手不算点击，不打开工具
    setTimeout(() => { suppressClick = false; }, 0);
    if (!t) { log('没放进流程，什么都不改'); return; }
    const to = flowById(t.card.dataset.flow);
    let index = t.index;
    if (d.fromFlow) {
      const from = flowById(d.fromFlow);
      if (from === to && (index === d.fromIndex || index === d.fromIndex + 1)) { log('放回原位，不变'); return; }
      from.steps.splice(d.fromIndex, 1);
      if (from === to && index > d.fromIndex) index--;
      log(from === to ? `「${to.name}」调顺序：第 ${d.fromIndex + 1} 步 → 第 ${index + 1} 步` : `从「${from.name}」挪到「${to.name}」第 ${index + 1} 步`, d.tool);
    } else {
      log(`加进「${to.name}」第 ${index + 1} 步`, d.tool);
    }
    to.steps.splice(index, 0, d.tool);
    saveFlows('拖放');
    renderFlows();
  }

  // ---------- 从哪一页回来的：图标闪几下 ----------
  const SCROLL_KEY = 'clear_handle_excel.homeScroll';
  window.addEventListener('pagehide', () => {
    try { sessionStorage.setItem(SCROLL_KEY, String(window.scrollY)); } catch (_) {}
    closeToolbox(); // 离开时收起，bfcache 恢复回来才不会带着展开的工具箱
  });
  window.addEventListener('pageshow', ev => { if (ev.persisted) closeToolbox(); });

  function flashEl(el) {
    el.classList.remove('is-back');
    void el.offsetWidth; // 重新开始动画
    el.classList.add('is-back');
    const stop = () => el.classList.remove('is-back');
    el.addEventListener('animationend', ev => { if (ev.animationName.startsWith('app-back-')) stop(); });
    setTimeout(stop, 3600); // 关了动画的浏览器也要把光圈收掉
  }

  function flashTile(id) {
    const apps = [...flowList.querySelectorAll(`.app[data-tool="${CSS.escape(id)}"]`)];
    if (apps.length) apps.forEach(flashEl);
    else flashEl(folderBtn);
  }

  renderFlows();

  const params = new URLSearchParams(location.search);
  const from = params.get('from');
  if (from) {
    params.delete('from');
    const rest = params.toString();
    history.replaceState(null, '', location.pathname + (rest ? `?${rest}` : '')); // 刷新主页不会再闪
    if (T.byId[from]) {
      try {
        const y = Number(sessionStorage.getItem(SCROLL_KEY));
        if (y > 0) { history.scrollRestoration = 'manual'; window.scrollTo({ top: y, behavior: 'instant' }); }
      } catch (_) {}
      flashTile(from);
    }
  }

  window.HomeTiles = { flash: flashTile, openToolbox, closeToolbox, flows: () => flows, render: renderFlows };
})();
