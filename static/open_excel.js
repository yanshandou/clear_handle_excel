// 网页上的 Excel 文件名都可以点：在可见 Excel 里打开（知道工作表 / 单元格时直接跳过去）
// 用 excelLinkHtml() 生成链接；点击、提示都在这里统一处理，提示浮在页面上，不挪动布局
(function () {
  const opening = new Set(); // 列表每次刷新都会重建元素，不能靠元素上的 class 防连点

  // 点过的链接（结果弹窗里标「✓ 已打开」，最后点的那个加框）：弹窗每次都重建，按文件 / 表 / 格记住；
  // 存 sessionStorage，刷新页面也还在，关掉浏览器就清空
  const VISITED_KEY = 'clear_handle_excel.openedLinks';
  const linkKey = t => [t.folder, t.file, t.sheet, t.cell, ...(t.compare ? [t.compare] : [])].map(v => v || '').join('\0');
  let visited = new Set();
  let lastKey = '';
  try {
    const saved = JSON.parse(sessionStorage.getItem(VISITED_KEY) || '{}');
    visited = new Set(saved.keys || []);
    lastKey = saved.last || '';
  } catch (_) {}
  function markVisited(link) {
    const key = linkKey(link.dataset);
    visited.add(key);
    lastKey = key;
    ++window.excelLinkHtml.revision;
    document.querySelectorAll('.excel-link.is-last').forEach(el => el.classList.remove('is-last'));
    document.querySelectorAll('.excel-link').forEach(el => {
      if (linkKey(el.dataset) === key) el.classList.add('is-visited', 'is-last');
    });
    try { sessionStorage.setItem(VISITED_KEY, JSON.stringify({ keys: [...visited].slice(-500), last: lastKey })); } catch (_) {}
  }

  function escapeAttr(value) {
    return String(value ?? '').replace(/[&<>"']/g, ch => (
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  }

  // labelHtml 需已转义；compare: links / colors / hidden，按服务端配对同时打开当前与备份。
  // blocked：后台还要处理这份文件时不能打开的原因，链接变灰、鼠标移上去显示原因，点了只提示不打开
  window.excelLinkHtml = function (labelHtml, target, extraClass = '') {
    const data = ['file', 'sheet', 'cell', 'folder', 'blocked', 'compare']
      .filter(key => target[key])
      .map(key => ` data-${key}="${escapeAttr(target[key])}"`).join('');
    const where = target.sheet ? `并跳到 ${target.sheet}${target.cell ? `!${target.cell}` : ''}` : '';
    const title = target.blocked ? `暂时不能打开：${target.blocked}` : target.compare
      ? `同时打开当前文件和参考文件（只读），左右对比${where}` : `点击在 Excel 中打开${where}`;
    const key = linkKey(target);
    const seen = visited.has(key) ? `is-visited ${key === lastKey ? 'is-last ' : ''}` : '';
    return `<a href="#" class="excel-link ${target.blocked ? 'is-blocked ' : ''}${seen}${extraClass}"${data}`
      + `${target.record ? ' data-record="1"' : ''}${target.blocked ? ' aria-disabled="true"' : ''}`
      + ` title="${escapeAttr(title)}">${labelHtml}</a>`;
  };
  // 工作台缓存链接 HTML 时，也要跟随访问标记更新（含暂时被筛选隐藏的卡片）。
  window.excelLinkHtml.revision = 0;

  // 提示条：彩色块 + 断线框 + 右边线条图标。颜色按种类：正在… 蓝、成功 绿、有附注 橙、打不开/出错 红
  const TOAST_ICONS = {
    busy: '<path d="M20 12a8 8 0 1 1-2.34-5.66"/><path d="M20 4v4.5h-4.5"/>',
    ok: '<circle cx="12" cy="12" r="9"/><path d="M7.8 12.3l2.8 2.8 5.6-5.8"/>',
    warn: '<path d="M12 3.5l9.5 16.5h-19z"/><path d="M12 10v4.5"/><path d="M12 17.3v.2"/>',
    error: '<circle cx="12" cy="12" r="9"/><path d="M8.8 8.8l6.4 6.4M15.2 8.8l-6.4 6.4"/>',
  };
  function toastKind(text, isError) {
    if (isError) return /^(无法|暂时不能|没能|请)/.test(text) ? 'error' : 'warn';
    return /^正在/.test(text) ? 'busy' : 'ok';
  }
  // 线框按像素断开几处：长线、空、短线、空……（rect 的周长从左上角顺时针量）
  const FRAME_DASH = '150 9 7 9 110 12 60 10 9 10 180 8 40 12';

  let toastTimer = null;
  function toast(text, isError) {
    let box = document.getElementById('excelOpenToast');
    if (!box) {
      box = document.createElement('div');
      box.id = 'excelOpenToast';
      box.className = 'excel-open-toast';
      box.innerHTML = '<svg class="toast-frame" aria-hidden="true"><rect x="0" y="0" width="100%" height="100%"'
        + ` stroke-dasharray="${FRAME_DASH}"/></svg>`
        + '<span class="toast-text"></span><span class="toast-icon" aria-hidden="true"></span>';
      document.body.appendChild(box);
    }
    // 弹窗（showModal）盖在页面最上层：提示要放进打开着的弹窗里才看得到
    const host = document.querySelector('dialog[open]') || document.body;
    if (box.parentElement !== host) host.appendChild(box);
    const kind = toastKind(text, isError);
    box.querySelector('.toast-text').textContent = text;
    // 图标和字一样：先画一层粗的深色线当描边，再在上面画细的白线
    box.querySelector('.toast-icon').innerHTML = `<svg viewBox="-2 -2 28 28"><g class="icon-edge">${TOAST_ICONS[kind]}</g><g class="icon-fill">${TOAST_ICONS[kind]}</g></svg>`;
    box.classList.remove('is-busy', 'is-ok', 'is-warn', 'is-error');
    box.classList.add(`is-${kind}`);
    // 每条新提示都重新淡入一次
    box.classList.remove('is-visible');
    void box.offsetWidth;
    box.classList.add('is-visible');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => box.classList.remove('is-visible'), isError ? 6000 : 2500);
  }

  window.excelToast = toast; // 别的页面自己打开 Excel 时也用同一个提示条

  async function openExcel(link) {
    const d = link.dataset;
    const key = linkKey(d);
    const compare = !!d.compare;
    if (opening.has(key)) return;
    opening.add(key);
    link.classList.add('opening');
    toast(compare ? '正在打开当前文件和参考文件…' : '正在用 Excel 打开…');
    try {
      const res = await fetch(compare ? '/api/excel/compare' : '/api/excel/open', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          file: d.file, sheet: d.sheet || '', cell: d.cell || '', folder: d.folder || '',
          ...(compare ? {kind: d.compare} : {record: d.record === '1'}),
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
      toast(data.note ? data.note : compare ? '已同时打开：左侧当前文件，右侧参考文件（只读）' : '已在 Excel 中打开', !!data.note);
    } catch (e) {
      toast(`无法打开：${e.message}`, true);
    } finally {
      link.classList.remove('opening');
      setTimeout(() => opening.delete(key), 1500);
    }
  }

  // 捕获阶段拦下：链接所在卡片的点击（切换状态）和拖动（左右滑筛选）都不触发
  document.addEventListener('pointerdown', e => {
    if (e.target.closest && e.target.closest('.excel-link')) e.stopPropagation();
  }, true);
  document.addEventListener('click', e => {
    const link = e.target.closest && e.target.closest('.excel-link');
    if (!link) return;
    e.preventDefault();
    e.stopPropagation();
    if (link.dataset.blocked) toast(`暂时不能打开：${link.dataset.blocked}`, true);
    else {
      markVisited(link);
      openExcel(link);
    }
  }, true);
})();
