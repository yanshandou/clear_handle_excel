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

  // labelHtml 需已转义；compare: links / colors，按服务端配对同时打开当前与备份。
  // blocked：后台还要处理这份文件时不能打开的原因，链接变灰、鼠标移上去显示原因，点了只提示不打开
  window.excelLinkHtml = function (labelHtml, target, extraClass = '') {
    const data = ['file', 'sheet', 'cell', 'folder', 'blocked', 'compare']
      .filter(key => target[key])
      .map(key => ` data-${key}="${escapeAttr(target[key])}"`).join('');
    const where = target.sheet ? `并跳到 ${target.sheet}${target.cell ? `!${target.cell}` : ''}` : '';
    const title = target.blocked ? `暂时不能打开：${target.blocked}` : target.compare
      ? `同时打开当前文件和原始备份（只读），左右对比${where}` : `点击在 Excel 中打开${where}`;
    const key = linkKey(target);
    const seen = visited.has(key) ? `is-visited ${key === lastKey ? 'is-last ' : ''}` : '';
    return `<a href="#" class="excel-link ${target.blocked ? 'is-blocked ' : ''}${seen}${extraClass}"${data}`
      + `${target.record ? ' data-record="1"' : ''}${target.blocked ? ' aria-disabled="true"' : ''}`
      + ` title="${escapeAttr(title)}">${labelHtml}</a>`;
  };
  // 工作台缓存链接 HTML 时，也要跟随访问标记更新（含暂时被筛选隐藏的卡片）。
  window.excelLinkHtml.revision = 0;

  // 水墨背景：5 张长短不同的图（宽/高比），挑形状最接近的那张，再拉到刚好盖住文字
  const INK_RATIOS = [1.86, 2.3, 2.61, 3.16, 3.05];
  function fitInk(box) {
    const cs = getComputedStyle(box);
    const textW = box.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight);
    const textH = box.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
    // 背景只是装饰：不必完全包住字，大致托在字下面就行
    const h = Math.max(110, Math.min(textH * 2, textH + 60));
    const w = Math.max(textW * 1.15, h * INK_RATIOS[0]);
    const want = w / h;
    let best = 0;
    INK_RATIOS.forEach((r, i) => {
      if (Math.abs(Math.log(r / want)) < Math.abs(Math.log(INK_RATIOS[best] / want))) best = i;
    });
    box.style.setProperty('--ink-img', `url("/static/icons/ca/ink_${best + 1}.png")`);
    box.style.setProperty('--ink-w', `${Math.round(w)}px`);
    box.style.setProperty('--ink-h', `${Math.round(h)}px`);
  }

  let toastTimer = null;
  function toast(text, isError) {
    let box = document.getElementById('excelOpenToast');
    if (!box) {
      box = document.createElement('div');
      box.id = 'excelOpenToast';
      box.className = 'excel-open-toast';
      document.body.appendChild(box);
    }
    // 弹窗（showModal）盖在页面最上层：提示要放进打开着的弹窗里才看得到
    const host = document.querySelector('dialog[open]') || document.body;
    if (box.parentElement !== host) host.appendChild(box);
    box.textContent = text;
    box.classList.toggle('is-error', !!isError);
    // 每条新提示都重新从小变大淡入一次（像游戏里每次击杀都重播播报）
    box.classList.remove('is-visible');
    fitInk(box);
    box.classList.add('is-visible');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => box.classList.remove('is-visible'), isError ? 6000 : 2500);
  }

  async function openExcel(link) {
    const d = link.dataset;
    const key = linkKey(d);
    const compare = !!d.compare;
    if (opening.has(key)) return;
    opening.add(key);
    link.classList.add('opening');
    toast(compare ? '正在打开当前文件和备份…' : '正在用 Excel 打开…');
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
      toast(data.note ? data.note : compare ? '已同时打开：左侧当前文件，右侧原始备份（只读）' : '已在 Excel 中打开', !!data.note);
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
