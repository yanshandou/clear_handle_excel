// 主页卡片页（用户 2026-11-11 要，布局参考开始菜单的大小磁贴）：流程、主要卡片、更多工具并成一张卡片页。
// - 大卡：精简 Sheet（顺便做 BL / 引用 / 隐藏页 / 字体）、上传 Excel（上传前清标记、恢复颜色、标准化）——同一份 Excel 只开存一次
// - 每一步又各有一张小卡，可以单独用；被大卡顺便做的小卡上写「已含在…里」
// - 卡片位置写死不排序，操作时不挪
// - 从某一页点「返回首页」回来（/?from=工具）：那张卡片边框闪几下后消失，页面不滚动、卡片不动
// 工具清单在 static/tools.js
(function () {
  const T = window.Tools;
  const grid = document.getElementById('tileGrid');
  if (!T || !grid) return;
  const esc = T.esc;

  function extrasHtml(t) {
    if (t.id === 'clean') {
      const on = T.chainOptions();
      return `<div class="tile-extras" aria-label="精简后顺便做"><b>顺便做</b>${T.CHAIN.map(c =>
        `<label><input type="checkbox" data-chain="${c.key}"${on[c.key] ? ' checked' : ''}> ${esc(c.label)}</label>`).join('')}</div>`;
    }
    if (t.id === 'upload') {
      return '<div class="tile-extras is-fixed" title="上传前每份都会做，不用单独点"><b>上传前都会做</b><span>✓ 清标记</span><span>✓ 恢复颜色</span><span>✓ 标准化</span></div>';
    }
    return '';
  }

  function tileHtml(t) {
    const icon = t.sprite
      ? `<span class="tile-icon garden-sprite sprite-${t.sprite}" aria-hidden="true"></span>`
      : `<span class="tile-icon tool-vector" aria-hidden="true"><svg class="ui-icon"><use href="/static/icons/garden-ui.svg#${t.vector}"></use></svg></span>`;
    const attrs = t.action ? `type="button" onclick="${t.action}"` : `href="${esc(t.url)}"${t.linkId ? ` id="${t.linkId}"` : ''}`;
    const tag = t.action ? 'button' : 'a';
    return `<li class="tile tile-${t.size} tone-${t.tone}" data-tool="${t.id}"${t.ops ? ` data-ops="${t.ops}"` : ''}>
      <${tag} ${attrs} class="tile-link">
        ${icon}
        <span class="tile-text">
          <span class="tile-title">${esc(t.title)}</span>
          <span class="tile-sub">${esc(t.sub)}</span>
          ${t.meta ? `<span class="tile-meta"${t.metaId ? ` id="${t.metaId}"` : ''}>${esc(t.meta)}</span>` : ''}
        </span>
      </${tag}>
      ${t.inc ? '<span class="tile-inc"></span>' : ''}
      ${extrasHtml(t)}
      <span class="tile-done hidden" role="img"></span>
    </li>`;
  }

  grid.innerHTML = T.TOOLS.map(tileHtml).join('');

  // 小卡上的「已含在…里 / 单独用」：只换文字，卡片不动
  function paintInc() {
    const chain = T.chainOptions();
    grid.querySelectorAll('.tile').forEach(li => {
      const t = T.byId[li.dataset.tool];
      const tag = li.querySelector('.tile-inc');
      if (!tag) return;
      const included = !t.incKey || chain[t.incKey];
      tag.textContent = included ? `${t.inc === 'clean' ? '精简' : '上传'}时顺便做` : '单独用';
      tag.classList.toggle('is-off', !included);
      tag.title = included ? `做「${T.byId[t.inc].title}」时会顺便做这一步，也可以点这张卡单独做` : `「${T.byId[t.inc].title}」卡上没勾这一项，要做就点这张卡单独做`;
    });
  }

  grid.addEventListener('change', e => {
    const box = e.target.closest('input[data-chain]');
    if (!box) return;
    T.setChainOption(box.dataset.chain, box.checked);
    paintInc();
  });
  paintInc();

  // ---------- 从哪一页回来的：那张卡片边框闪几下 ----------
  const SCROLL_KEY = 'clear_handle_excel.homeScroll';
  // 离开主页时记下滚到哪里，回来时原样停在那里（不是滚动动画，页面一出来就在原位）
  window.addEventListener('pagehide', () => { try { sessionStorage.setItem(SCROLL_KEY, String(window.scrollY)); } catch (_) {} });

  function flashTile(id) {
    const li = grid.querySelector(`.tile[data-tool="${CSS.escape(id)}"]`);
    if (!li) return;
    li.classList.remove('is-back');
    void li.offsetWidth; // 重新开始动画
    li.classList.add('is-back');
    const stop = () => li.classList.remove('is-back');
    li.addEventListener('animationend', e => { if (e.animationName.startsWith('tile-back-')) stop(); });
    setTimeout(stop, 3600); // 关了动画的浏览器也要把边框收掉
  }

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

  window.HomeTiles = { flash: flashTile, paintInc };
})();
