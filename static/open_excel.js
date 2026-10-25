// 网页上的 Excel 文件名都可以点：在可见 Excel 里打开（知道工作表 / 单元格时直接跳过去）
// 用 excelLinkHtml() 生成链接；点击、提示都在这里统一处理，提示浮在页面上，不挪动布局
(function () {
  const opening = new Set(); // 列表每次刷新都会重建元素，不能靠元素上的 class 防连点

  function escapeAttr(value) {
    return String(value ?? '').replace(/[&<>"']/g, ch => (
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  }

  // labelHtml 需已转义；target: { file, sheet, cell, folder, record }
  window.excelLinkHtml = function (labelHtml, target, extraClass = '') {
    const data = ['file', 'sheet', 'cell', 'folder']
      .filter(key => target[key])
      .map(key => ` data-${key}="${escapeAttr(target[key])}"`).join('');
    const where = target.sheet ? `并跳到 ${target.sheet}${target.cell ? `!${target.cell}` : ''}` : '';
    return `<a href="#" class="excel-link ${extraClass}"${data}${target.record ? ' data-record="1"' : ''}`
      + ` title="${escapeAttr(`点击在 Excel 中打开${where}`)}">${labelHtml}</a>`;
  };

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
    box.classList.add('is-visible');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => box.classList.remove('is-visible'), isError ? 6000 : 2500);
  }

  async function openExcel(link) {
    const d = link.dataset;
    const key = `${d.file}\0${d.sheet || ''}\0${d.cell || ''}`;
    if (opening.has(key)) return;
    opening.add(key);
    link.classList.add('opening');
    toast('正在用 Excel 打开…');
    try {
      const res = await fetch('/api/excel/open', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          file: d.file, sheet: d.sheet || '', cell: d.cell || '', folder: d.folder || '', record: d.record === '1',
        }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
      toast(data.note ? `已打开，但${data.note}` : '已在 Excel 中打开', !!data.note);
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
    openExcel(link);
  }, true);
})();
