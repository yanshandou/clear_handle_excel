// 一页对比（/excel-diff/all，用户 2026-11-12）里，每份文件的对比就是嵌进去的这个页面（带 embed=1）。
// 这里只做三件事：藏掉页头和输入框、把高度告诉外面、对比完 / 出错 / 上传成功时告诉外面。不带 embed=1 时什么都不做。
(function () {
  const params = new URLSearchParams(location.search);
  if (params.get('embed') !== '1' || window.parent === window) return;
  document.documentElement.classList.add('xd-embed');
  const rel = params.get('rel') || '';
  const tell = (event, extra = {}) => window.parent.postMessage({ type: 'excel-diff', rel, event, ...extra }, location.origin);

  window.addEventListener('xd:done', event => {
    const data = event.detail || {};
    tell('done', { identical: !!data.identical, summary: data.summary || {} });
  });
  const watch = (id, test, event) => {
    const el = document.getElementById(id);
    if (!el) return;
    new MutationObserver(() => { const text = el.textContent.trim(); if (test(text)) tell(event, { text }); })
      .observe(el, { childList: true, characterData: true, subtree: true });
  };
  watch('xdError', text => !!text, 'error');
  watch('xdUploadStatus', text => text.startsWith('✓'), 'uploaded');

  let last = 0;
  const size = () => {
    const height = Math.ceil(document.documentElement.scrollHeight);
    if (Math.abs(height - last) > 2) { last = height; tell('height', { height }); }
  };
  new ResizeObserver(size).observe(document.body);
  window.addEventListener('load', size);
})();
