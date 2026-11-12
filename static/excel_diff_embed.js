// 一页对比（/excel-diff/all，用户 2026-11-12）里，每份文件的对比就是嵌进去的这个页面（带 embed=1）。
// 藏掉页头和输入框，把内容高度、对比和 PDF 进度、上传结果告诉外面。
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
  window.addEventListener('xd:preview', event => tell('preview', { state: event.detail }));
  const watch = (id, test, event) => {
    const el = document.getElementById(id);
    if (!el) return;
    new MutationObserver(() => { const text = el.textContent.trim(); if (test(text)) tell(event, { text }); })
      .observe(el, { childList: true, characterData: true, subtree: true });
  };
  watch('xdError', text => !!text, 'error');
  watch('xdUploadStatus', text => text.startsWith('✓'), 'uploaded');

  let last = 0;
  const main = document.querySelector('main');
  const size = () => {
    // scrollHeight 至少有 iframe 视口那么高，收起内容时它不会缩小。量实际内容高度。
    const height = Math.ceil(main.getBoundingClientRect().bottom + window.scrollY);
    if (height > 0 && height !== last) { last = height; tell('height', { height }); }
  };
  new ResizeObserver(size).observe(main);
  window.addEventListener('load', size);
})();
