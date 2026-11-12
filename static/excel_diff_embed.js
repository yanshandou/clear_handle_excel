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

  // 外面卡片吸顶标题上的「⬆ 上传这一份」（用户 2026-11-12）：这里告诉外面能不能点、正在干什么；外面点了就替它点里面的按钮
  const panel = document.getElementById('xdUploadPanel'), button = document.getElementById('xdUpload');
  const status = document.getElementById('xdUploadStatus'), failed = document.getElementById('xdUploadError');
  let lastUpload = '';
  const uploadState = () => {
    if (!panel || !button) return;
    const state = { ready: !panel.hidden && !button.disabled, shown: !panel.hidden, label: button.textContent.trim(),
                    status: (status?.textContent || '').trim(), error: (failed?.textContent || '').trim() };
    const key = JSON.stringify(state);
    if (key !== lastUpload) { lastUpload = key; tell('upload-state', state); }
  };
  for (const el of [panel, button, status, failed]) {
    if (el) new MutationObserver(uploadState).observe(el, { attributes: true, childList: true, characterData: true, subtree: true });
  }
  window.addEventListener('message', event => {
    if (event.origin !== location.origin || event.source !== window.parent || event.data?.type !== 'excel-diff-cmd') return;
    if (event.data.cmd === 'upload' && panel && !panel.hidden && button && !button.disabled) {
      console.info('📤 一页对比：从卡片标题点了上传这一份', { 文件: rel });
      button.click();
    }
  });
  uploadState();

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
