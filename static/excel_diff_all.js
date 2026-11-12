// 一页对比（用户 2026-11-12）：当前任务要上传的文件都列在这一页，每份嵌一个单份对比页（/excel-diff?embed=1），
// 结果和「⬆ 没问题，上传这一份」都沿用单份页；这里只管列文件、排队一份一份比、收起、并排看已有的同名 PDF。
(function () {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  let task = null, files = [];
  const cards = new Map();   // rel → { el, state, frame, wait }
  let queue = null;          // 「全部对比」排着的 rel；null = 没在跑
  let paused = false;

  function joinPath(root, rel) {
    const sep = root.includes('\\') || /^[A-Za-z]:/.test(root) ? '\\' : '/';
    return root.replace(/[\\/]+$/, '') + sep + rel.split('/').join(sep);
  }
  async function post(path, body) {
    const response = await fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const data = await response.json();
    if (!response.ok || data.ok === false) throw new Error(data.error || '请求失败');
    return data;
  }
  // 和上传页清单一样：当前版本已经传过、也没有待重新处理的，就不用再看
  const needsLook = file => !(file.uploaded_current === true && !file.review);

  function diffUrl(rel) {
    return `/excel-diff?theirs=${encodeURIComponent(joinPath(task.shared, rel))}&mine=${encodeURIComponent(joinPath(task.work, rel))}`
      + `&task=${encodeURIComponent(task.id)}&rel=${encodeURIComponent(rel)}`;
  }

  function stateText(card) {
    return { idle: '还没比', running: '正在比…', same: '✓ 未发现差异', diff: '有不同', err: '! 比不了', up: '✓ 已上传' }[card.state] || '';
  }
  function setState(rel, state, note = '') {
    const card = cards.get(rel);
    if (!card) return;
    card.state = state;
    card.el.classList.remove('is-same', 'is-diff', 'is-err', 'is-up');
    if (state !== 'idle' && state !== 'running') card.el.classList.add('is-' + state);
    card.el.querySelector('.xa-state').textContent = stateText(card) + (note ? '：' + note : '');
    counts();
  }
  function counts() {
    const all = [...cards.values()];
    const n = s => all.filter(c => c.state === s).length;
    $('xaCount').textContent = `共 ${all.length} 份 · 比完 ${n('same') + n('diff') + n('err') + n('up')}（未发现差异 ${n('same')}、有不同 ${n('diff')}、比不了 ${n('err')}、已上传 ${n('up')}）`;
  }

  function cardHtml(file) {
    const status = file.review ? (file.review.ready ? '待重新处理' : '同步未完成') : file.uploaded_current === true ? '当前版本已上传'
      : file.uploaded_current === false ? '本地有新版本待上传' : file.uploaded ? '当前版本待核对' : '待上传';
    return `<div class="xa-head"><span class="xa-name">${window.excelLinkHtml ? window.excelLinkHtml(esc(file.rel), { folder: task.work, file: file.rel }) : esc(file.rel)}<small>${esc(status)}</small></span>`
      + `<span class="xa-state">还没比</span>`
      + `<button type="button" class="btn btn-outline btn-sm" data-act="diff">🔍 对比</button>`
      + `<button type="button" class="btn btn-outline btn-sm" data-act="pdf">📕 并排看 PDF</button>`
      + `<a class="btn btn-outline btn-sm" target="_blank" rel="noopener" href="${esc(diffUrl(file.rel))}">↗ 单独打开</a></div>`
      + `<div class="xa-body" hidden><div class="xa-pdfs" hidden></div></div>`;
  }

  function render() {
    for (const card of cards.values()) card.wait?.(null);
    cards.clear();
    const shown = $('xaShowAll').checked ? files : files.filter(needsLook);
    if (!shown.length) {
      $('xaList').innerHTML = `<p>${files.length ? '这个任务的文件当前版本都已经上传过了，没有要对比的。勾上面的「也列出当前版本已经上传过的」可以全部列出来。' : '处理文件夹里还没有 Excel 文件。'}</p>`;
      $('xaRunAll').disabled = true;
      counts();
      return;
    }
    const list = document.createElement('div');
    list.className = 'xa-list';
    for (const file of shown) {
      const el = document.createElement('article');
      el.className = 'xa-card';
      el.dataset.rel = file.rel;
      el.innerHTML = cardHtml(file);
      list.appendChild(el);
      cards.set(file.rel, { el, state: 'idle', frame: null, wait: null });
    }
    $('xaList').replaceChildren(list);
    $('xaRunAll').disabled = false;
    counts();
  }

  function body(card) { return card.el.querySelector('.xa-body'); }
  function syncOpen(card) {
    const open = !!card.frame || !card.el.querySelector('.xa-pdfs').hidden;
    body(card).hidden = !open;
    card.el.classList.toggle('is-open', open);
    card.el.querySelector('[data-act="diff"]').textContent = card.frame ? '收起对比' : '🔍 对比';
  }

  // 比一份：嵌进单份对比页，等它说比完 / 出错。返回 Promise，结果是 done / error / null（被收起）
  function compare(rel) {
    const card = cards.get(rel);
    if (!card) return Promise.resolve(null);
    if (card.frame) return Promise.resolve('done');
    const frame = document.createElement('iframe');
    frame.className = 'xa-diff';
    frame.title = '对比：' + rel;
    frame.src = diffUrl(rel) + '&embed=1';
    body(card).appendChild(frame);
    card.frame = frame;
    syncOpen(card);
    setState(rel, 'running');
    console.info('🔍 一页对比：开始', { 文件: rel });
    return new Promise(resolve => { card.wait = result => { card.wait = null; resolve(result); }; });
  }
  function collapse(rel) {
    const card = cards.get(rel);
    if (!card || !card.frame) return;
    card.frame.remove();
    card.frame = null;
    if (card.state === 'running') setState(rel, 'idle');
    card.wait?.(null);
    syncOpen(card);
  }

  window.addEventListener('message', event => {
    if (event.origin !== location.origin || !event.data || event.data.type !== 'excel-diff') return;
    const card = [...cards.values()].find(c => c.frame && c.frame.contentWindow === event.source);
    if (!card) return;
    const rel = card.el.dataset.rel, msg = event.data;
    if (msg.event === 'height') card.frame.style.height = `${Math.min(Math.max(msg.height, 120), 20000)}px`;
    else if (msg.event === 'done') {
      const s = msg.summary || {};
      const note = msg.identical ? '' : [s['有不同的表'] ? `${s['有不同的表']} 张表有改动` : '', s['只有远程有'] ? `${s['只有远程有']} 张只有远程有` : '',
        s['只有你这份有'] ? `${s['只有你这份有']} 张只有你这份有` : ''].filter(Boolean).join('，');
      if (card.state !== 'up') setState(rel, msg.identical ? 'same' : 'diff', note);
      console.info('🔍 一页对比：比完', { 文件: rel, 完全一样: msg.identical });
      card.wait?.('done');
    } else if (msg.event === 'error') {
      setState(rel, 'err', msg.text || '');
      console.warn('🔍 一页对比：比不了', { 文件: rel, 原因: msg.text });
      card.wait?.('error');
    } else if (msg.event === 'uploaded') {
      setState(rel, 'up');
      console.info('🔍 一页对比：已上传', { 文件: rel });
    }
  });

  async function runAll() {
    if (queue) return;
    paused = false;
    queue = [...cards.entries()].filter(([, c]) => c.state === 'idle').map(([rel]) => rel);
    $('xaRunAll').disabled = true;
    $('xaPause').hidden = false;
    $('xaPause').textContent = '⏸ 暂停';
    console.info('🔍 一页对比：全部对比', { 份数: queue.length });
    while (queue.length && !paused) {
      const rel = queue.shift();
      if (!cards.has(rel) || cards.get(rel).state !== 'idle') continue;
      const result = await compare(rel);
      if (result === 'done' && cards.get(rel)?.state === 'same') collapse(rel);  // 完全一样的收起来，状态留着
    }
    const stopped = paused && queue.length;
    queue = null;
    $('xaPause').hidden = true;
    $('xaRunAll').disabled = !cards.size;
    $('xaRunAll').textContent = stopped ? '▶ 接着比剩下的' : '▶ 全部对比';
    if (!stopped) window.gameFx?.celebrate?.();
  }

  async function togglePdf(rel) {
    const card = cards.get(rel);
    const box = card.el.querySelector('.xa-pdfs');
    if (!box.hidden) { box.hidden = true; box.replaceChildren(); syncOpen(card); return; }
    box.hidden = false;
    box.innerHTML = '<p>正在看两边有没有同名 PDF…</p>';
    syncOpen(card);
    try {
      const info = await post('/api/distribute/diff_pdfs', { task_id: task.id, rel });
      const side = (key, label, cls) => {
        const p = info[key];
        const url = `/api/distribute/diff_pdf?task=${encodeURIComponent(task.id)}&rel=${encodeURIComponent(rel)}&side=${key}`;
        return `<div class="xa-pdf ${cls}"><strong>${label}</strong><small>${esc(p.pdf)}${p.pdf_at ? ' · ' + esc(p.pdf_at) : ''}</small>`
          + (p.exists ? (p.stale ? '<span class="xa-warn">! 这份 PDF 比旁边的 Excel 旧，印出来的可能不是这一版</span>' : '')
            + `<iframe title="${esc(label)}" src="${url}#view=FitH"></iframe>`
            : `<div class="xa-none">没有同名 PDF。要看印出来的样子，先到 <a href="/pdf-bundle" target="_blank" rel="noopener">批量导出同名 PDF</a> 导一下。</div>`) + '</div>';
      };
      box.innerHTML = side('theirs', '− 远程现在这份（旧）', 'is-old') + side('mine', '+ 你要上传的这份（新）', 'is-new');
      console.info('📕 一页对比：并排看 PDF', { 文件: rel, 远程有: info.theirs.exists, 你这份有: info.mine.exists });
    } catch (error) {
      box.innerHTML = `<p class="batch-error">${esc(error.message)}</p>`;
    }
  }

  $('xaList').addEventListener('click', event => {
    const button = event.target.closest('button[data-act]');
    if (!button) return;
    const rel = button.closest('.xa-card').dataset.rel;
    if (button.dataset.act === 'pdf') { togglePdf(rel); return; }
    const card = cards.get(rel);
    if (card.frame) collapse(rel);
    else {
      if (card.state !== 'up') card.state = 'idle';
      compare(rel);
    }
  });
  $('xaRunAll').addEventListener('click', runAll);
  $('xaPause').addEventListener('click', () => { paused = true; $('xaPause').textContent = '等这一份比完就停…'; });
  $('xaShowAll').addEventListener('change', () => { if (!queue) render(); else $('xaShowAll').checked = !$('xaShowAll').checked; });
  $('xaReload').addEventListener('click', () => { if (!queue) loadFiles(); });

  async function loadFiles() {
    $('xaError').textContent = '';
    $('xaList').innerHTML = '<p>正在读取文件…</p>';
    try {
      files = (await post('/api/distribute/task_files', { task_id: task.id })).files || [];
      render();
    } catch (error) {
      $('xaError').textContent = error.message;
      $('xaList').innerHTML = '';
    }
  }

  async function start() {
    try {
      const result = await (await fetch('/api/tasks', { cache: 'no-store' })).json();
      task = (result.tasks || []).find(t => t.id === result.current) || null;
    } catch (_) { task = null; }
    if (!task) {
      $('xaTask').innerHTML = '还没有当前任务。请先到 <a href="/distribute/start">新任务 · 拿文件</a> 开一个。';
      $('xaList').innerHTML = '';
      return;
    }
    $('xaTitle').textContent = '当前任务：' + task.name;
    $('xaTask').textContent = `左 = 共享 ${task.shared}　右 = 处理文件夹 ${task.work}`;
    await loadFiles();
    if (new URLSearchParams(location.search).get('run') === '1' && cards.size) runAll();
  }
  start();
})();
