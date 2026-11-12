// 一页对比（用户 2026-11-12）：当前任务要上传的文件都列在这一页，每份嵌一个单份对比页（/excel-diff?embed=1），
// 自动逐份对比并提交预览 PDF，不等上一份导完；有差异的结果按列表顺序展开。
(function () {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  let task = null, files = [];
  const cards = new Map();
  let queue = null;          // 尚未开始逐格检查的文件；PDF 由后台 Excel 另行排队
  let paused = false, loading = false;

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
    return { idle: '等待对比', running: '正在对比…', same: '✓ 未发现差异', diff: '有不同', err: '! 比不了', up: '✓ 已上传' }[card.state] || '';
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
    const pdf = s => all.filter(c => c.preview === s).length;
    const checked = all.length - n('idle') - n('running');
    const progress = [pdf('queued') ? `${pdf('queued')} 份正在准备` : '', pdf('done') ? `${pdf('done')} 份已导好` : '',
      pdf('failed') ? `${pdf('failed')} 份导出有错误` : ''].filter(Boolean);
    $('xaCount').textContent = `逐格对比 ${checked} / ${all.length} 份（未发现差异 ${n('same')}、有不同 ${n('diff')}、比不了 ${n('err')}、已上传 ${n('up')}）`
      + (progress.length ? ' · 预览 PDF：' + progress.join('，') : '');
  }

  function cardHtml(file) {
    const status = file.review ? (file.review.ready ? '待重新处理' : '同步未完成') : file.uploaded_current === true ? '当前版本已上传'
      : file.uploaded_current === false ? '本地有新版本待上传' : file.uploaded ? '当前版本待核对' : '待上传';
    return `<div class="xa-head"><span class="xa-name">${window.excelLinkHtml ? window.excelLinkHtml(esc(file.rel), { folder: task.work, file: file.rel }) : esc(file.rel)}<small>${esc(status)}</small></span>`
      + `<span class="xa-state">等待对比</span>`
      + `<button type="button" class="btn btn-outline btn-sm" data-act="diff">🔍 对比</button>`
      + `<button type="button" class="btn btn-primary btn-sm" data-act="upload" hidden disabled title="对比完才能点；和下面对比结果里的上传按钮是同一个">⬆ 上传这一份</button>`
      + `<button type="button" class="btn btn-outline btn-sm" data-act="pdf">📕 并排看 PDF</button>`
      + `<a class="btn btn-outline btn-sm" target="_blank" rel="noopener" href="${esc(diffUrl(file.rel))}">↗ 单独打开</a></div>`
      + `<div class="xa-body" hidden><div class="xa-pdfs" hidden></div></div>`;
  }

  function render() {
    for (const card of cards.values()) card.wait?.();
    cards.clear();
    const shown = $('xaShowAll').checked ? files : files.filter(needsLook);
    if (!shown.length) {
      $('xaList').innerHTML = `<p>${files.length ? '这个任务的文件当前版本都已经上传过了，没有要对比的。勾上面的「也列出当前版本已经上传过的」可以全部列出来。' : '处理文件夹里还没有 Excel 文件。'}</p>`;
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
      cards.set(file.rel, { el, state: 'idle', frame: null, open: false, wait: null, preview: null });
    }
    $('xaList').replaceChildren(list);
    counts();
  }

  function body(card) { return card.el.querySelector('.xa-body'); }
  function syncOpen(card) {
    const open = card.open || !card.el.querySelector('.xa-pdfs').hidden;
    body(card).hidden = !open;
    if (card.frame) card.frame.hidden = !card.open;
    card.el.classList.toggle('is-open', open);
    const button = card.el.querySelector('[data-act="diff"]');
    button.textContent = card.state === 'err' ? '重新对比' : card.open ? '收起对比' : card.frame ? '展开对比' : '🔍 对比';
    button.setAttribute('aria-expanded', String(card.open));
    button.disabled = !card.frame && !!queue;
  }

  function syncControls() {
    $('xaShowAll').disabled = $('xaReload').disabled = loading || !!queue;
    $('xaRunAll').hidden = !!queue || ![...cards.values()].some(c => c.state === 'idle');
    $('xaRunAll').disabled = loading || !!queue;
    $('xaPause').hidden = !queue;
    for (const card of cards.values()) syncOpen(card);
  }

  // 等这份逐格检查并提交 PDF 请求后，就开始下一份；不等 PDF 导出。
  function compare(rel) {
    const card = cards.get(rel);
    const ready = new Promise(resolve => { card.wait = () => { card.wait = null; resolve(); }; });
    const frame = document.createElement('iframe');
    frame.className = 'xa-diff';
    frame.title = '对比：' + rel;
    frame.src = diffUrl(rel) + '&embed=1';
    card.frame?.remove();
    card.frame = frame;
    card.open = true;
    card.preview = null;
    setState(rel, 'running');
    syncOpen(card);
    body(card).appendChild(frame);
    console.info('🔍 一页对比：开始', { 文件: rel });
    return ready;
  }

  window.addEventListener('message', event => {
    if (event.origin !== location.origin || !event.data || event.data.type !== 'excel-diff') return;
    const card = [...cards.values()].find(c => c.frame && c.frame.contentWindow === event.source);
    if (!card) return;
    const rel = card.el.dataset.rel, msg = event.data;
    if (msg.event === 'height' && Number.isFinite(msg.height)) card.frame.style.height = `${Math.max(msg.height, 120)}px`;
    else if (msg.event === 'done') {
      const s = msg.summary || {};
      const note = msg.identical ? '' : [s['有不同的表'] ? `${s['有不同的表']} 张表有改动` : '', s['只有远程有'] ? `${s['只有远程有']} 张只有远程有` : '',
        s['只有你这份有'] ? `${s['只有你这份有']} 张只有你这份有` : ''].filter(Boolean).join('，');
      if (card.state !== 'up') setState(rel, msg.identical ? 'same' : 'diff', note);
      if (msg.identical) {
        card.open = false;
        card.wait?.();
      }
      syncOpen(card);
      console.info('🔍 一页对比：比完', { 文件: rel, 完全一样: msg.identical });
    } else if (msg.event === 'preview') {
      card.preview = msg.state;
      counts();
      card.wait?.();
    } else if (msg.event === 'error') {
      setState(rel, 'err', msg.text || '');
      syncOpen(card);
      console.warn('🔍 一页对比：比不了', { 文件: rel, 原因: msg.text });
      card.wait?.();
    } else if (msg.event === 'upload-state') {
      // 卡片标题吸顶，滑到哪里都能点上传（用户 2026-11-12）
      const up = card.el.querySelector('[data-act="upload"]');
      up.hidden = !msg.shown;
      up.disabled = !msg.ready;
      up.textContent = msg.ready ? msg.label || '⬆ 上传这一份' : msg.status && !msg.status.startsWith('✓') ? '上传中…' : msg.status.startsWith('✓') ? '✓ 已上传' : '⬆ 上传这一份';
      up.title = msg.error || msg.status || '对比完才能点；和下面对比结果里的上传按钮是同一个';
        } else if (msg.event === 'uploaded') {
      setState(rel, 'up');
      console.info('🔍 一页对比：已上传', { 文件: rel });
    }
  });

  // 卡片标题吸在页头下面（页头是固定的，吸在 0 会被它挡住；用户 2026-11-12 要滑到哪都能点上传）
  function stickyTop() {
    const header = document.querySelector('.app-header');
    const fixed = header && ['fixed', 'sticky'].includes(getComputedStyle(header).position);
    document.documentElement.style.setProperty('--xa-sticky-top', `${fixed ? Math.ceil(header.getBoundingClientRect().height) : 0}px`);
  }
  stickyTop();
  window.addEventListener('resize', stickyTop);

  async function runAll() {
    if (queue || loading) return;
    paused = false;
    queue = [...cards.entries()].filter(([, c]) => c.state === 'idle').map(([rel]) => rel);
    if (!queue.length) { queue = null; syncControls(); return; }
    $('xaPause').disabled = false;
    $('xaPause').textContent = '⏸ 暂停';
    syncControls();
    console.info('🔍 一页对比：全部对比', { 份数: queue.length });
    while (queue.length && !paused) {
      const rel = queue.shift();
      if (!cards.has(rel) || cards.get(rel).state !== 'idle') continue;
      await compare(rel);
    }
    queue = null;
    syncControls();
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
    if (button.dataset.act === 'upload') {
      const card = cards.get(rel);
      card?.frame?.contentWindow?.postMessage({ type: 'excel-diff-cmd', cmd: 'upload' }, location.origin);
      return;
    }
    const card = cards.get(rel);
    if (!card.frame || card.state === 'err') { compare(rel); syncControls(); return; }
    card.open = !card.open;
    syncOpen(card);
  });
  $('xaRunAll').addEventListener('click', runAll);
  $('xaPause').addEventListener('click', () => {
    paused = true;
    $('xaPause').disabled = true;
    $('xaPause').textContent = '等这一份提交后暂停…';
  });
  $('xaShowAll').addEventListener('change', () => { render(); runAll(); });
  $('xaReload').addEventListener('click', loadFiles);

  async function loadFiles() {
    if (queue || loading) return;
    loading = true;
    syncControls();
    $('xaError').textContent = '';
    cards.clear();
    counts();
    $('xaList').innerHTML = '<p>正在读取文件…</p>';
    try {
      files = (await post('/api/distribute/task_files', { task_id: task.id })).files || [];
      render();
    } catch (error) {
      files = [];
      $('xaError').textContent = error.message;
      $('xaList').innerHTML = '';
    } finally {
      loading = false;
      syncControls();
    }
    runAll();
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
  }
  start();
})();
