// 收尾页 /finish：完成任务前逐项检查，显示进度，过了就打勾（后端 core/work_task.py start_finish / finish_status）
// 用户 2026-11-06 要：入口（新任务）有页面，出口（完成任务）也要有页面，让用户知道在检查什么、到哪了。
(() => {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
  // 6 个检查项：第 1 项列文件，2、3 项列完就知道，4～6 项要逐个比对内容
  const STEPS = [
    { key: 'list', title: '读取处理文件夹和共享文件夹' },
    { key: '共享里没有', title: '处理里的文件，共享里都有', fix: '处理文件夹里有、共享里没有：新文件还没传。请到 <a href="/distribute">第 4 步上传</a>。' },
    { key: '待重新处理', title: '没有待重新处理的文件', fix: '上传时发现共享有更新，同步下来后还没重新处理。处理、复检后到 <a href="/distribute">第 4 步</a> 再上传。' },
    { key: '没上传', title: '改过的文件都已传回共享', fix: '改过但还没传回共享。请到 <a href="/distribute">第 4 步上传</a>。' },
    { key: '别人改过', title: '共享里的文件没被别人改过', fix: '你拿走之后，别人在共享里改过。先到 <a href="/distribute/start">新任务 · 拿文件</a> 拿最新版，改好再传。' },
    { key: '读不到', title: '每个文件都读得到', fix: '读不到的文件可能正在 Excel 里打开，关掉后再检查一遍。' },
  ];
  const EARLY = new Set(['共享里没有', '待重新处理']);
  const STATE_TEXT = { ok: '✓ 已传回共享，和共享一致', wait: '还没比对', '共享里没有': '共享里没有', '待重新处理': '待重新处理',
    '没上传': '改过还没上传', '别人改过': '别人在共享里改过', '读不到': '读不到' };
  let task = null, status = {}, timer = 0, filter = '', celebrated = false;

  function stepState(step) {
    const result = status.result, running = status.running, phase = status.phase || '';
    const issues = result?.issues || status.early || null;
    if (step.key === 'list') {
      if (result || ['listed', 'compare'].includes(phase)) {
        const w = result ? result.files : status.work_files, s = result ? result.shared_files : status.shared_files;
        return { cls: 'ok', icon: '✓', note: `处理 ${w ?? '—'} 个、共享 ${s ?? '—'} 个 Excel` };
      }
      return running ? { cls: 'doing', icon: '', note: '正在读取…' } : { cls: 'wait', icon: '○', note: '' };
    }
    const known = result || (EARLY.has(step.key) && status.early);
    if (known) {
      const list = issues[step.key] || [];
      return list.length ? { cls: 'bad', icon: '!', note: `${list.length} 个文件`, fix: step.fix }
        : { cls: 'ok', icon: '✓', note: '通过' };
    }
    if (running && phase === 'compare') return { cls: 'doing', icon: '', note: `正在比对 ${status.done || 0} / ${status.total || 0}` };
    return { cls: 'wait', icon: '○', note: running ? '等前面做完' : '' };
  }

  function renderSteps() {
    $('finSteps').innerHTML = STEPS.map((step, i) => {
      const s = stepState(step);
      return `<li class="fin-step is-${s.cls}"><span class="fin-step-icon" aria-hidden="true">${s.icon}</span>`
        + `<span class="fin-step-body"><span class="fin-step-title"><b>${i + 1}</b>${esc(step.title)}</span>`
        + `<span class="fin-step-note">${esc(s.note)}</span>${s.fix ? `<span class="fin-step-fix">${s.fix}</span>` : ''}</span></li>`;
    }).join('');
  }

  function renderProgress() {
    const fill = $('finProgressFill'), track = fill.parentElement;
    const running = !!status.running;
    track.classList.toggle('is-done', !running);
    let fraction = 0;
    if (running && status.phase === 'compare' && status.total) fraction = (status.done || 0) / status.total;
    fill.style.width = `${Math.round((running ? fraction : status.result ? 1 : 0) * 100)}%`;
    if (running) {
      $('finProgress').textContent = status.phase === 'compare'
        ? `正在逐个比对内容 ${status.done || 0} / ${status.total || 0}${status.current ? '：' + status.current : ''}`
        : '正在读取处理文件夹和共享文件夹…';
    } else if (status.error) {
      $('finProgress').textContent = '! 检查没做完';
    } else if (status.result) {
      const r = status.result;
      $('finProgress').textContent = r.passed ? `✓ 检查完了：${r.files} 个文件都已传回共享` : '! 检查完了，有文件还没传回共享（见上面标黄的项目）';
    }
    window.ProgressEta?.update(fill, 'finish_check', fraction, running && status.phase === 'compare');
    $('finError').textContent = status.error || '';
  }

  function renderVerdict() {
    const box = $('finVerdict'), r = status.result;
    if (!r || status.running) { box.hidden = true; return; }
    box.hidden = false;
    if (r.passed) {
      box.className = 'panel fin-verdict is-pass';
      box.innerHTML = `<h2>🏁 全部通过，任务「${esc(task?.name || '')}」已完成</h2>`
        + `<p>${r.files} 个文件都已传回共享。以后在主页「切换」里这个任务会显示「✓ 已完成」。</p>`
        + `<div class="link-actions"><a class="btn btn-primary" href="/">回主页</a><a class="btn" href="/distribute/start">开下一个新任务</a></div>`;
    } else {
      const n = Object.values(r.issues).reduce((sum, list) => sum + list.length, 0);
      box.className = 'panel fin-verdict is-fail';
      box.innerHTML = `<h2>还不能完成：${n} 个文件要先处理</h2>`
        + `<p>按上面标黄项目里的说明处理完，再点「再检查一遍」。下面「逐个文件」里能看到是哪些文件。</p>`;
    }
  }

  function renderFiles() {
    const r = status.result;
    $('finFilesPanel').hidden = !r || status.running;
    if (!r || status.running) return;
    const rows = r.file_states || [];
    const counts = {};
    for (const row of rows) counts[row.state] = (counts[row.state] || 0) + 1;
    const bad = rows.filter(row => row.state !== 'ok').length;
    if (!filter) filter = bad ? 'bad' : 'all';
    const chips = [['bad', `有问题 ${bad}`], ['all', `全部 ${rows.length}`], ['ok', `已一致 ${counts.ok || 0}`]]
      .concat(STEPS.slice(1).filter(s => counts[s.key]).map(s => [s.key, `${s.key} ${counts[s.key]}`]));
    $('finFilter').innerHTML = chips.map(([key, label]) => `<button type="button" data-fin-filter="${esc(key)}" aria-pressed="${filter === key}">${esc(label)}</button>`).join('');
    const shown = rows.filter(row => filter === 'all' || (filter === 'bad' ? row.state !== 'ok' : row.state === filter));
    const link = rel => (window.excelLinkHtml && task) ? window.excelLinkHtml(esc(rel), { folder: task.work, file: rel }) : esc(rel);
    $('finFiles').innerHTML = shown.length
      ? '<table class="link-table"><thead><tr><th>处理文件</th><th>结果</th></tr></thead><tbody>'
        + shown.map(row => `<tr class="${row.state === 'ok' ? 'is-ok' : 'is-bad'}"><td>${link(row.rel)}</td>`
          + `<td>${row.state === 'ok' ? '' : '! '}${esc(STATE_TEXT[row.state] || row.state)}${row.note ? '：' + esc(row.note) : ''}</td></tr>`).join('')
        + '</tbody></table>'
      : '<p>没有这一类的文件。</p>';
  }

  function render() {
    renderSteps();
    renderProgress();
    renderVerdict();
    renderFiles();
    $('finStart').disabled = !task || !!status.running;
    $('finStart').textContent = status.running ? '正在检查…' : status.result || status.error ? '🔄 再检查一遍' : '🏁 开始收尾检查';
    const passed = status.result?.passed && !status.running;
    $('finTaskDone').hidden = !(task?.done || passed);
    if (passed && !celebrated && status.fresh) {
      celebrated = true;
      window.gameFx?.celebrate('任务完成！', `${task?.name || ''} · ${status.result.files} 个文件都已传回共享`, 3);
    }
  }

  async function poll() {
    clearTimeout(timer);
    try {
      const response = await fetch('/api/tasks/finish/status', { cache: 'no-store' });
      const data = await response.json();
      const wasRunning = status.running;
      // 别的任务的检查结果不显示
      status = data.ok && data.task?.id === task?.id ? { ...data, fresh: status.fresh || (wasRunning && !data.running) } : {};
    } catch (_) {}
    render();
    if (status.running) timer = setTimeout(poll, 600);
  }

  async function start() {
    if (!task || status.running) return;
    $('finError').textContent = '';
    try {
      const response = await fetch('/api/tasks/finish/start', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ id: task.id }) });
      const data = await response.json();
      if (!response.ok || !data.ok) throw Error(data.error || '没能开始检查');
      filter = '';
      celebrated = false;
      status = { ...data, fresh: true };
      render();
      timer = setTimeout(poll, 400);
    } catch (error) { $('finError').textContent = error.message; }
  }

  async function init() {
    try {
      const data = await (await fetch('/api/tasks', { cache: 'no-store' })).json();
      task = (data.tasks || []).find(t => t.id === data.current) || null;
    } catch (_) {}
    if (!task) {
      $('finTaskName').textContent = '还没有任务';
      $('finNoTask').hidden = false;
    } else {
      $('finTaskName').textContent = task.name;
      $('finTaskPaths').innerHTML = `<dt>共享</dt><dd>${esc(task.shared)}</dd><dt>处理</dt><dd>${esc(task.work)}</dd>`;
    }
    await poll();
    // 从主页点「完成任务」进来（?start=1）：直接开始检查
    if (task && !status.running && new URLSearchParams(location.search).get('start') === '1') {
      history.replaceState(null, '', '/finish');
      start();
    }
  }

  $('finStart').addEventListener('click', start);
  $('finFilter').addEventListener('click', event => {
    const button = event.target.closest('[data-fin-filter]');
    if (!button) return;
    filter = button.dataset.finFilter;
    renderFiles();
  });
  init();
})();
