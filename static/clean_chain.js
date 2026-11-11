// 精简 Sheet 做完后自动接着做的检查（排查 BL 超链接 → 修复引用错误 →（勾选时）修复隐藏页 → 统一字体）：
// 显示每一步的进度和结果，改不了 / 出错的逐处列出来（点了用 Excel 打开跳到格子），最后只给一个「返回首页」按钮。
// 后台见 core/clean_chain.py；结果存在数据库里，重开程序也还在。
(function () {
  const DETAIL_PAGES = { bl: '/repair-bl-links', links: '/repair-links', hidden: '/repair-hidden', font: '/fix-fonts' };
  const ICONS = { waiting: '·', running: '⏳', done: '✓', warn: '!', failed: '✕', skipped: '–', off: '–' };
  const STATE_TEXT = {
    waiting: '等待中', running: '进行中', done: '完成', warn: '有改不了的', failed: '有出错的',
    skipped: '跳过', off: '没勾选',
  };
  let data = null;
  let wasRunning = false;
  let onFinished = null;

  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const norm = value => String(value || '').trim().replace(/^["']|["']$/g, '').replace(/[\\/]+/g, '/')
    .replace(/\/$/, '').toLowerCase();

  // 有配对到的参考文件（引用错误 / 隐藏页）：点了左边开处理后的文件、右边只读开参考文件，都跳到同一格
  function problemHtml(problem, step) {
    const where = problem.sheet ? `${problem.sheet}${problem.cell ? `!${problem.cell}` : ''}` : '';
    const label = esc(problem.file || '任务') + (where ? ` <span class="chain-where">${esc(where)}</span>` : '');
    const link = problem.file && window.excelLinkHtml
      ? window.excelLinkHtml(label, { folder: data.folder, file: problem.file, sheet: problem.sheet, cell: problem.cell,
        ...(problem.old ? { compare: `chain_${step.key}` } : {}) })
      : label;
    return `<li class="${problem.error ? 'is-error' : ''}">${link}<span class="chain-reason">${esc(problem.reason)}</span></li>`;
  }

  // 正在做的这一步：自己的进度条（第几份 / 共几份 + 预计几点完成）。
  // 数字和宽度不放进 HTML 里比较，每次轮询只改这几处，不重建整块。
  function progressSlot(step) {
    if (step.state !== 'running' || !step.progress || !step.progress.total) return '';
    return `<div class="chain-progress" data-step="${esc(step.key)}" data-phase="${esc(step.progress.phase)}">
        <div class="progress-track progress-track-sm"><div class="progress-fill" style="width: 0%;"></div></div>
        <p class="chain-progress-count"></p><p class="progress-eta chain-progress-eta"></p></div>`;
  }

  const etaKey = (key, phase) => `clean-chain.${key}.${phase}`;
  let etaKeys = new Set();

  function fillProgress(box) {
    const live = new Set();
    box.querySelectorAll('.chain-progress').forEach(slot => {
      const step = data.steps.find(item => item.key === slot.dataset.step);
      const progress = step && step.progress;
      if (!progress || !progress.total) return;
      const key = etaKey(step.key, progress.phase);
      live.add(key);
      const fraction = Math.min(1, progress.done / progress.total);
      const fill = slot.querySelector('.progress-fill');
      const width = `${Math.max(2, fraction * 100).toFixed(1)}%`;
      if (fill.style.width !== width) fill.style.width = width;
      const shown = Math.min(progress.total, progress.done + 1);
      const count = `${progress.phase} 第 ${shown} / ${progress.total} 份` + (progress.current ? ` · ${progress.current}` : '');
      const countEl = slot.querySelector('.chain-progress-count');
      if (countEl.textContent !== count) { countEl.textContent = count; countEl.title = count; }
      const eta = window.ProgressEta ? window.ProgressEta.text(key, fraction, undefined, '这一项') : '';
      const etaEl = slot.querySelector('.chain-progress-eta');
      if (etaEl.textContent !== eta) etaEl.textContent = eta;
      etaEl.hidden = !eta;
    });
    // 做完的那一步：清掉预估时间的记录
    etaKeys.forEach(key => { if (!live.has(key)) window.ProgressEta?.stop(key); });
    etaKeys = live;
  }

  function stepHtml(step, index) {
    const more = step.problem_total > step.problems.length
      ? `<li class="chain-more">还有 ${step.problem_total - step.problems.length} 处没列出，请点「看详细结果」</li>` : '';
    const errors = step.error_total ? `出错 ${step.error_total} 处` : '';
    const skipped = step.problem_total - step.error_total;
    const summary = [skipped ? `改不了 ${skipped} 处` : '', errors].filter(Boolean).join('，');
    const list = step.problem_total
      ? `<details class="chain-problems"><summary>${esc(summary)}（点开看是哪些）</summary><ul>${step.problems.map(problem => problemHtml(problem, step)).join('')}${more}</ul></details>`
      : '';
    const detail = ['done', 'warn', 'failed'].includes(step.state) && DETAIL_PAGES[step.key]
      ? `<a class="chain-detail" href="${DETAIL_PAGES[step.key]}" title="在原来的页面看这一步每个文件的详细结果">看详细结果 ↗</a>` : '';
    return `<li class="chain-step is-${esc(step.state)}">
      <span class="chain-icon" aria-hidden="true">${ICONS[step.state] || '·'}</span>
      <div class="chain-body">
        <div class="chain-line"><b>${index + 1}. ${esc(step.label)}</b>
          <span class="chain-state">${esc(STATE_TEXT[step.state] || step.state)}</span>
          <span class="chain-text">${esc(step.text || '')}</span>${detail}</div>
        ${progressSlot(step)}${list}
      </div>
    </li>`;
  }

  function render(box, visible) {
    if (!box) return;
    if (!visible || !data || !data.steps) {
      box.classList.add('hidden');
      return;
    }
    const running = Boolean(data.running);
    const bad = data.steps.some(step => step.state === 'failed');
    const warn = data.steps.some(step => step.state === 'warn');
    const title = running ? (data.stage === 'clean' ? '精简完后会自动接着检查' : '正在自动检查…')
      : bad ? '自动检查做完了，有出错的地方' : warn ? '自动检查做完了，有改不了的地方' : '自动检查全部做完了';
    const note = running ? '这几步不用再点，做完会在这里写结果。关掉网页也会继续。'
      : bad || warn ? '改不了、出错的都列在下面，点文件名可以用 Excel 打开跳到那一格。没问题的已经直接改好保存了。'
      : '没有要人工处理的地方。';
    // 做完只有「返回首页」（首页上精简 Sheet 那张卡片会闪几下）
    const next = running ? '' : `<div class="chain-next">
        <a class="btn btn-primary next-step-btn" href="/?from=clean">🏠 返回首页</a></div>`;
    const html = `<div class="chain-head"><h3>🤖 ${esc(title)}</h3><p>${esc(note)}</p></div>
      <ol class="chain-steps">${data.steps.map(stepHtml).join('')}</ol>${next}`;
    // 轮询时内容不变就不重建，免得打断点开的列表和选中复制
    if (box.dataset.html !== html) {
      const open = [...box.querySelectorAll('details')].map(el => el.open);
      box.dataset.html = html;
      box.innerHTML = html;
      box.querySelectorAll('details').forEach((el, i) => { if (open[i]) el.open = true; });
    }
    fillProgress(box);
    box.className = `panel clean-chain${running ? ' is-running' : bad ? ' is-failed' : warn ? ' is-warn' : ' is-ok'}`;
  }

  // 实时进度框（static/job_panel.js，用户 2026-11-09 要：所有批量操作都要有）：精简做完接着自动检查时，
  // 框里写正在做第几项、哪份文件，做完一项列一条结果
  function renderPanel() {
    if (!window.JobPanel || !data || !data.steps || data.stage === 'clean') return;
    const steps = data.steps.filter(step => step.state !== 'off');
    const log = [{ ok: true, name: '精简 Sheet', reason: '做完了' }].concat(steps
      .filter(step => ['done', 'warn', 'failed', 'skipped'].includes(step.state))
      .map(step => ({ ok: step.state === 'done' || step.state === 'skipped', name: step.label,
        reason: [STATE_TEXT[step.state], step.text].filter(Boolean).join('：') })));
    if (data.running) {
      const index = steps.findIndex(step => step.state === 'running');
      const step = steps[index] || {};
      const progress = step.progress || {};
      window.JobPanel.run('chain', {
        title: '正在自动检查（精简后）', count: index >= 0 ? `第 ${index + 1} / ${steps.length} 项` : '',
        now: step.label ? `${step.label}${progress.current ? ` · 📗 ${progress.current}` : ''}` : '准备中…',
        step: progress.total ? `${progress.phase || ''} 第 ${Math.min(progress.total, (progress.done || 0) + 1)} / ${progress.total} 份` : (step.text || ''),
        // 进度条和预计时间都只算正在做的这一项，和页面上这一项的卡片共用一份预估记录，两处时间一样。
        // 以前按「整条链」算：用第一项的速度去猜后面几项，后面几项每份快慢差很多，猜出来的时间不准，还和卡片对不上。
        fraction: progress.total ? Math.min(1, progress.done / progress.total) : 0,
        ...(progress.total ? { etaKey: etaKey(step.key, progress.phase), etaLabel: '这一项' } : {}),
        log, folder: data.folder,
      });
    } else {
      const bad = log.filter(item => !item.ok).length;
      window.JobPanel.finish('chain', {
        title: bad ? `自动检查做完了，${bad} 项有要人工看的` : '自动检查全部做完了', log, folder: data.folder,
        step: bad ? '改不了、出错的都列在页面上的「自动检查」里，点文件名可以用 Excel 打开跳到那一格。' : '',
      });
    }
  }

  async function refresh() {
    try {
      const res = await fetch('/api/clean/chain', { cache: 'no-store' });
      const body = await res.json();
      if (!res.ok || !body.ok) return data;
      data = body;
      renderPanel();
      if (wasRunning && !data.running && onFinished) onFinished(data);
      wasRunning = Boolean(data.running);
    } catch (_) { /* 读不到就保持上次的样子 */ }
    return data;
  }

  window.CleanChain = {
    refresh,
    render,
    get data() { return data; },
    get running() { return Boolean(data && data.running); },
    // 结果是不是这个文件夹的（页面上的文件夹框可能填的是没整理过的写法）
    isFor(folder) { return Boolean(data && data.folder && norm(data.folder) === norm(folder)); },
    onFinished(callback) { onFinished = callback; },
  };
})();
