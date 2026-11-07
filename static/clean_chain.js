// 精简 Sheet 做完后自动接着做的检查（排查 BL 超链接 → 修复引用错误 →（勾选时）修复隐藏页）：
// 显示每一步的进度和结果，改不了 / 出错的逐处列出来（点了用 Excel 打开跳到格子），最后给「下一步」按钮。
// 后台见 core/clean_chain.py；结果存在数据库里，重开程序也还在。
(function () {
  const DETAIL_PAGES = { bl: '/repair-bl-links', links: '/repair-links', hidden: '/repair-hidden' };
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

  function problemHtml(problem) {
    const where = problem.sheet ? `${problem.sheet}${problem.cell ? `!${problem.cell}` : ''}` : '';
    const label = esc(problem.file || '任务') + (where ? ` <span class="chain-where">${esc(where)}</span>` : '');
    const link = problem.file && window.excelLinkHtml
      ? window.excelLinkHtml(label, { folder: data.folder, file: problem.file, sheet: problem.sheet, cell: problem.cell })
      : label;
    return `<li class="${problem.error ? 'is-error' : ''}">${link}<span class="chain-reason">${esc(problem.reason)}</span></li>`;
  }

  function stepHtml(step, index) {
    const more = step.problem_total > step.problems.length
      ? `<li class="chain-more">还有 ${step.problem_total - step.problems.length} 处没列出，请点「看详细结果」</li>` : '';
    const errors = step.error_total ? `出错 ${step.error_total} 处` : '';
    const skipped = step.problem_total - step.error_total;
    const summary = [skipped ? `改不了 ${skipped} 处` : '', errors].filter(Boolean).join('，');
    const list = step.problem_total
      ? `<details class="chain-problems"><summary>${esc(summary)}（点开看是哪些）</summary><ul>${step.problems.map(problemHtml).join('')}${more}</ul></details>`
      : '';
    const detail = ['done', 'warn', 'failed'].includes(step.state) && DETAIL_PAGES[step.key]
      ? `<a class="chain-detail" href="${DETAIL_PAGES[step.key]}" title="在原来的页面看这一步每个文件的详细结果">看详细结果 ↗</a>` : '';
    return `<li class="chain-step is-${esc(step.state)}">
      <span class="chain-icon" aria-hidden="true">${ICONS[step.state] || '·'}</span>
      <div class="chain-body">
        <div class="chain-line"><b>${index + 1}. ${esc(step.label)}</b>
          <span class="chain-state">${esc(STATE_TEXT[step.state] || step.state)}</span>
          <span class="chain-text">${esc(step.text || '')}</span>${detail}</div>
        ${list}
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
    const next = running ? '' : `<div class="chain-next">
        <a class="btn btn-primary next-step-btn" href="/workbench">下一步：查看切字 ➜</a></div>`;
    const html = `<div class="chain-head"><h3>🤖 ${esc(title)}</h3><p>${esc(note)}</p></div>
      <ol class="chain-steps">${data.steps.map(stepHtml).join('')}</ol>${next}`;
    // 轮询时内容不变就不重建，免得打断点开的列表和选中复制
    if (box.dataset.html !== html) {
      const open = [...box.querySelectorAll('details')].map(el => el.open);
      box.dataset.html = html;
      box.innerHTML = html;
      box.querySelectorAll('details').forEach((el, i) => { if (open[i]) el.open = true; });
    }
    box.className = `panel clean-chain${running ? ' is-running' : bad ? ' is-failed' : warn ? ' is-warn' : ' is-ok'}`;
  }

  async function refresh() {
    try {
      const res = await fetch('/api/clean/chain', { cache: 'no-store' });
      const body = await res.json();
      if (!res.ok || !body.ok) return data;
      data = body;
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
