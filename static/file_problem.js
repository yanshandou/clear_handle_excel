/* ==========================================
   批量任务遇到有问题的文件：任务停下来等你选「重试」或「跳过」（后端 core/file_problem.py）
   每个页面都引入这个脚本，每 2 秒问一次后台；有等待中的问题就弹框：
   - 显示文件路径，点一下用 Excel 打开（任务暂停中也能打开这份文件）
   - 修好、保存并关闭后点「重试」；不修就点「跳过」，接着做下一份
   - 勾选「后面再出错都直接跳过」：这次任务不再弹框（夜里挂机用）
   按 Esc / × 可以先收起，页面左下角留一个提示条，点它再打开
   ========================================== */
(function () {
  const POLL_MS = 2000;
  let current = null;     // 正在显示的问题
  let dialog = null;
  let bar = null;
  let minimized = '';     // 收起的问题 id
  let sending = false;
  let timer = null;

  function el(tag, cls, text) {
    const node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text !== undefined) node.textContent = text;
    return node;
  }

  function waitedText(seconds) {
    if (seconds < 60) return '刚刚';
    const minutes = Math.floor(seconds / 60);
    return minutes < 60 ? `已经等了 ${minutes} 分钟` : `已经等了 ${Math.floor(minutes / 60)} 小时 ${minutes % 60} 分钟`;
  }

  function build() {
    dialog = el('dialog', 'app-dialog file-problem-dialog');
    dialog.dataset.tone = 'warn';
    dialog.setAttribute('aria-labelledby', 'fileProblemTitle');
    dialog.innerHTML = `
      <div class="app-dialog-head">
        <span class="app-dialog-icon" aria-hidden="true">⏸</span>
        <h2 id="fileProblemTitle" class="app-dialog-title">这份文件出错了，任务已暂停</h2>
        <button type="button" class="file-problem-hide" aria-label="先收起" title="先收起（任务仍然暂停，等你选择）">×</button>
      </div>
      <div class="file-problem-body">
        <p class="file-problem-task"></p>
        <div class="file-problem-label">出错的文件</div>
        <a href="#" class="file-problem-path" title="点击用 Excel 打开这份文件"></a>
        <div class="file-problem-label">原因</div>
        <div class="file-problem-error"></div>
        <p class="file-problem-hint">在 Excel 里修好后，请先<b>保存并关闭</b>这份文件，再点「重试」。
          点「跳过」会记下错误，接着处理下一份文件，以后可以「继续上次任务」再做它。</p>
        <label class="file-problem-auto"><input type="checkbox"> 这次任务后面再有文件出错，都直接跳过（不再弹框，适合夜里挂机）</label>
      </div>
      <div class="app-dialog-foot">
        <span class="file-problem-waited"></span>
        <button type="button" class="btn btn-outline file-problem-skip">跳过这份</button>
        <button type="button" class="btn btn-primary file-problem-retry">重试</button>
      </div>`;
    document.body.appendChild(dialog);
    dialog.querySelector('.file-problem-path').addEventListener('click', e => { e.preventDefault(); openFile(); });
    dialog.querySelector('.file-problem-skip').addEventListener('click', () => choose('skip'));
    dialog.querySelector('.file-problem-retry').addEventListener('click', () => choose('retry'));
    dialog.querySelector('.file-problem-hide').addEventListener('click', hide);
    dialog.addEventListener('cancel', e => { e.preventDefault(); hide(); });
    dialog.querySelector('.file-problem-auto input').addEventListener('change', e => {
      dialog.querySelector('.file-problem-retry').disabled = e.target.checked;
    });
    return dialog;
  }

  function buildBar() {
    bar = el('button', 'file-problem-bar');
    bar.type = 'button';
    bar.addEventListener('click', () => { minimized = ''; if (current) show(current); });
    document.body.appendChild(bar);
    return bar;
  }

  function hide() {
    if (!current) return;
    minimized = current.id;
    if (dialog && dialog.open) dialog.close();
    render(current);
  }

  function toast(text, isError) {
    if (typeof window.excelToast === 'function') window.excelToast(text, isError);
    else if (isError && typeof window.appAlert === 'function') window.appAlert(text, { tone: 'error' });
  }

  async function openFile() {
    if (!current) return;
    toast('正在用 Excel 打开…');
    try {
      const res = await fetch('/api/excel/open', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ file: current.path, folder: current.folder, problem: current.id }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
      toast(data.note || '已在 Excel 中打开：修好后保存并关闭，再点「重试」', !!data.note);
    } catch (e) {
      toast(`无法打开：${e.message}`, true);
    }
  }

  async function choose(choice) {
    if (!current || sending) return;
    sending = true;
    const autoSkip = choice === 'skip' && dialog.querySelector('.file-problem-auto input').checked;
    dialog.querySelectorAll('.app-dialog-foot .btn').forEach(b => { b.disabled = true; });
    try {
      const res = await fetch('/api/file-problem', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: current.id, choice, auto_skip: autoSkip }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok || !data.ok) throw new Error(data.error || `HTTP ${res.status}`);
      toast(choice === 'retry' ? `正在重试：${current.name}` : `已跳过：${current.name}，接着处理下一份`);
      close();
    } catch (e) {
      toast(e.message, true);
      close();
    } finally {
      sending = false;
      poll();
    }
  }

  function close() {
    current = null;
    minimized = '';
    if (dialog && dialog.open) dialog.close();
    if (bar) bar.hidden = true;
  }

  function fill(problem) {
    const d = dialog || build();
    d.querySelector('.file-problem-task').textContent = problem.task
      ? `「${problem.task}」处理到这份文件时出错，后面的文件在等你决定。` : '处理到这份文件时出错，后面的文件在等你决定。';
    d.querySelector('.file-problem-path').textContent = problem.path;
    d.querySelector('.file-problem-error').textContent = problem.error || '（没有说明）';
    d.querySelector('.file-problem-waited').textContent = waitedText(problem.waited || 0);
  }

  function show(problem) {
    const d = dialog || build();
    const fresh = !current || current.id !== problem.id;
    current = problem;
    fill(problem);
    if (fresh) {
      d.querySelector('.file-problem-auto input').checked = false;
      d.querySelectorAll('.app-dialog-foot .btn').forEach(b => { b.disabled = false; });
    }
    if (bar) bar.hidden = true;
    if (!d.open) {
      document.body.appendChild(d);
      d.showModal();
      d.querySelector('.file-problem-retry').focus();
    }
  }

  function render(problem) {
    if (!problem) { close(); return; }
    if (minimized === problem.id) {
      current = problem;
      fill(problem);
      const b = bar || buildBar();
      b.textContent = `⏸ 任务暂停：${problem.name} 出错，等你选择重试或跳过（${waitedText(problem.waited || 0)}）`;
      b.hidden = false;
      return;
    }
    show(problem);
  }

  async function poll() {
    clearTimeout(timer);
    try {
      const res = await fetch('/api/file-problem', { cache: 'no-store' });
      const data = await res.json();
      if (!sending) render(data.problem || null);
    } catch (_) {
      // 程序关掉了或者网络断了：下次再问
    }
    timer = setTimeout(poll, POLL_MS);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', poll);
  else poll();
})();
