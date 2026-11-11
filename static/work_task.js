// 主页「当前任务」：显示当前任务、切换以前的任务、收尾检查（后端 core/work_task.py）
// 开新任务在单独的「新任务 · 拿文件」页（/distribute/start，static/distribute.js）
// 开新任务或切换后，把 前 3 步共用的文件夹（lastFolder）设成处理文件夹、原始备份（oldFolder）设成参考文件夹
(() => {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
  let data = { tasks: [], current: '' };

  // 任务名只有最后一级文件夹，同名的分不清（如两个 06_SS障害者支援業務），列表里显示共享路径的最后 3 级
  const tailParts = t => String(t.shared || '').split(/[\\/]+/).filter(Boolean).slice(-3);
  const tailText = t => { const parts = tailParts(t); return parts.length ? parts.join('\\') : t.name; };
  const tailHtml = t => {
    const parts = tailParts(t);
    if (parts.length < 2) return esc(t.name);
    return `<span class="wt-parent">${parts.slice(0, -1).map(esc).join('\\')}\\</span>${esc(parts[parts.length - 1])}`;
  };
  const pathsHtml = t => `<dt>共享</dt><dd>${esc(t.shared)}</dd><dt>参考</dt><dd>${esc(t.reference)}</dd><dt>处理</dt><dd>${esc(t.work)}</dd>`;

  function applyTask(task) {
    try {
      localStorage.setItem('clear_handle_excel.lastFolder', task.work);
      localStorage.setItem('clear_handle_excel.oldFolder', task.reference);
    } catch (_) {}
    const input = $('taskFolderInput');
    if (input) input.value = task.work;
    if (typeof window.refreshDoneMarks === 'function') window.refreshDoneMarks();
    if (typeof window.refreshTaskFolder === 'function') window.refreshTaskFolder();
  }

  function render() {
    const current = data.tasks.find(t => t.id === data.current);
    const single = data.single === true;
    $('workTaskPanel').classList.toggle('single', single);
    $('workTaskSingle').textContent = single ? '📁 使用任务默认路径' : '📂 自选路径优先';
    $('workTaskName').textContent = single ? '📂 自选路径优先（各功能可独立指定）'
      : current ? tailText(current) + (current.done ? '（✓ 已完成）' : '') : '还没有任务';
    $('workTaskPaths').innerHTML = current ? pathsHtml(current) : '';
    $('workTaskHint').textContent = current
      ? '任务提供默认的处理、参考路径，各功能都可以自由修改。每个文件夹独立保存进度和处理记录，切回后可继续。「上传 Excel」按任务设置传回共享。'
      : '可以直接在各功能填写本地文件夹。需要从共享拿取、上传文件时，再点「🆕 新任务」建立共享、参考和处理路径。';
    $('workTaskName').title = current ? `共享：${current.shared}\n参考：${current.reference}\n处理：${current.work}` : '';
    const others = data.tasks.filter(t => t.id !== data.current);
    $('workTaskSwitchText').textContent = others.length ? `选择一个任务…（${others.length} 个）` : '（没有以前的任务）';
    $('workTaskSwitchList').innerHTML = others.map(t => `<li role="option" tabindex="-1" data-id="${esc(t.id)}" title="共享：${esc(t.shared)}">`
      + `<span class="wt-chip ${t.done ? 'is-done' : 'is-doing'}">${t.done ? '✓ 已完成' : '进行中'}</span>`
      + `<span class="wt-name">${tailHtml(t)}</span><span class="wt-time">${esc(t.used || t.created || '')}</span></li>`).join('');
    $('workTaskSwitch').disabled = !others.length;
    if (!others.length) closePicker();
  }

  async function load() {
    try {
      const response = await fetch('/api/tasks', { cache: 'no-store' });
      const result = await response.json();
      if (result.ok) { data = result; render(); }
    } catch (_) {}
  }

  async function post(path, payload) {
    const response = await fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    const result = await response.json();
    if (!response.ok || !result.ok) throw Error(result.error || '操作失败');
    return result;
  }

  // 终点「🏁 完成任务」是链接，打开收尾页 /finish（static/finish.js）
  // 路径偏好只决定默认值，两种偏好都允许直接修改输入框。
  $('workTaskSingle').addEventListener('click', async () => {
    const on = data.single !== true;
    try {
      data = { ...data, ...await post('/api/tasks/single', { on }) };
      render();
      if (!on) { const current = data.tasks.find(t => t.id === data.current); if (current) applyTask(current); }
      else if (typeof window.refreshTaskFolder === 'function') window.refreshTaskFolder();
    } catch (error) {
      window.appAlert(error.message, { tone: 'error' });
    }
  });
  // 「切换」下拉：自己画的列表（原生 select 弹出的列表改不了样式）
  const options = () => [...$('workTaskSwitchList').querySelectorAll('[role="option"]')];
  function closePicker(focusButton = false) {
    $('workTaskSwitchList').hidden = true;
    $('workTaskSwitch').setAttribute('aria-expanded', 'false');
    if (focusButton) $('workTaskSwitch').focus();
  }
  function openPicker() {
    if ($('workTaskSwitch').disabled) return;
    $('workTaskSwitchList').hidden = false;
    $('workTaskSwitch').setAttribute('aria-expanded', 'true');
    options()[0]?.focus();
  }
  async function pick(id) {
    closePicker(true);
    try {
      const { task } = await post('/api/tasks/switch', { id });
      applyTask(task);
      await load();
    } catch (error) {
      window.appAlert(error.message, { tone: 'error' });
    }
  }
  $('workTaskSwitch').addEventListener('click', () => ($('workTaskSwitchList').hidden ? openPicker() : closePicker()));
  $('workTaskSwitch').addEventListener('keydown', event => {
    if (event.key === 'ArrowDown') { event.preventDefault(); openPicker(); }
  });
  $('workTaskSwitchList').addEventListener('click', event => {
    const item = event.target.closest('[role="option"]');
    if (item) pick(item.dataset.id);
  });
  // 鼠标移到哪行就高亮哪行（和键盘选中是同一个高亮，不会同时亮两行）
  $('workTaskSwitchList').addEventListener('mouseover', event => event.target.closest('[role="option"]')?.focus());
  $('workTaskSwitchList').addEventListener('keydown', event => {
    const list = options(), index = list.indexOf(document.activeElement);
    if (event.key === 'ArrowDown') { event.preventDefault(); list[Math.min(index + 1, list.length - 1)]?.focus(); }
    else if (event.key === 'ArrowUp') { event.preventDefault(); list[Math.max(index - 1, 0)]?.focus(); }
    else if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); if (list[index]) pick(list[index].dataset.id); }
    else if (event.key === 'Escape') closePicker(true);
    else if (event.key === 'Tab') closePicker();
  });
  document.addEventListener('click', event => {
    if (!event.target.closest('.wt-pick')) closePicker();
  });
  load();
})();
