// 1～8 步的文件夹只能跟着「当前任务」走（用户 2026-11-03 定：不走新任务不能直接填路径）
// 文件夹框改成只读，填成当前任务的处理文件夹；原始备份框填成参考文件夹。还没开任务就提示回主页开。
// 后端 work_task.require_folder 也会拦，这里只是让页面上看得清楚、不用手填。
(() => {
  const WORK_IDS = ['folderInput', 'blFolder', 'cleanFolderInput', 'repairFolder', 'prFolder', 'taskFolderInput'];
  const REF_IDS = ['repairOldFolder', 'prOldFolder'];
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
  const inputs = ids => ids.map(id => document.getElementById(id)).filter(Boolean);
  const work = inputs(WORK_IDS), refs = inputs(REF_IDS);
  if (!work.length && !refs.length) return;

  const style = document.createElement('style');
  style.textContent = `
    input.task-locked { background: #eef3ee !important; color: #1d3b22 !important; cursor: default; }
    .task-folder-note { margin: 6px 0 0; font-size: 13px; line-height: 1.6; color: #1f5a2a; }
    .task-folder-note a { color: #14532d; font-weight: 700; }
    .task-folder-note.missing { color: #8a1c1c; font-weight: 700; }
    .task-folder-note.missing a { color: #8a1c1c; }`;
  document.head.appendChild(style);

  const placeholders = new Map([...work, ...refs].map(input => [input, input.placeholder]));
  const lock = on => {
    for (const input of [...work, ...refs]) {
      input.readOnly = on;
      input.classList.toggle('task-locked', on);
      if (!on) { input.placeholder = placeholders.get(input); input.title = ''; }
    }
  };
  lock(true);

  function setValue(input, value) {
    input.title = value ? `${value}\n（跟着当前任务走，不能手改；换任务请回主页）` : '';
    if (input.value === value) return;
    input.value = value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function note(input, html, missing) {
    // 注释放在输入框所在那一组（label / 输入槽）后面，一页只放一条
    // 输入框在横排（flex / grid）里时往外找，放到整排下面，不挤输入框
    let anchor = input.closest('label, .input-groove-wrap') || input;
    const inRow = el => {
      const css = getComputedStyle(el);
      return css.display.includes('grid') || (css.display.includes('flex') && css.flexDirection.startsWith('row'));
    };
    while (anchor.parentElement && anchor.parentElement !== document.body && inRow(anchor.parentElement)) anchor = anchor.parentElement;
    let el = document.getElementById('taskFolderNote');
    if (!el) {
      el = document.createElement('p');
      el.id = 'taskFolderNote';
      anchor.insertAdjacentElement('afterend', el);
    }
    el.className = 'task-folder-note' + (missing ? ' missing' : '');
    el.innerHTML = html;
  }

  async function apply() {
    let data = null;
    try {
      const response = await fetch('/api/tasks', { cache: 'no-store' });
      const result = await response.json();
      if (result.ok) data = result;
    } catch (_) {}
    if (!data) return;
    const task = (data.tasks || []).find(t => t.id === data.current);
    const first = work[0] || refs[0];
    if (data.single === true) {
      // 单项模式：自己填本地文件夹，后端照样拦 OneDrive
      lock(false);
      note(first, '🔧 单项模式：可以自己填本地文件夹；OneDrive（共享）里的文件夹不让改。要回到任务模式请去 <a href="/">主页</a>。', false);
      return;
    }
    lock(true);
    if (!task) {
      for (const input of [...work, ...refs]) { setValue(input, ''); input.placeholder = '还没开任务，请先回主页点「🆕 新任务」'; }
      note(first, '⚠ 还没开任务：程序只处理任务的本地文件夹，不能直接填路径。请先回 <a href="/">主页</a> 点「🆕 新任务」。', true);
      return;
    }
    try {
      localStorage.setItem('clear_handle_excel.lastFolder', task.work);
      localStorage.setItem('clear_handle_excel.oldFolder', task.reference);
    } catch (_) {}
    for (const input of work) setValue(input, task.work);
    for (const input of refs) setValue(input, task.reference);
    note(first, `🔒 跟着当前任务「${esc(task.name)}」走：Excel 文件夹 = 处理${refs.length ? '，原始备份 = 参考' : ''}。要换任务请回 <a href="/">主页</a>。`, false);
  }

  apply();
  window.refreshTaskFolder = apply;
})();
