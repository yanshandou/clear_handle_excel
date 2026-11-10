// 当前任务只提供默认路径；各功能始终可以独立修改文件夹。
(() => {
  const WORK_IDS = ['folderInput', 'blFolder', 'cleanFolderInput', 'repairFolder', 'prFolder', 'taskFolderInput', 'bundleFolder'];
  const REF_IDS = ['repairOldFolder', 'prOldFolder', 'cleanOldFolder'];
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
  const inputs = ids => ids.map(id => document.getElementById(id)).filter(Boolean);
  const work = inputs(WORK_IDS), refs = inputs(REF_IDS), all = [...work, ...refs];
  if (!all.length) return;
  for (const input of all) {
    input.readOnly = false;
    input.classList.remove('task-locked');
  }
  const style = document.createElement('style');
  style.textContent = '.task-folder-note { margin:6px 0 0; font-size:13px; line-height:1.6; color:#475569; }';
  document.head.appendChild(style);
  let sequence = 0;

  const norm = value => String(value || '').trim().replace(/[\\/]+/g, '/').replace(/\/$/, '').toLowerCase();
  const KEYS = { work:'clear_handle_excel.lastFolder', reference:'clear_handle_excel.oldFolder' };

  function setValue(input, value) {
    if (input.value === value) return;
    input.value = value;
    input.title = value;
    input.dispatchEvent(new Event('input', { bubbles:true }));
    input.dispatchEvent(new Event('change', { bubbles:true }));
  }

  function note(html) {
    const input = work[0] || refs[0];
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
    el.className = 'task-folder-note';
    el.innerHTML = html;
  }

  async function apply(useTaskDefaults = false) {
    const version = ++sequence;
    const before = new Map(all.map(input => [input, input.value]));
    let data;
    try {
      const response = await fetch('/api/tasks', { cache:'no-store' });
      data = await response.json();
      if (!data.ok || version !== sequence) return;
    } catch (_) { return; }
    const tasks = data.tasks || [];
    const task = tasks.find(t => t.id === data.current);
    if (task && !data.single) {
      for (const [group, field] of [[work, 'work'], [refs, 'reference']]) {
        // 框里是「别的任务」的默认路径（浏览器记住的是旧任务）→ 换成当前任务的；自己填的其他路径照样保留
        const otherTasks = new Set(tasks.filter(t => t.id !== task.id).map(t => norm(t[field])));
        for (const input of group) {
          const stale = otherTasks.has(norm(input.value)) && norm(input.value) !== norm(task[field]);
          // 异步加载任务信息期间输入的路径也要保留。
          if (input.value === before.get(input) && (useTaskDefaults || !input.value.trim() || stale)) {
            setValue(input, task[field]);
            if (stale) { try { localStorage.setItem(KEYS[field], task[field]); } catch (_) {} }
          }
        }
      }
      note(`当前任务「${esc(task.name)}」仅提供默认路径；可以直接修改文件夹，按实际选择执行。`);
    } else {
      note('可以直接填写本地文件夹；各文件夹的进度分别保存。');
    }
  }

  apply();
  window.refreshTaskFolder = () => apply(true);
})();
