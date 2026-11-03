// 主页「当前任务」：新任务（只填共享路径，自动生成参考 / 处理文件夹）、切换以前的任务（后端 core/work_task.py）
// 开新任务或切换后，把 1～8 步共用的文件夹（lastFolder）设成处理文件夹、原始备份（oldFolder）设成参考文件夹
(() => {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
  let data = { tasks: [], current: '' };
  let deriveTimer = 0, deriveSeq = 0, creating = false;

  const pathsHtml = t => `<dt>共享</dt><dd>${esc(t.shared)}</dd><dt>参考</dt><dd>${esc(t.reference)}</dd><dt>处理</dt><dd>${esc(t.work)}</dd>`;

  function applyTask(task) {
    try {
      localStorage.setItem('clear_handle_excel.lastFolder', task.work);
      localStorage.setItem('clear_handle_excel.oldFolder', task.reference);
    } catch (_) {}
    const input = $('taskFolderInput');
    if (input) input.value = task.work;
    if (typeof window.refreshDoneMarks === 'function') window.refreshDoneMarks();
  }

  function render() {
    const current = data.tasks.find(t => t.id === data.current);
    $('workTaskName').textContent = current ? current.name : '还没有任务';
    $('workTaskPaths').innerHTML = current ? pathsHtml(current) : '';
    $('workTaskHint').textContent = current
      ? '1～8 步的文件夹已经填成上面的「处理」文件夹，修复颜色 / 引用 / 隐藏页的原始备份是「参考」文件夹。第 9 步把处理好的文件传回共享。'
      : '点「🆕 新任务」填共享文件夹路径，工具会自动生成本地的参考、处理文件夹，并把 1～9 步的文件夹都填好。';
    const others = data.tasks.filter(t => t.id !== data.current);
    $('workTaskSwitch').innerHTML = others.length
      ? '<option value="">选择一个任务…</option>' + others.map(t => `<option value="${esc(t.id)}">${esc(t.name)}（${esc(t.used || t.created || '')}）</option>`).join('')
      : '<option value="">（没有以前的任务）</option>';
    $('workTaskSwitch').disabled = !others.length;
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

  function scheduleDerive() {
    clearTimeout(deriveTimer);
    deriveTimer = setTimeout(async () => {
      const seq = ++deriveSeq;
      const shared = $('workTaskShared').value.trim();
      $('workTaskError').textContent = '';
      if (!shared) { $('workTaskPreview').innerHTML = ''; return; }
      try {
        const result = await post('/api/tasks/derive', { shared, work_root: $('workTaskRoot').value.trim() });
        if (seq !== deriveSeq) return;
        $('workTaskPreview').innerHTML = pathsHtml(result.paths);
        if (!$('workTaskRoot').value.trim()) $('workTaskRoot').placeholder = result.paths.work_root;
      } catch (error) {
        if (seq !== deriveSeq) return;
        $('workTaskPreview').innerHTML = '';
        $('workTaskError').textContent = error.message;
      }
    }, 350);
  }

  function openDialog() {
    const current = data.tasks.find(t => t.id === data.current);
    $('workTaskShared').value = current ? current.shared : '';
    $('workTaskRoot').value = data.last_work_root || '';
    $('workTaskError').textContent = '';
    $('workTaskPreview').innerHTML = '';
    $('workTaskDialog').showModal();
    $('workTaskShared').focus();
    $('workTaskShared').select();
    scheduleDerive();
  }

  async function create() {
    if (creating) return;
    const shared = $('workTaskShared').value.trim();
    if (!shared) { $('workTaskError').textContent = '请填写共享文件夹路径'; return; }
    creating = true;
    $('workTaskCreate').disabled = true;
    $('workTaskCreate').textContent = '正在建文件夹…';
    try {
      const { task } = await post('/api/tasks/new', { shared, work_root: $('workTaskRoot').value.trim() });
      applyTask(task);
      // 第 0 步：去第 9 步的页面拿文件（自动生成拿取清单，确认后才拿）
      window.location.href = '/distribute?fetch=1';
    } catch (error) {
      $('workTaskError').textContent = error.message;
      $('workTaskCreate').disabled = false;
      $('workTaskCreate').textContent = '开始，去拿文件';
      creating = false;
    }
  }

  $('workTaskNew').addEventListener('click', openDialog);
  $('workTaskCancel').addEventListener('click', () => $('workTaskDialog').close());
  $('workTaskCreate').addEventListener('click', create);
  $('workTaskShared').addEventListener('input', scheduleDerive);
  $('workTaskRoot').addEventListener('input', scheduleDerive);
  $('workTaskShared').addEventListener('keydown', event => { if (event.key === 'Enter') create(); });
  $('workTaskSwitch').addEventListener('change', async event => {
    const id = event.target.value;
    if (!id) return;
    try {
      const { task } = await post('/api/tasks/switch', { id });
      applyTask(task);
      await load();
    } catch (error) {
      window.appAlert(error.message, { tone: 'error' });
      event.target.value = '';
    }
  });
  load();
})();
