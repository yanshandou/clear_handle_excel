// 「本轮只处理待重新处理文件」顶部提示（后端 core/work_task.py 的 focus，用户 2026-11-06 定）
// 上传页一键开启后，1～8 步列文件时只列清单里的文件；清单里的文件全部上传成功自动结束。
// 主页、1～8 步各页、上传页都显示这条提示，可以看清单、手动退出。
(() => {
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
  const style = document.createElement('style');
  style.textContent = `
.focus-banner { margin:0 0 14px; padding:10px 14px; border:2px solid #b45309; border-radius:12px; background:#fff7d6;
  color:#3b2604; font-size:14px; line-height:1.6; box-shadow:0 2px 0 #b45309; }
.focus-banner[hidden] { display:none; }
.focus-banner-row { display:flex; align-items:center; gap:10px; flex-wrap:wrap; }
.focus-banner-text { flex:1 1 320px; }
.focus-banner-text strong { color:#7c2d12; }
.focus-banner-tag { display:inline-block; margin-right:6px; padding:1px 9px; border-radius:999px; background:#b45309; color:#fff;
  font-weight:700; font-size:12px; letter-spacing:.04em; }
.focus-banner button { font:inherit; font-size:13px; padding:4px 12px; border-radius:8px; cursor:pointer; border:1.5px solid #7c2d12;
  background:#fff; color:#7c2d12; font-weight:600; }
.focus-banner button:hover { background:#ffedd5; }
.focus-banner button.focus-exit { background:#7c2d12; color:#fff; }
.focus-banner button.focus-exit:hover { background:#9a3412; }
.focus-banner ul { margin:8px 0 0; padding:8px 10px 8px 28px; max-height:160px; overflow:auto; background:#fff; border-radius:8px;
  border:1px solid #f0c674; font-family:Consolas, "MS Gothic", monospace; font-size:12.5px; color:#3b2604; }
.focus-banner ul[hidden] { display:none; }`;
  document.head.appendChild(style);

  let el = null, task = null;
  function place() {
    if (el) return el;
    el = document.createElement('section');
    el.id = 'focusBanner';
    el.className = 'focus-banner';
    el.hidden = true;
    el.setAttribute('role', 'status');
    const panel = document.getElementById('workTaskPanel');  // 主页：放在当前任务条下面
    const main = document.querySelector('main');
    if (panel) panel.insertAdjacentElement('afterend', el);
    else if (main) main.prepend(el);
    else document.body.prepend(el);
    el.addEventListener('click', event => {
      if (event.target.closest('.focus-list-toggle')) {
        const list = el.querySelector('ul');
        list.hidden = !list.hidden;
        event.target.closest('.focus-list-toggle').textContent = list.hidden ? '看清单 ▾' : '收起清单 ▴';
      } else if (event.target.closest('.focus-exit')) exit();
    });
    return el;
  }

  function render(data) {
    const current = (data.tasks || []).find(t => t.id === data.current);
    const files = (current?.focus?.files) || [];
    task = files.length ? current : null;
    if (!task && !el) return;
    const box = place();
    box.hidden = !task;
    if (!task) { box.innerHTML = ''; return; }
    box.innerHTML = `<div class="focus-banner-row"><span class="focus-banner-text"><span class="focus-banner-tag">本轮</span>`
      + `只处理 <strong>${files.length}</strong> 个待重新处理文件（任务「${esc(task.name)}」的处理 / 参考文件夹）。`
      + `其他文件这一轮不动；这几个全部上传成功后自动回到整个文件夹。</span>`
      + `<button type="button" class="focus-list-toggle">看清单 ▾</button>`
      + `<button type="button" class="focus-exit">退出，处理整个文件夹</button></div>`
      + `<ul hidden>${files.map(rel => `<li>${esc(rel)}</li>`).join('')}</ul>`;
  }

  async function load() {
    try {
      const response = await fetch('/api/tasks', { cache: 'no-store' });
      const data = await response.json();
      if (data.ok) render(data);
    } catch (_) {}
  }

  async function exit() {
    if (!task) return;
    const ask = window.appConfirm
      ? window.appConfirm('退出后，1～8 步重新处理整个文件夹的全部 Excel。\n还没上传的待重新处理文件仍在上传页标着「待重新处理」，随时可以再开。',
        { title: '退出本轮清单', okText: '退出', cancelText: '先不退' })
      : Promise.resolve(confirm('退出后 1～8 步处理整个文件夹，确定吗？'));
    if (!await ask) return;
    try {
      const response = await fetch('/api/tasks/focus/clear', { method:'POST', headers:{ 'Content-Type':'application/json' },
        body: JSON.stringify({ id: task.id }) });
      const result = await response.json();
      if (!response.ok || !result.ok) throw Error(result.error || '退出失败');
    } catch (error) {
      if (window.appAlert) window.appAlert(error.message, { title: '没能退出' }); else alert(error.message);
    }
    await load();
    document.dispatchEvent(new CustomEvent('focus-changed'));
  }

  window.refreshFocusBanner = load;
  load();
})();
