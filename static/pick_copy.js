// 挑文件复制（用户 2026-11-12 要）：贴一串路径 + 目标文件夹，保留目录结构只复制这些文件。后端 core/pick_copy.py
(() => {
  const $ = id => document.getElementById(id);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const KEY = 'clear_handle_excel.pickCopy';
  let busy = false;
  try {
    const saved = JSON.parse(localStorage.getItem(KEY) || '{}');
    // 默认填当前在处理的路径（各处理页共用的 lastFolder；用户 2026-11-12），没有才用上次贴的
    $('pickPaths').value = localStorage.getItem('clear_handle_excel.lastFolder') || saved.paths || '';
    $('pickDest').value = saved.dest || '';
  } catch (_) {}

  const input = () => ({ paths: $('pickPaths').value, dest: $('pickDest').value.trim() });
  const remember = () => { try { localStorage.setItem(KEY, JSON.stringify(input())); } catch (_) {} };

  async function post(url, body) {
    const response = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const result = await response.json().catch(() => ({}));
    if (!response.ok || !result.ok) throw Error(result.error || '请求失败');
    return result;
  }

  function controls() {
    for (const id of ['pickPaths', 'pickDest', 'pickPreview', 'pickRun']) $(id).disabled = busy;
  }

  function render(plan, done) {
    const groups = new Map();
    for (const file of plan.files) {
      const cut = file.rel.lastIndexOf('/');
      const dir = cut < 0 ? '' : file.rel.slice(0, cut);
      if (!groups.has(dir)) groups.set(dir, []);
      groups.get(dir).push({ ...file, name: file.rel.slice(cut + 1) });
    }
    const summary = [
      `<span class="is-ok">${done ? `复制了 ${plan.copied} 个文件` : `要复制 ${plan.files.length} 个文件`} · ${groups.size} 个文件夹</span>`,
      plan.duplicates ? `<span>重复的 ${plan.duplicates} 行只算一次</span>` : '',
      !done && plan.exists ? `<span class="is-warn">目标里已有 ${plan.exists} 个同名文件（复制时会问你要不要覆盖）</span>` : '',
      done && plan.failed ? `<span class="is-bad">${plan.failed} 个没复制成功</span>` : '',
      plan.missing.length ? `<span class="is-bad">${plan.missing.length} 行找不到</span>` : '',
    ].filter(Boolean).join('');
    const tree = [...groups].map(([dir, files]) => `<li class="pick-dir">📁 ${esc(dir || '（目标文件夹）')}</li>`
      + files.map(f => `<li class="pick-file ${f.status === 'ok' ? 'is-ok' : f.status === 'failed' ? 'is-failed' : f.exists ? 'is-exists' : ''}" title="${esc(f.source)}">${esc(f.name)}${f.error ? `：${esc(f.error)}` : ''}</li>`).join('')).join('');
    const missing = plan.missing.length ? `<h3>找不到的行（没有复制）</h3><ul class="pick-missing">${plan.missing.map(m => `<li>${esc(m.path)}<br><small>${esc(m.reason)}</small></li>`).join('')}</ul>` : '';
    $('pickResults').innerHTML = `<div class="pick-summary">${summary}</div>`
      + `<p class="toolbar-hint">目录结构从 <span class="pick-root">${esc(plan.root)}</span> 算起，放到 <span class="pick-root">${esc(plan.dest)}</span></p>`
      + missing + `<ul class="pick-tree">${tree}</ul>`;
  }

  async function go(run) {
    if (busy) return;
    $('pickError').textContent = '';
    busy = true;
    controls();
    remember();
    try {
      let plan = await post('/api/pick_copy/preview', input());
      if (run) {
        if (plan.exists && !await window.appConfirm(`目标里已经有 ${plan.exists} 个同名文件，要覆盖吗？`, { confirmText: '覆盖并复制' })) {
          render(plan, false);
          $('pickProgress').textContent = '没有复制。';
          return;
        }
        $('pickProgress').textContent = `正在复制 ${plan.files.length} 个文件…`;
        plan = await post('/api/pick_copy/run', { ...input(), overwrite: plan.exists > 0 });
      }
      render(plan, run);
      $('pickOpen').hidden = !run;
      const bad = plan.missing.length + (plan.failed || 0);
      $('pickProgress').textContent = !run ? `清单列好了：${plan.files.length} 个文件。没问题就点「📦 复制」。`
        : bad ? `! 复制了 ${plan.copied} 个，有 ${bad} 处要看一下（见下面）` : `✓ 复制好了，${plan.copied} 个文件`;
      if (run && !bad) window.gameFx?.celebrate?.('复制好了！', `${plan.copied} 个文件`);
    } catch (error) {
      $('pickError').textContent = error.message;
      $('pickProgress').textContent = '';
    } finally {
      busy = false;
      controls();
    }
  }

  $('pickPreview').addEventListener('click', () => go(false));
  $('pickRun').addEventListener('click', () => go(true));
  $('pickOpen').addEventListener('click', async () => {
    try { await post('/api/pick_copy/open', { dest: $('pickDest').value.trim() }); } catch (error) { $('pickError').textContent = error.message; }
  });
})();
