// 按上传记录 JSON 复制 Excel（后端 core/distribute.py）
// 两步：生成清单（只读） → 确认后复制。结果按人、按文件夹显示，三个数（JSON / 来源 / 目的地）对不上的标黄
(() => {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
  const api = '/api/distribute';
  const KEYS = { source: 'clear_handle_excel.distributeSource', dest: 'clear_handle_excel.distributeDest', plan: 'clear_handle_excel.distributePlan' };
  let latest = {}, submitting = false, loading = true, timer, pollError = '';
  let viewing = null;          // 正在看的历史记录（null = 看当前清单 / 最近一次结果）
  let filter = 'all';
  let wasRunning = false;

  const load = key => { try { return localStorage.getItem(key) || ''; } catch (_) { return ''; } };
  const save = (key, value) => { try { localStorage.setItem(key, value); } catch (_) {} };
  $('distSource').value = load(KEYS.source);
  $('distDest').value = load(KEYS.dest);
  $('distPlan').value = load(KEYS.plan);

  const size = n => n == null ? '-' : n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1048576).toFixed(1)} MB`;
  const badChecks = folder => (folder.checks || []).filter(c => !c.ok);
  const isBad = folder => !!(folder.error || badChecks(folder).length || (folder.files || []).some(f => f.status && f.status !== 'ok')
    || (folder.dest_extra || []).length);

  function controls() {
    const busy = loading || submitting || latest.running;
    for (const id of ['distSource', 'distDest', 'distPlan', 'distPick']) $(id).disabled = !!busy;
    $('distPreview').disabled = !!busy;
    const canStart = !busy && latest.preview && latest.preview.folders.some(f => f.files.length);
    $('distStart').disabled = !canStart;
    $('distPreview').textContent = submitting === 'preview' ? '正在生成清单…' : '📋 生成复制清单';
    $('distStart').textContent = latest.running ? '正在复制…' : '📥 按清单开始复制';
  }

  // 进度区：跑着显示进度条；结束后收起进度条，结果写成完成提示（有问题用黄色）
  function showProgress(state, title = '', text = '', warn = false) {
    const line = $('distProgress'), track = $('distProgressFill').parentElement;
    track.classList.toggle('is-done', state !== 'running');
    track.classList.toggle('is-paused', state !== 'running');
    line.classList.toggle('task-done-line', state === 'done');
    line.classList.toggle('is-warn', state === 'done' && warn);
    if (state !== 'done') { if (text) line.textContent = text; return; }
    const strong = document.createElement('strong');
    strong.textContent = title;
    line.replaceChildren(strong, document.createTextNode(text));
  }

  function nums(folder, applied) {
    const bad = new Set(badChecks(folder).flatMap(c => [c.a, c.b]));
    const one = (name, value) => `<span class="dist-num${bad.has(name) ? ' is-bad' : ''}">${name} ${value ?? '–'}</span>`;
    return `<span class="dist-nums">${one('JSON', folder.json_files ?? '没写')}${one('来源', folder.source_count)}${applied ? one('目的地', folder.dest_count) : ''}</span>`;
  }

  function fileLink(label, folder, rel) {
    return window.excelLinkHtml(esc(label), { folder, file: rel });
  }

  function folderCard(folder, data, applied) {
    const overwrite = folder.files.filter(f => f.action === 'overwrite');
    const fresh = folder.files.filter(f => f.action === 'new');
    const failed = folder.files.filter(f => f.status === 'failed');
    const notDone = applied ? folder.files.filter(f => !f.status) : [];
    const cls = folder.error ? ' is-err' : isBad(folder) ? ' is-bad' : '';
    const open = folder.error || isBad(folder) || (folder.overlaps || []).length;
    const parts = [];
    parts.push(`<summary><span class="dist-route">${esc(folder.source_rel)} → ${esc(folder.dest_rel)}<small>${folder.owner ? `作者 ${esc(folder.owner)} · ` : ''}${applied ? '复制' : '要复制'} ${applied ? folder.files.filter(f => f.status === 'ok').length : folder.files.length} 个（覆盖 ${overwrite.length}、新建 ${fresh.length}）</small></span>${nums(folder, applied)}</summary>`);
    if (folder.error) parts.push(`<p class="link-error">⛔ ${esc(folder.error)}，这个文件夹不复制。</p>`);
    for (const text of folder.overlaps || []) parts.push(`<p class="link-skip">⚠️ ${esc(text)}</p>`);
    const checks = folder.checks || [];
    if (checks.length) {
      parts.push('<ul class="dist-checks">' + checks.map(c => c.ok
        ? `<li class="good">✅ ${esc(c.pair)}：一致（${c.x} 个）</li>`
        : `<li class="bad">❌ ${esc(c.pair)}：${esc(c.text)}</li>`).join('') + '</ul>');
      if (checks.some(c => !c.ok && c.pair === 'JSON ↔ 来源') && folder.hint) parts.push(`<p class="dist-meta">💡 ${esc(folder.hint)}</p>`);
    }
    const extra = [];
    if ((folder.temp || []).length) extra.push(`~$ 临时文件 ${folder.temp.length} 个`);
    const other = Object.entries(folder.other || {});
    if (other.length) extra.push('非 Excel 文件 ' + other.map(([ext, n]) => `${ext} ${n} 个`).join('、'));
    if (extra.length) parts.push(`<p class="dist-meta">来源里另有（不复制、不算进来源数）：${esc(extra.join('，'))}</p>`);
    if ((folder.missing_in_dest || []).length) {
      parts.push(`<h4 class="dist-sub link-error">来源有、目的地没有（来源多 ${folder.missing_in_dest.length} 个）</h4><ul class="dist-list">`
        + folder.missing_in_dest.map(rel => `<li>${fileLink(rel, folder.source_path, rel)}</li>`).join('') + '</ul>');
    }
    if ((folder.dest_extra || []).length) {
      parts.push(`<h4 class="dist-sub link-skip">目的地原来就有、来源没有（目的地多 ${folder.dest_extra.length} 个，不会动它们）</h4><ul class="dist-list">`
        + folder.dest_extra.map(rel => `<li>${fileLink(rel, folder.dest_path, rel)}</li>`).join('') + '</ul>');
    }
    for (const err of folder.scan_errors || []) parts.push(`<p class="link-error">⚠️ 读不到：${esc(err)}</p>`);
    if (folder.files.length) {
      const rows = folder.files.map(f => {
        const action = f.action === 'overwrite'
          ? `<span class="tag-overwrite">🔁 覆盖</span>` : `<span class="tag-new">🆕 新建</span>`;
        const sizeText = f.action === 'overwrite' ? `${size(f.old_size)} → ${size(f.size)}` : size(f.size);
        let state = '';
        if (applied) state = f.status === 'ok' ? '<span class="link-ok">✅ 已复制</span>'
          : f.status === 'failed' ? `<span class="link-error">❌ ${esc(f.error)}</span>` : '<span class="link-skip">没有复制</span>';
        else if (f.source_open) state = '<span class="link-skip">⚠️ 来源正在 Excel 中打开，复制时会请你先关掉</span>';
        const base = applied && f.status === 'ok' ? folder.dest_path : folder.source_path;
        return `<tr><td>${fileLink(f.rel, base, applied && f.status === 'ok' ? f.dest_rel : f.rel)}</td><td>${action}</td><td>${sizeText}</td><td>${state}</td></tr>`;
      });
      parts.push(`<details${failed.length || notDone.length ? ' open' : ''}><summary class="dist-sub">文件清单 · ${folder.files.length} 个</summary><div class="link-table-wrap"><table class="link-table"><thead><tr><th>文件</th><th>操作</th><th>大小</th><th>${applied ? '结果' : '提示'}</th></tr></thead><tbody>${rows.join('')}</tbody></table></div></details>`);
    }
    parts.push(`<p class="dist-meta">来源：${esc(folder.source_path)}<br>目的地：${esc(folder.dest_path)}</p>`);
    return `<details class="link-file dist-folder${cls}"${open ? ' open' : ''}>${parts.join('')}</details>`;
  }

  function chips(data, applied) {
    const s = data.summary || {};
    const list = applied
      ? [['ok', `已复制 ${s['已复制'] ?? 0} 个`], ['info', `覆盖 ${s['覆盖'] ?? 0} · 新建 ${s['新建'] ?? 0}`],
         [s['没复制成功'] ? 'err' : '', `没复制成功 ${s['没复制成功']} 个`], [s['没有复制'] ? 'warn' : '', `没有复制 ${s['没有复制']} 个`],
         [s['数量不一致'] ? 'warn' : '', `数量不一致 ${s['数量不一致']} 个文件夹`], [s['有问题的文件夹'] ? 'err' : '', `找不到来源等 ${s['有问题的文件夹']} 个文件夹`]]
      : [['info', `${s['文件夹'] ?? 0} 个文件夹`], ['info', `要复制 ${s['要复制'] ?? 0} 个`], ['info', `覆盖 ${s['覆盖'] ?? 0} · 新建 ${s['新建'] ?? 0}`],
         [s['数量不一致'] ? 'warn' : '', `JSON 和来源数量不一致 ${s['数量不一致']} 个文件夹`],
         [s['目的地多出'] ? 'warn' : '', `目的地有来源没有的 Excel ${s['目的地多出']} 个文件夹`], [s['有问题'] ? 'err' : '', `有问题 ${s['有问题']} 个文件夹`]];
    return '<div class="dist-summary">' + list.filter(([cls]) => cls).map(([cls, text]) => `<span class="dist-chip dist-chip-${cls}">${esc(text)}</span>`).join('') + '</div>';
  }

  function alertBox(data, applied) {
    const items = [];
    if (data.stopped) items.push(data.stopped);
    for (const text of (data.stale || []).slice(0, 20)) items.push(text);
    for (const note of data.notes || []) items.push(note);
    for (const folder of data.folders || []) {
      const why = [];
      if (folder.error) why.push(folder.error);
      why.push(...badChecks(folder).map(c => c.text));
      const failed = folder.files.filter(f => f.status === 'failed').length;
      if (failed) why.push(`${failed} 个文件没复制成功`);
      if (!applied && (folder.dest_extra || []).length) why.push(`目的地原来就有 ${folder.dest_extra.length} 个来源没有的 Excel（不会动它们），复制后目的地会比来源多 ${folder.dest_extra.length} 个`);
      if (why.length) items.push(`${folder.dest_rel}：${why.join('；')}`);
    }
    if (!items.length) return '';
    const head = applied ? '有对不上的地方，请看下面标黄的文件夹，或下载复制记录查看：' : '生成清单时发现这些问题（不影响其他文件夹复制）：';
    return `<div class="dist-alert">⚠️ ${head}<ul>${items.map(t => `<li>${esc(t)}</li>`).join('')}</ul></div>`;
  }

  function renderData(data, applied) {
    const folders = data.folders || [];
    const groups = [['all', '全部', folders], ['bad', '对不上 / 有问题', folders.filter(isBad)],
      ['overwrite', '有覆盖', folders.filter(f => f.files.some(x => x.action === 'overwrite'))], ['ok', '没问题', folders.filter(f => !isBad(f))]];
    const shown = (groups.find(([key]) => key === filter) || groups[0])[2];
    const html = [chips(data, applied), alertBox(data, applied)];
    if (applied && data.id) html.push(`<p><a class="btn btn-outline btn-sm" href="${api}/markdown?id=${encodeURIComponent(data.id)}">⬇ 下载复制记录（Markdown）</a></p>`);
    html.push(`<p class="toolbar-hint">来源根目录：${esc(data.source_root)}　目的地根目录：${esc(data.dest_root)}</p>`);
    html.push('<div class="repair-filter-buttons" role="group" aria-label="文件夹筛选">' + groups.map(([key, label, list]) =>
      `<button type="button" class="btn ${key === filter ? 'btn-primary' : 'btn-outline'} btn-sm" data-dist-filter="${key}" aria-pressed="${key === filter}">${label}（${list.length}）</button>`).join('') + '</div>');
    let person = null;
    for (const folder of shown) {
      if (folder.person !== person) { person = folder.person; html.push(`<h3 class="dist-person">👤 ${esc(person)}</h3>`); }
      html.push(folderCard(folder, data, applied));
    }
    if (!shown.length) html.push('<p>这一类没有文件夹。</p>');
    return html.join('');
  }

  function renderHistory() {
    const runs = latest.history || [];
    if (!runs.length) { $('distHistory').innerHTML = '<p>还没有复制记录。</p>'; return; }
    $('distHistory').innerHTML = '<ul class="dist-history">' + runs.map(run => {
      const s = run.summary || {};
      const chip = run.attention ? '<span class="dist-chip dist-chip-warn">⚠️ 有对不上的地方</span>' : '<span class="dist-chip dist-chip-ok">✓ 全部一致</span>';
      return `<li><span class="dist-route">${esc(run.label || '（没有 label）')}<small>${esc(run.finished || '')} · 复制 ${s['已复制'] ?? 0} 个（覆盖 ${s['覆盖'] ?? 0}）· ${esc(run.dest_root || '')}</small></span>${chip}`
        + `<button type="button" class="btn btn-outline btn-sm" data-dist-view="${esc(run.id)}">查看</button>`
        + `<a class="btn btn-outline btn-sm" href="${api}/markdown?id=${encodeURIComponent(run.id)}">⬇ 下载</a></li>`;
    }).join('') + '</ul>';
  }

  function render() {
    controls();
    renderHistory();
    if (latest.running) {
      showProgress('running');
      $('distProgress').textContent = `正在复制 ${Math.min((latest.done || 0) + 1, latest.total || 0)} / ${latest.total || 0}：${latest.current || '准备中…'}`;
      $('distProgressFill').style.width = `${latest.total ? latest.done / latest.total * 100 : 0}%`;
      $('distResultTitle').textContent = '复制结果';
      $('distResults').textContent = '正在复制，完成后显示每个文件夹的核对结果。';
      return;
    }
    if (viewing) {
      $('distResultTitle').textContent = `复制记录：${viewing.label || ''}`;
      $('distResults').innerHTML = '<p><button type="button" class="btn btn-outline btn-sm" data-dist-back>← 回到当前清单</button></p>' + renderData(viewing, true);
      return;
    }
    const applied = latest.applied;
    const data = latest.preview || applied;
    if (!data) return;
    $('distProgressFill').style.width = '100%';
    if (latest.preview) {
      const s = data.summary;
      $('distResultTitle').textContent = `复制清单${data.label ? '：' + data.label : ''}`;
      showProgress('done', '清单已生成：', `${s['文件夹']} 个文件夹，要复制 ${s['要复制']} 个 Excel（覆盖 ${s['覆盖']}、新建 ${s['新建']}）。核对无误后点「📥 按清单开始复制」。`,
        !!(s['数量不一致'] || s['有问题'] || s['目的地多出'] || data.notes.length));
    } else {
      const s = applied.summary || {};
      $('distResultTitle').textContent = `复制结果${applied.label ? '：' + applied.label : ''}`;
      const warn = !!(applied.stopped || s['没复制成功'] || s['没有复制'] || s['数量不一致'] || s['有问题的文件夹'] || (applied.notes || []).length);
      showProgress('done', applied.stopped ? '没有复制：' : warn ? '复制结束，有对不上的地方：' : '复制完成：',
        applied.stopped ? applied.stopped : `已复制 ${s['已复制']} 个（覆盖 ${s['覆盖']}、新建 ${s['新建']}）`
          + (warn ? '。请看下面标黄的文件夹，或下载复制记录。' : '，每个文件夹三个数都一致。'), warn);
    }
    $('distResults').innerHTML = renderData(data, !latest.preview);
  }

  async function refresh() {
    clearTimeout(timer);
    let failed = false;
    try {
      const response = await fetch(api, { cache: 'no-store' });
      if (!response.ok) throw Error('读取任务进度失败，正在重试');
      latest = await response.json();
      if (pollError && $('distError').textContent === pollError) $('distError').textContent = '';
      pollError = '';
      loading = false;
      if (wasRunning && !latest.running && latest.applied && !latest.applied.stopped) {
        const s = latest.applied.summary || {};
        if (s['没复制成功'] || s['没有复制'] || s['数量不一致'] || s['有问题的文件夹']) {
          window.appAlert(`复制结束，但有 ${s['数量不一致'] || 0} 个文件夹数量对不上、${(s['没复制成功'] || 0) + (s['没有复制'] || 0)} 个文件没复制成功。\n请看页面上标黄的文件夹，或下载复制记录查看是哪些文件、哪边多了。`, { title: '请看复制记录' });
        }
      }
      wasRunning = !!latest.running;
      render();
    } catch (error) { failed = true; pollError = error.message; $('distError').textContent = pollError; }
    if (latest.running || failed) timer = setTimeout(refresh, failed ? 1800 : 900);
  }

  async function post(path, payload) {
    const response = await fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload || {}) });
    const result = await response.json();
    if (!response.ok || !result.ok) throw Error(result.error || '操作失败');
    return result;
  }

  async function makePreview() {
    if (loading || submitting || latest.running) return;
    const payload = { source_root: $('distSource').value.trim(), dest_root: $('distDest').value.trim(), plan: $('distPlan').value };
    $('distError').textContent = '';
    if (!payload.source_root || !payload.dest_root) { $('distError').textContent = '请填写来源根目录和目的地根目录'; return; }
    if (!payload.plan.trim()) { $('distError').textContent = '请粘贴上传记录 JSON，或选择 .json 文件'; return; }
    save(KEYS.source, payload.source_root); save(KEYS.dest, payload.dest_root); save(KEYS.plan, payload.plan);
    submitting = 'preview'; viewing = null; filter = 'all'; controls();
    showProgress('running', '', '正在列出来源和目的地的文件…');
    try {
      await post(api + '/preview', payload);
    } catch (error) {
      $('distError').textContent = error.message;
    } finally { submitting = false; await refresh(); }
  }

  async function startCopy() {
    if (loading || submitting || latest.running || !latest.preview) return;
    const s = latest.preview.summary;
    const lines = [`将复制 ${s['要复制']} 个 Excel 到「${latest.preview.dest_root}」。`,
      s['覆盖'] ? `其中 ${s['覆盖']} 个会覆盖目的地的同名文件（文件名不变，复制记录里会写明）。` : '目的地没有同名文件，不会覆盖。',
      '来源文件只复制、不移动。'];
    if (s['数量不一致'] || s['有问题'] || s['目的地多出']) lines.push(`注意：有 ${s['数量不一致']} 个文件夹 JSON 和来源数量对不上、${s['目的地多出']} 个文件夹目的地有来源没有的 Excel、${s['有问题']} 个文件夹有问题，复制后都会写进记录。`);
    const ok = await window.appConfirm(lines.join('\n'), { title: '按清单开始复制', okText: '开始复制', tone: s['覆盖'] ? 'danger' : undefined });
    if (!ok) return;
    submitting = 'start'; viewing = null; controls();
    try {
      await post(api + '/start', {});
      filter = 'all';
      wasRunning = true;
    } catch (error) {
      $('distError').textContent = error.message;
    } finally { submitting = false; await refresh(); }
  }

  async function viewRun(id) {
    try {
      const response = await fetch(`${api}/run?id=${encodeURIComponent(id)}`, { cache: 'no-store' });
      const result = await response.json();
      if (!result.ok) throw Error(result.error || '读取失败');
      viewing = result.run; filter = 'all'; render();
      $('distResultTitle').scrollIntoView({ block: 'start' });
    } catch (error) { $('distError').textContent = error.message; }
  }

  $('distPick').addEventListener('click', () => $('distFile').click());
  $('distFile').addEventListener('change', async () => {
    const file = $('distFile').files[0];
    if (!file) return;
    $('distPlan').value = await file.text();
    $('distFileName').textContent = `已读入 ${file.name}`;
    $('distFile').value = '';
  });
  $('distPreview').addEventListener('click', makePreview);
  $('distStart').addEventListener('click', startCopy);
  document.addEventListener('click', event => {
    const button = event.target.closest('[data-dist-filter]');
    if (button) { filter = button.dataset.distFilter; render(); $('distResults').querySelector('[aria-pressed="true"]')?.focus(); return; }
    const view = event.target.closest('[data-dist-view]');
    if (view) { viewRun(view.dataset.distView); return; }
    if (event.target.closest('[data-dist-back]')) { viewing = null; filter = 'all'; render(); }
  });
  controls();
  refresh();
})();
