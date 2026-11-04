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
  // 当前任务（主页「🆕 新任务」）：有任务时按任务的处理 / 共享文件夹，不用填 JSON；?mode=json 改回 JSON 复制
  const params = new URLSearchParams(location.search);
  let task = null;
  const jsonMode = params.get('mode') === 'json' || params.has('plan');
  document.querySelector('[data-dist-mode="' + (jsonMode ? 'json' : 'task') + '"]')?.setAttribute('aria-current', 'page');
  let needsPreview = params.has('plan');

  const load = key => { try { return localStorage.getItem(key) || ''; } catch (_) { return ''; } };
  const save = (key, value) => { try { localStorage.setItem(key, value); } catch (_) {} };
  $('distSource').value = load(KEYS.source);
  $('distDest').value = load(KEYS.dest);
  $('distPlan').value = load(KEYS.plan);

  const size = n => n == null ? '-' : n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1048576).toFixed(1)} MB`;
  const badChecks = folder => (folder.checks || []).filter(c => !c.ok);
  const conflicts = folder => (folder.files || []).filter(f => f.conflict || f.status === 'conflict');
  const isFetch = data => data && data.kind === 'fetch';
  const isBad = folder => !!(folder.error || conflicts(folder).length || (folder.dir_errors || []).length || badChecks(folder).length || (folder.files || []).some(f => f.status && f.status !== 'ok')
    || (folder.dest_extra || []).length);

  function controls() {
    const busy = loading || submitting || latest.running;
    for (const id of ['distSource', 'distDest', 'distPlan', 'distPick']) $(id).disabled = !!busy;
    $('distPreview').disabled = !!busy;
    $('distFetch').disabled = !!busy;
    const fetchMode = isFetch(latest.preview) || (latest.running && latest.kind === 'fetch');
    const canStart = !busy && !needsPreview && latest.preview && latest.preview.folders.some(f => f.files.length || (f.dirs_to_create || []).length);
    $('distStart').disabled = !canStart;
    $('distPreview').textContent = submitting === 'preview' ? '正在生成清单…' : task ? '📋 生成上传清单' : '📋 生成复制清单';
    $('distFetch').textContent = submitting === 'fetch' ? '正在对比共享和本地…' : task ? '⬇ 从共享拿最新版' : '⬇ 从目的地拿最新版';
    $('distStart').textContent = fetchMode ? (latest.running ? '正在拿…' : '⬇ 按清单拿文件')
      : (latest.running ? (task ? '正在上传…' : '正在复制…') : (task ? '📥 按清单上传到共享' : '📥 按清单开始复制'));
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
    const willCopy = folder.files.filter(f => !f.conflict);
    const overwrite = willCopy.filter(f => f.action === 'overwrite');
    const fresh = willCopy.filter(f => f.action === 'new');
    const failed = folder.files.filter(f => f.status === 'failed');
    const notDone = applied ? folder.files.filter(f => !f.status) : [];
    const clash = conflicts(folder);
    const cls = folder.error ? ' is-err' : clash.length ? ' is-conflict' : isBad(folder) ? ' is-bad' : '';
    const open = folder.error || isBad(folder) || (folder.overlaps || []).length;
    const parts = [];
    parts.push(`<summary><span class="dist-route">${esc(folder.source_rel)} → ${esc(folder.dest_rel)}<small>${folder.owner ? `作者 ${esc(folder.owner)} · ` : ''}${applied ? '复制' : '要复制'} ${applied ? folder.files.filter(f => f.status === 'ok').length : willCopy.length} 个（覆盖 ${overwrite.length}、新建 ${fresh.length}）</small></span>${nums(folder, applied)}</summary>`);
    const madeTop = applied ? (folder.created_dirs || []).includes('.') : (folder.dirs_to_create || []).includes('.');
    if (folder.error) parts.push(`<p class="link-error">⛔ ${esc(folder.error)}，这个文件夹不复制${madeTop ? (applied ? '，目的地已建成空文件夹' : '，只在目的地建一个空文件夹') : ''}。</p>`);
    const subDirs = (applied ? folder.created_dirs : folder.dirs_to_create || []).filter(d => d !== '.');
    if (!folder.error && (madeTop || subDirs.length)) {
      parts.push(`<p class="dist-meta">📁 ${applied ? '新建了' : '要新建'}${madeTop ? '目的地文件夹' : ''}${madeTop && subDirs.length ? '和' : ''}${subDirs.length ? `子文件夹 ${subDirs.length} 个：${esc(subDirs.join('、'))}` : ''}</p>`);
    }
    for (const err of folder.dir_errors || []) parts.push(`<p class="link-error">❌ 文件夹没建成：${esc(err)}</p>`);
    if (clash.length) {
      parts.push(`<h4 class="dist-sub dist-conflict">🛑 别人改过，${applied ? '没有覆盖' : '不会覆盖'}（${clash.length} 个）：请先「⬇ 从目的地拿最新版」，在新版上改好再复制</h4><ul class="dist-list">`
        + clash.map(f => `<li>${fileLink(f.dest_rel, folder.dest_path, f.dest_rel)}${f.guard === 'changed' && !/读不到/.test(f.conflict || '') ? '' : `　<span class="dist-conflict">${esc(f.conflict || f.error)}</span>`}</li>`).join('') + '</ul>');
    }
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
          : f.status === 'conflict' ? '<span class="dist-conflict">🛑 别人改过，没有覆盖</span>'
          : f.status === 'failed' ? `<span class="link-error">❌ ${esc(f.error)}</span>` : '<span class="link-skip">没有复制</span>';
        else if (f.conflict) state = '<span class="dist-conflict">🛑 别人改过，不覆盖</span>';
        else if (f.source_open) state = '<span class="link-skip">⚠️ 来源正在 Excel 中打开，复制时会请你先关掉</span>';
        else if (f.guard === 'same') state = '<span class="dist-guard-ok">✓ 拿走后没人改过</span>';
        else if (f.guard === 'identical') state = '<span class="dist-guard-ok">✓ 目的地已经和你的一样</span>';
        else if (f.guard === 'unknown') state = '<span class="dist-guard-unknown">没有拿取记录，没法检查别人改过没有</span>';
        const base = applied && f.status === 'ok' ? folder.dest_path : folder.source_path;
        return `<tr${f.conflict || f.status === 'conflict' ? ' class="is-conflict"' : ''}><td>${fileLink(f.rel, base, applied && f.status === 'ok' ? f.dest_rel : f.rel)}</td><td>${action}</td><td>${sizeText}</td><td>${state}</td></tr>`;
      });
      parts.push(`<details${failed.length || notDone.length || clash.length ? ' open' : ''}><summary class="dist-sub">文件清单 · ${folder.files.length} 个</summary><div class="link-table-wrap"><table class="link-table"><thead><tr><th>文件</th><th>操作</th><th>大小</th><th>${applied ? '结果' : '提示'}</th></tr></thead><tbody>${rows.join('')}</tbody></table></div></details>`);
    }
    parts.push(`<p class="dist-meta">来源：${esc(folder.source_path)}<br>目的地：${esc(folder.dest_path)}</p>`);
    return `<details class="link-file dist-folder${cls}"${open ? ' open' : ''}>${parts.join('')}</details>`;
  }

  function chips(data, applied) {
    const s = data.summary || {};
    const list = applied
      ? [['ok', `已复制 ${s['已复制'] ?? 0} 个`], ['info', `覆盖 ${s['覆盖'] ?? 0} · 新建 ${s['新建'] ?? 0}`],
         ['info', `新建文件夹 ${s['新建文件夹'] ?? 0} 个`],
         [s['别人改过没覆盖'] ? 'err' : '', `🛑 别人改过没覆盖 ${s['别人改过没覆盖']} 个`],
         [s['没复制成功'] ? 'err' : '', `没复制成功 ${s['没复制成功']} 个`], [s['没有复制'] ? 'warn' : '', `没有复制 ${s['没有复制']} 个`],
         [s['数量不一致'] ? 'warn' : '', `数量不一致 ${s['数量不一致']} 个文件夹`], [s['有问题的文件夹'] ? 'err' : '', `找不到来源等 ${s['有问题的文件夹']} 个文件夹`]]
      : [['info', `${s['文件夹'] ?? 0} 个文件夹`], ['info', `要复制 ${s['要复制'] ?? 0} 个`], ['info', `覆盖 ${s['覆盖'] ?? 0} · 新建 ${s['新建'] ?? 0}`],
         ['info', `要新建文件夹 ${s['新建文件夹'] ?? 0} 个`],
         [s['别人改过'] ? 'err' : '', `🛑 别人改过、不覆盖 ${s['别人改过']} 个`],
         [s['没有版本记录'] ? 'warn' : '', `没有拿取记录、没法检查 ${s['没有版本记录']} 个`],
         [s['数量不一致'] ? 'warn' : '', `JSON 和来源数量不一致 ${s['数量不一致']} 个文件夹`],
         [s['目的地多出'] ? 'warn' : '', `目的地有来源没有的 Excel ${s['目的地多出']} 个文件夹`], [s['有问题'] ? 'err' : '', `有问题 ${s['有问题']} 个文件夹`]];
    return '<div class="dist-summary">' + list.filter(([cls]) => cls).map(([cls, text]) => `<span class="dist-chip dist-chip-${cls}">${esc(text)}</span>`).join('') + '</div>';
  }

  function alertBox(data, applied) {
    const clash = (data.folders || []).flatMap(folder => conflicts(folder).map(f => `${folder.dest_rel}/${f.dest_rel}`));
    const red = clash.length ? `<div class="dist-alert is-conflict">🛑 ${clash.length} 个文件在你拿走之后被别人改过，${applied ? '没有覆盖' : '不会覆盖'}，别人的修改不会被盖掉。请点「⬇ 从目的地拿最新版」拿回新版（你改的那份会先备份），在新版上改好再复制：<ul>${clash.slice(0, 30).map(t => `<li>${esc(t)}</li>`).join('')}${clash.length > 30 ? `<li>……另外 ${clash.length - 30} 个见下面红框的文件夹</li>` : ''}</ul></div>` : '';
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
      if ((folder.dir_errors || []).length) why.push(`${folder.dir_errors.length} 个文件夹没建成`);
      if (!applied && (folder.dest_extra || []).length) why.push(`目的地原来就有 ${folder.dest_extra.length} 个来源没有的 Excel（不会动它们），复制后目的地会比来源多 ${folder.dest_extra.length} 个`);
      if (why.length) items.push(`${folder.dest_rel}：${why.join('；')}`);
    }
    if (!items.length) return red;
    const head = applied ? '有对不上的地方，请看下面标黄的文件夹，或下载复制记录查看：' : '生成清单时发现这些问题（不影响其他文件夹复制）：';
    return red + `<div class="dist-alert">⚠️ ${head}<ul>${items.map(t => `<li>${esc(t)}</li>`).join('')}</ul></div>`;
  }

  function renderData(data, applied) {
    const folders = data.folders || [];
    const groups = [['all', '全部', folders], ['bad', '对不上 / 有问题', folders.filter(isBad)],
      ...(folders.some(f => conflicts(f).length) ? [['conflict', '🛑 别人改过', folders.filter(f => conflicts(f).length)]] : []),
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

  // ───── 拿最新版（目的地 → 来源）的清单和结果 ─────
  const FETCH_ACTION = {
    copy: '<span class="tag-new">🆕 新拿</span>', replace: '<span class="tag-overwrite">🔁 替换本地</span>',
    keep: '<span class="tag-keep">✋ 保留你改的</span>', same: '<span class="tag-same">＝ 已一样</span>',
  };
  const FETCH_HINT = {
    copy: '本地没有，拿过来', replace: '目的地有新版，用它替换本地的', keep: '你改过、目的地没人动，不拿', same: '本地和目的地一样，只记下版本',
  };

  function fetchCard(folder, applied) {
    const replaced = folder.files.some(f => f.action === 'replace');
    const failed = folder.files.filter(f => f.status === 'failed');
    const take = folder.files.filter(f => f.action === 'copy' || f.action === 'replace');
    const cls = folder.error ? ' is-err' : failed.length ? ' is-bad' : '';
    const parts = [`<summary><span class="dist-route">${esc(folder.dest_rel)} → ${esc(folder.source_rel)}<small>${applied ? '拿了' : '要拿'} ${applied ? take.filter(f => f.status === 'ok').length : take.length} 个 · 保留你改的 ${folder.files.filter(f => f.action === 'keep').length} 个 · 已一样 ${folder.files.filter(f => f.action === 'same').length} 个</small></span></summary>`];
    if (folder.error) parts.push(`<p class="link-error">⛔ ${esc(folder.error)}</p>`);
    if ((folder.local_extra || []).length) {
      parts.push(`<h4 class="dist-sub link-skip">本地有、目的地没有（${folder.local_extra.length} 个，不会动）</h4><ul class="dist-list">`
        + folder.local_extra.map(rel => `<li>${fileLink(rel, folder.source_path, rel)}</li>`).join('') + '</ul>');
    }
    for (const err of folder.scan_errors || []) parts.push(`<p class="link-error">⚠️ 读不到：${esc(err)}</p>`);
    if (folder.files.length) {
      const rows = folder.files.map(f => {
        let state = esc(FETCH_HINT[f.action] || '');
        if (!applied && f.backup) state += '（你这份先备份）';
        if (!applied && f.shared_open) state += '<br><span class="link-skip">⚠️ 目的地这份有人正在 Excel 里打开，可能马上又会改</span>';
        if (applied) state = f.status === 'ok' ? `<span class="link-ok">✅</span>${f.backup_path ? ` 你的那份备份在 ${esc(f.backup_path)}` : ''}`
          : f.status === 'failed' ? `<span class="link-error">❌ ${esc(f.error)}</span>` : '<span class="link-skip">没有拿</span>';
        const base = applied ? folder.source_path : folder.dest_path;
        return `<tr><td>${fileLink(f.rel, base, applied ? f.local_rel : f.rel)}</td><td>${FETCH_ACTION[f.action] || esc(f.action)}</td><td>${size(f.size)}</td><td>${state}</td></tr>`;
      });
      parts.push(`<details${failed.length || replaced ? ' open' : ''}><summary class="dist-sub">文件清单 · ${folder.files.length} 个</summary><div class="link-table-wrap"><table class="link-table"><thead><tr><th>文件</th><th>做法</th><th>大小</th><th>${applied ? '结果' : '说明'}</th></tr></thead><tbody>${rows.join('')}</tbody></table></div></details>`);
    }
    parts.push(`<p class="dist-meta">目的地：${esc(folder.dest_path)}<br>拿回到：${esc(folder.source_path)}</p>`);
    return `<details class="link-file dist-folder${cls}"${folder.error || failed.length || replaced ? ' open' : ''}>${parts.join('')}</details>`;
  }

  function renderFetch(data, applied) {
    const s = data.summary || {};
    const list = applied
      ? [['ok', `已拿 ${s['已拿'] ?? 0} 个`], ['info', `新拿 ${s['新拿'] ?? 0} · 替换 ${s['替换'] ?? 0}`], ['info', `备份你的 ${s['备份'] ?? 0} 个`],
         ['info', `保留你改的 ${s['保留你改的'] ?? 0} 个`], ['info', `已一样 ${s['已一样'] ?? 0} 个`], ['ok', `记下版本 ${s['记下版本'] ?? 0} 个`],
         [s['没拿成功'] ? 'err' : '', `没拿成功 ${s['没拿成功']} 个`], [s['没有拿'] ? 'warn' : '', `没有拿 ${s['没有拿']} 个`],
         [s['有问题的文件夹'] ? 'err' : '', `找不到目的地等 ${s['有问题的文件夹']} 个文件夹`]]
      : [['info', `${s['文件夹'] ?? 0} 个文件夹`], ['info', `要拿 ${s['要拿'] ?? 0} 个（新拿 ${s['新拿'] ?? 0} · 替换 ${s['替换'] ?? 0}）`],
         [s['要备份'] ? 'warn' : '', `替换前备份你的 ${s['要备份']} 个`], ['info', `保留你改的 ${s['保留你改的'] ?? 0} 个`],
         ['info', `已一样 ${s['已一样'] ?? 0} 个`], [s['有问题'] ? 'err' : '', `有问题 ${s['有问题']} 个文件夹`]];
    const html = ['<div class="dist-summary">' + list.filter(([cls]) => cls).map(([cls, text]) => `<span class="dist-chip dist-chip-${cls}">${esc(text)}</span>`).join('') + '</div>'];
    if (applied && data.backup_root) html.push(`<p class="toolbar-hint">本地被替换的文件备份在：${esc(data.backup_root)}</p>`);
    if (applied && data.id) html.push(`<p><a class="btn btn-outline btn-sm" href="${api}/markdown?id=${encodeURIComponent(data.id)}">⬇ 下载拿取记录（Markdown）</a></p>`);
    html.push(`<p class="toolbar-hint">从目的地根目录：${esc(data.dest_root)}　拿回来源根目录：${esc(data.source_root)}</p>`);
    let person = null;
    for (const folder of data.folders || []) {
      if (folder.person !== person) { person = folder.person; html.push(`<h3 class="dist-person">👤 ${esc(person)}</h3>`); }
      html.push(fetchCard(folder, applied));
    }
    return html.join('');
  }

  const renderAny = (data, applied) => isFetch(data) ? renderFetch(data, applied) : renderData(data, applied);

  function renderHistory() {
    const runs = latest.history || [];
    if (!runs.length) { $('distHistory').innerHTML = '<p>还没有复制记录。</p>'; return; }
    $('distHistory').innerHTML = '<ul class="dist-history">' + runs.map(run => {
      const s = run.summary || {};
      const chip = run.attention ? (s['别人改过没覆盖'] ? '<span class="dist-chip dist-chip-err">🛑 有别人改过的</span>' : '<span class="dist-chip dist-chip-warn">⚠️ 有对不上的地方</span>')
        : '<span class="dist-chip dist-chip-ok">✓ 全部一致</span>';
      const what = run.kind === 'fetch' ? `⬇ 拿最新版 ${s['已拿'] ?? 0} 个（替换 ${s['替换'] ?? 0}）`
        : `复制 ${s['已复制'] ?? 0} 个（覆盖 ${s['覆盖'] ?? 0}${s['别人改过没覆盖'] ? `，别人改过没覆盖 ${s['别人改过没覆盖']}` : ''}）`;
      return `<li><span class="dist-route">${esc(run.label || '（没有 label）')}<small>${esc(run.finished || '')} · ${what} · ${esc(run.dest_root || '')}</small></span>${chip}`
        + `<button type="button" class="btn btn-outline btn-sm" data-dist-view="${esc(run.id)}">查看</button>`
        + `<a class="btn btn-outline btn-sm" href="${api}/markdown?id=${encodeURIComponent(run.id)}">⬇ 下载</a></li>`;
    }).join('') + '</ul>';
  }

  function render() {
    controls();
    renderHistory();
    if (latest.running) {
      showProgress('running');
      $('distProgress').textContent = `${latest.kind === 'fetch' ? '正在拿最新版' : '正在复制'} ${Math.min((latest.done || 0) + 1, latest.total || 0)} / ${latest.total || 0}：${latest.current || '准备中…'}`;
      $('distProgressFill').style.width = `${latest.total ? latest.done / latest.total * 100 : 0}%`;
      window.ProgressEta?.update('distProgressFill', `dist:${latest.kind}`, latest.total ? latest.done / latest.total : 0, true);
      $('distResultTitle').textContent = latest.kind === 'fetch' ? '拿取结果' : '复制结果';
      $('distResults').textContent = latest.kind === 'fetch' ? '正在拿最新版，完成后显示每个文件夹的结果。' : '正在复制，完成后显示每个文件夹的核对结果。';
      return;
    }
    window.ProgressEta?.update('distProgressFill', `dist:${latest.kind}`, 0, false);
    if (viewing) {
      $('distResultTitle').textContent = `${isFetch(viewing) ? '拿取记录' : '复制记录'}：${viewing.label || ''}`;
      $('distResults').innerHTML = '<p><button type="button" class="btn btn-outline btn-sm" data-dist-back>← 回到当前清单</button></p>' + renderAny(viewing, true);
      return;
    }
    const applied = latest.applied;
    const data = latest.preview || applied;
    if (!data) return;
    $('distProgressFill').style.width = '100%';
    if (isFetch(data)) {
      const s = data.summary || {};
      if (latest.preview) {
        $('distResultTitle').textContent = `拿取清单${data.label ? '：' + data.label : ''}`;
        showProgress('done', '拿取清单已生成：', `要从目的地拿 ${s['要拿']} 个（新拿 ${s['新拿']}、替换本地 ${s['替换']}，其中 ${s['要备份']} 个先备份你的），保留你改的 ${s['保留你改的']} 个，已一样 ${s['已一样']} 个。核对后点「⬇ 按清单拿文件」。`, !!s['有问题']);
      } else {
        const warn = !!(data.stopped || s['没拿成功'] || s['没有拿'] || s['有问题的文件夹']);
        $('distResultTitle').textContent = `拿取结果${data.label ? '：' + data.label : ''}`;
        showProgress('done', data.stopped ? '没有拿：' : warn ? '拿完了，有没拿成功的：' : '拿完了：',
          data.stopped || `拿了 ${s['已拿'] ?? 0} 个（替换 ${s['替换'] ?? 0}、备份你的 ${s['备份'] ?? 0}），记下 ${s['记下版本'] ?? 0} 个文件的版本，复制前会用它检查别人改过没有。`, warn);
      }
      $('distResults').innerHTML = renderFetch(data, !latest.preview);
      return;
    }
    if (latest.preview) {
      const s = data.summary;
      $('distResultTitle').textContent = `复制清单${data.label ? '：' + data.label : ''}`;
      showProgress('done', '清单已生成：', `${s['文件夹']} 个文件夹，要新建文件夹 ${s['新建文件夹']} 个，要复制 ${s['要复制']} 个 Excel（覆盖 ${s['覆盖']}、新建 ${s['新建']}）`
        + (s['别人改过'] ? `；🛑 ${s['别人改过']} 个别人改过，不会覆盖` : '') + '。核对无误后点「📥 按清单开始复制」。',
        !!(s['别人改过'] || s['数量不一致'] || s['有问题'] || s['目的地多出'] || data.notes.length));
    } else {
      const s = applied.summary || {};
      $('distResultTitle').textContent = `复制结果${applied.label ? '：' + applied.label : ''}`;
      const warn = !!(applied.stopped || s['没复制成功'] || s['没有复制'] || s['数量不一致'] || s['有问题的文件夹'] || (applied.notes || []).length);
      showProgress('done', applied.stopped ? '没有复制：' : warn ? '复制结束，有对不上的地方：' : '复制完成：',
        applied.stopped ? applied.stopped : `新建文件夹 ${s['新建文件夹'] ?? 0} 个，已复制 ${s['已复制']} 个（覆盖 ${s['覆盖']}、新建 ${s['新建']}）`
          + (s['别人改过没覆盖'] ? `；🛑 ${s['别人改过没覆盖']} 个别人改过，没有覆盖` : '')
          + (warn ? '。请看下面标红、标黄的文件夹，或下载复制记录。' : '，每个文件夹三个数都一致。'), warn);
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
      if (wasRunning && !latest.running && latest.applied && !latest.applied.stopped && !isFetch(latest.applied)) {
        const s = latest.applied.summary || {};
        if (s['别人改过没覆盖']) {
          window.appAlert(`有 ${s['别人改过没覆盖']} 个文件在你拿走之后被别人改过，没有覆盖（别人的修改还在）。\n请点「⬇ 从目的地拿最新版」拿回新版（你改的那份会先备份），在新版上改好再复制。\n其他文件已照常复制。`, { title: '有文件被别人改过', tone: 'warn' });
        } else if (s['没复制成功'] || s['没有复制'] || s['数量不一致'] || s['有问题的文件夹']) {
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

  async function makePreview(kind = 'preview') {
    if (loading || submitting || latest.running) return;
    const payload = { source_root: $('distSource').value.trim(), dest_root: $('distDest').value.trim(), plan: $('distPlan').value };
    $('distError').textContent = '';
    if (task) payload.task_id = task.id;
    else if (!payload.source_root || !payload.dest_root) { $('distError').textContent = '请填写来源根目录和目的地根目录'; return; }
    else if (!payload.plan.trim()) { $('distError').textContent = '请粘贴上传记录 JSON，或选择 .json 文件'; return; }
    if (!task) save(KEYS.source, payload.source_root); save(KEYS.dest, payload.dest_root); save(KEYS.plan, payload.plan);
    submitting = kind; viewing = null; filter = 'all'; controls();
    showProgress('running', '', kind === 'fetch' ? '正在对比目的地和本地的文件（要读每个文件算哈希，文件多时要等一会儿）…' : '正在列出来源和目的地的文件，检查要覆盖的文件别人改过没有…');
    try {
      await post(api + (kind === 'fetch' ? '/fetch_preview' : '/preview'), payload);
      needsPreview = false;
    } catch (error) {
      $('distError').textContent = error.message;
    } finally { submitting = false; await refresh(); }
  }

  async function startFetch() {
    const s = latest.preview.summary;
    const lines = [`从「${latest.preview.dest_root}」拿 ${s['要拿']} 个 Excel 回「${latest.preview.source_root}」：新拿 ${s['新拿']} 个，替换本地 ${s['替换']} 个。`];
    if (s['要备份']) lines.push(`替换的文件里有 ${s['要备份']} 个你本地改过，会先备份到${task ? '工作根目录下的「拿最新版前的备份」' : '来源根目录旁边的「…_拿最新版前的备份」'}文件夹，再换成目的地的新版。`);
    if (task) lines.push('拿下来的原件同时放一份到参考文件夹（修复颜色 / 引用 / 隐藏页对照用）。');
    if (s['保留你改的']) lines.push(`${s['保留你改的']} 个你改过、目的地没人动，不拿，保留你改的。`);
    lines.push('所有文件都会记下拿到的版本，之后复制（覆盖）前用来检查别人改过没有。目的地的文件不会被改动。');
    return window.appConfirm(lines.join('\n'), { title: '按清单拿最新版', okText: '开始拿', tone: s['替换'] ? 'danger' : undefined });
  }

  async function startCopy() {
    if (loading || submitting || latest.running || needsPreview || !latest.preview) return;
    if (isFetch(latest.preview)) {
      if (!await startFetch()) return;
      submitting = 'start'; viewing = null; controls();
      try { await post(api + '/start', {}); wasRunning = true; } catch (error) { $('distError').textContent = error.message; }
      finally { submitting = false; await refresh(); }
      return;
    }
    const s = latest.preview.summary;
    const lines = [`先按目录结构新建 ${s['新建文件夹']} 个文件夹（包括空文件夹），再复制 ${s['要复制']} 个 Excel 到「${latest.preview.dest_root}」。`,
      s['覆盖'] ? `其中 ${s['覆盖']} 个会覆盖目的地的同名文件（文件名不变，复制记录里会写明）。` : '目的地没有同名文件，不会覆盖。',
      '来源文件只复制、不移动。'];
    if (s['别人改过']) lines.push(`🛑 ${s['别人改过']} 个文件在你拿走之后被别人改过，不会覆盖（清单里标红）。`);
    if (s['没有版本记录']) lines.push(`${s['没有版本记录']} 个要覆盖的文件没有拿取记录，没法检查别人改过没有，会照旧覆盖。`);
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
    needsPreview = true;
    controls();
    $('distFileName').textContent = `已读入 ${file.name}`;
    $('distFile').value = '';
  });
  $('distPreview').addEventListener('click', () => makePreview('preview'));
  $('distFetch').addEventListener('click', () => makePreview('fetch'));
  $('distStart').addEventListener('click', startCopy);
  for (const id of ['distSource', 'distDest', 'distPlan']) {
    $(id).addEventListener('input', () => { needsPreview = true; controls(); });
  }
  document.addEventListener('click', event => {
    const button = event.target.closest('[data-dist-filter]');
    if (button) { filter = button.dataset.distFilter; render(); $('distResults').querySelector('[aria-pressed="true"]')?.focus(); return; }
    const view = event.target.closest('[data-dist-view]');
    if (view) { viewRun(view.dataset.distView); return; }
    if (event.target.closest('[data-dist-back]')) { viewing = null; filter = 'all'; render(); }
  });
  async function loadTask() {
    try {
      const result = await (await fetch('/api/tasks', { cache: 'no-store' })).json();
      const current = (result.tasks || []).find(t => t.id === result.current);
      if (current && !jsonMode) {
        task = current;
        $('distTaskName').textContent = current.name;
        $('distTaskPaths').innerHTML = `<dt>共享</dt><dd>${esc(current.shared)}</dd><dt>参考</dt><dd>${esc(current.reference)}</dd><dt>处理</dt><dd>${esc(current.work)}</dd>`;
        $('distTaskBox').hidden = false;
        $('distJsonFields').hidden = true;
        $('distJsonIntro').hidden = true;
      } else if (current) {
        $('distBackToTask').hidden = false;
      }
    } catch (_) {}
  }

  controls();
  loadTask().then(async () => {
    const planId = params.get('plan');
    if (planId) {
      try {
        const response = await fetch('/api/upload-manager/allocation-history/' + encodeURIComponent(planId) + '/manifest', { cache: 'no-store' });
        const manifest = await response.json();
        if (!response.ok) throw Error(manifest.error || '读取方案失败');
        $('distPlan').value = JSON.stringify(manifest, null, 2);
        $('distFileName').textContent = `已带入管理系统方案：${manifest.label || '上传方案'}。填写来源和目的地根目录后生成清单。`;
        save(KEYS.plan, $('distPlan').value);
        controls();
      } catch (error) {
        $('distError').textContent = error.message;
        // Failed handoff must not leave an older plan available for accidental use.
        $('distPlan').value = '';
        save(KEYS.plan, '');
        controls();
      }
    }
    await refresh();
    // 主页开了新任务过来：先自动生成拿取清单（确认后才真正拿）
    if (task && params.get('fetch') === '1' && !latest.running) {
      history.replaceState(null, '', '/distribute');
      makePreview('fetch');
    }
  });
})();
