// 修复图片（用户 2026-11-11 要）：先找出被行高列宽拉坏的截图、跑偏的红框（只读，有图片的页列出原版 / 新版 / PDF 三列整页图、
// 新版 Excel、对照打开原版、新版 PDF），看完再「修复这份」或「全部修复」。后端 core/image_restore.py
(() => {
  const $ = id => document.getElementById(id);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const API = '/api/image_restore';
  const FOLDER = 'clear_handle_excel.lastFolder';
  const OLD = 'clear_handle_excel.oldFolder';
  const PDF = 'clear_handle_excel.imgPdfFolder';
  const GROUP = 'clear_handle_excel.imgGroupOnFix';
  const shown = new Map();  // 文件名 → 已显示的那一块
  let listKey = '', onlyBad = false, hideDone = false;
  let latest = {}, lastSig = '', loading = true, submitting = false, timer = 0, initial = true, wasRunning = false;
  try {
    $('imgFolder').value = localStorage.getItem(FOLDER) || '';
    $('imgOldFolder').value = localStorage.getItem(OLD) || '';
    $('imgPdfFolder').value = localStorage.getItem(PDF) || '';
    $('imgGroupOnFix').checked = localStorage.getItem(GROUP) !== '0';
  } catch (_) {}

  const result = () => latest.result || null;
  const pending = () => (result()?.files || []).filter(f => !f.error && f.count && !f.fixed);

  function controls() {
    const busy = loading || submitting || latest.running;
    for (const id of ['imgFolder', 'imgOldFolder', 'imgPdfFolder', 'imgScan', 'imgGroup', 'imgGroupOnFix']) $(id).disabled = !!busy;
    const left = pending().length;
    $('imgFixAll').hidden = !result()?.scanned || !left;
    $('imgFixAll').disabled = !!busy;
    $('imgFixAll').textContent = `🖼️ 全部修复（${left} 份）`;
    $('imgScan').textContent = latest.running && latest.mode === 'scan' ? '正在找…' : '🔍 找出受影响的图片';
    document.querySelectorAll('[data-fix-file]').forEach(b => { b.disabled = !!busy; });
  }

  function renderProgress() {
    const total = latest.total || 0, done = latest.done || 0, running = !!latest.running;
    const fraction = total ? done / total : 0;
    $('imgFill').style.width = `${Math.round((running ? fraction : 1) * 100)}%`;
    $('imgTrack').hidden = !running;
    window.ProgressEta?.update($('imgFill'), 'image_restore_' + (latest.mode || 'scan'), fraction, running);
    const verb = { fix: '修复', group: '组合' }[latest.mode] || '查找';
    if (running) {
      $('imgProgress').textContent = `正在${verb} ${Math.min(done + 1, total)} / ${total} 份：${latest.current || '正在启动后台 Excel…'}` + (latest.step ? ` · ${latest.step}` : '');
      return;
    }
    const files = result()?.files || [];
    const bad = files.filter(f => f.error).length;
    const hit = files.filter(f => !f.error && f.count).length;
    const fixed = files.filter(f => f.fixed).length;
    const grouped = files.filter(f => f.grouped).length;
    const groupedText = grouped ? `，${grouped} 份已永久组合截图和标注` : '';
    $('imgProgress').textContent = !files.length ? '填写文件夹后点「找出受影响的图片」。'
      : !result()?.scanned ? `${bad ? '!' : '✓'} 处理了 ${files.length} 份${groupedText || '，没有要组合的截图'}` + (bad ? `，${bad} 份没处理` : '')
      : `${bad ? '!' : '✓'} 查了 ${files.length} 份：${hit} 份程序量出变化，${files.filter(f => !f.error && !f.count && f.review).length} 份有图片请目测` + (fixed ? `，已修复 ${fixed} 份` : '')
        + groupedText + (bad ? `，${bad} 份没查（看下面的原因）` : '');
  }

  function renderPanel() {
    if (!window.JobPanel || (!latest.running && !latest.result)) return;
    const log = latest.log || [];
    const verb = { fix: '修复图片', group: '组合截图和标注' }[latest.mode] || '找受影响的图片';
    if (latest.running) {
      window.JobPanel.run('image_restore', {
        title: `正在${verb}`, count: window.JobPanel.count(latest.done, latest.total, '份 Excel'),
        now: latest.current ? `📗 ${latest.current}` : '正在启动后台 Excel…', step: latest.step || '',
        fraction: latest.total ? latest.done / latest.total : 0, log, folder: latest.folder || '',
      });
    } else if (log.length) {
      const bad = log.filter(i => !i.ok).length;
      window.JobPanel.finish('image_restore', {
        title: `${verb}做完了${bad ? `，${bad} 份没处理` : ''}`, log, folder: latest.folder || '', warn: bad > 0,
        step: bad ? '黄色的是没处理的，原因写在后面；点文件名在 Excel 打开。' : '',
      });
    }
  }

  // 「复制路径」不经过后台：程序再忙、打开被拦，也能把路径贴到资源管理器 / Excel 里自己开（用户 2026-11-11 要「确保一定能打开」）
  const copyBtn = path => path ? ` <button class="img-copy" type="button" data-copy="${esc(path)}" title="复制这个路径，贴到资源管理器地址栏或 Excel 的打开框里">📋 复制路径</button>` : '';
  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); return true; } catch (_) {}
    const box = document.createElement('textarea');
    box.value = text; box.style.position = 'fixed'; box.style.opacity = '0';
    document.body.append(box); box.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (_) {}
    box.remove();
    return ok;
  }
  // pathHtml：这一列是哪个文件（完整路径，点了打开），用户 2026-11-11 要每页都能看到路径
  function shot(label, name, why, ratio, pathHtml = '') {
    return `<figure class="img-shot"><figcaption>${esc(label)}</figcaption><div class="img-path">${pathHtml}</div>`
      + (name ? `<img src="/api/batch/image/${encodeURIComponent(name)}?v=${result()?.stamp || 0}" alt="${esc(label)}" loading="lazy"${ratio ? ` style="aspect-ratio:${ratio}"` : ''}>`
        : `<div class="img-none"${ratio ? ` style="aspect-ratio:${ratio}"` : ''}>${esc(why || '没有')}</div>`) + '</figure>';
  }

  // 屏幕原样拍不了、退回导出效果时标出来（导出效果可能本身就有错位，不能当 Excel 原样看）
  const how = h => h === 'screen' ? '屏幕原样' : '导出效果，非屏幕原样';
  const groupTitle = f => (f.group_sheets || []).map(g => `【${g.sheet}】${g.groups.join('；')}`).join('\n');

  // 「无需处理」记在浏览器里（按 新版文件夹 + 文件 + 表 + 页），只改这一页的样子，不重画别的
  const MARKS = 'clear_handle_excel.imgNoNeed';
  let marks = {};
  try { marks = JSON.parse(localStorage.getItem(MARKS) || '{}') || {}; } catch (_) {}
  const markKey = (folder, file, sheet, page) => [folder, file, sheet, page].join('|');
  function applyMarks(root = $('imgResults')) {
    root.querySelectorAll('[data-mark]').forEach(el => {
      const on = !!marks[el.dataset.mark];
      el.classList.toggle('is-noneed', on);
      const button = el.querySelector('[data-mark-btn]');
      if (button) button.textContent = on ? '↩ 取消「无需处理」' : '✓ 无需处理';
    });
    const all = root.querySelectorAll('[data-mark]').length, done = root.querySelectorAll('[data-mark].is-noneed').length;
    if ($('imgMarkCount')) $('imgMarkCount').textContent = all ? `已标无需处理 ${done} / ${all} 页` : '';
  }

  // 一页一行三列整页图（用户 2026-11-11 要：看整页，能比对也能定位）。前两列是 Excel 屏幕原样（拍不了才退回导出效果）
  function renderFile(f, data) {
    const folder = data.folder;
    const openNew = (label, sheet = '', cell = '') => window.excelLinkHtml(label, { file: f.name, folder, sheet, cell });
    const compare = (label, sheet = '', cell = '') => window.excelLinkHtml(label, { file: f.name, folder, sheet, cell, compare: 'images' });
    const links = [
      openNew('📗 新版 Excel'),
      f.old ? compare('🆚 对照打开原版') : '',
      f.pdf ? `<a href="#" data-open-pdf="${esc(f.name)}" title="${esc(f.pdf)}">📄 新版 PDF</a>` : '<span title="Excel 旁边和 PDF 文件夹里都没找到同名 PDF">📄 没找到 PDF</span>',
    ].filter(Boolean).join('');
    const badge = f.fixed ? `<span class="img-badge">✓ 已修复 ${f.fixed_count || 0} 个图形</span>`
      : f.fix_error ? `<span class="img-badge is-bad">! 没修好：${esc(f.fix_error)}</span>`
      : f.count ? `<button class="btn btn-outline btn-sm" type="button" data-fix-file="${esc(f.name)}">🖼️ 修复这份</button>` : '';
    const groupBadge = f.grouped ? `<span class="img-badge is-group" title="${esc(groupTitle(f))}">📌 已组合 ${f.grouped} 组</span>` : '';
    const sheets = (f.sheets || []).map(s => {
      const details = (s.groups || []).flatMap(g => g.details || []);
      const cell = (s.groups || [])[0]?.cell || '';
      const head = `<div class="img-sheet">【${esc(s.sheet)}】${s.count ? `程序量出 ${s.count} 个图形变了` : '程序没量出变化，请看整页目测'}`
        + ` <span class="img-sheet-open">${compare('🆚 对照打开这张表', s.sheet, cell)}</span></div>`
        + (details.length ? `<div class="img-detail">${details.map(d => '· ' + esc(d)).join('\n')}</div>` : '');
      const pages = (s.pages || []).map(p => `<div class="img-page" data-mark="${esc(markKey(folder, f.name, s.sheet, p.page))}">`
        + `<div class="img-page-head"><span>这张表第 ${p.page} / ${p.pages} 页${p.note ? `（${esc(p.note)}）` : ''}</span>`
        + '<button class="btn btn-outline btn-sm" type="button" data-mark-btn>✓ 无需处理</button></div>'
        + `<div class="img-group">${shot(`原版 Excel${p.old_page ? `（第 ${p.old_page} 页 · ${how(p.old_how)}）` : ''}`, p.old_img, p.old_why, p.old_ratio || p.ratio,
          f.old ? compare(esc(f.old), s.sheet, cell) + copyBtn(f.old) : '没有对应的原版文件')}`
        + shot(`新版 Excel${f.fixed ? '（修复前）' : ''}（第 ${p.page} 页 · ${how(p.new_how)}）`, p.new_img, '', p.new_ratio || p.ratio, openNew(esc(f.path || f.name), s.sheet, cell) + copyBtn(f.path))
        + `${shot(`新版 PDF${p.pdf_page ? `（第 ${p.pdf_page} 页）` : ''}`, p.pdf_img, p.pdf_why, p.pdf_ratio || p.ratio,
          f.pdf ? `<a href="#" data-open-pdf="${esc(f.name)}" title="点击打开这份 PDF">${esc(f.pdf)}</a>` + copyBtn(f.pdf) : '没找到同名 PDF')}</div></div>`).join('');
      return head + (pages || `<div class="img-none">${esc(s.pages_why || '这张表没有导出页')}</div>`);
    }).join('');
    return `<div class="img-file${f.fixed ? ' is-fixed' : ''}"><div class="img-file-head"><span class="img-file-name">${esc(f.name)}</span>`
      + `<span class="img-links">${links}</span>${badge}${groupBadge}</div>${sheets}</div>`;
  }

  function renderResults() {
    const data = result();
    // 结果没变就不重画（修复进行中每秒刷新进度，整页重画会闪、图片会重新加载）
    const sig = JSON.stringify([latest.running && latest.mode === 'scan', data]);
    if (sig === lastSig) return;
    lastSig = sig;
    if (!data || !data.scanned) listKey = '';
    if (latest.running && latest.mode === 'scan' && !data) { $('imgResults').innerHTML = '<p>正在找，查完一份就会列在这里；进度看右下角的进度框。</p>'; return; }
    if (!data) { $('imgResults').innerHTML = '<p>尚未开始。</p>'; $('imgResultFolder').textContent = ''; return; }
    $('imgResultFolder').textContent = `新版：${data.folder}` + (data.old_folder ? ` · 原版：${data.old_folder}` : '') + (data.pdf_folder ? ` · PDF：${data.pdf_folder}` : '');
    const files = data.files || [];
    if (!data.scanned) {  // 只组合（没先找）：列出每份组合了几组
      const link = f => f.path ? window.excelLinkHtml(esc(f.name), { file: f.name, folder: data.folder }) : esc(f.name);
      const rows = files.filter(f => f.error || f.grouped).map(f => f.error
        ? `<li class="font-item is-bad"><span class="font-mark">!</span><span class="font-name">${link(f)}</span><span class="font-error">没处理：${esc(f.error)}</span></li>`
        : `<li class="font-item is-ok"><span class="font-mark">📌</span><span class="font-name">${link(f)}</span><span class="font-detail">组合了 ${f.grouped} 组\n${esc(groupTitle(f))}</span></li>`).join('');
      $('imgResults').innerHTML = rows ? `<ul class="font-list">${rows}</ul>` : '<p>没有截图上叠着红框、箭头的地方，什么都没改。</p>';
      return;
    }
    // 按查找顺序往下追加（用户 2026-11-11 要：查看时页面不动，新查完的接到最下面，不重排、不顶掉正在看的位置）。
    // 已经显示的那份只在它自己内容变了（比如点了修复）时原地换，别的不碰。
    const key = `${data.folder}|${data.stamp}`;
    if (key !== listKey || !$('imgList')) {
      listKey = key;
      shown.clear();
      $('imgResults').innerHTML = '<p id="imgSummary" class="font-summary"></p><p id="imgHint" class="toolbar-hint"></p>'
        + `<label class="img-group-opt"><input type="checkbox" id="imgOnlyBad"${onlyBad ? ' checked' : ''}>只看程序量出变化、没查的文件（其余先藏起来）</label>`
        + `<label class="img-group-opt"><input type="checkbox" id="imgHideDone"${hideDone ? ' checked' : ''}>藏起已标「无需处理」的页 <span id="imgMarkCount"></span></label>`
        + `<div id="imgList" class="${onlyBad ? 'only-bad ' : ''}${hideDone ? 'hide-done' : ''}"></div><p id="imgEmpty" hidden>没有图片。</p>`;
    }
    const bad = files.filter(f => f.error), hit = files.filter(f => !f.error && f.count);
    const review = files.filter(f => !f.error && !f.count && f.review);
    const same = files.filter(f => !f.error && !f.count && !f.review);
    $('imgSummary').innerHTML = `<span class="is-bad">${hit.length} 份程序量出变化</span>`
      + `<span class="is-bad">${review.length} 份有图片、请目测</span>`
      + `<span class="is-ok">${same.length} 份没有图片</span>` + (bad.length ? `<span class="is-bad">! ${bad.length} 份没查</span>` : '');
    $('imgHint').textContent = data.partial ? `还在查（已查 ${files.length} 份）：按查找顺序排，新查完的接在最下面，正在看的位置不会动。`
      : `按查找顺序排（共 ${files.length} 份）。勾下面「只看…」可以先藏起没量出问题的。`;
    const link = f => f.path ? window.excelLinkHtml(esc(f.name), { file: f.name, folder: data.folder }) : esc(f.name);
    const list = $('imgList');
    for (const f of files) {
      if (!f.error && !f.count && !f.review) continue;
      const html = f.error
        ? `<div class="img-file is-flag"><ul class="font-list"><li class="font-item is-bad"><span class="font-mark">!</span><span class="font-name">${link(f)}</span><span class="font-error">没查：${esc(f.error)}</span></li></ul></div>`
        : renderFile(f, data).replace('<div class="img-file', `<div class="img-file${f.count ? ' is-flag' : ''}`);
      const old = shown.get(f.name);
      if (old && old.html === html) continue;
      const box = document.createElement('div');
      box.innerHTML = html;
      const el = box.firstElementChild;
      if (old) old.el.replaceWith(el); else list.append(el);
      shown.set(f.name, { html, el });
    }
    applyMarks();
    $('imgEmpty').hidden = !!shown.size || !!data.partial;
  }

  // 查完时上面的进度条会收起，结果整体往上挪；用户已经往下看结果时把滚动补回去，眼前的内容不动
  function render() {
    const anchor = $('imgResults'), before = anchor.getBoundingClientRect().top;
    controls(); renderProgress(); renderPanel(); renderResults(); controls();
    const moved = anchor.getBoundingClientRect().top - before;
    if (before < 0 && Math.abs(moved) >= 1) window.scrollBy(0, moved);
  }

  async function refresh() {
    clearTimeout(timer);
    let failed = false;
    try {
      const response = await fetch(API, { cache: 'no-store' });
      if (!response.ok) throw Error('读取进度失败，正在重试');
      latest = await response.json();
      if ((latest.running || initial) && latest.folder) $('imgFolder').value = latest.folder;
      if ((latest.running || initial) && latest.old_folder) $('imgOldFolder').value = latest.old_folder;
      if (wasRunning && !latest.running) {
        const files = result()?.files || [];
        console.info('🖼️ 图片任务做完', { 动作: latest.mode, 份数: files.length, 量出变化: files.filter(f => f.count).length, 已修复: files.filter(f => f.fixed).length });
        if (latest.mode === 'group') window.gameFx?.celebrate?.('截图和标注组合好了！', `${files.filter(f => f.grouped).length} 份`);
        if (latest.mode === 'fix') window.gameFx?.celebrate?.('图片放回原样了！', `已修复 ${files.filter(f => f.fixed).length} 份`);
      }
      wasRunning = !!latest.running;
      initial = false;
      loading = false;
      render();
    } catch (error) { failed = true; $('imgError').textContent = error.message; }
    if (latest.running || failed) timer = setTimeout(refresh, failed ? 1800 : 900);
  }

  async function post(path, body) {
    const response = await fetch(API + path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const data = await response.json();
    if (!response.ok || !data.ok) throw Error(data.error || '操作失败');
    return data;
  }

  async function start(mode, names) {
    if (loading || submitting || latest.running) return;
    const folder = $('imgFolder').value.trim(), oldFolder = $('imgOldFolder').value.trim(), pdfFolder = $('imgPdfFolder').value.trim();
    $('imgError').textContent = '';
    if (!folder || (!oldFolder && mode !== 'group')) { $('imgError').textContent = '请填写新版文件夹和原版文件夹'; return; }
    const group = $('imgGroupOnFix').checked;
    if (mode === 'group') {
      const okGo = await window.appConfirm(
        `要把这个文件夹（含子文件夹）里每张截图和上面的红框、箭头、文字框永久组合，并保存 Excel 吗？\n\n${folder}\n\n`
        + '· 组合设成「随单元格移动、不改变大小」：以后调行高列宽、导出 PDF 都不会错开\n'
        + '· 不对照原版：已经被拉坏的截图会按坏的样子组合，请先「找出受影响的图片」修好\n\n请先保存并关闭这些 Excel。',
        { title: '组合截图和标注', okText: '开始组合' });
      if (!okGo) return;
    }
    if (mode === 'fix') {
      const okGo = await window.appConfirm(
        `要把 ${names ? names.length : pending().length} 份 Excel 里的截图和红框放回原样，并保存吗？\n\n`
        + '· 截图左上角不动，大小改回原版的\n· 截图上的红框、箭头按原版里相对截图的位置放回\n· 行高、列宽、格子内容都不动\n'
        + (group ? '· 放回后把截图和标注永久组合（随单元格移动、不改变大小）\n' : '') + '\n请先保存并关闭这些 Excel。',
        { title: '修复图片', okText: '开始修复' });
      if (!okGo) return;
    }
    submitting = true; controls();
    console.info('🖼️ 图片任务开始', { 动作: mode, 新版: folder, 原版: oldFolder, PDF: pdfFolder || '(Excel 旁边)', 份数: names?.length || '全部' });
    try {
      await post('/run', { folder, old_folder: oldFolder, pdf_folder: pdfFolder, mode, group, ...(names ? { names } : {}) });
      try {
        localStorage.setItem(FOLDER, folder); localStorage.setItem(PDF, pdfFolder); localStorage.setItem(GROUP, group ? '1' : '0');
        if (oldFolder) localStorage.setItem(OLD, oldFolder);
      } catch (_) {}
      wasRunning = true;
    } catch (error) {
      console.warn('🖼️ 图片任务没能开始', error.message);
      $('imgError').textContent = error.message;
    } finally {
      submitting = false;
      await refresh();
    }
  }

  $('imgScan').addEventListener('click', () => start('scan'));
  $('imgFixAll').addEventListener('click', () => start('fix', pending().map(f => f.name)));
  $('imgGroup').addEventListener('click', () => start('group'));
  $('imgGroupOnFix').addEventListener('change', () => { try { localStorage.setItem(GROUP, $('imgGroupOnFix').checked ? '1' : '0'); } catch (_) {} });
  $('imgResults').addEventListener('change', event => {
    if (event.target.id === 'imgOnlyBad') { onlyBad = event.target.checked; $('imgList')?.classList.toggle('only-bad', onlyBad); }
    if (event.target.id === 'imgHideDone') { hideDone = event.target.checked; $('imgList')?.classList.toggle('hide-done', hideDone); }
  });
  $('imgResults').addEventListener('click', async event => {
    const mark = event.target.closest('[data-mark-btn]');
    if (mark) {
      const key = mark.closest('[data-mark]').dataset.mark;
      if (marks[key]) delete marks[key]; else marks[key] = 1;
      try { localStorage.setItem(MARKS, JSON.stringify(marks)); } catch (_) {}
      console.info('🖼️ 标记无需处理', { 页: key, 标上: !!marks[key] });
      applyMarks();
      return;
    }
    const fix = event.target.closest('[data-fix-file]');
    if (fix) { start('fix', [fix.dataset.fixFile]); return; }
    const copy = event.target.closest('[data-copy]');
    if (copy) {
      const ok = await copyText(copy.dataset.copy);
      console.info('🖼️ 复制路径', { 路径: copy.dataset.copy, 成功: ok });
      copy.textContent = ok ? '✓ 已复制' : '复制不了，请手动选中路径';
      setTimeout(() => { copy.textContent = '📋 复制路径'; }, 1600);
      return;
    }
    const pdf = event.target.closest('[data-open-pdf]');
    if (pdf) {
      event.preventDefault();
      try {
        await post('/open_pdf', { file: pdf.dataset.openPdf });
        console.info('📄 打开新版 PDF', pdf.dataset.openPdf);
      } catch (error) { window.appAlert?.(error.message, { title: '打不开 PDF' }); }
      return;
    }
    const img = event.target.closest('.img-shot img');
    if (img) {
      const zoom = document.createElement('div');
      zoom.className = 'img-zoom';
      zoom.innerHTML = `<img src="${img.src}" alt="">`;
      zoom.addEventListener('click', () => zoom.remove());
      document.body.appendChild(zoom);
    }
  });
  controls();
  refresh();
})();
