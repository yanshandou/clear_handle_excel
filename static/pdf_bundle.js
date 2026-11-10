// 批量导出同名 PDF（用户 2026-11-10 要）：aaa.xlsx 旁边生成 aaa.pdf，所有工作表合成一份。后端 core/pdf_bundle.py
(() => {
  const $ = id => document.getElementById(id);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const API = '/api/pdf_bundle';
  const FOLDER = 'clear_handle_excel.lastFolder';
  const HIDDEN = 'clear_handle_excel.pdfBundleHidden';
  const SKIP = 'clear_handle_excel.pdfBundleSkip';
  const NO_HEADER = 'clear_handle_excel.pdfBundleNoHeader';
  const DEFAULT_SKIP = '変更履歴';
  const LABEL = '📕 批量导出同名 PDF';
  let latest = {}, loading = true, submitting = false, timer = 0, initial = true, wasRunning = false;
  try {
    $('bundleFolder').value = localStorage.getItem(FOLDER) || '';
    $('bundleHidden').checked = localStorage.getItem(HIDDEN) === '1';
    $('bundleNoHeader').checked = localStorage.getItem(NO_HEADER) !== '0';
  } catch (_) {}
  let savedSkip = null;
  try { savedSkip = localStorage.getItem(SKIP); } catch (_) {}
  $('bundleSkip').value = savedSkip ?? DEFAULT_SKIP;

  // 不打印的表名：一行一个，逗号、顿号、分号也能隔开
  const skipNames = () => [...new Set($('bundleSkip').value.split(/[\n,，、;；]+/).map(v => v.trim()).filter(Boolean))];
  function showSkip() {
    const names = skipNames();
    $('bundleSkipCount').textContent = names.length ? `共 ${names.length} 个表名会跳过` : '名单是空的：所有表都打印';
    try { localStorage.setItem(SKIP, $('bundleSkip').value); } catch (_) {}
  }
  $('bundleSkip').addEventListener('input', showSkip);
  $('bundleSkipReset').addEventListener('click', () => { $('bundleSkip').value = DEFAULT_SKIP; showSkip(); });
  showSkip();

  const baseName = p => String(p || '').split(/[\\/]/).pop();

  function controls() {
    const busy = loading || submitting || latest.running;
    for (const id of ['bundleFolder', 'bundleHidden', 'bundleNoHeader', 'bundleSkip', 'bundleSkipReset', 'bundleRun']) $(id).disabled = !!busy;
    $('bundleRun').textContent = submitting || latest.running ? '正在导出…' : LABEL;
  }

  function renderProgress() {
    const total = latest.total || 0, done = latest.done || 0;
    const running = !!latest.running;
    const fraction = total ? done / total : 0;
    $('bundleFill').style.width = `${Math.round((running ? fraction : 1) * 100)}%`;
    // 做完收起进度条，写 ✓ / ! 提示（项目约定 #127）
    $('bundleTrack').hidden = !running;
    window.ProgressEta?.update($('bundleFill'), 'pdf_bundle', fraction, running);
    if (running) {
      $('bundleProgress').textContent = `正在导出 ${Math.min(done + 1, total)} / ${total} 份：${latest.current || '正在启动后台 Excel…'}`
        + (latest.step ? ` · ${latest.step}` : '');
      return;
    }
    const files = latest.result?.files || [];
    const bad = files.filter(f => f.error).length;
    $('bundleProgress').textContent = !files.length ? '填写文件夹后点按钮开始。'
      : bad ? `! 导出完成，${files.length - bad} 份成功，${bad} 份没导出（看下面的原因）`
      : `✓ 导出完成，${files.length} 份 Excel 都已生成同名 PDF`;
  }

  function renderPanel() {
    if (!window.JobPanel || (!latest.running && !latest.result)) return;
    const log = latest.log || [];
    if (latest.running) {
      window.JobPanel.run('pdf_bundle', {
        title: '正在批量导出同名 PDF', count: window.JobPanel.count(latest.done, latest.total, '份 Excel'),
        now: latest.current ? `📗 ${latest.current}` : '正在启动后台 Excel…', step: latest.step || '',
        fraction: latest.total ? latest.done / latest.total : 0, log, folder: latest.folder || '',
      });
    } else {
      const bad = log.filter(i => !i.ok).length;
      window.JobPanel.finish('pdf_bundle', {
        title: `批量导出同名 PDF 做完了${bad ? `，${bad} 份没导出` : ''}`, log, folder: latest.folder || '', warn: bad > 0,
        step: bad ? '黄色的是没导出的，原因写在后面；点文件名在 Excel 打开。' : '',
      });
    }
  }

  function renderResults() {
    const result = latest.result;
    if (latest.running && !result) { $('bundleResults').innerHTML = '<p>正在导出，做完一份会出现在右下角的进度框里。</p>'; return; }
    if (!result) { $('bundleResults').innerHTML = '<p>尚未开始。</p>'; $('bundleResultFolder').textContent = ''; return; }
    const skipList = result.options?.skip_sheets || [];
    $('bundleResultFolder').textContent = `文件夹：${result.folder}${result.options?.include_hidden ? ' · 包含隐藏表' : ' · 跳过隐藏表'}`
      + `${result.options?.no_header_footer === false ? ' · 保留页眉页脚' : ' · 去掉页眉页脚'}`
      + ` · 不打印：${skipList.length ? skipList.join('、') : '无'}`;
    const files = result.files || [];
    const ok = files.filter(f => !f.error), bad = files.filter(f => f.error);
    const pages = ok.reduce((sum, f) => sum + (f.pages || 0), 0);
    const items = [...bad, ...ok].map(file => {
      const target = { file: file.name, folder: result.folder };
      const name = file.path ? window.excelLinkHtml(esc(file.name), target) : esc(file.name);
      if (file.error) {
        return `<li class="bundle-item is-bad"><span class="bundle-mark">!</span><span class="bundle-name">${name}</span>`
          + `<span class="bundle-error">没导出：${esc(file.error)}</span></li>`;
      }
      const skipped = (file.skipped || []).map(s => `${esc(s.sheet)}（${esc(s.reason)}）`).join('、');
      return `<li class="bundle-item is-ok"><span class="bundle-mark">✓</span><span class="bundle-name">${name}</span>`
        + `<span class="bundle-detail">→ <span class="bundle-pdf">${esc(baseName(file.pdf))}</span> · ${file.sheets.length} 张表 · ${file.pages} 页</span>`
        + (skipped ? `<span class="bundle-skip">没放进 PDF：${skipped}</span>` : '') + '</li>';
    }).join('');
    $('bundleResults').innerHTML = `<p class="bundle-summary"><span class="is-ok">✓ 已生成 ${ok.length} 份 PDF（共 ${pages} 页）</span>`
      + (bad.length ? `<span class="is-bad">! ${bad.length} 份没导出</span>` : '') + '</p>'
      + (items ? `<ul class="bundle-list">${items}</ul>` : '<p>文件夹里没有 Excel。</p>');
  }

  function render() { controls(); renderProgress(); renderPanel(); renderResults(); }

  async function refresh() {
    clearTimeout(timer);
    let failed = false;
    try {
      const response = await fetch(API, { cache: 'no-store' });
      if (!response.ok) throw Error('读取导出进度失败，正在重试');
      latest = await response.json();
      if (latest.running || initial) {
        if (latest.folder) $('bundleFolder').value = latest.folder;
        if (latest.options) {
          $('bundleHidden').checked = !!latest.options.include_hidden;
          $('bundleNoHeader').checked = latest.options.no_header_footer !== false;
          if (latest.running && Array.isArray(latest.options.skip_sheets)) $('bundleSkip').value = latest.options.skip_sheets.join('\n');
        }
      }
      if (wasRunning && !latest.running) {
        const files = latest.result?.files || [];
        console.info('📕 批量导出同名 PDF 做完', { 份数: files.length, 没导出: files.filter(f => f.error).length });
        if (files.length) window.gameFx?.celebrate?.('导出完成！', `${files.filter(f => !f.error).length} 份同名 PDF 已放在 Excel 旁边`);
      }
      wasRunning = !!latest.running;
      initial = false;
      loading = false;
      render();
    } catch (error) { failed = true; $('bundleError').textContent = error.message; }
    if (latest.running || failed) timer = setTimeout(refresh, failed ? 1800 : 900);
  }

  async function run() {
    if (loading || submitting || latest.running) return;
    const folder = $('bundleFolder').value.trim();
    const includeHidden = $('bundleHidden').checked;
    const skip = skipNames();
    const noHeader = $('bundleNoHeader').checked;
    $('bundleError').textContent = '';
    if (!folder) { $('bundleError').textContent = '请填写 Excel 文件夹'; return; }
    const okGo = await window.appConfirm(
      `要把这个文件夹（含子文件夹）里的每份 Excel 都导出成同名 PDF，放在 Excel 旁边吗？\n\n${folder}\n\n`
      + `· 已有的同名 PDF 会被覆盖\n· 不打印的表：${skip.length ? skip.join('、') : '无（所有表都打印）'}\n· 页眉页脚：${noHeader ? '去掉' : '保留'}\n· 隐藏的工作表：${includeHidden ? '也放进 PDF' : '跳过'}\n· 本工具加的标记色导出时临时去掉（用户自己的颜色不动）\n· Excel 只读打开，不会被改动`,
      { title: '批量导出同名 PDF', okText: '开始导出' });
    if (!okGo) return;
    submitting = true; controls();
    console.info('📕 批量导出同名 PDF 开始', { 文件夹: folder, 包含隐藏表: includeHidden, 不打印的表: skip, 去掉页眉页脚: noHeader });
    try {
      const response = await fetch(API + '/run', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ folder, options: { include_hidden: includeHidden, skip_sheets: skip, no_header_footer: noHeader } }),
      });
      const result = await response.json();
      if (!response.ok || !result.ok) throw Error(result.error || '操作失败');
      try { localStorage.setItem(FOLDER, folder); localStorage.setItem(HIDDEN, includeHidden ? '1' : ''); localStorage.setItem(NO_HEADER, noHeader ? '1' : '0'); } catch (_) {}
      wasRunning = true;
    } catch (error) {
      console.warn('📕 批量导出同名 PDF 没能开始', error.message);
      $('bundleError').textContent = error.message;
    } finally {
      submitting = false;
      await refresh();
    }
  }

  $('bundleRun').addEventListener('click', run);
  controls();
  refresh();
})();
