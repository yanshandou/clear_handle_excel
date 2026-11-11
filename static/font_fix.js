// 统一字体（用户 2026-11-10 要）：游ゴシック 换成ＭＳ Ｐゴシック，减少切字误报。后端 core/font_fix.py
(() => {
  const $ = id => document.getElementById(id);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const API = '/api/font_fix';
  const FOLDER = 'clear_handle_excel.lastFolder';
  const LABEL = '🔤 统一字体';
  let latest = {}, loading = true, submitting = false, timer = 0, initial = true, wasRunning = false;
  try { $('fontFolder').value = localStorage.getItem(FOLDER) || ''; } catch (_) {}

  function controls() {
    const busy = loading || submitting || latest.running;
    for (const id of ['fontFolder', 'fontRun']) $(id).disabled = !!busy;
    $('fontRun').textContent = submitting || latest.running ? '正在统一字体…' : LABEL;
  }

  const changedText = f => [f.shapes ? `${f.shapes} 个图形` : '', f.cells ? `${f.cells} 个格子` : ''].filter(Boolean).join('、');

  function renderProgress() {
    const total = latest.total || 0, done = latest.done || 0, running = !!latest.running;
    const fraction = total ? done / total : 0;
    $('fontFill').style.width = `${Math.round((running ? fraction : 1) * 100)}%`;
    $('fontTrack').hidden = !running;
    window.ProgressEta?.update($('fontFill'), 'font_fix', fraction, running);
    if (running) {
      $('fontProgress').textContent = `正在处理 ${Math.min(done + 1, total)} / ${total} 份：${latest.current || '正在启动后台 Excel…'}` + (latest.step ? ` · ${latest.step}` : '');
      return;
    }
    const files = latest.result?.files || [];
    const bad = files.filter(f => f.error).length;
    const changed = files.filter(f => !f.error && (f.shapes || f.cells)).length;
    $('fontProgress').textContent = !files.length ? '填写文件夹后点按钮开始。'
      : bad ? `! 做完了，${changed} 份改了字体，${bad} 份没处理（看下面的原因）`
      : `✓ 做完了，${changed} 份改了字体，其余字体都是对的`;
  }

  function renderPanel() {
    if (!window.JobPanel || (!latest.running && !latest.result)) return;
    const log = latest.log || [];
    if (latest.running) {
      window.JobPanel.run('font_fix', {
        title: '正在统一字体', count: window.JobPanel.count(latest.done, latest.total, '份 Excel'),
        now: latest.current ? `📗 ${latest.current}` : '正在启动后台 Excel…', step: latest.step || '',
        fraction: latest.total ? latest.done / latest.total : 0, log, folder: latest.folder || '',
      });
    } else {
      const bad = log.filter(i => !i.ok).length;
      window.JobPanel.finish('font_fix', {
        title: `统一字体做完了${bad ? `，${bad} 份没处理` : ''}`, log, folder: latest.folder || '', warn: bad > 0,
        step: bad ? '黄色的是没处理的，原因写在后面；点文件名在 Excel 打开。' : '',
      });
    }
  }

  function renderResults() {
    const result = latest.result;
    if (latest.running) { $('fontResults').innerHTML = '<p>正在处理，做完一份会出现在右下角的进度框里。</p>'; return; }
    if (!result) { $('fontResults').innerHTML = '<p>尚未开始。</p>'; $('fontResultFolder').textContent = ''; return; }
    $('fontResultFolder').textContent = `文件夹：${result.folder}`;
    const files = result.files || [];
    const bad = files.filter(f => f.error), changed = files.filter(f => !f.error && (f.shapes || f.cells));
    const same = files.filter(f => !f.error && !f.shapes && !f.cells);
    const link = f => f.path ? window.excelLinkHtml(esc(f.name), { file: f.name, folder: result.folder }) : esc(f.name);
    const items = [
      ...bad.map(f => `<li class="font-item is-bad"><span class="font-mark">!</span><span class="font-name">${link(f)}</span><span class="font-error">没处理：${esc(f.error)}</span></li>`),
      ...changed.map(f => `<li class="font-item is-ok"><span class="font-mark">✓</span><span class="font-name">${link(f)}</span>`
        + `<span class="font-detail">换成ＭＳ Ｐゴシック：${esc(changedText(f))}`
        + ` · ${(f.sheets || []).map(s => `${esc(s.sheet)}（${esc(changedText(s))}）`).join('、')}</span></li>`),
    ].join('');
    $('fontResults').innerHTML = `<p class="font-summary"><span class="is-ok">✓ ${changed.length} 份改了字体</span>`
      + `<span>${same.length} 份字体本来就对</span>` + (bad.length ? `<span class="is-bad">! ${bad.length} 份没处理</span>` : '') + '</p>'
      + (items ? `<ul class="font-list">${items}</ul>` : '<p>没有要改的字体。</p>');
  }

  function render() { controls(); renderProgress(); renderPanel(); renderResults(); }

  async function refresh() {
    clearTimeout(timer);
    let failed = false;
    try {
      const response = await fetch(API, { cache: 'no-store' });
      if (!response.ok) throw Error('读取进度失败，正在重试');
      latest = await response.json();
      if ((latest.running || initial) && latest.folder) window.WorkbookPaths?.show($('fontFolder'), latest.folder);
      if (wasRunning && !latest.running) {
        const files = latest.result?.files || [];
        console.info('🔤 统一字体做完', { 份数: files.length, 没处理: files.filter(f => f.error).length });
        if (files.length) window.gameFx?.celebrate?.('字体统一好了！', `${files.filter(f => !f.error && (f.shapes || f.cells)).length} 份改了字体`);
      }
      wasRunning = !!latest.running;
      initial = false;
      loading = false;
      render();
    } catch (error) { failed = true; $('fontError').textContent = error.message; }
    if (latest.running || failed) timer = setTimeout(refresh, failed ? 1800 : 900);
  }

  async function run() {
    if (loading || submitting || latest.running) return;
    const folder = $('fontFolder').value.trim();
    $('fontError').textContent = '';
    if (!folder) { $('fontError').textContent = '请填写 Excel 文件或文件夹路径'; return; }
    const okGo = await window.appConfirm(
      `要把所选 Excel 里的游ゴシック都换成ＭＳ Ｐゴシック，并保存 Excel 吗？\n\n${folder}\n\n`
      + '· 图形里的文字和有内容的格子都查\n· 只换游ゴシック一家，别的字体、字号不动\n· 不改「標準」样式',
      { title: '统一字体', okText: '开始' });
    if (!okGo) return;
    submitting = true; controls();
    console.info('🔤 统一字体开始', { 文件夹: folder });
    try {
      const response = await fetch(API + '/run', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ folder }) });
      const result = await response.json();
      if (!response.ok || !result.ok) throw Error(result.error || '操作失败');
      try { localStorage.setItem(FOLDER, result.folder || folder); } catch (_) {}
      wasRunning = true;
    } catch (error) {
      console.warn('🔤 统一字体没能开始', error.message);
      $('fontError').textContent = error.message;
    } finally {
      submitting = false;
      await refresh();
    }
  }

  $('fontRun').addEventListener('click', run);
  controls();
  refresh();
})();
