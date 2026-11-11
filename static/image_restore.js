// 修复图片（用户 2026-11-11 要）：对照参考文件，把被行高列宽拉变形的截图、跑偏的红框放回原样。后端 core/image_restore.py
(() => {
  const $ = id => document.getElementById(id);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const API = '/api/image_restore';
  const FOLDER = 'clear_handle_excel.lastFolder';
  const OLD = 'clear_handle_excel.oldFolder';
  const LABEL = '🖼️ 修复图片';
  let latest = {}, loading = true, submitting = false, timer = 0, initial = true, wasRunning = false;
  try { $('imgFolder').value = localStorage.getItem(FOLDER) || ''; $('imgOldFolder').value = localStorage.getItem(OLD) || ''; } catch (_) {}

  function controls() {
    const busy = loading || submitting || latest.running;
    for (const id of ['imgFolder', 'imgOldFolder', 'imgRun']) $(id).disabled = !!busy;
    $('imgRun').textContent = submitting || latest.running ? '正在修复图片…' : LABEL;
  }

  function renderProgress() {
    const total = latest.total || 0, done = latest.done || 0, running = !!latest.running;
    const fraction = total ? done / total : 0;
    $('imgFill').style.width = `${Math.round((running ? fraction : 1) * 100)}%`;
    $('imgTrack').hidden = !running;
    window.ProgressEta?.update($('imgFill'), 'image_restore', fraction, running);
    if (running) {
      $('imgProgress').textContent = `正在处理 ${Math.min(done + 1, total)} / ${total} 份：${latest.current || '正在启动后台 Excel…'}` + (latest.step ? ` · ${latest.step}` : '');
      return;
    }
    const files = latest.result?.files || [];
    const bad = files.filter(f => f.error).length;
    const changed = files.filter(f => !f.error && f.count).length;
    $('imgProgress').textContent = !files.length ? '填写两个文件夹后点按钮开始。'
      : bad ? `! 做完了，${changed} 份放回了图片，${bad} 份没处理（看下面的原因）`
      : `✓ 做完了，${changed} 份放回了图片，其余图片都没变`;
  }

  function renderPanel() {
    if (!window.JobPanel || (!latest.running && !latest.result)) return;
    const log = latest.log || [];
    if (latest.running) {
      window.JobPanel.run('image_restore', {
        title: '正在修复图片', count: window.JobPanel.count(latest.done, latest.total, '份 Excel'),
        now: latest.current ? `📗 ${latest.current}` : '正在启动后台 Excel…', step: latest.step || '',
        fraction: latest.total ? latest.done / latest.total : 0, log, folder: latest.folder || '',
      });
    } else {
      const bad = log.filter(i => !i.ok).length;
      window.JobPanel.finish('image_restore', {
        title: `修复图片做完了${bad ? `，${bad} 份没处理` : ''}`, log, folder: latest.folder || '', warn: bad > 0,
        step: bad ? '黄色的是没处理的，原因写在后面；点文件名在 Excel 打开。' : '',
      });
    }
  }

  function renderResults() {
    const result = latest.result;
    if (latest.running) { $('imgResults').innerHTML = '<p>正在处理，做完一份会出现在右下角的进度框里。</p>'; return; }
    if (!result) { $('imgResults').innerHTML = '<p>尚未开始。</p>'; $('imgResultFolder').textContent = ''; return; }
    $('imgResultFolder').textContent = `处理：${result.folder} · 参考：${result.old_folder}`;
    const files = result.files || [];
    const bad = files.filter(f => f.error), changed = files.filter(f => !f.error && f.count);
    const same = files.filter(f => !f.error && !f.count);
    const link = f => f.path ? window.excelLinkHtml(esc(f.name), { file: f.name, folder: result.folder }) : esc(f.name);
    const items = [
      ...bad.map(f => `<li class="font-item is-bad"><span class="font-mark">!</span><span class="font-name">${link(f)}</span><span class="font-error">没处理：${esc(f.error)}</span></li>`),
      ...changed.map(f => `<li class="font-item is-ok"><span class="font-mark">✓</span><span class="font-name">${link(f)}</span>`
        + `<span class="font-detail">${(f.sheets || []).map(s => `【${esc(s.sheet)}】放回 ${s.count} 个\n${(s.details || []).map(d => '· ' + esc(d)).join('\n')}`).join('\n')}</span></li>`),
    ].join('');
    $('imgResults').innerHTML = `<p class="font-summary"><span class="is-ok">✓ ${changed.length} 份放回了图片</span>`
      + `<span>${same.length} 份图片没变</span>` + (bad.length ? `<span class="is-bad">! ${bad.length} 份没处理</span>` : '') + '</p>'
      + (items ? `<ul class="font-list">${items}</ul>` : '<p>没有变形或跑偏的图片。</p>');
  }

  function render() { controls(); renderProgress(); renderPanel(); renderResults(); }

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
        const files = latest.result?.files || [];
        console.info('🖼️ 修复图片做完', { 份数: files.length, 放回: files.filter(f => !f.error && f.count).length, 没处理: files.filter(f => f.error).length });
        if (files.length) window.gameFx?.celebrate?.('图片放回原样了！', `${files.filter(f => !f.error && f.count).length} 份改了图片`);
      }
      wasRunning = !!latest.running;
      initial = false;
      loading = false;
      render();
    } catch (error) { failed = true; $('imgError').textContent = error.message; }
    if (latest.running || failed) timer = setTimeout(refresh, failed ? 1800 : 900);
  }

  async function run() {
    if (loading || submitting || latest.running) return;
    const folder = $('imgFolder').value.trim(), oldFolder = $('imgOldFolder').value.trim();
    $('imgError').textContent = '';
    if (!folder || !oldFolder) { $('imgError').textContent = '请填写处理后的文件夹和参考文件夹'; return; }
    const okGo = await window.appConfirm(
      `要对照参考文件，把截图和上面的红框放回原样，并保存 Excel 吗？\n\n处理：${folder}\n参考：${oldFolder}\n\n`
      + '· 截图左上角不动，大小改回原文件的\n· 截图上的红框、箭头按原来相对截图的位置放回\n· 行高、列宽、格子内容都不动',
      { title: '修复图片', okText: '开始' });
    if (!okGo) return;
    submitting = true; controls();
    console.info('🖼️ 修复图片开始', { 处理: folder, 参考: oldFolder });
    try {
      const response = await fetch(API + '/run', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ folder, old_folder: oldFolder }) });
      const result = await response.json();
      if (!response.ok || !result.ok) throw Error(result.error || '操作失败');
      try { localStorage.setItem(FOLDER, folder); localStorage.setItem(OLD, oldFolder); } catch (_) {}
      wasRunning = true;
    } catch (error) {
      console.warn('🖼️ 修复图片没能开始', error.message);
      $('imgError').textContent = error.message;
    } finally {
      submitting = false;
      await refresh();
    }
  }

  $('imgRun').addEventListener('click', run);
  controls();
  refresh();
})();
