// 三个页面共用：/distribute/start 新任务 + 从共享拿文件、/distribute 当前任务批量上传、/distribute?mode=json 按上传记录 JSON 复制 Excel（后端 core/distribute.py）
(() => {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[ch]));
  const api = '/api/distribute';
  const KEYS = { source: 'clear_handle_excel.distributeSource', dest: 'clear_handle_excel.distributeDest', plan: 'clear_handle_excel.distributePlan',
    freeSource: 'clear_handle_excel.freeUploadSource', freeDest: 'clear_handle_excel.freeUploadDest' };
  let latest = {}, submitting = false, loading = true, timer, pollError = '', creating = false;
  let viewing = null;          // 正在看的历史记录（null = 看当前清单 / 最近一次结果）
  let filter = 'all';
  let wasRunning = false;
  // 默认进入当前任务上传，JSON 复制从导航或保存的分配方案进入。
  const params = new URLSearchParams(location.search);
  let task = null;
  let taskFiles = [], readingFiles = false;
  const selected = new Set();
  const jsonMode = params.get('mode') === 'json' || params.has('plan');
  // 起点 → 终点 上传（用户 2026-11-07：几个账号分着做，汇总的人从团队 OneDrive 传到公司 OneDrive）
  const freeMode = !jsonMode && params.get('mode') === 'free';
  let needsFreePreview = true;
  // 做完只有「返回首页」（static/tools.js）：首页上这一页对应的卡片会闪几下
  const homeUrl = () => window.Tools?.homeUrl() || '/';
  const backHomeHtml = () => `<a class="btn btn-primary next-step-btn" href="${esc(homeUrl())}">🏠 返回首页</a>`;
  // 新任务 / 拿文件单独一页，上传页只管上传（以前都挤在 /distribute）
  const startMode = !jsonMode && /^\/distribute\/start\/?$/.test(location.pathname);
  if (!jsonMode && !startMode && params.get('fetch') === '1') { location.replace('/distribute/start?fetch=1'); return; }
  // 每页只看自己那一类：拿文件页只看拿取，上传页不看拿取
  const inScope = data => !data || jsonMode || (freeMode ? isFree(data) : startMode ? isFetch(data) : !isFetch(data) && !isFree(data));
  document.querySelector('[data-dist-mode="' + (jsonMode ? 'json' : freeMode ? 'free' : startMode ? 'start' : 'task') + '"]')?.setAttribute('aria-current', 'page');
  $('distToday').textContent = new Intl.DateTimeFormat('zh-CN', { dateStyle: 'long' }).format(new Date());
  window.lucide?.createIcons();
  let needsPreview = params.has('plan');
  if (jsonMode) {
    for (const id of ['distJsonIntro', 'distJsonFields', 'distJsonHelp', 'distActions']) $(id).hidden = false;
    $('distTaskNotice').hidden = true;
    document.title = '按 JSON 复制 Excel - Excel 整理工具箱';
    $('distNavOverview').textContent = '复制 Excel';
    $('distBreadcrumb').textContent = 'JSON 复制';
    $('distHeading').textContent = '先生成清单，确认后再复制';
    $('distSubtitle').textContent = '导入上传方案，按清单复制 Excel。';
    $('distControlsTitle').textContent = '设置复制清单';
    $('distProgress').textContent = '填好两个根目录和 JSON 后点「生成复制清单」。';
  } else if (freeMode) {
    $('distFreeFields').hidden = false;
    $('distTaskNotice').hidden = true;
    $('distJsonIntro').hidden = false;
    $('distJsonIntro').querySelector('h2').textContent = '起点 → 终点 上传';
    $('distJsonIntro').querySelector('.section-kicker').innerHTML = 'UPLOAD <span>01 / 从哪儿传到哪儿</span>';
    for (const el of $('distJsonIntro').querySelectorAll('p')) el.hidden = true;
    document.title = '起点 → 终点 上传 - Excel 整理工具箱';
    $('distNavOverview').textContent = '起点 → 终点 上传';
    $('distBreadcrumb').textContent = '起点 → 终点';
    $('distHeading').textContent = '从起点传到终点';
    $('distSubtitle').textContent = '自己填从哪儿传到哪儿，比如从团队 OneDrive 传到公司 OneDrive。先生成清单，确认后再传。';
    $('distControlsTitle').textContent = '填起点和终点';
    $('distResultTitle').textContent = '上传清单';
    $('distGoUpload').hidden = false;
  } else if (startMode) {
    document.body.classList.add('dist-start');
    document.title = '新任务 · 拿文件 - Excel 整理工具箱';
    $('distNavOverview').textContent = '新任务 · 拿文件';
    $('distBreadcrumb').textContent = '新任务 · 拿文件';
    $('distHeading').textContent = '新任务 · 从共享拿文件';
    $('distSubtitle').textContent = '填共享文件夹路径，建好任务，从共享拿文件。拿完点「返回首页」，在首页挑要用的卡片。';
    $('distControlsKicker').innerHTML = 'FETCH FILES <span>03 / 拿文件</span>';
    $('distControlsTitle').textContent = '从共享拿文件';
    $('distResultTitle').textContent = '拿取清单';
    $('distHistoryKicker').innerHTML = 'FETCH HISTORY <span>04 / 记录</span>';
    $('distHistoryTitle').textContent = '历史拿取记录';
    $('distNewTask').hidden = false;
    $('distGoManage').hidden = true;
    $('distBrandName').textContent = '新任务 · 入口';
    $('distBrandSub').textContent = 'Excel 整理工坊 · 起点';
    $('distNavResults').textContent = '拿取清单';
    $('distCrumbRoot').textContent = '起点';
    $('distFooter').textContent = 'EXCEL 整理工具箱 / 起点 · 新任务';
  } else {
    $('distGoStart').hidden = false;
  }
  if (!freeMode && !startMode) $('distGoFree').hidden = false;

  const load = key => { try { return localStorage.getItem(key) || ''; } catch (_) { return ''; } };
  const save = (key, value) => { try { localStorage.setItem(key, value); } catch (_) {} };
  $('distSource').value = load(KEYS.source);
  $('distDest').value = load(KEYS.dest);
  $('distPlan').value = load(KEYS.plan);
  // 起点默认填当前在处理的路径（各处理页共用的 lastFolder；用户 2026-11-12），没有才用上次填的
  const lastFolder = load('clear_handle_excel.lastFolder');
  $('distFreeSource').value = lastFolder || load(KEYS.freeSource);
  $('distFreeDest').value = load(KEYS.freeDest);
  // 处理页填的不是单个文件夹时，lastFolder 是程序数据里的「批量路径」目录：换回当时填的那几行（static/task_folder.js）
  if (lastFolder) window.WorkbookPaths?.show($('distFreeSource'), lastFolder);

  // ───── 导出同名 PDF 并一起上传（用户 2026-11-10）：导出设置沿用「批量导出同名 PDF」页（同一份 localStorage） ─────
  const PDF_ON = 'clear_handle_excel.uploadWithPdf';
  const PDF_DEFAULT_SKIP = '変更履歴';
  if (freeMode) {  // 起点 → 终点：放在「生成清单 / 按清单上传」按钮前后
    const actions = $('distFreeFields').querySelector('.link-actions');
    actions.before($('distPdfBox'));
    actions.after($('distPdfAsk'));
  }
  function pdfOptions() {
    let skip = null;  // 没存过 = 默认名单；存过空的 = 所有表都打印
    try { skip = localStorage.getItem('clear_handle_excel.pdfBundleSkip'); } catch (_) {}
    skip = skip ?? PDF_DEFAULT_SKIP;
    return {
      include_hidden: load('clear_handle_excel.pdfBundleHidden') === '1',
      skip_sheets: [...new Set(skip.split(/[\n,，、;；]+/).map(v => v.trim()).filter(Boolean))],
      no_header_footer: load('clear_handle_excel.pdfBundleNoHeader') !== '0',
    };
  }
  // 上传什么（用户 2026-11-10）：excel 只传表格 / both 表格和 PDF 一起 / pdf 只传 PDF
  const pdfMode = () => document.querySelector('input[name="distPdfMode"]:checked')?.value || 'excel';
  const pdfOn = () => pdfMode() !== 'excel';
  const isoTime = text => String(text || '').replace(' ', 'T');
  const pdfStale = file => !!(file.pdf_uploaded_at && file.uploaded_at && isoTime(file.pdf_uploaded_at) < isoTime(file.uploaded_at));
  function showPdfNote() {
    const o = pdfOptions();
    $('distPdfBox').classList.toggle('is-on', pdfOn());
    if (!pdfOn()) { $('distPdfNote').textContent = '只传 Excel，和以前一样。'; return; }
    $('distPdfNote').innerHTML = (pdfMode() === 'pdf'
      ? `<strong>Excel 不传</strong>（终点那份不动），处理文件夹里的 Excel 也不改：照样先核对共享版本，在本机临时副本上做上传前处理，再导出同名 PDF（所有表合一份、宽一页），放在终点 Excel 旁边。`
      : `每份 Excel 上传前处理做完后导出同名 PDF（所有表合一份、宽一页），Excel 传好后放在终点 Excel 旁边。`)
      + `按「批量导出同名 PDF」的设置：不打印的表 ${esc(o.skip_sheets.length ? o.skip_sheets.join('、') : '无')}，页眉页脚${o.no_header_footer ? '去掉' : '保留'}，隐藏表${o.include_hidden ? '也放进去' : '跳过'}`
      + `（<a href="/pdf-bundle" target="_blank" rel="noopener">改设置 ↗</a>）。<br>终点已有同名 PDF 时<strong>默认不覆盖</strong>：上传前先列出来，由你逐个选覆盖还是跳过。`
      + `<br>Excel 没改过、之前已经导过 PDF（上次上传或「批量导出同名 PDF」导的，设置也一样）的，<strong>直接用那份 PDF</strong>，不再重新导`
      + (pdfMode() === 'pdf' ? '，也不再做上传前处理。' : '（Excel 照样做上传前处理再传）。');
  }
  {
    const saved = load(PDF_ON);
    const mode = saved === '1' ? 'both' : ['excel', 'both', 'pdf'].includes(saved) ? saved : 'excel';
    document.querySelector(`input[name="distPdfMode"][value="${mode}"]`).checked = true;
  }
  const GUARD_KEY = 'clear_handle_excel.uploadGuard';
  const guardMode = () => document.querySelector('input[name="distGuardMode"]:checked')?.value || 'auto';
  function showGuardNote() {
    $('distGuardNote').textContent = guardMode() === 'content'
      ? '远程已有同名文件时，逐份展示 Excel 差异。看完并确认后上传这份。'
      : '基准哈希与远程一致的文件直接进入上传队列。基准缺失或对不上时，逐份展示 Excel 差异，确认后再上传。';
  }
  {
    const saved = load(GUARD_KEY);
    document.querySelector(`input[name="distGuardMode"][value="${saved === 'content' ? 'content' : 'auto'}"]`).checked = true;
    showGuardNote();
  }
  $('distPdfBox').addEventListener('change', event => {
    if (event.target.name === 'distGuardMode') {
      save(GUARD_KEY, guardMode());
      showGuardNote();
      if (freeMode && isFree(latest.preview)) needsFreePreview = true;  // 清单是按原来的方式核对的，要重新生成
      if (!loading) controls();
      return;
    }
    save(PDF_ON, pdfMode());
    console.info('📕 上传什么', { excel: '只传表格', both: '表格和 PDF 一起', pdf: '只传 PDF' }[pdfMode()], pdfOptions());
    showPdfNote();
    if (!loading) controls();  // 「全选待上传」按钮的说法跟着变
  });
  window.addEventListener('focus', showPdfNote);
  // 在对比页单独传了一份，回到这个标签页：清单里那一行要跟着没掉（只读一次状态，页面不滚动）
  document.addEventListener('visibilitychange', () => { if (!document.hidden && !loading && !submitting && !latest.running) refresh(); });  // 在另一个标签页改了导出设置，回来就更新
  showPdfNote();

  // 终点已有同名 PDF：在按钮下面列出来，每个选「跳过（默认）/ 覆盖」，点继续才往下走。返回 {rel: 检查时的哈希} 或 null（取消）
  function askPdfOverwrite(check) {
    const existing = check.files.filter(f => f.exists);
    const clash = check.files.filter(f => f.clash);
    const box = $('distPdfAsk');
    if (!existing.length && !clash.length) { box.hidden = true; return Promise.resolve({}); }
    const rows = existing.map((f, i) => `<tr><td>${esc(f.pdf)}<br><small>${esc(f.rel)} 的 PDF</small></td>`
      + `<td>${f.error ? `<span class="dist-pdf-other">读不出来：${esc(f.error)}</span>` : f.ours ? '<span class="dist-pdf-ours">上次是本工具传的，之后没人动过</span>' : '<span class="dist-pdf-other">不是本工具传的，或传上去之后被人改过</span>'}`
      + `<br><small>${esc(f.mtime || '')}${f.size != null ? ' · ' + size(f.size) : ''}</small></td>`
      + `<td><span class="dist-pdf-pick" role="radiogroup" aria-label="${esc(f.pdf)} 怎么处理">`
      + `<label data-pick="skip"><input type="radio" name="distPdfPick${i}" value="skip" checked>跳过</label>`
      + `<label data-pick="overwrite"><input type="radio" name="distPdfPick${i}" value="overwrite" ${f.error ? 'disabled' : ''}>覆盖</label></span></td></tr>`);
    box.innerHTML = `<h4 id="distPdfAskTitle">🛑 终点已经有 ${existing.length} 个同名 PDF，先确认再上传</h4>`
      + `<p>默认都<strong>跳过</strong>（终点那份不动，Excel 照常上传）。确认可以换掉的选「覆盖」；上传那一刻还会再比一次，检查之后又被人改过的照样不覆盖。</p>`
      + (existing.length ? `<div class="dist-pdf-bulk"><button type="button" class="btn btn-outline btn-sm" data-pdf-all="skip">全部跳过</button>`
        + `<button type="button" class="btn btn-outline btn-sm" data-pdf-all="ours">只覆盖本工具上次传的</button>`
        + `<button type="button" class="btn btn-outline btn-sm" data-pdf-all="overwrite">全部覆盖</button></div>`
        + `<div class="link-table-wrap"><table class="link-table"><thead><tr><th>终点已有的 PDF</th><th>是谁的</th><th>怎么处理</th></tr></thead><tbody>${rows.join('')}</tbody></table></div>` : '')
      + (clash.length ? `<p>⚠️ ${clash.map(f => `${esc(f.rel)} 和 ${esc(f.clash)}`).join('、')} 会得到同名 PDF，只传前一份的 PDF。</p>` : '')
      + `<div class="dist-pdf-go"><button type="button" class="button primary" data-pdf-go>按上面的选择继续</button>`
      + `<button type="button" class="button secondary" data-pdf-cancel>先不上传</button><span class="toolbar-hint" data-pdf-count></span></div>`;
    box.hidden = false;
    const picks = () => existing.map((f, i) => box.querySelector(`input[name="distPdfPick${i}"]:checked`)?.value || 'skip');
    const count = () => {
      const p = picks();
      box.querySelector('[data-pdf-count]').textContent = `覆盖 ${p.filter(v => v === 'overwrite').length} 个、跳过 ${p.filter(v => v === 'skip').length} 个`;
    };
    count();
    console.info('📕 终点已有同名 PDF，等你选择', existing.map(f => ({ PDF: f.pdf, 本工具传的: !!f.ours })));
    return new Promise(resolve => {
      const onClick = event => {
        const all = event.target.closest('[data-pdf-all]')?.dataset.pdfAll;
        if (all) {
          existing.forEach((f, i) => {
            const want = f.error ? 'skip' : all === 'ours' ? (f.ours ? 'overwrite' : 'skip') : all;
            const input = box.querySelector(`input[name="distPdfPick${i}"][value="${want}"]`);
            if (input) input.checked = true;
          });
          count();
          return;
        }
        const go = event.target.closest('[data-pdf-go]'), cancel = event.target.closest('[data-pdf-cancel]');
        if (!go && !cancel) return;
        box.removeEventListener('click', onClick);
        box.removeEventListener('change', count);
        box.hidden = true;
        if (cancel) { console.info('📕 同名 PDF：先不上传'); resolve(null); return; }
        const p = picks();
        const overwrite = Object.fromEntries(existing.filter((f, i) => p[i] === 'overwrite').map(f => [f.rel, f.sha256]));
        console.info('📕 同名 PDF 选择', { 覆盖: Object.keys(overwrite), 跳过: existing.filter((f, i) => p[i] !== 'overwrite').map(f => f.rel) });
        resolve(overwrite);
      };
      box.addEventListener('click', onClick);
      box.addEventListener('change', count);
    });
  }

  // 上传前：查终点同名 PDF、让用户选，返回要发给后台的 pdf 选项（null = 取消）；没勾选项返回 undefined
  async function pdfPlan(payload) {
    if (!pdfOn()) return undefined;
    $('distError').textContent = '';
    const box = $('distPdfAsk');
    box.innerHTML = '<p>正在看终点 Excel 旁边有没有同名 PDF…</p>';
    box.hidden = false;
    let check;
    try { check = await post(api + '/pdf_check', payload); }
    catch (error) { box.hidden = true; throw error; }
    console.info('📕 查终点同名 PDF', { 终点: check.dest, 文件: check.files.length, 已有: check.files.filter(f => f.exists).length });
    const overwrite = await askPdfOverwrite(check);
    if (overwrite === null) return null;
    const existing = check.files.filter(f => f.exists).length;
    const clash = check.files.filter(f => f.clash).length;
    const over = Object.keys(overwrite).length;
    const only = pdfMode() === 'pdf';
    return { options: pdfOptions(), overwrite, only,
      line: `📕 ${only ? '只传 PDF（Excel 不传）' : '同名 PDF 一起上传'}：新建 ${check.files.length - existing - clash} 个`
        + (over ? `，覆盖 ${over} 个` : '') + (existing - over ? `，终点已有、不覆盖 ${existing - over} 个` : '') + '。哪份导出 PDF 没做成，那份就不传，其余照传。' };
  }

  const size = n => n == null ? '-' : n < 1024 ? `${n} B` : n < 1048576 ? `${(n / 1024).toFixed(1)} KB` : `${(n / 1048576).toFixed(1)} MB`;
  const badChecks = folder => (folder.checks || []).filter(c => !c.ok);
  const conflicts = folder => (folder.files || []).filter(f => f.conflict || f.status === 'conflict');
  const isFetch = data => data && data.kind === 'fetch';
  const isUpload = data => data && data.kind === 'task_upload';
  function isFree(data) { return !!data && data.kind === 'free_upload'; }
  const isBad = folder => !!(folder.error || conflicts(folder).length || (folder.dir_errors || []).length || badChecks(folder).length || (folder.files || []).some(f => f.status && f.status !== 'ok')
    || (folder.dest_extra || []).length);

  function controls() {
    const busy = loading || submitting || latest.running;
    for (const id of ['distSource', 'distDest', 'distPlan', 'distPick']) $(id).disabled = !!busy;
    for (const input of document.querySelectorAll('input[name="distPdfMode"], input[name="distGuardMode"]')) input.disabled = !!busy;
    $('distPreview').disabled = !!busy;
    $('distFetch').disabled = !!busy;
    const fetchMode = isFetch(latest.preview) || (latest.running && latest.kind === 'fetch');
    const canStart = !busy && !needsPreview && latest.preview && (latest.preview.folders || []).some(f => f.files.length || (f.dirs_to_create || []).length);
    $('distStart').disabled = !canStart;
    $('distPreview').hidden = !jsonMode;
    $('distStart').hidden = !jsonMode && !fetchMode;
    $('distNewCreate').disabled = !!busy || creating;
    $('distNewCreate').textContent = creating ? '正在建文件夹…' : '🆕 建好任务，生成拿取清单';
    $('distPreview').textContent = submitting === 'preview' ? '正在生成清单…' : '📋 生成复制清单';
    $('distFetch').textContent = submitting === 'fetch' ? '正在对比共享和本地…' : task ? '⬇ 从共享拿最新版' : '⬇ 从目的地拿最新版';
    $('distStart').textContent = fetchMode ? (latest.running ? '正在拿…' : '⬇ 按清单拿文件')
      : (latest.running ? '正在复制…' : '📥 按清单开始复制');
    if (task && !startMode) taskControls(busy || readingFiles);
    if (freeMode) {
      $('distFreeSource').disabled = $('distFreeDest').disabled = !!busy;
      $('distFreePreview').disabled = !!busy;
      $('distFreePreview').textContent = submitting === 'free' ? '正在生成清单（要读每个文件算哈希）…' : '📋 生成上传清单';
      const p = isFree(latest.preview) ? latest.preview : null;
      $('distFreeStart').disabled = !!busy || needsFreePreview || !p || !p.summary['要上传'];
      $('distFreeStart').textContent = latest.running ? '正在上传…' : '🚚 按清单上传';
    }
  }

  function taskControls(busy) {
    const files = taskFiles.filter(file => selected.has(file.rel));
    const reviewCount = files.filter(file => file.review?.ready).length;
    for (const id of ['distTaskPickWaiting', 'distTaskPickAll', 'distTaskClear', 'distTaskReload']) $(id).disabled = !!busy;
    for (const box of $('distTaskFiles').querySelectorAll('input[type="checkbox"]')) {
      box.disabled = !!busy;
      box.checked = selected.has(box.value);
    }
    $('distTaskSelection').textContent = `已选 ${files.length} / ${taskFiles.length} 个`;
    const waitingCount = taskFiles.filter(waitingFilter()).length;
    $('distTaskPickWaiting').textContent = `☑ 全选${pdfMode() === 'pdf' ? '还没传 PDF 的' : '待上传'}（${waitingCount} 个）`;
    $('distTaskPickWaiting').disabled = !!busy || !waitingCount;
    $('distTaskPickAll').textContent = `全选全部（${taskFiles.length} 个）`;
    $('distTaskUpload').disabled = !!busy || !files.length;
    $('distTaskUpload').textContent = reviewCount ? '确认重新处理完毕，再检查并上传' : '检查并批量上传';
  }

  async function reloadTaskFiles() {
    readingFiles = true;
    controls();
    try {
      const data = await post(api + '/task_files', { task_id: task.id });
      taskFiles = data.files;
      $('distTaskTotal').textContent = taskFiles.length;
      $('distTaskUploaded').textContent = taskFiles.filter(file => file.uploaded_current && !file.review).length;
      $('distTaskPending').textContent = taskFiles.filter(file => file.review).length;
      const found = new Set(taskFiles.map(file => file.rel));
      for (const rel of selected) if (!found.has(rel)) selected.delete(rel);
      const rows = taskFiles.map(file => {
        const upload = (file.uploaded ? `<span class="link-ok">已上传${file.uploaded_at ? ' · ' + esc(file.uploaded_at.replace('T', ' ')) : ''}</span>` : '')
          + (file.pdf_uploaded_at ? `${file.uploaded ? '<br>' : ''}<small class="dist-pdf-line">📕 PDF ${pdfStale(file) ? '是旧的（之后又传过 Excel）' : '已传'} · ${esc(file.pdf_uploaded_at.replace('T', ' '))}</small>` : '');
        const status = file.review
          ? (file.review.ready ? '待重新处理' : '同步未完成，请重新检查')
          : file.uploaded_current === true ? '当前版本已上传'
          : file.uploaded_current === false ? '本地有新版本待上传'
          : file.uploaded ? '当前版本待核对' : '待上传';
        return `<tr><td><input type="checkbox" aria-label="选择 ${esc(file.rel)}" value="${esc(file.rel)}"></td><td>${fileLink(file.rel, task.work, file.rel)}${file.uploaded_current === true && !file.review ? '' : diffLink(task.work, task.shared, file.rel, file.rel, task.id)}</td><td>${status}</td><td>${upload || '—'}</td></tr>`;
      });
      $('distTaskFiles').innerHTML = (rows.length
        ? '<table class="link-table"><thead><tr><th>选择</th><th>处理文件</th><th>状态</th><th>上次上传</th></tr></thead><tbody>' + rows.join('') + '</tbody></table>'
        : '<p>处理文件夹里还没有 Excel 文件。请先到 <a href="/distribute/start">新任务 · 拿文件</a> 页从共享拿最新版。</p>')
        + data.errors.map(err => `<p class="link-error">${esc(err)}</p>`).join('');
      $('distTaskReview').hidden = !taskFiles.some(file => file.review?.ready);
      try {  // 上传成功的会从本轮清单划掉，重新读一下
        const tasks = await (await fetch('/api/tasks', { cache: 'no-store' })).json();
        const mine = (tasks.tasks || []).find(t => t.id === task.id);
        if (mine) task.focus = mine.focus || null;
      } catch (_) {}
      focusControls();
      if (window.refreshFocusBanner) window.refreshFocusBanner();
    } catch (error) {
      taskFiles = [];
      $('distTaskFiles').textContent = error.message;
      $('distTaskReview').hidden = true;
      for (const id of ['distTaskTotal', 'distTaskUploaded', 'distTaskPending']) $(id).textContent = '—';
    } finally {
      readingFiles = false;
      controls();
    }
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

  function fileLink(label, folder, rel, sheet = '', cell = '') {
    return window.excelLinkHtml(esc(label), { folder, file: rel, sheet, cell });
  }

  // 仿 Git 对比（用户 2026-11-11）：新标签页左右两栏看「你这份」和「远程那份」哪里不一样
  function joinPath(root, rel) {
    const sep = root.includes('\\') || /^[A-Za-z]:/.test(root) ? '\\' : '/';
    return root.replace(/[\\/]+$/, '') + sep + rel.split('/').join(sep);
  }
  // taskId 给了：对比页上多一个「上传这一份」按钮（用户 2026-11-12），只有当前任务上传的文件有
  // freeId = 起点 → 终点还没执行的清单：对比页同样有「上传这一份」，传完这份从清单里拿掉
  function diffLink(mineRoot, theirsRoot, rel, theirsRel = rel, taskId = '', freeId = '') {
    return ` <a class="dist-diff" target="_blank" rel="noopener" title="左右两栏看你这份和远程那份哪里不一样（文字、公式、字体、颜色、行高列宽、打印设置…）"`
      + ` href="/excel-diff?theirs=${encodeURIComponent(joinPath(theirsRoot, theirsRel))}&mine=${encodeURIComponent(joinPath(mineRoot, rel))}${taskId ? `&task=${encodeURIComponent(taskId)}&rel=${encodeURIComponent(rel)}` : ''}${freeId ? `&free=${encodeURIComponent(freeId)}&rel=${encodeURIComponent(rel)}` : ''}">🔍 对比</a>`;
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
          : f.status === 'conflict' ? `<span class="dist-conflict">🛑 ${f.guard === 'content_diff' ? '内容和目的地不一样' : '别人改过'}，没有覆盖</span>${f.guard === 'content_diff' ? `<br><small>${esc(f.error || f.conflict || '')}</small>` : ''}`
          : f.status === 'failed' ? `<span class="link-error">❌ ${esc(f.error)}</span>` : '<span class="link-skip">没有复制</span>';
        else if (f.conflict && f.guard === 'content_diff') state = `<span class="dist-conflict">🛑 内容和目的地不一样，不覆盖</span><br><small>${esc(f.conflict)}</small>`;
        else if (f.conflict) state = '<span class="dist-conflict">🛑 别人改过，不覆盖</span>';
        else if (f.source_open) state = '<span class="link-skip">⚠️ 来源正在 Excel 中打开，复制时会请你先关掉</span>';
        else if (f.guard === 'same') state = '<span class="dist-guard-ok">✓ 拿走后没人改过</span>';
        else if (f.guard === 'identical') state = '<span class="dist-guard-ok">✓ 目的地已经和你的一样</span>';
        else if (f.guard === 'content_same') state = `<span class="dist-guard-ok">✓ ${esc(f.content_check || '按文字和打印区域核对一致')}</span>`;
        else if (f.guard === 'unknown') state = '<span class="dist-guard-unknown">没有拿取记录，没法检查别人改过没有</span>';
        const base = applied && f.status === 'ok' ? folder.dest_path : folder.source_path;
        const diff = f.action === 'overwrite' && !(applied && f.status === 'ok') ? diffLink(folder.source_path, folder.dest_path, f.rel, f.dest_rel) : '';
        return `<tr${f.conflict || f.status === 'conflict' ? ' class="is-conflict"' : ''}><td>${fileLink(f.rel, base, applied && f.status === 'ok' ? f.dest_rel : f.rel)}${diff}</td><td>${action}</td><td>${sizeText}</td><td>${state}</td></tr>`;
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
    copy: '本地没有，拿过来', replace: '共享与本地文件内容不同，按拿取规则用共享这份替换本地；不代表是别人修改', keep: '你改过、目的地没人动，不拿', same: '共享与本地文件内容完全一致，无需替换（也可能是你已上传的版本），只记下版本',
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
        if (applied) state = f.status === 'ok' ? `<span class="link-ok">✅</span>${f.backup_path ? ` 你的那份备份在 ${esc(f.backup_path)}` : ''}${f.ref_backup_path ? `${f.backup_path ? '<br>' : ' '}旧参考备份在 ${esc(f.ref_backup_path)}` : ''}`
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
         [s['备份参考'] ? 'info' : '', `备份旧参考 ${s['备份参考']} 个`],
         ['info', `保留你改的 ${s['保留你改的'] ?? 0} 个`], ['info', `已一样 ${s['已一样'] ?? 0} 个`], ['ok', `记下版本 ${s['记下版本'] ?? 0} 个`],
         [s['没拿成功'] ? 'err' : '', `没拿成功 ${s['没拿成功']} 个`], [s['没有拿'] ? 'warn' : '', `没有拿 ${s['没有拿']} 个`],
         [s['有问题的文件夹'] ? 'err' : '', `找不到目的地等 ${s['有问题的文件夹']} 个文件夹`]]
      : [['info', `${s['文件夹'] ?? 0} 个文件夹`], ['info', `要拿 ${s['要拿'] ?? 0} 个（新拿 ${s['新拿'] ?? 0} · 替换 ${s['替换'] ?? 0}）`],
         [s['要备份'] ? 'warn' : '', `替换前备份你的 ${s['要备份']} 个`], ['info', `保留你改的 ${s['保留你改的'] ?? 0} 个`],
         ['info', `已一样 ${s['已一样'] ?? 0} 个`], [s['有问题'] ? 'err' : '', `有问题 ${s['有问题']} 个文件夹`]];
    const html = ['<div class="dist-summary">' + list.filter(([cls]) => cls).map(([cls, text]) => `<span class="dist-chip dist-chip-${cls}">${esc(text)}</span>`).join('') + '</div>'];
    if (applied && data.backup_root) html.push(`<p class="toolbar-hint">被替换的文件备份在：${esc(data.backup_root)}${data.task?.reference ? '（处理文件夹的在「处理」、参考文件夹的在「参考」子文件夹）' : ''}</p>`);
    if (applied && data.id) html.push(`<p><a class="btn btn-outline btn-sm" href="${api}/markdown?id=${encodeURIComponent(data.id)}">⬇ 下载拿取记录（Markdown）</a></p>`);
    html.push(`<p class="toolbar-hint">从目的地根目录：${esc(data.dest_root)}　拿回来源根目录：${esc(data.source_root)}</p>`);
    let person = null;
    for (const folder of data.folders || []) {
      if (folder.person !== person) { person = folder.person; html.push(`<h3 class="dist-person">👤 ${esc(person)}</h3>`); }
      html.push(fetchCard(folder, applied));
    }
    return html.join('');
  }

  const pdfSummary = s => s['PDF 已传'] == null ? '' : `同名 PDF 传了 ${s['PDF 已传']} 个`
    + (s['PDF 没覆盖'] ? `，终点已有没覆盖 ${s['PDF 没覆盖']} 个` : '') + (s['PDF 没传成'] ? `，没传成 ${s['PDF 没传成']} 个` : '') + '。';

  function renderUpload(data) {
    const s = data.summary || {};
    // 用户 2026-11-10：整批顺着跑完，没传上去的最后一起列出来
    const text = data.stopped || ((s['待重新处理'] || s['没复制成功'] || s['已跳过'])
      ? `这一批跑完了。${s['待重新处理'] ? `${s['待重新处理']} 个仍需在共享最新版上重新处理和复检。` : ''}`
        + (s['没复制成功'] ? `${s['没复制成功']} 个没传成，原因见下面。` : '')
        + (s['已跳过'] ? `${s['已跳过']} 个已跳过。` : '')
      : '本批文件已上传。');
    const rows = (data.files || []).map(file => {
      const status = file.status === 'ok' ? (file.pdf_only ? '✓ 只传 PDF，Excel 没传' : '✓ 已上传') : file.error || '未上传';
      const base = file.status === 'ok' ? data.dest_root : data.source_root;
      const diff = file.status !== 'ok' && data.shared ? diffLink(data.source_root, data.shared, file.rel) : '';
      return `<li>${fileLink(file.rel, base, file.rel)}${diff} · ${esc(status)}${file.content_check ? `<br><small>🔎 ${esc(file.content_check)}</small>` : ''}${prepHtml(file, data)}${pdfHtml(file)}${file.backup_path ? `<br><small>旧处理副本：${esc(file.backup_path)}</small>` : ''}${file.ref_backup_path ? `<br><small>旧参考副本：${esc(file.ref_backup_path)}</small>` : ''}</li>`;
    });
    const allDone = rows.length && (s['已复制'] || 0) === rows.length && !data.stopped;
    const pdfOnly = rows.length && data.files.every(file => file.pdf_only);
    return `<p>${pdfOnly ? `这批只传 PDF，处理了 ${s['已复制'] || 0} / ${rows.length} 个，Excel 都没传。` : `${esc(text)} 已上传 ${s['已复制'] || 0} / ${rows.length} 个。`}</p><ul class="dist-list">${rows.join('')}</ul>`
      + `<p class="next-step-row">${data.id ? `<a class="btn btn-outline btn-sm" href="${api}/markdown?id=${encodeURIComponent(data.id)}">⬇ 下载上传记录</a>` : ''}`
      + (allDone ? backHomeHtml() : '') + '</p>';
  }

  // 同名 PDF 这一份的结果（和后端 distribute._pdf_text 一样的说法）
  function pdfHtml(file) {
    const p = file.pdf;
    if (!p) return '';
    let text;
    if (p.status === 'ok') text = `同名 PDF ${p.name} 已${p.action === 'overwrite' ? '覆盖' : '新建'}（${p.sheets} 张表、${p.pages} 页${p.cached ? '，Excel 没变，沿用之前导出的' : ''}）`;
    else if (p.status === 'exported') text = `同名 PDF ${p.name} 已导出，但这份没传上去，PDF 也没传`;
    else if (p.status === 'stale') text = `同名 PDF ${p.name} 是按同步前那一版导的，没传，处理文件夹里那份也删掉了`;
    else text = p.error || '同名 PDF 没传';
    return `<br><small class="dist-pdf-line${['failed', 'skipped', 'kept'].includes(p.status) ? ' is-bad' : ''}">📕 ${esc(text)}</small>`;
  }

  // 上传前自动做的清标记 / 对照参考恢复颜色 / 视图标准化（core/upload_prep.py）
  function prepHtml(file, data) {
    const p = file.prep;
    if (!p) return '';
    const parts = [p.marks ? `清掉标记 ${p.marks} 条` : '没有标记要清'];
    if (p.tabs) parts.push(`还原表标签 ${p.tabs} 张`);
    if (p.no_reference) parts.push('没有参考，不对照颜色');
    else if (!p.reference) parts.push('参考里没有这份文件，没对照颜色');
    else parts.push(p.colors ? `按参考改回颜色 ${p.colors} 处` : '颜色和参考一致');
    if (p.normalized) parts.push('视图已标准化');
    const skipped = p.color_skipped ? `<details class="dist-prep-skip"><summary>颜色改不了 ${p.color_skipped} 处（点开看是哪些，请人工看一下）</summary><ul>`
      + (p.color_problems || []).map(cell => `<li>${fileLink(`${cell.sheet}!${cell.cell}`, data.source_root, file.rel, cell.sheet, cell.cell)} · ${esc(cell.reason)}</li>`).join('')
      + (p.color_skipped > (p.color_problems || []).length ? `<li>还有 ${p.color_skipped - p.color_problems.length} 处没列出</li>` : '') + '</ul></details>' : '';
    return `<br><small>🧽 ${file.prep_stale ? '同步前那一版做过的上传前处理（已作废）' : '上传前'}：${esc(parts.join('，'))}</small>${skipped}`;
  }

  const renderAny = (data, applied) => isFetch(data) ? renderFetch(data, applied)
    : isFree(data) ? renderFree(data, applied) : isUpload(data) ? renderUpload(data) : renderData(data, applied);

  // 起点 → 终点：清单（applied=false）和结果（applied=true）共用
  const FREE_ACTION = { new: ['tag-new', '新建'], overwrite: ['tag-overwrite', '覆盖'], review: ['tag-confirm', '待逐份确认'],
    conflict: ['dist-conflict', '不覆盖'], skipped: ['', '跳过'], failed: ['dist-conflict', '未上传'],
    confirm: ['tag-confirm', '确认覆盖'], force: ['dist-conflict', '强制覆盖'] };
  const FREE_GUARD = { manifest: '按拿取版本清单核对', baseline: '按相同起点的上传记录核对', none: '没有基准', content: '按文字和打印区域核对' };
  function renderFree(data, applied) {
    const s = data.summary || {};
    const chipsHtml = applied
      ? `<span class="dist-chip dist-chip-ok">已上传 ${s['已复制'] || 0}</span>`
        + (s['已跳过'] ? `<span class="dist-chip dist-chip-info">已跳过 ${s['已跳过']}</span>` : '')
        + (s['别人改过没覆盖'] ? `<span class="dist-chip dist-chip-err">🛑 别人改过没覆盖 ${s['别人改过没覆盖']}</span>` : '')
        + (s['没复制成功'] ? `<span class="dist-chip dist-chip-warn">没上传成功 ${s['没复制成功']}</span>` : '')
      : `<span class="dist-chip dist-chip-info">要上传 ${s['要上传']}（新建 ${s['新建']}、覆盖 ${s['覆盖']}）</span>`
        + (s['要确认'] ? `<span class="dist-chip dist-chip-warn">待逐份确认 ${s['要确认']}</span>` : '')
        + (s['别人改过'] ? `<span class="dist-chip dist-chip-err">🛑 不覆盖 ${s['别人改过']}</span>` : '')
        + `<span class="dist-chip ${s['有清单'] ? 'dist-chip-ok' : 'dist-chip-info'}">${s['有清单'] ? `${s['有清单']} 个按拿取版本清单核对` : '起点里没有拿取版本清单'}</span>`;
    const notes = (data.notes || []).map(note => `<li>${esc(note)}</li>`).join('');
    const rows = (data.files || []).map(file => {
      const action = applied && file.status !== 'ok' ? file.status || file.action
        : applied && file.review_confirmed ? (file.dest_sha256 ? 'overwrite' : 'new') : file.action;
      const [cls, label] = FREE_ACTION[action] || ['', action];
      const state = !applied ? '' : file.status === 'ok' ? (file.pdf_only ? '✓ 只传 PDF，Excel 没传' : '✓ 已上传') : file.status === 'conflict' ? '' : esc(file.error || '未上传');
      const why = file.action === 'conflict' || file.status === 'conflict' ? esc(file.error || '')
        : esc(file.action === 'force' ? file.forced_reason || '' : file.note || '');
      const link = applied && file.status === 'ok' ? fileLink(file.rel, data.dest_root, file.rel) : fileLink(file.rel, data.source_root, file.rel);
      const diff = file.dest_sha256 && file.status !== 'ok' ? diffLink(data.source_root, data.dest_root, file.rel, file.rel, '', !applied && file.action !== 'conflict' ? data.id : '') : '';
      return `<tr class="${file.action === 'conflict' || file.status === 'conflict' ? 'is-conflict' : ''}"><td>${link}${diff}</td><td><span class="${cls}">${label}</span></td>`
        + `<td>${state}${state && why ? '<br>' : ''}${why ? `<small>${why}</small>` : ''}${prepHtml(file, data)}${pdfHtml(file)}</td><td><small>${FREE_GUARD[file.guard] || ''}</small></td></tr>`;
    });
    const list = (title, items) => items && items.length ? `<h4 class="dist-sub">${title}（${items.length}）</h4><ul class="dist-list">${items.map(rel => `<li>${fileLink(rel, data.dest_root, rel)}</li>`).join('')}</ul>` : '';
    const relayTask = data.relay_task;
    const manifestLine = !relayTask ? ''
      : data.manifest_error ? `<p class="link-error">⚠️ 拿取版本.json 没写成：${esc(data.manifest_error)}</p>`
      : applied && data.manifest ? `<p class="dist-meta">🧾 已在终点写好拿取版本.json：${esc(data.manifest)}（从这里传回共享时用来核对，请不要删）</p>`
      : applied ? ''
      : `<p class="dist-meta">🧾 起点是任务「${esc(relayTask.name)}」的处理文件夹，传完会在终点写拿取版本.json，记下每份当初从共享拿的版本。</p>`;
    const unknownLine = applied && (data.manifest_unknown || []).length
      ? `<p class="link-error">⚠️ 这几份找不到当初从共享拿的版本，拿取版本.json 里没记：${data.manifest_unknown.map(esc).join('、')}</p>` : '';
    return `<p class="dist-meta">起点：${esc(data.source_root)}${(data.picks || []).length ? `（贴的 ${data.picks.length} 处的共同上级，只传贴的这些）` : ''}<br>终点：${esc(data.dest_root)}</p><div class="dist-summary">${chipsHtml}</div>`
      + manifestLine + unknownLine
      + (notes ? `<div class="dist-alert"><ul>${notes}</ul></div>` : '')
      + `<div class="link-table-wrap"><table class="link-table"><thead><tr><th>文件</th><th>做法</th><th>${applied ? '结果' : '说明'}</th><th>怎么核对的</th></tr></thead><tbody>${rows.join('')}</tbody></table></div>`
      + list('期间别人新加到终点的 Excel（不动）', data.added) + list('终点里有、起点没有的 Excel（不动）', data.extra)
      + (data.errors || []).map(err => `<p class="link-error">${esc(err)}</p>`).join('')
      + (applied && data.id ? `<p class="next-step-row"><a class="btn btn-outline btn-sm" href="${api}/markdown?id=${encodeURIComponent(data.id)}">⬇ 下载上传记录</a></p>` : '');
  }

  function renderHistory() {
    const runs = (latest.history || []).filter(inScope);
    if (!runs.length) { $('distHistory').innerHTML = '<p>还没有记录。</p>'; return; }
    $('distHistory').innerHTML = '<ul class="dist-history">' + runs.map(run => {
      const s = run.summary || {};
      const chip = run.attention ? (s['别人改过没覆盖'] ? '<span class="dist-chip dist-chip-err">🛑 有别人改过的</span>' : '<span class="dist-chip dist-chip-warn">⚠️ 有对不上的地方</span>')
        : '<span class="dist-chip dist-chip-ok">✓ 全部一致</span>';
      const what = isFree(run) ? `🚚 起点 → 终点 上传 ${s['已复制'] ?? 0} 个${s['别人改过没覆盖'] ? `，别人改过没覆盖 ${s['别人改过没覆盖']}` : ''}`
        : isUpload(run) ? `上传 ${s['已复制'] ?? 0} 个${s['待重新处理'] ? ' · 待重新处理 ' + s['待重新处理'] + ' 个' : ''}`
        : run.kind === 'fetch' ? `⬇ 拿最新版 ${s['已拿'] ?? 0} 个（替换 ${s['替换'] ?? 0}）`
        : `复制 ${s['已复制'] ?? 0} 个（覆盖 ${s['覆盖'] ?? 0}${s['别人改过没覆盖'] ? `，别人改过没覆盖 ${s['别人改过没覆盖']}` : ''}）`;
      return `<li><span class="dist-route">${esc(run.label || '（没有 label）')}<small>${esc(run.finished || '')} · ${what} · ${esc(run.dest_root || '')}</small></span>${chip}`
        + `<button type="button" class="button secondary" data-dist-view="${esc(run.id)}">查看</button>`
        + `<a class="button secondary" href="${api}/markdown?id=${encodeURIComponent(run.id)}">⬇ 下载</a></li>`;
    }).join('') + '</ul>';
  }

  // 实时进度框（static/job_panel.js，用户 2026-11-09 要：所有批量操作都要有）：正在拿 / 传哪份、每份的结果
  function renderPanel() {
    if (!window.JobPanel || !latest.kind) return;
    const id = `dist:${latest.kind}`;
    const what = isFree(latest) ? '起点 → 终点上传' : isUpload(latest) ? '检查并批量上传' : isFetch(latest) ? '从共享拿最新版' : '按 JSON 复制';
    const log = latest.log || [];
    if (latest.running) {
      const now = String(latest.current || '');
      const step = now.match(/^(检查|上传前处理|导出同名 PDF|上传同名 PDF|上传)\s/)?.[1] || '';
      window.JobPanel.run(id, {
        title: `正在${what}`, count: window.JobPanel.count(latest.done, latest.total, '个文件'),
        now: now ? `📗 ${step ? now.slice(step.length).trim() : now}` : '准备中…',
        step: step ? `正在${step}` : '', fraction: latest.total ? latest.done / latest.total : 0, log,
      });
    } else {
      const bad = log.filter(i => !i.ok).length;
      const stopped = latest.applied?.stopped;
      window.JobPanel.finish(id, {
        title: stopped ? `${what}停下了` : `${what}做完了${bad ? `，${bad} 个要看一下` : ''}`,
        warn: !!stopped || bad > 0, log, now: stopped || undefined,
        step: bad || stopped ? '详情看页面上标黄的文件夹或下载记录。' : '',
      });
    }
  }

  let visibleReviews = '';
  function renderUploadReviews() {
    const matches = freeMode ? isFree(latest) : !jsonMode && !startMode && isUpload(latest) && latest.review_task_id === task?.id;
    const reviews = matches ? latest.reviews || [] : [];
    $('distUploadReviews').hidden = !reviews.length;
    const keys = reviews.map(review => review.token).join(',');
    if (visibleReviews === keys) return;
    visibleReviews = keys;
    $('distUploadReviewList').innerHTML = reviews.map(review => `<li><strong>${esc(review.rel)}</strong> · ${esc(review.reason)}
      <a class="btn btn-outline btn-sm" href="/excel-diff?review=${encodeURIComponent(review.token)}" target="_blank" rel="noopener">${review.missing ? '查看并确认' : '查看差异并确认'}</a>
      <button class="btn btn-outline btn-sm" type="button" data-upload-skip="${esc(review.token)}">跳过</button></li>`).join('');
  }

  $('distUploadReviewList').addEventListener('click', async event => {
    const button = event.target.closest('button[data-upload-skip]');
    if (!button) return;
    button.disabled = true;
    $('distUploadReviewError').textContent = '';
    try {
      await post(api + '/upload_decision', { token: button.dataset.uploadSkip, decision: 'skip' });
      await refresh();
    } catch (error) {
      $('distUploadReviewError').textContent = error.message;
      button.disabled = false;
    }
  });

  function render() {
    renderUploadReviews();
    controls();
    renderHistory();
    renderPanel();
    if (latest.running) {
      showProgress('running');
      const uploading = isUpload(latest) || isFree(latest);
      const detail = latest.comparing ? `正在比对 ${latest.comparing}`
        : latest.reviews?.length ? `${latest.reviews.length} 份等待确认` : latest.current || '准备中…';
      $('distProgress').textContent = `${uploading ? '检查并上传' : isFetch(latest) ? '正在拿最新版' : '正在复制'} ${latest.done || 0} / ${latest.total || 0}：${detail}`;
      $('distProgressFill').style.width = `${latest.total ? latest.done / latest.total * 100 : 0}%`;
      window.ProgressEta?.update('distProgressFill', `dist:${latest.kind}`, latest.total ? latest.done / latest.total : 0, true);
      $('distResultTitle').textContent = uploading ? '上传进度' : isFetch(latest) ? '拿取结果' : '复制结果';
      $('distResults').textContent = uploading ? '已完成的比对会在这里逐份显示。' : '后台处理中，完成后显示结果。';
      return;
    }
    window.ProgressEta?.update('distProgressFill', `dist:${latest.kind}`, 0, false);
    if (viewing) {
      $('distResultTitle').textContent = `${isUpload(viewing) ? '上传记录' : isFetch(viewing) ? '拿取记录' : '复制记录'}：${viewing.label || ''}`;
      $('distResults').innerHTML = '<p><button type="button" class="btn btn-outline btn-sm" data-dist-back>← 回到当前清单</button></p>' + renderAny(viewing, true);
      return;
    }
    const applied = latest.applied;
    const data = latest.preview || applied;
    if (!data) {
      $('distResultTitle').textContent = freeMode ? '上传清单' : jsonMode ? '复制清单' : startMode ? '拿取清单' : '批量上传结果';
      $('distResults').textContent = jsonMode ? '尚未生成清单。' : freeMode ? '填好起点和终点后点「📋 生成上传清单」（只读不改）。'
        : startMode ? (task ? '点「⬇ 从共享拿最新版」先生成拿取清单（只读不改），核对后再拿。' : '先在上面开一个新任务。')
        : task ? '选好文件后，点击「检查并批量上传」。' : '还没有上传结果。';
      return;
    }
    $('distProgressFill').style.width = '100%';
    if (isFree(data)) {
      const s = data.summary || {};
      if (latest.preview) {
        $('distResultTitle').textContent = '上传清单';
        showProgress('done', '清单已生成：', `要上传 ${s['要上传']} 个（新建 ${s['新建']}、覆盖 ${s['覆盖']}）`
          + (s['别人改过'] ? `；🛑 ${s['别人改过']} 个别人改过，不会覆盖` : '') + '。核对后点「🚚 按清单上传」。', !!(s['别人改过'] || s['要确认'] || data.notes.length));
      } else {
        const warn = !!(data.stopped || s['别人改过没覆盖'] || s['没复制成功']);
        $('distResultTitle').textContent = '上传结果';
        showProgress('done', data.stopped ? '没有上传：' : warn ? '上传结束，有没传上去的：' : '上传完成：',
          data.stopped || `已上传 ${s['已复制'] || 0} / ${data.files.length} 个。${pdfSummary(s)}`, warn || !!s['PDF 没传成']);
      }
      $('distResults').innerHTML = renderFree(data, !latest.preview);
      return;
    }
    if (isUpload(data)) {
      const s = data.summary || {};
      const paused = data.stopped || s['待重新处理'] || s['没复制成功'];
      $('distResultTitle').textContent = '批量上传结果';
      const pdfOnly = data.files.length && data.files.every(file => file.pdf_only);
      showProgress('done', data.stopped ? '上传停下了。' : paused ? '这一批跑完了，有没传上去的：' : '上传完成。', pdfOnly ? ` 只传 PDF，Excel 没传。${pdfSummary(s)}`
        : ` 已上传 ${s['已复制'] || 0} / ${data.files.length} 个。${pdfSummary(s)}`, !!(paused || s['PDF 没传成']));
      $('distResults').innerHTML = renderUpload(data);
      return;
    }
    if (isFetch(data)) {
      const s = data.summary || {};
      if (latest.preview) {
        $('distResultTitle').textContent = `拿取清单${data.label ? '：' + data.label : ''}`;
        showProgress('done', '拿取清单已生成：', `要从目的地拿 ${s['要拿']} 个（新拿 ${s['新拿']}、替换本地 ${s['替换']}，其中 ${s['要备份']} 个先备份你的），保留你改的 ${s['保留你改的']} 个，已一样 ${s['已一样']} 个。核对后点「⬇ 按清单拿文件」。`, !!s['有问题']);
      } else {
        const warn = !!(data.stopped || s['没拿成功'] || s['没有拿'] || s['有问题的文件夹']);
        $('distResultTitle').textContent = `拿取结果${data.label ? '：' + data.label : ''}`;
        showProgress('done', data.stopped ? '没有拿：' : warn ? '拿完了，有没拿成功的：' : '拿完了：',
          data.stopped || `拿了 ${s['已拿'] ?? 0} 个（替换 ${s['替换'] ?? 0}、备份你的 ${s['备份'] ?? 0}${s['备份参考'] ? `、备份旧参考 ${s['备份参考']}` : ''}），记下 ${s['记下版本'] ?? 0} 个文件的版本，复制前会用它检查别人改过没有。`, warn);
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

  // 上传跑完：有没传上去的（共享有更新已同步、出错）一起列出来（用户 2026-11-10：不要一个一个卡住）
  function uploadDone(data) {
    const s = data.summary || {};
    const review = data.files.filter(file => file.status === 'review'), failed = data.files.filter(file => ['failed', 'conflict'].includes(file.status));
    console.info('📦 批量上传跑完', { 已上传: s['已复制'] || 0, 要重新处理: review.map(f => f.rel), 没传成: failed.map(f => `${f.rel}：${f.error}`) });
    if (!review.length && !failed.length) return;
    const list = files => files.slice(0, 10).map(file => '　' + file.rel).join('\n') + (files.length > 10 ? `\n　…还有 ${files.length - 10} 个` : '');
    const lines = [`已上传 ${s['已复制'] || 0} 个。`];
    if (review.length) lines.push(`\n🔄 ${review.length} 个仍需在共享最新版上重新处理和复检：\n${list(review)}`);
    if (failed.length) lines.push(`\n⚠️ ${failed.length} 个没传成（原因见页面上的结果）：\n${list(failed)}`);
    window.appAlert(lines.join('\n'), { title: '这一批跑完了，有几个要再处理', tone: 'warn' });
  }

  // 拿完文件：成功撒花（game_fx.js），再弹提示说拿了多少、下一步做什么；有没拿成的只弹黄色提示
  function fetchDone(data) {
    const s = data.summary || {};
    const got = s['已拿'] ?? 0;
    if (data.stopped) {
      window.appAlert(`没有拿：${data.stopped}`, { title: '没有拿文件', tone: 'warn' });
      return;
    }
    if (s['没拿成功'] || s['没有拿'] || s['有问题的文件夹']) {
      window.appAlert(`拿了 ${got} 个，但有 ${(s['没拿成功'] || 0) + (s['没有拿'] || 0)} 个文件没拿成功、${s['有问题的文件夹'] || 0} 个文件夹有问题。\n请看页面上标黄的文件夹（常见原因：文件正在 Excel 里打开），处理后再点「⬇ 从共享拿最新版」。`, { title: '拿完了，有没拿成功的', tone: 'warn' });
      return;
    }
    const lines = got
      ? [`从共享拿了 ${got} 个 Excel 到处理文件夹（新拿 ${s['新拿'] ?? 0}、替换 ${s['替换'] ?? 0}${s['备份'] ? `，替换前备份了你的 ${s['备份']} 个` : ''}${s['备份参考'] ? `，参考里 ${s['备份参考']} 个旧版也备份了` : ''}）。`,
         task ? '参考文件夹里也放了一份原件，修复颜色 / 引用 / 隐藏页时用来对照。' : '']
      : ['共享和本地已经一样，这次没有要拿的文件。'];
    if (s['保留你改的']) lines.push(`${s['保留你改的']} 个你改过、共享没人动的文件保留了你改的。`);
    if (got) window.gameFx?.celebrate('拿完了！', `${got} 个文件已拿到处理文件夹`, 3);
    // 先让横幅和撒花播一会儿，再弹提示
    setTimeout(async () => {
      const go = await window.appConfirm(lines.filter(Boolean).join('\n'), { title: '文件拿完了', tone: 'info', icon: '✓', okText: '🏠 返回首页', cancelText: '留在这页' });
      if (go) location.href = homeUrl();
    }, got ? 2600 : 0);
  }

  async function refresh() {
    clearTimeout(timer);
    let failed = false;
    try {
      const response = await fetch(api, { cache: 'no-store' });
      if (!response.ok) throw Error('读取任务进度失败，正在重试');
      latest = await response.json();
      const scope = jsonMode || freeMode ? '' : task?.id;
      if (latest.preview && (latest.preview.task?.id || '') !== scope) latest.preview = null;
      if (latest.applied && (latest.applied.task?.id || '') !== scope) latest.applied = null;
      if (!inScope(latest.preview)) latest.preview = null;
      if (!inScope(latest.applied)) latest.applied = null;
      if (freeMode && loading && isFree(latest.preview)) {  // 刷新页面：接着用还没执行的清单
        $('distFreeSource').value = latest.preview.source_text || latest.preview.source_root;
        $('distFreeDest').value = latest.preview.dest_root;
        needsFreePreview = false;
      }
      if (pollError && $('distError').textContent === pollError) $('distError').textContent = '';
      pollError = '';
      loading = false;
      if (task && !startMode && wasRunning && !latest.running) {
        if (isUpload(latest.applied)) {
          for (const file of latest.applied.files) if (file.status === 'ok') selected.delete(file.rel);
        }
        await reloadTaskFiles();
        if (isUpload(latest.applied) && !latest.applied.stopped) uploadDone(latest.applied);
      }
      if (wasRunning && !latest.running && isFetch(latest.applied)) fetchDone(latest.applied);
      if (wasRunning && !latest.running && isFree(latest.applied)) {
        const s = latest.applied.summary || {};
        if (s['别人改过没覆盖'] || s['没复制成功'] || latest.applied.stopped) {
          window.appAlert(latest.applied.stopped || `已上传 ${s['已复制'] || 0} 个。${s['别人改过没覆盖'] ? `\n${s['别人改过没覆盖']} 个需要重新生成清单并检查。` : ''}${s['没复制成功'] ? `\n${s['没复制成功']} 个没传上去，请看下面的说明。` : ''}`,
            { title: '有文件没传上去', tone: 'warn' });
        } else if (s['已复制']) {
          window.gameFx?.celebrate('传完了！', `${s['已复制']} 个文件已传到终点`, 3);
        }
        needsFreePreview = true;
      }
      if (wasRunning && !latest.running && latest.applied && !latest.applied.stopped && !isFetch(latest.applied) && !isUpload(latest.applied) && !isFree(latest.applied)) {
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
    if (loading || submitting || latest.running || (!jsonMode && !task)) return;
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
    if (task) lines.push('拿下来的原件同时放一份到参考文件夹（修复颜色 / 引用 / 隐藏页对照用），参考里原来的旧版也先备份。');
    if (s['保留你改的']) lines.push(`${s['保留你改的']} 个你改过、目的地没人动，不拿，保留你改的。`);
    lines.push('所有文件都会记下拿到的版本，之后复制（覆盖）前用来检查别人改过没有。目的地的文件不会被改动。');
    return window.appConfirm(lines.join('\n'), { title: '按清单拿最新版', okText: '开始拿', tone: s['替换'] ? 'danger' : undefined });
  }

  async function startCopy() {
    if (loading || submitting || latest.running || needsPreview || !latest.preview) return;
    const preview_id = latest.preview.id;
    if (isFetch(latest.preview)) {
      if (!await startFetch()) return;
      submitting = 'start'; viewing = null; controls();
      try { await post(api + '/start', { preview_id }); wasRunning = true; } catch (error) { $('distError').textContent = error.message; }
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
      await post(api + '/start', { preview_id });
      filter = 'all';
      wasRunning = true;
    } catch (error) {
      $('distError').textContent = error.message;
    } finally { submitting = false; await refresh(); }
  }

  async function startTaskUpload() {
    if (!task || loading || submitting || latest.running || readingFiles) return;
    const files = taskFiles.filter(file => selected.has(file.rel));
    if (!files.length) return;
    const reviews = files.filter(file => file.review?.ready);
    const destination = task.upload_to ? '中转文件夹「' + task.upload_to + '」' : '当前任务的共享文件夹';
    const legacy = reviews.length ? '请确认此前同步的文件已经保存并完成复检。\n\n' : '';
    const message = legacy + '检查选中的 ' + files.length + ' 个 Excel，检查通过的文件按本次设置上传到' + destination
      + '。需要复核的文件会逐份显示差异，确认一份就可以上传一份。';
    submitting = 'upload';
    controls();
    try {
      const pdf = await pdfPlan({ task_id: task.id, files: files.map(file => file.rel) });
      if (pdf === null) return;
      if (!await window.appConfirm(message + (pdf ? '\n\n' + pdf.line : ''), { title: pdf?.only ? '只传 PDF' : task.upload_to ? '批量上传到中转' : '批量上传到共享', okText: '开始检查', tone: 'danger' })) return;
      viewing = null;
      $('distError').textContent = '';
      const reviewed = Object.fromEntries(reviews.map(file => [file.rel, file.review.token]));
      await post(api + '/task_upload', { task_id: task.id, files: files.map(file => file.rel), reviewed, guard: guardMode(),
        pdf: pdf ? { options: pdf.options, overwrite: pdf.overwrite, only: pdf.only } : null });
      wasRunning = true;
    } catch (error) {
      $('distError').textContent = error.message;
    } finally {
      submitting = false;
      await refresh();
    }
  }

  // ───── 本轮只处理待重新处理文件（后端 work_task.set_focus；前 3 步只列这些文件，全部上传后自动结束） ─────
  const focusList = () => task?.focus?.files || [];
  const readyRels = () => taskFiles.filter(file => file.review?.ready).map(file => file.rel);
  const sameSet = (a, b) => a.length === b.length && a.every(rel => b.map(x => x.toLowerCase()).includes(rel.toLowerCase()));
  function focusControls() {
    if (!task || startMode) return;
    const ready = readyRels(), focus = focusList();
    $('distTaskFocusRow').hidden = !ready.length && !focus.length;
    const on = focus.length && sameSet(ready, focus);
    $('distTaskFocus').hidden = !ready.length || on;
    $('distTaskFocus').textContent = focus.length ? `🎯 把本轮清单换成这 ${ready.length} 个待重新处理文件`
      : `🎯 让 前 3 步只处理这 ${ready.length} 个待重新处理文件`;
    $('distTaskFocusNote').textContent = on ? `✓ 已开启：前 3 步这一轮只处理这 ${focus.length} 个文件。处理、复检完回来上传，传完自动回到整个文件夹。`
      : focus.length ? `本轮清单现在是 ${focus.length} 个文件（看页面顶部），和待重新处理的不一样。`
      : '其他文件这一轮不动，省得整个文件夹重跑一遍。';
  }
  async function startFocus() {
    const ready = readyRels();
    if (!task || !ready.length) return;
    const list = ready.slice(0, 8).join('\n') + (ready.length > 8 ? '\n…' : '');
    if (!await window.appConfirm(`这一轮 前 3 步（精简 Sheet、调整页面、查看切字）只处理这 ${ready.length} 个文件，其他文件不动：\n\n${list}\n\n`
      + '这几个全部上传成功后自动回到整个文件夹；也可以随时点页面顶部的「退出」。', { title: '本轮只处理待重新处理文件', okText: '开启', cancelText: '先不用' })) return;
    try {
      const result = await post('/api/tasks/focus', { id: task.id, files: ready });
      task.focus = result.task.focus;
    } catch (error) { $('distError').textContent = error.message; return; }
    focusControls();
    if (window.refreshFocusBanner) await window.refreshFocusBanner();
    if (await window.appConfirm(`已开启。回首页处理这 ${ready.length} 个文件，处理、复检完回这页上传。`,
      { title: '开启了', okText: '🏠 返回首页', cancelText: '留在这页' })) location.href = homeUrl();
  }
  $('distTaskFocus').addEventListener('click', startFocus);
  document.addEventListener('focus-changed', () => { if (task) { task.focus = null; focusControls(); } });

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
  async function makeFreePreview() {
    if (loading || submitting || latest.running) return;
    const source = $('distFreeSource').value.trim(), dest = $('distFreeDest').value.trim();
    $('distError').textContent = '';
    if (!source || !dest) { $('distError').textContent = '请填写起点和终点'; return; }
    save(KEYS.freeSource, source); save(KEYS.freeDest, dest);
    submitting = 'free'; viewing = null; controls();
    showProgress('running', '', '正在列出起点和终点的文件，检查终点那份期间有没有被人改过…');
    try {
      await post(api + '/free_preview', { source, dest, guard: guardMode() });
      needsFreePreview = false;
    } catch (error) { $('distError').textContent = error.message; }
    finally { submitting = false; await refresh(); }
  }

  async function startFree() {
    const p = latest.preview;
    if (!isFree(p) || loading || submitting || latest.running || needsFreePreview) return;
    const s = p.summary;
    submitting = 'start'; controls();
    try {
      const pdf = await pdfPlan({ preview_id: p.id });
      if (pdf === null) return;
      const lines = [pdf?.only ? '上传 ' + s['要上传'] + ' 份同名 PDF。'
        : '从「' + p.source_root + '」传到「' + p.dest_root + '」，共 ' + s['要上传'] + ' 个 Excel。'];
      if (!pdf?.only && s['要确认']) lines.push(s['要确认'] + ' 份需要逐个查看差异并确认，完成一份就可以上传一份。');
      if (pdf) lines.push(pdf.line);
      if (!await window.appConfirm(lines.join('\n'), { title: '按清单上传', okText: '开始检查', tone: 'danger' })) return;
      viewing = null;
      await post(api + '/free_start', { preview_id: p.id,
        pdf: pdf ? { options: pdf.options, overwrite: pdf.overwrite, only: pdf.only } : null });
      wasRunning = true;
    } catch (error) { $('distError').textContent = error.message; }
    finally { submitting = false; await refresh(); }
  }
  $('distFreePreview').addEventListener('click', makeFreePreview);
  $('distFreeStart').addEventListener('click', startFree);
  for (const id of ['distFreeSource', 'distFreeDest']) {
    $(id).addEventListener('input', () => { needsFreePreview = true; controls(); });
    // 起点是多行框：回车换行，Ctrl+Enter 才生成清单
    $(id).addEventListener('keydown', event => { if (event.key === 'Enter' && (id !== 'distFreeSource' || event.ctrlKey || event.metaKey)) makeFreePreview(); });
  }
  $('distPreview').addEventListener('click', () => makePreview('preview'));
  $('distFetch').addEventListener('click', () => makePreview('fetch'));
  $('distStart').addEventListener('click', startCopy);
  $('distTaskUpload').addEventListener('click', startTaskUpload);
  $('distTaskReload').addEventListener('click', reloadTaskFiles);
  $('distTaskFiles').addEventListener('change', event => {
    const box = event.target;
    if (box.checked) selected.add(box.value);
    else selected.delete(box.value);
    controls();
  });
  // 用户 2026-11-10：不要「前 N 个」，要么全选、要么自己勾。整批会顺着跑完，选多少都不会卡住
  // 只传 PDF：待上传 = 没传过 PDF，或传过之后 Excel 又传了新版；其他按 Excel 的当前版本传没传过
  function waitingFilter() {
    return pdfMode() === 'pdf' ? file => file.review || !file.pdf_uploaded_at || pdfStale(file)
      : file => file.review || !file.uploaded_current;
  }
  $('distTaskPickWaiting').addEventListener('click', () => {
    selected.clear();
    const waiting = taskFiles.filter(waitingFilter());
    for (const file of waiting) selected.add(file.rel);
    console.info('☑ 全选待上传', { 按: pdfMode() === 'pdf' ? 'PDF 有没有传过' : 'Excel 有没有传过当前版本', 个数: waiting.length });
    controls();
  });
  $('distTaskPickAll').addEventListener('click', () => {
    for (const file of taskFiles) selected.add(file.rel);
    console.info('☑ 全选全部', taskFiles.length);
    controls();
  });
  $('distTaskClear').addEventListener('click', () => { selected.clear(); controls(); });
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
  function showTaskPaths() {
    $('distTaskPaths').innerHTML = `<dt>共享</dt><dd>${esc(task.shared)}</dd><dt>参考</dt><dd>${esc(task.reference)}</dd><dt>处理</dt><dd>${esc(task.work)}</dd>`
      + ((task.picks || []).length ? `<dt>只拿</dt><dd class="dist-picks"><ul>${task.picks.map(p => `<li title="${esc(p)}">${esc(p)}</li>`).join('')}</ul></dd>` : '')
      + (task.upload_to ? `<dt>上传到</dt><dd>${esc(task.upload_to)}（中转）</dd>` : '');
    $('distTaskDest').placeholder = task.shared;
    $('distTaskDest').value = task.upload_to || '';
    $('distTaskDest').closest('.dist-dest-row').classList.toggle('is-relay', !!task.upload_to);
    $('distTaskDestNote').textContent = task.upload_to
      ? `现在传到中转：${task.upload_to}。公司共享那份照样先核对版本，但不会被改；中转里会放一份「拿取版本.json」，汇总的人上传到公司时用它检查期间有没有人改过，请不要删。`
      : `现在直接传回共享文件夹：${task.shared}`;
  }

  async function saveTaskDest() {
    if (!task || latest.running) return;
    $('distError').textContent = '';
    try {
      const result = await post('/api/tasks/upload_to', { id: task.id, path: $('distTaskDest').value.trim() });
      task = { ...task, upload_to: result.task.upload_to };
      showTaskPaths();
      await reloadTaskFiles();
    } catch (error) { $('distError').textContent = error.message; }
  }
  if (!jsonMode && !startMode && !freeMode) {
    $('distTaskDestSave').addEventListener('click', saveTaskDest);
    $('distTaskDest').addEventListener('keydown', event => { if (event.key === 'Enter') saveTaskDest(); });
  }

  async function loadTask() {
    try {
      const response = await fetch('/api/tasks', { cache: 'no-store' });
      const result = await response.json();
      if (!response.ok || !result.ok) throw Error('无法读取当前任务');
      const current = result.tasks.find(t => t.id === result.current);
      if (jsonMode || freeMode) {
        $('distBackToTask').hidden = !current;
        if (freeMode) $('distGoUpload').hidden = !current;
        return;
      }
      if (startMode) $('distNewRoot').value = result.last_work_root || '';
      if (!current) {
        if (startMode) { showLastShared(result.tasks[0]?.shared); focusShared(); }
        $('distTaskNoticeText').textContent = startMode ? '还没有当前任务。在下面填共享文件夹路径，开一个新任务。'
          : '还没有当前任务。请先开一个新任务，再选择处理好的 Excel。';
        $('distTaskHome').hidden = startMode;
        return;
      }
      task = current;
      // 已经有任务：开新任务的表单先收起来，拿文件在前面
      if (startMode) {
        $('distNewBox').open = false;
        $('distNewSummary').textContent = '当前任务是「' + task.name + '」。要换一个共享文件夹？展开开一个新任务';
        $('distNewTaskTitle').textContent = '开另一个新任务';
        // 预先填上次用的共享路径，展开时全选，直接粘贴就能覆盖
        showLastShared(task.shared);
      }
      $('distTaskName').textContent = task.name;
      showTaskPaths();
      $('distTaskNotice').hidden = true;
      $('distTaskSummary').hidden = false;
      if (startMode) {
        $('distActions').hidden = false;
        $('distProgress').textContent = '';
      } else {
        $('distTaskBox').hidden = false;
        $('distProgress').textContent = '等待选择文件。';
      }
    } catch (_) {
      if (!jsonMode) $('distTaskNoticeText').textContent = '无法读取当前任务，请刷新页面重试。';
    }
  }

  // ───── 开新任务（只填共享路径，参考 / 处理文件夹自动生成，后端 core/work_task.py） ─────
  let deriveTimer = 0, deriveSeq = 0;
  // 没填时也占好三行，填好后换成真路径，这块高度不变
  const NEW_EMPTY = ['共享', '参考', '处理'].map(name => `<dt>${name}</dt><dd class="is-empty">填好共享文件夹路径后显示</dd>`).join('');
  // 框里不预填，上次的路径只放在占位提示里：聚焦后提示还在，直接粘贴新路径
  const showLastShared = last => {
    if (!last) return;
    $('distNewLastHead').textContent = '上次用过';
    $('distNewLast').textContent = last;
    $('distNewLast').closest('.dist-last').title = `上次：${last}`;
  };
  const pathsHtml = t => [['共享', t.shared], ['参考', t.reference], ['处理', t.work]].map(([name, path]) => `<dt>${name}</dt><dd title="${esc(path)}">${esc(path)}</dd>`).join('')
    // 填了多个路径：共享取共同上级，只拿下面这些（用户 2026-11-12）
    + ((t.picks || []).length ? `<dt>只拿</dt><dd class="dist-picks"><ul>${t.picks.map(p => `<li title="${esc(p)}">${esc(p)}</li>`).join('')}</ul></dd>` : '');
  function scheduleDerive() {
    clearTimeout(deriveTimer);
    deriveTimer = setTimeout(async () => {
      const seq = ++deriveSeq;
      const shared = $('distNewShared').value.trim();
      $('distNewError').textContent = '';
      if (!shared) { $('distNewPaths').innerHTML = NEW_EMPTY; return; }
      try {
        const result = await post('/api/tasks/derive', { shared, work_root: $('distNewRoot').value.trim() });
        if (seq !== deriveSeq) return;
        $('distNewPaths').innerHTML = pathsHtml(result.paths);
        if (!$('distNewRoot').value.trim()) $('distNewRoot').placeholder = result.paths.work_root;
      } catch (error) {
        if (seq !== deriveSeq) return;
        $('distNewPaths').innerHTML = NEW_EMPTY;
        $('distNewError').textContent = error.message;
      }
    }, 350);
  }
  async function createTask() {
    if (creating || latest.running) return;
    const shared = $('distNewShared').value.trim();
    if (!shared) { $('distNewError').textContent = '请填写共享文件夹路径'; $('distNewShared').focus(); return; }
    creating = true;
    controls();
    try {
      const { task: made } = await post('/api/tasks/new', { shared, work_root: $('distNewRoot').value.trim() });
      // 前 3 步默认用这个任务的处理 / 参考文件夹
      try {
        localStorage.setItem('clear_handle_excel.lastFolder', made.work);
        localStorage.setItem('clear_handle_excel.oldFolder', made.reference);
      } catch (_) {}
      location.href = '/distribute/start?fetch=1';
    } catch (error) {
      $('distNewError').textContent = error.message;
      creating = false;
      controls();
    }
  }
  const focusShared = () => { const input = $('distNewShared'); input.focus({ preventScroll: true }); input.select(); };
  if (startMode) {
    $('distNewBox').addEventListener('toggle', () => { if ($('distNewBox').open) focusShared(); });
    $('distNewShared').addEventListener('input', scheduleDerive);
    $('distNewRoot').addEventListener('input', scheduleDerive);
    $('distNewShared').addEventListener('keydown', event => { if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) createTask(); });
    $('distNewCreate').addEventListener('click', createTask);
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
    if (task && !startMode) await reloadTaskFiles();
    // 刚开了新任务：先自动生成拿取清单（确认后才真正拿）
    if (startMode && task && params.get('fetch') === '1' && !latest.running) {
      history.replaceState(null, '', '/distribute/start');
      makePreview('fetch');
    }
  });
})();
