// 所有批量操作共用的实时进度框（用户 2026-11-09 要：所有耗时 / 批量操作都要实时进度提示）。
// 样子和用法照「调整页面」统一修改的进度框（#204）：fixed 浮在右下角，出现、收起、内容变化都不挪页面；
// 正在处理哪份、做到哪一步、进度条 + 预计几点完成（eta.js）；每份做完把结果（✓ 做好 / ! 没做好和原因）
// 加在列表最上面，文件名点一下在 Excel 打开；做完进度条收起，标题写 ✓ / !，点「知道了」关掉。
//
// 页面每次轮询到后台状态时调用：
//   JobPanel.run(id, { title, count, now, step, fraction, remaining, etaKey, etaLabel, waiting, log, empty, folder })  —— 正在跑
//   JobPanel.finish(id, { title, now, step, log, warn })  —— 跑完了（只有这一页看到它在跑、还没点「知道了」才显示）
// id 是这一类任务的名字（如 'batch'、'repair:links'）；log 每条 { ok, name, reason, path, sheet, folder }。
// 页面底部有悬浮条时，在 body 上设 --job-panel-bottom 让框浮在悬浮条上面。
// 跑着的时候有「⏭ 跳过这一份」（用户 2026-11-11 要：耗时太长的文件可以跳过）：结束正在做这份的后台 Excel，
// 任务记下「手动跳过」接着做下一份；旁边可以设「每份超过几分钟自动跳过」（0 = 不自动，存在后台）。
// 「⏹ 结束整个任务」：正在做的这份马上停，后面的都不做（记成「没做」），主页一类任务停成「可继续上次任务」。
// 页面传 skip: false 可以不显示（这个任务没有用 Excel 一份份处理的步骤）。
(function () {
  const OPEN_KEY = 'clear_handle_excel.jobPanel.open';   // sessionStorage：正在显示哪个任务（刷新页面还接着显示）
  const FOLD_KEY = 'clear_handle_excel.jobPanel.folded'; // localStorage：收起成一行的偏好
  const SHOW_MAX = 200;
  const esc = v => String(v ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const store = {
    get(area, key) { try { return window[area].getItem(key) || ''; } catch (_) { return ''; } },
    set(area, key, value) { try { value ? window[area].setItem(key, value) : window[area].removeItem(key); } catch (_) {} },
  };

  let box = null, el = {}, current = '', state = '', logSig = '', folder = '';

  function ensureCss() {
    if (document.querySelector('link[data-job-panel]')) return;
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = '/static/job_panel.css';
    link.dataset.jobPanel = '1';
    document.head.appendChild(link);
  }

  function build() {
    if (box) return;
    ensureCss();
    box = document.createElement('section');
    box.className = 'job-panel';
    box.hidden = true;
    box.setAttribute('role', 'status');
    box.setAttribute('aria-live', 'polite');
    box.innerHTML = `
      <div class="job-panel-head">
        <span class="job-panel-icon" aria-hidden="true">⏳</span>
        <h2 class="job-panel-title">正在处理</h2>
        <span class="job-panel-count"></span>
        <button type="button" class="job-panel-btn" data-job-fold title="收起成一行，不挡着页面；再点展开">收起</button>
      </div>
      <div class="job-panel-body">
        <p class="job-panel-now"></p>
        <p class="job-panel-step"></p>
        <div class="progress-track"><div class="progress-fill" style="width: 0%;"></div></div>
        <div class="job-panel-skip" hidden>
          <button type="button" class="job-panel-btn" data-job-skip title="这一份太慢不想等：结束后台 Excel，这份这次不做，接着做下一份">⏭ 跳过这一份</button>
          <button type="button" class="job-panel-btn job-panel-stop" data-job-stop title="整个任务不做了：正在做的这份马上停，后面的都不做，做好的保留">⏹ 结束整个任务</button>
          <label class="job-panel-limit" title="批量处理时一份文件做了这么久还没做完，就自动跳过（扫描查看切字有自己的时间上限，不看这个）">每份超过
            <input type="number" min="0" max="1440" step="1" inputmode="numeric" data-job-limit> 分钟自动跳过
            <span class="job-panel-limit-note">（0 = 不自动）</span></label>
        </div>
        <p class="job-panel-tally"><span class="is-ok">✓ 0</span><span class="is-bad is-zero">! 0</span><span class="job-panel-hint">点文件名在 Excel 打开</span></p>
        <ul class="job-panel-log"></ul>
        <div class="job-panel-foot" hidden><button type="button" class="btn btn-outline" data-job-close>知道了</button><a class="btn btn-primary" data-job-home href="/">🏠 返回首页</a></div>
      </div>`;
    document.body.appendChild(box);
    for (const name of ['icon', 'title', 'count', 'now', 'step', 'tally', 'log', 'foot', 'hint']) {
      el[name] = box.querySelector(`.job-panel-${name}`);
    }
    el.track = box.querySelector('.progress-track');
    el.fill = box.querySelector('.progress-fill');
    el.ok = box.querySelector('.job-panel-tally .is-ok');
    el.bad = box.querySelector('.job-panel-tally .is-bad');
    el.fold = box.querySelector('[data-job-fold]');
    el.skipRow = box.querySelector('.job-panel-skip');
    el.skip = box.querySelector('[data-job-skip]');
    el.limit = box.querySelector('[data-job-limit]');
    el.skip.addEventListener('click', skipCurrent);
    el.stop = box.querySelector('[data-job-stop]');
    el.stop.addEventListener('click', stopTask);
    el.limit.addEventListener('change', saveLimit);
    el.fold.addEventListener('click', () => setFold(!box.classList.contains('is-folded'), true));
    box.querySelector('[data-job-close]').addEventListener('click', close);
    // 做完以后「返回首页」：首页上这一页对应的卡片会闪几下（static/tools.js）；主页自己的任务不用回
    const home = box.querySelector('[data-job-home]');
    home.hidden = location.pathname === '/' || location.pathname === '/home.html';
    home.href = window.Tools?.homeUrl() || '/';
    setFold(store.get('localStorage', FOLD_KEY) === '1', false);
  }

  // 页面底部有悬浮条（精简 Sheet 的「确认执行」条、工作台的多选条、调整页面的底部条等）时浮在它上面，不挡按钮。
  // 只看贴着窗口底部、和框左右重叠的 fixed / sticky 元素；没变就不动
  let lift = -1;
  function placeAboveBars() {
    if (!box || box.hidden) return;
    const panel = box.getBoundingClientRect();
    const h = window.innerHeight;
    let need = 24;
    const seen = new Set();
    const points = [];
    for (const x of [panel.left + 12, (panel.left + panel.right) / 2, panel.right - 12]) {
      for (const y of [h - 6, h - 30, h - 60]) points.push([x, y]);
    }
    for (const [x, y] of points) {
      for (const hit of document.elementsFromPoint(x, y)) {
        if (box.contains(hit)) continue;
        for (let node = hit; node && node !== document.body; node = node.parentElement) {
          if (seen.has(node)) break;
          seen.add(node);
          const pos = getComputedStyle(node).position;
          if (pos !== 'fixed' && pos !== 'sticky') continue;
          if (node.classList.contains('excel-open-toast')) break;
          const r = node.getBoundingClientRect();
          if (r.height && r.top > h * 0.5) need = Math.max(need, Math.ceil(h - r.top) + 12);
          break;
        }
      }
    }
    if (need !== lift) {
      lift = need;
      box.style.setProperty('--job-panel-bottom', `${need}px`);
    }
  }
  let liftTimer = 0;
  const queueLift = () => { clearTimeout(liftTimer); liftTimer = setTimeout(placeAboveBars, 120); };
  window.addEventListener('scroll', queueLift, { passive: true });
  window.addEventListener('resize', queueLift);

  function setFold(folded, remember) {
    box.classList.toggle('is-folded', folded);
    el.fold.textContent = folded ? '展开' : '收起';
    if (remember) store.set('localStorage', FOLD_KEY, folded ? '1' : '');
  }

  function close() {
    if (!box) return;
    box.hidden = true;
    window.ProgressEta?.stop?.(`job_panel:${current}`);
    store.set('sessionStorage', OPEN_KEY, '');
    console.info('📋 进度框关掉', { 任务: current });
    current = ''; state = '';
  }

  // ---- 跳过这一份 / 每份超过几分钟自动跳过 ----
  const say = (msg, tone = 'info') => (window.appAlert ? window.appAlert(msg, { tone }) : Promise.resolve(alert(msg)));
  async function api(url, body) {
    const res = await fetch(url, body ? { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) } : {});
    const data = await res.json().catch(() => ({}));
    if (!res.ok || data.ok === false && data.error) throw new Error(data.error || `HTTP ${res.status}`);
    return data;
  }
  let limitLoaded = false;
  async function loadLimit() {
    if (limitLoaded) return;
    limitLoaded = true;
    try {
      const data = await api('/api/job/skip');
      if (document.activeElement !== el.limit) el.limit.value = data.settings?.file_limit_minutes ?? 0;
    } catch (err) {
      limitLoaded = false;
      console.warn('📋 读不到每份自动跳过的设置', err);
    }
  }
  async function saveLimit() {
    const minutes = Number(el.limit.value);
    if (!Number.isInteger(minutes) || minutes < 0 || minutes > 1440) {
      await say('每份最多几分钟要填 0–1440 之间的整数（0 = 不自动跳过）', 'warn');
      limitLoaded = false; loadLimit();
      return;
    }
    try {
      const data = await api('/api/job/skip/settings', { file_limit_minutes: minutes });
      el.limit.value = data.settings.file_limit_minutes;
      console.info('📋 每份超过几分钟自动跳过', { 分钟: minutes || '不自动' });
    } catch (err) {
      say(`没能保存：${err.message}`, 'error');
    }
  }
  async function skipCurrent() {
    if (el.skip.disabled) return;
    el.skip.disabled = true;
    try {
      const now = await api('/api/job/skip');
      if (!now.skippable) {
        console.info('📋 点了跳过这一份，但现在不能跳过', now);
        await say(now.running
          ? '现在这一步不能跳过（正在保存文件或做收尾，跳过会弄坏文件），等这一步做完再点。'
          : '这一步没有在用 Excel 一份份处理文件，不能跳过，等它做完就好。');
        return;
      }
      const mins = Math.floor((now.seconds || 0) / 60);
      const ok = await (window.appConfirm || (m => Promise.resolve(confirm(m))))(
        `跳过「${now.label}」？\n\n这份已经做了 ${mins ? `${mins} 分钟` : `${now.seconds || 0} 秒`}。`
        + '会结束后台 Excel，这份这次不做（没做完的不会保存进文件），结果里写「手动跳过」，之后可以单独再做。\n后面的文件接着做。',
        { title: '跳过这一份', okText: '跳过', cancelText: '接着等', tone: 'warn' });
      if (!ok) return;
      const res = await api('/api/job/skip', { label: now.label });
      console.info('📋 跳过这一份', res);
      if (!res.ok) await say(res.reason || '没能跳过', 'warn');
      else el.skip.textContent = '⏭ 正在跳过…';
    } catch (err) {
      say(`没能跳过：${err.message}`, 'error');
    } finally {
      setTimeout(() => { el.skip.disabled = stopping; el.skip.textContent = '⏭ 跳过这一份'; }, 2500);
    }
  }

  let stopping = false;
  async function stopTask() {
    if (stopping) return;
    const ok = await (window.appConfirm || (m => Promise.resolve(confirm(m))))(
      `结束「${el.title.textContent}」？\n\n${el.count.textContent ? `现在做到${el.count.textContent}。` : ''}`
      + '正在做的这份马上停（会结束后台 Excel，没做完的不保存），后面的都不做；已经做好的保留。\n'
      + '没做的在结果里写「没做」。主页上的任务（导出 PDF、精简等）之后可以点「继续上次任务」接着做。',
      { title: '结束整个任务', okText: '结束', cancelText: '接着做', tone: 'danger' });
    if (!ok) return;
    try {
      const res = await api('/api/job/stop', {});
      console.info('📋 结束整个任务', res);
      stopping = true;
      el.stop.disabled = el.skip.disabled = true;
      el.stop.textContent = '⏹ 正在结束…';
    } catch (err) {
      say(`没能结束：${err.message}`, 'error');
    }
  }
  function resetStop() {
    if (!el.stop) return;
    stopping = false;
    el.stop.disabled = el.skip.disabled = false;
    el.stop.textContent = '⏹ 结束整个任务';
  }

  function line(item) {
    const label = esc(item.name || '');
    const target = item.path || item.file ? { file: item.path || item.file, folder: item.folder || folder, sheet: item.sheet || '' } : null;
    const what = target && window.excelLinkHtml
      ? window.excelLinkHtml(`📗 ${label}${item.sheet ? ` › ${esc(item.sheet)}` : ''}`, target)
      : `<span class="job-panel-name">${label}${item.sheet ? ` › ${esc(item.sheet)}` : ''}</span>`;
    const note = item.reason ? `<span class="job-panel-note">${esc(item.reason)}</span>` : '';
    return `<li class="${item.ok ? 'is-ok' : 'is-bad'}"><i aria-hidden="true">${item.ok ? '✓' : '!'}</i><div>${what}${note}</div></li>`;
  }

  // 最新的在最上面：列表只往上长，不用滚动
  function paintLog(log, empty, force) {
    log = Array.isArray(log) ? log : [];
    const sig = `${log.length}|${log.length ? JSON.stringify(log[log.length - 1]) : ''}|${empty || ''}|${window.excelLinkHtml?.revision || 0}`;
    if (!force && sig === logSig) return;
    logSig = sig;
    const ok = log.filter(i => i.ok).length;
    el.ok.textContent = `✓ ${ok}`;
    el.bad.textContent = `! ${log.length - ok}`;
    el.bad.classList.toggle('is-zero', log.length === ok);
    el.hint.hidden = !log.some(i => i.path || i.file);
    if (!log.length) {
      el.log.innerHTML = `<li class="job-panel-empty">${esc(empty || '做完第一份就会列出结果…')}</li>`;
      return;
    }
    const shown = log.slice(-SHOW_MAX).reverse();
    el.log.innerHTML = shown.map(line).join('')
      + (log.length > SHOW_MAX ? `<li class="job-panel-more">……更早的 ${log.length - SHOW_MAX} 条没列出</li>` : '');
  }

  function open(id) {
    build();
    if (current !== id || state !== 'running') {
      // 新的一次（或上一次已经结束）：清空重来
      logSig = '';
      resetStop();
      el.track.classList.remove('is-done');
      el.foot.hidden = true;
      box.classList.remove('is-waiting');
      console.info('📋 进度框打开', { 任务: id });
    }
    current = id;
    state = 'running';
    box.dataset.state = 'running';
    box.hidden = false;
    store.set('sessionStorage', OPEN_KEY, id);
    queueLift();
  }

  function run(id, s = {}) {
    open(id);
    folder = s.folder || '';
    el.icon.textContent = '⏳';
    el.title.textContent = s.title || '正在处理';
    el.count.textContent = s.count || '';
    el.now.textContent = s.now || '准备中…';
    el.step.textContent = s.step ? `▸ ${s.step}` : '';
    box.classList.toggle('is-waiting', !!s.waiting);
    el.skipRow.hidden = s.skip === false;
    if (s.skip !== false) loadLimit();
    const fraction = Math.max(0, Math.min(1, Number(s.fraction) || 0));
    el.fill.style.width = `${(fraction * 100).toFixed(1)}%`;
    // etaKey：和页面上同一条进度共用一份预估记录，两处显示的时间就一样；etaLabel 写明算的是什么
    window.ProgressEta?.update(el.fill, s.etaKey || `job_panel:${id}`, fraction, true, s.remaining, s.etaLabel);
    paintLog(s.log, s.empty, false);
  }

  function finish(id, s = {}) {
    // 只收尾这一页看到在跑的任务（刷新页面时还没点「知道了」的也算）；别的页面早就做完的不弹出来
    if (current !== id && store.get('sessionStorage', OPEN_KEY) !== id) return;
    if (current === id && state !== 'running' && !box?.hidden && !s.force) {
      paintLog(s.log, s.empty, false);
      return;
    }
    build();
    current = id;
    state = 'done';
    folder = s.folder || folder;
    const log = Array.isArray(s.log) ? s.log : [];
    const warn = s.warn ?? log.some(i => !i.ok);
    box.dataset.state = warn ? 'warn' : 'done';
    box.classList.remove('is-waiting');
    box.hidden = false;
    window.ProgressEta?.update(el.fill, `job_panel:${id}`, 1, false);
    el.track.classList.add('is-done'); // 进度条收起，只留结果
    el.skipRow.hidden = true;
    resetStop();
    el.icon.textContent = warn ? '!' : '✓';
    el.title.textContent = s.title || (warn ? '做完了，有要看一下的' : '做完了');
    el.count.textContent = '';
    el.now.textContent = s.now || (log.length ? `${log.filter(i => i.ok).length} 份做好了${warn ? `，${log.filter(i => !i.ok).length} 份要看一下` : ''}` : '');
    el.step.textContent = s.step || '';
    el.foot.hidden = false;
    paintLog(log, s.empty || '没有要处理的', true);
    queueLift();
    console.info('📋 进度框做完', { 任务: id, 好: log.filter(i => i.ok).length, 要看: log.filter(i => !i.ok).length });
  }

  // 第几份 / 共几份的说法统一
  const count = (done, total, unit = '份') => total ? `第 ${Math.min((done || 0) + 1, total)} / ${total} ${unit}` : '准备中…';

  // 主页、精简 Sheet、执行记录这些页面轮询的 /api/overview：精简 Sheet、导出 PDF、视图标准化、扫描等
  // 主页一类的任务（后台记执行记录的），feed 是这一步每份做完的结果
  function overview(ov) {
    const job = ov && ov.job;
    if (!job || !job.kind || (job.kind.startsWith('batch_') && job.kind !== 'batch_scan')) return;
    const p = ov.progress || {};
    const id = `task:${job.kind}`;
    const name = String(p.task_name || '处理').replace(/（.*?）/g, '');
    const log = Array.isArray(ov.feed) ? ov.feed : [];
    const folder = job.folder || ov.folder || '';
    const total = Number(p.total_files) || 0;
    const finished = (Number(p.completed_files) || 0) + (Number(p.error_files) || 0);
    if (ov.busy && job.status === 'running') {
      const stepText = String(p.current_step || '');
      const m = stepText.match(/^\[(.+?)\]\s*(.*)$/);
      run(id, {
        title: `正在${name}`, count: count(finished, total, '份 Excel'),
        now: p.current_file || (m && m[1]) ? `📗 ${p.current_file || m[1]}` : '准备中…',
        step: m ? m[2] : stepText, waiting: /等你选择|出错了/.test(stepText),
        fraction: total ? finished / total : 0, log, folder,
      });
      return;
    }
    if (ov.busy) return;
    const bad = log.filter(i => !i.ok).length;
    const paused = job.status === 'interrupted';
    finish(id, {
      title: paused ? `${name}暂停了，可以继续上次任务` : `${name}做完了${bad ? `，${bad} 份出错` : ''}`,
      warn: paused || bad > 0, log, folder,
      step: bad ? '出错的可以修好后点「继续上次任务」再做。' : '',
    });
  }

  window.JobPanel = { run, finish, close, count, overview, isOpen: id => !!box && !box.hidden && current === id };
})();
