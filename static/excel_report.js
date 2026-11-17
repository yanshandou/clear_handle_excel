// Excel 对比报告页（用户 2026-11-17）：任选两份 Excel → POST /api/excel_report → 拿回单个 HTML 报告，
// 页面里预览、下载、新窗口打开。每边可以填路径，也可以选 / 拖文件（那样会把文件内容传给本机程序）。
(function () {
  const $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const KEY = { a: 'clear_handle_excel.reportA', b: 'clear_handle_excel.reportB' };
  const ETA_KEY = 'excel_report';
  const picked = { a: null, b: null };  // 选 / 拖进来的文件（File）；填了路径就不用它
  let report = null, running = false, timer = null;

  const read = key => { try { return localStorage.getItem(key) || ''; } catch (_) { return ''; } };
  const write = (key, value) => { try { localStorage.setItem(key, value); } catch (_) { /* 存不了就算了 */ } };
  const input = side => $(side === 'a' ? 'xpPathA' : 'xpPathB');
  const chosen = side => $(side === 'a' ? 'xpChosenA' : 'xpChosenB');
  const cleanPath = text => text.trim().replace(/^"(.*)"$/, '$1');

  input('a').value = read(KEY.a);
  input('b').value = read(KEY.b);

  function showPicked(side) {
    const file = picked[side];
    chosen(side).textContent = file ? `已选：${file.name}（${(file.size / 1024 / 1024).toFixed(1)} MB）` : '';
  }
  function setFile(side, file) {
    if (!file) return;
    if (!/\.(xlsx|xlsm|xltx|xltm)$/i.test(file.name)) {
      window.appAlert(`「${file.name}」没法对比：只支持 .xlsx / .xlsm / .xltx / .xltm。旧的 .xls 请先在 Excel 里另存为 .xlsx。`, { title: '这种文件不行' });
      return;
    }
    picked[side] = file;
    input(side).value = '';
    showPicked(side);
    console.info('📊 对比报告：选了文件', side, file.name, file.size);
  }

  document.querySelectorAll('[data-pick]').forEach(button => {
    button.addEventListener('click', () => $(button.dataset.pick === 'a' ? 'xpFileA' : 'xpFileB').click());
  });
  ['a', 'b'].forEach(side => {
    const fileInput = $(side === 'a' ? 'xpFileA' : 'xpFileB');
    fileInput.addEventListener('change', () => { setFile(side, fileInput.files[0]); fileInput.value = ''; });
    input(side).addEventListener('input', () => {
      if (input(side).value.trim() && picked[side]) { picked[side] = null; showPicked(side); }  // 改填路径就不用选的文件
      write(KEY[side], cleanPath(input(side).value));
    });
    const box = document.querySelector(`.xp-side[data-side="${side}"]`);
    box.addEventListener('dragover', event => { event.preventDefault(); box.classList.add('is-drop'); });
    box.addEventListener('dragleave', event => { if (!box.contains(event.relatedTarget)) box.classList.remove('is-drop'); });
    box.addEventListener('drop', event => {
      event.preventDefault();
      box.classList.remove('is-drop');
      const files = [...(event.dataTransfer.files || [])];
      if (files.length >= 2) { setFile('a', files[0]); setFile('b', files[1]); return; }  // 一次拖两份：第一份当 A
      setFile(side, files[0]);
    });
  });
  // 拖到框外面别让浏览器把文件打开
  window.addEventListener('dragover', event => event.preventDefault());
  window.addEventListener('drop', event => event.preventDefault());

  $('xpSwap').addEventListener('click', () => {
    [picked.a, picked.b] = [picked.b, picked.a];
    const a = input('a').value;
    input('a').value = input('b').value; input('b').value = a;
    write(KEY.a, cleanPath(input('a').value)); write(KEY.b, cleanPath(input('b').value));
    showPicked('a'); showPicked('b');
  });

  function toBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result).split(',', 2)[1] || '');
      reader.onerror = () => reject(new Error(`读不了「${file.name}」：${reader.error}`));
      reader.readAsDataURL(file);
    });
  }
  async function sidePayload(side) {
    const path = cleanPath(input(side).value);
    if (path) return { path };
    if (picked[side]) return { name: picked[side].name, data: await toBase64(picked[side]) };
    return null;
  }

  function busy(on, stage = '', fraction = 0) {
    $('xpBusy').hidden = !on;
    $('xpRun').disabled = on;
    if (!on) { window.ProgressEta.stop(ETA_KEY); return; }
    const pct = Math.round(Math.max(0, Math.min(1, fraction)) * 100);
    $('xpPercent').textContent = `${pct}%`;
    $('xpBar').style.width = `${pct}%`;
    $('xpStage').textContent = stage;
    $('xpEta').textContent = window.ProgressEta.text(ETA_KEY, fraction);
  }
  async function poll() {
    try {
      const response = await fetch('/api/excel_report/status');
      const state = await response.json();
      if (!running || !state.running) return;
      // 读两份文件算前 20%，逐张表比算后面 80%
      const fraction = state.total ? 0.2 + 0.8 * state.done / state.total : (state.stage.includes('读') ? 0.05 : 0);
      busy(true, `${state.stage}${state.total ? `（${state.done} / ${state.total} 张表）` : ''} · 已用 ${Math.round(state.elapsed)} 秒`, fraction);
    } catch (_) { /* 问进度失败不要紧，等结果 */ }
  }

  function verdictHtml(data) {
    const s = data.sheets;
    if (data.identical) return '<div class="xp-verdict is-same">✓ 两份文件在检查的项目里没有不同</div>';
    const parts = [s.changed ? `${s.changed} 张表有改动` : '', s.only_old ? `${s.only_old} 张表只有 A 有` : '', s.only_new ? `${s.only_new} 张表只有 B 有` : '',
      data.order_changed ? '表的顺序变了' : '', data.book ? `工作簿设置 ${data.book} 处不同` : ''].filter(Boolean);
    const kinds = Object.entries(data.totals).map(([kind, count]) => `${esc(kind)} ${count}`).join(' · ');
    return `<div class="xp-verdict is-diff">有不同：${esc(parts.join('，'))}${kinds ? `<br><small>${kinds}</small>` : ''}</div>`;
  }

  async function run() {
    if (running) return;
    $('xpError').textContent = '';
    let a, b;
    try {
      busy(true, '正在读选的文件…', 0);
      [a, b] = await Promise.all([sidePayload('a'), sidePayload('b')]);
    } catch (err) {
      busy(false);
      $('xpError').textContent = String(err.message || err);
      return;
    }
    if (!a || !b) {
      busy(false);
      $('xpError').textContent = `请先选文件 ${!a ? 'A' : 'B'}：填完整路径，或点「选择文件」，或把文件拖到框里。`;
      return;
    }
    running = true;
    window.ProgressEta.stop(ETA_KEY);
    busy(true, '正在读两份文件…', 0);
    timer = setInterval(poll, 800);
    console.info('📊 对比报告：开始', { a: a.path || a.name, b: b.path || b.name });
    try {
      const response = await fetch('/api/excel_report', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ a, b }) });
      const data = await response.json().catch(() => ({ ok: false, error: `程序没有回应（${response.status}）` }));
      if (!data.ok) throw new Error(data.error || '生成失败');
      report = data;
      $('xpVerdict').innerHTML = verdictHtml(data);
      $('xpSaved').textContent = data.saved ? `报告也存了一份：${data.saved}` : '';
      $('xpFrame').srcdoc = data.html;
      $('xpResultPanel').hidden = false;
      console.info('📊 对比报告：完成', { identical: data.identical, totals: data.totals, saved: data.saved, kb: Math.round(data.html.length / 1024) });
      if (window.gameFx && window.gameFx.celebrate) window.gameFx.celebrate();
    } catch (err) {
      console.warn('📊 对比报告：失败', err);
      $('xpError').textContent = String(err.message || err);
    } finally {
      running = false;
      clearInterval(timer);
      busy(false);
    }
  }
  $('xpRun').addEventListener('click', run);

  function blobUrl() {
    return URL.createObjectURL(new Blob([report.html], { type: 'text/html;charset=utf-8' }));
  }
  $('xpDownload').addEventListener('click', () => {
    if (!report) return;
    // 存好的那份让程序直接发下载（文件名不会乱）；没存成才用浏览器里的副本
    const url = report.saved ? `/api/excel_report/download?name=${encodeURIComponent(report.filename)}` : blobUrl();
    const link = document.createElement('a');
    link.href = url; link.download = report.filename;
    document.body.appendChild(link); link.click(); link.remove();
    if (!report.saved) setTimeout(() => URL.revokeObjectURL(url), 60000);
    console.info('📊 对比报告：下载', report.filename);
  });
  $('xpOpen').addEventListener('click', () => {
    if (!report) return;
    const url = blobUrl();
    window.open(url, '_blank');
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  });
})();
