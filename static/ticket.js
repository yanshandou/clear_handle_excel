(() => {
  const aliases = {
    ticket_id: ['チケット番号', 'タスク番号', '任务编号', 'Ticket ID'],
    title: ['題名', '件名', '任务标题', 'Subject'],
    parent: ['親チケット', '親タスク', '关联案件', 'Parent task'],
    status: ['ステータス', '状态', 'Status'],
    priority: ['優先度', '优先级', 'Priority'],
    assignee: ['担当者', '负责人', 'Assignee'],
    start_date: ['開始日', '开始日期', 'Start date'],
    due_date: ['期日', '截止日期', 'Due date'],
    progress: ['進捗率', '进度', '% Done', 'Progress'],
    version: ['対象バージョン', '目标版本', 'Target version'],
    screen_id: ['画面/帳票ID', '画面/报表 ID', 'Screen ID'],
    estimated_hours: ['予定工数', '预计工时', 'Estimated time'],
    actual_hours: ['実績工数', '实际工时', 'Spent time'],
    activity: ['创建 / 更新记录'],
  };
  const labelKey = label => label.toLowerCase().replace(/\s*[／/]\s*/g, '/');
  const fieldByLabel = new Map(Object.entries(aliases).flatMap(([key, labels]) => labels.map(label => [labelKey(label), key])));
  const fieldNames = [...fieldByLabel.keys()].map(label => label.replaceAll('/', '[ \\t　]*[／/][ \\t　]*')).join('|');
  const fieldPattern = new RegExp(`(?:^|[ \\t　])(${fieldNames})[ \\t　]*[:：]`, 'gi');
  const taskHeading = /^(?:(?:task|ticket|タスク|チケット|任务)\s*[:：]?\s*[#＃]?\s*|[#＃]\s*)([0-9０-９]+)(?:\s*[:：–—-]\s*|\s+|$)(.*)$/i;
  const parentHeading = /^(?:[^#＃]*?[-–—]\s*)?(?:案件|親チケット|親タスク|关联案件)\s*[#＃]\s*[0-9０-９]+/;
  const activityLine = /が.*(?:追加|作成)[.。．\s]|(?:前に|日時に)更新[.。．\s]*$/;
  const clean = text => text.replace(/\t+/g, ' ').trim();

  function parseTicket(source) {
    const ticket = {};
    let continuation = '';
    let foundDetails = false;
    for (const line of source.split(/\r?\n/).map(line => line.trim()).filter(Boolean)) {
      const heading = line.match(taskHeading);
      if (heading) {
        ticket.ticket_id ??= heading[1].normalize('NFKC');
        if (heading[2]) ticket.title ??= clean(heading[2]);
        continuation = '';
        continue;
      }
      if (parentHeading.test(line)) {
        ticket.parent ??= clean(line);
        continuation = '';
        continue;
      }
      const columns = [...line.matchAll(fieldPattern)];
      if (columns[0]?.index === 0) {
        foundDetails = true;
        continuation = '';
        for (const [index, column] of columns.entries()) {
          const key = fieldByLabel.get(labelKey(column[1]));
          const value = clean(line.slice(column.index + column[0].length, columns[index + 1]?.index));
          if (value) ticket[key] = value;
          else if (index === columns.length - 1) continuation = key;
        }
        continue;
      }
      if (activityLine.test(line)) {
        ticket.activity ??= clean(line);
      } else if (continuation && (continuation !== 'progress' || /^[0-9０-９]+\s*[%％]$/.test(line))) {
        ticket[continuation] = clean(line);
      } else if (!foundDetails && !/^[0-9０-９]+\s*[%％]$/.test(line)) {
        ticket.title ??= clean(line);
      }
      continuation = '';
    }
    return ticket;
  }

  const SAMPLE = `task #55345
機能追加 - 案件 #55022: 介舟ファミリーのタブレット対応／計画書の電子サイン対応（正式版）

TBSNG0000_手書きサイン
（開発）金 星（社） が [2026-11-11 15:00] 9日前に追加. [2026-11-20 19:41] 約4時間前に更新.

ステータス:\t進行中\t開始日:\t2026-11-20
優先度:\t通常\t期日:\t2026-12-03
担当者:\t（開発）白 家綾（現）\t進捗率:\t
0%

対象バージョン:\t介舟ファミリークラウド総合 - V440\t\t
画面／帳票ID:\tTBSNG0000\t実績工数:\t1.00
予定工数:\t9.00`;

  const CHIPS = [
    ['ticket_id', '编号'], ['title', '标题'], ['parent', '案件'], ['status', '状态'],
    ['priority', '优先级'], ['assignee', '负责人'], ['start_date', '开始'], ['due_date', '截止'],
    ['progress', '进度'], ['version', '版本'], ['screen_id', '画面ID'], ['estimated_hours', '预计'],
    ['actual_hours', '实际'], ['activity', '记录'],
  ];

  const form = document.getElementById('ticketForm');
  const source = document.getElementById('ticketSource');
  const bed = document.getElementById('scannerBed');
  const chipRow = document.getElementById('ticketChips');
  const boardCount = document.getElementById('boardCount');
  const boardTicks = document.getElementById('boardTicks');
  const headerCount = document.getElementById('headerCount');
  const recognition = document.getElementById('ticketRecognition');
  const pass = document.getElementById('boardingPass');
  const stamp = document.getElementById('passStamp');
  const stampText = document.getElementById('passStampText');
  const stubFill = document.getElementById('stubFill');
  const barcode = document.getElementById('stubBarcode');
  const create = document.getElementById('ticketCreate');
  const reset = document.getElementById('ticketReset');
  const success = document.getElementById('ticketSuccess');
  const successFolder = document.getElementById('successFolder');
  const error = document.getElementById('ticketError');
  const fields = [...form.querySelectorAll('input[name], textarea[name]')].filter(el => el !== source);
  const byName = Object.fromEntries(fields.map(el => [el.name, el]));
  const manual = new Set();
  let lastFolder = '';

  const wrapOf = el => el.closest('.pass-field, .pass-title, .stub-progress');
  const tagOf = el => wrapOf(el)?.querySelector('.auto-tag');

  boardTicks.innerHTML = CHIPS.map(() => '<i></i>').join('');
  const ticks = [...boardTicks.children];
  chipRow.innerHTML = CHIPS.map(([key, label]) =>
    `<button type="button" class="chip" data-key="${key}" title="跳到票面「${label}」">${label}</button>`).join('');
  const chips = [...chipRow.children];

  function setTag(el, kind) {
    const tag = tagOf(el);
    if (!tag) return;
    tag.hidden = !kind;
    tag.className = 'auto-tag' + (kind ? ` is-${kind}` : '');
    tag.textContent = kind === 'auto' ? '自动' : kind === 'manual' ? '手动' : '';
  }

  function pulse(el) {
    const wrap = wrapOf(el);
    if (!wrap) return;
    wrap.classList.remove('just-filled');
    void wrap.offsetWidth;
    wrap.classList.add('just-filled');
    wrap.addEventListener('animationend', () => wrap.classList.remove('just-filled'), { once: true });
  }

  function drawBarcode(seed) {
    const text = seed || 'HARBOR AIR';
    const bars = [];
    for (let i = 0; i < 26; i++) {
      const code = text.charCodeAt(i % text.length);
      bars.push(1 + ((code >> (i % 5)) & 3));
    }
    barcode.innerHTML = bars.map(w => `<i style="flex:0 0 ${w}px"></i>`).join('');
  }

  function drawLive() {
    const status = byName.status.value.trim();
    stamp.className = 'pass-stamp ' + (!status ? 'st-idle'
      : /完了|完成|終了|closed|done/i.test(status) ? 'st-done'
      : /新規|新建|new/i.test(status) ? 'st-new' : 'st-doing');
    stampText.textContent = status ? (status.length > 10 ? status.slice(0, 9) + '…' : status) : '待填写';
    const digits = byName.progress.value.match(/[0-9０-９]+/);
    const pct = digits ? Math.min(100, Math.max(0, parseInt(digits[0].normalize('NFKC'), 10))) : 0;
    stubFill.style.width = pct + '%';
    drawBarcode(byName.ticket_id.value.trim() || byName.title.value.trim());
  }

  function recognize() {
    const ticket = parseTicket(source.value);
    for (const el of fields) {
      if (manual.has(el.name)) continue;
      const value = ticket[el.name] ?? '';
      if (el.value !== value) {
        el.value = value;
        if (value) pulse(el);
      }
      setTag(el, value ? 'auto' : '');
    }
    const keys = new Set(Object.keys(ticket));
    chips.forEach((chip, i) => {
      chip.classList.toggle('is-on', keys.has(chip.dataset.key));
      ticks[i].classList.toggle('on', keys.has(chip.dataset.key));
    });
    const count = keys.size;
    boardCount.textContent = `${count} / ${CHIPS.length}`;
    headerCount.textContent = `识别 ${count} / ${CHIPS.length}`;
    recognition.textContent = count ? `识别到 ${count} 项，已填进票面，请核对后出票。`
      : source.value.trim() ? '没有识别到任务信息，可以直接手动填写票面。' : '';
    drawLive();
  }

  source.addEventListener('input', () => {
    bed.classList.remove('is-scanning');
    void bed.offsetWidth;
    bed.classList.add('is-scanning');
    recognize();
  });
  source.addEventListener('keydown', event => {
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') form.requestSubmit();
  });

  form.addEventListener('input', event => {
    const el = event.target;
    if (!el.name || el === source) return;
    manual.add(el.name);
    setTag(el, el.value.trim() ? 'manual' : '');
    drawLive();
  });

  chipRow.addEventListener('click', event => {
    const chip = event.target.closest('.chip');
    if (!chip) return;
    const el = byName[chip.dataset.key];
    el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    el.focus({ preventScroll: true });
  });

  document.getElementById('ticketSample').addEventListener('click', () => {
    source.value = SAMPLE;
    source.dispatchEvent(new Event('input', { bubbles: true }));
    // 重复载入同一份示例时字段值不变，也要给可见反馈，免得看起来像没响应
    fields.forEach(el => { if (el.value.trim()) pulse(el); });
    recognition.textContent = `已载入示例：识别 ${boardCount.textContent} 项，核对票面后即可出票。`;
    source.focus();
  });
  document.getElementById('ticketClearSource').addEventListener('click', () => {
    source.value = '';
    source.dispatchEvent(new Event('input', { bubbles: true }));
    source.focus();
  });

  form.addEventListener('reset', () => {
    manual.clear();
    requestAnimationFrame(() => {
      for (const el of fields) setTag(el, '');
      chips.forEach(chip => chip.classList.remove('is-on'));
      ticks.forEach(tick => tick.classList.remove('on'));
      boardCount.textContent = `0 / ${CHIPS.length}`;
      headerCount.textContent = `识别 0 / ${CHIPS.length}`;
      recognition.textContent = '';
      error.textContent = '';
      success.hidden = true;
      pass.classList.remove('is-created');
      drawLive();
      source.focus();
    });
  });

  document.getElementById('againBtn').addEventListener('click', () => form.reset());

  document.getElementById('openFolderBtn').addEventListener('click', async () => {
    if (!lastFolder) return;
    error.textContent = '';
    try {
      const response = await fetch('/api/ticket/open', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ folder: lastFolder }),
      });
      const data = await response.json();
      if (!response.ok || !data.ok) throw new Error(data.error || '打开文件夹失败。');
    } catch (err) {
      error.textContent = err.message || '打开文件夹失败。';
    }
  });

  form.addEventListener('submit', async event => {
    event.preventDefault();
    if (create.disabled) return;
    const values = Object.fromEntries(new FormData(form));
    for (const key of Object.keys(values)) values[key] = values[key].replace(/[\t\r\n]+/g, ' ').trim();
    create.disabled = reset.disabled = source.disabled = true;
    for (const el of fields) el.disabled = true;
    error.textContent = '';
    success.hidden = true;
    pass.classList.remove('is-created');
    try {
      const response = await fetch('/api/ticket/create', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(values),
      });
      const ticket = await response.json();
      if (!response.ok || !ticket.ok) throw new Error(ticket.error || '创建失败，请重试。');
      lastFolder = ticket.folder;
      successFolder.textContent = ticket.folder;
      success.hidden = false;
      void pass.offsetWidth;
      pass.classList.add('is-created');
      success.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    } catch (err) {
      error.textContent = err.message || '创建失败，请重试。';
    } finally {
      create.disabled = reset.disabled = source.disabled = false;
      for (const el of fields) el.disabled = false;
    }
  });

  drawLive();
})();