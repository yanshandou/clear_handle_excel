(() => {
  const aliases = {
    ticket_id: ['チケット番号', 'タスク番号', '任务编号', 'Ticket ID'],
    title: ['題名', '件名', '任务标题', 'Subject'],
    parent: ['親チケット', '親タスク', '关联案件', 'Parent task'],
    project: ['所属项目'], parent_project: ['上级项目名称'],
    parent_category: ['关联案件分类'], parent_title: ['关联案件标题'],
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
        const parent = line.match(/^(.*?)\s*[-–—]\s*(?:案件|親チケット|親タスク|关联案件)\s*[#＃]\s*([0-9０-９]+)\s*[:：]\s*(.*)$/);
        if (parent) { ticket.parent_category = clean(parent[1]); ticket.parent = parent[2].normalize('NFKC'); ticket.parent_title = clean(parent[3]); }
        else ticket.parent ??= clean(line);
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
        // 创建及更新文字不是票面字段。
      } else if (continuation && (continuation !== 'progress' || /^[0-9０-９]+\s*[%％]$/.test(line))) {
        ticket[continuation] = clean(line);
      } else if (!foundDetails && !/^[0-9０-９]+\s*[%％]$/.test(line)) {
        ticket.title ??= clean(line);
      }
      continuation = '';
    }
    return ticket;
  }

  // Parse only the issue summary and selected project; never use history or edit options as current values.
  function parseTicketHtml(source) {
    // Remove active/resource-bearing content before parsing the detached document.
    const safe = source.replace(/<(script|style|iframe|object|embed|link|img|source|video|audio)\b[^>]*>[\s\S]*?<\/\1\s*>|<(?:script|style|iframe|object|embed|link|img|source|video|audio)\b[^>]*\/?\s*>/gi, '')
      .replace(/\s(?:src|srcset|poster|background)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '');
    const doc = new DOMParser().parseFromString(safe, 'text/html');
    const result = {};
    const put = (key, value) => { value = (value || '').replace(/\s+/g, ' ').trim(); if (value) result[key] = value; };
    const issue = doc.querySelector('.issue.details');
    const heading = doc.querySelector('#content > h2, h2');
    const number = heading?.textContent.match(/(?:task|ticket|チケット|タスク)\s*#\s*(\d+)/i);
    if (number) put('ticket_id', number[1]);
    put('title', issue?.querySelector('.subject h3')?.textContent);
    const parent = issue?.querySelector('.subject p');
    const parentLink = parent?.querySelector('a.parent');
    const parentNumber = parentLink?.textContent.match(/#\s*(\d+)/);
    if (parentNumber) put('parent', parentNumber[1]);
    if (parent && parentLink) {
      const full = parent.textContent.replace(/\s+/g, ' ').trim();
      const linkText = parentLink.textContent.replace(/\s+/g, ' ').trim();
      const at = full.indexOf(linkText);
      if (at >= 0) {
        put('parent_category', full.slice(0, at).replace(/\s*[-–—]\s*$/, ''));
        put('parent_title', full.slice(at + linkText.length).replace(/^\s*[:：]\s*/, ''));
      }
    }
    const attrs = issue?.querySelector('table.attributes');
    const classes = {status:'status', priority:'priority', 'assigned-to':'assignee', 'start-date':'start_date', 'due-date':'due_date', progress:'progress', 'fixed-version':'version', cf_13:'screen_id', cf_84:'estimated_hours', cf_85:'actual_hours'};
    for (const [cls, key] of Object.entries(classes)) {
      const cell = attrs?.querySelector(`td.${cls}`);
      put(key, key === 'progress' ? cell?.querySelector('.percent')?.textContent : cell?.textContent);
    }
    const select = doc.querySelector('#issue_project_id');
    const selected = select?.querySelector('option[selected]');
    if (selected) {
      const label = selected.textContent.replace(/\s+/g, ' ').trim();
      put('project', label.replace(/^»\s*/, ''));
      if (label.startsWith('»')) {
        let previous = selected.previousElementSibling;
        while (previous && (previous.tagName !== 'OPTION' || previous.textContent.trim().startsWith('»'))) previous = previous.previousElementSibling;
        if (previous) put('parent_project', previous.textContent);
      } else if (selected.parentElement?.tagName === 'OPTGROUP') put('parent_project', selected.parentElement.label);
    }
    return result;
  }

  const SAMPLE = `<div id="content"><h2>task #55345</h2>
<div class="issue details"><div class="subject"><p>機能追加 - <a class="parent">案件 #55022</a>: 介舟ファミリーのタブレット対応／計画書の電子サイン対応（正式版）</p><h3>TBSNG0000_手書きサイン</h3></div>
<table class="attributes"><tr><td class="status">進行中</td><td class="start-date">2026/11/20</td><td class="priority">通常</td><td class="due-date">2026/12/03</td><td class="assigned-to">（開発）白 家綾（現）</td><td class="progress"><p class="percent">0%</p></td><td class="fixed-version">介舟ファミリークラウド総合 - V440</td><td class="cf_13">TBSNG0000</td><td class="cf_84">9.00</td><td class="cf_85">1.00</td></tr></table></div>
<select id="issue_project_id"><option>介舟ファミリークラウド総合</option><option selected>» 開発（UI）</option></select></div>`;

  const CHIPS = [
    ['ticket_id', '编号'], ['title', '标题'], ['project', '项目'], ['parent_project', '上级项目'],
    ['parent', '案件编号'], ['parent_category', '案件分类'], ['parent_title', '案件标题'],
    ['status', '状态'], ['priority', '优先级'], ['assignee', '负责人'], ['start_date', '开始'], ['due_date', '截止'],
    ['progress', '进度'], ['version', '版本'], ['screen_id', '画面ID'], ['estimated_hours', '预计'],
    ['actual_hours', '实际'],
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
  const iconField = document.getElementById('ticketIcon');
  const iconGrid = document.getElementById('ticketIconGrid');
  const defaultIcon = iconGrid.querySelector('input');
  const iconRefresh = document.getElementById('ticketIconRefresh');
  const iconStatus = document.getElementById('ticketIconStatus');
  const iconApplied = document.getElementById('ticketIconApplied');
  const iconWarning = document.getElementById('ticketIconWarning');
  const success = document.getElementById('ticketSuccess');
  const successFolder = document.getElementById('successFolder');
  const openFolder = document.getElementById('openFolderBtn');
  const openStatus = document.getElementById('ticketOpenStatus');
  const openError = document.getElementById('ticketOpenError');
  const error = document.getElementById('ticketError');
  const fields = [...form.querySelectorAll('input[name], textarea[name]')].filter(el => el !== source && !iconField.contains(el));
  const byName = Object.fromEntries(fields.map(el => [el.name, el]));
  const manual = new Set();
  let lastFolder = '';

  const wrapOf = el => el.closest('.pass-field, .pass-title, .stub-progress');
  const tagOf = el => wrapOf(el)?.querySelector('.auto-tag');

  function makeIconOption(icon) {
    const option = document.createElement('label');
    option.className = 'ticket-icon-option';
    option.title = icon.category ? `${icon.name}（${icon.category}）` : icon.name;

    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'icon_id';
    radio.value = icon.id;

    const tile = document.createElement('span');
    tile.className = 'ticket-icon-tile';
    const thumb = document.createElement('img');
    thumb.className = 'ticket-icon-thumb';
    thumb.alt = '';
    thumb.loading = 'lazy';
    thumb.src = `/api/icons/file?${new URLSearchParams({ id: icon.id, kind: 'png', v: icon.version })}`;
    const name = document.createElement('span');
    name.className = 'ticket-icon-name';
    name.textContent = icon.name;
    tile.append(thumb, name);
    option.append(radio, tile);
    return option;
  }

  async function loadIcons() {
    iconRefresh.disabled = true;
    iconStatus.textContent = '正在加载图标…';
    try {
      const response = await fetch('/api/icons', { cache: 'no-store' });
      const data = await response.json();
      if (!response.ok || !data.ok) throw new Error(data.error);
      const selected = iconGrid.querySelector('input:checked').value;
      const options = data.icons.map(makeIconOption);
      iconGrid.replaceChildren(defaultIcon.closest('label'), ...options);
      const selectedRadio = [...iconGrid.querySelectorAll('input')].find(radio => radio.value === selected);
      (selectedRadio || defaultIcon).checked = true;
      const selectionRemoved = !selectedRadio;
      iconStatus.textContent = selectionRemoved ? '原先选择的图标已从图标库移除。'
        : options.length ? '' : '图标库为空，可先到图标库添加图标。';
    } catch {
      iconStatus.textContent = '图标库加载失败，可点击「刷新图标」重试。';
    } finally {
      iconRefresh.disabled = false;
    }
  }

  iconRefresh.addEventListener('click', loadIcons);

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
    const text = seed || 'HARBOR LINE';
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

  function assigneeShort(value) {
    const name = value.normalize('NFKC').replace(/\([^()]*\)/g, '').replace(/\s+/g, '');
    const names = { '吴振科': '吴', '吳振科': '吴', '白家绫': '白', '白家綾': '白', '张晓静': '静', '張曉靜': '静' };
    return names[name] || '';
  }

  function updateAssigneeShort() {
    if (manual.has('assignee_short')) return;
    const el = byName.assignee_short;
    const value = assigneeShort(byName.assignee.value);
    if (el.value !== value) { el.value = value; if (value) pulse(el); }
    setTag(el, value ? 'auto' : '');
  }

  function formatWholeHours(value) {
    return value.replace(/^(\d+)\.0+$/, '$1');
  }

  function recognize() {
    const ticket = /<\s*(?:html|div|table|h2)\b/i.test(source.value) ? parseTicketHtml(source.value) : parseTicket(source.value);
    for (const el of fields) {
      if (manual.has(el.name) || el.name === 'assignee_short') continue;
      const rawValue = ticket[el.name] ?? '';
      const value = ['estimated_hours', 'actual_hours'].includes(el.name) ? formatWholeHours(rawValue) : rawValue;
      if (el.value !== value) {
        el.value = value;
        if (value) pulse(el);
      }
      setTag(el, value ? 'auto' : '');
    }
    updateAssigneeShort();
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
    if (!Object.hasOwn(byName, el.name)) return;
    manual.add(el.name);
    setTag(el, el.value.trim() ? 'manual' : '');
    if (el.name === 'assignee') updateAssigneeShort();
    drawLive();
  });

  const hoursButtons = [...form.querySelectorAll('[data-hours-for]')];
  function adjustHours(input, step) {
    const raw = input.value.normalize('NFKC').trim();
    if (raw && !/^(?:\d+(?:\.\d*)?|\.\d+)$/.test(raw)) { input.focus(); input.select(); return; }
    const digits = Math.min(10, Math.max(2, (raw.split('.')[1] || '').length));
    const value = Math.max(0, Number(raw || 0) + step);
    if (!Number.isFinite(value)) { input.focus(); input.select(); return; }
    input.value = formatWholeHours(value.toFixed(digits));
    input.dispatchEvent(new Event('input', { bubbles: true }));
  }
  hoursButtons.forEach(button => button.addEventListener('click', () => {
    adjustHours(document.getElementById(button.dataset.hoursFor), Number(button.dataset.hoursStep));
  }));
  for (const id of ['fEst', 'fAct']) {
    const input = document.getElementById(id);
    input.addEventListener('blur', () => { input.value = formatWholeHours(input.value); });
    input.addEventListener('wheel', event => {
      if (document.activeElement !== input || input.disabled || !event.deltaY) return;
      event.preventDefault();
      adjustHours(input, event.deltaY < 0 ? 1 : -1);
    }, { passive: false });
  }

  const progressInput = byName.progress;
  const progressSteps = [0, 50, 70, 100];
  progressInput.addEventListener('wheel', event => {
    if (document.activeElement !== progressInput || progressInput.disabled || !event.deltaY) return;
    event.preventDefault();
    const match = progressInput.value.normalize('NFKC').trim().match(/^(\d{1,3})\s*%?$/);
    const current = match ? Number(match[1]) : (event.deltaY < 0 ? -1 : 101);
    const next = event.deltaY < 0
      ? progressSteps.find(value => value > current) ?? 100
      : [...progressSteps].reverse().find(value => value < current) ?? 0;
    progressInput.value = `${next}%`;
    progressInput.dispatchEvent(new Event('input', { bubbles: true }));
  }, { passive: false });

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
    closeDatePop();
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

  // ---------- 船票风日历 ----------
  const DATE_CAPTIONS = { fStart: '启航 · FROM', fDue: '到港 · TO' };
  const dateTriggers = [...form.querySelectorAll('.date-trigger')];
  const pad2 = n => String(n).padStart(2, '0');
  const isoOf = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  const chev = (dir, double) => {
    const l = double ? 'M8 3 3 8l5 5M13 3 8 8l5 5' : 'M10 3 5 8l5 5';
    const r = double ? 'M3 3l5 5-5 5M8 3l5 5-5 5' : 'M6 3l5 5-5 5';
    return `<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="${dir < 0 ? l : r}"/></svg>`;
  };
  const datePop = document.createElement('div');
  datePop.className = 'date-pop';
  datePop.setAttribute('role', 'dialog');
  datePop.setAttribute('aria-modal', 'true');
  datePop.setAttribute('aria-label', '选择日期');
  datePop.hidden = true;
  datePop.innerHTML = `
    <div class="date-pop-band">
      <button type="button" class="date-nav" data-step="-12" aria-label="上一年" title="上一年">${chev(-1, true)}</button>
      <button type="button" class="date-nav" data-step="-1" aria-label="上个月" title="上个月">${chev(-1)}</button>
      <div class="date-pop-title"><b data-role="month"></b><small data-role="caption"></small></div>
      <button type="button" class="date-nav" data-step="1" aria-label="下个月" title="下个月">${chev(1)}</button>
      <button type="button" class="date-nav" data-step="12" aria-label="下一年" title="下一年">${chev(1, true)}</button>
    </div>
    <div class="date-paper">
      <div class="date-week" aria-hidden="true"><span>一</span><span>二</span><span>三</span><span>四</span><span>五</span><span class="wk-sat">六</span><span class="wk-sun">日</span></div>
      <div class="date-grid" data-role="grid" role="group" aria-label="日期"></div>
    </div>
    <div class="date-pop-perf" aria-hidden="true"><i class="notch left"></i><i class="notch right"></i></div>
    <div class="date-pop-foot">
      <span class="date-readout is-empty" data-role="readout">未选择</span>
      <button type="button" class="date-quick is-today" data-role="today">今天</button>
      <button type="button" class="date-quick is-clear" data-role="clear">清除</button>
    </div>
    <i class="date-pop-arrow" aria-hidden="true"></i>`;
  document.body.appendChild(datePop);
  const popMonth = datePop.querySelector('[data-role="month"]');
  const popCaption = datePop.querySelector('[data-role="caption"]');
  const popGrid = datePop.querySelector('[data-role="grid"]');
  const popReadout = datePop.querySelector('[data-role="readout"]');
  const popArrow = datePop.querySelector('.date-pop-arrow');
  let dateInput = null, viewY = 0, viewM = 0, placeQueued = false;

  function parseDate(value) {
    const m = value.normalize('NFKC').trim().match(/^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?$/);
    if (!m) return null;
    const [y, mo, d] = [+m[1], +m[2], +m[3]];
    const dt = new Date(y, mo - 1, d);
    return dt.getFullYear() === y && dt.getMonth() === mo - 1 && dt.getDate() === d ? dt : null;
  }

  function renderDatePop() {
    popMonth.textContent = `${viewY}年${viewM + 1}月`;
    popCaption.textContent = DATE_CAPTIONS[dateInput.id];
    const selected = parseDate(dateInput.value);
    const selIso = selected ? isoOf(selected) : '';
    const todayIso = isoOf(new Date());
    const offset = (new Date(viewY, viewM, 1).getDay() + 6) % 7;
    let html = '';
    for (let i = 0; i < 42; i++) {
      const d = new Date(viewY, viewM, 1 - offset + i);
      const dIso = isoOf(d);
      const dow = (d.getDay() + 6) % 7;
      const cls = ['date-day'];
      if (d.getMonth() !== viewM) cls.push('is-out');
      if (dow === 5) cls.push('is-sat');
      if (dow === 6) cls.push('is-sun');
      if (dIso === todayIso) cls.push('is-today');
      if (dIso === selIso) cls.push('is-sel');
      html += `<button type="button" class="${cls.join(' ')}" data-date="${dIso}" tabindex="-1" aria-pressed="${dIso === selIso}" aria-label="${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日" style="animation-delay:${Math.floor(i / 7) * 26}ms">${d.getDate()}</button>`;
    }
    popGrid.innerHTML = html;
    (popGrid.querySelector('.is-sel') || popGrid.querySelector('.is-today:not(.is-out)') || popGrid.querySelector('.date-day:not(.is-out)')).tabIndex = 0;
    popReadout.textContent = selected ? `${selected.getFullYear()}/${pad2(selected.getMonth() + 1)}/${pad2(selected.getDate())}` : '未选择';
    popReadout.classList.toggle('is-empty', !selected);
  }

  function placeDatePop() {
    const anchor = dateInput.closest('.date-control').getBoundingClientRect();
    const w = datePop.offsetWidth, h = datePop.offsetHeight;
    const vw = document.documentElement.clientWidth, vh = document.documentElement.clientHeight;
    const left = Math.min(Math.max(8, anchor.left), Math.max(8, vw - w - 8));
    let top = anchor.bottom + 10;
    const above = top + h > vh - 8 && anchor.top - h - 10 > 8;
    if (above) top = anchor.top - h - 10;
    datePop.style.left = `${left}px`;
    datePop.style.top = `${Math.min(Math.max(8, top), Math.max(8, vh - h - 8))}px`;
    datePop.classList.toggle('is-above', above);
    popArrow.style.left = `${Math.min(Math.max(anchor.left + anchor.width / 2 - left, 24), w - 24)}px`;
  }

  function openDatePop(input) {
    dateInput = input;
    const base = parseDate(input.value) || new Date();
    viewY = base.getFullYear();
    viewM = base.getMonth();
    renderDatePop();
    datePop.hidden = false;
    freshGrid();
    placeDatePop();
    dateTriggers.forEach(btn => {
      const on = btn.dataset.for === input.id;
      btn.setAttribute('aria-expanded', String(on));
      btn.closest('.date-control').classList.toggle('is-open', on);
    });
    popGrid.querySelector('[tabindex="0"]').focus({ preventScroll: true });
  }

  function closeDatePop(refocus = false) {
    if (datePop.hidden) return;
    const input = dateInput;
    datePop.hidden = true;
    dateInput = null;
    dateTriggers.forEach(btn => {
      btn.setAttribute('aria-expanded', 'false');
      btn.closest('.date-control').classList.remove('is-open');
    });
    if (refocus && input) dateTriggers.find(btn => btn.dataset.for === input.id)?.focus({ preventScroll: true });
  }

  function pickDate(dIso) {
    const [y, m, d] = dIso.split('-');
    dateInput.value = `${y}/${m}/${d}`;
    dateInput.dispatchEvent(new Event('input', { bubbles: true }));
    pulse(dateInput);
    const input = dateInput;
    closeDatePop();
    input.focus({ preventScroll: true });
  }

  function stepMonth(step) {
    const total = viewY * 12 + viewM + step;
    viewY = Math.floor(total / 12);
    viewM = total - viewY * 12;
    renderDatePop();
    freshGrid();
  }

  let freshTimer = 0;
  function freshGrid() {
    popGrid.classList.remove('is-fresh');
    void popGrid.offsetWidth;
    popGrid.classList.add('is-fresh');
    clearTimeout(freshTimer);
    freshTimer = setTimeout(() => popGrid.classList.remove('is-fresh'), 520);
  }

  dateTriggers.forEach(btn => {
    btn.addEventListener('click', () => {
      const input = document.getElementById(btn.dataset.for);
      if (!datePop.hidden && dateInput === input) closeDatePop(true);
      else openDatePop(input);
    });
    document.getElementById(btn.dataset.for).addEventListener('keydown', event => {
      if (event.key === 'ArrowDown' && datePop.hidden) {
        event.preventDefault();
        openDatePop(event.currentTarget);
      }
    });
  });

  datePop.addEventListener('click', event => {
    const nav = event.target.closest('.date-nav');
    if (nav) { stepMonth(+nav.dataset.step); return; }
    const day = event.target.closest('.date-day');
    if (day) { pickDate(day.dataset.date); return; }
    if (event.target.closest('[data-role="today"]')) { pickDate(isoOf(new Date())); return; }
    if (event.target.closest('[data-role="clear"]')) {
      dateInput.value = '';
      dateInput.dispatchEvent(new Event('input', { bubbles: true }));
      renderDatePop();
    }
  });

  datePop.addEventListener('keydown', event => {
    if (event.key === 'Escape') { event.preventDefault(); closeDatePop(true); return; }
    if (event.key === 'Tab') {
      const list = [...datePop.querySelectorAll('button')].filter(btn => btn.tabIndex !== -1);
      const first = list[0], last = list[list.length - 1];
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
      return;
    }
    if (event.key === 'PageUp' || event.key === 'PageDown') {
      event.preventDefault();
      stepMonth(event.key === 'PageUp' ? -1 : 1);
      popGrid.querySelector('[tabindex="0"]').focus();
      return;
    }
    const moves = { ArrowLeft: -1, ArrowRight: 1, ArrowUp: -7, ArrowDown: 7 };
    const day = event.target.closest('.date-day');
    if (!day || !(event.key in moves)) return;
    event.preventDefault();
    const d = parseDate(day.dataset.date);
    d.setDate(d.getDate() + moves[event.key]);
    if (d.getFullYear() !== viewY || d.getMonth() !== viewM) {
      viewY = d.getFullYear();
      viewM = d.getMonth();
      renderDatePop();
    }
    const target = popGrid.querySelector(`[data-date="${isoOf(d)}"]`);
    if (target) {
      popGrid.querySelector('[tabindex="0"]').tabIndex = -1;
      target.tabIndex = 0;
      target.focus();
    }
  });

  document.addEventListener('pointerdown', event => {
    if (datePop.hidden) return;
    if (datePop.contains(event.target) || event.target.closest('.date-trigger') || event.target === dateInput) return;
    closeDatePop();
  });

  form.addEventListener('input', event => {
    if (datePop.hidden || event.target !== dateInput) return;
    const d = parseDate(dateInput.value);
    if (d) { viewY = d.getFullYear(); viewM = d.getMonth(); }
    renderDatePop();
  });

  function queuePlace() {
    if (placeQueued || datePop.hidden) return;
    placeQueued = true;
    requestAnimationFrame(() => {
      placeQueued = false;
      if (!datePop.hidden) placeDatePop();
    });
  }
  window.addEventListener('scroll', queuePlace, true);
  window.addEventListener('resize', queuePlace);

  openFolder.addEventListener('click', async () => {
    if (!lastFolder || openFolder.disabled) return;
    openFolder.disabled = true;
    openStatus.textContent = '';
    openError.textContent = '';
    try {
      const response = await fetch('/api/ticket/open', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ folder: lastFolder }),
      });
      const data = await response.json();
      if (!response.ok || !data.ok) throw new Error(data.error || '打开文件夹失败。');
      if (data.foreground === false) {
        openStatus.textContent = '已请求打开文件夹。若窗口仍在后台，请从任务栏切换到资源管理器。';
      }
    } catch (err) {
      openError.textContent = err.message || '打开文件夹失败。';
    } finally {
      openFolder.disabled = false;
    }
  });

  form.addEventListener('submit', async event => {
    event.preventDefault();
    if (create.disabled) return;
    const values = Object.fromEntries(new FormData(form));
    for (const key of Object.keys(values)) values[key] = values[key].replace(/[\t\r\n]+/g, ' ').trim();
    if (!values.icon_id) delete values.icon_id;
    create.disabled = reset.disabled = source.disabled = iconField.disabled = true;
    for (const el of fields) el.disabled = true;
    dateTriggers.forEach(btn => { btn.disabled = true; });
    hoursButtons.forEach(btn => { btn.disabled = true; });
    closeDatePop();
    error.textContent = '';
    openStatus.textContent = '';
    openError.textContent = '';
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
      const applied = Boolean(values.icon_id && ticket.icon_id === values.icon_id);
      iconApplied.hidden = !applied;
      iconWarning.textContent = '';
      if (ticket.icon_error) {
        iconWarning.textContent = `文件夹图标设置失败。${ticket.icon_error}`;
      } else if (values.icon_id && !applied) {
        iconWarning.textContent = '传票已保存，服务未确认应用图标。请重启工具服务后保留当前票面和图标，再点一次「确认创建」。';
      }
      success.hidden = false;
      void pass.offsetWidth;
      pass.classList.add('is-created');
      success.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    } catch (err) {
      error.textContent = err.message || '创建失败，请重试。';
    } finally {
      create.disabled = reset.disabled = source.disabled = iconField.disabled = false;
      for (const el of fields) el.disabled = false;
      dateTriggers.forEach(btn => { btn.disabled = false; });
      hoursButtons.forEach(btn => { btn.disabled = false; });
    }
  });

  drawLive();
  loadIcons();
})();
