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

  const form = document.getElementById('ticketForm');
  const source = document.getElementById('ticketSource');
  const fieldset = document.getElementById('ticketFields');
  const create = document.getElementById('ticketCreate');
  const reset = document.getElementById('ticketReset');
  const recognition = document.getElementById('ticketRecognition');
  const result = document.getElementById('ticketResult');
  const error = document.getElementById('ticketError');
  const fields = [...fieldset.querySelectorAll('input, textarea')];
  const manualFields = new Set();

  source.addEventListener('input', () => {
    const ticket = parseTicket(source.value);
    for (const input of fields) {
      if (!manualFields.has(input.name)) input.value = ticket[input.name] ?? '';
    }
    const count = Object.keys(ticket).length;
    recognition.textContent = count ? `识别到 ${count} 项，请核对后确认。`
      : source.value.trim() ? '没有识别到任务信息，可以手动填写。' : '';
  });

  fieldset.addEventListener('input', event => manualFields.add(event.target.name));

  form.addEventListener('reset', () => {
    manualFields.clear();
    recognition.textContent = '';
    result.textContent = '';
    error.textContent = '';
    source.focus();
  });

  form.addEventListener('submit', async event => {
    event.preventDefault();
    if (create.disabled) return;
    const values = Object.fromEntries(new FormData(form));
    values.parent = values.parent.replace(/\r?\n/g, ' ');
    create.disabled = reset.disabled = fieldset.disabled = source.disabled = true;
    result.textContent = '';
    error.textContent = '';
    try {
      const response = await fetch('/api/ticket/create', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(values),
      });
      const ticket = await response.json();
      if (!response.ok || !ticket.ok) throw new Error(ticket.error || '创建失败，请重试。');
      result.textContent = `已保存 ${ticket.filename}\n${ticket.folder}`;
    } catch (err) {
      error.textContent = err.message || '创建失败，请重试。';
    } finally {
      create.disabled = reset.disabled = fieldset.disabled = source.disabled = false;
    }
  });
})();
