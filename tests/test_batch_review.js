const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '..', 'static', 'batch.js'), 'utf8');

function page() {
  const elements = new Map();
  function element() {
    return { value: '', innerHTML: '', dataset: {}, style: {}, classList: { toggle() {}, add() {} },
      listeners: {}, addEventListener(name, fn) { this.listeners[name] = fn; },
      scrollIntoView() {}, querySelector() { return element(); }, querySelectorAll() { return []; } };
  }
  const requests = [];
  const listeners = {};
  const context = vm.createContext({
    document: { activeElement: null, getElementById(id) { if (!elements.has(id)) elements.set(id, element()); return elements.get(id); },
      addEventListener(name, fn) { listeners[name] = fn; }, querySelectorAll() { return []; } },
    localStorage: { getItem() { return ''; } }, window: { addEventListener() {} }, appConfirm: async () => true,
    setInterval() {}, fetch() { return new Promise(() => {}); },
    MutationObserver: class { observe() {} },
    excelLinkHtml(label) { return `<a>${label}</a>`; }, Date, requests,
  });
  vm.runInContext(source, context);
  const run = code => vm.runInContext(code, context);
  run(`batchState = {busy: false, job: null, logs: [], progress: {}, folder: 'C:/Excel', cards: [
    {id:'a', seq:1, wb_name:'book.xlsx', sheet_name:'Sheet', cell_addr:'A5', color:'blue', verified:true, images:{before:'/before.png', after:'/after.png'}},
    {id:'b', seq:2, wb_name:'book.xlsx', sheet_name:'Sheet', cell_addr:'A6', color:'yellow', finalized:true, images:{}}
  ]}; poll = async () => {}; post = async (url, payload) => {requests.push({url, payload}); return {ok:true};};`);
  return { run, elements, requests, listeners };
}

test('both directions cycle through the three colors', () => {
  const p = page();
  assert.equal(p.run(`nextColor('blue', 1)`), 'green');
  assert.equal(p.run(`nextColor('green', 1)`), 'yellow');
  assert.equal(p.run(`nextColor('yellow', 1)`), 'blue');
  assert.equal(p.run(`nextColor('blue', -1)`), 'yellow');
  assert.equal(p.run(`nextColor('yellow', -1)`), 'green');
});

test('color changes submit only a marker; no restore request is sent', async () => {
  const p = page();
  await p.run(`changeColor('a')`);
  assert.deepEqual(p.requests.map(r => r.url), ['/api/batch/mark']);
  assert.equal(p.requests[0].payload.color, 'green');
});

test('an unverified yellow card can become blue with a single click', async () => {
  const p = page();
  await p.run(`changeColor('b')`);
  assert.equal(p.requests[0].payload.color, 'blue');
});

test('unverified cards without images cycle freely in both directions', async () => {
  const p = page();
  for (const direction of [1, -1]) {
    for (const color of ['blue', 'green', 'yellow']) {
      p.run(`batchState.cards[1].color = '${color}'; batchState.cards[1].verified = false;`);
      await p.run(`changeColor('b', ${direction})`);
      assert.equal(p.requests.at(-1).payload.color, p.run(`nextColor('${color}', ${direction})`));
    }
  }
});

test('small and large cards contain only the before and after images', () => {
  const p = page();
  for (const large of [false, true]) {
    const html = p.run(`cardHtml(batchState.cards[0], ${large})`);
    assert.equal((html.match(/<figure>/g) || []).length, 2);
    assert.match(html, /修改前/);
    assert.match(html, /修改后/);
  }
});

test('single-card navigation preserves card identity when changing colors', () => {
  const p = page();
  p.run(`openLarge('a'); moveCard(1);`);
  assert.equal(p.run('activeCardId'), 'b');
  p.run(`batchState.cards[1].color = 'green'; render();`);
  assert.equal(p.run('activeCardId'), 'b');
  p.run('moveCard(-1)');
  assert.equal(p.run('activeCardId'), 'a');
});

test('large mode adds page context while small cards keep only the comparison', () => {
  const p = page();
  p.run(`batchState.cards[0].context = {image:'/page.png', page_width:600, page_height:800, page_index:1, bbox:[60,160,180,200]}`);
  const large = p.run(`cardHtml(batchState.cards[0], true)`);
  assert.match(large, /batch-large-layout/);
  assert.match(large, /问题所在页面 · 第 2 页/);
  assert.match(large, /left:10%;top:20%;width:20%;height:5%/);
  assert.ok(large.indexOf('batch-shots') < large.indexOf('batch-context'));
  const small = p.run(`cardHtml(batchState.cards[0], false)`);
  assert.doesNotMatch(small, /batch-context|page\.png/);
});

test('large mode shows an honest placeholder when no page context exists', () => {
  const p = page();
  const large = p.run(`cardHtml(batchState.cards[0], true)`);
  assert.match(large, /暂无整页定位图/);
  assert.doesNotMatch(large, /batch-context-marker/);
});

test('failed trials are clearly labelled as unapplied in both review modes', () => {
  const p = page();
  p.run(`batchState.cards[0].after_kind = 'attempt'; batchState.cards[0].color = 'yellow';`);
  for (const large of [false, true]) {
    const html = p.run(`cardHtml(batchState.cards[0], ${large})`);
    assert.match(html, /<figcaption>最后一次尝试（未采用）<\/figcaption>/);
    assert.match(html, /alt="最后一次尝试（未采用）截图"/);
    assert.equal((html.match(/<figure>/g) || []).length, 2);
    assert.doesNotMatch(html, /<figcaption>修改后<\/figcaption>/);
  }
  p.run(`batchState.cards[0].after_kind = 'fixed'`);
  assert.match(p.run(`cardHtml(batchState.cards[0])`), /<figcaption>修改后<\/figcaption>/);
});

test('keyboard changes status and navigates while text inputs keep their keys', async () => {
  const p = page();
  p.run(`openLarge('a')`);
  const event = key => ({ key, target: { closest() { return false; } }, preventDefault() {} });
  p.listeners.keydown(event('ArrowDown'));
  assert.equal(p.run('activeCardId'), 'b');
  p.listeners.keydown(event('ArrowUp'));
  p.listeners.keydown(event('ArrowLeft'));
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(p.requests[0].payload.color, 'yellow');
  p.listeners.keydown({ ...event('ArrowDown'), target: { closest() { return true; } } });
  assert.equal(p.run('activeCardId'), 'a');
});

test('mouse wheel scrolls the page instead of switching cards', () => {
  const p = page();
  p.run(`openLarge('a')`);
  assert.equal(p.elements.get('largeView').listeners.wheel, undefined);
  assert.equal(p.run('activeCardId'), 'a');
});

test('only submitted yellow cards are eligible for manual actions', () => {
  const p = page();
  assert.deepEqual(Array.from(p.run(`manualCards().map(c => c.id)`)), ['b']);
  p.run(`batchState.cards[1].finalized = false`);
  assert.equal(p.run(`manualCards().length`), 0);
});

test('background work blocks markers, and failed marker writes show an error', async () => {
  const p = page();
  p.run('batchState.busy = true');
  await p.run(`changeColor('a')`);
  assert.equal(p.requests.length, 0);
  p.run(`batchState.busy = false; post = async () => {throw new Error('保存失败')}`);
  await p.run(`changeColor('a')`);
  assert.equal(p.elements.get('errorText').textContent, '保存失败');
  assert.equal(p.run('requestBusy'), false);
});

test('cards say what was detected and how much the row height changed', () => {
  const p = page();
  const html = p.run(`problemHtml({problems: [{label: '隐藏截断', reason: '后面的文字没显示', missing: '空'}]})`);
  assert.match(html, /隐藏截断/);
  assert.match(html, /「空」/);
  assert.equal(p.run(`handlingText({excel_row: 41, after_kind: 'fixed', row_height: {base: 15, applied: 21, saved: true}})`),
    '处理：第 41 行行高 15pt → 21pt（+6pt），已保存');
  assert.equal(p.run(`handlingText({excel_row: 41, after_kind: 'rechecked'})`),
    '处理：本行行高未修改（其他行调整后重新导出已完整显示）');
  assert.equal(p.run(`handlingText({excel_row: 41, after_kind: 'attempt', base_height: 15, attempt_height: 39})`),
    '处理：第 41 行行高试到 15pt → 39pt（+24pt）仍未通过，原文件未修改');
});

test('the selection circle toggles selection without changing color', () => {
  const p = page();
  const click = target => p.elements.get('cardList').listeners.click({ target });
  const card = { dataset: { id: 'a' } };
  const circle = { closest(sel) { return sel === '[data-id]' ? card : sel === '[data-action]' ? { dataset: { action: 'select' } } : null; } };
  click(circle);
  assert.deepEqual(Array.from(p.run('[...selectedIds]')), ['a']);
  click(circle);
  assert.equal(p.run('selectedIds.size'), 0);
  assert.equal(p.requests.length, 0);
});

test('select all and invert act on the cards in the current filter', () => {
  const p = page();
  p.run(`selectShown('all')`);
  assert.deepEqual(Array.from(p.run('[...selectedIds]')).sort(), ['a', 'b']);
  p.run(`selectedIds.delete('a'); selectShown('invert')`);
  assert.deepEqual(Array.from(p.run('[...selectedIds]')), ['a']);
  p.run(`selectedFilter = 'yellow'; selectShown('all')`);
  assert.deepEqual(Array.from(p.run('[...selectedIds]')).sort(), ['a', 'b']);
  p.run('render()');
  assert.deepEqual(Array.from(p.run('[...selectedIds]')), ['b']);
  p.run(`selectShown('none')`);
  assert.equal(p.run('selectedIds.size'), 0);
});

test('batch color change sends one request and keeps the selection', async () => {
  const p = page();
  p.run(`selectShown('all'); post = async (url, payload) => {requests.push({url, payload}); return {ok:true, changed:1, skipped:1};};`);
  await p.run(`changeSelectedColor('blue')`);
  assert.deepEqual(p.requests.map(r => r.url), ['/api/batch/mark_many']);
  assert.deepEqual(Array.from(p.requests[0].payload.card_ids).sort(), ['a', 'b']);
  assert.equal(p.requests[0].payload.color, 'blue');
  assert.equal(p.run('selectedIds.size'), 2);
  assert.equal(p.run('selectNotice'), '');
  p.run('batchState.busy = true');
  await p.run(`changeSelectedColor('green')`);
  assert.equal(p.requests.length, 1);
});

test('commit leaves cards that have not been processed yet out of the submission', async () => {
  const p = page();
  p.run(`batchState.cards.push({id:'c', seq:3, wb_name:'book.xlsx', sheet_name:'Sheet', cell_addr:'A7', color:'yellow', waiting:true, message:'等待自动处理', images:{}})`);
  p.run(`appConfirm = async text => { globalThis.asked = text; return true; }`);
  await p.elements.get('commitBtn').listeners.click();
  await new Promise(resolve => setImmediate(resolve));
  const asked = p.run('globalThis.asked');
  assert.match(asked, /另有 1 张还没轮到自动处理/);
  assert.equal(JSON.stringify(p.requests.at(-1).payload.card_ids), '["a","b"]');
});

test('committing while the automatic run is paused does not warn about losing its resume point', async () => {
  const p = page();
  p.run(`batchState.job = {kind:'batch_scan', status:'interrupted', completed:[]}; globalThis.asks = 0;
    appConfirm = async text => { if (text.includes('续跑点')) globalThis.asks++; return true; }`);
  await p.elements.get('commitBtn').listeners.click();
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(p.run('globalThis.asks'), 0);
  assert.equal(p.requests.at(-1).url, '/api/batch/commit');
});

test('scanning counts scanned files, and switching to manual is a plain request', async () => {
  const p = page();
  p.run(`batchState.busy = true; batchState.job = {kind:'batch_scan', status:'running', options:{folder:'C:/Excel'}, workbooks:['a','b'], scanned:['a'], completed:[]}; render();`);
  assert.match(p.elements.get('progressText').textContent, /正在扫描.*1 \/ 2 个文件/);
  p.run(`batchState.job.phase = 'repair'; batchState.mode = {to_manual: true}; render();`);
  assert.match(p.elements.get('progressText').textContent, /正在自动处理：0 \/ 2 个文件/);
  await p.elements.get('toManualBtn').listeners.click();
  assert.deepEqual(p.requests.map(r => r.url), ['/api/batch/to_manual']);
});

test('handed-over cards explain they were not tried and can go back to automatic', async () => {
  const p = page();
  p.run(`Object.assign(batchState.cards[1], {after_kind:'handover', handed_over:true});`);
  const html = p.run(`cardHtml(batchState.cards[1])`);
  assert.match(html, /转人工时自动还没处理到这一格/);
  assert.match(html, /可转回自动/);
  assert.match(html, /人工修改后（复检通过后显示）/);
  p.run(`batchState.mode = {to_auto: true}`);
  await p.elements.get('toAutoBtn').listeners.click();
  assert.equal(p.requests.at(-1).url, '/api/batch/to_auto');
});
