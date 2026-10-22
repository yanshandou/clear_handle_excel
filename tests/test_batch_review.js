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
      scrollIntoView() {}, querySelector() { return element(); } };
  }
  const requests = [];
  const listeners = {};
  const context = vm.createContext({
    document: { activeElement: null, getElementById(id) { if (!elements.has(id)) elements.set(id, element()); return elements.get(id); },
      addEventListener(name, fn) { listeners[name] = fn; } },
    localStorage: { getItem() { return ''; } }, window: { confirm() { return true; }, addEventListener() {} },
    setInterval() {}, fetch() { return new Promise(() => {}); },
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

test('an automatically failed yellow card can be ignored with a single click', async () => {
  const p = page();
  await p.run(`changeColor('b')`);
  assert.equal(p.requests[0].payload.color, 'green');
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

test('mouse wheel moves cards and throttles continuous trackpad events', () => {
  const p = page();
  p.run(`openLarge('a')`);
  const wheel = p.elements.get('largeView').listeners.wheel;
  wheel({ deltaY: 100, preventDefault() {} });
  assert.equal(p.run('activeCardId'), 'b');
  wheel({ deltaY: -100, preventDefault() {} });
  assert.equal(p.run('activeCardId'), 'b');
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
