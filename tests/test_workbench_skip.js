const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'static', 'app.js'), 'utf8');

function workbench() {
  const elements = new Map();
  const context = vm.createContext({
    document: {
      hidden: false,
      addEventListener() {},
      getElementById(id) {
        if (!elements.has(id)) elements.set(id, { dataset: {}, classList: { toggle() {} }, innerHTML: '' });
        return elements.get(id);
      },
    },
    window: { addEventListener() {} },
    localStorage: { getItem() { return null; }, setItem() {} },
    setInterval() {}, setTimeout() {}, clearTimeout() {},
    fetch() { return new Promise(() => {}); }, // 不访问正在运行的服务
    alerts: [], requests: [],
  });
  vm.runInContext(source, context);
  vm.runInContext(`
    alert = message => alerts.push(message);
    confirm = () => true;
    stateLoaded = true;
    stepTab = 'fix';
    latestItems = {
      i1: {id: 'i1', status: 'pending', cell_addr: 'A5', text: 'one'},
      i2: {id: 'i2', status: 'pending', cell_addr: 'B5', text: 'two'},
      i3: {id: 'i3', status: 'pending', cell_addr: 'A6', text: 'next'},
    };
    triageKept = new Set(['i1', 'i2', 'i3']);
    fixData = {auto: false, groups: [
      {id: 'g1', item_ids: ['i1', 'i2'], status: 'failed', attempts: [], manual: false, seq: 1, row: 5},
      {id: 'g2', item_ids: ['i3'], status: 'ready', attempts: [], manual: false, seq: 2, row: 6},
    ]};
    fixCurrentId = 'g1';
    syncSkips = async () => {};
    postAction = async (url, payload) => {
      requests.push({url, payload});
      return {ok: true, changed: payload.item_ids.length, errors: []};
    };
    pollState = async () => {
      for (const {payload} of requests) for (const id of payload.item_ids) latestItems[id].status = 'accepted';
    };
    pollAutofix = async () => renderFix();
  `, context);
  return { context, elements, run: code => vm.runInContext(code, context) };
}

test('skip button is available during preparation, after failure, and during manual review', () => {
  const w = workbench();
  for (const status of ['queued', 'running', 'ready', 'failed', 'error']) {
    for (const manual of [false, true]) {
      w.run(`fixData.groups[0].status = ${JSON.stringify(status)}; fixData.groups[0].manual = ${manual}; renderFix();`);
      assert.match(w.elements.get('fixStage').innerHTML, /跳过本行（无需修改）/);
      assert.match(w.elements.get('fixStage').innerHTML, /先看下一个/);
    }
  }
});

test('skip saves every problem in the row and advances; stale polling cannot restore the skipped row', async () => {
  const w = workbench();
  await w.run(`skipFixGroup('g1')`);
  assert.deepEqual(JSON.parse(JSON.stringify(w.context.requests)), [
    {url: '/api/autofix/skip', payload: {id: 'g1', item_ids: ['i1', 'i2']}},
  ]);
  assert.equal(w.run('fixCurrentId'), 'g2');
  assert.equal(w.run(`latestItems.i1.status`), 'accepted');
  w.run(`fixData.groups.unshift({id: 'g1', item_ids: ['i1', 'i2'], status: 'failed'});`);
  assert.deepEqual(Array.from(w.run('fixQueue().map(g => g.id)')), ['g2']);
});

test('failed persistence keeps the current row for retry', async () => {
  const w = workbench();
  w.run(`postAction = async () => { throw new Error('database unavailable'); };`);
  await w.run(`skipFixGroup('g1')`);
  assert.equal(w.run('fixCurrentId'), 'g1');
  assert.equal(w.run('latestItems.i1.status'), 'pending');
  assert.equal(w.run('fixSkipping.size'), 0);
  assert.match(w.context.alerts[0], /跳过本行失败/);
});

test('busy jobs and cancelled checkpoint replacement do not submit skip', async () => {
  const w = workbench();
  w.run('isBusy = true;');
  await w.run(`skipFixGroup('g1')`);
  w.run('isBusy = false; canResume = true; confirm = () => false;');
  await w.run(`skipFixGroup('g1')`);
  assert.equal(w.context.requests.length, 0);
  assert.equal(w.run('latestItems.i1.status'), 'pending');
});

test('a row queued for writing cannot be skipped', async () => {
  const w = workbench();
  w.run(`fixData.groups[0].status = 'applying';`);
  await w.run(`skipFixGroup('g1')`);
  assert.equal(w.context.requests.length, 0);
});

test('look at next only defers the row without recording skip', () => {
  const w = workbench();
  w.run(`deferFixGroup('g1');`);
  assert.equal(w.run('fixCurrentId'), 'g2');
  assert.equal(w.run('latestItems.i1.status'), 'pending');
  assert.equal(w.context.requests.length, 0);
  assert.equal(w.run('fixQueue().length'), 2);
});
