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
        if (!elements.has(id)) elements.set(id, { dataset: {}, style: {setProperty() {}},
          classList: { toggle() {}, add() {}, remove() {} }, addEventListener() {}, innerHTML: '' });
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
    originalPollState = pollState;
    originalPollAutofix = pollAutofix;
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
      return {ok: true, changed: payload.item_ids?.length || Object.keys(payload.decisions || {}).length, errors: []};
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

test('new recheck findings do not lock an already started fix session', () => {
  const w = workbench();
  w.run(`
    setStepTab('fix');
    latestItems.new = {id: 'new', status: 'manual_review', text: 'new finding'};
    renderFix();
  `);
  assert.equal(w.run('triageComplete()'), false);
  assert.equal(w.run('fixAvailable()'), true);
  assert.match(w.elements.get('fixStage').innerHTML, /当前还有 1 处可筛选/);
  assert.match(w.elements.get('fixStage').innerHTML, /第 5 行/);
  assert.doesNotMatch(w.elements.get('fixStage').innerHTML, /还有 1 处没筛选/);
  w.run(`setStepTab('triage'); setStepTab('fix'); renderFix();`);
  assert.match(w.elements.get('fixStage').innerHTML, /第 5 行/);
});

test('successful manual recheck advances to next row while failed recheck stays current', () => {
  const w = workbench();
  w.run(`manualAction = {kind: 'recheck', id: 'i1', submitted: true};`);
  assert.equal(w.run(`finishManualAction({busy: false, job: {kind: 'recheck', status: 'completed', options: {selected_ids: ['i1']}}})`), true);
  assert.equal(w.run('fixCurrentId'), 'g1');
  w.run(`
    manualAction = {kind: 'recheck', id: 'i1', submitted: true};
    latestItems.i1.status = latestItems.i2.status = 'fixed';
    finishManualAction({busy: false, job: {kind: 'recheck', status: 'completed', options: {selected_ids: ['i1']}}});
    renderFix();
  `);
  assert.equal(w.run('fixCurrentId'), 'g2');
  assert.equal(w.run('stepTab'), 'fix');
});

test('pending manual action shows progress and prevents duplicate submission or candidate writes', async () => {
  const w = workbench();
  w.run(`
    fixData.groups[0].manual = true;
    postAction = async (url, payload) => {requests.push({url, payload}); return {ok: true};};
    pollState = async () => {};
  `);
  await w.run(`manualRecheck('i1')`);
  w.run('renderFix()');
  assert.match(w.elements.get('fixStage').innerHTML, /正在复检/);
  await w.run(`manualRecheck('i1')`);
  await w.run(`manualOpen('i1')`);
  await w.run(`chooseAttempt('g2', 1)`);
  assert.equal(w.context.requests.length, 1);
  assert.equal(w.run('fixChosen.size'), 0);
  assert.equal(w.run(`finishManualAction({busy: false, job: {kind: 'manual_open', status: 'completed', options: {selected_ids: ['i1']}}})`), false);
  w.run(`manualAction.submitted = false;`);
  assert.equal(w.run(`finishManualAction({busy: false, job: {kind: 'recheck', status: 'completed', options: {selected_ids: ['i1']}}})`), false);
});

test('failed recheck renders fresh result as recheck without a recommendation', () => {
  const w = workbench();
  w.run(`
    Object.assign(fixData.groups[0], {manual: true, recheck_token: 'check1',
      attempts: [{height: 33, ok: false}], message: '本次复检仍有问题'});
    renderFix();
  `);
  const html = w.elements.get('fixStage').innerHTML;
  assert.match(html, /本次复检 33pt/);
  assert.match(html, /本次复检仍有问题/);
  assert.doesNotMatch(html, /fix-tile-badge/);
  assert.match(html, /跳过本行（无需修改）/);
});

test('scanning allows decisions and choosing, while non-scan jobs still block them', async () => {
  const w = workbench();
  w.run(`isBusy = true; scanRunning = true;`);
  await w.run(`chooseAttempt('g2', 1)`);
  assert.equal(w.context.requests[0].url, '/api/autofix/choose');
  w.run(`scanRunning = false;`);
  await w.run(`chooseAttempt('g1', 1)`);
  assert.equal(w.context.requests.length, 1);
});

test('triage page heartbeats request background pictures before all issues are screened', async () => {
  const w = workbench();
  w.run(`
    stepTab = 'triage'; isBusy = true; scanRunning = true;
    triageKept.delete('i3');
    fetch = async url => {requests.push({url}); return {ok: true, json: async () => fixData};};
  `);
  await w.run(`originalPollAutofix()`);
  assert.equal(w.context.requests[0].url, '/api/autofix?active=1');
  assert.equal(w.run('triageComplete()'), false);
});

test('two progress bars distinguish scanned files from currently available findings', () => {
  const w = workbench();
  w.run(`
    scanRunning = true; scanProgress = {total: 100, completed: 30, status: 'running'};
    triageKept.delete('i3'); renderPipelineProgress();
  `);
  assert.match(w.elements.get('scanPipelineText').textContent, /30 \/ 100 个文件 · 剩余 70/);
  assert.equal(w.elements.get('scanPipelineFill').style.width, '30%');
  assert.match(w.elements.get('supplyProgressText').textContent, /当前已产出 3 处 · 已筛选 2 处 · 可筛选 1 处/);
  w.run(`triageKept.add('i3'); renderTriage();`);
  assert.match(w.elements.get('triageStage').innerHTML, /等待扫描产生新结果/);
  assert.doesNotMatch(w.elements.get('triageStage').innerHTML, /已全部筛选完/);
});

test('an empty downstream queue while upstream is working shows waiting rather than completion', () => {
  const w = workbench();
  w.run(`scanRunning = true; latestItems = {}; fixData.groups = []; renderFix();`);
  assert.match(w.elements.get('fixStage').innerHTML, /等待扫描或筛选交接新行/);
  assert.doesNotMatch(w.elements.get('fixStage').innerHTML, /都处理完了/);
});

test('handed rows waiting for a queue snapshot are not reported as no work', () => {
  const w = workbench();
  w.run(`fixData.groups = []; renderFix();`);
  assert.match(w.elements.get('fixStage').innerHTML, /正在同步处理队列/);
  assert.doesNotMatch(w.elements.get('fixStage').innerHTML, /没有需要处理的切字/);
});

test('keep is persisted and a failed save restores the item to triage', async () => {
  const w = workbench();
  w.run(`
    serverItems = latestItems;
    for (const item of Object.values(serverItems)) item.triage_keep = true;
    serverItems.i3.triage_keep = false;
    triageKept.delete('i3');
    pollState = async () => {};
    postAction = async (url, payload) => {
      requests.push({url, payload});
      for (const [id, keep] of Object.entries(payload.decisions)) serverItems[id].triage_keep = keep;
      return {ok: true};
    };
    setKeep('i3', true);
  `);
  await w.run('keepRequest');
  assert.equal(w.context.requests[0].url, '/api/manual/keep');
  assert.equal(w.run('triageKept.has("i3")'), true);
  w.run(`postAction = async () => {throw new Error('disk full');}; setKeep('i3', false);`);
  await w.run('keepRequest');
  assert.equal(w.run('triageKept.has("i3")'), true);
  assert.match(w.context.alerts[0], /保存保留判断失败/);
});

test('later review hides the row immediately and moves on without opening Excel', async () => {
  const w = workbench();
  w.run(`renderFix();`);
  assert.match(w.elements.get('fixStage').innerHTML, /稍后人工复检/);
  assert.doesNotMatch(w.elements.get('fixStage').innerHTML, /人工处理（在 Excel 中打开）/);
  w.run(`pollState = async () => { latestItems.i1.later_review = true; latestItems.i2.later_review = true; };`);
  await w.run(`deferReview('g1')`);
  assert.deepEqual(JSON.parse(JSON.stringify(w.context.requests)), [
    {url: '/api/review/later', payload: {id: 'g1', item_ids: ['i1', 'i2']}},
  ]);
  assert.deepEqual(Array.from(w.run('fixQueue().map(g => g.id)')), ['g2']);
  assert.equal(w.run('fixCurrentId'), 'g2');
});

test('review panel lists queued files and offers open-all and batch recheck', () => {
  const w = workbench();
  w.run(`
    serverItems = {i1: {id: 'i1', status: 'manual_review', text: 'one', later_review: true},
                   i2: {id: 'i2', status: 'accepted', text: 'two', later_review: true}};
    latestItems.i1.wb_name = 'book.xlsx';
    reviewQueue = {'C:\\\\x\\\\🟡book.xlsx': {error: '', marks: [
      {sheet: 'S1', addr: 'A5', row: 5, item_ids: ['i1'], state: 'marked'},
      {sheet: 'S1', addr: 'B5', row: 5, item_ids: ['i2'], state: 'marked'},
    ]}};
    renderReview();
  `);
  const html = w.elements.get('reviewPanel').innerHTML;
  assert.match(html, /1 个文件 · 待修改 1 处/);
  assert.match(html, /🟡 book\.xlsx/);
  assert.match(html, /一次打开全部（1 个）/);
  assert.match(html, /统一复检/);
  assert.match(html, /已全部标黄/);
  assert.match(html, /已解决，复检时去掉黄色/);
});
