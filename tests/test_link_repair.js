const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

async function page(initial = {}, {kind = 'links', storage = new Map(), postError = '', failPoll = false} = {}) {
  const elements = {}, requests = [], timers = [], links = [];
  let status = initial;
  const get = id => elements[id] ||= {value:'', textContent:'', innerHTML:'', disabled:true, checked:true, style:{},
    events:{}, addEventListener(name, handler) {this.events[name] = handler;}};
  const context = vm.createContext({
    document: {getElementById:get, body:{dataset:{repairKind:kind}}},
    window: {excelLinkHtml:(label, target) => {links.push({label, target}); return label;}},
    localStorage: {getItem:key => storage.get(key), setItem:(key, value) => storage.set(key, value)},
    setTimeout:callback => {timers.push(callback); return timers.length;}, clearTimeout() {},
    fetch: async (url, options) => {
      requests.push({url, options});
      if (options?.method !== 'POST' && failPoll) { failPoll = false; throw Error('network unavailable'); }
      const post = options?.method === 'POST';
      return {ok:!(post && postError), json:async () => post ? {ok:!postError, error:postError} : status};
    },
  });
  for (const name of ['highlight_colors.js', 'repair.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../static', name), 'utf8'), context);
  }
  const flush = () => new Promise(resolve => setImmediate(resolve));
  await flush();
  return {get, requests, storage, timers, links, setStatus:value => {status = value;}, flush};
}

const hit = {sheet:'Main', addr:'A3', anchor:'A3', formula:'=#REF!*2', error:'引用错误'};
const result = {folder:'D:/current/001', old_folder:'D:/backup', files:[{name:'book.xlsx', old:'D:/backup/001/book.xlsx', cells:[{
  ...hit, action:'restore', status:'ok', source:'Main!A3', source_formula:'=Old!A1*2', value:20, display:'20', reason:'已恢复',
}, {...hit, addr:'A4', action:'skip', reason:'备份本身有错误'}]}]};

test('one click sends both folders in a single background repair request', async () => {
  const p = await page();
  p.get('repairFolder').value = result.folder;
  await p.get('repairRun').events.click();
  assert.equal(p.requests.filter(request => request.options?.method === 'POST').length, 0);
  assert.match(p.get('repairError').textContent, /备份文件夹/);
  p.get('repairOldFolder').value = result.old_folder;
  p.setStatus({applied:result});
  await p.get('repairRun').events.click();
  const post = p.requests.find(request => request.options?.method === 'POST');
  assert.equal(post.url, '/api/link_repair/run');
  assert.deepEqual(JSON.parse(post.options.body), {folder:result.folder, old_folder:result.old_folder});
  assert.equal(p.storage.get('clear_handle_excel.oldFolder'), result.old_folder);
  assert.match(p.get('repairProgress').textContent, /恢复正常 1 处；未修复 1 处/);
  assert.match(p.get('repairResults').innerHTML, /备份本身有错误/);
});

test('page reload restores results, source folders and recovery values', async () => {
  const p = await page({applied:result});
  assert.equal(p.get('repairFolder').value, result.folder);
  assert.equal(p.get('repairOldFolder').value, result.old_folder);
  assert.match(p.get('repairResults').innerHTML, /=Old!A1\*2/);
  assert.match(p.get('repairResults').innerHTML, /20/);
  assert.equal(p.get('repairRun').disabled, false);
});

test('file titles, backup paths and cell locations all open the same server-paired comparison', async () => {
  for (const kind of ['links', 'colors', 'hidden']) {
    const p = await page({applied:result}, {kind});
    assert.equal(p.links.length, 5);
    assert.ok(p.links.every(link => link.target.compare === kind && link.target.file === 'book.xlsx'));
    assert.equal(p.links[0].target.cell, 'A3');
    assert.match(p.get('repairResults').innerHTML, /同时打开当前与备份/);
    assert.match(p.get('repairResults').innerHTML, /右侧备份（只读）/);
  }
  const p = await page({applied:{folder:result.folder, files:[{name:'missing.xlsx',cells:[]}]}});
  assert.equal(p.links[0].target.compare, undefined);
  assert.ok(!p.get('repairResults').innerHTML.includes('同时打开'));
});

test('running task disables duplicate writes and folder changes, and resumes polling after reload', async () => {
  const p = await page({running:true, phase:'repair', stage:'preview', folder:'current', old_folder:'backup', done:1, total:3, current:'book.xlsx'});
  assert.equal(p.get('repairRun').disabled, true);
  assert.equal(p.get('repairFolder').disabled, true);
  assert.equal(p.get('repairOldFolder').value, 'backup');
  assert.match(p.get('repairProgress').textContent, /对比备份 2 \/ 3/);
  await p.get('repairRun').events.click();
  assert.ok(!p.requests.some(request => request.options?.method === 'POST'));
  p.setStatus({applied:result});
  await p.timers.pop()();
  assert.equal(p.get('repairRun').disabled, false);
  assert.match(p.get('repairResultTitle').textContent, /修复结果/);
});

test('colors page reuses validated presets and history and retains skipped colors', async () => {
  const storage = new Map([
    ['clear_handle_excel.highlightColors', JSON.stringify({fill:'#123456', font:'#fedcba'})],
    ['clear_handle_excel.highlightColorHistory', JSON.stringify([{fill:'#123456',font:'#FEDCBA'}, {fill:'red',font:'#000000'}])],
  ]);
  const p = await page({}, {kind:'colors', storage});
  assert.match(p.get('repairColorsSummary').textContent, /9 组/);
  assert.match(p.get('repairColors').innerHTML, /#123456/);
  p.get('repairFolder').value = result.folder; p.get('repairOldFolder').value = result.old_folder;
  p.setStatus({applied:{...result, files:[{name:'book.xlsx', cells:[{sheet:'Main', addr:'B7', fill:'#123456', font:'#FEDCBA', old_fill:'none', old_font:'#000000', status:'ok', reason:'已恢复'}, {sheet:'Main',addr:'B8', action:'skip', reason:'内容不一致'}]}]}});
  await p.get('repairRun').events.click();
  const post = p.requests.find(request => request.options?.method === 'POST');
  assert.equal(post.url, '/api/batch/find_marks/run');
  assert.equal(JSON.parse(post.options.body).colors.length, 9);
  assert.match(p.get('repairResults').innerHTML, /无填充/);
  assert.match(p.get('repairProgress').textContent, /恢复颜色 1 处；未修复 1 处/);
});

test('failed files never report pending repairs as successful, and workbook text is escaped', async () => {
  const p = await page({applied:{folder:'folder', files:[{name:'<img src=x>', error:'无法保存', cells:[{...result.files[0].cells[0], formula:'=<script>alert(1)</script>'}]}]}});
  assert.match(p.get('repairProgress').textContent, /恢复正常 0 处；未修复 1 处；1 份/);
  assert.ok(!p.get('repairResults').innerHTML.includes('<script>'));
  assert.ok(!p.get('repairResults').innerHTML.includes('<img'));
  assert.match(p.get('repairResults').innerHTML, /未修复：无法保存/);
});

test('initial polling failures retry while keeping duplicate operations disabled', async () => {
  const p = await page({}, {failPoll:true});
  assert.equal(p.get('repairRun').disabled, true);
  await p.get('repairRun').events.click();
  assert.ok(!p.requests.some(request => request.options?.method === 'POST'));
  assert.equal(p.timers.length, 1);
  await p.timers.pop()();
  assert.equal(p.get('repairRun').disabled, false);
  assert.equal(p.get('repairError').textContent, '');
});

test('server rejection is shown and a fresh scan can be retried', async () => {
  const p = await page({}, {postError:'后台任务正在执行'});
  p.get('repairFolder').value = 'current'; p.get('repairOldFolder').value = 'backup';
  await p.get('repairRun').events.click();
  assert.match(p.get('repairError').textContent, /后台任务正在执行/);
  assert.equal(p.get('repairRun').disabled, false);
});

test('hidden page sends selected scope and distinguishes hidden columns from narrow widths', async () => {
  const data = {...result, options:{sheets:true, columns:true, narrow:false}, files:[{name:'book.xlsx', old:'D:/backup/book.xlsx', cells:[
    {sheet:'Main', old_sheet:'OldMain', addr:'B:D', anchor:'B1', kind:'columns', status:'ok', current_state:'可见，列宽 20', old_state:'隐藏列', reason:'已恢复备份隐藏状态'},
    {sheet:'Main', old_sheet:'OldMain', addr:'E:E', anchor:'E1', kind:'narrow', action:'skip', current_state:'可见', old_state:'超窄列，列宽 0.25', reason:'内容不同'},
  ]}]};
  const p = await page({applied:data}, {kind:'hidden'});
  assert.equal(p.get('repairNarrow').checked, false);
  assert.match(p.get('repairResults').innerHTML, /隐藏列/);
  assert.match(p.get('repairResults').innerHTML, /超窄列/);
  assert.match(p.get('repairResults').innerHTML, /OldMain/);
  assert.match(p.get('repairProgress').textContent, /恢复隐藏设置 1 处；未修复 1 处/);
  assert.match(p.get('repairProgress').textContent, /按列段计数/);
  await p.get('repairRun').events.click();
  const post = p.requests.find(request => request.options?.method === 'POST');
  assert.equal(post.url, '/api/hidden_repair/run');
  assert.deepEqual(JSON.parse(post.options.body).options, data.options);
  assert.ok(p.links.every(link => link.target.compare === 'hidden'));
  assert.equal(p.links[0].target.cell, 'B1');
});

test('hidden page blocks empty scope and disables scope changes while running', async () => {
  const p = await page({}, {kind:'hidden'});
  p.get('repairFolder').value = 'current'; p.get('repairOldFolder').value = 'old';
  for (const id of ['repairSheets', 'repairColumns', 'repairNarrow']) p.get(id).checked = false;
  await p.get('repairRun').events.click();
  assert.ok(!p.requests.some(request => request.options?.method === 'POST'));
  assert.match(p.get('repairError').textContent, /至少选择/);
  const running = await page({running:true, phase:'repair', stage:'apply'}, {kind:'hidden'});
  for (const id of ['repairSheets', 'repairColumns', 'repairNarrow']) assert.equal(running.get(id).disabled, true);
});
