const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

async function page(initial = {}) {
  const elements = {};
  const requests = [];
  let status = initial;
  const get = id => elements[id] ||= {value:'', textContent:'', innerHTML:'', disabled:false,
    events:{}, addEventListener(name, handler) {this.events[name] = handler;}};
  const context = vm.createContext({
    document: {getElementById:get},
    window: {excelLinkHtml:label => label},
    localStorage: {getItem:() => '', setItem() {}},
    setTimeout: () => 1, clearTimeout() {},
    fetch: async (url, options) => {
      requests.push({url, options});
      return {ok:true, json:async () => options?.method === 'POST' ? {ok:true} : status};
    },
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../static/link_repair.js'), 'utf8'), context);
  const flush = () => new Promise(resolve => setImmediate(resolve));
  await flush();
  return {get, requests, setStatus:value => {status = value;}, flush};
}

const hit = {sheet:'Main', addr:'A3', anchor:'A3', formula:'=#REF!*2', error:'引用错误'};
const result = {folder:'D:/current/001', files:[{name:'book.xlsx', hits:[hit]}]};
const plan = {folder:result.folder, old_folder:'D:/backup', files:[{name:'book.xlsx', old:'D:/backup/001/book.xlsx', cells:[{
  ...hit, action:'restore', source:'Main!A3', source_formula:'=Old!A1*2', value:20, display:'20', reason:'可以恢复',
}]}]};

test('scan results enable backup comparison, and display the actual scanned folder', async () => {
  const p = await page({result});
  assert.equal(p.get('linkFolder').value, result.folder);
  assert.equal(p.get('linkPreview').disabled, false);
  assert.equal(p.get('linkApply').disabled, true);
  assert.match(p.get('linkResults').innerHTML, /#REF!\*2/);
});

test('preview shows original formula and value before enabling explicit repair', async () => {
  const p = await page({result, plan});
  assert.equal(p.get('linkApply').disabled, false);
  assert.match(p.get('linkApply').textContent, /1/);
  assert.match(p.get('linkResults').innerHTML, /=Old!A1\*2/);
  assert.match(p.get('linkResults').innerHTML, /20/);
  p.setStatus({result, applied:{folder:result.folder, files:[{name:'book.xlsx', cells:[{...plan.files[0].cells[0], status:'ok'}]}]}});
  await p.get('linkApply').events.click();
  assert.deepEqual(JSON.parse(p.requests.find(request => request.options?.method === 'POST').options.body), {});
  assert.match(p.get('linkResultTitle').textContent, /修复结果/);
});

test('editing either folder invalidates its old confirmation list', async () => {
  const p = await page({result, plan});
  p.get('linkOldFolder').events.input();
  assert.equal(p.get('linkApply').disabled, true);
  assert.equal(p.get('linkPreview').disabled, false);
  p.get('linkFolder').events.input();
  assert.equal(p.get('linkPreview').disabled, true);
  assert.equal(p.get('linkApply').disabled, true);
});

test('running task disables changes and escapes workbook contents in results', async () => {
  const p = await page({running:true, phase:'apply', done:0, total:2, current:'book.xlsx'});
  assert.equal(p.get('linkScan').disabled, true);
  assert.equal(p.get('linkFolder').disabled, true);
  assert.match(p.get('linkProgress').textContent, /1 \/ 2/);
  const q = await page({result:{folder:'folder', files:[{name:'<img src=x>', hits:[{...hit, formula:'=<script>alert(1)</script>'}]}]}});
  assert.ok(!q.get('linkResults').innerHTML.includes('<script>'));
  assert.ok(!q.get('linkResults').innerHTML.includes('<img'));
});
