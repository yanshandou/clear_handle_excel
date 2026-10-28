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

test('colors page compares all colors independently of saved presets and retains skipped colors', async () => {
  const storage = new Map([
    ['clear_handle_excel.highlightColors', JSON.stringify({fill:'#123456', font:'#fedcba'})],
    ['clear_handle_excel.highlightColorHistory', JSON.stringify([{fill:'#123456',font:'#FEDCBA'}, {fill:'red',font:'#000000'}])],
  ]);
  const p = await page({}, {kind:'colors', storage});
  p.get('repairFolder').value = result.folder; p.get('repairOldFolder').value = result.old_folder;
  p.setStatus({applied:{...result, files:[{name:'book.xlsx', cells:[{sheet:'Main', addr:'B7', fill:'#123456', font:'#FEDCBA', old_fill:'none', old_font:'#000000', status:'ok', reason:'已恢复'}, {sheet:'Main',addr:'B8', action:'skip', reason:'内容不一致'}]}]}});
  await p.get('repairRun').events.click();
  const post = p.requests.find(request => request.options?.method === 'POST');
  assert.equal(post.url, '/api/batch/find_marks/run');
  assert.equal(JSON.parse(post.options.body).colors, undefined);
  assert.match(p.get('repairResults').innerHTML, /无填充/);
  assert.match(p.get('repairProgress').textContent, /恢复颜色 1 处；未修复 1 处/);
});

test('font-only differences show both current and backup font values in preview and applied results', async () => {
  for (const phase of ['plan', 'applied']) {
    const p = await page({[phase]:{folder:'current', files:[{name:'book.xlsx', cells:[{
      sheet:'Main', addr:'A1', fill:'none', font:'#FF0000', old_fill:'none', old_font:'#123456',
      status:'ok', reason:'已恢复字色',
    }]}]}}, {kind:'colors'});
    assert.match(p.get('repairResults').innerHTML, /<td>无填充 \/ 字色 #FF0000<\/td><td>无填充 \/ 字色 #123456<\/td>/);
  }
});

test('skipped font-only differences retain both font values and the skip reason', async () => {
  const p = await page({applied:{folder:'current', files:[{name:'book.xlsx', cells:[{
    sheet:'Main', addr:'A2', fill:'none', font:'#FF0000', old_fill:'none', old_font:'#000000',
    action:'skip', reason:'附近内容不同，未修复',
  }]}]}}, {kind:'colors'});
  assert.match(p.get('repairResults').innerHTML, /<td>无填充 \/ 字色 #FF0000<\/td><td>无填充 \/ 字色 #000000<\/td>/);
  assert.match(p.get('repairResults').innerHTML, /class="link-skip">附近内容不同，未修复/);
  assert.match(p.get('repairProgress').textContent, /恢复颜色 0 处；未修复 1 处/);
});

test('theme or pattern fills also retain existing font values', async () => {
  const p = await page({applied:{folder:'current', files:[{name:'book.xlsx', cells:[{
    sheet:'Main', addr:'A1', fill:'other', font:'#112233', old_fill:'other', old_font:'#445566', status:'ok',
  }]}]}}, {kind:'colors'});
  assert.match(p.get('repairResults').innerHTML, /<td>主题色 \/ 图案 \/ 字色 #112233<\/td><td>主题色 \/ 图案 \/ 字色 #445566<\/td>/);
});

test('unknown fonts never produce a fabricated black value or an Aa preview', async () => {
  for (const font of [undefined, null, '']) {
    const p = await page({applied:{folder:'current', files:[{name:'book.xlsx', cells:
      ['none', 'other', '#AABBCC'].map((fill, index) => ({sheet:'Main', addr:`A${index+1}`,
        fill, font, old_fill:fill, old_font:font, status:'ok'})),
    }]}}, {kind:'colors'});
    const html = p.get('repairResults').innerHTML;
    assert.match(html, /<td>无填充<\/td><td>无填充<\/td>/);
    assert.match(html, /<td>主题色 \/ 图案<\/td><td>主题色 \/ 图案<\/td>/);
    assert.match(html, /<td>#AABBCC<\/td><td>#AABBCC<\/td>/);
    assert.doesNotMatch(html, /#000000|字色|>Aa<|style=/);
  }
});

test('sheet tab colors ignore font values even when the API supplies them', async () => {
  const p = await page({applied:{folder:'current', files:[{name:'book.xlsx', cells:
    ['none', 'other', '#FFC000'].map((fill, index) => ({kind:'tab', sheet:'Main', addr:`A${index+1}`,
      fill, font:'#112233', old_fill:fill, old_font:'#445566', status:'ok'})),
  }]}}, {kind:'colors'});
  const html = p.get('repairResults').innerHTML;
  assert.match(html, /<td>标签：无标签颜色<\/td><td>无标签颜色<\/td>/);
  assert.match(html, /<td>标签：主题色<\/td><td>主题色<\/td>/);
  assert.match(html, /style="background:#FFC000"/);
  assert.doesNotMatch(html, /字色|#112233|#445566|#000000|>Aa<|;color:/);
});

test('untrusted font strings are escaped as text and never used as CSS', async () => {
  const font = '#123456" onmouseover="alert(1)"><img src=x onerror=alert(1)>&\'';
  const escapedFont = '#123456&quot; onmouseover=&quot;alert(1)&quot;&gt;&lt;img src=x onerror=alert(1)&gt;&amp;&#39;';
  const cssFont = '#123456;background:url(javascript:alert(1))';
  const p = await page({applied:{folder:'current', files:[{name:'book.xlsx', cells:
    ['none', 'other', '#AABBCC'].map((fill, index) => ({sheet:'Main', addr:`A${index+1}`,
      fill, font, old_fill:fill, old_font:cssFont, status:'ok'})),
  }]}}, {kind:'colors'});
  const html = p.get('repairResults').innerHTML;
  assert.equal(html.split(escapedFont).length - 1, 3);
  assert.equal(html.split(cssFont).length - 1, 3);
  assert.doesNotMatch(html, /<img|<script|style=| onmouseover="| onerror="/);
});

test('failed files never report pending repairs as successful, and workbook text is escaped', async () => {
  const p = await page({applied:{folder:'folder', files:[{name:'<img src=x>', error:'无法保存', cells:[{...result.files[0].cells[0], formula:'=<script>alert(1)</script>'}]}]}});
  assert.match(p.get('repairProgress').textContent, /恢复正常 0 处；未修复 1 处；1 份/);
  assert.ok(!p.get('repairResults').innerHTML.includes('<script>'));
  assert.ok(!p.get('repairResults').innerHTML.includes('<img'));
  assert.match(p.get('repairResults').innerHTML, /未修复：无法保存/);
});

test('color results label sheet tabs separately and count restored tabs', async () => {
  const p = await page({applied:{folder:'current', files:[{name:'book.xlsx', old:'backup/book.xlsx', cells:[
    {sheet:'Main', old_sheet:'OldMain', kind:'tab', addr:'A1', anchor:'A1', fill:'#FFC000', old_fill:'none', status:'ok', reason:'已清除标签色'},
    {sheet:'Main', addr:'A1', fill:'#C0007A', old_fill:'none', status:'ok', reason:'已恢复单元格'},
  ]}]}}, {kind:'colors'});
  assert.match(p.get('repairResults').innerHTML, /Main · Sheet 标签/);
  assert.match(p.get('repairResults').innerHTML, /OldMain · Sheet 标签/);
  assert.match(p.get('repairResults').innerHTML, /无标签颜色/);
  assert.match(p.get('repairResults').innerHTML, /Main!A1/);
  assert.match(p.get('repairProgress').textContent, /恢复颜色 2 处/);
  assert.match(p.get('repairProgress').textContent, /Sheet 标签 1 张/);
  assert.ok(p.links.every(link => link.target.compare === 'colors'));
});

test('large color comparisons process and count all results while expanding the list in batches', async () => {
  const cells = Array.from({length:601}, (_, index) => ({sheet:'Main',addr:`A${index+1}`, mode:'all',
    fill:'#8ACDD1', old_fill:'none', status:'ok', reason:'已恢复'}));
  const p = await page({applied:{folder:'current', files:[{name:'book.xlsx', cells}]}}, {kind:'colors'});
  assert.match(p.get('repairProgress').textContent, /恢复颜色 601 处/);
  assert.equal((p.get('repairResults').innerHTML.match(/<tr><td>/g) || []).length, 200);
  assert.match(p.get('repairResults').innerHTML, /200 \/ 601/);
  p.get('repairResults').events.click({target:{closest:() => ({dataset:{repairMore:'book.xlsx'}})}});
  assert.equal((p.get('repairResults').innerHTML.match(/<tr><td>/g) || []).length, 400);
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
