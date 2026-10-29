const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

async function page(initial = {}, {kind = 'links', storage = new Map(), postError = '', failPoll = false} = {}) {
  const elements = {}, requests = [], timers = [], links = [];
  let status = initial;
  const get = id => elements[id] ||= {value:'', textContent:'', innerHTML:'', disabled:true, checked:true, style:{},
    events:{}, addEventListener(name, handler) {this.events[name] = handler;},
    querySelector() { return {focus() {}}; }};
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
  const selectFilter = repairFilter => get('repairFilters').events.click({target:{closest:() => ({dataset:{repairFilter}})}});
  return {get, requests, storage, timers, links, selectFilter, setStatus:value => {status = value;}, flush};
}

const hit = {sheet:'Main', addr:'A3', anchor:'A3', formula:'=#REF!*2', error:'引用错误'};
const result = {folder:'D:/current/001', old_folder:'D:/backup', files:[{name:'book.xlsx', old:'D:/backup/001/book.xlsx', cells:[{
  ...hit, action:'restore', status:'ok', source:'Main!A3', source_formula:'=Old!A1*2', value:20, display:'20', reason:'已恢复',
}, {...hit, addr:'A4', anchor:'A4', action:'skip', reason:'备份本身有错误'}]}]};

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

test('link result categories filter cells and retain file failures', async () => {
  const p = await page({applied:{folder:'current', files:[
    {name:'mixed.xlsx', cells:[
      {...hit, status:'ok', reason:'已恢复备份计算结果为固定值'},
      {...hit, addr:'A4', reason:'备份工作表无法唯一配对，请人工检查'},
      {...hit, addr:'A5', reason:'附近的内容和备份不同，可能移动过行列，请人工检查'},
    ]},
    {name:'failed.xlsx', error:'无法保存', cells:[{...hit, addr:'B1', status:'ok'}]},
    {name:'clean.xlsx', cells:[]},
  ]}});
  assert.equal(p.get('repairFilters').hidden, false);
  assert.match(p.get('repairFilters').innerHTML, /工作表无法配对（1 处）/);
  assert.match(p.get('repairFilters').innerHTML, /文件处理失败（1 份）/);
  p.selectFilter('link_sheet');
  assert.match(p.get('repairResults').innerHTML, /Main!A4/);
  assert.doesNotMatch(p.get('repairResults').innerHTML, /Main!A3|Main!A5|failed.xlsx/);
  p.selectFilter('unresolved');
  assert.match(p.get('repairResults').innerHTML, /failed.xlsx/);
  assert.doesNotMatch(p.get('repairResults').innerHTML, /Main!A3|clean.xlsx/);
  p.selectFilter('link_file_error');
  assert.match(p.get('repairResults').innerHTML, /Main!B1/);
  assert.match(p.get('repairResults').innerHTML, /无法保存/);
  p.selectFilter('no_difference');
  assert.match(p.get('repairResults').innerHTML, /clean.xlsx/);
  assert.doesNotMatch(p.get('repairResults').innerHTML, /mixed.xlsx/);
});

test('file titles, backup paths and cell locations all open the same server-paired comparison', async () => {
  for (const kind of ['links', 'colors', 'hidden']) {
    const p = await page({applied:result}, {kind});
    assert.equal(p.links.length, kind === 'colors' ? 4 : 5);
    assert.ok(p.links.every(link => link.target.compare === kind && link.target.file === 'book.xlsx'));
    assert.equal(p.links[0].target.cell, kind === 'colors' ? 'A4' : 'A3');
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
  assert.match(p.get('repairResults').innerHTML, /Main!B8/);
  assert.match(p.get('repairResults').innerHTML, /内容不一致/);
  assert.doesNotMatch(p.get('repairResults').innerHTML, /Main!B7|已恢复/);
  assert.match(p.get('repairProgress').textContent, /恢复颜色 1 处；未修复 1 处/);
});

test('font-only differences show both font values in preview and unresolved results', async () => {
  for (const phase of ['plan', 'applied']) {
    const p = await page({[phase]:{folder:'current', files:[{name:'book.xlsx', cells:[{
      sheet:'Main', addr:'A1', fill:'none', font:'#FF0000', old_fill:'none', old_font:'#123456',
      status:phase === 'plan' ? 'ok' : 'failed', reason:'需要核对字色',
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
    sheet:'Main', addr:'A1', fill:'other', font:'#112233', old_fill:'other', old_font:'#445566', status:'failed',
  }]}]}}, {kind:'colors'});
  assert.match(p.get('repairResults').innerHTML, /<td>主题色 \/ 图案 \/ 字色 #112233<\/td><td>主题色 \/ 图案 \/ 字色 #445566<\/td>/);
});

test('unknown fonts never produce a fabricated black value or an Aa preview', async () => {
  for (const font of [undefined, null, '']) {
    const p = await page({applied:{folder:'current', files:[{name:'book.xlsx', cells:
      ['none', 'other', '#AABBCC'].map((fill, index) => ({sheet:'Main', addr:`A${index+1}`,
        fill, font, old_fill:fill, old_font:font, status:'failed'})),
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
      fill, font:'#112233', old_fill:fill, old_font:'#445566', status:'failed'})),
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
      fill, font, old_fill:fill, old_font:cssFont, status:'failed'})),
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

test('color results hide successful cells and tabs while retaining unresolved locations and totals', async () => {
  const p = await page({applied:{folder:'current', files:[{name:'book.xlsx', old:'backup/book.xlsx', cells:[
    {sheet:'Done', kind:'tab', addr:'A1', anchor:'A1', fill:'#FFC000', old_fill:'none', status:'ok', code:'cleared_tab', reason:'已清除标签色'},
    {sheet:'Done', addr:'A1', fill:'#C0007A', old_fill:'none', status:'resolved', code:'restored_cell', reason:'已恢复单元格'},
    {sheet:'Main', old_sheet:'OldMain', kind:'tab', addr:'A1', anchor:'A1', fill:'#FFC000', old_fill:'none', status:'failed', code:'tab_verify_failed', reason:'标签颜色恢复失败'},
    {sheet:'Main', addr:'A1', fill:'#C0007A', old_fill:'none', action:'skip', code:'content_mismatch', reason:'内容不同'},
  ]}]}}, {kind:'colors'});
  const html = p.get('repairResults').innerHTML;
  assert.match(html, /Main · Sheet 标签/);
  assert.match(html, /OldMain · Sheet 标签/);
  assert.match(html, /无标签颜色/);
  assert.match(html, /Main!A1/);
  assert.match(html, /book.xlsx · 未修复 2 处/);
  assert.doesNotMatch(html, /Done|已清除标签色|已恢复单元格|link-ok/);
  for (const label of ['未处理好（1 份）', '全部（1 份）', '标签颜色恢复校验失败（1 份）', '单元格内容或公式不一致（1 处）', 'Sheet 标签标记色已清除（1 处）', '单元格配色已恢复（1 处）']) {
    assert.ok(p.get('repairFilters').innerHTML.includes(label));
  }
  assert.match(p.get('repairProgress').textContent, /恢复颜色 2 处；未修复 2 处/);
  assert.match(p.get('repairProgress').textContent, /Sheet 标签 1 张/);
  assert.ok(p.links.every(link => link.target.compare === 'colors'));
  assert.equal(p.links[0].target.cell, 'A1');
  p.selectFilter('cleared_tab');
  const done = p.get('repairResults').innerHTML;
  assert.match(done, /Done · Sheet 标签/);
  assert.match(done, /已清除标签色/);
  assert.doesNotMatch(done, /Main!A1|内容不同/);
});

test('color results omit empty and fully repaired files and show a clear empty state', async () => {
  for (const files of [[], [{name:'empty.xlsx', cells:[]}], [
    {name:'fixed.xlsx', cells:[{...hit, status:'ok'}, {...hit, status:'resolved'}]},
    {name:'empty.xlsx', cells:[]},
  ]]) {
    const p = await page({applied:{folder:'current', files}}, {kind:'colors'});
    assert.equal(p.get('repairResults').innerHTML, '<p>没有需要人工处理的配色问题。</p>');
    assert.equal(p.links.length, 0);
    assert.match(p.get('repairProgress').textContent, /未修复 0 处；0 份文件处理失败/);
  }
  const q = await page({applied:{folder:'current', files:[
    {name:'fixed.xlsx', cells:[{...hit, status:'ok'}, {...hit, status:'resolved'}]},
    {name:'empty.xlsx', cells:[]},
  ]}}, {kind:'colors'});
  for (const label of ['未处理好（0 份）', '全部（2 份）', '配色无差异（1 份）', '单元格配色已恢复（2 处）']) {
    assert.ok(q.get('repairFilters').innerHTML.includes(label));
  }
  q.selectFilter('no_difference');
  assert.match(q.get('repairResults').innerHTML, /empty.xlsx · 无需处理/);
  assert.doesNotMatch(q.get('repairResults').innerHTML, /fixed.xlsx|<table/);
  q.selectFilter('restored_cell');
  assert.match(q.get('repairResults').innerHTML, /fixed.xlsx · 2 处/);
  assert.match(q.get('repairResults').innerHTML, /Main!A3/);
  assert.doesNotMatch(q.get('repairResults').innerHTML, /empty.xlsx/);
});

test('color results retain file errors even with no cells or apparently successful cells', async () => {
  const p = await page({applied:{folder:'current', files:[
    {name:'empty.xlsx', cells:[]},
    {name:'fixed.xlsx', cells:[{...hit, status:'ok'}]},
    {name:'missing.xlsx', cells:[], error:'找不到备份', code:'backup_not_unique'},
    {name:'unsaved.xlsx', cells:[{...hit, status:'ok'}], error:'无法保存', code:'file_locked'},
  ]}}, {kind:'colors'});
  const html = p.get('repairResults').innerHTML;
  assert.doesNotMatch(html, /empty.xlsx|fixed.xlsx|没有需要处理的问题/);
  assert.match(html, /missing.xlsx · 处理失败/);
  assert.match(html, /找不到备份/);
  assert.match(html, /unsaved.xlsx · 1 处 · 处理失败/);
  assert.match(html, /未修复：无法保存/);
  for (const label of ['备份文件无法唯一匹配（1 份）', '文件被占用（1 份）']) {
    assert.ok(p.get('repairFilters').innerHTML.includes(label));
  }
  assert.match(p.get('repairProgress').textContent, /恢复颜色 1 处；未修复 1 处；2 份文件处理失败/);
  p.selectFilter('file_locked');
  const locked = p.get('repairResults').innerHTML;
  assert.match(locked, /unsaved.xlsx · 处理失败/);
  assert.match(locked, /无法保存/);
  assert.doesNotMatch(locked, /missing.xlsx|<table/);
});

test('color results keep incomplete scan warnings after all returned cells were repaired', async () => {
  const p = await page({applied:{folder:'current', files:[{
    name:'truncated.xlsx', truncated:true, cells:[{...hit, status:'ok'}],
  }]}}, {kind:'colors'});
  const html = p.get('repairResults').innerHTML;
  assert.match(html, /<details class="link-file" open>/);
  assert.match(html, /truncated.xlsx/);
  assert.match(html, /只检查前 500 处标记/);
  assert.doesNotMatch(html, /Main!A3| · 0 处|没有需要人工处理/);
  assert.match(p.get('repairFilters').innerHTML, /未处理好（1 份）/);
  assert.match(p.get('repairFilters').innerHTML, /单元格配色已恢复（1 处）/);
  assert.doesNotMatch(p.get('repairFilters').innerHTML, /配色无差异/);
});

test('color scan and preview keep pending cells and hits but omit files without issues', async () => {
  for (const phase of ['result', 'plan']) {
    const key = phase === 'result' ? 'hits' : 'cells';
    const p = await page({[phase]:{folder:'current', files:[
      {name:'pending.xlsx', [key]:[{...hit, action:'restore', status:'ok'}]},
      {name:'empty.xlsx', [key]:[]},
      {name:'failed.xlsx', [key]:[], error:'读取失败'},
    ]}}, {kind:'colors'});
    const html = p.get('repairResults').innerHTML;
    assert.match(html, /pending.xlsx · 1 处/);
    assert.match(html, /Main!A3/);
    assert.match(html, /读取失败/);
    assert.match(html, /empty.xlsx · 无需处理/);
    assert.equal(p.get('repairFilters').hidden, true);
    assert.match(p.get('repairProgress').textContent, /上次检查尚未修复/);
  }
});

test('large color results paginate only unresolved cells without changing full result totals', async () => {
  const cells = [
    ...Array.from({length:250}, (_, index) => ({sheet:'Done', addr:`A${index+1}`, status:'ok'})),
    ...Array.from({length:601}, (_, index) => ({sheet:'Main', addr:`A${index+1}`, mode:'all',
      fill:'#8ACDD1', old_fill:'none', status:'failed', reason:'未修复'})),
  ];
  const applied = {folder:'current', files:[{name:'book.xlsx', cells}]};
  const original = JSON.stringify(applied);
  const p = await page({applied}, {kind:'colors'});
  for (const shown of [200, 400, 600, 601]) {
    const html = p.get('repairResults').innerHTML;
    assert.match(p.get('repairProgress').textContent, /恢复颜色 250 处；未修复 601 处/);
    assert.equal((html.match(/<tr><td>/g) || []).length, shown);
    assert.match(html, /book.xlsx · 未修复 601 处/);
    assert.doesNotMatch(html, /Done|全部差异已处理/);
    if (shown < 601) {
      assert.ok(html.includes(`当前显示未修复项 ${shown} / 601 处`));
      p.get('repairResults').events.click({target:{closest:() => ({dataset:{repairMore:'book.xlsx'}})}});
    } else {
      assert.doesNotMatch(html, /data-repair-more/);
    }
  }
  assert.equal(JSON.stringify(applied), original);
});

test('color category chips list only reasons present in this run and show just that category', async () => {
  const applied = {folder:'current', files:[
    {name:'mixed.xlsx', old:'backup/mixed.xlsx', cells:[
      {...hit, status:'ok', code:'restored_cell'},
      {...hit, addr:'A4', anchor:'A4', action:'skip', code:'content_mismatch', reason:'内容不同'}]},
    {name:'fixed.xlsx', cells:[{...hit, status:'resolved', code:'restored_cell'}]},
    {name:'empty.xlsx', cells:[]},
    {name:'failed.xlsx', cells:[], error:'无法读取', code:'file_locked'},
  ]};
  const original = JSON.stringify(applied);
  const p = await page({applied}, {kind:'colors'});
  for (const label of ['未处理好（2 份）', '全部（4 份）', '单元格配色已恢复（2 处）', '单元格内容或公式不一致（1 处）', '文件被占用（1 份）', '配色无差异（1 份）']) {
    assert.ok(p.get('repairFilters').innerHTML.includes(label));
  }
  assert.doesNotMatch(p.get('repairFilters').innerHTML, /渐变|模板|合并/);
  const summary = p.get('repairProgress').textContent;
  for (const [filter, visible, absent] of [
    ['unresolved', /mixed.xlsx|failed.xlsx/, /fixed.xlsx|empty.xlsx|Main!A3/],
    ['restored_cell', /mixed.xlsx|fixed.xlsx/, /failed.xlsx|empty.xlsx|Main!A4|内容不同/],
    ['content_mismatch', /mixed.xlsx/, /fixed.xlsx|failed.xlsx|empty.xlsx|Main!A3/],
    ['file_locked', /failed.xlsx · 处理失败/, /mixed.xlsx|fixed.xlsx|empty.xlsx|<table/],
    ['no_difference', /empty.xlsx · 无需处理/, /mixed.xlsx|fixed.xlsx|failed.xlsx|<table/],
    ['all', /mixed.xlsx/, /没有检查结果/],
  ]) {
    p.selectFilter(filter);
    const html = p.get('repairResults').innerHTML;
    assert.match(html, visible);
    assert.doesNotMatch(html, absent);
    assert.match(p.get('repairFilters').innerHTML, new RegExp(`data-repair-filter="${filter}" aria-pressed="true"`));
    assert.equal(p.get('repairProgress').textContent, summary);
  }
  const all = p.get('repairResults').innerHTML;
  for (const file of applied.files) assert.ok(all.includes(file.name));
  assert.match(all, /Main!A3/);
  assert.match(all, /Main!A4/);
  assert.equal(p.requests.filter(request => request.options?.method === 'POST').length, 0);
  assert.equal(JSON.stringify(applied), original);
});

test('runs without any category keep both summary chips and starting a new run resets the filter', async () => {
  const p = await page({applied:{folder:'current', old_folder:'backup', files:[]}}, {kind:'colors'});
  assert.match(p.get('repairFilters').innerHTML, /未处理好（0 份）/);
  assert.match(p.get('repairFilters').innerHTML, /全部（0 份）/);
  for (const [filter, message] of [['unresolved', '没有需要人工处理的配色问题。'], ['all', '没有检查结果。']]) {
    p.selectFilter(filter);
    assert.equal(p.get('repairResults').innerHTML, `<p>${message}</p>`);
  }
  p.selectFilter('all');
  await p.get('repairRun').events.click();
  assert.match(p.get('repairFilters').innerHTML, /data-repair-filter="unresolved" aria-pressed="true"/);
  assert.equal(p.get('repairResults').innerHTML, '<p>没有需要人工处理的配色问题。</p>');
});

test('color filters are hidden before results, while running and in legacy preview phases', async () => {
  for (const initial of [{}, {running:true, phase:'repair', stage:'apply'}, {plan:{folder:'current', files:[{name:'book.xlsx', cells:[hit]}]}}]) {
    const p = await page(initial, {kind:'colors'});
    assert.equal(p.get('repairFilters').hidden, true);
  }
  for (const kind of ['links', 'hidden']) {
    const p = await page({applied:result}, {kind});
    assert.equal(typeof p.get('repairFilters').events.click, kind === 'links' ? 'function' : 'undefined');
    assert.match(p.get('repairResults').innerHTML, /Main!A3/);
    assert.match(p.get('repairResults').innerHTML, /Main!A4/);
  }
});

test('color category switches reset pagination for the newly selected cells', async () => {
  const cells = Array.from({length:450}, (_, index) => ({...hit, addr:`A${index+1}`,
    status:index < 225 ? 'ok' : 'failed', code:index < 225 ? 'restored_cell' : 'gradient'}));
  const p = await page({applied:{folder:'current', files:[{name:'large.xlsx', cells}]}}, {kind:'colors'});
  p.get('repairResults').events.click({target:{closest:() => ({dataset:{repairMore:'large.xlsx'}})}});
  assert.equal((p.get('repairResults').innerHTML.match(/<tr><td>/g) || []).length, 225);
  for (const filter of ['gradient', 'restored_cell', 'all']) {
    p.selectFilter(filter);
    assert.equal((p.get('repairResults').innerHTML.match(/<tr><td>/g) || []).length, 200);
  }
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
