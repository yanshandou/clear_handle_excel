const assert = require('node:assert/strict');
const test = require('node:test');
const { page } = require('./workbench_page');

test('appending and removing cards retain unchanged DOM nodes and file containers', () => {
  const p = page();
  p.run(`render(); const originalA = cardElements.get('a'), originalB = cardElements.get('b');
    const originalFile = fileElements.get('book.xlsx').element, originalRoot = $('cardList').firstElementChild;
    batchState.cards.push({...batchState.cards[0], id:'c', seq:3}); render();`);
  assert.equal(p.run(`cardElements.get('a') === originalA && cardElements.get('b') === originalB`), true);
  assert.equal(p.run(`fileElements.get('book.xlsx').element === originalFile && $('cardList').firstElementChild === originalRoot`), true);
  p.run(`batchState.cards = batchState.cards.filter(c => c.id !== 'a'); render();`);
  assert.equal(p.run(`cardElements.get('b') === originalB && !cardMemo.has('a') && !cardHeights.has('a')`), true);
});

test('empty-list messaging updates when a task starts or stops', () => {
  const p = page();
  p.run(`batchState.cards=[]; render();`);
  assert.match(p.elements.get('cardList').innerHTML, /当前没有待审核卡片/);
  p.run(`batchState.busy=true; render();`);
  assert.match(p.elements.get('cardList').innerHTML, /正在自动扫描与修复/);
});

test('memoization ignores logs, progress and unused job options but updates the affected file', () => {
  const p = page();
  p.run(`batchState.cards[0].wb_path='A.xlsx'; batchState.cards[1].wb_path='B.xlsx';
    batchState.busy=true; batchState.job={kind:'batch_scan', phase:'repair', options:{}, workbooks:['A.xlsx','B.xlsx']};
    render(); let calls=[]; const originalHtml=cardHtml; cardHtml=card=>{calls.push(card.id);return originalHtml(card);};
    batchState.logs.push('progress'); batchState.progress.completed=1; batchState.job.options.card_ids=['a','b']; render();`);
  assert.equal(p.run('calls.length'), 0);
  p.run(`batchState.open_blocked={'B.xlsx':'processing'}; render();`);
  assert.deepEqual(Array.from(p.run('calls')), ['b']);
});

test('pause and manual requests invalidate card prompts without other state changes', () => {
  const p = page();
  p.run(`batchState.busy=true; batchState.job={kind:'batch_scan',phase:'repair',options:{},workbooks:[]}; render();
    batchState.job.pause_requested=true; render();`);
  assert.match(p.run(`cardElements.get('b').outerHTML`), /正在暂停/);
  p.run(`batchState.job.pause_requested=false; batchState.job.to_manual=true; render();`);
  assert.match(p.run(`cardElements.get('b').outerHTML`), /转人工/);
});

test('selection survives card and file header replacement', () => {
  const p = page();
  p.run(`render(); toggleSelected('b'); batchState.cards[1].color='blue'; render();`);
  assert.equal(p.run(`cardElements.get('b').classList.contains('is-selected')`), true);
  assert.equal(p.run(`cardElements.get('b').querySelector('.batch-select').getAttribute('aria-pressed')`), 'true');
  p.run(`selectShown('all'); batchState.busy=true; render();`);
  assert.equal(p.run(`fileElements.get('book.xlsx').head.querySelector('[data-file-action="select"]').getAttribute('aria-pressed')`), 'true');
});

test('expanding a filter updates file selection even when all card nodes are reused', () => {
  const p = page();
  p.run(`render(); selectedFilter='blue'; filterKeep=new Set(); render(); selectShown('all');`);
  assert.equal(p.run(`fileElements.get('book.xlsx').head.querySelector('[data-file-action="select"]').getAttribute('aria-pressed')`), 'true');
  p.run(`selectedFilter='all'; filterKeep=new Set(); render();`);
  assert.equal(p.run(`fileElements.get('book.xlsx').head.querySelector('[data-file-action="select"]').getAttribute('aria-pressed')`), 'false');
});

test('image load and width changes refresh heights while preserving existing card nodes', () => {
  const p = page();
  p.run(`render(); const original=cardElements.get('a'); original.height=800;
    scheduleCardMeasurement({target:{closest:()=>original}});`);
  p.frames.shift()();
  assert.equal(p.run(`cardHeight('a')`), 818);
  p.run(`$('cardList').clientWidth=600; render();`);
  assert.equal(p.run(`cardElements.get('a') === original`), true);
  assert.equal(p.run(`fileElements.get('book.xlsx').columns.length`), 1);
  assert.equal(p.run(`cardHeightSum === [...cardHeights.values()].reduce((a,b)=>a+b,0)`), true);
});

test('cards returning from a filter refresh heights changed while detached', () => {
  const p = page();
  p.run(`render(); const hidden=cardElements.get('b'); selectedFilter='blue'; filterKeep=new Set(); render();
    hidden.height=900; selectedFilter='all'; filterKeep=new Set(); render();`);
  assert.equal(p.run(`cardHeight('b')`), 918);
});

test('large review does not rebuild its card for unrelated progress', () => {
  const p = page();
  p.run(`openLarge('a'); const revision=domRevision; batchState.logs.push('new log'); render();`);
  assert.equal(p.run(`revision === domRevision`), true);
});

test('cached card links notice changes to shared Excel link visit state', () => {
  const p = page();
  p.run(`render(); let calls=0; const original=cardHtml; cardHtml=card=>{calls++;return original(card);};
    excelLinkHtml.revision=1; render();`);
  assert.equal(p.run('calls'), 2);
});

test('ETag advances only with accepted data when a slower response is overtaken', async () => {
  const p = page();
  p.run(`lastEtag='v1'; let finishBody, count=0, validators=[];
    fetch=async (url, options)=>{validators.push(options.headers['If-None-Match']);
      if(options.headers['If-None-Match']==='v2') return {status:304};
      return {status:200,ok:true,headers:{get:()=> 'v2'},json:()=> ++count===1
        ? new Promise(resolve=>{finishBody=resolve;}) : Promise.resolve({...batchState,folder:'C:/Updated'})};};
    let first=pollUnderTest();`);
  await new Promise(resolve => setImmediate(resolve));
  await p.run('pollUnderTest()');
  p.run(`finishBody({...batchState,folder:'C:/Old'});`);
  await p.run('first');
  assert.equal(p.run('batchState.folder'), 'C:/Updated');
  assert.equal(p.run('lastEtag'), 'v2');
  assert.deepEqual(Array.from(p.run('validators')), ['v1', 'v1']);
  await p.run('pollUnderTest()');
  assert.equal(p.run('batchState.folder'), 'C:/Updated');
});

test('polls discarded during color saving cannot change ETag or state', async () => {
  const p = page();
  p.run(`lastEtag='old'; colorSaving=true; fetch=async()=>({status:200,ok:true,headers:{get:()=> 'new'},
    json:async()=>({...batchState,folder:'changed'})});`);
  await p.run('pollUnderTest()');
  assert.equal(p.run('lastEtag'), 'old');
  assert.equal(p.run('batchState.folder'), 'C:/Excel');
});

test('background polling does not accumulate requests while a poll is pending', async () => {
  const p = page();
  await new Promise(resolve => setImmediate(resolve));
  p.run(`let finishFetch, fetchCalls=0; fetch=()=>{fetchCalls++; return new Promise(resolve=>{finishFetch=resolve;});};
    let pendingPoll=pollUnderTest();`);
  await p.run('pollUnderTest(true)');
  assert.equal(p.run('fetchCalls'), 1);
  p.run(`finishFetch({status:304});`);
  await p.run('pendingPoll');
  assert.equal(p.run('pollInFlight'), 0);
});

test('a failed POST resets validators and restores controls from the next snapshot', async () => {
  const p = page();
  p.run(`lastEtag='old'; poll=pollUnderTest; let validators=[];
    fetch=async (url, options)=>options.method==='POST' ? {ok:false,json:async()=>({error:'rejected'})}
      : (validators.push(options.headers['If-None-Match']), {status:200,ok:true,headers:{get:()=> 'fresh'},json:async()=>batchState});`);
  await p.run(`action(()=>postUnderTest('/api/batch/mark',{id:'a',color:'blue'}))`);
  assert.equal(p.run('requestBusy'), false);
  assert.equal(p.elements.get('folderInput').disabled, false);
  assert.deepEqual(Array.from(p.run('validators')), [undefined]);
  assert.equal(p.elements.get('errorText').textContent, 'rejected');
});

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

test('cards without an after image show only one larger before image', () => {
  const p = page();
  p.run(`Object.assign(batchState.cards[1], {after_kind:'scan_only', images:{before:'/b.png'}})`);
  for (const large of [false, true]) {
    const html = p.run(`cardHtml(batchState.cards[1], ${large})`);
    assert.equal((html.match(/<figure>/g) || []).length, 1);
    assert.match(html, /batch-shots is-single/);
    assert.match(html, /<figcaption>修改前<\/figcaption>/);
    assert.doesNotMatch(html, /batch-vs|暂无截图/);
    assert.match(html, /人工修改后复检通过再显示/);
  }
  p.run(`Object.assign(batchState.cards[1], {after_kind:undefined, waiting:true})`);
  assert.match(p.run(`cardHtml(batchState.cards[1])`), /自动处理到这一格后显示/);
  p.run(`batchState.cards[1].images.after = '/a.png'`);
  const both = p.run(`cardHtml(batchState.cards[1])`);
  assert.equal((both.match(/<figure>/g) || []).length, 2);
  assert.doesNotMatch(both, /is-single/);
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
  assert.match(html, /人工修改后复检通过再显示/);
  p.run(`batchState.mode = {to_auto: true}`);
  await p.elements.get('toAutoBtn').listeners.click();
  assert.equal(p.requests.at(-1).url, '/api/batch/to_auto');
});

test('commit follows the selected filter color and names it on the button', async () => {
  const p = page();
  p.run(`batchState.cards.push({id:'g', seq:3, wb_name:'book.xlsx', sheet_name:'Sheet', cell_addr:'A8', excel_row:8, color:'green', images:{}});
    appConfirm = async text => { globalThis.asked = text; return true; }; selectedFilter = 'yellow'; renderControls();`);
  assert.equal(p.elements.get('commitBtn').textContent, '✓ 只处理黄色标记（1）');
  await p.elements.get('commitBtn').listeners.click();
  await new Promise(resolve => setImmediate(resolve));
  assert.match(p.run('globalThis.asked'), /只提交 1 张黄色卡片/);
  assert.equal(JSON.stringify(p.requests.at(-1).payload), '{"card_ids":["b"],"color":"yellow"}');
  p.run(`selectedFilter = 'all'; renderControls();`);
  assert.equal(p.elements.get('commitBtn').textContent, '✓ 确认并处理全部标记（3）');
});

test('a single-color commit leaves rows mixed with other colors for later', async () => {
  const p = page();
  p.run(`batchState.cards = [
      {id:'y1', seq:1, wb_path:'b', sheet_name:'S', excel_row:5, cell_addr:'A5', color:'yellow', row_height:{saved:true}},
      {id:'b1', seq:2, wb_path:'b', sheet_name:'S', excel_row:5, cell_addr:'B5', color:'blue', row_height:{saved:true}},
      {id:'y2', seq:3, wb_path:'b', sheet_name:'S', excel_row:6, cell_addr:'A6', color:'yellow', row_height:{saved:true}},
      {id:'y3', seq:4, wb_path:'b', sheet_name:'S', excel_row:7, cell_addr:'A7', color:'yellow'},
      {id:'b3', seq:5, wb_path:'b', sheet_name:'S', excel_row:7, cell_addr:'B7', color:'blue'}];
    appConfirm = async text => { globalThis.asked = text; return true; }; selectedFilter = 'yellow';`);
  await p.elements.get('commitBtn').listeners.click();
  await new Promise(resolve => setImmediate(resolve));
  assert.match(p.run('globalThis.asked'), /另有 1 张黄卡和别的颜色的卡共用行高或列宽/);
  assert.equal(JSON.stringify(p.requests.at(-1).payload.card_ids), '["y2","y3"]');
});

test('a single-color commit skips columns shared across different rows', () => {
  const p = page();
  p.run(`batchState.cards = [
    {id:'g', color:'green', wb_path:'b', sheet_name:'S', excel_row:5,
      col_width:{saved:true, ids:['B','C','D']}},
    {id:'b', color:'blue', wb_path:'b', sheet_name:'S', excel_row:9,
      col_width:{saved:true, ids:['C','D','E']}},
    {id:'free', color:'green', wb_path:'b', sheet_name:'S', excel_row:10,
      col_width:{saved:true, ids:['F']}}];`);
  assert.equal(p.run(`commitPlan('green').cards.map(c => c.id).join()`), 'free');
  assert.equal(p.run(`commitPlan('green').skipped.map(c => c.id).join()`), 'g');
  assert.equal(p.run(`commitPlan('all').cards.length`), 3);
});

test('column review follows transitive overlaps of merged areas', () => {
  const p = page();
  p.run(`batchState.cards = [
    {id:'a', color:'green', col_width:{saved:true, ids:['B','C']}},
    {id:'b', color:'green', col_width:{saved:true, ids:['C','D']}},
    {id:'c', color:'blue', col_width:{saved:true, ids:['D','E']}}];`);
  assert.equal(p.run(`commitPlan('green').cards.length`), 0);
  assert.equal(p.run(`commitPlan('green').skipped.map(c => c.id).join()`), 'a,b');
});

test('column and row height reviews on the same row are independent', () => {
  const p = page();
  p.run(`batchState.cards = [
    {id:'col', color:'blue', wb_path:'b', sheet_name:'S', excel_row:5,
      error_types:['hash_overflow'], col_width:{saved:true, ids:['C']}},
    {id:'row', color:'green', wb_path:'b', sheet_name:'S', excel_row:5, row_height:{saved:true}}];`);
  assert.equal(p.run(`commitPlan('green').cards.map(c => c.id).join()`), 'row');
  assert.equal(p.run(`commitPlan('green').skipped.length`), 0);
});

test('cards are grouped by file in processing order, and a file box rechecks that whole file', async () => {
  const p = page();
  p.run(`batchState.cards = [
      {id:'x1', seq:1, wb_path:'C:/Excel/x.xlsx', wb_name:'x.xlsx', sheet_name:'S', cell_addr:'A1', color:'yellow', finalized:true, images:{}},
      {id:'y1', seq:2, wb_path:'C:/Excel/sub/y.xlsx', wb_name:'y.xlsx', sheet_name:'S', cell_addr:'A1', color:'blue', images:{}},
      {id:'x2', seq:3, wb_path:'C:/Excel/x.xlsx', wb_name:'x.xlsx', sheet_name:'S', cell_addr:'A2', color:'yellow', finalized:true, images:{}},
      {id:'x3', seq:4, wb_path:'C:/Excel/x.xlsx', wb_name:'x.xlsx', sheet_name:'S', cell_addr:'A3', color:'green', images:{}}];`);
  assert.deepEqual(Array.from(p.run('shownCards().map(c => c.id)')), ['x1', 'x2', 'x3', 'y1']);
  p.run(`selectedFilter = 'blue'; filterKeep = new Set();`);
  assert.deepEqual(Array.from(p.run('shownCards().map(c => c.id)')), ['y1']);
  p.run(`selectedFilter = 'all'; render();`);
  const html = p.elements.get('cardList').innerHTML;
  assert.equal((html.match(/class="batch-file"/g) || []).length, 2);
  assert.ok(html.indexOf('x.xlsx') < html.indexOf('y.xlsx'));
  const head = p.run(`fileHeadHtml('C:/Excel/x.xlsx')`);
  assert.match(head, /黄 2/);
  assert.match(head, /绿 1/);
  assert.match(head, /共 3 格/);
  assert.match(p.run(`fileHeadHtml('C:/Excel/sub/y.xlsx')`), /batch-file-path">sub</);
  assert.match(p.run(`fileHeadHtml('C:/Excel/sub/y.xlsx')`), /disabled title="这个文件还没有已提交的黄色人工项"/);
  const button = { dataset: { fileAction: 'recheck', file: 'C:/Excel/x.xlsx' } };
  p.elements.get('cardList').listeners.click({ target: { closest(sel) { return sel === '[data-file-action]' ? button : null; } } });
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(p.requests.at(-1).url, '/api/batch/recheck');
  assert.equal(JSON.stringify(p.requests.at(-1).payload.card_ids), '["x1","x2"]');
  p.run(`toggleFileSelected('C:/Excel/x.xlsx')`);
  assert.deepEqual(Array.from(p.run('[...selectedIds]')).sort(), ['x1', 'x2', 'x3']);
  p.run(`toggleFileSelected('C:/Excel/x.xlsx')`);
  assert.equal(p.run('selectedIds.size'), 0);
});

test('auto adjust asks for the plan first, then starts only after confirming', async () => {
  const p = page();
  assert.match(p.run(`cardHtml(batchState.cards[1])`), /data-action="auto"/);
  assert.doesNotMatch(p.run(`cardHtml(batchState.cards[0])`), /data-action="auto"/);
  p.run(`post = async (url, payload) => { requests.push({url, payload});
    return url.endsWith('auto_plan') ? {ok:true, run:['b'], extra:[], skip:[]} : {ok:true}; }`);
  await p.run(`autoAdjust(['b'], '自动调整这一格')`);
  assert.deepEqual(p.requests.map(r => r.url), ['/api/batch/auto_plan', '/api/batch/auto_adjust']);
  assert.equal(JSON.stringify(p.requests[1].payload.card_ids), '["b"]');
});

test('auto adjust does not start when nothing can be adjusted', async () => {
  const p = page();
  p.run(`alerts = []; appAlert = async (m) => { alerts.push(m); };`);
  p.run(`post = async (url, payload) => { requests.push({url, payload});
    return {ok:true, run:[], extra:[], skip:[{id:'b', reason:'列宽不够'}]}; }`);
  await p.run(`autoAdjust(['b'], '自动调整这一格')`);
  assert.deepEqual(p.requests.map(r => r.url), ['/api/batch/auto_plan']);
  assert.match(p.run('alerts[0]'), /列宽不够/);
});

test('while auto adjust runs, other files can change color and queue auto adjust', () => {
  const p = page();
  p.run(`batchState.cards[0].wb_path = 'A.xlsx'; batchState.cards[1].wb_path = 'B.xlsx';
    batchState.busy = true; batchState.open_blocked = {'A.xlsx': '后台正在处理这份文件，处理完才能打开'};
    batchState.job = {kind: 'batch_scan', options: {picked: true}, workbooks: ['A.xlsx'], completed: []};`);
  assert.equal(p.run(`canMark(batchState.cards[0])`), false);
  assert.equal(p.run(`canMark(batchState.cards[1])`), true);
  assert.equal(p.run(`canAuto([batchState.cards[1]])`), true);
  assert.equal(p.run(`canAuto([batchState.cards[0], batchState.cards[1]])`), false);
  p.run(`batchState.job.closing = true`);
  assert.equal(p.run(`canAuto([batchState.cards[1]])`), false);
  // 提交等其他任务：照旧全部等它结束
  p.run(`batchState.job = {kind: 'batch_commit', options: {}, workbooks: ['A.xlsx']}`);
  assert.equal(p.run(`canMark(batchState.cards[1])`), false);
  p.run(`batchState.busy = false`);
  assert.equal(p.run(`canMark(batchState.cards[0]) && canAuto([batchState.cards[0]])`), true);
});

test('file box says which files are queued for auto adjust', () => {
  const p = page();
  p.run(`batchState.busy = true; batchState.job = {kind: 'batch_scan', options: {picked: true},
    workbooks: ['A.xlsx', 'B.xlsx', 'C.xlsx', 'D.xlsx'], completed: ['A.xlsx'], file_errors: {'B.xlsx': 'x'}, repairing: 'C.xlsx'};`);
  assert.match(p.run(`autoQueueBadge('A.xlsx')`), /自动调整完成/);
  assert.match(p.run(`autoQueueBadge('B.xlsx')`), /出错/);
  assert.match(p.run(`autoQueueBadge('C.xlsx')`), /正在自动调整/);
  assert.match(p.run(`autoQueueBadge('D.xlsx')`), /排队等自动调整 · 前面还有 1 个文件/);
  assert.equal(p.run(`autoQueueBadge('E.xlsx')`), '');
});

test('while scan only runs, scanned files can queue auto adjust until all files are scanned', () => {
  const p = page();
  p.run(`batchState.cards[0].wb_path = 'A.xlsx'; batchState.cards[1].wb_path = 'B.xlsx';
    batchState.cards[0].color = 'yellow'; batchState.cards[0].finalized = true;
    batchState.busy = true; batchState.open_blocked = {'A.xlsx': '后台正在处理这份文件，处理完才能打开'};
    batchState.job = {kind: 'batch_scan', options: {scan_only: true}, workbooks: ['B.xlsx', 'A.xlsx'], completed: ['B.xlsx']};`);
  assert.equal(p.run(`canAuto([batchState.cards[1]])`), true);
  assert.match(p.run(`autoButtonHtml(batchState.cards[1])`), /现在正在扫描：先排队，全部文件扫完后接着做/);
  // 还在扫的文件：按钮不可点，提示写明原因
  assert.equal(p.run(`canAuto([batchState.cards[0]])`), false);
  assert.match(p.run(`autoButtonHtml(batchState.cards[0])`), /disabled title="现在不能点：book.xlsx：后台正在处理这份文件，处理完再点自动调整"/);
  // 排上队的文件框写明扫完后开始
  p.run(`batchState.job.auto_queue = ['b']`);
  assert.match(p.run(`autoQueueBadge('B.xlsx')`), /排队等自动调整 · 全部文件扫完后开始/);
  assert.equal(p.run(`autoQueueBadge('A.xlsx')`), '');
  // 收尾、暂停、别的任务：都写明为什么不能点
  p.run(`batchState.job.closing = true`);
  assert.equal(p.run(`autoBlockReason([batchState.cards[1]])`), '只扫描正在收尾，结束后才能自动调整');
  p.run(`batchState.job.closing = false; batchState.job.pause_requested = true`);
  assert.equal(p.run(`autoBlockReason([batchState.cards[1]])`), '正在暂停，停下后才能自动调整');
  p.run(`batchState.job = {kind: 'batch_scan', options: {}, workbooks: ['A.xlsx'], completed: []}`);
  assert.equal(p.run(`autoBlockReason([batchState.cards[1]])`), '后台正在全自动处理，结束后才能自动调整');
  p.run(`batchState.job = {kind: 'batch_commit', options: {}, workbooks: ['A.xlsx'], completed: []}`);
  assert.equal(p.run(`autoBlockReason([batchState.cards[1]])`), '后台正在提交审核，结束后才能自动调整');
  assert.match(p.run(`fileHeadHtml('B.xlsx')`), /disabled title="现在不能点：后台正在提交审核，结束后才能自动调整"/);
});

test('recheck and commit queue behind other background tasks, except files the task still needs', () => {
  const p = page();
  p.run(`batchState.cards[0].wb_path = 'A.xlsx'; batchState.cards[1].wb_path = 'B.xlsx';
    batchState.cards[0].color = 'yellow'; batchState.cards[0].finalized = true;
    batchState.busy = true; batchState.open_blocked = {'A.xlsx': '后台正在处理这份文件，处理完才能打开'};
    batchState.job = {kind: 'batch_scan', options: {scan_only: true}, workbooks: ['B.xlsx', 'A.xlsx'], completed: ['B.xlsx']};`);
  // 扫完的文件：复检可以点，提示写明先排队
  assert.match(p.run(`fileHeadHtml('B.xlsx')`), /data-file-action="recheck" data-file="B.xlsx" title="[^"]*；后台正在只扫描：先排队，结束后接着复检"/);
  assert.match(p.run(`recheckButtonHtml(batchState.cards[1])`), /^<button[^>]*data-action="recheck" title=/);
  // 还在扫的文件：不能排，写明原因
  assert.match(p.run(`recheckButtonHtml(batchState.cards[0])`), /disabled title="现在不能点：book.xlsx：后台正在处理这份文件，处理完再点复检"/);
  assert.equal(p.run(`recheckAllCards().map(c => c.id).join()`), 'b');
  // 提交：还在扫的文件里的卡这次不交
  assert.equal(p.run(`commitPlan('all').cards.map(c => c.id).join()`), 'b');
  assert.equal(p.run(`commitPlan('all').held.map(c => c.id).join()`), 'a');
  assert.equal(p.run(`queueBlockReason([], '提交')`), '');
  assert.equal(p.run(`queueNote('提交')`), '；后台正在只扫描：先排队，结束后接着提交');
  // 复检正在跑：并进去，不用排到后面
  p.run(`batchState.job = {kind: 'batch_recheck', options: {}, workbooks: ['A.xlsx'], completed: []}`);
  assert.equal(p.run(`recheckBlockReason([batchState.cards[0]])`), '');
  assert.match(p.run(`queueNote('复检')`), /复检进行中/);
  // 提交 / 清除标色等：做完前所有文件都排着
  p.run(`batchState.job = {kind: 'batch_commit', options: {}, completed: []}; batchState.open_blocked = {'A.xlsx': '后台任务正在处理文件，结束后才能打开', 'B.xlsx': '后台任务正在处理文件，结束后才能打开'};`);
  assert.equal(p.run(`recheckBlockReason([batchState.cards[0], batchState.cards[1]])`), '');
  assert.equal(p.run(`commitPlan('all').held.length`), 0);
  // 其他页面的任务（精简 Sheet 等）：不排，写明原因
  p.run(`batchState.job = {kind: 'clean_sheets', options: {}}`);
  assert.equal(p.run(`queueBlockReason([], '提交')`), '后台正在处理别的任务，结束后才能提交');
  assert.match(p.run(`recheckButtonHtml(batchState.cards[1])`), /disabled title="现在不能点：后台正在处理别的任务，结束后才能复检"/);
});

test('queued rechecks and commits are listed with start-now and cancel', () => {
  const p = page();
  p.run(`batchState.cards[0].wb_path = 'A.xlsx'; batchState.cards[1].wb_path = 'B.xlsx'; batchState.busy = true;
    batchState.job = {kind: 'batch_scan', options: {scan_only: true}, workbooks: ['A.xlsx', 'B.xlsx'], completed: []};
    batchState.followups = [{kind: 'batch_recheck', card_ids: ['b'], full_paths: ['B.xlsx']}, {kind: 'batch_commit', card_ids: ['a', 'b'], color: 'blue'}];
    renderFollowups();`);
  assert.equal(p.elements.get('followupText').textContent,
    '⏳ 排队中：复检 1 个文件（1 格） → 提交蓝色 2 张卡。后台正在只扫描，结束后按顺序接着做。');
  assert.match(p.run(`followupBadge('B.xlsx')`), /排队等复检、提交 · 当前任务结束后开始/);
  assert.match(p.run(`followupBadge('A.xlsx')`), /排队等提交/);
  p.run(`batchState.busy = false; renderFollowups();`);
  assert.match(p.elements.get('followupText').textContent, /上一个任务暂停了，这些还没做/);
});

test('scan shows elapsed and remaining time, the current file, and files skipped for being slow', () => {
  const p = page();
  assert.equal(p.run(`duration(42)`), '42 秒');
  assert.equal(p.run(`duration(200)`), '3 分 20 秒');
  assert.equal(p.run(`duration(3720)`), '1 小时 2 分');
  assert.equal(p.run(`etaText(null)`), '');
  assert.equal(p.run(`etaText({elapsed: 30, remaining: null})`), ' · 已用 30 秒 · 扫完第一个文件后估算剩余时间');
  assert.match(p.run(`etaText({elapsed: 600, remaining: 480})`), /^ · 已用 10 分 · 预计还要约 8 分（\d\d:\d\d 左右扫完）$/);
  p.run(`renderScanNow({name: 'a.xlsx', sheet: 'レイアウト', index: 2, total: 5, seconds: 75, limit: 600})`);
  assert.equal(p.elements.get('scanNow').textContent,
    '正在检查 a.xlsx：第 2 / 5 张表「レイアウト」 · 这份已用 1 分 15 秒（上限 10 分，超过就跳过、留给人工）');
  p.run(`batchState.slow_files = [{path: 'C:/Excel/a.xlsx', name: 'a.xlsx', sheet: 'レイアウト', limit: 600}]; renderSlowFiles();`);
  assert.equal(p.elements.get('slowText').innerHTML,
    '⏱ 1 个文件扫描超时，已跳过，请人工检查：<a>a.xlsx</a>（超过 10 分，卡在表「レイアウト」）');
});

test('open all during a background task opens finished files and lists the rest', async () => {
  const p = page();
  p.run(`batchState.busy = true; batchState.job = {kind: 'batch_scan', options: {scan_only: true}, workbooks: [], completed: []};
    alerts = []; appAlert = async (m) => { alerts.push(m); };
    post = async (url, payload) => { requests.push({url, payload}); return {ok: true, opened: ['b.xlsx'],
      waiting: [{name: 'a.xlsx', reason: '还没轮到这份文件，处理完才能打开'}], errors: []}; };`);
  await p.run(`openReady()`);
  assert.deepEqual(p.requests.map(r => r.url), ['/api/batch/open']);
  assert.match(p.run('alerts[0]'), /已打开 1 个文件：b\.xlsx/);
  assert.match(p.run('alerts[0]'), /· a\.xlsx：还没轮到这份文件，处理完才能打开/);
});

test('fix error types render all 6 types, disable the 1 manual-only one, and pass selected types to jobs', async () => {
  const p = page();
  const html = p.elements.get('fixErrorTypes').innerHTML;
  assert.equal((html.match(/data-error-type="/g) || []).length, 5);
  assert.equal((html.match(/is-manual-only/g) || []).length, 1);
  assert.match(html, /### 列宽不足/);
  assert.match(html, /PDF 未匹配到文字（固定转人工）/);

  p.elements.get('folderInput').value = 'C:/Excel';
  p.run(`fixErrorTypes = new Set(['border_clip', 'page_break_cut'])`);
  await p.elements.get('startBtn').listeners.click();
  assert.equal(p.requests.at(-1).url, '/api/batch/start');
  assert.equal(p.requests.at(-1).payload.fix_error_types.join(','), 'border_clip,page_break_cut');

  // 全部取消勾选时点「开始全自动处理」提示至少勾选一种，不发请求
  const beforeCount = p.requests.length;
  p.run(`fixErrorTypes = new Set()`);
  await p.elements.get('startBtn').listeners.click();
  assert.equal(p.requests.length, beforeCount);
  assert.match(p.elements.get('errorText').textContent, /至少勾选一种/);
});

