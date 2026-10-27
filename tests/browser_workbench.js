// Real Chromium DOM/layout checks, without npm dependencies or the user's Excel service.
// node tests/browser_workbench.js [--baseline]
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const os = require('node:os');
const path = require('node:path');
const { spawn, execFileSync } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const baseline = process.argv.includes('--baseline');
const browser = [process.env.CHROME_BIN, process.env.EDGE_BIN,
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  'C:/Program Files/Google/Chrome/Application/chrome.exe', '/usr/bin/chromium', '/usr/bin/google-chrome']
  .find(candidate => candidate && fs.existsSync(candidate));
if (!browser) throw new Error('Set CHROME_BIN or EDGE_BIN to a Chromium browser executable.');
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const svg = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="600" height="160"><rect width="600" height="160" fill="#e4edf9"/><text x="30" y="80" font-size="28">Synthetic screenshot</text></svg>');
const fixture = { folder: 'C:/Synthetic', busy: true, logs: [], progress: {}, open_blocked: {},
  job: { kind: 'batch_scan', status: 'running', options: { scan_only: true }, completed: [], scanned: [],
    workbooks: Array.from({ length: 80 }, (_, i) => `C:/Synthetic/book-${i}.xlsx`) },
  cards: Array.from({ length: 1000 }, (_, i) => ({ id: String(i), seq: i,
    wb_path: `C:/Synthetic/book-${i % 80}.xlsx`, wb_name: `book-${i % 80}.xlsx`, sheet_name: 'Sheet1',
    cell_addr: `A${i + 1}`, excel_row: i + 1, color: 'yellow', finalized: true, verified: false,
    text: 'Synthetic review text '.repeat(2 + i % 5), images: { before: svg, after: svg } })) };
const pageSvg = 'data:image/svg+xml,' + encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="600" height="800">
  <rect width="600" height="800" fill="white"/><text x="30" y="50" font-size="24">Sheet1 - Page 1</text>
  ${Array.from({length:20}, (_, i) => `<path d="M30 ${90+i*32}H570" stroke="#aebdd0"/><text x="40" y="${112+i*32}" font-size="14">Row ${i+1} - sample cell content</text>`).join('')}
  <path d="M30 90V730M220 90V730M570 90V730" stroke="#aebdd0"/></svg>`);
for (const index of [0, 1]) fixture.cards[index].context = { image: pageSvg, page_width: 600, page_height: 800,
  page_index: 0, bbox: [30, 218, 220, 250] };
fixture.cards[1].images = { before: svg };
fixture.cards[1].after_kind = 'scan_only';
let script = baseline ? execFileSync('git', ['show', '940e58a:static/batch.js'], { cwd: root, encoding: 'utf8' })
  : fs.readFileSync(path.join(root, 'static/batch.js'), 'utf8');
script = script.replace('function cardHtml(card, large = false) {',
  'function cardHtml(card, large = false) { window.__htmlCalls = (window.__htmlCalls || 0) + 1;');
const server = http.createServer((request, response) => {
  const url = new URL(request.url, 'http://localhost');
  if (url.pathname === '/api/batch') {
    response.setHeader('Content-Type', 'application/json');
    response.setHeader('ETag', '"fixture"');
    if (request.headers['if-none-match'] === '"fixture"') { response.writeHead(304); response.end(); }
    else response.end(JSON.stringify(fixture));
  } else if (request.method === 'POST') {
    request.resume();
    response.setHeader('Content-Type', 'application/json');
    response.end('{"ok":true}');
  } else if (url.pathname === '/static/batch.js') {
    response.setHeader('Content-Type', 'application/javascript'); response.end(script);
  } else if (url.pathname.startsWith('/static/')) {
    const file = path.resolve(root, '.' + url.pathname);
    if (!file.startsWith(path.join(root, 'static') + path.sep) || !fs.existsSync(file)) {
      response.writeHead(404); response.end(); return;
    }
    response.setHeader('Content-Type', file.endsWith('.css') ? 'text/css' : 'application/javascript');
    response.end(fs.readFileSync(file));
  } else {
    response.setHeader('Content-Type', 'text/html; charset=utf-8');
    response.end(fs.readFileSync(path.join(root, 'static/batch.html'), 'utf8')
      .replace('<script src="/static/debug_log.js"></script>', ''));
  }
});

async function main() {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'workbench-browser-'));
  let child, socket;
  const pending = new Map(), errors = [];
  let nextId = 0;
  try {
    await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    child = spawn(browser, ['--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
      '--remote-debugging-port=0', `--user-data-dir=${profile}`, 'about:blank'], { windowsHide: true, stdio: 'ignore' });
    child.on('error', error => errors.push(error.message));
    const activePort = path.join(profile, 'DevToolsActivePort');
    for (let i = 0; !fs.existsSync(activePort) && i < 100; i++) await delay(100);
    if (!fs.existsSync(activePort)) throw new Error('Browser did not expose DevTools: ' + errors.join('; '));
    const port = fs.readFileSync(activePort, 'utf8').split('\n')[0];
    const target = await (await fetch(`http://127.0.0.1:${port}/json/new?about:blank`, { method: 'PUT' })).json();
    socket = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
    socket.onmessage = event => {
      const message = JSON.parse(event.data);
      if (message.id) {
        const entry = pending.get(message.id);
        if (!entry) return;
        pending.delete(message.id); clearTimeout(entry.timer);
        message.error ? entry.reject(new Error(JSON.stringify(message.error))) : entry.resolve(message.result);
      } else if (message.method === 'Runtime.exceptionThrown') errors.push(message.params.exceptionDetails.exception?.description || message.params.exceptionDetails.text);
    };
    const cdp = (method, params = {}) => new Promise((resolve, reject) => {
      const id = ++nextId;
      const timer = setTimeout(() => { pending.delete(id); reject(new Error('CDP timeout: ' + method)); }, 30000);
      pending.set(id, { resolve, reject, timer }); socket.send(JSON.stringify({ id, method, params }));
    });
    const evaluate = async expression => {
      const result = await cdp('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
      if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
      return result.result.value;
    };
    const pressKey = async (code, key, keyCode, location = 0, modifiers = 0) => {
      await cdp('Input.dispatchKeyEvent', { type: 'rawKeyDown', code, key, windowsVirtualKeyCode: keyCode, location, modifiers });
      await cdp('Input.dispatchKeyEvent', { type: 'keyUp', code, key, windowsVirtualKeyCode: keyCode, location });
    };
    const hover = async selector => {
      const point = await evaluate(`(async()=>{const el=document.querySelector(${JSON.stringify(selector)});
        el.scrollIntoView({block:'center'});await new Promise(resolve=>requestAnimationFrame(resolve));
        const r=el.getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2};})()`);
      await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', ...point });
    };
    const capture = async suffix => {
      if (!process.env.WORKBENCH_SCREENSHOT) return;
      const file = path.parse(process.env.WORKBENCH_SCREENSHOT);
      const screenshot = await cdp('Page.captureScreenshot', { format: 'png' });
      fs.writeFileSync(path.join(file.dir, `${file.name}-${suffix}.png`), Buffer.from(screenshot.data, 'base64'));
    };
    const assertMagnifierFits = async (selector = '#largeCard .batch-shots img') => {
      assert.equal(await evaluate(`(()=>{const image=$('imagePreview').querySelector('img'), r=image.getBoundingClientRect(),
        frame=$('imagePreview').getBoundingClientRect(), source=document.querySelector(${JSON.stringify(selector)}),
        sourceRect=source.getBoundingClientRect(), style=getComputedStyle(source),
        sourceScale=Math.min((source.clientWidth-parseFloat(style.paddingLeft)-parseFloat(style.paddingRight))/source.naturalWidth,
          (source.clientHeight-parseFloat(style.paddingTop)-parseFloat(style.paddingBottom))/source.naturalHeight),
        twiceWidth=source.naturalWidth*sourceScale*2, twiceHeight=source.naturalHeight*sourceScale*2;
        return !$('imagePreview').classList.contains('hidden') && frame.left>=8 && frame.top>=8
          && frame.right<=innerWidth-6 && frame.bottom<=innerHeight-6
          && Math.abs(r.width/r.height-image.naturalWidth/image.naturalHeight)<0.01
          && (twiceWidth<=innerWidth-64 && twiceHeight<=innerHeight-64
            ? Math.abs(r.width-twiceWidth)<1 && Math.abs(r.height-twiceHeight)<1
            : Math.abs(r.width-(innerWidth-64))<1 || Math.abs(r.height-(innerHeight-64))<1)
          && (Math.abs(frame.left+frame.width/2-sourceRect.left-sourceRect.width/2)<1
            || Math.abs(frame.left-8)<1 || Math.abs(frame.right-(innerWidth-12))<1)
          && (Math.abs(frame.top+frame.height/2-sourceRect.top-sourceRect.height/2)<1
            || Math.abs(frame.top-8)<1 || Math.abs(frame.bottom-(innerHeight-12))<1);})()`), true);
    };
    await cdp('Runtime.enable'); await cdp('Page.enable');
    await cdp('Emulation.setDeviceMetricsOverride', { width: 1280, height: 900, deviceScaleFactor: 1, mobile: false });
    await cdp('Page.navigate', { url: `http://127.0.0.1:${server.address().port}/workbench` });
    let ready = false;
    for (let i = 0; i < 100 && !ready; i++) {
      ready = await evaluate(`document.querySelectorAll('#cardList .batch-card').length === 1000`);
      if (!ready) await delay(100);
    }
    assert.equal(ready, true, errors.join('\n') || JSON.stringify(await evaluate(`({url:location.href,
      cards:document.querySelectorAll('#cardList .batch-card').length, body:document.body?.innerText.slice(0,800),
      error:document.getElementById('errorText')?.textContent})`)));
    await evaluate(`poll=async()=>{}; new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))`);
    assert.deepEqual(await evaluate(`(async()=>{const r=await fetch('/api/batch',{cache:'no-store',headers:{'If-None-Match':'"fixture"'}});return {status:r.status,body:await r.text()};})()`), { status: 304, body: '' });
    const report = { mode: baseline ? 'baseline' : 'optimized', browser: path.basename(browser), cards: 1000, files: 80 };
    report.progress = await evaluate(`(() => {
      const durations=[]; window.__htmlCalls=0;
      for(let i=0;i<12;i++){batchState.job.completed.push('progress-'+i);const t=performance.now();render();durations.push(performance.now()-t);}
      durations.sort((a,b)=>a-b); return {median_ms:durations[6],p95_ms:durations[11],html_calls:window.__htmlCalls};
    })()`);
    await evaluate(`window.__originalNodes=new Map([...document.querySelectorAll('#cardList .batch-card')].map(el=>[el.dataset.id,el]));
      window.__originalRoot=$('cardList').firstElementChild; window.__htmlCalls=0;`);
    report.append = await evaluate(`(() => {
      batchState.cards.push({...batchState.cards[0],id:'new-card',seq:1001});const t=performance.now();render();
      const duration=performance.now()-t, now=new Map([...document.querySelectorAll('#cardList .batch-card')].map(el=>[el.dataset.id,el]));
      return {ms:duration,html_calls:window.__htmlCalls,preserved_nodes:[...window.__originalNodes].filter(([id,el])=>now.get(id)===el).length,
        preserved_root:window.__originalRoot===$('cardList').firstElementChild};
    })()`);
    if (!baseline) {
      assert.equal(report.progress.html_calls, 0); assert.equal(report.append.html_calls, 1);
      assert.equal(report.append.preserved_nodes, 1000); assert.equal(report.append.preserved_root, true);
    }
    report.single_card = await evaluate(`(() => {
      window.__originalNodes=new Map([...document.querySelectorAll('#cardList .batch-card')].map(el=>[el.dataset.id,el]));
      window.__htmlCalls=0;batchState.cards[0].color='blue';const t=performance.now();render();const duration=performance.now()-t;
      const now=new Map([...document.querySelectorAll('#cardList .batch-card')].map(el=>[el.dataset.id,el]));
      return {ms:duration,html_calls:window.__htmlCalls,preserved_nodes:[...window.__originalNodes].filter(([id,el])=>now.get(id)===el).length};
    })()`);
    if (!baseline) assert.equal(report.single_card.html_calls, 1);
    assert.equal(report.single_card.preserved_nodes, 1000);
    report.file_block = await evaluate(`(() => {
      window.__htmlCalls=0;batchState.open_blocked[batchState.cards[1].wb_path]='processing';const t=performance.now();render();
      return {ms:performance.now()-t,html_calls:window.__htmlCalls,affected_cards:batchState.cards.filter(c=>c.wb_path===batchState.cards[1].wb_path).length};
    })()`);
    if (!baseline) {
      assert.equal(report.file_block.html_calls, report.file_block.affected_cards);
      await evaluate(`toggleSelected('0');batchState.cards[0].message='updated';render();`);
      assert.equal(await evaluate(`document.querySelector('#cardList [data-id="0"] .batch-select').getAttribute('aria-pressed')`), 'true');
      await evaluate(`batchState.job.pause_requested=true;render();`);
      assert.equal(await evaluate(`document.querySelector('#cardList [data-id="new-card"] [data-action="auto"]').title.includes('正在暂停')`), true);
      await evaluate(`batchState.job.pause_requested=false;batchState.job.to_manual=true;render();`);
      assert.equal(await evaluate(`document.querySelector('#cardList [data-id="new-card"] [data-action="auto"]').title.includes('转人工')`), true);
      await evaluate(`window.__beforeResize=document.querySelector('#cardList [data-id="0"]');`);
      await cdp('Emulation.setDeviceMetricsOverride', { width: 750, height: 900, deviceScaleFactor: 1, mobile: false });
      await delay(100);
      assert.equal(await evaluate(`window.__beforeResize === document.querySelector('#cardList [data-id="0"]')`), true);
      await evaluate(`document.querySelector('[data-filter="yellow"]').click();`);
      assert.equal(await evaluate(`document.querySelector('#cardList [data-id="0"]') === null && selectedIds.size === 0`), true);
      await evaluate(`document.querySelector('[data-filter="all"]').click();`);
      await cdp('Emulation.setDeviceMetricsOverride', { width: 1280, height: 900, deviceScaleFactor: 1, mobile: false });
      await delay(100);
      assert.equal(await evaluate(`document.querySelector('#largeBtn, [data-action="large"]') === null`), true);
      await hover('#cardList [data-id="0"] .batch-shots img');
      assert.equal(await evaluate(`!$('largeView').open && $('imagePreview').classList.contains('hidden')`), true);
      await pressKey('AltRight', 'Alt', 18, 2, 1);
      assert.equal(await evaluate(`$('largeView').open`), false);
      await evaluate(`window.__reviewScroll=scrollY;`);
      await pressKey('AltLeft', 'Alt', 18, 1, 1);
      assert.equal(await evaluate(`$('largeView').open && activeCardId==='0' && !$('cardList').classList.contains('hidden')`), true);
      await cdp('Input.dispatchKeyEvent', { type: 'rawKeyDown', code: 'AltLeft', key: 'Alt', windowsVirtualKeyCode: 18, location: 1, modifiers: 1, autoRepeat: true });
      await cdp('Input.dispatchKeyEvent', { type: 'keyUp', code: 'AltLeft', key: 'Alt', windowsVirtualKeyCode: 18, location: 1 });
      assert.equal(await evaluate(`$('largeView').open`), true);
      assert.equal(await evaluate(`(()=>{const shots=$('largeCard').querySelector('.batch-shots'), context=$('largeCard').querySelector('.batch-context');
        const figures=shots.querySelectorAll('figure');return shots.getBoundingClientRect().right<=context.getBoundingClientRect().left
          && figures[0].getBoundingClientRect().bottom<=figures[1].getBoundingClientRect().top
          && getComputedStyle(figures[0].querySelector('img')).borderTopColor==='rgb(220, 38, 38)'
          && getComputedStyle(figures[1].querySelector('img')).borderTopColor==='rgb(37, 99, 235)'
          && context.querySelector('img').getBoundingClientRect().height>500
          && context.querySelector('img').getBoundingClientRect().bottom<=$('largeView').getBoundingClientRect().bottom-12;})()`), true);
      await capture('comparison');
      for (const shot of ['first', 'last']) {
        const selector = `#largeCard .batch-shots figure:${shot}-of-type img`;
        await hover(selector);
        await evaluate(`$('imagePreview').querySelector('img').decode()`);
        await assertMagnifierFits(selector);
        const before = await evaluate(`(()=>{const r=$('imagePreview').getBoundingClientRect();
          return {left:r.left,top:r.top,width:r.width,height:r.height};})()`);
        const moved = await evaluate(`(()=>{const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();
          return {x:r.left+r.width*0.65,y:r.top+r.height*0.4};})()`);
        await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', ...moved });
        assert.deepEqual(await evaluate(`(()=>{const r=$('imagePreview').getBoundingClientRect();
          return {left:r.left,top:r.top,width:r.width,height:r.height};})()`), before);
      }
      await capture('comparison-magnifier');
      await evaluate(`window.__largeNode=$('largeCard').firstElementChild;batchState.job.completed.push('more-progress');render();`);
      assert.equal(await evaluate(`window.__largeNode === $('largeCard').firstElementChild`), true);
      await pressKey('AltLeft', 'Alt', 18, 1, 1);
      assert.equal(await evaluate(`!$('largeView').open && Math.abs(scrollY-window.__reviewScroll)<2`), true);
      await hover('#cardList [data-id="1"] .batch-shots img');
      await pressKey('AltLeft', 'Alt', 18, 1, 1);
      assert.equal(await evaluate(`activeCardId==='1' && $('largeCard').querySelectorAll('.batch-shots figure').length===1
        && !!$('largeCard').querySelector('.batch-context-marker')`), true);
      await capture('single');
      await hover('#largeCard .batch-shots img');
      await evaluate(`$('imagePreview').querySelector('img').decode()`);
      assert.equal(await evaluate(`!$('imagePreview').classList.contains('hidden') && $('imagePreview').closest('dialog')===$('largeView')`), true);
      await assertMagnifierFits();
      await capture('magnifier');
      // Native image size and contain-box whitespace must not affect the visual 2x magnification.
      for (const [width, height] of [[120, 40], [120, 600]]) {
        await evaluate(`(async()=>{const image=$('largeCard').querySelector('.batch-shots img');
          image.src='data:image/svg+xml,'+encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}"><rect width="100%" height="100%" fill="#e4edf9"/></svg>');
          await image.decode();})()`);
        await hover('#largeCard .batch-shots img');
        await evaluate(`$('imagePreview').querySelector('img').decode()`);
        await assertMagnifierFits();
      }
      await pressKey('Escape', 'Escape', 27);
      assert.equal(await evaluate(`!$('largeView').open && !largeMode && $('imagePreview').classList.contains('hidden')`), true);
      await hover('#cardList [data-id="0"] .batch-shots img');
      await pressKey('AltLeft', 'Alt', 18, 1, 1);
      await cdp('Emulation.setDeviceMetricsOverride', { width: 1100, height: 900, deviceScaleFactor: 1, mobile: false });
      await delay(100);
      await evaluate(`$('closeLargeBtn').click();`);
      assert.equal(await evaluate(`document.querySelector('#cardList [data-id="0"]') === window.__beforeResize`), true);
      assert.equal(await evaluate(`fileElements.values().next().value.columns.length === Math.max(1,Math.min(3,Math.floor(($('cardList').clientWidth-44+18)/328)))`), true);
      const accurateHeight = await evaluate(`(async()=>{
        const image=document.querySelector('#cardList [data-id="0"] img');
        image.src='data:image/svg+xml,'+encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="600" height="600"></svg>');
        await image.decode();await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
        return Math.abs(cardHeight('0')-cardElements.get('0').getBoundingClientRect().height-CARD_GAP)<1;
      })()`);
      assert.equal(accurateHeight, true);
      await evaluate(`document.querySelector('#cardList [data-id="0"] .excel-link').click();
        document.querySelector('[data-filter="yellow"]').click();document.querySelector('[data-filter="all"]').click();`);
      assert.equal(await evaluate(`document.querySelector('#cardList [data-id="0"] .excel-link').classList.contains('is-visited')`), true);
      const anchorDelta = await evaluate(`(async()=>{
        document.querySelector('.batch-file[data-file="C:/Synthetic/book-40.xlsx"]').scrollIntoView();
        await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
        const anchor=scrollAnchor();batchState.cards.push({...batchState.cards[0],id:'tall-card',text:'Long content '.repeat(200)});render();
        await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
        return Math.abs(cardElements.get(anchor.id).getBoundingClientRect().top-anchor.top);
      })()`);
      assert.ok(anchorDelta < 2, `scroll anchor drifted ${anchorDelta}px`);
      report.functional_checks = ['selection after replacement', 'pause/manual prompts', 'resize retains nodes',
        'filter selection', 'left/right Alt and key repeat', 'hover stays quiet in list', 'modal layout and colors',
        'Alt restores scroll position', 'single-image modal magnifier', '2x magnifier anchored to each comparison shot',
        'magnifier stays still as pointer moves', 'small and tall images preserve ratio and stay within viewport', 'Escape closes modal',
        'large-view node retention', 'resize while in large view', 'image-load height invalidation',
        'scroll anchor after insertion above viewport', 'native fetch 304', 'visited Excel links after filtering'];
    }
    assert.deepEqual(errors, []);
    if (process.env.WORKBENCH_SCREENSHOT) {
      const screenshot = await cdp('Page.captureScreenshot', { format: 'png' });
      fs.writeFileSync(process.env.WORKBENCH_SCREENSHOT, Buffer.from(screenshot.data, 'base64'));
    }
    console.log(JSON.stringify(report, null, 2));
  } finally {
    socket?.close();
    for (const entry of pending.values()) { clearTimeout(entry.timer); entry.reject(new Error('Browser closing')); }
    if (child && child.exitCode === null) { child.kill(); await delay(500); }
    await new Promise(resolve => server.close(resolve));
    // Only remove the uniquely created profile inside the known temporary directory.
    if (path.dirname(path.resolve(profile)) !== path.resolve(os.tmpdir())) throw new Error('Unexpected profile path');
    fs.rmSync(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
  }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
