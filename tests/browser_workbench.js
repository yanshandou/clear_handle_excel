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
    await cdp('Runtime.enable'); await cdp('Page.enable');
    await cdp('Emulation.setDeviceMetricsOverride', { width: 1280, height: 900, deviceScaleFactor: 1, mobile: false });
    await cdp('Page.navigate', { url: `http://127.0.0.1:${server.address().port}/workbench` });
    let ready = false;
    for (let i = 0; i < 100 && !ready; i++) {
      ready = await evaluate(`document.querySelectorAll('#cardList .batch-card').length === 1000`);
      if (!ready) await delay(100);
    }
    assert.equal(ready, true, errors.join('\n'));
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
      await evaluate(`document.querySelector('[data-filter="all"]').click();document.querySelector('#cardList [data-id="0"] [data-action="large"]').click();window.__largeNode=$('largeCard').firstElementChild;
        batchState.job.completed.push('more-progress');render();`);
      assert.equal(await evaluate(`window.__largeNode === $('largeCard').firstElementChild`), true);
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
        'filter selection', 'large-view node retention', 'resize while in large view', 'image-load height invalidation',
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
