const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn } = require('node:child_process');

const base = process.env.GARDEN_BASE_URL || 'http://127.0.0.1:8765';
const browser = [process.env.CHROME_BIN, process.env.EDGE_BIN,
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  'C:/Program Files/Google/Chrome/Application/chrome.exe', '/usr/bin/chromium', '/usr/bin/google-chrome']
  .find(file => file && fs.existsSync(file));
if (!browser) throw new Error('Set CHROME_BIN or EDGE_BIN to a Chromium executable.');
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

async function main() {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'garden-ui-browser-'));
  const screenshots = fs.mkdtempSync(path.join(os.tmpdir(), 'garden-ui-shots-'));
  const pending = new Map();
  const errors = [], failedRequests = [];
  let child, socket, nextId = 0;
  try {
    child = spawn(browser, ['--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
      '--remote-debugging-port=0', `--user-data-dir=${profile}`, 'about:blank'], { windowsHide: true, stdio: 'ignore' });
    child.on('error', error => errors.push(error.message));
    const portFile = path.join(profile, 'DevToolsActivePort');
    for (let i = 0; !fs.existsSync(portFile) && i < 100; i++) await delay(100);
    assert.ok(fs.existsSync(portFile), 'Chromium starts');
    const port = fs.readFileSync(portFile, 'utf8').split('\n')[0];
    const target = await (await fetch(`http://127.0.0.1:${port}/json/new?about:blank`, { method: 'PUT' })).json();
    socket = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
    socket.onmessage = event => {
      const message = JSON.parse(event.data);
      if (message.id) {
        const entry = pending.get(message.id);
        if (!entry) return;
        pending.delete(message.id);
        clearTimeout(entry.timer);
        message.error ? entry.reject(new Error(JSON.stringify(message.error))) : entry.resolve(message.result);
      } else if (message.method === 'Runtime.exceptionThrown') {
        errors.push(message.params.exceptionDetails.exception?.description || message.params.exceptionDetails.text);
      } else if (message.method === 'Network.responseReceived' && message.params.response.status >= 400) {
        failedRequests.push(`${message.params.response.status} ${message.params.response.url}`);
      }
    };
    const cdp = (method, params = {}) => new Promise((resolve, reject) => {
      const id = ++nextId;
      const timer = setTimeout(() => { pending.delete(id); reject(new Error(`Timeout: ${method}`)); }, 20000);
      pending.set(id, { resolve, reject, timer });
      socket.send(JSON.stringify({ id, method, params }));
    });
    const evaluate = async expression => {
      const result = await cdp('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
      if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
      return result.result.value;
    };
    const until = async expression => {
      for (let i = 0; i < 100; i++) {
        if (await evaluate(expression)) return;
        await delay(50);
      }
      throw new Error(`Condition did not become true: ${expression}`);
    };
    const navigate = async route => {
      await cdp('Page.navigate', { url: base + route });
      await until(`location.pathname + location.search === ${JSON.stringify(route)} && document.readyState === 'complete'`);
      await evaluate('document.fonts.ready.then(() => true)');
    };
    const resize = (width, height = 1000) => cdp('Emulation.setDeviceMetricsOverride', { width, height, deviceScaleFactor: 1, mobile: false });
    const capture = async name => {
      await evaluate('new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))');
      const { cssContentSize } = await cdp('Page.getLayoutMetrics');
      const shot = await cdp('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true,
        clip: { x: 0, y: 0, width: cssContentSize.width, height: cssContentSize.height, scale: 1 } });
      fs.writeFileSync(path.join(screenshots, `${name}.png`), Buffer.from(shot.data, 'base64'));
    };
    const click = async selector => {
      const point = await evaluate(`(async () => {
        const el = document.querySelector(${JSON.stringify(selector)});
        el.scrollIntoView({block:'center', behavior:'instant'});
        await new Promise(resolve => requestAnimationFrame(resolve));
        const r = el.getBoundingClientRect(); return {x:r.x+r.width/2, y:r.y+r.height/2};
      })()`);
      await cdp('Input.dispatchMouseEvent', { type: 'mousePressed', button: 'left', clickCount: 1, ...point });
      await cdp('Input.dispatchMouseEvent', { type: 'mouseReleased', button: 'left', clickCount: 1, ...point });
    };
    await cdp('Runtime.enable');
    await cdp('Page.enable');
    await cdp('Network.enable');
    await resize(1440);
    await navigate('/');
    await until("document.querySelectorAll('#tileGrid .tile').length === 15");
    await evaluate(`Promise.all(['/static/icons/garden-tools-v1.webp','/static/images/garden-workshop-v1.webp'].map(src => new Promise((resolve,reject) => { const image = new Image(); image.onload = resolve; image.onerror = reject; image.src = src; }))).then(() => true)`);
    assert.equal(await evaluate("document.querySelectorAll('#tileGrid .tile-big').length"), 2);
    assert.equal(await evaluate("document.querySelectorAll('#tileGrid .garden-sprite').length"), 6);
    assert.equal(await evaluate("document.querySelectorAll('#tileGrid svg use').length"), 9);
    assert.equal(await evaluate("document.querySelectorAll('.flow-bar, .flow-page-bar').length"), 0, 'workflow rail is gone');
    await capture('home-desktop');
    assert.equal(await evaluate('document.documentElement.scrollWidth <= innerWidth'), true, 'desktop has no horizontal overflow');
    await click('[data-tool="adjust"] .tile-link');
    await until("document.querySelector('#taskDialog').open");
    await capture('task-dialog');
    assert.ok(await evaluate("document.querySelector('#taskDialogTitle').textContent.includes('调整页面')"));
    await evaluate("document.querySelector('#taskFolderInput').value = ''");
    await click('#taskPrimaryBtn');
    assert.equal(await evaluate("document.querySelector('#taskDialog').open"), true, 'empty path keeps dialog open');
    await until("document.querySelectorAll('dialog[open]').length === 2");
    await click('.app-dialog-ok');
    await cdp('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
    await cdp('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
    await until("!document.querySelector('#taskDialog').open");
    await until("document.querySelector('[data-tool=adjust]').classList.contains('is-back')");
    const chainBefore = await evaluate("document.querySelector('[data-chain=hidden]').checked");
    await click('[data-chain=hidden]');
    assert.equal(await evaluate("Tools.chainOptions().hidden"), !chainBefore);
    await click('[data-chain=hidden]');
    // 从某一页回首页：那张卡片边框闪几下后消失，网址里的 from 去掉，卡片不挪位置
    await navigate('/clean');
    assert.equal(await evaluate("document.querySelector('.header-nav a.nav-link').getAttribute('href')"), '/?from=clean');
    await navigate('/workbench');
    assert.equal(await evaluate("document.querySelector('.nav-next').textContent.trim()"), '🏠 返回首页');
    assert.equal(await evaluate("document.querySelector('.nav-next').getAttribute('href')"), '/?from=cells');
    await cdp('Page.navigate', { url: base + '/?from=clean' });
    await until("location.pathname + location.search === '/' && document.querySelector('[data-tool=clean]')?.classList.contains('is-back')");
    const before = await evaluate("JSON.stringify(document.querySelector('[data-tool=clean]').getBoundingClientRect())");
    await until("!document.querySelector('[data-tool=clean]').classList.contains('is-back')");
    assert.equal(await evaluate("JSON.stringify(document.querySelector('[data-tool=clean]').getBoundingClientRect())"), before, 'flash does not move the tile');
    for (const width of [390, 768, 1024]) {
      await resize(width, 900);
      await evaluate('scrollTo({top:0, behavior:"instant"})');
      await capture(`home-${width}`);
      assert.equal(await evaluate('document.documentElement.scrollWidth <= innerWidth'), true, `no overflow at ${width}`);
      assert.equal(await evaluate("getComputedStyle(document.querySelector('#tileGrid')).display"), 'grid');
    }
    const routes = ['/clean', '/workbench', '/records', '/history', '/pdf-review', '/analysis', '/repair-colors',
      '/excel-diff', '/distribute', '/distribute/start', '/distribute?mode=json', '/distribute/manage'];
    await resize(1440);
    for (const route of routes) {
      await navigate(route);
      await capture(route.slice(1).replace(/[^a-z]+/g, '-'));
      assert.ok(await evaluate('document.body.innerText.length > 50'), `${route} renders`);
      assert.ok(await evaluate("document.querySelector('header.app-header .header-nav .ui-icon')"), `${route} uses the shared header`);
    }
    await resize(320, 900);
    for (const route of routes) {
      await navigate(route);
      assert.ok(await evaluate('document.documentElement.scrollWidth <= document.documentElement.clientWidth'), `${route} fits a 320px viewport`);
    }
    assert.deepEqual(errors, [], 'no runtime exceptions');
    assert.deepEqual(failedRequests, [], 'no failed resources or API responses');
    console.log(`Garden UI: sprites, SVGs, dialog, workflow interactions, shared headers and responsive layouts across ${routes.length} tool routes passed.`);
    console.log('Screenshots:', screenshots);
  } finally {
    if (socket) socket.close();
    if (child && child.exitCode === null) child.kill();
    for (const entry of pending.values()) clearTimeout(entry.timer);
    console.log('Screenshot directory:', screenshots);
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
