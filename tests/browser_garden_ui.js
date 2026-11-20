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
    await evaluate("localStorage.removeItem('clear_handle_excel.flows'), location.reload(), true");
    await until("document.readyState === 'complete' && window.Tools && document.querySelectorAll('#toolboxBody .app').length === Tools.TOOLS.length");
    await evaluate(`Promise.all(['/static/images/harbor/icons.webp', '/static/images/bg.png', '/static/images/ship-logo.png', '/static/images/dev-toolbox-logo.png'].map(src => {
      const image = new Image();
      image.src = src;
      return image.decode();
    })).then(() => true)`);
    assert.match(await evaluate("getComputedStyle(document.querySelector('.home-seascape')).backgroundImage"), /\/static\/images\/bg\.png/);
    assert.match(await evaluate("getComputedStyle(document.querySelector('.header-logo-badge')).backgroundImage"), /\/static\/images\/ship-logo\.png/);
    assert.equal(await evaluate(`document.querySelector('link[rel="icon"]').getAttribute('href')`), '/static/images/dev-toolbox-logo.png');
    assert.equal(await evaluate("document.querySelector('.header-title-text').textContent"), '开发百宝箱');
    // 工具箱：所有工具一样大的图标，点开展开；默认有一条流程
    assert.equal(await evaluate("document.querySelectorAll('#toolboxBody .harbor-icon').length === Tools.TOOLS.length"), true);
    assert.equal(await evaluate("document.querySelector('#toolboxPop').hidden"), true);
    assert.equal(await evaluate("document.querySelectorAll('#flowList .flow-card').length"), 1);
    await capture('home-desktop');
    assert.equal(await evaluate('document.documentElement.scrollWidth <= innerWidth'), true, 'desktop has no horizontal overflow');
    await click('#toolboxFolder');
    await until("!document.querySelector('#toolboxPop').hidden");
    await capture('home-toolbox');
    // 从工具箱拖「修复图片」进流程第 1 步前面
    const dragTo = async (fromSel, toSel, dx = -8) => {
      const p = await evaluate(`(() => { const a = document.querySelector(${JSON.stringify(fromSel)}).getBoundingClientRect();
        const b = document.querySelector(${JSON.stringify(toSel)}).getBoundingClientRect();
        return {x1: a.x + a.width / 2, y1: a.y + a.height / 2, x2: b.x + ${dx}, y2: b.y + b.height / 2}; })()`);
      await cdp('Input.dispatchMouseEvent', { type: 'mousePressed', button: 'left', buttons: 1, clickCount: 1, x: p.x1, y: p.y1 });
      for (let i = 1; i <= 8; i++) await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', button: 'left', buttons: 1, x: p.x1 + (p.x2 - p.x1) * i / 8, y: p.y1 + (p.y2 - p.y1) * i / 8 });
      await cdp('Input.dispatchMouseEvent', { type: 'mouseReleased', button: 'left', buttons: 0, clickCount: 1, x: p.x2, y: p.y2 });
    };
    await dragTo('#toolboxBody .app[data-tool=images] .app-icon', '#flowList .flow-step[data-index="0"] .app-icon');
    await until("HomeTiles.flows()[0].steps[0] === 'images'");
    assert.equal(await evaluate("location.pathname"), '/', 'dropping does not open the tool');
    assert.match(await evaluate("localStorage.getItem('clear_handle_excel.flows')"), /"images"/);
    await click('#toolboxClose');
    await until("document.querySelector('#toolboxPop').hidden");
    await click('#flowList .flow-step[data-index="0"] .flow-remove');
    await until("HomeTiles.flows()[0].steps[0] !== 'images'");
    await click('#flowList .app[data-tool="adjust"]');
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
    await until("document.querySelector('#flowList .app[data-tool=adjust]').classList.contains('is-back')");
    // 从某一页回首页：流程里的图标闪几下后消失，网址里的 from 去掉，图标不挪位置
    await navigate('/clean');
    assert.equal(await evaluate("document.querySelector('.header-nav a.nav-link').getAttribute('href')"), '/?from=clean');
    await navigate('/workbench');
    assert.equal(await evaluate("document.querySelector('.nav-next').textContent.trim()"), '🏠 返回首页');
    assert.equal(await evaluate("document.querySelector('.nav-next').getAttribute('href')"), '/?from=cells');
    await cdp('Page.navigate', { url: base + '/?from=clean' });
    await until("location.pathname + location.search === '/' && document.querySelector('#flowList .app[data-tool=clean]')?.classList.contains('is-back')");
    const before = await evaluate("JSON.stringify(document.querySelector('#flowList .app[data-tool=clean]').getBoundingClientRect())");
    await until("!document.querySelector('#flowList .app[data-tool=clean]').classList.contains('is-back')");
    assert.equal(await evaluate("JSON.stringify(document.querySelector('#flowList .app[data-tool=clean]').getBoundingClientRect())"), before, 'flash does not move the icon');
    await click('#toolboxFolder');
    await click('#toolboxBody .app[data-tool=ticket]');
    await until("location.pathname === '/ticket' && document.readyState === 'complete'");
    const ticketText = `機能追加 - 案件 #55022: 介舟ファミリーのタブレット対応／計画書の電子サイン対応（正式版）

TBSNG0000_手書きサイン
（開発）金 星（社） が9日前に追加. 約4時間前に更新.

ステータス:\t進行中\t開始日:\t2026/11/20
優先度:\t通常\t期日:\t2026/12/03
担当者:\t（開発）白 家綾（現）\t進捗率:\t
0%

対象バージョン:\t介舟ファミリークラウド総合 - V440
画面／帳票ID:\tTBSNG0000\t実績工数:\t1.00
予定工数:\t9.00`;
    const setTicketValue = (selector, value) => evaluate(`(() => {
      const input = document.querySelector(${JSON.stringify(selector)});
      input.value = ${JSON.stringify(value)};
      input.dispatchEvent(new Event('input', { bubbles: true }));
    })()`);
    const ticketFields = () => evaluate("Object.fromEntries(new FormData(document.querySelector('#ticketForm')))");
    await setTicketValue('#ticketSource', ticketText);
    assert.deepEqual(await ticketFields(), {
      title: 'TBSNG0000_手書きサイン', ticket_id: '', screen_id: 'TBSNG0000', parent: ticketText.split('\n')[0],
      status: '進行中', priority: '通常', assignee: '（開発）白 家綾（現）', progress: '0%', start_date: '2026/11/20',
      due_date: '2026/12/03', estimated_hours: '9.00', actual_hours: '1.00', version: '介舟ファミリークラウド総合 - V440',
      activity: '（開発）金 星（社） が9日前に追加. 約4時間前に更新.',
    });
    await setTicketValue('#ticketForm [name="assignee"]', '手動入力');
    await setTicketValue('#ticketSource', `task #55345\n${ticketText}`);
    assert.equal((await ticketFields()).ticket_id, '55345');
    assert.equal((await ticketFields()).assignee, '手動入力');
    await capture('ticket-filled');
    await evaluate(`window.fetch = async (url, options) => {
      if (url !== '/api/ticket/create' || options.method !== 'POST') throw new Error('Unexpected Ticket request');
      window.savedTicket = JSON.parse(options.body);
      return Response.json({ok: true, filename: 'Ticket.txt', folder: 'Desktop/Ticket #55345'});
    }`);
    await click('#ticketCreate');
    await until("document.querySelector('#ticketResult').textContent.includes('Desktop/Ticket #55345')");
    assert.equal(await evaluate('window.savedTicket.assignee'), '手動入力');
    await click('#ticketReset');
    assert.ok(Object.values(await ticketFields()).every(value => value === ''));
    await setTicketValue('#ticketForm [name="title"]', '手動入力だけ');
    await click('#ticketCreate');
    await until("window.savedTicket.title === '手動入力だけ' && document.querySelector('#ticketResult').textContent.includes('Ticket.txt')");
    assert.equal(await evaluate('window.savedTicket.ticket_id'), '');
    await navigate('/');
    for (const width of [1024, 1280, 1920]) { // 只做电脑版（pc-only）
      await resize(width, 900);
      await evaluate('scrollTo({top:0, behavior:"instant"})');
      await capture(`home-${width}`);
      assert.equal(await evaluate('document.documentElement.scrollWidth <= document.documentElement.clientWidth'), true, `no overflow at ${width}`);
    }
    const routes = ['/clean', '/workbench', '/records', '/history', '/pdf-review', '/analysis', '/repair-colors',
      '/excel-diff', '/distribute', '/distribute/start', '/distribute?mode=json', '/ticket', '/distribute/manage'];
    await resize(1440);
    for (const route of routes) {
      await navigate(route);
      await capture(route.slice(1).replace(/[^a-z]+/g, '-'));
      assert.ok(await evaluate('document.body.innerText.length > 50'), `${route} renders`);
      assert.ok(await evaluate("document.querySelector('header.app-header .header-nav .ui-icon')"), `${route} uses the shared header`);
    }
    assert.match(await evaluate("getComputedStyle(document.querySelector('#manager-task-upload .harbor-icon')).backgroundImage"), /harbor\/icons\.webp/);
    assert.match(await evaluate("document.querySelector('#empty-state span').textContent"), /导入 Excel 或添加文件夹/);
    await resize(320, 900);
    for (const route of routes) {
      await navigate(route);
      assert.ok(await evaluate('document.documentElement.scrollWidth <= document.documentElement.clientWidth'), `${route} fits a 320px viewport`);
    }
    await navigate('/pcl');
    assert.equal(await evaluate("document.documentElement.dataset.theme"), 'harbor');
    assert.equal(await evaluate("document.documentElement.scrollWidth <= document.documentElement.clientWidth"), true, 'PCL fits a 320px viewport');
    await resize(1440);
    await capture('pcl');
    await click('[data-theme-choice="neon"]');
    await until("document.documentElement.dataset.theme === 'neon'");
    await navigate('/pcl');
    assert.equal(await evaluate("document.documentElement.dataset.theme"), 'neon', 'PCL keeps the selected theme after reload');
    assert.deepEqual(errors, [], 'no runtime exceptions');
    assert.deepEqual(failedRequests, [], 'no failed resources or API responses');
    console.log(`Harbor UI: raster assets, dialog, workflow interactions, shared headers and responsive layouts across ${routes.length} tool routes passed.`);
    console.log('Screenshots:', screenshots);
  } finally {
    if (socket) socket.close();
    if (child && child.exitCode === null) child.kill();
    for (const entry of pending.values()) clearTimeout(entry.timer);
    console.log('Screenshot directory:', screenshots);
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
