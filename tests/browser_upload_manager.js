// Optional development test: real Chromium + isolated Python toolbox server, no npm install.
// node tests/browser_upload_manager.js
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {spawn} = require('node:child_process');
const root = path.resolve(__dirname, '..');
const browser = [process.env.CHROME_BIN, process.env.EDGE_BIN,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe']
  .find(value => value && fs.existsSync(value));
if (!browser) throw Error('Set CHROME_BIN or EDGE_BIN to run the optional browser test.');
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'upload-manager-browser-'));
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
async function until(fn, message) {
  const deadline = Date.now() + 20000;
  while (Date.now() < deadline) { const result = await fn(); if (result) return result; await pause(100); }
  throw Error(message);
}

async function main() {
  let python, chrome, socket;
  const pending = new Map();
  const errors = [];
  let sequence = 0;
  try {
    // An empty pre-created execution DB prevents importing the user's existing task state.
    const script = `import os, sqlite3
from pathlib import Path
from http.server import ThreadingHTTPServer
base = Path(os.environ['CLEAR_HANDLE_EXCEL_DATA'])
sqlite3.connect(base / 'execution.sqlite3').close()
from web_clip_fixer import WebHandler
from core import upload_manager
upload_manager.init_db()
with upload_manager.connect() as connection:
    upload_manager.merge_items(connection, [{'group':'parent','name':name,'total':total,'processed':total,'date':'2026-11-01','complete':True,'owner':'owner'} for name,total in [('first',4),('second',6),('third',10)]])
server = ThreadingHTTPServer(('127.0.0.1', 0), WebHandler)
print(server.server_address[1], flush=True)
server.serve_forever()
`;
    python = spawn(process.env.PYTHON || 'python', ['-u', '-c', script], {
      cwd: root, windowsHide:true,
      env:{...process.env, CLEAR_HANDLE_EXCEL_DATA:temp, UPLOAD_DESK_DB:path.join(temp,'manager.sqlite3'), PYTHONIOENCODING:'utf-8'}
    });
    let output = '', pythonErrors = '';
    python.stdout.on('data', chunk => { output += chunk; });
    python.stderr.on('data', chunk => { pythonErrors += chunk; });
    const port = await until(() => {
      if (python.exitCode !== null) throw Error('Python preview failed: ' + pythonErrors);
      return output.match(/^\d+$/m)?.[0];
    }, 'Python server did not start');
    const base = 'http://127.0.0.1:' + port;
    const profile = path.join(temp, 'profile');
    chrome = spawn(browser, ['--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
      '--remote-debugging-port=0', '--user-data-dir=' + profile, 'about:blank'], {windowsHide:true, stdio:'ignore'});
    const debugPort = await until(() => {
      const active = path.join(profile, 'DevToolsActivePort');
      return fs.existsSync(active) && fs.readFileSync(active, 'utf8').split('\n')[0];
    }, 'Chromium debugging did not start');
    const tabs = await (await fetch('http://127.0.0.1:' + debugPort + '/json/list')).json();
    const tab = tabs.find(tab => tab.type === 'page');
    socket = new WebSocket(tab.webSocketDebuggerUrl);
    await new Promise((resolve,reject) => { socket.onopen=resolve; socket.onerror=reject; });
    socket.onmessage = event => {
      const result = JSON.parse(event.data);
      if (result.id && pending.has(result.id)) {
        const entry = pending.get(result.id); pending.delete(result.id); clearTimeout(entry.timer);
        if (result.error) entry.reject(Error(JSON.stringify(result.error))); else entry.resolve(result.result);
      } else if (result.method === 'Runtime.exceptionThrown') {
        errors.push(result.params.exceptionDetails.exception?.description || result.params.exceptionDetails.text);
      }
    };
    function cdp(method, params={}) {
      return new Promise((resolve,reject) => {
        const id = ++sequence;
        const timer = setTimeout(() => { pending.delete(id); reject(Error('CDP timeout: '+method)); }, 15000);
        pending.set(id, {resolve,reject,timer}); socket.send(JSON.stringify({id,method,params}));
      });
    }
    async function evaluate(expression) {
      const result = await cdp('Runtime.evaluate', {expression, awaitPromise:true, returnByValue:true});
      if (result.exceptionDetails) throw Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
      return result.result.value;
    }
    async function navigate(url) {
      await cdp('Page.navigate', {url});
      await until(() => evaluate(`location.href===${JSON.stringify(url)} && document.readyState==='complete'`), 'Page did not load');
    }
    await cdp('Runtime.enable');
    await cdp('Page.enable');
    await cdp('Emulation.setDeviceMetricsOverride', {width:1440,height:1000,deviceScaleFactor:1,mobile:false});
    await navigate(base + '/distribute/manage');
    await until(() => evaluate(`document.getElementById('storage-status')?.textContent.includes('SQLite')`), 'Manager did not load its records');
    assert.equal(await evaluate(`items.length`), 3);
    assert.equal(await evaluate(`document.querySelectorAll('script[src*="node_modules"],link[href^="http"],script[src^="http"]').length`), 0);
    assert.equal(await evaluate(`document.getElementById('plan-date').type`), 'date');
    // Discover the controls from the actual product labels.
    await evaluate(`(()=>{const label=[...document.querySelectorAll('label')].find(label=>label.textContent.includes('目标'));if(!label)throw Error('Target field missing');label.querySelector('input').value='10';const generate=[...document.querySelectorAll('button')].find(button=>/生成.*方案/.test(button.textContent));if(!generate)throw Error('Generate button missing');generate.click();})()`);
    await until(() => evaluate(`document.getElementById('plan-create') && !document.getElementById('planner-results').hidden`), 'Allocation preview was not generated');
    assert.equal(await evaluate(`allocationPlans[0].totals.reduce((a,b)=>a+b,0)`), 10);
    await evaluate(`document.getElementById('plan-create').click()`);
    await until(() => evaluate(`!!document.querySelector('dialog[open]')`), 'Confirmation dialog missing');
    await evaluate(`(()=>{const confirm=[...document.querySelectorAll('dialog[open] button')].find(button=>button.textContent.includes('确认并保存'));if(!confirm)throw Error('Confirmation button missing');confirm.click();})()`);
    await until(() => evaluate(`!!document.getElementById('manager-next-copy')`), 'Saved plan has no copy handoff');
    const handoff = await evaluate(`document.getElementById('manager-next-copy').href`);
    assert.match(handoff, /mode=json&plan=/);
    const history = await (await fetch(base + '/api/upload-manager/allocation-history')).json();
    assert.equal(history.length, 1);
    assert.equal(history[0].assignments.reduce((sum,item)=>sum+item.uploaded,0), 0);
    const screenshot = await cdp('Page.captureScreenshot', {format:'png', captureBeyondViewport:false});
    const screenshotPath = path.join(temp, 'manager.png');
    fs.writeFileSync(screenshotPath, Buffer.from(screenshot.data,'base64'));
    // Narrow layout must keep its own document within the viewport; tables may scroll inside their container.
    await cdp('Emulation.setDeviceMetricsOverride', {width:390,height:844,deviceScaleFactor:1,mobile:false});
    assert.equal(await evaluate(`document.documentElement.scrollWidth <= innerWidth+1`), true);
    await cdp('Emulation.setDeviceMetricsOverride', {width:1440,height:1000,deviceScaleFactor:1,mobile:false});
    await navigate(handoff);
    await until(() => evaluate(`document.getElementById('distFileName')?.textContent.includes('已带入')`), 'Plan did not reach copy page');
    assert.deepEqual(await evaluate(`JSON.parse(document.getElementById('distPlan').value)`), history[0].manifest);
    assert.equal(await evaluate(`document.getElementById('distStart').disabled`), true);
    await evaluate(`document.getElementById('distSource').value='changed';document.getElementById('distSource').dispatchEvent(new Event('input'));`);
    assert.equal(await evaluate(`document.getElementById('distStart').disabled`), true);
    await navigate(base + '/distribute?mode=json&plan=missing');
    await until(() => evaluate(`document.getElementById('distError')?.textContent.includes('不存在')`), 'Missing plan did not show its error');
    assert.equal(await evaluate(`document.getElementById('distPlan').value`), '');

    async function post(route, payload) {
      const response = await fetch(base + route, {
        method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)
      });
      const result = await response.json();
      assert.equal(response.ok, true, result.error);
      return result;
    }
    const shared = path.join(temp, 'mock-shared');
    fs.mkdirSync(shared);
    fs.writeFileSync(path.join(shared, 'a.xlsx'), 'original');
    fs.writeFileSync(path.join(shared, 'b.xlsx'), 'unselected');
    const {task} = await post('/api/tasks/new', {shared, work_root:path.join(temp, 'local')});
    const {preview} = await post('/api/distribute/fetch_preview', {task_id:task.id});
    await post('/api/distribute/start', {preview_id:preview.id});
    await until(async () => !(await (await fetch(base + '/api/distribute')).json()).running, 'Initial fetch did not finish');
    fs.writeFileSync(path.join(shared, 'a.xlsx'), 'colleague edit');
    fs.writeFileSync(path.join(task.work, 'a.xlsx'), 'my old edit');

    await navigate(base + '/');
    await until(() => evaluate(`document.getElementById('workTaskUploadLink').pathname === '/distribute'`), 'Step 9 did not link to current-task upload');
    await navigate(base + '/distribute');
    await until(() => evaluate(`document.querySelectorAll('#distTaskFiles input').length === 2`), 'Task files did not load');
    await evaluate(`document.getElementById('distTaskCount').value='1';document.getElementById('distTaskPickCount').click();`);
    assert.deepEqual(await evaluate(`[...document.querySelectorAll('#distTaskFiles input:checked')].map(box=>box.value)`), ['a.xlsx']);
    const upload = async () => {
      await evaluate(`document.getElementById('distTaskUpload').click()`);
      await until(() => evaluate(`!!document.querySelector('dialog[open]')`), 'Upload confirmation missing');
      await evaluate(`[...document.querySelectorAll('dialog[open] button')].find(button=>button.textContent==='开始检查').click()`);
    };
    await upload();
    await until(() => evaluate(`document.getElementById('distTaskUpload').textContent.includes('确认重新处理完毕') && !document.getElementById('distTaskUpload').disabled`), 'Review pause missing');
    assert.equal(fs.readFileSync(path.join(task.reference, 'a.xlsx'), 'utf8'), 'colleague edit');
    const pdfLink = await evaluate(`document.querySelector('#distTaskReview a').href`);
    await navigate(pdfLink);
    await until(() => evaluate(`document.getElementById('taskDialog').open && document.getElementById('taskFolderInput').value===${JSON.stringify(task.work)}`), 'PDF export link did not open Step 7 for the working folder');
    fs.writeFileSync(path.join(task.work, 'a.xlsx'), 'my reviewed edit');
    await navigate(base + '/distribute');
    await until(() => evaluate(`document.querySelectorAll('#distTaskFiles input').length === 2`), 'Task files did not reload');
    await evaluate(`document.querySelector('#distTaskFiles input').click()`);
    await upload();
    await until(() => evaluate(`document.getElementById('distTaskFiles').textContent.includes('当前版本已上传')`), 'Upload record missing');
    assert.equal(fs.readFileSync(path.join(shared, 'a.xlsx'), 'utf8'), 'my reviewed edit');
    assert.equal(fs.readFileSync(path.join(shared, 'b.xlsx'), 'utf8'), 'unselected');
    await evaluate(`document.getElementById('distTaskPickCount').click()`);
    assert.deepEqual(await evaluate(`[...document.querySelectorAll('#distTaskFiles input:checked')].map(box=>box.value)`), ['b.xlsx']);
    assert.deepEqual(errors, []);
    console.log(JSON.stringify({checks:['manager loads SQLite records','no Node/CDN runtime','native date controls','allocation preview','save with confirmation','saved snapshot handoff','mobile layout','missing-plan handling','copy requires new preview','selected batch upload and review','PDF export link','persistent upload status'], screenshot:screenshotPath},null,2));
  } finally {
    socket?.close();
    for (const entry of pending.values()) { clearTimeout(entry.timer); entry.reject(Error('Browser closing')); }
    for (const child of [chrome,python]) if (child && child.exitCode===null) child.kill();
    // Keep the uniquely created temp folder and screenshot for review; never touch user data.
  }
}
main().catch(error=>{console.error(error);process.exitCode=1;});
