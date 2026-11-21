const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn } = require('node:child_process');
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
const browser = process.env.EDGE_BIN || 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';

(async () => {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'icon-assign-'));
  const server = spawn('python', ['-m', 'http.server', '18767', '--bind', '127.0.0.1'], { stdio: 'ignore', windowsHide: true });
  const edge = spawn(browser, ['--headless=new', '--disable-gpu', '--no-first-run', '--remote-debugging-port=0', `--user-data-dir=${profile}`, 'about:blank'], { stdio: 'ignore', windowsHide: true });
  let socket;
  try {
    for (let i = 0; i < 100 && !fs.existsSync(path.join(profile, 'DevToolsActivePort')); i++) await pause(100);
    const port = fs.readFileSync(path.join(profile, 'DevToolsActivePort'), 'utf8').split('\n')[0];
    const target = await (await fetch(`http://127.0.0.1:${port}/json/new?about:blank`, { method: 'PUT' })).json();
    socket = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
    let id = 0; const pending = new Map();
    socket.onmessage = event => { const m = JSON.parse(event.data); if (m.id && pending.has(m.id)) { const p = pending.get(m.id); pending.delete(m.id); m.error ? p.reject(m.error) : p.resolve(m.result); } };
    const call = (method, params = {}) => new Promise((resolve, reject) => { const n = ++id; pending.set(n, { resolve, reject }); socket.send(JSON.stringify({ id: n, method, params })); });
    const evalJS = async expression => { const r = await call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true }); if (r.exceptionDetails) throw Error(r.exceptionDetails.exception?.description || r.exceptionDetails.text); return r.result.value; };
    const waitFor = async expression => { for (let i = 0; i < 100 && !await evalJS(expression); i++) await pause(50); };
    const setInput = (selector, value) => evalJS(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.value=${JSON.stringify(value)};e.dispatchEvent(new Event('input',{bubbles:true}))})()`);
    await call('Page.enable');
    await call('Page.addScriptToEvaluateOnNewDocument', { source: `
      window.mockIcons = [
        {id:'aaaaaa',name:'A',category:'一',source:'a.png',folders:[],version:1,width:256,height:256},
        {id:'bbbbbb',name:'B',category:'二',source:'b.png',folders:[],version:1,width:256,height:256},
        {id:'cccccc',name:'C',category:'二',source:'c.png',folders:[],version:1,width:256,height:256}
      ];
      window.applyCalls = [];
      const realFetch = window.fetch;
      window.fetch = async (url, options) => {
        if (url === '/api/icons') {
          const counts = {}; for (const i of mockIcons) counts[i.category] = (counts[i.category] || 0) + 1;
          return Response.json({ok:true,icons:mockIcons,categories:Object.entries(counts).map(([name,count])=>({name,count})),tools:{},duplicates:[]});
        }
        if (url.startsWith('/api/icons/subfolders?')) {
          const folder = new URL(url, location.origin).searchParams.get('folder');
          if (folder.includes('gone')) return Response.json({ok:false,error:'找不到这个文件夹'}, {status:400});
          return Response.json({ok:true,dir:folder,total:2,folders:[
            {path:'C:\\\\mock\\\\c',icon_id:'cccccc'},
            {path:'C:\\\\mock\\\\a',icon_id:''}
          ]});
        }
        if (url === '/api/icons/apply') {
          const p = JSON.parse(options.body);
          window.applyCalls.push(p);
          const results = p.folders.map(f => f.toLowerCase() === 'c:\\\\mock\\\\b'
            ? {folder:f,ok:false,error:'找不到这个文件夹'} : {folder:f,ok:true});
          return Response.json({ok:true,results,ok_count:results.filter(r=>r.ok).length,windows:true});
        }
        return realFetch(url, options);
      };
    ` });
    await pause(500); await call('Page.navigate', { url: 'http://127.0.0.1:18767/static/icon_library.html' });
    await waitFor("document.querySelectorAll('.ic-card').length === 3 && !!document.querySelector('#ibRows')");
    assert.match(await evalJS("document.querySelector('#ibRows').textContent"), /列表还是空的/);
    assert.equal(await evalJS("document.querySelector('#ibApply').disabled"), true);

    // 粘贴多行路径：带引号、重复行、空行 → 去重后 2 行
    await setInput('#ibPaths', '"C:\\mock\\a"\nC:\\mock\\a\nC:\\mock\\b\n\n');
    await evalJS("document.querySelector('#ibFromPaths').click()");
    assert.equal(await evalJS("document.querySelectorAll('#ibRows .ib-row').length"), 2);
    assert.match(await evalJS("document.querySelector('#ibStatus').textContent"), /加进列表 2 个/);
    assert.deepEqual(await evalJS("[...document.querySelectorAll('#ibRows .ib-path')].map(e=>e.textContent)"), ['C:\\mock\\a', 'C:\\mock\\b']);

    // 弹网格点选：第一行挑 B
    await evalJS("document.querySelector('[data-slot=\"0\"]').click()");
    assert.equal(await evalJS("document.querySelector('#ibPicker').hidden"), false);
    assert.equal(await evalJS("document.querySelectorAll('#ibPicker .ib-pick').length"), 3);
    await evalJS("document.querySelector('#ibPicker [data-pick=\"bbbbbb\"]').click()");
    assert.equal(await evalJS("document.querySelector('#ibPicker').hidden"), true);
    assert.match(await evalJS("document.querySelector('[data-slot=\"0\"] img').src"), /id=bbbbbb/);
    assert.equal(await evalJS("document.querySelector('[data-slot=\"0\"]').classList.contains('is-set')"), true);

    // Esc 和点外面都关网格，不改选择
    await evalJS("document.querySelector('[data-slot=\"1\"]').click()");
    await call('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
    await call('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 });
    assert.equal(await evalJS("document.querySelector('#ibPicker').hidden"), true);
    assert.equal(await evalJS("document.querySelector('[data-slot=\"1\"]').textContent.trim()"), '+');
    await evalJS("document.querySelector('[data-slot=\"1\"]').click()");
    await evalJS("document.querySelector('h2').click()");
    assert.equal(await evalJS("document.querySelector('#ibPicker').hidden"), true);
    await evalJS("document.querySelector('[data-slot=\"1\"]').click()");
    await evalJS("document.querySelector('#ibPicker [data-pick=\"aaaaaa\"]').click()");
    assert.match(await evalJS("document.querySelector('[data-slot=\"1\"] img').src"), /id=aaaaaa/);
    assert.equal(await evalJS("document.querySelector('#ibApply').textContent"), '应用到 2 个文件夹');

    // 列子文件夹：报错回显；成功后预填当前图标、重复列不叠加、已挑的不被覆盖
    await setInput('#ibParent', 'C:\\gone');
    await evalJS("document.querySelector('#ibListSub').click()");
    await waitFor("document.querySelector('#ibStatus').textContent.includes('找不到这个文件夹')");
    await setInput('#ibParent', 'C:\\mock');
    await evalJS("document.querySelector('#ibListSub').click()");
    await waitFor("document.querySelectorAll('#ibRows .ib-row').length === 3");
    assert.match(await evalJS("document.querySelector('#ibStatus').textContent"), /列出 2 个子文件夹，新增 1 个/);
    assert.match(await evalJS("document.querySelector('[data-slot=\"2\"] img').src"), /id=cccccc/);
    assert.match(await evalJS("document.querySelector('[data-slot=\"0\"] img').src"), /id=bbbbbb/);
    await evalJS("document.querySelector('#ibListSub').click()");
    await waitFor("document.querySelector('#ibStatus').textContent.includes('新增 0 个')");
    assert.equal(await evalJS("document.querySelectorAll('#ibRows .ib-row').length"), 3);
    assert.equal(await evalJS("JSON.parse(localStorage.getItem('clear_handle_excel.iconLibrary')).assignRows.length"), 3);

    // 没挑图标的行应用时跳过；按图标分组调用；失败行单独回显
    await setInput('#ibPaths', 'C:\\mock\\d');
    await evalJS("document.querySelector('#ibFromPaths').click()");
    assert.equal(await evalJS("document.querySelectorAll('#ibRows .ib-row').length"), 4);
    await evalJS("document.querySelector('#ibApply').click()");
    await waitFor("window.applyCalls.length === 3 && document.querySelector('#ibStatus').textContent.includes('没成功')");
    assert.deepEqual(await evalJS('window.applyCalls'), [
      { id: 'bbbbbb', folders: ['C:\\mock\\a'] },
      { id: 'aaaaaa', folders: ['C:\\mock\\b'] },
      { id: 'cccccc', folders: ['C:\\mock\\c'] },
    ]);
    assert.match(await evalJS("document.querySelector('#ibStatus').textContent"), /套成功 2 个文件夹，1 个没成功/);
    assert.match(await evalJS("document.querySelector('#ibStatus').textContent"), /跳过 1 行没挑图标的/);
    assert.equal(await evalJS("document.querySelectorAll('#ibRows .ib-row')[0].classList.contains('is-ok')"), true);
    assert.equal(await evalJS("document.querySelectorAll('#ibRows .ib-row')[1].classList.contains('is-bad')"), true);
    assert.match(await evalJS("document.querySelectorAll('#ibRows .ib-row')[1].textContent"), /找不到这个文件夹/);
    assert.equal(await evalJS("document.querySelectorAll('#ibRows .ib-row')[2].classList.contains('is-ok')"), true);
    assert.equal(await evalJS("document.querySelectorAll('#ibRows .ib-row')[3].className"), 'ib-row');

    // 删行 / 清空
    await evalJS("document.querySelector('[data-drop=\"1\"]').click()");
    assert.equal(await evalJS("document.querySelectorAll('#ibRows .ib-row').length"), 3);
    await evalJS("document.querySelector('#ibClear').click()");
    assert.equal(await evalJS("document.querySelectorAll('#ibRows .ib-row').length"), 0);
    assert.match(await evalJS("document.querySelector('#ibRows').textContent"), /列表还是空的/);
    assert.equal(await evalJS("document.querySelector('#ibApply').disabled"), true);
    console.log('Per-folder icon assignment panel passed.');
  } finally { if (socket) socket.close(); edge.kill(); server.kill(); }
})().catch(error => { console.error(error); process.exitCode = 1; });