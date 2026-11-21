const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn } = require('node:child_process');
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
const browser = process.env.EDGE_BIN || 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';

(async () => {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'icon-batch-'));
  const server = spawn('python', ['-m', 'http.server', '18766', '--bind', '127.0.0.1'], { stdio: 'ignore', windowsHide: true });
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
    await call('Page.enable');
    await call('Page.addScriptToEvaluateOnNewDocument', { source: `
      window.mockIcons = [
        {id:'aaaaaa',name:'A',category:'一',source:'a.png',folders:['folder'],version:1,width:256,height:256},
        {id:'bbbbbb',name:'B',category:'二',source:'b.png',folders:[],version:1,width:256,height:256},
        {id:'cccccc',name:'C',category:'二',source:'c.png',folders:[],version:1,width:256,height:256}
      ];
      window.mockTools = {tool:'aaaaaa'};
      const realFetch = window.fetch;
      window.fetch = async (url, options) => {
        if (url === '/api/icons') {
          const counts = {}; for (const i of mockIcons) counts[i.category] = (counts[i.category] || 0) + 1;
          const duplicates = mockIcons.filter(i => ['aaaaaa','bbbbbb'].includes(i.id)).map(i => i.id);
          return Response.json({ok:true,icons:mockIcons,categories:Object.entries(counts).map(([name,count])=>({name,count})),tools:mockTools,duplicates:duplicates.length > 1 ? [duplicates] : []});
        }
        if (url === '/api/icons/update') { const p=JSON.parse(options.body); mockIcons.find(i=>i.id===p.id).category=p.category; return Response.json({ok:true}); }
        if (url === '/api/icons/delete') { const p=JSON.parse(options.body); window.mockIcons=mockIcons.filter(i=>i.id!==p.id); return Response.json({ok:true}); }
        if (url === '/api/icons/upload') return Response.json({ok:false,code:'duplicate_icon',error:'与图标库里的 C 重复'}, {status:400});
        if (url === '/api/icons/retrim') {
          const icon = mockIcons.find(i => i.id === JSON.parse(options.body).id);
          const info = {trimmed:true,render_width:200,render_height:100};
          Object.assign(icon, info, {version:icon.version+1});
          return Response.json({ok:true,icon,info});
        }
        return realFetch(url, options);
      };
    ` });
    await pause(500); await call('Page.navigate', { url: 'http://127.0.0.1:18766/static/icon_library.html' });
    for (let i = 0; i < 100 && !await evalJS("document.querySelectorAll('.ic-card').length === 3"); i++) await pause(100);
    assert.equal(await evalJS("document.querySelectorAll('.ic-card').length"), 3, await evalJS("location.href + ' ' + document.body.innerText.slice(0, 500)"));
    await evalJS("document.querySelector('#icDuplicates').click()");
    assert.equal(await evalJS("document.querySelectorAll('.ic-card').length"), 2);
    assert.match(await evalJS("document.querySelector('.ic-card').innerText"), /重复组 1/);
    await evalJS("document.querySelector('#icDuplicates').click()");
    await evalJS("document.querySelector('#icBatchToggle').click(); document.querySelector('[data-id=aaaaaa]').click(); document.querySelector('[data-cat=二]').click(); document.querySelector('[data-id=bbbbbb]').click()");
    assert.match(await evalJS("document.querySelector('#icBatchCount').textContent"), /2 个/);
    await evalJS("document.querySelector('#icBatchCategory').value='新分类'; document.querySelector('#icBatchMove').click()");
    for (let i = 0; i < 100 && !await evalJS("document.querySelector('#icBatchStatus').textContent.includes('2 个')"); i++) await pause(50);
    assert.deepEqual(await evalJS("mockIcons.map(i=>i.category)"), ['新分类', '新分类', '二']);
    await evalJS("document.querySelector('[data-cat=" + '"新分类"' + "]').click(); document.querySelector('#icBatchAll').click(); window.appConfirm=async text=>{window.confirmText=text;return false}; document.querySelector('#icBatchDelete').click()");
    await pause(100);
    assert.match(await evalJS('window.confirmText'), /2 个图标/);
    assert.match(await evalJS('window.confirmText'), /工具会恢复默认/);
    assert.equal(await evalJS('mockIcons.length'), 3);
    await evalJS("window.appConfirm=async()=>true; document.querySelector('#icBatchDelete').click()");
    for (let i = 0; i < 100 && await evalJS('mockIcons.length') !== 1; i++) await pause(50);
    assert.equal(await evalJS('mockIcons.length'), 1);
    for (let i = 0; i < 100 && !await evalJS("document.querySelector('#icBatchStatus').textContent.includes('已删除')"); i++) await pause(50);
    await evalJS(`{
      const input = document.querySelector('#icFile');
      const files = new DataTransfer();
      files.items.add(new File(['image'], 'copy.png', {type:'image/png'}));
      input.files = files.files;
      input.dispatchEvent(new Event('change'));
    }`);
    for (let i = 0; i < 100 && !await evalJS("document.querySelector('#icStatus').textContent.includes('已跳过')"); i++) await pause(50);
    assert.match(await evalJS("document.querySelector('#icStatus').textContent"), /已跳过 1 张重复图片/);
    assert.equal(await evalJS("document.querySelectorAll('.ic-card').length"), 1);
    await evalJS("document.querySelector('#icBatchToggle').click(); document.querySelector('[data-id=cccccc]').click(); document.querySelector('#icRetrim').click()");
    for (let i = 0; i < 100 && !await evalJS("document.querySelector('#icStatus').textContent.includes('裁剪后')"); i++) await pause(50);
    const metadata = await evalJS("document.querySelector('.ic-meta').innerText");
    assert.match(metadata, /裁剪后 200×100/);
    assert.match(metadata, /方形图标上下留透明边/);
    assert.match(await evalJS("document.querySelector('#icBig img').src"), /&v=2$/);
    console.log('Icon batch management, duplicate upload and retrim passed.');
  } finally { if (socket) socket.close(); edge.kill(); server.kill(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
