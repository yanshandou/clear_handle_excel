const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {spawn} = require('node:child_process');
const browser = process.env.EDGE_BIN || 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';
const pause = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'structures-browser-'));
  const server = spawn('python', ['-m', 'http.server', '18766', '--bind', '127.0.0.1'], {stdio:'ignore', windowsHide:true});
  const edge = spawn(browser, ['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0', `--user-data-dir=${profile}`, 'about:blank'], {stdio:'ignore', windowsHide:true});
  let socket;
  try {
    for (let i=0; i<100 && !fs.existsSync(path.join(profile,'DevToolsActivePort')); i++) await pause(100);
    const port=fs.readFileSync(path.join(profile,'DevToolsActivePort'),'utf8').split('\n')[0];
    const target=await (await fetch(`http://127.0.0.1:${port}/json/new?about:blank`,{method:'PUT'})).json();
    socket=new WebSocket(target.webSocketDebuggerUrl); await new Promise((r,j)=>{socket.onopen=r;socket.onerror=j});
    let id=0; const pending=new Map();
    socket.onmessage=e=>{const m=JSON.parse(e.data);if(m.id && pending.has(m.id)){const p=pending.get(m.id);pending.delete(m.id);m.error?p.reject(m.error):p.resolve(m.result)}};
    const call=(method,params={})=>new Promise((resolve,reject)=>{const n=++id;pending.set(n,{resolve,reject});socket.send(JSON.stringify({id:n,method,params}))});
    const evalJS=async expression=>{const r=await call('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.text);return r.result.value};
    await pause(500); await call('Page.navigate',{url:'http://127.0.0.1:18766/static/ticket_structures.html'});
    for(let i=0;i<100 && !await evalJS("document.readyState === 'complete' && !!document.querySelector('#tsCreate')");i++) await pause(100);
    await evalJS(`(() => {
      window.lib = {templates: []};
      window.seq = 0;
      const walk = (nodes, fn) => nodes.forEach(n => { fn(n); if (n.kind === 'dir') walk(n.children || [], fn); });
      const warningsOf = tree => { const out = []; walk(tree, n => { if (n.kind === 'file' && !n.source) out.push({id: n.id, name: n.name, reason: '未指定来源文件'}); }); return out; };
      window.fetch = async (url, options) => {
        if (url === '/api/ticket_structures') return Response.json({ok: true, templates: window.lib.templates, dir: '数据/目录结构'});
        if (url === '/api/icons') return Response.json({ok: true, icons: [{id: 'aaaaaa', name: '海星', category: '', version: 1}, {id: 'bbbbbb', name: '帆船', category: '船', version: 2}], tools: {}});
        const body = JSON.parse(options.body);
        if (url.endsWith('/create')) {
          if (window.lib.templates.some(t => t.name === body.name)) return Response.json({ok: false, error: '已经有叫「' + body.name + '」的模板了'}, {status: 400});
          const template = {id: 't' + (++window.seq), name: body.name, tree: [], warnings: []};
          window.lib.templates.unshift(template);
          return Response.json({ok: true, template});
        }
        const template = window.lib.templates.find(t => t.id === body.id);
        if (!template) return Response.json({ok: false, error: '没有这个目录结构模板'}, {status: 400});
        if (url.endsWith('/rename')) { template.name = body.name; return Response.json({ok: true, template}); }
        if (url.endsWith('/delete')) { window.lib.templates = window.lib.templates.filter(t => t.id !== body.id); return Response.json({ok: true, name: template.name}); }
        if (url.endsWith('/save')) {
          window.lastSave = body;
          let bad = '';
          walk(body.tree, n => { if (n.kind === 'file' && n.source && !n.source.startsWith('ok:')) bad = n.name; });
          if (bad) return Response.json({ok: false, error: '「' + bad + '」的来源请填写完整路径'}, {status: 400});
          template.tree = body.tree;
          template.warnings = warningsOf(body.tree);
          return Response.json({ok: true, template, warnings: template.warnings});
        }
        return Response.json({ok: false, error: '不认识的目录结构操作'}, {status: 400});
      };
    })()`);
    const setInput = (selector, value) => evalJS(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.value=${JSON.stringify(value)};e.dispatchEvent(new Event('input',{bubbles:true}))})()`);
    const clickText = (selector, text) => evalJS(`[...document.querySelectorAll(${JSON.stringify(selector)})].find(b=>b.textContent.trim()===${JSON.stringify(text)}).click()`);
    const dialogOk = () => evalJS("document.querySelector('.app-dialog-ok').click()");
    const dialogCancel = () => evalJS("document.querySelector('.app-dialog-cancel').click()");
    const waitFor = async expression => { for (let i = 0; i < 50; i++) { if (await evalJS(expression)) return; await pause(50); } throw Error('等待超时: ' + expression); };

    // 新建模板
    await setInput('#tsNewName', '基本設計');
    await evalJS("document.querySelector('#tsCreate').click()");
    await waitFor("document.querySelectorAll('#tsList .ts-item').length === 1");
    assert.match(await evalJS("document.querySelector('#tsEditor .ts-title').textContent"), /基本設計/);
    // 搭树：根目录 + 子文件
    await clickText('#tsEditor .ts-toolbar .btn', '+ 根目录');
    await setInput('#tsEditor .ts-node .ts-name', '01_要件');
    await clickText('#tsEditor .ts-node .ts-row .btn', '+ 文件');
    await setInput('#tsEditor .ts-tree .ts-tree .ts-name', 'ビュー定義書.xlsx');
    await setInput('#tsEditor .ts-tree .ts-tree .ts-source', 'ok:D:/模板/ビュー定義書.xlsx');
    assert.match(await evalJS("document.querySelector('#tsEditor .ts-dirty').textContent"), /未保存/);
    await clickText('#tsEditor .ts-toolbar .btn', '💾 保存');
    assert.equal(await evalJS("document.querySelector('#tsEditor .ts-dirty').textContent"), '');
    assert.equal(await evalJS("window.lastSave.tree[0].name"), '01_要件');
    assert.equal(await evalJS("window.lastSave.tree[0].children[0].name"), 'ビュー定義書.xlsx');
    assert.ok(await evalJS("window.lastSave.tree[0].id && window.lastSave.tree[0].children[0].id"));
    // 来源缺失警告 + 非法来源拒绝
    await clickText('#tsEditor .ts-node .ts-row .btn', '+ 文件');
    await setInput('#tsEditor .ts-tree .ts-tree .ts-node:nth-child(2) .ts-name', '空来源.txt');
    await clickText('#tsEditor .ts-toolbar .btn', '💾 保存');
    assert.match(await evalJS("document.querySelector('#tsStatus').textContent"), /1 个问题/);
    assert.equal(await evalJS("document.querySelectorAll('#tsEditor .ts-node.is-warn').length"), 1);
    await setInput('#tsEditor .ts-tree .ts-tree .ts-node:nth-child(2) .ts-source', '相对路径.txt');
    await clickText('#tsEditor .ts-toolbar .btn', '💾 保存');
    assert.match(await evalJS("document.querySelector('#tsStatus').textContent"), /完整路径/);
    await setInput('#tsEditor .ts-tree .ts-tree .ts-node:nth-child(2) .ts-source', '');
    // 删除节点
    await evalJS("document.querySelectorAll('#tsEditor .ts-tree .ts-tree .ts-node')[1].querySelector('.btn-danger').click()");
    assert.equal(await evalJS("document.querySelectorAll('#tsEditor .ts-tree .ts-tree .ts-node').length"), 1);
    await clickText('#tsEditor .ts-toolbar .btn', '💾 保存');
    assert.equal(await evalJS("document.querySelector('#tsStatus').textContent"), '已保存。');
    // 目录挑图标：点 📁 弹出图标库网格，选了以后存进模板，再选「默认图标」去掉
    await evalJS("window.dispatchEvent(new Event('focus'))");
    // 图标库是异步读的：读到之前弹出来只有「默认图标」，关掉再点，直到 3 个都在
    for (let i = 0; i < 50 && !await evalJS("document.querySelectorAll('.ts-icon-pop .ts-icon-opt').length === 3"); i++) {
      await evalJS("document.body.click(); document.querySelector('#tsEditor .ts-icon').click()");
      await pause(100);
    }
    await waitFor("document.querySelectorAll('.ts-icon-pop .ts-icon-opt').length === 3");
    await evalJS("[...document.querySelectorAll('.ts-icon-pop .ts-icon-opt')].find(b => b.textContent.includes('帆船')).click()");
    assert.equal(await evalJS("!!document.querySelector('.ts-icon-pop')"), false);
    assert.match(await evalJS("document.querySelector('#tsEditor .ts-icon img').src"), /id=bbbbbb/);
    await clickText('#tsEditor .ts-toolbar .btn', '💾 保存');
    assert.equal(await evalJS("window.lastSave.tree[0].icon"), 'bbbbbb');
    await evalJS("document.querySelector('#tsEditor .ts-icon').click()");
    await evalJS("[...document.querySelectorAll('.ts-icon-pop .ts-icon-opt')].find(b => b.textContent.includes('默认图标')).click()");
    await clickText('#tsEditor .ts-toolbar .btn', '💾 保存');
    assert.equal(await evalJS("'icon' in window.lastSave.tree[0]"), false);
    // 模板改名
    await setInput('#tsRename', '詳細設計');
    await evalJS("document.querySelector('#tsRenameBtn').click()");
    assert.match(await evalJS("document.querySelector('#tsList .ts-item-name').textContent"), /詳細設計/);
    // 未保存切换：先拦下，保存后再切
    await setInput('#tsNewName', 'テスト仕様');
    await evalJS("document.querySelector('#tsCreate').click()");
    await waitFor("document.querySelectorAll('#tsList .ts-item').length === 2");
    await clickText('#tsEditor .ts-toolbar .btn', '+ 根目录');
    await evalJS("document.querySelectorAll('#tsList .ts-item')[1].click()");
    assert.equal(await evalJS("!!document.querySelector('.app-dialog[open]')"), true);
    await dialogCancel();
    assert.match(await evalJS("document.querySelector('#tsEditor .ts-title').textContent"), /テスト仕様/);
    await clickText('#tsEditor .ts-toolbar .btn', '💾 保存');
    await evalJS("document.querySelectorAll('#tsList .ts-item')[1].click()");
    assert.match(await evalJS("document.querySelector('#tsEditor .ts-title').textContent"), /詳細設計/);
    // 删除模板（确认后）
    await evalJS("document.querySelector('#tsDelete').click()");
    assert.equal(await evalJS("!!document.querySelector('.app-dialog[open]')"), true);
    await dialogOk();
    await waitFor("document.querySelectorAll('#tsList .ts-item').length === 1");
    assert.match(await evalJS("document.querySelector('#tsList .ts-item-name').textContent"), /テスト仕様/);
    console.log('Structure template management: create, tree edit, save, warnings, folder icons, rename, switch guard and delete passed.');
  } finally {if(socket)socket.close(); edge.kill();server.kill();}
})().catch(e=>{console.error(e);process.exitCode=1});