const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {spawn} = require('node:child_process');
const browser = process.env.EDGE_BIN || 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';
const pause = ms => new Promise(r => setTimeout(r, ms));
(async () => {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'ticket-browser-'));
  const server = spawn('python', ['-m', 'http.server', '18765', '--bind', '127.0.0.1'], {stdio:'ignore', windowsHide:true});
  const edge = spawn(browser, ['--headless=new','--disable-gpu','--no-first-run','--remote-debugging-port=0', `--user-data-dir=${profile}`, 'about:blank'], {stdio:'ignore', windowsHide:true});
  let socket;
  try {
    for (let i=0; i<100 && !fs.existsSync(path.join(profile,'DevToolsActivePort')); i++) await pause(100);
    const port=fs.readFileSync(path.join(profile,'DevToolsActivePort'),'utf8').split('\n')[0];
    const target=await (await fetch(`http://127.0.0.1:${port}/json/new?about:blank`,{method:'PUT'})).json();
    socket=new WebSocket(target.webSocketDebuggerUrl); await new Promise((r,j)=>{socket.onopen=r;socket.onerror=j});
    let id=0; const pending=new Map();
    const call=(method,params={})=>new Promise((resolve,reject)=>{const n=++id;pending.set(n,{resolve,reject});socket.send(JSON.stringify({id:n,method,params}))});
    const iconFiles = ['appicon', 'json', 'pdfdiff', 'pick', 'report', 'reflink', 'view', 'widths', 'bl'];
    const icons = iconFiles.map((name, index) => ({id: `123ab${index}`, name, category: '传票', version: 2}));
    const iconBodies = new Map(icons.map(({id, name}) => [id, fs.readFileSync(`static/images/harbor/tools/${name}.png`, 'base64')]));
    socket.onmessage = e => {
      const m = JSON.parse(e.data);
      if (m.id && pending.has(m.id)) {
        const p = pending.get(m.id);
        pending.delete(m.id);
        m.error ? p.reject(m.error) : p.resolve(m.result);
      } else if (m.method === 'Fetch.requestPaused') {
        const {requestId, request} = m.params;
        const body = iconBodies.get(new URL(request.url).searchParams.get('id'));
        call('Fetch.fulfillRequest', {requestId, responseCode: body ? 200 : 404, body, responseHeaders: [{name: 'Content-Type', value: 'image/png'}]});
      }
    };
    const evalJS=async expression=>{const r=await call('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.text);return r.result.value};
    await call('Fetch.enable', {patterns: [{urlPattern: '*/api/icons/file?*'}]});
    await pause(500); await call('Page.navigate',{url:'http://127.0.0.1:18765/static/ticket.html'});
    for(let i=0;i<100 && !await evalJS("document.readyState === 'complete' && !!document.querySelector('#ticketSample')");i++) await pause(100);
    await evalJS("document.querySelector('#ticketSample').click()");
    const fields=()=>evalJS("Object.fromEntries(new FormData(document.querySelector('#ticketForm')))");
    const expected={ticket_id:'55345',title:'TBSNG0000_手書きサイン',project:'開発（UI）',parent_project:'介舟ファミリークラウド総合',parent:'55022',parent_category:'機能追加',parent_title:'介舟ファミリーのタブレット対応／計画書の電子サイン対応（正式版）',status:'進行中',priority:'通常',assignee:'（開発）白 家綾（現）',start_date:'2026/11/20',due_date:'2026/12/03',progress:'0%',version:'介舟ファミリークラウド総合 - V440',screen_id:'TBSNG0000',estimated_hours:'9',actual_hours:'1'};
    const check=async()=>{const result=await fields();for(const [key,value] of Object.entries(expected))assert.equal(result[key],value,key)};
    await check();
    assert.equal((await fields()).assignee_short, '白');
    await evalJS("document.querySelector('[data-hours-for=fEst][data-hours-step=\"1\"]').click()");
    assert.equal((await fields()).estimated_hours, '10');
    await evalJS("document.querySelector('[data-hours-for=fEst][data-hours-step=\"-1\"]').click()");
    assert.equal((await fields()).estimated_hours, '9');
    await evalJS("(()=>{const e=document.querySelector('#fAct'); e.value='1.125'; e.dispatchEvent(new Event('input',{bubbles:true})); document.querySelector('[data-hours-for=fAct][data-hours-step=\"1\"]').click()})()");
    assert.equal((await fields()).actual_hours, '2.125');
    assert.equal(await evalJS("(()=>{const e=document.querySelector('#fAct'); e.focus(); return !e.dispatchEvent(new WheelEvent('wheel',{deltaY:-100,bubbles:true,cancelable:true}))})()"), true);
    assert.equal((await fields()).actual_hours, '3.125');
    assert.equal(await evalJS("(()=>{const e=document.querySelector('#fAct'); e.blur(); return e.dispatchEvent(new WheelEvent('wheel',{deltaY:100,bubbles:true,cancelable:true}))})()"), true);
    assert.equal((await fields()).actual_hours, '3.125');
    for (const expectedProgress of ['50%', '70%', '100%', '100%']) {
      await evalJS("(()=>{const e=document.querySelector('#fProgress');e.focus();e.dispatchEvent(new WheelEvent('wheel',{deltaY:-100,bubbles:true,cancelable:true}))})()");
      assert.equal((await fields()).progress, expectedProgress);
    }
    assert.equal(await evalJS("document.querySelector('#stubFill').style.width"), '100%');
    await evalJS("document.querySelector('#fProgress').dispatchEvent(new WheelEvent('wheel',{deltaY:100,bubbles:true,cancelable:true}))");
    assert.equal((await fields()).progress, '70%');
    await evalJS("document.querySelector('#ticketReset').click()"); await pause(100);
    await evalJS("document.querySelector('#ticketSample').click()");
    assert.equal(await evalJS("document.querySelector('#boardCount').textContent"), '17 / 17');
    assert.equal(await evalJS("!!document.querySelector('[name=activity]')"), false);
    if(process.argv[2]) {
      const source=fs.readFileSync(process.argv[2],'utf8').split('```html')[1].split('```')[0];
      await evalJS(`(()=>{const e=document.querySelector('#ticketSource');e.value=${JSON.stringify(source)};e.dispatchEvent(new Event('input',{bubbles:true}))})()`);
      await check();
    }
    await evalJS(`(() => {
      window.testIcons = ${JSON.stringify(icons)};
      window.iconResult = 'failure';
      window.testStructures = [{id: 'st0001', name: '基本設計', tree: [
        {id: 'n1', kind: 'dir', name: '01_要件', children: [
          {id: 'n2', kind: 'file', name: 'ビュー定義書.xlsx', source: 'D:/模板/ビュー定義書.xlsx'},
        ]},
        {id: 'n3', kind: 'dir', name: '02_設計', children: []},
      ]}];
      window.structureResult = 'applied';
      window.fetch = async (url, options) => {
        if (url === '/api/icons') {
          return Response.json({ok: true, icons: window.testIcons});
        }
        if (url === '/api/ticket_structures') {
          return Response.json({ok: true, templates: window.testStructures});
        }
        if (url === '/api/ticket/open') {
          window.openedFolder = JSON.parse(options.body).folder;
          return window.folderAvailable ? Response.json({ok: true, foreground: window.folderForeground})
            : Response.json({ok: false, error: '文件夹不存在，可能已被移动或删除'}, {status: 400});
        }
        window.savedTicket = JSON.parse(options.body);
        const result = {ok: true, folder: 'Desktop/test'};
        if (window.savedTicket.icon_id && window.iconResult === 'failure') result.icon_error = '图标写入失败';
        if (window.savedTicket.icon_id && window.iconResult === 'applied') result.icon_id = window.savedTicket.icon_id;
        if (window.savedTicket.structure_id && window.structureResult === 'applied') {
          result.structure = {template: '基本設計', dirs: 1, files: 1, problems: [{path: '02_設計/欠.txt', reason: '来源文件不存在'}], items: []};
        }
        if (window.savedTicket.structure_id && window.structureResult === 'error') result.structure_error = '模板已删除';
        return Response.json(result);
      };
      document.querySelector('#ticketIconRefresh').click();
      document.querySelector('#ticketStructureRefresh').click();
    })()`);
    assert.equal((await fields()).icon_id, '');
    assert.equal(await evalJS("document.querySelectorAll('#ticketIconGrid img').length"), icons.length);
    assert.equal(await evalJS("document.querySelector('#ticketIconGrid img').getAttribute('src')"), '/api/icons/file?id=123ab0&kind=png&v=2');
    await evalJS("document.querySelector('#ticketIconGrid').scrollIntoView({block: 'center', behavior: 'instant'})");
    await evalJS("Promise.all([...document.querySelectorAll('#ticketIconGrid img')].map(image => image.decode()))");
    assert.deepEqual(await evalJS("[...document.querySelectorAll('#ticketIconGrid img')].map(image => [image.naturalWidth > 0, image.getBoundingClientRect().width])"), icons.map(() => [true, 46]));
    await evalJS("document.querySelector('#ticketIconGrid input:checked').focus()");
    await call('Input.dispatchKeyEvent', {type: 'keyDown', key: 'ArrowRight', code: 'ArrowRight', windowsVirtualKeyCode: 39});
    await call('Input.dispatchKeyEvent', {type: 'keyUp', key: 'ArrowRight', code: 'ArrowRight', windowsVirtualKeyCode: 39});
    assert.equal((await fields()).icon_id, '123ab0');
    await evalJS(`document.querySelector('#ticketIconGrid input[value="123ab1"]').labels[0].click()`);
    assert.equal((await fields()).icon_id, '123ab1');
    await evalJS("document.querySelector('#ticketIconRefresh').click()");
    assert.equal((await fields()).icon_id, '123ab1');
    if (process.env.TICKET_SCREENSHOT) { const shot = await call('Page.captureScreenshot', {captureBeyondViewport:true}); fs.writeFileSync(process.env.TICKET_SCREENSHOT, Buffer.from(shot.data, 'base64')); }
    await evalJS("document.querySelector('#ticketForm').requestSubmit()");
    assert.equal(await evalJS('window.savedTicket.icon_id'), '123ab1');
    assert.match(await evalJS("document.querySelector('#ticketIconWarning').textContent"), /图标写入失败/);
    assert.equal(await evalJS("document.querySelector('#ticketSuccess').hidden"), false);
    assert.equal(await evalJS("document.querySelector('#ticketIconApplied').hidden"), true);
    await evalJS("window.iconResult = 'old server'; document.querySelector('#ticketForm').requestSubmit()");
    assert.match(await evalJS("document.querySelector('#ticketIconWarning').textContent"), /服务未确认应用图标/);
    await evalJS("window.iconResult = 'applied'; document.querySelector('#ticketForm').requestSubmit()");
    assert.equal(await evalJS("document.querySelector('#ticketIconApplied').hidden"), false);
    assert.equal(await evalJS("document.querySelector('#ticketIconWarning').textContent"), '');
    await evalJS("document.querySelector('#openFolderBtn').click()");
    assert.equal(await evalJS("window.openedFolder"), 'Desktop/test');
    assert.match(await evalJS("document.querySelector('#ticketSuccess #ticketOpenError').textContent"), /文件夹不存在/);
    assert.equal(await evalJS("document.querySelector('#ticketError').textContent"), '');
    await evalJS("window.folderAvailable = true; window.folderForeground = false; document.querySelector('#openFolderBtn').click()");
    assert.equal(await evalJS("document.querySelector('#ticketOpenError').textContent"), '');
    assert.match(await evalJS("document.querySelector('#ticketOpenStatus').textContent"), /任务栏/);
    await evalJS("window.folderForeground = true; document.querySelector('#openFolderBtn').click()");
    assert.equal(await evalJS("document.querySelector('#ticketOpenStatus').textContent"), '');
    // ---------- 目录结构模板：选择、预览微调、出票反馈 ----------
    assert.equal(await evalJS("document.querySelector('#ticketStructurePreview').hidden"), true);
    await evalJS("(()=>{const e=document.querySelector('#ticketStructureSelect');e.value='st0001';e.dispatchEvent(new Event('change',{bubbles:true}))})()");
    assert.equal((await fields()).structure_id, 'st0001');
    assert.equal(await evalJS("document.querySelector('#ticketStructurePreview').hidden"), false);
    assert.equal(await evalJS("document.querySelectorAll('#ticketStructureTree .ticket-structure-row').length"), 3);
    await evalJS("document.querySelectorAll('#ticketStructureTree > li > .ticket-structure-row input[type=checkbox]')[1].click()");
    assert.equal(await evalJS("document.querySelectorAll('#ticketStructureTree > li')[1].classList.contains('is-excluded')"), true);
    await evalJS("document.querySelector('#ticketStructureTree > li > .ticket-structure-row input[type=checkbox]').click()");
    assert.equal(await evalJS("document.querySelector('#ticketStructureTree .ticket-structure-tree input[type=checkbox]').disabled"), true);
    assert.equal(await evalJS("document.querySelector('#ticketStructureTree .ticket-structure-tree .ticket-structure-name').disabled"), true);
    await evalJS("document.querySelector('#ticketStructureTree > li > .ticket-structure-row input[type=checkbox]').click()");
    await evalJS(`(()=>{const e=document.querySelector('#ticketStructureTree > li .ticket-structure-name');e.value='01_要件定義';e.dispatchEvent(new Event('input',{bubbles:true}))})()`);
    await evalJS("window.structureResult = 'applied'; document.querySelector('#ticketForm').requestSubmit()");
    assert.equal(await evalJS('window.savedTicket.structure_id'), 'st0001');
    assert.deepEqual(await evalJS('window.savedTicket.structure_overrides'), {exclude: ['n3'], rename: {n1: '01_要件定義'}});
    assert.match(await evalJS("document.querySelector('#ticketStructureApplied').textContent"), /基本設計.*1 个目录.*1 个文件/);
    assert.equal(await evalJS("document.querySelector('#ticketStructureApplied').hidden"), false);
    assert.match(await evalJS("document.querySelector('#ticketStructureWarning').textContent"), /没建成/);
    assert.equal(await evalJS("document.querySelectorAll('#ticketStructureProblems li').length"), 1);
    assert.match(await evalJS("document.querySelector('#ticketStructureProblems li').textContent"), /02_設計\/欠\.txt/);
    await evalJS("window.structureResult = 'error'; document.querySelector('#ticketForm').requestSubmit()");
    assert.match(await evalJS("document.querySelector('#ticketStructureWarning').textContent"), /组建失败.*模板已删除/);
    await evalJS("window.structureResult = 'old server'; document.querySelector('#ticketForm').requestSubmit()");
    assert.match(await evalJS("document.querySelector('#ticketStructureWarning').textContent"), /服务未确认组建目录结构/);
    await evalJS("document.querySelector('#ticketReset').click()"); await pause(100);
    assert.equal((await fields()).structure_id, '');
    assert.equal(await evalJS("document.querySelector('#ticketStructurePreview').hidden"), true);
    await evalJS("document.querySelector('#ticketSample').click()");
    await evalJS("document.querySelector('#ticketForm').requestSubmit()");
    assert.equal(await evalJS("Object.hasOwn(window.savedTicket, 'structure_id')"), false);
    assert.equal(await evalJS("Object.hasOwn(window.savedTicket, 'structure_overrides')"), false);
    assert.equal(await evalJS("document.querySelector('#ticketStructureApplied').hidden"), true);
    assert.equal(await evalJS("document.querySelector('#ticketStructureWarning').textContent"), '');
    assert.equal(await evalJS('window.savedTicket.parent_title'), expected.parent_title);
    assert.equal(await evalJS('window.savedTicket.parent_project'), expected.parent_project);
    assert.equal(await evalJS('window.savedTicket.assignee_short'), '白');
    const setAssignee = async value => evalJS(`(()=>{const e=document.querySelector('[name=assignee]');e.value=${JSON.stringify(value)};e.dispatchEvent(new Event('input',{bubbles:true}))})()`);
    await setAssignee('吴振科'); assert.equal((await fields()).assignee_short, '吴');
    await setAssignee('（開発）張 曉靜（現）'); assert.equal((await fields()).assignee_short, '静');
    await setAssignee('其他人'); assert.equal((await fields()).assignee_short, '');
    await evalJS("(()=>{const e=document.querySelector('[name=assignee_short]');e.value='自';e.dispatchEvent(new Event('input',{bubbles:true}))})()");
    await setAssignee('白家绫'); assert.equal((await fields()).assignee_short, '自');
    await evalJS("(()=>{const e=document.querySelector('[name=assignee]');e.value='手动';e.dispatchEvent(new Event('input',{bubbles:true}));document.querySelector('#ticketSample').click()})()");
    assert.equal((await fields()).assignee,'手动');
    await evalJS("document.querySelector('#ticketReset').click()"); await pause(100);
    assert.ok(Object.values(await fields()).every(x=>x===''));
    assert.equal(await evalJS("document.querySelector('#ticketIconGrid input:checked').value"), '');
    await evalJS(`document.querySelector('#ticketIconGrid input[value="123ab1"]').labels[0].click(); window.testIcons = []; document.querySelector('#ticketIconRefresh').click()`);
    assert.equal((await fields()).icon_id, '');
    assert.match(await evalJS("document.querySelector('#ticketIconStatus').textContent"), /图标已从图标库移除/);
    await evalJS("(()=>{const e=document.querySelector('#ticketStructureSelect');e.value='st0001';e.dispatchEvent(new Event('change',{bubbles:true}))})()");
    await evalJS("window.testStructures = []; document.querySelector('#ticketStructureRefresh').click()");
    assert.equal((await fields()).structure_id, '');
    assert.match(await evalJS("document.querySelector('#ticketStructureStatus').textContent"), /模板已删除/);
    assert.equal(await evalJS("document.querySelector('#ticketStructurePreview').hidden"), true);
    await evalJS(`(()=>{const e=document.querySelector('#ticketSource');e.value='<div id="content"><h2>task #7</h2><div class="issue details"><div class="subject"><h3>Only title</h3></div><table class="attributes"><tr><td class="status"></td></tr></table></div><div id="history">status: old</div></div>';e.dispatchEvent(new Event('input',{bubbles:true}))})()`);
    const partial=await fields(); assert.equal(partial.ticket_id,'7'); assert.equal(partial.title,'Only title'); assert.equal(partial.status,''); assert.equal(partial.parent_project,'');
    await evalJS("document.querySelector('#ticketForm').requestSubmit()");
    assert.equal(await evalJS("Object.hasOwn(window.savedTicket, 'icon_id')"), false);
    assert.equal(await evalJS("Object.hasOwn(window.savedTicket, 'structure_id')"), false);
    assert.equal(await evalJS("document.querySelector('#ticketIconApplied').hidden"), true);
    assert.equal(await evalJS("document.querySelector('#ticketIconWarning').textContent"), '');
    console.log('Ticket recognition, icon and structure selection, creation and reset passed.');
  } finally {if(socket)socket.close(); edge.kill();server.kill();}
})().catch(e=>{console.error(e);process.exitCode=1});
