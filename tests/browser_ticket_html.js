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
    socket.onmessage=e=>{const m=JSON.parse(e.data);if(m.id && pending.has(m.id)){const p=pending.get(m.id);pending.delete(m.id);m.error?p.reject(m.error):p.resolve(m.result)}};
    const call=(method,params={})=>new Promise((resolve,reject)=>{const n=++id;pending.set(n,{resolve,reject});socket.send(JSON.stringify({id:n,method,params}))});
    const evalJS=async expression=>{const r=await call('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(r.exceptionDetails)throw Error(r.exceptionDetails.text);return r.result.value};
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
    if (process.env.TICKET_SCREENSHOT) { const shot = await call('Page.captureScreenshot', {captureBeyondViewport:true}); fs.writeFileSync(process.env.TICKET_SCREENSHOT, Buffer.from(shot.data, 'base64')); }
    if(process.argv[2]) {
      const source=fs.readFileSync(process.argv[2],'utf8').split('```html')[1].split('```')[0];
      await evalJS(`(()=>{const e=document.querySelector('#ticketSource');e.value=${JSON.stringify(source)};e.dispatchEvent(new Event('input',{bubbles:true}))})()`);
      await check();
    }
    await evalJS(`(()=>{window.fetch=async (url, options)=>{window.savedTicket=JSON.parse(options.body); return Response.json({ok:true,folder:'Desktop/test'})}; document.querySelector('#ticketForm').requestSubmit()})()`);
    await pause(100);
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
    await evalJS(`(()=>{const e=document.querySelector('#ticketSource');e.value='<div id="content"><h2>task #7</h2><div class="issue details"><div class="subject"><h3>Only title</h3></div><table class="attributes"><tr><td class="status"></td></tr></table></div><div id="history">status: old</div></div>';e.dispatchEvent(new Event('input',{bubbles:true}))})()`);
    const partial=await fields(); assert.equal(partial.ticket_id,'7'); assert.equal(partial.title,'Only title'); assert.equal(partial.status,''); assert.equal(partial.parent_project,'');
    console.log('Ticket HTML recognition, sample, manual edits and reset passed.');
  } finally {if(socket)socket.close(); edge.kill();server.kill();}
})().catch(e=>{console.error(e);process.exitCode=1});
