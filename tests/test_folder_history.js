const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../static/history.js'), 'utf8');
const settle = () => new Promise(resolve => setImmediate(resolve));

async function page({ single = false, task = true, busy = false, failure = false } = {}) {
  const elements = new Map(), requests = [], values = new Map();
  const element = id => {
    if (!elements.has(id)) elements.set(id, { value:'', innerHTML:'', textContent:'', listeners:{}, addEventListener(name, fn) { this.listeners[name] = fn; } });
    return elements.get(id);
  };
  const folders = ['C:/A', 'C:/B'];
  const context = vm.createContext({
    document:{ getElementById:element }, window:{ location:{ href:'/history' } },
    localStorage:{ setItem(key, value) { values.set(key, value); } },
    fetch:async (url, options) => {
      requests.push({ url, options });
      let value;
      if (options?.method === 'POST') {
        if (failure) return { ok:false, status:400, json:async () => ({error:'当前任务正在执行，请先暂停'}) };
        value = {ok:true, task:{reference:'C:/reference'}};
      } else if (url === '/api/folders') value = {current:'C:/B', busy, folders:folders.map(folder => ({folder, issues:1, files:2, saved_at:'2026-11-05', status:'interrupted'}))};
      else if (url === '/api/tasks') value = {single, tasks: task ? [{id:'taskA', work:'C:/A'}] : []};
      else {
        const folder = new URL(url, 'http://localhost').searchParams.get('folder');
        value = { summary:{task_name:'扫描', status:'interrupted'}, report:{files:{'book.xlsx':{name:'book.xlsx', status:'needs_fix', error:'<script>bad</script>',
          deleted_backups:['旧表<img>'], fixes:{one:{sheet:'S',cell:'A1',reason:'截断',before:'before.png'}}}}},
          logs:[folder + ' log'], operations:[{name:'扫描',status:'completed',started_at:'2026-11-05'}], scans:[] };
      }
      return {ok:true, json:async () => value};
    }, URL, encodeURIComponent
  });
  vm.runInContext(source, context);
  await settle();
  return {elements, requests, values, context, element, async select(folder) {
    await element('folders').listeners.click({target:{closest:() => ({dataset:{folder}})}});
    await settle();
  }};
}

test('browsing history is read only, displays file details and escapes content', async () => {
  const p = await page();
  await p.select('C:/A');
  assert.match(p.element('detail').innerHTML, /book.xlsx/);
  assert.match(p.element('detail').innerHTML, /&lt;script&gt;bad&lt;\/script&gt;/);
  assert.match(p.element('detail').innerHTML, /C:\/A log/);
  assert.match(p.element('detail').innerHTML, /旧表&lt;img&gt;/);
  assert.match(p.element('detail').innerHTML, /\/api\/batch\/image\/before.png/);
  assert.ok(p.requests.every(r => r.options?.method !== 'POST'));
  assert.equal(p.context.window.location.href, '/history');
});

test('restoring a task switches to its saved workbench without starting a scan', async () => {
  const p = await page();
  await p.select('C:/A');
  await p.element('restore').listeners.click();
  const posts = p.requests.filter(r => r.options?.method === 'POST');
  assert.equal(posts.length, 1);
  assert.equal(posts[0].url, '/api/tasks/switch');
  assert.equal(JSON.parse(posts[0].options.body).id, 'taskA');
  assert.equal(p.values.get('clear_handle_excel.lastFolder'), 'C:/A');
  assert.equal(p.context.window.location.href, '/workbench');
});

test('custom folder progress can be restored without a task or a mode switch', async () => {
  const p = await page({single:false, task:false});
  await p.select('C:/A');
  assert.doesNotMatch(p.element('detail').innerHTML, /id="restore"[^>]*disabled|先到主页打开单项模式/);
  await p.element('restore').listeners.click();
  const posts = p.requests.filter(r => r.options?.method === 'POST');
  assert.equal(posts[0].url, '/api/folders/activate');
  assert.equal(JSON.parse(posts[0].options.body).folder, 'C:/A');
});

test('busy backend disables restore, while history remains readable', async () => {
  const p = await page({busy:true});
  await p.select('C:/A');
  assert.match(p.element('detail').innerHTML, /id="restore"[^>]*disabled/);
  assert.match(p.element('detail').innerHTML, /完成或暂停/);
  assert.match(p.element('detail').innerHTML, /book.xlsx/);
});

test('failed restore reports the error and keeps the history page', async () => {
  const p = await page({failure:true});
  await p.select('C:/A');
  await p.element('restore').listeners.click();
  assert.match(p.element('error').textContent, /请先暂停/);
  assert.equal(p.context.window.location.href, '/history');
});
