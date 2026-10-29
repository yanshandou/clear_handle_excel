const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function elements() {
  const items = {};
  const get = id => items[id] ||= {value:'', innerHTML:'', textContent:'', style:{}, events:{}, open:true,
    addEventListener(name, fn) { this.events[name] = fn; }, contains() { return false; },
    querySelector() { return {focus() {}}; }};
  return get;
}
const flush = () => new Promise(resolve => setImmediate(resolve));
const script = name => fs.readFileSync(path.join(__dirname, '../static', name), 'utf8');

test('BL filter changes display but clear still covers all hits', async () => {
  const get = elements(), requests = [], prompts = [];
  const data = {folder:'current', files:[{name:'hit.xlsx', hits:[{sheet:'S', addr:'A1', left_addr:'B1'}]},
    {name:'empty.xlsx', hits:[]}, {name:'failed.xlsx', error:'失败'}]};
  const context = vm.createContext({document:{getElementById:get}, localStorage:{getItem(){},setItem(){}},
    window:{excelLinkHtml:label => label, appConfirm:async text => {prompts.push(text); return true;}},
    setTimeout(){}, clearTimeout(){}, fetch:async (url, options) => {requests.push({url, options});
      return {ok:true, json:async () => options?.method === 'POST' ? {ok:true} : {result:data}};}});
  vm.runInContext(script('bl_links.js'), context);
  await flush();
  get('blResults').events.click({target:{closest:() => ({dataset:{blFilter:'failed'}})}});
  assert.match(get('blResults').innerHTML, /failed.xlsx/);
  assert.doesNotMatch(get('blResults').innerHTML, /hit.xlsx|empty.xlsx/);
  await get('blClear').events.click();
  assert.match(prompts[0], /删除 1 处/);
  assert.equal(requests.find(r => r.url.endsWith('/clear')).options.body, '{}');
});

test('print layout search and categories preserve original row identity for updates', async () => {
  const get = elements(), requests = [];
  get('printLayoutsFilter').value = 'all';
  const rows = [{file:'first.xlsx', sheet:'A', path:'first', key:'a', selected:1, original:1},
    {file:'second.xlsx', sheet:'B', path:'second', key:'b', selected:2, original:1, stale:true, manual:2}];
  const context = vm.createContext({document:{getElementById:get, activeElement:null}, localStorage:{getItem(){}}, setInterval(){},
    fetch:async (url, options) => {requests.push({url, options}); return {ok:true, json:async () => options?.method === 'POST' ? {ok:true} : {rows}};}});
  vm.runInContext(script('print_layout.js'), context);
  await get('printLayoutsDetails').events.toggle();
  get('printLayoutsFilter').value = 'stale';
  get('printLayoutsFilter').events.change();
  assert.match(get('printLayoutsList').innerHTML, /second.xlsx/);
  assert.doesNotMatch(get('printLayoutsList').innerHTML, /first.xlsx/);
  assert.match(get('printLayoutsList').innerHTML, /data-layout-index="1"/);
  get('printLayoutsList').events.change({target:{closest:() => ({dataset:{layoutIndex:'1'}, value:'portrait'})}});
  await flush();
  assert.equal(JSON.parse(requests.find(r => r.options?.method === 'POST').options.body).path, 'second');
  get('printLayoutsSearch').value = 'missing';
  get('printLayoutsSearch').events.input();
  assert.match(get('printLayoutsList').innerHTML, /没有符合筛选条件/);
});

