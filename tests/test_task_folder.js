const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../static/task_folder.js'), 'utf8');
const settle = () => new Promise(resolve => setImmediate(resolve));
const task = { id:'one', name:'任务一', work:'C:/task/work', reference:'C:/task/reference' };

function page({work = '', reference = '', data = {ok:true, current:'one', tasks:[task]}, paths, deferred = false, fail = false} = {}) {
  const elements = new Map();
  function element(value = '') {
    return { value, readOnly:true, placeholder:'原提示', classList:{remove(){}}, parentElement:null,
      closest(){return null;}, events:[], dispatchEvent(event){this.events.push(event.type);},
      insertAdjacentElement(where, node){elements.set(node.id,node);} };
  }
  elements.set('taskFolderInput', element(work));
  elements.set('repairOldFolder', element(reference));
  const requests = [];
  const window = {};
  const context = vm.createContext({ window, document:{getElementById:id => elements.get(id),
    createElement:() => element(), head:{appendChild(){}}, body:{}},
    Event:class {constructor(type){this.type=type;}}, getComputedStyle:() => ({display:'block'}),
    fetch:async url => {
      if(fail) throw Error('offline');
      if (url.startsWith('/api/workbook-paths?')) return {json:async () => ({ok:true, paths})};
      if(deferred) return new Promise(resolve => requests.push(value => resolve({json:async () => value})));
      return {json:async () => data};
    } });
  vm.runInContext(source, context);
  return {window, elements, requests, setData(value){data=value;}};
}

test('task paths fill empty inputs and stay editable', async () => {
  const p=page();
  assert.equal(p.elements.get('taskFolderInput').readOnly,false);
  await settle();
  assert.equal(p.elements.get('taskFolderInput').value,task.work);
  assert.equal(p.elements.get('repairOldFolder').value,task.reference);
  assert.equal(p.elements.get('repairOldFolder').readOnly,false);
  assert.match(p.elements.get('taskFolderNote').innerHTML,/可以直接修改/);
});

test('loading a current task keeps independently chosen work and reference paths', async () => {
  const p=page({work:'C:/chosen',reference:'D:/backup'});
  await settle();
  assert.equal(p.elements.get('taskFolderInput').value,'C:/chosen');
  assert.equal(p.elements.get('repairOldFolder').value,'D:/backup');
});

test('a saved batch displays editable source paths on separate lines', async () => {
  const scope = `C:\\data\\批量路径\\${'a'.repeat(24)}`;
  const paths = ['C:\\chosen\\book.xlsx', 'D:\\more'];
  const p = page({work:scope, paths});
  await settle();
  const input = p.elements.get('taskFolderInput');
  assert.equal(input.value, paths.join('\n'));
  assert.equal(input.readOnly, false);
});

test('no task is required, and saved paths and placeholders are preserved', async () => {
  const p=page({work:'C:/chosen',data:{ok:true,tasks:[],current:''}});
  await settle();
  assert.equal(p.elements.get('taskFolderInput').value,'C:/chosen');
  assert.equal(p.elements.get('repairOldFolder').placeholder,'原提示');
  assert.equal(p.elements.get('taskFolderInput').readOnly,false);
  assert.doesNotMatch(p.elements.get('taskFolderNote').innerHTML,/先.*新任务|锁定/);
});

test('paths entered while task information loads are not overwritten', async () => {
  const p=page({deferred:true});
  p.elements.get('taskFolderInput').value='C:/typed-before-response';
  p.requests[0]({ok:true,current:'one',tasks:[task]});
  await settle();
  assert.equal(p.elements.get('taskFolderInput').value,'C:/typed-before-response');
  assert.equal(p.elements.get('repairOldFolder').value,task.reference);
});

test('choosing task defaults fills both paths, then allows further edits', async () => {
  const p=page({work:'C:/custom',reference:'D:/custom'});
  await settle();
  await p.window.refreshTaskFolder();
  assert.equal(p.elements.get('taskFolderInput').value,task.work);
  assert.equal(p.elements.get('repairOldFolder').value,task.reference);
  assert.equal(p.elements.get('taskFolderInput').readOnly,false);
});

test('a failed task request leaves local paths usable', async () => {
  const p=page({work:'C:/custom',fail:true});
  await settle();
  assert.equal(p.elements.get('taskFolderInput').value,'C:/custom');
  assert.equal(p.elements.get('taskFolderInput').readOnly,false);
});

test('self-chosen preference does not force task defaults into empty inputs', async () => {
  const p=page({data:{ok:true,current:'one',tasks:[task],single:true}});
  await settle();
  assert.equal(p.elements.get('taskFolderInput').value,'');
  assert.equal(p.elements.get('taskFolderInput').readOnly,false);
});

test('paths left over from another task switch to the current task, custom paths stay', async () => {
  const old = { id:'two', name:'任务二', work:'C:\\old\\work', reference:'C:\\old\\reference' };
  const p=page({work:'c:/OLD/work/',reference:'D:/custom',data:{ok:true,current:'one',tasks:[task,old]}});
  await settle();
  assert.equal(p.elements.get('taskFolderInput').value,task.work);
  assert.equal(p.elements.get('repairOldFolder').value,'D:/custom');
});
