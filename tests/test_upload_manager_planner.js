const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const elements = new Map();
function elementFor(selector) {
  if (!elements.has(selector)) elements.set(selector, {
    handlers: {}, addEventListener(name, handler) { this.handlers[name] = handler; },
    classList: {toggle() {}}, after() {}
  });
  return elements.get(selector);
}
const context = {$: elementFor, document: {handlers: {}, addEventListener(name, handler) { (this.handlers[name] ||= []).push(handler); }}, items: [], Blob, atob, escapeHTML: value => String(value).replace(/[&<>"']/g, character => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[character]))};
const appSource = fs.readFileSync(path.join(__dirname, '../static/upload_manager/app.js'), 'utf8');
vm.runInNewContext(appSource.slice(appSource.indexOf('const folderStatuses'), appSource.indexOf('function percentOf(')), context);
vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../static/upload_manager/planner.js'), 'utf8'), context);

const today = new Date();
assert.equal(elementFor('#plan-date').value, `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`);
assert.equal(elementFor('#plan-people-count').textContent, '已选 2');
assert.ok(elementFor('#plan-people-summary').innerHTML.includes('person-picker-summary-chip'));
context.items = [{owner:'<测试>'}];
context.renderPersonPicker('people');
assert.ok(elementFor('#plan-people-options').innerHTML.includes('person-picker-chip'));
assert.ok(elementFor('#plan-people-options').innerHTML.includes('&lt;测试&gt;'));
assert.ok(!elementFor('#plan-people-options').innerHTML.includes('<测试>'));
const pickerColors = context.allocationPersonMarkers();
assert.notEqual(pickerColors.get('呉').hue, pickerColors.get('静').hue);
assert.ok(elementFor('#plan-people-options').innerHTML.includes(`--person-hue:${pickerColors.get('呉').hue}`));
assert.ok(elementFor('#plan-people-summary').innerHTML.includes(`--person-hue:${pickerColors.get('呉').hue}`));
context.items = [];
context.renderPersonPicker('owner');
assert.equal(elementFor('#plan-owner-count').textContent, '全部');
elementFor('#plan-people-options').handlers.change({target:{type:'checkbox', checked:true, value:'新人员'}});
assert.equal(elementFor('#plan-people-count').textContent, '已选 3');
assert.ok(elementFor('#plan-people-summary').innerHTML.includes('新人员'));
assert.equal((elementFor('#plan-people-summary').innerHTML.match(/class="person-picker-summary-chip/g) || []).length, 3);
assert.ok(!elementFor('#plan-people-summary').innerHTML.includes('person-picker-summary-more'));
assert.ok(elementFor('#plan-people-summary').innerHTML.includes('class="sr-only"'));
elementFor('#plan-people-options').handlers.change({target:{type:'checkbox', checked:false, value:'新人员'}});
assert.equal(elementFor('#plan-people-count').textContent, '已选 2');
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:true, value:''}});
assert.equal(elementFor('#plan-owner-count').textContent, '已选 1');
assert.ok(elementFor('#plan-owner-summary').innerHTML.includes('未分配'));
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:false, value:''}});
assert.equal(elementFor('#plan-owner-count').textContent, '全部');
context.items = [{owner:'呉'}];
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:true, value:'呉'}});
context.clearPlanPreview = () => {};
elementFor('#plan-person-add').value = '临时人员';
context.addPlanPerson();
const updatedMarkers = context.allocationPersonMarkers();
assert.ok(elementFor('#plan-people-summary').innerHTML.includes(`--person-hue:${updatedMarkers.get('呉').hue}`));
assert.ok(elementFor('#plan-owner-summary').innerHTML.includes(`--person-hue:${updatedMarkers.get('呉').hue}`));
elementFor('#plan-people-options').handlers.change({target:{type:'checkbox', checked:false, value:'临时人员'}});
assert.equal(context.allocationPersonMarkers().get('呉').hue, updatedMarkers.get('呉').hue);
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:false, value:'呉'}});
context.items.push({owner:'临时来源'});
context.renderPersonPicker('owner');
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:true, value:'临时来源'}});
const sourceMarkers = context.allocationPersonMarkers();
context.items = [{owner:'呉'}];
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:false, value:'临时来源'}});
assert.equal(context.allocationPersonMarkers().get('呉').hue, sourceMarkers.get('呉').hue);
const uploadedRecords = [
  {id:'a', group:'10要件定義\\10画面', name:'02_SG', total:6, uploaded:6, uploader:'呉', owner:'呉', date:'2026-10-20', parentOnly:false},
  {id:'b', group:'10要件定義¥10画面', name:'05_IR', total:22, uploaded:22, uploader:'静', owner:'呉', date:'', parentOnly:false},
  {id:'c', group:'独立', name:'独立', total:3, uploaded:3, uploader:'', owner:'', parentOnly:true}
];
const dateJson = JSON.parse(JSON.stringify(context.uploadDateJson('2026-10-29', uploadedRecords)));
assert.deepEqual(dateJson, {version:2, label:'2026-10-29 上传记录', uploadDate:'2026-10-29', people:[
  {name:'呉', folders:[{source:['10要件定義', '10画面', '02_SG'], destination:['呉', '10要件定義', '10画面', '02_SG'], owner:'呉', startDate:'2026-10-20', files:6, fileList:[]}]},
  {name:'静', folders:[{source:['10要件定義', '10画面', '05_IR'], destination:['静', '10要件定義', '10画面', '05_IR'], owner:'呉', startDate:'', files:22, fileList:[]}]},
  {name:'未分配', folders:[{source:['独立'], destination:['未分配', '独立'], owner:'', startDate:'', files:3, fileList:[]}]}
]});
const listedJson = context.uploadDateJson('2026-10-29', uploadedRecords, [['a.xlsx', '旧版/b.xls'], [], ['c.xlsx']]);
assert.deepEqual(listedJson.people[0].folders[0].fileList, ['a.xlsx', '旧版/b.xls']);
assert.deepEqual(Object.keys(listedJson.people[0].folders[0]), ['source', 'destination', 'owner', 'startDate', 'files', 'fileList']);
assert.deepEqual(Array.from(context.uploadDateSources(uploadedRecords)[2]), ['独立']);
const dateMarkup = context.dateExportMarkup('2026-10-29', uploadedRecords);
assert.ok(dateMarkup.includes('data-copy-date-json'));
assert.ok(dateMarkup.includes('3 个文件夹 · 31 个文件'));
assert.ok(!context.dateExportMarkup('2026-10-29', []).includes('data-copy-date-json'));
assert.ok(context.dateExportMarkup('2026-10-29', []).includes('date-history-empty'));
context.items = [];
context.window = {innerHeight:650};
const pickerPanel = {classList:{remove() {}, toggle(name, value) { this.above = value; }}, getBoundingClientRect() { return {height:350}; }};
const pickerList = {style:{}, getBoundingClientRect() { return {height:250}; }};
const positionedPicker = {querySelector(selector) { return selector === '.person-picker-panel' ? pickerPanel : pickerList; }, getBoundingClientRect() { return {top:300, bottom:340}; }};
context.positionPersonPicker(positionedPicker);
assert.equal(pickerPanel.classList.above, false);
assert.equal(pickerList.style.maxHeight, '195px');
positionedPicker.getBoundingClientRect = () => ({top:460, bottom:500});
context.positionPersonPicker(positionedPicker);
assert.equal(pickerPanel.classList.above, true);
assert.equal(pickerList.style.maxHeight, '250px');

function candidates(counts) {
  return counts.map((total, index) => ({id: String(index), name: String(index), total, date: '2026-01-01'}));
}
function totals(plans) {
  return plans.map(plan => plan.totals.reduce((sum, count) => sum + count, 0));
}

const exact = totals(context.generateAllocations(candidates([81, 80, 79]), 2, 160));
assert.equal(exact[0], 160);
assert.ok(exact.every(total => total >= 155 && total <= 165));
assert.ok(totals(context.generateAllocations(candidates([81, 80]), 2, 160)).every(total => total === 161));

const oversized = totals(context.generateAllocations(candidates([200, 90, 70]), 2, 160));
assert.ok(oversized.includes(160));
assert.ok(oversized.every(total => total >= 155 && total <= 165));

const belowLimit = totals(context.generateAllocations(candidates([40, 50]), 2, 160));
assert.equal(belowLimit.length, 0);
for (const total of [145, 155]) {
  const boundary = totals(context.generateAllocations(candidates([70, total - 70]), 2, 150));
  assert.ok(boundary.length > 0);
  assert.ok(boundary.every(value => value === total));
}
assert.equal(context.generateAllocations(candidates([70, 74]), 2, 150).length, 0);
assert.equal(context.generateAllocations(candidates([70, 86]), 2, 150).length, 0);
const bothSides = totals(context.generateAllocations(candidates([70, 75, 85]), 2, 150));
assert.ok(bothSides.includes(145) && bothSides.includes(155));
assert.ok(bothSides.every(total => total >= 145 && total <= 155));
const closestFirst = context.generateAllocations([
  {id:'a', name:'a', total:73, date:'2026-01-01', owner:'甲'},
  {id:'b', name:'b', total:77, date:'2026-01-01', owner:'乙'},
  {id:'c', name:'c', total:72, date:'2026-01-01', owner:'丙'},
  {id:'d', name:'d', total:74, date:'2026-01-01', owner:'丙'}
], 2, 150);
assert.equal(totals(closestFirst)[0], 150);
assert.ok(totals(closestFirst).every(total => total >= 145 && total <= 155));
assert.ok(totals(closestFirst).some(total => total !== 150));

const alternatives = context.generateAllocations(candidates([20, 20, 20]), 2, 60);
assert.equal(alternatives.length, 3);
assert.equal(new Set(alternatives.map(plan => plan.key)).size, 3);
assert.ok(alternatives.every(plan => plan.totals.every(total => total > 0) && plan.totals.reduce((sum, total) => sum + total, 0) >= 55 && plan.totals.reduce((sum, total) => sum + total, 0) <= 65));

const sourceOptions = context.generateAllocations([
  {id:'a', name:'a', total:20, date:'2026-01-01', owner:'甲'},
  {id:'b', name:'b', total:20, date:'2026-01-01', owner:'乙'},
  {id:'c', name:'c', total:20, date:'2026-01-01', owner:'甲'},
  {id:'d', name:'d', total:20, date:'2026-01-01', owner:'甲'}
], 2, 40);
assert.equal(sourceOptions.length, 3);
assert.equal(new Set(sourceOptions[0].assignments.map(entry => entry.item.owner)).size, 1);
assert.ok(sourceOptions.every((plan, index) => !index || sourceOptions[index - 1].ownerCount <= plan.ownerCount));

const diverseSources = context.generateAllocations([
  ...Array.from({length:25}, (_, index) => ({id:`a${index}`, name:`a${String(index).padStart(2, '0')}`, total:20, date:'2026-01-01', owner:'甲'})),
  {id:'b1', name:'b1', total:20, date:'2026-01-01', owner:'乙'},
  {id:'b2', name:'b2', total:20, date:'2026-01-01', owner:'乙'},
  {id:'b3', name:'b3', total:10, date:'2026-02-01', owner:'乙'}
], 2, 50);
assert.equal(diverseSources[0].ownerCount, 1);
assert.ok(diverseSources[0].assignments.every(entry => entry.item.owner === '乙'));

const ownerBeforeDate = context.generateAllocations([
  {id:'old-a', name:'old-a', total:40, date:'2026-01-01', owner:'甲'},
  {id:'old-b', name:'old-b', total:40, date:'2026-01-01', owner:'乙'},
  {id:'new-a', name:'new-a', total:40, date:'2026-02-01', owner:'丙'},
  {id:'new-b', name:'new-b', total:40, date:'2026-02-01', owner:'丙'}
], 2, 80);
assert.ok(ownerBeforeDate.length > 0);
assert.equal(ownerBeforeDate[0].ownerCount, 1);
assert.ok(ownerBeforeDate[0].assignments.every(entry => entry.item.owner === '丙'));

const ownerBeforeTotal = context.generateAllocations([
  {id:'a1', name:'a1', total:20, date:'2026-01-01', owner:'甲'},
  {id:'a2', name:'a2', total:20, date:'2026-01-01', owner:'甲'},
  {id:'b', name:'b', total:60, date:'2026-01-01', owner:'乙'},
  {id:'c', name:'c', total:40, date:'2026-01-01', owner:'丙'}
], 2, 100);
assert.equal(totals(ownerBeforeTotal)[0], 100);
assert.equal(ownerBeforeTotal[0].ownerCount, 2);

const parentBeforeDate = context.generateAllocations([
  {id:'old-a', name:'old-a', total:40, date:'2026-01-01', owner:'甲', group:'父一'},
  {id:'old-b', name:'old-b', total:40, date:'2026-01-01', owner:'甲', group:'父二'},
  {id:'new-a', name:'new-a', total:40, date:'2026-02-01', owner:'甲', group:'父三'},
  {id:'new-b', name:'new-b', total:40, date:'2026-02-01', owner:'甲', group:'父三'}
], 2, 80);
assert.ok(parentBeforeDate.length > 0);
assert.equal(parentBeforeDate[0].groupCount, 1);
assert.ok(parentBeforeDate[0].assignments.every(entry => entry.item.group === '父三'));

const parentBeforeTotal = context.generateAllocations([
  {id:'a1', name:'a1', total:20, date:'2026-01-01', owner:'甲', group:'父一'},
  {id:'a2', name:'a2', total:20, date:'2026-01-01', owner:'甲', group:'父一'},
  {id:'b', name:'b', total:60, date:'2026-01-01', owner:'甲', group:'父二'},
  {id:'c', name:'c', total:40, date:'2026-01-01', owner:'甲', group:'父三'}
], 2, 100);
assert.ok(parentBeforeTotal.length > 0);
assert.equal(totals(parentBeforeTotal)[0], 100);
assert.equal(parentBeforeTotal[0].groupCount, 2);

const ownerBeforeParent = context.generateAllocations([
  {id:'a', name:'a', total:20, date:'2026-01-01', owner:'甲', group:'父一'},
  {id:'b', name:'b', total:20, date:'2026-01-01', owner:'甲', group:'父二'},
  {id:'c', name:'c', total:20, date:'2026-01-01', owner:'乙', group:'父三'},
  {id:'d', name:'d', total:20, date:'2026-01-01', owner:'丙', group:'父三'}
], 2, 40);
assert.ok(ownerBeforeParent.length > 0);
assert.equal(ownerBeforeParent[0].ownerCount, 1);
assert.equal(ownerBeforeParent[0].groupCount, 2);

const dated = [
  {id:'old-a', name:'old-a', total:70, date:'2026-01-01'},
  {id:'old-b', name:'old-b', total:70, date:'2026-01-01'},
  {id:'new-a', name:'new-a', total:80, date:'2026-02-01'},
  {id:'new-b', name:'new-b', total:80, date:'2026-02-01'}
];
const early = context.generateAllocations(dated, 2, 160);
assert.ok(early.length > 0);
assert.ok(early.every(plan => plan.totals.reduce((sum, total) => sum + total, 0) === 160));
assert.ok(early.every(plan => plan.assignments.every(entry => entry.item.date === '2026-02-01')));
const withFiller = context.generateAllocations([...dated, {id:'new-small', name:'new-small', total:20, date:'2026-02-01'}], 2, 160);
assert.ok(withFiller.every(plan => plan.assignments.some(entry => entry.item.id === 'old-a') && plan.assignments.some(entry => entry.item.id === 'old-b')));
assert.ok(withFiller.every(plan => plan.totals.reduce((sum, total) => sum + total, 0) === 160));
const undated = context.generateAllocations([
  {id:'dated-a', name:'dated-a', total:40, date:'2026-01-01'},
  {id:'dated-b', name:'dated-b', total:40, date:'2026-01-01'},
  {id:'undated-a', name:'undated-a', total:50, date:''},
  {id:'undated-b', name:'undated-b', total:50, date:''}
], 2, 100);
assert.ok(undated.every(plan => plan.assignments.every(entry => !entry.item.date)));
const needsTwoPeople = context.generateAllocations([
  {id:'too-large', name:'too-large', total:160, date:'2026-01-01'},
  {id:'later-a', name:'later-a', total:78, date:'2026-02-01'},
  {id:'later-b', name:'later-b', total:82, date:'2026-02-01'}
], 2, 160);
assert.ok(needsTwoPeople.length > 0);
assert.ok(needsTwoPeople.every(plan => plan.assignments.every(entry => entry.item.id !== 'too-large')));

const historyMarkup = context.allocationHistoryMarkup([{id:'plan/1', label:'方案 <1>', plannedDate:'2026-11-02', createdAt:'2026-10-30T12:00:00', manifest:{people:[]}, hasArchive:true, assignments:[
  {group:'父<一>', name:'子&一', total:2, uploader:'甲<一>', owner:'来源甲', parentOnly:false},
  {group:'独立', name:'独立', total:1, uploader:'乙', owner:'', parentOnly:true}
]}]);
assert.ok(historyMarkup.includes('方案 &lt;1&gt;'));
assert.ok(historyMarkup.includes('甲&lt;一&gt;'));
assert.ok(historyMarkup.includes('父&lt;一&gt; / 子&amp;一'));
assert.ok(historyMarkup.includes('3 个文件 · 2 个文件夹 · 2 人'));
assert.ok(historyMarkup.includes('data-copy-history="plan/1"'));
assert.ok(historyMarkup.includes('/api/upload-manager/allocation-history/plan%2F1/archive'));
assert.ok(!historyMarkup.includes('父<一>'));
const jsonOnlyMarkup = context.allocationHistoryMarkup([{id:'new', label:'方案', plannedDate:'2026-11-02', createdAt:'2026-10-30T12:00:00', manifest:{people:[]}, hasArchive:false, assignments:[]}]);
assert.ok(jsonOnlyMarkup.includes('data-copy-history="new"'));
assert.ok(!jsonOnlyMarkup.includes('/archive'));
assert.ok(context.allocationHistoryMarkup([]).includes('暂无已保存的方案'));

context.isComplete = item => item.complete;
context.items = [
  {id: 'first', group: 'A', owner: '甲', total: 40, processed:40, date:'2026-10-30', uploaded: 0, complete: true},
  {id: 'second', group: 'B', owner: '乙', total: 50, processed:50, date:'2026-10-30', uploaded: 0, complete: true},
  {id: 'incomplete', group: 'A', owner: '甲', total: 20, uploaded: 0, complete: false},
  {id: 'uploaded', group: 'B', owner: '乙', total: 20, uploaded: 20, complete: true},
  {id: 'partial', group: 'B', owner: '乙', total: 20, uploaded: 10, complete: true},
  {id: 'empty', group: 'A', owner: '甲', total: 0, uploaded: 0, complete: true},
  {id: 'planned', group: 'A', owner: '甲', total: 20, processed:20, date:'2026-10-30', uploaded:0, complete:true, plannedDate:'2026-11-02'}
];
context.currentGroup = 'A';
assert.deepEqual(Array.from(context.plannerCandidates([]), item => item.id), ['first', 'second']);
assert.deepEqual(Array.from(context.plannerCandidates(['乙']), item => item.id), ['second']);
assert.deepEqual(Array.from(context.plannerCandidates(['甲', '乙']), item => item.id), ['first', 'second']);
const snapshot = context.plannerSnapshot();
context.currentGroup = 'B';
assert.equal(context.plannerSnapshot(), snapshot);
context.items[1].uploaded = 50;
assert.notEqual(context.plannerSnapshot(), snapshot);

context.items = [
  {id:'one', group:'A', name:'one', owner:'甲', total:40, processed:40, date:'2026-10-30', uploaded:0, complete:true},
  {id:'two', group:'A', name:'two', owner:'甲', total:50, processed:50, date:'2026-10-30', uploaded:0, complete:true}
];
elementFor('#plan-target').value = '90';
elementFor('#plan-date').value = '2026-11-02';
const renderAllocation = context.renderAllocation;
context.renderAllocation = () => {};
elementFor('#planner-form').handlers.submit({preventDefault() {}});
assert.equal(elementFor('#planner-results').hidden, false);
assert.equal(vm.runInNewContext('allocationDate', context), '2026-11-02');
elementFor('#plan-date').value = '2026-11-03';
elementFor('#planner-form').handlers.input();
assert.equal(elementFor('#planner-results').hidden, true);
assert.equal(vm.runInNewContext('allocationPlans.length', context), 0);
elementFor('#plan-date').value = '';
elementFor('#planner-form').handlers.submit({preventDefault() {}});
assert.equal(elementFor('#planner-message').textContent, '请选择计划上传日期。');
context.renderAllocation = renderAllocation;

async function testAllocationPreview() {
  context.escapeHTML = value => String(value).replace(/[&<>"']/g, character => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[character]));
  const preview = elementFor('#plan-hover-preview');
  preview.style = {};
  preview.offsetWidth = 200;
  preview.offsetHeight = 100;
  elementFor('#plan-preview-spacer').style = {};
  context.window = {scrollY:300, scrollX:0, innerHeight:720, scrollTo(x, y) { this.scrollY = y; }, scrollBy(x, y) { this.scrollY += y; }};
  elementFor('.planner-result-heading').getBoundingClientRect = () => ({left:10, top:100, width:300});
  context.renderAllocation = () => {};
  context.items.push({owner:'来源甲'});
  vm.runInNewContext(`
    allocationPlans = [
      {totals:[14, 4], assignments:[{person:0, item:{group:'父', name:'一', total:2, owner:'来源甲', date:'2026-01-01'}}, {person:0, item:{group:'父', name:'二', total:2, owner:'甲<一>'}}, {person:0, item:{group:'父', name:'三', total:2, owner:'甲<一>'}}, {person:0, item:{group:'父', name:'四', total:2}}, {person:0, item:{group:'父', name:'五', total:2}}, {person:0, item:{group:'父', name:'六', total:2}}, {person:0, item:{group:'父', name:'七', total:2}}, {person:1, item:{group:'独立', total:4, parentOnly:true}}]},
      {totals:[4, 6], assignments:[{person:0, item:{group:'父', name:'<测试>', total:4}}, {person:1, item:{group:'独立', total:6, parentOnly:true}}]}
    ];
    allocationPeople = ['甲<一>', '乙']; activeAllocation = 0;
  `, context);
  const first = {dataset:{plan:'0'}, textContent:'方案 1', closest() { return this; }, getBoundingClientRect() { return {left:20, top:140, height:35}; }};
  const second = {dataset:{plan:'1'}, textContent:'方案 2', closest() { return this; }, getBoundingClientRect() { return {left:300, top:180, height:35}; }};

  context.showAllocationPreview(first);
  assert.equal(preview.hidden, false);
  assert.ok(preview.innerHTML.includes('甲&lt;一&gt;'));
  assert.ok(preview.innerHTML.includes('14 个文件 · 7 个文件夹'));
  assert.ok(preview.innerHTML.includes('来源：'));
  assert.ok(preview.innerHTML.includes('来源甲'));
  assert.ok(preview.innerHTML.includes('开始：2026-01-01'));
  const markers = context.allocationPersonMarkers();
  const samePerson = context.allocationPersonMarker('甲<一>', markers);
  assert.equal(preview.innerHTML.split(samePerson).length - 1, 3);
  assert.notEqual(markers.get('甲<一>').number, markers.get('来源甲').number);
  assert.notEqual(markers.get('甲<一>').hue, markers.get('来源甲').hue);
  assert.ok(preview.innerHTML.includes('plan-person-unassigned'));
  context.items.push(...['甲', '丙', '丁', '戊', '己', '庚', '辛'].map(owner => ({owner})));
  const manyMarkers = context.allocationPersonMarkers();
  assert.equal(new Set([...manyMarkers.values()].map(marker => marker.number)).size, manyMarkers.size);
  context.items.reverse();
  const reordered = context.allocationPersonMarkers();
  for (const [name, marker] of manyMarkers) assert.equal(reordered.get(name).number, marker.number);
  context.items.reverse();
  vm.runInNewContext('allocationPeople.reverse()', context);
  const reversedPeople = context.allocationPersonMarkers();
  for (const [name, marker] of manyMarkers) {
    assert.equal(reversedPeople.get(name).number, marker.number);
    assert.equal(reversedPeople.get(name).hue, marker.hue);
  }
  vm.runInNewContext('allocationPeople.reverse()', context);
  assert.ok(preview.innerHTML.includes('>一</strong>'));
  assert.ok(preview.innerHTML.includes('>三</strong>'));
  assert.ok(preview.innerHTML.includes('>七</strong>'));
  assert.ok(preview.innerHTML.includes('独立'));
  assert.ok(preview.innerHTML.includes('8 个文件夹 · 3 个来源负责人 · 2 个父文件夹'));
  assert.equal(preview.style.top, '40px');
  assert.equal(preview.style.left, '10px');
  preview.offsetHeight = 1000;
  context.showAllocationPreview(first);
  assert.ok(preview.innerHTML.includes('>七</strong>'));
  assert.equal((preview.innerHTML.match(/class="plan-hover-folder"/g) || []).length, 8);
  assert.ok(!preview.innerHTML.includes('另有'));
  assert.equal(elementFor('#plan-preview-spacer').style.height, '568px');
  context.hideAllocationPreview();
  assert.equal(preview.hidden, true);
  assert.equal(elementFor('#plan-preview-spacer').style.height, '0px');
  assert.equal(context.window.scrollY, 300);
  preview.offsetHeight = 100;
  context.showAllocationPreview(second);
  assert.equal(elementFor('#plan-preview-spacer').style.height, '0px');
  assert.equal(vm.runInNewContext('activeAllocation', context), 0);
  assert.ok(preview.innerHTML.includes('方案 2'));
  assert.ok(preview.innerHTML.includes('&lt;测试&gt;'));
  assert.ok(!preview.innerHTML.includes('<img'));
  assert.equal(preview.style.top, '80px');
  assert.equal(preview.style.left, '100px');

  elementFor('#plan-options').handlers.click({target:second});
  assert.equal(vm.runInNewContext('activeAllocation', context), 1);
  context.resetAllocationPreviews();
  assert.equal(preview.hidden, true);
}

async function testAllocationJson() {
  let copied;
  let clipboardFails = false;
  let confirmationMessage;
  let confirmedPayload;
  context.navigator = {clipboard:{writeText: async text => { if (clipboardFails) throw new Error('denied'); copied = text; }}};
  context.askConfirm = async message => { confirmationMessage = message; return true; };
  context.request = async (url, options) => { confirmedPayload = JSON.parse(options.body); return {items:[]}; };
  context.document.body = {append() {}};
  context.document.createElement = () => ({setAttribute() {}, style:{}, select() {}, remove() {}});
  context.document.getElementById = () => null;
  context.document.execCommand = () => false;
  context.clearPlanPreview = () => {};
  context.render = () => {};
  context.notify = () => {};
  context.items = [
    {id:'json', group:'父', name:'子', total:2, uploaded:0, date:'2026-01-01', owner:'来源甲', complete:true},
    {id:'solo', group:'独立', name:'独立', total:3, uploaded:0, date:'', owner:'', parentOnly:true, complete:true}
  ];
  elementFor('#plan-date').value = '2026-11-02';
  const selectPlan = () => vm.runInNewContext(`
    allocationPlans = [{totals:[2, 3], assignments:[{item:items[0], person:0}, {item:items[1], person:1}]}];
    allocationPeople = ['甲', '乙'];
    activeAllocation = 0;
    allocationSnapshot = plannerSnapshot();
    allocationDate = '2026-11-02';
  `, context);
  selectPlan();

  await elementFor('#plan-create').handlers.click();
  assert.ok(confirmationMessage.includes('2026-11-02'));
  assert.ok(confirmationMessage.includes('不会标记为已上传'));
  assert.ok(confirmationMessage.includes('剪贴板'));
  assert.equal(confirmedPayload.assignments[0].uploadDate, '2026-11-02');
  assert.deepEqual(confirmedPayload.assignments.map(entry => entry.uploader), ['甲', '乙']);
  assert.equal(confirmedPayload.archive, undefined);
  const expected = {version:2, label:'方案 1 · 5 个', uploadDate:'2026-11-02', people:[
    {name:'甲', folders:[{source:['父', '子'], destination:['甲', '父', '子'], owner:'来源甲', startDate:'2026-01-01', files:2, fileList:[]}]},
    {name:'乙', folders:[{source:['独立'], destination:['乙', '独立'], owner:'', startDate:'', files:3, fileList:[]}]}
  ]};
  assert.deepEqual(confirmedPayload.manifest, expected);
  assert.deepEqual(JSON.parse(copied), expected);

  const copyHandler = context.document.handlers.click.at(-1);
  const dateCopyHandler = context.document.handlers.click.at(-2);
  context.planPreview = null;
  context.icons = () => {};
  context.updateDateExport('2026-10-29', uploadedRecords);
  assert.equal(elementFor('#date-history-bar').hidden, false);
  assert.ok(elementFor('#date-history-bar').innerHTML.includes('data-copy-date-json'));
  copied = undefined;
  await dateCopyHandler({target:{closest: () => ({})}});
  assert.deepEqual(JSON.parse(copied), dateJson);
  context.updateDateExport('', uploadedRecords);
  assert.equal(elementFor('#date-history-bar').hidden, true);
  context.updateDateExport('2026-10-29', uploadedRecords);
  context.planPreview = {};
  context.updateDateExport('2026-10-29', uploadedRecords);
  assert.equal(elementFor('#date-history-bar').hidden, true);
  context.planPreview = null;
  context.rememberHistory([{id:'history-1', manifest:{people:[], note:'历史'}}, {id:'legacy', manifest:null}]);
  copied = undefined;
  await copyHandler({target:{closest: () => ({dataset:{copyHistory:'history-1'}})}});
  assert.deepEqual(JSON.parse(copied), {people:[], note:'历史'});
  copied = undefined;
  await copyHandler({target:{closest: () => ({dataset:{copyHistory:'legacy'}})}});
  assert.equal(copied, undefined);

  let historyRefreshes = 0;
  const loadHistory = context.loadAllocationHistory;
  context.loadAllocationHistory = () => { historyRefreshes++; };
  elementFor('#allocation-history').open = true;
  clipboardFails = true;
  selectPlan();
  await elementFor('#plan-create').handlers.click();
  assert.equal(historyRefreshes, 1, '复制失败时已保存的历史也应刷新');
  assert.ok(elementFor('#planner-message').textContent.includes('已保存上传计划，但复制 JSON 失败'));
  elementFor('#allocation-history').open = false;
  context.loadAllocationHistory = loadHistory;
  clipboardFails = false;

  copied = undefined;
  context.request = async () => { throw new Error('conflict'); };
  selectPlan();
  await elementFor('#plan-create').handlers.click();
  assert.equal(copied, undefined, '确认失败时不能复制方案');
  assert.ok(elementFor('#planner-message').textContent.includes('未保存上传计划'));
  const manifest = [{source:['父', '子'], destination:['甲', '父', '子']}];
  const firstScript = fs.readFileSync(path.join(__dirname, '../static/upload_manager/step1_copy_files.py'), 'utf8');
  const secondScript = fs.readFileSync(path.join(__dirname, '../static/upload_manager/step2_copy_onedrive.py'), 'utf8');

  const os = require('node:os');
  const {spawnSync} = require('node:child_process');
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'upload-desk-'));
  try {
    const source = path.join(temp, '资源库 [10% &!]');
    const extracted = path.join(temp, '解压');
    const cloud = path.join(temp, '同步 [20% &!]');
    const documents = path.join(extracted, '分配资料');
    const sourceFolder = path.join(source, '父', '子');
    fs.mkdirSync(path.join(sourceFolder, '[目录]'), {recursive:true});
    fs.mkdirSync(path.join(sourceFolder, '空目录'));
    fs.mkdirSync(documents, {recursive:true});
    fs.mkdirSync(cloud);
    fs.writeFileSync(path.join(sourceFolder, '[目录]', '文件%&!.txt'), 'original');
    fs.writeFileSync(path.join(documents, 'allocation.json'), JSON.stringify(manifest));
    const firstPath = path.join(documents, 'step1_copy_files.py');
    const secondPath = path.join(documents, 'step2_copy_onedrive.py');
    fs.writeFileSync(firstPath, firstScript.replace('r"EDIT_SOURCE_ROOT"', JSON.stringify(source)));
    fs.writeFileSync(secondPath, secondScript.replace('r"EDIT_ONEDRIVE_SYNC_FOLDER"', JSON.stringify(cloud)));
    const runScript = script => spawnSync('python', [script], {encoding:'utf8', timeout:30000, env:{...process.env, PYTHONIOENCODING:'utf-8'}});
    let result = runScript(firstPath);
    assert.equal(result.status, 0, `第一步复制失败：${result.stdout} ${result.stderr}`);
    const relative = path.join('甲', '父', '子', '[目录]', '文件%&!.txt');
    assert.equal(fs.readFileSync(path.join(extracted, relative), 'utf8'), 'original');
    assert.equal(fs.existsSync(path.join(cloud, relative)), false);
    result = runScript(secondPath);
    assert.equal(result.status, 0, `第二步复制失败：${result.stdout} ${result.stderr}`);
    assert.equal(fs.readFileSync(path.join(cloud, relative), 'utf8'), 'original');
    assert.ok(fs.statSync(path.join(cloud, '甲', '父', '子', '空目录')).isDirectory());
    assert.equal(runScript(firstPath).status, 0, '第一步重复执行应成功');
    fs.writeFileSync(path.join(cloud, relative), 'different');
    result = runScript(secondPath);
    assert.equal(result.status, 0, `覆盖失败：${result.stdout} ${result.stderr}`);
    assert.equal(fs.readFileSync(path.join(cloud, relative), 'utf8'), 'original');
    const log = fs.readFileSync(path.join(documents, 'onedrive-overwrite-log.jsonl'), 'utf8').trim().split('\n').map(line => JSON.parse(line));
    assert.deepEqual(log.map(entry => entry.status), ['pending', 'completed']);
    fs.writeFileSync(path.join(sourceFolder, '[目录]', '文件%&!.txt'), 'different');
    result = runScript(firstPath);
    assert.notEqual(result.status, 0, '第一步不同内容的同名文件必须报错');
    assert.equal(fs.readFileSync(path.join(extracted, relative), 'utf8'), 'original');
    fs.writeFileSync(path.join(sourceFolder, '[目录]', '文件%&!.txt'), 'original');
    fs.mkdirSync(path.join(source, '父', '其他'));
    fs.writeFileSync(path.join(source, '父', '其他', '第二人.txt'), 'second');
    fs.writeFileSync(path.join(documents, 'allocation.json'), JSON.stringify([
      ...manifest, {source:['父', '其他'], destination:['乙', '父', '其他']}
    ]));
    assert.equal(runScript(firstPath).status, 0, '多人员第一步复制应成功');
    assert.equal(runScript(secondPath).status, 0, '多人员第二步复制应成功');
    assert.equal(fs.readFileSync(path.join(cloud, '乙', '父', '其他', '第二人.txt'), 'utf8'), 'second');
    fs.writeFileSync(path.join(documents, 'allocation.json'), JSON.stringify(manifest));
    fs.renameSync(sourceFolder, path.join(source, '父', 'moved'));
    assert.notEqual(runScript(firstPath).status, 0, '缺少源目录必须报错');
    fs.renameSync(path.join(source, '父', 'moved'), sourceFolder);
    fs.writeFileSync(firstPath, firstScript.replace('r"EDIT_SOURCE_ROOT"', JSON.stringify(extracted)));
    assert.notEqual(runScript(firstPath).status, 0, '源目录和目标目录重叠必须报错');
    fs.writeFileSync(firstPath, firstScript.replace('r"EDIT_SOURCE_ROOT"', JSON.stringify(source)));
    fs.writeFileSync(path.join(documents, 'allocation.json'), JSON.stringify([
      {source:['父'], destination:['甲', '父']}, {source:['父', '子'], destination:['乙', '父', '子']}
    ]));
    assert.notEqual(runScript(firstPath).status, 0, '分配的源目录互相包含时必须报错');
  } finally {
    fs.rmSync(temp, {recursive:true, force:true});
  }
}
testAllocationPreview().then(testAllocationJson).then(() => console.log('方案分配、悬停预览和方案 JSON 复制测试通过')).catch(error => {
  console.error(error);
  process.exitCode = 1;
});
