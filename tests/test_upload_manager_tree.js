const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../static/upload_manager/app.js'), 'utf8');
const context = {
  collapsedFolders: new Set(),
  currentOwners: new Set(),
  escapeHTML: value => String(value).replace(/[&<>"']/g, character => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[character])),
  isComplete: item => item.complete,
  percentOf: () => 0
};
vm.runInNewContext(source.slice(source.indexOf('const folderStatuses'), source.indexOf('function percentOf(')), context);
vm.runInNewContext(source.slice(source.indexOf('function folderTree('), source.indexOf('function render()')), context);

function item(id, group, name, parentOnly = false) {
  return {id, group, name, parentOnly, total: 1, uploaded: 0, complete: true, date: '', owner: ''};
}
const records = [
  item('parent', '根', '父'),
  item('child', '根/父', '子'),
  item('parent-only', '独立\\目录', '目录', true),
  item('separator', '///', '特殊', true),
  item('escaped', '根', '<测试>')
];
const base = {total:10, processed:0, date:'', uploaded:0, complete:false};
assert.equal(context.statusOf(base), 'unstarted');
assert.equal(context.statusOf({...base, date:'2026-10-30'}), 'processing');
assert.equal(context.statusOf({...base, date:'2026-10-30', processed:10}), 'ready');
assert.equal(context.statusOf({...base, plannedDate:'2026-11-02'}), 'uploading');
assert.equal(context.statusOf({...base, uploaded:10}), 'uploaded');
const manuallyReady = context.withFolderStatus(base, 'ready');
assert.equal(context.statusOf(manuallyReady), 'ready');
assert.equal(manuallyReady.processed, 0);
assert.equal(manuallyReady.date, '');
assert.equal(context.withFolderStatus({...base, plannedDate:'2026-11-02', uploader:'甲'}, 'uploaded').uploaded, 10);
assert.equal(context.withFolderStatus({...base, plannedDate:'2026-11-02', uploader:'甲'}, 'unstarted').plannedDate, '');
const cycle = [];
for (let item = {...base, total:10}, index = 0; index < 6; index++) {
  cycle.push(context.statusOf(item));
  item = context.withFolderStatus(item, context.nextStatus(item));
}
assert.deepEqual(cycle, ['unstarted', 'processing', 'ready', 'uploading', 'uploaded', 'unstarted']);
assert.equal(context.nextStatus({...base, statusOverride:'uploading', total:0, totalMissing:true}), 'unstarted');
const tree = context.folderTree(records);
const parent = tree.children.get('根').children.get('父');
assert.equal(parent.item.id, 'parent');
assert.equal(parent.children.get('子').item.id, 'child');
assert.equal(tree.children.get('独立').children.get('目录').item.id, 'parent-only');
assert.equal(tree.children.get('///').item.id, 'separator');

const expanded = context.renderFolderTree(tree);
assert.equal((expanded.match(/data-action="edit"/g) || []).length, records.length);
assert.ok(expanded.includes('&lt;测试&gt;'));
assert.ok(expanded.includes('data-status-id="parent"'));
assert.ok(expanded.includes('class="folder-status unstarted"'));
assert.ok(expanded.includes('点击切换为完成中...'));
assert.ok(expanded.includes('class="processed-count">—</span>'));
records[1].processed = 0;
records[1].uploader = '名义<人员>';
records[1].uploadDate = '2026-10-29';
assert.ok(context.renderFolderTree(tree).includes('class="processed-count">0</span>'));
assert.ok(context.renderFolderTree(tree).includes('<td>名义&lt;人员&gt;</td>'));
assert.ok(context.renderFolderTree(tree).includes('<td>2026 / 10 / 29</td>'));
assert.ok(context.renderFolderTree(tree).includes('colspan="10"'));
assert.equal(context.matchesUploadDate(records[1], '2026-10-29'), true);
assert.equal(context.matchesUploadDate(records[1], '2026-10-30'), false);
assert.equal(context.matchesUploadDate(records[0], '2026-10-29'), false);
assert.equal(context.matchesUploadDate(records[0], ''), true);
const preview = context.renderFolderTree(tree, new Map([['child', '甲<乙>']]));
assert.ok(preview.includes('分配给 甲&lt;乙&gt;'));
assert.equal((preview.match(/class="tree-assignment"/g) || []).length, 1);
assert.ok(expanded.includes('class="owner-tag owner-cell-tag" data-owner="" aria-pressed="false"'));
context.currentOwners.add('');
assert.ok(context.renderFolderTree(tree).includes('class="owner-tag owner-cell-tag" data-owner="" aria-pressed="true"'));
context.currentOwners.clear();
context.collapsedFolders.add(parent.path);
const collapsed = context.renderFolderTree(tree);
assert.ok(collapsed.includes('data-id="parent"'));
assert.ok(!collapsed.includes('data-id="child"'));
assert.ok(collapsed.includes('aria-expanded="false"'));
context.collapsedFolders.clear();
assert.ok(context.renderFolderTree(tree).includes('data-id="child"'));
context.setFoldersExpanded(tree, false);
assert.ok(context.collapsedFolders.has(tree.children.get('根').path));
assert.ok(context.collapsedFolders.has(parent.path));
assert.ok(!context.renderFolderTree(tree).includes('data-id="child"'));
context.collapsedFolders.delete(tree.children.get('根').path);
assert.ok(!context.renderFolderTree(tree).includes('data-id="child"'));
context.setFoldersExpanded(tree, true);
assert.ok(context.renderFolderTree(tree).includes('data-id="child"'));
const choices = context.folderChoices(records, '根/父', 'child');
assert.equal(choices.find(choice => choice.name === '子').unavailable, false);
assert.equal(context.folderChoices(records, '根/父', null).find(choice => choice.name === '子').unavailable, true);
assert.equal(context.folderChoices(records, '根', 'parent').find(choice => choice.name === '父').unavailable, false);
assert.equal(context.folderChoices(records, '独立\\目录', null)[0].unavailable, true);
assert.equal(context.folderChoices(records, '独立\\目录', 'parent-only')[0].unavailable, false);
assert.equal(context.folderChoices(records, '根/父', 'child').some(choice => choice.name === '父'), false);
assert.equal(context.folderChoices(records, '根/父', 'child').some(choice => choice.name === ''), false);
assert.equal(context.folderChoices(records, '不存在', 'child').length, 0);

console.log('目录树和展开折叠测试通过');
