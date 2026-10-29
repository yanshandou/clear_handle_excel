const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function cleanPage() {
  const source = fs.readFileSync(path.join(__dirname, '..', 'static', 'clean.js'), 'utf8');
  // Load page functions without starting DOM initialization or polling.
  const initialization = source.indexOf('const cleanFolderInput =');
  assert.ok(initialization > 0);
  const context = vm.createContext({window: {}});
  vm.runInContext(source.slice(0, initialization), context);
  return expression => vm.runInContext(expression, context);
}

test('sheet ID digits survive backup matching and highlighting', () => {
  const run = cleanPage();
  assert.equal(run("hasBackupMarker('雛型SSSEE0114')"), false);
  assert.equal(run("hasBackupMarker('雛型ＳＳＳＥＥ０１１４')"), false);
  assert.equal(run("hasBackupMarker('雛型SSSEE0114_20231023')"), true);
  assert.equal(run("analyzeSheetName('雛型SSSEE0114_20231023', {stripBackup:true}).key"), '雛型sssee0114');

  const highlights = JSON.parse(run(`JSON.stringify((() => {
    const file = {name:'templates.xlsx', sheets:[
      '雛型SSSEE0114', '雛型SSSEE0114_20231023', '雛型SSSEE0115', '雛型SSSEE0116',
    ], backups:['雛型SSSEE0114_20231023']};
    previewData = {files:[file]};
    initSheetDecisions();
    return computeFileSheetHighlights(file);
  })())`));
  const original = '雛型SSSEE0114';
  assert.deepEqual(highlights[0].mask, Array(original.length).fill(true));
  assert.equal(highlights[1].colorIdx, highlights[0].colorIdx);
  assert.deepEqual(highlights[1].mask, [
    ...Array(original.length).fill(true), ...Array('_20231023'.length).fill(false),
  ]);
  assert.equal(highlights[2], null);
  assert.equal(highlights[3], null);
});

test('MMDD still works after Japanese names or separators', () => {
  const run = cleanPage();
  for (const name of ['一覧0513', 'Report_0513', 'レイアウト(請求書1)0513']) {
    assert.equal(run(`hasBackupMarker(${JSON.stringify(name)})`), true, name);
  }
  assert.equal(run("analyzeSheetName('Report_0513', {stripBackup:true}).key"), 'report');
  assert.equal(run("hasBackupMarker('Report0513')"), false);
  assert.equal(run("hasBackupMarker('Report_1340')"), false);
});

test('rescan retains manual decisions only for an unchanged file', async () => {
  const run = cleanPage();
  run(`previewData = {folder:'folder', files:[{name:'a', sheets:['A','AB'], backups:[], uncertain:['AB'], version:{size:1}}]};
    initSheetDecisions(); sheetDecisions.a.AB = 'keep';
    readFolderInput = () => 'folder'; setPending = () => {}; saveLastFolder = () => {};
    renderPreview = () => {}; renderPage = () => {}; appAlert = e => { throw Error(e); };
    postAction = async () => ({folder:'folder', files:[{name:'a', sheets:['A','AB'], backups:[], uncertain:['AB'], version:{size:1}}]});`);
  await run('runPreview()');
  assert.equal(run('sheetDecisions.a.AB'), 'keep');
  run(`postAction = async () => ({folder:'folder', files:[{name:'a', sheets:['A','AB'], backups:[], uncertain:['AB'], version:{size:2}}]});`);
  await run('runPreview()');
  assert.equal(run('sheetDecisions.a.AB'), 'uncertain');
});
