const assert = require('node:assert/strict');
const test = require('node:test');
const { page } = require('./workbench_page');

const CURRENT = 'clear_handle_excel.highlightColors';
const HISTORY = 'clear_handle_excel.highlightColorHistory';
const custom = { fill: '#123456', font: '#FEDCBA' };
const readHistory = p => JSON.parse(p.storage.get(HISTORY) || '[]');

test('existing custom selection migrates to history and survives switching to a preset and reloading', () => {
  const storage = new Map([[CURRENT, JSON.stringify(custom)]]);
  const p = page({ storage });
  assert.deepEqual(readHistory(p), [custom]);
  assert.equal(p.elements.get('highlightHistoryRow').classList.contains('hidden'), false);
  p.run('setHighlight(HIGHLIGHT_PRESETS[4])');
  const reloaded = page({ storage });
  assert.match(reloaded.elements.get('highlightHistory').innerHTML, /#123456/);
  reloaded.elements.get('highlightHistory').listeners.click({ target: {
    closest: () => ({ dataset: { history: '0' } }),
  } });
  assert.equal(reloaded.elements.get('highlightFill').value, '#123456');
  assert.equal(reloaded.elements.get('highlightFont').value, '#fedcba');
  assert.match(reloaded.elements.get('highlightHistory').innerHTML, /aria-pressed="true"/);
  assert.deepEqual(JSON.parse(storage.get(CURRENT)), custom);
});

test('picker previews do not flood history and confirmed selections are saved as fill/font pairs', () => {
  const p = page();
  assert.equal(p.elements.get('highlightHistoryRow').classList.contains('hidden'), true);
  const fill = p.elements.get('highlightFill').listeners;
  for (const value of ['#123450', '#123451', '#123456']) fill.input({ target: { value } });
  assert.deepEqual(readHistory(p), []);
  fill.change({ target: { value: '#123456' } });
  const font = p.elements.get('highlightFont').listeners;
  font.input({ target: { value: '#fedcba' } });
  font.change({ target: { value: '#fedcba' } });
  assert.deepEqual(readHistory(p), [custom, { fill: '#123456', font: '#FFFFFF' }]);
  p.run('setHighlight(HIGHLIGHT_PRESETS[0])');
  assert.equal(readHistory(p).length, 2);
});

test('history validates, normalizes and deduplicates pairs while retaining different fonts for one fill', () => {
  const p = page({ storage: new Map([[HISTORY, JSON.stringify([
    { fill: '#abcdef', font: '#fedcba' }, { fill: '#ABCDEF', font: '#FEDCBA' },
    { fill: '#ABCDEF', font: '#000000' }, { fill: '#C0007A', font: '#FFFFFF' },
    { fill: 'red', font: '#FFFFFF' }, { fill: ['#123456'], font: '#FFFFFF' }, null, {},
  ])]]) });
  assert.equal(p.run('highlightHistory.length'), 2);
  assert.equal(p.run('highlightHistory[0].fill'), '#ABCDEF');
  p.run('setHighlight(highlightHistory[1])');
  assert.deepEqual(readHistory(p), [
    { fill: '#ABCDEF', font: '#000000' }, { fill: '#ABCDEF', font: '#FEDCBA' },
  ]);
});

test('find marks includes every history pair, current preview and presets without duplicates', async () => {
  const history = Array.from({ length: 35 }, (_, i) => ({
    fill: '#' + (0x123400 + i).toString(16).toUpperCase(), font: '#FFFFFF',
  }));
  history.push({ fill: '#C0007A', font: '#000000' });
  const p = page({ storage: new Map([[HISTORY, JSON.stringify(history)]]) });
  p.elements.get('highlightFill').listeners.input({ target: { value: custom.fill } });
  p.run('waitFindMarks = async () => {}');
  await p.run('findMarks()');
  const { url, payload } = p.requests[0];
  assert.equal(url, '/api/batch/find_marks');
  const pairs = Array.from(payload.colors, color => `${color.fill}/${color.font}`);
  for (const color of history) assert.ok(pairs.includes(`${color.fill}/${color.font}`));
  assert.ok(pairs.includes('#123456/#FFFFFF'));
  assert.ok(pairs.includes('#C0007A/#FFFFFF'));
  assert.equal(pairs.length, history.length + 8 + 1);
  assert.equal(new Set(pairs).size, pairs.length);
});

test('search results show the actual palette even when no cells match', () => {
  const p = page();
  p.run(`showFoundMarks({total: 1, folder: 'C:/Excel', files: [{name: 'book.xlsx', hits: []}],
    colors: [{fill: '#123456', font: '#FEDCBA'}]})`);
  const html = p.elements.get('findMarksBody').innerHTML;
  assert.match(html, /本次匹配 1 组配色/);
  assert.match(html, /底 #123456 \/ 字 #FEDCBA/);
});

test('malformed or unavailable browser storage does not prevent choosing and finding colors', async () => {
  for (const value of ['{bad json', '{}', 'null']) {
    const p = page({ storage: new Map([[HISTORY, value], [CURRENT, value]]) });
    p.run(`setHighlight({fill:'#123456',font:'#FEDCBA'})`);
    assert.deepEqual(readHistory(p), [custom]);
  }
  const p = page({ storage: { get() { throw Error('blocked'); }, set() { throw Error('blocked'); } } });
  p.run(`setHighlight({fill:'#123456',font:'#FEDCBA'}); setHighlight(HIGHLIGHT_PRESETS[0]);
    waitFindMarks = async () => {}`);
  await p.run('findMarks()');
  assert.equal(p.requests[0].payload.colors.length, 9);
});
