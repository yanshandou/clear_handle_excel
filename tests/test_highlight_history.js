const assert = require('node:assert/strict');
const test = require('node:test');
const { page } = require('./workbench_page');

const CURRENT = 'clear_handle_excel.highlightColors';
const HISTORY = 'clear_handle_excel.highlightColorHistory';
const TYPES = 'clear_handle_excel.typeHighlightColors';
const custom = { fill: '#123456', font: '#FEDCBA' };
const readHistory = p => JSON.parse(p.storage.get(HISTORY) || '[]');
const pick = (p, name, type, value) => p.elements.get('typeColors').listeners[name]({ target: { value, dataset: { typeColor: type } } });

test('each problem type has its own default color chip, and the scan sends them all', async () => {
  const p = page();
  const html = p.elements.get('typeColors').innerHTML;
  for (const label of ['边框压线', '文字截断', '超出裁剪框', '分页切开', '### 列宽不足', 'PDF 未匹配', '多种问题']) assert.match(html, new RegExp(label));
  assert.match(html, /background:#C0007A/);
  assert.match(html, /background:#000000;color:#FFFF00/);
  assert.equal(p.elements.get('typeColorsReset').classList.contains('hidden'), true);
  p.elements.get('folderInput').value = 'C:/Excel';
  p.run(`appConfirm = async () => true`);
  await p.elements.get('startBtn').listeners.click();
  const sent = p.requests.at(-1).payload.highlight.types;
  assert.equal(Object.keys(sent).length, 7);
  assert.equal(JSON.stringify(sent.text_truncated), JSON.stringify({ fill: '#1649CC', font: '#FFFFFF' }));
});

test('changing one type color previews without flooding history, picks a readable font and survives reloading', () => {
  const storage = new Map();
  const p = page({ storage });
  for (const value of ['#123450', '#123451', '#123456']) pick(p, 'input', 'border_clip', value);
  assert.deepEqual(readHistory(p), []);
  pick(p, 'change', 'border_clip', '#123456');
  assert.deepEqual(readHistory(p), [{ fill: '#123456', font: '#FFFFFF' }]);
  pick(p, 'change', 'hash_overflow', '#FFEE00');  // 浅色底配黑字
  assert.deepEqual(JSON.parse(storage.get(TYPES)).hash_overflow, { fill: '#FFEE00', font: '#000000' });
  const reloaded = page({ storage });
  assert.match(reloaded.elements.get('typeColors').innerHTML, /background:#123456;color:#FFFFFF/);
  assert.equal(reloaded.elements.get('typeColorsReset').classList.contains('hidden'), false);
  reloaded.elements.get('typeColorsReset').listeners.click();
  assert.match(reloaded.elements.get('typeColors').innerHTML, /background:#C0007A/);
  assert.equal(JSON.parse(storage.get(TYPES)).border_clip.fill, '#C0007A');
});

test('an old single custom color is kept in history so the repair page still finds it', () => {
  const p = page({ storage: new Map([[CURRENT, JSON.stringify(custom)]]) });
  assert.deepEqual(readHistory(p), [custom]);
});

test('repair palette includes presets, history and type colors without duplicates', () => {
  const history = Array.from({ length: 35 }, (_, i) => ({
    fill: '#' + (0x123400 + i).toString(16).toUpperCase(), font: '#FFFFFF',
  }));
  const p = page({ storage: new Map([[HISTORY, JSON.stringify(history)]]) });
  pick(p, 'change', 'box_overflow', '#654321');
  const pairs = Array.from(p.run('savedHighlightColors()'), color => `${color.fill}/${color.font}`);
  for (const color of history) assert.ok(pairs.includes(`${color.fill}/${color.font}`));
  assert.ok(pairs.includes('#654321/#FFFFFF'));
  assert.ok(pairs.includes('#1649CC/#FFFFFF'));
  assert.equal(new Set(pairs).size, pairs.length);
});

test('malformed or unavailable browser storage falls back to the default type colors', () => {
  for (const value of ['{bad json', '{}', 'null', '{"border_clip":{"fill":"red"}}']) {
    const p = page({ storage: new Map([[HISTORY, value], [CURRENT, value], [TYPES, value]]) });
    assert.match(p.elements.get('typeColors').innerHTML, /background:#C0007A/);
  }
  const p = page({ storage: { get() { throw Error('blocked'); }, set() { throw Error('blocked'); } } });
  pick(p, 'change', 'border_clip', '#123456');
  assert.match(p.elements.get('typeColors').innerHTML, /background:#123456/);
  assert.equal(p.run('savedHighlightColors().length'), 8);
});
