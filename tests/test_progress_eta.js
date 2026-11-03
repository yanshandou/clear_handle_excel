// 进度条预估时间（static/eta.js）：按已用时间和进度推算剩余时间、完成时刻
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function load() {
  let now = new Date(2026, 10, 3, 14, 0, 0).getTime();
  const storage = new Map();
  const RealDate = Date;
  class FakeDate extends RealDate {
    constructor(...args) { super(...(args.length ? args : [now])); }
    static now() { return now; }
  }
  const context = vm.createContext({
    Date: FakeDate, JSON, Math, Number, String,
    window: {}, document: { getElementById: () => null },
    localStorage: { getItem: k => storage.get(k) ?? null, setItem: (k, v) => storage.set(k, v), removeItem: k => storage.delete(k) },
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'static', 'eta.js'), 'utf8'), context);
  return { eta: context.window.ProgressEta, tick: ms => { now += ms; }, storage };
}

test('estimates finish time from elapsed time and progress', () => {
  const { eta, tick } = load();
  assert.equal(eta.text('job', 0), '⏱ 14:00 开始 → 正在估算完成时间（走完第一份后算出）…');
  tick(3 * 60000);
  // 3 分钟走了 25%：还要 9 分钟，14:12 完成
  assert.equal(eta.text('job', 0.25), '⏱ 14:00 开始 → 预计 14:12 完成 · 还要约 9 分钟');
  // 进度没变时完成时刻不变，剩余时间倒数
  tick(60000);
  assert.equal(eta.text('job', 0.25), '⏱ 14:00 开始 → 预计 14:12 完成 · 还要约 8 分钟');
});

test('pushes the finish time back when stuck past the estimate', () => {
  const { eta, tick } = load();
  eta.text('job', 0);
  tick(60000);
  eta.text('job', 0.5); // 预计 14:02 完成
  tick(5 * 60000);       // 到 14:06 还是 50%
  assert.equal(eta.text('job', 0.5), '⏱ 14:00 开始 → 预计 14:12 完成 · 还要约 6 分钟');
});

test('unknown start when first seen mid-way, restarts when progress goes back, and stop clears it', () => {
  const { eta, tick, storage } = load();
  assert.match(eta.text('job', 0.4), /^⏱ 正在估算/);
  tick(2 * 60000);
  assert.equal(eta.text('job', 0.6), '⏱ 约 13:56 开始 → 预计 14:06 完成 · 还要约 4 分钟');
  assert.match(eta.text('job', 0.01), /^⏱ 14:02 开始 → 正在估算/);
  eta.stop('job');
  assert.equal(storage.size, 0);
});

test('long jobs show hours and next-day finish', () => {
  const { eta, tick } = load();
  eta.text('job', 0);
  for (let i = 1; i <= 120; i++) { tick(60000); eta.text('job', i / 600); } // 页面一直在看：2 小时 20%
  assert.equal(eta.text('job', 0.2), '⏱ 14:00 开始 → 预计 明天 00:00 完成 · 还要约 8 小时');
  assert.equal(eta.left(4500), '约 1 小时 15 分钟');
});
