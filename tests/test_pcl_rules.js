const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

function rules() {
  const context = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../static/pcl_rules.js'), 'utf8'), context);
  return context.PclRules;
}
const part = (covered, total) => ({covered, total});
function file(short, lines, extra = {}) {
  return {short, name:'/app/' + short, lines, methods:part(0, 0), conditionals:part(0, 0), uncovered:[], partial:[],
    uncovered_methods:[], ...extra};
}

test('source rules still run after moving out of the page', () => {
  const {checkSources} = rules();
  const findings = checkSources([
    {name:'list.twig', type:'twig', code:'webix.ui({view:"datatable", drag:true, columns:[{id:"a", sort:"int"}]});'},
    {name:'repo.php', type:'php', code:'<?php $rows = $db->query(\'SELECT * FROM users\');'}
  ]);
  const warned = Array.from(findings.filter(f => f.status === 'warn'), f => f.id).sort();
  assert.deepEqual(warned, ['T04', 'T11', 'T22']);
  assert.deepEqual(Array.from(findings.find(f => f.id === 'T22').files[0].lines.keys()), [1]);
});

test('T21 is confirmed only when every line and branch ran', () => {
  const {coverageFinding} = rules();
  assert.equal(coverageFinding(null), null);
  const full = coverageFinding({totals:{lines:part(10, 10), conditionals:part(0, 0)}, files:[file('A.php', part(10, 10))]});
  assert.equal(full.status, 'confirmed');
  assert.match(full.message, /100%.*分支遗漏仍需走查/);

  const missed = coverageFinding({totals:{lines:part(1999, 2000), conditionals:part(1, 2)}, files:[
    file('A.php', part(5, 5)),
    file('B.php', part(1994, 1995), {uncovered:[[8, 9]], partial:[[7, 7]]})
  ]});
  assert.equal(missed.status, 'warn');
  assert.match(missed.message, /99\.9%/);  // 不能四舍五入成 100%
  assert.deepEqual(Array.from(missed.extra), ['B.php · 未执行第 8-9 行 · 分支没走全：第 7 行']);
});
