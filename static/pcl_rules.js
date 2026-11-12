// PCL 检查台的源码静态规则（从 static/UIPCL.html 拆出来；页面 static/pcl_check.js 和 tests/test_pcl_rules.js 共用）
// 只看静态写法，结论只适用于列出的代码位置；T21 的代码覆盖率由 core/pcl_check.py 解析 clover.xml 后在这里转成检查结论。
(function (root) {
'use strict';
const SOURCE_RULES = {
  T04: '拖拽表格的列开启了排序。',
  T07: '多个 Webix 组件声明了相同的静态 ID，请在 Debug 模式确认它们是否同时存在。',
  T08: 'foreach 的循环变量被整行赋值，请确认输出是否应逐字段组装。',
  T09: 'form 中的非 datatable 控件设置了 minWidth 或 maxWidth。',
  T10: '组件配置使用了 autowidth: true。',
  T11: '启用拖拽的 datatable 没有使用 drag: "order"。',
  T14: '发现单引号字符串，请核对 Twig 的引号规范。',
  T15: '发现双引号字符串，修改前请确认插值和转义。',
  T18: 'validate 中未识别到表单校验失败后提示 VL10001 并返回 false 的标准写法。',
  T22: '发现通配符取列，请核对实际需要的字段。'
};
const SOURCE_TYPES = {twig:'Twig / JavaScript', php:'PHP', sql:'SQL'};
const SOURCE_STATUS = {warn:'发现疑点', unknown:'信息不足', confirmed:'静态符合', skipped:'此配置不适用'};
const SOURCE_LIMIT = 200000;
function sourceType(code, filename = '') {
  if (/\.php$/i.test(filename)) return 'php';
  if (/\.sql$/i.test(filename)) return 'sql';
  if (/\.(?:twig|[cm]?js)$/i.test(filename)) return 'twig';
  if (/<\?(?:php\b|=)/i.test(code)) return 'php';
  const start = code.replace(/^(?:\s+|--[^\n]*|\/\*[\s\S]*?\*\/)+/, '');
  if (/^(?:select|with|insert|update|delete)\s/i.test(start)) return 'sql';
  if (/^(?:namespace\s|use\s+[\w\\]+;|\$[a-z_]|(?:public\s+|protected\s+|private\s+)?function\s+\w+\s*\(\s*\$|foreach\s*\()/i.test(start)) return 'php';
  return 'twig';
}

function sourceTokens(code, type, offset = 0, start = 0) {
  const quoted = /"(?:\\[\s\S]|[^"\\])*"|'(?:\\[\s\S]|[^'\\])*'|`(?:\\[\s\S]|[^`\\])*`/.source;
  const comments = {
    twig: String.raw`\/\/[^\n]*|\/\*[\s\S]*?(?:\*\/|(?![\s\S]))|\{#[\s\S]*?(?:#\}|(?![\s\S]))`,
    php: String.raw`\/\/[^\n]*?(?=\?>|$)|\/\*[\s\S]*?(?:\*\/|(?![\s\S]))|\#(?!\[)[^\n]*?(?=\?>|$)`,
    sql: String.raw`--[^\n]*|\/\*[\s\S]*?(?:\*\/|(?![\s\S]))|\#[^\n]*`
  };
  const heredoc = type === 'php'
    ? String.raw`<<<[ \t]*["']?([a-z_]\w*)["']?[^\n]*\n[\s\S]*?^[ \t]*\1\b|`
    : '';
  const pattern = new RegExp(heredoc + comments[type] + '|' + quoted +
    String.raw`|\$*[a-z_][\w$]*|\$+|\d+|=>|->|===|==|!==|!=|\?>|\.{3}|[^\s]`, 'gim');
  pattern.lastIndex = start;
  const regexLiteral = /\/(?:\\[^\n]|\[(?:\\[^\n]|[^\]\\\n])*\]|[^\/\\[\n])+\/[a-z]*/iy;
  const tokens = [];
  let match;
  while ((match = pattern.exec(code))) {
    let value = match[0];
    if (match[1] || /^(?:\/\/|\/\*|\{#|--|#(?!\[))/.test(value)) continue;
    const previous = tokens.at(-1)?.value;
    if (type === 'twig' && value === '/' &&
        (!previous || /^(?:[=(:,;[!{?]|return|=>)$/.test(previous))) {
      regexLiteral.lastIndex = match.index;
      const literal = regexLiteral.exec(code);
      if (literal) {
        pattern.lastIndex = regexLiteral.lastIndex;
        value = literal[0];
      }
    }
    tokens.push({
      value, index: offset + match.index,
      quote: /^["'`]/.test(value) && value.length > 1 ? value[0] : '',
      end: -1
    });
    if (type === 'php' && start > 0 && value === '?>') break;
  }
  const brackets = [];
  for (const [i, token] of tokens.entries()) {
    if (['{', '[', '('].includes(token.value)) brackets.push(i);
    const opener = {'}':'{', ']':'[', ')':'('}[token.value];
    if (opener && tokens[brackets.at(-1)]?.value === opener) tokens[brackets.pop()].end = i;
  }
  return tokens;
}

function sourceSections(code, type) {
  if (type === 'php') {
    const opening = /<\?(?:php\b|=)/gi;
    const sections = [];
    while (opening.exec(code)) {
      const tokens = sourceTokens(code, type, 0, opening.lastIndex);
      const last = tokens.at(-1);
      if (last?.value === '?>') tokens.pop();
      sections.push(tokens);
      if (last?.value !== '?>') break;
      opening.lastIndex = last.index + 2;
    }
    return sections.length ? sections : [sourceTokens(code, type)];
  }
  if (type !== 'twig' || !/^\s*<(?:!doctype|!--|\/?[a-z][\w:-]*(?=[\s/>]))/im.test(code)) {
    return [sourceTokens(code, type)];
  }
  // Scan script contents and Twig expressions. HTML text and tag attributes are exempt.
  const regions = /<!--[^]*?(?:-->|$)|\{#[^]*?(?:#\}|$)|<script\b(?:[^"'<>]|"[^"]*"|'[^']*')*>([^]*?)(?:<\/script\s*>|$)|<\/?[a-z](?:[^"'<>]|"[^"]*"|'[^']*')*>|(\{\{[^]*?\}\}|\{%[^]*?%\})/gi;
  return [...code.matchAll(regions)].flatMap(match => {
    const body = match[1] ?? match[2];
    if (!body) return [];
    const start = match[1] === undefined ? 0 : match[0].match(/^<script\b(?:[^"'<>]|"[^"]*"|'[^']*')*>/i)[0].length;
    return [sourceTokens(body, type, match.index + start)];
  });
}

function* sourceEntries(tokens, start, end) {
  for (let i = start; i < end; i++) {
    yield i;
    if (tokens[i].end > i) i = tokens[i].end;
  }
}

function sourceLiteral(token) {
  if (!token) return '';
  return token.quote ? token.value.slice(1, -1) : token.value;
}

function sourceObjects(tokens) {
  const objects = new Map();
  for (const [i, token] of tokens.entries()) {
    if (token.value !== '{' || token.end < 0) continue;
    const props = new Map();
    for (const j of sourceEntries(tokens, i + 1, token.end)) {
      if ((j === i + 1 || tokens[j - 1].value === ',') && tokens[j + 1]?.value === ':') {
        props.set(sourceLiteral(tokens[j]), tokens[j + 2]);
      }
    }
    if (props.size) objects.set(token, props);
  }
  return objects;
}

function* sourceMembers(tokens, start, end) {
  for (const i of sourceEntries(tokens, start, end)) {
    if (tokens[i].value !== ',') continue;
    if (start < i) yield {start, end:i};
    start = i + 1;
  }
  if (start < end) yield {start, end};
}

const UNKNOWN_PROPERTY = Symbol('unknown property');

function readSourceProperties(tokens, start, end) {
  const props = new Map();
  let hasUnknownKeys = false;
  for (const member of sourceMembers(tokens, start, end)) {
    const key = tokens[member.start];
    const next = tokens[member.start + 1];
    if (/^[a-z_$][\w$]*$/i.test(key.value) || /^["'][^\\{}]*["']$/.test(key.value)) {
      const hasValue = member.start + 2 < member.end && next?.value === ':';
      const body = next?.value === '(' && next.end >= 0 ? tokens[next.end + 1] : null;
      const method = body?.value === '{' && body.end === member.end - 1;
      if (hasValue || member.start + 1 === member.end || method) {
        let value = UNKNOWN_PROPERTY;
        if (hasValue) value = {start:member.start + 2, end:member.end};
        if (method) value = {...member, method:true};
        props.set(sourceLiteral(key), value);
        continue;
      }
    }
    // A spread or computed key may replace any preceding property.
    props.clear();
    hasUnknownKeys = true;
  }
  return key => props.get(key) ?? (hasUnknownKeys ? UNKNOWN_PROPERTY : null);
}

function sourceStaticValue(tokens, range) {
  if (!range || range === UNKNOWN_PROPERTY || range.end !== range.start + 1) return undefined;
  const {value} = tokens[range.start];
  if (value === 'true') return true;
  if (value === 'false') return false;
  if (value === 'null') return null;
  if (/^\d+$/.test(value)) return Number(value);
  if (/^(?:"[^"\\{}]*"|'[^'\\{}]*')$/.test(value)) return sourceLiteral(tokens[range.start]);
  return undefined;
}

function sourceSortEnabled(tokens, range) {
  if (range.method) return true;
  const value = sourceStaticValue(tokens, range);
  if (value !== undefined) return Boolean(value);
  if (range === UNKNOWN_PROPERTY) return undefined;
  let start = range.start;
  if (tokens[start].value === 'async' && start + 1 < range.end && tokens[start + 1].value !== '=>') start++;
  const first = tokens[start];
  if (first.value === 'function') {
    const params = [...sourceEntries(tokens, start + 1, range.end)].find(i => tokens[i].value === '(');
    const body = params === undefined ? null : tokens[tokens[params].end + 1];
    return body?.value === '{' && body.end === range.end - 1 ? true : undefined;
  }
  const parameterEnd = first.value === '(' ? first.end : /^[a-z_$][\w$]*$/i.test(first.value) ? start : -1;
  if (parameterEnd >= start && parameterEnd + 1 < range.end && tokens[parameterEnd + 1].value === '=>') return true;
  return undefined;
}

function checkDragSorting(tokens, object, report) {
  const props = readSourceProperties(tokens, object + 1, tokens[object].end);
  const view = sourceStaticValue(tokens, props('view'));
  const drag = props('drag');
  const dragValue = drag === null ? false : sourceStaticValue(tokens, drag);
  const index = tokens[object].index;
  if (view !== 'datatable' || dragValue === undefined) {
    const message = '表格类型或 drag 来自动态配置，需核对最终生效的配置。';
    report('T04', index, 'unknown', message);
    report('T11', index, 'unknown', message);
    return;
  }
  if (!dragValue) {
    const message = '此 datatable 的静态配置未开启拖拽。';
    report('T04', index, 'skipped', message);
    report('T11', index, 'skipped', message);
    return;
  }
  if (dragValue === 'order') {
    report('T11', tokens[drag.start].index, 'confirmed', '此 datatable 使用了 drag: "order"。');
  } else {
    report('T11', tokens[drag.start].index);
  }

  const columns = props('columns');
  const array = columns && tokens[columns.start];
  const hasInlineColumns = array?.value === '[' && array.end === columns.end - 1;
  const entries = hasInlineColumns ? [...sourceMembers(tokens, columns.start + 1, array.end)] : [];
  let unresolved = entries.length === 0;
  let enabled = false;
  for (const column of entries) {
    const token = tokens[column.start];
    if (token.value !== '{' || token.end !== column.end - 1) {
      unresolved = true;
      continue;
    }
    const sort = readSourceProperties(tokens, column.start + 1, token.end)('sort');
    if (sort === null) continue;
    const sorting = sourceSortEnabled(tokens, sort);
    if (sorting === undefined) {
      unresolved = true;
    } else if (sorting) {
      report('T04', tokens[sort.start].index);
      enabled = true;
    }
  }
  if (enabled) return;
  if (unresolved) {
    report('T04', index, 'unknown', '有列定义或 sort 值无法从此处确认，需核对最终生效的列配置。');
  } else {
    report('T04', index, 'confirmed', '此拖拽表格的内联列配置均未开启列排序。');
  }
}

function checkTwigSource(tokens, report, componentIds) {
  for (const token of tokens) {
    if (token.quote === "'") report('T14', token.index);
  }
  const objects = sourceObjects(tokens);
  const positions = new Map(tokens.map((token, i) => [token, i]));
  let formEnd = -1;
  for (const [object, props] of objects) {
    const view = props.get('view');
    const viewName = view?.quote ? sourceLiteral(view) : '';
    const id = props.get('id');
    if (viewName && id?.quote && id.quote !== '`') {
      const name = sourceLiteral(id);
      if (name && !/\{[{%]/.test(name)) {
        if (componentIds.has(name)) {
          const first = componentIds.get(name);
          first.report('T07', first.index);
          report('T07', id.index);
        } else {
          componentIds.set(name, {report, index:id.index});
        }
      }
    }
    if (viewName === 'form') formEnd = Math.max(formEnd, object.end);
    if (viewName && viewName !== 'form' && viewName !== 'datatable' && positions.get(object) < formEnd) {
      const width = props.get('minWidth') || props.get('maxWidth');
      if (width) report('T09', width.index);
    }
    for (const [name, value] of props) {
      if (name.toLowerCase() === 'autowidth' && value?.value === 'true') report('T10', value.index);
    }
    if (viewName === 'datatable') checkDragSorting(tokens, positions.get(object), report);
  }
  for (const [i, token] of tokens.entries()) {
    if (token.value !== 'validate') continue;
    let paramsIndex = i + 1;
    if ([':', '='].includes(tokens[paramsIndex]?.value)) {
      paramsIndex++;
      if (['async', 'function'].includes(tokens[paramsIndex]?.value)) paramsIndex++;
      if (tokens[paramsIndex]?.value === 'function') paramsIndex++;
    }
    const params = tokens[paramsIndex];
    if (params?.value !== '(' || params.end < 0) continue;
    const bodyIndex = params.end + (tokens[params.end + 1]?.value === '=>' ? 2 : 1);
    const body = tokens[bodyIndex];
    if (body?.value !== '{' || body.end < 0) continue;
    const standard = /^if \( ! \$\$ \( (?:"form"|'form') \) \. validate \( \) \) \{ webix \. message \( CONST \. MSG \. VL10001 \)(?: ;)? return false(?: ;)? \}/;
    const hasGuard = [...sourceEntries(tokens, bodyIndex + 1, body.end)].some(j =>
      tokens[j].value === 'if' && standard.test(tokens.slice(j, Math.min(j + 30, body.end)).map(t => t.value).join(' '))
    );
    if (!hasGuard) report('T18', token.index);
  }
}

function checkPhpSource(tokens, report) {
  for (const [i, token] of tokens.entries()) {
    if (token.quote === '"') report('T15', token.index);
    if (/^["']$/.test(token.quote) && /^\s*(?:select|with)\s/i.test(sourceLiteral(token))) {
      checkSqlSource(sourceTokens(sourceLiteral(token), 'sql', token.index + 1), report);
    }
    if (token.value.toLowerCase() === 'select' && tokens[i - 1]?.value === '->' &&
        tokens[i + 1]?.value === '(' && tokens[i + 2]?.quote &&
        /^(?:[\w.]+\.)?\*$/.test(sourceLiteral(tokens[i + 2]).trim())) report('T22', tokens[i + 2].index);

    const params = tokens[i + 1];
    if (token.value.toLowerCase() !== 'foreach' || params?.value !== '(' || params.end < 0) continue;
    const header = [...sourceEntries(tokens, i + 2, params.end)];
    const as = header.find(j => tokens[j].value.toLowerCase() === 'as');
    if (as === undefined) continue;
    const row = tokens[params.end - 1].value;
    if (!/^\$\w+$/.test(row)) continue;
    const body = tokens[params.end + 1];
    if (!body) continue;
    const start = params.end + 1;
    let end = body.end;
    if (body.value === ':') {
      let depth = 1;
      for (let j = start + 1; j < tokens.length; j++) {
        const word = tokens[j].value.toLowerCase();
        if (word === 'foreach' && tokens[tokens[j + 1]?.end + 1]?.value === ':') depth++;
        if (word === 'endforeach' && --depth === 0) { end = j; break; }
      }
    } else if (body.value !== '{') {
      for (const j of sourceEntries(tokens, start, tokens.length)) {
        if (tokens[j].value === ';') { end = j + 1; break; }
        if (tokens[j].value === '}') break;
      }
    }
    for (let j = start + 1; j < end; j++) {
      if (tokens[j].value === '=' && tokens[j + 1]?.value === row && tokens[j + 2]?.value === ';' &&
          /^(?:\$\w+|\])$/.test(tokens[j - 1].value)) report('T08', tokens[j].index);
    }
  }
}

function checkSqlSource(tokens, report) {
  for (const [i, token] of tokens.entries()) {
    if (token.value.toLowerCase() !== 'select') continue;
    let start = i + 1;
    if (/^(?:distinct|all)$/i.test(tokens[start]?.value)) start++;
    let column = [];
    for (const j of sourceEntries(tokens, start, tokens.length)) {
      const part = tokens[j];
      const stop = /^(?:from|into|select|\)|;)$/i.test(part.value);
      if (stop || part.value === ',') {
        reportWildcard(column, report);
        column = [];
        if (stop) break;
      } else {
        column.push(part);
      }
    }
    reportWildcard(column, report);
  }
}

function reportWildcard(column, report) {
  const last = column.at(-1);
  if (last?.value !== '*') return;
  const qualifier = column.slice(0, -1);
  if (qualifier.length % 2 === 0 && qualifier.every((token, i) =>
    i % 2 ? token.value === '.' : token.quote || /^\w+$/.test(token.value))) report('T22', last.index);
}

function collectSourceHits({code, type}, componentIds) {
  const hits = new Map();
  const report = (id, index, status = 'warn', message = SOURCE_RULES[id]) => {
    const key = id + ':' + status + ':' + message;
    if (!hits.has(key)) hits.set(key, {id, status, message, positions:new Set()});
    hits.get(key).positions.add(index);
  };
  for (const tokens of sourceSections(code, type)) {
    if (type === 'twig') checkTwigSource(tokens, report, componentIds);
    if (type === 'php') checkPhpSource(tokens, report);
    if (type === 'sql') checkSqlSource(tokens, report);
  }
  return hits;
}

function sourceExcerpts(code, hits) {
  const lines = code.split('\n');
  const findings = new Map();
  const sorted = [...hits].flatMap(([key, {positions}]) => [...positions].map(index => ({key, index})))
    .sort((a, b) => a.index - b.index);
  let line = 0, start = 0;
  for (const {key, index} of sorted) {
    while (line < lines.length - 1 && index >= start + lines[line].length + 1) {
      start += lines[line].length + 1;
      line++;
    }
    if (!findings.has(key)) findings.set(key, new Map());
    const excerptStart = Math.max(0, index - start - 65);
    const excerpt = lines[line].slice(excerptStart, excerptStart + 180).trim();
    findings.get(key).set(line + 1, (excerptStart ? '…' : '') + excerpt +
      (lines[line].length > excerptStart + 180 ? '…' : ''));
  }
  return findings;
}

function checkSources(files) {
  const componentIds = new Map();
  // A duplicate component ID can add a hit to an earlier file.
  const allHits = files.map(file => collectSourceHits(file, componentIds));
  const findings = new Map();
  for (const [i, file] of files.entries()) {
    const hits = allHits[i];
    for (const [key, lines] of sourceExcerpts(file.code, hits)) {
      const {id, status, message} = hits.get(key);
      if (!findings.has(key)) findings.set(key, {id, status, message, files:[]});
      findings.get(key).files.push({name:file.name, lines});
    }
  }
  const statusOrder = Object.keys(SOURCE_STATUS);
  return [...findings.values()].sort((a, b) =>
    statusOrder.indexOf(a.status) - statusOrder.indexOf(b.status) || a.id.localeCompare(b.id));
}

function coveragePercent(part) {
  return part && part.total ? Math.floor(part.covered / part.total * 1000) / 10 : null;
}

function coverageRanges(ranges, limit = 8) {
  const text = ranges.slice(0, limit).map(([a, b]) => a === b ? String(a) : a + '-' + b).join(', ');
  return ranges.length > limit ? text + ' …另 ' + (ranges.length - limit) + ' 段' : text;
}

// clover.xml 的结果 → T21 的自动检查结论（100% 只在每一行都执行到时才算）
function coverageFinding(coverage) {
  if (!coverage) return null;
  const {lines, conditionals} = coverage.totals;
  if (!lines.total) return {id:'T21', status:'unknown', message:'覆盖率文件里没有可执行的行，请确认测试范围。', files:[], extra:[]};
  const missed = coverage.files.filter(f => f.lines.covered < f.lines.total || f.conditionals.covered < f.conditionals.total);
  const branch = conditionals.total
    ? '条件分支 ' + coveragePercent(conditionals) + '%（' + conditionals.covered + ' / ' + conditionals.total + '）。'
    : 'clover.xml 不含分支数据，分支遗漏仍需走查确认。';
  if (!missed.length) {
    return {id:'T21', status:'confirmed', message:'代码行覆盖率 100%（' + lines.total + ' 行）。' + branch, files:[], extra:[]};
  }
  return {
    id:'T21', status:'warn',
    message:'代码行覆盖率 ' + coveragePercent(lines) + '%（' + lines.covered + ' / ' + lines.total + ' 行），' +
      missed.length + ' 个文件没到 100%。' + branch,
    files:[],
    extra:missed.slice(0, 5).map(f => f.short + (f.uncovered.length ? ' · 未执行第 ' + coverageRanges(f.uncovered) + ' 行' : '')
      + (f.partial?.length ? ' · 分支没走全：第 ' + coverageRanges(f.partial) + ' 行' : ''))
      .concat(missed.length > 5 ? ['另有 ' + (missed.length - 5) + ' 个文件没到 100%，见上方覆盖率面板'] : [])
  };
}

root.PclRules = {SOURCE_RULES, SOURCE_TYPES, SOURCE_STATUS, SOURCE_LIMIT, sourceType, checkSources,
  coveragePercent, coverageRanges, coverageFinding};
})(typeof window !== 'undefined' ? window : globalThis);
