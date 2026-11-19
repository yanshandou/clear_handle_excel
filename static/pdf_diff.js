// PDF 左右对比（用户 2026-11-19）：左右各放几份 PDF，按顺序接起来，像 Git 一样逐行比文字。
// 后端 core/pdf_diff.py：POST /api/pdf_diff/upload 存文件（正文就是 PDF），POST /api/pdf_diff 比，GET /api/pdf_diff/file?id= 打开。
(() => {
  const CONTEXT = 3;  // 折起一样的行时，不同处上下各留几行
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const log = (...args) => console.info('[PDF对比]', ...args);

  // 每边的文件：{key, name, id?, pages?, size?, status: 'busy' | 'ok' | 'error', error?}
  const sides = { left: [], right: [] };
  let keySeq = 0;
  let result = null;       // 最近一次对比结果
  let resultSig = '';      // 那次对比用的是哪些文件、什么顺序
  let running = false;
  let pickSide = 'left';
  let dragKey = null;      // 正在列表里拖的那一份

  const sideName = side => (side === 'left' ? '左边' : '右边');
  // 每份文件放进来时定一个字母（A、B、C…），换顺序、换边都不变，结果里的「A 3」就是 A 的第 3 页
  let letterSeq = 0;
  const nextLetter = () => { let n = letterSeq++, out = ''; do { out = String.fromCharCode(65 + (n % 26)) + out; n = Math.floor(n / 26) - 1; } while (n >= 0); return out; };
  const okFiles = side => sides[side].filter(f => f.status === 'ok');
  const signature = () => JSON.stringify([okFiles('left').map(f => f.id), okFiles('right').map(f => f.id)]);

  function sizeText(bytes) {
    if (!bytes && bytes !== 0) return '';
    return bytes >= 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`;
  }

  // ---------- 文件列表 ----------

  function renderSide(side) {
    const list = $(side === 'left' ? 'pdLeftList' : 'pdRightList');
    const files = sides[side];
    list.innerHTML = files.map((f, i) => {
      const info = f.status === 'busy' ? '<small>正在读取…</small>'
        : f.status === 'error' ? `<small class="is-error">${esc(f.error)}</small>`
        : `<small>${f.pages} 页 · ${sizeText(f.size)}</small>`;
      return `<li class="pd-item is-${f.status === 'ok' ? 'ok' : f.status}" draggable="${f.status === 'busy' ? 'false' : 'true'}" data-key="${f.key}">
        <span class="pd-letter">${esc(f.letter)}</span>
        <span class="pd-name">${esc(f.name)}${info}</span>
        <span class="pd-btns">
          <button type="button" data-act="up" title="往前挪" aria-label="${esc(f.name)} 往前挪" ${i === 0 ? 'disabled' : ''}>↑</button>
          <button type="button" data-act="down" title="往后挪" aria-label="${esc(f.name)} 往后挪" ${i === files.length - 1 ? 'disabled' : ''}>↓</button>
          <button type="button" data-act="remove" title="拿掉这一份" aria-label="拿掉 ${esc(f.name)}">✕</button>
        </span></li>`;
    }).join('');
    const ok = okFiles(side);
    const pages = ok.reduce((n, f) => n + f.pages, 0);
    $(side === 'left' ? 'pdLeftSum' : 'pdRightSum').textContent = ok.length ? `${ok.length} 份 · 共 ${pages} 页` : '';
  }

  function renderSides() {
    renderSide('left');
    renderSide('right');
    const busy = [...sides.left, ...sides.right].some(f => f.status === 'busy');
    $('pdRun').disabled = running || busy || !okFiles('left').length || !okFiles('right').length;
    if (result && signature() !== resultSig && !running) {
      $('pdStatus').innerHTML = '<span class="pd-stale">左右的文件或顺序变了，下面还是上一次的结果；点「🔍 开始对比」重新比。</span>';
    } else if (result && signature() === resultSig && !running) {
      $('pdStatus').textContent = '';
    }
  }

  function findFile(key) {
    for (const side of ['left', 'right']) {
      const index = sides[side].findIndex(f => f.key === key);
      if (index >= 0) return { side, index, file: sides[side][index] };
    }
    return null;
  }

  async function upload(file, entry) {
    try {
      const res = await fetch('/api/pdf_diff/upload', {
        method: 'POST', body: file,
        headers: { 'Content-Type': 'application/pdf', 'X-File-Name': encodeURIComponent(file.name) },
      });
      const data = await res.json().catch(() => ({ ok: false, error: `服务器没回应（${res.status}）` }));
      if (!data.ok) throw new Error(data.error || '读不了这份 PDF');
      Object.assign(entry, { id: data.id, pages: data.pages, size: data.size, status: 'ok' });
      log('收到', file.name, data.pages, '页', data.id);
    } catch (err) {
      Object.assign(entry, { status: 'error', error: err.message || String(err) });
      log('读取失败', file.name, err);
    }
    renderSides();
  }

  function addFiles(side, fileList, beforeKey = null) {
    const files = [...fileList];
    const pdfs = files.filter(f => /\.pdf$/i.test(f.name) || f.type === 'application/pdf');
    const skipped = files.length - pdfs.length;
    if (skipped) $('pdError').textContent = `有 ${skipped} 份不是 PDF，没有放进来：${files.filter(f => !pdfs.includes(f)).map(f => f.name).join('、')}`;
    else $('pdError').textContent = '';
    if (!pdfs.length) return;
    const entries = pdfs.map(f => ({ key: `f${++keySeq}`, letter: nextLetter(), name: f.name, size: f.size, status: 'busy' }));
    const list = sides[side];
    const at = beforeKey ? list.findIndex(f => f.key === beforeKey) : -1;
    list.splice(at >= 0 ? at : list.length, 0, ...entries);
    log('放到', sideName(side), pdfs.map(f => f.name));
    renderSides();
    // 一份一份传：顺序已经按放进来的排好了，传的先后不影响顺序
    (async () => { for (let i = 0; i < pdfs.length; i += 1) await upload(pdfs[i], entries[i]); })();
  }

  function moveFile(key, toSide, beforeKey) {
    const from = findFile(key);
    if (!from || key === beforeKey) return;
    sides[from.side].splice(from.index, 1);
    const list = sides[toSide];
    const at = beforeKey ? list.findIndex(f => f.key === beforeKey) : -1;
    list.splice(at >= 0 ? at : list.length, 0, from.file);
    log('挪动', from.file.name, sideName(from.side), '→', sideName(toSide));
    renderSides();
  }

  function onListClick(event) {
    const button = event.target.closest('button[data-act]');
    if (!button) return;
    const item = button.closest('.pd-item');
    const found = item && findFile(item.dataset.key);
    if (!found) return;
    const list = sides[found.side];
    if (button.dataset.act === 'remove') list.splice(found.index, 1);
    else {
      const to = found.index + (button.dataset.act === 'up' ? -1 : 1);
      if (to < 0 || to >= list.length) return;
      [list[found.index], list[to]] = [list[to], list[found.index]];
    }
    renderSides();
    // 按钮重画了，焦点放回同一份文件的同一个按钮，键盘连按不会丢
    const again = document.querySelector(`.pd-item[data-key="${found.file.key}"] button[data-act="${button.dataset.act}"]`);
    if (again && !again.disabled) again.focus({ preventScroll: true });
  }

  function clearDropMarks() {
    document.querySelectorAll('.pd-side.is-over').forEach(el => el.classList.remove('is-over'));
    document.querySelectorAll('.pd-item.is-drop-before').forEach(el => el.classList.remove('is-drop-before'));
  }

  function setupSide(box) {
    const side = box.dataset.side;
    box.addEventListener('dragover', event => {
      const isFiles = event.dataTransfer && [...event.dataTransfer.types].includes('Files');
      if (!isFiles && !dragKey) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = isFiles ? 'copy' : 'move';
      clearDropMarks();
      box.classList.add('is-over');
      const item = event.target.closest('.pd-item');
      if (item && item.dataset.key !== dragKey) item.classList.add('is-drop-before');
    });
    box.addEventListener('dragleave', event => {
      if (!box.contains(event.relatedTarget)) clearDropMarks();
    });
    box.addEventListener('drop', event => {
      event.preventDefault();
      const item = event.target.closest('.pd-item');
      const beforeKey = item ? item.dataset.key : null;
      clearDropMarks();
      if (event.dataTransfer.files && event.dataTransfer.files.length) addFiles(side, event.dataTransfer.files, beforeKey);
      else if (dragKey) moveFile(dragKey, side, beforeKey);
      dragKey = null;
    });
    const list = box.querySelector('.pd-list');
    list.addEventListener('click', onListClick);
    list.addEventListener('dragstart', event => {
      const item = event.target.closest('.pd-item');
      if (!item) return;
      dragKey = item.dataset.key;
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', item.dataset.key);
      requestAnimationFrame(() => item.classList.add('is-dragging'));
    });
    list.addEventListener('dragend', () => {
      dragKey = null;
      clearDropMarks();
      document.querySelectorAll('.pd-item.is-dragging').forEach(el => el.classList.remove('is-dragging'));
    });
  }

  // 拖到框外面松手：浏览器默认会直接打开这份 PDF、离开本页，拦下来
  window.addEventListener('dragover', event => { if (!event.target.closest('.pd-side')) event.preventDefault(); });
  window.addEventListener('drop', event => { if (!event.target.closest('.pd-side')) event.preventDefault(); });

  // ---------- 对比 ----------

  async function run() {
    if (running) return;
    const left = okFiles('left'), right = okFiles('right');
    if (!left.length || !right.length) return;
    running = true;
    $('pdError').textContent = '';
    const started = Date.now();
    $('pdStatus').innerHTML = '<span class="pd-busy"><span class="pd-spinner" aria-hidden="true"></span><span id="pdElapsed">正在抽文字、逐行对比…</span></span>';
    const timer = setInterval(() => {
      const el = $('pdElapsed');
      if (el) el.textContent = `正在抽文字、逐行对比… 已用 ${Math.round((Date.now() - started) / 1000)} 秒`;
    }, 1000);
    renderSides();
    const snapshot = { left: left.map(f => ({ ...f })), right: right.map(f => ({ ...f })) };
    const sig = signature();
    log('开始对比', snapshot.left.map(f => f.name), 'vs', snapshot.right.map(f => f.name));
    try {
      const res = await fetch('/api/pdf_diff', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          left: left.map(f => f.id), right: right.map(f => f.id),
          options: { ignore_space: $('pdIgnoreSpace').checked, ignore_header: $('pdIgnoreHeader').checked },
        }),
      });
      const data = await res.json().catch(() => ({ ok: false, error: `服务器没回应（${res.status}）` }));
      if (!data.ok) throw new Error(data.error || '对比失败');
      result = { ...data, snapshot };
      resultSig = sig;
      log('对比完成', data.stats);
      $('pdStatus').textContent = '';
      renderResult();
    } catch (err) {
      $('pdStatus').textContent = '';
      $('pdError').textContent = `对比失败：${err.message || err}`;
      log('对比失败', err);
    } finally {
      clearInterval(timer);
      running = false;
      renderSides();
    }
  }

  // ---------- 显示一次对比结果（上面拖放的一次、下面批量的每一组，都用这一套） ----------
  // 每个结果框 .pd-view 记着自己的结果：views.get(box) = {res, hunkIndex}
  const views = new WeakMap();

  // ops → 一行一行：{t: 'eq'|'mod'|'del'|'add'|'fold', l, r, hunk}
  function buildRows(res, foldSame) {
    const rows = [];
    let hunk = -1;
    const ops = res.ops;
    ops.forEach((op, opIndex) => {
      const [tag, i1, i2, j1, j2] = op;
      if (tag === 'e') {
        const n = i2 - i1;
        const keepHead = opIndex === 0 ? 0 : CONTEXT;
        const keepTail = opIndex === ops.length - 1 ? 0 : CONTEXT;
        if (foldSame && n > keepHead + keepTail + 1) {
          for (let k = 0; k < keepHead; k += 1) rows.push({ t: 'eq', l: i1 + k, r: j1 + k });
          rows.push({ t: 'fold', i1: i1 + keepHead, j1: j1 + keepHead, n: n - keepHead - keepTail });
          for (let k = n - keepTail; k < n; k += 1) rows.push({ t: 'eq', l: i1 + k, r: j1 + k });
        } else {
          for (let k = 0; k < n; k += 1) rows.push({ t: 'eq', l: i1 + k, r: j1 + k });
        }
        return;
      }
      hunk += 1;
      const start = rows.length;
      if (tag === 'r') {
        const pairs = Math.min(i2 - i1, j2 - j1);
        for (let k = 0; k < pairs; k += 1) rows.push({ t: 'mod', l: i1 + k, r: j1 + k, hunk });
        for (let k = i1 + pairs; k < i2; k += 1) rows.push({ t: 'del', l: k, r: null, hunk });
        for (let k = j1 + pairs; k < j2; k += 1) rows.push({ t: 'add', l: null, r: k, hunk });
      } else if (tag === 'd') {
        for (let k = i1; k < i2; k += 1) rows.push({ t: 'del', l: k, r: null, hunk });
      } else if (tag === 'i') {
        for (let k = j1; k < j2; k += 1) rows.push({ t: 'add', l: null, r: k, hunk });
      }
      rows[start].hunkStart = true;
    });
    return rows;
  }

  function marked(text, ranges) {
    if (!ranges || !ranges.length) return esc(text);
    let out = '', pos = 0;
    ranges.forEach(([a, b]) => { out += esc(text.slice(pos, a)) + `<mark>${esc(text.slice(a, b))}</mark>`; pos = b; });
    return out + esc(text.slice(pos));
  }

  const letterOfSnapshot = (res, side, fileIndex) => esc(res.snapshot[side][fileIndex].letter);

  // 行号格：行号 + 这一行在哪份文件第几页（换页的第一行才显示小标签）
  function noCell(res, side, index, cls, showSrc) {
    if (index === null || index === undefined) return `<td class="pd-no ${cls}"></td>`;
    const line = res[side][index];
    const file = res.snapshot[side][line.f];
    const letter = letterOfSnapshot(res, side, line.f);
    const tag = showSrc
      ? `<a class="pd-src is-${side === 'left' ? 'old' : 'new'}" href="/api/pdf_diff/file?id=${file.id}#page=${line.p}" target="_blank" rel="noopener" title="${esc(file.name)} 第 ${line.p} 页（点了在浏览器里打开这一页）">${letter} ${line.p}</a>`
      : '';
    return `<td class="pd-no ${cls}" title="${esc(file.name)} 第 ${line.p} 页">${tag}${index + 1}</td>`;
  }

  function rowHtml(res, row, prevPage) {
    if (row.t === 'fold') {
      return `<tr class="pd-fold" data-i="${row.i1}" data-j="${row.j1}" data-n="${row.n}"><td colspan="4">⋯ 一样的 ${row.n} 行 ⋯<button type="button" data-unfold>展开</button></td></tr>`;
    }
    const L = row.l !== null ? res.left[row.l] : null;
    const R = row.r !== null ? res.right[row.r] : null;
    const lPage = L ? `${L.f}:${L.p}` : null, rPage = R ? `${R.f}:${R.p}` : null;
    const lNew = L && lPage !== prevPage.left, rNew = R && rPage !== prevPage.right;
    if (L) prevPage.left = lPage;
    if (R) prevPage.right = rPage;
    const lCls = row.t === 'mod' || row.t === 'del' ? 'is-del' : row.t === 'add' ? 'is-empty' : '';
    const rCls = row.t === 'mod' || row.t === 'add' ? 'is-add' : row.t === 'del' ? 'is-empty' : '';
    const chars = row.t === 'mod' ? res.chars[`${row.l}:${row.r}`] : null;
    const attrs = [
      row.hunkStart ? `data-hunk="${row.hunk}"` : '',
      (lNew || rNew) ? 'class="pd-page"' : '',
    ].filter(Boolean).join(' ');
    return `<tr ${attrs}>${noCell(res, 'left', row.l, lCls, lNew)}<td class="pd-txt ${lCls}">${L ? marked(L.s, chars && chars[0]) : ''}</td>`
      + `${noCell(res, 'right', row.r, rCls, rNew)}<td class="pd-txt ${rCls}">${R ? marked(R.s, chars && chars[1]) : ''}</td></tr>`;
  }

  const rowsHtml = (res, rows, prevPage) => rows.map(row => rowHtml(res, row, prevPage)).join('');

  function verdictText(res) {
    return res.stats.same ? '✓ 两边的文字完全一样' : `有 ${res.stats.hunks} 处不同`;
  }

  // 把一次结果画进 box（box 带 .pd-view）；compact：批量卡片里不再重复大标题
  function renderView(box, res, compact = false) {
    box.classList.add('pd-view');
    const state = views.get(box) || {};
    views.set(box, { ...state, res, hunkIndex: -1, compact });
    const { stats, snapshot, files, dropped } = res;
    const names = side => snapshot[side].map((f, i) => `<span class="pd-hl">${letterOfSnapshot(res, side, i)}</span>${esc(f.name)}`).join(' + ');
    const pages = side => files[side].reduce((n, f) => n + f.pages, 0);
    const desc = `左边 ${snapshot.left.length} 份 ${pages('left')} 页 ${res.left.length} 行 · 右边 ${snapshot.right.length} 份 ${pages('right')} 页 ${res.right.length} 行`
      + (dropped.left || dropped.right ? ` · 没比的页眉页脚：左 ${dropped.left}、右 ${dropped.right}` : '');
    const warns = ['left', 'right'].flatMap(side => files[side].map((f, i) => (f.pages && !f.lines
      ? `${sideName(side)}的 ${letterOfSnapshot(res, side, i)} ${snapshot[side][i].name} 一个字都没抽到（可能是扫描件或整页图片），这份比不了文字。` : null)).filter(Boolean));
    const verdict = compact ? `<p class="pd-desc">${esc(desc)}</p>`
      : `<div class="pd-verdict ${stats.same ? 'is-same' : 'is-diff'}">${verdictText(res)}<small>${esc(desc)}</small></div>`;
    const chips = stats.same ? '' : [
      stats.changed ? `<span class="pd-chip is-mod">改了 ${stats.changed} 行</span>` : '',
      stats.removed ? `<span class="pd-chip is-del">左边多 ${stats.removed} 行</span>` : '',
      stats.added ? `<span class="pd-chip is-add">右边多 ${stats.added} 行</span>` : '',
    ].join('');
    const fold = box.dataset.fold !== '0';
    const bar = `<div class="pd-bar">
        <button type="button" class="btn btn-outline btn-sm" data-nav="-1" ${stats.same ? 'disabled' : ''}>▲ 上一处</button>
        <button type="button" class="btn btn-outline btn-sm" data-nav="1" ${stats.same ? 'disabled' : ''}>▼ 下一处</button>
        <span class="pd-pos">${stats.same ? '没有不同' : `共 ${stats.hunks} 处`}</span>
        <label><input type="checkbox" class="pd-fold-toggle" ${fold ? 'checked' : ''}> 只看不同（一样的行折起来）</label>
        ${chips}
      </div>`;
    const rows = buildRows(res, fold);
    const prevPage = { left: null, right: null };
    box.innerHTML = verdict + warns.map(w => `<p class="pd-warn">⚠️ ${esc(w)}</p>`).join('') + bar
      + `<div class="pd-heads"><div class="is-old">左边：${names('left')}</div><div class="is-new">右边：${names('right')}</div></div>`
      + `<table class="pd-table"><colgroup><col class="c-no"><col><col class="c-no"><col></colgroup><tbody class="pd-rows">${
        rows.length ? rowsHtml(res, rows, prevPage) : '<tr><td colspan="4" class="pd-empty-result">两边都没有抽到文字。</td></tr>'}</tbody></table>`;
  }

  function renderResult() {
    const box = $('pdResult');
    if (!result) { box.classList.remove('pd-view'); box.innerHTML = '<p class="pd-empty-result">左右两边都放好 PDF 后点「🔍 开始对比」。</p>'; return; }
    renderView(box, result);
  }

  function unfold(box, tr) {
    const { res } = views.get(box);
    const i1 = +tr.dataset.i, j1 = +tr.dataset.j, n = +tr.dataset.n;
    const rows = [];
    for (let k = 0; k < n; k += 1) rows.push({ t: 'eq', l: i1 + k, r: j1 + k });
    // 接着上一行的页：上一行是哪页，用来决定要不要显示换页标签
    const prevPage = { left: null, right: null };
    if (i1 > 0) { const p = res.left[i1 - 1]; prevPage.left = `${p.f}:${p.p}`; }
    if (j1 > 0) { const p = res.right[j1 - 1]; prevPage.right = `${p.f}:${p.p}`; }
    tr.insertAdjacentHTML('afterend', rowsHtml(res, rows, prevPage));
    tr.remove();
  }

  function navigate(box, step) {
    const state = views.get(box);
    const starts = [...box.querySelectorAll('.pd-rows tr[data-hunk]')];
    if (!starts.length) return;
    const i = state.hunkIndex;
    state.hunkIndex = i < 0 ? (step > 0 ? 0 : starts.length - 1) : (i + step + starts.length) % starts.length;
    box.querySelectorAll('.pd-rows tr.is-current').forEach(tr => tr.classList.remove('is-current'));
    const target = starts[state.hunkIndex];
    target.classList.add('is-current');
    target.scrollIntoView({ block: 'center' });  // 只在点按钮时滚过去
    box.querySelector('.pd-pos').textContent = `第 ${state.hunkIndex + 1} / ${starts.length} 处`;
  }

  document.addEventListener('click', event => {
    const box = event.target.closest('.pd-view');
    if (!box || !views.get(box)) return;
    const unfoldBtn = event.target.closest('[data-unfold]');
    if (unfoldBtn) { unfold(box, unfoldBtn.closest('tr')); return; }
    const nav = event.target.closest('[data-nav]');
    if (nav) navigate(box, +nav.dataset.nav);
  });
  document.addEventListener('change', event => {
    if (!event.target.classList.contains('pd-fold-toggle')) return;
    const box = event.target.closest('.pd-view');
    const state = box && views.get(box);
    if (!state) return;
    box.dataset.fold = event.target.checked ? '1' : '0';
    renderView(box, state.res, state.compact);
  });

  // 选项变了：已经比过就按新选项重比一次（文件已经在程序那边，很快）
  ['pdIgnoreSpace', 'pdIgnoreHeader'].forEach(id => $(id).addEventListener('change', () => {
    if (result && signature() === resultSig) run();
  }));

  $('pdRun').addEventListener('click', run);
  $('pdSwap').addEventListener('click', () => {
    [sides.left, sides.right] = [sides.right, sides.left];
    log('左右互换');
    renderSides();
  });
  $('pdClear').addEventListener('click', async () => {
    if (!sides.left.length && !sides.right.length && !result) return;
    if (!(await appConfirm('拿掉左右两边所有的 PDF，并清掉下面的对比结果？', { title: '全部清空', okText: '清空', tone: 'danger' }))) return;
    sides.left = [];
    sides.right = [];
    letterSeq = 0;
    result = null;
    resultSig = '';
    $('pdStatus').textContent = '';
    $('pdError').textContent = '';
    renderSides();
    renderResult();
  });
  document.querySelectorAll('[data-pick]').forEach(button => button.addEventListener('click', () => {
    pickSide = button.dataset.pick;
    $('pdFileInput').click();
  }));
  $('pdFileInput').addEventListener('change', event => {
    addFiles(pickSide, event.target.files);
    event.target.value = '';
  });
  // ---------- 批量对比：贴路径（用户 2026-11-19：A+B 比 C、C+D 比 E，一页里看完） ----------
  const BATCH_KEY = 'clear_handle_excel.pdfDiffBatch';
  const isDir = t => /[\\/]/.test(t);
  const dirOf = t => t.replace(/[\\/][^\\/]*$/, '');
  const baseName = t => t.replace(/^.*[\\/]/, '');
  const unquote = t => t.trim().replace(/^["'“”「]+|["'“”」]+$/g, '').trim();

  // 一边：只有紧跟 .pdf 或前后有空格的 + 才算分隔（文件名里本身可以有 +）
  const splitSide = text => text.split(/(?<=\.pdf["'”」]?)\s*\+\s*|\s+\+\s+/i).map(unquote).filter(Boolean);

  // 文字 → [{line, left:[路径], right:[路径], error?}]
  function parseBatch(text) {
    const trimmed = text.trim();
    if (!trimmed) return [];
    if (/^[[{]/.test(trimmed)) {
      let data;
      try { data = JSON.parse(trimmed); } catch (err) { throw new Error(`JSON 格式不对：${err.message}`); }
      const list = Array.isArray(data) ? data : [data];
      const side = v => (Array.isArray(v) ? v : [v]).filter(x => typeof x === 'string' && x.trim()).map(unquote);
      return list.map((item, i) => {
        const pair = Array.isArray(item) ? { left: item[0], right: item[1] } : (item || {});
        const g = { line: i + 1, left: side(pair.left ?? []), right: side(pair.right ?? []) };
        if (!g.left.length || !g.right.length) g.error = '左边或右边没有 PDF 路径';
        return g;
      });
    }
    const groups = [];
    let lastDir = '';
    text.split(/\r?\n/).forEach((raw, i) => {
      const line = raw.trim();
      if (!line || line.startsWith('#')) return;
      const parts = line.split(/(?<=\.pdf["'”」]?)\s*=\s*|\s+=\s+/i);
      if (parts.length !== 2) {
        groups.push({ line: i + 1, left: [], right: [], text: line, error: parts.length < 2 ? '这一行没有「=」，分不出左右' : '这一行有好几个「=」，分不出左右' });
        return;
      }
      const fill = list => list.map(p => {
        const full = isDir(p) || !lastDir ? p : `${lastDir}${lastDir.includes('\\') ? '\\' : '/'}${p}`;
        if (isDir(full)) lastDir = dirOf(full);
        return full;
      });
      const left = fill(splitSide(parts[0])), right = fill(splitSide(parts[1]));
      const g = { line: i + 1, left, right };
      if (!left.length || !right.length) g.error = '左边或右边没有 PDF 路径';
      groups.push(g);
    });
    return groups;
  }

  let batchRunning = false;
  const batchViews = new Map();  // 卡片 → 结果（第一次展开时才画，几十组也不卡）

  function groupName(g) {
    if (!g.left.length && !g.right.length) return esc(g.text || '');
    const side = list => list.map(p => esc(baseName(p))).join(' + ');
    return `${side(g.left)}<span class="pd-eq">=</span>${side(g.right)}`;
  }

  function setGroup(card, cls, state, small) {
    card.classList.remove('is-wait', 'is-busy', 'is-same', 'is-diff', 'is-bad');
    card.classList.add(cls);
    card.querySelector('.pd-gstate').textContent = state;
    if (small !== undefined) card.querySelector('.pd-gname small').textContent = small;
  }

  function batchSummary(counts, total, finished) {
    const box = $('pdBatchSum');
    box.hidden = false;
    box.innerHTML = `<span>${finished ? `${total} 组比完` : `共 ${total} 组`}</span>`
      + `<span class="is-same">✓ 一样 ${counts.same}</span><span class="is-diff">有不同 ${counts.diff}</span>`
      + (counts.bad ? `<span class="is-bad">出错 ${counts.bad}</span>` : '');
  }

  async function postJson(url, body) {
    const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const data = await res.json().catch(() => ({ ok: false, error: `服务器没回应（${res.status}）` }));
    if (!data.ok) throw new Error(data.error || '出错了');
    return data;
  }

  async function runBatch() {
    if (batchRunning) return;
    const text = $('pdBatchText').value;
    try { localStorage.setItem(BATCH_KEY, text); } catch (_) {}
    $('pdBatchError').textContent = '';
    let groups;
    try { groups = parseBatch(text); } catch (err) { $('pdBatchError').textContent = err.message; return; }
    if (!groups.length) { $('pdBatchError').textContent = '还没有贴路径：一行一组，例如 C:\\资料\\A.pdf + C:\\资料\\B.pdf = C:\\资料\\C.pdf'; return; }
    batchRunning = true;
    $('pdBatchRun').disabled = true;
    batchViews.clear();
    const results = $('pdBatchResults');
    results.innerHTML = groups.map((g, i) => `<details class="pd-group is-wait" data-g="${i}">
        <summary><span class="pd-gno">${i + 1}</span><span class="pd-gname">${groupName(g)}<small>第 ${g.line} 行</small></span><span class="pd-gstate">等待中</span></summary>
        <div class="pd-gbody"></div></details>`).join('');
    $('pdBatchTools').hidden = false;
    const cards = [...results.querySelectorAll('.pd-group')];
    const counts = { same: 0, diff: 0, bad: 0 };
    const stored = new Map();  // 路径 → 存好的文件（同一份在几组里出现只读一次）
    const letters = new Map(); // 路径 → 字母，整批里同一份文件字母一样
    const letterFor = p => { if (!letters.has(p)) letters.set(p, nextBatchLetter(letters.size)); return letters.get(p); };
    const options = { ignore_space: $('pdIgnoreSpace').checked, ignore_header: $('pdIgnoreHeader').checked };
    const track = $('pdBatchTrack'), fill = $('pdBatchFill');
    track.hidden = false;
    const progress = done => {
      fill.style.width = `${Math.round((done / groups.length) * 100)}%`;
      if (window.ProgressEta) ProgressEta.update(fill, 'pdf_diff_batch', done / groups.length, true);
    };
    progress(0);
    batchSummary(counts, groups.length, false);
    log('批量对比开始', groups.length, '组');
    for (let i = 0; i < groups.length; i += 1) {
      const g = groups[i], card = cards[i];
      $('pdBatchStatus').innerHTML = `<span class="pd-busy"><span class="pd-spinner" aria-hidden="true"></span>正在比第 ${i + 1} / ${groups.length} 组…</span>`;
      try {
        if (g.error) throw new Error(g.error);
        setGroup(card, 'is-busy', '正在读取…');
        const files = {};
        for (const side of ['left', 'right']) {
          files[side] = [];
          for (const p of g[side]) {
            if (!stored.has(p)) stored.set(p, postJson('/api/pdf_diff/path', { path: p }));
            const info = await stored.get(p);
            files[side].push({ ...info, letter: letterFor(p), status: 'ok' });
          }
        }
        setGroup(card, 'is-busy', '正在对比…');
        const data = await postJson('/api/pdf_diff', { left: files.left.map(f => f.id), right: files.right.map(f => f.id), options });
        const res = { ...data, snapshot: files };
        batchViews.set(card, res);
        const pages = side => data.files[side].reduce((n, f) => n + f.pages, 0);
        const noText = ['left', 'right'].some(side => data.files[side].some(f => f.pages && !f.lines));
        const small = `第 ${g.line} 行 · 左 ${pages('left')} 页、右 ${pages('right')} 页${noText ? ' · ⚠️ 有文件抽不到文字' : ''}`;
        if (data.stats.same) { counts.same += 1; setGroup(card, 'is-same', '✓ 一样', small); }
        else { counts.diff += 1; setGroup(card, 'is-diff', `有 ${data.stats.hunks} 处不同`, small); }
        if (card.open) renderView(card.querySelector('.pd-gbody'), res, true);
        log('第', i + 1, '组', data.stats);
      } catch (err) {
        counts.bad += 1;
        // 读不了的那份下次重新读
        g.left.concat(g.right).forEach(p => { const v = stored.get(p); if (v) v.catch(() => stored.delete(p)); });
        setGroup(card, 'is-bad', '出错');
        card.querySelector('.pd-gbody').innerHTML = `<p class="pd-error">${esc(err.message || err)}</p>`;
        log('第', i + 1, '组出错', err);
      }
      progress(i + 1);
      batchSummary(counts, groups.length, false);
    }
    batchSummary(counts, groups.length, true);
    if (window.ProgressEta) ProgressEta.update(fill, 'pdf_diff_batch', 1, false);
    track.hidden = true;
    $('pdBatchStatus').textContent = counts.bad ? `! ${groups.length} 组比完，${counts.bad} 组出错，点开卡片看原因。` : `✓ ${groups.length} 组都比完了。`;
    batchRunning = false;
    $('pdBatchRun').disabled = false;
    log('批量对比完成', counts);
  }

  function nextBatchLetter(n) {
    let out = '';
    do { out = String.fromCharCode(65 + (n % 26)) + out; n = Math.floor(n / 26) - 1; } while (n >= 0);
    return out;
  }

  // 第一次展开时才画
  $('pdBatchResults').addEventListener('toggle', event => {
    const card = event.target;
    if (!card.classList || !card.classList.contains('pd-group') || !card.open) return;
    const body = card.querySelector('.pd-gbody');
    const res = batchViews.get(card);
    if (res && !views.get(body)) renderView(body, res, true);
  }, true);
  $('pdBatchRun').addEventListener('click', runBatch);
  $('pdOpenDiff').addEventListener('click', () => {
    document.querySelectorAll('.pd-group.is-diff, .pd-group.is-bad').forEach(card => { card.open = true; });
  });
  $('pdCloseAll').addEventListener('click', () => {
    document.querySelectorAll('.pd-group[open]').forEach(card => { card.open = false; });
  });
  try { const saved = localStorage.getItem(BATCH_KEY); if (saved) $('pdBatchText').value = saved; } catch (_) {}

  document.querySelectorAll('.pd-side').forEach(setupSide);
  renderSides();
})();
