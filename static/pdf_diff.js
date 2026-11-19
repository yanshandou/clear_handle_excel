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
  let hunkIndex = -1;      // 「上一处 / 下一处」现在停在第几处
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
      hunkIndex = -1;
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

  // ops → 一行一行：{t: 'eq'|'mod'|'del'|'add'|'fold', l, r, hunk}
  function buildRows(foldSame) {
    const rows = [];
    let hunk = -1;
    const ops = result.ops;
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

  // 行号格：行号 + 这一行在哪份文件第几页（换页的第一行才显示小标签）
  function noCell(side, index, cls, showSrc) {
    if (index === null || index === undefined) return `<td class="pd-no ${cls}"></td>`;
    const line = result[side][index];
    const file = result.snapshot[side][line.f];
    const letter = letterOfSnapshot(side, line.f);
    const tag = showSrc
      ? `<a class="pd-src is-${side === 'left' ? 'old' : 'new'}" href="/api/pdf_diff/file?id=${file.id}#page=${line.p}" target="_blank" rel="noopener" title="${esc(file.name)} 第 ${line.p} 页（点了在浏览器里打开这一页）">${letter} ${line.p}</a>`
      : '';
    return `<td class="pd-no ${cls}" title="${esc(file.name)} 第 ${line.p} 页">${tag}${index + 1}</td>`;
  }

  function letterOfSnapshot(side, fileIndex) {
    return esc(result.snapshot[side][fileIndex].letter);
  }

  function rowHtml(row, prevPage) {
    if (row.t === 'fold') {
      return `<tr class="pd-fold" data-i="${row.i1}" data-j="${row.j1}" data-n="${row.n}"><td colspan="4">⋯ 一样的 ${row.n} 行 ⋯<button type="button" data-unfold>展开</button></td></tr>`;
    }
    const L = row.l !== null ? result.left[row.l] : null;
    const R = row.r !== null ? result.right[row.r] : null;
    const lPage = L ? `${L.f}:${L.p}` : null, rPage = R ? `${R.f}:${R.p}` : null;
    const lNew = L && lPage !== prevPage.left, rNew = R && rPage !== prevPage.right;
    if (L) prevPage.left = lPage;
    if (R) prevPage.right = rPage;
    const lCls = row.t === 'mod' || row.t === 'del' ? 'is-del' : row.t === 'add' ? 'is-empty' : '';
    const rCls = row.t === 'mod' || row.t === 'add' ? 'is-add' : row.t === 'del' ? 'is-empty' : '';
    const chars = row.t === 'mod' ? result.chars[`${row.l}:${row.r}`] : null;
    const attrs = [
      row.hunkStart ? `data-hunk="${row.hunk}"` : '',
      (lNew || rNew) ? 'class="pd-page"' : '',
    ].filter(Boolean).join(' ');
    return `<tr ${attrs}>${noCell('left', row.l, lCls, lNew)}<td class="pd-txt ${lCls}">${L ? marked(L.s, chars && chars[0]) : ''}</td>`
      + `${noCell('right', row.r, rCls, rNew)}<td class="pd-txt ${rCls}">${R ? marked(R.s, chars && chars[1]) : ''}</td></tr>`;
  }

  function rowsHtml(rows, prevPage) {
    return rows.map(row => rowHtml(row, prevPage)).join('');
  }

  function renderResult() {
    const box = $('pdResult');
    if (!result) { box.innerHTML = '<p class="pd-empty-result">左右两边都放好 PDF 后点「🔍 开始对比」。</p>'; return; }
    const { stats, snapshot, files, dropped } = result;
    const names = side => snapshot[side].map((f, i) => `<span class="pd-hl">${letterOfSnapshot(side, i)}</span>${esc(f.name)}`).join(' + ');
    const pages = side => files[side].reduce((n, f) => n + f.pages, 0);
    const desc = `左边 ${snapshot.left.length} 份 ${pages('left')} 页 ${result.left.length} 行 · 右边 ${snapshot.right.length} 份 ${pages('right')} 页 ${result.right.length} 行`
      + (dropped.left || dropped.right ? ` · 没比的页眉页脚：左 ${dropped.left}、右 ${dropped.right}` : '');
    const warns = ['left', 'right'].flatMap(side => files[side].map((f, i) => (f.pages && !f.lines
      ? `${sideName(side)}的 ${letterOfSnapshot(side, i)} ${snapshot[side][i].name} 一个字都没抽到（可能是扫描件或整页图片），这份比不了文字。` : null)).filter(Boolean));
    const verdict = stats.same
      ? `<div class="pd-verdict is-same">✓ 两边的文字完全一样<small>${esc(desc)}</small></div>`
      : `<div class="pd-verdict is-diff">有 ${stats.hunks} 处不同<small>${esc(desc)}</small></div>`;
    const chips = stats.same ? '' : [
      stats.changed ? `<span class="pd-chip is-mod">改了 ${stats.changed} 行</span>` : '',
      stats.removed ? `<span class="pd-chip is-del">左边多 ${stats.removed} 行</span>` : '',
      stats.added ? `<span class="pd-chip is-add">右边多 ${stats.added} 行</span>` : '',
    ].join('');
    const fold = box.dataset.fold !== '0';
    const bar = `<div class="pd-bar">
        <button type="button" class="btn btn-outline btn-sm" data-nav="-1" ${stats.same ? 'disabled' : ''}>▲ 上一处</button>
        <button type="button" class="btn btn-outline btn-sm" data-nav="1" ${stats.same ? 'disabled' : ''}>▼ 下一处</button>
        <span class="pd-pos" id="pdPos">${stats.same ? '没有不同' : `共 ${stats.hunks} 处`}</span>
        <label><input type="checkbox" id="pdFold" ${fold ? 'checked' : ''}> 只看不同（一样的行折起来）</label>
        ${chips}
      </div>`;
    const rows = buildRows(fold);
    const prevPage = { left: null, right: null };
    box.innerHTML = verdict + warns.map(w => `<p class="pd-warn">⚠️ ${esc(w)}</p>`).join('') + bar
      + `<div class="pd-heads"><div class="is-old">左边：${names('left')}</div><div class="is-new">右边：${names('right')}</div></div>`
      + `<table class="pd-table"><colgroup><col class="c-no"><col><col class="c-no"><col></colgroup><tbody id="pdBody">${rowsHtml(rows, prevPage)}</tbody></table>`;
    if (!rows.length) $('pdBody').innerHTML = '<tr><td colspan="4" class="pd-empty-result">两边都没有抽到文字。</td></tr>';
  }

  function unfold(tr) {
    const i1 = +tr.dataset.i, j1 = +tr.dataset.j, n = +tr.dataset.n;
    const rows = [];
    for (let k = 0; k < n; k += 1) rows.push({ t: 'eq', l: i1 + k, r: j1 + k });
    // 接着上一行的页：上一行是哪页，用来决定要不要显示换页标签
    const prevPage = { left: null, right: null };
    if (i1 > 0) { const p = result.left[i1 - 1]; prevPage.left = `${p.f}:${p.p}`; }
    if (j1 > 0) { const p = result.right[j1 - 1]; prevPage.right = `${p.f}:${p.p}`; }
    tr.insertAdjacentHTML('afterend', rowsHtml(rows, prevPage));
    tr.remove();
  }

  function navigate(step) {
    const starts = [...document.querySelectorAll('#pdBody tr[data-hunk]')];
    if (!starts.length) return;
    hunkIndex = hunkIndex < 0 ? (step > 0 ? 0 : starts.length - 1) : (hunkIndex + step + starts.length) % starts.length;
    document.querySelectorAll('#pdBody tr.is-current').forEach(tr => tr.classList.remove('is-current'));
    const target = starts[hunkIndex];
    target.classList.add('is-current');
    target.scrollIntoView({ block: 'center' });  // 只在点按钮时滚过去
    $('pdPos').textContent = `第 ${hunkIndex + 1} / ${starts.length} 处`;
  }

  $('pdResult').addEventListener('click', event => {
    const unfoldBtn = event.target.closest('[data-unfold]');
    if (unfoldBtn) { unfold(unfoldBtn.closest('tr')); return; }
    const nav = event.target.closest('[data-nav]');
    if (nav) navigate(+nav.dataset.nav);
  });
  $('pdResult').addEventListener('change', event => {
    if (event.target.id !== 'pdFold') return;
    $('pdResult').dataset.fold = event.target.checked ? '1' : '0';
    hunkIndex = -1;
    renderResult();
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
  document.querySelectorAll('.pd-side').forEach(setupSide);
  renderSides();
})();
