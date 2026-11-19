// PDF 左右对比（用户 2026-11-19）：左右各放几份 PDF，按顺序接起来，像 Git 一样逐行比文字。
// 后端 core/pdf_diff.py：POST /api/pdf_diff/upload 存文件（正文就是 PDF），POST /api/pdf_diff 比，GET /api/pdf_diff/file?id= 打开。
(() => {
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

  // ---------- 逐页看原图对比（用户 2026-11-19：两边就是两张图，不一样的地方用红绿框圈出来） ----------
  // 两边的页按顺序接起来，第 1 页对第 1 页、第 2 页对第 2 页……；每对在网页上逐像素比，
  // 不一样的地方连成一块块，左边画红框、右边画绿框。一边多出来的页整页盖色。
  const BLOCK = 6;            // 6 像素一格判断有没有不同
  const PIXEL_TOL = 48;       // 颜色差这么多以内算一样（抗锯齿、压缩噪点）
  const VISIBLE_TOL = 72;     // 超过这个、而且周围 1 像素内也找不到相近颜色，才算看得出的不同
  const GROW = 3;             // 相隔 3 格以内的不同处并成一个框
  const views = new WeakMap(); // 结果框 → {hunkIndex}
  let generation = 0;          // 重新对比时，上一次还没比完的就不要了

  const pageUrl = (id, p) => `/api/pdf_diff/page?id=${encodeURIComponent(id)}&p=${p}`;
  // 每页页码在哪（A+B 和 C 比，页码本来就不一样，用户 2026-11-19 说不要圈）；每份文件只取一次
  const pageNoCache = new Map();
  function pageNoBoxes(item) {
    if (!item) return Promise.resolve([]);
    if (!pageNoCache.has(item.id)) {
      pageNoCache.set(item.id, fetch(`/api/pdf_diff/pageno?id=${encodeURIComponent(item.id)}`)
        .then(r => r.json()).then(d => (d && d.ok ? d.pages : []))
        .catch(err => { log('页码位置取不到，页码照样比', item.name, err); return []; }));
    }
    return pageNoCache.get(item.id).then(pages => pages[item.p - 1] || []);
  }
  const fileUrl = (id, p) => `/api/pdf_diff/file?id=${encodeURIComponent(id)}#page=${p}`;
  function loadImage(url) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('页面图取不到（文件可能过期了，请重新放进来）'));
      img.src = url;
    });
  }

  // 一边所有的页：[{id, name, letter, p}]
  const pageList = files => files.flatMap(f => Array.from({ length: f.pages }, (_, i) => ({ id: f.id, name: f.name, letter: f.letter, p: i + 1 })));

  const dist = (a, i, b, j) => Math.abs(a[i] - b[j]) + Math.abs(a[i + 1] - b[j + 1]) + Math.abs(a[i + 2] - b[j + 2]);

  // 看得出的不同在哪些格（和 Excel 对比「真实效果」同一个判断法：容忍 1 像素偏移和抗锯齿）
  function detect(a, b, width, height) {
    const cols = Math.ceil(width / BLOCK), rows = Math.ceil(height / BLOCK);
    const unmatched = new Uint8Array(width * height);
    let any = false;
    for (let y = 0; y < height; y++) {
      const top = Math.max(y - 1, 0), bottom = Math.min(y + 1, height - 1);
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4;
        const d = dist(a, i, b, i);
        if (d <= VISIBLE_TOL) continue;
        let oldOk = false, newOk = false;
        const left = Math.max(x - 1, 0), right = Math.min(x + 1, width - 1);
        for (let ny = top; ny <= bottom && !(oldOk && newOk); ny++) {
          for (let j = (ny * width + left) * 4, end = (ny * width + right) * 4; j <= end; j += 4) {
            if (j === i) continue;
            oldOk ||= dist(a, i, b, j) <= VISIBLE_TOL;
            newOk ||= dist(b, i, a, j) <= VISIBLE_TOL;
            if (oldOk && newOk) break;
          }
        }
        if (!oldOk || !newOk) { unmatched[i / 4] = 1; any = true; }
      }
    }
    const hits = new Uint8Array(cols * rows);
    if (!any) return { cols, rows, hits, count: 0 };
    let count = 0;
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const i = y * width + x;
        if (!unmatched[i]) continue;
        // 孤立的一个点（噪点）不算
        const l = x > 0, r = x + 1 < width;
        const near = (l && unmatched[i - 1]) || (r && unmatched[i + 1])
          || (y > 0 && (unmatched[i - width] || (l && unmatched[i - width - 1]) || (r && unmatched[i - width + 1])))
          || (y + 1 < height && (unmatched[i + width] || (l && unmatched[i + width - 1]) || (r && unmatched[i + width + 1])));
        if (!near) continue;
        const k = Math.floor(y / BLOCK) * cols + Math.floor(x / BLOCK);
        if (!hits[k]) { hits[k] = 1; count++; }
      }
    }
    return { cols, rows, hits, count };
  }

  // 有不同的格连成一块块（隔 GROW 格以内算一块），每块一个框；坐标是占整页的比例
  function boxesOf({ cols, rows, hits }, width, height) {
    const seen = new Uint8Array(cols * rows), boxes = [];
    for (let k = 0; k < hits.length; k++) {
      if (!hits[k] || seen[k]) continue;
      let x0 = k % cols, x1 = x0, y0 = Math.floor(k / cols), y1 = y0;
      const stack = [k];
      seen[k] = 1;
      while (stack.length) {
        const c = stack.pop(), cx = c % cols, cy = Math.floor(c / cols);
        if (cx < x0) x0 = cx; if (cx > x1) x1 = cx; if (cy < y0) y0 = cy; if (cy > y1) y1 = cy;
        for (let ny = Math.max(cy - GROW, 0); ny <= Math.min(cy + GROW, rows - 1); ny++) {
          for (let nx = Math.max(cx - GROW, 0); nx <= Math.min(cx + GROW, cols - 1); nx++) {
            const n = ny * cols + nx;
            if (hits[n] && !seen[n]) { seen[n] = 1; stack.push(n); }
          }
        }
      }
      const pad = 4;
      const left = Math.max(x0 * BLOCK - pad, 0), top = Math.max(y0 * BLOCK - pad, 0);
      const right = Math.min((x1 + 1) * BLOCK + pad, width), bottom = Math.min((y1 + 1) * BLOCK + pad, height);
      boxes.push({ x: left / width, y: top / height, w: (right - left) / width, h: (bottom - top) / height });
    }
    return boxes;
  }

  // 比一对页：两张图缩放到同样宽，逐像素比
  async function comparePair(L, R) {
    const [a, b, noL, noR] = await Promise.all([L ? loadImage(pageUrl(L.id, L.p)) : null, R ? loadImage(pageUrl(R.id, R.p)) : null,
      pageNoBoxes(L), pageNoBoxes(R)]);
    const width = Math.max(a ? a.naturalWidth : 0, b ? b.naturalWidth : 0);
    const scaled = img => (img ? Math.round(img.naturalHeight * width / img.naturalWidth) : 0);
    const height = Math.max(scaled(a), scaled(b));
    if (!a || !b) return { kind: a ? 'left' : 'right', width, height, boxes: [] };
    const pixels = img => {
      const c = document.createElement('canvas');
      c.width = width; c.height = height;
      const g = c.getContext('2d', { willReadFrequently: true });
      g.fillStyle = '#fff'; g.fillRect(0, 0, width, height);
      g.drawImage(img, 0, 0, width, scaled(img));
      // 两边页码的位置都涂白再比，页码不一样不算不同
      const pad = 4;
      [[noL, scaled(a)], [noR, scaled(b)]].forEach(([list, h]) => list.forEach(([x0, y0, x1, y1]) => {
        g.fillRect(x0 * width - pad, y0 * h - pad, (x1 - x0) * width + 2 * pad, (y1 - y0) * h + 2 * pad);
      }));
      return g.getImageData(0, 0, width, height).data;
    };
    const found = detect(pixels(a), pixels(b), width, height);
    const boxes = found.count ? boxesOf(found, width, height) : [];
    return { kind: boxes.length ? 'diff' : 'same', width, height, boxes };
  }

  function frameHtml(side, page, pair) {
    const cls = side === 'left' ? 'is-old' : 'is-new';
    if (!page) return `<div class="pd-frame ${cls} is-missing" style="aspect-ratio:${pair.width} / ${pair.height}"><span>${side === 'left' ? '左边' : '右边'}没有这一页</span></div>`;
    const boxes = pair.kind === 'left' || pair.kind === 'right'
      ? '<span class="pd-box is-whole"></span>'
      : pair.boxes.map(b => `<span class="pd-box" data-b="${b.x},${b.y},${b.w},${b.h}" style="left:${(b.x * 100).toFixed(2)}%;top:${(b.y * 100).toFixed(2)}%;width:${(b.w * 100).toFixed(2)}%;height:${(b.h * 100).toFixed(2)}%"></span>`).join('');
    return `<a class="pd-frame ${cls}" href="${fileUrl(page.id, page.p)}" target="_blank" rel="noopener" title="${esc(page.name)} 第 ${page.p} 页（点了在浏览器里打开这一页）" style="aspect-ratio:${pair.width} / ${pair.height}">`
      + `<img src="${pageUrl(page.id, page.p)}" alt="${esc(page.name)} 第 ${page.p} 页" loading="lazy" draggable="false">${boxes}</a>`;
  }

  function pairHtml(index, L, R, pair) {
    const where = page => (page ? `<span class="pd-hl ${page === L ? 'pd-hl-old' : 'pd-hl-new'}">${esc(page.letter)}</span>第 ${page.p} 页` : '—');
    const state = pair.kind === 'same' ? '一样' : pair.kind === 'diff' ? `${pair.boxes.length} 处不同`
      : pair.kind === 'left' ? '只有左边有' : '只有右边有';
    return `<div class="pd-pg is-${pair.kind}" data-index="${index}" data-w="${pair.width}" data-h="${pair.height}">
        <div class="pd-pg-title"><b>第 ${index + 1} 页</b><span>左：${where(L)}</span><span>右：${where(R)}</span><span class="pd-pg-state">${state}</span></div>
        <div class="pd-pg-cols">${frameHtml('left', L, pair)}${frameHtml('right', R, pair)}</div></div>`;
  }

  // 把一次对比画进 box：先出框架，一页一页比完一页加一页
  async function compareInto(box, snapshot, { compact = false, onProgress = null } = {}) {
    const gen = compact ? 0 : ++generation;  // 批量的每组各比各的，不和上面拖放的抢
    box.classList.add('pd-view');
    box.classList.remove('show-same');
    views.set(box, { hunkIndex: -1, gen });
    const left = pageList(snapshot.left), right = pageList(snapshot.right);
    const total = Math.max(left.length, right.length);
    const names = side => snapshot[side].map(f => `<span class="pd-hl ${side === 'left' ? 'pd-hl-old' : 'pd-hl-new'}">${esc(f.letter)}</span>${esc(f.name)}`).join(' + ');
    box.innerHTML = `<div class="pd-verdict-slot"></div>
      <div class="pd-bar">
        <button type="button" class="btn btn-outline btn-sm" data-nav="-1" disabled>▲ 上一处</button>
        <button type="button" class="btn btn-outline btn-sm" data-nav="1" disabled>▼ 下一处</button>
        <span class="pd-pos">正在比…</span>
        <label><input type="checkbox" class="pd-same-toggle"> 一样的页也显示 <span class="pd-same-count"></span></label>
        <span class="pd-chips"></span>
      </div>
      <div class="pd-heads"><div class="is-old">左边：${names('left')}（${left.length} 页）</div><div class="is-new">右边：${names('right')}（${right.length} 页）</div></div>
      <div class="pd-pages"></div>
      <p class="pd-empty-result pd-all-same" hidden>每一页都一样。勾选上面的「一样的页也显示」可以看原图。</p>`;
    const pagesBox = box.querySelector('.pd-pages');
    const stats = { total, same: 0, diff: 0, onlyLeft: 0, onlyRight: 0, boxes: 0 };
    const paintStats = done => {
      const differ = stats.diff + stats.onlyLeft + stats.onlyRight;
      box.querySelector('.pd-same-count').textContent = `（${stats.same} 页）`;
      box.querySelector('.pd-chips').innerHTML = [
        stats.diff ? `<span class="pd-chip is-mod">${stats.diff} 页有不同</span>` : '',
        stats.onlyLeft ? `<span class="pd-chip is-del">左边多 ${stats.onlyLeft} 页</span>` : '',
        stats.onlyRight ? `<span class="pd-chip is-add">右边多 ${stats.onlyRight} 页</span>` : '',
      ].join('');
      box.querySelectorAll('[data-nav]').forEach(b => { b.disabled = !differ; });
      if (!done) { box.querySelector('.pd-pos').textContent = `正在比第 ${Math.min(stats.same + differ + 1, total)} / ${total} 页`; return; }
      box.querySelector('.pd-pos').textContent = differ ? `共 ${differ} 页不一样` : '没有不同';
      box.querySelector('.pd-all-same').hidden = differ > 0 || !total;
      const desc = `共 ${total} 页：一样 ${stats.same} 页${stats.diff ? `，有不同 ${stats.diff} 页（${stats.boxes} 处）` : ''}${stats.onlyLeft ? `，左边多 ${stats.onlyLeft} 页` : ''}${stats.onlyRight ? `，右边多 ${stats.onlyRight} 页` : ''}`;
      box.querySelector('.pd-verdict-slot').innerHTML = compact ? `<p class="pd-desc">${esc(desc)}</p>`
        : `<div class="pd-verdict ${differ ? 'is-diff' : 'is-same'}">${differ ? `有 ${differ} 页不一样` : '✓ 两边每一页看起来都一样'}<small>${esc(desc)}</small></div>`;
    };
    paintStats(false);
    for (let i = 0; i < total; i++) {
      if (!compact && gen !== generation) return null;  // 上面拖放的又重新比了
      if (onProgress) onProgress(i + 1, total);
      const L = left[i] || null, R = right[i] || null;
      const pair = await comparePair(L, R);
      if (pair.kind === 'same') stats.same += 1;
      else if (pair.kind === 'diff') { stats.diff += 1; stats.boxes += pair.boxes.length; }
      else if (pair.kind === 'left') stats.onlyLeft += 1;
      else stats.onlyRight += 1;
      pagesBox.insertAdjacentHTML('beforeend', pairHtml(i, L, R, pair));
      paintStats(false);
    }
    paintStats(true);
    log('逐页对比完成', stats);
    return stats;
  }

  // ---------- 鼠标移到红绿框上：弹出左右两边这一块的放大图（用户 2026-11-19） ----------
  // 浮在鼠标旁边（position: fixed），不挪页面；左右用同一块区域，方便对照
  const ZOOM_MAX = 4, ZOOM_BOX_W = 520, ZOOM_BOX_H = 380;
  const zoomPop = document.createElement('div');
  zoomPop.className = 'pd-zoom';
  zoomPop.hidden = true;
  zoomPop.setAttribute('aria-hidden', 'true');
  document.body.append(zoomPop);
  let zoomBox = null;

  function showZoom(boxEl) {
    const pg = boxEl.closest('.pd-pg');
    if (!pg) return;
    const W = +pg.dataset.w, H = +pg.dataset.h;
    const [bx, by, bw, bh] = boxEl.dataset.b.split(',').map(Number);
    // 框外面再留一点，看得到上下文
    const margin = Math.max(bw * W, bh * H) * 0.35 + 12;
    const X = Math.max(bx * W - margin, 0), Y = Math.max(by * H - margin, 0);
    const RW = Math.min(bx * W + bw * W + margin, W) - X, RH = Math.min(by * H + bh * H + margin, H) - Y;
    const k = Math.max(1, Math.min(ZOOM_MAX, ZOOM_BOX_W / RW, ZOOM_BOX_H / RH));
    const side = (frame, cls, label) => {
      const img = frame && frame.querySelector('img');
      const mark = `<span class="pd-zoom-mark" style="left:${(bx * W - X) * k}px;top:${(by * H - Y) * k}px;width:${bw * W * k}px;height:${bh * H * k}px"></span>`;
      const body = img
        ? `<div class="pd-zoom-crop" style="width:${RW * k}px;height:${RH * k}px"><img src="${img.getAttribute('src')}&sharp=1" alt="" style="width:${W * k}px;left:${-X * k}px;top:${-Y * k}px">${mark}</div>`
        : `<div class="pd-zoom-crop is-missing" style="width:${RW * k}px;height:${RH * k}px"><span>没有这一页</span></div>`;
      return `<div class="pd-zoom-side ${cls}"><b>${label}</b>${body}</div>`;
    };
    zoomPop.innerHTML = side(pg.querySelector('.pd-frame.is-old:not(.is-missing)'), 'is-old', '左边')
      + side(pg.querySelector('.pd-frame.is-new:not(.is-missing)'), 'is-new', '右边')
      + `<p class="pd-zoom-tip">放大 ${k.toFixed(1)} 倍</p>`;
    zoomPop.hidden = false;
  }

  function placeZoom(event) {
    const gap = 18, rect = zoomPop.getBoundingClientRect();
    let x = event.clientX + gap, y = event.clientY + gap;
    if (x + rect.width > innerWidth - 8) x = event.clientX - gap - rect.width;
    if (y + rect.height > innerHeight - 8) y = innerHeight - 8 - rect.height;
    zoomPop.style.left = `${Math.max(8, x)}px`;
    zoomPop.style.top = `${Math.max(8, y)}px`;
  }

  document.addEventListener('mouseover', event => {
    const boxEl = event.target.closest && event.target.closest('.pd-box[data-b]');
    if (boxEl === zoomBox) return;
    zoomBox = boxEl;
    if (!boxEl) { zoomPop.hidden = true; return; }
    showZoom(boxEl);
    placeZoom(event);
  });
  document.addEventListener('mousemove', event => { if (zoomBox) placeZoom(event); });
  document.addEventListener('scroll', () => { zoomBox = null; zoomPop.hidden = true; }, { passive: true });

  async function run() {
    if (running) return;
    const left = okFiles('left'), right = okFiles('right');
    if (!left.length || !right.length) return;
    running = true;
    $('pdError').textContent = '';
    const snapshot = { left: left.map(f => ({ ...f })), right: right.map(f => ({ ...f })) };
    const sig = signature();
    result = snapshot;
    resultSig = sig;
    renderSides();
    log('开始对比', snapshot.left.map(f => f.name), 'vs', snapshot.right.map(f => f.name));
    const box = $('pdResult');
    try {
      await compareInto(box, snapshot, {
        onProgress: (i, n) => {
          $('pdStatus').innerHTML = `<span class="pd-busy"><span class="pd-spinner" aria-hidden="true"></span>正在比第 ${i} / ${n} 页…</span>`;
        },
      });
      $('pdStatus').textContent = '';
    } catch (err) {
      $('pdStatus').textContent = '';
      $('pdError').textContent = `对比失败：${err.message || err}`;
      log('对比失败', err);
    } finally {
      running = false;
      renderSides();
    }
  }

  function renderResult() {
    const box = $('pdResult');
    box.classList.remove('pd-view');
    box.innerHTML = '<p class="pd-empty-result">左右两边都放好 PDF 后点「🔍 开始对比」。</p>';
  }

  // 上一处 / 下一处：在不一样的页之间跳（只在点按钮时滚动）
  function navigate(box, step) {
    const state = views.get(box);
    const targets = [...box.querySelectorAll('.pd-pg:not(.is-same)')];
    if (!targets.length) return;
    const i = state.hunkIndex;
    state.hunkIndex = i < 0 ? (step > 0 ? 0 : targets.length - 1) : (i + step + targets.length) % targets.length;
    box.querySelectorAll('.pd-pg.is-current').forEach(el => el.classList.remove('is-current'));
    const target = targets[state.hunkIndex];
    target.classList.add('is-current');
    target.scrollIntoView({ block: 'start' });
    box.querySelector('.pd-pos').textContent = `第 ${state.hunkIndex + 1} / ${targets.length} 页不一样`;
  }

  document.addEventListener('click', event => {
    const nav = event.target.closest('[data-nav]');
    const box = nav && nav.closest('.pd-view');
    if (box && views.get(box)) navigate(box, +nav.dataset.nav);
  });
  document.addEventListener('change', event => {
    if (!event.target.classList.contains('pd-same-toggle')) return;
    const box = event.target.closest('.pd-view');
    if (box) box.classList.toggle('show-same', event.target.checked);
  });

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
        const body = card.querySelector('.pd-gbody');
        const stats = await compareInto(body, files, { compact: true, onProgress: (n, total) => setGroup(card, 'is-busy', `正在比第 ${n} / ${total} 页`) });
        const differ = stats.diff + stats.onlyLeft + stats.onlyRight;
        const pages = side => files[side].reduce((n, f) => n + f.pages, 0);
        const small = `第 ${g.line} 行 · 左 ${pages('left')} 页、右 ${pages('right')} 页`;
        if (!differ) { counts.same += 1; setGroup(card, 'is-same', '✓ 一样', small); }
        else { counts.diff += 1; setGroup(card, 'is-diff', `${differ} 页不一样`, small); }
        log('第', i + 1, '组', stats);
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
