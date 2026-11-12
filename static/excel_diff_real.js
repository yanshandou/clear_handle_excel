// 对比的「真实效果」（用户 2026-11-12）：Excel 真导出来的 PDF 页面，两边逐像素比，不一样的地方左边盖红、右边盖绿。
// 对比页（static/excel_diff.js）比完会发 xd:done；这里拿有不同的表，请后台导 PDF（core/excel_diff_pdf.py），
// 默认标出明显差异，也可显示原始的细微像素差异。
(function () {
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const BLOCK = 6;
  const PIXEL_TOLERANCE = 48;
  const VISIBLE_TOLERANCE = 72;
  let generation = 0;     // 重新对比时，上一轮还没做完的就不要了

  async function post(path, body) {
    const response = await fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const data = await response.json();
    if (!response.ok || data.ok === false) throw new Error(data.error || '请求失败');
    return data;
  }
  const wait = ms => new Promise(resolve => setTimeout(resolve, ms));
  const pageUrl = (key, sheet, p) => `/api/excel_diff/page?key=${encodeURIComponent(key)}&sheet=${encodeURIComponent(sheet)}&p=${p}`;
  function loadImage(url) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('页面图取不到'));
      img.src = url;
    });
  }

  const colorDistance = (a, i, b, j) => Math.abs(a[i] - b[j]) + Math.abs(a[i + 1] - b[j + 1]) + Math.abs(a[i + 2] - b[j + 2]);

  function detectDifferences(a, b, width, height) {
    const cols = Math.ceil(width / BLOCK);
    const blockCount = cols * Math.ceil(height / BLOCK);
    const pixelHits = new Uint8Array(blockCount);
    const visibleHits = new Uint8Array(blockCount);
    const unmatched = new Uint8Array(width * height);
    let pixelCount = 0, visibleCount = 0, hasUnmatched = false;
    for (let y = 0; y < height; y++) {
      const blockRow = Math.floor(y / BLOCK) * cols;
      const top = Math.max(y - 1, 0), bottom = Math.min(y + 1, height - 1);
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4;
        const distance = colorDistance(a, i, b, i);
        if (distance <= PIXEL_TOLERANCE) continue;
        const block = blockRow + Math.floor(x / BLOCK);
        if (!pixelHits[block]) { pixelHits[block] = 1; pixelCount++; }
        if (distance <= VISIBLE_TOLERANCE) continue;

        // 双向在周围 1 像素内找相近颜色，容忍轻微偏移和抗锯齿差异。
        let oldMatches = false, newMatches = false;
        const left = Math.max(x - 1, 0), right = Math.min(x + 1, width - 1);
        for (let ny = top; ny <= bottom && !(oldMatches && newMatches); ny++) {
          const start = (ny * width + left) * 4, end = (ny * width + right) * 4;
          for (let j = start; j <= end; j += 4) {
            if (j === i) continue;
            oldMatches ||= colorDistance(a, i, b, j) <= VISIBLE_TOLERANCE;
            newMatches ||= colorDistance(b, i, a, j) <= VISIBLE_TOLERANCE;
            if (oldMatches && newMatches) break;
          }
        }
        if (!oldMatches || !newMatches) {
          unmatched[i / 4] = 1;
          hasUnmatched = true;
        }
      }
    }
    if (!hasUnmatched) return { cols, pixelHits, pixelCount, visibleHits, visibleCount };
    for (let y = 0; y < height; y++) {
      const blockRow = Math.floor(y / BLOCK) * cols;
      for (let x = 0; x < width; x++) {
        const i = y * width + x;
        if (!unmatched[i]) continue;
        // 孤立噪点只在细微差异模式下显示。相邻证据可以跨过蒙版网格边界。
        const left = x > 0, right = x + 1 < width;
        const hasNeighbor = (left && unmatched[i - 1]) || (right && unmatched[i + 1])
          || (y > 0 && (unmatched[i - width] || (left && unmatched[i - width - 1]) || (right && unmatched[i - width + 1])))
          || (y + 1 < height && (unmatched[i + width] || (left && unmatched[i + width - 1]) || (right && unmatched[i + width + 1])));
        if (!hasNeighbor) continue;
        const block = blockRow + Math.floor(x / BLOCK);
        if (!visibleHits[block]) { visibleHits[block] = 1; visibleCount++; }
      }
    }
    return { cols, pixelHits, pixelCount, visibleHits, visibleCount };
  }

  function compare(oldImg, newImg) {
    const width = Math.max(oldImg?.width || 0, newImg?.width || 0), height = Math.max(oldImg?.height || 0, newImg?.height || 0);
    const pixels = img => {
      const c = document.createElement('canvas');
      c.width = width; c.height = height;
      const g = c.getContext('2d');
      g.fillStyle = '#fff'; g.fillRect(0, 0, width, height);
      if (img) g.drawImage(img, 0, 0);
      return g.getImageData(0, 0, width, height).data;
    };
    const diff = oldImg && newImg ? detectDifferences(pixels(oldImg), pixels(newImg), width, height) : null;
    const side = img => {
      const c = document.createElement('canvas');
      c.width = width || 1; c.height = height || 1;
      const g = c.getContext('2d');
      g.fillStyle = img ? '#fff' : '#f3f4f6';
      g.fillRect(0, 0, c.width, c.height);
      if (img) g.drawImage(img, 0, 0);
      const mask = document.createElement('canvas');
      mask.width = c.width; mask.height = c.height;
      mask.className = 'xr-mask';
      c.className = 'xr-base';
      return { base: c, mask };
    };
    const old = side(oldImg), current = side(newImg);
    const draw = strict => {
      const hits = diff && (strict ? diff.pixelHits : diff.visibleHits);
      for (const [mask, color] of [[old.mask, '#e5303f'], [current.mask, '#1fae4b']]) {
        const g = mask.getContext('2d');
        g.clearRect(0, 0, width, height);
        g.fillStyle = color;
        if (!diff) { g.fillRect(0, 0, width, height); continue; }
        for (let k = 0; k < hits.length; k++) {
          if (!hits[k]) continue;
          const x = (k % diff.cols) * BLOCK, y = Math.floor(k / diff.cols) * BLOCK;
          g.fillRect(x - 2, y - 2, BLOCK + 4, BLOCK + 4);
        }
      }
    };
    return { old, new: current, draw, count: diff ? diff.visibleCount : -1, subtle: !!diff?.pixelCount };
  }

  // 等待中：转圈 + 已经等了多久（用户 2026-11-12：什么提示都没有，看不出是卡住了还是在导）
  const waitingSince = new WeakMap();
  const duration = seconds => seconds < 60 ? `约 ${Math.max(1, Math.round(seconds))} 秒` : `约 ${Math.round(seconds / 60)} 分钟`;
  function busy(box, text, progress = null) {
    if (!waitingSince.has(box)) waitingSince.set(box, Date.now());
    const seconds = Math.round((Date.now() - waitingSince.get(box)) / 1000);
    const percent = progress ? Math.min(99, Math.floor(progress.percent)) : null;
    box.innerHTML = `<div class="xr-wait"><p class="xr-status xr-busy"><span class="xr-spinner" aria-hidden="true"></span>`
      + (percent !== null ? `<b class="xr-percent">${percent}%</b>` : '') + `📄 真实效果：${esc(text)}</p>`
      + (progress ? `<div class="xr-bar" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${percent}"><span style="width:${percent}%"></span></div>` : '')
      + `<small class="xr-elapsed">${progress ? `已导 ${progress.done} / ${progress.total} 张表 · ` : ''}已等 ${seconds} 秒`
      + `${progress && progress.eta > 0 ? ` · 预计还要${duration(progress.eta)}` : ''}</small></div>`;
  }
  // 两边一共要导几张表、导了几张、还要多久（后台给每张表和打开 Excel 的实际平均用时）
  function progressOf(keys, want) {
    let total = 0, done = 0, eta = 0;
    for (const side of ['mine', 'theirs']) {
      const k = keys[side];
      if (!k) continue;
      const names = want[side];
      const have = names.filter(name => k.sheets[name]).length;
      const running = k.state === 'running' ? Math.min(k.done || 0, names.length - have) : 0;
      total += names.length;
      done += have + running;
      if (k.state === 'running' || k.state === 'queued') {
        const own = (names.length - have - running) * k.sheet_seconds + (k.state === 'queued' ? k.open_seconds : 0);
        const before = (k.ahead_sheets || 0) * k.sheet_seconds + (k.ahead_files || 0) * k.open_seconds;
        eta = Math.max(eta, own + before);  // 后导的那边，前面的活已经包括先导的那边
      }
    }
    return total ? { total, done, percent: done / total * 100, eta } : null;
  }

  function block(details) {
    let body = details.querySelector(':scope > .xd-body');
    if (!body) { body = document.createElement('div'); body.className = 'xd-body'; details.appendChild(body); }
    let box = body.querySelector(':scope > .xr-real');
    if (!box) {
      box = document.createElement('div');
      box.className = 'xr-real';
      body.prepend(box);
    }
    return box;
  }

  function paint(box, sheet, info, oldKey, newKey, gen) {
    const o = info.theirs, n = info.mine;
    const notes = [];
    if (o && o.error) notes.push(`远程这份：${o.error}`);
    if (n && n.error) notes.push(`你这份：${n.error}`);
    const oldPages = o && !o.error ? o.pages : 0, newPages = n && !n.error ? n.pages : 0;
    const total = Math.max(oldPages, newPages);
    if (!total) {  // 出错要说清楚：哪一边、什么原因（用户 2026-11-12）
      box.innerHTML = `<p class="${notes.length ? 'xr-warn' : 'xr-status'}">📄 真实效果：没有可以看的页面${notes.length ? `。${esc(notes.join('；'))}` : '（两边这张表都没有可打印的内容）'}</p>`;
      return;
    }
    if (o && n && oldPages !== newPages && oldPages && newPages) notes.push(`页数不一样：远程 ${oldPages} 页，你这份 ${newPages} 页`);
    box.innerHTML = `<div class="xr-head"><strong>📄 真实效果（Excel 导出的 PDF）</strong><span class="xr-status xr-busy"><span class="xr-spinner" aria-hidden="true"></span>正在取页面图、逐页比（共 ${total} 页）…</span>`
      + `<label><input type="checkbox" class="xr-toggle" checked> 显示红绿蒙版</label>`
      + `<label title="查看轻微偏移、抗锯齿和噪点等像素变化"><input type="checkbox" class="xr-strict"> 显示细微像素差异</label>`
      + `<button type="button" class="btn btn-outline btn-sm xr-all" hidden>显示所有页</button></div>`
      + (notes.length ? `<p class="xr-warn">! ${esc(notes.join('；'))}</p>` : '')
      + '<div class="xr-pages"></div>';
    const pages = box.querySelector('.xr-pages');
    const strict = box.querySelector('.xr-strict');
    const masks = [];
    box.querySelector('.xr-toggle').addEventListener('change', event => box.classList.toggle('is-nomask', !event.target.checked));
    strict.addEventListener('change', () => {
      box.classList.toggle('is-strict', strict.checked);
      for (const draw of masks) draw(strict.checked);
    });
    box.querySelector('.xr-all').addEventListener('click', event => {
      const showing = box.classList.toggle('is-allpages');
      event.target.textContent = showing ? '收起无差异的页' : `显示其余 ${pages.querySelectorAll('.xr-page.is-same, .xr-page.is-subtle').length} 页（无明显差异）`;
    });
    (async () => {
      let differ = 0, same = 0, subtle = 0, unavailable = 0, wholeShift = 0;
      for (let p = 0; p < total; p++) {
        if (gen !== generation) return;
        const [oldImg, newImg] = await Promise.all([
          p < oldPages ? loadImage(pageUrl(oldKey, sheet, p)) : null,
          p < newPages ? loadImage(pageUrl(newKey, sheet, p)) : null,
        ]).catch(() => [null, null]);
        if (gen !== generation) return;
        const progress = box.querySelector('.xr-status.xr-busy');
        if (progress) progress.lastChild.textContent = `正在取页面图、逐页比（第 ${p + 1} / ${total} 页）…`;
        if ((p < oldPages && !oldImg) || (p < newPages && !newImg)) { unavailable++; continue; }
        const result = compare(oldImg, newImg);
        if (result.count || result.subtle) {
          result.draw(strict.checked);
          masks.push(result.draw);
        }
        const row = document.createElement('div');
        row.className = 'xr-page' + (result.count === 0 ? ' is-same' : '');
        if (result.count === 0 && result.subtle) row.classList.add('is-subtle');
        const cell = (side, label, cls) => {
          const wrap = document.createElement('div');
          wrap.className = 'xr-side ' + cls;
          wrap.innerHTML = `<small>${label}</small>`;
          const frame = document.createElement('div');
          frame.className = 'xr-frame';
          frame.append(side.base, side.mask);
          wrap.appendChild(frame);
          return wrap;
        };
        const title = document.createElement('div');
        title.className = 'xr-page-title';
        title.textContent = `第 ${p + 1} 页` + (result.count === 0 ? (result.subtle ? ' · 仅有细微像素差异' : ' · 未发现差异')
          : result.count < 0 ? (oldImg ? ' · 你这份没有这一页' : ' · 远程没有这一页') : '');
        row.append(title, cell(result.old, oldImg ? '− 远程（旧）' : '− 远程没有这一页', 'is-old'), cell(result.new, newImg ? '+ 你这份（新）' : '+ 你这份没有这一页', 'is-new'));
        pages.appendChild(row);
        if (result.count !== 0) differ++;
        else if (result.subtle) subtle++;
        else same++;
        if (result.count > 0 && result.count > 0.4 * Math.ceil(result.old.base.width / BLOCK) * Math.ceil(result.old.base.height / BLOCK)) wholeShift++;
      }
      const status = box.querySelector('.xr-status');
      status.classList.remove('xr-busy');
      status.querySelector('.xr-spinner')?.remove();
      status.textContent = [
        differ ? `${differ} 页有明显差异` : total > unavailable ? '未发现明显的打印差异' : '',
        subtle ? `${subtle} 页仅有细微像素差异` : '',
        unavailable ? `${unavailable} 页预览加载失败` : '',
      ].filter(Boolean).join('，') + '。';
      if (wholeShift) status.textContent += ' 大片标色可能来自版式移动，具体改动可查下面的差异清单。';
      const all = box.querySelector('.xr-all');
      all.hidden = !(same || subtle);
      if (!differ && total > unavailable) {
        // 一页都没有看得出的差异：页面直接摆出来给人核对，不收起（用户 2026-11-12：看不到页面，分不清是没差异还是没做出来）
        box.classList.add('is-allpages');
        all.textContent = '收起无差异的页';
        status.textContent = `已导出并逐页比完：印出来的 ${total - unavailable} 页都没有看得出的差异。`
          + '上面逐格比出的不同多半没印出来（在打印区域外、在隐藏的行列里，或改动太小）。两边的页面都在下面，可以自己核对。'
          + (unavailable ? ` ${unavailable} 页预览加载失败。` : '');
      } else if (same || subtle) {
        all.textContent = `显示其余 ${same + subtle} 页（无明显差异）`;
      }
      if (unavailable) status.classList.add('xr-warn');
      console.info('📄 真实效果对比', { 表: sheet, 有明显差异的页: differ, 仅有细微差异的页: subtle, 未发现差异的页: same, 预览加载失败的页: unavailable });
    })();
  }

  window.addEventListener('xd:done', async event => {
    const data = event.detail || {};
    const gen = ++generation;
    if (data.identical || !Array.isArray(data.sheets)) return;
    const notify = state => {
      if (gen === generation) window.dispatchEvent(new CustomEvent('xd:preview', { detail: state }));
    };
    const want = { mine: [], theirs: [] };
    const boxes = new Map();
    for (const sheet of data.sheets) {
      if (sheet.doomed || sheet.status === 'same') continue;
      if (sheet.status !== 'only_new') want.theirs.push(sheet.name);
      if (sheet.status !== 'only_old') want.mine.push(sheet.name);
      const details = [...document.querySelectorAll('details.xd-sheet[data-sheet]')].find(d => d.dataset.sheet === sheet.name);
      if (details) {
        const box = block(details);
        busy(box, '正在请后台 Excel 导出 PDF…');
        boxes.set(sheet.name, box);
      }
    }
    if (!boxes.size) { notify('skipped'); return; }
    let keys;
    try {
      const started = await post('/api/excel_diff/pdf', { mine: data.mine, theirs: data.theirs, sheets: want });
      keys = { mine: started.mine, theirs: started.theirs };
      notify('queued');
    } catch (error) {
      for (const box of boxes.values()) box.innerHTML = `<p class="xr-warn">📄 真实效果做不了：${esc(error.message)}</p>`;
      notify('failed');
      return;
    }
    const painted = new Set();
    let hasPages = false, exportFailed = false, pollFailures = 0;
    for (;;) {
      if (gen !== generation) return;
      for (const [name, box] of boxes) {
        if (painted.has(name)) continue;
        const info = {};
        let ready = true;
        for (const side of ['mine', 'theirs']) {
          if (!want[side].includes(name)) continue;
          const k = keys[side];
          if (k.sheets[name]) info[side] = k.sheets[name];
          else if (k.state === 'error' || k.state === 'done') {
            exportFailed = true;
            info[side] = { pages: 0, error: k.state === 'error' ? '导出 PDF 失败：' + (k.error || '') : '这张表的 PDF 缺失，请重新对比' };
          } else ready = false;
        }
        if (ready) {
          painted.add(name);
          hasPages ||= !!(info.theirs?.pages || info.mine?.pages);
          paint(box, name, info, keys.theirs?.key, keys.mine?.key, gen);
        } else {
          const k = [keys.theirs, keys.mine].find(item => item && item.state !== 'done');
          busy(box, k && k.state === 'running' ? `后台 Excel 正在导出 PDF（${k.current || '准备中'}）…` : `排队等后台 Excel 导出 PDF${k && k.queue > 1 ? `（前面还有 ${k.queue - 1} 份）` : ''}…`, progressOf(keys, want));
        }
      }
      if (painted.size === boxes.size) {
        notify(exportFailed ? 'failed' : hasPages ? 'done' : 'skipped');
        return;
      }
      await wait(1500);
      try {
        const query = {};
        for (const side of ['mine', 'theirs']) {
          if (!keys[side]) continue;
          const { key } = keys[side];
          query[key] = [...new Set([...(query[key] || []), ...want[side]])];
        }
        const { status } = await post('/api/excel_diff/pdf_status', { keys: query });
        for (const side of ['mine', 'theirs']) if (keys[side]) keys[side] = status[keys[side].key] || keys[side];
        pollFailures = 0;
      } catch (error) {
        console.warn('📄 真实效果：问进度失败', error.message);
        if (++pollFailures >= 3) {
          for (const [name, box] of boxes) {
            if (!painted.has(name)) box.innerHTML = `<p class="xr-warn">📄 取不到 PDF 导出进度：${esc(error.message)}</p>`;
          }
          notify('failed');
          return;
        }
      }
    }
  });
})();
