// 对比的「真实效果」（用户 2026-11-12）：Excel 真导出来的 PDF 页面，两边逐像素比，不一样的地方左边盖红、右边盖绿。
// 对比页（static/excel_diff.js）比完会发 xd:done；这里拿有不同的表，请后台导 PDF（core/excel_diff_pdf.py），
// 导好后取页面图，在浏览器里逐像素比、画蒙版。完全一样的页默认收起来。
(function () {
  const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const BLOCK = 6;        // 每 6×6 像素算一格，有一个像素不一样就整格盖色
  const TOLERANCE = 48;   // 三个颜色通道差的和超过它才算不一样（避开抗锯齿的细微差别）
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

  // 两张页面图比一比：返回两张画好蒙版的画布和不一样的格数。一边没有这一页时整页盖色
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
    const a = pixels(oldImg), b = pixels(newImg);
    const cols = Math.ceil(width / BLOCK), rows = Math.ceil(height / BLOCK);
    const hit = new Uint8Array(cols * rows);
    let count = 0;
    for (let y = 0; y < height; y++) {
      const by = (y / BLOCK) | 0;
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4;
        if (Math.abs(a[i] - b[i]) + Math.abs(a[i + 1] - b[i + 1]) + Math.abs(a[i + 2] - b[i + 2]) > TOLERANCE) {
          const k = by * cols + ((x / BLOCK) | 0);
          if (!hit[k]) { hit[k] = 1; count++; }
          x = (((x / BLOCK) | 0) + 1) * BLOCK - 1;  // 这一格已经算不一样了，跳到下一格
        }
      }
    }
    const side = (img, color) => {
      const c = document.createElement('canvas');
      c.width = width || 1; c.height = height || 1;
      const g = c.getContext('2d');
      g.fillStyle = img ? '#fff' : '#f3f4f6';
      g.fillRect(0, 0, c.width, c.height);
      if (img) g.drawImage(img, 0, 0);
      const mask = document.createElement('canvas');
      mask.width = c.width; mask.height = c.height;
      const m = mask.getContext('2d');
      m.fillStyle = color;
      if (!img) m.fillRect(0, 0, c.width, c.height);
      else for (let k = 0; k < hit.length; k++) {
        if (!hit[k]) continue;
        const x = (k % cols) * BLOCK, y = ((k / cols) | 0) * BLOCK;
        m.fillRect(x - 2 * BLOCK, y - 2 * BLOCK, BLOCK * 5, BLOCK * 5);  // 往外扩两格，字的笔画连成一片，细小的改动也看得见
      }
      mask.className = 'xr-mask';
      c.className = 'xr-base';
      return { base: c, mask };
    };
    return { old: side(oldImg, '#e5303f'), new: side(newImg, '#1fae4b'), count: oldImg && newImg ? count : -1 };
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
    if (!total) { box.innerHTML = `<p class="xr-status">📄 真实效果：没有可以看的页面。${esc(notes.join('；'))}</p>`; return; }
    if (o && n && oldPages !== newPages && oldPages && newPages) notes.push(`页数不一样：远程 ${oldPages} 页，你这份 ${newPages} 页`);
    box.innerHTML = `<div class="xr-head"><strong>📄 真实效果（Excel 导出的 PDF）</strong><span class="xr-status">正在逐页比…</span>`
      + `<label><input type="checkbox" class="xr-toggle" checked> 显示红绿蒙版</label>`
      + `<button type="button" class="btn btn-outline btn-sm xr-all" hidden>显示完全一样的页</button></div>`
      + (notes.length ? `<p class="xr-warn">! ${esc(notes.join('；'))}</p>` : '')
      + '<div class="xr-pages"></div>';
    const pages = box.querySelector('.xr-pages');
    box.querySelector('.xr-toggle').addEventListener('change', event => box.classList.toggle('is-nomask', !event.target.checked));
    box.querySelector('.xr-all').addEventListener('click', event => {
      const showing = box.classList.toggle('is-allpages');
      event.target.textContent = showing ? '只看有不同的页' : '显示完全一样的页';
    });
    (async () => {
      let differ = 0, same = 0, wholeShift = 0;
      for (let p = 0; p < total; p++) {
        if (gen !== generation) return;
        const [oldImg, newImg] = await Promise.all([
          p < oldPages ? loadImage(pageUrl(oldKey, sheet, p)) : null,
          p < newPages ? loadImage(pageUrl(newKey, sheet, p)) : null,
        ]).catch(() => [null, null]);
        if (gen !== generation) return;
        if (!oldImg && !newImg) continue;
        const result = compare(oldImg, newImg);
        const row = document.createElement('div');
        row.className = 'xr-page' + (result.count === 0 ? ' is-same' : '');
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
        title.textContent = `第 ${p + 1} 页` + (result.count === 0 ? '：完全一样' : result.count < 0 ? (oldImg ? '：你这份没有这一页' : '：远程没有这一页') : '');
        row.append(title, cell(result.old, oldImg ? '− 远程（旧）' : '− 远程没有这一页', 'is-old'), cell(result.new, newImg ? '+ 你这份（新）' : '+ 你这份没有这一页', 'is-new'));
        pages.appendChild(row);
        if (result.count === 0) same++; else differ++;
        if (result.count > 0 && result.count > 0.4 * Math.ceil(result.old.base.width / BLOCK) * Math.ceil(result.old.base.height / BLOCK)) wholeShift++;
      }
      const status = box.querySelector('.xr-status');
      status.textContent = differ ? `${differ} 页有不同${same ? `，${same} 页完全一样（收起来了）` : ''}。` : '每一页印出来都一样（不同之处印不出来，比如只改了不打印的地方）。';
      if (wholeShift) status.textContent += ' 有的页大片都是红绿：多半是上面的行高 / 列宽变了，把下面挤动了，具体改了什么看下面的差异清单。';
      box.querySelector('.xr-all').hidden = !same;
      console.info('📄 真实效果对比', { 表: sheet, 有不同的页: differ, 一样的页: same });
    })();
  }

  window.addEventListener('xd:done', async event => {
    const data = event.detail || {};
    const gen = ++generation;
    if (data.identical || !Array.isArray(data.sheets)) return;
    const want = { mine: [], theirs: [] };
    const boxes = new Map();
    for (const sheet of data.sheets) {
      if (sheet.doomed || sheet.status === 'same') continue;
      if (sheet.status !== 'only_new') want.theirs.push(sheet.name);
      if (sheet.status !== 'only_old') want.mine.push(sheet.name);
      const details = [...document.querySelectorAll('details.xd-sheet[data-sheet]')].find(d => d.dataset.sheet === sheet.name);
      if (details) {
        const box = block(details);
        box.innerHTML = '<p class="xr-status">📄 真实效果：正在请后台 Excel 导出 PDF…</p>';
        boxes.set(sheet.name, box);
      }
    }
    if (!boxes.size) return;
    let keys;
    try {
      const started = await post('/api/excel_diff/pdf', { mine: data.mine, theirs: data.theirs, sheets: want });
      keys = { mine: started.mine, theirs: started.theirs };
    } catch (error) {
      for (const box of boxes.values()) box.innerHTML = `<p class="xr-warn">📄 真实效果做不了：${esc(error.message)}</p>`;
      return;
    }
    const painted = new Set();
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
          else if (k.state === 'error') info[side] = { pages: 0, error: '导出 PDF 失败：' + (k.error || '') };
          else if (k.state === 'done') info[side] = { pages: 0, error: '这张表的 PDF 缺失，请重新对比' };
          else ready = false;
        }
        if (ready) { painted.add(name); paint(box, name, info, keys.theirs?.key, keys.mine?.key, gen); }
        else {
          const k = [keys.theirs, keys.mine].find(item => item && item.state !== 'done');
          box.innerHTML = `<p class="xr-status">📄 真实效果：${k && k.state === 'running' ? `后台 Excel 正在导出 PDF（${esc(k.current || '准备中')}）…` : `排队等后台 Excel 导出 PDF${k && k.queue ? `（前面还有 ${k.queue - 1} 份）` : ''}…`}</p>`;
        }
      }
      if (painted.size === boxes.size) return;
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
      } catch (error) {
        console.warn('📄 真实效果：问进度失败', error.message);
      }
    }
  });
})();
