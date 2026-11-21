// 设置图标（用户 2026-11-20 要）：拖进透明底 PNG → 生成 Windows 图标（.ico）存进图标库；
// 图标库可以搜索、分类、改名、替换、删除；选中的图标可以套到 Windows 文件夹，或换掉主页工具箱里某个工具的图标。
// 后端 core/icon_library.py；调试日志都以「[设置图标]」开头打在浏览器控制台。
// 页面规矩：点图标只换颜色、右边详情框固定不挪，不自动滚动。
(() => {
  const $ = id => document.getElementById(id);
  const esc = v => String(v ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
  const log = (...args) => console.info('[设置图标]', ...args);
  const KEY = 'clear_handle_excel.iconLibrary';

  let lib = { icons: [], categories: [], tools: {}, windows: true };
  let selected = '';
  let category = null;     // null = 全部，'' = 未分类
  let folderDraft = '';    // 套到文件夹的路径框：换选别的图标也留着
  let lastResults = null;  // { id, html }
  let busy = false;
  let batchMode = false;
  let duplicatesOnly = false;
  let assignRows = [];   // 批量指定：{path, icon, state, message}
  let pickerRow = null;  // 正在挑图标的那一行
  const checked = new Set();

  const prefs = (() => { try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (_) { return {}; } })();
  const savePrefs = () => { try { localStorage.setItem(KEY, JSON.stringify(prefs)); } catch (_) {} };
  $('icNewCategory').value = prefs.newCategory || '';
  $('icTrim').checked = prefs.trim !== false;
  folderDraft = prefs.folders || '';
  assignRows = (Array.isArray(prefs.assignRows) ? prefs.assignRows : [])
    .filter(r => r && typeof r.path === 'string' && r.path.trim())
    .map(r => ({path: r.path, icon: typeof r.icon === 'string' ? r.icon : '', state: '', message: ''}));
  $('ibParent').value = prefs.ibParent || '';
  $('ibPaths').value = prefs.ibPaths || '';

  const pngUrl = icon => `/api/icons/file?id=${icon.id}&kind=png&v=${icon.version}`;
  const icoUrl = icon => `/api/icons/file?id=${icon.id}&kind=ico`;
  const catLabel = c => c || '未分类';
  const byId = id => lib.icons.find(i => i.id === id);

  async function post(url, body) {
    const response = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const result = await response.json().catch(() => ({}));
    if (!response.ok || !result.ok) throw Error(result.error || '请求失败');
    return result;
  }

  function status(html) { $('icStatus').innerHTML = html; }

  async function reload(reason) {
    const response = await fetch('/api/icons', { cache: 'no-store' });
    const data = await response.json();
    if (!data.ok) throw Error(data.error || '读不了图标库');
    lib = data;
    for (const id of checked) if (!byId(id)) checked.delete(id);
    if (selected && !byId(selected)) selected = '';
    for (const row of assignRows) if (row.icon && !byId(row.icon)) { row.icon = ''; row.state = ''; row.message = ''; }
    if (category !== null && !lib.categories.some(c => c.name === category)) category = null;
    log('图标库', reason, `${lib.icons.length} 个图标`, lib.dir);
    renderCats();
    renderGrid();
    renderBatch();
    renderDetail();
    renderAssign();
  }

  // ---------- 分类 + 搜索 ----------
  function renderCats() {
    const chips = [`<button type="button" class="ic-cat${category === null ? ' is-on' : ''}" data-cat="*">全部<small>${lib.icons.length}</small></button>`]
      .concat(lib.categories.map(c => `<button type="button" class="ic-cat${category === c.name ? ' is-on' : ''}" data-cat="${esc(c.name)}">${esc(catLabel(c.name))}<small>${c.count}</small></button>`));
    $('icCats').innerHTML = chips.join('');
    $('icCategoryList').innerHTML = lib.categories.filter(c => c.name).map(c => `<option value="${esc(c.name)}">`).join('');
    const canRename = category !== null && category !== '';
    $('icCatRename').disabled = $('icCatRenameBtn').disabled = !canRename;
    $('icCatRename').value = canRename ? category : '';
    $('icCatRename').placeholder = canRename ? '新的分类名' : '先点一个分类';
    $('icCount').textContent = `共 ${lib.icons.length} 个 · 存在 ${lib.dir || ''}`;
    const groups = lib.duplicates || [];
    $('icDuplicates').setAttribute('aria-pressed', String(duplicatesOnly));
    $('icDuplicates').textContent = duplicatesOnly ? '显示全部图标' : '查看重复图标';
    $('icDuplicateInfo').textContent = `${groups.length ? `发现 ${groups.length} 组重复` : '没有发现重复'}${lib.duplicate_unchecked?.length ? `；${lib.duplicate_unchecked.length} 个原图无法检查` : ''}`;
  }

  function visible() {
    const q = $('icSearch').value.trim().toLowerCase();
    const duplicateIds = new Set((lib.duplicates || []).flat());
    return lib.icons.filter(i => (!duplicatesOnly || duplicateIds.has(i.id)) && (category === null || i.category === category)
      && (!q || [i.name, i.category, i.source].some(v => String(v || '').toLowerCase().includes(q))));
  }

  function renderGrid() {
    const list = visible();
    if (!list.length) {
      $('icGrid').innerHTML = `<p class="ic-empty">${lib.icons.length ? '没有符合条件的图标' : '图标库还是空的：把透明底 PNG 拖到上面就能加'}</p>`;
      return;
    }
    $('icGrid').innerHTML = list.map(i => {
      const used = (i.folders || []).length + Object.values(lib.tools || {}).filter(v => v === i.id).length;
      const duplicateGroup = (lib.duplicates || []).findIndex(group => group.includes(i.id));
      return `<button type="button" class="ic-card${batchMode ? (checked.has(i.id) ? ' is-checked' : '') : (i.id === selected ? ' is-selected' : '')}" data-id="${i.id}" ${batchMode ? `aria-pressed="${checked.has(i.id)}"` : ''} title="${esc(i.name)}（${esc(catLabel(i.category))}）">
        ${batchMode ? `<span class="ic-check" aria-hidden="true">${checked.has(i.id) ? '☑ 已选' : '☐ 选择'}</span>` : ''}
        <span class="ic-thumb ic-checker"><img src="${pngUrl(i)}" alt="" loading="lazy" width="64" height="64"></span>
        <span class="ic-name">${esc(i.name)}</span>
        <span class="ic-sub">${esc(catLabel(i.category))}</span>
        ${duplicateGroup >= 0 ? `<span class="ic-badge">重复组 ${duplicateGroup + 1}</span>` : ''}
        ${used ? `<span class="ic-badge">用在 ${used} 处</span>` : ''}
      </button>`;
    }).join('');
  }

  function renderBatch() {
    $('icBatch').hidden = !batchMode;
    $('icBatchToggle').setAttribute('aria-pressed', String(batchMode));
    $('icBatchToggle').textContent = batchMode ? '完成批量管理' : '☑ 批量管理';
    $('icBatchCount').textContent = `已选 ${checked.size} 个（跨分类保留）`;
    $('icBatchMove').disabled = $('icBatchDelete').disabled = busy || !checked.size;
    $('icBatchAll').disabled = $('icBatchClear').disabled = busy;
    $('icBatchToggle').disabled = busy;
  }

  function setBusy(value) {
    busy = value;
    renderBatch();
    renderAssign();
  }

  async function batchAction(remove) {
    if (busy || !checked.size) return;
    const icons = [...checked].map(byId).filter(Boolean);
    const target = $('icBatchCategory').value.trim();
    if (remove) {
      const ids = new Set(icons.map(i => i.id));
      const folders = icons.reduce((n, i) => n + (i.folders || []).length, 0);
      const tools = Object.values(lib.tools || {}).filter(id => ids.has(id)).length;
      const ok = await window.appConfirm(`删除选中的 ${icons.length} 个图标？`
        + (folders ? `\n已套用图标的 ${folders} 处文件夹不受影响。` : '')
        + (tools ? `\n使用这些图标的 ${tools} 个工具会恢复默认图标。` : ''), { okText: '删除', tone: 'danger' });
      if (!ok || busy) return;
    }
    setBusy(true);
    const failed = [];
    let done = 0;
    try {
      for (const icon of icons) {
        $('icBatchStatus').textContent = `正在${remove ? '删除' : '移动'} ${done + failed.length + 1} / ${icons.length}：${icon.name}…`;
        try {
          await post(remove ? '/api/icons/delete' : '/api/icons/update', remove ? { id: icon.id } : { id: icon.id, category: target });
          checked.delete(icon.id);
          done++;
        } catch (e) { failed.push(`${icon.name}：${e.message}`); }
      }
      await reload('批量操作后');
      $('icBatchStatus').textContent = `${remove ? '已删除' : `已移入「${catLabel(target)}」`} ${done} 个${failed.length ? `；${failed.length} 个失败：${failed.join('；')}` : ''}`;
    } catch (e) {
      $('icBatchStatus').textContent = `已完成 ${done} 个，刷新图标库失败：${e.message}${failed.length ? `；失败：${failed.join('；')}` : ''}`;
    } finally { setBusy(false); }
  }

  function select(id) {
    if (selected === id) return;
    selected = id;
    // 只换颜色，不重画网格，页面不动
    document.querySelectorAll('.ic-card').forEach(card => card.classList.toggle('is-selected', card.dataset.id === id));
    renderDetail();
    log('选中', byId(id)?.name);
  }

  // ---------- 右边详情 ----------
  function toolOptions(icon) {
    const T = window.Tools;
    if (!T) return '';
    return T.TOOLS.map(t => {
      const cur = lib.tools[t.id];
      const mark = cur === icon.id ? '（正在用这个）' : cur ? '（已换过图标）' : '';
      return `<option value="${t.id}">${esc(t.title)}${mark}</option>`;
    }).join('');
  }

  function renderDetail() {
    const box = $('icDetail');
    const icon = byId(selected);
    if (!icon) {
      box.innerHTML = '<p class="ic-detail-empty">点左边一个图标，在这里改名、分类、替换、删除，或套到文件夹。</p>';
      return;
    }
    const warn = [];
    if (icon.transparent === false) warn.push('<span class="is-warn">这张图没有透明背景，图标会带底色</span>');
    const { render_width: renderWidth, render_height: renderHeight } = icon;
    if (Math.max(renderWidth || icon.width || 0, renderHeight || icon.height || 0) < 128) {
      warn.push('<span class="is-warn">用于生成图标的图片小于 128 像素，大图标可能模糊</span>');
    }
    const cropped = renderWidth && renderHeight && (renderWidth !== icon.width || renderHeight !== icon.height);
    const cropDetails = cropped ? `<br>裁剪后 ${renderWidth}×${renderHeight}` : '';
    const squareDetails = renderWidth && renderHeight && renderWidth !== renderHeight
      ? `<br>方形图标${renderWidth > renderHeight ? '上下' : '左右'}留透明边`
      : '';
    const usedTools = Object.entries(lib.tools || {}).filter(([, v]) => v === icon.id).map(([k]) => window.Tools?.byId[k]?.title || k);
    const folders = icon.folders || [];
    box.innerHTML = `
      <div class="ic-preview">
        <div class="ic-big ic-checker" id="icBig" title="把新 PNG 拖到这里 = 替换这个图标的图片"><img src="${pngUrl(icon)}" alt="${esc(icon.name)}"></div>
        <div>
          <div class="ic-sizes">${[48, 32, 24, 16].map(s => `<figure><span class="ic-s ic-checker"><img src="${pngUrl(icon)}" alt="" width="${s}" height="${s}"></span>${s}</figure>`).join('')}</div>
          <p class="ic-meta">原图 ${icon.width}×${icon.height}${icon.source ? ` · ${esc(icon.source)}` : ''}${cropDetails}${squareDetails}<br>加入 ${esc(icon.created || '')}${warn.length ? '<br>' + warn.join('<br>') : ''}</p>
        </div>
      </div>
      <div class="ic-fields">
        <label>名字<input id="icName" type="text" value="${esc(icon.name)}" maxlength="60" autocomplete="off"></label>
        <label>分类<input id="icCategory" type="text" value="${esc(icon.category)}" list="icCategoryList" maxlength="30" placeholder="未分类" autocomplete="off"></label>
      </div>
      <div class="ic-row">
        <button id="icReplace" class="btn" type="button">🔁 替换图片</button>
        <button id="icRetrim" class="btn" type="button">✂ 重新裁剪</button>
        <a class="btn" href="${icoUrl(icon)}" download>⬇ 下载 .ico</a>
        <button id="icDelete" class="btn btn-danger" type="button">🗑 删除</button>
      </div>
      <div class="ic-sec">
        <h3>📁 套到 Windows 文件夹</h3>
        <textarea id="icFolders" spellcheck="false" placeholder="每行一个文件夹路径，可以直接粘贴「复制为路径」带引号的">${esc(folderDraft)}</textarea>
        <div class="ic-row">
          <button id="icApply" class="btn btn-primary" type="button">📁 套用这个图标</button>
          <button id="icRestore" class="btn" type="button">↩ 恢复默认图标</button>
        </div>
        ${lib.windows ? '' : '<p class="ic-hint">现在不是在 Windows 上运行：会写 desktop.ini，但看不到图标效果。</p>'}
        <ul id="icResults" class="ic-results">${lastResults && lastResults.id === icon.id ? lastResults.html : ''}</ul>
        <p class="ic-hint">已经套了这个图标的文件夹（${folders.length} 个）</p>
        <ul class="ic-folders">${folders.map(f => `<li><span>${esc(f)}</span><button type="button" data-restore="${esc(f)}" title="这个文件夹恢复默认图标">↩</button></li>`).join('') || '<li><span>还没有</span></li>'}</ul>
      </div>
      <div class="ic-sec">
        <h3>🧰 工具箱图标</h3>
        <div class="ic-row"><select id="icTool" aria-label="工具箱里的工具">${toolOptions(icon)}</select></div>
        <div class="ic-row">
          <button id="icToolSet" class="btn" type="button">换成这个图标</button>
          <button id="icToolReset" class="btn" type="button">恢复默认</button>
        </div>
        <p class="ic-hint">${usedTools.length ? `正在用在：${usedTools.map(esc).join('、')}` : '主页工具箱里还没有工具用这个图标'}</p>
      </div>`;
    if (prefs.tool && window.Tools?.byId[prefs.tool]) $('icTool').value = prefs.tool;
    bindDetail(icon);
  }

  function bindDetail(icon) {
    const saveField = async () => {
      const name = $('icName').value.trim(), cat = $('icCategory').value.trim();
      if (name === icon.name && cat === icon.category) return;
      try {
        await post('/api/icons/update', { id: icon.id, name, category: cat });
        status(`<span class="is-ok">✓ 已保存「${esc(name || icon.name)}」</span>`);
        await reload('改名 / 分类');
      } catch (e) { status(`<span class="is-bad">! ${esc(e.message)}</span>`); }
    };
    $('icName').addEventListener('change', saveField);
    $('icCategory').addEventListener('change', saveField);
    $('icFolders').addEventListener('input', e => { folderDraft = e.target.value; prefs.folders = folderDraft; savePrefs(); });
    $('icReplace').addEventListener('click', () => $('icReplaceFile').click());
    $('icRetrim').addEventListener('click', () => retrimIcon(icon));
    $('icDelete').addEventListener('click', () => removeIcon(icon));
    $('icApply').addEventListener('click', () => applyFolders(icon, false));
    $('icRestore').addEventListener('click', () => applyFolders(icon, true));
    $('icTool').addEventListener('change', e => { prefs.tool = e.target.value; savePrefs(); });
    $('icToolSet').addEventListener('click', () => setTool(icon, true));
    $('icToolReset').addEventListener('click', () => setTool(icon, false));
    document.querySelectorAll('[data-restore]').forEach(btn => btn.addEventListener('click', () => restoreOne(btn.dataset.restore)));
    const big = $('icBig');
    big.addEventListener('dragover', e => { if (hasFiles(e)) { e.preventDefault(); e.stopPropagation(); big.classList.add('is-over'); } });
    big.addEventListener('dragleave', () => big.classList.remove('is-over'));
    big.addEventListener('drop', e => {
      e.preventDefault(); e.stopPropagation(); big.classList.remove('is-over');
      const file = [...(e.dataTransfer?.files || [])][0];
      if (file) replaceImage(icon, file);
    });
  }

  // ---------- 上传 / 替换 ----------
  async function upload(file, replaceId) {
    const headers = { 'X-File-Name': encodeURIComponent(file.name), 'X-Icon-Trim': $('icTrim').checked ? '1' : '0' };
    if (replaceId) headers['X-Icon-Replace'] = replaceId;
    else headers['X-Icon-Category'] = encodeURIComponent($('icNewCategory').value.trim());
    const response = await fetch('/api/icons/upload', { method: 'POST', headers, body: file });
    const result = await response.json().catch(() => ({}));
    if (!response.ok || !result.ok) throw Object.assign(Error(result.error || '上传失败'), { code: result.code });
    return result;
  }

  function infoWarn(info) {
    const out = [];
    if (!info.transparent) out.push('没有透明背景');
    if (info.small) out.push('用于生成图标的图片小于 128 像素，大图标可能模糊');
    return out;
  }

  async function addFiles(files) {
    files = [...files].filter(f => f && f.size);
    if (!files.length || busy) return;
    setBusy(true);
    prefs.newCategory = $('icNewCategory').value.trim(); prefs.trim = $('icTrim').checked; savePrefs();
    const done = [], skipped = [], bad = [], warns = [];
    try {
      for (const [n, file] of files.entries()) {
        status(`正在生成图标 ${n + 1} / ${files.length}：${esc(file.name)}…`);
        try {
          const r = await upload(file);
          done.push(r.icon);
          const w = infoWarn(r.info);
          if (w.length) warns.push(`${file.name}：${w.join('、')}`);
          log('加进图标库', file.name, r.info);
        } catch (e) {
          if (e.code === 'duplicate_icon') {
            skipped.push(`${file.name}：${e.message}`);
            continue;
          }
          bad.push(`${file.name}：${e.message}`);
          console.warn('[设置图标] 加不进去', file.name, e);
        }
      }
      await reload('加图标后');
      if (done.length === 1) select(done[0].id);
      const messages = [];
      if (done.length) messages.push(`<span class="is-ok">✓ 加进图标库 ${done.length} 个</span>`);
      if (skipped.length) messages.push(`<span class="is-warn">已跳过 ${skipped.length} 张重复图片</span>`);
      messages.push(...skipped.concat(warns).map(w => `<span class="is-warn">${esc(w)}</span>`));
      messages.push(...bad.map(b => `<span class="is-bad">! ${esc(b)}</span>`));
      status(messages.join('<br>'));
    } catch (e) {
      status(`<span class="is-bad">! ${esc(e.message)}</span>`);
    } finally { setBusy(false); }
  }

  async function replaceImage(icon, file) {
    if (busy) return;
    const n = (icon.folders || []).length;
    const ok = await window.appConfirm(`用「${file.name}」替换「${icon.name}」的图片？${n ? `\n已经套了这个图标的 ${n} 个文件夹会一起换成新图。` : ''}`, { okText: '替换' });
    if (!ok || busy) return;
    setBusy(true);
    status(`正在替换「${esc(icon.name)}」…`);
    try {
      const r = await upload(file, icon.id);
      const failed = (r.folders || []).filter(f => !f.ok);
      await reload('替换后');
      const w = infoWarn(r.info);
      status(`<span class="is-ok">✓ 已替换「${esc(icon.name)}」的图片${(r.folders || []).length ? `，文件夹一起换了 ${(r.folders.length - failed.length)} 个` : ''}</span>`
        + w.map(x => `<br><span class="is-warn">! ${esc(x)}</span>`).join('')
        + failed.map(f => `<br><span class="is-bad">! ${esc(f.folder)}：${esc(f.error)}</span>`).join(''));
      log('替换', icon.name, r);
    } catch (e) {
      status(`<span class="is-bad">! ${esc(e.message)}</span>`);
    } finally { setBusy(false); }
  }

  async function removeIcon(icon) {
    const n = (icon.folders || []).length;
    const tools = Object.values(lib.tools || {}).filter(v => v === icon.id).length;
    const ok = await window.appConfirm(`删除图标「${icon.name}」？`
      + (n ? `\n已经套了它的 ${n} 个文件夹不受影响（文件夹里有自己的一份）。` : '')
      + (tools ? `\n工具箱里用它的 ${tools} 个工具会恢复默认图标。` : ''), { okText: '删除', tone: 'danger' });
    if (!ok) return;
    try {
      await post('/api/icons/delete', { id: icon.id });
      status(`<span class="is-ok">✓ 已删除「${esc(icon.name)}」</span>`);
      selected = '';
      await reload('删除后');
    } catch (e) { status(`<span class="is-bad">! ${esc(e.message)}</span>`); }
  }

  async function retrimIcon(icon) {
    if (busy) return;
    const n = (icon.folders || []).length;
    const ok = await window.appConfirm(`用保存的原图重新裁剪「${icon.name}」？${n ? `\n已套用的 ${n} 个文件夹也会更新。` : ''}`, { okText: '重新裁剪' });
    if (!ok || busy) return;
    setBusy(true);
    try {
      const r = await post('/api/icons/retrim', { id: icon.id });
      const failed = (r.folders || []).filter(f => !f.ok);
      await reload('重新裁剪后');
      const { trimmed, render_width: width, render_height: height } = r.info;
      const message = trimmed ? `「${esc(icon.name)}」裁剪后为 ${width}×${height}` : `「${esc(icon.name)}」没有可裁剪的透明边`;
      status(`<span class="is-ok">✓ ${message}</span>`
        + failed.map(f => `<br><span class="is-bad">! ${esc(f.folder)}：${esc(f.error)}</span>`).join(''));
    } catch (e) { status(`<span class="is-bad">! ${esc(e.message)}</span>`); }
    finally { setBusy(false); }
  }

  // ---------- 文件夹 ----------
  function resultHtml(results) {
    return results.map(r => `<li class="${r.ok ? 'is-ok' : 'is-bad'}">${esc(r.folder)}${r.ok ? '' : `：${esc(r.error)}`}</li>`).join('');
  }

  async function applyFolders(icon, restore) {
    const folders = $('icFolders').value;
    if (!folders.trim()) { await window.appAlert('请先在框里填文件夹路径（每行一个）。'); return; }
    if (busy) return;
    setBusy(true);
    try {
      const r = restore ? await post('/api/icons/restore', { folders }) : await post('/api/icons/apply', { id: icon.id, folders });
      const bad = r.results.length - r.ok_count;
      lastResults = { id: icon.id, html: resultHtml(r.results) };
      status(bad ? `<span class="is-warn">! ${restore ? '恢复默认' : '套用'}成功 ${r.ok_count} 个，${bad} 个没成功（见右边）</span>`
        : `<span class="is-ok">✓ ${restore ? '已恢复默认图标' : `已把「${esc(icon.name)}」套到`} ${r.ok_count} 个文件夹${restore ? '' : '（看不到变化就在资源管理器按 F5）'}</span>`);
      log(restore ? '恢复默认' : '套到文件夹', r);
      await reload(restore ? '恢复默认后' : '套用后');
    } catch (e) {
      status(`<span class="is-bad">! ${esc(e.message)}</span>`);
    } finally { setBusy(false); }
  }

  async function restoreOne(folder) {
    const ok = await window.appConfirm(`把这个文件夹恢复成默认图标？\n${folder}`, { okText: '恢复默认' });
    if (!ok) return;
    try {
      const r = await post('/api/icons/restore', { folders: [folder] });
      const res = r.results[0];
      status(res.ok ? `<span class="is-ok">✓ 已恢复默认图标</span>` : `<span class="is-bad">! ${esc(res.error)}</span>`);
      await reload('单个恢复后');
    } catch (e) { status(`<span class="is-bad">! ${esc(e.message)}</span>`); }
  }

  // ---------- 工具箱 ----------
  async function setTool(icon, on) {
    const tool = $('icTool').value;
    const title = window.Tools?.byId[tool]?.title || tool;
    try {
      await post('/api/icons/tool', { tool, id: on ? icon.id : null });
      status(`<span class="is-ok">✓ 工具箱里「${esc(title)}」${on ? `换成了「${esc(icon.name)}」` : '恢复了默认图标'}（回主页就能看到）</span>`);
      await reload('工具箱图标');
    } catch (e) { status(`<span class="is-bad">! ${esc(e.message)}</span>`); }
  }

  // ---------- 批量指定：每个文件夹挑自己的图标 ----------
  const cleanPath = v => String(v ?? '').trim().replace(/^["']+|["']+$/g, '').trim();
  const pathKey = p => p.toLowerCase();
  const saveAssign = () => {
    prefs.assignRows = assignRows.map(r => ({path: r.path, icon: r.icon}));
    prefs.ibParent = $('ibParent').value;
    prefs.ibPaths = $('ibPaths').value;
    savePrefs();
  };

  function addAssignRows(entries) {
    let added = 0;
    for (const entry of entries) {
      const path = cleanPath(entry.path);
      if (!path) continue;
      const row = assignRows.find(r => pathKey(r.path) === pathKey(path));
      if (row) { if (!row.icon && entry.icon) row.icon = entry.icon; continue; }
      if (assignRows.length >= 200) break;
      assignRows.push({path, icon: entry.icon || '', state: '', message: ''});
      added++;
    }
    renderAssign(); saveAssign();
    return added;
  }

  function renderAssign() {
    const chosen = assignRows.filter(r => r.icon && byId(r.icon)).length;
    $('ibApply').disabled = busy || !chosen;
    $('ibApply').textContent = chosen ? `应用到 ${chosen} 个文件夹` : '应用到文件夹';
    $('ibRows').innerHTML = assignRows.length ? assignRows.map((r, n) => {
      const icon = byId(r.icon);
      const note = r.message ? `<span class="ib-note">${esc(r.message)}</span>` : r.state === 'ok' ? '<span class="ib-note">✓ 已应用</span>' : '';
      return `<li class="ib-row${r.state ? ` is-${r.state}` : ''}">
        <button type="button" class="ib-slot${icon ? ' is-set' : ''}" data-slot="${n}" title="给这个文件夹挑图标" aria-label="给 ${esc(r.path)} 挑图标">${icon ? `<img src="${pngUrl(icon)}" alt="${esc(icon.name)}">` : '+'}</button>
        <div><span class="ib-path">${esc(r.path)}</span>${note}</div>
        <button type="button" class="ib-drop-row" data-drop="${n}" title="移除这一行" aria-label="移除这一行">✕</button>
      </li>`;
    }).join('') : '<li class="ib-empty">列表还是空的：上面粘贴路径生成，或列一个父目录的子文件夹。</li>';
  }

  function openPicker(row, slot) {
    pickerRow = row;
    const picker = $('ibPicker');
    picker.innerHTML = lib.icons.map(i => `<button type="button" class="ib-pick" data-pick="${i.id}"><img src="${pngUrl(i)}" alt=""><span>${esc(i.name)}</span></button>`).join('')
      || '<p class="ib-empty">图标库还是空的</p>';
    picker.hidden = false;
    const rect = slot.getBoundingClientRect();
    const width = picker.offsetWidth, height = picker.offsetHeight;
    picker.style.left = `${Math.max(8, Math.min(rect.left, window.innerWidth - width - 8))}px`;
    picker.style.top = `${rect.bottom + 6 + height > window.innerHeight - 8 ? Math.max(8, rect.top - height - 6) : rect.bottom + 6}px`;
  }
  const closePicker = () => { $('ibPicker').hidden = true; pickerRow = null; };

  $('ibRows').addEventListener('click', e => {
    const slot = e.target.closest('[data-slot]');
    if (slot) { openPicker(assignRows[+slot.dataset.slot], slot); return; }
    const drop = e.target.closest('[data-drop]');
    if (drop) { assignRows.splice(+drop.dataset.drop, 1); closePicker(); renderAssign(); saveAssign(); }
  });
  $('ibPicker').addEventListener('click', e => {
    const tile = e.target.closest('[data-pick]');
    if (!tile || !pickerRow) return;
    pickerRow.icon = tile.dataset.pick; pickerRow.state = ''; pickerRow.message = '';
    closePicker(); renderAssign(); saveAssign();
  });
  document.addEventListener('click', e => {
    if (!$('ibPicker').hidden && !e.target.closest('#ibPicker') && !e.target.closest('[data-slot]')) closePicker();
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closePicker(); });
  $('ibFromPaths').addEventListener('click', () => {
    const added = addAssignRows($('ibPaths').value.split(/\r?\n/).map(path => ({path, icon: ''})));
    $('ibStatus').innerHTML = added ? `<span class="is-ok">✓ 加进列表 ${added} 个文件夹</span>`
      : '<span class="is-warn">没有新增文件夹（路径为空或和列表重复）</span>';
  });
  $('ibListSub').addEventListener('click', async () => {
    const folder = $('ibParent').value.trim();
    if (!folder) { await window.appAlert('请先填父目录路径。'); return; }
    try {
      const response = await fetch(`/api/icons/subfolders?folder=${encodeURIComponent(folder)}`, {cache: 'no-store'});
      const r = await response.json().catch(() => ({}));
      if (!response.ok || !r.ok) throw Error(r.error || '列不了子文件夹');
      const added = addAssignRows(r.folders.map(f => ({path: f.path, icon: f.icon_id})));
      $('ibStatus').innerHTML = `<span class="is-ok">✓ 列出 ${r.folders.length} 个子文件夹，新增 ${added} 个进列表</span>`
        + (r.total > r.folders.length ? `<br><span class="is-warn">! 共 ${r.total} 个子文件夹，只列前 200 个</span>` : '');
    } catch (e) { $('ibStatus').innerHTML = `<span class="is-bad">! ${esc(e.message)}</span>`; }
  });
  $('ibClear').addEventListener('click', () => { assignRows = []; closePicker(); renderAssign(); saveAssign(); $('ibStatus').textContent = ''; });
  $('ibPaths').addEventListener('input', saveAssign);
  $('ibParent').addEventListener('input', saveAssign);
  $('ibApply').addEventListener('click', async () => {
    const chosen = assignRows.filter(r => r.icon && byId(r.icon));
    if (!chosen.length || busy) return;
    setBusy(true);
    const groups = new Map();
    for (const row of chosen) { if (!groups.has(row.icon)) groups.set(row.icon, []); groups.get(row.icon).push(row.path); }
    const results = new Map();
    let done = 0;
    try {
      for (const [iconId, paths] of groups) {
        $('ibStatus').textContent = `正在套「${byId(iconId).name}」：第 ${done + 1} / ${groups.size} 组（${paths.length} 个文件夹）…`;
        try {
          const r = await post('/api/icons/apply', {id: iconId, folders: paths});
          for (const res of r.results) results.set(pathKey(res.folder), res);
        } catch (e) { for (const p of paths) results.set(pathKey(p), {ok: false, error: e.message}); }
        done++;
      }
      for (const row of assignRows) {
        const res = results.get(pathKey(row.path));
        if (!res) continue;
        row.state = res.ok ? 'ok' : 'bad'; row.message = res.ok ? '' : res.error;
      }
      const ok = [...results.values()].filter(r => r.ok).length;
      const bad = results.size - ok;
      const skipped = assignRows.length - chosen.length;
      $('ibStatus').innerHTML = (bad ? `<span class="is-warn">! 套成功 ${ok} 个文件夹，${bad} 个没成功（看行内说明）</span>`
        : `<span class="is-ok">✓ 已给 ${ok} 个文件夹各指定图标（看不到变化就在资源管理器按 F5）</span>`)
        + (skipped ? `<br><span class="is-warn">跳过 ${skipped} 行没挑图标的</span>` : '');
      log('批量指定', [...groups.entries()], [...results.values()]);
      await reload('批量指定后');
    } catch (e) {
      $('ibStatus').innerHTML = `<span class="is-bad">! ${esc(e.message)}</span>`;
    } finally { setBusy(false); }
  });

  // ---------- 拖放 / 事件 ----------
  const hasFiles = e => [...(e.dataTransfer?.types || [])].includes('Files');
  const drop = $('icDrop');
  drop.addEventListener('click', () => $('icFile').click());
  drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('icFile').click(); } });
  $('icFile').addEventListener('change', e => { addFiles(e.target.files); e.target.value = ''; });
  $('icReplaceFile').addEventListener('change', e => {
    const file = e.target.files[0]; e.target.value = '';
    const icon = byId(selected);
    if (file && icon) replaceImage(icon, file);
  });
  // 拖到页面任何地方（除了右边大图）都算加新图标；浏览器默认会直接打开图片，这里拦掉
  window.addEventListener('dragover', e => { if (hasFiles(e)) { e.preventDefault(); drop.classList.add('is-over'); } });
  window.addEventListener('dragleave', e => { if (!e.relatedTarget) drop.classList.remove('is-over'); });
  window.addEventListener('drop', e => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    drop.classList.remove('is-over');
    addFiles(e.dataTransfer.files);
  });

  $('icGrid').addEventListener('click', e => {
    const card = e.target.closest('.ic-card');
    if (!card || busy) return;
    if (!batchMode) { select(card.dataset.id); return; }
    if (checked.has(card.dataset.id)) checked.delete(card.dataset.id); else checked.add(card.dataset.id);
    renderGrid(); renderBatch();
    document.querySelectorAll('.ic-card').forEach(next => { if (next.dataset.id === card.dataset.id) next.focus(); });
  });
  $('icBatchToggle').addEventListener('click', () => {
    batchMode = !batchMode;
    if (!batchMode) checked.clear();
    $('icBatchStatus').textContent = '';
    renderBatch(); renderGrid();
  });
  $('icBatchAll').addEventListener('click', () => { visible().forEach(i => checked.add(i.id)); renderGrid(); renderBatch(); });
  $('icBatchClear').addEventListener('click', () => { checked.clear(); renderGrid(); renderBatch(); });
  $('icBatchMove').addEventListener('click', () => batchAction(false));
  $('icBatchDelete').addEventListener('click', () => batchAction(true));
  $('icDuplicates').addEventListener('click', () => {
    duplicatesOnly = !duplicatesOnly;
    if (duplicatesOnly) { category = null; $('icSearch').value = ''; }
    renderCats(); renderGrid();
  });
  $('icCats').addEventListener('click', e => {
    const chip = e.target.closest('.ic-cat');
    if (!chip) return;
    category = chip.dataset.cat === '*' ? null : chip.dataset.cat;
    renderCats();
    renderGrid();
    log('看分类', category === null ? '全部' : catLabel(category));
  });
  $('icSearch').addEventListener('input', renderGrid);
  $('icCatRenameBtn').addEventListener('click', async () => {
    const name = $('icCatRename').value.trim();
    if (category === null || name === category) return;
    try {
      const r = await post('/api/icons/category', { old: category, new: name });
      category = name;
      status(`<span class="is-ok">✓ 分类改名了，${r.count} 个图标改到「${esc(catLabel(name))}」</span>`);
      await reload('分类改名');
    } catch (e) { status(`<span class="is-bad">! ${esc(e.message)}</span>`); }
  });
  $('icOpenDir').addEventListener('click', async () => {
    try { const r = await post('/api/icons/open_dir', {}); status(`📂 已打开 ${esc(r.dir)}`); }
    catch (e) { status(`<span class="is-bad">! ${esc(e.message)}</span>`); }
  });

  reload('打开页面').catch(e => {
    $('icGrid').innerHTML = `<p class="ic-empty">读不了图标库：${esc(e.message)}</p>`;
    renderAssign();
    console.error('[设置图标] 读不了图标库', e);
  });
})();
