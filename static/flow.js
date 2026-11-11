// 本次流程（用户 2026-11-10 要）：主页工具箱里所有工具一张表，用户把要用的几张卡拖进「本次流程」轨道排成第 1 步、第 2 步……
// - 流程只管「这一页做完去哪一页」，每一页自己做的事不变（精简页照旧顺便做 BL / 引用 / 隐藏页 / 字体，上传页照旧顺便清标记、恢复颜色、标准化）
// - 按当前处理文件夹（任务的处理文件夹，localStorage lastFolder）各记一套，新任务默认 4 步：精简 → 调整页面 → 查看切字 → 上传
// - 工具页（非主页）顶部自动加一条流程条：第几步、做过的打 ✓、「下一步：X ➜」；这一页不在流程里就不显示
(function () {
  const FOLDER_KEY = 'clear_handle_excel.lastFolder';
  const FLOW_KEY = 'clear_handle_excel.flow';
  const DEFAULT_FLOW = ['clean', 'adjust', 'cells', 'upload'];

  // inc：这项已经被哪张大卡顺便做了（那张卡在流程里时，小卡上提示「已含在…里」）
  // ops：执行记录里的种类（core/operation_history.py），用来打 ✓；没有记录的工具不打 ✓
  const TOOLS = [
    { id: 'clean', title: '精简 Sheet', sub: '删备份表 · 顺便查 BL · 修引用错误', meta: '先预览，确认后一次做完', metaId: 'meta_clean_sheets',
      url: '/clean', ops: 'clean', icon: 'clean', tone: 'green', main: true },
    { id: 'adjust', title: '调整页面', sub: '宽一页 · 横竖 · 向下取整', meta: '导出 PDF，点按钮让程序改', metaId: 'meta_export_pdf',
      url: '/?task=export_pdf', action: "openTask('export_pdf')", pages: ['/pdf-review'], ops: 'pdf', icon: 'pdf', tone: 'teal', main: true },
    { id: 'cells', title: '查看切字', sub: '扫描标色 · 工具箱修格子 · 复检', meta: '尚未扫描', metaId: 'meta_clips',
      url: '/workbench', pages: ['/index.html'], ops: 'scan autofix', icon: 'inspect', tone: 'yellow', main: true },
    { id: 'upload', title: '上传 Excel', sub: '顺便清标记 · 恢复颜色 · 标准化', meta: '选择文件 · 核对共享版本 · 上传',
      url: '/distribute', ops: 'distribute', icon: 'distribute', tone: 'purple', main: true, linkId: 'workTaskUploadLink' },
    { id: 'bundle', title: '批量导出同名 PDF', sub: 'aaa.xlsx 旁边生成 aaa.pdf', meta: '所有表合成一份', url: '/pdf-bundle', emoji: '📕', tone: 'blue' },
    { id: 'json', title: '按 JSON 复制 Excel', sub: '按上传记录复制、建目录', meta: '只复制，上传仍人工', url: '/distribute?mode=json', emoji: '📋', tone: 'pink' },
    { id: 'colors', title: '修复颜色', sub: '对照参考恢复残留底色', meta: '上传时会顺便做', url: '/repair-colors', ops: 'colors', icon: 'colors', tone: 'orange', inc: 'upload' },
    { id: 'images', title: '修复图片', sub: '截图变形 · 红框错位', meta: '对照参考放回原样', url: '/fix-images', emoji: '🖼️', tone: 'teal' },
    { id: 'font', title: '统一字体', sub: '游ゴシック → ＭＳ Ｐゴシック', meta: '精简时可以顺便做', url: '/fix-fonts', emoji: '🔤', tone: 'green', inc: 'clean', incKey: 'font' },
    { id: 'links', title: '修复引用错误', sub: '#REF! · 指向被删的表', meta: '精简时可以顺便做', url: '/repair-links', ops: 'links', icon: 'references', tone: 'yellow', inc: 'clean', incKey: 'links' },
    { id: 'bl', title: '排查 BL 超链接', sub: '找出断掉的超链接', meta: '精简时可以顺便做', url: '/repair-bl-links', ops: 'bl_scan bl_clear', icon: 'bl-links', tone: 'purple', inc: 'clean', incKey: 'bl' },
    { id: 'hidden', title: '修复隐藏页', sub: '对照参考恢复隐藏 / 显示', meta: '精简时可以顺便做', url: '/repair-hidden', ops: 'hidden', icon: 'hidden', tone: 'blue', inc: 'clean', incKey: 'hidden' },
    { id: 'view', title: '视图标准化', sub: 'A1 · 100% · 停在第一张表', meta: '上传时会顺便做', url: '/?task=normalize_view', action: "openTask('normalize_view')", ops: 'view', icon: 'normalize', tone: 'pink', inc: 'upload' },
  ];
  const byId = Object.fromEntries(TOOLS.map(t => [t.id, t]));
  const FINISH = { id: 'finish', title: '完成任务', url: '/finish?start=1' };

  // 精简后顺便做的几项：和精简页的勾选框共用（static/clean.js）；隐藏页默认不做，其他默认做
  const CHAIN = [
    { key: 'bl', label: 'BL 超链接', storage: 'clear_handle_excel.cleanBl', def: true },
    { key: 'links', label: '引用错误', storage: 'clear_handle_excel.cleanLinks', def: true },
    { key: 'hidden', label: '隐藏页', storage: 'clear_handle_excel.cleanHidden', def: false },
    { key: 'font', label: '统一字体', storage: 'clear_handle_excel.cleanFont', def: true },
  ];

  function read(key) { try { return localStorage.getItem(key); } catch (_) { return null; } }
  function write(key, value) { try { localStorage.setItem(key, value); } catch (_) {} }
  function folder() { return (read(FOLDER_KEY) || '').trim(); }
  function flowKey() { const f = folder().toLowerCase(); return f ? `${FLOW_KEY}:${f}` : FLOW_KEY; }

  function load() {
    try {
      const saved = JSON.parse(read(flowKey()) || 'null');
      if (Array.isArray(saved)) return saved.filter(id => byId[id]).filter((id, i, all) => all.indexOf(id) === i);
    } catch (_) {}
    return DEFAULT_FLOW.slice();
  }

  function save(ids) {
    write(flowKey(), JSON.stringify(ids));
    window.dispatchEvent(new CustomEvent('flowchange', { detail: ids.slice() }));
  }

  function chainOptions() {
    const out = {};
    CHAIN.forEach(c => { const v = read(c.storage); out[c.key] = v === null ? c.def : v === '1'; });
    return out;
  }

  function setChainOption(key, on) {
    const c = CHAIN.find(x => x.key === key);
    if (c) write(c.storage, on ? '1' : '0');
  }

  // 这一页是哪个工具：按网址认（/distribute?mode=json 是按 JSON 复制，/distribute/start 是新任务页，不算）
  function currentToolId() {
    const path = location.pathname.replace(/\/+$/, '') || '/';
    if (path === '/distribute') return new URLSearchParams(location.search).get('mode') === 'json' ? 'json' : 'upload';
    const tool = TOOLS.find(t => t.url.split('?')[0] === path && !t.action || (t.pages || []).includes(path));
    return tool ? tool.id : '';
  }

  // 流程里 id 后面那一步；最后一步后面是「完成任务」；不在流程里返回 null
  function nextAfter(id) {
    const ids = load();
    const i = ids.indexOf(id);
    if (i < 0) return null;
    return i + 1 < ids.length ? byId[ids[i + 1]] : FINISH;
  }

  function firstStep() {
    const ids = load();
    return ids.length ? byId[ids[0]] : null;
  }

  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  async function doneMap() {
    const f = folder();
    if (!f) return {};
    try {
      const res = await fetch('/api/operation-history?folder=' + encodeURIComponent(f));
      return res.ok ? ((await res.json()).done || {}) : {};
    } catch (_) {
      return {};
    }
  }

  function isDone(tool, done) {
    return Boolean(tool.ops && tool.ops.split(' ').some(k => done[k] && done[k].completed));
  }

  // 工具页顶部的流程条
  async function renderPageBar() {
    const id = currentToolId();
    const ids = load();
    if (!id || !ids.includes(id) || document.getElementById('flowPageBar')) return;
    const header = document.querySelector('body > header');
    if (!header) return;
    const bar = document.createElement('nav');
    bar.id = 'flowPageBar';
    bar.className = 'flow-page-bar';
    bar.setAttribute('aria-label', '本次流程');
    header.insertAdjacentElement('afterend', bar);
    const paint = done => {
      const next = nextAfter(id);
      const steps = ids.map((sid, i) => {
        const t = byId[sid];
        const cls = sid === id ? 'cur' : (isDone(t, done) ? 'done' : '');
        const label = `${isDone(t, done) && sid !== id ? '✓' : i + 1} ${t.title}${sid === id ? '（这一页）' : ''}`;
        return `<a class="fpb-step ${cls}" href="${esc(t.url)}"${sid === id ? ' aria-current="step"' : ''}>${esc(label)}</a>`;
      });
      steps.push(`<a class="fpb-step" href="${FINISH.url}">🏁 ${FINISH.title}</a>`);
      bar.innerHTML = `<span class="fpb-label">🧭 本次流程</span>${steps.join('<span class="fpb-sep" aria-hidden="true">›</span>')}
        <a class="btn btn-primary btn-sm fpb-next" href="${esc(next.url)}">下一步：${esc(next.title)} ➜</a>
        <a class="fpb-edit" href="/" title="回主页拖动卡片改流程">改流程</a>`;
    };
    paint({});
    paint(await doneMap());
  }

  // 页头写死的「下一步」链接（data-flow-next="这一页的工具"）：这一页在流程里就改成流程里的下一步
  function updateNextLinks() {
    document.querySelectorAll('a[data-flow-next]').forEach(link => {
      const next = nextAfter(link.dataset.flowNext);
      if (!next) return;
      link.href = next.url;
      link.textContent = `下一步：${next.title} ➜`;
      link.title = `本次流程里的下一步：${next.title}`;
    });
  }

  window.Flow = { TOOLS, byId, FINISH, DEFAULT_FLOW, CHAIN, load, save, chainOptions, setChainOption, currentToolId, nextAfter, firstStep, doneMap, isDone, esc };

  if (location.pathname !== '/' && location.pathname !== '/home.html') {
    const init = () => { updateNextLinks(); renderPageBar(); };
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
  }
})();
