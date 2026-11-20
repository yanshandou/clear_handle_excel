// 工具清单：主页工具箱（static/home_tiles.js）和各页「返回首页」都按这张表。
// - 2026-11-20 起主页是手机桌面式：所有工具都是一样大的图标，收在「工具箱」里，点开展开；主页上把图标拖进「流程」
// - group：工具箱里分在哪一段（GROUPS）；task：在主页对话框里做的工具（home.js openTask）
// - size：旧的大小卡片用的，现在不用了
// - 没有「本次流程 / 下一步」了：每一页做完只有「返回首页」，回到主页时那一页对应的卡片边框闪几下（static/home_tiles.js）
// - inc：这项已经被哪张大卡顺便做了（同一份 Excel 只开存一次，见 excel_workbook_io_optimization_analysis.md），小卡上写「已含在…里」
// - 工具图标可以在「设置图标」页换成图标库里的图（存在程序数据\\图标库\\index.json 的 tools，主页 home_tiles.js 读 /api/icons 套上）
// - ops：执行记录里的种类（core/operation_history.py），用来在卡片右上角打 ✓；没有记录的工具不打 ✓
(function () {
  const FOLDER_KEY = 'clear_handle_excel.lastFolder';

  const TOOLS = [
    { id: 'clean', group: 'main', size: 'big', title: '精简 Sheet', sub: '删备份表、空白表，先预览再执行', meta: '先预览，确认后一次做完', metaId: 'meta_clean_sheets',
      url: '/clean', ops: 'clean', icon: 'clean', tone: 'green' },
    { id: 'adjust', group: 'main', size: 'wide', title: '调整页面', sub: '宽一页 · 横竖 · 向下取整', meta: '导出 PDF，点按钮让程序改', metaId: 'meta_export_pdf',
      url: '/?task=export_pdf', action: "openTask('export_pdf')", task: 'export_pdf', pages: ['/pdf-review'], ops: 'pdf', icon: 'adjust', tone: 'teal' },
    { id: 'cells', group: 'main', size: 'wide', title: '查看切字', sub: '扫描标色 · 工具箱修格子 · 复检', meta: '尚未扫描', metaId: 'meta_clips',
      url: '/workbench', pages: ['/index.html'], ops: 'scan autofix', icon: 'cells', tone: 'yellow' },
    { id: 'upload', group: 'main', size: 'big', title: '上传 Excel', sub: '选文件、核对共享版本、传回共享', meta: '上传前每份只开存一次',
      url: '/distribute', ops: 'distribute', icon: 'crate', tone: 'purple', linkId: 'workTaskUploadLink' },
    { id: 'bl', group: 'clean', size: 'small', title: '排查 BL 超链接', sub: '找出断掉的超链接', url: '/repair-bl-links', ops: 'bl_scan bl_clear', icon: 'bl', tone: 'green', inc: 'clean', incKey: 'bl' },
    { id: 'links', group: 'clean', size: 'small', title: '修复引用错误', sub: '#REF! · 指向被删的表', url: '/repair-links', ops: 'links', icon: 'reflink', tone: 'green', inc: 'clean', incKey: 'links' },
    { id: 'hidden', group: 'clean', size: 'small', title: '修复隐藏页', sub: '对照参考恢复隐藏 / 显示', url: '/repair-hidden', ops: 'hidden', icon: 'eye', tone: 'green', inc: 'clean', incKey: 'hidden' },
    { id: 'font', group: 'clean', size: 'small', title: '统一字体', sub: '游ゴシック → ＭＳ Ｐゴシック', url: '/fix-fonts', icon: 'type', tone: 'green', inc: 'clean', incKey: 'font' },
    { id: 'colors', group: 'upload', size: 'small', title: '修复颜色', sub: '对照参考恢复残留底色', url: '/repair-colors', ops: 'colors', icon: 'palette', tone: 'purple', inc: 'upload' },
    { id: 'view', group: 'upload', size: 'small', title: '视图标准化', sub: 'A1 · 100% · 停在第一张表', url: '/?task=normalize_view', action: "openTask('normalize_view')", task: 'normalize_view', ops: 'view', icon: 'view', tone: 'purple', inc: 'upload' },
    { id: 'new', group: 'task', size: 'small', title: '新任务 · 拿文件', sub: '填共享路径，从共享拿文件', url: '/distribute/start', icon: 'new', tone: 'lime', linkId: 'workTaskNew' },
    { id: 'images', group: 'fix', size: 'small', title: '修复图片', sub: '截图变形 · 红框错位', url: '/fix-images', icon: 'imgfix', tone: 'teal' },
    { id: 'widths', group: 'fix', size: 'small', title: '修复列宽', sub: '按参考恢复被改过的列宽', url: '/fix-images?tool=widths', icon: 'widths', tone: 'teal' },
    { id: 'bundle', group: 'fix', size: 'small', title: '批量导出同名 PDF', sub: 'aaa.xlsx 旁边生成 aaa.pdf', url: '/pdf-bundle', icon: 'pdfstack', tone: 'blue' },
    { id: 'pick', group: 'fix', size: 'small', title: '挑文件复制', sub: '贴路径，保留目录只复制这些', url: '/pick-copy', icon: 'pick', tone: 'pink' },
    { id: 'json', group: 'fix', size: 'small', title: '按 JSON 复制 Excel', sub: '按上传记录复制、建目录', url: '/distribute?mode=json', icon: 'json', tone: 'pink' },
    { id: 'pcl', group: 'check', size: 'small', title: 'PCL 单体测试检查', sub: '22 项清单 · 源码校验 · 覆盖率 · 测试报告', url: '/pcl', icon: 'check', tone: 'yellow' },
    { id: 'report', group: 'check', size: 'small', title: 'Excel 对比报告', sub: '任选两份 Excel · 生成 HTML 报告', url: '/excel-report', icon: 'report', tone: 'blue' },
    { id: 'pdfdiff', group: 'check', size: 'small', title: 'PDF 对比', sub: '拖进来左右对比 · 每边可接好几份', url: '/pdf-diff', icon: 'pdfdiff', tone: 'teal' },
    { id: 'icons', group: 'misc', size: 'small', title: '设置图标', sub: '透明 PNG 生成图标 · 图标库 · 套到文件夹', url: '/set-icons', icon: 'appicon', tone: 'pink' },
    { id: 'finish', group: 'task', size: 'small', title: '完成任务', sub: '收尾检查：都传回共享了没有', url: '/finish?start=1', pages: ['/finish'], icon: 'finish', tone: 'yellow', linkId: 'workTaskFinish' },
  ];
  const byId = Object.fromEntries(TOOLS.map(t => [t.id, t]));

  // 工具箱里的分组（2026-11-20 主页改成手机桌面式工具箱，点开以后按这个顺序分几段）
  const GROUPS = [
    { id: 'main', title: '常用' },
    { id: 'task', title: '任务开始 / 收尾' },
    { id: 'clean', title: '精简时可以顺便做' },
    { id: 'upload', title: '上传时会顺便做' },
    { id: 'fix', title: '修复 · 导出 · 复制' },
    { id: 'check', title: '检查 · 对比' },
    { id: 'misc', title: '其他小工具' },
  ];

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

  function chainOptions() {
    const out = {};
    CHAIN.forEach(c => { const v = read(c.storage); out[c.key] = v === null ? c.def : v === '1'; });
    return out;
  }

  function setChainOption(key, on) {
    const c = CHAIN.find(x => x.key === key);
    if (c) write(c.storage, on ? '1' : '0');
  }

  // 这一页是哪个工具：按网址认（/distribute?mode=json 是按 JSON 复制，/distribute/start 是新任务页）
  function currentToolId() {
    const path = location.pathname.replace(/\/+$/, '') || '/';
    if (path === '/distribute') return new URLSearchParams(location.search).get('mode') === 'json' ? 'json' : 'upload';
    if (path === '/fix-images') return new URLSearchParams(location.search).get('tool') === 'widths' ? 'widths' : 'images';
    const tool = TOOLS.find(t => !t.action && t.url.split('?')[0] === path || (t.pages || []).includes(path));
    return tool ? tool.id : '';
  }

  // 「返回首页」的网址：带上从哪一页回来的，主页让那张卡片的边框闪几下
  function homeUrl(id) {
    const from = id || currentToolId();
    return from ? `/?from=${encodeURIComponent(from)}` : '/';
  }

  // 各页做完以后唯一的按钮
  function homeButtonHtml(cls = 'btn btn-primary') {
    return `<a class="${cls} back-home-btn" href="${homeUrl()}">🏠 返回首页</a>`;
  }

  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  window.Tools = { TOOLS, byId, GROUPS, CHAIN, folder, chainOptions, setChainOption, currentToolId, homeUrl, homeButtonHtml, esc };

  // 工具页上所有回主页的链接（页头的「主页」、左上角的名字、「返回首页」）都带上这一页是谁
  if (location.pathname !== '/' && location.pathname !== '/home.html') {
    const mark = () => {
      const url = homeUrl();
      if (url === '/') return;
      document.querySelectorAll('a[href="/"]').forEach(a => { a.href = url; });
    };
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mark);
    else mark();
  }
})();
