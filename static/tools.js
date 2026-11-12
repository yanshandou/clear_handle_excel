// 工具清单（用户 2026-11-11 要）：主页是一张大小卡片拼起来的卡片页，这里是所有卡片的一张表。
// - size：big 大卡（常用流程，里面顺便把相关的几步一起做）、wide 横卡、small 小卡（每一步单独用）
// - 没有「本次流程 / 下一步」了：每一页做完只有「返回首页」，回到主页时那一页对应的卡片边框闪几下（static/home_tiles.js）
// - inc：这项已经被哪张大卡顺便做了（同一份 Excel 只开存一次，见 excel_workbook_io_optimization_analysis.md），小卡上写「已含在…里」
// - ops：执行记录里的种类（core/operation_history.py），用来在卡片右上角打 ✓；没有记录的工具不打 ✓
(function () {
  const FOLDER_KEY = 'clear_handle_excel.lastFolder';

  const TOOLS = [
    { id: 'clean', size: 'big', title: '精简 Sheet', sub: '删备份表、空白表，先预览再执行', meta: '先预览，确认后一次做完', metaId: 'meta_clean_sheets',
      url: '/clean', ops: 'clean', sprite: 'clean', tone: 'green' },
    { id: 'adjust', size: 'wide', title: '调整页面', sub: '宽一页 · 横竖 · 向下取整', meta: '导出 PDF，点按钮让程序改', metaId: 'meta_export_pdf',
      url: '/?task=export_pdf', action: "openTask('export_pdf')", pages: ['/pdf-review'], ops: 'pdf', sprite: 'adjust', tone: 'teal' },
    { id: 'cells', size: 'wide', title: '查看切字', sub: '扫描标色 · 工具箱修格子 · 复检', meta: '尚未扫描', metaId: 'meta_clips',
      url: '/workbench', pages: ['/index.html'], ops: 'scan autofix', sprite: 'cells', tone: 'yellow' },
    { id: 'upload', size: 'big', title: '上传 Excel', sub: '选文件、核对共享版本、传回共享', meta: '上传前每份只开存一次',
      url: '/distribute', ops: 'distribute', sprite: 'upload', tone: 'purple', linkId: 'workTaskUploadLink' },
    { id: 'bl', size: 'small', title: '排查 BL 超链接', sub: '找出断掉的超链接', url: '/repair-bl-links', ops: 'bl_scan bl_clear', vector: 'link', tone: 'green', inc: 'clean', incKey: 'bl' },
    { id: 'links', size: 'small', title: '修复引用错误', sub: '#REF! · 指向被删的表', url: '/repair-links', ops: 'links', vector: 'link', tone: 'green', inc: 'clean', incKey: 'links' },
    { id: 'hidden', size: 'small', title: '修复隐藏页', sub: '对照参考恢复隐藏 / 显示', url: '/repair-hidden', ops: 'hidden', vector: 'eye', tone: 'green', inc: 'clean', incKey: 'hidden' },
    { id: 'font', size: 'small', title: '统一字体', sub: '游ゴシック → ＭＳ Ｐゴシック', url: '/fix-fonts', vector: 'type', tone: 'green', inc: 'clean', incKey: 'font' },
    { id: 'colors', size: 'small', title: '修复颜色', sub: '对照参考恢复残留底色', url: '/repair-colors', ops: 'colors', vector: 'palette', tone: 'purple', inc: 'upload' },
    { id: 'view', size: 'small', title: '视图标准化', sub: 'A1 · 100% · 停在第一张表', url: '/?task=normalize_view', action: "openTask('normalize_view')", ops: 'view', vector: 'target', tone: 'purple', inc: 'upload' },
    { id: 'new', size: 'small', title: '新任务 · 拿文件', sub: '填共享路径，从共享拿文件', url: '/distribute/start', sprite: 'new', tone: 'lime', linkId: 'workTaskNew' },
    { id: 'images', size: 'small', title: '修复图片', sub: '截图变形 · 红框错位', url: '/fix-images', vector: 'image', tone: 'teal' },
    { id: 'widths', size: 'small', title: '修复列宽', sub: '按参考恢复被改过的列宽', url: '/fix-images?tool=widths', vector: 'grid', tone: 'teal' },
    { id: 'bundle', size: 'small', title: '批量导出同名 PDF', sub: 'aaa.xlsx 旁边生成 aaa.pdf', url: '/pdf-bundle', vector: 'document', tone: 'blue' },
    { id: 'pick', size: 'small', title: '挑文件复制', sub: '贴路径，保留目录只复制这些', url: '/pick-copy', vector: 'layers', tone: 'pink' },
    { id: 'json', size: 'small', title: '按 JSON 复制 Excel', sub: '按上传记录复制、建目录', url: '/distribute?mode=json', vector: 'layers', tone: 'pink' },
    { id: 'pcl', size: 'small', title: 'PCL 单体测试检查', sub: '22 项清单 · 源码校验 · 覆盖率 · 测试报告', url: '/pcl', vector: 'check', tone: 'yellow' },
    { id: 'finish', size: 'small', title: '完成任务', sub: '收尾检查：都传回共享了没有', url: '/finish?start=1', pages: ['/finish'], sprite: 'finish', tone: 'yellow', linkId: 'workTaskFinish' },
  ];
  const byId = Object.fromEntries(TOOLS.map(t => [t.id, t]));

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

  window.Tools = { TOOLS, byId, CHAIN, folder, chainOptions, setChainOption, currentToolId, homeUrl, homeButtonHtml, esc };

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
