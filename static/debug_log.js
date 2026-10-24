// 调试日志：选一次任务，一键复制（太长时只复制开头、全部报错和最后一段），或下载整个 log.txt。
// 用法：在页面里放 <div data-debug-log-tools></div>，引入本文件即可。
(() => {
  async function copyPlain(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      const area = Object.assign(document.createElement('textarea'), { value: text });
      area.style.cssText = 'position:fixed;left:-9999px;top:0';
      document.body.append(area);
      area.select();
      const ok = document.execCommand('copy');
      area.remove();
      return ok;
    }
  }

  function flash(button, text) {
    const label = button.dataset.label || (button.dataset.label = button.textContent);
    button.textContent = text;
    clearTimeout(button.flashTimer);
    button.flashTimer = setTimeout(() => { button.textContent = label; }, 1800);
  }

  const KINDS = {
    batch_scan: '全自动 / 只扫描', batch_commit: '提交审核', batch_recheck: '复检', batch_open: '打开 Excel',
    batch_unhighlight: '清除标记颜色', clean_sheets: '精简 Sheet', export_pdf: '导出 PDF', normalize_view: '视图标准化',
  };
  // {id: '20261024_183012', kind: 'batch_scan'} → 10-24 18:30 全自动 / 只扫描
  const label = item => {
    const m = item.id.match(/^\d{4}(\d{2})(\d{2})_(\d{2})(\d{2})/);
    return `${m ? `${m[1]}-${m[2]} ${m[3]}:${m[4]}` : item.id} ${KINDS[item.kind] || item.kind}（${item.kb} KB）`;
  };

  function mount(box) {
    box.classList.add('debug-log-tools');
    box.innerHTML = '<select class="debug-log-select" aria-label="选择哪一次任务的调试日志"></select>'
      + '<button type="button" class="btn btn-outline btn-sm debug-log-copy" title="出问题时把这段贴给 Claude 排查">复制调试日志</button>'
      + '<button type="button" class="btn btn-outline btn-sm debug-log-download">下载 log.txt</button>';
    const select = box.querySelector('select');
    const query = () => select.value ? `?id=${encodeURIComponent(select.value)}` : '';

    async function refresh() {
      try {
        const data = await (await fetch('/api/debug-log/list', { cache: 'no-store' })).json();
        const keep = select.value;
        select.innerHTML = data.tasks.length
          ? data.tasks.map((item, i) => `<option value="${item.id}">${i === 0 ? '最近：' : ''}${label(item)}</option>`).join('')
          : '<option value="">还没有调试日志</option>';
        if (keep && data.tasks.some(item => item.id === keep)) select.value = keep;
      } catch { /* 服务没响应时下次再取 */ }
    }

    // 列表要读整个 log.txt：只在点开下拉框时刷新，不定时轮询
    select.addEventListener('focus', refresh);
    box.querySelector('.debug-log-copy').addEventListener('click', async event => {
      const button = event.currentTarget;
      try {
        const text = await (await fetch(`/api/debug-log/copy${query()}`, { cache: 'no-store' })).text();
        flash(button, await copyPlain(text) ? `已复制 ✓（${Math.ceil(text.length / 1000)}k 字）` : '复制失败，请点下载');
      } catch {
        flash(button, '读取失败，请点下载');
      }
    });
    box.querySelector('.debug-log-download').addEventListener('click', () => {
      window.location.href = '/api/debug-log/download';  // 整个 log.txt（最近 12 小时）
    });
    refresh();
  }

  document.querySelectorAll('[data-debug-log-tools]').forEach(mount);
})();
