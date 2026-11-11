(() => {
  const reference = document.querySelector('#cleanOldFolder, #repairOldFolder, #prOldFolder, #imgOldFolder');
  const work = document.querySelector('#cleanFolderInput, #repairFolder, #prFolder, #imgFolder');
  if (!reference || !work) return;

  const key = 'clear_handle_excel.referenceFolders';
  const normalize = path => path.trim().replace(/\\/g, '/').toLowerCase();
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem(key) || '{}') || {}; } catch (_) {}

  const details = document.createElement('details');
  details.className = 'reference-folders';
  details.hidden = true;
  details.open = true;
  details.innerHTML = '<summary>为处理文件夹单独指定参考目录</summary>'
    + '<p>留空时使用默认参考文件夹。系统会按 Excel 文件名搜索所选参考目录及其子文件夹。</p>';
  const fields = document.createElement('fieldset');
  details.append(fields);
  const error = document.createElement('p');
  error.setAttribute('role', 'status');
  details.append(error);
  (reference.closest('label, .input-groove-wrap') || reference).insertAdjacentElement('afterend', details);

  let query = '', lookup = null, rendered = '', timer;
  const isBatch = value => value.includes('\n') || /[\\/]批量路径[\\/][a-f0-9]{24}$/.test(value);
  function list(value) {
    if (!value) return Promise.resolve([]);
    if (lookup && query === value) return lookup;
    query = value;
    const request = fetch('/api/workbook-paths', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ folder: value }),
    }).then(async response => {
      const data = await response.json();
      if (!response.ok || !data.ok) throw Error(data.error || '无法读取处理文件夹');
      return data.folders;
    }).catch(error => {
      if (lookup === request) lookup = null;
      throw error;
    });
    lookup = request;
    return lookup;
  }

  function render(folders) {
    const stamp = JSON.stringify(folders);
    details.hidden = !folders.length;
    error.textContent = '';
    if (rendered === stamp) return;
    rendered = stamp;
    fields.replaceChildren(...folders.map(folder => {
      const label = document.createElement('label');
      const path = document.createElement('span');
      path.textContent = folder;
      const input = document.createElement('input');
      input.type = 'text';
      input.placeholder = '使用默认参考文件夹';
      input.value = saved[normalize(folder)] || '';
      input.addEventListener('input', () => {
        const id = normalize(folder);
        if (input.value.trim()) saved[id] = input.value.trim();
        else delete saved[id];
        try { localStorage.setItem(key, JSON.stringify(saved)); } catch (_) {}
      });
      label.append(path, input);
      return label;
    }));
  }

  function refresh() {
    clearTimeout(timer);
    timer = setTimeout(async () => {
      const value = work.value.trim();
      try {
        const folders = await list(value);
        if (value === work.value.trim()) render(isBatch(value) || folders.some(path => saved[normalize(path)]) ? folders : []);
      } catch (e) {
        if (value !== work.value.trim()) return;
        render([]);
        details.hidden = false;
        error.textContent = e.message;
      }
    }, 200);
  }

  async function read(folder = work.value) {
    const folders = await list(folder.trim());
    return Object.fromEntries(folders.filter(path => saved[normalize(path)]).map(path => [path, saved[normalize(path)]]));
  }

  new MutationObserver(() => { fields.disabled = reference.disabled; }).observe(reference, {
    attributes: true, attributeFilter: ['disabled'],
  });
  fields.disabled = reference.disabled;
  work.addEventListener('input', refresh);
  window.ReferenceFolders = { read, refresh };
  refresh();
})();
