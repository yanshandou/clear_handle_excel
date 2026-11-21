/* 目录结构模板管理：左边模板列表，右边树编辑；整树保存到 /api/ticket_structures/save */
(function () {
  const list = document.getElementById('tsList');
  const editor = document.getElementById('tsEditor');
  const status = document.getElementById('tsStatus');
  const newName = document.getElementById('tsNewName');
  const createBtn = document.getElementById('tsCreate');
  const renameInput = document.getElementById('tsRename');
  const renameBtn = document.getElementById('tsRenameBtn');
  const deleteBtn = document.getElementById('tsDelete');

  let templates = [];
  let currentId = '';
  let tree = [];
  let dirty = false;
  let warnings = new Map();
  let icons = [];            // 图标库（只读）：目录节点挑图标用，管理图标去「设置图标」页
  let iconPop = null;

  const newId = () => crypto.randomUUID().replace(/-/g, '').slice(0, 12);
  const current = () => templates.find(t => t.id === currentId);
  const countNodes = nodes => nodes.reduce((sum, n) => sum + 1 + (n.kind === 'dir' ? countNodes(n.children || []) : 0), 0);

  async function api(path, payload) {
    const response = await fetch(path, payload === undefined ? { cache: 'no-store' }
      : { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    const data = await response.json();
    if (!response.ok || !data.ok) throw new Error(data.error || '操作失败，请重试。');
    return data;
  }

  function setStatus(text, kind) {
    status.textContent = text;
    status.className = `ts-status${kind ? ` is-${kind}` : ''}`;
  }

  function markDirty() {
    dirty = true;
    const hint = editor.querySelector('.ts-dirty');
    if (hint) hint.textContent = '有未保存的修改';
  }

  function setWarnings(items) {
    warnings = new Map((items || []).map(w => [w.id, w.reason]));
  }

  // ---------------------------------------------------------------- 模板列表

  function renderList() {
    list.replaceChildren();
    if (!templates.length) {
      const empty = document.createElement('li');
      empty.className = 'ts-empty';
      empty.textContent = '还没有模板，先起个名字新建一套。';
      list.append(empty);
    }
    for (const template of templates) {
      const item = document.createElement('li');
      const button = document.createElement('button');
      button.type = 'button';
      button.className = `ts-item${template.id === currentId ? ' is-on' : ''}`;
      const name = document.createElement('span');
      name.className = 'ts-item-name';
      name.textContent = template.name;
      const count = document.createElement('small');
      count.textContent = `${countNodes(template.tree || [])} 项`;
      button.append(name, count);
      const warnCount = (template.warnings || []).length;
      if (warnCount) {
        const badge = document.createElement('span');
        badge.className = 'ts-warn-badge';
        badge.textContent = `${warnCount} 个问题`;
        button.append(badge);
      }
      button.addEventListener('click', () => selectTemplate(template.id));
      item.append(button);
      list.append(item);
    }
    const hasCurrent = Boolean(current());
    renameInput.disabled = renameBtn.disabled = deleteBtn.disabled = !hasCurrent;
    if (hasCurrent) renameInput.placeholder = `给「${current().name}」改名`;
  }

  async function load(selectId) {
    const data = await api('/api/ticket_structures');
    templates = data.templates;
    const wanted = selectId || currentId;
    currentId = templates.some(t => t.id === wanted) ? wanted : (templates[0]?.id || '');
    const template = current();
    tree = template ? structuredClone(template.tree || []) : [];
    setWarnings(template?.warnings);
    dirty = false;
    renderList();
    renderEditor();
  }

  async function selectTemplate(id) {
    if (id === currentId) return;
    if (dirty && !await appConfirm('当前模板有未保存的修改，切换后会丢弃。确定切换吗？', { tone: 'warn', okText: '丢弃并切换' })) return;
    currentId = id;
    const template = current();
    tree = structuredClone(template.tree || []);
    setWarnings(template.warnings);
    dirty = false;
    renderList();
    renderEditor();
    setStatus('');
  }

  // ---------------------------------------------------------------- 目录图标

  const iconById = id => icons.find(i => i.id === id);
  const iconUrl = icon => `/api/icons/file?${new URLSearchParams({ id: icon.id, kind: 'png', v: icon.version })}`;

  async function loadIcons() {
    try {
      icons = (await api('/api/icons')).icons || [];
    } catch {
      icons = [];
    }
  }

  function closeIconPop() {
    if (!iconPop) return;
    iconPop.remove();
    iconPop = null;
  }

  function iconButton(node) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'ts-icon';
    const icon = node.icon ? iconById(node.icon) : null;
    if (icon) {
      const img = document.createElement('img');
      img.src = iconUrl(icon);
      img.alt = '';
      button.append(img);
      button.title = `文件夹图标：${icon.name}（点了换）`;
    } else {
      button.textContent = '📁';
      button.title = node.icon ? '这个图标已不在图标库里，建目录时用默认图标（点了重新挑）' : '默认文件夹图标（点了挑一个图标，建目录时顺便套上）';
      if (node.icon) button.classList.add('is-missing');
    }
    button.setAttribute('aria-label', `${node.name} 的文件夹图标`);
    button.addEventListener('click', event => { event.stopPropagation(); openIconPop(button, node); });
    return button;
  }

  function openIconPop(anchor, node) {
    closeIconPop();
    const pop = document.createElement('div');
    pop.className = 'ts-icon-pop';
    pop.setAttribute('role', 'dialog');
    pop.setAttribute('aria-label', `给「${node.name}」挑图标`);
    const head = document.createElement('div');
    head.className = 'ts-icon-pop-head';
    head.innerHTML = `<b></b><a href="/set-icons" target="_blank" rel="noopener">去图标库加图标</a>`;
    head.querySelector('b').textContent = `「${node.name}」用哪个图标`;
    const grid = document.createElement('div');
    grid.className = 'ts-icon-grid';
    const pick = id => {
      if (id) node.icon = id; else delete node.icon;
      markDirty();
      closeIconPop();
      renderEditor();
    };
    const option = (label, icon) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `ts-icon-opt${(node.icon || '') === (icon?.id || '') ? ' is-on' : ''}`;
      btn.title = icon ? (icon.category ? `${icon.name}（${icon.category}）` : icon.name) : '不套图标，用 Windows 默认文件夹图标';
      if (icon) {
        const img = document.createElement('img');
        img.src = iconUrl(icon);
        img.alt = '';
        img.loading = 'lazy';
        btn.append(img);
      } else {
        const glyph = document.createElement('span');
        glyph.className = 'ts-icon-default';
        glyph.textContent = '📁';
        btn.append(glyph);
      }
      const name = document.createElement('span');
      name.textContent = label;
      btn.append(name);
      btn.addEventListener('click', () => pick(icon?.id || ''));
      return btn;
    };
    grid.append(option('默认图标', null), ...icons.map(icon => option(icon.name, icon)));
    pop.append(head, grid);
    if (!icons.length) {
      const empty = document.createElement('p');
      empty.className = 'ts-hint';
      empty.textContent = '图标库还是空的：先到「设置图标」页拖进 PNG。';
      pop.append(empty);
    }
    document.body.append(pop);
    const rect = anchor.getBoundingClientRect();
    const width = pop.offsetWidth;
    pop.style.left = `${Math.max(8, Math.min(rect.left, window.innerWidth - width - 8))}px`;
    const below = rect.bottom + 6;
    pop.style.top = `${below + pop.offsetHeight > window.innerHeight - 8 ? Math.max(8, rect.top - pop.offsetHeight - 6) : below}px`;
    iconPop = pop;
    (pop.querySelector('.ts-icon-opt.is-on') || pop.querySelector('.ts-icon-opt'))?.focus();
  }

  document.addEventListener('click', event => { if (iconPop && !iconPop.contains(event.target)) closeIconPop(); });
  document.addEventListener('keydown', event => { if (event.key === 'Escape') closeIconPop(); });
  window.addEventListener('scroll', closeIconPop, true);
  window.addEventListener('resize', closeIconPop);

  // ---------------------------------------------------------------- 树编辑

  function makeNode(kind) {
    return kind === 'dir'
      ? { id: newId(), kind, name: '新建目录', children: [] }
      : { id: newId(), kind, name: '新建文件.txt', source: '' };
  }

  function addNode(children, kind) {
    const node = makeNode(kind);
    children.push(node);
    markDirty();
    renderEditor();
    const input = editor.querySelector(`[data-node-id="${node.id}"] .ts-name`);
    if (input) { input.focus(); input.select(); }
  }

  function removeNode(siblings, index) {
    siblings.splice(index, 1);
    markDirty();
    renderEditor();
  }

  function renderTree(nodes, container) {
    nodes.forEach((node, index) => {
      const item = document.createElement('li');
      item.className = `ts-node${warnings.has(node.id) ? ' is-warn' : ''}`;
      item.dataset.nodeId = node.id;
      const row = document.createElement('div');
      row.className = 'ts-row';
      let kind;
      if (node.kind === 'dir') {
        kind = iconButton(node);
      } else {
        kind = document.createElement('span');
        kind.className = 'ts-kind';
        kind.textContent = '📄';
        kind.setAttribute('aria-hidden', 'true');
      }
      const name = document.createElement('input');
      name.type = 'text';
      name.className = 'ts-name';
      name.maxLength = 60;
      name.value = node.name;
      name.setAttribute('aria-label', node.kind === 'dir' ? '目录名' : '文件名');
      name.addEventListener('input', () => { node.name = name.value; markDirty(); });
      row.append(kind, name);
      if (node.kind === 'file') {
        const source = document.createElement('input');
        source.type = 'text';
        source.className = 'ts-source';
        source.maxLength = 500;
        source.placeholder = '来源文件完整路径，如 D:\\模板\\xxx.xlsx';
        source.value = node.source || '';
        source.setAttribute('aria-label', `${node.name} 的来源路径`);
        source.title = warnings.get(node.id) || '';
        source.addEventListener('input', () => { node.source = source.value; markDirty(); });
        row.append(source);
      } else {
        const addDir = document.createElement('button');
        addDir.type = 'button';
        addDir.className = 'btn btn-outline';
        addDir.textContent = '+ 目录';
        addDir.addEventListener('click', () => addNode(node.children, 'dir'));
        const addFile = document.createElement('button');
        addFile.type = 'button';
        addFile.className = 'btn btn-outline';
        addFile.textContent = '+ 文件';
        addFile.addEventListener('click', () => addNode(node.children, 'file'));
        row.append(addDir, addFile);
      }
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'btn btn-danger';
      remove.textContent = '删除';
      remove.addEventListener('click', () => removeNode(nodes, index));
      row.append(remove);
      item.append(row);
      if (node.kind === 'dir') {
        const children = document.createElement('ul');
        children.className = 'ts-tree';
        renderTree(node.children || [], children);
        item.append(children);
      }
      container.append(item);
    });
  }

  function renderEditor() {
    editor.replaceChildren();
    const template = current();
    if (!template) {
      const empty = document.createElement('p');
      empty.className = 'ts-editor-empty';
      empty.textContent = '先在左边新建或选一套模板，再在这里编目录树。';
      editor.append(empty);
      return;
    }
    const toolbar = document.createElement('div');
    toolbar.className = 'ts-toolbar';
    const title = document.createElement('span');
    title.className = 'ts-title';
    title.textContent = `「${template.name}」的目录树`;
    const addDir = document.createElement('button');
    addDir.type = 'button';
    addDir.className = 'btn';
    addDir.textContent = '+ 根目录';
    addDir.addEventListener('click', () => addNode(tree, 'dir'));
    const addFile = document.createElement('button');
    addFile.type = 'button';
    addFile.className = 'btn';
    addFile.textContent = '+ 根文件';
    addFile.addEventListener('click', () => addNode(tree, 'file'));
    const save = document.createElement('button');
    save.type = 'button';
    save.className = 'btn btn-primary';
    save.textContent = '💾 保存';
    save.addEventListener('click', saveTree);
    const dirtyHint = document.createElement('span');
    dirtyHint.className = 'ts-dirty';
    dirtyHint.textContent = dirty ? '有未保存的修改' : '';
    toolbar.append(title, addDir, addFile, save, dirtyHint);
    editor.append(toolbar);

    const root = document.createElement('ul');
    root.className = 'ts-tree';
    renderTree(tree, root);
    editor.append(root);
    if (!tree.length) {
      const hint = document.createElement('p');
      hint.className = 'ts-hint';
      hint.textContent = '树还是空的：点「+ 根目录 / + 根文件」开始搭结构。';
      editor.append(hint);
    }

    const warnList = document.createElement('ul');
    warnList.className = 'ts-warnings';
    for (const [id, reason] of warnings) {
      const li = document.createElement('li');
      li.textContent = `⚠️ ${reason}`;
      li.dataset.nodeId = id;
      warnList.append(li);
    }
    editor.append(warnList);
    const hint = document.createElement('p');
    hint.className = 'ts-hint';
    hint.textContent = '点目录前面的 📁 挑图标，建目录时顺便套上；文件节点填来源完整路径，出票时复制过来；改完点「保存」。删除只影响模板，不会动已建好的 Ticket 文件夹。';
    editor.append(hint);
  }

  async function saveTree() {
    if (!current()) return;
    setStatus('正在保存…');
    try {
      const data = await api('/api/ticket_structures/save', { id: currentId, tree });
      const index = templates.findIndex(t => t.id === currentId);
      templates[index] = data.template;
      tree = structuredClone(data.template.tree || []);
      setWarnings(data.warnings);
      dirty = false;
      renderList();
      renderEditor();
      setStatus(data.warnings.length ? `已保存，但有 ${data.warnings.length} 个问题（见下方黄字）。` : '已保存。',
        data.warnings.length ? 'warn' : 'ok');
    } catch (err) {
      setStatus(err.message, 'bad');
    }
  }

  // ---------------------------------------------------------------- 新建 / 改名 / 删除

  createBtn.addEventListener('click', async () => {
    const name = newName.value.trim();
    if (!name) { setStatus('请先填新模板的名字。', 'warn'); newName.focus(); return; }
    if (dirty && !await appConfirm('当前模板有未保存的修改，新建后会丢弃。确定新建吗？', { tone: 'warn', okText: '丢弃并新建' })) return;
    try {
      const data = await api('/api/ticket_structures/create', { name });
      newName.value = '';
      await load(data.template.id);
      setStatus(`已新建「${data.template.name}」，开始搭目录树吧。`, 'ok');
    } catch (err) {
      setStatus(err.message, 'bad');
    }
  });

  renameBtn.addEventListener('click', async () => {
    const name = renameInput.value.trim();
    if (!current() || !name) { setStatus('请先填新的模板名。', 'warn'); renameInput.focus(); return; }
    try {
      const data = await api('/api/ticket_structures/rename', { id: currentId, name });
      const index = templates.findIndex(t => t.id === currentId);
      templates[index] = { ...templates[index], ...data.template, warnings: templates[index].warnings };
      renameInput.value = '';
      renderList();
      renderEditor();
      setStatus(`已改名为「${data.template.name}」。`, 'ok');
    } catch (err) {
      setStatus(err.message, 'bad');
    }
  });

  deleteBtn.addEventListener('click', async () => {
    const template = current();
    if (!template) return;
    if (!await appConfirm(`删除模板「${template.name}」？已经用它建好的 Ticket 文件夹不受影响。`, { tone: 'danger', okText: '删除' })) return;
    try {
      await api('/api/ticket_structures/delete', { id: currentId });
      await load('');
      setStatus(`已删除「${template.name}」。`, 'ok');
    } catch (err) {
      setStatus(err.message, 'bad');
    }
  });

  window.addEventListener('beforeunload', event => {
    if (dirty) event.preventDefault();
  });

  // 从图标库页加完图标回来，重新读一遍图标
  const iconKey = () => icons.map(i => `${i.id}:${i.version}:${i.name}`).join('|');
  window.addEventListener('focus', async () => {
    if (iconPop) return;
    const before = iconKey();
    await loadIcons();
    if (iconKey() !== before && current()) renderEditor();
  });

  loadIcons().then(load).catch(err => setStatus(err.message || '模板库加载失败。', 'bad'));
})();