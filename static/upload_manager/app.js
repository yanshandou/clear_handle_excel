const STORAGE_KEY = 'upload-desk-data-v1';
const seed = [
  ['10要件定義\\10画面','00_画面遷移図',2,2,'2026-10-20','剛'],
  ['10要件定義\\10画面','01_CM共通機能',145,59,'2026-10-20','剛'],
  ['10要件定義\\10画面','02_SG総合メニュー',6,6,'2026-10-20','呉'],
  ['10要件定義\\10画面','03_KS居宅支援業務',50,17,'2026-10-21','蒋'],
  ['10要件定義\\10画面','04_KH介護保険業務',63,63,'2026-10-20','呉'],
  ['10要件定義\\10画面','05_IR医療保険業務',22,22,'2026-10-20','呉'],
  ['10要件定義\\10画面','06_SS障害者支援業務',57,57,'2026-10-21','呉'],
  ['10要件定義\\10画面','07_SCスケジュール業務',41,41,'2026-10-23','呉'],
  ['10要件定義\\10画面','08_NY入金管理業務',13,13,'2026-10-23','呉'],
  ['10要件定義\\10画面','09_KY簡易給与計算業務',19,19,'2026-10-23','呉'],
  ['10要件定義\\10画面','10_KZ口座振替業務',6,6,'2026-10-23','呉'],
  ['10要件定義\\10画面','11_SYシステム管理業務',26,26,'2026-10-23','呉'],
  ['10要件定義\\10画面','12_KE契約管理業務',2,2,'2026-10-26','呉'],
  ['10要件定義\\10画面','13_TL有償機能',14,14,'2026-10-26','呉'],
  ['10要件定義\\10画面','14_個別機能',5,5,'2026-10-26','呉'],
  ['10要件定義\\10画面','15_FLスケジュール＆ウェイジ',3,3,'2026-10-26','呉'],
  ['10要件定義\\10画面','20_概要書詳細',31,31,'2026-10-26','呉'],
  ['10要件定義\\10画面','22_TBタブレット',14,14,'2026-10-26','呉'],
  ['10要件定義\\10画面','30_規則',28,28,'2026-10-26','呉'],
  ['10要件定義\\20帳票','00.資料',41,0,'','後回し'],
  ['10要件定義\\20帳票','01.共通',49,0,'','後回し'],
  ['10要件定義\\20帳票','02.居宅支援業務',21,0,'','後回し'],
  ['10要件定義\\20帳票','03.介護保険業務',52,0,'','後回し'],
  ['10要件定義\\20帳票','04.障害者福祉業務',106,0,'','後回し'],
  ['10要件定義\\20帳票','05.スケジュール業務',22,0,'','後回し'],
  ['10要件定義\\20帳票','30.規則',1,1,'2026-10-29','金'],
  ['20基本設計\\10DB','01.画面',424,354,'2026-10-20','徐'],
  ['20基本設計\\10DB','02.帳票',329,0,'',''],
  ['20基本設計\\10DB','03.バッチ',4,0,'','後回し'],
  ['20基本設計\\10DB','11.テーブル定義書',482,0,'','後回し'],
  ['20基本設計\\10DB','12.ビュー定義書',327,0,'','後回し'],
  ['20基本設計\\10DB','13.特別クラス概要書',103,0,'','後回し'],
  ['20基本設計\\10DB','21.ER図',276,0,'','後回し'],
  ['20基本設計\\20帳票','01_CM共通機能',55,55,'2026-10-27','呉'],
  ['20基本設計\\20帳票','03_KS居宅支援業務',61,61,'2026-10-21','李'],
  ['20基本設計\\20帳票','04_KH介護保険業務',102,102,'2026-10-22','呉'],
  ['20基本設計\\20帳票','05_IR医療保険業務',31,0,'2026-10-29','呉'],
  ['20基本設計\\20帳票','06_SS障害者支援業務',160,160,'2026-10-28','呉'],
  ['20基本設計\\20帳票','07_SCスケジュール業務',31,0,'2026-10-29','呉'],
  ['20基本設計\\20帳票','08_NY入金管理業務',8,8,'2026-10-20','李'],
  ['20基本設計\\20帳票','09_KY給与計算業務',16,0,'',''],
  ['20基本設計\\20帳票','10_KZ口座振替業務',4,4,'2026-10-27','李'],
  ['20基本設計\\20帳票','13_TL有償機能',19,19,'2026-10-26','李'],
  ['20基本設計\\20帳票','14_KO個別機能',11,11,'2026-10-20','李'],
  ['20基本設計\\30BL','BL',16,0,'',''],
  ['20基本設計\\40外部連携','外部連携',3,0,'',''],
  ['20基本設計\\50Common','Common',4,0,'','']
].map(([group,name,total,completedCount,date,owner],index) => ({id:`seed-${index}`,group,name,total,processed:completedCount,complete:total > 0 && completedCount === total,uploaded:0,date,owner}));

const $ = selector => document.querySelector(selector);
const escapeHTML = value => String(value).replace(/[&<>"']/g, character => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character]));
let items;
const fileMode = location.protocol === 'file:';
try { const stored = JSON.parse(localStorage.getItem(STORAGE_KEY)); items = fileMode && Array.isArray(stored) ? stored : fileMode ? seed : []; } catch { items = fileMode ? seed : []; }
let currentGroup = 'all';
const currentOwners = new Set();
let sortKey = 'name';
let sortDirection = 1;
let editingId = null;
let toastTimer;
const collapsedFolders = new Set();
let planPreview = null;
let visibleFolderTree = null;

async function request(path, options = {}) {
  if (fileMode) throw new Error('请通过本地服务打开网页，当前页面仅用于导出旧浏览器备份。');
  const response = await fetch(path, options);
  if (!response.ok) {
    const data = await response.json().catch(() => ({}));
    throw new Error(data.error || `请求失败（${response.status}）`);
  }
  return response.json();
}
async function refreshItems() {
  items = await request('/api/upload-manager/items');
  if (currentGroup !== 'all' && !items.some(item => item.group === currentGroup)) currentGroup = 'all';
  render();
}
async function writeItem(item, isNew = false) {
  const saved = await request(isNew ? '/api/upload-manager/items' : `/api/upload-manager/items/${encodeURIComponent(item.id)}`, {
    method: isNew ? 'POST' : 'PUT', headers: {'Content-Type':'application/json'}, body: JSON.stringify(item)
  });
  if (isNew) items.push(saved);
  else items = items.map(entry => entry.id === saved.id ? saved : entry);
  return saved;
}
function downloadBackup(data, filename, contentType) {
  const url = URL.createObjectURL(new Blob([data], {type:contentType}));
  const link = document.createElement('a'); link.href = url; link.download = filename; document.body.append(link); link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
function icons() { if (window.lucide) window.lucide.createIcons(); }
function askConfirm(message, {title = '请确认', okText = '确定', cancelText = '取消', danger = false, alertOnly = false} = {}) {
  const dialog = $('#confirm-dialog');
  $('#confirm-title').textContent = title;
  $('#confirm-message').textContent = message;
  const ok = $('#confirm-ok'), cancel = $('#confirm-cancel');
  ok.textContent = okText; cancel.textContent = cancelText;
  cancel.hidden = alertOnly;
  ok.classList.toggle('danger', danger);
  return new Promise(resolve => {
    const finish = result => { dialog.removeEventListener('close', onClose); dialog.removeEventListener('click', onBackdrop); ok.onclick = cancel.onclick = null; if (dialog.open) dialog.close(); resolve(result); };
    const onClose = () => finish(false);
    const onBackdrop = event => { if (event.target === dialog) finish(false); };
    ok.onclick = () => finish(true);
    cancel.onclick = () => finish(false);
    dialog.addEventListener('close', onClose);
    dialog.addEventListener('click', onBackdrop);
    dialog.showModal();
    (danger ? cancel : ok).focus();
  });
}
function notify(message) { const toast = $('#toast'); toast.textContent = message; toast.classList.add('show'); clearTimeout(toastTimer); toastTimer = setTimeout(() => toast.classList.remove('show'), 2200); }
function isComplete(item) { return item.complete ?? item.uploaded > 0; }
const folderStatuses = [['unstarted','未完成'], ['processing','完成中...'], ['ready','已完成'], ['uploading','上传中...'], ['uploaded','已上传']];
function statusOf(item) {
  if (item.statusOverride) return item.statusOverride;
  if (item.uploaded && item.uploaded === item.total) return 'uploaded';
  if (item.uploaded || item.plannedDate) return 'uploading';
  if (!item.date) return 'unstarted';
  return !item.totalMissing && item.processed != null && item.processed === item.total ? 'ready' : 'processing';
}
function statusLabel(item) { return folderStatuses.find(([value]) => value === statusOf(item))?.[1] || '未完成'; }
const statusIcons = {unstarted:'circle-dashed', processing:'loader', ready:'package-check', uploading:'cloud-upload', uploaded:'circle-check'};
function nextStatus(item) {
  const order = folderStatuses.map(([value]) => value);
  const next = order[(order.indexOf(statusOf(item)) + 1) % order.length];
  return next === 'uploaded' && (!item.total || item.totalMissing) ? order[0] : next;
}
function withFolderStatus(item, status) {
  const uploading = status === 'uploading' || status === 'uploaded';
  return {...item, statusOverride:status, complete:['ready','uploading','uploaded'].includes(status),
    uploaded:status === 'uploaded' ? item.total : status === 'uploading' && statusOf(item) === 'uploading' ? item.uploaded : 0,
    plannedDate:uploading ? item.plannedDate || '' : '', uploader:uploading ? item.uploader || '' : ''};
}
function percentOf(item) { return item.total ? Math.round(item.uploaded / item.total * 100) : 0; }
function groupItems() { return currentGroup === 'all' ? items : items.filter(item => item.group === currentGroup); }
function folderTree(records) {
  const root = {children: new Map(), records: []};
  for (const item of records) {
    let node = root;
    const groupParts = item.group.split(/[\\/¥]/).filter(Boolean);
    if (!groupParts.length) groupParts.push(item.group);
    const path = [];
    for (const part of item.parentOnly ? groupParts : [...groupParts, item.name]) {
      path.push(part);
      if (!node.children.has(part)) node.children.set(part, {name:part, path:JSON.stringify(path), children:new Map(), records:[], item:null});
      node = node.children.get(part);
    }
    if (!node.item) node.item = item;
    else node.records.push(item);
  }
  return root;
}
function setFoldersExpanded(root, expanded) {
  function visit(node) {
    if (node.path && (node.children.size || node.records.length)) {
      if (expanded) collapsedFolders.delete(node.path); else collapsedFolders.add(node.path);
    }
    for (const child of node.children.values()) visit(child);
  }
  visit(root);
}
function folderChoices(records, group, editingId) {
  const taken = name => records.some(item => item.id !== editingId && item.group === group && item.name === name);
  const grouped = records.filter(item => item.group === group);
  const names = [...new Set(grouped.filter(item => !item.parentOnly).map(item => item.name))].sort((first, second) => first.localeCompare(second, 'zh-CN'));
  const parent = grouped.find(item => item.parentOnly);
  return [...(parent ? [{name:'', unavailable:taken(parent.name)}] : []), ...names.map(name => ({name, unavailable:taken(name)}))];
}
function folderToggle(node) {
  const collapsed = collapsedFolders.has(node.path);
  return `<button type="button" class="tree-toggle" data-tree-path="${escapeHTML(node.path)}" aria-expanded="${!collapsed}" aria-label="${collapsed ? '展开' : '折叠'} ${escapeHTML(node.name)}" title="${collapsed ? '展开' : '折叠'}"><i data-lucide="${collapsed ? 'chevron-right' : 'chevron-down'}"></i></button>`;
}
function folderRow(item, depth, node = null, assignments = null) {
  const itemStatus = statusOf(item);
  const nextLabel = folderStatuses.find(([value]) => value === nextStatus(item))[1];
  const toggle = node && (node.children.size || node.records.length) ? folderToggle(node) : '<span class="tree-toggle-placeholder"></span>';
  const assignment = assignments?.has(item.id) ? `<span class="tree-assignment">分配给 ${escapeHTML(assignments.get(item.id))}</span>` : '';
  return `<tr><td><div class="folder-cell tree-cell" style="padding-left:${depth * 20}px">${toggle}<span class="folder-icon"><i data-lucide="folder"></i></span><div><span class="folder-name" title="${escapeHTML(item.name)}">${escapeHTML(item.name)}</span>${assignment}</div></div></td><td><span class="file-count">${item.totalMissing ? '—' : item.total.toLocaleString()}</span></td><td><span class="processed-count">${item.processed == null ? '—' : item.processed.toLocaleString()}</span></td><td>${item.date ? escapeHTML(item.date.replaceAll('-',' / ')) : '—'}</td><td><button type="button" class="owner-tag owner-cell-tag" data-owner="${escapeHTML(item.owner)}" aria-pressed="${currentOwners.has(item.owner)}" title="按负责人筛选">${escapeHTML(item.owner || '未分配')}</button></td><td>${escapeHTML(item.uploader || '—')}</td><td>${item.uploadDate ? escapeHTML(item.uploadDate.replaceAll('-',' / ')) : '—'}</td><td class="progress-cell"><div class="progress-info"><span>${item.uploaded.toLocaleString()} / ${item.total.toLocaleString()}</span><strong>${percentOf(item)}%</strong></div><div class="row-progress"><span style="width:${percentOf(item)}%"></span></div></td><td>${item.plannedDate ? escapeHTML(item.plannedDate.replaceAll('-',' / ')) : '—'}</td><td><button type="button" class="folder-status ${itemStatus}" data-status-id="${escapeHTML(item.id)}" aria-label="${escapeHTML(item.name)}的文件夹状态：${statusLabel(item)}，点击切换为${nextLabel}" title="点击切换为${nextLabel}"><i data-lucide="${statusIcons[itemStatus]}"></i><span>${statusLabel(item)}</span></button></td><td><div class="row-actions"><button class="icon-button" data-action="edit" data-id="${escapeHTML(item.id)}" title="编辑文件夹" aria-label="编辑 ${escapeHTML(item.name)}"><i data-lucide="pencil"></i></button><button class="icon-button" data-action="delete" data-id="${escapeHTML(item.id)}" title="删除文件夹" aria-label="删除 ${escapeHTML(item.name)}"><i data-lucide="trash-2"></i></button></div></td></tr>`;
}
function renderFolderTree(root, assignments = null) {
  const rows = [];
  function recordRow(item, depth, node = null) {
    return folderRow(item, depth, node, assignments);
  }
  function visit(node, depth) {
    if (node.item) rows.push(recordRow(node.item, depth, node));
    else rows.push(`<tr class="tree-parent-row"><td><div class="folder-cell tree-cell" style="padding-left:${depth * 20}px">${folderToggle(node)}<span class="folder-icon"><i data-lucide="folder"></i></span><span class="folder-name">${escapeHTML(node.name)}</span></div></td><td colspan="10"></td></tr>`);
    if (collapsedFolders.has(node.path)) return;
    for (const child of node.children.values()) visit(child, depth + 1);
    for (const item of node.records) rows.push(recordRow(item, depth + 1));
  }
  for (const child of root.children.values()) visit(child, 0);
  for (const item of root.records) rows.push(recordRow(item, 0));
  return rows.join('');
}
function matchesUploadDate(item, uploadDate) {
  return !uploadDate || item.uploadDate === uploadDate || item.plannedDate === uploadDate;
}
function setPlanPreview(plan, people, snapshot, label) {
  if (snapshot !== plannerSnapshot()) {
    clearPlanPreview();
    resetAllocationPreviews();
    allocationPlans = [];
    $('#planner-results').hidden = true;
    plannerMessage('文件夹数据已更改，请重新生成方案。', true);
    return;
  }
  planPreview = {assignments:new Map(plan.assignments.map(({item, person}) => [item.id, people[person]])), snapshot, label};
  collapsedFolders.clear();
  render();
}
function clearPlanPreview() {
  if (!planPreview) return;
  planPreview = null;
  render();
}
function render() {
  if (planPreview && planPreview.snapshot !== plannerSnapshot()) {
    planPreview = null;
    resetAllocationPreviews();
    allocationPlans = [];
    $('#planner-results').hidden = true;
    plannerMessage('文件夹数据已更改，请重新生成方案。', true);
  }
  if (typeof updatePersonPickerSummary === 'function') {
    updatePersonPickerSummary('people');
    updatePersonPickerSummary('owner');
  }
  const groups = [...new Set(items.map(item => item.group))];
  $('#folder-nav').innerHTML = groups.map(group => `<button class="nav-item ${currentGroup === group ? 'active' : ''}" data-group="${escapeHTML(group)}" title="${escapeHTML(group)}"><i data-lucide="folder-closed"></i><span>${escapeHTML(group)}</span></button>`).join('');
  document.querySelector('[data-group="all"]').classList.toggle('active', currentGroup === 'all');
  $('#breadcrumb-current').textContent = currentGroup === 'all' ? '总览' : currentGroup;
  $('#list-title').firstChild.textContent = planPreview ? '方案预览 ' : currentGroup === 'all' ? '文件夹明细 ' : `${currentGroup} `;
  $('#plan-preview-bar').hidden = !planPreview;
  $('#plan-preview-name').textContent = planPreview?.label || '';
  $('.table-toolbar').hidden = !!planPreview;
  $('.owner-filters').hidden = !!planPreview;
  const scoped = groupItems();
  const owners = [...new Set(scoped.map(item => item.owner))].sort((first,second) => first.localeCompare(second, 'zh-CN'));
  for (const owner of currentOwners) if (!owners.includes(owner)) currentOwners.delete(owner);
  $('#owner-tags').innerHTML = `<button type="button" class="owner-tag" data-owner-all aria-pressed="${currentOwners.size === 0}">全部</button>` + owners.map(owner => `<button type="button" class="owner-tag" data-owner="${escapeHTML(owner)}" aria-pressed="${currentOwners.has(owner)}">${escapeHTML(owner || '未分配')}</button>`).join('');
  const total = scoped.reduce((sum,item) => sum + item.total, 0);
  const uploaded = scoped.reduce((sum,item) => sum + item.uploaded, 0);
  const overall = total ? Math.round(uploaded / total * 100) : 0;
  $('#total-files').textContent = total.toLocaleString();
  $('#uploaded-files').textContent = uploaded.toLocaleString();
  $('#overall-percent').textContent = `${overall}%`;
  $('#overall-bar').style.width = `${overall}%`;
  const query = $('#search').value.trim().toLocaleLowerCase();
  const status = $('#status-filter').value;
  const uploadDate = $('#upload-date-filter').value;
  $('#clear-upload-date').hidden = !uploadDate;
  const visible = planPreview ? items.filter(item => planPreview.assignments.has(item.id)) : scoped.filter(item => (!query || `${item.name} ${item.group} ${item.owner}`.toLocaleLowerCase().includes(query)) && (status === 'all' || statusOf(item) === status) && (!currentOwners.size || currentOwners.has(item.owner)) && matchesUploadDate(item, uploadDate));
  visible.sort((first,second) => {
    const a = sortKey === 'status' ? statusOf(first) : first[sortKey] ?? '';
    const b = sortKey === 'status' ? statusOf(second) : second[sortKey] ?? '';
    return (typeof a === 'number' ? a - b : String(a).localeCompare(String(b), 'zh-CN', {numeric:true})) * sortDirection;
  });
  $('#row-count').textContent = visible.length;
  if (typeof updateDateExport === 'function') updateDateExport(uploadDate, visible);
  $('#table-summary').textContent = planPreview ? `方案包含 ${visible.length} 个文件夹` : `显示 ${visible.length} 个文件夹，共 ${scoped.length} 个`;
  $('#empty-state').hidden = visible.length > 0;
  $('#empty-state strong').textContent = items.length ? '没有找到相关文件夹' : '暂无文件夹';
  $('#empty-state span').textContent = items.length ? '试试调整搜索内容或筛选条件' : '导入 Excel 或添加文件夹';
  visibleFolderTree = folderTree(visible);
  $('#table-body').innerHTML = renderFolderTree(visibleFolderTree, planPreview?.assignments);
  if (planPreview) $('#table-body').querySelectorAll('button:not(.tree-toggle), select').forEach(control => { control.disabled = true; });
  icons();
}
function renderChildChoices(form, selected) {
  const choice = form.elements['name-choice'];
  const group = form.elements['group-choice'].value;
  choice.innerHTML = '<option value="" disabled>请选择子文件夹</option>' + folderChoices(items, group && group !== 'new' ? JSON.parse(group) : '', editingId).map(option => `<option value="${escapeHTML(JSON.stringify(option.name))}" ${option.unavailable ? 'disabled' : ''}>${escapeHTML(option.name || '仅父文件夹')}</option>`).join('');
  choice.value = selected === null ? '' : JSON.stringify(selected);
  if (!choice.selectedOptions.length || choice.selectedOptions[0].disabled) choice.value = '';
}
function syncNewField(form, field) {
  const input = form.elements[field];
  const creating = form.elements[`${field}-choice`].value === 'new';
  input.hidden = !creating;
  input.disabled = !creating;
  input.required = creating;
}
function openDialog(item = null, groupOnly = false) {
  editingId = item?.id || null;
  const form = $('#item-form'); form.reset();
  $('#dialog-title').textContent = item ? '编辑文件夹' : '添加文件夹';
  $('#form-error').textContent = '';
  for (const field of ['name','group','total','processed','uploaded','date','owner']) form.elements[field].value = item?.[field] ?? (field === 'group' && currentGroup !== 'all' && !groupOnly ? currentGroup : field === 'total' || field === 'uploaded' ? 0 : '');
  if (!item) form.elements.group.value = '';
  form.elements.name.hidden = !!item;
  form.elements.name.disabled = !!item;
  form.elements['name-choice'].hidden = !item;
  form.elements['name-choice'].disabled = !item;
  const groups = [...new Set(items.map(entry => entry.group))].sort((first, second) => first.localeCompare(second, 'zh-CN'));
  const owners = [...new Set(items.map(entry => entry.owner).filter(Boolean))].sort((first, second) => first.localeCompare(second, 'zh-CN'));
  form.elements['group-choice'].innerHTML = '<option value="" disabled>请选择父文件夹</option>' + groups.map(group => `<option value="${escapeHTML(JSON.stringify(group))}">${escapeHTML(group)}</option>`).join('') + (item ? '' : '<option value="new">新建父文件夹...</option>');
  form.elements['group-choice'].value = item ? JSON.stringify(item.group) : groupOnly || !groups.length ? 'new' : currentGroup !== 'all' && groups.includes(currentGroup) ? JSON.stringify(currentGroup) : '';
  form.elements['owner-choice'].innerHTML = `<option value="${escapeHTML(JSON.stringify(''))}">未分配</option>` + owners.map(owner => `<option value="${escapeHTML(JSON.stringify(owner))}">${escapeHTML(owner)}</option>`).join('') + (item ? '' : '<option value="new">新建负责人...</option>');
  form.elements['owner-choice'].value = JSON.stringify(item?.owner || '');
  syncNewField(form, 'group');
  syncNewField(form, 'owner');
  if (item) renderChildChoices(form, item.parentOnly ? '' : item.name);
  if (item?.parentOnly) form.elements.name.value = '';
  if (item?.totalMissing) form.elements.total.value = '';
  $('#item-dialog').showModal();
  form.elements[item ? 'name-choice' : groupOnly ? 'group' : 'name'].focus();
}
document.addEventListener('click', async event => {
  const treeToggle = event.target.closest('[data-tree-path]');
  if (treeToggle) {
    const path = treeToggle.dataset.treePath;
    if (collapsedFolders.has(path)) collapsedFolders.delete(path); else collapsedFolders.add(path);
    render();
    [...$('#table-body').querySelectorAll('[data-tree-path]')].find(button => button.dataset.treePath === path)?.focus();
    return;
  }
  const ownerTag = event.target.closest('.owner-tag');
  if (ownerTag) {
    const owner = ownerTag.dataset.owner;
    const rowId = ownerTag.closest('tr')?.querySelector('[data-action="edit"]')?.dataset.id;
    if (ownerTag.hasAttribute('data-owner-all')) currentOwners.clear();
    else if (currentOwners.has(owner)) currentOwners.delete(owner);
    else currentOwners.add(owner);
    collapsedFolders.clear();
    render();
    const rowButton = rowId && [...$('#table-body').querySelectorAll('[data-action="edit"]')].find(button => button.dataset.id === rowId)?.closest('tr')?.querySelector('.owner-cell-tag');
    const filterButton = [...$('#owner-tags').children].find(button => ownerTag.hasAttribute('data-owner-all') ? button.hasAttribute('data-owner-all') : button.dataset.owner === owner);
    (rowButton || filterButton)?.focus({preventScroll:true});
    return;
  }
  const nav = event.target.closest('[data-group]');
  if (nav) { planPreview = null; currentGroup = nav.dataset.group; currentOwners.clear(); render(); return; }
  const action = event.target.closest('[data-action]');
  if (!action) return;
  const item = items.find(entry => entry.id === action.dataset.id);
  if (!item) return;
  if (action.dataset.action === 'edit') { if (fileMode) { notify('请通过本地服务打开网页'); return; } openDialog(item); return; }
  if (action.dataset.action === 'delete') {
    if (!await askConfirm(`确定删除“${item.name}”吗？此操作无法撤销。`, {title:'删除文件夹',okText:'删除',danger:true})) return;
    action.disabled = true;
    try {
      await request(`/api/upload-manager/items/${encodeURIComponent(item.id)}`, {method:'DELETE',headers:{'Content-Type':'application/json'},body:JSON.stringify({version:item.version})});
      items = items.filter(entry => entry.id !== item.id);
      if (currentGroup !== 'all' && !items.some(entry => entry.group === currentGroup)) currentGroup = 'all';
      render(); notify('文件夹已删除');
    } catch (error) { notify(error.message); action.disabled = false; }
    return;
  }
});
$('#table-body').addEventListener('click', async event => {
  const control = event.target.closest('[data-status-id]');
  if (!control || control.disabled) return;
  const item = items.find(entry => entry.id === control.dataset.statusId);
  if (!item) return;
  const status = nextStatus(item);
  if ((item.uploaded && status !== 'uploaded') || (item.plannedDate && !['uploading','uploaded'].includes(status))) {
    if (!await askConfirm('切换状态将清除现有上传数量或上传计划，继续吗？', {title:'切换状态'})) return;
  }
  control.disabled = true;
  try {
    const saved = await writeItem(withFolderStatus(item, status));
    render(); notify(`状态已更新为${statusLabel(saved)}`);
    [...$('#table-body').querySelectorAll('[data-status-id]')].find(button => button.dataset.statusId === item.id)?.focus({preventScroll:true});
  } catch (error) { control.disabled = false; notify(error.message); }
});
$('#item-form').addEventListener('submit', async event => {
  event.preventDefault();
  const form = event.currentTarget;
  const child = editingId ? JSON.parse(form.elements['name-choice'].value) : form.elements.name.value.trim();
  const group = form.elements['group-choice'].value === 'new' ? form.elements.group.value.trim() : JSON.parse(form.elements['group-choice'].value);
  const parentOnly = !child;
  const name = child || group.split(/[\\/¥]/).pop().trim();
  const totalMissing = form.elements.total.value === '';
  const total = totalMissing ? 0 : Number(form.elements.total.value);
  const processed = form.elements.processed.value === '' ? null : Number(form.elements.processed.value);
  const uploaded = Number(form.elements.uploaded.value);
  if (!name || !group) { $('#form-error').textContent = '请填写有效的父文件夹路径。'; return; }
  if (items.some(item => item.id !== editingId && item.group === group && item.name === name)) { $('#form-error').textContent = '该父文件夹下已存在同名文件夹。'; return; }
  if (!Number.isInteger(total) || total < 0 || total > 1000000) { $('#form-error').textContent = '请输入有效的文件数量。'; return; }
  if (processed !== null && (!Number.isInteger(processed) || processed < 0 || processed > total)) { $('#form-error').textContent = '完成数量必须是非负整数，且不能大于文件数量。'; return; }
  if (!Number.isInteger(uploaded) || uploaded < 0 || uploaded > total) { $('#form-error').textContent = '已上传数量必须是非负整数，且不能大于文件数量。'; return; }
  const existing = items.find(entry => entry.id === editingId);
  const countsChanged = !existing || total !== existing.total || totalMissing !== !!existing.totalMissing || processed !== (existing.processed ?? null);
  const complete = countsChanged ? !totalMissing && processed !== null && processed === total : isComplete(existing);
  if (uploaded && !complete) { $('#form-error').textContent = '未完成的文件夹不能有已上传数量，请调整数量或先将上传数量清零。'; return; }
  const owner = form.elements['owner-choice'].value === 'new' ? form.elements.owner.value.trim() : JSON.parse(form.elements['owner-choice'].value);
  if (form.elements['owner-choice'].value === 'new' && !owner) { $('#form-error').textContent = '请输入新负责人名称。'; return; }
  if (existing?.plannedDate && !complete) { $('#form-error').textContent = '已有上传计划，请先取消计划再修改完成数量。'; return; }
  const uploadedChanged = existing && uploaded !== existing.uploaded;
  const statusOverride = uploadedChanged ? (uploaded === total && uploaded > 0 ? 'uploaded' : uploaded || existing.plannedDate ? 'uploading' : '')
    : existing && (existing.plannedDate || uploaded || (!countsChanged && form.elements.date.value === existing.date)) ? existing.statusOverride : '';
  const data = {...existing,name,group,total,totalMissing,processed,uploaded,complete,date:form.elements.date.value,owner,parentOnly,statusOverride};
  const submit = form.querySelector('[type="submit"]'); submit.disabled = true;
  try {
    await writeItem(data, !editingId);
    if (currentGroup !== 'all' && !items.some(entry => entry.group === currentGroup)) currentGroup = group;
    $('#item-dialog').close(); render(); notify(editingId ? '文件夹信息已更新' : '文件夹已添加');
  } catch (error) { $('#form-error').textContent = error.message; }
  finally { submit.disabled = false; }
});
$('#add-item').addEventListener('click', () => fileMode ? notify('请通过本地服务打开网页') : openDialog());
$('#add-group').addEventListener('click', () => fileMode ? notify('请通过本地服务打开网页') : openDialog(null,true));
$('#close-dialog').addEventListener('click', () => $('#item-dialog').close());
$('#cancel-dialog').addEventListener('click', () => $('#item-dialog').close());
$('#item-form').elements['group-choice'].addEventListener('change', event => {
  const form = event.currentTarget.form;
  syncNewField(form, 'group');
  if (editingId) {
    const choice = form.elements['name-choice'];
    renderChildChoices(form, choice.value ? JSON.parse(choice.value) : null);
  }
});
$('#item-form').elements['owner-choice'].addEventListener('change', event => syncNewField(event.currentTarget.form, 'owner'));
$('#exit-plan-preview').addEventListener('click', clearPlanPreview);
for (const [selector, expanded] of [['#expand-all-folders', true], ['#collapse-all-folders', false]]) {
  $(selector).addEventListener('click', () => { if (visibleFolderTree) { setFoldersExpanded(visibleFolderTree, expanded); render(); } });
}
function renderFiltered() { collapsedFolders.clear(); render(); }
$('#search').addEventListener('input', renderFiltered);
$('#status-filter').addEventListener('change', renderFiltered);
$('#upload-date-filter').addEventListener('change', renderFiltered);
$('#clear-upload-date').addEventListener('click', () => {
  $('#upload-date-filter').value = '';
  renderFiltered();
  $('#upload-date-filter').focus();
});
document.querySelectorAll('[data-sort]').forEach(button => button.addEventListener('click', () => { const key = button.dataset.sort; sortDirection = sortKey === key ? -sortDirection : 1; sortKey = key; render(); }));
$('#sort-order').addEventListener('click', () => { sortDirection *= -1; render(); notify(sortDirection === 1 ? '升序排列' : '降序排列'); });
document.addEventListener('keydown', event => { if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') { event.preventDefault(); $('#search').focus(); } });
$('#export').addEventListener('click', () => { if (fileMode) { notify('请通过本地服务导出 Excel'); return; } const link = document.createElement('a'); link.href = '/api/upload-manager/export'; link.download = 'upload-desk.xlsx'; document.body.append(link); link.click(); link.remove(); });
$('#export-customer').addEventListener('click', async () => {
  if (fileMode) { notify('请通过本地服务导出 Excel'); return; }
  try {
    const response = await fetch('/api/upload-manager/export/customer');
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data.error || `请求失败（${response.status}）`);
    }
    downloadBackup(await response.blob(), 'uploaded-folders.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  } catch (error) { notify(`客户表导出失败：${error.message}`); }
});
$('#import-excel').addEventListener('click', () => fileMode ? notify('请通过本地服务导入 Excel') : $('#excel-file').click());
$('#import-backup').addEventListener('click', () => fileMode ? notify('请在本地服务页面导入备份') : $('#backup-file').click());
$('#export-backup').addEventListener('click', async () => {
  if (fileMode) {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (!stored) { notify('当前浏览器没有旧数据'); return; }
    downloadBackup(stored, 'upload-desk-browser-backup.json', 'application/json;charset=utf-8');
    return;
  }
  try {
    const data = await request('/api/upload-manager/export/json');
    downloadBackup(JSON.stringify(data, null, 2), 'upload-desk-backup.json', 'application/json;charset=utf-8');
  } catch (error) { notify(`JSON 导出失败：${error.message}`); }
});
$('#excel-file').addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    const preview = await request('/api/upload-manager/import/preview', {method:'POST',headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},body:file});
    if (!await askConfirm(`工作表：${preview.sheet}\n共 ${preview.total} 条；新增 ${preview.inserted} 条，更新 ${preview.updated} 条。\n匹配的文件夹保留已有上传数量和名义上传者，其他记录保留。确定导入吗？`, {title:'导入 Excel',okText:'导入'})) return;
    const result = await request('/api/upload-manager/import', {method:'POST',headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},body:file});
    await refreshItems(); notify(`导入完成：新增 ${result.inserted} 条，更新 ${result.updated} 条`);
  } catch (error) { await askConfirm(`导入失败：${error.message}`, {title:'导入失败',alertOnly:true}); }
  finally { event.target.value = ''; }
});
$('#backup-file').addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    const records = JSON.parse(await file.text());
    if (!Array.isArray(records) || !records.length || records.some(item => !item || typeof item !== 'object' || Array.isArray(item))) throw new Error('备份为空或格式不正确');
    if (!await askConfirm(`将导入 ${records.length} 条 JSON 备份，匹配的文件夹会被更新，其他记录保留。建议先导出备份。确定继续吗？`, {title:'导入 JSON 备份',okText:'导入'})) return;
    const normalized = records.map(item => ({...item,complete:item.complete ?? item.uploaded > 0,processed:item.processed ?? null}));
    const result = await request('/api/upload-manager/migrate', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(normalized)});
    await refreshItems(); notify(`JSON 备份已导入：新增 ${result.inserted} 条，更新 ${result.updated} 条`);
  } catch (error) { await askConfirm(`导入失败：${error.message}`, {title:'导入失败',alertOnly:true}); }
  finally { event.target.value = ''; }
});
$('#today').textContent = new Intl.DateTimeFormat('zh-CN',{year:'numeric',month:'long',day:'numeric'}).format(new Date());
render();
if (fileMode) {
  $('#storage-status').textContent = '旧数据只读，请导出备份';
} else {
  refreshItems().then(() => { $('#storage-status').textContent = '数据保存在 SQLite'; }).catch(error => { $('#storage-status').textContent = '数据库连接失败'; notify(error.message); });
}
