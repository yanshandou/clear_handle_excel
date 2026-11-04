let allocationPlans = [];
let activeAllocation = 0;
let allocationPeople = [];
let allocationSnapshot = '';
let allocationDate = '';
let allocationPreviewSpace = 0;
const selectedPeople = new Set(['呉', '静']);
const selectedOwners = new Set();
const knownPeople = new Set();
const allocationTolerance = 5;
const plannerToday = new Date();
$('#plan-date').value = `${plannerToday.getFullYear()}-${String(plannerToday.getMonth() + 1).padStart(2, '0')}-${String(plannerToday.getDate()).padStart(2, '0')}`;

function plannerMessage(message, error = false) {
  $('#planner-message').textContent = message;
  $('#planner-message').classList.toggle('error', error);
}
function plannerSnapshot() {
  return JSON.stringify(items.map(item => [item.id,item.group,item.name,item.total,item.uploaded,item.date,item.owner,item.parentOnly,item.processed,item.plannedDate,statusOf(item)]));
}
function plannerCandidates(sourceOwners) {
  return items.filter(item => statusOf(item) === 'ready' && item.total > 0 && item.uploaded === 0 && (!sourceOwners.length || sourceOwners.includes(item.owner)));
}
function compareAllocationDates(first, second) {
  for (let index = 0; index < first.dateCounts.length; index++) {
    if (first.dateCounts[index] !== second.dateCounts[index]) return second.dateCounts[index] - first.dateCounts[index];
  }
  return 0;
}
function allocationScore(plan, target) {
  const total = plan.totals.reduce((sum, count) => sum + count, 0);
  const imbalance = Math.max(...plan.totals) - Math.min(...plan.totals);
  return Math.abs(total - target) * 2 + imbalance;
}
function allocationSelectionKey(selection) {
  return JSON.stringify([[...selection.owners].sort(), [...selection.groups].sort()]);
}
function limitAllocationSelections(selections, limit) {
  const groups = new Map();
  for (const selection of selections) {
    const key = allocationSelectionKey(selection);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(selection);
  }
  const ordered = [...groups.values()].sort((first,second) => first[0].owners.size - second[0].owners.size || first[0].groups.size - second[0].groups.size);
  const result = [];
  const priorities = new Set(ordered.map(group => `${group[0].owners.size},${group[0].groups.size}`));
  for (const priority of priorities) {
    const matching = ordered.filter(group => `${group[0].owners.size},${group[0].groups.size}` === priority);
    for (let index = 0; matching.some(group => index < group.length) && result.length < limit; index++) {
      for (const group of matching) {
        if (group[index]) result.push(group[index]);
        if (result.length === limit) return result;
      }
    }
  }
  return result;
}
function generateAllocations(candidates, peopleCount, target) {
  const maxPlans = 20;
  const minimum = Math.max(1, target - allocationTolerance);
  const maximum = target + allocationTolerance;
  const dates = [...new Set(candidates.map(item => item.date || '9999-99-99'))].sort();
  const dateRanks = new Map(dates.map((date,index) => [date,index]));
  const ordered = [...candidates].sort((first,second) => (first.date || '9999-99-99').localeCompare(second.date || '9999-99-99') || first.name.localeCompare(second.name, 'zh-CN'));
  let selections = new Map([['0,0', [{total:0, count:0, dateCounts:Array(dates.length).fill(0), owners:new Set(), groups:new Set(), previous:null, item:null}]]]);
  for (const item of ordered) {
    const next = new Map(selections);
    for (const bucket of selections.values()) {
      for (const selection of bucket) {
        const total = selection.total + item.total;
        if (total > maximum) continue;
        const count = Math.min(peopleCount, selection.count + 1);
        const dateCounts = [...selection.dateCounts]; dateCounts[dateRanks.get(item.date || '9999-99-99')] += item.total;
        const owners = new Set(selection.owners); owners.add(item.owner || '');
        const groups = new Set(selection.groups); groups.add(item.group || '');
        const option = {total, count, dateCounts, owners, groups, previous:selection, item};
        const key = `${total},${count}`;
        const existing = next.get(key) || [];
        const selectionKey = allocationSelectionKey(option);
        const sameSources = existing.filter(entry => allocationSelectionKey(entry) === selectionKey);
        const order = sameSources.length ? compareAllocationDates(option, sameSources[0]) : -1;
        if (order < 0) next.set(key, limitAllocationSelections([...existing.filter(entry => allocationSelectionKey(entry) !== selectionKey), option], maxPlans));
        else if (order === 0) next.set(key, limitAllocationSelections([...existing, option], maxPlans));
      }
    }
    selections = next;
  }
  const ranked = [...selections.values()].flat().filter(selection => selection.count === peopleCount && selection.total >= minimum).sort((first,second) => Math.abs(first.total - target) - Math.abs(second.total - target) || first.owners.size - second.owners.size || first.groups.size - second.groups.size || compareAllocationDates(first,second) || second.total - first.total);
  if (!ranked.length) return [];
  const assignments = [];
  for (const selection of ranked) {
    const selected = [];
    for (let current = selection; current.item; current = current.previous) selected.push(current.item);
    selected.reverse();
    let states = [{totals:Array(peopleCount).fill(0), assignments:[], dateCounts:selection.dateCounts, ownerCount:selection.owners.size, groupCount:selection.groups.size}];
    for (const [index, item] of selected.entries()) {
      const next = [];
      for (const state of states) for (let person = 0; person < peopleCount; person++) {
        const totals = [...state.totals]; totals[person] += item.total;
        if (totals.filter(total => total === 0).length <= selected.length - index - 1) {
          next.push({totals, assignments:[...state.assignments,{item,person}], dateCounts:selection.dateCounts, ownerCount:selection.owners.size, groupCount:selection.groups.size});
        }
      }
      states = next.sort((first,second) => allocationScore(first,target) - allocationScore(second,target)).slice(0, 100);
    }
    assignments.push(...states);
  }
  const results = [];
  assignments.sort((first,second) => Math.abs(first.totals.reduce((sum,count) => sum + count,0) - target) - Math.abs(second.totals.reduce((sum,count) => sum + count,0) - target) || first.ownerCount - second.ownerCount || first.groupCount - second.groupCount || compareAllocationDates(first,second) || allocationScore(first,target) - allocationScore(second,target));
  for (const plan of assignments) {
    if (results.length === 3) break;
    const key = JSON.stringify(plan.assignments.map(entry => [entry.item.id,entry.person]).sort((first,second) => first[0].localeCompare(second[0])));
    if (!results.some(result => result.key === key)) results.push({...plan,key});
  }
  return results;
}
function renderAllocation() {
  const plan = allocationPlans[activeAllocation];
  if (!plan) return;
  document.querySelectorAll('.plan-option').forEach((button,index) => button.setAttribute('aria-pressed', String(index === activeAllocation)));
  const total = plan.totals.reduce((sum,count) => sum + count,0);
  $('#plan-summary').innerHTML = `<span>计划上传日期 <strong>${escapeHTML(allocationDate)}</strong></span><span>合计 <strong>${total.toLocaleString()}</strong> 个文件 / ${plan.assignments.length} 个子文件夹</span>` + plan.totals.map((count,index) => `<span>${escapeHTML(allocationPeople[index])} <strong>${count.toLocaleString()}</strong> 个（${(count / total * 100).toFixed(1)}%）</span>`).join('');
  setPlanPreview(plan, allocationPeople, allocationSnapshot, document.querySelectorAll('.plan-option')[activeAllocation].textContent);
}
function directoryPart(value) {
  const name = value;
  if (!name.trim() || name === '.' || name === '..' || /[<>:"/\\|?*\x00-\x1f]/.test(name) || /[. ]$/.test(name) || /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i.test(name)) throw new Error(`目录名称不符合 Windows 规则：${value}`);
  return name;
}
async function fetchFileLists(sources) {
  try {
    const result = await request('/api/upload-manager/file-lists', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({sources})});
    return Array.isArray(result?.fileLists) ? result.fileLists : [];
  } catch { return []; }
}
function allocationJson(plan, people, paths, label, uploadDate, fileLists = []) {
  return {
    version: 2,
    label,
    uploadDate,
    people: people.map((person, index) => ({
      name: person,
      folders: plan.assignments.flatMap(({item, person:assigned}, assignmentIndex) => assigned === index ? [{
        source: paths[assignmentIndex].slice(1),
        destination: paths[assignmentIndex],
        owner: item.owner || '',
        startDate: item.date || '',
        files: item.total,
        fileList: fileLists[assignmentIndex] || []
      }] : [])
    })).filter(person => person.folders.length)
  };
}
async function copyText(text) {
  try { await navigator.clipboard.writeText(text); return; } catch {}
  const area = document.createElement('textarea');
  area.value = text; area.setAttribute('readonly', ''); area.style.position = 'fixed'; area.style.opacity = '0';
  document.body.append(area); area.select();
  const copied = document.execCommand('copy');
  area.remove();
  if (!copied) throw new Error('浏览器不允许写入剪贴板');
}
const historyManifests = new Map();
function rememberHistory(history) {
  for (const plan of history) if (plan.manifest) historyManifests.set(plan.id, plan.manifest);
}
function historyActions(plan, copyLabel, zipLabel) {
  const copy = plan.manifest ? `<button type="button" class="history-download" data-copy-history="${escapeHTML(plan.id)}"><i data-lucide="clipboard-copy"></i> ${copyLabel}</button>` : '';
  const zip = plan.hasArchive ? `<a class="history-download" href="/api/upload-manager/allocation-history/${encodeURIComponent(plan.id)}/archive"><i data-lucide="download"></i> ${zipLabel}</a>` : '';
  const use = plan.manifest ? `<a class="history-download" href="/distribute?mode=json&amp;plan=${encodeURIComponent(plan.id)}">带入复制页面 →</a>` : '';
  return use + copy + zip;
}
function allocationHistoryMarkup(history) {
  if (!history.length) return '<div class="history-empty">暂无已保存的方案</div>';
  return history.map(plan => {
    const groups = new Map();
    for (const item of plan.assignments) {
      if (!groups.has(item.uploader)) groups.set(item.uploader, []);
      groups.get(item.uploader).push(item);
    }
    const total = plan.assignments.reduce((sum, item) => sum + item.total, 0);
    const people = [...groups].map(([person, assigned]) => `<div class="history-person"><strong>${escapeHTML(person)}</strong> · ${assigned.reduce((sum,item) => sum + item.total,0).toLocaleString()} 个文件<ul>${assigned.map(item => `<li>${escapeHTML(item.group)}${item.parentOnly ? '' : ` / ${escapeHTML(item.name)}`} · ${item.total.toLocaleString()} 个 · 来源：${escapeHTML(item.owner || '未分配')}</li>`).join('')}</ul></div>`).join('');
    return `<details class="history-entry"><summary><strong>${escapeHTML(plan.label)}</strong><span>${escapeHTML(plan.plannedDate)}</span><span>${total.toLocaleString()} 个文件 · ${plan.assignments.length} 个文件夹 · ${groups.size} 人</span><span>${escapeHTML(plan.createdAt.replace('T', ' '))}</span><i data-lucide="chevron-down"></i></summary><div class="history-content"><div class="history-actions">${historyActions(plan, '复制方案 JSON', '下载旧版 ZIP')}</div>${people}</div></details>`;
  }).join('');
}
async function loadAllocationHistory() {
  $('#history-list').innerHTML = '<div class="history-empty">正在加载...</div>';
  try {
    const history = await request('/api/upload-manager/allocation-history');
    rememberHistory(history);
    $('#history-list').innerHTML = allocationHistoryMarkup(history);
    icons();
  } catch (error) {
    $('#history-list').innerHTML = `<div class="history-empty">加载失败：${escapeHTML(error.message)}</div>`;
  }
}
$('#allocation-history').addEventListener('toggle', event => { if (event.currentTarget.open) loadAllocationHistory(); });
$('#history-refresh').addEventListener('click', loadAllocationHistory);
let dateExport = null;
function uploadDateJson(uploadDate, records, fileLists = []) {
  const byPerson = new Map();
  records.forEach((item, index) => {
    const person = item.uploader || '未分配';
    const source = [...item.group.split(/[\\/¥]/).filter(Boolean), ...(item.parentOnly ? [] : [item.name])];
    if (!byPerson.has(person)) byPerson.set(person, []);
    byPerson.get(person).push({source, destination:[person, ...source], owner:item.owner || '', startDate:item.date || '', files:item.total, fileList:fileLists[index] || []});
  });
  return {
    version: 2,
    label: `${uploadDate} 上传记录`,
    uploadDate,
    people: [...byPerson].map(([name, folders]) => ({name, folders}))
  };
}
function uploadDateSources(records) {
  return records.map(item => [...item.group.split(/[\\/¥]/).filter(Boolean), ...(item.parentOnly ? [] : [item.name])]);
}
function dateExportMarkup(uploadDate, records) {
  const title = `<span class="date-history-title"><i data-lucide="history"></i>${escapeHTML(uploadDate)} 上传记录</span>`;
  if (!records.length) return `${title}<span class="date-history-empty">该日期没有符合当前筛选条件的文件夹</span>`;
  const total = records.reduce((sum, item) => sum + item.total, 0);
  return `${title}<span>${records.length} 个文件夹 · ${total.toLocaleString()} 个文件</span><button type="button" class="history-download" data-copy-date-json><i data-lucide="clipboard-copy"></i> 复制当日 JSON</button>`;
}
function updateDateExport(uploadDate, records) {
  const bar = $('#date-history-bar');
  dateExport = uploadDate && !planPreview ? {uploadDate, records} : null;
  bar.hidden = !dateExport;
  if (!dateExport) return;
  bar.innerHTML = dateExportMarkup(uploadDate, records);
  icons();
}
function hideAllocationPreview() {
  $('#plan-hover-preview').hidden = true;
  if (allocationPreviewSpace) {
    const scrollPosition = window.scrollY;
    $('#plan-preview-spacer').style.height = '0px';
    window.scrollTo(window.scrollX, Math.max(0, scrollPosition - allocationPreviewSpace));
    allocationPreviewSpace = 0;
  }
}
function resetAllocationPreviews() {
  hideAllocationPreview();
}
function allocationPersonMarkers() {
  for (const name of [...allocationPeople, ...selectedPeople, ...selectedOwners, ...items.map(item => item.owner)]) if (name) knownPeople.add(name);
  const names = [...knownPeople].sort((first, second) => first.localeCompare(second, 'zh-CN') || (first < second ? -1 : first > second ? 1 : 0));
  return new Map(names.map((name, index) => [name, {number:index + 1, hue:Math.round(index * 360 / names.length)}]));
}
function allocationPersonMarker(name, markers) {
  const marker = markers.get(name);
  if (!marker) return '<span class="plan-person-marker plan-person-unassigned">未分配</span>';
  return `<span class="plan-person-marker" style="--person-hue:${marker.hue}"><span class="plan-person-marker-number" aria-hidden="true">${marker.number}</span><span>${escapeHTML(name)}</span></span>`;
}
function allocationPreviewContent(plan, index) {
  const total = plan.totals.reduce((sum, count) => sum + count, 0);
  const owners = new Set(plan.assignments.map(({item}) => item.owner || ''));
  const groups = new Set(plan.assignments.map(({item}) => item.group || ''));
  const markers = allocationPersonMarkers();
  const people = allocationPeople.map((person, personIndex) => {
    const assigned = plan.assignments.filter(entry => entry.person === personIndex);
    const folders = assigned.map(({item}) => {
      const name = item.parentOnly ? item.group : item.name;
      const details = [`来源：${item.owner || '未分配'}`, `开始：${item.date || '未填写'}`, item.group].filter(Boolean).join(' · ');
      return `<li class="plan-hover-folder"><div class="plan-hover-folder-heading"><strong title="${escapeHTML(name)}">${escapeHTML(name)}</strong><span>${item.total.toLocaleString()} 个文件</span></div><div class="plan-hover-folder-details" title="${escapeHTML(details)}"><span>来源：</span>${allocationPersonMarker(item.owner, markers)}<span>· 开始：${escapeHTML(item.date || '未填写')}</span>${item.group ? `<span>· ${escapeHTML(item.group)}</span>` : ''}</div></li>`;
    }).join('');
    return `<section class="plan-hover-person" style="--person-hue:${markers.get(person).hue}"><div class="plan-hover-person-heading"><strong title="${escapeHTML(person)}">${allocationPersonMarker(person, markers)}</strong><span>${plan.totals[personIndex].toLocaleString()} 个文件 · ${assigned.length} 个文件夹</span></div><ul class="plan-hover-folders">${folders}</ul></section>`;
  }).join('');
  return `<div class="plan-hover-title"><strong>方案 ${index + 1}</strong><span>共 ${total.toLocaleString()} 个文件</span></div><div class="plan-hover-overview">${plan.assignments.length} 个文件夹 · ${owners.size} 个来源负责人 · ${groups.size} 个父文件夹</div><div class="plan-hover-people">${people}</div>`;
}
function showAllocationPreview(button) {
  const index = Number(button.dataset.plan);
  const plan = allocationPlans[index];
  if (!plan) return;
  const preview = $('#plan-hover-preview');
  const headingBounds = $('.planner-result-heading').getBoundingClientRect();
  const buttonBounds = button.getBoundingClientRect();
  preview.innerHTML = allocationPreviewContent(plan, index);
  preview.hidden = false;
  preview.style.top = `${buttonBounds.top - headingBounds.top}px`;
  preview.style.left = `${Math.max(0, Math.min(buttonBounds.left - headingBounds.left, headingBounds.width - preview.offsetWidth))}px`;
  const space = Math.max(0, preview.offsetHeight + 8 - (buttonBounds.top + window.scrollY - allocationPreviewSpace));
  if (space !== allocationPreviewSpace) {
    const scrollPosition = window.scrollY;
    $('#plan-preview-spacer').style.height = `${space}px`;
    window.scrollTo(window.scrollX, Math.max(0, scrollPosition + space - allocationPreviewSpace));
    allocationPreviewSpace = space;
  }
  const currentButtonTop = button.getBoundingClientRect().top;
  if (preview.offsetHeight <= window.innerHeight - button.getBoundingClientRect().height - 16 && currentButtonTop < preview.offsetHeight + 8) {
    window.scrollBy(0, currentButtonTop - preview.offsetHeight - 8);
  }
}
function updatePersonPickerSummary(kind) {
  const selected = kind === 'people' ? selectedPeople : selectedOwners;
  const names = [...selected];
  const markers = allocationPersonMarkers();
  $(`#plan-${kind}-summary`).innerHTML = names.length
    ? names.map(name => `<span class="person-picker-summary-chip${name ? '' : ' person-picker-unassigned'}"${name ? ` style="--person-hue:${markers.get(name).hue}"` : ''} aria-hidden="true">${escapeHTML(name || '未分配')}</span>`).join('') + `<span class="sr-only">${escapeHTML(names.map(name => name || '未分配').join('、'))}</span>`
    : escapeHTML(kind === 'owner' ? '全部负责人' : '选择分配人员');
  $(`#plan-${kind}-count`).textContent = selected.size ? `已选 ${selected.size}` : kind === 'owner' ? '全部' : '未选择';
}
function renderPersonPicker(kind) {
  const selected = kind === 'people' ? selectedPeople : selectedOwners;
  const owners = [...new Set([...items.map(item => item.owner).filter(owner => owner || kind === 'owner'), ...selected])].sort((first,second) => first.localeCompare(second, 'zh-CN'));
  const markers = allocationPersonMarkers();
  $(`#plan-${kind}-options`).innerHTML = owners.length ? owners.map(owner => `<label class="person-picker-option"><input type="checkbox" value="${escapeHTML(owner)}" ${selected.has(owner) ? 'checked' : ''}><span class="person-picker-chip${owner ? '' : ' person-picker-unassigned'}"${owner ? ` style="--person-hue:${markers.get(owner).hue}"` : ''}><span class="person-picker-chip-icon" aria-hidden="true"></span><span class="person-picker-chip-name">${escapeHTML(owner || '未分配')}</span></span></label>`).join('') : '<span class="person-picker-empty">暂无负责人</span>';
  updatePersonPickerSummary(kind);
}
function positionPersonPicker(picker) {
  const panel = picker.querySelector('.person-picker-panel');
  const list = picker.querySelector('.person-picker-list');
  panel.classList.remove('person-picker-above');
  list.style.maxHeight = '';
  const pickerBounds = picker.getBoundingClientRect();
  const panelHeight = panel.getBoundingClientRect().height;
  const listHeight = list.getBoundingClientRect().height;
  const fixedHeight = panelHeight - listHeight;
  const below = window.innerHeight - pickerBounds.bottom - 15;
  const above = pickerBounds.top - 15;
  const opensAbove = panelHeight > below && above > below;
  panel.classList.toggle('person-picker-above', opensAbove);
  list.style.maxHeight = `${Math.max(40, Math.min(listHeight, Math.floor((opensAbove ? above : below) - fixedHeight)))}px`;
}
for (const kind of ['people', 'owner']) {
  const picker = $(`#plan-${kind}-picker`);
  picker.addEventListener('click', event => { if (!picker.open && event.target.closest('summary')) renderPersonPicker(kind); });
  picker.addEventListener('toggle', () => { if (picker.open) positionPersonPicker(picker); });
  $(`#plan-${kind}-options`).addEventListener('change', event => {
    if (event.target.type !== 'checkbox') return;
    const selected = kind === 'people' ? selectedPeople : selectedOwners;
    if (event.target.checked) selected.add(event.target.value); else selected.delete(event.target.value);
    updatePersonPickerSummary(kind);
    if (picker.open) positionPersonPicker(picker);
  });
}
renderPersonPicker('people');
$('#plan-owner-clear').addEventListener('click', () => {
  selectedOwners.clear();
  renderPersonPicker('owner');
  if ($('#plan-owner-picker').open) positionPersonPicker($('#plan-owner-picker'));
  invalidatePlans();
});
function addPlanPerson() {
  const input = $('#plan-person-add');
  const name = input.value.trim();
  if (!name) return;
  const existing = [...selectedPeople].find(person => person.toLocaleLowerCase() === name.toLocaleLowerCase());
  if (!existing) selectedPeople.add(name);
  input.value = '';
  renderPersonPicker('people');
  updatePersonPickerSummary('owner');
  if ($('#plan-people-picker').open) positionPersonPicker($('#plan-people-picker'));
  invalidatePlans();
}
$('#plan-person-add-button').addEventListener('click', addPlanPerson);
$('#plan-person-add').addEventListener('keydown', event => { if (event.key === 'Enter') { event.preventDefault(); addPlanPerson(); } });
document.addEventListener('click', event => {
  for (const kind of ['people', 'owner']) {
    const picker = $(`#plan-${kind}-picker`);
    if (!picker.contains(event.target)) picker.open = false;
  }
});
$('#planner-form').addEventListener('keydown', event => {
  if (event.key === 'Escape') {
    const picker = event.target.closest('.person-picker');
    if (picker?.open) { picker.open = false; picker.querySelector('summary').focus(); }
  }
});
function invalidatePlans() { resetAllocationPreviews(); allocationPlans = []; $('#planner-results').hidden = true; clearPlanPreview(); }
$('#planner-form').addEventListener('input', () => { invalidatePlans(); plannerMessage('从已完成且未上传的子文件夹中，选择目标文件数上下 5 个以内的方案，优先接近目标。'); });
$('#planner-form').addEventListener('change', invalidatePlans);
$('#planner-form').addEventListener('submit', event => {
  event.preventDefault();
  const target = Number($('#plan-target').value);
  const uploadDate = $('#plan-date').value;
  const people = [...selectedPeople];
  if (!uploadDate) { plannerMessage('请选择计划上传日期。',true); return; }
  if (!Number.isInteger(target) || target < 1 || target > 1000000) { plannerMessage('请输入有效的目标文件数。',true); return; }
  if (people.length < 2 || people.length > 5 || new Set(people.map(name => name.toLocaleLowerCase())).size !== people.length) { plannerMessage('请输入 2 至 5 个不同的姓名，用逗号分隔。',true); return; }
  const candidates = plannerCandidates([...selectedOwners]);
  if (candidates.length < people.length) { invalidatePlans(); plannerMessage('符合条件的子文件夹不足，请调整来源条件。',true); return; }
  resetAllocationPreviews();
  allocationPeople = people;
  allocationPlans = generateAllocations(candidates,people.length,target);
  if (!allocationPlans.length) { invalidatePlans(); plannerMessage(`未找到总数在 ${Math.max(1,target - allocationTolerance)} 至 ${target + allocationTolerance} 个文件且每人都有文件夹的方案，请调整目标或来源条件。`,true); return; }
  allocationSnapshot = plannerSnapshot(); allocationDate = uploadDate; activeAllocation = 0;
  $('#plan-options').innerHTML = allocationPlans.map((plan,index) => `<button type="button" class="plan-option" data-plan="${index}" aria-pressed="${index === 0}" aria-describedby="plan-hover-preview">方案 ${index + 1} · ${plan.totals.reduce((sum,count) => sum + count,0).toLocaleString()} 个 · ${plan.ownerCount} 个来源 · ${plan.groupCount} 个父文件夹</button>`).join('');
  $('#planner-results').hidden = false;
  plannerMessage(`从 ${candidates.length} 个已完成且未上传的子文件夹中，生成 ${allocationPlans.length} 个方案（${Math.max(1,target - allocationTolerance)} 至 ${target + allocationTolerance} 个文件），优先接近目标，再考虑来源负责人、父文件夹和开始日期。`);
  renderAllocation();
});
for (const eventName of ['pointerover','focusin']) {
  $('#plan-options').addEventListener(eventName, event => {
    const button = event.target.closest('[data-plan]');
    if (!button) return;
    showAllocationPreview(button);
  });
}
$('#plan-options').addEventListener('click', event => {
  const button = event.target.closest('[data-plan]');
  if (!button) return;
  activeAllocation = Number(button.dataset.plan); renderAllocation();
});
$('.planner-result-heading').addEventListener('pointerleave', hideAllocationPreview);
$('#plan-options').addEventListener('focusout', event => {
  if (!event.currentTarget.contains(event.relatedTarget)) hideAllocationPreview();
});
$('#plan-create').addEventListener('click', async () => {
  const plan = allocationPlans[activeAllocation];
  if (!plan) return;
  const people = [...allocationPeople];
  const snapshot = allocationSnapshot;
  const uploadDate = allocationDate;
  const label = `方案 ${activeAllocation + 1} · ${plan.totals.reduce((sum,count) => sum + count,0)} 个`;
  if (snapshot !== plannerSnapshot()) { plannerMessage('文件夹数据已更改，请重新生成方案。',true); return; }
  if (!await askConfirm(`确认后保存 ${uploadDate} 的上传计划，状态设为“上传中”，不会标记为已上传，并把方案 JSON 复制到剪贴板。实际上传完成后请手动设为“已上传”。继续吗？`, {title:'确认上传计划',okText:'确认并保存'})) return;
  if (snapshot !== plannerSnapshot()) { plannerMessage('文件夹数据已更改，请重新生成方案。',true); return; }
  let paths;
  try {
    paths = plan.assignments.map(({item,person}) => [directoryPart(people[person]), ...item.group.split(/[\\/¥]/).map(directoryPart), ...(item.parentOnly ? [] : [directoryPart(item.name)])]);
    const keys = paths.map(path => path.map(part => part.toLocaleLowerCase()).join('/'));
    if (new Set(keys).size !== keys.length) throw new Error('方案中存在重复目录，请先调整文件夹名称。');
  } catch (error) { plannerMessage(error.message,true); return; }
  const createButton = $('#plan-create');
  createButton.disabled = true;
  let marked = false;
  try {
    const fileLists = await fetchFileLists(paths.map(path => path.slice(1)));
    const manifest = allocationJson(plan, people, paths, label, uploadDate, fileLists);
    const result = await request('/api/upload-manager/allocations/confirm', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({label,manifest,assignments:plan.assignments.map(({item,person}) => ({id:item.id,version:item.version,uploader:people[person],uploadDate}))})});
    const saved = result.items;
    const updated = new Map(saved.map(item => [item.id,item]));
    items = items.map(item => updated.get(item.id) || item);
    marked = true;
    const next = document.createElement('a');
    next.className = 'history-download';
    next.href = '/distribute?mode=json&plan=' + encodeURIComponent(result.historyId);
    next.textContent = '带入复制页面 →';
    document.getElementById('manager-next-copy')?.remove();
    next.id = 'manager-next-copy';
    $('#planner-message').after(next);
    if ($('#allocation-history').open) loadAllocationHistory();
    await copyText(JSON.stringify(manifest, null, 2));
    invalidatePlans();
    render();
    plannerMessage(`已为 ${paths.length} 个文件夹保存上传计划，方案 JSON 已复制到剪贴板；文件尚未标记为已上传。`);
    notify('方案 JSON 已复制到剪贴板');
  } catch (error) {
    if (marked) { invalidatePlans(); render(); }
    plannerMessage(marked ? `已保存上传计划，但复制 JSON 失败：${error.message}。可直接点击「带入复制页面」继续，或在“历史分配方案”中重新复制。` : `分配确认失败，未保存上传计划：${error.message}`, !marked);
  } finally {
    createButton.disabled = false;
  }
});
document.addEventListener('click', async event => {
  const button = event.target.closest('[data-copy-date-json]');
  if (!button || !dateExport) return;
  try { await copyText(JSON.stringify(uploadDateJson(dateExport.uploadDate, dateExport.records, await fetchFileLists(uploadDateSources(dateExport.records))), null, 2)); notify('当日 JSON 已复制到剪贴板'); }
  catch (error) { notify(`复制失败：${error.message}`); }
});
document.addEventListener('click', async event => {
  const button = event.target.closest('[data-copy-history]');
  const manifest = button && historyManifests.get(button.dataset.copyHistory);
  if (!manifest) return;
  try { await copyText(JSON.stringify(manifest, null, 2)); notify('方案 JSON 已复制到剪贴板'); }
  catch (error) { notify(`复制失败：${error.message}`); }
});
