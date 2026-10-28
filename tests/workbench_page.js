const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '..', 'static', 'batch.js'), 'utf8');

// A small structural DOM for state/identity assertions. Browser layout is tested separately.
function element(tag = 'div') {
  return {
    tag, value: '', dataset: {}, style: {}, attributes: {}, children: [], parentElement: null,
    _html: '', _raw: null, className: '', listeners: {}, _queries: new Map(),
    classList: { values: new Set(), toggle(name, value) { value ? this.values.add(name) : this.values.delete(name); },
      add(name) { this.values.add(name); }, remove(name) { this.values.delete(name); }, contains(name) { return this.values.has(name); } },
    get innerHTML() { return this._html + this.children.map(child => child.outerHTML).join(''); },
    set innerHTML(value) { this.children.forEach(child => { child.parentElement = null; }); this.children = []; this._html = value; this._queries.clear(); },
    get outerHTML() { return this._raw ?? `<${this.tag} class="${this.className}">${this.innerHTML}</${this.tag}>`; },
    get firstElementChild() {
      if (!this.children.length && this._html.trim().startsWith('<')) {
        const node = element(this._html.match(/^\s*<([\w-]+)/)[1]);
        node._raw = this._html;
        const opening = this._html.match(/^\s*<[^>]+>/)[0];
        node.className = opening.match(/class="([^"]*)"/)?.[1] || '';
        for (const match of opening.matchAll(/data-([\w-]+)="([^"]*)"/g)) {
          node.dataset[match[1].replace(/-([a-z])/g, (_, c) => c.toUpperCase())] = match[2];
        }
        this._html = '';
        this.append(node);
      }
      return this.children[0] || null;
    },
    get nextElementSibling() { const siblings = this.parentElement?.children || []; return siblings[siblings.indexOf(this) + 1] || null; },
    get isConnected() { return !!this.root || !!this.parentElement?.isConnected; },
    insertBefore(node, before) {
      node.remove();
      const index = before ? this.children.indexOf(before) : this.children.length;
      this.children.splice(index, 0, node);
      node.parentElement = this;
    },
    append(...nodes) { nodes.forEach(node => this.insertBefore(node, null)); },
    remove() { if (this.parentElement) { const siblings = this.parentElement.children; siblings.splice(siblings.indexOf(this), 1); this.parentElement = null; } },
    replaceWith(node) { if (this.parentElement) { this.parentElement.insertBefore(node, this); this.remove(); } },
    replaceChildren(...nodes) { this.innerHTML = ''; this.append(...nodes); },
    setAttribute(name, value) { this.attributes[name] = value; },
    getAttribute(name) { return this.attributes[name]; },
    addEventListener(name, fn) { this.listeners[name] = fn; },
    showModal() { this.open = true; }, close() { this.open = false; },
    scrollIntoView() {}, focus() {},
    getBoundingClientRect() { return { top: 100, bottom: 600, height: this.height || 500 }; },
    querySelectorAll(selector) {
      const matches = node => (selector.includes('.batch-card') && node.className === 'batch-card')
        || (selector === '.batch-file' && node.className === 'batch-file');
      return this.children.flatMap(child => [...(matches(child) ? [child] : []), ...child.querySelectorAll(selector)]);
    },
    querySelector(selector) {
      if (selector.includes('.batch-card')) { this.firstElementChild; return this.querySelectorAll(selector)[0] || null; }
      if (!this._queries.has(selector)) this._queries.set(selector, element('button'));
      return this._queries.get(selector);
    },
  };
}

function page({ collapsed = false } = {}) {
  const elements = new Map(), requests = [], listeners = {}, frames = [];
  const context = vm.createContext({
    document: { activeElement: null, createElement: element,
      getElementById(id) { if (!elements.has(id)) elements.set(id, Object.assign(element(), { root: true })); return elements.get(id); },
      addEventListener(name, fn) { listeners[name] = fn; }, querySelectorAll() { return []; },
      querySelector() { return null; }, elementFromPoint() { return this.pointElement || null; } },
    localStorage: { getItem() { return ''; } }, window: { addEventListener() {}, scrollBy() {} }, appConfirm: async () => true,
    setInterval() {}, setTimeout, clearTimeout, requestAnimationFrame(fn) { frames.push(fn); return frames.length; },
    fetch() { return Promise.resolve({ status: 304 }); },
    MutationObserver: class { observe() {} }, innerHeight: 900, innerWidth: 1200,
    excelLinkHtml(label) { return `<a>${label}</a>`; }, Date, requests,
  });
  vm.runInContext(source, context);
  const run = code => vm.runInContext(code, context);
  run(`const pollUnderTest = poll, postUnderTest = post; ++pollVersion;
    batchState = {busy: false, job: null, logs: [], progress: {}, folder: 'C:/Excel', cards: [
    {id:'a', seq:1, wb_name:'book.xlsx', sheet_name:'Sheet', cell_addr:'A5', color:'blue', verified:true, images:{before:'/before.png', after:'/after.png'}},
    {id:'b', seq:2, wb_name:'book.xlsx', sheet_name:'Sheet', cell_addr:'A6', color:'yellow', finalized:true, images:{}}
  ]}; poll = async () => {}; post = async (url, payload) => {requests.push({url, payload}); return {ok:true};};`);
  // 既有卡片交互用例按全部展开验证；折叠/懒加载用例明确使用真实默认状态。
  if (!collapsed) run(`const expandedRenderUnderTest = render; render = () => {
    batchState.cards.forEach(c => { expandedFiles.add(fileKey(c)); fileMountLimits.set(fileKey(c), Infinity); });
    expandedRenderUnderTest();
  };`);
  return { run, elements, requests, listeners, frames };
}

module.exports = { page };
