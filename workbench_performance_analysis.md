# `/workbench` 大量卡片场景下的卡顿代码分析与优化设计方案

## 一、问题现象与背景

在切字检查工作台（`/workbench`）处理大批量 Excel 文件、待审核卡片数量增长到数百至上千张时，页面会出现明显的交互延迟、滚动掉帧、点击切色响应变慢，甚至在后台持续扫描/试修期间浏览器主线程周期性冻结。

前端通过 `setInterval(poll, 1800)` 每 1.8 秒向后端请求一次 `/api/batch`（对应 [`core/batch.py`](./core/batch.py) 中的 `snapshot()`），并在每次收到响应或用户交互时执行 [`static/batch.js`](./static/batch.js) 中的 `render()`。经过对前后端全链路代码的逐行排查，卡顿并非单一原因造成，而是由 **后端锁内高频磁盘 I/O 与 $O(N \times M)$ 全表扫描** + **前端每次新增卡片触发整页 DOM 销毁重建、强制同步重排（Layout Thrashing）及多处 $O(N^2)$ 数组遍历** 叠加放大所致。

---

## 二、核心性能瓶颈代码级剖析

### 瓶颈 1（后端）：`snapshot()` 在全局锁 `STATE_LOCK` 内执行数千次磁盘 `stat()` 系统调用

#### 1.1 问题代码位置
* [`core/batch.py`](./core/batch.py) 第 400–455 行（`_image_url` 与 `snapshot`）：
  ```python
  def _image_url(name: str) -> str:
      # 同名文件会被后面的尝试覆盖：带版本号，内容变了地址才变，浏览器可以放心缓存
      return f"/api/batch/image/{name}?v={fix_images.version(name)}"
  ```
  ```python
  result["images"] = {
      which: _image_url(name) for which, name in card.get("fix_images", {}).items()
      if fix_images.image_path(name)
  }
  if result.get("context"):
      context = result["context"]
      name = context["image_file"]
      context["image"] = _image_url(name) if fix_images.image_path(name) else ""
  ```
* [`core/fix_images.py`](./core/fix_images.py) 第 50–62 行（`image_path` 与 `version`）：
  ```python
  def image_path(name: str) -> Path | None:
      if not name or Path(name).name != name:
          return None
      path = IMAGES_DIR / name
      return path if path.is_file() else None  # <-- 第 1 次磁盘 stat()

  def version(name: str) -> str:
      path = image_path(name)                  # <-- 第 2 次磁盘 stat()
      if path is None:
          return ""
      stat = path.stat()                       # <-- 第 3 次磁盘 stat()
      return f"{stat.st_mtime_ns:x}{stat.st_size:x}"
  ```
* [`core/batch.py`](./core/batch.py) 第 1940–1955 行（`_auto_retry_blocked`）：
  ```python
  def _auto_retry_blocked(card: dict) -> bool:
      path = card["wb_path"]
      try:
          current = wf._file_version(Path(path))  # <-- 每张已提交黄卡对原 Excel 做 1 次磁盘 stat()
  ```

#### 1.2 复杂度与耗时量化
* 每张卡片最多包含 3 张图片（`before`、`after`、`context`）。每生成 1 张图片的 URL，会调用 1 次 `fix_images.image_path(name)` + 1 次 `fix_images.version(name)`，底层触发 **3 次 `os.stat()`**，即 **每张卡片最多触发 9 次磁盘 `stat()`**。
* 当页面有 **600 张卡片** 时，前端每 **1.8 秒** 轮询一次 `/api/batch`，后端在持有全局 `wf.STATE_LOCK` 的情况下要执行 **$600 \times 9 + 600 = 6,000$ 次 Windows NTFS 文件系统 `stat()` 调用**。
* Windows 下单次 `stat()` 耗时约 $0.05\text{ms} \sim 0.15\text{ms}$，6,000 次 `stat()` 耗时高达 **$300\text{ms} \sim 900\text{ms}$**，不仅让接口响应变慢，还会在锁内阻塞后台扫描/修复线程和前端切色请求。

---

### 瓶颈 2（后端）：`snapshot()` 在全局锁内存在隐藏的 $O(N \times M)$ 全表扫描与全量 `deepcopy`

#### 2.1 问题代码位置
* [`core/batch.py`](./core/batch.py) 第 407–436 行：
  ```python
  by_cell: dict[tuple[str, str, str], list[dict]] = {}
  for item in wf.STATE["items"].values():
      by_cell.setdefault(_cell(item), []).append(item)
  for card in sorted(_data()["cards"].values(), key=lambda c: c["seq"]):
      if card.get("dismissed") or not wf.in_current_folder(card["wb_path"]):
          continue
      result = copy.deepcopy(card)
      ...
      if card["color"] == "yellow" and card.get("finalized"):
          result["auto_skip"] = _auto_skip_reason(card, by_cell.get(_cell(card), []))
  ```
* 虽然 `snapshot()` 开头特意构建了 `by_cell` 字典并传给了 `_auto_skip_reason(card, items)`，但看 `_auto_skip_reason`（第 1957 行）内部调用的 `_can_highlight(card)`（第 2229 行）：
  ```python
  def _can_highlight(card: dict) -> bool:
      return bool(card.get("cell_addr")) and all(
          it.get("location_confirmed") is not False for it in _items_at(card)
      )

  def _items_at(card: dict) -> list[dict]:
      return [it for it in wf.STATE["items"].values() if _cell(it) == _cell(card)]
  ```

#### 2.2 复杂度与耗时量化
* `_can_highlight(card)` 没有接收已经查好的 `items`，而是每次重新调用 `_items_at(card)` 遍历全量 `wf.STATE["items"].values()`。
* 此外，`_handed_over_paths()`（由 `_mode()` 在每次 `snapshot()` 中调用，第 1777 行）对每张 `handed_over` 卡片调用 `_unchecked_fix_reason(c)` -> `_card_fixable_types(c)` -> `_items_at(c)`，**又遍历了一遍全量 `wf.STATE["items"].values()`**。
* 当有 800 张黄卡、1,000 个检测项 `items` 时，单次 `snapshot()` 在 `STATE_LOCK` 内会空跑 **$800 \times 1,000 \times 2 = 1.6 \times 10^6$ 次字典比较循环**。
* 同时，对每张卡片执行 `copy.deepcopy(card)` 会深拷贝 `signatures`、`problems`、`item_ids`、`context`、`fix_images` 等所有嵌套结构，产生大量短生命周期 Python 对象与 GC 压力。

---

### 瓶颈 3（前端）：每当新增卡片或布局变化时，整页 `innerHTML` 销毁重建 + 强制同步重排（Layout Thrashing）

#### 3.1 问题代码位置
* [`static/batch.js`](./static/batch.js) 第 750–771 行（`render()` 瀑布流分支）：
  ```javascript
  const groups = fileGroups(shown);
  const fresh = placeCards(groups, columnCount);
  const html = new Map(shown.map(c => [c.id, cardHtml(c)]));
  const heads = new Map(groups.map(([key]) => [key, fileHeadHtml(key)]));
  let layout = layoutOf(groups, columnCount);
  if (layout === layoutKey && shown.length) {
      // 仅更新有变化的单卡 outerHTML ...
  } else {
      const anchor = scrollAnchor();
      $('cardList').innerHTML = shown.length
        ? `<div class="batch-files">${groups.map(...).join('')}</div>`
        : ...;
      if (rebalanceFresh(groups, columnCount, fresh)) layout = layoutOf(groups, columnCount);
      restoreAnchor(anchor);
  }
  ```
* [`static/batch.js`](./static/batch.js) 第 619–648 行（`rebalanceFresh`）：
  ```javascript
  function rebalanceFresh(groups, count, fresh) {
    const fileElements = $('cardList').querySelectorAll('.batch-file');
    if (fileElements.length !== groups.length) return false;
    const elements = new Map();
    for (const element of $('cardList').querySelectorAll('.batch-card[data-id]')) {
      elements.set(element.dataset.id, element);
      cardHeights.set(element.dataset.id, element.getBoundingClientRect().height + CARD_GAP);
    }
    ...
  ```

#### 3.2 为什么会导致严重卡顿
1. **增量扫描/修复期间每 1.8 秒触发一次整页销毁重建**：
   * 在扫描或修复过程中，只要新检出 1 张卡片，或 `batchState.busy` 状态改变，`layout = layoutOf(groups, columnCount)` 就会与 `layoutKey` 不同，从而进入 `else` 分支。
   * `$('cardList').innerHTML = ...` 会**一次性销毁页面上已有的几百上千张卡片 DOM 节点及数千个 `<img>` 元素**，再将数十万字符的 HTML 字符串重新解析生成几万个新 DOM 节点。
2. **强制同步布局（Forced Synchronous Layout / Reflow）**：
   * 刚设置完巨大的 `$('cardList').innerHTML`，浏览器尚未完成排版，下一行立刻调用 `rebalanceFresh()`，对页面上**每一张卡片**执行 `element.getBoundingClientRect().height`。
   * 这会强迫浏览器主线程立即同步计算整棵几万个节点的样式与几何布局，随后又通过 `columnElements[i].append(...)` 移动节点，紧接着 `restoreAnchor()` 再次读取 `getBoundingClientRect()`，造成连续多次全量重排。
3. **屏幕外大量节点未启用渲染隔离**：
   * [`static/batch.css`](./static/batch.css) 中的 `.batch-file` 与 `.batch-card` 未设置 `content-visibility: auto`，浏览器在每次滚动和窗口尺寸计算时，都必须对视口外几万像素高的所有卡片进行布局和绘制计算。

---

### 瓶颈 4（前端）：每次轮询与交互中存在多处 $O(N^2)$ 嵌套遍历与无变化时的全页 DOM 扫描

#### 4.1 问题代码位置
1. **每次轮询即使无任何变化（`renderKey === key`），仍全量扫描 DOM 与数组**：
   * [`static/batch.js`](./static/batch.js) 第 733 行：
     ```javascript
     if (renderKey === key) { applySelection(); return; }
     ```
   * 而 `applySelection()`（第 486–501 行）会：
     * `document.querySelectorAll('.batch-card[data-id]')` 遍历全页所有卡片 DOM；
     * 对每个文件框按钮，执行 `shown.filter(c => fileKey(c) === button.dataset.file)`（$O(F \times N)$ 遍历）。
2. **`renderControls()` 中的 $O(N^2)$ 数组 `includes` 查找**（第 242–243 行）：
   ```javascript
   const manual = manualCards(), recheckAll = recheckAllCards();
   const held = new Set(manual.filter(c => !recheckAll.includes(c)).map(fileKey)).size;
   ```
   当 `manual` 有 800 张黄卡时，`manual.filter(c => !recheckAll.includes(c))` 会进行 $800 \times 800 = 640,000$ 次引用比较。
3. **`fileHeadHtml(key)` 中的重复全表扫描**（第 653–689 行）：
   每渲染一个文件框标题栏（假设有 $F$ 个文件、共 $N$ 张卡）：
   * `const cards = batchState.cards.filter(c => fileKey(c) === key)` 扫描全表 $N$ 次；
   * 调用 `autoQueueBadge(key)`（第 80 行）：内部 `batchState.cards.some(c => fileKey(c) === path && queued.has(c.id))` 再次扫描全表 $N$ 次；
   * 调用 `followupBadge(key)`（第 70 行）：内部对每个排队任务都执行 `batchState.cards.some(c => fileKey(c) === path && ids.has(c.id))` 扫描全表 $N$ 次。
   * 渲染全部文件标题栏的总复杂度达到 $O(F \times N)$。
4. **`cardHeight(id)` 在 `placeCards()` 中的 $O(N^2)$ 求和**（第 584–589 行）：
   ```javascript
   function cardHeight(id) {
     if (cardHeights.has(id)) return cardHeights.get(id);
     let sum = 0;
     for (const value of cardHeights.values()) sum += value;
     return cardHeights.size ? sum / cardHeights.size : 520;
   }
   ```
   在 `placeCards` 循环放置新卡时，每张尚未测量高度的卡片都会把 `cardHeights.values()` 从头累加一遍。
5. **`renderKey` 每 1.8 秒对全量卡片对象做 `JSON.stringify`**（第 730 行）：
   ```javascript
   const key = JSON.stringify([batchState.cards, ...]);
   ```
   当 `batchState.cards` 有上千张卡片（含长文本、`problems`、`signatures`、`context`）时，每次轮询在主线程序列化数 MB 的 JSON 字符串，而且后端下发的 `card` 还携带了前端完全用不到的内部字段（如 `signatures`、`item_ids`）。

---

## 三、优化设计方案

本方案遵循 **「零功能变更、零视觉变更、前后端协同降维」** 的原则，分 4 个模块进行改造。

---

### 方案模块 1：后端内存级图片版本缓存（消除 `snapshot()` 中 99% 的磁盘 `stat()`）

#### 1.1 设计思路 ([`core/fix_images.py`](./core/fix_images.py))
* 所有报告/审核图片的创建、覆盖与删除全部经由 [`core/fix_images.py`](./core/fix_images.py) 的 `save()`、`forget()`、`clear()`、`import_all()` 四个入口。
* 在 [`core/fix_images.py`](./core/fix_images.py) 中维护一个内存字典 `_VERSION_CACHE: dict[str, str | None]`（`name -> version_str`，文件不存在时存 `None`）：
  1. **`save(item, which, data_url, ...)`**：写入字节成功时，直接更新 `_VERSION_CACHE[name]`。
  2. **`image_path(name)` 与 `version(name)`**：优先查 `_VERSION_CACHE`；仅当缓存未命中（如进程重启后首次访问旧图片）时才执行一次 `path.stat()` 并写入 `_VERSION_CACHE`。
  3. **`clear()` 与 `import_all()`**：同步清理或更新 `_VERSION_CACHE`。

#### 1.2 代码设计示例 ([`core/fix_images.py`](./core/fix_images.py))
```python
# 内存缓存：文件名 -> 版本号字符串（文件不存在记为 None），避免 snapshot 轮询时高频 stat 磁盘
_VERSION_CACHE: dict[str, str | None] = {}
_WRITE_SEQ = 0


def _probe_version(name: str) -> str | None:
    if not name or Path(name).name != name:
        return None
    if name in _VERSION_CACHE:
        return _VERSION_CACHE[name]
    path = IMAGES_DIR / name
    try:
        st = path.stat()
    except OSError:
        _VERSION_CACHE[name] = None
        return None
    ver = f"{st.st_mtime_ns:x}{st.st_size:x}"
    _VERSION_CACHE[name] = ver
    return ver


def has_image(name: str) -> bool:
    return _probe_version(name) is not None


def image_path(name: str) -> Path | None:
    return (IMAGES_DIR / name) if _probe_version(name) is not None else None


def version(name: str) -> str:
    return _probe_version(name) or ""
```
在 `save()` 写入文件成功后直接更新缓存：
```python
    global _WRITE_SEQ
    _WRITE_SEQ += 1
    raw = base64.b64decode(data_url.split(",", 1)[1])
    (IMAGES_DIR / name).write_bytes(raw)
    st = (IMAGES_DIR / name).stat()
    _VERSION_CACHE[name] = f"{st.st_mtime_ns:x}{st.st_size:x}_{_WRITE_SEQ:x}"
```
> **效果**：`snapshot()` 中每张卡片的 9 次磁盘 `stat()` 直接降为 **0 次磁盘 I/O**（纯内存字典 $O(1)$ 查找）。

---

### 方案模块 2：后端 `snapshot()` 消除 $O(N \times M)$ 全表扫描、缓存文件版本并精简序列化

#### 2.1 设计思路 ([`core/batch.py`](./core/batch.py))
1. **复用 `by_cell` 字典**：
   * 修改 `_can_highlight(card, items: list[dict] | None = None)`，当调用方传入 `items` 时直接使用，不再调用 `_items_at(card)` 全表扫描 `wf.STATE["items"]`。
   * 在 `_auto_skip_reason(card, items, enabled, file_versions)` 中将 `items` 透传给 `_can_highlight(card, items)`。
   * 在 `_handed_over_paths(by_cell)` 中同样复用 `by_cell`。
2. **单次 `snapshot()` 内按文件 `wb_path` 缓存 `_file_version`**：
   * 同一文件可能有几十张黄卡，在 `snapshot()` 局部变量中用 `file_versions: dict[str, Any] = {}` 缓存每个 `wb_path` 的 `wf._file_version(Path(path))`，每个 Excel 文件最多只 `stat` 一次，而不是每张卡片 `stat` 一次。
3. **用轻量字典浅拷贝替代全量 `copy.deepcopy(card)`**：
   * 在 `snapshot()` 中，无需深拷贝 `signatures`、`item_ids` 等前端不使用的列表，只构造前端展示所需的字段（对 `problems`、`context` 做浅拷贝即可），大幅降低 CPU 与 JSON 体积。

#### 2.2 代码设计示例 ([`core/batch.py`](./core/batch.py))
```python
def _can_highlight(card: dict, items: list[dict] | None = None) -> bool:
    items = _items_at(card) if items is None else items
    return bool(card.get("cell_addr")) and all(it.get("location_confirmed") is not False for it in items)


def _auto_retry_blocked(card: dict, file_versions: dict[str, Any] | None = None) -> bool:
    path = card["wb_path"]
    if file_versions is not None and path in file_versions:
        current = file_versions[path]
        if current is None:
            return False
    else:
        try:
            current = wf._file_version(Path(path))
        except OSError:
            if file_versions is not None:
                file_versions[path] = None
            return False
        if file_versions is not None:
            file_versions[path] = current
    done = _data().get("auto_versions", {}).get(path)
    return done == current if done is not None else _unchanged_since_check(path)
```

---

### 方案模块 3：前端增量 DOM 渲染与按需高度测量（彻底消除整页重建与强制同步重排）

#### 3.1 设计思路 ([`static/batch.js`](./static/batch.js))
当前最大的前端瓶颈是：只要 `layout !== layoutKey`（例如扫描中新增了 1 张卡片），就把整个 `$('cardList').innerHTML` 推倒重来，并对所有卡片调 `getBoundingClientRect()`。

我们可以把瀑布流渲染改为 **细粒度增量 DOM 更新**：
1. **文件框级增量复用**：
   * 如果 `#cardList > .batch-files` 容器已存在，且现有文件框的顺序是新 `groups` 的前缀（或相同文件集合）：
     * 对已存在的文件框 `<section class="batch-file" data-file="...">`：不销毁它，只更新有变化的 `.batch-file-head`；
     * 对新增的文件框：使用 `insertAdjacentHTML('beforeend', ...)` 追加到 `.batch-files` 末尾；
     * 对已移除的文件框（例如筛选切换或全部提交后）：直接 `section.remove()`。
2. **列内卡片级增量更新**：
   * 在每个文件框的 1~3 个 `.batch-column` 中：
     * 如果该列已有卡片是目标卡片列表的前缀（最常见场景：扫描时往列尾追加新卡片，或卡片只改了颜色/文字）：
       * 已存在的卡片：仅当 `cardHtmlCache.get(id) !== html.get(id)` 时替换该卡的 `outerHTML`；
       * 新增的卡片：直接 `columnEl.insertAdjacentHTML('beforeend', html.get(id))` 追加到该列末尾！
     * 只有当列数变化（窗口缩放）或筛选切换导致列内顺序重排时，才重绘对应的列。
3. **只测量未记录高度的新卡片（消除全量 `getBoundingClientRect`）**：
   * 在 `rebalanceFresh(groups, columnCount, fresh)` 中：
     * 之前已经测过高度并存在 `cardHeights` 里的卡片，**不再调用 `getBoundingClientRect()`**；
     * 只对本次 `fresh` 集合中的新卡片调用 `getBoundingClientRect().height`！
   * 维护全局变量 `cardHeightSum` 动态累加，让 `cardHeight(id)` 计算平均高度从 $O(N)$ 变为 **$O(1)$**。

#### 3.2 代码设计示例 ([`static/batch.js`](./static/batch.js))
```javascript
let cardHeightSum = 0;

function rememberCardHeight(id, h) {
  if (cardHeights.has(id)) cardHeightSum -= cardHeights.get(id);
  cardHeights.set(id, h);
  cardHeightSum += h;
}

function cardHeight(id) {
  if (cardHeights.has(id)) return cardHeights.get(id);
  return cardHeights.size ? cardHeightSum / cardHeights.size : 520;
}
```
在 `rebalanceFresh` 中只测量 `fresh` 卡片：
```javascript
function rebalanceFresh(groups, count, fresh) {
  if (!fresh.size) return false;
  const fileElements = $('cardList').querySelectorAll('.batch-file');
  if (fileElements.length !== groups.length) return false;
  const elements = new Map();
  for (const id of fresh) {
    const element = $('cardList').querySelector(`.batch-card[data-id="${cssEscape(id)}"]`);
    if (element) {
      elements.set(id, element);
      rememberCardHeight(id, element.getBoundingClientRect().height + CARD_GAP);
    }
  }
  ...
```

---

### 方案模块 4：消除前端所有 $O(N^2)$ 循环 + 启用 CSS 视口外跳过渲染 (`content-visibility`)

#### 4.1 消除 $O(N^2)$ 数组遍历 ([`static/batch.js`](./static/batch.js))
1. **`renderControls()` 中的 `recheckAll` 改用 `Set`**：
   ```javascript
   const manual = manualCards(), recheckAll = recheckAllCards();
   const recheckSet = new Set(recheckAll);
   const held = new Set(manual.filter(c => !recheckSet.has(c)).map(fileKey)).size;
   ```
2. **`fileHeadHtml(key, fileCards)` 直接传入该文件已分组好的 `cards`**：
   * 在 `render()` 中已经有 `groups = fileGroups(shown)`，同时预先构建一次全量按文件分组映射 `allByFile = new Map(fileGroups(batchState.cards))`。
   * `fileHeadHtml(key, allByFile.get(key) || [])`、`autoQueueBadge(key, fileCards)`、`followupBadge(key, fileCards)` 直接只检查该文件自己的 `fileCards`（通常几张到几十张），不再每次去遍历全量 `batchState.cards`。
3. **无变化时跳过 `applySelection()` 的全页 DOM 遍历**：
   * 记录 `lastSelectionKey = ''`（由 `[...selectedIds]` 和 `renderKey` 组成）。
   * 当 `renderKey === key` 且 `selectionKey === lastSelectionKey` 时，`render()` 直接 `return`，完全不触碰 DOM。
   * 在 `applySelection()` 中，预先把 `shownCards()` 按 `fileKey` 分组存入 `Map`，避免在 `for (const button of ...)` 循环内反复 `shown.filter(...)`。

#### 4.2 CSS 开启 `content-visibility: auto` ([`static/batch.css`](./static/batch.css))
在 [`static/batch.css`](./static/batch.css) 中为文件分组框 `.batch-file` 添加浏览器原生虚拟化属性：
```css
.batch-file {
  content-visibility: auto;
  contain-intrinsic-size: auto 620px;
}
```
> **作用**：浏览器会自动跳过当前滚动视口以外的所有 `.batch-file` 区块的子树样式计算、布局（Layout）与绘制（Paint），只渲染用户当前屏幕看到的几个文件框，内存与渲染帧率直接提升一个数量级，且完全不影响页内搜索、锚点恢复和瀑布流高度。

---

## 四、优化前后性能对比预估（以 80 个文件、1,000 张卡片为例）

| 指标 | 优化前 | 优化后 |
|---|---|---|
| **后端单次 `/api/batch` 磁盘 `stat()` 次数** | ~9,000 次（耗时 400~1000ms，持锁阻塞） | **0~80 次**（图片缓存命中后 0 次图片 `stat`，耗时 < 10ms） |
| **后端 `snapshot()` 内部循环次数** | $O(N \times M) \approx 1.6 \times 10^6$ 次 | **$O(N + M) \approx 2,000$ 次** |
| **前端无变化轮询时的 DOM 操作** | 每 1.8s 遍历 1,000 个卡片 DOM + $O(F \times N)$ 过滤 | **0 次 DOM 操作**（直接短路返回） |
| **前端扫描新增 1 张卡片时的 DOM 操作** | 销毁并重建全部 1,000 张卡片 DOM + 测量 1,000 张卡高度 | **仅追加 1 张新卡片 DOM + 仅测量 1 张新卡高度** |
| **前端视口外卡片渲染开销** | 全量 1,000 张卡片实时参与 Layout / Paint | **`content-visibility: auto` 仅渲染视口内 10~20 张卡片** |

---

## 五、附：单元格错误类型分级与 `###` 列宽自动调整方案（已实现）

### 5.1 6 种单元格错误类型统一规范 ([`core/autofix.py`](./core/autofix.py))
在 `CELL_ERROR_SPECS` 与 `item_error_type(item)` 中统一维护 6 种错误类型：

| 错误类型编码 (`code`) | 名称 (`label`) | 是否可自动修复 (`fixable`) | 自动修复策略 |
|---|---|---|---|
| `border_clip` | **边框压线** | ✅ 可勾选（默认选） | 自动逐次增加行高（每次 `+3pt`，通过后多留一步余量） |
| `text_truncated` | **文字截断** | ✅ 可勾选（默认选） | 未换行且横向受阻时先开启 `WrapText = True`，再逐次增加行高 |
| `box_overflow` | **超出裁剪框** | ✅ 可勾选（默认选） | 未换行且横向受阻时先开启 `WrapText = True`，再逐次增加行高 |
| `page_break_cut` | **分页切开合并单元格** | ✅ 可勾选（默认选） | 在合并区域第一行上方插入手动水平分页符 `XL_PAGEBREAK_MANUAL` |
| `hash_overflow` | **`###` 列宽不足** | ✅ 可勾选（默认选） | 单列直接加宽该列；跨多列合并时平均微调合并区域各列，直到 `###` 消失并留余量 |
| `text_missing` | **PDF 未匹配到文字** | ⚪ 固定转人工 | 缺少 PDF 定位且无法排除作者有意隐藏，整行不自动加高，直接转黄色人工卡 |

### 5.2 `###` 列宽不足（`hash_overflow`）自动调整与还原机制 ([`core/batch.py`](./core/batch.py))
1. **为什么数值/日期 `###` 需调列宽而非自动换行**：
   * 普通文本横向放不下时，可通过 `WrapText = True` + 增加行高折行显示（`_wrap_targets`）；
   * 但 Excel 对数字、日期类型的 `###` 溢出不会折行，即使开启 `WrapText` 并把行高拉满仍显示 `###`，因此需直接调整列宽。
2. **副本快速试算（`_prepare_column`）**：
   * 解析目标单元格合并区域 `sheet.Range(addr).MergeArea` 的列范围 `cols`：
     * 若为单列，每次增加 `2.0` 字符宽；
     * 若为跨 $K$ 列合并单元格（如日企方眼纸），每列每次平均增加 `max(0.5, round(2.0 / K, 2))`，避免单列突兀变宽。
   * 每步先在内存中读取 `sheet.Cells(row, cols[0]).Text`：若仍全为 `#`，直接跳到下一步继续加宽，**省去中间无效的 PDF 导出**；
   * 一旦 `Cells.Text` 恢复正常显示，再追加一点余量（`max(0.3, round(1.0 / K, 2))`），导出一次 PDF 并通过 `autofix._evaluate` 验证完整显示且无压线，截取修改后图片。
3. **独立列宽写入日志与三色审核还原（`_repair_column` / `_write_rows` / `_commit_review`）**：
   * 使用独立记录标识 `_col_rid(card)`（`kind="col"`），保存修改前各列宽度 `base_cols` 与修改后各列宽度 `applied_cols`，与同行其他单元格的行高修改互不干扰；
   * 写入原文件前核对各列当前宽度（容差 `COL_WIDTH_TOLERANCE = 0.4`），支持断点续跑；
   * 当用户在三色审核中将卡片标为**绿色（还原并忽略）**或**黄色（还原后转人工）**提交时，自动将 `applied_cols` 还原回 `base_cols`。

