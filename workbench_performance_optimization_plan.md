# `/workbench` 性能优化实施方案

**基准**：`main` @ `053aec8`（2026-10-27）。
**行号约定**：文中行号以上述提交为准，同时给出函数名；行号漂移后按函数名定位。
**与 `workbench_performance_analysis.md` 的关系**：那份文档指认的代码问题**基本都成立**（本次逐条核对，无虚构），但三处需要修正：量化数字偏高约一个数量级、`content-visibility` 方案会引入回归、遗漏了收益最大的前端与传输问题。修正记录见 [§1.5](#15-实测数据与对旧文档的修正)。本文是可执行的实施方案，旧文档保留作为背景。

---

## 一、有哪些性能需要优化

### 1.1 全景

轮询链路：`static/batch.js:1357` 的 `setInterval(poll, 1800)` → `poll()`（`static/batch.js:133`）→ `GET /api/batch`（`web_clip_fixer.py:169`）→ `batch.snapshot()`（`core/batch.py:572`，全程持有 `wf.STATE_LOCK`）→ 全量 JSON → `render()`（`static/batch.js:699`）。

| 编号 | 位置 | 问题 | 触发时机 | 代价（N=卡片数, M=检测项数, F=文件数） | 优先级 |
|---|---|---|---|---|---|
| **F1** | `static/batch.js:753` | 每轮 render 为**全部**卡片重建 HTML 字符串，之后才和缓存比对 | 卡片有变化时（扫描期间每 1.8s） | O(N) 次重模板拼接 + O(N) `JSON.stringify` | **P0** |
| **F2** | `static/batch.js:625` | `layoutOf` 把 `batchState.busy` 算进布局签名，busy 一变就整页 `innerHTML` 重建 | 每次任务开始/结束 | 销毁重建上千节点 + 数千 `<img>` | **P0**（1 行修复） |
| **F3** | `static/batch.js:628-633` | `rebalanceFresh` 对**页面上每一张卡**调 `getBoundingClientRect()` | 每次整页重建后 | 强制同步布局，O(N) 次 reflow 读 | **P0** |
| **B1** | `core/batch.py:2284` | `_auto_skip_reason` 已算好 `items` 却不传给 `_can_highlight`，后者重新全表扫 | 每轮 snapshot × 每张已提交黄卡 | O(N_yellow × M) 字典比较 | **P0**（3 行修复） |
| **B2** | `core/batch.py:2216-2222` | `_handed_over_paths` 同样走 `_items_at` 全表扫，且每轮 snapshot 都被 `_mode()` 调一次 | 每轮 snapshot × 每张 handed_over 卡 | O(N_handed × M) | **P0** |
| **B3** | `core/batch.py:583` | 无 `problems` 的旧卡片在 snapshot 里再走一次 `_items_at` 全表扫 | 每轮 snapshot × 每张旧卡 | O(N_old × M) | P1 |
| **F4** | `static/batch.js:700`、`262` | `renderControls()` 在短路之前无条件执行；其中 `recheckAll.includes(c)` 是 O(N²) | **每轮**，无论有无变化 | O(N²) 引用比较 + 多次 O(N) 过滤 | P1 |
| **F5** | `static/batch.js:660`、`90`、`81` | `fileHeadHtml` / `autoQueueBadge` / `followupBadge` 各自全表过滤 `batchState.cards` | 每轮 render × 每个文件框 | O(F × N) | P1 |
| **F6** | `static/batch.js:585-590` | `cardHeight` 未命中时遍历整个 `cardHeights` 求平均 | `placeCards` 里每张新卡一次 | O(fresh × N) | P1 |
| **F7** | `static/batch.js:491-505` | `applySelection` 每次轮询全页 DOM 扫描 + 循环内 `shown.filter` | **每轮**（含短路分支） | O(N) DOM + O(F × N) | P2 |
| **B4** | `core/batch.py:581` | 每张卡 `copy.deepcopy(card)`，并把 `signatures`、`item_ids` 等内部字段一起下发 | 每轮 snapshot × N | O(N × 卡内嵌套结构) CPU + GC + 报文体积 | P1 |
| **T1** | `web_clip_fixer.py:169`、`static/batch.js:135` | 每 1.8s 全量 JSON，`cache: 'no-store'`，无 304；空闲审核时也照传 | **每轮** | 数 MB 序列化 + 传输 + 客户端 `JSON.parse` | P1 |
| **B5** | `core/batch.py:2265-2274` | 每张已提交黄卡对原 Excel `stat()` 一次；旧卡片走 `_unchanged_since_check` 时再 `stat()` 一次 | 每轮 snapshot × N_yellow | 每轮最多 2 × N_yellow 次 stat | P2 |
| **B6** | `core/fix_images.py:53-66` | 每张卡的图片 URL 触发最多 10 次 `stat()` | 每轮 snapshot × N | 实测约 47ms / 6000 次（见 §1.5） | **P3**（收益被高估） |
| **C1** | `static/batch.css` | 视口外文件框未做渲染隔离 | 滚动、布局 | 全量参与 Layout/Paint | **P3**（有前置条件，见 §3.8） |
| **B7** | `core/batch.py:3309-3316` | `_wanted_marks` 内对每张卡调 `_can_highlight` → 全表扫 | 标色同步时（提交/复检） | O(N_file × M) | P3（非轮询热路径，旁边就是 Excel COM 操作） |

### 1.2 后端问题详述

#### B1 / B2 / B3：`snapshot()` 持锁内的全表扫描

`snapshot()` 开头已经建好了索引（`core/batch.py:575-577`）：

```python
by_cell: dict[tuple[str, str, str], list[dict]] = {}
for item in wf.STATE["items"].values():
    by_cell.setdefault(_cell(item), []).append(item)
```

并在 `core/batch.py:602` 正确地传了下去：`_auto_skip_reason(card, by_cell.get(_cell(card), []))`。
`_auto_skip_reason`（`core/batch.py:2277`）也在 2281 行接住了：`items = _items_at(card) if items is None else items`。
**但下一行就把它丢了**（`core/batch.py:2284`）：

```python
if not _can_highlight(card):          # ← items 没传进去
```

而 `_can_highlight`（`core/batch.py:2593`）内部无条件全表扫：

```python
def _can_highlight(card: dict) -> bool:
    return bool(card.get("cell_addr")) and all(it.get("location_confirmed") is not False for it in _items_at(card))
```

`_items_at`（`core/batch.py:506`）就是一次全量线性过滤：

```python
def _items_at(card: dict) -> list[dict]:
    return [it for it in wf.STATE["items"].values() if _cell(it) == _cell(card)]
```

同一模式还有两处：
- `_handed_over_paths()`（`core/batch.py:2216`）调 `_unchecked_fix_reason(c, enabled=enabled)`，`items` 缺省 → `_card_fixable_types`（`core/batch.py:270`）→ `_items_at` 全表扫。它被 `_mode()`（`core/batch.py:2225`）调用，而 `_mode()` 在每轮 snapshot 的 `core/batch.py:619` 都会执行。
- `snapshot()` 自己在 `core/batch.py:583` 为「没有 `problems` 的旧卡片」再走一次 `_items_at(card)`，此时 `by_cell` 就在手边。

**代价**：这三处都在 `wf.STATE_LOCK` 内，会同时阻塞后台扫描/修复线程和前端的切色请求。

#### B4：全量 `deepcopy` 与冗余字段

`core/batch.py:581` 对每张卡 `copy.deepcopy(card)`，会深拷贝 `signatures`、`item_ids`、`problems`、`context`、`fix_images` 等全部嵌套结构。其中 `signatures`（`core/batch.py:553`）和 `item_ids`（`core/batch.py:558`）是纯内部字段，前端完全不读（已核对：`static/` 与 `templates/` 下无任何引用）。

深拷贝在这里其实**不必要**：`snapshot()` 只写 `result` 的顶层键（`problems` / `error_types` / `page_break` / `col_width` / `row_height` / `adoptable` / `auto_skip` / `highlighted` / `images`），唯一的嵌套就地修改是 `context["image"]`（`core/batch.py:608-611`）。

#### B5：每张黄卡对原 Excel `stat()`

`_auto_retry_blocked`（`core/batch.py:2265`）每张卡都 `wf._file_version(Path(path))`（`core/workflow.py:671`，内含一次 `path.stat()`）。同一个文件往往有几十张黄卡，版本却只有一个。旧卡片分支还会调 `_unchanged_since_check(path)`（`core/batch.py:129`），**里面又 stat 一次**。

#### B6：图片 URL 的 `stat()`

`fix_images.image_path()` 一次 `is_file()`，`fix_images.version()` 内部再调一次 `image_path()` 加一次 `path.stat()`（`core/fix_images.py:53-66`）。`snapshot()` 里每张图先 `image_path` 判存在、再 `_image_url` 取版本（`core/batch.py:604-611`），所以**每张图 3 次 stat**；`before` + `after` + `context` 三张图 = **每卡最多 10 次**（旧文档写 9 次，漏了 context 图单独的那次 `image_path`）。

### 1.3 前端问题详述

#### F1：全量重建 HTML 字符串（**前端最大的一块**）

`render()` 的瀑布流分支（`static/batch.js:751-755`）：

```javascript
const groups = fileGroups(shown);
const fresh = placeCards(groups, columnCount);
const html = new Map(shown.map(c => [c.id, cardHtml(c)]));      // ← 全部卡片重新拼 HTML
const heads = new Map(groups.map(([key]) => [key, fileHeadHtml(key)]));  // ← 全部文件框重新拼
let layout = layoutOf(groups, columnCount);                      // ← 再全量 stringify 一次 id
```

`cardHtmlCache` 的差异比对发生在**这之后**（`static/batch.js:758-762`）——字符串已经全量生成了。`cardHtml`（`static/batch.js:464`）内部还要跑 `handlingText`、`problemHtml`、`recheckButtonHtml`、`autoButtonHtml`、多次 `esc()`，是重活。

关键事实：**卡片 HTML 并不依赖大多数高频变化的状态**。逐个核对 `cardHtml` 的外部依赖：

| 依赖 | 来源 | 扫描期间是否高频变化 |
|---|---|---|
| `batchState.folder` | `static/batch.js:466` | 否 |
| `openBlocked(card)` → `open_blocked[wb_path]` | `static/batch.js:27` | 否（按文件，任务推进时才变） |
| `canMark(card)` → `busy` + `markWhileBusy()` | `static/batch.js:30` | 否（busy 只在任务起止翻转） |
| `recheckButtonHtml` → `recheckJoins()` / `queueBlockReason` / `queueNote` → `busy`、`job.kind`、`job.closing`、`job.options` | `static/batch.js:449` | 否 |
| `autoButtonHtml` → `autoBlockReason` → `busy`、`job.pause_requested`、`job.closing`、`autoQueueOpen()` | `static/batch.js:457`、`46` | 否 |
| `job.completed` / `job.scanned` / `job.repairing` / `job.file_errors` / `job.auto_queue` / `progress` / `logs` / `scan_eta` / `followups` | — | **是，但卡片 HTML 完全不读它们** |

也就是说：扫描期间每 1.8s 变的 `job.scanned`、`job.completed`、`progress`、`logs` 只影响进度条和文件框徽章，**却导致全部卡片 HTML 被重新生成**。

#### F2：`busy` 被算进布局签名

`static/batch.js:625`：

```javascript
const layoutOf = (groups, count) => JSON.stringify([groups.map(...), batchState.busy]);
```

`layout !== layoutKey` 就走 `else` 分支整页 `innerHTML` 重建（`static/batch.js:769-777`）。但布局（哪个卡片在哪一列）与 `busy` 无关，而文件标题栏在 `if` 分支里本来就会被单独打补丁（`static/batch.js:763-768`）。所以 `busy` 在这里是**冗余且有害**的：每次任务开始/结束都白触发一次上千节点的销毁重建。

#### F3：强制同步布局

`rebalanceFresh`（`static/batch.js:628-633`）在整页 `innerHTML` 刚赋值之后，立刻遍历**所有**卡片读几何：

```javascript
for (const element of $('cardList').querySelectorAll('.batch-card[data-id]')) {
  elements.set(element.dataset.id, element);
  cardHeights.set(element.dataset.id, element.getBoundingClientRect().height + CARD_GAP);
}
```

浏览器被迫同步计算整棵树的样式与布局；随后 `columnElements[i].append(...)` 移动节点（`static/batch.js:653`）、`restoreAnchor` 再读一次 rect（`static/batch.js:575`），形成连续多次全量重排。

#### F4 / F5 / F6 / F7

- `renderControls()`（`static/batch.js:240`）在 `render()` 第一行就被调用（`static/batch.js:700`），**位于短路判断之前**，所以每轮必跑；内部 `manualCards()` 被反复调用（`static/batch.js:257`、`262` 及 `recheckAllCards()` 内部），还有 `commitPlan()`。其中 `static/batch.js:262`：
  ```javascript
  const held = new Set(manual.filter(c => !recheckAll.includes(c)).map(fileKey)).size;
  ```
  `includes` 对对象做线性引用比较 → O(N²)。
- `fileHeadHtml(key)`（`static/batch.js:660`）自己 `batchState.cards.filter(...)`，`autoQueueBadge`（`static/batch.js:90`）与 `followupBadge`（`static/batch.js:81`）内部又各扫一遍全表 → 渲染全部文件框 O(F × N)。
- `cardHeight(id)`（`static/batch.js:585-590`）未命中缓存时遍历 `cardHeights.values()` 求和，`placeCards`（`static/batch.js:597`）里每张新卡触发一次。
- `applySelection()`（`static/batch.js:491-505`）每轮（含短路分支）遍历全页 `.batch-card`，并对每个文件按钮做一次 `shown.filter(...)`。

### 1.4 传输问题（T1）

`poll()` 用 `fetch('/api/batch', { cache: 'no-store' })`（`static/batch.js:136`），服务端 `_send_json` 输出 `Cache-Control: no-store`（`web_clip_fixer.py:112`），没有任何条件请求。空闲审核阶段（无后台任务、用户在看卡片）状态完全不变，却仍然每 1.8s 传一遍全量 JSON——含 `logs[-100:]`、`followups` 与 `job` 的 `deepcopy`（`core/batch.py:620-625`）。

### 1.5 实测数据与对旧文档的修正

**实测方法**（本机，Windows，NTFS，SSD，文件已预热）：创建 2000 个小文件，对每个执行 `is_file()` + `stat()`，共 4000 次系统调用耗时 31.1ms，**单次约 7.8µs**。

| 旧文档说法 | 核对结果 |
|---|---|
| 单次 `stat()` 0.05~0.15ms，6000 次 = 300~900ms | 实测单次 7.8µs，6000 次 ≈ **47ms**。偏高 6~19 倍 |
| 因此把「图片 stat」列为瓶颈 1、优化模块 1 | 降级为 **P3**（B6）。持锁 47ms 仍不理想，但不是头号问题 |
| 每张卡最多 9 次 stat | 实为 **10 次**（context 图额外单独调了一次 `image_path`） |
| `content-visibility: auto`「完全不影响页内搜索、锚点恢复和瀑布流高度」 | **错误**，会污染 `cardHeights` 并破坏列分配，见 §3.8 |
| 模块 2「只构造前端展示所需的字段」 | 改为**浅拷贝 + 字段黑名单**，见 §3.3。白名单式投影要重写字段清单，卡片字段太多（`after_kind`/`attempt_height`/`col_ids`/`verified`/`handed_over`/…），漏一个就是线上显示异常 |
| `_auto_retry_blocked` 加 `file_versions` 缓存 | 不完整：`_unchanged_since_check`（`core/batch.py:129-136`）内部还会再 stat 一次，必须把版本一起传进去 |
| — | **遗漏 F1**（全量重建 HTML 字符串）、**F4 的 `renderControls` 位置**、**T1**（传输层），这三项收益都大于旧文档的模块 1 |
| 行号 | 旧文档基于更早的版本，`snapshot` 标为 400-455（实为 572）、`_can_highlight` 标为 2229（实为 2593） |

### 1.6 明确不做的事

1. **不做前端虚拟化/窗口化**（只渲染视口内卡片）。`content-visibility: auto` 是它的低成本替代，但有前置条件（§3.8）；真正的虚拟化要重写瀑布流与锚点逻辑，收益不匹配风险。
2. **不改 1.8s 轮询为 WebSocket/SSE**。单机本地服务，轮询足够；T1 的 304 已能消掉空闲期的传输成本。
3. **不动 `commitPlan()`**（`static/batch.js:1098`）。它刚在 `053aec8` 改成基于并查集的共享行高/列宽分组，逻辑正确且不是每轮热点（`renderControls` 每轮调一次，但内部是 O(N·α)）。
4. **不引入前端框架或构建步骤**。现有 `static/batch.js` 是纯脚本 + `vm` 沙箱测试（`tests/test_batch_review.js`），保持这个形态。

---

## 二、优化需求（详细）

### 2.0 全局不变量

所有步骤共同遵守：

- **零功能变更**：三色审核、单项复检、自动调整、排队、提交、标色还原的行为完全不变。
- **零视觉变更**：卡片外观、瀑布流列分配结果、滚动位置保持、文件框徽章文案不变。
- **不改数据结构**：`wf.STATE` 与进度备份文件的格式不变（老进度必须能继续用）。
- **回归测试必须全绿**：
  ```bash
  python -m unittest discover -s tests -p "test_*.py"
  node --test tests/test_batch_review.js
  ```
- **每步独立可回退**：一步一个提交，出问题 `git revert` 单个提交即可。

### 2.1 R0（Step 0）：先有基线，再谈优化

**目标**：用实测数字替换所有估算，让后续每步可验证。
**需求**：
1. 准备接近真实规模的数据：60~80 个 Excel、800~1000 张待审核卡片（可用现有文件夹，或构造）。
2. 后端指标：`snapshot()` 单次持锁耗时（ms）、`/api/batch` 响应字节数、单轮 `stat()` 次数。
3. 前端指标：`render()` 单次耗时、`cardHtml` 总耗时、`rebalanceFresh` 耗时、整页 `innerHTML` 重建次数、Performance 面板里的 Long Task 数量。
4. 分别在两个场景下各测一次：**扫描进行中**（卡片持续新增）与**空闲审核**（无后台任务，用户点击切色）。
**验收**：产出一张「优化前」表格，字段与 §四 各步的「验证」一致；测量用的临时打点在记录完成后删除，不进提交。

### 2.2 R1（Step 1）：后端消除 snapshot 内的全表扫描

**目标**：`snapshot()` 持锁内的 O(N × M) 降为 O(N + M)。
**需求**：
1. `_can_highlight` 支持调用方传入已算好的 `items`；不传时行为完全不变（另有 3 个调用点：`core/batch.py:2116`、`2284`、`3316`，以及 `tests/test_mode_switch.py` 间接依赖）。
2. `_auto_skip_reason` 把自己已经算好的 `items` 传下去。
3. `_handed_over_paths` 支持传入 `by_cell` 索引；`_mode` 同样支持，并透传。两者都必须保留无参调用（`core/batch.py:2241`、`2257`、`3622`，以及 `tests/test_mode_switch.py:125`、`136`、`746` 直接调 `batch._mode()`）。
4. `snapshot()` 里 `core/batch.py:583` 的 `_items_at(card)` 改用 `by_cell`。
5. 顺带修 B5：`_auto_retry_blocked` 支持单轮 `file_versions` 缓存，`_unchanged_since_check` 支持传入已取到的版本。
**必须保持的行为**：
- `by_cell.get(_cell(c), [])` 与 `_items_at(c)` 必须等价。已核对：两者遍历同一个 `wf.STATE["items"].values()`，元素顺序一致，因此 `_can_highlight` 的 `all()`、`autofix.review_only_reason()`、`_card_fixable_types` 里 `dict.fromkeys(raw)` 生成的标签顺序都不受影响（标签顺序会变就会改提示文案）。
- `_unchanged_since_check` 在 `stat` 失败时仍返回 `False`。
**验收**：unittest 全绿；R0 的 `snapshot()` 持锁耗时下降；`auto_skip` 文案在真机上逐字不变（抽查若干黄卡）。

### 2.3 R2（Step 2）：后端精简拷贝与下发字段

**目标**：降低 snapshot 的 CPU/GC，并直接减小每 1.8s 的报文。
**需求**：
1. `core/batch.py:581` 的 `copy.deepcopy(card)` 换成浅拷贝 `dict(card)`。
2. `context` 必须单独拷一份（它是 snapshot 唯一就地修改的嵌套结构）。
3. 下发前从 `result` 里删掉 `signatures`、`item_ids`。
**必须保持的行为**：
- **只删序列化输出，绝不动 `wf.STATE` 里的卡片**——`item_ids` 在内部逻辑和测试里都在用（`tests/test_batch.py:866` 断言 `_card_for()` 的结果，`tests/test_mode_switch.py:221` 读 STATE 里的卡片，两处都不是 snapshot 输出，已核对）。
- 浅拷贝后不得污染 STATE：`problems`、`error_types` 只在原值为假时才整体赋新列表；`fix_images` 只读。
**验收**：unittest 全绿（重点 `tests/test_batch.py:1257`、`1289` 断言 snapshot 的 `context`）；R0 记录的响应字节数下降；页面卡片显示无变化。

### 2.4 R3（Step 3）：前端只为真正变化的卡片生成 HTML

**目标**：消掉 F1 与 F2，让扫描期间的每轮 render 从「全量拼字符串」变成「只拼变化的那几张」。
**需求**：
1. 引入**按卡片记忆化**：只有卡片自身内容或其依赖的「chrome 签名」变化时，才重新执行 `cardHtml`。
2. chrome 签名必须**精确**覆盖 `cardHtml` 的全部外部依赖（见 §1.3 的依赖表），且**不得**包含 `job.completed`、`job.scanned`、`job.repairing`、`job.file_errors`、`job.auto_queue`、`progress`、`logs`、`scan_eta`、`followups`——否则签名每轮都变，记忆化等于没做。
3. 文件标题栏同理按文件记忆化；它的依赖更宽（含 `autoQueueBadge` 用到的 `job.completed` 等），单独一个签名。
4. 从 `layoutOf`（`static/batch.js:625`）里移除 `batchState.busy`。
**必须保持的行为**：
- busy 翻转、`open_blocked` 变化、暂停/收尾状态变化时，卡片上的「后台处理完这个文件后可标记」文案与按钮 disabled 状态必须立刻更新。
- `renderKey` 短路语义不变：完全无变化时仍走 `applySelection()` 后返回。
- 图片不得重新加载（`outerHTML` 替换的范围不能扩大）。
**验收**：`node --test tests/test_batch_review.js` 全绿（21 个用例覆盖 `cardHtml` 输出与交互）；真机验证清单见 §6.2；R0 的 `cardHtml` 总耗时在扫描期间从 O(N) 降到 O(变化卡片数)。

### 2.5 R4（Step 4）：前端只测量新卡高度

**目标**：消掉 F3 的强制同步全量重排，并为 C1 铺路。
**需求**：
1. `rebalanceFresh` 只对 `fresh` 集合里的新卡调 `getBoundingClientRect()`。
2. 已记录高度的卡片不再重复测量。
3. 维护 `cardHeightSum`，`cardHeight(id)` 求平均从 O(N) 变 O(1)；卡片被移除时要从 sum 里扣掉，避免长期漂移。
**必须保持的行为**：
- 瀑布流列分配结果不变（新卡仍然放进当前最矮的列）。
- 卡片高度会随内容变化（例如切色后文案变长）：需要明确策略——本方案选择**只在整页重建时测量新卡**，与当前行为一致（当前也只在 `rebalanceFresh` 里测），不引入新的重测时机。
**验收**：真机确认三列高度均衡、切换筛选与缩放窗口后不歪、`restoreAnchor` 不跳。

### 2.6 R5（Step 5）：前端清理剩余平方项

**目标**：消掉 F4 / F5 / F6 / F7。
**需求**：
1. `static/batch.js:262` 的 `includes` 改 `Set.has`。
2. `fileHeadHtml` / `autoQueueBadge` / `followupBadge` 改为接收已分组好的该文件卡片；`render()` 里用 `fileGroups(batchState.cards)`（`static/batch.js:109`）建一次映射。
3. `renderControls()` 内 `manualCards()`（`static/batch.js:124`）只算一次并复用。
4. `applySelection()` 循环前按 `fileKey` 分组一次；再加一层「选中集合与上轮相同就整体跳过」。
**必须保持的行为**：「选中本文件的卡」按钮的按下态与文案、批量改色范围、复检按钮的 held 计数全部不变。
**验收**：`node --test tests/test_batch_review.js`（含选中/全选/反选用例）+ 真机点击验证。

### 2.7 R6（Step 6）：传输层条件请求

**目标**：空闲审核时每 1.8s 的传输与客户端解析归零。
**需求**：
1. 服务端缓存上一次的响应体，用内容哈希做 `ETag`。
2. `If-None-Match` 命中时返回 304、不带 body。
3. 前端 `poll()` 保存上次 etag，主动带 `If-None-Match`，收到 304 时**不替换 `batchState`、不调 `render()`**。
4. **不采用手工维护的版本计数器**，理由见 §3.7。
**必须保持的行为**：任何状态变化（切色、扫描进度、日志、排队）都必须立刻反映到页面——这是本步唯一但严重的风险点。
**验收**：Network 面板确认空闲时 304 命中率接近 100%、有变化时立即 200；真机走一遍完整审核流程。

### 2.8 R7（Step 7，可选）：CSS 渲染隔离

**目标**：滚动与绘制只处理视口内的文件框。
**需求**：Step 4 必须先完成；`scrollAnchor`（`static/batch.js:560`）必须先改成不读视口外卡片的 rect。
**验收**：真机滚动 1000 张卡的列表，确认帧率提升、滚动条不跳、锚点不跳、三列不歪、Ctrl+F 仍能定位视口外卡片。**任一项不达标就整体回退这一步**（纯增益项，不做不影响前 6 步）。

### 2.9 R8（Step 8，最低优先级）：图片版本内存缓存

**目标**：省掉 B6 的每轮约 47ms。
**需求**：`save`（`core/fix_images.py:26`）、`clear`（`core/fix_images.py:69`）、`import_all`（`core/fix_images.py:94`）三处必须同步失效缓存。
**为什么排最后**：收益实测最小，而失效点分散，做错的后果是**浏览器一直显示旧截图**（用户可见的错误）。`tests/test_batch.py:1934-1936` 正是守这个行为的测试。

---

## 三、如何设计

### 3.1 后端：`items` 单向下传（B1/B2/B3）

设计原则：**索引在 `snapshot()` 顶部构建一次，沿调用链显式传递，所有新增参数都是可选的**。

```
snapshot()  ──建 by_cell──┬─→ _auto_skip_reason(card, items, file_versions)
                          │        ├─→ _can_highlight(card, items)      ← 新增可选参数
                          │        ├─→ _unchecked_fix_reason(card, items, enabled)   （已支持）
                          │        └─→ _auto_retry_blocked(card, file_versions)      ← 新增可选参数
                          │                 └─→ _unchanged_since_check(path, current) ← 新增可选参数
                          └─→ _mode(by_cell) → _handed_over_paths(by_cell)
                                   └─→ _unchecked_fix_reason(c, items, enabled)
```

**为什么用可选参数而不是全局/线程局部缓存**：`_can_highlight` 等函数有轮询路径之外的调用方（`core/batch.py:2116` 在 `_hand_over` 里、`3316` 在 `_wanted_marks` 里），那些地方没有 `by_cell`，也不在每 1.8s 的热路径上（旁边就是 Excel COM 调用）。可选参数让热路径拿到 O(1) 查表，冷路径保持原样，不引入隐式状态。

**为什么不顺手改 `_items_at` 本身**：它有 20+ 个调用点（`core/batch.py:271`、`583`、`830`、`837`、`843`、`926`、`1505`、`1733`、`1849`、`2030`、`2037`、`2110`、`2123`、`2330`、`2343`、`2366`、`2594`），大多在修复/复检流程里，那里 `by_cell` 不存在或已过期（`items` 会被就地修改）。全局缓存会引入失效问题，收益却只在轮询路径，不划算。

### 3.2 后端：单轮文件版本缓存（B5）

`file_versions: dict[str, dict[str, int] | None]` 是 **`snapshot()` 的局部变量**，生命周期只有一轮请求。不做跨轮缓存——原文件可能被用户在 Excel 里保存，跨轮缓存会让「自动试过没通过」的判断读到过期版本，属于正确性问题。`None` 表示 stat 失败（文件被占用/已删），与 `False` 区分开，避免把失败结果当成有效版本缓存住。

### 3.3 后端：浅拷贝 + 字段黑名单（B4）

```
深拷贝整卡  →  dict(card) 浅拷贝 + 单独拷 context + 删两个内部字段
```

安全性论证（已逐行核对 `snapshot()` 对 `result` 的全部写操作）：

| 写操作 | 位置 | 是否就地改嵌套对象 | 浅拷贝是否安全 |
|---|---|---|---|
| `result["problems"] = [...]` | 583 | 否，赋新列表 | ✅ |
| `result["error_types"] = [...]` | 585 | 否，赋新列表 | ✅ |
| `result["page_break"/"col_width"/"row_height"] = {...}` | 591/593/597 | 否，赋新字典 | ✅ |
| `result["adoptable"/"auto_skip"/"highlighted"/"images"] = ...` | 599/602/603/604 | 否 | ✅ |
| `context["image"] = ...` | 611 | **是** | ⚠️ 需 `result["context"] = dict(card["context"])` |

**为什么黑名单而不是白名单**：白名单要枚举卡片的全部展示字段，`cardHtml` / `handlingText` / `commitPlan` 读到的字段有 30+ 个（`after_kind`、`attempt_height`、`base_height`、`col_ids`、`verified`、`waiting`、`gen`、`handed_over`、`manual`、`marked`、`message`、`kind`、`seq`…），漏一个就是难查的显示异常。黑名单只删两个已确认无人读的字段，风险可控。

### 3.4 前端：三段式渲染签名（F1/F2）

核心思路：把「什么东西变了」拆成三个**互不重叠**的签名，各自决定要不要重算对应产物。

```
cardSig(card)   = JSON.stringify(card)                    → 决定这张卡要不要重拼 cardHtml
cardChromeSig   = [folder, busy, open_blocked, job.kind,   → 卡片 HTML 的全局依赖
                   job.closing, job.pause_requested,
                   job.to_manual, job.options,
                   markWhileBusy(), autoQueueOpen()]
headChromeSig   = cardChromeSig + [job.completed,          → 文件标题栏额外依赖
                   job.scanned, job.repairing,
                   job.file_errors, job.auto_queue,
                   job.status, followups]
layoutSig       = 各列的卡片 id 序列 + columnCount          → 决定要不要整页重建（不再含 busy）
```

记忆化表：`cardMemo: Map<id, {sig, chrome, html}>`、`headMemo: Map<fileKey, {sig, chrome, html}>`。

命中条件：`cardMemo.get(id).sig === cardSig(card) && cardMemo.get(id).chrome === cardChromeSig`。

**为什么 `cardChromeSig` 必须这么窄**：如果把 `job.completed` / `job.scanned` 放进去（当前 `renderKey` 就含 `job.completed`），扫描期间每完成一个文件签名就变，全部卡片记忆化失效，等于没优化。§1.3 的依赖表是逐项核对 `cardHtml` → `canMark` / `recheckButtonHtml` / `autoButtonHtml` → `queueBlockReason` / `autoBlockReason` / `autoQueueOpen` / `busyJobName` 得出的，不是猜的。

**为什么不用对象引用做记忆化键**：`poll()` 每轮 `batchState = value`（`static/batch.js:144`）整体替换，卡片对象每轮都是新的，引用相等永远不成立。

**代价**：`cardSig` 仍要对每张卡做一次 `JSON.stringify`。这比 `cardHtml` 便宜得多（后者含多次 `esc()`、模板拼接、`handlingText` 分支判断），而且全局 `renderKey` 可以直接由各卡 `cardSig` 拼接得到，不额外增加一趟序列化。彻底消掉这趟 stringify 属于 Step 6 之后的课题（需要服务端下发每卡版本号）。

**F2 的修复**：`layoutOf` 去掉 `batchState.busy`。正确性依据：busy 变化时 `renderKey` 仍然变化 → `render()` 不会短路 → `heads` 重新生成 → `if` 分支里的 `headHtmlCache` 比对会更新标题栏（`static/batch.js:763-768`）。所以 busy 变化仍然会反映到页面，只是不再触发整页重建。

### 3.5 前端：高度测量（F3/F6）

```
cardHeights: Map<id, height>     （已有）
cardHeightSum: number            （新增，与 cardHeights 同步维护）
rememberCardHeight(id, h)        （新增，唯一的写入口：先减旧值再设新值）
```

`rebalanceFresh` 的测量循环从「遍历页面全部卡片」改为「遍历 `fresh` 集合，逐个 `querySelector`」。`fresh` 通常只有 1~几张，`querySelector` 的成本远低于强制整树同步布局。

`elements` 映射（用于 `append` 挪动节点）只在**确实有卡需要挪动**时才收集，避免白遍历。

### 3.6 前端：分组映射一次构建（F5/F7）

`render()` 里已经有 `groups = fileGroups(shown)`（按筛选后的卡片分组）。再建一个不受筛选影响的全量映射给文件标题栏用（标题栏的数量统计按整个文件算，不随筛选变，见 `static/batch.js:658-659` 的注释）：

```javascript
const allByFile = new Map(fileGroups(batchState.cards));
```

一次 O(N)，之后 `fileHeadHtml(key, allByFile.get(key))`、`autoQueueBadge(key, fileCards)`、`followupBadge(key, fileCards)` 都只看本文件的几张~几十张卡。

`applySelection` 同理：循环前按 `fileKey` 分组 `shownCards()` 一次，循环内查 Map。

### 3.7 传输：内容哈希 ETag，而不是版本计数器（T1）

两种可选设计：

| 方案 | 服务端 CPU | 正确性风险 |
|---|---|---|
| A. 手工维护 epoch，状态变更处 `bump()`；命中则**完全跳过** snapshot 构建 | 最省 | **高**：漏掉任何一个变更点（`wf.STATE["logs"]`、`scan_progress`、`job` 的中间态大多不走 `_persist_state()`）就是「页面不刷新」的疑难 bug，且难复现 |
| B. 照常构建 snapshot，序列化后取 SHA-1 做 ETag，缓存 body；命中则 304 | 省网络与客户端解析，不省服务端 CPU | **无**：etag 由内容算出，构造性正确 |

**选 B**。理由：T1 的主要痛点是数 MB 报文的传输与客户端 `JSON.parse`（每 1.8s 一次，且 parse 在主线程、会阻塞点击响应），B 完整解决这部分；服务端 CPU 已由 Step 1/2 处理。A 的额外收益不值得引入「漏 bump → UI 冻结」这一类 bug。

实现要点：
- `_send_bytes`（`web_clip_fixer.py:104`）增加可选 `etag` 参数；304 分支不写 body、不发 `Content-Length`。
- 保留 `Cache-Control: no-store`，由 JS 主动管理 `If-None-Match`，不依赖浏览器 HTTP 缓存语义。
- **需要真机确认**：JS 手动发 `If-None-Match` 时浏览器是否把 304 原样交给 `fetch`（`response.status === 304`）。若被规范化成 200，退化为服务端返回 `{"unchanged": true}` + 200，前端同样跳过 `render()`。

### 3.8 CSS `content-visibility` 的前置条件（C1）

**旧文档的方案会引入回归，不能照抄。** 机制：`content-visibility: auto` 下，视口外子树跳过布局，其 `getBoundingClientRect()` 返回 `contain-intrinsic-size` 的占位值（旧文档设的 620px）。而本项目的瀑布流**依赖真实高度**：

- `rebalanceFresh` 会把测到的高度写进 `cardHeights`（`static/batch.js:632`）→ 占位值被永久缓存 → `placeCards` / `shortestColumn`（`static/batch.js:592`）之后全按错的高度分列，三列越分越歪。
- `scrollAnchor`（`static/batch.js:560-570`）遍历全部卡片读 rect 来找锚点，`restoreAnchor`（`static/batch.js:573-576`）再读一次；占位值会让锚点数学算错，重画后页面跳位。

**前置条件（缺一不可）**：
1. Step 4 已完成：只测新卡高度，不再全量测量。
2. `scrollAnchor` 改成只考虑视口内的元素（当前它有 `rect.bottom < 90 || rect.top > innerHeight` 的过滤，但过滤本身要先读 rect；改成先用 `element.offsetTop` 之类的廉价判断，或只遍历前 K 个）。
3. `contain-intrinsic-size` 用 `auto 620px` 形式（`auto` 关键字让浏览器记住已渲染过的真实高度，缓解滚动条跳动）。

**回退判据**：真机出现列分配不均、锚点跳位、滚动条明显跳动中的任一项，直接删掉这两行 CSS。

### 3.9 图片版本缓存（B6）

按旧文档模块 1 的思路做，但有两处修正：
1. 旧文档在 `save()` 里给版本号追加 `_WRITE_SEQ`（进程内自增）。后果：进程重启后所有图片 URL 变化一次，浏览器重新下载全部截图。可接受，但要知情；若不接受，改为只用 `st_mtime_ns + st_size`（与现有 `version()` 格式一致）。
2. 缓存 `None`（文件不存在）也要小心：`save()` 之前若已探测过同名文件为不存在，写入后必须覆盖缓存，否则卡片永远显示「暂无截图」。

---

## 四、如何写代码

### Step 0：基线测量

临时打点（**不进提交**）：

```python
# core/batch.py: snapshot() 首尾
def snapshot() -> dict:
    _t0 = time.perf_counter()
    with wf.STATE_LOCK:
        ...
    logging.info("snapshot 持锁 %.1f ms", (time.perf_counter() - _t0) * 1000)
```

```javascript
// static/batch.js: render() 首尾；cardHtml 累计
const _t0 = performance.now();
// ... render 主体 ...
console.debug('render', (performance.now() - _t0).toFixed(1), 'ms, cards', batchState.cards.length);
```

配合 Chrome Performance 面板录制 60 秒扫描过程，记录 Long Task 数量与 `recalculate style` / `layout` 耗时占比。填入下表（示例格式，数值待测）：

| 指标 | 扫描中 | 空闲审核 |
|---|---|---|
| `snapshot()` 持锁 | ? ms | ? ms |
| `/api/batch` 报文 | ? KB | ? KB |
| `render()` 单次 | ? ms | ? ms |
| 其中 `cardHtml` 合计 | ? ms | ? ms |
| `rebalanceFresh` | ? ms | 不触发 |
| 整页重建次数 / 分钟 | ? | ? |
| Long Task 数 / 分钟 | ? | ? |

### Step 1：后端消除全表扫描

**`core/batch.py:2593`** — `_can_highlight` 增加可选 `items`：

```python
def _can_highlight(card: dict, items: list[dict] | None = None) -> bool:
    if not card.get("cell_addr"):
        return False
    items = _items_at(card) if items is None else items
    return all(it.get("location_confirmed") is not False for it in items)
```

**`core/batch.py:2284`** — 把已算好的 `items` 传下去：

```python
    if not _can_highlight(card, items):
        return "单元格位置没确认，不能自动调整"
```

**`core/batch.py:129`** — `_unchanged_since_check` 支持传入版本：

```python
def _unchanged_since_check(path: str, current: dict[str, int] | None = None) -> bool:
    """上次整份检测之后原文件没再改过（检测规则也没变）：重新导出检测的结果和上次一样"""
    checks = _sheet_checks().get(path) or {}
    if current is None:
        try:
            current = wf._file_version(Path(path))
        except OSError:
            return False
    return checks.get("*") == current
```

**`core/batch.py:2265`** — `_auto_retry_blocked` 单轮缓存文件版本：

```python
def _auto_retry_blocked(card: dict, file_versions: dict[str, Any] | None = None) -> bool:
    """自动试过没通过的格子：文件之后没改过就不再试（同样的算法，结果也一样）。调用方持有 STATE_LOCK"""
    path = card["wb_path"]
    if file_versions is not None and path in file_versions:
        current = file_versions[path]
    else:
        try:
            current: Any = wf._file_version(Path(path))
        except OSError:
            current = None
        if file_versions is not None:
            file_versions[path] = current
    if current is None:
        return False
    done = _data().get("auto_versions", {}).get(path)
    # 旧版本没记自动处理完时的文件版本：按上次整份检测之后有没有改过判断
    return done == current if done is not None else _unchanged_since_check(path, current)
```

**`core/batch.py:2277`** — `_auto_skip_reason` 增加 `file_versions` 并透传：

```python
def _auto_skip_reason(card: dict, items: list[dict] | None = None, enabled: set[str] | None = None,
                      file_versions: dict[str, Any] | None = None) -> str:
    ...
    if not card.get("handed_over") and _auto_retry_blocked(card, file_versions):
        return "自动试过没通过，文件之后没改过，再试结果也一样"
    return ""
```

**`core/batch.py:2216`** — `_handed_over_paths` 复用索引：

```python
def _handed_over_paths(by_cell: dict[tuple[str, str, str], list[dict]] | None = None) -> list[str]:
    """调用方持有 STATE_LOCK：当前文件夹里还有可交回自动的问题的文件"""
    enabled = _active_fix_types()
    cards = sorted(_data()["cards"].values(), key=lambda c: c["seq"])
    return list(dict.fromkeys(
        c["wb_path"] for c in cards
        if c.get("handed_over") and not c.get("dismissed") and wf.in_current_folder(c["wb_path"])
        and not _unchecked_fix_reason(c, None if by_cell is None else by_cell.get(_cell(c), []), enabled)))
```

**`core/batch.py:2225`** — `_mode` 透传：

```python
def _mode(by_cell: dict[tuple[str, str, str], list[dict]] | None = None) -> dict[str, bool]:
    ...
        "to_auto": not wf.STATE["busy"] and not (auto and job.get("status") == "interrupted")
                   and bool(_handed_over_paths(by_cell)),
```

**`core/batch.py:572-619`** — `snapshot()` 内：

```python
def snapshot() -> dict:
    with wf.STATE_LOCK:
        cards = []
        by_cell: dict[tuple[str, str, str], list[dict]] = {}
        for item in wf.STATE["items"].values():
            by_cell.setdefault(_cell(item), []).append(item)
        file_versions: dict[str, Any] = {}          # 新增：单轮文件版本缓存
        for card in sorted(_data()["cards"].values(), key=lambda c: c["seq"]):
            if card.get("dismissed") or not wf.in_current_folder(card["wb_path"]):
                continue
            items = by_cell.get(_cell(card), [])     # 新增：本轮只查一次
            result = copy.deepcopy(card)
            if not result.get("problems"):
                result["problems"] = list({p["kind"]: p for p in map(_problem, items)}.values())   # 改：用 items
            ...
            if card["color"] == "yellow" and card.get("finalized"):
                result["auto_skip"] = _auto_skip_reason(card, items, file_versions=file_versions)  # 改：传缓存
            ...
                "mode": _mode(by_cell), "open_blocked": _open_blocked([c["wb_path"] for c in cards]),  # 改：传索引
```

> `_can_highlight` 的另外两个调用点（`core/batch.py:2116` 在 `_hand_over`、`3316` 在 `_wanted_marks`）**不改**：不在轮询热路径，且现场没有 `by_cell`。若要一并优化（B7），可在 `_wanted_marks` 里为该文件的卡片建一次局部索引。

**验证**：`python -m unittest discover -s tests -p "test_*.py"`，重点 `tests/test_batch.py`、`tests/test_mode_switch.py`、`tests/test_later_review.py`。

---

### Step 2：精简拷贝与字段

**`core/batch.py:581`**：

```python
            result = dict(card)                      # 改：浅拷贝即可，下面只改顶层键
            for internal in ("signatures", "item_ids"):   # 新增：前端不读的内部字段不下发
                result.pop(internal, None)
```

**`core/batch.py:608-611`**：

```python
            if result.get("context"):
                context = dict(card["context"])       # 改：唯一被就地修改的嵌套结构，单独拷一份
                result["context"] = context
                name = context["image_file"]
                context["image"] = _image_url(name) if fix_images.image_path(name) else ""
```

**验证**：`python -m unittest discover -s tests -p "test_*.py"`（重点 `tests/test_batch.py:1257`、`1289` 的 `context` 断言）+ `node --test tests/test_batch_review.js`；对比 Step 0 的报文体积。

---

### Step 3：前端记忆化渲染

**`static/batch.js` 顶部（约 10-25 行的状态区）新增**：

```javascript
// 卡片 HTML 依赖的全局状态：只放 cardHtml → canMark / recheckButtonHtml / autoButtonHtml 真正读到的字段。
// 不要加 job.completed / scanned / repairing / file_errors / auto_queue / progress / logs，
// 那些每轮都在变，加进来记忆化就失效了（它们只影响进度条和文件框徽章）
const cardChromeSig = () => JSON.stringify([
  batchState.folder, batchState.busy, batchState.open_blocked, markWhileBusy(), autoQueueOpen(),
  batchState.job?.kind, batchState.job?.closing, batchState.job?.pause_requested,
  batchState.job?.to_manual, batchState.job?.options,
]);
// 文件标题栏额外依赖任务推进（排队徽章、✅ 已处理完）
const headChromeSig = () => JSON.stringify([
  cardChromeSig(), batchState.job?.workbooks, batchState.job?.completed, batchState.job?.scanned,
  batchState.job?.repairing, batchState.job?.file_errors, batchState.job?.auto_queue,
  batchState.job?.status, batchState.followups,
]);
const cardMemo = new Map();   // id -> { sig, chrome, html }
const headMemo = new Map();   // fileKey -> { sig, chrome, html }

function cardHtmlMemo(card) {
  const sig = JSON.stringify(card), chrome = cardChromeSig();
  const hit = cardMemo.get(card.id);
  if (hit && hit.sig === sig && hit.chrome === chrome) return hit.html;
  const html = cardHtml(card);
  cardMemo.set(card.id, { sig, chrome, html });
  return html;
}
```

**`static/batch.js:625`** — 布局签名去掉 `busy`（F2）：

```javascript
const layoutOf = (groups, count) => JSON.stringify([groups.map(([key, cards]) => [key, cardColumns(cards, count).map(column => column.map(c => c.id))])]);
```

**`static/batch.js:751-755`** — 用记忆化版本，并按文件传入卡片：

```javascript
    const groups = fileGroups(shown);
    const fresh = placeCards(groups, columnCount);
    const allByFile = new Map(fileGroups(batchState.cards));     // Step 5 会用到，这里一起建
    const html = new Map(shown.map(c => [c.id, cardHtmlMemo(c)]));
    const headChrome = headChromeSig();
    const heads = new Map(groups.map(([key]) => [key, fileHeadHtmlMemo(key, allByFile.get(key) || [], headChrome)]));
    let layout = layoutOf(groups, columnCount);
```

配套 `fileHeadHtmlMemo`（`fileHeadHtml` 本身按 Step 5 改成接收 `fileCards`）：

```javascript
function fileHeadHtmlMemo(key, fileCards, chrome) {
  const sig = JSON.stringify(fileCards.map(c => [c.id, c.color, c.finalized, c.auto_skip, c.adoptable]));
  const hit = headMemo.get(key);
  if (hit && hit.sig === sig && hit.chrome === chrome) return hit.html;
  const html = fileHeadHtml(key, fileCards);
  headMemo.set(key, { sig, chrome, html });
  return html;
}
```

**清理**：卡片被移除后要清掉记忆化条目，避免 Map 无限增长——在 `render()` 末尾按当前 `batchState.cards` 的 id 集合裁剪 `cardMemo`（`headMemo` 按 `allByFile` 的键裁剪）。

**注意 `renderKey`**：`static/batch.js:730` 的全局 key 保持不变（它含 `batchState.cards` 全量 stringify）。它是「要不要进入渲染」的总闸；记忆化解决的是「进入渲染后要不要重拼每张卡」。Step 6 之后可以让 304 直接跳过整个 `render()`。

**验证**：`node --test tests/test_batch_review.js`；真机清单见 §6.2。

---

### Step 4：只测新卡高度

**`static/batch.js:581-590`**：

```javascript
const cardHeights = new Map();
let cardHeightSum = 0;

// 唯一写入口：先扣掉旧值，保证 cardHeightSum 与 cardHeights 始终一致
function rememberCardHeight(id, height) {
  if (cardHeights.has(id)) cardHeightSum -= cardHeights.get(id);
  cardHeights.set(id, height);
  cardHeightSum += height;
}

function forgetCardHeight(id) {
  if (!cardHeights.has(id)) return;
  cardHeightSum -= cardHeights.get(id);
  cardHeights.delete(id);
}

function cardHeight(id) {
  if (cardHeights.has(id)) return cardHeights.get(id);
  return cardHeights.size ? cardHeightSum / cardHeights.size : 520;
}
```

**`static/batch.js:628-635`** — 只测 `fresh`：

```javascript
function rebalanceFresh(groups, count, fresh) {
  if (!fresh.size) return false;
  const fileElements = $('cardList').querySelectorAll('.batch-file');
  if (fileElements.length !== groups.length) return false;
  const elements = new Map();
  for (const id of fresh) {
    const element = $('cardList').querySelector(`.batch-card[data-id="${cssEscape(id)}"]`);
    if (!element) continue;
    elements.set(id, element);
    rememberCardHeight(id, element.getBoundingClientRect().height + CARD_GAP);
  }
  ...
```

**`static/batch.js:653`** 的 `append` 需要**所有**参与挪动的卡片元素，不只是 `fresh`。因此挪动分支里按需补齐：

```javascript
    if (!changed) return;
    moved = true;
    cardColumns(cards, count).forEach((column, i) => columnElements[i].append(
      ...column.map(c => elements.get(c.id)
        || $('cardList').querySelector(`.batch-card[data-id="${cssEscape(c.id)}"]`)).filter(Boolean)));
```

（`append` 已有节点是移动而非重建，图片不会重新加载；这里只在确有卡片换列时才查 DOM。）

**清理**：卡片移除时调 `forgetCardHeight(id)`，位置与 `cardMemo` 的裁剪放在一起。

**验证**：真机确认三列均衡、筛选切换与窗口缩放后不歪、锚点不跳。

---

### Step 5：清理平方项

**`static/batch.js:262`**：

```javascript
  const manual = manualCards(), recheckAll = recheckAllCards(), recheckSet = new Set(recheckAll);
  const held = new Set(manual.filter(c => !recheckSet.has(c)).map(fileKey)).size;
```

同时把 `renderControls()` 里重复的 `manualCards()` 调用（`static/batch.js:257` 等）改为复用上面这个 `manual` 变量——注意需要把它的计算提到函数前部。

**`static/batch.js:81`、`90`、`660`** — 三个函数改为接收本文件卡片：

```javascript
function followupBadge(path, fileCards) {
  const kinds = (batchState.followups || []).filter(entry => {
    const ids = new Set(entry.card_ids);
    return fileCards.some(c => ids.has(c.id));
  }).map(entry => FOLLOWUP_NAMES[entry.kind]);
  ...
}

function autoQueueBadge(path, fileCards) {
  const job = batchState.job;
  if (batchState.busy && job?.kind === 'batch_scan' && job.options?.scan_only) {
    const queued = new Set(job.auto_queue || []);
    return fileCards.some(c => queued.has(c.id)) ? '<span ...>⏳ 排队等自动调整 · 全部文件扫完后开始</span>' : '';
  }
  ...   // 其余分支只用 path，不变
}

function fileHeadHtml(key, fileCards) {
  const cards = fileCards;      // 原来是 batchState.cards.filter(c => fileKey(c) === key)
  ...
  const badges = autoQueueBadge(key, cards) + followupBadge(key, cards);
  ...
}
```

> `fileHeadHtml` 的其它调用点（如果有）需要一并传参；`fileGroups(batchState.cards)` 的分组顺序与原来 `filter` 的顺序一致（都按 `batchState.cards` 的原序），所以 `cards[0]`、`manual[0]` 等取值不变。

**`static/batch.js:491-505`** — `applySelection`：

```javascript
let lastSelectionKey = '';

function applySelection() {
  const selectionKey = JSON.stringify([...selectedIds].sort()) + '|' + renderKey;
  if (selectionKey === lastSelectionKey) return;
  lastSelectionKey = selectionKey;
  for (const element of document.querySelectorAll('.batch-card[data-id]')) {
    const selected = selectedIds.has(element.dataset.id);
    element.classList.toggle('is-selected', selected);
    element.querySelector('.batch-select')?.setAttribute('aria-pressed', String(selected));
  }
  const byFile = new Map(fileGroups(shownCards()));
  for (const button of document.querySelectorAll('.batch-file [data-file-action="select"]')) {
    const ids = (byFile.get(button.dataset.file) || []).map(c => c.id);
    const all = ids.length > 0 && ids.every(id => selectedIds.has(id));
    button.setAttribute('aria-pressed', String(all));
    button.textContent = all ? '取消选中本文件' : '选中本文件的卡';
  }
  renderSelectBar();
}
```

> **注意**：整页 `innerHTML` 重建后 DOM 是全新的，此时即使 `selectionKey` 没变也必须重跑。所以 `render()` 的 `else` 分支（`static/batch.js:769`）里要把 `lastSelectionKey = ''` 重置，或在重建后直接调用一次强制刷新。这一点务必在实现时处理，否则会出现「重建后勾选圈消失」。

**验证**：`node --test tests/test_batch_review.js`（含 `the selection circle toggles selection without changing color`、`select all and invert act on the cards in the current filter`、`batch color change sends one request and keeps the selection`）+ 真机点击。

---

### Step 6：ETag / 304

**`web_clip_fixer.py:104-120`** — 响应层支持 etag：

```python
    def _send_bytes(
        self, body: bytes, content_type: str, status: int = 200, filename: str | None = None,
        cacheable: bool = False, etag: str = "",
    ) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        if status != 304:
            self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "private, max-age=86400" if cacheable else "no-store")
        if etag:
            self.send_header("ETag", etag)
        if filename is not None:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        if status != 304:
            self.wfile.write(body)

    def _send_json(self, data: dict, status: int = 200, etag: str = "") -> None:
        body = b"" if status == 304 else json.dumps(data, ensure_ascii=False).encode("utf-8")
        self._send_bytes(body, "application/json; charset=utf-8", status=status, etag=etag)
```

**`core/batch.py`** — 缓存上次序列化结果：

```python
_SNAPSHOT_BODY: dict[str, Any] = {"etag": "", "body": b""}


def snapshot_json() -> tuple[str, bytes]:
    """返回 (etag, body)。内容没变时 etag 不变，调用方可以据此回 304"""
    body = json.dumps(snapshot(), ensure_ascii=False).encode("utf-8")
    etag = f'"{hashlib.sha1(body).hexdigest()}"'
    _SNAPSHOT_BODY.update(etag=etag, body=body)
    return etag, body
```

**`web_clip_fixer.py:169`**：

```python
        elif parsed.path == "/api/batch":
            etag, body = batch.snapshot_json()
            if etag == self.headers.get("If-None-Match"):
                self._send_json({}, status=304, etag=etag)
                return
            self._send_bytes(body, "application/json; charset=utf-8", etag=etag)
```

**`static/batch.js:133-155`** — `poll()`：

```javascript
let lastEtag = '';

async function poll() {
  const version = ++pollVersion;
  try {
    const response = await fetch('/api/batch', {
      cache: 'no-store',
      headers: lastEtag ? { 'If-None-Match': lastEtag } : {},
    });
    if (response.status === 304) return;          // 状态没变：不碰 batchState，不重画
    if (!response.ok) throw new Error(`无法读取进度：HTTP ${response.status}`);
    lastEtag = response.headers.get('ETag') || '';
    const value = await response.json();
    ...
```

> 任何 POST 操作后要把 `lastEtag = ''`，确保下一次轮询一定拿到新数据（`action()` / `post()` 的收尾处，`static/batch.js:126-131`、`237`）。虽然 etag 由内容算出、理论上不会错，但显式失效能避免「本地已乐观更新、服务端还没落盘」的边界情况。

**验证**：Network 面板看空闲时 304 命中率；真机走完整审核流程（切色 → 提交 → 复检 → 标色还原），确认每步页面都及时更新。

---

### Step 7：CSS 渲染隔离（可选，前置条件见 §3.8）

**`static/batch.css`**（`.batch-file` 规则处）：

```css
.batch-file {
  content-visibility: auto;
  contain-intrinsic-size: auto 620px;
}
```

同时改 `scrollAnchor`（`static/batch.js:560`），不再对视口外元素读 rect。

---

### Step 8：图片版本缓存（最低优先级）

**`core/fix_images.py`**：

```python
_VERSION_CACHE: dict[str, str | None] = {}


def _probe_version(name: str) -> str | None:
    """文件不存在记为 None。只有缓存未命中时才 stat 磁盘"""
    if not name or Path(name).name != name:
        return None
    if name in _VERSION_CACHE:
        return _VERSION_CACHE[name]
    try:
        stat = (IMAGES_DIR / name).stat()
    except OSError:
        _VERSION_CACHE[name] = None
        return None
    version = f"{stat.st_mtime_ns:x}{stat.st_size:x}"
    _VERSION_CACHE[name] = version
    return version


def image_path(name: str) -> Path | None:
    return (IMAGES_DIR / name) if _probe_version(name) is not None else None


def version(name: str) -> str:
    """截图文件会被同名覆盖（每次尝试都写同一个文件）：地址带上它，浏览器才不会用旧图"""
    return _probe_version(name) or ""
```

三个失效点：

```python
# save()：写入成功后（core/fix_images.py:36 附近）
        (IMAGES_DIR / name).write_bytes(base64.b64decode(data_url.split(",", 1)[1]))
        _VERSION_CACHE.pop(name, None)      # 覆盖写：让下次访问重新 stat，拿到新的 mtime

# clear()：删除文件后（core/fix_images.py:77、81 附近）
        _VERSION_CACHE.clear()              # 批量删除，逐个对账不值得

# import_all()：写入后（core/fix_images.py:102 附近）
            _VERSION_CACHE.pop(name, None)
```

> `save()` 用 `pop` 而不是直接写入新版本号，是为了保持 `version()` 的格式与现在完全一致（`mtime_ns + size`，不带进程内序号），避免重启后所有图片 URL 变化导致全部重新下载。

**验证**：`python -m unittest tests/test_report_images.py`、`tests/test_batch.py`（`:1934-1936` 断言两次尝试的 `images["before"]` 地址不同）；真机确认复检覆盖截图后页面图片会更新。

---

## 五、实施顺序与依赖

```
Step 0（基线）
   ↓
Step 1（后端全表扫描）──→ Step 2（后端浅拷贝/字段）      [同改 snapshot()，串行避免冲突]
   ↓
Step 3（前端记忆化）──→ Step 6（ETag/304）                [6 依赖 3 的签名划分思路]
   ↓
Step 4（高度测量）──→ Step 7（content-visibility，可选）  [硬依赖：7 会返回占位高度]
   ↓
Step 5（平方项清理）
   
Step 8（图片缓存）：独立，随时可做，但优先级最低
```

**建议的提交粒度**：一步一个提交，提交信息写明对应的编号（如 `性能：snapshot 复用 by_cell 索引，消除 O(N×M) 全表扫描（B1/B2/B3/B5）`）。

**分支**：从 `main` 新开分支（例如 `perf/workbench-render`）。当前 `main` @ `053aec8` 刚合入列宽共享记录的改动，性能优化不要和它混在一起。

## 六、验证与回归清单

### 6.1 自动化

```bash
python -m unittest discover -s tests -p "test_*.py"
node --test tests/test_batch_review.js
```

各步重点用例：

| 步骤 | 重点测试 |
|---|---|
| Step 1 | `tests/test_batch.py`、`test_mode_switch.py`（`_mode()` 无参调用）、`test_later_review.py` |
| Step 2 | `tests/test_batch.py:1257`、`1289`（snapshot 的 `context`）、`:750`（`adoptable`） |
| Step 3/4/5 | `tests/test_batch_review.js` 全部 21 个用例 |
| Step 6 | `tests/test_persistence.py`（HTTP 层） |
| Step 8 | `tests/test_report_images.py`、`tests/test_batch.py:1934-1936` |

### 6.2 真机手工清单（每步做完都过一遍）

准备：60~80 个文件、800~1000 张卡的文件夹。

- [ ] 全自动处理过程中，卡片随扫描逐张出现，页面不跳动、滚动位置保持
- [ ] 扫描中截图不重新加载（Network 面板看 `/api/batch/image/` 没有重复请求）
- [ ] 任务开始/结束的瞬间（`busy` 翻转），卡片上的「后台处理完这个文件后可标记」文案与按钮 disabled 状态立即更新
- [ ] 文件框上的「⏳ 排队等复检」「✅ 自动调整完成」「🤖 正在自动调整」徽章随任务推进更新
- [ ] 单击切色立即变色；连点两下能 黄→蓝→绿
- [ ] 勾选/取消勾选卡片、选中本文件、全选、反选，按钮按下态与文案正确
- [ ] 切色后提交、复检、自动调整的行为与优化前一致
- [ ] 切换颜色筛选、切换大图审核、缩放窗口后瀑布流三列仍然均衡
- [ ] 黄卡上的「🙅 不能自动调整：…」原因文案逐字不变（Step 1 的重点回归项）
- [ ] 进度条、剩余时间、日志滚动正常
- [ ] （Step 6）空闲时 Network 面板显示 304；点一下切色后立刻恢复 200

### 6.3 性能对比

用 Step 0 的同一张表，在 Step 1/2 后、Step 3/4/5 后、Step 6 后各测一次，逐列填写。任何一项没有改善就说明该步的设计假设错了，回到 §1 重新核对，不要继续往下做。

## 七、风险与回退

| 步骤 | 主要风险 | 检测方式 | 回退 |
|---|---|---|---|
| Step 1 | `by_cell.get(..., [])` 与 `_items_at()` 顺序不一致 → `auto_skip` 文案变化 | 手工比对若干黄卡文案 | 单提交 revert |
| Step 1 | `_unchanged_since_check` 新参数传错 → 「自动试过没通过」判断失灵，本该跳过的格子反复重试 | `tests/test_batch.py:1533-1540` | 单提交 revert |
| Step 2 | 浅拷贝污染 STATE（`context` 忘了单独拷） | 切色后再次轮询看 `context.image` 是否异常；`tests/test_batch.py:1257` | 单提交 revert |
| Step 2 | 删掉的字段其实有人读 | 已核对 `static/`、`templates/` 无引用；再跑一遍前端测试 | 把字段加回来 |
| Step 3 | `cardChromeSig` 漏了某个依赖 → 按钮状态不更新 | §6.2 第 3、4 项 | 把漏掉的字段加进签名（宁可过宽，过宽只是少省一点） |
| Step 4 | 高度不再全量重测 → 长期运行后列分配漂移 | 长时间扫描后目测三列高度 | 单提交 revert |
| Step 5 | `applySelection` 短路过头 → 整页重建后勾选圈丢失 | §6.2 第 5 项 | 去掉 `lastSelectionKey` 短路，只保留分组优化 |
| Step 6 | 浏览器把 304 规范化 → 前端拿不到数据，页面卡死 | Network 面板 | 退化为 `{"unchanged": true}` + 200（§3.7） |
| Step 6 | etag 相同但内容不同（哈希碰撞） | 理论上 SHA-1 碰撞可忽略 | — |
| Step 7 | `content-visibility` 破坏高度测量与锚点 | §6.2 第 8 项 + 滚动条稳定性 | 删掉两行 CSS（§3.8 回退判据） |
| Step 8 | 缓存失效点漏了 → 浏览器一直显示旧截图 | `tests/test_batch.py:1934-1936` + 真机复检 | 单提交 revert |

---

## 附录：核对方法说明

本文的每一条问题都经过实际读码核对，不是从旧文档转述：

- 后端调用链：`grep -n` 定位定义与全部调用点，逐个 `Read` 确认参数是否传递。
- 前端依赖分析：从 `cardHtml`（`static/batch.js:464`）出发，逐层展开 `canMark`、`openBlocked`、`markWhileBusy`、`autoQueueOpen`、`recheckButtonHtml`、`autoButtonHtml`、`autoBlockReason`、`queueBlockReason`、`queueNote`、`busyJobName`，确认卡片 HTML 究竟读了哪些全局状态（§1.3 的依赖表）。
- `stat()` 耗时：本机实测（方法见 §1.5），非引用文献值。
- 字段是否可删：`grep -rn "signatures\|item_ids" static/ templates/`（无结果）+ `grep -rn` 测试目录（确认测试读的是内部状态而非 snapshot 输出）。

**注意**：核对期间 `core/batch.py` 曾被并发修改（`snapshot()` 从 548 行移到 572 行），本文所有行号已在 `main` @ `053aec8` 的干净工作区上重新核实。
