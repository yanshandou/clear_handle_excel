# `###` 列宽不足自动加宽：代码审查报告与修改意见

> **修复状态（2026-10-27）**：下文保留的是修复前的审查记录。问题 1–7 已落实代码修改；列宽采用方案 A，按文件、工作表和列共用记录。同列及部分重叠的合并区域一起审核，混色时还原相关列并转人工，按颜色提交时跳过混色组。改名、旧记录迁移、待续跑日志、审核接受状态和单列容差核对均有回归用例。真实 Excel 中的数字格式、日期显示与列宽像素取整仍需实测。

本次修复验证：

* 完整 Python 回归套件：**599 / 599 通过**，无失败、错误或跳过。
* 前端 `node --test tests/test_batch_review.js`：**38 / 38 通过**。
* 此前在修复前提交 `b40f31e` 上同样复现的 3 个测试失败也已修复：两个文件占用用例现在同时模拟锁文件和实际占用，并验证关闭后续跑、残留锁不阻塞；日志清理修复 Windows 下 CRLF 被再次转换而产生空行的问题（旧日志已删除，并非两条旧记录均被保留）。另补齐批处理、模式切换和稍后审核测试的后台报告落盘清理，避免临时 SQLite 数据库清理时仍被占用。

## 一、审查范围与结论

### 1.1 审查对象

修复前提交 `b40f31e` 中的列宽功能改动（审查时尚未提交），对照 [workbench_performance_analysis.md](./workbench_performance_analysis.md) 进行核对：

| 文件 | 改动概要 |
|---|---|
| [core/autofix.py](./core/autofix.py) | `hash_overflow` 改为 `fixable: True` |
| [core/batch.py](./core/batch.py) | 新增 `_col_rid` / `_cols_match` / `_cell_columns` / `_prepare_column` / `_repair_column`；`_row_value`、`_write_rows`、`_finish_pending`、`_repair_file`、`_commit_review`、`snapshot`、`_auto_skip_reason` 支持列宽记录 |
| [static/batch.js](./static/batch.js) | `### 列宽不足` 改为可勾选；卡片「处理」说明支持列宽 |
| [README.md](./README.md) | 说明改为支持 `###` 自动加宽列 |
| [tests/test_batch.py](./tests/test_batch.py) 等 | 更新 / 新增对应测试 |

### 1.2 总体结论

* 这次 diff 实现的是文档 **第五节（`###` 列宽不足自动调整与还原）**。**第一至第四节的性能优化方案均未包含在本 diff 中**（没有 `_VERSION_CACHE`、`_can_highlight(card, items)`、前端增量渲染、`content-visibility` 等），文档第四节的「优化后」对比表目前与代码不符。
* 第五节的**主体流程与文档描述一致**，但存在 **2 个高危问题、2 个中危问题、若干低危问题**，**暂不建议直接合并**。
* 这些问题**与第一至第四节未实现无关**。性能优化不改变功能语义，即使全部实现，以下问题仍然存在。问题来源分为三类：
  * **代码写错**：问题 1；
  * **文档设计本身有缺陷**：问题 2（文档写明按单元格建立 `_col_rid(card)`，但列宽是整列共用的）；
  * **集成时漏改**：问题 3、4、5、6（原来只认识「行高 / 分页」两种记录的逻辑，没有同步支持新的「列宽」记录）。

### 1.3 审查时（修复前）的验证情况

| 项目 | 结果 |
|---|---|
| `node --test tests/test_batch_review.js` | ✅ 35 / 35 通过 |
| `python -m unittest tests.test_batch tests.test_mode_switch` | ⚠️ 未能运行：本机 PyMuPDF 的 `_mupdf` DLL 被系统「应用程序控制策略」拦截，导入 `fitz` 失败 |
| 结论来源 | Python 部分结论来自代码阅读与调用链追踪，需在可运行环境中补跑测试，并按第五节补充测试用例 |

---

## 二、与文档第五节一致、实现正确的部分

| 文档要求 | 代码位置 | 结论 |
|---|---|---|
| `hash_overflow` 可勾选、默认选中，只剩 `text_missing` 固定转人工 | `autofix.CELL_ERROR_SPECS`、`batch.js` 的 `CELL_ERROR_TYPES` | ✅ |
| 单列每步 `+2.0`，跨 K 列每列 `max(0.5, round(2.0 / K, 2))` | `_prepare_column` | ✅ |
| 仍为 `###` 时不导出 PDF，直接继续加宽 | `_prepare_column` 读 `sheet.Cells(row, cols[0]).Text` | ✅ |
| 显示恢复后追加余量 `max(0.3, round(1.0 / K, 2))`，再导出 PDF 用 `autofix._evaluate` 验证 | `_prepare_column` | ✅ |
| 副本试算结束后还原副本列宽 | `_prepare_column` 的 `finally` | ✅ |
| 独立记录 `kind="col"`，保存 `base_cols` / `applied_cols` | `_repair_column` | ✅（但见问题 2） |
| 写入前核对各列宽度，容差 `COL_WIDTH_TOLERANCE = 0.4`，支持续跑 | `_cols_match` / `_write_rows` | ✅（但见问题 6） |
| 提交标绿 / 标黄时还原原列宽 | `_commit_review` | ✅（但见问题 2、4） |
| 未勾选 `###` 时转黄色人工卡，提示「未调整列宽」 | `_unchecked_fix_reason`、`_repair_file` | ✅ |
| 写入成功后卡片标蓝、问题项记为 `fixed` | `_finish_pending` 新增 `col` 分支 | ✅ |

---

## 三、问题清单

| # | 严重度 | 位置 | 问题 | 原因分类 |
|---|---|---|---|---|
| 1 | 🔴 高 | `core/batch.py` 第 1280 行 `_repair_file` | 列宽写入阶段的 `ValueError` 被吞掉，写入日志 `pending_adjustment` 会被后续写入覆盖丢失 | 代码写错 |
| 2 | 🔴 高 | `core/batch.py` 第 276 行 `_col_rid` 及其所有使用处 | 列宽修改记录按单元格建立，但列宽是整列共用的；同列多个 `###` 格子提交还原时核对失败或还原错误 | 文档设计缺陷 |
| 3 | 🟠 中 | `core/batch.py` 第 150–155 行（工作表改名迁移） | 列宽记录的编号被算成行高记录编号，可能覆盖同一行的行高记录，且之后再也找不到这条列宽记录 | 集成漏改 |
| 4 | 🟠 中 | `core/batch.py` 第 1559–1563 行 `_commit_review` 收尾 | 标蓝并提交后，列宽记录仍然是 `saved=True`，之后会用过期的原列宽去还原 | 集成漏改 |
| 5 | 🟡 低 | `batch.py` 第 862 行、第 1415 行；`autofix.py` 第 194/201 行；`_write_rows` 报错文案 | 文案写「加高 / 行高」；同行冲突判断不区分列宽 | 集成漏改 |
| 6 | 🟡 低（需实测） | `core/batch.py` `_write_rows` 中的 `at_target` 和 `expected` 判断 | 列宽记录用「各列宽度之和」与 `ROW_HEIGHT_TOLERANCE` 比较，合并的列很多时可能误判为「已被修改」 | 集成漏改 |
| 7 | ⚪ 极低 | `core/batch.py` 第 996 行 `_cell_columns` | 冗余的类型判断 | 代码风格 |

---

## 四、问题详解与修改意见

### 问题 1 🔴：列宽写入失败时异常被吞，写入日志丢失

#### 问题代码（`_repair_file`，第 1266–1284 行）
```python
try:
    workbook = _repair_column(excel, session, path, group, card, version, timing)
    version = wf._file_version(Path(path))
except ValueError as err:
    with wf.STATE_LOCK:
        card.update(color="yellow", verified=False, manual=True, waiting=False, message=str(err))
        wf.log_msg(...)
        wf._persist_state()
```

#### 原因分析
* `_repair_column` 内部会先把 `job["pending_adjustment"]` 写好并持久化，再调用 `_finish_pending` → `_write_rows`。
* `_write_rows` 发现文件被外部修改、或列宽已被人改过时，会抛出 **`ValueError`**（「文件在准备后被修改……」「行高已被其他操作修改……」）。这时原 Excel 有可能已经保存，也有可能还没保存，必须保留写入日志，后续靠「继续上次任务 / 放弃这次写入」来收尾。
* 新代码的 `except ValueError` 把这类错误和「试算不通过」的 `ValueError` 一起吞掉了：
  1. `pending_adjustment` 没有被清除，也没有被 `_park_pending` 隔离；
  2. 继续处理同一行的其他格子时，行高修复会**直接覆盖** `job["pending_adjustment"]`，列宽的写入日志就丢了；
  3. 卡片被标黄，但对应的 `rows` 记录状态没有收尾。
* 原有的行高 / 分页路径在外层 `except Exception` 中有明确约束：
  ```python
  # 有写入日志时不能吞掉异常：可能原文件已经保存，必须按日志续跑。
  if job.get("pending_adjustment"):
      raise
  ```
  列宽路径没有遵守这条约束。

#### 修改意见
在 `except ValueError` 里先检查写入日志，有日志就向外抛出，交给外层已有的处理逻辑：
```python
try:
    workbook = _repair_column(excel, session, path, group, card, version, timing)
    version = wf._file_version(Path(path))
except ValueError as err:
    if job.get("pending_adjustment"):
        raise  # 写入阶段出错：原文件可能已保存，必须保留写入日志按日志续跑
    with wf.STATE_LOCK:
        card.update(color="yellow", verified=False, manual=True, waiting=False, message=str(err))
        wf.log_msg(f"⚠️ {first['wb_name']} / {first['sheet_name']} 行{first['excel_row']}：{err}")
        wf._persist_state()
```
抛到外层后，会走到 `except Exception` 中的 `if job.get("pending_adjustment"): raise`，再由文件级流程 `_park_pending` 处理，和行高路径的行为一致。

#### 建议补充测试
* 用 `patch` 让 `_write_rows` 抛出 `ValueError("文件在准备后被修改")`。断言 `_repair_file` 向外抛出异常，并且 `job["pending_adjustment"]` 仍然是这条列宽写入日志（`kind == "col"`）。

---

### 问题 2 🔴：列宽修改记录按单元格建立，但列宽是整列共用的

#### 问题代码
```python
def _col_rid(card: dict) -> str:
    return _key(card["wb_path"], card["sheet_name"], card.get("cell_addr", ""), "col")
```
文档第五节同样写明「使用独立记录标识 `_col_rid(card)`」，所以这是**设计层面的问题**。

#### 触发场景
同一张表的 `C5`、`C9` 都显示 `###`：

1. 处理第 5 行：C 列 `8.0 → 11.3`，写入并保存，记录 `R5 = {base_cols: {C: 8.0}, applied_cols: {C: 11.3}}`。
2. 重新检测整表后，`C9` 的数字更长，仍是 `###`。处理第 9 行：C 列 `11.3 → 15.3`，记录 `R9 = {base_cols: {C: 11.3}, applied_cols: {C: 15.3}}`。
3. 提交审核时：

| 用户标记 | 结果 |
|---|---|
| C5 绿、C9 绿 | 两条还原放进同一份写入日志。`R5` 期望当前宽度为 11.3，实际是 15.3，`_cols_match` 失败，报「已被其他操作修改」，整份文件被挂起 |
| C5 绿、C9 蓝 | `R5` 期望 11.3，实际 15.3，同样失败 |
| C5 蓝、C9 绿 | `R9` 还原到 11.3，C5 的修改保留，结果正确，但只是碰巧 |

跨多列的合并单元格（如 `B3:D3` 和 `C7:E7`）列范围部分重叠时，问题相同。

行高为了避免同一类问题，用 `_rid` 让**同一行共用一条记录**，同行颜色不一致时整行转黄（`_linked_row`）。列宽需要同样的「共用资源」语义。

#### 修改意见

**方案 A（推荐，语义与行高一致）：按「表 + 列」建立列宽记录**

1. 记录编号改为按列：
   ```python
   def _col_key(path: str, sheet: str, col: int) -> str:
       return _key(path, sheet, int(col), "col")
   ```
2. 每一列一条记录，保存 `{"kind": "col", "path", "sheet", "col", "base", "applied", "saved", "cards": [...]}`。其中 `base` 是**第一次自动加宽前**的原始宽度；后续再加宽同一列时只更新 `applied`，不覆盖 `base`。
3. 卡片上记下自己影响了哪些列，例如 `card["col_ids"] = [...]`。`_row_record` / `snapshot` 根据 `col_ids` 汇总展示。
4. `_commit_review`：
   * 按列分组。一列上关联的所有卡片**都不是蓝色**时才还原该列，还原到 `base`。
   * 同一列关联的卡片颜色不一致时，参照 `_linked_row`，把这些卡片整体转黄并还原，提示「同列存在不同处理意见，整列还原后交给人工」。
5. 写入日志里的 `rows` 改为每列一项，`expected` / `target` 是单列宽度，不再用宽度之和（同时解决问题 6）。

**方案 B（最小改动）：同一列已有未提交的自动加宽时，不再自动加宽**

在 `_repair_file` 的列宽分支里，调用 `_repair_column` 之前先检查：当前格子的列范围如果和已有 `saved=True` 的列宽记录（同 path、同 sheet）重叠，就不再自动加宽，直接转黄：
```python
overlap = any(
    r.get("kind") == "col" and r["saved"] and r["path"] == path and r["sheet"] == sheet_name
    and set(r.get("applied_cols") or {}) & {str(c) for c in cols}
    for r in _data()["rows"].values()
)
if overlap:
    card.update(color="yellow", verified=False, manual=True, waiting=False,
                message="同一列已有自动加宽、尚未审核提交，未再次调整列宽，请提交后复检或人工处理")
    continue
```
方案 B 能避免还原出错，代价是同一列只会自动修第一个 `###` 格子。

**同时需要更新** [workbench_performance_analysis.md](./workbench_performance_analysis.md) 第五节第 3 条中「使用独立记录标识 `_col_rid(card)`」的描述。

#### 建议补充测试
* 同一列两个 `###` 格子先后加宽，分别标「绿 + 绿」「绿 + 蓝」「蓝 + 绿」后提交。断言最终列宽符合预期，并且不会把文件挂起。
* 两个范围部分重叠的合并单元格（如 `B3:D3`、`C7:E7`）做同样的断言。

---

### 问题 3 🟠：工作表改名时列宽记录编号被算错

#### 问题代码（第 150–155 行）
```python
for rid, row in list(data["rows"].items()):
    if row["path"] == path and row["sheet"] == old:
        row["sheet"] = new
        row["id"] = (_key(path, new, int(row["row"]), "break") if row.get("kind") == "break"
                     else _key(path, new, int(row["row"])))
        data["rows"][row["id"]] = data["rows"].pop(rid)
```

#### 原因分析
* `kind == "col"` 的记录会走 `else` 分支，被算成**行高记录的编号** `_key(path, new, row)`。
* 后果：
  1. 同一行如果还有行高记录，两条记录编号相同，后放进去的会**覆盖**先放进去的，丢失一条修改记录；
  2. 改名后 `_col_rid(card)` 按新表名计算的是 `_key(path, new, cell_addr, "col")`，与迁移后的编号不一致，这条列宽记录**再也找不到**。提交标绿时不会还原列宽，而 `_row_record` 有可能取到错的记录。

#### 修改意见
迁移时区分三种记录（采用方案 A 时改为按列计算编号）：
```python
kind = row.get("kind")
if kind == "break":
    row["id"] = _key(path, new, int(row["row"]), "break")
elif kind == "col":
    row["id"] = _key(path, new, row.get("cell_addr", ""), "col")  # 方案 A：_col_key(path, new, row["col"])
else:
    row["id"] = _key(path, new, int(row["row"]))
```
同时检查 `job["pending_adjustment"]` 和 `job["deferred_adjustments"]` 中 `rows[*]["id"]` 是否也需要同步迁移（这是原有逻辑的问题，但列宽记录让风险变大了）。

#### 建议补充测试
* 已加宽列（`saved=True`）后把工作表改名，再标绿提交。断言列宽被还原，`data["rows"]` 中同一行的行高记录没有被覆盖。

---

### 问题 4 🟠：标蓝提交后列宽记录未标成「已接受」

#### 问题代码（`_commit_review` 收尾，第 1559–1563 行）
```python
with wf.STATE_LOCK:
    for rid, peers in by_row.items():
        row = data["rows"].get(rid)
        if row and row["path"] == path and all(c.get("dismissed") and c["color"] == "blue" for c in peers):
            row["saved"] = False  # 已正式接受的历史修改不再属于可还原的待审核修改。
```

#### 原因分析
* 这里只处理按 `_rid(card)`（行）查到的行高记录。列宽记录的编号是 `_col_rid(card)`，**永远不会**被设成 `saved=False`。
* 后果：
  1. 已经正式接受的列宽修改一直被当作「待审核、可还原」的修改；
  2. 同一格之后又出现 `###` 卡片（卡片编号由单元格算出，与原来相同），标绿提交时会用**过期的 `base_cols`** 去还原。文件已被人改过时，会核对失败并挂起；核对恰好通过时，会把列宽改回很早以前的值；
  3. `_repair_file` 的 `manual_rows` 和 `_handed_over_paths` 中的「已保存记录」判断也会受到影响。

#### 修改意见
在同一个收尾块里一并处理列宽记录：
```python
with wf.STATE_LOCK:
    for rid, peers in by_row.items():
        row = data["rows"].get(rid)
        if row and row["path"] == path and all(c.get("dismissed") and c["color"] == "blue" for c in peers):
            row["saved"] = False
    for card in cards:
        if card["wb_path"] != path or not (card.get("dismissed") and card["color"] == "blue"):
            continue
        col_row = data["rows"].get(_col_rid(card))
        if col_row and col_row["path"] == path:
            col_row["saved"] = False  # 已正式接受的列宽修改不再属于可还原的待审核修改
    job["completed"].append(path)
    wf._persist_state()
```
采用方案 A 时，改为「一列关联的卡片全部是蓝色且已收起」才设 `saved=False`。

#### 建议补充测试
* 加宽列 → 标蓝提交，断言列宽记录 `saved is False`。之后同一格再出 `###` 卡并标绿提交，断言不会按旧的 `base_cols` 还原。

---

### 问题 5 🟡：文案与同行冲突判断沿用行高写法

| 位置 | 现状 | 修改意见 |
|---|---|---|
| `batch.py` 第 859–863 行（重新检测后自动标蓝） | 列宽卡片显示「已加高并保存，最新导出已通过检查」；用 `ROW_HEIGHT_TOLERANCE` 比较「列宽之和」 | 按 `row.get("kind")` 区分：`col` 显示「已加宽列并保存……」，用 `COL_WIDTH_TOLERANCE` 判断 |
| `autofix.validate_print_targets`（`autofix.py` 第 194、201 行） | 列宽路径也报「未调整行高，请人工核对」 | 增加参数，例如 `action: str = "未调整行高"`，`_prepare_column` 调用时传入 `"未调整列宽"` |
| `_write_rows` 中的报错 | 列宽核对失败时报「行高已被其他操作修改」 | 写入日志中含 `kind == "col"` 时改为「列宽已被其他操作修改」 |
| `abandon_pending` 文案 | 「请在 Excel 中核对这一行的行高」 | 列宽记录改为「核对这一列的列宽」 |
| `_linked_row`（第 1415 行） | 只看行高记录。同一行一格「列宽标蓝」、一格「行高标绿」时，整行转黄，列宽也被一起还原 | 列宽不属于「同一行共用」的资源，`by_row` 冲突判断应排除列宽卡片；列宽的冲突改为按列判断（见问题 2 方案 A） |

---

### 问题 6 🟡（需实测）：续跑核对用「列宽之和」与行高容差比较

#### 问题代码（`_write_rows`）
```python
at_target = (all(abs(h - r["target"]) < ROW_HEIGHT_TOLERANCE and _cols_match(book, r, "cols_target")
                 for h, r in zip(actual, journal["rows"])) ...)
...
if any(abs(h - r["expected"]) >= ROW_HEIGHT_TOLERANCE or not _cols_match(book, r, "cols_expected")
       for h, r in zip(actual, journal["rows"])):
    raise ValueError(...)
```
其中，列宽记录的 `h = _row_value(book, r)` 是**各列宽度之和**。

#### 原因分析
* Excel 写入 `ColumnWidth` 后会按像素取整，读回来的值会有一点偏差。单列偏差很小，`_cols_match` 已经按单列使用 `COL_WIDTH_TOLERANCE = 0.4` 判断。
* 但跨 K 列合并时，K 列的偏差**加在一起**，再和 `ROW_HEIGHT_TOLERANCE = 0.4` 比较，K 较大（如日企方眼纸常见的 6–10 列合并）时可能超过 0.4。结果是：
  * `at_target` 误判为 `False`，于是检查 `expected`，同样因为求和误差误判，报「已被其他操作修改」，文件被挂起。

#### 修改意见
列宽记录只按单列核对，不再用求和值比较：
```python
def _at(book: Any, h: float, r: dict, field: str, value_key: str) -> bool:
    if r.get("kind") == "col":
        return _cols_match(book, r, field)
    return abs(h - r[value_key]) < ROW_HEIGHT_TOLERANCE

at_target = (all(_at(book, h, r, "cols_target", "target") for h, r in zip(actual, journal["rows"])) and ...)
...
if any(not _at(book, h, r, "cols_expected", "expected") for h, r in zip(actual, journal["rows"])):
    raise ValueError(...)
```
`_row_value` 的求和结果只用于日志展示。

#### 建议验证
* 在真实 Excel 中对 8 列合并单元格写入加宽后的列宽并读回，确认单列偏差与总和偏差，必要时调整 `COL_WIDTH_TOLERANCE`。

---

### 问题 7 ⚪：`_cell_columns` 冗余判断

```python
c0, cnt = int(area.Column), int(area.Columns.Count)
if not isinstance(c0, bool) and not isinstance(cnt, bool) and isinstance(c0, int) and isinstance(cnt, int):
```
`int(...)` 的结果一定是 `int`，不会是 `bool`，类型判断可以删掉，只保留 `if c0 >= 1 and cnt >= 1:`。不影响功能。

---

## 五、其他需要实测确认的风险（未定性为缺陷）

1. **`_evaluate` 对 `###` 项的文字比对**：`###` 检测项的 `text` 由 `hidden_text._guess_display(cell["value"])` 推测得到，而 `_evaluate` 对 `kind == "hidden"` 的项要求 `normalize_text(evidence["text"]) == normalize_text(item["text"])`。日期、自定义数字格式、千分位等情况下，推测的显示文字可能与加宽后 PDF 实际印出的不一致，导致加宽成功却判为不通过，最终转人工。建议用真实样例（日期、货币、百分比、自定义格式）验证。
2. **合并单元格不是从左上角格子报出时**：`_prepare_column` 用 `sheet.Cells(row, cols[0]).Text` 判断是否仍为 `###`。`row` 取自 `cell_addr`，如果 `cell_addr` 不是合并区域的左上角格子，读到的可能是空字符串，每一步都会导出 PDF（只影响速度，不影响结果）。建议改为读 `sheet.Range(addr).MergeArea.Cells(1, 1).Text`。
3. **人工改列宽的检测**：行高路径通过 `_remember_handover` / `TOUCHED_MESSAGE` 识别「行高被人改过」。列宽目前没有对应机制，人工改了列宽以后只能靠写入核对失败才发现。建议确认这一点是否符合预期。

---

## 六、与性能分析文档的关系

* [workbench_performance_analysis.md](./workbench_performance_analysis.md) 第一至第四节（后端图片版本缓存、消除 `snapshot()` 中 O(N×M) 遍历、前端增量 DOM、`content-visibility`）**在本次 diff 中均未实现**。
* 本报告列出的问题全部属于第五节的功能正确性，**与性能优化是否实现无关**。
* 需要注意：如果之后实现第二节「用轻量字典浅拷贝替代 `copy.deepcopy(card)`」，需要保证 `snapshot()` 新增的 `col_width` 字段（`base` / `applied` / `saved` / `cols`）照常输出，否则 `static/batch.js` 中 `handlingText` 的列宽说明会消失。

---

## 七、修改优先级与检查清单

| 优先级 | 事项 | 依赖 |
|---|---|---|
| P0 | 问题 1：列宽写入阶段的 `ValueError` 有写入日志时向外抛出 | 无 |
| P0 | 问题 2：确定列宽记录按列共用（方案 A）或同列不重复加宽（方案 B），并更新文档第五节 | 需先定设计 |
| P1 | 问题 4：标蓝提交后列宽记录设为 `saved=False` | 问题 2 的方案决定具体写法 |
| P1 | 问题 3：改表名迁移支持 `kind == "col"` | 问题 2 的方案决定编号算法 |
| P2 | 问题 6：续跑核对只按单列判断 | 方案 A 下自然解决 |
| P2 | 问题 5：文案与同行冲突判断区分列宽 | 问题 2 |
| P3 | 问题 7：清理冗余判断 | 无 |
| P3 | 第五节中的三项实测确认 | 需要真实 Excel 环境 |

### 测试清单
- [ ] 写入阶段抛出 `ValueError` 时，`pending_adjustment` 被保留并向外抛出
- [ ] 同一列多个 `###` 格子，「绿 + 绿 / 绿 + 蓝 / 蓝 + 绿」提交后列宽正确
- [ ] 范围部分重叠的合并列，提交还原正确
- [ ] 工作表改名后标绿提交，列宽被还原，行高记录未被覆盖
- [ ] 标蓝提交后列宽记录 `saved is False`，之后同一格再出卡标绿不会用旧值还原
- [ ] 8 列合并单元格加宽后续跑，不会误报「已被其他操作修改」
- [ ] 在 PyMuPDF 能正常导入的环境中补跑 `python -m unittest tests.test_batch tests.test_mode_switch`
