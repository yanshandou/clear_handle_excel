# 「清除 Excel 标记颜色」迁入修复颜色页 —— 改动方案

> 状态：已实施（按下方方案落地，工作台按钮保留）
> 日期：2026-10-28

## 一、目标

把「切字检查工作台」里的「🧽 清除 Excel 标记颜色」能力，复制一份到「修复颜色」页（`/repair-colors`），
让该页面成为"颜色问题"的统一收口：

- **类型一（直接配色）**：走现有的"对照备份恢复"（不动）。
- **类型二（条件格式标记）**：新增"清除程序加的条件格式标记"（本次新增）。

工作台里的原按钮**本次保留不动**，等功能稳定后再决定是否删除。

## 二、核心结论（为什么不能直接复用工作台那条链路）

工作台清除走的是 `batch_unhighlight` job → `_sync_highlights(clear=True)`，
它依赖 `STATE` 里本次会话标过色的记录（`_marks()` / `_tabs()`）来决定"处理哪些文件、删哪些规则"。
**对一张在 STATE 里没有任何记录的文件，`_sync_highlights` 会直接早退、什么都不做。**

而「修复颜色」页的场景是：用户填一个文件夹，清除该文件夹下**所有**文件里程序加的条件格式标记。
路径来源是"文件夹全量扫描"，与 STATE 记录无关。

因此需要新增一个**不依赖 STATE 记录**的"全局清除"函数，复用底层删规则 / 还原标签的能力。

## 三、现有链路回顾（供对照）

| 能力 | 文件 | 关键函数 | 路径来源 |
|---|---|---|---|
| 加标记（转人工标色） | `core/batch.py` | `_add_rule()` | 工作台卡片 |
| 删标记规则 | `core/batch.py` | `_drop_rules(book, keep, …)` | 按 `CHE_MARK` 标识全局扫 |
| 还原标签色 | `core/batch.py` | `_sync_tabs()` / `_tab_is_ours()` / `TAB_HIGHLIGHT` | STATE `_tabs()` 记录 |
| 工作台清除 job | `core/batch.py` | `run()` 里 `batch_unhighlight` 分支 | `highlighted_paths()`（STATE 记录） |
| 修复颜色页 job | `core/mark_finder.py` | `repair()` / `_run_repair()` / `status()` | 文件夹扫描 |
| 路由 | `web_clip_fixer.py` | `/api/batch/unhighlight`、`/api/batch/find_marks*` | — |

关键复用点：`_drop_rules` 本身就是"按 `CHE_MARK` 标识全局扫"，**不依赖 STATE**；
`_tab_is_ours(sheet)` 只看标签色值是否等于 `TAB_HIGHLIGHT`（`#FFC000`），也不依赖 STATE。
所以"全局清除"只需把它们串起来。

## 四、后端改动

### 1. `core/batch.py` 新增 `clear_marks_in_file(app, path, report=None)`

一个不依赖 STATE 记录的"单文件全局清除"函数：

- 复用 `open_workbook_checked` 打开文件（`UpdateLinks=0`）。
- 调 `_drop_rules(book, keep=set(), tick, holders)`：删掉文件里**所有**带 `CHE_MARK` 标识的条件格式规则。
- 全局还原橙黄标签：遍历 `book.Worksheets`，凡是 `_tab_is_ours(sheet)` 为真的，
  执行 `sheet.Tab.ColorIndex = XL_COLOR_INDEX_NONE`（恢复成"无颜色"）。
  注意：没有 STATE 记录，无法还原"原色"，只能还原成"无颜色"——与工作台对"无记录但被标黄"的兜底一致。
- 统计并填 `report`：删了几条规则、几张标签、失败原因（工作表受保护等）。
- 复用 `close_saved_workbooks` 处理"文件还在 Excel 里开着"的情况。

> 放在 `batch.py` 的理由：`_drop_rules`、`_tab_is_ours`、`TAB_HIGHLIGHT`、`XL_COLOR_INDEX_NONE`、`open_workbook_checked` 都在这里。

### 2. `core/mark_finder.py` 新增 `clear_marks(folder)` + `_run_clear_marks(root, paths)`

复用「修复颜色」页已有的 job 机制（`_STATUS` + `_LOCK` + `list_workbooks` + 后台 Excel 启停）：

- `clear_marks(folder)`：
  - `list_workbooks(folder)` 拿到文件清单。
  - busy 检查（与 `repair()` 一致，避免和"对照备份恢复"同时跑）。
  - 置 `_STATUS`：`phase="clear"`、`stage="scan"`、`running=True`、`folder`、`total/done`。
  - 起线程跑 `_run_clear_marks`。
- `_run_clear_marks(root, paths)`：
  - 按需 `start_background_excel()`。
  - 逐文件调 `batch.clear_marks_in_file(...)`，收集结果到 `files`。
  - 收尾写 `_STATUS.applied = {"folder":…, "files": files}`，并复位 `running` / `wf.STATE["busy"]` / `maintenance`（照搬 `_run_repair` 的收尾）。

> 复用 `_STATUS` 的代价：清除标记与"对照备份恢复"共用同一把状态锁，两者不能同时跑——这是合理的。

### 3. `web_clip_fixer.py` 新增路由

- `POST /api/batch/clear_marks`：`mark_finder.clear_marks(_require_folder(payload))` → `{"ok": True}`。
- 轮询复用现有 `GET /api/batch/find_marks`（返回 `mark_finder.status()`），因为两者共用 `_STATUS`。

## 五、前端改动

### 4. `static/color_repair.html`

在「对照备份，检查全部配色差异」区块下方，新增一个独立卡片「🧽 清除 Excel 标记颜色」：

- 说明：一键删掉程序在切字工作台加的条件格式标记（`CHE_MARK` 规则）并还原被标橙黄的工作表标签；
  单元格直接设置的底色 / 字色不受影响（类型一请用上方"扫描并修复颜色"）。
- 复用「当前 Excel 文件夹」输入框 `repairFolder`（清除标记不需要备份文件夹）。
- 一个按钮：`🧽 清除标记颜色`（id 例如 `repairClearMarks`）。
- 结果复用现有 `repairResults` 结果区。

### 5. `static/repair.js`

- 新增 `clearMarks()`：校验 `repairFolder` 非空 → `POST /api/batch/clear_marks` → 复用 `refresh()` 轮询。
- 结果渲染分支：`latest.phase === 'clear'` 时，用"清除标记"的表头 / 文案渲染 `applied.files`（字段与现有 cells 类似：`sheet` / `addr` / `status` / `reason`）。
- 按钮 disabled 状态与现有 `controls()` 联动。

## 六、本次不改动

- `static/batch.html`、`static/batch.js`：工作台「🧽 清除 Excel 标记颜色」按钮保留。
- 现有"对照备份恢复"链路（`color_compare.py` / `restore_excel_backgrounds.py` / `mark_finder.repair()`）。

## 七、边界与风险

1. **标签色误伤**：`_tab_is_ours` 只看标签色 == `#FFC000`。用户若手动把标签设成这个橙黄，会被当程序标记清成"无颜色"。
   工作台本身也有此边界，可接受；必要时后续改成"带记录才还原"。
2. **原色降级**：全局清除没有 STATE 记录，标签只能还原成"无颜色"，而非"原色"。程序标的标签本来就橙黄，降级无碍。
3. **文件占用**：复用 `close_saved_workbooks` / `workbook_locked` 检查，开着的不动并提示。
4. **模板格式**：`.xltx / .xltm` 只读，参考现有逻辑跳过或提示另存。
5. **性能 / 卡死**：大文件逐条删规则可能慢。第一版复用 `start_background_excel` 简单模式（无 watchdog），
   与 `_run_repair` 一致；若后续发现卡死再引入 watchdog。

## 八、测试

- 新增 `tests/test_clear_marks.py`（或扩展 `tests/test_scan_only.py` 已有的 unhighlight 用例）：
  1. 带 `CHE_MARK` 标记的文件 → 清除后规则删除、橙黄标签还原为无颜色。
  2. 无标记文件 → 无变化、不报错。
  3. 受保护工作表 → 记录失败原因，不影响其他文件。
  4. 文件被 Excel 占用 → 报错跳过，不崩。

## 九、实施顺序

1. `core/batch.py`：新增 `clear_marks_in_file`。
2. `core/mark_finder.py`：新增 `clear_marks` + `_run_clear_marks`。
3. `web_clip_fixer.py`：新增 `/api/batch/clear_marks` 路由。
4. `static/color_repair.html` + `static/repair.js`：前端卡片 + 按钮 + 轮询渲染。
5. 测试 + 手动验证（造一份带标记的文件跑一遍）。
6. 稳定后，再决定是否删除工作台按钮（另行确认）。
