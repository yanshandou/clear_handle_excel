# Excel 批处理重复开存优化分析报告

> 分析日期：2026-11-11  
> 代码基线：本地 HEAD `16a1f63`，结合分析时的实际工作树；并非仅对已提交版本的审查。  
> 范围：核对优化建议、分析约束、提出实施和验收建议；本次不修改业务代码、不启动 Excel 处理用户文件。  
> 证据性质：打开次数来自实际调用路径的静态计数；时间收益是条件估算，不是性能实测。

## 1. 结论摘要

**“减少同一工作簿的重复打开与保存”是成立的优化方向，但原建议不能直接当作实施方案。**

主要结论如下：

1. **精简 Sheet 后自动检查确实存在较多重复开存。** 全部相关选项启用、每个修改分支均命中且没有重试时，当前文件可发生 **8 次 Excel COM 打开**，不是 7 次；参考文件另有 **2～3 次**，不应混在当前文件次数里。
2. **上传预处理确实可以集中写回。** 对普通 `.xlsx/.xlsm`，有颜色需要恢复时，常见路径是当前文件 3 次 COM 打开、参考文件 1 次，外加两次 openpyxl 解析；“每份 4 次”若指两份文件合计大致成立，若指当前文件则不准确。
3. **“只开一次”不等于“把现有函数放在一个循环里”。** 当前各函数自己负责打开、关闭、保存，还依赖磁盘版本、工作表历史、参考文件模拟和任务状态，需要拆开资源管理与具体操作。
4. **同名参考文件不能在同一个 Excel 实例中与当前文件同时打开。** 对引用修复，参考文件还会在内存中被删表、改名以模拟损坏，不能当成始终不变的共享工作簿。
5. **最后一次保存不等于完整事务。** 能避免正常执行异常前主动保存部分修改，但不能自动保证保存失败、进程被终止、旧版 `.xls` 写锁副作用以及数据库写入失败时原件和记录完全一致。
6. **“一项失败整份停止”是行为变更。** 当前精简自动检查是一步失败后继续后面的步骤；“无法确认所以不改”和真正的执行异常必须区分，不能把所有提示都升级为整份失败。
7. **建议先建立基线，再以⑤上传预处理验证集中开存，再推进①精简链路。** 这是按改动范围与验证难度排序，不是断言⑤能节省更多时间；若真实日志证明①占绝对大头，也可调整实施顺序。
8. **组合截图与标注、宽一页不要与第一轮 I/O 重构无条件捆绑。** 它们会改变 Excel 图形结构或打印设置，是独立的行为扩展，应分别验收。

## 2. 分析口径与边界

### 2.1 “打开一次”必须拆成不同指标

| 指标 | 含义 | 是否启动 Excel | 是否一定修改文件 |
| --- | --- | --- | --- |
| Excel 实例启动 | 创建一个后台 `Excel.Application` | 是 | 否 |
| 当前文件 COM 打开 | 对当前文件调用 `Workbooks.Open` | 通常复用实例 | 否 |
| 参考文件 COM 打开 | 对参考文件调用 `Workbooks.Open` | 通常复用实例 | 否 |
| openpyxl 解析 | Python 读取工作簿数据结构 | 否 | 本文涉及的读取不写回 |
| ZIP/XML 读取 | 读取工作表、列定义等 XML | 否 | 否 |
| `Save()` | 将工作簿内容保存到当前路径 | 否 | 是 |
| 临时 `SaveAs()` | 为解析或验证生成临时工作簿 | 否 | 写临时文件，不等于保存原件 |
| 文件复制、哈希、替换 | 文件系统读写 | 否 | 视操作而定 |

不能把这些操作全部当成“打开并保存一次”，再统一乘以同一个耗时。

项目实际写回主要使用 **win32com / Excel COM**，部分扫描、颜色计划和字体预检使用 **openpyxl / ZIP/XML**，不是全部操作都经过 Excel。

证据：

- [core/excel_ops.py:50 — open_workbook_checked](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/excel_ops.py#L50)
- [core/excel_ops.py:61 — start_background_excel](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/excel_ops.py#L61)
- [core/color_compare.py:121 — _plan_xml](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/color_compare.py#L121)
- [core/font_fix.py:42 — yu_cells](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/font_fix.py#L42)

### 2.2 本报告次数的适用条件

除特别说明外，次数均指：

- 单份业务文件的一次完整尝试；
- 不含失败重试、暂停后的重新执行；
- 不含用户在 Excel 中手动打开；
- 不含进入执行前的页面预览、选表扫描；
- COM 打开与 Python 解析分别计算；
- 临时副本承接当前文件修改时，归入“当前文件处理侧”，不误写为原路径被再次打开；
- 保存次数指业务修改的保存，临时格式转换另列。

**因此，“最多 8 次”是限定路径的上界，不是包含任意重试后的绝对上界。**

### 2.3 本次做了什么、没有做什么

已完成：

- 阅读主要流程及下游操作函数；
- 核对打开、保存、关闭及异常分支；
- 检查相关已有测试，识别已有行为约束；
- 检查仓库日志的覆盖范围；
- 复算原建议的时间算术。

未执行：

- 572 份真实 Excel 的性能基准；
- Excel COM 实机操作与故障注入；
- 自动化测试套件；
- 业务代码修改、部署或提交。

分析开始时工作树已有前端等未提交修改，期间还出现了后端路径范围适配等并行修改；本报告不改动这些现有工作，交付前已复核相关差异并更新引用位置。源码后续仍可能变化，请同时按所列函数名定位，不将此文档视为动态更新的代码索引。

## 3. 对原建议的逐项判断

| 原建议 | 判断 | 需要补充或修正 |
| --- | --- | --- |
| ①精简加自动检查是主要合并候选 | 成立 | 重复开存明确存在，但是否为真实耗时“大头”还需要计时 |
| ①同一份文件最多开 7 次 | 计数需修正 | 全选且各修改分支命中时，当前文件 COM 打开可达 8 次；参考另计 |
| ①当前可写一次、参考只读一次即可完成 | 目标有条件成立 | 同名冲突、参考模拟损坏、历史投影和未保存状态都需要处理 |
| ②调整页面已经恒为一次打开 | 仅主路径成立 | 缓存命中 0 次；主路径 1 次；稀疏页修复可能再开一次 |
| ③只有需要标色的文件才会第二次打开 | 不完整 | 清旧标记、核对标记、恢复标签及补存视图也可能触发可写打开 |
| ③扫描先不合并 | 赞同 | 当前检测阶段有专门的只读和超时隔离，不应为省一次 Open 直接去掉 |
| ④修改与复检本来就是两件事，没有可合并空间 | 前半句成立，后半句过强 | 必须保留验证阶段，但逻辑分离不必然要求物理重开；还有调整页面 PDF 重导分支 |
| ⑤上传预处理值得合并 | 成立 | 原函数有自行开存及部分成功/跳过语义，需重构而非直接拼接 |
| ⑤每份约开 4 次 | 取决于口径和格式 | 普通 XML 格式有修改时：当前 3 次、参考 1 次；二进制路径可能合计 6 次 |
| ⑥同名 PDF 与 JSON 复制都只开一次 | 分类不准确 | 同名 PDF 用 Excel，只读 1 次；JSON 复制是 Excel COM 0 次 |
| ⑥同名 PDF 已没有可省 | 过于绝对 | 独立同名 PDF 入口目前记录缓存，但没有在此入口命中后跳过导出 |
| ①能省一个半小时以上、⑤半小时左右 | 未被证实，且算术表述不准 | 原假设下分别为约 76～191 分钟、38～76 分钟，不是实测承诺 |
| 任一步出错，其他项不会半改半存 | 必须定义保证边界 | 最后统一 Save 只能提供有限的失败隔离，强保护还涉及临时副本与提交恢复 |

## 4. ①精简 Sheet 与精简后自动检查

### 4.1 当前不是“每份文件走完整条流水线”

当前流程大体是：

```text
先对目标文件执行精简 Sheet
    ↓
精简任务收尾
    ↓
对文件夹执行 BL 扫描，再清除命中的文件
    ↓
对文件夹执行引用错误扫描、参考比对、修复
    ↓
按选项执行隐藏状态恢复
    ↓
按选项执行统一字体
```

自动检查由精简任务收尾后启动，后续各模块有自己的任务入口、Excel 生命周期、状态和结果。

证据：

- [core/workflow.py:578 — 精简任务收尾后启动 clean_chain](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L578)
- [core/clean_chain.py:210 — BL 扫描与清除](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/clean_chain.py#L210)
- [core/clean_chain.py:298 — run 按步骤运行整个文件夹](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/clean_chain.py#L298)

合并后的结构则是：

```text
文件 A：精简 → BL → 引用 → 隐藏 → 字体 → 提交
文件 B：精简 → BL → 引用 → 隐藏 → 字体 → 提交
……
```

这不仅改变开存次数，也改变：

- 文件级完成状态的含义；
- 各步骤进度如何汇总；
- 失败后的续跑位置；
- 原独立工具页面能否继续显示同一批结果；
- 精简选中的文件范围与自动检查文件范围如何对应。

目前精简入口支持 `options.files` 过滤，而自动链只接收文件夹和选项，再由各模块列文件；合并前应明确是否让两者统一使用同一份清单，不能在重构中悄悄改变范围。

证据：[core/workflow.py:3804 — 精简文件范围过滤](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L3804)。

### 4.2 单文件打开与保存次数

下表在相应步骤启用、文件可处理的条件下计数：

| 步骤 | 当前文件 COM 打开 | 参考文件 COM 打开 | 当前处理结果保存 | 补充 |
| --- | ---: | ---: | ---: | --- |
| 精简 Sheet | 0～1 | 0 | 0～1 | 预检确认无需删表、改名、排序时，不打开 Excel |
| BL 超链接 | 1～2 | 0 | 0～1 | 只读扫描必需；有命中才可写清除 |
| 引用错误 | 1～2 | 0～1 | 0～1 | 当前先扫描；有错误才读参考；有可修项才写回 |
| 隐藏状态 | 1～2 | 1～2 | 0～1 | 推断出改名时，参考会带新投影再读一次 |
| 统一字体 | 1 | 0 | 0～1 | 还需检查图形，无目标单元格不代表可跳过 COM 打开 |
| 全部修改分支命中的合计 | **8** | **2～3** | **5** | 不含 Python 解析、临时转换、失败重试 |

对应证据：

- [core/workflow.py:3384 — 精简预检与按需打开](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L3384)
- [core/workflow.py:3450 — 精简保存与历史记录](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L3450)
- [core/bl_links.py:103 — scan_file](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/bl_links.py#L103)
- [core/bl_links.py:167 — clear_file](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/bl_links.py#L167)
- [core/link_repair.py:262 — repair 按 hits / ready 决定是否读参考和写回](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/link_repair.py#L262)
- [core/hidden_repair.py:236 — plan_file 与参考重读](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/hidden_repair.py#L236)
- [core/hidden_repair.py:258 — apply_file](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/hidden_repair.py#L258)
- [core/font_fix.py:122 — fix_workbook](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/font_fix.py#L122)

补充说明：

1. 隐藏步骤默认不勾选；默认启用的是 BL、引用和字体，因此不能用“全选上界”代表每份实际成本。
2. `.xlsx/.xlsm` 的精简预检和字体预检还分别有 Python 读取，不能算成两次 Excel 启动。
3. 隐藏恢复不仅恢复隐藏 Sheet，还包括隐藏列和超窄列；这些检查有内容签名、列定义和保护状态约束。
4. 二进制格式读取隐藏列定义时可能另存临时 `.xlsx`，这是临时转换，不是保存参考原件。
5. 各步骤复用自己批次内的 Excel 实例，但跨步骤未共享当前工作簿；优化对象是重复 Open/Save，而非假定每次 Open 都重启 Excel。

证据：

- [core/clean_chain.py:63 — 可选步骤及默认值](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/clean_chain.py#L63)
- [core/hidden_repair.py:26 — DEFAULT_OPTIONS](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/hidden_repair.py#L26)
- [core/hidden_repair.py:150 — 列定义与二进制临时转换](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/hidden_repair.py#L150)

### 4.3 合并时必须保留的正确性规则

#### A. 精简不是直接删表

当前删表前会：

1. 全量计算；
2. 找出依赖待删工作表的公式；
3. 捕获并保留计算结果；
4. 将需要保留的公式转成固定值；
5. 删除工作表；
6. 再计算，并拒绝引入新的引用错误。

合并后必须继续调用或保留这套逻辑，不能改成遍历 `Sheets.Delete()`。

证据：[core/formula_links.py:245 — delete_sheets_preserving_values](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/formula_links.py#L245)。

#### B. 未保存的删表历史也要供后续步骤使用

当前精简先保存，再将删表、改名、保留值和前后哈希写入历史；引用、隐藏、颜色对比会利用这些历史确认工作表身份和内容。

若改成最后才保存：

- 不能在文件提交前把计划冒充“已经完成的历史”；
- 也不能让后续步骤完全看不到本次删表、改名和转值；
- 需要向本文件后续操作传递本次暂存的映射及内容投影；
- 文件提交成功后，才登记最终历史和同步工作台记录。

证据：

- [core/sheet_history.py:38 — record 要求保存成功后登记](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/sheet_history.py#L38)
- [core/sheet_history.py:131 — Projection](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/sheet_history.py#L131)
- [core/sheet_history.py:162 — project 与历史阶段、哈希匹配](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/sheet_history.py#L162)
- [core/workflow.py:3297 — _sync_renamed_workbook_in_state](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L3297)

#### C. 字体预检不能继续读旧磁盘状态后按旧表名套用

目前 `yu_cells(path)` 从磁盘获得“表名 → 单元格地址”，之后重新打开同一个已保存文件修改。

合并后如果文件在内存中已删表或改名：

- 旧表名可能已经不存在；
- 被删除表的结果不应再应用；
- 新旧同名表的身份不能仅靠字符串猜测；
- 预检快照需要跟随本次实际改名映射，或者直接从当前工作簿获得等价数据。

不建议为了少一次 openpyxl 读取，就无条件改成逐格 COM 访问整个 UsedRange；这样可能减少打开次数却增加总耗时。

#### D. 内存公式与保存重开后的公式形态可能不同

代码已经处理删除工作表后内存中的 `#REF!A1` 与保存重开后 `#REF!` 的表示差异。

合并后不再经过中间保存与重开，必须验证引用扫描、模拟比较和最终验证仍使用一致的比较规则。

证据：[core/link_repair.py:96 — _formula_identity](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/link_repair.py#L96)。

## 5. “参考只读打开一次”的真实难度

### 5.1 同名限制已经在现有代码中显式规避

参考目录与处理目录通常保留相同文件名。当前颜色恢复先读参考、关闭参考，再打开当前文件；隐藏恢复也使用顺序打开。

因此以下伪流程并不可靠：

```text
同一个 Excel 实例
    打开 处理/a.xlsx
    再打开 参考/a.xlsx
    同时持有两个同名工作簿
```

现有代码明确说明第二次打开可能返回空对象。

证据：[core/mark_finder.py:631 — _read_old](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/mark_finder.py#L631)。

### 5.2 ReadOnly 只约束保存，不代表内存对象不变

引用修复的参考对比执行以下过程：

1. 只读打开参考；
2. 全量计算并捕获原公式、区域和计算值；
3. 在参考工作簿内存中模拟删除不保留的 Sheet；
4. 模拟改名并再次计算；
5. 比较模拟后的损坏公式与当前文件；
6. 不保存地关闭参考。

证据：[core/link_repair.py:112 — preview_file](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/link_repair.py#L112)。

**因此，不能直接把做完引用模拟后的参考工作簿交给隐藏恢复或颜色恢复。** 那时它已经不再表示原始参考内容。

### 5.3 对“先读参考快照，再打开当前文件”的进一步限定

前一轮简短答复提出的“先读参考快照”是可选技术方向，不是已经证明能覆盖①全部行为的完整方案。

- 对⑤颜色恢复，现有 `_read_old` 已经是按目标单元格读取参考快照的模式，较容易复用。
- 对①引用修复，当前实现依赖 Excel 的真实删表、改名与重新计算；普通的静态值快照不能直接取代它。
- 要一次读齐所有后续需要的数据，可能需要读取更多公式、值、列状态和样式，增加内存和 COM 通信成本。
- 隐藏恢复当前还会因推断改名而重算参考内容投影，不能只缓存一个不带阶段信息的最终签名。

更稳妥的目标分级：

| 目标 | 建议 |
| --- | --- |
| 当前文件集中到一次可写打开、至多一次提交 | 作为主要优化目标 |
| 参考的同一类数据只读取一次并在单文件内复用 | 作为明确可验收的次级目标 |
| ①所有参考操作严格只有一次 COM Open | 暂不作为首轮硬约束，先证明模拟与快照方案正确 |
| 将参考和当前放在不同 Excel 实例同时处理 | 仅作为备选，需要额外评估进程、计算环境与回收成本 |
| 复制并改名参考以规避同名限制 | 不能直接采用，需要验证外部链接、路径和公式身份不受影响 |

此外，`Application.CalculateFull()` 会作用于该 Excel 实例内的工作簿；若将原本分开打开的文件同时留在同一实例中，计算行为也需要实机验证。

## 6. ⑤上传前预处理

### 6.1 当前调用链

```text
prepare
    清标记
        打开当前文件
        删除工具条件格式标记、恢复工具标签色
        有修改时保存
        关闭
    对照参考恢复颜色
        解析当前和参考，生成计划
        有可恢复项时：读取参考实际颜色 → 关闭参考
        打开当前文件或工作副本 → 校验 → 写颜色 → 有修改时保存 → 关闭
    视图标准化
        再打开当前文件
        分页预览、A1、100%
        保存
        关闭
```

证据：

- [core/upload_prep.py:75 — prepare](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/upload_prep.py#L75)
- [core/batch.py:3725 — clear_marks_in_file](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/batch.py#L3725)
- [core/mark_finder.py:690 — _apply_colors](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/mark_finder.py#L690)
- [core/upload_prep.py:56 — _normalize](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/upload_prep.py#L56)

### 6.2 分格式、分命中情况的计数

| 场景 | 当前处理侧 COM 打开 | 参考 COM 打开 | Python 解析 | 备注 |
| --- | ---: | ---: | --- | --- |
| 无参考文件 | 2 | 0 | 无颜色对比解析 | 清标记与视图各一次 |
| 双方为 XML 格式，无可恢复颜色 | 2 | 0 | 当前、参考各一次 openpyxl | 不执行颜色写回 |
| 双方为 XML 格式，有可恢复颜色 | 3 | 1 | 当前、参考各一次 openpyxl | 两侧合计 4 次 COM Open |
| 进入二进制转换分支，无可恢复颜色 | 3 | 1 | 两个转换副本再由 openpyxl 解析 | 另有 2 次临时 SaveAs |
| 进入二进制转换分支，有可恢复颜色 | 4 | 2 | 两个转换副本再由 openpyxl 解析 | 两侧合计 6 次，另有 2 次临时 SaveAs |

二进制转换分支的条件不是“只有当前文件是二进制”：当两边未同时属于 `READABLE` 时，当前实现会依次转换两边。

对于 `.xls`，颜色写回另有“临时副本保存后替换原件”的保护，所以表中的当前处理侧次数包含该副本，并不表示每次都打开原路径。

证据：

- [core/color_compare.py:203 — plan_file 格式分支](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/color_compare.py#L203)
- [core/mark_finder.py:661 — apply_file 与 .xls 副本保护](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/mark_finder.py#L661)

普通成功路径中，清标记和颜色恢复都是有变化才保存，视图标准化则会调用保存；因此不能把每个 Open 都算成一次 Save。

### 6.3 清标记与颜色对比存在数据先后关系

“清标记 → 恢复颜色”这个顺序不能随意交换。

- 条件格式标记清除与直接底色/字色恢复是不同操作。
- 工具添加的 Sheet 标签色会被清除。
- 参考文件可能本来就有相同颜色的标签，后续对比需要按参考恢复。
- 若为了提前关掉参考，先生成颜色计划，再清标记，旧计划可能漏掉清标记刚产生的标签色差异。

如果保留 openpyxl 作为快速读取器，应将它视为**初始数据快照**：在内存中体现清标记带来的变化后再比较，或为已清除标签单独重算计划；不能继续读取未保存的磁盘文件并假设它代表当前工作簿。

颜色比较仍须保留原有约束：工作表唯一配对、内容/公式一致、合并区域一致、空格子所在行一致、主题色/图案色处理，以及不能安全恢复时的说明。

证据：[core/color_compare.py:142 — 标签与逐格颜色计划](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/color_compare.py#L142)。

### 6.4 现有“预处理失败”不保证本地文件未变

一个具体例子：

1. 清标记成功并保存；
2. 恢复颜色成功并保存；
3. 视图标准化抛出异常；
4. 上传层将该文件标记失败、不上传；
5. 但本地文件已经留下前两步修改。

另一个更细的例子：`clear_marks_in_file` 可能删除部分规则并保存，随后 `_clear_marks` 才检查 `leftover_failed` 并抛错。

因此，将写回集中到最后确实可以改善失败一致性，但这是**需要实现的新保证**，不是目前已经具备的性质。

证据：[core/upload_prep.py:20 — 清标记后检查失败报告](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/upload_prep.py#L20)。

### 6.5 现有警告不是全部阻断上传

当前规则包括：

- 有标记删不掉：抛异常，阻断该文件上传；
- 颜色无法安全恢复：记录 `color_skipped`，不一定阻断上传；
- 参考缺失：跳过颜色对比，仍执行其余步骤；
- 个别工作表视图设置失败：收集并写日志，随后仍可保存并返回 `normalized=True`；
- 真正抛出的预处理异常：上传层对该文件重试或失败，不影响其他文件。

这说明“全部三步成功”当前是流程返回意义上的成功，**不等于所有单元格与每张表的设置都成功**。

是否将上述警告升级为阻断，属于产品行为决策，不应借性能重构默默改变。

证据：

- [core/excel_ops.py:946 — normalize_workbook_view 的按表错误处理](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/excel_ops.py#L946)
- [core/upload_prep.py:32 — 颜色成功、跳过结果汇总](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/upload_prep.py#L32)
- [core/distribute.py:846 — 预处理失败与重试](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/distribute.py#L846)

### 6.6 不要破坏上传模式和 PDF 缓存边界

目前：

- 只传 PDF 时，在临时副本上做预处理，处理目录原 Excel 不应被修改。
- 只传 PDF 且缓存命中时，可直接跳过预处理与导出，达到该文件 0 次 COM 打开。
- Excel 与 PDF 一起上传时，即使 PDF 缓存可用，Excel 预处理仍可能执行。
- 无可复用 PDF 时，预处理后还有独立的同名 PDF 导出调用；这次打开不在前面的“三步预处理”计数内。
- 自由起点到终点上传的预处理使用临时副本，且没有参考颜色恢复。

合并方案必须维持这些边界，不能为了统一生命周期反而让缓存命中的文件重新打开。

证据：

- [core/distribute.py:796 — _prepare_task_files](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/distribute.py#L796)
- [core/distribute.py:1634 — _prepare_free](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/distribute.py#L1634)

## 7. ②③④⑥其他流程的补充核查

### 7.1 ②调整页面：主路径已合并，存在额外打开分支

当前主路径：

```text
缓存可用 → 直接返回，不打开 Excel
缓存不可用 → 打开 → 宽一页 → 导出 PDF → 分页预览 → 有变化才保存 → 关闭
```

但导出后若检测到可修复的稀疏页，会调用 `_fix_sparse_pages` 再打开当前文件，调整后重新导出并验证，有变化再保存。

因此更准确的描述是：

- 缓存命中：0 次打开；
- 主路径：1 次打开、至多 1 次保存；
- 可修稀疏页命中：可再增加 1 次打开和至多 1 次保存；
- 文件被占用：只读导出，不把页面设置保存到原件。

代码还保留打印方向写回的额外打开分支，但当前 `print_layout.ENABLED=False`，不能把这个关闭中的功能算成常规开销。

证据：

- [core/workflow.py:3463 — _export_workbook_pdfs](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L3463)
- [core/workflow.py:3616 — _fix_sparse_pages](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L3616)
- [core/print_layout.py:24 — 当前方向比较关闭](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/print_layout.py#L24)

结论：暂不优先改是合理的，但原因应是“主路径已合并、剩余是条件分支”，不是“绝无可省”。

### 7.2 ③切字扫描：保留只读阶段是合理的

当前扫描主链向 `_refresh` 传递 `compact_print_tail=False`，随后使用只读打开原文件进行检测，不是无条件创建 Excel 工作簿副本，也不是扫描时仍自动缩末页。

PDF 可以复用，但命中检查发生在工作簿已打开之后，主要省导出，不省本次只读 Open。

扫描之后，以下情况都可能要求可写打开：

- 新增需要标记的格子；
- 删除旧标记或更新标记颜色；
- 核对上次未确认保存的标记；
- 恢复或更新工作表标签；
- 没有标色工作，但仍需补存分页预览。

默认低风险黄色被排除，但用户打开“黄色也列出来”后会改变这一条件；是否标色还取决于人工待处理状态和位置确认，不能仅按“有无高风险项”判断。

普通完整路径可概括为 **1 次只读 + 0/1 次可写**，但不能由此推导“大部分文件一定只开一次”，这需要数据分布支持。

更重要的是，当前检测阶段有超时限制，写回前会解除该限制；把它们合并为一个可写工作簿，会扩大看门狗终止 Excel 时的影响范围。

证据：

- [core/batch.py:1315 — _refresh 传只读参数](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/batch.py#L1315)
- [core/batch.py:4215 — 检测超时与后续写回边界](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/batch.py#L4215)
- [core/batch.py:3340 — _sync_highlights 的打开条件](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/batch.py#L3340)
- [core/batch.py:145 — _hidden_low_risk](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/batch.py#L145)
- [core/workflow.py:1815 — 扫描中复用 PDF](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L1815)

结论：首轮不合并扫描只读与可写阶段，理由充分。

### 7.3 ④工具箱统一修改与复检：逻辑必须分开，物理重开未必必须

当前大体是：

```text
打开并修改 → 保存关闭
    ↓
若修改的表已有调整页面 PDF，先重新导出这些表
    ↓
整份复检（按需只读去标记、复用或导出 PDF）
    ↓
按需再次可写同步标记
```

工具箱内部批量改单元格已经集中保存一次；没改动时不做后续复检。

但是：

- 有调整页面 PDF 的文件，复检前可能多一次 `_export_workbook_pdfs(force=True)`；
- 复检可能使用同一个已去标记的只读工作簿，不应把 `_open_without_marks` 和 `_refresh` 重复计为两次打开；
- 独立复检发现文件自上次检测后没有变化时，会直接沿用检测结果，并非每次都必然导出。

保留“修改后验证”完全正确，但“因为是两件事所以技术上不能共用工作簿”不成立；是否共用，需要另外设计已保存版本验证、临时去标记和恢复、PDF 缓存等边界。

证据：

- [core/cell_toolbox.py:167 — apply_tools](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/cell_toolbox.py#L167)
- [core/cell_toolbox.py:272 — _fix_file](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/cell_toolbox.py#L272)
- [core/cell_toolbox.py:322 — _reexport_adjusted](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/cell_toolbox.py#L322)
- [core/batch.py:1656 — _recheck_file](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/batch.py#L1656)

结论：不列首轮，但不要宣布已经没有优化空间。

### 7.4 ⑥独立批量同名 PDF：使用 Excel，当前入口没有跳过缓存命中的导出

独立同名 PDF 入口每份只读打开一次，内存中清标记、处理选项并导出，最终不保存 Excel。

它与“调整页面”的 PDF 功能不是同一个入口：

- 调整页面可以在打开前命中缓存并跳过；
- 独立同名 PDF 的 `_run` 当前调用 `pdf_cache.remember` 记录导出结果，但没有先调用缓存查询跳过这份文件；
- 因而再次执行独立同名 PDF 批处理，仍会重新打开和导出。

**已有一次打开不代表没有优化空间：在输入、选项和输出校验都一致时，未来可考虑做到零打开、零导出。** 这是待实现的缓存优化，不是当前已有能力。

证据：

- [core/pdf_bundle.py:170 — export_workbook](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/pdf_bundle.py#L170)
- [core/pdf_bundle.py:256 — _run 与 remember](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/pdf_bundle.py#L256)

### 7.5 ⑥按 JSON 复制：Excel COM 是零次，不是一次

核心复制流程是复制到临时文件、校验哈希、替换目标，不需要 Excel COM。

但涉及覆盖冲突时，可能通过 openpyxl 和 ZIP/XML 比较文字及打印区域，所以“只是复制、完全不读工作簿内容”也不够准确。

证据：

- [core/distribute.py:513 — copy_one](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/distribute.py#L513)
- [core/distribute.py:290 — 覆盖判断中的内容比较](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/distribute.py#L290)
- [core/content_check.py:81 — snapshot](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/content_check.py#L81)

结论：没有 COM 重开可合并，文件复制和内容比较成本应单独度量。

## 8. 组合截图、红框与宽一页是否应一起加入

### 8.1 “组合截图红框”不是生成一张新的截图

现有实现直接操作 Excel 图形对象：将相邻或重叠的图片、标注、线条、文字框以及已有组合进行 Group。

两种用途不同：

| 用途 | 行为 | 是否写入 Excel |
| --- | --- | --- |
| 导出时临时组合 | 导出前组合及调整 Placement，结束后尝试拆开并恢复 | 只读导出不保存；可写调用方需注意恢复异常 |
| 永久组合 | 保留组合，设置“随单元格移动、不改变大小” | 由调用方保存 |

永久组合会排除某些打印用大框，以便导出时可以单独隐藏这些框。

证据：

- [core/print_overlays.py:123 — find_clusters](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/print_overlays.py#L123)
- [core/print_overlays.py:265 — grouped_overlays](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/print_overlays.py#L265)
- [core/print_overlays.py:304 — group_permanently](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/print_overlays.py#L304)
- [core/image_restore.py:483 — group_workbook](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/image_restore.py#L483)

### 8.2 永久组合不等于修复已经变形的图片

如果图片已经被拉伸或红框已经偏移，只组合会把当前错误状态组合起来，不会自动还原原始几何关系。

因此维持“图片修复作为问题出现后的独立功能”是合理的；若日后加入自动永久组合，应明确它是预防性结构调整，不是图片修复的替代品。

### 8.3 不能未经验证就宣称幂等，也不能断言每次都会重复组合

已组合图形属于允许参与聚类的类型，但一个孤立组合不会因为重复执行就必然再套一层，因为集群至少需要两个顶层对象。

另一方面，组合后的外接矩形、附近其他图形、后续人工编辑，都可能让下一次出现新的组合关系；当前测试不能据此证明任意工作簿反复执行都完全不变。

应新增或执行真实样本验证：

- 连续执行两次，不产生非预期嵌套或扩大组合范围；
- 红框与截图相对位置、叠放顺序不变；
- 打印大框仍可单独隐藏；
- 名称重复、受保护图形、已有多层组合可正确处理；
- 后续统一字体仍能访问需要处理的文字。

现有字体图形遍历只显式展开当前组合的一层，永久组合可能影响后续遍历可见的层级，尤其需要重复执行测试。

证据：[core/font_fix.py:80 — _shape_items](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/font_fix.py#L80)。

### 8.4 宽一页应保留当前规则，而不是统一重置所有页高

现有 `one_page_wide` 在已经满足 `Zoom=False`、`FitToPagesWide=1` 时直接返回，不强制重置 `FitToPagesTall`。

如果集成时改成每次都设置页高自动，会改变用户已经设好的固定页高，超出了减少开存的范围。

证据：[core/page_fill.py:700 — one_page_wide](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/page_fill.py#L700)。

建议：

1. 第一轮保持业务输出不变，只重构工作簿生命周期。
2. 永久组合和提前设置宽一页分别作为后续变更验证。
3. 若最终都放进同一文件流水线，可按“先内容与字体，再图形组合，再页面设置”的方向设计，但以 PDF 与工作簿实测结果为准。
4. 不因永久组合失败的现有日志警告，默认升级成全文件失败；严格程度需单独确定。

## 9. 失败、重试与“不会半改半存”的保证边界

### 9.1 现有流程不是统一的全有或全无

精简链路中：

- 一步失败会记入结果，后面的步骤仍可继续；
- 某张表改名失败时，已经完成的删表仍可能保存；
- 图形字体无法修改等情况可能只记日志；
- 引用与隐藏规则无法安全确认时会跳过相应项。

上传预处理则是“抛出的异常阻断这份上传”，但此前步骤可能已经保存；部分颜色和视图问题只是提示。

证据：

- [core/clean_chain.py:319 — 单步失败后继续](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/clean_chain.py#L319)
- [core/workflow.py:3425 — 改名失败后仍继续精简](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L3425)
- [core/font_fix.py:108 — 图形字体按段错误处理](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/font_fix.py#L108)

### 9.2 建议将结果明确分层

以下是建议语义，不是已经实现的新字段：

| 结果类别 | 例子 | 建议处理 |
| --- | --- | --- |
| 无需修改 | 没有 BL、字体已符合要求 | 正常继续 |
| 未启用 | 没勾隐藏恢复 | 正常跳过 |
| 无参考且规则允许跳过 | 上传颜色恢复无参考 | 记录原因，按现行规则继续 |
| 不能安全自动修复 | 内容不一致、合并区域不同 | 保持原值，记录人工核对项 |
| 明确执行失败 | COM 异常、写入回读不一致 | 中止本文件，释放资源，再决定重试/跳过 |
| 提交失败 | Save 或替换失败 | 不宣布成功，记录提交阶段和恢复信息 |
| 提交后记录失败 | 文件已保存，数据库更新失败 | 不声称已回滚；续跑前核对磁盘版本 |

“任何一项出错就停”应主要针对明确执行失败，不应把所有 `skip/warn` 混成错误。

而且当前不是所有错误都会弹重试框：`file_problem.should_ask` 会区分文件/Excel 故障与程序给出的检查结论。

证据：[core/file_problem.py:40 — should_ask](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/file_problem.py#L40)。

### 9.3 两种集中提交方案

#### 方案 A：直接打开当前文件，最后一次 Save

```text
打开当前文件
    内存中依次修改
    任一步抛错 → Close(False)
    全部完成 → Save → Close(False)
```

优点：复制开销较少，适合先验证生命周期集中管理。

局限：

- 必须确保下游函数都不偷偷 Save、SaveAs 或 Close；
- 被内部捕获的错误仍可能留下部分内存修改；
- Save 自身失败不等于已自动恢复旧文件；
- `.xls` 可写打开取得写锁时就可能改变容器字节；
- 不能由此承诺任何故障下原件字节都不变。

旧版 `.xls` 写锁副作用在现有代码中已有专门处理，不能视为纯理论问题。

证据：[core/hidden_repair.py:277 — .xls 写锁后的版本基准](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/hidden_repair.py#L277)。

#### 方案 B：工作副本处理，成功后校验并替换

```text
记录当前文件和参考版本
    ↓
创建工作副本，确认复制一致
    ↓
只在工作副本中处理与保存
    ↓
关闭工作簿，校验结果与原件版本
    ↓
在文件系统支持的条件下替换当前文件
    ↓
持久化历史、结果和相关缓存状态
```

优点：能将大部分处理、保存故障与原件隔离，和现有 `.xls` 颜色恢复的保护方向一致。

代价与约束：

- 增加复制、校验、磁盘空间成本，必须计入性能测量；
- 尽量保留格式、扩展名和业务文件身份，不能无条件转成 `.xlsx`；
- 临时路径可能影响相对外部链接与缓存身份，需要验证并显式区分业务路径和工作路径；
- 替换前必须释放 Excel 文件句柄并再次校验原件未被外部修改；
- 不能假设任意网络共享、同步盘或跨文件系统替换都有相同的原子性和持久性保证；
- 文件替换与数据库更新仍是两个提交动作，需要可恢复的中间状态。

**如果目标是“业务步骤异常和工作副本保存失败时原件保持不变”，更适合方案 B；如果只是先减少中间 Save，则方案 A 的承诺必须收窄。**

### 9.4 成功记录必须在提交成功后生效

需要避免两种错误：

1. 先把各步骤记成成功，最后 Save 失败，页面却显示已完成。
2. 文件已提交，但记录失败，重试时当成“未动过”从头执行非幂等步骤。

建议利用现有任务和文件报告机制区分：处理中、待提交、提交完成、记录待恢复；不必为此引入通用事务框架。

停止或重试前应关掉本任务工作簿和 Excel 实例，使用户可以打开文件处理；重试时基于新的磁盘状态重建快照，而非复用上次失败的 COM 对象或旧哈希。

## 10. 性能收益：能确认什么，不能确认什么

### 10.1 原估算的算术修正

沿用原假设：

```text
文件数 N = 572
每份减少 K 遍开存
每遍耗时 t = 2～4 秒
总节省分钟 = N × K × t / 60
```

| 减少遍数 | 每遍 2 秒 | 每遍 4 秒 |
| ---: | ---: | ---: |
| 2 | 38.13 分钟 | 76.27 分钟 |
| 4 | 76.27 分钟 | 152.53 分钟 |
| 5 | 95.33 分钟 | 190.67 分钟 |

因此：

- ①假定减少 4～5 遍，对应约 **76～191 分钟**，低端并不到 90 分钟。
- ⑤假定减少 2 遍，对应约 **38～76 分钟**，“半小时左右”没有覆盖该假设范围。
- 以上只是检验原表达的算术，不代表认可“每份都会少这些遍”或“每遍确实 2～4 秒”。

### 10.2 不应把全分支上界乘以全部 572 份

需要知道：

- 实际启用了哪些选项；
- 有多少文件确实删表、改名或排序；
- BL、引用、隐藏、字体分别有多少命中；
- XML 与二进制格式的比例；
- 文件大小、图形数量、公式复杂度；
- 参考是否存在、配对是否成功；
- PDF 缓存命中率；
- 重试、超时、用户等待占多少时间。

例如：无 BL 的文件只做一次扫描，不会做清除开存；没有可恢复颜色的 XML 文件没有颜色写回那次 Open；隐藏未勾选时整个步骤没有成本。

### 10.3 更合理的估算模型

```text
总耗时 =
    Excel 启动与回收
  + 当前工作簿 Open / Close
  + 参考工作簿 Open / Close
  + Save / 临时 SaveAs
  + 计算、扫描、比对、修改
  + openpyxl / XML 解析
  + 复制与哈希
  + PDF 导出与检测
  + 状态持久化
  + 失败重试与用户等待
```

合并的净收益应估为：

```text
消除的 Open / Close / Save 成本
+ 确实消除的重复读取和计算成本
- 新增快照、复制、校验、状态协调成本
```

尤其不要把“改成同一个工作簿”理解为可以删除全部 `CalculateFull()`：部分计算是依赖变化后的正确性要求。

### 10.4 当前仓库日志不足以验证原时间假设

本次读取的 [log.txt](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/log.txt) 快照为：

- 大小：52,648 字节；
- 行数：359 行；
- 时间范围：2026-11-11 22:09:52 至 22:32:10；
- 可见内容包含浏览器测试的 `mock-shared` 场景；
- 检索未发现本次所需的 `clean_chain`、`任务 clean_sheets`、`统一字体`、`上传前处理完成`、`保存了 Excel`、`打开工作簿` 或 `572` 相关记录。

这只说明**当前仓库这份日志不构成本次大批量性能证据**，不代表其他数据目录、历史归档里没有日志。

另外，`saved_excel` 主要记录保存来源、大小和修改时间，并没有测量 Save 调用本身的持续时间；即使找到保存记录，也不能直接把相邻两行的时间差当成 Save 耗时。

证据：

- [core/debug_log.py:168 — saved_excel](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/debug_log.py#L168)
- [core/debug_log.py:302 — timed](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/debug_log.py#L302)
- [core/workflow.py:1704 — 扫描已有 Open 计时](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/core/workflow.py#L1704)

### 10.5 建议的测量方案

先选一批有代表性的副本，不直接对用户正在编辑的原件跑基准。

至少覆盖：

- 小、中、大文件；
- `.xlsx/.xlsm` 与 `.xls/.xlsb`；
- 无修改、少量修改、多步骤均有修改；
- 大量图形与大量公式；
- 同名参考、无参考、无法唯一配对；
- 首次处理与已处理后的再次执行。

每份记录：

| 分类 | 指标 |
| --- | --- |
| 基本属性 | 格式、大小、工作表数、启用步骤 |
| 资源使用 | 当前/参考 Open 次数、Save 次数、临时 SaveAs 次数 |
| 分段耗时 | Open、Close、Save、计算、扫描、比对、修改、复制、哈希、记录 |
| 业务结果 | 每步命中数、修改数、跳过数、异常数 |
| 总体结果 | 总耗时、提交状态、重试数、缓存命中 |

旧流程与新流程必须从同一原始输入各复制一份开始，不能让新流程直接处理旧流程已经修改过的结果，再宣称加速。

按格式和命中情况分组报告样本数、总耗时、均值、P50、P95；面向 572 份的推算应按实际各组数量加权，而不是把一个中位数或最慢样本乘以全部文件。

## 11. 推荐实施顺序与边界

### 阶段 0：明确基线与失败语义

先确定：

1. 需要保留哪些现有 `skip/warn` 行为；
2. “整份失败”是否要求原件字节不变；
3. 第一轮是否只做生命周期优化，不加入永久组合和宽一页；
4. 用哪些样本做旧新结果对照；
5. 计时能否覆盖 Open、Save、参考读取和复制等主要部分。

本报告给出的默认建议是：**保留业务判定规则，明确执行异常的文件级隔离，新增图形/页面行为分开验收。**

### 阶段 1：先集中⑤上传预处理

重点不是创建通用流水线框架，而是对现有模块做最小拆分：

- 文件级入口保留独立工具的使用方式；
- 提取对已打开工作簿执行的清标记、颜色写入和视图标准化操作；
- 操作函数不负责 Open、Save、Close；
- `prepare` 负责生命周期、阶段信息和最终提交；
- 颜色计划明确处理清标记后的标签变化；
- 保留格式分支和 `.xls` 工作副本保护；
- 保留只传 PDF、无参考和缓存命中的现有边界。

目标可定义为：

> 实际需要预处理的文件，当前工作簿或其工作副本集中到一次可写打开；所有修改完成后至多一次保存/提交；参考、解析和转换成本单独计量。

这里的“至多一次保存”允许无需修改时不保存，但如果为此调整当前视图标准化总会保存的行为，也必须验证版本、缓存和结果显示不受影响。

### 阶段 2：再集中①精简自动检查

需要额外完成：

- 文件范围统一；
- 文件内暂存的删表、改名和保留值投影；
- 当前内存状态上的扫描与修复；
- 引用参考模拟的隔离；
- 字体预检数据随表身份变化；
- 统一的阶段结果与文件级提交状态；
- 与独立工具结果页面的衔接；
- 重试和续跑不重复登记已完成历史；
- 无参考、可选项关闭以及不能安全修复的路径。

不将“参考严格只开一次”作为阻碍正确性的硬性指标；先消除当前文件的重复写入生命周期，再按测量结果优化参考。

### 阶段 3：单独加入永久组合和页面预设

如果产品上确定要加入：

- 分别验收图形结构与打印结果；
- 验证重复执行；
- 保留已有固定页高；
- 不把尚未修复的变形图片当成已修复；
- 明确组合失败、视图失败等是警告还是整份阻断；
- 内容、字体、图形或打印设置变更后，不应仅靠更新文件版本来宣称旧 PDF 仍有效。

### 阶段 4：按真实收益选择后续优化

候选包括：

- 独立同名 PDF 入口加入严格缓存查询；
- 调整页面的稀疏页修复与主打开合并；
- 工具箱修改后的调整页面 PDF 重导与复检资源共享；
- 参考数据按单文件范围复用；
- 避免无变化的保存。

这些是候选，不代表本次已实施，也不代表优先级高于①⑤。

## 12. 验收清单

### 12.1 功能等价

| 场景 | 必须验证 |
| --- | --- |
| 全部步骤无命中 | 不产生额外内容变更，不增加不必要保存 |
| 删除备份表 | 依赖公式转值正确，不新增引用错误 |
| 删表加改名、同名复用 | 后续参考配对正确，不误认另一张旧表 |
| BL 合并区域和超链接 | 与原规则相同，只删除目标链接及内容 |
| 引用错误 | 数组/溢出区域、邻近锚点、公式身份和修复后验证保持一致 |
| 隐藏 Sheet/列/超窄列 | 内容签名、保护、至少保留一张可见表等约束保持 |
| 字体 | 不改标准样式、字号、颜色；已删表不再访问，改名后的表正确处理 |
| 清标记 | 仅清工具标记，不删除正常条件格式 |
| 恢复颜色 | 内容、合并关系、空行、主题色、图案色、标签色正确 |
| 无参考与选项关闭 | 继续保留原有跳过与提示行为 |
| 视图 | 分页预览、A1、100% 与预期一致，按表失败可解释 |
| 只传 PDF | 原 Excel 不被预处理修改；命中缓存仍可零打开 |
| 二进制与宏工作簿 | 不丢宏、不意外转换格式、不因元数据变化误报外部修改 |

### 12.2 故障与恢复

至少覆盖：

- 每个业务步骤抛错；
- 清标记部分失败；
- 参考在读取后发生变化；
- 当前文件在提交前被外部修改；
- 文件被占用或只读；
- Save 失败、临时文件保存失败、替换失败；
- 文件提交后状态记录失败；
- 用户重试、跳过、停止任务；
- 任务中断后续跑；
- 程序错误不会被包装成“无需修改”；
- 尚未提交的结果不会显示成已完成。

若采用工作副本强保护方案，在提交前失败时检查原件哈希不变；若采用直接原件最后保存方案，不应给旧版 `.xls` 声称未经验证的字节级保证。

### 12.3 保存后结果验证

性能验收与内容验收应分开：

- 成功结果比较工作簿语义，不只比较文件二进制哈希；
- 对代表性样本在测试中重新打开，确认实际保存内容；
- 检查公式/值、Sheet 顺序与可见性、列宽、格式、超链接、宏与图形；
- 加入永久组合或页面变化时，对比导出 PDF 的页数、文字、红框位置和图片比例；
- 记录专门用于验收的重开，不将其混入生产流水线的优化计数。

### 12.4 性能验收

- 当前处理侧 Open、Save 的次数符合各分支设计；
- 参考 Open、Python 解析与临时 SaveAs 均单独记录；
- 端到端总耗时确实下降，而不只是 Open 次数下降；
- 大文件、高公式量、高图形量的 P95 不出现明显退化；
- 缓存命中场景不退化成重新打开；
- 无命中场景不因强制集中可写打开而增加大量成本；
- 失败重试时间与等待时间单独报告。

## 13. 现有测试资产与待补覆盖

以下测试文件已存在，可作为实施时的回归入口；本次仅阅读或确认存在，**没有执行测试，也不宣称当前测试已经验证新方案**。

| 测试文件 | 用途 |
| --- | --- |
| [tests/test_clean_chain.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_clean_chain.py) | 自动链顺序、可选项、失败继续、状态恢复；文件级全有或全无会改变部分预期 |
| [tests/test_upload_prep.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_upload_prep.py) | 清标记→恢复颜色→视图顺序、无参考、步骤错误、单份失败不影响其他上传 |
| [tests/test_bl_links.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_bl_links.py) | BL 识别和清除规则 |
| [tests/test_formula_links.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_formula_links.py) | 公式与删表保护 |
| [tests/test_hidden_repair.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_hidden_repair.py) | 隐藏状态比较与修复 |
| [tests/test_font_fix.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_font_fix.py) | 字体识别与修改 |
| [tests/test_color_compare.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_color_compare.py) | 颜色计划 |
| [tests/test_mark_finder.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_mark_finder.py) | 参考颜色读取与恢复 |
| [tests/test_clear_marks.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_clear_marks.py) | 标记清理 |
| [tests/test_sheet_history.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_sheet_history.py) | 历史投影、身份映射 |
| [tests/test_print_layout.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_print_layout.py) | 页面设置与版本跟随 |
| [tests/test_print_overlays.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_print_overlays.py) | 临时/永久组合及 Placement |
| [tests/test_image_restore.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_image_restore.py) | 图形恢复与组合相关回归 |
| [tests/test_file_problem.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_file_problem.py) | 重试、跳过与问题分类 |
| [tests/test_upload_pdf.py](file:///C:/Users/XuJunhao/Documents/GitHub/clear_handle_excel/tests/test_upload_pdf.py) | 上传 PDF 与预处理、缓存边界 |

还应补充针对新生命周期的断言：

1. 操作函数不会自行 Open、Save、Close。
2. 同一文件只由一个提交入口负责保存。
3. 中间步骤使用的是修改后的内存状态，而非旧磁盘状态。
4. 失败后没有错误的成功记录或重复历史。
5. 同名参考在真实 Excel 中可工作。
6. 引用模拟不污染其他步骤使用的原始参考数据。
7. 连续执行两遍永久组合的结果经过实机验证。
8. 新旧流程从相同原始副本出发，输出语义等价。

Mock 测试可以验证调用次数和生命周期，但不能独立证明 Excel 的同名限制、公式重算、图形分组和打印效果；这些需要真实 Excel 样本。

## 14. 最终建议

认可①⑤作为重点优化候选，但将目标改写为：

> 在保留现有业务判定与参考比对安全规则的前提下，将同一文件的修改集中到受控的工作簿生命周期，减少中间打开和保存；先证明结果、失败恢复和缓存一致性，再用实际日志确认净收益。

默认顺序建议为：**测量与语义确认 → ⑤集中预处理 → ①集中精简链路 → 独立验收图形组合/页面行为 → 根据数据继续优化其他分支。**

当前不能承诺省多少小时，也不能承诺“参考只开一次就天然等价”；可以确认的是，重复开存真实存在，而且通过边界明确的重构有实质性的优化机会。
