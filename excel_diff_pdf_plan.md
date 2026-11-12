计划新增：Excel 对比改成「一页看全部文件」+「左右 PDF 对照、红绿蒙版标出不同」

实施状态：**未实施，2026-11-12 中途停下**。只做了两处后端铺垫（见第 2 节），主体代码都还没写。以下记下需求、已经查清的情况、定下的方案和剩下的工作，接着做时不用重新调查。

1. **需求（用户 2026-11-12）**

   - 所有文件的对比放在一个页面里；每个文件有自己的上传按钮，可以单独上传。
   - 左右两边用 PDF 显示 Excel 最终印出来的样子，同时在 PDF 上显示差异：哪里改了、哪里不同，用红色 / 绿色蒙版盖在上面。也就是「PDF 画面对比」和「逐格差异判断」合成一个画面。
   - 说的是「上传 Excel」页当前任务的文件清单（`static/distribute.html` 的 `#distTaskFiles`）：现在每份要点「🔍 对比」各开一个标签页，一个任务有三百多份。

2. **已经改了的（保留在工作区，没提交）**

   - `core/excel_diff.py`：
     - `diff()` 多一个参数 `spots=False`。传 `True` 时，有不同的表多带一个 `spots`：不一样的格子（`cells`）、行高 / 列宽不同的行列和说明（`rows`、`cols`）、不一样的图片图形占的范围（`boxes`）、每个行列在两边各挑的一个「会印出来的格子」（`edges`）。
     - 新函数 `_edge_cells()`：给行高 / 列宽不同的行列挑探针格子，避开隐藏行列、打印区域外、跨行（跨列）的合并格。
     - 默认不传 `spots`，现有调用方（`web_clip_fixer.py` 的 `/api/excel_diff`、`core/upload_review.py`）拿到的结果不变。
   - `core/excel_watchdog.py`：`ExcelWatchdog.start()` 多一个参数 `listed=True`。传 `False` 时不进 `_ACTIVE`，进度框的「跳过这一份 / 结束整个任务」管不到它，给不占任务槽的旁路活用。默认行为不变。

   改完后跑了 `python -m unittest tests.test_excel_diff tests.test_excel_watchdog tests.test_content_check`：35 个测试全部通过。新加的 `spots` 和 `_edge_cells()` 本身还没有测试。

3. **查清的现状**

   - 现有对比：`core/excel_diff.py` 用 openpyxl 只读比两份 `.xlsx / .xlsm`，不开 Excel。页面是 `static/excel_diff.html` + `static/excel_diff.js`，路由 `/excel-diff`。效果图是前端按格子自己画的，图片图形画不出来。
   - 现有单份上传：对比页带 `task` + `rel` 进来时有「⬆ 没问题，上传这一份」，调 `/api/distribute/task_upload`，带 `confirmed: {文件: {mine, theirs}}`（两份文件的 SHA-256）。上传那一刻两边还是这两个版本才直接传（`core/distribute.py` 的 `_check_task_file`）。
   - 现有 PDF 导出：`core/pdf_bundle.py` 的 `_export_sheet()` 是「同名 PDF」的印法（强制宽一页、隐藏大外框、截图和红框临时组合）；`_drop_marks()` 在内存里去掉本工具的标记色；`_sheet_blank()`、`nothing_to_print()` 处理印不出东西的表。PDF 页面图用 PyMuPDF 渲染，`core/pdf_review.py` 的 `page_png()` 是现成例子。
   - 后台 Excel：`core/excel_ops.py` 的 `start_background_excel()` 用 `DispatchEx` 另开进程；`core/file_problem.py` 的 `Guard` 用全局状态，旁路活不能用，所以才给看门狗加了 `listed=False`。
   - 服务是 `ThreadingHTTPServer`，可以做长轮询。
   - 测试：`python -m unittest`（没装 pytest），Excel 一律用测试替身。**跑测试要先把环境变量 `CLEAR_HANDLE_EXCEL_DATA` 指到临时文件夹**，否则 `core/workflow.py` 导入时会读真正的数据文件夹。
   - **开发这台机器没装 Excel**：后台 Excel 导出 PDF 在这里跑不了，只能用测试替身、PyMuPDF 生成的 PDF、浏览器分别验证，真实导出要在装了 Excel 的机器上看。
   - PyMuPDF 实测：页面带旋转时，`get_drawings()` 给的是没旋转的坐标，要乘 `page.rotation_matrix`；`get_text("words")` 的坐标也是没旋转的；渲染出来的图是旋转后的。

4. **定下的方案**

   **PDF 对照**

   - 两份 Excel 各复制一份到程序数据文件夹（`data_dir.DATA_DIR / "excel_diff_pdf"`），原文件不碰；副本只读打开、不保存，用完删掉，只留 PDF 和结果。
   - 先照旧逐格比（`excel_diff.diff(..., spots=True)`）。两份完全一样就不导 PDF；只导有不同的表和只有一边有的表，精简注定会删的表不导。
   - 每张表用 `pdf_bundle._export_sheet()` 导出，和上传的同名 PDF 是同一种印法；再拼成 `old.pdf`、`new.pdf`，记下每张表占哪几页。
   - 页面图由后端按需渲染：`GET /api/excel_diff/page?key=…&side=old|new&p=…`，地址带内容哈希，可以缓存。

   **红绿蒙版怎么定位（探针 PDF）**

   - 同一张表再导一次：先在内存里把每处不一样的格子涂上一种别处没用的颜色（一处一种），导出后用 `page.get_drawings()` 找这些颜色的色块，色块的位置就是格子在 PDF 上的位置。涂色不改行高列宽和分页，所以两份 PDF 版面一样。
   - 颜色按 8 一格取（每个通道 32 档），避开干净那份 PDF 里已经有的颜色；读回来时按同样的格子归档，容得下导出时的舍入误差。
   - 连成片、差异种类相同的格子合成一个长方形只涂一次（全表改字体时不会涂几千次）；每张表最多 1500 处，文字、公式优先。
   - 涂之前：`sheet.Unprotect("")`（带空密码，不会弹框）、删条件格式、关掉「单色打印」。
   - 行高 / 列宽不同：涂 `edges` 里挑的那个格子，前端画成页面左边缘 / 上边缘的色条。图片图形不同：涂它占的范围，前端画虚线框。
   - 探针 PDF 页数和干净那份对不上就整张表不标，并写进调试日志。每张表记下「要标几处、找到几处、几处不在打印范围」，真机上出问题时能查。

   **退路**

   - 探针标不出来的表（有保护涂不了色、一处也没找到），以及没法逐格比的格式（`.xls / .xlsb`）：改成比两份 PDF 上的字（`get_text("words")` + `difflib.SequenceMatcher`），只标得出文字不同，页面上写明原因。
   - PDF 整个导不出来（没装 Excel、Excel 没反应）：页面显示原因，退回现在按格子画的效果图。这种结果不进缓存。

   **后台任务**

   - 新模块 `core/excel_diff_pdf.py`：一个工作线程带一个自己的后台 Excel，排队一份一份做；不占 `STATE["busy"]`，上传等任务在跑时也能对比。空闲 90 秒退出 Excel。
   - 看门狗：`ExcelWatchdog(...).start(listed=False)`，Excel 弹框或 5 分钟没反应就结束它，这一份报错，下一份重开。
   - `Workbooks.Open` 带 `ReadOnly=True, UpdateLinks=0, Password="", IgnoreReadOnlyRecommended=True`，有密码的文件直接报错而不是弹框。
   - 结果按两份文件的 SHA-256 缓存（`<key>/result.json`、`summary.json`、`old.pdf`、`new.pdf`），文件没变再打开不用重导；留 7 天、最多 800 份。
   - 接口：`POST /api/excel_diff` 改成「确保有任务 + 问进度」，参数 `summary`（只要结论）、`front`（插到队首）、`fresh`（重新对比）、`wait`（长轮询秒数）。用长轮询而不是定时器：标签页放到后台时浏览器会把定时器压到一分钟一次，接连的 fetch 不受影响。排队的任务 60 秒没人问就丢掉（页面关了）。

   **前端**

   - 把 `static/excel_diff.js` 里画结果的部分抽成共用的 `static/excel_diff_view.js`（`ExcelDiffView.render(container, data)`），样式抽到 `static/excel_diff.css`；筛选、搜索按容器绑定，不再用固定 id。
   - PDF 对照放在每张表的卡片里，替换原来的效果图：一个滚动区，每行是「旧第 N 页 | 新第 N 页」，不用同步滚动。默认只显示有不同的页，可以切到全部页；可以放大缩小。蒙版用百分比定位盖在页面图上，鼠标移上去显示是哪一格、哪几类不同、文字旧 → 新。点类型按钮不看某一类时，对应的蒙版一起隐藏。
   - 新页面 `static/excel_diff_all.html` + `static/excel_diff_all.js`，路由 `/excel-diff/all?task=…`：
     - 列出任务的全部文件（`/api/distribute/task_files`），每份一张卡片；卡片标题吸顶，上面有状态、对比结论、「🔍 对比」「⬆ 上传这一份」。
     - 「▶ 全部对比」从上到下一份一份比，可以暂停；有不同的自动展开，完全一样的收着。卡片展开才取完整结果、收起就清掉，三百多份不会把页面撑爆。
     - 单份上传沿用现在的做法（`task_upload` 带 `confirmed`），对比完才能点；多份排队一份一份传。
   - 上传页（`static/distribute.js`）：清单上方加「🔍 一页对比全部」（带 `run=1` 进去自动开始），每行的「🔍 对比」改成进这一页并定位到那一份。
   - `/excel-diff` 单份页保留（起点 → 终点上传、逐份确认用），换成同一套显示。

5. **剩下的工作**

   1. 给 `spots` 和 `_edge_cells()` 补测试。
   2. 写 `core/excel_diff_pdf.py`（任务队列、导出、探针、文字退路、缓存、页面图）。
   3. `web_clip_fixer.py`：`/api/excel_diff` 换成新接口，加 `/api/excel_diff/page` 和 `/excel-diff/all`；这个接口的轮询不写进「🖱️ 网页操作」日志。
   4. 前端：抽共用显示、PDF 对照、新页面、上传页入口。
   5. 测试：探针色块定位（PyMuPDF 画的 PDF）、格子合并成长方形、取色避让、文字退路、任务流程（Excel 测试替身）、缓存命中。
   6. README 和各页「📖 程序会怎么做」里的说明。
   7. **在装了 Excel 的机器上实测**：探针色块能不能从 Excel 导出的 PDF 里读出来、颜色有没有偏、有保护的表和 `.xls` 走没走退路。这一步在开发机上做不了。

6. **没定、要问用户的**

   - 远程没有这份文件（上传后是新建）时，要不要单独显示你这份的 PDF。现在的想法是只写一句「远程没有这份」。
   - 对比用的 PDF 要不要去页眉页脚、要不要跳过 `変更履歴`。现在的想法是照 Excel 原样印，不去不跳。
   - 新页面要不要一进来就自动全部对比。现在的想法是只有从「一页对比全部」按钮进来才自动开始，从某一行的「对比」进来只比那一份。
