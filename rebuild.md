# Upload Desk：从零重建与复刻说明

## 1. 文档用途

本文件用于把 Upload Desk 交给 AI，从空目录重建当前项目。**只需这一份 Markdown**：第 5 节包含项目源码、测试、依赖清单、锁文件，以及两份 Excel 的 Base64 编码。无需另附 Excel 或手动粘贴代码。

下面的步骤是根据当前项目整理的**最终状态重建顺序**。目标是从本文恢复当前项目，不需要 Git 历史或中间开发记录。

### 两种目标

| 目标 | 必需材料 | 能恢复什么 |
| --- | --- | --- |
| 复刻程序 | 本文 | 代码、测试和两份原 Excel；默认从空数据库开始 |
| 恢复当前工作状态 | 先按第 8 节把数据库备份加入本文，再只交付这份 Markdown | 在程序之外恢复备份时的文件夹记录、上传信息、计划及历史方案 |

Excel 是二进制文件，直接粘贴文字会丢失格式；本文将原文件字节转成 Base64 文本，并附有 SHA-256。第 4 节脚本会解码并校验，恢复为原来的 `.xlsx` 文件。

**隐私提醒：Base64 只是编码，不是加密。** 本文包含原始 Excel 的全部内容，源码也可能含业务示例数据。分享给 AI 或其他人之前，请确认全部内容允许共享。SQLite 现有数据默认没有放入本文。

### 当前快照信息

- 基于 Git 提交 `350dc81`，包含当时尚未提交的 README 文档修改。
- 快照日期：2026-11-04。
- 当前系统 Python：3.14.7；openpyxl：3.1.5。
- 当前 Node.js：24.20.0；npm：11.19.0。
- 上述是记录快照时的工具版本，不代表本项目只支持这些版本。

注意：当前 Python 依赖使用版本范围，没有 Python 锁文件。若要求一致的依赖环境，应记录原环境的实际版本，并在该范围内安装同一版本。操作系统、浏览器、字体、换行符和外部资源也可能影响表现，不能只凭功能测试断言字节级一致。

## 2. 给接手 AI 的执行指令

把完整的本文交给 AI，并附上下面的指令：

> 请在新的空目录中重建 Upload Desk，不要覆盖已有项目。
>
> 1. 本文是唯一交付文件，源码和两份 Excel 都在第 5 节，不要要求额外附件。检查内容是否完整；缺失或校验失败时明确报告，不要猜测或自行补写。
> 2. 将本文保存为 `REBUILD.md`，优先执行第 4 节的还原脚本。脚本会从代码块恢复文件、解码 Excel 并校验哈希。不要手动重写或省略代码。
> 3. 不要更换框架、重写算法、合并样式表、重命名文件、改变依赖版本或修改业务规则。遇到错误先报告原因；必要修复须说明与原快照的差异。
> 4. 保留 UTF-8 编码、中文及日文名称、HTML 元素 ID、脚本加载顺序和相对路径。不要把本机绝对路径或凭据硬编码进程序。
> 5. 使用锁文件安装 Node.js 依赖，不要手写第三方库源码。第三方库由包管理器安装，保留其许可证。
> 6. 两份 Excel 必须从本文还原到指定文件名，不能用空表替代。若本文包含可选数据库备份则一并还原，否则从空数据库开始；不要自动导入 Excel 覆盖恢复的数据。
> 7. 完成后执行第 6、7 节的安装、测试和验收步骤，并汇报结果。测试通过不等于界面完全相同，还需浏览器检查。
> 8. 校验每个文件的字节数和 SHA-256；任何文件不一致都停止。不要将外部字体或图标服务视为已打包的离线资源。

本文较长，优先交给能够读取本地文件并执行命令的 AI，让其运行还原脚本，不必在对话中逐段生成代码。若对话窗口有长度限制，不要只粘贴被截断的内容；还原脚本会检测缺失和校验失败。

## 3. 项目目标与技术约束

### 项目是什么

Upload Desk 是在本机运行的文件夹管理工作台，管理文件数量、处理进度、上传状态和上传分配方案。

主要功能：

- 父文件夹树、明细列表、统计、搜索、筛选、排序和增删改。
- 处理进度和上传进度分别管理，支持手动切换状态。
- Excel 导入预览、合并导入、完整 Excel 导出、客户模板导出。
- JSON 备份导入导出，以及旧浏览器数据迁移。
- 按人员和来源负责人生成上传方案，保存计划、复制方案 JSON。
- 查看历史方案快照并重新复制 JSON，兼容旧版历史 ZIP 下载。
- Windows 双击启动，检查依赖并打开浏览器。

### 技术栈

- 前端：原生 HTML、CSS、JavaScript，无前端构建步骤。
- 后端：Python 标准库 HTTP 服务，不使用 Flask 或 Django。
- 存储：SQLite，Excel 处理使用 openpyxl。
- Node.js 仅用于安装浏览器依赖和运行 JavaScript 测试。
- 浏览器依赖：JSZip、flatpickr；确切安装版本以锁文件为准。
- 页面还引用外部 Lucide 图标脚本和 Google Fonts。离线或资源服务变化时，图标、字体可能与原界面不同；源码原样复刻不应擅自替换资源。
- 支持 Python 3.9 及以上；Node.js/npm 的具体复现版本填写在第 1 节。

### 必须保持的业务规则

- 完成数量、已上传数量、开始日期、计划日期、上传日期是不同字段。
- 状态分为未完成、完成中、已完成/未上传、上传中、已上传；保留现有自动判断及手动覆盖规则。
- 完成数量留空与填写 0 不同；不能丢弃缺失值语义。
- Excel 合并按父文件夹和子文件夹匹配；已有上传、计划和手动信息按当前实现保留，冲突时整批拒绝。
- 方案候选为已完成且未上传的文件夹，已有计划不能重复分配；文件夹不可拆分，每人至少分到一个文件夹。
- **整个方案**的文件总数位于目标文件数上下 5 个以内，不是每个人各自达到目标。
- 优先接近目标，再考虑来源负责人数、父文件夹数、开始日期和人数均衡；最多展示 5 个不同方案，不保证全局最优。
- 确认方案保存计划日期和名义上传者，使文件夹进入上传中，不自动计为已上传。
- 历史方案保存确认时的快照，不能因后续编辑或删除而改变。
- 方案 JSON 保留结构：
  `{version, label, plannedDate, totalFiles, folderCount, people:[{name, totalFiles, folders:[{source, destination, files, owner, startDate}]}]}`。
- `source` 为源路径数组，`destination` 为 `[人名, ...源路径]`；不要改成无法区分路径层级的字符串。
- 当前确认方案流程复制 JSON，不生成新的 ZIP；保留旧 ZIP 的兼容代码和复制脚本。
- 原 `file://` 页面与 HTTP 服务页面的数据来源不同，浏览器数据不会自动迁移到 SQLite。

以上只是关键约束，完整行为以源码、测试和 README 为准。

## 4. 从空目录开始的重建顺序

### 步骤 1：从本文还原文件

在一个**没有 `uploadtask` 子目录**的空目录中，将本文件原样保存为 `REBUILD.md`。在该目录打开 Windows PowerShell，运行下面完整命令。需要先安装 Python 3.9 或更高版本；此步骤不需要联网或其他附件。

脚本先在内存中检查所有必需文件、字节数和 SHA-256，全部通过后才创建 `uploadtask`。它不会覆盖已有目标；如果因磁盘等原因写入中途失败，请检查残留目录，选择新的空目录重试。

<!-- REBUILD_EXTRACTOR -->
```powershell
@'
import base64
import hashlib
import json
import re
from pathlib import Path, PureWindowsPath

document_path = Path.cwd() / 'REBUILD.md'
target = Path.cwd() / 'uploadtask'
if not document_path.is_file():
    raise SystemExit('REBUILD.md was not found in the current directory.')
if target.exists():
    raise SystemExit('The uploadtask directory already exists; choose an empty parent.')

workbook = '\u8a2d\u8a08\u66f8\u306e\u5c65\u6b74\u30b7\u30fc\u30c8\u524a\u9664\u4e00\u89a7'
required = {
    '.gitignore', 'README.md', 'requirements.txt', 'package.json',
    'package-lock.json', 'server.py', 'start.py', 'index.html',
    r'static\css\style.css', r'static\css\game-style.css',
    r'static\css\planner.css', r'static\css\cartoon-style.css',
    r'static\js\app.js', r'static\js\planner.js',
    r'data\step1_copy_files.py', r'data\step2_copy_onedrive.py',
    'data\\' + workbook + '.xlsx', 'data\\' + workbook + '-for-custom.xlsx',
    r'tests\test_launcher.py', r'tests\test_server.py',
    r'tests\test_planner.js', r'tests\test_tree.js',
}
optional_database = r'data\upload_data.sqlite3'
pattern = re.compile(
    r'^<!-- REBUILD_FILE (\{[^\n]*\}) -->\n'
    r'````[a-zA-Z0-9_-]*\n'
    r'(.*?)\n````(?:\n|$)',
    re.MULTILINE | re.DOTALL,
)

document_bytes = document_path.read_bytes()
document = document_bytes.decode('utf-8-sig').replace('\r\n', '\n').replace('\r', '\n')
files = {}
for match in pattern.finditer(document):
    info = json.loads(match.group(1))
    name = info['path']
    if name not in required and name != optional_database:
        raise SystemExit('Unexpected file entry: ' + ascii(name))
    if name in files:
        raise SystemExit('Duplicate file entry: ' + ascii(name))
    content = match.group(2)
    if info['encoding'] == 'base64':
        data = base64.b64decode(''.join(content.split()), validate=True)
    elif info['encoding'] == 'utf-8':
        if info['newline'] == 'CRLF':
            content = content.replace('\n', '\r\n')
        elif info['newline'] != 'LF':
            raise SystemExit('Unsupported newline format: ' + ascii(name))
        data = content.encode('utf-8')
    else:
        raise SystemExit('Unsupported encoding: ' + ascii(name))
    if len(data) != info['size'] or hashlib.sha256(data).hexdigest() != info['sha256']:
        raise SystemExit('File length or SHA-256 does not match: ' + ascii(name))
    files[name] = data

missing = required.difference(files)
if missing:
    raise SystemExit('Missing file entries: ' + ascii(sorted(missing)))

target.mkdir()
for name, data in files.items():
    destination = target.joinpath(*PureWindowsPath(name).parts)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open('xb') as output:
        output.write(data)
with (target / 'REBUILD.md').open('xb') as output:
    output.write(document_bytes)
print('Restored and verified', len(files), 'files. Optional database:', optional_database in files)
'@ | python -
```

进入生成的目录，后续命令均在该目录中执行：

```powershell
Set-Location .\uploadtask
```

最终交付结构如下；`.venv`、`node_modules`、缓存由安装或运行生成，不需要粘贴它们的内容。

```text
uploadtask\
|-- .gitignore
|-- README.md
|-- REBUILD.md
|-- requirements.txt
|-- package.json
|-- package-lock.json
|-- server.py
|-- start.py
|-- index.html
|-- static\
|   |-- css\
|   |   |-- style.css
|   |   |-- game-style.css
|   |   |-- planner.css
|   |   `-- cartoon-style.css
|   `-- js\
|       |-- app.js
|       `-- planner.js
|-- data\
|   |-- step1_copy_files.py
|   |-- step2_copy_onedrive.py
|   |-- 設計書の履歴シート削除一覧.xlsx
|   |-- 設計書の履歴シート削除一覧-for-custom.xlsx
|   `-- upload_data.sqlite3  [运行生成，或从备份恢复]
`-- tests\
    |-- test_launcher.py
    |-- test_server.py
    |-- test_planner.js
    `-- test_tree.js
```

### 步骤 2：写入项目配置

第 4 节脚本已从第 5 节写入 `.gitignore`、`requirements.txt`、`package.json`、`package-lock.json` 和 `README.md`。

必须保存完整锁文件，仅有 package.json 中的版本范围不能锁定所有间接依赖。

### 步骤 3：写入后端

脚本已写入 `server.py`；应保持静态资源映射、数据库初始化及迁移、事务、版本冲突检测、输入校验、导入导出和错误响应。

数据库由启动时初始化，不需要人工创建表。当前包含 `items` 和 `allocation_history` 两张表及兼容旧数据库的字段迁移。

### 步骤 4：写入页面与基础交互

脚本已写入 `index.html`、四份 CSS 和 `static\js\app.js`。

四份 CSS 的覆盖顺序属于当前界面实现的一部分。保持 HTML 中实际的顺序、内联样式、外部资源和脚本的 `defer` 属性；`app.js` 必须先于 `planner.js` 加载。

### 步骤 5：写入分配功能

脚本已写入 `static\js\planner.js`。它依赖基础脚本中的共享状态和函数，不要当作独立入口，也不要把当前受限候选搜索改成另一种算法。

### 步骤 6：写入启动器和兼容脚本

脚本已写入 `start.py`、`data\step1_copy_files.py` 和 `data\step2_copy_onedrive.py`。

两份复制脚本属于旧版 ZIP 流程，虽然不再随新方案生成，但仍被兼容逻辑和测试引用，应保留。不要在复刻时执行复制操作或自行填写其本机路径配置。

### 步骤 7：检查 Excel、测试和数据库

脚本已从本文还原两份 Excel 和四份测试源码。若事先按第 8 节将数据库备份加入本文，脚本也会恢复数据库；否则保持数据库不存在，由程序首次启动时创建空库。

### 步骤 8：安装、运行、验收

执行第 6、7 节，不需要 `npm run build`，当前项目没有该构建脚本。

## 5. 文件全文与内嵌数据

每个 `REBUILD_FILE` 标记对应一个文件，标记内记录目标路径、编码、字节数和 SHA-256。文本文件使用四个反引号的代码块；为让脚本准确还原末尾换行，结束围栏前额外保留一个分隔换行。请不要手动修改标记或代码块。

两份 Excel 使用 Base64 嵌入，原始数据与格式都保留。使用第 4 节脚本还原，不要把 Base64 文本直接改名为 `.xlsx`。

### 5.1 `.gitignore`

<!-- REBUILD_FILE {"path":".gitignore","newline":"LF","encoding":"utf-8","size":65,"sha256":"254ce2f308e741e402f69ed769a37c5987d03c0413944063d1649902649d07b9"} -->
````gitignore
node_modules/
*.sqlite3
*.sqlite3-shm
*.sqlite3-wal
__pycache__/

````

### 5.2 `requirements.txt`

<!-- REBUILD_FILE {"path":"requirements.txt","newline":"LF","encoding":"utf-8","size":17,"sha256":"9bba605fd7b6ea23038de15e5a1ae6714b3d3fb14ed2d6e6bedbfb0665ca2122"} -->
````text
openpyxl>=3.1,<4

````

### 5.3 `package.json`

<!-- REBUILD_FILE {"path":"package.json","newline":"LF","encoding":"utf-8","size":79,"sha256":"c7b5c368e684a4238ad1cc8bf9b7cd186a722df9a603b9770f9666e611561fab"} -->
````json
{
  "dependencies": {
    "flatpickr": "^4.6.13",
    "jszip": "^3.10.2"
  }
}

````

### 5.4 `package-lock.json`

<!-- REBUILD_FILE {"path":"package-lock.json","newline":"LF","encoding":"utf-8","size":4847,"sha256":"883ba0e1f7e0433361138c5af4cc9ab4251ed660ad601b9331e6346579f79c30"} -->
````json
{
  "name": "uploadTask",
  "lockfileVersion": 3,
  "requires": true,
  "packages": {
    "": {
      "dependencies": {
        "flatpickr": "^4.6.13",
        "jszip": "^3.10.2"
      }
    },
    "node_modules/core-util-is": {
      "version": "1.0.3",
      "resolved": "https://registry.npmjs.org/core-util-is/-/core-util-is-1.0.3.tgz",
      "integrity": "sha512-ZQBvi1DcpJ4GDqanjucZ2Hj3wEO5pZDS89BWbkcrvdxksJorwUDDZamX9ldFkp9aw2lmBDLgkObEA4DWNJ9FYQ==",
      "license": "MIT"
    },
    "node_modules/flatpickr": {
      "version": "4.6.13",
      "resolved": "https://registry.npmjs.org/flatpickr/-/flatpickr-4.6.13.tgz",
      "integrity": "sha512-97PMG/aywoYpB4IvbvUJi0RQi8vearvU0oov1WW3k0WZPBMrTQVqekSX5CjSG/M4Q3i6A/0FKXC7RyAoAUUSPw==",
      "license": "MIT"
    },
    "node_modules/immediate": {
      "version": "3.0.6",
      "resolved": "https://registry.npmjs.org/immediate/-/immediate-3.0.6.tgz",
      "integrity": "sha512-XXOFtyqDjNDAQxVfYxuF7g9Il/IbWmmlQg2MYKOH8ExIT1qg6xc4zyS3HaEEATgs1btfzxq15ciUiY7gjSXRGQ==",
      "license": "MIT"
    },
    "node_modules/inherits": {
      "version": "2.0.4",
      "resolved": "https://registry.npmjs.org/inherits/-/inherits-2.0.4.tgz",
      "integrity": "sha512-k/vGaX4/Yla3WzyMCvTQOXYeIHvqOKtnqBduzTHpzpQZzAskKMhZ2K+EnBiSM9zGSoIFeMpXKxa4dYeZIQqewQ==",
      "license": "ISC"
    },
    "node_modules/isarray": {
      "version": "1.0.0",
      "resolved": "https://registry.npmjs.org/isarray/-/isarray-1.0.0.tgz",
      "integrity": "sha512-VLghIWNM6ELQzo7zwmcg0NmTVyWKYjvIeM83yjp0wRDTmUnrM678fQbcKBo6n2CJEF0szoG//ytg+TKla89ALQ==",
      "license": "MIT"
    },
    "node_modules/jszip": {
      "version": "3.10.2",
      "resolved": "https://registry.npmjs.org/jszip/-/jszip-3.10.2.tgz",
      "integrity": "sha512-3l+rb15IOWtUhU0H5MFqES/T6Kh7abYwjosBey/vD6hDt8zoEffkSC5Ws5SGtgVw3gBx2NEbhTeSW1+kWkpyTQ==",
      "license": "(MIT OR GPL-3.0-or-later)",
      "dependencies": {
        "lie": "~3.3.0",
        "pako": "~1.0.2",
        "readable-stream": "~2.3.6",
        "setimmediate": "^1.0.5"
      }
    },
    "node_modules/lie": {
      "version": "3.3.0",
      "resolved": "https://registry.npmjs.org/lie/-/lie-3.3.0.tgz",
      "integrity": "sha512-UaiMJzeWRlEujzAuw5LokY1L5ecNQYZKfmyZ9L7wDHb/p5etKaxXhohBcrw0EYby+G/NA52vRSN4N39dxHAIwQ==",
      "license": "MIT",
      "dependencies": {
        "immediate": "~3.0.5"
      }
    },
    "node_modules/pako": {
      "version": "1.0.11",
      "resolved": "https://registry.npmjs.org/pako/-/pako-1.0.11.tgz",
      "integrity": "sha512-4hLB8Py4zZce5s4yd9XzopqwVv/yGNhV1Bl8NTmCq1763HeK2+EwVTv+leGeL13Dnh2wfbqowVPXCIO0z4taYw==",
      "license": "(MIT AND Zlib)"
    },
    "node_modules/process-nextick-args": {
      "version": "2.0.1",
      "resolved": "https://registry.npmjs.org/process-nextick-args/-/process-nextick-args-2.0.1.tgz",
      "integrity": "sha512-3ouUOpQhtgrbOa17J7+uxOTpITYWaGP7/AhoR3+A+/1e9skrzelGi/dXzEYyvbxubEF6Wn2ypscTKiKJFFn1ag==",
      "license": "MIT"
    },
    "node_modules/readable-stream": {
      "version": "2.3.8",
      "resolved": "https://registry.npmjs.org/readable-stream/-/readable-stream-2.3.8.tgz",
      "integrity": "sha512-8p0AUk4XODgIewSi0l8Epjs+EVnWiK7NoDIEGU0HhE7+ZyY8D1IMY7odu5lRrFXGg71L15KG8QrPmum45RTtdA==",
      "license": "MIT",
      "dependencies": {
        "core-util-is": "~1.0.0",
        "inherits": "~2.0.3",
        "isarray": "~1.0.0",
        "process-nextick-args": "~2.0.0",
        "safe-buffer": "~5.1.1",
        "string_decoder": "~1.1.1",
        "util-deprecate": "~1.0.1"
      }
    },
    "node_modules/safe-buffer": {
      "version": "5.1.2",
      "resolved": "https://registry.npmjs.org/safe-buffer/-/safe-buffer-5.1.2.tgz",
      "integrity": "sha512-Gd2UZBJDkXlY7GbJxfsE8/nvKkUEU1G38c1siN6QP6a9PT9MmHB8GnpscSmMJSoF8LOIrt8ud/wPtojys4G6+g==",
      "license": "MIT"
    },
    "node_modules/setimmediate": {
      "version": "1.0.5",
      "resolved": "https://registry.npmjs.org/setimmediate/-/setimmediate-1.0.5.tgz",
      "integrity": "sha512-MATJdZp8sLqDl/68LfQmbP8zKPLQNV6BIZoIgrscFDQ+RsvK/BxeDQOgyxKKoh0y/8h3BqVFnCqQ/gd+reiIXA==",
      "license": "MIT"
    },
    "node_modules/string_decoder": {
      "version": "1.1.1",
      "resolved": "https://registry.npmjs.org/string_decoder/-/string_decoder-1.1.1.tgz",
      "integrity": "sha512-n/ShnvDi6FHbbVfviro+WojiFzv+s8MPMHBczVePfUpDJLwoLT0ht1l4YwBCbi8pJAveEEdnkHyPyTP/mzRfwg==",
      "license": "MIT",
      "dependencies": {
        "safe-buffer": "~5.1.0"
      }
    },
    "node_modules/util-deprecate": {
      "version": "1.0.2",
      "resolved": "https://registry.npmjs.org/util-deprecate/-/util-deprecate-1.0.2.tgz",
      "integrity": "sha512-EPD5q1uXyFxJpCrLnCc1nHnq3gOa6DZBocAIiI2TaSCA7VCJ1UJDMagCzIkXNsUYfD1daK//LTEQ8xiIbrHtcw==",
      "license": "MIT"
    }
  }
}

````

### 5.5 `README.md`

<!-- REBUILD_FILE {"path":"README.md","newline":"LF","encoding":"utf-8","size":7085,"sha256":"a2e48b1eb163e60a403ea05db02055cb6177344cb3479fe1015cd16a2b26e735"} -->
````markdown
# Upload Desk

本机运行的文件夹管理工具。数据保存在 SQLite，支持 Excel 和 JSON 导入导出。

从空目录复刻项目，请参阅 [从零重建与复刻说明](./REBUILD.md)。该文档内含源码、Excel 的文本编码、还原步骤和验收方法；当前数据库默认不包含在内。文档包含原表数据，分享前请确认数据可以共享。

## 启动

Windows 下双击 `start.py` 即可启动并打开浏览器。需要提前安装 Python 3.9 或更高版本，并使 `.py` 文件关联 Python；缺少 ZIP 依赖时还需要 Node.js 和 npm。启动器会自动安装缺少的依赖，首次安装需要网络。保持启动窗口打开，关闭窗口会停止服务。如果浏览器没有自动打开，可手动访问窗口中显示的地址。

也可以在项目目录手动启动：

```bash
python -m pip install -r requirements.txt
npm install
python server.py
```

打开 http://127.0.0.1:8000。服务只监听本机；首次使用时数据库为空，在网页中选择“导入 Excel”，选择 `data/設計書の履歴シート削除一覧.xlsx` 并确认预览。

## 目录结构

- `index.html`：网页入口，保留原路径以便读取旧版浏览器数据。
- `static/css/`、`static/js/`：网页样式和脚本。
- `start.py`：双击启动入口，自动安装缺失依赖并打开浏览器。
- `server.py`：本地 Python 服务。
- `data/`：SQLite 数据库和原始 Excel 表格。
- `tests/`：自动化测试。

## 数据

- 数据库为 `data/upload_data.sqlite3`，请定期备份该文件；不要把它提交到版本控制。可通过 `UPLOAD_DESK_DB` 环境变量指定其他数据库路径。
- Excel 导入按父文件夹和子文件夹匹配：匹配项更新表格信息，但保留已有上传数量、计划日期、手动状态和名义上传者；新记录使用表格中的上传信息。若文件数量或完成状态与已有上传记录或计划冲突，整批导入会被拒绝。原表的“消化件数”和“完了”对应处理进度，不代表上传状态。
- 文件夹状态自动根据数据判断：无开始日期为“未完成”，有开始日期但完成数量未达到总数为“完成中”，完成数量等于总数为“已完成/未上传”；已有计划为“上传中”，全部上传后为“已上传”。历史部分上传记录也显示为“上传中”。明细中的状态按钮每点一次按“未完成 → 完成中 → 已完成/未上传 → 上传中 → 已上传”循环切换，不同状态颜色不同；未设置文件数量时跳过“已上传”。手动设置不修改开始日期和完成数量，重新编辑这些数据后恢复自动判断。切换到“上传中”可不填写计划日期；切换到“已上传”会将上传数量设为总数并记录当天日期。取消计划或上传时会提示确认，并清除相应的计划或上传信息。
- 明细中的“完成数量”对应原表“消化件数”，可在编辑窗口修改；留空与填写 0 不同。已上传数量与完成数量分别管理。
- 明细显示上传日期，并可按指定日期筛选；筛选仅影响当前列表，不影响完整 Excel 或客户表导出。按上传日期筛选时，列表上方会显示“复制当日 JSON”，直接用当前列表中的文件夹生成 JSON（结构同方案 JSON，按名义上传者分组，没有名义上传者的归入“未分配”），不需要该日期有已保存的方案；搜索、负责人和状态筛选也会影响导出范围。旧记录如未保存上传日期，显示为“—”且不会匹配具体日期；取消上传后日期会清除。
- 编辑时从已有父文件夹、所属子文件夹和负责人中选择；“仅父文件夹”表示原表中没有子文件夹的记录。同一父文件夹下已有的同名记录不可重复选择。新增时选择已有父文件夹和负责人，或选择新建后输入名称。
- 生成方案时可以选择计划上传日期，默认为本机当天；确认后仅保存计划日期和名义上传者，状态变为“上传中”，不计入已上传。候选为“已完成/未上传”的文件夹，已有计划的文件夹不会重复分配。文件总数必须在目标文件数上下 5 个以内，每人都有文件夹且文件夹不拆分。优先接近目标，再依次减少来源负责人数、减少涉及的父文件夹数、优先选择开始日期较早的文件夹，并兼顾人数均衡；无日期的最后处理。最多提供 5 个不同分配方案，候选数量有限，不能保证找到全局最优解；范围内没有方案时需调整目标或来源条件。悬停或聚焦方案可查看分配摘要，同一人在分配人员和来源负责人处使用相同编号和颜色；点击方案后选中并查看完整明细。
- 确认方案时会将当时的分配明细和方案 JSON 保存在数据库中，可在“历史分配方案”查看并重新复制 JSON；重新复制不会修改文件夹状态。历史明细不会随文件夹后续修改或删除而改变。启用该功能之前的计划没有完整方案快照，无法在历史中还原；旧版本保存的 ZIP 仍可在历史中下载。备份数据库时请一并保留历史记录。
- 完整 Excel 导出包含计划上传日期、文件夹状态、手动状态和上传记录，支持重新导入。客户表使用 `data/設計書の履歴シート削除一覧-for-custom.xlsx` 作为模板，保留其工作表与格式，列出所有当前文件夹，并在有实际上传数量的行填写上传日期、名义上传者和已上传数量；张晓静在客户表中显示为“静”。标记上传时记录本机日期，取消上传会清除日期。导入没有上传日期的旧数据时不会猜测日期；系统不记录具体文件名。
- 点击“导出 JSON 备份”可将当前 SQLite 中的文件夹明细导出为 JSON 文件；点击“导入 JSON 备份”可将 JSON 备份合并导入 SQLite（匹配的文件夹会被更新，新增记录会插入，其他未提及记录保留）。原 `file://` 页面中的浏览器数据不会自动转移到服务页面，可在原浏览器中打开 `index.html` 点击“导出 JSON 备份”，再在服务页面点击“导入 JSON 备份”。
- 点击“确定并复制方案 JSON”时，所选文件夹进入“上传中”，不会自动标记为已上传，同时把方案 JSON 复制到剪贴板（不再下载 ZIP）。JSON 结构为 `{version, label, plannedDate, totalFiles, folderCount, people:[{name, totalFiles, folders:[{source, destination, files, owner, startDate}]}]}`：`people` 按人名分开，`folders` 是该人负责的文件夹；`source` 是相对资源库根目录的路径数组，`destination` 是 `[人名, ...源路径]`，可据此识别并移动文件。`data/` 下的 `step1_copy_files.py` 和 `step2_copy_onedrive.py` 是旧版 ZIP 流程使用的复制脚本，现在不再随方案生成。

运行测试：`python -m unittest discover -s tests -v`、`node tests/test_planner.js` 和 `node tests/test_tree.js`。

````

### 5.6 `server.py`

<!-- REBUILD_FILE {"path":"server.py","newline":"LF","encoding":"utf-8","size":36834,"sha256":"288cea184fbf43d9bf5c2ea0185aa2e14a64c60917416ccdbdee9fe6cac23c3b"} -->
````python
import io
import json
import os
import socket
import sqlite3
import uuid
from contextlib import contextmanager
from copy import copy
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent
DATABASE = Path(os.environ.get('UPLOAD_DESK_DB', ROOT / 'data' / 'upload_data.sqlite3'))
CUSTOMER_TEMPLATE = ROOT / 'data' / '設計書の履歴シート削除一覧-for-custom.xlsx'
PORT = int(os.environ.get('PORT', '8000'))
MAX_BODY = 20 * 1024 * 1024
STATIC = {
    '/': ('index.html', 'text/html; charset=utf-8'),
    '/index.html': ('index.html', 'text/html; charset=utf-8'),
    '/static/js/app.js': ('static/js/app.js', 'text/javascript; charset=utf-8'),
    '/static/js/planner.js': ('static/js/planner.js', 'text/javascript; charset=utf-8'),
    '/static/css/style.css': ('static/css/style.css', 'text/css; charset=utf-8'),
    '/static/css/planner.css': ('static/css/planner.css', 'text/css; charset=utf-8'),
    '/static/css/game-style.css': ('static/css/game-style.css', 'text/css; charset=utf-8'),
    '/static/css/cartoon-style.css': ('static/css/cartoon-style.css', 'text/css; charset=utf-8'),
    '/data/step1_copy_files.py': ('data/step1_copy_files.py', 'text/plain; charset=utf-8'),
    '/data/step2_copy_onedrive.py': ('data/step2_copy_onedrive.py', 'text/plain; charset=utf-8'),
    '/node_modules/jszip/dist/jszip.min.js': ('node_modules/jszip/dist/jszip.min.js', 'text/javascript; charset=utf-8'),
    '/node_modules/flatpickr/dist/flatpickr.min.js': ('node_modules/flatpickr/dist/flatpickr.min.js', 'text/javascript; charset=utf-8'),
    '/node_modules/flatpickr/dist/l10n/zh.js': ('node_modules/flatpickr/dist/l10n/zh.js', 'text/javascript; charset=utf-8'),
    '/node_modules/flatpickr/dist/flatpickr.min.css': ('node_modules/flatpickr/dist/flatpickr.min.css', 'text/css; charset=utf-8'),
}


def list_source_files(parts):
    root = os.environ.get('SOURCE_ROOT', '').strip()
    if not root or not isinstance(parts, list) or not parts:
        return []
    if any(not isinstance(part, str) or part in ('', '.', '..') or any(char in '<>:"/\\|?*' for char in part) for part in parts):
        return []
    folder = Path(root).joinpath(*parts)
    if not folder.is_dir():
        return []
    return sorted(path.relative_to(folder).as_posix() for path in folder.rglob('*') if path.is_file())


class LocalHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = os.name != 'nt'

    def server_bind(self):
        if os.name == 'nt':
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


@contextmanager
def connect():
    DATABASE.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DATABASE, timeout=10)
    connection.row_factory = sqlite3.Row
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def init_db():
    with connect() as connection:
        connection.execute('''CREATE TABLE IF NOT EXISTS items (
            id TEXT PRIMARY KEY, parent TEXT NOT NULL, name TEXT NOT NULL,
            total INTEGER NOT NULL CHECK(total >= 0), total_missing INTEGER NOT NULL DEFAULT 0,
            start_date TEXT NOT NULL DEFAULT '', owner TEXT NOT NULL DEFAULT '',
            processed INTEGER, complete INTEGER NOT NULL DEFAULT 0,
            uploaded INTEGER NOT NULL DEFAULT 0, parent_only INTEGER NOT NULL DEFAULT 0,
            version INTEGER NOT NULL DEFAULT 1,
            UNIQUE(parent, name)
        )''')
        columns = {row['name'] for row in connection.execute('PRAGMA table_info(items)')}
        if 'uploader' not in columns:
            connection.execute("ALTER TABLE items ADD COLUMN uploader TEXT NOT NULL DEFAULT ''")
        if 'upload_date' not in columns:
            connection.execute("ALTER TABLE items ADD COLUMN upload_date TEXT NOT NULL DEFAULT ''")
        if 'planned_date' not in columns:
            connection.execute("ALTER TABLE items ADD COLUMN planned_date TEXT NOT NULL DEFAULT ''")
        if 'status_override' not in columns:
            connection.execute("ALTER TABLE items ADD COLUMN status_override TEXT NOT NULL DEFAULT ''")
        connection.execute('''CREATE TABLE IF NOT EXISTS allocation_history (
            id TEXT PRIMARY KEY, created_at TEXT NOT NULL, planned_date TEXT NOT NULL,
            label TEXT NOT NULL, assignments TEXT NOT NULL, archive BLOB NOT NULL,
            manifest TEXT NOT NULL DEFAULT ''
        )''')
        if 'manifest' not in {row['name'] for row in connection.execute('PRAGMA table_info(allocation_history)')}:
            connection.execute("ALTER TABLE allocation_history ADD COLUMN manifest TEXT NOT NULL DEFAULT ''")


def as_item(row):
    item = {'id': row['id'], 'group': row['parent'], 'name': row['name'],
            'total': row['total'], 'totalMissing': bool(row['total_missing']),
            'date': row['start_date'], 'owner': row['owner'], 'processed': row['processed'],
            'complete': bool(row['complete']), 'uploaded': row['uploaded'],
            'uploader': row['uploader'], 'uploadDate': row['upload_date'], 'plannedDate': row['planned_date'],
            'statusOverride': row['status_override'], 'parentOnly': bool(row['parent_only']), 'version': row['version']}
    item['status'] = folder_status(item)
    return item


def folder_status(item):
    if item.get('statusOverride'):
        return item['statusOverride']
    if item['uploaded'] and item['uploaded'] == item['total']:
        return 'uploaded'
    if item['uploaded'] or item.get('plannedDate'):
        return 'uploading'
    if not item.get('date'):
        return 'unstarted'
    if not item.get('totalMissing') and item.get('processed') is not None and item['processed'] == item['total']:
        return 'ready'
    return 'processing'


def all_items(connection):
    return [as_item(row) for row in connection.execute('SELECT * FROM items ORDER BY rowid')]


class AllocationConflict(ValueError):
    pass


def mark_allocations(connection, assignments):
    if not isinstance(assignments, list) or not assignments:
        raise ValueError('分配记录不能为空')
    seen = set()
    connection.execute('BEGIN IMMEDIATE')
    for assignment in assignments:
        if not isinstance(assignment, dict):
            raise ValueError('分配记录格式无效')
        item_id = assignment.get('id')
        uploader = assignment.get('uploader')
        if not isinstance(item_id, str) or not item_id or item_id in seen:
            raise ValueError('分配记录重复或标识无效')
        if not isinstance(uploader, str) or not uploader.strip() or len(uploader.strip()) > 100:
            raise ValueError('名义上传者无效')
        upload_date = assignment.get('uploadDate', date.today().isoformat())
        if not isinstance(upload_date, str) or not upload_date.strip():
            raise ValueError('上传日期不能为空')
        upload_date = parse_date(upload_date)
        version = integer(assignment.get('version'), '版本')
        seen.add(item_id)
        row = connection.execute('SELECT * FROM items WHERE id=?', (item_id,)).fetchone()
        if not row or folder_status(as_item(row)) != 'ready':
            raise AllocationConflict('分配中的文件夹已被修改或不再符合条件，请刷新并重新生成方案')
        result = connection.execute('''UPDATE items SET uploader=?,planned_date=?,complete=1,status_override='uploading',version=version+1
            WHERE id=? AND version=? AND uploaded=0 AND planned_date='' AND total>0''',
            (uploader.strip(), upload_date, item_id, version))
        if result.rowcount != 1:
            raise AllocationConflict('分配中的文件夹已被修改或不再符合条件，请刷新并重新生成方案')
    return [as_item(connection.execute('SELECT * FROM items WHERE id=?', (assignment['id'],)).fetchone()) for assignment in assignments]


def save_allocation_history(connection, payload):
    label = payload.get('label')
    if not isinstance(label, str) or not label.strip() or len(label) > 200:
        raise ValueError('方案名称无效')
    manifest = payload.get('manifest')
    if not isinstance(manifest, dict) or not isinstance(manifest.get('people'), list):
        raise ValueError('方案 JSON 格式无效')
    saved = mark_allocations(connection, payload.get('assignments'))
    dates = {item['plannedDate'] for item in saved}
    if len(dates) != 1:
        raise ValueError('同一方案的计划日期必须一致')
    history_id = str(uuid.uuid4())
    connection.execute('INSERT INTO allocation_history (id,created_at,planned_date,label,assignments,archive,manifest) VALUES (?,?,?,?,?,?,?)',
                       (history_id, datetime.now().isoformat(timespec='seconds'), dates.pop(), label.strip(),
                        json.dumps(saved, ensure_ascii=False), b'', json.dumps(manifest, ensure_ascii=False)))
    return saved, history_id


def integer(value, label, allow_empty=False):
    if value is None or value == '':
        if allow_empty:
            return None
        raise ValueError(f'{label}不能为空')
    if isinstance(value, bool) or not isinstance(value, (int, float)) or int(value) != value or not 0 <= value <= 1000000:
        raise ValueError(f'{label}必须是 0 到 1000000 的整数')
    return int(value)


def text(value):
    return str(value).strip() if value is not None else ''


def parse_date(value):
    if value is None or value == '':
        return ''
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    try:
        return date.fromisoformat(str(value).strip().replace('/', '-')).isoformat()
    except ValueError as error:
        raise ValueError(f'日期格式无效：{value}') from error


def validate_item(item):
    group, name = text(item.get('group')), text(item.get('name'))
    if not group or not name or len(group) > 500 or len(name) > 200:
        raise ValueError('父文件夹和子文件夹名称不能为空或过长')
    total = integer(item.get('total'), '文件数量')
    uploaded = integer(item.get('uploaded', 0), '已上传数量')
    processed = integer(item.get('processed'), '消化件数', allow_empty=True)
    complete = uploaded > 0 if item.get('complete') is None else item.get('complete')
    if not isinstance(complete, bool):
        raise ValueError('完成状态无效')
    if uploaded > total or (processed is not None and processed > total):
        raise ValueError('已上传数量或消化件数不能大于文件数量')
    if uploaded and not complete:
        raise ValueError('已有上传数量的文件夹必须为已完成')
    owner = text(item.get('owner'))
    if len(owner) > 100:
        raise ValueError('负责人名称过长')
    planned_date = parse_date(item.get('plannedDate'))
    status_override = item.get('statusOverride', '') or ''
    if status_override not in ('', 'unstarted', 'processing', 'ready', 'uploading', 'uploaded'):
        raise ValueError('文件夹状态无效')
    if status_override == 'uploaded' and (not total or uploaded != total):
        raise ValueError('已上传状态要求全部文件已上传')
    if status_override in ('unstarted', 'processing', 'ready') and (uploaded or planned_date):
        raise ValueError('请先取消上传计划或上传状态')
    if status_override == 'uploading' and total and uploaded == total:
        raise ValueError('上传中状态不能是全部文件已上传')
    if (planned_date or status_override in ('ready', 'uploading', 'uploaded')) and not complete:
        raise ValueError('已完成或上传中的文件夹必须为已完成')
    uploader = text(item.get('uploader')) if uploaded or planned_date or status_override == 'uploading' else ''
    if len(uploader) > 100:
        raise ValueError('名义上传者名称过长')
    return {'group': group, 'name': name, 'total': total, 'totalMissing': bool(item.get('totalMissing', False)),
            'date': parse_date(item.get('date')), 'owner': owner, 'processed': processed,
            'complete': complete, 'uploaded': uploaded, 'uploader': uploader,
            'uploadDate': parse_date(item.get('uploadDate')) if uploaded else '', 'plannedDate': planned_date,
            'statusOverride': status_override,
            'parentOnly': bool(item.get('parentOnly', False))}


def parse_workbook(data):
    try:
        workbook = load_workbook(io.BytesIO(data), data_only=False)
    except Exception as error:
        raise ValueError('无法读取 Excel 文件，请选择有效的 .xlsx 文件') from error
    required = {'親フォルダ', '子フォルダ', 'ファイル数', '開始日', '担当者', '消化件数', '完了'}
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 30)):
            headers = {text(cell.value): cell.column for cell in row if isinstance(cell.value, str)}
            if required.issubset(headers):
                break
        else:
            continue
        break
    else:
        raise ValueError('未找到包含親フォルダ、子フォルダ、ファイル数等列的工作表')
    parent_col = headers['親フォルダ']
    merged_parents = {}
    for merged in sheet.merged_cells.ranges:
        if merged.min_col == merged.max_col == parent_col:
            parent = sheet.cell(merged.min_row, parent_col).value
            for row_number in range(merged.min_row, merged.max_row + 1):
                merged_parents[row_number] = parent
    results = []
    paths = set()
    for row_number in range(row[0].row + 1, sheet.max_row + 1):
        def cell(column):
            return sheet.cell(row_number, headers[column]).value
        group = text(merged_parents.get(row_number, cell('親フォルダ')))
        child = text(cell('子フォルダ'))
        if not group:
            if child and not any(word in child for word in ('総件数', '合計')):
                raise ValueError(f'第 {row_number} 行缺少父文件夹')
            continue
        try:
            total_value = integer(cell('ファイル数'), '文件数量', allow_empty=True)
            total = total_value if total_value is not None else 0
            processed = integer(cell('消化件数'), '消化件数', allow_empty=True)
            marker = cell('完了')
            complete = (text(marker) in ('〇', '○', '已完成', '完成', '✓', '✔', 'TRUE', 'True') or marker is True)
            if marker is None or (isinstance(marker, str) and marker.startswith('=')):
                complete = total_value is not None and processed is not None and total == processed
            upload_column = next((name for name in ('已上传数量', 'アップロード数') if name in headers), None)
            uploaded = integer(cell(upload_column), '已上传数量', allow_empty=True) if upload_column else None
            parent_only = not child
            name = child or group.replace('¥', '\\').replace('/', '\\').split('\\')[-1]
            parsed = validate_item({'group': group, 'name': name, 'total': total, 'totalMissing': total_value is None,
                                    'date': parse_date(cell('開始日')), 'owner': text(cell('担当者')),
                                    'processed': processed, 'complete': complete, 'uploaded': uploaded or 0,
                                    'uploader': text(cell('名义上传者')) if '名义上传者' in headers else '',
                                    'uploadDate': parse_date(cell('上传日期')) if '上传日期' in headers else '',
                                    'plannedDate': parse_date(cell('计划上传日期')) if '计划上传日期' in headers else '',
                                    'statusOverride': text(cell('手动状态')) if '手动状态' in headers else '',
                                    'parentOnly': parent_only})
            parsed['uploadedProvided'] = upload_column is not None
            parsed['uploaderProvided'] = '名义上传者' in headers
            parsed['uploadDateProvided'] = '上传日期' in headers
            parsed['plannedDateProvided'] = '计划上传日期' in headers
            parsed['statusOverrideProvided'] = '手动状态' in headers
        except ValueError as error:
            raise ValueError(f'第 {row_number} 行：{error}') from error
        key = (parsed['group'], parsed['name'])
        if key in paths:
            raise ValueError(f'第 {row_number} 行的文件夹重复：{group} / {name}')
        paths.add(key)
        results.append(parsed)
    if not results:
        raise ValueError('工作表没有可导入的文件夹数据')
    return results, sheet.title


def merge_items(connection, records, preserve_upload=False):
    inserted = updated = 0
    for record in records:
        if not isinstance(record, dict):
            raise ValueError('备份数据格式不正确')
        existing = connection.execute('SELECT * FROM items WHERE parent=? AND name=?', (text(record.get('group')), text(record.get('name')))).fetchone()
        if existing and (preserve_upload or not record.get('uploadedProvided', True)):
            record = {**record, 'uploaded': existing['uploaded']}
        if existing and (preserve_upload or not record.get('uploaderProvided', 'uploader' in record)):
            record = {**record, 'uploader': existing['uploader']}
        if existing and (preserve_upload or not record.get('uploadDateProvided', 'uploadDate' in record)):
            record = {**record, 'uploadDate': existing['upload_date']}
        if existing and (preserve_upload or not record.get('plannedDateProvided', 'plannedDate' in record)):
            record = {**record, 'plannedDate': existing['planned_date']}
        if existing and (preserve_upload or not record.get('statusOverrideProvided', 'statusOverride' in record)):
            record = {**record, 'statusOverride': existing['status_override']}
        if preserve_upload and existing and existing['uploaded'] and (record['total'] < existing['uploaded'] or not record['complete']):
            raise ValueError(f'{record["group"]} / {record["name"]}：表格中的文件数量或完成状态与已有上传记录冲突，未导入任何记录')
        try:
            record = validate_item(record)
        except ValueError as error:
            raise ValueError(f'{record.get("group", "")} / {record.get("name", "")}：{error}') from error
        values = (record['group'], record['name'], record['total'], int(record['totalMissing']), record['date'], record['owner'],
                record['processed'], int(record['complete']), record['uploaded'], int(record['parentOnly']), record['uploader'], record['uploadDate'], record['plannedDate'], record['statusOverride'])
        if existing:
            connection.execute('UPDATE items SET parent=?,name=?,total=?,total_missing=?,start_date=?,owner=?,processed=?,complete=?,uploaded=?,parent_only=?,uploader=?,upload_date=?,planned_date=?,status_override=?,version=version+1 WHERE id=?', (*values, existing['id']))
            updated += 1
        else:
            connection.execute('INSERT INTO items (parent,name,total,total_missing,start_date,owner,processed,complete,uploaded,parent_only,uploader,upload_date,planned_date,status_override,id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', (*values, str(uuid.uuid4())))
            inserted += 1
    return inserted, updated


def export_workbook(records):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '文件夹明细'
    sheet.append(['親フォルダ', '子フォルダ', 'ファイル数', '開始日', '担当者', '消化件数', '完了', '已上传数量', '上传状态', '名义上传者', '上传日期', '计划上传日期', '文件夹状态', '手动状态'])
    for item in records:
        sheet.append([item['group'], '' if item['parentOnly'] else item['name'], None if item['totalMissing'] else item['total'],
                      item['date'], item['owner'], item['processed'], '〇' if item['complete'] else '未完成', item['uploaded'],
                      '已上传' if item['total'] > 0 and item['uploaded'] == item['total'] else '部分上传' if item['uploaded'] else '未上传', item.get('uploader', ''), item.get('uploadDate', ''),
                      item.get('plannedDate', ''), {'unstarted':'未完成', 'processing':'完成中', 'ready':'已完成/未上传', 'uploading':'上传中', 'uploaded':'已上传'}[folder_status(item)], item.get('statusOverride', '')])
        for cell in sheet[sheet.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = 's'
    for cell in sheet[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='28556A')
    for column, width in {'A': 35, 'B': 38, 'C': 13, 'D': 16, 'E': 16, 'F': 14, 'G': 12, 'H': 16, 'I': 14, 'J': 18, 'K': 16, 'L': 18, 'M': 20, 'N': 18}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = sheet.dimensions
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def export_customer_workbook(records):
    if not CUSTOMER_TEMPLATE.is_file():
        raise FileNotFoundError(f'客户表模板不存在：{CUSTOMER_TEMPLATE}')
    workbook = load_workbook(CUSTOMER_TEMPLATE)
    sheet = workbook.active
    if [sheet.cell(2, column).value for column in range(2, 9)] != ['親フォルダ', '子フォルダ', 'ファイル数', '消化日', '担当者', '消化件数', '完了']:
        raise ValueError('客户表模板的表头不符合预期')
    footer = sheet.max_row
    if sheet.cell(footer, 3).value != '総件数：' or sheet.cell(footer, 6).value != '完了件数：':
        raise ValueError('客户表模板的汇总行不符合预期')

    merged_parents = {}
    for merged in list(sheet.merged_cells.ranges):
        if merged.min_col != 2 or merged.max_col != 2 or merged.min_row < 3 or merged.max_row >= footer:
            raise ValueError('客户表模板含有无法处理的合并单元格')
        for row in range(merged.min_row, merged.max_row + 1):
            merged_parents[row] = sheet.cell(merged.min_row, 2).value
    row_styles = {}
    parent_styles = {}
    group_order = []
    for row in range(3, footer):
        group = text(merged_parents.get(row, sheet.cell(row, 2).value))
        child = text(sheet.cell(row, 3).value)
        if not group:
            raise ValueError(f'客户表模板第 {row} 行缺少父文件夹')
        name = child or group.replace('¥', '\\').replace('/', '\\').split('\\')[-1]
        if group not in parent_styles:
            parent_styles[group] = copy(sheet.cell(row, 2)._style)
            group_order.append(group)
        row_styles[(group, name)] = ([copy(sheet.cell(row, col)._style) for col in range(1, 9)],
                                     sheet.row_dimensions[row].height, bool(sheet.cell(row, 8).value))
    default_styles = [copy(sheet.cell(5, col)._style) for col in range(1, 9)]
    default_parent_style = copy(sheet.cell(43, 2)._style)
    date_style = copy(sheet.cell(5, 5)._style)
    footer_styles = [copy(sheet.cell(footer, col)._style) for col in range(1, 9)]
    footer_height = sheet.row_dimensions[footer].height

    grouped = {group: [] for group in group_order}
    for item in records:
        grouped.setdefault(item['group'], []).append(item)
    for merged in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged))
    sheet.delete_rows(3, footer - 2)

    for group, items_in_group in grouped.items():
        if not items_in_group:
            continue
        by_name = {item['name']: item for item in items_in_group}
        ordered = [by_name.pop(name) for template_group, name in row_styles if template_group == group and name in by_name]
        ordered.extend(item for item in items_in_group if item['name'] in by_name)
        first_row = sheet.max_row + 1
        for item in ordered:
            row = sheet.max_row + 1
            styles, height, had_formula = row_styles.get((group, item['name']), (default_styles, None, True))
            for col, style in enumerate(styles, 1):
                sheet.cell(row, col)._style = copy(style)
            sheet.cell(row, 2)._style = copy(parent_styles.get(group, default_parent_style))
            sheet.row_dimensions[row].height = height
            values = [None, group if row == first_row else None, '' if item['parentOnly'] else item['name'],
                      None if item['totalMissing'] else item['total'],
                      date.fromisoformat(item['uploadDate']) if item['uploaded'] and item.get('uploadDate') else None,
                      ('静' if item.get('uploader') == '张晓静' else item.get('uploader', '')) if item['uploaded'] else None,
                      item['uploaded'] if item['uploaded'] else None]
            for col, value in enumerate(values, 1):
                cell = sheet.cell(row, col, value)
                if isinstance(value, str):
                    cell.data_type = 's'
            if values[4] is not None:
                sheet.cell(row, 5)._style = copy(date_style)
            if had_formula or item['uploaded']:
                sheet.cell(row, 8, f'=IF(AND(D{row}<>"",G{row}<>"",D{row}=G{row}),"〇","")')
        last_row = sheet.max_row
        if last_row > first_row:
            sheet.merge_cells(start_row=first_row, start_column=2, end_row=last_row, end_column=2)

    footer = sheet.max_row + 1
    for col, style in enumerate(footer_styles, 1):
        sheet.cell(footer, col)._style = copy(style)
    sheet.row_dimensions[footer].height = footer_height
    sheet.cell(footer, 3, '総件数：')
    sheet.cell(footer, 4, f'=SUM(D3:D{footer - 1})' if footer > 3 else '=0')
    sheet.cell(footer, 6, '完了件数：')
    sheet.cell(footer, 7, f'=SUM(G3:G{footer - 1})' if footer > 3 else '=0')
    sheet.auto_filter.ref = f'A2:K{footer}'
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


class Handler(BaseHTTPRequestHandler):
    def respond(self, code, body, content_type='application/json; charset=utf-8', headers=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        for name, value in (headers or {}).items():
            self.send_header(name, value)
        self.end_headers()
        self.wfile.write(body)

    def body(self):
        length = int(self.headers.get('Content-Length', '0'))
        if length <= 0 or length > MAX_BODY:
            raise ValueError('文件为空或超过 20 MB')
        return self.rfile.read(length)

    def trusted_origin(self):
        origin = self.headers.get('Origin')
        host = self.headers.get('Host', '')
        return host in (f'127.0.0.1:{PORT}', f'localhost:{PORT}') and (not origin or origin in (f'http://127.0.0.1:{PORT}', f'http://localhost:{PORT}'))

    def do_GET(self):
        if not self.trusted_origin():
            return self.respond(403, {'error': '请求来源不被允许'})
        path = urlparse(self.path).path
        if path == '/api/health':
            return self.respond(200, {'service': 'upload-desk'})
        if path == '/api/items':
            with connect() as connection:
                return self.respond(200, all_items(connection))
        if path == '/api/allocation-history':
            with connect() as connection:
                rows = connection.execute('SELECT id,created_at,planned_date,label,assignments,manifest,length(archive) AS archive_size FROM allocation_history ORDER BY rowid DESC').fetchall()
                return self.respond(200, [{'id':row['id'], 'createdAt':row['created_at'], 'plannedDate':row['planned_date'],
                                           'label':row['label'], 'assignments':json.loads(row['assignments']),
                                           'manifest':json.loads(row['manifest']) if row['manifest'] else None,
                                           'hasArchive':bool(row['archive_size'])} for row in rows])
        if path.startswith('/api/allocation-history/') and path.endswith('/archive'):
            history_id = path.removeprefix('/api/allocation-history/').removesuffix('/archive')
            with connect() as connection:
                row = connection.execute('SELECT id,archive FROM allocation_history WHERE id=?', (history_id,)).fetchone()
            if row is None or not row['archive']:
                return self.respond(404, {'error': '历史方案不存在或没有 ZIP'})
            return self.respond(200, row['archive'], 'application/zip', {'Content-Disposition': f'attachment; filename="upload-allocation-{row["id"]}.zip"'})
        if path == '/api/export':
            with connect() as connection:
                data = export_workbook(all_items(connection))
            return self.respond(200, data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', {'Content-Disposition': 'attachment; filename="upload-desk.xlsx"'})
        if path == '/api/export/json':
            with connect() as connection:
                data = json.dumps(all_items(connection), ensure_ascii=False, indent=2).encode('utf-8')
            return self.respond(200, data, 'application/json; charset=utf-8', {'Content-Disposition': 'attachment; filename="upload-desk-backup.json"'})
        if path == '/api/export/customer':
            try:
                with connect() as connection:
                    data = export_customer_workbook(all_items(connection))
            except (OSError, ValueError) as error:
                return self.respond(500, {'error': f'客户表导出失败：{error}'})
            return self.respond(200, data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', {'Content-Disposition': 'attachment; filename="uploaded-folders.xlsx"'})
        if path in STATIC:
            filename, content_type = STATIC[path]
            try:
                return self.respond(200, (ROOT / filename).read_bytes(), content_type)
            except FileNotFoundError:
                return self.respond(404, {'error': '资源不存在，请先安装依赖'})
        self.respond(404, {'error': '未找到资源'})

    def do_POST(self):
        self.write_request('POST')

    def do_PUT(self):
        self.write_request('PUT')

    def do_DELETE(self):
        self.write_request('DELETE')

    def write_request(self, method):
        if not self.trusted_origin():
            return self.respond(403, {'error': '请求来源不被允许'})
        path = urlparse(self.path).path
        try:
            if method == 'POST' and path == '/api/file-lists':
                sources = json.loads(self.body()).get('sources')
                if not isinstance(sources, list):
                    raise ValueError('sources 格式无效')
                return self.respond(200, {'fileLists': [list_source_files(parts) for parts in sources]})
            if method == 'POST' and path == '/api/allocations/confirm':
                payload = json.loads(self.body())
                with connect() as connection:
                    if isinstance(payload, dict):
                        saved, history_id = save_allocation_history(connection, payload)
                        result = {'items':saved, 'historyId':history_id}
                    else:
                        result = mark_allocations(connection, payload)
                return self.respond(200, result)
            if method == 'POST' and path in ('/api/import/preview', '/api/import'):
                records, sheet = parse_workbook(self.body())
                with connect() as connection:
                    existing = {(row['parent'], row['name']) for row in connection.execute('SELECT parent,name FROM items')}
                    matched = sum((item['group'], item['name']) in existing for item in records)
                    if path.endswith('/preview'):
                        return self.respond(200, {'sheet': sheet, 'total': len(records), 'inserted': len(records) - matched, 'updated': matched})
                    inserted, updated = merge_items(connection, records, preserve_upload=True)
                return self.respond(200, {'inserted': inserted, 'updated': updated})
            if method == 'POST' and path == '/api/migrate':
                records = json.loads(self.body())
                if not isinstance(records, list) or not records:
                    raise ValueError('备份数据为空或格式不正确')
                with connect() as connection:
                    inserted, updated = merge_items(connection, records)
                return self.respond(200, {'inserted': inserted, 'updated': updated})
            if method == 'POST' and path == '/api/items':
                item = validate_item(json.loads(self.body()))
                with connect() as connection:
                    connection.execute('INSERT INTO items (id,parent,name,total,total_missing,start_date,owner,processed,complete,uploaded,parent_only,uploader,upload_date,planned_date,status_override) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                                       (str(uuid.uuid4()),item['group'],item['name'],item['total'],int(item['totalMissing']),item['date'],item['owner'],item['processed'],int(item['complete']),item['uploaded'],int(item['parentOnly']),item['uploader'],date.today().isoformat() if item['uploaded'] else '',item['plannedDate'],item['statusOverride']))
                    row = connection.execute('SELECT * FROM items WHERE rowid=last_insert_rowid()').fetchone()
                    connection.commit()
                    return self.respond(201, as_item(row))
            if path.startswith('/api/items/') and method in ('PUT', 'DELETE'):
                item_id = path.removeprefix('/api/items/')
                payload = json.loads(self.body())
                version = integer(payload.get('version'), '版本')
                with connect() as connection:
                    if method == 'DELETE':
                        result = connection.execute('DELETE FROM items WHERE id=? AND version=?', (item_id, version))
                    else:
                        item = validate_item(payload)
                        result = connection.execute('''UPDATE items SET parent=?,name=?,total=?,total_missing=?,start_date=?,owner=?,processed=?,complete=?,
                            upload_date=CASE WHEN ?=0 THEN '' WHEN uploaded=0 THEN ? ELSE upload_date END,
                            uploaded=?,parent_only=?,uploader=?,planned_date=?,status_override=?,version=version+1 WHERE id=? AND version=?''',
                            (item['group'],item['name'],item['total'],int(item['totalMissing']),item['date'],item['owner'],item['processed'],int(item['complete']),
                             item['uploaded'],date.today().isoformat(),item['uploaded'],int(item['parentOnly']),item['uploader'],item['plannedDate'],item['statusOverride'],item_id,version))
                    if result.rowcount == 0:
                        return self.respond(409, {'error': '记录已被修改或删除，请刷新后重试'})
                    connection.commit()
                    if method == 'DELETE':
                        return self.respond(200, {'ok': True})
                    return self.respond(200, as_item(connection.execute('SELECT * FROM items WHERE id=?', (item_id,)).fetchone()))
            self.respond(404, {'error': '未找到接口'})
        except AllocationConflict as error:
            self.respond(409, {'error': str(error)})
        except (ValueError, KeyError, json.JSONDecodeError, sqlite3.IntegrityError) as error:
            self.respond(400, {'error': str(error) or '数据格式无效'})
        except Exception:
            self.respond(500, {'error': '服务器处理失败，请检查服务器日志'})
            raise


if __name__ == '__main__':
    init_db()
    server = LocalHTTPServer(('127.0.0.1', PORT), Handler)
    print(f'Upload Desk: http://127.0.0.1:{PORT}', flush=True)
    server.serve_forever()

````

### 5.7 `start.py`

<!-- REBUILD_FILE {"path":"start.py","newline":"LF","encoding":"utf-8","size":3727,"sha256":"31946d4f21dc8d5fcaa874ac91128f89b271a9f2bb184ec6a3e900f5659b2091"} -->
````python
import errno
import json
import os
import re
import shutil
import subprocess
import sys
import webbrowser
from importlib import metadata
from pathlib import Path
from urllib.error import URLError
from urllib.request import ProxyHandler, build_opener

ROOT = Path(__file__).resolve().parent
JSZIP = ROOT / 'node_modules' / 'jszip' / 'dist' / 'jszip.min.js'


def configure_stdio():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            try:
                stream.reconfigure(encoding='utf-8', errors='replace')
            except Exception:
                try:
                    stream.reconfigure(errors='replace')
                except Exception:
                    pass


def ensure_dependencies():
    if sys.version_info < (3, 9):
        raise RuntimeError('需要 Python 3.9 或更高版本。')
    try:
        version = metadata.version('openpyxl')
        match = re.match(r'^(\d+)\.(\d+)', version)
        installed = bool(match and (3, 1) <= tuple(map(int, match.groups())) < (4, 0))
    except metadata.PackageNotFoundError:
        installed = False
    if not installed:
        print('正在安装 Python 依赖...', flush=True)
        subprocess.run([sys.executable, '-m', 'pip', 'install', '-r', str(ROOT / 'requirements.txt')], cwd=ROOT, check=True)
    if not JSZIP.is_file():
        npm = shutil.which('npm.cmd') or shutil.which('npm')
        if not npm:
            raise RuntimeError('缺少 Node.js/npm，请先安装 Node.js，然后重新运行。')
        print('正在安装 ZIP 依赖...', flush=True)
        subprocess.run([npm, 'install'], cwd=ROOT, check=True)
        if not JSZIP.is_file():
            raise RuntimeError('ZIP 依赖安装后仍未找到，请检查 npm 的输出。')


def existing_service(url):
    try:
        opener = build_opener(ProxyHandler({}))
        with opener.open(url + '/api/health', timeout=2) as response:
            return json.load(response).get('service') == 'upload-desk'
    except (URLError, OSError, ValueError, AttributeError):
        return False


def open_browser(url):
    print(f'访问地址：{url}', flush=True)
    try:
        opened = webbrowser.open(url)
    except (webbrowser.Error, OSError):
        opened = False
    if not opened:
        print('无法自动打开浏览器，请手动访问上面的地址。', flush=True)


def main():
    configure_stdio()
    ensure_dependencies()
    import server

    url = f'http://127.0.0.1:{server.PORT}'
    try:
        http_server = server.LocalHTTPServer(('127.0.0.1', server.PORT), server.Handler)
    except OSError as error:
        if error.errno in (errno.EADDRINUSE, errno.EACCES, 10013, 10048) and existing_service(url):
            print('服务已经运行。', flush=True)
            open_browser(url)
            return
        if error.errno in (errno.EADDRINUSE, errno.EACCES, 10013, 10048):
            raise RuntimeError(f'端口 {server.PORT} 不可用，请检查是否被占用或设置 PORT 环境变量。') from error
        raise
    with http_server:
        server.init_db()
        print('服务已启动，关闭此窗口将停止服务。', flush=True)
        open_browser(url)
        try:
            http_server.serve_forever()
        except KeyboardInterrupt:
            print('\n服务已停止。', flush=True)


if __name__ == '__main__':
    configure_stdio()
    try:
        main()
    except Exception as error:
        print(f'启动失败：{error}', flush=True)
        if os.name == 'nt' and getattr(sys.stdin, 'isatty', lambda: False)():
            try:
                input('按 Enter 关闭窗口...')
            except EOFError:
                pass
        sys.exit(1)

````

### 5.8 `index.html`

<!-- REBUILD_FILE {"path":"index.html","newline":"LF","encoding":"utf-8","size":16354,"sha256":"d76ed6f7644a53767c2f91b01d03139af4e826ae423bbabd11b837e0f7d7c164"} -->
````html
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="theme-color" content="#1c3f8c">
  <title>文件上传管理 | Upload Desk</title>
  <link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='64' height='64' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='14' fill='%230d7188'/%3E%3Crect x='3' y='3' width='58' height='58' rx='12' fill='%2343cbe6'/%3E%3Cpath d='M10 19h16l5 5h23v21a6 6 0 0 1-6 6H16a6 6 0 0 1-6-6z' fill='%23f4fdff'/%3E%3Cpath d='M32 43V27m-7 7 7-7 7 7' fill='none' stroke='%230b7189' stroke-width='4' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Noto+Sans+SC:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="static/css/style.css">
  <style>
    .owner-filters { display: flex; align-items: flex-start; gap: 12px; padding: 0 21px 16px; }
    .owner-filters-label { flex: none; padding-top: 6px; color: #a0ada2; font-size: 11px; }
    .owner-tags { display: flex; flex-wrap: wrap; gap: 7px; min-width: 0; }
    .owner-tag { max-width: 100%; padding: 5px 10px; border: 1px solid #e6ece4; border-radius: 5px; background: linear-gradient(#fff, #f8faf7); color: #79897c; font-size: 10px; overflow-wrap: anywhere; }
    .owner-tag:hover { border-color: #b6cbb9; color: #54745e; }
    .owner-tag[aria-pressed="true"] { border-color: #c9ddcb; background: #e8f1e7; color: #54785e; box-shadow: inset 0 1px 2px #718c741a; }
    @media (max-width: 700px) { html { overflow-x: clip; } .owner-filters { padding: 0 13px 13px; } }
  </style>
  <link rel="stylesheet" href="node_modules/flatpickr/dist/flatpickr.min.css">
  <link rel="stylesheet" href="static/css/game-style.css">
  <link rel="stylesheet" href="static/css/planner.css">
  <link rel="stylesheet" href="static/css/cartoon-style.css">
  <script src="https://unpkg.com/lucide@0.468.0/dist/umd/lucide.min.js" defer></script>
  <script src="node_modules/jszip/dist/jszip.min.js" defer></script>
  <script src="node_modules/flatpickr/dist/flatpickr.min.js" defer></script>
  <script src="node_modules/flatpickr/dist/l10n/zh.js" defer></script>
  <script src="static/js/app.js" defer></script>
  <script src="static/js/planner.js" defer></script>
</head>
<body>
  <div class="app-shell">
    <aside class="sidebar">
      <div class="brand"><div class="brand-mark"><i data-lucide="layers-2"></i></div><div><strong>Upload Desk</strong><span>文件管理工作台</span></div></div>
      <div class="sidebar-label">工作空间 <span>WORKSPACE</span></div>
      <nav class="navigation" aria-label="文件夹导航">
        <button class="nav-item active" data-group="all"><i data-lucide="layout-dashboard"></i><span>总览</span><i class="nav-arrow" data-lucide="chevron-right"></i></button>
        <div class="nav-caption">项目文件夹 <button id="add-group" class="subtle-icon" title="新建文件夹" aria-label="新建文件夹"><i data-lucide="plus"></i></button></div>
        <div id="folder-nav"></div>
      </nav>
      <div class="sidebar-bottom"><div class="storage-icon"><i data-lucide="hard-drive"></i></div><div><strong>本地工作空间</strong><span id="storage-status">正在连接数据库</span></div><span class="online-dot"></span></div>
    </aside>

    <main class="main-content">
      <header class="topbar"><div class="breadcrumb"><span>工作空间</span><i data-lucide="chevron-right"></i><strong id="breadcrumb-current">总览</strong></div><div class="topbar-right"><span class="today"><i data-lucide="calendar-days"></i><span id="today"></span></span><div class="avatar">管</div></div></header>
      <div class="content">
        <section class="page-heading"><div><div class="eyebrow"><span></span> WORKSPACE OVERVIEW</div><h1>文件上传管理 <span class="heading-spark">✳</span></h1><p>所有文件的上传进度，一目了然。</p></div><button id="add-item" class="button primary"><i data-lucide="plus"></i> 添加文件夹</button></section>

        <section class="overview" aria-label="上传概况">
          <div class="overview-intro"><div class="section-kicker">PROJECT PROGRESS <span>01 / OVERVIEW</span></div><div class="overview-title"><div><h2>项目进度概览</h2><p>掌握每一份文件的上传情况</p></div><div class="mini-chart"><i data-lucide="chart-no-axes-column-increasing"></i></div></div></div>
          <div class="metric"><div class="metric-label"><span class="metric-icon blue"><i data-lucide="files"></i></span> 文件总数</div><div class="metric-value"><strong id="total-files">0</strong><span>个文件</span></div><div class="metric-foot">当前范围内的全部文件</div></div>
          <div class="metric"><div class="metric-label"><span class="metric-icon green"><i data-lucide="circle-check"></i></span> 已上传</div><div class="metric-value"><strong id="uploaded-files">0</strong><span>个文件</span></div><div class="metric-foot"><span class="tiny-dot green-dot"></span> 已完成上传的文件</div></div>
          <div class="metric progress-metric"><div class="metric-label"><span class="metric-icon amber"><i data-lucide="chart-pie"></i></span> 上传完成率</div><div class="metric-value"><strong id="overall-percent">0%</strong></div><div class="progress-line"><span id="overall-bar"></span></div><div class="metric-foot">每一次更新，都离完成更近一步</div></div>
        </section>

        <div id="plan-preview-spacer" aria-hidden="true"></div>
        <section class="planner" aria-labelledby="planner-title">
          <div class="planner-heading"><div><div class="section-kicker">DAILY PLAN <span>02 / ALLOCATION</span></div><h2 id="planner-title">上传分配</h2></div><span class="planner-badge"><i data-lucide="calendar-check"></i> 按子文件夹分配</span></div>
          <form id="planner-form" class="planner-form">
            <label>计划上传日期<input id="plan-date" type="text" autocomplete="off" placeholder="YYYY-MM-DD" data-datepicker required></label>
            <label>目标文件数<input id="plan-target" type="number" min="1" max="1000000" value="150" required></label>
            <div class="planner-field planner-people"><span id="plan-people-label">分配给</span><details class="person-picker" id="plan-people-picker"><summary aria-labelledby="plan-people-label plan-people-summary"><span id="plan-people-summary"></span><i data-lucide="chevron-down"></i></summary><div class="person-picker-panel"><div class="person-picker-heading"><span>分配人员</span><strong id="plan-people-count" aria-live="polite"></strong></div><div class="person-picker-list" id="plan-people-options" role="group" aria-label="分配人员"></div><div class="person-picker-add"><input id="plan-person-add" type="text" maxlength="40" placeholder="添加其他人员" aria-label="添加其他人员"><button id="plan-person-add-button" type="button">添加</button></div></div></details></div>
            <div class="planner-field"><span id="plan-owner-label">来源负责人</span><details class="person-picker" id="plan-owner-picker"><summary aria-labelledby="plan-owner-label plan-owner-summary"><span id="plan-owner-summary">全部负责人</span><i data-lucide="chevron-down"></i></summary><div class="person-picker-panel"><div class="person-picker-heading"><span>来源负责人</span><strong id="plan-owner-count" aria-live="polite">全部</strong></div><div class="person-picker-list" id="plan-owner-options" role="group" aria-label="来源负责人"></div><button id="plan-owner-clear" class="person-picker-clear" type="button">清除选择</button></div></details></div>
            <button type="submit" class="button primary"><i data-lucide="sparkles"></i> 生成方案</button>
          </form>
          <p id="planner-message" class="planner-message" role="status">从已完成且未上传的子文件夹中，选择目标文件数上下 5 个以内的方案，优先接近目标。</p>
          <div id="planner-results" hidden>
            <div class="planner-result-heading"><div id="plan-options" class="plan-options" role="group" aria-label="分配方案"></div><div id="plan-hover-preview" class="plan-hover-preview" role="tooltip" hidden></div><button id="plan-create" type="button" class="button primary"><i data-lucide="clipboard-copy"></i> 确定并复制方案 JSON</button></div>
            <div id="plan-summary" class="plan-summary"></div>
            <p class="planner-note">确认后保存上传计划，文件夹进入上传中；实际上传完成后请手动切换为已上传。</p>
          </div>
          <details id="allocation-history" class="allocation-history">
            <summary><i data-lucide="history"></i><span>历史分配方案</span><i data-lucide="chevron-down"></i></summary>
            <div class="history-toolbar"><strong>已保存的方案</strong><button id="history-refresh" type="button" class="icon-button" title="刷新历史方案" aria-label="刷新历史方案"><i data-lucide="refresh-cw"></i></button></div>
            <div id="history-list" class="history-list" aria-live="polite"></div>
          </details>
        </section>

        <section class="list-section"><div class="list-heading"><div><div class="section-kicker">FILE DIRECTORY <span>03 / DETAILS</span></div><h2 id="list-title">文件夹明细 <span id="row-count">0</span></h2><p>查看并管理各文件夹的上传状态</p></div><div class="data-actions"><button id="import-excel" class="button secondary"><i data-lucide="file-down"></i> 导入 Excel</button><button id="export" class="button secondary"><i data-lucide="file-up"></i> 导出完整 Excel</button><button id="export-customer" class="button secondary"><i data-lucide="file-up"></i> 导出客户表</button><button id="import-backup" class="button secondary"><i data-lucide="file-down"></i> 导入 JSON 备份</button><button id="export-backup" class="button secondary"><i data-lucide="file-up"></i> 导出 JSON 备份</button><input id="excel-file" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden><input id="backup-file" type="file" accept=".json,application/json" hidden></div></div>
          <div id="plan-preview-bar" class="plan-preview-bar" hidden><span>正在预览：<strong id="plan-preview-name"></strong></span><button id="exit-plan-preview" type="button" class="button secondary">退出预览</button></div>
          <div id="date-history-bar" class="date-history-bar" aria-live="polite" hidden></div>
          <div class="table-card"><div class="table-toolbar"><label class="search-box"><i data-lucide="search"></i><input id="search" type="search" placeholder="搜索文件夹或负责人..." aria-label="搜索文件夹或负责人"><kbd>⌘ K</kbd></label><div class="toolbar-actions"><label class="upload-date-filter"><span>上传日期</span><input id="upload-date-filter" type="text" autocomplete="off" placeholder="YYYY-MM-DD" data-datepicker aria-label="按上传日期筛选"></label><button id="clear-upload-date" type="button" class="clear-upload-date" aria-label="清除上传日期筛选" hidden>清除</button><label class="filter-select"><i data-lucide="list-filter"></i><select id="status-filter" aria-label="按文件夹状态筛选"><option value="all">全部状态</option><option value="unstarted">未完成</option><option value="processing">完成中...</option><option value="ready">已完成/未上传</option><option value="uploading">上传中...</option><option value="uploaded">已上传</option></select><i data-lucide="chevron-down"></i></label><span class="toolbar-divider"></span><button id="sort-order" class="icon-button" title="切换排序方向" aria-label="切换排序方向"><i data-lucide="arrow-down-wide-narrow"></i></button></div></div>
            <div class="owner-filters"><span class="owner-filters-label">负责人</span><div id="owner-tags" class="owner-tags" role="group" aria-label="按负责人筛选"></div></div>
            <div class="tree-actions"><button id="expand-all-folders" type="button" class="button secondary"><i data-lucide="chevrons-up-down"></i> 全部展开</button><button id="collapse-all-folders" type="button" class="button secondary"><i data-lucide="chevrons-down-up"></i> 全部折叠</button></div>
            <div class="table-wrap"><table><thead><tr><th><button data-sort="name">文件夹名称 <i data-lucide="arrow-up-down"></i></button></th><th><button data-sort="total">文件数量 <i data-lucide="arrow-up-down"></i></button></th><th>完成数量</th><th><button data-sort="date">开始日期 <i data-lucide="arrow-up-down"></i></button></th><th><button data-sort="owner">负责人 <i data-lucide="arrow-up-down"></i></button></th><th>名义上传者</th><th><button data-sort="uploadDate">上传日期 <i data-lucide="arrow-up-down"></i></button></th><th>上传进度</th><th><button data-sort="plannedDate">计划上传日期 <i data-lucide="arrow-up-down"></i></button></th><th><button data-sort="status">文件夹状态 <i data-lucide="arrow-up-down"></i></button></th><th><span class="sr-only">操作</span></th></tr></thead><tbody id="table-body"></tbody></table><div id="empty-state" class="empty-state" hidden><i data-lucide="folder-search"></i><strong>没有找到相关文件夹</strong><span>试试调整搜索内容或筛选条件</span></div></div>
            <div class="table-footer"><span id="table-summary"></span></div>
          </div>
        </section>
        <footer class="page-footer"><span>UPLOAD DESK <span> / </span> 文件管理工作台</span><span>让每一份进度都有迹可循 <span class="footer-star">✳</span></span></footer>
      </div>
    </main>
  </div>
  <dialog id="item-dialog"><form id="item-form"><div class="dialog-header"><div><span class="section-kicker">FOLDER DETAILS</span><h2 id="dialog-title">添加文件夹</h2></div><button type="button" class="icon-button" id="close-dialog" aria-label="关闭"><i data-lucide="x"></i></button></div><div class="form-grid"><label class="full">子文件夹名称<input name="name" maxlength="200" placeholder="留空表示父文件夹记录"><select name="name-choice" aria-label="子文件夹名称" required hidden disabled></select></label><label class="full">所属父文件夹<select name="group-choice" aria-label="所属父文件夹" required></select><input name="group" aria-label="新建父文件夹名称" maxlength="500" placeholder="输入新父文件夹路径" hidden disabled></label><label>文件数量<input name="total" type="number" min="0" max="1000000" step="1" placeholder="未填写"></label><label>完成数量（消化件数）<input name="processed" type="number" min="0" max="1000000" step="1" placeholder="未填写"></label><label>已上传数量<input name="uploaded" type="number" min="0" max="1000000" step="1" required value="0"></label><label>开始日期<input name="date" type="text" autocomplete="off" placeholder="YYYY-MM-DD" data-datepicker></label><label>负责人<select name="owner-choice" aria-label="负责人"></select><input name="owner" aria-label="新建负责人名称" maxlength="100" placeholder="输入新负责人" hidden disabled></label></div><p id="form-error" class="form-error" role="alert"></p><div class="dialog-actions"><button type="button" id="cancel-dialog" class="button secondary">取消</button><button type="submit" class="button primary">保存更改</button></div></form></dialog>
  <dialog id="confirm-dialog" class="confirm-dialog" aria-labelledby="confirm-title" aria-describedby="confirm-message"><form method="dialog"><div class="dialog-header"><div><span class="section-kicker">CONFIRM</span><h2 id="confirm-title">请确认</h2></div></div><p id="confirm-message" class="confirm-message"></p><div class="dialog-actions"><button type="button" id="confirm-cancel" class="button secondary">取消</button><button type="button" id="confirm-ok" class="button primary">确定</button></div></form></dialog>
  <div id="toast" class="toast" role="status"></div>
</body>
</html>

````

### 5.9 `static\css\style.css`

<!-- REBUILD_FILE {"path":"static\\css\\style.css","newline":"LF","encoding":"utf-8","size":13579,"sha256":"8db260c977337d402b3440302a63cedc3739d1076d995542faba62b6641d1c85"} -->
````css
:root{font-family:"DM Sans","Noto Sans SC",sans-serif;color:#293334;background:#f2f1ec;font-synthesis:none;--ink:#293334;--muted:#8b9692;--line:#e8ebe6;--green:#597b69;--orange:#d49d62}*{box-sizing:border-box}body{margin:0}button,input,select{font:inherit}button{cursor:pointer}button:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid #6c947c;outline-offset:2px}svg{width:18px;height:18px;stroke-width:1.7}.app-shell{display:flex;min-height:100vh}.sidebar{width:232px;flex:none;background:linear-gradient(180deg,#f9faf7,#f3f5f0);border-right:1px solid #dfe4dc;display:flex;flex-direction:column;padding:28px 17px 18px;box-shadow:3px 0 12px #37443a08}.brand{display:flex;align-items:center;gap:11px;padding:0 6px 41px}.brand-mark{width:39px;height:39px;display:grid;place-items:center;color:#fff;background:linear-gradient(145deg,#718c7a,#4c6b59);border:1px solid #526f5d;border-radius:11px;box-shadow:inset 0 1px 1px #ffffff80,0 3px 5px #52685630}.brand-mark svg{width:21px;height:21px}.brand strong,.brand span,.sidebar-bottom strong,.sidebar-bottom span{display:block}.brand strong{font-size:15px;letter-spacing:.1px}.brand span{font-size:10px;color:#99a29b;margin-top:3px}.sidebar-label,.nav-caption{font-size:10px;color:#a6afa7;font-weight:700;letter-spacing:.5px;padding:0 9px}.sidebar-label span{font-size:9px;font-weight:500;float:right;letter-spacing:.8px}.navigation{margin-top:15px}.nav-item{width:100%;height:39px;display:flex;align-items:center;gap:11px;border:0;background:transparent;color:#77857a;border-radius:6px;padding:0 10px;text-align:left;font-size:12px;font-weight:500}.nav-item svg{width:16px;height:16px;flex:none}.nav-item span{min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.nav-item.active{color:#4d705b;background:#e4ede3;box-shadow:inset 0 0 0 1px #dce8da}.nav-item.active:before{content:"";position:absolute;left:17px;width:3px;height:18px;background:#769982;border-radius:3px}.nav-arrow{margin-left:auto;width:13px!important}.nav-caption{display:flex;align-items:center;justify-content:space-between;margin:31px 0 10px}.subtle-icon,.icon-button{border:0;background:transparent;display:grid;place-items:center;color:#9daaa0}.subtle-icon svg{width:15px;height:15px}#folder-nav .nav-item{font-size:11px}.sidebar-bottom{margin-top:auto;border-top:1px solid #e0e6dd;padding:20px 4px 0;display:flex;align-items:center;gap:9px}.storage-icon{width:32px;height:32px;background:#e7ece4;border:1px solid #dce4da;border-radius:7px;display:grid;place-items:center;color:#75907a}.storage-icon svg{width:15px}.sidebar-bottom strong{font-size:10px}.sidebar-bottom span{font-size:9px;color:#a3aca4;margin-top:4px}.online-dot{margin-left:auto;width:6px;height:6px;background:#83ad88;border-radius:50%}.main-content{min-width:0;flex:1}.topbar{height:66px;background:#fcfcfa;border-bottom:1px solid #e4e7e0;display:flex;justify-content:space-between;align-items:center;padding:0 42px}.breadcrumb,.topbar-right,.today{display:flex;align-items:center;gap:11px}.breadcrumb{font-size:11px;color:#a2ada4}.breadcrumb strong{color:#536257;font-weight:600}.breadcrumb svg{width:13px}.topbar-right{gap:20px}.today{font-size:11px;color:#87968a}.today svg{width:15px}.avatar{width:29px;height:29px;background:#dce7dc;border:2px solid white;box-shadow:0 0 0 1px #d6e0d5;border-radius:50%;display:grid;place-items:center;color:#63806b;font-size:11px;font-weight:700}.content{max-width:1510px;margin:auto;padding:43px 42px 24px}.page-heading{display:flex;justify-content:space-between;align-items:flex-end;margin-bottom:29px}.eyebrow,.section-kicker{font-size:9px;letter-spacing:1.1px;font-weight:700;color:#9ba99e}.eyebrow{display:flex;align-items:center;gap:7px}.eyebrow span{width:6px;height:6px;border-radius:50%;background:#83a78b}.page-heading h1{font-size:26px;letter-spacing:0;margin:10px 0 7px;font-weight:700;color:#293932}.heading-spark{color:#d8aa70;font-size:17px;margin-left:4px}.page-heading p,.list-heading p,.overview-title p{color:#9ba69e;font-size:11px;margin:0}.button{height:36px;border-radius:6px;padding:0 15px;display:inline-flex;align-items:center;justify-content:center;gap:7px;font-size:11px;font-weight:600;white-space:nowrap}.button svg{width:15px;height:15px}.primary{color:white;background:linear-gradient(#789785,#668875);border:1px solid #63816e;box-shadow:inset 0 1px #ffffff4d,0 2px 3px #4c695128}.primary:hover{background:#5e806b}.secondary{background:linear-gradient(#fff,#f8f9f6);border:1px solid #dfe5dc;color:#657469;box-shadow:0 2px 3px #3a4b3810}.secondary:hover{border-color:#aabdad}.overview{background:#fbfcf9;border:1px solid #e0e6dd;border-radius:8px;box-shadow:0 3px 8px #485d4410,inset 0 1px #fff;display:grid;grid-template-columns:1.25fr 1fr 1fr 1.15fr;padding:22px 0;margin-bottom:39px}.overview-intro,.metric{padding:0 25px;min-width:0}.metric{border-left:1px solid #e9eee7}.section-kicker span{font-size:8px;font-weight:500;color:#bdc7be;margin-left:8px}.overview-title{display:flex;justify-content:space-between;align-items:center;margin-top:21px}.overview-title h2,.list-heading h2{margin:0 0 7px;color:#35453b;font-size:15px}.mini-chart{height:37px;width:37px;border:1px solid #e5ebe3;background:#f4f7f1;border-radius:8px;display:grid;place-items:center;color:#8aa393}.metric-label{display:flex;align-items:center;gap:8px;color:#829086;font-size:11px}.metric-icon{width:26px;height:26px;border-radius:6px;display:grid;place-items:center}.metric-icon svg{width:14px;height:14px}.blue{color:#8199a0;background:#ebf1f0}.green{color:#81a48c;background:#eaf2e9}.amber{color:#c6a370;background:#f7f0e2}.metric-value{display:flex;align-items:baseline;gap:7px;margin:13px 0 8px}.metric-value strong{font-size:25px;line-height:1;font-weight:700;color:#34443a}.metric-value span,.metric-foot{font-size:10px;color:#a7b1a9}.tiny-dot{display:inline-block;width:5px;height:5px;border-radius:50%;margin-right:4px}.green-dot{background:#91b99a}.progress-line,.row-progress{height:5px;border-radius:10px;background:#e9eee7;overflow:hidden}.progress-line{margin:11px 0 8px}.progress-line span,.row-progress span{display:block;height:100%;background:linear-gradient(90deg,#91b59b,#6f9a7b);border-radius:10px}.list-heading{display:flex;align-items:flex-end;justify-content:space-between;margin-bottom:18px}.list-heading h2{font-size:18px;margin:9px 0 6px}.list-heading h2 span{display:inline-grid;place-items:center;vertical-align:2px;margin-left:4px;background:#e6eee5;color:#77937c;border-radius:4px;font-size:10px;min-width:21px;height:19px}.table-card{background:#fff;border:1px solid #e0e6dd;border-radius:8px;box-shadow:0 3px 9px #485d4410;overflow:hidden}.table-toolbar{min-height:68px;padding:16px 21px;display:flex;justify-content:space-between;gap:12px;align-items:center}.search-box{height:35px;width:260px;max-width:100%;display:flex;align-items:center;gap:9px;padding:0 10px;background:#f9faf8;border:1px solid #e8ede7;border-radius:5px;color:#a6b1a8}.search-box svg{width:15px;flex:none}.search-box input{width:100%;min-width:0;border:0;outline:0;background:transparent;font-size:11px;color:#48584d}.search-box input::placeholder{color:#b3bdb5}.search-box kbd{font-family:inherit;font-size:9px;white-space:nowrap;border:1px solid #e4e9e3;border-radius:3px;background:#fff;padding:2px 4px}.toolbar-actions{display:flex;align-items:center;gap:14px}.filter-select{display:flex;align-items:center;gap:7px;color:#7e8d81}.filter-select svg{width:14px}.filter-select select{appearance:none;border:0;background:transparent;color:#657469;font-size:11px;cursor:pointer;outline:0}.toolbar-divider{width:1px;height:20px;background:#e7ece6}.icon-button{height:30px;width:30px;border-radius:5px}.icon-button:hover,.subtle-icon:hover{background:#edf3eb;color:#668473}.icon-button svg{width:16px;height:16px}.table-wrap{overflow-x:auto}table{width:100%;border-collapse:collapse;min-width:850px;text-align:left}thead{background:#f8faf7;border-top:1px solid #eef1ed;border-bottom:1px solid #e9eee8}th{height:39px;color:#a0ada2;font-weight:500;font-size:10px;white-space:nowrap}th:first-child,td:first-child{padding-left:22px}th:last-child,td:last-child{padding-right:18px}th button{border:0;padding:0;display:flex;align-items:center;gap:5px;background:none;color:inherit;font-size:inherit}th button svg{width:11px;height:11px}td{height:60px;border-bottom:1px solid #f0f3ef;font-size:11px;color:#718075}tbody tr:last-child td{border-bottom:0}tbody tr:hover{background:#fafcf9}.folder-cell{display:flex;align-items:center;gap:11px;min-width:230px}.folder-icon{width:33px;height:33px;display:grid;place-items:center;background:linear-gradient(#fbf4e7,#f4e9d8);border:1px solid #f0e4d1;border-radius:6px;color:#cda16b;flex:none}.folder-icon svg{width:16px;fill:#d7ad7830}.folder-name{display:block;max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px;color:#45574a;font-weight:600}.folder-parent{display:block;margin-top:4px;font-size:9px;color:#aab5ac}.file-count{font-weight:600;color:#687a6d}.owner{display:flex;align-items:center;gap:7px}.owner-avatar{width:24px;height:24px;display:grid;place-items:center;background:#eaf0e8;color:#819687;border-radius:50%;font-size:10px;font-weight:700}.owner-avatar.empty{background:#f3f4f1;color:#bdc4be}.progress-cell{width:150px;min-width:140px;padding-right:24px}.progress-info{display:flex;justify-content:space-between;gap:12px;margin-bottom:7px;color:#a3afa5;font-size:9px}.progress-info strong{color:#718c78;font-weight:600}.row-progress{height:4px}.status{border:1px solid transparent;border-radius:4px;padding:4px 7px;font-size:10px;white-space:nowrap}.status:before{content:"";display:inline-block;width:5px;height:5px;border-radius:50%;margin:0 5px 1px 0;background:currentColor}.status.done{color:#709b79;background:#edf5ed;border-color:#e4f0e4}.status.progress{color:#c79b61;background:#fbf4e9;border-color:#f6ead7}.status.pending{color:#9aa7a0;background:#f2f4f1;border-color:#ebeeea}.row-actions{display:flex;justify-content:flex-end;gap:2px}.row-actions .icon-button{width:27px;height:27px}.table-footer{height:47px;border-top:1px solid #edf1ec;padding:0 22px;display:flex;align-items:center;justify-content:space-between;color:#a3aea5;font-size:10px}.footer-hint{display:flex;align-items:center;gap:6px}.footer-hint svg{width:12px}.empty-state{min-width:850px;padding:48px;text-align:center;color:#a3afa4}.empty-state:not([hidden]){display:flex;align-items:center;flex-direction:column;gap:8px}.empty-state svg{width:30px;height:30px;color:#b9c9b9}.empty-state strong{color:#738477;font-size:12px}.empty-state span{font-size:11px}.page-footer{display:flex;justify-content:space-between;padding-top:24px;color:#b5bfb6;font-size:9px;letter-spacing:.3px}.page-footer span span{color:#d5ddd4}.footer-star{color:#d5ae79!important;font-size:13px}.sr-only{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:nowrap;border:0}dialog{border:1px solid #dfe7dd;border-radius:9px;padding:0;width:min(460px,calc(100% - 32px));box-shadow:0 20px 70px #26392b35;color:#34443a}dialog::backdrop{background:#26372db3}dialog form{padding:24px}.dialog-header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:22px}.dialog-header h2{font-size:19px;margin:8px 0 0}.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:15px}.form-grid label{font-size:11px;color:#78877b;font-weight:600}.form-grid .full{grid-column:1/-1}.form-grid input{display:block;width:100%;height:38px;border:1px solid #e0e7df;border-radius:5px;padding:0 10px;margin-top:7px;color:#405347;background:#fbfcfa}.form-error{font-size:11px;color:#bd6d59;min-height:14px;margin:12px 0}.dialog-actions{display:flex;justify-content:flex-end;gap:8px}.toast{position:fixed;right:24px;bottom:24px;background:#405c4b;color:#fff;border-radius:6px;padding:10px 15px;font-size:12px;box-shadow:0 5px 20px #354a3930;opacity:0;transform:translateY(8px);transition:.2s;pointer-events:none}.toast.show{opacity:1;transform:translateY(0)}@media(max-width:1050px){.overview{grid-template-columns:1fr 1fr}.overview-intro,.metric{padding:15px 25px}.overview-intro,.metric:nth-child(2){border-bottom:1px solid #e9eee7}.metric:nth-child(3){border-left:0}.overview-title{margin-top:12px}.overview{padding:7px 0}}@media(max-width:700px){.app-shell{display:block}.sidebar{width:100%;padding:12px 16px;display:block;border-right:0;border-bottom:1px solid #dfe4dc}.brand{padding:0}.brand-mark{width:33px;height:33px}.brand strong{font-size:13px}.sidebar-label,.nav-caption,.sidebar-bottom{display:none}.navigation{margin-top:10px;display:flex;overflow-x:auto;gap:4px}.nav-item{width:auto;flex:none;padding:0 10px;height:34px}.nav-item.active:before,.nav-arrow{display:none}#folder-nav{display:flex}.topbar{height:50px;padding:0 18px}.today{display:none}.content{padding:28px 16px}.page-heading{align-items:flex-start;gap:10px}.page-heading h1{font-size:22px}.page-heading .button{margin-top:18px}.overview{margin-bottom:30px}.overview-intro,.metric{padding:14px}.metric-value strong{font-size:22px}.mini-chart{display:none}.list-heading h2{font-size:16px}.table-toolbar{padding:13px;flex-wrap:wrap}.search-box{width:100%}.toolbar-actions{width:100%;justify-content:flex-end}.table-footer{padding:0 13px}.footer-hint{display:none}.page-footer{gap:10px}.page-footer>span:last-child{text-align:right}}@media(max-width:380px){.page-heading{flex-direction:column}.page-heading .button{margin-top:0}.overview{grid-template-columns:1fr 1fr}.overview-title p,.metric-foot{font-size:9px}.page-footer{font-size:8px}}

````

### 5.10 `static\css\game-style.css`

<!-- REBUILD_FILE {"path":"static\\css\\game-style.css","newline":"LF","encoding":"utf-8","size":17900,"sha256":"91f7819091242e527a371cd90ce9f60b36c563473ae5ca5b9adf96808c74968d"} -->
````css
:root {
  color: #35505e;
  background: #f0f8f8;
  --ink: #35505e;
  --muted: #829ba2;
  --line: #dfedf0;
}

body {
  background-color: #f0f8f8;
  background-image: radial-gradient(#c9e8e8 1px, transparent 1px);
  background-size: 20px 20px;
}

button:focus-visible, input:focus-visible, select:focus-visible {
  outline: 3px solid #ffbc41;
  outline-offset: 3px;
}

.sidebar {
  background: linear-gradient(180deg, #f8ffff 0%, #e8f7f5 100%);
  border-right: 2px solid #cde9e8;
  box-shadow: 4px 0 0 #ffffffa6, 8px 0 18px #65aab418;
}

.brand-mark {
  color: white;
  background: linear-gradient(#42d8fa 0%, #0ba8e9 55%, #0787cc 100%);
  border: 2px solid #078fc8;
  border-radius: 12px;
  box-shadow: inset 0 3px 1px #ffffff80, inset 0 -4px 0 #057ab3, 0 4px 5px #178ab044;
}

.brand strong { color: #2f6176; font-size: 16px; }
.brand span, .sidebar-label, .nav-caption { color: #8caab1; }
.nav-item { color: #66878f; font-size: 12px; font-weight: 600; }
.nav-item:hover { background: #e0f5f7; color: #238aa8; }
.nav-item.active {
  color: white;
  background: linear-gradient(#40c8f1, #0da7dc 60%, #078fc4);
  box-shadow: inset 0 2px 0 #ffffff80, inset 0 -3px 0 #077cae, 0 3px 5px #148caf35;
  text-shadow: 0 1px 1px #08749c;
}
.nav-item.active:before { display: none; }
.sidebar-bottom { border-color: #cfe7e7; }
.storage-icon { color: #1b9cc0; background: linear-gradient(#e9fdff, #cceff6); border-color: #b6e4ed; }
.sidebar-bottom span { color: #91aaae; }
.online-dot { background: #55c72f; box-shadow: 0 0 0 3px #dff4d7; }

.topbar { background: #ffffffeb; border-bottom: 2px solid #e1f1f0; }
.breadcrumb { color: #91aab0; }
.breadcrumb strong { color: #4a7685; }
.today { color: #789aa4; }
.avatar {
  background: linear-gradient(#ffe17c, #ffae25);
  color: #895609;
  border: 2px solid #fff2ca;
  box-shadow: 0 2px 0 #e69b18, 0 3px 5px #ba891c44;
}

.eyebrow, .section-kicker { color: #67b4c6; }
.eyebrow span { background: #6ed536; box-shadow: 0 0 0 3px #dbf5ce; }
.section-kicker span { color: #a4c6cb; }
.page-heading h1 { color: #325765; font-weight: 700; }
.heading-spark, .footer-star { color: #ffb930 !important; }
.page-heading p, .list-heading p, .overview-title p { color: #8aa8ae; }

.button {
  height: 39px;
  border-radius: 8px;
  font-size: 12px;
  font-weight: 700;
  transition: transform .12s, box-shadow .12s;
}
.button:active, .owner-tag:active, .status:active { transform: translateY(2px); }
.primary, .primary:hover {
  color: white;
  background: linear-gradient(#81ea31 0%, #55cf0e 48%, #39b700 100%);
  border: 2px solid #42b908;
  box-shadow: inset 0 3px 0 #ffffff75, inset 0 -4px 0 #299b00, 0 4px 5px #3a9d3750;
  text-shadow: 0 1px 1px #278b09;
}
.secondary, .secondary:hover {
  color: #417d91;
  background: linear-gradient(#fff, #eaf9fb);
  border: 2px solid #c3e6ed;
  box-shadow: inset 0 2px 0 #fff, inset 0 -3px 0 #c6e7ed, 0 3px 4px #5da3b62b;
}

.overview, .table-card {
  background: #fff;
  border: 2px solid #d5ebec;
  border-radius: 8px;
  box-shadow: 0 5px 0 #c6e3e3, 0 8px 15px #70aeb51c, inset 0 2px 0 #fff;
}
.overview { padding: 22px 0; }
.overview-intro, .metric { padding-left: 25px; padding-right: 25px; }
.metric { border-color: #e5f1f0; }
.overview-title h2, .list-heading h2 { color: #375d69; }
.mini-chart, .metric-icon {
  border-radius: 8px;
  box-shadow: inset 0 2px 0 #ffffffb3, 0 2px 0 #00000012;
}
.mini-chart { background: linear-gradient(#e5faff, #bcecf7); color: #20a7cd; border: 1px solid #a8e3f0; }
.metric-icon.blue { background: linear-gradient(#6fddfb, #20b3e8); color: white; }
.metric-icon.green { background: linear-gradient(#a2ec5a, #5fc82b); color: white; }
.metric-icon.amber { background: linear-gradient(#ffe17a, #ffb82e); color: white; }
.metric-label { color: #7898a0; font-weight: 600; }
.metric-value strong { color: #375966; }
.metric-value span, .metric-foot { color: #a0b7ba; }
.green-dot { background: #6ed33d; }
.progress-line, .row-progress {
  height: 8px;
  background: #dcebed;
  border: 1px solid #c8dfe1;
  box-shadow: inset 0 1px 2px #81a8ad40;
}
.progress-line span, .row-progress span {
  background: linear-gradient(#75eaff, #05bbeb 55%, #079ecc);
  box-shadow: inset 0 1px 0 #ffffffa6;
}
.list-heading h2 span {
  color: white;
  background: linear-gradient(#ffca54, #f5a928);
  box-shadow: inset 0 1px 0 #fff8, 0 2px 0 #df9418;
}

.table-toolbar { padding-top: 20px; }
.search-box {
  height: 38px;
  color: #83b1bc;
  background: #f1f9fa;
  border: 2px solid #d4ebed;
  box-shadow: inset 0 2px 3px #79abb31f;
}
.search-box input { color: #426d79; font-size: 12px; }
.search-box input::placeholder { color: #a2bdc1; }
.search-box kbd { border-color: #dceced; color: #9ab5ba; }
.filter-select { color: #70a6b0; }
.filter-select select { color: #557f89; font-size: 12px; font-weight: 600; }
.toolbar-divider { background: #deeeee; }
.icon-button, .subtle-icon { color: #80aab1; }
.icon-button:hover, .subtle-icon:hover { color: #168fb5; background: #e6f8fb; }
.owner-filters-label { color: #83a3a9; font-weight: 700; }
.owner-tag {
  color: #6f929a;
  background: linear-gradient(#fff, #edf8fa);
  border: 1px solid #cfe8ed;
  border-radius: 6px;
  box-shadow: inset 0 2px 0 #fff, 0 2px 0 #d3e8ea;
  font-size: 11px;
  font-weight: 600;
}
.owner-tag:hover { color: #1687ad; border-color: #94d8e9; }
.owner-tag[aria-pressed="true"] {
  color: white;
  background: linear-gradient(#54d6fa, #0cafe3 60%, #0796cc);
  border-color: #0c9dcd;
  box-shadow: inset 0 2px 0 #ffffff80, 0 2px 0 #087fac;
  text-shadow: 0 1px 1px #0785ac;
}

thead { background: #edf9fa; border-color: #dceef0; }
th { color: #80a6ae; font-size: 11px; font-weight: 700; }
td { color: #71919a; font-size: 12px; border-color: #edf5f5; }
tbody tr:hover { background: #f4fcfc; }
.folder-icon {
  color: #d88910;
  background: linear-gradient(#ffe578 0%, #ffc648 60%, #ffad27 100%);
  border: 1px solid #efb43f;
  box-shadow: inset 0 2px 0 #fff9, inset 0 -3px 0 #e99e20, 0 3px 4px #e9a43b33;
}
.folder-icon svg { fill: #fff4bf80; stroke-width: 2; }
.tree-cell { gap: 7px; }
.tree-parent-row { background: #eaf7f7; }
.tree-parent-row .folder-name { font-weight: 700; }
.tree-toggle, .tree-toggle-placeholder { display: inline-grid; place-items: center; flex: none; width: 24px; height: 24px; }
.tree-toggle { padding: 0; color: #245c6b; background: transparent; border: 0; border-radius: 4px; cursor: pointer; }
.tree-toggle:hover, .tree-toggle:focus-visible { background: #ccebf0; }
.tree-toggle svg { width: 16px; height: 16px; }
.tree-actions { display: flex; flex-wrap: wrap; gap: 8px; padding: 10px 21px; border-top: 1px solid #d6e9eb; }
.tree-actions .button { min-height: 32px; padding: 6px 10px; font-size: 11px; }
.tree-actions .button svg { width: 14px; height: 14px; }
.tree-assignment { display: block; width: fit-content; max-width: 100%; margin-top: 4px; padding: 2px 6px; color: #164b5e; background: #d7f0f3; border: 1px solid #a7d6df; border-radius: 4px; font-size: 11px; overflow-wrap: anywhere; }
.folder-name { color: #466c77; font-size: 12px; }
.folder-parent { color: #a0b9bd; }
.file-count { color: #56808a; }
.owner-avatar { color: #168caf; background: linear-gradient(#e2f9ff, #bcecf6); box-shadow: inset 0 1px 0 #fff, 0 2px 0 #a6dce7; }
.owner-avatar.empty { color: #a8b8ba; background: #eef4f3; box-shadow: 0 2px 0 #dce8e7; }
.progress-info strong { color: #1aa9c9; }
.status { border-radius: 6px; font-weight: 700; box-shadow: inset 0 2px 0 #ffffff91, 0 2px 0 #00000012; }
.status.done { color: #459e25; background: linear-gradient(#e6fbd8, #c4eeaa); border-color: #b3e599; }
.status.progress { color: #ba7916; background: linear-gradient(#fff5d5, #ffe2a0); border-color: #f5d384; }
.status.pending { color: #879ca1; background: linear-gradient(#f4f8f8, #e4eeee); border-color: #d8e6e7; }
.table-footer { border-color: #e4f1f1; color: #9bb5b9; }
.page-footer { color: #a0b9bc; }

dialog { border: 2px solid #c4e6ea; box-shadow: 0 6px 0 #a7d7de, 0 20px 70px #315c6b55; }
dialog::backdrop { background: #335865a6; }
.confirm-dialog { width: min(440px, calc(100% - 32px)); }
.confirm-dialog .dialog-header { margin-bottom: 14px; }
.confirm-message { margin: 0 0 22px; color: #45646d; font-size: 13px; line-height: 1.8; white-space: pre-line; overflow-wrap: anywhere; }
.confirm-dialog .button { min-width: 84px; padding: 0 16px; }
.confirm-dialog .button.danger { background: linear-gradient(#ff8a6b, #ee5a3a 55%, #d9431f); border: 2px solid #b8341a; box-shadow: inset 0 3px 0 #ffffff66, inset 0 -4px 0 #b8341a, 0 4px 5px #b8341a50; color: white; }
#item-dialog { overflow: visible; }
.flatpickr-calendar { width: 292px; padding: 8px 10px 10px; border: 2px solid #c4e6ea; border-radius: 12px; background: #fff; color: #183b48; font-family: inherit; box-shadow: 0 5px 0 #a7d7de, 0 16px 40px #315c6b40; }
.flatpickr-calendar::before, .flatpickr-calendar::after { display: none; }
.flatpickr-months { align-items: center; padding-bottom: 4px; }
.flatpickr-months .flatpickr-month { height: 36px; color: #183b48; fill: #183b48; }
.flatpickr-current-month { display: flex; align-items: center; justify-content: center; gap: 4px; padding: 0; font-size: 14px; font-weight: 700; line-height: 1; }
.flatpickr-current-month .cur-month, .flatpickr-current-month input.cur-year { color: #183b48; font-weight: 700; }
.flatpickr-months .flatpickr-prev-month, .flatpickr-months .flatpickr-next-month { position: static; display: grid; width: 30px; height: 30px; padding: 0; place-items: center; border-radius: 8px; color: #67b4c6; fill: #67b4c6; }
.flatpickr-months .flatpickr-prev-month:hover, .flatpickr-months .flatpickr-next-month:hover { background: #e4f4f6; }
.flatpickr-months .flatpickr-prev-month:hover svg, .flatpickr-months .flatpickr-next-month:hover svg { fill: #189bc1; }
.flatpickr-months .flatpickr-prev-month svg, .flatpickr-months .flatpickr-next-month svg { width: 12px; height: 12px; }
span.flatpickr-weekday { color: #67b4c6; font-size: 12px; font-weight: 700; }
.flatpickr-innerContainer, .flatpickr-rContainer, .flatpickr-days, .dayContainer { width: 272px; min-width: 272px; max-width: 272px; }
.dayContainer { padding: 2px 0 0; }
.flatpickr-day { max-width: 36px; height: 36px; margin: 1px 0; border-radius: 8px; color: #183b48; font-weight: 600; line-height: 34px; }
.flatpickr-day:hover, .flatpickr-day:focus { border-color: #d9eeee; background: #e4f4f6; }
.flatpickr-day.prevMonthDay, .flatpickr-day.nextMonthDay, .flatpickr-day.flatpickr-disabled { color: #b9cdd1; }
.flatpickr-day.today { border-color: #67b4c6; }
.flatpickr-day.today:hover { border-color: #67b4c6; background: #e4f4f6; color: #183b48; }
.flatpickr-day.selected, .flatpickr-day.selected:hover { border-color: #087fa5; background: #189bc1; color: #fff; box-shadow: 0 2px 0 #087fa5; }
.fp-today { display: block; width: 100%; height: 32px; margin-top: 8px; border: 1px solid #d9eeee; border-radius: 8px; background: linear-gradient(#fff, #eef8f9); color: #2a7f93; font: inherit; font-size: 12px; font-weight: 700; cursor: pointer; }
.fp-today:hover { border-color: #a7d7de; background: #e4f4f6; }
.dialog-header h2 { color: #375d69; }
.form-grid label { color: #6d9199; }
.form-grid input { background: #f5fbfb; border: 2px solid #d9eeee; box-shadow: inset 0 2px 3px #79abb31a; }
.toast { background: #189bc1; box-shadow: inset 0 2px 0 #ffffff66, 0 4px 0 #087fa5, 0 6px 12px #158bac33; }

@media (max-width: 1050px) {
  .overview { padding: 7px 0; }
  .overview-intro, .metric:nth-child(2) { border-bottom-color: #e5f1f0; }
}
@media (max-width: 700px) {
  .sidebar { border-bottom: 2px solid #cde9e8; }
  .table-toolbar { padding-top: 13px; }
  .button { font-size: 11px; }
}

:root { color: #183b48; --ink: #183b48; --muted: #45646d; }
button:focus-visible, input:focus-visible, select:focus-visible { outline-color: #163f78; }

.brand strong, .page-heading h1, .overview-title h2, .list-heading h2,
.metric-value strong, .dialog-header h2 { color: #173d4a; }
.brand span, .sidebar-label, .nav-caption, .sidebar-bottom span,
.breadcrumb, .breadcrumb strong, .today, .eyebrow, .section-kicker,
.section-kicker span, .page-heading p, .list-heading p, .overview-title p,
.metric-label, .metric-value span, .metric-foot, .table-footer, .page-footer,
.page-footer span span, .form-grid label { color: #45646d; }
.brand span, .sidebar-label, .nav-caption, .sidebar-bottom span,
.eyebrow, .section-kicker, .section-kicker span, .metric-foot,
.table-footer, .page-footer { font-size: 11px; }
.sidebar-bottom strong, .sidebar-bottom span, .brand span { line-height: 1.4; }
.nav-item, #folder-nav .nav-item { color: #315965; font-size: 12px; }
.nav-item:hover { color: #163e4d; }
.nav-item.active, .nav-item.active:hover {
  color: #092f40;
  background: linear-gradient(#89e8ff, #4bc9ef 60%, #20aeda);
  text-shadow: none;
}
.sidebar-bottom .online-dot { background: #23831b; }

.primary, .primary:hover {
  color: #153d0a;
  background: linear-gradient(#a0f35b, #70d82e 48%, #4fbe12);
  text-shadow: none;
}
.secondary, .secondary:hover { color: #234f5f; border-color: #8dbfc9; }
.metric-icon.blue { color: #07394c; }
.metric-icon.green { color: #173e0b; }
.metric-icon.amber { color: #553200; }
.mini-chart, .storage-icon { color: #155469; }
.list-heading h2 span { color: #4b2c00; }

.search-box { color: #426d79; border-color: #a7cdd3; }
.search-box input { color: #183b48; }
.search-box input::placeholder { color: #55757d; opacity: 1; }
.search-box kbd { color: #45646d; border-color: #aecbd0; }
.filter-select, .filter-select select, .icon-button, .subtle-icon,
.owner-filters-label { color: #315e6b; }
.table-toolbar, .toolbar-actions { flex-wrap: wrap; }
.upload-date-filter { display: flex; align-items: center; gap: 7px; color: #315e6b; font-size: 12px; font-weight: 600; white-space: nowrap; }
.upload-date-filter input { min-width: 0; padding: 5px 7px; border: 1px solid #a7cdd3; border-radius: 5px; background: #f5fbfb; color: #183b48; font: inherit; }
.clear-upload-date { padding: 4px; border: 0; background: none; color: #14586f; font-size: 12px; cursor: pointer; }
.clear-upload-date[hidden] { display: none; }

@media (max-width: 1050px) {
  .table-toolbar .search-box { width: 100%; max-width: none; }
  .toolbar-actions { width: 100%; justify-content: flex-start; }
}
.icon-button:hover, .subtle-icon:hover { color: #123f50; }
.owner-tag { color: #274e5b; border-color: #9cc9d0; }
.owner-tag:hover { color: #163e4d; border-color: #348ba2; }
.owner-tag[aria-pressed="true"] {
  color: #092f40;
  background: linear-gradient(#89e8ff, #4bc9ef 60%, #20aeda);
  border-color: #087da5;
  text-shadow: none;
}

thead { background: #d7eff2; border-color: #a8cdd3; }
th { color: #294f5b; font-size: 12px; }
td { color: #294d58; border-color: #dcebed; }
.folder-name, .file-count { color: #183f4d; }
.folder-parent { color: #4b6972; font-size: 11px; }
.folder-icon { color: #694000; }
.owner-avatar { color: #164b5e; }
.owner-avatar.empty { color: #4b6469; }
.progress-info { color: #496771; font-size: 11px; }
.progress-info strong { color: #125a70; }
.status { font-size: 11px; }
.status.done { color: #205e13; border-color: #80bb65; }
.status.progress { color: #704300; border-color: #d5aa4a; }
.status.pending { color: #415b60; border-color: #aebfc2; }
.status { cursor: pointer; }
.folder-cell { min-width: 0; }
.folder-name { max-width: none; min-width: 7em; overflow: visible; white-space: normal; overflow-wrap: anywhere; }
.progress-cell { width: 120px; min-width: 110px; padding-right: 10px; }
.folder-status {
  --status-top: #c3ced2; --status-bottom: #aebbc1; --status-edge: #3f525b; --status-dark: #8b9ba3;
  display: inline-flex; align-items: center; justify-content: center; gap: 6px;
  width: 124px; min-height: 34px; padding: 4px 12px;
  border: 2px solid var(--status-edge); border-radius: 999px;
  background: linear-gradient(var(--status-top), var(--status-bottom));
  box-shadow: inset 0 2px 0 #ffffff59, inset 0 -4px 0 var(--status-dark);
  color: #fff; font: inherit; font-size: 12px; font-weight: 700; text-shadow: 0 1px 0 var(--status-edge), 0 0 2px var(--status-edge);
  cursor: pointer; transition: transform .1s, box-shadow .1s, filter .1s;
}
.folder-status svg { width: 14px; height: 14px; flex: none; filter: drop-shadow(0 1px 0 var(--status-edge)); }
.folder-status:hover:not(:disabled) { filter: brightness(1.06); }
.folder-status:active:not(:disabled) { transform: translateY(2px); box-shadow: inset 0 2px 0 #ffffff40, inset 0 -2px 0 var(--status-dark); }
.folder-status:disabled { opacity: .6; cursor: not-allowed; }
.folder-status.processing { --status-top: #ffb94d; --status-bottom: #f7a23a; --status-edge: #6e2a10; --status-dark: #d98324; }
.folder-status.ready { --status-top: #34aae8; --status-bottom: #2399dc; --status-edge: #0d2f7a; --status-dark: #1472bd; }
.folder-status.uploading { --status-top: #bf72e3; --status-bottom: #ab5ed2; --status-edge: #3f1766; --status-dark: #8a42b3; }
.folder-status.uploaded { --status-top: #86c94a; --status-bottom: #74bd3b; --status-edge: #16622a; --status-dark: #4ea234; }
.folder-status:focus-visible { outline: 3px solid #ffbc41; outline-offset: 2px; }
.empty-state, .empty-state strong, .empty-state span { color: #45646d; }
.form-grid input { color: #183b48; border-color: #a7cdd3; }
.form-grid select { display: block; width: 100%; height: 39px; margin-top: 7px; padding: 0 9px; color: #183b48; background: #f5fbfb; border: 2px solid #a7cdd3; border-radius: 5px; font-size: 12px; }
.form-grid [hidden] { display: none; }
.form-grid input::placeholder { color: #55757d; opacity: 1; }
.toast { color: #fff; background: #14586f; }
.data-actions { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 8px; }
.data-actions [hidden] { display: none; }
.table-wrap { position: relative; }
@media (max-width: 700px) { .list-heading { align-items: flex-start; flex-direction: column; gap: 12px; } .data-actions { justify-content: flex-start; } }

````

### 5.11 `static\css\planner.css`

<!-- REBUILD_FILE {"path":"static\\css\\planner.css","newline":"LF","encoding":"utf-8","size":13666,"sha256":"e68d44c7a304e3c836640dc7c0f2da99cac082291833fe0290ed4887df5efe05"} -->
````css
.planner { margin-bottom: 39px; padding: 22px 25px; background: #fff; border: 2px solid #d5ebec; border-radius: 8px; box-shadow: 0 5px 0 #c6e3e3, 0 8px 15px #70aeb51c; }
.planner-heading, .planner-result-heading { display: flex; justify-content: space-between; align-items: center; gap: 14px; }
.planner-heading h2 { margin: 8px 0 0; color: #173d4a; font-size: 18px; }
.planner-badge { display: inline-flex; align-items: center; gap: 6px; color: #155469; background: #e5f8fa; border: 1px solid #add8df; border-radius: 6px; padding: 7px 10px; font-size: 11px; font-weight: 700; white-space: nowrap; }
.planner-badge svg { width: 15px; height: 15px; }
.planner-form { display: grid; grid-template-columns: minmax(140px,1fr) minmax(100px,.7fr) minmax(150px,1.3fr) minmax(110px,1fr) auto; align-items: end; gap: 12px; margin-top: 21px; }
.planner-form label, .planner-field { min-width: 0; color: #294f5b; font-size: 11px; font-weight: 700; }
.planner-form input, .planner-form select { display: block; width: 100%; height: 39px; margin-top: 7px; padding: 0 9px; color: #183b48; background: #f5fbfb; border: 2px solid #a7cdd3; border-radius: 5px; font-size: 12px; }
.planner-form input::placeholder { color: #55757d; opacity: 1; }
.person-picker { position: relative; margin-top: 7px; }
.person-picker summary { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-height: 39px; padding: 5px 11px; color: #183b48; background: #f5fbfb; border: 2px solid #a7cdd3; border-radius: 5px; cursor: pointer; list-style: none; font-size: 12px; transition: border-color .15s, background-color .15s; }
.person-picker summary::-webkit-details-marker { display: none; }
.person-picker summary:focus-visible { outline: 2px solid #087da5; outline-offset: 2px; }
.person-picker summary > span { display: flex; flex: 1 1 auto; flex-wrap: wrap; align-items: center; gap: 5px; min-width: 0; }
.person-picker-summary-chip, .person-picker-chip { --picker-text: hsl(var(--person-hue, 190), 65%, 25%); --picker-border: hsl(var(--person-hue, 190), 43%, 70%); --picker-background: hsl(var(--person-hue, 190), 65%, 96%); --picker-hover: hsl(var(--person-hue, 190), 60%, 90%); }
.person-picker-summary-chip { box-sizing: border-box; max-width: 100%; padding: 3px 8px; color: #fff; background: var(--picker-text); border: 1px solid var(--picker-text); border-radius: 999px; font-size: 11px; line-height: 1.2; overflow-wrap: anywhere; }
.person-picker-unassigned { --picker-text: #526461; --picker-border: #b6c8c3; --picker-background: #f0f5f3; --picker-hover: #e2eeea; }
.person-picker summary svg { width: 15px; height: 15px; flex: none; color: #46808c; transition: transform .15s; }
.person-picker summary:hover { border-color: #68aab7; background: #fff; }
.person-picker[open] summary { border-color: #1687a3; background: #fff; }
.person-picker[open] summary svg { transform: rotate(180deg); }
.person-picker-panel { position: absolute; z-index: 20; top: calc(100% + 7px); left: 0; box-sizing: border-box; width: max(100%, 240px); max-width: min(360px, calc(100vw - 32px)); padding: 8px; background: #fff; border: 1px solid #a7cdd3; border-radius: 6px; box-shadow: 0 4px 0 #c6e3e3, 0 12px 28px #25535b33; }
.person-picker-panel.person-picker-above { top: auto; bottom: calc(100% + 7px); }
.person-picker-heading { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 4px 8px 9px; color: #244e59; font-size: 12px; font-weight: 700; }
.person-picker-heading strong { color: #568078; font-size: 11px; font-weight: 600; white-space: nowrap; }
.person-picker-list { display: flex; flex-wrap: wrap; align-content: flex-start; gap: 9px; max-height: min(250px, 42vh); padding: 6px; overflow-y: auto; overscroll-behavior: contain; scrollbar-width: thin; scrollbar-color: #a7cdd3 transparent; }
.planner-form .person-picker-option { position: relative; display: inline-flex; max-width: 100%; padding: 0; cursor: pointer; }
.planner-form .person-picker-option input { position: absolute; inset: 0; width: 100%; height: 100%; margin: 0; padding: 0; opacity: 0; cursor: pointer; scroll-margin: 6px; }
.person-picker-chip { display: inline-flex; align-items: center; gap: 7px; box-sizing: border-box; max-width: 100%; min-height: 32px; padding: 4px 11px 4px 5px; color: var(--picker-text); background: var(--picker-background); border: 1px solid var(--picker-border); border-radius: 999px; font-size: 12px; line-height: 1.3; }
.person-picker-chip-name { min-width: 0; overflow-wrap: anywhere; }
.person-picker-chip-icon { display: grid; place-items: center; width: 21px; height: 21px; flex: none; color: var(--picker-text); background: #fff; border: 1px solid var(--picker-border); border-radius: 50%; font-size: 15px; line-height: 1; }
.person-picker-chip-icon::before { content: '+'; }
.person-picker-option:hover .person-picker-chip { border-color: var(--picker-text); background: var(--picker-hover); }
.person-picker-option input:checked + .person-picker-chip { color: #fff; background: var(--picker-text); border-color: var(--picker-text); }
.person-picker-option input:checked + .person-picker-chip .person-picker-chip-icon { color: var(--picker-text); border-color: #fff; }
.person-picker-option input:checked + .person-picker-chip .person-picker-chip-icon::before { content: '\2713'; font-size: 13px; font-weight: 700; }
.person-picker-option:focus-within .person-picker-chip { outline: 2px solid #087da5; outline-offset: 2px; }
.person-picker-empty { display: block; padding: 12px 10px; color: #667d80; }
.person-picker-add { display: flex; gap: 7px; margin-top: 7px; padding: 9px 2px 2px; border-top: 1px solid #dfefed; }
.planner-form .person-picker-add input { min-width: 0; height: 36px; margin: 0; background: #f7fbfa; border-width: 1px; font-weight: 400; }
.planner-form .person-picker-add input:focus { outline: 2px solid #087da5; outline-offset: 1px; }
.person-picker-add button, .person-picker-clear { flex: none; padding: 0 13px; color: #fff; background: #137e91; border: 1px solid #0e6f80; border-radius: 4px; cursor: pointer; font-size: 11px; font-weight: 700; }
.person-picker-clear { width: 100%; height: 34px; margin-top: 8px; color: #315d64; background: #f2f8f7; border-color: #d0e4e1; }
.person-picker-add button:hover { background: #0e6a7c; }
.person-picker-clear:hover { background: #e3f2ef; }
.planner-message, .planner-note { margin: 13px 0 0; color: #45646d; font-size: 11px; line-height: 1.6; }
.planner-message.error { color: #a42f26; font-weight: 700; }
.allocation-history { margin-top: 20px; border-top: 1px solid #d6e9eb; }
.allocation-history > summary { display: flex; align-items: center; gap: 8px; padding: 15px 0 0; color: #245665; font-size: 13px; font-weight: 700; cursor: pointer; list-style: none; }
.allocation-history > summary::-webkit-details-marker { display: none; }
.allocation-history > summary svg { width: 16px; height: 16px; }
.allocation-history > summary svg:last-child { margin-left: auto; }
.allocation-history[open] > summary svg:last-child { transform: rotate(180deg); }
.history-toolbar { display: flex; align-items: center; justify-content: space-between; margin: 14px 0 6px; color: #45646d; font-size: 11px; }
.history-list { border-top: 1px solid #dcebea; }
.history-entry { border-bottom: 1px solid #dcebea; }
.history-entry > summary { display: grid; grid-template-columns: minmax(130px, 1fr) minmax(100px, .7fr) minmax(170px, 1.2fr) minmax(150px, 1fr) auto; align-items: center; gap: 12px; min-height: 46px; color: #31535c; font-size: 12px; cursor: pointer; list-style: none; }
.history-entry > summary::-webkit-details-marker { display: none; }
.history-entry > summary strong { color: #173d4a; }
.history-entry > summary svg { width: 15px; height: 15px; }
.history-entry[open] > summary svg { transform: rotate(180deg); }
.history-content { padding: 4px 0 15px; }
.history-actions { display: flex; justify-content: flex-end; margin-bottom: 8px; }
.history-download { display: inline-flex; align-items: center; gap: 6px; padding: 7px 10px; color: #14566a; border: 1px solid #a7d6df; border-radius: 4px; background: #eaf8fa; text-decoration: none; font-size: 11px; font-weight: 700; }
.history-download svg { width: 14px; height: 14px; }
.history-person { padding: 9px 0; border-top: 1px solid #edf4f4; color: #31535c; font-size: 12px; }
.history-person strong { color: #173d4a; }
.history-person ul { margin: 6px 0 0; padding-left: 20px; }
.history-person li { padding: 3px 0; overflow-wrap: anywhere; }
.history-empty { padding: 15px 0; color: #55757d; font-size: 12px; }
#planner-results { margin-top: 18px; border-top: 1px solid #d6e9eb; padding-top: 17px; }
.planner-result-heading { position: relative; }
.plan-options { display: flex; flex-wrap: wrap; gap: 8px; }
.plan-option { padding: 8px 12px; border: 2px solid #a4d3db; border-radius: 6px; color: #214f5c; background: linear-gradient(#fff, #e5f7f9); font-size: 11px; font-weight: 700; box-shadow: inset 0 2px #fff, 0 2px #c4e3e6; }
.plan-option[aria-pressed="true"] { color: #092f40; border-color: #087da5; background: linear-gradient(#89e8ff, #4bc9ef 60%, #20aeda); box-shadow: inset 0 2px #fff8, 0 2px #087fac; }
#plan-preview-spacer { overflow-anchor: none; }
.plan-hover-preview { position: absolute; z-index: 15; top: 0; left: 0; box-sizing: border-box; width: min(720px, 100%); padding: 16px 18px; color: #28525d; background: #fff; border: 1px solid #a7cdd3; border-radius: 6px; box-shadow: 0 8px 24px #25535b33; transform: translateY(-100%); font-size: 13px; line-height: 1.4; }
.plan-hover-preview[hidden] { display: none; }
.plan-hover-title, .plan-hover-person-heading, .plan-hover-folder-heading { display: flex; align-items: baseline; justify-content: space-between; gap: 10px; }
.plan-hover-title { color: #173d4a; font-size: 16px; }
.plan-hover-title span { flex: none; color: #087da5; font-size: 13px; font-weight: 700; }
.plan-hover-overview { margin-top: 4px; color: #55757d; font-size: 12px; }
.plan-hover-people { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px 20px; margin-top: 12px; }
.plan-hover-person { min-width: 0; padding-top: 9px; border-top: 2px solid hsl(var(--person-hue), 43%, 65%); }
.plan-hover-person-heading strong, .plan-hover-folder-heading strong, .plan-hover-folder-details { min-width: 0; overflow-wrap: anywhere; }
.plan-hover-person-heading strong { font-size: 14px; }
.plan-person-marker { --person-color: hsl(var(--person-hue), 65%, 23%); --person-border: hsl(var(--person-hue), 43%, 65%); --person-background: hsl(var(--person-hue), 65%, 94%); display: inline-flex; align-items: baseline; gap: 5px; max-width: 100%; padding: 3px 8px; color: var(--person-color); background: var(--person-background); border: 1px solid var(--person-border); border-radius: 4px; font-weight: 700; }
.plan-person-marker > span:last-child { min-width: 0; overflow-wrap: anywhere; }
.plan-person-marker-number { flex: none; padding-right: 5px; border-right: 1px solid currentColor; font-size: 10px; line-height: 1.2; }
.plan-person-unassigned { --person-color: #536365; --person-border: #bbc8c7; --person-background: #f1f5f4; }
.plan-hover-person-heading > span, .plan-hover-folder-heading > span { flex: none; color: #36545b; font-size: 12px; font-weight: 600; white-space: nowrap; }
.plan-hover-folders { margin: 7px 0 0; padding: 0; list-style: none; }
.plan-hover-folder { padding: 6px 0; border-top: 1px solid #edf4f4; }
.plan-hover-folder-heading strong { color: #173d4a; font-size: 13px; font-weight: 700; }
.plan-hover-folder-heading span { padding: 2px 6px; color: #315f85; background: #eaf3fb; border-radius: 4px; }
.plan-hover-folder-details { display: flex; flex-wrap: wrap; align-items: baseline; gap: 2px 3px; margin-top: 2px; color: #5b6d69; font-size: 11px; }
.plan-hover-folder-details .plan-person-marker { padding: 0 4px; font-size: 11px; }
.plan-summary { display: flex; flex-wrap: wrap; gap: 8px; margin: 15px 0 12px; }
.plan-summary span { padding: 6px 9px; color: #28525d; background: #eef9fa; border: 1px solid #cae5e8; border-radius: 5px; font-size: 11px; }
.plan-summary strong { color: #173d4a; }
.plan-preview-bar { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 12px; padding: 10px 12px; color: #164b5e; background: #e5f7f9; border: 1px solid #a7d6df; border-radius: 5px; font-size: 12px; }
.plan-preview-bar[hidden], .table-toolbar[hidden], .owner-filters[hidden] { display: none; }
.plan-preview-bar .button { flex: none; }
.date-history-bar { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 10px; margin-bottom: 12px; padding: 10px 12px; color: #164b5e; background: #f3fafb; border: 1px solid #cfe7ea; border-radius: 5px; font-size: 12px; }
.date-history-bar[hidden] { display: none; }
.date-history-title { display: inline-flex; align-items: center; gap: 6px; font-weight: 700; }
.date-history-title svg { width: 14px; height: 14px; }
.date-history-empty { color: #6f8e96; }
button.history-download { font: inherit; cursor: pointer; }
@media (max-width: 1100px) { .planner-form { grid-template-columns: repeat(3, minmax(0, 1fr)); } }
@media (max-width: 700px) { .planner { padding: 18px 14px; margin-bottom: 30px; } .planner-form { grid-template-columns: repeat(2, minmax(0, 1fr)); } .planner-people { grid-column: span 2; } .planner-result-heading { align-items: flex-start; flex-direction: column; } .planner-badge { font-size: 10px; } .plan-hover-preview { padding: 12px; } .plan-hover-people { grid-template-columns: minmax(0, 1fr); gap: 8px; margin-top: 8px; } }
@media (max-width: 380px) { .planner-heading { align-items: flex-start; flex-direction: column; } }

````

### 5.12 `static\css\cartoon-style.css`

<!-- REBUILD_FILE {"path":"static\\css\\cartoon-style.css","newline":"CRLF","encoding":"utf-8","size":16634,"sha256":"2c5405b03ee17f9d625142b49560b94351a89da7e5fecf07637f90c70993e6eb"} -->
````css
/* Cartoon mobile-game theme: royal-blue backdrop, thick navy outlines, gold/orange/purple panels. Colors only; layout lives in the other stylesheets. */
:root {
  color: #1b2a55;
  background: #2b84df;
  --ink: #1b2a55;
  --muted: #5a6f9a;
  --line: #b8c8e8;
  --outline: #1b2a55;
  --shadow: #0e1a3d;
  --gold-edge: #7a4306;
}

body {
  background-color: #2b84df;
  background-image:
    radial-gradient(circle at 16% 20%, #ffffff14 0 90px, transparent 91px),
    radial-gradient(circle at 84% 36%, #ffffff12 0 130px, transparent 131px),
    radial-gradient(circle at 28% 86%, #ffffff12 0 110px, transparent 111px),
    radial-gradient(circle, #ffffff10 2px, transparent 3px);
  background-size: auto, auto, auto, 28px 28px;
  background-attachment: fixed;
}

button:focus-visible, input:focus-visible, select:focus-visible { outline: 3px solid #fff; outline-offset: 2px; box-shadow: 0 0 0 6px var(--outline); }

/* Sidebar / topbar */
.sidebar {
  background: linear-gradient(180deg, #2a6fd0, #1b4fa6);
  border-right: 3px solid var(--outline);
  box-shadow: 4px 0 0 #0e1a3d40;
}
.brand-mark {
  color: #7b2cbf;
  background: linear-gradient(#ffe65f, #ffb81c);
  border: 3px solid var(--gold-edge);
  border-radius: 12px;
  box-shadow: inset 0 3px 0 #fff9, 0 3px 0 var(--shadow);
}
.brand strong { color: #fff; font-size: 16px; text-shadow: 0 2px 0 var(--outline), 0 0 3px var(--outline); }
.brand span, .sidebar-label, .nav-caption, .sidebar-bottom span { color: #c6defc; }
.nav-item, #folder-nav .nav-item { color: #eaf3ff; font-weight: 700; border-radius: 12px; }
.nav-item:hover { background: #ffffff26; color: #fff; }
.nav-item.active, .nav-item.active:hover {
  color: #fff;
  background: linear-gradient(#ffb04a, #f0761e);
  border: 3px solid var(--outline);
  box-shadow: inset 0 3px 0 #ffffff66, 0 3px 0 var(--shadow);
  text-shadow: 0 2px 0 #7a3200;
}
.nav-caption .subtle-icon { color: #fff; }
.nav-caption .subtle-icon:hover { background: #ffffff33; color: #fff; }
.sidebar-bottom { border-color: #ffffff40; }
.sidebar-bottom strong { color: #fff; }
.storage-icon { color: #7b2cbf; background: linear-gradient(#ffe65f, #ffb81c); border: 2px solid var(--gold-edge); border-radius: 8px; box-shadow: inset 0 2px 0 #fff9; }
.sidebar-bottom .online-dot { background: #5ed827; border: 2px solid var(--outline); width: 12px; height: 12px; box-shadow: none; }

.topbar {
  background: linear-gradient(#2a5db5, #1c3f8c);
  border-bottom: 3px solid var(--outline);
}
.breadcrumb, .today { color: #c6defc; }
.breadcrumb strong { color: #fff; }
.avatar {
  color: #5a2d00;
  background: linear-gradient(#ffe65f, #ffb81c);
  border: 3px solid var(--outline);
  box-shadow: 0 2px 0 var(--shadow);
}

/* Page heading (sits on the blue backdrop) */
.eyebrow { color: #ffe45c; font-weight: 700; }
.eyebrow span { background: #5ed827; border: 2px solid var(--outline); box-shadow: none; width: 10px; height: 10px; }
.section-kicker { color: #ffe45c; font-weight: 700; }
.section-kicker span { color: #c6defc; }
.page-heading h1, .overview-title h2, .planner-heading h2, .list-heading h2 {
  color: #fff;
  font-weight: 700;
  -webkit-text-stroke: 5px var(--outline);
  paint-order: stroke fill;
  filter: drop-shadow(0 3px 0 var(--outline));
}
.page-heading h1 { font-size: 28px; }
.page-heading p, .list-heading p { color: #e6f1ff; font-size: 12px; font-weight: 500; }
.list-heading h2 span {
  color: #fff;
  background: linear-gradient(#ff6a62, #e5383b);
  border: 2px solid var(--outline);
  border-radius: 8px;
  box-shadow: inset 0 2px 0 #ffffff66;
  -webkit-text-stroke: 0;
  font-weight: 700;
  min-width: 26px;
  height: 22px;
  font-size: 12px;
}
.heading-spark, .footer-star { color: #ffe45c !important; }

/* Buttons */
.button {
  height: 42px;
  border-radius: 12px;
  font-size: 13px;
  font-weight: 700;
}
.primary, .primary:hover {
  color: #5a2d00;
  background: linear-gradient(#ffe65f, #ffc226 55%, #ffa800);
  border: 3px solid var(--gold-edge);
  box-shadow: inset 0 3px 0 #fff9, inset 0 -4px 0 #e08a00, 0 4px 0 #3b2203;
  text-shadow: 0 1px 0 #fff6;
}
.primary:hover { filter: brightness(1.06); }
.secondary, .secondary:hover {
  color: var(--outline);
  background: linear-gradient(#fff, #cfe3ff);
  border: 3px solid var(--outline);
  box-shadow: inset 0 3px 0 #fff, inset 0 -3px 0 #a9c6f2, 0 4px 0 var(--shadow);
}
.secondary:hover { filter: brightness(1.04); }
.button:active { transform: translateY(3px); }
.primary:active { box-shadow: inset 0 3px 0 #fff9, inset 0 -2px 0 #e08a00, 0 1px 0 #3b2203; }
.secondary:active { box-shadow: inset 0 3px 0 #fff, inset 0 -2px 0 #a9c6f2, 0 1px 0 var(--shadow); }
.confirm-dialog .button.danger { background: linear-gradient(#ff7a6b, #e5383b 55%, #c4262a); border: 3px solid #5a0e10; box-shadow: inset 0 3px 0 #ffffff66, inset 0 -4px 0 #a01a1d, 0 4px 0 #3b0709; }
.tree-actions .button { height: 36px; }

/* Overview: purple halftone card */
.overview {
  background: radial-gradient(circle, #ffffff1f 2px, transparent 2.5px) 0 0 / 14px 14px, linear-gradient(180deg, #9b45dd, #6d24b8);
  border: 3px solid var(--outline);
  border-radius: 16px;
  box-shadow: 0 6px 0 var(--shadow), inset 0 3px 0 #ffffff40;
}
.overview .section-kicker { color: #ffe45c; }
.overview .section-kicker span, .overview-title p, .metric-value span, .metric-foot { color: #e6d3ff; }
.metric { border-color: #ffffff40; border-left-style: dashed; border-left-width: 2px; }
.metric-label { color: #fff; font-weight: 700; }
.metric-value strong { color: #fff; text-shadow: 0 3px 0 #3d0f6e; }
.green-dot { background: #5ed827; box-shadow: 0 0 0 2px var(--outline); }
.mini-chart, .metric-icon, .metric-icon.blue, .metric-icon.green, .metric-icon.amber {
  background: linear-gradient(#ffe65f, #ffb81c);
  border: 2px solid var(--gold-edge);
  border-radius: 8px;
  box-shadow: inset 0 2px 0 #fff9, 0 2px 0 #3b2203;
}
.mini-chart { color: #7b2cbf; }
.metric-icon.blue { color: #2a7fe0; }
.metric-icon.green { color: #2aa31a; }
.metric-icon.amber { color: #c26a00; }
.progress-line, .row-progress {
  height: 14px;
  background: #1d2d5c;
  border: 2px solid var(--outline);
  border-radius: 999px;
  box-shadow: inset 0 2px 3px #0008;
}
.row-progress { height: 12px; }
.progress-line span, .row-progress span {
  background: linear-gradient(#b4f56a, #5ed827 55%, #36b01a);
  box-shadow: inset 0 2px 0 #ffffff80;
}

/* Planner: orange halftone card */
.planner {
  background: radial-gradient(circle, #ffffff26 2px, transparent 2.5px) 0 0 / 14px 14px, linear-gradient(180deg, #ffa43c, #f0761e);
  border: 3px solid var(--outline);
  border-radius: 16px;
  box-shadow: 0 6px 0 var(--shadow), inset 0 3px 0 #ffffff55;
}
.planner .section-kicker { color: #fff; text-shadow: 0 1px 0 #7a3200; }
.planner .section-kicker span { color: #ffe6c2; }
.planner-badge { color: #fff; background: #2b2a7a; border: 2px solid var(--outline); border-radius: 8px; }
.planner-form label, .planner-field { color: #4a1d00; font-size: 12px; }
.planner-form input, .planner-form select, .form-grid input, .form-grid select, .upload-date-filter input {
  color: var(--outline);
  background: #fff;
  border: 2px solid var(--outline);
  border-radius: 10px;
  box-shadow: inset 0 2px 3px #1b2a5526;
}
.planner-form input::placeholder, .form-grid input::placeholder, .search-box input::placeholder { color: #7a8aa8; }
.person-picker summary { color: var(--outline); background: #fff; border: 2px solid var(--outline); border-radius: 10px; box-shadow: inset 0 2px 3px #1b2a5526; }
.person-picker summary svg { color: var(--outline); }
.person-picker summary:hover { border-color: var(--outline); background: #fff3c9; }
.person-picker[open] summary { border-color: var(--outline); background: #fff; }
.person-picker-panel { background: #fff; border: 3px solid var(--outline); border-radius: 12px; box-shadow: 0 5px 0 var(--shadow), 0 14px 30px #0006; }
.person-picker-heading { color: var(--outline); }
.person-picker-heading strong { color: var(--muted); }
.person-picker-list { scrollbar-color: #8fa8d8 transparent; }
.person-picker-add { border-top: 2px dashed #c9d8f2; }
.planner-form .person-picker-add input { border-width: 2px; }
.person-picker-add button { color: #fff; background: linear-gradient(#4aa0f7, #2370d0); border: 2px solid var(--outline); border-radius: 8px; }
.person-picker-add button:hover { background: linear-gradient(#5aaaff, #2a7fe0); }
.person-picker-clear { color: var(--outline); background: linear-gradient(#fff, #cfe3ff); border: 2px solid var(--outline); border-radius: 8px; }
.person-picker-clear:hover { background: #dbeaff; }
.planner-message, .planner-note { color: #4a1d00; font-size: 12px; font-weight: 600; }
.planner-message.error { display: inline-block; padding: 6px 12px; color: #fff; background: #d12b2f; border: 2px solid var(--outline); border-radius: 8px; }

.allocation-history, #planner-results {
  background: #fff7e6;
  border: 3px solid var(--outline);
  border-radius: 12px;
  box-shadow: 0 4px 0 var(--shadow);
}
.allocation-history { padding: 0 16px 14px; }
#planner-results { padding: 16px; }
.allocation-history > summary, .history-entry > summary strong, .history-person strong, .plan-hover-title, .plan-hover-folder-heading strong, .plan-summary strong { color: var(--outline); }
.history-toolbar, .history-entry > summary, .history-person, .history-empty { color: #3a4f7a; }
.history-list, .history-entry { border-color: #ecdcb8; }
.history-person { border-top-color: #f0e4c8; }
.history-download { color: var(--outline); background: #dbeaff; border: 2px solid var(--outline); border-radius: 8px; }
.plan-option { color: var(--outline); border: 2px solid var(--outline); border-radius: 10px; background: linear-gradient(#fff, #cfe3ff); box-shadow: inset 0 2px #fff, 0 3px 0 var(--shadow); }
.plan-option[aria-pressed="true"] { color: #5a2d00; border-color: var(--gold-edge); background: linear-gradient(#ffe65f, #ffb81c); box-shadow: inset 0 2px #fff9, 0 3px 0 #3b2203; }
.plan-hover-preview { color: #3a4f7a; background: #fff; border: 3px solid var(--outline); border-radius: 12px; box-shadow: 0 5px 0 var(--shadow), 0 14px 30px #0006; }
.plan-hover-title span { color: #2a7fe0; }
.plan-hover-overview, .plan-hover-person-heading > span, .plan-hover-folder-details { color: var(--muted); }
.plan-hover-folder { border-top-color: #e3ecfa; }
.plan-hover-folder-heading span { color: #2a5db5; background: #dbeaff; }
.plan-summary span { color: #3a4f7a; background: #fff; border: 2px solid #b8c8e8; border-radius: 8px; }
.plan-preview-bar, .date-history-bar { color: var(--outline); background: #fff3c9; border: 2px solid var(--outline); border-radius: 10px; box-shadow: 0 3px 0 var(--shadow); }
.date-history-empty { color: var(--muted); }

/* Table card */
.table-card {
  background: #f3f8ff;
  border: 3px solid var(--outline);
  border-radius: 16px;
  box-shadow: 0 6px 0 var(--shadow);
}
.table-toolbar, .owner-filters, .tree-actions { background: #e3efff; }
.tree-actions { border-top: 2px dashed #b8cdf0; }
.search-box { color: var(--outline); background: #fff; border: 2px solid var(--outline); border-radius: 12px; box-shadow: inset 0 2px 3px #1b2a5526; }
.search-box input { color: var(--outline); }
.search-box kbd { color: var(--muted); border-color: #b8c8e8; }
.filter-select, .filter-select select, .icon-button, .subtle-icon, .owner-filters-label, .upload-date-filter { color: var(--outline); }
.toolbar-divider { background: #b8cdf0; }
.icon-button:hover, .subtle-icon:hover { color: var(--outline); background: #cfe2ff; }
.clear-upload-date { color: #c4262a; font-weight: 700; }
.owner-tag {
  color: var(--outline);
  background: linear-gradient(#fff, #e3efff);
  border: 2px solid var(--outline);
  border-radius: 999px;
  box-shadow: 0 2px 0 var(--shadow);
  font-weight: 700;
  white-space: nowrap;
}
.owner-tag:hover { color: var(--outline); background: #fff3c9; border-color: var(--outline); }
.owner-tag[aria-pressed="true"] {
  color: #5a2d00;
  background: linear-gradient(#ffe65f, #ffb81c);
  border-color: var(--outline);
  box-shadow: inset 0 2px 0 #fff9, 0 2px 0 var(--shadow);
  text-shadow: none;
}

thead { background: linear-gradient(#3a8ef0, #2370d0); border-top: 0; border-bottom: 3px solid var(--outline); }
th { color: #fff; font-size: 12px; text-shadow: 0 1px 0 var(--outline); }
td { color: var(--outline); border-color: #d3e2f8; }
tbody tr:hover { background: #fff3c9; }
.tree-parent-row, .tree-parent-row:hover { background: #dcebff; }
.folder-icon { color: #694000; border: 2px solid var(--gold-edge); border-radius: 8px; box-shadow: inset 0 2px 0 #fff9, 0 2px 0 #3b2203; }
.tree-toggle { color: var(--outline); }
.tree-toggle:hover, .tree-toggle:focus-visible { background: #cfe2ff; }
.tree-assignment { color: var(--outline); background: #fff3c9; border-color: #c9a24a; }
.folder-name, .file-count { color: var(--outline); font-weight: 700; }
.folder-parent { color: var(--muted); }
.owner-avatar { color: #fff; background: linear-gradient(#c58bf0, #8a35cc); border: 2px solid var(--outline); box-shadow: inset 0 1px 0 #fff8; }
.owner-avatar.empty { color: var(--muted); background: #dfe6f3; border-color: #9aa9c8; box-shadow: none; }
.progress-info { color: var(--muted); }
.progress-info strong { color: var(--outline); }
.status.done { color: #205e13; background: linear-gradient(#e6fbd8, #b8ee94); border-color: #4ea234; }
.status.progress { color: #704300; background: linear-gradient(#fff5d5, #ffd877); border-color: #d5aa4a; }
.status.pending { color: #3a4f7a; background: linear-gradient(#f4f8ff, #d3e2f8); border-color: #9aa9c8; }
.folder-status { border-width: 3px; }
.table-footer { color: var(--muted); background: #e3efff; border-top: 2px solid #b8cdf0; }
.empty-state, .empty-state strong, .empty-state span { color: var(--muted); }
.empty-state svg { color: #8fa8d8; }

.page-footer, .page-footer span span { color: #c6defc; }

/* Dialogs */
dialog {
  color: var(--outline);
  background: #f3f8ff;
  border: 3px solid var(--outline);
  border-radius: 18px;
  box-shadow: 0 8px 0 var(--shadow), 0 24px 60px #0008;
}
dialog::backdrop { background: #0a1946b3; }
.dialog-header h2 { color: var(--outline); font-weight: 700; }
.dialog-header .section-kicker { color: #e8741a; }
.form-grid label { color: var(--outline); }
.form-error { color: #d12b2f; font-weight: 700; }
.confirm-message { color: var(--outline); }

/* Date picker */
.flatpickr-calendar { background: #fff; color: var(--outline); border: 3px solid var(--outline); border-radius: 14px; box-shadow: 0 6px 0 var(--shadow), 0 16px 40px #0006; }
.flatpickr-months .flatpickr-month, .flatpickr-current-month .cur-month, .flatpickr-current-month input.cur-year { color: var(--outline); fill: var(--outline); }
.flatpickr-months .flatpickr-prev-month, .flatpickr-months .flatpickr-next-month { color: #2a7fe0; fill: #2a7fe0; }
.flatpickr-months .flatpickr-prev-month:hover, .flatpickr-months .flatpickr-next-month:hover { background: #dbeaff; }
.flatpickr-months .flatpickr-prev-month:hover svg, .flatpickr-months .flatpickr-next-month:hover svg { fill: #1b5fb8; }
span.flatpickr-weekday { color: #2a7fe0; }
.flatpickr-day { color: var(--outline); }
.flatpickr-day:hover, .flatpickr-day:focus { border-color: #b8c8e8; background: #dbeaff; }
.flatpickr-day.prevMonthDay, .flatpickr-day.nextMonthDay, .flatpickr-day.flatpickr-disabled { color: #a4b4d0; }
.flatpickr-day.today, .flatpickr-day.today:hover { border-color: #f0761e; color: var(--outline); background: #fff3c9; }
.flatpickr-day.selected, .flatpickr-day.selected:hover { color: #fff; border-color: var(--outline); background: linear-gradient(#ffb04a, #f0761e); box-shadow: 0 2px 0 var(--shadow); text-shadow: 0 1px 0 #7a3200; }
.fp-today { color: #5a2d00; border: 2px solid var(--gold-edge); background: linear-gradient(#ffe65f, #ffb81c); }
.fp-today:hover { border-color: var(--gold-edge); background: linear-gradient(#fff07a, #ffc226); }

.toast {
  color: #fff;
  background: linear-gradient(#4aa0f7, #2370d0);
  border: 3px solid var(--outline);
  border-radius: 12px;
  box-shadow: inset 0 2px 0 #ffffff66, 0 4px 0 var(--shadow);
  font-weight: 700;
}

@media (max-width: 1050px) {
  .overview-intro, .metric:nth-child(2) { border-bottom: 2px dashed #ffffff40; }
}
@media (max-width: 700px) {
  .sidebar { border-bottom: 3px solid var(--outline); border-right: 0; }
}

````

### 5.13 `static\js\app.js`

<!-- REBUILD_FILE {"path":"static\\js\\app.js","newline":"LF","encoding":"utf-8","size":36408,"sha256":"814c51d62c59f6f7fbe6047841ac2dcc36af052c545d18f7f8bdf2ab229e3d60"} -->
````javascript
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
  items = await request('/api/items');
  if (currentGroup !== 'all' && !items.some(item => item.group === currentGroup)) currentGroup = 'all';
  render();
}
async function writeItem(item, isNew = false) {
  const saved = await request(isNew ? '/api/items' : `/api/items/${encodeURIComponent(item.id)}`, {
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
      await request(`/api/items/${encodeURIComponent(item.id)}`, {method:'DELETE',headers:{'Content-Type':'application/json'},body:JSON.stringify({version:item.version})});
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
$('#export').addEventListener('click', () => { if (fileMode) { notify('请通过本地服务导出 Excel'); return; } const link = document.createElement('a'); link.href = '/api/export'; link.download = 'upload-desk.xlsx'; document.body.append(link); link.click(); link.remove(); });
$('#export-customer').addEventListener('click', async () => {
  if (fileMode) { notify('请通过本地服务导出 Excel'); return; }
  try {
    const response = await fetch('/api/export/customer');
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
    const data = await request('/api/export/json');
    downloadBackup(JSON.stringify(data, null, 2), 'upload-desk-backup.json', 'application/json;charset=utf-8');
  } catch (error) { notify(`JSON 导出失败：${error.message}`); }
});
$('#excel-file').addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    const preview = await request('/api/import/preview', {method:'POST',headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},body:file});
    if (!await askConfirm(`工作表：${preview.sheet}\n共 ${preview.total} 条；新增 ${preview.inserted} 条，更新 ${preview.updated} 条。\n匹配的文件夹保留已有上传数量和名义上传者，其他记录保留。确定导入吗？`, {title:'导入 Excel',okText:'导入'})) return;
    const result = await request('/api/import', {method:'POST',headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'},body:file});
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
    const result = await request('/api/migrate', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(normalized)});
    await refreshItems(); notify(`JSON 备份已导入：新增 ${result.inserted} 条，更新 ${result.updated} 条`);
  } catch (error) { await askConfirm(`导入失败：${error.message}`, {title:'导入失败',alertOnly:true}); }
  finally { event.target.value = ''; }
});
if (typeof flatpickr !== 'undefined') {
  document.querySelectorAll('[data-datepicker]').forEach(input => flatpickr(input, {
    locale: flatpickr.l10ns.zh, dateFormat: 'Y-m-d', allowInput: true, disableMobile: true, monthSelectorType: 'static',
    appendTo: input.closest('dialog') || undefined,
    onReady: (_, __, picker) => {
      const today = document.createElement('button');
      today.type = 'button'; today.className = 'fp-today'; today.textContent = '今天';
      today.addEventListener('click', () => { picker.setDate(new Date(), true); picker.close(); });
      picker.calendarContainer.append(today);
    },
    onOpen: (_, __, picker) => picker.setDate(picker.input.value, false)
  }));
}
$('#today').textContent = new Intl.DateTimeFormat('zh-CN',{year:'numeric',month:'long',day:'numeric'}).format(new Date());
render();
if (fileMode) {
  $('#storage-status').textContent = '旧数据只读，请导出备份';
} else {
  refreshItems().then(() => { $('#storage-status').textContent = '数据保存在 SQLite'; }).catch(error => { $('#storage-status').textContent = '数据库连接失败'; notify(error.message); });
}

````

### 5.14 `static\js\planner.js`

<!-- REBUILD_FILE {"path":"static\\js\\planner.js","newline":"LF","encoding":"utf-8","size":29960,"sha256":"59f6ddcaf546b474115bd2b2fbfbb997fae696642ef61d5369f3f3e84d6fbff6"} -->
````javascript
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
    const result = await request('/api/file-lists', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({sources})});
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
  const zip = plan.hasArchive ? `<a class="history-download" href="/api/allocation-history/${encodeURIComponent(plan.id)}/archive"><i data-lucide="download"></i> ${zipLabel}</a>` : '';
  return copy + zip;
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
    const history = await request('/api/allocation-history');
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
    const result = await request('/api/allocations/confirm', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({label,manifest,assignments:plan.assignments.map(({item,person}) => ({id:item.id,version:item.version,uploader:people[person],uploadDate}))})});
    const saved = result.items;
    const updated = new Map(saved.map(item => [item.id,item]));
    items = items.map(item => updated.get(item.id) || item);
    marked = true;
    if ($('#allocation-history').open) loadAllocationHistory();
    await copyText(JSON.stringify(manifest, null, 2));
    invalidatePlans();
    render();
    plannerMessage(`已为 ${paths.length} 个文件夹保存上传计划，方案 JSON 已复制到剪贴板；文件尚未标记为已上传。`);
    notify('方案 JSON 已复制到剪贴板');
  } catch (error) {
    if (marked) { invalidatePlans(); render(); }
    plannerMessage(marked ? `已保存上传计划，但复制 JSON 失败：${error.message}。可在“历史分配方案”中重新复制。` : `分配确认失败，未保存上传计划：${error.message}`,true);
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

````

### 5.15 `data\step1_copy_files.py`

<!-- REBUILD_FILE {"path":"data\\step1_copy_files.py","newline":"LF","encoding":"utf-8","size":8175,"sha256":"42fdfc8a7c4561acf0c227560b04bbb2e18ebe69de3704ebd7cdc345d556e954"} -->
````python
import ctypes
import hashlib
import json
import os
import shutil
import sys
import uuid
from pathlib import Path


SOURCE_ROOT = r"EDIT_SOURCE_ROOT"

SCRIPT_DIR = Path(__file__).resolve().parent
LOCAL_ROOT = SCRIPT_DIR.parent
MANIFEST_PATH = SCRIPT_DIR / "allocation.json"


def configure_stdio():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                try:
                    stream.reconfigure(errors="replace")
                except Exception:
                    pass


def get_desktop_path():
    if os.name != "nt":
        raise ValueError("仅 Windows 支持使用桌面文件夹名称，请填写绝对路径。")
    desktop_id = ctypes.create_string_buffer(
        uuid.UUID("B4BFCC3A-DB2C-424C-B029-7FE99A87C641").bytes_le, 16
    )
    shell = ctypes.WinDLL("shell32", use_last_error=True)
    memory = ctypes.WinDLL("ole32")
    shell.SHGetKnownFolderPath.argtypes = [
        ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)
    ]
    shell.SHGetKnownFolderPath.restype = ctypes.c_long
    memory.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    memory.CoTaskMemFree.restype = None
    desktop = ctypes.c_void_p()
    result = shell.SHGetKnownFolderPath(ctypes.byref(desktop_id), 0, None, ctypes.byref(desktop))
    if result != 0:
        raise OSError(f"无法获取 Windows 桌面路径，错误代码：0x{result & 0xFFFFFFFF:08X}")
    try:
        return Path(ctypes.wstring_at(desktop.value))
    finally:
        memory.CoTaskMemFree(desktop)


def resolve_configured_path(value):
    raw_value = os.fspath(value)
    path = Path(raw_value).expanduser()
    if path.is_absolute():
        return path
    if path.name != raw_value or raw_value in ("", ".", ".."):
        raise ValueError(f"请填写完整绝对路径，或仅填写桌面上的文件夹名称：{raw_value}")
    return get_desktop_path() / path


def load_allocations():
    with MANIFEST_PATH.open(encoding="utf-8-sig") as manifest_file:
        allocations = json.load(manifest_file)
    if not isinstance(allocations, list) or not allocations:
        raise ValueError("分配清单为空或格式不正确。")
    for allocation in allocations:
        if not isinstance(allocation, dict):
            raise ValueError("分配清单格式不正确。")
        for field in ("source", "destination"):
            parts = allocation.get(field)
            if not isinstance(parts, list) or not parts:
                raise ValueError("分配清单中的路径格式不正确。")
            for part in parts:
                if (not isinstance(part, str) or not part.strip()
                        or part in (".", "..") or part.endswith((" ", "."))
                        or any(char in '<>:"/\\|?*' or ord(char) < 32 for char in part)):
                    raise ValueError("分配清单中包含无效路径。")
    return allocations


def path_key(path):
    return os.path.normcase(os.path.abspath(path))


def paths_overlap(first, second):
    first_key = path_key(first)
    second_key = path_key(second)
    try:
        common = os.path.commonpath((first_key, second_key))
    except ValueError:
        return False
    return common == first_key or common == second_key


def is_linked_folder(path):
    return path.is_symlink() or bool(getattr(path.lstat(), "st_file_attributes", 0) & 0x400)


def file_hash(path):
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.digest()


def check_existing_file(source, destination, allow_overwrite=False):
    if destination.is_dir():
        raise ValueError(f"目标位置已有同名文件夹：{destination}")
    if not allow_overwrite and destination.exists() and (source.stat().st_size != destination.stat().st_size
                                                         or file_hash(source) != file_hash(destination)):
        raise ValueError(f"目标位置已有不同的文件，不会覆盖：{destination}")


def copy_allocations(source_root, destination_root, allocations, source_field, overwrite_existing=None):
    source_root = resolve_configured_path(source_root)
    destination_root = resolve_configured_path(destination_root)
    if not source_root.is_absolute() or not source_root.is_dir():
        raise ValueError(f"源目录必须是已存在的绝对路径：{source_root}")
    if not destination_root.is_absolute() or not destination_root.is_dir():
        raise ValueError(f"目标目录必须是已存在的绝对路径：{destination_root}")
    source_root = source_root.resolve()
    destination_root = destination_root.resolve()
    if paths_overlap(source_root, destination_root):
        raise ValueError("源目录和目标目录不能相同或互相包含。")

    folders = {}
    files = {}
    source_folders = []
    for allocation in allocations:
        source = source_root.joinpath(*allocation[source_field])
        destination = destination_root.joinpath(*allocation["destination"])
        if not source.is_dir():
            raise ValueError(f"找不到源文件夹：{source}")
        if is_linked_folder(source):
            raise ValueError(f"不支持链接文件夹：{source}")
        if any(paths_overlap(source, existing) for existing in source_folders):
            raise ValueError(f"分配清单中的源文件夹互相重叠：{source}")
        source_folders.append(source)

        def raise_walk_error(error):
            raise error

        for current, directory_names, file_names in os.walk(source, onerror=raise_walk_error):
            current_path = Path(current)
            target_folder = destination / current_path.relative_to(source)
            folders[path_key(target_folder)] = target_folder
            for directory_name in directory_names:
                directory = current_path / directory_name
                if is_linked_folder(directory):
                    raise ValueError(f"不支持链接文件夹：{directory}")
            for file_name in file_names:
                source_file = current_path / file_name
                target_file = target_folder / file_name
                key = path_key(target_file)
                if key in files:
                    raise ValueError(f"多个源文件对应同一目标：{target_file}")
                files[key] = (source_file, target_file)

    for folder in folders.values():
        current = folder
        while current != destination_root:
            if current.is_file() or path_key(current) in files:
                raise ValueError(f"目标文件夹的位置已有文件：{current}")
            current = current.parent
    for source_file, target_file in files.values():
        check_existing_file(source_file, target_file, allow_overwrite=overwrite_existing is not None)

    for folder in folders.values():
        folder.mkdir(parents=True, exist_ok=True)
    copied = 0
    for source_file, target_file in files.values():
        check_existing_file(source_file, target_file, allow_overwrite=overwrite_existing is not None)
        if target_file.exists() and overwrite_existing is not None:
            overwrite_existing(source_file, target_file)
            copied += 1
        elif not target_file.exists():
            with source_file.open("rb") as source_stream, target_file.open("xb") as target_stream:
                shutil.copyfileobj(source_stream, target_stream)
            copied += 1
    return copied, len(files)


def main():
    configure_stdio()
    try:
        copied, total = copy_allocations(SOURCE_ROOT, LOCAL_ROOT, load_allocations(), "source")
        print(f"第一步完成：共 {total} 个文件，新复制 {copied} 个文件。")
    except (OSError, ValueError, KeyError) as error:
        print(f"第一步失败：{error}")
        print("复制可能未完成；请解决问题后重新运行。")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
````

### 5.16 `data\step2_copy_onedrive.py`

<!-- REBUILD_FILE {"path":"data\\step2_copy_onedrive.py","newline":"LF","encoding":"utf-8","size":2445,"sha256":"f6dea181f4322e1afe4491f9dd72cb44188cf3c1c77cca8bfd81fac181e75fec"} -->
````python
import json
import os
import shutil
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path

from step1_copy_files import (
    LOCAL_ROOT,
    SCRIPT_DIR,
    configure_stdio,
    copy_allocations,
    file_hash,
    load_allocations,
)


ONEDRIVE_ROOT = r"EDIT_ONEDRIVE_SYNC_FOLDER"
LOG_PATH = SCRIPT_DIR / "onedrive-overwrite-log.jsonl"


def write_log(record):
    record["time"] = datetime.now(timezone.utc).isoformat()
    with LOG_PATH.open("a", encoding="utf-8") as log_file:
        log_file.write(json.dumps(record, ensure_ascii=False) + "\n")
        log_file.flush()
        os.fsync(log_file.fileno())


def overwrite_with_log(source, destination):
    temporary_path = None
    operation_id = uuid.uuid4().hex
    try:
        with source.open("rb") as source_file:
            with tempfile.NamedTemporaryFile(dir=destination.parent, delete=False) as temporary_file:
                temporary_path = temporary_file.name
                shutil.copyfileobj(source_file, temporary_file)
        write_log({
            "id": operation_id,
            "status": "pending",
            "source": str(source),
            "destination": str(destination),
            "old_sha256": file_hash(destination).hex(),
            "new_sha256": file_hash(Path(temporary_path)).hex(),
        })
        try:
            os.replace(temporary_path, destination)
        except OSError as error:
            write_log({"id": operation_id, "status": "failed", "error": str(error)})
            raise
        temporary_path = None
        write_log({"id": operation_id, "status": "completed"})
    finally:
        if temporary_path is not None:
            os.unlink(temporary_path)


def main():
    configure_stdio()
    try:
        copied, total = copy_allocations(
            LOCAL_ROOT, ONEDRIVE_ROOT, load_allocations(), "destination", overwrite_with_log
        )
        print(f"第二步完成：共 {total} 个文件，已复制或覆盖 {copied} 个文件到 OneDrive 同步文件夹。")
        print(f"覆盖记录（发生覆盖时生成）：{LOG_PATH}")
        print("请检查 OneDrive 客户端，确认云端同步已完成。")
    except (OSError, ValueError, KeyError) as error:
        print(f"第二步失败：{error}")
        print("复制可能未完成；请解决问题后重新运行。")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
````

### 5.17 `tests\test_launcher.py`

<!-- REBUILD_FILE {"path":"tests\\test_launcher.py","newline":"LF","encoding":"utf-8","size":4239,"sha256":"f52a0eec7a704aa6fce9c572518aa890d9e1985c6e2b532e6495a665b59c1689"} -->
````python
import errno
import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import server
import start


class LauncherTests(unittest.TestCase):
    def test_existing_dependencies_do_not_install(self):
        with tempfile.TemporaryDirectory() as directory:
            jszip = Path(directory) / 'jszip.min.js'
            jszip.touch()
            with patch.object(start, 'JSZIP', jszip), patch.object(start.metadata, 'version', return_value='3.1.5'), patch.object(start.subprocess, 'run') as run:
                start.ensure_dependencies()
                run.assert_not_called()

    def test_missing_dependencies_are_installed(self):
        with tempfile.TemporaryDirectory() as directory:
            jszip = Path(directory) / 'jszip.min.js'

            def install(command, **kwargs):
                if command[-1] == 'install':
                    jszip.touch()

            with patch.object(start, 'JSZIP', jszip), patch.object(start.metadata, 'version', side_effect=start.metadata.PackageNotFoundError), patch.object(start.shutil, 'which', return_value='npm'), patch.object(start.subprocess, 'run', side_effect=install) as run:
                start.ensure_dependencies()
                self.assertEqual(run.call_count, 2)
                self.assertEqual(run.call_args_list[0].args[0][1:4], ['-m', 'pip', 'install'])
                self.assertEqual(run.call_args_list[1].args[0], ['npm', 'install'])

    def test_existing_service_opens_browser(self):
        with patch.object(start, 'ensure_dependencies'), patch.object(server, 'LocalHTTPServer', side_effect=OSError(errno.EADDRINUSE, 'in use')), patch.object(start, 'existing_service', return_value=True), patch.object(start, 'open_browser') as browser:
            start.main()
            browser.assert_called_once_with(f'http://127.0.0.1:{server.PORT}')

    def test_other_service_is_not_opened(self):
        with patch.object(start, 'ensure_dependencies'), patch.object(server, 'LocalHTTPServer', side_effect=OSError(errno.EADDRINUSE, 'in use')), patch.object(start, 'existing_service', return_value=False), patch.object(start, 'open_browser') as browser:
            with self.assertRaisesRegex(RuntimeError, '端口'):
                start.main()
            browser.assert_not_called()

    def test_browser_failure_does_not_stop_service(self):
        with patch.object(start.webbrowser, 'open', side_effect=start.webbrowser.Error('no browser')):
            start.open_browser('http://127.0.0.1:8000')

    def test_start_initializes_server_and_opens_browser(self):
        with patch.object(start, 'ensure_dependencies'), patch.object(server, 'LocalHTTPServer') as server_class, patch.object(server, 'init_db') as init_db, patch.object(start, 'open_browser') as browser:
            start.main()
            init_db.assert_called_once()
            browser.assert_called_once_with(f'http://127.0.0.1:{server.PORT}')
            server_class.return_value.serve_forever.assert_called_once()

    def test_cp932_stdio_does_not_fail_on_chinese_output(self):
        raw_stdout = io.BytesIO()
        raw_stderr = io.BytesIO()
        cp932_stdout = io.TextIOWrapper(raw_stdout, encoding='cp932', errors='strict')
        cp932_stderr = io.TextIOWrapper(raw_stderr, encoding='cp932', errors='strict')
        with patch.object(start.sys, 'stdout', cp932_stdout), patch.object(start.sys, 'stderr', cp932_stderr):
            with patch.object(start, 'ensure_dependencies'), patch.object(server, 'LocalHTTPServer') as server_class, patch.object(server, 'init_db'), patch.object(start.webbrowser, 'open', return_value=True):
                start.main()
                server_class.return_value.serve_forever.assert_called_once()
        cp932_stdout.flush()
        output = raw_stdout.getvalue().decode('utf-8')
        self.assertIn('服务已启动，关闭此窗口将停止服务。', output)

    def test_server_port_is_exclusive(self):
        with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
            with self.assertRaises(OSError):
                server.LocalHTTPServer(http_server.server_address, server.Handler).server_close()


if __name__ == '__main__':
    unittest.main()

````

### 5.18 `tests\test_server.py`

<!-- REBUILD_FILE {"path":"tests\\test_server.py","newline":"LF","encoding":"utf-8","size":37120,"sha256":"3693b823fc56db51366547735de2eecaf1e8cb2996756e9dab6e4ab927467622"} -->
````python
import io
import json
import sqlite3
import tempfile
import threading
import unittest
from datetime import date, datetime
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from unittest.mock import patch

from openpyxl import load_workbook

import server
from server import export_workbook, parse_workbook

WORKBOOK = Path(__file__).resolve().parent.parent / 'data' / '設計書の履歴シート削除一覧.xlsx'


MANIFEST = {'version': 1, 'label': '方案 1', 'people': [{'name': '甲', 'totalFiles': 10, 'folders': [{'source': ['父', '一'], 'destination': ['甲', '父', '一'], 'files': 10}]}]}


class AllocationDateTests(unittest.TestCase):
    def test_list_source_files_is_relative_sorted_and_safe(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory) / '父' / '子'
            (folder / '旧版').mkdir(parents=True)
            (folder / 'b.xlsx').write_text('x')
            (folder / '旧版' / 'a.xls').write_text('x')
            with patch.dict('os.environ', {'SOURCE_ROOT': directory}):
                self.assertEqual(server.list_source_files(['父', '子']), ['b.xlsx', '旧版/a.xls'])
                self.assertEqual(server.list_source_files(['父', '..']), [])
                self.assertEqual(server.list_source_files(['不存在']), [])
            with patch.dict('os.environ', {'SOURCE_ROOT': ''}):
                self.assertEqual(server.list_source_files(['父', '子']), [])

    def test_invalid_history_rolls_back_all_assignments(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [
                    {'group':'父', 'name':'一', 'total':10, 'processed':10, 'date':'2026-10-30', 'complete':True},
                    {'group':'父', 'name':'二', 'total':10, 'processed':10, 'date':'2026-10-30', 'complete':True},
                ])
                first, second = server.all_items(connection)
            assignments = [
                {'id':first['id'], 'version':first['version'], 'uploader':'甲', 'uploadDate':'2026-11-02'},
                {'id':second['id'], 'version':second['version'], 'uploader':'乙', 'uploadDate':'2026-11-03'},
            ]
            payload = {'label':'方案 1', 'manifest':MANIFEST, 'assignments':assignments}
            invalid_payloads = [
                payload,
                {**payload, 'manifest':'not-json'},
                {**payload, 'manifest':{'version':1}},
                {**payload, 'manifest':None, 'assignments':[{**assignments[0]}, {**assignments[1], 'uploadDate':'2026-11-02'}]},
            ]
            for invalid in invalid_payloads:
                with self.assertRaises(ValueError):
                    with server.connect() as connection:
                        server.save_allocation_history(connection, invalid)
                with server.connect() as connection:
                    self.assertEqual([item['status'] for item in server.all_items(connection)], ['ready', 'ready'])
                    self.assertEqual(connection.execute('SELECT count(*) FROM allocation_history').fetchone()[0], 0)

    def test_history_http_returns_manifest_and_keeps_legacy_archive(self):
        legacy_archive = b'PK legacy zip bytes'
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [{'group':'父', 'name':'一', 'total':10, 'processed':10, 'date':'2026-10-30', 'complete':True}])
                item = server.all_items(connection)[0]
                connection.execute("INSERT INTO allocation_history (id,created_at,planned_date,label,assignments,archive) VALUES ('legacy','2026-10-01T00:00:00','2026-10-02','旧方案','[]',?)", (legacy_archive,))
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        url = f'http://127.0.0.1:{server.PORT}'
                        payload = {'label':'方案 1', 'manifest':MANIFEST,
                                   'assignments':[{'id':item['id'], 'version':item['version'], 'uploader':'甲', 'uploadDate':'2026-11-02'}]}
                        request = Request(f'{url}/api/allocations/confirm', data=json.dumps(payload).encode(), headers={'Content-Type':'application/json'}, method='POST')
                        with urlopen(request) as response:
                            saved = json.load(response)
                        with urlopen(f'{url}/api/allocation-history') as response:
                            history = {entry['id']: entry for entry in json.load(response)}
                        self.assertEqual(len(history), 2)
                        current = history[saved['historyId']]
                        self.assertEqual(current['assignments'][0]['uploader'], '甲')
                        self.assertEqual(current['plannedDate'], '2026-11-02')
                        self.assertEqual(current['manifest'], MANIFEST)
                        self.assertFalse(current['hasArchive'])
                        self.assertIsNone(history['legacy']['manifest'])
                        self.assertTrue(history['legacy']['hasArchive'])
                        with urlopen(f'{url}/api/allocation-history/legacy/archive') as response:
                            self.assertEqual(response.headers.get_content_type(), 'application/zip')
                            self.assertEqual(response.read(), legacy_archive)
                        with self.assertRaises(HTTPError) as error:
                            urlopen(f"{url}/api/allocation-history/{saved['historyId']}/archive")
                        self.assertEqual(error.exception.code, 404)
                        error.exception.close()
                        with server.connect() as connection:
                            self.assertEqual(server.all_items(connection), saved['items'])
                    finally:
                        http_server.shutdown()
                        thread.join()

    def test_history_stores_snapshot_and_manifest_atomically(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [{'group':'父', 'name':'一', 'total':10, 'processed':10, 'date':'2026-10-30', 'complete':True}])
                item = server.all_items(connection)[0]
            assignment = {'id':item['id'], 'version':item['version'], 'uploader':'甲', 'uploadDate':'2026-11-02'}
            payload = {'label':'方案 1', 'assignments':[assignment], 'manifest':MANIFEST}
            with server.connect() as connection:
                saved, history_id = server.save_allocation_history(connection, payload)
                self.assertEqual(saved[0]['status'], 'uploading')
                self.assertEqual(saved[0]['uploaded'], 0)
                history = connection.execute('SELECT * FROM allocation_history WHERE id=?', (history_id,)).fetchone()
                self.assertEqual(json.loads(history['manifest']), MANIFEST)
                self.assertEqual(history['archive'], b'')
                self.assertEqual(json.loads(history['assignments'])[0]['name'], '一')
                connection.execute('DELETE FROM items WHERE id=?', (item['id'],))
            with self.assertRaises(server.AllocationConflict):
                with server.connect() as connection:
                    server.save_allocation_history(connection, payload)
            with server.connect() as connection:
                self.assertEqual(connection.execute('SELECT count(*) FROM allocation_history').fetchone()[0], 1)
                history = connection.execute('SELECT * FROM allocation_history WHERE id=?', (history_id,)).fetchone()
                self.assertEqual(json.loads(history['assignments'])[0]['name'], '一')
                self.assertEqual(json.loads(history['manifest']), MANIFEST)

    def test_folder_status_rules_and_manual_override(self):
        item = {'total': 10, 'processed': 0, 'date': '', 'uploaded': 0}
        self.assertEqual(server.folder_status(item), 'unstarted')
        self.assertEqual(server.folder_status({**item, 'date': '2026-10-30'}), 'processing')
        self.assertEqual(server.folder_status({**item, 'date': '2026-10-30', 'processed': 10}), 'ready')
        self.assertEqual(server.folder_status({**item, 'plannedDate': '2026-11-02'}), 'uploading')
        self.assertEqual(server.folder_status({**item, 'uploaded': 10}), 'uploaded')
        self.assertEqual(server.folder_status({**item, 'statusOverride': 'ready'}), 'ready')

    def test_allocation_date_and_legacy_default(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [
                    {'group': '父', 'name': '一', 'total': 10, 'processed': 10, 'date': '2026-10-30', 'complete': True},
                    {'group': '父', 'name': '二', 'total': 10, 'processed': 10, 'date': '2026-10-30', 'complete': True},
                ])
                first, second = server.all_items(connection)
            assignments = [
                {'id': first['id'], 'version': first['version'], 'uploader': '甲', 'uploadDate': '2026-11-02'},
                {'id': second['id'], 'version': second['version'], 'uploader': '乙', 'uploadDate': 'invalid'},
            ]
            with self.assertRaises(ValueError):
                with server.connect() as connection:
                    server.mark_allocations(connection, assignments)
            with server.connect() as connection:
                self.assertEqual([item['uploaded'] for item in server.all_items(connection)], [0, 0])
                self.assertEqual([item['plannedDate'] for item in server.all_items(connection)], ['', ''])
                del assignments[1]['uploadDate']
                saved = server.mark_allocations(connection, assignments)
                self.assertEqual([item['plannedDate'] for item in saved], ['2026-11-02', date.today().isoformat()])
                self.assertEqual([item['uploaded'] for item in saved], [0, 0])
                self.assertEqual([item['uploadDate'] for item in saved], ['', ''])
                self.assertEqual([item['status'] for item in saved], ['uploading', 'uploading'])
                exported = server.export_workbook(saved)
            parsed, _ = server.parse_workbook(exported)
            self.assertEqual([(item['plannedDate'], item['statusOverride'], item['uploader']) for item in parsed],
                             [('2026-11-02', 'uploading', '甲'), (date.today().isoformat(), 'uploading', '乙')])
            with server.connect() as connection:
                server.merge_items(connection, parsed, preserve_upload=True)
                self.assertEqual([item['status'] for item in server.all_items(connection)], ['uploading', 'uploading'])


class ScriptResourceTests(unittest.TestCase):
    def test_script_resource_downloads(self):
        with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
            with patch.object(server, 'PORT', http_server.server_address[1]):
                thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                thread.start()
                try:
                    for name in ('step1_copy_files.py', 'step2_copy_onedrive.py'):
                        with urlopen(f'http://127.0.0.1:{server.PORT}/data/{name}') as response:
                            self.assertEqual(response.read(), (server.ROOT / 'data' / name).read_bytes())
                            self.assertEqual(response.headers.get_content_type(), 'text/plain')
                finally:
                    http_server.shutdown()
                    thread.join()


class WorkbookTests(unittest.TestCase):
    def test_excel_import_api_keeps_upload_and_rolls_back_conflicts(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        original = {**records[0], 'uploaded': records[0]['total'], 'uploader': '原上传者'}
        incoming = {**records[0], 'uploaded': 0, 'uploader': '', 'owner': '新负责人'}
        new_item = {**records[2], 'uploaded': records[2]['total'], 'uploader': '新上传者'}
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [original])
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        url = f'http://127.0.0.1:{server.PORT}/api/import'
                        def import_excel(items):
                            request = Request(url, data=export_workbook(items),
                                              headers={'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}, method='POST')
                            return urlopen(request)
                        with import_excel([incoming, new_item]) as response:
                            self.assertEqual(json.load(response), {'inserted': 1, 'updated': 1})
                        with server.connect() as connection:
                            first, second = server.all_items(connection)
                        self.assertEqual((first['uploaded'], first['uploader'], first['owner']), (original['uploaded'], '原上传者', '新负责人'))
                        self.assertEqual((second['uploaded'], second['uploader']), (new_item['uploaded'], '新上传者'))
                        conflict = {**incoming, 'complete': False}
                        with self.assertRaises(HTTPError) as error:
                            import_excel([{**new_item, 'owner': '不应保存'}, conflict])
                        self.assertEqual(error.exception.code, 400)
                        self.assertIn('未导入任何记录', json.load(error.exception)['error'])
                        error.exception.close()
                        with server.connect() as connection:
                            first, second = server.all_items(connection)
                        self.assertEqual((first['uploaded'], first['uploader'], first['complete']), (original['uploaded'], '原上传者', True))
                        self.assertNotEqual(second['owner'], '不应保存')
                    finally:
                        http_server.shutdown()
                        thread.join()

    def test_excel_merge_preserves_uploaded_values_even_when_columns_exist(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        original = {**records[0], 'uploaded': records[0]['total'], 'uploader': '原上传者'}
        incoming = {**records[0], 'uploaded': 0, 'uploader': '', 'owner': '新负责人'}
        parsed, _ = parse_workbook(export_workbook([incoming]))
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [original])
            with server.connect() as connection:
                server.merge_items(connection, parsed, preserve_upload=True)
                saved = server.all_items(connection)[0]
            self.assertEqual((saved['uploaded'], saved['uploader'], saved['owner']), (original['uploaded'], '原上传者', '新负责人'))
            parsed[0]['complete'] = False
            with self.assertRaises(ValueError):
                with server.connect() as connection:
                    server.merge_items(connection, parsed, preserve_upload=True)
            with server.connect() as connection:
                saved = server.all_items(connection)[0]
                self.assertEqual((saved['uploaded'], saved['uploader'], saved['complete']), (original['uploaded'], '原上传者', True))

    def test_allocation_is_atomic_and_records_uploader(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [records[0], records[2]])
                first, second = server.all_items(connection)
            assignments = [{'id': first['id'], 'version': first['version'], 'uploader': '甲'},
                           {'id': second['id'], 'version': second['version'] + 1, 'uploader': '乙'}]
            with self.assertRaises(server.AllocationConflict):
                with server.connect() as connection:
                    server.mark_allocations(connection, assignments)
            with server.connect() as connection:
                self.assertEqual([item['uploaded'] for item in server.all_items(connection)], [0, 0])
                assignments[1]['version'] = second['version']
                saved = server.mark_allocations(connection, assignments)
                self.assertEqual([(item['uploaded'], item['uploader']) for item in saved],
                                 [(0, '甲'), (0, '乙')])
                self.assertEqual([item['status'] for item in saved], ['uploading', 'uploading'])

    def test_customer_export_and_backup_uploader_round_trip(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        records[0].update(uploaded=records[0]['total'], uploader='甲', uploadDate='2026-10-29')
        records[1].update(total=145, processed=145, complete=True, uploaded=10, uploader='张晓静', uploadDate='2026-10-30')
        result = load_workbook(io.BytesIO(server.export_customer_workbook(records)))
        template = load_workbook(server.CUSTOMER_TEMPLATE)
        customer = result.active
        self.assertEqual(result.sheetnames, template.sheetnames)
        self.assertEqual(customer.max_row, 87)
        self.assertEqual([cell.value for cell in customer[2]], [cell.value for cell in template.active[2]])
        self.assertEqual(str(customer.merged_cells), str(template.active.merged_cells))
        self.assertEqual(customer.auto_filter.ref, template.active.auto_filter.ref)
        self.assertEqual(customer.column_dimensions['B'].width, template.active.column_dimensions['B'].width)
        self.assertEqual((customer['E3'].value, customer['F3'].value, customer['G3'].value), (datetime(2026, 10, 29), '甲', 2))
        self.assertEqual((customer['E4'].value, customer['F4'].value, customer['G4'].value), (datetime(2026, 10, 30), '静', 10))
        self.assertEqual([customer.cell(5, col).value for col in (5, 6, 7)], [None, None, None])
        self.assertEqual(customer['H4'].value, '=IF(AND(D4<>"",G4<>"",D4=G4),"〇","")')
        self.assertEqual(customer['D87'].value, '=SUM(D3:D86)')
        self.assertEqual(customer['G87'].value, '=SUM(G3:G86)')
        self.assertEqual(customer['E3'].number_format, template.active['E5'].number_format)
        other = result['変更対象']
        expected_other = template['変更対象']
        self.assertEqual([[cell.value for cell in row] for row in other], [[cell.value for cell in row] for row in expected_other])
        reloaded, _ = parse_workbook(export_workbook(records))
        self.assertEqual([item['uploader'] for item in reloaded[:2]], ['甲', '张晓静'])
        self.assertEqual([item['uploadDate'] for item in reloaded[:2]], ['2026-10-29', '2026-10-30'])
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, reloaded)
                saved = server.all_items(connection)
                self.assertEqual([item['uploader'] for item in saved[:2]], ['甲', '张晓静'])
                self.assertEqual(saved[0]['uploadDate'], '2026-10-29')
            originals, _ = parse_workbook(WORKBOOK.read_bytes())
            with server.connect() as connection:
                server.merge_items(connection, originals[:1])
                self.assertEqual(server.all_items(connection)[0]['uploader'], '甲')
                self.assertEqual(server.all_items(connection)[0]['uploadDate'], '2026-10-29')
            self.assertEqual(server.validate_item({**saved[0], 'uploaded': 0})['uploader'], '')
            self.assertEqual(server.validate_item({**saved[0], 'uploaded': 0})['uploadDate'], '')

    def test_customer_export_added_deleted_and_special_folders(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        first = records[0]
        extra = {**first, 'name': '=新文件夹', 'total': 4, 'uploaded': 2, 'uploader': '新担当', 'uploadDate': '2026-11-01'}
        new_group = {**extra, 'group': '新增父文件夹', 'name': '子文件夹'}
        customer = load_workbook(io.BytesIO(server.export_customer_workbook([first, extra, new_group]))).active
        self.assertEqual(customer.max_row, 6)
        self.assertEqual(str(customer.merged_cells), 'B3:B4')
        self.assertEqual(customer['C4'].value, '=新文件夹')
        self.assertEqual(customer['C4'].data_type, 's')
        self.assertEqual((customer['E4'].value, customer['F4'].value, customer['G4'].value), (datetime(2026, 11, 1), '新担当', 2))
        self.assertEqual(customer['B5'].value, '新增父文件夹')
        self.assertEqual(customer['D6'].value, '=SUM(D3:D5)')
        empty = load_workbook(io.BytesIO(server.export_customer_workbook([]))).active
        self.assertEqual(empty['D3'].value, '=0')

    def test_customer_export_matches_complete_example(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        group = '10要件定義\\10画面'
        uploads = {
            '02_SG総合メニュー': (6, '呉'), '04_KH介護保険業務': (63, '呉'),
            '05_IR医療保険業務': (22, '张晓静'), '06_SS障害者支援業務': (57, '张晓静'),
            '12_KE契約管理業務': (2, '呉')
        }
        for item in records:
            if item['group'] == group and item['name'] in uploads:
                item['uploaded'], item['uploader'] = uploads[item['name']]
                item['uploadDate'] = '2026-10-29'
        expected = load_workbook(server.CUSTOMER_TEMPLATE)
        actual = load_workbook(io.BytesIO(server.export_customer_workbook(records)))
        self.assertEqual(actual.sheetnames, expected.sheetnames)
        for expected_sheet, actual_sheet in zip(expected, actual):
            self.assertEqual(str(actual_sheet.merged_cells), str(expected_sheet.merged_cells))
            self.assertEqual((actual_sheet.max_row, actual_sheet.max_column), (expected_sheet.max_row, expected_sheet.max_column))
            for row in expected_sheet:
                for cell in row:
                    with self.subTest(sheet=expected_sheet.title, cell=cell.coordinate):
                        self.assertEqual(actual_sheet[cell.coordinate].value, cell.value)
                        self.assertEqual(actual_sheet[cell.coordinate]._style, cell._style)

    def test_allocation_and_customer_http_endpoints(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [records[0], records[2], records[1]])
                first, second, _ = server.all_items(connection)
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        url = f'http://127.0.0.1:{server.PORT}'
                        assignments = [{'id': first['id'], 'version': first['version'], 'uploader': '甲', 'uploadDate': '2026-11-02'},
                                       {'id': second['id'], 'version': second['version'], 'uploader': '乙'}]
                        def confirm():
                            request = Request(f'{url}/api/allocations/confirm', data=json.dumps(assignments).encode(),
                                              headers={'Content-Type': 'application/json'}, method='POST')
                            return urlopen(request)
                        with confirm() as response:
                            saved = json.load(response)
                        self.assertEqual([item['uploader'] for item in saved], ['甲', '乙'])
                        self.assertEqual([item['plannedDate'] for item in saved], ['2026-11-02', date.today().isoformat()])
                        self.assertEqual([item['uploadDate'] for item in saved], ['', ''])
                        self.assertEqual([item['status'] for item in saved], ['uploading', 'uploading'])
                        with self.assertRaises(HTTPError) as error:
                            confirm()
                        self.assertEqual(error.exception.code, 409)
                        error.exception.close()
                        with urlopen(f'{url}/api/export/customer') as response:
                            customer = load_workbook(io.BytesIO(response.read())).active
                        self.assertEqual(customer.max_row, 6)
                        self.assertEqual([customer.cell(row, 5).value for row in (3, 5)], [None, None])
                        for item in saved:
                            payload = {**item, 'statusOverride': 'uploaded', 'uploaded': item['total']}
                            request = Request(f"{url}/api/items/{item['id']}", data=json.dumps(payload).encode(),
                                              headers={'Content-Type': 'application/json'}, method='PUT')
                            with urlopen(request) as response:
                                updated = json.load(response)
                            self.assertEqual(updated['status'], 'uploaded')
                            self.assertEqual(updated['plannedDate'], item['plannedDate'])
                        with urlopen(f'{url}/api/export/customer') as response:
                            customer = load_workbook(io.BytesIO(response.read())).active
                        self.assertEqual([customer.cell(row, 6).value for row in (3, 5)], ['甲', '乙'])
                        self.assertEqual(customer.cell(3, 5).value.date(), date.today())
                        self.assertEqual(customer.cell(5, 5).value.date(), date.today())
                    finally:
                        http_server.shutdown()
                        thread.join()

    def test_original_workbook(self):
        records, sheet = parse_workbook(WORKBOOK.read_bytes())
        self.assertEqual(sheet, 'Sheet1')
        self.assertEqual(len(records), 84)
        self.assertEqual(sum(item['total'] for item in records), 3490)
        self.assertEqual(records[1]['processed'], 59)
        self.assertFalse(records[1]['complete'])
        self.assertEqual(records[1]['uploaded'], 0)
        self.assertTrue(records[0]['complete'])
        self.assertTrue(records[40]['parentOnly'])

    def test_export_round_trip(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        records[0]['uploaded'] = records[0]['total']
        data = export_workbook(records)
        reloaded, _ = parse_workbook(data)
        self.assertEqual(len(reloaded), len(records))
        self.assertEqual(reloaded[0]['uploaded'], records[0]['total'])
        self.assertEqual(reloaded[1]['processed'], 59)
        self.assertEqual(reloaded[40]['parentOnly'], records[40]['parentOnly'])
        self.assertEqual(load_workbook(io.BytesIO(data)).active['A1'].value, '親フォルダ')

    def test_explicit_incomplete_preserved(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        records[0]['complete'] = False
        reloaded, _ = parse_workbook(export_workbook(records))
        self.assertFalse(reloaded[0]['complete'])

    def test_edited_values_round_trip(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        edited = {**records[1], 'total': 145, 'totalMissing': False, 'processed': 145,
                  'complete': True, 'uploaded': 10, 'uploadedProvided': True,
                  'date': '2026-11-01', 'owner': '测试'}
        parent = {**records[40], 'total': 0, 'totalMissing': True, 'processed': None,
                  'complete': False, 'uploaded': 0, 'uploadedProvided': True}
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [edited, parent])
                saved = server.all_items(connection)
            reloaded, _ = parse_workbook(export_workbook(saved))
        for field in ('group', 'name', 'total', 'totalMissing', 'processed', 'complete', 'uploaded', 'date', 'owner', 'parentOnly'):
            self.assertEqual(reloaded[0][field], edited[field])
            self.assertEqual(reloaded[1][field], parent[field])

    def test_edit_api_exports_updated_counts(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [records[1]])
                item = server.all_items(connection)[0]
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        item.update(processed=item['total'], complete=True, uploaded=10)
                        url = f'http://127.0.0.1:{server.PORT}'
                        request = Request(f'{url}/api/items/{item["id"]}', data=json.dumps(item).encode(),
                                          headers={'Content-Type': 'application/json'}, method='PUT')
                        with urlopen(request) as response:
                            saved = json.load(response)
                        with urlopen(f'{url}/api/export') as response:
                            exported, _ = parse_workbook(response.read())
                        reset_payload = {**saved, 'uploaded': 0}
                        request = Request(f'{url}/api/items/{saved["id"]}', data=json.dumps(reset_payload).encode(),
                                          headers={'Content-Type': 'application/json'}, method='PUT')
                        with urlopen(request) as response:
                            reset = json.load(response)
                    finally:
                        http_server.shutdown()
                        thread.join()
            self.assertEqual((saved['processed'], saved['complete'], saved['uploaded']), (145, True, 10))
            self.assertEqual((exported[0]['processed'], exported[0]['complete'], exported[0]['uploaded']), (145, True, 10))
            self.assertEqual(saved['uploadDate'], date.today().isoformat())
            self.assertEqual(exported[0]['uploadDate'], saved['uploadDate'])
            self.assertEqual(reset['uploadDate'], '')
            self.assertEqual(reset['uploader'], '')

    def test_import_preserves_upload_and_rolls_back_conflicts(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                self.assertEqual(server.merge_items(connection, records), (84, 0))
                connection.execute('UPDATE items SET uploaded=total WHERE rowid=1')
            with server.connect() as connection:
                self.assertEqual(server.merge_items(connection, records), (0, 84))
                self.assertEqual(server.all_items(connection)[0]['uploaded'], 2)
            records[1]['owner'] = 'changed'
            records[0]['complete'] = False
            with self.assertRaises(ValueError):
                with server.connect() as connection:
                    server.merge_items(connection, [records[1], records[0], *records[2:]])
            with server.connect() as connection:
                self.assertEqual(server.all_items(connection)[1]['owner'], '剛')

    def test_json_export_and_import_round_trip(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        first = {**records[0], 'uploaded': records[0]['total'], 'uploader': '甲', 'uploadDate': '2026-10-29', 'statusOverride': 'uploaded'}
        second = {**records[1], 'processed': 145, 'complete': True, 'plannedDate': '2026-11-02', 'uploader': '乙', 'statusOverride': 'uploading'}
        third = records[40]
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [first, second, third])
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        url = f'http://127.0.0.1:{server.PORT}'
                        with urlopen(f'{url}/api/export/json') as response:
                            self.assertEqual(response.headers.get_content_type(), 'application/json')
                            self.assertIn('upload-desk-backup.json', response.headers.get('Content-Disposition', ''))
                            exported = json.load(response)
                        self.assertEqual(len(exported), 3)
                        with server.connect() as connection:
                            connection.execute('DELETE FROM items WHERE name=?', (second['name'],))
                            connection.execute("UPDATE items SET owner='旧负责人', uploaded=0, uploader='', upload_date='', status_override='' WHERE name=?", (first['name'],))
                        exported[0]['owner'] = '新负责人'
                        request = Request(f'{url}/api/migrate', data=json.dumps(exported).encode('utf-8'),
                                          headers={'Content-Type': 'application/json'}, method='POST')
                        with urlopen(request) as response:
                            self.assertEqual(json.load(response), {'inserted': 1, 'updated': 2})
                        with urlopen(f'{url}/api/export/json') as response:
                            restored = {item['name']: item for item in json.load(response)}
                    finally:
                        http_server.shutdown()
                        thread.join()
            self.assertEqual(restored[first['name']]['owner'], '新负责人')
            self.assertEqual(restored[first['name']]['uploaded'], first['total'])
            self.assertEqual(restored[first['name']]['uploader'], '甲')
            self.assertEqual(restored[first['name']]['uploadDate'], '2026-10-29')
            self.assertEqual(restored[first['name']]['status'], 'uploaded')
            self.assertEqual(restored[second['name']]['plannedDate'], '2026-11-02')
            self.assertEqual(restored[second['name']]['uploader'], '乙')
            self.assertEqual(restored[second['name']]['status'], 'uploading')
            self.assertTrue(restored[third['name']]['parentOnly'])


if __name__ == '__main__':
    unittest.main()

````

### 5.19 `tests\test_planner.js`

<!-- REBUILD_FILE {"path":"tests\\test_planner.js","newline":"LF","encoding":"utf-8","size":33705,"sha256":"12f4d546e2359acd4d0fee0045556266076949e693eb8969b74946e66fa614d1"} -->
````javascript
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const elements = new Map();
function elementFor(selector) {
  if (!elements.has(selector)) elements.set(selector, {
    handlers: {}, addEventListener(name, handler) { this.handlers[name] = handler; },
    classList: {toggle() {}}
  });
  return elements.get(selector);
}
const context = {$: elementFor, document: {handlers: {}, addEventListener(name, handler) { (this.handlers[name] ||= []).push(handler); }}, items: [], Blob, atob, escapeHTML: value => String(value).replace(/[&<>"']/g, character => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[character]))};
const appSource = fs.readFileSync(path.join(__dirname, '../static/js/app.js'), 'utf8');
vm.runInNewContext(appSource.slice(appSource.indexOf('const folderStatuses'), appSource.indexOf('function percentOf(')), context);
vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../static/js/planner.js'), 'utf8'), context);

const today = new Date();
assert.equal(elementFor('#plan-date').value, `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`);
assert.equal(elementFor('#plan-people-count').textContent, '已选 2');
assert.ok(elementFor('#plan-people-summary').innerHTML.includes('person-picker-summary-chip'));
context.items = [{owner:'<测试>'}];
context.renderPersonPicker('people');
assert.ok(elementFor('#plan-people-options').innerHTML.includes('person-picker-chip'));
assert.ok(elementFor('#plan-people-options').innerHTML.includes('&lt;测试&gt;'));
assert.ok(!elementFor('#plan-people-options').innerHTML.includes('<测试>'));
const pickerColors = context.allocationPersonMarkers();
assert.notEqual(pickerColors.get('呉').hue, pickerColors.get('静').hue);
assert.ok(elementFor('#plan-people-options').innerHTML.includes(`--person-hue:${pickerColors.get('呉').hue}`));
assert.ok(elementFor('#plan-people-summary').innerHTML.includes(`--person-hue:${pickerColors.get('呉').hue}`));
context.items = [];
context.renderPersonPicker('owner');
assert.equal(elementFor('#plan-owner-count').textContent, '全部');
elementFor('#plan-people-options').handlers.change({target:{type:'checkbox', checked:true, value:'新人员'}});
assert.equal(elementFor('#plan-people-count').textContent, '已选 3');
assert.ok(elementFor('#plan-people-summary').innerHTML.includes('新人员'));
assert.equal((elementFor('#plan-people-summary').innerHTML.match(/class="person-picker-summary-chip/g) || []).length, 3);
assert.ok(!elementFor('#plan-people-summary').innerHTML.includes('person-picker-summary-more'));
assert.ok(elementFor('#plan-people-summary').innerHTML.includes('class="sr-only"'));
elementFor('#plan-people-options').handlers.change({target:{type:'checkbox', checked:false, value:'新人员'}});
assert.equal(elementFor('#plan-people-count').textContent, '已选 2');
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:true, value:''}});
assert.equal(elementFor('#plan-owner-count').textContent, '已选 1');
assert.ok(elementFor('#plan-owner-summary').innerHTML.includes('未分配'));
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:false, value:''}});
assert.equal(elementFor('#plan-owner-count').textContent, '全部');
context.items = [{owner:'呉'}];
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:true, value:'呉'}});
context.clearPlanPreview = () => {};
elementFor('#plan-person-add').value = '临时人员';
context.addPlanPerson();
const updatedMarkers = context.allocationPersonMarkers();
assert.ok(elementFor('#plan-people-summary').innerHTML.includes(`--person-hue:${updatedMarkers.get('呉').hue}`));
assert.ok(elementFor('#plan-owner-summary').innerHTML.includes(`--person-hue:${updatedMarkers.get('呉').hue}`));
elementFor('#plan-people-options').handlers.change({target:{type:'checkbox', checked:false, value:'临时人员'}});
assert.equal(context.allocationPersonMarkers().get('呉').hue, updatedMarkers.get('呉').hue);
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:false, value:'呉'}});
context.items.push({owner:'临时来源'});
context.renderPersonPicker('owner');
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:true, value:'临时来源'}});
const sourceMarkers = context.allocationPersonMarkers();
context.items = [{owner:'呉'}];
elementFor('#plan-owner-options').handlers.change({target:{type:'checkbox', checked:false, value:'临时来源'}});
assert.equal(context.allocationPersonMarkers().get('呉').hue, sourceMarkers.get('呉').hue);
const uploadedRecords = [
  {id:'a', group:'10要件定義\\10画面', name:'02_SG', total:6, uploaded:6, uploader:'呉', owner:'呉', date:'2026-10-20', parentOnly:false},
  {id:'b', group:'10要件定義¥10画面', name:'05_IR', total:22, uploaded:22, uploader:'静', owner:'呉', date:'', parentOnly:false},
  {id:'c', group:'独立', name:'独立', total:3, uploaded:3, uploader:'', owner:'', parentOnly:true}
];
const dateJson = JSON.parse(JSON.stringify(context.uploadDateJson('2026-10-29', uploadedRecords)));
assert.deepEqual(dateJson, {version:2, label:'2026-10-29 上传记录', uploadDate:'2026-10-29', people:[
  {name:'呉', folders:[{source:['10要件定義', '10画面', '02_SG'], destination:['呉', '10要件定義', '10画面', '02_SG'], owner:'呉', startDate:'2026-10-20', files:6, fileList:[]}]},
  {name:'静', folders:[{source:['10要件定義', '10画面', '05_IR'], destination:['静', '10要件定義', '10画面', '05_IR'], owner:'呉', startDate:'', files:22, fileList:[]}]},
  {name:'未分配', folders:[{source:['独立'], destination:['未分配', '独立'], owner:'', startDate:'', files:3, fileList:[]}]}
]});
const listedJson = context.uploadDateJson('2026-10-29', uploadedRecords, [['a.xlsx', '旧版/b.xls'], [], ['c.xlsx']]);
assert.deepEqual(listedJson.people[0].folders[0].fileList, ['a.xlsx', '旧版/b.xls']);
assert.deepEqual(Object.keys(listedJson.people[0].folders[0]), ['source', 'destination', 'owner', 'startDate', 'files', 'fileList']);
assert.deepEqual(Array.from(context.uploadDateSources(uploadedRecords)[2]), ['独立']);
const dateMarkup = context.dateExportMarkup('2026-10-29', uploadedRecords);
assert.ok(dateMarkup.includes('data-copy-date-json'));
assert.ok(dateMarkup.includes('3 个文件夹 · 31 个文件'));
assert.ok(!context.dateExportMarkup('2026-10-29', []).includes('data-copy-date-json'));
assert.ok(context.dateExportMarkup('2026-10-29', []).includes('date-history-empty'));
context.items = [];
context.window = {innerHeight:650};
const pickerPanel = {classList:{remove() {}, toggle(name, value) { this.above = value; }}, getBoundingClientRect() { return {height:350}; }};
const pickerList = {style:{}, getBoundingClientRect() { return {height:250}; }};
const positionedPicker = {querySelector(selector) { return selector === '.person-picker-panel' ? pickerPanel : pickerList; }, getBoundingClientRect() { return {top:300, bottom:340}; }};
context.positionPersonPicker(positionedPicker);
assert.equal(pickerPanel.classList.above, false);
assert.equal(pickerList.style.maxHeight, '195px');
positionedPicker.getBoundingClientRect = () => ({top:460, bottom:500});
context.positionPersonPicker(positionedPicker);
assert.equal(pickerPanel.classList.above, true);
assert.equal(pickerList.style.maxHeight, '250px');

function candidates(counts) {
  return counts.map((total, index) => ({id: String(index), name: String(index), total, date: '2026-01-01'}));
}
function totals(plans) {
  return plans.map(plan => plan.totals.reduce((sum, count) => sum + count, 0));
}

const exact = totals(context.generateAllocations(candidates([81, 80, 79]), 2, 160));
assert.equal(exact[0], 160);
assert.ok(exact.every(total => total >= 155 && total <= 165));
assert.ok(totals(context.generateAllocations(candidates([81, 80]), 2, 160)).every(total => total === 161));

const oversized = totals(context.generateAllocations(candidates([200, 90, 70]), 2, 160));
assert.ok(oversized.includes(160));
assert.ok(oversized.every(total => total >= 155 && total <= 165));

const belowLimit = totals(context.generateAllocations(candidates([40, 50]), 2, 160));
assert.equal(belowLimit.length, 0);
for (const total of [145, 155]) {
  const boundary = totals(context.generateAllocations(candidates([70, total - 70]), 2, 150));
  assert.ok(boundary.length > 0);
  assert.ok(boundary.every(value => value === total));
}
assert.equal(context.generateAllocations(candidates([70, 74]), 2, 150).length, 0);
assert.equal(context.generateAllocations(candidates([70, 86]), 2, 150).length, 0);
const bothSides = totals(context.generateAllocations(candidates([70, 75, 85]), 2, 150));
assert.ok(bothSides.includes(145) && bothSides.includes(155));
assert.ok(bothSides.every(total => total >= 145 && total <= 155));
const closestFirst = context.generateAllocations([
  {id:'a', name:'a', total:73, date:'2026-01-01', owner:'甲'},
  {id:'b', name:'b', total:77, date:'2026-01-01', owner:'乙'},
  {id:'c', name:'c', total:72, date:'2026-01-01', owner:'丙'},
  {id:'d', name:'d', total:74, date:'2026-01-01', owner:'丙'}
], 2, 150);
assert.equal(totals(closestFirst)[0], 150);
assert.ok(totals(closestFirst).every(total => total >= 145 && total <= 155));
assert.ok(totals(closestFirst).some(total => total !== 150));

const alternatives = context.generateAllocations(candidates([20, 20, 20]), 2, 60);
assert.equal(alternatives.length, 3);
assert.equal(new Set(alternatives.map(plan => plan.key)).size, 3);
assert.ok(alternatives.every(plan => plan.totals.every(total => total > 0) && plan.totals.reduce((sum, total) => sum + total, 0) >= 55 && plan.totals.reduce((sum, total) => sum + total, 0) <= 65));

const sourceOptions = context.generateAllocations([
  {id:'a', name:'a', total:20, date:'2026-01-01', owner:'甲'},
  {id:'b', name:'b', total:20, date:'2026-01-01', owner:'乙'},
  {id:'c', name:'c', total:20, date:'2026-01-01', owner:'甲'},
  {id:'d', name:'d', total:20, date:'2026-01-01', owner:'甲'}
], 2, 40);
assert.equal(sourceOptions.length, 3);
assert.equal(new Set(sourceOptions[0].assignments.map(entry => entry.item.owner)).size, 1);
assert.ok(sourceOptions.every((plan, index) => !index || sourceOptions[index - 1].ownerCount <= plan.ownerCount));

const diverseSources = context.generateAllocations([
  ...Array.from({length:25}, (_, index) => ({id:`a${index}`, name:`a${String(index).padStart(2, '0')}`, total:20, date:'2026-01-01', owner:'甲'})),
  {id:'b1', name:'b1', total:20, date:'2026-01-01', owner:'乙'},
  {id:'b2', name:'b2', total:20, date:'2026-01-01', owner:'乙'},
  {id:'b3', name:'b3', total:10, date:'2026-02-01', owner:'乙'}
], 2, 50);
assert.equal(diverseSources[0].ownerCount, 1);
assert.ok(diverseSources[0].assignments.every(entry => entry.item.owner === '乙'));

const ownerBeforeDate = context.generateAllocations([
  {id:'old-a', name:'old-a', total:40, date:'2026-01-01', owner:'甲'},
  {id:'old-b', name:'old-b', total:40, date:'2026-01-01', owner:'乙'},
  {id:'new-a', name:'new-a', total:40, date:'2026-02-01', owner:'丙'},
  {id:'new-b', name:'new-b', total:40, date:'2026-02-01', owner:'丙'}
], 2, 80);
assert.ok(ownerBeforeDate.length > 0);
assert.equal(ownerBeforeDate[0].ownerCount, 1);
assert.ok(ownerBeforeDate[0].assignments.every(entry => entry.item.owner === '丙'));

const ownerBeforeTotal = context.generateAllocations([
  {id:'a1', name:'a1', total:20, date:'2026-01-01', owner:'甲'},
  {id:'a2', name:'a2', total:20, date:'2026-01-01', owner:'甲'},
  {id:'b', name:'b', total:60, date:'2026-01-01', owner:'乙'},
  {id:'c', name:'c', total:40, date:'2026-01-01', owner:'丙'}
], 2, 100);
assert.equal(totals(ownerBeforeTotal)[0], 100);
assert.equal(ownerBeforeTotal[0].ownerCount, 2);

const parentBeforeDate = context.generateAllocations([
  {id:'old-a', name:'old-a', total:40, date:'2026-01-01', owner:'甲', group:'父一'},
  {id:'old-b', name:'old-b', total:40, date:'2026-01-01', owner:'甲', group:'父二'},
  {id:'new-a', name:'new-a', total:40, date:'2026-02-01', owner:'甲', group:'父三'},
  {id:'new-b', name:'new-b', total:40, date:'2026-02-01', owner:'甲', group:'父三'}
], 2, 80);
assert.ok(parentBeforeDate.length > 0);
assert.equal(parentBeforeDate[0].groupCount, 1);
assert.ok(parentBeforeDate[0].assignments.every(entry => entry.item.group === '父三'));

const parentBeforeTotal = context.generateAllocations([
  {id:'a1', name:'a1', total:20, date:'2026-01-01', owner:'甲', group:'父一'},
  {id:'a2', name:'a2', total:20, date:'2026-01-01', owner:'甲', group:'父一'},
  {id:'b', name:'b', total:60, date:'2026-01-01', owner:'甲', group:'父二'},
  {id:'c', name:'c', total:40, date:'2026-01-01', owner:'甲', group:'父三'}
], 2, 100);
assert.ok(parentBeforeTotal.length > 0);
assert.equal(totals(parentBeforeTotal)[0], 100);
assert.equal(parentBeforeTotal[0].groupCount, 2);

const ownerBeforeParent = context.generateAllocations([
  {id:'a', name:'a', total:20, date:'2026-01-01', owner:'甲', group:'父一'},
  {id:'b', name:'b', total:20, date:'2026-01-01', owner:'甲', group:'父二'},
  {id:'c', name:'c', total:20, date:'2026-01-01', owner:'乙', group:'父三'},
  {id:'d', name:'d', total:20, date:'2026-01-01', owner:'丙', group:'父三'}
], 2, 40);
assert.ok(ownerBeforeParent.length > 0);
assert.equal(ownerBeforeParent[0].ownerCount, 1);
assert.equal(ownerBeforeParent[0].groupCount, 2);

const dated = [
  {id:'old-a', name:'old-a', total:70, date:'2026-01-01'},
  {id:'old-b', name:'old-b', total:70, date:'2026-01-01'},
  {id:'new-a', name:'new-a', total:80, date:'2026-02-01'},
  {id:'new-b', name:'new-b', total:80, date:'2026-02-01'}
];
const early = context.generateAllocations(dated, 2, 160);
assert.ok(early.length > 0);
assert.ok(early.every(plan => plan.totals.reduce((sum, total) => sum + total, 0) === 160));
assert.ok(early.every(plan => plan.assignments.every(entry => entry.item.date === '2026-02-01')));
const withFiller = context.generateAllocations([...dated, {id:'new-small', name:'new-small', total:20, date:'2026-02-01'}], 2, 160);
assert.ok(withFiller.every(plan => plan.assignments.some(entry => entry.item.id === 'old-a') && plan.assignments.some(entry => entry.item.id === 'old-b')));
assert.ok(withFiller.every(plan => plan.totals.reduce((sum, total) => sum + total, 0) === 160));
const undated = context.generateAllocations([
  {id:'dated-a', name:'dated-a', total:40, date:'2026-01-01'},
  {id:'dated-b', name:'dated-b', total:40, date:'2026-01-01'},
  {id:'undated-a', name:'undated-a', total:50, date:''},
  {id:'undated-b', name:'undated-b', total:50, date:''}
], 2, 100);
assert.ok(undated.every(plan => plan.assignments.every(entry => !entry.item.date)));
const needsTwoPeople = context.generateAllocations([
  {id:'too-large', name:'too-large', total:160, date:'2026-01-01'},
  {id:'later-a', name:'later-a', total:78, date:'2026-02-01'},
  {id:'later-b', name:'later-b', total:82, date:'2026-02-01'}
], 2, 160);
assert.ok(needsTwoPeople.length > 0);
assert.ok(needsTwoPeople.every(plan => plan.assignments.every(entry => entry.item.id !== 'too-large')));

const historyMarkup = context.allocationHistoryMarkup([{id:'plan/1', label:'方案 <1>', plannedDate:'2026-11-02', createdAt:'2026-10-30T12:00:00', manifest:{people:[]}, hasArchive:true, assignments:[
  {group:'父<一>', name:'子&一', total:2, uploader:'甲<一>', owner:'来源甲', parentOnly:false},
  {group:'独立', name:'独立', total:1, uploader:'乙', owner:'', parentOnly:true}
]}]);
assert.ok(historyMarkup.includes('方案 &lt;1&gt;'));
assert.ok(historyMarkup.includes('甲&lt;一&gt;'));
assert.ok(historyMarkup.includes('父&lt;一&gt; / 子&amp;一'));
assert.ok(historyMarkup.includes('3 个文件 · 2 个文件夹 · 2 人'));
assert.ok(historyMarkup.includes('data-copy-history="plan/1"'));
assert.ok(historyMarkup.includes('/api/allocation-history/plan%2F1/archive'));
assert.ok(!historyMarkup.includes('父<一>'));
const jsonOnlyMarkup = context.allocationHistoryMarkup([{id:'new', label:'方案', plannedDate:'2026-11-02', createdAt:'2026-10-30T12:00:00', manifest:{people:[]}, hasArchive:false, assignments:[]}]);
assert.ok(jsonOnlyMarkup.includes('data-copy-history="new"'));
assert.ok(!jsonOnlyMarkup.includes('/archive'));
assert.ok(context.allocationHistoryMarkup([]).includes('暂无已保存的方案'));

context.isComplete = item => item.complete;
context.items = [
  {id: 'first', group: 'A', owner: '甲', total: 40, processed:40, date:'2026-10-30', uploaded: 0, complete: true},
  {id: 'second', group: 'B', owner: '乙', total: 50, processed:50, date:'2026-10-30', uploaded: 0, complete: true},
  {id: 'incomplete', group: 'A', owner: '甲', total: 20, uploaded: 0, complete: false},
  {id: 'uploaded', group: 'B', owner: '乙', total: 20, uploaded: 20, complete: true},
  {id: 'partial', group: 'B', owner: '乙', total: 20, uploaded: 10, complete: true},
  {id: 'empty', group: 'A', owner: '甲', total: 0, uploaded: 0, complete: true},
  {id: 'planned', group: 'A', owner: '甲', total: 20, processed:20, date:'2026-10-30', uploaded:0, complete:true, plannedDate:'2026-11-02'}
];
context.currentGroup = 'A';
assert.deepEqual(Array.from(context.plannerCandidates([]), item => item.id), ['first', 'second']);
assert.deepEqual(Array.from(context.plannerCandidates(['乙']), item => item.id), ['second']);
assert.deepEqual(Array.from(context.plannerCandidates(['甲', '乙']), item => item.id), ['first', 'second']);
const snapshot = context.plannerSnapshot();
context.currentGroup = 'B';
assert.equal(context.plannerSnapshot(), snapshot);
context.items[1].uploaded = 50;
assert.notEqual(context.plannerSnapshot(), snapshot);

context.items = [
  {id:'one', group:'A', name:'one', owner:'甲', total:40, processed:40, date:'2026-10-30', uploaded:0, complete:true},
  {id:'two', group:'A', name:'two', owner:'甲', total:50, processed:50, date:'2026-10-30', uploaded:0, complete:true}
];
elementFor('#plan-target').value = '90';
elementFor('#plan-date').value = '2026-11-02';
const renderAllocation = context.renderAllocation;
context.renderAllocation = () => {};
elementFor('#planner-form').handlers.submit({preventDefault() {}});
assert.equal(elementFor('#planner-results').hidden, false);
assert.equal(vm.runInNewContext('allocationDate', context), '2026-11-02');
elementFor('#plan-date').value = '2026-11-03';
elementFor('#planner-form').handlers.input();
assert.equal(elementFor('#planner-results').hidden, true);
assert.equal(vm.runInNewContext('allocationPlans.length', context), 0);
elementFor('#plan-date').value = '';
elementFor('#planner-form').handlers.submit({preventDefault() {}});
assert.equal(elementFor('#planner-message').textContent, '请选择计划上传日期。');
context.renderAllocation = renderAllocation;

async function testAllocationPreview() {
  context.escapeHTML = value => String(value).replace(/[&<>"']/g, character => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[character]));
  const preview = elementFor('#plan-hover-preview');
  preview.style = {};
  preview.offsetWidth = 200;
  preview.offsetHeight = 100;
  elementFor('#plan-preview-spacer').style = {};
  context.window = {scrollY:300, scrollX:0, innerHeight:720, scrollTo(x, y) { this.scrollY = y; }, scrollBy(x, y) { this.scrollY += y; }};
  elementFor('.planner-result-heading').getBoundingClientRect = () => ({left:10, top:100, width:300});
  context.renderAllocation = () => {};
  context.items.push({owner:'来源甲'});
  vm.runInNewContext(`
    allocationPlans = [
      {totals:[14, 4], assignments:[{person:0, item:{group:'父', name:'一', total:2, owner:'来源甲', date:'2026-01-01'}}, {person:0, item:{group:'父', name:'二', total:2, owner:'甲<一>'}}, {person:0, item:{group:'父', name:'三', total:2, owner:'甲<一>'}}, {person:0, item:{group:'父', name:'四', total:2}}, {person:0, item:{group:'父', name:'五', total:2}}, {person:0, item:{group:'父', name:'六', total:2}}, {person:0, item:{group:'父', name:'七', total:2}}, {person:1, item:{group:'独立', total:4, parentOnly:true}}]},
      {totals:[4, 6], assignments:[{person:0, item:{group:'父', name:'<测试>', total:4}}, {person:1, item:{group:'独立', total:6, parentOnly:true}}]}
    ];
    allocationPeople = ['甲<一>', '乙']; activeAllocation = 0;
  `, context);
  const first = {dataset:{plan:'0'}, textContent:'方案 1', closest() { return this; }, getBoundingClientRect() { return {left:20, top:140, height:35}; }};
  const second = {dataset:{plan:'1'}, textContent:'方案 2', closest() { return this; }, getBoundingClientRect() { return {left:300, top:180, height:35}; }};

  context.showAllocationPreview(first);
  assert.equal(preview.hidden, false);
  assert.ok(preview.innerHTML.includes('甲&lt;一&gt;'));
  assert.ok(preview.innerHTML.includes('14 个文件 · 7 个文件夹'));
  assert.ok(preview.innerHTML.includes('来源：'));
  assert.ok(preview.innerHTML.includes('来源甲'));
  assert.ok(preview.innerHTML.includes('开始：2026-01-01'));
  const markers = context.allocationPersonMarkers();
  const samePerson = context.allocationPersonMarker('甲<一>', markers);
  assert.equal(preview.innerHTML.split(samePerson).length - 1, 3);
  assert.notEqual(markers.get('甲<一>').number, markers.get('来源甲').number);
  assert.notEqual(markers.get('甲<一>').hue, markers.get('来源甲').hue);
  assert.ok(preview.innerHTML.includes('plan-person-unassigned'));
  context.items.push(...['甲', '丙', '丁', '戊', '己', '庚', '辛'].map(owner => ({owner})));
  const manyMarkers = context.allocationPersonMarkers();
  assert.equal(new Set([...manyMarkers.values()].map(marker => marker.number)).size, manyMarkers.size);
  context.items.reverse();
  const reordered = context.allocationPersonMarkers();
  for (const [name, marker] of manyMarkers) assert.equal(reordered.get(name).number, marker.number);
  context.items.reverse();
  vm.runInNewContext('allocationPeople.reverse()', context);
  const reversedPeople = context.allocationPersonMarkers();
  for (const [name, marker] of manyMarkers) {
    assert.equal(reversedPeople.get(name).number, marker.number);
    assert.equal(reversedPeople.get(name).hue, marker.hue);
  }
  vm.runInNewContext('allocationPeople.reverse()', context);
  assert.ok(preview.innerHTML.includes('>一</strong>'));
  assert.ok(preview.innerHTML.includes('>三</strong>'));
  assert.ok(preview.innerHTML.includes('>七</strong>'));
  assert.ok(preview.innerHTML.includes('独立'));
  assert.ok(preview.innerHTML.includes('8 个文件夹 · 3 个来源负责人 · 2 个父文件夹'));
  assert.equal(preview.style.top, '40px');
  assert.equal(preview.style.left, '10px');
  preview.offsetHeight = 1000;
  context.showAllocationPreview(first);
  assert.ok(preview.innerHTML.includes('>七</strong>'));
  assert.equal((preview.innerHTML.match(/class="plan-hover-folder"/g) || []).length, 8);
  assert.ok(!preview.innerHTML.includes('另有'));
  assert.equal(elementFor('#plan-preview-spacer').style.height, '568px');
  context.hideAllocationPreview();
  assert.equal(preview.hidden, true);
  assert.equal(elementFor('#plan-preview-spacer').style.height, '0px');
  assert.equal(context.window.scrollY, 300);
  preview.offsetHeight = 100;
  context.showAllocationPreview(second);
  assert.equal(elementFor('#plan-preview-spacer').style.height, '0px');
  assert.equal(vm.runInNewContext('activeAllocation', context), 0);
  assert.ok(preview.innerHTML.includes('方案 2'));
  assert.ok(preview.innerHTML.includes('&lt;测试&gt;'));
  assert.ok(!preview.innerHTML.includes('<img'));
  assert.equal(preview.style.top, '80px');
  assert.equal(preview.style.left, '100px');

  elementFor('#plan-options').handlers.click({target:second});
  assert.equal(vm.runInNewContext('activeAllocation', context), 1);
  context.resetAllocationPreviews();
  assert.equal(preview.hidden, true);
}

async function testAllocationJson() {
  let copied;
  let clipboardFails = false;
  let confirmationMessage;
  let confirmedPayload;
  context.navigator = {clipboard:{writeText: async text => { if (clipboardFails) throw new Error('denied'); copied = text; }}};
  context.askConfirm = async message => { confirmationMessage = message; return true; };
  context.request = async (url, options) => { confirmedPayload = JSON.parse(options.body); return {items:[]}; };
  context.document.body = {append() {}};
  context.document.createElement = () => ({setAttribute() {}, style:{}, select() {}, remove() {}});
  context.document.execCommand = () => false;
  context.clearPlanPreview = () => {};
  context.render = () => {};
  context.notify = () => {};
  context.items = [
    {id:'json', group:'父', name:'子', total:2, uploaded:0, date:'2026-01-01', owner:'来源甲', complete:true},
    {id:'solo', group:'独立', name:'独立', total:3, uploaded:0, date:'', owner:'', parentOnly:true, complete:true}
  ];
  elementFor('#plan-date').value = '2026-11-02';
  const selectPlan = () => vm.runInNewContext(`
    allocationPlans = [{totals:[2, 3], assignments:[{item:items[0], person:0}, {item:items[1], person:1}]}];
    allocationPeople = ['甲', '乙'];
    activeAllocation = 0;
    allocationSnapshot = plannerSnapshot();
    allocationDate = '2026-11-02';
  `, context);
  selectPlan();

  await elementFor('#plan-create').handlers.click();
  assert.ok(confirmationMessage.includes('2026-11-02'));
  assert.ok(confirmationMessage.includes('不会标记为已上传'));
  assert.ok(confirmationMessage.includes('剪贴板'));
  assert.equal(confirmedPayload.assignments[0].uploadDate, '2026-11-02');
  assert.deepEqual(confirmedPayload.assignments.map(entry => entry.uploader), ['甲', '乙']);
  assert.equal(confirmedPayload.archive, undefined);
  const expected = {version:2, label:'方案 1 · 5 个', uploadDate:'2026-11-02', people:[
    {name:'甲', folders:[{source:['父', '子'], destination:['甲', '父', '子'], owner:'来源甲', startDate:'2026-01-01', files:2, fileList:[]}]},
    {name:'乙', folders:[{source:['独立'], destination:['乙', '独立'], owner:'', startDate:'', files:3, fileList:[]}]}
  ]};
  assert.deepEqual(confirmedPayload.manifest, expected);
  assert.deepEqual(JSON.parse(copied), expected);

  const copyHandler = context.document.handlers.click.at(-1);
  const dateCopyHandler = context.document.handlers.click.at(-2);
  context.planPreview = null;
  context.icons = () => {};
  context.updateDateExport('2026-10-29', uploadedRecords);
  assert.equal(elementFor('#date-history-bar').hidden, false);
  assert.ok(elementFor('#date-history-bar').innerHTML.includes('data-copy-date-json'));
  copied = undefined;
  await dateCopyHandler({target:{closest: () => ({})}});
  assert.deepEqual(JSON.parse(copied), dateJson);
  context.updateDateExport('', uploadedRecords);
  assert.equal(elementFor('#date-history-bar').hidden, true);
  context.updateDateExport('2026-10-29', uploadedRecords);
  context.planPreview = {};
  context.updateDateExport('2026-10-29', uploadedRecords);
  assert.equal(elementFor('#date-history-bar').hidden, true);
  context.planPreview = null;
  context.rememberHistory([{id:'history-1', manifest:{people:[], note:'历史'}}, {id:'legacy', manifest:null}]);
  copied = undefined;
  await copyHandler({target:{closest: () => ({dataset:{copyHistory:'history-1'}})}});
  assert.deepEqual(JSON.parse(copied), {people:[], note:'历史'});
  copied = undefined;
  await copyHandler({target:{closest: () => ({dataset:{copyHistory:'legacy'}})}});
  assert.equal(copied, undefined);

  let historyRefreshes = 0;
  const loadHistory = context.loadAllocationHistory;
  context.loadAllocationHistory = () => { historyRefreshes++; };
  elementFor('#allocation-history').open = true;
  clipboardFails = true;
  selectPlan();
  await elementFor('#plan-create').handlers.click();
  assert.equal(historyRefreshes, 1, '复制失败时已保存的历史也应刷新');
  assert.ok(elementFor('#planner-message').textContent.includes('已保存上传计划，但复制 JSON 失败'));
  elementFor('#allocation-history').open = false;
  context.loadAllocationHistory = loadHistory;
  clipboardFails = false;

  copied = undefined;
  context.request = async () => { throw new Error('conflict'); };
  selectPlan();
  await elementFor('#plan-create').handlers.click();
  assert.equal(copied, undefined, '确认失败时不能复制方案');
  assert.ok(elementFor('#planner-message').textContent.includes('未保存上传计划'));
  const manifest = [{source:['父', '子'], destination:['甲', '父', '子']}];
  const firstScript = fs.readFileSync(path.join(__dirname, '../data/step1_copy_files.py'), 'utf8');
  const secondScript = fs.readFileSync(path.join(__dirname, '../data/step2_copy_onedrive.py'), 'utf8');

  const os = require('node:os');
  const {spawnSync} = require('node:child_process');
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'upload-desk-'));
  try {
    const source = path.join(temp, '资源库 [10% &!]');
    const extracted = path.join(temp, '解压');
    const cloud = path.join(temp, '同步 [20% &!]');
    const documents = path.join(extracted, '分配资料');
    const sourceFolder = path.join(source, '父', '子');
    fs.mkdirSync(path.join(sourceFolder, '[目录]'), {recursive:true});
    fs.mkdirSync(path.join(sourceFolder, '空目录'));
    fs.mkdirSync(documents, {recursive:true});
    fs.mkdirSync(cloud);
    fs.writeFileSync(path.join(sourceFolder, '[目录]', '文件%&!.txt'), 'original');
    fs.writeFileSync(path.join(documents, 'allocation.json'), JSON.stringify(manifest));
    const firstPath = path.join(documents, 'step1_copy_files.py');
    const secondPath = path.join(documents, 'step2_copy_onedrive.py');
    fs.writeFileSync(firstPath, firstScript.replace('r"EDIT_SOURCE_ROOT"', JSON.stringify(source)));
    fs.writeFileSync(secondPath, secondScript.replace('r"EDIT_ONEDRIVE_SYNC_FOLDER"', JSON.stringify(cloud)));
    const runScript = script => spawnSync('python', [script], {encoding:'utf8', timeout:30000, env:{...process.env, PYTHONIOENCODING:'utf-8'}});
    let result = runScript(firstPath);
    assert.equal(result.status, 0, `第一步复制失败：${result.stdout} ${result.stderr}`);
    const relative = path.join('甲', '父', '子', '[目录]', '文件%&!.txt');
    assert.equal(fs.readFileSync(path.join(extracted, relative), 'utf8'), 'original');
    assert.equal(fs.existsSync(path.join(cloud, relative)), false);
    result = runScript(secondPath);
    assert.equal(result.status, 0, `第二步复制失败：${result.stdout} ${result.stderr}`);
    assert.equal(fs.readFileSync(path.join(cloud, relative), 'utf8'), 'original');
    assert.ok(fs.statSync(path.join(cloud, '甲', '父', '子', '空目录')).isDirectory());
    assert.equal(runScript(firstPath).status, 0, '第一步重复执行应成功');
    fs.writeFileSync(path.join(cloud, relative), 'different');
    result = runScript(secondPath);
    assert.equal(result.status, 0, `覆盖失败：${result.stdout} ${result.stderr}`);
    assert.equal(fs.readFileSync(path.join(cloud, relative), 'utf8'), 'original');
    const log = fs.readFileSync(path.join(documents, 'onedrive-overwrite-log.jsonl'), 'utf8').trim().split('\n').map(line => JSON.parse(line));
    assert.deepEqual(log.map(entry => entry.status), ['pending', 'completed']);
    fs.writeFileSync(path.join(sourceFolder, '[目录]', '文件%&!.txt'), 'different');
    result = runScript(firstPath);
    assert.notEqual(result.status, 0, '第一步不同内容的同名文件必须报错');
    assert.equal(fs.readFileSync(path.join(extracted, relative), 'utf8'), 'original');
    fs.writeFileSync(path.join(sourceFolder, '[目录]', '文件%&!.txt'), 'original');
    fs.mkdirSync(path.join(source, '父', '其他'));
    fs.writeFileSync(path.join(source, '父', '其他', '第二人.txt'), 'second');
    fs.writeFileSync(path.join(documents, 'allocation.json'), JSON.stringify([
      ...manifest, {source:['父', '其他'], destination:['乙', '父', '其他']}
    ]));
    assert.equal(runScript(firstPath).status, 0, '多人员第一步复制应成功');
    assert.equal(runScript(secondPath).status, 0, '多人员第二步复制应成功');
    assert.equal(fs.readFileSync(path.join(cloud, '乙', '父', '其他', '第二人.txt'), 'utf8'), 'second');
    fs.writeFileSync(path.join(documents, 'allocation.json'), JSON.stringify(manifest));
    fs.renameSync(sourceFolder, path.join(source, '父', 'moved'));
    assert.notEqual(runScript(firstPath).status, 0, '缺少源目录必须报错');
    fs.renameSync(path.join(source, '父', 'moved'), sourceFolder);
    fs.writeFileSync(firstPath, firstScript.replace('r"EDIT_SOURCE_ROOT"', JSON.stringify(extracted)));
    assert.notEqual(runScript(firstPath).status, 0, '源目录和目标目录重叠必须报错');
    fs.writeFileSync(firstPath, firstScript.replace('r"EDIT_SOURCE_ROOT"', JSON.stringify(source)));
    fs.writeFileSync(path.join(documents, 'allocation.json'), JSON.stringify([
      {source:['父'], destination:['甲', '父']}, {source:['父', '子'], destination:['乙', '父', '子']}
    ]));
    assert.notEqual(runScript(firstPath).status, 0, '分配的源目录互相包含时必须报错');
  } finally {
    fs.rmSync(temp, {recursive:true, force:true});
  }
}
testAllocationPreview().then(testAllocationJson).then(() => console.log('方案分配、悬停预览和方案 JSON 复制测试通过')).catch(error => {
  console.error(error);
  process.exitCode = 1;
});

````

### 5.20 `tests\test_tree.js`

<!-- REBUILD_FILE {"path":"tests\\test_tree.js","newline":"LF","encoding":"utf-8","size":6191,"sha256":"9d98c352862e0feecf84624b16f876b942cca338ac93a26ba6c3778f758544bc"} -->
````javascript
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../static/js/app.js'), 'utf8');
const context = {
  collapsedFolders: new Set(),
  currentOwners: new Set(),
  escapeHTML: value => String(value).replace(/[&<>"']/g, character => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'}[character])),
  isComplete: item => item.complete,
  percentOf: () => 0
};
vm.runInNewContext(source.slice(source.indexOf('const folderStatuses'), source.indexOf('function percentOf(')), context);
vm.runInNewContext(source.slice(source.indexOf('function folderTree('), source.indexOf('function render()')), context);

function item(id, group, name, parentOnly = false) {
  return {id, group, name, parentOnly, total: 1, uploaded: 0, complete: true, date: '', owner: ''};
}
const records = [
  item('parent', '根', '父'),
  item('child', '根/父', '子'),
  item('parent-only', '独立\\目录', '目录', true),
  item('separator', '///', '特殊', true),
  item('escaped', '根', '<测试>')
];
const base = {total:10, processed:0, date:'', uploaded:0, complete:false};
assert.equal(context.statusOf(base), 'unstarted');
assert.equal(context.statusOf({...base, date:'2026-10-30'}), 'processing');
assert.equal(context.statusOf({...base, date:'2026-10-30', processed:10}), 'ready');
assert.equal(context.statusOf({...base, plannedDate:'2026-11-02'}), 'uploading');
assert.equal(context.statusOf({...base, uploaded:10}), 'uploaded');
const manuallyReady = context.withFolderStatus(base, 'ready');
assert.equal(context.statusOf(manuallyReady), 'ready');
assert.equal(manuallyReady.processed, 0);
assert.equal(manuallyReady.date, '');
assert.equal(context.withFolderStatus({...base, plannedDate:'2026-11-02', uploader:'甲'}, 'uploaded').uploaded, 10);
assert.equal(context.withFolderStatus({...base, plannedDate:'2026-11-02', uploader:'甲'}, 'unstarted').plannedDate, '');
const cycle = [];
for (let item = {...base, total:10}, index = 0; index < 6; index++) {
  cycle.push(context.statusOf(item));
  item = context.withFolderStatus(item, context.nextStatus(item));
}
assert.deepEqual(cycle, ['unstarted', 'processing', 'ready', 'uploading', 'uploaded', 'unstarted']);
assert.equal(context.nextStatus({...base, statusOverride:'uploading', total:0, totalMissing:true}), 'unstarted');
const tree = context.folderTree(records);
const parent = tree.children.get('根').children.get('父');
assert.equal(parent.item.id, 'parent');
assert.equal(parent.children.get('子').item.id, 'child');
assert.equal(tree.children.get('独立').children.get('目录').item.id, 'parent-only');
assert.equal(tree.children.get('///').item.id, 'separator');

const expanded = context.renderFolderTree(tree);
assert.equal((expanded.match(/data-action="edit"/g) || []).length, records.length);
assert.ok(expanded.includes('&lt;测试&gt;'));
assert.ok(expanded.includes('data-status-id="parent"'));
assert.ok(expanded.includes('class="folder-status unstarted"'));
assert.ok(expanded.includes('点击切换为完成中...'));
assert.ok(expanded.includes('class="processed-count">—</span>'));
records[1].processed = 0;
records[1].uploader = '名义<人员>';
records[1].uploadDate = '2026-10-29';
assert.ok(context.renderFolderTree(tree).includes('class="processed-count">0</span>'));
assert.ok(context.renderFolderTree(tree).includes('<td>名义&lt;人员&gt;</td>'));
assert.ok(context.renderFolderTree(tree).includes('<td>2026 / 10 / 29</td>'));
assert.ok(context.renderFolderTree(tree).includes('colspan="10"'));
assert.equal(context.matchesUploadDate(records[1], '2026-10-29'), true);
assert.equal(context.matchesUploadDate(records[1], '2026-10-30'), false);
assert.equal(context.matchesUploadDate(records[0], '2026-10-29'), false);
assert.equal(context.matchesUploadDate(records[0], ''), true);
const preview = context.renderFolderTree(tree, new Map([['child', '甲<乙>']]));
assert.ok(preview.includes('分配给 甲&lt;乙&gt;'));
assert.equal((preview.match(/class="tree-assignment"/g) || []).length, 1);
assert.ok(expanded.includes('class="owner-tag owner-cell-tag" data-owner="" aria-pressed="false"'));
context.currentOwners.add('');
assert.ok(context.renderFolderTree(tree).includes('class="owner-tag owner-cell-tag" data-owner="" aria-pressed="true"'));
context.currentOwners.clear();
context.collapsedFolders.add(parent.path);
const collapsed = context.renderFolderTree(tree);
assert.ok(collapsed.includes('data-id="parent"'));
assert.ok(!collapsed.includes('data-id="child"'));
assert.ok(collapsed.includes('aria-expanded="false"'));
context.collapsedFolders.clear();
assert.ok(context.renderFolderTree(tree).includes('data-id="child"'));
context.setFoldersExpanded(tree, false);
assert.ok(context.collapsedFolders.has(tree.children.get('根').path));
assert.ok(context.collapsedFolders.has(parent.path));
assert.ok(!context.renderFolderTree(tree).includes('data-id="child"'));
context.collapsedFolders.delete(tree.children.get('根').path);
assert.ok(!context.renderFolderTree(tree).includes('data-id="child"'));
context.setFoldersExpanded(tree, true);
assert.ok(context.renderFolderTree(tree).includes('data-id="child"'));
const choices = context.folderChoices(records, '根/父', 'child');
assert.equal(choices.find(choice => choice.name === '子').unavailable, false);
assert.equal(context.folderChoices(records, '根/父', null).find(choice => choice.name === '子').unavailable, true);
assert.equal(context.folderChoices(records, '根', 'parent').find(choice => choice.name === '父').unavailable, false);
assert.equal(context.folderChoices(records, '独立\\目录', null)[0].unavailable, true);
assert.equal(context.folderChoices(records, '独立\\目录', 'parent-only')[0].unavailable, false);
assert.equal(context.folderChoices(records, '根/父', 'child').some(choice => choice.name === '父'), false);
assert.equal(context.folderChoices(records, '根/父', 'child').some(choice => choice.name === ''), false);
assert.equal(context.folderChoices(records, '不存在', 'child').length, 0);

console.log('目录树和展开折叠测试通过');

````

### 5.21 `data\設計書の履歴シート削除一覧.xlsx`

<!-- REBUILD_FILE {"path":"data\\設計書の履歴シート削除一覧.xlsx","encoding":"base64","size":16967,"sha256":"859c351ed078c9765e22a4a3c10bb2ef73849d877b254c362429a6c59fba1bbc"} -->
````base64
UEsDBBQABgAIAAAAIQBgknq+gAEAAAwGAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAAC
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADE
VM1KAzEQvgu+w5Kr7Kb2ICLdevDnqELrA6TJdDc0m4TMVNu3dzbaIlJbigUvWbLJfD+T5BvdrjpX
vEFCG3wtLquBKMDrYKxvavE6fSyvRYGkvFEueKjFGlDcjs/PRtN1BCy42mMtWqJ4IyXqFjqFVYjg
eWUeUqeIp6mRUemFakAOB4MrqYMn8FRSjyHGo3uYq6Wj4mHFvz+VzKwXxd3nvp6qFipGZ7UiFirf
vPlBUob53GowQS87hq4wJlAGWwDqXBWTZcY0ASI2hkLu5Ezg8DjSL1cVV2Zh2NqIF2z9F4Z+5XdX
X3XPfBzJGiheVKIn1bF3uXLyPaTFLIRFtR/k2NbkFlWdsn6jew9/3owyfy5PLKT3l4GP1DH8Jx3E
dx1kHv/eigxzwDjS2gGe+vgz6CHmViUwE+JX1JxcwHfsAzq0cvqu5at64iZscffxc7S8pBCR0yvB
8QI2UdFXl5GBIJGFbVjsenRbRo6+PzuGPlsNmB3cMmf5+AMAAP//AwBQSwMEFAAGAAgAAAAhALVV
MCP0AAAATAIAAAsACAJfcmVscy8ucmVscyCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACskk1PwzAMhu9I/IfI99XdkBBCS3dBSLsh
VH6ASdwPtY2jJBvdvyccEFQagwNHf71+/Mrb3TyN6sgh9uI0rIsSFDsjtnethpf6cXUHKiZylkZx
rOHEEXbV9dX2mUdKeSh2vY8qq7iooUvJ3yNG0/FEsRDPLlcaCROlHIYWPZmBWsZNWd5i+K4B1UJT
7a2GsLc3oOqTz5t/15am6Q0/iDlM7NKZFchzYmfZrnzIbCH1+RpVU2g5abBinnI6InlfZGzA80Sb
vxP9fC1OnMhSIjQS+DLPR8cloPV/WrQ08cudecQ3CcOryPDJgosfqN4BAAD//wMAUEsDBBQABgAI
AAAAIQC35zZ2MAQAAEAKAAAPAAAAeGwvd29ya2Jvb2sueG1srFbbbuM2EH0v0H9QhbzKIiXqYiHO
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
IQDtVdTIxQcAABUiAAATAAAAAAAAAAAAAAAAAE4gAAB4bC90aGVtZS90aGVtZTEueG1sUEsBAi0A
FAAGAAgAAAAhAPUTkqLcBAAAkRsAAA0AAAAAAAAAAAAAAAAARCgAAHhsL3N0eWxlcy54bWxQSwEC
LQAUAAYACAAAACEAVlKU7o4HAAB/OQAAFAAAAAAAAAAAAAAAAABLLQAAeGwvc2hhcmVkU3RyaW5n
cy54bWxQSwECLQAUAAYACAAAACEAO20yS8EAAABCAQAAIwAAAAAAAAAAAAAAAAALNQAAeGwvd29y
a3NoZWV0cy9fcmVscy9zaGVldDEueG1sLnJlbHNQSwECLQAUAAYACAAAACEAB0M5VbEAAAA0BAAA
JwAAAAAAAAAAAAAAAAANNgAAeGwvcHJpbnRlclNldHRpbmdzL3ByaW50ZXJTZXR0aW5nczEuYmlu
UEsBAi0AFAAGAAgAAAAhAGFAiFBpAQAAowYAABAAAAAAAAAAAAAAAAAAAzcAAHhsL2NhbGNDaGFp
bi54bWxQSwECLQAUAAYACAAAACEAYLRKgrEBAAAnAwAAEQAAAAAAAAAAAAAAAACaOAAAZG9jUHJv
cHMvY29yZS54bWxQSwECLQAUAAYACAAAACEAhDIpAc8BAAB2AwAAEAAAAAAAAAAAAAAAAACCOwAA
ZG9jUHJvcHMvYXBwLnhtbFBLBQYAAAAADgAOAKoDAACHPgAAAAA=
````

### 5.22 `data\設計書の履歴シート削除一覧-for-custom.xlsx`

<!-- REBUILD_FILE {"path":"data\\設計書の履歴シート削除一覧-for-custom.xlsx","encoding":"base64","size":16671,"sha256":"20187fb3e0c0059b759be049505387fdb91b1b288bf3e1721985fca62a51ad47"} -->
````base64
UEsDBBQABgAIAAAAIQBgknq+gAEAAAwGAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAAC
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADE
VM1KAzEQvgu+w5Kr7Kb2ICLdevDnqELrA6TJdDc0m4TMVNu3dzbaIlJbigUvWbLJfD+T5BvdrjpX
vEFCG3wtLquBKMDrYKxvavE6fSyvRYGkvFEueKjFGlDcjs/PRtN1BCy42mMtWqJ4IyXqFjqFVYjg
eWUeUqeIp6mRUemFakAOB4MrqYMn8FRSjyHGo3uYq6Wj4mHFvz+VzKwXxd3nvp6qFipGZ7UiFirf
vPlBUob53GowQS87hq4wJlAGWwDqXBWTZcY0ASI2hkLu5Ezg8DjSL1cVV2Zh2NqIF2z9F4Z+5XdX
X3XPfBzJGiheVKIn1bF3uXLyPaTFLIRFtR/k2NbkFlWdsn6jew9/3owyfy5PLKT3l4GP1DH8Jx3E
dx1kHv/eigxzwDjS2gGe+vgz6CHmViUwE+JX1JxcwHfsAzq0cvqu5at64iZscffxc7S8pBCR0yvB
8QI2UdFXl5GBIJGFbVjsenRbRo6+PzuGPlsNmB3cMmf5+AMAAP//AwBQSwMEFAAGAAgAAAAhALVV
MCP0AAAATAIAAAsACAJfcmVscy8ucmVscyCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACskk1PwzAMhu9I/IfI99XdkBBCS3dBSLsh
VH6ASdwPtY2jJBvdvyccEFQagwNHf71+/Mrb3TyN6sgh9uI0rIsSFDsjtnethpf6cXUHKiZylkZx
rOHEEXbV9dX2mUdKeSh2vY8qq7iooUvJ3yNG0/FEsRDPLlcaCROlHIYWPZmBWsZNWd5i+K4B1UJT
7a2GsLc3oOqTz5t/15am6Q0/iDlM7NKZFchzYmfZrnzIbCH1+RpVU2g5abBinnI6InlfZGzA80Sb
vxP9fC1OnMhSIjQS+DLPR8cloPV/WrQ08cudecQ3CcOryPDJgosfqN4BAAD//wMAUEsDBBQABgAI
AAAAIQDsvTI9JAQAADAKAAAPAAAAeGwvd29ya2Jvb2sueG1srFbbbuM2EH0v0H9QhbzKIiXqYiHO
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
cC54bWxQSwUGAAAAAA4ADgCqAwAAXz0AAAAA
````

### 5.23 可选数据库

默认不包含数据库，首次启动将创建空数据库。如果需要把现有工作数据也放入唯一的 Markdown，请先阅读第 8 节；加入后该位置将包含数据库备份的 Base64 数据块。

<!-- OPTIONAL_DATABASE_NOT_INCLUDED -->

本文件 `REBUILD.md` 会由还原脚本原样复制到新项目；无需递归嵌入自身内容。

## 6. 安装与运行

以下命令在新项目根目录的 PowerShell 中执行；每条成功后再继续，失败时检查错误输出。

### 6.1 手动安装（推荐用于验证复刻）

```powershell
python --version
node --version
npm --version
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r .\requirements.txt
npm ci
```

若要匹配原环境的 openpyxl 版本，在安装后执行下面的命令，先把占位替换成第 1 节记录的真实版本；不追求一致版本时跳过：

```powershell
.\.venv\Scripts\python.exe -m pip install "openpyxl==【原环境实际版本】"
```

`npm ci` 需要完整且与 package.json 一致的锁文件。失败时不要删除锁文件改用其他版本，应先检查占位是否补全、锁文件是否复制完整、网络和 npm 版本是否符合原环境。

### 6.2 启动服务

```powershell
.\.venv\Scripts\python.exe .\server.py
```

在浏览器打开 <http://127.0.0.1:8000>。保持窗口打开，使用 Ctrl+C 停止服务。

另一个 PowerShell 窗口可检查健康接口：

```powershell
Invoke-RestMethod -Uri 'http://127.0.0.1:8000/api/health'
```

响应的 `service` 字段应为 `upload-desk`。

### 6.3 验证启动器

先停止手动启动的服务，再执行：

```powershell
.\.venv\Scripts\python.exe .\start.py
```

应打开浏览器并保持服务运行。双击 `start.py` 使用 Windows 的 Python 文件关联，不一定使用 `.venv`，会检查并可能安装关联解释器所需的依赖。

### 6.4 可选环境变量

| 变量 | 默认值 | 作用 |
| --- | --- | --- |
| `PORT` | `8000` | 本机监听端口 |
| `UPLOAD_DESK_DB` | 项目内 `data\upload_data.sqlite3` | 指定数据库文件路径 |
| `SOURCE_ROOT` | 未配置 | 旧版相关逻辑读取源目录文件列表的根目录；新方案 JSON 复制不要求它 |

例如，在同一个 PowerShell 窗口设置端口后启动：

```powershell
$env:PORT = '8001'
.\.venv\Scripts\python.exe .\start.py
```

此时访问 <http://127.0.0.1:8001>。环境变量只作用于该进程环境，不是自动保存的项目配置。

## 7. 验证是否复刻成功

### 自动化测试

在项目根目录运行：

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
node .\tests\test_planner.js
node .\tests\test_tree.js
```

三个命令都应成功。不要删除失败的测试来获得“通过”；记录完整错误、定位原因，并区分原材料缺失、环境差异和恢复错误。

### 浏览器验收清单

- [ ] 未内嵌数据库时，全新启动显示空数据库，不把旧 `file://` 种子数据当作服务端现有数据。
- [ ] Excel 导入先显示预览，确认后出现父文件夹树、明细和统计。
- [ ] 搜索、负责人、状态、上传日期筛选和排序正常。
- [ ] 新增、编辑、删除及刷新后持久化正常。
- [ ] 完成数量留空和 0 的区别保留，状态切换及取消上传确认正常。
- [ ] 在新建验收数据中准备总数 10、已完成未上传的一个文件夹，选择一人、目标 10，生成的方案总数在 5 至 15 之间。
- [ ] 验收方案不拆分文件夹，各人都有文件夹；确认后进入上传中而非已上传。
- [ ] 剪贴板 JSON 可以解析，字段和路径数组符合第 3 节约定；无剪贴板权限时明确提示。
- [ ] 历史方案可查看、重新复制，后续修改文件夹不改变历史快照。
- [ ] Excel、客户表及 JSON 导出成功；在独立空数据库中检查重新导入。
- [ ] 服务停止重启后，已保存的记录仍存在。
- [ ] 对照原项目截图检查字体、图标、布局、颜色、弹窗和方案预览。
- [ ] 如恢复现有数据库，恢复后的记录数量和历史方案与备份前一致，不使用真实数据进行破坏性验收。

本文没有包含界面截图；如果需要精确对照，可另行将截图编码嵌入同一份 Markdown。

## 8. 恢复当前数据与完整性检查

### 可选：把数据库也放进唯一的 Markdown

默认的本文**不含当前 SQLite 工作数据**。如果只复刻程序，跳过本节。

如果也需要恢复文件夹记录、计划及历史方案：

1. 确认数据库内容可以交给接收方。数据库可能包含敏感业务数据；Base64 无法加密或脱敏。
2. 在原项目中停止服务及其他数据库写入进程。不要仅复制运行中的 `.sqlite3` 主文件；下方使用 SQLite 备份 API 制作一致性快照。
3. 在原项目根目录执行以下 PowerShell 命令。默认备份 `data\upload_data.sqlite3`；若使用 `UPLOAD_DESK_DB`，先将第一行改为实际数据库路径。命令不会覆盖现有的 `REBUILD_WITH_DATA.md`。

<!-- REBUILD_DATABASE_PACKAGER -->
```powershell
$env:REBUILD_SOURCE_DB = (Resolve-Path -LiteralPath '.\data\upload_data.sqlite3').Path
@'
import base64
import hashlib
import json
import os
import sqlite3
import tempfile
import textwrap
from contextlib import closing
from pathlib import Path

source_path = Path(os.environ['REBUILD_SOURCE_DB']).resolve(strict=True)
document_path = Path.cwd() / 'REBUILD.md'
output_path = Path.cwd() / 'REBUILD_WITH_DATA.md'
if output_path.exists():
    raise SystemExit('REBUILD_WITH_DATA.md already exists; choose an empty directory.')
marker = '<!-- OPTIONAL_' + 'DATABASE_NOT_INCLUDED -->'
document = document_path.read_text(encoding='utf-8-sig')
if document.count(marker) != 1:
    raise SystemExit('Optional database placeholder missing or duplicated.')

with tempfile.TemporaryDirectory() as directory:
    snapshot_path = Path(directory) / 'snapshot.sqlite3'
    with closing(sqlite3.connect(source_path.as_uri() + '?mode=ro', uri=True)) as source:
        with closing(sqlite3.connect(snapshot_path)) as snapshot:
            source.backup(snapshot)
    data = snapshot_path.read_bytes()

info = {
    'path': r'data\upload_data.sqlite3',
    'encoding': 'base64',
    'size': len(data),
    'sha256': hashlib.sha256(data).hexdigest(),
}
encoded = '\n'.join(textwrap.wrap(base64.b64encode(data).decode('ascii'), 76))
block = '<!-- REBUILD_FILE ' + json.dumps(info) + ' -->\n````base64\n' + encoded + '\n````\n'
with output_path.open('x', encoding='utf-8', newline='\n') as output:
    output.write(document.replace(marker, block, 1))
print('Created REBUILD_WITH_DATA.md with a verified SQLite snapshot.')
'@ | python -
```

4. 只交付生成的 `REBUILD_WITH_DATA.md`，不需要另交数据库文件。接收方将它保存为 `REBUILD.md`，再执行第 4 节还原脚本。
5. 启动后核对记录数量、计划和历史方案。这个数据库只代表执行备份时的状态，不会随原项目后续修改而更新。

JSON 备份主要恢复文件夹明细，不能替代包含历史分配快照的完整 SQLite 备份。原 Excel 也不能恢复当前全部上传记录和历史方案。

### 校验和更新快照

第 4 节脚本会逐个校验 20 个文本文件和两份 Excel 的字节数、SHA-256，以及可选数据库。校验通过才写入；文本文件按记录的 LF 或 CRLF 还原。若文件传输时被截断、编辑器改写了源码，或手动更换内容而未更新元数据，脚本会报错。

SHA-256 用于发现内容变化，不证明文档来源可信。若需要证明来源，另行核对可信渠道提供的文档哈希。

本文是制作时的项目快照，不会自动跟随项目更新。源码或 Excel 修改后，需要让 AI 同步更新第 5 节对应内容、字节数、SHA-256 和文本换行标记；Excel 重新编码为 Base64。然后在新空目录重新执行还原脚本，并将全部还原文件与当前项目逐一比较。不要擅自把正在使用的数据库或依赖目录加入快照。
