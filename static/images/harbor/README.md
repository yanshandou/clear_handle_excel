# 海岛工具箱素材

使用内置 ImageGen，参考用户提供的三张海岛和帆船图片。底板和图标经裁切、缩放后以 WebP 随项目发布。
底板以 360 × 144 存储，使用 66 像素九宫格切片；图标精灵图为 5 × 5，每格 120 像素。

## 底板

Use case: ui-mockup. Input images are style and palette references only. Create a game UI skin sprite sheet in the exact cheerful soft stylized 3D cartoon toy rendering style of the tropical island and sailboat reference images, tropical lagoon and sky colors. Arrange nine identical-size very wide 2.6:1 blank front-on rounded rectangles in a precisely aligned three column, three row grid, all with real alpha outside, with wide clean transparent gaps. No perspective, no objects, no text, no scenery. Each entire rectangle is a smooth molded shell/enamel surface with a restrained shallow softly lit rim about 2% of its height. Broad nearly uniform interiors so live text is readable and texture stretches well. Row 1: (1) seafoam-white neutral button with medium aqua rim, (2) vivid saturated ocean cobalt-blue primary button, center dark enough for white lettering around #075db0, (3) rich tropical coral-red danger button, center dark enough for white lettering around #ad322b. Row 2: (4) cloud-white near-white light aqua panel and subtle frost blue rim, (5) light clear lagoon mint panel, (6) light sky-periwinkle panel. Row 3: (7) light sky blue panel, (8) pale sand-yellow panel with sunny warm rim, (9) light tropical coral pink panel. Smooth clean rounded toy edges, no paper texture or vintage materials, no thick cast shadows, no shadows outside the silhouettes. Transparent background. Landscape canvas.

## 图标

Use case: ui-mockup. Input images are reference for style and palette, not layout. Generate 25 unique game inventory icons in the reference's soft bright stylized 3D tropical island cartoon toy rendering: ocean blue, turquoise lagoon, foamy white, tropical lime, sandy gold, touches of coral; soft sunny highlights, glossy ceramic and rounded molded parts. A perfectly aligned uniform 5 columns × 5 rows game sprite atlas, one compact icon centered per square cell, each icon takes 70–77% of its cell with substantial clear gutters on all sides, never touching any cell boundary. Exact positions LEFT TO RIGHT by row: 1 seaside cabin with bright blue roof, green spreadsheet grid, cobalt magnifier, colorful rising bar chart, yellow-orange folder. 2 stack of colored documents, down arrow into blue tray, up arrow out of blue tray, turquoise gear, two joined gold chain links. 3 open stylized eye, artist paint palette, white-blue-coral bullseye target, paper document with coral folded corner, blue and lime typography blocks. 4 framed coastal picture, turquoise clipboard with one bold checkmark, spreadsheet stack and toy cleaning brush, sheet with yellow set square and ruler, magnifier over spreadsheet with highlighted orange cell. 5 bright blue treasure shipping crate with lime up arrow, open blue and gold treasure chest with a file peeking out, azure pennant trophy, rolled nautical map, blue and green compass rose. Do not add labels, background shapes, scenery, drop shadows, or objects between icons. Transparent true alpha. Crisp simple readability at small size. Square canvas.

## 工具专属图标（2026-11-20）

主页工具箱里一个工具一张图，不再共用精灵图：原来「排查 BL / 修复引用」共用链条、「调整页面 / 修复列宽」共用尺子、
「批量 PDF / PDF 对比」共用纸张、「修复颜色 / 设置图标」共用调色板，另有几个图标和工具功能不贴切。
现 11 个工具各有专属图，存 `tools/<名字>.png`（256 × 256 透明底），CSS 在 `garden-shell.css` 用
`--icon-image / --icon-size / --icon-x / --icon-y` 变量盖掉精灵图（`theme-cartoon.css` 的对话框头同样走变量）。
生成时以本目录 `icons.webp` 作风格参考；出图是白底，按「贴近背景色的连通域」抠掉背景（含图标内部镂空），
裁切留 7% 边距后缩到 256；appicon / pdfstack 底部多画的投影按低饱和度区域单独清掉（精灵图原图无投影）。

对应关系：bl=排查 BL 超链接（断开的金链 + 珊瑚红警示星）、reflink=修复引用错误（金链 + 绿松石扳手）、
widths=修复列宽（表格 + 蓝色横向双向箭头）、pdfstack=批量导出同名 PDF（一叠珊瑚角纸 + 青柠导出箭头）、
pdfdiff=PDF 对比（两张纸 + 蓝色对换箭头 + 红绿差异点）、report=Excel 对比报告（两张表格衬底 + 柱状图报告页）、
json=按 JSON 复制 Excel（绿松石花括号括住表格）、pick=挑文件复制（金文件夹 + 带勾选中的文件）、
view=视图标准化（显示器里表格首格沙金高亮 + 蓝色十字准星）、appicon=设置图标（蓝色圆角应用块 + 齿轮）、
imgfix=修复图片（海岛相框 + 珊瑚红校正框 + 扳手）。

通用提示词（每张单独出图，方形透明底）：Single game inventory icon: <主体描述>. Match exactly the soft
bright stylized 3D tropical-island cartoon toy rendering of the reference sprite sheet: ocean cobalt blue,
turquoise lagoon, foamy white, tropical lime, sandy gold, touches of coral; glossy ceramic rounded molded
parts, soft sunny highlights. One compact icon centered on a fully transparent background (true alpha),
icon fills about 75% of the square canvas, no text, no labels, no background shapes, no scenery, no drop
shadow outside the silhouette, crisp and readable at small size.

## PCL 检查台

Use case: ui-mockup. Input images are visual references for style and palette. One isolated game prop on transparent alpha: front-facing bright blue and turquoise harbor inspection console with rounded toy shapes, cobalt blue shell matching the sailboat, curved creamy white trim, little sand-gold dials on the right, subtle miniature life ring or ship wheel. A large completely blank dark navy screen for live overlay text, left of center (spanning x=15% to 72%, y=25% to 75%). Style must closely match the bright sunny, smooth, stylized 3D tropical island cartoon toy CGI of the references. Gentle soft rim lighting, clean materials, no paper texture, nothing behind it, no floor, no thick cast shadow, no words, no text, no numbers. 4:3 landscape.

## 海下背景 undersea.webp（2026-11-20）

首页横幅（bg.png，海面上）保持不动；水面以下由 `undersea.webp` 竖向无缝平铺补全，
左右留白处露出珊瑚礁，整页读作一张连续的海（参考用户提供的竖向海上海下构图）。

- 出图：1536×1024，风格参考 bg.png、构图参考用户图；两侧礁石群 + 中部开阔水域（气泡、小鱼群）。
  第一版珊瑚顶到上下边缘无法平铺，重生成时强制「上下 18% 全宽纯海水」，第二版通过。
- 后处理：取顶行水色 W=rgb(26,251,253) 为接缝色；顶 10% / 底 8% 强制为纯 W，
  10–17% 与 84–92% 用 smoothstep 混回原图，首末行像素完全相同 → repeat-y 无缝。
- 接法：`.home-stage`（position:relative）包住横幅 + `.home-undersea` + 工作台；
  undersea `inset:0; z-index:0`，`background-position-y: clamp(184px,24vw,316px)`（= 横幅高度）
  使贴图第 0 行正好落在横幅底边；横幅 `::after` 68px 渐隐收尾到同一 W，接缝无色差。
  深度渐变 ::before 前 30% 保持全透明以免污染接缝；::after 为缓慢摆动的上帝光，
  prefers-reduced-motion 时静止；undersea `overflow:hidden` 防止光束平移造成横向溢出。
- 验证：headless Edge 1440/1024/390/320 截图，接缝与平铺缝目视无色差、无横向溢出、
  无控制台报错与失败请求；tests/browser_garden_ui.js 全过。
