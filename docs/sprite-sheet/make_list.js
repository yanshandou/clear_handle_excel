// 用 icons.js 生成 清单.md：node docs/sprite-sheet/make_list.js
const fs = require('fs');
const path = require('path');
const { icons, groups } = require('./icons.js');
const num = n => String(n).padStart(3, '0');
const cell = s => String(s).replace(/\|/g, '\\|');
let md = `# 精灵图清单（SVG 草稿）\n\n`
  + `一共 **${icons.length}** 个图标，分 ${groups.length} 组。打开同目录的 \`index.html\` 能看到每个图标的 SVG 草图（透明底、带编号），可以直接导出 PNG 精灵图和位置表。\n\n`
  + `| 组 | 个数 | 显示尺寸 |\n|---|---|---|\n`
  + groups.map(g => { const l = icons.filter(i => i.g === g); return `| ${g} | ${l.length} | ${l[0].size} |`; }).join('\n')
  + `\n\n## 怎么用\n\n`
  + `- 编号固定，英文名以后对应 CSS 的 \`.icon-英文名\`；美化时按编号对照，每格同样大小、图标居中、四周留约 10% 空白、透明底。\n`
  + `- 建议画成两张图：① 主页工具卡 + 空状态插画用 256px 一格（显示到 136px，高清屏够用）；② 其余小图标用 128px 一格（显示 16–40px）。\n`
  + `- 进度条底槽、进度条填充、完成横幅丝带要能横向拉伸：画成 256×64，左右两头各留 1/4 不拉伸。\n`
  + `- 标「可复用」的几个（说明块的放大镜、齿轮、循环）可以直接用同一张图，不用另画。\n`
  + `- 「上传管理页」那组现在用 lucide 线条图标、自成一套风格，要不要换由你定。\n\n`;
for (const g of groups) {
  md += `## ${g}\n\n| # | 英文名 | 名字 | 用在哪 | 现在用什么 | 出现在哪些页面 | 建议样子 |\n|---|---|---|---|---|---|---|\n`;
  md += icons.filter(i => i.g === g).map(i => `| ${num(i.no)} | \`${i.id}\` | ${cell(i.name)} | ${cell(i.use)} | ${cell(i.now)} | ${cell(i.pages)} | ${cell(i.look)} |`).join('\n') + '\n\n';
}
fs.writeFileSync(path.join(__dirname, '清单.md'), md);
console.log('写好', icons.length, '个');
