// 精灵图草稿：全站要换成精灵图的图标清单 + 每个图标的 SVG 草图（viewBox 0 0 64 64，透明底）。
// index.html 读这个文件画网格、导出 PNG；make_list.js 用它生成 清单.md。
// 每项：id 英文名（以后 CSS 里 .icon-<id>）、name 中文名、use 用途、now 现在用什么、pages 出现在哪些页面、size 显示尺寸、look 建议样子、svg 草图。
(function () {
  const INK = '#17325a';
  const C = {
    blue: '#2f86e8', deep: '#1d5fc0', sky: '#8fd0ff', teal: '#22c3b5', mint: '#9eecd6', lime: '#8ad13b', green: '#2fae5b',
    gold: '#ffc83a', orange: '#ff9a2e', coral: '#ff6b5e', red: '#e8443a', pink: '#ff8fb8', purple: '#8b5cf6', lav: '#c9b6ff',
    gray: '#b8c4d4', dgray: '#7d8aa0', white: '#ffffff', paper: '#f4f8ff', fold: '#d7e3f5', brown: '#b9773a',
  };
  const O = `stroke="${INK}" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"`;
  const L = (w = 3, col = INK) => `stroke="${col}" stroke-width="${w}" fill="none" stroke-linecap="round" stroke-linejoin="round"`;
  // 描边粗线：先画深色粗线，再画白（或彩色）细线，小尺寸也看得清
  const glyph = (d, col = '#fff', w = 6) => `<path d="${d}" ${L(w + 5)}/><path d="${d}" ${L(w, col)}/>`;
  const shine = d => `<path d="${d}" ${L(3, 'rgba(255,255,255,.8)')}/>`;
  const txt = (x, y, s, size = 14, col = INK, weight = 900) =>
    `<text x="${x}" y="${y}" font-family="Arial Black,Arial,sans-serif" font-weight="${weight}" font-size="${size}" text-anchor="middle" fill="${col}">${s}</text>`;
  const otxt = (x, y, s, size = 16, col = '#fff') =>
    `<text x="${x}" y="${y}" font-family="Arial Black,Arial,sans-serif" font-weight="900" font-size="${size}" text-anchor="middle" fill="${col}" stroke="${INK}" stroke-width="4" paint-order="stroke">${s}</text>`;

  // ---- 零件 ----
  function paper(x, y, w, h, fill = C.paper, f = 10) {
    return `<path d="M${x} ${y}h${w - f}l${f} ${f}v${h - f}h${-w}z" fill="${fill}" ${O}/><path d="M${x + w - f} ${y}v${f}h${f}" fill="${C.fold}" ${O}/>`;
  }
  function lines(x, y, w, n, gap = 6, col = C.gray) {
    let s = '';
    for (let i = 0; i < n; i++) s += `<path d="M${x} ${y + i * gap}h${i === n - 1 ? w * 0.6 : w}" ${L(3, col)}/>`;
    return s;
  }
  function sheet(x, y, w, h, head = C.green, cols = 3, rows = 3) {
    const hh = h / (rows + 1);
    let s = `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="5" fill="#fff"/>`;
    s += `<path d="M${x} ${y + hh}V${y + 5}a5 5 0 0 1 5-5h${w - 10}a5 5 0 0 1 5 5V${y + hh}z" fill="${head}"/>`;
    for (let i = 1; i <= rows; i++) if (i > 1) s += `<path d="M${x} ${y + hh * i}h${w}" ${L(2, '#9fb3cc')}/>`;
    for (let j = 1; j < cols; j++) s += `<path d="M${x + (w / cols) * j} ${y + hh}v${h - hh}" ${L(2, '#9fb3cc')}/>`;
    s += `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="5" fill="none" ${O}/>`;
    return s;
  }
  function folder(x, y, w, h, col = C.gold, open = false) {
    let s = `<path d="M${x} ${y + 4}a4 4 0 0 1 4-4h${w * 0.3}l5 6h${w * 0.7 - 13}a4 4 0 0 1 4 4v${h - 10}h${-w}z" fill="${C.orange}" ${O}/>`;
    if (open) s += `<path d="M${x - 3} ${y + 14}h${w + 4}l-5 ${h - 14}h${-w + 6}z" fill="${col}" ${O}/>`;
    else s += `<rect x="${x}" y="${y + 11}" width="${w}" height="${h - 11}" rx="4" fill="${col}" ${O}/>`;
    return s;
  }
  function mag(cx, cy, r, glass = '#cdeeff') {
    const a = r * 0.72, b = r * 1.75;
    return `<path d="M${cx + a} ${cy + a}L${cx + b} ${cy + b}" ${L(11)}/><path d="M${cx + a + 2} ${cy + a + 2}L${cx + b - 1} ${cy + b - 1}" ${L(5, C.orange)}/>`
      + `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${glass}" stroke="${INK}" stroke-width="3"/><circle cx="${cx}" cy="${cy}" r="${r - 4}" fill="none" stroke="${C.blue}" stroke-width="4"/>`
      + shine(`M${cx - r * 0.45} ${cy - r * 0.1}a${r * 0.5} ${r * 0.5} 0 0 1 ${r * 0.4}-${r * 0.38}`);
  }
  const disc = (col, r = 26, cx = 32, cy = 32) =>
    `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${col}" ${O}/>` + shine(`M${cx - r * 0.62} ${cy - r * 0.2}a${r * 0.68} ${r * 0.68} 0 0 1 ${r * 0.48}-${r * 0.5}`);
  const tile = (col, x = 8, y = 8, w = 48, h = 48) =>
    `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="12" fill="${col}" ${O}/>` + shine(`M${x + 7} ${y + 17}v-4a6 6 0 0 1 6-6h6`);
  const check = (d = 'M20 33l8 8 16-17', col = '#fff') => glyph(d, col, 6);
  const fatArrow = (dir, col, cx = 32, cy = 32, s = 1) => {
    // 向上的胖箭头，再旋转
    const p = `M${cx} ${cy - 20 * s}l${16 * s} ${17 * s}h${-9 * s}v${21 * s}h${-14 * s}v${-21 * s}h${-9 * s}z`;
    const rot = { up: 0, right: 90, down: 180, left: 270 }[dir];
    return `<path d="${p}" fill="${col}" ${O} transform="rotate(${rot} ${cx} ${cy})"/>`;
  };
  const chain = (broken = false) => broken
    ? `<rect x="4" y="22" width="28" height="18" rx="9" ${L(11)} transform="rotate(-30 18 31)"/><rect x="4" y="22" width="28" height="18" rx="9" ${L(5, C.gold)} transform="rotate(-30 18 31)"/>`
      + `<rect x="32" y="26" width="28" height="18" rx="9" ${L(11)} transform="rotate(-30 46 35)"/><rect x="32" y="26" width="28" height="18" rx="9" ${L(5, C.gold)} transform="rotate(-30 46 35)"/>`
      + `<path d="M30 12l2 8M38 50l-2 -7M24 46l5-4M41 16l-4 5" ${L(3, C.red)}/>`
    : `<rect x="6" y="24" width="30" height="18" rx="9" ${L(11)} transform="rotate(-35 21 33)"/><rect x="6" y="24" width="30" height="18" rx="9" ${L(5, C.gold)} transform="rotate(-35 21 33)"/>`
      + `<rect x="28" y="22" width="30" height="18" rx="9" ${L(11)} transform="rotate(-35 43 31)"/><rect x="28" y="22" width="30" height="18" rx="9" ${L(5, C.gold)} transform="rotate(-35 43 31)"/>`;
  const pdfDoc = (x, y, w, h, col = C.red) => paper(x, y, w, h) + `<rect x="${x - 3}" y="${y + h - 20}" width="${w * 0.8}" height="13" rx="3" fill="${col}" ${O}/>`
    + `<text x="${x - 3 + w * 0.4}" y="${y + h - 9.5}" font-family="Arial Black,Arial" font-weight="900" font-size="9" text-anchor="middle" fill="#fff">PDF</text>`;
  const xlsDoc = (x, y, w, h) => paper(x, y, w, h) + `<rect x="${x - 3}" y="${y + h - 22}" width="17" height="17" rx="4" fill="${C.green}" ${O}/>`
    + `<path d="M${x + 1} ${y + h - 18}l9 9M${x + 10} ${y + h - 18}l-9 9" ${L(2.6, '#fff')}/>` + lines(x + 18, y + h - 20, w - 24, 3, 5);
  const spin = (col = C.blue) => `<g transform="rotate(-100 32 32)"><circle cx="32" cy="32" r="20" fill="none" stroke="${C.gray}" stroke-width="8" opacity=".6"/><circle cx="32" cy="32" r="20" ${L(11)} stroke-dasharray="70 200"/><circle cx="32" cy="32" r="20" fill="none" stroke="${col}" stroke-width="6" stroke-linecap="round" stroke-dasharray="70 200"/></g>`;
  const star = (cx, cy, r, col = C.gold) => {
    let d = '';
    for (let i = 0; i < 10; i++) { const a = Math.PI / 5 * i - Math.PI / 2, rr = i % 2 ? r * 0.45 : r; d += `${i ? 'L' : 'M'}${(cx + rr * Math.cos(a)).toFixed(1)} ${(cy + rr * Math.sin(a)).toFixed(1)}`; }
    return `<path d="${d}z" fill="${col}" ${O}/>`;
  };
  const person = (cx, cy, col = C.blue, s = 1) => `<circle cx="${cx}" cy="${cy - 9 * s}" r="${8 * s}" fill="#ffd9b8" ${O}/><path d="M${cx - 14 * s} ${cy + 16 * s}a${14 * s} ${13 * s} 0 0 1 ${28 * s} 0z" fill="${col}" ${O}/>`;
  const hand = `<path d="M22 56c-6-6-10-12-10-20v-6a3 3 0 0 1 6 0v6V14a3.5 3.5 0 0 1 7 0v14V10a3.5 3.5 0 0 1 7 0v18V12a3.5 3.5 0 0 1 7 0v18v-12a3.5 3.5 0 0 1 7 0v20c0 10-5 18-13 18z" fill="#ffd9b8" ${O}/>`;
  const cell = (x, y, w, h, col) => `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="${col}" ${O}/>`;
  const pageLand = (x, y, w, h, col = C.paper) => `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="3" fill="${col}" ${O}/>`;

  const I = [];
  const add = (g, id, name, use, now, pages, size, look, svg) => I.push({ g, id, name, use, now, pages, size, look, svg });

  // ===== 1 主页工具卡 =====
  const G1 = '主页工具卡';
  add(G1, 'tool-clean', '精简 Sheet', '主页大卡、页头', '海岛精灵图 icon-clean', '主页 /、/clean', '大卡 136px，画 256px', '几张表叠在一起 + 一把清洁刷',
    sheet(8, 14, 34, 30, C.teal) + sheet(14, 20, 34, 30) + `<path d="M40 50l14-24" ${L(9)}/><path d="M40 50l14-24" ${L(4, C.brown)}/><path d="M34 46l12 7-4 7-14-6z" fill="${C.gold}" ${O}/><path d="M33 52l-3 6M37 54l-2 6M41 56l-1 5" ${L(2.5)}/>`);
  add(G1, 'tool-adjust', '调整页面', '主页横卡', '海岛精灵图 icon-adjust', '主页、/pdf-review', '横卡 104px，画 256px', '一张纸 + 黄色三角尺和尺子',
    paper(12, 8, 34, 44) + lines(18, 20, 20, 3) + `<path d="M30 56V30l26 26z" fill="${C.gold}" ${O}/><path d="M36 50v-6l6 6z" fill="#fff8d8" ${O}/><rect x="6" y="48" width="30" height="9" rx="2" fill="${C.sky}" ${O}/><path d="M12 48v4M18 48v3M24 48v4M30 48v3" ${L(2)}/>`);
  add(G1, 'tool-cells', '查看切字', '主页横卡', '海岛精灵图 icon-cells', '主页、/workbench', '横卡 104px，画 256px', '放大镜照着表格，里面一格橙色高亮',
    sheet(6, 8, 42, 38) + cell(20, 26, 14, 10, C.orange) + mag(38, 36, 12));
  add(G1, 'tool-upload', '上传 Excel', '主页大卡', '海岛精灵图 icon-crate', '主页、/distribute', '大卡 136px，画 256px', '蓝色货箱 + 绿色向上箭头',
    `<rect x="8" y="30" width="48" height="26" rx="4" fill="${C.blue}" ${O}/><path d="M8 40h48M24 30v26M40 30v26" ${L(3)}/>` + shine('M13 35h8') + fatArrow('up', C.lime, 32, 20, 0.75));
  add(G1, 'tool-bl', '排查 BL 超链接', '主页小卡', '海岛精灵图 icon-link（和引用错误共用）', '主页、/repair-bl-links', '小卡 52px，画 128px', '断开的金色链条 + 红色断裂线',
    chain(true));
  add(G1, 'tool-links', '修复引用错误', '主页小卡', '海岛精灵图 icon-link（和 BL 共用）', '主页、/repair-links', '小卡 52px，画 128px', '一格写着 #REF! + 修好的链条',
    `<rect x="6" y="8" width="40" height="22" rx="5" fill="#fff" ${O}/>` + txt(26, 25, '#REF!', 11, C.red) + `<g transform="translate(16 20) scale(.75)">${chain()}</g>`);
  add(G1, 'tool-hidden', '修复隐藏页', '主页小卡', '海岛精灵图 icon-eye', '主页、/repair-hidden', '小卡 52px，画 128px', '一排表标签 + 眼睛（一半闭着）',
    `<path d="M6 50h14l3-8h10l3 8h22" fill="none" ${O}/><path d="M23 42h10l3 8H20z" fill="${C.mint}" ${O}/><path d="M38 50l3-8h10l3 8z" fill="${C.gray}" ${O} stroke-dasharray="4 3"/>`
    + `<path d="M6 27c9-14 43-14 52 0c-9 14-43 14-52 0z" fill="#fff" ${O}/><circle cx="32" cy="29" r="10" fill="${C.blue}" ${O}/><circle cx="32" cy="30" r="4" fill="${INK}"/><path d="M6 27c9-14 43-14 52 0c-9-6-43-6-52 0z" fill="${C.lav}" ${O}/><path d="M10 31l-4 4M58 31l4 4" ${L(2.5)}/>`);
  add(G1, 'tool-font', '统一字体', '主页小卡', '海岛精灵图 icon-type', '主页、/fix-fonts', '小卡 52px，画 128px', '蓝色大 A + 绿色小 a，字块感',
    `<rect x="6" y="10" width="30" height="40" rx="7" fill="${C.blue}" ${O}/>` + otxt(21, 42, 'A', 30) + `<rect x="32" y="24" width="26" height="30" rx="7" fill="${C.lime}" ${O}/>` + otxt(45, 47, 'a', 24));
  add(G1, 'tool-colors', '修复颜色', '主页小卡', '海岛精灵图 icon-palette', '主页、/repair-colors', '小卡 52px，画 128px', '调色盘，几团颜料',
    `<path d="M32 6C16 6 6 18 6 31c0 14 12 26 26 26c5 0 6-4 3-8c-2-3 0-7 4-7h7c7 0 12-5 12-12C58 16 47 6 32 6z" fill="#ffe3b8" ${O}/>`
    + `<circle cx="20" cy="26" r="5" fill="${C.red}" ${O}/><circle cx="30" cy="16" r="5" fill="${C.gold}" ${O}/><circle cx="43" cy="19" r="5" fill="${C.green}" ${O}/><circle cx="47" cy="31" r="5" fill="${C.blue}" ${O}/><circle cx="20" cy="41" r="5" fill="${C.purple}" ${O}/>`);
  add(G1, 'tool-view', '视图标准化', '主页小卡', '海岛精灵图 icon-compass', '主页（对话框）', '小卡 52px，画 128px', '指南针，指针指向左上角的 A1',
    disc(C.sky, 25) + `<circle cx="32" cy="32" r="18" fill="#fff" ${O}/><path d="M32 32L20 20l16 6z" fill="${C.red}" ${O}/><path d="M32 32l12 12-16-6z" fill="${C.blue}" ${O}/>` + otxt(18, 20, 'A1', 12, C.gold));
  add(G1, 'tool-new', '新任务 · 拿文件', '主页小卡、当前任务条', '海岛精灵图 icon-new + emoji 🆕', '主页、/distribute/start', '小卡 52px，画 128px', '打开的宝箱，里面露出文件 + 闪光',
    `<path d="M8 30h48v24H8z" fill="${C.blue}" ${O}/><path d="M8 30l6-14h36l6 14z" fill="${C.deep}" ${O}/>` + paper(20, 14, 22, 24) + `<rect x="26" y="34" width="12" height="10" rx="2" fill="${C.gold}" ${O}/><path d="M8 40h48" ${L(3)}/>` + star(52, 12, 7));
  add(G1, 'tool-images', '修复图片', '主页小卡', '海岛精灵图 icon-image', '主页、/fix-images', '小卡 52px，画 128px', '画框里一张风景，右上角红色标注框歪了',
    `<rect x="6" y="12" width="44" height="38" rx="4" fill="${C.sky}" ${O}/><path d="M6 44l12-12 10 9 8-6 14 11v4H6z" fill="${C.green}" ${O}/><circle cx="38" cy="24" r="5" fill="${C.gold}" ${O}/>`
    + `<rect x="32" y="6" width="26" height="20" rx="2" ${L(7)} transform="rotate(10 45 16)"/><rect x="32" y="6" width="26" height="20" rx="2" ${L(3.5, C.red)} transform="rotate(10 45 16)"/>`);
  add(G1, 'tool-widths', '修复列宽', '主页小卡', '海岛精灵图 icon-adjust（和调整页面共用）', '主页、/fix-images?tool=widths', '小卡 52px，画 128px', '表格中间一列 + 左右双向箭头',
    sheet(6, 10, 52, 44, C.green, 3, 3) + cell(23.3, 21, 17.4, 33, C.mint) + glyph('M16 38h32M22 32l-6 6 6 6M42 32l6 6-6 6', C.blue, 4));
  add(G1, 'tool-bundle', '批量导出同名 PDF', '主页小卡', '海岛精灵图 icon-document', '主页、/pdf-bundle', '小卡 52px，画 128px', '一叠红色 PDF 文件',
    paper(18, 4, 34, 42, C.fold) + paper(12, 10, 34, 42) + pdfDoc(8, 16, 34, 42));
  add(G1, 'tool-pick', '挑文件复制', '主页小卡', '海岛精灵图 icon-map', '主页、/pick-copy', '小卡 52px，画 128px', '一列文件，其中两个打勾 + 复制符号',
    `<rect x="6" y="8" width="38" height="48" rx="5" fill="#fff" ${O}/>` + [16, 28, 40].map((y, i) => `<rect x="11" y="${y}" width="9" height="9" rx="2" fill="${i === 1 ? '#fff' : C.green}" ${O}/>` + (i === 1 ? '' : `<path d="M13 ${y + 5}l2 2 4-5" ${L(2.4, '#fff')}/>`) + `<path d="M24 ${y + 4.5}h14" ${L(3, C.gray)}/>`).join('')
    + `<rect x="34" y="32" width="22" height="24" rx="4" fill="${C.pink}" ${O}/><rect x="40" y="26" width="18" height="22" rx="4" fill="#fff" ${O}/>`);
  add(G1, 'tool-json', '按 JSON 复制 Excel', '主页小卡', '海岛精灵图 icon-layers', '主页、/distribute?mode=json', '小卡 52px，画 128px', '{ } 文件 + 后面一叠 Excel',
    xlsDoc(24, 6, 32, 40) + paper(6, 16, 32, 42, '#fff7d6') + txt(22, 46, '{ }', 18, C.purple));
  add(G1, 'tool-pcl', 'PCL 单体测试检查', '主页小卡', '海岛精灵图 icon-check', '主页、/pcl', '小卡 52px，画 128px', '青绿色剪贴板 + 一个大勾',
    `<rect x="10" y="10" width="44" height="48" rx="6" fill="${C.teal}" ${O}/><rect x="16" y="16" width="32" height="36" rx="3" fill="#fff" ${O}/><rect x="22" y="5" width="20" height="10" rx="3" fill="${C.gold}" ${O}/>` + check('M22 34l7 7 13-14', C.green));
  add(G1, 'tool-finish', '完成任务', '主页小卡、当前任务条', '海岛精灵图 icon-finish + emoji 🏁', '主页、/finish', '小卡 52px，画 128px', '金色奖杯 + 小旗',
    `<path d="M18 8h28v12c0 10-6 16-14 16S18 30 18 20z" fill="${C.gold}" ${O}/><path d="M18 12H9c0 9 4 13 10 13M46 12h9c0 9-4 13-10 13" ${L(3)}/><path d="M28 36h8v8h-8z" fill="${C.orange}" ${O}/><rect x="18" y="44" width="28" height="10" rx="3" fill="${C.blue}" ${O}/>` + shine('M23 13v8') + star(32, 20, 6, '#fff3b0'));

  // ===== 2 顶部导航 =====
  const G2 = '顶部导航';
  const navs = [
    ['nav-home', '主页', '页头「主页」、各页「🏠 返回首页」', '海岛精灵图 icon-home + emoji 🏠', '所有页面', '小房子，蓝屋顶白墙',
      `<path d="M10 30L32 10l22 20" fill="${C.blue}" ${O}/><path d="M16 26v28h32V26L32 13z" fill="#fff" ${O}/><path d="M8 31L32 9l24 22" ${L(5)}/><path d="M8 31L32 9l24 22" ${L(2, C.blue)}/><rect x="27" y="38" width="10" height="16" rx="2" fill="${C.orange}" ${O}/>`],
    ['nav-clean', '精简 Sheet', '页头导航', '海岛精灵图 icon-grid', '所有页面页头', '绿色表格（小号，和主页卡区分开可更简单）',
      sheet(8, 10, 48, 44)],
    ['nav-cells', '切字检查', '页头导航', '海岛精灵图 icon-search', '所有页面页头', '放大镜',
      mag(27, 27, 15)],
    ['nav-analysis', '扫描分析', '页头导航、「📈 查看本次报告」', '海岛精灵图 icon-chart + emoji 📈', '所有页面页头、/workbench', '彩色柱状图往上走',
      `<path d="M8 8v48h48" ${L(4)}/>` + [[14, 34, C.blue], [26, 24, C.teal], [38, 14, C.lime]].map(([x, y, c]) => `<rect x="${x}" y="${y}" width="9" height="${52 - y}" rx="2" fill="${c}" ${O}/>`).join('') + glyph('M14 30l12-10 12-6 10-6', C.orange, 3)],
    ['nav-history', '文件夹历史', '页头导航', '海岛精灵图 icon-folder', '所有页面页头、/history', '黄色文件夹 + 小时钟',
      folder(6, 12, 42, 36) + disc('#fff', 12, 46, 44) + `<path d="M46 37v7h6" ${L(3)}/>`],
    ['nav-records', '执行记录', '页头导航', '海岛精灵图 icon-layers', '所有页面页头、/records', '一叠彩色文件',
      `<path d="M32 8l24 12-24 12L8 20z" fill="${C.sky}" ${O}/><path d="M8 30l24 12 24-12" ${L(3)}/><path d="M8 30l24 12 24-12v0" ${L(3)}/><path d="M8 40l24 12 24-12" ${L(3)}/><path d="M8 20v0" ${L(3)}/>`],
    ['nav-diff', 'Excel 对比', '上传页「🔍 对比」、/excel-diff 页头', 'emoji 🔍', '/distribute、/excel-diff、/excel-diff/all', '左右两张表，中间红绿差异条',
      `<rect x="6" y="10" width="24" height="44" rx="4" fill="#fff" ${O}/><rect x="34" y="10" width="24" height="44" rx="4" fill="#fff" ${O}/><rect x="10" y="24" width="16" height="7" fill="#ffb3ab" ${O}/><rect x="38" y="24" width="16" height="7" fill="#a8eab9" ${O}/>` + lines(11, 18, 14, 1) + lines(39, 18, 14, 1) + lines(11, 40, 14, 2) + lines(39, 40, 14, 2)],
    ['nav-more', '更多工具', '页头/主页的工具入口', '海岛精灵图 icon-grid', '页头', '四个圆角方块（2×2）',
      [[8, 8, C.blue], [34, 8, C.teal], [8, 34, C.gold], [34, 34, C.coral]].map(([x, y, c]) => `<rect x="${x}" y="${y}" width="22" height="22" rx="6" fill="${c}" ${O}/>`).join('')],
    ['nav-settings', '设置 / 路径设置', '主页「路径设置」「⏸ 出错时」等设置按钮', '海岛精灵图 icon-settings', '主页', '青绿色齿轮',
      `<path d="M28 6h8l1.5 6 5 2.5 5.5-3 5.5 5.5-3 5.5 2.5 5 6 1.5v8l-6 1.5-2.5 5 3 5.5-5.5 5.5-5.5-3-5 2.5L36 58h-8l-1.5-6-5-2.5-5.5 3-5.5-5.5 3-5.5L11 36l-6-1.5v-8L11 25l2.5-5-3-5.5 5.5-5.5 5.5 3 5-2.5z" fill="${C.teal}" ${O}/><circle cx="32" cy="31" r="9" fill="#fff" ${O}/>`],
  ];
  navs.forEach(n => add(G2, n[0], n[1], n[2], n[3], n[4], '16–20px，画 128px', n[5], n[6]));

  // ===== 3 状态徽章 =====
  const G3 = '状态徽章';
  const st = [
    ['st-done', '做过 / 完成 ✓', '主页卡右上角绿圈 ✓、进度框完成、检查通过', 'CSS 圆 + 文字 ✓ / emoji ✅', '主页、所有批量进度框、/finish、/clean', '绿色圆徽章 + 白勾', disc(C.green) + check()],
    ['st-warn', '有问题 !', '主页卡右上角黄圈 !、进度框做完但有出错', 'CSS 圆 + 文字 !', '主页、所有批量进度框', '黄色圆徽章 + 深色 !', disc(C.gold) + `<path d="M32 17v16" ${L(7)}/><circle cx="32" cy="44" r="4" fill="${INK}"/>`],
    ['st-error', '出错 ✗', '结果列表的 ❌、PCL 的 NG', 'emoji ❌ / 字符 ✗', '/distribute、/records、/pcl、/workbench', '红色圆徽章 + 白叉', disc(C.red) + glyph('M23 23l18 18M41 23L23 41')],
    ['st-stop', '禁止 / 不能做 ⛔', '对话框出错、文件夹没建成、不复制', 'emoji ⛔ 🛑', '全站对话框、/distribute', '红色圆 + 白横杠', disc(C.red) + glyph('M19 32h26')],
    ['st-busy', '进行中', '正在打开 / 正在扫描 / 正在导出的标记', 'emoji ⏳ / CSS 转圈动画', '所有批量页面、提示条', '蓝色圆 + 白色转圈箭头', disc(C.blue) + glyph('M44 32a12 12 0 1 1-4-9M41 15v9h-9')],
    ['st-wait', '等待 / 没开始 ○', '收尾检查里还没轮到的项、排队中', '字符 ○', '/finish、排队状态', '灰色空心圆 + 三个点', `<circle cx="32" cy="32" r="24" fill="#fff" ${O} stroke-dasharray="6 5"/><circle cx="22" cy="32" r="3.5" fill="${C.dgray}"/><circle cx="32" cy="32" r="3.5" fill="${C.dgray}"/><circle cx="42" cy="32" r="3.5" fill="${C.dgray}"/>`],
    ['st-ready', '就绪', '页面右上角状态胶囊「✅ 就绪」', 'emoji ✅', '/clean、/workbench、/clean/preview', '绿色圆角方块 + 白勾', tile(C.green) + check()],
    ['st-idle', '空闲', '主页任务条「⚪ 空闲」', 'emoji ⚪', '主页、/records', '灰色圆 + 白色小月亮 / zZ', disc(C.gray) + otxt(32, 40, 'z', 22)],
    ['st-skip', '跳过', '进度框「⏭ 跳过这一份」、跳过的文件', 'emoji ⏭', '所有批量进度框', '橙色圆 + 白色快进', disc(C.orange) + `<path d="M19 21l12 11-12 11zM31 21l12 11-12 11z" fill="#fff" ${O}/><path d="M45 21v22" ${L(4)}/>`],
    ['st-pause', '暂停 / 出错时', '「⏸ 出错时」设置、出错弹框', 'emoji ⏸', '主页、所有批量弹框', '蓝紫色圆 + 白色双竖条', disc(C.purple) + `<rect x="22" y="20" width="7" height="24" rx="2" fill="#fff" ${O}/><rect x="35" y="20" width="7" height="24" rx="2" fill="#fff" ${O}/>`],
    ['st-stopall', '结束整个任务', '进度框「⏹ 结束整个任务」', 'emoji ⏹', '所有批量进度框', '红色圆 + 白方块', disc(C.coral) + `<rect x="22" y="22" width="20" height="20" rx="3" fill="#fff" ${O}/>`],
    ['st-locked', '暂不能打开', '查看切字「🔒 暂不能打开」、先核对', 'emoji 🔒', '/workbench、/repair-hidden', '金色小锁', `<path d="M21 28v-7a11 11 0 0 1 22 0v7" ${L(9)}/><path d="M21 28v-7a11 11 0 0 1 22 0v7" ${L(4, C.gray)}/><rect x="14" y="27" width="36" height="28" rx="6" fill="${C.gold}" ${O}/><circle cx="32" cy="39" r="4" fill="${INK}"/><path d="M32 41v6" ${L(4)}/>`],
    ['st-stale', 'Excel 改过，PDF 要重导', '调整页面 PDF 卡上的「已过期」标记', 'emoji 🔄 + 文字', '/pdf-review', '黄色圆 + 两个循环箭头', disc(C.gold) + `<g transform="translate(11 11) scale(.66)">` + glyph('M14 28a18 18 0 0 1 32-9M50 36a18 18 0 0 1-32 9', '#fff', 6) + `<path d="M48 8v14H34z" fill="#fff" ${O}/><path d="M16 56V42h14z" fill="#fff" ${O}/></g>`],
    ['st-opened', '已打开 Excel', '文件名旁「✓ 已打开」、已打开待复检', 'CSS ::after 文字', '/workbench、/pdf-review', '绿色 Excel 小标 + 打开的门/箭头', tile(C.green, 6, 10, 40, 44) + `<path d="M16 22l14 18M30 22L16 40" ${L(5, '#fff')}/>` + fatArrow('right', C.gold, 46, 34, 0.6)],
    ['st-new', '新 NEW', '新任务、刚生成的条目角标', 'emoji 🆕', '主页、/distribute/start、/finish', '橙红色小旗形角标写 NEW', `<path d="M4 18h56l-6 14 6 14H4z" fill="${C.coral}" ${O}/>` + otxt(30, 38, 'NEW', 16)],
  ];
  st.forEach(n => add(G3, n[0], n[1], n[2], n[3], n[4], '16–28px，画 128px', n[5], n[6]));

  // ===== 4 提示条 / 对话框 =====
  const G4 = '提示条和对话框';
  const tw = [
    ['msg-busy', '提示条：正在…', '底部提示条（打开 Excel 等）蓝色', '提示条 SVG（open_excel.js TOAST_ICONS）', '全站', '蓝色圆 + 转圈箭头（比徽章大、更精致）', spin(C.sky)],
    ['msg-ok', '提示条：已完成', '底部提示条绿色', '提示条 SVG', '全站', '绿色圆 + 白勾 + 小星光', disc(C.green, 24) + check() + star(52, 12, 7, '#fff3b0')],
    ['msg-warn', '提示条：注意', '底部提示条橙色、对话框「注意」', '提示条 SVG + emoji ⚠️', '全站', '橙黄三角 + 深色 !', `<path d="M32 6l27 48H5z" fill="${C.gold}" ${O}/><path d="M32 22v16" ${L(7)}/><circle cx="32" cy="46" r="4" fill="${INK}"/>`],
    ['msg-error', '提示条：打不开 / 出错了', '底部提示条红色、对话框「出错了」', '提示条 SVG + emoji ⛔', '全站', '红色八角牌 + 白叉', `<path d="M22 6h20l14 14v20L42 56H22L8 40V20z" fill="${C.red}" ${O}/>` + glyph('M23 22l18 18M41 22L23 40')],
    ['msg-info', '对话框：提示', 'appAlert 默认图标', 'emoji ℹ️', '全站', '蓝色圆 + 白色 i', disc(C.blue) + `<circle cx="32" cy="19" r="4.5" fill="#fff" ${O}/>` + glyph('M32 29v16')],
    ['msg-confirm', '对话框：请确认', 'appConfirm 默认图标', 'emoji ℹ️/⚠️', '全站', '紫色对话气泡 + 白色 ?', `<path d="M10 10h44a4 4 0 0 1 4 4v28a4 4 0 0 1-4 4H28l-12 10v-10h-6a4 4 0 0 1-4-4V14a4 4 0 0 1 4-4z" fill="${C.purple}" ${O}/>` + otxt(32, 39, '?', 28)],
    ['msg-danger', '对话框：危险操作', '删除、覆盖前的确认', 'emoji ⚠️', '/pdf-review、/clean、/distribute', '红色三角 + 垃圾桶', `<path d="M32 6l27 48H5z" fill="${C.coral}" ${O}/><path d="M24 30h16l-2 18H26z" fill="#fff" ${O}/><path d="M22 30h20M29 26h6" ${L(3)}/>`],
    ['msg-tip', '小提示 💡', '结果里的提示行', 'emoji 💡', '/distribute', '黄色灯泡', `<path d="M32 6a17 17 0 0 0-10 31c2 2 3 4 3 7h14c0-3 1-5 3-7A17 17 0 0 0 32 6z" fill="${C.gold}" ${O}/><rect x="24" y="44" width="16" height="10" rx="3" fill="${C.gray}" ${O}/>` + shine('M24 18a10 10 0 0 1 6-5')],
  ];
  tw.forEach(n => add(G4, n[0], n[1], n[2], n[3], n[4], '提示条 24–28px，对话框 32–40px，画 128px', n[5], n[6]));

  // ===== 5 按钮动作 =====
  const G5 = '按钮动作';
  const btn = [
    ['act-scan', '扫描', '「🔎 扫描备份表」「开始扫描」', 'emoji 🔎', '/clean、/workbench、/repair-bl-links', '放大镜扫过一条亮光', mag(26, 26, 15) + `<path d="M6 50h22" ${L(4, C.teal)}/><path d="M6 57h14" ${L(4, C.teal)}/>`],
    ['act-run', '执行 / 清理', '「🧹 执行」、清掉残留标记', 'emoji 🧹', '/clean、/workbench', '扫帚', `<path d="M44 6L30 34" ${L(9)}/><path d="M44 6L30 34" ${L(4, C.brown)}/><path d="M22 30l16 8-4 18-24-10z" fill="${C.gold}" ${O}/><path d="M16 46l5-9M22 50l5-9M29 53l4-10" ${L(2.5)}/>`],
    ['act-compare', '开始对比', '「🔍 开始对比」', 'emoji 🔍', '/excel-diff、/excel-diff/all', '两张纸 + 放大镜', paper(6, 6, 26, 34) + paper(22, 12, 26, 34) + mag(40, 40, 10)],
    ['act-refresh', '刷新 / 重新生成', '「🔄 刷新卡片」「🔄 重新生成」「↻ 重新读取」', 'emoji 🔄 ↻', '/workbench、/pdf-review、/history、/pcl', '两个绕圈的蓝色箭头', glyph('M14 28a18 18 0 0 1 32-9M50 36a18 18 0 0 1-32 9', C.blue, 6) + `<path d="M48 8v14H34z" fill="${C.blue}" ${O}/><path d="M16 56V42h14z" fill="${C.blue}" ${O}/>`],
    ['act-recheck', '复检', '「🔁 复检」、扫完以后', 'emoji 🔁', '/workbench、主页', '循环箭头中间一个勾', glyph('M14 28a18 18 0 0 1 32-9M50 36a18 18 0 0 1-32 9', C.teal, 6) + `<path d="M48 8v14H34z" fill="${C.teal}" ${O}/><path d="M16 56V42h14z" fill="${C.teal}" ${O}/>` + check('M24 32l6 6 11-11', C.green)],
    ['act-delete', '删除 / 删除标红', '「🗑 删除标红」、删表', 'emoji 🗑', '/pdf-review、/clean', '红色垃圾桶', `<path d="M14 18h36l-4 38H18z" fill="${C.coral}" ${O}/><rect x="9" y="12" width="46" height="8" rx="3" fill="${C.red}" ${O}/><path d="M26 12V7h12v5" ${L(3)}/><path d="M25 27v21M32 27v21M39 27v21" ${L(3)}/>`],
    ['act-fix', '统一修改 / 只修这一类', '「🛠 统一修改」「🛠 只修这一类」', 'emoji 🛠', '/pdf-review、/workbench', '扳手 + 螺丝刀交叉', `<path d="M14 50L44 20" ${L(11)}/><path d="M14 50L44 20" ${L(5, C.gray)}/><path d="M46 8a10 10 0 0 0-12 13l6 4 4 6a10 10 0 0 0 13-12l-6 6-6-2-1-6z" fill="${C.gray}" ${O}/><path d="M18 14l30 30" ${L(11)}/><path d="M18 14l30 30" ${L(5, C.orange)}/><path d="M46 42l10 10-4 4-10-10z" fill="${C.red}" ${O}/>`],
    ['act-export-pdf', '导出 PDF', '「📕 批量导出同名 PDF」「导出 PDF」', 'emoji 📕', '/pdf-bundle、主页对话框、/excel-diff/all', 'PDF 文件 + 向外的箭头', pdfDoc(8, 6, 34, 46) + fatArrow('right', C.lime, 48, 30, 0.6)],
    ['act-copy', '复制路径 / 复制清单', '「复制路径」「📋 复制剩下路径」', 'emoji 📋 / 文字', '/fix-images、/pick-copy、/distribute、进度框', '剪贴板 + 前面一张纸', `<rect x="8" y="10" width="34" height="44" rx="5" fill="${C.orange}" ${O}/><rect x="16" y="5" width="18" height="9" rx="3" fill="${C.gray}" ${O}/>` + paper(26, 20, 30, 38) + lines(31, 34, 18, 3)],
    ['act-upload', '上传', '「⬆ 上传这一份」「🚚 按清单上传」', 'emoji ⬆ 🚚', '/distribute、/excel-diff', '托盘 + 向上箭头', `<path d="M6 40v14h52V40" ${L(9)}/><path d="M6 40v14h52V40" ${L(4, C.blue)}/>` + fatArrow('up', C.lime, 32, 28, 0.8)],
    ['act-fetch', '拿最新版 / 下载', '「⬇ 从目的地拿最新版」「从共享拿文件」', 'emoji ⬇', '/distribute、/distribute/start', '托盘 + 向下箭头', `<path d="M6 40v14h52V40" ${L(9)}/><path d="M6 40v14h52V40" ${L(4, C.blue)}/>` + fatArrow('down', C.sky, 32, 28, 0.8)],
    ['act-copy-files', '复制文件', '「📦 复制」「📥 按清单开始复制」', 'emoji 📦 📥', '/pick-copy、/distribute?mode=json', '纸箱 + 两张文件飞进去', `<path d="M8 28h48v28H8z" fill="${C.brown}" ${O}/><path d="M8 28l8-10h32l8 10" fill="#d9a05a" ${O}/>` + paper(18, 6, 18, 22) + paper(30, 10, 18, 22) + `<path d="M24 38h16" ${L(4)}/>`],
    ['act-select-all', '全选', '「☑ 全选待上传」「全选全部」', 'emoji ☑', '/distribute', '三个叠在一起的勾选框', [[16, 6], [11, 18], [6, 30]].map(([x, y], i) => `<rect x="${x + 6}" y="${y}" width="36" height="26" rx="6" fill="${['#fff', '#fff', C.green][i]}" ${O}/>`).join('') + check('M18 43l6 6 12-12')],
    ['act-next', '下一处 / 下一步', '「➡ 下一处」「下一张」', 'emoji ➡', '/workbench、/pdf-review', '胖箭头向右', fatArrow('right', C.blue)],
    ['act-prev', '上一张', '大图「← →」翻页', '字符 ←', '/workbench、/pdf-review', '胖箭头向左', fatArrow('left', C.blue)],
    ['act-close', '关闭 ✕', '对话框、大图右上角 ✕', '字符 ✕', '全站对话框、大图', '灰色圆 + 白叉', disc(C.dgray, 22) + glyph('M24 24l16 16M40 24L24 40', '#fff', 5)],
    ['act-reset', '恢复默认 ↺', '「↺ 恢复默认」', '字符 ↺', '/clean、/clean/preview', '逆时针箭头', glyph('M18 22a18 18 0 1 1-2 18', C.orange, 6) + `<path d="M8 8v18h18z" fill="${C.orange}" ${O}/>`],
    ['act-undo', '撤销 ↩', '「↩ 取消「无需处理」」、再点一次撤销', '字符 ↩', '/fix-images、/pdf-review', '回头的箭头', glyph('M22 22h18a14 14 0 0 1 0 28H24', C.purple, 6) + `<path d="M26 8L8 22l18 14z" fill="${C.purple}" ${O}/>`],
    ['act-rename', '改名 / 编辑 ✏', '预览里改表名、「✏ 新名字」', 'emoji ✏', '/clean、/clean/preview', '黄色铅笔', `<path d="M10 54l4-14L42 12l10 10-28 28z" fill="${C.gold}" ${O}/><path d="M42 12l10 10 4-4a4 4 0 0 0 0-6l-4-4a4 4 0 0 0-6 0z" fill="${C.pink}" ${O}/><path d="M10 54l4-14 10 10z" fill="#ffe3b8" ${O}/>`],
    ['act-group', '组合 / 只组合', '「📌 组合」「只组合」', 'emoji 📌', '/fix-images', '一张截图 + 红框，中间别针', `<rect x="6" y="12" width="38" height="30" rx="3" fill="${C.sky}" ${O}/><rect x="22" y="26" width="34" height="26" rx="2" ${L(7)}/><rect x="22" y="26" width="34" height="26" rx="2" ${L(3.5, C.red)}/><path d="M30 8l12 12-6 2-8 8-4-4 8-8z" fill="${C.red}" ${O}/>`],
    ['act-side-by-side', '对照打开', '「🆚 对照打开原版」「↔ 同时打开当前与备份」', 'emoji 🆚 ↔', '/fix-images、/repair-*、自动检查清单', '两个窗口并排 + 双向箭头', `<rect x="4" y="12" width="26" height="36" rx="4" fill="#fff" ${O}/><rect x="34" y="12" width="26" height="36" rx="4" fill="${C.paper}" ${O}/><path d="M4 20h26M34 20h26" ${L(3)}/>` + glyph('M14 38h36M20 32l-6 6 6 6M44 32l6 6-6 6', C.blue, 4)],
    ['act-external', '在新窗口打开 ↗', '「查看执行记录 ↗」等外链', '字符 ↗', '/clean、/distribute、主页', '方框 + 斜上箭头', `<path d="M28 12H12v40h40V36" ${L(5)}/>` + glyph('M30 34L52 12M36 12h16v16', C.blue, 5)],
    ['act-focus', '只处理这几个', '「🎯 让前 3 步只处理这几个待重新处理文件」', 'emoji 🎯 / 海岛 icon-target', '/distribute、主页当前任务', '靶心', disc('#fff', 26) + `<circle cx="32" cy="32" r="18" fill="${C.coral}" ${O}/><circle cx="32" cy="32" r="10" fill="#fff" ${O}/><circle cx="32" cy="32" r="4" fill="${C.coral}" ${O}/>`],
    ['act-zoom', '看大图', '单击小图看大图、「大 / 中 / 小图」', '文字按钮', '/pdf-review、/workbench', '放大镜里一个 +', mag(26, 26, 16) + glyph('M19 26h14M26 19v14', C.blue, 3.5)],
    ['act-sort', '排序 / 顺序', '「↕ 保存工作表顺序」、上传页排序', '字符 ↕ / lucide arrow-up-down', '/clean、上传管理', '上下两个箭头', glyph('M22 50V14M12 24l10-10 10 10', C.blue, 5) + glyph('M42 14v36M32 40l10 10 10-10', C.teal, 5)],
    ['act-expand', '展开 / 收起', '「📖 程序会怎么做」折叠、下拉 ▾ ▴、看清单', '字符 ▶ ▼ ▾ ▴', '全站说明块、主页任务下拉', '圆角方块里的向下 V（转 180° 就是收起）', tile(C.sky) + glyph('M20 27l12 12 12-12', '#fff', 5)],
    ['act-filter', '筛选', '分类筛选、只看对不上的', 'lucide list-filter / 文字', '/workbench、/distribute、上传管理', '漏斗', `<path d="M8 10h48L38 32v18l-12 6V32z" fill="${C.sky}" ${O}/>`],
    ['act-add', '新增 +', '加路径、加一行', 'lucide plus / 文字', '/pcl、上传管理、新任务', '绿色圆 + 白十字', disc(C.green) + glyph('M32 19v26M19 32h26')],
    ['act-fx', '特效开关 🎉', '页头「🎉 特效」开关', 'emoji 🎉', '页头（game_fx.js）', '彩带礼炮', `<path d="M8 56l12-34 22 22z" fill="${C.gold}" ${O}/><path d="M14 40l10 10M18 30l16 16" ${L(3, C.coral)}/>` + `<path d="M34 18c4-6 10-6 12-12M42 28c6-2 12 2 16-2" ${L(3, C.purple)}/>` + star(50, 16, 5, C.pink) + `<circle cx="30" cy="10" r="3" fill="${C.teal}" ${O}/>`],
    ['act-download-file', '保存文件 💾', '保存、下载报告', 'emoji 💾', '/records、/color-repair 说明', '蓝色软盘', `<path d="M8 8h40l8 8v40H8z" fill="${C.blue}" ${O}/><rect x="16" y="8" width="26" height="16" rx="2" fill="#fff" ${O}/><rect x="34" y="11" width="5" height="10" fill="${INK}"/><rect x="14" y="34" width="36" height="20" rx="2" fill="${C.paper}" ${O}/>`],
  ];
  btn.forEach(n => add(G5, n[0], n[1], n[2], n[3], n[4], '16–20px，画 128px', n[5], n[6]));

  // ===== 6 文件和对象 =====
  const G6 = '文件和对象';
  const obj = [
    ['obj-excel', 'Excel 文件', '「📗 文件名」、打开 Excel、进度框当前文件', 'emoji 📗', '所有批量页面、进度框、对比页', '白纸 + 绿色 X 标', xlsDoc(14, 6, 38, 50)],
    ['obj-pdf', 'PDF 文件', '「📄 新版 PDF」「📕」、PDF 卡标题', 'emoji 📄 📕', '/pdf-review、/fix-images、/pdf-bundle、/distribute', '白纸 + 红色 PDF 标', pdfDoc(14, 6, 38, 50)],
    ['obj-doc', '文档 / 报告', '执行记录报告、record.md', 'emoji 📄 / 海岛 icon-document', '/records、/analysis', '白纸 + 几行字', paper(12, 6, 40, 52) + lines(20, 22, 24, 5)],
    ['obj-folder', '文件夹', '「📁 目录名」、路径输入框前', 'emoji 📁', '/pick-copy、/distribute、/pdf-bundle、/records', '黄色文件夹（合着）', folder(6, 12, 52, 42)],
    ['obj-folder-open', '打开文件夹 / 选路径', '「📂 自选路径优先」、输入框前的 📂', 'emoji 📂', '主页、/clean、/workbench、/distribute', '黄色文件夹（打开，露出纸）', folder(6, 12, 52, 42, C.gold, true).replace(`<path d="M${3}`, `${paper(16, 6, 28, 28)}<path d="M${3}`)],
    ['obj-sheet-tab', '工作表标签', '「🏷 表名」、标签颜色', 'emoji 🏷', '/workbench、/clean', '一排表标签，当前那张橙黄', `<path d="M4 20h56v20H4z" fill="#fff" ${O}/><path d="M6 40l4 14h16l4-14z" fill="${C.gray}" ${O}/><path d="M28 40l4 14h18l4-14z" fill="${C.gold}" ${O}/>` + lines(10, 28, 44, 2)],
    ['obj-image', '图片 / 截图', '「🖼 全部修复」、对比页图片', 'emoji 🖼', '/fix-images、/clean/preview、/excel-diff', '画框 + 山和太阳', `<rect x="6" y="10" width="52" height="44" rx="5" fill="${C.sky}" ${O}/><path d="M6 46l14-14 10 10 8-6 20 14v4H6z" fill="${C.green}" ${O}/><circle cx="44" cy="22" r="6" fill="${C.gold}" ${O}/>`],
    ['obj-person', '人', '「👤 某人」分人的文件夹', 'emoji 👤', '/distribute', '一个小人', person(32, 32, C.blue, 1.4)],
    ['obj-team', '几个人分着做', '「👥 几个人分着做」', 'emoji 👥', '/distribute 说明', '两个小人', person(22, 34, C.teal, 1.1) + person(42, 32, C.blue, 1.2)],
    ['obj-json', 'JSON / 拿取版本记录', '「🧾 拿取版本.json」、上传记录', 'emoji 🧾', '/distribute', '小票纸 + { }', `<path d="M14 6h36v52l-6-4-6 4-6-4-6 4-6-4-6 4z" fill="#fff7d6" ${O}/>` + txt(32, 34, '{ }', 18, C.purple) + lines(20, 44, 24, 1)],
    ['obj-hash', '核对三个数 🔢', '「🔢 核对三个数」', 'emoji 🔢', '/distribute', '蓝色方块写 123', tile(C.blue) + otxt(32, 41, '123', 17)],
    ['obj-shape', '图形 / 标注框', '修复图片里的图形、空心框', '无（文字）', '/fix-images', '红色空心框 + 箭头标注', `<rect x="8" y="10" width="34" height="26" rx="3" ${L(8)}/><rect x="8" y="10" width="34" height="26" rx="3" ${L(4, C.red)}/>` + glyph('M36 54L50 40M50 40h-9M50 40v9', C.red, 3.5)],
  ];
  obj.forEach(n => add(G6, n[0], n[1], n[2], n[3], n[4], '14–20px，画 128px', n[5], n[6]));

  // ===== 7 说明块小标题 =====
  const G7 = '说明块小标题';
  const how = [
    ['how-book', '程序会怎么做 📖', '每页「📖 程序会怎么做」折叠块的标题', 'emoji 📖', '几乎所有工具页（约 20 页）', '打开的书', `<path d="M32 16c-8-6-18-6-26-4v40c8-2 18-2 26 4c8-6 18-6 26-4V12c-8-2-18-2-26 4z" fill="#fff" ${O}/><path d="M32 16v40" ${L(3)}/><path d="M6 12v40c8-2 18-2 26 4" ${L(3)}/>` + lines(12, 24, 14, 3) + lines(38, 24, 14, 3) + `<path d="M48 12v14l-4-3-4 3V13" fill="${C.red}" ${O}/>`],
    ['how-find', '找哪些 / 怎么查 🔍', '说明块「🔍 找哪些错误」等', 'emoji 🔍 🔎', '/repair-*、/excel-diff、/workbench', '同「切字检查」放大镜，可复用 nav-cells', mag(27, 27, 15)],
    ['how-list', '怎么填 / 比哪些 📋', '说明块「📋 怎么填」「📋 比哪些」', 'emoji 📋', '/distribute、/excel-diff、/pick-copy', '剪贴板 + 几行', `<rect x="10" y="10" width="44" height="48" rx="6" fill="${C.orange}" ${O}/><rect x="16" y="16" width="32" height="36" rx="3" fill="#fff" ${O}/><rect x="22" y="5" width="20" height="10" rx="3" fill="${C.gray}" ${O}/>` + lines(21, 26, 22, 4)],
    ['how-manual', '留给人工 ✋', '说明块「✋ 留给人工的情况」', 'emoji ✋', '/repair-links、/repair-hidden、/repair-colors、/pdf-review', '举起的手', hand],
    ['how-forbid', '不碰的 🚫', '说明块「🚫 不碰的」「🚫 扫描不动的」', 'emoji 🚫', '/workbench、/fix-fonts、/pick-copy', '红圈斜杠', `<circle cx="32" cy="32" r="24" fill="#fff" ${O}/><circle cx="32" cy="32" r="20" fill="none" stroke="${C.red}" stroke-width="7"/><path d="M18 18l28 28" ${L(7, C.red)}/>`],
    ['how-why', '为什么 🤔', '说明块「🤔 为什么会坏」「🤔 为什么要换」', 'emoji 🤔', '/fix-images、/fix-fonts', '紫色圆 + 白色 ?（和「请确认」区分：不带气泡尾巴）', disc(C.purple) + otxt(32, 42, '?', 30)],
    ['how-question', '待定 ❓', '精简 Sheet 默认「待定」', 'emoji ❓', '/clean、/clean/preview', '橙色圆角方块 + 白色 ?', tile(C.orange) + otxt(32, 43, '?', 30)],
    ['how-shield', '先核对 / 覆盖前检查 🛡🔒', '说明块「🛡 覆盖前检查别人改过没有」「🔒 先核对」', 'emoji 🛡 🔒', '/distribute、/repair-hidden', '蓝色盾牌 + 勾', `<path d="M32 4l24 8v16c0 16-10 26-24 32C18 54 8 44 8 28V12z" fill="${C.blue}" ${O}/>` + check('M21 31l8 8 14-15')],
    ['how-ruler', '版面怎么定 📐 / 列宽 📏', '说明块「📐 版面怎么定」「📏 修复列宽」', 'emoji 📐 📏', '/pdf-bundle、/fix-images、主页对话框', '三角尺 + 直尺', `<path d="M8 56V12l44 44z" fill="${C.gold}" ${O}/><path d="M16 48V32l16 16z" fill="#fff8d8" ${O}/><rect x="32" y="6" width="12" height="40" rx="2" fill="${C.sky}" ${O} transform="rotate(30 38 26)"/>`],
    ['how-pages', 'PDF 里放哪些表 📑', '说明块「📑 PDF 里放哪些表」', 'emoji 📑', '/pdf-bundle', '一叠纸，侧面彩色书签', paper(14, 4, 36, 44, C.fold) + paper(10, 12, 36, 46) + `<rect x="44" y="20" width="10" height="7" rx="2" fill="${C.coral}" ${O}/><rect x="44" y="31" width="10" height="7" rx="2" fill="${C.lime}" ${O}/><rect x="44" y="42" width="10" height="7" rx="2" fill="${C.sky}" ${O}/>` + lines(16, 28, 22, 4)],
    ['how-recycle', '和上传的关系 ♻', '说明块「♻ 和上传的关系」「🔗 和精简的关系」', 'emoji ♻ 🔗', '/pdf-bundle、/repair-*', '三个绿色箭头围成三角', glyph('M32 10l10 17M50 40H30M14 40l8-14', C.green, 6) + `<path d="M46 22l-2 12-10-6z" fill="${C.green}" ${O}/><path d="M26 46l-10-6 10-6z" fill="${C.green}" ${O}/><path d="M18 28l6-10 6 10z" fill="${C.green}" ${O}/>`],
    ['how-look', '怎么看 👀', '说明块「👀 怎么看」', 'emoji 👀', '/excel-diff', '一双大眼睛', `<ellipse cx="21" cy="32" rx="13" ry="17" fill="#fff" ${O}/><ellipse cx="43" cy="32" rx="13" ry="17" fill="#fff" ${O}/><circle cx="24" cy="35" r="6" fill="${INK}"/><circle cx="46" cy="35" r="6" fill="${INK}"/><circle cx="26" cy="33" r="2" fill="#fff"/><circle cx="48" cy="33" r="2" fill="#fff"/>`],
    ['how-click', '你可以怎么改 🖱', '说明块「🖱 你可以怎么改」、Ctrl+单击提示', 'emoji 🖱', '/clean、/pdf-review', '鼠标', `<rect x="16" y="6" width="32" height="52" rx="16" fill="#fff" ${O}/><path d="M16 26h32M32 6v20" ${L(3)}/><path d="M18 24V22a14 14 0 0 1 14-14v18z" fill="${C.blue}" ${O}/>`],
    ['how-after', '点确认以后 ⚙', '说明块「⚙ 点「确认执行」以后」', 'emoji ⚙', '/clean', '同「设置」齿轮，可复用 nav-settings', `<path d="M28 6h8l1.5 6 5 2.5 5.5-3 5.5 5.5-3 5.5 2.5 5 6 1.5v8l-6 1.5-2.5 5 3 5.5-5.5 5.5-5.5-3-5 2.5L36 58h-8l-1.5-6-5-2.5-5.5 3-5.5-5.5 3-5.5L11 36l-6-1.5v-8L11 25l2.5-5-3-5.5 5.5-5.5 5.5 3 5-2.5z" fill="${C.gray}" ${O}/><circle cx="32" cy="31" r="9" fill="#fff" ${O}/>`],
    ['how-sponge', '标记色 🧽', '说明块「🧽 标记色和 Excel 本身」、清除标记', 'emoji 🧽', '/workbench、/repair-bl-links、/repair-colors、/pdf-bundle、/distribute', '黄色海绵 + 泡泡', `<rect x="8" y="22" width="44" height="30" rx="8" fill="${C.gold}" ${O}/><rect x="8" y="22" width="44" height="10" rx="5" fill="${C.green}" ${O}/><circle cx="20" cy="42" r="3" fill="${C.orange}"/><circle cx="34" cy="44" r="2.5" fill="${C.orange}"/><circle cx="42" cy="38" r="2" fill="${C.orange}"/><circle cx="48" cy="12" r="6" fill="#e8f7ff" ${O}/><circle cx="56" cy="22" r="4" fill="#e8f7ff" ${O}/>`],
    ['how-alert', '问题分几类 ❗', '说明块「❗ 查出来的问题分这几类」', 'emoji ❗', '/finish', '红色大感叹号', `<path d="M24 8h16l-3 32H27z" fill="${C.red}" ${O}/><circle cx="32" cy="50" r="6" fill="${C.red}" ${O}/>`],
    ['how-button', '每张表的按钮 🔘', '说明块「🔘 每张表的按钮」', 'emoji 🔘', '/pdf-review', '一个圆按钮被按下', `<ellipse cx="32" cy="40" rx="24" ry="12" fill="${C.deep}" ${O}/><ellipse cx="32" cy="32" rx="24" ry="12" fill="${C.blue}" ${O}/>` + shine('M16 30a16 7 0 0 1 12-6')],
    ['how-loop', '扫完以后 🔁', '说明块「🔁 扫完以后」', 'emoji 🔁', '/workbench', '同「复检」，可复用 act-recheck', glyph('M14 28a18 18 0 0 1 32-9M50 36a18 18 0 0 1-32 9', C.teal, 6) + `<path d="M48 8v14H34z" fill="${C.teal}" ${O}/><path d="M16 56V42h14z" fill="${C.teal}" ${O}/>`],
  ];
  how.forEach(n => add(G7, n[0], n[1], n[2], n[3], n[4], '18–22px，画 128px', n[5], n[6]));

  // ===== 8 分级和标记色 =====
  const G8 = '分级和标记色';
  const gem = (col, shape = 'circle') => shape === 'square'
    ? `<rect x="12" y="12" width="40" height="40" rx="9" fill="${col}" ${O}/>` + shine('M19 26v-3a4 4 0 0 1 4-4h4')
    : `<circle cx="32" cy="32" r="20" fill="${col}" ${O}/>` + shine('M20 28a13 13 0 0 1 8-9');
  const sev = [
    ['mark-red', '红：必须改 / 整张表删', '切字分级 🔴、调整页面标红', 'emoji 🔴 / CSS 色点', '/workbench、/pdf-review', '红色圆点（像宝石）', gem(C.red)],
    ['mark-yellow', '黄：可以不管 / 这一页要改', '切字分级 🟡、调整页面标黄、转人工', 'CSS 画的色点（日文 Win10 没有 🟡）', '/workbench、/pdf-review', '黄色圆点', gem(C.gold)],
    ['mark-gray', '灰：扫一眼', '切字分级 ⚪', 'emoji ⚪', '/workbench', '灰色圆点', gem(C.gray)],
    ['mark-blue', '蓝：保留 / 不用管', '三色审核蓝、调整页面蓝', 'CSS 色块', '/workbench、/pdf-review', '蓝色圆点', gem(C.blue)],
    ['mark-green', '绿：还原 / 忽略', '三色审核绿', 'CSS 色块', '/workbench', '绿色圆点', gem(C.green)],
    ['mark-orange', '橙：表标签标色', '表标签橙黄', 'CSS 色块', '/workbench', '橙色方块', gem(C.orange, 'square')],
    ['mark-combo', '一格多种问题', 'Excel 里黑底黄字的格子', 'CSS 色块', '/workbench（图例）', '黑色方块 + 黄色 ×2', `<rect x="12" y="12" width="40" height="40" rx="9" fill="${INK}" ${O}/>` + txt(32, 40, '×2', 20, C.gold)],
    ['mark-star', '★ 临时标记', '文件名前的 ★', '字符 ★', '/pdf-bundle、文件名', '金色星星', star(32, 33, 25)],
  ];
  sev.forEach(n => add(G8, n[0], n[1], n[2], n[3], n[4], '12–16px，画 128px', n[5], n[6]));

  // ===== 9 调整页面工具 =====
  const G9 = '调整页面工具';
  const pg = [
    ['page-fit-width', '宽一页', '调整页面按钮 / 分类', '文字按钮', '/pdf-review', '一页纸 + 左右顶边的箭头', pageLand(14, 6, 36, 52) + glyph('M8 32h48M14 26l-6 6 6 6M50 26l6 6-6 6', C.blue, 4)],
    ['page-portrait', '竖向', '调整页面「竖向」按钮 / 分类', '文字按钮', '/pdf-review', '竖着的纸 + 「縦」或竖箭头', pageLand(16, 6, 32, 52) + lines(22, 16, 20, 2) + otxt(32, 48, '縦', 18, C.blue)],
    ['page-landscape', '横向', '调整页面「横向」按钮 / 分类', '文字按钮', '/pdf-review', '横着的纸 + 「横」', pageLand(6, 14, 52, 36) + lines(12, 22, 40, 1) + otxt(32, 45, '横', 18, C.teal)],
    ['page-floor', '向下取整', '「🟡 末页只多一点」、取整按钮', '文字按钮 + CSS 黄点', '/pdf-review', '「4.1 → 4」，末页只剩一条', pageLand(4, 10, 24, 34) + pageLand(10, 16, 24, 34, C.fold) + `<rect x="10" y="44" width="24" height="6" fill="${C.gold}" ${O}/>` + otxt(48, 36, '4', 24, C.green) + glyph('M38 54h14', C.green, 3)],
    ['page-clear-breaks', '清手动分页', '「✂ 清手动分页」', 'emoji ✂', '/pdf-review', '纸中间一条蓝色虚线 + 剪刀', pageLand(18, 4, 40, 56) + `<path d="M18 32h40" stroke="${C.blue}" stroke-width="4" stroke-dasharray="6 4"/><circle cx="10" cy="22" r="6" fill="#fff" ${O}/><circle cx="10" cy="42" r="6" fill="#fff" ${O}/><path d="M14 26l20 10M14 38l20-10" ${L(4)}/>`],
    ['page-fixed-height', '页高定死的', '分类「页高定死的」', '文字', '/pdf-review', '纸 + 上下被锁住的箭头', pageLand(14, 10, 30, 44) + glyph('M52 12v40M46 18l6-6 6 6M46 46l6 6 6-6', C.coral, 3.5) + `<rect x="22" y="26" width="14" height="12" rx="3" fill="${C.gold}" ${O}/><path d="M25 26v-3a4 4 0 0 1 8 0v3" ${L(2.5)}/>`],
    ['page-sparse', '内容特别少的页', '复检页标黄「空页」', '文字', '/pdf-review', '一张几乎空白的纸 + 黄色 ?', pageLand(12, 6, 40, 52) + lines(18, 16, 14, 1) + otxt(38, 48, '?', 22, C.gold)],
    ['page-delete-sheet', '整张表删', '调整页面红色「删表」', 'emoji 🗑 + CSS ::before', '/pdf-review', '纸 + 红色大叉', pageLand(12, 6, 40, 52) + lines(18, 16, 28, 3) + glyph('M20 32l24 20M44 32L20 52', C.red, 5)],
  ];
  pg.forEach(n => add(G9, n[0], n[1], n[2], n[3], n[4], '20–24px，画 128px', n[5], n[6]));

  // ===== 10 查看切字：工具箱和问题分类 =====
  const G10 = '查看切字工具和分类';
  const rowIcon = label => `<rect x="6" y="14" width="40" height="36" rx="3" fill="#fff" ${O}/><rect x="6" y="24" width="40" height="16" fill="${C.mint}" ${O}/>` + glyph('M54 18v28M48 24l6-6 6 6M48 40l6 6 6-6', C.blue, 3) + otxt(26, 37, label, 12);
  const cut = [
    ['cell-row-12', '行高 ×1.2', '工具箱按钮', '文字按钮', '/workbench', '一行格子 + 上下箭头 + ×1.2', rowIcon('×1.2')],
    ['cell-row-15', '行高 ×1.5', '工具箱按钮', '文字按钮', '/workbench', '同上，×1.5', rowIcon('×1.5')],
    ['cell-row-2', '行高 ×2', '工具箱按钮', '文字按钮', '/workbench', '同上，×2', rowIcon('×2')],
    ['cell-tighten', '收紧行高', '「↕ 收紧行高并复检」', '字符 ↕', '/workbench', '一行格子 + 两个箭头往里压', `<rect x="6" y="16" width="52" height="32" rx="3" fill="#fff" ${O}/><rect x="6" y="26" width="52" height="12" fill="${C.mint}" ${O}/>` + glyph('M32 4v12M26 10l6 6 6-6M32 60V48M26 54l6-6 6 6', C.orange, 3.5)],
    ['cell-shrink-font', '缩小字体', '工具箱按钮', '文字按钮', '/workbench', '大 A → 小 a + 向下箭头', otxt(20, 42, 'A', 30, C.blue) + otxt(46, 46, 'a', 20, C.teal) + glyph('M30 18l8 8', C.orange, 3)],
    ['cell-add-break', '上面加分页线', '工具箱按钮', '文字按钮', '/workbench', '格子上面一条蓝色虚线 + 加号', `<rect x="6" y="28" width="44" height="26" rx="3" fill="#fff" ${O}/><path d="M4 22h48" stroke="${C.blue}" stroke-width="4" stroke-dasharray="6 4"/>` + disc(C.green, 9, 52, 14) + glyph('M52 9v10M47 14h10', '#fff', 2.5)],
    ['cell-keep', '不动', '工具箱「不动」', '文字按钮', '/workbench', '灰色圆 + 横杠（停一下）', disc(C.gray) + glyph('M20 32h24')],
    ['cut-strike', '压线（半截字）', '问题分类', '文字', '/workbench', '字被横线压住，只露一半', `<rect x="6" y="14" width="52" height="36" rx="3" fill="#fff" ${O}/>` + txt(32, 41, 'ABC', 18, C.blue) + `<path d="M6 35h52" ${L(4, C.red)}/>`],
    ['cut-clip', '被裁 / 被挡住', '问题分类', '文字', '/workbench', '字超出格子，右边被裁掉', `<rect x="6" y="18" width="38" height="28" fill="#fff" ${O}/>` + txt(30, 41, 'ABCD', 20, C.blue) + `<rect x="44" y="12" width="16" height="40" fill="${C.gray}" ${O}/>`],
    ['cut-half', '只印出前半段', '问题分类', '文字', '/workbench', '一行字后面变成虚线', `<rect x="4" y="18" width="56" height="28" rx="3" fill="#fff" ${O}/>` + txt(20, 40, 'AB', 18, C.blue) + `<path d="M34 34h22" stroke="${C.dgray}" stroke-width="4" stroke-dasharray="3 4"/>`],
    ['cut-hash', '###（列太窄）', '问题分类', '文字', '/workbench', '格子里写 ###', `<rect x="6" y="16" width="52" height="32" rx="3" fill="#fff" ${O}/>` + txt(32, 41, '###', 22, C.red)],
    ['cut-page-split', '被分页切开', '问题分类', '文字', '/workbench', '一格被蓝色分页虚线从中间切开', `<rect x="10" y="10" width="44" height="44" rx="3" fill="#fff" ${O}/>` + txt(32, 40, 'AB', 22, C.blue) + `<path d="M4 34h56" stroke="${C.blue}" stroke-width="4" stroke-dasharray="6 4"/>`],
    ['cut-border', '边框压线', '问题分类', '文字', '/workbench', '字贴着粗边框', `<rect x="8" y="14" width="48" height="36" fill="#fff" ${O}/><path d="M8 44h48" ${L(6)}/>` + txt(32, 46, 'Ag', 22, C.blue)],
    ['cut-horizontal', '跨格横排', '问题分类（旧标志）', '文字', '/workbench', '一行字跨过好几格', `<rect x="4" y="18" width="56" height="28" fill="#fff" ${O}/><path d="M22 18v28M40 18v28" ${L(2, '#9fb3cc')}/>` + txt(32, 39, 'ABCDE', 15, C.blue)],
  ];
  cut.forEach(n => add(G10, n[0], n[1], n[2], n[3], n[4], '20–24px，画 128px', n[5], n[6]));

  // ===== 11 进度和动效 =====
  const G11 = '进度和动效';
  const fx = [
    ['fx-spinner', '转圈（加载中）', '文件正在打开、导出中、加载中', 'CSS 动画（batch-file-spin、clean-spin、fin-spin、toast-spin）', '/workbench、/clean、/finish、提示条、/excel-diff/all', '一圈渐变的弧（做成 8 帧也行）', spin()],
    ['fx-hourglass', '沙漏 ⏳', '进度框「⏳」、正在导出、等待中', 'emoji ⏳', '所有批量进度框、/pdf-review、/clean', '金色沙漏', `<path d="M14 6h36M14 58h36" ${L(5)}/><path d="M18 8c0 14 14 18 14 24S18 42 18 56h28c0-14-14-18-14-24s14-10 14-24z" fill="#e8f7ff" ${O}/><path d="M24 50c2-6 8-8 8-12 0 4 6 6 8 12z" fill="${C.gold}" ${O}/><path d="M24 14h16l-8 8z" fill="${C.gold}" ${O}/>`],
    ['fx-timer', '计时 ⏱', '「⏱ 预计完成」、超时跳过', 'emoji ⏱', '所有进度条下的预计时间（eta.js）', '秒表', disc('#fff', 23, 32, 36) + `<rect x="26" y="4" width="12" height="7" rx="2" fill="${C.coral}" ${O}/><path d="M32 36V22" ${L(4)}/><path d="M32 36l8 6" ${L(4, C.coral)}/>`],
    ['fx-progress-track', '进度条底槽', '所有进度条的底', 'CSS 渐变', '所有批量页面', '圆角长条凹槽（做成可横向拉伸：左中右三段）', `<rect x="2" y="24" width="60" height="16" rx="8" fill="#dbe6f5" ${O}/><rect x="6" y="28" width="52" height="4" rx="2" fill="rgba(23,50,90,.15)"/>`],
    ['fx-progress-fill', '进度条填充', '所有进度条的已完成部分', 'CSS 渐变 + 斜纹动画（fx-stripes）', '所有批量页面', '亮绿色圆角长条 + 斜纹（三段可拉伸）', `<rect x="2" y="24" width="60" height="16" rx="8" fill="${C.lime}" ${O}/>` + [10, 22, 34, 46].map(x => `<path d="M${x} 38l8-12" ${L(4, 'rgba(255,255,255,.55)')}/>`).join('') + shine('M8 29h48')],
    ['fx-confetti', '撒花碎片', '完成时撒花（game_fx.js 画布）', 'canvas 画的彩纸', '完成任务时全站', '一组彩纸：长条、圆片、三角、波浪、星（各一格）', `<rect x="6" y="8" width="8" height="18" rx="2" fill="${C.coral}" ${O} transform="rotate(-20 10 17)"/><circle cx="40" cy="14" r="6" fill="${C.teal}" ${O}/><path d="M12 44l10-8 4 12z" fill="${C.gold}" ${O}/><path d="M34 40c4-4 8 4 12 0s8 4 12 0" ${L(4, C.purple)}/>` + star(46, 54, 7, C.pink)],
    ['fx-star', '星星 / 闪光', '完成横幅上的星（fx-star-pop）、闪光', 'CSS + 字符 ★', '完成时全站', '金色胖星星', star(32, 34, 26)],
    ['fx-ribbon', '完成横幅丝带', '完成时屏幕中间的横幅（fx-banner-ribbon）', 'CSS 画的丝带', '完成任务时全站', '红色丝带横幅，两头燕尾（中间留空放字，可九宫格拉伸）', `<path d="M2 24h12v20H2l5-10z" fill="${C.red}" ${O}/><path d="M62 24H50v20h12l-5-10z" fill="${C.red}" ${O}/><rect x="10" y="18" width="44" height="22" rx="3" fill="${C.coral}" ${O}/>` + shine('M15 23h34')],
    ['fx-sparkle', '闪光 ✨', '按钮高光、新功能提示', 'CSS 动画（btn-sheen、fx-sheen）/ lucide sparkles', '主页、上传管理', '两颗四角星', `<path d="M24 6c2 12 6 16 18 18-12 2-16 6-18 18-2-12-6-16-18-18 12-2 16-6 18-18z" fill="${C.gold}" ${O}/><path d="M48 36c1 7 3 9 10 10-7 1-9 3-10 10-1-7-3-9-10-10 7-1 9-3 10-10z" fill="#fff3b0" ${O}/>`],
  ];
  fx.forEach(n => add(G11, n[0], n[1], n[2], n[3], n[4], n[0].startsWith('fx-progress') || n[0] === 'fx-ribbon' ? '按需拉伸，画 256×64 再切三段' : '24–48px，画 128px', n[5], n[6]));

  // ===== 12 空状态插画 =====
  const G12 = '空状态插画';
  const empty = [
    ['empty-no-folder', '还没选文件夹', '页面一打开、没填路径时', '只有文字', '/clean、/workbench、各修复页', '空文件夹 + 问号', folder(8, 14, 48, 40, C.gold, true) + otxt(32, 50, '?', 22)],
    ['empty-no-scan', '尚未扫描', '查看切字没扫过、主页卡「尚未扫描」', '只有文字', '主页、/workbench', '放大镜 + 一张白表 + zZ', sheet(6, 14, 38, 36) + mag(44, 30, 10) + otxt(52, 14, 'z', 14, C.sky)],
    ['empty-all-done', '全部处理完了', '卡片都审完、没有问题', '只有文字（✓ 提示）', '/workbench、/pdf-review、/fix-images', '奖杯/大勾 + 撒花', disc(C.green, 22, 32, 36) + check('M22 37l7 7 13-14') + star(10, 12, 6, C.gold) + star(54, 10, 5, C.pink) + `<circle cx="52" cy="54" r="3" fill="${C.teal}" ${O}/>`],
    ['empty-waiting', '等第一份出来', '「⏳ 正在导出 PDF，第一份导好就会出现」', 'emoji ⏳ + 文字', '/pdf-review、/workbench、/excel-diff/all', '传送带上一份纸 + 沙漏', `<rect x="4" y="44" width="56" height="10" rx="5" fill="${C.gray}" ${O}/><circle cx="12" cy="49" r="2.5" fill="${INK}"/><circle cx="32" cy="49" r="2.5" fill="${INK}"/><circle cx="52" cy="49" r="2.5" fill="${INK}"/>` + paper(10, 12, 24, 30) + `<g transform="translate(34 6) scale(.45)"><path d="M14 6h36M14 58h36" ${L(5)}/><path d="M18 8c0 14 14 18 14 24S18 42 18 56h28c0-14-14-18-14-24s14-10 14-24z" fill="#e8f7ff" ${O}/><path d="M24 50c2-6 8-8 8-12 0 4 6 6 8 12z" fill="${C.gold}" ${O}/></g>`],
    ['empty-no-task', '还没有任务', '「还没有任务，请回主页点新任务」', '只有文字', '/finish、主页任务条、/distribute', '空的宝箱，盖子开着', `<path d="M8 30h48v24H8z" fill="${C.blue}" ${O}/><path d="M10 30L4 12h40l12 18z" fill="${C.deep}" ${O}/><path d="M8 40h48" ${L(3)}/><rect x="26" y="34" width="12" height="10" rx="2" fill="${C.gold}" ${O}/>`],
    ['empty-failed', '没加载出来', '「这一页没加载出来，刷新页面再试」', 'CSS ::after 文字', '/clean/preview、/pdf-review 小图', '一张破了角的纸 + 刷新箭头', paper(12, 6, 40, 52) + `<path d="M12 40l8-4 6 6 8-6 6 6 12-4v20H12z" fill="#fff" ${O}/>` + glyph('M22 24a10 10 0 0 1 18-4M40 14v7h-7', C.blue, 3.5)],
  ];
  empty.forEach(n => add(G12, n[0], n[1], n[2], n[3], n[4], '120–200px，画 256px', n[5], n[6]));

  // ===== 13 上传管理页（现在用 lucide 线条图标）=====
  const G13 = '上传管理页（可选）';
  const um = [
    ['um-history', '历史', '侧栏「历史」', 'lucide history', '上传管理 /static/upload_manager', '时钟 + 逆时针箭头', disc('#fff', 22) + `<path d="M32 20v12l8 5" ${L(4)}/>` + `<path d="M8 12v12h12z" fill="${C.blue}" ${O}/>`],
    ['um-dashboard', '总览', '侧栏「总览」', 'lucide layout-dashboard', '上传管理', '四块不等大的面板', `<rect x="6" y="6" width="24" height="30" rx="5" fill="${C.blue}" ${O}/><rect x="34" y="6" width="24" height="18" rx="5" fill="${C.teal}" ${O}/><rect x="6" y="40" width="24" height="18" rx="5" fill="${C.gold}" ${O}/><rect x="34" y="28" width="24" height="30" rx="5" fill="${C.coral}" ${O}/>`],
    ['um-drive', '硬盘 / 本机', '来源是本机硬盘', 'lucide hard-drive', '上传管理', '灰色硬盘盒 + 绿灯', `<rect x="6" y="20" width="52" height="26" rx="6" fill="${C.gray}" ${O}/><circle cx="46" cy="33" r="3.5" fill="${C.lime}" ${O}/><path d="M14 33h20" ${L(3)}/>`],
    ['um-folder-search', '找文件夹', '「选文件夹」', 'lucide folder-search', '上传管理', '文件夹 + 放大镜', folder(4, 10, 44, 38) + mag(42, 40, 9)],
    ['um-files', '多个文件', '文件清单', 'lucide files', '上传管理', '两张纸叠', paper(20, 6, 32, 42, C.fold) + paper(10, 14, 32, 44)],
    ['um-table-import', '导入表格', '导入 Excel 清单', 'lucide table-import', '上传管理', '表格 + 向里箭头', sheet(22, 10, 36, 44) + fatArrow('right', C.lime, 14, 32, 0.55)],
    ['um-table-export', '导出表格', '导出 Excel 清单', 'lucide table-export', '上传管理', '表格 + 向外箭头', sheet(6, 10, 36, 44) + fatArrow('right', C.orange, 50, 32, 0.55)],
    ['um-json-import', '导入 JSON', '导入上传记录', 'lucide file-json-import', '上传管理', '{ } 纸 + 向里箭头', paper(24, 6, 32, 50, '#fff7d6') + txt(40, 38, '{ }', 14, C.purple) + fatArrow('right', C.lime, 14, 32, 0.5)],
    ['um-json-export', '导出 JSON', '导出上传记录', 'lucide file-json-export', '上传管理', '{ } 纸 + 向外箭头', paper(6, 6, 32, 50, '#fff7d6') + txt(22, 38, '{ }', 14, C.purple) + fatArrow('right', C.orange, 50, 32, 0.5)],
    ['um-contact', '负责人', '负责人一栏', 'lucide contact-round', '上传管理', '名片 + 小人', `<rect x="4" y="12" width="56" height="40" rx="6" fill="#fff" ${O}/>` + person(20, 34, C.blue, 0.8) + lines(32, 28, 20, 3)],
    ['um-calendar', '日期', '日期选择、计划', 'lucide calendar-days / calendar-check', '上传管理', '日历', `<rect x="6" y="10" width="52" height="46" rx="6" fill="#fff" ${O}/><path d="M6 16a6 6 0 0 1 6-6h40a6 6 0 0 1 6 6v8H6z" fill="${C.coral}" ${O}/><path d="M20 4v10M44 4v10" ${L(4)}/>` + check('M22 39l7 7 13-13', C.green)],
    ['um-pie', '统计饼图', '进度统计', 'lucide chart-pie', '上传管理', '彩色饼图', `<circle cx="32" cy="32" r="24" fill="${C.sky}" ${O}/><path d="M32 32V8a24 24 0 0 1 22 14z" fill="${C.gold}" ${O}/><path d="M32 32l22-10a24 24 0 0 1-8 30z" fill="${C.coral}" ${O}/>`],
    ['um-collapse', '全部展开 / 收起', '「全部展开」「全部收起」', 'lucide chevrons-up-down / chevrons-down-up', '上传管理', '两个 V 上下相对', glyph('M20 22l12-10 12 10M20 42l12 10 12-10', C.blue, 5)],
  ];
  um.forEach(n => add(G13, n[0], n[1], n[2], n[3], n[4], '16px，画 128px', n[5], n[6]));

  const SYM = /\s*[\u2190-\u2BFF\u2713-\u2718\u{1F000}-\u{1FAFF}]\uFE0F?/gu;
  I.forEach((ic, i) => { ic.no = i + 1; ic.name = ic.name.replace(SYM, '').replace(/\s+$/, ''); });
  const groups = [...new Set(I.map(i => i.g))];
  const api = { icons: I, groups, INK };
  if (typeof module !== 'undefined') module.exports = api; else window.SpriteIcons = api;
})();
