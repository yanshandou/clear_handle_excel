'use strict';
// 切字标色与独立颜色修复共用配色，浏览器中的历史记录保持兼容。
const HIGHLIGHT_KEY = 'clear_handle_excel.highlightColors';
const HIGHLIGHT_HISTORY_KEY = 'clear_handle_excel.highlightColorHistory';
const HIGHLIGHT_PRESETS = [
  { label: '洋红底白字（推荐）', fill: '#C0007A', font: '#FFFFFF' },
  { label: '宝蓝底白字', fill: '#1649CC', font: '#FFFFFF' },
  { label: '墨绿底白字', fill: '#006633', font: '#FFFFFF' },
  { label: '深青底白字', fill: '#006A71', font: '#FFFFFF' },
  { label: '深紫底白字', fill: '#6D28D9', font: '#FFFFFF' },
  { label: '橙红底白字', fill: '#C2410C', font: '#FFFFFF' },
  { label: '深灰底白字', fill: '#333333', font: '#FFFFFF' },
  { label: '纯黑底黄字', fill: '#000000', font: '#FFFF00' },
];
// 切字工作台按问题类型分色（用户 2026-11-03 定）：默认色和 core/batch.py DEFAULT_TYPE_HIGHLIGHT 一致，都是上面的预设色
const TYPE_HIGHLIGHT_KEY = 'clear_handle_excel.typeHighlightColors';
const TYPE_HIGHLIGHTS = [
  { type: 'border_clip', label: '边框压线', fill: '#C0007A', font: '#FFFFFF' },
  { type: 'text_truncated', label: '文字截断', fill: '#1649CC', font: '#FFFFFF' },
  { type: 'box_overflow', label: '超出裁剪框', fill: '#006633', font: '#FFFFFF' },
  { type: 'page_break_cut', label: '分页切开', fill: '#C2410C', font: '#FFFFFF' },
  { type: 'hash_overflow', label: '### 列宽不足', fill: '#6D28D9', font: '#FFFFFF' },
  { type: 'text_missing', label: 'PDF 未匹配', fill: '#333333', font: '#FFFFFF' },
  { type: 'multi', label: '多种问题', fill: '#000000', font: '#FFFF00' },
];
const HEX = /^#[0-9a-f]{6}$/i;
const sameHighlight = (a, b) => a.fill === b.fill && a.font === b.font;
function normalizeHighlight(color) {
  return typeof color?.fill === 'string' && typeof color?.font === 'string' && HEX.test(color.fill) && HEX.test(color.font)
    ? { fill: color.fill.toUpperCase(), font: color.font.toUpperCase() } : null;
}
function uniqueHighlights(colors) {
  const seen = new Set();
  return colors.map(normalizeHighlight).filter(color => {
    if (!color) return false;
    const key = `${color.fill}/${color.font}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
function savedHighlightColors() {
  let current, history;
  try { current = JSON.parse(localStorage.getItem(HIGHLIGHT_KEY)); } catch (_) {}
  try { history = JSON.parse(localStorage.getItem(HIGHLIGHT_HISTORY_KEY)); } catch (_) {}
  let types;
  try { types = JSON.parse(localStorage.getItem(TYPE_HIGHLIGHT_KEY)); } catch (_) {}
  const typeColors = types && typeof types === 'object' ? Object.values(types) : [];
  return uniqueHighlights([...HIGHLIGHT_PRESETS, current, ...typeColors, ...(Array.isArray(history) ? history : [])]);
}
