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
  return uniqueHighlights([...HIGHLIGHT_PRESETS, current, ...(Array.isArray(history) ? history : [])]);
}
