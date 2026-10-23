'use strict';
// 界面风格切换：贴纸风（theme.css）/ 黏土风（theme-clay.css）/ 电竞风（theme-esports.css），只换样式表不改功能。
// 选择记在浏览器 localStorage 里，各页面共用；其他标签页切换时这里跟着换。
(function () {
  const KEY = 'clear_handle_excel.theme';
  const THEMES = {
    sticker: { css: '/static/theme.css', icon: '/static/favicon.svg', name: '贴纸风' },
    clay: { css: '/static/theme-clay.css', icon: '/static/icons/clay/app.svg', name: '黏土风' },
    esports: { css: '/static/theme-esports.css', icon: '/static/icons/esports/app.svg', name: '电竞风' },
  };
  const link = document.getElementById('themeCss');
  let current = 'sticker';
  try {
    const saved = localStorage.getItem(KEY);
    if (THEMES[saved]) current = saved;
  } catch (e) { /* 读不到就用默认 */ }

  function apply(name) {
    current = THEMES[name] ? name : 'sticker';
    if (link) link.href = THEMES[current].css;
    const icon = document.querySelector('link[rel="icon"]');
    if (icon) icon.href = THEMES[current].icon;
    const select = document.getElementById('themeSelect');
    if (select) select.value = current;
  }

  // 页面还没画出来就换好样式表，不闪一下默认风格
  if (current !== 'sticker') apply(current);

  document.addEventListener('DOMContentLoaded', () => {
    const nav = document.querySelector('.header-nav');
    if (nav) {
      // 下拉选择风格：外层沿用导航链接的样式
      const label = document.createElement('label');
      label.className = 'nav-link theme-toggle';
      label.title = '切换界面风格（只改外观，选择会记在这台电脑的浏览器里）';
      const select = document.createElement('select');
      select.id = 'themeSelect';
      select.setAttribute('aria-label', '界面风格');
      for (const [value, theme] of Object.entries(THEMES)) select.append(new Option(theme.name, value));
      select.addEventListener('change', () => {
        try { localStorage.setItem(KEY, select.value); } catch (e) { /* 存不了就只对当前页生效 */ }
        apply(select.value);
      });
      label.append('🎨 ', select);
      const status = document.getElementById('statusText');
      nav.insertBefore(label, status && status.parentElement === nav ? status : null);
    }
    apply(current);
  });

  window.addEventListener('storage', event => {
    if (event.key === KEY) apply(THEMES[event.newValue] ? event.newValue : 'sticker');
  });
})();
