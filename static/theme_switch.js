'use strict';
// 界面风格切换：贴纸风（theme.css）/ 黏土风（theme-clay.css），只换样式表不改功能。
// 选择记在浏览器 localStorage 里，各页面共用；其他标签页切换时这里跟着换。
(function () {
  const KEY = 'clear_handle_excel.theme';
  const THEMES = {
    sticker: { css: '/static/theme.css', icon: '/static/favicon.svg', name: '贴纸风' },
    clay: { css: '/static/theme-clay.css', icon: '/static/icons/clay/app.svg', name: '黏土风' },
  };
  const link = document.getElementById('themeCss');
  let current = 'sticker';
  try { if (localStorage.getItem(KEY) === 'clay') current = 'clay'; } catch (e) { /* 读不到就用默认 */ }

  function apply(name) {
    current = THEMES[name] ? name : 'sticker';
    if (link) link.href = THEMES[current].css;
    const icon = document.querySelector('link[rel="icon"]');
    if (icon) icon.href = THEMES[current].icon;
    const button = document.getElementById('themeToggleBtn');
    if (button) button.textContent = `🎨 换成${THEMES[current === 'clay' ? 'sticker' : 'clay'].name}`;
  }

  // 页面还没画出来就换好样式表，不闪一下默认风格
  if (current !== 'sticker') apply(current);

  document.addEventListener('DOMContentLoaded', () => {
    const nav = document.querySelector('.header-nav');
    if (nav) {
      const button = document.createElement('button');
      button.type = 'button';
      button.id = 'themeToggleBtn';
      button.className = 'nav-link theme-toggle';
      button.title = '切换界面风格（只改外观，选择会记在这台电脑的浏览器里）';
      button.addEventListener('click', () => {
        const next = current === 'clay' ? 'sticker' : 'clay';
        try { localStorage.setItem(KEY, next); } catch (e) { /* 存不了就只对当前页生效 */ }
        apply(next);
      });
      const status = document.getElementById('statusText');
      nav.insertBefore(button, status && status.parentElement === nav ? status : null);
    }
    apply(current);
  });

  window.addEventListener('storage', event => {
    if (event.key === KEY) apply(event.newValue === 'clay' ? 'clay' : 'sticker');
  });
})();
