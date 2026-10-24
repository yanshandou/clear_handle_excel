'use strict';
// 游戏感特效：任务完成撒花 + 「任务完成」缎带横幅、按钮点击星星。
// 只做外观，不挡操作（全部 pointer-events: none），播放完自动移除；顶栏「🎉 特效」开关可关掉，选择记在 localStorage。
// 系统设置了「减少动画」时不撒花，只显示横幅。
(function () {
  const FX_KEY = 'clear_handle_excel.fx';
  // 以前有三套主题可切换，现在只剩游戏风：清掉旧选择，免得留着无用的值
  try { localStorage.removeItem('clear_handle_excel.theme'); } catch (_) {}

  const COLORS = ['#4fbf1f', '#7ee04a', '#2f80ff', '#5aa6ff', '#ffc21f', '#ffe066', '#ff8a1f', '#7c5cf0', '#ec2049'];
  const reduceMotion = () => window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  function enabled() {
    try { return localStorage.getItem(FX_KEY) !== '0'; } catch (_) { return true; }
  }

  /* ---------- 撒花：整页一块透明画布，约 2 秒后移除 ---------- */
  let confettiCanvas = null;
  let confettiFrame = 0;

  function confetti() {
    if (confettiCanvas) { cancelAnimationFrame(confettiFrame); confettiCanvas.remove(); }
    const canvas = confettiCanvas = document.createElement('canvas');
    canvas.className = 'fx-confetti';
    canvas.setAttribute('aria-hidden', 'true');
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    const width = window.innerWidth;
    const height = window.innerHeight;
    canvas.width = width * dpr;
    canvas.height = height * dpr;
    document.body.appendChild(canvas);
    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);

    // 从左右两个下角斜着往上喷
    const pieces = [];
    for (let i = 0; i < 150; i++) {
      const left = i % 2 === 0;
      const angle = (left ? -60 : -120) + (Math.random() - 0.5) * 40;
      const speed = 11 + Math.random() * 9;
      pieces.push({
        x: left ? -10 : width + 10,
        y: height * 0.85,
        vx: Math.cos(angle * Math.PI / 180) * speed,
        vy: Math.sin(angle * Math.PI / 180) * speed,
        size: 7 + Math.random() * 7,
        spin: Math.random() * Math.PI,
        vspin: (Math.random() - 0.5) * 0.35,
        color: COLORS[i % COLORS.length],
        star: i % 6 === 0,
      });
    }

    const start = performance.now();
    const DURATION = 2200;
    function star(size) {
      ctx.beginPath();
      for (let k = 0; k < 10; k++) {
        const r = k % 2 ? size * 0.42 : size;
        const a = -Math.PI / 2 + k * Math.PI / 5;
        ctx.lineTo(Math.cos(a) * r, Math.sin(a) * r);
      }
      ctx.closePath();
      ctx.fill();
      ctx.stroke();
    }
    function frame(now) {
      const t = now - start;
      ctx.clearRect(0, 0, width, height);
      ctx.globalAlpha = t > DURATION - 500 ? Math.max(0, (DURATION - t) / 500) : 1;
      ctx.lineWidth = 1.5;
      ctx.strokeStyle = '#1c2858';
      for (const p of pieces) {
        p.vy += 0.32;
        p.vx *= 0.99;
        p.x += p.vx;
        p.y += p.vy;
        p.spin += p.vspin;
        ctx.save();
        ctx.translate(p.x, p.y);
        ctx.rotate(p.spin);
        ctx.fillStyle = p.color;
        if (p.star) star(p.size * 0.8);
        else {
          // 纸片翻转：宽度随角度变化
          const w = p.size * Math.abs(Math.cos(p.spin * 2));
          ctx.fillRect(-w / 2, -p.size * 0.3, w, p.size * 0.6);
        }
        ctx.restore();
      }
      if (t < DURATION) confettiFrame = requestAnimationFrame(frame);
      else { canvas.remove(); if (confettiCanvas === canvas) confettiCanvas = null; }
    }
    confettiFrame = requestAnimationFrame(frame);
  }

  /* ---------- 横幅：缎带标题 + 三颗星依次弹出 ---------- */
  let bannerTimer = 0;

  function banner(title, sub, stars) {
    document.querySelectorAll('.fx-banner').forEach(el => el.remove());
    clearTimeout(bannerTimer);
    const el = document.createElement('div');
    el.className = 'fx-banner';
    el.setAttribute('role', 'status');
    const count = Math.max(0, Math.min(3, stars ?? 3));
    el.innerHTML = '<div class="fx-banner-stars" aria-hidden="true">'
      + [0, 1, 2].map(i => `<span class="fx-star${i < count ? ' is-on' : ''}" style="--i:${i}"></span>`).join('')
      + '</div><div class="fx-banner-ribbon"></div><div class="fx-banner-sub"></div>';
    el.querySelector('.fx-banner-ribbon').textContent = title;
    const subEl = el.querySelector('.fx-banner-sub');
    if (sub) subEl.textContent = sub; else subEl.remove();
    document.body.appendChild(el);
    bannerTimer = setTimeout(() => {
      el.classList.add('is-leaving');
      setTimeout(() => el.remove(), 400);
    }, 2600);
  }

  /* ---------- 公开接口 ---------- */
  // 任务完成：撒花 + 横幅。stars 用来表示结果好坏（3 = 全部成功）
  function celebrate(title = '任务完成！', sub = '', stars = 3) {
    if (!enabled() || !document.body) return;
    banner(title, sub, stars);
    if (!reduceMotion()) confetti();
  }

  // 点击位置冒几颗小星星，0.5 秒后消失
  function sparkle(x, y) {
    if (!enabled() || reduceMotion() || !document.body) return;
    for (let i = 0; i < 8; i++) {
      const dot = document.createElement('span');
      dot.className = 'fx-spark';
      dot.style.left = `${x}px`;
      dot.style.top = `${y}px`;
      dot.style.background = COLORS[(i * 2) % COLORS.length];
      document.body.appendChild(dot);
      const angle = (i / 8) * Math.PI * 2 + Math.random() * 0.4;
      const dist = 26 + Math.random() * 18;
      const anim = dot.animate([
        { transform: 'translate(-50%, -50%) scale(1)', opacity: 1 },
        { transform: `translate(calc(-50% + ${Math.cos(angle) * dist}px), calc(-50% + ${Math.sin(angle) * dist}px)) scale(0.2)`, opacity: 0 },
      ], { duration: 480, easing: 'cubic-bezier(.2,.7,.3,1)' });
      anim.onfinish = () => dot.remove();
    }
  }

  window.gameFx = { celebrate, sparkle, enabled };

  document.addEventListener('pointerdown', event => {
    if (event.button !== 0) return;
    const target = event.target.closest('.btn-primary, .btn-success, .feature-card');
    if (!target || target.disabled || target.getAttribute('aria-disabled') === 'true') return;
    sparkle(event.clientX, event.clientY);
  });

  /* ---------- 顶栏开关 ---------- */
  document.addEventListener('DOMContentLoaded', () => {
    const nav = document.querySelector('.header-nav');
    if (!nav) return;
    const label = document.createElement('label');
    label.className = 'nav-link fx-toggle';
    label.title = '完成任务时撒花、点按钮冒星星（只是外观，选择记在这台电脑的浏览器里）';
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = enabled();
    box.setAttribute('aria-label', '游戏特效');
    box.addEventListener('change', () => {
      try { localStorage.setItem(FX_KEY, box.checked ? '1' : '0'); } catch (_) {}
      if (box.checked) celebrate('特效已开启', '完成任务时会撒花庆祝');
    });
    label.append(box, '🎉 特效');
    const status = document.getElementById('statusText');
    nav.insertBefore(label, status && status.parentElement === nav ? status : null);
  });

  window.addEventListener('storage', event => {
    const box = document.querySelector('.fx-toggle input');
    if (event.key === FX_KEY && box) box.checked = enabled();
  });
})();
