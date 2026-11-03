// 进度条下面的预估时间：「14:08 开始 → 预计 14:20 完成 · 还要约 12 分钟」
// 按这次已用时间和已走的进度推算（匀速），进度每变一次重算一次；两次之间完成时刻不变，「还要」自己倒数。
// 起点存在 localStorage，刷新页面、换页面看同一个任务都接着算；任务结束（stop）就清掉。
(function () {
  const PREFIX = 'clear_handle_excel.eta.';
  const pad = n => String(n).padStart(2, '0');

  function clock(ms) {
    const d = new Date(ms), today = new Date();
    const hm = `${pad(d.getHours())}:${pad(d.getMinutes())}`;
    if (d.toDateString() === today.toDateString()) return hm;
    const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
    return d.toDateString() === tomorrow.toDateString() ? `明天 ${hm}` : `${d.getMonth() + 1}/${d.getDate()} ${hm}`;
  }

  function left(seconds) {
    const m = Math.ceil(seconds / 60);
    if (seconds < 60) return '不到 1 分钟';
    if (m < 60) return `约 ${m} 分钟`;
    const h = Math.floor(m / 60), rest = m % 60;
    return `约 ${h} 小时${rest ? ` ${rest} 分钟` : ''}`;
  }

  function load(key) {
    try { return JSON.parse(localStorage.getItem(PREFIX + key)) || null; } catch (e) { return null; }
  }
  function save(key, rec) {
    try { localStorage.setItem(PREFIX + key, JSON.stringify(rec)); } catch (e) { /* 存不了就只在本页算 */ }
  }
  function clear(key) {
    try { localStorage.removeItem(PREFIX + key); } catch (e) { /* 忽略 */ }
  }

  // 记一次进度，返回 { start, finish, remaining }（毫秒 / 秒）；还算不出来时 finish 为 null
  // remainingSeconds：后台自己估好的剩余秒数（切字扫描按文件大小估），给了就用它
  function estimate(key, fraction, remainingSeconds) {
    const now = Date.now();
    const p = Math.min(1, Math.max(0, Number(fraction) || 0));
    let rec = load(key);
    // 进度往回走（新开了一次任务）、超过 10 分钟没看过（可能是上次没收尾的旧记录）就重新开始算
    if (!rec || p + 1e-9 < rec.p || now - (rec.seen || rec.t0) > 600000) {
      rec = { t0: now, p0: p, p, tp: now, start: p <= 0.02 ? now : null };
    } else if (p > rec.p) {
      rec.p = p; rec.tp = now;
    }
    rec.seen = now;
    save(key, rec);
    if (remainingSeconds != null && Number.isFinite(Number(remainingSeconds))) {
      const r = Math.max(0, Number(remainingSeconds));
      return { start: rec.start, finish: now + r * 1000, remaining: r };
    }
    const moved = rec.p - rec.p0, used = rec.tp - rec.t0;
    if (moved <= 0 || used < 1500) return { start: rec.start, finish: null, remaining: null };
    // 上次进度变化时推出的完成时刻；卡在一个大文件上超过了它，就按到现在的速度重算（往后推）
    let finish = rec.tp + (1 - rec.p) * used / moved;
    if (finish < now) finish = now + (1 - rec.p) * (now - rec.t0) / moved;
    // 中途才打开页面时不知道几点开始的：按同样的速度往回推一个「约几点开始」
    const guessed = rec.start == null ? rec.t0 - rec.p0 * used / moved : null;
    return { start: rec.start ?? guessed, guessed: guessed != null, finish, remaining: (finish - now) / 1000 };
  }

  function text(key, fraction, remainingSeconds) {
    const e = estimate(key, fraction, remainingSeconds);
    const from = e.start ? `${e.guessed ? '约 ' : ''}${clock(e.start)} 开始 → ` : '';
    if (e.finish == null) return `⏱ ${from}正在估算完成时间（走完第一份后算出）…`;
    return `⏱ ${from}预计 ${clock(e.finish)} 完成 · 还要${left(e.remaining)}`;
  }

  // 在进度条下面放一行预估时间（只建一次，跑着的时候固定占一行，文字变化不挪页面）
  function line(fill) {
    const track = fill && fill.closest && fill.closest('.progress-track');
    if (!track) return null;
    // 主页任务条：进度条下面已经有「12.5% · 5/40 个文件」，时间放在它后面
    const anchor = track.parentElement && track.parentElement.classList.contains('home-strip-progress')
      ? track.parentElement.lastElementChild : track;
    let el = anchor.classList.contains('progress-eta') ? anchor : anchor.nextElementSibling;
    if (!el || !el.classList.contains('progress-eta')) {
      el = document.createElement('p');
      el.className = 'progress-eta';
      el.setAttribute('aria-live', 'polite');
      if (!anchor.after) return null;
      anchor.after(el);
    }
    return el;
  }

  // 每次刷新进度时调用：running 为 false 时收起这一行并清掉记录
  function update(fillOrId, key, fraction, running, remainingSeconds) {
    const fill = typeof fillOrId === 'string' ? document.getElementById(fillOrId) : fillOrId;
    const el = line(fill);
    if (!running) {
      clear(key);
      if (el) { el.hidden = true; el.textContent = ''; }
      return;
    }
    const value = text(key, fraction, remainingSeconds);
    if (el) { el.hidden = false; if (el.textContent !== value) el.textContent = value; }
    return value;
  }

  window.ProgressEta = { update, text, stop: clear, clock, left };
})();
