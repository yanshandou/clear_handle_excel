/* ==========================================
   站内弹窗：替代浏览器原生 alert / confirm
   appAlert(message, options)   → Promise<void>
   appConfirm(message, options) → Promise<boolean>（取消、Esc 都算 false）
   options: title, okText, cancelText, tone（info | warn | error | danger）
   多个弹窗同时请求时排队，一次只显示一个；文字按纯文本显示，\n 换行。
   ========================================== */
(function () {
  const TONES = {
    info: { icon: 'ℹ️', alertTitle: '提示', confirmTitle: '请确认' },
    warn: { icon: '⚠️', alertTitle: '注意', confirmTitle: '请确认' },
    error: { icon: '⛔', alertTitle: '出错了', confirmTitle: '请确认' },
    danger: { icon: '⚠️', alertTitle: '注意', confirmTitle: '请确认' },
  };
  let queue = Promise.resolve();
  let dialog = null;

  function build() {
    dialog = document.createElement('dialog');
    dialog.className = 'app-dialog';
    dialog.setAttribute('aria-labelledby', 'appDialogTitle');
    dialog.setAttribute('aria-describedby', 'appDialogMessage');
    dialog.innerHTML = '<div class="app-dialog-head"><span class="app-dialog-icon" aria-hidden="true"></span>'
      + '<h2 id="appDialogTitle" class="app-dialog-title"></h2></div>'
      + '<p id="appDialogMessage" class="app-dialog-message"></p>'
      + '<div class="app-dialog-foot"><button type="button" class="btn btn-outline app-dialog-cancel"></button>'
      + '<button type="button" class="btn btn-primary app-dialog-ok"></button></div>';
    document.body.appendChild(dialog);
    return dialog;
  }

  function show(kind, message, options = {}) {
    const tone = TONES[options.tone] ? options.tone : (kind === 'confirm' ? 'warn' : 'info');
    const preset = TONES[tone];
    return new Promise(resolve => {
      const el = dialog && dialog.isConnected ? dialog : build();
      const ok = el.querySelector('.app-dialog-ok');
      const cancel = el.querySelector('.app-dialog-cancel');
      el.dataset.tone = tone;
      el.querySelector('.app-dialog-icon').textContent = options.icon || preset.icon;
      el.querySelector('.app-dialog-title').textContent = options.title || (kind === 'confirm' ? preset.confirmTitle : preset.alertTitle);
      el.querySelector('.app-dialog-message').textContent = String(message ?? '');
      ok.textContent = options.okText || (kind === 'confirm' ? '确定' : '知道了');
      ok.className = `btn ${tone === 'danger' ? 'btn-danger' : 'btn-primary'} app-dialog-ok`;
      cancel.textContent = options.cancelText || '取消';
      cancel.hidden = kind !== 'confirm';

      const previousFocus = document.activeElement;
      let done = false;
      function finish(result) {
        if (done) return;
        done = true;
        ok.removeEventListener('click', onOk);
        cancel.removeEventListener('click', onCancel);
        el.removeEventListener('cancel', onEsc);
        if (el.open) el.close();
        if (previousFocus && previousFocus.isConnected && typeof previousFocus.focus === 'function') {
          try { previousFocus.focus({ preventScroll: true }); } catch (_) {}
        }
        resolve(kind === 'confirm' ? result : undefined);
      }
      const onOk = () => finish(true);
      const onCancel = () => finish(false);
      // Esc：确认框算取消，提示框算关闭
      const onEsc = event => { event.preventDefault(); finish(false); };
      ok.addEventListener('click', onOk);
      cancel.addEventListener('click', onCancel);
      el.addEventListener('cancel', onEsc);

      // 每次都重新 showModal，保证盖在页面已打开的其它对话框上面
      if (el.open) el.close();
      document.body.appendChild(el);
      el.showModal();
      el.querySelector('.app-dialog-message').scrollTop = 0;
      ok.focus();
    });
  }

  function enqueue(kind, message, options) {
    const next = queue.then(() => show(kind, message, options));
    queue = next.catch(() => {});
    return next;
  }

  window.appAlert = (message, options) => enqueue('alert', message, options);
  window.appConfirm = (message, options) => enqueue('confirm', message, options);
})();
