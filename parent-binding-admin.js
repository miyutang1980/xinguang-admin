/* Same-origin isolated review frame: no credentials in URLs/postMessage/storage.
 * The frame uses its own POST fetch, isolated from legacy admin GET middleware.
 * Gateway independently validates staff/admin on every read and decision. */
(function () {
  'use strict';
  let frame = null;
  const allowed = () => typeof _loggedIn !== 'undefined' && _loggedIn &&
    typeof _currentUser !== 'undefined' && _currentUser &&
    ['admin','staff'].includes(String(_currentUser.role).toLowerCase());
  function reset() {
    if (frame) { frame.remove(); frame = null; }
  }
  window.XGParentBinding = {
    allowed, reset,
    session(child) {
      if (!allowed() || currentSection !== 'parentBinding' ||
          !frame || !frame.isConnected || frame.contentWindow !== child) return null;
      let password = '';
      try { password = sessionStorage.getItem('xg_session_pwd') || localStorage.getItem('xg_admin_pwd') || ''; } catch (_) {}
      if (!password) return null;
      return {adminUser:_currentUser.username, adminPass:password, epoch:_authEpoch};
    },
    resize(child, height) {
      if (frame && frame.contentWindow === child && Number.isFinite(height)) {
        frame.style.height = Math.max(440, Math.min(height + 16, 200000)) + 'px';
      }
    }
  };
  window.renderParentBinding = function (container) {
    reset();
    if (!allowed()) { container.textContent = '請使用管理員或行政帳號登入。'; return; }
    container.replaceChildren();
    const info = document.createElement('p');
    info.className = 'binding-session-info';
    info.textContent = '沿用目前後台登入。核對家長身分與每位孩子後，才會送出整份核准；開啟此頁不會自動授權。';
    frame = document.createElement('iframe');
    frame.id = 'parentBindingFrame';
    frame.title = '家長綁定申請審核';
    frame.src = './parent-binding/?embed=1&v=20260923-5';
    frame.referrerPolicy = 'no-referrer';
    container.append(info, frame);
    if (!document.getElementById('bindingMobileMenu')) {
      const button = document.createElement('button');
      button.id = 'bindingMobileMenu'; button.className = 'btn btn-outline';
      button.textContent = '展開後台選單'; button.setAttribute('aria-expanded','false');
      button.onclick = () => {
        const open = document.body.classList.toggle('binding-menu-expanded');
        button.setAttribute('aria-expanded',String(open));
        button.textContent = open ? '收起後台選單' : '展開後台選單';
      };
      document.querySelector('.sidebar').prepend(button);
    }
  };
  window.addEventListener('pagehide', reset);
})();
