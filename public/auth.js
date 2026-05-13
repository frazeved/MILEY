/* auth.js — shared across all protected pages
   - 30-minute inactivity auto-logout
   - Username on the LEFT of the nav bar
   - Log Out button on the RIGHT of the nav bar
*/
(function () {
  'use strict';

  const TIMEOUT_MS = 30 * 60 * 1000;
  let timer;

  function doLogout() {
    window.location.href = '/api/logout';
  }

  function resetTimer() {
    clearTimeout(timer);
    timer = setTimeout(doLogout, TIMEOUT_MS);
  }

  ['mousemove', 'keydown', 'mousedown', 'scroll', 'touchstart'].forEach(function (ev) {
    document.addEventListener(ev, resetTimer, { passive: true });
  });
  resetTimer();

  document.addEventListener('DOMContentLoaded', async function () {
    const nav = document.querySelector('nav');
    if (!nav) return;

    // Get logged-in user
    let userName = '';
    try {
      const r = await fetch('/api/me');
      if (r.ok) { const u = await r.json(); userName = u.name || ''; }
    } catch (_) {}

    // Styles
    const s = document.createElement('style');
    s.textContent =
      '.auth-username{font-size:12px;font-weight:600;color:var(--text-secondary,#6b7280);' +
      'white-space:nowrap;flex-shrink:0;letter-spacing:0.2px;}' +

      '.auth-logout{font-size:12px;font-weight:700;padding:5px 14px;border-radius:6px;' +
      'background:#fee2e2;color:#dc2626;border:none;cursor:pointer;font-family:inherit;' +
      'transition:background 0.15s;white-space:nowrap;flex-shrink:0;margin-left:auto;}' +
      '.auth-logout:hover{background:#fecaca;}' +

      /* On mobile: hide username, remove auto-margin from logout (hamburger handles the push) */
      '@media(max-width:640px){' +
        '.auth-username{display:none;}' +
        '.auth-logout{margin-left:0;font-size:11px;padding:5px 10px;}' +
      '}';
    document.head.appendChild(s);

    // USERNAME — prepend as first child so it sits on the far left
    if (userName) {
      const nameEl = document.createElement('span');
      nameEl.className = 'auth-username';
      nameEl.textContent = userName;
      nav.prepend(nameEl);
    }

    // LOG OUT — append as last child so it sits on the far right (margin-left:auto)
    const btn = document.createElement('button');
    btn.className = 'auth-logout';
    btn.textContent = 'Log Out';
    btn.addEventListener('click', doLogout);
    nav.appendChild(btn);
  });
})();
