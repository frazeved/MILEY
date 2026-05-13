/* auth.js — shared across all protected pages
   - 30-minute inactivity auto-logout
   - Injects logout button + username into the nav bar
*/
(function () {
  'use strict';

  const TIMEOUT_MS = 30 * 60 * 1000; // 30 minutes
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

  // ── Inject logout button into nav ───────────────────────────────────────────
  document.addEventListener('DOMContentLoaded', async function () {
    const nav = document.querySelector('nav');
    if (!nav) return;

    // Fetch logged-in user name (best-effort)
    let userName = '';
    try {
      const r = await fetch('/api/me');
      if (r.ok) { const u = await r.json(); userName = u.name || ''; }
    } catch (_) {}

    // Inject styles
    const style = document.createElement('style');
    style.textContent = [
      '.nav-user-area{display:flex;align-items:center;gap:10px;margin-left:auto;flex-shrink:0;}',
      '.nav-username{font-size:12px;font-weight:500;color:var(--text-secondary,#6b7280);white-space:nowrap;display:none;}',
      '.nav-logout{font-size:12px;font-weight:600;padding:5px 12px;border-radius:6px;',
      'background:#fee2e2;color:#dc2626;border:none;cursor:pointer;font-family:inherit;',
      'transition:background 0.15s;white-space:nowrap;line-height:1.4;}',
      '.nav-logout:hover{background:#fecaca;}',
      '.nav-hamburger{margin-left:0!important;}', // reset so user-area provides the push
      '@media(min-width:641px){.nav-username{display:block;}}',
      '@media(max-width:640px){.nav-user-area{gap:6px;}.nav-logout{padding:5px 9px;font-size:11px;}}',
    ].join('');
    document.head.appendChild(style);

    // Build user area
    const area = document.createElement('div');
    area.className = 'nav-user-area';

    if (userName) {
      const span = document.createElement('span');
      span.className = 'nav-username';
      span.textContent = userName;
      area.appendChild(span);
    }

    const btn = document.createElement('button');
    btn.className = 'nav-logout';
    btn.textContent = 'Log Out';
    btn.addEventListener('click', doLogout);
    area.appendChild(btn);

    // Insert before hamburger so it pushes hamburger to the far right on mobile
    const hamburger = nav.querySelector('.nav-hamburger');
    if (hamburger) {
      nav.insertBefore(area, hamburger);
    } else {
      nav.appendChild(area);
    }
  });
})();
