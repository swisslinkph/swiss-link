/**
 * router.js — Simple hash-based SPA router
 * Routes: #dashboard | #members | #events | #frontdesk | #email
 */

const Router = (() => {
  const _routes = {};
  let _current  = null;

  function register(hash, handler) {
    _routes[hash] = handler;
  }

  function navigate(hash) {
    window.location.hash = hash;
  }

  function _resolve() {
    const hash = window.location.hash.replace('#', '') || 'dashboard';
    const handler = _routes[hash];

    // Update sidebar active state
    document.querySelectorAll('.nav-item').forEach(el => {
      el.classList.toggle('active', el.dataset.route === hash);
    });

    // Hide all views
    document.querySelectorAll('.view').forEach(el => el.setAttribute('hidden', ''));

    // Show target view
    const view = document.getElementById(`view-${hash}`);
    if (view) view.removeAttribute('hidden');

    // Update page title
    const label = document.querySelector(`[data-route="${hash}"] .nav-label`);
    document.title = `${label?.textContent || hash} — Swiss Club Admin`;

    if (handler && _current !== hash) {
      _current = hash;
      handler();
    }
  }

  function init() {
    window.addEventListener('hashchange', _resolve);
    // Sidebar click listeners
    document.querySelectorAll('.nav-item[data-route]').forEach(el => {
      el.addEventListener('click', () => navigate(el.dataset.route));
    });
    _resolve();
  }

  return { register, navigate, init };
})();
