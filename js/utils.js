/**
 * utils.js — Shared helpers, formatters, and UI utilities
 */

const Utils = (() => {

  // ── Number / currency ─────────────────────────────────────────────────────
  function formatPHP(amount) {
    const n = parseFloat(amount) || 0;
    return '₱' + n.toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function parsePHP(str) {
    return parseFloat(String(str).replace(/[₱,\s]/g, '')) || 0;
  }

  // ── Dates ─────────────────────────────────────────────────────────────────
  function formatDate(val) {
    if (!val) return '';
    const d = new Date(val);
    if (isNaN(d)) return val;
    return d.toLocaleDateString('en-PH', { year: 'numeric', month: 'short', day: 'numeric' });
  }

  function toISODate(val) {
    if (!val) return '';
    const d = new Date(val);
    if (isNaN(d)) return '';
    return d.toISOString().split('T')[0];
  }

  function today() {
    return new Date().toISOString().split('T')[0];
  }

  function currentYear() {
    return new Date().getFullYear();
  }

  // ── Strings ───────────────────────────────────────────────────────────────
  function escape(str) {
    return String(str ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function initials(name) {
    return (name || '').split(' ')
      .filter(Boolean)
      .slice(0, 2)
      .map(w => w[0].toUpperCase())
      .join('');
  }

  function slugify(str) {
    return str.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '');
  }

  // ── Status helpers ────────────────────────────────────────────────────────
  const STATUS_CLASSES = {
    'Member':     'badge-member',
    'Exempt':     'badge-exempt',
    'Non-member': 'badge-non-member',
    'Non-Member': 'badge-non-member',
    'TBC':        'badge-tbc',
  };

  function statusBadge(status) {
    const cls = STATUS_CLASSES[status] || 'badge-tbc';
    return `<span class="badge ${cls}">${escape(status || 'TBC')}</span>`;
  }

  function typeBadge(type) {
    const cls = type === 'Family' ? 'badge-family' : 'badge-individual';
    return `<span class="badge ${cls}">${escape(type || '—')}</span>`;
  }

  // ── Toast notifications ───────────────────────────────────────────────────
  let _toastTimer;
  function toast(message, type = 'success') {
    clearTimeout(_toastTimer);
    const el = document.getElementById('toast');
    if (!el) return;
    el.textContent  = message;
    el.className    = `toast toast-${type} show`;
    _toastTimer = setTimeout(() => el.classList.remove('show'), 4000);
  }

  // ── Modal ─────────────────────────────────────────────────────────────────
  function showModal(id) {
    const el = document.getElementById(id);
    if (el) { el.classList.add('open'); el.removeAttribute('hidden'); }
  }

  function hideModal(id) {
    const el = document.getElementById(id);
    if (el) { el.classList.remove('open'); el.setAttribute('hidden', ''); }
  }

  // ── Loading overlay ───────────────────────────────────────────────────────
  function setLoading(active, message = 'Loading…') {
    const el = document.getElementById('loading-overlay');
    if (!el) return;
    if (active) {
      el.querySelector('.loading-msg').textContent = message;
      el.removeAttribute('hidden');
    } else {
      el.setAttribute('hidden', '');
    }
  }

  // ── Confirm dialog (returns promise) ─────────────────────────────────────
  function confirm(message) {
    return new Promise(resolve => {
      const el  = document.getElementById('confirm-modal');
      const msg = document.getElementById('confirm-message');
      const yes = document.getElementById('confirm-yes');
      const no  = document.getElementById('confirm-no');
      if (!el) return resolve(window.confirm(message));
      msg.textContent = message;
      el.classList.add('open');
      const cleanup = (result) => {
        el.classList.remove('open');
        yes.replaceWith(yes.cloneNode(true));
        no.replaceWith(no.cloneNode(true));
        resolve(result);
      };
      document.getElementById('confirm-yes').onclick = () => cleanup(true);
      document.getElementById('confirm-no').onclick  = () => cleanup(false);
    });
  }

  // ── Table helpers ─────────────────────────────────────────────────────────
  function sortTable(rows, key, dir = 'asc') {
    return [...rows].sort((a, b) => {
      const av = (a[key] || '').toString().toLowerCase();
      const bv = (b[key] || '').toString().toLowerCase();
      const n = av.localeCompare(bv);
      return dir === 'asc' ? n : -n;
    });
  }

  function filterRows(rows, query, fields) {
    if (!query.trim()) return rows;
    const q = query.toLowerCase();
    return rows.filter(r => fields.some(f => (r[f] || '').toLowerCase().includes(q)));
  }

  // ── Computed: total paid YTD from Transactions ────────────────────────────
  function totalPaidYTD(memberKey, transactions) {
    const year = currentYear();
    return transactions
      .filter(t => t.MemberKey === memberKey && String(t.Year || '').startsWith(String(year)))
      .reduce((sum, t) => sum + parsePHP(t.AmountPaid), 0);
  }

  // ── Debounce ──────────────────────────────────────────────────────────────
  function debounce(fn, ms = 300) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
  }

  // ── Generate random token for RSVP links ─────────────────────────────────
  function rsvpToken(eventId, memberKey) {
    const raw = btoa(`${eventId}:${memberKey}:${CONFIG.SHEET_ID.slice(0, 8)}`);
    return raw.replace(/=/g, '');
  }

  return {
    formatPHP, parsePHP, formatDate, toISODate, today, currentYear,
    escape, initials, slugify,
    statusBadge, typeBadge,
    toast, showModal, hideModal, setLoading, confirm,
    sortTable, filterRows, totalPaidYTD,
    debounce, rsvpToken,
  };
})();
