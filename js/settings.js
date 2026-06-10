/**
 * settings.js — Settings view: manage admin access
 */

const Settings = (() => {
  let _admins = []; // rows from Admins sheet

  // ── Render ────────────────────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading settings…');
    try {
      _admins = await Sheets.getAll(CONFIG.SHEETS.ADMINS);
    } catch (e) {
      if (e.message && e.message.toLowerCase().includes('unable to parse range')) {
        // Admins sheet doesn't exist yet — ensureSheets() will create it
        _admins = [];
      } else {
        Utils.toast(e.message, 'error');
      }
    } finally {
      _renderList();
      Utils.setLoading(false);
    }
  }

  function _renderList() {
    const tbody = document.getElementById('settings-admin-tbody');
    if (!tbody) return;

    const builtInRows = CONFIG.ALLOWED_EMAILS.map(email => `
      <tr class="settings-builtin-row">
        <td>${Utils.escape(email)}</td>
        <td class="text-muted">—</td>
        <td class="text-muted">—</td>
        <td><span class="badge badge-exempt">Built-in</span></td>
        <td></td>
      </tr>`).join('');

    const dynamicRows = _admins.map(a => `
      <tr>
        <td>${Utils.escape(a['Email'] || '')}</td>
        <td>${Utils.escape(a['Name'] || '—')}</td>
        <td>${Utils.formatDate(a['Added Date']) || '—'}</td>
        <td><span class="badge badge-member">Admin</span></td>
        <td>
          <button class="btn btn-sm btn-danger"
                  onclick="Settings.confirmRemove('${Utils.escape(a['Email'])}')">
            Remove
          </button>
        </td>
      </tr>`).join('');

    tbody.innerHTML = builtInRows + (dynamicRows || `
      <tr class="dynamic-empty-row">
        <td colspan="5" class="text-muted" style="text-align:center;padding:12px;">
          No additional admins yet.
        </td>
      </tr>`);

    // Keep auth module in sync
    Auth.setAdminEmails(_admins.map(a => a['Email']));
  }

  // ── Add admin ─────────────────────────────────────────────────────────────
  async function addAdmin() {
    const emailEl = document.getElementById('settings-new-email');
    const nameEl  = document.getElementById('settings-new-name');
    const email   = (emailEl?.value || '').trim().toLowerCase();
    const name    = (nameEl?.value  || '').trim();

    if (!email || !email.includes('@')) {
      Utils.toast('Please enter a valid email address.', 'error'); return;
    }

    const allEmails = [
      ...CONFIG.ALLOWED_EMAILS,
      ..._admins.map(a => (a['Email'] || '').toLowerCase()),
    ];
    if (allEmails.includes(email)) {
      Utils.toast('This email already has admin access.', 'error'); return;
    }

    const btn = document.getElementById('settings-add-btn');
    if (btn) btn.disabled = true;

    try {
      await Sheets.append(CONFIG.SHEETS.ADMINS, {
        Email:        email,
        Name:         name,
        'Added Date': Utils.today(),
        'Added By':   Auth.getUserEmail(),
      });
      if (emailEl) emailEl.value = '';
      if (nameEl)  nameEl.value  = '';
      Utils.toast(`${email} added as admin`);
      await render();
    } catch (e) {
      Utils.toast('Failed to add admin: ' + e.message, 'error');
    } finally {
      if (btn) btn.disabled = false;
    }
  }

  // ── Remove admin ──────────────────────────────────────────────────────────
  async function confirmRemove(email) {
    const ok = await Utils.confirm(
      `Remove ${email} as admin?\nThey will lose access on their next login.`
    );
    if (!ok) return;

    const admin = _admins.find(a => (a['Email'] || '').toLowerCase() === email.toLowerCase());
    if (!admin) return;

    try {
      await Sheets.deleteRow(CONFIG.SHEETS.ADMINS, admin._rowIndex);
      Utils.toast(`${email} removed`);
      await render();
    } catch (e) {
      Utils.toast('Failed to remove admin: ' + e.message, 'error');
    }
  }

  function init() {
    document.getElementById('settings-add-btn')
      ?.addEventListener('click', addAdmin);
  }

  return { render, init, confirmRemove };
})();
