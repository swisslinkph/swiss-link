/**
 * settings.js — Settings view: manage admin access
 */

const Settings = (() => {
  let _admins = []; // rows from Admins sheet

  const DEFAULT_RATES = [
    { id: 'single-jr',   label: 'Single Jr. (18–25)',             amount: 1500 },
    { id: 'mm-single',   label: 'Metro Manila – Individual',       amount: 2800 },
    { id: 'mm-family',   label: 'Metro Manila – Family',           amount: 3500 },
    { id: 'prov-single', label: 'Province/Overseas – Individual',  amount: 2000 },
    { id: 'prov-family', label: 'Province/Overseas – Family',      amount: 2300 },
  ];

  // ── Render ────────────────────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading settings…');
    try {
      [_admins] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.ADMINS).catch(e => {
          if (!e.message?.toLowerCase().includes('unable to parse range')) Utils.toast(e.message, 'error');
          return [];
        }),
        _renderRates(),
      ]);
    } finally {
      _renderList();
      Utils.setLoading(false);
    }
  }

  async function _renderRates() {
    const container = document.getElementById('settings-rates-body');
    if (!container) return;

    let rates = [];
    let fromSheet = false;
    try {
      const rows = await Sheets.getAll(CONFIG.SHEETS.RATES);
      if (rows.length) {
        rates = rows
          .map(r => ({ id: r['Tier ID']?.trim(), label: r['Label']?.trim(), amount: parseFloat(r['Amount']) }))
          .filter(r => r.id && r.label && !isNaN(r.amount));
        fromSheet = rates.length > 0;
      }
    } catch {}

    if (!fromSheet) rates = DEFAULT_RATES;

    const badge = fromSheet
      ? `<span class="rates-source-badge rates-live">✓ Live from Rates sheet</span>`
      : `<span class="rates-source-badge rates-fallback">Using defaults — Rates sheet not found</span>`;

    container.innerHTML = `
      ${badge}
      <table class="rates-table">
        <thead><tr><th>Membership Tier</th><th class="rates-amount-col">Annual Rate</th></tr></thead>
        <tbody>
          ${rates.map(r => `
            <tr>
              <td>${Utils.escape(r.label)}</td>
              <td class="rates-amount-col">${Utils.formatPHP(r.amount)}</td>
            </tr>`).join('')}
        </tbody>
      </table>`;
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

  // ── Membership year rollover ──────────────────────────────────────────────
  async function startNewYear() {
    const yearEl = document.getElementById('settings-rollover-year');
    const year   = parseInt(yearEl?.value, 10);
    if (!year || year < 2020 || year > 2100) {
      Utils.toast('Please enter a valid year (e.g. 2027).', 'error'); return;
    }

    Utils.setLoading(true, 'Checking members…');
    let members;
    try {
      members = await Sheets.getAll(CONFIG.SHEETS.MEMBERS);
    } catch (e) {
      Utils.toast('Failed to load members: ' + e.message, 'error');
      Utils.setLoading(false);
      return;
    }
    Utils.setLoading(false);

    const toReset = members.filter(m => {
      const status      = m['Membership Status'];
      const renewalYear = parseInt(m['Renewal Year'], 10) || 0;
      return status !== 'Exempt' && renewalYear < year;
    });

    if (!toReset.length) {
      Utils.toast(`All eligible members already have ${year} renewal on record.`);
      return;
    }

    const ok = await Utils.confirm(
      `Start ${year} membership year?\n\n` +
      `${toReset.length} member(s) will be set to TBC.\n` +
      `Exempt members will not be affected.\n\nThis cannot be undone.`
    );
    if (!ok) return;

    const btn = document.getElementById('settings-rollover-btn');
    if (btn) btn.disabled = true;
    Utils.setLoading(true, `Resetting ${toReset.length} members to TBC…`);
    try {
      for (const m of toReset) {
        await Sheets.update(CONFIG.SHEETS.MEMBERS, m._rowIndex, {
          ...m, 'Membership Status': 'TBC',
        });
      }
      Utils.toast(`${toReset.length} member(s) set to TBC for ${year}.`);
      if (yearEl) yearEl.value = '';
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      if (btn) btn.disabled = false;
      Utils.setLoading(false);
    }
  }

  function init() {
    document.getElementById('settings-add-btn')
      ?.addEventListener('click', addAdmin);
    document.getElementById('settings-rollover-btn')
      ?.addEventListener('click', startNewYear);
    const link = document.getElementById('settings-sheet-link');
    if (link) link.href = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}`;
  }

  return { render, init, confirmRemove };
})();
