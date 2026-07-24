/**
 * registrations.js — Event Registrations view
 * Lists registrations for a given event, allows Wilma to confirm payments
 * and add walk-in registrations.
 */

const Registrations = (() => {
  let _eventId   = null;
  let _event     = null;
  let _all       = [];   // all registrations for this event
  let _filtered  = [];
  let _events    = [];   // all events (for name lookup)

  const C = {
    ID:        'RegistrationID',
    TS:        'Timestamp',
    SOURCE:    'Source',
    EVID:      'EventID',
    EVNAME:    'EventName',
    LAST:      'LastName',
    FIRST:     'FirstName',
    EMAIL:     'Email',
    MKEY:      'MemberKey',
    MEM_QTY:   'MemberQty',
    GUEST_QTY: 'GuestQty',
    KIDS_QTY:  'KidsQty',
    WALKIN:    'IsWalkIn',
    TOTAL:     'TotalDue',
    PAY_NOTE:  'PaymentNote',
    STATUS:    'PaymentStatus',
    PAY_MODE:  'PaymentMode',
    AMOUNT:    'AmountPaid',
    NOTES:     'AdminNotes',
  };

  // ── Render (called by Router) ─────────────────────────────────────────────
  async function render() {
    _eventId = sessionStorage.getItem('reg_event');

    Utils.setLoading(true, 'Loading registrations…');
    try {
      [_all, _events] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).catch(() => []),
        Sheets.getAll(CONFIG.SHEETS.EVENTS).catch(() => []),
      ]);

      _event = _events.find(e => e.EventID === _eventId) || null;

      // Filter to this event
      _all = _all.filter(r => r[C.EVID] === _eventId);

      _applyFilter();
      _renderHeader();
      _renderSummary();
      _renderTable();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  function _renderHeader() {
    const titleEl = document.getElementById('reg-event-title');
    if (titleEl) titleEl.textContent = _event ? _event.Title : (_eventId || 'Registrations');

    // Show/hide Kids column based on whether KidsFee is set
    const hasKids = _event && parseFloat(_event.KidsFee) > 0;
    document.querySelectorAll('.reg-kids-col').forEach(el => {
      el.style.display = hasKids ? '' : 'none';
    });

    // Show walk-in rates hint in the walk-in modal
    const ratesEl = document.getElementById('reg-walkin-rates');
    if (ratesEl && _event) {
      const parts = [];
      if (_event.WalkInMemberFee) parts.push(`Member ₱${Utils.formatPHP(_event.WalkInMemberFee)}`);
      if (_event.WalkInGuestFee)  parts.push(`Guest ₱${Utils.formatPHP(_event.WalkInGuestFee)}`);
      if (_event.KidsFee)         parts.push(`Kids ₱${Utils.formatPHP(_event.KidsFee)} (pre-reg rate)`);
      ratesEl.textContent = parts.length ? 'Walk-in rates: ' + parts.join(' · ') : '';
    }
    const kidsRowEl = document.getElementById('reg-wi-kids-row');
    if (kidsRowEl) kidsRowEl.style.display = hasKids ? '' : 'none';
  }

  function _renderSummary() {
    const el = document.getElementById('reg-summary');
    if (!el) return;

    const total      = _all.length;
    const confirmed  = _all.filter(r => r[C.STATUS] === 'Confirmed').length;
    const pending    = _all.filter(r => r[C.STATUS] === 'Pending').length;
    const cancelled  = _all.filter(r => r[C.STATUS] === 'Cancelled').length;
    const walkIns    = _all.filter(r => r[C.WALKIN] === 'Yes').length;
    const revenue    = _all
      .filter(r => r[C.STATUS] === 'Confirmed')
      .reduce((s, r) => s + Utils.parsePHP(r[C.AMOUNT] || r[C.TOTAL]), 0);

    const memberPax  = _all.reduce((s, r) => s + (parseInt(r[C.MEM_QTY], 10) || 0), 0);
    const guestPax   = _all.reduce((s, r) => s + (parseInt(r[C.GUEST_QTY], 10) || 0), 0);
    const kidsPax    = _all.reduce((s, r) => s + (parseInt(r[C.KIDS_QTY], 10) || 0), 0);
    const totalPax   = memberPax + guestPax + kidsPax;

    el.innerHTML = `
      <div class="reg-stat"><span class="reg-stat-num">${total}</span><span class="reg-stat-label">Registrations</span></div>
      <div class="reg-stat"><span class="reg-stat-num">${totalPax}</span><span class="reg-stat-label">Total Pax</span></div>
      <div class="reg-stat"><span class="reg-stat-num reg-confirmed">${confirmed}</span><span class="reg-stat-label">Confirmed</span></div>
      <div class="reg-stat"><span class="reg-stat-num reg-pending">${pending}</span><span class="reg-stat-label">Pending</span></div>
      ${cancelled ? `<div class="reg-stat"><span class="reg-stat-num reg-cancelled">${cancelled}</span><span class="reg-stat-label">Cancelled</span></div>` : ''}
      ${walkIns   ? `<div class="reg-stat"><span class="reg-stat-num">${walkIns}</span><span class="reg-stat-label">Walk-ins</span></div>` : ''}
      <div class="reg-stat"><span class="reg-stat-num">${Utils.formatPHP(revenue)}</span><span class="reg-stat-label">Collected</span></div>`;
  }

  function _renderTable() {
    const tbody = document.getElementById('reg-tbody');
    if (!tbody) return;

    if (!_filtered.length) {
      tbody.innerHTML = `<tr><td colspan="8" class="empty-state" style="text-align:center;padding:24px;">No registrations found.</td></tr>`;
      return;
    }

    tbody.innerHTML = _filtered.map(r => {
      const name   = [r[C.LAST], r[C.FIRST]].filter(Boolean).join(', ');
      const pax    = _paxSummary(r);
      const status = r[C.STATUS] || 'Pending';
      const isWI   = r[C.WALKIN] === 'Yes';

      return `<tr>
        <td>
          <strong>${Utils.escape(name || '—')}</strong>
          ${r[C.EMAIL] ? `<br><span class="text-muted" style="font-size:12px;">${Utils.escape(r[C.EMAIL])}</span>` : ''}
          ${isWI ? ' <span class="badge badge-walkin">Walk-in</span>' : ''}
        </td>
        <td>${Utils.escape(r[C.MKEY] || '—')}</td>
        <td>${pax}</td>
        <td class="amount">${Utils.formatPHP(r[C.TOTAL])}</td>
        <td><span class="badge badge-reg-${status.toLowerCase()}">${status}</span></td>
        <td>${Utils.escape(r[C.PAY_NOTE] || r[C.NOTES] || '—')}</td>
        <td>
          ${status !== 'Confirmed' ? `
            <button class="btn btn-sm btn-primary" onclick="Registrations.openConfirmPayment('${Utils.escape(r[C.ID])}')">
              Confirm
            </button>` : `
            <span class="text-muted" style="font-size:12px;">
              ${Utils.formatPHP(r[C.AMOUNT])} via ${Utils.escape(r[C.PAY_MODE] || '—')}
            </span>`}
          ${status !== 'Cancelled' ? `
            <button class="btn btn-sm btn-danger-outline" style="margin-left:4px;"
                    onclick="Registrations.cancelRegistration('${Utils.escape(r[C.ID])}')">
              ✕
            </button>` : ''}
        </td>
      </tr>`;
    }).join('');
  }

  function _paxSummary(r) {
    const parts = [];
    const m = parseInt(r[C.MEM_QTY], 10) || 0;
    const g = parseInt(r[C.GUEST_QTY], 10) || 0;
    const k = parseInt(r[C.KIDS_QTY], 10) || 0;
    if (m) parts.push(`${m}M`);
    if (g) parts.push(`${g}G`);
    if (k) parts.push(`${k}K`);
    return parts.join(' + ') || '—';
  }

  function _applyFilter() {
    const q      = (document.getElementById('reg-search')?.value || '').toLowerCase();
    const status = document.getElementById('reg-filter-status')?.value || '';
    const source = document.getElementById('reg-filter-source')?.value || '';

    _filtered = _all.filter(r => {
      if (status && r[C.STATUS] !== status) return false;
      if (source && r[C.SOURCE] !== source) return false;
      if (q) {
        const hay = [r[C.LAST], r[C.FIRST], r[C.EMAIL], r[C.MKEY], r[C.PAY_NOTE], r[C.NOTES]]
          .join(' ').toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }

  function applyFilter() {
    _applyFilter();
    _renderTable();
  }

  // ── Confirm Payment modal ─────────────────────────────────────────────────
  function openConfirmPayment(regId) {
    const r = _all.find(x => x[C.ID] === regId);
    if (!r) return;

    const name = [r[C.LAST], r[C.FIRST]].filter(Boolean).join(', ') || regId;
    document.getElementById('reg-confirm-id').value       = regId;
    document.getElementById('reg-confirm-name').textContent = name;
    document.getElementById('reg-confirm-tickets').textContent = _paxSummary(r);
    document.getElementById('reg-confirm-due').textContent = 'Total Due: ' + Utils.formatPHP(r[C.TOTAL]);
    document.getElementById('reg-confirm-amount').value   = r[C.TOTAL] || '';
    document.getElementById('reg-confirm-notes').value    = r[C.NOTES] || '';
    Utils.showModal('reg-confirm-modal');
  }

  async function saveConfirmPayment() {
    const btn   = document.getElementById('reg-confirm-save-btn');
    btn.disabled = true;
    try {
      const regId  = document.getElementById('reg-confirm-id').value;
      const amount = document.getElementById('reg-confirm-amount').value.trim();
      const mode   = document.getElementById('reg-confirm-mode').value;
      const notes  = document.getElementById('reg-confirm-notes').value.trim();

      if (!amount) { Utils.toast('Please enter the amount paid.', 'error'); return; }

      const r = _all.find(x => x[C.ID] === regId);
      if (!r) return;

      await Sheets.update(CONFIG.SHEETS.REGISTRATIONS, r._rowIndex, {
        ...r,
        [C.STATUS]:   'Confirmed',
        [C.PAY_MODE]: mode,
        [C.AMOUNT]:   amount,
        [C.NOTES]:    notes,
      });

      Utils.hideModal('reg-confirm-modal');
      Utils.toast('Payment confirmed.');
      await render();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  // ── Cancel registration ───────────────────────────────────────────────────
  async function cancelRegistration(regId) {
    const r = _all.find(x => x[C.ID] === regId);
    if (!r) return;
    const name = [r[C.LAST], r[C.FIRST]].filter(Boolean).join(' ') || regId;
    const ok = await Utils.confirm(`Cancel registration for ${name}?`);
    if (!ok) return;
    try {
      await Sheets.update(CONFIG.SHEETS.REGISTRATIONS, r._rowIndex, {
        ...r, [C.STATUS]: 'Cancelled',
      });
      Utils.toast('Registration cancelled.');
      await render();
    } catch (e) {
      Utils.toast(e.message, 'error');
    }
  }

  // ── Walk-in modal ─────────────────────────────────────────────────────────
  function openAddWalkIn() {
    document.getElementById('reg-walkin-modal')
      ?.querySelectorAll('input')
      .forEach(el => { el.value = el.type === 'number' ? '0' : ''; });
    Utils.showModal('reg-walkin-modal');
  }

  async function saveWalkIn() {
    const btn = document.getElementById('reg-walkin-save-btn');
    btn.disabled = true;
    try {
      const get = id => document.getElementById(id)?.value?.trim() || '';
      const last = get('reg-wi-last');
      if (!last) { Utils.toast('Last name is required.', 'error'); return; }

      const mQty = parseInt(get('reg-wi-member-qty'), 10) || 0;
      const gQty = parseInt(get('reg-wi-guest-qty'), 10)  || 0;
      const kQty = parseInt(get('reg-wi-kids-qty'), 10)   || 0;

      if (mQty + gQty + kQty === 0) {
        Utils.toast('Please enter at least 1 ticket.', 'error'); return;
      }

      const wiMemberFee = parseFloat(_event?.WalkInMemberFee || _event?.MemberFee || 0);
      const wiGuestFee  = parseFloat(_event?.WalkInGuestFee  || _event?.GuestFee  || 0);
      const kidsFee     = parseFloat(_event?.KidsFee || 0);
      const total       = (mQty * wiMemberFee) + (gQty * wiGuestFee) + (kQty * kidsFee);

      const amount  = get('reg-wi-amount') || String(total);
      const rows    = await Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).catch(() => []);
      const maxNum  = rows
        .map(r => parseInt((r[C.ID] || '').replace(/\D/g, ''), 10))
        .filter(n => !isNaN(n));
      const nextNum = maxNum.length ? Math.max(...maxNum) + 1 : 1;
      const regId   = 'REG-' + String(nextNum).padStart(4, '0');

      await Sheets.append(CONFIG.SHEETS.REGISTRATIONS, {
        [C.ID]:        regId,
        [C.TS]:        new Date().toISOString(),
        [C.SOURCE]:    'Walk-in',
        [C.EVID]:      _eventId,
        [C.EVNAME]:    _event?.Title || '',
        [C.LAST]:      last,
        [C.FIRST]:     get('reg-wi-first'),
        [C.EMAIL]:     get('reg-wi-email'),
        [C.MEM_QTY]:  mQty,
        [C.GUEST_QTY]: gQty,
        [C.KIDS_QTY]:  kQty,
        [C.WALKIN]:    'Yes',
        [C.TOTAL]:     total,
        [C.STATUS]:    amount ? 'Confirmed' : 'Pending',
        [C.PAY_MODE]:  get('reg-wi-mode'),
        [C.AMOUNT]:    amount,
        [C.NOTES]:     get('reg-wi-notes'),
      });

      Utils.hideModal('reg-walkin-modal');
      Utils.toast('Walk-in added.');
      await render();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  // ── Add Registration modal ────────────────────────────────────────────────
  function openAddRegistration() {
    const form = document.getElementById('reg-add-modal');
    if (!form) return;
    form.querySelectorAll('input[type=text], input[type=email], input[type=number]')
      .forEach(el => { el.value = el.type === 'number' ? '0' : ''; });
    const statusEl = document.getElementById('reg-add-status');
    if (statusEl) statusEl.value = 'Pending';
    const sourceEl = document.getElementById('reg-add-source');
    if (sourceEl) sourceEl.value = 'Manual';
    const walkinEl = document.getElementById('reg-add-walkin');
    if (walkinEl) walkinEl.checked = false;
    _toggleAddPaymentFields(false);
    const hasKids = _event && parseFloat(_event.KidsFee) > 0;
    const kidsRow = document.getElementById('reg-add-kids-row');
    if (kidsRow) kidsRow.style.display = hasKids ? '' : 'none';
    Utils.showModal('reg-add-modal');
  }

  function onAddSourceChange() {
    const source  = document.getElementById('reg-add-source')?.value;
    const walkinEl = document.getElementById('reg-add-walkin');
    if (walkinEl) walkinEl.checked = source === 'Walk-in';
    recalcAddTotal();
  }

  function onAddWalkInChange() {
    recalcAddTotal();
  }

  function recalcAddTotal() {
    if (!_event) return;
    const isWalkIn = document.getElementById('reg-add-walkin')?.checked;
    const mFee = parseFloat(isWalkIn ? (_event.WalkInMemberFee || _event.MemberFee) : _event.MemberFee) || 0;
    const gFee = parseFloat(isWalkIn ? (_event.WalkInGuestFee  || _event.GuestFee)  : _event.GuestFee)  || 0;
    const kFee = parseFloat(_event.KidsFee) || 0;
    const mQty = parseInt(document.getElementById('reg-add-member-qty')?.value, 10) || 0;
    const gQty = parseInt(document.getElementById('reg-add-guest-qty')?.value,  10) || 0;
    const kQty = parseInt(document.getElementById('reg-add-kids-qty')?.value,   10) || 0;
    const total = (mQty * mFee) + (gQty * gFee) + (kQty * kFee);
    const totalEl = document.getElementById('reg-add-total');
    if (totalEl) totalEl.value = total || '';
  }

  function onAddStatusChange() {
    const status = document.getElementById('reg-add-status')?.value;
    _toggleAddPaymentFields(status === 'Confirmed');
  }

  function _toggleAddPaymentFields(show) {
    document.getElementById('reg-add-mode-row')  ?.style && (document.getElementById('reg-add-mode-row').style.display   = show ? '' : 'none');
    document.getElementById('reg-add-amount-row')?.style && (document.getElementById('reg-add-amount-row').style.display = show ? '' : 'none');
  }

  async function saveAddRegistration() {
    const btn = document.getElementById('reg-add-save-btn');
    btn.disabled = true;
    try {
      const get = id => document.getElementById(id)?.value?.trim() || '';
      const last = get('reg-add-last');
      if (!last) { Utils.toast('Last name is required.', 'error'); return; }

      const mQty    = parseInt(get('reg-add-member-qty'), 10) || 0;
      const gQty    = parseInt(get('reg-add-guest-qty'),  10) || 0;
      const kQty    = parseInt(get('reg-add-kids-qty'),   10) || 0;
      if (mQty + gQty + kQty === 0) { Utils.toast('Enter at least 1 ticket.', 'error'); return; }

      const isWalkIn = document.getElementById('reg-add-walkin')?.checked;
      const source   = get('reg-add-source') || 'Manual';
      const status   = get('reg-add-status') || 'Pending';
      const total    = get('reg-add-total');
      const amount   = status === 'Confirmed' ? (get('reg-add-amount') || total) : '';

      const rows   = await Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).catch(() => []);
      const maxNum = rows.map(r => parseInt((r[C.ID] || '').replace(/\D/g, ''), 10)).filter(n => !isNaN(n));
      const nextNum = maxNum.length ? Math.max(...maxNum) + 1 : 1;
      const regId   = 'REG-' + String(nextNum).padStart(4, '0');

      await Sheets.append(CONFIG.SHEETS.REGISTRATIONS, {
        [C.ID]:        regId,
        [C.TS]:        new Date().toISOString(),
        [C.SOURCE]:    source,
        [C.EVID]:      _eventId,
        [C.EVNAME]:    _event?.Title || '',
        [C.LAST]:      last,
        [C.FIRST]:     get('reg-add-first'),
        [C.EMAIL]:     get('reg-add-email'),
        [C.MKEY]:      get('reg-add-member-key'),
        [C.MEM_QTY]:   mQty,
        [C.GUEST_QTY]: gQty,
        [C.KIDS_QTY]:  kQty,
        [C.WALKIN]:    isWalkIn ? 'Yes' : 'No',
        [C.TOTAL]:     total,
        [C.PAY_NOTE]:  get('reg-add-pay-note'),
        [C.STATUS]:    status,
        [C.PAY_MODE]:  status === 'Confirmed' ? get('reg-add-mode') : '',
        [C.AMOUNT]:    amount,
        [C.NOTES]:     get('reg-add-notes'),
      });

      Utils.hideModal('reg-add-modal');
      Utils.toast('Registration saved.');
      await render();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  // ── Back to events ────────────────────────────────────────────────────────
  function backToEvents() {
    Router.navigate('events');
  }

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    document.getElementById('reg-back-btn')
      ?.addEventListener('click', backToEvents);
    document.getElementById('reg-walkin-btn')
      ?.addEventListener('click', openAddWalkIn);
    document.getElementById('reg-search')
      ?.addEventListener('input', applyFilter);
    document.getElementById('reg-filter-status')
      ?.addEventListener('change', applyFilter);
    document.getElementById('reg-filter-source')
      ?.addEventListener('change', applyFilter);
  }

  return {
    render, init,
    applyFilter,
    openConfirmPayment, saveConfirmPayment,
    cancelRegistration,
    openAddRegistration, onAddSourceChange, onAddWalkInChange,
    recalcAddTotal, onAddStatusChange, saveAddRegistration,
    openAddWalkIn, saveWalkIn,
    backToEvents,
  };
})();
