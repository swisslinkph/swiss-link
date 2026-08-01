/**
 * registrations.js — Event Registrations view
 * Lists registrations for a given event, allows Wilma to confirm payments
 * and add walk-in registrations.
 */

const Registrations = (() => {
  let _eventId        = null;
  let _event          = null;
  let _all            = [];   // all registrations for this event
  let _filtered       = [];
  let _events         = [];   // all events (for name lookup)
  let _members        = [];   // for member key lookup
  let _editingRegId     = null;
  let _slotAddContext    = null;  // { type: 'confirm'|'edit', slotIdx: number }
  let _slotAddPrimaryKey = null; // primary member key for optional family link

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
    SLOTS:     'MemberSlots',
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
      _all   = _all.filter(r => r[C.EVID] === _eventId);

      _applyFilter();
      _renderHeader();
      _renderSummary();
      _renderTable();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }

    // Background form sync — page is already visible; update if new rows arrive
    if (_event?.FormSheetID) {
      _syncFromForm().then(n => {
        if (!n) return;
        Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).then(rows => {
          _all = rows.filter(r => r[C.EVID] === _eventId);
          _applyFilter();
          _renderSummary();
          _renderTable();
          Utils.toast(`${n} new registration${n !== 1 ? 's' : ''} synced from Google Form.`);
        });
      }).catch(() => {});
    }
  }

  // ── Google Form sync helpers ───────────────────────────────────────────────
  function _detectFormCols(headers) {
    const h = headers.map(s => s.toLowerCase());
    const idx = fn => h.findIndex(fn);
    return {
      ts:       idx(s => s.includes('timestamp')),
      email:    idx(s => s.includes('email') && !s.includes('payment')),
      last:     idx(s => s.includes('family') || s.includes('last name') || s.includes('surname')),
      first:    idx(s => s.includes('first name') || s.includes('given name')),
      status:   idx(s => s.includes('member status')),
      memQty:   idx(s => s.includes('adult') && s.includes('member') && !s.includes('non') && !s.includes('guest')),
      guestQty: idx(s => s.includes('adult') && (s.includes('non') || s.includes('guest'))),
      kidsQty:  idx(s => s.includes('kid') || s.includes('child')),
      comments: idx(s => s.includes('comment') || s.includes('question')),
    };
  }

  function _feeFromHeader(header) {
    const m = (header || '').match(/[\d,]+\s*(?:php|₱)/i);
    return m ? parseInt(m[0].replace(/[^0-9]/g, ''), 10) : 0;
  }

  async function _syncFromForm() {
    if (!_event?.FormSheetID || !_eventId) return 0;
    const tabName = _event.FormSheetTab || 'Form Responses 1';

    const { headers, rows } = await Sheets.getFromSheet(_event.FormSheetID, tabName, 1);
    if (!headers.length || !rows.length) return 0;

    const cols = _detectFormCols(headers);
    const get  = (row, i) => i >= 0 ? (row[i] || '').toString().trim() : '';

    // Load ALL registrations for dedup (not just this event)
    const allRegs = await Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).catch(() => []);
    const seenTs  = new Set(allRegs.map(r => r[C.TS]));
    const maxNum  = allRegs.reduce((m, r) => {
      const n = parseInt((r[C.ID] || '').replace(/\D/g, ''), 10);
      return isNaN(n) ? m : Math.max(m, n);
    }, 0);

    // Fee per ticket type — event config takes priority, then fall back to column header price
    const mFee = parseFloat(_event.MemberFee) || _feeFromHeader(headers[cols.memQty]);
    const gFee = parseFloat(_event.GuestFee)  || _feeFromHeader(headers[cols.guestQty]);
    const kFee = parseFloat(_event.KidsFee)   || _feeFromHeader(headers[cols.kidsQty]);

    let nextNum  = maxNum;
    let imported = 0;

    for (const row of rows) {
      const ts = get(row, cols.ts);
      if (!ts || seenTs.has(ts)) continue;

      const mQty = parseInt(get(row, cols.memQty),   10) || 0;
      const gQty = parseInt(get(row, cols.guestQty), 10) || 0;
      const kQty = parseInt(get(row, cols.kidsQty),  10) || 0;
      if (mQty + gQty + kQty === 0) continue;

      nextNum++;
      await Sheets.append(CONFIG.SHEETS.REGISTRATIONS, {
        [C.ID]:        'REG-' + String(nextNum).padStart(4, '0'),
        [C.TS]:        ts,
        [C.SOURCE]:    'Google Form',
        [C.EVID]:      _eventId,
        [C.EVNAME]:    _event.Title,
        [C.LAST]:      get(row, cols.last),
        [C.FIRST]:     get(row, cols.first),
        [C.EMAIL]:     get(row, cols.email),
        [C.MKEY]:      '',
        [C.MEM_QTY]:   mQty,
        [C.GUEST_QTY]: gQty,
        [C.KIDS_QTY]:  kQty,
        [C.WALKIN]:    'No',
        [C.TOTAL]:     (mQty * mFee) + (gQty * gFee) + (kQty * kFee),
        [C.PAY_NOTE]:  get(row, cols.comments),
        [C.STATUS]:    'Pending',
        [C.PAY_MODE]:  '',
        [C.AMOUNT]:    '',
        [C.NOTES]:     get(row, cols.status) ? `Form status: ${get(row, cols.status)}` : '',
      });

      seenTs.add(ts); // guard against duplicate timestamps within the same form sheet
      imported++;
    }

    return imported;
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

      const safeId = Utils.escape(r[C.ID]);
      const mkey   = r[C.MKEY] || '';
      return `<tr>
        <td>
          <strong>${Utils.escape(name || '—')}</strong>
          ${r[C.EMAIL] ? `<br><span class="text-muted" style="font-size:12px;">${Utils.escape(r[C.EMAIL])}</span>` : ''}
          ${isWI ? ' <span class="badge badge-walkin">Walk-in</span>' : ''}
        </td>
        <td>
          ${mkey
            ? `<span style="font-family:monospace;font-size:12px;">${Utils.escape(mkey)}</span>`
            : `<span class="text-muted" style="font-size:12px;">—</span>`}
        </td>
        <td style="white-space:nowrap;">${pax}</td>
        <td class="amount">${Utils.formatPHP(r[C.TOTAL])}</td>
        <td><span class="badge badge-reg-${status.toLowerCase()}">${status}</span></td>
        <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"
            title="${Utils.escape(r[C.PAY_NOTE] || r[C.NOTES] || '')}">
          ${Utils.escape(r[C.PAY_NOTE] || r[C.NOTES] || '—')}
        </td>
        <td style="white-space:nowrap;">
          <button class="btn btn-sm btn-outline" onclick="Registrations.openEditRegistration('${safeId}')">
            Edit
          </button>
          ${status !== 'Confirmed' ? `
            <button class="btn btn-sm btn-primary" style="margin-left:4px;"
                    onclick="Registrations.openConfirmPayment('${safeId}')">
              Confirm
            </button>` : `
            <span class="text-muted" style="font-size:12px;margin-left:4px;">
              ${Utils.formatPHP(r[C.AMOUNT])} via ${Utils.escape(r[C.PAY_MODE] || '—')}
            </span>`}
          ${status !== 'Cancelled' ? `
            <button class="btn btn-sm btn-danger-outline" style="margin-left:4px;"
                    onclick="Registrations.cancelRegistration('${safeId}')">
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
    if (m) parts.push(`${m} Member${m !== 1 ? 's' : ''}`);
    if (g) parts.push(`${g} Guest${g !== 1 ? 's' : ''}`);
    if (k) parts.push(`${k} Kid${k !== 1 ? 's' : ''}`);
    return parts.join(' · ') || '—';
  }

  function _feeBreakdown(r) {
    const isWI = r[C.WALKIN] === 'Yes';
    const mFee = parseFloat(isWI ? (_event?.WalkInMemberFee || _event?.MemberFee) : _event?.MemberFee) || 0;
    const gFee = parseFloat(isWI ? (_event?.WalkInGuestFee  || _event?.GuestFee)  : _event?.GuestFee)  || 0;
    const kFee = parseFloat(_event?.KidsFee) || 0;
    const m    = parseInt(r[C.MEM_QTY],   10) || 0;
    const g    = parseInt(r[C.GUEST_QTY], 10) || 0;
    const k    = parseInt(r[C.KIDS_QTY],  10) || 0;
    const parts = [];
    if (m && mFee) parts.push(`${m} × ₱${mFee.toLocaleString()}`);
    if (g && gFee) parts.push(`${g} × ₱${gFee.toLocaleString()}`);
    if (k && kFee) parts.push(`${k} × ₱${kFee.toLocaleString()}`);
    const total  = (m * mFee) + (g * gFee) + (k * kFee);
    return { line: parts.join(' + '), total };
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

    if (!r[C.MKEY]) {
      Utils.toast('Assign a member in Edit before confirming payment.', 'error');
      return;
    }

    const name      = [r[C.LAST], r[C.FIRST]].filter(Boolean).join(', ') || regId;
    const breakdown = _feeBreakdown(r);
    const storedTotal = parseFloat(r[C.TOTAL]) || 0;
    const calcTotal   = breakdown.total;

    document.getElementById('reg-confirm-id').value         = regId;
    document.getElementById('reg-confirm-name').textContent = name;

    const ticketsEl = document.getElementById('reg-confirm-tickets');
    ticketsEl.innerHTML = `${Utils.escape(_paxSummary(r))}` +
      (breakdown.line ? `<br><span style="color:var(--text-muted)">${Utils.escape(breakdown.line)}</span>` : '');

    const dueEl = document.getElementById('reg-confirm-due');
    if (storedTotal && calcTotal && storedTotal !== calcTotal) {
      dueEl.innerHTML = `Total Due: <strong>${Utils.formatPHP(storedTotal)}</strong>` +
        ` <span style="color:var(--text-muted);font-size:12px;">(calculated: ${Utils.formatPHP(calcTotal)})</span>`;
    } else {
      dueEl.textContent = 'Total Due: ' + Utils.formatPHP(storedTotal || calcTotal);
    }

    document.getElementById('reg-confirm-amount').value = r[C.TOTAL] || (calcTotal || '');
    document.getElementById('reg-confirm-notes').value  = r[C.NOTES] || '';

    // Member assignment slots (shown when MemberQty > 1)
    const qty = parseInt(r[C.MEM_QTY], 10) || 1;
    const section = document.getElementById('reg-confirm-members-section');
    if (qty > 1) {
      const savedSlots = (r[C.SLOTS] || '').split(',').map(s => s.trim()).filter(Boolean);
      _renderConfirmSlots(qty, r[C.MKEY], savedSlots);
      section.style.display = 'block';
    } else {
      section.style.display = 'none';
    }

    Utils.showModal('reg-confirm-modal');
  }

  function _renderConfirmSlots(qty, primaryKey, savedSlots = []) {
    if (!_members.length) Sheets.getAll(CONFIG.SHEETS.MEMBERS).then(m => { _members = m; });
    const pm    = _members.find(m => m['Member Key'] === primaryKey);
    const pName = pm ? `${pm['First Name']} ${pm['Last Name']}`.trim() : primaryKey;

    const container = document.getElementById('reg-confirm-members');
    container.innerHTML = `
      <div class="confirm-slot">
        <span class="slot-num">1</span>
        <div class="slot-assigned">
          <span class="slot-name">${Utils.escape(pName)}</span>
          <span class="slot-key">${Utils.escape(primaryKey)}</span>
        </div>
        <input type="hidden" class="slot-key-input" value="${Utils.escape(primaryKey)}">
      </div>
      ${Array.from({ length: qty - 1 }, (_, i) => {
        const slotIdx = i + 1;
        const saved   = savedSlots[i];
        if (saved) {
          const sm    = _members.find(m => m['Member Key'] === saved);
          const sName = sm ? `${sm['First Name']} ${sm['Last Name']}`.trim() : saved;
          return `
          <div class="confirm-slot" id="confirm-slot-${slotIdx}">
            <span class="slot-num">${slotIdx + 1}</span>
            <div class="slot-search-wrap">
              <div class="slot-assigned">
                <span class="slot-name">${Utils.escape(sName)}</span>
                <span class="slot-key">${Utils.escape(saved)}</span>
                <button type="button" class="slot-clear" onclick="Registrations.clearConfirmSlot(${slotIdx})">✕</button>
              </div>
            </div>
            <input type="hidden" class="slot-key-input" value="${Utils.escape(saved)}">
          </div>`;
        }
        return `
        <div class="confirm-slot" id="confirm-slot-${slotIdx}">
          <span class="slot-num">${slotIdx + 1}</span>
          <div class="slot-search-wrap">
            <input type="text" class="form-control slot-search-input"
              placeholder="Search member…" autocomplete="off"
              oninput="Registrations.searchConfirmSlot(${slotIdx}, this.value)">
            <div class="slot-suggestions" id="confirm-slot-sugg-${slotIdx}"></div>
          </div>
          <input type="hidden" class="slot-key-input" value="">
        </div>`;
      }).join('')}`;
  }

  function searchConfirmSlot(slotIdx, query) {
    const suggBox = document.getElementById(`confirm-slot-sugg-${slotIdx}`);
    if (!suggBox) return;
    const q = query.trim();
    if (!q) { suggBox.innerHTML = ''; return; }

    const found = Utils.filterRows(_members, q, ['First Name','Last Name','Member Key']).slice(0, 5);
    const memberItems = found.map(m => {
      const key  = m['Member Key'];
      const name = `${m['First Name']} ${m['Last Name']}`.trim();
      return `<div class="slot-sugg-item" onclick="Registrations.selectConfirmSlot(${slotIdx},'${Utils.escape(key)}','${Utils.escape(name)}')">
        <strong>${Utils.escape(name)}</strong> <span>${Utils.escape(key)}</span>
      </div>`;
    });
    const addNew = `<div class="slot-sugg-item slot-add-new-opt"
        onclick="Registrations.openAddFromSlot(${slotIdx},'confirm','${Utils.escape(q)}')">
      + Add "<strong>${Utils.escape(q)}</strong>" as new member
    </div>`;
    suggBox.innerHTML = memberItems.join('') + addNew;
  }

  function selectConfirmSlot(slotIdx, key, name) {
    const slot = document.getElementById(`confirm-slot-${slotIdx}`);
    if (!slot) return;
    slot.querySelector('.slot-key-input').value = key;
    const wrap = slot.querySelector('.slot-search-wrap');
    wrap.innerHTML = `
      <div class="slot-assigned">
        <span class="slot-name">${Utils.escape(name)}</span>
        <span class="slot-key">${Utils.escape(key)}</span>
        <button type="button" class="slot-clear" onclick="Registrations.clearConfirmSlot(${slotIdx})">✕</button>
      </div>`;
    slot.querySelector('.slot-key-input').value = key;
  }

  function clearConfirmSlot(slotIdx) {
    const slot = document.getElementById(`confirm-slot-${slotIdx}`);
    if (!slot) return;
    slot.querySelector('.slot-key-input').value = '';
    slot.querySelector('.slot-search-wrap').innerHTML = `
      <input type="text" class="form-control slot-search-input"
        placeholder="Search member…" autocomplete="off"
        oninput="Registrations.searchConfirmSlot(${slotIdx}, this.value)">
      <div class="slot-suggestions" id="confirm-slot-sugg-${slotIdx}"></div>`;
  }

  // ── Edit-modal slot helpers ────────────────────────────────────────────────
  // Slot 0 renders into #reg-edit-slot0-wrap (inline with registrant fields).
  // Slots 1..N-1 render into #reg-edit-slots-container as "Member 2", "Member 3"…
  function _renderEditSlots(qty, primaryKey, savedSlots) {
    // Slot 0 — link-to-record field for Member 1
    const slot0Wrap = document.getElementById('reg-edit-slot0-wrap');
    if (slot0Wrap) {
      if (primaryKey) {
        const pm    = _members.find(m => m['Member Key'] === primaryKey);
        const pName = pm ? `${pm['First Name']} ${pm['Last Name']}`.trim() : primaryKey;
        slot0Wrap.innerHTML = `
          <div id="edit-slot-0">
            <div class="slot-search-wrap">
              <div class="slot-assigned">
                <span class="slot-name">${Utils.escape(pName)}</span>
                <span class="slot-key">${Utils.escape(primaryKey)}</span>
                <button type="button" class="slot-clear" onclick="Registrations.clearEditSlot(0)">✕</button>
              </div>
            </div>
            <input type="hidden" class="slot-key-input" value="${Utils.escape(primaryKey)}">
          </div>`;
      } else {
        slot0Wrap.innerHTML = `
          <div id="edit-slot-0">
            <div class="slot-search-wrap">
              <input type="text" class="form-control slot-search-input"
                placeholder="Search to link to member record (optional)…" autocomplete="off"
                oninput="Registrations.searchEditSlot(0, this.value)">
              <div class="slot-suggestions" id="edit-slot-sugg-0"></div>
            </div>
            <input type="hidden" class="slot-key-input" value="">
          </div>`;
      }
      slot0Wrap.style.display = 'block';
    }

    // Slots 1..N-1 — Additional Members
    const section   = document.getElementById('reg-edit-slots-section');
    const container = document.getElementById('reg-edit-slots-container');
    if (!container || !section) return;

    if (qty <= 1) { section.style.display = 'none'; return; }
    section.style.display = 'block';

    container.innerHTML = Array.from({ length: qty - 1 }, (_, i) => {
      const slotIdx = i + 1;
      const saved   = savedSlots[i];
      if (saved) {
        const sm    = _members.find(m => m['Member Key'] === saved);
        const sName = sm ? `${sm['First Name']} ${sm['Last Name']}`.trim() : saved;
        return `
        <div class="edit-member-group" id="edit-slot-${slotIdx}">
          <div class="edit-member-label">Member ${slotIdx + 1}</div>
          <div class="slot-search-wrap">
            <div class="slot-assigned">
              <span class="slot-name">${Utils.escape(sName)}</span>
              <span class="slot-key">${Utils.escape(saved)}</span>
              <button type="button" class="slot-clear" onclick="Registrations.clearEditSlot(${slotIdx})">✕</button>
            </div>
          </div>
          <input type="hidden" class="slot-key-input" value="${Utils.escape(saved)}">
        </div>`;
      }
      return `
      <div class="edit-member-group" id="edit-slot-${slotIdx}">
        <div class="edit-member-label">Member ${slotIdx + 1}</div>
        <div class="slot-search-wrap">
          <input type="text" class="form-control slot-search-input"
            placeholder="Search or add member…" autocomplete="off"
            oninput="Registrations.searchEditSlot(${slotIdx}, this.value)">
          <div class="slot-suggestions" id="edit-slot-sugg-${slotIdx}"></div>
        </div>
        <input type="hidden" class="slot-key-input" value="">
      </div>`;
    }).join('');
  }

  function searchEditSlot(slotIdx, query) {
    const suggBox = document.getElementById(`edit-slot-sugg-${slotIdx}`);
    if (!suggBox) return;
    const q = query.trim();
    if (!q) { suggBox.innerHTML = ''; return; }

    const found = Utils.filterRows(_members, q, ['First Name','Last Name','Member Key']).slice(0, 5);
    const primaryKey = (document.getElementById('reg-add-member-key')?.value || '').trim();
    const memberItems = found.map(m => {
      const key  = m['Member Key'];
      const name = `${m['First Name']} ${m['Last Name']}`.trim();
      return `<div class="slot-sugg-item" onclick="Registrations.selectEditSlot(${slotIdx},'${Utils.escape(key)}','${Utils.escape(name)}')">
        <strong>${Utils.escape(name)}</strong> <span>${Utils.escape(key)}</span>
      </div>`;
    });
    const addNew = `<div class="slot-sugg-item slot-add-new-opt"
        onclick="Registrations.openAddFromSlot(${slotIdx},'edit','${Utils.escape(q)}')">
      + Add "<strong>${Utils.escape(q)}</strong>" as new member
    </div>`;
    suggBox.innerHTML = memberItems.join('') + addNew;
  }

  function selectEditSlot(slotIdx, key, name) {
    const slot = document.getElementById(`edit-slot-${slotIdx}`);
    if (!slot) return;
    slot.querySelector('.slot-key-input').value = key;
    slot.querySelector('.slot-search-wrap').innerHTML = `
      <div class="slot-assigned">
        <span class="slot-name">${Utils.escape(name)}</span>
        <span class="slot-key">${Utils.escape(key)}</span>
        <button type="button" class="slot-clear" onclick="Registrations.clearEditSlot(${slotIdx})">✕</button>
      </div>`;
  }

  function clearEditSlot(slotIdx) {
    const slot = document.getElementById(`edit-slot-${slotIdx}`);
    if (!slot) return;
    slot.querySelector('.slot-key-input').value = '';
    slot.querySelector('.slot-search-wrap').innerHTML = `
      <input type="text" class="form-control slot-search-input"
        placeholder="Search or add member…" autocomplete="off"
        oninput="Registrations.searchEditSlot(${slotIdx}, this.value)">
      <div class="slot-suggestions" id="edit-slot-sugg-${slotIdx}"></div>`;
  }

  async function openAddFromSlot(slotIdx, context, query) {
    if (!_members.length) {
      _members = await Sheets.getAll(CONFIG.SHEETS.MEMBERS).catch(() => []);
    }

    // Parse name: last word → Last Name, rest → First Name
    const parts = query.trim().split(/\s+/);
    const last   = parts.length > 1 ? parts.pop() : '';
    const first  = parts.join(' ');

    // Auto-generate next member key
    const nums = _members
      .map(m => m['Member Key'])
      .filter(k => /^MBR-\d+$/.test(k))
      .map(k => parseInt(k.slice(4), 10));
    const nextNum = nums.length ? Math.max(...nums) + 1 : 1;
    const suggestedKey = `MBR-${String(nextNum).padStart(4, '0')}`;

    // Identify primary member for the family toggle label
    _slotAddContext = { type: context, slotIdx };
    if (context === 'confirm') {
      _slotAddPrimaryKey = document.querySelector('#reg-confirm-members .slot-key-input')?.value?.trim() || null;
    } else {
      // In edit mode, primary is slot 0 of the edit slots section
      _slotAddPrimaryKey = document.querySelector('#edit-slot-0 .slot-key-input')?.value?.trim() || null;
    }

    let primaryName = '';
    if (_slotAddPrimaryKey) {
      const pm = _members.find(m => m['Member Key'] === _slotAddPrimaryKey);
      primaryName = pm ? `${pm['First Name']} ${pm['Last Name']}`.trim() : _slotAddPrimaryKey;
    }

    document.getElementById('reg-slot-add-first').value = first;
    document.getElementById('reg-slot-add-last').value  = last;
    document.getElementById('reg-slot-add-key').value   = suggestedKey;
    document.getElementById('reg-slot-add-family').checked = !!_slotAddPrimaryKey;

    const famRow = document.getElementById('reg-slot-add-fam-row');
    const famLabel = document.getElementById('reg-slot-add-fam-label');
    if (famRow) famRow.style.display = _slotAddPrimaryKey ? '' : 'none';
    if (famLabel) famLabel.textContent = `Link as family under ${primaryName}`;

    Utils.showModal('reg-slot-add-modal');
  }

  async function saveAddFromSlot() {
    const btn = document.getElementById('reg-slot-add-save-btn');
    btn.disabled = true;
    try {
      const first      = document.getElementById('reg-slot-add-first').value.trim();
      const last       = document.getElementById('reg-slot-add-last').value.trim();
      const key        = document.getElementById('reg-slot-add-key').value.trim();
      const linkFamily = document.getElementById('reg-slot-add-family').checked;

      if (!last) { Utils.toast('Last name is required.', 'error'); btn.disabled = false; return; }
      if (!key)  { Utils.toast('Member Key is required.', 'error'); btn.disabled = false; return; }

      if (!_members.length) {
        _members = await Sheets.getAll(CONFIG.SHEETS.MEMBERS).catch(() => []);
      }
      if (_members.some(m => m['Member Key'] === key)) {
        Utils.toast('That Member Key is already in use.', 'error'); btn.disabled = false; return;
      }

      const famHead = linkFamily && _slotAddPrimaryKey ? _slotAddPrimaryKey : '';

      await Sheets.append(CONFIG.SHEETS.MEMBERS, {
        'Member Key':  key,
        'First Name':  first,
        'Last Name':   last,
        'Family Head': famHead,
        Status:        'Member',
      });

      // Make the primary member a self-referential head if not already set
      if (linkFamily && _slotAddPrimaryKey) {
        const primary = _members.find(m => m['Member Key'] === _slotAddPrimaryKey);
        if (primary && !primary['Family Head']) {
          await Sheets.update(CONFIG.SHEETS.MEMBERS, primary._rowIndex, {
            ...primary, 'Family Head': _slotAddPrimaryKey,
          });
        }
      }

      _members = []; // invalidate cache
      const name = [first, last].filter(Boolean).join(' ');
      Utils.hideModal('reg-slot-add-modal');
      const { type, slotIdx } = _slotAddContext || {};
      if (type === 'edit') {
        selectEditSlot(slotIdx, key, name);
      } else {
        selectConfirmSlot(slotIdx, key, name);
      }
      Utils.toast(`${name} added as ${key}${linkFamily ? ' · linked to family.' : '.'}`);
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  async function saveConfirmPayment() {
    const btn = document.getElementById('reg-confirm-save-btn');
    btn.disabled = true;
    try {
      const regId  = document.getElementById('reg-confirm-id').value;
      const amount = parseFloat(document.getElementById('reg-confirm-amount').value.trim()) || 0;
      const mode   = document.getElementById('reg-confirm-mode').value;
      const notes  = document.getElementById('reg-confirm-notes').value.trim();

      if (!amount) { Utils.toast('Please enter the amount paid.', 'error'); btn.disabled = false; return; }

      const r = _all.find(x => x[C.ID] === regId);
      if (!r) return;

      // Collect assigned member keys from slots (if member assignment section is visible)
      const section  = document.getElementById('reg-confirm-members-section');
      const slots    = section?.style.display !== 'none'
        ? [...document.querySelectorAll('#reg-confirm-members .slot-key-input')].map(el => el.value.trim()).filter(Boolean)
        : [r[C.MKEY]];

      // Create one transaction per assigned member.
      // Each secondary member gets the event member fee; slot 1 gets the remainder
      // (absorbs guest/kids fees or any adjustment).
      if (slots.length > 0) {
        const mFee     = parseFloat(_event?.MemberFee) || 0;
        const secTotal = mFee * (slots.length - 1);
        const now      = new Date();

        for (let i = 0; i < slots.length; i++) {
          const key    = slots[i];
          const member = _members.find(m => m['Member Key'] === key);
          const mName  = member
            ? `${member['First Name']} ${member['Last Name']}`.trim()
            : key;
          const txnAmt = i === 0 ? Math.max(0, amount - secTotal) : mFee;

          await Sheets.append(CONFIG.SHEETS.TRANSACTIONS, {
            TransactionID: await Sheets.nextId(CONFIG.SHEETS.TRANSACTIONS, 'TXN'),
            Timestamp:     now.toISOString(),
            MemberKey:     key,
            MemberName:    mName,
            EventID:       _eventId,
            EventName:     _event?.Title || '',
            AmountPaid:    txnAmt,
            PaymentMode:   mode,
            Category:      'Event',
            Year:          now.getFullYear(),
            Month:         now.getMonth() + 1,
            HeadCount:     1,
            Notes:         notes,
            RecordedBy:    Auth.getUserEmail(),
          });
        }
      }

      await Sheets.update(CONFIG.SHEETS.REGISTRATIONS, r._rowIndex, {
        ...r,
        [C.STATUS]:   'Confirmed',
        [C.PAY_MODE]: mode,
        [C.AMOUNT]:   amount,
        [C.NOTES]:    notes,
      });

      Utils.hideModal('reg-confirm-modal');
      Utils.toast(slots.length > 1
        ? `Payment confirmed · ${slots.length} member records created.`
        : 'Payment confirmed.');
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

  // ── Edit Registration (reuses Add Registration modal) ─────────────────────
  async function openEditRegistration(regId) {
    const r = _all.find(x => x[C.ID] === regId);
    if (!r) return;

    _editingRegId = regId;

    const titleEl = document.querySelector('#reg-add-modal .modal-header h2');
    const btnEl   = document.getElementById('reg-add-save-btn');
    if (titleEl) titleEl.textContent = 'Edit Registration';
    if (btnEl)   btnEl.textContent   = 'Save Changes';

    const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
    set('reg-add-last',       r[C.LAST]);
    set('reg-add-first',      r[C.FIRST]);
    set('reg-add-email',      r[C.EMAIL]);
    set('reg-add-member-key', r[C.MKEY]);
    set('reg-add-source',     r[C.SOURCE] || 'Manual');
    set('reg-add-member-qty', r[C.MEM_QTY]   || 0);
    set('reg-add-guest-qty',  r[C.GUEST_QTY] || 0);
    set('reg-add-kids-qty',   r[C.KIDS_QTY]  || 0);
    set('reg-add-total',      r[C.TOTAL]);
    set('reg-add-pay-note',   r[C.PAY_NOTE]);
    set('reg-add-status',     r[C.STATUS] || 'Pending');
    set('reg-add-mode',       r[C.PAY_MODE]);
    set('reg-add-amount',     r[C.AMOUNT]);
    set('reg-add-notes',      r[C.NOTES]);

    const walkinEl = document.getElementById('reg-add-walkin');
    if (walkinEl) walkinEl.checked = r[C.WALKIN] === 'Yes';

    _toggleAddPaymentFields(r[C.STATUS] === 'Confirmed');

    const hasKids = _event && parseFloat(_event.KidsFee) > 0;
    const kidsRow = document.getElementById('reg-add-kids-row');
    if (kidsRow) kidsRow.style.display = hasKids ? '' : 'none';

    // Show Edit-mode structural elements (hidden by default)
    const m1Head = document.getElementById('reg-edit-member1-head');
    if (m1Head) m1Head.style.display = 'block';

    // Load members, then render unified slot section
    if (!_members.length) {
      _members = await Sheets.getAll(CONFIG.SHEETS.MEMBERS).catch(() => []);
    }

    const qty        = Math.max(1, parseInt(r[C.MEM_QTY], 10) || 1);
    const savedSlots = (r[C.SLOTS] || '').split(',').map(s => s.trim()).filter(Boolean);
    _renderEditSlots(qty, r[C.MKEY], savedSlots);
    const editSlotsSection = document.getElementById('reg-edit-slots-section');
    if (editSlotsSection) editSlotsSection.style.display = 'block';

    Utils.showModal('reg-add-modal');
  }

  function _memberTypeCls(type) {
    return { Family: 'sug-family', Individual: 'sug-individual', Single: 'sug-single', Honorary: 'sug-honorary' }[type] || 'sug-other';
  }

  function _memberTypeBadge(m) {
    const type      = m['Membership Type'] || '';
    const isFamHead = m['Family Head'] === m['Member Key'];
    const label     = type === 'Family' ? (isFamHead ? 'Family' : 'Family member') : type;
    return label ? `<span class="suggestion-type ${_memberTypeCls(type)}">${Utils.escape(label)}</span>` : '';
  }

  function _memberTypeScore(m, wantsFamily) {
    const type      = m['Membership Type'] || '';
    const isFamily  = type === 'Family';
    const isFamHead = m['Family Head'] === m['Member Key'];
    if (wantsFamily) {
      if (isFamily && isFamHead) return 0;
      if (isFamily)              return 1;
      return 2;
    } else {
      if (!isFamily) return 0;
      if (isFamHead) return 1;
      return 2;
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

  // ── Member Key search ─────────────────────────────────────────────────────
  function searchMemberKey() {
    const last  = (document.getElementById('reg-add-last')?.value  || '').toLowerCase().trim();
    const first = (document.getElementById('reg-add-first')?.value || '').toLowerCase().trim();
    const box   = document.getElementById('reg-add-suggestions');
    if (!box) return;

    if (!last && !first) { box.style.display = 'none'; return; }

    const mQty        = parseInt(document.getElementById('reg-add-member-qty')?.value, 10) || 1;
    const wantsFamily = mQty > 1;

    const matches = _members
      .filter(m => {
        const mLast  = (m['Last Name']  || '').toLowerCase();
        const mFirst = (m['First Name'] || '').toLowerCase();
        const mAlt   = (m['Alternative Name'] || '').toLowerCase();
        const lastOk  = !last  || mLast.includes(last);
        const firstOk = !first || mFirst.includes(first) || mAlt.includes(first);
        return lastOk && firstOk;
      })
      .sort((a, b) => _memberTypeScore(a, wantsFamily) - _memberTypeScore(b, wantsFamily))
      .slice(0, 6);

    if (!matches.length) { box.style.display = 'none'; return; }

    box.innerHTML = matches.map(m => {
      const name = `${m['First Name'] || ''} ${m['Last Name'] || ''}`.trim();
      const key  = m['Member Key'] || '';
      return `<div class="member-key-suggestion-item"
                   onclick="Registrations.selectMemberSuggestion('${Utils.escape(key)}', '${Utils.escape(name)}')">
        <span class="suggestion-name">${Utils.escape(name)}</span>
        <span style="display:flex;gap:6px;align-items:center;flex-shrink:0;">
          ${_memberTypeBadge(m)}
          <span class="suggestion-key">${Utils.escape(key)}</span>
        </span>
      </div>`;
    }).join('');
    box.style.display = 'block';
  }

  function selectMemberSuggestion(key, name) {
    const keyEl = document.getElementById('reg-add-member-key');
    if (keyEl) keyEl.value = key;
    clearMemberSuggestions();
  }

  function clearMemberSuggestions() {
    const box = document.getElementById('reg-add-suggestions');
    if (box) box.style.display = 'none';
  }

  // ── Add Registration modal ────────────────────────────────────────────────
  function openAddRegistration() {
    _editingRegId = null;

    // Reset title and button
    const titleEl = document.querySelector('#reg-add-modal .modal-header h2');
    const btnEl   = document.getElementById('reg-add-save-btn');
    if (titleEl) titleEl.textContent = 'Add Registration';
    if (btnEl)   btnEl.textContent   = 'Save Registration';

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
    clearMemberSuggestions();
    const hasKids = _event && parseFloat(_event.KidsFee) > 0;
    const kidsRow = document.getElementById('reg-add-kids-row');
    if (kidsRow) kidsRow.style.display = hasKids ? '' : 'none';
    // Show member key field (hidden by default; only used in Add mode)
    const mkeyGroup = document.getElementById('reg-add-mkey-group');
    if (mkeyGroup) mkeyGroup.style.display = '';
    // Hide Edit-mode structural elements
    const m1Head = document.getElementById('reg-edit-member1-head');
    if (m1Head) m1Head.style.display = 'none';
    const slot0Wrap = document.getElementById('reg-edit-slot0-wrap');
    if (slot0Wrap) slot0Wrap.style.display = 'none';
    const editSlotsEl = document.getElementById('reg-edit-slots-section');
    if (editSlotsEl) editSlotsEl.style.display = 'none';

    if (!_members.length) {
      Sheets.getAll(CONFIG.SHEETS.MEMBERS).then(rows => { _members = rows; }).catch(() => {});
    }

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

      // Collect member assignments from slot UI (Edit mode) or Member Key field (Add mode)
      let memberKey, memberSlots;
      if (_editingRegId) {
        memberKey   = document.querySelector('#reg-edit-slot0-wrap .slot-key-input')?.value.trim() || '';
        memberSlots = [...document.querySelectorAll('#reg-edit-slots-container .slot-key-input')]
          .map(el => el.value.trim()).filter(Boolean).join(',');
      } else {
        memberKey   = get('reg-add-member-key');
        memberSlots = '';
      }

      const obj = {
        [C.SOURCE]:    source,
        [C.EVID]:      _eventId,
        [C.EVNAME]:    _event?.Title || '',
        [C.LAST]:      last,
        [C.FIRST]:     get('reg-add-first'),
        [C.EMAIL]:     get('reg-add-email'),
        [C.MKEY]:      memberKey,
        [C.SLOTS]:     memberSlots,
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
      };

      if (_editingRegId) {
        const existing = _all.find(x => x[C.ID] === _editingRegId);
        if (!existing) throw new Error('Registration not found.');
        await Sheets.update(CONFIG.SHEETS.REGISTRATIONS, existing._rowIndex, { ...existing, ...obj });

        // Sync registration email to member record if the member has none
        const regEmail = obj[C.EMAIL];
        if (memberKey && regEmail) {
          if (!_members.length) _members = await Sheets.getAll(CONFIG.SHEETS.MEMBERS).catch(() => []);
          const member = _members.find(m => m['Member Key'] === memberKey);
          if (member && !member['Email']) {
            await Sheets.update(CONFIG.SHEETS.MEMBERS, member._rowIndex, { ...member, Email: regEmail });
            _members = [];
            Utils.toast('Registration updated · Email saved to member record.');
          } else {
            Utils.toast('Registration updated.');
          }
        } else {
          Utils.toast('Registration updated.');
        }
      } else {
        const rows    = await Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).catch(() => []);
        const maxNum  = rows.map(r => parseInt((r[C.ID] || '').replace(/\D/g, ''), 10)).filter(n => !isNaN(n));
        const nextNum = maxNum.length ? Math.max(...maxNum) + 1 : 1;
        obj[C.ID] = 'REG-' + String(nextNum).padStart(4, '0');
        obj[C.TS] = new Date().toISOString();
        await Sheets.append(CONFIG.SHEETS.REGISTRATIONS, obj);
        Utils.toast('Registration saved.');
      }

      _editingRegId = null;
      Utils.hideModal('reg-add-modal');
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
    searchConfirmSlot, selectConfirmSlot, clearConfirmSlot,
    searchEditSlot, selectEditSlot, clearEditSlot,
    openAddFromSlot, saveAddFromSlot,
    cancelRegistration,
    openAddRegistration, onAddSourceChange, onAddWalkInChange,
    recalcAddTotal, onAddStatusChange, saveAddRegistration,
    openEditRegistration,
    searchMemberKey, selectMemberSuggestion, clearMemberSuggestions,
    openAddWalkIn, saveWalkIn,
    backToEvents,
  };
})();
