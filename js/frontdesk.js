/**
 * frontdesk.js — Front Desk (Door Check-In) Mode
 * Registration-centric: each search result is a registration card showing
 * all ticket slots (members, guests, kids) with per-slot check-in.
 * Walk-in members (not registered) fall through to the existing payment modal.
 */

const FrontDesk = (() => {
  let _event         = null;
  let _events        = [];
  let _members       = [];
  let _txns          = [];
  let _registrations = []; // event registrations for the current event
  let _famMembers    = [];
  let _famSelected   = new Set();
  let _famHead       = null;
  let _assignRegId   = null;
  let _assignSlotIdx = null;

  // ── Render / init ─────────────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading front desk…');
    try {
      [_events, _members, _txns] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.EVENTS),
        Sheets.getAll(CONFIG.SHEETS.MEMBERS),
        Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS),
      ]);
      _events = _events.filter(e => new Date(e.Date) >= new Date(Date.now() - 86400000 * 3));

      const preSelected = sessionStorage.getItem('fd_event');
      if (preSelected) {
        sessionStorage.removeItem('fd_event');
        await _selectEvent(preSelected);
      } else {
        _renderEventPicker();
      }
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  // ── Event Picker ──────────────────────────────────────────────────────────
  function _renderEventPicker() {
    const container = document.getElementById('fd-content');
    if (!container) return;

    if (!_events.length) {
      container.innerHTML = `
        <div class="fd-empty">
          <div class="fd-empty-icon">📅</div>
          <p>No recent or upcoming events found.</p>
          <button class="btn btn-primary" onclick="Router.navigate('events')">Create an Event</button>
        </div>`;
      return;
    }

    container.innerHTML = `
      <div class="fd-event-picker">
        <h2>Select Event</h2>
        <div class="event-picker-grid">
          ${_events.map(e => `
            <div class="event-pick-card" onclick="FrontDesk.selectEvent('${Utils.escape(e.EventID)}')">
              <div class="epc-date">${Utils.formatDate(e.Date)}</div>
              <div class="epc-title">${Utils.escape(e.Title)}</div>
              <div class="epc-loc">${Utils.escape(e.Location)}</div>
              <div class="epc-fee">Members: ${Utils.formatPHP(e.MemberFee)}</div>
            </div>`).join('')}
        </div>
      </div>`;
  }

  function selectEvent(id) { _selectEvent(id); }

  async function _selectEvent(id) {
    _event = _events.find(e => e.EventID === id) || null;
    if (!_event && _events.length) _event = _events[0];
    if (!_event) { _renderEventPicker(); return; }

    const allRegs = await Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).catch(() => []);
    _registrations = allRegs.filter(r => r.EventID === _event.EventID);

    _renderFrontDesk();
  }

  // ── Stats helpers ─────────────────────────────────────────────────────────
  function _checkedInCount() {
    return _registrations.reduce((sum, r) => {
      return sum + (r.CheckedIn || '').split(',').filter(Boolean).length;
    }, 0);
  }

  function _revenue() {
    return _txns
      .filter(t => t.EventID === _event.EventID &&
        (t.Category === 'Event' || t.Category === 'Walk-in Guest' || t.Category === 'Walk-in'))
      .reduce((s, t) => s + Utils.parsePHP(t.AmountPaid), 0);
  }

  // ── Main Front Desk UI ────────────────────────────────────────────────────
  function _renderFrontDesk() {
    const container = document.getElementById('fd-content');
    if (!container) return;

    container.innerHTML = `
      <div class="fd-header">
        <div class="fd-event-info">
          <button class="btn btn-sm fd-change-event-btn" onclick="FrontDesk.changeEvent()">⬅ Change Event</button>
          <div class="fd-event-name">${Utils.escape(_event.Title)}</div>
          <div class="fd-event-meta">${Utils.formatDate(_event.Date)} · ${Utils.escape(_event.Location)}</div>
        </div>
        <div class="fd-live-stats">
          <div class="fd-stat"><span id="fd-count">${_checkedInCount()}</span><label>Checked In</label></div>
          <div class="fd-stat"><span id="fd-revenue">${Utils.formatPHP(_revenue())}</span><label>Collected</label></div>
          <div class="fd-stat"><span>${Utils.formatPHP(_event.MemberFee)}</span><label>Member Fee</label></div>
          <div class="fd-stat"><span>${Utils.formatPHP(_event.GuestFee)}</span><label>Guest Fee</label></div>
        </div>
      </div>

      <div class="fd-search-bar">
        <input type="search" id="fd-search" class="fd-search-input"
               placeholder="🔍 Search by name, email, or member key…" autocomplete="off" autofocus>
        <button class="btn btn-outline fd-walkin-guest-btn" onclick="FrontDesk.openWalkinGuest()">
          + Walk-in Guest
        </button>
      </div>

      <div id="fd-results" class="fd-results"></div>

      <div class="fd-checkedin-list">
        <h3>Checked In (${_checkedInCount()})</h3>
        <div id="fd-checkedin"></div>
      </div>`;

    document.getElementById('fd-search')
      ?.addEventListener('input', Utils.debounce(e => _search(e.target.value), 200));

    _renderCheckedIn();
  }

  // ── Search ────────────────────────────────────────────────────────────────
  function _search(query) {
    const results = document.getElementById('fd-results');
    if (!results) return;
    const q = query.trim().toLowerCase();
    if (!q) { results.innerHTML = ''; return; }

    // Match registrations by registrant name/email OR any slot member key/name
    const matchedRegs = _registrations.filter(r => {
      const regName = `${r.LastName || ''} ${r.FirstName || ''}`.toLowerCase();
      if (regName.includes(q) || (r.Email || '').toLowerCase().includes(q)) return true;
      const keys = _slotKeys(r);
      return keys.some(key => {
        if (key.toLowerCase().includes(q)) return true;
        const m = _members.find(m => m['Member Key'] === key);
        return m && `${m['First Name'] || ''} ${m['Last Name'] || ''}`.toLowerCase().includes(q);
      });
    });

    // Walk-in members not covered by any registration
    const coveredKeys = new Set(_registrations.flatMap(_slotKeys));
    const walkIns = Utils.filterRows(_members, query, ['First Name','Last Name','Alternative Name','Member Key'])
      .filter(m => !coveredKeys.has(m['Member Key']))
      .slice(0, 3);

    if (!matchedRegs.length && !walkIns.length) {
      results.innerHTML = '<p class="fd-no-results">No registration or member found.</p>';
      return;
    }

    results.innerHTML = [
      ...matchedRegs.slice(0, 5).map(r => _regCardHtml(r)),
      ...walkIns.map(m => _memberCardHtml(m)),
    ].join('');
  }

  // All assigned member keys in a registration
  function _slotKeys(r) {
    return [r.MemberKey, ...(r.MemberSlots || '').split(',').map(s => s.trim())]
      .filter(Boolean);
  }

  // ── Registration Card ─────────────────────────────────────────────────────
  function _regCardHtml(reg) {
    const name         = [reg.LastName, reg.FirstName].filter(Boolean).join(', ');
    const status       = reg.PaymentStatus || 'Pending';
    const checkedSlots = new Set((reg.CheckedIn || '').split(',').filter(Boolean));
    const mQty         = parseInt(reg.MemberQty) || 0;
    const gQty         = parseInt(reg.GuestQty)  || 0;
    const kQty         = parseInt(reg.KidsQty)   || 0;
    const totalSlots   = mQty + gQty + kQty;
    const assignedKeys = _slotKeys(reg);

    const rows = [];

    // Member slots
    for (let i = 0; i < mQty; i++) {
      const slotId = `m${i}`;
      const key    = assignedKeys[i] || '';
      const m      = key ? _members.find(x => x['Member Key'] === key) : null;
      const label  = m
        ? `${m['First Name'] || ''} ${m['Last Name'] || ''}`.trim()
        : key || `Member ${i + 1} — unassigned`;
      const isIn   = checkedSlots.has(slotId);
      rows.push(_slotRowHtml(reg.RegistrationID, slotId, label, key, isIn, 'member'));
    }

    // Guest slots
    for (let i = 0; i < gQty; i++) {
      const slotId = `g${i + 1}`;
      const isIn   = checkedSlots.has(slotId);
      rows.push(_slotRowHtml(reg.RegistrationID, slotId, `Guest ${i + 1}`, '', isIn, 'guest'));
    }

    // Kids slots
    for (let i = 0; i < kQty; i++) {
      const slotId = `k${i + 1}`;
      const isIn   = checkedSlots.has(slotId);
      rows.push(_slotRowHtml(reg.RegistrationID, slotId, `Child ${i + 1}`, '', isIn, 'kids'));
    }

    const allIn    = checkedSlots.size >= totalSlots && totalSlots > 0;
    const safeId   = Utils.escape(reg.RegistrationID);
    const paidLine = status === 'Confirmed'
      ? `${Utils.formatPHP(reg.AmountPaid)} · ${Utils.escape(reg.PaymentMode || '—')}`
      : `${Utils.formatPHP(reg.TotalDue)} due`;

    return `<div class="fd-reg-card ${allIn ? 'fd-reg-all-in' : ''}" id="fd-reg-${safeId}">
      <div class="fd-reg-header">
        <div class="fd-reg-title">
          <span class="fd-reg-name">${Utils.escape(name)}</span>
          <span class="badge badge-reg-${status.toLowerCase()}">${Utils.escape(status)}</span>
        </div>
        <div class="fd-reg-meta">
          ${reg.Email ? `<span class="text-muted" style="font-size:12px;">${Utils.escape(reg.Email)}</span>` : ''}
          <span style="font-size:12px;">${paidLine}</span>
          <span class="fd-slot-count ${allIn ? 'all-in' : ''}">${checkedSlots.size}/${totalSlots} checked in</span>
        </div>
      </div>
      <div class="fd-slot-list">${rows.join('')}</div>
      ${status !== 'Confirmed' ? `
        <div class="fd-reg-footer">
          <button class="btn btn-primary btn-sm"
            onclick="FrontDesk.openRegPayment('${safeId}')">Collect Payment</button>
        </div>` : ''}
    </div>`;
  }

  function _slotRowHtml(regId, slotId, label, memberKey, isIn, type) {
    const avatarText = memberKey ? Utils.initials(label) : (type === 'guest' ? 'G' : type === 'kids' ? 'K' : '?');
    const safeRegId  = Utils.escape(regId);
    const safeSlot   = Utils.escape(slotId);
    return `<div class="fd-slot-row ${isIn ? 'slot-checked-in' : ''}" id="fd-slot-${safeRegId}-${safeSlot}">
      <div class="fd-avatar sm fd-slot-avatar ${type}">${avatarText}</div>
      <div class="fd-slot-info">
        <span class="fd-slot-name">${Utils.escape(label)}</span>
        ${memberKey ? `<span class="fd-slot-key">${Utils.escape(memberKey)}</span>` : ''}
      </div>
      ${isIn
        ? `<div class="fd-slot-in-wrap">
             <span class="fd-checked-badge">✅ In</span>
             <button class="fd-slot-undo-btn" onclick="FrontDesk.undoSlot('${safeRegId}','${safeSlot}')" title="Undo check-in">Undo</button>
           </div>`
        : `<div class="fd-slot-actions">
             ${!memberKey && type === 'member'
               ? `<button class="btn btn-sm btn-outline fd-assign-btn"
                    onclick="FrontDesk.openAssignSlot('${safeRegId}',${parseInt(safeSlot.slice(1))})">Assign</button>`
               : ''}
             <button class="btn btn-sm ${type === 'member' ? 'btn-primary' : 'btn-outline'} fd-slot-btn"
               onclick="FrontDesk.checkinSlot('${safeRegId}','${safeSlot}')">Check In</button>
           </div>`}
    </div>`;
  }

  // ── Walk-in member card (not in any registration) ─────────────────────────
  function _memberCardHtml(m) {
    const key       = m['Member Key'];
    const name      = `${m['First Name']} ${m['Last Name']}`.trim();
    const status    = m['Membership Status'] || 'TBC';
    const type      = m['Membership Type'] || '';
    const isExempt  = status.toLowerCase() === 'exempt';
    const famHead   = m['Family Head'] || '';
    const famPending = famHead
      ? _members.filter(fm => fm['Family Head'] === famHead).length
      : 0;
    const showFamBtn = famHead && famPending > 1;

    return `<div class="fd-member-card fd-walkin-card">
      <div class="fd-member-info">
        <div class="fd-avatar">${Utils.initials(name)}</div>
        <div>
          <div class="fd-member-name">${Utils.escape(name)}</div>
          <div class="fd-member-meta">
            ${Utils.statusBadge(status)} ${Utils.typeBadge(type)}
            <span class="badge" style="background:var(--text-muted);color:#fff;font-size:10px;">No registration</span>
          </div>
        </div>
      </div>
      <div class="fd-card-actions">
        <button class="btn btn-primary fd-quick-btn"
          onclick="FrontDesk.quickCheckin('${Utils.escape(key)}', this)">
          ✅ ${isExempt ? 'Exempt · ₱0' : Utils.formatPHP(_event.MemberFee) + ' · Cash'}
        </button>
        ${showFamBtn ? `<button class="btn fd-fam-btn"
          onclick="FrontDesk.openFamilyCheckin('${Utils.escape(key)}')">👨‍👩‍👧 Family</button>` : ''}
        <button class="btn fd-custom-btn" title="Customize"
          onclick="FrontDesk.openCheckin('${Utils.escape(key)}')">⚙</button>
      </div>
    </div>`;
  }

  // ── Slot check-in ─────────────────────────────────────────────────────────
  async function checkinSlot(regId, slotId) {
    const reg = _registrations.find(r => r.RegistrationID === regId);
    if (!reg) return;

    const current = (reg.CheckedIn || '').split(',').filter(Boolean);
    if (current.includes(slotId)) return;
    current.push(slotId);
    const newCheckedIn = current.join(',');

    // Optimistically update UI first
    reg.CheckedIn = newCheckedIn;
    const cardEl = document.getElementById(`fd-reg-${regId}`);
    if (cardEl) cardEl.outerHTML = _regCardHtml(reg);

    _updateLiveStats();
    _renderCheckedIn();

    try {
      await Sheets.update(CONFIG.SHEETS.REGISTRATIONS, reg._rowIndex, { ...reg, CheckedIn: newCheckedIn });
      // Resolve slot label for toast
      const mIdx = slotId.startsWith('m') ? parseInt(slotId.slice(1)) : -1;
      let label = slotId;
      if (mIdx >= 0) {
        const key = _slotKeys(reg)[mIdx];
        const m   = key ? _members.find(x => x['Member Key'] === key) : null;
        label = m ? `${m['First Name']} ${m['Last Name']}`.trim() : (key || `Member ${mIdx + 1}`);
      } else if (slotId.startsWith('g')) {
        label = `Guest ${slotId.slice(1)}`;
      } else if (slotId.startsWith('k')) {
        label = `Child ${slotId.slice(1)}`;
      }
      Utils.toast(`✅ ${label} checked in`);
    } catch (e) {
      // Roll back
      reg.CheckedIn = current.slice(0, -1).join(',');
      const cardEl2 = document.getElementById(`fd-reg-${regId}`);
      if (cardEl2) cardEl2.outerHTML = _regCardHtml(reg);
      _updateLiveStats();
      Utils.toast(e.message, 'error');
    }
  }

  // ── Collect payment for pending registration ──────────────────────────────
  function openRegPayment(regId) {
    const reg = _registrations.find(r => r.RegistrationID === regId);
    if (!reg) return;

    const name = [reg.LastName, reg.FirstName].filter(Boolean).join(', ');
    document.getElementById('checkin-member-name').textContent   = name;
    document.getElementById('checkin-avatar').textContent        = Utils.initials(name);
    document.getElementById('checkin-event-name').textContent    = _event.Title;
    document.getElementById('checkin-default-fee').textContent   = `Total Due: ${Utils.formatPHP(reg.TotalDue)}`;
    document.getElementById('checkin-member-key').value          = reg.MemberKey || '';
    document.getElementById('checkin-amount').value              = parseFloat(reg.TotalDue) || 0;
    document.getElementById('checkin-guests').value              = 0;
    document.getElementById('checkin-kids').value                = 0;
    document.getElementById('checkin-guest-fee').textContent     = `${Utils.formatPHP(_event.GuestFee)} each`;
    document.getElementById('checkin-kids-fee').textContent      = `${Utils.formatPHP(_event.KidsFee || 0)} each`;
    document.getElementById('checkin-mode').value                = 'Cash';
    document.getElementById('checkin-notes').value               = '';
    document.getElementById('checkin-notes').style.display       = 'none';
    document.getElementById('checkin-exempt-note').style.display = 'none';

    // Store reg ID so submitCheckin can find it
    document.getElementById('checkin-member-key').dataset.regId = regId;

    const kidsRow = document.getElementById('checkin-kids-row');
    if (kidsRow) kidsRow.style.display = _event.KidsFee ? 'flex' : 'none';

    document.querySelectorAll('.pay-pill').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.mode === 'Cash');
    });
    const notesToggle = document.querySelector('.notes-toggle');
    if (notesToggle) notesToggle.textContent = '+ Add note';

    _updateCheckinTotal();
    Utils.showModal('checkin-modal');
  }

  // ── Quick Check-In for walk-in members ────────────────────────────────────
  async function quickCheckin(memberKey, btn) {
    if (btn) btn.disabled = true;
    try {
      const member = _members.find(m => m['Member Key'] === memberKey);
      if (!member) return;
      const isExempt = (member['Membership Status'] || '').toLowerCase() === 'exempt';
      const amount   = isExempt ? 0 : (Utils.parsePHP(_event.MemberFee) || 0);
      const name     = `${member['First Name']} ${member['Last Name']}`.trim();

      const txnId = await Sheets.nextId(CONFIG.SHEETS.TRANSACTIONS, 'TXN');
      const now   = new Date();
      await Sheets.append(CONFIG.SHEETS.TRANSACTIONS, {
        TransactionID: txnId,
        Timestamp:     now.toISOString(),
        MemberKey:     memberKey,
        MemberName:    name,
        EventID:       _event.EventID,
        EventName:     _event.Title,
        AmountPaid:    amount,
        PaymentMode:   'Cash',
        Category:      'Event',
        Year:          now.getFullYear(),
        Month:         now.getMonth() + 1,
        HeadCount:     1,
        Notes:         'Walk-in',
        RecordedBy:    Auth.getUserEmail(),
      });

      _txns.push({ TransactionID: txnId, MemberKey: memberKey, MemberName: name,
        EventID: _event.EventID, AmountPaid: amount, PaymentMode: 'Cash', Category: 'Event', HeadCount: 1 });
      _updateLiveStats();
      _renderCheckedIn();
      document.getElementById('fd-search').value = '';
      document.getElementById('fd-results').innerHTML = '';
      document.getElementById('fd-search')?.focus();
      Utils.toast(`✅ ${name} · ${isExempt ? 'Exempt' : Utils.formatPHP(amount) + ' Cash'}`);
    } catch (e) {
      Utils.toast(e.message, 'error');
      if (btn) btn.disabled = false;
    }
  }

  // ── Custom check-in modal (walk-in members) ───────────────────────────────
  function openCheckin(memberKey) {
    const member = _members.find(m => m['Member Key'] === memberKey);
    if (!member) return;

    const name       = `${member['First Name']} ${member['Last Name']}`.trim();
    const defaultAmt = Utils.parsePHP(_event.MemberFee) || 0;
    const isExempt   = (member['Membership Status'] || '').toLowerCase() === 'exempt';

    document.getElementById('checkin-member-name').textContent    = name;
    document.getElementById('checkin-avatar').textContent          = Utils.initials(name);
    document.getElementById('checkin-event-name').textContent      = _event.Title;
    document.getElementById('checkin-default-fee').textContent     = `Default: ${Utils.formatPHP(_event.MemberFee)}`;
    document.getElementById('checkin-member-key').value            = memberKey;
    document.getElementById('checkin-member-key').dataset.regId    = '';
    document.getElementById('checkin-amount').value                = isExempt ? 0 : defaultAmt;
    document.getElementById('checkin-guests').value                = 0;
    document.getElementById('checkin-kids').value                  = 0;
    document.getElementById('checkin-guest-fee').textContent       = `${Utils.formatPHP(_event.GuestFee)} each`;
    document.getElementById('checkin-kids-fee').textContent        = `${Utils.formatPHP(_event.KidsFee || 0)} each`;
    document.getElementById('checkin-mode').value                  = 'Cash';
    document.getElementById('checkin-notes').value                 = '';
    document.getElementById('checkin-notes').style.display         = 'none';
    document.getElementById('checkin-exempt-note').style.display   = isExempt ? 'block' : 'none';

    const kidsRow = document.getElementById('checkin-kids-row');
    if (kidsRow) kidsRow.style.display = _event.KidsFee ? 'flex' : 'none';

    document.querySelectorAll('.pay-pill').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.mode === 'Cash');
    });
    const notesToggle = document.querySelector('.notes-toggle');
    if (notesToggle) notesToggle.textContent = '+ Add note';

    _updateCheckinTotal();
    Utils.showModal('checkin-modal');
  }

  function stepCount(fieldId, delta) {
    const el = document.getElementById(fieldId);
    if (!el) return;
    el.value = Math.max(0, (parseInt(el.value, 10) || 0) + delta);
    _updateCheckinTotal();
  }

  function selectPayMode(mode) {
    document.getElementById('checkin-mode').value = mode;
    document.querySelectorAll('.pay-pill').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.mode === mode);
    });
  }

  function toggleNotes() {
    const inp = document.getElementById('checkin-notes');
    const btn = document.querySelector('.notes-toggle');
    if (!inp) return;
    const shown = inp.style.display !== 'none';
    inp.style.display = shown ? 'none' : 'block';
    if (btn) btn.textContent = shown ? '+ Add note' : '− Remove note';
    if (!shown) inp.focus();
  }

  function _updateCheckinTotal() {
    const amount   = Utils.parsePHP(document.getElementById('checkin-amount')?.value || 0);
    const guests   = parseInt(document.getElementById('checkin-guests')?.value || 0, 10);
    const kids     = parseInt(document.getElementById('checkin-kids')?.value || 0, 10);
    const guestFee = Utils.parsePHP(_event?.GuestFee || 0);
    const kidsFee  = Utils.parsePHP(_event?.KidsFee || 0);
    const total    = amount + (guests * guestFee) + (kids * kidsFee);
    const el       = document.getElementById('checkin-total');
    if (el) el.textContent = Utils.formatPHP(total);
  }

  async function submitCheckin() {
    const btn = document.getElementById('checkin-submit-btn');
    btn.disabled = true;
    try {
      const memberKeyEl = document.getElementById('checkin-member-key');
      const memberKey   = memberKeyEl?.value || '';
      const regId       = memberKeyEl?.dataset.regId || '';
      const member      = memberKey ? _members.find(m => m['Member Key'] === memberKey) : null;
      const reg0        = regId ? _registrations.find(r => r.RegistrationID === regId) : null;
      const name        = member
        ? `${member['First Name']} ${member['Last Name']}`.trim()
        : (reg0 ? [reg0.LastName, reg0.FirstName].filter(Boolean).join(', ') : memberKey);
      const amount      = Utils.parsePHP(document.getElementById('checkin-amount')?.value || 0);
      const guests      = parseInt(document.getElementById('checkin-guests')?.value || 0, 10);
      const kids        = parseInt(document.getElementById('checkin-kids')?.value || 0, 10);
      const guestFee    = Utils.parsePHP(_event.GuestFee || 0);
      const kidsFee     = Utils.parsePHP(_event.KidsFee || 0);
      const total       = amount + (guests * guestFee) + (kids * kidsFee);
      const mode        = document.getElementById('checkin-mode')?.value || 'Cash';
      const notes       = document.getElementById('checkin-notes')?.value?.trim() || '';

      const txnId = await Sheets.nextId(CONFIG.SHEETS.TRANSACTIONS, 'TXN');
      const now   = new Date();
      await Sheets.append(CONFIG.SHEETS.TRANSACTIONS, {
        TransactionID: txnId,
        Timestamp:     now.toISOString(),
        MemberKey:     memberKey,
        MemberName:    name,
        EventID:       _event.EventID,
        EventName:     _event.Title,
        AmountPaid:    total,
        PaymentMode:   mode,
        Category:      'Event',
        Year:          now.getFullYear(),
        Month:         now.getMonth() + 1,
        HeadCount:     1 + guests + kids,
        Notes:         [notes, regId ? `Reg: ${regId}` : ''].filter(Boolean).join(' — '),
        RecordedBy:    Auth.getUserEmail(),
      });

      _txns.push({ TransactionID: txnId, MemberKey: memberKey, MemberName: name,
        EventID: _event.EventID, AmountPaid: total, PaymentMode: mode,
        Category: 'Event', HeadCount: 1 + guests + kids });

      // If triggered from a pending registration, mark slot 0 checked in + update status
      if (regId) {
        const reg = _registrations.find(r => r.RegistrationID === regId);
        if (reg) {
          const slots = (reg.CheckedIn || '').split(',').filter(Boolean);
          if (!slots.includes('m0')) slots.push('m0');
          reg.CheckedIn      = slots.join(',');
          reg.PaymentStatus  = 'Confirmed';
          reg.PaymentMode    = mode;
          reg.AmountPaid     = total;
          await Sheets.update(CONFIG.SHEETS.REGISTRATIONS, reg._rowIndex, { ...reg });
        }
      }

      Utils.hideModal('checkin-modal');
      _updateLiveStats();
      _renderCheckedIn();
      document.getElementById('fd-search').value = '';
      document.getElementById('fd-results').innerHTML = '';
      Utils.toast(`✅ ${name} checked in.`);
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  // ── Stats + Checked-In list ───────────────────────────────────────────────
  function _updateLiveStats() {
    document.getElementById('fd-count').textContent   = _checkedInCount();
    document.getElementById('fd-revenue').textContent = Utils.formatPHP(_revenue());
    // Update section header
    const h3 = document.querySelector('.fd-checkedin-list h3');
    if (h3) h3.textContent = `Checked In (${_checkedInCount()})`;
  }

  function _renderCheckedIn() {
    const container = document.getElementById('fd-checkedin');
    if (!container) return;

    const rows = [];
    _registrations.forEach(reg => {
      const slots = (reg.CheckedIn || '').split(',').filter(Boolean);
      slots.forEach(slotId => {
        const mIdx = slotId.startsWith('m') ? parseInt(slotId.slice(1)) : -1;
        let label = slotId;
        if (mIdx >= 0) {
          const key = _slotKeys(reg)[mIdx];
          const m   = key ? _members.find(x => x['Member Key'] === key) : null;
          label = m ? `${m['First Name']} ${m['Last Name']}`.trim() : (key || `Member ${mIdx + 1}`);
        } else if (slotId.startsWith('g')) {
          label = `Guest ${slotId.slice(1)}`;
        } else if (slotId.startsWith('k')) {
          label = `Child ${slotId.slice(1)}`;
        }
        const regName = [reg.LastName, reg.FirstName].filter(Boolean).join(', ');
        rows.push({ label, regName, status: reg.PaymentStatus, regId: reg.RegistrationID, slotId });
      });
    });

    // Walk-in transactions (member walk-ins + walk-in guests)
    const regKeys = new Set(_registrations.flatMap(_slotKeys));
    _txns.filter(t => t.EventID === _event?.EventID &&
        (t.Category === 'Event' || t.Category === 'Walk-in Guest' || t.Category === 'Walk-in') &&
        !regKeys.has(t.MemberKey))
      .forEach(t => rows.push({
        label:   t.MemberName || t.MemberKey,
        regName: t.Category === 'Walk-in Guest' ? 'Walk-in guest' : 'Walk-in',
        txnId:   t.TransactionID,
        amount:  Utils.parsePHP(t.AmountPaid),
      }));

    if (!rows.length) {
      container.innerHTML = '<p class="empty-state">No one checked in yet.</p>';
      return;
    }

    container.innerHTML = rows.slice(-30).reverse().map(r => `
      <div class="fd-checkedin-item">
        <div class="fd-avatar sm">${Utils.initials(r.label)}</div>
        <span class="fd-ci-name">${Utils.escape(r.label)}</span>
        <span class="fd-ci-reg">${Utils.escape(r.regName)}</span>
        ${r.txnId
          ? `<button class="fd-undo-btn" onclick="FrontDesk.undoCheckin('${Utils.escape(r.txnId)}')">Undo</button>`
          : `<button class="fd-undo-btn" onclick="FrontDesk.undoSlot('${Utils.escape(r.regId)}','${Utils.escape(r.slotId)}')">Undo</button>`}
      </div>`).join('');
  }

  // ── Undo slot check-in (registered attendee) ─────────────────────────────
  async function undoSlot(regId, slotId) {
    const reg = _registrations.find(r => r.RegistrationID === regId);
    if (!reg) return;
    const current = (reg.CheckedIn || '').split(',').filter(Boolean);
    if (!current.includes(slotId)) return;
    const updated = current.filter(s => s !== slotId);

    // Optimistic UI
    reg.CheckedIn = updated.join(',');
    const cardEl = document.getElementById(`fd-reg-${regId}`);
    if (cardEl) cardEl.outerHTML = _regCardHtml(reg);
    _updateLiveStats();
    _renderCheckedIn();

    try {
      await Sheets.update(CONFIG.SHEETS.REGISTRATIONS, reg._rowIndex, { ...reg, CheckedIn: reg.CheckedIn });
      Utils.toast('Check-in undone.');
    } catch (e) {
      // Roll back
      reg.CheckedIn = current.join(',');
      const cardEl2 = document.getElementById(`fd-reg-${regId}`);
      if (cardEl2) cardEl2.outerHTML = _regCardHtml(reg);
      _updateLiveStats();
      Utils.toast(e.message, 'error');
    }
  }

  // ── Undo (walk-in txns only) ──────────────────────────────────────────────
  async function undoCheckin(txnId) {
    const ok = await Utils.confirm('Remove this check-in? The transaction record will be deleted.');
    if (!ok) return;
    try {
      const txn = _txns.find(t => t.TransactionID === txnId);
      if (!txn) { Utils.toast('Transaction not found.', 'error'); return; }
      await Sheets.deleteRow(CONFIG.SHEETS.TRANSACTIONS, txn._rowIndex);
      _txns = _txns.filter(t => t.TransactionID !== txnId);
      _updateLiveStats();
      _renderCheckedIn();
      Utils.toast('Check-in removed.');
    } catch (e) {
      Utils.toast(e.message, 'error');
    }
  }

  // ── Walk-in guest / drop-in attendee ─────────────────────────────────────
  function openWalkinGuest() {
    if (!_event) return;
    document.getElementById('fd-wi-name').value       = '';
    document.getElementById('fd-wi-member-key').value = '';
    document.getElementById('fd-wi-amount').value     = parseFloat(_event.GuestFee) || 0;
    document.getElementById('fd-wi-notes').value      = '';
    document.querySelectorAll('.wi-pay-pill').forEach(b => {
      b.classList.toggle('active', b.dataset.mode === 'Cash');
    });
    document.getElementById('fd-wi-mode').value = 'Cash';
    Utils.showModal('fd-walkin-modal');
    document.getElementById('fd-wi-name')?.focus();
  }

  function selectWiMode(mode) {
    document.getElementById('fd-wi-mode').value = mode;
    document.querySelectorAll('.wi-pay-pill').forEach(b => {
      b.classList.toggle('active', b.dataset.mode === mode);
    });
  }

  async function submitWalkinGuest() {
    const btn    = document.getElementById('fd-wi-submit');
    const name   = document.getElementById('fd-wi-name')?.value.trim();
    if (!name) { Utils.toast('Enter a name for the walk-in guest.', 'error'); return; }
    btn.disabled = true;
    try {
      const memberKey = document.getElementById('fd-wi-member-key')?.value.trim() || '';
      const amount    = parseFloat(document.getElementById('fd-wi-amount')?.value) || 0;
      const mode      = document.getElementById('fd-wi-mode')?.value || 'Cash';
      const notes     = document.getElementById('fd-wi-notes')?.value.trim() || '';
      const txnId     = await Sheets.nextId(CONFIG.SHEETS.TRANSACTIONS, 'TXN');
      const now       = new Date();
      await Sheets.append(CONFIG.SHEETS.TRANSACTIONS, {
        TransactionID: txnId,
        Timestamp:     now.toISOString(),
        MemberKey:     memberKey,
        MemberName:    name,
        EventID:       _event.EventID,
        EventName:     _event.Title,
        AmountPaid:    amount,
        PaymentMode:   mode,
        Category:      'Walk-in Guest',
        Year:          now.getFullYear(),
        Month:         now.getMonth() + 1,
        HeadCount:     1,
        Notes:         notes || 'Door walk-in',
        RecordedBy:    Auth.getUserEmail(),
      });
      _txns.push({ TransactionID: txnId, MemberKey: memberKey, MemberName: name,
        EventID: _event.EventID, AmountPaid: amount, PaymentMode: mode,
        Category: 'Walk-in Guest', HeadCount: 1 });
      Utils.hideModal('fd-walkin-modal');
      _updateLiveStats();
      _renderCheckedIn();
      Utils.toast(`✅ ${name} logged${amount > 0 ? ' · ' + Utils.formatPHP(amount) : ''}`);
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  // ── Family Group Check-In (walk-in) ───────────────────────────────────────
  function openFamilyCheckin(memberKey) {
    const member = _members.find(m => m['Member Key'] === memberKey);
    if (!member) return;
    _famHead    = member['Family Head'] || memberKey;
    _famMembers = _members
      .filter(m => m['Family Head'] === _famHead)
      .sort((a, b) => (a['Member Key'] === _famHead ? -1 : b['Member Key'] === _famHead ? 1 : 0));

    if (_famMembers.length < 2) { Utils.toast('No other family members found.', 'error'); return; }

    _famSelected = new Set(_famMembers.map(m => m['Member Key']));

    const headMember = _famMembers.find(m => m['Member Key'] === _famHead);
    const groupLabel = headMember
      ? `${headMember['Last Name'] || headMember['First Name']} Family`
      : 'Family Group';
    document.getElementById('fd-fam-group').textContent = groupLabel;
    document.getElementById('fd-fam-guests').value      = 0;
    document.getElementById('fd-fam-guest-fee').textContent = Utils.formatPHP(_event.GuestFee) + ' each';
    document.getElementById('fd-fam-mode').value        = 'Cash';
    document.querySelectorAll('.fam-pay-pill').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.mode === 'Cash');
    });
    _renderFamilyMembers();
    _updateFamilyTotal();
    Utils.showModal('fd-family-modal');
  }

  function _renderFamilyMembers() {
    const container = document.getElementById('fd-fam-members');
    container.innerHTML = _famMembers.map(m => {
      const key      = m['Member Key'];
      const name     = `${m['First Name']} ${m['Last Name']}`.trim();
      const isHead   = key === _famHead;
      const isExempt = (m['Membership Status'] || '').toLowerCase() === 'exempt';
      const selected = _famSelected.has(key);
      return `<div class="fd-fam-member ${selected ? 'selected' : ''}"
          onclick="FrontDesk.toggleFamilyMember('${Utils.escape(key)}')">
        <div class="fd-fam-check">${selected ? '☑' : '☐'}</div>
        <div class="fd-avatar sm">${Utils.initials(name)}</div>
        <div class="fd-fam-info">
          <span class="fd-fam-name">${Utils.escape(name)}</span>
          ${isHead ? '<span class="fam-head-badge">Head</span>' : ''}
          ${isExempt ? '<span class="badge badge-exempt" style="font-size:10px;">Exempt</span>' : ''}
        </div>
        <span class="fd-fam-fee">${isExempt ? '₱0' : Utils.formatPHP(_event.MemberFee)}</span>
      </div>`;
    }).join('');
  }

  function toggleFamilyMember(key) {
    if (_famSelected.has(key)) _famSelected.delete(key);
    else _famSelected.add(key);
    _renderFamilyMembers();
    _updateFamilyTotal();
  }

  function _updateFamilyTotal() {
    const mFee   = Utils.parsePHP(_event?.MemberFee || 0);
    const gFee   = Utils.parsePHP(_event?.GuestFee  || 0);
    const guests = parseInt(document.getElementById('fd-fam-guests')?.value || 0, 10);
    let total    = 0;
    _famSelected.forEach(key => {
      const m = _famMembers.find(fm => fm['Member Key'] === key);
      if ((m?.['Membership Status'] || '').toLowerCase() !== 'exempt') total += mFee;
    });
    total += guests * gFee;
    const el = document.getElementById('fd-fam-total');
    if (el) el.textContent = Utils.formatPHP(total);
  }

  function stepFamGuests(delta) {
    const el = document.getElementById('fd-fam-guests');
    if (!el) return;
    el.value = Math.max(0, (parseInt(el.value, 10) || 0) + delta);
    _updateFamilyTotal();
  }

  function selectFamMode(mode) {
    document.getElementById('fd-fam-mode').value = mode;
    document.querySelectorAll('.fam-pay-pill').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.mode === mode);
    });
  }

  async function submitFamilyCheckin() {
    const btn = document.getElementById('fd-fam-submit');
    btn.disabled = true;
    try {
      const keys     = [..._famSelected];
      if (!keys.length) { Utils.toast('Select at least one member.', 'error'); btn.disabled = false; return; }
      const mode     = document.getElementById('fd-fam-mode')?.value || 'Cash';
      const guests   = parseInt(document.getElementById('fd-fam-guests')?.value || 0, 10);
      const mFee     = Utils.parsePHP(_event?.MemberFee || 0);
      const gFee     = Utils.parsePHP(_event?.GuestFee  || 0);
      const guestAmt = guests * gFee;
      const now      = new Date();
      const guestHolder = keys.includes(_famHead) ? _famHead : keys[0];

      for (const key of keys) {
        const m         = _famMembers.find(fm => fm['Member Key'] === key);
        const name      = m ? `${m['First Name']} ${m['Last Name']}`.trim() : key;
        const isExempt  = (m?.['Membership Status'] || '').toLowerCase() === 'exempt';
        const memberAmt = isExempt ? 0 : mFee;
        const extraAmt  = key === guestHolder ? guestAmt : 0;
        const txnId     = await Sheets.nextId(CONFIG.SHEETS.TRANSACTIONS, 'TXN');
        await Sheets.append(CONFIG.SHEETS.TRANSACTIONS, {
          TransactionID: txnId,
          Timestamp:     now.toISOString(),
          MemberKey:     key, MemberName: name,
          EventID:       _event.EventID, EventName: _event.Title,
          AmountPaid:    memberAmt + extraAmt,
          PaymentMode:   mode, Category: 'Event',
          Year:          now.getFullYear(), Month: now.getMonth() + 1,
          HeadCount:     1 + (key === guestHolder ? guests : 0),
          Notes:         'Walk-in family',
          RecordedBy:    Auth.getUserEmail(),
        });
        _txns.push({ TransactionID: txnId, MemberKey: key, MemberName: name,
          EventID: _event.EventID, AmountPaid: memberAmt + extraAmt,
          PaymentMode: mode, Category: 'Event', HeadCount: 1 });
      }

      Utils.hideModal('fd-family-modal');
      _updateLiveStats();
      _renderCheckedIn();
      document.getElementById('fd-search').value = '';
      document.getElementById('fd-results').innerHTML = '';
      document.getElementById('fd-search')?.focus();
      const headM   = _famMembers.find(m => m['Member Key'] === _famHead);
      const famName = headM ? `${headM['Last Name'] || headM['First Name']} Family` : 'Family';
      Utils.toast(`✅ ${famName} · ${keys.length} member${keys.length > 1 ? 's' : ''} checked in`);
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  // ── Assign member to unassigned slot ─────────────────────────────────────
  function openAssignSlot(regId, slotIdx) {
    _assignRegId   = regId;
    _assignSlotIdx = slotIdx;
    const reg  = _registrations.find(r => r.RegistrationID === regId);
    const name = reg ? [reg.LastName, reg.FirstName].filter(Boolean).join(', ') : regId;
    document.getElementById('fd-assign-reg-name').textContent = name;
    document.getElementById('fd-assign-slot-label').textContent = `Member ${slotIdx + 1}`;
    document.getElementById('fd-assign-search').value = '';
    document.getElementById('fd-assign-suggestions').innerHTML = '';
    Utils.showModal('fd-assign-modal');
    setTimeout(() => document.getElementById('fd-assign-search')?.focus(), 100);
  }

  function onAssignInput() {
    const q = document.getElementById('fd-assign-search').value.trim().toLowerCase();
    const box = document.getElementById('fd-assign-suggestions');
    if (!q) { box.innerHTML = ''; return; }
    const hits = _members
      .filter(m => {
        const full = `${m['First Name']} ${m['Last Name']} ${m['Member Key']}`.toLowerCase();
        return full.includes(q);
      })
      .slice(0, 6);
    if (!hits.length) {
      box.innerHTML = '<div class="fd-assign-empty">No members found</div>'; return;
    }
    box.innerHTML = hits.map(m => {
      const key  = Utils.escape(m['Member Key']);
      const name = Utils.escape(`${m['First Name']} ${m['Last Name']}`.trim());
      return `<div class="fd-assign-opt" onclick="FrontDesk.selectAssignMember('${key}','${name}')">
        <div class="fd-avatar sm" style="background:var(--navy);color:#fff;flex-shrink:0;">${Utils.initials(name)}</div>
        <div>
          <div style="font-size:13px;font-weight:600;">${name}</div>
          <div style="font-size:11px;color:var(--text-muted);font-family:monospace;">${key}</div>
        </div>
      </div>`;
    }).join('');
  }

  async function selectAssignMember(memberKey, memberName) {
    const reg = _registrations.find(r => r.RegistrationID === _assignRegId);
    if (!reg) return;

    const btn = document.getElementById('fd-assign-confirm');
    btn.disabled = true;

    try {
      if (_assignSlotIdx === 0) {
        reg.MemberKey = memberKey;
      } else {
        const slots = (reg.MemberSlots || '').split(',').map(s => s.trim());
        while (slots.length < _assignSlotIdx) slots.push('');
        slots[_assignSlotIdx - 1] = memberKey;
        reg.MemberSlots = slots.join(',');
      }

      await Sheets.update(CONFIG.SHEETS.REGISTRATIONS, reg._rowIndex, { ...reg });

      Utils.hideModal('fd-assign-modal');
      const cardEl = document.getElementById(`fd-reg-${_assignRegId}`);
      if (cardEl) cardEl.outerHTML = _regCardHtml(reg);
      Utils.toast(`${memberName} assigned to slot ${_assignSlotIdx + 1}.`);
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
      _assignRegId = null;
      _assignSlotIdx = null;
    }
  }

  function changeEvent() { _event = null; _registrations = []; _renderEventPicker(); }

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    document.getElementById('checkin-submit-btn')
      ?.addEventListener('click', submitCheckin);
    document.getElementById('checkin-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('checkin-modal'));
    document.getElementById('checkin-amount')
      ?.addEventListener('input', _updateCheckinTotal);
  }

  return {
    render, init, selectEvent, changeEvent,
    checkinSlot, undoSlot, openRegPayment,
    openAssignSlot, onAssignInput, selectAssignMember,
    openCheckin, quickCheckin, submitCheckin, undoCheckin, stepCount, selectPayMode, toggleNotes,
    openFamilyCheckin, toggleFamilyMember, stepFamGuests, selectFamMode, submitFamilyCheckin,
    openWalkinGuest, selectWiMode, submitWalkinGuest,
  };
})();
