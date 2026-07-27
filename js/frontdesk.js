/**
 * frontdesk.js — Front Desk (Door Check-In) Mode
 * Optimised for quick use at event entry:
 *  1. Select or auto-load an event
 *  2. Search for a member by name
 *  3. Record payment + guest count → writes to Transactions
 */

const FrontDesk = (() => {
  let _event    = null;   // selected event object
  let _events   = [];
  let _members  = [];
  let _txns     = [];
  let _checkedIn  = new Set(); // member keys already checked in this session
  let _famMembers = [];         // family members shown in family modal
  let _famSelected = new Set(); // selected keys in family modal
  let _famHead    = null;       // family head key

  const PAYMENT_MODES = ['Cash','GCash','BDO','PayPal','Bank Transfer','Other'];

  // ── Render / init ─────────────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading front desk…');
    try {
      [_events, _members, _txns] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.EVENTS),
        Sheets.getAll(CONFIG.SHEETS.MEMBERS),
        Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS),
      ]);
      _events = _events.filter(e => new Date(e.Date) >= new Date(Date.now() - 86400000 * 3)); // last 3 days + future

      // Check if an event was pre-selected from the Events page
      const preSelected = sessionStorage.getItem('fd_event');
      if (preSelected) {
        sessionStorage.removeItem('fd_event');
        _selectEvent(preSelected);
      } else {
        _renderEventPicker();
      }
      _buildCheckedInSet();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  function _buildCheckedInSet() {
    if (!_event) return;
    _checkedIn = new Set(
      _txns.filter(t => t.EventID === _event.EventID && t.Category === 'Event').map(t => t.MemberKey)
    );
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

  function _selectEvent(id) {
    _event = _events.find(e => e.EventID === id) || null;
    if (!_event && _events.length) _event = _events[0];
    if (!_event) { _renderEventPicker(); return; }
    _buildCheckedInSet();
    _renderFrontDesk();
  }

  // ── Main Front Desk UI ────────────────────────────────────────────────────
  function _renderFrontDesk() {
    const container = document.getElementById('fd-content');
    if (!container) return;

    const checkedCount = _checkedIn.size;
    const revenue      = _txns
      .filter(t => t.EventID === _event.EventID && t.Category === 'Event')
      .reduce((s,t) => s + Utils.parsePHP(t.AmountPaid), 0);

    container.innerHTML = `
      <div class="fd-header">
        <div class="fd-event-info">
          <button class="btn btn-sm btn-outline" onclick="FrontDesk.changeEvent()">⬅ Change Event</button>
          <div class="fd-event-name">${Utils.escape(_event.Title)}</div>
          <div class="fd-event-meta">${Utils.formatDate(_event.Date)} · ${Utils.escape(_event.Location)}</div>
        </div>
        <div class="fd-live-stats">
          <div class="fd-stat"><span id="fd-count">${checkedCount}</span><label>Checked In</label></div>
          <div class="fd-stat"><span id="fd-revenue">${Utils.formatPHP(revenue)}</span><label>Collected</label></div>
          <div class="fd-stat"><span>${Utils.formatPHP(_event.MemberFee)}</span><label>Member Fee</label></div>
          <div class="fd-stat"><span>${Utils.formatPHP(_event.GuestFee)}</span><label>Guest Fee</label></div>
        </div>
      </div>

      <div class="fd-search-bar">
        <input type="search" id="fd-search" class="fd-search-input"
               placeholder="🔍 Search member by name or key…" autocomplete="off" autofocus>
      </div>

      <div id="fd-results" class="fd-results"></div>

      <div class="fd-checkedin-list">
        <h3>Checked In (${checkedCount})</h3>
        <div id="fd-checkedin"></div>
      </div>`;

    // Wire up search
    document.getElementById('fd-search')
      ?.addEventListener('input', Utils.debounce(e => _searchMembers(e.target.value), 200));

    _renderCheckedIn();
  }

  // ── Member Search ─────────────────────────────────────────────────────────
  function _searchMembers(query) {
    const results = document.getElementById('fd-results');
    if (!results) return;
    if (!query.trim()) { results.innerHTML = ''; return; }

    const found = Utils.filterRows(_members, query, [
      'First Name','Last Name','Alternative Name','Member Key','Family Group',
    ]).slice(0, 8);

    if (!found.length) {
      results.innerHTML = '<p class="fd-no-results">No member found.</p>';
      return;
    }

    results.innerHTML = found.map(m => {
      const key       = m['Member Key'];
      const name      = `${m['First Name']} ${m['Last Name']}`.trim();
      const status    = m['Membership Status'] || 'TBC';
      const type      = m['Membership Type'] || '';
      const fam       = m['Family Group'] || '';
      const famHead   = m['Family Head'] || '';
      const alreadyIn = _checkedIn.has(key);
      const isExempt  = status.toLowerCase() === 'exempt';
      const quickLabel = isExempt ? 'Exempt · ₱0' : Utils.formatPHP(_event.MemberFee) + ' · Cash';

      // Use Family Head field (reliable key-based link) to find family members
      const famPending = famHead
        ? _members.filter(fm => fm['Family Head'] === famHead && !_checkedIn.has(fm['Member Key'])).length
        : 0;
      const showFamBtn = famHead && famPending > 1;

      return `<div class="fd-member-card ${alreadyIn ? 'already-in' : ''}">
        <div class="fd-member-info">
          <div class="fd-avatar">${Utils.initials(name)}</div>
          <div>
            <div class="fd-member-name">${Utils.escape(name)}</div>
            <div class="fd-member-meta">
              ${Utils.statusBadge(status)} ${Utils.typeBadge(type)}
              ${fam ? `<span class="badge badge-fam">👨‍👩‍👧 ${Utils.escape(fam)}</span>` : ''}
            </div>
          </div>
        </div>
        ${alreadyIn
          ? `<span class="fd-checked-badge">✅ Checked In</span>`
          : `<div class="fd-card-actions">
               <button class="btn btn-primary fd-quick-btn"
                 onclick="FrontDesk.quickCheckin('${Utils.escape(key)}', this)">
                 ✅ ${Utils.escape(quickLabel)}
               </button>
               ${showFamBtn ? `<button class="btn fd-fam-btn"
                 onclick="FrontDesk.openFamilyCheckin('${Utils.escape(key)}')">
                 👨‍👩‍👧 Family
               </button>` : ''}
               <button class="btn fd-custom-btn" title="Customize amount, guests, payment"
                 onclick="FrontDesk.openCheckin('${Utils.escape(key)}')">
                 ⚙
               </button>
             </div>`}
      </div>`;
    }).join('');
  }

  // ── Quick Check-In (no modal — cash, default fee, no guests) ─────────────
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
        Notes:         '',
        RecordedBy:    Auth.getUserEmail(),
      });

      _checkedIn.add(memberKey);
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

  // ── Stepper for quantity fields ───────────────────────────────────────────
  function stepCount(fieldId, delta) {
    const el = document.getElementById(fieldId);
    if (!el) return;
    el.value = Math.max(0, (parseInt(el.value, 10) || 0) + delta);
    _updateCheckinTotal();
  }

  // ── Payment pill toggle ───────────────────────────────────────────────────
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

  // ── Check-in Payment Modal ────────────────────────────────────────────────
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
    document.getElementById('checkin-amount').value                = isExempt ? 0 : defaultAmt;
    document.getElementById('checkin-guests').value                = 0;
    document.getElementById('checkin-kids').value                  = 0;
    document.getElementById('checkin-guest-fee').textContent       = `${Utils.formatPHP(_event.GuestFee)} each`;
    document.getElementById('checkin-kids-fee').textContent        = `${Utils.formatPHP(_event.KidsFee || 0)} each`;
    document.getElementById('checkin-mode').value                  = 'Cash';
    document.getElementById('checkin-notes').value                 = '';
    document.getElementById('checkin-notes').style.display         = 'none';
    document.getElementById('checkin-exempt-note').style.display   = isExempt ? 'block' : 'none';

    // Show kids row only if event has a kids fee set
    const kidsRow = document.getElementById('checkin-kids-row');
    if (kidsRow) kidsRow.style.display = _event.KidsFee ? 'flex' : 'none';

    const notesToggle = document.querySelector('.notes-toggle');
    if (notesToggle) notesToggle.textContent = '+ Add note';

    // Reset payment pills to Cash
    document.querySelectorAll('.pay-pill').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.mode === 'Cash');
    });

    _updateCheckinTotal();
    Utils.showModal('checkin-modal');
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
      const memberKey = document.getElementById('checkin-member-key')?.value;
      const member    = _members.find(m => m['Member Key'] === memberKey);
      const name      = member
        ? `${member['First Name']} ${member['Last Name']}`.trim()
        : memberKey;
      const amount    = Utils.parsePHP(document.getElementById('checkin-amount')?.value || 0);
      const guests    = parseInt(document.getElementById('checkin-guests')?.value || 0, 10);
      const kids      = parseInt(document.getElementById('checkin-kids')?.value || 0, 10);
      const guestFee  = Utils.parsePHP(_event.GuestFee || 0);
      const kidsFee   = Utils.parsePHP(_event.KidsFee || 0);
      const total     = amount + (guests * guestFee) + (kids * kidsFee);
      const mode      = document.getElementById('checkin-mode')?.value || 'Cash';
      const notes     = document.getElementById('checkin-notes')?.value?.trim() || '';

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
        Notes:         notes,
        RecordedBy:    Auth.getUserEmail(),
      });

      _checkedIn.add(memberKey);
      _txns.push({
        TransactionID: txnId, MemberKey: memberKey, MemberName: name,
        EventID: _event.EventID, AmountPaid: total, PaymentMode: mode,
        Category: 'Event', HeadCount: 1 + guests + kids,
      });

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

  function _updateLiveStats() {
    const revenue = _txns
      .filter(t => t.EventID === _event.EventID && t.Category === 'Event')
      .reduce((s,t) => s + Utils.parsePHP(t.AmountPaid), 0);
    document.getElementById('fd-count').textContent   = _checkedIn.size;
    document.getElementById('fd-revenue').textContent = Utils.formatPHP(revenue);
  }

  function _renderCheckedIn() {
    const container = document.getElementById('fd-checkedin');
    if (!container) return;
    const list = _txns
      .filter(t => t.EventID === _event?.EventID && t.Category === 'Event')
      .slice(-20).reverse();
    if (!list.length) { container.innerHTML = '<p class="empty-state">No one checked in yet.</p>'; return; }
    container.innerHTML = list.map(t => `
      <div class="fd-checkedin-item">
        <div class="fd-avatar sm">${Utils.initials(t.MemberName || '')}</div>
        <span class="fd-ci-name">${Utils.escape(t.MemberName || t.MemberKey)}</span>
        <span class="amount">${Utils.formatPHP(t.AmountPaid)}</span>
        <span class="fd-mode">${Utils.escape(t.PaymentMode || '')}</span>
        <button class="fd-undo-btn" title="Undo check-in"
          onclick="FrontDesk.undoCheckin('${Utils.escape(t.TransactionID)}')">↩ Undo</button>
      </div>`).join('');
  }

  // ── Undo a check-in ───────────────────────────────────────────────────────
  async function undoCheckin(txnId) {
    const ok = await Utils.confirm('Remove this check-in? The transaction record will be deleted.');
    if (!ok) return;
    try {
      // Re-fetch to get row index (locally-added items don't carry _rowIndex)
      const allTxns = await Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS);
      const txn = allTxns.find(t => t.TransactionID === txnId);
      if (!txn) { Utils.toast('Transaction not found.', 'error'); return; }

      await Sheets.deleteRow(CONFIG.SHEETS.TRANSACTIONS, txn._rowIndex);

      const memberKey = txn.MemberKey;
      _txns = _txns.filter(t => t.TransactionID !== txnId);

      const stillIn = _txns.some(t =>
        t.EventID === _event.EventID && t.Category === 'Event' && t.MemberKey === memberKey
      );
      if (!stillIn) _checkedIn.delete(memberKey);

      _updateLiveStats();
      _renderCheckedIn();
      Utils.toast('Check-in removed.');
    } catch (e) {
      Utils.toast(e.message, 'error');
    }
  }

  // ── Family Group Check-In ─────────────────────────────────────────────────
  function openFamilyCheckin(memberKey) {
    const member = _members.find(m => m['Member Key'] === memberKey);
    if (!member) return;

    _famHead = member['Family Head'] || memberKey;

    _famMembers = _members
      .filter(m => m['Family Head'] === _famHead)
      .sort((a, b) => {
        const aHead = a['Member Key'] === _famHead;
        const bHead = b['Member Key'] === _famHead;
        return aHead === bHead ? 0 : aHead ? -1 : 1;
      });

    if (_famMembers.length < 2) { Utils.toast('No other family members found.', 'error'); return; }

    // Pre-select all members not yet checked in
    _famSelected = new Set(
      _famMembers.filter(m => !_checkedIn.has(m['Member Key'])).map(m => m['Member Key'])
    );

    const headMember = _famMembers.find(m => m['Member Key'] === _famHead);
    const groupLabel = headMember
      ? `${headMember['Last Name'] || headMember['First Name']} Family`
      : (member['Family Group'] || 'Family Group');
    document.getElementById('fd-fam-group').textContent = groupLabel;
    document.getElementById('fd-fam-guests').value        = 0;
    document.getElementById('fd-fam-guest-fee').textContent = Utils.formatPHP(_event.GuestFee) + ' each';
    document.getElementById('fd-fam-mode').value          = 'Cash';
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
      const key       = m['Member Key'];
      const name      = `${m['First Name']} ${m['Last Name']}`.trim();
      const isHead    = key === m['Family Head'];
      const isExempt  = (m['Membership Status'] || '').toLowerCase() === 'exempt';
      const alreadyIn = _checkedIn.has(key);
      const selected  = _famSelected.has(key);

      if (alreadyIn) {
        return `<div class="fd-fam-member fd-fam-done">
          <div class="fd-fam-check">✅</div>
          <div class="fd-avatar sm">${Utils.initials(name)}</div>
          <div class="fd-fam-info">
            <span class="fd-fam-name">${Utils.escape(name)}</span>
            ${isHead ? '<span class="fam-head-badge">Head</span>' : ''}
          </div>
          <span class="fd-fam-already">Already checked in</span>
        </div>`;
      }

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
    const gFee   = Utils.parsePHP(_event?.GuestFee || 0);
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
      const keys   = [..._famSelected];
      if (!keys.length) { Utils.toast('Select at least one member.', 'error'); btn.disabled = false; return; }

      const mode     = document.getElementById('fd-fam-mode')?.value || 'Cash';
      const guests   = parseInt(document.getElementById('fd-fam-guests')?.value || 0, 10);
      const mFee     = Utils.parsePHP(_event?.MemberFee || 0);
      const gFee     = Utils.parsePHP(_event?.GuestFee || 0);
      const guestAmt = guests * gFee;
      const now      = new Date();

      // Guest fees go to the family head if selected, otherwise first selected member
      const guestHolder = keys.includes(_famHead) ? _famHead : keys[0];

      for (const key of keys) {
        const m        = _famMembers.find(fm => fm['Member Key'] === key);
        const name     = m ? `${m['First Name']} ${m['Last Name']}`.trim() : key;
        const isExempt = (m?.['Membership Status'] || '').toLowerCase() === 'exempt';
        const memberAmt = isExempt ? 0 : mFee;
        const extraAmt  = key === guestHolder ? guestAmt : 0;

        const txnId = await Sheets.nextId(CONFIG.SHEETS.TRANSACTIONS, 'TXN');
        await Sheets.append(CONFIG.SHEETS.TRANSACTIONS, {
          TransactionID: txnId,
          Timestamp:     now.toISOString(),
          MemberKey:     key,
          MemberName:    name,
          EventID:       _event.EventID,
          EventName:     _event.Title,
          AmountPaid:    memberAmt + extraAmt,
          PaymentMode:   mode,
          Category:      'Event',
          Year:          now.getFullYear(),
          Month:         now.getMonth() + 1,
          HeadCount:     1 + (key === guestHolder ? guests : 0),
          Notes:         '',
          RecordedBy:    Auth.getUserEmail(),
        });

        _checkedIn.add(key);
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

  function changeEvent() { _event = null; _renderEventPicker(); }

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    document.getElementById('checkin-submit-btn')
      ?.addEventListener('click', submitCheckin);
    document.getElementById('checkin-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('checkin-modal'));
    document.getElementById('checkin-amount')
      ?.addEventListener('input', _updateCheckinTotal);
  }

  return { render, init, selectEvent, openCheckin, quickCheckin, submitCheckin, undoCheckin,
    openFamilyCheckin, toggleFamilyMember, stepFamGuests, selectFamMode, submitFamilyCheckin,
    changeEvent, stepCount, selectPayMode, toggleNotes };
})();
