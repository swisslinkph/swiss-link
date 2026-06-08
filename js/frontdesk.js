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
  let _checkedIn = new Set(); // member keys already checked in this session

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
      'Full Name','First Name','Last Name','Member Key','Family Group',
    ]).slice(0, 8);

    if (!found.length) {
      results.innerHTML = '<p class="fd-no-results">No member found. You can check in a guest instead.</p>';
      return;
    }

    results.innerHTML = found.map(m => {
      const key       = m['Member Key'];
      const name      = m['Full Name'] || `${m['First Name']} ${m['Last Name']}`;
      const status    = m['2026 Membership Status (Member, Exempt, Non-member, TBC)'] || 'TBC';
      const type      = m['Membership Type'] || '';
      const fam       = m['Family Group'] || '';
      const alreadyIn = _checkedIn.has(key);

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
          : `<button class="btn btn-primary fd-checkin-btn"
               onclick="FrontDesk.openCheckin('${Utils.escape(key)}')">
               Check In
             </button>`}
      </div>`;
    }).join('');
  }

  // ── Check-in Payment Modal ────────────────────────────────────────────────
  function openCheckin(memberKey) {
    const member = _members.find(m => m['Member Key'] === memberKey);
    if (!member) return;

    const name        = member['Full Name'] || `${member['First Name']} ${member['Last Name']}`;
    const defaultAmt  = Utils.parsePHP(_event.MemberFee) || 0;
    const isExempt    = (member['2026 Membership Status (Member, Exempt, Non-member, TBC)'] || '').toLowerCase() === 'exempt';

    document.getElementById('checkin-member-name').textContent  = name;
    document.getElementById('checkin-event-name').textContent   = _event.Title;
    document.getElementById('checkin-default-fee').textContent  = Utils.formatPHP(_event.MemberFee);
    document.getElementById('checkin-member-key').value         = memberKey;
    document.getElementById('checkin-amount').value             = isExempt ? 0 : defaultAmt;
    document.getElementById('checkin-guests').value             = 0;
    document.getElementById('checkin-guest-fee').textContent    = Utils.formatPHP(_event.GuestFee);
    document.getElementById('checkin-mode').value               = 'Cash';
    document.getElementById('checkin-notes').value              = '';
    document.getElementById('checkin-exempt-note').style.display = isExempt ? 'block' : 'none';

    _updateCheckinTotal();
    Utils.showModal('checkin-modal');
  }

  function _updateCheckinTotal() {
    const amount   = Utils.parsePHP(document.getElementById('checkin-amount')?.value || 0);
    const guests   = parseInt(document.getElementById('checkin-guests')?.value || 0, 10);
    const guestFee = Utils.parsePHP(_event?.GuestFee || 0);
    const total    = amount + (guests * guestFee);
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
        ? (member['Full Name'] || `${member['First Name']} ${member['Last Name']}`)
        : memberKey;
      const amount    = Utils.parsePHP(document.getElementById('checkin-amount')?.value || 0);
      const guests    = parseInt(document.getElementById('checkin-guests')?.value || 0, 10);
      const guestFee  = Utils.parsePHP(_event.GuestFee || 0);
      const guestAmt  = guests * guestFee;
      const mode      = document.getElementById('checkin-mode')?.value || 'Cash';
      const notes     = document.getElementById('checkin-notes')?.value?.trim() || '';

      // Main entry
      const txnId = await Sheets.nextId(CONFIG.SHEETS.TRANSACTIONS, 'TXN');
      await Sheets.append(CONFIG.SHEETS.TRANSACTIONS, {
        TransactionID: txnId,
        Timestamp:     new Date().toISOString(),
        MemberKey:     memberKey,
        MemberName:    name,
        EventID:       _event.EventID,
        EventName:     _event.Title,
        AmountPaid:    amount + guestAmt,
        PaymentMode:   mode,
        Category:      'Event',
        Year:          new Date().getFullYear(),
        HeadCount:     1 + guests,
        Notes:         notes,
        RecordedBy:    Auth.getUserEmail(),
      });

      _checkedIn.add(memberKey);
      _txns.push({
        TransactionID: txnId, MemberKey: memberKey, MemberName: name,
        EventID: _event.EventID, AmountPaid: amount + guestAmt,
        Category: 'Event', HeadCount: 1 + guests,
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
        <span>${Utils.escape(t.MemberName || t.MemberKey)}</span>
        <span class="amount">${Utils.formatPHP(t.AmountPaid)}</span>
        <span class="fd-mode">${Utils.escape(t.PaymentMode || '')}</span>
      </div>`).join('');
  }

  function changeEvent() { _event = null; _renderEventPicker(); }

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    document.getElementById('checkin-submit-btn')
      ?.addEventListener('click', submitCheckin);
    document.getElementById('checkin-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('checkin-modal'));
    ['checkin-amount','checkin-guests'].forEach(id => {
      document.getElementById(id)?.addEventListener('input', _updateCheckinTotal);
    });
  }

  return { render, init, selectEvent, openCheckin, submitCheckin, changeEvent };
})();
