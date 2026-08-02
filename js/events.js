/**
 * events.js — Event Management view
 * Create / edit / delete events; view attendance summaries.
 */

const Events = (() => {
  let _all      = [];
  let _txns     = [];
  let _regs     = [];
  let _editingRow = null;

  // ── Render ────────────────────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading events…');
    try {
      [_all, _txns, _regs] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.EVENTS),
        Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS),
        Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).catch(() => []),
      ]);
      _renderList();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  function _renderList() {
    const container = document.getElementById('events-list');
    if (!container) return;

    const sorted = [..._all].sort((a, b) => new Date(b.Date) - new Date(a.Date));

    if (!sorted.length) {
      container.innerHTML = '<p class="empty-state">No events yet. Create your first event!</p>';
      return;
    }

    container.innerHTML = sorted.map(e => {
      const attendeeTxns  = _txns.filter(t => t.EventID === e.EventID && t.Category === 'Event');
      const rsvpCount     = _txns.filter(t => t.EventID === e.EventID && t.Category === 'RSVP').length;
      const eventRegs     = _regs.filter(r => r.EventID === e.EventID);
      const hasRegs       = eventRegs.length > 0;
      const isPast        = new Date(e.Date) < new Date();
      const statusClass   = isPast ? 'event-card-past' : 'event-card-upcoming';

      let statsHTML;
      if (hasRegs) {
        const confirmed  = eventRegs.filter(r => r.PaymentStatus === 'Confirmed').length;
        const pending    = eventRegs.filter(r => r.PaymentStatus === 'Pending').length;
        const totalPax   = eventRegs.reduce((s, r) =>
          s + (parseInt(r.MemberQty, 10) || 0) + (parseInt(r.GuestQty, 10) || 0) + (parseInt(r.KidsQty, 10) || 0), 0);
        const collected  = eventRegs
          .filter(r => r.PaymentStatus === 'Confirmed')
          .reduce((s, r) => s + Utils.parsePHP(r.AmountPaid || r.TotalDue), 0);
        statsHTML = `
          <div class="stat-box">
            <span class="stat-num">${eventRegs.length}</span>
            <span class="stat-label">Registrations</span>
          </div>
          <div class="stat-box">
            <span class="stat-num">${totalPax}</span>
            <span class="stat-label">Total Pax</span>
          </div>
          <div class="stat-box" style="color:#166534">
            <span class="stat-num">${confirmed}</span>
            <span class="stat-label">Confirmed</span>
          </div>
          <div class="stat-box" style="color:#92400e">
            <span class="stat-num">${pending}</span>
            <span class="stat-label">Pending</span>
          </div>
          <div class="stat-box">
            <span class="stat-num">${Utils.formatPHP(collected)}</span>
            <span class="stat-label">Collected</span>
          </div>`;
      } else {
        const attendeeCount = attendeeTxns.length;
        const revenue       = attendeeTxns.reduce((s, t) => s + Utils.parsePHP(t.AmountPaid), 0);
        statsHTML = `
          <div class="stat-box">
            <span class="stat-num">${attendeeCount}</span>
            <span class="stat-label">Attendees</span>
          </div>
          <div class="stat-box">
            <span class="stat-num">${rsvpCount}</span>
            <span class="stat-label">RSVPs</span>
          </div>
          <div class="stat-box">
            <span class="stat-num">${Utils.formatPHP(revenue)}</span>
            <span class="stat-label">Revenue</span>
          </div>
          <div class="stat-box">
            <span class="stat-num">${Utils.formatPHP(e.MemberFee)}</span>
            <span class="stat-label">Member Fee</span>
          </div>`;
      }

      return `<div class="event-card ${statusClass}">
        <div class="event-card-header">
          <div>
            <h3 class="event-card-title">${Utils.escape(e.Title)}</h3>
            <div class="event-card-meta">
              📅 ${Utils.formatDate(e.Date)} &nbsp;|&nbsp; 📍 ${Utils.escape(e.Location)}
            </div>
          </div>
          <div class="event-card-status">
            <span class="badge badge-${isPast ? 'past' : 'upcoming'}">${isPast ? 'Past' : 'Upcoming'}</span>
          </div>
        </div>

        <div class="event-card-stats">${statsHTML}</div>

        ${e.Description ? `<p class="event-card-desc">${Utils.escape(e.Description)}</p>` : ''}

        <div class="event-card-actions">
          <button class="btn btn-sm btn-primary" onclick="Events.openRegistrations('${Utils.escape(e.EventID)}')">
            📋 Registrations
          </button>
          <button class="btn btn-sm" onclick="Events.openFrontDesk('${Utils.escape(e.EventID)}')">
            🎫 Front Desk
          </button>
          <button class="btn btn-sm btn-outline" onclick="Events.openStats('${Utils.escape(e.EventID)}')">
            📊 Stats
          </button>
          <button class="btn btn-sm btn-outline" onclick="Events.openEdit('${Utils.escape(e.EventID)}')">
            ✏️ Edit
          </button>
          <button class="btn btn-sm btn-outline" onclick="Email.openForEvent('${Utils.escape(e.EventID)}')">
            📧 Send Invites
          </button>
          <button class="btn btn-sm btn-outline" onclick="Events.viewAttendees('${Utils.escape(e.EventID)}')">
            👥 Attendees
          </button>
          <button class="btn btn-sm btn-danger-outline" onclick="Events.confirmDelete('${Utils.escape(e.EventID)}')">
            🗑️ Delete
          </button>
        </div>
      </div>`;
    }).join('');
  }

  // ── Open Add modal ────────────────────────────────────────────────────────
  function openAdd() {
    _editingRow = null;
    document.getElementById('event-modal-title').textContent = 'New Event';
    document.getElementById('event-form').reset();
    document.getElementById('ef-date').value = Utils.today();
    Utils.showModal('event-modal');
  }

  // ── Open Edit modal ───────────────────────────────────────────────────────
  function openEdit(eventId) {
    const event = _all.find(e => e.EventID === eventId);
    if (!event) return;
    _editingRow = event._rowIndex;
    document.getElementById('event-modal-title').textContent = 'Edit Event';
    const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
    set('ef-title',          event.Title);
    set('ef-date',           Utils.toISODate(event.Date));
    set('ef-location',       event.Location);
    set('ef-desc',           event.Description);
    set('ef-member-fee',     event.MemberFee);
    set('ef-guest-fee',      event.GuestFee);
    set('ef-kids-fee',       event.KidsFee);
    set('ef-walkin-member',   event.WalkInMemberFee);
    set('ef-walkin-guest',    event.WalkInGuestFee);
    set('ef-form-sheet-id',   event.FormSheetID);
    set('ef-form-sheet-tab',  event.FormSheetTab);
    Utils.showModal('event-modal');
  }

  // ── Save ──────────────────────────────────────────────────────────────────
  async function save() {
    const btn = document.getElementById('event-save-btn');
    btn.disabled = true;
    try {
      const get  = id => document.getElementById(id)?.value?.trim() || '';
      const title = get('ef-title');
      if (!title) { Utils.toast('Event title is required.', 'error'); btn.disabled = false; return; }

      const obj = {
        Title:            title,
        Date:             get('ef-date'),
        Location:         get('ef-location'),
        Description:      get('ef-desc'),
        MemberFee:        get('ef-member-fee'),
        GuestFee:         get('ef-guest-fee'),
        KidsFee:          get('ef-kids-fee'),
        WalkInMemberFee:  get('ef-walkin-member'),
        WalkInGuestFee:   get('ef-walkin-guest'),
        FormSheetID:      get('ef-form-sheet-id'),
        FormSheetTab:     get('ef-form-sheet-tab'),
        Status:           new Date(get('ef-date')) >= new Date() ? 'Upcoming' : 'Completed',
      };

      if (_editingRow) {
        const existing = _all.find(e => e._rowIndex === _editingRow);
        await Sheets.update(CONFIG.SHEETS.EVENTS, _editingRow, { ...existing, ...obj });
        Utils.toast('Event updated.');
      } else {
        obj.EventID      = await Sheets.nextId(CONFIG.SHEETS.EVENTS, 'EVT');
        obj.CreatedDate  = new Date().toISOString();
        obj.CreatedBy    = Auth.getUserEmail();
        await Sheets.append(CONFIG.SHEETS.EVENTS, obj);
        Utils.toast('Event created.');
      }

      Utils.hideModal('event-modal');
      await render();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  // ── Delete ────────────────────────────────────────────────────────────────
  async function confirmDelete(eventId) {
    const event = _all.find(e => e.EventID === eventId);
    if (!event) return;
    const ok = await Utils.confirm(`Delete "${event.Title}"? All transaction records for this event will remain.`);
    if (!ok) return;
    Utils.setLoading(true, 'Deleting…');
    try {
      await Sheets.deleteRow(CONFIG.SHEETS.EVENTS, event._rowIndex);
      Utils.toast('Event deleted.');
      await render();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  // ── View attendees ────────────────────────────────────────────────────────
  function viewAttendees(eventId) {
    const event = _all.find(e => e.EventID === eventId);
    const attendees = _txns.filter(t => t.EventID === eventId && t.Category === 'Event');

    const container = document.getElementById('attendee-list');
    const title     = document.getElementById('attendee-modal-title');
    if (!container || !title) return;

    title.textContent = `Attendees — ${event?.Title || eventId}`;

    if (!attendees.length) {
      container.innerHTML = '<p class="empty-state">No attendees recorded yet.</p>';
    } else {
      const total = attendees.reduce((s, t) => s + Utils.parsePHP(t.AmountPaid), 0);
      container.innerHTML = `
        <p class="attendee-summary">${attendees.length} attendees &nbsp;|&nbsp; Total: ${Utils.formatPHP(total)}</p>
        <table class="data-table compact">
          <thead><tr><th>Member</th><th>Amount</th><th>Mode</th><th>Guests</th><th>Notes</th></tr></thead>
          <tbody>
            ${attendees.map(t => `<tr>
              <td>${Utils.escape(t.MemberName || t.MemberKey)}</td>
              <td class="amount">${Utils.formatPHP(t.AmountPaid)}</td>
              <td>${Utils.escape(t.PaymentMode)}</td>
              <td>${Utils.escape(t.HeadCount)}</td>
              <td>${Utils.escape(t.Notes)}</td>
            </tr>`).join('')}
          </tbody>
        </table>`;
    }
    Utils.showModal('attendee-modal');
  }

  // ── Navigate to front desk for this event ────────────────────────────────
  function openFrontDesk(eventId) {
    sessionStorage.setItem('fd_event', eventId);
    Router.navigate('frontdesk');
  }

  // ── Navigate to registrations view for this event ─────────────────────────
  function openRegistrations(eventId) {
    sessionStorage.setItem('reg_event', eventId);
    Router.navigate('registrations');
  }

  // ── Event Stats modal ─────────────────────────────────────────────────────
  function openStats(eventId) {
    const event = _all.find(e => e.EventID === eventId);
    if (!event) return;
    const regs = _regs.filter(r => r.EventID === eventId);

    document.getElementById('event-stats-title').textContent = event.Title;
    document.getElementById('event-stats-subtitle').textContent =
      `${Utils.formatDate(event.Date)}  ·  ${event.Location}`;

    document.getElementById('event-stats-body').innerHTML = _buildStatsHtml(event, regs);
    Utils.showModal('event-stats-modal');
  }

  function _buildStatsHtml(event, regs) {
    if (!regs.length) return '<p class="empty-state">No registrations yet.</p>';

    // ── Ticket counts ─────────────────────────────────────────────────
    const memberPax = regs.reduce((s, r) => s + (parseInt(r.MemberQty,  10) || 0), 0);
    const guestPax  = regs.reduce((s, r) => s + (parseInt(r.GuestQty,   10) || 0), 0);
    const kidsPax   = regs.reduce((s, r) => s + (parseInt(r.KidsQty,    10) || 0), 0);
    const totalPax  = memberPax + guestPax + kidsPax;

    // ── Payment status ────────────────────────────────────────────────
    const confirmed  = regs.filter(r => r.PaymentStatus === 'Confirmed');
    const pending    = regs.filter(r => r.PaymentStatus !== 'Confirmed');
    const collected  = confirmed.reduce((s, r) => s + Utils.parsePHP(r.AmountPaid || r.TotalDue), 0);
    const outstanding = pending.reduce((s, r) => s + Utils.parsePHP(r.TotalDue), 0);
    const totalDue   = regs.reduce((s, r) => s + Utils.parsePHP(r.TotalDue), 0);
    const collectionRate = totalDue > 0 ? Math.round(collected / totalDue * 100) : 0;

    // ── Payment modes ─────────────────────────────────────────────────
    const modeCounts = {};
    confirmed.forEach(r => {
      const mode = r.PaymentMode || 'Unknown';
      if (!modeCounts[mode]) modeCounts[mode] = { count: 0, amount: 0 };
      modeCounts[mode].count++;
      modeCounts[mode].amount += Utils.parsePHP(r.AmountPaid || r.TotalDue);
    });
    const modeEntries = Object.entries(modeCounts).sort((a, b) => b[1].amount - a[1].amount);

    // ── Check-in progress ─────────────────────────────────────────────
    const checkedInSlots = new Set();
    regs.forEach(r => {
      if (r.CheckedIn) r.CheckedIn.split(',').map(s => s.trim()).filter(Boolean).forEach(s => checkedInSlots.add(r.RegistrationID + ':' + s));
    });
    const checkedInPax = checkedInSlots.size;
    const checkInRate  = totalPax > 0 ? Math.round(checkedInPax / totalPax * 100) : 0;

    // Walk-in registrations
    const walkInRegs   = regs.filter(r => r.WalkIn === 'Yes' || (r.Category || '').toLowerCase().includes('walk'));
    const walkInPax    = walkInRegs.reduce((s, r) =>
      s + (parseInt(r.MemberQty, 10) || 0) + (parseInt(r.GuestQty, 10) || 0) + (parseInt(r.KidsQty, 10) || 0), 0);
    const preRegPax    = totalPax - walkInPax;

    // ── Helpers ───────────────────────────────────────────────────────
    function bar(value, max, colorClass) {
      const pct = max > 0 ? Math.min(100, Math.round(value / max * 100)) : 0;
      return `<div class="stats-bar-track"><div class="stats-bar-fill ${colorClass}" style="width:${pct}%"></div></div>`;
    }
    function pct(n, d) { return d > 0 ? Math.round(n / d * 100) : 0; }

    return `
    <div class="stats-grid">

      <!-- Pax Breakdown -->
      <div class="stats-card stats-card-wide">
        <div class="stats-card-title">Ticket Breakdown</div>
        <div class="stats-kpi-row">
          <div class="stats-kpi">
            <span class="stats-kpi-num">${totalPax}</span>
            <span class="stats-kpi-label">Total Pax</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num">${regs.length}</span>
            <span class="stats-kpi-label">Registrations</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num">${memberPax}</span>
            <span class="stats-kpi-label">Member Tickets</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num">${guestPax}</span>
            <span class="stats-kpi-label">Guest Tickets</span>
          </div>
          ${kidsPax > 0 ? `<div class="stats-kpi"><span class="stats-kpi-num">${kidsPax}</span><span class="stats-kpi-label">Kids Tickets</span></div>` : ''}
        </div>
        <div class="stats-bar-legend">
          <span class="stats-bar-dot stats-dot-member"></span>Members ${pct(memberPax, totalPax)}%
          <span class="stats-bar-dot stats-dot-guest" style="margin-left:12px;"></span>Guests ${pct(guestPax, totalPax)}%
          ${kidsPax > 0 ? `<span class="stats-bar-dot stats-dot-kids" style="margin-left:12px;"></span>Kids ${pct(kidsPax, totalPax)}%` : ''}
        </div>
        <div class="stats-bar-stacked">
          <div class="stats-bar-seg stats-dot-member" style="width:${pct(memberPax, totalPax)}%" title="Members: ${memberPax}"></div>
          <div class="stats-bar-seg stats-dot-guest"  style="width:${pct(guestPax,  totalPax)}%" title="Guests: ${guestPax}"></div>
          ${kidsPax > 0 ? `<div class="stats-bar-seg stats-dot-kids" style="width:${pct(kidsPax, totalPax)}%" title="Kids: ${kidsPax}"></div>` : ''}
        </div>
      </div>

      <!-- Walk-in vs Pre-registered -->
      <div class="stats-card">
        <div class="stats-card-title">Registration Type</div>
        <div class="stats-split-row">
          <div class="stats-split-item">
            <span class="stats-split-num">${preRegPax}</span>
            <span class="stats-split-label">Pre-registered</span>
          </div>
          <div class="stats-split-div"></div>
          <div class="stats-split-item">
            <span class="stats-split-num">${walkInPax}</span>
            <span class="stats-split-label">Walk-ins</span>
          </div>
        </div>
        ${bar(preRegPax, totalPax, 'stats-fill-blue')}
        <div class="stats-bar-note">${pct(preRegPax, totalPax)}% pre-registered</div>
      </div>

      <!-- Check-in Progress -->
      <div class="stats-card">
        <div class="stats-card-title">Check-in Progress</div>
        <div class="stats-split-row">
          <div class="stats-split-item">
            <span class="stats-split-num stats-color-green">${checkedInPax}</span>
            <span class="stats-split-label">Checked In</span>
          </div>
          <div class="stats-split-div"></div>
          <div class="stats-split-item">
            <span class="stats-split-num stats-color-muted">${totalPax - checkedInPax}</span>
            <span class="stats-split-label">Not Yet Arrived</span>
          </div>
        </div>
        ${bar(checkedInPax, totalPax, 'stats-fill-green')}
        <div class="stats-bar-note">${checkInRate}% checked in</div>
      </div>

      <!-- Payment Status -->
      <div class="stats-card">
        <div class="stats-card-title">Payment Status</div>
        <div class="stats-split-row">
          <div class="stats-split-item">
            <span class="stats-split-num stats-color-green">${confirmed.length}</span>
            <span class="stats-split-label">Confirmed</span>
          </div>
          <div class="stats-split-div"></div>
          <div class="stats-split-item">
            <span class="stats-split-num stats-color-amber">${pending.length}</span>
            <span class="stats-split-label">Pending</span>
          </div>
        </div>
        ${bar(confirmed.length, regs.length, 'stats-fill-green')}
        <div class="stats-bar-note">${pct(confirmed.length, regs.length)}% of registrations confirmed</div>
      </div>

      <!-- Collection -->
      <div class="stats-card">
        <div class="stats-card-title">Collection</div>
        <div class="stats-kpi-row" style="gap:16px;">
          <div class="stats-kpi">
            <span class="stats-kpi-num stats-color-green">${Utils.formatPHP(collected)}</span>
            <span class="stats-kpi-label">Collected</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num stats-color-amber">${Utils.formatPHP(outstanding)}</span>
            <span class="stats-kpi-label">Outstanding</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num stats-color-muted">${Utils.formatPHP(totalDue)}</span>
            <span class="stats-kpi-label">Total Due</span>
          </div>
        </div>
        ${bar(collected, totalDue, 'stats-fill-green')}
        <div class="stats-bar-note">${collectionRate}% collected</div>
      </div>

      <!-- Payment Modes -->
      <div class="stats-card">
        <div class="stats-card-title">Payment Modes</div>
        ${modeEntries.length ? `<div class="stats-mode-list">
          ${modeEntries.map(([mode, d]) => `
            <div class="stats-mode-row">
              <span class="stats-mode-name">${Utils.escape(mode)}</span>
              <span class="stats-mode-count">${d.count} reg${d.count !== 1 ? 's' : ''}</span>
              <span class="stats-mode-amt">${Utils.formatPHP(d.amount)}</span>
            </div>
            ${bar(d.amount, collected, 'stats-fill-blue')}
          `).join('')}
        </div>` : '<p class="stats-empty">No confirmed payments yet.</p>'}
      </div>

    </div>`;
  }

  // ── Get all events (for Email picker) ────────────────────────────────────
  function getAll() { return _all; }

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    document.getElementById('add-event-btn')
      ?.addEventListener('click', openAdd);
    document.getElementById('event-save-btn')
      ?.addEventListener('click', save);
    document.getElementById('event-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('event-modal'));
    document.getElementById('attendee-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('attendee-modal'));
  }

  return { render, init, openAdd, openEdit, save, confirmDelete, viewAttendees, openFrontDesk, openRegistrations, openStats, getAll };
})();
