/**
 * events.js — Event Management view
 * Create / edit / delete events; view attendance summaries.
 */

const Events = (() => {
  let _all        = [];
  let _txns       = [];
  let _regs       = [];
  let _editingRow = null;
  let _printData  = null;

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
      const eventRegs  = _regs.filter(r => r.EventID === e.EventID);
      const hasRegs    = eventRegs.length > 0;
      const isPast     = new Date(e.Date) < new Date();
      const safeId     = Utils.escape(e.EventID);

      // ── Core metrics ───────────────────────────────────────────────
      const confirmed   = eventRegs.filter(r => r.PaymentStatus === 'Confirmed').length;
      const pending     = eventRegs.filter(r => r.PaymentStatus === 'Pending').length;
      const activeRegs  = eventRegs.filter(r => r.PaymentStatus !== 'Cancelled');
      const totalPax    = activeRegs.reduce((s, r) =>
        s + (parseInt(r.MemberQty, 10) || 0) + (parseInt(r.GuestQty, 10) || 0) + (parseInt(r.KidsQty, 10) || 0), 0);
      const collected   = eventRegs
        .filter(r => r.PaymentStatus === 'Confirmed')
        .reduce((s, r) => s + Utils.parsePHP(r.AmountPaid || r.TotalDue), 0);
      const totalDue    = activeRegs.reduce((s, r) => s + Utils.parsePHP(r.TotalDue), 0);
      const outstanding = activeRegs
        .filter(r => r.PaymentStatus !== 'Confirmed')
        .reduce((s, r) => s + Utils.parsePHP(r.TotalDue), 0);
      const walkIns     = eventRegs.filter(r => r.WalkIn === 'Yes').length;

      // ── Check-in progress ──────────────────────────────────────────
      const checkedInPax = eventRegs.reduce((s, r) => {
        return s + (r.CheckedIn || '').split(',').map(x => x.trim()).filter(Boolean).length;
      }, 0);
      const checkInPct  = totalPax > 0 ? Math.round(checkedInPax / totalPax * 100) : 0;
      const collectPct  = totalDue  > 0 ? Math.round(collected    / totalDue  * 100) : 0;

      // ── Days label ─────────────────────────────────────────────────
      const diffDays  = Math.round((new Date(e.Date) - new Date()) / 86400000);
      const daysLabel = isPast
        ? (Math.abs(diffDays) === 1 ? 'Yesterday' : `${Math.abs(diffDays)} days ago`)
        : (diffDays === 0 ? 'Today' : diffDays === 1 ? 'Tomorrow' : `In ${diffDays} days`);

      // ── Stat boxes ─────────────────────────────────────────────────
      const statsHTML = hasRegs ? `
        <div class="ec-stats">
          <div class="ec-stat">
            <span class="ec-stat-num">${eventRegs.length}</span>
            <span class="ec-stat-label">Registrations</span>
          </div>
          <div class="ec-stat">
            <span class="ec-stat-num">${totalPax}</span>
            <span class="ec-stat-label">Total Pax</span>
          </div>
          <div class="ec-stat ec-stat-green">
            <span class="ec-stat-num">${confirmed}</span>
            <span class="ec-stat-label">Confirmed</span>
          </div>
          ${pending ? `<div class="ec-stat ec-stat-amber">
            <span class="ec-stat-num">${pending}</span>
            <span class="ec-stat-label">Pending</span>
          </div>` : ''}
          ${walkIns ? `<div class="ec-stat">
            <span class="ec-stat-num">${walkIns}</span>
            <span class="ec-stat-label">Walk-ins</span>
          </div>` : ''}
          <div class="ec-stat-divider"></div>
          <div class="ec-stat">
            <span class="ec-stat-num">${checkedInPax}/${totalPax}</span>
            <span class="ec-stat-label">Checked In</span>
          </div>
          <div class="ec-stat-divider"></div>
          <div class="ec-stat ec-stat-green">
            <span class="ec-stat-num">${Utils.formatPHP(collected)}</span>
            <span class="ec-stat-label">Collected</span>
          </div>
          ${outstanding > 0 ? `<div class="ec-stat ec-stat-amber">
            <span class="ec-stat-num">${Utils.formatPHP(outstanding)}</span>
            <span class="ec-stat-label">Outstanding</span>
          </div>` : ''}
        </div>` : `
        <div class="ec-stats">
          <div class="ec-stat">
            <span class="ec-stat-num">${Utils.formatPHP(e.MemberFee || 0)}</span>
            <span class="ec-stat-label">Member Fee</span>
          </div>
          <div class="ec-stat">
            <span class="ec-stat-num">${Utils.formatPHP(e.GuestFee || 0)}</span>
            <span class="ec-stat-label">Guest Fee</span>
          </div>
          <div class="ec-stat">
            <span class="ec-stat-num">${daysLabel}</span>
            <span class="ec-stat-label">${isPast ? 'Date' : 'Coming up'}</span>
          </div>
        </div>`;

      return `<div class="event-card ${isPast ? 'event-card-past' : 'event-card-upcoming'}">
        <div class="event-card-header">
          <div class="ec-title-block">
            <h3 class="event-card-title">${Utils.escape(e.Title)}</h3>
            <div class="event-card-meta">
              📅 ${Utils.formatDate(e.Date)}
              ${!isPast ? `<span class="ec-days-chip">${daysLabel}</span>` : ''}
              &nbsp;·&nbsp; 📍 ${Utils.escape(e.Location)}
            </div>
          </div>
          <span class="badge badge-${isPast ? 'past' : 'upcoming'}">${isPast ? 'Past' : 'Upcoming'}</span>
        </div>

        ${statsHTML}

        <div class="event-card-actions">
          <button class="btn btn-sm btn-primary" onclick="Events.openRegistrations('${safeId}')">📋 Registrations</button>
          <button class="btn btn-sm btn-secondary" onclick="Events.openFrontDesk('${safeId}')">🎫 Front Desk</button>
          <button class="btn btn-sm btn-outline" onclick="Events.openStats('${safeId}')">📊 Stats</button>
          <button class="btn btn-sm btn-outline" onclick="Events.openEdit('${safeId}')">✏️ Edit</button>
          <button class="btn btn-sm btn-outline" onclick="Email.openForEvent('${safeId}')">📧 Send Invites</button>
          <button class="btn btn-sm btn-danger-outline" onclick="Events.confirmDelete('${safeId}')">🗑️ Delete</button>
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
  async function openStats(eventId) {
    const event = _all.find(e => e.EventID === eventId);
    if (!event) return;
    const regs = _regs.filter(r => r.EventID === eventId);

    document.getElementById('event-stats-title').textContent = event.Title;
    document.getElementById('event-stats-subtitle').textContent =
      `${Utils.formatDate(event.Date)}  ·  ${event.Location}`;

    const members = await Sheets.getAll(CONFIG.SHEETS.MEMBERS).catch(() => []);
    _printData = { event, regs, members };
    document.getElementById('event-stats-body').innerHTML = _buildStatsHtml(event, regs, members);
    Utils.showModal('event-stats-modal');
  }

  function _buildStatsHtml(event, regs, members) {
    if (!regs.length) return '<p class="empty-state">No registrations yet.</p>';

    // ── Pax (exclude cancelled) ───────────────────────────────────────
    const activeRegs = regs.filter(r => r.PaymentStatus !== 'Cancelled');
    const memberPax  = activeRegs.reduce((s, r) => s + (parseInt(r.MemberQty, 10) || 0), 0);
    const guestPax   = activeRegs.reduce((s, r) => s + (parseInt(r.GuestQty,  10) || 0), 0);
    const kidsPax    = activeRegs.reduce((s, r) => s + (parseInt(r.KidsQty,   10) || 0), 0);
    const totalPax   = memberPax + guestPax + kidsPax;

    // ── Payment status ────────────────────────────────────────────────
    const confirmed   = regs.filter(r => r.PaymentStatus === 'Confirmed');
    const pending     = regs.filter(r => r.PaymentStatus === 'Pending');
    const collected   = confirmed.reduce((s, r) => s + Utils.parsePHP(r.AmountPaid || r.TotalDue), 0);
    const outstanding = pending.reduce((s, r) => s + Utils.parsePHP(r.TotalDue), 0);
    const totalDue    = activeRegs.reduce((s, r) => s + Utils.parsePHP(r.TotalDue), 0);
    const collectionRate = totalDue > 0 ? Math.round(collected / totalDue * 100) : 0;

    // ── Pre-paid vs front desk ────────────────────────────────────────
    const prePaidRegs    = confirmed.filter(r => r.WalkIn !== 'Yes');
    const frontDeskRegs  = confirmed.filter(r => r.WalkIn === 'Yes');
    const prePaidTotal   = prePaidRegs.reduce((s, r) => s + Utils.parsePHP(r.AmountPaid || r.TotalDue), 0);
    const frontDeskTotal = frontDeskRegs.reduce((s, r) => s + Utils.parsePHP(r.AmountPaid || r.TotalDue), 0);

    // ── Walk-in pax ───────────────────────────────────────────────────
    const walkInRegs = activeRegs.filter(r => r.WalkIn === 'Yes');
    const walkInPax  = walkInRegs.reduce((s, r) =>
      s + (parseInt(r.MemberQty, 10) || 0) + (parseInt(r.GuestQty, 10) || 0) + (parseInt(r.KidsQty, 10) || 0), 0);
    const preRegPax  = totalPax - walkInPax;

    // ── Payment modes ─────────────────────────────────────────────────
    const modeCounts = {};
    confirmed.forEach(r => {
      const mode = r.PaymentMode || 'Unknown';
      if (!modeCounts[mode]) modeCounts[mode] = { count: 0, amount: 0, walkInCount: 0 };
      modeCounts[mode].count++;
      modeCounts[mode].amount += Utils.parsePHP(r.AmountPaid || r.TotalDue);
      if (r.WalkIn === 'Yes') modeCounts[mode].walkInCount++;
    });
    const modeEntries = Object.entries(modeCounts).sort((a, b) => b[1].amount - a[1].amount);

    // ── Cash reconciliation ───────────────────────────────────────────
    const cashRegs    = confirmed.filter(r => (r.PaymentMode || '').toLowerCase() === 'cash');
    const cashTotal   = cashRegs.reduce((s, r) => s + Utils.parsePHP(r.AmountPaid || r.TotalDue), 0);
    const cashWalkIn  = cashRegs.filter(r => r.WalkIn === 'Yes');
    const cashPreReg  = cashRegs.filter(r => r.WalkIn !== 'Yes');
    const cashWITotal = cashWalkIn.reduce((s, r) => s + Utils.parsePHP(r.AmountPaid || r.TotalDue), 0);
    const cashPRTotal = cashPreReg.reduce((s, r) => s + Utils.parsePHP(r.AmountPaid || r.TotalDue), 0);

    // ── Check-in ──────────────────────────────────────────────────────
    const checkedInSlots = new Set();
    regs.forEach(r => {
      if (r.CheckedIn) r.CheckedIn.split(',').map(s => s.trim()).filter(Boolean)
        .forEach(s => checkedInSlots.add(r.RegistrationID + ':' + s));
    });
    const checkedInPax = checkedInSlots.size;
    const checkInRate  = totalPax > 0 ? Math.round(checkedInPax / totalPax * 100) : 0;

    // ── No-shows (only meaningful once check-in has been used) ────────
    const noShows   = checkedInPax > 0
      ? confirmed.filter(r => !r.CheckedIn || !r.CheckedIn.trim())
      : [];
    const noShowPax = noShows.reduce((s, r) =>
      s + (parseInt(r.MemberQty, 10) || 0) + (parseInt(r.GuestQty, 10) || 0) + (parseInt(r.KidsQty, 10) || 0), 0);

    // ── New members ───────────────────────────────────────────────────
    const eventDate  = (event.Date || '').slice(0, 10);
    const newMembers = (members || []).filter(m => (m['Date Added'] || '').slice(0, 10) === eventDate);

    // ── Helpers ───────────────────────────────────────────────────────
    function bar(value, max, colorClass) {
      const p = max > 0 ? Math.min(100, Math.round(value / max * 100)) : 0;
      return `<div class="stats-bar-track"><div class="stats-bar-fill ${colorClass}" style="width:${p}%"></div></div>`;
    }
    function pct(n, d) { return d > 0 ? Math.round(n / d * 100) : 0; }

    return `
    <div class="stats-grid">

      <!-- 1. Attendance -->
      <div class="stats-card stats-card-wide">
        <div class="stats-card-title">Attendance</div>
        <div class="stats-kpi-row">
          <div class="stats-kpi">
            <span class="stats-kpi-num">${totalPax}</span>
            <span class="stats-kpi-label">Expected Pax</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num stats-color-green">${checkedInPax}</span>
            <span class="stats-kpi-label">Checked In</span>
          </div>
          ${noShows.length ? `
          <div class="stats-kpi">
            <span class="stats-kpi-num stats-color-amber">${noShowPax}</span>
            <span class="stats-kpi-label">Paid, No-show</span>
          </div>` : ''}
          <div class="stats-kpi">
            <span class="stats-kpi-num">${preRegPax}</span>
            <span class="stats-kpi-label">Pre-registered</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num">${walkInPax}</span>
            <span class="stats-kpi-label">Walk-ins</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num stats-color-muted">${regs.length}</span>
            <span class="stats-kpi-label">Registrations</span>
          </div>
        </div>
        ${bar(checkedInPax, totalPax, 'stats-fill-green')}
        <div class="stats-bar-note">${checkInRate}% check-in rate</div>
      </div>

      <!-- 2. Ticket Breakdown -->
      <div class="stats-card stats-card-wide">
        <div class="stats-card-title">Ticket Breakdown</div>
        <div class="stats-kpi-row">
          <div class="stats-kpi">
            <span class="stats-kpi-num">${memberPax}</span>
            <span class="stats-kpi-label">Members</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num">${guestPax}</span>
            <span class="stats-kpi-label">Guests</span>
          </div>
          ${kidsPax > 0 ? `<div class="stats-kpi"><span class="stats-kpi-num">${kidsPax}</span><span class="stats-kpi-label">Kids</span></div>` : ''}
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

      <!-- 3. Revenue -->
      <div class="stats-card stats-card-wide">
        <div class="stats-card-title">Revenue</div>
        <div class="stats-kpi-row">
          <div class="stats-kpi">
            <span class="stats-kpi-num stats-color-green">${Utils.formatPHP(collected)}</span>
            <span class="stats-kpi-label">Total Collected</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num">${Utils.formatPHP(prePaidTotal)}</span>
            <span class="stats-kpi-label">Pre-paid</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num">${Utils.formatPHP(frontDeskTotal)}</span>
            <span class="stats-kpi-label">Front Desk</span>
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
        <div class="stats-bar-note">${collectionRate}% collected · ${confirmed.length} confirmed, ${pending.length} pending</div>
      </div>

      <!-- 4. Payment Modes -->
      <div class="stats-card">
        <div class="stats-card-title">Payment Modes</div>
        ${modeEntries.length ? `<div class="stats-mode-list">
          ${modeEntries.map(([mode, d]) => `
            <div class="stats-mode-row">
              <span class="stats-mode-name">${Utils.escape(mode)}</span>
              <span class="stats-mode-count">${d.count} reg${d.count !== 1 ? 's' : ''}${d.walkInCount ? ` · ${d.walkInCount} walk-in` : ''}</span>
              <span class="stats-mode-amt">${Utils.formatPHP(d.amount)}</span>
            </div>
            ${bar(d.amount, collected, 'stats-fill-blue')}
          `).join('')}
        </div>` : '<p class="stats-empty">No confirmed payments yet.</p>'}
      </div>

      <!-- 5. Cash Reconciliation -->
      <div class="stats-card">
        <div class="stats-card-title">Cash Reconciliation</div>
        ${cashRegs.length ? `
        <div class="stats-kpi-row" style="gap:20px;margin-bottom:14px;">
          <div class="stats-kpi">
            <span class="stats-kpi-num stats-color-green">${Utils.formatPHP(cashTotal)}</span>
            <span class="stats-kpi-label">Total Cash</span>
          </div>
          <div class="stats-kpi">
            <span class="stats-kpi-num">${cashRegs.length}</span>
            <span class="stats-kpi-label">Payers</span>
          </div>
        </div>
        <div class="stats-mode-list">
          ${cashWalkIn.length ? `
          <div class="stats-mode-row">
            <span class="stats-mode-name">Front desk</span>
            <span class="stats-mode-count">${cashWalkIn.length} payer${cashWalkIn.length !== 1 ? 's' : ''}</span>
            <span class="stats-mode-amt">${Utils.formatPHP(cashWITotal)}</span>
          </div>` : ''}
          ${cashPreReg.length ? `
          <div class="stats-mode-row">
            <span class="stats-mode-name">Pre-registered</span>
            <span class="stats-mode-count">${cashPreReg.length} payer${cashPreReg.length !== 1 ? 's' : ''}</span>
            <span class="stats-mode-amt">${Utils.formatPHP(cashPRTotal)}</span>
          </div>` : ''}
        </div>` : '<p class="stats-empty">No cash payments recorded.</p>'}
      </div>

      <!-- 6. Paid no-shows -->
      ${noShows.length ? `
      <div class="stats-card stats-card-wide">
        <div class="stats-card-title">Paid · Did Not Show Up (${noShows.length} reg${noShows.length !== 1 ? 's' : ''}, ${noShowPax} pax)</div>
        <div class="stats-mode-list">
          ${noShows.map(r => `
            <div class="stats-mode-row">
              <span class="stats-mode-name">${Utils.escape(r.LastName || '')}${r.FirstName ? ', ' + Utils.escape(r.FirstName) : ''}</span>
              <span class="stats-mode-count">${Utils.escape(r.PaymentMode || '')} · ${(parseInt(r.MemberQty, 10) || 0) + (parseInt(r.GuestQty, 10) || 0) + (parseInt(r.KidsQty, 10) || 0)} pax</span>
              <span class="stats-mode-amt">${Utils.formatPHP(r.AmountPaid || r.TotalDue)}</span>
            </div>
          `).join('')}
        </div>
      </div>` : ''}

      <!-- 7. New members -->
      ${newMembers.length ? `
      <div class="stats-card stats-card-wide">
        <div class="stats-card-title">New Members Added (${newMembers.length})</div>
        <div class="stats-mode-list">
          ${newMembers.map(m => `
            <div class="stats-mode-row">
              <span class="stats-mode-name">${Utils.escape(m['First Name'] || '')} ${Utils.escape(m['Last Name'] || '')}</span>
              <span class="stats-mode-count">${Utils.escape(m['Member Key'] || '')}</span>
              <span class="stats-mode-amt">${Utils.escape(m['Membership Type'] || '')}</span>
            </div>
          `).join('')}
        </div>
      </div>` : ''}

    </div>`;
  }

  function printStats() {
    if (!_printData) return;
    const { event, regs, members } = _printData;
    const title    = document.getElementById('event-stats-title').textContent;
    const subtitle = document.getElementById('event-stats-subtitle').textContent;
    const body     = _buildStatsHtml(event, regs, members);
    const generated = new Date().toLocaleDateString('en-PH', { dateStyle: 'long' });

    const win = window.open('', '_blank');
    win.document.write(`<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>${Utils.escape(title)} — Event Report</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:14px;color:#111;background:#fff;padding:32px}
h1{font-size:22px;font-weight:700;margin-bottom:4px}
.subtitle{font-size:14px;color:#555;margin-bottom:4px}
.generated{font-size:11px;color:#999;margin-bottom:24px}
.stats-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.stats-card{border:1px solid #e5e7eb;border-radius:8px;padding:16px}
.stats-card-wide{grid-column:1/-1}
.stats-card-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#6b7280;margin-bottom:12px}
.stats-kpi-row{display:flex;gap:24px;flex-wrap:wrap;margin-bottom:12px}
.stats-kpi{display:flex;flex-direction:column;gap:2px}
.stats-kpi-num{font-size:26px;font-weight:700;line-height:1;font-variant-numeric:tabular-nums}
.stats-kpi-label{font-size:11px;color:#6b7280}
.stats-split-row{display:flex;align-items:center;gap:12px;margin-bottom:10px}
.stats-split-item{display:flex;flex-direction:column;gap:2px;flex:1}
.stats-split-num{font-size:22px;font-weight:700;font-variant-numeric:tabular-nums}
.stats-split-label{font-size:11px;color:#6b7280}
.stats-split-div{width:1px;height:36px;background:#e5e7eb;flex-shrink:0}
.stats-bar-stacked{display:flex;height:10px;border-radius:5px;overflow:hidden;background:#f3f4f6;margin-bottom:6px}
.stats-bar-seg{height:100%}
.stats-bar-track{height:8px;border-radius:4px;background:#f3f4f6;overflow:hidden;margin-bottom:6px}
.stats-bar-fill{height:100%;border-radius:4px}
.stats-bar-note{font-size:11px;color:#6b7280}
.stats-bar-legend{font-size:11px;color:#6b7280;display:flex;align-items:center;margin-bottom:6px}
.stats-bar-dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:4px}
.stats-dot-member,.stats-bar-seg.stats-dot-member{background:#2563eb}
.stats-dot-guest,.stats-bar-seg.stats-dot-guest{background:#7c3aed}
.stats-dot-kids,.stats-bar-seg.stats-dot-kids{background:#0891b2}
.stats-fill-green{background:#16a34a}
.stats-fill-blue{background:#2563eb}
.stats-fill-amber{background:#d97706}
.stats-color-green{color:#16a34a}
.stats-color-amber{color:#d97706}
.stats-color-muted{color:#6b7280}
.stats-mode-list{display:flex;flex-direction:column;gap:8px}
.stats-mode-row{display:flex;align-items:center;gap:6px;font-size:13px}
.stats-mode-name{flex:1;font-weight:500}
.stats-mode-count{font-size:11px;color:#6b7280;white-space:nowrap}
.stats-mode-amt{font-variant-numeric:tabular-nums;font-weight:600;white-space:nowrap;min-width:80px;text-align:right}
.stats-empty{font-size:13px;color:#6b7280;margin:0}
@media print{body{padding:16px}@page{margin:16mm}}
</style>
</head>
<body>
<h1>${Utils.escape(title)}</h1>
<div class="subtitle">${Utils.escape(subtitle)}</div>
<div class="generated">Report generated ${Utils.escape(generated)}</div>
${body}
<script>window.onload=function(){window.print();}<\/script>
</body>
</html>`);
    win.document.close();
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

  return { render, init, openAdd, openEdit, save, confirmDelete, viewAttendees, openFrontDesk, openRegistrations, openStats, printStats, getAll };
})();
