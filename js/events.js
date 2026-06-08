/**
 * events.js — Event Management view
 * Create / edit / delete events; view attendance summaries.
 */

const Events = (() => {
  let _all      = [];
  let _txns     = [];
  let _editingRow = null;

  // ── Render ────────────────────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading events…');
    try {
      [_all, _txns] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.EVENTS),
        Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS),
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
      const attendeeTxns = _txns.filter(t => t.EventID === e.EventID && t.Category === 'Event');
      const attendeeCount = attendeeTxns.length;
      const revenue       = attendeeTxns.reduce((s, t) => s + Utils.parsePHP(t.AmountPaid), 0);
      const isPast        = new Date(e.Date) < new Date() && e.Status !== 'Upcoming';
      const statusClass   = isPast ? 'event-card-past' : 'event-card-upcoming';
      const rsvpCount     = _txns.filter(t => t.EventID === e.EventID && t.Category === 'RSVP').length;

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

        <div class="event-card-stats">
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
          </div>
        </div>

        ${e.Description ? `<p class="event-card-desc">${Utils.escape(e.Description)}</p>` : ''}

        <div class="event-card-actions">
          <button class="btn btn-sm" onclick="Events.openFrontDesk('${Utils.escape(e.EventID)}')">
            🎫 Front Desk
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
    set('ef-title',       event.Title);
    set('ef-date',        Utils.toISODate(event.Date));
    set('ef-location',    event.Location);
    set('ef-desc',        event.Description);
    set('ef-member-fee',  event.MemberFee);
    set('ef-guest-fee',   event.GuestFee);
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
        Title:       title,
        Date:        get('ef-date'),
        Location:    get('ef-location'),
        Description: get('ef-desc'),
        MemberFee:   get('ef-member-fee'),
        GuestFee:    get('ef-guest-fee'),
        Status:      new Date(get('ef-date')) >= new Date() ? 'Upcoming' : 'Completed',
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

  return { render, init, openAdd, openEdit, save, confirmDelete, viewAttendees, openFrontDesk, getAll };
})();
