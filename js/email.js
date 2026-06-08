/**
 * email.js — Email Invite System
 * Select an event, choose recipients (all members / filtered / individual),
 * preview the email, and send via Gmail API.
 */

const Email = (() => {
  let _members = [];
  let _events  = [];
  let _selectedEvent  = null;
  let _selectedMembers = [];  // filtered recipient list
  let _sentCount = 0;

  // ── Render ────────────────────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading…');
    try {
      [_members, _events] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.MEMBERS),
        Sheets.getAll(CONFIG.SHEETS.EVENTS),
      ]);
      _members = _members.filter(m => m['Member Key'] && m['Email']?.includes('@'));
      _renderEventPicker();
      _renderRecipients();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  function _renderEventPicker() {
    const select = document.getElementById('email-event-select');
    if (!select) return;
    const options = _events.map(e =>
      `<option value="${Utils.escape(e.EventID)}">${Utils.escape(e.Title)} — ${Utils.formatDate(e.Date)}</option>`
    );
    select.innerHTML = '<option value="">— Select an event —</option>' + options.join('');
    if (_selectedEvent) select.value = _selectedEvent.EventID;
  }

  function _renderRecipients() {
    const container = document.getElementById('email-recipient-list');
    if (!container) return;

    const query    = document.getElementById('email-search')?.value || '';
    const filter   = document.getElementById('email-filter-status')?.value || '';

    _selectedMembers = _members
      .filter(m => !filter || m['2026 Membership Status (Member, Exempt, Non-member, TBC)'] === filter)
      .filter(m => !query || [m['Full Name'],m['First Name'],m['Email']].some(
        v => (v||'').toLowerCase().includes(query.toLowerCase())));

    document.getElementById('email-recipient-count')
      .textContent = `${_selectedMembers.length} recipients`;

    container.innerHTML = _selectedMembers.length
      ? `<table class="data-table compact">
           <thead><tr><th>Name</th><th>Email</th><th>Status</th></tr></thead>
           <tbody>
             ${_selectedMembers.map(m => `<tr>
               <td>${Utils.escape(m['Full Name'] || m['First Name'])}</td>
               <td>${Utils.escape(m['Email'])}</td>
               <td>${Utils.statusBadge(m['2026 Membership Status (Member, Exempt, Non-member, TBC)'])}</td>
             </tr>`).join('')}
           </tbody>
         </table>`
      : '<p class="empty-state">No matching members with email addresses.</p>';
  }

  // ── Preview ───────────────────────────────────────────────────────────────
  function preview() {
    if (!_selectedEvent) {
      Utils.toast('Please select an event first.', 'error');
      return;
    }
    const sample = _selectedMembers[0];
    if (!sample) { Utils.toast('No recipients selected.', 'error'); return; }

    const html = Gmail.buildInviteEmail(_selectedEvent, sample);
    const frame = document.getElementById('email-preview-frame');
    if (frame) {
      frame.srcdoc = html;
      Utils.showModal('email-preview-modal');
    }
  }

  // ── Send Invites ──────────────────────────────────────────────────────────
  async function sendAll() {
    if (!_selectedEvent) { Utils.toast('Select an event first.', 'error'); return; }
    if (!_selectedMembers.length) { Utils.toast('No recipients.', 'error'); return; }

    const ok = await Utils.confirm(
      `Send invites to ${_selectedMembers.length} member(s) for "${_selectedEvent.Title}"? ` +
      `This will send ${_selectedMembers.length} emails from your Gmail account.`
    );
    if (!ok) return;

    const btn = document.getElementById('email-send-btn');
    btn.disabled = true;
    Utils.setLoading(true, 'Sending emails…');

    _sentCount = 0;
    const failed = [];

    for (const member of _selectedMembers) {
      try {
        const html    = Gmail.buildInviteEmail(_selectedEvent, member);
        const subject = `Invitation: ${_selectedEvent.Title} — ${Utils.formatDate(_selectedEvent.Date)}`;
        await Gmail.send({ to: member['Email'], subject, htmlBody: html });
        _sentCount++;
        // Small delay to avoid hitting Gmail rate limits
        await new Promise(r => setTimeout(r, 200));
      } catch (e) {
        failed.push(`${member['Full Name'] || member['Email']}: ${e.message}`);
      }
    }

    Utils.setLoading(false);
    btn.disabled = false;

    if (failed.length) {
      Utils.toast(`Sent ${_sentCount}, failed ${failed.length}. Check console for details.`, 'error');
      console.error('Failed sends:', failed);
    } else {
      Utils.toast(`✅ ${_sentCount} invites sent successfully!`);
    }
  }

  // ── Open pre-selected to an event (called from Events page) ──────────────
  function openForEvent(eventId) {
    _selectedEvent = null; // will be set after render
    sessionStorage.setItem('email_event', eventId);
    Router.navigate('email');
  }

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    document.getElementById('email-event-select')?.addEventListener('change', e => {
      _selectedEvent = _events.find(ev => ev.EventID === e.target.value) || null;
    });
    document.getElementById('email-search')
      ?.addEventListener('input', Utils.debounce(_renderRecipients));
    document.getElementById('email-filter-status')
      ?.addEventListener('change', _renderRecipients);
    document.getElementById('email-preview-btn')
      ?.addEventListener('click', preview);
    document.getElementById('email-send-btn')
      ?.addEventListener('click', sendAll);
    document.getElementById('email-preview-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('email-preview-modal'));

    // Pre-select event if navigated from Events page
    const preEvent = sessionStorage.getItem('email_event');
    if (preEvent) {
      sessionStorage.removeItem('email_event');
      setTimeout(() => {
        _selectedEvent = _events.find(e => e.EventID === preEvent) || null;
        const sel = document.getElementById('email-event-select');
        if (sel && _selectedEvent) sel.value = _selectedEvent.EventID;
      }, 0);
    }
  }

  return { render, init, preview, sendAll, openForEvent };
})();
