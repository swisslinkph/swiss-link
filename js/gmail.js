/**
 * gmail.js — Gmail API helpers (send invite emails)
 * Uses the logged-in admin's Gmail account as sender.
 */

const Gmail = (() => {
  const API = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/send';

  function _makeEmail({ to, subject, htmlBody }) {
    const from    = Auth.getUserEmail();
    const headers = [
      `From: ${Auth.getUserName()} <${from}>`,
      `To: ${to}`,
      `Subject: ${subject}`,
      'MIME-Version: 1.0',
      'Content-Type: text/html; charset=UTF-8',
    ].join('\r\n');
    const raw = `${headers}\r\n\r\n${htmlBody}`;
    return btoa(unescape(encodeURIComponent(raw)))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  }

  async function send({ to, subject, htmlBody }) {
    const raw  = _makeEmail({ to, subject, htmlBody });
    const token = Auth.getToken();
    const res   = await fetch(API, {
      method: 'POST',
      headers: {
        Authorization:  `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ raw }),
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error?.message || 'Gmail send failed');
    }
    return res.json();
  }

  /**
   * Build an event invite email for a single member.
   * Includes a personalised RSVP link.
   */
  function buildInviteEmail(event, member) {
    const rsvpToken = Utils.rsvpToken(event.EventID, member['Member Key'] || member.FullName);
    const rsvpUrl   = `${CONFIG.APP_BASE_URL}/rsvp.html?event=${encodeURIComponent(event.EventID)}&member=${encodeURIComponent(member['Member Key'] || '')}&token=${rsvpToken}`;

    const memberFee   = Utils.formatPHP(event.MemberFee);
    const guestFee    = Utils.formatPHP(event.GuestFee);
    const eventDate   = Utils.formatDate(event.Date);

    return `<!DOCTYPE html>
<html><body style="font-family:Arial,sans-serif;color:#2d2d2d;max-width:600px;margin:0 auto;padding:20px">
  <div style="background:#CC0000;padding:20px;border-radius:8px 8px 0 0;text-align:center">
    <img src="${CONFIG.APP_BASE_URL}/assets/logo.png" alt="Swiss Club" style="height:60px" onerror="this.style.display='none'">
    <h1 style="color:#fff;margin:10px 0 0;font-size:22px">Swiss Club of the Philippines</h1>
  </div>
  <div style="background:#fff;border:1px solid #ddd;border-top:none;padding:30px;border-radius:0 0 8px 8px">
    <p>Dear ${Utils.escape(member.FirstName || member.FullName)},</p>
    <p>You are cordially invited to <strong>${Utils.escape(event.Title)}</strong>.</p>

    <table style="width:100%;border-collapse:collapse;margin:20px 0;background:#f9f9f9;border-radius:6px">
      <tr><td style="padding:10px 15px;font-weight:bold;width:140px">📅 Date</td>
          <td style="padding:10px 15px">${eventDate}</td></tr>
      <tr style="background:#f0f0f0"><td style="padding:10px 15px;font-weight:bold">📍 Location</td>
          <td style="padding:10px 15px">${Utils.escape(event.Location)}</td></tr>
      <tr><td style="padding:10px 15px;font-weight:bold">💳 Member fee</td>
          <td style="padding:10px 15px">${memberFee}</td></tr>
      ${event.GuestFee ? `<tr style="background:#f0f0f0"><td style="padding:10px 15px;font-weight:bold">👥 Guest fee</td>
          <td style="padding:10px 15px">${guestFee}</td></tr>` : ''}
      ${event.Description ? `<tr><td style="padding:10px 15px;font-weight:bold">ℹ️ Info</td>
          <td style="padding:10px 15px">${Utils.escape(event.Description)}</td></tr>` : ''}
    </table>

    <div style="text-align:center;margin:30px 0">
      <a href="${rsvpUrl}" style="background:#CC0000;color:#fff;padding:14px 40px;border-radius:6px;text-decoration:none;font-size:16px;font-weight:bold;display:inline-block">
        ✅ RSVP Now
      </a>
    </div>

    <p style="font-size:13px;color:#888">
      If the button does not work, copy this link into your browser:<br>
      <a href="${rsvpUrl}" style="color:#CC0000;word-break:break-all">${rsvpUrl}</a>
    </p>
    <p style="font-size:13px;color:#888">This invitation was sent by ${Utils.escape(Auth.getUserName())} via the Swiss Club Admin App.</p>
  </div>
</body></html>`;
  }

  return { send, buildInviteEmail };
})();
