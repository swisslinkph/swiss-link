/**
 * gmail.js — Gmail API helpers (send invite emails)
 * Uses the logged-in admin's Gmail account as sender.
 */

const Gmail = (() => {
  const API = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/send';

  // Encode non-ASCII subject lines per RFC 2047 (required by email spec)
  function _encodeSubject(text) {
    if (/^[\x00-\x7F]*$/.test(text)) return text;
    return `=?UTF-8?B?${btoa(unescape(encodeURIComponent(text)))}?=`;
  }

  function _makeEmail({ to, subject, htmlBody }) {
    const from    = Auth.getUserEmail();
    const headers = [
      `From: Swiss Club of the Philippines <${from}>`,
      `To: ${to}`,
      `Subject: ${_encodeSubject(subject)}`,
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
    const key        = member['Member Key'] || '';
    const firstName  = member['First Name'] || member['Last Name'] || '';
    const rsvpToken  = Utils.rsvpToken(event.EventID, key);
    const baseParams = `event=${encodeURIComponent(event.EventID)}&member=${encodeURIComponent(key)}&token=${rsvpToken}&name=${encodeURIComponent(firstName)}`;
    const rsvpUrl    = `${CONFIG.APP_BASE_URL}/rsvp.html?${baseParams}`;
    const rsvpYes    = `${rsvpUrl}&response=yes`;
    const rsvpNo     = `${rsvpUrl}&response=no`;

    const memberFee = Utils.formatPHP(event.MemberFee);
    const guestFee  = Utils.formatPHP(event.GuestFee);
    const eventDate = Utils.formatDate(event.Date);

    return `<!DOCTYPE html>
<html><body style="font-family:Arial,sans-serif;color:#2d2d2d;max-width:600px;margin:0 auto;padding:20px;background:#f4f4f4">
  <div style="background:#CC0000;padding:24px 20px;border-radius:8px 8px 0 0;text-align:center">
    <div style="font-size:36px;margin-bottom:8px">🇨🇭</div>
    <h1 style="color:#fff;margin:0;font-size:20px;font-weight:700;letter-spacing:0.3px">Swiss Club of the Philippines</h1>
  </div>
  <div style="background:#fff;border:1px solid #ddd;border-top:none;padding:32px;border-radius:0 0 8px 8px">
    <p style="font-size:16px;margin:0 0 8px">Dear <strong>${Utils.escape(firstName)}</strong>,</p>
    <p style="color:#555;margin:0 0 24px">You are cordially invited to the following event:</p>

    <div style="background:#f8f8f8;border-left:4px solid #CC0000;border-radius:0 6px 6px 0;padding:16px 20px;margin-bottom:24px">
      <div style="font-size:18px;font-weight:700;margin-bottom:12px">${Utils.escape(event.Title)}</div>
      <table style="border-collapse:collapse;font-size:14px;color:#444">
        <tr><td style="padding:3px 12px 3px 0;white-space:nowrap">📅 Date</td><td style="padding:3px 0"><strong>${eventDate}</strong></td></tr>
        <tr><td style="padding:3px 12px 3px 0;white-space:nowrap">📍 Location</td><td style="padding:3px 0">${Utils.escape(event.Location)}</td></tr>
        <tr><td style="padding:3px 12px 3px 0;white-space:nowrap">💰 Member fee</td><td style="padding:3px 0">${memberFee}</td></tr>
        ${event.GuestFee ? `<tr><td style="padding:3px 12px 3px 0;white-space:nowrap">👥 Guest fee</td><td style="padding:3px 0">${guestFee}</td></tr>` : ''}
      </table>
      ${event.Description ? `<p style="margin:12px 0 0;font-size:13px;color:#666">${Utils.escape(event.Description)}</p>` : ''}
    </div>

    <p style="text-align:center;font-weight:600;margin:0 0 16px;font-size:15px">Will you attend?</p>
    <div style="text-align:center;margin-bottom:28px">
      <a href="${rsvpYes}" style="background:#16a34a;color:#fff;padding:13px 28px;border-radius:6px;text-decoration:none;font-size:15px;font-weight:bold;display:inline-block;margin:4px 6px">✅ Yes, I'll be there!</a>
      <a href="${rsvpNo}"  style="background:#fff;color:#666;padding:12px 28px;border-radius:6px;text-decoration:none;font-size:15px;font-weight:bold;display:inline-block;margin:4px 6px;border:1px solid #ccc">✕ Sorry, can't make it</a>
    </div>

    <p style="font-size:12px;color:#aaa;text-align:center;margin:0">
      This invitation was sent by ${Utils.escape(Auth.getUserName())} via the Swiss Link portal.<br>
      Having trouble with the buttons? <a href="${rsvpUrl}" style="color:#CC0000">Open RSVP page</a>
    </p>
  </div>
</body></html>`;
  }

  return { send, buildInviteEmail };
})();
