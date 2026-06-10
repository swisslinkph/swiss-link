/**
 * rsvp-relay.gs — Google Apps Script RSVP Relay
 *
 * Deployed as a Web App:
 *   Execute as: Me (your Google account)
 *   Who has access: Anyone (no sign-in required)
 *
 * Receives a POST from rsvp.html and appends a row to the
 * Transactions sheet with Category = 'RSVP'.
 *
 * Setup:
 *   1. Open the Google Sheet → Extensions → Apps Script
 *   2. Paste this script, save.
 *   3. Deploy → New deployment → Web App
 *      - Execute as: Me
 *      - Who can access: Anyone
 *   4. Copy the Web App URL into CONFIG.RSVP_RELAY_URL in js/config.js
 */

const SHEET_ID         = '1xOs_Gii-7N-y97wdaR5sr-5ancGQqz7BA2YNF85UauA';
const TRANSACTIONS_TAB = 'Transactions';
const EVENTS_TAB       = 'Events';
const MEMBERS_TAB      = 'Members';

/**
 * Handle POST from rsvp.html
 * Expected body (JSON):
 *   { eventId, memberKey, token, response ('yes'|'no'), timestamp }
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const { eventId, memberKey, token, response, timestamp } = data;

    // ── Basic validation ─────────────────────────────────────────
    if (!eventId || !memberKey || !token || !response) {
      return _json({ success: false, error: 'Missing required fields.' });
    }

    // ── Verify token (must match rsvp.html: base64 with = stripped) ─
    const expectedToken = Utilities.base64Encode(
      eventId + ':' + memberKey + ':' + SHEET_ID.slice(0, 8)
    ).replace(/=/g, '');
    if (token !== expectedToken) {
      return _json({ success: false, error: 'Invalid token.' });
    }

    // ── Open spreadsheet ──────────────────────────────────────────
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // ── Fetch event name ──────────────────────────────────────────
    const evtSheet  = ss.getSheetByName(EVENTS_TAB);
    let   eventName = eventId;
    if (evtSheet) {
      const evtData   = evtSheet.getDataRange().getValues();
      const evtHdrs   = evtData[0];
      const evtIdCol  = evtHdrs.indexOf('EventID');
      const evtTitCol = evtHdrs.indexOf('Title');
      if (evtIdCol >= 0 && evtTitCol >= 0) {
        const evtRow = evtData.find((r, i) => i > 0 && r[evtIdCol] === eventId);
        if (evtRow) eventName = evtRow[evtTitCol];
      }
    }

    // ── Fetch member name from Members sheet (MBR-NNNN format) ────
    let memberName = memberKey;
    const mbrSheet = ss.getSheetByName(MEMBERS_TAB);
    if (mbrSheet) {
      const mbrData   = mbrSheet.getDataRange().getValues();
      const mbrHdrs   = mbrData[0];
      const mbrKeyCol = mbrHdrs.indexOf('Member Key');
      const mbrFstCol = mbrHdrs.indexOf('First Name');
      const mbrLstCol = mbrHdrs.indexOf('Last Name');
      if (mbrKeyCol >= 0 && mbrFstCol >= 0 && mbrLstCol >= 0) {
        const mbrRow = mbrData.find((r, i) => i > 0 && r[mbrKeyCol] === memberKey);
        if (mbrRow) memberName = (mbrRow[mbrFstCol] + ' ' + mbrRow[mbrLstCol]).trim();
      }
    }

    // ── Build transaction row ─────────────────────────────────────
    const txnSheet = ss.getSheetByName(TRANSACTIONS_TAB);
    if (!txnSheet) {
      return _json({ success: false, error: 'Transactions sheet not found.' });
    }

    const headers  = txnSheet.getRange(1, 1, 1, txnSheet.getLastColumn()).getValues()[0];
    const nextId   = _nextTxnId(txnSheet, headers);
    const year     = new Date().getFullYear();

    const row = {
      TransactionID: nextId,
      Timestamp:     timestamp || new Date().toISOString(),
      MemberKey:     memberKey,
      MemberName:    memberName,
      EventID:       eventId,
      EventName:     eventName,
      AmountPaid:    0,
      PaymentMode:   '',
      Category:      'RSVP',
      Year:          year,
      HeadCount:     response === 'yes' ? 1 : 0,
      Notes:         response === 'yes' ? 'RSVP: attending' : 'RSVP: not attending',
      RecordedBy:    'rsvp-relay',
    };

    // Map to column order
    const values = headers.map(h => row[h] !== undefined ? row[h] : '');
    txnSheet.appendRow(values);

    return _json({ success: true, transactionId: nextId });

  } catch (err) {
    return _json({ success: false, error: err.message });
  }
}

/** Generate next TXN-YYYY-NNN id */
function _nextTxnId(sheet, headers) {
  const year   = new Date().getFullYear();
  const prefix = 'TXN-' + year + '-';
  const idCol  = headers.indexOf('TransactionID');
  if (idCol < 0) return prefix + '001';

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return prefix + '001';

  const ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues()
    .flat()
    .filter(v => String(v).startsWith(prefix))
    .map(v => parseInt(String(v).split('-').pop(), 10) || 0);

  const next = ids.length ? Math.max(...ids) + 1 : 1;
  return prefix + String(next).padStart(3, '0');
}

/** Return JSON response */
function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Optional: handle GET for CORS preflight (some browsers send it) */
function doGet(e) {
  return _json({ status: 'Swiss Club RSVP relay is running.' });
}
