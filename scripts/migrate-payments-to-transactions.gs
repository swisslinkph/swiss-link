/**
 * migrate-payments-to-transactions.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * One-time migration script.
 * Reads payment data from the Members sheet and creates matching rows in the
 * Transactions sheet for:
 *   1. 2025 Membership Dues
 *   2. 2026 Membership Dues
 *   3. AGM 2026 Event attendance
 *
 * HOW TO RUN:
 *   1. Open the Google Sheet
 *   2. Extensions → Apps Script
 *   3. Paste this entire file into the editor (replace any existing code)
 *   4. Click Save, then click Run → migratePaymentsToTransactions
 *   5. Approve the permissions popup when prompted
 *   6. Check the popup summary when it finishes
 *
 * SAFE TO RE-RUN: The script checks for existing TransactionIDs and skips
 * any row that already exists — so running it twice won't create duplicates.
 * ─────────────────────────────────────────────────────────────────────────────
 */

function migratePaymentsToTransactions() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName('Members');
  const txnSheet     = ss.getSheetByName('Transactions');

  if (!membersSheet) { SpreadsheetApp.getUi().alert('❌ Could not find a sheet named "Members".'); return; }
  if (!txnSheet)     { SpreadsheetApp.getUi().alert('❌ Could not find a sheet named "Transactions".'); return; }

  // ── Read Members ──────────────────────────────────────────────────────────
  const membersData    = membersSheet.getDataRange().getValues();
  const membersHeaders = membersData[0];

  // ── Read Transactions (to check for existing IDs) ─────────────────────────
  const txnData    = txnSheet.getDataRange().getValues();
  const txnHeaders = txnData[0];

  const TXN_COLS = [
    'TransactionID','Timestamp','MemberKey','MemberName','EventID','EventName',
    'AmountPaid','PaymentMode','Category','Year','HeadCount','Notes','RecordedBy'
  ];

  // Verify Transactions sheet has expected headers
  const missingCols = TXN_COLS.filter(h => !txnHeaders.includes(h));
  if (missingCols.length) {
    SpreadsheetApp.getUi().alert(
      '❌ Transactions sheet is missing columns: ' + missingCols.join(', ') +
      '\nMake sure the Transactions sheet has the correct headers first.'
    );
    return;
  }

  // Build set of existing TransactionIDs to prevent duplicates
  const txnIdColIdx  = txnHeaders.indexOf('TransactionID');
  const existingIds  = new Set(txnData.slice(1).map(r => String(r[txnIdColIdx]).trim()).filter(Boolean));

  // ── Column lookup helpers ─────────────────────────────────────────────────
  const mc = name => membersHeaders.indexOf(name); // member column index

  const M = {
    key:          mc('Member Key'),
    first:        mc('First Name'),
    last:         mc('Last Name'),
    // 2025 Membership
    amount2025:   mc('2025 Membership Amount (PHP)'),
    mode2025:     mc('2025 Payment Mode'),
    date2025:     mc('2025 Payment Date'),
    notes2025:    mc('2025 Payment Notes'),
    invoice2025:  mc('2025 Invoice Number'),
    // 2026 Membership
    amount2026:   mc('2026 Membership Amount (PHP)'),
    mode2026:     mc('2026 Payment Mode'),
    date2026:     mc('2026 Payment Date'),
    notes2026:    mc('2026 Payment Notes'),
    // AGM 2026
    agmAmount:    mc('AGM 2026 Paid Amount'),
    agmMode:      mc('AGM 2026 Payment Mode'),
    agmDate:      mc('AGM 2026 Payment Date'),
    agmHeadCount: mc('AGM 2026 Total Head Count'),
    agmRemarks:   mc('AGM 2026 Remarks'),
  };

  // Warn if any expected Members column is missing
  const missingMemberCols = Object.entries(M)
    .filter(([,v]) => v === -1)
    .map(([k]) => k);
  if (missingMemberCols.length) {
    Logger.log('Warning: some Members columns not found: ' + missingMemberCols.join(', '));
  }

  // ── ID counters (start after any that already exist) ─────────────────────
  let seq2025 = 1, seq2026 = 1, seqAGM = 1;
  existingIds.forEach(id => {
    let m;
    if ((m = id.match(/^MEM-2025-(\d+)$/))) seq2025 = Math.max(seq2025, +m[1] + 1);
    if ((m = id.match(/^MEM-2026-(\d+)$/))) seq2026 = Math.max(seq2026, +m[1] + 1);
    if ((m = id.match(/^AGM-2026-(\d+)$/))) seqAGM  = Math.max(seqAGM,  +m[1] + 1);
  });
  const pad = n => String(n).padStart(3, '0');

  // ── Build output rows ─────────────────────────────────────────────────────
  const newRows = [];
  let count2025 = 0, count2026 = 0, countAGM = 0, skipped = 0;

  const safeStr  = v => String(v || '').trim();
  const safeNum  = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
  const safeDate = v => {
    if (!v) return '';
    try { const d = new Date(v); return isNaN(d) ? '' : d.toISOString(); }
    catch(e) { return ''; }
  };

  const makeRow = (id, ts, eventId, eventName, amount, mode, category, year, headCount, notes, memberKey, memberName) => {
    const obj = {};
    TXN_COLS.forEach(h => obj[h] = '');
    obj.TransactionID = id;
    obj.Timestamp     = ts;
    obj.MemberKey     = memberKey;
    obj.MemberName    = memberName;
    obj.EventID       = eventId;
    obj.EventName     = eventName;
    obj.AmountPaid    = amount;
    obj.PaymentMode   = mode;
    obj.Category      = category;
    obj.Year          = year;
    obj.HeadCount     = headCount;
    obj.Notes         = notes;
    obj.RecordedBy    = 'Migration Script';
    return TXN_COLS.map(h => obj[h]);
  };

  membersData.slice(1).forEach(row => {
    const memberKey  = safeStr(row[M.key]);
    if (!memberKey) return;

    const firstName  = safeStr(row[M.first]);
    const lastName   = safeStr(row[M.last]);
    const memberName = [firstName, lastName].filter(Boolean).join(' ');

    // ── 2025 Membership Dues ────────────────────────────────────────────────
    const amt2025 = safeNum(row[M.amount2025]);
    if (amt2025 > 0) {
      // Use invoice number as ID if available, otherwise generate one
      const invoiceNum = safeStr(row[M.invoice2025]);
      const txnId      = invoiceNum || `MEM-2025-${pad(seq2025++)}`;

      if (existingIds.has(txnId)) {
        skipped++;
      } else {
        newRows.push(makeRow(
          txnId,
          safeDate(row[M.date2025]) || '2025-01-01T00:00:00.000Z',
          '', '2025 Membership Dues',
          amt2025, safeStr(row[M.mode2025]),
          'Membership', 2025, 1,
          safeStr(row[M.notes2025]),
          memberKey, memberName
        ));
        existingIds.add(txnId); // prevent dupe if same invoice number appears twice
        count2025++;
      }
    }

    // ── 2026 Membership Dues ────────────────────────────────────────────────
    const amt2026 = safeNum(row[M.amount2026]);
    if (amt2026 > 0) {
      const txnId = `MEM-2026-${pad(seq2026++)}`;

      if (existingIds.has(txnId)) {
        skipped++;
      } else {
        newRows.push(makeRow(
          txnId,
          safeDate(row[M.date2026]) || '2026-01-01T00:00:00.000Z',
          '', '2026 Membership Dues',
          amt2026, safeStr(row[M.mode2026]),
          'Membership', 2026, 1,
          safeStr(row[M.notes2026]),
          memberKey, memberName
        ));
        existingIds.add(txnId);
        count2026++;
      }
    }

    // ── AGM 2026 Event ──────────────────────────────────────────────────────
    const agmAmt = safeNum(row[M.agmAmount]);
    if (agmAmt > 0) {
      const txnId     = `AGM-2026-${pad(seqAGM++)}`;
      const headCount = parseInt(row[M.agmHeadCount]) || 1;

      if (existingIds.has(txnId)) {
        skipped++;
      } else {
        newRows.push(makeRow(
          txnId,
          safeDate(row[M.agmDate]) || '2026-02-12T00:00:00.000Z',
          'AGM-2026', 'AGM 2026',
          agmAmt, safeStr(row[M.agmMode]),
          'Event', 2026, headCount,
          safeStr(row[M.agmRemarks]),
          memberKey, memberName
        ));
        existingIds.add(txnId);
        countAGM++;
      }
    }
  });

  // ── Write to Transactions sheet ───────────────────────────────────────────
  if (newRows.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No new rows to add.\n' +
      (skipped ? `${skipped} row(s) already existed and were skipped.` : 'No payment data found in the Members sheet.')
    );
    return;
  }

  const startRow = txnSheet.getLastRow() + 1;
  txnSheet.getRange(startRow, 1, newRows.length, TXN_COLS.length).setValues(newRows);

  SpreadsheetApp.getUi().alert(
    '✅ Migration complete!\n\n' +
    `2025 membership dues:  ${count2025} rows\n` +
    `2026 membership dues:  ${count2026} rows\n` +
    `AGM 2026 attendance:   ${countAGM} rows\n` +
    (skipped ? `Skipped (already existed): ${skipped}\n` : '') +
    `\nTotal rows added: ${newRows.length}\n\n` +
    'Next step: verify the Transactions sheet looks correct,\n' +
    'then delete columns K–AL from the Members sheet.'
  );
}
