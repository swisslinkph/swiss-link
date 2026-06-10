/**
 * backfill-transaction-months.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * Adds a "Month" column to the Transactions sheet and populates it for every
 * existing row by reading the Timestamp field.
 *
 * Month is stored as a 1-based integer (1 = Jan … 12 = Dec).
 * Rows that already have a Month value are skipped.
 *
 * HOW TO RUN:
 *   1. Open the Google Sheet → Extensions → Apps Script
 *   2. Paste this file, click Save
 *   3. Run → backfillTransactionMonths
 * ─────────────────────────────────────────────────────────────────────────────
 */

function backfillTransactionMonths() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');

  if (!sheet) { ui.alert('❌ Could not find a sheet named "Transactions".'); return; }

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];

  const tsCol    = headers.indexOf('Timestamp');
  let   monthCol = headers.indexOf('Month');
  const yearCol  = headers.indexOf('Year');

  if (tsCol === -1) { ui.alert('❌ "Timestamp" column not found.'); return; }

  // ── Add Month column if it doesn't exist ─────────────────────────────────
  if (monthCol === -1) {
    // Insert Month right after Year (or at the end if Year not found)
    const insertAfter = yearCol !== -1 ? yearCol : headers.length - 1;
    monthCol = insertAfter + 1;

    sheet.insertColumnAfter(insertAfter + 1); // 1-based
    sheet.getRange(1, monthCol + 1).setValue('Month');

    ui.alert('ℹ️ "Month" column added. Running backfill now…');
  }

  // ── Backfill Month values ─────────────────────────────────────────────────
  let filled = 0, skipped = 0;
  const updates = [];

  data.slice(1).forEach((row, i) => {
    const existing = String(row[monthCol] || '').trim();
    if (existing !== '' && !isNaN(parseInt(existing))) { skipped++; return; }

    const ts = row[tsCol];
    if (!ts) { skipped++; return; }

    let month;
    if (ts instanceof Date) {
      month = ts.getMonth() + 1; // getMonth() is 0-based
    } else {
      // Parse ISO string: extract MM from YYYY-MM-DD
      const m = String(ts).match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (m) {
        month = parseInt(m[2], 10);
      } else {
        const d = new Date(ts);
        month = isNaN(d) ? null : d.getMonth() + 1;
      }
    }

    if (!month) { skipped++; return; }

    updates.push({ row: i + 2, month }); // +2: 1-based + skip header
    filled++;
  });

  if (updates.length === 0) {
    ui.alert(`✅ All rows already have a Month value (${skipped} skipped).`);
    return;
  }

  // Batch write for performance
  updates.forEach(u => {
    sheet.getRange(u.row, monthCol + 1).setValue(u.month);
  });

  ui.alert(
    `✅ Done!\n\nFilled: ${filled} row(s)\nSkipped (already set or no timestamp): ${skipped}`
  );
}
