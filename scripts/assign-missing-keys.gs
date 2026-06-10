/**
 * assign-missing-keys.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * Assigns MBR-NNNN keys to any Members rows that have a blank Member Key.
 * Safe to run multiple times — only touches rows that are still blank.
 *
 * HOW TO RUN:
 *   1. Open the Google Sheet → Extensions → Apps Script
 *   2. Paste this file, click Save
 *   3. Run → assignMissingKeys
 * ─────────────────────────────────────────────────────────────────────────────
 */

function assignMissingKeys() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName('Members');

  if (!membersSheet) { ui.alert('❌ Could not find a sheet named "Members".'); return; }

  const data    = membersSheet.getDataRange().getValues();
  const headers = data[0];

  const keyCol   = headers.indexOf('Member Key');
  const lastCol  = headers.indexOf('Last Name');
  const firstCol = headers.indexOf('First Name');

  if (keyCol === -1)  { ui.alert('❌ "Member Key" column not found.'); return; }
  if (lastCol === -1) { ui.alert('❌ "Last Name" column not found.'); return; }

  // Find highest existing MBR number
  let maxNum = 0;
  data.slice(1).forEach(row => {
    const m = String(row[keyCol] || '').match(/^MBR-(\d+)$/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });

  // Collect rows with blank keys, sorted by Last + First name
  const blanks = [];
  data.slice(1).forEach((row, i) => {
    const key = String(row[keyCol] || '').trim();
    if (key) return; // already has a key
    blanks.push({
      sheetRow:  i + 2,
      lastName:  String(row[lastCol]  || '').trim().toLowerCase(),
      firstName: String(row[firstCol] || '').trim().toLowerCase(),
    });
  });

  if (!blanks.length) {
    ui.alert('ℹ️ No blank Member Keys found. Nothing to do.');
    return;
  }

  blanks.sort((a, b) => {
    if (a.lastName  !== b.lastName)  return a.lastName  < b.lastName  ? -1 : 1;
    if (a.firstName !== b.firstName) return a.firstName < b.firstName ? -1 : 1;
    return 0;
  });

  // Assign new keys
  const pad = n => String(n).padStart(4, '0');
  let counter = maxNum + 1;
  blanks.forEach(r => { r.newKey = `MBR-${pad(counter++)}`; });

  // Confirm
  const sample = blanks.slice(0, 5)
    .map(r => `  Row ${r.sheetRow}: ${r.lastName}, ${r.firstName} → ${r.newKey}`)
    .join('\n');
  const confirm = ui.alert(
    'Assign Missing Keys',
    `${blanks.length} member(s) have no key. Assign MBR-${pad(maxNum + 1)} through MBR-${pad(counter - 1)}?\n\nSample:\n${sample}` +
    (blanks.length > 5 ? `\n  … and ${blanks.length - 5} more` : '') + '\n\nProceed?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // Write keys back to sheet
  blanks.forEach(r => {
    membersSheet.getRange(r.sheetRow, keyCol + 1).setValue(r.newKey);
  });

  ui.alert(`✅ Done! Assigned ${blanks.length} new Member Key(s):\nMBR-${pad(maxNum + 1)} through MBR-${pad(counter - 1)}.`);
}
