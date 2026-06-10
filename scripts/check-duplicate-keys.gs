/**
 * check-duplicate-keys.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * Scans the Members sheet for duplicate Member Keys and highlights them.
 * Read-only — does not change any data.
 *
 * HOW TO RUN:
 *   1. Open the Google Sheet → Extensions → Apps Script
 *   2. Paste this file, click Save
 *   3. Run → checkDuplicateKeys
 * ─────────────────────────────────────────────────────────────────────────────
 */

function checkDuplicateKeys() {
  const ui   = SpreadsheetApp.getUi();
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Members');

  if (!sheet) { ui.alert('❌ Could not find a sheet named "Members".'); return; }

  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyCol  = headers.indexOf('Member Key');
  const lastCol = headers.indexOf('Last Name');
  const firstCol = headers.indexOf('First Name');

  if (keyCol === -1) { ui.alert('❌ "Member Key" column not found.'); return; }

  // Build a map of key → rows that share it
  const keyMap = {};
  data.slice(1).forEach((row, i) => {
    const key = String(row[keyCol] || '').trim();
    if (!key) return;
    if (!keyMap[key]) keyMap[key] = [];
    keyMap[key].push({
      sheetRow:  i + 2,
      lastName:  String(row[lastCol]  || '').trim(),
      firstName: String(row[firstCol] || '').trim(),
    });
  });

  const dupes = Object.entries(keyMap).filter(([, rows]) => rows.length > 1);

  if (!dupes.length) {
    ui.alert('✅ No duplicate Member Keys found. All keys are unique!');
    return;
  }

  // Highlight duplicate rows in red for easy identification
  const RED = '#fca5a5';
  sheet.getDataRange().setBackground(null); // clear existing highlights first
  dupes.forEach(([, rows]) => {
    rows.forEach(r => {
      sheet.getRange(r.sheetRow, 1, 1, sheet.getLastColumn()).setBackground(RED);
    });
  });

  const report = dupes.map(([key, rows]) =>
    `${key}:\n` + rows.map(r => `  Row ${r.sheetRow}: ${r.firstName} ${r.lastName}`).join('\n')
  ).join('\n\n');

  ui.alert(
    `⚠️ Found ${dupes.length} duplicate key(s) — highlighted in red:\n\n${report}\n\n` +
    'Review the highlighted rows and delete or reassign the duplicates manually.'
  );
}
