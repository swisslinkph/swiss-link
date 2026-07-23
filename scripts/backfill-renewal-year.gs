/**
 * backfill-renewal-year.gs
 *
 * Run ONCE from the Apps Script editor (Extensions → Apps Script in the Google Sheet).
 *
 * What it does:
 *   1. Adds a "Renewal Year" column to the Members sheet (after "Membership Status")
 *      if the column doesn't already exist.
 *   2. For each member, finds their latest Membership transaction in the Transactions
 *      sheet and writes that year as the Renewal Year.
 *
 * Members with no Membership transactions are left blank (they can be filled
 * manually or via the app when dues are next recorded).
 */
function backfillRenewalYear() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName('Members');
  const txSheet      = ss.getSheetByName('Transactions');

  if (!membersSheet) { SpreadsheetApp.getUi().alert('Sheet "Members" not found.'); return; }
  if (!txSheet)      { SpreadsheetApp.getUi().alert('Sheet "Transactions" not found.'); return; }

  // ── Step 1: Ensure "Renewal Year" column exists in Members ───────────────
  const mHeaders = membersSheet
    .getRange(1, 1, 1, membersSheet.getLastColumn())
    .getValues()[0];

  let renewalCol = mHeaders.indexOf('Renewal Year') + 1; // 1-based, 0 = not found

  if (!renewalCol) {
    const statusCol = mHeaders.indexOf('Membership Status') + 1;
    if (!statusCol) {
      SpreadsheetApp.getUi().alert('"Membership Status" column not found in Members sheet.');
      return;
    }
    // Insert immediately after Membership Status
    membersSheet.insertColumnAfter(statusCol);
    renewalCol = statusCol + 1;
    membersSheet.getRange(1, renewalCol).setValue('Renewal Year');
    Logger.log('Inserted "Renewal Year" column at position ' + renewalCol);
  } else {
    Logger.log('"Renewal Year" column already exists at position ' + renewalCol);
  }

  // ── Step 2: Build map of memberKey → latest membership year ──────────────
  const txHeaders = txSheet
    .getRange(1, 1, 1, txSheet.getLastColumn())
    .getValues()[0];

  const txLastRow = txSheet.getLastRow();
  if (txLastRow < 2) {
    SpreadsheetApp.getUi().alert('No transactions found. Nothing to backfill.');
    return;
  }

  const txData = txSheet
    .getRange(2, 1, txLastRow - 1, txSheet.getLastColumn())
    .getValues();

  const memberKeyCol = txHeaders.indexOf('MemberKey');
  const categoryCol  = txHeaders.indexOf('Category');
  const yearCol      = txHeaders.indexOf('Year');

  if (memberKeyCol < 0 || categoryCol < 0 || yearCol < 0) {
    SpreadsheetApp.getUi().alert(
      'Missing required Transactions columns.\n' +
      'Expected: MemberKey, Category, Year.\n' +
      'Found: ' + txHeaders.join(', ')
    );
    return;
  }

  const latestYear = {};
  txData.forEach(row => {
    const key  = String(row[memberKeyCol] || '').trim();
    const cat  = String(row[categoryCol]  || '').trim();
    const year = parseInt(row[yearCol], 10);
    if (key && cat === 'Membership' && year) {
      latestYear[key] = Math.max(latestYear[key] || 0, year);
    }
  });

  Logger.log('Members with membership transactions: ' + Object.keys(latestYear).length);

  // ── Step 3: Write Renewal Year for each member ────────────────────────────
  const mLastRow = membersSheet.getLastRow();
  if (mLastRow < 2) {
    SpreadsheetApp.getUi().alert('No member rows found.');
    return;
  }

  // Re-read headers (column may have shifted after insert)
  const mHeadersFresh = membersSheet
    .getRange(1, 1, 1, membersSheet.getLastColumn())
    .getValues()[0];
  const keyColM = mHeadersFresh.indexOf('Member Key') + 1;

  const mData = membersSheet
    .getRange(2, keyColM, mLastRow - 1, 1)
    .getValues();

  let filled = 0, skipped = 0;
  mData.forEach((row, i) => {
    const key  = String(row[0] || '').trim();
    if (!key) { skipped++; return; }
    const year = latestYear[key];
    if (year) {
      membersSheet.getRange(i + 2, renewalCol).setValue(year);
      filled++;
    } else {
      skipped++;
    }
  });

  SpreadsheetApp.getUi().alert(
    'Backfill complete!\n\n' +
    'Renewal Year filled: ' + filled + ' member(s)\n' +
    'Skipped (no membership tx): ' + skipped + ' member(s)'
  );
}
