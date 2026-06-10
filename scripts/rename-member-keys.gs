/**
 * rename-member-keys.gs
 * ─────────────────────────────────────────────────────────────────────────────
 * One-time migration script.
 * Renames all Member Keys from "lastname|firstname" format to "MBR-NNNN",
 * consistent with EventID (EVT-YYYY-NNN) and TransactionID (TXN-YYYY-NNN).
 *
 * Updates:
 *   - "Member Key" column in the Members sheet
 *   - "MemberKey" column in the Transactions sheet
 *
 * Members are sorted alphabetically by Last Name before numbering, so
 * MBR-0001 = first alphabetically.
 *
 * HOW TO RUN:
 *   1. Open the Google Sheet
 *   2. Extensions → Apps Script
 *   3. Paste this file into the editor, click Save
 *   4. Run → renameMemberKeys
 *   5. Click Yes on the confirmation dialog
 *
 * SAFE TO RE-RUN: Already-renamed keys (MBR-NNNN) are left untouched.
 * ─────────────────────────────────────────────────────────────────────────────
 */

function renameMemberKeys() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const membersSheet = ss.getSheetByName('Members');
  const txnSheet     = ss.getSheetByName('Transactions');

  if (!membersSheet) { ui.alert('❌ Could not find a sheet named "Members".'); return; }
  if (!txnSheet)     { ui.alert('❌ Could not find a sheet named "Transactions".'); return; }

  // ── Confirm ───────────────────────────────────────────────────────────────
  const confirm = ui.alert(
    'Rename Member Keys',
    'This will rename all Member Keys from "lastname|firstname" format to "MBR-0001" style.\n\n' +
    'Both the Members sheet and Transactions sheet will be updated.\n\n' +
    'Make sure you have a Members_BACKUP tab before continuing.\n\nProceed?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // ── Read data ─────────────────────────────────────────────────────────────
  const membersRange = membersSheet.getDataRange();
  const membersData  = membersRange.getValues();
  const membersHdr   = membersData[0];

  const txnRange = txnSheet.getDataRange();
  const txnData  = txnRange.getValues();
  const txnHdr   = txnData[0];

  const mKeyCol    = membersHdr.indexOf('Member Key');
  const mLastCol   = membersHdr.indexOf('Last Name');
  const mFirstCol  = membersHdr.indexOf('First Name');
  const tKeyCol    = txnHdr.indexOf('MemberKey');

  const missing = [];
  if (mKeyCol   === -1) missing.push('"Member Key" in Members');
  if (mLastCol  === -1) missing.push('"Last Name" in Members');
  if (mFirstCol === -1) missing.push('"First Name" in Members');
  if (tKeyCol   === -1) missing.push('"MemberKey" in Transactions');
  if (missing.length) { ui.alert('❌ Missing columns:\n' + missing.join('\n')); return; }

  // ── Find highest existing MBR number (idempotency) ────────────────────────
  let maxNum = 0;
  membersData.slice(1).forEach(row => {
    const m = String(row[mKeyCol] || '').match(/^MBR-(\d+)$/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });

  // ── Collect members that need renaming, sorted by Last + First name ───────
  const toRename = [];
  membersData.slice(1).forEach((row, i) => {
    const key = String(row[mKeyCol] || '').trim();
    if (!key || /^MBR-\d+$/.test(key)) return; // skip blank or already done
    toRename.push({
      sheetRow: i + 2,       // 1-based sheet row (row 1 = headers)
      oldKey:   key,
      lastName: String(row[mLastCol]  || '').trim().toLowerCase(),
      firstName: String(row[mFirstCol] || '').trim().toLowerCase(),
    });
  });

  if (!toRename.length) {
    ui.alert('ℹ️ All Member Keys are already in MBR-NNNN format. Nothing to do.');
    return;
  }

  toRename.sort((a, b) => {
    if (a.lastName  !== b.lastName)  return a.lastName  < b.lastName  ? -1 : 1;
    if (a.firstName !== b.firstName) return a.firstName < b.firstName ? -1 : 1;
    return 0;
  });

  // ── Assign new keys ───────────────────────────────────────────────────────
  const pad    = n  => String(n).padStart(4, '0');
  const keyMap = {}; // oldKey → newKey

  let counter = maxNum + 1;
  toRename.forEach(m => {
    m.newKey       = `MBR-${pad(counter++)}`;
    keyMap[m.oldKey] = m.newKey;
  });

  // ── Batch-update Members sheet (entire key column in one write) ───────────
  const memberKeyCol = membersData.map((row, i) => {
    if (i === 0) return [row[mKeyCol]]; // header unchanged
    const old = String(row[mKeyCol] || '').trim();
    return [keyMap[old] ?? row[mKeyCol]]; // replace if in map, else keep
  });
  membersSheet.getRange(1, mKeyCol + 1, memberKeyCol.length, 1).setValues(memberKeyCol);

  // ── Batch-update Transactions sheet ──────────────────────────────────────
  let txnUpdated = 0;
  const txnKeyColData = txnData.map((row, i) => {
    if (i === 0) return [row[tKeyCol]]; // header unchanged
    const old = String(row[tKeyCol] || '').trim();
    if (keyMap[old]) { txnUpdated++; return [keyMap[old]]; }
    return [row[tKeyCol]];
  });
  txnSheet.getRange(1, tKeyCol + 1, txnKeyColData.length, 1).setValues(txnKeyColData);

  // ── Summary ───────────────────────────────────────────────────────────────
  const sample = toRename.slice(0, 6)
    .map(m => `  ${m.oldKey.padEnd(35)} → ${m.newKey}`)
    .join('\n');

  ui.alert(
    '✅ Rename complete!\n\n' +
    `Members renamed:          ${toRename.length}\n` +
    `Transaction rows updated: ${txnUpdated}\n\n` +
    'Sample:\n' + sample +
    (toRename.length > 6 ? `\n  … and ${toRename.length - 6} more` : '')
  );
}
