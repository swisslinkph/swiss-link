/**
 * migrate-to-family-head.gs
 *
 * Run ONCE from the Apps Script editor (Extensions → Apps Script).
 *
 * What it does:
 *   1. Adds a "Family Head" column to the Members sheet (after "Family Group").
 *   2. For each existing named family group, sorts members by Last Name then
 *      First Name and auto-picks the first as the family head.
 *   3. Sets Family Head = head's own Member Key for the head,
 *      and Family Head = head's Member Key for all other members in the group.
 *
 * The original "Family Group" column is left untouched for reference.
 * Members with no Family Group are left with a blank Family Head.
 */
function migrateToFamilyHead() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Members');

  if (!sheet) { SpreadsheetApp.getUi().alert('Sheet "Members" not found.'); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // ── Step 1: Ensure "Family Head" column exists ────────────────────────────
  let famHeadCol = headers.indexOf('Family Head') + 1; // 1-based; 0 = not found

  if (!famHeadCol) {
    const famGroupCol = headers.indexOf('Family Group') + 1;
    if (!famGroupCol) {
      SpreadsheetApp.getUi().alert('"Family Group" column not found in Members sheet.');
      return;
    }
    sheet.insertColumnAfter(famGroupCol);
    famHeadCol = famGroupCol + 1;
    sheet.getRange(1, famHeadCol).setValue('Family Head');
    Logger.log('Inserted "Family Head" column at position ' + famHeadCol);
  } else {
    Logger.log('"Family Head" column already exists at position ' + famHeadCol);
  }

  // ── Step 2: Re-read headers after possible column insert ──────────────────
  const headersFresh  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyCol        = headersFresh.indexOf('Member Key')   + 1;
  const lastNameCol   = headersFresh.indexOf('Last Name')    + 1;
  const firstNameCol  = headersFresh.indexOf('First Name')   + 1;
  const famGroupCol   = headersFresh.indexOf('Family Group') + 1;
  const famHeadColF   = headersFresh.indexOf('Family Head')  + 1;

  if (!keyCol || !famGroupCol) {
    SpreadsheetApp.getUi().alert('Required columns missing: Member Key or Family Group.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('No member rows found.'); return; }

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // ── Step 3: Build groups: groupName → [{ key, lastName, firstName, row }] ─
  const groups = {};
  data.forEach((row, i) => {
    const key   = String(row[keyCol - 1]       || '').trim();
    const grp   = String(row[famGroupCol - 1]  || '').trim();
    const last  = String(row[lastNameCol - 1]  || '').trim();
    const first = String(row[firstNameCol - 1] || '').trim();
    if (key && grp) {
      if (!groups[grp]) groups[grp] = [];
      groups[grp].push({ key, lastName: last, firstName: first, rowIndex: i + 2 });
    }
  });

  // ── Step 4: For each group, pick head and write Family Head column ─────────
  let groupsProcessed = 0;
  let membersUpdated  = 0;

  Object.entries(groups).forEach(([grp, members]) => {
    // Sort by last name, then first name (alphabetical)
    members.sort((a, b) => {
      const cmp = a.lastName.localeCompare(b.lastName);
      return cmp !== 0 ? cmp : a.firstName.localeCompare(b.firstName);
    });

    const headKey = members[0].key;
    Logger.log(`Group "${grp}": head = ${headKey} (${members[0].firstName} ${members[0].lastName})`);

    members.forEach(m => {
      sheet.getRange(m.rowIndex, famHeadColF).setValue(headKey);
      membersUpdated++;
    });
    groupsProcessed++;
  });

  SpreadsheetApp.getUi().alert(
    'Migration complete!\n\n' +
    'Family groups migrated: ' + groupsProcessed + '\n' +
    'Members updated: ' + membersUpdated + '\n\n' +
    'The original "Family Group" column is unchanged.'
  );
}
