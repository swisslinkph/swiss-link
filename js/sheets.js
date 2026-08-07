/**
 * sheets.js — Google Sheets API v4 wrapper
 * Provides easy read/write access to the Members, Events, and Transactions sheets.
 * All rows are returned as plain objects keyed by column header.
 */

const Sheets = (() => {
  const BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

  // ── Internal request helper ───────────────────────────────────────────────
  async function request(path, options = {}) {
    const token = Auth.getToken();
    if (!token) throw new Error('Not authenticated');

    const url = path.startsWith('http') ? path : `${BASE}/${CONFIG.SHEET_ID}${path}`;
    const res  = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        ...(options.headers || {}),
      },
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: { message: res.statusText }}));
      throw new Error(err.error?.message || `Sheets API error ${res.status}`);
    }
    return options.method === 'DELETE' ? null : res.json();
  }

  // ── Cache for header rows and sheet gids ─────────────────────────────────
  const _headerCache = {};
  const _sheetIdCache = {};

  async function getHeaders(sheetName) {
    if (_headerCache[sheetName]) return _headerCache[sheetName];
    const range = encodeURIComponent(`${sheetName}!1:1`);
    const data  = await request(`/values/${range}`);
    const headers = (data.values?.[0] || []).map(h => h.trim());
    _headerCache[sheetName] = headers;
    return headers;
  }

  function clearHeaderCache(sheetName) {
    delete _headerCache[sheetName];
  }

  // ── Convert array row → object ────────────────────────────────────────────
  function rowToObj(headers, row) {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] ?? ''; });
    obj._rowIndex = undefined; // set by caller
    return obj;
  }

  // ── Convert object → array row (in header order) ─────────────────────────
  function objToRow(headers, obj) {
    return headers.map(h => obj[h] ?? '');
  }

  // ── READ all rows ─────────────────────────────────────────────────────────
  async function getAll(sheetName) {
    const headers = await getHeaders(sheetName);
    const range   = encodeURIComponent(`${sheetName}!A:${colLetter(headers.length)}`);
    const data    = await request(`/values/${range}`);
    const rows    = data.values || [];
    return rows.slice(1).map((row, i) => {
      const obj = rowToObj(headers, row);
      obj._rowIndex = i + 2; // 1-based, row 1 = headers
      return obj;
    }).filter(obj => {
      // Filter out completely empty rows
      return headers.some(h => h !== '_rowIndex' && obj[h]?.trim());
    });
  }

  // ── APPEND a new row ──────────────────────────────────────────────────────
  async function append(sheetName, obj) {
    const headers  = await getHeaders(sheetName);
    const row      = objToRow(headers, obj);
    // Find the true last row by scanning column A — avoids blank separator
    // rows tricking the Sheets API into inserting before existing data.
    const colA     = encodeURIComponent(`${sheetName}!A:A`);
    const colAData = await request(`/values/${colA}`);
    const nextRow  = (colAData.values?.length ?? 1) + 1;
    const range    = encodeURIComponent(`${sheetName}!A${nextRow}:${colLetter(headers.length)}${nextRow}`);
    await request(`/values/${range}?valueInputOption=USER_ENTERED`, {
      method: 'PUT',
      body: JSON.stringify({ values: [row] }),
    });
  }

  // ── UPDATE an existing row by row index ───────────────────────────────────
  async function update(sheetName, rowIndex, obj) {
    const headers = await getHeaders(sheetName);
    const row     = objToRow(headers, obj);
    const range   = encodeURIComponent(`${sheetName}!A${rowIndex}:${colLetter(headers.length)}${rowIndex}`);
    await request(`/values/${range}?valueInputOption=USER_ENTERED`, {
      method: 'PUT',
      body: JSON.stringify({ values: [row] }),
    });
  }

  // ── DELETE a row by row index (clears it; use with caution) ──────────────
  async function deleteRow(sheetName, rowIndex) {
    // Resolve numeric sheet gid, caching to avoid repeated metadata fetches
    // Use `=== undefined` not `!` — sheetId can legitimately be 0 (first sheet)
    if (_sheetIdCache[sheetName] === undefined) {
      const meta = await request('');
      meta.sheets.forEach(s => { _sheetIdCache[s.properties.title] = s.properties.sheetId; });
    }
    const sheetId = _sheetIdCache[sheetName];
    if (sheetId === undefined) throw new Error(`Sheet "${sheetName}" not found`);

    // Note: structural batchUpdate uses :batchUpdate (colon), not /batchUpdate
    await request(':batchUpdate', {
      method: 'POST',
      body: JSON.stringify({
        requests: [{
          deleteDimension: {
            range: {
              sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex - 1,
              endIndex: rowIndex,
            },
          },
        }],
      }),
    });
  }

  // ── BATCH UPDATE multiple cells ───────────────────────────────────────────
  async function batchUpdate(updates) {
    // updates: [{ range: 'Sheet!A1', values: [[v]] }, ...]
    await request(`/values:batchUpdate`, {
      method: 'POST',
      body: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: updates,
      }),
    });
  }

  // ── ENSURE required sheets exist ─────────────────────────────────────────
  async function ensureSheets() {
    const meta     = await request('');
    const existing = meta.sheets.map(s => s.properties.title);
    const needed   = Object.values(CONFIG.SHEETS).filter(n => !existing.includes(n));

    if (needed.length) {
      await request(':batchUpdate', {
        method: 'POST',
        body: JSON.stringify({
          requests: needed.map(title => ({
            addSheet: { properties: { title } },
          })),
        }),
      });
    }

    // Migrate existing Registrations sheet — add any missing columns
    if (existing.includes(CONFIG.SHEETS.REGISTRATIONS)) {
      await ensureColumn(CONFIG.SHEETS.REGISTRATIONS, 'MemberSlots');
      await ensureColumn(CONFIG.SHEETS.REGISTRATIONS, 'CheckedIn');
      await ensureColumn(CONFIG.SHEETS.REGISTRATIONS, 'GuestNames');
      await ensureColumn(CONFIG.SHEETS.REGISTRATIONS, 'SlotPayments');
      await ensureColumn(CONFIG.SHEETS.REGISTRATIONS, 'WalkIn');
      await ensureColumn(CONFIG.SHEETS.REGISTRATIONS, 'KidsNames');
    }

    // Migrate existing Events sheet — add any missing columns
    if (existing.includes(CONFIG.SHEETS.EVENTS)) {
      await ensureColumn(CONFIG.SHEETS.EVENTS, 'KidsFee');
      await ensureColumn(CONFIG.SHEETS.EVENTS, 'WalkInMemberFee');
      await ensureColumn(CONFIG.SHEETS.EVENTS, 'WalkInGuestFee');
      await ensureColumn(CONFIG.SHEETS.EVENTS, 'FormSheetID');
      await ensureColumn(CONFIG.SHEETS.EVENTS, 'FormSheetTab');
    }

    // Write headers for any sheet that exists but is empty
    const allSheets = [...needed, ...existing];
    for (const name of allSheets) {
      if (name === CONFIG.SHEETS.EVENTS) {
        const h = await getHeaders(name);
        if (!h.length) await _writeEventsHeaders();
      }
      if (name === CONFIG.SHEETS.TRANSACTIONS) {
        const h = await getHeaders(name);
        if (!h.length) await _writeTxHeaders();
      }
      if (name === CONFIG.SHEETS.ADMINS) {
        const h = await getHeaders(name);
        if (!h.length) await _writeAdminsHeaders();
      }
      if (name === CONFIG.SHEETS.REGISTRATIONS) {
        const h = await getHeaders(name);
        if (!h.length) await _writeRegistrationsHeaders();
      }
    }
  }

  async function _writeRegistrationsHeaders() {
    const headers = [
      'RegistrationID','Timestamp','Source','EventID','EventName',
      'LastName','FirstName','Email','MemberKey','MemberSlots',
      'MemberQty','GuestQty','KidsQty','WalkIn',
      'TotalDue','PaymentNote','PaymentStatus','PaymentMode','AmountPaid','AdminNotes','CheckedIn','GuestNames','SlotPayments',
    ];
    const range = encodeURIComponent(`${CONFIG.SHEETS.REGISTRATIONS}!A1`);
    await request(`/values/${range}?valueInputOption=USER_ENTERED`, {
      method: 'PUT',
      body: JSON.stringify({ values: [headers] }),
    });
    clearHeaderCache(CONFIG.SHEETS.REGISTRATIONS);
  }

  async function _writeAdminsHeaders() {
    const headers = ['Email', 'Name', 'Added Date', 'Added By'];
    const range   = encodeURIComponent(`${CONFIG.SHEETS.ADMINS}!A1`);
    await request(`/values/${range}?valueInputOption=USER_ENTERED`, {
      method: 'PUT',
      body: JSON.stringify({ values: [headers] }),
    });
    clearHeaderCache(CONFIG.SHEETS.ADMINS);
  }

  async function _writeEventsHeaders() {
    const headers = [
      'EventID','Title','Date','Location','Description',
      'MemberFee','GuestFee','KidsFee','WalkInMemberFee','WalkInGuestFee',
      'RSVPFormURL','Status','CreatedDate','CreatedBy',
    ];
    const range = encodeURIComponent(`${CONFIG.SHEETS.EVENTS}!A1`);
    await request(`/values/${range}?valueInputOption=USER_ENTERED`, {
      method: 'PUT',
      body: JSON.stringify({ values: [headers] }),
    });
    clearHeaderCache(CONFIG.SHEETS.EVENTS);
  }

  async function _writeTxHeaders() {
    const headers = [
      'TransactionID','Timestamp','MemberKey','MemberName','EventID','EventName',
      'AmountPaid','PaymentMode','Category','Year','Month','HeadCount','Notes','RecordedBy',
    ];
    const range = encodeURIComponent(`${CONFIG.SHEETS.TRANSACTIONS}!A1`);
    await request(`/values/${range}?valueInputOption=USER_ENTERED`, {
      method: 'PUT',
      body: JSON.stringify({ values: [headers] }),
    });
    clearHeaderCache(CONFIG.SHEETS.TRANSACTIONS);
  }

  // ── READ from an external spreadsheet (e.g. Google Form responses) ──────
  async function getFromSheet(sheetId, tabName, headerRow) {
    headerRow = headerRow || 1;
    const base = `${BASE}/${sheetId}`;
    const hRange = encodeURIComponent(`${tabName}!${headerRow}:${headerRow}`);
    const hData  = await request(`${base}/values/${hRange}`);
    const headers = (hData.values?.[0] || []).map(h => h.toString().trim());
    if (!headers.length) return { headers: [], rows: [] };
    const dRange = encodeURIComponent(`${tabName}!A${headerRow + 1}:${colLetter(headers.length)}9999`);
    const data   = await request(`${base}/values/${dRange}`);
    return { headers, rows: data.values || [] };
  }

  // ── ADD a column header to a sheet if it doesn't already exist ───────────
  async function ensureColumn(sheetName, columnName) {
    const headers = await getHeaders(sheetName);
    if (headers.includes(columnName)) return;
    const nextCol = colLetter(headers.length + 1);
    const range   = encodeURIComponent(`${sheetName}!${nextCol}1`);
    await request(`/values/${range}?valueInputOption=USER_ENTERED`, {
      method: 'PUT',
      body: JSON.stringify({ values: [[columnName]] }),
    });
    clearHeaderCache(sheetName);
  }

  // ── GENERATE next ID for Events / Transactions ────────────────────────────
  async function nextId(sheetName, prefix) {
    const rows = await getAll(sheetName);
    const year = new Date().getFullYear();
    const existing = rows
      .map(r => r[sheetName === CONFIG.SHEETS.EVENTS ? 'EventID' : 'TransactionID'])
      .filter(id => id?.startsWith(`${prefix}-${year}`))
      .map(id => parseInt(id.split('-').pop(), 10))
      .filter(n => !isNaN(n));
    const next = existing.length ? Math.max(...existing) + 1 : 1;
    return `${prefix}-${year}-${String(next).padStart(3, '0')}`;
  }

  // ── Helper: column letter from 1-based index ──────────────────────────────
  function colLetter(n) {
    let s = '';
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s || 'A';
  }

  return {
    getAll, append, update, deleteRow, batchUpdate,
    getHeaders, clearHeaderCache, ensureSheets, nextId,
    getFromSheet, ensureColumn,
  };
})();
