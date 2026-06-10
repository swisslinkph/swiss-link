/**
 * members.js — Member Management view
 * Searchable table with Add / Edit / Delete via modal.
 */

const Members = (() => {
  let _all      = [];   // all member rows from sheet
  let _txns     = [];   // all transactions (for YTD totals)
  let _filtered = [];
  let _sortKey  = 'Last Name';
  let _sortDir  = 'asc';
  let _editingRow = null; // row index being edited (null = new member)

  // Column name constants (must match sheet headers exactly)
  const C = {
    KEY:    'Member Key',
    LAST:   'Last Name',
    ALT:    'Alternative Name',
    FIRST:  'First Name',
    EMAIL:  'Email',
    LOC:    'Location (Metro Manila/Province)',
    STATUS: '2026 Membership Status (Member, Exempt, Non-member, TBC)',
    TYPE:   'Membership Type',
    FAM:    'Family Group',
  };

  // ── Render (initial load) ─────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading members…');
    try {
      [_all, _txns] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.MEMBERS),
        Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS).catch(() => []),
      ]);
      _all = _all.filter(m => m[C.KEY]?.trim()); // skip empty / header rows
      _applyFilter(document.getElementById('member-search')?.value || '');
      _renderTable();
      _updateCount();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  // ── Filter + sort ─────────────────────────────────────────────────────────
  function _applyFilter(query) {
    _filtered = Utils.filterRows(_all, query, [C.FIRST, C.LAST, C.ALT, C.EMAIL, C.FAM, C.KEY]);
    _filtered = Utils.sortTable(_filtered, _sortKey, _sortDir);
  }

  function _updateCount() {
    const el = document.getElementById('member-count');
    if (el) el.textContent = `${_filtered.length} of ${_all.length} members`;
  }

  // ── Render table ──────────────────────────────────────────────────────────
  function _renderTable() {
    const tbody = document.querySelector('#members-table tbody');
    if (!tbody) return;

    if (!_filtered.length) {
      tbody.innerHTML = '<tr><td colspan="8" class="empty-state">No members found.</td></tr>';
      return;
    }

    tbody.innerHTML = _filtered.map(m => {
      const ytd    = Utils.totalPaidYTD(m[C.KEY], _txns);
      const fullName = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
      const initials = Utils.initials(m[C.FIRST] || m[C.LAST] || '?');
      return `<tr data-key="${Utils.escape(m[C.KEY])}">
        <td>
          <div class="member-cell">
            <div class="avatar">${initials}</div>
            <div>
              <div class="member-name">${Utils.escape(fullName)}</div>
              <div class="member-key">${Utils.escape(m[C.KEY])}</div>
            </div>
          </div>
        </td>
        <td>${Utils.escape(m[C.EMAIL])}</td>
        <td>${Utils.escape(m[C.LOC])}</td>
        <td>${Utils.statusBadge(m[C.STATUS])}</td>
        <td>${Utils.typeBadge(m[C.TYPE])}</td>
        <td>${Utils.escape(m[C.FAM])}</td>
        <td class="amount">${Utils.formatPHP(ytd)}</td>
        <td class="actions">
          <button class="btn-icon" title="Edit" onclick="Members.openEdit('${Utils.escape(m[C.KEY])}')">✏️</button>
          <button class="btn-icon btn-danger" title="Delete" onclick="Members.confirmDelete('${Utils.escape(m[C.KEY])}')">🗑️</button>
        </td>
      </tr>`;
    }).join('');
  }

  // ── Search handler (debounced) ────────────────────────────────────────────
  const _onSearch = Utils.debounce(query => {
    _applyFilter(query);
    _renderTable();
    _updateCount();
  });

  // ── Sort handler ──────────────────────────────────────────────────────────
  function sort(key) {
    if (_sortKey === key) _sortDir = _sortDir === 'asc' ? 'desc' : 'asc';
    else { _sortKey = key; _sortDir = 'asc'; }
    _applyFilter(document.getElementById('member-search')?.value || '');
    _renderTable();
    // Update sort icons
    document.querySelectorAll('#members-table th[data-sort]').forEach(th => {
      th.dataset.dir = th.dataset.sort === key ? _sortDir : '';
    });
  }

  // ── Open modal for Add ────────────────────────────────────────────────────
  function openAdd() {
    _editingRow = null;
    document.getElementById('member-modal-title').textContent = 'Add Member';
    document.getElementById('member-form').reset();
    document.getElementById('mf-key').value = '';
    Utils.showModal('member-modal');
  }

  // ── Open modal for Edit ───────────────────────────────────────────────────
  function openEdit(key) {
    const member = _all.find(m => m[C.KEY] === key);
    if (!member) return;
    _editingRow = member._rowIndex;
    document.getElementById('member-modal-title').textContent = 'Edit Member';
    _populateForm(member);
    Utils.showModal('member-modal');
  }

  function _populateForm(m) {
    const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val || ''; };
    set('mf-key',    m[C.KEY]);
    set('mf-last',   m[C.LAST]);
    set('mf-first',  m[C.FIRST]);
    set('mf-alt',    m[C.ALT]);
    set('mf-email',  m[C.EMAIL]);
    set('mf-loc',    m[C.LOC]);
    set('mf-status', m[C.STATUS]);
    set('mf-type',   m[C.TYPE]);
    set('mf-fam',    m[C.FAM]);
  }

  // ── Save (Add or Edit) ────────────────────────────────────────────────────
  async function save() {
    const btn = document.getElementById('member-save-btn');
    btn.disabled = true;
    try {
      const get = id => document.getElementById(id)?.value?.trim() || '';
      const key = get('mf-key') || `${get('mf-last').toLowerCase()}|${get('mf-first').toLowerCase()}`.replace(/\s+/g, '-');

      const obj = {
        [C.KEY]:    key,
        [C.LAST]:   get('mf-last'),
        [C.ALT]:    get('mf-alt'),
        [C.FIRST]:  get('mf-first'),
        [C.EMAIL]:  get('mf-email'),
        [C.LOC]:    get('mf-loc'),
        [C.STATUS]: get('mf-status'),
        [C.TYPE]:   get('mf-type'),
        [C.FAM]:    get('mf-fam'),
      };

      if (_editingRow) {
        // Preserve existing columns not in the form
        const existing = _all.find(m => m[C.KEY] === get('mf-key'));
        const merged   = existing ? { ...existing, ...obj } : obj;
        await Sheets.update(CONFIG.SHEETS.MEMBERS, _editingRow, merged);
        Utils.toast('Member updated successfully.');
      } else {
        // Check for duplicate key
        if (_all.some(m => m[C.KEY] === key)) {
          Utils.toast('A member with this key already exists.', 'error');
          btn.disabled = false;
          return;
        }
        await Sheets.append(CONFIG.SHEETS.MEMBERS, obj);
        Utils.toast('Member added successfully.');
      }

      Utils.hideModal('member-modal');
      await render();
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
    }
  }

  // ── Delete ────────────────────────────────────────────────────────────────
  async function confirmDelete(key) {
    const member = _all.find(m => m[C.KEY] === key);
    if (!member) return;
    const ok = await Utils.confirm(`Delete member "${`${member[C.FIRST]} ${member[C.LAST]}`.trim() || key}"? This cannot be undone.`);
    if (!ok) return;
    Utils.setLoading(true, 'Deleting…');
    try {
      await Sheets.deleteRow(CONFIG.SHEETS.MEMBERS, member._rowIndex);
    } catch (e) {
      Utils.toast('Delete failed: ' + e.message, 'error');
      Utils.setLoading(false);
      return;
    }

    // Immediately update UI without waiting for a server round-trip
    _all = _all.filter(m => m[C.KEY] !== key);
    _applyFilter(document.getElementById('member-search')?.value || '');
    _renderTable();
    _updateCount();
    Utils.setLoading(false);
    Utils.toast('Member deleted.');

    // Reload from server in background to sync row indices
    render().catch(() => {});
  }

  // ── Export CSV ────────────────────────────────────────────────────────────
  function exportCSV() {
    const cols   = [C.KEY, C.LAST, C.FIRST, C.EMAIL, C.LOC, C.STATUS, C.TYPE, C.FAM];
    const header = cols.join(',');
    const rows   = _filtered.map(m =>
      cols.map(c => `"${(m[c] || '').replace(/"/g, '""')}"`).join(',')
    );
    const blob   = new Blob([header + '\n' + rows.join('\n')], { type: 'text/csv' });
    const url    = URL.createObjectURL(blob);
    const a      = Object.assign(document.createElement('a'), {
      href: url, download: `swiss-club-members-${Utils.today()}.csv`,
    });
    a.click(); URL.revokeObjectURL(url);
  }

  // ── Init (wire up event listeners, called once) ────────────────────────────
  function init() {
    document.getElementById('member-search')
      ?.addEventListener('input', e => _onSearch(e.target.value));
    document.getElementById('add-member-btn')
      ?.addEventListener('click', openAdd);
    document.getElementById('member-save-btn')
      ?.addEventListener('click', save);
    document.getElementById('member-export-btn')
      ?.addEventListener('click', exportCSV);
    document.getElementById('member-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('member-modal'));
    // Sort headers
    document.querySelectorAll('#members-table th[data-sort]').forEach(th => {
      th.addEventListener('click', () => sort(th.dataset.sort));
    });
  }

  return { render, init, openAdd, openEdit, confirmDelete, exportCSV };
})();
