/**
 * members.js — Member Management view
 * Searchable table with Add / Edit / Delete via modal.
 * Family Groups view: manage groupings and apply bulk status updates.
 */

const Members = (() => {
  let _all      = [];
  let _txns     = [];
  let _filtered = [];
  let _sortKey  = 'Last Name';
  let _sortDir  = 'asc';
  let _editingRow = null;
  let _view     = 'list'; // 'list' | 'families'

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
    NAME:   'Full Name',
  };

  // ── Load data ─────────────────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading members…');
    try {
      [_all, _txns] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.MEMBERS),
        Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS).catch(() => []),
      ]);
      _all = _all.filter(m => m[C.KEY]?.trim());
      if (_view === 'families') {
        _renderFamilyGroups();
      } else {
        _applyFilter(document.getElementById('member-search')?.value || '');
        _renderTable();
        _updateCount();
      }
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  // ── View toggle ───────────────────────────────────────────────────────────
  function switchView(view) {
    _view = view;
    const isList = view === 'list';
    document.getElementById('members-list-view').style.display   = isList ? '' : 'none';
    document.getElementById('members-family-view').style.display = isList ? 'none' : '';
    document.getElementById('members-list-controls').style.display   = isList ? 'flex' : 'none';
    document.getElementById('families-list-controls').style.display  = isList ? 'none'  : 'flex';
    document.querySelectorAll('.view-toggle-btn').forEach(b =>
      b.classList.toggle('active', b.dataset.view === view)
    );
    if (isList) {
      _applyFilter(document.getElementById('member-search')?.value || '');
      _renderTable();
      _updateCount();
    } else {
      _renderFamilyGroups();
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

  // ── Render list table ─────────────────────────────────────────────────────
  function _renderTable() {
    const tbody = document.querySelector('#members-table tbody');
    if (!tbody) return;

    if (!_filtered.length) {
      tbody.innerHTML = '<tr><td colspan="8" class="empty-state">No members found.</td></tr>';
      return;
    }

    tbody.innerHTML = _filtered.map(m => {
      const ytd      = Utils.totalPaidYTD(m[C.KEY], _txns);
      const fullName = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
      const initials = Utils.initials(m[C.FIRST] || m[C.LAST] || '?');
      return `<tr data-key="${Utils.escape(m[C.KEY])}">
        <td>
          <div class="member-cell">
            <div class="avatar">${initials}</div>
            <div>
              <div class="member-name">${Utils.escape(fullName)}</div>
            </div>
          </div>
        </td>
        <td>${Utils.escape(m[C.EMAIL])}</td>
        <td>${Utils.escape(m[C.LOC])}</td>
        <td>${Utils.statusBadge(m[C.STATUS])}</td>
        <td>${Utils.typeBadge(m[C.TYPE])}</td>
        <td>${m[C.FAM] ? `<span class="badge badge-fam" title="${Utils.escape(m[C.FAM])}">👨‍👩‍👧 ${Utils.escape(m[C.FAM])}</span>` : ''}</td>
        <td class="amount">${Utils.formatPHP(ytd)}</td>
        <td class="actions">
          <button class="btn-icon" title="Edit" onclick="Members.openEdit('${Utils.escape(m[C.KEY])}')">✏️</button>
          <button class="btn-icon btn-danger" title="Delete" onclick="Members.confirmDelete('${Utils.escape(m[C.KEY])}')">🗑️</button>
        </td>
      </tr>`;
    }).join('');
  }

  // ── Family Groups view ────────────────────────────────────────────────────
  function _renderFamilyGroups() {
    const grid      = document.getElementById('family-grid');
    const unassSec  = document.getElementById('unassigned-section');
    if (!grid) return;

    // Bucket members by family group
    const groups = {};
    const unassigned = [];
    _all.forEach(m => {
      const grp = m[C.FAM]?.trim();
      if (grp) {
        if (!groups[grp]) groups[grp] = [];
        groups[grp].push(m);
      } else {
        unassigned.push(m);
      }
    });

    const groupNames = Object.keys(groups).sort();

    if (!groupNames.length) {
      grid.innerHTML = '<p class="empty-state" style="grid-column:1/-1">No family groups yet. Edit a member and set their Family Group to get started.</p>';
    } else {
      grid.innerHTML = groupNames.map(name => _familyCardHTML(name, groups[name], groupNames)).join('');
    }

    // Unassigned section
    if (unassSec) {
      unassSec.querySelector('.unassigned-title').textContent = `Unassigned Members (${unassigned.length})`;
      const list = document.getElementById('unassigned-list');
      if (list) {
        if (!unassigned.length) {
          list.innerHTML = '<p class="empty-state">All members are assigned to a family group.</p>';
        } else {
          const opts = groupNames.map(g => `<option value="${Utils.escape(g)}">${Utils.escape(g)}</option>`).join('');
          list.innerHTML = unassigned.map(m => {
            const name = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
            return `<div class="unassigned-row">
              <div class="unassigned-info">
                <div class="avatar sm">${Utils.initials(name)}</div>
                <span>${Utils.escape(name)}</span>
                ${Utils.statusBadge(m[C.STATUS])}
              </div>
              <select class="form-control unassigned-select"
                      data-key="${Utils.escape(m[C.KEY])}"
                      onchange="Members.assignToGroup(this.dataset.key, this.value); this.value=''">
                <option value="">— Assign to group —</option>
                ${opts}
              </select>
            </div>`;
          }).join('');
        }
      }
    }
  }

  function _familyCardHTML(groupName, members, allGroupNames) {
    const memberRows = members.map(m => {
      const name = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
      return `<div class="fam-member-row">
        <div class="fam-member-info">
          <div class="avatar sm">${Utils.initials(name)}</div>
          <span class="fam-member-name">${Utils.escape(name)}</span>
          ${Utils.statusBadge(m[C.STATUS])}
        </div>
        <button class="btn-icon btn-danger fam-remove-btn" title="Remove from group"
                onclick="Members.removeFromGroup('${Utils.escape(m[C.KEY])}')">✕</button>
      </div>`;
    }).join('');

    return `<div class="family-card">
      <div class="family-card-header">
        <span class="family-card-name">👨‍👩‍👧 ${Utils.escape(groupName)}</span>
        <span class="family-card-count">${members.length} member${members.length !== 1 ? 's' : ''}</span>
      </div>
      <div class="family-card-members">${memberRows}</div>
      <div class="family-card-footer">
        <button class="btn btn-sm btn-outline"
                onclick="Members.openFamilyStatusModal('${Utils.escape(groupName)}')">
          Update Status
        </button>
      </div>
    </div>`;
  }

  // ── Assign / remove family group ──────────────────────────────────────────
  async function assignToGroup(memberKey, groupName) {
    if (!groupName) return;
    const member = _all.find(m => m[C.KEY] === memberKey);
    if (!member) return;
    Utils.setLoading(true, 'Saving…');
    try {
      await Sheets.update(CONFIG.SHEETS.MEMBERS, member._rowIndex, { ...member, [C.FAM]: groupName });
      member[C.FAM] = groupName;
      _renderFamilyGroups();
      Utils.toast(`${member[C.FIRST]} ${member[C.LAST]} added to "${groupName}".`);
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  async function removeFromGroup(memberKey) {
    const member = _all.find(m => m[C.KEY] === memberKey);
    if (!member) return;
    const name = `${member[C.FIRST]} ${member[C.LAST]}`.trim();
    const ok = await Utils.confirm(`Remove ${name} from family group "${member[C.FAM]}"?`);
    if (!ok) return;
    Utils.setLoading(true, 'Saving…');
    try {
      await Sheets.update(CONFIG.SHEETS.MEMBERS, member._rowIndex, { ...member, [C.FAM]: '' });
      member[C.FAM] = '';
      _renderFamilyGroups();
      Utils.toast(`${name} removed from family group.`);
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  // ── Family status modal ───────────────────────────────────────────────────
  function openFamilyStatusModal(groupName) {
    const members = _all.filter(m => m[C.FAM] === groupName);
    document.getElementById('fsm-title').textContent = `Update Status — ${groupName}`;
    document.getElementById('fsm-desc').textContent =
      `Sets the 2026 membership status for all ${members.length} member${members.length !== 1 ? 's' : ''} in this group.`;
    document.getElementById('fsm-members-list').innerHTML = members.map(m => {
      const name = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
      return `<div class="fsm-member-row">
        <div class="avatar sm">${Utils.initials(name)}</div>
        <span>${Utils.escape(name)}</span>
        ${Utils.statusBadge(m[C.STATUS])}
      </div>`;
    }).join('');
    document.getElementById('fsm-save-btn').dataset.group = groupName;
    Utils.showModal('family-status-modal');
  }

  async function saveFamilyStatus() {
    const btn       = document.getElementById('fsm-save-btn');
    const groupName = btn.dataset.group;
    const status    = document.getElementById('fsm-status').value;
    const members   = _all.filter(m => m[C.FAM] === groupName);

    btn.disabled = true;
    Utils.setLoading(true, 'Updating family…');
    try {
      for (const m of members) {
        await Sheets.update(CONFIG.SHEETS.MEMBERS, m._rowIndex, { ...m, [C.STATUS]: status });
        m[C.STATUS] = status;
      }
      Utils.hideModal('family-status-modal');
      _renderFamilyGroups();
      if (_view === 'list') _renderTable();
      Utils.toast(`Updated ${members.length} members to "${status}".`);
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      btn.disabled = false;
      Utils.setLoading(false);
    }
  }

  // ── Search (debounced) ────────────────────────────────────────────────────
  const _onSearch = Utils.debounce(query => {
    _applyFilter(query);
    _renderTable();
    _updateCount();
  });

  // ── Sort ──────────────────────────────────────────────────────────────────
  function sort(key) {
    if (_sortKey === key) _sortDir = _sortDir === 'asc' ? 'desc' : 'asc';
    else { _sortKey = key; _sortDir = 'asc'; }
    _applyFilter(document.getElementById('member-search')?.value || '');
    _renderTable();
    document.querySelectorAll('#members-table th[data-sort]').forEach(th => {
      th.dataset.dir = th.dataset.sort === key ? _sortDir : '';
    });
  }

  // ── Datalist helper ───────────────────────────────────────────────────────
  function _populateFamilyDatalist() {
    const dl = document.getElementById('family-groups-list');
    if (!dl) return;
    const groups = [...new Set(_all.map(m => m[C.FAM]).filter(Boolean))].sort();
    dl.innerHTML = groups.map(g => `<option value="${Utils.escape(g)}">`).join('');
  }

  // ── Add modal ─────────────────────────────────────────────────────────────
  function openAdd() {
    _editingRow = null;
    document.getElementById('member-modal-title').textContent = 'Add Member';
    document.getElementById('member-form').reset();
    document.getElementById('mf-key').value = '';
    _populateFamilyDatalist();
    Utils.showModal('member-modal');
  }

  // ── Edit modal ────────────────────────────────────────────────────────────
  function openEdit(key) {
    const member = _all.find(m => m[C.KEY] === key);
    if (!member) return;
    _editingRow = member._rowIndex;
    document.getElementById('member-modal-title').textContent = 'Edit Member';
    _populateFamilyDatalist();
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

  // ── Save ──────────────────────────────────────────────────────────────────
  async function save() {
    const btn = document.getElementById('member-save-btn');
    btn.disabled = true;
    try {
      const get = id => document.getElementById(id)?.value?.trim() || '';
      const first = get('mf-first');
      const last  = get('mf-last');
      const key   = get('mf-key') || `${last.toLowerCase()}|${first.toLowerCase()}`.replace(/\s+/g, '-');

      const obj = {
        [C.KEY]:    key,
        [C.NAME]:   `${first} ${last}`.trim().toUpperCase(),
        [C.LAST]:   last,
        [C.ALT]:    get('mf-alt'),
        [C.FIRST]:  first,
        [C.EMAIL]:  get('mf-email'),
        [C.LOC]:    get('mf-loc'),
        [C.STATUS]: get('mf-status'),
        [C.TYPE]:   get('mf-type'),
        [C.FAM]:    get('mf-fam'),
      };

      if (_editingRow) {
        const existing = _all.find(m => m[C.KEY] === get('mf-key'));
        const merged   = existing ? { ...existing, ...obj } : obj;
        await Sheets.update(CONFIG.SHEETS.MEMBERS, _editingRow, merged);
        Utils.toast('Member updated.');
      } else {
        if (_all.some(m => m[C.KEY] === key)) {
          Utils.toast('A member with this key already exists.', 'error');
          btn.disabled = false;
          return;
        }
        await Sheets.append(CONFIG.SHEETS.MEMBERS, obj);
        Utils.toast('Member added.');
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
    _all = _all.filter(m => m[C.KEY] !== key);
    _applyFilter(document.getElementById('member-search')?.value || '');
    _renderTable();
    _updateCount();
    Utils.setLoading(false);
    Utils.toast('Member deleted.');
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

  // ── Init ──────────────────────────────────────────────────────────────────
  function init() {
    document.getElementById('member-search')
      ?.addEventListener('input', e => _onSearch(e.target.value));
    document.getElementById('add-member-btn')
      ?.addEventListener('click', openAdd);
    document.getElementById('add-member-btn-fam')
      ?.addEventListener('click', openAdd);
    document.getElementById('member-save-btn')
      ?.addEventListener('click', save);
    document.getElementById('member-export-btn')
      ?.addEventListener('click', exportCSV);
    document.getElementById('member-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('member-modal'));
    document.getElementById('fsm-save-btn')
      ?.addEventListener('click', saveFamilyStatus);
    document.getElementById('fsm-close')
      ?.addEventListener('click', () => Utils.hideModal('family-status-modal'));
    document.querySelectorAll('#members-table th[data-sort]').forEach(th => {
      th.addEventListener('click', () => sort(th.dataset.sort));
    });
  }

  return {
    render, init, switchView,
    openAdd, openEdit, confirmDelete, exportCSV,
    assignToGroup, removeFromGroup, openFamilyStatusModal,
  };
})();
