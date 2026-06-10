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
  let _view       = 'list'; // 'list' | 'families'
  let _detailKey  = null;

  const C = {
    KEY:    'Member Key',
    LAST:   'Last Name',
    ALT:    'Alternative Name',
    FIRST:  'First Name',
    EMAIL:  'Email',
    LOC:    'Location (Metro Manila/Province)',
    STATUS: 'Membership Status',
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
          <div class="member-cell member-cell-link" onclick="Members.openDetail('${Utils.escape(m[C.KEY])}')">
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

    const memberKeys = members.map(m => m[C.KEY]);
    const year       = Utils.currentYear();
    const yearPaid   = _txns
      .filter(t => memberKeys.includes(t['MemberKey']) && String(t['Year']) === String(year) && t['Category'] === 'Membership')
      .reduce((s, t) => s + (parseFloat(t['AmountPaid']) || 0), 0);
    const grp = Utils.escape(groupName);

    return `<div class="family-card">
      <div class="family-card-header">
        <span class="family-card-name">👨‍👩‍👧 ${grp}</span>
        <div style="display:flex;align-items:center;gap:6px;">
          <span class="family-card-count">${members.length} member${members.length !== 1 ? 's' : ''}</span>
          <button class="btn-icon" title="Rename group"
                  onclick="Members.openRenameGroup('${grp}')">✏️</button>
        </div>
      </div>
      <div class="family-card-members">${memberRows}</div>
      <div class="family-card-footer">
        <div class="family-card-paid">
          <span class="family-paid-amount">${Utils.formatPHP(yearPaid)}</span>
          <span class="family-paid-label">${year} membership dues</span>
        </div>
        <button class="btn btn-sm btn-outline"
                onclick="Members.openFamilyStatusModal('${grp}')">
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

  // ── Rename group modal ────────────────────────────────────────────────────
  function openRenameGroup(groupName) {
    document.getElementById('rename-group-old').value            = groupName;
    document.getElementById('rename-group-current-display').textContent = groupName;
    document.getElementById('rename-group-input').value          = groupName;
    Utils.showModal('rename-group-modal');
    setTimeout(() => {
      const inp = document.getElementById('rename-group-input');
      inp.focus();
      inp.select();
    }, 50);
  }

  async function saveRenameGroup() {
    const oldName = document.getElementById('rename-group-old').value.trim();
    const newName = document.getElementById('rename-group-input').value.trim();

    if (!newName)             { Utils.toast('Please enter a new group name.', 'error'); return; }
    if (newName === oldName)  { Utils.hideModal('rename-group-modal'); return; }

    const affected = _all.filter(m => m[C.FAM] === oldName);
    if (!affected.length) { Utils.toast('No members found in this group.', 'error'); return; }

    const ok = await Utils.confirm(
      `Rename group "${oldName}" → "${newName}"?\n${affected.length} member(s) will be updated.`
    );
    if (!ok) return;

    const btn = document.getElementById('rename-group-save-btn');
    btn.disabled = true;
    Utils.setLoading(true, 'Renaming group…');

    try {
      for (const m of affected) {
        await Sheets.update(CONFIG.SHEETS.MEMBERS, m._rowIndex, { ...m, [C.FAM]: newName });
        m[C.FAM] = newName;
      }
      Utils.hideModal('rename-group-modal');
      Utils.toast(`Group renamed to "${newName}"`);
      _renderFamilyGroups();
    } catch (e) {
      Utils.toast('Save failed: ' + e.message, 'error');
    } finally {
      btn.disabled = false;
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

  // ── Member Detail view ────────────────────────────────────────────────────
  function openDetail(key) {
    _detailKey = key;
    document.getElementById('members-page-header').style.display = 'none';
    document.getElementById('members-list-view').style.display   = 'none';
    document.getElementById('members-family-view').style.display = 'none';
    document.getElementById('member-detail-view').style.display  = '';
    _renderDetail(key);
  }

  function closeDetail() {
    _detailKey = null;
    document.getElementById('members-page-header').style.display = '';
    document.getElementById('member-detail-view').style.display  = 'none';
    switchView(_view);
  }

  function _renderDetail(key) {
    const member = _all.find(m => m[C.KEY] === key);
    if (!member) return;

    const memberTxns = _txns
      .filter(t => t['MemberKey'] === key)
      .sort((a, b) => new Date(b['Timestamp']) - new Date(a['Timestamp']));

    const totalPaid  = memberTxns.reduce((s, t) => s + (parseFloat(t['AmountPaid']) || 0), 0);
    const year       = Utils.currentYear();
    const yearPaid   = memberTxns
      .filter(t => String(t['Year']) === String(year))
      .reduce((s, t) => s + (parseFloat(t['AmountPaid']) || 0), 0);
    const fullName   = `${member[C.FIRST]} ${member[C.LAST]}`.trim();

    const txnRows = memberTxns.length
      ? memberTxns.map(t => {
          const date  = t['Timestamp'] ? Utils.formatDate(t['Timestamp']) : '—';
          const desc  = Utils.escape(t['EventName'] || t['Category'] || '—');
          const notes = t['Notes'] ? `<div class="txn-notes">${Utils.escape(t['Notes'])}</div>` : '';
          const tid   = Utils.escape(t['TransactionID']);
          return `<tr>
            <td>${date}</td>
            <td>${desc}${notes}</td>
            <td>${_categoryBadge(t['Category'])}</td>
            <td>${Utils.escape(t['PaymentMode'] || '—')}</td>
            <td class="amount">${Utils.formatPHP(t['AmountPaid'])}</td>
            <td class="actions">
              <button class="btn-icon" title="Edit" onclick="Members.openEditTxn('${tid}')">✏️</button>
              <button class="btn-icon btn-danger" title="Delete" onclick="Members.confirmDeleteTxn('${tid}')">🗑️</button>
            </td>
          </tr>`;
        }).join('')
      : `<tr><td colspan="6" class="empty-state">No transactions recorded yet.</td></tr>`;

    document.getElementById('member-detail-content').innerHTML = `
      <div class="member-detail-header">
        <div class="avatar lg">${Utils.initials(fullName)}</div>
        <div class="member-detail-info">
          <h2 class="member-detail-name">${Utils.escape(fullName)}</h2>
          <div class="member-detail-meta">
            ${[member[C.EMAIL], member[C.LOC]].filter(Boolean).map(Utils.escape).join('<span class="meta-sep">·</span>')}
          </div>
          <div class="member-detail-badges">
            ${Utils.statusBadge(member[C.STATUS])}
            ${Utils.typeBadge(member[C.TYPE])}
            ${member[C.FAM] ? `<span class="badge badge-fam">👨‍👩‍👧 ${Utils.escape(member[C.FAM])}</span>` : ''}
            <span class="badge badge-key">${Utils.escape(member[C.KEY])}</span>
          </div>
        </div>
        <div style="display:flex;gap:8px;flex-shrink:0;">
          <button class="btn btn-primary btn-sm" onclick="Members.openRecordDues('${Utils.escape(key)}')">💰 Record Dues</button>
          <button class="btn btn-outline btn-sm" onclick="Members.openEdit('${Utils.escape(key)}')">✏️ Edit</button>
        </div>
      </div>

      <div class="member-detail-stats">
        <div class="stat-box"><span class="stat-num">${Utils.formatPHP(totalPaid)}</span><span class="stat-label">Total Paid (All Time)</span></div>
        <div class="stat-box"><span class="stat-num">${Utils.formatPHP(yearPaid)}</span><span class="stat-label">${year} Paid</span></div>
        <div class="stat-box"><span class="stat-num">${memberTxns.length}</span><span class="stat-label">Transactions</span></div>
      </div>

      <div class="member-detail-transactions">
        <h3 class="section-title">Transaction History</h3>
        <table class="data-table">
          <thead>
            <tr>
              <th>Date</th><th>Description</th><th>Category</th><th>Mode</th><th>Amount</th><th></th>
            </tr>
          </thead>
          <tbody>${txnRows}</tbody>
        </table>
      </div>`;
  }

  function _categoryBadge(cat) {
    const map = { Membership: 'badge-dues', Event: 'badge-event', RSVP: 'badge-upcoming' };
    return cat ? `<span class="badge ${map[cat] || 'badge-tbc'}">${Utils.escape(cat)}</span>` : '';
  }

  // ── Edit / Delete transaction ─────────────────────────────────────────────
  function openEditTxn(txnId) {
    const t = _txns.find(x => x['TransactionID'] === txnId);
    if (!t) return;
    document.getElementById('txn-id-display').textContent = txnId;
    document.getElementById('txn-id').value               = txnId;
    document.getElementById('txn-event-name').value       = t['EventName']   || '';
    document.getElementById('txn-category').value         = t['Category']    || 'Membership';
    document.getElementById('txn-year').value             = t['Year']        || Utils.currentYear();
    document.getElementById('txn-amount').value           = t['AmountPaid']  || '';
    document.getElementById('txn-mode').value             = t['PaymentMode'] || 'Cash';
    document.getElementById('txn-date').value             = Utils.toISODate(t['Timestamp']);
    document.getElementById('txn-headcount').value        = t['HeadCount']   || 1;
    document.getElementById('txn-notes').value            = t['Notes']       || '';
    Utils.showModal('txn-modal');
  }

  async function saveTxn() {
    const btn   = document.getElementById('txn-save-btn');
    const txnId = document.getElementById('txn-id').value;
    const t     = _txns.find(x => x['TransactionID'] === txnId);
    if (!t) return;

    btn.disabled = true;
    Utils.setLoading(true, 'Saving…');
    try {
      const date    = document.getElementById('txn-date').value;
      const updated = {
        ...t,
        EventName:   document.getElementById('txn-event-name').value.trim(),
        Category:    document.getElementById('txn-category').value,
        Year:        parseInt(document.getElementById('txn-year').value, 10),
        AmountPaid:  parseFloat(document.getElementById('txn-amount').value) || 0,
        PaymentMode: document.getElementById('txn-mode').value,
        Timestamp:   date ? new Date(date + 'T00:00:00').toISOString() : t['Timestamp'],
        HeadCount:   parseInt(document.getElementById('txn-headcount').value, 10) || 1,
        Notes:       document.getElementById('txn-notes').value.trim(),
      };
      await Sheets.update(CONFIG.SHEETS.TRANSACTIONS, t._rowIndex, updated);
      Object.assign(t, updated);
      Utils.hideModal('txn-modal');
      Utils.toast('Transaction updated.');
      if (_detailKey) _renderDetail(_detailKey);
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      btn.disabled = false;
      Utils.setLoading(false);
    }
  }

  async function confirmDeleteTxn(txnId) {
    const t = _txns.find(x => x['TransactionID'] === txnId);
    if (!t) return;
    const desc = t['EventName'] || t['Category'] || txnId;
    const ok   = await Utils.confirm(`Delete transaction "${desc}" (${txnId})? This cannot be undone.`);
    if (!ok) return;
    Utils.setLoading(true, 'Deleting…');
    try {
      await Sheets.deleteRow(CONFIG.SHEETS.TRANSACTIONS, t._rowIndex);
      _txns = _txns.filter(x => x['TransactionID'] !== txnId);
      Utils.toast('Transaction deleted.');
      if (_detailKey) _renderDetail(_detailKey);
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  // ── Record Dues modal ─────────────────────────────────────────────────────
  function openRecordDues(key) {
    const member = _all.find(m => m[C.KEY] === key);
    if (!member) return;
    const fullName = `${member[C.FIRST]} ${member[C.LAST]}`.trim();
    document.getElementById('dues-member-name').textContent        = fullName;
    document.getElementById('dues-member-key-display').textContent = member[C.KEY];
    document.getElementById('dues-member-key').value               = key;
    document.getElementById('dues-year').value                     = Utils.currentYear();
    document.getElementById('dues-amount').value                   = '';
    document.getElementById('dues-mode').value                     = 'Cash';
    document.getElementById('dues-date').value                     = Utils.today();
    document.getElementById('dues-notes').value                    = '';
    document.getElementById('dues-mark-member').checked            = true;
    Utils.showModal('dues-modal');
  }

  async function saveRecordDues() {
    const btn    = document.getElementById('dues-save-btn');
    const key    = document.getElementById('dues-member-key').value;
    const year   = parseInt(document.getElementById('dues-year').value, 10);
    const amount = parseFloat(document.getElementById('dues-amount').value);

    if (!amount || amount <= 0) {
      Utils.toast('Please enter a valid amount.', 'error');
      return;
    }

    const member = _all.find(m => m[C.KEY] === key);
    if (!member) return;

    btn.disabled = true;
    Utils.setLoading(true, 'Recording payment…');
    try {
      const txnId    = await Sheets.nextId(CONFIG.SHEETS.TRANSACTIONS, 'MEM');
      const fullName = `${member[C.FIRST]} ${member[C.LAST]}`.trim();
      const date     = document.getElementById('dues-date').value;
      const mode     = document.getElementById('dues-mode').value;
      const notes    = document.getElementById('dues-notes').value.trim();
      const markMember = document.getElementById('dues-mark-member').checked;

      await Sheets.append(CONFIG.SHEETS.TRANSACTIONS, {
        TransactionID: txnId,
        Timestamp:     date ? new Date(date + 'T00:00:00').toISOString() : new Date().toISOString(),
        MemberKey:     key,
        MemberName:    fullName,
        EventID:       '',
        EventName:     `${year} Membership Dues`,
        AmountPaid:    amount,
        PaymentMode:   mode,
        Category:      'Membership',
        Year:          year,
        HeadCount:     1,
        Notes:         notes,
        RecordedBy:    Auth.getUserEmail(),
      });

      if (markMember) {
        await Sheets.update(CONFIG.SHEETS.MEMBERS, member._rowIndex, { ...member, [C.STATUS]: 'Member' });
        member[C.STATUS] = 'Member';
      }

      Utils.hideModal('dues-modal');
      Utils.toast(`Payment recorded: ${txnId}`);

      _txns = await Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS).catch(() => _txns);
      if (_detailKey === key) _renderDetail(key);
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      btn.disabled = false;
      Utils.setLoading(false);
    }
  }

  // ── Generate next MBR-NNNN key ────────────────────────────────────────────
  function _nextMemberId() {
    const nums = _all
      .map(m => m[C.KEY])
      .filter(k => /^MBR-\d+$/.test(k))
      .map(k => parseInt(k.slice(4), 10));
    const max = nums.length ? Math.max(...nums) : 0;
    return `MBR-${String(max + 1).padStart(4, '0')}`;
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
      const key   = get('mf-key') || _nextMemberId();

      const obj = {
        [C.KEY]:    key,
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
      if (_detailKey) _renderDetail(_detailKey);
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
    document.getElementById('dues-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('dues-modal'));
    document.getElementById('dues-save-btn')
      ?.addEventListener('click', saveRecordDues);
    document.getElementById('txn-modal-close')
      ?.addEventListener('click', () => Utils.hideModal('txn-modal'));
    document.getElementById('txn-save-btn')
      ?.addEventListener('click', saveTxn);
    document.getElementById('rename-group-close')
      ?.addEventListener('click', () => Utils.hideModal('rename-group-modal'));
    document.getElementById('rename-group-save-btn')
      ?.addEventListener('click', saveRenameGroup);
    document.querySelectorAll('#members-table th[data-sort]').forEach(th => {
      th.addEventListener('click', () => sort(th.dataset.sort));
    });
  }

  return {
    render, init, switchView,
    openDetail, closeDetail,
    openAdd, openEdit, confirmDelete, exportCSV,
    openRecordDues,
    openEditTxn, confirmDeleteTxn,
    assignToGroup, removeFromGroup, openFamilyStatusModal, openRenameGroup,
  };
})();
