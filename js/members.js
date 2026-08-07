/**
 * members.js — Member Management view
 * Searchable table with Add / Edit / Delete via modal.
 * Family Groups view: manage groupings and apply bulk status updates.
 */

const Members = (() => {
  let _all      = [];
  let _txns     = [];
  let _events   = [];
  let _filtered = [];
  let _sortKey  = 'Last Name';
  let _sortDir  = 'asc';
  let _editingRow = null;
  let _view       = 'list'; // 'list' | 'families'
  let _detailKey  = null;
  let _afterSaveCallback = null; // optional cross-module callback set by openAdd(prefill, cb)
  let _rates      = {}; // loaded from Rates sheet, falls back to MEMBERSHIP_RATES
  let _regs       = []; // lazy-loaded when member detail is opened

  // Merge state
  let _mergeCanonicalKey = null;
  let _mergeDupKey       = null;
  let _mergeStep         = 1;
  let _mergeConflicts    = []; // [{field, keepVal, dupVal, chosen}]
  let _mergeResolved     = {}; // field -> chosen value

  const C = {
    KEY:      'Member Key',
    LAST:     'Last Name',
    ALT:      'Alternative Name',
    FIRST:    'First Name',
    EMAIL:    'Email',
    MOBILE:   'Mobile',
    LOC:      'Location (Metro Manila/Province)',
    STATUS:   'Membership Status',
    RENEWAL:  'Renewal Year',
    TYPE:     'Membership Type',
    FAM:      'Family Group',   // legacy column kept for search
    FAM_HEAD: 'Family Head',    // member key of the family head
    DATE_ADDED: 'Date Added',
  };

  // Annual membership rates — keyed by tier id
  const MEMBERSHIP_RATES = {
    'single-jr':   { label: 'Single Jr. (18–25)',           amount: 1500 },
    'mm-single':   { label: 'Metro Manila – Individual',     amount: 2800 },
    'mm-family':   { label: 'Metro Manila – Family',         amount: 3500 },
    'prov-single': { label: 'Province/Overseas – Individual', amount: 2000 },
    'prov-family': { label: 'Province/Overseas – Family',    amount: 2300 },
  };

  function _suggestTier(member) {
    const type = (member[C.TYPE] || '').toLowerCase();
    const loc  = member[C.LOC]  || '';
    if (type === 'single')     return 'single-jr';
    const isMM     = loc === 'Metro Manila';
    if (type === 'family')     return isMM ? 'mm-family'  : 'prov-family';
    if (type === 'individual') return isMM ? 'mm-single'  : 'prov-single';
    return '';
  }

  async function _loadRates() {
    try {
      const rows = await Sheets.getAll(CONFIG.SHEETS.RATES);
      if (!rows.length) return { ...MEMBERSHIP_RATES };
      const loaded = {};
      rows.forEach(r => {
        const id     = r['Tier ID']?.trim();
        const label  = r['Label']?.trim();
        const amount = parseFloat(r['Amount']);
        if (id && label && !isNaN(amount)) loaded[id] = { label, amount };
      });
      return Object.keys(loaded).length ? loaded : { ...MEMBERSHIP_RATES };
    } catch {
      return { ...MEMBERSHIP_RATES };
    }
  }

  // ── Load data ─────────────────────────────────────────────────────────────
  async function render() {
    Utils.setLoading(true, 'Loading members…');
    try {
      [_all, _txns, _rates] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.MEMBERS),
        Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS).catch(() => []),
        _loadRates(),
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
    const statusFilter = document.getElementById('member-status-filter')?.value || '';
    _filtered = Utils.filterRows(_all, query, [C.FIRST, C.LAST, C.ALT, C.EMAIL, C.FAM, C.KEY]);
    if (statusFilter) _filtered = _filtered.filter(m => m[C.STATUS] === statusFilter);
    if (_sortKey === C.STATUS) {
      // Compound sort: Status (toggleable) then Renewal Year descending within same status
      _filtered = [..._filtered].sort((a, b) => {
        const sa  = (a[C.STATUS] || '').toLowerCase();
        const sb  = (b[C.STATUS] || '').toLowerCase();
        const cmp = sa.localeCompare(sb);
        if (cmp !== 0) return _sortDir === 'asc' ? cmp : -cmp;
        const ya  = parseInt(a[C.RENEWAL], 10) || 0;
        const yb  = parseInt(b[C.RENEWAL], 10) || 0;
        return yb - ya; // most recent year first within same status
      });
    } else {
      _filtered = Utils.sortTable(_filtered, _sortKey, _sortDir);
    }
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
      tbody.innerHTML = '<tr><td colspan="10" class="empty-state">No members found.</td></tr>';
      return;
    }

    // Pre-build headKey → name map for the family column
    const headNames = {};
    _all.forEach(m => {
      if (m[C.FAM_HEAD] === m[C.KEY]) headNames[m[C.KEY]] = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
    });

    tbody.innerHTML = _filtered.map(m => {
      const ytd     = Utils.totalPaidYTD(m[C.KEY], _txns);
      const key     = Utils.escape(m[C.KEY]);
      const isHead  = m[C.FAM_HEAD] === m[C.KEY];
      const famName = headNames[m[C.FAM_HEAD]] || '';
      const famCell = isHead
        ? `<span class="badge badge-head">👑 Head</span>`
        : famName ? `<span class="badge badge-fam">👨‍👩‍👧 ${Utils.escape(famName)}</span>` : '';
      return `<tr data-key="${key}">
        <td class="col-key" style="font-family:monospace;font-size:12px;">${Utils.escape(key)}</td>
        <td class="member-cell-link col-first" onclick="Members.openDetail('${key}')">${Utils.escape(m[C.FIRST])}</td>
        <td class="member-cell-link col-last" onclick="Members.openDetail('${key}')">${Utils.escape(m[C.LAST])}</td>
        <td class="col-email">${Utils.escape(m[C.EMAIL])}</td>
        <td class="col-location">${Utils.escape(m[C.LOC])}</td>
        <td class="col-status">${Utils.statusBadge(m[C.STATUS], m[C.RENEWAL])}</td>
        <td class="col-type">${Utils.typeBadge(m[C.TYPE])}</td>
        <td class="col-family">${famCell}</td>
        <td class="amount col-ytd">${Utils.formatPHP(ytd)}</td>
        <td class="actions">
          <button class="btn-icon btn-history" title="Transaction history" onclick="Members.openDetail('${key}')">🧾</button>
          <button class="btn-icon" title="Edit" onclick="Members.openEdit('${key}')">✏️</button>
          <button class="btn-icon btn-danger" title="Delete" onclick="Members.confirmDelete('${key}')">🗑️</button>
        </td>
      </tr>`;
    }).join('');
  }

  // ── Family Groups view ────────────────────────────────────────────────────
  function _renderFamilyGroups() {
    const grid     = document.getElementById('family-grid');
    const unassSec = document.getElementById('unassigned-section');
    if (!grid) return;

    // Group by Family Head key
    const groups = {};
    const unassigned = [];
    _all.forEach(m => {
      const hk = m[C.FAM_HEAD]?.trim();
      if (hk) {
        if (!groups[hk]) groups[hk] = [];
        groups[hk].push(m);
      } else {
        unassigned.push(m);
      }
    });

    // Sort families by head's last name
    const headKeys = Object.keys(groups).sort((a, b) => {
      const ha = _all.find(m => m[C.KEY] === a);
      const hb = _all.find(m => m[C.KEY] === b);
      const na = ha ? `${ha[C.LAST]} ${ha[C.FIRST]}` : a;
      const nb = hb ? `${hb[C.LAST]} ${hb[C.FIRST]}` : b;
      return na.localeCompare(nb);
    });

    if (!headKeys.length) {
      grid.innerHTML = '<p class="empty-state" style="grid-column:1/-1">No family groups yet. Select an unassigned member below and click "Start New Family".</p>';
    } else {
      grid.innerHTML = headKeys.map(hk => _familyCardHTML(hk, groups[hk])).join('');
    }

    // Unassigned section
    if (unassSec) {
      unassSec.querySelector('.unassigned-title').textContent = `Unassigned Members (${unassigned.length})`;
      const list = document.getElementById('unassigned-list');
      if (list) {
        if (!unassigned.length) {
          list.innerHTML = '<p class="empty-state">All members are assigned to a family.</p>';
        } else {
          const headOpts = headKeys.map(hk => {
            const head = _all.find(m => m[C.KEY] === hk);
            const name = head ? `${head[C.FIRST]} ${head[C.LAST]}`.trim() : hk;
            return `<option value="${Utils.escape(hk)}">${Utils.escape(name)}'s family</option>`;
          }).join('');

          list.innerHTML = unassigned.map(m => {
            const name = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
            const key  = Utils.escape(m[C.KEY]);
            return `<div class="unassigned-row">
              <div class="unassigned-info">
                <span>${Utils.escape(name)}</span>
                ${Utils.statusBadge(m[C.STATUS], m[C.RENEWAL])}
              </div>
              <div class="unassigned-actions">
                ${headOpts ? `<select class="form-control unassigned-select"
                        data-key="${key}"
                        onchange="Members.assignToFamily(this.dataset.key, this.value); this.value=''">
                  <option value="">— Add to existing family —</option>
                  ${headOpts}
                </select>` : ''}
                <button class="btn btn-sm btn-outline"
                        onclick="Members.createFamily('${key}')">
                  Start New Family
                </button>
              </div>
            </div>`;
          }).join('');
        }
      }
    }
  }

  function _familyCardHTML(headKey, members) {
    const hk = Utils.escape(headKey);

    // Head selector options — used in the card header
    const headOpts = members.map(m => {
      const name = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
      const sel  = m[C.KEY] === headKey ? 'selected' : '';
      return `<option value="${Utils.escape(m[C.KEY])}" ${sel}>${Utils.escape(name)}</option>`;
    }).join('');

    const memberRows = members.map(m => {
      const name   = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
      const isHead = m[C.KEY] === headKey;
      const key    = Utils.escape(m[C.KEY]);
      return `<div class="fam-member-row">
        <div class="fam-member-info">
          <span class="fam-member-name">${Utils.escape(name)}</span>
          ${Utils.statusBadge(m[C.STATUS], m[C.RENEWAL])}
        </div>
        ${!isHead
          ? `<button class="btn-icon btn-danger fam-remove-btn" title="Remove from family"
                     onclick="Members.removeFromFamily('${key}')">✕</button>`
          : '<span class="fam-remove-placeholder"></span>'}
      </div>`;
    }).join('');

    const memberKeys = members.map(m => m[C.KEY]);
    const year       = Utils.currentYear();
    const yearPaid   = _txns
      .filter(t => memberKeys.includes(t['MemberKey']) && String(t['Year']) === String(year) && t['Category'] === 'Membership')
      .reduce((s, t) => s + (parseFloat(t['AmountPaid']) || 0), 0);

    return `<div class="family-card">
      <div class="family-card-header">
        <div class="family-head-select-wrap">
          <span class="family-card-icon">👨‍👩‍👧</span>
          <select class="fam-head-select"
                  title="Change family head"
                  data-head="${hk}"
                  onchange="Members.setAsHead(this.value, this)">
            ${headOpts}
          </select>
        </div>
        <span class="family-card-count">${members.length} member${members.length !== 1 ? 's' : ''}</span>
      </div>
      <div class="family-card-members">${memberRows}</div>
      <div class="family-card-footer">
        <div class="family-card-paid">
          <span class="family-paid-amount">${Utils.formatPHP(yearPaid)}</span>
          <span class="family-paid-label">${year} membership dues</span>
        </div>
        <button class="btn btn-sm btn-outline"
                onclick="Members.openFamilyStatusModal('${hk}')">
          Update Status
        </button>
      </div>
    </div>`;
  }

  // ── Family assignment / head management ───────────────────────────────────
  async function assignToFamily(memberKey, headKey) {
    if (!headKey) return;
    const member = _all.find(m => m[C.KEY] === memberKey);
    if (!member) return;
    Utils.setLoading(true, 'Saving…');
    try {
      await Sheets.update(CONFIG.SHEETS.MEMBERS, member._rowIndex, { ...member, [C.FAM_HEAD]: headKey });
      member[C.FAM_HEAD] = headKey;
      _renderFamilyGroups();
      const head = _all.find(m => m[C.KEY] === headKey);
      const headName = head ? `${head[C.FIRST]} ${head[C.LAST]}`.trim() : headKey;
      Utils.toast(`${member[C.FIRST]} ${member[C.LAST]} added to ${headName}'s family.`);
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  async function removeFromFamily(memberKey) {
    const member = _all.find(m => m[C.KEY] === memberKey);
    if (!member) return;
    const name   = `${member[C.FIRST]} ${member[C.LAST]}`.trim();
    const isHead = member[C.FAM_HEAD] === memberKey;

    if (isHead) {
      const others = _all.filter(m => m[C.FAM_HEAD] === memberKey && m[C.KEY] !== memberKey);
      if (others.length) {
        Utils.toast(`${name} is the family head — set a new head first before removing them.`, 'error');
        return;
      }
    }

    const ok = await Utils.confirm(`Remove ${name} from their family?`);
    if (!ok) return;
    Utils.setLoading(true, 'Removing…');
    try {
      await Sheets.update(CONFIG.SHEETS.MEMBERS, member._rowIndex, { ...member, [C.FAM_HEAD]: '' });
      member[C.FAM_HEAD] = '';
      _renderFamilyGroups();
      Utils.toast(`${name} removed from family.`);
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  async function setAsHead(memberKey, selectEl) {
    const member = _all.find(m => m[C.KEY] === memberKey);
    if (!member) return;
    const oldHeadKey = member[C.FAM_HEAD];
    if (!oldHeadKey || oldHeadKey === memberKey) {
      if (selectEl) selectEl.value = oldHeadKey || memberKey;
      return;
    }

    const name = `${member[C.FIRST]} ${member[C.LAST]}`.trim();
    const ok   = await Utils.confirm(`Set ${name} as the new family head?`);
    if (!ok) {
      if (selectEl) selectEl.value = oldHeadKey;
      return;
    }

    const familyMembers = _all.filter(m => m[C.FAM_HEAD] === oldHeadKey);
    Utils.setLoading(true, 'Updating family head…');
    try {
      for (const m of familyMembers) {
        await Sheets.update(CONFIG.SHEETS.MEMBERS, m._rowIndex, { ...m, [C.FAM_HEAD]: memberKey });
        m[C.FAM_HEAD] = memberKey;
      }
      Utils.toast(`${name} is now the family head.`);
      _renderFamilyGroups();
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  async function createFamily(memberKey) {
    const member = _all.find(m => m[C.KEY] === memberKey);
    if (!member) return;
    const name = `${member[C.FIRST]} ${member[C.LAST]}`.trim();
    const ok   = await Utils.confirm(`Create a new family with ${name} as the head?`);
    if (!ok) return;
    Utils.setLoading(true, 'Creating family…');
    try {
      await Sheets.update(CONFIG.SHEETS.MEMBERS, member._rowIndex, { ...member, [C.FAM_HEAD]: memberKey });
      member[C.FAM_HEAD] = memberKey;
      _renderFamilyGroups();
      Utils.toast(`${name}'s family created.`);
    } catch (e) {
      Utils.toast('Error: ' + e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  // ── Family status modal ───────────────────────────────────────────────────
  function openFamilyStatusModal(headKey) {
    const members  = _all.filter(m => m[C.FAM_HEAD] === headKey);
    const head     = _all.find(m => m[C.KEY] === headKey);
    const headName = head ? `${head[C.FIRST]} ${head[C.LAST]}`.trim() : headKey;
    document.getElementById('fsm-title').textContent = `Update Status — ${headName}'s family`;
    document.getElementById('fsm-desc').textContent =
      `Sets the membership status for all ${members.length} member${members.length !== 1 ? 's' : ''} in this family.`;
    document.getElementById('fsm-members-list').innerHTML = members.map(m => {
      const name = `${m[C.FIRST]} ${m[C.LAST]}`.trim();
      return `<div class="fsm-member-row">
        <div class="avatar sm">${Utils.initials(name)}</div>
        <span>${Utils.escape(name)}</span>
        ${Utils.statusBadge(m[C.STATUS], m[C.RENEWAL])}
      </div>`;
    }).join('');
    document.getElementById('fsm-save-btn').dataset.group = headKey;
    Utils.showModal('family-status-modal');
  }

  async function saveFamilyStatus() {
    const btn     = document.getElementById('fsm-save-btn');
    const headKey = btn.dataset.group;
    const status  = document.getElementById('fsm-status').value;
    const members = _all.filter(m => m[C.FAM_HEAD] === headKey);

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

  function applyStatusFilter() {
    _applyFilter(document.getElementById('member-search')?.value || '');
    _renderTable();
    _updateCount();
  }

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
  // ── Member Detail view ────────────────────────────────────────────────────
  async function openDetail(key) {
    _detailKey = key;
    document.getElementById('members-page-header').style.display = 'none';
    document.getElementById('members-list-view').style.display   = 'none';
    document.getElementById('members-family-view').style.display = 'none';
    document.getElementById('member-detail-view').style.display  = '';
    _renderDetail(key);
    if (!_regs.length) {
      _regs = await Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).catch(() => []);
      if (_detailKey === key) _renderDetail(key);
    }
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
          const hc = parseInt(t['HeadCount'], 10) || 1;
          return `<tr>
            <td>${date}</td>
            <td>${desc}${notes}</td>
            <td>${_categoryBadge(t['Category'])}</td>
            <td style="text-align:center;">${Utils.escape(t['Year'] || '—')}</td>
            <td style="text-align:center;">${hc}</td>
            <td>${Utils.escape(t['PaymentMode'] || '—')}</td>
            <td class="amount">${Utils.formatPHP(t['AmountPaid'])}</td>
            <td class="actions">
              <button class="btn-icon" title="Edit" onclick="Members.openEditTxn('${tid}')">✏️</button>
              <button class="btn-icon btn-danger" title="Delete" onclick="Members.confirmDeleteTxn('${tid}')">🗑️</button>
            </td>
          </tr>`;
        }).join('')
      : `<tr><td colspan="8" class="empty-state">No transactions recorded yet.</td></tr>`;

    document.getElementById('member-detail-content').innerHTML = `
      <div class="member-detail-header">
        <div class="avatar lg">${Utils.initials(fullName)}</div>
        <div class="member-detail-info">
          <h2 class="member-detail-name">${Utils.escape(fullName)}</h2>
          <div class="member-detail-meta">
            ${[member[C.EMAIL], member[C.MOBILE], member[C.LOC]].filter(Boolean).map(Utils.escape).join('<span class="meta-sep">·</span>')}
          </div>
          <div class="member-detail-badges">
            ${Utils.statusBadge(member[C.STATUS], member[C.RENEWAL])}
            ${Utils.typeBadge(member[C.TYPE])}
            ${(() => { const hd = member[C.FAM_HEAD] ? _all.find(m => m[C.KEY] === member[C.FAM_HEAD]) : null; return hd ? `<span class="badge badge-fam">👨‍👩‍👧 ${Utils.escape(`${hd[C.FIRST]} ${hd[C.LAST]}`.trim())}'s family</span>` : ''; })()}
            <span class="badge badge-key">${Utils.escape(member[C.KEY])}</span>
          </div>
        </div>
        <div style="display:flex;gap:8px;flex-shrink:0;">
          <button class="btn btn-primary btn-sm" onclick="Members.openRecordDues('${Utils.escape(key)}')">💰 Record Transaction</button>
          <button class="btn btn-outline btn-sm" onclick="Members.openEdit('${Utils.escape(key)}')">✏️ Edit</button>
          <button class="btn btn-outline btn-sm" onclick="Members.openMerge('${Utils.escape(key)}')">⇌ Merge Duplicate</button>
        </div>
      </div>

      <div class="member-detail-stats">
        <div class="stat-box"><span class="stat-num">${Utils.formatPHP(totalPaid)}</span><span class="stat-label">Total Paid (All Time)</span></div>
        <div class="stat-box"><span class="stat-num">${Utils.formatPHP(yearPaid)}</span><span class="stat-label">${year} Paid</span></div>
        <div class="stat-box"><span class="stat-num">${memberTxns.length}</span><span class="stat-label">Transactions</span></div>
        <div class="stat-box"><span class="stat-num">${_regs.length ? _eventsAttended(key).length : '…'}</span><span class="stat-label">Events Attended</span></div>
      </div>

      <div class="member-detail-transactions">
        <h3 class="section-title">Transaction History</h3>
        <table class="data-table">
          <thead>
            <tr>
              <th>Date</th><th>Description</th><th>Category</th><th>Year</th><th>Pax</th><th>Mode</th><th>Amount</th><th></th>
            </tr>
          </thead>
          <tbody>${txnRows}</tbody>
        </table>
      </div>

      <div class="member-detail-events">
        <h3 class="section-title">Events Attended</h3>
        ${_regs.length ? _renderEventsAttended(key) : '<p class="empty-state" style="font-size:13px;padding:12px 0;">Loading…</p>'}
      </div>`;
  }

  function _categoryBadge(cat) {
    const map = { Membership: 'badge-dues', Event: 'badge-event', RSVP: 'badge-upcoming' };
    return cat ? `<span class="badge ${map[cat] || 'badge-tbc'}">${Utils.escape(cat)}</span>` : '';
  }

  function _renderEventsAttended(key) {
    const attended = _eventsAttended(key);
    if (!attended.length) return '<p class="empty-state" style="font-size:13px;padding:12px 0;">No events attended yet.</p>';

    return `<div class="events-attended-list">${attended.map(({ r, event, isPrimary, primaryName }) => {
      const eventName = event?.Title || r.EventName || r.EventID || '—';
      const date      = event?.Date  ? Utils.formatDate(event.Date) : '—';
      const location  = event?.Location || '';
      const regId     = r.RegistrationID || '';
      const mQty      = parseInt(r.MemberQty,  10) || 0;
      const gQty      = parseInt(r.GuestQty,   10) || 0;
      const kQty      = parseInt(r.KidsQty,    10) || 0;
      const paxParts  = [
        mQty ? `${mQty}M` : '', gQty ? `${gQty}G` : '', kQty ? `${kQty}K` : '',
      ].filter(Boolean).join('+');
      const roleTag   = isPrimary
        ? `<span class="ea-role ea-role-primary">Primary registrant</span>`
        : `<span class="ea-role ea-role-guest">On ${Utils.escape(primaryName)}'s registration</span>`;

      return `<div class="ea-row">
        <div class="ea-date">${Utils.escape(date)}</div>
        <div class="ea-info">
          <div class="ea-name">${Utils.escape(eventName)}</div>
          <div class="ea-meta">
            ${location ? `<span>${Utils.escape(location)}</span><span class="meta-sep">·</span>` : ''}
            ${paxParts ? `<span>${paxParts}</span><span class="meta-sep">·</span>` : ''}
            ${regId ? `<span class="ea-regid">${Utils.escape(regId)}</span>` : ''}
          </div>
        </div>
        <div class="ea-tags">${roleTag}</div>
      </div>`;
    }).join('')}</div>`;
  }

  function _eventsAttended(key) {
    // Returns registrations where this member was checked in (as primary or slot)
    return _regs
      .filter(r => {
        const checked = new Set((r.CheckedIn || '').split(',').map(s => s.trim()).filter(Boolean));
        if (!checked.size) return false;
        // Primary registrant — any slot checked in counts (they may have registered
        // with guest tickets only, so m0 may not exist)
        if (r.MemberKey === key) return true;
        // Additional member slots
        const slots = (r.MemberSlots || '').split(',').map(s => s.trim()).filter(Boolean);
        const slotIdx = slots.indexOf(key);
        if (slotIdx !== -1 && checked.has(`m${slotIdx + 1}`)) return true;
        return false;
      })
      .map(r => {
        const event   = _events.find(e => e.EventID === r.EventID);
        const date    = event?.Date || r.EventID || '';
        const isPrimary = r.MemberKey === key;
        const primaryName = isPrimary ? null
          : [r.LastName, r.FirstName].filter(Boolean).join(', ') || r.MemberKey;
        return { r, event, date, isPrimary, primaryName };
      })
      .sort((a, b) => new Date(b.date) - new Date(a.date));
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
        Month:       date ? parseInt(date.split('-')[1], 10) : (parseInt(t['Month'], 10) || new Date().getMonth() + 1),
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

  // ── Record Payment modal ──────────────────────────────────────────────────
  async function openRecordDues(key) {
    const member = _all.find(m => m[C.KEY] === key);
    if (!member) return;
    const fullName = `${member[C.FIRST]} ${member[C.LAST]}`.trim();

    document.getElementById('dues-member-name').textContent        = fullName;
    document.getElementById('dues-member-key-display').textContent = member[C.KEY];
    document.getElementById('dues-member-key').value               = key;
    document.getElementById('dues-category').value                 = 'Membership';
    document.getElementById('dues-year').value                     = Utils.currentYear();
    document.getElementById('dues-mode').value                     = 'Cash';
    document.getElementById('dues-date').value                     = Utils.today();
    document.getElementById('dues-notes').value                    = '';
    document.getElementById('dues-headcount').value                = '1';
    document.getElementById('dues-description').value              = '';
    document.getElementById('dues-mark-member').checked            = true;

    // Build tier options from live rates (may differ from hardcoded defaults)
    const tierSel = document.getElementById('dues-tier');
    tierSel.innerHTML = '<option value="">— Select tier —</option>' +
      Object.entries(_rates).map(([id, r]) =>
        `<option value="${Utils.escape(id)}">${Utils.escape(r.label)} — ₱${r.amount.toLocaleString()}</option>`
      ).join('');

    // Auto-suggest tier and pre-fill amount
    const suggestedTier = _suggestTier(member);
    tierSel.value = _rates[suggestedTier] ? suggestedTier : '';
    const rate = _rates[suggestedTier];
    document.getElementById('dues-amount').value = rate ? rate.amount : '';

    onDuesCategoryChange();

    // Populate event dropdown (lazy-load once)
    if (!_events.length) {
      _events = await Sheets.getAll(CONFIG.SHEETS.EVENTS).catch(() => []);
    }
    const sel = document.getElementById('dues-event-select');
    sel.innerHTML = '<option value="">— Select event —</option>' +
      _events.map(e => `<option value="${Utils.escape(e.EventID)}" data-name="${Utils.escape(e.Title)}">
        ${Utils.escape(e.Title)} (${Utils.formatDate(e.Date)})
      </option>`).join('');

    Utils.showModal('dues-modal');
  }

  function onDuesCategoryChange() {
    const cat = document.getElementById('dues-category').value;
    document.getElementById('dues-membership-fields').style.display  = cat === 'Membership' ? '' : 'none';
    document.getElementById('dues-event-fields').style.display       = cat === 'Event'      ? '' : 'none';
    document.getElementById('dues-other-fields').style.display       = cat === 'Other'      ? '' : 'none';
    document.getElementById('dues-mark-member-row').style.display    = cat === 'Membership' ? '' : 'none';
  }

  function onDuesTierChange() {
    const tier = document.getElementById('dues-tier').value;
    const rate = _rates[tier];
    if (rate) document.getElementById('dues-amount').value = rate.amount;
  }

  async function saveRecordDues() {
    const btn      = document.getElementById('dues-save-btn');
    const key      = document.getElementById('dues-member-key').value;
    const category = document.getElementById('dues-category').value;
    const amount   = parseFloat(document.getElementById('dues-amount').value);

    if (!amount || amount <= 0) { Utils.toast('Please enter a valid amount.', 'error'); return; }

    const member = _all.find(m => m[C.KEY] === key);
    if (!member) return;

    // Category-specific validation
    let eventId = '', eventName = '', headCount = 1, txnPrefix = 'TXN';
    if (category === 'Membership') {
      txnPrefix = 'MEM';
      const year = parseInt(document.getElementById('dues-year').value, 10);
      if (!year) { Utils.toast('Please enter a membership year.', 'error'); return; }
      const tier    = document.getElementById('dues-tier').value;
      const tierObj = _rates[tier];
      eventName = tierObj ? `${year} Membership Dues – ${tierObj.label}` : `${year} Membership Dues`;
    } else if (category === 'Event') {
      const sel = document.getElementById('dues-event-select');
      eventId   = sel.value;
      eventName = sel.options[sel.selectedIndex]?.dataset?.name || '';
      headCount = parseInt(document.getElementById('dues-headcount').value, 10) || 1;
      if (!eventId) { Utils.toast('Please select an event.', 'error'); return; }
      txnPrefix = 'TXN';
    } else {
      eventName = document.getElementById('dues-description').value.trim();
      if (!eventName) { Utils.toast('Please enter a description.', 'error'); return; }
    }

    btn.disabled = true;
    Utils.setLoading(true, 'Recording payment…');
    try {
      const txnId    = await Sheets.nextId(CONFIG.SHEETS.TRANSACTIONS, txnPrefix);
      const fullName = `${member[C.FIRST]} ${member[C.LAST]}`.trim();
      const date     = document.getElementById('dues-date').value;
      const mode     = document.getElementById('dues-mode').value;
      const notes    = document.getElementById('dues-notes').value.trim();
      const dateObj  = date ? new Date(date + 'T00:00:00') : new Date();

      await Sheets.append(CONFIG.SHEETS.TRANSACTIONS, {
        TransactionID: txnId,
        Timestamp:     dateObj.toISOString(),
        MemberKey:     key,
        MemberName:    fullName,
        EventID:       eventId,
        EventName:     eventName,
        AmountPaid:    amount,
        PaymentMode:   mode,
        Category:      category,
        Year:          dateObj.getFullYear(),
        Month:         date ? parseInt(date.split('-')[1], 10) : dateObj.getMonth() + 1,
        HeadCount:     headCount,
        Notes:         notes,
        RecordedBy:    Auth.getUserEmail(),
      });

      if (category === 'Membership' && document.getElementById('dues-mark-member').checked) {
        const renewalYear = parseInt(document.getElementById('dues-year').value, 10) || dateObj.getFullYear();
        await Sheets.update(CONFIG.SHEETS.MEMBERS, member._rowIndex, {
          ...member,
          [C.STATUS]:  'Member',
          [C.RENEWAL]: renewalYear,
        });
        member[C.STATUS]  = 'Member';
        member[C.RENEWAL] = renewalYear;
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
  function onTypeChange() {
    const type   = document.getElementById('mf-type')?.value || '';
    const famRow = document.getElementById('mf-fam-head')?.closest('.form-group');
    if (famRow) famRow.style.opacity = type === 'Family' ? '1' : '0.4';
    const sel = document.getElementById('mf-fam-head');
    if (sel && sel.dataset.lockedHead !== 'true') sel.disabled = type !== 'Family';
  }

  function _buildFamilySelect(memberKey, currentHeadKey) {
    const sel  = document.getElementById('mf-fam-head');
    const hint = document.getElementById('mf-fam-hint');
    if (!sel) return;

    const isHead = memberKey && currentHeadKey === memberKey;

    if (isHead) {
      sel.innerHTML = `<option value="${Utils.escape(memberKey)}">👑 Head of own family</option>`;
      sel.disabled  = true;
      sel.dataset.lockedHead = 'true';
      if (hint) { hint.textContent = 'To reassign the head, use the Family Groups view.'; hint.style.display = ''; }
    } else {
      sel.disabled = false;
      sel.dataset.lockedHead = 'false';
      if (hint) hint.style.display = 'none';
      const heads = _all
        .filter(m => m[C.FAM_HEAD] === m[C.KEY] && m[C.KEY] !== memberKey)
        .sort((a, b) => a[C.LAST].localeCompare(b[C.LAST]) || a[C.FIRST].localeCompare(b[C.FIRST]));
      const selfOption = memberKey
        ? `<option value="${Utils.escape(memberKey)}" ${currentHeadKey === memberKey ? 'selected' : ''}>➕ Start new family (make me head)</option>`
        : '';
      sel.innerHTML = '<option value="">— No family —</option>' +
        selfOption +
        heads.map(h => {
          const name     = `${h[C.FIRST]} ${h[C.LAST]}`.trim();
          const selected = currentHeadKey === h[C.KEY] ? 'selected' : '';
          return `<option value="${Utils.escape(h[C.KEY])}" ${selected}>${Utils.escape(name)}'s family</option>`;
        }).join('');
    }
  }

  function openAdd(prefill, afterSaveCallback) {
    _editingRow = null;
    _afterSaveCallback = afterSaveCallback || null;
    document.getElementById('member-modal-title').textContent = 'Add Member';
    document.getElementById('member-form').reset();
    document.getElementById('mf-key').value = '';
    _buildFamilySelect('', '');
    onTypeChange();
    if (prefill) {
      const set = (id, val) => { if (val) { const el = document.getElementById(id); if (el) el.value = val; } };
      set('mf-key',    prefill.key);
      set('mf-first',  prefill.firstName);
      set('mf-last',   prefill.lastName);
      set('mf-email',  prefill.email);
      set('mf-status', prefill.status);
    }
    Utils.showModal('member-modal');
  }

  // ── Edit modal ────────────────────────────────────────────────────────────
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
    set('mf-mobile', m[C.MOBILE]);
    set('mf-loc',    m[C.LOC]);
    set('mf-status', m[C.STATUS]);
    set('mf-type',   m[C.TYPE]);
    _buildFamilySelect(m[C.KEY], m[C.FAM_HEAD]);
    onTypeChange();
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

      const existing = _editingRow ? _all.find(m => m[C.KEY] === get('mf-key')) : null;
      const obj = {
        [C.KEY]:     key,
        [C.LAST]:    last,
        [C.ALT]:     get('mf-alt'),
        [C.FIRST]:   first,
        [C.EMAIL]:   get('mf-email'),
        [C.MOBILE]:  get('mf-mobile'),
        [C.LOC]:     get('mf-loc'),
        [C.STATUS]:  get('mf-status'),
        [C.RENEWAL]: existing?.[C.RENEWAL] || '',
        [C.TYPE]:     get('mf-type'),
        [C.FAM_HEAD]: get('mf-fam-head'),
        // FAM (legacy) preserved from existing via merge below
      };

      if (_editingRow) {
        // Guard against accidentally reassigning a key that belongs to a different row
        const conflict = _all.find(m => m[C.KEY] === key && m._rowIndex !== _editingRow);
        if (conflict) {
          const cName = `${conflict[C.FIRST] || ''} ${conflict[C.LAST] || ''}`.trim() || key;
          Utils.toast(`Member key "${key}" is already used by ${cName}. Choose a different key.`, 'error');
          btn.disabled = false;
          return;
        }
        const merged = existing ? { ...existing, ...obj } : obj;
        await Sheets.update(CONFIG.SHEETS.MEMBERS, _editingRow, merged);
        Utils.toast('Member updated.');
      } else {
        if (_all.some(m => m[C.KEY] === key)) {
          Utils.toast('A member with this key already exists.', 'error');
          btn.disabled = false;
          return;
        }
        obj[C.DATE_ADDED] = new Date().toISOString().slice(0, 10);
        Sheets.clearHeaderCache(CONFIG.SHEETS.MEMBERS);
        await Sheets.append(CONFIG.SHEETS.MEMBERS, obj);
        Utils.toast('Member added.');
      }

      Utils.hideModal('member-modal');
      await render();
      if (_detailKey) _renderDetail(_detailKey);

      if (!_editingRow && _afterSaveCallback) {
        const cb = _afterSaveCallback;
        _afterSaveCallback = null;
        await cb(key);
      }
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
  // ── Column visibility ─────────────────────────────────────────────────────
  const _ALL_COLS = ['first','last','email','location','status','type','family','ytd'];
  let _hiddenCols = new Set(JSON.parse(localStorage.getItem('members-hidden-cols') || '[]'));

  function _applyColVisibility() {
    let el = document.getElementById('members-col-style');
    if (!el) {
      el = document.createElement('style');
      el.id = 'members-col-style';
      document.head.appendChild(el);
    }
    el.textContent = [..._hiddenCols]
      .map(k => `#members-table .col-${k} { display: none; }`)
      .join('\n');
    _ALL_COLS.forEach(k => {
      const cb = document.getElementById(`col-toggle-${k}`);
      if (cb) cb.checked = !_hiddenCols.has(k);
    });
  }

  function toggleCol(key) {
    if (_hiddenCols.has(key)) _hiddenCols.delete(key);
    else _hiddenCols.add(key);
    localStorage.setItem('members-hidden-cols', JSON.stringify([..._hiddenCols]));
    _applyColVisibility();
  }

  function toggleColPanel() {
    const panel = document.getElementById('members-col-panel');
    if (!panel) return;
    const open = panel.style.display !== 'none';
    panel.style.display = open ? 'none' : 'block';
    if (!open) {
      // Close on outside click
      setTimeout(() => {
        const close = e => {
          if (!panel.contains(e.target) && !e.target.closest('[onclick*="toggleColPanel"]')) {
            panel.style.display = 'none';
          }
          document.removeEventListener('click', close);
        };
        document.addEventListener('click', close);
      }, 0);
    }
  }

  // ── Merge Duplicate ───────────────────────────────────────────────────────

  const MERGE_FIELDS = [
    { field: C.LAST,    label: 'Last Name' },
    { field: C.FIRST,   label: 'First Name' },
    { field: C.ALT,     label: 'Alternative Name' },
    { field: C.EMAIL,   label: 'Email' },
    { field: C.MOBILE,  label: 'Mobile' },
    { field: C.LOC,     label: 'Location' },
    { field: C.STATUS,  label: 'Membership Status' },
    { field: C.RENEWAL, label: 'Renewal Year' },
    { field: C.TYPE,    label: 'Membership Type' },
    { field: C.FAM_HEAD,label: 'Family Head' },
  ];

  function _memberCard(m) {
    const name = `${m[C.FIRST] || ''} ${m[C.LAST] || ''}`.trim() || m[C.KEY];
    const lines = [
      `<div class="merge-card-name">${Utils.escape(name)}</div>`,
      `<div class="merge-card-key">${Utils.escape(m[C.KEY])}</div>`,
      m[C.EMAIL]  ? `<div class="merge-card-detail">${Utils.escape(m[C.EMAIL])}</div>` : '',
      m[C.STATUS] ? Utils.statusBadge(m[C.STATUS], m[C.RENEWAL]) : '',
      m[C.TYPE]   ? Utils.typeBadge(m[C.TYPE]) : '',
    ].filter(Boolean);
    return lines.join('');
  }

  function openMerge(canonicalKey) {
    _mergeCanonicalKey = canonicalKey;
    _mergeDupKey       = null;
    _mergeStep         = 1;
    _mergeConflicts    = [];
    _mergeResolved     = {};

    const canonical = _all.find(m => m[C.KEY] === canonicalKey);
    if (!canonical) return;

    document.getElementById('merge-step1').style.display = '';
    document.getElementById('merge-step2').style.display = 'none';
    document.getElementById('merge-step3').style.display = 'none';
    document.getElementById('merge-search-input').value = '';
    document.getElementById('merge-search-sugg').innerHTML = '';
    document.getElementById('merge-selected-preview').style.display = 'none';
    document.getElementById('merge-next-btn').style.display = 'none';
    document.getElementById('merge-confirm-btn').style.display = 'none';

    document.getElementById('merge-keep-card-body').innerHTML = _memberCard(canonical);

    Utils.showModal('merge-modal');
  }

  function onMergeSearch(q) {
    const sugg = document.getElementById('merge-search-sugg');
    if (!q || q.length < 2) { sugg.innerHTML = ''; return; }
    const results = Utils.filterRows(_all, q, [C.FIRST, C.LAST, C.ALT, C.EMAIL, C.KEY])
      .filter(m => m[C.KEY] !== _mergeCanonicalKey)
      .slice(0, 6);
    if (!results.length) { sugg.innerHTML = '<div class="suggestion-item">No matches</div>'; return; }
    sugg.innerHTML = results.map(m => {
      const name = `${m[C.FIRST] || ''} ${m[C.LAST] || ''}`.trim() || m[C.KEY];
      return `<div class="suggestion-item" onclick="Members.selectMergeDup('${Utils.escape(m[C.KEY])}')">
        <span class="sug-name">${Utils.escape(name)}</span>
        <span class="sug-key">${Utils.escape(m[C.KEY])}</span>
      </div>`;
    }).join('');
  }

  function selectMergeDup(dupKey) {
    _mergeDupKey = dupKey;
    document.getElementById('merge-search-sugg').innerHTML = '';

    const dup = _all.find(m => m[C.KEY] === dupKey);
    if (!dup) return;

    const input = document.getElementById('merge-search-input');
    input.value = `${dup[C.FIRST] || ''} ${dup[C.LAST] || ''}`.trim() || dupKey;

    document.getElementById('merge-dup-card-body').innerHTML = _memberCard(dup);
    document.getElementById('merge-selected-preview').style.display = '';
    document.getElementById('merge-next-btn').style.display = '';
  }

  function mergeNext() {
    if (_mergeStep === 1) {
      if (!_mergeDupKey) return;
      const canonical = _all.find(m => m[C.KEY] === _mergeCanonicalKey);
      const dup       = _all.find(m => m[C.KEY] === _mergeDupKey);
      if (!canonical || !dup) return;

      // Find conflicts
      _mergeConflicts = MERGE_FIELDS.filter(f => {
        const cv = (canonical[f.field] || '').trim();
        const dv = (dup[f.field]       || '').trim();
        return cv && dv && cv !== dv;
      }).map(f => ({
        field:   f.field,
        label:   f.label,
        keepVal: (canonical[f.field] || '').trim(),
        dupVal:  (dup[f.field]       || '').trim(),
        chosen:  'keep',
      }));

      // Pre-fill resolved with auto-merge values
      _mergeResolved = {};
      MERGE_FIELDS.forEach(f => {
        const cv = (canonical[f.field] || '').trim();
        const dv = (dup[f.field]       || '').trim();
        // Conflict: user must pick — start with canonical
        if (cv && dv && cv !== dv) { _mergeResolved[f.field] = cv; return; }
        // No conflict: take whichever is non-empty
        _mergeResolved[f.field] = cv || dv;
      });

      if (_mergeConflicts.length) {
        _mergeStep = 2;
        _renderConflicts();
        document.getElementById('merge-step1').style.display = 'none';
        document.getElementById('merge-step2').style.display = '';
        document.getElementById('merge-next-btn').textContent = 'Review Summary';
      } else {
        _mergeStep = 3;
        _renderSummary();
        document.getElementById('merge-step1').style.display = 'none';
        document.getElementById('merge-step3').style.display = '';
        document.getElementById('merge-next-btn').style.display = 'none';
        document.getElementById('merge-confirm-btn').style.display = '';
      }
    } else if (_mergeStep === 2) {
      // Collect conflict choices
      document.querySelectorAll('.merge-conflict-row').forEach(row => {
        const field  = row.dataset.field;
        const chosen = row.querySelector('.merge-choice-btn.active')?.dataset.choice;
        if (field && chosen) {
          _mergeResolved[field] = chosen === 'keep'
            ? _mergeConflicts.find(c => c.field === field)?.keepVal
            : _mergeConflicts.find(c => c.field === field)?.dupVal;
        }
      });
      _mergeStep = 3;
      _renderSummary();
      document.getElementById('merge-step2').style.display = 'none';
      document.getElementById('merge-step3').style.display = '';
      document.getElementById('merge-next-btn').style.display = 'none';
      document.getElementById('merge-confirm-btn').style.display = '';
    }
  }

  function _renderConflicts() {
    const container = document.getElementById('merge-conflicts');
    container.innerHTML = _mergeConflicts.map(c => `
      <div class="merge-conflict-row" data-field="${Utils.escape(c.field)}">
        <div class="merge-conflict-label">${Utils.escape(c.label)}</div>
        <div class="merge-conflict-choices">
          <button class="merge-choice-btn active" data-choice="keep"
            onclick="Members._mergePickChoice(this, 'keep')">
            <span class="merge-choice-source">Keep</span>
            <span class="merge-choice-val">${Utils.escape(c.keepVal)}</span>
          </button>
          <button class="merge-choice-btn" data-choice="dup"
            onclick="Members._mergePickChoice(this, 'dup')">
            <span class="merge-choice-source">Duplicate</span>
            <span class="merge-choice-val">${Utils.escape(c.dupVal)}</span>
          </button>
        </div>
      </div>`).join('');
  }

  function _mergePickChoice(btn, choice) {
    const row = btn.closest('.merge-conflict-row');
    row.querySelectorAll('.merge-choice-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  }

  function _renderSummary() {
    const dup    = _all.find(m => m[C.KEY] === _mergeDupKey);
    const dupKey = _mergeDupKey;

    const txnCount  = _txns.filter(t => t.MemberKey === dupKey).length;
    const regs      = [];
    // We'll count these after we load registrations — for now show what we know
    const container = document.getElementById('merge-summary');
    container.innerHTML = `
      <div class="merge-summary-section">
        <div class="merge-summary-title">Merged member record</div>
        ${MERGE_FIELDS.map(f => {
          const val = _mergeResolved[f.field];
          if (!val) return '';
          return `<div class="merge-summary-row"><span class="merge-summary-field">${Utils.escape(f.label)}</span><span class="merge-summary-val">${Utils.escape(val)}</span></div>`;
        }).filter(Boolean).join('')}
      </div>
      <div class="merge-summary-section">
        <div class="merge-summary-title">References to re-link from <code>${Utils.escape(dupKey)}</code></div>
        <div class="merge-summary-row"><span class="merge-summary-field">Transactions</span><span class="merge-summary-val">${txnCount} record${txnCount !== 1 ? 's' : ''}</span></div>
        <div class="merge-summary-row"><span class="merge-summary-field">Duplicate member row</span><span class="merge-summary-val">Will be deleted</span></div>
        <div class="merge-summary-row merge-summary-warn"><span class="merge-summary-field">Registrations + Family Head refs</span><span class="merge-summary-val">Scanned &amp; updated on confirm</span></div>
      </div>`;
  }

  async function executeMerge() {
    const canonical = _all.find(m => m[C.KEY] === _mergeCanonicalKey);
    const dup       = _all.find(m => m[C.KEY] === _mergeDupKey);
    if (!canonical || !dup) return;

    const btn = document.getElementById('merge-confirm-btn');
    btn.disabled = true;
    Utils.setLoading(true, 'Merging…');

    try {
      const canonicalKey = _mergeCanonicalKey;
      const dupKey       = _mergeDupKey;

      // 1. Update the canonical member record with resolved values
      const updatedMember = { ...canonical };
      MERGE_FIELDS.forEach(f => {
        if (_mergeResolved[f.field] !== undefined) updatedMember[f.field] = _mergeResolved[f.field];
      });
      await Sheets.update(CONFIG.SHEETS.MEMBERS, canonical._rowIndex, updatedMember);

      // 2. Re-link Transactions
      const allTxns = await Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS).catch(() => []);
      for (const t of allTxns) {
        if (t.MemberKey === dupKey) {
          await Sheets.update(CONFIG.SHEETS.TRANSACTIONS, t._rowIndex, { ...t, MemberKey: canonicalKey });
        }
      }

      // 3. Re-link Registrations (MemberKey + MemberSlots)
      const allRegs = await Sheets.getAll(CONFIG.SHEETS.REGISTRATIONS).catch(() => []);
      for (const r of allRegs) {
        let changed = false;
        const updated = { ...r };
        if (r.MemberKey === dupKey) {
          updated.MemberKey = canonicalKey;
          changed = true;
        }
        if (r.MemberSlots) {
          const slots = r.MemberSlots.split(',').map(s => s.trim());
          const newSlots = slots.map(s => s === dupKey ? canonicalKey : s);
          if (newSlots.join(',') !== slots.join(',')) {
            updated.MemberSlots = newSlots.join(',');
            changed = true;
          }
        }
        if (changed) await Sheets.update(CONFIG.SHEETS.REGISTRATIONS, r._rowIndex, updated);
      }

      // 4. Re-link Family Head references on other members
      const allMembers = await Sheets.getAll(CONFIG.SHEETS.MEMBERS).catch(() => []);
      for (const m of allMembers) {
        if (m[C.FAM_HEAD] === dupKey && m[C.KEY] !== dupKey) {
          await Sheets.update(CONFIG.SHEETS.MEMBERS, m._rowIndex, { ...m, [C.FAM_HEAD]: canonicalKey });
        }
      }

      // 5. Delete the duplicate row
      await Sheets.deleteRow(CONFIG.SHEETS.MEMBERS, dup._rowIndex);

      Utils.hideModal('merge-modal');
      Utils.toast('Members merged successfully.');
      await render();
      if (_detailKey === dupKey) {
        _detailKey = canonicalKey;
      }
      if (_detailKey) _renderDetail(_detailKey);

    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      btn.disabled = false;
      Utils.setLoading(false);
    }
  }

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
    document.querySelectorAll('#members-table th[data-sort]').forEach(th => {
      th.addEventListener('click', () => sort(th.dataset.sort));
    });
    _applyColVisibility();
  }

  return {
    render, init, switchView,
    openDetail, closeDetail,
    openAdd, openEdit, confirmDelete, exportCSV, onTypeChange,
    openRecordDues, onDuesCategoryChange, onDuesTierChange,
    openEditTxn, confirmDeleteTxn,
    assignToFamily, removeFromFamily, setAsHead, createFamily, openFamilyStatusModal,
    applyStatusFilter, toggleCol, toggleColPanel,
    openMerge, onMergeSearch, selectMergeDup, mergeNext, _mergePickChoice, executeMerge,
  };
})();
