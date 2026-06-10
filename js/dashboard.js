/**
 * dashboard.js — Dashboard view
 * Shows KPI cards, revenue chart, membership breakdown, upcoming events, recent transactions.
 */

const Dashboard = (() => {
  let _charts = {};

  async function render() {
    Utils.setLoading(true, 'Loading dashboard…');
    try {
      const [members, events, txns] = await Promise.all([
        Sheets.getAll(CONFIG.SHEETS.MEMBERS),
        Sheets.getAll(CONFIG.SHEETS.EVENTS),
        Sheets.getAll(CONFIG.SHEETS.TRANSACTIONS),
      ]);
      _drawKPIs(members, events, txns);
      _drawRevenueChart(txns);
      _drawMembershipChart(members);
      _drawUpcomingEvents(events);
      _drawRecentTxns(txns, members);
    } catch (e) {
      Utils.toast(e.message, 'error');
    } finally {
      Utils.setLoading(false);
    }
  }

  // ── KPI Cards ─────────────────────────────────────────────────────────────
  function _drawKPIs(members, events, txns) {
    const year       = Utils.currentYear();
    const realMembers = members.filter(m => m['Member Key'] && m['Member Key'] !== 'Member Key');
    const active     = realMembers.filter(m => {
      const s = m['Membership Status'] || '';
      return s === 'Member' || s === 'Exempt';
    }).length;
    const totalMembers = realMembers.length;

    const yearTxns   = txns.filter(t => String(t.Year || t.Timestamp || '').includes(String(year)));
    const duesRev    = yearTxns.filter(t => t.Category === 'Membership').reduce((s,t) => s + Utils.parsePHP(t.AmountPaid), 0);
    const eventRev   = yearTxns.filter(t => t.Category === 'Event').reduce((s,t) => s + Utils.parsePHP(t.AmountPaid), 0);
    const totalRev   = duesRev + eventRev;
    const upcoming   = events.filter(e => e.Status === 'Upcoming' || new Date(e.Date) >= new Date()).length;

    _setKPI('kpi-total-members',  totalMembers);
    _setKPI('kpi-active-members', active);
    _setKPI('kpi-total-revenue',  Utils.formatPHP(totalRev));
    _setKPI('kpi-upcoming-events', upcoming);
    _setKPI('kpi-dues-revenue',   Utils.formatPHP(duesRev));
    _setKPI('kpi-event-revenue',  Utils.formatPHP(eventRev));
  }

  function _setKPI(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  }

  // ── Monthly Revenue Chart ─────────────────────────────────────────────────
  function _drawRevenueChart(txns) {
    const year   = Utils.currentYear();
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const dues   = Array(12).fill(0);
    const events = Array(12).fill(0);

    txns.forEach(t => {
      const ts = t.Timestamp || '';
      if (!ts.includes(String(year))) return;
      const month = new Date(ts).getMonth();
      if (isNaN(month)) return;
      const amount = Utils.parsePHP(t.AmountPaid);
      if (t.Category === 'Membership') dues[month]   += amount;
      if (t.Category === 'Event')      events[month] += amount;
    });

    const ctx = document.getElementById('revenue-chart');
    if (!ctx) return;
    if (_charts.revenue) _charts.revenue.destroy();

    _charts.revenue = new Chart(ctx, {
      type: 'bar',
      data: {
        labels:   months,
        datasets: [
          { label: 'Membership Dues', data: dues,   backgroundColor: '#CC0000' },
          { label: 'Event Revenue',   data: events, backgroundColor: '#1a3a5c' },
        ],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'top' } },
        scales: {
          x: { stacked: true },
          y: { stacked: true, ticks: { callback: v => '₱' + v.toLocaleString() } },
        },
      },
    });
  }

  // ── Membership Distribution Chart ─────────────────────────────────────────
  function _drawMembershipChart(members) {
    const realMembers = members.filter(m => m['Member Key'] && m['Member Key'] !== 'Member Key');
    const counts = { Member: 0, Exempt: 0, 'Non-member': 0, TBC: 0 };
    realMembers.forEach(m => {
      const s = m['Membership Status'] || 'TBC';
      const key = s.includes('Non') ? 'Non-member' : (counts[s] !== undefined ? s : 'TBC');
      counts[key] = (counts[key] || 0) + 1;
    });

    const ctx = document.getElementById('membership-chart');
    if (!ctx) return;
    if (_charts.membership) _charts.membership.destroy();

    _charts.membership = new Chart(ctx, {
      type: 'doughnut',
      data: {
        labels:   Object.keys(counts),
        datasets: [{
          data: Object.values(counts),
          backgroundColor: ['#16a34a','#2563eb','#9ca3af','#eab308'],
          borderWidth: 2,
        }],
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'right' } },
        cutout: '65%',
      },
    });
  }

  // ── Upcoming Events ───────────────────────────────────────────────────────
  function _drawUpcomingEvents(events) {
    const container = document.getElementById('upcoming-events-list');
    if (!container) return;
    const upcoming = events
      .filter(e => new Date(e.Date) >= new Date() || e.Status === 'Upcoming')
      .sort((a, b) => new Date(a.Date) - new Date(b.Date))
      .slice(0, 4);

    if (!upcoming.length) {
      container.innerHTML = '<p class="empty-state">No upcoming events.</p>';
      return;
    }
    container.innerHTML = upcoming.map(e => `
      <div class="event-item" onclick="Router.navigate('events')">
        <div class="event-date">${Utils.formatDate(e.Date)}</div>
        <div class="event-info">
          <strong>${Utils.escape(e.Title)}</strong>
          <span>${Utils.escape(e.Location)}</span>
        </div>
        <div class="event-fee">${Utils.formatPHP(e.MemberFee)}</div>
      </div>`).join('');
  }

  // ── Recent Transactions ───────────────────────────────────────────────────
  function _drawRecentTxns(txns, members) {
    const container = document.getElementById('recent-txns-list');
    if (!container) return;
    const recent = [...txns]
      .sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp))
      .slice(0, 8);

    if (!recent.length) {
      container.innerHTML = '<p class="empty-state">No transactions yet.</p>';
      return;
    }
    container.innerHTML = `
      <table class="data-table compact">
        <thead><tr>
          <th>Date</th><th>Member</th><th>Category</th><th>Amount</th>
        </tr></thead>
        <tbody>
          ${recent.map(t => `<tr>
            <td>${Utils.formatDate(t.Timestamp)}</td>
            <td>${Utils.escape(t.MemberName || t.MemberKey)}</td>
            <td><span class="badge badge-${(t.Category || '').toLowerCase()}">${Utils.escape(t.Category)}</span></td>
            <td class="amount">${Utils.formatPHP(t.AmountPaid)}</td>
          </tr>`).join('')}
        </tbody>
      </table>`;
  }

  function init() {} // no listeners needed — render() is called by router

  return { render, init };
})();
