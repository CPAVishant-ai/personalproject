/* ============================================================
   RUPAIYA - Personal Expense Tracker | app.js
   ============================================================ */

'use strict';

/* ============================================================ STATE */
let state = {
  expenses: [],
  categories: ['Food & Dining', 'Transport', 'Shopping', 'Entertainment', 'Health & Medical', 'Utilities', 'Rent & Housing', 'Education', 'Travel', 'Personal Care', 'Gifts & Donations', 'Business', 'Investments', 'Other'],
  nature: ['Personal', 'Business', 'Family', 'Travel', 'Emergency'],
  paidBy: ['Cash', 'UPI', 'Credit Card', 'Debit Card', 'Net Banking', 'Wallet', 'Cheque'],
  budgets: {},
  savedCharts: [],
  sortKey: 'date',
  sortDir: 'desc',
};

let currentUserId = null;  // set after Firebase auth resolves

let importData_raw = null;  // holds parsed Excel data during import
let importHeaders = [];
let chartInstances = {};    // track chart.js instances
let selectedIds = new Set(); // tracks selected expense IDs for bulk operations

/* ============================================================ CONSTANTS */
const CATEGORY_COLORS = [
  '#6c63ff','#10b981','#f59e0b','#3b82f6','#ef4444',
  '#8b5cf6','#06b6d4','#84cc16','#f97316','#ec4899',
  '#14b8a6','#a855f7','#22c55e','#eab308',
];

const CATEGORY_EMOJIS = {
  'Food & Dining': '🍽️', 'Transport': '🚗', 'Shopping': '🛍️',
  'Entertainment': '🎬', 'Health & Medical': '💊', 'Utilities': '💡',
  'Rent & Housing': '🏠', 'Education': '📚', 'Travel': '✈️',
  'Personal Care': '💄', 'Gifts & Donations': '🎁', 'Business': '💼',
  'Investments': '📈', 'Other': '📦',
};

const PALETTES = {
  vibrant: ['#6c63ff','#ef4444','#10b981','#f59e0b','#3b82f6','#ec4899','#8b5cf6','#06b6d4','#84cc16','#f97316'],
  pastel:  ['#c4b5fd','#fca5a5','#6ee7b7','#fde68a','#93c5fd','#f9a8d4','#ddd6fe','#67e8f9','#d9f99d','#fed7aa'],
  warm:    ['#ef4444','#f97316','#f59e0b','#eab308','#84cc16','#fb923c','#fbbf24','#facc15','#a3e635','#4ade80'],
  cool:    ['#3b82f6','#06b6d4','#8b5cf6','#6366f1','#14b8a6','#0ea5e9','#7c3aed','#4f46e5','#0891b2','#0d9488'],
  mono:    ['#1e293b','#334155','#475569','#64748b','#94a3b8','#cbd5e1','#e2e8f0','#f1f5f9','#f8fafc','#020617'],
};

const DOW_LABELS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

/* ============================================================ INIT */
document.addEventListener('DOMContentLoaded', () => {
  initFontSize(); // apply saved font size immediately
});

// Firebase resolves auth state and dispatches this event (see firebase.js)
document.addEventListener('firebase:authstate', async (e) => {
  const user = e.detail;
  document.getElementById('loadingOverlay').style.display = 'none';

  if (user) {
    currentUserId = user.uid;

    // Show user info in sidebar footer
    const avatar = document.getElementById('userAvatar');
    const initials = document.getElementById('userInitials');
    if (user.photoURL) {
      avatar.src = user.photoURL;
      avatar.style.display = 'block';
      if (initials) initials.style.display = 'none';
    } else if (initials) {
      initials.textContent = (user.displayName || user.email || 'U')[0].toUpperCase();
    }
    document.getElementById('userName').textContent = user.displayName || user.email || 'User';

    // Show app shell, hide login
    document.getElementById('loginOverlay').style.display = 'none';
    document.getElementById('appShell').style.display = '';

    await loadFromFirestore();
    initApp();
  } else {
    currentUserId = null;
    document.getElementById('loginOverlay').style.display = 'flex';
    document.getElementById('appShell').style.display = 'none';
  }
});

async function loadFromFirestore() {
  try {
    const { settings, expenses } = await window._fb.loadUserData(currentUserId);
    state.expenses = expenses;
    if (settings.categories && settings.categories.length) state.categories = settings.categories;
    if (settings.nature    && settings.nature.length)    state.nature    = settings.nature;
    if (settings.paidBy    && settings.paidBy.length)    state.paidBy    = settings.paidBy;
    if (settings.budgets)     state.budgets     = settings.budgets;
    if (settings.savedCharts) state.savedCharts = settings.savedCharts;
  } catch (err) {
    console.error('Firestore load error:', err);
    showToast('Failed to load data. Check your connection.', 'error');
  }
}

function initApp() {
  initSidebar();
  populateSelects();
  populateGlobalMonthFilter();
  navigate('dashboard');
  setDefaultDate();
  initDropZone();
  renderSidebarStats();

  // Set current date in header
  const dateEl = document.getElementById('currentDate');
  if (dateEl) {
    const now = new Date();
    dateEl.textContent = now.toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
  }

  document.getElementById('globalMonthFilter').addEventListener('change', () => {
    refreshCurrentPage();
  });
}

// saveState now persists only settings (categories, nature, paidBy, budgets, savedCharts).
// Individual expense add/delete/update go through their own Firestore calls.
function saveState() {
  if (!currentUserId) return;
  window._fb.saveSettings(currentUserId, {
    categories:  state.categories,
    nature:      state.nature,
    paidBy:      state.paidBy,
    budgets:     state.budgets,
    savedCharts: state.savedCharts,
  }).catch(e => console.error('Firestore settings save error:', e));
}

/* ============================================================ SIDEBAR STATS */
function renderSidebarStats() {
  const now = new Date();
  const curMonth = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const monthExp = state.expenses.filter(e => e.date.startsWith(curMonth));
  const monthTotal = monthExp.reduce((s, e) => s + e.amount, 0);

  const sbTotal = document.getElementById('sbMonthTotal');
  const sbCount = document.getElementById('sbMonthCount');
  if (sbTotal) sbTotal.textContent = formatINR(monthTotal);
  if (sbCount) sbCount.textContent = `${monthExp.length} expense${monthExp.length !== 1 ? 's' : ''}`;

  renderSidebarDonut();
}

function renderSidebarDonut() {
  const canvas = document.getElementById('sidebarDonut');
  if (!canvas) return;

  const now = new Date();
  const curMonth = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const filtered = state.expenses.filter(e => e.date.startsWith(curMonth));

  destroyChart('sidebarDonut');

  if (!filtered.length) {
    const legend = document.getElementById('sbLegend');
    if (legend) legend.innerHTML = '';
    return;
  }

  const byCat = groupBy(filtered, 'category');
  const entries = Object.entries(byCat)
    .map(([cat, exps]) => ({ cat, total: exps.reduce((s, e) => s + e.amount, 0) }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 5);
  const total = entries.reduce((s, e) => s + e.total, 0);

  chartInstances['sidebarDonut'] = new Chart(canvas.getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: entries.map(e => e.cat),
      datasets: [{
        data: entries.map(e => e.total),
        backgroundColor: entries.map(e => getCategoryColor(e.cat)),
        borderWidth: 0,
        hoverOffset: 4,
      }]
    },
    options: {
      responsive: false,
      plugins: { legend: { display: false }, tooltip: { enabled: false } },
      cutout: '72%',
    }
  });

  const legend = document.getElementById('sbLegend');
  if (legend) {
    legend.innerHTML = entries.slice(0, 4).map(e => `
      <div class="sb-legend-item">
        <div class="sb-legend-dot" style="background:${getCategoryColor(e.cat)}"></div>
        <span class="sb-legend-cat">${e.cat}</span>
        <span class="sb-legend-pct">${((e.total / total) * 100).toFixed(0)}%</span>
      </div>
    `).join('');
  }
}

/* ============================================================ NAVIGATION */
function navigate(page) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.querySelectorAll('.bnav-item').forEach(n => n.classList.remove('active'));

  const pageEl = document.getElementById('page-' + page);
  if (pageEl) pageEl.classList.add('active');

  const navEl = document.querySelector(`.nav-item[data-page="${page}"]`);
  if (navEl) navEl.classList.add('active');

  const bnavEl = document.querySelector(`.bnav-item[data-page="${page}"]`);
  if (bnavEl) bnavEl.classList.add('active');

  const titles = {
    dashboard: 'Dashboard', expenses: 'Expenses', add: 'Add Expense',
    import: 'Import Excel', analytics: 'Analytics',
    'custom-charts': 'Custom Charts', settings: 'Settings',
  };
  document.getElementById('pageTitle').textContent = titles[page] || page;

  // Close mobile sidebar on navigation
  closeMobileSidebar();

  refreshCurrentPage(page);
  return false;
}

document.querySelectorAll('.nav-item').forEach(item => {
  item.addEventListener('click', e => {
    e.preventDefault();
    navigate(item.dataset.page);
  });
});

function refreshCurrentPage(forcePage) {
  const activePage = forcePage || document.querySelector('.page.active')?.id?.replace('page-', '');
  if (!activePage) return;

  switch (activePage) {
    case 'dashboard':    renderDashboard(); break;
    case 'expenses':     renderExpensesTable(); break;
    case 'add':          renderTodayStats(); break;
    case 'analytics':    renderAnalytics(); break;
    case 'custom-charts': updateCustomChart(); break;
    case 'settings':     renderSettings(); break;
    case 'import':       renderImportHistory(); break;
  }
}

/* ============================================================ MOBILE SIDEBAR */
function openMobileSidebar() {
  document.getElementById('sidebar').classList.add('mobile-open');
  document.getElementById('sidebarOverlay').classList.add('visible');
}

function closeMobileSidebar() {
  document.getElementById('sidebar').classList.remove('mobile-open');
  document.getElementById('sidebarOverlay').classList.remove('visible');
}

/* ============================================================ SIDEBAR TOGGLE */
function initSidebar() {
  const btn = document.getElementById('sidebarToggle');
  if (!btn) return;
  btn.addEventListener('click', () => {
    document.getElementById('sidebar').classList.toggle('collapsed');
  });
}

/* ============================================================ FONT SIZE */
function initFontSize() {
  const saved = parseInt(localStorage.getItem('rupaiya_font_size'), 10) || 16;
  document.documentElement.style.setProperty('--base-font-size', saved + 'px');
  // Highlight the active button once settings page renders
  document.addEventListener('rupaiya:settingsRendered', () => syncFontSizeBtns(saved));
}

function setFontSize(px, btn) {
  document.documentElement.style.setProperty('--base-font-size', px + 'px');
  localStorage.setItem('rupaiya_font_size', px);
  syncFontSizeBtns(px);
}

function syncFontSizeBtns(px) {
  document.querySelectorAll('.font-size-btn').forEach(b => {
    b.classList.toggle('active', parseInt(b.dataset.size, 10) === px);
  });
}

/* ============================================================ GLOBAL MONTH FILTER */
function populateGlobalMonthFilter() {
  const sel = document.getElementById('globalMonthFilter');
  const months = getMonthOptions();
  // Clear and add "All Time"
  sel.innerHTML = '<option value="all">All Time</option>';
  months.forEach(m => {
    sel.add(new Option(m.label, m.value));
  });
  // Default to current month if data exists
  const now = new Date();
  const curVal = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  if (months.find(m => m.value === curVal)) sel.value = curVal;
}

function getMonthOptions() {
  const months = new Set();
  state.expenses.forEach(e => {
    const d = e.date.substring(0, 7); // YYYY-MM
    months.add(d);
  });
  // Also add current month
  const now = new Date();
  months.add(`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`);
  return [...months].sort().reverse().map(m => {
    const [y, mo] = m.split('-');
    return { value: m, label: `${MONTHS_SHORT[parseInt(mo)-1]} ${y}` };
  });
}

function getFilteredExpenses() {
  const monthFilter = document.getElementById('globalMonthFilter').value;
  if (monthFilter === 'all') return [...state.expenses];
  return state.expenses.filter(e => e.date.startsWith(monthFilter));
}

/* ============================================================ EXPENSE CRUD */
function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

function saveExpense(e) {
  e.preventDefault();
  const category = document.getElementById('expCategory').value.trim();
  const nature = document.getElementById('expNature').value.trim();

  // Auto-add new category / nature on the fly
  let newAdded = [];
  if (category && !state.categories.includes(category)) {
    state.categories.push(category);
    newAdded.push(`category "${category}"`);
  }
  if (nature && !state.nature.includes(nature)) {
    state.nature.push(nature);
    newAdded.push(`nature "${nature}"`);
  }

  const expense = {
    id: generateId(),
    date: document.getElementById('expDate').value,
    amount: parseFloat(document.getElementById('expAmount').value),
    description: '',
    category: category || 'Other',
    nature: nature || 'Personal',
    paidBy: document.getElementById('expPaidBy').value,
    notes: '',
    createdAt: Date.now(),
  };
  state.expenses.push(expense);
  if (currentUserId) window._fb.addExpense(currentUserId, expense).catch(console.error);
  saveState(); // persists any new category/nature added above
  populateGlobalMonthFilter();
  populateSelects();
  renderSidebarStats();
  if (newAdded.length) {
    showToast(`Expense added! New ${newAdded.join(' & ')} saved.`, 'success');
  } else {
    showToast('Expense added!', 'success');
  }
  resetForm();
  renderTodayStats();
}

function resetForm() {
  document.getElementById('expenseForm').reset();
  document.getElementById('expCategory').value = '';
  document.getElementById('expNature').value = '';
  setDefaultDate();
}

function setDefaultDate() {
  const today = new Date().toISOString().split('T')[0];
  const el = document.getElementById('expDate');
  if (el) el.value = today;
}

function deleteExpense(id) {
  if (!confirm('Delete this expense?')) return;
  state.expenses = state.expenses.filter(e => e.id !== id);
  if (currentUserId) window._fb.deleteExpense(currentUserId, id).catch(console.error);
  renderExpensesTable();
  populateGlobalMonthFilter();
  renderSidebarStats();
  showToast('Expense deleted.', 'warning');
}

function openEditModal(id) {
  const exp = state.expenses.find(e => e.id === id);
  if (!exp) return;
  document.getElementById('editExpId').value = exp.id;
  document.getElementById('editDate').value = exp.date;
  document.getElementById('editAmount').value = exp.amount;

  const editCat = document.getElementById('editCategory');
  const editNat = document.getElementById('editNature');
  const editPay = document.getElementById('editPaidBy');
  populateSelectEl(editCat, state.categories, exp.category);
  populateSelectEl(editNat, state.nature, exp.nature);
  populateSelectEl(editPay, state.paidBy, exp.paidBy);

  document.getElementById('editModal').style.display = 'flex';
}

function updateExpense(e) {
  e.preventDefault();
  const id = document.getElementById('editExpId').value;
  const idx = state.expenses.findIndex(ex => ex.id === id);
  if (idx < 0) return;
  const updated = {
    date:     document.getElementById('editDate').value,
    amount:   parseFloat(document.getElementById('editAmount').value),
    category: document.getElementById('editCategory').value,
    nature:   document.getElementById('editNature').value,
    paidBy:   document.getElementById('editPaidBy').value,
  };
  state.expenses[idx] = { ...state.expenses[idx], ...updated };
  if (currentUserId) window._fb.updateExpense(currentUserId, id, updated).catch(console.error);
  closeModal('editModal');
  renderExpensesTable();
  if (document.getElementById('page-dashboard').classList.contains('active')) renderDashboard();
  renderSidebarStats();
  showToast('Expense updated!', 'success');
}

function closeModal(id) {
  document.getElementById(id).style.display = 'none';
}

/* ============================================================ DASHBOARD */
function renderDashboard() {
  const all = getFilteredExpenses();
  const now = new Date();
  const curMonth = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;

  const total = all.reduce((s, e) => s + e.amount, 0);
  const thisMonth = all.filter(e => e.date.startsWith(curMonth)).reduce((s,e) => s+e.amount, 0);

  // Daily avg this month
  const daysElapsed = Math.max(1, now.getDate());
  const dailyAvg = thisMonth / daysElapsed;

  // Top category
  const byCat = groupBy(all, 'category');
  let topCat = '—', topCatAmt = 0;
  Object.entries(byCat).forEach(([cat, exps]) => {
    const sum = exps.reduce((s,e) => s+e.amount, 0);
    if (sum > topCatAmt) { topCatAmt = sum; topCat = cat; }
  });

  // Last month for comparison
  const lastMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const lastMonth = `${lastMonthDate.getFullYear()}-${String(lastMonthDate.getMonth()+1).padStart(2,'0')}`;
  const lastMonthTotal = state.expenses.filter(e => e.date.startsWith(lastMonth)).reduce((s,e)=>s+e.amount,0);
  const changePercent = lastMonthTotal > 0 ? ((thisMonth - lastMonthTotal) / lastMonthTotal * 100).toFixed(1) : null;

  document.getElementById('totalSpent').textContent = formatINR(total);
  document.getElementById('totalCount').textContent = `${all.length} transactions`;
  document.getElementById('monthSpent').textContent = formatINR(thisMonth);
  document.getElementById('dailyAvg').textContent = formatINR(dailyAvg);
  document.getElementById('topCategory').textContent = topCat !== '—' ? `${getCategoryEmoji(topCat)} ${topCat}` : '—';
  document.getElementById('topCategoryAmt').textContent = topCat !== '—' ? formatINR(topCatAmt) : '';

  if (changePercent !== null) {
    const up = parseFloat(changePercent) >= 0;
    document.getElementById('monthChange').innerHTML = `<span style="color:${up?'var(--danger)':'var(--success)'}">` +
      `${up?'▲':'▼'} ${Math.abs(changePercent)}% vs last month</span>`;
  } else {
    document.getElementById('monthChange').textContent = '';
  }

  renderTrendChart();
  renderCategoryChart();
  renderPaymentChart();
  renderDOWChart();
  renderNatureChart();
  renderRecentExpenses(all);
}

function renderRecentExpenses(expenses) {
  const container = document.getElementById('recentExpenses');
  const recent = [...expenses].sort((a,b) => b.date.localeCompare(a.date) || b.createdAt - a.createdAt).slice(0, 8);
  if (recent.length === 0) {
    container.innerHTML = `<div class="empty-state"><div class="empty-icon">💸</div><p>No expenses yet.</p><p class="hint">Click "Add Expense" to get started!</p></div>`;
    return;
  }
  container.innerHTML = recent.map(e => `
    <div class="recent-item">
      <div class="recent-icon" style="background:${hexAlpha(getCategoryColor(e.category),0.13)};color:${getCategoryColor(e.category)}">
        ${getCategoryEmoji(e.category)}
      </div>
      <div class="recent-info">
        <div class="recent-desc">${e.description || e.category}</div>
        <div class="recent-meta">${e.category} &bull; ${e.nature} &bull; ${e.paidBy}</div>
      </div>
      <div>
        <div class="recent-amount">${formatINR(e.amount)}</div>
        <div class="recent-date">${formatDate(e.date)}</div>
      </div>
    </div>
  `).join('');
}

/* ============================================================ CHARTS - DASHBOARD */
function renderTrendChart() {
  const type = document.getElementById('trendChartType').value;
  const all = state.expenses;
  const monthMap = {};
  all.forEach(e => {
    const m = e.date.substring(0,7);
    monthMap[m] = (monthMap[m] || 0) + e.amount;
  });
  const sorted = Object.keys(monthMap).sort();
  const labels = sorted.map(m => { const [y,mo]=m.split('-'); return `${MONTHS_SHORT[parseInt(mo)-1]} ${y.slice(2)}`; });
  const data = sorted.map(m => monthMap[m]);

  const isDark = document.body.classList.contains('dark');
  const gridColor = isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)';
  const textColor = isDark ? '#a0a8c0' : '#6b7280';

  destroyChart('trendChart');
  if (data.length === 0) return;
  const ctx = document.getElementById('trendChart').getContext('2d');
  chartInstances['trendChart'] = new Chart(ctx, {
    type,
    data: {
      labels,
      datasets: [{
        label: 'Spending (₹)',
        data,
        borderColor: '#6c63ff',
        backgroundColor: type === 'line' ? 'rgba(108,99,255,0.12)' : 'rgba(108,99,255,0.7)',
        borderWidth: 2,
        fill: type === 'line',
        tension: 0.4,
        pointBackgroundColor: '#6c63ff',
        pointRadius: 5,
        pointHoverRadius: 7,
        borderRadius: type === 'bar' ? 6 : 0,
      }]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { color: gridColor }, ticks: { color: textColor } },
        y: {
          grid: { color: gridColor },
          ticks: { color: textColor, callback: v => '₹' + formatNum(v) },
          beginAtZero: true,
        }
      }
    }
  });
}

function renderCategoryChart() {
  const all = getFilteredExpenses();
  const byCat = groupBy(all, 'category');
  const labels = Object.keys(byCat);
  const data = labels.map(k => byCat[k].reduce((s,e)=>s+e.amount,0));
  const colors = labels.map(l => getCategoryColor(l));

  destroyChart('categoryChart');
  if (data.length === 0) return;
  const ctx = document.getElementById('categoryChart').getContext('2d');
  chartInstances['categoryChart'] = new Chart(ctx, {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: colors, borderWidth: 2, borderColor: document.body.classList.contains('dark') ? '#1e2130' : '#ffffff' }] },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'right', labels: { font: { size: 11 }, color: getTextColor(), padding: 10, boxWidth: 14 } },
        tooltip: { callbacks: { label: c => ` ${c.label}: ${formatINR(c.raw)} (${((c.raw/c.dataset.data.reduce((a,b)=>a+b,0))*100).toFixed(1)}%)` } }
      },
      cutout: '62%',
    }
  });
}

function renderPaymentChart() {
  const all = getFilteredExpenses();
  const byPay = groupBy(all, 'paidBy');
  const labels = Object.keys(byPay);
  const data = labels.map(k => byPay[k].reduce((s,e)=>s+e.amount,0));

  destroyChart('paymentChart');
  if (data.length === 0) return;
  const ctx = document.getElementById('paymentChart').getContext('2d');
  chartInstances['paymentChart'] = new Chart(ctx, {
    type: 'pie',
    data: { labels, datasets: [{ data, backgroundColor: PALETTES.cool, borderWidth: 2, borderColor: document.body.classList.contains('dark') ? '#1e2130' : '#ffffff' }] },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 10 }, color: getTextColor(), padding: 8, boxWidth: 12 } },
      }
    }
  });
}

function renderDOWChart() {
  const all = getFilteredExpenses();
  const dowTotals = Array(7).fill(0);
  all.forEach(e => {
    const dow = new Date(e.date).getDay();
    dowTotals[dow] += e.amount;
  });

  const isDark = document.body.classList.contains('dark');
  const gridColor = isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)';
  const textColor = isDark ? '#a0a8c0' : '#6b7280';

  destroyChart('dowChart');
  const ctx = document.getElementById('dowChart').getContext('2d');
  chartInstances['dowChart'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: DOW_LABELS,
      datasets: [{
        data: dowTotals,
        backgroundColor: dowTotals.map((_, i) => `hsl(${250 + i*12}, 70%, 62%)`),
        borderRadius: 6,
      }]
    },
    options: {
      responsive: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { color: textColor } },
        y: { grid: { color: gridColor }, ticks: { color: textColor, callback: v => '₹' + formatNum(v) }, beginAtZero: true }
      }
    }
  });
}

function renderNatureChart() {
  const all = getFilteredExpenses();
  const byNat = groupBy(all, 'nature');
  const labels = Object.keys(byNat);
  const data = labels.map(k => byNat[k].reduce((s,e)=>s+e.amount,0));

  destroyChart('natureChart');
  if (data.length === 0) return;
  const ctx = document.getElementById('natureChart').getContext('2d');
  chartInstances['natureChart'] = new Chart(ctx, {
    type: 'polarArea',
    data: { labels, datasets: [{ data, backgroundColor: PALETTES.warm.map(c => c + 'cc'), borderColor: PALETTES.warm, borderWidth: 1 }] },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 10 }, color: getTextColor(), padding: 8, boxWidth: 12 } },
      },
      scales: { r: { ticks: { display: false }, grid: { color: 'rgba(128,128,128,0.15)' } } }
    }
  });
}

/* ============================================================ EXPENSES TABLE */
function renderExpensesTable() {
  const monthFilter = document.getElementById('globalMonthFilter').value;
  let expenses = monthFilter === 'all' ? [...state.expenses] : state.expenses.filter(e => e.date.startsWith(monthFilter));

  // Apply filters
  const search = document.getElementById('searchInput').value.toLowerCase();
  const cat = document.getElementById('filterCategory').value;
  const nat = document.getElementById('filterNature').value;
  const pay = document.getElementById('filterPaidBy').value;
  const from = document.getElementById('filterFrom').value;
  const to = document.getElementById('filterTo').value;

  if (search) expenses = expenses.filter(e =>
    (e.description||'').toLowerCase().includes(search) ||
    e.category.toLowerCase().includes(search) ||
    e.nature.toLowerCase().includes(search) ||
    e.paidBy.toLowerCase().includes(search) ||
    String(e.amount).includes(search)
  );
  if (cat) expenses = expenses.filter(e => e.category === cat);
  if (nat) expenses = expenses.filter(e => e.nature === nat);
  if (pay) expenses = expenses.filter(e => e.paidBy === pay);
  if (from) expenses = expenses.filter(e => e.date >= from);
  if (to) expenses = expenses.filter(e => e.date <= to);

  // Sort
  expenses.sort((a, b) => {
    let va = a[state.sortKey], vb = b[state.sortKey];
    if (state.sortKey === 'amount') { va = parseFloat(va); vb = parseFloat(vb); }
    if (va < vb) return state.sortDir === 'asc' ? -1 : 1;
    if (va > vb) return state.sortDir === 'asc' ? 1 : -1;
    return 0;
  });

  const tbody = document.getElementById('expenseTableBody');
  if (expenses.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7"><div class="empty-state"><div class="empty-icon">🔍</div><p>No expenses found.</p></div></td></tr>`;
    document.getElementById('tableCount').textContent = '0 records';
    document.getElementById('tableTotal').textContent = '';
    updateBulkBar();
    return;
  }

  // Keep only selected IDs that are in current filtered view
  const visibleIds = new Set(expenses.map(e => e.id));
  selectedIds = new Set([...selectedIds].filter(id => visibleIds.has(id)));

  tbody.innerHTML = expenses.map(e => `
    <tr class="${selectedIds.has(e.id) ? 'row-selected' : ''}">
      <td class="col-check">
        <label class="custom-checkbox">
          <input type="checkbox" ${selectedIds.has(e.id) ? 'checked' : ''} onchange="toggleSelectRow('${e.id}', this)" />
          <span class="checkmark"></span>
        </label>
      </td>
      <td>${formatDate(e.date)}</td>
      <td>
        <span class="badge-category" style="background:${hexAlpha(getCategoryColor(e.category),0.13)};color:${getCategoryColor(e.category)}">
          ${getCategoryEmoji(e.category)} ${e.category}
        </span>
      </td>
      <td><span class="badge-category" style="background:var(--surface2);color:var(--text2)">${e.nature}</span></td>
      <td style="font-weight:700;color:var(--primary)">${formatINR(e.amount)}</td>
      <td>${e.paidBy}</td>
      <td>
        <div class="action-btns">
          <button class="btn-icon btn-edit" onclick="openEditModal('${e.id}')" title="Edit">&#9998;</button>
          <button class="btn-icon btn-delete" onclick="deleteExpense('${e.id}')" title="Delete">&#128465;</button>
        </div>
      </td>
    </tr>
  `).join('');

  // Sync select-all checkbox state
  const selectAllEl = document.getElementById('selectAll');
  if (selectAllEl) {
    selectAllEl.checked = expenses.length > 0 && expenses.every(e => selectedIds.has(e.id));
    selectAllEl.indeterminate = selectedIds.size > 0 && !selectAllEl.checked;
  }

  updateBulkBar();

  const total = expenses.reduce((s,e)=>s+e.amount,0);
  document.getElementById('tableCount').textContent = `${expenses.length} record${expenses.length!==1?'s':''}`;
  document.getElementById('tableTotal').textContent = `Total: ${formatINR(total)}`;
}

function sortTable(key) {
  if (state.sortKey === key) state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
  else { state.sortKey = key; state.sortDir = 'asc'; }
  renderExpensesTable();
}

function clearFilters() {
  document.getElementById('searchInput').value = '';
  document.getElementById('filterCategory').value = '';
  document.getElementById('filterNature').value = '';
  document.getElementById('filterPaidBy').value = '';
  document.getElementById('filterFrom').value = '';
  document.getElementById('filterTo').value = '';
  renderExpensesTable();
}

/* ============================================================ BULK SELECTION */
function toggleSelectAll(el) {
  const rows = document.querySelectorAll('#expenseTableBody tr');
  rows.forEach(row => {
    const cb = row.querySelector('input[type="checkbox"]');
    if (!cb) return;
    const id = cb.getAttribute('onchange').match(/'([^']+)'/)?.[1];
    if (!id) return;
    if (el.checked) {
      selectedIds.add(id);
      row.classList.add('row-selected');
    } else {
      selectedIds.delete(id);
      row.classList.remove('row-selected');
    }
    cb.checked = el.checked;
  });
  updateBulkBar();
}

function toggleSelectRow(id, el) {
  if (el.checked) {
    selectedIds.add(id);
    el.closest('tr').classList.add('row-selected');
  } else {
    selectedIds.delete(id);
    el.closest('tr').classList.remove('row-selected');
  }
  // Sync select-all
  const allCbs = document.querySelectorAll('#expenseTableBody input[type="checkbox"]');
  const allChecked = allCbs.length > 0 && [...allCbs].every(c => c.checked);
  const selectAllEl = document.getElementById('selectAll');
  if (selectAllEl) {
    selectAllEl.checked = allChecked;
    selectAllEl.indeterminate = selectedIds.size > 0 && !allChecked;
  }
  updateBulkBar();
}

function updateBulkBar() {
  const bar = document.getElementById('bulkBar');
  const countEl = document.getElementById('bulkCount');
  if (!bar) return;
  if (selectedIds.size > 0) {
    countEl.textContent = selectedIds.size;
    bar.classList.add('visible');
  } else {
    bar.classList.remove('visible');
  }
}

function clearSelection() {
  selectedIds.clear();
  document.querySelectorAll('#expenseTableBody tr').forEach(row => {
    row.classList.remove('row-selected');
    const cb = row.querySelector('input[type="checkbox"]');
    if (cb) cb.checked = false;
  });
  const selectAllEl = document.getElementById('selectAll');
  if (selectAllEl) { selectAllEl.checked = false; selectAllEl.indeterminate = false; }
  updateBulkBar();
}

function bulkDelete() {
  const count = selectedIds.size;
  if (count === 0) return;
  if (!confirm(`Delete ${count} selected expense${count !== 1 ? 's' : ''}? This cannot be undone.`)) return;
  const deletedIds = [...selectedIds];
  state.expenses = state.expenses.filter(e => !selectedIds.has(e.id));
  selectedIds.clear();
  if (currentUserId) window._fb.batchDeleteExpenses(currentUserId, deletedIds).catch(console.error);
  populateGlobalMonthFilter();
  renderExpensesTable();
  if (document.getElementById('page-dashboard').classList.contains('active')) renderDashboard();
  renderSidebarStats();
  showToast(`${count} expense${count !== 1 ? 's' : ''} deleted.`, 'warning');
}

function openBulkEditModal() {
  if (selectedIds.size === 0) return;
  document.getElementById('bulkEditCount').textContent = `(${selectedIds.size} selected)`;

  const bulkCat = document.getElementById('bulkCategory');
  const bulkNat = document.getElementById('bulkNature');
  const bulkPay = document.getElementById('bulkPaidBy');
  populateSelectEl(bulkCat, state.categories, '', '— Keep existing —');
  populateSelectEl(bulkNat, state.nature, '', '— Keep existing —');
  populateSelectEl(bulkPay, state.paidBy, '', '— Keep existing —');

  document.getElementById('bulkDate').value = '';
  document.getElementById('bulkEditModal').style.display = 'flex';
}

function applyBulkEdit(e) {
  e.preventDefault();
  const newCat = document.getElementById('bulkCategory').value;
  const newNat = document.getElementById('bulkNature').value;
  const newPay = document.getElementById('bulkPaidBy').value;
  const newDate = document.getElementById('bulkDate').value;

  if (!newCat && !newNat && !newPay && !newDate) {
    showToast('No fields selected to update.', 'warning');
    return;
  }

  let count = 0;
  const firestoreUpdates = [];
  state.expenses.forEach(exp => {
    if (!selectedIds.has(exp.id)) return;
    const data = {};
    if (newCat)  { exp.category = newCat;  data.category = newCat; }
    if (newNat)  { exp.nature   = newNat;  data.nature   = newNat; }
    if (newPay)  { exp.paidBy   = newPay;  data.paidBy   = newPay; }
    if (newDate) { exp.date     = newDate; data.date     = newDate; }
    firestoreUpdates.push({ id: exp.id, data });
    count++;
  });

  if (currentUserId && firestoreUpdates.length) {
    window._fb.batchUpdateExpenses(currentUserId, firestoreUpdates).catch(console.error);
  }
  closeModal('bulkEditModal');
  clearSelection();
  renderExpensesTable();
  if (document.getElementById('page-dashboard').classList.contains('active')) renderDashboard();
  renderSidebarStats();
  showToast(`${count} expense${count !== 1 ? 's' : ''} updated!`, 'success');
}

/* ============================================================ TODAY STATS */
function renderTodayStats() {
  const today = new Date().toISOString().split('T')[0];
  const todayExp = state.expenses.filter(e => e.date === today);
  const todayTotal = todayExp.reduce((s,e)=>s+e.amount,0);

  const curMonth = today.substring(0,7);
  const monthExp = state.expenses.filter(e => e.date.startsWith(curMonth));
  const monthTotal = monthExp.reduce((s,e)=>s+e.amount,0);

  document.getElementById('todayStats').innerHTML = `
    <div class="stat-item"><div class="stat-value">${formatINR(todayTotal)}</div><div class="stat-label">Today</div></div>
    <div class="stat-item"><div class="stat-value">${todayExp.length}</div><div class="stat-label">Today's Entries</div></div>
    <div class="stat-item"><div class="stat-value">${formatINR(monthTotal)}</div><div class="stat-label">This Month</div></div>
    <div class="stat-item"><div class="stat-value">${monthExp.length}</div><div class="stat-label">Month Entries</div></div>
  `;
}

/* ============================================================ EXCEL IMPORT */
function initDropZone() {
  const dz = document.getElementById('dropZone');
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', e => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) processImportFile(file);
  });
}

function handleFileUpload(e) {
  const file = e.target.files[0];
  if (file) processImportFile(file);
}

function processImportFile(file) {
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const data = new Uint8Array(ev.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

      if (json.length < 2) { showToast('File is empty or has no data rows.', 'error'); return; }

      importHeaders = json[0].map(h => String(h || '').trim());
      importData_raw = json.slice(1).filter(row => row.some(c => c));

      showMappingUI();
      showPreview(json.slice(0, 6));
    } catch (err) {
      showToast('Error reading file: ' + err.message, 'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function showMappingUI() {
  const fields = [
    { key: 'date', label: 'Date *' },
    { key: 'amount', label: 'Amount *' },
    { key: 'description', label: 'Description' },
    { key: 'category', label: 'Category' },
    { key: 'nature', label: 'Nature' },
    { key: 'paidBy', label: 'Paid By' },
    { key: 'notes', label: 'Notes' },
  ];

  const mappingGrid = document.getElementById('mappingGrid');
  mappingGrid.innerHTML = fields.map(f => `
    <div class="mapping-item">
      <label>${f.label}</label>
      <select class="input" id="map_${f.key}">
        <option value="">-- Skip --</option>
        ${importHeaders.map((h, i) => `<option value="${i}" ${autoMap(h, f.key) ? 'selected' : ''}>${h}</option>`).join('')}
      </select>
    </div>
  `).join('');

  document.getElementById('mappingSection').style.display = 'block';
  document.getElementById('importResult').style.display = 'none';
}

function autoMap(header, field) {
  const h = header.toLowerCase().replace(/[\s_-]/g, '');
  const maps = {
    date: ['date','dt','transactiondate','txndate'],
    amount: ['amount','amt','price','cost','value','inr','rs','rupees','debit'],
    description: ['description','desc','narration','details','particulars','name','item'],
    category: ['category','cat','type','head'],
    nature: ['nature','kind','typeof','classification'],
    paidBy: ['paidby','paymentmethod','mode','via','payby','modeofpayment'],
    notes: ['notes','note','remarks','remark','comment'],
  };
  return (maps[field] || []).includes(h);
}

function showPreview(rows) {
  const headers = rows[0] || [];
  const dataRows = rows.slice(1);
  const table = `
    <table class="expense-table">
      <thead><tr>${headers.map(h=>`<th>${h||''}</th>`).join('')}</tr></thead>
      <tbody>${dataRows.map(r=>`<tr>${headers.map((_,i)=>`<td>${r[i]||''}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>
  `;
  document.getElementById('previewContainer').innerHTML = table;
  document.getElementById('previewContainer').style.display = 'block';
}

function importData() {
  const getCol = key => {
    const el = document.getElementById('map_' + key);
    return el && el.value !== '' ? parseInt(el.value) : null;
  };

  const dateCol = getCol('date');
  const amtCol = getCol('amount');
  if (dateCol === null || amtCol === null) {
    showToast('Date and Amount columns are required!', 'error'); return;
  }

  let imported = 0, skipped = 0;
  const now = Date.now();
  const importId = 'imp_' + now;   // unique tag for this batch
  const prevLength = state.expenses.length;

  const descCol  = getCol('description');
  const catCol   = getCol('category');
  const natCol   = getCol('nature');
  const payCol   = getCol('paidBy');
  const notesCol = getCol('notes');

  importData_raw.forEach((row, idx) => {
    const rawDate = row[dateCol];
    const rawAmt = row[amtCol];
    if (!rawDate || !rawAmt) { skipped++; return; }

    const parsedDate = parseExcelDate(rawDate);
    const parsedAmt = parseFloat(String(rawAmt).replace(/[^0-9.-]/g, ''));
    if (!parsedDate || isNaN(parsedAmt) || parsedAmt <= 0) { skipped++; return; }

    const rawCat = catCol !== null ? (row[catCol] || '') : '';
    const rawNat = natCol !== null ? (row[natCol] || '') : '';
    const rawPay = payCol !== null ? (row[payCol] || '') : '';

    // Auto-add unknown categories/natures/paidBy
    const cat = matchOrAdd(rawCat, state.categories, 'Other');
    const nat = matchOrAdd(rawNat, state.nature, 'Personal');
    const pay = matchOrAdd(rawPay, state.paidBy, 'Cash');

    state.expenses.push({
      id: generateId() + idx,
      importId,
      date: parsedDate,
      amount: parsedAmt,
      description: descCol !== null ? (row[descCol] || '') : '',
      category: cat,
      nature: nat,
      paidBy: pay,
      notes: notesCol !== null ? (row[notesCol] || '') : '',
      createdAt: now + idx,
    });
    imported++;
  });

  const newExpenses = state.expenses.slice(prevLength);
  if (currentUserId && newExpenses.length) {
    window._fb.batchAddExpenses(currentUserId, newExpenses).catch(console.error);
  }
  saveState(); // persist any new categories/natures added during import
  populateGlobalMonthFilter();
  populateSelects();
  renderSidebarStats();

  const resultEl = document.getElementById('importResult');
  resultEl.className = 'import-result ' + (imported > 0 ? 'success' : 'error');
  resultEl.innerHTML = `<strong>${imported > 0 ? '✅' : '❌'} Import Complete</strong><br>
    Imported: <strong>${imported}</strong> rows &bull; Skipped: <strong>${skipped}</strong> rows`;
  resultEl.style.display = 'block';
  document.getElementById('mappingSection').style.display = 'none';

  if (imported > 0) {
    showToast(`${imported} expenses imported!`, 'success');
    renderImportHistory();
  }
}

function matchOrAdd(value, list, defaultVal) {
  if (!value) return defaultVal;
  const v = value.trim();
  if (!v) return defaultVal;
  const found = list.find(item => item.toLowerCase() === v.toLowerCase());
  if (found) return found;
  // Add if new
  if (v.length > 0 && !list.includes(v)) list.push(v);
  return v || defaultVal;
}

function cancelImport() {
  document.getElementById('mappingSection').style.display = 'none';
  document.getElementById('importResult').style.display = 'none';
  importData_raw = null;
  document.getElementById('fileInput').value = '';
}

function renderImportHistory() {
  const container = document.getElementById('importHistory');
  if (!container) return;

  // Group expenses by importId
  const batches = {};
  state.expenses.forEach(e => {
    if (!e.importId) return;
    if (!batches[e.importId]) {
      batches[e.importId] = { id: e.importId, timestamp: parseInt(e.importId.replace('imp_', '')), count: 0, total: 0 };
    }
    batches[e.importId].count++;
    batches[e.importId].total += e.amount;
  });

  const sorted = Object.values(batches).sort((a, b) => b.timestamp - a.timestamp);

  if (sorted.length === 0) {
    container.innerHTML = '<p class="hint" style="padding:8px 0">No Excel imports found. Future imports will appear here.</p>';
    return;
  }

  container.innerHTML = sorted.map(b => {
    const d = new Date(b.timestamp);
    const dateStr = d.toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
    const timeStr = d.toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' });
    return `
      <div class="import-history-row" id="imp-row-${b.id}">
        <div class="import-history-icon">📥</div>
        <div class="import-history-info">
          <div class="import-history-date">${dateStr} · ${timeStr}</div>
          <div class="import-history-meta">${b.count} expense${b.count !== 1 ? 's' : ''} · ${formatINR(b.total)}</div>
        </div>
        <button class="btn btn-sm btn-danger-outline" onclick="deleteImport('${b.id}', ${b.count})">&#128465; Delete</button>
      </div>`;
  }).join('');
}

async function deleteImport(importId, count) {
  const confirmed = confirm(`Delete all ${count} expenses from this import?\n\nThis cannot be undone.`);
  if (!confirmed) return;

  const ids = state.expenses.filter(e => e.importId === importId).map(e => e.id);
  if (!ids.length) return;

  state.expenses = state.expenses.filter(e => e.importId !== importId);

  if (currentUserId) {
    window._fb.batchDeleteExpenses(currentUserId, ids).catch(console.error);
  }

  renderImportHistory();
  populateGlobalMonthFilter();
  renderSidebarStats();
  refreshCurrentPage('import');
  showToast(`Deleted ${ids.length} imported expenses.`, 'success');
}

function parseExcelDate(val) {
  if (!val) return null;
  const s = String(val).trim();

  // Try ISO
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0,10);

  // Try DD/MM/YYYY or DD-MM-YYYY
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) return `${dmy[3]}-${dmy[2].padStart(2,'0')}-${dmy[1].padStart(2,'0')}`;

  // Try MM/DD/YYYY
  const mdy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (mdy) {
    const y = mdy[3].length===2 ? '20'+mdy[3] : mdy[3];
    return `${y}-${mdy[1].padStart(2,'0')}-${mdy[2].padStart(2,'0')}`;
  }

  // Try Excel serial date
  const num = parseFloat(s);
  if (!isNaN(num) && num > 1000) {
    const d = new Date((num - 25569) * 86400 * 1000);
    if (!isNaN(d.getTime())) return d.toISOString().split('T')[0];
  }

  // Try natural parsing
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().split('T')[0];

  return null;
}

function downloadTemplate() {
  const wb = XLSX.utils.book_new();
  const headers = [['Date', 'Amount', 'Description', 'Category', 'Nature', 'Paid By', 'Notes']];
  const sample = [
    ['2025-01-15', '450', 'Lunch at restaurant', 'Food & Dining', 'Personal', 'UPI', ''],
    ['2025-01-16', '1200', 'Cab to airport', 'Transport', 'Business', 'Credit Card', 'Official trip'],
    ['2025-01-17', '5000', 'Monthly groceries', 'Shopping', 'Family', 'Debit Card', ''],
  ];
  const ws = XLSX.utils.aoa_to_sheet([...headers, ...sample]);
  ws['!cols'] = [15,10,25,20,15,15,25].map(w => ({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws, 'Expenses');
  XLSX.writeFile(wb, 'rupaiya_template.xlsx');
}

/* ============================================================ EXPORT */
function exportAllExcel() {
  const expenses = getFilteredExpenses();
  if (expenses.length === 0) { showToast('No data to export.', 'warning'); return; }

  const rows = [['Date','Description','Category','Nature','Amount (INR)','Paid By','Notes']];
  expenses.sort((a,b)=>b.date.localeCompare(a.date)).forEach(e => {
    rows.push([e.date, e.description||'', e.category, e.nature, e.amount, e.paidBy, e.notes||'']);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [12,25,20,15,15,15,25].map(w => ({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws, 'Expenses');

  const filename = `rupaiya_export_${new Date().toISOString().split('T')[0]}.xlsx`;
  XLSX.writeFile(wb, filename);
  showToast('Exported to Excel!', 'success');
}

function exportJSON() {
  const data = JSON.stringify(state, null, 2);
  const blob = new Blob([data], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = 'rupaiya_backup.json';
  a.click(); URL.revokeObjectURL(url);
  showToast('JSON backup exported!', 'success');
}

function importJSON() {
  document.getElementById('jsonImportFile').click();
}

function handleJSONImport(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    try {
      const data = JSON.parse(ev.target.result);
      if (data.expenses) {
        if (confirm(`Import ${data.expenses.length} expenses? This will MERGE with existing data.`)) {
          const existingIds = new Set(state.expenses.map(e => e.id));
          const newExps = data.expenses.filter(e => !existingIds.has(e.id));
          state.expenses.push(...newExps);
          if (data.categories) state.categories = [...new Set([...state.categories, ...data.categories])];
          if (data.nature) state.nature = [...new Set([...state.nature, ...data.nature])];
          if (data.paidBy) state.paidBy = [...new Set([...state.paidBy, ...data.paidBy])];
          if (currentUserId && newExps.length) {
            window._fb.batchAddExpenses(currentUserId, newExps).catch(console.error);
          }
          saveState();
          populateSelects();
          populateGlobalMonthFilter();
          renderSidebarStats();
          showToast(`${newExps.length} new expenses imported!`, 'success');
        }
      } else showToast('Invalid backup file.', 'error');
    } catch (err) { showToast('Error reading JSON: ' + err.message, 'error'); }
  };
  reader.readAsText(file);
  e.target.value = '';
}

function clearAllData() {
  if (confirm('Are you sure? This will delete ALL expenses permanently!')) {
    if (confirm('Really delete everything? This cannot be undone!')) {
      state.expenses = [];
      state.budgets = {};
      state.savedCharts = [];
      if (currentUserId) window._fb.clearAllExpenses(currentUserId).catch(console.error);
      saveState();
      populateGlobalMonthFilter();
      renderSidebarStats();
      showToast('All data cleared.', 'warning');
      navigate('dashboard');
    }
  }
}

/* ============================================================ ANALYTICS */
function renderAnalytics() {
  const all = getFilteredExpenses();
  try { renderMonthCategoryBreakdown(); } catch(e) { console.error('catBreakdown:', e); }
  try { renderSmartInsights(all); } catch(e) { console.error('smartInsights:', e); }
  try { render6MonthTrend(); } catch(e) { console.error('6monthTrend:', e); }
  try { renderTopExpensesChart(all); } catch(e) { console.error('topExpenses:', e); }
  try { renderPaidByAnalyticsChart(all); } catch(e) { console.error('paidBy:', e); }
  try { renderBudgetBars(all); } catch(e) { console.error('budgetBars:', e); }
}

function renderMonthCategoryBreakdown() {
  const now = new Date();
  const curMonth = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const selected = document.getElementById('globalMonthFilter').value;
  const month = (selected && selected !== 'all') ? selected : curMonth;

  const [y, m] = month.split('-').map(Number);
  const prevDate = new Date(y, m - 2, 1);
  const prevMonth = `${prevDate.getFullYear()}-${String(prevDate.getMonth()+1).padStart(2,'0')}`;

  const monthExp = state.expenses.filter(e => e.date.startsWith(month));
  const prevMonthExp = state.expenses.filter(e => e.date.startsWith(prevMonth));
  const total = monthExp.reduce((s,e) => s + e.amount, 0);

  const byCat = {};
  monthExp.forEach(e => {
    if (!byCat[e.category]) byCat[e.category] = { amount: 0, count: 0 };
    byCat[e.category].amount += e.amount;
    byCat[e.category].count++;
  });

  const prevByCat = {};
  prevMonthExp.forEach(e => { prevByCat[e.category] = (prevByCat[e.category] || 0) + e.amount; });

  const sorted = Object.entries(byCat).sort((a,b) => b[1].amount - a[1].amount);
  const maxAmt = sorted.length ? sorted[0][1].amount : 1;

  const [yr, mo] = month.split('-');
  const label = `${MONTHS_SHORT[parseInt(mo)-1]} ${yr}`;
  const periodEl = document.getElementById('catBreakdownPeriod');
  const totalEl  = document.getElementById('catBreakdownTotal');
  if (periodEl) periodEl.textContent = label;
  if (totalEl)  totalEl.textContent  = formatINR(total);

  const list = document.getElementById('catBreakdownList');
  if (!list) return;

  if (sorted.length === 0) {
    list.innerHTML = '<p class="hint" style="padding:28px;text-align:center">No expenses for this period.</p>';
    return;
  }

  list.innerHTML = sorted.map(([cat, data]) => {
    const pct    = total > 0 ? (data.amount / total * 100) : 0;
    const barW   = (data.amount / maxAmt * 100).toFixed(1);
    const prevAmt = prevByCat[cat] || 0;
    const change  = prevAmt > 0 ? ((data.amount - prevAmt) / prevAmt * 100) : null;
    const changeHtml = change !== null
      ? `<span class="cat-change ${change > 0 ? 'up' : 'down'}">${change > 0 ? '↑' : '↓'}${Math.abs(change).toFixed(0)}%</span>`
      : `<span class="cat-change new-cat">New</span>`;
    const color = getCategoryColor(cat);

    return `<div class="cat-row">
      <div class="cat-row-emoji" style="background:${hexAlpha(color,0.09)};border-color:${hexAlpha(color,0.18)}">${getCategoryEmoji(cat)}</div>
      <div class="cat-row-main">
        <div class="cat-row-top">
          <span class="cat-row-name">${cat}</span>
          <div class="cat-row-right">
            <span class="cat-row-amount" style="color:${color}">${formatINR(data.amount)}</span>
            ${changeHtml}
          </div>
        </div>
        <div class="cat-row-bar-row">
          <div class="cat-row-bar-wrap">
            <div class="cat-row-bar" style="width:${barW}%;background:${color}"></div>
          </div>
          <span class="cat-row-pct">${pct.toFixed(0)}%</span>
        </div>
        <span class="cat-row-meta">${data.count} transaction${data.count > 1 ? 's' : ''}</span>
      </div>
    </div>`;
  }).join('');
}

function renderSmartInsights(expenses) {
  const grid = document.getElementById('smartInsightsGrid');
  if (!grid) return;
  if (expenses.length === 0) { grid.innerHTML = ''; return; }

  const now = new Date();
  const curMonth = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const prevDate  = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const prevMonth = `${prevDate.getFullYear()}-${String(prevDate.getMonth()+1).padStart(2,'0')}`;

  const thisMonth = state.expenses.filter(e => e.date.startsWith(curMonth));
  const lastMonth = state.expenses.filter(e => e.date.startsWith(prevMonth));
  const thisTotal = thisMonth.reduce((s,e) => s + e.amount, 0);
  const lastTotal = lastMonth.reduce((s,e) => s + e.amount, 0);

  const insights = [];

  // Month-over-month
  if (lastTotal > 0) {
    const diff = thisTotal - lastTotal;
    const pct  = Math.abs(diff / lastTotal * 100).toFixed(0);
    insights.push({
      icon: diff >= 0 ? '📈' : '📉',
      title: 'Month over Month',
      value: (diff >= 0 ? '+' : '') + formatINR(diff),
      sub: `${pct}% ${diff >= 0 ? 'higher' : 'lower'} than ${MONTHS_SHORT[prevDate.getMonth()]}`,
      color: diff >= 0 ? '#EF4444' : '#10B981',
    });
  }

  // Highest spend day of week
  const byDow = Array(7).fill(0);
  const cntDow = Array(7).fill(0);
  expenses.forEach(e => { const d = new Date(e.date + 'T00:00:00').getDay(); byDow[d] += e.amount; cntDow[d]++; });
  const maxDow = byDow.indexOf(Math.max(...byDow));
  if (byDow[maxDow] > 0) {
    insights.push({
      icon: '📅',
      title: 'Highest Spend Day',
      value: DOW_LABELS[maxDow],
      sub: `${formatINR(byDow[maxDow])} total · ${cntDow[maxDow]} transactions`,
      color: '#3B82F6',
    });
  }

  // Weekend vs weekday daily avg
  const wkndExp = expenses.filter(e => [0,6].includes(new Date(e.date + 'T00:00:00').getDay()));
  const wkdyExp = expenses.filter(e => ![0,6].includes(new Date(e.date + 'T00:00:00').getDay()));
  const wkndDays = Math.max(new Set(wkndExp.map(e => e.date)).size, 1);
  const wkdyDays = Math.max(new Set(wkdyExp.map(e => e.date)).size, 1);
  const wkndAvg  = wkndExp.reduce((s,e) => s + e.amount, 0) / wkndDays;
  const wkdyAvg  = wkdyExp.reduce((s,e) => s + e.amount, 0) / wkdyDays;
  insights.push({
    icon: '⚖️',
    title: 'Weekend vs Weekday',
    value: `${formatINR(wkndAvg)}/day`,
    sub: `Weekend avg · Weekdays: ${formatINR(wkdyAvg)}/day`,
    color: '#8B5CF6',
  });

  // Top 3 category concentration
  const byCat = {};
  expenses.forEach(e => { byCat[e.category] = (byCat[e.category] || 0) + e.amount; });
  const catTotal = Object.values(byCat).reduce((s,v) => s + v, 0);
  const top3sum  = Object.values(byCat).sort((a,b) => b-a).slice(0,3).reduce((s,v) => s + v, 0);
  const top3cats = Object.entries(byCat).sort((a,b) => b[1]-a[1]).slice(0,3).map(([k]) => k).join(', ');
  const conc = catTotal > 0 ? (top3sum / catTotal * 100).toFixed(0) : 0;
  insights.push({
    icon: '🎯',
    title: 'Top 3 Concentration',
    value: `${conc}% of spend`,
    sub: top3cats || '—',
    color: '#F59E0B',
  });

  // Avg transaction + outliers
  const avg = expenses.reduce((s,e) => s + e.amount, 0) / expenses.length;
  const big = expenses.filter(e => e.amount > avg * 2).length;
  insights.push({
    icon: '💡',
    title: 'Avg Transaction',
    value: formatINR(avg),
    sub: `${big} transaction${big !== 1 ? 's' : ''} are 2× above average`,
    color: '#14B8A6',
  });

  // Biggest month ever (all-time)
  const byMonth = {};
  state.expenses.forEach(e => { byMonth[e.date.substring(0,7)] = (byMonth[e.date.substring(0,7)] || 0) + e.amount; });
  if (Object.keys(byMonth).length > 0) {
    const top = Object.entries(byMonth).sort((a,b) => b[1]-a[1])[0];
    const [tmy, tmm] = top[0].split('-');
    insights.push({
      icon: '🏔️',
      title: 'Biggest Month Ever',
      value: `${MONTHS_SHORT[parseInt(tmm)-1]} ${tmy}`,
      sub: formatINR(top[1]) + ' spent',
      color: '#EC4899',
    });
  }

  grid.innerHTML = insights.map(i => `
    <div class="smart-insight-card" style="--si-accent:${i.color}">
      <div class="si-icon">${i.icon}</div>
      <div class="si-title">${i.title}</div>
      <div class="si-value">${i.value}</div>
      <div class="si-sub">${i.sub}</div>
    </div>
  `).join('');
}

function getDailyAvg(expenses) {
  if (!expenses.length) return 0;
  const dates = new Set(expenses.map(e => e.date));
  return expenses.reduce((s,e)=>s+e.amount,0) / dates.size;
}

function render6MonthTrend() {
  destroyChart('an6MonthChart');
  const canvas = document.getElementById('an6MonthChart');
  if (!canvas) return;

  const now = new Date();
  const months = [];
  const totals = [];
  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    const label = `${MONTHS_SHORT[d.getMonth()]} ${d.getFullYear()}`;
    const total = state.expenses.filter(e => e.date.startsWith(key)).reduce((s,e) => s + e.amount, 0);
    months.push(label);
    totals.push(total);
  }

  if (totals.every(v => v === 0)) {
    canvas.style.display = 'none';
    canvas.insertAdjacentHTML('afterend', '<div class="chart-empty"><div class="chart-empty-icon">📊</div>No data for the last 6 months</div>');
    return;
  }

  const isDark = document.body.classList.contains('dark');
  const textColor = isDark ? '#a0a8c0' : '#6b7280';
  const gridColor = isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)';

  const ctx = canvas.getContext('2d');
  chartInstances['an6MonthChart'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: months,
      datasets: [{
        data: totals,
        backgroundColor: totals.map((v, i) => {
          const isCurrentMonth = i === 5;
          return isCurrentMonth ? '#2563EB' : 'rgba(37,99,235,0.2)';
        }),
        borderRadius: 8,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: c => ` ${formatINR(c.raw)}` } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { color: textColor, font: { size: 11 } } },
        y: { grid: { color: gridColor }, ticks: { color: textColor, callback: v => '₹' + formatNum(v) }, beginAtZero: true }
      }
    }
  });
}

function renderHeatmap() {
  const container = document.getElementById('heatmapContainer');
  const today = new Date();
  const sixMonthsAgo = new Date(today.getFullYear(), today.getMonth() - 5, 1);

  // Group by date
  const byDate = {};
  state.expenses.forEach(e => { byDate[e.date] = (byDate[e.date] || 0) + e.amount; });

  if (Object.keys(byDate).length === 0) {
    const max = Math.max(...Object.values(byDate), 1);
    container.innerHTML = '<p class="hint">No data for heatmap.</p>'; return;
  }

  const amounts = Object.values(byDate);
  const max = Math.max(...amounts);
  const getLevel = amt => {
    if (!amt) return 0;
    const pct = amt / max;
    if (pct < 0.25) return 1;
    if (pct < 0.5) return 2;
    if (pct < 0.75) return 3;
    return 4;
  };

  // Build by month
  let html = '<div class="heatmap">';
  for (let m = 0; m < 6; m++) {
    const monthDate = new Date(today.getFullYear(), today.getMonth() - 5 + m, 1);
    const monthStr = `${MONTHS_SHORT[monthDate.getMonth()]} ${monthDate.getFullYear()}`;
    const year = monthDate.getFullYear();
    const month = monthDate.getMonth();
    const daysInMonth = new Date(year, month+1, 0).getDate();
    const firstDow = new Date(year, month, 1).getDay();

    html += `<div class="heatmap-month"><div class="heatmap-month-label">${MONTHS_SHORT[month]}</div><div class="heatmap-weeks">`;

    // Group into weeks
    const cells = [];
    for (let blank = 0; blank < firstDow; blank++) cells.push(null);
    for (let d = 1; d <= daysInMonth; d++) cells.push(d);

    const weeks = [];
    for (let i = 0; i < cells.length; i += 7) weeks.push(cells.slice(i, i+7));

    weeks.forEach(week => {
      html += '<div class="heatmap-week">';
      week.forEach(day => {
        if (day === null) { html += '<div class="heatmap-day" style="opacity:0"></div>'; return; }
        const dateStr = `${year}-${String(month+1).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
        const amt = byDate[dateStr] || 0;
        const level = getLevel(amt);
        html += `<div class="heatmap-day" data-level="${level}" title="${dateStr}: ${formatINR(amt)}"></div>`;
      });
      html += '</div>';
    });

    html += '</div></div>';
  }
  html += '</div>';
  container.innerHTML = html;
}

function renderCatMonthChart() {
  const type = document.getElementById('catMonthChartType').value;
  const expenses = state.expenses;
  if (expenses.length === 0) { destroyChart('catMonthChart'); return; }

  const months = [...new Set(expenses.map(e => e.date.substring(0,7)))].sort();
  const cats = [...new Set(expenses.map(e => e.category))];

  const datasets = cats.map((cat, i) => ({
    label: cat,
    data: months.map(m => {
      const filtered = expenses.filter(e => e.category === cat && e.date.startsWith(m));
      return filtered.reduce((s,e)=>s+e.amount,0);
    }),
    backgroundColor: getCategoryColor(cat) + (type === 'bar' ? 'cc' : '33'),
    borderColor: getCategoryColor(cat),
    borderWidth: 2,
    fill: type === 'line',
    tension: 0.4,
    stack: type === 'bar' ? 'stack' : undefined,
    borderRadius: type === 'bar' ? 4 : 0,
  }));

  const isDark = document.body.classList.contains('dark');
  const gridColor = isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)';
  const textColor = isDark ? '#a0a8c0' : '#6b7280';
  const labels = months.map(m => { const [y,mo]=m.split('-'); return `${MONTHS_SHORT[parseInt(mo)-1]} ${y.slice(2)}`; });

  destroyChart('catMonthChart');
  const ctx = document.getElementById('catMonthChart').getContext('2d');
  chartInstances['catMonthChart'] = new Chart(ctx, {
    type,
    data: { labels, datasets },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom', labels: { color: getTextColor(), boxWidth: 12 } },
        tooltip: { callbacks: { label: c => ` ${c.dataset.label}: ${formatINR(c.raw)}` } }
      },
      scales: {
        x: { grid: { color: gridColor }, ticks: { color: textColor } },
        y: {
          stacked: type === 'bar',
          grid: { color: gridColor },
          ticks: { color: textColor, callback: v => '₹'+formatNum(v) },
          beginAtZero: true,
        }
      }
    }
  });
}

function renderTopExpensesChart(expenses) {
  const top10 = [...expenses].sort((a,b)=>b.amount-a.amount).slice(0,10);
  const labels = top10.map(e => (e.description || e.category).substring(0,20));
  const data = top10.map(e => e.amount);
  const colors = top10.map(e => getCategoryColor(e.category));

  const isDark = document.body.classList.contains('dark');
  const gridColor = isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)';
  const textColor = isDark ? '#a0a8c0' : '#6b7280';

  destroyChart('topExpensesChart');
  const wrap = document.getElementById('topExpensesWrap');
  if (data.length === 0) {
    if (wrap) wrap.innerHTML = '<div class="chart-empty"><div class="chart-empty-icon">🏆</div>No expenses to display</div>';
    return;
  }
  if (wrap) wrap.innerHTML = '<canvas id="topExpensesChart"></canvas>';
  const ctx = document.getElementById('topExpensesChart').getContext('2d');
  chartInstances['topExpensesChart'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ data, backgroundColor: colors, borderRadius: 6, borderSkipped: false }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ` ${formatINR(c.raw)}` } } },
      scales: {
        x: { grid: { color: gridColor }, ticks: { color: textColor, callback: v => '₹'+formatNum(v) }, beginAtZero: true },
        y: { grid: { display: false }, ticks: { color: textColor } }
      }
    }
  });
}

function renderPaidByAnalyticsChart(expenses) {
  const byPay = groupBy(expenses, 'paidBy');
  const labels = Object.keys(byPay);
  const data = labels.map(k => byPay[k].reduce((s,e)=>s+e.amount,0));

  destroyChart('paidByAnalyticsChart');
  const wrap = document.getElementById('paidByWrap');
  if (data.length === 0) {
    if (wrap) wrap.innerHTML = '<div class="chart-empty"><div class="chart-empty-icon">💳</div>No payment data</div>';
    return;
  }
  if (wrap) wrap.innerHTML = '<canvas id="paidByAnalyticsChart"></canvas>';
  const ctx = document.getElementById('paidByAnalyticsChart').getContext('2d');
  const bgColor = document.body.classList.contains('dark') ? '#1e2130' : '#ffffff';
  chartInstances['paidByAnalyticsChart'] = new Chart(ctx, {
    type: 'doughnut',
    data: { labels, datasets: [{ data, backgroundColor: PALETTES.cool, borderWidth: 2, borderColor: bgColor }] },
    options: {
      responsive: true,
      cutout: '60%',
      plugins: { legend: { position: 'bottom', labels: { color: getTextColor(), boxWidth: 12, font:{size:11} } } }
    }
  });
}

function renderBudgetBars(expenses) {
  const now = new Date();
  const curMonth = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const monthExp = expenses.filter(e => e.date.startsWith(curMonth));
  const byCat = groupBy(monthExp, 'category');

  const container = document.getElementById('budgetBars');
  if (Object.keys(byCat).length === 0) {
    container.innerHTML = '<p class="hint">No expenses this month.</p>'; return;
  }

  container.innerHTML = Object.entries(byCat).map(([cat, exps]) => {
    const spent = exps.reduce((s,e)=>s+e.amount,0);
    const budget = state.budgets[cat] || 0;
    const pct = budget > 0 ? Math.min((spent/budget)*100, 100) : 0;
    const color = pct >= 90 ? 'var(--danger)' : pct >= 70 ? 'var(--warning)' : getCategoryColor(cat);
    return `
      <div class="budget-item">
        <div class="budget-header">
          <span class="budget-name">${getCategoryEmoji(cat)} ${cat}</span>
          <span class="budget-info">${formatINR(spent)} ${budget > 0 ? `/ ${formatINR(budget)} (${pct.toFixed(0)}%)` : '(no budget set)'}</span>
        </div>
        ${budget > 0 ? `
          <div class="budget-bar-bg">
            <div class="budget-bar-fill" style="width:${pct}%;background:${color}"></div>
          </div>
        ` : ''}
      </div>
    `;
  }).join('');
}

function openBudgetModal() {
  const now = new Date();
  const curMonth = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const monthExp = getFilteredExpenses().filter(e => e.date.startsWith(curMonth));
  const cats = [...new Set(monthExp.map(e => e.category))];

  document.getElementById('budgetModalContent').innerHTML = `
    <div class="form-grid" style="margin-bottom:1rem">
      ${state.categories.map(cat => `
        <div class="form-group">
          <label>${getCategoryEmoji(cat)} ${cat}</label>
          <input type="number" class="input" id="budget_${cat.replace(/\s/g,'_')}"
            value="${state.budgets[cat] || ''}" placeholder="Set budget (₹)" min="0" step="100" />
        </div>
      `).join('')}
    </div>
  `;
  document.getElementById('budgetModal').style.display = 'flex';
}

function saveBudgets() {
  state.categories.forEach(cat => {
    const el = document.getElementById('budget_' + cat.replace(/\s/g,'_'));
    if (el) {
      const val = parseFloat(el.value);
      if (!isNaN(val) && val > 0) state.budgets[cat] = val;
      else delete state.budgets[cat];
    }
  });
  saveState();
  closeModal('budgetModal');
  renderBudgetBars(getFilteredExpenses());
  showToast('Budgets saved!', 'success');
}

/* ============================================================ CUSTOM CHARTS */
let customChartInstance = null;

function updateCustomChart() {
  const type = document.getElementById('ccType').value;
  const groupBy_ = document.getElementById('ccGroupBy').value;
  const measure = document.getElementById('ccMeasure').value;
  const filterCat = document.getElementById('ccFilterCategory').value;
  const filterNat = document.getElementById('ccFilterNature').value;
  const dateFrom = document.getElementById('ccDateFrom').value;
  const dateTo = document.getElementById('ccDateTo').value;
  const palette = document.getElementById('ccPalette').value;
  const title = document.getElementById('ccTitle').value || 'My Chart';

  let data = [...state.expenses];
  if (filterCat) data = data.filter(e => e.category === filterCat);
  if (filterNat) data = data.filter(e => e.nature === filterNat);
  if (dateFrom) data = data.filter(e => e.date >= dateFrom);
  if (dateTo) data = data.filter(e => e.date <= dateTo);

  if (data.length === 0) {
    document.getElementById('customChartHint').textContent = 'No data matches your filters.';
    if (customChartInstance) { customChartInstance.destroy(); customChartInstance = null; }
    return;
  }

  // Group
  const grouped = {};
  data.forEach(e => {
    let key;
    switch(groupBy_) {
      case 'category': key = e.category; break;
      case 'nature': key = e.nature; break;
      case 'paidBy': key = e.paidBy; break;
      case 'month': { const [y,m]=e.date.split('-'); key = `${MONTHS_SHORT[parseInt(m)-1]} ${y}`; break; }
      case 'week': { const d=new Date(e.date); const wk=getWeekNumber(d); key = `W${wk} ${d.getFullYear()}`; break; }
      case 'dow': key = DOW_LABELS[new Date(e.date).getDay()]; break;
      case 'description': key = e.description || e.category; break;
      default: key = e.category;
    }
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(e.amount);
  });

  const colors = PALETTES[palette] || PALETTES.vibrant;
  const labels = Object.keys(grouped);
  const values = labels.map(k => {
    const vals = grouped[k];
    switch(measure) {
      case 'sum': return vals.reduce((a,b)=>a+b,0);
      case 'count': return vals.length;
      case 'avg': return vals.reduce((a,b)=>a+b,0)/vals.length;
      case 'max': return Math.max(...vals);
      default: return vals.reduce((a,b)=>a+b,0);
    }
  });

  const bgColors = labels.map((_,i) => colors[i % colors.length]);
  const isDark = document.body.classList.contains('dark');
  const gridColor = isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)';
  const textColor = isDark ? '#a0a8c0' : '#6b7280';

  const isPolar = ['pie','doughnut','polarArea'].includes(type);

  if (customChartInstance) { customChartInstance.destroy(); customChartInstance = null; }
  document.getElementById('customChartHint').textContent = '';
  const ctx = document.getElementById('customChart').getContext('2d');

  customChartInstance = new Chart(ctx, {
    type,
    data: {
      labels,
      datasets: [{
        label: title,
        data: values,
        backgroundColor: isPolar ? bgColors : bgColors.map(c=>c+'cc'),
        borderColor: isPolar ? bgColors.map(c=>c+'ff') : bgColors,
        borderWidth: isPolar ? 2 : 1.5,
        fill: type === 'line' ? true : undefined,
        tension: 0.4,
        borderRadius: type === 'bar' ? 6 : 0,
        pointRadius: type === 'line' ? 5 : undefined,
        pointHoverRadius: type === 'line' ? 7 : undefined,
      }]
    },
    options: {
      responsive: true,
      plugins: {
        title: { display: !!title, text: title, color: textColor, font: { size: 14, weight: 'bold' } },
        legend: { position: 'bottom', labels: { color: textColor, boxWidth: 12 } },
        tooltip: { callbacks: {
          label: c => measure === 'count' ? ` ${c.label}: ${c.raw} items` : ` ${c.label}: ${formatINR(c.raw)}`
        }}
      },
      scales: isPolar ? undefined : {
        x: { grid: { color: gridColor }, ticks: { color: textColor } },
        y: {
          grid: { color: gridColor },
          ticks: {
            color: textColor,
            callback: v => measure === 'count' ? v : '₹'+formatNum(v)
          },
          beginAtZero: true,
        }
      },
      cutout: type === 'doughnut' ? '55%' : undefined,
    }
  });
}

function saveCustomChart() {
  if (!customChartInstance) { showToast('Build a chart first!', 'warning'); return; }
  const title = document.getElementById('ccTitle').value || `Chart ${state.savedCharts.length+1}`;
  const config = {
    id: generateId(),
    title,
    type: document.getElementById('ccType').value,
    groupBy: document.getElementById('ccGroupBy').value,
    measure: document.getElementById('ccMeasure').value,
    filterCat: document.getElementById('ccFilterCategory').value,
    filterNat: document.getElementById('ccFilterNature').value,
    dateFrom: document.getElementById('ccDateFrom').value,
    dateTo: document.getElementById('ccDateTo').value,
    palette: document.getElementById('ccPalette').value,
    savedAt: Date.now(),
  };
  state.savedCharts.push(config);
  saveState();
  renderSavedCharts();
  showToast('Chart saved!', 'success');
}

function downloadCustomChart() {
  if (!customChartInstance) { showToast('Build a chart first!', 'warning'); return; }
  const canvas = document.getElementById('customChart');
  const a = document.createElement('a');
  a.href = canvas.toDataURL('image/png');
  a.download = (document.getElementById('ccTitle').value || 'chart') + '.png';
  a.click();
  showToast('Chart downloaded!', 'success');
}

function renderSavedCharts() {
  const section = document.getElementById('savedChartsSection');
  if (state.savedCharts.length === 0) { section.style.display = 'none'; return; }
  section.style.display = 'block';

  const grid = document.getElementById('savedChartsGrid');
  grid.innerHTML = state.savedCharts.map(cfg => `
    <div class="saved-chart-item">
      <div class="saved-chart-header">
        <span class="saved-chart-title">${cfg.title}</span>
        <div class="saved-chart-actions">
          <button class="btn btn-sm btn-outline" onclick="loadSavedChart('${cfg.id}')">Load</button>
          <button class="btn-icon btn-delete" onclick="deleteSavedChart('${cfg.id}')">🗑️</button>
        </div>
      </div>
      <p class="hint">${cfg.type} &bull; By ${cfg.groupBy} &bull; ${cfg.measure}</p>
    </div>
  `).join('');
}

function loadSavedChart(id) {
  const cfg = state.savedCharts.find(c => c.id === id);
  if (!cfg) return;
  document.getElementById('ccType').value = cfg.type;
  document.getElementById('ccGroupBy').value = cfg.groupBy;
  document.getElementById('ccMeasure').value = cfg.measure;
  document.getElementById('ccFilterCategory').value = cfg.filterCat || '';
  document.getElementById('ccFilterNature').value = cfg.filterNat || '';
  document.getElementById('ccDateFrom').value = cfg.dateFrom || '';
  document.getElementById('ccDateTo').value = cfg.dateTo || '';
  document.getElementById('ccPalette').value = cfg.palette;
  document.getElementById('ccTitle').value = cfg.title;
  updateCustomChart();
}

function deleteSavedChart(id) {
  state.savedCharts = state.savedCharts.filter(c => c.id !== id);
  saveState();
  renderSavedCharts();
}

/* ============================================================ SETTINGS */
function renderSettings() {
  renderTagList('categoriesList', state.categories, 'category');
  renderTagList('natureTagsList', state.nature, 'nature');
  renderTagList('paidByList', state.paidBy, 'paidby');
  // Sync font size buttons
  const saved = parseInt(localStorage.getItem('rupaiya_font_size'), 10) || 16;
  syncFontSizeBtns(saved);
}

function renderTagList(containerId, items, type) {
  document.getElementById(containerId).innerHTML = items.map(item => `
    <span class="tag">
      ${item}
      <button class="tag-remove" onclick="removeTag('${type}','${item.replace(/'/g,"\\'")}')" title="Remove">✕</button>
    </span>
  `).join('');
}

function removeTag(type, value) {
  if (type === 'category') {
    if (state.expenses.some(e => e.category === value)) {
      if (!confirm(`Some expenses use "${value}". Remove anyway?`)) return;
    }
    state.categories = state.categories.filter(c => c !== value);
  } else if (type === 'nature') {
    state.nature = state.nature.filter(n => n !== value);
  } else if (type === 'paidby') {
    state.paidBy = state.paidBy.filter(p => p !== value);
  }
  saveState();
  populateSelects();
  renderSettings();
}

function addCategory() {
  const val = document.getElementById('newCategory').value.trim();
  if (!val) return;
  if (state.categories.includes(val)) { showToast('Already exists!', 'warning'); return; }
  state.categories.push(val);
  saveState(); populateSelects(); renderSettings();
  document.getElementById('newCategory').value = '';
  showToast(`Category "${val}" added!`, 'success');
}

function addNature() {
  const val = document.getElementById('newNature').value.trim();
  if (!val) return;
  if (state.nature.includes(val)) { showToast('Already exists!', 'warning'); return; }
  state.nature.push(val);
  saveState(); populateSelects(); renderSettings();
  document.getElementById('newNature').value = '';
  showToast(`Nature "${val}" added!`, 'success');
}

function addPaidBy() {
  const val = document.getElementById('newPaidBy').value.trim();
  if (!val) return;
  if (state.paidBy.includes(val)) { showToast('Already exists!', 'warning'); return; }
  state.paidBy.push(val);
  saveState(); populateSelects(); renderSettings();
  document.getElementById('newPaidBy').value = '';
  showToast(`Payment method "${val}" added!`, 'success');
}

/* ============================================================ SELECT POPULATION */
function populateSelects() {
  // Category + Nature use datalist (free text + suggestions)
  const catList = document.getElementById('categoryList');
  const natList = document.getElementById('natureList');
  if (catList) catList.innerHTML = state.categories.map(c => `<option value="${c}"></option>`).join('');
  if (natList) natList.innerHTML = state.nature.map(n => `<option value="${n}"></option>`).join('');
  populateSelectEl(document.getElementById('expPaidBy'), state.paidBy, '', 'Select Payment Method');

  // Filter selects on expenses page
  const filterCat = document.getElementById('filterCategory');
  const filterNat = document.getElementById('filterNature');
  const filterPay = document.getElementById('filterPaidBy');
  if (filterCat) populateSelectEl(filterCat, state.categories, '', 'All Categories');
  if (filterNat) populateSelectEl(filterNat, state.nature, '', 'All Nature');
  if (filterPay) populateSelectEl(filterPay, state.paidBy, '', 'All Paid By');

  // Custom chart filters
  const ccFilterCat = document.getElementById('ccFilterCategory');
  const ccFilterNat = document.getElementById('ccFilterNature');
  if (ccFilterCat) populateSelectEl(ccFilterCat, state.categories, '', 'All');
  if (ccFilterNat) populateSelectEl(ccFilterNat, state.nature, '', 'All');
}

function populateSelectEl(el, items, selectedValue, placeholder) {
  if (!el) return;
  const current = selectedValue !== undefined ? selectedValue : el.value;
  const ph = placeholder || '';
  el.innerHTML = (ph ? `<option value="">${ph}</option>` : '') +
    items.map(i => `<option value="${i}" ${i===current?'selected':''}>${i}</option>`).join('');
}

/* ============================================================ HELPERS */
function formatINR(n) {
  if (isNaN(n)) return '₹0';
  return '₹' + Number(n).toLocaleString('en-IN', { maximumFractionDigits: 0 });
}

function formatNum(n) {
  if (n >= 100000) return (n/100000).toFixed(1) + 'L';
  if (n >= 1000) return (n/1000).toFixed(1) + 'K';
  return n.toFixed(0);
}

function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('en-IN', { day:'2-digit', month:'short', year:'numeric' });
}

function groupBy(arr, key) {
  return arr.reduce((acc, item) => {
    const k = item[key] || 'Unknown';
    if (!acc[k]) acc[k] = [];
    acc[k].push(item);
    return acc;
  }, {});
}

function getCategoryColor(category) {
  const idx = state.categories.indexOf(category);
  return CATEGORY_COLORS[idx >= 0 ? idx % CATEGORY_COLORS.length : 0];
}

function hexAlpha(hex, alpha) {
  const r = parseInt(hex.slice(1,3), 16);
  const g = parseInt(hex.slice(3,5), 16);
  const b = parseInt(hex.slice(5,7), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function getCategoryEmoji(category) {
  return CATEGORY_EMOJIS[category] || '📦';
}

function getTextColor() {
  return document.body.classList.contains('dark') ? '#a0a8c0' : '#6b7280';
}

function getWeekNumber(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  date.setUTCDate(date.getUTCDate() + 4 - (date.getUTCDay()||7));
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(),0,1));
  return Math.ceil((((date - yearStart) / 86400000) + 1)/7);
}

function destroyChart(id) {
  if (chartInstances[id]) {
    chartInstances[id].destroy();
    delete chartInstances[id];
  }
}

/* ============================================================ TOAST */
function showToast(message, type = '') {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.className = `toast ${type} visible`;
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { toast.classList.remove('visible'); }, 3000);
}

/* ============================================================ KEYBOARD SHORTCUTS */
document.addEventListener('keydown', e => {
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') return;
  if (e.key === 'Escape') {
    closeModal('editModal');
    closeModal('budgetModal');
    closeModal('bulkEditModal');
  }
  if (e.key === 'n' && !e.ctrlKey && !e.metaKey) navigate('add');
});
