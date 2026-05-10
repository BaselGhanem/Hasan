/**
 * app.js — دار الدواء Analytics
 * المتحكم في الرسوم البيانية والفلاتر
 */

let salesChartInst = null;
let teamsChartInst = null;
let ALL_TABLE_ROWS = [];

const CHART_COLORS = ['#00c896', '#3d8bff', '#f97316', '#a855f7', '#ec4899', '#22c55e', '#eab308', '#06b6d4'];

document.addEventListener('DOMContentLoaded', () => {
  // تفعيل ملحق ملصقات البيانات
  Chart.register(ChartDataLabels);
  initCharts();
  bindEvents();
});

function bindEvents() {
  document.getElementById('excelFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    await loadData(() => parseExcelFile(file), `تحميل: ${file.name}`);
    e.target.value = ''; 
  });

  document.getElementById('toggleValue').addEventListener('click', () => switchDisplayMode('value'));
  document.getElementById('toggleQty').addEventListener('click', () => switchDisplayMode('qty'));

  document.getElementById('darkModeBtn').addEventListener('click', () => {
    document.body.classList.toggle('light');
    document.getElementById('darkModeBtn').textContent = document.body.classList.contains('light') ? '🌙' : '☀️';
  });

  ['filterTeam', 'filterArea', 'filterRep', 'filterItem', 'filterSpecialty', 'dateFrom', 'dateTo']
    .forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', runFilters);
    });

  document.getElementById('resetFilters').addEventListener('click', resetFilters);
}

async function loadData(fetcher, label) {
  showLoading(true);
  try {
    const rows = await fetcher();
    RAW_DATA = rows;
    FILTERED_DATA = rows;
    updateFilterSelects(extractFilterOptions(RAW_DATA));
    renderAll();
    setStatus(`✅ ${label}`, 'success');
  } catch (err) {
    setStatus(`❌ خطأ: ${err.message}`, 'error');
  } finally {
    showLoading(false);
  }
}

function runFilters() {
  const filters = {
    team:      document.getElementById('filterTeam').value,
    area:      document.getElementById('filterArea').value,
    rep:       document.getElementById('filterRep').value,
    item:      document.getElementById('filterItem').value,
    specialty: document.getElementById('filterSpecialty').value,
    dateFrom:  document.getElementById('dateFrom').value ? new Date(document.getElementById('dateFrom').value) : null,
    dateTo:    document.getElementById('dateTo').value ? new Date(document.getElementById('dateTo').value) : null,
  };
  applyFilters(filters);
  updateFilterSelects(extractFilterOptions(FILTERED_DATA), filters);
  renderAll();
}

function updateFilterSelects(opts, current = {}) {
  fillSmartSelect('filterTeam', opts.teams, current.team, 'جميع الفرق');
  fillSmartSelect('filterArea', opts.areas, current.area, 'جميع المناطق');
  fillSmartSelect('filterRep', opts.reps, current.rep, 'جميع المندوبين');
  fillSmartSelect('filterItem', opts.items, current.item, 'جميع الأصناف');
  fillSmartSelect('filterSpecialty', opts.specialties, current.specialty, 'جميع التخصصات');
}

function fillSmartSelect(id, items, currentVal, placeholder) {
  const sel = document.getElementById(id);
  if (!sel) return;
  sel.innerHTML = `<option value="">${placeholder}</option>`;
  items.forEach(v => {
    const selected = (v === currentVal) ? 'selected' : '';
    sel.insertAdjacentHTML('beforeend', `<option value="${v}" ${selected}>${v}</option>`);
  });
}

function initCharts() {
  const isDark = !document.body.classList.contains('light');
  const textColor = isDark ? '#8b949e' : '#4a5568';

  const commonOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      datalabels: {
        anchor: 'end',
        align: 'top',
        color: textColor,
        font: { weight: 'bold', family: 'Almarai' },
        formatter: (v) => formatNum(v)
      }
    }
  };

  const salesCtx = document.getElementById('salesChart')?.getContext('2d');
  if (salesCtx) {
    salesChartInst = new Chart(salesCtx, {
      type: 'bar',
      data: { labels: [], datasets: [] },
      options: commonOptions
    });
  }

  const teamsCtx = document.getElementById('teamsChart')?.getContext('2d');
  if (teamsCtx) {
    teamsChartInst = new Chart(teamsCtx, {
      type: 'doughnut',
      data: { labels: [], datasets: [] },
      options: { ...commonOptions, cutout: '65%' }
    });
  }
}

function renderAll() {
  const kpis = computeKPIs(FILTERED_DATA);
  const monthly = computeMonthlySales(FILTERED_DATA);
  const teams = computeTeamsData(FILTERED_DATA);
  const repRows = computeRepTable(FILTERED_DATA);
  const insights = generateInsights(kpis, teams);

  renderKPIs(kpis);
  renderSalesChart(monthly);
  renderTeamsChart(teams);
  renderTable(repRows);
  renderInsights(insights);
}

function renderKPIs(kpis) {
  const isVal = DISPLAY_MODE === 'value';
  animCounter('kpiTotalSales', isVal ? kpis.totalValue : kpis.totalQty, n => formatNum(n));
  animCounter('kpiTotalVisits', kpis.totalVisits, n => n.toLocaleString());
  animCounter('kpiReps', kpis.repCount, n => n.toLocaleString());
  
  const pct = kpis.targetPct;
  document.getElementById('kpiTarget').textContent = pct.toFixed(1) + '%';
  document.getElementById('kpiProgressBar').style.width = Math.min(pct, 100) + '%';
}

function renderTable(rows) {
  ALL_TABLE_ROWS = rows;
  const tbody = document.getElementById('visitsTableBody');
  if (!tbody) return;
  tbody.innerHTML = rows.map(r => `
    <tr>
      <td>${esc(r.rep)}</td>
      <td>${esc(r.team)}</td>
      <td>${esc(r.area)}</td>
      <td>${esc(r.specialty)}</td>
      <td>${r.visits.toLocaleString()}</td>
      <td>${formatNum(r.value)}</td>
      <td><span class="badge ${r.target && (r.value/r.target*100 >= 90) ? 'badge-green' : 'badge-red'}">${r.target ? (r.value/r.target*100).toFixed(0) : 0}%</span></td>
    </tr>
  `).join('');
}

function renderSalesChart(monthly) {
  if (!salesChartInst) return;
  salesChartInst.data.labels = monthly.labels;
  salesChartInst.data.datasets = [{
    data: DISPLAY_MODE === 'value' ? monthly.values : monthly.qtys,
    backgroundColor: CHART_COLORS[0]
  }];
  salesChartInst.update();
}

function renderTeamsChart(teams) {
  if (!teamsChartInst) return;
  teamsChartInst.data.labels = teams.labels;
  teamsChartInst.data.datasets = [{
    data: DISPLAY_MODE === 'value' ? teams.values : teams.qtys,
    backgroundColor: CHART_COLORS
  }];
  teamsChartInst.update();
}

function renderInsights(insights) {
  const list = document.getElementById('swotList');
  if (list) list.innerHTML = insights.map(i => `<li class="${i.type}">${i.text}</li>`).join('');
}

function resetFilters() {
  ['filterTeam', 'filterArea', 'filterRep', 'filterItem'].forEach(id => document.getElementById(id).value = '');
  applyFilters({});
  renderAll();
}

function switchDisplayMode(mode) {
  DISPLAY_MODE = mode;
  document.getElementById('toggleValue').classList.toggle('active', mode === 'value');
  document.getElementById('toggleQty').classList.toggle('active', mode === 'qty');
  renderAll();
}

function setStatus(msg, type) {
  const txt = document.getElementById('statusText');
  if (txt) txt.textContent = msg;
}

function showLoading(show) {
  let loader = document.getElementById('loadingOverlay');
  if (!loader && show) {
    loader = document.createElement('div');
    loader.id = 'loadingOverlay';
    loader.className = 'loading-overlay';
    loader.innerHTML = '<div class="spinner"></div>';
    document.body.appendChild(loader);
  }
  if (loader) loader.classList.toggle('show', show);
}

function animCounter(id, target, fmt) {
  const el = document.getElementById(id);
  if (el) el.textContent = fmt(target);
}

function esc(s) { return s ? String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;') : '—'; }
