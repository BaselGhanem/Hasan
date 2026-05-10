/**
 * app.js — دار الدواء Analytics
 * المتحكم الرئيسي: الفلاتر الذكية، المخططات، الجداول، KPIs، والتبديل بين الصفحات
 */

// ============================================================
// المتغيرات العامة والمخططات
// ============================================================
let salesChartInst = null;
let teamsChartInst = null;
let ALL_TABLE_ROWS = []; // لتخزين بيانات الجدول للبحث السريع

const CHART_COLORS = [
  '#00c896', '#3d8bff', '#f97316', '#a855f7',
  '#ec4899', '#22c55e', '#eab308', '#06b6d4',
];

// ============================================================
// تهيئة التطبيق عند تحميل الصفحة
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  initCharts();
  bindEvents();
  setStatus('⬆ ارفع ملف Excel لبدء التحليل الذكي', 'neutral');
});

// ============================================================
// ربط الأحداث (Event Bindings)
// ============================================================
function bindEvents() {
  // رفع ملف Excel
  document.getElementById('excelFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    await loadData(() => parseExcelFile(file), `تم تحميل: ${file.name}`);
    e.target.value = ''; 
  });

  // مزامنة GitHub
  document.getElementById('githubSyncBtn')?.addEventListener('click', () => {
    document.getElementById('githubModal').classList.add('open');
  });

  // التبديل بين القيمة والكمية
  document.getElementById('toggleValue').addEventListener('click', () => switchDisplayMode('value'));
  document.getElementById('toggleQty').addEventListener('click', () => switchDisplayMode('qty'));

  // الوضع الليلي
  document.getElementById('darkModeBtn').addEventListener('click', () => {
    document.body.classList.toggle('light');
    document.getElementById('darkModeBtn').textContent =
      document.body.classList.contains('light') ? '🌙' : '☀️';
  });

  // الفلاتر الذكية (استدعاء runFilters عند أي تغيير)
  ['filterTeam', 'filterArea', 'filterRep', 'filterItem', 'filterSpecialty', 'dateFrom', 'dateTo']
    .forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', runFilters);
    });

  document.getElementById('resetFilters').addEventListener('click', resetFilters);

  // البحث في الجدول
  document.getElementById('tableSearch')?.addEventListener('input', (e) => {
    filterTable(e.target.value);
  });
}

// ============================================================
// تحميل البيانات
// ============================================================
async function loadData(fetcher, label) {
  showLoading(true);
  try {
    const rows = await fetcher();
    if (!rows.length) throw new Error('الملف فارغ أو غير متوافق');

    RAW_DATA = rows;
    FILTERED_DATA = rows;

    // ملء الفلاتر لأول مرة بناءً على كل البيانات
    const initialOptions = extractFilterOptions(RAW_DATA);
    updateFilterSelects(initialOptions);
    
    renderAll();
    setStatus(`✅ ${label} — جاهز للتحليل`, 'success');
  } catch (err) {
    setStatus(`❌ خطأ: ${err.message}`, 'error');
    console.error(err);
  } finally {
    showLoading(false);
  }
}

// ============================================================
// الفلترة الذكية (Cascading Filters)
// ============================================================
function runFilters() {
  const filters = {
    team:      document.getElementById('filterTeam').value,
    area:      document.getElementById('filterArea').value,
    specialty: document.getElementById('filterSpecialty')?.value || '',
    rep:       document.getElementById('filterRep').value,
    item:      document.getElementById('filterItem').value,
    dateFrom:  document.getElementById('dateFrom').value ? new Date(document.getElementById('dateFrom').value) : null,
    dateTo:    document.getElementById('dateTo').value ? new Date(document.getElementById('dateTo').value) : null,
  };

  // 1. تطبيق الفلترة على البيانات الخام
  applyFilters(filters);

  // 2. تحديث قوائم الفلاتر بناءً على المتاح حالياً (الذكاء المتقاطع)
  const currentOptions = extractFilterOptions(FILTERED_DATA);
  updateFilterSelects(currentOptions, filters);

  renderAll();
}

function updateFilterSelects(opts, current = {}) {
  fillSmartSelect('filterTeam',      opts.teams,       current.team,      'جميع الفرق');
  fillSmartSelect('filterArea',      opts.areas,       current.area,      'جميع المناطق');
  fillSmartSelect('filterSpecialty', opts.specialties, current.specialty, 'جميع التخصصات');
  fillSmartSelect('filterRep',       opts.reps,        current.rep,       'جميع المندوبين');
  fillSmartSelect('filterItem',      opts.items,       current.item,      'جميع الأصناف');
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

function resetFilters() {
  ['filterTeam', 'filterArea', 'filterSpecialty', 'filterRep', 'filterItem'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  document.getElementById('dateFrom').value = '2026-01-01';
  document.getElementById('dateTo').value = '2026-12-31';
  
  applyFilters({});
  updateFilterSelects(extractFilterOptions(RAW_DATA));
  renderAll();
}

// ============================================================
// الرسوم البيانية (Charts)
// ============================================================
function initCharts() {
  // تسجيل الملحق عالمياً
  Chart.register(ChartDataLabels);

  const isDark = !document.body.classList.contains('light');
  const textColor = isDark ? '#8b949e' : '#4a5568';

  const salesCtx = document.getElementById('salesChart')?.getContext('2d');
  if (salesCtx) {
    salesChartInst = new Chart(salesCtx, {
      type: 'bar',
      data: { labels: [], datasets: [] },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          // إعدادات ملصقات البيانات
          datalabels: {
            anchor: 'end',
            align: 'top',
            color: textColor,
            font: { weight: 'bold', family: 'Almarai' },
            formatter: (value) => formatNum(value)
          }
        },
        scales: {
          y: { 
            beginAtZero: true,
            ticks: { callback: v => formatNum(v) } 
          }
        }
      }
    });
  }
  
  // قم بعمل نفس الشيء لمخطط الفرق (teamsChart) مع تغيير الموقع ليكون 'center'
}
  const teamsCtx = document.getElementById('teamsChart')?.getContext('2d');
  if (teamsCtx) {
    teamsChartInst = new Chart(teamsCtx, {
      type: 'doughnut',
      data: { labels: [], datasets: [] },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { position: 'bottom' } },
        cutout: '65%'
      }
    });
  }
}

function renderSalesChart(monthly) {
  if (!salesChartInst) return;
  const vals = DISPLAY_MODE === 'value' ? monthly.values : monthly.qtys;
  salesChartInst.data.labels = monthly.labels;
  salesChartInst.data.datasets = [{
    data: vals,
    backgroundColor: CHART_COLORS[0] + '99',
    borderColor: CHART_COLORS[0],
    borderWidth: 2,
    borderRadius: 5
  }];
  salesChartInst.update();
}

function renderTeamsChart(teams) {
  if (!teamsChartInst) return;
  const vals = DISPLAY_MODE === 'value' ? teams.values : teams.qtys;
  teamsChartInst.data.labels = teams.labels;
  teamsChartInst.data.datasets = [{
    data: vals,
    backgroundColor: CHART_COLORS.slice(0, teams.labels.length),
    borderColor: document.body.classList.contains('light') ? '#fff' : '#161b22',
    borderWidth: 2
  }];
  teamsChartInst.update();
}

// ============================================================
// العرض النهائي (Rendering)
// ============================================================
function renderAll() {
  const data = FILTERED_DATA;
  const kpis = computeKPIs(data);
  const monthly = computeMonthlySales(data);
  const teams = computeTeamsData(data);
  const repRows = computeRepTable(data);
  const insights = generateInsights(kpis, teams);

  renderKPIs(kpis);
  renderSalesChart(monthly);
  renderTeamsChart(teams);
  renderTable(repRows);
  renderInsights(insights);
}

function renderKPIs(kpis) {
  const isVal = DISPLAY_MODE === 'value';
  animCounter('kpiTotalSales',  isVal ? kpis.totalValue : kpis.totalQty, n => formatNum(n) + (isVal ? ' د.أ' : ' وحدة'));
  animCounter('kpiTotalVisits', kpis.totalVisits, n => n.toLocaleString('ar-EG'));
  animCounter('kpiReps',        kpis.repCount,    n => n.toLocaleString('ar-EG'));
  if (document.getElementById('kpiAvgSales')) {
     animCounter('kpiAvgSales', kpis.avgSales, n => formatNum(n) + ' د.أ');
  }

  const pct = kpis.targetPct;
  const targetEl = document.getElementById('kpiTarget');
  if (targetEl) targetEl.textContent = pct ? pct.toFixed(1) + '%' : '0%';
  
  const bar = document.getElementById('kpiProgressBar');
  if (bar) {
    setTimeout(() => { bar.style.width = Math.min(pct, 100) + '%'; }, 100);
  }
}

function renderTable(rows) {
  ALL_TABLE_ROWS = rows;
  const tbody = document.getElementById('visitsTableBody');
  if (!tbody) return;

  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="table-empty">لا توجد نتائج</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map(r => {
    // حساب الأداء بناءً على الهدف (إذا كان موجوداً)
    const pct = r.target ? (r.value / r.target * 100) : 0;
    const badgeClass = pct >= 90 ? 'badge-green' : pct >= 70 ? 'badge-yellow' : 'badge-red';

    return `<tr>
      <td>${esc(r.rep)}</td>
      <td>${esc(r.team)}</td>
      <td>${esc(r.area)}</td>
      <td>${esc(r.specialty)}</td>
      <td>${r.visits.toLocaleString('ar-EG')}</td>
      <td>${formatNum(r.value)} د.أ</td>
      <td><span class="badge ${badgeClass}">${pct.toFixed(0)}%</span></td>
    </tr>`;
  }).join('');
}

function filterTable(query) {
  const q = query.toLowerCase();
  const filtered = ALL_TABLE_ROWS.filter(r => 
    [r.rep, r.team, r.area, r.specialty].some(v => v.toLowerCase().includes(q))
  );
  renderTable(filtered);
}

function renderInsights(insights) {
  const list = document.getElementById('swotList');
  if (!list) return;
  list.innerHTML = insights.length ? 
    insights.map(i => `<li class="${i.type}">${i.text}</li>`).join('') :
    '<li class="swot-empty">لا توجد رؤى حالياً</li>';
}

// ============================================================
// مساعدات الواجهة (UI Helpers)
// ============================================================
function switchDisplayMode(mode) {
  DISPLAY_MODE = mode;
  document.getElementById('toggleValue').classList.toggle('active', mode === 'value');
  document.getElementById('toggleQty').classList.toggle('active', mode === 'qty');
  if (FILTERED_DATA.length) renderAll();
}

function setStatus(msg, type) {
  const bar = document.getElementById('statusBar');
  const txt = document.getElementById('statusText');
  if (txt) txt.textContent = msg;
  bar.className = `status-bar ${type}`;
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
  if (!el) return;
  let start = 0;
  const duration = 1000;
  const startTime = performance.now();
  const update = (now) => {
    const progress = Math.min((now - startTime) / duration, 1);
    const value = start + (target - start) * (1 - Math.pow(1 - progress, 3)); // Ease Out
    el.textContent = fmt(value);
    if (progress < 1) requestAnimationFrame(update);
  };
  requestAnimationFrame(update);
}

function esc(s) {
  if (!s) return '—';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
