/**
 * app.js — دار الدواء Analytics
 * ✅ نسخة مُصلحة — جميع الأخطاء معالجة
 */

let salesChartInst = null;
let teamsChartInst = null;
let ALL_TABLE_ROWS = [];

const CHART_COLORS = [
  '#00c896', '#3d8bff', '#f97316', '#a855f7',
  '#ec4899', '#22c55e', '#eab308', '#06b6d4'
];

// ============================================================
// تهيئة
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  // ✅ التحقق من وجود ChartDataLabels قبل register
  if (typeof ChartDataLabels !== 'undefined') {
    Chart.register(ChartDataLabels);
  }
  initCharts();
  bindEvents();
});

// ============================================================
// ربط الأحداث
// ============================================================
function bindEvents() {
  // رفع ملف Excel
  document.getElementById('excelFile').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    await loadData(() => parseExcelFile(file), `تم تحميل: ${file.name}`);
    e.target.value = '';
  });

  // ✅ GitHub Sync — ربط المودال
  document.getElementById('githubSyncBtn').addEventListener('click', () => {
    document.getElementById('githubModal').classList.add('open');
  });
  document.getElementById('closeModalBtn').addEventListener('click', () => {
    document.getElementById('githubModal').classList.remove('open');
  });
  document.getElementById('githubModal').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) e.currentTarget.classList.remove('open');
  });
  document.getElementById('loadGithubBtn').addEventListener('click', async () => {
    const url = document.getElementById('githubUrl').value.trim();
    if (!url) { alert('الرجاء إدخال رابط الملف'); return; }
    document.getElementById('githubModal').classList.remove('open');
    await loadData(() => fetchExcelFromUrl(url), `GitHub: ${url.split('/').pop()}`);
  });

  // Toggle Value / Qty
  document.getElementById('toggleValue').addEventListener('click', () => switchDisplayMode('value'));
  document.getElementById('toggleQty').addEventListener('click',  () => switchDisplayMode('qty'));

  // Dark mode
  document.getElementById('darkModeBtn').addEventListener('click', () => {
    document.body.classList.toggle('light');
    document.getElementById('darkModeBtn').textContent =
      document.body.classList.contains('light') ? '🌙' : '☀️';
    // ✅ إعادة رسم المخططات بعد تغيير الثيم
    destroyCharts();
    initCharts();
    if (FILTERED_DATA.length) renderAll();
  });

  // الفلاتر
  ['filterTeam', 'filterArea', 'filterSpecialty', 'filterRep', 'filterItem', 'dateFrom', 'dateTo']
    .forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', runFilters);
    });

  document.getElementById('resetFilters').addEventListener('click', resetFilters);

  // بحث في الجدول
  const searchEl = document.getElementById('tableSearch');
  if (searchEl) searchEl.addEventListener('input', (e) => filterTable(e.target.value));
}

// ============================================================
// تحميل البيانات
// ============================================================
async function loadData(fetcher, label) {
  showLoading(true);
  try {
    const rows = await fetcher();
    FILTERED_DATA = rows; // RAW_DATA يُعيَّن داخل data.js
    updateFilterSelects(extractFilterOptions(RAW_DATA));
    renderAll();
    setStatus(`✅ ${label} — ${rows.length.toLocaleString('ar-EG')} سجل`, 'success');
    document.getElementById('lastUpdate').textContent =
      `آخر تحديث: ${new Date().toLocaleTimeString('ar-EG')}`;
  } catch (err) {
    setStatus(`❌ خطأ: ${err.message}`, 'error');
    console.error(err);
  } finally {
    showLoading(false);
  }
}

// ============================================================
// تشغيل الفلاتر
// ============================================================
function runFilters() {
  const filters = {
    team:      document.getElementById('filterTeam').value,
    area:      document.getElementById('filterArea').value,
    specialty: document.getElementById('filterSpecialty')?.value || '',
    rep:       document.getElementById('filterRep').value,
    item:      document.getElementById('filterItem').value,
    dateFrom:  document.getElementById('dateFrom').value
               ? new Date(document.getElementById('dateFrom').value) : null,
    dateTo:    document.getElementById('dateTo').value
               ? new Date(document.getElementById('dateTo').value)   : null,
  };
  applyFilters(filters);
  // ✅ إعادة بناء الفلاتر مع الحفاظ على القيم المختارة
  updateFilterSelects(extractFilterOptions(FILTERED_DATA), filters);
  renderAll();
}

function resetFilters() {
  ['filterTeam', 'filterArea', 'filterSpecialty', 'filterRep', 'filterItem']
    .forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  document.getElementById('dateFrom').value = '2026-01-01';
  document.getElementById('dateTo').value   = '2026-12-31';
  applyFilters({});
  updateFilterSelects(extractFilterOptions(RAW_DATA));
  renderAll();
}

// ============================================================
// ملء قوائم الفلاتر
// ============================================================
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
  (items || []).forEach(v => {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    if (v === currentVal) opt.selected = true;
    sel.appendChild(opt);
  });
}

// ============================================================
// تهيئة المخططات
// ============================================================
function initCharts() {
  const isDark    = !document.body.classList.contains('light');
  const textColor = isDark ? '#8b949e' : '#4a5568';
  const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.08)';

  Chart.defaults.font.family = 'Almarai';
  Chart.defaults.color       = textColor;

  // ✅ إعداد datalabels آمن — يتفعل فقط لو الـ plugin محمّل
  const dataLabelsPlugin = (typeof ChartDataLabels !== 'undefined') ? {
    datalabels: {
      anchor: 'end', align: 'top',
      color: textColor,
      font: { weight: 'bold', family: 'Almarai', size: 10 },
      formatter: (v) => v > 0 ? formatNum(v) : '',
      clamp: true,
    }
  } : { datalabels: { display: false } };

  // Sales Chart (Bar)
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
          ...dataLabelsPlugin,
        },
        scales: {
          x: {
            grid: { color: gridColor },
            ticks: { color: textColor, maxRotation: 45 },
          },
          y: {
            grid: { color: gridColor },
            ticks: { color: textColor, callback: v => formatNum(v) },
            beginAtZero: true,
          },
        },
      },
    });
  }

  // Teams Chart (Doughnut)
  const teamsCtx = document.getElementById('teamsChart')?.getContext('2d');
  if (teamsCtx) {
    teamsChartInst = new Chart(teamsCtx, {
      type: 'doughnut',
      data: { labels: [], datasets: [] },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '65%',
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: textColor, font: { family: 'Almarai', size: 11 }, padding: 12 },
          },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.label}: ${formatNum(ctx.raw)}`
            }
          },
          // ✅ datalabels مخصص للـ doughnut
          datalabels: (typeof ChartDataLabels !== 'undefined') ? {
            color: '#fff',
            font: { weight: 'bold', size: 10 },
            formatter: (v, ctx) => {
              const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
              const pct = total ? ((v / total) * 100).toFixed(0) : 0;
              return pct > 5 ? pct + '%' : '';
            },
          } : { display: false },
        },
      },
    });
  }
}

function destroyCharts() {
  if (salesChartInst) { salesChartInst.destroy(); salesChartInst = null; }
  if (teamsChartInst) { teamsChartInst.destroy(); teamsChartInst = null; }
}

// ============================================================
// رسم كل شيء
// ============================================================
function renderAll() {
  const data     = FILTERED_DATA;
  const kpis     = computeKPIs(data);
  const monthly  = computeMonthlySales(data);
  const teams    = computeTeamsData(data);
  const repRows  = computeRepTable(data);
  const insights = generateInsights(kpis, teams);

  renderKPIs(kpis);
  renderSalesChart(monthly);
  renderTeamsChart(teams);
  renderTable(repRows);
  renderInsights(insights);
}

// ============================================================
// KPIs
// ============================================================
function renderKPIs(kpis) {
  const isVal = DISPLAY_MODE === 'value';

  // إجمالي المبيعات أو الكميات
  animCounter('kpiTotalSales', isVal ? kpis.totalValue : kpis.totalQty,
    n => formatNum(n) + (isVal ? ' د.أ' : ' وحدة'));

  // KPI الكمية (إن وُجد)
  const kpiQtyEl = document.getElementById('kpiTotalQty');
  if (kpiQtyEl) kpiQtyEl.textContent = formatNum(kpis.totalQty) + ' وحدة';

  // الزيارات والمندوبين
  animCounter('kpiTotalVisits', kpis.totalVisits, n => Math.round(n).toLocaleString('ar-EG'));
  animCounter('kpiReps',        kpis.repCount,    n => Math.round(n).toLocaleString('ar-EG'));

  // متوسط الزيارات
  const avgVisitsEl = document.getElementById('kpiAvgVisits');
  if (avgVisitsEl) {
    const avg = kpis.repCount ? kpis.totalVisits / kpis.repCount : 0;
    avgVisitsEl.textContent = avg.toFixed(1);
  }

  // نسبة الهدف
  const pct = kpis.targetPct || 0;
  const targetEl = document.getElementById('kpiTarget');
  if (targetEl) targetEl.textContent = pct.toFixed(1) + '%';
  setTimeout(() => {
    const bar = document.getElementById('kpiProgressBar');
    if (bar) bar.style.width = Math.min(pct, 100) + '%';
  }, 100);
}

// ============================================================
// مخطط المبيعات الشهري
// ============================================================
function renderSalesChart(monthly) {
  if (!salesChartInst) return;
  const vals  = DISPLAY_MODE === 'value' ? monthly.values : monthly.qtys;
  const label = DISPLAY_MODE === 'value' ? 'المبيعات (د.أ)' : 'الكميات';

  salesChartInst.data.labels = monthly.labels;
  salesChartInst.data.datasets = [{
    label,
    data:            vals,
    backgroundColor: CHART_COLORS[0] + 'cc',
    borderColor:     CHART_COLORS[0],
    borderWidth:     2,
    borderRadius:    6,
  }];
  salesChartInst.update();
}

// ============================================================
// مخطط الفرق
// ============================================================
function renderTeamsChart(teams) {
  if (!teamsChartInst) return;
  // ✅ استخدام qtys المُصلحة من computeTeamsData
  const vals = DISPLAY_MODE === 'value' ? teams.values : teams.qtys;

  teamsChartInst.data.labels = teams.labels;
  teamsChartInst.data.datasets = [{
    data:            vals,
    backgroundColor: CHART_COLORS.slice(0, teams.labels.length),
    borderWidth:     2,
    borderColor:     document.body.classList.contains('light') ? '#fff' : '#161b22',
  }];
  teamsChartInst.update();
}

// ============================================================
// الجدول
// ============================================================
function renderTable(rows) {
  ALL_TABLE_ROWS = rows;
  drawTable(rows);
}

function drawTable(rows) {
  const tbody = document.getElementById('visitsTableBody');
  if (!tbody) return;
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="table-empty">لا توجد نتائج للفلاتر المحددة</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map(r => {
    const pct    = r.target ? (r.value / r.target * 100) : null;
    const badge  = pct === null ? '—'
      : pct >= 90 ? `<span class="badge badge-green">${pct.toFixed(0)}%</span>`
      : pct >= 70 ? `<span class="badge badge-yellow">${pct.toFixed(0)}%</span>`
      :             `<span class="badge badge-red">${pct.toFixed(0)}%</span>`;

    return `<tr>
      <td>${esc(r.rep)}</td>
      <td>${esc(r.team)}</td>
      <td>${esc(r.area)}</td>
      <td>${esc(r.specialty)}</td>
      <td>${Math.round(r.visits).toLocaleString('ar-EG')}</td>
      <td>${formatNum(r.value)} ${r.value ? 'د.أ' : ''}</td>
      <td>${badge}</td>
    </tr>`;
  }).join('');
}

function filterTable(query) {
  if (!query.trim()) { drawTable(ALL_TABLE_ROWS); return; }
  const q = query.toLowerCase();
  drawTable(ALL_TABLE_ROWS.filter(r =>
    [r.rep, r.team, r.area, r.specialty].some(v => v && v.toLowerCase().includes(q))
  ));
}

// ============================================================
// رؤى SWOT
// ============================================================
function renderInsights(insights) {
  const list = document.getElementById('swotList');
  if (!list) return;
  list.innerHTML = insights.length
    ? insights.map(i => `<li class="${i.type}">${i.text}</li>`).join('')
    : '<li class="swot-empty">لا توجد بيانات كافية</li>';
}

// ============================================================
// تبديل وضع العرض
// ============================================================
function switchDisplayMode(mode) {
  DISPLAY_MODE = mode;
  document.getElementById('toggleValue').classList.toggle('active', mode === 'value');
  document.getElementById('toggleQty').classList.toggle('active',  mode === 'qty');
  if (FILTERED_DATA.length) renderAll();
}

// ============================================================
// مساعدات UI
// ============================================================
function setStatus(msg, type = 'neutral') {
  const bar = document.getElementById('statusBar');
  const txt = document.getElementById('statusText');
  if (txt) txt.textContent = msg;
  if (bar) bar.className = 'status-bar' + (type === 'success' ? ' success' : type === 'error' ? ' error' : '');
}

function showLoading(show) {
  let overlay = document.getElementById('loadingOverlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'loadingOverlay';
    overlay.className = 'loading-overlay';
    overlay.innerHTML = '<div class="spinner"></div>';
    document.body.appendChild(overlay);
  }
  overlay.classList.toggle('show', show);
}

function animCounter(id, target, fmt) {
  const el = document.getElementById(id);
  if (!el || !target) { if (el) el.textContent = fmt(0); return; }
  const duration = 700;
  const start    = performance.now();
  const update   = (now) => {
    const t   = Math.min((now - start) / duration, 1);
    const val = target * (1 - Math.pow(1 - t, 3));
    el.textContent = fmt(val);
    if (t < 1) requestAnimationFrame(update);
  };
  requestAnimationFrame(update);
}

function esc(s) {
  if (!s) return '—';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
