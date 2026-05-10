/**
 * app.js — دار الدواء Analytics
 * المتحكم الرئيسي: الفلاتر، المخططات، الجداول، KPIs، الوضع الليلي
 */

// ============================================================
// المخططات البيانية
// ============================================================
let salesChartInst = null;
let teamsChartInst = null;

// ألوان المخططات
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
  setStatus('⬆ ارفع ملف Excel أو اربط مصدر GitHub لبدء التحليل', 'neutral');
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
    e.target.value = ''; // السماح برفع نفس الملف مرة أخرى
  });

  // زر GitHub
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
    if (!url) { alert('الرجاء إدخال الرابط'); return; }
    document.getElementById('githubModal').classList.remove('open');
    await loadData(() => fetchExcelFromUrl(url), `GitHub: ${url.split('/').pop()}`);
  });

  // Toggle Value/Qty
  document.getElementById('toggleValue').addEventListener('click', () => {
    DISPLAY_MODE = 'value';
    document.getElementById('toggleValue').classList.add('active');
    document.getElementById('toggleQty').classList.remove('active');
    if (FILTERED_DATA.length) renderAll();
  });
  document.getElementById('toggleQty').addEventListener('click', () => {
    DISPLAY_MODE = 'qty';
    document.getElementById('toggleQty').classList.add('active');
    document.getElementById('toggleValue').classList.remove('active');
    if (FILTERED_DATA.length) renderAll();
  });

  // Dark mode
  document.getElementById('darkModeBtn').addEventListener('click', () => {
    document.body.classList.toggle('light');
    document.getElementById('darkModeBtn').textContent =
      document.body.classList.contains('light') ? '🌙' : '☀️';
  });

  // الفلاتر
  ['filterTeam','filterArea','filterSpecialty','filterRep','filterItem','dateFrom','dateTo']
    .forEach(id => document.getElementById(id).addEventListener('change', runFilters));

  document.getElementById('resetFilters').addEventListener('click', resetFilters);

  // بحث في الجدول
  document.getElementById('tableSearch').addEventListener('input', (e) => {
    filterTable(e.target.value);
  });
}

// ============================================================
// تحميل البيانات (رفع ملف أو GitHub)
// ============================================================
async function loadData(fetcher, label) {
  showLoading(true);
  try {
    const rows = await fetcher();
    if (!rows.length) throw new Error('الملف فارغ أو لا يحتوي على بيانات مطابقة');

    RAW_DATA = rows;
    FILTERED_DATA = rows;

    populateFilters(extractFilterOptions(rows));
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
    specialty: document.getElementById('filterSpecialty').value,
    rep:       document.getElementById('filterRep').value,
    item:      document.getElementById('filterItem').value,
    dateFrom:  document.getElementById('dateFrom').value ? new Date(document.getElementById('dateFrom').value) : null,
    dateTo:    document.getElementById('dateTo').value   ? new Date(document.getElementById('dateTo').value)   : null,
  };
  applyFilters(filters);
  renderAll();
}

function resetFilters() {
  ['filterTeam','filterArea','filterSpecialty','filterRep','filterItem'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('dateFrom').value = '2026-01-01';
  document.getElementById('dateTo').value   = '2026-12-31';
  applyFilters({});
  renderAll();
}

// ============================================================
// ملء قوائم الفلاتر
// ============================================================
function populateFilters(opts) {
  fillSelect('filterTeam',      opts.teams,      'جميع الفرق');
  fillSelect('filterArea',      opts.areas,      'جميع المناطق');
  fillSelect('filterSpecialty', opts.specialties,'جميع التخصصات');
  fillSelect('filterRep',       opts.reps,       'جميع المندوبين');
  fillSelect('filterItem',      opts.items,      'جميع الأصناف');
}

function fillSelect(id, items, placeholder) {
  const sel = document.getElementById(id);
  sel.innerHTML = `<option value="">${placeholder}</option>`;
  items.forEach(v => sel.insertAdjacentHTML('beforeend', `<option value="${esc(v)}">${esc(v)}</option>`));
}

// ============================================================
// رسم كل شيء
// ============================================================
function renderAll() {
  const data = FILTERED_DATA;
  const kpis = computeKPIs(data);
  const monthly = computeMonthlySales(data);
  const teams   = computeTeamsData(data);
  const repRows = computeRepTable(data);
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

  animCounter('kpiTotalSales',  isVal ? kpis.totalValue : kpis.totalQty,
    isVal ? (n => formatNum(n) + ' د.أ') : (n => formatNum(n) + ' وحدة'));
  animCounter('kpiTotalVisits', kpis.totalVisits, n => n.toLocaleString('ar-EG'));
  animCounter('kpiReps',        kpis.repCount,    n => n.toLocaleString('ar-EG'));
  animCounter('kpiAvgSales',    kpis.avgSales,    n => formatNum(n) + ' د.أ');

  const pct = kpis.targetPct;
  document.getElementById('kpiTarget').textContent = pct ? pct.toFixed(1) + '%' : '—';
  setTimeout(() => {
    document.getElementById('kpiProgressBar').style.width = Math.min(pct, 100) + '%';
  }, 100);
}

function animCounter(id, target, fmt) {
  const el = document.getElementById(id);
  if (!el) return;
  const start = 0;
  const duration = 800;
  const startTime = performance.now();
  const update = (now) => {
    const t = Math.min((now - startTime) / duration, 1);
    const val = start + (target - start) * easeOut(t);
    el.textContent = fmt(val);
    if (t < 1) requestAnimationFrame(update);
  };
  requestAnimationFrame(update);
}
function easeOut(t) { return 1 - Math.pow(1 - t, 3); }

// ============================================================
// مخطط المبيعات الشهري
// ============================================================
function initCharts() {
  const isDark = !document.body.classList.contains('light');
  const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';
  const textColor = isDark ? '#8b949e' : '#4a5568';

  Chart.defaults.color = textColor;
  Chart.defaults.font.family = 'Almarai';

  const salesCtx = document.getElementById('salesChart').getContext('2d');
  salesChartInst = new Chart(salesCtx, {
    type: 'bar',
    data: { labels: [], datasets: [] },
    options: {
      responsive: true, maintainAspectRatio: true,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { color: gridColor }, ticks: { color: textColor } },
        y: { grid: { color: gridColor }, ticks: { color: textColor, callback: v => formatNum(v) } },
      },
    },
  });

  const teamsCtx = document.getElementById('teamsChart').getContext('2d');
  teamsChartInst = new Chart(teamsCtx, {
    type: 'doughnut',
    data: { labels: [], datasets: [] },
    options: {
      responsive: true, maintainAspectRatio: true,
      plugins: {
        legend: { position: 'bottom', labels: { color: textColor, font: { family: 'Almarai', size: 11 } } },
        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${formatNum(ctx.raw)} د.أ` } },
      },
      cutout: '62%',
    },
  });
}

function renderSalesChart(monthly) {
  if (!salesChartInst) return;
  const vals = DISPLAY_MODE === 'value' ? monthly.values : monthly.qtys;
  const label = DISPLAY_MODE === 'value' ? 'المبيعات (د.أ)' : 'الكميات';

  salesChartInst.data.labels = monthly.labels;
  salesChartInst.data.datasets = [{
    label,
    data: vals,
    backgroundColor: CHART_COLORS[0] + '99',
    borderColor: CHART_COLORS[0],
    borderWidth: 2,
    borderRadius: 6,
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
    borderWidth: 2,
    borderColor: document.body.classList.contains('light') ? '#fff' : '#161b22',
  }];
  teamsChartInst.update();
}

// ============================================================
// الجدول
// ============================================================
let ALL_ROWS = [];

function renderTable(rows) {
  ALL_ROWS = rows;
  drawTable(rows);
}

function drawTable(rows) {
  const tbody = document.getElementById('visitsTableBody');
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="table-empty">لا توجد نتائج للفلاتر المحددة</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map(r => {
    const pct = r.target ? (r.value / r.target * 100) : null;
    const badge = pct === null ? '' :
      pct >= 90 ? `<span class="badge badge-green">${pct.toFixed(0)}%</span>` :
      pct >= 70 ? `<span class="badge badge-yellow">${pct.toFixed(0)}%</span>` :
                  `<span class="badge badge-red">${pct.toFixed(0)}%</span>`;

    return `<tr>
      <td>${esc(r.rep)}</td>
      <td>${esc(r.team)}</td>
      <td>${esc(r.area)}</td>
      <td>${esc(r.specialty)}</td>
      <td>${r.visits.toLocaleString('ar-EG')}</td>
      <td>${formatNum(r.value)} ${r.value ? 'د.أ' : ''}</td>
      <td>${badge || '—'}</td>
    </tr>`;
  }).join('');
}

function filterTable(query) {
  if (!query) { drawTable(ALL_ROWS); return; }
  const q = query.toLowerCase();
  drawTable(ALL_ROWS.filter(r =>
    [r.rep, r.team, r.area, r.specialty].some(v => v && v.toLowerCase().includes(q))
  ));
}

// ============================================================
// رؤى SWOT
// ============================================================
function renderInsights(insights) {
  const list = document.getElementById('swotList');
  if (!insights.length) {
    list.innerHTML = '<li class="swot-empty">لا توجد بيانات كافية لتوليد رؤى</li>';
    return;
  }
  list.innerHTML = insights.map(i =>
    `<li class="${i.type}">${i.text}</li>`
  ).join('');
}

// ============================================================
// مساعدات UI
// ============================================================
function setStatus(msg, type = 'neutral') {
  const bar = document.getElementById('statusBar');
  document.getElementById('statusText').textContent = msg;
  bar.className = 'status-bar ' + (type === 'success' ? 'success' : type === 'error' ? 'error' : '');
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

function esc(str) {
  if (!str) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
// دالة لتحديث الفلاتر بناءً على البيانات المختارة حالياً
function updateSmartFilters(data) {
    const currentFilters = {
        team: document.getElementById('filterTeam').value,
        area: document.getElementById('filterArea').value,
        // ... بقية الفلاتر
    };

    // 1. استخراج الخيارات المتاحة بناءً على ما هو مفلتر حالياً
    const availableOptions = extractFilterOptions(data);

    // 2. تحديث قائمة "المندوب" مثلاً لتظهر فقط مناديب الفريق المختار
    populateSelect('filterRep', availableOptions.reps, currentFilters.rep);
    populateSelect('filterArea', availableOptions.areas, currentFilters.area);
    // وهكذا للبقية...
}
