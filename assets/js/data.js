// assets/js/data.js
// ✅ نسخة مُصلحة — جميع الأخطاء معالجة

const COLUMN_MAP = {
  rep:       ['المندوب', 'Medreps', 'مندوب', 'Rep'],
  team:      ['الفريق', 'Team', 'فريق'],
  area:      ['المنطقة', 'Area', 'منطقة'],
  specialty: ['التخصص', 'Specialty', 'تخصص'],
  item:      ['الصنف', 'Item Desc', 'Item', 'صنف', 'المنتج'],
  date:      ['التاريخ', 'Date', 'تاريخ', 'الشهر'],
  value:     ['القيمة', 'Value', 'المبيعات', 'Sales'],
  qty:       ['الكمية', 'Qty', 'كمية', 'Quantity'],
  // ✅ إصلاح تايبو: 'يارات' → 'الزيارات'
  visits:    ['الزيارات', 'Visits', 'زيارات'],
  // ✅ إضافة aliases أكثر للهدف
  target:    ['الهدف', 'Target', 'الهدف بالقيمة', 'هدف القيمة', 'الهدف (القيمة)'],
  targetQty: ['الهدف بالكمية', 'Target Qty', 'هدف الكمية', 'الهدف (الكمية)']
};

let RAW_DATA = [];
let FILTERED_DATA = [];
let DISPLAY_MODE = 'value';

// ============================================================
// قراءة ملف Excel محلي — كل الشيتات
// ============================================================
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
        let allRows = [];
        wb.SheetNames.forEach(name => {
          const json = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
          if (json.length) allRows = allRows.concat(normalizeRows(json));
        });
        if (!allRows.length) throw new Error('الملف فارغ أو لا يحتوي بيانات مطابقة للأعمدة المطلوبة');
        RAW_DATA = allRows.map((r, i) => ({ ...r, id: i }));
        resolve(RAW_DATA);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('فشل قراءة الملف'));
    reader.readAsBinaryString(file);
  });
}

// ============================================================
// ✅ تمت إضافة fetchExcelFromUrl المفقودة كلياً
// ============================================================
async function fetchExcelFromUrl(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`فشل التحميل من GitHub: ${res.status} ${res.statusText}`);
  const buffer = await res.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  let allRows = [];
  wb.SheetNames.forEach(name => {
    const json = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
    if (json.length) allRows = allRows.concat(normalizeRows(json));
  });
  if (!allRows.length) throw new Error('الملف فارغ أو الأعمدة غير مطابقة');
  RAW_DATA = allRows.map((r, i) => ({ ...r, id: i }));
  return RAW_DATA;
}

// ============================================================
// تطبيع الصفوف
// ============================================================
function normalizeRows(rows) {
  if (!rows.length) return [];
  const headers = Object.keys(rows[0]);
  const map = {};
  for (const [key, aliases] of Object.entries(COLUMN_MAP)) {
    const found = headers.find(h =>
      aliases.some(a => a.trim().toLowerCase() === h.trim().toLowerCase())
    );
    if (found) map[key] = found;
  }

  return rows.map(row => ({
    rep:       String(row[map.rep]       || '').trim(),
    team:      String(row[map.team]      || '').trim(),
    area:      String(row[map.area]      || '').trim(),
    specialty: String(row[map.specialty] || '').trim(),
    item:      String(row[map.item]      || '').trim(),
    date:      parseDate(row[map.date]),
    value:     parseFloat(row[map.value])     || 0,
    qty:       parseFloat(row[map.qty])       || 0,
    visits:    parseFloat(row[map.visits])    || 0,
    target:    parseFloat(row[map.target])    || 0,
    targetQty: parseFloat(row[map.targetQty]) || 0,
  })).filter(r => r.rep || r.team);
}

// ============================================================
// فلترة البيانات
// ============================================================
function applyFilters(f = {}) {
  FILTERED_DATA = RAW_DATA.filter(r => {
    if (f.team      && r.team      !== f.team)      return false;
    if (f.area      && r.area      !== f.area)       return false;
    if (f.specialty && r.specialty !== f.specialty)  return false;
    if (f.rep       && r.rep       !== f.rep)        return false;
    if (f.item      && r.item      !== f.item)       return false;
    if (f.dateFrom  && r.date && r.date < f.dateFrom) return false;
    if (f.dateTo    && r.date && r.date > f.dateTo)   return false;
    return true;
  });
  return FILTERED_DATA;
}

// ============================================================
// خيارات الفلاتر — ✅ تمت إضافة specialties المفقودة
// ============================================================
function extractFilterOptions(data) {
  const getUniq = (key) => [...new Set(data.map(r => r[key]))].filter(Boolean).sort();
  return {
    teams:       getUniq('team'),
    areas:       getUniq('area'),
    specialties: getUniq('specialty'), // ✅ كانت مفقودة
    reps:        getUniq('rep'),
    items:       getUniq('item'),
  };
}

// ============================================================
// KPIs
// ============================================================
function computeKPIs(data) {
  const v  = data.reduce((s, r) => s + r.value, 0);
  const q  = data.reduce((s, r) => s + r.qty,   0);
  const tV = data.reduce((s, r) => s + r.target,    0);
  const tQ = data.reduce((s, r) => s + r.targetQty, 0);
  const curTarget = DISPLAY_MODE === 'qty' ? tQ : tV;
  const curActual = DISPLAY_MODE === 'qty' ? q  : v;
  const reps = [...new Set(data.map(r => r.rep))].filter(Boolean);
  return {
    totalValue:  v,
    totalQty:    q,
    totalVisits: data.reduce((s, r) => s + r.visits, 0),
    repCount:    reps.length,
    avgSales:    reps.length ? v / reps.length : 0,
    targetPct:   curTarget ? Math.min((curActual / curTarget) * 100, 150) : 0,
  };
}

// ============================================================
// مبيعات شهرية
// ============================================================
function computeMonthlySales(data) {
  const months = {};
  data.forEach(row => {
    if (!row.date || isNaN(row.date)) return;
    const key = `${row.date.getFullYear()}-${String(row.date.getMonth() + 1).padStart(2, '0')}`;
    if (!months[key]) months[key] = { value: 0, qty: 0 };
    months[key].value += row.value;
    months[key].qty   += row.qty;
  });
  const keys = Object.keys(months).sort();
  return {
    labels: keys.map(k => {
      const [y, m] = k.split('-');
      return new Date(+y, +m - 1, 1)
        .toLocaleDateString('ar-EG', { month: 'short', year: '2-digit' });
    }),
    values: keys.map(k => months[k].value),
    qtys:   keys.map(k => months[k].qty),
  };
}

// ============================================================
// ✅ إصلاح computeTeamsData: إضافة qtys المفقودة
// ============================================================
function computeTeamsData(data) {
  const teams = {};
  data.forEach(row => {
    if (!row.team) return;
    if (!teams[row.team]) teams[row.team] = { value: 0, qty: 0 };
    teams[row.team].value += row.value;
    teams[row.team].qty   += row.qty;
  });
  const keys = Object.keys(teams).sort((a, b) => teams[b].value - teams[a].value);
  return {
    labels: keys,
    values: keys.map(k => teams[k].value),
    qtys:   keys.map(k => teams[k].qty),  // ✅ كانت مفقودة
  };
}

// ============================================================
// ✅ إضافة computeRepTable المفقودة من data.js
// ============================================================
function computeRepTable(data) {
  const reps = {};
  data.forEach(row => {
    const k = row.rep || '—';
    if (!reps[k]) {
      reps[k] = {
        rep: k,
        team: row.team,
        area: row.area,
        specialty: row.specialty,
        visits: 0, value: 0, qty: 0, target: 0, targetQty: 0
      };
    }
    reps[k].visits    += row.visits;
    reps[k].value     += row.value;
    reps[k].qty       += row.qty;
    reps[k].target    += row.target;
    reps[k].targetQty += row.targetQty;
  });
  return Object.values(reps).sort((a, b) => b.value - a.value);
}

// ============================================================
// رؤى SWOT تلقائية — ✅ إضافة teamsData parameter
// ============================================================
function generateInsights(kpis, teamsData) {
  const insights = [];

  if (kpis.targetPct >= 100)
    insights.push({ type: 'strength',    text: `✅ تجاوز الهدف — تحقيق ${kpis.targetPct.toFixed(0)}%` });
  else if (kpis.targetPct >= 80)
    insights.push({ type: 'strength',    text: `✅ أداء جيد — تحقيق ${kpis.targetPct.toFixed(0)}% من الهدف` });
  else if (kpis.targetPct >= 60)
    insights.push({ type: 'warning',     text: `⚠ تحقيق ${kpis.targetPct.toFixed(0)}% من الهدف — يحتاج متابعة` });
  else if (kpis.targetPct > 0)
    insights.push({ type: 'weakness',    text: `🔴 تحقيق ${kpis.targetPct.toFixed(0)}% فقط — أداء دون المستهدف` });

  if (kpis.repCount > 0) {
    const avgVisits = kpis.totalVisits / kpis.repCount;
    if (avgVisits >= 20)
      insights.push({ type: 'strength',   text: `📋 معدل ${avgVisits.toFixed(0)} زيارة/مندوب — نشاط ميداني مرتفع` });
    else if (avgVisits > 0)
      insights.push({ type: 'opportunity',text: `📋 معدل ${avgVisits.toFixed(0)} زيارة/مندوب — يمكن تعزيز النشاط` });
  }

  if (teamsData && teamsData.labels.length > 0)
    insights.push({ type: 'opportunity', text: `🏆 الفريق الأفضل: ${teamsData.labels[0]}` });

  if (!insights.length)
    insights.push({ type: 'warning', text: 'لا توجد بيانات هدف — يُنصح بإضافة عمود الهدف في الملف' });

  return insights;
}

// ============================================================
// Helpers
// ============================================================
function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function formatNum(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  if (n >= 1_000_000) return (n / 1_000_000).toFixed(1) + 'M';
  if (n >= 1_000)     return (n / 1_000).toFixed(1) + 'K';
  return Number(n).toLocaleString('ar-EG');
}
