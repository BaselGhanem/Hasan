/**
 * data.js — دار الدواء Analytics
 * مسؤول عن: قراءة Excel، معالجة البيانات، الفلترة
 * * تم التعديل لدعم قراءة البيانات من عدة صفحات (Sheets) داخل نفس الملف.
 * (المبيعات، الأهداف، والزيارات في صفحات منفصلة)
 */

const COLUMN_MAP = {
  rep:       ['المندوب', 'Rep', 'rep', 'مندوب'],
  team:      ['الفريق', 'Team', 'team', 'فريق'],
  area:      ['المنطقة', 'Area', 'area', 'منطقة'],
  specialty: ['التخصص', 'Specialty', 'specialty', 'تخصص'],
  item:      ['الصنف', 'Item', 'item', 'صنف', 'المنتج', 'Product'],
  date:      ['التاريخ', 'Date', 'date', 'تاريخ', 'الشهر', 'اليوم'],
  value:     ['القيمة', 'Value', 'value', 'المبيعات', 'Sales'],
  qty:       ['الكمية', 'Qty', 'qty', 'كمية', 'Quantity'],
  visits:    ['الزيارات', 'Visits', 'visits', 'زيارات'],
  target:    ['الهدف', 'Target', 'target', 'هدف'],
};

// البيانات الخام وبعد الفلترة
let RAW_DATA = [];
let FILTERED_DATA = [];
let DISPLAY_MODE = 'value'; // 'value' | 'qty'

// ============================================================
// قراءة ملف Excel من ملف محلي
// ============================================================
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
        const rows = processAllSheets(wb);
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsBinaryString(file);
  });
}

// ============================================================
// قراءة ملف Excel من رابط GitHub Raw
// ============================================================
async function fetchExcelFromUrl(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`فشل التحميل: ${res.status}`);
  const buffer = await res.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  return processAllSheets(wb);
}

// ============================================================
// دمج جميع الصفحات (المبيعات، الزيارات، الأهداف) في مصفوفة واحدة
// ============================================================
function processAllSheets(wb) {
  let allRows = [];
  
  wb.SheetNames.forEach(sheetName => {
    const sheet = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (json.length > 0) {
      allRows = allRows.concat(normalizeRows(json));
    }
  });

  // إعادة ترقيم المعرفات بعد الدمج
  return allRows.map((r, i) => { 
    r.id = i; 
    return r; 
  });
}

// ============================================================
// تطبيع الصفوف — تحويل أسماء الأعمدة إلى مفاتيح موحدة
// ============================================================
function normalizeRows(rows) {
  if (!rows.length) return [];
  const headers = Object.keys(rows[0]);

  // اكتشاف أسماء الأعمدة تلقائياً لكل صفحة
  const map = {};
  for (const [key, aliases] of Object.entries(COLUMN_MAP)) {
    for (const alias of aliases) {
      const found = headers.find(h => h.trim() === alias.trim());
      if (found) { map[key] = found; break; }
    }
  }

  return rows.map((row) => ({
    rep:       str(row[map.rep]),
    team:      str(row[map.team]),
    area:      str(row[map.area]),
    specialty: str(row[map.specialty]),
    item:      str(row[map.item]),
    date:      parseDate(row[map.date]),
    value:     num(row[map.value]),
    qty:       num(row[map.qty]),
    visits:    num(row[map.visits]),
    target:    num(row[map.target]),
  })).filter(r => r.rep || r.team); // تجاهل الصفوف الفارغة
}

// ============================================================
// فلترة البيانات
// ============================================================
function applyFilters(filters = {}) {
  FILTERED_DATA = RAW_DATA.filter(row => {
    if (filters.team      && row.team      !== filters.team)      return false;
    if (filters.area      && row.area      !== filters.area)      return false;
    if (filters.specialty && row.specialty !== filters.specialty) return false;
    if (filters.rep       && row.rep       !== filters.rep)       return false;
    if (filters.item      && row.item      !== filters.item)      return false;
    if (filters.dateFrom  && row.date && row.date < filters.dateFrom) return false;
    if (filters.dateTo    && row.date && row.date > filters.dateTo)   return false;
    return true;
  });
  return FILTERED_DATA;
}

// ============================================================
// حساب KPIs
// ============================================================
function computeKPIs(data) {
  const totalValue   = data.reduce((s, r) => s + r.value, 0);
  const totalQty     = data.reduce((s, r) => s + r.qty, 0);
  const totalVisits  = data.reduce((s, r) => s + r.visits, 0);
  const totalTarget  = data.reduce((s, r) => s + r.target, 0);
  const reps         = unique(data.map(r => r.rep)).filter(Boolean);
  const avgSales     = reps.length ? totalValue / reps.length : 0;
  const targetPct    = totalTarget ? Math.min((totalValue / totalTarget) * 100, 150) : 0;

  return {
    totalValue, totalQty, totalVisits, totalTarget,
    repCount: reps.length,
    avgSales, targetPct,
  };
}

// ============================================================
// بيانات الرسم البياني الشهري
// ============================================================
function computeMonthlySales(data) {
  const months = {};
  data.forEach(row => {
    if (!row.date) return;
    const key = `${row.date.getFullYear()}-${String(row.date.getMonth() + 1).padStart(2, '0')}`;
    if (!months[key]) months[key] = { value: 0, qty: 0 };
    months[key].value += row.value;
    months[key].qty   += row.qty;
  });

  const keys = Object.keys(months).sort();
  return {
    labels: keys.map(k => {
      const [y, m] = k.split('-');
      return new Date(+y, +m - 1, 1).toLocaleDateString('ar-EG', { month: 'short', year: '2-digit' });
    }),
    values: keys.map(k => months[k].value),
    qtys:   keys.map(k => months[k].qty),
  };
}

// ============================================================
// بيانات مساهمة الفرق
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
    qtys:   keys.map(k => teams[k].qty),
  };
}

// ============================================================
// بيانات الجدول (تجميع حسب المندوب)
// ============================================================
function computeRepTable(data) {
  const reps = {};
  data.forEach(row => {
    const k = row.rep || '—';
    if (!reps[k]) reps[k] = { rep: k, team: row.team, area: row.area, specialty: row.specialty, visits: 0, value: 0, target: 0 };
    
    reps[k].visits += row.visits;
    reps[k].value  += row.value;
    reps[k].target += row.target;
    
    // الاحتفاظ بالفريق والمنطقة إذا كانت موجودة في أحد الصفوف وغير موجودة في البقية
    if (!reps[k].team && row.team) reps[k].team = row.team;
    if (!reps[k].area && row.area) reps[k].area = row.area;
    if (!reps[k].specialty && row.specialty) reps[k].specialty = row.specialty;
  });
  return Object.values(reps).sort((a, b) => b.value - a.value);
}

// ============================================================
// توليد رؤى SWOT تلقائية
// ============================================================
function generateInsights(kpis, teamsData) {
  const insights = [];

  if (kpis.targetPct >= 90)
    insights.push({ type: 'strength', text: `✅ تحقيق ${kpis.targetPct.toFixed(1)}% من الهدف — أداء ممتاز` });
  else if (kpis.targetPct >= 70)
    insights.push({ type: 'warning', text: `⚠ تحقيق ${kpis.targetPct.toFixed(1)}% من الهدف — يحتاج متابعة` });
  else if (kpis.targetPct > 0)
    insights.push({ type: 'weakness', text: `🔴 تحقيق ${kpis.targetPct.toFixed(1)}% فقط — أداء دون المستهدف` });

  if (kpis.repCount > 0) {
    const avgVisits = kpis.totalVisits / kpis.repCount;
    if (avgVisits > 20)
      insights.push({ type: 'strength', text: `📋 معدل ${avgVisits.toFixed(0)} زيارة/مندوب — نشاط ميداني مرتفع` });
    else if (avgVisits > 0)
      insights.push({ type: 'opportunity', text: `📋 معدل ${avgVisits.toFixed(0)} زيارة/مندوب — يمكن تعزيز النشاط` });
  }

  if (teamsData.labels.length > 1) {
    const top = teamsData.labels[0];
    insights.push({ type: 'opportunity', text: `🏆 الفريق الأفضل أداءً: ${top}` });
  }

  if (kpis.totalValue > 0)
    insights.push({ type: 'strength', text: `💰 إجمالي المبيعات: ${formatNum(kpis.totalValue)} د.أ` });

  return insights;
}

// ============================================================
// الاستخلاص من البيانات لملء الفلاتر
// ============================================================
function extractFilterOptions(data) {
  return {
    teams:      unique(data.map(r => r.team)).filter(Boolean).sort(),
    areas:      unique(data.map(r => r.area)).filter(Boolean).sort(),
    specialties:unique(data.map(r => r.specialty)).filter(Boolean).sort(),
    reps:       unique(data.map(r => r.rep)).filter(Boolean).sort(),
    items:      unique(data.map(r => r.item)).filter(Boolean).sort(),
  };
}

// ============================================================
// HELPERS
// ============================================================
function str(v)  { return v !== undefined && v !== null ? String(v).trim() : ''; }
function num(v)  { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function unique(arr) { return [...new Set(arr)]; }

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function formatNum(n, compact = true) {
  if (!n && n !== 0) return '—';
  if (compact && n >= 1_000_000) return (n / 1_000_000).toFixed(1) + 'م';
  if (compact && n >= 1_000)     return (n / 1_000).toFixed(1) + 'ك';
  return n.toLocaleString('ar-EG');
}
