/**
 * data.js — دار الدواء Analytics
 * المحرك المسؤول عن معالجة البيانات من عدة صفحات
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
  target:    ['الهدف', 'Target', 'target', 'الهدف بالقيمة', 'الهدف قيمة'],
  targetQty: ['الهدف بالكمية', 'Target Qty', 'الهدف كمية'],
};

let RAW_DATA = [];
let FILTERED_DATA = [];
let DISPLAY_MODE = 'value';

function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
        resolve(processAllSheets(wb));
      } catch (err) { reject(err); }
    };
    reader.readAsBinaryString(file);
  });
}

async function fetchExcelFromUrl(url) {
  const res = await fetch(url);
  const buffer = await res.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  return processAllSheets(wb);
}

function processAllSheets(wb) {
  let allRows = [];
  wb.SheetNames.forEach(name => {
    const json = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
    if (json.length) allRows = allRows.concat(normalizeRows(json));
  });
  return allRows.map((r, i) => ({ ...r, id: i }));
}

function normalizeRows(rows) {
  const headers = Object.keys(rows[0]);
  const map = {};
  for (const [key, aliases] of Object.entries(COLUMN_MAP)) {
    const found = headers.find(h => aliases.some(a => a.trim() === h.trim()));
    if (found) map[key] = found;
  }

  return rows.map(row => ({
    rep:       str(row[map.rep]),
    team:      str(row[map.team]),
    area:      str(row[map.area]),
    specialty: str(row[map.specialty]),
    item:      str(row[map.item]),
    date:      parseDate(row[map.date]),
    value:     num(row[map.value]),
    qty:       num(row[map.qty]),
    visits:    num(row[map.visits]),
    target:    num(row[row[map.target] ? map.target : 'الهدف بالقيمة']), 
    targetQty: num(row[map.targetQty]),
  })).filter(r => r.rep || r.team || r.area || r.item);
}

function applyFilters(f) {
  FILTERED_DATA = RAW_DATA.filter(r => {
    if (f.team && r.team !== f.team) return false;
    if (f.area && r.area !== f.area) return false;
    if (f.specialty && r.specialty !== f.specialty) return false;
    if (f.rep && r.rep !== f.rep) return false;
    if (f.item && r.item !== f.item) return false;
    if (f.dateFrom && r.date && r.date < f.dateFrom) return false;
    if (f.dateTo && r.date && r.date > f.dateTo) return false;
    return true;
  });
  return FILTERED_DATA;
}

function computeKPIs(data) {
  const v = data.reduce((s, r) => s + r.value, 0);
  const q = data.reduce((s, r) => s + r.qty, 0);
  const vis = data.reduce((s, r) => s + r.visits, 0);
  const tV = data.reduce((s, r) => s + r.target, 0);
  const tQ = data.reduce((s, r) => s + r.targetQty, 0);
  const reps = unique(data.map(r => r.rep)).filter(Boolean);
  
  const curT = DISPLAY_MODE === 'qty' ? tQ : tV;
  const curA = DISPLAY_MODE === 'qty' ? q : v;

  return {
    totalValue: v, totalQty: q, totalVisits: vis,
    repCount: reps.length,
    avgSales: reps.length ? v / reps.length : 0,
    targetPct: curT ? Math.min((curA / curT) * 100, 150) : 0
  };
}

// الدوال المساعدة (نفسها ستبقى)
function str(v) { return v ? String(v).trim() : ''; }
function num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function unique(arr) { return [...new Set(arr)]; }
function parseDate(v) { 
  if (!v) return null; if (v instanceof Date) return v;
  const d = new Date(v); return isNaN(d.getTime()) ? null : d;
}
function formatNum(n) {
  if (n >= 1000000) return (n/1000000).toFixed(1) + 'M';
  if (n >= 1000) return (n/1000).toFixed(1) + 'K';
  return n.toLocaleString('ar-EG');
}

// الدوال المطلوبة لـ app.js
function computeMonthlySales(data) { /* كود الرسم البياني */ 
    const months = {};
    data.forEach(row => {
        if (!row.date) return;
        const key = `${row.date.getFullYear()}-${String(row.date.getMonth() + 1).padStart(2, '0')}`;
        if (!months[key]) months[key] = { value: 0, qty: 0 };
        months[key].value += row.value;
        months[key].qty += row.qty;
    });
    const keys = Object.keys(months).sort();
    return {
        labels: keys.map(k => {
            const [y, m] = k.split('-');
            return new Date(+y, +m - 1, 1).toLocaleDateString('ar-EG', { month: 'short', year: '2-digit' });
        }),
        values: keys.map(k => months[k].value),
        qtys: keys.map(k => months[k].qty),
    };
}

function computeTeamsData(data) {
    const teams = {};
    data.forEach(row => {
        if (!row.team) return;
        if (!teams[row.team]) teams[row.team] = { value: 0, qty: 0 };
        teams[row.team].value += row.value;
        teams[row.team].qty += row.qty;
    });
    const keys = Object.keys(teams).sort((a, b) => teams[b].value - teams[a].value);
    return {
        labels: keys,
        values: keys.map(k => teams[k].value),
        qtys: keys.map(k => teams[k].qty),
    };
}

function computeRepTable(data) {
    const reps = {};
    data.forEach(row => {
        const k = row.rep || '—';
        if (!reps[k]) reps[k] = { rep: k, team: row.team, area: row.area, specialty: row.specialty, visits: 0, value: 0, target: row.target };
        reps[k].visits += row.visits;
        reps[k].value += row.value;
        if (!reps[k].team && row.team) reps[k].team = row.team;
    });
    return Object.values(reps).sort((a, b) => b.value - a.value);
}

function extractFilterOptions(data) {
    return {
        teams: unique(data.map(r => r.team)).filter(Boolean).sort(),
        areas: unique(data.map(r => r.area)).filter(Boolean).sort(),
        specialties: unique(data.map(r => r.specialty)).filter(Boolean).sort(),
        reps: unique(data.map(r => r.rep)).filter(Boolean).sort(),
        items: unique(data.map(r => r.item)).filter(Boolean).sort(),
    };
}
// ============================================================
// توليد رؤى SWOT تلقائية (Executive Insights)
// ============================================================
function generateInsights(kpis, teamsData) {
  const insights = [];

  // 1. تحليل نسبة تحقيق الهدف
  if (kpis.targetPct >= 90)
    insights.push({ type: 'strength', text: `✅ تحقيق ${kpis.targetPct.toFixed(1)}% من الهدف — أداء ممتاز ومستقر` });
  else if (kpis.targetPct >= 70)
    insights.push({ type: 'warning', text: `⚠ تحقيق ${kpis.targetPct.toFixed(1)}% من الهدف — الأداء يحتاج إلى دفعة بسيطة` });
  else if (kpis.targetPct > 0)
    insights.push({ type: 'weakness', text: `🔴 تحقيق ${kpis.targetPct.toFixed(1)}% فقط — أداء دون المستوى المطلوب، يتطلب تدخل إداري` });

  // 2. تحليل الكثافة الميدانية (الزيارات)
  if (kpis.repCount > 0) {
    const avgVisits = kpis.totalVisits / kpis.repCount;
    if (avgVisits > 20)
      insights.push({ type: 'strength', text: `📋 معدل ${avgVisits.toFixed(0)} زيارة لكل مندوب — تغطية ميدانية قوية جداً` });
    else if (avgVisits > 0 && avgVisits <= 10)
      insights.push({ type: 'opportunity', text: `📋 معدل الزيارات (${avgVisits.toFixed(0)}) منخفض — فرصة لزيادة الكثافة الميدانية` });
  }

  // 3. تحليل الفريق الأفضل
  if (teamsData && teamsData.labels && teamsData.labels.length > 0) {
    const topTeam = teamsData.labels[0];
    insights.push({ type: 'opportunity', text: `🏆 الفريق الأعلى مساهمة حالياً: ${topTeam}` });
  }

  // 4. تحليل المبيعات الإجمالي
  if (kpis.totalValue > 0) {
    insights.push({ type: 'strength', text: `💰 تم رصد تدفق مبيعات إجمالي بقيمة ${formatNum(kpis.totalValue)} د.أ` });
  }

  return insights;
}
