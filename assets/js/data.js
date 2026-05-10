// assets/js/data.js

const COLUMN_MAP = {
  rep:       ['المندوب', 'Medreps', 'مندوب', 'Rep'],
  team:      ['الفريق', 'Team', 'فريق'],
  area:      ['المنطقة', 'Area', 'منطقة'],
  specialty: ['التخصص', 'Specialty', 'تخصص'],
  item:      ['الصنف', 'Item Desc', 'Item', 'صنف'],
  date:      ['التاريخ', 'Date', 'تاريخ', 'الشهر'],
  value:     ['القيمة', 'Value', 'المبيعات', 'Sales'],
  qty:       ['الكمية', 'Qty', 'كمية'],
  visits:    ['الزيارات', 'يارات', 'زيارات'],
  target:    ['الهدف', 'Target', 'دف (القيمة)', 'الهدف بالقيمة'],
  targetQty: ['الهدف بالكمية', 'Target Qty', 'دف (الكمية)']
};

let RAW_DATA = [];
let FILTERED_DATA = [];
let DISPLAY_MODE = 'value';

function parseExcelFile(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
      let allRows = [];
      wb.SheetNames.forEach(name => {
        const json = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
        if (json.length) allRows = allRows.concat(normalizeRows(json));
      });
      RAW_DATA = allRows.map((r, i) => ({ ...r, id: i }));
      resolve(RAW_DATA);
    };
    reader.readAsBinaryString(file);
  });
}

function normalizeRows(rows) {
  const headers = Object.keys(rows[0]);
  const map = {};
  for (const [key, aliases] of Object.entries(COLUMN_MAP)) {
    const found = headers.find(h => aliases.some(a => a.trim() === h.trim()));
    if (found) map[key] = found;
  }
  return rows.map(row => ({
    rep: String(row[map.rep] || '').trim(),
    team: String(row[map.team] || '').trim(),
    area: String(row[map.area] || '').trim(),
    specialty: String(row[map.specialty] || '').trim(),
    item: String(row[map.item] || '').trim(),
    date: row[map.date] instanceof Date ? row[map.date] : new Date(row[map.date]),
    value: parseFloat(row[map.value]) || 0,
    qty: parseFloat(row[map.qty]) || 0,
    visits: parseFloat(row[map.visits]) || 0,
    target: parseFloat(row[map.target]) || 0,
    targetQty: parseFloat(row[map.targetQty]) || 0
  })).filter(r => r.rep || r.team);
}

function applyFilters(f) {
  FILTERED_DATA = RAW_DATA.filter(r => {
    if (f.team && r.team !== f.team) return false;
    if (f.area && r.area !== f.area) return false;
    if (f.rep && r.rep !== f.rep) return false;
    if (f.item && r.item !== f.item) return false;
    if (f.dateFrom && r.date < f.dateFrom) return false;
    if (f.dateTo && r.date > f.dateTo) return false;
    return true;
  });
  return FILTERED_DATA;
}

function extractFilterOptions(data) {
  const getUniq = (key) => [...new Set(data.map(r => r[key]))].filter(Boolean).sort();
  return { teams: getUniq('team'), areas: getUniq('area'), reps: getUniq('rep'), items: getUniq('item') };
}

function computeKPIs(data) {
  const v = data.reduce((s, r) => s + r.value, 0);
  const q = data.reduce((s, r) => s + r.qty, 0);
  const tV = data.reduce((s, r) => s + r.target, 0);
  const tQ = data.reduce((s, r) => s + r.targetQty, 0);
  const curT = DISPLAY_MODE === 'qty' ? tQ : tV;
  const curA = DISPLAY_MODE === 'qty' ? q : v;
  return {
    totalValue: v, totalQty: q, totalVisits: data.reduce((s, r) => s + r.visits, 0),
    repCount: [...new Set(data.map(r => r.rep))].filter(Boolean).length,
    targetPct: curT ? Math.min((curA / curT) * 100, 150) : 0
  };
}

function computeMonthlySales(data) {
  const months = {};
  data.forEach(row => {
    if (!row.date || isNaN(row.date)) return;
    const key = `${row.date.getFullYear()}-${String(row.date.getMonth() + 1).padStart(2, '0')}`;
    if (!months[key]) months[key] = { value: 0, qty: 0 };
    months[key].value += row.value;
    months[key].qty += row.qty;
  });
  const keys = Object.keys(months).sort();
  return {
    labels: keys.map(k => k),
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
  const keys = Object.keys(teams).sort((a,b) => teams[b].value - teams[a].value);
  return { labels: keys, values: keys.map(k => teams[k].value) };
}

function generateInsights(kpis) {
  const insights = [];
  if (kpis.targetPct >= 90) insights.push({ type: 'strength', text: `✅ تحقيق ${kpis.targetPct.toFixed(0)}% من الهدف - أداء ممتاز` });
  else if (kpis.targetPct > 0) insights.push({ type: 'warning', text: `🔴 تحقيق ${kpis.targetPct.toFixed(0)}% فقط - يحتاج متابعة` });
  return insights;
}

function formatNum(n) {
  if (n >= 1000000) return (n/1000000).toFixed(1) + 'M';
  if (n >= 1000) return (n/1000).toFixed(1) + 'K';
  return n.toLocaleString('ar-EG');
}
