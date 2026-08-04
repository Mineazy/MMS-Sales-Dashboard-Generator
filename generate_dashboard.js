const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('kobo_data.xlsx');
const sheet = workbook.Sheets['MMS Sales Tracker'];
const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });

function excelDateToJS(serial) {
  if (!serial || isNaN(serial)) return null;
  return new Date((serial - 25569) * 86400 * 1000);
}

const repSales = {};
const dailySales = {};
const monthlySales = {};
let totalAmount = 0;
let totalInvoices = 0;
let invoiceSet = new Set();

data.forEach(row => {
  const amt = parseFloat(row['Amount']) || 0;
  const rep = row['Sales Rep Number'];
  const date = excelDateToJS(row['Date']);
  const inv = row['Invoice Number'];
  
  totalAmount += amt;
  totalInvoices++;
  if (inv) invoiceSet.add(inv);
  
  if (rep) {
    if (!repSales[rep]) repSales[rep] = { amount: 0, count: 0 };
    repSales[rep].amount += amt;
    repSales[rep].count++;
  }
  
  if (date) {
    const key = date.toISOString().split('T')[0];
    if (!dailySales[key]) dailySales[key] = { amount: 0, count: 0 };
    dailySales[key].amount += amt;
    dailySales[key].count++;
  }
  
  if (date) {
    const m = date.toISOString().substring(0, 7);
    if (!monthlySales[m]) monthlySales[m] = { amount: 0, count: 0 };
    monthlySales[m].amount += amt;
    monthlySales[m].count++;
  }
});

const topReps = Object.entries(repSales)
  .sort((a, b) => b[1].amount - a[1].amount)
  .slice(0, 15);

const dailySorted = Object.entries(dailySales).sort((a, b) => a[0].localeCompare(b[0]));
const dailyLabels = dailySorted.map(d => d[0]);
const dailyAmounts = dailySorted.map(d => Math.round(d[1].amount));

const monthlySorted = Object.entries(monthlySales).sort((a, b) => a[0].localeCompare(b[0]));
const monthLabels = monthlySorted.map(m => {
  const [y, mo] = m[0].split('-');
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return months[parseInt(mo) - 1] + ' ' + y;
});
const monthAmounts = monthlySorted.map(m => Math.round(m[1].amount));

const repLabels = topReps.map(r => 'Rep #' + r[0]);
const repAmounts = topReps.map(r => Math.round(r[1].amount));
const repCount = Object.keys(repSales).length;

const maxDaily = Math.max(...dailyAmounts);
const avgDaily = Math.round(totalAmount / dailySorted.length);
const peakDay = dailySorted.find(d => d[1].amount === maxDaily)?.[0] || '';

const topRep = topReps[0];
const dates = dailySorted.map(d => d[0]);
const dateRange = dates.length > 0 ? dates[0] + ' to ' + dates[dates.length - 1] : 'N/A';

function fmt(v) {
  return v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

const tableRows = topReps.map((r, i) => {
  const pct = (r[1].amount / totalAmount * 100);
  const barW = Math.max(pct * 2, 2);
  const rankClass = i < 3 ? 'rank-' + (i + 1) : 'rank-other';
  const avgInv = r[1].amount / r[1].count;
  return `<tr>
    <td><span class="rep-rank ${rankClass}">${i + 1}</span></td>
    <td><strong>Rep #${r[0]}</strong></td>
    <td class="text-right">$${fmt(r[1].amount)}</td>
    <td class="text-right">${r[1].count.toLocaleString()}</td>
    <td class="text-right">$${fmt(avgInv)}</td>
    <td class="text-right">${pct.toFixed(1)}%</td>
    <td><span class="bar" style="width:${barW}px"></span></td>
  </tr>`;
}).join('');

const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MMS Sales Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"><\/script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #f0f2f5; color: #333; padding: 20px; }
.container { max-width: 1400px; margin: 0 auto; }
h1 { font-size: 28px; margin-bottom: 6px; color: #1a1a2e; }
.subtitle { color: #666; font-size: 14px; margin-bottom: 24px; }
.metrics { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 24px; }
.metric-card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
.metric-card .label { font-size: 13px; color: #888; text-transform: uppercase; letter-spacing: .5px; margin-bottom: 8px; }
.metric-card .value { font-size: 28px; font-weight: 700; color: #1a1a2e; }
.metric-card .sub { font-size: 12px; color: #999; margin-top: 4px; }
.charts { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
.chart-card { background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,.08); }
.chart-card.full { grid-column: 1 / -1; }
.chart-card h3 { font-size: 15px; color: #444; margin-bottom: 16px; }
.chart-wrap { position: relative; width: 100%; }
.chart-wrap canvas { width: 100% !important; }
.table-wrap { overflow-x: auto; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th { background: #f8f9fa; color: #555; padding: 10px 12px; text-align: left; font-weight: 600; border-bottom: 2px solid #e9ecef; }
td { padding: 10px 12px; border-bottom: 1px solid #f0f0f0; }
tr:hover td { background: #f8f9ff; }
.text-right { text-align: right; }
.bar { display: inline-block; height: 8px; border-radius: 4px; background: #4f6ef7; min-width: 4px; vertical-align: middle; margin-right: 8px; }
.rep-rank { display: inline-flex; align-items: center; justify-content: center; width: 24px; height: 24px; border-radius: 50%; font-size: 11px; font-weight: 700; color: #fff; margin-right: 8px; }
.rank-1 { background: #f5a623; }
.rank-2 { background: #7b8ba8; }
.rank-3 { background: #cd7f32; }
.rank-other { background: #ddd; color: #888; }
@media (max-width: 900px) { .charts { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<div class="container">
<h1>MMS Sales Dashboard</h1>
<p class="subtitle">Data from Kobo Toolbox &bull; ${dateRange} &bull; ${data.length} records</p>

<div class="metrics">
<div class="metric-card">
<div class="label">Total Sales</div>
<div class="value">$${(totalAmount / 1000000).toFixed(2)}M</div>
<div class="sub">$${fmt(totalAmount)}</div>
</div>
<div class="metric-card">
<div class="label">Total Invoices</div>
<div class="value">${totalInvoices.toLocaleString()}</div>
<div class="sub">${invoiceSet.size.toLocaleString()} unique invoice numbers</div>
</div>
<div class="metric-card">
<div class="label">Sales Reps</div>
<div class="value">${repCount}</div>
<div class="sub">Active in this period</div>
</div>
<div class="metric-card">
<div class="label">Avg Daily Sales</div>
<div class="value">$${(avgDaily / 1000).toFixed(1)}K</div>
<div class="sub">Peak: $${(maxDaily / 1000).toFixed(0)}K on ${peakDay}</div>
</div>
<div class="metric-card">
<div class="label">Best Performer</div>
<div class="value">Rep #${topRep[0]}</div>
<div class="sub">$${(topRep[1].amount / 1000000).toFixed(2)}M total</div>
</div>
<div class="metric-card">
<div class="label">Avg Invoice</div>
<div class="value">$${(totalAmount / totalInvoices / 1000).toFixed(1)}K</div>
<div class="sub">${totalInvoices} invoices total</div>
</div>
</div>

<div class="charts">
<div class="chart-card">
<h3>Monthly Sales Trend</h3>
<div class="chart-wrap" style="height:300px"><canvas id="monthlyChart"></canvas></div>
</div>
<div class="chart-card">
<h3>Top 15 Sales Reps</h3>
<div class="chart-wrap" style="height:300px"><canvas id="repChart"></canvas></div>
</div>
<div class="chart-card full">
<h3>Daily Sales</h3>
<div class="chart-wrap" style="height:250px"><canvas id="dailyChart"></canvas></div>
</div>
<div class="chart-card full">
<h3>Sales Rep Performance Details</h3>
<div class="table-wrap">
<table>
<thead><tr><th>Rank</th><th>Sales Rep</th><th>Total Sales</th><th>Invoice Count</th><th>Avg Invoice</th><th>% of Total</th><th></th></tr></thead>
<tbody>
${tableRows}
</tbody>
</table>
</div>
</div>
</div>
</div>

<script>
function fmtChart(v) {
  if (v >= 1000000) return '$' + (v / 1000000).toFixed(1) + 'M';
  if (v >= 1000) return '$' + (v / 1000).toFixed(1) + 'K';
  return '$' + v;
}

new Chart(document.getElementById('monthlyChart'), {
  type: 'bar',
  data: {
    labels: ${JSON.stringify(monthLabels)},
    datasets: [{
      label: 'Sales',
      data: ${JSON.stringify(monthAmounts)},
      backgroundColor: ['rgba(79,110,247,.7)','rgba(79,110,247,.7)','rgba(79,110,247,.7)','rgba(79,110,247,.7)','rgba(79,110,247,.7)'],
      borderColor: '#4f6ef7',
      borderWidth: 1,
      borderRadius: 4
    }]
  },
  options: {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: { y: { beginAtZero: true, ticks: { callback: fmtChart } } }
  }
});

new Chart(document.getElementById('repChart'), {
  type: 'bar',
  data: {
    labels: ${JSON.stringify(repLabels)},
    datasets: [{
      label: 'Sales',
      data: ${JSON.stringify(repAmounts)},
      backgroundColor: '#36a2eb',
      borderRadius: 3
    }]
  },
  options: {
    responsive: true,
    maintainAspectRatio: false,
    indexAxis: 'y',
    plugins: { legend: { display: false } },
    scales: { x: { beginAtZero: true, ticks: { callback: fmtChart } } }
  }
});

new Chart(document.getElementById('dailyChart'), {
  type: 'line',
  data: {
    labels: ${JSON.stringify(dailyLabels)},
    datasets: [{
      label: 'Daily Sales',
      data: ${JSON.stringify(dailyAmounts)},
      borderColor: '#4f6ef7',
      backgroundColor: 'rgba(79,110,247,.08)',
      fill: true,
      tension: .3,
      pointRadius: 0,
      pointHitRadius: 10
    }]
  },
  options: {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: { y: { beginAtZero: true, ticks: { callback: fmtChart } } },
    interaction: { mode: 'index', intersect: false }
  }
});
<\/script>
</body>
</html>`;

fs.writeFileSync('dashboard.html', html);
console.log('dashboard.html generated successfully');
