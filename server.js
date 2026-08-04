const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const http = require('http');
const https = require('https');

const app = express();
const PORT = 3001;
const DATA_URL = 'https://kf.kobotoolbox.org/api/v2/assets/ai5wFipp7kxWNA5RJd5jre/export-settings/esdqM5MrXnavHqJcgWRZwcc/data.xlsx';
const DATA_FILE = path.join(__dirname, 'kobo_data.xlsx');

let cachedData = null;

function excelDateToJS(serial) {
  if (!serial || isNaN(serial)) return null;
  return new Date((serial - 25569) * 86400 * 1000);
}

function fmtNum(v) {
  return Number(v).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function extractReps(row) {
  const multi = row['Sales Rep Number(s)'];
  if (multi && String(multi).trim() !== '') {
    return String(multi).split(/[\/;\\]+/).map(s => s.trim()).filter(Boolean).map(s => s.padStart(2, '0'));
  }
  const primary = row['Sales Rep Number'];
  if (primary !== undefined && primary !== '' && primary !== null) {
    return [String(primary).padStart(2, '0')];
  }
  return [];
}

function processData(raw) {
  const records = raw.map((row, idx) => {
    const date = excelDateToJS(row['Date']);
    const reps = extractReps(row);
    return {
      id: row['_id'],
      uuid: row['_uuid'] || '',
      date: date ? date.toISOString().split('T')[0] : '',
      dateTs: date ? date.getTime() : 0,
      month: date ? date.toISOString().substring(0, 7) : '',
      branch: row['Branch'] || '',
      rep: reps.length > 0 ? reps[0] : '',
      reps: reps,
      repLabel: reps.length > 0 ? 'Rep #' + reps.join(', #') : '-',
      invoice: String(row['Invoice Number'] || ''),
      amount: parseFloat(row['Amount']) || 0,
      index: idx + 1
    };
  });

  const repMap = new Map();
  const dailyMap = new Map();
  const monthlyMap = new Map();
  let totalAmount = 0;
  const invSet = new Set();

  for (const r of records) {
    totalAmount += r.amount;
    if (r.invoice) invSet.add(r.invoice);
    if (r.reps.length > 0) {
      for (const rep of r.reps) {
        const e = repMap.get(rep) || { amount: 0, count: 0 };
        e.amount += r.amount;
        e.count++;
        repMap.set(rep, e);
      }
    }
    if (r.date) {
      const e = dailyMap.get(r.date) || { amount: 0, count: 0 };
      e.amount += r.amount;
      e.count++;
      dailyMap.set(r.date, e);
    }
    if (r.month) {
      const e = monthlyMap.get(r.month) || { amount: 0, count: 0 };
      e.amount += r.amount;
      e.count++;
      monthlyMap.set(r.month, e);
    }
  }
  const repSorted = [...repMap.entries()].sort((a, b) => b[1].count - a[1].count);

  const dailySorted = [...dailyMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const monthlySorted = [...monthlyMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  const summary = {
    totalAmount: Math.round(totalAmount * 100) / 100,
    totalRecords: records.length,
    uniqueInvoices: invSet.size,
    repCount: repMap.size,
    tradingDays: dailySorted.length,
    avgDaily: Math.round(totalAmount / (dailySorted.length || 1)),
    maxDaily: dailySorted.length ? Math.round(Math.max(...dailySorted.map(d => d[1].amount))) : 0,
    peakDay: dailySorted.length ? dailySorted.reduce((a, b) => a[1].amount > b[1].amount ? a : b)[0] : '',
    avgInvoice: Math.round((totalAmount / (records.length || 1)) * 100) / 100,
    dateRange: dailySorted.length ? dailySorted[0][0] + ' to ' + dailySorted[dailySorted.length - 1][0] : ''
  };

  return {
    summary,
    monthly: monthlySorted.map(m => ({
      month: m[0],
      label: months[parseInt(m[0].split('-')[1]) - 1] + ' ' + m[0].split('-')[0],
      amount: Math.round(m[1].amount),
      count: m[1].count
    })),
    daily: dailySorted.map(d => ({ date: d[0], amount: Math.round(d[1].amount), count: d[1].count })),
    topReps: repSorted.map(r => ({
      rep: r[0],
      amount: Math.round(r[1].amount * 100) / 100,
      count: r[1].count,
      avgInvoice: Math.round((r[1].amount / r[1].count) * 100) / 100,
      pct: Math.round((r[1].count / records.length) * 1000) / 10
    })),
    repList: repSorted.map(r => r[0]),
    monthList: monthlySorted.map(m => m[0]),
    branchList: [...new Set(records.map(r => r.branch).filter(Boolean))].sort(),
    records
  };
}

function filterData(query, data) {
  let records = data.records;
  if (query.search) {
    const q = query.search.toLowerCase();
    records = records.filter(r =>
      r.invoice.toLowerCase().includes(q) ||
      r.repLabel.toLowerCase().includes(q) ||
      r.branch.toLowerCase().includes(q) ||
      r.date.includes(q)
    );
  }
  if (query.month) records = records.filter(r => r.month === query.month);
  if (query.rep) records = records.filter(r => r.reps.includes(query.rep));
  if (query.invoice) records = records.filter(r => r.invoice.toLowerCase().includes(query.invoice.toLowerCase()));
  if (query.dateFrom) records = records.filter(r => r.date >= query.dateFrom);
  if (query.dateTo) records = records.filter(r => r.date <= query.dateTo);

  const result = computeFiltered(records);
  if (query.top) result.topReps = result.topReps.slice(0, parseInt(query.top));
  return result;
}

function computeFiltered(records) {
  const repMap = new Map();
  const dailyMap = new Map();
  const monthlyMap = new Map();
  let totalAmount = 0;
  const invSet = new Set();
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

  for (const r of records) {
    totalAmount += r.amount;
    if (r.invoice) invSet.add(r.invoice);
    if (r.reps.length > 0) {
      for (const rep of r.reps) {
        const e = repMap.get(rep) || { amount: 0, count: 0 };
        e.amount += r.amount; e.count++;
        repMap.set(rep, e);
      }
    }
    if (r.date) {
      const e = dailyMap.get(r.date) || { amount: 0, count: 0 };
      e.amount += r.amount; e.count++;
      dailyMap.set(r.date, e);
    }
    if (r.month) {
      const e = monthlyMap.get(r.month) || { amount: 0, count: 0 };
      e.amount += r.amount; e.count++;
      monthlyMap.set(r.month, e);
    }
  }

  const repSorted = [...repMap.entries()].sort((a, b) => b[1].count - a[1].count);
  const dailySorted = [...dailyMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
  const monthlySorted = [...monthlyMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));

  return {
    summary: {
      totalAmount: Math.round(totalAmount * 100) / 100,
      totalRecords: records.length,
      uniqueInvoices: invSet.size,
      repCount: repMap.size,
      tradingDays: dailySorted.length,
      avgDaily: Math.round(totalAmount / (dailySorted.length || 1)),
      maxDaily: dailySorted.length ? Math.round(Math.max(...dailySorted.map(d => d[1].amount))) : 0,
      peakDay: dailySorted.length ? dailySorted.reduce((a, b) => a[1].amount > b[1].amount ? a : b)[0] : '',
      avgInvoice: Math.round((totalAmount / (records.length || 1)) * 100) / 100,
      dateRange: dailySorted.length ? dailySorted[0][0] + ' to ' + dailySorted[dailySorted.length - 1][0] : ''
    },
    monthly: monthlySorted.map(m => ({
      month: m[0],
      label: months[parseInt(m[0].split('-')[1]) - 1] + ' ' + m[0].split('-')[0],
      amount: Math.round(m[1].amount),
      count: m[1].count
    })),
    daily: dailySorted.map(d => ({ date: d[0], amount: Math.round(d[1].amount), count: d[1].count })),
    topReps: repSorted.map(r => ({
      rep: r[0],
      amount: Math.round(r[1].amount * 100) / 100,
      count: r[1].count,
      avgInvoice: Math.round((r[1].amount / r[1].count) * 100) / 100,
      pct: Math.round((r[1].count / records.length) * 1000) / 10
    })),
    records: records.slice(0, 500),
    totalFiltered: records.length
  };
}

function loadFromFile() {
  if (!fs.existsSync(DATA_FILE)) return null;
  const workbook = XLSX.readFile(DATA_FILE);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  console.log('Loaded ' + raw.length + ' records from cache');
  return processData(raw);
}

function resolveUrl(base, relative) {
  if (relative.startsWith('http://') || relative.startsWith('https://')) return relative;
  const u = new URL(base);
  return u.origin + relative;
}

function downloadFile(url, depth) {
  depth = depth || 0;
  if (depth > 5) return Promise.reject(new Error('Too many redirects'));
  return new Promise((resolve, reject) => {
    const mod = url.startsWith('https') ? https : http;
    mod.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' }, timeout: 60000 }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        const next = resolveUrl(url, res.headers.location);
        res.resume();
        return downloadFile(next, depth + 1).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error('HTTP ' + res.statusCode + ' from ' + url));
      }
      const chunks = [];
      res.on('data', c => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

async function fetchFromKobo() {
  console.log('Fetching data from Kobo...');
  const buf = await downloadFile(DATA_URL);
  fs.writeFileSync(DATA_FILE, buf);
  console.log('Downloaded ' + buf.length + ' bytes');
  const workbook = XLSX.read(buf, { type: 'buffer' });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  return processData(raw);
}

function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function generateHTML(data) {
  const s = data.summary;
  const topRep = data.topReps[0] || { rep: '-', amount: 0, count: 0, pct: 0 };

  const repRows = data.topReps.map((r, i) => {
    const rc = i < 3 ? 'rank-' + (i + 1) : 'rank-other';
    const bw = Math.max(r.pct * 2, 2);
    return '<tr><td><span class="rep-rank ' + rc + '">' + (i + 1) + '</span></td><td><strong>Rep #' + esc(r.rep) + '</strong></td><td class="tr">$' + fmtNum(r.amount) + '</td><td class="tr">' + r.count.toLocaleString() + '</td><td class="tr">$' + fmtNum(r.avgInvoice) + '</td><td class="tr">' + r.pct + '%</td><td><span class="bar" style="width:' + bw + 'px"></span></td></tr>';
  }).join('');

  const recRows = data.records.slice(0, 500).map(r =>
    '<tr><td>' + r.index + '</td><td>' + esc(r.date) + '</td><td>' + esc(r.invoice) + '</td><td>' + esc(r.repLabel) + '</td><td>' + esc(r.branch) + '</td><td class="tr">$' + fmtNum(r.amount) + '</td><td style="font-size:10px;color:#999">' + esc(r.uuid).substring(0, 10) + '...</td></tr>'
  ).join('');

  const monthOpts = data.monthList.map(m => '<option value="' + esc(m) + '">' + esc(m) + '</option>').join('');
  const repOpts = data.repList.map(r => '<option value="' + esc(r) + '">Rep #' + esc(r) + '</option>').join('');

  const repMonthOpts = '<option value="">All Time</option>' + monthOpts;

  return '<!DOCTYPE html>' +
'<html lang="en"><head>' +
'<meta charset="UTF-8">' +
'<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
'<title>MMS Sales Dashboard</title>' +
'<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"><' + '/script>' +
'<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"><' + '/script>' +
'<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.3/jspdf.plugin.autotable.min.js"><' + '/script>' +
'<style>' +
'*{margin:0;padding:0;box-sizing:border-box}' +
'body{font-family:"Segoe UI",Tahoma,sans-serif;background:#f0f2f5;color:#333;padding:20px}' +
'.container{max-width:1500px;margin:0 auto}' +
'h1{font-size:28px;margin-bottom:4px;color:#1a1a2e}' +
'.sub{color:#888;font-size:13px;margin-bottom:20px}' +
'.pull-section{background:#fff;border-radius:12px;padding:14px 20px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:16px;display:flex;align-items:center;gap:14px;flex-wrap:wrap}' +
'.filters{background:#fff;border-radius:12px;padding:14px 20px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:16px;display:flex;flex-wrap:wrap;gap:10px;align-items:end}' +
'.fg{display:flex;flex-direction:column;gap:3px}' +
'.fg label{font-size:10px;color:#888;text-transform:uppercase;letter-spacing:.3px;font-weight:600}' +
'.fg input,.fg select{padding:6px 10px;border:1px solid #ddd;border-radius:6px;font-size:13px;min-width:130px;outline:none;background:#fff}' +
'.fg input:focus,.fg select:focus{border-color:#4f6ef7;box-shadow:0 0 0 2px rgba(79,110,247,.15)}' +
'.fg input[type="date"]{min-width:140px}' +
'.btn{padding:7px 16px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600;white-space:nowrap}' +
'.btn-p{background:#4f6ef7;color:#fff}' +
'.btn-p:hover{background:#3d5bd9}' +
'.btn-o{background:transparent;color:#555;border:1px solid #ddd}' +
'.btn-o:hover{background:#f5f5f5}' +
'.btn-s{background:#22c55e;color:#fff}' +
'.btn-s:hover{background:#16a34a}' +
'.btn-d{background:#8b5cf6;color:#fff}' +
'.btn-d:hover{background:#7c3aed}' +
'.btn-x{background:#f97316;color:#fff}' +
'.btn-x:hover{background:#ea580c}' +
'.ps{font-size:13px;color:#666}' +
'.metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:16px}' +
'.mc{background:#fff;border-radius:12px;padding:16px 18px;box-shadow:0 2px 8px rgba(0,0,0,.06)}' +
'.mc .l{font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.4px;margin-bottom:5px}' +
'.mc .v{font-size:24px;font-weight:700;color:#1a1a2e}' +
'.mc .s{font-size:11px;color:#aaa;margin-top:2px}' +
'.charts{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}' +
'.cc{background:#fff;border-radius:12px;padding:16px;box-shadow:0 2px 8px rgba(0,0,0,.06)}' +
'.cc.full{grid-column:1/-1}' +
'.cc h3{font-size:14px;color:#444;margin-bottom:10px}' +
'.cc-hdr{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:10px}' +
'.cc-hdr h3{margin-bottom:0}' +
'.cc-hdr .fg{flex-direction:row;align-items:center;gap:6px}' +
'.cc-hdr .fg label{margin-bottom:0}' +
'.tw{overflow-x:auto;max-height:500px;overflow-y:auto}' +
'table{width:100%;border-collapse:collapse;font-size:12px}' +
'th{background:#f8f9fa;color:#555;padding:7px 10px;text-align:left;font-weight:600;border-bottom:2px solid #e9ecef;position:sticky;top:0;z-index:1}' +
'td{padding:7px 10px;border-bottom:1px solid #f0f0f0}' +
'tr:hover td{background:#f8f9ff}' +
'.tr{text-align:right}' +
'.bar{display:inline-block;height:7px;border-radius:3px;background:#4f6ef7;min-width:3px;vertical-align:middle;margin-right:6px}' +
'.rr{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;font-size:10px;font-weight:700;color:#fff;margin-right:6px}' +
'.r1{background:#f5a623}' +
'.r2{background:#7b8ba8}' +
'.r3{background:#cd7f32}' +
'.ro{background:#e5e7eb;color:#888}' +
'.rc{font-size:12px;color:#999;padding:4px 0}' +
'@media(max-width:900px){.charts{grid-template-columns:1fr}}' +
'.sb{position:relative}' +
'.sb input{padding-left:28px!important}' +
'.btn-group{display:flex;gap:6px;flex-wrap:wrap}' +
'@media print{body{background:#fff;padding:10px}.pull-section,.filters,.btn{display:none!important}.cc{box-shadow:none;border:1px solid #ddd;break-inside:avoid}}' +
'</style></head><body>' +
'<div class="container">' +
'<h1>MMS Sales Dashboard</h1>' +
'<p class="sub" id="sub">Data from Kobo Toolbox &bull; ' + s.dateRange + ' &bull; ' + s.totalRecords + ' records</p>' +

'<div class="pull-section">' +
  '<button class="btn btn-s" onclick="pullData()">&#x27F3; Pull Latest Data</button>' +
  '<span class="ps" id="ps">Last refreshed from Kobo API</span>' +
'</div>' +

'<div class="filters">' +
  '<div class="fg sb"><label>Search</label><input type="text" id="q" placeholder="Invoice, Rep, Branch..." oninput="df()"></div>' +
  '<div class="fg"><label>From</label><input type="date" id="df" onchange="df()"></div>' +
  '<div class="fg"><label>To</label><input type="date" id="dt" onchange="df()"></div>' +
  '<div class="fg"><label>Month</label><select id="mf" onchange="af()"><option value="">All Months</option>' + monthOpts + '</select></div>' +
  '<div class="fg"><label>Sales Rep</label><select id="rf" onchange="af()"><option value="">All Reps</option>' + repOpts + '</select></div>' +
  '<div class="fg"><label>Invoice</label><input type="text" id="invf" placeholder="Invoice #" oninput="df()"></div>' +
  '<button class="btn btn-o" onclick="rst()">Reset</button>' +
  '<div class="btn-group">' +
    '<button class="btn btn-d" onclick="exportPDF()" title="Download filtered data as PDF">&#x1F4C4; PDF</button>' +
    '<button class="btn btn-x" onclick="exportXLSX()" title="Download filtered data as Excel">&#x1F4C3; Excel</button>' +
  '</div>' +
'</div>' +

'<div class="metrics">' +
  '<div class="mc"><div class="l">Total Sales</div><div class="v" id="mT">$' + (s.totalAmount / 1000000).toFixed(2) + 'M</div><div class="s">$' + fmtNum(s.totalAmount) + '</div></div>' +
  '<div class="mc"><div class="l">Invoices</div><div class="v" id="mI">' + s.totalRecords.toLocaleString() + '</div><div class="s">' + s.uniqueInvoices.toLocaleString() + ' unique</div></div>' +
  '<div class="mc"><div class="l">Sales Reps</div><div class="v" id="mR">' + s.repCount + '</div><div class="s">Active this period</div></div>' +
  '<div class="mc"><div class="l">Avg Daily</div><div class="v" id="mAD">$' + (s.avgDaily / 1000).toFixed(1) + 'K</div><div class="s">Peak: $' + (s.maxDaily / 1000).toFixed(0) + 'K</div></div>' +
  '<div class="mc"><div class="l">Best Rep</div><div class="v" id="mBR">Rep #' + topRep.rep + '</div><div class="s">$' + (topRep.amount / 1000000).toFixed(2) + 'M</div></div>' +
  '<div class="mc"><div class="l">Avg Invoice</div><div class="v" id="mAI">$' + (s.avgInvoice / 1000).toFixed(1) + 'K</div><div class="s">Across ' + s.totalRecords + ' records</div></div>' +
'</div>' +

'<div class="charts">' +
  '<div class="cc"><h3>Monthly Sales Trend</h3><div style="height:280px"><canvas id="mc"></canvas></div></div>' +
  '<div class="cc"><div class="cc-hdr"><h3>All Sales Reps</h3><div class="fg"><label>Month</label><select id="rmf" onchange="repFilter()">' + repMonthOpts + '</select></div></div><div style="height:280px"><canvas id="rc"></canvas></div></div>' +
  '<div class="cc full"><h3>Daily Sales</h3><div style="height:200px"><canvas id="dc"></canvas></div></div>' +
  '<div class="cc full"><div class="cc-hdr"><h3>All Rep Performance</h3><div class="fg"><label>Month</label><select id="rmt" onchange="repFilter()">' + repMonthOpts + '</select></div></div><div class="tw" style="max-height:500px"><table id="repTable"><thead><tr><th>Rank</th><th>Sales Rep</th><th>Total Sales</th><th>Invoices</th><th>Avg Invoice</th><th>% of Invoices</th><th></th></tr></thead><tbody>' + repRows + '</tbody></table></div></div>' +
  '<div class="cc full"><h3>Transactions</h3><p class="rc" id="rcnt">Showing ' + Math.min(data.records.length, 500) + ' of ' + data.records.length + ' records</p><div class="tw" style="max-height:400px"><table><thead><tr><th>#</th><th>Date</th><th>Invoice</th><th>Rep</th><th>Branch</th><th class="tr">Amount</th><th>UUID</th></tr></thead><tbody id="rb">' + recRows + '</tbody></table></div></div>' +
'</div></div>' +

'<script>' +
'var D=' + JSON.stringify(data) + ';' +
'function fn(v){return v.toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2})}' +
'function fc(v){if(v>=1e6)return"$"+(v/1e6).toFixed(1)+"M";if(v>=1e3)return"$"+(v/1e3).toFixed(1)+"K";return"$"+v}' +
'var _mc,_rc,_dc;var _ft;' +
'var _repMonth="";' +
'function df(){clearTimeout(_ft);_ft=setTimeout(af,300)}' +
'async function af(){var p=new URLSearchParams();' +
  'var q=document.getElementById("q").value.trim();if(q)p.set("search",q);' +
  'var df=document.getElementById("df").value;if(df)p.set("dateFrom",df);' +
  'var dt=document.getElementById("dt").value;if(dt)p.set("dateTo",dt);' +
  'var mf=document.getElementById("mf").value;if(mf)p.set("month",mf);' +
  'var rf=document.getElementById("rf").value;if(rf)p.set("rep",rf);' +
  'var invf=document.getElementById("invf").value.trim();if(invf)p.set("invoice",invf);' +
  'if(_repMonth)p.set("rmonth",_repMonth);' +
  'try{var r=await fetch("/api/data?"+p.toString());var d=await r.json();ud(d)}catch(e){console.error(e)}' +
'}' +
'function ud(d){' +
  'var s=d.summary;var tr=d.topReps[0]||{rep:"-",amount:0,count:0};' +
  'document.getElementById("sub").textContent="Data from Kobo Toolbox &bull; "+s.dateRange+" &bull; "+s.totalRecords+" records";' +
  'document.getElementById("mT").textContent="$"+(s.totalAmount/1e6).toFixed(2)+"M";' +
  'document.getElementById("mT").nextElementSibling.textContent="$"+fn(s.totalAmount);' +
  'document.getElementById("mI").textContent=s.totalRecords.toLocaleString();' +
  'document.getElementById("mI").nextElementSibling.textContent=s.uniqueInvoices.toLocaleString()+" unique";' +
  'document.getElementById("mR").textContent=s.repCount;' +
  'document.getElementById("mAD").textContent="$"+(s.avgDaily/1e3).toFixed(1)+"K";' +
  'document.getElementById("mAD").nextElementSibling.textContent="Peak: $"+(s.maxDaily/1e3).toFixed(0)+"K";' +
  'document.getElementById("mAI").textContent="$"+(s.avgInvoice/1e3).toFixed(1)+"K";' +
  'uc(d);urt(d);urt2(d);' +
'}' +
'function uc(d){' +
  'if(_mc)_mc.destroy();if(_rc)_rc.destroy();if(_dc)_dc.destroy();' +
  '_mc=new Chart(document.getElementById("mc"),{type:"bar",data:{labels:d.monthly.map(function(m){return m.label}),datasets:[{label:"Sales",data:d.monthly.map(function(m){return m.amount}),backgroundColor:"rgba(79,110,247,.7)",borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{callback:fc}}}}});' +
  'var repData=_repMonth?d.topReps:D.topReps;' +
  '_rc=new Chart(document.getElementById("rc"),{type:"bar",data:{labels:repData.map(function(r){return"Rep #"+r.rep}),datasets:[{label:"Sales",data:repData.map(function(r){return r.amount}),backgroundColor:"#36a2eb",borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{callback:fc}}}}});' +
  '_dc=new Chart(document.getElementById("dc"),{type:"line",data:{labels:d.daily.map(function(x){return x.date}),datasets:[{label:"Daily Sales",data:d.daily.map(function(x){return x.amount}),borderColor:"#4f6ef7",backgroundColor:"rgba(79,110,247,.08)",fill:true,tension:.3,pointRadius:0,pointHitRadius:10}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,ticks:{callback:fc}}},interaction:{mode:"index",intersect:false}}});' +
'}' +
'async function repFilter(){' +
  'var rmf=document.getElementById("rmf").value;' +
  'document.getElementById("rmt").value=rmf;' +
  '_repMonth=rmf;' +
  'var p=new URLSearchParams();' +
  'if(rmf)p.set("month",rmf);' +
  'try{var r=await fetch("/api/data?"+p.toString());var d=await r.json();' +
  'if(_rc)_rc.destroy();' +
  'var rd=d.topReps;' +
  '_rc=new Chart(document.getElementById("rc"),{type:"bar",data:{labels:rd.map(function(r){return"Rep #"+r.rep}),datasets:[{label:"Sales",data:rd.map(function(r){return r.amount}),backgroundColor:"#36a2eb",borderRadius:3}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,ticks:{callback:fc}}}}});' +
  'var tb=document.querySelector("#repTable tbody");' +
  'tb.innerHTML=rd.map(function(r,i){var rc=i<3?"rank-"+(i+1):"rank-other";var bw=Math.max(r.pct*2,2);return"<tr><td><span class=\\"rep-rank "+rc+"\\">"+(i+1)+"</span></td><td><strong>Rep #"+r.rep+"</strong></td><td class=\\"tr\\">$"+fn(r.amount)+"</td><td class=\\"tr\\">"+r.count.toLocaleString()+"</td><td class=\\"tr\\">$"+fn(r.avgInvoice)+"</td><td class=\\"tr\\">"+r.pct+' +
  '"%</td><td><span class=\\"bar\\" style=\\"width:"+bw+"px\\"></span></td></tr>"}).join("");' +
  'var best=rd[0]||{rep:"-",amount:0};' +
  'document.getElementById("mBR").textContent="Rep #"+best.rep;' +
  'document.getElementById("mBR").nextElementSibling.textContent="$"+(best.amount/1e6).toFixed(2)+"M";' +
  '}catch(e){console.error(e)}' +
'}' +
'function urt(d){' +
  'var tb=document.querySelector("#repTable tbody");if(!tb)return;' +
  'tb.innerHTML=d.topReps.map(function(r,i){var rc=i<3?"rank-"+(i+1):"rank-other";var bw=Math.max(r.pct*2,2);return"<tr><td><span class=\\"rep-rank "+rc+"\\">"+(i+1)+"</span></td><td><strong>Rep #"+r.rep+"</strong></td><td class=\\"tr\\">$"+fn(r.amount)+"</td><td class=\\"tr\\">"+r.count.toLocaleString()+"</td><td class=\\"tr\\">$"+fn(r.avgInvoice)+"</td><td class=\\"tr\\">"+r.pct+' +
'"%</td><td><span class=\\"bar\\" style=\\"width:"+bw+"px\\"></span></td></tr>"}).join("");' +
'}' +
'function urt2(d){' +
  'var b=document.getElementById("rb");var i=document.getElementById("rcnt");var rec=d.records||[];' +
  'i.textContent="Showing "+rec.length+" of "+(d.totalFiltered||rec.length)+" records";' +
  'b.innerHTML=rec.map(function(r){return"<tr><td>"+r.index+"</td><td>"+r.date+"</td><td>"+r.invoice+"</td><td>"+(r.repLabel||"-")+"</td><td>"+r.branch+"</td><td class=\\"tr\\">$"+fn(r.amount)+"</td><td style=\\"font-size:10px;color:#999\\">"+(r.uuid||"").substring(0,10)+"...</td></tr>"}).join("");' +
'}' +
'function rst(){' +
  'document.getElementById("q").value="";document.getElementById("df").value="";' +
  'document.getElementById("dt").value="";document.getElementById("mf").value="";' +
  'document.getElementById("rf").value="";document.getElementById("invf").value="";' +
  '_repMonth="";document.getElementById("rmf").value="";document.getElementById("rmt").value="";' +
  'af();repFilter();' +
'}' +
'async function pullData(){' +
  'var btn=event.target;var st=document.getElementById("ps");' +
  'btn.disabled=true;btn.textContent="&#x27F3; Pulling...";' +
  'st.textContent="Downloading latest data from Kobo...";st.style.color="#4f6ef7";' +
  'try{var r=await fetch("/api/refresh");var d=await r.json();' +
  'if(d.success){st.textContent="Refreshed - "+d.records+" records";st.style.color="#22c55e";setTimeout(function(){location.reload()},1500)}' +
  'else{throw new Error(d.error)}}catch(e){st.textContent="Error: "+e.message;st.style.color="#ef4444"}' +
  'finally{btn.disabled=false;btn.textContent="&#x27F3; Pull Latest Data"}}' +
'function getFilterParams(){' +
  'var p=new URLSearchParams();' +
  'var q=document.getElementById("q").value.trim();if(q)p.set("search",q);' +
  'var df=document.getElementById("df").value;if(df)p.set("dateFrom",df);' +
  'var dt=document.getElementById("dt").value;if(dt)p.set("dateTo",dt);' +
  'var mf=document.getElementById("mf").value;if(mf)p.set("month",mf);' +
  'var rf=document.getElementById("rf").value;if(rf)p.set("rep",rf);' +
  'var invf=document.getElementById("invf").value.trim();if(invf)p.set("invoice",invf);' +
  'if(_repMonth)p.set("rmonth",_repMonth);' +
  'return p.toString();' +
'}' +
'function exportPDF(){' +
  'var p=getFilterParams();' +
  'fetch("/api/data?"+p).then(function(r){return r.json()}).then(function(d){' +
    'var s=d.summary;var repRows=d.topReps;' +
    'var doc=new jspdf.jsPDF({orientation:"portrait",unit:"mm",format:"a4"});' +
    'doc.setFontSize(16);doc.text("MMS Sales Report",14,15);' +
    'doc.setFontSize(9);doc.text("Date: "+s.dateRange+" | Records: "+s.totalRecords+" | Total: $"+fn(s.totalAmount),14,22);' +
    'doc.setFontSize(10);doc.text("Top 20 Sales Reps",14,30);' +
    'var cols=[{header:"Rank",dataKey:"rank"},{header:"Sales Rep",dataKey:"rep"},{header:"Total Sales",dataKey:"amount"},{header:"Invoices",dataKey:"count"},{header:"Avg Invoice",dataKey:"avg"},{header:"% of Invoices",dataKey:"pct"}];' +
    'var rows=repRows.map(function(r,i){return{rank:i+1,rep:"Rep #"+r.rep,amount:"$"+fn(r.amount),count:r.count,avg:"$"+fn(r.avgInvoice),pct:r.pct+"%"}});' +
    'doc.autoTable({startY:33,columns:cols,body:rows,headStyles:{fillColor:[79,110,247],fontSize:8},bodyStyles:{fontSize:7},columnStyles:{amount:{halign:"right"},count:{halign:"right"},avg:{halign:"right"},pct:{halign:"right"}}});' +
    'doc.save("MMS_Sales_Report_"+new Date().toISOString().substring(0,10)+".pdf");' +
  '}).catch(function(e){alert("PDF export error: "+e.message)});' +
'}' +
'function exportXLSX(){' +
  'var p=getFilterParams();' +
  'window.location.href="/api/export/xlsx?"+p;' +
'}' +
'uc(D);' +
'<' + '/script>' +
'</body></html>';
}

// --- Routes ---

app.use(express.json());

app.get('/api/data', (req, res) => {
  if (!cachedData) return res.status(503).json({ error: 'Data not loaded' });
  res.json(filterData(req.query, cachedData));
});

app.get('/api/export/xlsx', (req, res) => {
  if (!cachedData) return res.status(503).json({ error: 'Data not loaded' });
  try {
    const filtered = filterData(req.query, cachedData);
    const wb = XLSX.utils.book_new();

    // Sheet 1: Summary
    const s = filtered.summary;
    const summaryData = [
      ['Metric', 'Value'],
      ['Date Range', s.dateRange],
      ['Total Records', s.totalRecords],
      ['Unique Invoices', s.uniqueInvoices],
      ['Active Sales Reps', s.repCount],
      ['Total Sales', s.totalAmount],
      ['Avg Daily Sales', s.avgDaily],
      ['Max Daily Sales', s.maxDaily],
      ['Avg Invoice', s.avgInvoice]
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, ws1, 'Summary');

    // Sheet 2: Top Reps
    const repData = filtered.topReps.map((r, i) => ({
      Rank: i + 1,
      'Sales Rep': 'Rep #' + r.rep,
      'Total Sales': r.amount,
      'Invoice Count': r.count,
      'Avg Invoice': r.avgInvoice,
      '% of Total': r.pct + '%'
    }));
    const ws2 = XLSX.utils.json_to_sheet(repData);
    XLSX.utils.book_append_sheet(wb, ws2, 'Top Reps');

    // Sheet 3: Transactions
    const txData = filtered.records.map(r => ({
      '#': r.index,
      Date: r.date,
      Invoice: r.invoice,
      Rep: r.rep ? 'Rep #' + r.rep : '',
      Branch: r.branch,
      Amount: r.amount
    }));
    const ws3 = XLSX.utils.json_to_sheet(txData);
    XLSX.utils.book_append_sheet(wb, ws3, 'Transactions');

    // Sheet 4: Monthly
    const monData = filtered.monthly.map(m => ({
      Month: m.label,
      Sales: m.amount,
      Records: m.count
    }));
    const ws4 = XLSX.utils.json_to_sheet(monData);
    XLSX.utils.book_append_sheet(wb, ws4, 'Monthly');

    // Sheet 5: Daily
    const dayData = filtered.daily.map(d => ({
      Date: d.date,
      Sales: d.amount,
      Records: d.count
    }));
    const ws5 = XLSX.utils.json_to_sheet(dayData);
    XLSX.utils.book_append_sheet(wb, ws5, 'Daily');

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="MMS_Sales_Export_' + new Date().toISOString().substring(0, 10) + '.xlsx"');
    res.send(buf);
  } catch (e) {
    console.error('Export error:', e);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/refresh', async (req, res) => {
  try {
    cachedData = await fetchFromKobo();
    res.json({ success: true, records: cachedData.summary.totalRecords });
  } catch (e) {
    console.error('Refresh failed:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/', (req, res) => {
  if (!cachedData) return res.send('<h1>Loading data...</h1><script>setTimeout(()=>location.reload(),2000)</script>');
  res.send(generateHTML(cachedData));
});

async function start() {
  cachedData = loadFromFile();
  if (!cachedData) {
    console.log('No cache file. Fetching from Kobo...');
    try { cachedData = await fetchFromKobo(); } catch (e) { console.error('Initial fetch failed:', e.message); }
  }
  if (cachedData) {
    console.log('Summary: ' + cachedData.summary.totalRecords + ' records, $' + (cachedData.summary.totalAmount / 1e6).toFixed(2) + 'M total');
  }
  app.listen(PORT, () => console.log('Dashboard: http://localhost:' + PORT));
}

start();
