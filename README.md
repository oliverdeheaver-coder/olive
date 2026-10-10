/**
 * Oliver's Health Dashboard — ONE FILE, all-in-one
 * ------------------------------------------------
 * 1. Reads the CSV files Health Sync puts in Google Drive
 * 2. Archives every day into this spreadsheet permanently (never deletes)
 * 3. Serves a dashboard web page straight from this script
 *
 * HOW TO USE
 *   a) Create a new Google Sheet → Extensions → Apps Script
 *   b) Delete what's there, paste this whole file, click Save
 *   c) Pick "setup" in the function dropdown → Run → approve permissions
 *   d) Deploy → New deployment → type: Web app
 *        Execute as: Me    Who has access: Only myself
 *      → Deploy → open the URL it gives you (bookmark it on your phone)
 *
 * Usually nothing needs editing. If a metric shows no data, run "inspectFiles",
 * read the log, and fill in the overrides in CONFIG below.
 */

const CONFIG = {
  // Drive folder Health Sync writes to. If not found, the whole Drive is searched.
  folderName: 'Health Sync',

  // fileMatch: words that appear in the CSV file names (any of them, case-insensitive)
  // dateCol / valueCol: leave null for auto-detect, or set a column number (0 = first)
  // agg: combine several rows on the same day → 'sum' or 'avg'
  metrics: [
    { sheet: 'Steps',     label: 'Steps',      unit: 'steps', chart: 'bar',
      fileMatch: ['step'],            valueHints: ['step', 'count', 'value'],
      dateCol: null, valueCol: null, agg: 'sum' },
    { sheet: 'Sleep',     label: 'Sleep',      unit: 'min',   chart: 'bar',
      fileMatch: ['sleep'],           valueHints: ['duration', 'minute', 'sleep', 'total', 'value'],
      dateCol: null, valueCol: null, agg: 'sum' },
    { sheet: 'HeartRate', label: 'Heart rate', unit: 'bpm',   chart: 'line',
      fileMatch: ['heart', 'pulse'],  valueHints: ['bpm', 'heart', 'rate', 'pulse', 'value'],
      dateCol: null, valueCol: null, agg: 'avg' },
  ],
};

// =====================================================================
//  RUN THESE
// =====================================================================

/** Run once. Imports now and then every hour from then on. */
function setup() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'syncAll'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  syncAll();
}

/** Pulls the latest CSVs into the archive tabs. Runs hourly by itself. */
function syncAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  CONFIG.metrics.forEach(function (m) {
    const files = findCsvFiles_(m.fileMatch);
    if (!files.length) { console.warn(m.label + ': no CSV files found'); return; }
    files.sort(function (a, b) { return a.getLastUpdated() - b.getLastUpdated(); }); // newest wins

    const incoming = new Map();
    files.forEach(function (f) {
      perDay_(f, m).forEach(function (v, d) { incoming.set(d, v); });
    });
    const added = mergeIntoSheet_(ss, m.sheet, incoming);
    console.log(m.label + ': ' + incoming.size + ' days read, ' + added + ' new days archived');
  });
}

/** Troubleshooting: logs every CSV's name and first rows with column numbers. */
function inspectFiles() {
  findCsvFiles_(['']).forEach(function (f) {
    const rows = Utilities.parseCsv(f.getBlob().getDataAsString()).slice(0, 3);
    console.log(f.getName() + '\n' + rows.map(function (r) {
      return r.map(function (c, j) { return '[' + j + '] ' + c; }).join(' | ');
    }).join('\n'));
  });
}

// =====================================================================
//  WEB PAGE
// =====================================================================

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const payload = CONFIG.metrics.map(function (m) {
    const sh = ss.getSheetByName(m.sheet);
    const rows = sh && sh.getLastRow() > 1
      ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
          .map(function (r) { return [String(r[0]), Number(r[1])]; })
      : [];
    return { label: m.label, unit: m.unit, chart: m.chart, rows: rows };
  });
  const json = JSON.stringify(payload).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput(PAGE_.replace('__DATA__', json))
    .setTitle("Oliver's Health")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const PAGE_ = String.raw`
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  :root { --bg:#f6f7f9; --card:#fff; --ink:#1c1f24; --muted:#6b7280; --line:#e5e7eb; --accent:#2563eb; }
  @media (prefers-color-scheme: dark) {
    :root { --bg:#111317; --card:#1a1d23; --ink:#eef0f3; --muted:#9aa1ad; --line:#2a2e36; --accent:#60a5fa; }
  }
  * { box-sizing: border-box; }
  body { margin:0; background:var(--bg); color:var(--ink); font:15px/1.4 system-ui,-apple-system,Segoe UI,sans-serif; }
  main { max-width:960px; margin:0 auto; padding:24px 16px 48px; }
  header { display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between; margin-bottom:20px; }
  h1 { font-size:22px; margin:0; }
  .ranges button { border:1px solid var(--line); background:var(--card); color:var(--ink); padding:6px 12px; border-radius:8px; cursor:pointer; }
  .ranges button.on { background:var(--accent); border-color:var(--accent); color:#fff; }
  section { background:var(--card); border:1px solid var(--line); border-radius:12px; padding:16px; margin-bottom:16px; }
  section h2 { font-size:16px; margin:0 0 8px; }
  .stats { display:flex; gap:20px; flex-wrap:wrap; color:var(--muted); font-size:13px; margin-bottom:8px; }
  .stats b { display:block; color:var(--ink); font-size:20px; }
  .chart { position:relative; height:240px; }
</style>
<main>
  <header>
    <h1>Oliver's Health</h1>
    <div class="ranges">
      <button data-days="7">7d</button><button data-days="30" class="on">30d</button>
      <button data-days="90">90d</button><button data-days="0">All</button>
    </div>
  </header>
  <div id="metrics"></div>
</main>
<script>
  var DATA = __DATA__;
  var rangeDays = 30, charts = {};
  var accent = getComputedStyle(document.documentElement).getPropertyValue('--accent').trim();

  function fmt(n) { return Math.round(n).toLocaleString('en-GB'); }
  function inRange(rows) {
    if (!rangeDays) return rows;
    var cutoff = new Date(Date.now() - rangeDays * 864e5).toISOString().slice(0, 10);
    return rows.filter(function (r) { return r[0] >= cutoff; });
  }
  function render(m, i) {
    var rows = inRange(m.rows), vals = rows.map(function (r) { return r[1]; });
    var avg = vals.length ? vals.reduce(function (a, b) { return a + b; }, 0) / vals.length : 0;
    var last = rows[rows.length - 1];
    document.getElementById('stats' + i).innerHTML = last
      ? '<span><b>' + fmt(last[1]) + '</b>latest (' + last[0] + ')</span>' +
        '<span><b>' + fmt(avg) + '</b>average ' + m.unit + '</span>' +
        '<span><b>' + fmt(Math.max.apply(null, vals)) + '</b>best</span>' +
        '<span><b>' + rows.length + '</b>days</span>'
      : 'No data in this range yet.';
    var labels = rows.map(function (r) {
      return new Date(r[0]).toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });
    });
    if (charts[i]) {
      charts[i].data.labels = labels; charts[i].data.datasets[0].data = vals; charts[i].update(); return;
    }
    charts[i] = new Chart(document.getElementById('chart' + i), {
      type: m.chart,
      data: { labels: labels, datasets: [{ data: vals, backgroundColor: accent, borderColor: accent, tension: 0.3, pointRadius: 2 }] },
      options: { maintainAspectRatio: false, plugins: { legend: { display: false } },
                 scales: { y: { beginAtZero: m.chart === 'bar' } } }
    });
  }
  var host = document.getElementById('metrics');
  DATA.forEach(function (m, i) {
    host.insertAdjacentHTML('beforeend',
      '<section><h2>' + m.label + '</h2><div class="stats" id="stats' + i + '"></div>' +
      '<div class="chart"><canvas id="chart' + i + '"></canvas></div></section>');
    render(m, i);
  });
  document.querySelectorAll('.ranges button').forEach(function (b) {
    b.addEventListener('click', function () {
      document.querySelectorAll('.ranges button').forEach(function (x) { x.classList.remove('on'); });
      b.classList.add('on'); rangeDays = Number(b.dataset.days);
      DATA.forEach(render);
    });
  });
</script>`;

// =====================================================================
//  HELPERS (no need to touch)
// =====================================================================

function findCsvFiles_(words) {
  const out = [];
  let it;
  const folders = DriveApp.getFoldersByName(CONFIG.folderName);
  if (folders.hasNext()) {
    it = folders.next().getFiles();
  } else {
    it = DriveApp.searchFiles("title contains '.csv' and trashed = false");
  }
  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName().toLowerCase();
    if (!name.endsWith('.csv')) continue;
    if (words.some(function (w) { return name.indexOf(w) !== -1; })) out.push(f);
  }
  return out;
}

/** Turns one CSV into a Map of yyyy-mm-dd -> daily value, auto-detecting columns. */
function perDay_(file, m) {
  const rows = Utilities.parseCsv(file.getBlob().getDataAsString());
  if (rows.length < 2) return new Map();
  const header = rows[0].map(function (h) { return String(h).toLowerCase(); });
  const body = rows.slice(1);

  const dateCol = m.dateCol !== null ? m.dateCol : detectDateCol_(body);
  const valueCol = m.valueCol !== null ? m.valueCol : detectValueCol_(header, body, dateCol, m.valueHints);
  if (dateCol < 0 || valueCol < 0) {
    console.warn(m.label + ': could not detect columns in ' + file.getName() + ' — run inspectFiles');
    return new Map();
  }

  const buckets = new Map();
  body.forEach(function (r) {
    const d = normaliseDate_(r[dateCol]);
    const v = toNumber_(r[valueCol]);
    if (!d || v === null) return;
    if (!buckets.has(d)) buckets.set(d, []);
    buckets.get(d).push(v);
  });

  const result = new Map();
  buckets.forEach(function (vals, d) {
    const sum = vals.reduce(function (a, b) { return a + b; }, 0);
    result.set(d, Math.round((m.agg === 'avg' ? sum / vals.length : sum) * 10) / 10);
  });
  return result;
}

function detectDateCol_(body) {
  const sample = body.slice(0, 20);
  for (let c = 0; c < (body[0] || []).length; c++) {
    const hits = sample.filter(function (r) { return normaliseDate_(r[c]); }).length;
    if (hits >= sample.length * 0.8) return c;
  }
  return -1;
}

function detectValueCol_(header, body, dateCol, hints) {
  const sample = body.slice(0, 20);
  const numeric = function (c) {
    return c !== dateCol && sample.filter(function (r) { return toNumber_(r[c]) !== null; }).length >= sample.length * 0.8;
  };
  for (let h = 0; h < hints.length; h++) {
    for (let c = 0; c < header.length; c++) {
      if (header[c].indexOf(hints[h]) !== -1 && numeric(c)) return c;
    }
  }
  for (let c = 0; c < header.length; c++) if (numeric(c)) return c;
  return -1;
}

/** 2026-10-10, 2026-10-10 14:00, 10/10/2026, 10.10.2026 (UK day-first) → yyyy-mm-dd */
function normaliseDate_(raw) {
  if (raw === undefined || raw === null) return null;
  const s = String(raw).trim();
  let x = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (x) return x[1] + '-' + pad_(x[2]) + '-' + pad_(x[3]);
  x = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
  if (x) return x[3] + '-' + pad_(x[2]) + '-' + pad_(x[1]);
  return null;
}

/** Numbers like 8432, 72.5, 72,5 and durations like 7:30 or 07:30:00 (→ minutes). */
function toNumber_(raw) {
  if (raw === undefined || raw === null) return null;
  const s = String(raw).trim();
  const t = s.match(/^(\d+):(\d{2})(?::(\d{2}))?$/);
  if (t) return Number(t[1]) * 60 + Number(t[2]) + (t[3] ? Number(t[3]) / 60 : 0);
  if (!/^-?\d+([.,]\d+)?$/.test(s)) return null;
  return parseFloat(s.replace(',', '.'));
}

function pad_(n) { return ('0' + n).slice(-2); }

/** Adds new days, updates changed ones (e.g. today's steps), never deletes. */
function mergeIntoSheet_(ss, name, incoming) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange('A:A').setNumberFormat('@');
    sh.appendRow(['date', 'value']);
  }
  const data = sh.getDataRange().getValues();
  const existing = new Map();
  for (let i = 1; i < data.length; i++) existing.set(String(data[i][0]), i + 1);

  const add = [];
  incoming.forEach(function (v, d) {
    if (existing.has(d)) {
      const row = existing.get(d);
      if (data[row - 1][1] !== v) sh.getRange(row, 2).setValue(v);
    } else {
      add.push([d, v]);
    }
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 2).setValues(add);
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).sort(1);
  return add.length;
}
