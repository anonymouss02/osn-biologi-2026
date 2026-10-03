#!/usr/bin/env node
// Reads data/<year>.xlsx for every year, computes every derived figure, and renders one page per
// year (dist/<year>/index.html) from template.html, plus the newest year again at dist/index.html.
// Everything the page shows is derived here; the spreadsheets only ever hold raw marks + config.

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, 'data');
const DIST = path.join(ROOT, 'dist');

class DataError extends Error {}
const bad = (msg) => { throw new DataError(msg); };

// ---------- statistics ----------
// Sample standard deviation (n-1) — matches the convention the source data was scored with.
function zscores(values) {
  const n = values.length;
  if (n < 2) bad(`need at least 2 scores to compute a z-score, got ${n}`);
  const mean = values.reduce((a, b) => a + b, 0) / n;
  const sd = Math.sqrt(values.reduce((a, b) => a + (b - mean) ** 2, 0) / (n - 1));
  if (sd === 0) bad('every competitor scored identically, so a z-score is undefined');
  return values.map((v) => (v - mean) / sd);
}

// Assign z-scores back onto the objects that actually have a value for `read`.
function scoreInto(rows, read, write) {
  const sat = rows.filter((r) => read(r) != null);
  if (!sat.length) return;
  const z = zscores(sat.map(read));
  sat.forEach((r, i) => write(r, z[i]));
}

// Marks come out of a spreadsheet with a decimal place or two, but adding them in binary
// floating point leaves noise: 41 + 38.8 and 40.6 + 39.2 are both 79.8 yet differ by 1.4e-14.
// Left alone that noise silently decides ranks, so every total is snapped to 6 decimals first.
const snap = (v) => (v == null ? null : Math.round(v * 1e6) / 1e6);

// Standard competition ranking (1, 2, 2, 4) by a numeric key, descending.
// Genuinely tied competitors must share a rank: medals are handed out by rank, so splitting a
// tie on arbitrary ordering would be the difference between one competitor's gold and another's
// silver. Rows with no value stay explicitly unranked.
function rankBy(rows, key, rankKey) {
  for (const r of rows) r[rankKey] = null;
  const sat = rows.filter((r) => r[key] != null).sort((a, b) => b[key] - a[key]);
  sat.forEach((r, i) => {
    r[rankKey] = i > 0 && r[key] === sat[i - 1][key] ? sat[i - 1][rankKey] : i + 1;
  });
}

// ---------- spreadsheet ----------
function sheet(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) bad(`data.xlsx has no sheet named "${name}" (found: ${wb.SheetNames.join(', ')})`);
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

const str = (v) => String(v ?? '').trim();

function num(v, where) {
  if (v === '' || v == null) return null;
  const n = typeof v === 'number' ? v : Number(String(v).trim().replace(',', '.'));
  if (!Number.isFinite(n)) bad(`${where}: "${v}" is not a number`);
  return n;
}

function readExams(wb) {
  const rows = sheet(wb, 'Exams');
  if (!rows.length) bad('Sheet "Exams" is empty — it must list at least one exam');
  const seen = new Set();
  return rows.map((r, i) => {
    const at = `Sheet "Exams", row ${i + 2}`;
    const key = str(r.key);
    if (!key) bad(`${at}: "key" is required`);
    if (seen.has(key)) bad(`${at}: duplicate exam key "${key}"`);
    seen.add(key);
    const group = str(r.group).toLowerCase();
    if (group !== 'practical' && group !== 'theory') {
      bad(`${at}: "group" must be "practical" or "theory", got "${r.group}"`);
    }
    const max = num(r.max, `${at}, column "max"`);
    if (max == null || max <= 0) bad(`${at}: "max" must be a positive number`);
    return { key, label: str(r.label) || key, group, max, color: str(r.color) || '#8E8F91', tag: str(r.tag) };
  });
}

function readConfig(wb) {
  const cfg = {};
  for (const r of sheet(wb, 'Config')) {
    const k = str(r.key);
    if (k) cfg[k] = typeof r.value === 'string' ? r.value.trim() : r.value;
  }
  return cfg;
}

function readResults(wb, exams) {
  const rows = sheet(wb, 'Results');
  if (!rows.length) bad('Sheet "Results" is empty — add one row per competitor');

  const examKeys = new Set(exams.map((e) => e.key));
  const headers = Object.keys(rows[0]);
  const known = new Set(['code', 'name', 'country', ...examKeys]);
  for (const h of headers) {
    if (!known.has(h)) {
      bad(`Sheet "Results" has column "${h}", which is not "code"/"name"/"country" and not an exam key in the Exams sheet`);
    }
  }
  for (const e of exams) {
    if (!headers.includes(e.key)) bad(`Sheet "Results" is missing a column for exam "${e.key}" (${e.label})`);
  }

  const codes = new Set();
  return rows.map((r, i) => {
    const at = `Sheet "Results", row ${i + 2}`;
    const code = str(r.code);
    const name = str(r.name);
    if (!code) bad(`${at}: "code" is required`);
    if (!name) bad(`${at}: "name" is required`);
    if (codes.has(code)) bad(`${at}: duplicate code "${code}"`);
    codes.add(code);

    const s = { code, name, country: str(r.country) || '—', raw: {}, z: {} };
    for (const e of exams) {
      const v = num(r[e.key], `${at}, column "${e.key}"`);
      if (v != null && (v < 0 || v > e.max)) {
        bad(`${at}: "${e.key}" is ${v}, outside the 0–${e.max} range set in the Exams sheet`);
      }
      s.raw[e.key] = v;
    }
    return s;
  });
}

// ---------- the maths ----------
// Per exam: z of the raw mark. Per group: one composite z, then the overall score.
// A group composite is either the z of the summed z-scores ("z-sum", right when exams have
// unequal maxes) or the z of the summed raw marks ("raw-sum", right when they are comparable).
function compute(students, exams, cfg) {
  for (const e of exams) {
    scoreInto(students, (s) => s.raw[e.key], (s, z) => { s.z[e.key] = z; });
  }

  const groups = ['practical', 'theory'];
  for (const g of groups) {
    const gExams = exams.filter((e) => e.group === g);
    if (!gExams.length) continue;
    const mode = str(cfg[`composite.${g}`]) || (g === 'theory' ? 'raw-sum' : 'z-sum');
    if (mode !== 'z-sum' && mode !== 'raw-sum') {
      bad(`Config "composite.${g}" must be "z-sum" or "raw-sum", got "${mode}"`);
    }
    // A competitor only gets a composite if they sat every exam in the group.
    const total = (s) => {
      let sum = 0;
      for (const e of gExams) {
        const v = mode === 'z-sum' ? s.z[e.key] : s.raw[e.key];
        if (v == null) return null;
        sum += v;
      }
      return sum;
    };
    for (const s of students) { s[`_${g}`] = snap(total(s)); s[`z${g}`] = null; }
    scoreInto(students, (s) => s[`_${g}`], (s, z) => { s[`z${g}`] = z; });
  }

  const base = num(cfg['score.base'], 'Config "score.base"') ?? 50;
  const scale = num(cfg['score.scale'], 'Config "score.scale"') ?? 20;
  for (const s of students) {
    const parts = groups.map((g) => s[`z${g}`]).filter((v) => v != null);
    s.overall = parts.length === groups.length
      ? snap(base + scale * (parts.reduce((a, b) => a + b, 0) / parts.length))
      : null;
  }

  rankBy(students, 'overall', 'pos');
  rankBy(students, 'zpractical', 'pracRank');
  rankBy(students, 'ztheory', 'theoRank');
  students.sort((a, b) => (a.pos ?? Infinity) - (b.pos ?? Infinity) || a.code.localeCompare(b.code));

  // Gap to the competitor one place below — the "leads #N by X" line in the modal.
  for (let i = 0; i < students.length; i++) {
    const s = students[i], next = students[i + 1];
    s.leads = s.overall != null && next?.overall != null ? s.overall - next.overall : null;
  }
  return students;
}

function assignMedals(students, cfg) {
  const cut = (k) => {
    const v = num(cfg[`medal.${k}_through`], `Config "medal.${k}_through"`);
    if (v != null && (!Number.isInteger(v) || v < 0)) bad(`Config "medal.${k}_through" must be a whole rank, got ${v}`);
    return v;
  };
  const tiers = [['gold', cut('gold')], ['silver', cut('silver')], ['bronze', cut('bronze')], ['merit', cut('merit')]];
  let prev = 0;
  for (const [name, through] of tiers) {
    if (through == null) continue;
    if (through < prev) bad(`Config "medal.${name}_through" (${through}) must be at least the previous tier's cut-off (${prev})`);
    prev = through;
  }
  for (const s of students) {
    s.medal = 'none';
    if (s.pos == null) continue;
    for (const [name, through] of tiers) {
      if (through != null && s.pos <= through) { s.medal = name; break; }
    }
  }
  // Tiers left blank in the spreadsheet are dropped from the filter bar and legend entirely.
  return tiers.filter(([, t]) => t != null).map(([name]) => name);
}

// ---------- render ----------
const round = (v, d = 3) => (v == null ? null : Math.round(v * 10 ** d) / 10 ** d);

// ---------- years ----------
// One workbook per year — data/2026.xlsx, data/2025.xlsx … — each complete with its own Exams and
// Config sheets, because papers, max marks and medal cut-offs can change from one year to the next.
const YEAR_FILE = /^(\d{4})\.xlsx$/;

function findYears() {
  if (!fs.existsSync(DATA_DIR)) {
    bad("there is no data/ folder — put each year's spreadsheet there, named by year: data/2026.xlsx");
  }
  // Excel drops a "~$2026.xlsx" lock file beside any workbook it has open; that is never data.
  const files = fs.readdirSync(DATA_DIR)
    .filter((f) => f.endsWith('.xlsx') && !f.startsWith('~$') && !f.startsWith('.'));
  const misnamed = files.find((f) => !YEAR_FILE.test(f));
  if (misnamed) bad(`data/${misnamed} is not named after a year — rename it to the 4-digit year it covers, e.g. data/2026.xlsx`);
  if (!files.length) bad('data/ has no spreadsheets — add one per year, named like data/2026.xlsx');
  return files.map((f) => f.match(YEAR_FILE)[1]).sort().reverse(); // newest first
}

function readYear(year) {
  const file = `data/${year}.xlsx`;
  try {
    return readWorkbook(XLSX.readFile(path.join(ROOT, file)));
  } catch (err) {
    if (err instanceof DataError) err.file = file;
    throw err;
  }
}

function readWorkbook(wb) {
  const exams = readExams(wb);
  const cfg = readConfig(wb);
  const students = compute(readResults(wb, exams), exams, cfg);
  const tiers = assignMedals(students, cfg);

  const medalCounts = {};
  for (const m of ['gold', 'silver', 'bronze', 'merit', 'none']) {
    medalCounts[m] = students.filter((s) => s.medal === m).length;
  }

  const data = {
    meta: {
      total: students.length,
      countries: new Set(students.map((s) => s.country)).size,
      eyebrow: str(cfg['event.eyebrow']),
      title: str(cfg['event.title']) || 'Results Overview',
      subtitle: str(cfg['event.subtitle']),
      metricsLabel: str(cfg['event.metrics_label']),
      footer: str(cfg['event.footer']),
      // What the "country" column is called. Countries for an international event, provinces
      // or states for a national one — it is only ever a grouping label.
      region: str(cfg['label.region']) || 'Country',
      regionPlural: str(cfg['label.region_plural']) || 'countries',
    },
    exams,
    tiers,
    medalCounts,
    students: students.map((s) => ({
      code: s.code,
      name: s.name,
      country: s.country,
      medal: s.medal,
      pos: s.pos,
      raw: Object.fromEntries(exams.map((e) => [e.key, round(s.raw[e.key], 4)])),
      z: Object.fromEntries(exams.map((e) => [e.key, round(s.z[e.key])])),
      zprac: round(s.zpractical),
      ztheo: round(s.ztheory),
      overall: round(s.overall, 2),
      leads: round(s.leads, 2),
      pracRank: s.pracRank,
      theoRank: s.theoRank,
    })),
  };

  return { data, pageTitle: str(cfg['page.title']) || data.meta.title };
}

function render(template, { data, pageTitle }, nav) {
  return template
    .replace('{{PAGE_TITLE}}', escapeHtml(pageTitle))
    // Year pages sit one folder down (2025/index.html) and the root copy doesn't, so asset paths
    // and year-switch links carry a prefix that makes them resolve from either place.
    .replaceAll('{{BASE}}', nav.base)
    // JSON goes inside a <script>; only "</" can break out of it.
    .replace('{{DATA}}', () => JSON.stringify({ ...data, nav }).replace(/<\//g, '<\\/'));
}

function build() {
  const years = findYears();
  // Read and check every year before writing anything, so one bad spreadsheet can't leave a
  // half-rebuilt site behind.
  const pages = years.map((year) => ({ year, ...readYear(year) }));
  const template = fs.readFileSync(path.join(ROOT, 'template.html'), 'utf8');

  fs.rmSync(DIST, { recursive: true, force: true });
  const write = (rel, html) => {
    const out = path.join(DIST, rel);
    fs.mkdirSync(path.dirname(out), { recursive: true });
    fs.writeFileSync(out, html);
  };
  for (const p of pages) {
    write(`${p.year}/index.html`, render(template, p, { year: p.year, years, base: '../' }));
  }
  // The bare site URL always shows the newest year.
  write('index.html', render(template, pages[0], { year: pages[0].year, years, base: '' }));

  if (fs.existsSync(path.join(ROOT, 'assets'))) {
    fs.cpSync(path.join(ROOT, 'assets'), path.join(DIST, 'assets'), { recursive: true });
  }
  // Tells GitHub Pages to serve the files as-is instead of running them through Jekyll.
  fs.writeFileSync(path.join(DIST, '.nojekyll'), '');

  for (const p of pages) {
    const { meta, exams, tiers, medalCounts } = p.data;
    const kb = (fs.statSync(path.join(DIST, p.year, 'index.html')).size / 1024).toFixed(0);
    console.log(`data/${p.year}.xlsx -> dist/${p.year}/index.html  ${kb} KB${p === pages[0] ? '  (also the site root)' : ''}`);
    console.log(`  ${meta.total} competitors · ${meta.countries} ${meta.regionPlural} · ${exams.length} exams`);
    console.log(`  ${tiers.map((t) => `${t} ${medalCounts[t]}`).join(' · ')} · no medal ${medalCounts.none}`);
  }
  return pages;
}

const escapeHtml = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

if (require.main === module) {
  try {
    build();
  } catch (err) {
    if (err instanceof DataError) {
      console.error(`\n  Build stopped — problem in ${err.file ?? 'the data/ folder'}:\n  ${err.message}\n`);
      process.exit(1);
    }
    throw err;
  }
}

module.exports = { build, zscores, compute, assignMedals, readExams, readResults, readConfig };
