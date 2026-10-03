# Results site — GitHub Pages edition

A static competition-results page — podium, score histogram, searchable/sortable ranking table,
and a per-competitor exam breakdown — generated from one Excel file per year, with a **Year**
dropdown to switch between them.

```bash
nvm use           # Node 22, from .nvmrc — older Node versions cannot run build.js
npm install
npm run build     # data/<year>.xlsx  ->  dist/
npm test          # check the scoring maths
npm run pages     # build, then copy dist/ into docs/ for publishing
```

---

## Years

Each year is one spreadsheet in `data/`, named by the year it covers:

```
data/
├── 2026.xlsx
└── 2025.xlsx
```

Each one becomes its own page, and the newest year is also served at the bare URL:

| URL | shows |
|---|---|
| `…/osn-biologi-2026/` | the newest year |
| `…/osn-biologi-2026/2026/` | 2026 |
| `…/osn-biologi-2026/2025/` | 2025 |

The **Year** dropdown in the header lists every year found in `data/`, newest first. Choosing
one opens that year's page.

Every workbook is complete on its own, with its own `Results`, `Exams` and `Config` sheets. So
each year can have different papers, max marks, medal cut-offs and wording.

### Adding a year

1. Copy last year's workbook as a starting point, for example `data/2026.xlsx` → `data/2027.xlsx`.
2. Replace the rows in its `Results` sheet with the new year's competitors.
3. Update its `Exams` sheet if the papers or max marks changed.
4. Update its `Config` sheet: `page.title`, `event.eyebrow` and `event.footer` (they mention the
   year and city), plus the `medal.*_through` cut-offs if they changed.
5. Publish (below).

The filename must be exactly the 4-digit year. A name like `2027 (1).xlsx` stops the build with
a message saying which file to rename. Excel's own `~$2027.xlsx` lock files are ignored.

---

## Publishing

Pages for this repo is set to **Deploy from a branch → `main` / `/docs`**, so the site is whatever
is committed in `docs/`. After any change to a spreadsheet, the template or the images:

```bash
npm run pages
git add -A
git commit -m "update results"
git push
```

`npm run pages` rebuilds every year and replaces `docs/` with the result. If you skip it, the
push still goes through, but the live site keeps showing the old `docs/`.

### Automatic builds (once GitHub Actions is available on the account)

`.github/workflows/deploy.yml` can do the build on GitHub instead: on every push it runs
`npm ci`, `npm test` and `npm run build`, then publishes `dist/`. To use it, set
**Settings → Pages → Source → GitHub Actions**. From then on a plain `git push` publishes, and
`docs/` is no longer used.

### Public repository

Free GitHub Pages requires a public repository, so every `data/*.xlsx` file, with the raw marks
for every competitor in every year, is publicly downloadable from the repo.

### Paths

The site lives under a subpath (`username.github.io/repo-name/`), and year pages sit one folder
deeper. Every asset and link in the page is relative, with the build adding the `../` that year
pages need, so it works at any depth with no configuration.

---

## Editing a year's data

You only ever enter raw marks — every z-score, rank, medal and overall score is recomputed at
build time.

### Sheet `Results` — one row per competitor

| code | name | country | mb | ap | am | pc | ta | tb |
|---|---|---|---|---|---|---|---|---|
| KPR-S1 | DERICKSON LIE | Prov. Kepulauan Riau | 37.7 | 48 | 73 | 127 | 24.8 | 34.2 |

`code`, `name` and `country` are fixed. After them comes **one column per exam**, its header
matching a `key` from the `Exams` sheet. Leave a cell blank if the competitor did not sit that exam.

The part of `code` before the first `-` is shown as the region tag (`KPR-S1` → `KPR`).

The column is named `country` for historical reasons but it is just a grouping label. Set
`label.region` / `label.region_plural` in `Config` to display it as Province, State, School, or
anything else. The column header itself must stay `country`.

### Sheet `Exams` — what the exams are

| key | label | group | max | color | tag |
|---|---|---|---|---|---|
| mb | Biologi Molekuler dan Biokimia | practical | 100 | #408DC4 | LAB1 |
| ta | Theory A | theory | 50 | #E8BC05 | |

- **`group`** must be `practical` or `theory`.
- **`max`** is the paper's maximum mark. It drives the progress bars and range-checks your data.
- **`tag`** is the small coloured pill in the breakdown; leave blank for a plain dot.

Add or remove rows freely. The table columns, the modal and the footer text all follow.

### Sheet `Config` — copy and rules

| key | meaning |
|---|---|
| `page.title` | browser tab title |
| `event.eyebrow` | small green line above the heading |
| `event.title` | the big heading |
| `event.subtitle` | line under the heading |
| `event.metrics_label` | caption over the medal counts |
| `event.footer` | sentence under the table |
| `label.region` | what the `country` column is called — `Country`, `Province`, `State`… |
| `label.region_plural` | its plural, used in "60 competitors · 17 provinces" |
| `medal.gold_through` | best rank through which Gold is awarded (`5` → ranks 1–5) |
| `medal.silver_through` | …through which Silver is awarded (`15` → ranks 6–15) |
| `medal.bronze_through` | …through which Bronze is awarded (`30` → ranks 16–30) |
| `medal.merit_through` | optional extra tier; leave blank and everyone past Bronze gets no medal |
| `score.base` / `score.scale` | overall score = `base + scale × mean(practical z, theory z)` |
| `composite.practical` | `z-sum` or `raw-sum` (see below) |
| `composite.theory` | `z-sum` or `raw-sum` |

---

## How the scoring works

1. **Per exam** — `z = (mark − mean) / stdev`, using the sample standard deviation (n−1), over
   everyone who sat that exam.
2. **Per group** — one composite z per group, controlled by `composite.<group>`:
   - `z-sum` — add up the group's *z-scores*, then z-score that sum. Weights each exam equally,
     which suits exams with different maximums or spreads.
   - `raw-sum` — add up the group's *raw marks*, then z-score that sum. Suits papers that are
     directly comparable.
3. **Overall** — `base + scale × mean(practical z, theory z)`, so an average competitor scores
   exactly `base`.
4. **Rank** — by overall score, descending.
5. **Medals** — by rank, using the `medal.*_through` cut-offs.

Scores are computed within a year only; years are never compared or mixed.

A competitor who misses any exam in a group gets no composite for that group, and therefore no
overall score and no rank. They still appear in the table, marked `—`.

**Ties share a rank.** Two identical scores both get rank 5, and the next competitor is 7th.
Medals are handed out by rank, so splitting a tie would decide Gold versus Silver on spreadsheet
row order.

---

## When the build fails

The build refuses to produce pages from data it does not understand, and says which file and
cell to look at:

```
Build stopped — problem in data/2025.xlsx:
Sheet "Results", row 3, column "mb": "abc" is not a number
```

It checks for misnamed files in `data/`, unknown columns, missing exam columns, duplicate
competitor codes, missing names, non-numeric marks, and marks outside `0…max`. Every year is
checked before anything is written, so a failed build leaves the previous `dist/` and `docs/`
untouched.

---

## Swapping the branding

`assets/logo.png` and `assets/garland.png` are the header logo and the decorative strip beneath it,
shared by every year. Replace the files and nothing else needs changing.

> `assets/garland.png` is the IBO 2026 event's own artwork, carried over from the site this page
> was modelled on.

---

## Files

| | |
|---|---|
| `data/<year>.xlsx` | one workbook per year — the only files you edit day to day |
| `build.js` | reads every workbook, computes everything, writes `dist/` |
| `template.html` | page markup, CSS and browser JS — edit to restyle |
| `verify.js` | `npm test`; checks the maths, including against 302 rows of real published results |
| `docs/` | the published copy of `dist/`, refreshed by `npm run pages` |
| `.github/workflows/deploy.yml` | automatic build-and-publish, for when Actions is available |
| `.nvmrc` | pins Node 22 for `nvm use` |
