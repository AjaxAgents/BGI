# BGI

Budgeting General Interface for STO Ajax Building Company.

## Underbilling Report (prototype, v0.1.0)

A single Windows EXE that turns the month's underbilling inputs into a report people can read:

- **Schedule** mirrors the month-end Adjusted Underbillings schedule: the total, where it sits, what moved since last month, and every job line.
- **How it was calculated** opens when you click a line. It shows each formula with the numbers filled in and the Excel column each number came from, plus AP by category (GCs, GRs, Trades, Insurance).
- **Data check** re-runs every rule on every row of the inputs and says which ones hold.

Every figure is read from an input file or produced by a fixed formula listed in [docs/RULES.md](docs/RULES.md). Nothing is estimated.

### Use it

1. Put `BGI-Underbilling.exe` in a folder with the two input files:
   - the Excel export (`UNDERBILLING_CASHPOSITION_RPT.xlsx`, any `.xlsx` with "underbilling" in the name)
   - the schedule file (any `.csv` with "schedule" in the name, layout below)
2. Double-click the EXE. It loads the newest of each and opens the report.

You can also drag files onto the window. With only the Excel export loaded, the schedule page is calculated from the export alone.

`BGI-Underbilling.html` is the same report as one file with no EXE: open it in Edge or Chrome and drag the two files in.

### Schedule file layout

One row per line of the schedule. Amount columns are named for their month-end date; the last one is the current month.

```
row_type,section,job_code,job_name,area_manager,project_executive,underbilling_2026-09-30,underbilling_2026-10-31,const_start,net_billing_2026-10,billed_to_date_pct,comment
job,IN_PROCESS,1000,Example Tower,A. Manager,P. Executive,100000,150000,2025-01-15,25000,45.0,
small,SMALL,,Jobs with underbillings less than $10000,,,5000,4000,,,,
subtotal,IN_PROCESS,,Subtotal,,,100000,150000,,,,
total,,,Total,,,105000,154000,,,,
```

`row_type` is `job`, `small`, `subtotal` or `total`. `section` is `NOT_STARTED`, `SMALL` or `IN_PROCESS`. Subtotal and total rows are optional; when present the report checks that the lines add up to them.

### Security

- The EXE listens on `127.0.0.1` only, on a random port, and makes no outbound connections.
- Every request needs a random per-run token and the exact `Host` header, so other sites and local pages cannot read the data.
- It can serve exactly two files from its own folder (the newest export and the newest schedule). There is no path parameter.
- The page runs under a strict Content-Security-Policy with a per-response nonce. All text from input files is escaped.
- No third-party code: the Excel reader, the charts and the server are in this repository. Go standard library only.
- The EXE is not code-signed yet. Windows SmartScreen may warn the first time it runs from a download.

### Data stays out of this repository

This repository is public. Never commit exports, schedules, PDFs or emails. `.gitignore` blocks the usual file types, and the tests use made-up jobs.

### Build

Needs Go 1.24+ and Node 18+.

```
./scripts/build.sh      # tests, bundles the page, cross-compiles dist/BGI-Underbilling.exe
```

### Layout

```
web/src/xlsx.js       dependency-free .xlsx reader
web/src/engine.js     the rules: parsing, formulas, checks, AP categories
web/src/app.js        the three views
web/src/styles.css
cmd/bgi-underbilling  Go launcher that embeds and serves the page
tests/                node --test, synthetic fixtures
docs/RULES.md         every rule in plain language
```

### Roadmap

- Read the export straight from SharePoint instead of a dropped file (the page and engine stay as they are; only the loader changes).
- Keep each month's figures so the prior-month columns come from history instead of a schedule file.
- Code-sign the EXE.
