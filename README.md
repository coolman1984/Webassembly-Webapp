# BOM Confirmation Plan dashboard

## For users
1. **Double-click `start_dashboard.bat`.** The dashboard opens in your browser (keep the black window open while you use it).
2. Open the **Update Data** tab, attach the 3 Excel files (any order), click **Process & update dashboard**.
   It takes about 40 seconds. Every number, chart, table and the project filter refresh.
3. Next time, just start it again - the last data is remembered.

**You need:** Windows and desktop Microsoft Excel. **You do not need to install anything else** - Python and every
library are inside the `runtime` folder. Copy the *whole* project folder to move it to another PC.

If something goes wrong the page says what (for example *"'X.xlsx' looks like the SEEG SOP file but is missing column(s): Cate"*).
Your previous data always stays in place when a refresh is rejected. Technical details: `data\pipeline.log`.

## What the program checks on the 3 files
* Which file is which is decided by **content** (column names), not file names - the timestamp in a name may change freely.
  Moving, inserting or adding columns is fine; header case, spacing and hidden characters are ignored; title rows above the
  headers are fine; week columns may be in any order.
* Rejected with a specific message: a missing/renamed required column, the same file twice, two files of the same kind,
  a missing kind, a password-protected/corrupt/non-Excel file, an empty data sheet, two files with the same name.
* Accepted but reported as **Notes** after the refresh: duplicate model rows, unknown SOP categories (only `Ship`/`SOP` are
  used), invalid ISO week columns, unreadable dates (shown as unknown, never guessed - `06/07/2026` is ambiguous so it is
  not interpreted), BOM models absent from the SOP sheet, a sudden drop of >30% in model count.
* Excel details handled: 1904 date system, numbers stored as text, error cells (#N/A), hidden/filtered rows, sheets whose
  formatting reaches row 1,048,576, files locked by another Excel, stale external links, macros never run, no prompts.

## Architecture
`3 Excel files -> Excel COM (one DispatchEx instance, ReadOnly, UpdateLinks=0, bulk Value2, row-chunked)`
`-> validate + normalise -> SQLite (data\dashboard.db, atomic swap) -> /api/data -> browser KPIs/filters/charts`

| Path | Role |
|---|---|
| `runtime\` | private Python 3.12 + pywin32 (no install; run in isolated mode `-I`) |
| `pipeline\excel_com.py` | COM session: hardening, watchdog, orphan cleanup, real last-cell detection |
| `pipeline\sources.py` | content-based recognition/validation + extraction |
| `pipeline\transform.py` | all business rules |
| `pipeline\store.py` | SQLite layer (versioned schema, atomic replace with retry) |
| `pipeline\run.py` | orchestration, data-quality checks, CLI |
| `server.py` | local service (127.0.0.1 only) |
| `verify.py` | regression check against the original dashboard data |
| `tests\` | edge-case, service, stress and browser tests |
| `tools\build_runtime.ps1` | (maintainers) rebuilds `runtime\` from an existing Python install |
| `backup\` | the untouched original dashboard |

Security of the local service: loopback only; every request needs a loopback `Host` header (DNS-rebinding defence);
state-changing calls also need a same-origin `Origin` and an `X-Requested-With` header; uploads are sanitised and
size-checked; nothing leaves the machine.

## The three sources
| Role | Detected by | Used for |
|---|---|---|
| **DASH** status workbook | headers `MODEL_CODE`, `MKT_PROJECT`, `PROJECT_NAME` | model list; `project`, `projectName` |
| **SEEG** SOP workbook | a sheet with `Item, Project, Inch, Cate, Version` + weekly columns like `202601` | `version`, `firstSop`, `mp`, SOP models |
| **New Model** workbook | two-row header: `Project/Model/Type/Inch` + `BOM > HQ/LOCAL` + `SET PLANT > MP` | BOM scope (`Plan` rows), BOM dates, `inch` |

Join key = model code (trimmed, upper-case, invisible characters removed, numeric codes normalised).

## Calculation rules (verified against the original data by `verify.py`)
* `version` latest SOP version of the item; `firstSop` earliest version with Cate `SOP`.
* `mp` first *week* (chronological, regardless of column order) with qty>0 in the latest `Ship` version; else the latest `SOP`
  version; else the New Model `MP` date. `mpFrom` records which one was used.
* `hqTarget` = MP - 13 weeks, `localTarget` = MP - 12 weeks (ISO weeks, Monday dates). The week counts live in
  `pipeline/transform.py` (`HQ_WEEKS`, `LOCAL_WEEKS`, `SOP_LEAD_WEEKS`) and the page reads them from the data.
* `hq/localStatus` MATCH / LATER / EARLIER from the week difference between BOM plan and target; N/A without a date.
* `firstAppearMpGap` = weeks from the `firstSop` version week to the MP week (BOM models only); **SOP issue** when < 12.
* All Models scope: DASH models + SOP models with qty>0 in any version + BOM models.
* Inch is normalised (`43`, `43.0` -> `43`); `-` is treated as "unknown" for project and inch.

## Dashboard logic
* **On time** = LOCAL BOM planned no later than the target week (MATCH *or* EARLIER). *Late* = LATER. The timing donut is
  exclusive (on time + late + no plan = scope).
* **Planning week** = the ISO week of today's date (no longer a hard-coded label).
* **Progress** of each BOM model (exclusive): *Confirmed* (the New Model `Actual` row holds a date or OK/Done/Y),
  *Plan passed* (plan week before this week, not confirmed), *This week*, *Upcoming*, *No plan*.
* **Delay cause**: a late model whose delay is fully explained by the SOP having moved MP earlier than the New Model file's
  `SET PLANT MP` is marked *MP pull-in*; the rest are late BOM plans.
* **Data & logic checks**: MP mismatch between SOP and New Model, HQ BOM planned after LOCAL BOM, missing plan/MP,
  plan weeks passed, due within 4 weeks. Every check, KPI card, timeline bar and project row opens the matching model list.
* Tables: click a header to sort, chips for quick filters, 35/100/250 rows per page, click a row for a detail panel that
  explains the calculation. CSV export includes every field, opens correctly in Excel (UTF-8 BOM) and neutralises formulas.

## Database (`data\dashboard.db`, schema 3)
* Datasets: `bom_data`, `master_data` (+ raw source tables `dash`, `newmodel_plan` incl. Actual cells, `sop_item`, `sop_qty`).
* `refresh_history`: headline KPIs of every refresh (kept across refreshes, last 200) - shown on the Update Data page and as
  the on-time trend.
* `bom_changes`: model-level changes between consecutive refreshes (added / removed / MP, plan, status, actual changed) -
  shown as "Since last refresh" and the *Changed* filter.
* A schema-2 database from the previous version is upgraded automatically at start-up without needing Excel.

## Tests (maintainers)
```
runtime\python.exe -I tests\test_logic.py -v        # rules, change tracking, history, migration (no Excel, seconds)
runtime\python.exe -I tests\test_edge_cases.py -v   # edge cases incl. fuzzed storage variants (real Excel)
runtime\python.exe -I tests\test_service.py -v      # security, hostile uploads, concurrency, crash recovery, ports
runtime\python.exe -I tests\test_stress.py -v       # 500k-row SOP, 60k models, Excel's 1,048,576-row limit
python tests\test_browser.py                        # 300k-row page, hostile strings, CSV injection (needs Playwright)
runtime\python.exe -I verify.py FILE FILE FILE      # equality with the original dashboard data
```

## Limits
* Needs desktop Excel (COM). Without it the page says so and keeps the old data.
* The planning week follows the PC clock.
* The dashboard's table columns are fixed; a new Excel column is ignored unless the page is extended.

## Further reading
* `docs/EXPERIENCE.md` - narrative log of how this was built, what was found and fixed during hardening, and
  what a future maintainer should know.
* `.claude/skills/excel-com-automation/` - a Claude Code skill distilling the reusable lessons (Excel COM
  hardening, self-contained Windows packaging, local-service security, testing strategy) for future projects.
  Not specific to this repo; safe to copy into other projects.
