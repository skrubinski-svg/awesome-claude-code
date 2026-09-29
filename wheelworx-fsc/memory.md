# memory.md — Wheelworx FSC Report App

Project memory: decisions, constants, and what was learned during the build.
Source of truth for the spec: `WHEELWORX_FSC_REPORT_BUILD_SPEC.md` (provided by Steve).

## Constants (config.json; derived values are computed in code, never hard-coded)
- Band width $0.05; Rate Table $3.60 → $11.05 = 150 rows (3 blocks of 50)
- New Wheelworx: start $0.23 at $3.60, 5.0 MPG → **peg = 3.60 − 0.23 × 5.0 = $2.45**
- DAT flatbed: peg $1.25, 5.0 MPG
- **Gap vs DAT = (2.45 − 1.25) / 5 = $0.24/mile at every price**
- Step: $0.05 / 5.0 MPG = $0.01 per band
- 2025 table: $0.00 below $3.60; else cents = floor((band − 3.60) / 0.055) + 1
- Steve chose the $0.23 start, so the gap is $0.24. If he wants a $0.23 gap: peg $2.40, start $0.24
  (change `start_surcharge_at_table_start` to 0.24). Confirm with Steve before changing.

## Calculation decisions
- Round DOE price to the cent (ROUND_HALF_UP), then floor to the $0.05 band. All math in `Decimal`;
  config is parsed with `parse_float=Decimal`; EIA floats go through `repr()` then quantize to 0.001.
- Surcharges never negative. Formula continues above the table end.
- Effective date = EIA week + 1 day.
- `config.json` load fails if the two MPGs differ (the gap would not be constant).

## Report decisions
- Values only, no live formulas (spec §6.7).
- Report file name uses **today's date**; `_2`, `_3`… appended if it already exists; never overwrites.
  This is done in Python (`unique_report_path`), so both the CLI and the launcher are covered.
- YTD = every EIA week whose date is in the latest week's year. Label every 4th row (`Jan 5`, …),
  `""` otherwise; `tickLblSkip=1` is also set so renderers don't thin labels further.
- YTD has a "Gap vs DAT" column placed before the chart-label column.
- Rate Table side panel at Q/R (spec). Highlight = band of the latest week's rounded price.
  If the price is outside the printed table, no highlight and a red note in the panel.
- DAT Flatbed tab: 150 rows from the $1.25 peg ($1.25 → $8.70), same 3-block layout.
- 2025 tab: rows from `data/2025_table.json` (blocks of 37) + footer text stored in that JSON.
- Chart fonts: Arial, axis labels 11 pt (rotation 0), axis titles 12 pt, on both charts.
- Output check after save: re-open the file and compare every surcharge cell with Python. openpyxl
  does not load charts, so the chart check reads `xl/charts/chart*.xml` directly (3 `<c:ser>` each,
  referencing `12-Week FSC` / `YTD FSC`).

## `--check`
- Runs pytest on `tests/`. Pure golden tests always run. Real-data tests (golden weekly raw prices
  must match the EIA file; DAT monthly cross-check Mar/Apr/Aug 2026 ±1¢) use the live EIA download,
  or `--csv` when given. If the download fails they are SKIPPED and the reason is printed (`-rs`).

## Launcher
- `run_fsc_report.ps1` creates `.venv` on first run (prefers `py -3`), activates it, installs
  requirements, then runs `fsc_report.py --open`. If pip fails, `.venv` is removed so the next run retries.
- `run_fsc_report.bat` = double-click wrapper (`-ExecutionPolicy Bypass`, `pause` at the end).
- NOT yet tested on Windows (built in a Linux container without PowerShell).

## Build log / status
- 2026-09-29: first build in a cloud container. The container's network policy blocked eia.gov
  (proxy 403), so the live download was not tested yet. Steve will allow eia.gov.
- 2026-09-29: `Fuel_Surcharge_2025_3 (2).pdf` was not provided yet. `data/2025_table.json` is NOT
  created; Steve will upload the PDF for transcription. Until then
  `test_2025_rule_matches_every_published_row` fails and the report refuses to build (by design:
  no invented data).
- Charts were visually checked by rendering scratch data through LibreOffice (not a real report).
