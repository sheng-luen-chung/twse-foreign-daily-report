# Release Notes v0.2.0

## Summary

This release expands the project from a TWSE foreign-investor daily report tool into a dual-report workflow that also supports TDCC shareholder distribution analysis.

The main outcome is that we can now compare recent weekly changes in shareholder counts and holdings by position-size bucket, while still preserving the original foreign net buy / sell report flow.

## Highlights

- Added TDCC shareholder distribution analysis
  Query by stock code or stock name.
  Supports pulling the most recent N available weeks, with 3 weeks as the default.

- Added weekly shareholder change outputs
  New outputs include:
  - full shareholder detail
  - recent two-week change view
  - multi-week change view
  - shareholder summary

- Added recent two-week deltas in both people and lots
  The new `recent_changes` output shows:
  - shareholder count change
  - total shares change
  - total lots change
  This makes it easier to read short-term movement by holding bucket.

- Added integrated Excel workbook support
  When `--holders-targets` is used, the generated `latest_report.xlsx` now includes:
  - `holders_overview`
  - `holders_summary`
  - `holders_recent_changes`
  - `holders_changes`
  - `holders_detail`

- Added dedicated shareholder workbook and CSV outputs
  New latest files:
  - `latest_shareholders_detail.csv`
  - `latest_shareholders_recent_changes.csv`
  - `latest_shareholders_changes.csv`
  - `latest_shareholders_summary.csv`
  - `latest_shareholders_report.xlsx`

- Added project documentation
  - expanded `README.md`
  - added `PROJECT_OVERVIEW.md`

## Main Files Changed

- `twse_foreign_report/daily_report.py`
  Added shareholder-mode CLI integration and combined workbook generation.

- `twse_foreign_report/shareholder_distribution.py`
  Added the TDCC query flow, stock resolution, weekly comparison logic, and shareholder report export pipeline.

- `README.md`
  Added clearer run modes, output explanations, and field interpretation examples.

- `PROJECT_OVERVIEW.md`
  Added an end-to-end architectural overview of the project.

- `.gitignore`
  Added a few common Python development cache / artifact ignores for cleaner repository maintenance.

## Suggested Usage

Foreign-only report:

```bash
python -X utf8 main.py --date 20260417 --top 25 --outdir output
```

Foreign report plus shareholder distribution:

```bash
python -X utf8 main.py --date 20260419 --holders-targets 鴻海 臻鼎 --holders-weeks 3 --outdir output
```

## Version Intent

`v0.2.0` is an appropriate version tag because the project now includes a meaningful new feature set and expanded reporting outputs, while keeping the original foreign-report workflow compatible.
