# Changelog

All notable changes to this project are documented here.

## [Unreleased]

- Next version TBD.

## [v0.3.0] - 2026-04-25

### Added

- Added dark HTML chip dashboard output:
  - `output/latest_visual_dashboard.html`
  - `output/archive/YYYYMMDD/visual_dashboard_YYYYMMDD.html`
- Added `twse_foreign_report/visual_report.py` for HTML report generation.
- Added 10-week shareholder visualization for multiple target stocks on the same page.
- Added interactive dashboard tabs:
  - `持股比`
  - `金字塔`
  - `股東均張`
  - `股東人數`
- Added interactive pyramid date switching. Selecting a week updates holder counts, holding ratios, and weekly changes by holding bucket.
- Added cumulative TDCC shareholder detail history:
  - `output/history/shareholders_detail_history.csv`
- Added TWSE `STOCK_DAY` quote/volume history cache:
  - `output/history/price_history.csv`
- Added shareholder-average-lots view:
  - `股東均張 = 集保總股數 / 總股東人數 / 1000`
  - `均張增減 = 本週股東均張 - 前一週股東均張`
  - `週轉率 = 週期間成交股數合計 / 當週集保總股數 * 100`
- Added stock price and trading volume columns to the shareholder-count view.

### Changed

- Changed default `--holders-weeks` from `3` to `10`.
- Dashboard tables now show the latest date first, then older dates below.
- Dashboard now uses cumulative history files so new weekly data can be appended over time.
- Price color convention now follows Taiwan/China market convention:
  - red for up
  - green for down
- Removed non-real bid/ask placeholder values from the visual dashboard.
- Historical prices are no longer backfilled from the latest close. They are sourced from TWSE `STOCK_DAY`.
- If a TDCC weekly date is not a trading day, the dashboard uses the latest previous trading-day close and preserves `price_date` in the cache.
- `.gitignore` now ignores generated HTML report files.

### Notes

- TDCC holder counts, shares, and holding ratios are sourced from the TDCC shareholder distribution table.
- Prices, price changes, and trading volumes are sourced from TWSE `STOCK_DAY`.
- The dashboard turnover ratio is this project's explicit estimate: `weekly traded shares / TDCC total shares * 100`. Broker apps may use a different private definition.

## [v0.2.0] - 2026-04-20

### Added

- Added TDCC shareholder distribution report support.
- Added stock-code and stock-name target resolution.
- Added weekly shareholder change outputs:
  - detail
  - recent two-week changes
  - multi-week changes
  - summary
- Added standalone shareholder CSV and Excel outputs.
- Added `holders_*` worksheets to the integrated `latest_report.xlsx` when `--holders-targets` is used.
- Added project documentation and release notes.

### Changed

- Extended the original TWSE foreign-investor report workflow with optional shareholder-distribution reporting.
- Updated README usage examples for TDCC shareholder mode.

## [v0.1.0] - 2026-04-15

### Added

- Initial TWSE foreign-investor daily report workflow.
- Added TWSE TWT38U foreign buy/sell data retrieval.
- Added TWSE quote retrieval.
- Added CSV and Excel exports.
- Added top buy, top sell, full report, summary, and summary history outputs.
