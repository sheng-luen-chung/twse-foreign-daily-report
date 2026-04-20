# Changelog

本檔案用來記錄這個專案的重要版本變更。

格式參考 `Keep a Changelog`，版本號以 Git tag 為主。

## [Unreleased]

- 尚未整理成下一版的正式變更

## [v0.2.0] - 2026-04-20

### Added

- 新增 TDCC 集保戶股權分散表分析功能
- 新增以股票名稱或代號查詢股東分散資料的流程
- 新增最近兩週股東人數增減與張數增減輸出
- 新增股東分散表專用 CSV：
  - `latest_shareholders_detail.csv`
  - `latest_shareholders_recent_changes.csv`
  - `latest_shareholders_changes.csv`
  - `latest_shareholders_summary.csv`
- 新增股東分散表專用 Excel：
  - `latest_shareholders_report.xlsx`
- 新增將股東分散表工作表整合進 `latest_report.xlsx` 的能力
- 新增專案文件：
  - `PROJECT_OVERVIEW.md`
  - `RELEASE_NOTES_v0.2.0.md`

### Changed

- `main.py` / `daily_report.py` 現在支援外資模式與股東分散模式的雙流程
- `latest_report.xlsx` 在使用 `--holders-targets` 時會額外包含：
  - `holders_overview`
  - `holders_summary`
  - `holders_recent_changes`
  - `holders_changes`
  - `holders_detail`
- `README.md` 補充了：
  - 外資模式與股東分散模式的使用方式
  - `recent_changes` 欄位解讀
  - 以鴻海為例的實際解讀方式
- `.gitignore` 補上常見 Python 開發暫存與快取忽略規則

### Notes

- 這版是專案從「TWSE 外資買賣超工具」擴充到「外資日報 + TDCC 股東分散分析工具」的重要版本
- Git tag：`v0.2.0`
- GitHub Release 已建立

## [v0.1.0] - 2026-04-15

### Added

- 初始版 TWSE 外資買賣超日報工具
- 支援抓取 TWSE 外資及陸資買賣超資料
- 支援抓取 TWSE 每日收盤行情資料
- 產生：
  - 買超排行
  - 賣超排行
  - 完整明細
  - 摘要
  - Excel 報表

### Changed

- 2026-04-16 補強 Windows / UTF-8 終端輸出穩定性
- 2026-04-16 調整 import 與執行環境設定，改善 Windows 中文顯示體驗
