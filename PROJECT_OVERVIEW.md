# Project Overview

本專案用來整理台股公開籌碼資料，並輸出可讀性較高的 CSV、Excel 與 HTML 視覺報表。

## Architecture

主要進入點：

- `main.py`：CLI entry point。
- `twse_foreign_report/daily_report.py`：處理參數、TWSE 外資買賣超、TWSE 股價成交量、輸出流程。
- `twse_foreign_report/shareholder_distribution.py`：處理 TDCC 股權分散資料、股票名稱查詢、週資料與歷史檔。
- `twse_foreign_report/visual_report.py`：產生深色系 HTML 視覺報表。

## Data Flow

一般外資報表流程：

```text
main.py
  -> daily_report.run()
  -> TWSE 外資買賣超資料
  -> CSV / Excel
```

股權分散與視覺報表流程：

```text
main.py
  -> daily_report.run()
  -> shareholder_distribution.build_shareholder_distribution_report()
  -> TDCC 股權分散週資料
  -> update_price_history()
  -> TWSE STOCK_DAY 股價與成交量
  -> save_shareholder_outputs()
  -> save_visual_outputs()
  -> CSV / Excel / HTML / history CSV
```

## Output Structure

```text
output/
  latest_foreign_report.csv
  latest_foreign_report.xlsx
  latest_shareholders_distribution.csv
  latest_shareholders_distribution.xlsx
  latest_visual_dashboard.html
  archive/
    YYYYMMDD/
      visual_dashboard_YYYYMMDD.html
  history/
    shareholders_summary_history.csv
    shareholders_detail_history.csv
    price_history.csv
```

HTML 報表是產物，已在 `.gitignore` 排除。

## Visual Dashboard

視覺報表採用一檔股票一張卡片。每張卡片包含：

- 股票代號與中文名稱。
- TWSE 收盤價、漲跌與漲跌幅。
- TDCC 股權分散資料日期。
- `持股比`、`金字塔`、`股東均張`、`股東人數` 四個頁籤。

互動行為：

- 點擊頁籤會切換圖表與表格內容。
- `金字塔` 頁籤的日期按鈕會切換該週級距資料。
- 表格日期排序採最近日期在上，較舊日期往下。

## Indicator Definitions

大戶持股：

```text
1000 張以上持股級距的持股數 / TDCC 集保總股數 * 100
```

散戶持股：

```text
50 張以下持股級距的持股數 / TDCC 集保總股數 * 100
```

股東均張：

```text
TDCC 集保總股數 / 總股東人數 / 1000
```

週轉率：

```text
該週區間 TWSE 成交股數合計 / 當週 TDCC 集保總股數 * 100
```

股價：

```text
TWSE STOCK_DAY 收盤價。若 TDCC 週資料日期不是交易日，使用該日期以前最近一個交易日。
```

成交量：

```text
TWSE STOCK_DAY 成交股數，報表顯示為張數。
```

## Color Convention

台灣與中國市場慣例：

- 上漲：紅色。
- 下跌：綠色。
- 持平：中性色。

這套規則套用於首頁價格、漲跌幅、表格變動欄位，以及圖表中的漲跌相關標示。

## Operational Notes

- 建議用 `python -X utf8` 執行，降低 Windows 中文輸出亂碼風險。
- TDCC 與 TWSE 都是公開網站資料，若網站格式更動，parser 可能需要同步調整。
- `週轉率` 是專案內定義的週區間成交量比率，券商 App 可能採用不同資料源或計算方式，因此數字可能不同。
