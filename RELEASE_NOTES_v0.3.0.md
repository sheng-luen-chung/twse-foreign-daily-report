# Release Notes v0.3.0

Release date: 2026-04-25

這版把原本偏 Excel/CSV 的輸出，升級成可以直接檢視籌碼趨勢的 HTML 視覺報表，並補上真實股價、成交量與多週歷史串接。

## Highlights

- 新增 `output/latest_visual_dashboard.html`，可直接用瀏覽器查看最新籌碼視覺報表。
- 支援同時顯示多檔股票，目前常用範例為 `2317 鴻海` 與 `4958 臻鼎-KY`。
- `--holders-weeks` 預設由 3 週改為 10 週。
- 新增歷史串接，後續執行會累積 TDCC 股權分散明細與 TWSE 股價資料。
- 新增互動頁籤：`持股比`、`金字塔`、`股東均張`、`股東人數`。
- `金字塔` 上方日期按鈕可切換不同週期，顯示對應週別的級距人數、持股比例與變動。
- `股東人數` 與相關表格補入股價與成交量。
- 改用台灣市場顏色慣例：漲為紅色，跌為綠色。

## New Data Outputs

新增或強化下列輸出：

- `output/latest_visual_dashboard.html`
- `output/archive/YYYYMMDD/visual_dashboard_YYYYMMDD.html`
- `output/history/shareholders_detail_history.csv`
- `output/history/price_history.csv`

HTML 產物已加入 `.gitignore`，避免自動產生的報表進入 commit。

## Data Sources

- TDCC 股權分散表：股東人數、級距人數、級距持股數。
- TWSE `STOCK_DAY`：每日成交股數、收盤價、漲跌價差。
- TWSE ISIN：股票中文名稱。

## Indicator Definitions

`股東均張`：

```text
TDCC 集保總股數 / 總股東人數 / 1000
```

`週轉率`：

```text
該週區間 TWSE 成交股數合計 / 當週 TDCC 集保總股數 * 100
```

週區間從上一筆 TDCC 日期之後開始，包含本週 TDCC 日期。若 TDCC 日期遇到非交易日，收盤價會採用最近一個可用 TWSE 交易日。

## Compatibility Notes

- 舊的 CSV 與 Excel 輸出仍保留。
- `--holders-weeks` 仍可手動指定，未指定時使用 10 週。
- 舊版 `RELEASE_NOTES_v0.2.0.md` 檔名暫時保留，但內容已更新為本次 v0.3.0 發行說明，方便既有流程引用。

## Verification

本版已用下列方式驗證：

```powershell
python -m py_compile main.py twse_foreign_report\daily_report.py twse_foreign_report\shareholder_distribution.py twse_foreign_report\visual_report.py
python main.py --help
python -X utf8 main.py --holders-targets 2317 4958 --holders-weeks 10 --outdir output
```

實際資料曾成功抓取 2026-02-13 到 2026-04-24 的 10 週 TDCC 週資料，並產生 2317 與 4958 的視覺報表。
