# PROJECT OVERVIEW

## 1. 專案定位

這個專案是一個以 Python 撰寫的台股資料整理工具，目標是把原始官方資料整理成比較容易閱讀與後續分析的 CSV / Excel 報表。

目前專案主要有兩條功能線：

1. 外資買賣超日報
   來源是 TWSE 的外資及陸資買賣超資料與每日收盤行情。
   用來產生外資買超排行、賣超排行、完整明細、摘要與歷史摘要。

2. 股東分散表分析
   來源是 TDCC 的集保戶股權分散表查詢頁。
   用來比較最近幾週不同持股級距的股東人數變化，以及各級距總股數 / 張數變化。


## 2. 專案功能總覽

### 2.1 外資買賣超功能

- 抓取 TWSE `TWT38U` 外資及陸資買賣超資料
- 抓取 TWSE `MI_INDEX` 收盤行情資料
- 合併出每檔股票的：
  - 買進股數
  - 賣出股數
  - 買賣超股數
  - 收盤價
  - 漲跌價差
  - 漲跌幅
  - 買賣超金額估算
- 產出：
  - 買超前 N 名
  - 賣超前 N 名
  - 完整明細
  - 摘要
  - 歷史摘要
  - Excel 圖表報表

### 2.2 股東分散表功能

- 以股票代號或股票名稱查詢 TDCC 集保戶股權分散表
- 可抓查詢日以前最近 N 週的資料，預設 3 週
- 比較各持股級距的：
  - 股東人數
  - 股數
  - 張數
  - 比例
  - 最近兩週增減
  - 多週逐週增減
- 產出：
  - 明細表
  - 最近兩週重點變化表
  - 多週變化表
  - 摘要表
  - 獨立 Excel
  - 整合到主報表 `latest_report.xlsx` 的 `holders_*` 工作表


## 3. 專案結構與檔案角色

### 3.1 主要檔案

| 檔案 | 角色 |
|---|---|
| [main.py](./main.py) | 專案執行入口，負責設定 UTF-8 環境，最後呼叫 `twse_foreign_report.daily_report.main()` |
| [twse_foreign_report/__init__.py](./twse_foreign_report/__init__.py) | 套件入口，對外暴露 `main` |
| [twse_foreign_report/daily_report.py](./twse_foreign_report/daily_report.py) | 主流程控制器。外資買賣超的抓取、整理、輸出都在這裡，股東分散表模式也由這裡分流啟動 |
| [twse_foreign_report/shareholder_distribution.py](./twse_foreign_report/shareholder_distribution.py) | 股東分散表專用模組，負責股票代號解析、TDCC 查詢、資料整理、CSV/Excel 輸出 |
| [README.md](./README.md) | 使用說明與執行方式 |
| [PROJECT_OVERVIEW.md](./PROJECT_OVERVIEW.md) | 專案總覽與交接說明 |
| [requirements.txt](./requirements.txt) | 依賴套件清單 |
| [run_latest.bat](./run_latest.bat) / [run_latest.sh](./run_latest.sh) | 快速執行最近交易日報表 |
| [run_with_date.bat](./run_with_date.bat) / [run_with_date.sh](./run_with_date.sh) | 指定日期執行 |

### 3.2 輸出資料夾

| 路徑 | 角色 |
|---|---|
| [output/](./output/) | 所有輸出檔的根目錄 |
| `output/archive/YYYYMMDD/` | 每次依實際資料日期建立的歷史存檔資料夾 |
| `output/history/summary_history.csv` | 外資摘要歷史累積檔 |


## 4. 模組關係

### 4.1 呼叫關係

整體呼叫鏈大致如下：

```text
main.py
  -> twse_foreign_report.daily_report.main()
     -> parse_args()
     -> normalize_date()
     -> 分成兩種模式：
        A. 純外資買賣超模式
           -> resolve_build()
           -> build_rank_table()
           -> save_outputs()
        B. 外資買賣超 + 股東分散表模式
           -> resolve_build()
           -> build_shareholder_distribution_report()
           -> save_shareholder_outputs()
           -> save_outputs(..., shareholder_result=...)
```

### 4.2 `daily_report.py` 與 `shareholder_distribution.py` 的分工

`daily_report.py` 是主控台。

- 負責 CLI 參數解析
- 決定目前是純外資模式，還是包含股東分散表模式
- 建立外資報表
- 輸出主 Excel 與主 CSV
- 若有股東分散表資料，將 `holders_*` 工作表整合進主 Excel

`shareholder_distribution.py` 是股東分散表子系統。

- 將輸入的股票名稱或代號解析成正式股票代號
- 從 TDCC 查詢頁取得最近幾週資料
- 整理成 detail / recent_changes / changes / summary
- 產出股東分散表專用 CSV 與 Excel
- 提供工作表寫入函式，供主報表整合使用


## 5. 資料來源

### 5.1 外資買賣超資料來源

- TWSE `TWT38U`
  外資及陸資買賣超彙總表

- TWSE `MI_INDEX`
  每日收盤行情

程式會同時支援多個網址版本與 JSON / CSV fallback，以增加穩定性。

### 5.2 股東分散表資料來源

- TDCC 集保戶股權分散表查詢頁
  用來逐週查詢各股票的持股級距分布

- TWSE ISIN 查詢頁
  用來把使用者輸入的股票名稱解析成正式證券代號


## 6. 運作流程

### 6.1 純外資買賣超模式流程

當使用者沒有帶 `--holders-targets` 時，流程如下：

1. `parse_args()` 解析 CLI 參數
2. `normalize_date()` 將日期標準化成 `YYYYMMDD`
3. `resolve_build()` 依查詢日往前尋找最近可用交易日
4. `build_rank_table()`：
   - 抓 TWSE 外資買賣超資料
   - 抓 TWSE 收盤行情資料
   - 兩者 merge
   - 計算外資買超 / 賣超 / 淨張數 / 淨值估算
   - 建立 top buy / top sell / summary
5. `save_outputs()`：
   - 輸出 CSV
   - 輸出 Excel
   - 更新 `summary_history.csv`
   - 複製一份到 `latest_*`

### 6.2 股東分散表模式流程

當使用者帶 `--holders-targets` 時，流程如下：

1. 先照正常流程建立外資買賣超報表
2. `build_shareholder_distribution_report()`：
   - `resolve_targets()` 將股票名稱轉成正式代號
   - `_fetch_filtered_shareholder_rows()` 逐股票、逐週向 TDCC 查詢
   - `_complete_detail_grid()` 補齊每檔股票每週各持股級距資料
   - `_build_recent_change_table()` 建立最近兩週重點變化
   - `_build_change_table()` 建立多週完整變化
   - `_build_summary_table()` 建立整體摘要
3. `save_shareholder_outputs()` 產出股東分散表專用 CSV / Excel
4. `save_outputs(..., shareholder_result=...)` 將 `holders_*` 工作表整合進主 Excel


## 7. 重要資料結構

### 7.1 `BuildResult`

定義在 [daily_report.py](./twse_foreign_report/daily_report.py)。

用來承接外資買賣超模式的主要結果：

- `date`
- `merged`
- `buys`
- `sells`
- `summary`

### 7.2 `ShareholderBuildResult`

定義在 [shareholder_distribution.py](./twse_foreign_report/shareholder_distribution.py)。

用來承接股東分散表模式的主要結果：

- `base_date`
- `selected_dates`
- `targets`
- `detail`
- `changes`
- `recent_changes`
- `summary`


## 8. 產生檔案總覽

### 8.1 外資買賣超相關檔案

| 檔名 | 說明 |
|---|---|
| `latest_full.csv` | 外資買賣超完整明細 |
| `latest_top_buy.csv` | 外資買超前 N 名 |
| `latest_top_sell.csv` | 外資賣超前 N 名 |
| `latest_summary.csv` | 外資日摘要 |
| `latest_report.xlsx` | 主報表 Excel |
| `output/archive/YYYYMMDD/twse_foreign_rank_YYYYMMDD.xlsx` | 純外資歷史 Excel |
| `output/archive/YYYYMMDD/twse_foreign_rank_with_holders_YYYYMMDD.xlsx` | 含股東分散表工作表的整合歷史 Excel |
| `output/history/summary_history.csv` | 外資摘要歷史累積檔 |

### 8.2 股東分散表相關檔案

| 檔名 | 說明 |
|---|---|
| `latest_shareholders_detail.csv` | 各股票、各週、各持股級距的完整明細 |
| `latest_shareholders_recent_changes.csv` | 最近兩週的重點變化，最適合快速閱讀 |
| `latest_shareholders_changes.csv` | 多週完整變化表 |
| `latest_shareholders_summary.csv` | 每檔股票的總體摘要 |
| `latest_shareholders_report.xlsx` | 股東分散表專用 Excel |


## 9. 各輸出檔如何解讀

### 9.1 `latest_summary.csv`

這是外資日摘要，通常用來快速看：

- 當天統計到幾檔股票
- 外資總買進 / 賣出 / 淨買賣超張數
- 買超檔數與賣超檔數
- 買超第一名與賣超第一名
- 估算買超與賣超總金額

適合拿來做 daily briefing 或長期歷史追蹤。

### 9.2 `latest_full.csv`

這是外資買賣超完整明細，每列是一檔股票，常見欄位包括：

- `code`：股票代號
- `name`：股票名稱
- `foreign_buy_lots`：外資買進張數
- `foreign_sell_lots`：外資賣出張數
- `foreign_net_lots`：外資買賣超張數
- `close`：收盤價
- `change`：漲跌價差
- `pct`：漲跌幅
- `foreign_net_value_est`：買賣超金額估算

### 9.3 `latest_top_buy.csv` / `latest_top_sell.csv`

這兩份是外資排行表。

- `latest_top_buy.csv`：外資買超前 N 名
- `latest_top_sell.csv`：外資賣超前 N 名

通常是最常看的輸出。

### 9.4 `latest_report.xlsx`

這是主報表 Excel。

純外資模式下通常包含：

- `dashboard`
- `summary`
- `top_buy`
- `top_sell`
- `full`
- `summary_history`

若有帶 `--holders-targets`，還會額外出現：

- `holders_overview`
- `holders_summary`
- `holders_recent_changes`
- `holders_changes`
- `holders_detail`

### 9.5 `latest_shareholders_detail.csv`

這是股東分散表完整明細。

每列代表：

- 一檔股票
- 某一週
- 某一個持股級距

常用欄位：

- `holding_range`：持股級距，例如 `1-999`、`1,000-5,000`
- `holders`：該級距股東人數
- `shares`：該級距總股數
- `ratio_pct`：占集保庫存數比例

### 9.6 `latest_shareholders_recent_changes.csv`

這是股東分散表最推薦閱讀的檔案，重點是最近兩週的增減。

常見欄位：

- `holders_YYYYMMDD`
  某週該級距股東人數

- `holders_change_新週_vs_舊週`
  最近兩週股東人數增減

- `shares_YYYYMMDD`
  某週該級距總股數

- `shares_change_新週_vs_舊週`
  最近兩週總股數增減

- `shares_lots_YYYYMMDD`
  將總股數換算為張數

- `lots_change_新週_vs_舊週`
  最近兩週總股數增減張數

解讀範例：

- 如果 `holders_change_20260417_vs_20260410 = -5129`
  代表這個持股級距最近一週比前一週少了 5,129 位股東

- 如果 `lots_change_20260417_vs_20260410 = -1152.565`
  代表這個持股級距合計持股量比前一週少了 1,152.565 張

### 9.7 `latest_shareholders_changes.csv`

這份是多週版本的完整變化表，會保留最近 N 週所有欄位與逐週增減欄位。

適合：

- 後續自己做分析
- 匯入 Excel 篩選
- 做簡單量化觀察

### 9.8 `latest_shareholders_summary.csv`

這份是每檔股票整體摘要，著重在總股東數、總股數、總張數的週變化。


## 10. 常用執行方式

### 10.1 只看外資買賣超

```bash
python -X utf8 main.py --top 25 --outdir output
```

### 10.2 指定日期的外資買賣超

```bash
python -X utf8 main.py --date 20260417 --top 25 --outdir output
```

### 10.3 查股東分散表

```bash
python -X utf8 main.py --date 20260419 --holders-targets 鴻海 臻鼎 --holders-weeks 3 --outdir output
```


## 11. 錯誤處理與穩定性設計

專案目前已處理幾個實務上常見的問題：

- TWSE 可能沒有當日資料
  會自動往前回推最近交易日

- TWSE JSON / CSV 結構可能不一致
  外資與收盤行情模組都有 fallback 設計

- SSL 憑證相容性問題
  會在必要時啟用 SSL fallback

- 股票名稱不一定能直接查到唯一代號
  會先走 ISIN 查詢做名稱解析，必要時提醒使用者改用代號


## 12. 維護與擴充建議

如果之後要再擴充，建議延續目前分層方式：

1. `daily_report.py` 保持為主流程與外資功能主控
2. `shareholder_distribution.py` 保持為股東分散表子系統
3. 若新增借券、融資融券、法人合併分析，可再切成獨立模組
4. Excel 匯出與 chart 邏輯若再變大，可獨立抽成 `exporters.py`
5. 若要做排程，可在現有 CLI 基礎上直接加 Windows 工作排程器或 GitHub Actions


## 13. 一句話總結

這個專案的核心價值，是把 TWSE 與 TDCC 的原始官方資料整理成可以直接閱讀、追蹤與比較的報表，同時保留固定檔名輸出與歷史存檔，方便每天持續使用。
