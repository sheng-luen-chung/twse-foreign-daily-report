# TWSE Foreign Daily Report

台股籌碼資料整理與視覺化工具。專案會抓取 TWSE 三大法人買賣超、TWSE 股價/成交量，以及 TDCC 集保股權分散資料，輸出 CSV、Excel 與深色系 HTML 視覺報表。

## 主要功能

- 產生每日外資買賣超排行榜。
- 依指定股票產生 TDCC 股權分散報表。
- 預設抓取最近 10 週股權分散資料，後續執行會持續累積歷史資料。
- 同時顯示多檔股票，例如 `2317 鴻海`、`4958 臻鼎-KY`。
- 產生 `持股比`、`金字塔`、`股東均張`、`股東人數` 四種視覺頁籤。
- 補入 TWSE 真實收盤價與成交量，不再只輸出 Excel。
- 依台灣市場慣例顯示顏色：上漲為紅色，下跌為綠色。

## 安裝

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 使用方式

只產生外資買賣超報表：

```powershell
python -X utf8 main.py --top 25 --outdir output
```

指定日期：

```powershell
python -X utf8 main.py --date 20260424 --top 25 --outdir output
```

產生股權分散與視覺報表。`--holders-targets` 後面可以接多檔台股股票代號或名稱：

```powershell
python -X utf8 main.py --holders-targets 2317 4958 --outdir output
```

### 多檔比較範例

比較兩檔，例如 `3189 景碩` 與 `4958 臻鼎-KY`：

```powershell
python -X utf8 main.py --holders-targets 3189 4958 --outdir output
```

比較三檔，例如 `3037 欣興`、`4958 臻鼎-KY`、`2368 金像電`：

```powershell
python -X utf8 main.py --holders-targets 3037 4958 2368 --outdir output
```

比較三檔並指定最近 10 週資料：

```powershell
python -X utf8 main.py --holders-targets 3037 4958 2368 --holders-weeks 10 --outdir output
```

### 常用 options

| Option | 說明 | 範例 |
| --- | --- | --- |
| `--holders-targets` | 要比較的股票標的，可接多檔代號或名稱。 | `--holders-targets 3037 4958 2368` |
| `--holders-weeks` | 股權分散表比較週數，預設 10，至少要 2。 | `--holders-weeks 10` |
| `--date` | 查詢基準日，未指定時使用台北時間今天，並自動往前找最近交易日。 | `--date 20260424` |
| `--outdir` | 輸出目錄，預設建議使用 `output`。 | `--outdir output` |
| `--top` | 外資買賣超排行榜顯示檔數。 | `--top 25` |

跑完多檔比較後，直接開啟 `output/latest_visual_dashboard.html` 查看互動視覺報表。桌機寬螢幕目前以兩欄排列，因此三檔會顯示成第一列兩檔、第二列一檔。

## 輸出檔案

主要輸出會放在 `output/`：

- `latest_foreign_report.csv`
- `latest_foreign_report.xlsx`
- `latest_shareholders_distribution.csv`
- `latest_shareholders_distribution.xlsx`
- `latest_visual_dashboard.html`
- `archive/YYYYMMDD/visual_dashboard_YYYYMMDD.html`

歷史資料會放在 `output/history/`：

- `shareholders_summary_history.csv`
- `shareholders_detail_history.csv`
- `price_history.csv`

`latest_visual_dashboard.html` 與其他 HTML 報表屬於產物，已加入 `.gitignore`，避免每次產生報表時污染版本控制。

## 視覺報表

HTML 視覺報表目前包含四個頁籤：

- `持股比`：顯示 1000 張以上大戶持股比例、50 張以下散戶持股比例、股價。
- `金字塔`：顯示各持股級距的股東人數、持股比例與週變化；點擊上方日期會切換到該週真實資料。
- `股東均張`：顯示每位股東平均持有張數、均張增減、週轉率與股價。
- `股東人數`：顯示總股東人數、人數增減、股價與成交量。

### 指標定義

`股東均張`：

```text
TDCC 集保總股數 / 總股東人數 / 1000
```

`週轉率`：

```text
該週區間 TWSE 成交股數合計 / 當週 TDCC 集保總股數 * 100
```

週區間採用上一筆 TDCC 週資料日期之後，到本週 TDCC 日期為止。若 TDCC 日期不是交易日，股價欄位會使用該日期以前最近一個 TWSE 交易日收盤資料。

## 資料來源

- TWSE 外資買賣超資料。
- TWSE `STOCK_DAY` 每日股價與成交量。
- TDCC 股權分散表。
- TWSE ISIN 基本資料，用於補中文股票名稱。

## 注意事項

- 報表資料仰賴公開網站回應，若 TWSE 或 TDCC 調整格式，可能需要同步更新 parser。
- `週轉率` 是本專案明確定義的週區間成交量比率，券商 App 可能使用不同分母、期間或內部資料，因此數字不一定完全一致。
- 建議使用 `python -X utf8` 執行，避免 Windows 終端機中文編碼問題。
