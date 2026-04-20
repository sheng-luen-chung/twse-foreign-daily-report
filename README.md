# TWSE 外資買賣超日報工具 v4

這是一個可直接放進 **VS Code** 與 **GitHub** 維護的 Python 專案，用來抓取台灣證券交易所（TWSE）官方資料，產生外資買賣超日報、排行、摘要與 Excel 報表。

## 目前功能

- 抓取 **TWT38U**：外資及陸資買賣超彙總表
- 抓取 **MI_INDEX**：每日收盤行情
- 抓取 **TDCC 集保戶股權分散表**：比較各持股級距股東人數變化
- 自動合併成買超／賣超排行
- 支援 **自動回推最近交易日**
- 輸出 CSV、Excel
- 保留歷史輸出與 latest 固定檔名
- 已處理常見問題：
  - SSL 憑證相容性
  - TWSE CSV 多層／重複表頭
  - MI_INDEX JSON 結構差異

## 建議 repo 名稱

- `twse-foreign-daily-report`

## 專案結構

```text
twse-foreign-daily-report/
├─ .vscode/
│  ├─ launch.json
│  └─ settings.json
├─ data/
│  └─ .gitkeep
├─ output/
│  └─ .gitkeep
├─ twse_foreign_report/
│  ├─ __init__.py
│  └─ daily_report.py
├─ .gitignore
├─ main.py
├─ README.md
├─ requirements.txt
├─ run_latest.bat
├─ run_latest.sh
├─ run_with_date.bat
└─ run_with_date.sh
```

## 在 VS Code 中使用

### 1. 建立虛擬環境

在專案根目錄開 Terminal：

```bash
python -m venv .venv
```

Windows 啟用：

在 **Git Bash**：

```bash
source .venv/Scripts/activate
```

在 **PowerShell**：

```powershell
.\.venv\Scripts\Activate.ps1
```

在 **cmd**：

```cmd
.venv\Scripts\activate.bat
```

### 2. 安裝套件

```bash
pip install -r requirements.txt
```

### 3. 執行

**建議一律使用 UTF-8 模式執行**，避免不同 Windows 電腦的終端機編碼不一致，導致中文股票名稱亂碼。

#### 只看外資買賣超

自動找最近交易日：

```bash
python -X utf8 main.py --top 25 --outdir output
```

指定日期：

```bash
python -X utf8 main.py --date 20260414 --top 25 --outdir output
```

不自動回推前一個交易日：

```bash
python -X utf8 main.py --date 20260414 --top 25 --outdir output --no-auto-prev
```

說明：

- 不帶 `--holders-targets` 時，就是純外資買賣超模式
- 主要看 `latest_report.xlsx`、`latest_top_buy.csv`、`latest_top_sell.csv`

#### 看股東分散表

查集保戶股權分散表，可直接輸入股票名稱：

```bash
python -X utf8 main.py --date 20260419 --holders-targets 鴻海 臻鼎 --holders-weeks 3 --outdir output
```

也可直接輸入代號：

```bash
python -X utf8 main.py --holders-targets 2317 4958 --holders-weeks 3 --outdir output
```

說明：

- `--holders-targets` 會額外查 TDCC 股權分散表，並把 `holders_*` 工作表整合進 `latest_report.xlsx`
- `--holders-weeks` 預設為 `3`，會抓查詢日以前最近 3 週可用資料
- `--date` 在這個模式下代表「往前找到這一天以前最近可用的週資料」
- 另外也會保留獨立的 `latest_shareholders_report.xlsx` 與對應 CSV
- 其中 `recent_changes` / `latest_shareholders_recent_changes.csv` 會直接列出最近兩週各持股級距的人數增減與張數增減
- 如果只想看外資買賣超，不要加 `--holders-targets`

`recent_changes` 主要欄位：

- `holders_YYYYMMDD`：該週該持股級距的股東人數
- `holders_change_新週_vs_舊週`：最近兩週股東人數增減
- `shares_YYYYMMDD`：該週該持股級距的總股數
- `shares_change_新週_vs_舊週`：最近兩週總股數增減
- `shares_lots_YYYYMMDD`：把總股數換算成張數
- `lots_change_新週_vs_舊週`：最近兩週總股數增減張數

範例解讀：

- 以 `2317 鴻海`、持股分級 `1-999` 為例：
- `holders_20260417 = 432714`、`holders_20260410 = 437843`，表示這個持股級距的股東人數從前一週的 `437,843` 人，變成最新一週的 `432,714` 人
- `holders_change_20260417_vs_20260410 = -5129`，表示最近兩週這個級距少了 `5,129` 位股東
- `shares_20260417 = 104425588`、`shares_20260410 = 105578153`，表示這個級距合計持有股數從 `105,578,153` 股降到 `104,425,588` 股
- `lots_change_20260417_vs_20260410 = -1152.565`，表示換算後最近兩週少了 `1,152.565` 張

## Git Bash 中文顯示建議

若在某台 Windows 電腦的 **Git Bash** 看到中文亂碼，先執行：

```bash
export LANG=zh_TW.UTF-8
export LC_ALL=zh_TW.UTF-8
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
```

若想永久生效，可加入 `~/.bashrc`：

```bash
export LANG=zh_TW.UTF-8
export LC_ALL=zh_TW.UTF-8
export PYTHONUTF8=1
export PYTHONIOENCODING=utf-8
```

加入後執行：

```bash
source ~/.bashrc
```

然後再測試：

```bash
python -X utf8 -c "print('台積電 鴻海 聯發科 富邦金 華航')"
```

## 輸出結果

程式會在 `output/` 之下產生：

- `latest_full.csv`
- `latest_top_buy.csv`
- `latest_top_sell.csv`
- `latest_summary.csv`
- `latest_report.xlsx`
- `latest_shareholders_detail.csv`
- `latest_shareholders_recent_changes.csv`
- `latest_shareholders_changes.csv`
- `latest_shareholders_summary.csv`
- `latest_shareholders_report.xlsx`
- `archive/YYYYMMDD/` 每日歷史檔
- `history/summary_history.csv`

## VS Code 執行方式

你可以直接用：

- `Run and Debug` → `TWSE 日報：自動最近交易日`
- `Run and Debug` → `TWSE 日報：指定日期`

## 建議的 Git 初始化流程

```bash
git init
git branch -M main
git add .
git commit -m "Initial commit: TWSE foreign daily report v4"
```

如果你已經先在 GitHub 建好空 repo：

```bash
git remote add origin git@github.com:YOUR_ACCOUNT/twse-foreign-daily-report.git
git push -u origin main
```

## 下一步建議

這個版本已經能當成小型資料產品的骨架。真正值得再升級的方向有：

1. 加入 **上市 + 上櫃** 合併排行
2. 加入 **外資連買／連賣天數**
3. 加入 **借券賣出餘額** 交叉分析
4. 加入 **排程**（Windows 工作排程器 / GitHub Actions）
5. 加入 **多日趨勢圖** 與簡易 dashboard

## 你可以怎麼擴充

若你之後要發展成研究或教學 demo，可以把功能分成：

- `fetch`：資料抓取
- `parse`：格式解析
- `transform`：統計整理
- `export`：CSV / Excel / 圖表輸出
- `schedule`：每日自動執行

這樣後續維護會比把所有邏輯擠在單一腳本中容易很多。
