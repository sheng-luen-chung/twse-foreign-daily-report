# TWSE 外資買賣超日報工具 v4

這是一個可直接放進 **VS Code** 與 **GitHub** 維護的 Python 專案，用來抓取台灣證券交易所（TWSE）官方資料，產生外資買賣超日報、排行、摘要與 Excel 報表。

## 目前功能

- 抓取 **TWT38U**：外資及陸資買賣超彙總表
- 抓取 **MI_INDEX**：每日收盤行情
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
