@echo off
set /p INPUT_DATE=請輸入日期（YYYYMMDD）: 
python main.py --date %INPUT_DATE% --top 25 --outdir output
pause
