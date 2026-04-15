#!/usr/bin/env bash
read -p "請輸入日期（YYYYMMDD）: " INPUT_DATE
python main.py --date "$INPUT_DATE" --top 25 --outdir output
