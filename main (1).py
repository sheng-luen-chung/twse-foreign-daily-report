#!/usr/bin/env python3
import os
import sys

# 盡量固定本程式的輸出為 UTF-8，降低 Windows / Git Bash 中文亂碼機率。
# 注意：這不能完全取代終端機本身的編碼設定，所以 README 仍建議用：
#   python -X utf8 main.py ...
os.environ.setdefault("PYTHONUTF8", "1")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from twse_foreign_report.daily_report import main

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
