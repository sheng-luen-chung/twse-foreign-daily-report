#!/usr/bin/env python3
import os
import sys

# Keep UTF-8 output stable on Windows shells.
os.environ.setdefault("PYTHONUTF8", "1")
os.environ.setdefault("PYTHONIOENCODING", "utf-8")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from twse_foreign_report.daily_report import main

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
