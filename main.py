#!/usr/bin/env python3
import sys
from twse_foreign_report.daily_report import main


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
