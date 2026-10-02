#!/usr/bin/env python3
"""
make_month.py — Pre-create a monthly workbook reports/51la_YYYY-MM.xlsx.

It rolls over the latest existing monthly workbook, preserving its layout,
styles and formulas, and fills in every day of the target month.

Usage:
  python3 make_month.py             # create next month's workbook
  python3 make_month.py 2026-11     # create the November 2026 workbook
  python3 make_month.py --list      # list existing monthly workbooks
"""

import argparse
import sys
from datetime import date
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

from config import load
from workbook_manager import ensure_month_workbook, list_monthly_workbooks


def _next_month(today: date) -> date:
    if today.month == 12:
        return date(today.year + 1, 1, 1)
    return date(today.year, today.month + 1, 1)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("month", nargs="?", help="Target month as YYYY-MM (default: next month)")
    parser.add_argument("--list", action="store_true", help="List existing monthly workbooks and exit")
    args = parser.parse_args()

    cfg = load()

    if args.list:
        items = list_monthly_workbooks(cfg.reports_dir, cfg.report_file_pattern)
        if not items:
            print(f"No monthly workbooks found in {cfg.reports_dir}/")
            return 0
        for key, path in items:
            print(f"{key}  {path}")
        return 0

    if args.month:
        try:
            year, month = (int(x) for x in args.month.split("-"))
            target = date(year, month, 1)
        except ValueError:
            print(f"[ERROR] Invalid month '{args.month}', expected YYYY-MM")
            return 1
    else:
        target = _next_month(date.today())

    try:
        path = ensure_month_workbook(target, cfg.reports_dir, cfg.report_file_pattern)
    except Exception as e:
        print(f"[ERROR] {e}")
        return 1

    print(f"[OK] Month workbook ready: {path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
