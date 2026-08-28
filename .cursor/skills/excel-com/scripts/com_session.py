# -*- coding: utf-8 -*-
"""Open an Excel workbook and keep the window for interactive review.

Usage:
  python excel-com/scripts/com_session.py path/to.xlsx
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from excel_session import attach_excel, open_workbook


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Open Excel workbook and keep visible")
    parser.add_argument("path")
    args = parser.parse_args(argv)

    excel = attach_excel(visible=True)
    wb = open_workbook(excel, Path(args.path))
    print(f"opened: {wb.FullName}")
    print(f"sheets: {wb.Worksheets.Count}")
    print("Excel left running (not Quit).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
