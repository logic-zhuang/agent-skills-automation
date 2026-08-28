# -*- coding: utf-8 -*-
"""Inspect Excel form sheets → structure report v1.

Usage:
  python excel-com/scripts/inspect_sheet.py path/to.xlsx
  python excel-com/scripts/inspect_sheet.py path/to.xlsx --json slots.json
  python excel-com/scripts/inspect_sheet.py path/to.xlsx --sheet Sheet1 --json slots.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from excel_session import attach_excel, open_workbook
from sheet_utils import inspect_workbook


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inspect Excel sheet form structure (v1)")
    parser.add_argument("path")
    parser.add_argument("--sheet", default=None, help="Sheet name or 1-based index; default=all")
    parser.add_argument("--json", metavar="OUT.json")
    parser.add_argument("--hidden", action="store_true")
    args = parser.parse_args(argv)

    path = Path(args.path)
    sheet: str | int | None = args.sheet
    if sheet is not None and str(sheet).isdigit():
        sheet = int(sheet)

    excel = attach_excel(visible=not args.hidden)
    wb = open_workbook(excel, path, read_only=True)
    payload = inspect_workbook(wb, sheet=sheet)

    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.json:
        Path(args.json).write_text(text, encoding="utf-8")
        print("wrote", args.json)
    else:
        print(text)

    print(
        f"sheets={len(payload['sheets'])} slots={len(payload['slots'])} "
        f"fingerprint={payload['form_fingerprint']}",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
