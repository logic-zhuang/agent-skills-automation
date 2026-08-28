# -*- coding: utf-8 -*-
"""Write fill.json into an Excel workbook via COM.

Primitives:
  - cells[] / ops write_cell — set Range.Value2
  - ops append_row — append a data row (table_rows layouts)

Usage:
  python excel-com/scripts/fill_sheet.py path/to.xlsx --json fill.json --save
  python excel-com/scripts/fill_sheet.py source.xlsx --json fill.json --draft draft.xlsx
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from excel_session import attach_excel, clone_for_edit, open_workbook
from sheet_utils import a1_address, verify_fill

HIGHLIGHT_COLOR = 0xCCFFCC


def _write_one_cell(ws: Any, address: str, text: Any, *, highlight: bool) -> bool:
    rng = ws.Range(address)
    try:
        if rng.MergeCells:
            rng = rng.MergeArea.Cells(1, 1)
    except Exception:
        pass
    try:
        if bool(rng.HasFormula):
            print(f"skip formula cell {ws.Name}!{address}", file=sys.stderr)
            return False
    except Exception:
        pass
    rng.Value2 = text
    if highlight:
        try:
            rng.Interior.Color = HIGHLIGHT_COLOR
        except Exception:
            pass
    return True


def append_row(wb: Any, op: dict[str, Any], *, highlight: bool = False) -> int:
    """Append one row. values: {\"B\": \"...\", \"F\": \"...\"} or list of {col|address, text}."""
    sheet = op.get("sheet")
    ws = wb.Worksheets(sheet) if sheet else wb.Worksheets(1)
    used = ws.UsedRange
    last = int(used.Row) + int(used.Rows.Count) - 1
    new_row = int(op.get("row") or (last + 1))

    if op.get("copy_format_from_last", True) and last >= 1:
        try:
            ws.Rows(last).Copy()
            ws.Rows(new_row).PasteSpecial(Paste=-4122)  # xlPasteFormats
            wb.Application.CutCopyMode = False
        except Exception as exc:
            print(f"format copy skipped: {exc}", file=sys.stderr)

    n = 0
    values = op.get("values") or {}
    if isinstance(values, dict):
        for col_letter, val in values.items():
            addr = f"{col_letter}{new_row}" if col_letter.isalpha() else col_letter
            if isinstance(col_letter, str) and col_letter[0].isdigit():
                addr = a1_address(new_row, int(col_letter))
            elif isinstance(col_letter, str) and col_letter.isalpha():
                addr = f"{col_letter.upper()}{new_row}"
            if _write_one_cell(ws, addr, val, highlight=highlight):
                n += 1
    elif isinstance(values, list):
        for item in values:
            if "address" in item:
                addr = item["address"]
            elif "col" in item:
                addr = a1_address(new_row, int(item["col"]))
            else:
                continue
            if _write_one_cell(ws, addr, item.get("text", item.get("value", "")), highlight=highlight):
                n += 1
    op["_written_row"] = new_row
    return n


def fill_sheet(wb: Any, fill_payload: dict[str, Any], *, highlight: bool = False) -> int:
    n = 0
    for item in fill_payload.get("cells") or []:
        sheet = item.get("sheet")
        address = item.get("address")
        text = item.get("text", item.get("value", ""))
        if not address:
            continue
        ws = wb.Worksheets(sheet) if sheet else wb.Worksheets(1)
        if _write_one_cell(ws, address, text, highlight=highlight):
            n += 1

    for op in fill_payload.get("ops") or []:
        kind = op.get("op") or op.get("type")
        if kind == "append_row":
            n += append_row(wb, op, highlight=highlight)
        elif kind == "write_cell":
            sheet = op.get("sheet")
            ws = wb.Worksheets(sheet) if sheet else wb.Worksheets(1)
            if _write_one_cell(ws, op["address"], op.get("text", op.get("value", "")), highlight=highlight):
                n += 1
        else:
            print(f"skip unknown op: {kind}", file=sys.stderr)
    return n


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Fill Excel from Agent fill.json (cells + ops)")
    parser.add_argument("path", help="Working copy, or source when using --draft")
    parser.add_argument("--json", required=True, help="fill.json or plan.json (uses .fill)")
    parser.add_argument("--hidden", action="store_true")
    parser.add_argument("--save", action="store_true")
    parser.add_argument("--draft", metavar="OUT.xlsx", help="Clone path→OUT, fill OUT only")
    parser.add_argument("--highlight", action="store_true")
    parser.add_argument("--no-verify", action="store_true")
    args = parser.parse_args(argv)

    raw = json.loads(Path(args.json).read_text(encoding="utf-8"))
    fill_payload = raw.get("fill") if isinstance(raw.get("fill"), dict) else raw
    if "cells" not in fill_payload and "ops" not in fill_payload and "fill" in raw:
        fill_payload = raw["fill"]

    excel = attach_excel(visible=not args.hidden)
    src = Path(args.path)
    if args.draft:
        work = Path(args.draft).resolve()
        clone_for_edit(excel, src, work)
        wb = open_workbook(excel, work)
    else:
        work = src.resolve()
        wb = open_workbook(excel, work)

    filled = fill_sheet(wb, fill_payload, highlight=args.highlight)

    ver = None
    if not args.no_verify and fill_payload.get("cells"):
        ver = verify_fill(wb, fill_payload["cells"])

    if args.draft or args.save:
        wb.Save()
        print("saved", work, file=sys.stderr)
    else:
        print("not saved (pass --draft or --save)", file=sys.stderr)

    print(json.dumps({"filled": filled, "working_copy": str(work), "verification": ver}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
