# -*- coding: utf-8 -*-
"""Create a minimal demo registration .xlsx via Excel COM.

Usage:
  python excel-com/scripts/create_form_template.py --out form.xlsx
"""
from __future__ import annotations

import argparse
from pathlib import Path

from excel_session import XL_OPEN_XML_WORKBOOK, attach_excel, find_open_workbook


def create_register_template(path: Path) -> None:
    excel = attach_excel(visible=True)
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)

    existing = find_open_workbook(excel, path)
    if existing is not None:
        existing.Close(SaveChanges=False)

    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "登记表"

    rows = [
        (1, "姓名", ""),
        (2, "学号", ""),
        (3, "学院", ""),
        (4, "专业", ""),
        (5, "联系电话", ""),
        (6, "邮箱", ""),
    ]
    for r, label, value in rows:
        ws.Cells(r, 1).Value2 = label
        ws.Cells(r, 2).Value2 = value if value else None

    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 28

    if path.exists():
        try:
            path.unlink()
        except OSError:
            pass
    wb.SaveAs(Filename=str(path), FileFormat=XL_OPEN_XML_WORKBOOK)
    print("wrote", path)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Create demo Excel form template")
    parser.add_argument("--out", default="form.xlsx")
    args = parser.parse_args(argv)
    create_register_template(Path(args.out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
