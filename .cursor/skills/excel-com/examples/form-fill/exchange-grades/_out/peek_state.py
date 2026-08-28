# -*- coding: utf-8 -*-
import sys
from pathlib import Path

sys.path.insert(0, r"d:\Desktop\desktop-rip\excel-com\scripts")
from excel_session import attach_excel, find_open_workbook, open_workbook
from sheet_utils import cell_text

orig = Path(r"D:\Downloads\交流生成绩录入模板 (1) (1).xlsx")
draft = Path(
    r"d:\Desktop\desktop-rip\excel-com\examples\form-fill\exchange-grades\_out\draft_add_course.xlsx"
)

excel = attach_excel(visible=True)


def peek(path: Path, label: str) -> None:
    wb = find_open_workbook(excel, path)
    already = wb is not None
    if wb is None:
        wb = open_workbook(excel, path, read_only=True)
    ws = wb.Worksheets(1)
    used = ws.UsedRange
    last = int(used.Row) + int(used.Rows.Count) - 1
    print(f"--- {label} ---")
    print("path:", path)
    print("size:", path.stat().st_size if path.exists() else None)
    print("already_open:", already)
    print("FullName:", wb.FullName)
    print("Saved:", getattr(wb, "Saved", "?"))
    print("last_row:", last)
    for r in range(max(3, last - 1), last + 1):
        f = cell_text(ws.Range(f"F{r}").Value2)
        t = cell_text(ws.Range(f"T{r}").Value2)
        print(f"  R{r}: F={f!r} T={t!r}")
    if not already:
        wb.Close(SaveChanges=False)


peek(orig, "ORIGINAL")
peek(draft, "DRAFT")
