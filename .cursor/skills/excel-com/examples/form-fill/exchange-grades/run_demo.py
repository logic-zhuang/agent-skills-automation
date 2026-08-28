# -*- coding: utf-8 -*-
"""Demo: table_rows grade sheet — Agent fill via append_row (never edit source in place).

Source (optional): Downloads 交流生成绩录入模板.
Falls back to existing _out/form.xlsx if source missing.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parents[2]
SCRIPTS = SKILL / "scripts"
OUT = HERE / "_out"
SRC_CANDIDATES = [
    Path(r"D:\Downloads\交流生成绩录入模板 (1) (1).xlsx"),
    Path(r"D:\Downloads\交流生成绩录入模板.xlsx"),
]
FILL = HERE / "fill_add_course.json"


def run(args: list[str]) -> None:
    print("+", " ".join(args))
    subprocess.check_call(args, cwd=str(SKILL))


def main() -> int:
    OUT.mkdir(exist_ok=True)
    form = OUT / "form.xlsx"
    draft = OUT / "draft_add_course.xlsx"
    inspect_json = OUT / "inspect.json"

    # Avoid Win32 lock races from prior demos
    sys.path.insert(0, str(SCRIPTS))
    from excel_session import attach_excel, find_open_workbook

    excel = attach_excel(visible=True)
    for p in (form, draft):
        wb = find_open_workbook(excel, p)
        if wb is not None:
            wb.Close(SaveChanges=False)

    src = next((p for p in SRC_CANDIDATES if p.exists()), None)
    if src is not None:
        src_open = find_open_workbook(excel, src)
        if src_open is not None:
            src_open.Close(SaveChanges=False)
        shutil.copy2(src, form)
        print("cloned source →", form)
    elif not form.exists():
        print("ERROR: missing source template and", form)
        return 1
    else:
        print("reuse", form)

    py = sys.executable
    run([py, str(SCRIPTS / "inspect_sheet.py"), str(form), "--json", str(inspect_json)])
    data = json.loads(inspect_json.read_text(encoding="utf-8"))
    hint = (data.get("facts") or {}).get("layout_hint")
    print("layout_hint:", hint)
    if hint != "table_rows":
        print("WARN: expected table_rows; Agent should still prefer append_row if header+data rows")

    run(
        [
            py,
            str(SCRIPTS / "auto_fill_sheet.py"),
            str(form),
            "--fill",
            str(FILL),
            "--draft",
            str(draft),
            "--highlight",
        ]
    )
    print("draft:", draft)
    print("source untouched:", src or form)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
