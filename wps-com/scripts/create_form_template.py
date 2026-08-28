# -*- coding: utf-8 -*-
"""Create example counselor form template for form-fill demo (WPS)."""
from __future__ import annotations

from pathlib import Path

wdCharacter = 1

OUT = Path(__file__).resolve().parents[1] / "templates" / "form" / "导员登记表.docx"


def attach_app():
    from wps_session import attach_wps as _attach

    return _attach(visible=False)


def set_cell(t, r, c, text):
    cell = t.Cell(r, c)
    rng = cell.Range
    rng.MoveEnd(wdCharacter, -1)
    rng.Text = text


def main() -> int:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    app = attach_app()
    doc = app.Documents.Add()
    try:
        p = doc.Paragraphs(1)
        p.Range.Text = "导员登记表\r"
        p.Range.Font.NameFarEast = "黑体"
        p.Range.Font.Size = 16
        p.Format.Alignment = 1

        rows = (
            ("姓名", "（请填写）"),
            ("学号", "（请填写）"),
            ("学院", "（请填写）"),
            ("联系电话", "（请填写）"),
        )
        t = doc.Tables.Add(doc.Paragraphs(2).Range, len(rows), 2)
        for i, (label, val) in enumerate(rows, start=1):
            set_cell(t, i, 1, label)
            set_cell(t, i, 2, val)

        doc.SaveAs2(str(OUT.resolve()))
        print("created", OUT)
    finally:
        doc.Close(False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
