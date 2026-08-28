# -*- coding: utf-8 -*-
"""Create hybrid application template (table + 3 sections) via WPS."""
from __future__ import annotations

from pathlib import Path

wdCharacter = 1

OUT = Path(__file__).resolve().parents[1] / "templates" / "hybrid" / "申请书模板.docx"


def attach_app():
    from wps_session import attach_wps as _attach

    return _attach(visible=False)


def add_para(doc, text, style="正文", bold=False):
    p = doc.Paragraphs.Add()
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    rng.Text = text
    try:
        p.Style = style
    except Exception:
        pass
    if bold:
        rng.Font.Bold = True
    return p


def set_cell(t, r, c, text):
    rng = t.Cell(r, c).Range
    rng.MoveEnd(wdCharacter, -1)
    rng.Text = text


def main() -> int:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    app = attach_app()
    doc = app.Documents.Add()
    try:
        add_para(doc, "申请表", "标题 1")

        rows = (("姓名", "（请填写）"), ("学院", "（请填写）"), ("申请项目", "（请填写）"))
        t = doc.Tables.Add(doc.Paragraphs(doc.Paragraphs.Count).Range, len(rows), 2)
        for i, (a, b) in enumerate(rows, start=1):
            set_cell(t, i, 1, a)
            set_cell(t, i, 2, b)

        add_para(doc, "")
        for title in ("第一部分  申请理由", "第二部分  个人情况", "第三部分  承诺"):
            add_para(doc, title, "标题 2")
            add_para(doc, "（请在此填写正文）")

        doc.SaveAs2(str(OUT.resolve()))
        print("created", OUT)
    finally:
        doc.Close(False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
