# -*- coding: utf-8 -*-
"""Create a short Chinese essay .docx for revise demos."""
from __future__ import annotations

from pathlib import Path

import win32com.client as win32

OUT = Path(__file__).resolve().parent / "short_essay.docx"
wdCharacter = 1


def attach_word():
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    return word


def set_para_text(p, text: str) -> None:
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    rng.Text = text
    rng.Font.NameFarEast = "宋体"
    rng.Font.NameAscii = "Times New Roman"
    rng.Font.Size = 12


def main() -> int:
    word = attach_word()
    doc = word.Documents.Add()
    paras = [
        "人工智能正在改变教育方式。",
        "许多学校只好采用线上平台完成作业与考试，教师也只能通过屏幕点评学生作文。",
        "然而技术不能完全替代面对面的讨论与启发，师生互动仍然十分重要。",
    ]
    set_para_text(doc.Paragraphs(1), paras[0])
    for text in paras[1:]:
        doc.Paragraphs.Add()
        set_para_text(doc.Paragraphs(doc.Paragraphs.Count), text)

    # Drop trailing empty paras if any
    while doc.Paragraphs.Count > 3:
        last = doc.Paragraphs(doc.Paragraphs.Count)
        if not (last.Range.Text or "").replace("\r", "").strip():
            last.Range.Delete()
        else:
            break

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.SaveAs2(str(OUT.resolve()))
    print("created", OUT)
    print("paragraphs", doc.Paragraphs.Count)
    for i in range(1, doc.Paragraphs.Count + 1):
        t = (doc.Paragraphs(i).Range.Text or "").replace("\r", "").strip()
        print(f"  p{i}: {t}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
