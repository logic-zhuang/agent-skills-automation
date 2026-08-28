# -*- coding: utf-8 -*-
"""Merge extra next-page sections; keep a single page break before 审阅.

Reuses the open demo document. Does not Quit.
"""
from __future__ import annotations

from pathlib import Path

import win32com.client as win32

wdCollapseStart = 1
wdPageBreak = 7
wdCharacter = 1

OUT = Path(__file__).resolve().parents[4] / "example" / "word_com_格式演示.docx"


def attach_word():
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    try:
        word.ScreenUpdating = True
    except Exception:
        pass
    return word


def get_doc(word):
    target = str(OUT.resolve()).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            full = str(Path(d.FullName).resolve()).lower()
        except Exception:
            continue
        if full == target or "格式演示" in d.Name:
            return d
    if not OUT.exists():
        raise FileNotFoundError(OUT)
    return word.Documents.Open(str(OUT), ReadOnly=False)


def delete_all_page_breaks(doc):
    n = 0
    while n < 30:
        rng = doc.Content
        f = rng.Find
        f.ClearFormatting()
        f.Text = "^m"
        f.Forward = True
        f.Wrap = 0
        if not f.Execute():
            break
        rng.Delete()
        n += 1
    return n


def main():
    word = attach_word()
    doc = get_doc(word)
    doc.Activate()
    print("before sections", doc.Sections.Count, "pages", doc.ComputeStatistics(2))

    for i in range(1, doc.Fields.Count + 1):
        fld = doc.Fields(i)
        if fld.Code.Text.strip().startswith("INDEX"):
            fld.Code.Text = ' INDEX \\z "2052" '
            try:
                fld.Update()
            except Exception:
                pass

    while doc.Sections.Count > 1:
        s = doc.Sections(2)
        doc.Range(s.Range.Start - 1, s.Range.Start).Delete()
    print("merged to", doc.Sections.Count, "section(s)")

    print("removed page breaks", delete_all_page_breaks(doc))

    h1 = None
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        try:
            t = p.Range.Text.replace("\r", "")
            if p.Style.NameLocal == "标题 1" and t.startswith("增补页：审阅、脚注、引用与导航") and "续" not in t:
                h1 = p
                break
        except Exception:
            continue
    if h1 is None:
        raise RuntimeError("review heading not found")

    rng = h1.Range
    rng.Collapse(wdCollapseStart)
    rng.InsertBreak(wdPageBreak)

    hr = h1.Range
    hr.MoveEnd(wdCharacter, -1)
    if doc.Bookmarks.Exists("SecReview"):
        doc.Bookmarks("SecReview").Delete()
    doc.Bookmarks.Add("SecReview", hr)

    try:
        doc.Save()
    except Exception:
        doc.SaveAs2(str(OUT), FileFormat=16)
    print("after sections", doc.Sections.Count, "pages", doc.ComputeStatistics(2), "FF", doc.Content.Text.count("\x0c"))
    print("Word window left open.")


if __name__ == "__main__":
    main()
