# -*- coding: utf-8 -*-
"""Quick structural checks for CUPK Phase 3 output."""
from __future__ import annotations

import re
import sys
from pathlib import Path

import win32com.client as win32

wdCharacter = 1
wdHeaderFooterPrimary = 1
wdHeaderFooterEvenPages = 3


def attach_word():
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    return word


def open_doc(word, path: Path):
    path = path.resolve()
    target = str(path).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            if str(Path(d.FullName).resolve()).lower() == target:
                return d, False
        except Exception:
            continue
    return word.Documents.Open(str(path), ReadOnly=True), True


def para_text(p) -> str:
    return (p.Range.Text or "").replace("\r", "").replace("\x07", "").strip()


def footer_text(sec, which: int) -> str:
    try:
        rng = sec.Footers(which).Range
        rng.MoveEnd(wdCharacter, -1)
        return (rng.Text or "").replace("\r", "").strip()
    except Exception:
        return ""


def header_text(sec, which: int) -> str:
    try:
        rng = sec.Headers(which).Range
        rng.MoveEnd(wdCharacter, -1)
        return (rng.Text or "").replace("\r", "").strip()
    except Exception:
        return ""


def main(argv: list[str] | None = None) -> int:
    path = Path(argv[0] if argv else sys.argv[1])
    word = attach_word()
    owned = False
    issues: list[str] = []
    ok: list[str] = []
    try:
        doc, owned = open_doc(word, path)
        ok.append(f"sections={doc.Sections.Count}")
        ok.append(f"paragraphs={doc.Paragraphs.Count}")
        ok.append(f"omaths={doc.OMaths.Count}")

        tips = 0
        for i in range(1, doc.Paragraphs.Count + 1):
            if para_text(doc.Paragraphs(i)).startswith("【格式提示】"):
                tips += 1
        if tips:
            issues.append(f"format tips remaining: {tips}")
        else:
            ok.append("no 【格式提示】")

        cover_title = ""
        if doc.Tables.Count >= 1:
            t = doc.Tables(1)
            try:
                cover_title = para_text(t.Cell(1, 2).Range.Paragraphs(1))
            except Exception:
                pass
        if "Phase3" in cover_title:
            ok.append(f"cover title ok: {cover_title[:40]}")
        else:
            issues.append(f"cover title unexpected: {cover_title!r}")

        sec2_footer = footer_text(doc.Sections(2), wdHeaderFooterPrimary)
        if re.search(r"-\s*[IVXLC]+\s*-", sec2_footer):
            ok.append(f"abstract footer roman: {sec2_footer}")
        else:
            issues.append(f"abstract footer not roman: {sec2_footer!r}")

        body_started = False
        for si in range(5, doc.Sections.Count + 1):
            sec = doc.Sections(si)
            ft = footer_text(sec, wdHeaderFooterPrimary)
            ht_odd = header_text(sec, wdHeaderFooterPrimary)
            ht_even = header_text(sec, wdHeaderFooterEvenPages)
            first_para = para_text(sec.Range.Paragraphs(1)) if sec.Range.Paragraphs.Count else ""
            if re.search(r"-\s*1\s*-", ft) and not body_started:
                body_started = True
                ok.append(f"body page 1 starts section {si} ({first_para[:20]})")
            if "中国石油大学（北京）本科毕业设计（论文）" in ht_odd:
                ok.append(f"section {si} odd header ok")
            if ht_even and first_para:
                ok.append(f"section {si} even header: {ht_even[:30]}")

        if not body_started:
            issues.append("body arabic page 1 footer not found from section 5")

        h1_left = []
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            try:
                st = str(p.Style.NameLocal)
            except Exception:
                continue
            if st == "标题 1" and p.Format.Alignment != 1:
                h1_left.append(para_text(p)[:30])
        if h1_left:
            issues.append(f"non-centered 标题 1: {h1_left}")
        else:
            ok.append("all 标题 1 centered")

        print("=== CUPK Phase 3 verify:", path.name, "===")
        for line in ok:
            print("OK ", line)
        for line in issues:
            print("FAIL", line)
        return 1 if issues else 0
    finally:
        if owned:
            try:
                doc.Close(False)
            except Exception:
                pass
            try:
                if word.Documents.Count == 0:
                    word.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
