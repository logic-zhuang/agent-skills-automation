# -*- coding: utf-8 -*-
"""Inspect a Word document via COM (read-only).

Usage:
  python .cursor/skills/word-com/scripts/inspect_doc.py path/to.docx
  python .cursor/skills/word-com/scripts/inspect_doc.py path/to.docx --start 1 --count 40
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import win32com.client as win32


def attach_word():
    reused = True
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        reused = False
        word = win32.DispatchEx("Word.Application")
    word.DisplayAlerts = 0
    # Reuse visible session: do not hide or Quit Word afterward.
    if not reused:
        word.Visible = False
    else:
        try:
            word.ScreenUpdating = True
        except Exception:
            pass
    return word, reused


def open_readonly(word, path: Path):
    path = path.resolve()
    target = str(path).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            if str(Path(d.FullName).resolve()).lower() == target:
                return d, False
        except Exception:
            continue
    doc = word.Documents.Open(str(path), ReadOnly=True)
    return doc, True


def para_text(p) -> str:
    return (p.Range.Text or "").replace("\r", "").replace("\x07", "")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    parser.add_argument("--start", type=int, default=1)
    parser.add_argument("--count", type=int, default=30)
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    path = Path(args.path)
    if path.name.startswith("~$"):
        print("skip Word lock file", path, file=sys.stderr)
        return 2

    word, reused_word = attach_word()
    owned = False
    try:
        doc, owned = open_readonly(word, path)
        if not owned:
            try:
                word.Activate()
                doc.Activate()
            except Exception:
                pass
        info = {
            "path": str(path.resolve()),
            "paragraphs": doc.Paragraphs.Count,
            "sections": doc.Sections.Count,
            "tables": doc.Tables.Count,
            "comments": doc.Comments.Count,
            "revisions": doc.Revisions.Count,
            "shapes": doc.Shapes.Count,
            "inline_shapes": doc.InlineShapes.Count,
            "omaths": doc.OMaths.Count,
            "bookmarks": doc.Bookmarks.Count,
        }
        try:
            info["pages"] = doc.ComputeStatistics(2)
        except Exception:
            info["pages"] = None

        paras = []
        start = max(1, args.start)
        end = min(doc.Paragraphs.Count, start + args.count - 1)
        for i in range(start, end + 1):
            p = doc.Paragraphs(i)
            t = para_text(p)
            if not t.strip():
                continue
            try:
                style = str(p.Style.NameLocal)
            except Exception:
                style = "?"
            paras.append({"i": i, "style": style, "text": t[:200]})

        headers = []
        for s in range(1, doc.Sections.Count + 1):
            sec = doc.Sections(s)
            row = {"section": s}
            for name, idx in (("odd", 1), ("even", 3)):
                try:
                    row[name] = (sec.Headers(idx).Range.Text or "").replace("\r", " ").strip()[:60]
                except Exception:
                    row[name] = ""
            headers.append(row)

        comments = []
        for i in range(1, min(doc.Comments.Count, 20) + 1):
            c = doc.Comments(i)
            comments.append(
                {
                    "i": i,
                    "author": c.Author,
                    "body": (c.Range.Text or "").strip()[:200],
                    "scope": (c.Scope.Text or "").replace("\r", "")[:80],
                }
            )

        shapes = []
        for i in range(1, min(doc.Shapes.Count, 30) + 1):
            sh = doc.Shapes(i)
            try:
                txt = (sh.TextFrame.TextRange.Text or "").replace("\r", " ").strip()[:80]
            except Exception:
                txt = ""
            try:
                name = sh.Name
            except Exception:
                name = "?"
            shapes.append({"i": i, "name": name, "text": txt})

        payload = {
            "info": info,
            "paragraphs": paras,
            "headers": headers,
            "comments": comments,
            "shapes": shapes,
        }
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print("===", path.name, "===")
            for k, v in info.items():
                print(f"  {k}: {v}")
            print("--- paragraphs ---")
            for p in paras:
                print(f"{p['i']:4d} [{p['style']}] {p['text'][:80]}")
            print("--- section headers ---")
            for h in headers:
                print(h)
            if comments:
                print("--- comments ---")
                for c in comments:
                    print(c)
            if shapes:
                print("--- shapes ---")
                for s in shapes:
                    print(s)
        if owned:
            doc.Close(False)
    finally:
        # Only tear down a Word instance we started for headless inspect.
        if owned and not reused_word:
            try:
                if word.Documents.Count == 0:
                    word.Quit()
            except Exception:
                pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
