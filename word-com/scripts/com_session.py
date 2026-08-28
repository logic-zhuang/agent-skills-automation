# -*- coding: utf-8 -*-
"""Open a .docx via Word COM, reuse if already open, print basic stats.

Keeps Word visible and does not Quit (for interactive sessions).

Usage:
  python .cursor/skills/word-com/scripts/com_session.py path/to.docx
  python .cursor/skills/word-com/scripts/com_session.py path/to.docx --quit
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import win32com.client as win32


def attach_word(visible: bool = True):
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = bool(visible)
    word.DisplayAlerts = 0
    try:
        word.ScreenUpdating = True
    except Exception:
        pass
    return word


def find_open_doc(word, path: Path):
    target = str(path.resolve()).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            full = str(Path(d.FullName).resolve()).lower()
        except Exception:
            continue
        if full == target:
            return d
    return None


def open_doc(word, path: Path, read_only: bool = False):
    path = path.resolve()
    if not path.exists():
        raise FileNotFoundError(path)
    doc = find_open_doc(word, path)
    if doc is None:
        doc = word.Documents.Open(str(path), ReadOnly=read_only)
    try:
        word.Activate()
        doc.Activate()
    except Exception:
        pass
    return doc


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Word COM session helper")
    parser.add_argument("path", help="Absolute or relative .docx path")
    parser.add_argument("--quit", action="store_true", help="Close doc and Quit Word")
    parser.add_argument("--hidden", action="store_true")
    args = parser.parse_args(argv)

    path = Path(args.path)
    word = attach_word(visible=not args.hidden)
    doc = open_doc(word, path, read_only=False)
    print("app", word.Version)
    print("doc", doc.FullName)
    print("paragraphs", doc.Paragraphs.Count)
    print("sections", doc.Sections.Count)
    print("tables", doc.Tables.Count)
    print("comments", doc.Comments.Count)
    print("shapes", doc.Shapes.Count)
    print("inline_shapes", doc.InlineShapes.Count)
    print("omaths", doc.OMaths.Count)
    if args.quit:
        doc.Close(SaveChanges=0)
        word.Quit()
    else:
        print("Word window left open. Pass --quit to close.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
