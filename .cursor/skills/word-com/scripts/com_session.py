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

from word_session import (
    app_install_path,
    attach_word,
    backend_progid,
    open_doc,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Word COM session helper")
    parser.add_argument("path", help="Absolute or relative .docx path")
    parser.add_argument("--quit", action="store_true", help="Close doc and Quit Word")
    parser.add_argument("--hidden", action="store_true")
    args = parser.parse_args(argv)

    path = Path(args.path)
    word = attach_word(visible=not args.hidden)
    doc = open_doc(word, path, read_only=False)
    print("backend", backend_progid() or "Word.Application")
    print("app", word.Version)
    print("path", app_install_path(word))
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
