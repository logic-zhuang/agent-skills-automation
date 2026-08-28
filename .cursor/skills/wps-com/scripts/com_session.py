# -*- coding: utf-8 -*-
"""Open a .docx via WPS Writer COM, reuse if already open, print basic stats.

Keeps WPS visible and does not Quit (for interactive sessions).

Usage:
  python wps-com/scripts/com_session.py path/to.docx
  python wps-com/scripts/com_session.py path/to.docx --quit
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from wps_session import attach_wps, backend_progid, find_open_doc, open_doc


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="WPS Writer COM session helper")
    parser.add_argument("path", help="Absolute or relative .docx path")
    parser.add_argument("--quit", action="store_true", help="Close doc and Quit WPS")
    parser.add_argument("--hidden", action="store_true")
    args = parser.parse_args(argv)

    path = Path(args.path)
    wps = attach_wps(visible=not args.hidden)
    doc = open_doc(wps, path, read_only=False)
    print("backend", backend_progid() or "?")
    print("app", wps.Version)
    print("doc", doc.FullName)
    print("paragraphs", doc.Paragraphs.Count)
    print("sections", doc.Sections.Count)
    print("tables", doc.Tables.Count)
    print("comments", doc.Comments.Count)
    print("shapes", doc.Shapes.Count)
    print("inline_shapes", doc.InlineShapes.Count)
    try:
        print("omaths", doc.OMaths.Count)
    except Exception as exc:
        print("omaths", f"unavailable ({exc})")
    if args.quit:
        doc.Close(SaveChanges=0)
        wps.Quit()
    else:
        print("WPS window left open. Pass --quit to close.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
