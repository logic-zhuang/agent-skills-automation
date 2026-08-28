# -*- coding: utf-8
"""Fill hybrid doc: table fields + section headings/body."""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

from com_cells import replace_cell_text_keep_format
from fill_form import fill_by_fields, normalize_label
from word_session import attach_word, open_doc, para_text

wdCharacter = 1


def find_paragraph_by_heading(doc, heading: str) -> int | None:
    target = normalize_label(heading)
    for i in range(1, doc.Paragraphs.Count + 1):
        t = normalize_label(para_text(doc.Paragraphs(i)))
        if t == target or target in t:
            return i
    return None


def set_para_body(p, text: str) -> None:
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    rng.Text = text


def fill_section_after_heading(doc, heading: str, body: str) -> bool:
    idx = find_paragraph_by_heading(doc, heading)
    if idx is None:
        return False
    if idx + 1 <= doc.Paragraphs.Count:
        next_p = doc.Paragraphs(idx + 1)
        next_style = str(next_p.Style.NameLocal)
        if next_style.startswith("标题") or "Heading" in next_style:
            ins = doc.Paragraphs(idx).Range
            ins.Collapse(0)
            ins.InsertParagraphAfter()
            set_para_body(doc.Paragraphs(idx + 1), body)
            return True
        set_para_body(next_p, body)
        return True
    return False


def hybrid_fill(doc, payload: dict) -> dict:
    result = {"table_filled": 0, "sections_filled": 0, "missing_sections": []}
    if payload.get("table_fields") or payload.get("fields"):
        fields = payload.get("table_fields") or payload.get("fields")
        result["table_filled"] = fill_by_fields(doc, int(payload.get("table") or 1), fields)
    for sec in payload.get("sections") or []:
        heading = sec.get("heading") or ""
        body = sec.get("body") or sec.get("text") or ""
        if isinstance(body, list):
            body = "\r".join(body)
        if fill_section_after_heading(doc, heading, body):
            result["sections_filled"] += 1
        else:
            result["missing_sections"].append(heading)
    return result


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Fill hybrid Word document")
    parser.add_argument("path")
    parser.add_argument("--json", required=True)
    parser.add_argument("--hidden", action="store_true")
    parser.add_argument("--no-save", action="store_true")
    args = parser.parse_args(argv)

    path = Path(args.path)
    payload = json.loads(Path(args.json).read_text(encoding="utf-8"))
    word = attach_word(visible=not args.hidden)
    doc = open_doc(word, path)
    summary = hybrid_fill(doc, payload)
    if not args.no_save:
        doc.Save()
    summary["saved"] = not args.no_save
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
