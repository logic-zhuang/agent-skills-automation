# -*- coding: utf-8 -*-
"""Inspect Word table forms → structure report v2.

Usage:
  python word-com/scripts/inspect_form.py path/to.docx
  python word-com/scripts/inspect_form.py path/to.docx --json slots.json
  python word-com/scripts/inspect_form.py path/to.docx --table 1 --json slots.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from form_utils import (
    build_coverage,
    build_facts,
    detect_slots_in_table,
    dump_table_cells,
    inspect_document,
    normalize_label,
)
from word_session import attach_word, open_doc


def inspect_table(doc, table_index: int = 1) -> dict:
    if doc.Tables.Count < table_index:
        raise IndexError(f"document has {doc.Tables.Count} tables, requested {table_index}")
    t = doc.Tables(table_index)
    dump = dump_table_cells(t, table_index)
    slots, unresolved = detect_slots_in_table(t, table_index)
    rows = []
    for s in slots:
        rows.append(
            {
                "row": s.get("label_row", s["row"]),
                "label": s["label"],
                "value": s.get("current", ""),
                "label_col": s.get("label_col"),
                "value_col": s["col"],
                "value_row": s["row"],
                "fill_mode": s.get("fill_mode"),
                "confidence": s.get("confidence"),
            }
        )
    coverage = build_coverage(dump["cells"], slots, unresolved, dump["inaccessible"])
    tables_meta = [
        {
            "table": table_index,
            "rows": dump["rows"],
            "cols": dump["cols"],
            "slot_count": len(slots),
            "unresolved_count": len(unresolved),
            "inaccessible_count": dump["inaccessible_count"],
            "inaccessible": dump["inaccessible"],
        }
    ]
    facts = build_facts(dump["cells"], tables_meta=tables_meta)
    return {
        "schema": "word-com/form-inspect/v2.2",
        "table": table_index,
        "rows": rows,
        "row_count": t.Rows.Count,
        "cells": dump["cells"],
        "inaccessible": dump["inaccessible"],
        "inaccessible_count": dump["inaccessible_count"],
        "facts": facts,
        "suggestions": slots,
        "slots": slots,
        "coverage": coverage,
        "agent_note": "Read facts (+ coverage). suggestions are optional heuristics; Agent owns fill plan.",
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inspect Word table form structure (v2)")
    parser.add_argument("path")
    parser.add_argument("--table", type=int, default=0, help="Single table index; 0 = all tables")
    parser.add_argument("--json", metavar="OUT.json")
    parser.add_argument("--hidden", action="store_true")
    args = parser.parse_args(argv)

    path = Path(args.path)
    word = attach_word(visible=not args.hidden)
    doc = open_doc(word, path, read_only=True)

    if args.table and args.table > 0:
        payload = inspect_table(doc, args.table)
        payload["path"] = str(path.resolve())
        payload["labels"] = sorted({normalize_label(s["label"]) for s in payload["slots"]})
    else:
        payload = inspect_document(doc)
        payload["path"] = str(path.resolve())

    if args.json:
        Path(args.json).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print("wrote", args.json)
        summary = {
            "schema": payload.get("schema"),
            "slot_count": payload.get("slot_count") or len(payload.get("slots") or []),
            "form_fingerprint": payload.get("form_fingerprint"),
            "modes": {},
            "coverage": {
                "unresolved_label_count": (payload.get("coverage") or {}).get("unresolved_label_count"),
                "unbound_empty_count": (payload.get("coverage") or {}).get("unbound_empty_count"),
                "label_coverage": (payload.get("coverage") or {}).get("label_coverage"),
            },
        }
        for s in payload.get("slots") or []:
            m = s.get("fill_mode") or s.get("kind") or "?"
            summary["modes"][m] = summary["modes"].get(m, 0) + 1
        print(json.dumps(summary, ensure_ascii=False))
    else:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
