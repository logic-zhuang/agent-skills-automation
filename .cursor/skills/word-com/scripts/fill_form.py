# -*- coding: utf-8 -*-
"""Fill table cells from JSON. Default: TrackRevisions on, do NOT finalize Save.

Usage:
  python word-com/scripts/fill_form.py path/to.docx --json fill.json
  python word-com/scripts/fill_form.py path/to.docx --json fill.json --save
  python word-com/scripts/fill_form.py path/to.docx --json fill.json --draft out.docx
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

from com_cells import replace_cell_text_keep_format
from form_utils import verify_fill
from word_session import attach_word, cell_text, open_doc


def normalize_label(text: str) -> str:
    return re.sub(r"\s+", "", text).strip("：:")


def fill_by_cells(doc, table_index: int, cells: list[dict]) -> int:
    n = 0
    for item in cells:
        t_idx = int(item.get("table") or table_index)
        t = doc.Tables(t_idx)
        r, c = int(item["row"]), int(item.get("col", 2))
        text = item.get("text", "")
        replace_cell_text_keep_format(t.Cell(r, c), str(text))
        n += 1
    return n


def fill_by_fields(doc, table_index: int, fields: dict[str, str]) -> int:
    t = doc.Tables(table_index)
    norm_fields = {normalize_label(k): v for k, v in fields.items()}
    n = 0
    for r in range(1, t.Rows.Count + 1):
        try:
            label = cell_text(t.Cell(r, 1)) if t.Columns.Count >= 2 else ""
        except Exception:
            continue
        key = normalize_label(label)
        if not key or key not in norm_fields:
            continue
        col = 2 if t.Columns.Count >= 2 else 1
        replace_cell_text_keep_format(t.Cell(r, col), norm_fields[key])
        n += 1
    return n


def fill_form(doc, payload: dict) -> int:
    cells = payload.get("cells")
    if cells:
        default_table = int(payload.get("table") or 1)
        return fill_by_cells(doc, default_table, cells)
    table_index = int(payload.get("table") or 1)
    if payload.get("fields"):
        return fill_by_fields(doc, table_index, payload["fields"])
    raise ValueError("JSON needs 'fields' or 'cells'")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Fill Word table (review-first)")
    parser.add_argument("path")
    parser.add_argument("--json", required=True)
    parser.add_argument("--hidden", action="store_true")
    parser.add_argument(
        "--save",
        action="store_true",
        help="Finalize Save to original path (default: do not save final)",
    )
    parser.add_argument("--draft", metavar="OUT.docx", help="SaveAs draft copy for review")
    parser.add_argument("--no-track", action="store_true", help="Disable TrackRevisions")
    parser.add_argument("--verify", action="store_true", help="Re-read cells after fill")
    # backward compat: --no-save was old skip; now default is no-save
    parser.add_argument("--no-save", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args(argv)

    path = Path(args.path)
    payload = json.loads(Path(args.json).read_text(encoding="utf-8"))
    word = attach_word(visible=not args.hidden)
    doc = open_doc(word, path)

    if not args.no_track:
        try:
            doc.TrackRevisions = True
        except Exception:
            pass

    count = fill_form(doc, payload)

    verification = None
    if args.verify and payload.get("cells"):
        verification = verify_fill(doc, payload["cells"])

    saved = False
    draft_path = None
    if args.draft:
        draft_path = str(Path(args.draft).resolve())
        doc.SaveAs2(draft_path)
        saved = True
        draft_path = draft_path
    elif args.save and not args.no_save:
        doc.Save()
        saved = True

    result: dict[str, Any] = {
        "filled": count,
        "saved": saved,
        "track_revisions": not args.no_track,
        "review_required": not args.save,
        "message": (
            "已写入（修订开启）。请在 Word 审阅中核对；确认后使用 --save 定稿，禁止 AcceptAll。"
            if not args.save
            else "已定稿 Save。仍建议人工核对；未自动 AcceptAll。"
        ),
    }
    if draft_path:
        result["draft"] = draft_path
    if verification is not None:
        result["verification"] = verification
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
