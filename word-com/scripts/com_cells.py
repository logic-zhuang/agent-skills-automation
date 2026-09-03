# -*- coding: utf-8 -*-
"""Table cell helpers — preserve font/alignment when replacing cell text."""
from __future__ import annotations

from difflib import SequenceMatcher
from typing import Any

wdCharacter = 1


def cell_body_range(cell: Any) -> Any:
    rng = cell.Range
    rng.MoveEnd(wdCharacter, -1)
    return rng


def replace_cell_text_keep_format(cell: Any, text: str) -> None:
    """Same pattern as cupk_com: MoveEnd(-1) keeps cell mark; restore font."""
    rng = cell_body_range(cell)
    al = rng.ParagraphFormat.Alignment
    fe = rng.Font.NameFarEast
    ascii_name = rng.Font.NameAscii
    size = rng.Font.Size
    bold = rng.Font.Bold
    rng.Text = text
    rng.ParagraphFormat.Alignment = al
    try:
        rng.Font.NameFarEast = fe
        rng.Font.NameAscii = ascii_name
        rng.Font.Size = size
        rng.Font.Bold = bold
    except Exception:
        pass


def replace_cell_checkbox_toggle(cell: Any, new_text: str) -> None:
    """Write multi-□ cell by replacing the whole body with TrackRevisions off.

    Char-diff under TrackRevisions often yields ``☑□选项`` ghosts; atomic replace
    with tracking paused keeps all options and only changes □/☑ in the saved text.
    """
    doc = cell.Range.Document
    prev = bool(doc.TrackRevisions)
    try:
        doc.TrackRevisions = False
        replace_cell_text_keep_format(cell, new_text)
    finally:
        try:
            doc.TrackRevisions = prev
        except Exception:
            pass


def replace_cell_text_minimal_diff(cell: Any, new_text: str) -> None:
    """Replace cell body via character-level diff (TrackRevisions-friendly for prose)."""
    rng = cell_body_range(cell)
    old = str(rng.Text or "")
    new_text = str(new_text or "")
    if old == new_text:
        return

    # Checkbox templates: never char-diff under TrackRevisions
    if ("□" in old or "☑" in old or "☐" in old) and ("□" in new_text or "☑" in new_text):
        replace_cell_checkbox_toggle(cell, new_text)
        return

    doc = rng.Document
    base = int(rng.Start)
    matcher = SequenceMatcher(None, old, new_text, autojunk=False)
    for tag, i1, i2, j1, j2 in reversed(matcher.get_opcodes()):
        if tag == "equal":
            continue
        target = doc.Range(base + i1, base + i2)
        piece = new_text[j1:j2]
        if tag == "delete":
            target.Delete()
        elif tag == "insert":
            target.InsertBefore(piece)
        else:
            target.Text = piece
