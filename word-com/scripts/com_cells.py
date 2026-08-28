# -*- coding: utf-8
"""Table cell helpers — preserve font/alignment when replacing cell text."""
from __future__ import annotations

from typing import Any

wdCharacter = 1


def replace_cell_text_keep_format(cell: Any, text: str) -> None:
    """Same pattern as cupk_com: MoveEnd(-1) keeps cell mark; restore font."""
    rng = cell.Range
    rng.MoveEnd(wdCharacter, -1)
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
