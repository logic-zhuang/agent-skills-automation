# -*- coding: utf-8 -*-
"""Shared Word COM session helpers for word-com scripts."""
from __future__ import annotations

import time
from pathlib import Path
from typing import Any

import win32com.client as win32

wdCharacter = 1


def _ping_word(word: Any) -> None:
    """Raise if the COM proxy is stale (e.g. Word just Quit in another process)."""
    _ = word.Version
    _ = word.Documents.Count


def attach_word(*, visible: bool = True, retries: int = 2) -> Any:
    word: Any | None = None
    last_err: Exception | None = None
    for attempt in range(retries + 1):
        try:
            try:
                word = win32.GetActiveObject("Word.Application")
            except Exception:
                word = win32.DispatchEx("Word.Application")
            _ping_word(word)
            break
        except Exception as exc:
            last_err = exc
            word = None
            if attempt < retries:
                time.sleep(0.4 * (attempt + 1))
                continue
            raise RuntimeError("Word COM unavailable") from last_err

    word.Visible = bool(visible)
    word.DisplayAlerts = 0
    try:
        word.ScreenUpdating = True
    except Exception:
        pass
    return word


def find_open_doc(word: Any, path: Path) -> Any | None:
    target = str(path.resolve()).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            if str(Path(d.FullName).resolve()).lower() == target:
                return d
        except Exception:
            continue
    return None


def open_doc(word: Any, path: Path, *, read_only: bool = False) -> Any:
    path = path.resolve()
    if path.name.startswith("~$"):
        raise ValueError(f"skip Word lock file: {path}")
    if not path.exists():
        raise FileNotFoundError(path)

    last_err: Exception | None = None
    for attempt in range(3):
        try:
            doc = find_open_doc(word, path)
            if doc is None:
                doc = word.Documents.Open(FileName=str(path), ReadOnly=read_only)
            try:
                word.Activate()
                doc.Activate()
            except Exception:
                pass
            _ = doc.Tables.Count
            return doc
        except Exception as exc:
            last_err = exc
            if attempt < 2:
                time.sleep(0.5 * (attempt + 1))
                word = attach_word(visible=word.Visible)
                continue
            break
    raise RuntimeError(f"could not open document: {path}") from last_err


def para_text(p) -> str:
    return (p.Range.Text or "").replace("\r", "").replace("\x07", "").strip()


def cell_text(cell) -> str:
    return para_text(cell.Range.Paragraphs(1))


def clone_for_edit(word: Any, src: Path, dest: Path) -> Path:
    """Copy src → dest on disk. Never write the original.

    OpenWorker: dest under session workspace. Demos may use examples/_out (temporary).
    """
    import shutil

    src = src.resolve()
    dest = dest.resolve()
    if src.name.startswith("~$") or dest.name.startswith("~$"):
        raise ValueError("skip Word lock file")
    if not src.exists():
        raise FileNotFoundError(src)
    if src == dest:
        raise ValueError("dest must differ from src (refuse in-place edit of source)")

    dest.parent.mkdir(parents=True, exist_ok=True)
    existing_dest = find_open_doc(word, dest)
    if existing_dest is not None:
        existing_dest.Close(SaveChanges=False)
    shutil.copy2(src, dest)
    return dest
