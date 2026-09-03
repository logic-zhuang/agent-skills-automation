# -*- coding: utf-8 -*-
"""Shared Word COM session helpers for word-com scripts."""
from __future__ import annotations

import time
from pathlib import Path
from typing import Any

import win32com.client as win32

wdCharacter = 1

WORD_PROGID = "Word.Application"

# WPS often registers itself as Word.Application in the ROT; detect via install path.
_WPS_PATH_MARKERS = (
    "kingsoft",
    "\\wps office\\",
    "/wps office/",
    "\\wps office/",
)

_last_backend: str | None = None


def backend_progid() -> str | None:
    """ProgID reported after the last successful attach_word call."""
    return _last_backend


def _ping_word(word: Any) -> None:
    """Raise if the COM proxy is stale (e.g. Word just Quit in another process)."""
    _ = word.Version
    _ = word.Documents.Count


def app_install_path(app: Any) -> str:
    try:
        return str(app.Path or "")
    except Exception:
        return ""


def is_wps_application(app: Any) -> bool:
    """True when a COM app claiming Word.Application is actually WPS Writer."""
    path = app_install_path(app).lower().replace("/", "\\")
    if any(marker in path for marker in _WPS_PATH_MARKERS):
        return True
    # Some builds expose Build like 12.1.0.xxxxx under Kingsoft only; path is primary.
    return False


def _wps_hijack_error(path: str) -> RuntimeError:
    return RuntimeError(
        "Word.Application resolved to WPS Writer "
        f"(Path={path or '?'}), not Microsoft Word. "
        "Close WPS or use the wps-com skill (Kwps.Application). "
        "word-com does not drive WPS."
    )


def _try_active_word() -> Any | None:
    """Return running Microsoft Word, or None if absent / is WPS masquerade."""
    try:
        word = win32.GetActiveObject(WORD_PROGID)
    except Exception:
        return None
    _ping_word(word)
    if is_wps_application(word):
        return None
    return word


def _dispatch_real_word() -> Any:
    """Start a new Word.Application via LocalServer32; refuse WPS hijack."""
    word = win32.DispatchEx(WORD_PROGID)
    _ping_word(word)
    if is_wps_application(word):
        raise _wps_hijack_error(app_install_path(word))
    return word


def attach_word(*, visible: bool = True, retries: int = 2) -> Any:
    """Attach Microsoft Word only.

    Skips GetActiveObject when the ROT entry is WPS pretending to be Word,
    then DispatchEx to launch real WINWORD. Never falls back to Kwps.Application.
    """
    global _last_backend

    word: Any | None = None
    last_err: Exception | None = None
    for attempt in range(retries + 1):
        try:
            word = _try_active_word()
            if word is None:
                word = _dispatch_real_word()
            break
        except Exception as exc:
            last_err = exc
            word = None
            if attempt < retries:
                time.sleep(0.4 * (attempt + 1))
                continue
            if isinstance(last_err, RuntimeError) and "WPS Writer" in str(last_err):
                raise last_err
            raise RuntimeError(
                "Microsoft Word COM unavailable (Word.Application). "
                "Install/repair Microsoft Office Word. "
                "If only WPS is installed, use the wps-com skill instead."
            ) from last_err

    assert word is not None
    word.Visible = bool(visible)
    word.DisplayAlerts = 0
    try:
        word.ScreenUpdating = True
    except Exception:
        pass
    _last_backend = WORD_PROGID
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
    """Plain cell text; prefer Final view so deleted TrackRevisions ghosts are excluded."""
    try:
        doc = cell.Range.Document
        app = doc.Application
        view = app.ActiveWindow.View
        # wdRevisionsViewFinal = 0
        prev = view.RevisionsView
        view.RevisionsView = 0
        try:
            return para_text(cell.Range.Paragraphs(1))
        finally:
            view.RevisionsView = prev
    except Exception:
        return para_text(cell.Range.Paragraphs(1))


def clone_for_edit(word: Any, src: Path, dest: Path) -> Path:
    """Copy src → dest on disk. Never write the original.

    OpenWorker / user session: dest under session workspace.
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
