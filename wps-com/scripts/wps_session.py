# -*- coding: utf-8 -*-
"""Shared WPS Writer COM session helpers for wps-com scripts."""
from __future__ import annotations

import time
from pathlib import Path
from typing import Any

import win32com.client as win32

wdCharacter = 1

# WPS Writer ProgIDs only — never fall back to Word.Application.
WPS_PROGIDS = ("Kwps.Application", "wps.Application")

_last_backend: str | None = None


def backend_progid() -> str | None:
    """ProgID used by the last successful attach_wps call."""
    return _last_backend


def _ping_wps(wps: Any) -> None:
    """Raise if the COM proxy is stale (e.g. WPS just Quit in another process)."""
    _ = wps.Version
    _ = wps.Documents.Count


def _attach_progid(progid: str) -> Any:
    try:
        return win32.GetActiveObject(progid)
    except Exception:
        return win32.DispatchEx(progid)


def attach_wps(*, visible: bool = True, retries: int = 2) -> Any:
    global _last_backend

    wps: Any | None = None
    last_err: Exception | None = None

    for attempt in range(retries + 1):
        for progid in WPS_PROGIDS:
            try:
                wps = _attach_progid(progid)
                _ping_wps(wps)
                _last_backend = progid
                break
            except Exception as exc:
                last_err = exc
                wps = None
        if wps is not None:
            break
        if attempt < retries:
            time.sleep(0.4 * (attempt + 1))
            continue
        raise RuntimeError(
            "WPS Writer COM unavailable. Install WPS Office and ensure "
            "Kwps.Application or wps.Application is registered. "
            "This skill does not fall back to Microsoft Word."
        ) from last_err

    wps.Visible = bool(visible)
    wps.DisplayAlerts = 0
    try:
        wps.ScreenUpdating = True
    except Exception:
        pass
    return wps


def find_open_doc(wps: Any, path: Path) -> Any | None:
    target = str(path.resolve()).lower()
    for i in range(1, wps.Documents.Count + 1):
        d = wps.Documents(i)
        try:
            if str(Path(d.FullName).resolve()).lower() == target:
                return d
        except Exception:
            continue
    return None


def open_doc(wps: Any, path: Path, *, read_only: bool = False) -> Any:
    path = path.resolve()
    if path.name.startswith("~$"):
        raise ValueError(f"skip WPS lock file: {path}")
    if not path.exists():
        raise FileNotFoundError(path)

    last_err: Exception | None = None
    for attempt in range(3):
        try:
            doc = find_open_doc(wps, path)
            if doc is None:
                doc = wps.Documents.Open(FileName=str(path), ReadOnly=read_only)
            try:
                wps.Activate()
                doc.Activate()
            except Exception:
                pass
            _ = doc.Tables.Count
            return doc
        except Exception as exc:
            last_err = exc
            if attempt < 2:
                time.sleep(0.5 * (attempt + 1))
                wps = attach_wps(visible=wps.Visible)
                continue
            break
    raise RuntimeError(f"could not open document: {path}") from last_err


def para_text(p) -> str:
    return (p.Range.Text or "").replace("\r", "").replace("\x07", "").strip()


def cell_text(cell) -> str:
    return para_text(cell.Range.Paragraphs(1))


def clone_for_edit(wps: Any, src: Path, dest: Path) -> Path:
    """Copy src → dest on disk. Never write the original.

    OpenWorker / user session: dest under session workspace.
    """
    import shutil

    src = src.resolve()
    dest = dest.resolve()
    if src.name.startswith("~$") or dest.name.startswith("~$"):
        raise ValueError("skip WPS lock file")
    if not src.exists():
        raise FileNotFoundError(src)
    if src == dest:
        raise ValueError("dest must differ from src (refuse in-place edit of source)")

    dest.parent.mkdir(parents=True, exist_ok=True)
    existing_dest = find_open_doc(wps, dest)
    if existing_dest is not None:
        existing_dest.Close(SaveChanges=False)
    shutil.copy2(src, dest)
    return dest
