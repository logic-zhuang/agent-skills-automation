# -*- coding: utf-8 -*-
"""Shared Excel COM session helpers for excel-com scripts."""
from __future__ import annotations

import shutil
import time
from pathlib import Path
from typing import Any

import win32com.client as win32

XL_OPEN_XML_WORKBOOK = 51  # .xlsx


def _ping_excel(excel: Any) -> None:
    _ = excel.Version
    _ = excel.Workbooks.Count


def attach_excel(*, visible: bool = True, retries: int = 2) -> Any:
    excel: Any | None = None
    last_err: Exception | None = None
    for attempt in range(retries + 1):
        try:
            try:
                excel = win32.GetActiveObject("Excel.Application")
            except Exception:
                excel = win32.DispatchEx("Excel.Application")
            _ping_excel(excel)
            break
        except Exception as exc:
            last_err = exc
            excel = None
            if attempt < retries:
                time.sleep(0.4 * (attempt + 1))
                continue
            raise RuntimeError("Excel COM unavailable") from last_err

    excel.Visible = bool(visible)
    excel.DisplayAlerts = False
    try:
        excel.ScreenUpdating = True
    except Exception:
        pass
    return excel


def find_open_workbook(excel: Any, path: Path) -> Any | None:
    target = str(path.resolve()).lower()
    for i in range(1, excel.Workbooks.Count + 1):
        wb = excel.Workbooks(i)
        try:
            if str(Path(wb.FullName).resolve()).lower() == target:
                return wb
        except Exception:
            continue
    return None


def open_workbook(excel: Any, path: Path, *, read_only: bool = False) -> Any:
    path = path.resolve()
    if path.name.startswith("~$"):
        raise ValueError(f"skip Excel lock file: {path}")
    if not path.exists():
        raise FileNotFoundError(path)

    last_err: Exception | None = None
    for attempt in range(3):
        try:
            wb = find_open_workbook(excel, path)
            if wb is None:
                wb = excel.Workbooks.Open(
                    Filename=str(path),
                    ReadOnly=read_only,
                    UpdateLinks=0,
                    IgnoreReadOnlyRecommended=True,
                )
            if wb is None:
                raise RuntimeError(f"Workbooks.Open returned None for {path}")
            _ = int(wb.Worksheets.Count)
            try:
                excel.Activate()
                wb.Activate()
            except Exception:
                pass
            return wb
        except Exception as exc:
            last_err = exc
            time.sleep(0.3 * (attempt + 1))
    raise RuntimeError(f"failed to open workbook: {path}") from last_err


def clone_for_edit(excel: Any, src: Path, dest: Path) -> Path:
    """Copy src → dest on disk. Never write the original.

    OpenWorker / user session: dest under session workspace.
    """
    src = src.resolve()
    dest = dest.resolve()
    if src.name.startswith("~$") or dest.name.startswith("~$"):
        raise ValueError("skip Excel lock file")
    if not src.exists():
        raise FileNotFoundError(src)
    if src == dest:
        raise ValueError("dest must differ from src (refuse in-place edit of source)")

    dest.parent.mkdir(parents=True, exist_ok=True)

    existing_dest = find_open_workbook(excel, dest)
    if existing_dest is not None:
        existing_dest.Close(SaveChanges=False)

    shutil.copy2(src, dest)
    time.sleep(0.15)
    return dest


def save_as_xlsx(wb: Any, path: Path) -> None:
    """SaveAs to a new path (changes active workbook identity)."""
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        try:
            path.unlink()
        except OSError:
            pass
    wb.SaveAs(Filename=str(path), FileFormat=XL_OPEN_XML_WORKBOOK)


def save_copy_xlsx(wb: Any, path: Path) -> None:
    """Write a copy for review; keep the original workbook path/identity."""
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        try:
            path.unlink()
        except OSError:
            excel = wb.Application
            other = find_open_workbook(excel, path)
            if other is not None and other.FullName != wb.FullName:
                other.Close(SaveChanges=False)
            try:
                path.unlink()
            except OSError:
                pass
    wb.SaveCopyAs(Filename=str(path))
