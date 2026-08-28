"""Find and launch Stata for the CLI skill (no Soft Pack / ui_map)."""

from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Optional

EXE_NAMES = (
    "StataMP-64.exe",
    "StataSE-64.exe",
    "StataBE-64.exe",
    "StataMP-64",
    "StataSE-64",
    "StataBE-64",
)

# Prefer STATA_EXE / --exe. These are last-resort guesses only.
_CANDIDATE_REL = (
    ("Stata19", "StataMP-64.exe"),
    ("Stata19", "StataSE-64.exe"),
    ("Stata19", "StataBE-64.exe"),
    ("Stata18", "StataMP-64.exe"),
    ("Stata18", "StataSE-64.exe"),
    ("Stata17", "StataMP-64.exe"),
)


def _program_files_candidates() -> list[Path]:
    out: list[Path] = []
    for key in ("ProgramFiles", "ProgramFiles(x86)"):
        root = os.environ.get(key)
        if not root:
            continue
        base = Path(root)
        for folder, exe in _CANDIDATE_REL:
            out.append(base / folder / exe)
    # Common custom install locations (optional)
    out.extend(
        [
            Path(r"D:\soft\stata\StataMP-64.exe"),
            Path(r"C:\Program Files\Stata19\StataMP-64.exe"),
            Path(r"C:\Program Files\Stata18\StataMP-64.exe"),
        ]
    )
    return out


def _existing(path: Optional[Path]) -> Optional[Path]:
    if path is None:
        return None
    try:
        p = path.expanduser()
        if p.is_file():
            return p
    except OSError:
        return None
    return None


def resolve_stata_exe(explicit: str = "") -> Optional[Path]:
    if explicit:
        return _existing(Path(explicit))
    env = (os.environ.get("STATA_EXE") or "").strip()
    if env:
        found = _existing(Path(env))
        if found:
            return found
    for cand in _program_files_candidates():
        found = _existing(cand)
        if found:
            return found
    return None


def _tasklist_has(image: str) -> bool:
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", f"IMAGENAME eq {image}", "/FO", "CSV", "/NH"],
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except Exception:
        return False
    return image.lower() in (out or "").lower()


def stata_already_running() -> bool:
    if sys.platform != "win32":
        return False
    for name in EXE_NAMES:
        if _tasklist_has(name):
            return True
    return False


def wait_for_stata_window(*, wait_s: float = 60.0, settle_s: float = 1.0) -> bool:
    """Poll until the Stata main window is visible to UIA (process alone is not enough)."""
    if str(Path(__file__).resolve().parent) not in sys.path:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
    try:
        import uiautomation as uia
        from stata_session import find_stata_window
    except Exception:
        return False
    deadline = time.monotonic() + max(3.0, float(wait_s or 60))
    while time.monotonic() < deadline:
        try:
            find_stata_window(uia)
            time.sleep(max(0.3, float(settle_s or 1.0)))
            return True
        except Exception:
            time.sleep(0.4)
    return False


def launch_stata(*, exe: str = "", wait_s: float = 60.0) -> dict[str, Any]:
    """Start Stata if needed; wait until the main window is UIA-ready."""
    wait = max(5.0, float(wait_s or 60))
    if stata_already_running():
        if wait_for_stata_window(wait_s=min(20.0, wait)):
            return {"ok": True, "launched": False, "reused": True, "window_ready": True}
        return {
            "ok": False,
            "error": "stata_process_running_but_window_not_found",
            "hint": "Bring Stata to the foreground, or close orphan Stata processes.",
        }

    path = resolve_stata_exe(exe)
    if path is None:
        return {
            "ok": False,
            "error": "stata_exe_not_found",
            "hint": (
                "Set STATA_EXE to your StataMP-64.exe (or SE/BE), pass --exe, "
                "or start Stata manually."
            ),
        }
    try:
        subprocess.Popen(
            [str(path)],
            cwd=str(path.parent),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        return {"ok": False, "error": f"launch_failed:{exc}", "exe": str(path)}

    # Process first, then main window (UI can lag several seconds behind tasklist).
    proc_deadline = time.monotonic() + wait
    while time.monotonic() < proc_deadline:
        if stata_already_running():
            break
        time.sleep(0.4)
    else:
        return {
            "ok": False,
            "error": "launch_timeout_process",
            "exe": str(path),
            "wait_s": wait,
        }

    remaining = max(5.0, proc_deadline - time.monotonic())
    if not wait_for_stata_window(wait_s=remaining, settle_s=1.2):
        return {
            "ok": False,
            "error": "launch_timeout_window",
            "exe": str(path),
            "wait_s": wait,
            "hint": "Stata process started but main window not ready; retry connect.",
        }
    return {
        "ok": True,
        "launched": True,
        "exe": str(path),
        "window_ready": True,
    }
