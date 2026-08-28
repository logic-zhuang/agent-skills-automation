"""Keep the target app maximized and park the cursor safely inside the client."""

from __future__ import annotations

import ctypes
import sys
from typing import Any, Optional

_IS_WINDOWS = sys.platform == "win32"
SW_MAXIMIZE = 3
SW_RESTORE = 9


def ensure_maximized(hwnd: int) -> bool:
    """Maximize ``hwnd`` (no-op if already maximized). Returns True on success."""
    if not _IS_WINDOWS or not hwnd:
        return False
    try:
        user32 = ctypes.windll.user32
        hwnd = int(hwnd)
        if user32.IsZoomed(hwnd):
            user32.SetForegroundWindow(hwnd)
            return True
        user32.ShowWindow(hwnd, SW_MAXIMIZE)
        user32.SetForegroundWindow(hwnd)
        return bool(user32.IsZoomed(hwnd))
    except Exception:
        return False


def client_center(hwnd: int) -> Optional[tuple[int, int]]:
    """Screen coordinates of the window client-area center."""
    if not _IS_WINDOWS or not hwnd:
        return None
    try:
        from ctypes import wintypes

        user32 = ctypes.windll.user32
        rect = wintypes.RECT()
        if not user32.GetClientRect(int(hwnd), ctypes.byref(rect)):
            return None
        pt = wintypes.POINT(rect.left + (rect.right - rect.left) // 2, rect.top + (rect.bottom - rect.top) // 2)
        user32.ClientToScreen(int(hwnd), ctypes.byref(pt))
        return int(pt.x), int(pt.y)
    except Exception:
        return None


def park_in_client(hwnd: int) -> None:
    """Move cursor to client center — never screen (0,0) (resize / Aero edge hazards)."""
    if not _IS_WINDOWS:
        return
    pt = client_center(hwnd)
    if not pt:
        return
    try:
        ctypes.windll.user32.SetCursorPos(pt[0], pt[1])
    except Exception:
        pass


def window_rect(hwnd: int) -> dict[str, int]:
    if not _IS_WINDOWS or not hwnd:
        return {}
    try:
        from ctypes import wintypes

        rect = wintypes.RECT()
        if not ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(rect)):
            return {}
        return {
            "left": int(rect.left),
            "top": int(rect.top),
            "right": int(rect.right),
            "bottom": int(rect.bottom),
        }
    except Exception:
        return {}


def is_maximized(hwnd: int) -> bool:
    if not _IS_WINDOWS or not hwnd:
        return False
    try:
        return bool(ctypes.windll.user32.IsZoomed(int(hwnd)))
    except Exception:
        return False


def stabilize_main_window(hwnd: int, *, note_sink: Optional[list[Any]] = None) -> None:
    """Re-maximize if the explore run accidentally restored/resized the frame."""
    if not hwnd:
        return
    if is_maximized(hwnd):
        return
    ok = ensure_maximized(hwnd)
    if note_sink is not None:
        note_sink.append(f"re_maximized={ok}")
