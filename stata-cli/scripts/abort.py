"""Emergency stop for the Stata CLI skill — Ctrl+Shift+Q aborts this process only."""

from __future__ import annotations

import ctypes
import os
import sys
import threading
import time
from typing import Callable, Optional

_VK_CONTROL = 0x11
_VK_SHIFT = 0x10
_VK_Q = 0x51
_IS_WINDOWS = sys.platform == "win32"
_HOLD_S = 0.15
_down_since: Optional[float] = None
_watchdog_started = False
_ABORT_HOTKEY = "Ctrl+Shift+Q"


class ExploreAborted(Exception):
    """User requested stop via Ctrl+Shift+Q."""


def abort_hotkey_name() -> str:
    return _ABORT_HOTKEY


def _key_down(vk: int) -> bool:
    if not _IS_WINDOWS:
        return False
    try:
        return bool(ctypes.windll.user32.GetAsyncKeyState(vk) & 0x8000)
    except Exception:
        return False


def abort_chord_down() -> bool:
    return _key_down(_VK_CONTROL) and _key_down(_VK_SHIFT) and _key_down(_VK_Q)


def abort_chord_held(hold_s: float = _HOLD_S) -> bool:
    global _down_since
    if not abort_chord_down():
        _down_since = None
        return False
    now = time.monotonic()
    if _down_since is None:
        _down_since = now
        return False
    return (now - _down_since) >= hold_s


def check_abort(extra: Optional[Callable[[], bool]] = None) -> None:
    if abort_chord_held() or (extra and extra()):
        raise ExploreAborted(_ABORT_HOTKEY)


def start_abort_watchdog(
    *,
    out_dir: Optional[os.PathLike | str] = None,
    auto=None,
    hwnd: int = 0,
    pid: int = 0,
) -> None:
    """Daemon: Ctrl+Shift+Q exits this CLI process (does not kill the parent shell)."""
    global _watchdog_started
    if _watchdog_started:
        return
    _watchdog_started = True

    def _loop() -> None:
        while True:
            try:
                if abort_chord_held():
                    print(
                        f"\n*** STOP ({_ABORT_HOTKEY}) — exiting stata-cli ***\n",
                        flush=True,
                    )
                    if out_dir:
                        try:
                            Path = __import__("pathlib").Path
                            (Path(str(out_dir)) / "ABORTED").write_text(
                                f"{_ABORT_HOTKEY}\n", encoding="utf-8"
                            )
                        except Exception:
                            pass
                    os._exit(130)
            except Exception:
                pass
            time.sleep(0.05)

    threading.Thread(target=_loop, name="stata-cli-abort", daemon=True).start()
