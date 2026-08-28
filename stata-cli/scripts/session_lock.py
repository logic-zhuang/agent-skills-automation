"""Exclusive lock so only one cli.py touches Stata's Command window at a time.

Uses an atomic lock file (O_EXCL) so parallel *processes* cannot both type into
Stata. Stale locks (dead PID) are removed automatically.
"""

from __future__ import annotations

import os
import sys
import time
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional


class SessionBusy(Exception):
    """Another stata-cli process holds the session lock."""


def lock_path_for(out_dir: Path | str) -> Path:
    return Path(out_dir).expanduser() / ".stata_cli.lock"


def _pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform == "win32":
        try:
            import ctypes

            SYNCHRONIZE = 0x00100000
            handle = ctypes.windll.kernel32.OpenProcess(SYNCHRONIZE, 0, int(pid))
            if handle:
                ctypes.windll.kernel32.CloseHandle(handle)
                return True
            # Access denied often means process exists
            err = ctypes.windll.kernel32.GetLastError()
            return err == 5
        except Exception:
            return True
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False
    except Exception:
        return True


def _read_lock_pid(path: Path) -> Optional[int]:
    try:
        text = path.read_text(encoding="utf-8", errors="replace").strip()
    except OSError:
        return None
    for token in text.replace("=", " ").split():
        if token.isdigit():
            return int(token)
    return None


def _clear_stale(path: Path) -> bool:
    if not path.is_file():
        return False
    pid = _read_lock_pid(path)
    if pid is not None and _pid_alive(pid):
        return False
    try:
        path.unlink()
        return True
    except OSError:
        return False


@contextmanager
def exclusive_session(
    out_dir: Path | str,
    *,
    timeout_s: float = 180.0,
    poll_s: float = 0.3,
) -> Iterator[Path]:
    """Block until this process owns the session lock (or raise SessionBusy)."""
    path = lock_path_for(out_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    deadline = time.monotonic() + max(1.0, float(timeout_s or 180))
    fd: Optional[int] = None
    while True:
        try:
            fd = os.open(str(path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, f"pid={os.getpid()}\n".encode("utf-8"))
            break
        except FileExistsError:
            _clear_stale(path)
            if time.monotonic() >= deadline:
                holder = _read_lock_pid(path)
                raise SessionBusy(
                    f"stata-cli busy (another run holds {path}"
                    + (f", pid={holder}" if holder else "")
                    + "). Never run cli.py connect/run/read-* in parallel."
                )
            time.sleep(float(poll_s or 0.3))
        except OSError as exc:
            if time.monotonic() >= deadline:
                raise SessionBusy(
                    f"stata-cli busy (cannot create {path}: {exc}). "
                    "Never run cli.py connect/run/read-* in parallel."
                ) from exc
            time.sleep(float(poll_s or 0.3))
    assert fd is not None
    try:
        yield path
    finally:
        try:
            os.close(fd)
        except OSError:
            pass
        try:
            path.unlink(missing_ok=True)
        except TypeError:
            try:
                if path.is_file():
                    path.unlink()
            except OSError:
                pass
        except OSError:
            pass
