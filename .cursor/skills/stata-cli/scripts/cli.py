"""Stata CLI skill entrypoint — Command window, Results, graphs. No menus/roadmaps."""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any, Optional

_SCRIPTS = Path(__file__).resolve().parent
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))


def _com_ready() -> None:
    try:
        import pythoncom

        pythoncom.CoInitialize()
    except Exception:
        pass


def _out_dir(explicit: str = "") -> Path:
    if explicit:
        dest = Path(explicit).expanduser()
    else:
        env = (os.environ.get("STATA_CLI_OUT") or "").strip()
        dest = Path(env) if env else Path.cwd() / ".coworker" / "stata_cli"
    dest.mkdir(parents=True, exist_ok=True)
    (dest / "graphs").mkdir(parents=True, exist_ok=True)
    return dest


def _emit(payload: dict[str, Any], *, exit_ok: bool = True) -> int:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False, indent=2) + "\n")
    sys.stdout.flush()
    return 0 if payload.get("ok", exit_ok) else 1


def _cd_workspace(session: Any, workspace: str) -> Optional[str]:
    if not workspace:
        return None
    try:
        root = Path(workspace).expanduser().resolve()
    except OSError:
        return None
    if not root.is_dir():
        return None
    target = str(root).replace("\\", "/")
    try:
        session.run(f'cd "{target}"', settle_s=0.7)
        return target
    except Exception:
        return None


def _connect(
    *,
    out_dir: str,
    workspace: str,
    exe: str,
    wait_s: float,
    clear: bool = False,
) -> dict[str, Any]:
    from launch import launch_stata, wait_for_stata_window
    from session_lock import SessionBusy, exclusive_session
    from stata_session import StataSession

    dest = _out_dir(out_dir)
    try:
        with exclusive_session(dest, timeout_s=max(60.0, float(wait_s or 60) + 30)):
            _com_ready()
            launched: dict[str, Any] | None = None
            try:
                session = StataSession(out_dir=dest, settle_s=1.0).connect()
            except Exception as first:
                launched = launch_stata(exe=exe, wait_s=wait_s)
                if not launched.get("ok"):
                    return {
                        "ok": False,
                        "error": str(first),
                        "launch": launched,
                        "hint": (
                            "Open Stata manually, set STATA_EXE, or pass --exe. "
                            "Then retry connect."
                        ),
                    }
                if not wait_for_stata_window(wait_s=15.0):
                    return {
                        "ok": False,
                        "error": "window_not_ready_after_launch",
                        "launch": launched,
                        "hint": "Stata started; wait a few seconds and retry connect.",
                    }
                session = StataSession(out_dir=dest, settle_s=1.0).connect()
            workspace_cd = _cd_workspace(session, workspace)
            cleared: dict[str, Any] | None = None
            if clear:
                cleared = session.clear_panes()
                # Re-apply cwd after clear all (clear does not change pwd, but be safe)
                if workspace:
                    workspace_cd = _cd_workspace(session, workspace) or workspace_cd
            out: dict[str, Any] = {
                "ok": True if not cleared else bool(cleared.get("ok")),
                "app": session.app,
                "pid": session.pid,
                "hwnd": session.hwnd,
                "out_dir": str(session.out_dir),
            }
            if workspace_cd:
                out["workspace_cd"] = workspace_cd
                out["hint"] = (
                    "Stata cwd set to workspace. Prefer absolute paths for do/export/save."
                )
            if launched is not None:
                out["launched"] = True
                out["launch"] = launched
            if cleared is not None:
                out["cleared"] = cleared
                if not cleared.get("ok"):
                    out["ok"] = False
                    out.setdefault(
                        "hint",
                        "Pane clear had errors; check cleared.steps before continuing.",
                    )
            return out
    except SessionBusy as exc:
        return {"ok": False, "error": str(exc), "hint": "Serialize cli.py calls."}


def _clear_panes(
    *, out_dir: str, workspace: str, exe: str
) -> dict[str, Any]:
    from launch import launch_stata, wait_for_stata_window
    from session_lock import SessionBusy, exclusive_session
    from stata_session import StataSession

    dest = _out_dir(out_dir)
    try:
        with exclusive_session(dest, timeout_s=180.0):
            _com_ready()
            try:
                session = StataSession(out_dir=dest, settle_s=1.0).connect()
            except Exception as first:
                launched = launch_stata(exe=exe, wait_s=60)
                if not launched.get("ok"):
                    return {"ok": False, "error": str(first), "launch": launched}
                if not wait_for_stata_window(wait_s=15.0):
                    return {
                        "ok": False,
                        "error": "window_not_ready_after_launch",
                        "launch": launched,
                    }
                session = StataSession(out_dir=dest, settle_s=1.0).connect()
            if workspace:
                _cd_workspace(session, workspace)
            cleared = session.clear_panes()
            if workspace:
                workspace_cd = _cd_workspace(session, workspace)
            else:
                workspace_cd = None
            out: dict[str, Any] = {
                "ok": bool(cleared.get("ok")),
                "app": session.app,
                "out_dir": str(session.out_dir),
                "cleared": cleared,
                "hint": (
                    "Panes reset for a new task (clear all + cls + history). "
                    "Next: run analysis commands."
                ),
            }
            if workspace_cd:
                out["workspace_cd"] = workspace_cd
            return out
    except SessionBusy as exc:
        return {"ok": False, "error": str(exc), "hint": "Serialize cli.py calls."}


def _run(
    command: str, *, out_dir: str, settle_s: float, workspace: str, exe: str
) -> dict[str, Any]:
    from launch import launch_stata, wait_for_stata_window
    from session_lock import SessionBusy, exclusive_session
    from stata_session import StataSession

    if not (command or "").strip():
        return {"ok": False, "error": "command is required"}
    dest = _out_dir(out_dir)
    try:
        with exclusive_session(dest, timeout_s=180.0):
            _com_ready()
            try:
                session = StataSession(out_dir=dest, settle_s=1.0).connect()
            except Exception as first:
                launched = launch_stata(exe=exe, wait_s=60)
                if not launched.get("ok"):
                    return {"ok": False, "error": str(first), "launch": launched}
                if not wait_for_stata_window(wait_s=15.0):
                    return {
                        "ok": False,
                        "error": "window_not_ready_after_launch",
                        "launch": launched,
                    }
                session = StataSession(out_dir=dest, settle_s=1.0).connect()
                _cd_workspace(session, workspace)
            if workspace:
                _cd_workspace(session, workspace)
            rec = session.run(command.strip(), settle_s=float(settle_s or 1.2))
            return {
                "ok": bool(rec.get("ok")),
                "step": rec.get("step"),
                "command": rec.get("command"),
                "output": rec.get("output"),
                "tables": rec.get("tables"),
                "graphs": rec.get("graphs"),
                "out_dir": str(session.out_dir),
            }
    except SessionBusy as exc:
        return {"ok": False, "error": str(exc), "hint": "Serialize cli.py calls."}


def _read(kind: str, *, out_dir: str) -> dict[str, Any]:
    from session_lock import SessionBusy, exclusive_session
    from stata_session import StataSession

    dest = _out_dir(out_dir)
    try:
        with exclusive_session(dest, timeout_s=120.0):
            _com_ready()
            session = StataSession(out_dir=dest, settle_s=1.0).connect()
            if kind == "results":
                return {"ok": True, **session.read_results()}
            if kind == "history":
                return {"ok": True, **session.read_history()}
            if kind == "variables":
                return {"ok": True, **session.read_variables()}
            return {"ok": False, "error": f"unknown_read:{kind}"}
    except SessionBusy as exc:
        return {"ok": False, "error": str(exc), "hint": "Serialize cli.py calls."}


def _add_common(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("--out-dir", default="", help="Session output folder (graphs/jsonl).")
    sp.add_argument("--workspace", default="", help="cd Stata pwd to this folder.")
    sp.add_argument("--exe", default="", help="Path to StataMP-64.exe (or STATA_EXE).")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="stata-cli",
        description="Drive Stata via the Command window; read Results and graphs.",
    )
    sub = p.add_subparsers(dest="cmd", required=True)

    c = sub.add_parser("connect", help="Attach or launch Stata")
    _add_common(c)
    c.add_argument("--wait-s", type=float, default=60.0)
    c.add_argument(
        "--clear",
        action="store_true",
        help="After connect, clear Results/variables/properties/history (new task).",
    )

    nt = sub.add_parser(
        "new-task",
        help="Connect (launch if needed) then clear panes — start of every new task.",
    )
    _add_common(nt)
    nt.add_argument("--wait-s", type=float, default=60.0)

    cp = sub.add_parser(
        "clear-panes",
        help="Clear Results (cls), variables/properties (clear all), and History.",
    )
    _add_common(cp)

    r = sub.add_parser("run", help='Send one command (or do "C:/abs/file.do")')
    _add_common(r)
    r.add_argument("command")
    r.add_argument("--settle-s", type=float, default=1.2)

    for name, help_s in (
        ("read-results", "Copy Results window text"),
        ("read-history", "Copy History pane commands"),
        ("read-variables", "Copy Variables pane names"),
    ):
        sp = sub.add_parser(name, help=help_s)
        _add_common(sp)

    m = sub.add_parser("materials", help="Look up CLI command dictionary")
    m.add_argument("action", choices=["index", "search"])
    m.add_argument("query", nargs="?", default="")
    m.add_argument("--family", default="")
    m.add_argument("--limit", type=int, default=6)
    return p


def main(argv: Optional[list[str]] = None) -> int:
    args = _build_parser().parse_args(argv)
    try:
        if args.cmd == "connect":
            return _emit(
                _connect(
                    out_dir=args.out_dir,
                    workspace=args.workspace,
                    exe=args.exe,
                    wait_s=args.wait_s,
                    clear=bool(getattr(args, "clear", False)),
                )
            )
        if args.cmd == "new-task":
            return _emit(
                _connect(
                    out_dir=args.out_dir,
                    workspace=args.workspace,
                    exe=args.exe,
                    wait_s=args.wait_s,
                    clear=True,
                )
            )
        if args.cmd == "clear-panes":
            return _emit(
                _clear_panes(
                    out_dir=args.out_dir,
                    workspace=args.workspace,
                    exe=args.exe,
                )
            )
        if args.cmd == "run":
            return _emit(
                _run(
                    args.command,
                    out_dir=args.out_dir,
                    settle_s=args.settle_s,
                    workspace=args.workspace,
                    exe=args.exe,
                )
            )
        if args.cmd == "read-results":
            return _emit(_read("results", out_dir=getattr(args, "out_dir", "")))
        if args.cmd == "read-history":
            return _emit(_read("history", out_dir=getattr(args, "out_dir", "")))
        if args.cmd == "read-variables":
            return _emit(_read("variables", out_dir=getattr(args, "out_dir", "")))
        if args.cmd == "materials":
            # Dictionary lookup does not touch Stata — no session lock.
            from materials import family_index, search_materials

            if args.action == "index":
                return _emit({"ok": True, **family_index()}, exit_ok=True)
            rec = search_materials(args.query, family=args.family, limit=args.limit)
            rec["ok"] = "error" not in rec
            return _emit(rec)
        return _emit({"ok": False, "error": f"unknown_cmd:{args.cmd}"})
    except Exception as exc:
        return _emit({"ok": False, "error": str(exc)})


if __name__ == "__main__":
    raise SystemExit(main())
