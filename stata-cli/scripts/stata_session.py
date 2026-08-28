"""Stata interactive session: run commands via the Command window + read outputs.

Capabilities
------------
- ``run(cmd)``: type into Command window (Scintilla) and press Enter
- ``read_results()``: Results window via Ctrl+A/C clipboard
- ``read_history()`` / ``read_variables()`` / ``read_properties()``: reuse pane readers
- Graphs:
    1. Prefer ``graph export`` to PNG under the session out dir
    2. Also screenshot any Stata ``Graph`` windows found under the process
- Tables: parse ASCII tables from the results text delta

Window binding prefers real Stata (``Stata/MP``, ``StataSE``, process name),
never Cursor/editor windows whose title merely contains ``Stata``.
"""

from __future__ import annotations

import json
import re
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

try:
    from . import panes as pane_mod
except ImportError:  # python scripts/cli.py
    import panes as pane_mod


@dataclass
class StataSession:
    title_contains: str = "Stata/"
    out_dir: Path = field(default_factory=lambda: Path(".coworker/stata_cli"))
    settle_s: float = 0.9
    auto_graph_export: bool = True

    root: Any = field(default=None, repr=False)
    uia: Any = field(default=None, repr=False)
    hwnd: int = 0
    pid: int = 0
    app: str = ""
    _last_results: str = ""
    _step: int = 0

    def connect(self) -> "StataSession":
        import uiautomation as uia

        try:
            from .abort import start_abort_watchdog
            from .window_geom import ensure_maximized, park_in_client
        except ImportError:
            from abort import start_abort_watchdog
            from window_geom import ensure_maximized, park_in_client

        self.uia = uia
        self.root = find_stata_window(uia, self.title_contains)
        self.hwnd = int(self.root.NativeWindowHandle or 0)
        self.pid = int(getattr(self.root, "ProcessId", 0) or 0)
        self.app = self.root.Name or "Stata"
        self.out_dir = Path(self.out_dir)
        self.out_dir.mkdir(parents=True, exist_ok=True)
        (self.out_dir / "graphs").mkdir(parents=True, exist_ok=True)
        (self.out_dir / "shots").mkdir(parents=True, exist_ok=True)
        try:
            self.root.SetFocus()
        except Exception:
            pass
        ensure_maximized(self.hwnd)
        park_in_client(self.hwnd)
        start_abort_watchdog(
            out_dir=self.out_dir, auto=uia, hwnd=self.hwnd, pid=self.pid
        )
        self._last_results = self.read_results().get("text") or ""
        return self

    # ------------------------------------------------------------------ run / read

    def run(self, command: str, *, settle_s: Optional[float] = None) -> dict[str, Any]:
        """Execute one command; return text delta, tables, graphs."""
        try:
            from .abort import check_abort
        except ImportError:
            from abort import check_abort

        check_abort()
        if self.root is None:
            self.connect()
        settle = self.settle_s if settle_s is None else settle_s
        self._step += 1
        step = self._step
        before_graphs = {g["hwnd"] for g in self.list_graph_windows()}

        # Marker in Results so we can slice this step's output reliably.
        marker = f"OPENWORK_STEP_{step}_{int(time.time())}"
        pane_mod._run_command(
            self.uia, self.root, f'display "{marker}"', settle_s=min(0.45, settle)
        )
        time.sleep(0.15)

        ok = pane_mod._run_command(self.uia, self.root, command, settle_s=settle)
        time.sleep(max(0.2, settle * 0.5))

        # If this looks like a graph command, export PNG via Stata itself.
        exported: list[dict[str, Any]] = []
        if self.auto_graph_export and _looks_like_graph_command(command):
            export_path = (self.out_dir / "graphs" / f"{step:03d}_export.png").resolve()
            # Use forward slashes — safer inside Stata strings on Windows.
            stata_path = str(export_path).replace("\\", "/")
            exp_cmd = f'graph export "{stata_path}", replace'
            pane_mod._run_command(self.uia, self.root, exp_cmd, settle_s=settle)
            time.sleep(0.35)
            if export_path.exists() and export_path.stat().st_size > 0:
                exported.append(
                    {
                        "path": str(export_path),
                        "rel": f"graphs/{export_path.name}",
                        "source": "graph_export",
                    }
                )

        after = self.read_results()
        text = after.get("text") or ""
        delta = _text_after_marker(text, marker)
        if not delta.strip():
            delta = _text_delta(self._last_results, text)
        self._last_results = text

        tables = extract_tables(delta)
        shots = self.capture_new_graphs(before_hwnds=before_graphs, step=step)
        graphs = exported + shots

        record = {
            "step": step,
            "command": command,
            "ok": bool(ok),
            "ts": datetime.now(timezone.utc).isoformat(),
            "output": delta,
            "tables": tables,
            "graphs": graphs,
            "results_line_count": after.get("line_count"),
        }
        self._append_jsonl(record)
        self._write_step_md(record)
        return record

    def read_results(self) -> dict[str, Any]:
        if self.root is None:
            self.connect()
        content = pane_mod._read_results(self.uia, self.root)
        text = content.get("raw_clipboard") or content.get("text_tail") or ""
        return {
            "text": text,
            "line_count": content.get("line_count") or len(text.splitlines()),
            "source": content.get("source"),
        }

    def read_history(self) -> dict[str, Any]:
        if self.root is None:
            self.connect()
        return pane_mod._read_history(self.uia, self.root)

    def read_variables(self) -> dict[str, Any]:
        if self.root is None:
            self.connect()
        return pane_mod._read_variables(self.uia, self.root)

    def clear_panes(
        self,
        *,
        clear_data: bool = True,
        clear_results: bool = True,
        clear_history: bool = True,
        drop_graphs: bool = True,
    ) -> dict[str, Any]:
        """Reset panes for a new task: data/vars/properties, Results, History, graphs.

        - Variables + Properties: via ``clear all`` (destroys data in memory)
        - Results: via ``cls``
        - History: UIA context-menu / Delete (no Stata command)
        - Graphs: ``graph drop _all`` when requested
        """
        try:
            from .abort import check_abort
        except ImportError:
            from abort import check_abort

        check_abort()
        if self.root is None:
            self.connect()

        report: dict[str, Any] = {"ok": True, "steps": []}

        def _note(name: str, rec: dict[str, Any]) -> None:
            report["steps"].append({"name": name, **rec})
            if not rec.get("ok", True):
                report["ok"] = False

        def _cmd(command: str, settle_s: float = 0.8) -> bool:
            return bool(
                pane_mod._run_command(
                    self.uia, self.root, command, settle_s=settle_s
                )
            )

        # Data first so Variables / Properties empty before we clear Results text.
        if clear_data:
            try:
                ok = _cmd("clear all", 0.9)
                _note("clear_all", {"ok": ok, "command": "clear all"})
            except Exception as exc:
                _note("clear_all", {"ok": False, "error": str(exc)})

        if drop_graphs:
            try:
                ok = _cmd("capture graph drop _all", 0.7)
                _note("graph_drop", {"ok": ok, "command": "capture graph drop _all"})
            except Exception as exc:
                _note("graph_drop", {"ok": False, "error": str(exc)})

        if clear_results:
            try:
                ok = _cmd("cls", 0.6)
                _note("cls", {"ok": ok, "command": "cls"})
                time.sleep(0.2)
                self._last_results = self.read_results().get("text") or ""
            except Exception as exc:
                _note("cls", {"ok": False, "error": str(exc)})

        # History last so clear/cls themselves do not remain in the History pane.
        if clear_history:
            try:
                hist = pane_mod._clear_history(self.uia, self.root)
                _note("history", hist)
            except Exception as exc:
                _note("history", {"ok": False, "error": str(exc)})

        try:
            _focus = pane_mod._focus_pane
            _focus(self.uia, self.root, "command")
        except Exception:
            pass

        try:
            report["variables"] = self.read_variables()
        except Exception as exc:
            report["variables"] = {"error": str(exc)}
        try:
            report["history"] = self.read_history()
        except Exception as exc:
            report["history"] = {"error": str(exc)}
        try:
            res = self.read_results()
            report["results_line_count"] = res.get("line_count")
            report["results_tail"] = (res.get("text") or "")[-300:]
        except Exception as exc:
            report["results"] = {"error": str(exc)}

        return report

    def read_properties(self) -> dict[str, Any]:
        if self.root is None:
            self.connect()
        pane_mod._select_first_variable_row(self.uia, self.root)
        time.sleep(0.2)
        return pane_mod._read_properties(self.root)

    # ------------------------------------------------------------------ graphs

    def list_graph_windows(self) -> list[dict[str, Any]]:
        """Graph windows belonging to this Stata pid (MDI or top-level)."""
        if self.root is None:
            return []
        found: list[dict[str, Any]] = []
        # Under main window
        for ctrl in _walk(self.root, maxd=6):
            rec = _as_graph_rec(ctrl, self.pid)
            if rec:
                found.append(rec)
        # Top-level desktop (some graphs float)
        try:
            for w in self.uia.GetRootControl().GetChildren():
                rec = _as_graph_rec(w, self.pid)
                if rec and rec["hwnd"] not in {f["hwnd"] for f in found}:
                    found.append(rec)
        except Exception:
            pass
        return found

    def capture_new_graphs(
        self, *, before_hwnds: set[int], step: int
    ) -> list[dict[str, Any]]:
        try:
            from .capture import capture_control
        except ImportError:
            from capture import capture_control

        out: list[dict[str, Any]] = []
        for g in self.list_graph_windows():
            if g["hwnd"] in before_hwnds:
                continue
            fname = f"{step:03d}_win_{g['hwnd']}.png"
            dest = self.out_dir / "graphs" / fname
            meta = capture_control(g["ctrl"], dest)
            if meta.get("captured"):
                out.append(
                    {
                        "path": str(dest),
                        "rel": f"graphs/{fname}",
                        "title": g.get("title"),
                        "hwnd": g["hwnd"],
                        "source": "uia_capture",
                        "bbox": meta.get("bbox"),
                    }
                )
        # If no "new" hwnd but a Graph exists, still capture the front-most once.
        if not out:
            graphs = self.list_graph_windows()
            if graphs:
                g = graphs[-1]
                fname = f"{step:03d}_win_current.png"
                dest = self.out_dir / "graphs" / fname
                meta = capture_control(g["ctrl"], dest)
                if meta.get("captured"):
                    out.append(
                        {
                            "path": str(dest),
                            "rel": f"graphs/{fname}",
                            "title": g.get("title"),
                            "hwnd": g["hwnd"],
                            "source": "uia_capture_current",
                            "bbox": meta.get("bbox"),
                        }
                    )
        return out

    # ------------------------------------------------------------------ persist

    def save_summary(self, steps: list[dict[str, Any]]) -> Path:
        summary = {
            "app": self.app,
            "hwnd": self.hwnd,
            "pid": self.pid,
            "explored_at": datetime.now(timezone.utc).isoformat(),
            "steps": steps,
            "variables": self.read_variables(),
            "history": self.read_history(),
            "properties": self.read_properties(),
        }
        path = self.out_dir / "session.json"
        path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
        md = [f"# Stata session — {self.app}", ""]
        for s in steps:
            md.append(f"## Step {s['step']}: `{s['command']}`")
            md.append("")
            md.append("### Output")
            md.append("```")
            md.append((s.get("output") or "").rstrip())
            md.append("```")
            md.append("")
            if s.get("tables"):
                md.append(f"### Tables ({len(s['tables'])})")
                for i, t in enumerate(s["tables"], 1):
                    md.append(f"#### Table {i}")
                    md.append("```")
                    md.append(t.get("text") or "")
                    md.append("```")
            if s.get("graphs"):
                md.append("### Graphs")
                for g in s["graphs"]:
                    rel = g.get("rel") or g.get("path")
                    md.append(f"- {g.get('source')}: `{rel}`")
                    if g.get("rel"):
                        md.append(f"  ![{rel}]({rel})")
            md.append("")
        md_path = self.out_dir / "session.md"
        md_path.write_text("\n".join(md), encoding="utf-8")
        return md_path

    def _append_jsonl(self, row: dict[str, Any]) -> None:
        slim = {k: v for k, v in row.items() if k != "ok" or True}
        # Drop huge duplicated fields from jsonl? keep output.
        with (self.out_dir / "session.jsonl").open("a", encoding="utf-8") as f:
            f.write(json.dumps(slim, ensure_ascii=False) + "\n")

    def _write_step_md(self, row: dict[str, Any]) -> None:
        p = self.out_dir / "steps" / f"{row['step']:03d}.md"
        p.parent.mkdir(parents=True, exist_ok=True)
        lines = [
            f"# Step {row['step']}",
            "",
            f"Command: `{row['command']}`",
            "",
            "## Output",
            "```",
            (row.get("output") or "").rstrip(),
            "```",
            "",
        ]
        for g in row.get("graphs") or []:
            if g.get("rel"):
                lines += [f"## Graph ({g.get('source')})", "", f"![]({g['rel']})", ""]
        p.write_text("\n".join(lines), encoding="utf-8")


# ------------------------------------------------------------------ helpers


def find_stata_window(uia, title_contains: str = "Stata/"):
    """Prefer real Stata main window over editors whose title mentions Stata."""
    needle = (title_contains or "Stata").lower()
    candidates = []
    for w in uia.GetRootControl().GetChildren():
        name = w.Name or ""
        cls = w.ClassName or ""
        if not name:
            continue
        low = name.lower()
        # Skip IDE / explorer false positives
        if "cursor" in low or "visual studio" in low or "文件资源管理器" in name:
            continue
        if "chrome_widget" in cls.lower():
            continue
        score = 0
        if needle in low:
            score += 2
        if low.startswith("stata/") or low.startswith("stata ") or "stata/mp" in low or "stata/se" in low or "stata/be" in low:
            score += 5
        if cls.startswith("Afx:"):
            score += 3
        if "stata" in low:
            score += 1
        if score >= 3:
            candidates.append((score, w))
    if not candidates:
        # Fallback: any Afx window with Stata in title
        for w in uia.GetRootControl().GetChildren():
            name = (w.Name or "").lower()
            cls = w.ClassName or ""
            if "stata" in name and cls.startswith("Afx:"):
                return w
        raise RuntimeError("Stata main window not found")
    candidates.sort(key=lambda x: -x[0])
    return candidates[0][1]


def _walk(root, maxd: int = 6):
    stack = [(root, 0)]
    while stack:
        x, d = stack.pop()
        yield x
        if d < maxd:
            try:
                for ch in x.GetChildren():
                    stack.append((ch, d + 1))
            except Exception:
                pass


def _as_graph_rec(ctrl, pid: int) -> Optional[dict[str, Any]]:
    try:
        if getattr(ctrl, "ControlTypeName", "") != "WindowControl":
            return None
        name = ctrl.Name or ""
        cls = ctrl.ClassName or ""
        cpid = int(getattr(ctrl, "ProcessId", 0) or 0)
        if pid and cpid and cpid != pid:
            return None
        low = name.lower()
        # Stata graph windows are typically titled "Graph" (sometimes with suffix)
        if not (low == "graph" or low.startswith("graph ") or name == "图形"):
            return None
        hwnd = int(ctrl.NativeWindowHandle or 0)
        if not hwnd:
            return None
        return {"ctrl": ctrl, "hwnd": hwnd, "title": name, "class_name": cls}
    except Exception:
        return None


def _looks_like_graph_command(cmd: str) -> bool:
    c = (cmd or "").strip().lower()
    if not c or c.startswith("graph export"):
        return False
    prefixes = (
        "graph ",
        "scatter",
        "line ",
        "twoway",
        "histogram",
        "kdensity",
        "box ",
        "hbox",
        "qnorm",
        "symplot",
        "lowess",
        "lpoly",
        "ac ",
        "pac ",
    )
    return c.startswith(prefixes) or c.split(",")[0].strip() in {
        "scatter",
        "histogram",
        "twoway",
        "graph",
    }


def _text_after_marker(text: str, marker: str) -> str:
    """Return Results text after the display marker line."""
    if not text or not marker:
        return text or ""
    # Stata echoes: . display "MARKER" then MARKER on next line
    idx = text.rfind(marker)
    if idx < 0:
        return ""
    # Skip the marker token itself and following newline
    rest = text[idx + len(marker) :]
    return rest.lstrip("\r\n")


def _text_delta(before: str, after: str) -> str:
    if not before:
        return after or ""
    if after.startswith(before):
        return after[len(before) :]
    # Fallback: last N lines of after
    lines = (after or "").splitlines()
    return "\n".join(lines[-60:])


_TABLE_LINE = re.compile(r".*\S.*\s{2,}\S")  # at least two columns separated by 2+ spaces


def extract_tables(text: str) -> list[dict[str, Any]]:
    """Heuristic: blocks of space-aligned columns in Results text."""
    lines = (text or "").splitlines()
    tables: list[dict[str, Any]] = []
    buf: list[str] = []

    def flush() -> None:
        nonlocal buf
        if len(buf) >= 3:
            tables.append({"text": "\n".join(buf), "n_lines": len(buf)})
        buf = []

    for ln in lines:
        if _TABLE_LINE.match(ln) or (buf and ln.strip().startswith("-" * 5)):
            buf.append(ln.rstrip())
        else:
            flush()
    flush()
    return tables
