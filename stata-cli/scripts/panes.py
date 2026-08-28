"""Read/write Stata dock panes (command / history / variables / results).

Supports Chinese and English UI labels. Strategy:
  - Command: click Scintilla + SendKeys (+ Enter)
  - History / Variables / Results: focus pane + Ctrl+A/C → clipboard
  - Properties: LegacyIAccessible Name/Value on XTPPropertyGrid items
"""

from __future__ import annotations

import subprocess
import time
from typing import Any, Optional

# Logical pane → hotkey + possible UIA names (zh-CN then en).
_PANES: dict[str, dict[str, Any]] = {
    "command": {
        "hotkey": "{Ctrl}1",
        "names": ("命令窗口", "Command", "Command Window"),
    },
    "results": {
        "hotkey": "{Ctrl}2",
        "names": ("结果窗口", "Results", "Results Window"),
    },
    "history": {
        "hotkey": "{Ctrl}3",
        "names": ("历史窗口", "History", "History Window"),
    },
    "variables": {
        "hotkey": "{Ctrl}4",
        "names": ("变量窗口", "Variables", "Variables Window"),
    },
    "properties": {
        "hotkey": "{Ctrl}5",
        "names": ("属性窗口", "Properties", "Properties Window"),
    },
}


def _find_named(
    root, names: tuple[str, ...] | str, control_type: Optional[str] = None, maxd: int = 10
):
    want = (names,) if isinstance(names, str) else tuple(names)
    stack = [(root, 0)]
    while stack:
        x, d = stack.pop()
        try:
            n = x.Name or ""
            if n in want and (
                control_type is None or x.ControlTypeName == control_type
            ):
                return x
        except Exception:
            pass
        if d < maxd:
            try:
                for ch in x.GetChildren():
                    stack.append((ch, d + 1))
            except Exception:
                pass
    return None


def _find_pane(
    root, kind: str, control_type: Optional[str] = None, maxd: int = 10
):
    meta = _PANES.get(kind) or {}
    names = tuple(meta.get("names") or ())
    if not names:
        return None
    return _find_named(root, names, control_type=control_type, maxd=maxd)


def _find_class(root, class_name: str, maxd: int = 10):
    stack = [(root, 0)]
    while stack:
        x, d = stack.pop()
        try:
            if (x.ClassName or "") == class_name:
                return x
        except Exception:
            pass
        if d < maxd:
            try:
                for ch in x.GetChildren():
                    stack.append((ch, d + 1))
            except Exception:
                pass
    return None


def _focus_pane(uia, root, kind: str) -> None:
    meta = _PANES.get(kind) or {}
    hot = meta.get("hotkey")
    if hot:
        try:
            root.SetFocus()
        except Exception:
            pass
        time.sleep(0.05)
        uia.SendKeys(hot)
        time.sleep(0.15)
    node = _find_pane(root, kind)
    if node is None:
        return
    try:
        r = node.BoundingRectangle
        uia.Click((r.left + r.right) // 2, min(r.top + 24, (r.top + r.bottom) // 2))
    except Exception:
        pass
    time.sleep(0.1)


def _run_command(uia, root, command: str, *, settle_s: float = 0.8) -> bool:
    _focus_pane(uia, root, "command")
    tab = _find_pane(root, "command", "TabControl")
    sci = _find_class(tab or root, "Scintilla")
    if sci is None:
        # English builds sometimes nest Scintilla differently — last resort scan.
        sci = _find_class(root, "Scintilla", maxd=12)
    if sci is None:
        return False
    try:
        r = sci.BoundingRectangle
        uia.Click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
        time.sleep(0.12)
        uia.SendKeys("{Ctrl}a")
        time.sleep(0.05)
        uia.SendKeys(_escape_sendkeys(command), waitTime=0.01)
        time.sleep(0.08)
        uia.SendKeys("{Enter}")
        time.sleep(max(0.3, settle_s))
        return True
    except Exception:
        return False


def _escape_sendkeys(text: str) -> str:
    out = []
    for ch in text:
        if ch in "+^%~{}()":
            out.append("{" + ch + "}")
        else:
            out.append(ch)
    return "".join(out)


def _clipboard_get() -> str:
    try:
        return subprocess.check_output(
            ["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"],
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except Exception:
        return ""


def _clipboard_set(text: str) -> None:
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command", "Set-Clipboard -Value ''"],
            check=False,
            capture_output=True,
        )
    except Exception:
        pass


def _copy_from_control(uia, ctrl) -> str:
    if ctrl is None:
        return ""
    _clipboard_set("")
    try:
        r = ctrl.BoundingRectangle
        y = r.top + max(30, (r.bottom - r.top) // 3)
        uia.Click((r.left + r.right) // 2, min(y, r.bottom - 8))
        time.sleep(0.15)
        uia.SendKeys("{Ctrl}a")
        time.sleep(0.08)
        uia.SendKeys("{Ctrl}c")
        time.sleep(0.25)
        return _clipboard_get()
    except Exception:
        return ""


def _read_history(uia, root) -> dict[str, Any]:
    _focus_pane(uia, root, "history")
    tab = _find_pane(root, "history", "TabControl")
    table = _find_class(tab or root, "XTPReport")
    raw = _copy_from_control(uia, table or tab)
    lines = [ln.strip() for ln in (raw or "").splitlines() if ln.strip()]
    return {
        "headers": ["#", "Command", "_rc"],
        "row_count_uia": _count_report_rows(table),
        "commands": lines,
        "raw_clipboard": raw,
        "source": "clipboard",
    }


def _read_variables(uia, root) -> dict[str, Any]:
    _focus_pane(uia, root, "variables")
    tab = _find_pane(root, "variables", "TabControl")
    table = _find_class(tab or root, "XTPReport")
    raw = _copy_from_control(uia, table or tab)
    names: list[str] = []
    for ln in (raw or "").splitlines():
        ln = ln.strip()
        if not ln:
            continue
        names.extend([p for p in ln.split() if p])
    seen = set()
    ordered = []
    for n in names:
        if n not in seen:
            seen.add(n)
            ordered.append(n)
    return {
        "headers": ["name", "label"],
        "row_count_uia": _count_report_rows(table),
        "names": ordered,
        "raw_clipboard": raw,
        "source": "clipboard",
    }


def _select_first_variable_row(uia, root) -> None:
    _focus_pane(uia, root, "variables")
    tab = _find_pane(root, "variables", "TabControl")
    table = _find_class(tab or root, "XTPReport")
    if table is None:
        return
    try:
        for ch in table.GetChildren():
            if getattr(ch, "ControlTypeName", "") == "CustomControl":
                r = ch.BoundingRectangle
                uia.Click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
                return
        r = table.BoundingRectangle
        uia.Click((r.left + r.right) // 2, r.top + 40)
    except Exception:
        pass


def _read_properties(root) -> dict[str, Any]:
    tab = _find_pane(root, "properties", "TabControl")
    if tab is None:
        return {"properties": {}, "source": "uia", "error": "pane_not_found"}
    props: dict[str, str] = {}
    stack = [(tab, 0)]
    while stack:
        x, d = stack.pop()
        try:
            if getattr(x, "ControlTypeName", "") == "CustomControl":
                name = (x.Name or "").strip()
                if name:
                    try:
                        li = x.GetLegacyIAccessiblePattern()
                        val = li.Value
                    except Exception:
                        val = ""
                    if name not in props or (not props[name] and val):
                        props[name] = "" if val is None else str(val)
        except Exception:
            pass
        if d < 10:
            try:
                for ch in x.GetChildren():
                    stack.append((ch, d + 1))
            except Exception:
                pass
    return {"properties": props, "source": "uia_legacy_iaccessible"}


def _read_results(uia, root) -> dict[str, Any]:
    _focus_pane(uia, root, "results")
    win = _find_pane(root, "results", "WindowControl")
    if win is None:
        win = _find_pane(root, "results")
    raw = _copy_from_control(uia, win)
    lines = (raw or "").splitlines()
    tail = "\n".join(lines[-80:]) if lines else ""
    return {
        "line_count": len(lines),
        "text_tail": tail,
        "raw_clipboard": raw,
        "source": "clipboard",
    }


def _count_report_rows(table) -> int:
    if table is None:
        return 0
    n = 0
    try:
        for ch in table.GetChildren():
            if getattr(ch, "ControlTypeName", "") == "CustomControl":
                n += 1
    except Exception:
        pass
    return n


_CLEAR_HISTORY_LABELS = (
    "Clear History",
    "Clear history",
    "Clear All",
    "Clear all",
    "清除历史",
    "清除历史记录",
    "全部清除",
    "清除",
)


def _menu_label(ctrl) -> str:
    try:
        return (ctrl.Name or "").strip()
    except Exception:
        return ""


def _click_clear_history_menu(uia) -> bool:
    """Click a Clear/清除 item on the current context/popup menu."""
    try:
        desktop = uia.GetRootControl()
    except Exception:
        return False
    # Prefer freshly opened popup bars near the cursor.
    candidates = []
    try:
        for w in desktop.GetChildren():
            try:
                cls = w.ClassName or ""
                ctype = w.ControlTypeName or ""
            except Exception:
                continue
            if "XTPPopupBar" not in cls and ctype not in ("MenuControl", "Menu"):
                # Still walk shallow children of unnamed popups
                if "Popup" not in cls and "Menu" not in cls:
                    continue
            stack = [(w, 0)]
            while stack:
                x, d = stack.pop()
                label = _menu_label(x)
                low = label.lower()
                if label and any(
                    label == t or t.lower() in low for t in _CLEAR_HISTORY_LABELS
                ):
                    # Prefer more specific labels over bare "清除"
                    score = 2 if ("history" in low or "历史" in label) else 1
                    if label in ("清除", "Clear", "Clear all", "Clear All"):
                        score = 1
                    candidates.append((score, len(label), x))
                if d < 4:
                    try:
                        for ch in x.GetChildren():
                            stack.append((ch, d + 1))
                    except Exception:
                        pass
    except Exception:
        return False
    if not candidates:
        return False
    candidates.sort(key=lambda t: (-t[0], -t[1]))
    target = candidates[0][2]
    try:
        r = target.BoundingRectangle
        uia.Click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
        time.sleep(0.2)
        return True
    except Exception:
        try:
            target.Click(simulateMove=False)
            time.sleep(0.2)
            return True
        except Exception:
            return False


def _clear_history(uia, root) -> dict[str, Any]:
    """Clear the History pane (no official Stata command — UIA / keys)."""
    _focus_pane(uia, root, "history")
    tab = _find_pane(root, "history", "TabControl")
    table = _find_class(tab or root, "XTPReport")
    target = table or tab or _find_pane(root, "history")
    if target is None:
        return {"ok": False, "method": "none", "error": "history_pane_not_found"}

    try:
        r = target.BoundingRectangle
        x = (r.left + r.right) // 2
        y = min(r.top + 40, (r.top + r.bottom) // 2)
        uia.Click(x, y)
        time.sleep(0.12)
        uia.RightClick(x, y)
        time.sleep(0.25)
        if _click_clear_history_menu(uia):
            time.sleep(0.2)
            return {"ok": True, "method": "context_menu"}
    except Exception as exc:
        try:
            uia.SendKeys("{Esc}")
        except Exception:
            pass
        # fall through to key fallback
        _ = exc

    # Fallback: select all rows + Delete
    try:
        uia.SendKeys("{Esc}")
        time.sleep(0.05)
        _focus_pane(uia, root, "history")
        if target is not None:
            r = target.BoundingRectangle
            uia.Click((r.left + r.right) // 2, min(r.top + 40, (r.top + r.bottom) // 2))
        time.sleep(0.1)
        uia.SendKeys("{Ctrl}a")
        time.sleep(0.08)
        uia.SendKeys("{Delete}")
        time.sleep(0.2)
        return {"ok": True, "method": "ctrl_a_delete"}
    except Exception as exc:
        return {"ok": False, "method": "ctrl_a_delete", "error": str(exc)}
