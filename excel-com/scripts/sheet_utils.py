# -*- coding: utf-8 -*-
"""Sheet structure helpers: label detection, aliases, fingerprint."""
from __future__ import annotations

import hashlib
import json
import re
from typing import Any

# label_norm → profile key (under student / top-level)
PROFILE_KEY_MAP: dict[str, str] = {
    "姓名": "name",
    "名字": "name",
    "学生姓名": "name",
    "学号": "student_id",
    "学院": "college",
    "院系": "college",
    "专业": "major",
    "班级": "class_name",
    "年级": "grade",
    "联系电话": "phone",
    "手机": "phone",
    "电话": "phone",
    "手机号": "phone",
    "身份证": "id_number",
    "身份证号": "id_number",
    "邮箱": "email",
    "电子邮箱": "email",
}


def normalize_label(text: str) -> str:
    t = (text or "").strip()
    t = re.sub(r"[\s\u3000]+", "", t)
    t = t.strip("：:=:（）()【】[]")
    return t


def cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def a1_address(row: int, col: int) -> str:
    """1-based row/col → A1."""
    n = col
    letters = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters)) + str(row)


_SKIP_TOP_KEYS = {"student", "missing", "schema", "form_fingerprint"}


def flatten_kb(kb: dict[str, Any]) -> dict[str, Any]:
    flat: dict[str, Any] = {}
    if not isinstance(kb, dict):
        return flat
    student = kb.get("student")
    if isinstance(student, dict):
        for k, v in student.items():
            if v is not None and v != "":
                flat[k] = v
    for k, v in kb.items():
        if k in _SKIP_TOP_KEYS:
            continue
        if isinstance(v, dict):
            continue
        if v is not None and v != "":
            flat[k] = v
    return flat


def resolve_kb_value(label: str, kb_flat: dict[str, Any]) -> tuple[str | None, Any]:
    norm = normalize_label(label)
    key = PROFILE_KEY_MAP.get(norm)
    if key and key in kb_flat:
        return key, kb_flat[key]
    # direct key match
    if norm in kb_flat:
        return norm, kb_flat[norm]
    for k, v in kb_flat.items():
        if normalize_label(str(k)) == norm:
            return k, v
    return None, None


def _used_bounds(ws: Any) -> tuple[int, int, int, int]:
    used = ws.UsedRange
    r0 = int(used.Row)
    c0 = int(used.Column)
    r1 = r0 + int(used.Rows.Count) - 1
    c1 = c0 + int(used.Columns.Count) - 1
    return r0, c0, r1, c1


def build_sheet_facts(ws: Any, *, sheet_name: str | None = None) -> dict[str, Any]:
    """Agent-facing cell facts + layout_hint (no final fill decisions)."""
    name = sheet_name or ws.Name
    r0, c0, r1, c1 = _used_bounds(ws)
    cells: list[dict[str, Any]] = []
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            rng = ws.Cells(r, c)
            text = cell_text(rng.Value2)
            has_formula = False
            try:
                has_formula = bool(rng.HasFormula)
            except Exception:
                pass
            right_text = cell_text(ws.Cells(r, c + 1).Value2) if c <= c1 else ""
            below_text = cell_text(ws.Cells(r + 1, c).Value2) if r <= r1 else ""
            # peek one col beyond used for classic label|empty
            if c == c1:
                right_text = cell_text(ws.Cells(r, c + 1).Value2)
            cells.append(
                {
                    "sheet": name,
                    "row": r,
                    "col": c,
                    "address": a1_address(r, c),
                    "text": text,
                    "empty": not text,
                    "has_formula": has_formula,
                    "right_empty": not right_text,
                    "below_empty": not below_text,
                    "looks_like_label": bool(text) and len(text) <= 40 and not re.fullmatch(r"[\d./\-]+", text),
                }
            )

    # layout_hint: many non-empty on first used row + data rows below → table_rows
    header_nonempty = sum(1 for x in cells if x["row"] == r0 and x["text"])
    data_rows = r1 - r0
    label_right = sum(1 for x in cells if x["looks_like_label"] and x["right_empty"] and not x["empty"])
    layout_hint = "unknown"
    if header_nonempty >= 5 and data_rows >= 1:
        layout_hint = "table_rows"
    elif label_right >= 3:
        layout_hint = "label_left_value_right"

    return {
        "sheet": name,
        "used": {"r0": r0, "c0": c0, "r1": r1, "c1": c1},
        "cells": cells,
        "cell_count": len(cells),
        "layout_hint": layout_hint,
        "note": "facts only; Agent chooses write/append. suggestions/slots are optional heuristics.",
    }


def detect_slots_on_sheet(ws: Any, *, sheet_name: str | None = None) -> list[dict[str, Any]]:
    """Heuristic suggestions: non-empty label with empty cell immediately to the right."""
    name = sheet_name or ws.Name
    r0, c0, r1, c1 = _used_bounds(ws)
    slots: list[dict[str, Any]] = []
    slot_i = 0

    # Labels may sit in UsedRange while empty value cells sit one column outside it.
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            label_rng = ws.Cells(r, c)
            try:
                if label_rng.MergeCells:
                    ma = label_rng.MergeArea
                    if int(label_rng.Row) != int(ma.Row) or int(label_rng.Column) != int(ma.Column):
                        continue
            except Exception:
                pass

            label = cell_text(label_rng.Value2)
            if not label or len(label) > 40:
                continue
            # skip pure numbers / dates as labels
            if re.fullmatch(r"[\d./\-]+", label):
                continue

            value_col = c + 1
            value_rng = ws.Cells(r, value_col)
            try:
                if value_rng.MergeCells:
                    value_rng = value_rng.MergeArea.Cells(1, 1)
            except Exception:
                pass

            current = cell_text(value_rng.Value2)
            has_formula = False
            try:
                has_formula = bool(value_rng.HasFormula)
            except Exception:
                pass

            # only treat as fillable if empty (or whitespace) and not formula
            if current or has_formula:
                continue

            norm = normalize_label(label)
            if not norm:
                continue

            conf = 0.75
            fill_mode = "single"
            reason = "label_left_empty_right"
            if norm in PROFILE_KEY_MAP:
                conf = 0.9
            elif len(norm) < 2:
                fill_mode = "unknown"
                conf = 0.3
                reason = "short_label"

            slot_i += 1
            addr = a1_address(int(value_rng.Row), int(value_rng.Column))
            slots.append(
                {
                    "slot_id": f"{name}!{addr}",
                    "sheet": name,
                    "label": label,
                    "label_norm": norm,
                    "label_address": a1_address(r, c),
                    "address": addr,
                    "row": int(value_rng.Row),
                    "col": int(value_rng.Column),
                    "current": current,
                    "fill_mode": fill_mode,
                    "confidence": conf,
                    "reason": reason,
                    "has_formula": has_formula,
                    "targets": [{"sheet": name, "address": addr}],
                }
            )
    return slots


def inspect_workbook(wb: Any, *, sheet: str | int | None = None) -> dict[str, Any]:
    sheets_out: list[dict[str, Any]] = []
    all_slots: list[dict[str, Any]] = []
    all_facts: list[dict[str, Any]] = []
    layout_hints: list[str] = []

    indices: list[int]
    if sheet is None:
        indices = list(range(1, wb.Worksheets.Count + 1))
    elif isinstance(sheet, int):
        indices = [sheet]
    else:
        indices = [wb.Worksheets(sheet).Index]

    for idx in indices:
        ws = wb.Worksheets(idx)
        facts = build_sheet_facts(ws)
        all_facts.append(facts)
        layout_hints.append(facts.get("layout_hint") or "unknown")
        slots = detect_slots_on_sheet(ws)
        # On table_rows layouts, left→right suggestions are usually noise — demote
        if facts.get("layout_hint") == "table_rows":
            for s in slots:
                s["confidence"] = min(float(s.get("confidence") or 0.5), 0.4)
                s["fill_mode"] = "unknown"
                s["reason"] = "table_rows_layout_demoted"
        all_slots.extend(slots)
        r0, c0, r1, c1 = _used_bounds(ws)
        sheets_out.append(
            {
                "index": idx,
                "name": ws.Name,
                "used": {"r0": r0, "c0": c0, "r1": r1, "c1": c1},
                "slot_count": len(slots),
                "layout_hint": facts.get("layout_hint"),
            }
        )

    fingerprint = hashlib.sha256(
        json.dumps(
            [{"s": s["sheet"], "l": s["label_norm"], "a": s["address"]} for s in all_slots],
            ensure_ascii=False,
            sort_keys=True,
        ).encode("utf-8")
    ).hexdigest()[:16]

    primary_layout = "table_rows" if "table_rows" in layout_hints else (
        layout_hints[0] if layout_hints else "unknown"
    )

    return {
        "schema": "excel-com/form-inspect/v1.1",
        "sheets": sheets_out,
        "facts": {"sheets": all_facts, "layout_hint": primary_layout},
        "suggestions": all_slots,
        "slots": all_slots,
        "form_fingerprint": fingerprint,
        "agent_note": "Read facts.layout_hint. table_rows → Agent append_row; suggestions optional.",
    }


def verify_fill(wb: Any, cells: list[dict[str, Any]]) -> dict[str, Any]:
    ok = 0
    bad: list[dict[str, Any]] = []
    for item in cells:
        sheet = item.get("sheet")
        address = item.get("address")
        expected = cell_text(item.get("text", item.get("value", "")))
        ws = wb.Worksheets(sheet) if sheet else wb.Worksheets(1)
        actual = cell_text(ws.Range(address).Value2)
        if actual == expected:
            ok += 1
        else:
            bad.append({"sheet": sheet, "address": address, "expected": expected, "actual": actual})
    return {"ok": ok, "total": len(cells), "mismatches": bad}
