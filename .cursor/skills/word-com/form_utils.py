# -*- coding: utf-8 -*-
"""Shared helpers for universal form inspect + plan + fill (v2)."""
from __future__ import annotations

import hashlib
import re
from typing import Any, Iterator

PLACEHOLDER_PATTERNS = (
    r"^[\s\u3000]*$",
    r"^[\(_（【\[]?请填写[\)_）】\]]?$",
    r"^[\(_（【\[]?请在此填写[\)_）】\]]?$",
    r"^[_＿\-—\.…·/\s]+$",
    r"^[（\(].*[填写填].*[）\)]$",
)
PLACEHOLDER_RES = [re.compile(p) for p in PLACEHOLDER_PATTERNS]

# Template cells like 「年　月　日」 / 「年　月」(not empty, but writable)
DATE_YMD_RE = re.compile(r"^[\s\u3000]*年[\s\u3000]*月[\s\u3000]*日[\s\u3000]*$")
DATE_YM_RE = re.compile(r"^[\s\u3000]*年[\s\u3000]*月[\s\u3000]*$")
BOX_MARK_RE = re.compile(r"[□☐☑✓]")

SECTION_HEADERS = frozenset(
    {
        "基本情况",
        "学习情况",
        "在校期间主要获奖情况",
        "论文发表情况",
        "申请理由",
        "学院意见",
        "学校意见",
        "院系意见",
        "中国石油奖学金领导小组意见",
    }
)

DEFAULT_ALIASES: dict[str, list[str]] = {
    "姓名": ["name", "student_name", "xm", "real_name"],
    "性别": ["gender", "sex", "xb"],
    "民族": ["ethnicity", "nation", "mz", "minzu"],
    "出生年月": ["birthday", "birth_date", "birth", "csny", "出生日期"],
    "出生日期": ["birthday", "birth_date", "birth", "csny", "出生年月"],
    "政治面貌": ["political_status", "politics", "zzmm"],
    "加入年月": ["political_join_date", "league_join_date", "join_date", "加入时间"],
    "学历": ["education", "degree", "xl"],
    "入学时间": ["enrollment_date", "enroll_date", "rxsj", "入学年月"],
    "入学年月": ["enrollment_date", "enroll_date", "rxsj", "入学时间"],
    "专业": ["major", "specialty", "zy"],
    "学制": ["schooling_length", "program_length", "xz"],
    "联系电话": ["phone", "mobile", "tel", "contact_phone", "lxdh"],
    "手机号": ["phone", "mobile", "tel"],
    "电话": ["phone", "mobile", "tel"],
    "手机": ["phone", "mobile", "tel"],
    "身份证号": ["id_number", "id_card", "identity", "sfzh"],
    "身份证号码": ["id_number", "id_card", "identity", "sfzh"],
    "身份证": ["id_number", "id_card", "identity", "sfzh"],
    "学号": ["student_id", "student_no", "sid", "xh"],
    "学院": ["college", "school", "department", "xy", "所在学院"],
    "所在学院": ["college", "school", "department", "xy"],
    "院系": ["college", "department", "yx"],
    "专业": ["major", "specialty", "zy"],
    "专业（类）": ["major", "specialty", "zy"],
    "班级": ["class_name", "class", "bj"],
    "籍贯": ["native_place", "jg"],
    "生源地": ["source_region", "student_origin", "syd"],
    "毕业中学": ["high_school", "graduate_school", "byzx"],
    "健康状况": ["health", "health_status"],
    "宗教信仰": ["religion"],
    "港澳台侨": ["hmt_status", "gangaotai"],
    "外语（语种）": ["foreign_language", "language"],
    "信教开始时间": ["religion_start"],
    "年级": ["grade", "nj"],
    "申请项目": ["application_item", "project", "sqxm"],
    "申请理由": ["application_reason", "reason", "sqly"],
    "申请理由(200字)": ["application_reason", "reason", "sqly"],
    "成绩排名": ["score_rank_text", "rank_display", "rank"],
    "绩点": ["gpa", "jd"],
    "邮箱": ["email", "mail"],
    "电子邮箱": ["email", "mail"],
    "家庭住址": ["address", "home_address"],
    "籍贯": ["native_place", "jg"],
    "导师": ["advisor", "supervisor", "ds"],
    "实行综合考评排名": ["comprehensive_ranking", "comprehensive_eval", "zhkp"],
    "必修课": ["required_courses_text", "required_courses"],
}

# Chinese key / alias → internal student field (for parse_profile)
PROFILE_KEY_MAP: dict[str, str] = {
    "姓名": "name",
    "名字": "name",
    "性别": "gender",
    "民族": "nation",
    "出生年月": "birthday",
    "出生日期": "birthday",
    "生日": "birthday",
    "政治面貌": "political_status",
    "加入年月": "political_join_date",
    "加入时间": "political_join_date",
    "学历": "education",
    "入学时间": "enrollment_date",
    "入学年月": "enrollment_date",
    "专业": "major",
    "学制": "schooling_length",
    "联系电话": "phone",
    "手机号": "phone",
    "手机": "phone",
    "电话": "phone",
    "身份证号": "id_number",
    "身份证号码": "id_number",
    "身份证": "id_number",
    "学号": "student_id",
    "学院": "college",
    "所在学院": "college",
    "院系": "college",
    "专业": "major",
    "专业（类）": "major",
    "班级": "class_name",
    "籍贯": "native_place",
    "生源地": "source_region",
    "毕业中学": "high_school",
    "健康状况": "health",
    "宗教信仰": "religion",
    "港澳台侨": "hmt_status",
    "外语（语种）": "foreign_language",
    "信教开始时间": "religion_start",
    "申请理由": "application_reason",
    "申请项目": "application_item",
    "成绩排名": "score_rank",
    "排名": "rank",
    "总人数": "rank_total",
    "必修课": "required_courses",
    "及格门数": "required_passed",
    "综合考评": "comprehensive_ranking",
    "实行综合考评排名": "comprehensive_ranking",
}

ID_LIKE_LABELS = frozenset({"身份证号", "身份证", "学号", "考生号", "准考证号"})
PHONE_LIKE_LABELS = frozenset({"联系电话", "手机号", "手机", "电话", "联系方式"})

CHECK_MARK = "☑"
CHECKBOX_PAIR_RE = re.compile(r"([是否])\s*([□☐☑✓])")
INLINE_RANK_RE = re.compile(r"(.*?排名：)\s*([^/]*?)\s*/\s*([^（]*?)（名次/总人数）")
INLINE_COURSE_RE = re.compile(r"(必修课)\s*(\S*)\s*(门，其中及格以上)\s*(\S*)\s*(门)")


def normalize_label(text: str) -> str:
    return re.sub(r"\s+", "", (text or "")).strip("：:;；")


def is_placeholder(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return True
    for rx in PLACEHOLDER_RES:
        if rx.match(t):
            return True
    if len(t) <= 4 and set(t) <= set("_＿-—./…· "):
        return True
    return False


def is_date_template(text: str) -> bool:
    """True for blank date scaffolds like 年/月/日 (not yet filled)."""
    t = (text or "").strip()
    if not t:
        return False
    if DATE_YMD_RE.match(t) or DATE_YM_RE.match(t):
        return True
    compact = re.sub(r"[\s\u3000]", "", t)
    return compact in {"年月日", "年月"}


def date_template_style(text: str) -> str:
    compact = re.sub(r"[\s\u3000]", "", (text or ""))
    return "ymd" if "日" in compact else "ym"


def count_box_marks(text: str) -> int:
    return len(BOX_MARK_RE.findall(text or ""))


def parse_date_parts(value: Any) -> tuple[str, str, str]:
    """Return (year, month, day); day may be '' if only year-month given."""
    s = str(value or "").strip()
    if not s:
        return "", "", ""
    m = re.match(r"^(\d{4})\D+(\d{1,2})(?:\D+(\d{1,2}))?", s)
    if m:
        day = m.group(3) or ""
        return m.group(1), m.group(2).zfill(2), day.zfill(2) if day else ""
    m = re.match(r"^(\d{4})年(\d{1,2})月(?:(\d{1,2})日)?", s)
    if m:
        day = m.group(3) or ""
        return m.group(1), m.group(2).zfill(2), day.zfill(2) if day else ""
    if re.fullmatch(r"\d{8}", s):
        return s[:4], s[4:6], s[6:8]
    if re.fullmatch(r"\d{6}", s):
        return s[:4], s[4:6], ""
    return "", "", ""


def apply_year_month_day(value: Any, style: str = "ymd") -> str:
    y, m, d = parse_date_parts(value)
    if not y:
        return str(value or "").strip()
    if style == "ym" or not d:
        return f"{y}年{m}月"
    return f"{y}年{m}月{d}日"


def looks_like_label(text: str) -> bool:
    t = (text or "").strip()
    if not t or is_placeholder(t):
        return False
    if is_date_template(t):
        return False
    # Checkbox option rows are not field labels
    if t.startswith(("□", "☐", "☑")) or (count_box_marks(t) >= 2 and len(re.sub(r"[\s□☐☑✓\t]", "", t)) <= 12):
        # allow「港澳台侨」style short labels with boxes only if boxes are few; multi-option lines out
        if count_box_marks(t) >= 2:
            return False
    norm = normalize_label(t)
    if norm in SECTION_HEADERS:
        return False
    if len(norm) < 2 and norm not in DEFAULT_ALIASES:
        return False
    if len(norm) > 40:
        return False
    if len(t) > 25 and "。" in t:
        return False
    if re.search(r"[\u4e00-\u9fff]", t):
        return True
    if re.fullmatch(r"[A-Za-z][A-Za-z0-9 _/-]{0,30}", t):
        return True
    return len(norm) <= 20


def is_section_header(text: str) -> bool:
    norm = normalize_label(text)
    if norm in SECTION_HEADERS:
        return True
    return norm.endswith("情况") and len(norm) <= 12


def is_digit_cell(text: str) -> bool:
    t = (text or "").strip()
    if not t:
        return True
    return len(t) <= 1 and (t.isdigit() or t in {"X", "x", "*"})


def is_phone_label(norm: str) -> bool:
    if norm in PHONE_LIKE_LABELS:
        return True
    return "电话" in norm or "手机" in norm


def is_id_like_label(norm: str) -> bool:
    if norm in ID_LIKE_LABELS:
        return True
    return "身份证" in norm or norm in {"学号", "考生号", "准考证号"}


def _cell_text(cell) -> str:
    from word_session import cell_text

    return cell_text(cell)


def iter_accessible_cells(table) -> Iterator[tuple[int, int, str]]:
    for r in range(1, table.Rows.Count + 1):
        for c in range(1, table.Columns.Count + 1):
            try:
                yield r, c, _cell_text(table.Cell(r, c))
            except Exception:
                continue


def dump_table_cells(table, table_index: int) -> dict[str, Any]:
    """Full cell grid dump for structure report."""
    cells: list[dict[str, Any]] = []
    inaccessible: list[dict[str, int]] = []
    rows, cols = table.Rows.Count, table.Columns.Count
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            try:
                text = _cell_text(table.Cell(r, c))
                cells.append({"table": table_index, "row": r, "col": c, "text": text})
            except Exception:
                inaccessible.append({"table": table_index, "row": r, "col": c})
    return {
        "table": table_index,
        "rows": rows,
        "cols": cols,
        "cells": cells,
        "inaccessible": inaccessible,
        "inaccessible_count": len(inaccessible),
    }


def extract_leading_label(text: str) -> str:
    head = (text or "").split("：", 1)[0].split(":", 1)[0]
    return normalize_label(head)


def has_checkbox_marks(text: str) -> bool:
    return bool(CHECKBOX_PAIR_RE.search(text or ""))


def consecutive_digit_cols(row_cells: list[tuple[int, str]], start_i: int) -> list[int]:
    cols: list[int] = []
    for j in range(start_i + 1, len(row_cells)):
        c_j, t_j = row_cells[j]
        if is_digit_cell(t_j):
            cols.append(c_j)
        else:
            break
    return cols


def classify_fill_zone(
    label_norm: str,
    label_text: str,
    row_cells: list[tuple[int, str]],
    label_i: int,
) -> dict[str, Any] | None:
    """Classify fill mode from label + rightward cells. Returns slot dict fields or None."""
    # Same-cell specials (checkbox / inline) take the cell itself
    if has_checkbox_marks(label_text):
        return None  # handled in special pass using cell text

    grid_cols = consecutive_digit_cols(row_cells, label_i)
    n = len(grid_cols)

    # ID-like: ≥15 consecutive digit cells → digit_grid
    if is_id_like_label(label_norm) and n >= 15:
        return {
            "fill_mode": "digit_grid",
            "kind": "digit_grid",
            "cols": grid_cols,
            "targets": [{"row": None, "col": c} for c in grid_cols],  # row filled by caller
            "confidence": 0.95,
            "reason": f"id_like+consecutive_empty={n}",
        }

    # Phone-like: 7–11 → digit_grid; 1 → single
    if is_phone_label(label_norm):
        if 7 <= n <= 11:
            return {
                "fill_mode": "digit_grid",
                "kind": "digit_grid",
                "cols": grid_cols,
                "targets": [{"row": None, "col": c} for c in grid_cols],
                "confidence": 0.9,
                "reason": f"phone_like+consecutive_empty={n}",
            }
        if n >= 1 and n < 7:
            # Prefer first empty as single if only 1 meaningful fill cell
            if n == 1 or (label_i + 1 < len(row_cells) and is_placeholder(row_cells[label_i + 1][1])):
                c_val = row_cells[label_i + 1][0]
                return {
                    "fill_mode": "single",
                    "kind": "field",
                    "cols": [c_val],
                    "targets": [{"row": None, "col": c_val}],
                    "confidence": 0.85,
                    "reason": f"phone_like+single_cell n={n}",
                }
        if n == 0 and label_i + 1 < len(row_cells):
            c_val, val_text = row_cells[label_i + 1]
            if is_placeholder(val_text) or not val_text.strip():
                return {
                    "fill_mode": "single",
                    "kind": "field",
                    "cols": [c_val],
                    "targets": [{"row": None, "col": c_val}],
                    "confidence": 0.8,
                    "reason": "phone_like+right_empty",
                }

    # Generic digit grid (学号 etc.) ≥11
    if n >= 11 and (is_id_like_label(label_norm) or label_norm in ID_LIKE_LABELS):
        return {
            "fill_mode": "digit_grid",
            "kind": "digit_grid",
            "cols": grid_cols,
            "targets": [{"row": None, "col": c} for c in grid_cols],
            "confidence": 0.85,
            "reason": f"digit_grid_min11 n={n}",
        }

    # Multiline body fields
    if "理由" in label_norm or "说明" in label_norm or "简介" in label_norm:
        if label_i + 1 < len(row_cells):
            c_val, val_text = row_cells[label_i + 1]
            if is_placeholder(val_text) or not val_text.strip():
                return {
                    "fill_mode": "multiline",
                    "kind": "multiline",
                    "cols": [c_val],
                    "targets": [{"row": None, "col": c_val}],
                    "confidence": 0.85,
                    "reason": "multiline_label",
                }

    # Right neighbor: date scaffold 「年　月　日」
    if label_i + 1 < len(row_cells):
        c_val, val_text = row_cells[label_i + 1]
        if is_date_template(val_text):
            style = date_template_style(val_text)
            return {
                "fill_mode": "year_month_day",
                "kind": "year_month_day",
                "cols": [c_val],
                "targets": [{"row": None, "col": c_val}],
                "confidence": 0.88,
                "reason": f"label+right_date_template:{style}",
                "date_style": style,
                "current": val_text.strip(),
            }
        # Multi □ in the value cell (not simple 是/否 pair)
        if count_box_marks(val_text) >= 2 and not has_checkbox_marks(val_text):
            return {
                "fill_mode": "unknown",
                "kind": "unknown",
                "cols": [c_val],
                "targets": [{"row": None, "col": c_val}],
                "confidence": 0.35,
                "reason": "multi_checkbox_right",
                "current": val_text.strip(),
            }

    # Standard single: next cell empty/placeholder
    if label_i + 1 < len(row_cells):
        c_val, val_text = row_cells[label_i + 1]
        if is_placeholder(val_text) or (not val_text.strip() and c_val != row_cells[label_i][0]):
            # Ambiguous many empties without id/phone label
            if n >= 7 and not is_id_like_label(label_norm) and not is_phone_label(label_norm):
                return {
                    "fill_mode": "unknown",
                    "kind": "unknown",
                    "cols": grid_cols[:n],
                    "targets": [{"row": None, "col": c} for c in grid_cols[:n]],
                    "confidence": 0.4,
                    "reason": f"ambiguous_grid n={n}",
                }
            return {
                "fill_mode": "single",
                "kind": "field",
                "cols": [c_val],
                "targets": [{"row": None, "col": c_val}],
                "confidence": 0.8,
                "reason": "label+right_empty",
                "current": val_text.strip(),
            }

    # Colon in same cell
    if "：" in label_text or ":" in label_text:
        parts = re.split(r"[：:]", label_text, maxsplit=1)
        if len(parts) == 2 and is_placeholder(parts[1]):
            c_label = row_cells[label_i][0]
            return {
                "fill_mode": "single",
                "kind": "field",
                "cols": [c_label],
                "targets": [{"row": None, "col": c_label}],
                "confidence": 0.7,
                "reason": "same_cell_after_colon",
                "label_override": parts[0].strip(),
                "current": parts[1].strip(),
            }

    return None


def try_below_fill_zone(
    label_norm: str,
    label_text: str,
    label_row: int,
    label_col: int,
    grid: dict[tuple[int, int], str],
) -> dict[str, Any] | None:
    """Label above / value below (same column)."""
    below_r = label_row + 1
    below_text = grid.get((below_r, label_col))
    if below_text is None:
        return None
    # Don't steal a cell that is itself a label for the next row
    if looks_like_label(below_text) and not is_date_template(below_text) and not is_placeholder(below_text):
        return None
    if is_date_template(below_text):
        style = date_template_style(below_text)
        return {
            "fill_mode": "year_month_day",
            "kind": "year_month_day",
            "cols": [label_col],
            "targets": [{"row": below_r, "col": label_col}],
            "confidence": 0.82,
            "reason": f"label+below_date_template:{style}",
            "date_style": style,
            "current": below_text.strip(),
        }
    if count_box_marks(below_text) >= 2 and not has_checkbox_marks(below_text):
        return {
            "fill_mode": "unknown",
            "kind": "unknown",
            "cols": [label_col],
            "targets": [{"row": below_r, "col": label_col}],
            "confidence": 0.35,
            "reason": "multi_checkbox_below",
            "current": below_text.strip(),
        }
    if is_placeholder(below_text):
        mode = "multiline" if ("理由" in label_norm or "说明" in label_norm or "简介" in label_norm) else "single"
        return {
            "fill_mode": mode,
            "kind": "multiline" if mode == "multiline" else "field",
            "cols": [label_col],
            "targets": [{"row": below_r, "col": label_col}],
            "confidence": 0.78,
            "reason": "label+below_empty",
            "current": below_text.strip(),
        }
    return None


def detect_special_slots_in_table(table, table_index: int) -> list[dict[str, Any]]:
    special: list[dict[str, Any]] = []
    seen: set[tuple[int, int, int]] = set()

    for r, c, text in iter_accessible_cells(table):
        if not text or not text.strip():
            continue

        if has_checkbox_marks(text):
            label_norm = extract_leading_label(text)
            if not label_norm:
                continue
            key = (table_index, r, c)
            if key in seen:
                continue
            seen.add(key)
            special.append(
                {
                    "slot_id": f"t{table_index}_r{r}_c{c}_cb",
                    "table": table_index,
                    "row": r,
                    "col": c,
                    "cols": [c],
                    "targets": [{"row": r, "col": c}],
                    "label": text.split("：", 1)[0].split(":", 1)[0].strip(),
                    "label_norm": label_norm,
                    "current": text.strip(),
                    "kind": "checkbox",
                    "fill_mode": "checkbox",
                    "confidence": 0.9,
                    "reason": "checkbox_marks",
                }
            )
            continue

        if INLINE_RANK_RE.search(text):
            label_norm = extract_leading_label(text) or "成绩排名"
            key = (table_index, r, c)
            if key in seen:
                continue
            seen.add(key)
            special.append(
                {
                    "slot_id": f"t{table_index}_r{r}_c{c}_rank",
                    "table": table_index,
                    "row": r,
                    "col": c,
                    "cols": [c],
                    "targets": [{"row": r, "col": c}],
                    "label": label_norm,
                    "label_norm": label_norm,
                    "current": text.strip(),
                    "kind": "inline_rank",
                    "fill_mode": "inline_rank",
                    "confidence": 0.9,
                    "reason": "inline_rank_pattern",
                }
            )
            continue

        if INLINE_COURSE_RE.search(text):
            key = (table_index, r, c)
            if key in seen:
                continue
            seen.add(key)
            special.append(
                {
                    "slot_id": f"t{table_index}_r{r}_c{c}_course",
                    "table": table_index,
                    "row": r,
                    "col": c,
                    "cols": [c],
                    "targets": [{"row": r, "col": c}],
                    "label": "必修课",
                    "label_norm": "必修课",
                    "current": text.strip(),
                    "kind": "inline_courses",
                    "fill_mode": "inline_courses",
                    "confidence": 0.9,
                    "reason": "inline_courses_pattern",
                }
            )

    return special


def detect_slots_in_table(table, table_index: int) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Unified structure detection → (slots, unresolved_labels)."""
    grid: dict[tuple[int, int], str] = {}
    by_row: dict[int, list[tuple[int, str]]] = {}
    for r, c, text in iter_accessible_cells(table):
        grid[(r, c)] = text
        by_row.setdefault(r, []).append((c, text))

    slots: list[dict[str, Any]] = []
    unresolved: list[dict[str, Any]] = []
    occupied: set[tuple[int, int, int]] = set()

    for r in sorted(by_row):
        row_cells = sorted(by_row[r], key=lambda x: x[0])
        for i, (c_label, label_text) in enumerate(row_cells):
            if not looks_like_label(label_text):
                continue
            if is_section_header(label_text):
                continue
            # Skip cells that are themselves checkbox/inline bodies
            if has_checkbox_marks(label_text) or INLINE_RANK_RE.search(label_text) or INLINE_COURSE_RE.search(
                label_text
            ):
                continue

            norm = normalize_label(label_text)
            zone = classify_fill_zone(norm, label_text, row_cells, i)
            if zone is None:
                zone = try_below_fill_zone(norm, label_text, r, c_label, grid)
            if zone is None:
                unresolved.append(
                    {
                        "table": table_index,
                        "label": label_text.strip(),
                        "label_norm": norm,
                        "label_row": r,
                        "label_col": c_label,
                        "reason": "no_right_or_below_writable",
                    }
                )
                continue

            override = zone.pop("label_override", None)
            label_out = override or label_text.strip()
            if override:
                norm = normalize_label(label_out)

            cols = zone.get("cols") or []
            targets = []
            for t in zone.get("targets") or []:
                tr = r if t.get("row") is None else t["row"]
                tc = t["col"]
                targets.append({"row": tr, "col": tc})

            if not targets:
                unresolved.append(
                    {
                        "table": table_index,
                        "label": label_text.strip(),
                        "label_norm": norm,
                        "label_row": r,
                        "label_col": c_label,
                        "reason": "zone_without_targets",
                    }
                )
                continue

            first = (table_index, targets[0]["row"], targets[0]["col"])
            if first in occupied:
                continue
            for tg in targets:
                occupied.add((table_index, tg["row"], tg["col"]))

            fill_mode = zone["fill_mode"]
            kind = zone.get("kind") or fill_mode
            slot_id = (
                f"t{table_index}_r{r}_grid_{norm}"
                if fill_mode == "digit_grid"
                else f"t{table_index}_r{targets[0]['row']}_c{targets[0]['col']}"
            )
            slot: dict[str, Any] = {
                "slot_id": slot_id,
                "table": table_index,
                "row": targets[0]["row"],
                "col": targets[0]["col"],
                "cols": cols,
                "targets": targets,
                "grid_length": len(cols) if fill_mode == "digit_grid" else 1,
                "label": label_out,
                "label_norm": normalize_label(label_out),
                "label_row": r,
                "label_col": c_label,
                "current": zone.get("current", ""),
                "kind": kind,
                "fill_mode": fill_mode,
                "confidence": zone.get("confidence", 0.5),
                "reason": zone.get("reason", ""),
            }
            if zone.get("date_style"):
                slot["date_style"] = zone["date_style"]
            slots.append(slot)

    special = detect_special_slots_in_table(table, table_index)
    for s in special:
        key = (s["table"], s["row"], s["col"])
        if key in occupied:
            continue
        occupied.add(key)
        slots.append(s)

    return slots, unresolved


def form_fingerprint_v2(slots: list[dict[str, Any]]) -> str:
    parts = []
    for s in slots:
        mode = s.get("fill_mode") or s.get("kind") or "field"
        n = len(s.get("targets") or s.get("cols") or [1])
        parts.append(f"{s.get('label_norm','')}:{mode}:{n}")
    payload = "|".join(sorted(parts))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def form_fingerprint(labels: list[str]) -> str:
    payload = "|".join(sorted(normalize_label(x) for x in labels if x))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def _cell_role(text: str) -> str:
    if is_date_template(text):
        return "date_template"
    if has_checkbox_marks(text) or count_box_marks(text) >= 2:
        return "checkbox_body"
    if is_section_header(text):
        return "section"
    if looks_like_label(text):
        return "label"
    if is_placeholder(text):
        return "empty_writable"
    if not (text or "").strip():
        return "empty_writable"
    return "content"


def build_coverage(
    all_cells: list[dict[str, Any]],
    slots: list[dict[str, Any]],
    unresolved_labels: list[dict[str, Any]],
    inaccessible: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Structure-side completeness (independent of profile/KB)."""
    claimed: set[tuple[int, int, int]] = set()
    for s in slots:
        claimed.add((int(s["table"]), int(s.get("label_row") or s["row"]), int(s.get("label_col") or s["col"])))
        for t in s.get("targets") or [{"row": s["row"], "col": s["col"]}]:
            claimed.add((int(s["table"]), int(t["row"]), int(t["col"])))

    unbound_empties: list[dict[str, Any]] = []
    role_counts: dict[str, int] = {}
    for cell in all_cells:
        role = _cell_role(cell.get("text") or "")
        role_counts[role] = role_counts.get(role, 0) + 1
        key = (int(cell["table"]), int(cell["row"]), int(cell["col"]))
        if role in {"empty_writable", "date_template"} and key not in claimed:
            unbound_empties.append(
                {
                    "table": cell["table"],
                    "row": cell["row"],
                    "col": cell["col"],
                    "role": role,
                    "text": (cell.get("text") or "")[:40],
                }
            )

    resolved = len(slots)
    unresolved_n = len(unresolved_labels)
    unbound_n = len(unbound_empties)
    # Heuristic: resolved labels vs resolved+unresolved; empties still unbound are soft debt
    denom = resolved + unresolved_n
    label_coverage = (resolved / denom) if denom else 1.0

    return {
        "resolved_slots": resolved,
        "unresolved_labels": unresolved_labels,
        "unresolved_label_count": unresolved_n,
        "unbound_empties": unbound_empties[:80],  # cap for report size
        "unbound_empty_count": unbound_n,
        "unbound_empties_truncated": unbound_n > 80,
        "inaccessible_count": len(inaccessible or []),
        "role_counts": role_counts,
        "label_coverage": round(label_coverage, 3),
        "scan_complete_enough": unresolved_n == 0 and unbound_n == 0,
        "note": "扫全看 unresolved_labels / unbound_empties，不看档案；档案多出来的字段正常。",
    }


def inspect_document(doc) -> dict[str, Any]:
    """Inspect all tables → structure report v2 (+ coverage)."""
    all_slots: list[dict[str, Any]] = []
    all_unresolved: list[dict[str, Any]] = []
    tables_meta: list[dict[str, Any]] = []
    all_cells: list[dict[str, Any]] = []
    all_inaccessible: list[dict[str, Any]] = []

    for ti in range(1, doc.Tables.Count + 1):
        t = doc.Tables(ti)
        dump = dump_table_cells(t, ti)
        all_cells.extend(dump["cells"])
        all_inaccessible.extend(dump["inaccessible"])
        slots, unresolved = detect_slots_in_table(t, ti)
        all_slots.extend(slots)
        all_unresolved.extend(unresolved)
        tables_meta.append(
            {
                "table": ti,
                "rows": dump["rows"],
                "cols": dump["cols"],
                "slot_count": len(slots),
                "unresolved_count": len(unresolved),
                "inaccessible_count": dump["inaccessible_count"],
                "inaccessible": dump["inaccessible"],
            }
        )

    labels = sorted({s["label_norm"] for s in all_slots})
    coverage = build_coverage(all_cells, all_slots, all_unresolved, all_inaccessible)
    return {
        "schema": "word-com/form-inspect/v2.1",
        "tables": tables_meta,
        "table_count": doc.Tables.Count,
        "cells": all_cells,
        "cell_count": len(all_cells),
        "slots": all_slots,
        "slot_count": len(all_slots),
        "labels": labels,
        "coverage": coverage,
        "form_fingerprint": form_fingerprint_v2(all_slots),
    }


def apply_digit_grid(value: Any, cols: list[int]) -> list[tuple[int, str]]:
    raw = re.sub(r"[\s\-]", "", str(value or ""))
    out: list[tuple[int, str]] = []
    for i, _col in enumerate(cols):
        out.append((i, raw[i] if i < len(raw) else ""))
    return out


def apply_checkbox_choice(text: str, choice: Any) -> str:
    yes = choice in (True, "true", "True", "yes", "Yes", "是", "y", "1", 1)
    if yes:
        return CHECKBOX_PAIR_RE.sub(lambda m: "是 ☑" if m.group(1) == "是" else "否 □", text)
    return CHECKBOX_PAIR_RE.sub(lambda m: "否 ☑" if m.group(1) == "否" else "是 □", text)


def apply_inline_rank(text: str, rank: Any, total: Any) -> str:
    return INLINE_RANK_RE.sub(lambda m: f"{m.group(1)}{rank} / {total}（名次/总人数）", text, count=1)


def apply_inline_courses(text: str, required: Any, passed: Any) -> str:
    return INLINE_COURSE_RE.sub(
        lambda m: f"{m.group(1)}{required}{m.group(3)}{passed}{m.group(5)}",
        text,
        count=1,
    )


def coerce_bool(value: Any) -> bool | None:
    if isinstance(value, bool):
        return value
    if value is None:
        return None
    s = str(value).strip().lower()
    if s in {"是", "yes", "true", "1", "y"}:
        return True
    if s in {"否", "no", "false", "0", "n"}:
        return False
    return None


def parse_rank_pair(value: Any) -> tuple[str | None, str | None]:
    if isinstance(value, dict):
        return str(value.get("rank", "") or value.get("numerator", "")), str(
            value.get("total", "") or value.get("denominator", "")
        )
    if isinstance(value, str) and "/" in value:
        a, b = value.split("/", 1)
        return a.strip(), b.strip()
    return None, None


def flatten_kb(kb: dict[str, Any]) -> dict[str, Any]:
    flat: dict[str, Any] = {}

    def put(key: str, value: Any) -> None:
        if value is None:
            return
        if isinstance(value, str) and not value.strip():
            return
        flat[key] = value
        flat[normalize_label(key)] = value

    def walk(obj: Any, prefix: str = "") -> None:
        if isinstance(obj, dict):
            for k, v in obj.items():
                key = f"{prefix}.{k}" if prefix else str(k)
                if isinstance(v, dict):
                    walk(v, key)
                else:
                    put(key, v)
                    put(str(k), v)
        else:
            put(prefix, obj)

    walk(kb)
    return flat


def lookup_kb_value(label_norm: str, kb_flat: dict[str, Any]) -> tuple[str | None, Any]:
    if label_norm in kb_flat:
        return label_norm, kb_flat[label_norm]
    for key in DEFAULT_ALIASES.get(label_norm, []):
        if key in kb_flat:
            return key, kb_flat[key]
        nk = normalize_label(key)
        if nk in kb_flat:
            return key, kb_flat[nk]
    for k, v in kb_flat.items():
        if normalize_label(str(k)) == label_norm:
            return str(k), v
    return None, None


def render_slot_text(slot: dict[str, Any], kb_flat: dict[str, Any]) -> tuple[str | None, str | None, Any]:
    kind = slot.get("fill_mode") or slot.get("kind") or "field"
    label_norm = slot.get("label_norm") or normalize_label(slot.get("label", ""))
    current = slot.get("current") or ""

    if kind == "checkbox":
        for key in DEFAULT_ALIASES.get(label_norm, []) + [label_norm]:
            if key not in kb_flat:
                continue
            raw = kb_flat[key]
            choice = coerce_bool(raw)
            if choice is None:
                continue
            return key, apply_checkbox_choice(current, choice), raw
        return None, None, None

    if kind == "inline_rank":
        rank = kb_flat.get("rank")
        total = kb_flat.get("rank_total")
        key = "rank"
        if rank is None or total is None:
            for alias_key in DEFAULT_ALIASES.get(label_norm, []) + ["score_rank", "rank_display"]:
                if alias_key in kb_flat:
                    rank, total = parse_rank_pair(kb_flat[alias_key])
                    key = alias_key
                    break
        if rank is None or total is None:
            return None, None, None
        return key, apply_inline_rank(current, rank, total), f"{rank}/{total}"

    if kind == "inline_courses":
        req = kb_flat.get("required_courses")
        passed = kb_flat.get("required_passed")
        if req is None or passed is None:
            return None, None, None
        return "required_courses", apply_inline_courses(current, req, passed), f"{req}/{passed}"

    kb_key, value = lookup_kb_value(label_norm, kb_flat)
    if kb_key is None or value is None:
        return None, None, None
    text = value if isinstance(value, str) else str(value)
    return kb_key, text, value


def render_slot_fill_items(slot: dict[str, Any], kb_flat: dict[str, Any]) -> tuple[str | None, list[dict[str, Any]], Any]:
    """Return (kb_key, cell_writes, raw_value). unknown → no writes."""
    fill_mode = slot.get("fill_mode") or slot.get("kind") or "field"
    if fill_mode == "unknown":
        return None, [], None

    label_norm = slot.get("label_norm") or normalize_label(slot.get("label", ""))
    table = int(slot.get("table") or 1)
    row = int(slot.get("row") or 0)
    kind = slot.get("kind") or fill_mode

    if fill_mode == "digit_grid" or kind == "digit_grid":
        kb_key, value = lookup_kb_value(label_norm, kb_flat)
        if kb_key is None or value is None:
            return None, [], None
        cols = slot.get("cols") or [t["col"] for t in slot.get("targets") or []]
        items = []
        for idx, ch in apply_digit_grid(value, cols):
            items.append(
                {
                    "table": table,
                    "row": row,
                    "col": cols[idx],
                    "text": ch,
                    "slot_id": slot.get("slot_id"),
                    "label": slot.get("label"),
                    "kind": "digit_grid",
                    "fill_mode": "digit_grid",
                }
            )
        return kb_key, items, value

    if fill_mode == "year_month_day" or kind == "year_month_day":
        kb_key, value = lookup_kb_value(label_norm, kb_flat)
        if kb_key is None or value is None:
            return None, [], None
        style = slot.get("date_style") or "ymd"
        text = apply_year_month_day(value, style)
        return kb_key, [
            {
                "table": table,
                "row": int(slot["row"]),
                "col": int(slot["col"]),
                "text": text,
                "slot_id": slot.get("slot_id"),
                "label": slot.get("label"),
                "kind": "year_month_day",
                "fill_mode": "year_month_day",
            }
        ], value

    kb_key, text, raw = render_slot_text(slot, kb_flat)
    if kb_key is None or text is None:
        return None, [], None
    return kb_key, [
        {
            "table": table,
            "row": int(slot["row"]),
            "col": int(slot["col"]),
            "text": text,
            "slot_id": slot.get("slot_id"),
            "label": slot.get("label"),
            "kind": kind,
            "fill_mode": fill_mode,
        }
    ], raw


def verify_fill(doc, cells: list[dict[str, Any]]) -> dict[str, Any]:
    """Re-read written cells and compare to planned text."""
    ok: list[dict[str, Any]] = []
    mismatch: list[dict[str, Any]] = []
    for item in cells:
        t_idx = int(item.get("table") or 1)
        r, c = int(item["row"]), int(item["col"])
        expected = str(item.get("text", ""))
        try:
            actual = _cell_text(doc.Tables(t_idx).Cell(r, c))
        except Exception as exc:
            mismatch.append({**item, "actual": None, "error": str(exc)})
            continue
        # digit cells: compare stripped first char / full
        if actual.strip() == expected.strip() or (expected and actual.strip().startswith(expected)):
            ok.append({"table": t_idx, "row": r, "col": c, "text": expected})
        else:
            mismatch.append({"table": t_idx, "row": r, "col": c, "expected": expected, "actual": actual})
    return {"verified": len(ok), "mismatch_count": len(mismatch), "mismatches": mismatch}
