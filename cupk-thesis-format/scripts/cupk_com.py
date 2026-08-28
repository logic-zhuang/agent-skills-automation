# -*- coding: utf-8 -*-
"""Shared Word COM helpers for CUPK 套版. Do not rewrite PAGE fields."""
from __future__ import annotations

from pathlib import Path

import win32com.client as win32

wdCharacter = 1
wdAlignParagraphCenter = 1
wdHeaderFooterPrimary = 1
wdHeaderFooterEvenPages = 3
wdSectionBreakNextPage = 2
wdCollapseEnd = 0
CM = 28.35

ODD_HEADER = "中国石油大学（北京）本科毕业设计（论文）"


FILLABLE_NAME = "中国石油大学（北京）本科毕业设计（论文）标准模板（可填写）.docx"


def skill_root() -> Path:
    return Path(__file__).resolve().parents[1]


def fillable_template_path() -> Path:
    """Skill-bundled 可填写模板 (`templates/`). No runtime fallback outside skill."""
    path = skill_root() / "templates" / FILLABLE_NAME
    if not path.exists():
        raise FileNotFoundError(
            f"fillable template missing: {path}\n"
            "Copy the official 可填写 template into templates/, or run "
            "build_standard_template.py (see convert.md §模板维护)."
        )
    return path


def attach_word(*, visible: bool = True):
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = bool(visible)
    word.DisplayAlerts = 0
    try:
        word.ScreenUpdating = True
    except Exception:
        pass
    return word


def open_or_reuse(word, path: Path, *, read_only: bool = False):
    path = path.resolve()
    if path.name.startswith("~$"):
        raise ValueError(f"skip Word lock file: {path}")
    target = str(path).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            if str(Path(d.FullName).resolve()).lower() == target:
                return d, False
        except Exception:
            continue
    doc = word.Documents.Open(
        str(path),
        ConfirmConversions=False,
        ReadOnly=read_only,
        AddToRecentFiles=False,
    )
    return doc, True


def para_text(p) -> str:
    return (p.Range.Text or "").replace("\r", "").replace("\x07", "")


def set_para_text(p, text: str):
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    al = rng.ParagraphFormat.Alignment
    rng.Text = text
    rng.ParagraphFormat.Alignment = al


def apply_heading1(p):
    """范本「标题 1」样式 Alignment=3（两端），正文用法必须居中。"""
    doc = p.Range.Document
    try:
        p.Style = doc.Styles("标题 1")
    except Exception:
        pass
    p.Format.Alignment = wdAlignParagraphCenter
    p.Format.FirstLineIndent = 0
    try:
        p.Format.CharacterUnitFirstLineIndent = 0
    except Exception:
        pass


def center_all_heading1(doc) -> int:
    n = 0
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        try:
            st = str(p.Style.NameLocal)
        except Exception:
            continue
        if st not in ("标题 1", "Heading 1"):
            continue
        t = para_text(p).replace("\x0c", "").strip()
        if not t:
            continue
        apply_heading1(p)
        n += 1
    return n


def replace_cell_text_keep_format(cell, text: str):
    rng = cell.Range
    rng.MoveEnd(wdCharacter, -1)
    al = rng.ParagraphFormat.Alignment
    fe = rng.Font.NameFarEast
    ascii_name = rng.Font.NameAscii
    size = rng.Font.Size
    bold = rng.Font.Bold
    rng.Text = text
    rng.ParagraphFormat.Alignment = al
    try:
        rng.Font.NameFarEast = fe
        rng.Font.NameAscii = ascii_name
        rng.Font.Size = size
        rng.Font.Bold = bold
    except Exception:
        pass


def clear_after_para(sec, para_index: int):
    first_end = sec.Range.Paragraphs(para_index).Range.End
    end = sec.Range.End - 1
    if end > first_end:
        sec.Range.Document.Range(first_end, end).Delete()


def append_in_section(sec, text: str, style="正文"):
    doc = sec.Range.Document
    pos = sec.Range.End - 1
    rng = doc.Range(pos, pos)
    rng.InsertBefore(text + "\r")
    p = rng.Paragraphs(1)
    try:
        p.Style = doc.Styles(style)
    except Exception:
        pass
    if style in ("标题 1", "Heading 1"):
        apply_heading1(p)
    return p


def set_even_header_text(sec, text: str):
    """Only change even-page header text. Never touch footers / PAGE fields."""
    h = sec.Headers(wdHeaderFooterEvenPages)
    h.LinkToPrevious = False
    rng = h.Range
    rng.MoveEnd(wdCharacter, -1)
    fe = rng.Font.NameFarEast or "黑体"
    ascii_name = rng.Font.NameAscii or "Times New Roman"
    size = rng.Font.Size or 10.5
    rng.Text = text
    try:
        rng.Font.NameFarEast = fe
        rng.Font.NameAscii = ascii_name
        rng.Font.Name = fe
        rng.Font.Size = size
        rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
    except Exception:
        pass


def ensure_odd_header(sec, text: str = ODD_HEADER):
    h = sec.Headers(wdHeaderFooterPrimary)
    h.LinkToPrevious = False
    rng = h.Range
    rng.MoveEnd(wdCharacter, -1)
    cur = (rng.Text or "").replace("\r", "").strip()
    if cur == text:
        return
    if "硕士" in cur or not cur:
        fe = rng.Font.NameFarEast or "黑体"
        ascii_name = rng.Font.NameAscii or "Times New Roman"
        size = rng.Font.Size or 10.5
        rng.Text = text
        try:
            rng.Font.NameFarEast = fe
            rng.Font.NameAscii = ascii_name
            rng.Font.Name = fe
            rng.Font.Size = size
            rng.ParagraphFormat.Alignment = wdAlignParagraphCenter
        except Exception:
            pass


def unlink_headers_keep_footers(sec):
    for idx in (wdHeaderFooterPrimary, wdHeaderFooterEvenPages):
        try:
            sec.Headers(idx).LinkToPrevious = False
        except Exception:
            pass
        try:
            f = sec.Footers(idx)
            f.LinkToPrevious = False
            try:
                f.PageNumbers.RestartNumberingAtSection = False
            except Exception:
                pass
        except Exception:
            pass


def apply_keywords_zh_format(p):
    """「关键词：」黑体小三加粗；其后黑体小四加粗。"""
    doc = p.Range.Document
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    try:
        rng.Font.NameFarEast = "黑体"
        rng.Font.NameAscii = "黑体"
    except Exception:
        pass
    rng.Font.Bold = True
    p.Format.FirstLineIndent = 0
    try:
        p.Format.CharacterUnitFirstLineIndent = 0
    except Exception:
        pass
    start = p.Range.Start
    lab = doc.Range(start, start + 4)
    lab.Font.Size = 15
    lab.Font.Bold = True
    rest_end = p.Range.End - 1
    if rest_end > start + 4:
        rest = doc.Range(start + 4, rest_end)
        rest.Font.Size = 12
        rest.Font.Bold = True
        try:
            rest.Font.NameFarEast = "黑体"
        except Exception:
            pass


def apply_keywords_en_format(p):
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    try:
        rng.Font.NameFarEast = "Times New Roman"
        rng.Font.NameAscii = "Times New Roman"
        rng.Font.Name = "Times New Roman"
    except Exception:
        pass
    rng.Font.Size = 12
    rng.Font.Bold = True
    p.Format.FirstLineIndent = 0
    try:
        p.Format.CharacterUnitFirstLineIndent = 0
    except Exception:
        pass
