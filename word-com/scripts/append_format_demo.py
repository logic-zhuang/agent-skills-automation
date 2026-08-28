# -*- coding: utf-8 -*-
"""Append more formats and OMath formulas to the existing COM demo doc.

Reuses the already-open Word document (no second Open of the same path).
Leaves the Word window open.
"""
from __future__ import annotations

from pathlib import Path

import win32com.client as win32

wdCharacter = 1
wdCollapseEnd = 0
wdAlignLeft = 0
wdAlignCenter = 1
wdAlignJustify = 3
wdSectionBreakNextPage = 2
wdYellow = 6  # wdBrightGreen=4; 7=yellow
wdTurquoise = 3
wdLineStyleSingle = 1
wdLineWidth150pt = 12
wdLineWidth050pt = 4
wdBorderTop = -1
wdBorderLeft = -2
wdBorderBottom = -3
wdBorderRight = -4
wdBorderHorizontal = -5
wdBorderVertical = -6
wdUndefined = 9999999

OUT = Path(__file__).resolve().parents[4] / "example" / "word_com_格式演示.docx"


def rgb(r: int, g: int, b: int) -> int:
    return r + (g << 8) + (b << 16)


def attach_word():
    try:
        word = win32.GetActiveObject("Word.Application")
    except Exception:
        word = win32.DispatchEx("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0
    try:
        word.ScreenUpdating = True
    except Exception:
        pass
    return word


def get_doc(word):
    target = str(OUT.resolve()).lower()
    for i in range(1, word.Documents.Count + 1):
        d = word.Documents(i)
        try:
            full = str(Path(d.FullName).resolve()).lower()
        except Exception:
            continue
        if full == target or d.Name.endswith("格式演示.docx"):
            return d
    if not OUT.exists():
        raise FileNotFoundError(OUT)
    return word.Documents.Open(str(OUT), ReadOnly=False)


def add_para(doc, text: str):
    end = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end.InsertAfter(text + "\r")
    for i in range(doc.Paragraphs.Count, 0, -1):
        cand = doc.Paragraphs(i)
        ct = (cand.Range.Text or "").replace("\r", "").replace("\x07", "")
        if ct == text:
            return cand
    return doc.Paragraphs(doc.Paragraphs.Count - 1)


def body_of(p):
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    return rng


def style_heading(p, size=16, color=None):
    try:
        p.Style = p.Range.Document.Styles("标题 2")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.NameAscii = "Arial"
    r.Font.Size = size
    r.Font.Bold = True
    if color is not None:
        r.Font.Color = color
    p.Format.SpaceBefore = 12
    p.Format.SpaceAfter = 6


def insert_eq(doc, linear: str, number: str):
    p = add_para(doc, linear)
    r = body_of(p)
    r.Font.NameAscii = "Cambria Math"
    r.Font.Size = 12
    p.Format.Alignment = wdAlignCenter
    p.Format.SpaceBefore = 6
    p.Format.FirstLineIndent = 0
    try:
        p.Format.CharacterUnitFirstLineIndent = 0
    except Exception:
        pass
    try:
        om = doc.OMaths.Add(r)
        try:
            om.BuildUp()
        except Exception:
            pass
    except Exception as e:
        print("OMath add fail", number, e)
    n = add_para(doc, number)
    nr = body_of(n)
    nr.Font.NameAscii = "Times New Roman"
    nr.Font.Size = 11
    n.Format.Alignment = wdAlignCenter
    n.Format.SpaceAfter = 8
    return p


def main():
    word = attach_word()
    doc = get_doc(word)
    try:
        word.Activate()
        doc.Activate()
    except Exception:
        pass

    # one page break so extra content continues without a half-empty page
    end = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end.InsertBreak(7)  # wdPageBreak

    p = add_para(doc, "增补页：更多字符格式与数学公式")
    try:
        p.Style = doc.Styles("标题 1")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 18
    r.Font.Color = rgb(128, 0, 0)
    p.Format.Alignment = wdAlignCenter

    # --- character formats ---
    p = add_para(doc, "五、更多字符与段落格式")
    style_heading(p, color=rgb(0, 82, 147))

    p = add_para(doc, "删除线：这段文字带删除线，表示修订删除。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    r.Font.StrikeThrough = True
    r.Font.Color = rgb(128, 128, 128)

    p = add_para(doc, "双删除线 + 阴文：DoubleStrike 演示。")
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 12
    try:
        r.Font.DoubleStrikeThrough = True
    except Exception:
        r.Font.StrikeThrough = True

    p = add_para(doc, "字符间距加宽 2 磅，小型大写英文 SMALL CAPS Demo。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.NameAscii = "Times New Roman"
    r.Font.Size = 12
    try:
        r.Font.Spacing = 2
        r.Font.SmallCaps = True
    except Exception:
        pass

    p = add_para(doc, "上标与下标：水的分子式为 H2O，爱因斯坦质能关系涉及 c2。")
    # apply sub/super on digits via Find within paragraph
    rng = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    f = rng.Find
    f.ClearFormatting()
    if f.Execute(FindText="H2O", Forward=True, Wrap=0):
        # range now is H2O; shrink to the 2
        two = doc.Range(rng.Start + 1, rng.Start + 2)
        two.Font.Subscript = True
        two.Font.Size = 12
    rng2 = body_of(p)
    f2 = rng2.Find
    f2.ClearFormatting()
    if f2.Execute(FindText="c2", Forward=True, Wrap=0):
        two = doc.Range(rng2.Start + 1, rng2.Start + 2)
        two.Font.Superscript = True

    p = add_para(doc, "段落底纹：本段浅蓝底，用于提示或摘录。")
    r = body_of(p)
    r.Font.NameFarEast = "楷体"
    r.Font.Size = 12
    r.Font.Color = rgb(0, 51, 102)
    try:
        p.Range.Shading.BackgroundPatternColor = rgb(218, 232, 252)
    except Exception:
        pass
    p.Format.LeftIndent = 14
    p.Format.RightIndent = 14
    p.Format.SpaceBefore = 6
    p.Format.SpaceAfter = 6

    p = add_para(doc, "边框段落：四周边框 + 段前段后间距。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    try:
        for bidx in (wdBorderTop, wdBorderBottom, wdBorderLeft, wdBorderRight):
            b = p.Range.Borders(bidx)
            b.LineStyle = wdLineStyleSingle
            b.LineWidth = wdLineWidth050pt
            b.Color = rgb(0, 82, 147)
    except Exception:
        pass

    p = add_para(doc, "项目编号列表第一项。")
    try:
        p.Range.ListFormat.ApplyNumberDefault()
    except Exception:
        pass
    p = add_para(doc, "项目编号列表第二项。")
    try:
        p.Range.ListFormat.ApplyNumberDefault()
    except Exception:
        pass
    p = add_para(doc, "项目符号列表。")
    try:
        p.Range.ListFormat.ApplyBulletDefault()
    except Exception:
        pass

    p = add_para(doc, "超链接文本指向 Microsoft 文档。")
    r = body_of(p)
    try:
        doc.Hyperlinks.Add(Anchor=r, Address="https://learn.microsoft.com/office/vba/api/overview/word")
    except Exception as e:
        print("hyperlink skip", e)

    # comment on last heading-ish para
    try:
        c = doc.Comments.Add(Range=body_of(p), Text="这是 Word Comments 批注对象，不是封面文本框。")
        c.Author = "Word COM Demo"
    except Exception as e:
        print("comment skip", e)

    # --- three-line table ---
    p = add_para(doc, "六、三线表示例")
    style_heading(p, color=rgb(0, 82, 147))
    p = add_para(doc, "表 1  常见希腊字母（三线表：顶线、栏头线、底线）")
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 10.5
    p.Format.Alignment = wdAlignCenter

    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    table = doc.Tables.Add(rng, 4, 3)
    data = [
        ["符号", "名称", "常见用途"],
        ["α, β, γ", "alpha, beta, gamma", "角度 / 系数"],
        ["σ, τ", "sigma, tau", "应力 / 时间常数"],
        ["ω, θ", "omega, theta", "角频率 / 角度"],
    ]
    for i, row in enumerate(data, 1):
        for c, val in enumerate(row, 1):
            cell = table.Cell(i, c)
            cr = cell.Range
            cr.MoveEnd(wdCharacter, -1)
            cr.Text = val
            cr.Font.NameFarEast = "宋体" if i > 1 else "黑体"
            cr.Font.NameAscii = "Times New Roman"
            cr.Font.Size = 10.5
            cr.Font.Bold = i == 1
            cell.Range.ParagraphFormat.Alignment = wdAlignCenter
            cell.VerticalAlignment = 1
    # three-line: clear all, then top/header-bottom/bottom
    try:
        table.Borders.Enable = False
    except Exception:
        pass
    try:
        table.Rows(1).Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        table.Rows(1).Borders(wdBorderTop).LineWidth = wdLineWidth150pt
        table.Rows(1).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        table.Rows(1).Borders(wdBorderBottom).LineWidth = wdLineWidth050pt
        table.Rows(table.Rows.Count).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        table.Rows(table.Rows.Count).Borders(wdBorderBottom).LineWidth = wdLineWidth150pt
    except Exception as e:
        print("three-line border skip", e)
    table.Rows.Alignment = 1
    add_para(doc, "")

    # text box shape
    try:
        # left/top in points from page; keep moderate
        sh = doc.Shapes.AddTextbox(1, 120, 40, 280, 48)  # msoTextOrientationHorizontal=1
        sh.TextFrame.TextRange.Text = "浮动文本框：Shapes 对象，可拖动。"
        sh.TextFrame.TextRange.Font.NameFarEast = "楷体"
        sh.TextFrame.TextRange.Font.Size = 10.5
        sh.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignCenter
        sh.Fill.ForeColor.RGB = rgb(255, 242, 204)
        sh.Line.ForeColor.RGB = rgb(191, 143, 0)
    except Exception as e:
        print("textbox skip", e)

    # --- formulas ---
    p = add_para(doc, "七、更多数学公式（OMath）")
    style_heading(p, color=rgb(128, 0, 0))
    p = add_para(doc, "下列公式使用 Word 公式对象；编号在公式外单独成段。若仍为线性文本，可在公式上右键转为「专业」。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 10.5
    r.Font.Color = rgb(89, 89, 89)
    p.Format.CharacterUnitFirstLineIndent = 2

    formulas = [
        (r"a/b+c/d=(ad+bc)/bd", "(2)"),
        (r"\sqrt(a^2+b^2)", "(3)"),
        (r"\sum_(i=1)^n i=n(n+1)/2", "(4)"),
        (r"\int_0^1 x^2 dx=1/3", "(5)"),
        (r"lim_(x\to0) sin(x)/x=1", "(6)"),
        (r"e^(i\pi)+1=0", "(7)"),
        (r"\partial u/\partial t=k \partial^2 u/\partial x^2", "(8)"),
        (r"F=G (m_1 m_2)/r^2", "(9)"),
        (r"\alpha+\beta+\gamma=\pi", "(10)"),
        (r"\vec(a)·\vec(b)=|a||b|cos\theta", "(11)"),
        (r"P(A|B)=P(B|A)P(A)/P(B)", "(12)"),
        (r"x_{1,2}=(-b\pm\sqrt(b^2-4ac))/(2a)", "(13)"),
        (r"\oint_C \vec(F)·d\vec(l)=0", "(14)"),
        (r"\bar(x)=1/n \sum_(i=1)^n x_i", "(15)"),
        (r"A=\pmatrix(a&b@c&d)", "(16)"),
    ]
    for linear, num in formulas:
        insert_eq(doc, linear, num)

    # BuildUp any leftover linear omaths
    try:
        for i in range(1, doc.OMaths.Count + 1):
            try:
                doc.OMaths(i).BuildUp()
            except Exception:
                pass
    except Exception:
        pass

    doc.Fields.Update()
    try:
        doc.Save()
    except Exception:
        doc.SaveAs2(str(OUT), FileFormat=16)
    print("updated", OUT)
    print(
        "paragraphs",
        doc.Paragraphs.Count,
        "tables",
        doc.Tables.Count,
        "omath",
        doc.OMaths.Count,
        "comments",
        doc.Comments.Count,
        "hyperlinks",
        doc.Hyperlinks.Count,
        "shapes",
        doc.Shapes.Count,
        "sections",
        doc.Sections.Count,
    )
    print("Word window left open.")


if __name__ == "__main__":
    main()
