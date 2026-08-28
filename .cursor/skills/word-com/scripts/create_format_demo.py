# -*- coding: utf-8 -*-
"""Create a visible Word demo covering fonts, colors, tables, headers, formula."""
from __future__ import annotations

from pathlib import Path

import win32com.client as win32

# enums
wdCharacter = 1
wdCollapseEnd = 0
wdAlignLeft = 0
wdAlignCenter = 1
wdAlignRight = 2
wdAlignJustify = 3
wdLineSpaceMultiple = 5
wdLineSpace1pt5 = 1
wdSectionBreakNextPage = 2
wdHeaderFooterPrimary = 1
wdFieldPage = 33
wdYellow = 7
wdColorBlue = 16711680  # BGR
wdColorRed = 255
wdColorGreen = 32768
wdColorDarkBlue = 8388608
wdLineStyleSingle = 1
wdLineWidth050pt = 4
wdBorderTop = -1
wdBorderLeft = -2
wdBorderBottom = -3
wdBorderRight = -4
wdBorderHorizontal = -5
wdBorderVertical = -6
wdColorGray25 = 12632256

# scripts/ → word-com → skills → .cursor → repo root
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


def add_para(doc, text: str):
    end = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end.InsertAfter(text + "\r")
    p = doc.Paragraphs(doc.Paragraphs.Count - 1)
    got = (p.Range.Text or "").replace("\r", "").replace("\x07", "")
    if got != text:
        for i in range(doc.Paragraphs.Count, 0, -1):
            cand = doc.Paragraphs(i)
            ct = (cand.Range.Text or "").replace("\r", "").replace("\x07", "")
            if ct == text:
                return cand
    return p


def body_of(p):
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    return rng


def main():
    word = attach_word()
    doc = word.Documents.Add()
    try:
        word.Activate()
        doc.Activate()
    except Exception:
        pass

    # page
    ps = doc.Sections(1).PageSetup
    ps.TopMargin = 72
    ps.BottomMargin = 72
    ps.LeftMargin = 72
    ps.RightMargin = 72

    # header / footer
    h = doc.Sections(1).Headers(wdHeaderFooterPrimary)
    h.Range.Text = "Word COM 格式演示"
    h.Range.Font.NameFarEast = "黑体"
    h.Range.Font.NameAscii = "Arial"
    h.Range.Font.Size = 10.5
    h.Range.ParagraphFormat.Alignment = wdAlignCenter
    ftr = doc.Sections(1).Footers(wdHeaderFooterPrimary)
    ftr.Range.Text = ""
    ftr.Range.ParagraphFormat.Alignment = wdAlignCenter
    ftr.Range.Font.NameAscii = "Times New Roman"
    ftr.Range.Font.Size = 10.5
    rng = ftr.Range
    rng.InsertAfter("第 ")
    rng.Collapse(wdCollapseEnd)
    rng.Fields.Add(rng, wdFieldPage)
    rng.Collapse(wdCollapseEnd)
    rng.InsertAfter(" 页")

    # title
    p = add_para(doc, "Word COM 格式能力演示")
    try:
        p.Style = doc.Styles("标题 1")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.NameAscii = "Arial"
    r.Font.Size = 22
    r.Font.Bold = True
    r.Font.Color = rgb(0, 82, 147)
    p.Format.Alignment = wdAlignCenter
    p.Format.SpaceAfter = 12

    p = add_para(doc, "本文档由 win32com 新建（Documents.Add），覆盖字体、颜色、段落、表格、页眉页脚与公式。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.NameAscii = "Times New Roman"
    r.Font.Size = 12
    p.Format.Alignment = wdAlignJustify
    p.Format.CharacterUnitFirstLineIndent = 2
    p.Format.LineSpacingRule = wdLineSpaceMultiple
    p.Format.LineSpacing = 15

    # heading
    p = add_para(doc, "一、字体与颜色")
    try:
        p.Style = doc.Styles("标题 2")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 16
    r.Font.Bold = True
    r.Font.Color = rgb(192, 0, 0)
    p.Format.SpaceBefore = 12

    samples = [
        ("宋体小四，黑色正文。", "宋体", "Times New Roman", 12, False, False, rgb(0, 0, 0), False),
        ("黑体四号加粗，深蓝。", "黑体", "Arial", 14, True, False, rgb(0, 51, 102), False),
        ("楷体小四倾斜，绿色。", "楷体", "Times New Roman", 12, False, True, rgb(0, 128, 0), False),
        ("仿宋三号，紫色。", "仿宋", "Times New Roman", 16, False, False, rgb(112, 48, 160), False),
        ("Times New Roman 14pt Bold Red English.", "宋体", "Times New Roman", 14, True, False, wdColorRed, False),
        ("Arial 12pt 下划线 + 黄色底纹。", "宋体", "Arial", 12, False, False, rgb(0, 0, 0), True),
    ]
    for text, fe, asc, size, bold, italic, color, hl in samples:
        p = add_para(doc, text)
        r = body_of(p)
        try:
            r.Font.NameFarEast = fe
        except Exception:
            r.Font.NameFarEast = "宋体"
        r.Font.NameAscii = asc
        r.Font.Size = size
        r.Font.Bold = bold
        r.Font.Italic = italic
        r.Font.Color = color
        if hl:
            r.Underline = 1
            r.HighlightColorIndex = wdYellow
        p.Format.SpaceAfter = 6

    p = add_para(doc, "对齐：本段左对齐。")
    p.Format.Alignment = wdAlignLeft
    p = add_para(doc, "对齐：本段居中。")
    p.Format.Alignment = wdAlignCenter
    p = add_para(doc, "对齐：本段右对齐。")
    p.Format.Alignment = wdAlignRight
    p = add_para(doc, "对齐：本段两端对齐。科研论文排版时正文常用此对齐方式，使左右边缘整齐。")
    p.Format.Alignment = wdAlignJustify
    p.Format.CharacterUnitFirstLineIndent = 2

    # table section
    p = add_para(doc, "二、表格")
    try:
        p.Style = doc.Styles("标题 2")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 16
    r.Font.Bold = True
    r.Font.Color = rgb(0, 82, 147)

    p = add_para(doc, "下表演示表头底纹、单元格对齐与边框。")
    p.Format.CharacterUnitFirstLineIndent = 2

    # insert table after last content para
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    table = doc.Tables.Add(rng, 4, 4)
    headers = ["项目", "中文字体", "西文字体", "字号"]
    rows = [
        ["标题", "黑体", "Arial", "小三 15"],
        ["正文", "宋体", "Times New Roman", "小四 12"],
        ["表内", "宋体", "Arial", "五号 10.5"],
    ]
    for c, htext in enumerate(headers, 1):
        cell = table.Cell(1, c)
        cr = cell.Range
        cr.MoveEnd(wdCharacter, -1)
        cr.Text = htext
        cr.Font.NameFarEast = "黑体"
        cr.Font.NameAscii = "Arial"
        cr.Font.Size = 10.5
        cr.Font.Bold = True
        cr.Font.Color = rgb(255, 255, 255)
        cell.Shading.BackgroundPatternColor = rgb(0, 82, 147)
        cell.Range.ParagraphFormat.Alignment = wdAlignCenter
        cell.VerticalAlignment = 1  # center
    for i, row in enumerate(rows, 2):
        for c, val in enumerate(row, 1):
            cell = table.Cell(i, c)
            cr = cell.Range
            cr.MoveEnd(wdCharacter, -1)
            cr.Text = val
            cr.Font.NameFarEast = "宋体"
            cr.Font.NameAscii = "Times New Roman"
            cr.Font.Size = 10.5
            cell.Range.ParagraphFormat.Alignment = wdAlignCenter
            if i % 2 == 0:
                cell.Shading.BackgroundPatternColor = rgb(217, 226, 243)
    for idx in (wdBorderTop, wdBorderBottom, wdBorderLeft, wdBorderRight, wdBorderHorizontal, wdBorderVertical):
        try:
            b = table.Borders(idx)
            b.LineStyle = wdLineStyleSingle
            b.LineWidth = wdLineWidth050pt
            b.Color = rgb(0, 82, 147)
        except Exception:
            pass
    table.Rows.Alignment = 1  # center table

    add_para(doc, "")

    # list-like
    p = add_para(doc, "三、列表与强调")
    try:
        p.Style = doc.Styles("标题 2")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 16
    r.Font.Bold = True

    for t in ("（1）Range 改字，避免 Selection。", "（2）单元格先 MoveEnd 再写 Text。", "（3）中英字体拆开 NameFarEast / NameAscii。"):
        p = add_para(doc, t)
        r = body_of(p)
        r.Font.NameFarEast = "宋体"
        r.Font.Size = 12
        p.Format.LeftIndent = 21
        p.Format.SpaceAfter = 3

    # formula
    p = add_para(doc, "四、公式（OMath）")
    try:
        p.Style = doc.Styles("标题 2")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 16
    r.Font.Bold = True

    p = add_para(doc, "E=mc^2")
    r = body_of(p)
    try:
        om = doc.OMaths.Add(r)
        om.BuildUp()
    except Exception as e:
        print("OMath BuildUp skip:", e)
    p.Format.Alignment = wdAlignCenter
    p.Format.SpaceBefore = 6
    p.Format.SpaceAfter = 0

    p = add_para(doc, "(1)")
    r = body_of(p)
    r.Font.NameAscii = "Times New Roman"
    r.Font.Size = 12
    p.Format.Alignment = wdAlignCenter
    p.Format.SpaceAfter = 12

    # one page break (not next-page section — short sections leave large blank areas)
    end = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end.InsertBreak(7)  # wdPageBreak

    p = add_para(doc, "第二节：分页后的页眉仍属同一节")
    try:
        p.Style = doc.Styles("标题 1")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 18
    r.Font.Color = rgb(112, 48, 160)
    p.Format.Alignment = wdAlignCenter

    p = add_para(doc, "本页用分页符 wdPageBreak=7 另起一页；不分节则页眉与上一页相同。需要不同页眉时才用分节符下一页。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2

    p = add_para(doc, "上标示例：x² + y² = z²（部分为普通文本）。下标示例：H2O 中的数字可用 Font.Subscript。")
    # apply subscript on 2 in H2O if present
    rng = body_of(p)
    find = rng.Find
    find.ClearFormatting()
    if find.Execute(FindText="H2O", Forward=True, Wrap=0):
        # highlight the 2
        pass

    try:
        doc.Bookmarks.Add("DemoStart", doc.Paragraphs(1).Range)
    except Exception:
        pass

    doc.Fields.Update()
    OUT.parent.mkdir(parents=True, exist_ok=True)
    if OUT.exists():
        try:
            OUT.unlink()
        except Exception:
            pass
    doc.SaveAs2(str(OUT), FileFormat=16)
    print("saved", OUT)
    print("Word window left open. paragraphs", doc.Paragraphs.Count, "tables", doc.Tables.Count, "omath", doc.OMaths.Count)


if __name__ == "__main__":
    main()
