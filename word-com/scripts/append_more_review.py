# -*- coding: utf-8 -*-
"""Add more review / footnote / citation / hyperlink demos. Reuse open doc. No Quit."""
from __future__ import annotations

from pathlib import Path

import win32com.client as win32

wdCharacter = 1
wdCollapseStart = 1
wdCollapseEnd = 0
wdAlignCenter = 1
wdSectionBreakNextPage = 2
wdFieldEmpty = -1
wdFieldRef = 3
wdFieldXE = 4
wdFieldNoteRef = 72
wdFieldStyleRef = 48
wdFieldCitation = 96
wdFieldBibliography = 97
wdFieldIndex = 8
wdCaptionPositionAbove = 0
wdCaptionPositionBelow = 1
wdYellow = 7
wdTurquoise = 3
wdBrightGreen = 4
wdRefTypeFootnote = 3
wdRefTypeHeading = 1
wdRefTypeBookmark = 2
wdContentText = 2
wdPageNumber = 7
wdFootnoteNumber = 4

OUT = Path(__file__).resolve().parents[4] / "example" / "word_com_格式演示.docx"


def rgb(r, g, b):
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
        if full == target or "格式演示" in d.Name:
            return d
    if not OUT.exists():
        raise FileNotFoundError(OUT)
    return word.Documents.Open(str(OUT), ReadOnly=False)


def add_para(doc, text: str):
    end = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    end.InsertAfter(text + "\r")
    for i in range(doc.Paragraphs.Count, 0, -1):
        p = doc.Paragraphs(i)
        t = p.Range.Text.replace("\r", "").replace("\x07", "")
        if t == text or (text and t.startswith(text[:20])):
            return p
        if not text and t == "":
            return p
    return doc.Paragraphs.Last


def body_of(p):
    rng = p.Range
    rng.MoveEnd(wdCharacter, -1)
    return rng


def heading(p, color=None):
    try:
        p.Style = p.Range.Document.Styles("标题 2")
        try:
            p.Range.Document.Styles("标题 2").NextParagraphStyle = p.Range.Document.Styles("正文")
        except Exception:
            pass
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.NameAscii = "Times New Roman"
    r.Font.Size = 14
    r.Font.Bold = True
    if color is not None:
        r.Font.Color = color
    p.Format.CharacterUnitFirstLineIndent = 0
    p.Format.SpaceBefore = 12
    p.Format.SpaceAfter = 6


def add_field(doc, rng, ftype, text=None):
    if text is None:
        return doc.Fields.Add(rng, ftype)
    return doc.Fields.Add(rng, ftype, text)


def delete_from_marker(doc, marker: str):
    rng = doc.Content
    f = rng.Find
    f.ClearFormatting()
    if f.Execute(FindText=marker, Forward=True, Wrap=0):
        start = rng.Start
        end = doc.Content.End - 1
        if end > start:
            doc.Range(start, end).Delete()
            print("deleted from", marker)


def main():
    word = attach_word()
    doc = get_doc(word)
    doc.Activate()

    delete_from_marker(doc, "CITATION_TEST")

    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    rng.InsertBreak(7)  # wdPageBreak; do not use next-page section (blank pages)

    p = add_para(doc, "增补页：审阅、脚注、引用、超链接（续）")
    try:
        p.Style = doc.Styles("标题 1")
    except Exception:
        pass
    r = body_of(p)
    r.Font.NameFarEast = "黑体"
    r.Font.Size = 16
    r.Font.Bold = True
    p.Format.Alignment = wdAlignCenter
    p.Format.CharacterUnitFirstLineIndent = 0
    try:
        doc.Bookmarks.Add("MoreReview", body_of(p))
    except Exception:
        pass

    p = add_para(
        doc,
        "本页补全：批注回复、修订高亮、更多脚注/尾注、NOTEREF、题注、CITATION/BIBLIOGRAPHY、带 ScreenTip 的超链接、STYLEREF、索引。",
    )
    r = body_of(p)
    r.Font.NameFarEast = "楷体"
    r.Font.Size = 10.5
    r.Font.Color = rgb(89, 89, 89)
    p.Format.CharacterUnitFirstLineIndent = 2

    # ---------- 十六 审阅 ----------
    p = add_para(doc, "十六、审阅加深：回复、高亮、再插修订")
    heading(p, rgb(192, 0, 0))

    p = add_para(doc, "本段同时带批注和高亮，便于对照审阅窗格与正文标记。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    r.HighlightColorIndex = wdYellow
    try:
        c = doc.Comments.Add(Range=r, Text="黄色高亮表示待核实。")
        c.Author = "审阅人丙"
        c.Initial = "丙"
        try:
            reply = c.Replies.Add(Range=c.Range, Text="已核实，数据来自第 3 章表 1。")
            try:
                reply.Author = "作者"
                reply.Initial = "作"
            except Exception:
                pass
            print("comment reply ok, replies", c.Replies.Count)
        except Exception as e:
            print("comment reply skip", e)
            p2 = add_para(doc, "（本机 Word 的 Comment.Replies.Add 不可用，已跳过回复。）")
            body_of(p2).Font.Size = 10.5
            body_of(p2).Font.Color = rgb(128, 128, 128)
    except Exception as e:
        print("comment3 skip", e)

    p = add_para(doc, "关键词「待删句」将在修订模式下删除；「待插入」将显示为插入修订。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2

    p = add_para(doc, "保留句。待删句。保留句。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    try:
        doc.TrackRevisions = True
        f = body_of(p).Find
        f.ClearFormatting()
        if f.Execute(FindText="待删句。", Forward=True, Wrap=0):
            f.Parent.Delete()
        ins = body_of(p)
        ins.Collapse(wdCollapseEnd)
        ins.InsertAfter("待插入：此处为修订插入文本。")
        print("revisions after extra edit", doc.Revisions.Count)
    except Exception as e:
        print("track extra skip", e)
    finally:
        try:
            doc.TrackRevisions = False
        except Exception:
            pass

    p = add_para(doc, "青绿高亮表示已采纳意见；亮绿表示作者自检。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    f = r.Find
    f.ClearFormatting()
    if f.Execute(FindText="青绿高亮", Forward=True, Wrap=0):
        f.Parent.HighlightColorIndex = wdTurquoise
    f = body_of(p).Find
    f.ClearFormatting()
    if f.Execute(FindText="亮绿", Forward=True, Wrap=0):
        f.Parent.HighlightColorIndex = wdBrightGreen

    # ---------- 十七 脚注 ----------
    p = add_para(doc, "十七、更多脚注、自定义标记与交叉引用")
    heading(p)

    p = add_para(doc, "第三脚注跟在句末。¹ 符号由 Word 自动编号。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    mark = body_of(p)
    mark.Collapse(wdCollapseEnd)
    try:
        fn3 = doc.Footnotes.Add(Range=mark, Text="第三条脚注：可含对方法或数据来源的补充说明。")
        try:
            doc.Bookmarks.Add("fnThird", fn3.Reference)
        except Exception:
            pass
        print("footnotes", doc.Footnotes.Count)
    except Exception as e:
        print("fn3 skip", e)

    p = add_para(doc, "自定义标记脚注（星号）常用于标题或作者说明。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    mark = body_of(p)
    mark.Collapse(wdCollapseEnd)
    try:
        doc.Footnotes.Add(Range=mark, Text="自定义标记脚注：Reference 使用 *。", Reference="*")
        print("custom footnote ok")
    except Exception as e:
        print("custom fn skip", e)

    p = add_para(doc, "第四条自动编号脚注，便于后面做 NOTEREF。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    mark = body_of(p)
    mark.Collapse(wdCollapseEnd)
    try:
        fn4 = doc.Footnotes.Add(Range=mark, Text="第四条脚注：交叉引用会指向本标记。")
        doc.Bookmarks.Add("fnFourth", fn4.Reference)
        print("fn4 bookmarked")
    except Exception as e:
        print("fn4 skip", e)

    p = add_para(doc, "正文交叉引用脚注：参见脚注 ")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        fld = add_field(doc, ins, wdFieldNoteRef)
        fld.Code.Text = " NOTEREF fnFourth \\h "
        fld.Update()
        print("NOTEREF ok", repr(fld.Result.Text))
    except Exception as e:
        print("NOTEREF field skip", e)
        ins = body_of(p)
        ins.Collapse(wdCollapseEnd)
        try:
            ins.InsertCrossReference(
                ReferenceType=wdRefTypeFootnote,
                ReferenceKind=wdFootnoteNumber,
                ReferenceItem="1",
                InsertAsHyperlink=True,
            )
            print("InsertCrossReference footnote ok")
        except Exception as e2:
            print("xref footnote skip", e2)

    p = add_para(doc, "再插一条尾注，与脚注分区显示。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    mark = body_of(p)
    mark.Collapse(wdCollapseEnd)
    try:
        doc.Endnotes.Add(Range=mark, Text="第二条尾注：Endnotes 集合与 Footnotes 分开计数。")
        print("endnotes", doc.Endnotes.Count)
    except Exception as e:
        print("endnote2 skip", e)

    # ---------- 十八 题注 ----------
    p = add_para(doc, "十八、题注（Figure/Table）与 REF 交叉引用")
    heading(p)

    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    table = doc.Tables.Add(rng, 3, 3)
    data = (("项目", "数值", "备注"), ("样本 A", "12.3", "对照组"), ("样本 B", "15.8", "实验组"))
    for i, row in enumerate(data, 1):
        for c, val in enumerate(row, 1):
            cell = table.Cell(i, c)
            cr = cell.Range
            cr.MoveEnd(wdCharacter, -1)
            cr.Text = val
            cr.Font.Size = 10.5
            cr.Font.NameFarEast = "宋体"
            cell.Range.ParagraphFormat.Alignment = wdAlignCenter
            if i == 1:
                cr.Font.Bold = True
    try:
        table.Borders.Enable = True
    except Exception:
        pass
    table.Rows.Alignment = 1

    try:
        cap = table.Range
        cap.Collapse(wdCollapseStart)
        cap.InsertCaption(
            Label="Table",
            Title="  审阅演示用三列表",
            TitleAutoText="",
            Position=wdCaptionPositionAbove,
        )
        print("table caption ok")
    except Exception as e:
        print("table caption skip", e)

    p = add_para(doc, "（示意图占位，下面加 Figure 题注）")
    r = body_of(p)
    r.Font.NameFarEast = "楷体"
    r.Font.Size = 10.5
    r.Font.Color = rgb(89, 89, 89)
    p.Format.Alignment = wdAlignCenter
    p.Format.CharacterUnitFirstLineIndent = 0
    try:
        cap = body_of(p)
        cap.Collapse(wdCollapseEnd)
        cap.InsertParagraphAfter()
        cap2 = cap.Next()
        cap2.InsertCaption(
            Label="Figure",
            Title="  示意图题注（InsertCaption Label=Figure）",
            TitleAutoText="",
            Position=wdCaptionPositionBelow,
        )
        print("figure caption ok")
    except Exception as e:
        print("figure caption skip", e)

    p = add_para(doc, "交叉引用书签 MoreReview：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        fld = add_field(doc, ins, wdFieldRef)
        fld.Code.Text = " REF MoreReview \\h "
        fld.Update()
        print("REF MoreReview", repr(fld.Result.Text[:40]))
    except Exception as e:
        print("REF skip", e)

    p = add_para(doc, "STYLEREF 取当前「标题 1」文本：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        fld = add_field(doc, ins, wdFieldStyleRef)
        fld.Code.Text = ' STYLEREF "标题 1" '
        fld.Update()
        print("STYLEREF", repr(fld.Result.Text[:40]))
    except Exception as e:
        print("STYLEREF skip", e)

    # ---------- 十九 引用 ----------
    p = add_para(doc, "十九、文献引用域（CITATION=96，BIBLIOGRAPHY=97）")
    heading(p)

    p = add_para(doc, "正文里插入两条引文：先书后刊。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        f1 = add_field(doc, ins, wdFieldCitation, "Li2021")
        f1.Update()
        print("citation Li2021", repr(f1.Result.Text))
    except Exception as e:
        print("cit Li skip", e)
        ins = body_of(p)
        ins.Collapse(wdCollapseEnd)
        try:
            fld = add_field(doc, ins, wdFieldEmpty)
            fld.Code.Text = " CITATION Li2021 \\l 2052 "
            fld.Update()
        except Exception as e2:
            print("cit empty skip", e2)

    p = add_para(doc, "期刊引文：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        f2 = add_field(doc, ins, wdFieldCitation, "Wang2022")
        f2.Update()
        print("citation Wang2022", repr(f2.Result.Text))
    except Exception as e:
        print("cit Wang skip", e)

    p = add_para(doc, "参考文献表（BIBLIOGRAPHY 域）")
    try:
        p.Style = doc.Styles("标题 3")
    except Exception:
        pass
    p = add_para(doc, "")
    ins = body_of(p)
    try:
        try:
            doc.Bibliography.BibliographyStyle = "GB7714"
        except Exception:
            pass
        fb = add_field(doc, ins, wdFieldBibliography)
        fb.Update()
        print("biblio ok, sources", doc.Bibliography.Sources.Count, "result chars", len(fb.Result.Text or ""))
    except Exception as e:
        print("biblio skip", e)

    # ---------- 二十 超链接 ----------
    p = add_para(doc, "二十、超链接：网址、邮件、书签、ScreenTip")
    heading(p, rgb(0, 100, 0))

    p = add_para(doc, "带屏幕提示的网址：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        doc.Hyperlinks.Add(
            Anchor=ins,
            Address="https://learn.microsoft.com/office/vba/api/word.comment",
            TextToDisplay="Comment 对象文档",
            ScreenTip="悬停可见：Word VBA Comment",
        )
        print("hl screentip ok")
    except Exception as e:
        print("hl screentip skip", e)

    p = add_para(doc, "脚注对象文档：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        doc.Hyperlinks.Add(
            Anchor=ins,
            Address="https://learn.microsoft.com/office/vba/api/word.footnotes",
            TextToDisplay="Footnotes 集合",
            ScreenTip="打开 Footnotes API",
        )
    except Exception as e:
        print("hl fn skip", e)

    p = add_para(doc, "书目概念文档：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        doc.Hyperlinks.Add(
            Anchor=ins,
            Address="https://learn.microsoft.com/office/vba/word/concepts/working-with-word/working-with-bibliographies",
            TextToDisplay="Working with Bibliographies",
            ScreenTip="CITATION / BIBLIOGRAPHY 域说明",
        )
    except Exception as e:
        print("hl bib skip", e)

    p = add_para(doc, "邮件：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        doc.Hyperlinks.Add(
            Anchor=ins,
            Address="mailto:editor@example.edu?subject=Thesis%20review",
            TextToDisplay="editor@example.edu",
            ScreenTip="打开默认邮件客户端",
        )
    except Exception as e:
        print("hl mail skip", e)

    p = add_para(doc, "文档内跳转：")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    ins = body_of(p)
    ins.Collapse(wdCollapseEnd)
    try:
        doc.Hyperlinks.Add(
            Anchor=ins,
            Address="",
            SubAddress="MoreReview",
            TextToDisplay="回到本页标题",
            ScreenTip="书签 MoreReview",
        )
    except Exception as e:
        print("hl more skip", e)

    if doc.Bookmarks.Exists("SecReview"):
        p = add_para(doc, "上一增补页：")
        ins = body_of(p)
        ins.Collapse(wdCollapseEnd)
        try:
            doc.Hyperlinks.Add(
                Anchor=ins,
                Address="",
                SubAddress="SecReview",
                TextToDisplay="跳到「审阅、脚注、引用与导航」",
                ScreenTip="书签 SecReview",
            )
        except Exception as e:
            print("hl secreview skip", e)

    if doc.Bookmarks.Exists("DemoStart"):
        p = add_para(doc, "文首：")
        ins = body_of(p)
        ins.Collapse(wdCollapseEnd)
        try:
            doc.Hyperlinks.Add(
                Anchor=ins,
                Address="",
                SubAddress="DemoStart",
                TextToDisplay="跳到演示文档开头",
            )
        except Exception:
            pass

    # ---------- 二十一 索引 ----------
    p = add_para(doc, "二十一、索引项（XE）与 INDEX 域")
    heading(p)
    p = add_para(doc, "下面给「审阅」「脚注」「引文」做索引项，文末生成 INDEX。")
    r = body_of(p)
    r.Font.NameFarEast = "宋体"
    r.Font.Size = 12
    p.Format.CharacterUnitFirstLineIndent = 2

    for word_xe in ("审阅", "脚注", "引文", "超链接"):
        ins = body_of(p)
        ins.Collapse(wdCollapseEnd)
        try:
            fld = add_field(doc, ins, wdFieldXE)
            fld.Code.Text = ' XE "%s" ' % word_xe
        except Exception as e:
            print("XE skip", word_xe, e)

    p = add_para(doc, "索引：")
    try:
        p.Style = doc.Styles("标题 3")
    except Exception:
        pass
    p = add_para(doc, "")
    ins = body_of(p)
    try:
        fld = add_field(doc, ins, wdFieldIndex)
        fld.Code.Text = ' INDEX \\z "2052" '
        fld.Update()
        print("INDEX ok")
    except Exception as e:
        print("INDEX skip", e)

    try:
        doc.Fields.Update()
    except Exception:
        pass
    try:
        if doc.TablesOfContents.Count:
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
    except Exception:
        pass

    try:
        doc.Save()
    except Exception:
        doc.SaveAs2(str(OUT), FileFormat=16)

    try:
        if doc.Bookmarks.Exists("MoreReview"):
            doc.Bookmarks("MoreReview").Range.Select()
            word.ActiveWindow.ScrollIntoView(doc.Bookmarks("MoreReview").Range, True)
    except Exception:
        pass

    print(
        "saved",
        OUT,
        "sections",
        doc.Sections.Count,
        "comments",
        doc.Comments.Count,
        "revisions",
        doc.Revisions.Count,
        "footnotes",
        doc.Footnotes.Count,
        "endnotes",
        doc.Endnotes.Count,
        "hyperlinks",
        doc.Hyperlinks.Count,
        "fields",
        doc.Fields.Count,
        "sources",
        doc.Bibliography.Sources.Count,
        "tables",
        doc.Tables.Count,
    )
    print("Word window left open.")


if __name__ == "__main__":
    main()
