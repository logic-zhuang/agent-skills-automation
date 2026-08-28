"""Microsoft Word COM bridge.

所有修改都通过 Word 的对象模型完成，因此开启 TrackRevisions 后会生成真正的
Revision，而不是用颜色或删除线模拟的视觉效果。
"""

from __future__ import annotations

from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Any


class WordBridgeError(RuntimeError):
    """Word 未运行、没有活动文档或 COM 操作失败。"""


@dataclass(frozen=True, slots=True)
class RevisionInfo:
    """Word Revision 的可序列化摘要。"""

    revision_type: int
    text: str
    author: str
    date: str


class WordBridge:
    """连接当前用户已打开的 Word 实例，并提供原生修订操作。"""

    def __init__(self, word: Any, document: Any) -> None:
        self.word = word
        self.document = document

    @classmethod
    def connect_active(cls) -> "WordBridge":
        """连接当前活动 Word；不会启动新 Word，也不会关闭 Word。"""
        try:
            import win32com.client
        except ImportError as exc:
            raise WordBridgeError("未安装 pywin32，请先执行 uv sync。") from exc

        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except Exception as exc:  # COM 的具体异常类型随 Office 版本变化
            raise WordBridgeError("找不到正在运行的 Word，请先打开 Word。") from exc

        try:
            document = word.ActiveDocument
            if document is None:
                raise WordBridgeError("Word 当前没有活动文档。")
        except WordBridgeError:
            raise
        except Exception as exc:
            raise WordBridgeError("无法读取 Word 的活动文档。") from exc
        return cls(word, document)

    @property
    def track_revisions(self) -> bool:
        """当前文档是否正在记录修订。"""
        return bool(self.document.TrackRevisions)

    @track_revisions.setter
    def track_revisions(self, enabled: bool) -> None:
        self.document.TrackRevisions = bool(enabled)

    def enable_tracking(self) -> None:
        """开启 Word 原生修订记录。"""
        self.track_revisions = True

    def get_document_text(self) -> str:
        """读取文档正文（包含 Word 的段落结束符）。"""
        return str(self.document.Content.Text)

    def get_selection_text(self) -> str:
        """读取当前 Word 选择区域的文本。"""
        return str(self.word.Selection.Text)

    def find_text(self, text: str) -> Any | None:
        """查找第一次精确匹配，返回独立 Range；找不到返回 None。"""
        if not text:
            raise ValueError("text 不能为空。")
        # Duplicate 避免 Find.Execute 改变文档 Content 这个共享 Range。
        search_range = self.document.Content.Duplicate
        find = search_range.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Text = text
        find.Forward = True
        find.Wrap = 0  # wdFindStop
        find.MatchCase = True
        find.MatchWholeWord = False
        return search_range if find.Execute() else None

    def replace_text(
        self, old: str, new: str, *, all_occurrences: bool = False
    ) -> int:
        """用新文本替换匹配内容，并让 Word 记录删除/插入修订。

        返回实际替换次数。默认只替换第一次，避免 AI 误改同文档中的其他同文。
        """
        if not old:
            raise ValueError("old 不能为空。")
        self.enable_tracking()
        if all_occurrences:
            count = 0
            search_start = 0
            while True:
                match = self._find_text_in_range(old, search_start)
                if match is None:
                    return count
                self._replace_range_minimal(match, old, new)
                count += 1
                # 从新内容之后继续，避免 new 包含 old 时再次命中刚插入的文本。
                search_start = match.End
        match = self.find_text(old)
        if match is None:
            return 0
        self._replace_range_minimal(match, old, new)
        return 1

    def _replace_range_minimal(self, match: Any, old: str, new: str) -> None:
        """只修改 old/new 的差异部分，让 Word 产生更细粒度的修订。"""
        matcher = SequenceMatcher(None, old, new, autojunk=False)
        document_start = match.Start

        # 必须从后往前改，否则前面的插入/删除会改变后续 COM 位置。
        for tag, old_start, old_end, new_start, new_end in reversed(
            matcher.get_opcodes()
        ):
            if tag == "equal":
                continue
            target = self.document.Range(
                document_start + old_start,
                document_start + old_end,
            )
            replacement = new[new_start:new_end]
            if tag == "delete":
                target.Delete()
            elif tag == "insert":
                target.InsertBefore(replacement)
            else:  # replace：只替换这一小段不同文本
                target.Text = replacement

    def replace_text_in_paragraph(
        self, paragraph_number: int, old: str, new: str
    ) -> int:
        """只在指定段落中替换一次文本，并产生 Word 原生修订。

        paragraph_number 使用 Word 用户习惯的 1-based 编号：第 1 段传 1，
        第 2 段传 2。这样可以避免全文中出现相同句子时误改其他位置。
        """
        if paragraph_number < 1:
            raise ValueError("paragraph_number 必须从 1 开始。")
        if not old:
            raise ValueError("old 不能为空。")
        try:
            paragraph = self.document.Paragraphs.Item(paragraph_number)
            paragraph_range = paragraph.Range.Duplicate
        except Exception as exc:
            raise IndexError(f"找不到第 {paragraph_number} 段。") from exc

        find = paragraph_range.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Text = old
        find.Forward = True
        find.Wrap = 0  # wdFindStop
        find.MatchCase = True
        find.MatchWholeWord = False
        if not find.Execute():
            return 0
        self.enable_tracking()
        self._replace_range_minimal(paragraph_range, old, new)
        return 1

    def replace_document_text(
        self, new_text: str, *, expected_old_text: str | None = None
    ) -> None:
        """用全文最小差异更新文档，并保留 Word 原生修订。

        new_text 可以来自大模型的整篇重写结果。方法会保留文档末尾段落标记，
        并把换行统一为 Word 使用的回车符。传入 expected_old_text 可防止在
        Word 内容已被其他人修改后，误将旧的 AI 结果写回文档。
        """
        content = self.document.Content.Duplicate
        # Word 文档正文末尾有一个不可删除的段落标记，不交给模型参与比较。
        target = self.document.Range(content.Start, content.End - 1)
        old_text = str(target.Text)
        normalized_new = new_text.replace("\r\n", "\r").replace("\n", "\r")
        if expected_old_text is not None:
            normalized_expected = expected_old_text.replace(
                "\r\n", "\r"
            ).replace("\n", "\r")
            if normalized_expected.rstrip("\r") != old_text.rstrip("\r"):
                raise WordBridgeError(
                    "Word 当前正文已变化，与 AI 修改前的原文不一致，已停止更新。"
                )
        self.enable_tracking()
        self._replace_range_minimal(target, old_text, normalized_new)

    def _find_text_in_range(self, text: str, start: int) -> Any | None:
        """在文档的 [start, 末尾] 范围内查找，供全部替换使用。"""
        search_range = self.document.Range(start, self.document.Content.End)
        find = search_range.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.Text = text
        find.Forward = True
        find.Wrap = 0  # wdFindStop
        find.MatchCase = True
        find.MatchWholeWord = False
        return search_range if find.Execute() else None

    def insert_text(self, text: str, position: int | None = None) -> None:
        """在字符位置插入文本；position 省略时插入当前选择位置。"""
        self.enable_tracking()
        target = (
            self.word.Selection.Range
            if position is None
            else self.document.Range(position, position)
        )
        target.InsertBefore(text)

    def delete_text(self, text: str) -> int:
        """删除第一次匹配，并让 Word 记录删除修订。"""
        match = self.find_text(text)
        if match is None:
            return 0
        self.enable_tracking()
        match.Delete()
        return 1

    def set_style(self, text: str, style_name: str) -> int:
        """将第一次匹配的文本应用 Word 样式，返回是否成功应用。"""
        if not style_name:
            raise ValueError("style_name 不能为空。")
        match = self.find_text(text)
        if match is None:
            return 0
        self.enable_tracking()
        match.Style = style_name
        return 1

    def set_font(
        self,
        text: str,
        *,
        name: str | None = None,
        size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        color: int | None = None,
    ) -> int:
        """设置第一次匹配文本的指定字体属性，未提供的属性保持不变。"""
        match = self.find_text(text)
        if match is None:
            return 0
        self.enable_tracking()
        font = match.Font
        if name is not None:
            font.Name = name
        if size is not None:
            font.Size = size
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if color is not None:
            font.Color = color
        return 1

    def get_revisions(self) -> list[RevisionInfo]:
        """返回当前文档修订的快照，便于交给 Agent/UI。"""
        return [
            RevisionInfo(
                revision_type=int(revision.Type),
                text=str(revision.Range.Text),
                author=str(revision.Author),
                date=str(revision.Date),
            )
            for revision in self.document.Revisions
        ]

    def accept_all_revisions(self) -> None:
        """接受当前文档全部修订。"""
        self.document.Revisions.AcceptAll()

    def reject_all_revisions(self) -> None:
        """拒绝当前文档全部修订。"""
        self.document.Revisions.RejectAll()

    def accept_revision(self, index: int) -> None:
        """按当前修订快照的零基索引接受一条修订。"""
        self._revision_at(index).Accept()

    def reject_revision(self, index: int) -> None:
        """按当前修订快照的零基索引拒绝一条修订。"""
        self._revision_at(index).Reject()

    def _revision_at(self, index: int) -> Any:
        if index < 0 or index >= self.document.Revisions.Count:
            raise IndexError(f"修订索引越界：{index}")
        # COM 集合是 1-based，Python API 对外统一使用 0-based。
        return self.document.Revisions.Item(index + 1)
