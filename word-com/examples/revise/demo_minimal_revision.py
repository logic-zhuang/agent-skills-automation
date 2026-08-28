# -*- coding: utf-8 -*-
"""Demo minimal revision: polish paragraph 2 with TrackRevisions (visible marks).

Always recreates a clean short_essay.docx, then applies one polish job.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

CASE = Path(__file__).resolve().parent
ROOT = CASE.parents[1]
ESSAY = CASE / "short_essay.docx"
JOB = CASE / "revise_job.json"

NEW_P2 = "不少学校只能通过线上平台完成作业与考试，教师也只好借助屏幕点评学生作文。"


def close_essay_if_open() -> None:
    try:
        import win32com.client as win32

        word = win32.GetActiveObject("Word.Application")
        for i in range(word.Documents.Count, 0, -1):
            d = word.Documents(i)
            try:
                if "short_essay" in str(d.Name).lower():
                    d.Close(False)
            except Exception:
                pass
    except Exception:
        pass


def main() -> int:
    close_essay_if_open()
    if ESSAY.exists():
        try:
            ESSAY.unlink()
        except PermissionError:
            print("error: short_essay.docx 被占用，请先关闭 Word 中的该文件", file=sys.stderr)
            return 2

    subprocess.check_call([sys.executable, str(CASE / "create_short_essay.py")])

    sys.path.insert(0, str(ROOT / "scripts"))
    from word_session import attach_word, open_doc, para_text

    word = attach_word(visible=True)
    doc = open_doc(word, ESSAY)
    # Original view without mixing delete+insert glyphs in Range.Text
    try:
        word.ActiveWindow.View.ShowRevisionsAndComments = False
    except Exception:
        pass
    old = para_text(doc.Paragraphs(2))
    print("before p2:", old)

    JOB.write_text(
        json.dumps(
            {
                "paragraph": 2,
                "old": old,
                "new": NEW_P2,
                "expected_old": old,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    subprocess.check_call(
        [
            sys.executable,
            str(ROOT / "scripts" / "minimal_revision.py"),
            str(ESSAY),
            "--job",
            str(JOB),
        ]
    )

    doc = open_doc(word, ESSAY)
    try:
        word.ActiveWindow.View.RevisionsView = 0  # final
        word.ActiveWindow.View.ShowRevisionsAndComments = False
    except Exception:
        pass
    print("after p2 (final view):", para_text(doc.Paragraphs(2)))
    print("revisions:", doc.Revisions.Count)
    try:
        word.ActiveWindow.View.ShowRevisionsAndComments = True
        word.ActiveWindow.View.RevisionsView = 0
    except Exception:
        pass
    print("请在 Word「审阅 → 所有标记」中查看删除线/下划线；勿自动 AcceptAll。")
    print("essay:", ESSAY)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
