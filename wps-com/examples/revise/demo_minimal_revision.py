# -*- coding: utf-8 -*-
"""Demo minimal revision via WPS: polish paragraph 2 with TrackRevisions."""
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
    sys.path.insert(0, str(ROOT / "scripts"))
    try:
        from wps_session import WPS_PROGIDS, find_open_doc

        import win32com.client as win32

        wps = None
        for progid in WPS_PROGIDS:
            try:
                wps = win32.GetActiveObject(progid)
                break
            except Exception:
                continue
        if wps is None:
            return
        doc = find_open_doc(wps, ESSAY)
        if doc is not None:
            doc.Close(False)
    except Exception:
        pass


def main() -> int:
    close_essay_if_open()
    if ESSAY.exists():
        try:
            ESSAY.unlink()
        except PermissionError:
            print("error: short_essay.docx 被占用，请先关闭 WPS 中的该文件", file=sys.stderr)
            return 2

    subprocess.check_call([sys.executable, str(CASE / "create_short_essay.py")])

    sys.path.insert(0, str(ROOT / "scripts"))
    from wps_session import attach_wps, backend_progid, open_doc, para_text

    wps = attach_wps(visible=True)
    print("backend", backend_progid())
    doc = open_doc(wps, ESSAY)
    try:
        wps.ActiveWindow.View.ShowRevisionsAndComments = False
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

    doc = open_doc(wps, ESSAY)
    try:
        wps.ActiveWindow.View.RevisionsView = 0
        wps.ActiveWindow.View.ShowRevisionsAndComments = False
    except Exception:
        pass
    print("after p2 (final view):", para_text(doc.Paragraphs(2)))
    print("revisions:", doc.Revisions.Count)
    try:
        wps.ActiveWindow.View.ShowRevisionsAndComments = True
        wps.ActiveWindow.View.RevisionsView = 0
    except Exception:
        pass
    print("请在 WPS「审阅 → 所有标记」中查看删除线/下划线；勿自动 AcceptAll。")
    print("essay:", ESSAY)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
