"""Per-control screenshots via UIA CaptureToImage (no OCR)."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def capture_control(ctrl, dest: Path) -> dict[str, Any]:
    """Capture one UIA control to ``dest``. Returns metadata."""
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    meta: dict[str, Any] = {
        "image": str(dest),
        "captured": False,
        "bbox": {},
        "error": "",
    }
    try:
        rect = ctrl.BoundingRectangle
        w = int(rect.right - rect.left)
        h = int(rect.bottom - rect.top)
        meta["bbox"] = {
            "left": int(rect.left),
            "top": int(rect.top),
            "right": int(rect.right),
            "bottom": int(rect.bottom),
            "width": w,
            "height": h,
        }
        if w < 2 or h < 2:
            meta["error"] = "bbox_too_small"
            return meta
        ok = bool(ctrl.CaptureToImage(str(dest)))
        meta["captured"] = ok and dest.exists() and dest.stat().st_size > 0
        if not meta["captured"]:
            meta["error"] = "capture_failed"
    except Exception as exc:
        meta["error"] = str(exc)
    return meta
