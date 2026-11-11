"""Undo the most recent Office action (application.Undo).

Undo is the safe side of the write path: it rolls work back instead of adding
new changes. It still reverts whatever the user did last, so the same
confirm/expected_document rules as every other write apply, and dry-run always
reports whether anything is undoable first.
"""

from __future__ import annotations

from typing import Any

import pywintypes

from ..errors import BridgeError, ErrorCode
from ..target import ResolvedTarget
from .common import check_expected_document, require_confirm, unsaved_changes


def _undo_error(message: str, exc: BaseException) -> BridgeError:
    hresult = getattr(exc, "hresult", None)
    details: dict[str, Any] = {"error": str(exc)}
    if hresult is not None:
        details["hresult"] = f"0x{int(hresult) & 0xFFFFFFFF:08X}"
    return BridgeError(ErrorCode.COM_ERROR, message, 502, details)


def _undoable(target: ResolvedTarget) -> bool | None:
    """application.Undoable, or None when the application does not expose it."""
    try:
        return bool(target.application.Undoable)
    except Exception:
        return None


def undo(
    target: ResolvedTarget,
    *,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    check_expected_document(target, params)
    undoable_before = _undoable(target)

    result: dict[str, Any] = {
        "operation": "undo",
        "dry_run": dry_run,
        "executed": False,
        "undoable": undoable_before,
    }
    if dry_run:
        if undoable_before is False:
            result["note"] = "There is nothing to undo right now."
        return result

    require_confirm(params, "undo")
    if undoable_before is False:
        result["note"] = "There is nothing to undo; nothing changed."
        return result

    try:
        target.application.Undo()
    except pywintypes.com_error as exc:
        raise _undo_error("The undo call failed.", exc) from exc
    except Exception as exc:
        raise _undo_error("The undo call failed.", exc) from exc

    result.update(
        {
            "executed": True,
            "undoable_after": _undoable(target),
            "unsaved_changes": unsaved_changes(target),
        }
    )
    return result
