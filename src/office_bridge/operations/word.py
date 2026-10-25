from __future__ import annotations

from typing import Any

import pywintypes

from ..errors import BridgeError, ErrorCode, invalid_request, text_too_large
from ..target import ResolvedTarget
from .common import (
    check_expected_document,
    optional_bool,
    require_confirm,
    truncate_text,
)

_SELECTION_TYPES = {
    0: "NoSelection",
    1: "InsertionPoint",
    2: "Normal",
    3: "Frame",
    4: "Column",
    5: "Row",
    6: "Block",
    7: "InlineShape",
    8: "Shape",
}


def _selection_error(exc: BaseException) -> BridgeError:
    hresult = getattr(exc, "hresult", None)
    details: dict[str, Any] = {"error": str(exc)}
    if hresult is not None:
        details["hresult"] = f"0x{int(hresult) & 0xFFFFFFFF:08X}"
    return BridgeError(
        ErrorCode.COM_ERROR,
        "The Word selection could not be accessed.",
        502,
        details,
    )


def _selection(target: ResolvedTarget) -> Any:
    try:
        return target.window.Selection
    except pywintypes.com_error as exc:
        raise _selection_error(exc) from exc


def selection(target: ResolvedTarget, *, max_text_chars: int) -> dict[str, Any]:
    selected = _selection(target)
    try:
        text = str(selected.Text)
        start = int(selected.Start)
        end = int(selected.End)
        type_number = int(selected.Type)
    except pywintypes.com_error as exc:
        raise _selection_error(exc) from exc
    except Exception as exc:
        raise _selection_error(exc) from exc

    full_length = len(text)
    text, truncated = truncate_text(text, max_text_chars)
    return {
        "application": "Word",
        "selection_type": _SELECTION_TYPES.get(type_number, f"Unknown({type_number})"),
        "selection_type_code": type_number,
        "start": start,
        "end": end,
        "characters": full_length,
        "text": text,
        "truncated": truncated,
    }


def document(target: ResolvedTarget) -> dict[str, Any]:
    doc = target.document
    info: dict[str, Any] = {
        "application": "Word",
        "name": target.info.get("document", ""),
        "full_name": target.info.get("full_name", ""),
        "path": target.info.get("path", ""),
    }
    counters = {
        "paragraphs": lambda: int(doc.Paragraphs.Count),
        "words": lambda: int(doc.Words.Count),
        "characters": lambda: int(doc.Characters.Count),
    }
    for key, getter in counters.items():
        try:
            info[key] = getter()
        except Exception:
            info[key] = None
    for key, getter in {
        "saved": lambda: bool(doc.Saved),
        "read_only": lambda: bool(doc.ReadOnly),
    }.items():
        try:
            info[key] = getter()
        except Exception:
            info[key] = None
    return info


def insert_text(
    target: ResolvedTarget,
    *,
    text: Any,
    replace: bool,
    dry_run: bool,
    max_text_chars: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    if not isinstance(text, str) or not text:
        raise invalid_request("Parameter 'text' must be a non-empty string.")
    if len(text) > max_text_chars:
        raise text_too_large(len(text), max_text_chars)
    check_expected_document(target, params)

    selected = _selection(target)
    try:
        start = int(selected.Start)
        end = int(selected.End)
    except Exception as exc:
        raise _selection_error(exc) from exc

    if not dry_run:
        require_confirm(params, "word.selection.insert_text")

    result = {
        "application": "Word",
        "dry_run": dry_run,
        "inserted": False,
        "replace": replace,
        "start": start,
        "end": end,
        "characters": len(text),
    }
    if dry_run:
        return result

    try:
        if not replace:
            selected.Collapse(0)
        selected.TypeText(text)
    except pywintypes.com_error as exc:
        raise _selection_error(exc) from exc

    try:
        result["selection_start"] = int(selected.Start)
        result["selection_end"] = int(selected.End)
    except Exception:
        pass
    result["inserted"] = True
    return result
