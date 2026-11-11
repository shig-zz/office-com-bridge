from __future__ import annotations

import re
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

# WdConstants verified against the installed Word type library.
_WD_FIND_STOP = 0
_WD_REPLACE_ONE = 1
_WD_REPLACE_ALL = 2
_MAX_PARAGRAPH_WINDOW = 200

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


# ---------------------------------------------------------------------------
# Phase 3B: paragraphs, find and replace
#
# Word automation quirks drive these choices: Range.Sort-style Find runs are
# unreliable when driven through dynamic dispatch, so the search path scans the
# paragraphs directly (giving paragraph numbers and context for free), while
# the replace path drives Word's own Find.Execute so formatting and tracked
# changes behave as they do in the UI, then re-scans to verify.
# ---------------------------------------------------------------------------


def _word_error(message: str, exc: BaseException) -> BridgeError:
    hresult = getattr(exc, "hresult", None)
    details: dict[str, Any] = {"error": str(exc)}
    if hresult is not None:
        details["hresult"] = f"0x{int(hresult) & 0xFFFFFFFF:08X}"
    return BridgeError(ErrorCode.COM_ERROR, message, 502, details)


def _paragraph_text(paragraph: Any) -> str:
    """Paragraph.Range.Text keeps a trailing paragraph mark; drop it."""
    text = str(paragraph.Range.Text)
    if text.endswith("\r") or text.endswith("\x07"):
        text = text[:-1]
    return text


def _paragraph_style(paragraph: Any) -> str:
    try:
        style = paragraph.Style
        if isinstance(style, str):
            return style
        return str(style.NameLocal)
    except Exception:
        return ""


def _is_heading(style_name: str, paragraph: Any) -> bool:
    low = style_name.lower()
    if "heading" in low or "title" in low or "見出し" in style_name or "表題" in style_name:
        return True
    try:
        level = int(paragraph.OutlineLevel)  # wdOutlineLevel1..9 are headings, 10 is body
        return 1 <= level <= 9
    except Exception:
        return False


def _match_spans(
    text: str, needle: str, *, match_case: bool, whole_words: bool
) -> list[tuple[int, int]]:
    if whole_words:
        flags = re.UNICODE | (0 if match_case else re.IGNORECASE)
        pattern = re.compile(rf"(?<!\w){re.escape(needle)}(?!\w)", flags)
        return [(m.start(), m.end()) for m in pattern.finditer(text)]
    spans: list[tuple[int, int]] = []
    source = text if match_case else text.lower()
    target = needle if match_case else needle.lower()
    position = source.find(target)
    while position != -1:
        spans.append((position, position + len(target)))
        position = source.find(target, position + len(target))
    return spans


def _context(text: str, start: int, end: int, width: int = 60) -> str:
    left = max(0, start - width)
    right = min(len(text), end + width)
    prefix = "…" if left > 0 else ""
    suffix = "…" if right < len(text) else ""
    return prefix + text[left:right].replace("\r", " ").replace("\x07", " ") + suffix


def paragraphs(
    target: ResolvedTarget,
    *,
    start: Any,
    count: Any,
    max_text_chars: int,
) -> dict[str, Any]:
    """Read a window of paragraphs with style name and heading flag."""
    try:
        start_int = int(start)
        count_int = int(count)
    except (TypeError, ValueError) as exc:
        raise invalid_request(
            "Parameters 'start' and 'count' must be numbers.",
            {"start": start, "count": count},
        ) from exc
    if start_int < 1:
        raise invalid_request(
            "Parameter 'start' must be a 1-based paragraph number.", {"start": start}
        )
    if count_int < 1 or count_int > _MAX_PARAGRAPH_WINDOW:
        raise invalid_request(
            f"Parameter 'count' must be between 1 and {_MAX_PARAGRAPH_WINDOW}.",
            {"count": count},
        )
    doc = target.document
    try:
        total = int(doc.Paragraphs.Count)
    except pywintypes.com_error as exc:
        raise _word_error("The Word document paragraphs could not be read.", exc) from exc

    items: list[dict[str, Any]] = []
    truncated_any = False
    end = min(total, start_int + count_int - 1)
    try:
        for index in range(start_int, end + 1):
            paragraph = doc.Paragraphs(index)
            text = _paragraph_text(paragraph)
            style = _paragraph_style(paragraph)
            display, truncated = truncate_text(text, max_text_chars)
            truncated_any = truncated_any or truncated
            items.append(
                {
                    "index": index,
                    "characters": len(text),
                    "text": display,
                    "truncated": truncated,
                    "style": style,
                    "is_heading": _is_heading(style, paragraph),
                }
            )
    except pywintypes.com_error as exc:
        raise _word_error("The Word document paragraphs could not be read.", exc) from exc

    return {
        "application": "Word",
        "total_paragraphs": total,
        "start": start_int,
        "count": len(items),
        "truncated": truncated_any,
        "paragraphs": items,
    }


def find(
    target: ResolvedTarget,
    *,
    find_text: Any,
    match_case: bool,
    whole_words: bool,
    max_matches: int,
    max_text_chars: int,
) -> dict[str, Any]:
    """Search the body text and report each match with its paragraph number."""
    if not isinstance(find_text, str) or not find_text:
        raise invalid_request("Parameter 'find' must be a non-empty string.")
    doc = target.document
    try:
        total = int(doc.Paragraphs.Count)
    except pywintypes.com_error as exc:
        raise _word_error("The Word document paragraphs could not be read.", exc) from exc

    locations: list[dict[str, Any]] = []
    overflow = False
    scanned = 0
    try:
        for index in range(1, total + 1):
            scanned = index
            text = _paragraph_text(doc.Paragraphs(index))
            for begin, finish in _match_spans(
                text, find_text, match_case=match_case, whole_words=whole_words
            ):
                if len(locations) >= max_matches:
                    overflow = True
                    break
                locations.append(
                    {
                        "paragraph": index,
                        "context": truncate_text(
                            _context(text, begin, finish), max_text_chars
                        )[0],
                    }
                )
            if overflow:
                break
    except pywintypes.com_error as exc:
        raise _word_error("The Word document could not be searched.", exc) from exc

    return {
        "application": "Word",
        "find": find_text,
        "match_case": match_case,
        "whole_words": whole_words,
        "paragraphs_scanned": scanned,
        "matches": len(locations),
        "more_matches": overflow,
        "locations": locations,
    }


def _run_find(
    range_object: Any,
    find_text: str,
    replace_with: str,
    *,
    match_case: bool,
    whole_words: bool,
    replace_mode: int,
) -> bool:
    try:
        finder = range_object.Find
        finder.ClearFormatting()
        finder.Replacement.ClearFormatting()
        found = finder.Execute(
            find_text,
            bool(match_case),
            bool(whole_words),
            False,  # MatchWildcards
            False,  # MatchSoundsLike
            False,  # MatchAllWordForms
            True,   # Forward
            _WD_FIND_STOP,
            False,  # Format (search for formatting off)
            replace_with,
            replace_mode,
        )
    except pywintypes.com_error as exc:
        raise _word_error("The Word find/replace call failed.", exc) from exc
    return bool(found)


def _count_matches(
    doc: Any, needle: str, *, match_case: bool, whole_words: bool
) -> int:
    total = int(doc.Paragraphs.Count)
    count = 0
    for index in range(1, total + 1):
        text = _paragraph_text(doc.Paragraphs(index))
        count += len(
            _match_spans(text, needle, match_case=match_case, whole_words=whole_words)
        )
    return count


def replace(
    target: ResolvedTarget,
    *,
    find_text: Any,
    replacement: Any,
    match_case: bool,
    whole_words: bool,
    nth: Any,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Replace matches through Word's own Find, then re-scan to verify."""
    if not isinstance(find_text, str) or not find_text:
        raise invalid_request("Parameter 'find' must be a non-empty string.")
    if not isinstance(replacement, str):
        raise invalid_request("Parameter 'replace' must be a string (empty to delete).")
    check_expected_document(target, params)
    doc = target.document
    try:
        total = int(doc.Paragraphs.Count)
    except pywintypes.com_error as exc:
        raise _word_error("The Word document paragraphs could not be read.", exc) from exc

    occurrences: list[tuple[int, int, int]] = []
    try:
        for index in range(1, total + 1):
            text = _paragraph_text(doc.Paragraphs(index))
            for begin, finish in _match_spans(
                text, find_text, match_case=match_case, whole_words=whole_words
            ):
                occurrences.append((index, begin, finish))
    except pywintypes.com_error as exc:
        raise _word_error("The Word document could not be searched.", exc) from exc

    track_revisions = None
    try:
        track_revisions = bool(doc.TrackRevisions)
    except Exception:
        pass

    result: dict[str, Any] = {
        "application": "Word",
        "dry_run": dry_run,
        "replaced": False,
        "find": find_text,
        "match_case": match_case,
        "whole_words": whole_words,
        "occurrences_before": len(occurrences),
        "first_paragraph": occurrences[0][0] if occurrences else None,
        "track_revisions": track_revisions,
    }
    if track_revisions:
        result["note"] = (
            "Track Changes is on: replacements are recorded as tracked changes."
        )
    if nth is not None:
        if isinstance(nth, bool) or not isinstance(nth, int) or nth < 1:
            raise invalid_request(
                "Parameter 'nth' must be a 1-based match number.", {"nth": nth}
            )
        result["nth"] = nth
        if nth > len(occurrences):
            raise invalid_request(
                "Parameter 'nth' is beyond the number of matches found.",
                {"nth": nth, "occurrences": len(occurrences)},
            )
    if dry_run:
        return result
    require_confirm(params, "word.replace")

    if not occurrences:
        result["verified"] = True
        result["occurrences_after"] = 0
        return result

    try:
        if nth is not None:
            paragraph_index = occurrences[nth - 1][0]
            found = _run_find(
                doc.Paragraphs(paragraph_index).Range,
                find_text,
                replacement,
                match_case=match_case,
                whole_words=whole_words,
                replace_mode=_WD_REPLACE_ONE,
            )
            result["replaced"] = bool(found)
            remaining = _count_matches(
                doc, find_text, match_case=match_case, whole_words=whole_words
            )
            result["verified"] = remaining == len(occurrences) - 1
        else:
            _run_find(
                doc.Content,
                find_text,
                replacement,
                match_case=match_case,
                whole_words=whole_words,
                replace_mode=_WD_REPLACE_ALL,
            )
            remaining = _count_matches(
                doc, find_text, match_case=match_case, whole_words=whole_words
            )
            result["replaced"] = True
            result["verified"] = remaining == 0
        result["occurrences_after"] = remaining
    except pywintypes.com_error as exc:
        raise _word_error("The Word replace operation failed.", exc) from exc
    return result


def replace_paragraph_text(
    target: ResolvedTarget,
    *,
    paragraph: Any,
    text: Any,
    dry_run: bool,
    max_text_chars: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Replace one paragraph's body text, keeping the paragraph break."""
    if isinstance(paragraph, bool) or not isinstance(paragraph, int) or paragraph < 1:
        raise invalid_request(
            "Parameter 'paragraph' must be a 1-based paragraph number.",
            {"parameter": "paragraph"},
        )
    if not isinstance(text, str):
        raise invalid_request("Parameter 'text' must be a string.")
    if len(text) > max_text_chars:
        raise text_too_large(len(text), max_text_chars)
    check_expected_document(target, params)

    doc = target.document
    try:
        total = int(doc.Paragraphs.Count)
    except pywintypes.com_error as exc:
        raise _word_error("The Word document paragraphs could not be read.", exc) from exc
    if paragraph > total:
        raise invalid_request(
            "Parameter 'paragraph' is beyond the end of the document.",
            {"paragraph": paragraph, "paragraph_count": total},
        )
    try:
        before = _paragraph_text(doc.Paragraphs(paragraph))
    except pywintypes.com_error as exc:
        raise _word_error("The paragraph could not be read.", exc) from exc

    normalized = text.replace("\r\n", "\r").replace("\n", "\r")
    added_breaks = normalized.count("\r")

    result: dict[str, Any] = {
        "application": "Word",
        "dry_run": dry_run,
        "replaced": False,
        "paragraph": paragraph,
        "paragraph_count_before": total,
        "paragraph_count_after": total + added_breaks,
        "characters_before": len(before),
        "text_preview": truncate_text(before, max_text_chars)[0],
    }
    if dry_run:
        return result
    require_confirm(params, "word.paragraph.replace_text")

    try:
        # Paragraph.Range.Text includes the trailing paragraph mark; assigning
        # text without it would merge the paragraph into the next one, so the
        # replacement always ends with a paragraph break.
        doc.Paragraphs(paragraph).Range.Text = normalized + "\r"
        after_total = int(doc.Paragraphs.Count)
        after = _paragraph_text(doc.Paragraphs(paragraph))
    except pywintypes.com_error as exc:
        raise _word_error("The paragraph could not be changed.", exc) from exc

    first_line = normalized.split("\r")[0] if "\r" in normalized else normalized
    result.update(
        {
            "replaced": True,
            "paragraph_count_after": after_total,
            "characters_after": len(after),
            "verified": after == first_line,
            "text_preview": truncate_text(after, max_text_chars)[0],
        }
    )
    return result
