from __future__ import annotations

import re
from typing import Any

import pywintypes

from ..errors import (
    BridgeError,
    ErrorCode,
    autofit_refused,
    invalid_request,
    text_frame_missing,
)
from ..target import ResolvedTarget
from .common import (
    check_expected_document,
    jsonable,
    require_confirm,
    truncate_text,
)

_UNDO_HINT = (
    "roll this back with the 'undo' operation only after the user approves; "
    "it reverts the most recent Office action"
)


def _slide_error(message: str, exc: BaseException) -> BridgeError:
    hresult = getattr(exc, "hresult", None)
    details: dict[str, Any] = {"error": str(exc)}
    if hresult is not None:
        details["hresult"] = f"0x{int(hresult) & 0xFFFFFFFF:08X}"
    return BridgeError(ErrorCode.COM_ERROR, message, 502, details)


def _current_slide(target: ResolvedTarget) -> Any:
    if target.window is None:
        raise BridgeError(
            ErrorCode.NO_ACTIVE_SLIDE,
            "The PowerPoint presentation has no document window.",
            409,
        )
    try:
        return target.window.View.Slide
    except pywintypes.com_error as exc:
        raise BridgeError(
            ErrorCode.NO_ACTIVE_SLIDE,
            "PowerPoint is not showing a slide view for the selected presentation.",
            409,
            {"error": str(exc)},
        ) from exc


def _shape_texts(shape: Any) -> str:
    parts: list[str] = []
    try:
        if bool(shape.HasTextFrame) and bool(shape.TextFrame.HasText):
            text = str(shape.TextFrame.TextRange.Text)
            if text:
                parts.append(text)
    except Exception:
        pass
    try:
        if bool(shape.HasTable):
            table = shape.Table
            for row in range(1, int(table.Rows.Count) + 1):
                for column in range(1, int(table.Columns.Count) + 1):
                    cell = table.Cell(row, column)
                    try:
                        if bool(cell.Shape.HasTextFrame) and bool(cell.Shape.TextFrame.HasText):
                            text = str(cell.Shape.TextFrame.TextRange.Text)
                            if text:
                                parts.append(text)
                    except Exception:
                        continue
    except Exception:
        pass
    return "\n".join(parts)


def _slide_title(slide: Any) -> str:
    try:
        title = slide.Shapes.Title
        if title is None:
            return ""
        if bool(title.HasTextFrame) and bool(title.TextFrame.HasText):
            return str(title.TextFrame.TextRange.Text)
    except Exception:
        pass
    return ""


def _slide_body_texts(slide: Any) -> str:
    """Join the text of every non-title shape on a slide.

    The table-of-contents op uses this as the pool to search for sub-topics, and
    the title shape is left out on purpose: a title that is a substring of another
    title ("概要" inside "製品概要") says nothing about hierarchy, while a title
    listed in another slide's body is the section-slide pattern worth acting on.
    """
    parts: list[str] = []
    try:
        title = slide.Shapes.Title
    except Exception:
        title = None
    title_id = getattr(title, "Id", None) if title is not None else None
    try:
        count = int(slide.Shapes.Count)
    except Exception:
        count = 0
    for index in range(1, count + 1):
        try:
            shape = slide.Shapes(index)
        except Exception:
            continue
        if title_id is not None and getattr(shape, "Id", None) == title_id:
            continue
        text = _shape_texts(shape)
        if text:
            parts.append(text)
    return "\n".join(parts)


def _slide_result(target: ResolvedTarget, slide: Any, max_text_chars: int) -> dict[str, Any]:
    presentation = target.document
    try:
        slide_count = int(presentation.Slides.Count)
    except Exception:
        slide_count = 0
    try:
        slide_index = int(slide.SlideIndex)
    except Exception:
        slide_index = 0
    try:
        slide_id = int(slide.SlideID)
    except Exception:
        slide_id = 0

    layout = ""
    try:
        layout = str(slide.CustomLayout.Name)
    except Exception:
        try:
            layout = str(slide.Layout)
        except Exception:
            layout = ""

    shapes: list[dict[str, Any]] = []
    remaining = max_text_chars
    truncated = False
    try:
        shape_count = int(slide.Shapes.Count)
    except Exception:
        shape_count = 0
    for index in range(1, shape_count + 1):
        try:
            shape = slide.Shapes(index)
            text = _shape_texts(shape)
            if remaining <= 0:
                text = ""
                truncated = True
            elif len(text) > remaining:
                text = text[:remaining]
                truncated = True
            remaining -= len(text)
            shapes.append(
                {
                    "index": index,
                    "id": int(shape.Id),
                    "name": str(shape.Name),
                    "type": int(shape.Type),
                    "text": text,
                }
            )
        except Exception:
            continue

    title = _slide_title(slide)
    title, title_truncated = truncate_text(title, max_text_chars)
    return {
        "application": "PowerPoint",
        "slide_index": slide_index,
        "slide_count": slide_count,
        "slide_id": slide_id,
        "layout": layout,
        "title": title,
        "title_truncated": title_truncated,
        "shapes": shapes,
        "truncated": truncated,
    }


def slide(target: ResolvedTarget, *, max_text_chars: int) -> dict[str, Any]:
    return _slide_result(target, _current_slide(target), max_text_chars)


def document(target: ResolvedTarget, *, max_text_chars: int) -> dict[str, Any]:
    presentation = target.document
    info: dict[str, Any] = {
        "application": "PowerPoint",
        "name": target.info.get("document", ""),
        "full_name": target.info.get("full_name", ""),
        "path": target.info.get("path", ""),
        "slide_count": 0,
        "current_slide": target.info.get("slide_index", 0),
        "slide_titles": [],
    }
    try:
        info["slide_count"] = int(presentation.Slides.Count)
    except Exception:
        pass
    try:
        info["slide_width"] = float(presentation.PageSetup.SlideWidth)
        info["slide_height"] = float(presentation.PageSetup.SlideHeight)
    except Exception:
        pass

    titles: list[dict[str, Any]] = []
    remaining = max_text_chars
    try:
        count = int(presentation.Slides.Count)
    except Exception:
        count = 0
    for index in range(1, min(count, 1000) + 1):
        if remaining <= 0:
            break
        try:
            item = presentation.Slides(index)
            title = _slide_title(item)
            title, truncated = truncate_text(title, min(remaining, 500))
            remaining -= len(title)
            titles.append(
                {"slide_index": index, "title": title, "truncated": truncated}
            )
        except Exception:
            continue
    info["slide_titles"] = titles
    return info


# ---------------------------------------------------------------------------
# Geometry helpers (all Office shape geometry is reported in points)

UNIT = "pt"
_TEXTBOX_ORIENTATION_HORIZONTAL = 1  # msoTextOrientationHorizontal
_LAYOUT_TEXT = 2  # ppLayoutText
_MAX_SLIDE_WINDOW = 100


def _slide_size(presentation: Any) -> tuple[float | None, float | None]:
    try:
        return float(presentation.PageSetup.SlideWidth), float(presentation.PageSetup.SlideHeight)
    except Exception:
        return None, None


def _shape_geometry(shape: Any) -> dict[str, Any]:
    geometry: dict[str, Any] = {"unit": UNIT}
    for key in ("Left", "Top", "Width", "Height"):
        try:
            geometry[key.lower()] = float(getattr(shape, key))
        except Exception:
            geometry[key.lower()] = None
    return geometry


def _notes_text(slide: Any) -> str:
    try:
        placeholder = slide.NotesPage.Shapes.Placeholders(2)
        return str(placeholder.TextFrame.TextRange.Text)
    except Exception:
        return ""


def _slide_count(presentation: Any) -> int:
    try:
        return int(presentation.Slides.Count)
    except Exception:
        return 0


def _slide_at(presentation: Any, index: int) -> Any:
    if isinstance(index, bool) or not isinstance(index, int) or index < 1:
        raise invalid_request("Parameter 'slide' must be a 1-based slide number.")
    try:
        return presentation.Slides(index)
    except Exception as exc:
        raise _slide_error(f"Slide {index} could not be resolved.", exc) from exc


def _shape_at(slide: Any, ref: Any) -> Any:
    if isinstance(ref, bool) or not isinstance(ref, (int, str)) or not str(ref).strip():
        raise invalid_request(
            "Parameter 'shape' must be a 1-based shape index or a shape name."
        )
    try:
        return slide.Shapes(ref)
    except Exception as exc:
        raise _slide_error(f"Shape {ref!r} could not be resolved on this slide.", exc) from exc


def _has_text_frame(shape: Any) -> bool | None:
    try:
        return bool(shape.HasTextFrame)
    except Exception:
        return None


def _shape_identity(shape: Any) -> dict[str, Any]:
    identity: dict[str, Any] = {}
    for key, member in (("shape_name", "Name"), ("shape_type", "Type"), ("shape_index", "Index")):
        try:
            identity[key] = jsonable(getattr(shape, member))
        except Exception:
            continue
    return identity


def _text_range(shape: Any, slide: int, ref: Any) -> Any:
    """TextRange of a shape that currently holds text."""
    if _has_text_frame(shape) is False:
        raise text_frame_missing(
            f"Shape {ref!r} on slide {slide} has no text frame (HasTextFrame = 0), "
            "so it holds no text to read or change.",
            {"slide": slide, "shape": ref, "has_text_frame": 0, **_shape_identity(shape)},
        )
    try:
        if not bool(shape.TextFrame.HasText):
            raise invalid_request(
                "The shape has no text to change.",
                {
                    "slide": slide,
                    "shape": ref,
                    "has_text_frame": 1,
                    "has_text": 0,
                    **_shape_identity(shape),
                    "hint": (
                        "The text frame is empty. Writing into it is allowed; use a write "
                        "operation (powerpoint text set) instead of a read or a format write."
                    ),
                },
            )
        return shape.TextFrame.TextRange
    except BridgeError:
        raise
    except Exception as exc:
        raise _slide_error(f"Shape {ref!r} has no usable text frame.", exc) from exc


def _write_text_range(shape: Any, slide: int, ref: Any) -> Any:
    """TextRange to write into: a text frame is required, existing text is not.

    Filling an empty placeholder (a blank TOC body, for example) is a normal
    write, so unlike _text_range this accepts HasText = 0.
    """
    if _has_text_frame(shape) is False:
        raise text_frame_missing(
            f"Shape {ref!r} on slide {slide} has no text frame (HasTextFrame = 0), so no "
            "text can be written into it. Pick a text-bearing placeholder or add a text "
            "box with 'office powerpoint textbox add'.",
            {"slide": slide, "shape": ref, "has_text_frame": 0, **_shape_identity(shape)},
        )
    try:
        return shape.TextFrame.TextRange
    except Exception as exc:
        raise _slide_error(f"Shape {ref!r} has no usable text frame.", exc) from exc


# ---------------------------------------------------------------------------
# Paragraph text helpers
#
# PowerPoint separates paragraphs with CR (Char(13)); LF is an ordinary
# character inside a paragraph. Shells and JSON files routinely deliver CRLF
# or LF, so every text write normalizes line endings and reports the count.
# ---------------------------------------------------------------------------

PARAGRAPH_BREAK = "\r"


def _normalize_paragraph_breaks(text: str) -> tuple[str, int, int]:
    """Return the CR-separated text, its paragraph-break count, and LF conversions."""
    converted = text.count("\n")
    normalized = text.replace("\r\n", PARAGRAPH_BREAK).replace("\n", PARAGRAPH_BREAK)
    return normalized, normalized.count(PARAGRAPH_BREAK), converted


def _paragraph_count(text_range: Any) -> int | None:
    try:
        return int(text_range.Paragraphs().Count)
    except Exception:
        return None


def _paragraph_slice(text_range: Any, start: int, count: int) -> Any:
    """Paragraphs(start) for one paragraph, Paragraphs(start, count) for a block.

    The two-argument form addresses a whole block in one call, which is how a
    level or size is applied to a range of paragraphs at once.
    """
    if count == 1:
        return text_range.Paragraphs(start)
    return text_range.Paragraphs(start, count)


def _optional_number(owner: Any, name: str) -> float | None:
    try:
        return float(getattr(owner, name))
    except Exception:
        return None


def _optional_int(owner: Any, name: str) -> int | None:
    try:
        return int(getattr(owner, name))
    except Exception:
        return None


def _optional_text(owner: Any, name: str) -> str | None:
    try:
        return str(getattr(owner, name))
    except Exception:
        return None


def _optional_member(owner: Any, name: str) -> Any:
    try:
        return getattr(owner, name)
    except Exception:
        return None


def _parse_paragraph_ranges(value: Any) -> list[tuple[int, int]]:
    """Parse a paragraph selector into 1-based (start, count) blocks.

    Accepts "4", "4-10", "4:10", "4..10", or a list of those, and a two-element
    [start, end] pair. Blocks are returned in ascending order.
    """
    if isinstance(value, bool) or value is None:
        raise invalid_request(
            "Parameter 'paragraph' must be a 1-based paragraph number or a "
            "\"first-last\" range.",
            {"parameter": "paragraph"},
        )
    specs: list[str] = []
    if isinstance(value, (list, tuple)):
        specs = [str(item) for item in value]
    elif isinstance(value, int):
        specs = [str(value)]
    else:
        text = str(value).strip()
        if not text:
            raise invalid_request("Parameter 'paragraph' must not be empty.")
        specs = [part.strip() for part in text.split(",") if part.strip()]
    if not specs:
        raise invalid_request("Parameter 'paragraph' must name at least one paragraph.")

    ranges: list[tuple[int, int]] = []
    for spec in specs:
        match = re.fullmatch(r"(\d+)\s*(?:-|\.\.|:)?\s*(\d+)?", spec)
        if not match:
            raise invalid_request(
                f"Parameter 'paragraph' value {spec!r} is not a paragraph number or "
                "\"first-last\" range.",
                {"parameter": "paragraph", "value": spec},
            )
        first = int(match.group(1))
        last = int(match.group(2)) if match.group(2) else first
        if first < 1 or last < first:
            raise invalid_request(
                f"Parameter 'paragraph' value {spec!r} must be a 1-based range whose "
                "end is not before its start.",
                {"parameter": "paragraph", "value": spec},
            )
        ranges.append((first, last - first + 1))
    return sorted(ranges)


def _paragraph_label(ranges: list[tuple[int, int]]) -> str:
    return ",".join(
        str(start) if count == 1 else f"{start}-{start + count - 1}" for start, count in ranges
    )


# ---------------------------------------------------------------------------
# Read: every slide with shape geometry, text, and speaker notes


def slides(
    target: ResolvedTarget,
    *,
    start: int,
    count: int,
    max_text_chars: int,
) -> dict[str, Any]:
    presentation = target.document
    if isinstance(start, bool) or not isinstance(start, int) or start < 1:
        raise invalid_request("Parameter 'start' must be a 1-based slide number.")
    if isinstance(count, bool) or not isinstance(count, int) or count < 1:
        raise invalid_request("Parameter 'count' must be a positive number of slides.")
    count = min(count, _MAX_SLIDE_WINDOW)

    total = _slide_count(presentation)
    slide_width, slide_height = _slide_size(presentation)
    items: list[dict[str, Any]] = []
    remaining = max_text_chars
    truncated = False
    last = min(total, start + count - 1)
    for index in range(start, last + 1):
        try:
            slide = presentation.Slides(index)
        except Exception:
            continue
        shapes: list[dict[str, Any]] = []
        try:
            shape_count = int(slide.Shapes.Count)
        except Exception:
            shape_count = 0
        for shape_index in range(1, shape_count + 1):
            try:
                shape = slide.Shapes(shape_index)
            except Exception:
                continue
            text = _shape_texts(shape)
            if remaining <= 0 and text:
                text = ""
                truncated = True
            elif len(text) > remaining:
                text = text[:remaining]
                truncated = True
            remaining -= len(text)
            entry: dict[str, Any] = {
                "index": shape_index,
                "name": str(shape.Name),
                "text": text,
            }
            try:
                entry["id"] = int(shape.Id)
                entry["type"] = int(shape.Type)
            except Exception:
                pass
            entry["geometry"] = _shape_geometry(shape)
            shapes.append(entry)
        notes = _notes_text(slide)
        if len(notes) > remaining:
            notes = notes[:remaining]
            truncated = True
        remaining -= len(notes)
        items.append(
            {
                "slide_index": index,
                "title": _slide_title(slide),
                "notes": notes,
                "shapes": shapes,
            }
        )

    return {
        "application": "PowerPoint",
        "unit": UNIT,
        "slide_count": total,
        "slide_width": slide_width,
        "slide_height": slide_height,
        "start": start,
        "count": len(items),
        "slides": items,
        "truncated": truncated,
    }


# ---------------------------------------------------------------------------
# Write: in-place text replacement


def text_replace(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
    find: Any,
    replacement: Any,
    dry_run: bool,
    max_text_chars: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    if not isinstance(find, str) or not find:
        raise invalid_request("Parameter 'find' must be a non-empty string.")
    if not isinstance(replacement, str):
        raise invalid_request("Parameter 'replace' must be a string.")
    find, find_breaks, find_converted = _normalize_paragraph_breaks(find)
    replacement, replace_breaks, replace_converted = _normalize_paragraph_breaks(replacement)
    check_expected_document(target, params)

    slide = _slide_at(target.document, slide_index)
    shape = _shape_at(slide, shape_ref)
    text_range = _text_range(shape, slide_index, shape_ref)
    before = str(text_range.Text)
    occurrences = before.count(find)

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "changed": False,
        "slide": slide_index,
        "shape": shape_ref,
        "find": find,
        "replace": replacement,
        "occurrences": occurrences,
        "characters_before": len(before),
        "paragraph_breaks": replace_breaks,
        "text_preview": before[:max_text_chars] if max_text_chars > 0 else "",
    }
    if find_converted or replace_converted:
        result["newline_conversions"] = find_converted + replace_converted
        result["note"] = (
            "Line feeds in the arguments were converted to CR: PowerPoint separates "
            "paragraphs with CR, and LF would otherwise sit inside one paragraph."
        )
    if dry_run:
        return result
    require_confirm(params, "powerpoint.text.replace")
    if occurrences == 0:
        result["verified"] = False
        result["note"] = "The text was not found on this shape; nothing changed."
        return result

    try:
        shape.TextFrame.TextRange.Replace(find, replacement)
        after = str(shape.TextFrame.TextRange.Text)
    except Exception as exc:
        raise _slide_error("The text replacement failed.", exc) from exc

    result.update(
        {
            "changed": after != before,
            "characters_after": len(after),
            "find_remaining": after.count(find),
            "replace_count": after.count(replacement) if replacement else 0,
            "verified": after != before,
            "text_preview": after[:max_text_chars] if max_text_chars > 0 else "",
        }
    )
    return result


# ---------------------------------------------------------------------------
# Write: add a text box


def textbox_add(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    left: Any,
    top: Any,
    width: Any,
    height: Any,
    text: Any,
    font_size: Any,
    dry_run: bool,
    max_text_chars: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    numbers = {"left": left, "top": top, "width": width, "height": height}
    for name, value in numbers.items():
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise invalid_request(f"Parameter {name!r} must be a number of points.")
        if float(value) < 0:
            raise invalid_request(f"Parameter {name!r} must not be negative.")
    if width is not None and float(width) <= 0:
        raise invalid_request("Parameter 'width' must be positive.")
    if height is not None and float(height) <= 0:
        raise invalid_request("Parameter 'height' must be positive.")
    if not isinstance(text, str) or not text:
        raise invalid_request("Parameter 'text' must be a non-empty string.")
    text, paragraph_breaks, converted = _normalize_paragraph_breaks(text)
    if len(text) > max_text_chars and max_text_chars > 0:
        raise invalid_request(
            "Parameter 'text' is longer than the configured text limit.",
            {"characters": len(text), "limit": max_text_chars},
        )
    if font_size is not None and (
        isinstance(font_size, bool)
        or not isinstance(font_size, (int, float))
        or not 1 <= float(font_size) <= 400
    ):
        raise invalid_request("Parameter 'font_size' must be a point size between 1 and 400.")
    check_expected_document(target, params)

    slide = _slide_at(target.document, slide_index)
    slide_width, slide_height = _slide_size(target.document)
    planned = {
        "left": float(left),
        "top": float(top),
        "width": float(width),
        "height": float(height),
        "unit": UNIT,
    }
    fits = True
    if slide_width is not None and slide_height is not None:
        fits = (
            planned["left"] + planned["width"] <= slide_width + 0.01
            and planned["top"] + planned["height"] <= slide_height + 0.01
        )

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "added": False,
        "slide": slide_index,
        "geometry": planned,
        "slide_width": slide_width,
        "slide_height": slide_height,
        "fits_in_slide": fits,
        "characters": len(text),
        "paragraph_breaks": paragraph_breaks,
    }
    if converted:
        result["newline_conversions"] = converted
    if not fits:
        result["warning"] = "The planned box extends past the slide edge."
    if dry_run:
        return result
    require_confirm(params, "powerpoint.textbox.add")

    try:
        shape = slide.Shapes.AddTextbox(
            _TEXTBOX_ORIENTATION_HORIZONTAL,
            planned["left"],
            planned["top"],
            planned["width"],
            planned["height"],
        )
        shape.TextFrame.TextRange.Text = text
        if font_size is not None:
            shape.TextFrame.TextRange.Font.Size = float(font_size)
        written = str(shape.TextFrame.TextRange.Text)
        created_index = int(shape.Index)
        created_id = int(shape.Id)
        created_name = str(shape.Name)
        geometry = _shape_geometry(shape)
    except Exception as exc:
        raise _slide_error("The text box could not be added.", exc) from exc

    result.update(
        {
            "added": True,
            "verified": written == text,
            "shape": {
                "index": created_index,
                "id": created_id,
                "name": created_name,
                "geometry": geometry,
            },
        }
    )
    return result


# ---------------------------------------------------------------------------
# Write: add a slide


def slide_add(
    target: ResolvedTarget,
    *,
    index: Any,
    layout: Any,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    presentation = target.document
    if layout is not None and (
        isinstance(layout, bool) or not isinstance(layout, int) or layout < 1
    ):
        raise invalid_request("Parameter 'layout' must be a PowerPoint layout number.")
    check_expected_document(target, params)

    total = _slide_count(presentation)
    if index is None:
        position = total + 1
    else:
        if isinstance(index, bool) or not isinstance(index, int) or index < 1:
            raise invalid_request("Parameter 'index' must be a 1-based slide number.")
        if index > total + 1:
            raise invalid_request(
                "Parameter 'index' is beyond the end of the presentation.",
                {"index": index, "slide_count": total},
            )
        position = index
    layout_number = _LAYOUT_TEXT if layout is None else layout

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "added": False,
        "slide_index": position,
        "layout": layout_number,
        "slide_count_before": total,
    }
    if dry_run:
        return result
    require_confirm(params, "powerpoint.slide.add")

    try:
        presentation.Slides.Add(position, layout_number)
        slide = presentation.Slides(position)
        shapes = []
        for shape_index in range(1, int(slide.Shapes.Count) + 1):
            shape = slide.Shapes(shape_index)
            shapes.append({"index": shape_index, "name": str(shape.Name)})
    except Exception as exc:
        raise _slide_error("The slide could not be added.", exc) from exc

    result.update(
        {
            "added": True,
            "verified": _slide_count(presentation) == total + 1,
            "slide_count_after": _slide_count(presentation),
            "shapes": shapes,
        }
    )
    return result


# ---------------------------------------------------------------------------
# Read/write: speaker notes


def _notes_text_range(slide: Any, slide_index: int) -> Any:
    try:
        placeholder = slide.NotesPage.Shapes.Placeholders(2)
        return placeholder.TextFrame.TextRange
    except Exception as exc:
        raise _slide_error(
            f"Slide {slide_index} has no notes placeholder.", exc
        ) from exc


def _same_text(left: str, right: str) -> bool:
    def normalize(value: str) -> str:
        return value.replace("\r\n", "\r").replace("\n", "\r")

    return normalize(left) == normalize(right)


def notes(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    max_text_chars: int,
) -> dict[str, Any]:
    slide = _slide_at(target.document, slide_index)
    text_range = _notes_text_range(slide, slide_index)
    try:
        text = str(text_range.Text)
    except Exception as exc:
        raise _slide_error("The notes could not be read.", exc) from exc
    notes_text, truncated = truncate_text(text, max_text_chars)
    return {
        "application": "PowerPoint",
        "slide": slide_index,
        "notes": notes_text,
        "characters": len(text),
        "truncated": truncated,
    }


def notes_set(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    text: Any,
    dry_run: bool,
    max_text_chars: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    if not isinstance(text, str):
        raise invalid_request("Parameter 'text' must be a string.")
    text, paragraph_breaks, converted = _normalize_paragraph_breaks(text)
    if max_text_chars > 0 and len(text) > max_text_chars:
        raise invalid_request(
            "Parameter 'text' is longer than the configured text limit.",
            {"characters": len(text), "limit": max_text_chars},
        )
    check_expected_document(target, params)

    slide = _slide_at(target.document, slide_index)
    text_range = _notes_text_range(slide, slide_index)
    before = str(text_range.Text)

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "changed": False,
        "slide": slide_index,
        "characters_before": len(before),
        "characters_after": len(text),
        "paragraph_breaks": paragraph_breaks,
        "notes_preview": truncate_text(before, max_text_chars)[0],
    }
    if converted:
        result["newline_conversions"] = converted
    if dry_run:
        return result
    require_confirm(params, "powerpoint.notes.set")

    try:
        text_range.Text = text
        after = str(text_range.Text)
    except Exception as exc:
        raise _slide_error("The notes could not be changed.", exc) from exc

    result.update(
        {
            "changed": after != before,
            "verified": _same_text(after, text),
            "notes_preview": truncate_text(after, max_text_chars)[0],
        }
    )
    return result


# ---------------------------------------------------------------------------
# Write: delete a slide


def slide_delete(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    presentation = target.document
    check_expected_document(target, params)

    slide = _slide_at(presentation, slide_index)
    total = _slide_count(presentation)
    title = _slide_title(slide)

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "deleted": False,
        "slide": slide_index,
        "slide_title": title,
        "slide_count_before": total,
    }
    if dry_run:
        return result
    require_confirm(params, "powerpoint.slide.delete")

    try:
        slide.Delete()
    except Exception as exc:
        raise _slide_error("The slide could not be deleted.", exc) from exc

    after = _slide_count(presentation)
    result.update(
        {
            "deleted": True,
            "verified": after == total - 1,
            "slide_count_after": after,
            "undo_hint": _UNDO_HINT,
        }
    )
    return result


# ---------------------------------------------------------------------------
# Write: move a slide
# ---------------------------------------------------------------------------


def slide_move(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    to_index: Any,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Move a slide with Slides(n).MoveTo and verify through the SlideID."""
    presentation = target.document
    check_expected_document(target, params)
    slide = _slide_at(presentation, slide_index)
    if isinstance(to_index, bool) or not isinstance(to_index, int) or to_index < 1:
        raise invalid_request("Parameter 'to' must be a 1-based slide number.")
    if to_index == slide_index:
        raise invalid_request(
            "Parameter 'to' must differ from 'slide'.",
            {"slide": slide_index, "to": to_index},
        )
    total = _slide_count(presentation)
    if to_index > total:
        raise invalid_request(
            "Parameter 'to' is beyond the last slide.",
            {"to": to_index, "slide_count": total},
        )
    try:
        slide_id = int(slide.SlideID)
    except Exception:
        slide_id = None
    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "moved": False,
        "slide": slide_index,
        "to": to_index,
        "direction": "up" if to_index < slide_index else "down",
        "slide_title": _slide_title(slide),
        "slide_count": total,
    }
    if dry_run:
        return result
    require_confirm(params, "powerpoint.slide.move")

    try:
        slide.MoveTo(to_index)
    except Exception as exc:
        raise _slide_error("The slide could not be moved.", exc) from exc
    try:
        landed = presentation.Slides(to_index)
        landed_index = int(landed.SlideIndex)
        landed_id = int(landed.SlideID)
    except Exception as exc:
        raise _slide_error(
            "The slide move call ran but the moved slide could not be read back.", exc
        ) from exc
    result.update(
        {
            "moved": True,
            "slide_title_after": _slide_title(landed),
            "verified": landed_index == to_index
            and (slide_id is None or landed_id == slide_id),
        }
    )
    return result


# ---------------------------------------------------------------------------
# Write: delete a shape
# ---------------------------------------------------------------------------


def shape_delete(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
    dry_run: bool,
    max_text_chars: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    presentation = target.document
    check_expected_document(target, params)
    slide = _slide_at(presentation, slide_index)
    shape = _shape_at(slide, shape_ref)
    try:
        count_before = int(slide.Shapes.Count)
    except Exception as exc:
        raise _slide_error("The shapes on this slide could not be counted.", exc) from exc
    try:
        name = str(shape.Name)
    except Exception:
        name = ""
    try:
        shape_type: int | None = int(shape.Type)
    except Exception:
        shape_type = None
    text_preview = truncate_text(_shape_texts(shape), max_text_chars)[0]

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "deleted": False,
        "slide": slide_index,
        "shape": shape_ref,
        "shape_name": name,
        "shape_type": shape_type,
        "shape_count_before": count_before,
        "text_preview": text_preview,
    }
    if dry_run:
        return result
    require_confirm(params, "powerpoint.shape.delete")

    try:
        shape.Delete()
    except Exception as exc:
        raise _slide_error("The shape could not be deleted.", exc) from exc
    try:
        count_after: int | None = int(slide.Shapes.Count)
    except Exception:
        count_after = None
    result.update(
        {
            "deleted": True,
            "shape_count_after": count_after,
            "verified": count_after is not None and count_after == count_before - 1,
            "undo_hint": _UNDO_HINT,
        }
    )
    return result


# ---------------------------------------------------------------------------
# Write: set text font formatting on one shape
# ---------------------------------------------------------------------------

_MSO_TRUE = -1
_MSO_FALSE = 0


def _parse_color(value: Any, key: str) -> tuple[int, str]:
    """Parse RRGGBB text into the BGR integer Office stores; return both."""
    if not isinstance(value, str):
        raise invalid_request(
            f"Parameter {key!r} must be a 'RRGGBB' or '#RRGGBB' hex color string.",
            {"parameter": key},
        )
    text = value.strip().lstrip("#")
    if len(text) != 6:
        raise invalid_request(
            f"Parameter {key!r} must be a 'RRGGBB' or '#RRGGBB' hex color string.",
            {"parameter": key},
        )
    try:
        rgb = int(text, 16)
    except ValueError:
        raise invalid_request(
            f"Parameter {key!r} is not a valid hex color: {value!r}.",
            {"parameter": key},
        ) from None
    red = (rgb >> 16) & 0xFF
    green = (rgb >> 8) & 0xFF
    blue = rgb & 0xFF
    # Office stores colors as BGR.
    return blue * 0x10000 + green * 0x100 + red, f"{text.upper()}"


def _color_hex(value: Any) -> str | None:
    try:
        raw = int(value) & 0xFFFFFF
    except (TypeError, ValueError):
        return None
    return f"{raw & 0xFF:02X}{(raw >> 8) & 0xFF:02X}{(raw >> 16) & 0xFF:02X}"


def _read_font(font: Any) -> dict[str, Any]:
    entry: dict[str, Any] = {}
    try:
        entry["name"] = str(font.Name)
    except Exception:
        entry["name"] = None
    try:
        entry["size"] = float(font.Size)
    except Exception:
        entry["size"] = None
    try:
        bold = int(font.Bold)
    except Exception:
        entry["bold"] = None
    else:
        entry["bold"] = True if bold == _MSO_TRUE else False if bold == _MSO_FALSE else "mixed"
    try:
        entry["color"] = _color_hex(font.Color.RGB)
    except Exception:
        entry["color"] = None
    return entry


def shape_set_format(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
    font_size: Any,
    font_name: Any,
    color: Any,
    bold: Any,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Write the allowed TextRange.Font properties on one shape.

    Only the text of the shape is touched: the shape outline and fill
    (Line/Fill) are deliberately left alone. Colors are written as the BGR
    integer Office uses, taken from an RRGGBB string like the Excel
    set_format command, and read back through _color_hex for verification.
    """
    requested: dict[str, Any] = {}
    if font_size is not None:
        if isinstance(font_size, bool) or not isinstance(font_size, (int, float)) or not 1 <= float(font_size) <= 400:
            raise invalid_request("Parameter 'font_size' must be a point size between 1 and 400.")
        requested["font_size"] = float(font_size)
    if font_name is not None:
        if not isinstance(font_name, str) or not font_name.strip():
            raise invalid_request("Parameter 'font_name' must be a font name.")
        requested["font_name"] = font_name.strip()
    if color is not None:
        requested["color"] = _parse_color(color, "color")
    if bold is not None:
        if not isinstance(bold, bool):
            raise invalid_request("Parameter 'bold' must be true or false.")
        requested["bold"] = bold
    if not requested:
        raise invalid_request(
            "At least one of 'font_size', 'font_name', 'color' or 'bold' is required."
        )
    check_expected_document(target, params)

    slide = _slide_at(target.document, slide_index)
    shape = _shape_at(slide, shape_ref)
    text_range = _text_range(shape, slide_index, shape_ref)
    try:
        font_before = _read_font(text_range.Font)
    except Exception as exc:
        raise _slide_error("The shape's current font could not be read.", exc) from exc

    planned: dict[str, Any] = {}
    if "font_size" in requested:
        planned["font_size"] = requested["font_size"]
    if "font_name" in requested:
        planned["font_name"] = requested["font_name"]
    if "color" in requested:
        planned["color"] = requested["color"][1]
    if "bold" in requested:
        planned["bold"] = requested["bold"]

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "changed": False,
        "apply_to": "text_range",
        "slide": slide_index,
        "shape": shape_ref,
        "font_before": font_before,
        "planned": planned,
    }
    if dry_run:
        return result
    require_confirm(params, "powerpoint.shape.set_format")

    try:
        font = text_range.Font
        if "font_name" in requested:
            font.Name = requested["font_name"]
        if "font_size" in requested:
            font.Size = requested["font_size"]
        if "bold" in requested:
            font.Bold = _MSO_TRUE if requested["bold"] else _MSO_FALSE
        if "color" in requested:
            font.Color.RGB = requested["color"][0]
        font_after = _read_font(text_range.Font)
    except Exception as exc:
        raise _slide_error("The shape formatting could not be changed.", exc) from exc

    checks: list[bool] = []
    if "font_name" in requested:
        checks.append(font_after.get("name") == requested["font_name"])
    if "font_size" in requested:
        checks.append(font_after.get("size") == requested["font_size"])
    if "bold" in requested:
        checks.append(font_after.get("bold") is requested["bold"])
    if "color" in requested:
        checks.append(font_after.get("color") == requested["color"][1])
    result.update(
        {
            "changed": True,
            "font_after": font_after,
            "verified": all(checks),
        }
    )
    return result


# ---------------------------------------------------------------------------
# Text: read a shape's text with paragraph levels, sizes, and autofit
#
# These commands exist so a text edit does not turn into twenty `com get`
# calls: one read returns the text, the paragraph levels and sizes, the frame
# geometry, and the autofit state of one shape.
# ---------------------------------------------------------------------------

_MAX_PARAGRAPH_REPORT = 200
_MAX_PARAGRAPH_ENTRIES = 500

_AUTOFIT_BEHAVIORS = {"none": 0, "shape": 1, "text": 2}
_AUTOFIT_LABELS = {"none": "none", "shape": "shape_fits_text", "text": "text_fits_shape"}
_MSO_AUTOSIZE_NAMES = {
    0: "none",
    1: "shape_fits_text",
    2: "text_fits_shape",
    -2: "mixed",
}
_PP_AUTOSIZE_NAMES = {0: "none", 1: "shape_fits_text", -2: "mixed"}
_AUTOFIT_ENUMS = {
    "TextFrame.AutoSize (PpAutoSize)": {"none": 0, "shape_fits_text": 1, "mixed": -2},
    "TextFrame2.AutoSize (MsoAutoSize)": {
        "none": 0,
        "shape_fits_text": 1,
        "text_fits_shape": 2,
        "mixed": -2,
    },
}
_AUTOFIT_HINT = (
    "TextFrame.AutoSize takes PpAutoSize and cannot express 'shrink text on overflow'; "
    "TextFrame2.AutoSize takes MsoAutoSize where text_fits_shape is 2. The curated command "
    "tries TextFrame2 first (dynamic dispatch, then a type-library cast) and reads the "
    "result back. If every attempt fails, ask the user to set it once in the UI "
    "(Text Options > Autofit) or carry the setting in the template layout."
)


def _paragraph_report(text_range: Any, limit: int = _MAX_PARAGRAPH_REPORT) -> list[dict[str, Any]]:
    """Report IndentLevel and font size per paragraph so levels can be planned."""
    total = _paragraph_count(text_range)
    if total is None:
        return []
    report: list[dict[str, Any]] = []
    for index in range(1, min(total, limit) + 1):
        entry: dict[str, Any] = {"paragraph": index}
        try:
            paragraph = text_range.Paragraphs(index)
        except Exception:
            entry.update({"text": None, "indent_level": None, "font_size": None})
            report.append(entry)
            continue
        font = _optional_member(paragraph, "Font")
        entry["text"] = _optional_text(paragraph, "Text")
        entry["indent_level"] = _optional_int(paragraph, "IndentLevel")
        entry["font_size"] = _optional_number(font, "Size") if font is not None else None
        report.append(entry)
    return report


def _autofit_names(value: int | None, table: dict[int, str]) -> str | None:
    if value is None:
        return None
    return table.get(int(value), f"unknown({int(value)})")


def _autofit_state(shape: Any) -> dict[str, Any]:
    """Read both autofit members plus word wrap; missing members stay None."""
    state: dict[str, Any] = {
        "text_frame_autosize": None,
        "text_frame2_autosize": None,
        "word_wrap": None,
        "has_text_frame": None,
        "has_text": None,
        "autofit": None,
        "autofit_member": None,
    }
    if hasattr(shape, "HasTextFrame"):
        state["has_text_frame"] = _optional_int(shape, "HasTextFrame")
    text_frame = _optional_member(shape, "TextFrame")
    if text_frame is not None:
        state["text_frame_autosize"] = _optional_int(text_frame, "AutoSize")
        state["word_wrap"] = _optional_int(text_frame, "WordWrap")
        state["has_text"] = _optional_int(text_frame, "HasText")
        state["autofit"] = _autofit_names(state["text_frame_autosize"], _PP_AUTOSIZE_NAMES)
        if state["autofit"] is not None:
            state["autofit_member"] = "TextFrame.AutoSize"
    text_frame2 = _optional_member(shape, "TextFrame2")
    if text_frame2 is not None:
        value = _optional_int(text_frame2, "AutoSize")
        state["text_frame2_autosize"] = value
        if state["word_wrap"] is None:
            state["word_wrap"] = _optional_int(text_frame2, "WordWrap")
        if value is not None:
            # TextFrame2 is the richer table (it knows text_fits_shape), so it wins.
            state["autofit"] = _autofit_names(value, _MSO_AUTOSIZE_NAMES)
            state["autofit_member"] = "TextFrame2.AutoSize"
    return state


def text_read(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
    max_text_chars: int,
) -> dict[str, Any]:
    """Read one shape's text, paragraph levels, font, geometry, and autofit."""
    slide = _slide_at(target.document, slide_index)
    shape = _shape_at(slide, shape_ref)
    text_range = _text_range(shape, slide_index, shape_ref)
    text = str(text_range.Text)
    preview, truncated = truncate_text(text, max_text_chars)
    font = _optional_member(text_range, "Font")
    placeholder = _optional_member(shape, "PlaceholderFormat")
    return {
        "application": "PowerPoint",
        "unit": UNIT,
        "slide": slide_index,
        "shape": shape_ref,
        "shape_name": _optional_text(shape, "Name"),
        "shape_type": _optional_int(shape, "Type"),
        "placeholder_type": _optional_int(placeholder, "Type") if placeholder is not None else None,
        "text": preview,
        "characters": len(text),
        "truncated": truncated,
        "paragraph_count": _paragraph_count(text_range),
        "paragraphs": _paragraph_report(text_range),
        "font": _read_font(font) if font is not None else {},
        "indent_level": _optional_int(text_range, "IndentLevel"),
        "geometry": _shape_geometry(shape),
        "autofit": _autofit_state(shape),
    }


def autofit_read(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
) -> dict[str, Any]:
    """Read the autofit state of one shape (no confirmation, nothing changes)."""
    slide = _slide_at(target.document, slide_index)
    shape = _shape_at(slide, shape_ref)
    return {
        "application": "PowerPoint",
        "unit": UNIT,
        "slide": slide_index,
        "shape": shape_ref,
        "shape_name": _optional_text(shape, "Name"),
        "autofit": _autofit_state(shape),
        "geometry": _shape_geometry(shape),
        "autosize_enums": _AUTOFIT_ENUMS,
        "note": (
            "'autofit' is the readable behaviour name; text_fits_shape is the UI's "
            "'Shrink text on overflow' and only TextFrame2.AutoSize can report it."
        ),
    }


def _cast_to_textframe2(text_frame: Any) -> Any:
    """Ask the type library for the TextFrame2 interface of a TextFrame.

    Dynamic dispatch resolves member names through IDispatch::GetIDsOfNames, which
    does not hand out every member of a CoClass. A cast takes the generated
    type-library wrapper for TextFrame2, so members the dynamic path will not name
    become reachable. Only real COM dispatch objects can be cast.
    """
    from win32com.client import CastTo, CDispatch  # imported lazily: only a cast needs it

    if not isinstance(text_frame, CDispatch):
        raise TypeError(
            f"{type(text_frame).__name__} is not a COM dispatch object, so it cannot be "
            "cast to TextFrame2"
        )
    return CastTo(text_frame, "TextFrame2")


def _apply_autofit(shape: Any, behavior: str) -> tuple[list[dict[str, Any]], str | None]:
    """Set the autofit behaviour through the first member that accepts it.

    TextFrame2.AutoSize is tried first because it is the only autofit member whose
    enum contains 'text fits shape'; TextFrame.AutoSize takes PpAutoSize and is only
    tried for the behaviours it can express. Every attempt is logged so the caller
    can see which member Office accepted or refused.
    """
    attempts: list[dict[str, Any]] = []
    number = _AUTOFIT_BEHAVIORS[behavior]
    text_frame = _optional_member(shape, "TextFrame")

    def attempt(label: str, holder: Any) -> bool:
        try:
            holder.AutoSize = number
        except Exception as exc:
            entry: dict[str, Any] = {"member": label, "ok": False, "value": number, "error": str(exc)}
            hresult = getattr(exc, "hresult", None)
            if hresult is not None:
                try:
                    entry["hresult"] = f"0x{int(hresult) & 0xFFFFFFFF:08X}"
                except (TypeError, ValueError):
                    pass
            attempts.append(entry)
            return False
        attempts.append({"member": label, "ok": True, "value": number})
        return True

    text_frame2 = _optional_member(shape, "TextFrame2")
    if text_frame2 is not None and attempt("TextFrame2.AutoSize", text_frame2):
        return attempts, "TextFrame2.AutoSize"
    if text_frame is not None:
        try:
            cast = _cast_to_textframe2(text_frame)
        except Exception as exc:
            attempts.append(
                {
                    "member": "TextFrame2.AutoSize (type-library cast)",
                    "ok": False,
                    "value": number,
                    "error": str(exc),
                }
            )
            cast = None
        if cast is not None and attempt("TextFrame2.AutoSize (type-library cast)", cast):
            return attempts, "TextFrame2.AutoSize (type-library cast)"
    if behavior != "text" and text_frame is not None and attempt("TextFrame.AutoSize", text_frame):
        return attempts, "TextFrame.AutoSize"
    return attempts, None


def text_set_autofit(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
    behavior: Any,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Set the autofit behaviour of one shape and read it back."""
    if not isinstance(behavior, str) or behavior.strip().lower() not in _AUTOFIT_BEHAVIORS:
        raise invalid_request(
            "Parameter 'behavior' must be one of: none, shape, text.",
            {"parameter": "behavior", "values": list(_AUTOFIT_BEHAVIORS)},
        )
    wanted = behavior.strip().lower()
    check_expected_document(target, params)

    slide = _slide_at(target.document, slide_index)
    shape = _shape_at(slide, shape_ref)
    before = _autofit_state(shape)
    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "changed": False,
        "slide": slide_index,
        "shape": shape_ref,
        "shape_name": _optional_text(shape, "Name"),
        "behavior": wanted,
        "behavior_label": _AUTOFIT_LABELS[wanted],
        "autofit_before": before,
        "geometry_before": _shape_geometry(shape),
        "autosize_enums": _AUTOFIT_ENUMS,
    }
    if dry_run:
        return result
    require_confirm(params, "powerpoint.text.set_autofit")

    attempts, applied_via = _apply_autofit(shape, wanted)
    after = _autofit_state(shape)
    verified = after.get("autofit") == _AUTOFIT_LABELS[wanted]
    result.update(
        {
            "changed": after != before,
            "autofit_after": after,
            "applied_via": applied_via,
            "attempts": attempts,
            "geometry_after": _shape_geometry(shape),
            "verified": verified,
        }
    )
    if not verified:
        raise autofit_refused(
            f"PowerPoint refused every autofit member tried for behavior {wanted!r}; "
            f"the frame still reports {after.get('autofit')!r}.",
            {
                "slide": slide_index,
                "shape": shape_ref,
                "behavior": wanted,
                "attempts": attempts,
                "autofit_before": before,
                "autofit_after": after,
                "autosize_enums": _AUTOFIT_ENUMS,
                "hint": _AUTOFIT_HINT,
            },
        )
    return result


def text_measure(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
    text: Any,
    font_size: Any,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Measure the height a frame would need by growing a throw-away duplicate.

    The shape is duplicated, the duplicate is switched to 'shape fits text' so
    PowerPoint itself lays the text out, its height is read, and the duplicate is
    deleted. Because the duplicate carries the real font and frame width, the
    number is PowerPoint's own measurement rather than an estimate.
    """
    if text is not None and not isinstance(text, str):
        raise invalid_request("Parameter 'text' must be a string.")
    if font_size is not None and (
        isinstance(font_size, bool)
        or not isinstance(font_size, (int, float))
        or not 1 <= float(font_size) <= 400
    ):
        raise invalid_request("Parameter 'font_size' must be a point size between 1 and 400.")
    check_expected_document(target, params)

    slide = _slide_at(target.document, slide_index)
    shape = _shape_at(slide, shape_ref)
    text_range = _write_text_range(shape, slide_index, shape_ref)
    current_text = str(text_range.Text)
    probe_text, breaks, converted = _normalize_paragraph_breaks(
        current_text if text is None else text
    )
    geometry = _shape_geometry(shape)

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "unit": UNIT,
        "dry_run": dry_run,
        "measured": False,
        "slide": slide_index,
        "shape": shape_ref,
        "shape_name": _optional_text(shape, "Name"),
        "geometry": geometry,
        "characters": len(probe_text),
        "paragraphs": 1 + breaks,
        "font_size": float(font_size) if font_size is not None else None,
        "note": (
            "Measuring duplicates the shape, lets PowerPoint grow the duplicate to fit "
            "its text, reads the height, and deletes the duplicate. The slide's shape "
            "count changes twice; one undo cannot undo the whole measurement."
        ),
    }
    if converted:
        result["newline_conversions"] = converted
    if dry_run:
        return result
    require_confirm(params, "powerpoint.text.measure")

    duplicate: Any = None
    deleted = False
    try:
        try:
            duplicate = shape.Duplicate()
        except Exception as exc:
            raise _slide_error("The shape could not be duplicated for measuring.", exc) from exc
        probe_range = _optional_member(duplicate, "TextFrame")
        if probe_range is None:
            raise BridgeError(
                ErrorCode.UNSUPPORTED_OPERATION,
                "The duplicated shape has no text frame, so its required height cannot be measured.",
                409,
                {"slide": slide_index, "shape": shape_ref},
            )
        duplicate_range = probe_range.TextRange
        if text is not None:
            duplicate_range.Text = probe_text
        if font_size is not None:
            duplicate_range.Font.Size = float(font_size)
        attempts, applied_via = _apply_autofit(duplicate, "shape")
        required_height = _optional_number(duplicate, "Height")
        required_width = _optional_number(duplicate, "Width")
        measured_size = _optional_number(_optional_member(duplicate_range, "Font"), "Size")
        try:
            duplicate.Delete()
            deleted = True
        except Exception as exc:
            result["duplicate_delete_error"] = str(exc)
    finally:
        if duplicate is not None and not deleted:
            try:
                duplicate.Delete()
                deleted = True
            except Exception:
                pass

    if required_height is None:
        raise BridgeError(
            ErrorCode.UNSUPPORTED_OPERATION,
            "PowerPoint did not report a height for the duplicated shape, so the text "
            "could not be measured. Check whether the frame has autofit locked.",
            409,
            {"slide": slide_index, "shape": shape_ref, "attempts": attempts},
        )

    height = geometry.get("height")
    fits = height is None or required_height <= float(height) + 0.01
    result.update(
        {
            "measured": True,
            "required_height": required_height,
            "required_width": required_width,
            "current_height": height,
            "fits": fits,
            "margin": None if height is None else round(float(height) - required_height, 2),
            "measured_font_size": measured_size,
            "applied_via": applied_via,
            "attempts": attempts,
            "duplicate_deleted": deleted,
            "undo_hint": _UNDO_HINT if not deleted else None,
        }
    )
    if not fits:
        result["note"] = (
            f"The text needs {required_height:.1f} pt but the frame is {height:.1f} pt: "
            "shrink the font, reduce the paragraphs, or set autofit to text_fits_shape."
        )
    return result


def _normalize_paragraph_entries(value: Any) -> list[dict[str, Any]]:
    """Validate a paragraph array into text / indent_level / font_size entries."""
    if not isinstance(value, list) or not value:
        raise invalid_request(
            "Parameter 'paragraphs' must be a non-empty array of paragraph strings or objects.",
            {"parameter": "paragraphs"},
        )
    if len(value) > _MAX_PARAGRAPH_ENTRIES:
        raise invalid_request(
            f"Parameter 'paragraphs' has {len(value)} items, which exceeds the limit of "
            f"{_MAX_PARAGRAPH_ENTRIES}.",
            {"parameter": "paragraphs", "limit": _MAX_PARAGRAPH_ENTRIES},
        )
    entries: list[dict[str, Any]] = []
    for position, item in enumerate(value, start=1):
        if isinstance(item, str):
            raw: dict[str, Any] = {"text": item}
        elif isinstance(item, dict):
            raw = item
        else:
            raise invalid_request(
                "Each paragraph must be a string or an object with a 'text' key.",
                {"paragraph": position},
            )
        unknown = sorted(set(raw) - {"text", "level", "indent_level", "font_size"})
        if unknown:
            raise invalid_request(
                f"Paragraph {position} has unsupported keys: {', '.join(unknown)}.",
                {
                    "paragraph": position,
                    "allowed": ["text", "level", "indent_level", "font_size"],
                },
            )
        if not isinstance(raw.get("text"), str):
            raise invalid_request(
                f"Paragraph {position} needs a 'text' string.",
                {"paragraph": position},
            )
        indent = raw.get("indent_level", raw.get("level"))
        if indent is not None and (
            isinstance(indent, bool) or not isinstance(indent, int) or not 1 <= indent <= 8
        ):
            raise invalid_request(
                "Parameter 'indent_level' ('level') must be a paragraph level between 1 and 8.",
                {"paragraph": position},
            )
        size = raw.get("font_size")
        if size is not None and (
            isinstance(size, bool) or not isinstance(size, (int, float)) or not 1 <= float(size) <= 400
        ):
            raise invalid_request(
                "Parameter 'font_size' must be a point size between 1 and 400.",
                {"paragraph": position},
            )
        text, breaks, converted = _normalize_paragraph_breaks(raw["text"])
        entries.append(
            {
                "text": text,
                "breaks": breaks,
                "indent_level": indent,
                "font_size": float(size) if size is not None else None,
                "converted": converted,
            }
        )
    return entries


def _write_paragraphs(
    text_range: Any,
    entries: list[dict[str, Any]],
    joined: str,
) -> tuple[str, list[dict[str, Any]], bool]:
    """Write the joined text, then apply per-entry levels and sizes.

    Returns the read-back text, a per-entry report of what was applied, and
    whether everything verified. Paragraph positions are cumulative because one
    entry can carry several CR-separated paragraphs of its own.
    """
    try:
        text_range.Text = joined
    except Exception as exc:
        raise _slide_error("The shape text could not be written.", exc) from exc

    reports: list[dict[str, Any]] = []
    verified = _same_text(str(text_range.Text), joined)
    position = 1
    for entry in entries:
        count = 1 + entry["breaks"]
        report: dict[str, Any] = {
            "paragraph": position,
            "paragraph_count": count,
            "text": entry["text"],
        }
        if entry["indent_level"] is not None:
            report["indent_level"] = entry["indent_level"]
        if entry["font_size"] is not None:
            report["font_size"] = entry["font_size"]
        if entry["indent_level"] is not None or entry["font_size"] is not None:
            try:
                block = _paragraph_slice(text_range, position, count)
                if entry["indent_level"] is not None:
                    block.IndentLevel = entry["indent_level"]
                if entry["font_size"] is not None:
                    block.Font.Size = entry["font_size"]
                if entry["indent_level"] is not None:
                    after_level = _optional_int(block, "IndentLevel")
                    report["indent_level_after"] = after_level
                    verified = verified and after_level == entry["indent_level"]
                if entry["font_size"] is not None:
                    after_size = _optional_number(
                        _optional_member(block, "Font"), "Size"
                    )
                    report["font_size_after"] = after_size
                    verified = verified and after_size == entry["font_size"]
            except Exception as exc:
                report["error"] = str(exc)
                verified = False
        reports.append(report)
        position += count
    return str(text_range.Text), reports, verified


def text_set(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
    paragraphs: Any,
    dry_run: bool,
    max_text_chars: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Replace one shape's whole text with a paragraph array, levels and sizes included.

    A whole-frame rewrite through `com set TextRange.Text` loses the previous
    content unless the caller remembered to read it, and cannot set per-paragraph
    levels. This operation reports before/after counts and a preview, and applies
    IndentLevel / Font.Size per paragraph through the two-argument Paragraphs form.
    """
    entries = _normalize_paragraph_entries(paragraphs)
    joined = PARAGRAPH_BREAK.join(entry["text"] for entry in entries)
    converted = sum(entry["converted"] for entry in entries)
    if max_text_chars > 0 and len(joined) > max_text_chars:
        raise invalid_request(
            "The paragraphs add up to more text than the configured limit.",
            {"characters": len(joined), "limit": max_text_chars},
        )
    check_expected_document(target, params)

    slide = _slide_at(target.document, slide_index)
    shape = _shape_at(slide, shape_ref)
    text_range = _write_text_range(shape, slide_index, shape_ref)
    before = str(text_range.Text)
    paragraphs_before = _paragraph_count(text_range)

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "changed": False,
        "slide": slide_index,
        "shape": shape_ref,
        "shape_name": _optional_text(shape, "Name"),
        "characters_before": len(before),
        "characters_after": len(joined),
        "paragraphs_before": paragraphs_before,
        "paragraphs_after": sum(1 + entry["breaks"] for entry in entries),
        "paragraph_breaks": joined.count(PARAGRAPH_BREAK),
        "autofit_before": _autofit_state(shape),
        "text_preview": truncate_text(before, max_text_chars)[0],
        "planned_text": truncate_text(joined, max_text_chars)[0],
        "planned": [
            {
                "paragraph": position,
                **({"indent_level": entry["indent_level"]} if entry["indent_level"] is not None else {}),
                **({"font_size": entry["font_size"]} if entry["font_size"] is not None else {}),
            }
            for position, entry in enumerate(entries, start=1)
        ],
    }
    if converted:
        result["newline_conversions"] = converted
    if dry_run:
        return result
    require_confirm(params, "powerpoint.text.set")

    after, reports, verified = _write_paragraphs(text_range, entries, joined)
    result.update(
        {
            "changed": after != before,
            "characters_after": len(after),
            "paragraphs_after": _paragraph_count(text_range),
            "verified": verified,
            "text_preview": truncate_text(after, max_text_chars)[0],
            "paragraphs": reports,
            "autofit_after": _autofit_state(shape),
        }
    )
    if after != joined and _same_text(after, joined):
        result["note"] = (
            "PowerPoint returned the text with different line breaks than written; the "
            "CR-normalized text matches."
        )
    return result


def text_set_format(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
    paragraph: Any,
    font_size: Any,
    font_name: Any,
    color: Any,
    bold: Any,
    indent_level: Any,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Set font and paragraph level on a whole text frame or on paragraph ranges.

    `paragraph` selects 1-based paragraphs ("4", "4-10", or a list of both); with
    `indent_level` this is how a TOC body gets its level structure without writing
    raw COM expressions.
    """
    requested: dict[str, Any] = {}
    if font_size is not None:
        if (
            isinstance(font_size, bool)
            or not isinstance(font_size, (int, float))
            or not 1 <= float(font_size) <= 400
        ):
            raise invalid_request("Parameter 'font_size' must be a point size between 1 and 400.")
        requested["font_size"] = float(font_size)
    if font_name is not None:
        if not isinstance(font_name, str) or not font_name.strip():
            raise invalid_request("Parameter 'font_name' must be a font name.")
        requested["font_name"] = font_name.strip()
    if color is not None:
        requested["color"] = _parse_color(color, "color")
    if bold is not None:
        if not isinstance(bold, bool):
            raise invalid_request("Parameter 'bold' must be true or false.")
        requested["bold"] = bold
    if indent_level is not None:
        if (
            isinstance(indent_level, bool)
            or not isinstance(indent_level, int)
            or not 1 <= indent_level <= 8
        ):
            raise invalid_request("Parameter 'indent_level' must be a paragraph level between 1 and 8.")
        requested["indent_level"] = indent_level
    if not requested:
        raise invalid_request(
            "At least one of 'font_size', 'font_name', 'color', 'bold' or 'indent_level' "
            "is required."
        )
    ranges = _parse_paragraph_ranges(paragraph) if paragraph is not None else []
    check_expected_document(target, params)

    slide = _slide_at(target.document, slide_index)
    shape = _shape_at(slide, shape_ref)
    text_range = _text_range(shape, slide_index, shape_ref)

    blocks: list[tuple[str, Any]] = []
    if not ranges:
        blocks.append(("text_range", text_range))
    else:
        for start, count in ranges:
            label = f"paragraphs({_paragraph_label([(start, count)])})"
            try:
                blocks.append((label, _paragraph_slice(text_range, start, count)))
            except Exception as exc:
                raise _slide_error(
                    f"Paragraph range {start}-{start + count - 1} could not be resolved.",
                    exc,
                ) from exc

    planned: dict[str, Any] = {}
    if "font_size" in requested:
        planned["font_size"] = requested["font_size"]
    if "font_name" in requested:
        planned["font_name"] = requested["font_name"]
    if "color" in requested:
        planned["color"] = requested["color"][1]
    if "bold" in requested:
        planned["bold"] = requested["bold"]
    if "indent_level" in requested:
        planned["indent_level"] = requested["indent_level"]

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "changed": False,
        "apply_to": "paragraphs" if ranges else "text_range",
        "slide": slide_index,
        "shape": shape_ref,
        "paragraph": _paragraph_label(ranges) if ranges else None,
        "planned": planned,
        "blocks": [
            {
                "target": label,
                "font_before": _read_font(holder.Font) if _optional_member(holder, "Font") else {},
                "indent_level_before": _optional_int(holder, "IndentLevel"),
            }
            for label, holder in blocks
        ],
    }
    if dry_run:
        return result
    require_confirm(params, "powerpoint.text.set_format")

    checks: list[bool] = []
    for entry, (label, holder) in zip(result["blocks"], blocks):
        try:
            if "font_name" in requested:
                holder.Font.Name = requested["font_name"]
            if "font_size" in requested:
                holder.Font.Size = requested["font_size"]
            if "bold" in requested:
                holder.Font.Bold = _MSO_TRUE if requested["bold"] else _MSO_FALSE
            if "color" in requested:
                holder.Font.Color.RGB = requested["color"][0]
            if "indent_level" in requested:
                holder.IndentLevel = requested["indent_level"]
        except Exception as exc:
            raise _slide_error(f"The formatting of {label} could not be changed.", exc) from exc
        font = _optional_member(holder, "Font")
        entry["font_after"] = _read_font(font) if font is not None else {}
        entry["indent_level_after"] = _optional_int(holder, "IndentLevel")
        if "font_name" in requested:
            checks.append(entry["font_after"].get("name") == requested["font_name"])
        if "font_size" in requested:
            checks.append(entry["font_after"].get("size") == requested["font_size"])
        if "bold" in requested:
            checks.append(entry["font_after"].get("bold") is requested["bold"])
        if "color" in requested:
            checks.append(entry["font_after"].get("color") == requested["color"][1])
        if "indent_level" in requested:
            checks.append(entry["indent_level_after"] == requested["indent_level"])

    single = result["blocks"][0]
    result.update(
        {
            "changed": True,
            "font_before": single["font_before"],
            "font_after": single["font_after"],
            "indent_level_before": single["indent_level_before"],
            "indent_level_after": single["indent_level_after"],
            "verified": all(checks),
        }
    )
    return result


# ---------------------------------------------------------------------------
# Table of contents: fill one body placeholder from the deck's slide titles
# ---------------------------------------------------------------------------

_FULLWIDTH_DIGITS = "０１２３４５６７８９"
_FULLWIDTH_TO_ASCII = str.maketrans(_FULLWIDTH_DIGITS, "0123456789")
_ASCII_TO_FULLWIDTH = str.maketrans("0123456789", _FULLWIDTH_DIGITS)
_NUMBERING_STYLES = ("auto", "fullwidth-dot", "ascii-dot", "digits", "none")
_NUMBER_PATTERNS = (
    (re.compile(rf"^([{_FULLWIDTH_DIGITS}]{{1,3}})([．。][ \u3000]?)(.*)$"), "fullwidth", None),
    (re.compile(r"^([0-9]{1,3})(\.[ \u3000]?)(.*)$"), "ascii", None),
    (re.compile(rf"^([{_FULLWIDTH_DIGITS}]{{1,3}})([\u3000 ]+)(.*)$"), "fullwidth", None),
    (re.compile(r"^([0-9]{1,3})([\u3000 ]+)(.*)$"), "ascii", None),
)


def _detect_numbering(text: str) -> dict[str, Any] | None:
    """Read the numbering scheme an existing TOC already uses.

    A TOC placeholder usually carries a couple of seeded entries ("０．改定履歴",
    "１． motivation") that fix the style: full- or half-width digits, the
    separator, and the starting number. New entries have to match it.
    """
    for line in text.split(PARAGRAPH_BREAK):
        stripped = line.strip()
        if not stripped:
            continue
        for pattern, digits, literal_separator in _NUMBER_PATTERNS:
            match = pattern.match(stripped)
            if not match:
                continue
            number_text = match.group(1)
            separator = literal_separator if literal_separator is not None else match.group(2)
            start_text = number_text.translate(_FULLWIDTH_TO_ASCII)
            try:
                start = int(start_text)
            except ValueError:
                continue
            return {
                "detected_from": stripped,
                "digits": digits,
                "separator": separator,
                "start": start,
            }
    return None


def _apply_numbering(title: str, index: int, style: dict[str, Any] | None) -> str:
    if not style or style["digits"] == "none":
        return title
    digits = str(index)
    if style["digits"] == "fullwidth":
        digits = digits.translate(_ASCII_TO_FULLWIDTH)
    return f"{digits}{style['separator']}{title}"


def _parse_exclude(value: Any) -> set[int]:
    if value is None:
        return set()
    items: list[Any] = []
    if isinstance(value, (list, tuple)):
        items = list(value)
    elif isinstance(value, int) and not isinstance(value, bool):
        items = [value]
    else:
        items = [part.strip() for part in str(value).split(",") if part.strip()]
    excluded: set[int] = set()
    for item in items:
        if isinstance(item, bool) or not isinstance(item, int):
            try:
                item = int(str(item))
            except ValueError as exc:
                raise invalid_request(
                    f"Parameter 'exclude' value {item!r} is not a slide number.",
                    {"parameter": "exclude"},
                ) from exc
        if item < 1:
            raise invalid_request(
                "Parameter 'exclude' must hold 1-based slide numbers.",
                {"parameter": "exclude"},
            )
        excluded.add(int(item))
    return excluded


def _parse_levels(value: Any) -> dict[str, Any]:
    """Parse a level specification into {slide_number: level} plus an auto flag."""
    if value is None:
        return {"auto": False, "map": {}}
    if isinstance(value, str) and value.strip().lower() in {"auto", "true"}:
        return {"auto": True, "map": {}}
    pairs: list[str] = []
    mapping: dict[str, Any] = {}
    if isinstance(value, str):
        pairs = [part.strip() for part in value.split(",") if part.strip()]
    elif isinstance(value, (list, tuple)):
        for item in value:
            if isinstance(item, dict):
                mapping.update(item)
            else:
                pairs.append(str(item))
    elif isinstance(value, dict):
        mapping.update(value)
    else:
        raise invalid_request(
            "Parameter 'levels' must be 'auto', a spec string such as \"6-12=2\", or an "
            "array/object of slide-range to level pairs.",
            {"parameter": "levels"},
        )
    result: dict[int, int] = {}
    for pair in pairs:
        match = re.fullmatch(r"(\d+)\s*(?:-|:|\.\.)?\s*(\d+)?\s*=\s*(\d+)", pair)
        if not match:
            raise invalid_request(
                f"Parameter 'levels' entry {pair!r} is not \"first-last=level\" or \"slide=level\".",
                {"parameter": "levels", "entry": pair},
            )
        first, last, level = int(match.group(1)), int(match.group(2) or match.group(1)), int(match.group(3))
        if first < 1 or last < first or level < 1:
            raise invalid_request(
                f"Parameter 'levels' entry {pair!r} must name a 1-based slide range and a level of 1 or more.",
                {"parameter": "levels", "entry": pair},
            )
        for slide_number in range(first, last + 1):
            result[slide_number] = level
    for key, level in mapping.items():
        try:
            key_text = str(key)
            level_value = int(level)
        except (TypeError, ValueError) as exc:
            raise invalid_request(
                f"Parameter 'levels' pair {key!r}={level!r} is not a slide range and level.",
                {"parameter": "levels"},
            ) from exc
        match = re.fullmatch(r"(\d+)\s*(?:-|:|\.\.)?\s*(\d+)?", key_text)
        if not match or level_value < 1:
            raise invalid_request(
                f"Parameter 'levels' pair {key!r}={level!r} is not a 1-based slide range and level.",
                {"parameter": "levels", "key": key},
            )
        first, last = int(match.group(1)), int(match.group(2) or match.group(1))
        for slide_number in range(first, last + 1):
            result[slide_number] = level_value
    return {"auto": False, "map": result}


def _auto_levels(entries: list[tuple[int, str, str]]) -> dict[int, int]:
    """Guess level 2 for titles that read as a sub-topic of another slide.

    The signal is the section-slide pattern: a title that is listed in another
    slide's body text (the classic "部品の種類" slide listing アルミフレーム and
    パネル). A wrong guess changes indentation, never content.
    """
    titles = {index: title.strip().lower() for index, title, _ in entries}
    blobs = {index: text.strip().lower() for index, _, text in entries}
    levels: dict[int, int] = {}
    for index, title, _ in entries:
        needle = titles[index]
        if not needle:
            continue
        for other_index in titles:
            if other_index == index:
                continue
            if needle == titles[other_index]:
                continue
            if needle in blobs[other_index]:
                levels[index] = 2
                break
    return levels


def toc_fill(
    target: ResolvedTarget,
    *,
    slide_index: Any,
    shape_ref: Any,
    start: Any,
    end: Any,
    exclude: Any,
    levels: Any,
    numbering: Any,
    font_size: Any,
    sub_font_size: Any,
    dry_run: bool,
    max_text_chars: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Fill a TOC placeholder from the deck's slide titles, matching its numbering.

    Collects the titles of the slide range, keeps the numbering scheme already
    present in the placeholder (full-width digits and a full-width stop, for
    example), demotes sub-topics to level 2 on request, and writes the result
    through the paragraph-aware text write so levels and sizes land in one step.
    """
    presentation = target.document
    total = _slide_count(presentation)
    if start is None:
        start_value = 1
    else:
        if isinstance(start, bool) or not isinstance(start, int) or start < 1:
            raise invalid_request("Parameter 'from' must be a 1-based slide number.")
        start_value = start
    if end is None:
        end_value = total
    else:
        if isinstance(end, bool) or not isinstance(end, int) or end < 1:
            raise invalid_request("Parameter 'to' must be a 1-based slide number.")
        end_value = min(end, total) if total else end
    if end_value < start_value:
        raise invalid_request(
            "Parameter 'to' must not be before 'from'.",
            {"from": start_value, "to": end_value},
        )
    excluded = _parse_exclude(exclude)
    level_spec = _parse_levels(levels)
    if numbering is not None and (
        not isinstance(numbering, str) or numbering.strip().lower() not in _NUMBERING_STYLES
    ):
        raise invalid_request(
            f"Parameter 'numbering' must be one of: {', '.join(_NUMBERING_STYLES)}.",
            {"parameter": "numbering", "values": list(_NUMBERING_STYLES)},
        )
    for name, size in (("font_size", font_size), ("sub_font_size", sub_font_size)):
        if size is not None and (
            isinstance(size, bool) or not isinstance(size, (int, float)) or not 1 <= float(size) <= 400
        ):
            raise invalid_request(
                f"Parameter {name!r} must be a point size between 1 and 400.",
                {"parameter": name},
            )
    check_expected_document(target, params)

    slide = _slide_at(presentation, slide_index)
    shape = _shape_at(slide, shape_ref)
    text_range = _write_text_range(shape, slide_index, shape_ref)
    before = str(text_range.Text)
    paragraphs_before = _paragraph_count(text_range)
    detected = _detect_numbering(before)
    style_choice = (numbering or "auto").strip().lower()
    if style_choice == "auto":
        style = detected or {"digits": "none", "separator": "", "start": 1}
    elif style_choice == "none":
        style = {"digits": "none", "separator": "", "start": 1}
    elif style_choice == "fullwidth-dot":
        style = {"digits": "fullwidth", "separator": "．", "start": detected["start"] if detected else 0}
    elif style_choice == "ascii-dot":
        style = {"digits": "ascii", "separator": ". ", "start": detected["start"] if detected else 1}
    else:  # digits, no separator
        style = {"digits": "ascii", "separator": " ", "start": detected["start"] if detected else 1}

    collected: list[tuple[int, str, str]] = []
    skipped: list[dict[str, Any]] = []
    for index in range(start_value, end_value + 1):
        if index == slide_index or index in excluded:
            skipped.append({"slide_index": index, "reason": "excluded"})
            continue
        try:
            item = presentation.Slides(index)
        except Exception:
            skipped.append({"slide_index": index, "reason": "unresolvable"})
            continue
        title = _normalize_paragraph_breaks(_slide_title(item).strip())[0]
        if not title:
            skipped.append({"slide_index": index, "reason": "no_title"})
            continue
        collected.append((index, title, _slide_body_texts(item)))
    if not collected:
        raise invalid_request(
            "No slide titles were found in the requested range, so there is nothing to write.",
            {"from": start_value, "to": end_value, "skipped": skipped},
        )

    level_map = level_spec["map"] or (_auto_levels(collected) if level_spec["auto"] else {})
    entries: list[dict[str, Any]] = []
    plan: list[dict[str, Any]] = []
    number = int(style["start"])
    for index, title, _text in collected:
        level = int(level_map.get(index, 1))
        text = _apply_numbering(title, number, style) if level == 1 else title
        text, breaks, _converted = _normalize_paragraph_breaks(text)
        if level == 1 and style["digits"] != "none":
            number += 1
        entry: dict[str, Any] = {
            "text": text,
            "breaks": breaks,
            "indent_level": level,
        }
        if font_size is not None and level == 1:
            entry["font_size"] = float(font_size)
        if sub_font_size is not None and level > 1:
            entry["font_size"] = float(sub_font_size)
        entries.append(entry)
        plan.append(
            {
                "slide_index": index,
                "level": level,
                "number": number - 1 if level == 1 and style["digits"] != "none" else None,
                "text": text,
            }
        )

    joined = PARAGRAPH_BREAK.join(entry["text"] for entry in entries)
    if max_text_chars > 0 and len(joined) > max_text_chars:
        raise invalid_request(
            "The generated table of contents is longer than the configured text limit.",
            {"characters": len(joined), "limit": max_text_chars},
        )

    result: dict[str, Any] = {
        "application": "PowerPoint",
        "dry_run": dry_run,
        "changed": False,
        "slide": slide_index,
        "shape": shape_ref,
        "shape_name": _optional_text(shape, "Name"),
        "slide_range": {"from": start_value, "to": end_value},
        "numbering_style": {
            "digits": style["digits"],
            "separator": style["separator"],
            "start": style["start"],
            "requested": style_choice,
            "detected": detected,
        },
        "levels": {"mode": "auto" if level_spec["auto"] else ("specified" if level_spec["map"] else "flat"), "map": level_map},
        "entries": plan,
        "entry_count": len(plan),
        "skipped": skipped,
        "characters_before": len(before),
        "characters_after": len(joined),
        "paragraphs_before": paragraphs_before,
        "paragraphs_after": sum(1 + entry["breaks"] for entry in entries),
        "text_preview": truncate_text(before, max_text_chars)[0],
        "preview_text": truncate_text(joined, max_text_chars)[0],
    }
    if dry_run:
        return result
    require_confirm(params, "powerpoint.toc.fill")

    after, reports, verified = _write_paragraphs(text_range, entries, joined)
    result.update(
        {
            "changed": after != before,
            "characters_after": len(after),
            "paragraphs_after": _paragraph_count(text_range) or len(entries),
            "verified": verified,
            "text_preview": truncate_text(after, max_text_chars)[0],
            "paragraphs": reports,
        }
    )
    return result


# ---------------------------------------------------------------------------
# Read: section names with the slide ranges they cover
# ---------------------------------------------------------------------------


def sections(target: ResolvedTarget) -> dict[str, Any]:
    """Read the presentation's sections as name plus first slide and slide count.

    PowerPoint exposes sections through the parameterized `SectionProperties`
    object (Name / FirstSlide / SlidesCount take the section index); the
    `Slides(n).Section` path the object model does not resolve, so the ranges come
    from SectionProperties when it is available.
    """
    presentation = target.document
    properties = _optional_member(presentation, "SectionProperties")
    if properties is None:
        raise BridgeError(
            ErrorCode.UNSUPPORTED_OPERATION,
            "This PowerPoint build does not expose SectionProperties, so the section "
            "list cannot be read.",
            409,
            {"hint": "Ask the user which slides belong to which section instead of guessing."},
        )
    try:
        count = int(properties.Count)
    except Exception as exc:
        raise _slide_error("The presentation's sections could not be counted.", exc) from exc

    items: list[dict[str, Any]] = []
    for index in range(1, max(count, 0) + 1):
        first: int | None = None
        try:
            first = int(properties.FirstSlide(index))
        except Exception:
            pass
        try:
            slide_count = int(properties.SlidesCount(index))
        except Exception:
            slide_count = None
        entry: dict[str, Any] = {
            "section_index": index,
            "name": None,
            "first_slide": first,
            "slide_count": slide_count,
            "last_slide": (first + slide_count - 1) if first and slide_count else None,
        }
        try:
            entry["name"] = str(properties.Name(index))
        except Exception:
            pass
        try:
            entry["section_id"] = str(properties.SectionID(index))
        except Exception:
            pass
        items.append(entry)

    return {
        "application": "PowerPoint",
        "section_count": max(count, 0),
        "slide_count": _slide_count(presentation),
        "sections": items,
        "note": (
            "Sections are logical groups of consecutive slides; slide_count 0 or -1 "
            "means the section is empty on this build."
        ),
    }
