from __future__ import annotations

from typing import Any

import pywintypes

from ..errors import BridgeError, ErrorCode, invalid_request
from ..target import ResolvedTarget
from .common import (
    check_expected_document,
    require_confirm,
    truncate_text,
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


def _text_range(shape: Any, slide: int, ref: Any) -> Any:
    try:
        if not bool(shape.HasTextFrame) or not bool(shape.TextFrame.HasText):
            raise invalid_request(
                "The shape has no text to change.",
                {"slide": slide, "shape": ref},
            )
        return shape.TextFrame.TextRange
    except BridgeError:
        raise
    except Exception as exc:
        raise _slide_error(f"Shape {ref!r} has no usable text frame.", exc) from exc


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
        "text_preview": before[:max_text_chars] if max_text_chars > 0 else "",
    }
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
    }
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
