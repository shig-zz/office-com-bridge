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
        "notes_preview": truncate_text(before, max_text_chars)[0],
    }
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
