from __future__ import annotations

from typing import Any

import pywintypes

from ..errors import BridgeError, ErrorCode, invalid_request
from ..target import ResolvedTarget
from .common import truncate_text


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
