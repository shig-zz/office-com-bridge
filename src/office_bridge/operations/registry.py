from __future__ import annotations

from typing import Any

from ..config import Settings
from ..errors import invalid_request, unsupported_application, unsupported_operation
from ..target import AppKind, ResolvedTarget
from . import com_ops, excel, powerpoint, undo, word
from .com_ops import MODES, ROOTS
from .common import (
    document_result,
    optional_bool,
    require_str,
    selection_result,
    unsaved_changes,
)


def _require_app(target: ResolvedTarget, expected: AppKind, operation: str) -> None:
    if target.app is not expected:
        raise unsupported_application(target.app.value, operation)


def inspect(target: ResolvedTarget, kind: str, params: dict[str, Any], settings: Settings) -> dict[str, Any]:
    if kind == "selection":
        return selection_result(
            target,
            limit=settings.max_cells,
            max_text_chars=settings.max_text_chars,
        )
    if kind == "document":
        return document_result(target, max_text_chars=settings.max_text_chars)

    if kind in {"excel.selection", "word.selection", "powerpoint.slide"}:
        if kind == "excel.selection":
            _require_app(target, AppKind.EXCEL, kind)
            return excel.selection(target, limit=settings.max_cells)
        if kind == "word.selection":
            _require_app(target, AppKind.WORD, kind)
            return word.selection(target, max_text_chars=settings.max_text_chars)
        _require_app(target, AppKind.POWERPOINT, kind)
        return powerpoint.slide(target, max_text_chars=settings.max_text_chars)

    if kind == "excel.range":
        _require_app(target, AppKind.EXCEL, kind)
        address = require_str(params, "address")
        include_formulas = optional_bool(params, "include_formulas", False)
        return excel.read_range(
            target,
            address=address,
            include_formulas=include_formulas,
            max_cells=settings.max_cells,
            sheet_name=params.get("sheet"),
        )

    if kind == "excel.range.format":
        _require_app(target, AppKind.EXCEL, kind)
        address = require_str(params, "address")
        return excel.read_format(
            target,
            address=address,
            max_cells=settings.max_cells,
            include_borders=optional_bool(params, "borders", False),
            sheet_name=params.get("sheet"),
        )

    if kind in {"excel.workbook", "word.document", "powerpoint.document"}:
        if kind == "excel.workbook":
            _require_app(target, AppKind.EXCEL, kind)
            return excel.document(target)
        if kind == "word.document":
            _require_app(target, AppKind.WORD, kind)
            return word.document(target)
        _require_app(target, AppKind.POWERPOINT, kind)
        return powerpoint.document(target, max_text_chars=settings.max_text_chars)

    if kind == "word.paragraphs":
        _require_app(target, AppKind.WORD, kind)
        return word.paragraphs(
            target,
            start=params.get("start", 1),
            count=params.get("count", 50),
            max_text_chars=settings.max_text_chars,
        )

    if kind == "word.find":
        _require_app(target, AppKind.WORD, kind)
        return word.find(
            target,
            find_text=params.get("find"),
            match_case=optional_bool(params, "match_case", False),
            whole_words=optional_bool(params, "whole_words", False),
            max_matches=int(params.get("max", 50) or 50),
            max_text_chars=settings.max_text_chars,
        )

    if kind == "powerpoint.slides":
        _require_app(target, AppKind.POWERPOINT, kind)
        return powerpoint.slides(
            target,
            start=params.get("start", 1),
            count=params.get("count", 20),
            max_text_chars=settings.max_text_chars,
        )

    if kind == "powerpoint.notes":
        _require_app(target, AppKind.POWERPOINT, kind)
        return powerpoint.notes(
            target,
            slide_index=params.get("slide"),
            max_text_chars=settings.max_text_chars,
        )

    if kind == "powerpoint.text":
        _require_app(target, AppKind.POWERPOINT, kind)
        return powerpoint.text_read(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
            max_text_chars=settings.max_text_chars,
        )

    if kind == "powerpoint.text.autofit":
        _require_app(target, AppKind.POWERPOINT, kind)
        return powerpoint.autofit_read(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
        )

    if kind == "powerpoint.sections":
        _require_app(target, AppKind.POWERPOINT, kind)
        return powerpoint.sections(target)

    raise unsupported_operation(kind)


def invoke(
    target: ResolvedTarget,
    operation: str,
    params: dict[str, Any],
    settings: Settings,
) -> dict[str, Any]:
    result = _invoke(target, operation, params, settings)
    # Applied writes report whether the document now holds unsaved changes so
    # agents can tell the user to save in Office without being asked.
    if isinstance(result, dict) and not optional_bool(params, "dry_run", False):
        result.setdefault("unsaved_changes", unsaved_changes(target))
    return result


def _invoke(
    target: ResolvedTarget,
    operation: str,
    params: dict[str, Any],
    settings: Settings,
) -> dict[str, Any]:
    if operation == "undo":
        return undo.undo(
            target,
            dry_run=optional_bool(params, "dry_run", False),
            params=params,
        )

    if operation == "excel.range.write":
        _require_app(target, AppKind.EXCEL, operation)
        address = require_str(params, "address")
        if "values" not in params:
            raise invalid_request("Parameter 'values' is required.")
        dry_run = optional_bool(params, "dry_run", False)
        return excel.write_range(
            target,
            address=address,
            values=params.get("values"),
            dry_run=dry_run,
            max_cells=settings.max_cells,
            params=params,
            sheet_name=params.get("sheet"),
        )

    if operation == "word.selection.insert_text":
        _require_app(target, AppKind.WORD, operation)
        if "text" not in params:
            raise invalid_request("Parameter 'text' is required.")
        dry_run = optional_bool(params, "dry_run", False)
        replace = optional_bool(params, "replace", True)
        return word.insert_text(
            target,
            text=params.get("text"),
            replace=replace,
            dry_run=dry_run,
            max_text_chars=settings.max_text_chars,
            params=params,
        )

    if operation == "excel.range.copy_format":
        _require_app(target, AppKind.EXCEL, operation)
        source_address = require_str(params, "source_address")
        address = require_str(params, "address")
        dry_run = optional_bool(params, "dry_run", False)
        return excel.copy_format(
            target,
            source_address=source_address,
            address=address,
            dry_run=dry_run,
            max_cells=settings.max_cells,
            params=params,
            include_borders=optional_bool(params, "borders", False),
            sheet_name=params.get("sheet"),
        )

    if operation == "excel.range.set_format":
        _require_app(target, AppKind.EXCEL, operation)
        address = require_str(params, "address")
        dry_run = optional_bool(params, "dry_run", False)
        return excel.set_format(
            target,
            address=address,
            dry_run=dry_run,
            max_cells=settings.max_cells,
            params=params,
            sheet_name=params.get("sheet"),
        )

    if operation == "excel.range.add_hyperlink":
        _require_app(target, AppKind.EXCEL, operation)
        address = require_str(params, "address")
        dry_run = optional_bool(params, "dry_run", False)
        return excel.add_hyperlink(
            target,
            address=address,
            dry_run=dry_run,
            max_cells=settings.max_cells,
            params=params,
            sheet_name=params.get("sheet"),
        )

    if operation == "excel.range.sort":
        _require_app(target, AppKind.EXCEL, operation)
        address = require_str(params, "address")
        if "keys" not in params:
            raise invalid_request("Parameter 'keys' is required.")
        dry_run = optional_bool(params, "dry_run", False)
        return excel.sort_range(
            target,
            address=address,
            keys=params.get("keys"),
            dry_run=dry_run,
            max_cells=settings.max_cells,
            params=params,
            sheet_name=params.get("sheet"),
        )

    if operation in {
        "excel.rows.insert",
        "excel.rows.delete",
        "excel.columns.insert",
        "excel.columns.delete",
    }:
        _require_app(target, AppKind.EXCEL, operation)
        dimension, mode = operation.split(".")[1:]
        spec_key = "row" if dimension == "rows" else "column"
        if spec_key not in params:
            raise invalid_request(f"Parameter {spec_key!r} is required.")
        dry_run = optional_bool(params, "dry_run", False)
        return excel.insert_or_delete(
            target,
            dimension=dimension,
            spec=params.get(spec_key),
            shift=params.get("shift"),
            mode=mode,
            dry_run=dry_run,
            params=params,
        )

    if operation == "excel.rows.move":
        _require_app(target, AppKind.EXCEL, operation)
        if "row" not in params or "to" not in params:
            raise invalid_request("Parameters 'row' and 'to' are required.")
        dry_run = optional_bool(params, "dry_run", False)
        return excel.move_rows(
            target,
            row=params.get("row"),
            to=params.get("to"),
            sheet_name=params.get("sheet"),
            dry_run=dry_run,
            params=params,
        )

    if operation == "word.replace":
        _require_app(target, AppKind.WORD, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return word.replace(
            target,
            find_text=params.get("find"),
            replacement=params.get("replace"),
            match_case=optional_bool(params, "match_case", False),
            whole_words=optional_bool(params, "whole_words", False),
            nth=params.get("nth"),
            dry_run=dry_run,
            params=params,
        )

    if operation == "word.paragraph.replace_text":
        _require_app(target, AppKind.WORD, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return word.replace_paragraph_text(
            target,
            paragraph=params.get("paragraph"),
            text=params.get("text"),
            dry_run=dry_run,
            max_text_chars=settings.max_text_chars,
            params=params,
        )

    if operation == "powerpoint.text.replace":
        _require_app(target, AppKind.POWERPOINT, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return powerpoint.text_replace(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
            find=params.get("find"),
            replacement=params.get("replace"),
            dry_run=dry_run,
            max_text_chars=settings.max_text_chars,
            params=params,
        )

    if operation == "powerpoint.textbox.add":
        _require_app(target, AppKind.POWERPOINT, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return powerpoint.textbox_add(
            target,
            slide_index=params.get("slide"),
            left=params.get("left"),
            top=params.get("top"),
            width=params.get("width"),
            height=params.get("height"),
            text=params.get("text"),
            font_size=params.get("font_size"),
            dry_run=dry_run,
            max_text_chars=settings.max_text_chars,
            params=params,
        )

    if operation == "powerpoint.slide.add":
        _require_app(target, AppKind.POWERPOINT, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return powerpoint.slide_add(
            target,
            index=params.get("index"),
            layout=params.get("layout"),
            dry_run=dry_run,
            params=params,
        )

    if operation == "powerpoint.notes.set":
        _require_app(target, AppKind.POWERPOINT, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return powerpoint.notes_set(
            target,
            slide_index=params.get("slide"),
            text=params.get("text"),
            dry_run=dry_run,
            max_text_chars=settings.max_text_chars,
            params=params,
        )

    if operation == "powerpoint.slide.delete":
        _require_app(target, AppKind.POWERPOINT, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return powerpoint.slide_delete(
            target,
            slide_index=params.get("slide"),
            dry_run=dry_run,
            params=params,
        )

    if operation == "powerpoint.slide.move":
        _require_app(target, AppKind.POWERPOINT, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return powerpoint.slide_move(
            target,
            slide_index=params.get("slide"),
            to_index=params.get("to"),
            dry_run=dry_run,
            params=params,
        )

    if operation == "powerpoint.shape.delete":
        _require_app(target, AppKind.POWERPOINT, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return powerpoint.shape_delete(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
            dry_run=dry_run,
            max_text_chars=settings.max_text_chars,
            params=params,
        )

    if operation == "powerpoint.shape.set_format":
        _require_app(target, AppKind.POWERPOINT, operation)
        dry_run = optional_bool(params, "dry_run", False)
        return powerpoint.shape_set_format(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
            font_size=params.get("font_size"),
            font_name=params.get("font_name"),
            color=params.get("color"),
            bold=params.get("bold"),
            dry_run=dry_run,
            params=params,
        )

    if operation == "powerpoint.text.set":
        _require_app(target, AppKind.POWERPOINT, operation)
        if "paragraphs" not in params:
            raise invalid_request("Parameter 'paragraphs' is required.")
        return powerpoint.text_set(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
            paragraphs=params.get("paragraphs"),
            dry_run=optional_bool(params, "dry_run", False),
            max_text_chars=settings.max_text_chars,
            params=params,
        )

    if operation == "powerpoint.text.set_format":
        _require_app(target, AppKind.POWERPOINT, operation)
        return powerpoint.text_set_format(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
            paragraph=params.get("paragraph"),
            font_size=params.get("font_size"),
            font_name=params.get("font_name"),
            color=params.get("color"),
            bold=params.get("bold"),
            indent_level=params.get("indent_level"),
            dry_run=optional_bool(params, "dry_run", False),
            params=params,
        )

    if operation == "powerpoint.text.set_autofit":
        _require_app(target, AppKind.POWERPOINT, operation)
        return powerpoint.text_set_autofit(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
            behavior=params.get("behavior"),
            dry_run=optional_bool(params, "dry_run", False),
            params=params,
        )

    if operation == "powerpoint.text.measure":
        _require_app(target, AppKind.POWERPOINT, operation)
        return powerpoint.text_measure(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
            text=params.get("text"),
            font_size=params.get("font_size"),
            dry_run=optional_bool(params, "dry_run", False),
            params=params,
        )

    if operation == "powerpoint.toc.fill":
        _require_app(target, AppKind.POWERPOINT, operation)
        return powerpoint.toc_fill(
            target,
            slide_index=params.get("slide"),
            shape_ref=params.get("shape"),
            start=params.get("from"),
            end=params.get("to"),
            exclude=params.get("exclude"),
            levels=params.get("levels"),
            numbering=params.get("numbering"),
            font_size=params.get("font_size"),
            sub_font_size=params.get("sub_font_size"),
            dry_run=optional_bool(params, "dry_run", False),
            max_text_chars=settings.max_text_chars,
            params=params,
        )

    raise unsupported_operation(operation)


def set_property(
    target: ResolvedTarget,
    name: str,
    value: Any,
    params: dict[str, Any],
    settings: Settings,
) -> dict[str, Any]:
    if name == "excel.active_sheet":
        _require_app(target, AppKind.EXCEL, name)
        result = excel.set_active_sheet(target, value, params)
        if isinstance(result, dict):
            result.setdefault("unsaved_changes", unsaved_changes(target))
        return result
    raise unsupported_operation(name)


def capabilities(settings: Settings) -> dict[str, Any]:
    return {
        "limits": {
            "max_cells": settings.max_cells,
            "max_text_chars": settings.max_text_chars,
        },
        "security": {
            "localhost_only": True,
            "writes_require_confirmation": True,
            "allowed_operations_only": False,
            "general_com_access": True,
            "blocked_calls": com_ops.blocked_summary(settings),
        },
        "inspections": [
            {
                "name": "selection",
                "applications": ["Excel", "Word", "PowerPoint"],
                "mode": "read",
                "description": "Read the current selection. Excel returns the selected range; Word returns selected text; PowerPoint returns the current slide.",
                "params": {},
            },
            {
                "name": "document",
                "applications": ["Excel", "Word", "PowerPoint"],
                "mode": "read",
                "description": "Read basic information about the attached document, workbook, or presentation.",
                "params": {},
            },
            {
                "name": "excel.range",
                "applications": ["Excel"],
                "mode": "read",
                "description": "Read values (and optionally formulas) from an A1-style range on a named sheet, defaulting to the active sheet.",
                "params": {
                    "address": {"type": "string", "required": True},
                    "include_formulas": {"type": "boolean", "required": False},
                    "sheet": {"type": "string", "required": False, "note": "Sheet name; defaults to the active sheet."},
                },
            },
            {
                "name": "excel.range.format",
                "applications": ["Excel"],
                "mode": "read",
                "description": (
                    "Read cell formatting (fill, font, number format, alignment) "
                    "cell by cell, with colors converted from Excel's BGR integers "
                    "to RRGGBB hex. Pass borders: true to include per-side border "
                    "style, weight and color. Multi-cell Interior.Color/Borders "
                    "reads are unreliable on this build, so every cell is read "
                    "individually; keep the range small."
                ),
                "params": {
                    "address": {"type": "string", "required": True},
                    "borders": {"type": "boolean", "required": False},
                    "sheet": {"type": "string", "required": False, "note": "Sheet name; defaults to the active sheet."},
                },
            },
            {
                "name": "excel.workbook",
                "applications": ["Excel"],
                "mode": "read",
                "description": "Read workbook metadata and worksheet names.",
                "params": {},
            },
            {
                "name": "word.selection",
                "applications": ["Word"],
                "mode": "read",
                "description": "Read the selected text and its character offsets.",
                "params": {},
            },
            {
                "name": "word.document",
                "applications": ["Word"],
                "mode": "read",
                "description": "Read document metadata and basic counts.",
                "params": {},
            },
            {
                "name": "word.find",
                "applications": ["Word"],
                "mode": "read",
                "description": (
                    "Search the document body and report each match's paragraph "
                    "number and a text snippet. Read-only (no confirmation). The "
                    "match count reflects the bridge's own scan of paragraph text."
                ),
                "params": {
                    "find": {"type": "string", "required": True},
                    "match_case": {"type": "boolean", "required": False},
                    "whole_words": {"type": "boolean", "required": False},
                    "max": {"type": "number", "required": False, "note": "Maximum matches to report, default 50."},
                },
            },
            {
                "name": "word.paragraphs",
                "applications": ["Word"],
                "mode": "read",
                "description": (
                    "Read a window of paragraphs with their style name and heading "
                    "flag. Up to 200 paragraphs per call."
                ),
                "params": {
                    "start": {"type": "number", "required": False, "note": "1-based paragraph number, default 1."},
                    "count": {"type": "number", "required": False, "note": "Paragraphs to read, default 50, max 200."},
                },
            },
            {
                "name": "powerpoint.slide",
                "applications": ["PowerPoint"],
                "mode": "read",
                "description": "Read the current slide, its title, and shape texts.",
                "params": {},
            },
            {
                "name": "powerpoint.document",
                "applications": ["PowerPoint"],
                "mode": "read",
                "description": "Read presentation metadata and slide titles.",
                "params": {},
            },
            {
                "name": "powerpoint.slides",
                "applications": ["PowerPoint"],
                "mode": "read",
                "description": (
                    "Read a window of slides with speaker notes and every shape "
                    "(index, id, name, type, text, and left/top/width/height "
                    "geometry in points). Use it to plan text edits without "
                    "reading the file directly."
                ),
                "params": {
                    "start": {"type": "number", "required": False, "note": "1-based slide number, default 1."},
                    "count": {"type": "number", "required": False, "note": "Slides to read, default 20, max 100."},
                },
            },
            {
                "name": "powerpoint.notes",
                "applications": ["PowerPoint"],
                "mode": "read",
                "description": "Read the speaker notes of one slide.",
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                },
            },
            {
                "name": "powerpoint.text",
                "applications": ["PowerPoint"],
                "mode": "read",
                "description": (
                    "Read one shape's text with its paragraph count, per-paragraph "
                    "IndentLevel and font size, the shape geometry, and the autofit "
                    "state. Use this instead of a chain of com get calls when planning a "
                    "text edit: one call answers text, levels, sizes, and frame height."
                ),
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "shape": {"type": "number|string", "required": True, "note": "1-based shape index or shape name."},
                },
            },
            {
                "name": "powerpoint.text.autofit",
                "applications": ["PowerPoint"],
                "mode": "read",
                "description": (
                    "Read a shape's autofit behaviour from both members (TextFrame.AutoSize "
                    "uses PpAutoSize, TextFrame2.AutoSize uses MsoAutoSize) plus word wrap "
                    "and geometry. 'text_fits_shape' is the UI's 'Shrink text on overflow'."
                ),
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "shape": {"type": "number|string", "required": True, "note": "1-based shape index or shape name."},
                },
            },
            {
                "name": "powerpoint.sections",
                "applications": ["PowerPoint"],
                "mode": "read",
                "description": (
                    "Read the presentation's sections through SectionProperties: name, "
                    "first slide, slide count, and last slide for every section. The "
                    "Slides(n).Section member is not reachable on this build, so the "
                    "ranges come from the parameterized SectionProperties object."
                ),
                "params": {},
            },
        ],
        "operations": [
            {
                "name": "excel.range.write",
                "applications": ["Excel"],
                "mode": "write",
                "description": "Write a rectangular two-dimensional array of literal values to an A1-style range.",
                "requires_confirmation": True,
                "params": {
                    "address": {"type": "string", "required": True},
                    "values": {"type": "array", "required": True, "items": "row arrays"},
                    "sheet": {"type": "string", "required": False, "note": "Sheet name; defaults to the active sheet."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "word.selection.insert_text",
                "applications": ["Word"],
                "mode": "write",
                "description": "Insert or replace text at the current selection. Newlines are allowed.",
                "requires_confirmation": True,
                "params": {
                    "text": {"type": "string", "required": True},
                    "replace": {"type": "boolean", "required": False},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "word.replace",
                "applications": ["Word"],
                "mode": "write",
                "description": (
                    "Replace occurrences of a literal string throughout the body "
                    "using Word's own Find (formatting preserved; tracked when "
                    "Track Changes is on). Dry-run reports the match count and the "
                    "first paragraph before anything changes; the run re-scans and "
                    "reports 'verified'. Pass nth to replace a single numbered "
                    "match instead of all of them."
                ),
                "requires_confirmation": True,
                "params": {
                    "find": {"type": "string", "required": True},
                    "replace": {"type": "string", "required": True, "note": "Replacement text; empty string deletes the match."},
                    "match_case": {"type": "boolean", "required": False},
                    "whole_words": {"type": "boolean", "required": False},
                    "nth": {"type": "number", "required": False, "note": "1-based match number to replace only this one; omit to replace all."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "word.paragraph.replace_text",
                "applications": ["Word"],
                "mode": "write",
                "description": (
                    "Replace one paragraph's body text by 1-based number, keeping "
                    "the paragraph break. Dry-run previews the current text and the "
                    "paragraph count before/after (newlines split paragraphs); the "
                    "run reads the paragraph back and reports 'verified'. Editing a "
                    "paragraph inside a table or field is unreliable: prefer "
                    "word.replace for in-place edits."
                ),
                "requires_confirmation": True,
                "params": {
                    "paragraph": {"type": "number", "required": True, "note": "1-based paragraph number."},
                    "text": {"type": "string", "required": True},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "excel.range.copy_format",
                "applications": ["Excel"],
                "mode": "write",
                "description": "Copy formatting only (font, fill, borders, number format, alignment) from a source range to a target range of the same size. Cell values are not changed. This build does not propagate cell borders through a format paste: pass borders=true and the operation re-applies them cell by cell (reported as borders_applied).",
                "requires_confirmation": True,
                "params": {
                    "source_address": {"type": "string", "required": True},
                    "address": {"type": "string", "required": True},
                    "sheet": {"type": "string", "required": False, "note": "Sheet name for both ranges; defaults to the active sheet."},
                    "borders": {"type": "boolean", "required": False},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "excel.range.add_hyperlink",
                "applications": ["Excel"],
                "mode": "write",
                "description": (
                    "Add a real hyperlink object to an A1-style range. The link "
                    "covers the whole range and the top-left cell value is used as "
                    "display text. If the range already has a hyperlink, its "
                    "address (and optionally screen tip and sub address) is updated "
                    "instead, so the operation is idempotent."
                ),
                "requires_confirmation": True,
                "params": {
                    "address": {"type": "string", "required": True},
                    "url": {
                        "type": "string",
                        "required": True,
                        "note": "Destination, for example 'https://example.com/'. Must not start with '='.",
                    },
                    "sub_address": {
                        "type": "string",
                        "required": False,
                        "note": "Anchor inside the destination or a sheet reference such as 'Sheet2!A1'.",
                    },
                    "screen_tip": {"type": "string", "required": False},
                    "sheet": {"type": "string", "required": False, "note": "Sheet name; defaults to the active sheet."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "excel.range.sort",
                "applications": ["Excel"],
                "mode": "write",
                "description": (
                    "Sort the rows of an A1-style range in place with a stable multi-key "
                    "sort (up to 3 keys). Rows are moved as whole rows (values, cell "
                    "formatting and hyperlinks travel with the row) using Excel row "
                    "insert/delete moves, because this Excel build ignores Range.Sort calls "
                    "whose key is a Range object passed through automation. The range must "
                    "cover every column whose contents should stay aligned; rows outside "
                    "the range are not touched."
                ),
                "requires_confirmation": True,
                "params": {
                    "address": {"type": "string", "required": True},
                    "keys": {
                        "type": "array",
                        "required": True,
                        "items": "objects",
                        "note": (
                            "1 to 3 objects, applied in order (first key is primary). Each "
                            "has 'column' (a 1-based offset from the range's left edge, or a "
                            "sheet column letter such as 'C' that lies inside the range) and "
                            "an optional 'order' ('ascending' default or 'descending')."
                        ),
                    },
                    "sheet": {"type": "string", "required": False, "note": "Sheet name; defaults to the active sheet."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "excel.rows.insert",
                "applications": ["Excel"],
                "mode": "write",
                "description": (
                    "Insert whole rows on the active sheet at a row number or the "
                    "first row of a \"first:last\" range (\"61\" or \"2:4\"). The "
                    "shift enum is applied for you (down by default) and the used "
                    "range is re-read afterwards to report 'verified'."
                ),
                "requires_confirmation": True,
                "params": {
                    "row": {"type": "number or string", "required": True, "note": "1-based row number or \"first:last\" range."},
                    "shift": {"type": "string", "required": False, "note": "down (default) or right."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "excel.rows.delete",
                "applications": ["Excel"],
                "mode": "write",
                "description": (
                    "Delete whole rows on the active sheet by row number or "
                    "\"first:last\" range. Range-limited deletes stay unavailable: "
                    "only whole rows are deleted. The used range is re-read to "
                    "report 'verified', and the result carries an undo hint."
                ),
                "requires_confirmation": True,
                "params": {
                    "row": {"type": "number or string", "required": True, "note": "1-based row number or \"first:last\" range."},
                    "shift": {"type": "string", "required": False, "note": "omit for the reliable no-shift call; up or left to force a shift."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "excel.columns.insert",
                "applications": ["Excel"],
                "mode": "write",
                "description": (
                    "Insert whole columns on the active sheet at a column letter, "
                    "a 1-based index, or an \"A:C\" range (\"C\", \"3\", or \"A:C\"). "
                    "The shift enum is applied for you (right by default)."
                ),
                "requires_confirmation": True,
                "params": {
                    "column": {"type": "string or number", "required": True, "note": "Column letter, 1-based index, or \"first:last\" range."},
                    "shift": {"type": "string", "required": False, "note": "right (default) or down."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "excel.columns.delete",
                "applications": ["Excel"],
                "mode": "write",
                "description": (
                    "Delete whole columns on the active sheet by letter, 1-based "
                    "index, or \"A:C\" range. The used range is re-read to report "
                    "'verified', and the result carries an undo hint."
                ),
                "requires_confirmation": True,
                "params": {
                    "column": {"type": "string or number", "required": True, "note": "Column letter, 1-based index, or \"first:last\" range."},
                    "shift": {"type": "string", "required": False, "note": "omit for the reliable no-shift call; up or left to force a shift."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "excel.rows.move",
                "applications": ["Excel"],
                "mode": "write",
                "description": (
                    "Move one whole row to a new position with the proven copy/"
                    "insert/clear/delete choreography: values, formatting and "
                    "hyperlinks travel with the row. Reports completed_steps, and "
                    "a mid-move failure returns RANGE_INVALID with the steps done "
                    "so the sheet can be re-read before cleanup; one undo would "
                    "only revert the last step. Uses the clipboard, so no other "
                    "clipboard work may run concurrently."
                ),
                "requires_confirmation": True,
                "params": {
                    "row": {"type": "number", "required": True, "note": "1-based row to move."},
                    "to": {"type": "number", "required": True, "note": "1-based destination row; the row lands exactly at this number."},
                    "sheet": {"type": "string", "required": False, "note": "Sheet name; default is the active sheet."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "excel.range.set_format",
                "applications": ["Excel"],
                "mode": "write",
                "description": "Set explicit formatting properties on a range: font name/size/bold/italic/underline/color, number format, alignment, wrap, fill, and borders. At least one property is required.",
                "requires_confirmation": True,
                "params": {
                    "address": {"type": "string", "required": True},
                    "sheet": {"type": "string", "required": False, "note": "Sheet name; defaults to the active sheet."},
                    "font_name": {"type": "string", "required": False, "note": "Use a pan-CJK font (e.g. MS Gothic, Yu Gothic, Meiryo) to cover both Japanese and Western text."},
                    "font_size": {"type": "number", "required": False},
                    "bold": {"type": "boolean", "required": False},
                    "italic": {"type": "boolean", "required": False},
                    "underline": {"type": "boolean", "required": False},
                    "color": {"type": "string", "required": False, "format": "RRGGBB"},
                    "number_format": {"type": "string", "required": False},
                    "halign": {"type": "string", "required": False, "values": ["left", "center", "right"]},
                    "valign": {"type": "string", "required": False, "values": ["top", "middle", "bottom"]},
                    "wrap": {"type": "boolean", "required": False},
                    "fill": {"type": "string", "required": False, "format": "RRGGBB"},
                    "border": {"type": "boolean|string", "required": False, "values": [True, "continuous"]},
                    "border_color": {"type": "string", "required": False, "format": "RRGGBB"},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.text.replace",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Replace text inside one shape's text frame in place with "
                    "TextRange.Replace, so formatting and layout are preserved. "
                    "Dry-run reports how many times the text occurs; the run "
                    "reads the shape back and returns 'verified' plus a preview."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "shape": {"type": "number|string", "required": True, "note": "1-based shape index or shape name."},
                    "find": {"type": "string", "required": True},
                    "replace": {"type": "string", "required": True, "note": "Replacement text (empty string deletes)."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.textbox.add",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Add a text box to a slide. Geometry is in points (matches "
                    "Left/Top/Width/Height, not EMU). Dry-run reports whether the "
                    "box fits the slide; the run returns the created shape's "
                    "index, id, and name so later calls do not have to guess."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "left": {"type": "number", "required": True, "note": "Points from the slide's left edge."},
                    "top": {"type": "number", "required": True, "note": "Points from the slide's top edge."},
                    "width": {"type": "number", "required": True},
                    "height": {"type": "number", "required": True},
                    "text": {"type": "string", "required": True},
                    "font_size": {"type": "number", "required": False, "note": "Optional point size."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.slide.add",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Add a new slide at a 1-based position (default: end) with a "
                    "layout number (default 2, Title and Content). Returns the "
                    "new slide's index and its placeholder shapes."
                ),
                "requires_confirmation": True,
                "params": {
                    "index": {"type": "number", "required": False, "note": "1-based insertion position; omit to append."},
                    "layout": {"type": "number", "required": False, "note": "ppSlideLayout number, default 2."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.notes.set",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Replace the speaker notes of one slide. Dry-run previews the "
                    "current notes; the run reads them back and returns 'verified'."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "text": {"type": "string", "required": True, "note": "New notes text (empty string clears)."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.slide.delete",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Delete one slide by 1-based number. Dry-run reports the title "
                    "of the slide that would be removed; the run verifies the "
                    "slide count dropped by one. Only rollbackable with 'undo' "
                    "while the delete is still the most recent Office action, so "
                    "ask the user to confirm the exact slide up front."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.slide.move",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Move one slide to a new 1-based position (Slides(n).MoveTo). "
                    "Dry-run reports the direction and the slide title; the run "
                    "verifies the SlideID lands at the requested index. Typical "
                    "use: gathering correction-note slides at the end of a deck."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "to": {"type": "number", "required": True, "note": "1-based destination position."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.shape.delete",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Delete one shape by 1-based index or name. Dry-run reports "
                    "the shape name, type and a text preview so the exact shape "
                    "can be confirmed with the user; the run verifies the shape "
                    "count dropped by one and returns an undo hint."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "shape": {"type": "number or string", "required": True, "note": "1-based shape index or shape name."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.shape.set_format",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Set text font properties (font_size, font_name, RRGGBB "
                    "color, bold) on the text of one shape. Only "
                    "TextFrame.TextRange.Font is written; the shape outline and "
                    "fill are left alone. Dry-run returns the current values and "
                    "the planned change; the run reads the font back and reports "
                    "'verified'. Mixed values across the text read back as "
                    "'mixed' in font_before."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "shape": {"type": "number or string", "required": True, "note": "1-based shape index or shape name."},
                    "font_size": {"type": "number", "required": False, "note": "Point size 1-400."},
                    "font_name": {"type": "string", "required": False},
                    "color": {"type": "string", "required": False, "note": "RRGGBB hex; converted to Office's BGR integer."},
                    "bold": {"type": "boolean", "required": False},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.text.set",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Replace one shape's whole text with an array of paragraphs, applying "
                    "per-paragraph indent level and font size in the same step. Unlike a "
                    "raw TextRange.Text assignment it reports the previous text, the "
                    "before/after paragraph counts, and 'verified'. Line feeds in the "
                    "text are converted to PowerPoint's CR paragraph break and reported "
                    "as paragraph_breaks. CLI: --paragraphs takes JSON or '@path'."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "shape": {"type": "number|string", "required": True, "note": "1-based shape index or shape name."},
                    "paragraphs": {
                        "type": "array",
                        "required": True,
                        "items": "strings or {text, level|indent_level, font_size} objects",
                        "note": (
                            "One array item becomes one paragraph (extra CRs inside an "
                            "item add paragraphs). 'level' is the PowerPoint IndentLevel (1-8)."
                        ),
                    },
                    "dry_run": {"type": "boolean", "required": False, "note": "Dry-run previews the full text and the planned levels/sizes."},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.text.set_format",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Set font properties and the paragraph IndentLevel on one shape's "
                    "text, limited to a paragraph range when 'paragraph' is given "
                    "(\"4\", \"4-10\", or a list). This is the curated way to give a list "
                    "levels and per-level sizes; dry-run reports the current values of "
                    "every targeted block."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "shape": {"type": "number|string", "required": True, "note": "1-based shape index or shape name."},
                    "paragraph": {"type": "number|string", "required": False, "note": "1-based paragraph number, \"first-last\" range, or a list; omit for the whole text frame."},
                    "font_size": {"type": "number", "required": False, "note": "Point size 1-400."},
                    "font_name": {"type": "string", "required": False},
                    "color": {"type": "string", "required": False, "note": "RRGGBB hex; converted to Office's BGR integer."},
                    "bold": {"type": "boolean", "required": False},
                    "indent_level": {"type": "number", "required": False, "note": "Paragraph level 1-8 (PowerPoint IndentLevel)."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.text.set_autofit",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Set a shape's autofit behaviour: none, shape (grow the frame to fit "
                    "the text) or text (the UI's 'Shrink text on overflow'). TextFrame2."
                    "AutoSize is tried first because it is the only member whose enum "
                    "(MsoAutoSize) contains text_fits_shape; a type-library cast of "
                    "TextFrame to TextFrame2 is tried next for builds where dynamic "
                    "dispatch will not name the member, and TextFrame.AutoSize "
                    "(PpAutoSize) last for the behaviours it can express. Every attempt "
                    "is reported and the result is read back; a refusal is "
                    "AUTOFIT_REFUSED with both enum tables in details."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "shape": {"type": "number|string", "required": True, "note": "1-based shape index or shape name."},
                    "behavior": {"type": "string", "required": True, "values": ["none", "shape", "text"]},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.text.measure",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Measure whether text fits a frame: the shape is duplicated, the "
                    "duplicate is switched to 'shape fits text' so PowerPoint lays the "
                    "text out itself, its height is read as required_height, and the "
                    "duplicate is deleted. Pass text and font_size to test a draft or a "
                    "point size before writing it. Reported in points, with fits and "
                    "margin. The duplicate changes the shape count twice, so this is a "
                    "confirm-required write even though the deck ends unchanged."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "1-based slide number."},
                    "shape": {"type": "number|string", "required": True, "note": "1-based shape index or shape name."},
                    "text": {"type": "string", "required": False, "note": "Draft text to measure; omit to measure the shape's current text."},
                    "font_size": {"type": "number", "required": False, "note": "Optional point size to test."},
                    "dry_run": {"type": "boolean", "required": False, "note": "Reports the frame and text without duplicating."},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "powerpoint.toc.fill",
                "applications": ["PowerPoint"],
                "mode": "write",
                "description": (
                    "Fill a table-of-contents placeholder with the deck's slide titles: "
                    "collects the titles of the slide range (the TOC slide itself is "
                    "skipped), keeps the numbering style already present in the "
                    "placeholder (full-width digits and a full-width stop, for example, "
                    "including its starting number), numbers only level-1 entries, and "
                    "writes everything through the paragraph-aware text write so levels "
                    "and per-level sizes land in one step. Dry-run previews the exact "
                    "final text plus the slide-to-paragraph mapping."
                ),
                "requires_confirmation": True,
                "params": {
                    "slide": {"type": "number", "required": True, "note": "The TOC slide holding the placeholder."},
                    "shape": {"type": "number|string", "required": True, "note": "1-based shape index or shape name of the body placeholder."},
                    "from": {"type": "number", "required": False, "note": "First slide to collect titles from; default 1."},
                    "to": {"type": "number", "required": False, "note": "Last slide to collect; default the last slide."},
                    "exclude": {"type": "array|csv", "required": False, "note": "Slide numbers to skip, in addition to the TOC slide itself."},
                    "levels": {"type": "string|object", "required": False, "note": "\"auto\" (a title listed in another slide's body text becomes level 2), or a spec such as \"6-12=2,15=2\"."},
                    "numbering": {"type": "string", "required": False, "values": ["auto", "fullwidth-dot", "ascii-dot", "digits", "none"], "note": "auto reuses the style detected in the existing text."},
                    "font_size": {"type": "number", "required": False, "note": "Point size for level-1 entries."},
                    "sub_font_size": {"type": "number", "required": False, "note": "Point size for level 2 and deeper."},
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
            {
                "name": "undo",
                "applications": ["Excel", "Word", "PowerPoint"],
                "mode": "write",
                "description": (
                    "Undo the most recent Office action (application.Undo). "
                    "Dry-run reports whether anything is undoable "
                    "(application.Undoable). Undoing reverts the user's own "
                    "actions too, so confirm the exact action with the user "
                    "before running. Per-app status on this build: PowerPoint "
                    "calls work, Word's Undo call itself raises an error for "
                    "any change (the error carries a hint to use the UI's "
                    "Ctrl+Z instead), Excel is untested — plan writes so they "
                    "do not depend on undo being possible."
                ),
                "requires_confirmation": True,
                "params": {
                    "dry_run": {"type": "boolean", "required": False},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            },
        ],
        "properties": [
            {
                "name": "excel.active_sheet",
                "applications": ["Excel"],
                "mode": "write",
                "description": "Activate a worksheet in the attached workbook.",
                "requires_confirmation": True,
                "params": {
                    "value": {"type": "string", "required": True},
                    "expected_document": {"type": "string", "required": False},
                    "confirm": {"type": "boolean", "required": True},
                },
            }
        ],
        "com": {
            "endpoint": "POST /com",
            "batch_endpoint": "POST /com/batch",
            "description": (
                "General Office object-model access. Reads any property, calls any "
                "method, and sets any property on the attached document. Use it for "
                "anything the curated commands above do not cover; save and close "
                "operations are blocked (see security.blocked_calls). The batch "
                "endpoint runs an array of steps in one request with one "
                "confirmation, stopping at the first failing step. A batch is not one "
                "undo action: the result reports undo_steps and an undo_note because "
                "Office records every applied step separately."
            ),
            "modes": [
                {
                    "name": "get",
                    "mode": "read",
                    "description": "Read a property. The expression must end with a property name (no arguments).",
                },
                {
                    "name": "call",
                    "mode": "write",
                    "description": "Call a method. The expression must end with (arguments). Supports dry_run.",
                    "requires_confirmation": True,
                },
                {
                    "name": "set",
                    "mode": "write",
                    "description": (
                        "Set a property from 'value'. The expression must end with a "
                        "property name (no arguments). Supports dry_run. Dry-run and "
                        "applied results both carry value_before when the current value "
                        "can be read, so a write can be reversed without a separate read."
                    ),
                    "requires_confirmation": True,
                },
            ],
            "roots": ["application", "window", "document"],
            "examples": {
                "get": 'window.ActiveSheet.Range("A1").Value2',
                "call": 'window.ActiveSheet.Range("2:4").Insert(-4107)',
                "set": 'window.Selection.Font.Bold = true',
            },
            "params": {
                "mode": {"type": "string", "required": True, "values": list(MODES)},
                "root": {"type": "string", "required": False, "values": list(ROOTS), "default": "window"},
                "expression": {"type": "string", "required": True},
                "value": {
                    "type": "any",
                    "required": "set mode only",
                    "file_reference": (
                        "CLI: --value @path reads the value from a UTF-8 file (a byte-order "
                        "mark is accepted). Prefer it for multi-line or Japanese text: "
                        "shell quoting mangles CR paragraph breaks and non-ASCII arguments."
                    ),
                },
                "dry_run": {"type": "boolean", "required": False},
                "expected_document": {"type": "string", "required": False},
                "confirm": {"type": "boolean", "required": True},
            },
            "file_arguments": {
                "com batch --file": (
                    "Path to a UTF-8 JSON array of steps "
                    '[{"mode": "get"|"call"|"set", "root": "window", "expression": "...", "value": ...}]. '
                    "A leading byte-order mark is stripped, so files written by Windows "
                    "PowerShell 5.1 (Set-Content -Encoding UTF8) work. This is the reliable "
                    "way to send CR-separated PowerPoint paragraphs: put them in the JSON as "
                    "\\r instead of building the argument in a shell."
                ),
                "com set --value": "@path reads the JSON or plain-text value from a UTF-8 file.",
                "powerpoint text set --paragraphs": "@path reads the paragraph array from a UTF-8 JSON file.",
                "powerpoint notes set --text / textbox add --text / text replace --replace": (
                    "@path reads the text from a UTF-8 file (kept as literal text, not parsed as JSON)."
                ),
            },
        },
    }
