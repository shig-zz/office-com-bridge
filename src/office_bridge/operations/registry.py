from __future__ import annotations

from typing import Any

from ..config import Settings
from ..errors import invalid_request, unsupported_application, unsupported_operation
from ..target import AppKind, ResolvedTarget
from . import com_ops, excel, powerpoint, word
from .com_ops import MODES, ROOTS
from .common import (
    document_result,
    optional_bool,
    require_str,
    selection_result,
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

    raise unsupported_operation(kind)


def invoke(
    target: ResolvedTarget,
    operation: str,
    params: dict[str, Any],
    settings: Settings,
) -> dict[str, Any]:
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
        return excel.set_active_sheet(target, value, params)
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
                "description": "Read values (and optionally formulas) from an A1-style range on the window's active sheet.",
                "params": {
                    "address": {"type": "string", "required": True},
                    "include_formulas": {"type": "boolean", "required": False},
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
                "name": "excel.range.copy_format",
                "applications": ["Excel"],
                "mode": "write",
                "description": "Copy formatting only (font, fill, borders, number format, alignment) from a source range to a target range of the same size. Cell values are not changed.",
                "requires_confirmation": True,
                "params": {
                    "source_address": {"type": "string", "required": True},
                    "address": {"type": "string", "required": True},
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
            "description": (
                "General Office object-model access. Reads any property, calls any "
                "method, and sets any property on the attached document. Use it for "
                "anything the curated commands above do not cover; save and close "
                "operations are blocked (see security.blocked_calls)."
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
                    "description": "Set a property from 'value'. The expression must end with a property name (no arguments). Supports dry_run.",
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
                "value": {"type": "any", "required": "set mode only"},
                "dry_run": {"type": "boolean", "required": False},
                "expected_document": {"type": "string", "required": False},
                "confirm": {"type": "boolean", "required": True},
            },
        },
    }
