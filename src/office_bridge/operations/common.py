from __future__ import annotations

import math
import os
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Iterable

from ..errors import (
    BridgeError,
    confirmation_required,
    document_mismatch,
    invalid_request,
    range_too_large,
    unsupported_application,
    unsupported_operation,
)
from ..target import AppKind, ResolvedTarget


def jsonable(value: Any) -> Any:
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        if math.isfinite(value):
            return value
        return None
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (list, tuple)):
        return [jsonable(item) for item in value]
    if isinstance(value, dict):
        return {str(key): jsonable(item) for key, item in value.items()}
    try:
        return str(value)
    except Exception:
        return None


def require_str(params: dict[str, Any], key: str, *, allow_empty: bool = False) -> str:
    value = params.get(key)
    if not isinstance(value, str):
        raise invalid_request(f"Parameter {key!r} must be a string.", {"parameter": key})
    if not allow_empty and not value.strip():
        raise invalid_request(f"Parameter {key!r} must not be empty.", {"parameter": key})
    return value


def optional_bool(params: dict[str, Any], key: str, default: bool = False) -> bool:
    value = params.get(key, default)
    if not isinstance(value, bool):
        raise invalid_request(f"Parameter {key!r} must be a boolean.", {"parameter": key})
    return value


def optional_int(params: dict[str, Any], key: str, default: int) -> int:
    value = params.get(key, default)
    if isinstance(value, bool) or not isinstance(value, int):
        raise invalid_request(f"Parameter {key!r} must be an integer.", {"parameter": key})
    return value


def require_confirm(params: dict[str, Any], operation: str) -> None:
    if params.get("confirm") is not True:
        raise confirmation_required(operation)


def check_expected_document(target: ResolvedTarget, params: dict[str, Any]) -> None:
    expected = params.get("expected_document")
    if expected is None:
        return
    if not isinstance(expected, str) or not expected.strip():
        raise invalid_request(
            "Parameter 'expected_document' must be a non-empty string.",
            {"parameter": "expected_document"},
        )

    candidates = {
        str(target.info.get("document", "")),
        str(target.info.get("full_name", "")),
    }
    full_name = str(target.info.get("full_name", ""))
    if full_name:
        candidates.add(os.path.basename(full_name))
    candidates = {candidate.lower() for candidate in candidates if candidate}
    if expected.strip().lower() not in candidates:
        raise document_mismatch(expected, str(target.info.get("document", "")))


def check_cell_limit(cells: int, limit: int) -> None:
    if cells > limit:
        raise range_too_large(cells, limit)


def truncate_text(text: str, limit: int) -> tuple[str, bool]:
    if limit <= 0 or len(text) <= limit:
        return text, False
    return text[:limit], True


def validate_matrix(values: Any) -> tuple[list[list[Any]], int, int]:
    if not isinstance(values, list) or not values:
        raise invalid_request("Parameter 'values' must be a non-empty two-dimensional array.")
    width: int | None = None
    normalized: list[list[Any]] = []
    for row_index, row in enumerate(values):
        if not isinstance(row, list) or not row:
            raise invalid_request(
                "Parameter 'values' must contain non-empty rows.",
                {"row": row_index},
            )
        if width is None:
            width = len(row)
        elif len(row) != width:
            raise invalid_request(
                "Parameter 'values' must be rectangular (all rows have the same length).",
                {"row": row_index, "expected_columns": width, "actual_columns": len(row)},
            )
        normalized_row: list[Any] = []
        for column_index, cell in enumerate(row):
            if cell is None or isinstance(cell, (str, bool, int)):
                normalized_row.append(cell)
            elif isinstance(cell, float):
                if not math.isfinite(cell):
                    raise invalid_request(
                        "Cell values must be finite numbers.",
                        {"row": row_index, "column": column_index},
                    )
                normalized_row.append(cell)
            else:
                raise invalid_request(
                    "Cell values must be strings, numbers, booleans, or null.",
                    {"row": row_index, "column": column_index, "type": type(cell).__name__},
                )
        normalized.append(normalized_row)
    return normalized, len(normalized), int(width or 0)


def selection_result(target: ResolvedTarget, *, limit: int, max_text_chars: int) -> dict[str, Any]:
    if target.app is AppKind.EXCEL:
        from . import excel

        return excel.selection(target, limit=limit)
    if target.app is AppKind.WORD:
        from . import word

        return word.selection(target, max_text_chars=max_text_chars)
    if target.app is AppKind.POWERPOINT:
        from . import powerpoint

        return powerpoint.slide(target, max_text_chars=max_text_chars)
    raise unsupported_application(target.app.value, "selection")


def document_result(target: ResolvedTarget, *, max_text_chars: int) -> dict[str, Any]:
    if target.app is AppKind.EXCEL:
        from . import excel

        return excel.document(target)
    if target.app is AppKind.WORD:
        from . import word

        return word.document(target)
    if target.app is AppKind.POWERPOINT:
        from . import powerpoint

        return powerpoint.document(target, max_text_chars=max_text_chars)
    raise unsupported_application(target.app.value, "document")


def unsupported(target: ResolvedTarget, operation: str) -> BridgeError:
    return unsupported_application(target.app.value, operation)


def dedupe_preserve_order(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result


def operation_placeholder(name: str) -> BridgeError:
    return unsupported_operation(name)
