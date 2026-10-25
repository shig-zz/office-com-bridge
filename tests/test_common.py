from __future__ import annotations

from datetime import datetime
from types import SimpleNamespace

import pytest

from office_bridge.errors import BridgeError, ErrorCode
from office_bridge.operations.common import (
    check_expected_document,
    jsonable,
    truncate_text,
    validate_matrix,
)
from office_bridge.target import AppKind


def make_target(document: str = "Book1.xlsx", full_name: str = r"C:\docs\Book1.xlsx") -> SimpleNamespace:
    return SimpleNamespace(
        app=AppKind.EXCEL,
        info={"application": "Excel", "document": document, "full_name": full_name},
    )


def test_jsonable_converts_tuples_and_datetimes() -> None:
    value = {"a": (1, 2), "b": (datetime(2026, 1, 2, 3, 4, 5),)}
    assert jsonable(value) == {"a": [1, 2], "b": ["2026-01-02T03:04:05"]}


def test_validate_matrix_accepts_rectangular_values() -> None:
    matrix, rows, columns = validate_matrix([[1, "two"], [True, None]])
    assert matrix == [[1, "two"], [True, None]]
    assert rows == 2
    assert columns == 2


def test_validate_matrix_rejects_ragged_arrays() -> None:
    with pytest.raises(BridgeError) as excinfo:
        validate_matrix([[1, 2], [3]])
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_validate_matrix_rejects_non_finite_numbers() -> None:
    with pytest.raises(BridgeError):
        validate_matrix([[float("nan")]])


def test_validate_matrix_rejects_nested_objects() -> None:
    with pytest.raises(BridgeError):
        validate_matrix([[{"x": 1}]])


def test_truncate_text_reports_truncation() -> None:
    assert truncate_text("abcdef", 3) == ("abc", True)
    assert truncate_text("abc", 3) == ("abc", False)


def test_expected_document_accepts_name_and_basename() -> None:
    target = make_target()
    check_expected_document(target, {"expected_document": "Book1.xlsx"})
    check_expected_document(target, {"expected_document": r"c:\DOCS\book1.XLSX"})


def test_expected_document_rejects_mismatch() -> None:
    with pytest.raises(BridgeError) as excinfo:
        check_expected_document(make_target(), {"expected_document": "Other.xlsx"})
    assert excinfo.value.code == ErrorCode.DOCUMENT_MISMATCH


def test_expected_document_requires_string() -> None:
    with pytest.raises(BridgeError):
        check_expected_document(make_target(), {"expected_document": 123})
