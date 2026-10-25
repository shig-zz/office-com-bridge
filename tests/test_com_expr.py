from __future__ import annotations

import pytest

from office_bridge.com_expr import parse_expression, parse_value
from office_bridge.errors import BridgeError, ErrorCode


def test_parse_simple_property_chain() -> None:
    segments = parse_expression("Selection.Font.Bold")
    assert [(s.name, s.args) for s in segments] == [
        ("Selection", None),
        ("Font", None),
        ("Bold", None),
    ]


def test_parse_call_arguments() -> None:
    segments = parse_expression('ActiveSheet.Range("A1:B2").Copy()')
    assert [(s.name, s.args) for s in segments] == [
        ("ActiveSheet", None),
        ("Range", ("A1:B2",)),
        ("Copy", ()),
    ]


def test_parse_single_quoted_and_multiple_args() -> None:
    segments = parse_expression("Cells(2, 1)")
    assert [(s.name, s.args) for s in segments] == [("Cells", (2, 1))]

    segments = parse_expression("Worksheets('Sheet1').Name")
    assert [(s.name, s.args) for s in segments] == [
        ("Worksheets", ("Sheet1",)),
        ("Name", None),
    ]


def test_parse_escapes() -> None:
    segments = parse_expression(r'TypeText("a\nb\\c\'d")')
    assert segments[0].args == ("a\nb\\c'd",)


def test_parse_rejects_empty() -> None:
    with pytest.raises(BridgeError) as excinfo:
        parse_expression("   ")
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_parse_rejects_bad_start() -> None:
    with pytest.raises(BridgeError) as excinfo:
        parse_expression(".Selection")
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_parse_rejects_unterminated_string() -> None:
    with pytest.raises(BridgeError) as excinfo:
        parse_expression('Range("A1')
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_parse_rejects_trailing_garbage() -> None:
    with pytest.raises(BridgeError) as excinfo:
        parse_expression("A.B = 1")
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_parse_rejects_unterminated_call() -> None:
    with pytest.raises(BridgeError) as excinfo:
        parse_expression("Copy(")
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_parse_value_scalars() -> None:
    assert parse_value("42") == 42
    assert parse_value("-1.5") == -1.5
    assert parse_value("true") is True
    assert parse_value("False") is False
    assert parse_value("null") is None
    assert parse_value('"text"') == "text"
    assert parse_value("'text with spaces'") == "text with spaces"


def test_parse_value_containers() -> None:
    assert parse_value("[[1, 2], [3, 4]]") == [[1, 2], [3, 4]]
    assert parse_value("[]") == []
    assert parse_value('{"a": 1, "b": [true, null]}') == {"a": 1, "b": [True, None]}


def test_parse_value_rejects_trailing_garbage() -> None:
    with pytest.raises(BridgeError) as excinfo:
        parse_value("42 x")
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_parse_value_rejects_bare_word() -> None:
    with pytest.raises(BridgeError) as excinfo:
        parse_value("hello")
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
