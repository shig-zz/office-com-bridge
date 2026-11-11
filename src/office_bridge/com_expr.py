"""Parser for Office object-model expressions used by the /com endpoint.

An expression is a chain of member accesses starting from a root object:

    ActiveSheet.Range("A1:B2").Copy()
    Selection.Font.Bold
    Paragraphs(5).Range.Text

Each segment is either a property name or a method call ``name(args)``.
Arguments and assignment values are literals: strings (single or double
quoted), numbers, booleans, null, arrays, and objects.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .errors import BridgeError, invalid_request

_IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
_NUMBER_RE = re.compile(r"-?(?:\d+\.\d*|\.\d+|\d+)(?:[eE][+-]?\d+)?")
_ESCAPES = {
    "n": "\n",
    "t": "\t",
    "r": "\r",
    "v": "\v",
    "0": "\0",
    "\\": "\\",
    "/": "/",
    "'": "'",
    '"': '"',
}

# Unicode escape (\uXXXX) length, handled separately from the table above.
_UNICODE_ESCAPE_LENGTH = 4


@dataclass(frozen=True, slots=True)
class Segment:
    """One member access. args is None for a plain property read."""

    name: str
    args: tuple[Any, ...] | None


def _skip_ws(text: str, pos: int) -> int:
    while pos < len(text) and text[pos] in " \t\r\n":
        pos += 1
    return pos


def _fail(pos: int, message: str, extra: dict[str, Any] | None = None) -> BridgeError:
    details = {"position": pos, "message": message}
    if extra:
        details.update(extra)
    return invalid_request(f"Invalid expression: {message}", details)


def parse_expression(text: str) -> list[Segment]:
    """Parse a member chain such as ``ActiveSheet.Range("A1").Value2``."""
    if not isinstance(text, str) or not text.strip():
        raise invalid_request("The expression must be a non-empty string.")
    segments: list[Segment] = []
    pos = 0
    length = len(text)
    while True:
        pos = _skip_ws(text, pos)
        if pos >= length:
            break
        match = _IDENT_RE.match(text, pos)
        if not match:
            raise _fail(pos, f"expected a member name, found {text[pos]!r}")
        name = match.group(0)
        pos = _skip_ws(text, match.end())
        args: tuple[Any, ...] | None = None
        if pos < length and text[pos] == "(":
            args, pos = _parse_call_args(text, pos)
        segments.append(Segment(name=name, args=args))
        pos = _skip_ws(text, pos)
        if pos >= length:
            break
        if text[pos] != ".":
            raise _fail(pos, f"expected '.' between members, found {text[pos]!r}")
        pos += 1
    if not segments:
        raise invalid_request("The expression must contain at least one member name.")
    return segments


def _parse_call_args(text: str, pos: int) -> tuple[tuple[Any, ...], int]:
    assert text[pos] == "("
    pos = _skip_ws(text, pos + 1)
    args: list[Any] = []
    if pos < len(text) and text[pos] == ")":
        return (), pos + 1
    while True:
        value, pos = _parse_value(text, pos, allow_container=True)
        args.append(value)
        pos = _skip_ws(text, pos)
        if pos >= len(text):
            raise _fail(pos, "unterminated argument list: missing ')'")
        if text[pos] == ",":
            pos = _skip_ws(text, pos + 1)
            continue
        if text[pos] == ")":
            return tuple(args), pos + 1
        raise _fail(pos, f"expected ',' or ')' in argument list, found {text[pos]!r}")


def parse_value(text: str) -> Any:
    """Parse a literal value (string, number, boolean, null, array, object)."""
    if not isinstance(text, str) or not text.strip():
        raise invalid_request("The value must be a non-empty string.")
    value, pos = _parse_value(text, 0, allow_container=True)
    pos = _skip_ws(text, pos)
    if pos != len(text):
        raise _fail(pos, f"unexpected trailing characters in value: {text[pos]!r}")
    return value


def _parse_value(text: str, pos: int, *, allow_container: bool) -> tuple[Any, int]:
    pos = _skip_ws(text, pos)
    if pos >= len(text):
        raise _fail(pos, "expected a value")
    ch = text[pos]
    if ch in "\"'":
        return _parse_string(text, pos)
    if ch == "[":
        if not allow_container:
            raise _fail(pos, "arrays are not allowed here")
        return _parse_array(text, pos)
    if ch == "{":
        if not allow_container:
            raise _fail(pos, "objects are not allowed here")
        return _parse_object(text, pos)
    if ch.isdigit() or ch in "-+.":
        match = _NUMBER_RE.match(text, pos)
        if match:
            token = match.group(0)
            try:
                return int(token), match.end()
            except ValueError:
                return float(token), match.end()
    word = _IDENT_RE.match(text, pos)
    if word:
        lowered = word.group(0).lower()
        if lowered == "true":
            return True, word.end()
        if lowered == "false":
            return False, word.end()
        if lowered == "null":
            return None, word.end()
        raise _fail(pos, f"unexpected word {word.group(0)!r}; use quotes for strings")
    raise _fail(pos, f"expected a value, found {ch!r}")


def _parse_string(text: str, pos: int) -> tuple[str, int]:
    quote = text[pos]
    pos += 1
    parts: list[str] = []
    while True:
        if pos >= len(text):
            raise _fail(pos, "unterminated string literal")
        ch = text[pos]
        if ch == quote:
            return "".join(parts), pos + 1
        if ch == "\\":
            if pos + 1 >= len(text):
                raise _fail(pos, "unterminated escape in string literal")
            escape = text[pos + 1]
            if escape == "u":
                digits = text[pos + 2 : pos + 2 + _UNICODE_ESCAPE_LENGTH]
                if len(digits) < _UNICODE_ESCAPE_LENGTH or any(
                    digit not in "0123456789abcdefABCDEF" for digit in digits
                ):
                    raise _fail(
                        pos,
                        "\\u must be followed by four hex digits in string literal",
                        {"escape": f"\\u{digits}"},
                    )
                parts.append(chr(int(digits, 16)))
                pos += 2 + _UNICODE_ESCAPE_LENGTH
                continue
            if escape not in _ESCAPES:
                raise _fail(pos, f"unsupported escape '\\{escape}' in string literal")
            parts.append(_ESCAPES[escape])
            pos += 2
            continue
        parts.append(ch)
        pos += 1


def _parse_array(text: str, pos: int) -> tuple[list[Any], int]:
    assert text[pos] == "["
    pos = _skip_ws(text, pos + 1)
    items: list[Any] = []
    if pos < len(text) and text[pos] == "]":
        return [], pos + 1
    while True:
        value, pos = _parse_value(text, pos, allow_container=True)
        items.append(value)
        pos = _skip_ws(text, pos)
        if pos >= len(text):
            raise _fail(pos, "unterminated array: missing ']'")
        if text[pos] == ",":
            pos = _skip_ws(text, pos + 1)
            continue
        if text[pos] == "]":
            return items, pos + 1
        raise _fail(pos, f"expected ',' or ']' in array, found {text[pos]!r}")


def _parse_object(text: str, pos: int) -> tuple[dict[str, Any], int]:
    assert text[pos] == "{"
    pos = _skip_ws(text, pos + 1)
    result: dict[str, Any] = {}
    if pos < len(text) and text[pos] == "}":
        return {}, pos + 1
    while True:
        pos = _skip_ws(text, pos)
        if pos >= len(text):
            raise _fail(pos, "unterminated object: missing '}'")
        key, pos = _parse_string(text, pos)
        pos = _skip_ws(text, pos)
        if pos >= len(text) or text[pos] != ":":
            raise _fail(pos, "expected ':' after object key")
        value, pos = _parse_value(text, pos + 1, allow_container=True)
        result[key] = value
        pos = _skip_ws(text, pos)
        if pos >= len(text):
            raise _fail(pos, "unterminated object: missing '}'")
        if text[pos] == ",":
            pos += 1
            continue
        if text[pos] == "}":
            return result, pos + 1
        raise _fail(pos, f"expected ',' or '}}' in object, found {text[pos]!r}")
