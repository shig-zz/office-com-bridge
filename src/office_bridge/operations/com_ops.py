"""General Office object-model operations (POST /com).

Any property can be read, any method can be called, and any property can be
set on the attached document, subject to the blocklist. The blocklist
intentionally covers only the save and close families (plus quit/exit):
member names starting with "save" or "close", plus "quit" and "exit", are
rejected for method calls and property sets. Property reads are never
blocked. Additional names can be blocked with OFFICE_BRIDGE_BLOCKED_CALLS.
"""

from __future__ import annotations

import difflib
from typing import Any

import pywintypes

from ..com_expr import Segment
from ..config import Settings
from ..errors import (
    BridgeError,
    ErrorCode,
    com_error,
    invalid_request,
    operation_blocked,
)
from ..target import AppKind, ResolvedTarget
from .common import (
    check_expected_document,
    jsonable,
    require_confirm,
    truncate_text,
    unsaved_changes,
)

MODES = ("get", "call", "set")
ROOTS = ("application", "window", "document")

# Methods whose first two arguments are Before/After sheet references.
_SHEET_PLACEMENT_METHODS = {"add", "copy", "move"}


def _com_value(value: Any) -> Any:
    """Normalise a value for COM property assignment.

    A two-dimensional array of sequences (list or tuple of list/tuple) is passed to COM as the
    tuple-of-tuples shape the object model expects (as used by excel.range.write). Scalars,
    booleans and plain numbers are returned unchanged; array-shaped values for Range
    targets are validated and reshaped first by _normalize_array_value.
    """
    if isinstance(value, (list, tuple)) and value and all(isinstance(row, (list, tuple)) for row in value):
        return tuple(tuple(row) for row in value)
    return value


def _range_dimensions(owner: Any) -> tuple[int, int] | None:
    """Return the (rows, columns) of a Range-shaped object, else None."""
    try:
        rows = int(owner.Rows.Count)
        columns = int(owner.Columns.Count)
    except Exception:
        return None
    if rows < 1 or columns < 1:
        return None
    return rows, columns


def _normalize_array_value(owner: Any, name: str, value: Any) -> Any:
    """Shape array values to the Range that receives them before a COM set.

    Excel does not reject a mismatched Variant array: it coerces it. A flat
    list assigned to a multi-cell Range becomes a one-row array, the leading
    cells receive the items, and the remaining cells are cleared silently.
    When the assignment target is a Range, validate the array against the
    Range's dimensions and convert matching flat lists (a single row's or a
    single column's worth of values) into the two-dimensional shape Excel
    expects, so partial silent writes cannot happen.
    """
    if not isinstance(value, (list, tuple)) or not value:
        return value
    dims = _range_dimensions(owner)
    if dims is None:
        return value
    rows, columns = dims
    if all(isinstance(item, (list, tuple)) for item in value):
        widths = sorted({len(item) for item in value})
        if len(value) == rows and widths == [columns]:
            if rows == 1 and columns == 1:
                return value[0][0]
            return tuple(tuple(row) for row in value)
        raise BridgeError(
            ErrorCode.SIZE_MISMATCH,
            f"The array assigned to {name!r} is {len(value)}x{'/'.join(str(w) for w in widths)} "
            f"but the range is {rows}x{columns}. Pass an array that matches the range.",
            400,
            {
                "property": name,
                "range_rows": rows,
                "range_columns": columns,
                "array_rows": len(value),
                "array_widths": widths,
            },
        )
    if len(value) == 1:
        # A single item means "this value everywhere"; Excel's scalar
        # assignment fills every cell of the range.
        return value[0]
    if rows == 1 and len(value) == columns:
        return (tuple(value),)
    if columns == 1 and len(value) == rows:
        return tuple((item,) for item in value)
    raise invalid_request(
        f"The flat array assigned to {name!r} has {len(value)} items but the range is "
        f"{rows}x{columns}. Excel would write only the leading cells and clear the rest. "
        "Pass a two-dimensional array of rows x columns, or a flat array of exactly one "
        "row's or one column's worth of values.",
        {
            "property": name,
            "items": len(value),
            "range_rows": rows,
            "range_columns": columns,
        },
    )

_BLOCKED_PREFIXES = ("save", "close")
_BLOCKED_NAMES = ("quit", "exit")

# IDispatch error codes for an unknown member name.
_UNKNOWN_MEMBER_HRESULTS = {0x8002000B, 0x80020009, 0x80020027}


def blocked_policy(settings: Settings) -> dict[str, Any]:
    extra = tuple(
        dict.fromkeys(
            part.strip().lower()
            for part in (settings.blocked_calls or "").split(",")
            if part.strip()
        )
    )
    return {
        "prefixes": list(_BLOCKED_PREFIXES),
        "names": sorted(set(_BLOCKED_NAMES) | set(extra)),
    }


def is_blocked(name: str, policy: dict[str, Any]) -> bool:
    lowered = name.lower()
    if lowered.startswith(tuple(policy["prefixes"])):
        return True
    return lowered in policy["names"]


def blocked_summary(settings: Settings) -> dict[str, Any]:
    policy = blocked_policy(settings)
    return {
        "prefixes": policy["prefixes"],
        "names": policy["names"],
        "note": (
            "Member names starting with 'save' or 'close', plus 'quit' and 'exit' "
            "(and any names added via OFFICE_BRIDGE_BLOCKED_CALLS), are rejected "
            "for method calls and property sets. Property reads are never blocked."
        ),
    }


def execute(
    target: ResolvedTarget,
    mode: str,
    root: str,
    segments: list[Segment],
    value: Any,
    dry_run: bool,
    params: dict[str, Any],
    settings: Settings,
) -> dict[str, Any]:
    if mode not in MODES:
        raise invalid_request(f"Mode must be one of: {', '.join(MODES)}.", {"mode": mode})
    if root not in ROOTS:
        raise invalid_request(f"Root must be one of: {', '.join(ROOTS)}.", {"root": root})
    if not segments:
        raise invalid_request("The expression must contain at least one member name.")

    last = segments[-1]
    if mode == "call" and last.args is None:
        raise invalid_request(
            "A 'call' expression must end with a method call, for example Range(\"A1\").Copy().",
            {"expression": _render(segments)},
        )
    if mode in ("get", "set") and last.args is not None:
        raise invalid_request(
            f"A {mode!r} expression must end with a property name (no call arguments). "
            "To read a parameterized property, end the expression with the property "
            "name it exposes, or use 'com call' with the arguments.",
            {"expression": _render(segments)},
        )
    if mode == "set" and value is None:
        raise invalid_request("Mode 'set' requires a value.")
    root, segments = _resolve_segments(root, segments)
    last = segments[-1]
    last_args: tuple[Any, ...] = last.args if last.args is not None else ()

    policy = blocked_policy(settings)
    for segment in segments:
        if segment.args is not None and is_blocked(segment.name, policy):
            raise operation_blocked(segment.name)
    if mode in ("call", "set") and is_blocked(last.name, policy):
        raise operation_blocked(last.name)

    check_expected_document(target, params)
    root_object = _resolve_root(target, root)

    if mode == "get":
        result = _walk(root_object, segments)
        truncated = False
        note: str | None = None
        if callable(result):
            note = (
                "The value is a callable member (some Office members are parameterized "
                "properties). Read it with a 'call' expression, for example "
                f"'{_render(segments)}(...)', or use the property it exposes."
            )
        if isinstance(result, str) and settings.max_text_chars > 0:
            result, truncated = truncate_text(result, settings.max_text_chars)
        payload = {
            "mode": "get",
            "root": root,
            "value": jsonable(result),
            "truncated": truncated,
        }
        if note:
            payload["note"] = note
        return payload

    if not dry_run:
        require_confirm(params, f"com.{mode}")

    if mode == "call":
        owner = _walk(root_object, segments[:-1])
        member = _member(owner, last.name)
        if not callable(member):
            raise invalid_request(f"Member {last.name!r} is not a method.")
        call_args, resolved_notes = _resolve_placement_arguments(target, owner, last)
        call_args, anchor_notes = _resolve_hyperlink_anchor(target, segments, owner, last, call_args)
        resolved_notes = resolved_notes + anchor_notes
        if dry_run:
            preview = {
                "mode": "call",
                "root": root,
                "dry_run": True,
                "executed": False,
                "method": last.name,
                "arguments": [jsonable(item) for item in last_args],
            }
            if resolved_notes:
                preview["resolved_arguments"] = resolved_notes
            return preview
        result = _call(member, Segment(name=last.name, args=call_args))
        payload = {
            "mode": "call",
            "root": root,
            "dry_run": False,
            "executed": True,
            "method": last.name,
            "arguments": [jsonable(item) for item in last_args],
            "result": jsonable(result),
            "result_type": _result_type(result),
            "unsaved_changes": unsaved_changes(target),
        }
        if resolved_notes:
            payload["resolved_arguments"] = resolved_notes
        return payload

    owner = _walk(root_object, segments[:-1])
    value = _normalize_array_value(owner, last.name, value)
    com_value = _com_value(value)
    if dry_run:
        member = _member(owner, last.name)
        preview = {
            "mode": "set",
            "root": root,
            "dry_run": True,
            "applied": False,
            "property": last.name,
            "value": jsonable(com_value),
        }
        if callable(member):
            preview["applied_via"] = "parameterized_property"
        return preview
    applied_via = "property"
    try:
        setattr(owner, last.name, com_value)
    except (AttributeError, pywintypes.com_error) as exc:
        if not _is_unknown_member_error(exc):
            raise com_error(f"The COM property {last.name!r} could not be set.", exc) from exc
        # The name is not reachable as a plain property; some Office members are
        # parameterized properties and are set by calling them with the value.
        member = _member(owner, last.name)
        if not callable(member):
            raise com_error(f"The COM property {last.name!r} could not be set.", exc) from exc
        try:
            member(com_value)
        except Exception as call_exc:
            raise com_error(
                f"The COM property {last.name!r} could not be set.", call_exc
            ) from call_exc
        applied_via = "parameterized_property"
    except Exception as exc:
        raise com_error(f"The COM property {last.name!r} could not be set.", exc) from exc
    return {
        "mode": "set",
        "root": root,
        "dry_run": False,
        "applied": True,
        "property": last.name,
        "value": jsonable(com_value),
        "applied_via": applied_via,
        "unsaved_changes": unsaved_changes(target),
    }


def _render(segments: list[Segment]) -> str:
    parts: list[str] = []
    for segment in segments:
        if segment.args is None:
            parts.append(segment.name)
        else:
            parts.append(f"{segment.name}({_render_args(segment.args)})")
    return ".".join(parts)


def _render_args(args: tuple[Any, ...]) -> str:
    return ", ".join(repr(item) for item in args)


def _resolve_segments(root: str, segments: list[Segment]) -> tuple[str, list[Segment]]:
    """Allow the expression itself to name the root, for example 'window.Selection'.

    A leading root name is used as the root (and dropped) unless the caller
    explicitly requested a different one.
    """
    if segments and segments[0].name.lower() in ROOTS:
        expression_root = segments[0].name.lower()
        rest = segments[1:]
        if not rest:
            raise invalid_request(
                "The expression names only a root; continue it with a member, for example window.Selection.",
                {"expression": _render(segments)},
            )
        if root.lower() != "window" and root.lower() != expression_root:
            raise invalid_request(
                f"The expression starts with root {expression_root!r} but root {root!r} was requested.",
                {"root": root, "expression_root": expression_root},
            )
        return expression_root, rest
    return root, segments


def _resolve_root(target: ResolvedTarget, root: str) -> Any:
    if root == "application":
        return target.application
    if root == "window":
        if target.window is None:
            raise BridgeError(
                ErrorCode.NO_DOCUMENT,
                "The attached target has no document window; use root 'application' "
                "or 'document' instead of 'window'.",
                409,
            )
        return target.window
    return target.document


def _walk(root_object: Any, segments: list[Segment]) -> Any:
    current = root_object
    for segment in segments:
        current = _member(current, segment.name)
        if segment.args is not None:
            current = _call(current, segment)
    return current


def _member(owner: Any, name: str) -> Any:
    try:
        return getattr(owner, name)
    except AttributeError as exc:
        raise _unknown_member(owner, name, exc) from exc
    except pywintypes.com_error as exc:
        if _is_unknown_member_error(exc):
            raise _unknown_member(owner, name, exc) from exc
        raise com_error(f"The COM member {name!r} could not be read.", exc) from exc
    except Exception as exc:
        raise com_error(f"The COM member {name!r} could not be read.", exc) from exc


def _is_unknown_member_error(exc: BaseException) -> bool:
    if isinstance(exc, AttributeError):
        return True
    hresult = int(getattr(exc, "hresult", 0) or 0) & 0xFFFFFFFF
    return hresult in _UNKNOWN_MEMBER_HRESULTS


def _is_callable_member(owner: Any, name: str) -> bool:
    try:
        return callable(getattr(owner, name))
    except Exception:
        return False


def _result_type(result: Any) -> str:
    """Name the returned type so a stringified COM object cannot be mistaken
    for a scalar of the same printed form (for example TextRange.Replace)."""
    return type(result).__name__


def _member_candidates(owner: Any, name: str, limit: int = 12) -> list[str]:
    names: list[str] = []
    try:
        names.extend(str(item) for item in dir(owner))
    except Exception:
        pass
    for attribute in ("_propnames", "_methods_"):
        try:
            names.extend(str(item) for item in getattr(owner, attribute) or ())
        except Exception:
            continue
    unique = sorted({item for item in names if item and not item.startswith("_")})
    if not unique:
        return []
    close = difflib.get_close_matches(name, unique, n=limit, cutoff=0.5)
    if len(close) < limit:
        prefix = name.lower()[: max(2, len(name) - 1)]
        close.extend(item for item in unique if item.lower().startswith(prefix) and item not in close)
    return close[:limit]


def _member_hint(owner: Any, name: str) -> str | None:
    lowered = name.lower()
    if lowered == "text" and _is_callable_member(owner, "TextRange"):
        return (
            "This object exposes TextRange: read or write TextRange.Text instead. "
            "On some Office objects the plain name is a parameterized property and is "
            "not reachable as a property."
        )
    return (
        "Some Office members are parameterized properties: read them with a 'call' "
        "expression that passes the arguments, or use the property they expose."
    )


def _unknown_member(owner: Any, name: str, exc: BaseException) -> BridgeError:
    details: dict[str, Any] = {"member": name, "error": str(exc)}
    candidates = _member_candidates(owner, name)
    if candidates:
        details["did_you_mean"] = candidates
    hint = _member_hint(owner, name)
    if hint:
        details["hint"] = hint
    return invalid_request(
        f"Member {name!r} was not found on the target object.", details
    )


def _call(member: Any, segment: Segment) -> Any:
    args = segment.args if segment.args is not None else ()
    try:
        return member(*args)
    except pywintypes.com_error as exc:
        raise com_error(f"The COM call {segment.name!r} failed.", exc) from exc
    except Exception as exc:
        raise com_error(f"The COM call {segment.name!r} failed.", exc) from exc


def _placement_sheets(owner: Any) -> Any | None:
    """Return the workbook's Worksheets collection when the call owner lives in one.

    A Worksheet (Copy/Move) and a Worksheets collection (Add) both expose a
    parent workbook. Ranges expose a Worksheet parent, which has no
    Worksheets member, so they are excluded.
    """
    try:
        sheets = owner.Parent.Worksheets
        sheets.Count
    except Exception:
        return None
    return sheets


def _resolve_placement_arguments(
    target: ResolvedTarget, owner: Any, last: Segment
) -> tuple[tuple[Any, ...], list[dict[str, Any]]]:
    """Resolve Add/Copy/Move Before/After literals into real sheet objects.

    Excel's Before/After parameters are Variants that must wrap an actual
    Worksheet object; a raw sheet index or sheet name passed through COM is
    rejected. Because the bridge holds the owning workbook during the walk,
    an int (1-based sheet index) or a sheet-name string in the first two
    argument positions is resolved to the live sheet object before the call.
    Unresolvable values are passed through unchanged so Excel reports the
    failure.
    """
    args = last.args if last.args is not None else ()
    if target.app is not AppKind.EXCEL or last.name.lower() not in _SHEET_PLACEMENT_METHODS:
        return args, []
    sheets = _placement_sheets(owner)
    if sheets is None:
        return args, []
    resolved: list[Any] = []
    notes: list[dict[str, Any]] = []
    for position, arg in enumerate(args):
        sheet: Any = None
        if position < 2 and isinstance(arg, (int, str)) and not isinstance(arg, bool):
            try:
                if not isinstance(arg, int) or 1 <= arg <= sheets.Count:
                    sheet = sheets(arg)
            except Exception:
                sheet = None
        if sheet is None:
            resolved.append(arg)
            continue
        resolved.append(sheet)
        try:
            sheet_name = jsonable(sheet.Name)
        except Exception:
            sheet_name = None
        notes.append({"position": position, "argument": jsonable(arg), "sheet": sheet_name})
    return tuple(resolved), notes


def _resolve_hyperlink_anchor(
    target: ResolvedTarget,
    segments: list[Segment],
    owner: Any,
    last: Segment,
    args: tuple[Any, ...],
) -> tuple[tuple[Any, ...], list[dict[str, Any]]]:
    """Give Excel's Hyperlinks.Add the Anchor object its caller cannot pass.

    Hyperlinks.Add takes the linked Range as its first (Anchor) parameter, but
    com call arguments are JSON literals, so a Range object can never reach
    it and the call fails with a type mismatch. The Hyperlinks collection's
    Parent is exactly the Range that owns it, so for a call walked through a
    'Hyperlinks.Add' on a Range-scoped collection the anchor is prepended.
    Collections whose parent is not a Range (for example a Worksheet's
    Hyperlinks) are left alone so Excel reports the missing anchor itself.
    """
    if (
        target.app is not AppKind.EXCEL
        or len(segments) < 2
        or segments[-2].name.lower() != "hyperlinks"
        or last.name.lower() != "add"
    ):
        return args, []
    try:
        anchor = owner.Parent
        anchor_address = str(anchor.Address)
    except Exception:
        return args, []
    note: dict[str, Any] = {"position": 0, "anchor": anchor_address}
    try:
        note["sheet"] = jsonable(anchor.Parent.Name)
    except Exception:
        pass
    return (anchor,) + args, [note]
