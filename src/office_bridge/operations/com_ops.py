"""General Office object-model operations (POST /com).

Any property can be read, any method can be called, and any property can be
set on the attached document, subject to the blocklist. The blocklist
intentionally covers only the save and close families (plus quit/exit):
member names starting with "save" or "close", plus "quit" and "exit", are
rejected for method calls and property sets. Property reads are never
blocked. Additional names can be blocked with OFFICE_BRIDGE_BLOCKED_CALLS.
"""

from __future__ import annotations

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
from ..target import ResolvedTarget
from .common import check_expected_document, jsonable, require_confirm, truncate_text

MODES = ("get", "call", "set")
ROOTS = ("application", "window", "document")

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
            f"A {mode!r} expression must end with a property name (no call arguments).",
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
        if isinstance(result, str) and settings.max_text_chars > 0:
            result, truncated = truncate_text(result, settings.max_text_chars)
        return {
            "mode": "get",
            "root": root,
            "value": jsonable(result),
            "truncated": truncated,
        }

    if not dry_run:
        require_confirm(params, f"com.{mode}")

    if mode == "call":
        if dry_run:
            owner = _walk(root_object, segments[:-1])
            member = _member(owner, last.name)
            if not callable(member):
                raise invalid_request(f"Member {last.name!r} is not a method.")
            return {
                "mode": "call",
                "root": root,
                "dry_run": True,
                "executed": False,
                "method": last.name,
                "arguments": [jsonable(item) for item in last_args],
            }
        result = _walk(root_object, segments)
        return {
            "mode": "call",
            "root": root,
            "dry_run": False,
            "executed": True,
            "method": last.name,
            "arguments": [jsonable(item) for item in last_args],
            "result": jsonable(result),
        }

    owner = _walk(root_object, segments[:-1])
    _member(owner, last.name)
    if dry_run:
        return {
            "mode": "set",
            "root": root,
            "dry_run": True,
            "applied": False,
            "property": last.name,
            "value": jsonable(value),
        }
    try:
        setattr(owner, last.name, value)
    except pywintypes.com_error as exc:
        raise com_error(f"The COM property {last.name!r} could not be set.", exc) from exc
    except Exception as exc:
        raise com_error(f"The COM property {last.name!r} could not be set.", exc) from exc
    return {
        "mode": "set",
        "root": root,
        "dry_run": False,
        "applied": True,
        "property": last.name,
        "value": jsonable(value),
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
        raise _unknown_member(name, exc) from exc
    except pywintypes.com_error as exc:
        hresult = int(getattr(exc, "hresult", 0) or 0) & 0xFFFFFFFF
        if hresult in _UNKNOWN_MEMBER_HRESULTS:
            raise _unknown_member(name, exc) from exc
        raise com_error(f"The COM member {name!r} could not be read.", exc) from exc
    except Exception as exc:
        raise com_error(f"The COM member {name!r} could not be read.", exc) from exc


def _call(member: Any, segment: Segment) -> Any:
    args = segment.args if segment.args is not None else ()
    try:
        return member(*args)
    except pywintypes.com_error as exc:
        raise com_error(f"The COM call {segment.name!r} failed.", exc) from exc
    except Exception as exc:
        raise com_error(f"The COM call {segment.name!r} failed.", exc) from exc


def _unknown_member(name: str, exc: BaseException) -> BridgeError:
    return invalid_request(
        f"Member {name!r} was not found on the target object.",
        {"member": name, "error": str(exc)},
    )
