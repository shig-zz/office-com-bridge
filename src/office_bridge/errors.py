from __future__ import annotations

from dataclasses import dataclass
from typing import Any


class ErrorCode:
    INVALID_REQUEST = "INVALID_REQUEST"
    NOT_OFFICE_WINDOW = "NOT_OFFICE_WINDOW"
    NO_DOCUMENT = "NO_DOCUMENT"
    NO_ACTIVE_SLIDE = "NO_ACTIVE_SLIDE"
    TARGET_NOT_SET = "TARGET_NOT_SET"
    TARGET_NOT_FOUND = "TARGET_NOT_FOUND"
    TARGET_LOST = "TARGET_LOST"
    DOCUMENT_MISMATCH = "DOCUMENT_MISMATCH"
    RESOLVE_FAILED = "RESOLVE_FAILED"
    INSTANCE_MISMATCH = "INSTANCE_MISMATCH"
    UNSUPPORTED_OPERATION = "UNSUPPORTED_OPERATION"
    UNSUPPORTED_APPLICATION = "UNSUPPORTED_APPLICATION"
    OPERATION_BLOCKED = "OPERATION_BLOCKED"
    RANGE_INVALID = "RANGE_INVALID"
    RANGE_TOO_LARGE = "RANGE_TOO_LARGE"
    SIZE_MISMATCH = "SIZE_MISMATCH"
    TEXT_TOO_LARGE = "TEXT_TOO_LARGE"
    CONFIRMATION_REQUIRED = "CONFIRMATION_REQUIRED"
    COM_ERROR = "COM_ERROR"
    OFFICE_BUSY = "OFFICE_BUSY"
    OFFICE_NO_RESPONSE = "OFFICE_NO_RESPONSE"
    BRIDGE_NOT_RUNNING = "BRIDGE_NOT_RUNNING"
    BRIDGE_UNREACHABLE = "BRIDGE_UNREACHABLE"
    INTERNAL_ERROR = "INTERNAL_ERROR"


@dataclass(slots=True)
class BridgeError(Exception):
    code: str
    message: str
    status_code: int = 400
    details: dict[str, Any] | None = None

    def __str__(self) -> str:
        return f"{self.code}: {self.message}"

    def to_dict(self) -> dict[str, Any]:
        error: dict[str, Any] = {"code": self.code, "message": self.message}
        if self.details:
            error["details"] = self.details
        return error


def invalid_request(message: str, details: dict[str, Any] | None = None) -> BridgeError:
    return BridgeError(ErrorCode.INVALID_REQUEST, message, 400, details)


def not_office_window(message: str, details: dict[str, Any] | None = None) -> BridgeError:
    return BridgeError(ErrorCode.NOT_OFFICE_WINDOW, message, 409, details)


def no_document(message: str, details: dict[str, Any] | None = None) -> BridgeError:
    return BridgeError(ErrorCode.NO_DOCUMENT, message, 409, details)


def target_not_set() -> BridgeError:
    return BridgeError(
        ErrorCode.TARGET_NOT_SET,
        "No Office target is selected. Focus the Office document window and press the bridge hotkey.",
        409,
    )


def target_lost(
    message: str = "The selected Office document is no longer available.",
    details: dict[str, Any] | None = None,
) -> BridgeError:
    return BridgeError(ErrorCode.TARGET_LOST, message, 409, details)


def target_not_found(target_id: str) -> BridgeError:
    return BridgeError(
        ErrorCode.TARGET_NOT_FOUND,
        f"No registered target matches id {target_id!r}. Run 'office target list' to see registered documents.",
        404,
        {"target_id": target_id},
    )


def document_mismatch(expected: str, actual: str) -> BridgeError:
    return BridgeError(
        ErrorCode.DOCUMENT_MISMATCH,
        f"Target document mismatch: expected {expected!r} but the bridge is attached to {actual!r}.",
        409,
        {"expected": expected, "actual": actual},
    )


def resolve_failed(message: str, details: dict[str, Any] | None = None) -> BridgeError:
    return BridgeError(ErrorCode.RESOLVE_FAILED, message, 409, details)


def instance_mismatch(message: str, details: dict[str, Any] | None = None) -> BridgeError:
    return BridgeError(ErrorCode.INSTANCE_MISMATCH, message, 409, details)


def unsupported_operation(name: str) -> BridgeError:
    return BridgeError(
        ErrorCode.UNSUPPORTED_OPERATION,
        f"Operation {name!r} is not allowed.",
        400,
        {"operation": name},
    )


def operation_blocked(name: str) -> BridgeError:
    return BridgeError(
        ErrorCode.OPERATION_BLOCKED,
        f"The operation {name!r} is blocked by the bridge. Saving and closing are "
        "intentionally unavailable; ask the user to perform them in Office.",
        403,
        {"operation": name},
    )


def unsupported_application(app: str, operation: str) -> BridgeError:
    return BridgeError(
        ErrorCode.UNSUPPORTED_APPLICATION,
        f"Operation {operation!r} is not supported for {app}.",
        400,
        {"application": app, "operation": operation},
    )


def range_too_large(count: int, limit: int) -> BridgeError:
    return BridgeError(
        ErrorCode.RANGE_TOO_LARGE,
        f"Range has {count} cells, which exceeds the limit of {limit}.",
        413,
        {"cells": count, "limit": limit},
    )


def text_too_large(count: int, limit: int) -> BridgeError:
    return BridgeError(
        ErrorCode.TEXT_TOO_LARGE,
        f"Text has {count} characters, which exceeds the limit of {limit}.",
        413,
        {"characters": count, "limit": limit},
    )


def confirmation_required(operation: str) -> BridgeError:
    return BridgeError(
        ErrorCode.CONFIRMATION_REQUIRED,
        f"Operation {operation!r} modifies a document and requires confirm=true.",
        409,
        {"operation": operation},
    )


def com_error(message: str, exc: BaseException) -> BridgeError:
    return BridgeError(
        ErrorCode.COM_ERROR,
        message,
        502,
        _exception_details(exc),
    )


def office_busy(message: str, exc: BaseException | None = None) -> BridgeError:
    details = _exception_details(exc) if exc is not None else None
    return BridgeError(ErrorCode.OFFICE_BUSY, message, 503, details)


def office_no_response(timeout: float) -> BridgeError:
    return BridgeError(
        ErrorCode.OFFICE_NO_RESPONSE,
        "The Office COM worker did not respond in time. Office may be busy; do not retry a write blindly.",
        504,
        {"timeout_seconds": timeout},
    )


def bridge_not_running() -> BridgeError:
    return BridgeError(ErrorCode.BRIDGE_NOT_RUNNING, "The COM worker thread is not running.", 503)


def internal_error(message: str, details: dict[str, Any] | None = None) -> BridgeError:
    return BridgeError(ErrorCode.INTERNAL_ERROR, message, 500, details)


def _exception_details(exc: BaseException | None) -> dict[str, Any] | None:
    if exc is None:
        return None
    details: dict[str, Any] = {"type": type(exc).__name__, "message": str(exc)}
    hresult = getattr(exc, "hresult", None)
    if hresult is not None:
        try:
            details["hresult"] = f"0x{int(hresult) & 0xFFFFFFFF:08X}"
        except (TypeError, ValueError):
            pass
    return details
