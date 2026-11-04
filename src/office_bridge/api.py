from __future__ import annotations

import logging
from typing import Any

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import JSONResponse
from pydantic import BaseModel, ConfigDict, Field
from starlette.exceptions import HTTPException as StarletteHTTPException

from . import __version__
from .com_worker import BridgeCore
from .config import Settings
from .errors import BridgeError, ErrorCode

logger = logging.getLogger(__name__)


class InspectRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: str = Field(min_length=1)
    params: dict[str, Any] = Field(default_factory=dict)


class InvokeRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    operation: str = Field(min_length=1)
    params: dict[str, Any] = Field(default_factory=dict)


class SetPropertyRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    property: str = Field(min_length=1)
    value: Any = None
    confirm: bool = False
    expected_document: str | None = None


class ComRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    mode: str = Field(min_length=1)
    root: str = Field(default="window", min_length=1)
    expression: str = Field(min_length=1)
    value: Any = None
    dry_run: bool = False
    confirm: bool = False
    expected_document: str | None = None


class CaptureRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    hwnd: int | None = None


def _envelope_success(result: Any, context: dict[str, Any] | None) -> dict[str, Any]:
    return {"ok": True, "context": context, "result": result}


def _envelope_error(error: BridgeError, context: dict[str, Any] | None) -> dict[str, Any]:
    return {"ok": False, "context": context, "error": error.to_dict()}


def create_app(core: BridgeCore, settings: Settings) -> FastAPI:
    app = FastAPI(
        title="Office COM Bridge",
        version=__version__,
        docs_url="/docs",
        redoc_url=None,
    )

    def context_snapshot() -> dict[str, Any] | None:
        try:
            return core.cached_summary()
        except Exception:
            return None

    def target_count() -> int:
        try:
            return len(core.cached_targets())
        except Exception:
            return 0

    @app.exception_handler(BridgeError)
    async def handle_bridge_error(_request: Request, exc: BridgeError) -> JSONResponse:
        return JSONResponse(
            status_code=exc.status_code,
            content=_envelope_error(exc, context_snapshot()),
        )

    @app.exception_handler(RequestValidationError)
    async def handle_validation_error(
        _request: Request, exc: RequestValidationError
    ) -> JSONResponse:
        errors = []
        for item in exc.errors():
            errors.append(
                {
                    "location": list(item.get("loc", [])),
                    "message": item.get("msg", ""),
                    "type": item.get("type", ""),
                }
            )
        error = BridgeError(
            ErrorCode.INVALID_REQUEST,
            "Request validation failed.",
            400,
            {"errors": errors},
        )
        return JSONResponse(
            status_code=400,
            content=_envelope_error(error, context_snapshot()),
        )

    @app.exception_handler(StarletteHTTPException)
    async def handle_http_error(
        _request: Request, exc: StarletteHTTPException
    ) -> JSONResponse:
        code = ErrorCode.INVALID_REQUEST if exc.status_code < 500 else ErrorCode.INTERNAL_ERROR
        error = BridgeError(code, str(exc.detail), exc.status_code)
        return JSONResponse(
            status_code=exc.status_code,
            content=_envelope_error(error, context_snapshot()),
        )

    @app.exception_handler(Exception)
    async def handle_unexpected_error(_request: Request, exc: Exception) -> JSONResponse:
        logger.exception("unhandled API error")
        error = BridgeError(
            ErrorCode.INTERNAL_ERROR,
            "An unexpected error occurred.",
            500,
            {"type": type(exc).__name__, "message": str(exc)},
        )
        return JSONResponse(
            status_code=500,
            content=_envelope_error(error, context_snapshot()),
        )

    @app.get("/health")
    def health() -> dict[str, Any]:
        return _envelope_success(
            {
                "status": "running",
                "version": __version__,
                "hotkey": settings.hotkey if settings.hotkey_enabled else None,
                "hotkey_enabled": settings.hotkey_enabled,
                "target": context_snapshot(),
                "target_count": target_count(),
                "worker_alive": bool(core.worker.is_alive),
            },
            context_snapshot(),
        )

    @app.get("/context")
    def context() -> dict[str, Any]:
        info = core.context()
        return _envelope_success(info, context_snapshot())

    @app.get("/capabilities")
    def capabilities() -> dict[str, Any]:
        result = core.capabilities()
        result["version"] = __version__
        result["hotkey"] = settings.hotkey if settings.hotkey_enabled else None
        return _envelope_success(result, context_snapshot())

    @app.post("/inspect")
    def inspect(request: InspectRequest) -> dict[str, Any]:
        result = core.inspect(request.kind, request.params)
        return _envelope_success(result, context_snapshot())

    @app.post("/invoke")
    def invoke(request: InvokeRequest) -> dict[str, Any]:
        result = core.invoke(request.operation, request.params)
        return _envelope_success(result, context_snapshot())

    @app.post("/set-property")
    def set_property(request: SetPropertyRequest) -> dict[str, Any]:
        params: dict[str, Any] = {"confirm": request.confirm}
        if request.expected_document is not None:
            params["expected_document"] = request.expected_document
        result = core.set_property(request.property, request.value, params)
        return _envelope_success(result, context_snapshot())

    @app.post("/com")
    def com(request: ComRequest) -> dict[str, Any]:
        params: dict[str, Any] = {"confirm": request.confirm}
        if request.expected_document is not None:
            params["expected_document"] = request.expected_document
        result = core.com(
            request.mode,
            request.root,
            request.expression,
            request.value,
            request.dry_run,
            params,
        )
        return _envelope_success(result, context_snapshot())

    @app.post("/targets/capture")
    def capture_target(request: CaptureRequest) -> dict[str, Any]:
        result = core.capture(request.hwnd)
        return _envelope_success(result["target"], context_snapshot())

    @app.get("/targets")
    def list_targets() -> dict[str, Any]:
        result = core.targets()
        return _envelope_success(result, context_snapshot())

    @app.get("/targets/{target_id}")
    def target_info(target_id: str) -> dict[str, Any]:
        result = core.target_info(target_id)
        return _envelope_success(result["target"], context_snapshot())

    @app.delete("/targets/{target_id}")
    def remove_target(target_id: str) -> dict[str, Any]:
        result = core.remove_target(target_id)
        return _envelope_success(result, context_snapshot())

    return app
