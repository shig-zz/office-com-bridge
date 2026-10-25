from __future__ import annotations

import ctypes
import logging
import queue
import threading
import time
from concurrent.futures import Future, TimeoutError as FutureTimeout
from ctypes import wintypes
from typing import Any, Callable, TypeVar

import pythoncom
import pywintypes

from .com_expr import parse_expression
from .config import Settings
from .errors import (
    BridgeError,
    com_error,
    internal_error,
    invalid_request,
    office_busy,
    office_no_response,
    target_lost,
    target_not_set,
)
from .operations import com_ops, registry
from .target import ResolvedTarget, get_foreground_window, resolve_target

logger = logging.getLogger(__name__)

T = TypeVar("T")
_STOP = object()

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_user32.PeekMessageW.restype = wintypes.BOOL
_user32.PeekMessageW.argtypes = [
    ctypes.POINTER(wintypes.MSG),
    wintypes.HWND,
    wintypes.UINT,
    wintypes.UINT,
    wintypes.UINT,
]
_user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
_user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]

_PM_REMOVE = 0x0001
_BUSY_HRESULTS = {0x8001010A, 0x800AC472, 0x80010001}


def _pump_messages() -> None:
    message = wintypes.MSG()
    while _user32.PeekMessageW(ctypes.byref(message), None, 0, 0, _PM_REMOVE):
        _user32.TranslateMessage(ctypes.byref(message))
        _user32.DispatchMessageW(ctypes.byref(message))


class ComWorker:
    """Dedicated STA thread that owns every Office COM object."""

    def __init__(self, name: str = "office-com-sta") -> None:
        self._name = name
        self._queue: queue.Queue[Any] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._ready = threading.Event()
        self._start_error: BaseException | None = None

    @property
    def is_alive(self) -> bool:
        return bool(self._thread and self._thread.is_alive())

    def start(self, timeout: float = 10.0) -> None:
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._run, name=self._name, daemon=True)
        self._thread.start()
        if not self._ready.wait(timeout):
            raise RuntimeError("The COM worker thread did not start in time.")
        if self._start_error is not None:
            raise RuntimeError(f"The COM worker thread failed to start: {self._start_error}")

    def _run(self) -> None:
        try:
            pythoncom.CoInitialize()
        except BaseException as exc:
            self._start_error = exc
            self._ready.set()
            return

        self._ready.set()
        logger.debug("COM worker started (thread=%s)", threading.current_thread().name)
        try:
            while True:
                try:
                    item = self._queue.get(timeout=0.05)
                except queue.Empty:
                    _pump_messages()
                    continue

                if item is _STOP:
                    break

                function, args, kwargs, future = item
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    result = function(*args, **kwargs)
                except BaseException as exc:
                    future.set_exception(exc)
                else:
                    future.set_result(result)
                _pump_messages()
        finally:
            _pump_messages()
            pythoncom.CoUninitialize()
            logger.debug("COM worker stopped")

    def submit(
        self,
        function: Callable[..., T],
        *args: Any,
        timeout: float | None = None,
        **kwargs: Any,
    ) -> T:
        if not self.is_alive:
            raise BridgeError(
                "BRIDGE_NOT_RUNNING",
                "The COM worker thread is not running.",
                503,
            )
        future: Future[T] = Future()
        self._queue.put((function, args, kwargs, future))
        try:
            return future.result(timeout=timeout)
        except FutureTimeout as exc:
            raise office_no_response(timeout or 0.0) from exc

    def call(
        self,
        function: Callable[..., T],
        *args: Any,
        timeout: float | None = None,
        **kwargs: Any,
    ) -> T:
        return self.submit(function, *args, timeout=timeout, **kwargs)

    def stop(self, timeout: float = 10.0) -> None:
        thread = self._thread
        if thread is None:
            return
        if thread.is_alive():
            self._queue.put(_STOP)
            thread.join(timeout)
            if thread.is_alive():
                logger.warning("COM worker did not stop within %.1fs", timeout)
        self._thread = None


class BridgeCore:
    def __init__(self, settings: Settings, worker: ComWorker | None = None) -> None:
        self.settings = settings
        self.worker = worker or ComWorker()
        self._target: ResolvedTarget | None = None
        self._cached_summary: dict[str, Any] | None = None
        self._cache_lock = threading.Lock()

    def start(self) -> None:
        self.worker.start()

    def stop(self) -> None:
        try:
            if self.worker.is_alive:
                self.worker.call(self._shutdown, timeout=5.0)
        except Exception:
            pass
        self.worker.stop()

    def _shutdown(self) -> None:
        self._target = None
        self._set_cache(None)

    def cached_summary(self) -> dict[str, Any] | None:
        with self._cache_lock:
            return dict(self._cached_summary) if self._cached_summary else None

    def _set_cache(self, summary: dict[str, Any] | None) -> None:
        with self._cache_lock:
            self._cached_summary = dict(summary) if summary else None

    def select_foreground(self) -> dict[str, Any]:
        return self.select_hwnd(get_foreground_window())

    def select_hwnd(self, hwnd: int) -> dict[str, Any]:
        return self.worker.call(
            self._select_hwnd,
            int(hwnd),
            timeout=self.settings.request_timeout,
        )

    def _select_hwnd(self, hwnd: int) -> dict[str, Any]:
        target = resolve_target(hwnd)
        info = target.refresh_info()

        previous = self._target
        self._target = target
        summary = target.summary()
        self._set_cache(summary)

        if previous is not None:
            try:
                previous_summary = previous.summary()
            except Exception:
                previous_summary = None
            logger.info(
                "target switched: %s/%s -> %s/%s (hwnd=%s pid=%s)",
                (previous_summary or {}).get("application"),
                (previous_summary or {}).get("document"),
                summary.get("application"),
                summary.get("document"),
                target.hwnd,
                target.pid,
            )
        else:
            logger.info(
                "target selected: application=%s document=%s hwnd=%s pid=%s",
                summary.get("application"),
                summary.get("document"),
                target.hwnd,
                target.pid,
            )
        return {"target": info}

    def context(self) -> dict[str, Any]:
        return self.worker.call(self._context, timeout=self.settings.request_timeout)

    def _context(self) -> dict[str, Any]:
        target = self._require_target()
        info = target.refresh_info()
        self._set_cache(target.summary())
        return info

    def inspect(self, kind: str, params: dict[str, Any]) -> dict[str, Any]:
        return self.worker.call(
            self._inspect,
            kind,
            params,
            timeout=self.settings.request_timeout,
        )

    def _inspect(self, kind: str, params: dict[str, Any]) -> dict[str, Any]:
        return self._execute(
            f"inspect:{kind}",
            lambda target: registry.inspect(target, kind, params, self.settings),
        )

    def invoke(self, operation: str, params: dict[str, Any]) -> dict[str, Any]:
        return self.worker.call(
            self._invoke,
            operation,
            params,
            timeout=self.settings.write_timeout,
        )

    def _invoke(self, operation: str, params: dict[str, Any]) -> dict[str, Any]:
        return self._execute(
            operation,
            lambda target: registry.invoke(target, operation, params, self.settings),
        )

    def set_property(
        self,
        name: str,
        value: Any,
        params: dict[str, Any],
    ) -> dict[str, Any]:
        return self.worker.call(
            self._set_property,
            name,
            value,
            params,
            timeout=self.settings.write_timeout,
        )

    def _set_property(self, name: str, value: Any, params: dict[str, Any]) -> dict[str, Any]:
        return self._execute(
            f"set-property:{name}",
            lambda target: registry.set_property(
                target, name, value, params, self.settings
            ),
        )

    def capabilities(self) -> dict[str, Any]:
        return registry.capabilities(self.settings)

    def com(
        self,
        mode: str,
        root: str,
        expression: str,
        value: Any,
        dry_run: bool,
        params: dict[str, Any],
    ) -> dict[str, Any]:
        if mode not in com_ops.MODES:
            raise invalid_request(
                f"Mode must be one of: {', '.join(com_ops.MODES)}.",
                {"mode": mode},
            )
        segments = parse_expression(expression)
        timeout = self.settings.request_timeout if mode == "get" else self.settings.write_timeout
        return self.worker.call(
            self._com,
            mode,
            root,
            segments,
            value,
            dry_run,
            params,
            timeout=timeout,
        )

    def _com(
        self,
        mode: str,
        root: str,
        segments: Any,
        value: Any,
        dry_run: bool,
        params: dict[str, Any],
    ) -> dict[str, Any]:
        return self._execute(
            f"com.{mode}",
            lambda target: com_ops.execute(
                target, mode, root, segments, value, dry_run, params, self.settings
            ),
        )

    def _require_target(self) -> ResolvedTarget:
        if self._target is None:
            raise target_not_set()
        if not self._target.is_alive():
            self._clear_target("the document or application is no longer available")
            raise target_lost()
        return self._target

    def _clear_target(self, reason: str) -> None:
        summary = self._target.summary() if self._target is not None else None
        self._target = None
        self._set_cache(None)
        logger.warning("target invalidated: %s (was %s)", reason, summary)

    def _execute(
        self,
        operation: str,
        function: Callable[[ResolvedTarget], T],
    ) -> T:
        target = self._require_target()
        started = time.perf_counter()
        success = False
        try:
            result = function(target)
            success = True
            self._set_cache(target.summary())
            return result
        except BridgeError as exc:
            logger.warning(
                "operation=%s target=%s/%s result=error code=%s",
                operation,
                target.info.get("application"),
                target.info.get("document"),
                exc.code,
            )
            raise
        except BaseException as exc:
            if not target.is_alive():
                self._clear_target("the document or application is no longer available")
                raise target_lost() from exc
            if isinstance(exc, pywintypes.com_error):
                hresult = int(getattr(exc, "hresult", 0) or 0) & 0xFFFFFFFF
                if hresult in _BUSY_HRESULTS:
                    logger.warning(
                        "operation=%s target=%s/%s result=busy hresult=0x%08X",
                        operation,
                        target.info.get("application"),
                        target.info.get("document"),
                        hresult,
                    )
                    raise office_busy(
                        "Office is busy or rejected the COM call. Wait a moment and try again.",
                        exc,
                    ) from exc
                logger.warning(
                    "operation=%s target=%s/%s result=com_error hresult=0x%08X",
                    operation,
                    target.info.get("application"),
                    target.info.get("document"),
                    hresult,
                )
                raise com_error("The Office COM call failed.", exc) from exc
            logger.exception("operation=%s failed unexpectedly", operation)
            raise internal_error(
                "An unexpected error occurred while executing the operation.",
                {"type": type(exc).__name__, "message": str(exc)},
            ) from exc
        finally:
            duration_ms = (time.perf_counter() - started) * 1000.0
            logger.info(
                "operation=%s target=%s/%s ok=%s duration_ms=%.1f",
                operation,
                target.info.get("application"),
                target.info.get("document"),
                success,
                duration_ms,
            )
