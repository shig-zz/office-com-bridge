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
import win32gui

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
    target_not_found,
    target_not_set,
)
from .operations import com_ops, registry
from .target import ResolvedTarget, get_foreground_window, resolve_target
from .targets import TargetRegistry, TargetState, dedupe_key, utc_now

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


def _window_exists(hwnd: int) -> bool:
    try:
        return bool(win32gui.IsWindow(int(hwnd)))
    except Exception:
        return True


def _common_metadata(info: dict[str, Any]) -> dict[str, Any]:
    return {
        "full_name": str(info.get("full_name") or ""),
        "path": str(info.get("path") or ""),
    }


def _app_metadata(info: dict[str, Any]) -> dict[str, Any]:
    return {
        key: value
        for key, value in info.items()
        if key not in {"document", "full_name", "path"}
    }


class BridgeCore:
    def __init__(self, settings: Settings, worker: ComWorker | None = None) -> None:
        self.settings = settings
        self.worker = worker or ComWorker()
        self._registry = TargetRegistry()
        self._com_refs: dict[str, ResolvedTarget] = {}
        self._cached_summary: dict[str, Any] | None = None
        self._cached_targets: list[dict[str, Any]] = []
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
        self._registry.clear()
        self._com_refs.clear()
        self._refresh_caches()

    def cached_summary(self) -> dict[str, Any] | None:
        with self._cache_lock:
            return dict(self._cached_summary) if self._cached_summary else None

    def cached_targets(self) -> list[dict[str, Any]]:
        with self._cache_lock:
            return [dict(target) for target in self._cached_targets]

    def _refresh_caches(self) -> None:
        record = self._registry.selected()
        summary = None
        if record is not None and record.state is TargetState.ACTIVE:
            summary = record.summary()
        with self._cache_lock:
            self._cached_summary = summary
            self._cached_targets = self._registry.list_public()

    def capture(self, hwnd: int | None = None) -> dict[str, Any]:
        if hwnd is None:
            hwnd = get_foreground_window()
        return self.worker.call(
            self._capture_hwnd,
            int(hwnd),
            timeout=self.settings.request_timeout,
        )

    def capture_foreground(self) -> dict[str, Any]:
        return self.capture()

    def capture_hwnd(self, hwnd: int) -> dict[str, Any]:
        return self.capture(int(hwnd))

    def _capture_hwnd(self, hwnd: int) -> dict[str, Any]:
        resolved = resolve_target(hwnd)
        info = resolved.refresh_info()
        now = utc_now()
        key = dedupe_key(
            str(info.get("application") or resolved.app.value),
            int(info.get("pid") or resolved.pid),
            str(info.get("full_name") or ""),
            resolved.hwnd,
        )
        existing = self._registry.find_active_by_dedupe(key)
        if existing is not None:
            record = self._registry.update_metadata(
                existing.target_id,
                display_name=str(info.get("document") or existing.display_name),
                metadata=_common_metadata(info),
                app_metadata=_app_metadata(info),
                now=now,
            )
            action = "re-selected"
        else:
            record = self._registry.add(
                app_type=str(info.get("application") or resolved.app.value).lower(),
                display_name=str(info.get("document") or ""),
                dedupe_key=key,
                metadata=_common_metadata(info),
                app_metadata=_app_metadata(info),
                now=now,
            )
            action = "registered"
            for removed in self._registry.purge_expired():
                self._com_refs.pop(removed, None)
        if record is None:  # pragma: no cover - defensive
            raise internal_error("The target record disappeared during capture.")
        self._com_refs[record.target_id] = resolved
        self._registry.select(record.target_id, now=now)
        self._refresh_caches()
        logger.info(
            "target %s: id=%s application=%s document=%s hwnd=%s pid=%s (targets=%d)",
            action,
            record.target_id,
            record.app_type,
            record.display_name,
            resolved.hwnd,
            resolved.pid,
            len(self._registry),
        )
        return {"target": record.public()}

    def targets(self) -> dict[str, Any]:
        return self.worker.call(self._targets, timeout=self.settings.request_timeout)

    def _targets(self) -> dict[str, Any]:
        for record in self._registry.active_records():
            hwnd = record.app_metadata.get("hwnd")
            if hwnd and not _window_exists(int(hwnd)):
                self._expire_target(record.target_id, "window_gone")
        self._refresh_caches()
        return {
            "targets": self._registry.list_public(),
            "selected_target_id": self._registry.selected_target_id,
        }

    def target_info(self, target_id: str) -> dict[str, Any]:
        return self.worker.call(
            self._target_info,
            str(target_id),
            timeout=self.settings.request_timeout,
        )

    def _target_info(self, target_id: str) -> dict[str, Any]:
        record = self._registry.get(target_id)
        if record is None:
            raise target_not_found(target_id)
        if record.state is TargetState.ACTIVE:
            resolved = self._com_refs.get(target_id)
            if resolved is None or not resolved.is_alive():
                self._expire_target(target_id, "com_reference_lost")
            else:
                self._sync_record(target_id, resolved.refresh_info())
        self._refresh_caches()
        record = self._registry.get(target_id)
        if record is None:  # pragma: no cover - defensive
            raise target_not_found(target_id)
        return {"target": record.public()}

    def remove_target(self, target_id: str) -> dict[str, Any]:
        return self.worker.call(
            self._remove_target,
            str(target_id),
            timeout=self.settings.request_timeout,
        )

    def _remove_target(self, target_id: str) -> dict[str, Any]:
        record = self._registry.remove(target_id)
        if record is None:
            raise target_not_found(target_id)
        self._com_refs.pop(target_id, None)
        self._refresh_caches()
        logger.info("target removed: id=%s document=%s", target_id, record.display_name)
        return {"removed": target_id, "target": record.public()}

    def context(self) -> dict[str, Any]:
        return self.worker.call(self._context, timeout=self.settings.request_timeout)

    def _context(self) -> dict[str, Any]:
        target = self._require_target()
        info = target.refresh_info()
        selected_id = self._registry.selected_target_id
        if selected_id is not None:
            self._sync_record(selected_id, info)
        self._refresh_caches()
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
        selected_id = self._registry.selected_target_id
        if selected_id is None:
            raise target_not_set()
        record = self._registry.get(selected_id)
        if record is None or record.state is not TargetState.ACTIVE:
            self._registry.clear_selection()
            self._refresh_caches()
            raise target_lost(
                details={"target_id": selected_id, "reason": "expired"}
            )
        resolved = self._com_refs.get(selected_id)
        if resolved is None or not resolved.is_alive():
            self._expire_target(selected_id, "com_reference_lost")
            raise target_lost(
                details={"target_id": selected_id, "reason": "com_reference_lost"}
            )
        return resolved

    def _expire_target(self, target_id: str, reason: str) -> None:
        record = self._registry.expire(target_id, reason)
        self._com_refs.pop(target_id, None)
        self._refresh_caches()
        if record is not None:
            logger.warning(
                "target invalidated: id=%s reason=%s (was %s/%s)",
                target_id,
                reason,
                record.app_type,
                record.display_name,
            )

    def _sync_record(self, target_id: str, info: dict[str, Any]) -> None:
        record = self._registry.get(target_id)
        if record is None:
            return
        self._registry.update_metadata(
            target_id,
            display_name=str(info.get("document") or record.display_name),
            metadata=_common_metadata(info),
            app_metadata=_app_metadata(info),
        )

    def _execute(
        self,
        operation: str,
        function: Callable[[ResolvedTarget], T],
    ) -> T:
        target = self._require_target()
        target_id = self._registry.selected_target_id
        started = time.perf_counter()
        success = False
        try:
            result = function(target)
            success = True
            if target_id is not None:
                self._sync_record(target_id, target.info)
            self._refresh_caches()
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
                if target_id is not None:
                    self._expire_target(target_id, "com_reference_lost")
                raise target_lost(
                    details={"target_id": target_id, "reason": "com_reference_lost"}
                ) from exc
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
