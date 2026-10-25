from __future__ import annotations

import ctypes
import logging
import threading
from ctypes import wintypes
from typing import Callable

logger = logging.getLogger(__name__)

WM_HOTKEY = 0x0312
WM_QUIT = 0x0012

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
MOD_NOREPEAT = 0x4000

_HOTKEY_ID = 0xB12D

_NAMED_KEYS = {
    "space": 0x20,
    "tab": 0x09,
    "enter": 0x0D,
    "return": 0x0D,
    "esc": 0x1B,
    "escape": 0x1B,
    "backspace": 0x08,
    "delete": 0x2E,
    "insert": 0x2D,
    "home": 0x24,
    "end": 0x23,
    "pageup": 0x21,
    "pagedown": 0x22,
    "left": 0x25,
    "up": 0x26,
    "right": 0x27,
    "down": 0x28,
    "pause": 0x13,
    "printscreen": 0x2C,
    "comma": 0xBC,
    "period": 0xBE,
    "slash": 0xBF,
    "semicolon": 0xBA,
    "quote": 0xDE,
    "backtick": 0xC0,
    "minus": 0xBD,
    "equals": 0xBB,
    "bracketleft": 0xDB,
    "bracketright": 0xDD,
    "backslash": 0xDC,
}

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

_user32.RegisterHotKey.restype = wintypes.BOOL
_user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
_user32.UnregisterHotKey.restype = wintypes.BOOL
_user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
_user32.GetMessageW.restype = ctypes.c_int
_user32.GetMessageW.argtypes = [
    ctypes.POINTER(wintypes.MSG),
    wintypes.HWND,
    wintypes.UINT,
    wintypes.UINT,
]
_user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
_user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
_user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
_kernel32.GetCurrentThreadId.restype = wintypes.DWORD


class HotkeyError(Exception):
    pass


def parse_hotkey(specification: str) -> tuple[int, int]:
    if not isinstance(specification, str) or not specification.strip():
        raise HotkeyError("Hotkey specification is empty.")
    parts = [part.strip().lower() for part in specification.split("+") if part.strip()]
    if len(parts) < 1:
        raise HotkeyError("Hotkey specification is empty.")

    modifiers = 0
    key_name = ""
    for part in parts:
        if part in {"ctrl", "control"}:
            modifiers |= MOD_CONTROL
        elif part == "alt":
            modifiers |= MOD_ALT
        elif part == "shift":
            modifiers |= MOD_SHIFT
        elif part in {"win", "windows", "super"}:
            modifiers |= MOD_WIN
        elif part:
            if key_name:
                raise HotkeyError(f"Hotkey {specification!r} contains more than one key.")
            key_name = part

    if not key_name:
        raise HotkeyError(f"Hotkey {specification!r} does not include a key.")

    key_code = _virtual_key(key_name)
    is_function_key = 0x70 <= key_code <= 0x87
    if modifiers == 0 and not is_function_key:
        raise HotkeyError(
            "A global hotkey must include at least one modifier (Ctrl, Alt, Shift, Win) "
            "unless it is a function key."
        )
    return modifiers, key_code


def _virtual_key(name: str) -> int:
    if name in _NAMED_KEYS:
        return _NAMED_KEYS[name]
    if len(name) == 1:
        character = name.upper()
        if "A" <= character <= "Z" or "0" <= character <= "9":
            return ord(character)
    if name.startswith("f") and name[1:].isdigit():
        number = int(name[1:])
        if 1 <= number <= 24:
            return 0x70 + number - 1
    raise HotkeyError(f"Unsupported key name: {name!r}")


class HotkeyListener(threading.Thread):
    def __init__(self, specification: str, callback: Callable[[], None]) -> None:
        super().__init__(name="office-hotkey", daemon=True)
        self.specification = specification
        self._modifiers, self._virtual_key = parse_hotkey(specification)
        self._callback = callback
        self._ready = threading.Event()
        self._stopping = threading.Event()
        self._error: BaseException | None = None
        self._thread_id: int | None = None

    def start(self, timeout: float = 5.0) -> None:
        super().start()
        if not self._ready.wait(timeout):
            raise HotkeyError("The hotkey listener did not start in time.")
        if self._error is not None:
            raise HotkeyError(str(self._error))

    def run(self) -> None:
        self._thread_id = int(_kernel32.GetCurrentThreadId())
        registered = False
        try:
            flags = self._modifiers | MOD_NOREPEAT
            if not _user32.RegisterHotKey(None, _HOTKEY_ID, flags, self._virtual_key):
                error_code = ctypes.get_last_error()
                raise HotkeyError(
                    f"RegisterHotKey failed for {self.specification!r} "
                    f"(Windows error {error_code}). Another application may already use it."
                )
            registered = True
            self._ready.set()
            logger.info("hotkey registered: %s", self.specification)

            message = wintypes.MSG()
            while not self._stopping.is_set():
                result = _user32.GetMessageW(ctypes.byref(message), None, 0, 0)
                if result == 0 or result == -1:
                    break
                if message.message == WM_HOTKEY and int(message.wParam) == _HOTKEY_ID:
                    try:
                        self._callback()
                    except Exception:
                        logger.exception("hotkey callback failed")
                    continue
                _user32.TranslateMessage(ctypes.byref(message))
                _user32.DispatchMessageW(ctypes.byref(message))
        except BaseException as exc:
            self._error = exc
            self._ready.set()
            logger.error("hotkey listener failed: %s", exc)
        finally:
            if registered:
                _user32.UnregisterHotKey(None, _HOTKEY_ID)
            logger.info("hotkey listener stopped")

    def stop(self, timeout: float = 5.0) -> None:
        self._stopping.set()
        if self.is_alive() and self._thread_id is not None:
            _user32.PostThreadMessageW(self._thread_id, WM_QUIT, 0, 0)
            self.join(timeout)
