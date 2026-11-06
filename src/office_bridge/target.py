from __future__ import annotations

import ctypes
import logging
import os
import uuid
from ctypes import wintypes
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Callable

import pythoncom
import win32com.client
from win32com.client import dynamic as win32_dynamic  # type: ignore[attr-defined]
import win32gui
import win32process

from .errors import (
    BridgeError,
    instance_mismatch,
    no_document,
    not_office_window,
    resolve_failed,
)

logger = logging.getLogger(__name__)


class AppKind(str, Enum):
    EXCEL = "Excel"
    WORD = "Word"
    POWERPOINT = "PowerPoint"


OBJID_NATIVEOM = 0xFFFFFFF0
GA_ROOT = 2
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


def _guid(value: str) -> _GUID:
    parsed = uuid.UUID(value)
    result = _GUID()
    result.Data1 = parsed.time_low
    result.Data2 = parsed.time_mid
    result.Data3 = parsed.time_hi_version
    raw = parsed.bytes
    result.Data4 = (ctypes.c_ubyte * 8)(*raw[8:])
    return result


_IID_IDISPATCH = _guid("{00020400-0000-0000-C000-000000000046}")

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_oleacc = ctypes.OleDLL("oleacc")
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

_oleacc.AccessibleObjectFromWindow.restype = ctypes.c_long
_oleacc.AccessibleObjectFromWindow.argtypes = [
    wintypes.HWND,
    wintypes.DWORD,
    ctypes.POINTER(_GUID),
    ctypes.POINTER(ctypes.c_void_p),
]

_kernel32.OpenProcess.restype = wintypes.HANDLE
_kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
_kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
_kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.LPWSTR,
    ctypes.POINTER(wintypes.DWORD),
]
_kernel32.CloseHandle.argtypes = [wintypes.HANDLE]


OFFICE_EXECUTABLES = {
    "EXCEL.EXE": AppKind.EXCEL,
    "WINWORD.EXE": AppKind.WORD,
    "POWERPNT.EXE": AppKind.POWERPOINT,
}

APP_WINDOW_CLASSES = {
    "XLMAIN": AppKind.EXCEL,
    "OpusApp": AppKind.WORD,
    "PPTFrameClass": AppKind.POWERPOINT,
}


def get_foreground_window() -> int:
    return int(win32gui.GetForegroundWindow())


def root_window(hwnd: int) -> int:
    try:
        return int(win32gui.GetAncestor(int(hwnd), GA_ROOT))
    except Exception:
        return int(hwnd)


def window_text(hwnd: int) -> str:
    try:
        return win32gui.GetWindowText(int(hwnd))
    except Exception:
        return ""


def window_class(hwnd: int) -> str:
    try:
        return win32gui.GetClassName(int(hwnd)) or ""
    except Exception:
        return ""


def window_pid(hwnd: int) -> int:
    try:
        return int(win32process.GetWindowThreadProcessId(int(hwnd))[1])
    except Exception:
        return 0


def process_image_name(pid: int) -> str:
    if not pid:
        return ""
    handle = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
    if not handle:
        return ""
    try:
        buffer = ctypes.create_unicode_buffer(32768)
        size = wintypes.DWORD(len(buffer))
        if not _kernel32.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
            return ""
        return os.path.basename(buffer.value or "")
    finally:
        _kernel32.CloseHandle(handle)


def app_kind_for_window(hwnd: int) -> tuple[AppKind, int, str, str]:
    top = root_window(hwnd)
    cls = window_class(top)
    title = window_text(top)
    pid = window_pid(top)
    exe = process_image_name(pid).upper()

    kind = APP_WINDOW_CLASSES.get(cls)
    if kind is not None:
        return kind, top, title, exe

    if exe in OFFICE_EXECUTABLES:
        raise not_office_window(
            "The foreground window belongs to Office but is not a document window "
            "(for example, a dialog or the start screen).",
            {"hwnd": top, "class": cls, "process": exe, "title": title},
        )
    raise not_office_window(
        "The foreground window is not a Microsoft Word, Excel, or PowerPoint document window.",
        {"hwnd": top, "class": cls, "process": exe or None, "title": title},
    )


def find_descendant(root: int, class_name: str) -> int | None:
    found: list[int] = []

    def callback(hwnd: int, _lparam: Any) -> bool:
        if window_class(hwnd) == class_name:
            found.append(int(hwnd))
            return False
        return True

    try:
        win32gui.EnumChildWindows(int(root), callback, None)
    except Exception:
        pass
    return found[0] if found else None


def native_object(hwnd: int) -> Any:
    pointer = ctypes.c_void_p()
    try:
        hresult = _oleacc.AccessibleObjectFromWindow(
            int(hwnd),
            wintypes.DWORD(OBJID_NATIVEOM),
            ctypes.byref(_IID_IDISPATCH),
            ctypes.byref(pointer),
        )
    except OSError as exc:
        raise resolve_failed(
            "Could not obtain the Office native object for the window.",
            {"hwnd": int(hwnd), "error": str(exc)},
        ) from exc

    if hresult != 0 or not pointer.value:
        raise resolve_failed(
            "Could not obtain the Office native object for the window.",
            {"hwnd": int(hwnd), "hresult": f"0x{hresult & 0xFFFFFFFF:08X}"},
        )

    try:
        dispatch = pythoncom.ObjectFromAddress(pointer.value, pythoncom.IID_IDispatch)
        return win32_dynamic.Dispatch(dispatch)
    except Exception as exc:
        raise resolve_failed(
            "Could not wrap the Office native object.",
            {"hwnd": int(hwnd), "error": str(exc)},
        ) from exc


def _safe_str(getter: Callable[[], Any], default: str = "") -> str:
    try:
        value = getter()
    except Exception:
        return default
    if value is None:
        return default
    return str(value)


def _base_info(kind: AppKind, top: int, pid: int, title: str) -> dict[str, Any]:
    return {
        "application": kind.value,
        "document": "",
        "full_name": "",
        "path": "",
        "hwnd": top,
        "pid": pid,
        "window_title": title,
    }


def _excel_info(workbook: Any, window: Any, base: dict[str, Any]) -> dict[str, Any]:
    info = dict(base)
    info["document"] = _safe_str(lambda: workbook.Name)
    info["full_name"] = _safe_str(lambda: workbook.FullName)
    info["path"] = _safe_str(lambda: workbook.Path)
    info["sheet"] = _safe_str(lambda: window.ActiveSheet.Name)
    return info


def _word_info(document: Any, base: dict[str, Any]) -> dict[str, Any]:
    info = dict(base)
    info["document"] = _safe_str(lambda: document.Name)
    info["full_name"] = _safe_str(lambda: document.FullName)
    info["path"] = _safe_str(lambda: document.Path)
    return info


def _powerpoint_info(presentation: Any, window: Any | None, base: dict[str, Any]) -> dict[str, Any]:
    info = dict(base)
    info["document"] = _safe_str(lambda: presentation.Name)
    info["full_name"] = _safe_str(lambda: presentation.FullName)
    info["path"] = _safe_str(lambda: presentation.Path)
    info["slide_count"] = _safe_int(lambda: int(presentation.Slides.Count), 0)
    if window is not None:
        info["slide_index"] = _safe_int(lambda: int(window.View.Slide.SlideIndex), 0)
    else:
        info["slide_index"] = 0
    return info


def _safe_int(getter: Callable[[], int], default: int = 0) -> int:
    try:
        return int(getter())
    except Exception:
        return default


@dataclass(slots=True)
class ResolvedTarget:
    app: AppKind
    hwnd: int
    pid: int
    window_title: str
    application: Any
    document: Any
    window: Any
    info: dict[str, Any] = field(default_factory=dict)

    def summary(self) -> dict[str, Any]:
        return {
            "application": self.info.get("application", self.app.value),
            "document": self.info.get("document", ""),
            "full_name": self.info.get("full_name", ""),
            "hwnd": self.hwnd,
            "pid": self.pid,
        }

    def refresh_info(self) -> dict[str, Any]:
        if self.app is AppKind.EXCEL:
            self.info = _excel_info(self.document, self.window, self.info)
        elif self.app is AppKind.WORD:
            self.info = _word_info(self.document, self.info)
        elif self.app is AppKind.POWERPOINT:
            self.info = _powerpoint_info(self.document, self.window, self.info)
        return dict(self.info)

    def is_alive(self) -> bool:
        try:
            _ = self.application.Name
            _ = self.document.Name
            return True
        except Exception:
            return False


def resolve_target(hwnd: int) -> ResolvedTarget:
    kind, top, title, _exe = app_kind_for_window(hwnd)
    pid = window_pid(top)

    if kind is AppKind.EXCEL:
        return _resolve_excel(top, pid, title)
    if kind is AppKind.WORD:
        return _resolve_word(top, pid, title)
    if kind is AppKind.POWERPOINT:
        return _resolve_powerpoint(top, pid, title)
    raise not_office_window("Unsupported Office application.", {"hwnd": top, "title": title})


def _resolve_excel(top: int, pid: int, title: str) -> ResolvedTarget:
    grid = find_descendant(top, "EXCEL7")
    if not grid:
        raise no_document(
            "Excel is running but no workbook window is open in the selected window.",
            {"application": "Excel", "hwnd": top, "title": title},
        )

    native = native_object(grid)
    try:
        application = native.Application
        app_hwnd = int(application.Hwnd)
    except Exception as exc:
        raise resolve_failed(
            "Could not access the Excel application through the workbook window.",
            {"hwnd": top, "error": str(exc)},
        ) from exc

    if app_hwnd != top:
        raise instance_mismatch(
            "The Excel application resolved from the window does not match the foreground window.",
            {"window_hwnd": top, "application_hwnd": app_hwnd},
        )

    try:
        sheet = native.ActiveSheet
        workbook = sheet.Parent
        _ = workbook.Name
    except Exception as exc:
        raise no_document(
            "Excel does not have an active workbook in the selected window.",
            {"application": "Excel", "hwnd": top, "error": str(exc)},
        ) from exc

    base = _base_info(AppKind.EXCEL, top, pid, title)
    info = _excel_info(workbook, native, base)
    return ResolvedTarget(
        app=AppKind.EXCEL,
        hwnd=top,
        pid=pid,
        window_title=title,
        application=application,
        document=workbook,
        window=native,
        info=info,
    )


def _resolve_word(top: int, pid: int, title: str) -> ResolvedTarget:
    document_window = find_descendant(top, "_WwG")
    if not document_window:
        raise no_document(
            "Word is running but no document window is open in the selected window.",
            {"application": "Word", "hwnd": top, "title": title},
        )

    native = native_object(document_window)
    try:
        application = native.Application
        document = native.Document
        window_hwnd = int(native.Hwnd)
    except Exception as exc:
        raise resolve_failed(
            "Could not access the Word document through the window.",
            {"hwnd": top, "error": str(exc)},
        ) from exc

    if window_hwnd != top:
        raise instance_mismatch(
            "The Word window resolved from the document does not match the foreground window.",
            {"window_hwnd": top, "document_window_hwnd": window_hwnd},
        )

    base = _base_info(AppKind.WORD, top, pid, title)
    info = _word_info(document, base)
    return ResolvedTarget(
        app=AppKind.WORD,
        hwnd=top,
        pid=pid,
        window_title=title,
        application=application,
        document=document,
        window=native,
        info=info,
    )


_PPT_TITLE_SUFFIXES = (" - PowerPoint", " - Microsoft PowerPoint")


def _strip_powerpoint_suffix(title: str) -> str:
    base = title.strip()
    for suffix in _PPT_TITLE_SUFFIXES:
        if base.lower().endswith(suffix.lower()):
            return base[: -len(suffix)].strip()
    return base


def _presentation_score(presentation: Any, base_title: str) -> int:
    base = base_title.lower()
    if not base:
        return 0
    names: list[str] = []
    try:
        names.append(str(presentation.Name))
    except Exception:
        pass
    try:
        full_name = str(presentation.FullName)
        if full_name:
            names.append(os.path.basename(full_name))
    except Exception:
        pass

    best = 0
    for name in names:
        if not name:
            continue
        lowered = name.lower()
        if base == lowered:
            best = max(best, 100)
        elif base.startswith(lowered):
            best = max(best, 80)
        elif lowered in base:
            best = max(best, 50)
    return best


def _presentation_key(presentation: Any) -> tuple[str, str]:
    full_name = _safe_str(lambda: presentation.FullName)
    name = _safe_str(lambda: presentation.Name)
    return ((full_name or name).lower(), name.lower())


def _active_powerpoint() -> Any | None:
    try:
        return win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        return None


def _rot_presentations() -> list[Any]:
    results: list[Any] = []
    try:
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        context = pythoncom.CreateBindCtx(0)
    except Exception:
        return results

    while True:
        try:
            item = enum.Next()
        except pythoncom.com_error:
            break
        if not item:
            break
        moniker = item[0]  # type: ignore[index]
        try:
            display_name = moniker.GetDisplayName(context, None)  # type: ignore[call-arg]
        except Exception:
            continue
        lowered = display_name.lower()
        if not lowered.endswith((".ppt", ".pptx", ".pps", ".ppsx", ".pptm", ".ppsm")):
            continue
        try:
            unknown = rot.GetObject(moniker)
            dispatch = unknown.QueryInterface(pythoncom.IID_IDispatch)
            presentation = win32_dynamic.Dispatch(dispatch)
            _ = presentation.Name
            results.append(presentation)
        except Exception:
            continue
    return results


def _resolve_powerpoint_presentation(top: int, title: str) -> tuple[Any, Any]:
    base_title = _strip_powerpoint_suffix(title)
    candidates: dict[tuple[str, str], tuple[int, Any, Any]] = {}

    def add(presentation: Any, score: int) -> None:
        try:
            application = presentation.Application
            key = _presentation_key(presentation)
        except Exception:
            return
        current = candidates.get(key)
        if current is None or score > current[0]:
            candidates[key] = (score, presentation, application)

    active = _active_powerpoint()
    if active is not None:
        try:
            count = int(active.Presentations.Count)
        except Exception:
            count = 0
        for index in range(1, count + 1):
            try:
                presentation = active.Presentations(index)
                score = _presentation_score(presentation, base_title)
                if score > 0:
                    add(presentation, score)
            except Exception:
                continue

    if not candidates:
        for presentation in _rot_presentations():
            score = _presentation_score(presentation, base_title)
            if score > 0:
                add(presentation, score)

    if candidates:
        ranked = sorted(candidates.values(), key=lambda entry: entry[0], reverse=True)
        best_score = ranked[0][0]
        best = [entry for entry in ranked if entry[0] == best_score]
        if len(best) > 1:
            raise resolve_failed(
                "Multiple PowerPoint presentations match the foreground window. "
                "Open only the target presentation and press the hotkey again.",
                {
                    "hwnd": top,
                    "title": title,
                    "matches": [entry[1].Name for entry in best],
                },
            )
        return best[0][1], best[0][2]

    if active is not None:
        try:
            if int(active.Presentations.Count) == 1:
                presentation = active.Presentations(1)
                return presentation, presentation.Application
        except Exception:
            pass

    raise resolve_failed(
        "Could not determine which PowerPoint presentation belongs to the foreground window.",
        {"hwnd": top, "title": title},
    )


def _resolve_powerpoint(top: int, pid: int, title: str) -> ResolvedTarget:
    presentation, application = _resolve_powerpoint_presentation(top, title)
    window = None
    try:
        if int(presentation.Windows.Count) >= 1:
            window = presentation.Windows(1)
    except Exception:
        window = None

    base = _base_info(AppKind.POWERPOINT, top, pid, title)
    info = _powerpoint_info(presentation, window, base)
    return ResolvedTarget(
        app=AppKind.POWERPOINT,
        hwnd=top,
        pid=pid,
        window_title=title,
        application=application,
        document=presentation,
        window=window,
        info=info,
    )
