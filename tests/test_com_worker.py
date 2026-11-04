from __future__ import annotations

from types import SimpleNamespace

import pytest

import office_bridge.com_worker as com_worker
from office_bridge.config import Settings
from office_bridge.errors import BridgeError, ErrorCode
from office_bridge.target import AppKind, ResolvedTarget


class InlineWorker:
    """Runs submitted functions on the calling thread."""

    is_alive = True

    def call(self, function, *args, timeout=None, **kwargs):  # noqa: ANN001, ANN002, ANN003
        return function(*args, **kwargs)


class FakeDocument:
    def __init__(self, name: str, full_name: str, alive: bool = True) -> None:
        self._name = name
        self._full_name = full_name
        self.alive = alive

    @property
    def Name(self) -> str:  # noqa: N802
        if not self.alive:
            raise RuntimeError("document is closed")
        return self._name

    @property
    def FullName(self) -> str:  # noqa: N802
        if not self.alive:
            raise RuntimeError("document is closed")
        return self._full_name

    @property
    def Path(self) -> str:  # noqa: N802
        return r"C:\docs"


def make_resolved(
    document: str = "Book1.xlsx",
    full_name: str = r"C:\docs\Book1.xlsx",
    pid: int = 100,
    hwnd: int = 1000,
    alive: bool = True,
) -> ResolvedTarget:
    return ResolvedTarget(
        app=AppKind.EXCEL,
        hwnd=hwnd,
        pid=pid,
        window_title=document,
        application=SimpleNamespace(Name="Excel", Hwnd=hwnd),
        document=FakeDocument(document, full_name, alive=alive),
        window=SimpleNamespace(ActiveSheet=SimpleNamespace(Name="Sheet1")),
        info={
            "application": "Excel",
            "document": document,
            "full_name": full_name,
            "path": r"C:\docs",
            "hwnd": hwnd,
            "pid": pid,
            "window_title": document,
        },
    )


@pytest.fixture()
def core(monkeypatch) -> com_worker.BridgeCore:
    core = com_worker.BridgeCore(Settings(), worker=InlineWorker())
    resolved_registry: dict[int, ResolvedTarget] = {}

    def fake_resolve(hwnd: int) -> ResolvedTarget:
        return resolved_registry[int(hwnd)]

    monkeypatch.setattr(com_worker, "resolve_target", fake_resolve)
    core._resolved_registry = resolved_registry  # type: ignore[attr-defined]
    return core


def capture(core: com_worker.BridgeCore, monkeypatch, hwnd: int, **kwargs) -> dict:
    resolved = make_resolved(hwnd=hwnd, **kwargs)
    core._resolved_registry[hwnd] = resolved  # type: ignore[attr-defined]
    monkeypatch.setattr(com_worker, "_window_exists", lambda _hwnd: True)
    return core.capture_hwnd(hwnd)


def test_capture_registers_document_and_selects_it(core, monkeypatch) -> None:
    result = capture(core, monkeypatch, 1000)
    target = result["target"]
    assert target["app_type"] == "excel"
    assert target["display_name"] == "Book1.xlsx"
    assert target["state"] == "active"
    assert target["target_id"].startswith("t-")
    assert core._registry.selected_target_id == target["target_id"]  # noqa: SLF001

    summary = core.cached_summary()
    assert summary is not None
    assert summary["document"] == "Book1.xlsx"
    assert summary["target_id"] == target["target_id"]


def test_capture_adds_second_document_without_replacing_first(core, monkeypatch) -> None:
    first = capture(core, monkeypatch, 1000)["target"]
    second = capture(
        core, monkeypatch, 2000, document="Book2.xlsx", full_name=r"C:\docs\Book2.xlsx"
    )["target"]

    assert first["target_id"] != second["target_id"]
    assert len(core._registry) == 2  # noqa: SLF001
    assert core._registry.selected_target_id == second["target_id"]  # noqa: SLF001
    assert core._registry.get(first["target_id"]).state.value == "active"  # noqa: SLF001

    listing = core.targets()
    ids = [entry["target_id"] for entry in listing["targets"]]
    assert ids == [first["target_id"], second["target_id"]]
    assert listing["selected_target_id"] == second["target_id"]


def test_capture_same_document_twice_reuses_target_id(core, monkeypatch) -> None:
    first = capture(core, monkeypatch, 1000)["target"]
    second = capture(core, monkeypatch, 1000)["target"]
    assert first["target_id"] == second["target_id"]
    assert len(core._registry) == 1  # noqa: SLF001


def test_capture_unsaved_documents_are_distinct_by_hwnd(core, monkeypatch) -> None:
    first = capture(core, monkeypatch, 1000, document="Book1", full_name="")["target"]
    second = capture(core, monkeypatch, 2000, document="Book1", full_name="")["target"]
    assert first["target_id"] != second["target_id"]
    assert len(core._registry) == 2  # noqa: SLF001


def test_closed_selected_document_fails_without_fallback(core, monkeypatch) -> None:
    first = capture(core, monkeypatch, 1000)["target"]
    second_doc = make_resolved(
        document="Book2.xlsx", full_name=r"C:\docs\Book2.xlsx", hwnd=2000
    )
    core._resolved_registry[2000] = second_doc  # type: ignore[attr-defined]
    second = core.capture_hwnd(2000)["target"]

    second_doc.document.alive = False  # simulate the user closing Book2.xlsx

    with pytest.raises(BridgeError) as excinfo:
        core.context()
    assert excinfo.value.code == ErrorCode.TARGET_LOST
    assert excinfo.value.details["target_id"] == second["target_id"]
    assert excinfo.value.details["reason"] == "com_reference_lost"

    # No silent fallback: selection is cleared, Book1 stays registered but
    # operations still fail until a new capture.
    assert core._registry.selected_target_id is None  # noqa: SLF001
    assert core._registry.get(first["target_id"]).state.value == "active"  # noqa: SLF001
    assert core.cached_summary() is None
    with pytest.raises(BridgeError) as excinfo:
        core.context()
    assert excinfo.value.code == ErrorCode.TARGET_NOT_SET


def test_targets_list_marks_window_gone(core, monkeypatch) -> None:
    first = capture(core, monkeypatch, 1000)["target"]
    monkeypatch.setattr(com_worker, "_window_exists", lambda _hwnd: False)
    listing = core.targets()
    entry = next(t for t in listing["targets"] if t["target_id"] == first["target_id"])
    assert entry["state"] == "expired"
    assert entry["expired_reason"] == "window_gone"


def test_target_info_unknown_id_raises_not_found(core) -> None:
    with pytest.raises(BridgeError) as excinfo:
        core.target_info("t-nope00000000")
    assert excinfo.value.code == ErrorCode.TARGET_NOT_FOUND


def test_target_info_refreshes_and_detects_closed_document(core, monkeypatch) -> None:
    resolved = make_resolved(hwnd=1000)
    core._resolved_registry[1000] = resolved  # type: ignore[attr-defined]
    first = core.capture_hwnd(1000)["target"]

    info = core.target_info(first["target_id"])["target"]
    assert info["state"] == "active"

    resolved.document.alive = False
    info = core.target_info(first["target_id"])["target"]
    assert info["state"] == "expired"
    assert info["expired_reason"] == "com_reference_lost"


def test_remove_target_unregisters_and_clears_selection(core, monkeypatch) -> None:
    first = capture(core, monkeypatch, 1000)["target"]
    result = core.remove_target(first["target_id"])
    assert result["removed"] == first["target_id"]
    assert core._registry.get(first["target_id"]) is None  # noqa: SLF001
    assert core._registry.selected_target_id is None  # noqa: SLF001
    assert core.cached_targets() == []
    assert core.cached_summary() is None

    with pytest.raises(BridgeError) as excinfo:
        core.remove_target(first["target_id"])
    assert excinfo.value.code == ErrorCode.TARGET_NOT_FOUND


def test_shutdown_clears_registry_and_refs(core, monkeypatch) -> None:
    capture(core, monkeypatch, 1000)
    core._shutdown()  # noqa: SLF001
    assert len(core._registry) == 0  # noqa: SLF001
    assert core._com_refs == {}  # noqa: SLF001
    assert core.cached_summary() is None
