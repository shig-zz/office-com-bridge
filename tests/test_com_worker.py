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


class FakeWorksheets:
    def __init__(self, names: tuple[str, ...] = ("Sheet1",)) -> None:
        self._names = names

    @property
    def Count(self) -> int:  # noqa: N802
        return len(self._names)

    def __call__(self, index: int) -> SimpleNamespace:
        return SimpleNamespace(Name=self._names[index - 1])


class FakeDocument:
    def __init__(self, name: str, full_name: str, alive: bool = True) -> None:
        self._name = name
        self._full_name = full_name
        self.alive = alive
        self.Worksheets = FakeWorksheets()  # noqa: N815
        self.ReadOnly = False  # noqa: N815

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

def test_inspect_routes_by_target_id_without_changing_selection(core, monkeypatch) -> None:
    a = capture(core, monkeypatch, 1000)["target"]
    b = capture(
        core, monkeypatch, 2000, document="Book2.xlsx", full_name=r"C:\docs\Book2.xlsx"
    )["target"]
    # Selection is now B. Read explicitly against A.
    result, summary = core.inspect("excel.workbook", {}, a["target_id"])
    assert summary is not None and summary["target_id"] == a["target_id"]
    assert result["target"]["target_id"] == a["target_id"]
    assert result["name"] == "Book1.xlsx"
    assert core._registry.selected_target_id == b["target_id"]  # noqa: SLF001

    # Default operation still uses the selected target B.
    result_b, summary_b = core.inspect("excel.workbook", {}, None)
    assert summary_b is not None and summary_b["target_id"] == b["target_id"]
    assert result_b["name"] == "Book2.xlsx"


def test_inspect_unknown_target_id_fails(core, monkeypatch) -> None:
    capture(core, monkeypatch, 1000)
    with pytest.raises(BridgeError) as excinfo:
        core.inspect("excel.workbook", {}, "t-nope00000000")
    assert excinfo.value.code == ErrorCode.TARGET_NOT_FOUND
    assert excinfo.value.details["target_id"] == "t-nope00000000"


def test_inspect_expired_target_id_fails_without_switch(core, monkeypatch) -> None:
    a = capture(core, monkeypatch, 1000)["target"]
    b = capture(
        core, monkeypatch, 2000, document="Book2.xlsx", full_name=r"C:\docs\Book2.xlsx"
    )["target"]
    core._resolved_registry[1000].document.alive = False  # type: ignore[attr-defined]

    with pytest.raises(BridgeError) as excinfo:
        core.inspect("excel.workbook", {}, a["target_id"])
    assert excinfo.value.code == ErrorCode.TARGET_LOST
    assert excinfo.value.details["target_id"] == a["target_id"]

    # No fallback: selection stays on B, A is marked expired.
    assert core._registry.selected_target_id == b["target_id"]  # noqa: SLF001
    assert core._registry.get(a["target_id"]).state.value == "expired"  # noqa: SLF001


def test_operation_errors_carry_the_used_target_id(core, monkeypatch) -> None:
    a = capture(core, monkeypatch, 1000)["target"]
    # expected_document checked against the explicitly routed target, not the
    # selected one: asking for a Word range on an Excel document fails cleanly
    # and names the target that was used.
    with pytest.raises(BridgeError) as excinfo:
        core.invoke(
            "word.selection.insert_text",
            {"text": "x", "confirm": True, "expected_document": "Nope.docx"},
            a["target_id"],
        )
    assert excinfo.value.code == ErrorCode.UNSUPPORTED_APPLICATION
    assert excinfo.value.details["target_id"] == a["target_id"]
    assert excinfo.value.details["application"] == "Excel"


def test_selection_is_validated_on_explicit_operations(core, monkeypatch) -> None:
    # Even with an explicit target_id, an unrelated dead *selected* target must
    # not break the request (routing is per request, not global state).
    a = capture(core, monkeypatch, 1000)["target"]
    b = capture(
        core, monkeypatch, 2000, document="Book2.xlsx", full_name=r"C:\docs\Book2.xlsx"
    )["target"]
    core._resolved_registry[1000].document.alive = False  # type: ignore[attr-defined]

    result, summary = core.inspect("excel.workbook", {}, b["target_id"])
    assert summary is not None and summary["target_id"] == b["target_id"]
    assert result["name"] == "Book2.xlsx"

def test_com_error_is_translated_and_worker_keeps_running(core, monkeypatch) -> None:
    import pywintypes

    from office_bridge.operations import registry as ops_registry

    state = {"fail": True}
    real_inspect = ops_registry.inspect

    def flaky_inspect(target, kind, params, settings):
        if state["fail"]:
            raise pywintypes.com_error(
                -2147352567,
                "Excel rejected the call",
                (0, None, None, None, None, -2146827284),
                None,
            )
        return real_inspect(target, kind, params, settings)

    monkeypatch.setattr(com_worker.registry, "inspect", flaky_inspect)
    capture(core, monkeypatch, 1000)

    with pytest.raises(BridgeError) as excinfo:
        core.inspect("excel.workbook", {})
    assert excinfo.value.code == ErrorCode.COM_ERROR

    # The target stays registered and usable after a COM error.
    state["fail"] = False
    result, summary = core.inspect("excel.workbook", {})
    assert summary is not None
    assert result["name"] == "Book1.xlsx"


def test_busy_hresult_maps_to_office_busy(core, monkeypatch) -> None:
    import pywintypes

    def busy_inspect(target, kind, params, settings):
        raise pywintypes.com_error(0x8001010A, "busy", None, None)

    monkeypatch.setattr(com_worker.registry, "inspect", busy_inspect)
    capture(core, monkeypatch, 1000)
    with pytest.raises(BridgeError) as excinfo:
        core.inspect("excel.workbook", {})
    assert excinfo.value.code == ErrorCode.OFFICE_BUSY


def test_serial_operations_across_two_targets_keep_records(core, monkeypatch) -> None:
    a = capture(core, monkeypatch, 1000)["target"]
    b = capture(
        core, monkeypatch, 2000, document="Book2.xlsx", full_name=r"C:\docs\Book2.xlsx"
    )["target"]
    for target_id, expect in (
        (a["target_id"], "Book1.xlsx"),
        (b["target_id"], "Book2.xlsx"),
        (a["target_id"], "Book1.xlsx"),
        (b["target_id"], "Book2.xlsx"),
    ):
        result, summary = core.inspect("excel.workbook", {}, target_id)
        assert result["name"] == expect
        assert summary["target_id"] == target_id
    # Records kept both documents; selection untouched.
    assert [r.display_name for r in core._registry.records()] == ["Book1.xlsx", "Book2.xlsx"]  # noqa: SLF001
    assert core._registry.selected_target_id == b["target_id"]  # noqa: SLF001


def test_com_batch_runs_get_steps_without_confirm(core, monkeypatch) -> None:
    capture(core, monkeypatch, 1000)
    calls: list[tuple[str, str]] = []

    def fake_execute(target, mode, root, segments, value, dry_run, params, settings):  # noqa: ANN001
        calls.append((mode, segments[0].name))
        return {"mode": mode, "value": segments[0].name}

    monkeypatch.setattr(com_worker.com_ops, "execute", fake_execute)
    result, _summary = core.com_batch(
        [
            {"mode": "get", "expression": "Name"},
            {"mode": "get", "expression": 'Sheets(2).Name'},
        ],
        {},
    )
    assert result["stopped"] is False
    assert result["applied"] == 0
    assert [item["ok"] for item in result["steps"]] == [True, True]
    assert calls == [("get", "Name"), ("get", "Sheets")]


def test_com_batch_requires_confirm_for_write_steps(core, monkeypatch) -> None:
    capture(core, monkeypatch, 1000)
    with pytest.raises(BridgeError) as excinfo:
        core.com_batch([{"mode": "call", "expression": 'Rows("2:4").Insert(-4107)'}], {})
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED


def test_com_batch_validates_steps_before_running(core, monkeypatch) -> None:
    capture(core, monkeypatch, 1000)
    with pytest.raises(BridgeError) as empty_exc:
        core.com_batch([], {})
    assert empty_exc.value.code == ErrorCode.INVALID_REQUEST

    with pytest.raises(BridgeError) as bad_exc:
        core.com_batch(
            [{"mode": "get", "expression": "Name"}, {"mode": "get", "expression": "Copy("}],
            {},
        )
    assert bad_exc.value.code == ErrorCode.INVALID_REQUEST
    assert bad_exc.value.details["step"] == 1


def test_com_batch_stops_at_first_failing_step(core, monkeypatch) -> None:
    capture(core, monkeypatch, 1000)

    def fake_execute(target, mode, root, segments, value, dry_run, params, settings):  # noqa: ANN001
        if segments[0].name == "Boom":
            raise BridgeError(ErrorCode.COM_ERROR, "boom", 502)
        return {"mode": mode, "value": segments[0].name}

    monkeypatch.setattr(com_worker.com_ops, "execute", fake_execute)
    result, _summary = core.com_batch(
        [
            {"mode": "get", "expression": "Name"},
            {"mode": "get", "expression": "Boom"},
            {"mode": "get", "expression": "Name"},
        ],
        {},
    )
    assert result["stopped"] is True
    assert [item["ok"] for item in result["steps"]] == [True, False]
    assert result["steps"][1]["error"]["code"] == "COM_ERROR"


def test_prune_targets_removes_expired_records(core, monkeypatch) -> None:
    first = capture(core, monkeypatch, 1000)["target"]
    second = capture(core, monkeypatch, 2000, pid=200, document="Book2.xlsx", full_name=r"C:\docs\Book2.xlsx")["target"]
    core._registry.expire(first["target_id"], "window_gone")  # noqa: SLF001
    result = core.prune_targets()
    assert result["removed"] == [first["target_id"]]
    assert result["count"] == 1
    remaining = core.targets()["targets"]
    assert [entry["target_id"] for entry in remaining] == [second["target_id"]]


def test_reconnect_target_reattaches_live_window(core, monkeypatch) -> None:
    target = capture(core, monkeypatch, 1000)["target"]
    resolved = core._com_refs[target["target_id"]]  # noqa: SLF001
    resolved.document.alive = False
    fresh = make_resolved(hwnd=1000)
    core._resolved_registry[1000] = fresh  # noqa: SLF001
    monkeypatch.setattr(com_worker, "_window_exists", lambda _hwnd: True)

    result = core.reconnect_target(target["target_id"])
    assert result["reconnected"] is True
    assert result["target"]["state"] == "active"
    assert core._com_refs[target["target_id"]] is fresh  # noqa: SLF001


def test_reconnect_target_reports_window_gone(core, monkeypatch) -> None:
    target = capture(core, monkeypatch, 1000)["target"]
    core._com_refs.pop(target["target_id"])  # noqa: SLF001
    monkeypatch.setattr(com_worker, "_window_exists", lambda _hwnd: False)
    with pytest.raises(BridgeError) as excinfo:
        core.reconnect_target(target["target_id"])
    assert excinfo.value.code == ErrorCode.TARGET_LOST
    assert excinfo.value.details["reason"] == "window_gone"


def test_reconnect_target_refuses_changed_document(core, monkeypatch) -> None:
    target = capture(core, monkeypatch, 1000)["target"]
    core._com_refs.pop(target["target_id"])  # noqa: SLF001
    monkeypatch.setattr(com_worker, "_window_exists", lambda _hwnd: True)
    core._resolved_registry[1000] = make_resolved(  # noqa: SLF001
        hwnd=1000, document="Other.xlsx", full_name=r"C:\docs\Other.xlsx"
    )
    with pytest.raises(BridgeError) as excinfo:
        core.reconnect_target(target["target_id"])
    assert excinfo.value.code == ErrorCode.TARGET_LOST
    assert excinfo.value.details["reason"] == "document_changed"


def test_target_lost_reports_reconnect_available(core, monkeypatch) -> None:
    target = capture(core, monkeypatch, 1000)["target"]
    core._com_refs[target["target_id"]].document.alive = False  # noqa: SLF001
    monkeypatch.setattr(com_worker, "_window_exists", lambda _hwnd: True)
    with pytest.raises(BridgeError) as excinfo:
        core.context()
    assert excinfo.value.code == ErrorCode.TARGET_LOST
    assert excinfo.value.details["reason"] == "com_reference_lost"
    assert excinfo.value.details["reconnect_available"] is True
