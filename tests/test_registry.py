from __future__ import annotations

from types import SimpleNamespace

import pytest

from office_bridge.config import Settings
from office_bridge.errors import BridgeError, ErrorCode
from office_bridge.operations import registry
from office_bridge.target import AppKind

_settings = Settings()


class FakeTarget:
    app: AppKind = AppKind.EXCEL

    def __init__(self, app: AppKind) -> None:
        self.app = app

    def is_alive(self) -> bool:
        return True


def make_target(app: AppKind) -> FakeTarget:
    target = FakeTarget(app)
    target.info = {
        "application": app.value.capitalize(),
        "document": f"{app.value.capitalize()} doc",
        "full_name": rf"C:\docs\{app.value.capitalize()} doc",
        "hwnd": 1234,
        "pid": 5678,
        "window_title": f"Doc - {app.value.capitalize()}",
    }
    return target


def test_inspections_listed_in_capabilities() -> None:
    cap = registry.capabilities(_settings)
    names = [item["name"] for item in cap["inspections"]]
    assert "selection" in names
    assert "document" in names
    assert "excel.range" in names
    assert "powerpoint.slide" in names


def test_operations_listed_in_capabilities() -> None:
    cap = registry.capabilities(_settings)
    names = [item["name"] for item in cap["operations"]]
    assert "excel.range.write" in names
    assert "word.selection.insert_text" in names
    assert "excel.range.add_hyperlink" in names
    assert "powerpoint.text.replace" in names
    assert "powerpoint.textbox.add" in names
    assert "powerpoint.slide.add" in names
    assert "powerpoint.notes.set" in names
    assert "powerpoint.slide.delete" in names
    assert "undo" in names
    inspections = [item["name"] for item in cap["inspections"]]
    assert "powerpoint.slides" in inspections
    assert "powerpoint.notes" in inspections
    assert cap["com"]["batch_endpoint"] == "POST /com/batch"


def test_write_results_report_unsaved_changes() -> None:
    from test_powerpoint_ops import FakeSlide
    from test_powerpoint_ops import make_target as make_deck_target

    deck = make_deck_target([FakeSlide(1, "Title", "old text")])
    result = registry.invoke(
        deck,
        "powerpoint.text.replace",
        {"slide": 1, "shape": 2, "find": "old", "replace": "new", "confirm": True},
        _settings,
    )
    assert "unsaved_changes" in result
    dry = registry.invoke(
        deck,
        "powerpoint.text.replace",
        {"slide": 1, "shape": 2, "find": "old", "replace": "new", "dry_run": True},
        _settings,
    )
    assert "unsaved_changes" not in dry


def test_powerpoint_notes_inspection_and_operations() -> None:
    from test_powerpoint_ops import FakeSlide
    from test_powerpoint_ops import make_target as make_deck_target

    deck = make_deck_target([FakeSlide(1, "Title", "body", notes="old notes")])
    read = registry.inspect(deck, "powerpoint.notes", {"slide": 1}, _settings)
    assert read["notes"] == "old notes"

    result = registry.invoke(
        deck,
        "powerpoint.notes.set",
        {"slide": 1, "text": "new notes", "confirm": True},
        _settings,
    )
    assert result["verified"] is True
    assert deck.document.Slides(1).notes_range.Text == "new notes"

    deleted = registry.invoke(
        deck,
        "powerpoint.slide.delete",
        {"slide": 1, "confirm": True},
        _settings,
    )
    assert deleted["verified"] is True
    assert deleted["slide_count_after"] == 0


def test_notes_and_slide_delete_require_powerpoint() -> None:
    with pytest.raises(BridgeError) as excinfo:
        registry.inspect(make_target(AppKind.EXCEL), "powerpoint.notes", {"slide": 1}, _settings)
    assert excinfo.value.code == ErrorCode.UNSUPPORTED_APPLICATION


def test_undo_dry_run_and_execution() -> None:
    target = make_target(AppKind.EXCEL)
    dry = registry.invoke(target, "undo", {"dry_run": True}, _settings)
    assert dry["operation"] == "undo"
    assert dry["executed"] is False
    assert dry["undoable"] is None

    with pytest.raises(BridgeError) as excinfo:
        registry.invoke(target, "undo", {}, _settings)
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED

    calls: list[bool] = []
    target.application = SimpleNamespace(
        Undoable=True, Undo=lambda: calls.append(True)
    )
    result = registry.invoke(target, "undo", {"confirm": True}, _settings)
    assert result["executed"] is True
    assert calls == [True]


def test_slides_inspection_dispatches_to_powerpoint() -> None:
    from test_powerpoint_ops import FakeSlide
    from test_powerpoint_ops import make_target as make_deck_target

    deck = make_deck_target([FakeSlide(1, "One", "body", notes="n1")])
    result = registry.inspect(deck, "powerpoint.slides", {"start": 1, "count": 5}, _settings)
    assert result["unit"] == "pt"
    assert result["slides"][0]["slide_index"] == 1
    assert result["slides"][0]["notes"] == "n1"


def test_slides_inspection_requires_powerpoint() -> None:
    with pytest.raises(BridgeError) as excinfo:
        registry.inspect(make_target(AppKind.EXCEL), "powerpoint.slides", {}, _settings)
    assert excinfo.value.code == ErrorCode.UNSUPPORTED_APPLICATION


def test_powerpoint_text_replace_dispatches_to_operation() -> None:
    from test_powerpoint_ops import FakeSlide
    from test_powerpoint_ops import make_target as make_deck_target

    deck = make_deck_target([FakeSlide(1, "Title", "alpha 1963")])
    result = registry.invoke(
        deck,
        "powerpoint.text.replace",
        {"slide": 1, "shape": 2, "find": "1963", "replace": "1968", "confirm": True},
        _settings,
    )
    assert result["verified"] is True
    assert deck.document.Slides(1).Shapes(2).TextFrame.TextRange.Text == "alpha 1968"


def test_powerpoint_write_requires_powerpoint_app() -> None:
    for operation in ("powerpoint.text.replace", "powerpoint.textbox.add", "powerpoint.slide.add"):
        with pytest.raises(BridgeError) as excinfo:
            registry.invoke(make_target(AppKind.WORD), operation, {}, _settings)
        assert excinfo.value.code == ErrorCode.UNSUPPORTED_APPLICATION


def test_add_hyperlink_dispatches_to_excel_operation() -> None:
    from types import SimpleNamespace

    from test_excel_ops import FakeHyperlinks, FakeRange, FakeWindow

    link_range = FakeRange(1, 1, "https://example.com", address="E35")
    window = FakeWindow(FakeRange(1, 1, None), {"E35": link_range})
    target = make_target(AppKind.EXCEL)
    target.window = window

    result = registry.invoke(
        target,
        "excel.range.add_hyperlink",
        {"address": "E35", "url": "https://example.com", "confirm": True},
        _settings,
    )
    assert result["applied"] is True
    assert link_range.Hyperlinks.Count == 1
    assert link_range.Hyperlinks(1).Address == "https://example.com"


def test_add_hyperlink_requires_excel_app() -> None:
    with pytest.raises(BridgeError) as excinfo:
        registry.invoke(
            make_target(AppKind.WORD),
            "excel.range.add_hyperlink",
            {"address": "A1", "url": "https://example.com"},
            _settings,
        )
    assert excinfo.value.code == ErrorCode.UNSUPPORTED_APPLICATION


def test_inspect_rejects_unknown_kind() -> None:
    with pytest.raises(BridgeError) as excinfo:
        registry.inspect(FakeTarget(AppKind.EXCEL), "unknown.kind", {}, _settings)
    assert excinfo.value.code == ErrorCode.UNSUPPORTED_OPERATION


def test_inspections_validate_matching_app() -> None:
    with pytest.raises(BridgeError) as app_exc:
        registry.inspect(
            make_target(AppKind.WORD),  # Wrong app for excel.range
            "excel.range",
            {"address": "A1"},
            _settings,
        )
    assert app_exc.value.code == ErrorCode.UNSUPPORTED_APPLICATION


def test_invoke_rejects_unknown_operation() -> None:
    target = make_target(AppKind.EXCEL)

    with pytest.raises(BridgeError) as excinfo:
        registry.invoke(target, "excel.random_fake_op", {"address": "A1"}, _settings)


def test_invoke_requires_app_match() -> None:
    word_target = make_target(AppKind.WORD)

    with pytest.raises(BridgeError) as excinfo:
        registry.invoke(word_target, "excel.range.write", {"address": "A1"}, _settings)


def test_set_property_rejects_unknown_name() -> None:
    target = make_target(AppKind.EXCEL)

    with pytest.raises(BridgeError) as excinfo:
        registry.set_property(target, "unknown.property", None, {}, _settings)


def test_security_flags_in_capabilities() -> None:
    cap = registry.capabilities(_settings)
    assert cap["security"]["localhost_only"] is True
    assert cap["security"]["writes_require_confirmation"] is True
    assert cap["security"]["allowed_operations_only"] is False
    assert cap["security"]["general_com_access"] is True
    blocked = cap["security"]["blocked_calls"]
    assert blocked["prefixes"] == ["save", "close"]
    assert "quit" in blocked["names"]
    assert "exit" in blocked["names"]


def test_com_section_in_capabilities() -> None:
    cap = registry.capabilities(_settings)
    com = cap["com"]
    assert com["endpoint"] == "POST /com"
    assert [mode["name"] for mode in com["modes"]] == ["get", "call", "set"]
    assert com["roots"] == ["application", "window", "document"]
    assert "window" in com["examples"]["get"]
