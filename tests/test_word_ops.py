from __future__ import annotations

from types import SimpleNamespace

import pytest

from office_bridge.errors import BridgeError, ErrorCode
from office_bridge.operations import word
from office_bridge.target import AppKind


class FakeSelection:
    def __init__(self, text: str = "Hello", start: int = 0, end: int = 5) -> None:
        self.Text = text
        self.Start = start
        self.End = end
        self.Type = 2
        self.typed = None
        self.collapsed = None

    def TypeText(self, text: str) -> None:
        self.typed = text

    def Collapse(self, direction: int) -> None:
        self.collapsed = direction


def make_target(selection: FakeSelection, document: str = "Report.docx") -> SimpleNamespace:
    return SimpleNamespace(
        app=AppKind.WORD,
        window=SimpleNamespace(Selection=selection),
        info={
            "application": "Word",
            "document": document,
            "full_name": rf"C:\docs\{document}",
        },
    )


def test_selection_reads_text_and_offsets() -> None:
    result = word.selection(make_target(FakeSelection()), max_text_chars=100)
    assert result["text"] == "Hello"
    assert result["characters"] == 5
    assert result["selection_type"] == "Normal"
    assert result["truncated"] is False


def test_selection_truncates_long_text() -> None:
    result = word.selection(
        make_target(FakeSelection("abcdefghij")), max_text_chars=4
    )
    assert result["text"] == "abcd"
    assert result["truncated"] is True


def test_insert_dry_run_does_not_change_selection() -> None:
    selection = FakeSelection()
    result = word.insert_text(
        make_target(selection),
        text="Replacement",
        replace=True,
        dry_run=True,
        max_text_chars=100,
        params={},
    )
    assert result["inserted"] is False
    assert selection.typed is None


def test_insert_requires_confirmation() -> None:
    selection = FakeSelection()
    with pytest.raises(BridgeError) as excinfo:
        word.insert_text(
            make_target(selection),
            text="Replacement",
            replace=True,
            dry_run=False,
            max_text_chars=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert selection.typed is None


def test_insert_replaces_when_confirmed() -> None:
    selection = FakeSelection()
    result = word.insert_text(
        make_target(selection),
        text="Replacement",
        replace=True,
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert result["inserted"] is True
    assert selection.typed == "Replacement"
    assert selection.collapsed is None


def test_insert_without_replace_collapses_to_start() -> None:
    selection = FakeSelection()
    word.insert_text(
        make_target(selection),
        text="Added",
        replace=False,
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert selection.collapsed == 0
    assert selection.typed == "Added"


def test_insert_rejects_oversized_text() -> None:
    with pytest.raises(BridgeError) as excinfo:
        word.insert_text(
            make_target(FakeSelection()),
            text="abcdef",
            replace=True,
            dry_run=True,
            max_text_chars=3,
            params={},
        )
    assert excinfo.value.code == ErrorCode.TEXT_TOO_LARGE
