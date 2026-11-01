from __future__ import annotations

from types import SimpleNamespace
from typing import Any

import pytest

from office_bridge.com_expr import parse_expression
from office_bridge.config import Settings
from office_bridge.errors import BridgeError, ErrorCode
from office_bridge.operations import com_ops
from office_bridge.target import AppKind


class Node:
    """COM-like fake: attribute container whose members may be callable."""

    def __init__(self, name: str = "node", members: dict[str, Any] | None = None, call_result: Any = None) -> None:
        object.__setattr__(self, "_name", name)
        object.__setattr__(self, "_members", dict(members or {}))
        object.__setattr__(self, "_call_result", call_result)
        self.calls: list[tuple] = []
        self.sets: dict[str, Any] = {}

    def __getattr__(self, item: str) -> Any:
        if item.startswith("_"):
            raise AttributeError(item)
        members = self.__dict__["_members"]
        if item in members:
            return members[item]
        raise AttributeError(f"{self._name} has no member {item!r}")

    def __setattr__(self, key: str, value: Any) -> None:
        if key.startswith("_") or key in ("calls", "sets"):
            object.__setattr__(self, key, value)
        else:
            self.sets[key] = value

    def __call__(self, *args: Any) -> Any:
        self.calls.append(args)
        return self._call_result


def make_target(
    *,
    application: Any = None,
    document: Any = None,
    window: Any = None,
    document_name: str = "Book1.xlsx",
) -> SimpleNamespace:
    return SimpleNamespace(
        app=AppKind.EXCEL,
        application=application if application is not None else Node("application"),
        document=document if document is not None else Node("document"),
        window=window,
        info={
            "application": "Excel",
            "document": document_name,
            "full_name": rf"C:\docs\{document_name}",
        },
    )


def run(
    mode: str,
    expression: str,
    target: SimpleNamespace,
    *,
    root: str = "window",
    value: Any = None,
    dry_run: bool = False,
    confirm: bool = False,
    settings: Settings | None = None,
) -> dict[str, Any]:
    params: dict[str, Any] = {}
    if confirm:
        params["confirm"] = True
    return com_ops.execute(
        target,
        mode,
        root,
        parse_expression(expression),
        value,
        dry_run,
        params,
        settings or Settings(),
    )


def test_get_reads_property_through_call_chain() -> None:
    range_node = Node("range", {"Value2": 42})
    sheet = Node("sheet", {"Range": Node("Range", call_result=range_node)})
    window = Node("window", {"ActiveSheet": sheet})
    target = make_target(window=window)
    result = run("get", 'ActiveSheet.Range("A1").Value2', target)
    assert result["value"] == 42
    assert result["truncated"] is False
    assert window.ActiveSheet.Range.calls == [("A1",)]


def test_get_truncates_long_strings() -> None:
    document = Node("document", {"Text": "x" * 50})
    target = make_target(document=document)
    result = run("get", "Text", target, root="document", settings=Settings(max_text_chars=10))
    assert result["value"] == "x" * 10
    assert result["truncated"] is True


def test_get_rejects_trailing_call() -> None:
    target = make_target(window=Node("window"))
    with pytest.raises(BridgeError) as excinfo:
        run("get", 'Range("A1").Copy()', target)
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_call_requires_trailing_call() -> None:
    target = make_target(window=Node("window", {"ActiveSheet": Node("sheet")}))
    with pytest.raises(BridgeError) as excinfo:
        run("call", "ActiveSheet.Name", target, confirm=True)
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_call_executes_when_confirmed() -> None:
    text = Node("TypeText")
    selection = Node("selection", {"TypeText": text})
    window = Node("window", {"Selection": selection})
    target = make_target(window=window)
    result = run("call", 'Selection.TypeText("hello")', target, confirm=True)
    assert result["executed"] is True
    assert text.calls == [("hello",)]


def test_call_returns_result_value() -> None:
    address = Node("Address", call_result="$A$1")
    range_node = Node("range", {"Address": address})
    sheet = Node("sheet", {"Range": Node("Range", call_result=range_node)})
    window = Node("window", {"ActiveSheet": sheet})
    target = make_target(window=window)
    result = run("call", 'ActiveSheet.Range("A1").Address(False, False)', target, confirm=True)
    assert result["executed"] is True
    assert result["result"] == "$A$1"
    assert result["arguments"] == [False, False]


def test_call_requires_confirmation() -> None:
    text = Node("TypeText")
    selection = Node("selection", {"TypeText": text})
    window = Node("window", {"Selection": selection})
    target = make_target(window=window)
    with pytest.raises(BridgeError) as excinfo:
        run("call", 'Selection.TypeText("hello")', target)
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert text.calls == []


def test_call_dry_run_does_not_execute() -> None:
    text = Node("TypeText")
    selection = Node("selection", {"TypeText": text})
    window = Node("window", {"Selection": selection})
    target = make_target(window=window)
    result = run("call", 'Selection.TypeText("hello")', target, dry_run=True)
    assert result["dry_run"] is True
    assert result["executed"] is False
    assert text.calls == []


def test_call_dry_run_rejects_non_method() -> None:
    range_node = Node("range", {"Value2": 42})
    sheet = Node("sheet", {"Range": Node("Range", call_result=range_node)})
    window = Node("window", {"ActiveSheet": sheet})
    target = make_target(window=window)
    with pytest.raises(BridgeError) as excinfo:
        run("call", 'ActiveSheet.Range("A1").Value2()', target, dry_run=True)
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_set_applies_when_confirmed() -> None:
    font = Node("font", {"Bold": None})
    selection = Node("selection", {"Font": font})
    window = Node("window", {"Selection": selection})
    target = make_target(window=window)
    result = run("set", "Selection.Font.Bold", target, value=True, confirm=True)
    assert result["applied"] is True
    assert font.sets == {"Bold": True}


def test_set_value2_2d_array_is_passed_as_tuple_of_tuples() -> None:
    range_node = Node("range", {"Value2": None})
    sheet = Node("sheet", {"Range": Node("Range", call_result=range_node)})
    window = Node("window", {"ActiveSheet": sheet})
    target = make_target(window=window)
    result = run("set", 'ActiveSheet.Range("A1:B2").Value2', target, value=[[1, 2], [3, 4]], confirm=True)
    assert result["applied"] is True
    assert result["dry_run"] is False
    assert result["value"] == [[1, 2], [3, 4]]
    assert range_node.sets == {"Value2": ((1, 2), (3, 4))}


def test_set_value2_1d_array_is_passed_unchanged() -> None:
    range_node = Node("range", {"Value2": None})
    sheet = Node("sheet", {"Range": Node("Range", call_result=range_node)})
    window = Node("window", {"ActiveSheet": sheet})
    target = make_target(window=window)
    result = run("set", 'ActiveSheet.Range("A1:B1").Value2', target, value=[5, 6], confirm=True)
    assert result["applied"] is True
    assert range_node.sets == {"Value2": [5, 6]}


def test_set_single_segment_uses_root_directly() -> None:
    document = Node("document", {"ReadOnly": False})
    target = make_target(document=document)
    result = run("set", "ReadOnly", target, root="document", value=False, confirm=True)
    assert result["applied"] is True
    assert document.sets == {"ReadOnly": False}


def test_set_requires_value() -> None:
    target = make_target(window=Node("window", {"Selection": Node("selection")}))
    with pytest.raises(BridgeError) as excinfo:
        run("set", "Selection.Font.Bold", target)
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_set_requires_confirmation() -> None:
    font = Node("font", {"Bold": None})
    selection = Node("selection", {"Font": font})
    window = Node("window", {"Selection": selection})
    target = make_target(window=window)
    with pytest.raises(BridgeError) as excinfo:
        run("set", "Selection.Font.Bold", target, value=True)
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert font.sets == {}


def test_set_dry_run_does_not_apply() -> None:
    font = Node("font", {"Bold": None})
    selection = Node("selection", {"Font": font})
    window = Node("window", {"Selection": selection})
    target = make_target(window=window)
    result = run("set", "Selection.Font.Bold", target, value=True, dry_run=True)
    assert result["applied"] is False
    assert font.sets == {}


def test_set_rejects_unknown_member() -> None:
    selection = Node("selection")
    window = Node("window", {"Selection": selection})
    target = make_target(window=window)
    with pytest.raises(BridgeError) as excinfo:
        run("set", "Selection.NotAProperty", target, value=1, dry_run=True)
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_blocklist_blocks_save_call() -> None:
    document = Node("document", {"Save": Node("Save")})
    target = make_target(document=document)
    with pytest.raises(BridgeError) as excinfo:
        run("call", "Save()", target, root="document", confirm=True)
    assert excinfo.value.code == ErrorCode.OPERATION_BLOCKED


def test_blocklist_blocks_save_variants_and_close_family() -> None:
    document = Node("document", {"SaveAs": Node("SaveAs"), "Close": Node("Close")})
    application = Node("application", {"Quit": Node("Quit"), "CloseAll": Node("CloseAll")})
    target = make_target(document=document, application=application)
    for expression, root in [
        ('SaveAs("other.docx")', "document"),
        ("Close()", "document"),
        ("Quit()", "application"),
        ("CloseAll()", "application"),
    ]:
        with pytest.raises(BridgeError) as excinfo:
            run("call", expression, target, root=root, confirm=True)
        assert excinfo.value.code == ErrorCode.OPERATION_BLOCKED


def test_blocklist_blocks_mid_path_save_call() -> None:
    document = Node("document", {"Save": Node("Save", call_result=Node("saved"))})
    target = make_target(document=document)
    with pytest.raises(BridgeError) as excinfo:
        run("get", "Save().Name", target, root="document")
    assert excinfo.value.code == ErrorCode.OPERATION_BLOCKED


def test_blocklist_allows_reading_saved_property() -> None:
    document = Node("document", {"Saved": True})
    target = make_target(document=document)
    result = run("get", "Saved", target, root="document")
    assert result["value"] is True


def test_blocklist_blocks_setting_saved_property() -> None:
    document = Node("document", {"Saved": False})
    target = make_target(document=document)
    with pytest.raises(BridgeError) as excinfo:
        run("set", "Saved", target, root="document", value=True, confirm=True)
    assert excinfo.value.code == ErrorCode.OPERATION_BLOCKED


def test_blocklist_custom_names() -> None:
    document = Node("document", {"Rename": Node("Rename")})
    target = make_target(document=document)
    settings = Settings(blocked_calls="rename, printout")
    for expression in ['Rename("other.docx")', "PrintOut()"]:
        with pytest.raises(BridgeError) as excinfo:
            run("call", expression, target, root="document", confirm=True, settings=settings)
        assert excinfo.value.code == ErrorCode.OPERATION_BLOCKED


def test_expected_document_mismatch() -> None:
    font = Node("font", {"Bold": None})
    selection = Node("selection", {"Font": font})
    window = Node("window", {"Selection": selection})
    target = make_target(window=window, document_name="Book1.xlsx")
    params = {"confirm": True, "expected_document": "Other.xlsx"}
    with pytest.raises(BridgeError) as excinfo:
        com_ops.execute(
            target, "set", "window", parse_expression("Selection.Font.Bold"),
            True, False, params, Settings(),
        )
    assert excinfo.value.code == ErrorCode.DOCUMENT_MISMATCH
    assert font.sets == {}


def test_expression_root_prefix_is_used_as_root() -> None:
    font = Node("font", {"Bold": 0})
    selection = Node("selection", {"Font": font})
    window = Node("window", {"Selection": selection})
    target = make_target(window=window)
    result = run("get", "window.Selection.Font.Bold", target)
    assert result["value"] == 0
    assert result["root"] == "window"


def test_expression_root_prefix_on_document() -> None:
    document = Node("document", {"Name": "doc1"})
    target = make_target(document=document)
    result = run("get", "document.Name", target)
    assert result["value"] == "doc1"
    assert result["root"] == "document"


def test_expression_root_prefix_conflicts_with_explicit_root() -> None:
    target = make_target(window=Node("window"))
    with pytest.raises(BridgeError) as excinfo:
        run("get", "window.Name", target, root="document")
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_expression_naming_only_a_root_is_rejected() -> None:
    target = make_target(window=Node("window"))
    with pytest.raises(BridgeError) as excinfo:
        run("get", "window", target)
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_unknown_root_is_rejected() -> None:
    target = make_target(window=Node("window"))
    with pytest.raises(BridgeError) as excinfo:
        run("get", "Name", target, root="workbook")
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_window_root_requires_window() -> None:
    target = make_target(window=None)
    with pytest.raises(BridgeError) as excinfo:
        run("get", "Name", target)
    assert excinfo.value.code == ErrorCode.NO_DOCUMENT


def test_blocked_summary_lists_prefixes_and_names() -> None:
    summary = com_ops.blocked_summary(Settings())
    assert summary["prefixes"] == ["save", "close"]
    assert "quit" in summary["names"]
    assert "exit" in summary["names"]
    assert com_ops.blocked_summary(Settings(blocked_calls="rename"))["names"]


def test_modes_and_roots_are_exported() -> None:
    assert com_ops.MODES == ("get", "call", "set")
    assert com_ops.ROOTS == ("application", "window", "document")
