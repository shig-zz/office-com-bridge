from __future__ import annotations

import json
from typing import Any

import pytest
from typer.testing import CliRunner

import office_cli.main as cli


class FakeClient:
    def __init__(self, *, health=None, context=None, invoke=None) -> None:
        self._health = health or {"ok": True, "context": None, "result": {"status": "running"}}
        self._context = context or {
            "ok": True,
            "context": {"application": "Excel", "document": "Book1.xlsx"},
            "result": {
                "application": "Excel",
                "document": "Book1.xlsx",
                "full_name": r"C:\docs\Book1.xlsx",
            },
        }
        self._invoke = invoke or (lambda operation, params: {"ok": True, "result": {"operation": operation}})
        self.invocations: list[tuple[str, dict[str, Any]]] = []
        self.set_calls: list[tuple[str, Any, bool, Any]] = []
        self.com_calls: list[tuple[str, str, str, Any, bool, bool, Any]] = []

    def health(self):
        return self._health

    def context(self):
        return self._context

    def capabilities(self):
        return {"ok": True, "context": None, "result": {"operations": []}}

    def inspect(self, kind, params=None):
        return {"ok": True, "context": None, "result": {"kind": kind, "params": params or {}}}

    def invoke(self, operation, params):
        self.invocations.append((operation, params))
        return self._invoke(operation, params)

    def set_property(self, name, value, confirm, expected_document):
        self.set_calls.append((name, value, confirm, expected_document))
        return {"ok": True, "result": {"property": name}}

    def com(self, mode, root, expression, value, dry_run, confirm, expected_document):
        self.com_calls.append((mode, root, expression, value, dry_run, confirm, expected_document))
        return {"ok": True, "result": {"mode": mode}}


def install_client(monkeypatch, client: FakeClient) -> None:
    monkeypatch.setattr(cli, "BridgeClient", lambda url, timeout: client)


def parse_stdout(result) -> dict[str, Any]:
    return json.loads(result.stdout)


def test_health_success(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["health"])
    assert result.exit_code == 0
    assert parse_stdout(result)["ok"] is True


def test_api_error_exits_one(monkeypatch) -> None:
    client = FakeClient(
        health={"ok": False, "context": None, "error": {"code": "TARGET_LOST", "message": "gone"}}
    )
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["health"])
    assert result.exit_code == 1
    assert parse_stdout(result)["error"]["code"] == "TARGET_LOST"


def test_unreachable_exits_two(monkeypatch) -> None:
    client = FakeClient()
    client.health = lambda: (_ for _ in ()).throw(cli.BridgeUnreachable("down"))
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["health"])
    assert result.exit_code == 2
    assert parse_stdout(result)["error"]["code"] == "BRIDGE_UNREACHABLE"


def test_range_read_passes_formulas(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    captured: dict[str, Any] = {}

    def inspect(kind, params=None):
        captured["kind"] = kind
        captured["params"] = params
        return {"ok": True, "context": None, "result": {}}

    client.inspect = inspect
    result = CliRunner().invoke(
        cli.app, ["excel", "range", "read", "A1:B2", "--formulas"]
    )
    assert result.exit_code == 0
    assert captured == {
        "kind": "excel.range",
        "params": {"address": "A1:B2", "include_formulas": True},
    }


def test_range_write_dry_run_does_not_confirm(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        [
            "excel",
            "range",
            "write",
            "A1:B2",
            "--json",
            "[[1,2],[3,4]]",
            "--dry-run",
        ],
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "excel.range.write"
    assert params["dry_run"] is True
    assert "confirm" not in params


def test_range_write_requires_confirmation_without_tty(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        ["excel", "range", "write", "A1:B2", "--json", "[[1,2],[3,4]]"],
    )
    assert result.exit_code == 2
    assert client.invocations == []


def test_range_write_with_yes_sends_confirm(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        [
            "excel",
            "range",
            "write",
            "A1:B2",
            "--json",
            "[[1,2],[3,4]]",
            "--yes",
            "--expected-document",
            "Book1.xlsx",
        ],
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "excel.range.write"
    assert params["confirm"] is True
    assert params["expected_document"] == "Book1.xlsx"


def test_add_hyperlink_dry_run_does_not_confirm(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        [
            "excel",
            "range",
            "add-hyperlink",
            "E35",
            "--url",
            "https://example.com",
            "--screen-tip",
            "公式ページ",
            "--dry-run",
        ],
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "excel.range.add_hyperlink"
    assert params["address"] == "E35"
    assert params["url"] == "https://example.com"
    assert params["screen_tip"] == "公式ページ"
    assert params["dry_run"] is True
    assert "confirm" not in params


def test_add_hyperlink_requires_url(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["excel", "range", "add-hyperlink", "E35"])
    assert result.exit_code == 2
    assert client.invocations == []


def test_add_hyperlink_requires_confirmation_without_tty(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        ["excel", "range", "add-hyperlink", "E35", "--url", "https://example.com"],
    )
    assert result.exit_code == 2
    assert client.invocations == []


def test_add_hyperlink_with_yes_sends_confirm(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        [
            "excel",
            "range",
            "add-hyperlink",
            "E35",
            "--url",
            "https://example.com",
            "--sub-address",
            "Sheet2!A1",
            "--yes",
            "--expected-document",
            "Book1.xlsx",
        ],
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "excel.range.add_hyperlink"
    assert params["confirm"] is True
    assert params["sub_address"] == "Sheet2!A1"
    assert params["expected_document"] == "Book1.xlsx"


def test_word_insert_requires_text(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app, ["word", "selection", "insert-text", "hello", "--dry-run"]
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "word.selection.insert_text"
    assert params["text"] == "hello"
    assert params["dry_run"] is True


def test_com_get_passes_payload(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["com", "get", 'ActiveSheet.Name', "--root", "window"])
    assert result.exit_code == 0
    assert client.com_calls == [("get", "window", "ActiveSheet.Name", None, False, False, None)]


def test_com_call_requires_confirmation_without_tty(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["com", "call", 'ActiveSheet.Range("A1").Copy()'])
    assert result.exit_code == 2
    assert client.com_calls == []


def test_com_call_with_yes_sends_confirm(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        [
            "com",
            "call",
            'ActiveSheet.Range("A1").Copy()',
            "--yes",
            "--expected-document",
            "Book1.xlsx",
        ],
    )
    assert result.exit_code == 0
    assert client.com_calls == [
        (
            "call",
            "window",
            'ActiveSheet.Range("A1").Copy()',
            None,
            False,
            True,
            "Book1.xlsx",
        )
    ]


def test_com_call_dry_run_does_not_confirm(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app, ["com", "call", 'ActiveSheet.Range("A1").Copy()', "--dry-run"]
    )
    assert result.exit_code == 0
    assert client.com_calls == [
        ("call", "window", 'ActiveSheet.Range("A1").Copy()', None, True, False, None)
    ]


def test_com_set_parses_json_value(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        ["com", "set", 'Range("A1:B2").Value2', "--value", "[[1,2],[3,4]]", "--yes"],
    )
    assert result.exit_code == 0
    mode, root, expression, value, dry_run, confirm, expected = client.com_calls[0]
    assert (mode, root, expression) == ("set", "window", 'Range("A1:B2").Value2')
    assert value == [[1, 2], [3, 4]]
    assert (dry_run, confirm) == (False, True)


def test_com_set_parses_plain_string_value(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        ["com", "set", "Selection.Font.Name", "--value", "MS Gothic", "--yes"],
    )
    assert result.exit_code == 0
    assert client.com_calls[0][3] == "MS Gothic"


def test_com_set_rejects_null_value(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app, ["com", "set", "Selection.Font.Name", "--value", "null", "--yes"]
    )
    assert result.exit_code == 2
    assert client.com_calls == []


def test_load_values_accepts_values_object() -> None:
    assert cli._load_values('{"values": [[1, 2]]}', None) == [[1, 2]]


def test_load_values_rejects_both_sources(tmp_path) -> None:
    path = tmp_path / "values.json"
    path.write_text("[[1]]", encoding="utf-8")
    with pytest.raises(cli.typer.BadParameter):
        cli._load_values("[[1]]", path)
