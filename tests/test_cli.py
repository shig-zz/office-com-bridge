from __future__ import annotations

import json
from types import SimpleNamespace
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
        self._targets = {
            "ok": True,
            "result": {
                "targets": [
                    {"target_id": "t-aaaa1111aaaa", "display_name": "Book1.xlsx", "state": "active"},
                    {"target_id": "t-bbbb2222bbbb", "display_name": "Doc1.docx", "state": "expired"},
                ],
                "selected_target_id": "t-aaaa1111aaaa",
            },
        }
        self.invocations: list[tuple[str, dict[str, Any]]] = []
        self.set_calls: list[tuple[str, Any, bool, Any]] = []
        self.com_calls: list[tuple[str, str, str, Any, bool, bool, Any]] = []
        self.batch_calls: list[tuple[list[dict[str, Any]], bool, Any]] = []
        self.prune_calls: list[bool] = []
        self.reconnect_calls: list[str] = []
        self.capture_calls: list[Any] = []
        self.info_calls: list[str] = []
        self.remove_calls: list[str] = []
        self.target_id: str | None = None

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

    def com_batch(self, steps, confirm, expected_document):
        self.batch_calls.append((steps, confirm, expected_document))
        return {"ok": True, "result": {"steps": [{"step": i, "ok": True} for i in range(len(steps))], "stopped": False}}

    def prune_targets(self):
        self.prune_calls.append(True)
        return {"ok": True, "result": {"removed": ["t-expired00000"], "count": 1}}

    def reconnect_target(self, target_id):
        self.reconnect_calls.append(target_id)
        return {"ok": True, "result": {"target": {"target_id": target_id, "state": "active"}, "reconnected": True}}

    def capture(self, hwnd=None):
        self.capture_calls.append(hwnd)
        return {"ok": True, "result": {"target_id": "t-aaaa1111aaaa", "state": "active"}}

    def targets(self):
        return self._targets

    def target_info(self, target_id):
        self.info_calls.append(target_id)
        return {
            "ok": True,
            "result": {
                "target_id": target_id,
                "state": "active",
                "app_type": "excel",
                "display_name": "Book1.xlsx",
                "metadata": {"full_name": r"C:\docs\Book1.xlsx"},
            },
        }

    def remove_target(self, target_id):
        self.remove_calls.append(target_id)
        return {"ok": True, "result": {"removed": target_id}}


def install_client(monkeypatch, client: FakeClient) -> None:
    def factory(url, timeout, target_id=None):
        client.target_id = target_id
        return client

    monkeypatch.setattr(cli, "BridgeClient", factory)


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


def test_com_set_value_from_utf8_file(tmp_path, monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    path = tmp_path / "note.txt"
    raw = "日本語メモ\n段落区切り"
    path.write_text(raw, encoding="utf-8")
    result = CliRunner().invoke(
        cli.app, ["com", "set", "TextRange.Text", "--value", f"@{path}", "--yes"]
    )
    assert result.exit_code == 0
    assert client.com_calls[0][3] == raw


def test_com_set_value_from_json_file(tmp_path, monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    path = tmp_path / "values.json"
    path.write_text("[[1, 2], [3, 4]]", encoding="utf-8")
    result = CliRunner().invoke(
        cli.app, ["com", "set", 'Range("A1:B2").Value2', "--value", f"@{path}", "--yes"]
    )
    assert result.exit_code == 0
    assert client.com_calls[0][3] == [[1, 2], [3, 4]]


def test_com_set_value_missing_file_fails(tmp_path, monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        ["com", "set", "TextRange.Text", "--value", f"@{tmp_path / 'nope.txt'}", "--yes"],
    )
    assert result.exit_code != 0
    assert client.com_calls == []


def test_powerpoint_slides_read_passes_window(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    captured: dict[str, Any] = {}

    def inspect(kind, params=None):
        captured["kind"] = kind
        captured["params"] = params
        return {"ok": True, "context": None, "result": {}}

    client.inspect = inspect
    result = CliRunner().invoke(
        cli.app, ["powerpoint", "slides", "read", "--start", "21", "--count", "10"]
    )
    assert result.exit_code == 0
    assert captured == {"kind": "powerpoint.slides", "params": {"start": 21, "count": 10}}


def test_powerpoint_text_replace_with_yes_sends_confirm(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        [
            "powerpoint", "text", "replace", "3",
            "--shape", "7",
            "--find", "1979",
            "--replace", "1982",
            "--yes",
        ],
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "powerpoint.text.replace"
    assert params == {
        "slide": 3,
        "shape": 7,
        "find": "1979",
        "replace": "1982",
        "dry_run": False,
        "confirm": True,
    }


def test_powerpoint_text_replace_keeps_shape_name_as_string(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        [
            "powerpoint", "text", "replace", "3",
            "--shape", "TextBox 7",
            "--find", "a",
            "--replace", "b",
            "--dry-run",
        ],
    )
    assert result.exit_code == 0
    _, params = client.invocations[0]
    assert params["shape"] == "TextBox 7"
    assert params["dry_run"] is True
    assert "confirm" not in params


def test_powerpoint_text_replace_value_from_file(tmp_path, monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    path = tmp_path / "replacement.txt"
    path.write_text("置換後テキスト", encoding="utf-8")
    result = CliRunner().invoke(
        cli.app,
        [
            "powerpoint", "text", "replace", "3",
            "--shape", "2",
            "--find", "old",
            "--replace", f"@{path}",
            "--yes",
        ],
    )
    assert result.exit_code == 0
    assert client.invocations[0][1]["replace"] == "置換後テキスト"


def test_powerpoint_textbox_add_with_yes(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        [
            "powerpoint", "textbox", "add", "3",
            "--left", "20",
            "--top", "460",
            "--width", "920",
            "--height", "60",
            "--text", "補足",
            "--font-size", "12",
            "--yes",
        ],
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "powerpoint.textbox.add"
    assert params["left"] == 20.0
    assert params["top"] == 460.0
    assert params["text"] == "補足"
    assert params["font_size"] == 12.0
    assert params["confirm"] is True


def test_powerpoint_slide_add_with_yes(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app, ["powerpoint", "slide", "add", "--index", "70", "--yes"]
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "powerpoint.slide.add"
    assert params["index"] == 70
    assert "layout" not in params
    assert params["confirm"] is True


def test_powerpoint_slide_add_appends_without_index(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["powerpoint", "slide", "add", "--dry-run"])
    assert result.exit_code == 0
    _, params = client.invocations[0]
    assert "index" not in params
    assert params["dry_run"] is True


def test_powerpoint_notes_read_passes_slide(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    captured: dict[str, Any] = {}

    def inspect(kind, params=None):
        captured["kind"] = kind
        captured["params"] = params
        return {"ok": True, "context": None, "result": {}}

    client.inspect = inspect
    result = CliRunner().invoke(cli.app, ["powerpoint", "notes", "read", "70"])
    assert result.exit_code == 0
    assert captured == {"kind": "powerpoint.notes", "params": {"slide": 70}}


def test_powerpoint_notes_set_with_yes(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app,
        ["powerpoint", "notes", "set", "70", "--text", "校訂ノート", "--yes"],
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "powerpoint.notes.set"
    assert params == {"slide": 70, "text": "校訂ノート", "dry_run": False, "confirm": True}


def test_powerpoint_notes_set_text_from_file(tmp_path, monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    path = tmp_path / "notes.txt"
    path.write_text("一行目\n二行目", encoding="utf-8")
    result = CliRunner().invoke(
        cli.app,
        ["powerpoint", "notes", "set", "70", "--text", f"@{path}", "--yes"],
    )
    assert result.exit_code == 0
    assert client.invocations[0][1]["text"] == "一行目\n二行目"


def test_powerpoint_slide_delete_with_yes(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app, ["powerpoint", "slide", "delete", "3", "--yes"]
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "powerpoint.slide.delete"
    assert params == {"slide": 3, "dry_run": False, "confirm": True}


def test_undo_with_yes(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(
        cli.app, ["undo", "--yes", "--expected-document", "Deck.pptx"]
    )
    assert result.exit_code == 0
    operation, params = client.invocations[0]
    assert operation == "undo"
    assert params == {"dry_run": False, "confirm": True, "expected_document": "Deck.pptx"}


def test_undo_dry_run_does_not_confirm(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["undo", "--dry-run"])
    assert result.exit_code == 0
    _, params = client.invocations[0]
    assert params == {"dry_run": True}


def test_target_prune_and_reconnect(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["target", "prune"])
    assert result.exit_code == 0
    assert client.prune_calls == [True]

    result = CliRunner().invoke(cli.app, ["target", "reconnect", "t-aaaa1111aaaa"])
    assert result.exit_code == 0
    assert client.reconnect_calls == ["t-aaaa1111aaaa"]


def test_com_batch_file_sends_steps(tmp_path, monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    path = tmp_path / "steps.json"
    path.write_text(
        '[{"mode": "get", "expression": "document.Name"},'
        ' {"mode": "set", "root": "document", "expression": "ReadOnly", "value": false}]',
        encoding="utf-8",
    )
    result = CliRunner().invoke(
        cli.app, ["com", "batch", "--file", str(path), "--yes"]
    )
    assert result.exit_code == 0
    steps, confirm, expected = client.batch_calls[0]
    assert confirm is True
    assert expected is None
    assert steps[0] == {
        "mode": "get",
        "root": "window",
        "expression": "document.Name",
        "value": None,
        "dry_run": False,
    }
    assert steps[1]["mode"] == "set"
    assert steps[1]["value"] is False


def test_com_batch_get_only_needs_no_confirm(tmp_path, monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    path = tmp_path / "steps.json"
    path.write_text(
        '[{"mode": "get", "expression": "document.Name"},'
        ' {"mode": "get", "expression": "Slides.Count"}]',
        encoding="utf-8",
    )
    result = CliRunner().invoke(cli.app, ["com", "batch", "--file", str(path)])
    assert result.exit_code == 0
    steps, confirm, _expected = client.batch_calls[0]
    assert confirm is False
    assert steps[0]["dry_run"] is False


def test_com_batch_dry_run_marks_every_step(tmp_path, monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    path = tmp_path / "steps.json"
    path.write_text('[{"mode": "call", "expression": "Rows(2).Delete()"}]', encoding="utf-8")
    result = CliRunner().invoke(
        cli.app, ["com", "batch", "--file", str(path), "--dry-run"]
    )
    assert result.exit_code == 0
    steps, confirm, _expected = client.batch_calls[0]
    assert confirm is False
    assert steps[0]["dry_run"] is True


def test_com_batch_rejects_bad_file(tmp_path, monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    path = tmp_path / "steps.json"
    path.write_text("not json", encoding="utf-8")
    result = CliRunner().invoke(cli.app, ["com", "batch", "--file", str(path), "--yes"])
    assert result.exit_code != 0
    assert client.batch_calls == []


def test_load_values_accepts_values_object() -> None:
    assert cli._load_values('{"values": [[1, 2]]}', None) == [[1, 2]]


def test_load_values_rejects_both_sources(tmp_path) -> None:
    path = tmp_path / "values.json"
    path.write_text("[[1]]", encoding="utf-8")
    with pytest.raises(cli.typer.BadParameter):
        cli._load_values("[[1]]", path)

def test_target_capture_uses_foreground_by_default(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["target", "capture"])
    assert result.exit_code == 0
    assert client.capture_calls == [None]
    assert parse_stdout(result)["result"]["target_id"] == "t-aaaa1111aaaa"


def test_target_capture_passes_hwnd(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["target", "capture", "--hwnd", "1234"])
    assert result.exit_code == 0
    assert client.capture_calls == [1234]


def test_target_list_shows_entries(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["target", "list"])
    assert result.exit_code == 0
    body = parse_stdout(result)["result"]
    assert body["selected_target_id"] == "t-aaaa1111aaaa"
    assert [entry["state"] for entry in body["targets"]] == ["active", "expired"]


def test_target_info_and_remove_pass_ids_through(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["target", "info", "t-aaaa1111aaaa"])
    assert result.exit_code == 0
    assert client.info_calls == ["t-aaaa1111aaaa"]

    result = CliRunner().invoke(cli.app, ["target", "remove", "t-bbbb2222bbbb"])
    assert result.exit_code == 0
    assert client.remove_calls == ["t-bbbb2222bbbb"]

def test_target_option_stored_on_client(monkeypatch) -> None:
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["--target", "t-abc123", "inspect", "selection"])
    assert result.exit_code == 0
    assert client.target_id == "t-abc123"


def test_target_option_reads_env_var(monkeypatch) -> None:
    monkeypatch.setenv("OFFICE_BRIDGE_TARGET", "t-env")
    client = FakeClient()
    install_client(monkeypatch, client)
    result = CliRunner().invoke(cli.app, ["health"])
    assert result.exit_code == 0
    assert client.target_id == "t-env"


def test_bridge_client_injects_target_into_payloads(monkeypatch) -> None:
    captured: list[tuple[str, str, dict]] = []

    class Resp:
        status_code = 200

        def json(self):
            return {"ok": True}

    class FakeHttp:
        def __init__(self, base_url, timeout):
            pass

        def __enter__(self):
            return self

        def __exit__(self, *exc):
            return False

        def request(self, method, path, json=None):
            captured.append((method, path, json))
            return Resp()

    monkeypatch.setattr(cli.httpx, "Client", FakeHttp)
    client = cli.BridgeClient("http://127.0.0.1:1", 1.0, target_id="t-999999999999")
    client.inspect("selection", {})
    client.invoke("excel.range.write", {"address": "A1"})
    client.set_property("excel.active_sheet", "Sheet2", True, None)
    client.com("get", "window", "Name", None, False, False, None)

    assert [path for _m, path, _p in captured] == [
        "/inspect",
        "/invoke",
        "/set-property",
        "/com",
    ]
    for _method, path, payload in captured:
        assert payload["target_id"] == "t-999999999999", path


def test_confirm_write_uses_target_record(monkeypatch) -> None:
    client = FakeClient()
    ctx = SimpleNamespace(obj=cli.ClientOptions(url="u", timeout=1.0, target="t-aaaa1111aaaa"))
    prompts: list = []
    monkeypatch.setattr(cli.typer, "confirm", lambda *args, **kwargs: prompts.append(args) or True)
    assert cli._confirm_write(ctx, client, "write A1:B2", "Book1.xlsx") is True
    assert client.info_calls == ["t-aaaa1111aaaa"]
    assert prompts  # the operator was prompted once


def test_confirm_write_refuses_expired_target(monkeypatch) -> None:
    client = FakeClient()
    client.target_info = lambda target_id: {
        "ok": True,
        "result": {"target_id": target_id, "state": "expired", "expired_reason": "closed"},
    }
    ctx = SimpleNamespace(obj=cli.ClientOptions(url="u", timeout=1.0, target="t-aaaa1111aaaa"))
    monkeypatch.setattr(
        cli.typer, "confirm", lambda *args, **kwargs: pytest.fail("must not prompt")
    )
    assert cli._confirm_write(ctx, client, "write A1:B2", None) is False


def test_confirm_write_target_mismatch_refuses(monkeypatch) -> None:
    client = FakeClient()
    ctx = SimpleNamespace(obj=cli.ClientOptions(url="u", timeout=1.0, target="t-aaaa1111aaaa"))
    monkeypatch.setattr(
        cli.typer, "confirm", lambda *args, **kwargs: pytest.fail("must not prompt")
    )
    assert cli._confirm_write(ctx, client, "write A1:B2", "Other.xlsx") is False
