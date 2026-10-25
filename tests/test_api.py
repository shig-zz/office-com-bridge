from __future__ import annotations

from typing import Any

import pytest
from fastapi.testclient import TestClient

from office_bridge.api import create_app
from office_bridge.config import Settings
from office_bridge.errors import BridgeError, ErrorCode


class FakeWorker:
    is_alive = True


class FakeCore:
    def __init__(self) -> None:
        self.worker = FakeWorker()
        self.summary = {"application": "Excel", "document": "Book1.xlsx"}
        self.set_calls: list[tuple[str, Any, dict[str, Any]]] = []
        self.com_calls: list[tuple[str, str, str, Any, bool, dict[str, Any]]] = []

    def cached_summary(self):
        return dict(self.summary)

    def context(self):
        return {
            "application": "Excel",
            "document": "Book1.xlsx",
            "full_name": r"C:\docs\Book1.xlsx",
        }

    def capabilities(self):
        return {"inspections": [{"name": "selection"}], "operations": []}

    def inspect(self, kind: str, params: dict[str, Any]):
        if kind == "missing":
            raise BridgeError(ErrorCode.TARGET_LOST, "gone", 409)
        return {"kind": kind, "params": params}

    def invoke(self, operation: str, params: dict[str, Any]):
        if operation == "denied":
            raise BridgeError(ErrorCode.UNSUPPORTED_OPERATION, "denied", 400)
        return {"operation": operation, "params": params}

    def set_property(self, name: str, value: Any, params: dict[str, Any]):
        self.set_calls.append((name, value, params))
        return {"property": name, "value": value}

    def com(
        self,
        mode: str,
        root: str,
        expression: str,
        value: Any,
        dry_run: bool,
        params: dict[str, Any],
    ):
        self.com_calls.append((mode, root, expression, value, dry_run, params))
        return {"mode": mode, "root": root, "expression": expression}


@pytest.fixture()
def client() -> TestClient:
    app = create_app(FakeCore(), Settings())
    return TestClient(app)


def test_health_includes_target_and_version(client: TestClient) -> None:
    response = client.get("/health")
    assert response.status_code == 200
    body = response.json()
    assert body["ok"] is True
    assert body["result"]["status"] == "running"
    assert body["result"]["target"]["document"] == "Book1.xlsx"
    assert body["context"]["application"] == "Excel"


def test_context_returns_target(client: TestClient) -> None:
    response = client.get("/context")
    assert response.status_code == 200
    assert response.json()["result"]["document"] == "Book1.xlsx"


def test_capabilities(client: TestClient) -> None:
    response = client.get("/capabilities")
    assert response.status_code == 200
    body = response.json()["result"]
    assert body["inspections"][0]["name"] == "selection"
    assert body["version"]


def test_inspect_passes_kind_and_params(client: TestClient) -> None:
    response = client.post("/inspect", json={"kind": "selection", "params": {"x": 1}})
    assert response.status_code == 200
    assert response.json()["result"] == {"kind": "selection", "params": {"x": 1}}


def test_bridge_error_is_rendered_as_envelope(client: TestClient) -> None:
    response = client.post("/inspect", json={"kind": "missing"})
    assert response.status_code == 409
    body = response.json()
    assert body["ok"] is False
    assert body["error"]["code"] == ErrorCode.TARGET_LOST
    assert body["context"]["document"] == "Book1.xlsx"


def test_validation_error_is_rendered_as_envelope(client: TestClient) -> None:
    response = client.post("/inspect", json={"params": {}})
    assert response.status_code == 400
    body = response.json()
    assert body["ok"] is False
    assert body["error"]["code"] == ErrorCode.INVALID_REQUEST
    assert body["error"]["details"]["errors"]


def test_unknown_route_is_json(client: TestClient) -> None:
    response = client.get("/does-not-exist")
    assert response.status_code == 404
    assert response.json()["ok"] is False


def test_invoke_error(client: TestClient) -> None:
    response = client.post("/invoke", json={"operation": "denied", "params": {}})
    assert response.status_code == 400
    assert response.json()["error"]["code"] == ErrorCode.UNSUPPORTED_OPERATION


def test_set_property_passes_confirm_and_expected_document() -> None:
    core = FakeCore()
    client = TestClient(create_app(core, Settings()))
    response = client.post(
        "/set-property",
        json={
            "property": "excel.active_sheet",
            "value": "Sheet2",
            "confirm": True,
            "expected_document": "Book1.xlsx",
        },
    )
    assert response.status_code == 200
    assert core.set_calls == [
        (
            "excel.active_sheet",
            "Sheet2",
            {"confirm": True, "expected_document": "Book1.xlsx"},
        )
    ]


def test_com_passes_mode_root_expression_and_flags() -> None:
    core = FakeCore()
    client = TestClient(create_app(core, Settings()))
    response = client.post(
        "/com",
        json={
            "mode": "set",
            "root": "document",
            "expression": "Paragraphs(1).Range.Text",
            "value": "hello",
            "dry_run": True,
            "expected_document": "Book1.xlsx",
        },
    )
    assert response.status_code == 200
    body = response.json()
    assert body["ok"] is True
    assert body["result"]["mode"] == "set"
    assert core.com_calls == [
        (
            "set",
            "document",
            "Paragraphs(1).Range.Text",
            "hello",
            True,
            {"confirm": False, "expected_document": "Book1.xlsx"},
        )
    ]


def test_com_rejects_missing_expression() -> None:
    client = TestClient(create_app(FakeCore(), Settings()))
    response = client.post("/com", json={"mode": "get"})
    assert response.status_code == 400
    assert response.json()["ok"] is False
