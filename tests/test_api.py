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
        self.set_calls: list[tuple[str, Any, dict[str, Any], str | None]] = []
        self.com_calls: list[tuple[str, str, str, Any, bool, dict[str, Any], str | None]] = []
        self.com_batch_calls: list[tuple[list[dict[str, Any]], dict[str, Any], str | None]] = []
        self.prune_calls: list[bool] = []
        self.reconnect_calls: list[str] = []
        self.inspect_calls: list[tuple[str, dict[str, Any], str | None]] = []
        self.invoke_calls: list[tuple[str, dict[str, Any], str | None]] = []
        self.capture_calls: list[int | None] = []

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

    def inspect(self, kind: str, params: dict[str, Any], target_id: str | None = None):
        self.inspect_calls.append((kind, params, target_id))
        if kind == "missing":
            raise BridgeError(ErrorCode.TARGET_LOST, "gone", 409)
        return {"kind": kind, "params": params}, dict(self.summary)

    def invoke(self, operation: str, params: dict[str, Any], target_id: str | None = None):
        self.invoke_calls.append((operation, params, target_id))
        if operation == "denied":
            raise BridgeError(ErrorCode.UNSUPPORTED_OPERATION, "denied", 400)
        return {"operation": operation, "params": params}, dict(self.summary)

    def set_property(
        self,
        name: str,
        value: Any,
        params: dict[str, Any],
        target_id: str | None = None,
    ):
        self.set_calls.append((name, value, params, target_id))
        return {"property": name, "value": value}, dict(self.summary)

    def com(
        self,
        mode: str,
        root: str,
        expression: str,
        value: Any,
        dry_run: bool,
        params: dict[str, Any],
        target_id: str | None = None,
    ):
        self.com_calls.append(
            (mode, root, expression, value, dry_run, params, target_id)
        )
        return {"mode": mode, "root": root, "expression": expression}, dict(self.summary)

    def com_batch(
        self,
        steps: list[dict[str, Any]],
        params: dict[str, Any],
        target_id: str | None = None,
    ):
        self.com_batch_calls.append((steps, params, target_id))
        return {"steps": [{"step": i, "ok": True} for i in range(len(steps))], "stopped": False}, dict(self.summary)

    def prune_targets(self):
        self.prune_calls.append(True)
        return {"removed": ["t-expired00000"], "count": 1}

    def reconnect_target(self, target_id: str):
        self.reconnect_calls.append(target_id)
        if target_id == "t-missing00000":
            raise BridgeError(ErrorCode.TARGET_NOT_FOUND, "gone", 404)
        return {"target": {"target_id": target_id, "state": "active"}, "reconnected": True}

    def cached_targets(self):
        return [dict(self.summary)]

    def capture(self, hwnd: int | None = None):
        self.capture_calls.append(hwnd)
        return {
            "target": {
                "target_id": "t-aaaa1111aaaa",
                "app_type": "excel",
                "display_name": "Book1.xlsx",
                "state": "active",
                "expired_reason": None,
            }
        }

    def targets(self):
        return {
            "targets": [
                {
                    "target_id": "t-aaaa1111aaaa",
                    "app_type": "excel",
                    "display_name": "Book1.xlsx",
                    "state": "active",
                    "expired_reason": None,
                }
            ],
            "selected_target_id": "t-aaaa1111aaaa",
        }

    def target_info(self, target_id: str):
        if target_id == "t-missing00000":
            raise BridgeError(ErrorCode.TARGET_NOT_FOUND, "gone", 404)
        return {"target": {"target_id": target_id, "state": "active"}}

    def remove_target(self, target_id: str):
        if target_id == "t-missing00000":
            raise BridgeError(ErrorCode.TARGET_NOT_FOUND, "gone", 404)
        return {"removed": target_id, "target": {"target_id": target_id}}


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
            None,
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
            None,
        )
    ]


def test_com_rejects_missing_expression() -> None:
    client = TestClient(create_app(FakeCore(), Settings()))
    response = client.post("/com", json={"mode": "get"})
    assert response.status_code == 400
    assert response.json()["ok"] is False


def test_com_batch_passes_steps_and_flags() -> None:
    core = FakeCore()
    client = TestClient(create_app(core, Settings()))
    response = client.post(
        "/com/batch",
        json={
            "steps": [
                {"mode": "get", "expression": "Name"},
                {"mode": "call", "root": "document", "expression": "Reorder()"},
            ],
            "confirm": True,
            "expected_document": "Book1.xlsx",
        },
    )
    assert response.status_code == 200
    body = response.json()
    assert body["ok"] is True
    assert body["result"]["stopped"] is False
    steps, params, target_id = core.com_batch_calls[0]
    assert steps[0]["expression"] == "Name"
    assert steps[1]["mode"] == "call"
    assert steps[1]["root"] == "document"
    assert params == {"confirm": True, "expected_document": "Book1.xlsx"}
    assert target_id is None


def test_com_batch_requires_steps() -> None:
    client = TestClient(create_app(FakeCore(), Settings()))
    response = client.post("/com/batch", json={"steps": []})
    assert response.status_code == 400
    assert response.json()["ok"] is False


def test_targets_prune_and_reconnect_endpoints() -> None:
    core = FakeCore()
    client = TestClient(create_app(core, Settings()))
    response = client.post("/targets/prune")
    assert response.status_code == 200
    assert response.json()["result"] == {"removed": ["t-expired00000"], "count": 1}

    response = client.post("/targets/t-aaaa1111aaaa/reconnect")
    assert response.status_code == 200
    assert response.json()["result"]["reconnected"] is True
    assert core.reconnect_calls == ["t-aaaa1111aaaa"]

    response = client.post("/targets/t-missing00000/reconnect")
    assert response.status_code == 404


def test_health_includes_target_count(client: TestClient) -> None:
    response = client.get("/health")
    assert response.status_code == 200
    assert response.json()["result"]["target_count"] == 1


def test_capture_target_returns_record(client: TestClient) -> None:
    response = client.post("/targets/capture", json={})
    assert response.status_code == 200
    target = response.json()["result"]
    assert target["target_id"] == "t-aaaa1111aaaa"
    assert target["state"] == "active"


def test_capture_target_passes_hwnd() -> None:
    core = FakeCore()
    client = TestClient(create_app(core, Settings()))
    response = client.post("/targets/capture", json={"hwnd": 1234})
    assert response.status_code == 200
    assert core.capture_calls == [1234]


def test_list_targets(client: TestClient) -> None:
    response = client.get("/targets")
    assert response.status_code == 200
    body = response.json()["result"]
    assert body["selected_target_id"] == "t-aaaa1111aaaa"
    assert body["targets"][0]["target_id"] == "t-aaaa1111aaaa"


def test_target_info_and_missing_id(client: TestClient) -> None:
    response = client.get("/targets/t-aaaa1111aaaa")
    assert response.status_code == 200
    assert response.json()["result"]["target_id"] == "t-aaaa1111aaaa"

    response = client.get("/targets/t-missing00000")
    assert response.status_code == 404
    assert response.json()["error"]["code"] == ErrorCode.TARGET_NOT_FOUND


def test_remove_target(client: TestClient) -> None:
    response = client.delete("/targets/t-aaaa1111aaaa")
    assert response.status_code == 200
    assert response.json()["result"]["removed"] == "t-aaaa1111aaaa"

    response = client.delete("/targets/t-missing00000")
    assert response.status_code == 404
    assert response.json()["error"]["code"] == ErrorCode.TARGET_NOT_FOUND


def test_operations_pass_target_id_to_core() -> None:
    core = FakeCore()
    client = TestClient(create_app(core, Settings()))

    response = client.post(
        "/inspect",
        json={"kind": "excel.range", "params": {"address": "A1"}, "target_id": "t-84102044b1b6"},
    )
    assert response.status_code == 200
    assert core.inspect_calls[-1] == ("excel.range", {"address": "A1"}, "t-84102044b1b6")
    # The envelope context is the summary of the target the operation used.
    assert response.json()["context"]["document"] == "Book1.xlsx"

    response = client.post(
        "/invoke",
        json={
            "operation": "excel.range.write",
            "params": {"address": "A1", "values": [[1]]},
            "target_id": "t-write0000000",
        },
    )
    assert response.status_code == 200
    assert core.invoke_calls[-1][2] == "t-write0000000"

    response = client.post(
        "/com",
        json={"mode": "get", "expression": "Name", "target_id": "t-read00000000"},
    )
    assert response.status_code == 200
    assert core.com_calls[-1][6] == "t-read00000000"

    response = client.post(
        "/set-property",
        json={
            "property": "excel.active_sheet",
            "value": "Sheet2",
            "confirm": True,
            "target_id": "t-prop00000000",
        },
    )
    assert response.status_code == 200
    assert core.set_calls[-1][3] == "t-prop00000000"
