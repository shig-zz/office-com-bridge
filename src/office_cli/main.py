from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

import httpx
import typer

from office_bridge.config import DEFAULT_PORT

app = typer.Typer(
    name="office",
    help="Query and edit the Microsoft Office document selected by the office bridge.",
    no_args_is_help=True,
    add_completion=False,
)

excel_app = typer.Typer(help="Excel commands.", no_args_is_help=True)
excel_range_app = typer.Typer(help="Excel range commands.", no_args_is_help=True)
excel_workbook_app = typer.Typer(help="Excel workbook commands.", no_args_is_help=True)
word_app = typer.Typer(help="Word commands.", no_args_is_help=True)
word_selection_app = typer.Typer(help="Word selection commands.", no_args_is_help=True)
word_document_app = typer.Typer(help="Word document commands.", no_args_is_help=True)
powerpoint_app = typer.Typer(help="PowerPoint commands.", no_args_is_help=True)
powerpoint_slide_app = typer.Typer(help="PowerPoint slide commands.", no_args_is_help=True)
powerpoint_slides_app = typer.Typer(
    help="PowerPoint presentation-wide slide commands.", no_args_is_help=True
)
powerpoint_text_app = typer.Typer(help="PowerPoint text commands.", no_args_is_help=True)
powerpoint_textbox_app = typer.Typer(
    help="PowerPoint text box commands.", no_args_is_help=True
)
powerpoint_document_app = typer.Typer(
    help="PowerPoint presentation commands.", no_args_is_help=True
)
com_app = typer.Typer(
    help="General Office object-model commands (get, call, set).", no_args_is_help=True
)
target_app = typer.Typer(
    help="Manage the documents registered with the bridge.", no_args_is_help=True
)

app.add_typer(excel_app, name="excel")
excel_app.add_typer(excel_range_app, name="range")
excel_app.add_typer(excel_workbook_app, name="workbook")
app.add_typer(word_app, name="word")
word_app.add_typer(word_selection_app, name="selection")
word_app.add_typer(word_document_app, name="document")
app.add_typer(powerpoint_app, name="powerpoint")
powerpoint_app.add_typer(powerpoint_slide_app, name="slide")
powerpoint_app.add_typer(powerpoint_slides_app, name="slides")
powerpoint_app.add_typer(powerpoint_text_app, name="text")
powerpoint_app.add_typer(powerpoint_textbox_app, name="textbox")
powerpoint_app.add_typer(powerpoint_document_app, name="document")
app.add_typer(com_app, name="com")
app.add_typer(target_app, name="target")


@dataclass(slots=True)
class ClientOptions:
    url: str
    timeout: float
    target: str | None = None


class BridgeUnreachable(Exception):
    pass


class BridgeClient:
    def __init__(self, base_url: str, timeout: float, target_id: str | None = None) -> None:
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.target_id = target_id

    def _with_target(self, payload: dict[str, Any]) -> dict[str, Any]:
        if self.target_id:
            payload = {**payload, "target_id": self.target_id}
        return payload

    def request(
        self,
        method: str,
        path: str,
        payload: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        try:
            with httpx.Client(base_url=self.base_url, timeout=self.timeout) as client:
                response = client.request(method, path, json=payload)
        except httpx.HTTPError as exc:
            raise BridgeUnreachable(
                f"Could not reach the office bridge at {self.base_url}: {exc}"
            ) from exc

        try:
            data = response.json()
        except ValueError as exc:
            raise BridgeUnreachable(
                f"The office bridge returned a non-JSON response (HTTP {response.status_code})."
            ) from exc
        if not isinstance(data, dict):
            raise BridgeUnreachable("The office bridge returned an unexpected JSON payload.")
        return data

    def health(self) -> dict[str, Any]:
        return self.request("GET", "/health")

    def context(self) -> dict[str, Any]:
        return self.request("GET", "/context")

    def capabilities(self) -> dict[str, Any]:
        return self.request("GET", "/capabilities")

    def capture(self, hwnd: int | None = None) -> dict[str, Any]:
        payload: dict[str, Any] = {}
        if hwnd is not None:
            payload["hwnd"] = hwnd
        return self.request("POST", "/targets/capture", payload)

    def targets(self) -> dict[str, Any]:
        return self.request("GET", "/targets")

    def target_info(self, target_id: str) -> dict[str, Any]:
        return self.request("GET", f"/targets/{target_id}")

    def remove_target(self, target_id: str) -> dict[str, Any]:
        return self.request("DELETE", f"/targets/{target_id}")

    def inspect(self, kind: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        return self.request(
            "POST", "/inspect", self._with_target({"kind": kind, "params": params or {}})
        )

    def invoke(self, operation: str, params: dict[str, Any]) -> dict[str, Any]:
        return self.request(
            "POST", "/invoke", self._with_target({"operation": operation, "params": params})
        )

    def set_property(
        self,
        name: str,
        value: Any,
        confirm: bool,
        expected_document: str | None,
    ) -> dict[str, Any]:
        payload: dict[str, Any] = {"property": name, "value": value, "confirm": confirm}
        if expected_document:
            payload["expected_document"] = expected_document
        return self.request("POST", "/set-property", self._with_target(payload))

    def com(
        self,
        mode: str,
        root: str,
        expression: str,
        value: Any,
        dry_run: bool,
        confirm: bool,
        expected_document: str | None,
    ) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "mode": mode,
            "root": root,
            "expression": expression,
            "dry_run": dry_run,
            "confirm": confirm,
        }
        if value is not None:
            payload["value"] = value
        if expected_document:
            payload["expected_document"] = expected_document
        return self.request("POST", "/com", self._with_target(payload))


def _default_url() -> str:
    configured = os.environ.get("OFFICE_BRIDGE_URL")
    if configured:
        return configured
    port = os.environ.get("OFFICE_BRIDGE_PORT", str(DEFAULT_PORT))
    return f"http://127.0.0.1:{port}"


@app.callback()
def configure(
    ctx: typer.Context,
    url: str = typer.Option(
        None,
        "--url",
        envvar="OFFICE_BRIDGE_URL",
        help="Base URL of the office bridge.",
    ),
    timeout: float = typer.Option(
        180.0,
        "--timeout",
        envvar="OFFICE_BRIDGE_CLI_TIMEOUT",
        help="HTTP timeout in seconds.",
    ),
    target: str = typer.Option(
        None,
        "--target",
        envvar="OFFICE_BRIDGE_TARGET",
        help=(
            "Target id (from 'office target list') that this command operates on. "
            "Omit to use the last captured target."
        ),
    ),
) -> None:
    ctx.obj = ClientOptions(url=url or _default_url(), timeout=timeout, target=target)


def _emit(payload: dict[str, Any]) -> None:
    typer.echo(json.dumps(payload, ensure_ascii=False, indent=2, default=str))


def _run(ctx: typer.Context, action: Callable[[BridgeClient], dict[str, Any]]) -> None:
    options: ClientOptions = ctx.obj
    client = BridgeClient(options.url, options.timeout, options.target)
    try:
        payload = action(client)
    except BridgeUnreachable as exc:
        payload = {
            "ok": False,
            "context": None,
            "error": {"code": "BRIDGE_UNREACHABLE", "message": str(exc)},
        }
        _emit(payload)
        typer.echo(str(exc), err=True)
        raise typer.Exit(2) from exc

    _emit(payload)
    if not payload.get("ok", False):
        raise typer.Exit(1)


def _parse_param_values(raw: list[str]) -> dict[str, Any]:
    params: dict[str, Any] = {}
    for item in raw:
        if "=" not in item:
            raise typer.BadParameter(f"Expected key=value, got {item!r}.")
        key, value = item.split("=", 1)
        key = key.strip()
        if not key:
            raise typer.BadParameter(f"Empty parameter name in {item!r}.")
        params[key] = _parse_json_or_string(value)
    return params


def _parse_json_or_string(value: str) -> Any:
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return value


def _resolve_value_argument(value: str) -> Any:
    """Parse a --value argument, reading it from a UTF-8 file when prefixed with '@'.

    Long or multi-line text (Japanese text, PowerPoint paragraphs) is unsafe to pass
    through a shell command line, so '@path' loads the value from a file instead.
    """
    if not value.startswith("@"):
        return _parse_json_or_string(value)
    path = Path(value[1:])
    try:
        raw = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise typer.BadParameter(f"Could not read {path}: {exc}") from exc
    return _parse_json_or_string(raw)


def _shape_argument(value: str) -> Any:
    text = value.strip()
    if text.isdigit():
        return int(text)
    return value


def _resolve_text_argument(value: str) -> str:
    """Keep PowerPoint text as text; '@path' loads it from a UTF-8 file."""
    if not value.startswith("@"):
        return value
    path = Path(value[1:])
    try:
        return path.read_text(encoding="utf-8")
    except OSError as exc:
        raise typer.BadParameter(f"Could not read {path}: {exc}") from exc


def _load_values(json_text: str | None, json_file: Path | None) -> Any:
    if bool(json_text) == bool(json_file):
        raise typer.BadParameter("Provide exactly one of --json or --file.")
    try:
        if json_file is not None:
            raw = json_file.read_text(encoding="utf-8")
        else:
            raw = json_text or ""
        data = json.loads(raw)
    except OSError as exc:
        raise typer.BadParameter(f"Could not read {json_file}: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise typer.BadParameter(f"Invalid JSON: {exc}") from exc
    if isinstance(data, dict) and "values" in data:
        data = data["values"]
    return data


def _matrix_summary(values: Any) -> str:
    if not isinstance(values, list) or not values:
        return "invalid array"
    rows = len(values)
    columns = len(values[0]) if isinstance(values[0], list) else 0
    return f"{rows}x{columns} ({rows * columns} cells)"


def _reject_unsafe_console() -> None:
    if sys.stdin is None or not sys.stdin.isatty():
        raise typer.BadParameter(
            "Confirmation requires an interactive terminal. Re-run with --yes after "
            "confirming the target document and change with the user."
        )


def _confirm_write(
    ctx: typer.Context,
    client: BridgeClient,
    action_description: str,
    expected_document: str | None,
) -> bool:
    options: ClientOptions = ctx.obj
    if options.target:
        response = client.target_info(options.target)
        if not response.get("ok"):
            _emit(response)
            raise typer.Exit(1)
        record = response.get("result") or {}
        if record.get("state") != "active":
            typer.echo(
                f"Target {options.target} is {record.get('state')} "
                f"({record.get('expired_reason') or 'unavailable'}); refusing to write.",
                err=True,
            )
            return False
        actual_document = str(record.get("display_name", ""))
        actual_application = str(record.get("app_type", ""))
        actual_full_name = str((record.get("metadata") or {}).get("full_name", ""))
        label = f"{options.target} ({actual_application} / {actual_document})"
    else:
        context_response = client.context()
        if not context_response.get("ok"):
            _emit(context_response)
            raise typer.Exit(1)
        result = context_response.get("result") or {}
        actual_document = str(result.get("document", ""))
        actual_application = str(result.get("application", ""))
        actual_full_name = str(result.get("full_name", ""))
        label = f"last selected target (no --target): {actual_application} / {actual_document}"
        listing = client.targets()
        targets = (listing.get("result") or {}).get("targets") or []
        active = [entry for entry in targets if entry.get("state") == "active"]
        if len(active) > 1:
            typer.echo(
                f"Warning: {len(active)} targets are registered; pass --target to be "
                "explicit about the write destination.",
                err=True,
            )
    if expected_document and expected_document.lower() not in {
        actual_document.lower(),
        actual_full_name.lower(),
    }:
        typer.echo(
            f"Target mismatch: expected {expected_document!r}, bridge is attached to "
            f"{actual_document!r}.",
            err=True,
        )
        return False
    typer.echo(
        f"Target: {label}\nAction: {action_description}",
        err=True,
    )
    return typer.confirm("Proceed with this change?", default=False)


@app.command()
def health(ctx: typer.Context) -> None:
    """Check whether the bridge is running."""
    _run(ctx, lambda client: client.health())


@app.command()
def context(ctx: typer.Context) -> None:
    """Show the current Office target and document information."""
    _run(ctx, lambda client: client.context())


@app.command()
def capabilities(ctx: typer.Context) -> None:
    """List the allowed inspection and write operations."""
    _run(ctx, lambda client: client.capabilities())


@target_app.command("capture")
def target_capture(
    ctx: typer.Context,
    hwnd: int = typer.Option(
        None,
        "--hwnd",
        help="Window handle to attach instead of the foreground window (diagnostics).",
    ),
) -> None:
    """Register the foreground Office window (same as the bridge hotkey)."""
    _run(ctx, lambda client: client.capture(hwnd))


@target_app.command("list")
def target_list(ctx: typer.Context) -> None:
    """List the documents registered with the bridge and their state."""
    _run(ctx, lambda client: client.targets())


@target_app.command("info")
def target_info(
    ctx: typer.Context,
    target_id: str = typer.Argument(..., help="Target id printed by 'office target list'."),
) -> None:
    """Show one registered document, refreshed and validated against Office."""
    _run(ctx, lambda client: client.target_info(target_id))


@target_app.command("remove")
def target_remove(
    ctx: typer.Context,
    target_id: str = typer.Argument(..., help="Target id to unregister."),
) -> None:
    """Unregister a document from the bridge."""
    _run(ctx, lambda client: client.remove_target(target_id))


@app.command()
def inspect(
    ctx: typer.Context,
    kind: str = typer.Argument(
        ...,
        help="Inspection kind: selection, document, excel.range, excel.workbook, word.selection, word.document, powerpoint.slide, powerpoint.slides, powerpoint.document.",
    ),
    param: list[str] = typer.Option(
        [],
        "--param",
        help="Extra parameter in key=value form (JSON values accepted).",
    ),
) -> None:
    """Run an allowed read operation."""
    params = _parse_param_values(param)
    _run(ctx, lambda client: client.inspect(kind, params))


@excel_range_app.command("read")
def excel_range_read(
    ctx: typer.Context,
    address: str = typer.Argument(..., help="A1-style range, for example A1:D20."),
    formulas: bool = typer.Option(
        False, "--formulas", help="Include formula strings alongside values."
    ),
) -> None:
    """Read a range from the active sheet of the attached workbook."""
    params = {"address": address, "include_formulas": formulas}
    _run(ctx, lambda client: client.inspect("excel.range", params))


@excel_range_app.command("write")
def excel_range_write(
    ctx: typer.Context,
    address: str = typer.Argument(..., help="A1-style range, for example A1:B2."),
    json_text: str = typer.Option(
        None, "--json", help="Two-dimensional JSON array of literal values."
    ),
    json_file: Path = typer.Option(
        None, "--file", help="Path to a JSON file containing a two-dimensional array."
    ),
    dry_run: bool = typer.Option(
        False, "--dry-run", help="Validate the range and array without writing."
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to write unless the attached document matches this name.",
    ),
) -> None:
    """Write a two-dimensional array of literal values to an Excel range."""
    values = _load_values(json_text, json_file)
    params: dict[str, Any] = {
        "address": address,
        "values": values,
        "dry_run": dry_run,
    }
    if expected_document:
        params["expected_document"] = expected_document

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"write {address} with {_matrix_summary(values)}",
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            params["confirm"] = True
        return client.invoke("excel.range.write", params)

    _run(ctx, action)


@excel_range_app.command("copy-format")
def excel_range_copy_format(
    ctx: typer.Context,
    source_address: str = typer.Argument(
        ..., help="Source range that has the desired formatting, for example A1:B2."
    ),
    address: str = typer.Argument(
        ..., help="Target range that receives the formatting (same dimensions)."
    ),
    dry_run: bool = typer.Option(
        False, "--dry-run", help="Validate the ranges without changing formatting."
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to write unless the attached document matches this name.",
    ),
) -> None:
    """Copy formatting only (font, fill, borders, number format) between ranges."""
    params: dict[str, Any] = {
        "source_address": source_address,
        "address": address,
        "dry_run": dry_run,
    }
    if expected_document:
        params["expected_document"] = expected_document

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"copy formatting from {source_address} to {address} (values unchanged)",
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            params["confirm"] = True
        return client.invoke("excel.range.copy_format", params)

    _run(ctx, action)


@excel_range_app.command("set-format")
def excel_range_set_format(
    ctx: typer.Context,
    address: str = typer.Argument(..., help="A1-style range, for example A1:B2."),
    font_name: str = typer.Option(
        None,
        "--font-name",
        help="Font name, for example 'MS Gothic'. Use a pan-CJK font to cover both Japanese and Western text.",
    ),
    font_size: float = typer.Option(None, "--font-size", help="Font size in points."),
    bold: bool = typer.Option(None, "--bold/--no-bold", help="Bold on or off."),
    italic: bool = typer.Option(None, "--italic/--no-italic", help="Italic on or off."),
    color: str = typer.Option(None, "--color", help="Font color as RRGGBB hex, for example FF0000."),
    number_format: str = typer.Option(
        None, "--number-format", help="Excel number format code, for example 'yyyy/m/d'."
    ),
    halign: str = typer.Option(None, "--halign", help="Horizontal alignment: left, center, right."),
    valign: str = typer.Option(None, "--valign", help="Vertical alignment: top, middle, bottom."),
    wrap: bool = typer.Option(None, "--wrap/--no-wrap", help="Wrap text on or off."),
    fill: str = typer.Option(None, "--fill", help="Fill color as RRGGBB hex, for example FFFF00."),
    border: bool = typer.Option(None, "--border/--no-border", help="Add a continuous border."),
    border_color: str = typer.Option(None, "--border-color", help="Border color as RRGGBB hex."),
    dry_run: bool = typer.Option(
        False, "--dry-run", help="Validate the options without changing formatting."
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to write unless the attached document matches this name.",
    ),
) -> None:
    """Set explicit formatting properties on an Excel range."""
    params: dict[str, Any] = {"address": address, "dry_run": dry_run}
    if font_name is not None:
        params["font_name"] = font_name
    if font_size is not None:
        params["font_size"] = font_size
    if bold is not None:
        params["bold"] = bold
    if italic is not None:
        params["italic"] = italic
    if color is not None:
        params["color"] = color
    if number_format is not None:
        params["number_format"] = number_format
    if halign is not None:
        params["halign"] = halign
    if valign is not None:
        params["valign"] = valign
    if wrap is not None:
        params["wrap"] = wrap
    if fill is not None:
        params["fill"] = fill
    if border is not None:
        params["border"] = border
    if border_color is not None:
        params["border_color"] = border_color
    if len(params) <= 2:
        raise typer.BadParameter(
            "Provide at least one formatting option "
            "(--font-name, --font-size, --bold, --italic, --color, --number-format, "
            "--halign, --valign, --wrap, --fill, --border, --border-color)."
        )
    if expected_document:
        params["expected_document"] = expected_document

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"set formatting on {address}: "
                    + ", ".join(f"{k}={v!r}" for k, v in params.items() if k not in {"address", "dry_run"}),
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            params["confirm"] = True
        return client.invoke("excel.range.set_format", params)

    _run(ctx, action)


@excel_range_app.command("add-hyperlink")
def excel_range_add_hyperlink(
    ctx: typer.Context,
    address: str = typer.Argument(..., help="A1-style range, for example E35."),
    url: str = typer.Option(
        ...,
        "--url",
        help="Hyperlink destination, for example 'https://example.com/'.",
    ),
    sub_address: str = typer.Option(
        None,
        "--sub-address",
        help="Anchor inside the destination or a sheet reference, for example 'Sheet2!A1'.",
    ),
    screen_tip: str = typer.Option(
        None, "--screen-tip", help="Tooltip text shown when hovering over the link."
    ),
    dry_run: bool = typer.Option(
        False, "--dry-run", help="Validate the range without adding the link."
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to write unless the attached document matches this name.",
    ),
) -> None:
    """Add a hyperlink object to an Excel range (or update the existing one)."""
    params: dict[str, Any] = {"address": address, "url": url, "dry_run": dry_run}
    if sub_address is not None:
        params["sub_address"] = sub_address
    if screen_tip is not None:
        params["screen_tip"] = screen_tip
    if expected_document:
        params["expected_document"] = expected_document

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"add hyperlink {url} to {address}",
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            params["confirm"] = True
        return client.invoke("excel.range.add_hyperlink", params)

    _run(ctx, action)


@excel_range_app.command("sort")
def excel_range_sort(
    ctx: typer.Context,
    address: str = typer.Argument(
        ..., help="A1-style range to sort, for example A2:E54 (exclude header rows)."
    ),
    column: list[str] = typer.Option(
        ...,
        "--column",
        "-c",
        help=(
            "Sort column: a number (1-based from the range's left edge) or a sheet column "
            "letter such as C. Repeat for up to 3 keys, in priority order."
        ),
    ),
    order: list[str] = typer.Option(
        (),
        "--order",
        "-o",
        help="ascending or descending for the corresponding key; defaults to ascending.",
    ),
    dry_run: bool = typer.Option(
        False, "--dry-run", help="Report the planned row order without changing the sheet."
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to write unless the attached document matches this name.",
    ),
) -> None:
    """Sort the rows of an Excel range in place (stable, up to 3 keys)."""
    if not 1 <= len(column) <= 3:
        raise typer.BadParameter("Provide between 1 and 3 --column options.")
    if len(order) > len(column):
        raise typer.BadParameter("--order appears once per --column, in the same order.")
    keys: list[dict[str, Any]] = []
    for index, col in enumerate(column):
        key: dict[str, Any] = {"column": int(col) if col.isdigit() else col}
        if index < len(order):
            key["order"] = order[index].lower()
        keys.append(key)
    params: dict[str, Any] = {"address": address, "keys": keys, "dry_run": dry_run}
    if expected_document:
        params["expected_document"] = expected_document

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"sort the rows of {address} by "
                    + ", ".join(
                        f"{key['column']} {key.get('order', 'ascending')}" for key in keys
                    ),
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            params["confirm"] = True
        return client.invoke("excel.range.sort", params)

    _run(ctx, action)


@excel_workbook_app.command("inspect")
def excel_workbook_inspect(ctx: typer.Context) -> None:
    """Read workbook metadata and worksheet names."""
    _run(ctx, lambda client: client.inspect("excel.workbook", {}))


@word_selection_app.command("read")
def word_selection_read(ctx: typer.Context) -> None:
    """Read the selected text in Word."""
    _run(ctx, lambda client: client.inspect("word.selection", {}))


@word_selection_app.command("insert-text")
def word_selection_insert_text(
    ctx: typer.Context,
    text: str = typer.Argument(..., help="Text to insert."),
    replace: bool = typer.Option(
        True,
        "--replace/--no-replace",
        help="Replace the selection (default) or collapse to its start and insert.",
    ),
    dry_run: bool = typer.Option(
        False, "--dry-run", help="Validate without changing the document."
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to write unless the attached document matches this name.",
    ),
) -> None:
    """Insert or replace text at the current Word selection without saving."""
    params: dict[str, Any] = {
        "text": text,
        "replace": replace,
        "dry_run": dry_run,
    }
    if expected_document:
        params["expected_document"] = expected_document

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"insert {len(text)} characters at the selection "
                    f"(replace={replace})",
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            params["confirm"] = True
        return client.invoke("word.selection.insert_text", params)

    _run(ctx, action)


@word_document_app.command("read")
def word_document_read(ctx: typer.Context) -> None:
    """Read Word document metadata and basic counts."""
    _run(ctx, lambda client: client.inspect("word.document", {}))


@powerpoint_slide_app.command("inspect")
def powerpoint_slide_inspect(ctx: typer.Context) -> None:
    """Read the current PowerPoint slide and its shape texts."""
    _run(ctx, lambda client: client.inspect("powerpoint.slide", {}))


@powerpoint_document_app.command("read")
def powerpoint_document_read(ctx: typer.Context) -> None:
    """Read presentation metadata and slide titles."""
    _run(ctx, lambda client: client.inspect("powerpoint.document", {}))


@powerpoint_slides_app.command("read")
def powerpoint_slides_read(
    ctx: typer.Context,
    start: int = typer.Option(
        1, "--start", help="1-based first slide to read."
    ),
    count: int = typer.Option(
        20, "--count", help="Number of slides to read (max 100)."
    ),
) -> None:
    """Read a window of slides with notes and shape geometry in points."""
    _run(ctx, lambda client: client.inspect("powerpoint.slides", {"start": start, "count": count}))


@powerpoint_text_app.command("replace")
def powerpoint_text_replace(
    ctx: typer.Context,
    slide: int = typer.Argument(..., help="1-based slide number."),
    shape: str = typer.Option(
        ...,
        "--shape",
        help="1-based shape index or shape name on the slide.",
    ),
    find: str = typer.Option(..., "--find", help="Text to find in the shape."),
    replace: str = typer.Option(
        ...,
        "--replace",
        help="Replacement text, or '@path' to read it from a UTF-8 file.",
    ),
    dry_run: bool = typer.Option(
        False,
        "--dry-run",
        help="Report how many times the text occurs without changing it.",
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to write unless the attached document matches this name.",
    ),
) -> None:
    """Replace text inside one shape in place, keeping formatting."""
    params: dict[str, Any] = {
        "slide": slide,
        "shape": _shape_argument(shape),
        "find": find,
        "replace": _resolve_text_argument(replace),
        "dry_run": dry_run,
    }
    if expected_document:
        params["expected_document"] = expected_document

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"replace {find!r} with {params['replace']!r} on slide {slide}",
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            params["confirm"] = True
        return client.invoke("powerpoint.text.replace", params)

    _run(ctx, action)


@powerpoint_textbox_app.command("add")
def powerpoint_textbox_add(
    ctx: typer.Context,
    slide: int = typer.Argument(..., help="1-based slide number."),
    left: float = typer.Option(..., "--left", help="Left position in points."),
    top: float = typer.Option(..., "--top", help="Top position in points."),
    width: float = typer.Option(..., "--width", help="Width in points."),
    height: float = typer.Option(..., "--height", help="Height in points."),
    text: str = typer.Option(
        ...,
        "--text",
        help="Text to write, or '@path' to read it from a UTF-8 file.",
    ),
    font_size: float = typer.Option(
        None, "--font-size", help="Optional font size in points."
    ),
    dry_run: bool = typer.Option(
        False,
        "--dry-run",
        help="Validate the position without adding the text box.",
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to write unless the attached document matches this name.",
    ),
) -> None:
    """Add a text box to a slide. Geometry is in points."""
    params: dict[str, Any] = {
        "slide": slide,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "text": _resolve_text_argument(text),
        "dry_run": dry_run,
    }
    if font_size is not None:
        params["font_size"] = font_size
    if expected_document:
        params["expected_document"] = expected_document

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"add a text box on slide {slide} at ({left}, {top}) "
                    f"{width}x{height} pt with {len(params['text'])} characters",
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            params["confirm"] = True
        return client.invoke("powerpoint.textbox.add", params)

    _run(ctx, action)


@powerpoint_slide_app.command("add")
def powerpoint_slide_add(
    ctx: typer.Context,
    index: int = typer.Option(
        None,
        "--index",
        help="1-based position for the new slide; omit to append at the end.",
    ),
    layout: int = typer.Option(
        None,
        "--layout",
        help="Layout number (default 2, Title and Content).",
    ),
    dry_run: bool = typer.Option(
        False, "--dry-run", help="Validate without adding the slide."
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to write unless the attached document matches this name.",
    ),
) -> None:
    """Add a slide to the presentation without saving."""
    params: dict[str, Any] = {"dry_run": dry_run}
    if index is not None:
        params["index"] = index
    if layout is not None:
        params["layout"] = layout
    if expected_document:
        params["expected_document"] = expected_document

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            position = f"position {index}" if index is not None else "the end"
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"add a slide at {position}",
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            params["confirm"] = True
        return client.invoke("powerpoint.slide.add", params)

    _run(ctx, action)


@app.command("set-property")
def set_property(
    ctx: typer.Context,
    name: str = typer.Argument(..., help="Allowed property name, for example excel.active_sheet."),
    value: str = typer.Option(
        ...,
        "--value",
        help="JSON value or plain string, or '@path' to read the value from a UTF-8 file.",
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to change the property unless the attached document matches.",
    ),
) -> None:
    """Change an allowed document property."""
    parsed_value = _resolve_value_argument(value)

    def action(client: BridgeClient) -> dict[str, Any]:
        if not yes:
            _reject_unsafe_console()
            if not _confirm_write(
                ctx,
                client,
                f"set {name} to {parsed_value!r}",
                expected_document,
            ):
                typer.echo("Aborted.", err=True)
                raise typer.Exit(1)
        return client.set_property(name, parsed_value, True, expected_document)

    _run(ctx, action)


def _com_root_option():
    return typer.Option(
        "window",
        "--root",
        help="Root object: application, window, or document.",
    )


@com_app.command("get")
def com_get(
    ctx: typer.Context,
    expression: str = typer.Argument(
        ...,
        help='Property to read, for example \'ActiveSheet.Range("A1").Value2\'. Must end with a property name.',
    ),
    root: str = _com_root_option(),
) -> None:
    """Read any property from the attached Office document."""
    _run(
        ctx,
        lambda client: client.com("get", root, expression, None, False, False, None),
    )


@com_app.command("call")
def com_call(
    ctx: typer.Context,
    expression: str = typer.Argument(
        ...,
        help='Method call, for example \'ActiveSheet.Range("2:4").Insert(-4107)\'. Must end with (arguments).',
    ),
    root: str = _com_root_option(),
    dry_run: bool = typer.Option(
        False, "--dry-run", help="Validate the object and method without calling it."
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to call unless the attached document matches this name.",
    ),
) -> None:
    """Call any method on the Office object model."""

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx, client, f"com call: {expression}", expected_document
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            confirm = True
        else:
            confirm = False
        return client.com("call", root, expression, None, dry_run, confirm, expected_document)

    _run(ctx, action)


@com_app.command("set")
def com_set(
    ctx: typer.Context,
    expression: str = typer.Argument(
        ...,
        help='Property to set, for example \'Range("A1").Value2\'. Must end with a property name (no parentheses).',
    ),
    value: str = typer.Option(
        ...,
        "--value",
        help="JSON value or plain string, for example 42, true, \"text\", [[1,2],[3,4]].",
    ),
    root: str = _com_root_option(),
    dry_run: bool = typer.Option(
        False, "--dry-run", help="Validate the property without changing it."
    ),
    yes: bool = typer.Option(
        False, "--yes", help="Skip the interactive confirmation prompt."
    ),
    expected_document: str = typer.Option(
        None,
        "--expected-document",
        help="Refuse to set unless the attached document matches this name.",
    ),
) -> None:
    """Set any property on the Office object model."""
    parsed_value = _resolve_value_argument(value)
    if parsed_value is None:
        raise typer.BadParameter(
            "Setting a property to null is not supported; pass an empty string instead."
        )

    def action(client: BridgeClient) -> dict[str, Any]:
        if not dry_run:
            if not yes:
                _reject_unsafe_console()
                if not _confirm_write(
                    ctx,
                    client,
                    f"com set: {expression} = {parsed_value!r}",
                    expected_document,
                ):
                    typer.echo("Aborted.", err=True)
                    raise typer.Exit(1)
            confirm = True
        else:
            confirm = False
        return client.com("set", root, expression, parsed_value, dry_run, confirm, expected_document)

    _run(ctx, action)


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
        except Exception:
            pass
    app()


if __name__ == "__main__":
    main()
