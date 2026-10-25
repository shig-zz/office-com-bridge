# Office COM Bridge

Local HTTP and CLI bridge between AI agents and running Microsoft Office
desktop applications. Select a live Word, Excel, or PowerPoint window, then
read and write the current selection or document. Curated commands cover the
common cases; a general object-model layer (`POST /com`, `office com ...`)
exposes every other Word/Excel/PowerPoint operation.

The bridge never launches Office and never saves or closes documents: member
names starting with `save` or `close` (plus `quit`/`exit`) are rejected for
method calls and property sets. Everything runs on loopback and all COM access
goes through a single STA worker thread.

## Requirements

- Windows 10 or 11
- Microsoft Office desktop (Word, Excel, PowerPoint)
- Python 3.11 or 3.12
- An OpenCode installation for the agent launcher

## Install

```powershell
python -m venv --system-site-packages .venv
.\.venv\Scripts\python -m pip install -e ".[dev]"
```

The `--system-site-packages` flag is useful when `pywin32` is already installed
globally for the same Python version.

## Start the bridge

```powershell
.\.venv\Scripts\office-bridge
```

Defaults:

- Address: `127.0.0.1:8733`
- Hotkey: `Ctrl+Alt+O`
- Logs: `%LOCALAPPDATA%\office-bridge\logs`

Open a document in Word, Excel, or PowerPoint, focus its window, and press the
hotkey. The selected document becomes the bridge target, and OpenCode starts in
a new Windows Terminal window with the Office skill available, so the agent can
immediately read or edit that document. Use `--no-hotkey` for health checks or
tests that do not need target selection:

```powershell
.\.venv\Scripts\office-bridge --no-hotkey --port 8733
```

To disable the automatic OpenCode launch on hotkey while keeping target capture,
set `OFFICE_BRIDGE_LAUNCH_AGENT=0` in the environment before starting the
bridge.

## CLI

Activate the virtual environment first, or call `.\.venv\Scripts\office.exe`
directly:

```powershell
.\.venv\Scripts\Activate.ps1
office health
office context
office capabilities
office inspect excel.selection
office excel range read A1:B2 --formulas
office excel range write A1:B2 --json "[[1,2],[3,4]]" --dry-run
office excel range write A1:B2 --json "[[1,2],[3,4]]" --yes
office excel range copy-format A1:B2 C1:D2 --yes
office excel range set-format A1:B2 --font-name "MS Gothic" --font-size 11 --halign center --yes
office word selection read
office word selection insert-text "Replacement" --yes
office powerpoint slide inspect
office com get 'window.ActiveSheet.Range("A1").Value2'
office com call 'window.ActiveSheet.Rows("2:4").Insert(-4107)' --yes
office com set 'window.Selection.Font.Bold' --value true --dry-run
```

Writes require confirmation. `--dry-run` validates without changing the
document; `--yes` skips the interactive prompt. Use `--expected-document` to
guard against acting on a different document.

`office com get|call|set` runs any property read, method call, or property
set on the attached document (see "General operations" below).

## HTTP API

All endpoints return an envelope:

```json
{
  "ok": true,
  "context": {"application": "Excel", "document": "Book1.xlsx"},
  "result": {}
}
```

| Endpoint | Purpose |
| --- | --- |
| `GET /health` | Liveness and current target summary |
| `GET /context` | Current target details |
| `GET /capabilities` | Supported inspections, operations, properties, and general-operation info |
| `POST /inspect` | Read a document/selection view |
| `POST /invoke` | Run a curated write operation |
| `POST /set-property` | Change a curated property |
| `POST /com` | General object-model access: `get`, `call`, or `set` on the attached document |

Errors use `{"ok": false, "context": ..., "error": {"code", "message", "details"}}`.

## Curated writes

- `excel.range.write` with literal values (formulas are rejected)
- `excel.range.copy_format` copies formatting only (font, fill, borders,
  number format, alignment) from a source range to a same-sized target range;
  values are never changed
- `excel.range.set_format` sets explicit formatting: font name (including
  far-east), font size, bold, italic, underline, font color, number format,
  horizontal/vertical alignment, wrap, fill color, continuous border and
  border color
- `word.selection.insert_text`
- `excel.active_sheet` through `/set-property`

Write operations support `dry_run` and, except for dry runs, require `confirm`.
Property changes also require `confirm`.

## General operations

`POST /com` (and `office com get|call|set`) exposes the full Word/Excel/
PowerPoint object model for operations the curated commands do not cover. A
request names a root (`application`, `window`, or `document`) and an
expression chaining members, for example:

```json
{"mode": "get", "root": "window", "expression": "ActiveSheet.Range(\"A1\").Value2"}
{"mode": "call", "expression": "window.ActiveSheet.Rows(\"2:4\").Insert(-4107)", "confirm": true}
{"mode": "set", "expression": "Selection.Font.Bold", "value": true, "confirm": true}
```

`get` reads a property, `call` invokes a method, and `set` assigns a property
from `value`. The leading root segment in the expression selects the root when
no conflicting `root` is given. `call` and `set` support `dry_run` and, except
for dry runs, require `confirm`.

Only saving and closing are intentionally unavailable: member names starting
with `save` or `close` (plus `quit` and `exit`) are rejected with
`OPERATION_BLOCKED` for calls and property sets. Property reads (such as
`document.Saved`) are always allowed. Add names to the blocklist with
`OFFICE_BRIDGE_BLOCKED_CALLS` (comma separated).

## OpenCode launcher

```powershell
.\scripts\start-agent.ps1
```

The launcher starts the bridge if needed, copies the Office skill to
`.opencode/skills/office/SKILL.md`, and starts OpenCode in the current directory.
The skill teaches the agent when and how to use the CLI.

## Tests

```powershell
.\.venv\Scripts\python -m pytest
```

## Safety

- Bind to loopback only.
- Requests are processed serially in one STA worker.
- The target is captured at selection time and invalidated when the document
 or application closes.
- Reads are capped at 10,000 cells and 100,000 characters by default.
- Save and close operations are blocked for general operations; all other
  writes require `confirm` and support `dry_run`.
