# Office COM Bridge

Local HTTP and CLI bridge between AI agents and running Microsoft Office
desktop applications. Capture live Word, Excel, or PowerPoint windows with a
hotkey; each becomes a **registered target** with its own `target_id`, and
every command routes to the target you name (or the last captured one when you
do not). Read and write the current selection or document. Curated commands
cover the common cases; a general object-model layer (`POST /com`, `office com
...`) exposes every other Word/Excel/PowerPoint operation.

The bridge never launches Office and never saves or closes documents: member
names starting with `save` or `close` (plus `quit`/`exit`) are rejected for
method calls and property sets. Everything runs on loopback and all COM access
goes through a single STA worker thread.

## Requirements

- Windows 10 or 11
- Microsoft Office desktop (Word, Excel, PowerPoint)
- Python 3.11 or 3.12
- An agent CLI for the launcher: OpenCode, Claude Code, Codex, or Hermes Agent
  (or a custom command)

## Install

```powershell
python -m venv --system-site-packages .venv
.\.venv\Scripts\python -m pip install -e ".[dev]"
```

The `--system-site-packages` flag is useful when `pywin32` is already installed
globally for the same Python version.

### Batch shortcuts (Windows)

- `install.bat` creates `.venv` with a supported Python (3.11 or 3.12), installs
  the package and dev dependencies, and copies the Office skill for the
  configured agent (OpenCode by default).
- `configure.bat` selects OpenCode, Claude Code, Codex, Hermes Agent, or a
  custom command, and stores the choice in
  `%LOCALAPPDATA%\office-bridge\agent.json`. The interactive flow can also
  install the skill immediately.
- `start.bat` starts the bridge in the current console. Server options are passed
  through, for example `start.bat --no-hotkey --port 8733`.
- `start-agent.bat` starts the bridge if needed and launches the configured
  agent with the Office skill available in the current terminal. Use
  `start-agent.bat -Agent hermes` for a one-off override.

## Start the bridge

```powershell
.\.venv\Scripts\office-bridge
```

Defaults:

- Address: `127.0.0.1:8733`
- Hotkey: `Ctrl+Alt+O`
- Logs: `%LOCALAPPDATA%\office-bridge\logs`

Open a document in Word, Excel, or PowerPoint, focus its window, and press the
hotkey. That document is **captured into the target registry** (existing
targets are kept), and the configured agent starts in a new Windows Terminal
window with the Office skill available, so it can immediately read or edit that
document. Capture each document a task touches, then route commands with
`--target <id>`. Use `--no-hotkey` for health checks or tests that do not need
target capture:

```powershell
.\.venv\Scripts\office-bridge --no-hotkey --port 8733
```

To disable the automatic agent launch on hotkey while keeping target capture,
set `OFFICE_BRIDGE_LAUNCH_AGENT=0` in the environment before starting the
bridge.

## CLI

Activate the virtual environment first, or call `.\.venv\Scripts\office.exe`
directly:

```powershell
.\.venv\Scripts\Activate.ps1
office health
office target list
office target info t-1a2b3c4d5e6f
office context
office capabilities
office inspect excel.selection
office --target t-1a2b3c4d5e6f excel range read A1:B2 --formulas
office --target t-1a2b3c4d5e6f excel range write A1:B2 --json "[[1,2],[3,4]]" --dry-run
office --target t-1a2b3c4d5e6f excel range write A1:B2 --json "[[1,2],[3,4]]" --yes
office excel range copy-format A1:B2 C1:D2 --yes
office excel range set-format A1:B2 --font-name "MS Gothic" --font-size 11 --halign center --yes
office excel range add-hyperlink E35 --url "https://example.com/" --yes
office word selection read
office word selection insert-text "Replacement" --yes
office powerpoint slide inspect
office com get 'window.ActiveSheet.Range("A1").Value2'
office com call 'window.ActiveSheet.Rows("2:4").Insert(-4107)' --yes
office com set 'window.Selection.Font.Bold' --value true --dry-run
```

`--target <target_id>` (a global option, before the subcommand) routes the
command to a captured target. Without it the bridge falls back to the last
captured target; that is a compatibility default, so pass `--target` on every
command in multi-document work. `office target list` shows every registered
target with its `target_id` and `state`; `office target remove <id>` unregisters
one.

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
| `GET /context` | Current (default) target details |
| `GET /capabilities` | Supported inspections, operations, properties, and general-operation info |
| `GET /targets` | Registered targets with `target_id` and `state` |
| `POST /targets/capture` | Capture a target by `hwnd` (tests; normally the hotkey captures) |
| `GET /targets/{target_id}` | One target's details with a full liveness check |
| `DELETE /targets/{target_id}` | Unregister a target |
| `POST /inspect` | Read a document/selection view |
| `POST /invoke` | Run a curated write operation |
| `POST /set-property` | Change a curated property |
| `POST /com` | General object-model access: `get`, `call`, or `set` on the attached document |

Operation requests accept an optional top-level `target_id`; when omitted they
use the last captured target. Operation results start with a `target` block and
operation errors carry `error.details.target_id`. Unknown ids fail with
`TARGET_NOT_FOUND`; closed documents with `TARGET_LOST` (no fallback to another
document).

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
- `excel.range.add_hyperlink` adds a real hyperlink object to a range (blue
  underlined text alone is not clickable). The link covers the whole range and
  the top-left cell value is used as display text. If the range already has a
  hyperlink, its address is updated instead, so the operation is idempotent.
  Optional `sub_address` (anchor or sheet reference) and `screen_tip` (tooltip)
  are supported; formula-style strings beginning with `=` are rejected
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

## Agent selection

Run `configure.bat` without arguments for an interactive choice, or use the
`office-agent` CLI directly:

```powershell
.\.venv\Scripts\office-agent list
.\.venv\Scripts\office-agent set hermes
.\.venv\Scripts\office-agent show
.\.venv\Scripts\office-agent install-skill
```

The same operations are available without the console script:

```powershell
.\.venv\Scripts\python -m office_bridge.agent_config set hermes
```

Presets and their default skill locations:

| Agent | Command | Project skill | User skill | Default scope |
| --- | --- | --- | --- | --- |
| `opencode` | `opencode` | `.opencode/skills/office` | `~/.config/opencode/skills/office` | project |
| `claude` | `claude` | `.claude/skills/office` | `~/.claude/skills/office` | project |
| `codex` | `codex` | `.agents/skills/office` | `~/.agents/skills/office` | project |
| `hermes` | `hermes -s office` | `.hermes/skills/office` | `%HERMES_HOME%\skills\office` | user |
| `custom` | `--command` | `--skill-dir` or disabled | `--skill-dir` | project |

Hermes defaults to the user skill directory because project-local skills require
`hermes skills trust`; it also preloads the skill with `-s office`. For a custom
agent, provide an executable and extra arguments explicitly:

```powershell
.\.venv\Scripts\office-agent set custom --command "C:\Tools\agent.exe" --arg=--flag --no-install-skill
```

Use `--scope user` to install a preset skill globally, or `--skill-dir <path>` to
choose the destination. `--no-preload` suppresses the Hermes `-s office`
argument. `--no-install-skill` leaves skill management to you.

The saved configuration is `%LOCALAPPDATA%\office-bridge\agent.json`. Override it
for one process with `OFFICE_BRIDGE_AGENT` (for example `OFFICE_BRIDGE_AGENT=codex`)
or point at a different file with `OFFICE_BRIDGE_AGENT_CONFIG`.

The launcher starts the bridge if needed, installs the skill for the selected
agent, and starts that agent in the current directory:

```powershell
.\scripts\start-agent.ps1
.\scripts\start-agent.ps1 -Agent hermes
```

The skill teaches the agent when and how to use the CLI.

## Tests

```powershell
.\.venv\Scripts\python -m pytest
```

## Safety

- Bind to loopback only.
- Requests are processed serially in one STA worker.
- The target registry keeps every captured document until it closes; each
  target has a stable `target_id` and is invalidated (never silently replaced)
  when the document or application closes.
- Reads are capped at 10,000 cells and 100,000 characters by default.
- Save and close operations are blocked for general operations; all other
  writes require `confirm` and support `dry_run`.
