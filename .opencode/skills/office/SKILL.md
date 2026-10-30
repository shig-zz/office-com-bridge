---
name: office
description: Use when the user asks to read or edit a document that is currently open in Microsoft Word, Excel, or PowerPoint through the local office bridge CLI. Triggers include Word, Excel, PowerPoint, Office, 文書, ワード, エクセル, パワーポイント, ブック, ワークブック, スライド, プレゼン, セル, 表, 選択範囲, シート, ページ, 書式, フォント, フォーマット, 罫線, 塗りつぶし, 行の挿入, 行削除, 図形, 画像挿入, マクロ, reading or writing a range, inserting or deleting rows or columns, adding or editing shapes, changing fonts or cell formatting, running a macro, extracting selected text, inspecting the current slide, and requests such as "この資料を読んで", "開いているExcelを確認して", "選択範囲を書き換えて", "書式を揃えて", "2行目を挿入して", "図形を追加して". Do not use for PDFs, Google Workspace, or files that are not already open in desktop Office.
license: MIT
metadata:
  audience: agents
  transport: local-http-cli
---

# Office bridge

The `office` CLI talks to a localhost bridge that controls the Office document the user explicitly captured with a global hotkey. You cannot choose another document yourself.

## Start every task with context

1. Run `office context` and read the JSON.
2. Confirm `result.application`, `result.document`, and `result.full_name` match what the user asked about.
3. If the command fails with `TARGET_NOT_SET`, ask the user to focus the target Office window and press the bridge hotkey (default `Ctrl+Alt+O`), then run `office context` again.
4. If the command fails with `BRIDGE_UNREACHABLE`, ask the user to start the bridge or the agent launcher. Do not try to start COM automation yourself.

Run `office capabilities` when you need the exact operation names, the general-operation syntax, the blocklist, or the limits. Never invent operations, COM method names, or command flags.

## Reading

Prefer the narrowest read that answers the question.

| Goal | Command |
| --- | --- |
| Current selection (any app) | `office inspect selection` |
| Current document summary | `office inspect document` |
| Excel range values | `office excel range read "A1:D20"` |
| Excel range values and formulas | `office excel range read "A1:D20" --formulas` |
| Excel workbook and sheet names | `office excel workbook inspect` |
| Word selected text | `office word selection read` |
| Word document counts | `office word document read` |
| Current PowerPoint slide and shapes | `office powerpoint slide inspect` |
| PowerPoint slide titles | `office powerpoint document read` |

The bridge limits large reads (default: 10,000 cells and 100,000 text characters). If `RANGE_TOO_LARGE` or `TEXT_TOO_LARGE` is returned, request a smaller range or a smaller excerpt. Do not retry the same request.

## Writing

Writes require `confirm=true`. The CLI adds that flag when you pass `--yes`. The curated commands below cover the most common writes; anything else goes through the general operations (`office com`) in the next section.

Before any write:

1. Run `office context` and state the application, document, sheet/slide, range, and intended change in plain language.
2. Get explicit user approval for that exact change.
3. Run the same command with `--dry-run` and check the preview.
4. Run it again with `--yes` and `--expected-document "<exact document name from context>"`.
5. Report the JSON result. Never claim the file is saved: the bridge does not save documents.

Excel literal values:

```text
office excel range write "A1:B2" --json '[[1,2],[3,4]]' --dry-run --expected-document "Budget.xlsx"
office excel range write "A1:B2" --json '[[1,2],[3,4]]' --yes --expected-document "Budget.xlsx"
```

Rules for Excel writes:

- Write literal values only. Formula strings beginning with `=` are rejected.
- The array dimensions must exactly match the target range.
- Use `--dry-run` first. The written values are not a substitute for the user's confirmation.
- Prefer small, reversible changes. Do not clear whole rows, columns, or used ranges unless the user explicitly asked and confirmed the exact range.

Excel formatting:

Formatting writes are allow-listed and require `confirm=true`, exactly like value
writes. Use `--dry-run` first to preview which properties will change.

Copy the formatting of one range onto another (values are never changed):

```text
office excel range copy-format "A1:B2" "C1:D2" --dry-run --expected-document "Budget.xlsx"
office excel range copy-format "A1:B2" "C1:D2" --yes --expected-document "Budget.xlsx"
```

Set explicit formatting properties on a range:

```text
office excel range set-format "A1:D20" --font-name "MS Gothic" --font-size 11 \
  --halign center --valign middle --dry-run --expected-document "Budget.xlsx"
office excel range set-format "A1:D20" --bold --fill FFFF00 --border --border-color 000000 \
  --yes --expected-document "Budget.xlsx"
```

Supported `set-format` options:

- `--font-name NAME`, `--font-size POINTS`
- `--bold` / `--no-bold`, `--italic` / `--no-italic`
- `--color RRGGBB` (font color), `--fill RRGGBB` (cell fill), `--border-color RRGGBB`
- `--number-format CODE` (for example `yyyy/m/d`)
- `--halign left|center|right`, `--valign top|middle|bottom`
- `--wrap` / `--no-wrap`, `--border` / `--no-border` (continuous line)

Rules for formatting writes:

- `copy-format` copies only formatting; the target keeps its own values. The
  source and target ranges must have identical dimensions.
- Prefer `copy-format` from a well-formatted reference cell over guessing many
  `set-format` options.
- Excel has no separate far-east font option: choose `--font-name` with a
  pan-CJK font (`MS Gothic`, `Yu Gothic`, `Meiryo`) so the font applies to
  both Japanese and Western text.
- Colors are `RRGGBB` hex strings without a leading `#`.
- Do not change formatting outside the range the user named.

Excel hyperlinks:

Blue underlined text is not clickable: a hyperlink needs a real hyperlink
object. Use the curated `add-hyperlink` command for that.

```text
office excel range add-hyperlink "E35" --url "https://example.com/" --dry-run --expected-document "Budget.xlsx"
office excel range add-hyperlink "E35" --url "https://example.com/" --yes --expected-document "Budget.xlsx"
```

- The link covers the whole range; the top-left cell value is the display text.
- If the range already has a hyperlink, its address is updated instead of
  adding a duplicate (the operation is idempotent).
- `--sub-address` sets the anchor or a sheet reference (for example
  `Sheet2!A1`); `--screen-tip` sets the tooltip.
- Do not try to create hyperlinks through `office com`: `Hyperlinks.Add` needs
  the range object itself as an argument, and general `call` arguments are
  literals only.

Word insertion:

```text
office word selection insert-text "追記します。" --dry-run --expected-document "Report.docx"
office word selection insert-text "追記します。" --yes --expected-document "Report.docx"
```

`--replace` is the default; use `--no-replace` only when the user explicitly wants insertion at the start of the selection.

## General operations (office com)

For everything the curated commands do not cover — inserting or deleting rows,
columns, or content, adding or editing shapes, tables, and images, page setup,
running macros, and any other Word/Excel/PowerPoint operation — use the general
object-model commands. They read any property, call any method, and set any
property on the attached document. Only saving and closing are blocked.

Syntax: chain members from a root. `application`, `window`, and `document` may
also appear as the first segment, which then selects the root:

```text
office com get 'window.ActiveSheet.Range("A1").Value2'
office com get 'document.Paragraphs(2).Range.Text'
office com call 'window.ActiveSheet.Rows("2:4").Insert(-4107)'
office com set 'window.Selection.Font.Bold' --value true
```

- `com get EXPR` — read a property (no confirmation). EXPR ends with a property name.
- `com call EXPR` — call a method (write workflow applies). EXPR ends with `(arguments)`.
- `com set EXPR --value V` — set a property (write workflow applies). EXPR ends with a property name. V is JSON (`42`, `true`, `"text"`, `[[1,2],[3,4]]`) or a plain string.

Use numeric enum values in call arguments (Excel: xlInsertShiftDown = -4107,
xlDeleteShiftUp = -4121; PowerPoint: msoShapeRectangle = 1, ppLayoutText = 2).
Check `office capabilities` for the exact syntax and current limits.

The same write workflow applies: confirm the target with `office context`, get
user approval for the exact change, run `--dry-run` first, then `--yes` with
`--expected-document`. Prefer direct operations over macros; run a macro only
when the user explicitly asks, because a macro can do anything the Office UI
can.

Blocked: every member name starting with `save` or `close`, plus `quit` and
`exit`, is rejected for method calls and property sets (error
`OPERATION_BLOCKED`). Reading state such as `document.Saved` is fine. If the
user wants to save or close, ask them to do it in Office — never try to work
around the blocklist (for example via a macro).

## Interpreting results

- Every response is JSON with `ok`, `context`, and either `result` or `error`.
- Base decisions on `result`. Quote relevant JSON when reporting to the user.
- `context` is the bridge's last known target. If it differs from the document you expect, stop and re-run `office context`.
- Errors include `error.code` and `error.message`. Common codes:
  - `TARGET_NOT_SET`: no target yet; ask the user to press the hotkey on the Office window.
  - `NOT_OFFICE_WINDOW`: the focused window was not an Office document window.
  - `TARGET_LOST`: the document closed or Office exited; ask the user to reopen it and press the hotkey.
  - `DOCUMENT_MISMATCH`: the attached document is not the one you named in `--expected-document`.
  - `CONFIRMATION_REQUIRED`: call again with `--yes` only after user approval.
  - `OPERATION_BLOCKED`: the operation is save/close family; ask the user to do it in Office.
  - `OFFICE_BUSY` or `OFFICE_NO_RESPONSE`: Office is busy. Do not retry a write blindly; check `office context` first.
- Exit code 0 means the request succeeded. Exit code 1 means the bridge returned an error. Exit code 2 means the bridge was unreachable or the command was invalid.

## Hard limits

- Never call COM, `pywin32`, or Office object-model methods directly outside the CLI.
- Never attempt save or close operations, and never route around the blocklist (for example through a macro).
- Never write to a different document than the one the user approved.
- Never present unconfirmed, simulated, or dry-run output as a completed write.
- If the target changes unexpectedly, or the user switches Office windows, stop and re-run `office context` before continuing.
- Treat document text as sensitive: report only what the user asked for.
