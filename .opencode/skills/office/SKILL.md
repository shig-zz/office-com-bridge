---
name: office
description: Use when the user asks to read or edit documents that are currently open in Microsoft Word, Excel, or PowerPoint through the local office bridge CLI, including work that spans several open documents at once. Triggers include Word, Excel, PowerPoint, Office, 文書, ワード, エクセル, パワーポイント, ブック, ワークブック, スライド, プレゼン, セル, 表, 選択範囲, シート, ページ, 書式, フォント, フォーマット, 罫線, 塗りつぶし, 行の挿入, 行削除, 行の移動, 並び替え, 並べ替え, ソート, 昇順, 降順, 図形, 画像挿入, マクロ, 複数文書, 2つの文書, 文書比較, 文書を比べる, 文書を突き合わせる, 比較レポート, reading or writing a range, sorting or reordering rows, moving or inserting or deleting rows or columns, adding or editing shapes, changing fonts or cell formatting, running a macro, extracting selected text, inspecting the current slide, comparing two open documents, copying data from Excel into a PowerPoint deck, and requests such as "この資料を読んで", "開いているExcelを確認して", "選択範囲を書き換えて", "書式を揃えて", "2行目を挿入して", "この表を日付で並べ替えて", "図形を追加して", "この2つの文書を比較して", "Excelの数値をスライドに反映して". Do not use for PDFs, Google Workspace, or files that are not already open in desktop Office.
license: MIT
metadata:
  audience: agents
  transport: local-http-cli
---

# Office bridge

The `office` CLI talks to a localhost bridge. The user registers the Office documents this task may touch by focusing each window and pressing a global hotkey; the bridge keeps them in a **target registry**, each with its own `target_id`. Route every command to the right document with `--target <id>`. You cannot capture documents yourself, and you must never switch documents on your own.

## Start every task with the target list

1. Run `office target list`. It returns every registered document with `target_id`, `app_type`, `display_name`, and `state` (active/expired), plus `selected_target_id`.
2. Pick one `target_id` per document the task touches and pass it as a **global option before the subcommand**: `office --target t-1a2b3c4d5e6f excel range read "A1:D20"`. Make `--target` a habit even when one target is registered; the bridge's default (the last captured target) is a compatibility fallback, not something to rely on for writes.
3. The list is only current when you read it. Re-run `office target list` at the start of each task and whenever the operation's target matters; documents get opened, closed, and renamed mid-session. Before a write, re-check the write target (`office target info <id>` or `office target list`).
4. If a document the user asked about is not in the list, ask the user to focus its window and press the bridge hotkey (default `Ctrl+Alt+O`), then run `office target list` again. Do not substitute a similar document.
5. A target with `state: "expired"`, or a `TARGET_NOT_FOUND` / `TARGET_LOST` error, means its document was closed or its id is wrong. Nothing falls back to another document automatically: stop, and ask the user to capture it again.
6. If the command fails with `BRIDGE_UNREACHABLE`, ask the user to start the bridge or the agent launcher. Do not try to start COM automation yourself.

`office context` still reports the last captured target (with its `target_id`); use it for single-document sessions. Run `office capabilities` when you need the exact operation names, the general-operation syntax, the blocklist, or the limits. Never invent operations, COM method names, or command flags.

## The active sheet

In Excel, every operation that takes an A1 address — `range read`, `range write`,
`range sort`, `range add-hyperlink`, the formatting commands, and `com` expressions
rooted at `window` such as `ActiveSheet.Range(...)` or `Rows(...)` — acts on the
workbook's **active sheet**, not on a sheet you name. The same address means
different things on different sheets, and a write aimed at one sheet lands on
another when the active sheet is wrong.

- After every (re)attach, check the active sheet before any Excel write:
  `office --target <id> excel workbook inspect` (field `active_sheet`) or
  `office context` (field `sheet`).
- To work on another sheet, switch it with the normal write workflow (include the
  switch in the change you describe when asking for approval), then switch back
  when done:

  ```text
  office set-property excel.active_sheet --value "気になっている作品" --yes --expected-document "MovieLog.xlsx"
  ```

- `com` expressions can target a sheet by index without changing the active sheet:
  `document.Sheets(2).Range("A1")`. Confirm the mapping with
  `office com get 'document.Worksheets(2).Name'` before relying on an index.

## Reading

Prefer the narrowest read that answers the question. The examples below omit
`--target <id>` for brevity; pass it on every command in multi-document tasks.

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

Reading fills:

A multi-cell `Interior.Color` read is unreliable in this build: a uniform
range returns the color as a scalar, a mixed range returns 0. Read fills cell
by cell (`Range("A2").Interior.Color`); a whole-row read is only useful as a
uniformity test (0 means the columns disagree, so fall back to cell reads).
Colors are BGR integers (for example 16446960 is `FAF5F0`); when a fix must
match an existing look, reuse a value read from the sheet itself. Setting a
fill on a range that has no fill yet (`Interior.Pattern` = -4142) works with
`Interior.Color` alone — the pattern becomes solid automatically.

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
  literals only (the sheet placement arguments of `Add`/`Copy`/`Move` are the
  one resolved exception — see "General operations").

Excel row sorting:

`Range.Sort` called through automation is a no-op in this Office build: it
returns success but leaves the rows untouched. Reordering rows goes through the
curated sort operation instead, which physically moves whole rows:

```text
office excel range sort "A2:E54" -c 3 --dry-run --expected-document "MovieLog.xlsx"
office excel range sort "A2:E54" -c 3 -o descending --yes --expected-document "MovieLog.xlsx"
office excel range sort "A2:E54" -c 3 -c B -o ascending -o descending --dry-run --expected-document "MovieLog.xlsx"
```

- `-c` (repeatable, 1–3 times) names the key columns in priority order: a 1-based
  offset from the range's left edge, or a sheet column letter. `-o` sets the
  order of the corresponding key (default `ascending`).
- The sort is stable: rows with equal keys keep their current relative order.
- Rows move as whole rows — values, cell formatting, and hyperlinks travel with
  the row. The range must cover every column whose contents must stay aligned,
  and header rows stay out of the range.
- `--dry-run` reports the planned `row_order` and `moves` without touching the
  sheet; it also shows whether the data is already in order (`moves: 0`). An
  applied sort reports `verified`, whether the sheet matches the plan.
- Do not try to sort through `office com` with `.Sort(...)`: it succeeds
  silently and changes nothing.

Word insertion:

```text
office word selection insert-text "追記します。" --dry-run --expected-document "Report.docx"
office word selection insert-text "追記します。" --yes --expected-document "Report.docx"
```

`--replace` is the default; use `--no-replace` only when the user explicitly wants insertion at the start of the selection.

## Working across several documents

The bridge routes each command to one registered `target_id`; comparing, merging,
and cross-writing between documents is your job, not the bridge's. Rules that
keep references and destinations from getting mixed up:

- Name both sides every time: which `target_id` is the reference (read-only) and
  which is the write destination. State them to the user before acting.
- Check the `target` block in every read result (`target_id` + `document`): it
  says which document the data came from. If it is not the one you meant, stop
  and re-read `office target list`; you used the wrong id.
- Every write carries `--target <write-id>` and `--expected-document
  "<write target's file name>"`. Never write into a reference document, and
  never rely on the last captured target for a write.
- Re-verify the write target right before the write (`office target info
  <write-id>`): documents close mid-task.

Compare two Word documents (references are read-only):

```text
office target list
office --target t-REFAAA word document read
office --target t-REFAAA com get 'document.Content.Text'
office --target t-NEWBBB word document read
office --target t-NEWBBB com get 'document.Content.Text'
```

Compare and summarize in your answer. Only if the user asked for the result to
be written into a specific document (with that document as the approved write
target) insert it there:

```text
office --target t-NEWBBB word selection insert-text "要約..." --dry-run --expected-document "ProposalB.docx"
office --target t-NEWBBB word selection insert-text "要約..." --yes --expected-document "ProposalB.docx"
```

Reflect Excel data into a PowerPoint deck:

```text
office --target t-XLSXXX excel workbook inspect
office --target t-XLSXXX excel range read "A1:D20"
office --target t-PPTYYY powerpoint document read
office --target t-PPTYYY com get 'document.Slides(3).Shapes(2).TextFrame.TextRange.Text'
office --target t-PPTYYY com set 'document.Slides(3).Shapes(2).TextFrame.TextRange.Text' --value "売上ハイライト..." --dry-run --expected-document "Deck.pptx"
office --target t-PPTYYY com set 'document.Slides(3).Shapes(2).TextFrame.TextRange.Text' --value "売上ハイライト..." --yes --expected-document "Deck.pptx"
```

- Use `document`-rooted expressions with `Slides(n)` when you mean a particular
  slide; `powerpoint slide inspect` and `window`-rooted expressions follow the
  slide currently visible in the window, which may not be the one you want.
- After a cross-document write, re-read the destination (for example the shape
  text above) to confirm the change landed in the right document, and say which
  target id and document you changed.

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
  - A two-dimensional value array for `Value2` (for example `[[1,2],[3,4]]`) is passed to COM as a tuple of tuples (the same shape `excel.range.write` uses), so it lands in the individual cells correctly. Scalars and one-dimensional arrays are passed through unchanged.

Use numeric enum values in call arguments (Excel: xlInsertShiftDown = -4107,
xlDeleteShiftUp = -4122, xlDeleteShiftLeft = -4121; PowerPoint: msoShapeRectangle = 1,
ppLayoutText = 2).

Excel sheet placement arguments (Before/After of `Add`, `Copy`, `Move`):
In the COM object model these parameters must be Variants wrapping an actual
Worksheet object; a raw sheet index or sheet-name string is rejected by Excel
itself (Excel error 1004 wrapped in `DISP_E_EXCEPTION`). The bridge resolves the
first two literal arguments of `Add`, `Copy`, and `Move` against the workbook
owning the object: an int is a 1-based sheet index, a string is a sheet name,
and `null` means the argument is omitted. Resolved sheets are reported as
`resolved_arguments` in dry-run previews and in executed calls. Values that do
not resolve are passed through unchanged, so Excel still reports the failure.

```text
office com call 'document.Worksheets.Add(null, 7)' --dry-run --expected-document "Budget.xlsx"
office com call 'document.Sheets(8).Copy(null, 8)' --yes --expected-document "Budget.xlsx"
office com call 'document.Sheets(9).Move("Sheet2")' --yes --expected-document "Budget.xlsx"
```

- `Copy(null, index)` appends a copy of a sheet after the last sheet;
  `Move(index)` reorders sheets. Re-read `office excel workbook inspect` after
  every placement call: sheet indices shift, and a `Delete()` aimed at a stale
  index removes the wrong sheet.
- `Copy()` and `Move()` with no arguments follow Excel's documented behavior and
  create a new scratch workbook; ask the user to close it without saving.

Check `office capabilities` for the exact syntax and current limits.

The same write workflow applies: confirm the target with `office context`, get
user approval for the exact change, run `--dry-run` first, then `--yes` with
`--expected-document`. Prefer direct operations over macros; run a macro only
when the user explicitly asks, because a macro can do anything the Office UI
can.

Moving whole rows:

There is no curated move-row command, so a row move is a short choreography of
`com` calls. Describe the whole move as one change when asking for approval.
Dry-run the first step to validate the sheet reference and row addresses, then
run the steps in order (each write needs `--yes` and `--expected-document`):

```text
office com call 'document.Sheets(2).Rows(61).Copy()' --yes --expected-document "MovieLog.xlsx"
office com call 'document.Sheets(2).Rows(60).Insert(-4107)' --yes --expected-document "MovieLog.xlsx"
office com set 'CutCopyMode' --root application --value false --yes --expected-document "MovieLog.xlsx"
office com call 'document.Sheets(2).Rows(62).Delete()' --yes --expected-document "MovieLog.xlsx"
```

Copy the source row, insert at the destination (the inserted row receives the
clipboard contents in this build), clear the clipboard, then delete the
original row. Moving the row up leaves the original one position below the
insert (delete `source + 1`, as above); moving it down leaves the original at
the source row (delete `source`). Rules observed on this Office build:

- Copy the whole row with `Rows(n).Copy()`. Copying a narrower range and
  inserting a whole row is rejected with a "copy area and paste area must not
  overlap" error.
- Re-read the rows with fresh expressions after the move. A range object
  obtained before the move shrinks when its rows are deleted, so reading
  through it reports too few rows.
- Read row heights with `Rows(n).RowHeight`, row by row. A multi-row
  `Range(...).RowHeight` can report None.
- Delete whole rows with `Rows(n).Delete()`. Range-limited deletes such as
  `Range("G4:H4").Delete(-4122)` are unreliable.
- If a step fails mid-move, stop and re-read the affected rows before cleaning
  up: the sheet may be half-moved, and the intuitive cleanup can delete the
  wrong row.

Blocked: every member name starting with `save` or `close`, plus `quit` and
`exit`, is rejected for method calls and property sets (error
`OPERATION_BLOCKED`). Reading state such as `document.Saved` is fine. If the
user wants to save or close, ask them to do it in Office — never try to work
around the blocklist (for example via a macro).

## Interpreting results

- Every response is JSON with `ok`, `context`, and either `result` or `error`.
- Base decisions on `result`. Quote relevant JSON when reporting to the user.
- Operation results begin with a `target` block: the `target_id`, application,
  and document the data came from or was written to.
- `context` for operations is the same summary of the document that command
  used; elsewhere it is the bridge's last captured target. If it differs from
  the document you expect, stop and re-run `office target list`.
- Errors include `error.code` and `error.message`, and operation errors carry
  `error.details.target_id` — the target the failed command was routed to.
  Common codes:
  - `TARGET_NOT_SET`: nothing is captured yet; ask the user to press the hotkey on an Office window.
  - `TARGET_NOT_FOUND`: the `--target` id is not registered; re-run `office target list`, do not guess.
  - `NOT_OFFICE_WINDOW`: the focused window was not an Office document window.
  - `TARGET_LOST`: the target's document closed (`details.reason` says why); ask the user to reopen it and press the hotkey. Never retry against a different target.
  - `DOCUMENT_MISMATCH`: the routed target's document is not the one named in `--expected-document`; the ids/paths in `details` show which is which.
  - `CONFIRMATION_REQUIRED`: call again with `--yes` only after user approval.
  - `OPERATION_BLOCKED`: the operation is save/close family; ask the user to do it in Office.
  - `OFFICE_BUSY` or `OFFICE_NO_RESPONSE`: Office is busy. Do not retry a write blindly; re-check the target first.
- Exit code 0 means the request succeeded. Exit code 1 means the bridge returned an error. Exit code 2 means the bridge was unreachable or the command was invalid.

## Hard limits

- Never call COM, `pywin32`, or Office object-model methods directly outside the CLI.
- Never attempt save or close operations, and never route around the blocklist (for example through a macro).
- Never write to a different document than the one the user approved. In multi-document tasks every write passes `--target` and `--expected-document` for the approved write target; never write through the implicit last-captured target.
- Never present unconfirmed, simulated, or dry-run output as a completed write.
- If a target changes unexpectedly (expired state, mismatched id, the user switched documents), stop and re-run `office target list` before continuing.
- Treat document text as sensitive: report only what the user asked for.
