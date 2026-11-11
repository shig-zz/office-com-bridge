# Phase 3 計画：Excel 構造操作・Word 編集系・PowerPoint 補完（三アプリ横断）

作成日: 2026-11-11
**状態: 3A・3B・3C とも実装完了（2026-11-12）。pytest 369 件全緑、capabilities・CLI・SKILL.md・README 日英 更新済み。実文書スモークのみ未実施（ブリッジ再起動とユーザーのホットキー登録が必要）。**

起点: Phase 2 までの実運用（PowerPoint 校訂 33 書き込み、Excel 行移動・罫線・塗りつぶし作業、複数文書比較）と、SKILL.md に固定した「この Office ビルドの挙動」から、キュレーション欠如・手組み回避が残っている領域を抽出した。対象は Word / Excel / PowerPoint の三アプリ。

## 観測した問題

| # | 事象 | 根拠 | 影響 |
|---|---|---|---|
| 1 | 行の挿入・削除・移動が `com` 手組み | SKILL.md「Moving whole rows」節（Copy→Insert→CutCopyMode→Delete の choreography と 5 種の注意書き） | enum（-4107 等）の誤り、半移動状態、行番号の stale 参照など事故ポイントが多い。undo 導入後はラップ価値がさらに上がった |
| 2 | 書式の読み取りが unreliable | SKILL.md「Reading fills / borders」節：複数セルの `Interior.Color` は uniform/mixed で挙動が変わる、色は BGR 整数、`Borders(n)` の引数は位置 index | 書式監査・復元のための手作業セル読み＋BGR→RRGGBB 変換が毎回発生。`copy-format` は罫線を運ばないことも既知 |
| 3 | Word の編集経路が `selection.insert_text` だけ | registry の word 系は `word.selection` / `word.document`（read）と insert_text のみ | 置換・検索・段落単位編集が全て `com` 手組み（`Content.Text` の全文 dump からの人力 grep）。校訂系タスクの再発コストが高い |
| 4 | PowerPoint の並べ替え・図形削除・書式設定が手組み | `powerpoint.slide.move`（MoveTo）・`shape delete`・`Font.*` set が未配線 | 校訂ノート挿入済みの 74 枚 deck でも「スライド移動」「図形フォント-size 調整」は com 生呼びが必要 |
| 5 | 破壊的操作の結果に取り消し案内が無い | `undo`（Phase 2C）導入済みだが、`slide delete` 等の結果 payload に `undo_hint` が無い | 「ブリッジ経由では undo 不可」と SKILL.md に書いてある記述が古くなり、ユーザー側 undo を勧められていない |
| 6 | PowerPoint の `Application.Undoable` が読めず `undo --dry-run` が `undoable: null` | 2C 実文書スモーク（2026-11-11） | dry-run の事前報告がアプリにより機能しない。防御設計で crash はしていないので優先度は低 |

## Phase 3A：Excel 構造操作と書式キュレーション（最優先）

`operations/excel.py` に追加し、`registry.invoke` / `inspect` と CLI に配線する。すべて `expected_document`・`dry_run`・`confirm` の既存規約に従う。

### A1. write: `excel.rows.insert` / `excel.rows.delete` / `excel.columns.insert` / `excel.columns.delete`
- CLI: `office excel rows insert "2:4" [--shift down|up]`、`office excel rows delete 61 [--shift up|left]`（columns も同形）
- 内部で shift enum を解決して渡す（insert: xlShiftDown=-4107 / xlShiftToRight=-4161、delete: xlShiftUp=-4122 / xlShiftToLeft=-4121）。引数指定なしのときは既定（down / up）を使い、enum をユーザーに見せない。
- dry-run は対象行（`Rows(n).RowHeight` 等の読み取りで存在確認した旨）と shift を報告。実行後は used range 行数の差分で `verified` を返す（`range sort` と同じ検証方針）。
- SKILL.md の注意書きどおり、delete は範囲限定（`Range(...).Delete(...)`）を許さず行・列単位のみとする。

### A2. write: `excel.rows.move`
- CLI: `office excel row move 61 --to 60 [--sheet 気になっている作品]`
- 実測済みの choreography をラップする：`Rows(src).Copy()` → `Rows(dst).Insert(-4107)` → `CutCopyMode=false`（application ルート）→ 原行を `Delete()`（上方向移動は `src+1`、下方向は `src`）。`--sheet` 省略時は active sheet（既存キュレーションと同じ前提）。
- 実行前に `Rows` 数を記録し、実行後 (a) 行数が移動前後で不変、(b) 移動後行の A 列値が元の内容と一致、を検証して `verified` を返す。
- 途中失敗時は `partial: true` と完了 step 一覧を返し、以後の cleanup 判断をユーザーに委ねる（自動巻き戻しはしない。必要なら `office undo` を案内する `undo_hint` を付ける）。
- クリップボードを使うため、同一ブリッジで他のクリップボード操作と同時実行できない旨を capabilities に注記する。

### A3. read: `excel.range.format`（inspect）
- CLI: `office excel range read-format "A1:D20"`
- セル単位で読み、`{address, cells[]: {cell, fill: rrggbb|null, pattern, font: {name, size, bold, color}, number_format, borders: {left, top, bottom, right, inside_h, inside_v}}}` を返す。色は BGR→RRGGBB 変換済み。
- 複数セル一括読み（uniformity テスト専用）は内部実装に隠し、返答からは「どのセルが不統一か」だけを拾える形にする。罫線は SKILL.md 記載の位置 index（7–12）を名前付きフィールドへ変換する。

### A4. fix: `excel.range.copy_format` の罫線
- このビルドで罫線が伝わらない既定挙動を応答に明示する（`borders_applied: false`）。
- `--borders` 指定時だけ、コピー後に A3 の罫線読み取り→セル単位の `Borders(...)` 書き込みで補完する（`set-format --border` 相当ではなく、source の weight/color を再現）。

## Phase 3B：Word 編集系

`operations/word.py` に追加。Word は Find/Replace の automation 挙動が未検証なので、**まず読み取り（B1）と検索（B2）を先に実装し、書き込み（B3）は実文書検証後に有効化する**。

### B1. read: `word.paragraphs`（inspect）
- CLI: `office word paragraphs read [--start N] [--count M]`
- `{paragraph_index, text, style, is_heading, characters}` の配列。`--count` 既定 50・上限 200、`max_text_chars` で切り捨て。
- 構造把握を `com get 'document.Content.Text'` の dump から解放するのが目的。見出し（Heading 1–9 style）に `is_heading` を付ける。

### B2. read（準読み取り）: `word.find`
- CLI: `office word find "検索語" [--match-case] [--whole-words] [--max N]`
- 変更を伴わないので confirm 不要。`Find.Execute` を検索のみでループし、`matches`（件数）と各該当の段落 index・前後コンテキストを返す。`--max` 既定 50。

### B3. write: `word.replace`
- CLI: `office word replace "F" --with "R" [--all | --nth K] [--match-case]`
- `Find.Execute(..., Replace:=wdReplaceAll)`（wdReplaceAll=2）。実行前に B2 と同じ検索で `occurrences` を dry-run 報告し、実行後は再度検索して `verified: occurrences_after == 0`（`--nth` は当該置換後に該当数が 1 減ことを確認）を返す。
- 判別できない文字列（全角/半角、正規化差）での取りこぼしに備え、検索フラグは Word の既定（大文字小文字無視・部分一致）を明示オプションで絞る形にする。ワイルドカード・正規表現は本期間では不提供。
- 変更履歴（Track Changes）が立っている文書での挙動が未検証。実行前に `document.TrackRevisions` を読み、true なら `note` を添えて確認を求める。

### B4. write: `word.paragraph.replace_text`
- CLI: `office word paragraph replace <index> --text "…" [--text @file]`
- 段落 index の本文を `Paragraphs(n).Range.Text` の set で置き換える（Phase 2A の parameterized 挙動に注意し、`Paragraphs(n).Range.Text` 経由）。実行後は当該段落を読み戻して `verified`。
- 段落数が想定と変わる（改行を又ぐテキスト等）ため、dry-run で段落数の before/after を報告する。

## Phase 3C：PowerPoint 補完と横断改善

### C1. write: `powerpoint.slide.move`
- CLI: `office powerpoint slide move 72 --to 3`
- `Slides(n).MoveTo(m)`。dry-run は移動前後の index と移動されるスライドのタイトルを報告。実行後は `Slides(m).SlideIndex == m` と移動先のタイトル一致で `verified`。
- 校訂ノートスライド末尾まとめ（実務での典型）を 1 コマンドにするのが目的。

### C2. write: `powerpoint.shape.delete`
- CLI: `office powerpoint shape delete <slide> --shape <index|name>`
- 削除対象の `{index, name, type, text_preview}` を dry-run 報告、実行後は `Shapes.Count` 減少で `verified`。
- 実行結果に `undo_hint: "run 'office undo --yes' to revert if the user confirms"` を付ける（undo 導入済みであり、SKILL.md の「the bridge cannot undo」記述を C4 で更新する）。

### C3. write: `powerpoint.shape.set_format`
- CLI: `office powerpoint shape set-format <slide> --shape <index|name> [--font-size N] [--font-name NAME] [--color RRGGBB] [--bold/--no-bold]`
- `TextFrame.TextRange.Font` の allow-list 書き込み。`apply_to: "text_range"` を返し、図形枠（Line/Fill）には触らない。dry-run は変更予定プロパティと現在値の一覧を返す。
- 色は RRGGBB 文字列を受け取り BGR 整数へ変換（Excel の set-format と同じ表記に統一）。

### C4. 横断: 破壊的操作の応答に `undo_hint`
- `slide delete`・`shape delete`（予定）・`rows delete`（予定）など、実行結果の `undoable`（アプリ側Undoable）と併せて `undo_hint` を返す。
- SKILL.md の「the bridge cannot undo the delete」「confirm first」の記述を「undo 可能だがユーザーの承認が必須」へ更新。
- 低優先: `Application.Undoable` 非公開アプリ（この PowerPoint ビルド）では `undoable: null` のまま。dry-run の `note` で「unknown on this build」と明示するだけにする。

## 実装方針（全項目共通）

- 新規操作は `operations/<app>.py` に実装 → `registry.py` で配線（`_require_app` 検証）→ `capabilities` に登録 → `office_cli/main.py` に CLI（dry-run → 対話確認 → `--yes --expected-document`）→ テスト → SKILL.md / README 日英 更新 → 実文書スモーク。
- 読み取り inspect 系は confirm 不要、`max_text_chars` / 件数上限で切る。書き込み系は `dry_run` 対応と `confirm` 必須の既存規約を守る。
- 保存・終了は引き続き不提供。破壊的 write にも undo を除外しない（undo は 2C で実装済み）。

## テスト方針

- `tests/test_excel_ops.py`: rows insert/delete/columns 系（enum 解決・verified）、rows move（choreography の step 記録と partial）、format read（BGR→RRGGBB・mixed 検出）、copy_format の `borders_applied`
- `tests/test_word_ops.py`: paragraphs 読み（style・上限）、find（件数・max）、replace（occurrences→verified、TrackRevisions note）、paragraph replace
- `tests/test_powerpoint_ops.py`: slide move（MoveTo・verified）、shape delete（count・undo_hint）、shape set-format（apply_to・現在値プレビュー）
- `tests/test_registry.py` / `test_cli.py` / `test_api.py`: 新 operation の配線・capabilities・コマンドペイロード
- 実文書スモーク: Excel 実ブックで rows insert/delete 1 組と row move（要ユーザー承認）、Word 実文書で find（読み取りのみ）、PowerPoint deck で slide move の dry-run と shape delete の dry-run

## 完了条件

- `pytest` が既存テスト含めて全緑（新規 30 件前後を想定）
- `office capabilities` に全 read/write が載り、`requires_confirmation` が正しく付く
- SKILL.md: 三アプリそれぞれの節に新コマンド、Excel の手組み choreography 節を「キュレーション済み操作を使う」へ書き換え、undo 記述を整合
- README 日英: コマンド例・エンドポイント表・用意済み操作リストの更新
- 各アプリ 1 項目以上の実文書スモーク（書き込みはユーザー承認の上）

## やらないこと（本期間外）

- Excel 数式文字列の書き込み（`=…` は引き続き拒否。undo があるから解禁可能だが、検証・監査コストが高い）
- Word の表（Tables）・画像挿入・目次更新、PowerPoint の画像挿入（`AddPicture`）— パラメータが多く dry-run で安全な検証ができる形が未設計
- フィルター・ピボット・マクロ系（`com` 経由の既存経路に委ねる）
- 変更履歴（Track Changes）の読み書きキュレーション（B3 では「立っていたら注意する」だけ）
