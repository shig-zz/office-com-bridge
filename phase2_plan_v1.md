# Phase 2 計画：COM エラー体験の改善と PowerPoint 書込系の追加

作成日: 2026-11-11
**状態: 2A・2B 実装完了（2026-11-11）。pytest 269件全緑。実文書でのスモークテスト済み（slides read / text replace / textbox add / slide add のdry-run、`did_you_mean`、`--value @file`）。Phase 2C は未着手。**

起点: 20210416_3次元グラフィックスの基礎知識.pptx のファクトチェック／校訂作業（読み取り69枚、書き込み約33回、失敗2回、検証読み戻し約30回）で観測した実運用上の摩擦。

## 観測した問題

| # | 事象 | 根拠 | 影響 |
|---|---|---|---|
| 1 | `TextFrame.Text` が get/set とも `Member 'Text' was not found` | `com_ops.py:292 _unknown_member()` は候補も代替も示さない | 引数付きプロパティだと分かるまで数往復。代替経路の調査コスト |
| 2 | 引数付き（parameterized）プロパティの set 手段が無い | `execute()` の set は `setattr` のみ（com_ops.py:191-196） | `TextRange.Text` へ書き換えるしかない。`Paragraphs.Count` も同じ理由で読めない |
| 3 | 文字列リテラルの escape が不足 | `com_expr.py:24 _ESCAPES` に `\v` と `\uXXXX` が無い | PowerPoint の段落（CR）や縦タブを CLI 経由で安全に渡せない。PowerShell の `` `r `` 依存 |
| 4 | 長文・日本語の `--value` がシェル引用符とコードページに左右される | CLI は `--json`/`--file` を excel 系だけサポート（main.py:249） | com set での多段落書き込みが事故りやすい |
| 5 | 書き込み後に「未保存」かがレスポンスに出ない | 書き込み系 payload に `Saved` が無い | ユーザーへ保存を促す判断を毎回人力でやる必要がある |
| 6 | PowerPoint にキュレーションされた read/write が無い | `operations/powerpoint.py` は read のみ（5.8KB）。`slide inspect` は表示中スライドのみ | 全文把握のためファイル直接読み込み（python-pptx）に回避。書き込みは全て com 手組み（置換7・テキストボックス8・スライド5追加で約25回の書き込み＋約30回の検証読み） |
| 7 | 図形 index がオフライン情報と COM で対応している保証が無い | 追加系 call が生成物の index を返さない | `Shapes(7)` を別図形に誤爆させるリスク。検証読みが必須になる |

## Phase 2A：COM 層の改善（今回実装）

### A1. 未知メンバーの候補とヒント（`com_ops.py`）
- `_unknown_member(name, exc, owner)` に変更し、`details.did_you_mean` として `dir(owner)` 由来の候補を最大12件返す。
- 同じ名前のメンバーが実際には呼び出し可能（=parameterized property）なら `details.hint` に「call として呼ぶ／代替経路」を出す。
- owner が `TextRange` を持つのに `Text` が読めないケース（PowerPoint TextFrame）は、代替として `TextRange.Text` を hint に明記する。

### A2. parameterized property の set フォールバック（`com_ops.py`）
- `setattr` が unknown-member（`_UNKNOWN_MEMBER_HRESULTS` もしくは AttributeError）で失敗した場合、`getattr(owner, name)(value)` を試行し、成功時に `applied_via: "parameterized_property"` を返す。
- 失敗時は従来の `COM_ERROR` を返す（動作の後方互換）。
- `com get` の末尾括弧は引き続き不可（`call` に委ね、エラーメッセージで parameterized property を案内する）。

### A3. escape の拡張（`com_expr.py`）
- `_ESCAPES` に `"v"`（縦タブ）と `"/"` を追加。
- `\uXXXX`（4桁 hex）を `_parse_string` で処理する。不正な hex は従来どおり `INVALID_REQUEST`。

### A4. CLI の `--value @file`（`office_cli/main.py`）
- `com set` と `set-property` の `--value` が `@` で始まる場合、そのパスを UTF-8 のテキスト/JSON として読み込む（`--json-file` と同じ読み方。先頭が `[`/`{`/`"` なら JSON、そうでなければ素の文字列）。
- 日本語の長文・複数行テキストをシェル引用符なしで安全に渡せるようにする。

### A5. 書き込み応答に未保存状態（`com_ops.py`）
- `call`/`set` の実行時 payload に `unsaved_changes`（`document.Saved` の反転。読取不可なら `null`）を追加する。
- 保存・終了自体は引き続き block しない（ブロックリスト変更なし）。

## Phase 2B：PowerPoint キュレーション操作（今回実装）

`operations/powerpoint.py` に追加し、`registry.invoke` と CLI に配線する。すべて `expected_document`・`dry_run`・`confirm` の既存規約に従う。

### B1. read: `powerpoint.slides`
- CLI: `office powerpoint slides read [--start N] [--count M]`
- 戻り値: `slides[] = {slide_index, title, notes, shapes[]}`、`shapes[] = {index, id, name, type, text, left, top, width, height, unit: "pt"}`
- `--start`（既定1）/`--count`（既定20、上限100）で分割。`max_text_chars` で切り捨て、`truncated` を返す。
- 単位はすべてポイント（`PageSetup.SlideWidth/Height` も併記）。

### B2. write: `powerpoint.text.replace`
- CLI: `office powerpoint text replace <slide> --shape <index|name> --find F --replace R`
- dry-run では「検索語の該当数（0/1/複数）」とプレビューを返し、実行後は本文を読み戻して `verified` と `text_preview` を返す（`range sort` と同じ検証方針）。
- `TextRange.Replace` を使うため書式とレイアウトを保持する。

### B3. write: `powerpoint.textbox.add`
- CLI: `office powerpoint textbox add <slide> --left L --top T --width W --height H --text "…" [--font-size N]`
- `AddTextbox(1, …)`（単位はポイント）→ `TextRange.Text` → `Font.Size` の順に適用。生成図形の `index`/`id`/`name`/`geometry` を返し、後続呼び出しの index 推定を不要にする。
- dry-run ではスライドサイズ内に収まるか（`fits_in_slide`）と重なり候補（同一スライドの他図形の矩形）を報告する。

### B4. write: `powerpoint.slide.add`
- CLI: `office powerpoint slide add [--index N] [--layout 2]`
- `Slides.Add(index, layout)`。実行後は新スライドの index と placeholder（`Shapes` の index/name）を返す。
- Excel のシート配置引数解決（com_ops）は PowerPoint 側では作用しないことをテストで固定する。

## Phase 2C：今回やらない（次フェーズ案）

- `office undo` / `application.Undoable` の公開（取り消し経路）
- `com batch`（steps を1リクエストで、承認1回）
- `office target prune` と expired 重複の畳み込み、`com_reference_lost` 時の hwnd/pid 再接続試行
- `powerpoint.notes read|set`、`powerpoint.slide delete`
- COM 戻り値の型情報（`result_type`）：`Replace` の戻り値が `bool` なのか `TextRange` の文字列化なのか判別できない問題
- 図形名前指定（`Shapes("TextBox 7")`）のドキュメント化

## テスト方針

- `tests/test_com_expr.py`: `\v`・`\uXXXX`・不正 escape
- `tests/test_com_ops.py`: `did_you_mean`/`hint`、parameterized set フォールバック、`unsaved_changes`
- `tests/test_powerpoint_ops.py`: `slides` 読み取り（分割・切り捨て）、`text.replace`（dry-run の該当数／verified）、`textbox.add`（geometry・created index・fits_in_slide）、`slide.add`
- `tests/test_registry.py`: 新 operation の配線とアプリ検証
- `tests/test_cli.py`: 新コマンドの invoke ペイロード、`--value @file`

## 完了条件

- `pytest` が既存テスト含めて全緑
- `office capabilities` に新 read/write が載る
- スキル `.opencode/skills/office/SKILL.md` を新操作に合わせて更新する（PowerPoint 節の書き換え：`TextFrame.Text` 回避策は新操作で不要になった旨、`did_you_mean` の存在、`--value @file`）
