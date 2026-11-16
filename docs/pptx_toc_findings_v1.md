# PowerPoint 目次記入タスクで分かった問題点と改善案

- 対象バージョン: `0.1.0`（office-com-bridge）
- 検証環境: Windows 10/11、PowerPoint デスクトップ、OpenCode + Office Skill、PowerShell 5.1
- 検証対象文書: `documents/シンFRAMES/CADCore開発状況_テンプレ適用.pptx`（全25スライド、960×540pt）
- 実施タスク: 「目次の内容を埋めてください」→ 3枚目スライドの本文プレースホルダーに 22 段落の目次を記入
- 結果: 記入成功（レベル1=16pt / レベル2=13pt、下位7項目は `IndentLevel=2`）。**オートフィットのみ自動化できなかった**

本書は、この実作業中に発生した詰まりを再現手順ついでに整理し、それぞれの解決案をフェーズ計画に反映するための記録である。
実装の全体像は `docs/implementation.md`、計画系は `phase2_plan_v1.md` / `phase3_plan_v1.md` を参照。

---

## 1. サマリ

| # | 分類 | 問題 | 影響 | 深刻度 |
| --- | --- | --- | --- | --- |
| A1 | 機能ギャップ | 目次生成のキュレーション操作がない | 22段落分を `com set` の生書きで対応。番号・階層の判断を毎回その場でやる | 高 |
| A2 | 機能ギャップ | 「overflow 時に文字を縮小」が設定も読み取りもできない | ユーザーが選んだ方針（自動縮小）を実現できず、フォントサイズ手で決めに切り替え | 高 |
| A3 | 機能ギャップ | テキストが框に収まるかを測る手段がない | はみ出し可否を推定で判断するしかない | 中 |
| A4 | 機能ギャップ | 段落単位の書式（`IndentLevel` / `Font.Size`）のキュレーション操作がない | レベル付けとサイズ指定で com 生書きの batch 2本（計4ステップ）になった | 中 |
| A5 | 機能ギャップ | 枠全体のテキスト置換 op がない（`text.replace` は部分置換のみ） | `com set TextRange.Text` 経由になり、**変更前の値と検証読み戻しが残らない** | 中 |
| A6 | 機能ギャップ | セクション（`Slides(i).Section`）が読めない | 目次の階層推断にセクションを使えず、タイトル一覧だけの手動分類になった | 低 |
| A7 | エラー品質 | テキスト枠なし shape の `TextRange.Text` 読みが内部例外のまま出る | 原因（HasTextFrame=0）が読み取りにくい | 低 |
| B1 | シェル起因 | PowerShell 5.1 が `--value "{...}\r..."` のエスケープを壊す | 段落区切り CR が失われ、`\r` が文字列として届いた | 高 |
| B2 | シェル起因 | BOM なし UTF-8 の `.ps1` が ANSI として読まれ日本語引数が文字化け | `--expected-document` が `DOCUMENT_MISMATCH` になる | 中 |
| B3 | CLI起因 | `Set-Content -Encoding UTF8` が書く BOM を `com batch --file` が拒否 | `Invalid JSON ... Unexpected UTF-8 BOM`。ファイル作成手段が制限される | 中 |
| B4 | 実行環境起因 | portable shell scanner が複雑な PowerShell 一行コマンドを解析不能 | `invalid-structure` で実行自体できない | 中 |
| C1 | 発見性 | `--value @path`（値をファイルから読む形式）が help / capabilities に出ていない | B1・B4 の正解手段が分からず、遠回り（PowerShell 経由）した | 高 |
| C2 | 発見性 | Skill の読み取り表に `office powerpoint slides read` がない | geometry と notes 一括取得を使わず、`com get` を20回以上叩いた | 中 |
| C3 | 安全性 | batch の複数ステップは undo 1手分しか戻せないことが結果に見えない | 巻き戻し前提の設計ができない | 中 |

---

## 2. 問題点の詳述と再現手順

### A1. 目次（TOC）生成のキュレーション操作がない

今回必要だったのは「各スライドのタイトルを収集し、既存の目次プレースホルダーの番号様式（`０．` / `１．` の全角数字＋全角ピリオド）に合わせて書き、関連スライドをレベル2に下げる」ことだった。
これは PowerPoint 文書で非常に頻出の作業なのに、`office powerpoint` サブコマンド（`slide` / `slides` / `text` / `shape` / `textbox` / `notes` / `document`）に対応 op がない。

- 実際に使った手段: `office com batch --file toc_batch.json`（`set ...TextRange.Text` + `set ...Paragraphs(4, 7).IndentLevel`）
- 代价: 番号様式の推定、階層の判断、段落数の検算がすべてエージェント側の手作業に落ちる

### A2. オートフィットが設定も読み取りもできない（最重要の機能ギャップ）

PowerPoint の「テキストオプション → 自動調整 → **溢出時に文字を縮小**」に相当する設定が、このブリッジ経由では触れない。

試した経路と結果:

| 試した表現 | 結果 |
| --- | --- |
| `com set 'document.Slides(3).Shapes(3).TextFrame.AutoSize' --value -1` | `COM_ERROR` / `hresult 0x80020009`（DISP_E_EXCEPTION相当、理由の情報なし） |
| `com set ... TextFrame.AutoSize --value 0` | **成功**（現状値 0 と同じ。書き込み自体は通る） |
| `com set ... TextFrame2.MsoAutoFitBehavior --value 2` | `INVALID_REQUEST` / `Member 'MsoAutoFitBehavior' was not found` |
| `com get ... TextFrame2.MsoAutoFitBehavior` | 同上（読み取りも不可） |
| `com get ... TextFrame.FitText` | `Member 'FitText' was not found` |
| `com get ... TextFrame2.WordWrap` | 成功（`= -1`）→ TextFrame2 自体は辿れている |

つまり `TextFrame2` へは到達できるが、`MsoAutoFitBehavior` は IDispatch の名前解決（`GetIDsOfNames`）で出てこない（隠しメンバー）。`TextFrame.AutoSize` は書けるが `-1`（框を伸ばす）は PowerPoint 側に拒否される。
結果として「自動縮小」を自動適用できず、`16pt / 13pt` への明示的な縮小に切り替えた（ユーザーの承認変更あり）。

### A3. テキストが框に収まるかの測定手段がない

- 框: `Left 146.0` / `Top 75.7` / `Width 422.1` / `Height 442.25`、フォント メイリオ
- 22段落 × 20pt は確実に超過するが、**必要高さを実測できない**
- `AutoSize = -1` が失敗するので「伸ばして収まるか」という間接プローブも使えない
- 結果、行高の推定（16pt/13pt で約440pt ≒ 框442pt）で判断することになり、「ギリギリ」という不確実な報告になった

### A4. 段落単位の書式 op がない

`office powerpoint shape set-format` は shape 全体の `TextRange.Font` のみ。レベル1=16pt / レベル2=13pt のような**段落レベルの指定**は com 生書きが必要だった。

```text
office com batch --file toc_font.json   # TextRange.Font.Size=16 → Paragraphs(4,7).Font.Size=13
```

なお `Paragraphs(始点, 個数)` の2引数形は `IndentLevel`・`Font.Size` とも機能した（良い点。Skill に明記する価値がある）。

### A5. 枠全体テキストの置換 op がない

- `office powerpoint text replace` は部分置換のみ
- 全体書き換えは `com set '...TextFrame.TextRange.Text'` になるが、dry-run も実行結果も **書き換え前の値を返さない**（`powerpoint.notes.set` には `notes_preview` / `verified` がある）
- 旧テキスト（`０．改定履歴\r１．`）は事前に自分で読んでおく必要があり、忘れると失われる

### A6. セクションが読めない

```text
office com get 'document.Slides(4).Section.Name'   # → INVALID_REQUEST
```

目次の階層を決める手がかりとして PowerPoint のセクション分割を調べたが、辿れなかった。タイトル21本の手動分類になった。

### A7. テキスト枠を持たない shape の読み取りエラー

```text
office com get 'document.Slides(25).Shapes(2).TextFrame.TextRange.Text'
→ INVALID_REQUEST "Member 'TextRange' was not found"
  error: 例外が発生しました。(0, None, '指定された値は境界を超えています。', ...)
```

実際は `HasTextFrame = 0` が原因。`did_you_mean: [TextRange, HasText]` は出ていて親切だが、`TEXT_FRAME_MISSING` のような専用コードと `HasTextFrame` の値が欲しい。

### B1. PowerShell 5.1 経由で `--value` の CR が失われた

`ConvertTo-Json` で作った `"..."\r"..."` を PowerShell 経由で渡すと、dry-run の `value` が `"０．改定履歴\\r１．..."`（＝**バックスラッシュと r の2文字**）になった。段落区切り CR が届かず、この経路では正しく書けない。

### B2. BOM なし UTF-8 の `.ps1` 内の日本語が文字化け

`powershell -File script.ps1`（`script.ps1` は BOM なし UTF-8）で `--expected-document "CADCore開発状況_テンプレ適用.pptx"` を書くと、

```json
"message": "Target document mismatch: expected 'CADCore髢狗匱迥ｶ豕＼繝・Φ繝励Ξ驕ｩ逕ｨ.pptx' ..."
```

PowerShell 5.1 が `.ps1` を ANSI（CP932）として読むため。CLI を直接呼ぶ場合は日本語引数は問題ない。

### B3. `com batch --file` が BOM 付き UTF-8 を拒否する

```text
Invalid value: Invalid JSON in ...\toc_apply.json: Unexpected UTF-8 BOM (decode using utf-8-sig)
```

`src/office_cli/main.py` の batch ファイル読み込みは `read_text(encoding="utf-8")`（B3 相当の該当箇所）。PowerShell 5.1 の `Set-Content -Encoding UTF8` は必ず BOM を書くので、**同じシェルの中でファイルを用意する定番手順が通らない**。

### B4. portable shell scanner が解析できないコマンドがある

```powershell
$lines = @('a','b'); $val = ConvertTo-Json ($lines -join "`r") -Compress; office com set ... --value $val
```

→ `Portable shell scanner cannot analyze command: invalid-structure` で実行前に拒否される。変数＋配列＋`-join` を含む一行コマンドは使えない。

### C1. `--value @path` が discoverable でない

`src/office_cli/main.py` の `_resolve_value_argument()` は `@` プレフィックスで UTF-8 ファイルから値を読み込む（コメントにも「日本語テキスト・PowerPoint の段落はコマンドライン経由は不安全」とある）。**まさに今回の正解**だが、

- `office com set --help` の `--value` 説明: 「JSON value or plain string, for example 42, true, "text", [[1,2],[3,4]].」
- `office capabilities` の `com.params.value`: 「type: any」

のどちらも `@path` を触れていない。Skill の記載にもなく、結果的に B1・B2・B4 を全部踏んだ。

### C2. 読み取りの一括手段を Skill が案内していない

`office powerpoint slides read`（スライド窓＋notes＋全 shape の index / name / type / text / geometry）が存在するのに、Skill の読み取り表に `office powerpoint slide inspect` と `office powerpoint document read` しか無い。
ため shape 3個ぶんの `Name` / `Type` / `PlaceholderFormat.Type` / `Text` / `Left` / `Top` / `Width` / `Height` / `Font.Size` を個別 `com get` で集め、HTTP リクエストが20回以上になった。

### C3. batch の undo 単位が見えない

batch は 2〜4 ステップを 1 回の承認で適用するが、`undo` は「直近1アクション」しか戻せず、結果 JSON にも巻き戻し単位が示されない（`powerpoint.slide.delete` には undo hint がある）。

---

## 3. 解決案

### 3.1 即時（Skill / ドキュメント修正のみで効くもの）

1. **`--value @path` と `com batch --file` を第一選択として明記する**（Skill の Writing 節、`office capabilities` の com.params、`com set --help` の `--value` 説明）。
   - 多行・日本語の値はシェルに渡さない、を原則にする
   - 実装は既に用意されている（`com set` の `--value` と `powerpoint text replace` / `textbox add` / `notes set` の `--text` が `@path` 対応、`_resolve_value_argument()` / `_resolve_text_argument()`）。**未記載なだけ**なので、help と capabilities に同じ文言を落とす
   - `--value @file` に加え `--json @file` / `--file` も同様の表記なのか help に列挙する
2. **PowerPoint の段落区切りは CR (`\r`) である旨を Skill に書く**。JSON ファイル経由なら `\r` エスケープで意図どおり届くことを例示する。
3. **PowerShell 呼び出しの注意を Skill に追記する**
   - `.ps1` に日本語リテラルを書かない（ASCII のにし、日本語は UTF-8 の JSON データファイルに逃がす）
   - `Set-Content -Encoding UTF8` は BOM を書くので、CLI に JSON を食わせない
   - 出力を見るなら `[Console]::OutputEncoding = [Text.Encoding]::UTF8`
4. **読み取りは `office powerpoint slides read` / `office word paragraphs` / `office excel range format` を優先**する表に更新する（`com get` の多用を避ける）。
5. `Paragraphs(始点, 個数)` の2引数形が効くことを Skill に記載する（レベル変更・段落書式の一括指定に有効）。

### 3.2 CLI / ブリッジの小修正（低コスト・高効果）

| 案 | 内容 | 対象箇所 |
| --- | --- | --- |
| BOM 吸収 | batch / `@file` / `--file` の読み込みを `utf-8-sig` にする（Excel/Word 側の JSON 読み込みも統一） | `src/office_cli/main.py`（`_load_values`、com batch、`_resolve_value_argument`） |
| 改行正規化 | PowerPoint 側テキスト書き込みで `\r\n` → `\r`、`\n` → `\r` を適用し、結果に `paragraph_breaks` を報告する（Word 側の `paragraph.replace_text` は既に正規化済み。揃える） | `src/office_bridge/operations/powerpoint.py` |
| 専用エラー | `TextFrame.AutoSize` の設定失敗を `AUTOFIT_REFUSED` にマップし、`hresult` と「框をスライド外へ伸ばすため拒否された可能性」「UI での手動操作を推奨」を hint に入れる | `com_ops.py` の `_call` / set 経路 |
| 専用エラー | `HasTextFrame = 0` の shape で `TextFrame.TextRange` を辿ったら `TEXT_FRAME_MISSING`（`has_text_frame: 0`, `shape_type` を添える） | `operations/powerpoint.py::_text_range`、`com_ops._member` |
| 変更前の値 | `com set` の dry-run / 実行結果に `value_before` を入れる（読めるプロパティなら）。少なくとも PowerPoint のテキスト系 set では必須にする | `com_ops.py` の set 経路 |
| undo 単位 | batch 結果に `undo_steps`（適用ステップ数）と「まとめて1回では戻せない」旨の note を入れる | `operations/com_ops.py`（batch） |
| help 強化 | `--value` に `@path`、`com batch` に「BOM なし UTF-8 を推奨」を記載。`capabilities` の `com.params.value` に `file_reference: "@path は UTF-8 ファイルから読み込む"` を追加 | `src/office_cli/main.py`、`operations/registry.py` |

### 3.3 新機能（キュレーション op として追加すべきもの）

1. **`office powerpoint toc fill`**（A1 対策、本命）
   - `--slide 3 --shape 3`（目次スライド／本文プレースホルダー）を指定し、`--from-titles`（既定）で該当スライド範囲のタイトルを収集して記入
   - 既存テキストの番号様式（全角/半角、`０．` の開始番号、区切り）を解析して踏襲
   - `--exclude 3,25`、`--levels "6-12=2"` のような階層指定、または `--levels auto`（タイトルが他タイトルを含む／スライド種別で判断）
   - 結果に `paragraphs_before` / `paragraphs_after` / `numbering_style` / `entries[]`（スライド番号対応）/ `verified` を返す
   - dry-run で最終テキストを全文プレビュー
2. **`office powerpoint text set`**（A5 対策）
   - shape 1個のテキストを段落配列で一括書き換え。`paragraphs: [{"text": "...", "level": 1|2, "font_size": 16}]` を許容し、内部で `Text` 設定 → `Paragraphs(a, b).IndentLevel` / `Font.Size` を順に適用
   - `notes.set` と同じ `preview` / `verified` / `characters_before/after` を返す
3. **`office powerpoint text set-format` の段落拡張**（A4 対策）
   - 既存 `shape set-format` に `--paragraph 4-10`（範囲指定）と `--indent-level 2` を追加
4. **`office powerpoint text set-autofit`**（A2 対策）
   - `--behavior none|shape|text` を受け、次の順で試す:
     1. `TextFrame2.MsoAutoFitBehavior` を **早期バインドで**設定（`win32com.client.CastTo(shape.TextFrame, "TextFrame2")` 等、type library 経由の dispID 呼び出し。動的名前解決が出さない隠しメンバーに備える）
     2. だめなら `TextFrame.AutoSize`（`shape` 相当）を試し、拒否されたら `AUTOFIT_REFUSED` を返す
   - いずれにせよ読み戻し（適用後の振る舞い）を結果に含める。読み取り専用 `... text get-autofit` も併設
5. **`office powerpoint text measure`**（A3 対策）
   - `Shape.Duplicate()` で複製を作り、複製に autofit（框を伸ばす）をかけて **必要高さを実測** → 複製を削除して `required_height` / `fits: true|false` / `margin` を返す
   - 原文字形・フォント・框幅は複製が使으므로精度が高い。「収まるか分からないまま書く」をなくす
   - オプションで `--font-size` を与えると、そのサイズでの必要高さを返す（サイズ選択の自動決定に使える）
6. **`office powerpoint sections read`**（A6 対策）
   - `Slides(i).Section` を辿る現経路は失败するため、`document.Sections`（PowerPoint の Sections コレクション）から「セクション名 → スライド番号範囲」を集めて返す

### 3.4 目次スライドのテンプレ側改善（文書運用）

ブリッジ側の修正と並行で、テンプレート自体にも手を入れると今後の目次作業が楽になる。

- 目次本文プレースホルダーのレベル既定値を **1レベル 16pt / 2レベル 13pt** にしておく（20pt 統一だと22行は必ずはみ出す）
- レイアウト側に「溢出時に文字を縮小」を仕込んでおく（ブリッジで設定できないので、**テンプレで一度だけ UI 設定**するのが最短）
- 項目が多いときは 2 列化（右側に同じサイズのテキストボックス）が有効。今回 `powerpoint textbox add` で対応可能だった（未実施）

---

## 4. 実際の操作手順（再現用・付録）

正解として機能したのは JSON ファイル + `com batch`。一時ファイルの中身は以下のとおり。

`toc_batch.json`（BOM なし UTF-8。`--file toc_batch.json --dry-run` → `--yes` の順に実行）

```json
[
  {
    "mode": "set",
    "root": "document",
    "expression": "document.Slides(3).Shapes(3).TextFrame.TextRange.Text",
    "value": "０．改定履歴\r１．モチベーション\r２．部品の種類\rアルミフレーム\rパネル\r安全柵用配線ラック\r部品と型番の粒度\r直角接続ブラケット\rアジャスタパッド＆キャスタユニット\r打掛錠\r３．Problems\r４．Solution – OpEngine Architecture\r５．スクリプト言語\r６．デモ\r７．CADCoreの階層構造\r８．CADCoreの実装技術\r９．ちょっと脱線…OSS CADカーネル\r１０．実装方針\r１１．モデリング機能\r１２．今後の予定\r１３．スクリプト言語によるテスト基盤\r１４．テスト基盤があるとなにが良いか？"
  },
  {
    "mode": "set",
    "root": "document",
    "expression": "document.Slides(3).Shapes(3).TextFrame.TextRange.Paragraphs(4, 7).IndentLevel",
    "value": 2
  }
]
```

`toc_font.json`（フォント縮小。同上の手順で dry-run → yes）

```json
[
  {
    "mode": "set",
    "root": "document",
    "expression": "document.Slides(3).Shapes(3).TextFrame.TextRange.Font.Size",
    "value": 16
  },
  {
    "mode": "set",
    "root": "document",
    "expression": "document.Slides(3).Shapes(3).TextFrame.TextRange.Paragraphs(4, 7).Font.Size",
    "value": 13
  }
]
```

確認した読み戻し値:

| 項目 | 値 |
| --- | --- |
| `TextRange.Paragraphs().Count` | 22 |
| `Paragraphs(3).Font.Size` / `Paragraphs(4).Font.Size` | 16.0 / 13.0 |
| `Paragraphs(10).IndentLevel` / `Paragraphs(11).IndentLevel` | 2 / 1（下位項目は 4〜10） |
| `TextRange.Font.Name` | メイリオ |
| `Top` / `Height` | 75.73 / 442.25（変更なし） |
| `TextFrame.AutoSize` | 0（`-1` は設定失敗） |

---

## 5. 優先度つきの着手順

1. `--value @path` / `com batch --file` を help・capabilities・Skill に明記（C1。実働ゼロ分、最も費用対効果が高い）
2. batch / `@file` 読み込みを `utf-8-sig` に（B3）
3. `powerpoint.text.set`（段落配列＋レベル＋サイズ、before/after 付き）（A4・A5）
4. `powerpoint.text.set-autofit`（早期バインド）＋ `AUTOFIT_REFUSED` 専用エラー（A2）
5. `powerpoint.text.measure`（Duplicate 実測）（A3）
6. `powerpoint.toc.fill`（A1。3〜5 の上に作る）
7. `com set` の `value_before`、batch の `undo_steps`、`TEXT_FRAME_MISSING`、`sections read`（A5・C3・A7・A6）
8. Skill の読み取り表を `powerpoint slides read` / `word paragraphs` 優先に更新（C2）
