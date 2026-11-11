# Office COM Bridge

[English](README.md) | 日本語

AIエージェントと、起動中のMicrosoft Officeデスクトップアプリをつなぐローカルブリッジです。HTTPとCLIから操作できます。

ホットキーでWord・Excel・PowerPointのウィンドウを1回押すと、その文書が**ターゲットとして登録**され、`target_id`が振られます。以降のコマンドは名前を挙げたターゲットに届きます。挙げなければ、最後に登録したターゲットです。現在の選択範囲や文書の読み書きができます。よくある操作はコマンドとして用意済みで、それを除いた操作は汎用オブジェクトモデル層(`POST /com`、`office com ...`)が通します。Word・Excel・PowerPointのオブジェクトモデルなら何でも呼べる状態です。

ブリッジがやらないことを先に書いておきます。Officeを起動しません。文書の保存や終了も一切しません。名前の先頭が`save`または`close`のメソッド呼び出しとプロパティ設定(`quit`/`exit`も)は差し戻します。待ち受けはループバックアドレスのみで、COMアクセスはSTAワーカースレッド1本に集約して扱います。

## 動作要件

- Windows 10または11
- Microsoft Officeデスクトップ版(Word、Excel、PowerPoint)
- Python 3.11または3.12
- ランチャーから起動できるエージェントCLI:OpenCode、Claude Code、Codex、Hermes Agentのいずれか(独自コマンドも可)

## インストール

```powershell
python -m venv --system-site-packages .venv
.\.venv\Scripts\python -m pip install -e ".[dev]"
```

`--system-site-packages`は、同じバージョンのPythonに`pywin32`がすでにグローバルへインストール済みの場合に効いてきます。

### バッチファイルによる手軽な手順(Windows)

- `install.bat` — 対応Python(3.11/3.12)で`.venv`を作り、パッケージと開発依存をインストールし、設定済みエージェント(既定はOpenCode)用のOfficeスキルをコピーします
- `configure.bat` — OpenCode、Claude Code、Codex、Hermes Agent、独自コマンドのいずれかを選び、`%LOCALAPPDATA%\office-bridge\agent.json`に保存します。対話操作ではスキルそのもののインストールもその場で実行できます
- `start.bat` — 現在のコンソールでブリッジを起動します。サーバーオプションはそのまま渡せます。例:`start.bat --no-hotkey --port 8733`
- `start-agent.bat` — 必要ならブリッジを起動したうえで、現在のターミナルでOfficeスキルを有効にして設定済みエージェントを立ち上げます。この1回だけ変えたいときは`start-agent.bat -Agent hermes`

## ブリッジの起動

```powershell
.\.venv\Scripts\office-bridge
```

既定値:

- アドレス:`127.0.0.1:8733`
- ホットキー:`Ctrl+Alt+O`
- ログ:`%LOCALAPPDATA%\office-bridge\logs`

Word・Excel・PowerPointで文書を開いてウィンドウにフォーカスし、ホットキーを押してください。その文書が**ターゲットレジストリに登録**されます(既存のターゲットは残ります)。設定済みエージェントが新しいWindows Terminalウィンドウに起動し、Officeスキルが使える状態になっているので、エージェントは間髪入れずにその文書を読み書きできます。作業が関わる文書は1つずつ登録し、`--target <id>`でコマンドを振り分けてください。ターゲット登録が不要なヘルスチェックやテストには`--no-hotkey`を使います:

```powershell
.\.venv\Scripts\office-bridge --no-hotkey --port 8733
```

ホットキーでのターゲット登録は残しながら、エージェントの自動起動だけ止めたい場合は、ブリッジ起動前に環境変数`OFFICE_BRIDGE_LAUNCH_AGENT=0`を設定します。

## CLI

仮想環境を有効化してから使うか、`.\.venv\Scripts\office.exe`を直接呼んでください:

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
office excel range copy-format A1:B2 C1:D2 --borders --yes
office excel range set-format A1:B2 --font-name "MS Gothic" --font-size 11 --halign center --yes
office excel range read-format A1:B2 --borders
office excel range add-hyperlink E35 --url "https://example.com/" --yes
office excel rows insert "2:4" --yes
office excel rows delete 61 --yes
office excel rows move 61 --to 60 --yes
office excel columns insert C --yes
office excel columns delete "A:C" --yes
office word selection read
office word selection insert-text "Replacement" --yes
office word paragraphs read --start 1 --count 50
office word find "検索語" --match-case
office word replace "旧称" --replace-with "新称" --dry-run
office word replace "旧称" --replace-with "新称" --yes
office word paragraph replace 12 --text "新しい本文" --yes
office powerpoint slide inspect
office powerpoint slides read --start 1 --count 20
office powerpoint text replace 3 --shape 2 --find "1979" --replace "1982" --dry-run
office powerpoint textbox add 3 --left 57 --top 452 --width 830 --height 41 --text "補足" --yes
office powerpoint slide add --index 70 --yes
office powerpoint slide move 72 --to 3 --yes
office powerpoint shape delete 5 --shape 2 --dry-run
office powerpoint shape set-format 5 --shape "TextBox 3" --font-size 14 --color FF8000 --yes
office powerpoint notes read 70
office powerpoint notes set 70 --text "発表者ノート" --yes
office powerpoint slide delete 71 --dry-run
office undo --dry-run
office com get 'window.ActiveSheet.Range("A1").Value2'
office com call 'window.ActiveSheet.Rows("2:4").Insert(-4107)' --yes
office com set 'window.Selection.Font.Bold' --value true --dry-run
office com batch --file steps.json --yes
office target prune
office target reconnect t-1a2b3c4d5e6f
```

`--target <target_id>`はサブコマンドの前に置くグローバルオプションで、コマンドを登録済みターゲットへ送ります。省略すると最後のターゲットへ戻る仕様ですが、これは後方互換のための既定動作です。複数の文書を扱うときは全コマンドに`--target`を付けてください。`office target list`は登録済みターゲットすべてを`target_id`と`state`付きで一覧表示し、`office target remove <id>`で登録解除できます。`office target prune`は失効(expired)の記録を一覧から削除します。ウィンドウは開いたままCOM参照だけ失われた場合(SharePoint同期や長いアイドル)は、`office target reconnect <id>`が保存済みのウィンドウハンドルから再接続します。ウィンドウに別の文書が表示されていれば拒否します。

書き込みは確認を要求します。`--dry-run`は検証だけで文書を変えせず、`--yes`は対話確認を省略します。`--expected-document`は、違う文書へ操作してしまう事故を防ぐために使います。

`office com get|call|set`は接続済み文書に対するプロパティ読み・メソッド呼び出し・プロパティ設定のどれでも実行できます(下記の「汎用操作」参照)。

## HTTP API

すべてのエンドポイントはエンベロープ形式で返します:

```json
{
  "ok": true,
  "context": {"application": "Excel", "document": "Book1.xlsx"},
  "result": {}
}
```

| エンドポイント | 役割 |
| --- | --- |
| `GET /health` | ライブネス確認と現在のターゲット概要 |
| `GET /context` | 現在(既定)のターゲットの詳細 |
| `GET /capabilities` | 対応するinspect、操作、プロパティ、汎用操作の情報 |
| `GET /targets` | 登録済みターゲットの`target_id`と`state` |
| `POST /targets/capture` | `hwnd`でのターゲット登録(テスト用。通常はホットキーが登録する) |
| `POST /targets/prune` | 失効(expired)のターゲットを一覧から削除する |
| `GET /targets/{target_id}` | ターゲット1件の詳細。ライブネス確認を伴う |
| `POST /targets/{target_id}/reconnect` | ウィンドウが残っている失効ターゲットを同じ文書へ再接続する |
| `DELETE /targets/{target_id}` | ターゲットの登録解除 |
| `POST /inspect` | 文書・選択範囲ビューの読み取り |
| `POST /invoke` | 用意済み書き込み操作の実行 |
| `POST /set-property` | 用意済みプロパティの変更 |
| `POST /com` | 汎用オブジェクトモデルアクセス(接続済み文書への`get`・`call`・`set`) |
| `POST /com/batch` | 複数のcom stepを1リクエストで順に実行し、最初の失敗で停止する |

操作リクエストは省略可能なトップレベルの`target_id`を受け取り、未指定なら最後のターゲットを使います。操作結果は`target`ブロックから始まり、操作エラーには`error.details.target_id`が付きます。IDが存在しなければ`TARGET_NOT_FOUND`、文書が閉じられていれば`TARGET_LOST`で失敗します。別の文書へフォールバックすることはありません。

エラーは`{"ok": false, "context": ..., "error": {"code", "message", "details"}}`の形です。

## 用意済み書き込み操作

- `excel.range.write` — 値をそのまま書き込む(数式文字列は差し戻し)
- `excel.range.copy_format` — 書式(フォント、塗りつぶし、罫線、数値フォーマット、配置)だけを、同じサイズのターゲット範囲へコピーする。値は変わらない。このビルドは書式貼り付けで罫線を運ばないため、`borders: true`を付けると実行後にセル単位で罫線を再適用する(結果は`borders_applied`で報告)
- `excel.range.set_format` — 書式を明示的に設定する:フォント名(和文フォント対応)、フォントサイズ、太字、斜体、下線、文字色、数値フォーマット、水平・垂直配置、自動折返し、塗りつぶし色、連続罫線と罫線色
- `excel.range.add_hyperlink` — 範囲に本物のハイパーリンクオブジェクトを追加する(青字下線に似せただけの文字はクリックできない)。リンクは範囲全体に掛かり、表示テキストは左上セルの値を使う。範囲にすでにハイパーリンクがあればアドレスだけ更新するので、同じ操作を繰り返しても結果は同じになる。`sub_address`(アンカーまたはシート参照)と`screen_tip`(ツールチップ)は省略可能なオプション。数式風の`=`始まりの文字列は差し戻す
- `word.selection.insert_text`
- `word.replace` — 全文から一致語をWord自身の検索・置換で置き換える(書式は保持され、変更履歴が立っていれば変更履歴として記録される)。dry-runは該当数と最初の段落番号を返し、実行後は再検索して`verified`を付ける。`nth`を渡すと該当の1件だけを置き換える
- `word.paragraph.replace_text` — 段落番号(1始まり)を指定して本文を置き換える(段落終わりは保持)。dry-runは段落数の前後を返し、テキスト内の改行で段落が分裂するのを伝える
- `powerpoint.text.replace` — 指定図形の中身を検索語に置換する(`TextRange.Replace`を使い書式とレイアウトを保持)。dry-runは該当数を返し、実行後は読み戻して`verified`を付ける
- `powerpoint.textbox.add` — 指定スライドにテキストボックスを追加する(位置・サイズはすべてポイント)。dry-runはスライド内に収まるかを判定し、実行後は生成図形の`index`・`id`・`name`を返す
- `powerpoint.slide.add` — 指定位置(既定は末尾)にスライドを追加する(レイアウト番号は既定2=タイトルとコンテンツ)。実行後は新スライドのindexとプレースホルダーを返す
- `powerpoint.notes.set` — 指定スライドの発表者ノートを置き換える。dry-runは現在のノートをプレビューし、実行後は読み戻して`verified`を付ける
- `powerpoint.slide.delete` — 指定番号(1始まり)のスライドを削除する。dry-runは削除対象のタイトルを返し、実行後はスライド数が1減ったことを検証する
- `powerpoint.slide.move` — 指定スライドを新しい位置へ移動する(`Slides(n).MoveTo`)。dry-runはタイトルと移動方向を返し、実行後は同じSlideIDが目的地indexに着いたことを検証する
- `powerpoint.shape.delete` — 図形をindexまたは名前で削除する。dry-runは図形名・種類・テキストプレビューを返し、実行後は図形数が1減ったことを検証して取り消し案内(`undo_hint`)を付ける
- `powerpoint.shape.set_format` — 指定図形のテキスト書式(`font_size`・`font_name`・RRGGBBの`color`・`bold`)を設定する。書き込むのは`TextFrame.TextRange.Font`だけで、図形枠や塗りには触れない。dry-runは現在値と変更予定を返し、実行後は読み戻して`verified`を付ける
- `undo` — 接続済みアプリの直近の操作を取り消す(`application.Undo`)。dry-runは取り消し可能性があるか(`application.Undoable`)を返す
- `excel.active_sheet` — `/set-property`経由で変更

書き込み操作は`dry_run`に対応し、dry_run以外は`confirm`を要求します。プロパティ変更にも`confirm`が必要です。実行した書き込みの結果には`unsaved_changes`(未保存の変更があるか)が付くので、エージェントはユーザーにOfficeで保存するよう伝えられます。

用意済み読み取りには`excel.range.format`(セル単位の塗り・フォント・罫線読み。色はRRGGBBへ変換済み)、`word.paragraphs`(スタイル名と見出しフラグ付きの段落ウィンドウ)、`word.find`(読み取り専用。段落番号と前後の文脈付きで検索語の該当を列挙、確認不要)があります。破壊的な書き込み(`slide delete`・`shape delete`・`rows/columns delete`)の結果には`undo_hint`が付きます。

## 汎用操作

`POST /com`(`office com get|call|set`も)は、Word・Excel・PowerPointのオブジェクトモデル全体をそのまま公開します。用意済みコマンドでカバーされていない操作はここを通してください。リクエストはルート(`application`、`window`、`document`)と、メンバーを連鎖する式で指定します。例:

```json
{"mode": "get", "root": "window", "expression": "ActiveSheet.Range(\"A1\").Value2"}
{"mode": "call", "expression": "window.ActiveSheet.Rows(\"2:4\").Insert(-4107)", "confirm": true}
{"mode": "set", "expression": "Selection.Font.Bold", "value": true, "confirm": true}
```

`get`はプロパティの読み取り、`call`はメソッドの実行、`set`は`value`をプロパティへ代入します。`root`を明示しない限り、式の先頭セグメントがルートを兼ねます。`call`と`set`は`dry_run`に対応し、dry_run以外は`confirm`を要求します。実行した`call`の結果には`result_type`(返り値の型名)も付くため、文字列化されたCOMオブジェクトを同じ表示のスカラ取り違えることはありません。

存在しないメンバー名は`INVALID_REQUEST`で差し戻し、`details.did_you_mean`に対象オブジェクトの実メンバー名の候補、`details.hint`に定番の代替経路(例: `TextFrame.Text`なら`TextRange.Text`)を載せます。`com set`の`--value`は`@path`でUTF-8ファイルから値を読み込めるため、日本語の長文・複数行テキストをシェルの引用符やコードページに左右されず渡せます。式内の文字列リテラルは`\n`、`\t`、`\r`、`\v`(縦タブ)、`\/`、`\uXXXX`のエスケープに対応します。実行した`call`と`set`の結果には`unsaved_changes`(未保存の変更があるか)が付きます。

`POST /com/batch`(CLIは`office com batch --file steps.json`)は同じ項目に加えて空でない`steps`配列を受け取り、1つのターゲットへ全stepを順に実行し、最初の失敗で停止します(`stopped: true`)。全stepは実行前に構文とブロックリストの検査を終えるため、途中で不正なstepがあれば何も実行されずに`details.step`付きでエラーになります。`confirm`1回でバッチ全体を承認でき、`get`のみのバッチに確認は不要です。

あえて通さないのは保存と終了の2種類だけです。名前の先頭が`save`または`close`のメンバー(`quit`と`exit`も)は、メソッド呼び出しとプロパティ設定で`OPERATION_BLOCKED`として差し戻します。プロパティの読み取り(`document.Saved`など)はいくらでもできます。拒否リストへの追加は`OFFICE_BRIDGE_BLOCKED_CALLS`(カンマ区切り)で。

## エージェントの選択

`configure.bat`を引数なしで実行すると対話メニューで選べます。`office-agent` CLIでも同じ操作ができます:

```powershell
.\.venv\Scripts\office-agent list
.\.venv\Scripts\office-agent set hermes
.\.venv\Scripts\office-agent show
.\.venv\Scripts\office-agent install-skill
```

コンソールスクリプトを経由しない形でも操作できます:

```powershell
.\.venv\Scripts\python -m office_bridge.agent_config set hermes
```

プリセットとスキル配置先の既定:

| エージェント | コマンド | プロジェクトスキル | ユーザースキル | 既定スコープ |
| --- | --- | --- | --- | --- |
| `opencode` | `opencode` | `.opencode/skills/office` | `~/.config/opencode/skills/office` | project |
| `claude` | `claude` | `.claude/skills/office` | `~/.claude/skills/office` | project |
| `codex` | `codex` | `.agents/skills/office` | `~/.agents/skills/office` | project |
| `hermes` | `hermes -s office` | `.hermes/skills/office` | `%HERMES_HOME%\skills\office` | user |
| `custom` | `--command` | `--skill-dir`または無効 | `--skill-dir` | project |

Hermesは既定スコープがuserです。プロジェクトローカルのスキルは`hermes skills trust`が必要なため、ユーザー側における配置を選びました。スキルは`-s office`でプリロードもします。独自エージェントでは、実行ファイルと追加引数を明示してください:

```powershell
.\.venv\Scripts\office-agent set custom --command "C:\Tools\agent.exe" --arg=--flag --no-install-skill
```

`--scope user`でプリセットのスキルをグローバルへ入ります。`--skill-dir <path>`は配置先の直接指定です。`--no-preload`はHermesの`-s office`引数を立てないにします。`--no-install-skill`はスキル管理をユーザー側が握ります。

保存される設定ファイルは`%LOCALAPPDATA%\office-bridge\agent.json`です。1プロセスだけ変えたいときは`OFFICE_BRIDGE_AGENT`(例:`OFFICE_BRIDGE_AGENT=codex`)、参照するファイルそのものを変えるときは`OFFICE_BRIDGE_AGENT_CONFIG`を使います。

ランチャーは、必要ならブリッジを起動し、選んだエージェントへスキルをインストールし、そのエージェントを現在のディレクトリに立ち上げます:

```powershell
.\scripts\start-agent.ps1
.\scripts\start-agent.ps1 -Agent hermes
```

スキルには、CLIをいつどう使うかという使い分けが書き込んであります。

## テスト

```powershell
.\.venv\Scripts\python -m pytest
```

## 安全策

- ループバックアドレスのみで待ち受ける
- リクエストはSTAワーカースレッド1本で直列に処理する
- ターゲットレジストリは、閉じられるまで登録済み文書をすべて保持する。ターゲットごとの`target_id`は不変で、文書またはアプリが閉じるとターゲットは無効になる(黙って差し替えることはない)
- 読み取りは既定で10,000セル・100,000文字まで
- 汎用操作で保存と終了は全面的に差し戻す。それ以外の書き込みはすべて`confirm`を要求し、`dry_run`に対応する
