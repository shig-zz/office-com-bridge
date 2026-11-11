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
office excel range set-format A1:B2 --font-name "MS Gothic" --font-size 11 --halign center --yes
office excel range add-hyperlink E35 --url "https://example.com/" --yes
office word selection read
office word selection insert-text "Replacement" --yes
office powerpoint slide inspect
office powerpoint slides read --start 1 --count 20
office powerpoint text replace 3 --shape 2 --find "1979" --replace "1982" --dry-run
office powerpoint textbox add 3 --left 57 --top 452 --width 830 --height 41 --text "補足" --yes
office powerpoint slide add --index 70 --yes
office com get 'window.ActiveSheet.Range("A1").Value2'
office com call 'window.ActiveSheet.Rows("2:4").Insert(-4107)' --yes
office com set 'window.Selection.Font.Bold' --value true --dry-run
```

`--target <target_id>`はサブコマンドの前に置くグローバルオプションで、コマンドを登録済みターゲットへ送ります。省略すると最後のターゲットへ戻る仕様ですが、これは後方互換のための既定動作です。複数の文書を扱うときは全コマンドに`--target`を付けてください。`office target list`は登録済みターゲットすべてを`target_id`と`state`付きで一覧表示し、`office target remove <id>`で登録解除できます。

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
| `GET /targets/{target_id}` | ターゲット1件の詳細。ライブネス確認を伴う |
| `DELETE /targets/{target_id}` | ターゲットの登録解除 |
| `POST /inspect` | 文書・選択範囲ビューの読み取り |
| `POST /invoke` | 用意済み書き込み操作の実行 |
| `POST /set-property` | 用意済みプロパティの変更 |
| `POST /com` | 汎用オブジェクトモデルアクセス(接続済み文書への`get`・`call`・`set`) |

操作リクエストは省略可能なトップレベルの`target_id`を受け取り、未指定なら最後のターゲットを使います。操作結果は`target`ブロックから始まり、操作エラーには`error.details.target_id`が付きます。IDが存在しなければ`TARGET_NOT_FOUND`、文書が閉じられていれば`TARGET_LOST`で失敗します。別の文書へフォールバックすることはありません。

エラーは`{"ok": false, "context": ..., "error": {"code", "message", "details"}}`の形です。

## 用意済み書き込み操作

- `excel.range.write` — 値をそのまま書き込む(数式文字列は差し戻し)
- `excel.range.copy_format` — 書式(フォント、塗りつぶし、罫線、数値フォーマット、配置)だけを、同じサイズのターゲット範囲へコピーする。値は変わらない
- `excel.range.set_format` — 書式を明示的に設定する:フォント名(和文フォント対応)、フォントサイズ、太字、斜体、下線、文字色、数値フォーマット、水平・垂直配置、自動折返し、塗りつぶし色、連続罫線と罫線色
- `excel.range.add_hyperlink` — 範囲に本物のハイパーリンクオブジェクトを追加する(青字下線に似せただけの文字はクリックできない)。リンクは範囲全体に掛かり、表示テキストは左上セルの値を使う。範囲にすでにハイパーリンクがあればアドレスだけ更新するので、同じ操作を繰り返しても結果は同じになる。`sub_address`(アンカーまたはシート参照)と`screen_tip`(ツールチップ)は省略可能なオプション。数式風の`=`始まりの文字列は差し戻す
- `word.selection.insert_text`
- `powerpoint.text.replace` — 指定図形の中身を検索語に置換する(`TextRange.Replace`を使い書式とレイアウトを保持)。dry-runは該当数を返し、実行後は読み戻して`verified`を付ける
- `powerpoint.textbox.add` — 指定スライドにテキストボックスを追加する(位置・サイズはすべてポイント)。dry-runはスライド内に収まるかを判定し、実行後は生成図形の`index`・`id`・`name`を返す
- `powerpoint.slide.add` — 指定位置(既定は末尾)にスライドを追加する(レイアウト番号は既定2=タイトルとコンテンツ)。実行後は新スライドのindexとプレースホルダーを返す
- `excel.active_sheet` — `/set-property`経由で変更

書き込み操作は`dry_run`に対応し、dry_run以外は`confirm`を要求します。プロパティ変更にも`confirm`が必要です。

## 汎用操作

`POST /com`(`office com get|call|set`も)は、Word・Excel・PowerPointのオブジェクトモデル全体をそのまま公開します。用意済みコマンドでカバーされていない操作はここを通してください。リクエストはルート(`application`、`window`、`document`)と、メンバーを連鎖する式で指定します。例:

```json
{"mode": "get", "root": "window", "expression": "ActiveSheet.Range(\"A1\").Value2"}
{"mode": "call", "expression": "window.ActiveSheet.Rows(\"2:4\").Insert(-4107)", "confirm": true}
{"mode": "set", "expression": "Selection.Font.Bold", "value": true, "confirm": true}
```

`get`はプロパティの読み取り、`call`はメソッドの実行、`set`は`value`をプロパティへ代入します。`root`を明示しない限り、式の先頭セグメントがルートを兼ねます。`call`と`set`は`dry_run`に対応し、dry_run以外は`confirm`を要求します。

存在しないメンバー名は`INVALID_REQUEST`で差し戻し、`details.did_you_mean`に対象オブジェクトの実メンバー名の候補、`details.hint`に定番の代替経路(例: `TextFrame.Text`なら`TextRange.Text`)を載せます。`com set`の`--value`は`@path`でUTF-8ファイルから値を読み込めるため、日本語の長文・複数行テキストをシェルの引用符やコードページに左右されず渡せます。式内の文字列リテラルは`\n`、`\t`、`\r`、`\v`(縦タブ)、`\/`、`\uXXXX`のエスケープに対応します。実行した`call`と`set`の結果には`unsaved_changes`(未保存の変更があるか)が付きます。

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
