# Office COMブリッジ計画案（Python版）

> 修正（v2）：オーナー判断により、Word / Excel / PowerPoint に対する操作を
> 原則すべて解放した。`POST /com`（`office com get|call|set`）でオブジェクト
> モデル全体へのアクセスを提供し、ブロック対象は保存（save*）とクローズ
> （close*、quit / exit）のみに限定した。既存の整理済みコマンドはそのまま
> 利用できる。追加のブロック名は `OFFICE_BRIDGE_BLOCKED_CALLS` で指定可能。

-   対象OS：Windows 11
-   対象：起動済みの Microsoft Word / Excel / PowerPoint
-   エージェント：OpenCodeを既定とし、将来は他のエージェントにも対応
-   方針：Pythonのみで実装。MCPは必須にせず、SkillからCLIを呼び出す

## 1. 目的

Officeアプリ上でホットキーを押すと、その時点のOfficeウィンドウを捕捉する。ブリッジは対応するCOMオブジェクトを保持し、ローカルHTTP
APIとCLIを通じて情報取得・操作を提供する。最終的には、Office用Skillを利用できる状態でOpenCodeを起動し、エージェントがSkillの手順に従ってCLIを使えるようにする。

## 2. スコープ

### 実装対象

-   ホットキー押下時の前景ウィンドウHWNDの取得
-   Word / Excel /
    PowerPointの対象ウィンドウとCOMオブジェクトの対応付け・保持
-   COM操作を専用STAスレッドに集約
-   ローカルHTTP APIとJSON形式
-   CLI
-   エージェント向けOffice Skill
-   Skillを利用できる状態でOpenCodeを起動するPowerShellランチャー
-   基本的な読み取り操作、許可リストに基づく限定的な書き込み
-   ログ、エラー処理、テスト

### 今回は対象外

-   MCPサーバー
-   Office以外のアプリ操作
-   画像認識によるUI操作
-   ~~任意のCOMメソッドを無制限に呼び出すAPI~~（v2で`POST /com`として実装。保存・クローズのみブロック）
-   インターネット越しのAPI公開
-   LLM自体による文書理解処理

## 3. 全体構成

``` text
[Word / Excel / PowerPoint]
           |
         hotkey
           v
[office-bridge: Python常駐プロセス]
  - ホットキー監視 / HWND取得
  - 対象COMオブジェクトの保持
  - 専用STAワーカーと要求キュー
  - localhost限定HTTP API
           ^
           | HTTP + JSON
           |
[office CLI]
           ^
           | コマンド実行
           |
[OpenCode + Office Skill]
```

-   **ブリッジ**：ホットキー、対象管理、COM操作、HTTP APIを担当。
-   **CLI**：HTTP
    APIを呼び出し、エージェントや人間向けのコマンドを提供。
-   **Skill**：コマンドの使い方、操作手順、安全上のルールをエージェントに伝える。
-   **ランチャー**：ブリッジの稼働確認後、Office
    Skillを使える状態でOpenCodeを起動。

初期版ではブリッジとHTTP
APIを1プロセスにまとめ、CLIとエージェントは別プロセスにする。

## 4. 技術選定

  ------------------------------------------------------------------------------------------------
  用途                    採用候補                                      理由
  ----------------------- --------------------------------------------- --------------------------
  Office COM              `pywin32`                                     PythonからOffice
                                                                        Automationを利用

  HTTP API                FastAPI + Uvicorn                             JSON APIを簡潔に実装

  CLI                     Typer                                         サブコマンドを簡潔に実装

  入出力検証              Pydantic                                      APIスキーマを明示

  ホットキー              Win32 `RegisterHotKey`（pywin32経由）         グローバルホットキー

  HWND・プロセス情報      Win32 API（pywin32経由）                      対象ウィンドウを識別

  テスト                  pytest                                        単体・統合テスト

  配布                    当初はPython仮想環境、必要ならPyInstaller等   開発・診断しやすさを優先
  ------------------------------------------------------------------------------------------------

依存関係は`pyproject.toml`で管理し、Pythonのバージョンを固定して再現性を確保する。

## 5. COM接続と対象管理

### 5.1 ホットキー処理

1.  グローバルホットキーを受信。
2.  その時点の前景ウィンドウHWNDを取得。
3.  対象がWord / Excel / PowerPointのウィンドウか検証。
4.  プロセスIDとアプリ種別を確認。
5.  HWND・プロセスID・Officeのウィンドウ情報を照合し、対応するApplicationとDocument
    / Workbook / Presentation等を解決。
6.  ブリッジの「現在の対象」として登録。
7.  `office context`で対象アプリ、文書名、ウィンドウ情報を返す。

### 5.2 COMスレッド規約

`GetActiveObject()`だけでは、複数インスタンスや複数文書がある場合に対象を誤る可能性がある。対象解決処理を独立モジュールにし、複数インスタンスを早期に検証する。

COMオブジェクトは専用STAスレッド上で生成・使用・解放する。HTTPリクエスト処理スレッドから直接COMを操作せず、キュー経由でSTAスレッドに処理を依頼する。STAスレッドではCOM初期化と必要なメッセージポンプを実装する。

文書が閉じられた、またはOfficeが終了した場合はCOM例外を捕捉して対象を無効化し、古い参照を使い続けない。長時間停止したCOM呼び出しを安全に強制中断できると仮定しない。

### 5.3 対象切り替えルール

-   ホットキーを押すたびに、その時点の前景Officeウィンドウへ対象を切り替える。
-   Office以外のウィンドウなら、現在の対象を暗黙に変更せず、エラーまたは通知を返す。
-   対象切り替えはログに記録する。
-   API応答には対象アプリ名と文書名を含め、誤操作を検知しやすくする。

## 6. HTTP API案

`127.0.0.1`だけで待ち受け、外部ネットワークには公開しない。

  メソッド   パス              用途
  ---------- ----------------- --------------------------------------
  `GET`      `/health`         稼働状態
  `GET`      `/context`        現在の対象情報
  `GET`      `/capabilities`   利用可能な操作一覧
  `POST`     `/inspect`        選択範囲・文書・スライド等の情報取得
  `POST`     `/invoke`         許可リストにある操作の実行
  `POST`     `/set-property`   許可されたプロパティの変更

応答はJSONに統一し、`ok`、`result`または`error`、`context`を含める。エラーには機械可読コードと人間向け説明を付ける。

``` json
{
  "ok": true,
  "context": {
    "application": "Excel",
    "document": "Budget.xlsx"
  },
  "result": {
    "address": "A1:B2",
    "values": [[1, 2], [3, 4]]
  }
}
```

設計方針：

-   任意のCOMメソッド名を自由に実行できるAPIは作らない。
-   操作名・引数をスキーマ化し、許可リストで制限する。
-   読み取りと書き込みを明確に分ける。
-   大きな範囲の読み取りにはサイズ上限を設ける。
-   保存・削除など影響の大きい操作は専用コマンドにし、初期版では提供しないか明示確認を必須にする。
-   必要になった段階で`/api/v1`などのバージョン管理を導入する。

## 7. CLI案

コマンド名は仮に`office`とする。

``` text
office health
office context
office capabilities
office inspect selection
office excel range read "A1:D20"
office excel range write "A1:B2" --json '[[1,2],[3,4]]'
office word selection read
office powerpoint slide inspect
```

CLIはJSONを標準出力へ出し、診断メッセージは標準エラーへ出す。終了コードも成功・失敗で統一する。書き込みコマンドは意図が明確な名前にし、必要に応じて`--dry-run`を提供する。

## 8. 初期対応操作

### 共通

-   現在の対象アプリと文書名
-   選択状態の基本情報
-   利用可能な操作一覧

### Excel

-   選択範囲のアドレスと値
-   指定範囲の値
-   指定範囲への二次元配列の書き込み
-   必要に応じて数式と表示値を分けて取得

### Word

-   選択テキストの取得
-   文書の基本情報
-   後続候補として、選択範囲へのテキスト挿入

### PowerPoint

-   現在スライド番号とスライド数
-   現在スライドの図形テキスト
-   スライドタイトルや図形一覧

初期版は読み取りを優先し、書き込みはExcelの範囲更新など限定した操作から導入する。

## 9. SkillとOpenCode起動

### 9.1 Skillに書くルール

-   操作前に`office context`を実行して対象文書を確認する。
-   必要に応じて`office capabilities`を確認する。
-   読み取りには読み取りコマンドを使う。
-   書き込み前に対象文書・範囲・変更内容を確認する。
-   大規模な変更や不可逆操作ではユーザーに確認する。
-   CLIのJSON結果を根拠として次の操作を決める。
-   対象不明、対象変更、Office応答停止時は推測で続行せず状態を再確認する。
-   未定義操作や任意のCOM呼び出しを試みない。

### 9.2 推奨ディレクトリ構成

``` text
office-agent/
  pyproject.toml
  README.md
  src/
    office_bridge/
      __init__.py
      server.py
      com_worker.py
      target.py
      api.py
      operations/
        common.py
        excel.py
        word.py
        powerpoint.py
    office_cli/
      __init__.py
      main.py
  skills/
    office/
      SKILL.md
  scripts/
    start-bridge.ps1
    start-agent.ps1
  tests/
```

Skill原本は`skills/office/SKILL.md`に置く。OpenCodeが検出する配置場所やメタデータ形式は、インストール済みバージョンの仕様を確認して確定する。必要なら、原本からOpenCode対応の場所へコピーするセットアップ処理を用意する。

### 9.3 ランチャーの処理

`start-agent.ps1`を入口とし、次の処理を行う。

1.  プロジェクトの仮想環境を有効化。
2.  `office health`でブリッジの稼働を確認。
3.  停止中ならブリッジを起動。
4.  healthの成功を期限付きで待機。固定時間のsleepだけに依存しない。
5.  SkillがOpenCodeから参照できることを確認。
6.  OpenCodeを起動。
7.  OpenCode終了後、ブリッジを常駐させるか終了させるかを設定に従って処理。

概念例：

``` powershell
# scripts/start-agent.ps1（概念例）
$ErrorActionPreference = "Stop"

. .\.venv\Scripts\Activate.ps1

office health
if ($LASTEXITCODE -ne 0) {
    Start-Process -FilePath "python" `
        -ArgumentList "-m office_bridge.server" `
        -WorkingDirectory (Get-Location)
    # 実装時はhealth成功を期限付きで待つ
}

# Skillの探索場所・起動オプションはOpenCodeの使用バージョンに合わせる
opencode
```

これは構成の概念例であり、OpenCodeのSkill指定オプションを確定した実装コードではない。実装フェーズで現在のOpenCodeのSkill検出仕様を確認する。二重起動防止、ログ、タイムアウト、終了コードも実装する。

### 9.4 他エージェントへの拡張

ブリッジとCLIはOpenCode固有機能に依存させない。最初にOpenCode向けSkillを用意し、必要になったらHermes
Agent等の形式に合わせた薄いアダプターを追加する。

## 10. セキュリティと安全性

-   HTTP APIは`127.0.0.1`のみで待ち受ける。
-   不要なCORS許可を行わず、ブラウザーからの直接利用は初期版の対象外。
-   API側に操作許可リストを設ける。
-   読み取りと書き込みを区別する。
-   保存、上書き、削除、メール送信など外部への影響がある操作は初期版では提供しないか、明示確認を必須にする。
-   ログには操作名、対象、成否、所要時間を記録する。文書内容や選択テキストは無条件に記録しない。
-   外部ネットワークへ公開する場合は認証・認可を別途設計する。
-   Skillの説明だけに依存せず、ブリッジ側でも許可操作を制限する。

## 11. エラーと運用

想定エラー：

-   対象ウィンドウがOfficeではない
-   対象COMオブジェクトを解決できない
-   文書が閉じられた／Officeが終了した
-   Officeが応答しない、COM呼び出しが失敗した
-   API入力が不正
-   範囲が大きすぎる
-   未許可操作
-   ブリッジ未起動／二重起動

エラーは統一JSON形式で返し、CLIは終了コードを返す。COM呼び出しが停止した際に安全に強制中断できると仮定しない。状態の無効化、ユーザー通知、必要に応じたブリッジ再起動を回復方針とする。

## 12. 実装フェーズ

### Phase 1：COM接続の技術検証

-   Pythonから起動済みExcelに接続
-   アプリ、ウィンドウ、ブック、選択範囲を取得
-   複数Excelインスタンスの識別を検証
-   Word / PowerPointでも接続を確認

**完了条件：** 対象を取り違えず基本情報を取得できる。

### Phase 2：ブリッジ基盤

-   STAワーカーと要求キュー
-   対象コンテキスト管理
-   HTTP API、health、context、capabilities
-   JSON形式、ログ、例外処理、二重起動防止

**完了条件：** HTTPクライアントから対象情報を取得できる。

### Phase 3：CLIと基本操作

-   Typer CLI
-   Excel範囲読み取り
-   Word選択テキスト読み取り
-   PowerPointのスライド・図形テキスト取得
-   許可リストによる限定的な書き込み

**完了条件：** 安定したJSON出力をエージェントから利用できる。

### Phase 4：ホットキー

-   グローバルホットキー登録
-   HWNDとOffice対象の照合
-   対象切り替えとログ
-   無効なCOM参照の処理

**完了条件：**
複数のOffice文書を開き、ホットキーで対象を切り替えられる。

### Phase 5：SkillとOpenCodeランチャー

-   Office Skillを作成
-   OpenCodeのSkill配置・検出仕様を実機確認
-   ブリッジ自動起動とhealth待機
-   Skillを利用可能な状態でOpenCodeを起動
-   エージェントに読み取りタスクを実行させる
-   限定的な書き込みタスクを検証

**完了条件：**
ランチャーからOpenCodeを起動し、Skill経由でCLIを呼び出してOffice文書情報を取得できる。

### Phase 6：安定化

-   pytestによる単体テスト
-   Office実機での統合テスト
-   文書切り替え、Office終了、大きな範囲などの異常系テスト
-   ログと診断コマンド
-   必要に応じて実行ファイル化を検討

## 13. テスト計画

1.  Excelでホットキーを押し、`office context`が対象ブックを返す。
2.  選択範囲の値とアドレスが正しい。
3.  別のExcelブックでホットキーを押すと対象が切り替わる。
4.  Wordで選択テキストを取得できる。
5.  PowerPointで現在スライドのテキストを取得できる。
6.  対象文書を閉じてもブリッジがクラッシュしない。
7.  Office以外のウィンドウでは誤操作しない。
8.  不正入力、未知の操作、大きすぎる範囲を拒否する。
9.  OpenCodeがSkillを認識し、CLIの利用ルールに従う。
10. 書き込み前に対象と変更内容を確認する。

## 14. 初期リリースの完了条件

-   ホットキーでWord / Excel / PowerPointの対象を捕捉できる。
-   COM対象を保持し、対象切り替えができる。
-   CLIで対象情報と基本的な文書情報をJSON取得できる。
-   各Officeアプリに少なくとも1つの実用的な読み取り操作がある。
-   書き込み操作は許可リストで制御される。
-   OpenCodeがOffice Skillを認識し、Skill経由でCLIを実行できる。
-   APIはlocalhost限定で、主要エラーを診断できる。

## 15. 先に確認すべき技術リスク

1.  **HWNDから正しいCOM対象への対応付け**：最大のリスク。複数インスタンスを早期に検証する。
2.  **COMのSTA運用**：専用スレッド、初期化、メッセージポンプ、参照寿命を一貫させる。
3.  **OpenCodeのSkill検出方法**：インストール済みバージョンの仕様を確認してからランチャーを確定する。
4.  **操作安全性**：Skillの説明だけに頼らず、API側で操作を制限する。
5.  **Officeの応答停止**：強制中断を過信せず、診断と回復策を設計する。

## 16. 最初に実施する作業

最初の作業はホットキーやエージェント起動ではなく、**起動済みExcelをPythonから特定し、選択範囲を読む小さな検証プログラム**とする。

ここで複数インスタンスの識別方法を検証する。その後、STAワーカー、HTTP
API、CLI、ホットキー、Skill、OpenCodeランチャーの順に積み上げる。最大の技術リスクを、全体を作り込む前に検証できる。
