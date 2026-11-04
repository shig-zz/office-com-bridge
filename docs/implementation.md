# Office COM Bridge 実装解説

- 対象バージョン: `0.1.0`
- 対象環境: Windows 10/11、Microsoft Office デスクトップ（Word / Excel / PowerPoint）、Python 3.11 / 3.12
- 本稿執筆時点の検証: `python -m pytest -q` → **140 passed**
- ライセンス: MIT

本書は、起動済み Office アプリケーションを AI エージェントから読み書きするために実装した
`office-com-bridge` の技術的な全体像を、設計判断・処理フロー・制約まで含めて説明する。
利用方法の要約は `README.md`、当初計画は `office_com_bridge_plan.md` にある。

---

## 1. 完成したもの

起動済みの Word / Excel / PowerPoint を対象に、次の機能を提供するローカルブリッジである。

1. **グローバルホットキーによる対象捕捉**
   既定の `Ctrl+Alt+O` を押した瞬間の前景 Office ウィンドウを特定し、そのウィンドウに
   対応する COM オブジェクトを「現在の対象」として保持する。
2. **localhost 限定の HTTP API**
   `127.0.0.1:8733` で JSON API を提供する。対象情報、読み取り、許可済み書き込み、
   汎用 COM 操作を提供する。
3. **CLI**
   `office` コマンドで上記 API を呼ぶ。エージェントにも人間にも扱いやすい JSON を返す。
4. **キュレーション済み操作**
   Excel の範囲読み書き・書式コピー・書式設定、Word の選択範囲へのテキスト挿入、
   Excel のアクティブシート変更など、頻度の高い操作を型付きで提供する。
5. **汎用オブジェクトモデル操作**
   `POST /com` / `office com get|call|set` で、Word / Excel / PowerPoint の任意の
   プロパティ読み取り、メソッド呼び出し、プロパティ設定を行う。保存・クローズ系のみブロックする。
6. **エージェント連携**
   Office Skill をエージェント別のディレクトリに配置し、ブリッジと OpenCode / Claude Code /
   Codex / Hermes Agent / 任意コマンドをまとめて起動するランチャーを提供する。
7. **安全性と診断**
   loopback バインド、STA スレッドへの COM 集約、読み取り上限、書き込み確認、
   対象文書名の照合、構造化エラー、ローテーションログ、単一インスタンス制御を実装する。

### 1.1 計画フェーズとの対応

| フェーズ | 内容 | 状態 |
| --- | --- | --- |
| Phase 1 | COM 接続の技術検証（HWND → Application / Document / Selection） | 実装済み。実機ログで Excel の捕捉を確認 |
| Phase 2 | STA ワーカー、要求キュー、対象管理、HTTP API、ログ、二重起動防止 | 実装済み |
| Phase 3 | Typer CLI、Excel 読み取り、Word / PowerPoint 読み取り、限定書き込み | 実装済み |
| Phase 4 | グローバルホットキー、HWND 照合、対象切替、無効参照処理 | 実装済み。実機ログでホットキー登録と対象切替を確認 |
| Phase 5 | Office Skill、エージェント設定、ブリッジ自動起動 | 実装済み |
| Phase 6 | pytest、異常系、実機統合、診断、実行ファイル化 | 単体テストと手動実機確認まで。実機自動テストと配布用パッケージ化は未着手 |

### 1.2 意図的な非目標

- MCP サーバーにはしない。CLI + Skill でエージェント非依存に近づける。
- Office を起動しない。既に開かれている文書だけを対象にする。
- 文書を保存・クローズしない。これは利用者の明示操作に残す。
- インターネット公開、認証、TLS、CORS は対象外。loopback 専用。
- 画像認識による UI 操作、文書内容の LLM 解析、複数対象の同時保持は対象外。
- 文書の新規作成、ファイルを開く操作、印刷などの外部影響操作はキュレーション対象外。

---

## 2. 全体アーキテクチャ

```text
[Word / Excel / PowerPoint]
            │  ユーザーがホットキーを押す
            ▼
┌─────────────────────────────────────────────────────────────┐
│ office-bridge プロセス                                      │
│                                                             │
│  hotkey thread ──► BridgeCore ──► COM worker (STA 1本)      │
│  RegisterHotKey      対象レジストリ       pywin32 / COM     │
│  GetForegroundWindow target_id 振り分け   Office オブジェクト │
│                                                             │
│  FastAPI + Uvicorn on 127.0.0.1:8733                        │
└─────────────────────────────────────────────────────────────┘
            ▲
            │ HTTP + JSON
┌───────────┴───────────┐
│ office CLI (Typer)    │
└───────────▲───────────┘
            │ コマンド実行
┌───────────┴───────────┐
│ Configured agent +    │
│ Office Skill          │
└───────────────────────┘
```

### 2.1 プロセスとスレッド

| 実行単位 | 役割 | 備考 |
| --- | --- | --- |
| ブリッジプロセス | HTTP サーバー、ホットキー、対象保持、COM 委譲 | 常駐 |
| Uvicorn / asyncio | HTTP 受付 | FastAPI の同期エンドポイントはスレッドプールで実行 |
| hotkey スレッド | `RegisterHotKey` と Windows メッセージループ | `src/office_bridge/hotkey.py:129` |
| COM STA ワーカー | すべての COM 生成・参照・呼び出し・解放 | `src/office_bridge/com_worker.py:58` |
| CLI プロセス | HTTP クライアント、確認プロンプト、JSON 出力 | ブリッジとは別プロセス |
| エージェントプロセス | Skill を読み、CLI を実行 | ランチャーまたはホットキーから起動 |

COM は原則としてシングルスレッドアパートメント（STA）で扱う必要がある。Office の
オブジェクトを複数スレッドから直接触ると、マーシャリングや再入の問題が起きやすい。
そのため、COM オブジェクトの生成から破棄までを 1 本の STA スレッドに閉じ込め、
HTTP ハンドラーやホットキーはキュー経由で処理を依頼する設計にしている。

### 2.2 コンポーネント責務

| モジュール | 責務 |
| --- | --- |
| `src/office_bridge/server.py` | 起動引数、ログ、単一インスタンス、ホットキー起動、Uvicorn 起動、終了処理 |
| `src/office_bridge/config.py` | 環境変数と既定値、loopback 制約、ベース URL |
| `src/office_bridge/agent_config.py` | エージェントのプリセット、設定ファイル、Skill 配置、起動コマンド解決 |
| `src/office_bridge/hotkey.py` | ホットキー文字列の解析、`RegisterHotKey`、メッセージループ |
| `src/office_bridge/target.py` | HWND → Office 種別判定、COM ネイティブオブジェクト取得、文書解決 |
| `src/office_bridge/targets.py` | `TargetRecord` / `TargetRegistry`: 対象 ID、重複キー、寿命・監査（COM 非依存） |
| `src/office_bridge/com_worker.py` | STA ワーカー、要求キュー、対象レジストリと COM 参照保持、`target_id` 振り分け、操作実行、エラー分類、ログ |
| `src/office_bridge/com_expr.py` | `/com` 用のメンバー連鎖・リテラル解析 |
| `src/office_bridge/api.py` | FastAPI エンドポイント、Pydantic スキーマ、JSON エンベロープ |
| `src/office_bridge/errors.py` | エラーコード、HTTP ステータス、例外詳細 |
| `src/office_bridge/operations/registry.py` | 操作名のディスパッチ、ケイパビリティ定義 |
| `src/office_bridge/operations/excel.py` | Excel の読み取り・書き込み・書式操作 |
| `src/office_bridge/operations/word.py` | Word の選択範囲読み取り・テキスト挿入 |
| `src/office_bridge/operations/powerpoint.py` | PowerPoint のスライド・図形テキスト読み取り |
| `src/office_bridge/operations/com_ops.py` | 汎用 COM の get / call / set、ブロックリスト |
| `src/office_bridge/operations/common.py` | 入力検証、JSON 変換、上限、文書照合などの共通処理 |
| `src/office_cli/main.py` | Typer CLI、HTTP クライアント、書き込み確認、終了コード |
| `skills/office/SKILL.md` | エージェント向けの手順・禁止事項 |
| `scripts/*.ps1` | Skill 配置、ブリッジ起動、エージェント起動 |

### 2.3 起動シーケンス

1. `Settings.from_env` で引数と環境変数を解決する。loopback 以外のホストは拒否する。
2. ログを初期化する。コンソール（stderr）とローテーションファイルに出力する。
3. ポート番号を含む名前付きミューテックスで二重起動を防ぐ。
4. ポートが空いているか確認する。
5. `BridgeCore.start()` で COM ワーカースレッドを起動し、`CoInitialize()` する。
6. ホットキーを登録する。登録に失敗した場合は終了コード 4 で終了する。
7. FastAPI アプリを生成し、Uvicorn で待ち受ける。
8. 終了時にホットキー解除、COM ワーカー停止、ミューテックス解放を行う。

起動オプションは `--host`、`--port`、`--hotkey`、`--no-hotkey`、`--log-level`。
`--no-hotkey` はヘルスチェックやテスト用で、この場合は前景ウィンドウからの捕捉経路が
ない。テスト・診断に限り `POST /targets/capture`（HWND 指定）で登録できる。

---

## 3. HWND から COM 対象への解決

対象解決は最大の誤操作リスクである。`GetActiveObject()` だけでは、複数インスタンスや
複数文書があるときに別の文書を操作する可能性がある。実装では前景ウィンドウから
子ウィンドウをたどり、そのウィンドウに紐づくネイティブ COM オブジェクトを取得している。

### 3.1 ウィンドウ種別の判定

1. `GetForegroundWindow()` で前景 HWND を取得する。
2. `GetAncestor(hwnd, GA_ROOT)` でトップレベルウィンドウに正規化する。
3. ウィンドウクラス名を取得する。
4. `GetWindowThreadProcessId` で PID を取得する。
5. `OpenProcess` + `QueryFullProcessImageNameW` で実行ファイル名を取得する。

| アプリ | トップレベルクラス | 実行ファイル |
| --- | --- | --- |
| Excel | `XLMAIN` | `EXCEL.EXE` |
| Word | `OpusApp` | `WINWORD.EXE` |
| PowerPoint | `PPTFrameClass` | `POWERPNT.EXE` |

クラス名が一致しない場合、プロセスが Office 系なら「Office だが文書ウィンドウではない」
（ダイアログ、スタート画面など）として `NOT_OFFICE_WINDOW` を返す。Office 以外なら
同じコードで拒否する。前景が Office 以外のときに、現在の対象を暗黙に変更しない。

### 3.2 ネイティブオブジェクトの取得

Excel と Word は、文書用の子ウィンドウから `AccessibleObjectFromWindow` を使う。

- `OBJID_NATIVEOM`（`0xFFFFFFF0`）
- `IID_IDispatch`
- 返った `IDispatch` ポインタを `pythoncom.ObjectFromAddress` でラップ
- `win32com.client.dynamic.Dispatch` で遅延バインディングのオブジェクトにする

実装は `src/office_bridge/target.py:186`。遅延バインディングなので、Office の
バージョンごとにタイプライブラリ由来のラッパーを生成する必要がない。一方で、
メンバー名や引数の誤りは実行時に初めて分かる。

### 3.3 アプリ別の解決

| アプリ | 起点 | 文書の取り方 | 整合性チェック |
| --- | --- | --- | --- |
| Excel | 子ウィンドウ `EXCEL7` | `native.Application` → `native.ActiveSheet.Parent` | `Application.Hwnd == 前景トップ HWND` |
| Word | 子ウィンドウ `_WwG` | `native.Document` | `native.Hwnd == 前景トップ HWND` |
| PowerPoint | なし（タイトル / ROT） | タイトル照合で `Presentation` を選択 | スコア上位が一意であること |

Excel の `Application.Hwnd` 照合と Word の `Native.Hwnd` 照合により、前景ウィンドウと
解決した COM オブジェクトが同じインスタンスであることを確認する。食い違えば
`INSTANCE_MISMATCH` で失敗する。子ウィンドウが無い場合は `NO_DOCUMENT`。

PowerPoint は Excel / Word のような `OBJID_NATIVEOM` 経由の文書ウィンドウ解決が
安定しないため、次のヒューリスティックを使う（`src/office_bridge/target.py:503`）。

1. `GetActiveObject("PowerPoint.Application")` の `Presentations` を列挙する。
2. ウィンドウタイトルから `" - PowerPoint"` / `" - Microsoft PowerPoint"` を除去する。
3. プレゼンテーションの `Name` と `FullName` の basename をタイトルと比較する。
   - 完全一致: 100 点
   - タイトルが名前で始まる: 80 点
   - 名前に含まれる: 50 点
4. 候補が無ければ Running Object Table を走査し、`.ppt` / `.pptx` / `.pps` /
   `.ppsx` / `.pptm` / `.ppsm` のモニカーから `Presentation` を取得して同様に採点する。
5. 最高点が同点で複数ある場合は `RESOLVE_FAILED`。候補が無く、アクティブインスタンスの
   プレゼンテーションが 1 件だけならそれを採用する。
6. `presentation.Windows(1)` を文書ウィンドウとして保持する。

このため PowerPoint は、タイトルが未保存状態や互換モードなどで変わる場合、同名プレゼンが
複数ある場合に解決が弱い。曖昧なときは推測せずエラーにする方針である。

### 3.4 対象の保持と無効化

対象は `TargetRegistry`（`src/office_bridge/targets.py`）に複数保持する。各対象は
`TargetRecord` で、ブリッッジが発行する `target_id`（`t-` + 16 進 12 桁）、`AppKind`、
ウィンドウタイトル、PID、フルパス、メタデータ（シート名・スライド番号など）、
状態（`active` / `expired`）と時刻を持つ。COM オブジェクト参照そのものは
`BridgeCore._com_refs` に記録とは別に保持し、ワーカーのスレッドでのみ触る。

- ホットキーは前景 Office ウィンドウを **捕捉（capture）** して登録する。既存の対象は
  置き換わらず、捕捉時のみ `selected_target_id`（既定の対象）が更新される。
  同じウィンドウの再捕捉は新しい ID を発行せず、既存レコードの時刻を更新する。
- 重複判定は `(pid, full_name 小文字)` キー。フルパスが解けない未保存文書は
  `hwnd:{hwnd}` を使う。同名ファイルでも別プロセス（例: Word を `/x` で複数起動）なら
  別対象として登録される。
- 操作のたびに COM 参照の生存確認（`Application.Name` / `Document.Name`）を行う。
  失効すると理由（`closed` / `window_gone` / `com_reference_lost`）を記録して
  `expired` にする。失効レコードは監査のため保持し、50 件を超えない範囲で古いものから
  削除する（active は削除しない）。失効対象への操作は `TARGET_LOST` を返し、
  他の対象へフォールバックすることはない。
- 対象情報にはアプリ種別、文書名、フルパス、HWND、PID、ウィンドウタイトル、Excel の
  アクティブシート、PowerPoint のスライド番号などを含める。

---

## 4. COM 実行モデル

### 4.1 STA ワーカー

`ComWorker` は 1 本のスレッドで次を行う（`src/office_bridge/com_worker.py:61`）。

1. `pythoncom.CoInitialize()` で STA を初期化する。
2. キューから `(function, args, kwargs, future)` を取り出す。
3. `Future.set_running_or_notify_cancel()` で実行権を取る。
4. 関数を実行し、結果または例外を `Future` に設定する。
5. 次の要求を待つ間に `PeekMessage` / `TranslateMessage` / `DispatchMessage` で
   Windows メッセージをポンプする。
6. 終了時は `CoUninitialize()` する。

キュー待ちは 50ms タイムアウトで、待っている間もメッセージポンプを行う。COM 呼び出しの
途中で別の要求が割り込むことはなく、すべて直列化される。HTTP ハンドラーは `Future` の
完了を待つため、同時アクセスしても Office 側は 1 操作ずつ処理される。

### 4.2 タイムアウト

| 種別 | 既定値 | 用途 |
| --- | --- | --- |
| `request_timeout` | 30 秒 | 対象選択、context、読み取り、`/com get` |
| `write_timeout` | 120 秒 | 書き込み、`/com call|set` |
| CLI HTTP タイムアウト | 180 秒 | CLI 全体の HTTP 待ち |

`Future.result(timeout=...)` がタイムアウトすると `OFFICE_NO_RESPONSE`（HTTP 504）を返す。
ただし **COM 呼び出し自体は中断されない**。ワーカースレッドは呼び出しが返るまで
ブロックし続け、後続の要求はキューに溜まる。Office のダイアログ表示やビジー状態で
呼び出しが止まった場合、強制中断はできない。これは計画でも前提にしていた制約である。

### 4.3 エラー分類

`BridgeCore._execute`（`src/office_bridge/com_worker.py:567`）が操作を包み、
以下に分類する。

- `BridgeError`: そのまま返す。操作ログにコードを残す。操作系エラーは
  `details.target_id` を付す。
- 振り向け先の対象の生存確認に失敗: その対象を理由付きで `expired` にして
  `TARGET_LOST`。他の対象へは振り向けない。
- `pywintypes.com_error` で HRESULT がビジー系:
  `0x8001010A`、`0x800AC472`、`0x80010001` → `OFFICE_BUSY`（503）。
- その他の COM エラー: `COM_ERROR`（502）。HRESULT を 16 進で詳細に付ける。
- 想定外の例外: `INTERNAL_ERROR`（500）。

操作ログには `operation`、`target=Application/Document`、`id=target_id`、`ok`、
`duration_ms` を残す。文書内容やセル値は原則としてログに残さない。

### 4.4 対象のライフサイクル

- 対象はプロセス内のメモリに複数保持する（`TargetRegistry` + `_com_refs`）。
  ホットキーは捕捉であり、既存対象を置き換えない。既定の対象（`selected_target_id`）は
  捕捉時だけ更新される。
- `--target`/`target_id` を省略した要求は既定の対象へ送る（単一対象運用との後方互換）。
- 文書クローズや Office 終了時は、一覧・詳細・操作の生存確認で検出し、理由付きで
  `expired` にする。`GET /targets` はウィンドウ生存（`IsWindow`）の軽い確認のみで、
  フルな COM 生存確認は詳細取得と操作時に行う。
- 失効対象は再捕捉されない限り復活しない。操作は `TARGET_LOST`、ID が無ければ
  `TARGET_NOT_FOUND` を返し、他の対象へ自動的に切り替えることはない。
- ブリッジ再起動で対象は失われる。再度ホットキーを押す必要がある。
- `BridgeCore.stop()` はワーカー上でレジストリ・COM 参照・キャッシュをクリアしてから
  スレッドを止める。

注意点として、生存確認はあらゆる例外を「参照失効」とみなす。Office が一時的に
ビジーで `Name` 参照に失敗した場合も `com_reference_lost` になり得る。その場合は
同じウィンドウでホットキーを押し直せば同じ ID が再選択される。

---

## 5. HTTP API

FastAPI で実装し、`/docs` に OpenAPI UI を出す。Redoc は無効。すべての応答は
次のエンベロープに統一する。

```json
{
  "ok": true,
  "context": {"application": "Excel", "document": "Book1.xlsx"},
  "result": {}
}
```

エラー時:

```json
{
  "ok": false,
  "context": {"application": "Excel", "document": "Book1.xlsx"},
  "error": {
    "code": "CONFIRMATION_REQUIRED",
    "message": "Operation 'excel.range.write' modifies a document and requires confirm=true.",
    "details": {"operation": "excel.range.write"}
  }
}
```

`context` は対象のスナップショットで、`/health` やエラー応答でも返る。操作系エンドポイント
（`/inspect` `/invoke` `/set-property` `/com`）では **そのコマンドが実際に使った対象**
（`target_id` で振り向けた対象、省略時は既定の対象）の要約であり、それ以外の場所では
既定の対象（`selected_target_id`）の要約である。`/context` と `GET /targets/{id}` だけは
実際に COM を読んで最新化する。操作系の `result` は先頭に `target` ブロック
（`target_id` / `application` / `document` など）を持ち、データがどの文書から来たか・
どこへ書き込まれたかを明示する。操作系エラーは `error.details.target_id` に振り向け先を
載せる。

### 5.1 エンドポイント

| メソッド | パス | 用途 | 確認 |
| --- | --- | --- | --- |
| GET | `/health` | 稼働状態、バージョン、ホットキー、既定の対象、`target_count`、ワーカー生存 | 不要 |
| GET | `/context` | 既定の対象の最新情報 | 不要 |
| GET | `/capabilities` | 読み取り・書き込み・プロパティ・`/com` の定義、上限、ブロックリスト | 不要 |
| GET | `/targets` | 登録済み対象の一覧（失効判定は軽い確認のみ） | 不要 |
| POST | `/targets/capture` | HWND を指定して対象を登録（テスト用。通常はホットキー） | 不要 |
| GET | `/targets/{target_id}` | 対象の詳細（フルな COM 生存確認を行う） | 不要 |
| DELETE | `/targets/{target_id}` | 対象を登録から外す | 不要 |
| POST | `/inspect` | 読み取り | 不要 |
| POST | `/invoke` | キュレーション済み書き込み | `params.confirm` |
| POST | `/set-property` | 許可済みプロパティ変更 | `confirm` |
| POST | `/com` | 汎用 get / call / set | call / set は `confirm` |

Pydantic モデルは `extra="forbid"` で、未知のフィールドを拒否する。

### 5.2 リクエスト例

読み取り:

```json
{"kind": "excel.range", "params": {"address": "A1:B2", "include_formulas": true}}
```

キュレーション書き込み:

```json
{
  "operation": "excel.range.write",
  "params": {
    "address": "A1:B2",
    "values": [[1, 2], [3, 4]],
    "confirm": true,
    "expected_document": "Budget.xlsx"
  }
}
```

プロパティ変更:

```json
{"property": "excel.active_sheet", "value": "Sheet2", "confirm": true}
```

汎用 COM:

```json
{"mode": "get", "root": "window", "expression": "ActiveSheet.Range(\"A1\").Value2"}
```

`/invoke` は操作パラメータと `confirm` / `dry_run` / `expected_document` を、`/inspect` は
読み取りパラメータを `params` の中に置く。`/set-property` と `/com` は `confirm` /
`expected_document` をトップレベルに置く。API の形が操作種別で非対称なので、
クライアント実装時は注意が必要である。

操作系 4 エンドポイントはいずれもトップレベルの `target_id`（省略可）を受け付け、
指定されればその対象へ振り向ける。省略時は既定の対象（最後に捕捉された対象）を使い、
これは単一対象運用との後方互換のための挙動である。登録に無い ID や失効対象を指定すると
`TARGET_NOT_FOUND` / `TARGET_LOST` となり、他の対象へフォールバックすることはない。

### 5.3 例外ハンドリング

- `BridgeError`: 定義済みの HTTP ステータスとエンベロープで返す。
- `RequestValidationError`: 400 の `INVALID_REQUEST` に変換し、`loc` / `msg` / `type` を詳細に残す。
- `StarletteHTTPException`: 4xx は `INVALID_REQUEST`、5xx は `INTERNAL_ERROR` に変換する。
- その他の例外: 500 の `INTERNAL_ERROR`。スタックトレースはログに残す。

---

## 6. CLI

Typer でサブコマンドを構成し、HTTP クライアントは `httpx` を使う。結果は標準出力に
JSON（`ensure_ascii=False`、インデント 2）、診断は標準エラーに出す。`main()` では
標準出力を UTF-8 に再設定する。

### 6.1 主なコマンド

| コマンド | 内容 |
| --- | --- |
| `office health` | ブリッジの稼働確認 |
| `office target list` | 登録済み対象の一覧（`target_id`、状態、既定の対象） |
| `office target info <target_id>` | 対象の詳細（COM 生存確認付き） |
| `office target remove <target_id>` | 対象を登録から外す |
| `office context` | 既定の対象のアプリ・文書・パスなどの確認 |
| `office capabilities` | 操作一覧、上限、ブロックリストの確認 |
| `office inspect selection|document` | 共通の選択範囲・文書情報 |
| `office excel range read A1:D20 [--formulas]` | アクティブシートの範囲読み取り |
| `office excel range write A1:B2 --json ... [--dry-run] [--yes]` | 値の書き込み |
| `office excel range copy-format SRC DST` | 書式のみコピー |
| `office excel range set-format A1:B2 ...` | 明示的な書式設定 |
| `office excel workbook inspect` | シート名などのワークブック情報 |
| `office word selection read` | 選択テキストの読み取り |
| `office word selection insert-text ...` | 選択範囲へのテキスト挿入・置換 |
| `office word document read` | 段落数・単語数などの文書情報 |
| `office powerpoint slide inspect` | 現在スライドと図形テキスト |
| `office powerpoint document read` | スライド数・タイトル一覧 |
| `office set-property excel.active_sheet --value Sheet2` | 許可済みプロパティ変更 |
| `office com get|call|set ...` | 汎用オブジェクトモデル操作 |

全コマンドはサブコマンドの前に付ける大域オプション `--target <target_id>` を受け付ける
（環境変数 `OFFICE_BRIDGE_TARGET` でも指定可）。指定するとその対象へ振り向け、省略時は
既定の対象が使われる。複数文書の作業では常に `--target` を付けるのが運用規則である。

### 6.2 書き込みワークフロー

CLI の書き込みは次の流れを強制する。

1. `office target list` / `office target info` で書き込み先の対象を確認する（Skill の手順）。
2. `--dry-run` で範囲・配列・書式オプションを検証する。このとき `confirm` は送らない。
3. 対話端末なら、振り向け先の対象（`target_id` と文書名）と変更内容を stderr に表示し、
   `typer.confirm` で承認を取る。既定は No。対象が失効していれば承認を求めずエラーになる。
   対象が複数登録されている場合は、書き込み先に指定した ID を明示的に表示する。
4. `--yes` 指定時はプロンプトを省略し、`confirm: true` を送る。
5. `--expected-document` 指定時は、振り向け先の `document` / `full_name` と照合する。
6. サーバー側でも `confirm` と `expected_document` を再検証する。

対話端末でない場合に `--yes` を付けずに書き込もうとすると、エラー終了する。
これはエージェントが勝手に書き込むことを防ぐための CLI 側のガードである。ただし
後述するとおり、HTTP API 利用者は `confirm: true` を自分で送れるため、これは
暗号学的・認証的な承認ではない。

### 6.3 終了コード

| コード | 意味 |
| --- | --- |
| 0 | 成功（`ok: true`） |
| 1 | ブリッジがエラーを返した、またはユーザーが中止した |
| 2 | ブリッジに到達できない、またはコマンド引数が不正 |

`BridgeUnreachable` の場合は CLI がローカルで `BRIDGE_UNREACHABLE` エンベロープを
生成し、stderr にメッセージを出して終了コード 2 を返す。

---

## 7. キュレーション済み操作

### 7.1 読み取り

| 操作名 | 対象 | 返すもの | 主な制約 |
| --- | --- | --- | --- |
| `selection` | 3 アプリ共通 | Excel は選択範囲、Word は選択テキスト、PowerPoint は現在スライド | Excel は `RangeSelection` |
| `document` | 3 アプリ共通 | 文書情報 | アプリごとに項目が異なる |
| `excel.range` | Excel | `address`、`values`、必要なら `formulas`、複数領域は `areas` | アクティブシートのみ、セル数上限 |
| `excel.workbook` | Excel | 文書名、パス、シート名、`read_only` | ワークシートのみ |
| `word.selection` | Word | テキスト、`Start` / `End`、選択種別、文字数、切り詰めフラグ | 文字数上限 |
| `word.document` | Word | 段落数、単語数、文字数、`saved`、`read_only` | 取得失敗時は `null` |
| `powerpoint.slide` | PowerPoint | スライド番号、枚数、ID、レイアウト、タイトル、図形テキスト | 現在スライドのみ |
| `powerpoint.document` | PowerPoint | スライド数、ページサイズ、タイトル一覧 | 最大 1000 スライド、タイトルは各 500 文字 |

Excel の範囲読み取りは `Range.Value2` を使い、必要に応じて `Range.Formula` も読む。
COM が返す 1 次元・2 次元・スカラーの揺れを `_matrix` / `_rectangular` で
常に rows × columns の行列に正規化する。複数領域（`Areas.Count > 1`）の場合は
各領域を `areas` に分け、合計セル数で上限を判定する。

PowerPoint の図形テキストは、テキストフレームと表のセルを対象にする。グループ図形の
子要素を再帰的にたどる処理はしていない。スライド全体のテキスト量には上限を設け、
図形ごとに残り予算を消費する。タイトルは別枠で切り詰める。

### 7.2 書き込み

| 操作名 | 対象 | 動作 | 事前条件 |
| --- | --- | --- | --- |
| `excel.range.write` | Excel | 2 次元配列を `Value2` に代入 | 配列が矩形、範囲と寸法一致、値はリテラル、セル数上限、`confirm` |
| `excel.range.copy_format` | Excel | コピー元の書式をコピー先へ `PasteSpecial` | コピー元と先の寸法一致、セル数上限、`confirm` |
| `excel.range.set_format` | Excel | フォント・色・配置・罫線などを設定 | 1 つ以上のオプション、セル数上限、`confirm` |
| `word.selection.insert_text` | Word | 選択範囲に `TypeText` | 非空文字列、文字数上限、`confirm` |
| `excel.active_sheet` | Excel | シートをアクティブ化 | シート名が存在、`confirm` |

#### Excel 値書き込み

`write_range`（`src/office_bridge/operations/excel.py:259`）の順序は次のとおり。

1. `expected_document` の照合。
2. 2 次元配列の検証（非空、矩形、有限数、文字列・数値・真偽値・null のみ）。
3. セル数上限の確認。
4. `=` で始まる文字列（数式）を拒否。
5. 範囲を解決し、実際の行数・列数と配列の寸法を完全一致で比較。
6. dry-run でなければ `confirm` を要求。
7. 1×1 ならスカラー、それ以外はタプルのタプルを `Value2` に代入。

数式を書きたい場合はキュレーション操作では扱わず、`office com set` で
`...Formula` を設定する。シートを指定する引数はなく、常に `ActiveSheet` が対象。
別シートなら先に `excel.active_sheet` を変更するか、`/com` を使う。

#### 書式コピー

`copy_format`（`src/office_bridge/operations/excel.py:329`）は `source.Copy()` と
`dest.PasteSpecial(-4122)`（`xlPasteFormats`）を使う。値は変更しない。
処理後に `Application.CutCopyMode = False` でコピーモードを解除する。Excel の
コピー機構を使うため、ユーザーのクリップボード状態に影響し得る。

#### 書式設定

`set_format`（`src/office_bridge/operations/excel.py:391`）は、指定された
オプションだけを順に適用する。対応するのは以下。

- フォント名、フォントサイズ（0 より大きく 1000 以下）
- 太字、斜体、下線
- フォント色、塗りつぶし色、罫線色（`RRGGBB` または `#RRGGBB`）
- 表示形式（`NumberFormat`）
- 水平・垂直配置、折り返し
- 連続線の罫線

色は RGB で受け取り、Excel の BGR 表現に変換する。塗りつぶしは
`Interior.Pattern = xlPatternSolid` を設定してから色を入れる。
`Font.NameFarEast` は Excel の Font オブジェクトに設定可能なプロパティとして
存在しないため、指定された場合はエラーにし、`MS Gothic` / `Yu Gothic` / `Meiryo`
などの pan-CJK フォントを `font_name` に指定するよう案内する。

書式設定は「全オプションを検証してから適用」するが、適用は 1 プロパティずつであり
トランザクションではない。途中で COM エラーになると、それまでの変更は残る。
dry-run は「どのプロパティをどの値で設定するか」を返すだけで、Office 側の保護や
結合セルなどの最終的な可否までは保証しない。

#### Word テキスト挿入

`insert_text`（`src/office_bridge/operations/word.py:104`）は `Selection.Start` /
`End` を記録し、dry-run なら何もしない。実行時は `Selection.TypeText(text)` を呼ぶ。
既定の `replace=True` では選択範囲を置換する。`replace=False` では
`Selection.Collapse(0)`（選択範囲の先頭に折りたたむ）してから挿入する。
実行後の選択範囲の `Start` / `End` も返す。

---

## 8. 汎用 COM 操作 (`POST /com`)

キュレーション操作で足りないもの（行・列の挿入削除、図形・表・画像、ページ設定、
マクロ実行など）は、汎用オブジェクトモデル操作で行う。

### 8.1 式の文法

`src/office_bridge/com_expr.py` が次の文法を解析する。

```text
expression := segment ( "." segment )*
segment    := IDENT [ "(" [ value ( "," value )* ] ")" ]
value      := string | number | true | false | null | array | object
```

- 識別子は `[A-Za-z_][A-Za-z0-9_]*`。日本語のメンバー名は使えない。
- 文字列はシングルクォートまたはダブルクォート。エスケープは
  `\n`、`\t`、`\r`、`\0`、`\\`、`\'`、`\"`。
- 数値は整数・小数・指数表記。負数も可。
- 配列 `[...]` とオブジェクト `{"key": value}` は引数・設定値に使える。
- 添字構文 `[...]` はメンバーアクセスとしては使えない。`Cells(1, 1)` や
  `Item(...)` のようにメソッド／既定メンバー呼び出しを使う。
- 名前付き引数は未対応。位置引数のみ。
- 列挙値は数値で渡す。例: Excel の `xlInsertShiftDown = -4107`、
  PowerPoint の `msoShapeRectangle = 1`。

例:

```text
window.ActiveSheet.Range("A1").Value2
document.Paragraphs(2).Range.Text
window.ActiveSheet.Rows("2:4").Insert(-4107)
window.Selection.Font.Bold
```

### 8.2 ルート

| ルート | 実体 |
| --- | --- |
| `application` | `target.application`（Office Application） |
| `window` | `target.window`（Excel Window / Word Window / PowerPoint DocumentWindow） |
| `document` | `target.document`（Workbook / Document / Presentation） |

既定ルートは `window`。式の先頭が `application` / `window` / `document` の場合、
その名前をルートとして採用し、先頭セグメントを式から取り除く。ただし明示的な
`root` が `window` 以外で式の先頭と食い違う場合はエラーにする。
PowerPoint で `window` が無い場合（ウィンドウを持たないプレゼンテーション）は
`NO_DOCUMENT` を返すので、`application` か `document` を使う。

### 8.3 モード

| モード | 式の末尾 | 動作 | 確認 |
| --- | --- | --- | --- |
| `get` | プロパティ名（引数なし） | プロパティを読む | 不要 |
| `call` | メソッド呼び出し `name(args)` | メソッドを呼ぶ | `confirm`（dry-run を除く） |
| `set` | プロパティ名（引数なし） | `value` を代入する | `confirm`（dry-run を除く） |

- `get` は途中にメソッド呼び出しを含んでよい。最後だけプロパティである必要がある。
- `call` は最後が `name(args)` でなければならない。引数なしなら `Copy()` のように空括弧。
- `set` は `value` が必須で、`null` は受け付けない。CLI も `--value null` を拒否する。
- `get` の結果が文字列の場合、`max_text_chars` で切り詰め、`truncated` を返す。
- `call` の戻り値と `set` の設定値は JSON 化して返す。

### 8.4 dry-run

- `call` の dry-run は、対象オブジェクトとメソッドが存在し callable であることだけを
  確認する。引数の型や Office 側の状態までは検証しない。
- `set` の dry-run は、対象オブジェクトとプロパティが存在することを確認する。
  型の互換性や保護状態までは検証しない。
- 途中のプロパティ参照は dry-run でも実行される。

### 8.5 ブロックリスト

保存とクローズだけを意図的に禁止する（`src/office_bridge/operations/com_ops.py:32`）。

- メソッド呼び出しとプロパティ設定について、メンバー名が `save` / `close` で
  始まるもの、または `quit` / `exit` を拒否する。大文字小文字は区別しない。
- 式の途中にあるメソッド呼び出しも検査する。
- プロパティの読み取りはブロックしない。たとえば `document.Saved` の読み取りは許可。
- `OFFICE_BRIDGE_BLOCKED_CALLS` にカンマ区切りで追加の名前を指定できる。

拒否時は `OPERATION_BLOCKED`（HTTP 403）。

**重要な限界**: これはメンバー名ベースの防御であり、マクロなどの迂回路を完全には
防げない。`Application.Run("SaveAndClose")` のような呼び出しは名前が `Run` なので
ブロックリストを通る。Skill ではマクロ経由の回避を禁止しているが、サーバー側で
マクロの内容を検査しているわけではない。

### 8.6 結果の形

```json
{
  "mode": "get",
  "root": "window",
  "value": "A1",
  "truncated": false
}
```

```json
{
  "mode": "call",
  "root": "window",
  "dry_run": false,
  "executed": true,
  "method": "Insert",
  "arguments": [-4107],
  "result": null
}
```

```json
{
  "mode": "set",
  "root": "window",
  "dry_run": false,
  "applied": true,
  "property": "Bold",
  "value": true
}
```

`/capabilities` の `com.examples.set` は `window.Selection.Font.Bold = true` という
代入形で表示されるが、実際の `expression` は左辺の `window.Selection.Font.Bold` のみを
渡し、右辺は `value` に入れる。

---

## 9. 安全設計

### 9.1 loopback 限定

`Settings.from_env` は `127.0.0.1`、`localhost`、`::1` 以外のホストを拒否する
（`src/office_bridge/config.py:97`）。Uvicorn もそのホストで待ち受ける。
外部ネットワークからは到達できない前提である。CORS は設定していない。

### 9.2 確認と文書照合

- キュレーション書き込みと `/com call|set` は、dry-run を除き `confirm: true` を要求する。
- `expected_document` は任意。指定時は振り向け先の `document`、`full_name`、`full_name` の
  basename を小文字で比較する。一致しなければ `DOCUMENT_MISMATCH`。
- CLI は `--yes` が無い場合に対話端末を要求し、振り向け先の対象と変更内容を表示して承認を取る。
  指定した `target_id` が失効している場合は承認を求めずに拒否する。

ただし **`confirm` はサーバー側の人間承認ではない**。HTTP API を直接呼ぶローカル
プロセスは `confirm: true` を送れる。CLI のプロンプトは事故防止のための協調的な
ガードであり、認証・認可ではない。ブリッジを信頼できないプロセスと共有する環境では、
別の認証層が必要になる。

### 9.3 読み取り上限

| 設定 | 既定値 | 適用範囲 |
| --- | --- | --- |
| `max_cells` | 10,000 | Excel の範囲読み取り・書き込み・書式操作 |
| `max_text_chars` | 100,000 | Word 選択テキスト、PowerPoint のテキスト、`/com get` の文字列結果 |

上限を超えると `RANGE_TOO_LARGE` / `TEXT_TOO_LARGE`（HTTP 413）。

`/com get` の切り詰めは、最終結果が文字列の場合だけに適用される。リストや
オブジェクトを返すプロパティは、`jsonable` で文字列化された後は上限の対象外に
なり得る。大きなコレクションを読むときは、呼び出し側で範囲や件数を絞る必要がある。

### 9.4 保存・クローズの禁止

ブリッジは文書を保存しない。COM の `save*` / `close*` / `quit` / `exit` を
ブロックリストで拒否する。ユーザーが保存・クローズしたい場合は Office の UI で行う。
これは、エージェントの誤操作でファイルが上書きされたり、未保存の変更が失われたりする
ことを避けるための設計である。ブリッジ経由の変更は文書をダーティ状態にするが、
保存はしない。

### 9.5 ログ

- 出力先: `%LOCALAPPDATA%\office-bridge\logs\bridge.log`（`OFFICE_BRIDGE_LOG_DIR` で変更可）。
- ローテーション: 2 MB × 3 世代。
- コンソールにも stderr で出力する。
- 操作ログ: 操作名、対象アプリ／文書、`id=target_id`、成否、所要時間。
- 対象登録ログ: 捕捉・再選択・失効・削除を `id` 付きで記録。失効は理由も残す。
- 文書本文やセル値は原則として記録しない。

---

## 10. 設定・ログ・運用

### 10.1 環境変数

| 変数 | 既定値 | 意味 |
| --- | --- | --- |
| `OFFICE_BRIDGE_HOST` | `127.0.0.1` | 待ち受けホスト。loopback のみ |
| `OFFICE_BRIDGE_PORT` | `8733` | 待ち受けポート |
| `OFFICE_BRIDGE_HOTKEY` | `ctrl+alt+o` | グローバルホットキー |
| `OFFICE_BRIDGE_HOTKEY_ENABLED` | `1` | ホットキー有効化 |
| `OFFICE_BRIDGE_MAX_CELLS` | `10000` | Excel のセル数上限 |
| `OFFICE_BRIDGE_MAX_TEXT_CHARS` | `100000` | テキスト文字数上限 |
| `OFFICE_BRIDGE_REQUEST_TIMEOUT` | `30` | 読み取り系の COM 待ち秒数 |
| `OFFICE_BRIDGE_WRITE_TIMEOUT` | `120` | 書き込み系の COM 待ち秒数 |
| `OFFICE_BRIDGE_LOG_LEVEL` | `INFO` | ログレベル |
| `OFFICE_BRIDGE_LOG_DIR` | `%LOCALAPPDATA%\office-bridge\logs` | ログディレクトリ |
| `OFFICE_BRIDGE_BEEP` | `1` | ホットキー失敗時のビープ |
| `OFFICE_BRIDGE_LAUNCH_AGENT` | `1` | ホットキー時のエージェント自動起動 |
| `OFFICE_BRIDGE_AGENT` | 未設定 | 設定ファイルのエージェントを一時的に上書き |
| `OFFICE_BRIDGE_AGENT_CONFIG` | `%LOCALAPPDATA%\office-bridge\agent.json` | エージェント設定ファイルのパス |
| `OFFICE_BRIDGE_SKILL_SOURCE` | 未設定 | Skill 原本 `SKILL.md` の明示パス |
| `OFFICE_BRIDGE_BLOCKED_CALLS` | 空 | 追加のブロック名（カンマ区切り） |
| `OFFICE_BRIDGE_URL` | 未設定 | CLI が使うブリッジ URL。ホットキー起動時にも子プロセスへ渡す |
| `OFFICE_BRIDGE_CLI_TIMEOUT` | `180` | CLI の HTTP タイムアウト秒数 |
| `OFFICE_BRIDGE_TARGET` | 未設定 | CLI が既定で振り向ける `target_id`（`--target` の代わり） |

真偽値は `0` / `false` / `no` / `off` 以外を真として扱う。

### 10.2 単一インスタンスと終了コード

`server.py` は `Local\OfficeComBridge-<port>` という名前付きミューテックスを作成する。
既に存在すれば「別のブリッジが同じポートで動いている」として終了コード 3 を返す。
ミューテックス取得後にポートの空きも確認する。

| 終了コード | 意味 |
| --- | --- |
| 0 | 正常終了 |
| 1 | Uvicorn の SystemExit など |
| 2 | 設定エラー（loopback 以外など） |
| 3 | 二重起動、またはポート使用中 |
| 4 | ホットキー登録失敗 |

### 10.3 ホットキーとエージェント起動

`HotkeyListener` は専用スレッドで `RegisterHotKey(None, id, modifiers | MOD_NOREPEAT, vk)`
を呼び、`GetMessageW` ループで `WM_HOTKEY` を待つ。停止時は
`PostThreadMessageW(WM_QUIT)` でループを抜け、`UnregisterHotKey` する。

ホットキー受信時は次の処理を行う。

1. `GetForegroundWindow()` で前景 HWND を取得。
2. COM ワーカーにそのウィンドウの捕捉（capture）を依頼。
3. 成功すれば対象をレジストリに登録（既存なら再選択）、`selected_target_id` に設定し、
   ログに記録。エージェントは `OFFICE_BRIDGE_URL` 付きで起動する。
4. `OFFICE_BRIDGE_LAUNCH_AGENT` が有効なら、設定されたエージェントを新しい端末で起動。
5. 失敗時は警告ログとビープ。既存の対象は変更しない。

起動コマンドは `agent_config.load_agent_config()` が返す設定から解決する。既定は
OpenCode で、`configure.bat` または `office-agent set` で Claude Code / Codex /
Hermes Agent / 任意コマンドに変更できる。実行ファイルは PATH または絶対パスから探し、
`.ps1` は PowerShell、`.cmd` / `.bat` は `cmd.exe /d /c` 経由で起動する。
Windows Terminal (`wt`) があれば `wt -d <cwd> ...`、無ければ `CREATE_NEW_CONSOLE`
で起動する。子プロセスには `OFFICE_BRIDGE_URL` を渡し、`sys.executable` の
ディレクトリ（通常は venv の `Scripts`）を PATH の先頭に追加する。これにより新しい
端末で `office` CLI が見つけやすい。

ホットキーを押すたびに起動するため、既にエージェントが開いていても新しいウィンドウが
増える。自動起動を止めたい場合は `OFFICE_BRIDGE_LAUNCH_AGENT=0` にする。

### 10.4 PowerShell スクリプト

| スクリプト | 役割 |
| --- | --- |
| `scripts/install-skill.ps1` | `skills/office/SKILL.md` を設定されたエージェントの Skill ディレクトリへコピー。`-Global` でユーザー領域、`-Agent` / `-Scope` で一時上書き |
| `scripts/start-bridge.ps1` | venv を有効化してブリッジ起動。`-Background` でログを `.office-bridge/` に出す |
| `scripts/start-agent.ps1` | Skill 配置、ブリッジの health 待機（最大 20 秒）、設定されたエージェントの起動。終了時に自分が起動したブリッジを停止（`-KeepBridge` で維持） |

`start-agent.ps1` は `OFFICE_BRIDGE_URL` を設定し、`/health` の `ok` を確認する。
固定 sleep ではなく期限付きポーリング（400ms 間隔）で待つ。`-Agent` で設定を一時的に
上書きできる。`office-agent resolve --json` で起動コマンドと Skill 配置先を取得し、
実行ファイルが見つからない、Skill が配置されていないなどの場合は明示的に失敗する。

ルートのバッチファイルは、この PowerShell 群を意識せずに使うための入口である。

| バッチ | 動作 |
| --- | --- |
| `install.bat` | Python 3.11/3.12 を検出し、`.venv` 作成、`pip install -e ".[dev]"`、Skill 配置、import 検証を行う |
| `configure.bat` | 引数なしで対話設定。`configure.bat set hermes` のように `office-agent` へ引数を透過する |
| `start.bat` | `.venv` と `pywin32` を確認し、`python -m office_bridge.server` を前景で起動する。`%*` で `--no-hotkey` などを透過する |
| `start-agent.bat` | `.venv` を確認し、`scripts/start-agent.ps1` を実行してブリッジと設定されたエージェントを起動する。`-Agent` で一時上書き |

バッチ内のメッセージは Windows 既定のコードページで文字化けしないよう ASCII のみにしている。

### 10.5 エージェント設定

`office-agent` は、ホットキー起動と `start-agent.bat` が使うエージェントを管理する。
設定は `%LOCALAPPDATA%\office-bridge\agent.json` に保存する。

| プリセット | 起動コマンド | プロジェクト Skill | ユーザー Skill | 既定 scope |
| --- | --- | --- | --- | --- |
| `opencode` | `opencode` | `.opencode/skills/office` | `~/.config/opencode/skills/office` | project |
| `claude` | `claude` | `.claude/skills/office` | `~/.claude/skills/office` | project |
| `codex` | `codex` | `.agents/skills/office` | `~/.agents/skills/office` | project |
| `hermes` | `hermes -s office` | `.hermes/skills/office` | `%HERMES_HOME%\skills\office` | user |
| `custom` | `--command` | `--skill-dir` または無効 | `--skill-dir` | project |

Hermes はプロジェクト Skill の読み込みに `hermes skills trust` が必要なため、既定で
ユーザー領域へ配置し、`-s office` でプリロードする。設定ファイルの主なフィールドは
`agent`、`command`、`args`、`skill_scope`、`skill_dir`、`install_skill`、
`preload_skill`。`command` が null ならプリセットのコマンドを使う。`OFFICE_BRIDGE_AGENT`
は設定ファイルより優先され、`command` / `skill_scope` / `skill_dir` をリセットして
プリセットへ切り替える。

主な CLI:

```text
office-agent list
office-agent show [--json]
office-agent set hermes [--scope user] [--arg=--tui]
office-agent install-skill [--agent codex] [--scope project]
office-agent resolve --json
office-agent reset
```

Skill の原本は `skills/office/SKILL.md`。`OFFICE_BRIDGE_SKILL_SOURCE` で明示できる。
`install_skill` が false の場合は Skill 配置を行わない。

### 10.6 Office Skill

`skills/office/SKILL.md` は、エージェントに対して次を指示する。

- タスク開始時に `office context` で対象を確認する。
- `TARGET_NOT_SET` ならホットキー押下をユーザーに依頼する。
- 読み取りは最も狭い操作を選ぶ。
- 書き込み前に変更内容を説明し、ユーザーの承認を得る。
- `--dry-run` → `--yes` + `--expected-document` の順で実行する。
- 保存・クローズは行わない。ブロックリストを迂回しない。
- 結果の JSON を根拠に判断し、未確認の書き込みを完了と報告しない。

Skill は選択したエージェントのディレクトリ（例: `.opencode/skills/office/SKILL.md`、
`.claude/skills/office/SKILL.md`、`.agents/skills/office/SKILL.md`、
`%HERMES_HOME%\skills\office\SKILL.md`）にコピーして使う。原本は
`skills/office/SKILL.md`。

---

## 11. エラーコード

| コード | HTTP | 意味 |
| --- | --- | --- |
| `INVALID_REQUEST` | 400 | 入力不正、式の構文エラー、未知のメンバーなど |
| `UNSUPPORTED_OPERATION` | 400 | 未定義の操作名 |
| `UNSUPPORTED_APPLICATION` | 400 | 対象アプリでは使えない操作 |
| `RANGE_INVALID` | 400 | 範囲の解決・読み書きに失敗 |
| `SIZE_MISMATCH` | 400 | 範囲と配列、コピー元と先の寸法不一致 |
| `NOT_OFFICE_WINDOW` | 409 | 前景が Office 文書ウィンドウではない |
| `NO_DOCUMENT` | 409 | 文書ウィンドウ／文書が無い |
| `NO_ACTIVE_SLIDE` | 409 | PowerPoint がスライド表示ではない |
| `TARGET_NOT_SET` | 409 | 既定の対象が未登録 |
| `TARGET_NOT_FOUND` | 404 | 指定 `target_id` が登録に無い |
| `TARGET_LOST` | 409 | 文書クローズ／Office 終了などで対象が失効 |
| `DOCUMENT_MISMATCH` | 409 | `expected_document` と不一致 |
| `RESOLVE_FAILED` | 409 | COM 対象を解決できない、PowerPoint が曖昧 |
| `INSTANCE_MISMATCH` | 409 | 前景ウィンドウと COM インスタンスが不一致 |
| `CONFIRMATION_REQUIRED` | 409 | `confirm` が無い |
| `OPERATION_BLOCKED` | 403 | 保存・クローズ系の操作 |
| `RANGE_TOO_LARGE` | 413 | セル数上限超過 |
| `TEXT_TOO_LARGE` | 413 | 文字数上限超過 |
| `COM_ERROR` | 502 | COM 呼び出し失敗（HRESULT 付き） |
| `OFFICE_BUSY` | 503 | Office がビジー／呼び出し拒否 |
| `BRIDGE_NOT_RUNNING` | 503 | COM ワーカーが停止 |
| `OFFICE_NO_RESPONSE` | 504 | COM ワーカーの応答タイムアウト |
| `INTERNAL_ERROR` | 500 | 想定外の例外 |
| `BRIDGE_UNREACHABLE` | - | CLI がローカルで生成。HTTP には出ない |

`COM_ERROR` の `details` には `type`、`message`、可能なら `hresult`（`0x........`）が入る。
操作系のエラー（HTTP 500 未満）は `details.target_id` に振り向け先を対象 ID で残す。

---

## 12. テストと検証

### 12.1 自動テスト

`python -m pytest -q` で **228 件成功**（本稿執筆時点）。テストは実際の Office を
起動せず、フェイクの COM オブジェクトと FastAPI の `TestClient`、Typer の
`CliRunner`、`monkeypatch` を使う。

| ファイル | 主な検証内容 |
| --- | --- |
| `tests/test_agent_config.py` | プリセット、設定の保存・読込、環境変数／CLI 上書き、Skill 配置、起動コマンド解決 |
| `tests/test_targets.py` | `TargetRegistry` と `TargetRecord`: ID 採番、捕捉の追加/上書きでない性質、`(pid, full_name)` 重複排除、未保存文書の `hwnd:` キー、再選択で ID 不変、失効理由、失効レコードの保持上限 |
| `tests/test_api.py` | エンベロープ、health / context / capabilities / targets、`target_id` 振り分けと `TARGET_NOT_FOUND`、validation、`/com` の引数伝播 |
| `tests/test_cli.py` | サブコマンド、`target` コマンド、`--target` の送信、書き込み確認、`--yes`、dry-run、JSON 値、終了コード |
| `tests/test_common.py` | `jsonable`、行列検証、文字数切り詰め、文書名照合 |
| `tests/test_com_expr.py` | 式パーサ（文字列、エスケープ、配列、オブジェクト、エラー） |
| `tests/test_com_ops.py` | get / call / set、ブロックリスト、ルート、dry-run、未知メンバー |
| `tests/test_com_worker.py` | フェイク COM での BridgeCore 操作、`_resolve_target` のフォールバック無し、COM 例外時のワーカー継続、`OFFICE_BUSY` 判定 |
| `tests/test_excel_ops.py` | 範囲読み取り、複数領域、上限、書き込み、寸法不一致、数式拒否、書式コピー／設定 |
| `tests/test_word_ops.py` | 選択読み取り、挿入、文字数上限、確認 |
| `tests/test_powerpoint_ops.py` | スライド情報、図形テキスト、タイトル、切り詰め |
| `tests/test_registry.py` | 操作ディスパッチ、アプリ制約、ケイパビリティ |
| `tests/test_hotkey.py` | ホットキー文字列の解析 |

### 12.2 実機確認

`.office-bridge/bridge.stderr.log` には、実機でブリッジを起動し、ホットキーを登録し、
Excel ブックを対象として捕捉・切替した記録が残っている。これは自動テストではなく
手動確認の証跡である。Phase 1（複数対象）では追加で次の実機確認を行った
（テスト用ポート 8799 のブリッジで `target capture --hwnd` を使い、`office` CLI 経由で検証）。

- Word 2 文書（Alpha / Beta）+ Excel + PowerPoint の 4 対象を同時登録し、
  `--target` で比較読み取り → 要約の書き込みは Beta のみ、参照側は無変更。
- Excel 範囲読み取り → PowerPoint スライド本文への反映は Deck のみ、Data.xlsx は無変更。
- 別フォルダの同名 `Same.xlsx` 2 ブックが別 `target_id` として両方登録される
  （`full_name` 基準の重複排除）。
- 失効対象・未知 ID への操作は `TARGET_LOST` / `TARGET_NOT_FOUND` で、フォールバック無し。
- 書き込み先の `--expected-document` 不一致は `DOCUMENT_MISMATCH` で拒否。
- 26 万セルの読み取りは従来どおり `RANGE_TOO_LARGE`（`details.target_id` 付き）。

同一文書の複数ウィンドウ（Word の `NewWindow`）は IDispatch 経由でメソッドが見つからず、
`/com` で新しいウィンドウを作れない。このケースは捕捉の重複キー（`hwnd:{hwnd}` フォールバック
を含む）のユニットテストで担保している。

Word と PowerPoint の日常操作ログは、コードとフェイクテストはあるが、リポジトリ内に
実機ログは無い。

### 12.3 テストの空白

- `target.py` の HWND 解決、`server.py` の起動・ミューテックス、`com_worker.py` の
  STA スレッド、`HotkeyListener` の実登録は自動テストの対象外。
- 実際の Word / Excel / PowerPoint を使う統合テストは未整備。
- CI、lint（ruff 等）、型チェック（mypy 等）の設定は無い。
- テスト実行時に Starlette / anyio の非推奨警告が 1 件出るが、失敗ではない。

---

## 13. 既知の制約と注意点

### 13.1 COM とタイムアウト

- タイムアウトしても COM 呼び出しは中断できない。Office がダイアログ待ちやビジーだと、
  ワーカーが長時間ブロックし、後続要求がキューに溜まる。
- 応答しない書き込みを盲目的に再試行しない。まず `office context` で状態を確認する。
- `is_alive()` は例外を対象消失とみなすため、一時的なビジーが `TARGET_LOST` に
  見えることがある。再度ホットキーで対象を選び直す。

### 13.2 確認モデル

- `confirm: true` はクライアントの自己申告であり、サーバー側の人間承認ではない。
- loopback 上の任意のプロセスが API を呼べる。認証は無い。
- `expected_document` は任意。指定しない場合は文書照合を行わない。
- 信頼できないプロセスと同居する環境では、追加の認証・ACL・専用ユーザー分離が必要。

### 13.3 ブロックリストの限界

- 名前ベースの防御であり、マクロ（`Application.Run`）や他の間接呼び出しを完全には
  防げない。
- 保存・クローズを禁止しても、文書をダーティにすることはできる。保存はユーザーが行う。
- 追加のブロック名は環境変数で設定できるが、許可リストではなく拒否リストである。

### 13.4 操作の原子性

- `excel.range.set_format` はプロパティを順に適用する。途中で失敗すると部分適用が残る。
- `copy_format` は Excel のコピー／貼り付けを使うため、クリップボード状態に影響し得る。
- キュレーション書き込みにロールバックやトランザクションは無い。
- dry-run は構文・寸法・オプションを検証するが、シート保護、結合セル、Office の
  一時状態などによる失敗までは保証しない。
- `/com` の dry-run は構造チェックであり、引数型やプロパティ型の互換性は保証しない。

### 13.5 対象解決

- Excel / Word は子ウィンドウのクラス名（`EXCEL7`、`_WwG`）に依存する。Office の
  バージョンや UI の実装変更で影響を受ける可能性がある。
- PowerPoint はタイトルと ROT によるヒューリスティック。同名プレゼン、未保存タイトル、
  複数インスタンスでは解決に失敗することがある。曖昧な場合は推測せずエラーにする。
- 複数対象をレジストリで保持する。同名ファイルでもパスが違えば別対象、同一ウィンドウ
  の再捕捉は同じ ID に寄約される。同一文書の複数ウィンドウ（Word の新ウィンドウ等）は
  同一対象として扱われる。
- 通常の対象捕捉はホットキー経由。`--no-hotkey` でも `POST /targets/capture`（HWND 指定、
  テスト・診断用）で登録できるが、運用としてはホットキーを正とする。
- ホットキー起動のエージェントはブリッジプロセスのカレントディレクトリで開く。

### 13.6 API / CLI の細かな不整合

- `excel.range.set_format` はバックエンドと API で `underline` を受け付けるが、
  CLI には `--underline` オプションが無い。CLI から下線を設定するには `/com` を使う。
- `/capabilities` の `com.examples.set` は代入形で表示されるが、実際の API では
  左辺を `expression`、右辺を `value` に分ける。
- `/invoke` は `confirm` / `dry_run` / `expected_document` を `params` 内に、`/inspect` は
  読み取りパラメータを `params` 内に、`/set-property` / `/com` はトップレベルに置く。
- `/com get` の文字数上限は最終結果が文字列のときだけ適用される。

### 13.7 書き込み範囲の制限

- Excel のキュレーション書き込みは常にアクティブシートが対象。シート名を直接指定できない。
- Excel の値書き込みはリテラルのみ。数式文字列は拒否する。数式は `/com` で `Formula` を設定する。
- Word の書き込みは選択範囲へのテキスト挿入のみ。書式や表の操作は `/com` を使う。
- PowerPoint の書き込みはキュレーションされておらず、`/com` を使う。
- 保存・クローズは提供しない。

### 13.8 読み取り範囲

- PowerPoint の図形テキストはテキストフレームと表のセルが対象。グループ図形の子要素は
  再帰的にたどらない。
- `powerpoint.document` は最大 1000 スライド、各タイトル最大 500 文字。
- `word.document` の段落数・単語数・文字数は取得失敗時に `null` になる。
- Excel の `excel.workbook` はワークシートのみで、チャートシートは列挙しない。

---

## 14. ファイル構成

```text
office_bridge/
  pyproject.toml                  # パッケージ定義、依存、entry points、pytest 設定
  README.md                       # 利用者向け概要
  office_com_bridge_plan.md       # 当初計画
  install.bat                     # venv 作成、依存インストール、Skill 配置
  configure.bat                   # エージェント選択（office-agent への入口）
  start.bat                       # ブリッジ起動（サーバー引数を透過）
  start-agent.bat                 # ブリッジ + 設定されたエージェント起動
  docs/
    implementation.md             # 本書
  src/
    office_bridge/
      __init__.py                 # __version__ = "0.1.0"
      __main__.py                 # python -m office_bridge
      config.py                   # Settings、環境変数、loopback 制約
      agent_config.py             # エージェント設定、プリセット、Skill 配置、起動解決
      errors.py                   # ErrorCode、BridgeError、エラー生成
      server.py                   # 起動、ログ、ミューテックス、ホットキー、Uvicorn
      api.py                      # FastAPI アプリ、Pydantic モデル、エンベロープ
      hotkey.py                   # RegisterHotKey、メッセージループ
      target.py                   # HWND 判定、OBJID_NATIVEOM、文書解決
      targets.py                  # TargetRecord / TargetRegistry（COM 非依存）
      com_worker.py               # STA ワーカー、BridgeCore、操作実行
      com_expr.py                 # /com 式パーサ
      operations/
        __init__.py
        common.py                 # 検証、JSON 変換、上限、文書照合
        registry.py               # 操作ディスパッチ、ケイパビリティ
        excel.py                  # Excel 読み書き・書式
        word.py                   # Word 読み取り・挿入
        powerpoint.py             # PowerPoint 読み取り
        com_ops.py                # 汎用 COM、ブロックリスト
    office_cli/
      __init__.py
      __main__.py
      main.py                     # Typer CLI、HTTP クライアント
  skills/
    office/
      SKILL.md                    # Skill 原本
  .opencode/
    skills/office/SKILL.md        # 配置先の一例（エージェント別に変わる）
  scripts/
    install-skill.ps1
    start-bridge.ps1
    start-agent.ps1
  tests/
    conftest.py
    test_agent_config.py
    test_api.py
    test_cli.py
    test_common.py
    test_com_expr.py
    test_com_ops.py
    test_excel_ops.py
    test_hotkey.py
    test_powerpoint_ops.py
    test_registry.py
    test_targets.py
    test_word_ops.py
```

`pyproject.toml` の entry points:

- `office = office_cli.main:main`
- `office-bridge = office_bridge.server:main`
- `office-agent = office_bridge.agent_config:main`

依存関係は `pywin32`（Windows のみ）、`fastapi`、`uvicorn`、`typer`、`pydantic`、
`httpx`、`rich`。開発用に `pytest`、`pytest-cov`。Python は `>=3.11,<3.13`。

---

## 15. 拡張候補

現状の設計を壊さずに追加できるもの。対象選択 API と複数対象の保持は Phase 1 で
実装済み（§3.4、§5.1）。

1. **認証付き loopback API**
   ランダムトークンを起動時に生成し、CLI と子プロセスに渡す。`confirm` の
   協調性を強化できる。
2. **実機統合テスト**
   Office を起動した状態で、Excel / Word / PowerPoint の読み書きを自動検証する。
   テスト用文書を用意し、保存・クローズはテスト側で行う。
3. **配布用パッケージ**
   PyInstaller 等で `office-bridge` と `office` を単体実行ファイル化する。
4. **読み取りサイズ制御の強化**
   `/com get` のコレクション結果にも上限を設ける。ページングや件数指定を追加する。
5. **PowerPoint 対象解決の改善**
   ウィンドウハンドルとプレゼンテーションの対応をより直接的に取る。
6. **監査ログ**
   操作名、対象、変更内容のハッシュ、承認者などを構造化して記録する。
7. **CLI の不足オプション**
   `--underline` など、バックエンドが受け付ける書式オプションを CLI に追加する。
8. **ロールバック／アンドゥ連携**
    書式設定の部分適用を避けるため、変更前の値を取得して復元する仕組みを検討する。
9. **Office 以外アプリの対象化（roadmap Phase 2）**
   Outlook やブラウザ・PDF など Office 以外の対象種別を追加する。`AppKind` と
   キャプチャ経路の拡張で対応できる形にしているが、Phase 1 では意図的に手付けていない。

---

## 付録: 代表的な処理フロー

### A. ホットキーで対象を捕捉する

```text
ユーザー: Ctrl+Alt+O
  → hotkey thread: GetForegroundWindow()
  → BridgeCore.capture_foreground()
  → COM worker: _capture_hwnd(hwnd)
      → resolve_target(hwnd)
          → トップレベル HWND / クラス / PID / EXE
          → Excel: EXCEL7 → AccessibleObjectFromWindow → Application / ActiveSheet.Parent
          → Word:  _WwG   → AccessibleObjectFromWindow → Document
          → PPT:   タイトル / ROT → Presentation → Windows(1)
          → Application.Hwnd などの整合性チェック
      → target.refresh_info()
      → registry に捕捉（新規 ID か再選択かを重複キーで判定）
      → _com_refs に COM 参照を保存、selected_target_id に設定、summary をキャッシュ
  → 成功: ログ、必要なら設定されたエージェントを起動
  → 失敗: 警告ログ、ビープ、既存の対象は変更しない
```

### B. Excel の範囲を読む

```text
office excel range read A1:B2 --formulas
  → CLI: GET /context はせず、POST /inspect
  → FastAPI: InspectRequest(kind="excel.range", params={...})
  → BridgeCore.inspect()
  → COM worker: registry.inspect()
      → excel.read_range()
          → target.window.ActiveSheet.Range("A1:B2")
          → Rows.Count × Columns.Count、上限チェック
          → Value2、必要なら Formula
          → 行列を正規化
  → JSON エンベロープ
  → CLI: stdout に JSON、exit 0
```

### C. 汎用 COM で行を挿入する

```text
office com call 'window.ActiveSheet.Rows("2:4").Insert(-4107)' --dry-run
  → 式を com_expr.parse_expression でセグメント化
  → CLI: dry-run なので confirm を送らない
  → POST /com {mode:"call", root:"window", expression:..., dry_run:true}
  → com_ops.execute()
      → 式の末尾が call か検証
      → root / segments 解決
      → ブロックリスト検査（Insert は許可）
      → expected_document 検査
      → dry-run: 対象メソッドが callable か確認して返す
  → 実際に実行する場合は --yes で confirm:true
      → COM worker 上でメソッドを呼ぶ
      → 結果を JSON 化
```

### D. 書き込み時の確認

```text
office excel range write A1:B2 --json '[[1,2],[3,4]]'
  → CLI: --yes 無し、--dry-run 無し
  → 対話端末か確認（非対話ならエラー）
  → GET /context で対象を取得
  → expected_document の事前照合
  → 対象と変更内容を stderr に表示
  → typer.confirm（既定 No）
  → 承認時のみ params.confirm = true
  → POST /invoke
  → サーバー側でも confirm と expected_document を再検証
  → Excel の寸法一致・数式拒否・セル数上限を検証
  → Value2 に代入
  → JSON で結果を返す（保存はしない）
```
