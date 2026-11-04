# Phase 1 実装計画：Office の複数ドキュメント対応

対象: `office-com-bridge` 現行実装 → AppBridge Phase 1
前提: `AppBridge_roadmap_v1.md` の決定事項（Phase 1 は Office 3 アプリのみ、過剰な汎用化をしない）に従う。

---

## 0. スコープ

**含める**

- 対象レジストリ（複数 Word / Excel / PowerPoint 文書の登録・一覧・状態管理）
- `target_id` による操作ルーティング（API / CLI 全域）
- ホットキーの「置き換え」→「登録・再選択」変更
- 複数文書用 Skill（Word 比較、Excel→PowerPoint）
- 安全失敗（失効対象への操作はエラー、フォールバックなし）

**含めない**（Phase 2 まで据える）

- Office 以外のアプリ、アダプター機構、共通ランタイムへの切り出し
- 文書一覧ウィンドウ（GUI）
- CLI 名・パッケージ名・リポジトリ名の AppBridge への変更

---

## 1. 現状実装との差分

| Phase 1 要求 | 現状の実装位置 | 変更内容 |
|---|---|---|
| 複数対象の保持 | `BridgeCore._target`（単一の `ResolvedTarget`、`com_worker.py`） | 対象レジストリへ置換 |
| 安定した `target_id` | なし（「現在の対象」のみ） | ブリッジ発行の ID を新設 |
| ホットキー＝登録・再選択 | `select_foreground` → `_select_hwnd` が常に対象を**置き換え**（`com_worker.py`、`server.py` の `on_hotkey`） | capture 意味（追加 or 再選択 + 選択状態更新）へ拡張 |
| 操作ごとの対象指定 | `inspect` / `invoke` / `set_property` / `com` はすべて `_require_target()` で単一対象を使用 | 各要求に `target_id` を追加 |
| 対象一覧・詳細 API | `GET /context`（現在対象 1 件）のみ | `/targets` 系エンドポイント新設 |
| 結果への対象識別子 | エンベロープの `context` は「最後の対象」の要約のみ | 対象 ID・文書識別情報を付与 |
| 失効処理 | 単一対象を `_clear_target` で破棄し `TARGET_LOST` | レコードを `expired` として保持、選択対象でも自動フォールバックなし |
| CLI の対象指定 | `--expected-document` による事後照合のみ（`main.py`） | `target` サブコマンドと `--target` オプション |

既存で再利用できるもの：`ResolvedTarget` と HWND→COM 解決（`target.py`）、STA ワーカーとエラー翻訳（`ComWorker`）、`expected_document` 照合（`common.py`）、確認・dry-run・ログ・読み取り上限。

---

## 2. 設計

### 2.1 対象モデル

新規モジュール `src/office_bridge/targets.py`（**COM を import しない純 Python**）。Phase 2 の共通化を見据え、Office 固有情報は `app_metadata` に分離するが、アダプター機構は作らない。

```python
class TargetState(str, Enum):
    ACTIVE = "active"      # 有効
    EXPIRED = "expired"    # 消失・参照無効（expired_reason を伴う）

@dataclass
class TargetRecord:
    target_id: str          # "t-" + 12桁16進（uuid4 から衝突チェック付きで発行）
    app_type: str           # "word" | "excel" | "powerpoint"（小文字固定）
    display_name: str       # 文書ファイル名
    state: TargetState
    expired_reason: str | None   # "closed" | "window_gone" | "com_reference_lost"
    created_at: str         # ISO 8601（UTC）
    last_selected_at: str | None
    last_refreshed_at: str | None
    dedupe_key: str         # 重複登録検出キー（2.3）
    metadata: dict          # 共通: display_name, full_name, path（取得できた場合）
    app_metadata: dict      # Office 固有: hwnd, pid, window_title, sheet / slide_count 等
```

- **COM オブジェクト参照は `TargetRecord` に持たせない。** ワーカー側で `dict[target_id, ResolvedTarget]` を別途保持し、公開レコードから参照を露出しない。
- `target_id` はウィンドウハンドル・ファイル名から導出しない（ roadmap §12.4 準拠）。
- 状態は roadmap の「有効 / 消失 / 無効」を `active` / `expired` + `expired_reason` に統合する。Phase 1 で区別すべき復旧可能性がないため、3 状態は持たない。

### 2.2 対象レジストリ

同モジュールに `TargetRegistry` を置く。

- 内部は `dict[target_id, TargetRecord]` + `last_selected_id`。挿入順を保持。
- API: `add_or_touch(dedupe_key, ...) -> TargetRecord`、`get`、`list`、`remove`、`set_selected`、`purge_expired(keep=50)`。
- **アクセスはすべて STA ワーカーのスレッド内**（`BridgeCore` の既存パターンどおり `worker.call` 経由）。HTTP スレッドはレジストリインスタンスに直接触れない。スレッドセーフはワーカー直列実行に依存し、ロックは追加しない。
- 一覧用の軽量スナップショット（`cached_targets`）を `cached_summary` と同じロック付きキャッシュで持ち、`/health` がワーカーを待たずに応答できるようにする。

### 2.3 capture と重複登録

ホットキー（および `POST /targets/capture`）の流れ：

1. 前景 HWND を `resolve_target()` で解決（現状のまま）。
2. 重複判定キーを計算：
   - `full_name` が非空 → `f"{pid}:{full_name.lower()}"`
   - 未保存文書（`full_name` 空）→ `f"hwnd:{top_hwnd}"`
3. キーに一致する **active** レコードがあれば重複登録せず、そのレコードを再選択し、`app_metadata`（hwnd・ウィンドウタイトル・シート名等）を更新。
4. なければ新 `target_id` を発行して登録し、選択状態に設定。
5. 旧対象は消さない（選択状態が変わるだけ）。

同名ファイル・別フォルダは `full_name` が異なるため別対象。同一ファイルの再作成（閉じて再度開く）は旧レコードが expired になるため新対象として登録される。

### 2.4 状態管理と失効

- **操作直前**：`_require_target` 相当が対象解決時に `is_alive()`（COM 応答確認）と `win32gui.IsWindow(hwnd)` を実行し、失効を検出したら `expired(com_reference_lost | window_gone)` にして `TARGET_LOST` を返す。**別対象へ切り替えない。**
- **一覧**（`GET /targets`）：キャッシュ済み状態 + `IsWindow` の安価な事前チェックのみ表示。詳細（`GET /targets/{id}`）と操作直前で完全な COM 検証を行う。対象数が多くても一覧が COM 往復で遅くならないため。
- expired レコードは監査のため一覧に残る（`state=expired`, `expired_reason`）。`office target remove` で手動削除、または active/expired 合計 50 件超過時に古い expired から自動除去。

### 2.5 操作ルーティングと互換

- `BridgeCore._execute` は要求パラメータの `target_id` を受け取る：
  - `target_id` あり → レコード検索。なければ `TARGET_NOT_FOUND`、expired なら `TARGET_LOST`（details に `target_id`・理由・文書名）。
  - `target_id` なし → `last_selected_id` の対象（従来互換）。選択対象が expired なら `TARGET_LOST`。
  - 対象未登録時は従来どおり `TARGET_NOT_SET`。
- アプリ種別と操作の整合は既存の `_require_app` / `UNSUPPORTED_APPLICATION` をそのまま使う（Word 対象に `excel.range.read` を送れば明確なエラー）。
- `expected_document` 照合は維持し、明示 `target_id` の追加検査として機能させる。
- **複数対象 Skill は `target_id` 明示を必須**（Skill 側に規定）。ブリッジ側の既定（last selected）は単一対象の従来利用・相互運用専用。

### 2.6 HTTP API

新規エンドポイント（`api.py`）：

| メソッド | パス | 内容 |
|---|---|---|
| POST | `/targets/capture` | 前景ウィンドウを登録・再選択（ホットキーと同じ処理。テスト・CLI 用） |
| GET | `/targets` | `{targets: [...], selected_target_id}` |
| GET | `/targets/{target_id}` | レコード詳細（完全な鮮度確認付き） |
| DELETE | `/targets/{target_id}` | 登録解除（COM 参照も解放） |

既存要求モデルに `target_id: str | None = None` を追加（`/inspect`、`/invoke`、`/set-property`、`/com`）。

エンベロープの変更（後方互換のため追加のみ）：

- 成功応答の `context` を、その操作が対象とした文書の要約（`target_id` 含む）に変更。未指定時は従来どおり last-selected。
- 読み取り `result` には `target_id`・`document`・`full_name` を必ず含める（ roadmap §5.4）。実装上は共通ラッパで付与し、各 operation を触らない。
- 新エラーコード：`TARGET_NOT_FOUND`（404）。`TARGET_LOST` の details に `target_id` / `reason` を追加。

### 2.7 CLI

```text
office target capture                 # 前景ウィンドウを登録（ホットキーの CLI 版）
office target list                    # 登録対象の一覧（選択状態・状態・失効理由を表示）
office target info <target-id>        # 詳細
office target remove <target-id>      # 登録解除

office --target t-ab12cd34ef56 excel range read A1:D20
office --target t-99aa88bb7766 excel range write A1:B2 --json ... --yes
```

- グローバル `--target` オプションを `configure` コールバックに追加し、`ClientOptions` を経由して全操作のペイロードへ入れる。省略時は従来動作。
- 確認プロンプト（`_confirm_write`）は `context` 参照をやめ、**解決される対象の情報**（`target_id` / アプリ / 文書名）を表示する。複数対象環境で誤認しにくいため、`--target` なしの書き込み時は選択対象名に加え「明示指定なし」と警告表示。
- `health` には登録対象数を含める。

### 2.8 ホットキー

- `server.py` の `on_hotkey` は `core.capture_foreground()`（capture 意味）を呼ぶ。-beep・`launch_agent` 挙動は現状維持。
- 登録成功時は対象数と文書名をログに残す（`hotkey target: ...` を拡張）。

---

## 3. 実装マイルストーン

### M1: 対象レジストリと capture UX

- 新規 `src/office_bridge/targets.py`（`TargetRecord` / `TargetRegistry`）
- `com_worker.py` の `BridgeCore` をレジストリ所有に変更（`_target` → レジストリ + `_com_refs`）
- ホットキーを「登録・再選択」へ変更（`server.py`）
- `/targets/capture`・`/targets`・`/targets/{id}`（GET/DELETE）+ `/health` 拡張
- CLI `target` サブコマンド
- テスト：`tests/test_targets.py`（レジストリ純粋ロジック）、`test_api.py` / `test_cli.py` 追加、capture の重複登録テスト

**完了条件**：ホットキー 2 回で 2 文書が登録され、`office target list` に両方が見える。同一文書の再捕捉で ID が増えない。

### M2: `target_id` によるルーティング

- 各要求モデル・`BridgeCore.inspect/invoke/set_property/com` に `target_id` を透過
- `_execute` の対象解決をレジストリ基準に変更（失効・誤指定は安全に失敗、フォールバックなし）
- 読み取り結果への `target_id` / 文書識別子の共通付与
- エラーコード `TARGET_NOT_FOUND` 追加（`errors.py`）
- テスト：不正 ID・失効対象・別アプリ操作のエラー応答

**完了条件**：`--target A excel range read` → `--target B word document read` が交互に正しく動作し、失効対象への操作は `TARGET_LOST`。

### M3: 既存操作の移行完了と互換維持

- CLI `--target` 全局オプション、確認プロンプトの対象情報表示
- `/context`・`TARGET_NOT_SET` の互換動作を維持したまま回帰テスト追加
- STA ワーカーで複数対象を直列操作するテスト（fake core / fake ResolvedTarget でよい）
- COM エラー後のワーカー継続性テスト（既存 test 拡張）

**完了条件**：既存の読み書き・確認手順・テストが全て残り、`office context` 単一対象フローも従来どおり動く。

### M4: 複数文書用 Skill

- `skills/office/SKILL.md`（と `.opencode/skills/office/SKILL.md`）に「複数文書の扱い」節を追加：
  - タスク開始時に `office target list` で最新一覧を取得（起動時の一覧を信じない）
  - 全操作で `--target` 明示、書き込み先は必ず明示＋確認
- ワークフロー 1：2 つの Word 文書比較（別 ID で読み取り→エージェントが比較→必要なら書き込み先を明示）
- ワークフロー 2：Excel 範囲読み取り → PowerPoint への反映
- Skill の手順中に「参照元・書き込み先の取り違え検証」ステップを含める

**完了条件**：ロードマップ §7 の Skill 完了条件 2 件が実機で動作。

### M5: 安全性・統合テストとドキュメント

- 手動統合チェックリスト（実機）：
  - [ ] 文書閉じ後の全操作が安全に失敗
  - [ ] 同名ファイル（別フォルダ）・同一文書の複数ウィンドウ
  - [ ] 対象 ID 誤指定・存在しない ID
  - [ ] 書き込み確認・dry-run・`expected_document`
  - [ ] 長い読み取り結果とタイムアウト
  - [ ] COM エラー後のワーカー継続性
- `docs/implementation.md` に「対象レジストリ」章を追加、README 更新

---

## 4. テスト戦略

- **ユニット**（COM なし）：`TargetRegistry` の ID 発行・重複判定・失効遷移・選択・purge。`BridgeCore` は fake `ComWorker` / fake `ResolvedTarget` で capture の追加倍（置き換えにならないこと）・失効時のフォールバックなしを検証。
- **API**：既存 `FakeCore` パターンを拡張し、`/targets` 系と `target_id` 透過、`TARGET_NOT_FOUND` を検証。
- **CLI**：`test_cli.py` の既存偽クライアントで `target` コマンドと `--target` ペイロードを検証。
- **実機統合**：M5 のチェックリスト。自動化的に難しい失効・ウィンドウ対応はチェックリスト運用とする（現行 `docs/implementation.md` §12.2 の方針を継承）。

---

## 5. リスクと設計上の論点

| 項目 | 方針 |
|---|---|
| PowerPoint のウィンドウ→プレゼン対応はタイトル照合の経験則（`target.py`） | capture 時の 1 回だけなので許容。解決失敗時は現行どおり明確なエラー。レジストリ登録済み対象は COM 参照を保持するため再解決が不要 |
| 未保存文書（`Document1` 等） | `full_name` 空 → hwnd ベースの重複判定。初回保存後に refresh で `full_name` を更新 |
| 一覧の鮮度確認コスト | 一覧はキャッシュ + `IsWindow`、完全 COM 検証は詳細・操作直前（§2.4） |
| 同一ファイルの複数プロセス保持 | 現行同様サポート外。発生時は capture 失敗または最初の解決対象となり、一覧で確認可能 |
| `context` の意味変化 | 「その操作の対象」に変わるため、Skill 記述を M4 で追従。後方互換のためキーの削除はしない |
| ワークフロー中のセッション開始時一覧 | §11.4 準拠：Skill は操作直前に `target list` を取得する手順を義務化 |

## 6. Phase 1 完了条件との対応

| roadmap §7 のチェックボックス | マイルストーン |
|---|---|
| 複数文書の同時登録 | M1 |
| 一意な `target_id` と一覧・詳細 | M1 |
| すべての操作での対象明示 | M2 / M3 |
| 閉じられた文書への操作が安全に失敗 | M2 / M5 |
| 参照元・書き込み先を分けたワークフロー | M4 |
| Word 2 文書比較 Skill | M4 |
| Excel→PowerPoint Skill | M4 |
| 既存機能・確認手順・テストの維持 | M3 / M5 |

## 7. 進め方の注意

- 作業ツリーには未コミットの変更（`com_ops.py` のシート配置引数解決など）がある。Phase 1 のコミットと混ざらないよう、先に独立コミットとして分離してから着手する。
- モジュール名・CLI 名は Phase 1 では `office_bridge` / `office` のまま（ roadmap §11 準拠）。リネームは Phase 2 開始時に別途。
- 各マイルストーンの完了ごとにコミットし、M1→M2→M3 の順を守る（ルーティングを M1 に混ぜない）。
