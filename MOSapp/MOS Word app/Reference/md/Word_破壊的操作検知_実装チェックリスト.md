# Word 破壊的操作検知 — 実装チェックリスト

本チェックリストは `Reference/md/Word_破壊的操作検知_実装案.md` を実装に落とし込むための作業用です。  
チェックは `[ ]` -> `[x]` で更新してください。

---

## 0. 事前確認（着手前）

- [ ] `Word_破壊的操作検知_実装案.md` の「本書で固定する前提」をチーム内で合意する
- [ ] `TaskStart` 記録主体を「アプリ側」で確定（VSTO 側重複記録をしない）
- [ ] `attemptNo` 運用ルールを確定（通常 `0`、結果画面再採点 `>=1`）
- [ ] 対象ファイルの現状バックアップ方針を決める（ブランチ/小分けコミット）
- [ ] 実装の受け入れ順（フェーズ 0 -> 1 -> 2 -> 3）を確定

---

## 1. フェーズ 0（配線のみ）

### 1-1. パス/共有ファイル API（Libraries/LogReader.cs）

- [ ] `GetCurrentTaskFilePath()` を追加（`mos_word_current_task.txt` のパス定数）
- [ ] `GetDestructiveErrorLogPath()` / `GetSnapshotFilePath()` を追加
- [ ] `ClearDestructiveLog()` / `ClearSnapshot()` / `ClearCurrentTaskFile()` を追加
- [ ] `LogTaskStart(project, task, attempt)` を追加（**アプリ側の記録主体**。フォーマット: `[ts] [TaskStart] P-T-A`）
- [ ] `mos_word_current_task.txt` の書式を `ProjectId,TaskId,ExemptFlags,AttemptNo` で統一

### 1-2. UI から current_task 書き出し

- [ ] `Views/UiTestAppBarWindow.xaml.cs` の `WriteCurrentTaskFile` を `attemptNo` 含みで統一
- [ ] `Views/AppBarWindow.xaml.cs` 側も同等実装にそろえる
- [ ] 画面遷移時（タスク切替）に必ず `WriteCurrentTaskFile` が呼ばれることを確認

### 1-3. VSTO ポーリング配線

- [ ] `New_MOSWordVSTOAddIn/New_MOSWordVSTOAddIn/ThisAddIn.cs` に `current_task` ポーリング（約500ms）を配置
- [ ] task 変更検知時に「新タスクの snapshot 取得」だけを先に実装
- [ ] フェーズ0では比較ロジックを無効（または常に空エラー）にする

### 1-4. 採点ゲートの枠

- [ ] `Libraries/WordGradingGate.cs`（新規）を作成し、暫定で常に「合格」実装
- [ ] `Libraries/WordBatchScoring.cs` の `InvokeCheckTask` 前に `WordGradingGate` 呼び出しを挿入
- [ ] `ScoreSingleTask` と一括採点の両経路で同じ入口を通ることを確認

### 1-5. フェーズ0受け入れ

- [ ] ビルド成功（MOS Word app + VSTO）
- [ ] タスク切替で `%TEMP%` に `mos_word_current_task.txt` が生成/更新される（書式 `P,T,Flags,A`）
- [ ] `%TEMP%` に snapshot ファイルが生成される
- [ ] `mos_word_log.txt` に `[TaskStart]` 行が書き出される
- [ ] プロジェクトリセットで `mos_word_current_task.txt` / `mos_word_destructive_errors.log` / `mos_word_snapshot.txt` が削除される（`MainViewModel.cs`（ルート直下）から呼び出し）
- [ ] 既存採点（破壊検知なし）が従来どおり動く

---

## 2. フェーズ 1（ログベース破壊検知）

### 2-1. ログフォーマット拡張

- [ ] `TaskStart` 行フォーマットを `[TaskStart] P-T-A` で実装
- [ ] `[Op]` 行フォーマットを `[Task P-T-A] [Op] Type Detail` で実装
- [ ] `Executed` 行の既存フォーマットを維持（互換性を壊さない）

### 2-2. TaskStart 記録

- [ ] `UiTestAppBarWindow` でタスク表示・切替時に `LogTaskStart(project, task, attempt)` 呼び出し
- [ ] `AppBarWindow` でも同様に記録
- [ ] 同一タスクで重複記録しないガード（直前値比較）を追加

### 2-3. VSTO `[Op]` 記録

- [ ] `New_MOSWordVSTOAddIn/Logger.cs` に `LogOperation(type, detail)` 追加
- [ ] `SetCurrentTaskContext(project, task, attempt)` を追加し `[Op]` 行に反映
- [ ] まずは `RibbonCommand`（idMso 正規化）を記録
- [ ] 既存フックから `Cut` / `Paste` / `FileSaveAs` を優先記録

### 2-4. LogReader パース/判定

- [ ] `GetOperationsForTask(p, t, a)` 実装
- [ ] `HasDisallowedOperations(p, t, a, allowedTypes)` 実装
- [ ] 未知の OpType をどう扱うか方針確定（推奨: fail-closed）
- [ ] ログ欠損時の挙動を確定（推奨: fail-open + 警告ログ）

### 2-5. タスク設定（最小）

- [ ] `Libraries/WordTaskValidationConfig.cs`（新規）作成
  - フェーズ1ではまず `GetAllowedOperationTypes` のみ実装。`WordValidationExemptFlags` 列挙体と `GetExemptFlags` はフェーズ2（3-4）で追加する
- [ ] `GetAllowedOperationTypes(project, task)` 実装（最小タスクから）
- [ ] まず P7（7-1, 7-3, 7-4, 7-5）の allowlist を先行定義

### 2-6. 採点ゲート接続

- [ ] `WordGradingGate` に `HasDisallowedOperations` 判定を接続
- [ ] 判定 NG 時は `WordChecker` 未実行で false を返す
- [ ] 再採点経路（`ScoreSingleTask`）で同様動作を確認

### 2-7. フェーズ1受け入れ

- [ ] 正解操作のみ -> ○
- [ ] 無関係操作（例: `InsertTable`）混入 -> ✖
- [ ] ログから該当 `P-T-A` 区間が抽出できる

---

## 3. フェーズ 2（スナップショット差分）

### 3-1. スナップショットモデル

- [ ] `Libraries/WordSnapshotChecker.cs`（新規）を作成
- [ ] snapshot データ構造を定義（Sections, BodyTextLength, Shapes, Comments 等）
- [ ] `AttemptNo` を snapshot のキーに含める
- [ ] 対象文書 `Project{N}` の実ファイル解決ロジックを実装

### 3-2. 取得ロジック

- [ ] `TakeSnapshot(project, task, attempt)` 実装
- [ ] `document.Content` 由来の本文長計測を実装（ヘッダー除外ルールを明確化）
- [ ] `HeaderFooterFingerprint` 取得を既存 7-3 指紋ロジックと整合
- [ ] 例外時のフォールバック（空 snapshot / リトライ）を決める

### 3-3. 比較ロジック

- [ ] `CompareAndGetErrors(project, task, attempt, exemptFlags)` 実装
- [ ] フラグ除外 (`WordValidationExemptFlags`) を適用
- [ ] 差分エラー文言を `destructive_errors.log` に保存可能な形式へ正規化
- [ ] 同一 attempt 内の重複エラー記録を抑制

### 3-4. 免除設定（`WordTaskValidationConfig.cs` に追記）

> フェーズ1（2-5）で新規作成済みのファイルに免除フラグ機能を追加する。

- [ ] `WordValidationExemptFlags` 列挙体を定義（`[Flags]` 属性付き）
- [ ] `WordTaskValidationConfig.GetExemptFlags(project, task)` 実装
- [ ] CSV（`Reference/CSV/MOS模擬アプリ正誤判定表251120_.csv`）を根拠に初期値設定
- [ ] P7（7-1, 7-3, 7-4, 7-5）の免除を先行で確定

### 3-5. destructive_errors 連携

- [ ] `HasLoggedDestructiveError(project, task, attempt)` 実装
- [ ] `AppendDestructiveErrors(project, task, attempt, errors)` 実装
- [ ] 採点前ゲートの先頭で destructive_errors をチェック

### 3-6. フェーズ2受け入れ

- [ ] P7 T3 のみ実施 -> ○
- [ ] P7 T1 -> T3 -> ○
- [ ] 結果画面から P7 T3 復習再採点 -> ○
- [ ] T3中に無関係編集（セクション追加等） -> ✖

---

## 4. フェーズ 3（調整と安定化）

### 4-1. 誤検知/見逃し調整

- [ ] 誤検知の多いタスクを列挙し、許可 Op または免除フラグを見直す
- [ ] `BodyTextLength` の許容差分（必要ならタスク別）を導入
- [ ] 未取得 Op（リボン外）を snapshot 差分で補完できているか確認

### 4-2. 運用性

- [ ] ログ肥大化時の読み取りコストを計測
- [ ] 必要に応じてローテーション/セッション開始時アーカイブを検討
- [ ] デバッグログ（開発用）と本番判定ログ（採点用）を分離する

### 4-3. ドキュメント更新

- [ ] `Wordtasks/md/VSTO_ログコマンド一覧.md` に `[TaskStart]` / `[Op]` を追記
- [ ] `Word_破壊的操作検知_実装案.md` のステータスを更新
- [ ] 誤検知既知リストと暫定回避策を追記

---

## 5. テストチェックリスト（最終）

### 5-1. 単体テスト相当（ロジック）

- [ ] LogReader が `TaskStart` / `[Op]` / `Executed` を正しく分類
- [ ] `P-T-A` 境界で抽出が混ざらない
- [ ] `HasDisallowedOperations` が allowlist どおり動作
- [ ] `CompareAndGetErrors` が exemptFlags を正しく反映

### 5-2. 結合テスト相当（アプリ + VSTO）

- [ ] タスク切替で `current_task` と `TaskStart` が同期
- [ ] VSTO が正しい `P-T-A` で `[Op]` を記録
- [ ] snapshot が task 変更時に更新される
- [ ] 採点時に 3段ゲート順で評価される

### 5-3. 回帰テスト（既存機能）

- [ ] 既存 `Executed` 判定（特に 7-2, 7-3）が劣化しない
- [ ] 結果画面の再採点フローが維持される
- [ ] プロジェクトリセットで関連ファイルが全消去される

---

## 6. 既知リスク確認（リリース前）

- [ ] 複数 Word インスタンス起動時の想定外挙動を許容するか判断
- [ ] ログ欠損時の fail-open / fail-closed 方針を最終決定
- [ ] 7-4 / 7-5 の別文書操作で誤検知しないことを確認
- [ ] 互換モード（7-1）とヘッダー指紋（7-3）が干渉しないことを確認

---

## 7. 未導入項目バックログ（後で実装）

> すぐ導入しないが、導入忘れを防ぐための記録欄。  
> 新規候補が出たら、同じ形式で追記すること。

### 7-1. 現在の候補

- [x] **P4（4-5/4-6）Watermark / PageBorder の厳格化**
  - 背景: Op ゲート・Snapshot 指紋・タスク別免除を導入済み（リボン透かしギャラリーは Word スキーマ上不可のためポーリング）。
  - 目標: 4-5/4-6 以外で透かし・ページ罫線操作を行った場合は破壊的操作として ✖ にできる状態へ。
  - 推奨導入順:
    1. [x] Op 側（`[Op]`）で許可/不許可を明確化 — `4-5`: `Watermark` のみ、`4-6`: `PageBorders` のみ。VSTO が `[Op]` に記録（`RibbonCommand` 汎用にしない）
    2. [x] Snapshot 側に `WatermarkFingerprint` / `PageBorderFingerprint` 比較を追加（`WordSnapshotChecker` + VSTO `WordDestructiveMonitor`、v2 保存）
    3. [x] タスク別免除を最小化（4-5 は Watermark、4-6 は PageBorder）
    - [x] 透かし指紋ポーリングで `[Op] Watermark` を 4-1 等でも記録（Word は `commands` に gallery 不可のためリボン XML 未接続）
  - 確認観点:
    - [x] 正答手順（4-5/4-6）は ○ を維持（手動確認済み）
    - [x] 無関係タスク（4-1）で透かし挿入 → ✖（手動確認済み）
    - [ ] 4-7（検査で削除）との干渉がない

- [ ] **破壊ログのタスク帰属揺れ対策（切替境界）**
  - 背景: `current_task` ポーリング（約500ms）でタスク切替境界に比較が走るため、まれに `destructive_errors.log` の記録タスク番号が体感とずれることがある（例: 5-3/5-4 操作が 5-8 で記録）。
  - [x] `[Op]` 帰属: 行内 `[Task P-T-A]` を TaskStart カーソルより優先（`LogReader.GetOperationsForTask`、4-5 Watermark が 4-6 に誤帰属する不具合）
  - [x] タスク切替時に透かしポーリングベースライン再同期（`SyncWatermarkPollingBaselineOnTaskSwitch`、4-6 で残留透かしの誤 `[Op] Watermark` 抑制）
  - 目標: 破壊ログの記録タスク番号と実操作タスクの一致率を上げ、解析しやすくする。
  - 推奨導入順:
    1. タスク切替直前フラッシュ比較（UI 側）を追加し、旧タスクの比較を明示実行
    2. `TaskStart` 境界と `attemptNo` を使った記録整合チェックを追加
    3. 競合時の優先ルール（「採点対象タスク優先」など）を文書化
  - 確認観点:
    - 正答手順で誤タスク番号への記録が減る
    - 誤操作時に採点対象タスクで安定して ✖ になる
    - 再採点（`attemptNo >= 1`）でも帰属が崩れない

### 7-2. 追記テンプレート（コピペ用）

- [ ] **P?-? 追加候補タイトル**
  - 背景:
  - 目標:
  - 推奨導入順:
    1.
    2.
    3.
  - 確認観点:
    - 正答手順:
    - 誤操作検知:
    - 既存タスクへの副作用:

---

## 8. 完了記録

- 着手日:
- 完了日:
- 実装ブランチ:
- 主要変更ファイル:
- 未解決課題:

