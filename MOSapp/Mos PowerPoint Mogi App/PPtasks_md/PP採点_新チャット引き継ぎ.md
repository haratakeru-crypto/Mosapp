# PowerPoint 採点 新チャット引き継ぎ

新しいチャットでプロジェクト採点（P6 以降）を進める際に、このファイルを添付して参照させる。

**最終更新**: 2026-06-23（P6 完了・P7 Phase A 完了）

---

## 1. このリポジトリでやっていること

MOS PowerPoint 模擬試験アプリは、次の2層で採点する。

1. **ゲート（破壊的操作・ログ）** — 指示外の変更がないか
2. **COM チェッカー** — 問題の正解操作ができているか

詳細アーキテクチャ: [`validation_architecture.md`](validation_architecture.md)  
（※ `mos_ppt_current_task.txt` の書式説明は本 MD の「2. 確定済み修正」を優先）

---

## 2. 確定済み修正（2026-06 時点・維持すること）

### 2.1 `current_task` プロトコル（5項目固定）

| 項目 | 意味 |
|------|------|
| 1 | ProjectId |
| 2 | TaskId |
| 3 | ExemptFlags（整数） |
| 4 | AttemptNo |
| 5 | SnapshotGen（UI=0、採点時>0） |

- 書き込み: `PPLogReader.WriteCurrentTaskFile`（原子的書き込み）
- 読み取り: 5項目ちょうどでない行は **無視**（途中読み対策）
- UI: `UiTestAppBarWindow.WriteCurrentTaskFile` → 常に `gen=0`
- 採点: `PowerPointGrader.StartTask` → `gen>0`（`StartTaskAndWaitForSnapshot`）

### 2.2 VSTO 境界の破壊的操作比較

`ThisAddIn.TaskFilePollTimer_Tick` で、次の **すべて** を満たすときだけ `CheckAndLogDestructiveOperations` を実行する。

- `taskIdentityChanged`（project / task / attempt のいずれかが変化）
- 同一 project 内遷移
- スナップショットファイルが存在（`!forceSnapshot`）
- **`_currentSnapshotGen == 0` かつ `snapshotGen == 0`**（通常 UI 遷移のみ）

採点ループ中の `gen>0` 切替や、採点直後の混線で誤ログが出ないようにするための核心。

### 2.3 採点中の UI 書き込み停止

`UiTestAppBarWindow`:

- `BeginScoringSession` / `EndScoringSession`
- 採点中は `WriteCurrentTaskFile` をスキップ（終了時 `force: true` で UI 状態を復元）
- 対象: `ScoreButton_Click`、`ShowResultWindowAsync`、`TryRescorePendingRetryTasks`

### 2.4 破壊的操作ログの責務分担（Word 方式）

| 経路 | 破壊的ログ追記 |
|------|----------------|
| VSTO タスク境界（0→0） | `ThisAddIn.CheckAndLogDestructiveOperations` |
| 採点ゲート③ | `PPSnapshotChecker` → `PPLogReader.AppendDestructiveErrors` |
| UI プロジェクト遷移 | **追記しない** |

採点ゲート順（`PowerPointGrader.GradeTask`）:

1. `HasLoggedDestructiveError`
2. `FailsLogChecks`（許可外リボン操作など）
3. `PPSnapshotChecker.CompareAndGetErrors`
4. COM チェッカー（`PowerPointChecker1_X`）

### 2.5 検証済み（P5）

- P5-1〜7 連続操作 → 採点: **4回連続○**
- `%TEMP%\mos_ppt_destructive_errors.log`: 誤検知なし
- 意図的破壊操作: **× になる**（検知維持）
- 採点後 `mos_ppt_current_task.txt` 例: `5,7,16,1,0`（正常）

---

## 3. 共有ファイル（%TEMP%）

| ファイル | 役割 |
|----------|------|
| `mos_ppt_current_task.txt` | 現在タスク（5項目固定） |
| `mos_ppt_snapshot.txt` | タスク開始時スナップショット |
| `mos_ppt_destructive_errors.log` | 破壊的操作記録 `P,T,A:理由` |
| `mos_ppt_log.txt` | 操作ログ |
| `mos_ppt_task_evidence.txt` | 印刷・グレースケール等の採点証跡 |

---

## 4. プロジェクト実装の進め方（テンプレート）

### 4.1 参照する既存ドキュメント

| ファイル | 用途 |
|----------|------|
| [`tasks/PP_類題採点対応表.md`](../tasks/PP_類題採点対応表.md) | **新 P番号 ↔ 旧タスク ↔ Legacy チェッカー** の対応表（最重要） |
| [`PP問題文.csv`](../PP問題文.csv) | 問題文・操作手順の参照 |
| [`validation_architecture.md`](validation_architecture.md) | 採点・破壊検知の全体像 |
| [`PP採点チェッカーとVSTO要件.md`](PP採点チェッカーとVSTO要件.md) | VSTO 証跡が必要なタスク |
| [`Release検証手順_PowerPoint.md`](Release検証手順_PowerPoint.md) | リリース前確認 |

### 4.2 実装時に触るファイル（典型）

| 順 | ファイル | 作業内容 |
|----|----------|----------|
| 1 | `Libraries/Group1/PowerPointChecker1_{N}.cs` | `CheckTask_1_{N}_0X` を実装（**上書き**。Legacy は参照のみ） |
| 2 | `PowerPointGrader.cs` | `case {N}:` に taskId 1〜X を振り分け |
| 3 | `Libraries/PPTaskValidationConfig.cs` | `GetExemptFlags` とデルタ/位置上限を **新 projectId・taskId** で定義 |
| 4 | `PowerPointAddIn1/ThisAddIn.cs` | 印刷・光彩等、VSTO 証跡が必要なら **新番号** に合わせて拡張 |
| 5 | `tasks/PP_類題採点対応表.md` | P{N} セクション・チェックリストを更新 |

### 4.3 実装ルール

- **`Libraries/Legacy/*` は編集しない**（参照・コピー元のみ）
- メソッド名: `CheckTask_1_{projectId}_{taskId:00}`（例: P6-3 → `CheckTask_1_6_03`）
- COM オブジェクトは `PowerPointCheckerCommon` のパターンに従い `Marshal.ReleaseComObject`
- 破壊的操作の免除は「フラグだけ」でなく、必要なら **スライド別デルタ**・**既存図形位置上限** も設定（P5 参照）
- 印刷・グレースケール・キオスク等は **COM だけでは取れない** ことがある → VSTO 証跡ログを確認

### 4.4 ビルド・動作確認

```text
ソリューション: MOSapp/Mos PowerPoint Mogi App/MOS PowerPoint app.sln
構成: Debug
```

- **VSTO 変更後は PowerPoint 再起動必須**
- 正常系: タスク順操作 → 採点 ○
- 破壊的ログ: 誤検知が出ないこと
- 負例: 意図的な指示外操作で ×
- 連続採点: 2回以上で再発しないこと

---

## 5. 完了済みプロジェクト

| プロジェクト | 状態 | メモ |
|-------------|------|------|
| P1 | 確認中 | |
| P2 | 完了 | P2-4 累積採点の既知課題あり |
| P3 | 完了 | |
| P4 | 実装済・検証待ち | |
| **P5** | **完了・検証済** | 破壊的操作誤検知修正も含む |
| **P6** | **完了・検証済** | P6-1〜7 連続採点○。破壊的操作免除は全タスクなし |
| **P7** | **Phase A 完了** | Checker スタブ・Grader 01-05・config 付替・問題文 JSON |
| P8〜P11 | 未着手 | |

---

## 6. 完了: プロジェクト6（P6）最終仕様メモ

### 6.0 検証結果

- P6-1〜P6-7 を一連で操作・採点 → **全タスク ○**
- 破壊的操作免除の追加調整は **不要**（全タスク `ExemptFlags = None`、スライド別デルタ・位置上限もなし）
- 旧 `projectId=6 taskId=3/4`（6-3/6-4 3Dモデル）の免除は Phase A で削除済み（P3-4/5 に移植済み）

### 6.1 タスク一覧と合格条件

`projectId = 6`。チェッカー: `PowerPointChecker1_6.cs`（`CheckTask_1_6_0X`）。

| タスク | 旧 | メソッド | 合格条件（COM） | VSTO 証跡 |
|--------|-----|----------|----------------|-----------|
| **P6-1** | 旧10-1 | `CheckTask_1_6_01` | 全スライド `Comments.Count == 0`。BuiltIn プロパティ（Author, Manager, Company, Last Author, Title, Subject, Keywords, Comments）がすべて空 | 不要 |
| **P6-2** | 旧8-5 | `CheckTask_1_6_02` | SaveCopyAs → OpenXML で `readOnlyRecommended` 検出、または `pres.ReadOnly == msoTrue` | 不要 |
| **P6-3** | 旧7-4 | `CheckTask_1_6_03` | `SlideShowSettings.ShowType == ppShowTypeKiosk` | `[Task6-3] Kiosk`（main log のみ。COM フォールバックあり） |
| **P6-4** | 旧10-2 | `CheckTask_1_6_04` | 目的別スライドショー名 **「書式のポイント」** にスライド **4・5・6** がこの順で含まれる（Legacy「教育」から名称変更） | 不要 |
| **P6-5** | 旧5-1 | `CheckTask_1_6_05` | `OutputType == ppPrintOutputOutline`、`NumberOfCopies == 6`、`Collate == msoTrue`（部単位） | `[Task6-5] Print` |
| **P6-6** | 旧11-7 | `CheckTask_1_6_06` | `OutputType == ppPrintOutputNotesPages`、`NumberOfCopies == 3`、`Collate == msoFalse`（**ページ単位**） | `[Task6-6] Print` |
| **P6-7** | 旧5-1 | `CheckTask_1_6_07` | `OutputType == ppPrintOutputThreeSlideHandouts`、`NumberOfCopies == 4`、`PrintColorType == ppPrintBlackAndWhite`（印刷オプションのグレースケール）。**Collate はチェックしない** | `[Task6-7] Print` |

> **P6-5 / P6-6 / P6-7** はハイブリッド判定: VSTO 証跡があれば即 ○、なければ COM で上記条件を確認。

### 6.2 重要な差分・注意点

| 項目 | 内容 |
|------|------|
| P6-4 ショー名 | Legacy 旧10-2 は「教育」→ 新問題文は **「書式のポイント」** |
| P6-5 vs 旧5-1 | 同型（アウトライン・6部・部単位）。証跡タグは `[Task6-5] Print`（旧 `[Task5-1] Print` とは別） |
| P6-6 vs 旧11-7 | **Collate 条件が逆**。旧11-7 は部単位（`msoTrue`）、P6-6 は問題文どおり **ページ単位（`msoFalse`）** |
| P6-7 vs 旧5-1 | 配布資料3スライド/頁・4部に加え **グレースケール**（`PrintColorType`）が必須。10-4 の表示グレースケール（`BlackWhiteMode`）とは別 |
| `PrintColorType` | Interop には `ppPrintColorGrayscale` なし。日本語UI「グレースケール」= **`ppPrintBlackAndWhite`（値2）** |
| 印刷タスク | 印刷実行は不要。設定変更の検知が目的（VSTO ポーリング + タスク離脱時の境界再評価） |

### 6.3 破壊的操作免除（`PPTaskValidationConfig`）

**全タスク（P6-1〜7）: 免除フラグなし（`PPValidationExemptFlags.None`）**

| チェック種別 | P6 での設定 |
|-------------|------------|
| ExemptFlags | なし（全7タスク） |
| 図形数デルタ | なし |
| 文字数デルタ | なし |
| 既存図形位置上限 | なし |

P6 の操作はスライド内容・図形の変更を伴わない（ドキュメント検査・読み取り専用・スライドショー設定・印刷オプション）ため、P5 のような免除調整は不要。連続採点で誤検知が出ないことも検証済み。

### 6.4 VSTO 証跡（`ThisAddIn` / `Logger` / `PPLogReader`）

| タスク | Logger | PPLogReader | 境界処理 |
|--------|--------|-------------|----------|
| P6-3 | `LogTask6_3Kiosk()` | `HasTask6_3KioskExecuted()` | キオスク設定ポーリング |
| P6-5 | `LogTask6_5Print()` | `HasTask6_5PrintExecuted()` | `TryLogTask6_5PrintOnTaskBoundary()` |
| P6-6 | `LogTask6_6Print()` | `HasTask6_6PrintExecuted()` | `TryLogTask6_6PrintOnTaskBoundary()` |
| P6-7 | `LogTask6_7Print()` | `HasTask6_7PrintExecuted()` | `TryLogTask6_7PrintOnTaskBoundary()`（P6 最終タスク・レビュー遷移時の取りこぼし対策あり） |

印刷ポーリング（`PrintOptionsPollTimer`）は P6-5/6/7 を監視。P6-7 では `PrintColorType` の変化も変更検知に含む（`_lastPrintColorType`）。

証跡リセット: `PPLogReader.GetEvidenceLineMarkersToRemoveForProject(6)` に `[Task6-5] Print`, `[Task6-6] Print`, `[Task6-7] Print` を登録済み。

### 6.5 触った主要ファイル

| ファイル | 内容 |
|----------|------|
| `Libraries/Group1/PowerPointChecker1_6.cs` | P6-1〜7 の COM チェッカー |
| `PowerPointGrader.cs` | `case 6:` → `CheckTask_1_6_01`〜`07` |
| `Libraries/PPTaskValidationConfig.cs` | `projectId==6` を P6 用に付け替え（旧6-3/6-4 削除） |
| `PowerPointAddIn1/ThisAddIn.cs` | 印刷・キオスク証跡、境界処理 |
| `PowerPointAddIn1/Logger.cs` | Task6-3/5/6/7 タグ |
| `Libraries/PPLogReader.cs` | HasTask6_* 系 |
| `MOS模擬アプリ問題文一覧_PowerPoint.json` | projectId=6 を7問に更新 |

---

## 7. 次の作業: プロジェクト7（P7）

### 7.0 ハイブリッド進め方

| フェーズ | 内容 | 状態 |
|----------|------|------|
| **Phase A** | 問題文・対応表・config 骨格・Grader 振り分け・Checker スタブ（全 `return false`） | **完了** |
| **Phase B** | タスク単位: Legacy 移植 → VSTO（要否）→ 検証 → 次タスク | P7-1 から順 |

**Phase B 推奨順**: P7-1 → P7-2 → P7-3 → P7-4 → P7-5

### 7.1 タスク一覧（新番号）

`projectId = 7`。`PowerPointChecker1_7.cs` を **P7 用に上書き**する（旧7-1/7-2/7-4 は新P7に含まれない）。

| 新タスク | 旧タスク | 移植元 Legacy | 概要 |
|----------|----------|---------------|------|
| P7-1 | 旧6-1 | `PowerPointChecker1_6.Legacy.cs` `CheckTask_1_6_01` | スライド1コメント（文言は問題文差分） |
| P7-2 | 旧9-6 | `PowerPointChecker1_9.Legacy.cs` `CheckTask_1_9_06` | ハイパーリンク（対象文字列・URL差分） |
| P7-3 | 旧7-3 | `PowerPointChecker1_7.Legacy.cs` `CheckTask_1_7_03` | アウトライン挿入（スライド6後・文書「まとめ」） |
| P7-4 | 旧9-4 | `PowerPointChecker1_9.Legacy.cs` `CheckTask_1_9_04` | フッター（rabbitway.jp・タイトル以外） |
| P7-5 | 旧9-5 | `PowerPointChecker1_9.Legacy.cs` `CheckTask_1_9_05` | フッター（スライド5・6のみ「参考事例」） |

### 7.2 破壊的操作免除（Phase A 設定済み）

| タスク | 免除 | 備考 |
|--------|------|------|
| P7-1 | なし | |
| P7-2 | Text, Position | 文字数デルタ -57 は暫定（旧9-6流用）。Phase B で再確認 |
| P7-3 | Slides, Shapes, Text, Position | 旧7-3 から移植 |
| P7-4, P7-5 | Shapes, Text, Position | 旧9-4/5 から移植 |

### 7.3 Phase B で注意する差分

| タスク | Legacy との差分 |
|--------|----------------|
| P7-1 | コメント文言「情報発信の責任を考える」（旧6-1 は「MOSの説明は詳しく」） |
| P7-2 | 対象「情報学習支援」、URL `https://rabbitway.jp/`（旧9-6 は別文字列） |
| P7-3 | 挿入位置スライド6後、文書「まとめ」（旧7-3 はスライド5後・「弊社の他の講座一覧」） |
| P7-4 | フッター「rabbitway.jp」（旧9-4 は www.MOS.jp） |
| P7-5 | スライド5・6「参考事例」（旧9-5 はスライド5のみ「集中的に」） |

---

## 8. やってはいけないこと

- `Libraries/Legacy/*` の編集
- 境界比較を `snapshotGen>0` の遷移でも走らせる変更
- 採点中に UI から `gen=0` の `current_task` を書くこと
- `mismatch` / `timeout` 時にスナップショットチェックを無条件スキップすること
- ビルド生成物（`bin/` `obj/` `.vs/`）のコミット

---

## 9. トラブルシュート早見表

| 症状 | 疑う箇所 |
|------|----------|
| 正解なのに ×（破壊的ログあり） | VSTO 境界の誤比較、`current_task` 混線、免除不足 |
| `5,7,1:ShapesCount...` 系 | `ThisAddIn` 境界比較が採点中に走っていないか（`0→0` 条件） |
| COM は ○ だが全体 × | `HasLoggedDestructiveError` または `FailsLogChecks` |
| 印刷タスクだけ × | VSTO 証跡 `[Task*-] Print` が記録されているか |
| 採点が不安定 | `StartTaskAndWaitForSnapshot` タイムアウト、VSTO 未再起動 |

デバッグ出力キーワード: `[DestructiveBoundary]` `[Grader]` `[Perf]` `[UiTestAppBarWindow] BeginScoringSession`

---

## 10. 新チャット初回プロンプト（コピペ用）

以下を **そのまま** 新チャットの1通目に貼り、添付ファイルとして本 MD と関連ファイルを指定する。

```text
PowerPoint MOS 模擬アプリの採点実装を続けます。まず添付の「PP採点_新チャット引き継ぎ.md」を読み、現行アーキテクチャと確定済み修正（current_task 5項目・境界比較 0→0・採点中UI書き込み停止）を前提に作業してください。

## 今回のゴール
プロジェクト6（P6-1〜P6-7）の採点ロジックを実装・検証してください。

## 必読ファイル
- MOSapp/Mos PowerPoint Mogi App/PPtasks_md/PP採点_新チャット引き継ぎ.md（本引き継ぎ）
- MOSapp/Mos PowerPoint Mogi App/tasks/PP_類題採点対応表.md（旧→新対応）
- MOSapp/Mos PowerPoint Mogi App/PPtasks_md/validation_architecture.md
- MOSapp/Mos PowerPoint Mogi App/Libraries/Group1/PowerPointChecker1_5.cs（完了済みの実装例）
- MOSapp/Mos PowerPoint Mogi App/Libraries/PPTaskValidationConfig.cs

## 制約
- Libraries/Legacy/* は編集しない（参照・移植のみ）
- 破壊的操作の誤検知修正（0→0境界、採点中WriteCurrentTaskFile停止、5項目current_task）は壊さない
- メソッド名は CheckTask_1_6_0X（P6-X）に統一
- VSTO変更後はPowerPoint再起動が必要な旨を作業末尾に記載

## 作業手順（この順で）
1. P6 各タスクの Legacy チェッカーと問題文を照合し、実装方針を短く提示
2. PowerPointChecker1_6.cs を P6 用に実装（Grader の case 6 に 05〜07 追加）
3. PPTaskValidationConfig の projectId==6 を P6 用に更新（旧6-3/6-4 設定の置き換え）
4. 印刷・キオスク等、VSTO 証跡が必要なタスクがあれば ThisAddIn / Logger / PPLogReader を拡張
5. Debug ビルド
6. 検証手順（正常系・連続採点・意図的破壊の負例）を提示

## 検証の合格基準
- P6-1〜7 を順に操作して採点し、正解操作で ○
- mos_ppt_destructive_errors.log に誤検知が出ない（2回以上の連続採点）
- 意図的な指示外操作で × になる
- 採点後の mos_ppt_current_task.txt が 5項目かつ最後が ,0 であること

不明点があれば実装前に質問してください。推測で破壊的操作まわりの既存ガードを弱めないでください。
```

### 添付推奨ファイル（@ で指定）

- `MOSapp/Mos PowerPoint Mogi App/PPtasks_md/PP採点_新チャット引き継ぎ.md`
- `MOSapp/Mos PowerPoint Mogi App/tasks/PP_類題採点対応表.md`
- `MOSapp/Mos PowerPoint Mogi App/PP問題文.csv`
- `MOSapp/Mos PowerPoint Mogi App/Libraries/Group1/PowerPointChecker1_6.cs`（上書き対象）
- 移植元 Legacy（P6 表に従い必要なもの）:
  - `Libraries/Group1/Legacy/PowerPointChecker1_10.Legacy.cs`
  - `Libraries/Group1/Legacy/PowerPointChecker1_8.Legacy.cs`
  - `Libraries/Group1/Legacy/PowerPointChecker1_7.Legacy.cs`
  - `Libraries/Group1/Legacy/PowerPointChecker1_5.Legacy.cs`
  - `Libraries/Group1/Legacy/PowerPointChecker1_11.Legacy.cs`

---

## 11. コミットメッセージ案（破壊的操作修正まとめ）

```text
Fix false destructive-operation logs during PowerPoint grading.

Unify current_task to five fields with atomic writes, run VSTO boundary checks only on UI transitions (snapshotGen 0→0), and suppress UI current_task updates while scoring.
```
