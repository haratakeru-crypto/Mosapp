# Legacy — 破壊的操作免除設定（PPTaskValidationConfig）

P4 着手前（P3 完了時点）の `PPTaskValidationConfig.cs` 参照用。**ビルド対象外**（`.csproj` に含めない）。

## Baseline

| 項目 | 値 |
|------|-----|
| 作成日 | 2026-06-17 |
| 最終更新 | 2026-06-17（P3-7 既存図形位置: 上限1 → 無制限に現行 config と同期） |
| スナップショット元 | `Libraries/PPTaskValidationConfig.cs`（P3 全タスク完了時点） |
| git HEAD（参考） | `4840fd5` |
| 人間向け一覧 | `tasks/PP_類題採点対応表.md` →「破壊的操作免除一覧」 |

## ファイル

| ファイル | 内容 |
|----------|------|
| `PPTaskValidationConfig.Legacy.cs` | 上記時点の免除設定一式（GetExemptFlags / デルタ判定 / 既存図形上限など） |

## 使い方

1. **日常**: `tasks/PP_類題採点対応表.md` の免除表を見る（主）
2. **旧 taskId のコード確認**: 本 Legacy ファイルを検索（例: `projectId == 6 && taskId == 3`）
3. **最終手段**: `git log -p -- Libraries/PPTaskValidationConfig.cs`

## Px 実装時の運用

- **P4 着手前**: 必要なら本フォルダを再度コピーして baseline を更新
- **taskId 付け替え後**: 対応表の免除行を現行 config に合わせて更新。旧設定は本 Legacy に残る
- **チェッカー Legacy**（`Group1/Legacy/`）とは別管理。クラス名は同じ `PPTaskValidationConfig` のため、本フォルダを Compile に含めると重複エラーになる

## 注意

- VSTO（`PowerPointAddIn1/ThisAddIn.cs`）にもデルタ判定の重複実装あり。config 変更時は **両方** を更新すること
- 現行 config は P1/P2/P3 のみ **新 taskId（Px-X）** で記述。projectId 4〜11 は **旧番号のまま**（P4 実装時に付け替え予定）
