# Legacy PowerPoint Checkers

P1 改修前の旧採点ロジック参照用。**ビルド対象外**（`.csproj` に含めない）。

## Baseline

| 項目 | 値 |
|------|-----|
| 作成日 | 2026-06-15 |
| P1 改修前 commit | `d03e149` |
| P1 改修 commit（参考） | `81ecce9` |

## ファイル一覧

| ファイル | 内容 |
|----------|------|
| `PowerPointChecker1_1.Legacy.cs` | commit `d03e149` から復元（P1 改修前・7タスク） |
| `PowerPointChecker1_2.Legacy.cs` 〜 `1_11.Legacy.cs` | 2026-06-15 時点の `Libraries/Group1/` をそのままコピー |

## 使い方

1. `tasks/PP_類題採点対応表.md` の「旧」列で Legacy ファイル・メソッドを特定
2. 該当 `.Legacy.cs` からロジックをコピーし、新問題文に合わせて定数・スライド番号等を調整
3. 実装先は `PowerPointChecker1_{新プロジェクト}.cs`（Legacy は変更しない）

## git 確認メモ（依頼1実施時）

- `PowerPointChecker1_1.cs`: HEAD（`81ecce9`）に P1 改修済み。作業ツリーに**未コミットの追加差分**あり
- `PowerPointChecker1_2.cs` 〜 `1_11.cs`: 改修なし（Legacy コピー = 現行と同一）
- Legacy の `1_1` は HEAD ではなく **`d03e149`** を使用（HEAD は既に新 P1）

## 注意

- クラス名は現行と同じ `PowerPointChecker1_X` のため、本フォルダを Compile に含めると重複エラーになる
- P{N} 実装完了時に `PowerPointChecker1_N.cs` が上書きされても、Legacy は残る（**編集しない**）
- 破壊的操作の Legacy は `Libraries/Legacy/PPTaskValidationConfig.Legacy.cs`（Checker Legacy とは別フォルダ。同じく**凍結**）
