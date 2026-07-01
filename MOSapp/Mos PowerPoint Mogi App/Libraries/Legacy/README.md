# Legacy — 破壊的操作免除設定（PPTaskValidationConfig）

P3 完了時点の `PPTaskValidationConfig.cs` を**凍結保存**した参照用コピー。**ビルド対象外**（`.csproj` に含めない）。

`Libraries/Group1/Legacy/*.Legacy.cs`（チェッカー）と**同じ運用**: Px 実装中は**編集しない**。

## 何のためにあるか

| 用途 | 参照先 |
|------|--------|
| **日常・現行タスクの免除一覧** | `tasks/PP_類題採点対応表.md`（主） |
| **実行時の設定** | `Libraries/PPTaskValidationConfig.cs` |
| **VSTO のデルタ判定** | `PowerPointAddIn1/ThisAddIn.cs`（config と同期） |
| **旧 `(projectId, taskId)` のコード確認** | 本ファイル（例: `projectId == 6 && taskId == 3`） |
| **最終手段** | `git log -p -- Libraries/PPTaskValidationConfig.cs` |

## Px 実装時に更新するファイル（本 Legacy は含めない）

```
✅ Libraries/PPTaskValidationConfig.cs
✅ PowerPointAddIn1/ThisAddIn.cs   （デルタ判定の重複あり）
✅ tasks/PP_類題採点対応表.md       （破壊的操作免除一覧・P4 セクション等）
❌ Libraries/Legacy/PPTaskValidationConfig.Legacy.cs  ← 触らない
```

## Baseline（凍結時点）

| 項目 | 値 |
|------|-----|
| 作成日 | 2026-06-17 |
| 凍結内容 | P3 完了時点。`projectId` 1〜3 は新 taskId（P1/P2/P3-X）、**4〜11 は旧 taskId のまま** |
| git 復元基準 | `37275f4`（PP_Project3採点完了） |
| 人間向け一覧 | `tasks/PP_類題採点対応表.md` |

## チェッカー Legacy との対応

| 種類 | パス | 凍結の意味 |
|------|------|-----------|
| チェッカー | `Group1/Legacy/PowerPointChecker1_*.Legacy.cs` | 上書き前の**採点ロジック** |
| 破壊的操作 | `Libraries/Legacy/PPTaskValidationConfig.Legacy.cs` | P4 着手前の**旧番号 config** |

どちらも「新タスク実装のたびに同期更新」は**しない**。新設定は現行ファイルと MD に書く。

## 再スナップショットが必要なときだけ

次のような**意図的なマイルストーン**でのみ、管理者が手動でコピーし直す（日常の Px 1 件完了ごとではない）。

- 旧 projectId 4〜11 をすべて新番号に付け替え終えた大きな区切り
- 対応表の「旧→新移植メモ」が Legacy の旧番号と乖離したと判断したとき

その際も **git で履歴を残す**こと。実行中の `PPTaskValidationConfig.cs` を無断で Legacy に上書きしない。

## 採点への影響

本ファイルを編集しても**採点結果は変わらない**（コンパイルされないため）。  
誤って Legacy だけを現行 config と同期すると、**将来の移植参照**（例: 旧 `(4,5)` → P5-1）が壊れるだけ。

## 注意

- クラス名は現行と同じ `PPTaskValidationConfig` のため、Compile に含めると重複エラーになる
- `projectId == 4` のブロックは**旧 4-4/4-5/4-6** のまま。P4-X の設定は現行 config と対応表を見る
