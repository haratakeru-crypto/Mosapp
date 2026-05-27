## Excelプロジェクト入れ替え 変更メモ（Group1/演習）

### 矢印の意味（読み方B）

**「プロジェクトN→M」＝ 以前のプロジェクトNの内容が、スロットMに載る（移動先がM）**

| 指定 | 意味 |
|------|------|
| 1→3 | 旧P1 → スロット3 |
| 2→1 | 旧P2 → スロット1 |
| 3→2 | 旧P3 → スロット2 |
| 9→7 | 旧P9 → スロット7 |
| 7→8 | 旧P7 → スロット8 |
| 8→9 | 旧P8 → スロット9 |

### 画面上のスロットに載る内容

| スロット | 旧# | library | taskCount | 1問目の目印 |
|---------|-----|---------|-----------|------------|
| 1 | 2 | ExcelChecker1_2 | 5 | 試験結果・テーブル |
| 2 | 3 | ExcelChecker1_3 | 7 | 下半期売上・A2 |
| 3 | 1 | ExcelChecker1_1 | 7 | 売上一覧・印刷 |
| 7 | 9 | ExcelChecker1_9 | 7 | 売上報告・数式表示 |
| 8 | 7 | ExcelChecker1_7 | 7 | イベント売上・H5 |
| 9 | 8 | ExcelChecker1_8 | 6 | 学生名簿・名前定義 |

---

### 変更ファイル一覧

1. **`MOSapp/mos_xaml_app/References/JSON/MOS演習問題文一覧.json`**  
   `projectId` 1/2/3/7/8/9 の `tasks` を上表どおり入れ替え

2. **`MOSapp/mos_xaml_app/Assets/config.json`**  
   `tabs["1"].projects["1"|"2"|"3"|"7"|"8"|"9"]` の `library` / `taskCount`

3. **`MOSapp/mos_xaml_app/Libraries/ExcelTaskValidationConfig.cs`**  
   `GetExemptFlags` / `GetAllowedRanges` の case 1/2/3/7/8/9

4. **`MOSapp/mos_xaml_app/References/Answers/Group1/Project1〜3,7〜9/`**  
   フォルダ単位で循環リネーム（中身ごと移動）

5. **`MOSapp/mos_xaml_app/Ui/ViewModels/MainViewModel.cs`**  
   `ExecuteScoreAsync`：`projectConfig["library"]` を優先

6. **`MOSapp/mos_xaml_app/Views/ReviewPageWindow.xaml.cs`**（一括採点）  
   - 破壊的操作検証：`ApplyDestructiveValidation` に **スロット番号**（`project.projectId`）を渡す（従来は `ExcelChecker1_N` の N を渡していた）  
   - チェッカー呼び出し：`private CheckTask_* (string filePath)` があれば `expectedFilePath` で採点（ActiveWorkbook 依存を低減）

7. **`MOSapp/mos_xaml_app/Libraries/ExcelTaskValidationConfig.cs`** + **`ExcelLogReader.cs`**（提案A）  
   - `GetTargetSheets(projectId, taskId)` でタスク想定シートを定義  
   - 一括採点の破壊的操作判定で、**対象シート外**の `[Op]` は違反にしない（例: 10-1 文脈の `受注明細!SetPageBreak`）

8. **`MOSapp/mos_xaml_app/Ui/ViewModels/MainViewModel.cs`**（その場採点）  
   - `ExecuteScoreAsync` でチェッカー結果後に `ApplyDestructiveValidationForTask` を適用（一括採点と同じ破壊的操作ルール）

---

### ユーザー側対応（未実施）

- `C:\MOSTest\Excel365\Tab1\Initial\project{N}.xlsx` 等の実ファイルリネーム（上表の旧#→スロット対応に合わせる）

### ビルド後の確認

- `mos_xaml_app` を **Release** でリビルドし、起動 exe 配下の  
  `References\JSON\MOS演習問題文一覧.json` が更新されていることを確認
- MOSapp から起動する場合、`MOSapp/Assets/config.json` の `excelExePath` は Release を指している
