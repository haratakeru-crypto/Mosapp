## Excelプロジェクト入れ替え 変更メモ（Group1/演習）

入れ替え対応（スロット番号に対する元コンテンツの対応）

- Project1 → 3
- Project2 → 1
- Project3 → 2
- Project7 → 8
- Project8 → 9
- Project9 → 7

備考

- ここでいう「ProjectN」は `Group1` の表示上のプロジェクト番号（`config.json` の `tabs["1"].projects["N"]` や、問題文JSONの `projectId:N` 等）です。
- 実Excelファイル（`C:\MOSTest\...projectN.xlsx`）のリネームはユーザー側で対応する前提です。

---

### 変更ファイル一覧と箇所

#### 1) 問題文 JSON（AppBar表示 / 問題文一覧）
- **ファイル**: `MOSapp/mos_xaml_app/References/JSON/MOS演習問題文一覧.json`
- **箇所**: `projects` 配列内の `projectId` が **1 / 2 / 3 / 7 / 8 / 9** の各 `tasks` ブロック
- **対応内容**: 各 `tasks[].description` のセットを、下表の元コンテンツと入れ替え
  - `projectId:1` ← 旧 `projectId:3`
  - `projectId:2` ← 旧 `projectId:1`
  - `projectId:3` ← 旧 `projectId:2`
  - `projectId:7` ← 旧 `projectId:8`
  - `projectId:8` ← 旧 `projectId:9`
  - `projectId:9` ← 旧 `projectId:7`

#### 2) 採点DLLの選択 / タスク数（config）
- **ファイル**: `MOSapp/mos_xaml_app/Assets/config.json`
- **箇所**: `tabs["1"].projects["1"|"2"|"3"|"7"|"8"|"9"]`
- **対応内容**:
  - `library` を入れ替え
  - `taskCount` も入れ替え後の内容に合わせて調整

入れ替え後の設定（Group1/Tab1）

| スロット | library | taskCount |
|---|---|---|
| 1 | `ExcelChecker1_3` | 7 |
| 2 | `ExcelChecker1_1` | 7 |
| 3 | `ExcelChecker1_2` | 5 |
| 7 | `ExcelChecker1_8` | 6 |
| 8 | `ExcelChecker1_9` | 7 |
| 9 | `ExcelChecker1_7` | 7 |

#### 3) 破壊的操作検知の免除/許可範囲
- **ファイル**: `MOSapp/mos_xaml_app/Libraries/ExcelTaskValidationConfig.cs`
- **箇所1**: `GetExemptFlags(int projectId, int taskId)` 内の `switch (projectId)` **case 1 / 2 / 3 / 7 / 8 / 9**
- **箇所2**: `GetAllowedRanges(int projectId, int taskId)` 内の `switch (projectId)` **case 1 / 2 / 7 / 8 / 9**（入れ替え反映）
- **対応内容**: config.json と同じ対応関係になるように、各 case 内の免除フラグ/許可セル範囲を入れ替え

#### 4) 採点結果に表示する解答画像（結果ダイアログ等）
- **ファイル（フォルダ）**:
  - `MOSapp/mos_xaml_app/References/Answers/Group1/Project1/`
  - `MOSapp/mos_xaml_app/References/Answers/Group1/Project2/`
  - `MOSapp/mos_xaml_app/References/Answers/Group1/Project3/`
  - `MOSapp/mos_xaml_app/References/Answers/Group1/Project7/`
  - `MOSapp/mos_xaml_app/References/Answers/Group1/Project8/`
  - `MOSapp/mos_xaml_app/References/Answers/Group1/Project9/`
- **箇所**: 上記フォルダ名（ProjectN）に紐づく `Task{taskId}.png` を、循環入れ替えするためのフォルダスワップ
- **対応内容（循環）**:
  - `Project1` ← 旧 `Project3`
  - `Project2` ← 旧 `Project1`
  - `Project3` ← 旧 `Project2`
  - `Project7` ← 旧 `Project8`
  - `Project8` ← 旧 `Project9`
  - `Project9` ← 旧 `Project7`

※表示側は `Task{taskId}.png` を `Group{groupId}/Project{projectId}` 配下から組み立てるため、CSVは基本的にそのまま運用できます。

#### 5) アプリバーのその場採点（MainViewModel 単体採点）
- **ファイル**: `MOSapp/mos_xaml_app/Ui/ViewModels/MainViewModel.cs`
- **箇所**: `ExecuteScoreAsync(...)` 内の `libraryName` 決定
- **対応内容**:
  - 変更前: `libraryName = $"ExcelChecker{groupId}_{projectId}"`
  - 変更後: `projectConfig["library"]` を優先し、空なら従来命名にフォールバック

---

### 未実施 / ユーザー側対応（前提）

- 実Excelファイル（`C:\MOSTest\Excel365\Tab{groupId}\Initial\project{N}.xlsx`、および関連 Template/Tab2 等）の **リネーム/差し替え（中身入れ替え）**

