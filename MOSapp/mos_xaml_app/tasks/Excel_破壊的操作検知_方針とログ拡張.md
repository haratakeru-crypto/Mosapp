# Excel 破壊的操作検知 — 方針とログ拡張

Excel 模擬アプリにおける「破壊的操作」の扱い、採点との関係、PowerPoint との比較、今後の拡張の進め方をまとめる。

### `ExcelTaskValidationConfig` のコード配置（参照用）

`ExcelTaskValidationConfig.cs` 1 ファイル内を **`#region`** で次のように区切っている（免除 / 許可 / 禁止 / 許可範囲）。

採点フロー（方式の分岐）は `ReviewPageWindow.ApplyDestructiveValidation` の XML コメントを参照。

---

## 0. 用語（このドキュメントでの呼び方）

| 読みやすい説明 | 内容 |
|----------------|------|
| **免除ベースの判定** | タスクごとに「これは不正にしない」操作カテゴリ（**免除**）を決め、**ログに記録された操作のうち、免除に当てはまらないものが1件でもあれば誤答**とする考え方。許可リストで列挙するのではなく、「免除以外はすべてアウト」に近い。 |
| **コード上の名前** | `ExcelLogReader.TryGetFirstNonExemptViolation`。**プロジェクト1**の採点でのみ使う。 |

※ 開発中の内部呼称だった「方式A」は、このドキュメントでは使わず、上記の表現に置き換えている。

---

## 1. 採点の二層構造（プロジェクト1のタスクなど）

1. **チェッカー DLL（例: `ExcelChecker1_1`）**  
   保存済みブックを COM で開き、問題ごとの正誤条件（印刷の向き、範囲、数式など）を判定する。  
   メソッド例: `CheckTask_1_1_01` … `CheckTask_1_1_07`。

2. **破壊的操作検知（免除ベースの判定）**  
   チェッカーが正解でも、**操作ログ上に「そのタスクの免除に該当しない操作」が含まれる**と **誤答** に落とす。  
   実装: `ReviewPageWindow.ApplyDestructiveValidation` → `ExcelLogReader.TryGetFirstNonExemptViolation`（**プロジェクト1のみ**このルールを適用）。

---

## 2. ログファイルの役割

| ファイル | 役割 |
|----------|------|
| `%TEMP%\mos_excel_log.txt` | VSTO が追記する **主ログ**。**`[Op] 操作種別 detail`** 形式の行が、**免除ベースの判定の入力**になる。 |
| `%TEMP%\mos_excel_current_task.txt` | アプリバーが現在の `projectId,taskId` を書き、VSTO がポーリングしてタスク文脈を同期する。 |
| `%TEMP%\mos_excel_destructive_errors.log` | 採点時に違反が確定したときに **理由を追記** する用途。**採点判定の必須入力ではない**（免除ベースの判定は `mos_excel_log.txt` を読んで再計算するため、このファイルが無くても正誤は決まる）。 |

---

## 3. プロジェクト1の「免除ベースの判定」の意味

- **`ExcelTaskValidationConfig.GetExemptFlags(projectId, taskId)`** でタスクごとの **免除フラグ**（`PrintAndPage`、`CellFormatOnly` など）を取得する。
- **`GetExemptCategoryForOperation`** で各 `ExcelOperationType` がどの免除カテゴリに属するかを決める。
- ログから該当タスク区間の `[Op]` を集め、**免除に入らない操作が1つでもあれば違反**。  
  許可範囲が定義されているタスク（例: 1-7 の G4）は、**`TryGetFirstNonExemptViolation` 内で許可範囲判定**も行う。

**重要:** 免除したい「正しい操作」も、**ログに `ExcelOperationType` と一致する種別名で `[Op]` として出ていないと**、免除判定に乗らない（＝正規操作が無い扱いにならない）。  
現状は **`Application.SheetChange` 由来のセル編集（`EditCellValue` / `EditCellFormula`）が主**で、印刷設定などは **ログに出ない** 場合がある。

---

## 4. PowerPoint・Word との関係

- **PowerPoint** は試験中にスナップショット比較し、違反時に **`mos_ppt_destructive_errors.log` に追記**し、採点時は **そのファイルの有無を先に見る** 流れがある。
- **Excel** は **採点時にメインログを読んで再計算する**方針を優先し、**Office ごとに挙動を完全一致させる必要はない**とする。
- **Word** も別実装になり得る。利用者向けの説明だけ揃え、内部は **Excel はメインログ拡張＋採点時判定**でよい、という整理が妥当。

---

## 5. 推奨方針：メインログ拡張（現状の採点を大きく壊さない）

- **採点の骨格**（`mos_excel_log.txt` ＋ `ExcelTaskValidationConfig` ＋ **免除ベースの判定**）は維持する。
- **不足しているのは「観測」**なので、VSTO 側で **`Logger.LogOperation(operationType, detail)`** を増やし、**`ExcelOperationType` の列挙子名と一致する `[Op]`** を出していく。
- PowerPoint のような **ブック全体スナップショット比較を Excel にそのまま移植**するのは、**コスト・複雑さ**の面で優先度を下げやすい。

---

## 6. 進め方の目安（フェーズ）

### 第一段階（最優先）

- **`ExcelTaskValidationConfig` / `ExcelOperationType` に載っている操作種別**と、ログの **`[Op]` の種別文字列**を一致させる。
- **セル編集以外**も、免除・不正判定に必要なものから順に、**Excel アドインでイベントを取り、`LogOperation` で追記**する（例: 印刷・ページ設定系に `SetPrintArea` 等の種別名）。

**実装済み（第1段階・ページレイアウト系）**

- `ExcelOperationType` に `SetPageOrientation` / `SetPageMargins` / `SetPageScaling` を追加（いずれも免除カテゴリは `PrintAndPage`）。`BuildAllowedForProject1` にも追加。
- VSTO `ThisAddIn`：`SheetActivate` / `WindowActivate` / ブックオープン時に **PageSetup・改ページ数・ウィンドウ枠**のスナップショットを比較し、変化があれば `SetPrintArea` / `SetPrintTitle` / `SetHeaderFooter` / `SetPageOrientation` / `SetPageMargins` / `SetPageScaling` / `SetPageBreak` / `SetFreezePanes` を `Logger.LogOperation` で追記（`ThisAddIn.LayoutMonitoring.cs`）。
- **未対応（今後の拡張）**: 図形・行/列挿入・条件付き書式・セル書式のみの変更（`EditCellFormat`）などは、別イベント／別検知が必要。

### 第二段階（進行中）

- タスクごとの **`GetExemptFlags`・`GetAllowedRanges`** を、実際の問題（プロジェクト1〜4）と突き合わせて調整する。
- **実装済み**:
    - **プロジェクト1**: 1-1〜1-7 の全タスク。
    - **プロジェクト3**: 3-1〜3-7 の全タスク（書式変更タスクは `None`＋許可範囲で厳密化、コピー＆ペーストは `RangeEdit` で緩和）。
    - **プロジェクト4**: 4-1〜4-4 の全タスク（図形・グラフ系操作のため `ShapeOrImage` を免除）。

### 第三段階（必要なら）

- 二重ログ・ノイズ・取りこぼしの調整（イベントの選定、フィルタなど）。

---

## 7. 関連コード（参照用）

| 内容 | 場所 |
|------|------|
| 免除・操作種別・許可範囲 | `Libraries/ExcelTaskValidationConfig.cs` |
| ログ読取・免除ベースの判定 | `Libraries/ExcelLogReader.cs`（`TryGetFirstNonExemptViolation` など） |
| 採点時に免除ベースの判定を適用（プロジェクト1） | `Views/ReviewPageWindow.xaml.cs`（`ApplyDestructiveValidation`） |
| VSTO ログ出力 | `ExcelAddIn1/Logger.cs`、`ExcelAddIn1/ThisAddIn.cs` |

---

## 8. 改訂履歴

| 日付 | 内容 |
|------|------|
| 2026-03-30 | 初版（方針整理・拡張フェーズの目安） |
| 2026-03-30 | 「方式A」をやめ、**免除ベースの判定**と説明に統一。用語セクションを追加。 |
| 2026-03-30 | 第1段階（ページレイアウト系 `[Op]`）を VSTO に実装。`ExcelOperationType` に `SetPageOrientation` 等を追加。 |
| 2026-04-30 | プロジェクト3の破壊的操作検知（免除フラグ・許可範囲）を実装。 |
| 2026-04-30 | プロジェクト4の破壊的操作検知（免除フラグ）を実装。 |
