# ExcelChecker → WordChecker 変換チェックリスト

## 1. 名前空間・参照の変更

### 1.1 Interop参照の変更
- [ ] `Microsoft.Office.Interop.Excel` → `Microsoft.Office.Interop.Word`
- [ ] プロジェクトファイル（.csproj）の参照を変更
  - `Microsoft.Office.Interop.Excel` → `Microsoft.Office.Interop.Word`
  - PublicKeyToken: `71e9bce111e9429c` (同じ可能性あり、要確認)

### 1.2 usingディレクティブの変更
- [ ] すべての`.cs`ファイルで `using Microsoft.Office.Interop.Excel;` → `using Microsoft.Office.Interop.Word;`

## 2. クラス名・インターフェース名の変更

### 2.1 インターフェース
- [ ] `IExcelCheckerService` → `IWordCheckerService`
- [ ] `IExcelCheckerRepository` → `IWordCheckerRepository`

### 2.2 実装クラス
- [ ] `ExcelCheckerService` → `WordCheckerService`
- [ ] `ExcelCheckerRepository` → `WordCheckerRepository`

### 2.3 チェッカークラス
- [ ] `ExcelChecker1_1` → `WordChecker1_1`
- [ ] `ExcelChecker1_2` → `WordChecker1_2`
- [ ] ... (すべてのExcelCheckerクラスをWordCheckerに変更)

## 3. メソッド名・変数名の変更

### 3.1 サービスメソッド
- [ ] `CheckExcel()` → `CheckWord()`
- [ ] `ExecuteCheck()` メソッド内の処理をWord用に変更

### 3.2 チェッカーメソッド
- [ ] `GetCurrentExcelFilePath()` → `GetCurrentWordFilePath()`
- [ ] メソッド内のExcelオブジェクト操作をWordオブジェクト操作に変更

## 4. オブジェクトモデルの変更

### 4.1 Applicationオブジェクト
- [ ] `Application excelApp` → `Application wordApp`
- [ ] `Microsoft.Office.Interop.Excel.Application` → `Microsoft.Office.Interop.Word.Application`

### 4.2 ドキュメントオブジェクト
- [ ] `Workbook` → `Document`
- [ ] `Workbooks` → `Documents`
- [ ] `Workbook.Open()` → `Documents.Open()`

### 4.3 シート・セル操作
- [ ] `Worksheet` → 削除（Wordにはシート概念がない）
- [ ] `Range` → `Range` (Wordでも使用可能だが、使用方法が異なる)
- [ ] `Cells` → WordのRange操作に変更
- [ ] `Worksheet.Cells[row, col]` → WordのRange操作に変更

### 4.4 その他のオブジェクト
- [ ] `Sheets` → 削除
- [ ] `PageSetup` → Wordの`PageSetup`（異なるプロパティ）
- [ ] `PrintArea` → Wordの印刷範囲設定に変更

## 5. ファイルパス・拡張子の変更

### 5.1 ファイル拡張子
- [ ] `.xlsx` → `.docx`
- [ ] `.xls` → `.doc`
- [ ] ファイル検索ロジックで拡張子を変更

### 5.2 ファイルパス
- [ ] `C:\MOSTest\Excel365\` → `C:\MOSTest\Word365\`
- [ ] プロジェクトファイルのパスを変更

## 6. プロセス名・ウィンドウ検索の変更

### 6.1 プロセス名
- [ ] `EXCEL` → `WINWORD`
- [ ] `excel.exe` → `winword.exe`
- [ ] `FindTopWindowByProcessName("EXCEL")` → `FindTopWindowByProcessName("WINWORD")`

### 6.2 実行パス
- [ ] Excel実行パス候補 → Word実行パス候補
  - `excel.exe` → `winword.exe`
  - `EXCEL.EXE` → `WINWORD.EXE`
  - Office16のパスも変更

## 7. ファイル名・フォルダ名の変更

### 7.1 ソースファイル
- [ ] `ExcelChecker1_1.cs` → `WordChecker1_1.cs`
- [ ] `ExcelCheckerService.cs` → `WordCheckerService.cs`
- [ ] `ExcelCheckerRepository.cs` → `WordCheckerRepository.cs`
- [ ] `IExcelCheckerService.cs` → `IWordCheckerService.cs`
- [ ] `IExcelCheckerRepository.cs` → `IWordCheckerRepository.cs`

### 7.2 プロジェクトファイル
- [ ] 各チェッカーの`.csproj`ファイル名を変更
- [ ] プロジェクトファイル内の参照を更新

### 7.3 DLL名
- [ ] `ExcelChecker1_1.dll` → `WordChecker1_1.dll`
- [ ] ビルドスクリプト（build_dlls.bat等）を更新

## 8. 文字列リテラルの変更

### 8.1 ライブラリ名
- [ ] `"ExcelChecker{groupId}_{projectId}"` → `"WordChecker{groupId}_{projectId}"`
- [ ] `libraryName.Replace("ExcelChecker", "")` → `libraryName.Replace("WordChecker", "")`

### 8.2 エラーメッセージ
- [ ] "Excelファイル" → "Wordファイル"
- [ ] "Excelが既に開いています" → "Wordが既に開いています"

### 8.3 名前空間文字列
- [ ] `"Libraries.Group{groupId}.ExcelChecker..."` → `"Libraries.Group{groupId}.WordChecker..."`

## 9. プロジェクト構造の変更

### 9.1 フォルダ構造
- [ ] `Libraries/Group1/ExcelChecker1_*.cs` → `Libraries/Group1/WordChecker1_*.cs`
- [ ] 必要に応じてフォルダ構造を確認

### 9.2 ビルドスクリプト
- [ ] `build_dlls.bat` をWord用に更新
- [ ] DLL参照パスを更新
- [ ] Interop.Wordの参照パスを更新

## 10. 設定ファイルの変更

### 10.1 config.json
- [ ] `excelFile` → `wordFile` (または `documentFile`)
- [ ] ファイルパスをWord用に更新
- [ ] ライブラリ名をWordCheckerに変更

### 10.2 App.config
- [ ] 必要に応じて設定を更新

## 11. その他の注意事項

### 11.1 Word固有の操作
- [ ] WordにはWorksheetがないため、ドキュメント全体を操作
- [ ] セル操作（Cells）の代わりにRange操作を使用
- [ ] 印刷設定、ページ設定などのAPIが異なる
- [ ] コメント（Comment）の操作が異なる可能性

### 11.2 互換性
- [ ] Wordのバージョン（Word 2016, 2019, 365等）を確認
- [ ] Interop.Wordのバージョンを確認

### 11.3 テスト
- [ ] すべてのチェッカーメソッドをWord用に実装し直す
- [ ] WordオブジェクトモデルのAPIを確認
- [ ] エラーハンドリングを確認

## 12. 具体的な変更例

### 12.1 Excel → Word オブジェクト変換例
```csharp
// Excel版
Application excelApp = new Application();
Workbook workbook = excelApp.Workbooks.Open(filePath);
Worksheet worksheet = workbook.Worksheets[1];
Range range = worksheet.Cells[1, 1];

// Word版
Application wordApp = new Application();
Document document = wordApp.Documents.Open(filePath);
Range range = document.Range(0, 0);
```

### 12.2 プロセス検索の変更例
```csharp
// Excel版
hWnd = FindTopWindowByProcessName("EXCEL", 6000);

// Word版
hWnd = FindTopWindowByProcessName("WINWORD", 6000);
```

