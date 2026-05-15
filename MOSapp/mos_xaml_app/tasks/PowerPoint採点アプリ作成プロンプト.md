# PowerPoint採点アプリ作成プロンプト

## 概要
現在のMOS_XAML_APP（Excel採点アプリ）と同様の構造で、PowerPoint採点アプリを作成してください。UIの配置や機能はExcel版と同じですが、起動するアプリケーションとCheckerがPowerPoint用に変更されます。

## 同じUI・機能（Excel版から継承）

### 1. アプリバー（AppBarWindow）
- **配置**: 画面下部に固定表示（1920x258サイズ、Y=774位置）
- **ヘッダーバー（上部）**:
  - 左側: タイマー表示（🕐 50:00:00形式）、レビューページボタン/結果画面に戻るボタン
  - 中央: プロジェクト情報表示（例: "プロジェクト 1"）
  - 右側: 採点ボタン、一時停止ボタン、リセットボタン、次のプロジェクトボタン、アイコンボタン（↓、□）
- **メインコンテンツエリア（下部）**:
  - ナビゲーション: 前へボタン（←）、タスク番号ボタン（実際のPowerPointの問題数に合わせる）、次へボタン（→）
  - 問題文表示エリア: 現在のタスクの問題文を表示（"で囲まれた部分に下線を付けて表示、クリックでクリップボードにコピー）
  - アクションボタン: 解答済みボタン、あとで見直すボタン（フラグ）

**重要**: タスク番号ボタンの数は、Excel版の「1-7」ではなく、**PowerPoint版の実際の問題数に合わせてください**。各プロジェクトの問題数は異なる可能性があります。問題数はCSVファイルで管理されます（のちほど作成）。

### 2. レビュー画面（ReviewPageWindow）
- **レイアウト**: 全画面表示、ヘッダーにタイマーと閉じるボタン
- **コンテンツ**: 
  - プロジェクト一覧を表示
  - 各プロジェクト内のタスク一覧をテーブル形式で表示
  - タスククリックで該当タスクにジャンプ
  - 解答済み/フラグ状態を表示

### 3. 結果画面（ResultWindow）
- **レイアウト**: 中央表示、ヘッダーにタイトルと閉じるボタン
- **コンテンツ**:
  - 正答率表示
  - 間違えた問題数の表示
  - プロジェクト別の採点結果一覧
  - 各タスクの正誤判定結果（○/×表示）

## 変更が必要な箇所

### 1. アプリケーション起動
**変更前（Excel）**:
```csharp
private void LaunchAndPositionExcel(int appBarHeight)
{
    // Excel起動処理
    string[] candidates = new[]
    {
        "excel.exe",
        @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
        @"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    };
    // ...
}
```

**変更後（PowerPoint）**:
```csharp
private void LaunchAndPositionPowerPoint(int appBarHeight)
{
    // PowerPoint起動処理
    string[] candidates = new[]
    {
        "powerpnt.exe",
        @"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        @"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
    };
    // PowerPointのウィンドウクラス名は "PPFrameClass" または "PPTFrameClass" を確認
    // ...
}
```

### 2. COMオブジェクトの参照
**変更前（Excel）**:
```csharp
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
```

**変更後（PowerPoint）**:
```csharp
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using PowerPointPresentation = Microsoft.Office.Interop.PowerPoint.Presentation;
```

### 3. ファイルパスの取得
**変更前（Excel）**:
```csharp
string expectedPath = @"C:\MOSTest\Excel365\Tab1\project2.xlsx";
```

**変更後（PowerPoint）**:
```csharp
string expectedPath = @"C:\MOSTest\PowerPoint365\Tab1\project2.pptx";
// または適切なPowerPointファイルのパス
```

### 4. 問題文のJSONファイル
**変更前（Excel）**:
- `MOS演習問題文一覧.json`
- `MOS模擬試験①問題文一覧.json`
- `MOS模擬試験②問題文一覧.json`

**変更後（PowerPoint）**:
- `MOS演習問題文一覧_PowerPoint.json`（または適切な命名規則）
- `MOS模擬試験①問題文一覧_PowerPoint.json`
- `MOS模擬試験②問題文一覧_PowerPoint.json`

JSONファイルの構造は同じ（`projects`配列内に`projectId`と`tasks`配列、各`task`に`taskId`と`description`）

### 5. Checkerクラスの呼び出し
**変更前（Excel）**:
```csharp
// Libraries/Group1/ExcelChecker1_2.cs など
string libraryName = "ExcelChecker1_2";
string dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", GetGroupFolder(libraryName), $"{libraryName}.dll");
Type checkerType = assembly.GetType(libraryName);
```

**変更後（PowerPoint）**:
```csharp
// Libraries/Group1/PowerPointChecker1_2.cs など（のちほど作成）
string libraryName = "PowerPointChecker1_2";
string dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", GetGroupFolder(libraryName), $"{libraryName}.dll");
Type checkerType = assembly.GetType(libraryName);
```

### 6. ウィンドウクラス名の検索
**変更前（Excel）**:
```csharp
if (className.ToString().Contains("XLMAIN"))
{
    excelHwnd = windowHandle;
    return false;
}
```

**変更後（PowerPoint）**:
```csharp
if (className.ToString().Contains("PPFrameClass") || className.ToString().Contains("PPTFrameClass"))
{
    powerpointHwnd = windowHandle;
    return false;
}
```

### 7. プロセス名の検索
**変更前（Excel）**:
```csharp
var excelProcesses = Process.GetProcessesByName("EXCEL");
```

**変更後（PowerPoint）**:
```csharp
var powerpointProcesses = Process.GetProcessesByName("POWERPNT");
```

### 8. ファイル拡張子の処理
**変更前（Excel）**:
- `.xlsx`, `.xlsm` などのExcelファイル拡張子

**変更後（PowerPoint）**:
- `.pptx`, `.pptm` などのPowerPointファイル拡張子

### 9. アプリバー内の問題文表示
- JSONファイルから読み込む問題文はPowerPoint用の内容に変更
- 問題文の表示ロジック自体は変更不要（`SetTextWithUnderline`メソッドなど）

### 10. レビュー画面の問題文表示
- 同様にPowerPoint用のJSONファイルから読み込む
- 表示ロジックは変更不要

## 実装手順

1. **プロジェクト名の変更**
   - `MOSExcelMogiApp` → `MOSPowerPointMogiApp`（または適切な名前）
   - 名前空間も同様に変更

2. **ファイル名の変更**
   - `AppBarWindow.xaml` → そのまま（内容を変更）
   - `MainWindow.xaml` → そのまま（内容を変更）
   - `ReviewPageWindow.xaml` → そのまま（内容を変更）
   - `ResultWindow.xaml` → そのまま（内容を変更）

3. **参照の追加**
   - Microsoft.Office.Interop.PowerPoint への参照を追加
   - Excelへの参照を削除（または両方保持）

4. **Checkerクラスの準備**
   - `Libraries/Group1/PowerPointChecker1_2.cs` などを作成（のちほど実装）
   - チェックメソッド名: `CheckTask_1_2_01()`, `CheckTask_1_2_02()` など（Excel版と同じ命名規則）

5. **JSONファイルの準備**
   - PowerPoint用の問題文JSONファイルを作成
   - `References/JSON/` フォルダに配置

6. **ファイルパスの設定**
   - PowerPointファイルのパスを設定（`GetProjectFilePath`メソッドなど）

## config.jsonファイルの変更

### Assets/config.json
**変更前（Excel）**:
```json
{
  "tabs": {
    "1": {
      "projects": {
        "1": {
          "taskCount": 7,
          "library": "ExcelChecker1_1",
          "excelFile": "C:\\MOSTest\\Excel365\\Tab1\\project1.xlsx"
        }
      }
    }
  }
}
```

**変更後（PowerPoint）**:
```json
{
  "tabs": {
    "1": {
      "projects": {
        "1": {
          "taskCount": 7,
          "library": "PowerPointChecker1_1",
          "powerpointFile": "C:\\MOSTest\\PowerPoint365\\Tab1\\project1.pptx"
        }
      }
    }
  }
}
```

**重要**: 
- `excelFile` → `powerpointFile` に変更し、ファイルパスも `.xlsx` → `.pptx` に変更
- **`taskCount`の値は、Excel版の値ではなく、PowerPoint版の実際の問題数に合わせてください**。各プロジェクトごとに問題数が異なる可能性があります。
- **問題数の管理**: 各プロジェクトの問題数は、のちほどCSVファイルで管理できるようになります。CSVファイルから問題数を読み込んで`config.json`の`taskCount`を設定するか、またはCSVファイルを直接参照する実装にしてください。

## 問題数の管理（CSVファイル）

各プロジェクトの問題数は、CSVファイルで管理されます。CSVファイルの形式は以下の通りです：

```
GroupId,ProjectId,TaskCount
1,1,7
1,2,5
1,3,7
...
```

**実装方法**:
- CSVファイルを`References/CSV/`フォルダに配置（例: `PowerPoint問題数一覧.csv`）
- アプリ起動時またはプロジェクト選択時にCSVファイルから問題数を読み込む
- `config.json`の`taskCount`を動的に設定するか、CSVファイルを直接参照する実装にする
- アプリバーのタスク番号ボタンは、読み込んだ問題数に基づいて動的に表示する

## 問題文の読み込み（CSVファイル対応）

模試①（GroupId=2）など、CSVファイルから問題文を読み込む場合は、以下の実装を追加してください：

### 1. CSVファイルから問題文を読み込むメソッドの追加

`AppBarWindow.xaml.cs`に`LoadTasksFromCsv()`メソッドを追加：

```csharp
private void LoadTasksFromCsv(int groupId)
{
    // CSVファイルのパス
    string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "CSV", "解答手順あり模擬試験①問題文.csv");
    
    // CSVファイルを読み込む
    var projects = new Dictionary<int, List<TaskInfo>>();
    string[] lines = File.ReadAllLines(csvPath, Encoding.UTF8);
    
    // ヘッダー行をスキップ（1行目）
    for (int i = 1; i < lines.Length; i++)
    {
        string line = lines[i].Trim();
        if (string.IsNullOrEmpty(line))
            continue;
        
        // CSVのパース（カンマ区切り、引用符内のカンマを考慮）
        string[] fields = ParseCsvLine(line);
        if (fields.Length < 3)
            continue;
        
        // グループ,プロジェクト,問題文,解答操作
        if (int.TryParse(fields[0].Trim(), out int csvGroupId) && 
            int.TryParse(fields[1].Trim(), out int projectId) &&
            csvGroupId == groupId)
        {
            string description = fields[2].Trim();
            if (string.IsNullOrEmpty(description))
                continue;
            
            if (!projects.ContainsKey(projectId))
            {
                projects[projectId] = new List<TaskInfo>();
            }
            
            int taskId = projects[projectId].Count + 1;
            projects[projectId].Add(new TaskInfo
            {
                TaskId = taskId,
                Description = description
            });
        }
    }
    
    // ProjectData形式に変換
    _projectData = new ProjectData
    {
        Projects = projects.Select(kvp => new ProjectInfo
        {
            ProjectId = kvp.Key,
            Tasks = kvp.Value
        }).ToList()
    };
    
    // 現在のプロジェクトのタスクを取得
    var currentProject = _projectData.Projects.Find(p => p.ProjectId == _currentProjectId);
    if (currentProject != null)
    {
        _tasks = currentProject.Tasks;
        // プロジェクトの状態を初期化
    }
}
```

### 2. CSVパーサーの実装

引用符内のカンマに対応したCSVパーサーを追加：

```csharp
private string[] ParseCsvLine(string line)
{
    var fields = new List<string>();
    bool inQuotes = false;
    StringBuilder currentField = new StringBuilder();
    
    for (int i = 0; i < line.Length; i++)
    {
        char c = line[i];
        
        if (c == '"')
        {
            if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
            {
                // エスケープされた引用符
                currentField.Append('"');
                i++; // 次の文字をスキップ
            }
            else
            {
                // 引用符の開始/終了
                inQuotes = !inQuotes;
            }
        }
        else if (c == ',' && !inQuotes)
        {
            // フィールドの区切り
            fields.Add(currentField.ToString());
            currentField.Clear();
        }
        else
        {
            currentField.Append(c);
        }
    }
    
    // 最後のフィールドを追加
    fields.Add(currentField.ToString());
    
    return fields.ToArray();
}
```

### 3. LoadTasks()メソッドの修正

GroupIdに応じてCSVまたはJSONから読み込むように修正：

```csharp
private void LoadTasks()
{
    int groupId = 1; // デフォルト
    if (_viewModel?.CurrentProject != null)
    {
        string groupStr = _viewModel.CurrentProject.Group.Replace("Group ", "");
        int.TryParse(groupStr, out groupId);
    }
    
    // 模試①（GroupId=2）の場合はCSVファイルから読み込む
    if (groupId == 2)
    {
        LoadTasksFromCsv(groupId);
    }
    else
    {
        // その他の場合はJSONファイルから読み込む
        LoadTasksFromJson(groupId);
    }
}
```

### 4. .csprojにCSVファイルのコピー設定を追加

`MOSPowerPointMogiApp.csproj`に以下を追加：

```xml
<Content Include="References\CSV\解答手順あり模擬試験①問題文.csv">
  <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
</Content>
```

### 5. プロジェクト変更時の処理

`OnCurrentProjectChanged`メソッドで、プロジェクトが変更されたときに`LoadTasks()`が呼ばれるようにします：

```csharp
private void OnCurrentProjectChanged(object sender, EventArgs e)
{
    if (_viewModel.CurrentProject != null)
    {
        int newProjectId = _viewModel.CurrentProject.ProjectNumber;
        
        if (newProjectId != _currentProjectId)
        {
            _currentProjectId = newProjectId;
            LoadTasks(); // 問題文を再読み込み
            _currentTaskId = 1; // 最初のタスクにリセット
            UpdateTaskDisplay();
        }
    }
}
```

### 6. ReviewPageWindow.xaml.csの対応

レビュー画面でも同様に、GroupId=2の場合はCSVから読み込むように`LoadProjectsFromCsv()`メソッドを追加してください。

## 注意事項

- **問題数の設定**: PowerPoint版の各プロジェクトの問題数は、Excel版と異なる可能性があります。`config.json`の`taskCount`や、アプリバーのタスク番号ボタンの数は、**PowerPoint版の実際の問題数に合わせて設定してください**。問題数はCSVファイルで管理されるため、CSVファイルから読み込んだ値を使用してください。
- PowerPointCheckerはのちほど作成するため、Checkerクラスの実装は後回しでOK
- UIの配置やボタンの機能はExcel版と完全に同じにする
- タイマー機能、レビューページ機能、結果画面機能はすべて同じロジックを使用
- アプリバーのウィンドウ位置やサイズもExcel版と同じ（1920x258、Y=774）
- PowerPointのウィンドウクラス名は環境によって異なる可能性があるため、複数のパターンを確認する
- 別フォルダで作成するため、Excel版のコードをコピーして変更を適用してください
- 名前空間の変更は一括で行うことを推奨（Visual Studioの「検索と置換」機能を使用）

