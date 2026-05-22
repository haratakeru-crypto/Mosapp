# Word/PowerPoint版アプリ作成ガイド

このドキュメントは、現在のExcelアプリケーション（MOSExcelMogiApp）と同じUI・機能を持つWord版またはPowerPoint版のアプリケーションを作成する際のプロンプト例を提供します。

---

## 📋 基本情報

### 異なる点
1. **問題文**: Word（またはPowerPoint）用の問題文JSONファイルを使用
2. **Checkerクラス**: WordChecker（またはPowerPointChecker）クラスを使用
   - Microsoft.Office.Interop.Wordを参照
   - GetCurrentWordFilePathメソッドで開いているWordファイルを取得
   - CheckTask_X_X_XX メソッドでWord文書の設定をチェック

### 同じ点
1. UIレイアウト（AppBarWindow.xaml）
2. タイマー、タスクナビゲーション、ボタン配置
3. 一時停止、リセット、次のプロジェクト、レビューページなどの機能
4. プロジェクト管理とタスク状態管理
5. レビューページのUIと機能（問題文の内容のみ異なる）
6. 結果画面のUIと機能（問題文の内容のみ異なる）
7. 問題文の下線表示とクリップボードコピー機能
8. レビューページボタンと結果画面に戻るボタンの切り替え機能

### アプリバー内の問題文と問題番号について

**重要な注意事項**: アプリバー内の各プロジェクトの問題番号と問題文は、読み込む問題文ファイル（CSVまたはJSON）の内容に基づいて動的に表示されます。

- **問題番号**: プロジェクト番号とタスク番号の表示形式は同じ（例: "プロジェクト 1"、"タスク 1"）
- **問題文**: 各Officeアプリケーション（Excel/Word/PowerPoint）で異なる内容を表示
- **問題文の読み込み**: `AppBarWindow.xaml.cs`の`LoadTasks()`メソッドで、GroupIdに応じて適切なファイルから読み込む
  - Excel版: `References/CSV/解答手順あり模擬試験①問題文.csv` または `References/JSON/MOS模擬アプリ問題文一覧.json`
  - Word版: `References/CSV/解答手順あり模擬試験①問題文_Word.csv` または `References/JSON/MOSWord問題文一覧.json`
  - PowerPoint版: `References/CSV/解答手順あり模擬試験①問題文_PowerPoint.csv` または `References/JSON/MOSPowerPoint問題文一覧.json`

---

## 📐 ウィンドウ配置とサイズ設定

### Excel版のウィンドウ配置

現在のExcelアプリケーションでは、以下のウィンドウ配置が使用されています：

#### 画面全体の構成
- **画面サイズ**: 1920px × 1032px（想定）
- **Excelウィンドウ**: 画面の上3/4（774px）
- **アプリバー**: 画面の下1/4（258px）

#### Excelウィンドウの配置（PositionExcelWindowメソッド）

現在の実装では、`PositionExcelWindow()`メソッドでExcelウィンドウを配置します：

```csharp
private void PositionExcelWindow()
{
    try
    {
        // 実行中のExcelプロセスを取得
        var excelProcesses = Process.GetProcessesByName("EXCEL");
        if (excelProcesses.Length == 0) return;

        Process excelProcess = excelProcesses[0];
        
        // Excelのメインウィンドウハンドルを取得（リトライロジック）
        IntPtr excelHwnd = IntPtr.Zero;
        uint processId = (uint)excelProcess.Id;
        int retryCount = 0;
        const int maxRetries = 20; // 最大20回リトライ（10秒）
        
        while (excelHwnd == IntPtr.Zero && retryCount < maxRetries)
        {
            // プロセスIDからウィンドウハンドルを検索
            EnumWindows((windowHandle, lParam) =>
            {
                GetWindowThreadProcessId(windowHandle, out uint windowProcessId);
                if (windowProcessId == processId)
                {
                    // Excelのメインウィンドウを特定（クラス名で判定）
                    StringBuilder className = new StringBuilder(256);
                    GetClassName(windowHandle, className, className.Capacity);
                    if (className.ToString().Contains("XLMAIN"))
                    {
                        excelHwnd = windowHandle;
                        return false; // 見つかったので列挙を停止
                    }
                }
                return true; // 続行
            }, IntPtr.Zero);
            
            if (excelHwnd == IntPtr.Zero)
            {
                Thread.Sleep(500); // 500ms待機してリトライ
                retryCount++;
            }
        }
        
        // ウィンドウハンドルが見つかった場合、リサイズ
        if (excelHwnd != IntPtr.Zero)
        {
            // Excelのウィンドウの境界線サイズを取得
            GetWindowRect(excelHwnd, out RECT excelWindowRect);
            GetClientRect(excelHwnd, out RECT excelClientRect);
            
            int excelBorderWidth = (excelWindowRect.right - excelWindowRect.left) - excelClientRect.right;
            int excelBorderHeight = (excelWindowRect.bottom - excelWindowRect.top) - excelClientRect.bottom;
            
            // Excelのウィンドウを左上 X=0, Y=0、右下 X=1920, Y=774 にリサイズ
            // 高さ: 258 * 3 = 774 (1032 / 4 * 3)
            // 境界線を考慮して位置を調整（マージンをゼロにする）
            int excelX = -excelBorderWidth / 2;
            int excelY = -excelBorderHeight / 2;
            int excelWidth = 1920 + excelBorderWidth;
            int excelHeight = 774 + excelBorderHeight;
            
            MoveWindow(excelHwnd, excelX, excelY, excelWidth, excelHeight, true);
        }
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Error positioning Excel window: {ex.Message}");
    }
}
```

**Excelウィンドウの仕様**:
- **位置**: X=0, Y=0（画面左上）
- **サイズ**: 1920px × 774px
- **ウィンドウクラス名**: `XLMAIN`（Excelのメインウィンドウ）
- **プロセス名**: `EXCEL`
- **リトライロジック**: 最大20回、500ms間隔でウィンドウハンドルを検索

#### アプリバーの配置（SetWindowPositionメソッド）

現在の実装では、`SetWindowPosition()`メソッドでアプリバーを配置します：

```csharp
private void SetWindowPosition()
{
    // Excelのウィンドウを配置
    PositionExcelWindow();

    // ウィンドウハンドルを取得
    IntPtr hWnd = new WindowInteropHelper(this).Handle;
    if (hWnd == IntPtr.Zero)
    {
        // ハンドルが取得できない場合はWPFプロパティで設定
        this.Width = 1920;
        this.Height = 258;
        this.Left = 0;
        this.Top = 774;
        this.Topmost = true;
        return;
    }

    // 現在のウィンドウサイズを取得して境界線のサイズを計算
    GetWindowRect(hWnd, out RECT windowRect);
    GetClientRect(hWnd, out RECT clientRect);

    int borderWidth = (windowRect.right - windowRect.left) - clientRect.right;
    int borderHeight = (windowRect.bottom - windowRect.top) - clientRect.bottom;

    // アプリバーのウィンドウを1920x258サイズで、Excelの下に配置
    // 高さ: 258 (1032 / 4)
    // 位置: Y=774 (Excelの下)
    // 境界線を考慮して位置を調整
    int x = -borderWidth / 2; // 左側の境界線を考慮
    int y = 774 - borderHeight / 2; // 上側の境界線を考慮（Excelの下）
    int width = 1920 + borderWidth; // 境界線を含めた幅
    int height = 258 + borderHeight; // 境界線を含めた高さ

    MoveWindow(hWnd, x, y, width, height, true);
    
    // ウィンドウを最前面に表示
    this.Topmost = true;
}
```

**アプリバーの仕様**:
- **位置**: X=0, Y=774（Excelウィンドウの下）
- **サイズ**: 1920px × 258px
- **ウィンドウスタイル**: `WindowStyle="None"`, `AllowsTransparency="True"`, `Topmost="True"`
- **背景**: `#F5F5F5`（角丸8px、ドロップシャドウ付き）
- **配置タイミング**: `OnContentRendered`イベントで`Dispatcher.BeginInvoke`を使用して配置

**アプリバー内の問題文と問題番号について**:
- **問題文の表示**: `TaskDescriptionTextBlock`（XAML内のTextBlock）に現在のタスクの問題文を表示
- **プロジェクト情報の表示**: `ProjectInfoTextBlock`（XAML内のTextBlock）に現在のプロジェクト名を表示
- **問題文の読み込み**: `LoadTasks()`メソッドで、GroupIdに応じて以下のファイルから読み込む
  - GroupId=2（模試①）: `References/CSV/解答手順あり模擬試験①問題文.csv`
  - その他: `References/JSON/MOS演習問題文一覧.json` など（GroupIdに応じて異なるJSONファイル）
- **Word/PowerPoint版での変更点**: 
  - 問題文の内容は各Officeアプリケーション（Excel/Word/PowerPoint）で異なるため、対応する問題文ファイル（CSVまたはJSON）を用意する必要がある
  - アプリバー内の各プロジェクトの問題番号と問題文は、読み込む問題文ファイルの内容に基づいて動的に表示される
  - プロジェクト番号とタスク番号の表示形式は同じ（例: "プロジェクト 1"、"タスク 1"）だが、問題文の内容は各Officeで異なる

### 問題文の表示機能（SetTextWithUnderline）

現在の実装では、問題文内の"で囲まれた部分に下線を付けて表示し、クリックするとクリップボードにコピーする機能が実装されています。

#### SetTextWithUnderlineメソッド

```csharp
/// <summary>
/// テキスト内の"で囲まれた部分に下線を付けてTextBlockに設定します
/// "自体は表示せず、その中のテキストだけに下線を付けます
/// </summary>
private void SetTextWithUnderline(TextBlock textBlock, string text)
{
    if (string.IsNullOrEmpty(text))
    {
        textBlock.Text = string.Empty;
        return;
    }

    textBlock.Inlines.Clear();
    textBlock.Text = string.Empty; // TextプロパティをクリアしてInlinesを使用
    
    // "で囲まれた部分を検索して下線を付ける
    int startIndex = 0;
    bool foundQuotes = false;
    
    while (startIndex < text.Length)
    {
        // "の開始位置を検索（半角ダブルクォート）
        int quoteStart = text.IndexOf('"', startIndex);
        if (quoteStart == -1)
        {
            // "が見つからない場合は残りのテキストをそのまま追加
            if (startIndex < text.Length)
            {
                string remainingText = text.Substring(startIndex);
                if (!string.IsNullOrEmpty(remainingText))
                {
                    textBlock.Inlines.Add(new Run(remainingText));
                }
            }
            break;
        }
        
        foundQuotes = true;
        
        // "の前のテキストを追加
        if (quoteStart > startIndex)
        {
            textBlock.Inlines.Add(new Run(text.Substring(startIndex, quoteStart - startIndex)));
        }
        
        // "の終了位置を検索
        int quoteEnd = text.IndexOf('"', quoteStart + 1);
        if (quoteEnd == -1)
        {
            // "が見つからない場合は残りをそのまま追加
            textBlock.Inlines.Add(new Run(text.Substring(quoteStart)));
            break;
        }
        
        // "で囲まれた部分のテキスト（"を除く）に下線を付けて追加
        string quotedText = text.Substring(quoteStart + 1, quoteEnd - quoteStart - 1);
        
        var run = new Run(quotedText);
        run.TextDecorations = TextDecorations.Underline;
        run.Cursor = Cursors.Hand; // マウスカーソルをポインターに変更
        run.MouseDown += (sender, e) => OnUnderlinedTextClick(quotedText, e);
        textBlock.Inlines.Add(run);
        
        startIndex = quoteEnd + 1;
    }
    
    // "が見つからなかった場合は通常のテキストとして設定
    if (!foundQuotes)
    {
        textBlock.Inlines.Clear();
        textBlock.Text = text;
    }
}

/// <summary>
/// 下線付きテキストがクリックされたときに呼び出されます
/// クリップボードにテキストをコピーします
/// </summary>
private void OnUnderlinedTextClick(string text, MouseButtonEventArgs e)
{
    try
    {
        Clipboard.SetText(text);
        System.Diagnostics.Debug.WriteLine($"[OnUnderlinedTextClick] Copied to clipboard: {text}");
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"[OnUnderlinedTextClick] Error copying to clipboard: {ex.Message}");
        MessageBox.Show($"クリップボードへのコピーに失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}
```

**機能の説明**:
- 問題文内の"で囲まれた部分に下線を付けて表示
- 下線付きテキストにマウスを合わせるとカーソルがポインターに変わる
- 下線付きテキストをクリックすると、そのテキストがクリップボードにコピーされる
- "自体は表示されず、その中のテキストだけが下線付きで表示される

**クリップボードにコピーされるテキストの例**:

`References/Fix/問題文_抽出結果.csv`を参考に、実際にコピーされるテキストの例を以下に示します：

- 数値や単位: 「"単位：円"」「"10000"」「"5000"」「"5/10"」
- テキスト情報: 「"最新の商品情報"」「"有楽町店の売上グラフ"」「"売上構成比"」「"売上"」「"氏名"」「"在庫を補充"」
- URL: 「"https://rabbitway.jp/service_mos"」
- メールアドレス: 「"@rabbit.ac.jp"」
- 条件分岐の結果: 「"あり"」「"なし"」

**使用例**:
問題文が「シート［売上一覧］のセル【G4】にメモを挿入し「"最新の商品情報"」と入力します。」の場合：
- 「最新の商品情報」の部分に下線が表示される
- この部分をクリックすると、「最新の商品情報」（"は除く）がクリップボードにコピーされる
- ユーザーはコピーしたテキストをOfficeアプリケーションに貼り付けて使用できる

**Word/PowerPoint版での実装**:
- Excel版と同じ`SetTextWithUnderline`メソッドを使用
- 問題文の内容は各Officeで異なるが、表示機能は同じ
- Word版やPowerPoint版の問題文にも同様に"で囲まれた部分が含まれる場合、同じようにクリップボードにコピーできる

### Word版のウィンドウ配置

Word版でも同じ配置を使用します。変更点は以下の通りです：

#### Wordウィンドウの配置（PositionWordWindowメソッド）
```csharp
private void PositionWordWindow()
{
    try
    {
        // 実行中のWordプロセスを取得
        var wordProcesses = Process.GetProcessesByName("WINWORD");
        if (wordProcesses.Length == 0) return;

        Process wordProcess = wordProcesses[0];
        
        // Wordのメインウィンドウハンドルを取得
        IntPtr wordHwnd = IntPtr.Zero;
        uint processId = (uint)wordProcess.Id;
        int retryCount = 0;
        const int maxRetries = 20;
        
        while (wordHwnd == IntPtr.Zero && retryCount < maxRetries)
        {
            EnumWindows((windowHandle, lParam) =>
            {
                GetWindowThreadProcessId(windowHandle, out uint windowProcessId);
                if (windowProcessId == processId)
                {
                    // Wordのメインウィンドウを特定（クラス名で判定）
                    StringBuilder className = new StringBuilder(256);
                    GetClassName(windowHandle, className, className.Capacity);
                    if (className.ToString().Contains("OpusApp")) // Wordのウィンドウクラス名
                    {
                        wordHwnd = windowHandle;
                        return false;
                    }
                }
                return true;
            }, IntPtr.Zero);
            
            if (wordHwnd == IntPtr.Zero)
            {
                Thread.Sleep(500);
                retryCount++;
            }
        }
        
        if (wordHwnd != IntPtr.Zero)
        {
            // Wordのウィンドウの境界線サイズを取得
            GetWindowRect(wordHwnd, out RECT wordWindowRect);
            GetClientRect(wordHwnd, out RECT wordClientRect);
            
            int wordBorderWidth = (wordWindowRect.right - wordWindowRect.left) - wordClientRect.right;
            int wordBorderHeight = (wordWindowRect.bottom - wordWindowRect.top) - wordClientRect.bottom;
            
            // Wordのウィンドウを左上 X=0, Y=0、右下 X=1920, Y=774 にリサイズ
            int wordX = -wordBorderWidth / 2;
            int wordY = -wordBorderHeight / 2;
            int wordWidth = 1920 + wordBorderWidth;
            int wordHeight = 774 + wordBorderHeight;
            
            MoveWindow(wordHwnd, wordX, wordY, wordWidth, wordHeight, true);
        }
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Error positioning Word window: {ex.Message}");
    }
}
```

**Wordウィンドウの仕様**:
- **位置**: X=0, Y=0（画面左上、Excelと同じ）
- **サイズ**: 1920px × 774px（Excelと同じ）
- **ウィンドウクラス名**: `OpusApp`（Wordのメインウィンドウ）
- **プロセス名**: `WINWORD`

#### アプリバーの配置（Word版も同じ）
```csharp
private void SetWindowPosition()
{
    // Wordのウィンドウを配置
    PositionWordWindow(); // Excel版では PositionExcelWindow()

    // アプリバーの配置（Excel版と同じ）
    IntPtr hWnd = new WindowInteropHelper(this).Handle;
    if (hWnd == IntPtr.Zero)
    {
        this.Width = 1920;
        this.Height = 258;
        this.Left = 0;
        this.Top = 774;
        this.Topmost = true;
        return;
    }

    GetWindowRect(hWnd, out RECT windowRect);
    GetClientRect(hWnd, out RECT clientRect);

    int borderWidth = (windowRect.right - windowRect.left) - clientRect.right;
    int borderHeight = (windowRect.bottom - windowRect.top) - clientRect.bottom;

    int x = -borderWidth / 2;
    int y = 774 - borderHeight / 2;
    int width = 1920 + borderWidth;
    int height = 258 + borderHeight;

    MoveWindow(hWnd, x, y, width, height, true);
    this.Topmost = true;
}
```

**アプリバーの仕様（Word版も同じ）**:
- **位置**: X=0, Y=774（Wordウィンドウの下）
- **サイズ**: 1920px × 258px
- **ウィンドウスタイル**: `WindowStyle="None"`, `AllowsTransparency="True"`, `Topmost="True"`

### PowerPoint版のウィンドウ配置

PowerPoint版でも同じ配置を使用します。変更点は以下の通りです：

#### PowerPointウィンドウの配置（PositionPowerPointWindowメソッド）
```csharp
private void PositionPowerPointWindow()
{
    try
    {
        // 実行中のPowerPointプロセスを取得
        var pptProcesses = Process.GetProcessesByName("POWERPNT");
        if (pptProcesses.Length == 0) return;

        Process pptProcess = pptProcesses[0];
        
        // PowerPointのメインウィンドウハンドルを取得
        IntPtr pptHwnd = IntPtr.Zero;
        uint processId = (uint)pptProcess.Id;
        int retryCount = 0;
        const int maxRetries = 20;
        
        while (pptHwnd == IntPtr.Zero && retryCount < maxRetries)
        {
            EnumWindows((windowHandle, lParam) =>
            {
                GetWindowThreadProcessId(windowHandle, out uint windowProcessId);
                if (windowProcessId == processId)
                {
                    // PowerPointのメインウィンドウを特定（クラス名で判定）
                    StringBuilder className = new StringBuilder(256);
                    GetClassName(windowHandle, className, className.Capacity);
                    if (className.ToString().Contains("PPTFrameClass")) // PowerPointのウィンドウクラス名
                    {
                        pptHwnd = windowHandle;
                        return false;
                    }
                }
                return true;
            }, IntPtr.Zero);
            
            if (pptHwnd == IntPtr.Zero)
            {
                Thread.Sleep(500);
                retryCount++;
            }
        }
        
        if (pptHwnd != IntPtr.Zero)
        {
            // PowerPointのウィンドウの境界線サイズを取得
            GetWindowRect(pptHwnd, out RECT pptWindowRect);
            GetClientRect(pptHwnd, out RECT pptClientRect);
            
            int pptBorderWidth = (pptWindowRect.right - pptWindowRect.left) - pptClientRect.right;
            int pptBorderHeight = (pptWindowRect.bottom - pptWindowRect.top) - pptClientRect.bottom;
            
            // PowerPointのウィンドウを左上 X=0, Y=0、右下 X=1920, Y=774 にリサイズ
            int pptX = -pptBorderWidth / 2;
            int pptY = -pptBorderHeight / 2;
            int pptWidth = 1920 + pptBorderWidth;
            int pptHeight = 774 + pptBorderHeight;
            
            MoveWindow(pptHwnd, pptX, pptY, pptWidth, pptHeight, true);
        }
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Error positioning PowerPoint window: {ex.Message}");
    }
}
```

**PowerPointウィンドウの仕様**:
- **位置**: X=0, Y=0（画面左上、Excel/Wordと同じ）
- **サイズ**: 1920px × 774px（Excel/Wordと同じ）
- **ウィンドウクラス名**: `PPTFrameClass`（PowerPointのメインウィンドウ）
- **プロセス名**: `POWERPNT`

### ウィンドウ配置の比較表

| 項目 | Excel版 | Word版 | PowerPoint版 |
|-----|---------|--------|--------------|
| **Officeアプリの位置** | X=0, Y=0 | X=0, Y=0 | X=0, Y=0 |
| **Officeアプリのサイズ** | 1920×774px | 1920×774px | 1920×774px |
| **アプリバーの位置** | X=0, Y=774 | X=0, Y=774 | X=0, Y=774 |
| **アプリバーのサイズ** | 1920×258px | 1920×258px | 1920×258px |
| **ウィンドウクラス名** | `XLMAIN` | `OpusApp` | `PPTFrameClass` |
| **プロセス名** | `EXCEL` | `WINWORD` | `POWERPNT` |

### 注意事項

1. **境界線の考慮**: ウィンドウの境界線（ボーダー）を考慮して位置とサイズを調整する必要があります。
2. **リトライロジック**: Officeアプリケーションのウィンドウハンドル取得には時間がかかる場合があるため、リトライロジックを実装しています（最大20回、500ms間隔）。
3. **ウィンドウクラス名の確認**: 環境によってウィンドウクラス名が異なる可能性があるため、デバッグログで実際のクラス名を確認してください。
4. **画面解像度**: 1920×1032pxを想定していますが、異なる解像度の場合は適宜調整が必要です。

---

## 🔄 リセット機能の実装

### リセットボタンの機能

アプリバーのリセットボタンをクリックすると、現在のプロジェクトファイルをテンプレートファイルから復元します。

#### リセット処理で参照されるファイル

リセット処理は、以下のファイルを参照・操作します：

**1. リセット対象ファイル（プロジェクトファイル）のパス取得（優先順位順）**

1. **CurrentProject.FilePath**（ViewModelから取得）
   - `_viewModel.CurrentProject.FilePath`で現在開いているプロジェクトファイルのパス

2. **config.jsonの`initialDataFile`**
   - `Assets/config.json`の`tabs[groupId].projects[projectId].initialDataFile`

3. **config.jsonの`excelFile`（Excel版）/ `wordFile`（Word版）/ `powerPointFile`（PowerPoint版）**
   - `Assets/config.json`の`tabs[groupId].projects[projectId].excelFile`（または`wordFile`/`powerPointFile`）

4. **Initialフォルダの固定パス**（フォールバック）
   - Excel版: `C:\MOSTest\Excel365\Tab{groupId}\Initial\project{projectId}.xlsx`
   - Word版: `C:\MOSTest\Word365\Tab{groupId}\Initial\project{projectId}.docx`
   - PowerPoint版: `C:\MOSTest\PowerPoint365\Tab{groupId}\Initial\project{projectId}.pptx`

**2. テンプレートファイルの検索（優先順位順）**

1. **動的検索**（Templatesフォルダ内）
   - Excel版: `C:\MOSTest\Excel365\Templates\Tab{groupId}`フォルダ内
   - Word版: `C:\MOSTest\Word365\Templates\Tab{groupId}`フォルダ内
   - PowerPoint版: `C:\MOSTest\PowerPoint365\Templates\Tab{groupId}`フォルダ内
   - 条件: ファイル名に`project{projectId}`を含むファイル（大文字小文字を区別しない）
   - 拡張子: `.xlsx`（Excel版）/ `.docx`（Word版）/ `.pptx`（PowerPoint版）

2. **固定パターン**（動的検索で見つからない場合）
   - パターン1: `C:\MOSTest\{OfficeApp}\Templates\Tab{groupId}\project{projectId}.{ext}`
   - パターン2: `C:\MOSTest\{OfficeApp}\Templates\project{projectId}.{ext}`
   - パターン3: `C:\MOSTest\{OfficeApp}\Templates\Tab{groupId}\Tab{groupId}_project{projectId}.{ext}`
   - `{OfficeApp}`: `Excel365`（Excel版）/ `Word365`（Word版）/ `PowerPoint365`（PowerPoint版）
   - `{ext}`: `.xlsx`（Excel版）/ `.docx`（Word版）/ `.pptx`（PowerPoint版）

**3. Initialフォルダへの保存**

リセット後、プロジェクトファイルをInitialフォルダにもコピーします：
- Excel版: `C:\MOSTest\Excel365\Tab{groupId}\Initial\project{projectId}.xlsx`
- Word版: `C:\MOSTest\Word365\Tab{groupId}\Initial\project{projectId}.docx`
- PowerPoint版: `C:\MOSTest\PowerPoint365\Tab{groupId}\Initial\project{projectId}.pptx`

#### リセット処理の実装

```csharp
// MainWindow.xaml.cs
public void ResetProject(int groupId, int projectId, bool showMessage = true)
{
    try
    {
        // リセット対象ファイルのパスを取得
        string projectFilePath = null;
        
        // 1. CurrentProjectから実際のファイルパスを取得（最優先）
        if (_viewModel?.CurrentProject != null)
        {
            projectFilePath = _viewModel.CurrentProject.FilePath;
        }
        
        // 2. config.jsonから取得
        if (string.IsNullOrEmpty(projectFilePath))
        {
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
            if (File.Exists(configPath))
            {
                string jsonContent = File.ReadAllText(configPath);
                JObject config = JObject.Parse(jsonContent);
                
                var projectConfig = config["tabs"]?[groupId.ToString()]?["projects"]?[projectId.ToString()];
                if (projectConfig != null)
                {
                    // initialDataFileを優先
                    projectFilePath = projectConfig["initialDataFile"]?.ToString();
                    
                    // initialDataFileがない場合、excelFile/wordFile/powerPointFileを使用
                    if (string.IsNullOrEmpty(projectFilePath))
                    {
                        // Excel版: excelFile
                        projectFilePath = projectConfig["excelFile"]?.ToString();
                        // Word版: wordFile
                        // projectFilePath = projectConfig["wordFile"]?.ToString();
                        // PowerPoint版: powerPointFile
                        // projectFilePath = projectConfig["powerPointFile"]?.ToString();
                    }
                    
                    // それでもない場合、Initialフォルダのパスを生成
                    if (string.IsNullOrEmpty(projectFilePath))
                    {
                        // Excel版
                        projectFilePath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\Initial\\project{projectId}.xlsx";
                        // Word版
                        // projectFilePath = $"C:\\MOSTest\\Word365\\Tab{groupId}\\Initial\\project{projectId}.docx";
                        // PowerPoint版
                        // projectFilePath = $"C:\\MOSTest\\PowerPoint365\\Tab{groupId}\\Initial\\project{projectId}.pptx";
                    }
                }
            }
        }
        
        // テンプレートファイルのパスを検索
        string templatesFolder = $"C:\\MOSTest\\Excel365\\Templates\\Tab{groupId}";
        // Word版: $"C:\\MOSTest\\Word365\\Templates\\Tab{groupId}"
        // PowerPoint版: $"C:\\MOSTest\\PowerPoint365\\Templates\\Tab{groupId}"
        
        string templatePath = null;
        string fileExtension = ".xlsx"; // Word版: ".docx", PowerPoint版: ".pptx"
        
        // Templatesフォルダ内のファイルを動的に検索
        if (Directory.Exists(templatesFolder))
        {
            var files = Directory.GetFiles(templatesFolder, $"*{fileExtension}", SearchOption.TopDirectoryOnly);
            string searchPattern = $"project{projectId}".ToLower();
            
            foreach (var file in files)
            {
                string fileName = Path.GetFileNameWithoutExtension(file).ToLower();
                if (fileName.Contains(searchPattern) || fileName == searchPattern)
                {
                    templatePath = file;
                    break;
                }
            }
        }
        
        // 固定パターンで検索
        if (string.IsNullOrEmpty(templatePath))
        {
            string[] patterns = {
                $"C:\\MOSTest\\Excel365\\Templates\\Tab{groupId}\\project{projectId}{fileExtension}",
                $"C:\\MOSTest\\Excel365\\Templates\\project{projectId}{fileExtension}",
                $"C:\\MOSTest\\Excel365\\Templates\\Tab{groupId}\\Tab{groupId}_project{projectId}{fileExtension}"
            };
            // Word版/PowerPoint版の場合も同様にパスを変更
            
            foreach (var pattern in patterns)
            {
                if (File.Exists(pattern))
                {
                    templatePath = pattern;
                    break;
                }
            }
        }
        
        // テンプレートファイルを読み取り専用で保護
        FileInfo templateFileInfo = new FileInfo(templatePath);
        if (!templateFileInfo.IsReadOnly)
        {
            templateFileInfo.IsReadOnly = true;
        }
        
        // Officeアプリケーションが開いている場合は閉じる
        CloseAllOfficeWorkbooks(); // Excel版: CloseAllExcelWorkbooks(), Word版: CloseAllWordDocuments(), PowerPoint版: CloseAllPowerPointPresentations()
        
        // テンプレートファイルをプロジェクトファイルにコピー
        File.Copy(templatePath, projectFilePath, overwrite: true);
        
        // Initialフォルダにもコピー
        string initialFolderPath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\Initial";
        // Word版: $"C:\\MOSTest\\Word365\\Tab{groupId}\\Initial"
        // PowerPoint版: $"C:\\MOSTest\\PowerPoint365\\Tab{groupId}\\Initial"
        
        string initialFilePath = Path.Combine(initialFolderPath, $"project{projectId}{fileExtension}");
        
        if (!Directory.Exists(initialFolderPath))
        {
            Directory.CreateDirectory(initialFolderPath);
        }
        
        File.Copy(projectFilePath, initialFilePath, overwrite: true);
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"[ResetProject] Error: {ex.Message}");
        if (showMessage)
        {
            MessageBox.Show($"プロジェクトリセット中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
```

#### Word/PowerPoint版での変更点

| 項目 | Excel版 | Word版 | PowerPoint版 |
|-----|---------|--------|--------------|
| **フォルダパス** | `Excel365` | `Word365` | `PowerPoint365` |
| **ファイル拡張子** | `.xlsx` | `.docx` | `.pptx` |
| **config.jsonのキー** | `excelFile` | `wordFile` | `powerPointFile` |
| **ファイル閉じる処理** | `CloseAllExcelWorkbooks()` | `CloseAllWordDocuments()` | `CloseAllPowerPointPresentations()` |
| **テンプレートフォルダ** | `C:\MOSTest\Excel365\Templates\Tab{groupId}` | `C:\MOSTest\Word365\Templates\Tab{groupId}` | `C:\MOSTest\PowerPoint365\Templates\Tab{groupId}` |
| **Initialフォルダ** | `C:\MOSTest\Excel365\Tab{groupId}\Initial` | `C:\MOSTest\Word365\Tab{groupId}\Initial` | `C:\MOSTest\PowerPoint365\Tab{groupId}\Initial` |

**重要な注意事項**:
- テンプレートファイルは読み取り専用に設定されるため、誤って編集されないように保護されます
- リセット処理中は、Officeアプリケーションで開いているファイルがすべて閉じられます
- リセット後、プロジェクトファイルとInitialフォルダの両方にファイルがコピーされます
- Word版やPowerPoint版を作成する際は、フォルダ構造とファイル拡張子を適切に変更してください

---

## 📋 レビューページの実装

### ReviewPageWindowの機能

レビューページは、全プロジェクトと全タスクを一覧表示し、任意の問題に戻ることができる機能を提供します。

#### レビューページの主な機能

1. **全プロジェクト・全タスクの一覧表示**
   - プロジェクトごとにグループ化して表示
   - 各タスクに「タスク番号」と「問題文」を表示
   - 「解答済み」と「あとで見直す」の状態を表示

2. **問題文の読み込み**
   - GroupIdに応じてCSVまたはJSONファイルから読み込む
   - GroupId=2（模試①）: `References/CSV/解答手順あり模擬試験①問題文.csv`
   - その他: `References/JSON/MOS演習問題文一覧.json` など

3. **タスクへのナビゲーション**
   - タスクをクリックすると、該当のプロジェクトとタスクに移動
   - `OnNavigateToTask`コールバックでAppBarWindowに通知
   - レビューページを閉じてAppBarWindowに戻る

4. **状態管理**
   - `_projectTaskCompletedStates`: 各プロジェクトの各タスクの「解答済み」状態
   - `_projectTaskFlaggedStates`: 各プロジェクトの各タスクの「あとで見直す」状態
   - 状態は配列で管理（TaskIdは1始まり、配列インデックスは0始まり）

5. **試験終了処理（EndExamButton_Click）**
   - 「試験終了」ボタンをクリックすると、全プロジェクトを自動採点
   - 採点結果を`ExamResultStorage`に保存
   - Officeアプリケーション（Excel/Word/PowerPoint）を閉じる
   - 結果画面（ResultWindow）を表示

**重要な注意事項**:
- **試験終了処理と全プロジェクト採点処理は、Checkerクラスが作成されていないと正常に動作しません**
- アプリバー（AppBarWindow）を作成した後に、Checkerクラス（WordChecker/PowerPointChecker）を作成する予定です
- 実装の順序: **アプリバー → Checker → 採点処理**
- Checkerクラスが未作成の状態では、採点処理は実行されますが、すべてのタスクが「×」として判定されます
- Checkerクラスを作成する前に、アプリバーとレビューページの基本機能（問題文表示、タスクナビゲーションなど）が正常に動作することを確認してください

#### 試験終了処理の実装

```csharp
private async void EndExamButton_Click(object sender, RoutedEventArgs e)
{
    try
    {
        // ボタンを無効化して再クリックを防止
        if (sender is Button button)
        {
            button.IsEnabled = false;
            button.Content = "処理中...";
        }
        
        // タイマーを停止
        _timer?.Stop();
        
        // UI更新の機会を与える
        await Task.Delay(100);
        
        // AppBarWindowを閉じる
        await Dispatcher.InvokeAsync(() => CloseAppBarWindows(), DispatcherPriority.Background);
        
        // すべてのプロジェクトの採点を実行（バックグラウンドで実行）
        System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Starting to score all projects...");
        await Task.Run(() => ScoreAllProjects());
        
        // 採点が完了したらOfficeアプリケーションを閉じる
        await Task.Run(() => CloseOfficeApplication()); // Excel版ではCloseExcelApplication()
        
        // 保存されている採点結果を取得
        var allResults = Models.ExamResultStorage.GetAllResults();
        
        // 結果画面ウィンドウを表示
        ResultWindow resultWindow = null;
        await Dispatcher.InvokeAsync(() =>
        {
            resultWindow = new ResultWindow(allResults, _groupId);
            resultWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            resultWindow.Topmost = true;
            
            // OnNavigateToTaskを設定
            resultWindow.OnNavigateToTask = (projectId, taskId) =>
            {
                // AppBarWindowを確実に表示
                var appBarWindow = Application.Current.Windows.OfType<AppBarWindow>().FirstOrDefault();
                if (appBarWindow != null)
                {
                    appBarWindow.Show();
                    appBarWindow.Activate();
                }
                
                // NavigateToTaskを呼び出す
                OnNavigateToTask(projectId, taskId);
            };
            
            resultWindow.Show();
            resultWindow.Activate();
        }, DispatcherPriority.Normal);
        
        // ReviewPageWindowを非表示にする
        await Dispatcher.InvokeAsync(() =>
        {
            this.Hide();
        }, DispatcherPriority.Normal);
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in EndExamButton_Click: {ex.Message}");
        await Dispatcher.InvokeAsync(() =>
        {
            MessageBox.Show($"試験終了処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        });
    }
}
```

#### 全プロジェクト採点処理（ScoreAllProjects）

**重要な注意事項**:
- **この処理はCheckerクラス（WordChecker/PowerPointChecker）が作成されていないと正常に動作しません**
- `ExecuteScoringForProject`メソッド内で、Checkerクラスのメソッド（例: `CheckTask_1_1_01`）を呼び出して採点を行います
- Checkerクラスが未作成の場合、`ExecuteScoringForProject`はCheckerクラスを見つけられず、すべてのタスクが「×」（false）として判定されます
- アプリバーを作成した後、Phase 4で説明するCheckerクラスを作成してから、この採点処理を実装してください

```csharp
private void ScoreAllProjects()
{
    try
    {
        // ExamResultStorageをクリア
        Models.ExamResultStorage.Clear();
        
        // config.jsonを読み込む
        string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
        if (!File.Exists(configPath))
        {
            System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] config.json not found");
            return;
        }
        
        string jsonContent = File.ReadAllText(configPath);
        JObject config = JObject.Parse(jsonContent);
        
        // 現在のグループのすべてのプロジェクトを取得
        var groupProjects = config["tabs"]?[_groupId.ToString()]?["projects"];
        if (groupProjects == null)
        {
            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] No projects found for group {_groupId}");
            return;
        }
        
        // プロジェクトリストを先に作成
        var projectList = new List<(int projectId, int taskCount, string libraryName, string filePath)>();
        foreach (var project in groupProjects)
        {
            var projectProperty = project as JProperty;
            if (projectProperty != null)
            {
                int projectId = int.Parse(projectProperty.Name);
                var projectConfig = projectProperty.Value;
                int taskCount = projectConfig["taskCount"]?.Value<int>() ?? 0;
                string libraryName = projectConfig["library"]?.ToString() ?? $"ExcelChecker{_groupId}_{projectId}";
                string filePath = GetProjectFilePath(_groupId, projectId, projectConfig);
                
                projectList.Add((projectId, taskCount, libraryName, filePath));
            }
        }
        
        // 必要なファイルを開く
        foreach (string filePath in filesToOpen)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
                
                // ファイルが開くまで待つ
                System.Threading.Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error opening file: {ex.Message}");
            }
        }
        
        // すべてのプロジェクトを順番に採点
        for (int idx = 0; idx < projectList.Count; idx++)
        {
            var project = projectList[idx];
            
            // ファイルを明示的にアクティブにする
            bool activated = ActivateOfficeFile(project.filePath); // Excel版ではActivateExcelFile
            if (!activated)
            {
                // アクティブ化に失敗した場合もfalseを保存
                var falseResults = new List<bool>();
                for (int i = 0; i < project.taskCount; i++)
                {
                    falseResults.Add(false);
                }
                Models.ExamResultStorage.SaveProjectResult(project.projectId, falseResults);
                continue;
            }
            
            // ファイルがアクティブになるまで十分に待つ
            System.Threading.Thread.Sleep(1500);
            
            // 採点を実行
            var results = ExecuteScoringForProject(project.libraryName, project.taskCount);
            
            // 採点結果を保存
            Models.ExamResultStorage.SaveProjectResult(project.projectId, results);
        }
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in ScoreAllProjects: {ex.Message}");
    }
}
```

**Word/PowerPoint版での変更点**:
- `CloseExcelApplication()` → `CloseWordApplication()` / `ClosePowerPointApplication()`
- `ActivateExcelFile()` → `ActivateWordFile()` / `ActivatePowerPointFile()`
- `ExecuteScoringForProject()`内の`libraryName`が`WordChecker`または`PowerPointChecker`になる
- その他のロジックはExcel版と同じ

#### Word/PowerPoint版での実装ポイント

1. **問題文ファイルの準備**
   - Word版: `References/CSV/解答手順あり模擬試験①問題文_Word.csv` または `References/JSON/MOSWord問題文一覧.json`
   - PowerPoint版: `References/CSV/解答手順あり模擬試験①問題文_PowerPoint.csv` または `References/JSON/MOSPowerPoint問題文一覧.json`

2. **ReviewPageWindow.xaml.csの修正**
   ```csharp
   // GroupIdに応じて適切なファイルを読み込む
   private ProjectData LoadProjectsFromCsv(int groupId)
   {
       // Word版の場合
       string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, 
           "References", "CSV", "解答手順あり模擬試験①問題文_Word.csv");
       // ...
   }
   
   private ProjectData LoadProjectsFromJson(int groupId)
   {
       // Word版の場合
       string jsonFileName = groupId switch
       {
           1 => "MOSWord演習問題文一覧.json",
           2 => "MOSWord模擬試験①問題文一覧.json",
           3 => "MOSWord模擬試験②問題文一覧.json",
           _ => "MOSWord問題文一覧.json"
       };
       // ...
   }
   ```

3. **プロジェクトタイトルとタスクタイトルの表示**
   - プロジェクトタイトル: `$"プロジェクト {project.ProjectId}"`（Excel版と同じ形式）
   - タスクタイトル: `$"タスク {task.TaskId}"`（Excel版と同じ形式）
   - 問題文: `task.Description`（各Officeで異なる内容）

4. **状態管理の実装**
   - Excel版と同じロジックを使用
   - TaskId（1始まり）と配列インデックス（0始まり）の変換に注意
   ```csharp
   int arrayIndex = task.TaskId - 1;
   bool isCompleted = completedStates[arrayIndex];
   ```

---

## 📊 結果画面の実装

### ResultWindowの機能

結果画面は、試験終了後の採点結果を表示し、間違えた問題に戻ることができる機能を提供します。

> **復習後の正答率・×数更新（PowerPoint 同等）**  
> Word 向けの推奨実装案・現状ギャップ・作業順序は [`Word_結果画面復習後更新_実装案.md`](Word_結果画面復習後更新_実装案.md) を参照。

#### 結果画面の主な機能

1. **採点結果の表示**
   - 正答率と間違えた問題数を表示
   - 各プロジェクト・各タスクの結果を〇/×で表示
   - 問題文も一緒に表示

2. **問題文の読み込み**
   - JSONファイルから読み込み（GroupIdに応じて異なるファイル）
   - GroupId=1: `References/JSON/MOS演習問題文一覧.json`
   - GroupId=2: `References/JSON/MOS模擬試験①問題文一覧.json`
   - GroupId=3: `References/JSON/MOS模擬試験②問題文一覧.json`

3. **タスクへのナビゲーション**
   - タスク（特に×のタスク）をクリックするとAppBarWindowに戻る
   - `OnNavigateToTask`コールバックを使用
   - 該当のプロジェクトとタスクに移動

4. **フィルター機能**
   - 「間違えた問題のみ表示」ボタンでフィルター
   - フィルターON/OFFを切り替え可能

#### 結果画面の実装例

```csharp
public partial class ResultWindow : Window
{
    private Dictionary<int, List<bool>> _allProjectResults;
    private int _groupId = 1;
    private List<ResultProjectInfo> _allProjects; // すべてのプロジェクトを保持
    private bool _showingWrongOnly = false; // フィルター状態
    public Action<int, int> OnNavigateToTask { get; set; } // ProjectId, TaskId

    public ResultWindow(Dictionary<int, List<bool>> allProjectResults = null, int groupId = 1)
    {
        InitializeComponent();
        _allProjectResults = allProjectResults ?? new Dictionary<int, List<bool>>();
        _groupId = groupId;
        
        // ウィンドウが読み込まれた後にデータを読み込む（非同期）
        this.Loaded += ResultWindow_Loaded;
    }

    private async void ResultWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await Task.Delay(50);
        await LoadResultsAsync();
    }

    private async Task LoadResultsAsync()
    {
        try
        {
            // 正答率を計算
            int totalTasks = 0;
            int correctTasks = 0;
            int wrongTasks = 0;

            foreach (var projectResult in _allProjectResults)
            {
                foreach (var result in projectResult.Value)
                {
                    totalTasks++;
                    if (result)
                    {
                        correctTasks++;
                    }
                    else
                    {
                        wrongTasks++;
                    }
                }
            }

            // 正答率を計算
            double correctRate = totalTasks > 0 ? (double)correctTasks / totalTasks * 100 : 100.0;
            
            // UIスレッドで更新
            await Dispatcher.InvokeAsync(() =>
            {
                CorrectRateTextBlock.Text = $"{correctRate:F1}%";
                WrongCountTextBlock.Text = $"{wrongTasks}問";
            });

            // 結果を表示（非同期で読み込む）
            await LoadProjectDataAsync();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error loading results: {ex.Message}");
        }
    }
}
```

#### タスクへのナビゲーション実装

```csharp
private async void TaskRow_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
{
    if (sender is FrameworkElement element && element.DataContext is ResultTaskInfo taskInfo)
    {
        System.Diagnostics.Debug.WriteLine($"[ResultWindow] Task clicked: ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");
        
        if (OnNavigateToTask != null && taskInfo.ProjectId > 0 && taskInfo.TaskId > 0)
        {
            try
            {
                // AppBarWindowを確実に表示してからNavigateToTaskを呼び出す
                var appBarWindow = Application.Current.Windows.OfType<AppBarWindow>().FirstOrDefault();
                
                // AppBarWindowが存在しない、または閉じられている場合は新しく作成
                if (appBarWindow == null || !appBarWindow.IsLoaded)
                {
                    // MainWindowからViewModelを取得
                    var mainWindow = Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                    if (mainWindow == null)
                    {
                        mainWindow = new MainWindow();
                        Application.Current.MainWindow = mainWindow;
                    }
                    
                    var viewModel = mainWindow.DataContext as Ui.ViewModels.MainViewModel;
                    if (viewModel == null)
                    {
                        MessageBox.Show("ViewModelが見つかりませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    
                    // 新しいAppBarWindowを作成
                    appBarWindow = new AppBarWindow(viewModel);
                }
                
                // 結果画面から来たことを記録
                appBarWindow.SetFromResultWindow(true);
                appBarWindow.SetResultWindow(this);
                
                // AppBarWindowを表示
                if (!appBarWindow.IsVisible)
                {
                    appBarWindow.Show();
                }
                appBarWindow.Activate();
                
                // UI更新の機会を与える
                await Task.Delay(50);
                
                OnNavigateToTask(taskInfo.ProjectId, taskInfo.TaskId);
                
                // 結果画面を非表示にする（閉じずに保持）
                this.Hide();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error navigating to task: {ex.Message}");
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
```

#### フィルター機能の実装

```csharp
private void ShowWrongOnlyButton_Click(object sender, RoutedEventArgs e)
{
    _showingWrongOnly = !_showingWrongOnly;
    
    if (_showingWrongOnly)
    {
        // ×の問題のみ表示
        var wrongOnlyProjects = FilterWrongTasks(_allProjects);
        ProjectsItemsControl.ItemsSource = wrongOnlyProjects;
        ShowWrongOnlyButton.Content = "全て表示";
        ShowWrongOnlyButton.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 64, 175)); // #1E40AF
    }
    else
    {
        // 全て表示
        ProjectsItemsControl.ItemsSource = _allProjects;
        ShowWrongOnlyButton.Content = "間違えた問題のみ表示";
        ShowWrongOnlyButton.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 38, 38)); // #DC2626
    }
}

private List<ResultProjectInfo> FilterWrongTasks(List<ResultProjectInfo> projects)
{
    if (projects == null) return new List<ResultProjectInfo>();
    
    var filteredProjects = new List<ResultProjectInfo>();
    
    foreach (var project in projects)
    {
        var wrongTasks = project.Tasks?.Where(t => t.ResultMark == "×").ToList();
        
        // ×のタスクがある場合のみプロジェクトを追加
        if (wrongTasks != null && wrongTasks.Count > 0)
        {
            filteredProjects.Add(new ResultProjectInfo
            {
                ProjectTitle = project.ProjectTitle,
                Tasks = wrongTasks
            });
        }
    }
    
    return filteredProjects;
}
```

**Word/PowerPoint版での実装**:
- Excel版と同じ`ResultWindow`クラスを使用
- 問題文の読み込みで、Word版は`MOSWord問題文一覧.json`、PowerPoint版は`MOSPowerPoint問題文一覧.json`を使用
- その他の機能（タスクへのナビゲーション、フィルター）は同じ

### レビューページボタンと結果画面に戻るボタンの切り替え

アプリバーには、通常時と結果画面から来た場合でボタンを切り替える機能があります。

#### 実装方法

```csharp
// AppBarWindow.xaml.cs
private bool _fromResultWindow = false; // 結果画面から来たかどうか
private ResultWindow _resultWindow = null; // 結果画面への参照

public void SetFromResultWindow(bool fromResultWindow)
{
    _fromResultWindow = fromResultWindow;
    UpdateReviewPageButtonVisibility();
}

public void SetResultWindow(ResultWindow resultWindow)
{
    _resultWindow = resultWindow;
}

private void UpdateReviewPageButtonVisibility()
{
    // レビューページボタンまたは結果画面に戻るボタンの表示を切り替え
    var reviewPageButton = FindName("ReviewPageButton") as Button;
    var returnToResultButton = FindName("ReturnToResultButton") as Button;
    
    if (reviewPageButton != null && returnToResultButton != null)
    {
        if (_fromResultWindow)
        {
            reviewPageButton.Visibility = Visibility.Collapsed;
            returnToResultButton.Visibility = Visibility.Visible;
        }
        else
        {
            reviewPageButton.Visibility = Visibility.Visible;
            returnToResultButton.Visibility = Visibility.Collapsed;
        }
    }
}

private void ReturnToResultButton_Click(object sender, RoutedEventArgs e)
{
    try
    {
        // Officeアプリケーションの保存・終了処理は非同期で実行
        // （Excel版ではExcel、Word版ではWord、PowerPoint版ではPowerPoint）
        
        // 結果画面を表示
        if (_resultWindow != null && !_resultWindow.IsVisible)
        {
            _resultWindow.Show();
            _resultWindow.Activate();
            _resultWindow.Focus();
        }
        
        // AppBarWindowを非表示にする
        this.Hide();
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"結果画面表示エラー: {ex.Message}");
        MessageBox.Show("結果画面の表示に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}
```

**Word/PowerPoint版での実装**:
- Excel版と同じ`SetFromResultWindow`と`SetResultWindow`メソッドを使用
- `ReturnToResultButton_Click`内で、Excel版では`CloseExcelApplication()`、Word版では`CloseWordApplication()`、PowerPoint版では`ClosePowerPointApplication()`を呼び出す
- その他のロジックは同じ

---

## 🎯 Phase 1: 初期設定とプロジェクト構造の確認

### プロンプト例

```
現在のExcelアプリケーション（MOSExcelMogiApp）と同じUI・機能を持つ
Word版（またはPowerPoint版）のアプリケーションを作成したいです。

異なる点：
1. 問題文: Word（またはPowerPoint）用の問題文JSONファイルを使用
2. Checkerクラス: WordChecker（またはPowerPointChecker）クラスを使用
   - Microsoft.Office.Interop.Wordを参照
   - GetCurrentWordFilePathメソッドで開いているWordファイルを取得
   - CheckTask_X_X_XX メソッドでWord文書の設定をチェック

同じ点：
1. UIレイアウト（AppBarWindow.xaml）
2. タイマー、タスクナビゲーション、ボタン配置
3. 一時停止、リセット、次のプロジェクト、レビューページなどの機能
4. プロジェクト管理とタスク状態管理

まず、現在のプロジェクト構造を確認して、
どのファイルをコピー/修正すべきか提案してください。
```

---

## 🎯 Phase 2: プロジェクトファイルとアセンブリ参照の設定

### プロンプト例

```
新しいWPFプロジェクト「MOSWordMogiApp」を作成します。

1. MOSExcelMogiApp.csprojをベースに、以下を変更：
   - AssemblyName: MOSWordMogiApp
   - RootNamespace: MOSWordMogiApp
   - Microsoft.Office.Interop.Word への参照を追加
   - Libraries\WordGroup1\WordChecker1_1.cs～WordChecker1_10.cs を追加

2. 必要なファイル構成：
   - App.xaml / App.xaml.cs
   - MainWindow.xaml / MainWindow.xaml.cs（Excelアプリと同じUI）
   - AppBarWindow.xaml / AppBarWindow.xaml.cs（Excelアプリと同じUI）
   - Views\ScoringResultDialog.xaml（採点結果ダイアログ）
   - Views\ReviewPageWindow.xaml（レビューページ）
   - Ui\ViewModels\MainViewModel.cs（採点ロジックをWordChecker用に修正）
   - Assets\config.json（Word用のプロジェクト設定）
   - References\JSON\MOSWord問題文一覧.json（Word用問題文）

プロジェクトファイルを作成してください。
```

---

## 🎯 Phase 3: 問題文ファイルの作成

### 問題文ファイルの種類

Word/PowerPoint版では、Excel版と同様に、問題文をCSVまたはJSONファイルから読み込みます。

#### CSVファイル形式（模試①の場合）

模試①（GroupId=2）の場合は、CSVファイルを使用します。

**ファイルパス**: `References\CSV\解答手順あり模擬試験①問題文_Word.csv`（Word版の場合）

**CSV形式**:
```csv
グループ,プロジェクト,問題文,解答操作
2,1,"「販売実績」シートの印刷の向きを横向きに設定します。","..."
2,1,"「販売実績」シートの印刷範囲をA1:I20に設定します。","..."
```

**Word版での変更点**:
- ファイル名に`_Word`を追加（例: `解答手順あり模擬試験①問題文_Word.csv`）
- 問題文の内容をWord用に変更（例: 「シート」→「文書」、「印刷範囲」→「余白設定」など）

#### JSONファイル形式（その他の場合）

演習や模試②の場合は、JSONファイルを使用します。

**ファイルパス**: `References\JSON\MOSWord問題文一覧.json`（Word版の場合）

### プロンプト例

```
Word用の問題文JSONファイルを作成します。

ファイル: References\JSON\MOSWord問題文一覧.json

Excelの「References\JSON\MOS模擬アプリ問題文一覧.json」と
同じ構造で、以下のWord用問題文を設定してください：

プロジェクト1-1:
- タスク1: 「文書のページ設定を余白：狭い、に設定します。」
- タスク2: 「文書の向きを横向きに変更します。」
- タスク3: 「段落の行間を1.5行に設定します。」
- タスク4: 「見出し1のスタイルを適用します。」
- タスク5: 「ページ番号を挿入します。」
- タスク6: 「目次を挿入します。」
- タスク7: 「透かし「社外秘」を挿入します。」

（実際の問題文は後で設定するので、サンプルでOK）

JSONファイルを作成してください。
```

### JSON構造例

```json
{
  "projects": [
    {
      "projectId": 1,
      "tasks": [
        {
          "taskId": 1,
          "description": "文書のページ設定を余白：狭い、に設定します。"
        },
        {
          "taskId": 2,
          "description": "文書の向きを横向きに変更します。"
        }
      ]
    }
  ]
}
```

### AppBarWindow.xaml.csでの問題文読み込み

Word/PowerPoint版では、`LoadTasks()`メソッドと`LoadTasksFromCsv()`/`LoadTasksFromJson()`メソッドを修正して、適切なファイルを読み込むようにします。

```csharp
private void LoadTasksFromCsv(int groupId)
{
    // Word版の場合
    string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, 
        "References", "CSV", "解答手順あり模擬試験①問題文_Word.csv");
    // ...
}

private void LoadTasksFromJson(int groupId)
{
    // Word版の場合
    string jsonFileName = groupId switch
    {
        1 => "MOSWord演習問題文一覧.json",
        2 => "MOSWord模擬試験①問題文一覧.json",
        3 => "MOSWord模擬試験②問題文一覧.json",
        _ => "MOSWord問題文一覧.json"
    };
    // ...
}
```

### ReviewPageWindow.xaml.csでの問題文読み込み

レビューページでも同様に、適切なファイルを読み込むように修正します。

```csharp
private ProjectData LoadProjectsFromCsv(int groupId)
{
    // Word版の場合
    string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, 
        "References", "CSV", "解答手順あり模擬試験①問題文_Word.csv");
    // ...
}

private ProjectData LoadProjectsFromJson(int groupId)
{
    // Word版の場合
    string jsonFileName = groupId switch
    {
        1 => "MOSWord演習問題文一覧.json",
        2 => "MOSWord模擬試験①問題文一覧.json",
        3 => "MOSWord模擬試験②問題文一覧.json",
        _ => "MOSWord問題文一覧.json"
    };
    // ...
}
```

---

## 🎯 Phase 4: WordCheckerクラスの基本構造作成

### プロンプト例

```
Libraries\WordGroup1\WordChecker1_1.cs を作成します。

ExcelChecker1_1.cs の構造をベースに、以下の点を変更：

1. ネームスペース: Libraries.WordGroup1
2. クラス名: WordChecker1_1
3. 参照: Microsoft.Office.Interop.Word
4. GetCurrentWordFilePath(): 開いているWord文書のパスを取得
5. CheckTask_1_1_01(): Word文書の余白をチェック
6. CheckTask_1_1_02(): Word文書の向きをチェック

以下のメソッドシグネチャで作成：
- public bool CheckTask_1_1_01()
- private bool CheckTask_1_1_01(string filePath)
- private string GetCurrentWordFilePath()

ExcelCheckerと同じエラーハンドリングとデバッグログを含めてください。
```

### WordCheckerクラスの基本構造

```csharp
using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace Libraries.WordGroup1
{
    public class WordChecker1_1
    {
        // パブリックメソッド（パラメータなし）
        public bool CheckTask_1_1_01()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_1_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        // プライベートメソッド（実際のチェックロジック）
        private bool CheckTask_1_1_01(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                // 既に開いているWordアプリケーションを取得
                try
                {
                    wordApp = (Application)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[WordChecker1_1_01] Wordアプリケーションが見つかりません");
                    return false;
                }
                
                // デバッグログ
                System.Diagnostics.Debug.WriteLine($"[WordChecker1_1_01] Looking for file: {filePath}");
                
                // 既に開いている文書を検索
                document = null;
                string fileName = Path.GetFileName(filePath);
                
                foreach (Document doc in wordApp.Documents)
                {
                    System.Diagnostics.Debug.WriteLine($"[WordChecker1_1_01] Checking document: {doc.FullName}");
                    
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        System.Diagnostics.Debug.WriteLine($"[WordChecker1_1_01] Found matching document!");
                        break;
                    }
                }
                
                if (document == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[WordChecker1_1_01] Document not found. Using ActiveDocument instead.");
                    document = wordApp.ActiveDocument;
                }
                
                if (document == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[WordChecker1_1_01] No document available");
                    return false;
                }
                
                // 余白をチェック（例：狭い = 12.7mm = 36ポイント）
                float topMargin = document.PageSetup.TopMargin;
                float bottomMargin = document.PageSetup.BottomMargin;
                float leftMargin = document.PageSetup.LeftMargin;
                float rightMargin = document.PageSetup.RightMargin;
                
                System.Diagnostics.Debug.WriteLine($"[WordChecker1_1_01] Margins: Top={topMargin}, Bottom={bottomMargin}, Left={leftMargin}, Right={rightMargin}");
                
                // 狭い余白の場合、全て約36ポイント（12.7mm）
                bool result = Math.Abs(topMargin - 36) < 1 && 
                              Math.Abs(bottomMargin - 36) < 1 && 
                              Math.Abs(leftMargin - 36) < 1 && 
                              Math.Abs(rightMargin - 36) < 1;
                
                System.Diagnostics.Debug.WriteLine($"[WordChecker1_1_01] Task result: {result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordChecker1_1_01] Exception: {ex.Message}");
                return false;
            }
        }

        // 開いているWord文書のパスを取得
        private string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try
            {
                // 実行中のWordアプリケーションを取得
                wordApp = (Application)Marshal.GetActiveObject("Word.Application");
                
                // アクティブな文書のパスを取得
                if (wordApp.ActiveDocument != null)
                {
                    return wordApp.ActiveDocument.FullName;
                }
                
                return null;
            }
            catch (COMException)
            {
                // Wordが起動していない場合
                return null;
            }
            finally
            {
                if (wordApp != null)
                    Marshal.ReleaseComObject(wordApp);
            }
        }
    }
}
```

---

## 🎯 Phase 5: MainViewModelの修正

### プロンプト例

```
Ui\ViewModels\MainViewModel.cs のExecuteScoringDirectメソッドを
Word版に対応させます。

変更点：
1. libraryName: "ExcelChecker1_1" → "WordChecker1_1"
2. namespaceName: "Libraries.Group1" → "Libraries.WordGroup1"
3. その他のロジックはExcel版と同じ

デバッグログも同じように出力してください。
```

### 変更箇所

```csharp
// Excel版
string libraryName = $"ExcelChecker{groupId}_{projectId}";
string namespaceName = $"Libraries.Group{groupId}";

// Word版
string libraryName = $"WordChecker{groupId}_{projectId}";
string namespaceName = $"Libraries.WordGroup{groupId}";

// PowerPoint版
string libraryName = $"PowerPointChecker{groupId}_{projectId}";
string namespaceName = $"Libraries.PowerPointGroup{groupId}";
```

---

## 🎯 Phase 6: config.jsonの作成

### プロンプト例

```
Assets\config.json をWord用に作成します。

Excel版の構造をベースに、以下を変更：
1. wordFile: Word文書ファイルのパス
2. taskCount: 各プロジェクトのタスク数
3. checkerDll: WordCheckerのDLLパス（使用しない場合は空でOK）

例：
{
  "tabs": {
    "1": {
      "projects": {
        "1": {
          "wordFile": "C:\\MOSTest\\Word365\\Tab1\\project1.docx",
          "taskCount": 7
        },
        "2": {
          "wordFile": "C:\\MOSTest\\Word365\\Tab1\\project2.docx",
          "taskCount": 7
        }
      }
    }
  }
}

作成してください。
```

---

## 🎯 Phase 7: プロジェクトのビルドと統合

### プロンプト例

```
全てのWordCheckerクラス（WordChecker1_1～1_10）を
MOSWordMogiApp.csprojに追加してください。

<Compile Include="Libraries\WordGroup1\WordChecker1_1.cs" />
<Compile Include="Libraries\WordGroup1\WordChecker1_2.cs" />
<Compile Include="Libraries\WordGroup1\WordChecker1_3.cs" />
<Compile Include="Libraries\WordGroup1\WordChecker1_4.cs" />
<Compile Include="Libraries\WordGroup1\WordChecker1_5.cs" />
<Compile Include="Libraries\WordGroup1\WordChecker1_6.cs" />
<Compile Include="Libraries\WordGroup1\WordChecker1_7.cs" />
<Compile Include="Libraries\WordGroup1\WordChecker1_8.cs" />
<Compile Include="Libraries\WordGroup1\WordChecker1_9.cs" />
<Compile Include="Libraries\WordGroup1\WordChecker1_10.cs" />

追加後、ビルドキャッシュをクリアしてリビルドしてください。
```

---

## 💡 効率的なプロンプトのポイント

### ✅ DO（推奨）

1. **既存のファイル名を明示する**
   ```
   「ExcelChecker1_1.cs」をベースに「WordChecker1_1.cs」を作成
   ```

2. **変更点と同じ点を明確に区別**
   ```
   変更点：問題文、Checker機能
   同じ点：UIレイアウト、ボタン機能
   ```

3. **段階的に進める**
   ```
   Phase 1: 構造確認
   Phase 2: プロジェクト作成
   Phase 3: 問題文
   Phase 4: Checker
   Phase 5: ViewModel
   Phase 6: config
   Phase 7: ビルド
   ```

4. **具体的なファイルパスとメソッド名を指定**
   ```
   「Ui\ViewModels\MainViewModel.cs」の
   「ExecuteScoringDirect」メソッドを修正
   ```

5. **デバッグログの重要性を強調**
   ```
   ExcelCheckerと同じエラーハンドリングとデバッグログを含めてください
   ```

### ❌ DON'T（非推奨）

1. **曖昧な指示**
   ```
   ❌ 「Word版を作ってください」
   ✅ 「ExcelアプリのUIとレイアウトを保持し、Checker機能のみWord対応にしてください」
   ```

2. **一度に全てを依頼**
   ```
   ❌ 「Word版アプリを完成させてください」
   ✅ 「まずWordChecker1_1.csを作成してください」
   ```

3. **変更範囲が不明確**
   ```
   ❌ 「Word用に変えてください」
   ✅ 「Microsoft.Office.Interop.ExcelをWordに変更し、GetCurrentExcelFilePathをGetCurrentWordFilePathに変更してください」
   ```

4. **ファイルパスを省略**
   ```
   ❌ 「MainViewModelを修正してください」
   ✅ 「Ui\ViewModels\MainViewModel.csのExecuteScoringDirectメソッドを修正してください」
   ```

---

## 📝 PowerPoint版の場合

同じ構造で、以下を置き換えてください：

| Excel/Word版 | PowerPoint版 |
|-------------|-------------|
| `Word` | `PowerPoint` |
| `WordChecker` | `PowerPointChecker` |
| `Microsoft.Office.Interop.Word` | `Microsoft.Office.Interop.PowerPoint` |
| `Document` | `Presentation` |
| `GetCurrentWordFilePath` | `GetCurrentPowerPointFilePath` |
| `wordFile` | `powerPointFile` |
| `.docx` | `.pptx` |

### PowerPoint用のconfig.json例

```json
{
  "tabs": {
    "1": {
      "projects": {
        "1": {
          "powerPointFile": "C:\\MOSTest\\PowerPoint365\\Tab1\\project1.pptx",
          "taskCount": 7
        }
      }
    }
  }
}
```

---

## 🎯 Phase 8: タスク状態管理（重要な注意事項）

### タスクの旗/チェックマーク表示の実装

**重要**: タスクの「後で見直す」と「解答済み」の状態管理では、TaskId（1始まり）と配列インデックス（0始まり）の変換が必要です。

### 正しい実装方法

#### 1. 配列の初期化

```csharp
// ✅ 正しい配列サイズ
// TaskIdは1始まりだが、配列は0始まりなので maxTaskCount のみ
int maxTaskCount = Math.Max(_tasks.Count, 8);
_projectTaskCompletedStates[_currentProjectId] = new bool[maxTaskCount];
_projectTaskFlaggedStates[_currentProjectId] = new bool[maxTaskCount];

// ❌ 間違い（+1は不要）
_projectTaskCompletedStates[_currentProjectId] = new bool[maxTaskCount + 1];
```

**修正箇所**:
- `LoadTasksFromCsv`メソッド（430-440行目付近）
- `LoadTasksFromJson`メソッド（485-495行目付近）

#### 2. ボタンクリック時の配列アクセス

```csharp
// ✅ 正しい（FlagButton_Clickメソッド）
private void FlagButton_Click(object sender, RoutedEventArgs e)
{
    if (!_projectTaskFlaggedStates.ContainsKey(_currentProjectId))
    {
        _projectTaskFlaggedStates[_currentProjectId] = new bool[8];
    }

    bool[] flaggedStates = _projectTaskFlaggedStates[_currentProjectId];

    // TaskIdは1始まり、配列は0始まりなので -1
    int index = _currentTaskId - 1;
    if (index >= 0 && index < flaggedStates.Length)
    {
        flaggedStates[index] = !flaggedStates[index];  // トグル
    }

    UpdateTaskDisplay();
}
```

**修正箇所**:
- `FlagButton_Click`メソッド（1148-1176行目付近）
- `CompleteButton_Click`メソッド（1178-1203行目付近）

#### 3. ボタンテキストの更新

```csharp
// ✅ 正しい（UpdateButtonTextsメソッド）
private void UpdateButtonTexts()
{
    bool[] completedStates = _projectTaskCompletedStates.ContainsKey(_currentProjectId) ?
        _projectTaskCompletedStates[_currentProjectId] : new bool[8];
    bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(_currentProjectId) ?
        _projectTaskFlaggedStates[_currentProjectId] : new bool[8];

    var completeButton = FindName("CompleteButton") as Button;
    if (completeButton != null)
    {
        // TaskIdは1始まり、配列は0始まりなので -1
        if (_currentTaskId > 0 && _currentTaskId <= completedStates.Length && 
            completedStates[_currentTaskId - 1])
        {
            completeButton.Content = "✓ 解答済み";
        }
        else
        {
            completeButton.Content = "解答済みにする";
        }
    }

    var flagButton = FindName("FlagButton") as Button;
    if (flagButton != null)
    {
        // TaskIdは1始まり、配列は0始まりなので -1
        if (_currentTaskId > 0 && _currentTaskId <= flaggedStates.Length && 
            flaggedStates[_currentTaskId - 1])
        {
            flagButton.Content = "フラグを外す";
        }
        else
        {
            flagButton.Content = "あとで見直す";
        }
    }
}
```

**修正箇所**:
- `UpdateButtonTexts`メソッド（753-790行目付近）

#### 4. タスクボタンの旗/チェックマーク表示

```csharp
// ✅ 正しい（UpdateTaskButtonsメソッド）
private void UpdateTaskButtons()
{
    bool[] completedStates = _projectTaskCompletedStates.ContainsKey(_currentProjectId) ?
        _projectTaskCompletedStates[_currentProjectId] : new bool[8];
    bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(_currentProjectId) ?
        _projectTaskFlaggedStates[_currentProjectId] : new bool[8];

    for (int i = 1; i <= 7; i++)
    {
        var button = FindName($"TaskButton{i}") as Button;
        if (button != null)
        {
            bool isVisible = i <= _tasks.Count;
            button.Visibility = isVisible ? Visibility.Visible : Visibility.Collapsed;

            if (isVisible)
            {
                // TaskIdは1始まり、配列は0始まりなので i-1 でアクセス
                int arrayIndex = i - 1;
                
                // チェックマークの表示制御
                var checkTextBlock = FindName($"Check{i}") as TextBlock;
                if (checkTextBlock != null)
                {
                    checkTextBlock.Visibility = (arrayIndex >= 0 && 
                        arrayIndex < completedStates.Length && 
                        completedStates[arrayIndex]) ?
                        Visibility.Visible : Visibility.Collapsed;
                }

                // 旗マークの表示制御
                var flagTextBlock = FindName($"Flag{i}") as TextBlock;
                if (flagTextBlock != null)
                {
                    flagTextBlock.Visibility = (arrayIndex >= 0 && 
                        arrayIndex < flaggedStates.Length && 
                        flaggedStates[arrayIndex]) ?
                        Visibility.Visible : Visibility.Collapsed;
                }
            }
        }
    }
}
```

**修正箇所**:
- `UpdateTaskButtons`メソッド（699-751行目付近）

#### 5. ReviewPageWindowでの状態読み取り

```csharp
// ✅ 正しい（ReviewPageWindow.xaml.cs）
int arrayIndex = task.TaskId - 1;
bool isCompleted = false;
bool isFlagged = false;

if (_projectTaskCompletedStates?.ContainsKey(task.ProjectId) == true)
{
    var completedArray = _projectTaskCompletedStates[task.ProjectId];
    if (arrayIndex >= 0 && arrayIndex < completedArray.Length)
    {
        isCompleted = completedArray[arrayIndex];
    }
}

if (_projectTaskFlaggedStates?.ContainsKey(task.ProjectId) == true)
{
    var flaggedArray = _projectTaskFlaggedStates[task.ProjectId];
    if (arrayIndex >= 0 && arrayIndex < flaggedArray.Length)
    {
        isFlagged = flaggedArray[arrayIndex];
    }
}
```

**修正箇所**:
- `UpdateTaskStates`メソッド（280-340行目付近）

**Word/PowerPoint版での注意点**:
- ReviewPageWindowのコンストラクタで、`_projectTaskCompletedStates`と`_projectTaskFlaggedStates`を受け取る
- これらの状態はAppBarWindowから渡されるため、Word/PowerPoint版でも同じロジックを使用
- 問題文の内容は各Officeで異なるが、状態管理のロジックは同じ

### Word/PowerPoint版での実装ポイント

Word版やPowerPoint版を作成する際は、以下のポイントを必ず守ってください：

| 項目 | 正しい実装 | 間違った実装 |
|------|-----------|-------------|
| **配列サイズ** | `new bool[maxTaskCount]` | `new bool[maxTaskCount + 1]` |
| **保存時** | `array[_currentTaskId - 1]` | `array[_currentTaskId]` |
| **読み取り時** | `array[_currentTaskId - 1]` | `array[_currentTaskId]` |
| **ループ処理** | `array[i - 1]` (iが1始まりの場合) | `array[i]` |
| **範囲チェック** | `index >= 0 && index < array.Length` | チェックなし |

### プロンプト例

```
Word版のAppBarWindow.xaml.csで、タスクの「後で見直す」と「解答済み」の状態を管理します。

重要な実装ポイント：

1. 配列初期化:
   _projectTaskCompletedStates[_currentProjectId] = new bool[maxTaskCount];
   // +1は不要。TaskIdは1始まりだが、配列は0始まりで対応

2. 状態の保存（FlagButton_Click, CompleteButton_Click）:
   int index = _currentTaskId - 1;
   if (index >= 0 && index < flaggedStates.Length)
   {
       flaggedStates[index] = true;
   }

3. ボタンテキストの更新（UpdateButtonTexts）:
   if (_currentTaskId > 0 && _currentTaskId <= completedStates.Length && 
       completedStates[_currentTaskId - 1])
   {
       completeButton.Content = "✓ 解答済み";
   }

4. タスクボタンのマーク表示（UpdateTaskButtons）:
   for (int i = 1; i <= 7; i++)
   {
       int arrayIndex = i - 1;
       checkTextBlock.Visibility = completedStates[arrayIndex] ? 
           Visibility.Visible : Visibility.Collapsed;
   }

5. ReviewPageWindowでの読み取り:
   int arrayIndex = task.TaskId - 1;
   bool isCompleted = completedStates[arrayIndex];

これらの修正をすべて含めて、Word版のAppBarWindow.xaml.csを作成してください。
```

---

## 🔧 トラブルシューティング

### よくある問題と解決方法

#### 1. Checkerクラスが見つからない

**症状:**
```
[DEBUG] Type not found: Libraries.WordGroup1.WordChecker1_1
```

**解決方法:**
```
WordChecker1_1.csがMOSWordMogiApp.csprojに追加されているか確認してください。

<Compile Include="Libraries\WordGroup1\WordChecker1_1.cs" />

追加後、ビルドキャッシュをクリアしてリビルドしてください。
```

#### 2. Word文書が見つからない

**症状:**
```
[WordChecker1_1_01] Document not found. Using ActiveDocument instead.
[WordChecker1_1_01] No document available
```

**解決方法:**
```
1. Word文書が開いているか確認
2. config.jsonのファイルパスが正しいか確認
3. GetCurrentWordFilePathメソッドにデバッグログを追加して、
   実際のファイルパスを確認
```

#### 3. 採点結果が全て×になる

**症状:**
正しく設定しても採点結果が×

**解決方法:**
```
1. デバッグログで実際の値を確認
2. 許容誤差を調整（例：Math.Abs(topMargin - 36) < 2）
3. Word/PowerPointのバージョンによる違いを確認
```

#### 4. タスクボタンの旗/チェックが前のタスクに表示される

**症状:**
```
タスク3を選択 → 「後で見直す」を押す
→ タスク2のボタンに旗マークが表示される
→ レビューページでは正しくタスク3に表示される
```

**原因:**
- `UpdateTaskButtons`メソッドで、ループ変数`i`（1始まり）を直接配列インデックスとして使用している
- `completedStates[i]`ではなく`completedStates[i - 1]`を使用する必要がある

**解決方法:**
```csharp
// UpdateTaskButtonsメソッドを修正
for (int i = 1; i <= 7; i++)
{
    // ❌ 間違い
    checkTextBlock.Visibility = completedStates[i] ? Visibility.Visible : Visibility.Collapsed;
    
    // ✅ 正しい
    int arrayIndex = i - 1;
    checkTextBlock.Visibility = (arrayIndex >= 0 && arrayIndex < completedStates.Length && 
        completedStates[arrayIndex]) ? Visibility.Visible : Visibility.Collapsed;
}
```

#### 5. ビルドエラー「ファイルがロックされています」

**症状:**
```
error MSB3027: "obj\Debug\MOSExcelMogiApp.exe" を "bin\Debug\MOSExcelMogiApp.exe" にコピーできませんでした。
別のプロセスで使用されているため、プロセスはファイル 'bin\Debug\MOSExcelMogiApp.exe' にアクセスできません。
```

**解決方法:**
```
1. アプリケーションを完全に終了する
2. Visual Studioのデバッガーを停止する
3. タスクマネージャーでプロセスが残っていないか確認
4. 再度ビルドを実行
```

---

## 📚 参考情報

### Excel版からの主な変更点まとめ

| 項目 | Excel版 | Word版 | PowerPoint版 |
|-----|---------|--------|--------------|
| 参照 | Microsoft.Office.Interop.Excel | Microsoft.Office.Interop.Word | Microsoft.Office.Interop.PowerPoint |
| アプリケーション | Application | Application | Application |
| ドキュメント | Workbook | Document | Presentation |
| シート/スライド | Worksheet | - | Slide |
| ファイル拡張子 | .xlsx | .docx | .pptx |
| GetActiveObject | "Excel.Application" | "Word.Application" | "PowerPoint.Application" |

### 便利なWord APIリファレンス

- **余白設定**: `Document.PageSetup.TopMargin`, `BottomMargin`, `LeftMargin`, `RightMargin`
- **ページ向き**: `Document.PageSetup.Orientation` (wdOrientPortrait / wdOrientLandscape)
- **行間**: `Paragraph.LineSpacing`, `LineSpacingRule`
- **スタイル**: `Paragraph.Style`
- **ページ番号**: `Document.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].PageNumbers`
- **目次**: `Document.TablesOfContents`
- **透かし**: `Document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes.AddTextEffect`

### 便利なPowerPoint APIリファレンス

- **スライドサイズ**: `Presentation.PageSetup.SlideWidth`, `SlideHeight`
- **スライド向き**: `Presentation.PageSetup.SlideOrientation`
- **テーマ**: `Presentation.TemplateName`
- **スライド数**: `Presentation.Slides.Count`
- **テキスト**: `Slide.Shapes[index].TextFrame.TextRange.Text`
- **画像**: `Slide.Shapes.AddPicture`
- **アニメーション**: `Slide.TimeLine.MainSequence`

---

## ✨ まとめ

このガイドを使用して、Excel版と同じUI・機能を持つWord版またはPowerPoint版のアプリケーションを効率的に作成できます。

**重要なポイント:**
1. 段階的に進める（Phase 1～7）
2. 既存のコードをベースにする
3. デバッグログを必ず含める
4. ファイルパスとメソッド名を明示する
5. 変更点と同じ点を明確に区別する

**アプリバーとレビューページでの問題文・問題番号の扱い:**
1. **問題番号**: プロジェクト番号とタスク番号の表示形式は各Officeで同じ（例: "プロジェクト 1"、"タスク 1"）
2. **問題文**: 各Officeアプリケーション（Excel/Word/PowerPoint）で異なる内容を表示
3. **問題文ファイル**: 各Office用の問題文ファイル（CSVまたはJSON）を用意する必要がある
   - Excel版: `References/CSV/解答手順あり模擬試験①問題文.csv` など
   - Word版: `References/CSV/解答手順あり模擬試験①問題文_Word.csv` など
   - PowerPoint版: `References/CSV/解答手順あり模擬試験①問題文_PowerPoint.csv` など
4. **問題文の読み込み**: `AppBarWindow.xaml.cs`と`ReviewPageWindow.xaml.cs`の`LoadTasks()`/`LoadProjectsFromCsv()`/`LoadProjectsFromJson()`メソッドで、適切なファイルを読み込むように修正
5. **状態管理**: 問題文の内容は異なるが、タスクの「解答済み」と「あとで見直す」の状態管理ロジックは同じ

**Word/PowerPoint版での主な変更点まとめ:**

| 項目 | Excel版 | Word版 | PowerPoint版 |
|-----|---------|--------|--------------|
| **ウィンドウ配置メソッド** | `PositionExcelWindow()` | `PositionWordWindow()` | `PositionPowerPointWindow()` |
| **ウィンドウクラス名** | `XLMAIN` | `OpusApp` | `PPTFrameClass` |
| **プロセス名** | `EXCEL` | `WINWORD` | `POWERPNT` |
| **Checkerクラス** | `ExcelChecker` | `WordChecker` | `PowerPointChecker` |
| **ネームスペース** | `Libraries.Group{id}` | `Libraries.WordGroup{id}` | `Libraries.PowerPointGroup{id}` |
| **ファイル拡張子** | `.xlsx` | `.docx` | `.pptx` |
| **config.jsonのキー** | `excelFile` | `wordFile` | `powerPointFile` |
| **問題文ファイル** | `MOS模擬アプリ問題文一覧.json` | `MOSWord問題文一覧.json` | `MOSPowerPoint問題文一覧.json` |
| **Officeアプリ終了** | `CloseExcelApplication()` | `CloseWordApplication()` | `ClosePowerPointApplication()` |
| **ファイルアクティブ化** | `ActivateExcelFile()` | `ActivateWordFile()` | `ActivatePowerPointFile()` |

**変更不要な機能:**
- UIレイアウト（AppBarWindow.xaml、ReviewPageWindow.xaml、ResultWindow.xaml）
- タイマー機能
- タスク状態管理（解答済み、あとで見直す）
- 問題文の下線表示とクリップボードコピー
- レビューページボタンと結果画面に戻るボタンの切り替え
- 全プロジェクト採点処理（Checkerクラス名とファイルパスのみ変更）
- 結果画面のフィルター機能

このプロンプト構造を使えば、Excel版と同じ品質でWord/PowerPoint版を効率的に作成できます！

