# Word・PowerPoint アプリ UI 改善 実装手順

Excel アプリで実施した以下 3 つの改善を、Word と PowerPoint にも適用するための実装手順書です。

1. [保護ビュー（Protected View）を自動解除する](#1-保護ビューの自動解除)
2. [リセット中にメッセージを表示する](#2-リセット中メッセージの表示)
3. [レビューページからタスクを開く際にメッセージを表示する](#3-タスクを開いていますメッセージの表示)

---

## 1. 保護ビューの自動解除

### 背景

ダウンロードや別ドライブからコピーしたファイルには NTFS の `Zone.Identifier` が付き、
Office が保護ビューで開く。保護ビューでは読み取り専用になり採点が正常に動作しない。

### 対応 A：既存ファイルの Zone.Identifier を一括削除（PowerShell）

PowerShell を**管理者権限**で開き、以下を実行。パスはアプリが参照する実際のフォルダに合わせる。

```powershell
# Word
Get-ChildItem "C:\MOSTest\Word365" -Recurse -Filter "*.docx" | Unblock-File
# PowerPoint
Get-ChildItem "C:\MOSTest\PowerPoint365" -Recurse -Filter "*.pptx" | Unblock-File
```

> `Unblock-File` は Zone.Identifier ADS を削除するだけで、ファイル内容は変更しない。

---

### 対応 B：アプリ側でプログラム的に保護ビューを解除する

Word と PowerPoint の両 COM API には `ProtectedViewWindows` コレクションがある。
ファイルを開く前後でこのコレクションを確認し、保護ビューで開かれていれば `Edit()` を呼んで通常モードに切り替える。

#### Word (`MOS Word app/Views/AppBarWindow.xaml.cs`)

**① using エイリアスを追加（ファイル先頭）**

```csharp
using WordProtectedViewWindow = Microsoft.Office.Interop.Word.ProtectedViewWindow;
```

**② ファイルを開く処理（`OpenProjectDocument` または `NavigateToTask` 内）を修正**

現在 `wordApp.Documents.Open(filePath)` を呼んでいる箇所の直前に以下のブロックを挿入する。

```csharp
// 保護ビューで既に開かれている場合は編集モードに切り替える
bool openedFromProtectedView = false;
try
{
    string targetFileName = Path.GetFileName(filePath);
    foreach (WordProtectedViewWindow pvw in wordApp.ProtectedViewWindows)
    {
        try
        {
            if (string.Equals(pvw.Caption, targetFileName, StringComparison.OrdinalIgnoreCase))
            {
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Word protected view detected, switching to edit: {filePath}");
                targetDocument = pvw.Edit(); // 保護ビューを解除して編集可能な Document を返す
                openedFromProtectedView = true;
                break;
            }
        }
        catch { }
    }
}
catch { }

if (!openedFromProtectedView)
{
    targetDocument = wordApp.Documents.Open(filePath, ReadOnly: false);
}
```

> `pvw.FullName` は Word の `ProtectedViewWindow` API に存在しないため、必ず `pvw.Caption`（ファイル名部分）で比較する。

---

#### PowerPoint (`Mos PowerPoint Mogi App/Views/UiTestAppBarWindow.xaml.cs`)

**① using エイリアスを追加（ファイル先頭）**

```csharp
using PptProtectedViewWindow = Microsoft.Office.Interop.PowerPoint.ProtectedViewWindow;
```

**② `OpenProjectDocument` 内でファイルを開く処理を修正**

```csharp
// 保護ビューで既に開かれている場合は編集モードに切り替える
bool openedFromProtectedView = false;
try
{
    string targetFileName = Path.GetFileName(filePath);
    foreach (PptProtectedViewWindow pvw in pptApp.ProtectedViewWindows)
    {
        try
        {
            if (string.Equals(pvw.Caption, targetFileName, StringComparison.OrdinalIgnoreCase))
            {
                System.Diagnostics.Debug.WriteLine($"[UiTestAppBarWindow] PPT protected view detected, switching to edit: {filePath}");
                targetPresentation = pvw.Edit(); // 保護ビューを解除して Presentation を返す
                openedFromProtectedView = true;
                break;
            }
        }
        catch { }
    }
}
catch { }

if (!openedFromProtectedView)
{
    targetPresentation = pptApp.Presentations.Open(filePath, ReadOnly: Microsoft.Office.Core.MsoTriState.msoFalse);
}
```

---

## 2. リセット中メッセージの表示

### 背景

リセット処理（ファイルコピー・Word/PPT 起動）は重い。ボタンクリック直後に UI スレッドをブロックすると
「リセット中です」ウィンドウが描画されない。また、WPF オブジェクトにバックグラウンドスレッドからアクセスすると
クロススレッド例外が発生する。

### 解決策（Excel と同じパターン）

- ボタンハンドラを `async void` にする
- オーバーレイウィンドウを `Show()` した後に `Dispatcher.InvokeAsync(..., DispatcherPriority.Background)` でリセット処理を実行
- `DispatcherPriority.Background` により、先にオーバーレイの描画が完了してからリセット処理が走る

---

#### Word (`MOS Word app/Views/AppBarWindow.xaml.cs`)

`ResetButton_Click` を以下のように書き換える。

```csharp
private async void ResetButton_Click(object sender, RoutedEventArgs e)
{
    try
    {
        var result = MessageBox.Show(
            $"プロジェクト{_currentProjectId}をリセットしますか？\n現在の変更内容は失われます。",
            "リセット確認",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes) return;

        // 「リセット中です」オーバーレイを作成して即時表示
        var waitWindow = new Window
        {
            Title = "リセット中",
            Width = 300,
            Height = 120,
            WindowStyle = WindowStyle.None,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ShowInTaskbar = false,
            ResizeMode = ResizeMode.NoResize,
            Topmost = true,
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = System.Windows.Media.Brushes.SteelBlue,
            BorderThickness = new Thickness(2)
        };
        var stack = new System.Windows.Controls.StackPanel
        {
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            Margin = new Thickness(16)
        };
        stack.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = $"リセット中です...\nプロジェクト {_currentProjectId}",
            FontSize = 14,
            TextAlignment = System.Windows.TextAlignment.Center,
            Foreground = System.Windows.Media.Brushes.SteelBlue
        });
        waitWindow.Content = stack;
        waitWindow.Show();

        Exception resetError = null;
        // Word ドキュメントを閉じてリセットを UI スレッド上で実行
        // DispatcherPriority.Background でオーバーレイの描画を先に確定させる
        await Application.Current.Dispatcher.InvokeAsync(() =>
        {
            try
            {
                if (!CloseAllWordDocuments())
                    TryQuitWord();
                System.Threading.Thread.Sleep(500);
                ResetProject(_groupId, _currentProjectId);
            }
            catch (Exception ex) { resetError = ex; }
        }, System.Windows.Threading.DispatcherPriority.Background);

        waitWindow.Close();

        if (resetError != null)
        {
            MessageBox.Show($"リセット中にエラーが発生しました: {resetError.Message}",
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        else
        {
            // リセット後、Wordドキュメントを再読み込み
            OpenProjectDocument(_currentProjectId, _groupId);

            bool originalTopmost = this.Topmost;
            try
            {
                this.Topmost = true;
                this.Activate();
                MessageBox.Show(this, "プロジェクトをリセットしました。", "リセット完了",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally { this.Topmost = originalTopmost; }
        }
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"[ResetButton_Click] Error: {ex.Message}");
        MessageBox.Show($"リセット中にエラーが発生しました: {ex.Message}",
            "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}
```

> `Thread.Sleep(500)` は `Dispatcher.InvokeAsync` の中で実行されるため UI はブロックされない。

---

#### PowerPoint (`Mos PowerPoint Mogi App/Views/UiTestAppBarWindow.xaml.cs`)

`ResetButton_Click` を以下のように書き換える。

```csharp
private async void ResetButton_Click(object sender, RoutedEventArgs e)
{
    try
    {
        var result = MessageBox.Show(
            $"プロジェクト{_currentProjectId}をリセットしますか？\n現在の変更内容は失われます。",
            "リセット確認",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes) return;

        // 「リセット中です」オーバーレイを作成して即時表示
        var waitWindow = new Window
        {
            Title = "リセット中",
            Width = 300,
            Height = 120,
            WindowStyle = WindowStyle.None,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ShowInTaskbar = false,
            ResizeMode = ResizeMode.NoResize,
            Topmost = true,
            Background = System.Windows.Media.Brushes.White,
            BorderBrush = System.Windows.Media.Brushes.SteelBlue,
            BorderThickness = new Thickness(2)
        };
        var stack = new System.Windows.Controls.StackPanel
        {
            VerticalAlignment = System.Windows.VerticalAlignment.Center,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            Margin = new Thickness(16)
        };
        stack.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = $"リセット中です...\nプロジェクト {_currentProjectId}",
            FontSize = 14,
            TextAlignment = System.Windows.TextAlignment.Center,
            Foreground = System.Windows.Media.Brushes.SteelBlue
        });
        waitWindow.Content = stack;
        waitWindow.Show();

        Exception resetError = null;
        // リセット処理を UI スレッド上で実行（WPF オブジェクトへのアクセスを保護）
        await Application.Current.Dispatcher.InvokeAsync(() =>
        {
            try
            {
                Libraries.PPLogReader.ClearSnapshot();
                ResetProject(_groupId, _currentProjectId);
                PowerPointChecker1_1.ResetTask4SlideDeletionState();
            }
            catch (Exception ex) { resetError = ex; }
        }, System.Windows.Threading.DispatcherPriority.Background);

        waitWindow.Close();

        if (resetError != null)
        {
            MessageBox.Show($"リセット中にエラーが発生しました: {resetError.Message}",
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        else
        {
            OpenProjectDocument(_currentProjectId, _groupId);
            WriteCurrentTaskFile();
            MessageBox.Show("プロジェクトをリセットしました。", "リセット完了",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"[ResetButton_Click] Error: {ex.Message}");
        MessageBox.Show($"リセット中にエラーが発生しました: {ex.Message}",
            "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}
```

---

## 3. タスクを開いていますメッセージの表示

### 背景

レビューページからタスクをクリックすると `OnNavigateToTask` が呼ばれ、
アプリバー側が Word/PowerPoint を起動する。この処理には数秒かかることがあるが、
現状はクリックから何も表示されないため操作が効いているか分からない。

### 解決策（Excel と同じパターン）

`ReviewPageWindow.NavigateToTask` の**先頭**でオーバーレイを `Show()` し、
`await Dispatcher.InvokeAsync(() => {}, DispatcherPriority.Render)` で描画を確定させてから
重い処理（`OnNavigateToTask` 呼び出し）を行う。

---

#### Word (`MOS Word app/Views/ReviewPageWindow.xaml.cs`)

```csharp
private async void NavigateToTask(ReviewTaskInfo taskInfo)
{
    System.Diagnostics.Debug.WriteLine($"NavigateToTask called: ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");

    if (OnNavigateToTask == null || taskInfo.ProjectId <= 0 || taskInfo.TaskId <= 0)
    {
        System.Diagnostics.Debug.WriteLine($"ナビゲーション条件不一致");
        return;
    }

    // クリック直後にオーバーレイを表示（重い処理より前）
    var openingOverlay = new Window
    {
        Title = "タスクを開いています",
        Width = 320,
        Height = 120,
        WindowStyle = WindowStyle.None,
        WindowStartupLocation = WindowStartupLocation.CenterScreen,
        ShowInTaskbar = false,
        ResizeMode = ResizeMode.NoResize,
        Topmost = true,
        Background = System.Windows.Media.Brushes.White,
        BorderBrush = System.Windows.Media.Brushes.SteelBlue,
        BorderThickness = new Thickness(2)
    };
    var stack = new System.Windows.Controls.StackPanel
    {
        VerticalAlignment = System.Windows.VerticalAlignment.Center,
        HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
        Margin = new Thickness(16)
    };
    stack.Children.Add(new System.Windows.Controls.TextBlock
    {
        Text = "タスクを開いています...",
        FontSize = 14,
        TextAlignment = System.Windows.TextAlignment.Center,
        Foreground = System.Windows.Media.Brushes.SteelBlue
    });
    openingOverlay.Content = stack;
    openingOverlay.Show();

    // 描画を確定させてからナビゲーション処理へ
    await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);

    try
    {
        _timer?.Stop();
        OnNavigateToTask(taskInfo.ProjectId, taskInfo.TaskId);
        openingOverlay.Close();
        this.Close();
    }
    catch (Exception ex)
    {
        openingOverlay.Close();
        System.Diagnostics.Debug.WriteLine($"ナビゲーション実行エラー: {ex.Message}");
        MessageBox.Show($"タスクの移動中にエラーが発生しました: {ex.Message}",
            "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}
```

また、`TaskButton_Click` のシグネチャを非同期に対応させる。

```csharp
private async void TaskButton_Click(object sender, RoutedEventArgs e)
{
    var button = sender as Button;
    if (button != null && button.DataContext is ReviewTaskInfo taskInfo)
    {
        await System.Threading.Tasks.Task.Run(() => { }); // コンパイラに async を認識させる
        NavigateToTask(taskInfo); // await 不要（NavigateToTask 自体が async void）
    }
}
```

> `NavigateToTask` が `async void` のため `await` 不要。`TaskButton_Click` も `async void` に変更することで、
> 将来的に await が必要になったときに対応しやすくなる。

---

#### PowerPoint (`Mos PowerPoint Mogi App/Views/ReviewPageWindow.xaml.cs`)

PowerPoint の `ReviewPageWindow` は `TaskButton_Click` 内で直接 `OnNavigateToTask` を呼んでいる。
`NavigateToTask` ラッパーメソッドを追加してオーバーレイを表示する。

```csharp
private void TaskButton_Click(object sender, RoutedEventArgs e)
{
    if (sender is Button btn && btn.DataContext is ReviewTaskInfo taskInfo)
    {
        if (OnNavigateToTask != null && taskInfo.ProjectId > 0 && taskInfo.TaskId > 0)
        {
            NavigateToTaskWithOverlay(taskInfo);
        }
    }
}

private async void NavigateToTaskWithOverlay(ReviewTaskInfo taskInfo)
{
    // クリック直後にオーバーレイを表示
    var openingOverlay = new Window
    {
        Title = "タスクを開いています",
        Width = 320,
        Height = 120,
        WindowStyle = WindowStyle.None,
        WindowStartupLocation = WindowStartupLocation.CenterScreen,
        ShowInTaskbar = false,
        ResizeMode = ResizeMode.NoResize,
        Topmost = true,
        Background = System.Windows.Media.Brushes.White,
        BorderBrush = System.Windows.Media.Brushes.SteelBlue,
        BorderThickness = new Thickness(2)
    };
    var stack = new System.Windows.Controls.StackPanel
    {
        VerticalAlignment = System.Windows.VerticalAlignment.Center,
        HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
        Margin = new Thickness(16)
    };
    stack.Children.Add(new System.Windows.Controls.TextBlock
    {
        Text = "タスクを開いています...",
        FontSize = 14,
        TextAlignment = System.Windows.TextAlignment.Center,
        Foreground = System.Windows.Media.Brushes.SteelBlue
    });
    openingOverlay.Content = stack;
    openingOverlay.Show();

    // 描画を確定させてからナビゲーション処理へ
    await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);

    try
    {
        OnNavigateToTask(taskInfo.ProjectId, taskInfo.TaskId);
        openingOverlay.Close();
        this.Close();
    }
    catch (Exception ex)
    {
        openingOverlay.Close();
        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Navigation error: {ex.Message}");
        MessageBox.Show($"タスクの移動中にエラーが発生しました: {ex.Message}",
            "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
    }
}
```

> PowerPoint の `UiTestAppBarWindow.NavigateToTask` は既に `RunWithDelayedPrepareOverlayAsync` で
> 「準備中」オーバーレイを表示している。`ReviewPageWindow` 側のオーバーレイは **クリックから
> `OnNavigateToTask` 呼び出しまでの間**だけ表示される短い待機表示として機能する。

---

## 変更ファイル一覧

| ファイル | 変更内容 |
|---|---|
| `MOS Word app/Views/AppBarWindow.xaml.cs` | 保護ビュー解除コード追加、`ResetButton_Click` を async 化してオーバーレイ表示 |
| `MOS Word app/Views/ReviewPageWindow.xaml.cs` | `NavigateToTask` を async void 化してオーバーレイを先頭で表示 |
| `Mos PowerPoint Mogi App/Views/UiTestAppBarWindow.xaml.cs` | 保護ビュー解除コード追加、`ResetButton_Click` を async 化してオーバーレイ表示 |
| `Mos PowerPoint Mogi App/Views/ReviewPageWindow.xaml.cs` | `NavigateToTaskWithOverlay` メソッド追加、`TaskButton_Click` から呼び出し変更 |

---

## 注意事項

- **`pvw.Caption` で比較する**：Word・PowerPoint の `ProtectedViewWindow` オブジェクトには `FullName` プロパティがない。`Caption` にはファイル名（拡張子付き）が格納されている。
- **`Dispatcher.InvokeAsync(..., DispatcherPriority.Background)` の優先度**：`Background` を指定することで、UI レンダリング（オーバーレイ描画）が完了してからリセット処理が始まる。`Normal` では描画前に処理が始まる可能性がある。
- **ビルドと動作確認**：変更後は必ずアプリを閉じてからビルドし、`bin\Debug` へのコピーが完了していることを確認してから起動すること（ファイルロックによるビルドエラーを防ぐため）。
