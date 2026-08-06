using System;
using System.Windows;
using System.Windows.Threading;
using Newtonsoft.Json;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using MOSExcelMogiApp;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace MOSExcelMogiApp.Views
{
    /// <summary>
    /// UiTestAppBarWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class UiTestAppBarWindow : Window
    {
        private DispatcherTimer _timer;
        private TimeSpan _remainingTime;
        private int _currentProjectId = 1;
        private bool _isPaused = false;
        private bool _timerDisabled = false; // タイマー無効化フラグ

        [DllImport("user32.dll")]
        private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        private const int SW_RESTORE = 9;
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }
        
        public UiTestAppBarWindow()
        {
            InitializeComponent();
            InitializeTimer();
            UpdateProjectTitle();
            SetWindowPosition();
        }
        
        private void SetWindowPosition()
        {
            // 画面のサイズを取得
            var screenWidth = SystemParameters.PrimaryScreenWidth;
            var screenHeight = SystemParameters.PrimaryScreenHeight;
            
            // ウィンドウのサイズを設定
            this.Width = screenWidth;
            this.Height = 120;
            
            // 位置を設定（画面の下部）
            this.Left = 0;
            this.Top = screenHeight - this.Height;
            
            // ウィンドウを最前面に表示
            this.Topmost = true;
        }

        private void AdjustScreenButton_Click(object sender, RoutedEventArgs e)
        {
            SetWindowPosition();
            PositionExcelWindow();
        }

        private void PositionExcelWindow()
        {
            try
            {
                var excelProcess = Process.GetProcessesByName("EXCEL")
                    .OrderByDescending(process =>
                    {
                        try { return process.StartTime; }
                        catch { return DateTime.MinValue; }
                    })
                    .FirstOrDefault();
                if (excelProcess == null)
                    return;

                IntPtr excelHwnd = IntPtr.Zero;
                uint processId = (uint)excelProcess.Id;
                try
                {
                    EnumWindows((windowHandle, _) =>
                    {
                        GetWindowThreadProcessId(windowHandle, out uint windowProcessId);
                        if (windowProcessId != processId)
                            return true;

                        var className = new StringBuilder(256);
                        GetClassName(windowHandle, className, className.Capacity);
                        if (!className.ToString().Contains("XLMAIN"))
                            return true;

                        excelHwnd = windowHandle;
                        return false;
                    }, IntPtr.Zero);
                }
                finally
                {
                    excelProcess.Dispose();
                }

                if (excelHwnd == IntPtr.Zero)
                    return;

                ShowWindow(excelHwnd, SW_RESTORE);
                GetWindowRect(excelHwnd, out RECT windowRect);
                GetClientRect(excelHwnd, out RECT clientRect);

                int borderWidth = (windowRect.right - windowRect.left) - clientRect.right;
                int borderHeight = (windowRect.bottom - windowRect.top) - clientRect.bottom;
                int screenWidth = (int)SystemParameters.PrimaryScreenWidth;
                int screenHeight = (int)SystemParameters.PrimaryScreenHeight;

                MoveWindow(
                    excelHwnd,
                    -borderWidth / 2,
                    -borderHeight / 2,
                    screenWidth + borderWidth,
                    screenHeight - (int)this.Height + borderHeight,
                    true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[UiTestAppBarWindow] Error positioning Excel window: {ex.Message}");
            }
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            SetWindowPosition();
        }
        
        private void InitializeTimer()
        {
            // 50分（3000秒）からカウントダウン開始
            _remainingTime = TimeSpan.FromMinutes(50);
            UpdateTimerDisplay();
            
            // タイマーを1秒間隔で更新
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(1);
            _timer.Tick += Timer_Tick;
            
            // MainWindowの「タイマーなし」チェックボックスの状態を確認
            _timerDisabled = MainWindow.IsTimerDisabled;
            
            System.Diagnostics.Debug.WriteLine($"UiTestAppBarWindow: InitializeTimer called, IsTimerDisabled = {_timerDisabled}");
            
            if (_timerDisabled)
            {
                // タイマーは開始しない
                _timer.Stop();
                System.Diagnostics.Debug.WriteLine("UiTestAppBarWindow: Timer disabled, not starting");
            }
            else
            {
                _timer.Start();
                System.Diagnostics.Debug.WriteLine("UiTestAppBarWindow: Timer enabled, starting");
            }
            
            // 「一時停止」ボタンの状態を更新
            UpdatePauseButtonState();
        }
        
        private void Timer_Tick(object sender, EventArgs e)
        {
            if (_remainingTime.TotalSeconds > 0)
            {
                _remainingTime = _remainingTime.Subtract(TimeSpan.FromSeconds(1));
                UpdateTimerDisplay();
            }
            else
            {
                _timer.Stop();
                UpdateTimerDisplay();
            }
        }
        
        private void UpdateTimerDisplay()
        {
            // タイマーの表示を更新
            var timerTextBlock = FindName("TimerTextBlock") as System.Windows.Controls.TextBlock;
            if (timerTextBlock != null)
            {
                timerTextBlock.Text = _remainingTime.ToString(@"hh\:mm\:ss");
            }
        }
        
        private void UpdateProjectTitle()
        {
            var projectInfoTextBlock = FindName("ProjectInfoTextBlock") as System.Windows.Controls.TextBlock;
            if (projectInfoTextBlock != null)
            {
                projectInfoTextBlock.Text = $"プロジェクト {_currentProjectId}/10";
            }
        }
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            _timer?.Stop();
            this.Close();
        }
        
        private void ReviewPageButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MessageBox.Show("レビューページ機能は準備中です。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"レビューページ表示エラー: {ex.Message}");
                MessageBox.Show("レビューページの表示に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void ScoreButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // config.jsonから現在のプロジェクトのタスク数を取得
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (!File.Exists(configPath))
                {
                    MessageBox.Show("config.jsonが見つかりませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                string configContent = File.ReadAllText(configPath);
                dynamic config = Newtonsoft.Json.JsonConvert.DeserializeObject(configContent);
                
                // Tab1を想定
                string tabNumber = "1";
                int taskCount = 0;
            
            try
            {
                    taskCount = config.tabs[tabNumber].projects[_currentProjectId.ToString()].taskCount;
                }
                catch
                {
                    MessageBox.Show($"プロジェクト{_currentProjectId}の設定が見つかりませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // 採点を実行
                string libraryName = $"ExcelChecker{tabNumber}_{_currentProjectId}";
                var results = ExecuteScoringDirect(libraryName, taskCount);
                
                // 採点結果ダイアログを表示（groupIdとprojectIdを渡す）
                int groupId = int.Parse(tabNumber); // tabNumberがgroupIdに対応
                MOSExcelMogiApp.Views.ScoringResultDialog.ShowResults(this, taskCount, results, groupId, _currentProjectId);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"採点中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"採点エラー: {ex.Message}\n{ex.StackTrace}");
            }
        }
        
        private List<bool> ExecuteScoringDirect(string libraryName, int taskCount)
        {
            var results = new List<bool>();
            
            try
            {
                System.Diagnostics.Debug.WriteLine($"ExecuteScoringDirect called with libraryName: {libraryName}, taskCount: {taskCount}");
                
                // Extract group and project numbers from library name
                var parts = libraryName.Replace("ExcelChecker", "").Split('_');
                if (parts.Length >= 2)
                {
                    string groupId = parts[0];
                    string projectId = parts[1];
                    
                    // Create namespace and type name
                    string namespaceName = $"Libraries.Group{groupId}";
                    string fullTypeName = $"{namespaceName}.{libraryName}";
                    
                    System.Diagnostics.Debug.WriteLine($"Looking for type: {fullTypeName}");
                    
                    // Try to get the type from loaded assemblies
                    System.Type checkerType = System.Type.GetType(fullTypeName);
                    if (checkerType == null)
                {
                        foreach (var assembly in System.AppDomain.CurrentDomain.GetAssemblies())
                        {
                            checkerType = assembly.GetType(fullTypeName);
                            if (checkerType != null) break;
                        }
                    }
                    
                    // If type not found, try to load from DLL
                    if (checkerType == null)
                {
                        System.Diagnostics.Debug.WriteLine($"Type not found in loaded assemblies, trying to load DLL");
                        
                        // Try multiple DLL paths in order of preference
                        string[] dllPaths = {
                            Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "bin", "Debug", "Libraries", $"Group{groupId}", $"{libraryName}.dll"),
                            Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Libraries", $"Group{groupId}", $"{libraryName}.dll"),
                            Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Libraries", "bin", "Debug", "net48", $"{libraryName}.dll")
                        };
                        
                        string dllPath = null;
                        foreach (string path in dllPaths)
                        {
                            System.Diagnostics.Debug.WriteLine($"Trying DLL path: {path}");
                            if (File.Exists(path))
                        {
                                dllPath = path;
                                System.Diagnostics.Debug.WriteLine($"Found DLL at: {dllPath}");
                                break;
                            }
                        }
                        
                        if (dllPath != null)
                        {
                            try
                            {
                                System.Diagnostics.Debug.WriteLine($"Loading DLL from: {dllPath}");
                                System.Reflection.Assembly assembly = System.Reflection.Assembly.LoadFrom(dllPath);
                                checkerType = assembly.GetTypes().FirstOrDefault(t => t.Name == libraryName);
                                System.Diagnostics.Debug.WriteLine($"Found type from DLL: {checkerType?.Name ?? "null"}");
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error loading DLL: {ex.Message}");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"DLL not found in any of the searched paths");
                        }
                    }
                    
                    if (checkerType != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Type found: {checkerType.Name}");
                        
                        object checkerInstance = System.Activator.CreateInstance(checkerType);
                        
                        for (int i = 1; i <= taskCount; i++)
        {
                            // 正しいメソッド名の形式を試す
                            // Group1の場合は CheckTask_1_{projectId}_{taskNumber:D2} の形式
                            string[] methodNames;
                            if (groupId == "1")
                            {
                                methodNames = new string[] {
                                    $"CheckTask_1_{projectId}_{i:D2}",
                                    $"CheckTask_1_{projectId}_0{i}",
                                    $"CheckTask_{groupId}_{projectId}_{i:D2}",
                                    $"CheckTask_{groupId}_{projectId}_0{i}"
                                };
                            }
                            else
                            {
                                methodNames = new string[] {
                                    $"CheckTask_{groupId}_{projectId}_{i:D2}",
                                    $"CheckTask_{groupId}_{projectId}_0{i}",
                                    $"CheckTask_1_{projectId}_{i:D2}",
                                    $"CheckTask_1_{projectId}_0{i}"
                                };
                            }
                            
                            bool methodFound = false;
                            foreach (string methodName in methodNames)
                            {
                                System.Diagnostics.Debug.WriteLine($"Looking for method: {methodName}");
                                System.Reflection.MethodInfo method = checkerType.GetMethod(methodName);
                                
                                if (method != null)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Method found: {methodName}");
                                    bool result = (bool)method.Invoke(checkerInstance, null);
                                    System.Diagnostics.Debug.WriteLine($"Method {methodName} result: {result}");
                                    results.Add(result);
                                    methodFound = true;
                                    break;
                }
                else
                {
                                    System.Diagnostics.Debug.WriteLine($"Method not found: {methodName}");
                }
            }
            
                            if (!methodFound)
                            {
                                System.Diagnostics.Debug.WriteLine($"No method found for task {i}, returning false");
                                results.Add(false);
                            }
                        }
                }
                else
                {
                        System.Diagnostics.Debug.WriteLine($"Type not found: {fullTypeName}");
                        // Fill with false results if type not found
                        for (int i = 0; i < taskCount; i++)
                {
                            results.Add(false);
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Invalid library name format: {libraryName}");
                    // Fill with false results if invalid format
                    for (int i = 0; i < taskCount; i++)
            {
                        results.Add(false);
            }
        }
            }
            catch (Exception ex)
        {
                System.Diagnostics.Debug.WriteLine($"Direct execution error: {ex.Message}");
                MessageBox.Show($"採点中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                // Fill with false results in case of error
                for (int i = 0; i < taskCount; i++)
                {
                    results.Add(false);
                }
            }
            
            return results;
        }
        
        private void PauseButton_Click(object sender, RoutedEventArgs e)
        {
            // タイマーが無効化されている場合は何もしない
            if (_timerDisabled)
            {
                return;
            }
            
            if (_isPaused)
            {
                // タイマーを再開
                _timer?.Start();
                _isPaused = false;
                
                // ボタンのテキストを「一時停止」に変更
                if (sender is System.Windows.Controls.Button button)
                {
                    button.Content = "一時停止";
                }
            }
            else
            {
                // タイマーを停止
                _timer?.Stop();
                _isPaused = true;
                
                // ボタンのテキストを「再開」に変更
                if (sender is System.Windows.Controls.Button button)
                {
                    button.Content = "再開";
                }
            }
        }
        
        private void UpdatePauseButtonState()
        {
            if (PauseButton != null)
            {
                if (_timerDisabled)
                {
                    // ボタンを無効化（グレーアウト）
                    PauseButton.IsEnabled = false;
                    PauseButton.Opacity = 0.5;
                    System.Diagnostics.Debug.WriteLine("UiTestAppBarWindow: PauseButton disabled (grayed out)");
                }
                else
                {
                    // ボタンを有効化
                    PauseButton.IsEnabled = true;
                    PauseButton.Opacity = 1.0;
                    System.Diagnostics.Debug.WriteLine("UiTestAppBarWindow: PauseButton enabled");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("UiTestAppBarWindow: PauseButton is null!");
            }
        }
        
        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            // 確認ダイアログを表示
            var result = MessageBox.Show(
                $"プロジェクト {_currentProjectId} をリセットしますか？",
                "リセット確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    MessageBox.Show(
                        $"プロジェクト {_currentProjectId} をリセットしました。",
                        "リセット完了",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"リセット中にエラーが発生しました：\n{ex.Message}",
                        "エラー",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }
        
        private void NextProject_Click(object sender, RoutedEventArgs e)
        {
            // 次のプロジェクトに移動
            if (_currentProjectId < 10)
            {
                _currentProjectId++;
                UpdateProjectTitle();
                MessageBox.Show($"プロジェクト {_currentProjectId} に移動しました。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("これが最後のプロジェクトです。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        
        private void PreviousTask_Click(object sender, RoutedEventArgs e)
        {
            // 前のタスクに移動する処理
            MessageBox.Show("前のタスクに移動します。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void TaskButton_Click(object sender, RoutedEventArgs e)
        {
            // タスクボタンがクリックされたときの処理
            if (sender is System.Windows.Controls.Button button && button.Tag != null)
            {
                string taskNumber = button.Tag.ToString();
                MessageBox.Show($"タスク {taskNumber} を選択しました。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void NextTask_Click(object sender, RoutedEventArgs e)
        {
            // 次のタスクに移動する処理
            MessageBox.Show("次のタスクに移動します。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void CompleteButton_Click(object sender, RoutedEventArgs e)
        {
            // 解答済みにする処理
            MessageBox.Show("タスクを解答済みにしました。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void FlagButton_Click(object sender, RoutedEventArgs e)
        {
            // あとで見直すフラグを設定する処理
            MessageBox.Show("タスクにフラグを設定しました。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        protected override void OnClosed(EventArgs e)
        {
            _timer?.Stop();
            base.OnClosed(e);
        }
    }
}
