using System;
using System.Windows;
using System.Windows.Threading;
using Newtonsoft.Json;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows.Media;
using System.Text.RegularExpressions;
using System.Windows.Documents;
using System.Windows.Input;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Reflection;
using System.Windows.Interop;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using PowerPointPresentation = Microsoft.Office.Interop.PowerPoint.Presentation;
using Microsoft.Office.Interop.PowerPoint;
using System.Configuration;

namespace MOS_PowerPoint_app.Views
{
    /// <summary>
    /// UiTestAppBarWindow.xaml の相互作用ロジック（PowerPointアプリ用）
    /// </summary>
    public partial class UiTestAppBarWindow : System.Windows.Window
    {
        // Windows API用の定義
        [DllImport("user32.dll")]
        static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
        
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        
        [DllImport("user32.dll")]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        
        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        
        [DllImport("user32.dll")]
        static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);
        
        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        
        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        
        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }
        private DispatcherTimer _timer;
        private TimeSpan _remainingTime;
        private int _currentProjectId = 1;
        private int _currentTaskId = 1;
        private int _groupId = 1; // グループIDを保存
        private List<TaskInfo> _tasks;
        private ProjectData _projectData;
        private Dictionary<int, bool[]> _projectTaskCompletedStates = new Dictionary<int, bool[]>(); // プロジェクトごとの解答済み状態
        private Dictionary<int, bool[]> _projectTaskFlaggedStates = new Dictionary<int, bool[]>(); // プロジェクトごとのフラグ状態
        private Dictionary<int, bool[]> _projectTaskViewedStates = new Dictionary<int, bool[]>(); // プロジェクトごとの閲覧状態（未読問題の追跡用）
        private DateTime _projectStartTime; // プロジェクト開始時刻
        private DispatcherTimer _projectTimer; // プロジェクト用タイマー（5分制限）
        private Dictionary<int, Dictionary<int, string>> _clipboardTargets = new Dictionary<int, Dictionary<int, string>>(); // クリップボード対象（プロジェクトID → タスクID → 問題文）
        private bool _isPaused = false; // 一時停止状態
        private DateTime _pauseStartTime; // 一時停止開始時刻（プロジェクトタイマー用）
        private List<System.Windows.Controls.Button> _dynamicTaskButtons = new List<System.Windows.Controls.Button>(); // 動的に生成されたタスクボタン（8番目以降）
        
        public UiTestAppBarWindow(int projectId = 1, int groupId = 1)
        {
            InitializeComponent();
            _currentProjectId = projectId;
            _groupId = groupId; // グループIDを保存
            InitializeTimer();
            InitializeProjectTimer();
            LoadClipboardTargets(); // クリップボード対象を先に読み込む
            LoadTasks();
            UpdateTaskDisplay();
            SetWindowPosition();
            // 注意: PowerPointプレゼンテーションはMainViewModelのExecuteOpenProjectで既に開かれている
            // ここでは開かない（PositionPowerPointWindowはSetWindowPositionで呼ばれる）
        }
        
        private void SetWindowPosition()
        {
            // PowerPointウィンドウを配置
            PositionPowerPointWindow();
            
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

            // アプリバーのウィンドウを1920x258サイズで、PowerPointの下に配置
            // 高さ: 258 (1032 / 4)
            // 位置: Y=774 (PowerPointの下)
            // 境界線を考慮して位置を調整
            int x = -borderWidth / 2; // 左側の境界線を考慮
            int y = 774 - borderHeight / 2; // 上側の境界線を考慮（PowerPointの下）
            int width = 1920 + borderWidth; // 境界線を含めた幅
            int height = 258 + borderHeight; // 境界線を含めた高さ

            MoveWindow(hWnd, x, y, width, height, true);
            
            // ウィンドウを最前面に表示
            this.Topmost = true;
        }
        
        private void PositionPowerPointWindow()
        {
            try
            {
                // 実行中のPowerPointプロセスを取得
                var pptProcesses = Process.GetProcessesByName("POWERPNT");
                if (pptProcesses.Length == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[AppBarWindow] PowerPoint process not found");
                    return;
                }

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
                            if (className.ToString().Contains("PPTFrameClass"))
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
                    // 高さ: 258 * 3 = 774 (1032 / 4 * 3)
                    // 境界線を考慮して位置を調整（マージンをゼロにする）
                    int pptX = -pptBorderWidth / 2;
                    int pptY = -pptBorderHeight / 2;
                    int pptWidth = 1920 + pptBorderWidth;
                    int pptHeight = 774 + pptBorderHeight;
                    
                    MoveWindow(pptHwnd, pptX, pptY, pptWidth, pptHeight, true);
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] PowerPoint window positioned: {pptWidth}x{pptHeight} at ({pptX}, {pptY})");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[AppBarWindow] PowerPoint window handle not found");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Error positioning PowerPoint window: {ex.Message}");
            }
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            // ウィンドウハンドルが利用可能になるまで少し待機してから配置
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SetWindowPosition();
            }), DispatcherPriority.Loaded);
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
            
            // MainWindowのタイマー設定を確認
            bool timerDisabled = MainWindow.IsTimerDisabled;
            
            System.Diagnostics.Debug.WriteLine($"UiTestAppBarWindow: InitializeTimer called, IsTimerDisabled = {timerDisabled}");
            
            if (timerDisabled)
            {
                // タイマーは開始しない
                _timer.Stop();
                System.Diagnostics.Debug.WriteLine("UiTestAppBarWindow: Timer disabled, not starting");
                
                // 一時停止ボタンを無効化
                UpdatePauseButtonState(true);
            }
            else
            {
                _timer.Start();
                System.Diagnostics.Debug.WriteLine("UiTestAppBarWindow: Timer enabled, starting");
                
                // 一時停止ボタンを有効化
                UpdatePauseButtonState(false);
            }
        }
        
        private void UpdatePauseButtonState(bool isDisabled)
        {
            var pauseButton = FindName("PauseButton") as System.Windows.Controls.Button;
            if (pauseButton != null)
            {
                pauseButton.IsEnabled = !isDisabled;
            }
        }
        
        private void InitializeProjectTimer()
        {
            // プロジェクト用タイマーを初期化（5分制限）
            _projectTimer = new DispatcherTimer();
            _projectTimer.Interval = TimeSpan.FromSeconds(1);
            _projectTimer.Tick += ProjectTimer_Tick;
            
            // 最初のプロジェクト開始時刻を設定
            _projectStartTime = DateTime.Now;
            
            // MainWindowのタイマー設定を確認
            bool timerDisabled = MainWindow.IsTimerDisabled;
            
            if (!timerDisabled)
            {
                _projectTimer.Start();
            }
        }
        
        private void Timer_Tick(object sender, EventArgs e)
        {
            // タイマーが無効化されている場合は何もしない
            if (MainWindow.IsTimerDisabled)
            {
                return;
            }
            
            if (_remainingTime.TotalSeconds > 0)
            {
                _remainingTime = _remainingTime.Subtract(TimeSpan.FromSeconds(1));
                UpdateTimerDisplay();
            }
            else
            {
                _timer.Stop();
                UpdateTimerDisplay();
                // 試験終了処理：結果画面を表示
                ShowResultWindow();
            }
        }
        
        /// <summary>
        /// レビューページ（問題一覧）を表示する
        /// </summary>
        private void ShowReviewPageWindow()
        {
            try
            {
                var reviewWindow = new ReviewPageWindow(
                    _remainingTime,
                    _projectTaskCompletedStates,
                    _projectTaskFlaggedStates,
                    _projectTaskViewedStates,
                    _groupId);

                reviewWindow.OnNavigateToTask = (projectId, taskId) =>
                {
                    var appBar = System.Windows.Application.Current.Windows.OfType<UiTestAppBarWindow>().FirstOrDefault();
                    if (appBar != null)
                    {
                        appBar.Show();
                        appBar.Activate();
                        appBar.NavigateToTask(projectId, taskId);
                    }
                };
                reviewWindow.OnShowResultRequested = () =>
                {
                    ShowResultWindow();
                };
                reviewWindow.OnBackRequested = () =>
                {
                    this.Show();
                    this.Activate();
                };

                reviewWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                reviewWindow.Show();
                this.Hide();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UiTestAppBarWindow] ShowReviewPageWindow error: {ex.Message}");
                MessageBox.Show($"レビューページの表示中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 結果画面を表示する
        /// </summary>
        private void ShowResultWindow()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] Showing result window");
                
                // 結果画面を作成
                var resultWindow = new ResultWindow(_projectTaskFlaggedStates, _projectTaskViewedStates, _groupId);
                resultWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                resultWindow.Topmost = true;
                
                // OnEndRequestedを設定（終了時: PowerPoint終了・メイン画面に戻る）
                resultWindow.OnEndRequested = () =>
                {
                    CloseAllPowerPointPresentations();
                    try
                    {
                        var pptProcesses = Process.GetProcessesByName("POWERPNT");
                        foreach (var proc in pptProcesses)
                        {
                            proc.Kill();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OnEndRequested] PowerPoint プロセス終了エラー: {ex.Message}");
                    }
                    this.Close();
                    var main = System.Windows.Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                    if (main != null)
                    {
                        main.Show();
                        main.Activate();
                    }
                };
                
                // OnNavigateToTaskを設定
                resultWindow.OnNavigateToTask = (projectId, taskId) =>
                {
                    // AppBarWindowを確実に表示
                    var appBarWindow = System.Windows.Application.Current.Windows.OfType<UiTestAppBarWindow>().FirstOrDefault();
                    if (appBarWindow != null)
                    {
                        appBarWindow.Show();
                        appBarWindow.Activate();
                        appBarWindow.NavigateToTask(projectId, taskId);
                    }
                };
                
                resultWindow.Show();
                resultWindow.Activate();
                
                // AppBarWindowを非表示にする
                this.Hide();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UiTestAppBarWindow] Error showing result window: {ex.Message}");
                MessageBox.Show($"結果画面の表示中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// タスクに移動する（結果画面から呼び出される）
        /// </summary>
        public void NavigateToTask(int projectId, int taskId)
        {
            System.Diagnostics.Debug.WriteLine($"NavigateToTask called: ProjectId={projectId}, TaskId={taskId}");
            
            try
            {
                // プロジェクトを変更
                if (projectId != _currentProjectId)
                {
                    System.Diagnostics.Debug.WriteLine($"プロジェクト変更: {_currentProjectId} -> {projectId}");
                    _currentProjectId = projectId;
                    LoadCurrentProjectTasks();
                }
                
                // タスクを変更
                if (taskId != _currentTaskId && taskId >= 1 && taskId <= _tasks.Count)
                {
                    System.Diagnostics.Debug.WriteLine($"タスク変更: {_currentTaskId} -> {taskId}");
                    _currentTaskId = taskId;
                }
                
                // UIを更新
                UpdateTaskDisplay();
                
                // メインウィンドウを表示
                this.Show();
                this.WindowState = WindowState.Normal;
                this.Activate();
                this.Focus();
                
                System.Diagnostics.Debug.WriteLine($"プロジェクト{projectId}のタスク{taskId}に移動しました");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ナビゲーションエラー: {ex.Message}");
            }
        }
        
        private void ProjectTimer_Tick(object sender, EventArgs e)
        {
            // タイマーが無効化されている場合は何もしない
            if (MainWindow.IsTimerDisabled)
            {
                return;
            }
            
            // プロジェクト開始から5分経過したかチェック
            var elapsed = DateTime.Now - _projectStartTime;
            if (elapsed.TotalMinutes >= 5.0)
            {
                _projectTimer.Stop();
                MoveToNextProjectWithMessage();
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
        
        private void PauseButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // タイマーが無効化されている場合は何もしない
                if (MainWindow.IsTimerDisabled)
                {
                    MessageBox.Show("タイマーは無効化されています。", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                if (!_isPaused)
                {
                    // 一時停止
                    _timer?.Stop();
                    _projectTimer?.Stop();
                    _pauseStartTime = DateTime.Now;
                    _isPaused = true;
                    
                    // ボタンテキストを「再開」に変更
                    var pauseButton = sender as System.Windows.Controls.Button;
                    if (pauseButton != null)
                    {
                        pauseButton.Content = "再開";
                    }
                    
                    System.Diagnostics.Debug.WriteLine("[PauseButton] 一時停止しました");
                }
                else
                {
                    // 再開
                    // 一時停止時間を計算
                    var pauseDuration = DateTime.Now - _pauseStartTime;
                    
                    // プロジェクト開始時刻を調整（一時停止時間分を加算）
                    _projectStartTime = _projectStartTime.Add(pauseDuration);
                    
                    // タイマーを再開
                    _timer?.Start();
                    _projectTimer?.Start();
                    _isPaused = false;
                    
                    // ボタンテキストを「一時停止」に変更
                    var pauseButton = sender as System.Windows.Controls.Button;
                    if (pauseButton != null)
                    {
                        pauseButton.Content = "一時停止";
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[PauseButton] 再開しました（一時停止時間: {pauseDuration.TotalSeconds}秒）");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PauseButton] Error: {ex.Message}");
                MessageBox.Show($"一時停止/再開処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void ScoreButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[ScoreButton] 採点を開始: プロジェクト{_currentProjectId}, グループ{_groupId} (PowerPoint)");
                
                // 現在のプロジェクトのタスク数を取得
                int taskCount = _tasks != null ? _tasks.Count : 0;
                if (taskCount == 0)
                {
                    MessageBox.Show("タスクが見つかりません。", "採点エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // PowerPointChecker DLLのパスを構築（後で実装）
                // 現在は採点機能は準備中
                int passedCount = 0;
                int totalTasks = taskCount;
                var taskResults = new List<(int taskId, bool result, string error)>();
                
                // 採点機能は後で実装
                for (int taskNum = 1; taskNum <= taskCount; taskNum++)
                {
                    taskResults.Add((taskNum, false, "採点機能は準備中です"));
                }
                
                // 結果を表示
                StringBuilder resultMessage = new StringBuilder();
                resultMessage.AppendLine($"採点完了: {passedCount}/{totalTasks} タスク合格");
                resultMessage.AppendLine($"得点: {passedCount}点 / {totalTasks}点");
                resultMessage.AppendLine();
                resultMessage.AppendLine("詳細:");
                
                foreach (var (taskId, result, error) in taskResults)
                {
                    if (!string.IsNullOrEmpty(error))
                    {
                        resultMessage.AppendLine($"  タスク{taskId}: × (エラー: {error})");
                    }
                    else
                    {
                        resultMessage.AppendLine($"  タスク{taskId}: {(result ? "✓ 合格" : "× 不合格")}");
                    }
                }
                
                MessageBox.Show(resultMessage.ToString(), "採点結果", MessageBoxButton.OK, MessageBoxImage.Information);
                
                System.Diagnostics.Debug.WriteLine($"[ScoreButton] 採点完了: {passedCount}/{totalTasks}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ScoreButton] Error: {ex.Message}");
                MessageBox.Show($"採点中にエラーが発生しました:\n{ex.Message}", "採点エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            _timer?.Stop();
            _projectTimer?.Stop();
            SaveAllPowerPointPresentations();
            // 結果画面を表示
            ShowResultWindow();
        }
        
        private void ReviewPageButton_Click(object sender, RoutedEventArgs e)
        {
            ShowReviewPageWindow();
        }
        
        
        private void LoadTasks()
        {
            try
            {
                // JSONファイルから問題文を読み込む（プロジェクトルートのファイルを使用）
                string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOS模擬アプリ問題文一覧_PowerPoint.json");
                
                // ファイルが存在しない場合は、References/JSONフォルダも試す
                if (!File.Exists(jsonPath))
                {
                    jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", "MOS模擬アプリ問題文一覧_PowerPoint.json");
                }
                
                // ファイルが存在しない場合はエラー
                if (!File.Exists(jsonPath))
                {
                    System.Diagnostics.Debug.WriteLine($"JSONファイルが見つかりません: {jsonPath}");
                    _tasks = new List<TaskInfo>();
                    return;
                }
                
                // JSONファイルを読み込む
                LoadTasksFromJson(jsonPath);
                
                // 現在のプロジェクトのタスクを取得
                LoadCurrentProjectTasks();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"タスク読み込みエラー: {ex.Message}");
                _tasks = new List<TaskInfo>();
            }
        }
        
        private void LoadTasksFromJson(string jsonPath)
        {
            try
            {
                string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
                _projectData = JsonConvert.DeserializeObject<ProjectData>(jsonContent);
                
                System.Diagnostics.Debug.WriteLine($"JSONから{_projectData?.Projects?.Count ?? 0}個のプロジェクト、合計{_projectData?.Projects?.Sum(p => p.Tasks?.Count ?? 0) ?? 0}個のタスクを読み込みました");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"JSON読み込みエラー: {ex.Message}");
                _projectData = new ProjectData { Projects = new List<ProjectInfo>() };
            }
        }
        
        private void LoadClipboardTargets()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[LoadClipboardTargets] 開始");
                System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] AppDomain.CurrentDomain.BaseDirectory: {AppDomain.CurrentDomain.BaseDirectory}");
                System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] Assembly.Location: {System.Reflection.Assembly.GetExecutingAssembly().Location}");
                
                // CSVファイル名
                string csvFileName = "MOS模擬アプリ正誤判定表251120_挿入入力のみ.csv";
                
                // 複数のパス候補を試す
                List<string> pathCandidates = new List<string>();
                
                // 1. 実行ディレクトリからの相対パス
                pathCandidates.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reference", "CSV", csvFileName));
                
                // 2. 実行ディレクトリからの相対パス（References）
                pathCandidates.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "CSV", csvFileName));
                
                // 3. アセンブリの場所からの相対パス
                string assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string assemblyDir = Path.GetDirectoryName(assemblyLocation);
                pathCandidates.Add(Path.Combine(assemblyDir, "Reference", "CSV", csvFileName));
                pathCandidates.Add(Path.Combine(assemblyDir, "References", "CSV", csvFileName));
                
                // 4. プロジェクトルートからの相対パス
                string projectRoot = Path.GetDirectoryName(assemblyDir);
                if (projectRoot != null)
                {
                    pathCandidates.Add(Path.Combine(projectRoot, "Reference", "CSV", csvFileName));
                    pathCandidates.Add(Path.Combine(projectRoot, "References", "CSV", csvFileName));
                }
                
                // 5. さらに上の階層も試す
                if (projectRoot != null)
                {
                    string projectRootParent = Path.GetDirectoryName(projectRoot);
                    if (projectRootParent != null)
                    {
                        pathCandidates.Add(Path.Combine(projectRootParent, "Reference", "CSV", csvFileName));
                        pathCandidates.Add(Path.Combine(projectRootParent, "References", "CSV", csvFileName));
                    }
                }
                
                string csvPath = null;
                foreach (string candidatePath in pathCandidates)
                {
                    System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] パス候補を確認: {candidatePath}");
                    if (File.Exists(candidatePath))
                    {
                        csvPath = candidatePath;
                        System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] CSVファイルが見つかりました: {csvPath}");
                        break;
                    }
                }
                
                if (csvPath == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] クリップボード対象CSVファイルが見つかりません。試したパス:");
                    foreach (string candidatePath in pathCandidates)
                    {
                        System.Diagnostics.Debug.WriteLine($"  - {candidatePath}");
                    }
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] CSVファイルを読み込みます: {csvPath}");
                string[] lines = File.ReadAllLines(csvPath, Encoding.UTF8);
                System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] CSVファイルの行数: {lines.Length}");
                
                int loadedCount = 0;
                // ヘッダー行をスキップ（1行目）
                for (int i = 1; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (string.IsNullOrEmpty(line))
                        continue;
                    
                    // デバッグ: 元の行の内容を確認
                    System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] 行 {i + 1} の元の内容: {line}");
                    
                    // CSVのパース（カンマ区切り、引用符内のカンマを考慮）
                    string[] fields = ParseCsvLine(line);
                    System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] 行 {i + 1} のパース結果: {fields.Length}個のフィールド");
                    for (int j = 0; j < fields.Length; j++)
                    {
                        System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets]   フィールド{j}: {fields[j].Substring(0, Math.Min(100, fields[j].Length))}...");
                    }
                    
                    if (fields.Length < 3)
                    {
                        System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] フィールド数が不足: {fields.Length} (行 {i + 1})");
                        continue;
                    }
                    
                    // プロジェクト,タスク,問題文,解答操作
                    if (int.TryParse(fields[0].Trim(), out int projectId) &&
                        int.TryParse(fields[1].Trim(), out int taskId))
                    {
                        string description = fields.Length > 2 ? fields[2].Trim() : "";
                        if (string.IsNullOrEmpty(description))
                            continue;
                        
                        // デバッグ: 問題文に""が含まれているか確認
                        bool containsQuotes = description.Contains('"');
                        System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] 問題文に\"が含まれているか: {containsQuotes}");
                        if (containsQuotes)
                        {
                            int quoteCount = description.Count(c => c == '"');
                            System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] 問題文内の\"の数: {quoteCount}");
                        }
                        
                        // プロジェクトが存在しない場合は作成
                        if (!_clipboardTargets.ContainsKey(projectId))
                        {
                            _clipboardTargets[projectId] = new Dictionary<int, string>();
                        }
                        
                        // 問題文を保存（「"」で囲まれた部分を含む）
                        _clipboardTargets[projectId][taskId] = description;
                        loadedCount++;
                        System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] 読み込み: プロジェクト{projectId}, タスク{taskId}, 問題文: {description.Substring(0, Math.Min(50, description.Length))}...");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] プロジェクトIDまたはタスクIDのパースに失敗 (行 {i + 1}): {fields[0]}, {fields[1]}");
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] クリップボード対象を{_clipboardTargets.Sum(p => p.Value.Count)}件読み込みました (実際の読み込み数: {loadedCount})");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] クリップボード対象読み込みエラー: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] スタックトレース: {ex.StackTrace}");
            }
        }
        
        private void LoadTasksFromCsv(string csvPath)
        {
            try
            {
                // MOSボタンアプリ正誤判定表.csvは「タスク,問題文,解答操作」形式
                // プロジェクト情報がないため、全タスクをプロジェクト1に割り当て
                var tasks = new List<TaskInfo>();
                
                string[] lines = File.ReadAllLines(csvPath, Encoding.UTF8);
                
                // ヘッダー行をスキップ（1行目）
                for (int i = 1; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (string.IsNullOrEmpty(line))
                        continue;
                    
                    // CSVのパース（カンマ区切り、引用符内のカンマを考慮）
                    string[] fields = ParseCsvLine(line);
                    if (fields.Length < 2)
                        continue;
                    
                    // タスク,問題文,解答操作
                    if (int.TryParse(fields[0].Trim(), out int taskId))
                    {
                        string description = fields.Length > 1 ? fields[1].Trim() : "";
                        if (string.IsNullOrEmpty(description))
                            continue;
                        
                        tasks.Add(new TaskInfo
                        {
                            TaskId = taskId,
                            Description = description
                        });
                    }
                }
                
                // タスクをタスクID順にソート
                tasks.Sort((a, b) => a.TaskId.CompareTo(b.TaskId));
                
                // プロジェクト1として設定
                _projectData = new ProjectData
                {
                    Projects = new List<ProjectInfo>
                    {
                        new ProjectInfo
                        {
                            ProjectId = 1,
                            Tasks = tasks
                        }
                    }
                };
                
                System.Diagnostics.Debug.WriteLine($"CSVからプロジェクト1に{tasks.Count}個のタスクを読み込みました");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CSV読み込みエラー: {ex.Message}");
                _projectData = new ProjectData { Projects = new List<ProjectInfo>() };
            }
        }
        
        private string[] ParseCsvLine(string line)
        {
            var fields = new List<string>();
            bool inQuotes = false;
            int fieldStartIndex = 0; // 現在のフィールドの開始位置
            StringBuilder currentField = new StringBuilder();
            
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                
                if (c == '"')
                {
                    if (!inQuotes)
                    {
                        // フィールドの開始引用符の可能性
                        // 次の文字を確認して、フィールド全体が引用符で囲まれているか判断
                        bool isFieldStart = (i == 0 || line[i - 1] == ',');
                        if (isFieldStart)
                        {
                            // フィールドの開始引用符
                            inQuotes = true;
                            fieldStartIndex = i;
                            // 開始引用符は追加しない（CSV標準）
                        }
                        else
                        {
                            // フィールド内の引用符（フィールド全体が引用符で囲まれていない場合）
                            // この場合は引用符を保持する
                            currentField.Append('"');
                        }
                    }
                    else if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // エスケープされた引用符（""）
                        currentField.Append('"');
                        i++; // 次の文字をスキップ
                    }
                    else if (i + 1 < line.Length && (line[i + 1] == ',' || i + 1 == line.Length))
                    {
                        // フィールドの終了引用符（次の文字がカンマまたは行末）
                        inQuotes = false;
                        // 終了引用符は追加しない（CSV標準）
                    }
                    else
                    {
                        // フィールド内の引用符（フィールド全体が引用符で囲まれていない場合）
                        // この場合は引用符を保持する
                        currentField.Append('"');
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    // フィールドの区切り
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                    fieldStartIndex = i + 1;
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
        
        private void LoadCurrentProjectTasks()
        {
            if (_projectData != null)
            {
                var currentProject = _projectData.Projects.Find(p => p.ProjectId == _currentProjectId);
                if (currentProject != null)
                {
                    _tasks = currentProject.Tasks;
                    _currentTaskId = 1; // 新しいプロジェクトの最初のタスクにリセット
                    
                    // 新しいプロジェクトの状態を初期化（タスク数に合わせて動的にサイズを設定）
                    int taskCount = _tasks != null ? _tasks.Count : 0;
                    int arraySize = Math.Max(taskCount, 1); // タスク数、最小1（配列は0始まりなので+1は不要）
                    
                    if (!_projectTaskCompletedStates.ContainsKey(_currentProjectId))
                    {
                        _projectTaskCompletedStates[_currentProjectId] = new bool[arraySize];
                    }
                    else
                    {
                        // 既存の配列のサイズが不足している場合は拡張
                        if (_projectTaskCompletedStates[_currentProjectId].Length < arraySize)
                        {
                            var oldArray = _projectTaskCompletedStates[_currentProjectId];
                            var newArray = new bool[arraySize];
                            Array.Copy(oldArray, newArray, oldArray.Length);
                            _projectTaskCompletedStates[_currentProjectId] = newArray;
                        }
                    }
                    
                    if (!_projectTaskFlaggedStates.ContainsKey(_currentProjectId))
                    {
                        _projectTaskFlaggedStates[_currentProjectId] = new bool[arraySize];
                    }
                    else
                    {
                        // 既存の配列のサイズが不足している場合は拡張
                        if (_projectTaskFlaggedStates[_currentProjectId].Length < arraySize)
                        {
                            var oldArray = _projectTaskFlaggedStates[_currentProjectId];
                            var newArray = new bool[arraySize];
                            Array.Copy(oldArray, newArray, oldArray.Length);
                            _projectTaskFlaggedStates[_currentProjectId] = newArray;
                        }
                    }
                    
                    // 閲覧状態も初期化
                    if (!_projectTaskViewedStates.ContainsKey(_currentProjectId))
                    {
                        _projectTaskViewedStates[_currentProjectId] = new bool[arraySize];
                    }
                    else
                    {
                        // 既存の配列のサイズが不足している場合は拡張
                        if (_projectTaskViewedStates[_currentProjectId].Length < arraySize)
                        {
                            var oldArray = _projectTaskViewedStates[_currentProjectId];
                            var newArray = new bool[arraySize];
                            Array.Copy(oldArray, newArray, oldArray.Length);
                            _projectTaskViewedStates[_currentProjectId] = newArray;
                        }
                    }
                }
                else
                {
                    _tasks = new List<TaskInfo>();
                }
            }
            
            // プロジェクトタイトルを更新
            UpdateProjectTitle();
        }
        
        private void UpdateProjectTitle()
        {
            if (_projectData?.Projects != null)
            {
                int totalProjects = _projectData.Projects.Count;
                var projectInfoTextBlock = FindName("ProjectInfoTextBlock") as System.Windows.Controls.TextBlock;
                if (projectInfoTextBlock != null)
                {
                    projectInfoTextBlock.Text = $"プロジェクト {_currentProjectId}/{totalProjects}";
                }
            }
        }
        
        private void UpdateTaskDisplay()
        {
            System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] 開始: プロジェクト{_currentProjectId}, タスク{_currentTaskId}");
            
            // 閲覧状態を記録（未読問題の追跡用）
            MarkTaskAsViewed(_currentProjectId, _currentTaskId);
            
            // タスク説明の表示を更新
            var taskDescriptionTextBlock = FindName("TaskDescriptionTextBlock") as System.Windows.Controls.TextBlock;
            if (taskDescriptionTextBlock == null)
            {
                System.Diagnostics.Debug.WriteLine("[UpdateTaskDisplay] TaskDescriptionTextBlockが見つかりません");
                return;
            }
            
            if (_tasks == null)
            {
                System.Diagnostics.Debug.WriteLine("[UpdateTaskDisplay] _tasksがnullです");
                return;
            }
            
            if (_currentTaskId > _tasks.Count)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] _currentTaskId({_currentTaskId})が_tasks.Count({_tasks.Count})を超えています");
                return;
            }
            
            var currentTask = _tasks.Find(t => t.TaskId == _currentTaskId);
            if (currentTask == null)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] タスク{_currentTaskId}が見つかりません");
                return;
            }
            
            System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] _clipboardTargetsにプロジェクト{_currentProjectId}が含まれているか: {_clipboardTargets.ContainsKey(_currentProjectId)}");
            if (_clipboardTargets.ContainsKey(_currentProjectId))
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] プロジェクト{_currentProjectId}のタスク数: {_clipboardTargets[_currentProjectId].Count}");
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] プロジェクト{_currentProjectId}にタスク{_currentTaskId}が含まれているか: {_clipboardTargets[_currentProjectId].ContainsKey(_currentTaskId)}");
            }
            
            // 現在のプロジェクト・タスクがクリップボード対象に含まれているかチェック
            string clipboardTargetDescription = null;
            if (_clipboardTargets.ContainsKey(_currentProjectId) &&
                _clipboardTargets[_currentProjectId].ContainsKey(_currentTaskId))
            {
                clipboardTargetDescription = _clipboardTargets[_currentProjectId][_currentTaskId];
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] クリップボード対象の問題文を取得: {clipboardTargetDescription.Substring(0, Math.Min(100, clipboardTargetDescription.Length))}...");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] クリップボード対象が見つかりません。通常の問題文を使用します");
            }
            
            // クリップボード対象がある場合は下線付き表示、ない場合は通常表示
            if (clipboardTargetDescription != null)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] SetTextWithUnderlineを呼び出します（クリップボード対象）");
                SetTextWithUnderline(taskDescriptionTextBlock, clipboardTargetDescription);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] SetTextWithUnderlineを呼び出します（通常の問題文）");
                SetTextWithUnderline(taskDescriptionTextBlock, currentTask.Description);
            }
            
            // タスクボタンの状態を更新（チェックマークと旗マークも含む）
            UpdateTaskButtons();
            
            // ボタンのテキストを更新
            UpdateButtonTexts();
        }
        
        /// <summary>
        /// タスクを閲覧済みとして記録（未読問題の追跡用）
        /// </summary>
        private void MarkTaskAsViewed(int projectId, int taskId)
        {
            // 現在のプロジェクトの状態を取得または初期化
            if (!_projectTaskViewedStates.ContainsKey(projectId))
            {
                int taskCount = _tasks != null ? _tasks.Count : 0;
                int arraySize = Math.Max(taskCount, 1);
                _projectTaskViewedStates[projectId] = new bool[arraySize];
            }
            
            bool[] viewedStates = _projectTaskViewedStates[projectId];
            // タスクIDは1始まり、配列は0始まりなので -1
            int arrayIndex = taskId - 1;
            
            if (arrayIndex >= 0 && arrayIndex < viewedStates.Length)
            {
                viewedStates[arrayIndex] = true;
                System.Diagnostics.Debug.WriteLine($"[MarkTaskAsViewed] プロジェクト{projectId}のタスク{taskId}を閲覧済みとして記録しました");
            }
        }
        
        /// <summary>
        /// テキスト内の"で囲まれた部分に下線を付けてTextBlockに設定します
        /// "自体は表示せず、その中のテキストだけに下線を付けます
        /// </summary>
        private void SetTextWithUnderline(TextBlock textBlock, string text)
        {
            System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] 開始: テキスト長={text?.Length ?? 0}");
            
            if (string.IsNullOrEmpty(text))
            {
                System.Diagnostics.Debug.WriteLine("[SetTextWithUnderline] テキストが空です");
                textBlock.Text = string.Empty;
                return;
            }

            System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] テキスト内容: {text.Substring(0, Math.Min(100, text.Length))}...");
            
            textBlock.Inlines.Clear();
            textBlock.Text = string.Empty; // TextプロパティをクリアしてInlinesを使用
            
            // "で囲まれた部分を検索して下線を付ける
            int startIndex = 0;
            bool foundQuotes = false;
            int quoteCount = 0;
            
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
                quoteCount++;
                System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] \"を検出 (位置 {quoteStart}, {quoteCount}個目)");
                
                // "の前のテキストを追加
                if (quoteStart > startIndex)
                {
                    string beforeText = text.Substring(startIndex, quoteStart - startIndex);
                    textBlock.Inlines.Add(new Run(beforeText));
                    System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] \"の前のテキストを追加: {beforeText.Substring(0, Math.Min(50, beforeText.Length))}...");
                }
                
                // "の終了位置を検索
                int quoteEnd = text.IndexOf('"', quoteStart + 1);
                if (quoteEnd == -1)
                {
                    // "が見つからない場合は残りをそのまま追加
                    System.Diagnostics.Debug.WriteLine("[SetTextWithUnderline] 終了の\"が見つかりません");
                    textBlock.Inlines.Add(new Run(text.Substring(quoteStart)));
                    break;
                }
                
                // "で囲まれた部分のテキスト（"を除く）に下線を付けて追加
                string quotedText = text.Substring(quoteStart + 1, quoteEnd - quoteStart - 1);
                System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] \"で囲まれたテキストを検出: {quotedText}");
                
                var run = new Run(quotedText);
                run.TextDecorations = TextDecorations.Underline;
                run.Cursor = Cursors.Hand; // マウスカーソルをポインターに変更
                run.Foreground = new SolidColorBrush(Colors.Blue);
                run.MouseDown += (sender, e) => OnUnderlinedTextClick(quotedText, e);
                textBlock.Inlines.Add(run);
                System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] 下線付きテキストを追加: {quotedText}");
                
                startIndex = quoteEnd + 1;
            }
            
            System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] 検出された\"の数: {quoteCount}, foundQuotes: {foundQuotes}");
            
            // "が見つからなかった場合は通常のテキストとして設定
            if (!foundQuotes)
            {
                System.Diagnostics.Debug.WriteLine("[SetTextWithUnderline] \"が見つからなかったため、通常のテキストとして設定");
                textBlock.Inlines.Clear();
                textBlock.Text = text;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] 完了: {textBlock.Inlines.Count}個のInline要素を追加");
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
        
        private void UpdateTaskButtons()
        {
            if (_tasks == null) return;
            
            // 現在のプロジェクトの状態を取得
            int maxTaskCount = _tasks != null ? _tasks.Count : 0;
            bool[] completedStates = _projectTaskCompletedStates.ContainsKey(_currentProjectId) ? 
                _projectTaskCompletedStates[_currentProjectId] : new bool[Math.Max(maxTaskCount, 1)];
            bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(_currentProjectId) ? 
                _projectTaskFlaggedStates[_currentProjectId] : new bool[Math.Max(maxTaskCount, 1)];

            // XAMLで定義されているタスクボタンの数（7つ）まで処理
            int maxStaticButtonCount = 7; // XAMLで定義されているボタンの数
            for (int i = 1; i <= maxStaticButtonCount; i++)
            {
                var button = FindName($"TaskButton{i}") as System.Windows.Controls.Button;
                if (button != null)
                {
                    // 現在のプロジェクトのタスク数に応じて表示/非表示を制御
                    bool isVisible = i <= _tasks.Count;
                    button.Visibility = isVisible ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                    
                    if (isVisible)
                    {
                        UpdateTaskButtonState(button, i, completedStates, flaggedStates);
                    }
                }
            }
            
            // 8つ目以降のタスクボタンを動的に生成（タスク数が7より多い場合）
            var container = FindName("TaskButtonsContainer") as System.Windows.Controls.StackPanel;
            if (container != null && maxTaskCount > maxStaticButtonCount)
            {
                // 既存の動的ボタンを削除
                foreach (var btn in _dynamicTaskButtons)
                {
                    container.Children.Remove(btn);
                }
                _dynamicTaskButtons.Clear();
                
                // TaskButton7のインデックスを取得
                var taskButton7 = FindName("TaskButton7") as System.Windows.Controls.Button;
                int insertIndex = taskButton7 != null ? container.Children.IndexOf(taskButton7) + 1 : container.Children.Count - 1;
                
                // 8つ目以降のボタンを生成
                for (int i = maxStaticButtonCount + 1; i <= maxTaskCount; i++)
                {
                    var button = CreateTaskButton(i);
                    _dynamicTaskButtons.Add(button);
                    
                    // TaskButton7の後に挿入
                    container.Children.Insert(insertIndex, button);
                    insertIndex++; // 次の挿入位置を更新
                }
            }
            else if (container != null && maxTaskCount <= maxStaticButtonCount)
            {
                // タスク数が7以下になった場合は動的ボタンを削除
                foreach (var btn in _dynamicTaskButtons)
                {
                    container.Children.Remove(btn);
                }
                _dynamicTaskButtons.Clear();
            }
            
            // 動的ボタンの状態も更新
            foreach (var button in _dynamicTaskButtons)
            {
                int taskId = (int)button.Tag;
                if (taskId <= maxTaskCount)
                {
                    UpdateTaskButtonState(button, taskId, completedStates, flaggedStates);
                }
            }
        }
        
        private void UpdateTaskButtonState(System.Windows.Controls.Button button, int taskId, bool[] completedStates, bool[] flaggedStates)
        {
            // 数字のTextBlockを更新
            var grid = button.Content as System.Windows.Controls.Grid;
            if (grid != null)
            {
                // 静的ボタン（XAMLで定義）の場合はNameで特定、動的ボタンの場合はTextで判定
                var taskTextBlock = grid.Children.OfType<System.Windows.Controls.TextBlock>()
                    .FirstOrDefault(tb => tb.Name == $"TaskText{taskId}" || (tb.Name == null && tb.Text != "✓" && tb.Text != "🚩"));
                if (taskTextBlock != null)
                {
                    taskTextBlock.Text = taskId.ToString();
                }
                
                // チェックマークの表示制御（タスクIDは1始まり、配列は0始まりなので -1）
                var checkTextBlock = grid.Children.OfType<System.Windows.Controls.TextBlock>()
                    .FirstOrDefault(tb => tb.Name == $"Check{taskId}" || (tb.Name == null && tb.Text == "✓"));
                if (checkTextBlock != null)
                {
                    int arrayIndex = taskId - 1;
                    checkTextBlock.Visibility = (arrayIndex >= 0 && arrayIndex < completedStates.Length && completedStates[arrayIndex]) ? 
                        System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                }
                
                // 旗マークの表示制御（タスクIDは1始まり、配列は0始まりなので -1）
                var flagTextBlock = grid.Children.OfType<System.Windows.Controls.TextBlock>()
                    .FirstOrDefault(tb => tb.Name == $"Flag{taskId}" || (tb.Name == null && tb.Text == "🚩"));
                if (flagTextBlock != null)
                {
                    int arrayIndex = taskId - 1;
                    flagTextBlock.Visibility = (arrayIndex >= 0 && arrayIndex < flaggedStates.Length && flaggedStates[arrayIndex]) ? 
                        System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;
                }
            }
            
            // ボタンの選択状態を更新
            if (taskId == _currentTaskId)
            {
                // 現在のタスクボタンは選択状態
                button.Background = System.Windows.Media.Brushes.White;
                button.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Gray);
                button.BorderThickness = new System.Windows.Thickness(2);
                button.Foreground = System.Windows.Media.Brushes.Black;
            }
            else
            {
                // 他のタスクボタンは非選択状態
                button.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightGray);
                button.BorderThickness = new System.Windows.Thickness(0);
                button.Foreground = System.Windows.Media.Brushes.Black;
            }
        }
        
        private System.Windows.Controls.Button CreateTaskButton(int taskId)
        {
            var button = new System.Windows.Controls.Button
            {
                Width = 120,
                Height = 32,
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightGray),
                BorderThickness = new System.Windows.Thickness(0),
                Margin = new System.Windows.Thickness(0, 0, 5, 0),
                Tag = taskId
            };
            
            button.Click += TaskButton_Click;
            
            var grid = new System.Windows.Controls.Grid
            {
                Width = 120,
                Height = 32
            };
            
            // タスク番号のTextBlock
            var taskTextBlock = new System.Windows.Controls.TextBlock
            {
                Text = taskId.ToString(),
                FontSize = 16,
                Foreground = System.Windows.Media.Brushes.Black,
                FontWeight = System.Windows.FontWeights.Bold,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center
            };
            grid.Children.Add(taskTextBlock);
            
            // チェックマークのTextBlock
            var checkTextBlock = new System.Windows.Controls.TextBlock
            {
                Text = "✓",
                FontSize = 16,
                Foreground = System.Windows.Media.Brushes.Green,
                FontWeight = System.Windows.FontWeights.Bold,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new System.Windows.Thickness(0, 0, 15, 0),
                Visibility = System.Windows.Visibility.Collapsed
            };
            grid.Children.Add(checkTextBlock);
            
            // 旗マークのTextBlock
            var flagTextBlock = new System.Windows.Controls.TextBlock
            {
                Text = "🚩",
                FontSize = 16,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new System.Windows.Thickness(15, 0, 0, 0),
                Visibility = System.Windows.Visibility.Collapsed
            };
            grid.Children.Add(flagTextBlock);
            
            button.Content = grid;
            
            return button;
        }
        
        
        private void UpdateButtonTexts()
        {
            // 現在のプロジェクトの状態を取得
            int taskCount = _tasks != null ? _tasks.Count : 0;
            bool[] completedStates = _projectTaskCompletedStates.ContainsKey(_currentProjectId) ? 
                _projectTaskCompletedStates[_currentProjectId] : new bool[Math.Max(taskCount, 1)];
            bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(_currentProjectId) ? 
                _projectTaskFlaggedStates[_currentProjectId] : new bool[Math.Max(taskCount, 1)];
            
            // 解答済みボタンのテキストを更新
            var completeButton = FindName("CompleteButton") as System.Windows.Controls.Button;
            var completeButtonFooter = FindName("CompleteButtonFooter") as System.Windows.Controls.Button;
            
            if (completeButton != null)
            {
                // タスクIDは1始まり、配列は0始まりなので -1
                int arrayIndex = _currentTaskId - 1;
                if (arrayIndex >= 0 && arrayIndex < completedStates.Length && completedStates[arrayIndex])
                {
                    completeButton.Content = "✓ 解答済み";
                }
                else
                {
                    completeButton.Content = "解答済みにする";
                }
            }
            
            if (completeButtonFooter != null)
            {
                // タスクIDは1始まり、配列は0始まりなので -1
                int arrayIndex = _currentTaskId - 1;
                if (arrayIndex >= 0 && arrayIndex < completedStates.Length && completedStates[arrayIndex])
                {
                    completeButtonFooter.Content = "✓ 解答済み";
                }
                else
                {
                    completeButtonFooter.Content = "解答済みにする";
                }
            }
            
            // フラグボタンのテキストを更新
            var flagButton = FindName("FlagButton") as System.Windows.Controls.Button;
            var flagButtonFooter = FindName("FlagButtonFooter") as System.Windows.Controls.Button;
            
            if (flagButton != null)
            {
                // タスクIDは1始まり、配列は0始まりなので -1
                int arrayIndex = _currentTaskId - 1;
                if (arrayIndex >= 0 && arrayIndex < flaggedStates.Length && flaggedStates[arrayIndex])
                {
                    flagButton.Content = "フラグを外す";
                }
                else
                {
                    flagButton.Content = "あとで見直す";
                }
            }
            
            if (flagButtonFooter != null)
            {
                // タスクIDは1始まり、配列は0始まりなので -1
                int arrayIndex = _currentTaskId - 1;
                if (arrayIndex >= 0 && arrayIndex < flaggedStates.Length && flaggedStates[arrayIndex])
                {
                    flagButtonFooter.Content = "フラグを外す";
                }
                else
                {
                    flagButtonFooter.Content = "あとで見直す";
                }
            }
        }
        
        private void PreviousTask_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTaskId > 1)
            {
                _currentTaskId--;
                UpdateTaskDisplay();
            }
        }
        
        private void NextTask_Click(object sender, RoutedEventArgs e)
        {
            if (_tasks != null && _currentTaskId < _tasks.Count)
            {
                _currentTaskId++;
                UpdateTaskDisplay();
            }
        }
        
        private void TaskButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Tag != null)
            {
                int taskId = int.Parse(button.Tag.ToString());
                if (taskId >= 1 && taskId <= _tasks.Count)
                {
                    _currentTaskId = taskId;
                    UpdateTaskDisplay();
                }
            }
        }
        
        private void NextProject_Click(object sender, RoutedEventArgs e)
        {
            MoveToNextProject();
        }
        
        private void MoveToNextProject()
        {
            // プロジェクトの最大数をチェック（JSONファイルの最大プロジェクトID）
            int maxProjectId = _projectData?.Projects?.Max(p => p.ProjectId) ?? 1;
            
            // 次のプロジェクトに移動
            _currentProjectId++;
            
            if (_currentProjectId > maxProjectId)
            {
                // 最後のプロジェクトを超えた場合はメッセージを表示
                System.Diagnostics.Debug.WriteLine($"プロジェクト{maxProjectId}を超えました");
                MessageBox.Show("すべてのプロジェクトが完了しました。", "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            // 新しいプロジェクトのPowerPointプレゼンテーションを開く
            OpenProjectDocument(_currentProjectId, _groupId);
            
            // プロジェクト変更時は状態をリセットしない（Dictionaryで管理）
            
            // 新しいプロジェクトのタスクを読み込み
            LoadCurrentProjectTasks();
            UpdateTaskDisplay();
            
            // プロジェクトタイマーをリセット
            ResetProjectTimer();
            
            System.Diagnostics.Debug.WriteLine($"プロジェクト{_currentProjectId}に移動しました");
        }
        
        private void OpenProjectDocument(int projectId, int groupId)
        {
            try
            {
                // MainViewModelと同じロジックでファイルパスを構築
                // App.configからパスを読み込む（存在しない場合はデフォルト値を使用）
                string basePath = ConfigurationManager.AppSettings["PowerPointDataPath"] 
                    ?? @"C:\MOSTest\PowerPoint365";
                string tabFolder = Path.Combine(basePath, $"Tab{groupId}");
                
                // Project1.pptxからProject10.pptxを検索
                string[] possibleNames = { $"Project{projectId}.pptx", $"Project{projectId}.ppt" };
                string filePath = null;
                
                foreach (var fileName in possibleNames)
                {
                    string fullPath = Path.Combine(tabFolder, fileName);
                    if (File.Exists(fullPath))
                    {
                        filePath = fullPath;
                        break;
                    }
                }
                
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    System.Diagnostics.Debug.WriteLine($"プロジェクト{projectId}のファイルが見つかりません: {tabFolder}");
                    return;
                }
                
                // PowerPointアプリケーションを取得または作成
                PowerPointApp pptApp = null;
                try
                {
                    pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                }
                catch
                {
                    pptApp = new PowerPointApp();
                    pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                }
                
                // 現在開いているプレゼンテーションを閉じてから新しいプレゼンテーションを開く
                PowerPointPresentation presentation = null;
                try
                {
                    // 既に開いているプレゼンテーションがある場合は閉じる
                    while (pptApp.Presentations.Count > 0)
                    {
                        PowerPointPresentation openPres = pptApp.Presentations[1]; // 1-based index
                        try
                        {
                            openPres.Close();
                        }
                        catch (Exception closeEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"プレゼンテーションを閉じる際のエラー: {closeEx.Message}");
                            // エラーが発生しても次に進む
                        }
                        finally
                        {
                            // COMオブジェクトの参照を解放
                            try
                            {
                                if (openPres != null)
                                {
                                    Marshal.ReleaseComObject(openPres);
                                }
                            }
                            catch { }
                        }
                    }
                    
                    // 新しいプレゼンテーションを開く
                    presentation = pptApp.Presentations.Open(filePath, WithWindow: Microsoft.Office.Core.MsoTriState.msoTrue);
                    System.Diagnostics.Debug.WriteLine($"プレゼンテーションを開きました: {filePath}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"プレゼンテーションを開く際のエラー: {ex.Message}");
                    return;
                }
                finally
                {
                    // COMオブジェクトの参照を解放
                    if (presentation != null)
                    {
                        try
                        {
                            Marshal.ReleaseComObject(presentation);
                        }
                        catch { }
                    }
                }
                
                // PowerPointウィンドウを画面の上部2/3に配置
                PositionPowerPointWindow();
                
                System.Diagnostics.Debug.WriteLine($"プロジェクト{projectId}のプレゼンテーションを開きました: {filePath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"プロジェクトプレゼンテーションを開く際のエラー: {ex.Message}");
            }
        }
        
        private void MoveToNextProjectWithMessage()
        {
            // メッセージを表示
            MessageBox.Show("5分経ったので次のプロジェクトに移動します", "時間切れ", 
                          MessageBoxButton.OK, MessageBoxImage.Information);
            
            // 次のプロジェクトに移動
            MoveToNextProject();
        }
        
        private void ResetProjectTimer()
        {
            // プロジェクトタイマーをリセット
            _projectTimer?.Stop();
            _projectStartTime = DateTime.Now;
            _projectTimer?.Start();
        }
        
        private void FlagButton_Click(object sender, RoutedEventArgs e)
        {
            // 現在のプロジェクトの状態を取得または初期化
            if (!_projectTaskFlaggedStates.ContainsKey(_currentProjectId))
            {
                int taskCount = _tasks != null ? _tasks.Count : 0;
                int arraySize = Math.Max(taskCount, 1); // 配列は0始まりなので+1は不要
                _projectTaskFlaggedStates[_currentProjectId] = new bool[arraySize];
            }
            
            bool[] flaggedStates = _projectTaskFlaggedStates[_currentProjectId];
            // タスクIDは1始まり、配列は0始まりなので -1
            int arrayIndex = _currentTaskId - 1;
            
            if (arrayIndex >= 0 && arrayIndex < flaggedStates.Length)
            {
                System.Diagnostics.Debug.WriteLine($"FlagButton_Click: flaggedStates[{arrayIndex}]={flaggedStates[arrayIndex]}, _currentTaskId={_currentTaskId}");
                
                if (!flaggedStates[arrayIndex])
                {
                    // フラグを設定
                    flaggedStates[arrayIndex] = true;
                    System.Diagnostics.Debug.WriteLine($"タスク{_currentTaskId}のフラグを設定しました");
                }
                else
                {
                    // フラグを解除
                    flaggedStates[arrayIndex] = false;
                    System.Diagnostics.Debug.WriteLine($"タスク{_currentTaskId}のフラグを解除しました");
                }
            }
            
            // UIを更新
            UpdateTaskDisplay();
        }
        
        
        private void CompleteButton_Click(object sender, RoutedEventArgs e)
        {
            // 現在のプロジェクトの状態を取得または初期化
            if (!_projectTaskCompletedStates.ContainsKey(_currentProjectId))
            {
                int taskCount = _tasks != null ? _tasks.Count : 0;
                int arraySize = Math.Max(taskCount, 1); // 配列は0始まりなので+1は不要
                _projectTaskCompletedStates[_currentProjectId] = new bool[arraySize];
            }
            
            bool[] completedStates = _projectTaskCompletedStates[_currentProjectId];
            // タスクIDは1始まり、配列は0始まりなので -1
            int arrayIndex = _currentTaskId - 1;
            
            if (arrayIndex >= 0 && arrayIndex < completedStates.Length)
            {
                System.Diagnostics.Debug.WriteLine($"CompleteButton_Click: completedStates[{arrayIndex}]={completedStates[arrayIndex]}, _currentTaskId={_currentTaskId}");
                
                if (!completedStates[arrayIndex])
                {
                    // チェックマークを設定
                    completedStates[arrayIndex] = true;
                    System.Diagnostics.Debug.WriteLine($"タスク{_currentTaskId}のチェックマークを設定しました");
                }
                else
                {
                    // チェックマークを解除
                    completedStates[arrayIndex] = false;
                    System.Diagnostics.Debug.WriteLine($"タスク{_currentTaskId}のチェックマークを解除しました");
                }
            }
            
            // UIを更新
            UpdateTaskDisplay();
        }
        
        
        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var result = MessageBox.Show(
                    $"プロジェクト{_currentProjectId}をリセットしますか？\n現在の変更内容は失われます。",
                    "リセット確認",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    ResetProject(_groupId, _currentProjectId);
                    MessageBox.Show("プロジェクトをリセットしました。", "リセット完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    
                    // リセット後、PowerPointプレゼンテーションを再読み込み
                    OpenProjectDocument(_currentProjectId, _groupId);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResetButton_Click] Error: {ex.Message}");
                MessageBox.Show($"リセット中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void ResetProject(int groupId, int projectId)
        {
            try
            {
                // リセット対象ファイルのパスを取得
                // App.configからパスを読み込む（存在しない場合はデフォルト値を使用）
                string basePath = ConfigurationManager.AppSettings["PowerPointDataPath"] 
                    ?? @"C:\MOSTest\PowerPoint365";
                string tabFolder = Path.Combine(basePath, $"Tab{groupId}");
                
                // Project{projectId}.pptx または Project{projectId}.ppt を検索
                string[] possibleNames = { $"Project{projectId}.pptx", $"Project{projectId}.ppt" };
                string projectFilePath = null;
                
                foreach (var fileName in possibleNames)
                {
                    string fullPath = Path.Combine(tabFolder, fileName);
                    if (File.Exists(fullPath))
                    {
                        projectFilePath = fullPath;
                        break;
                    }
                }
                
                if (string.IsNullOrEmpty(projectFilePath))
                {
                    // ファイルが見つからない場合は固定パスを生成
                    projectFilePath = Path.Combine(tabFolder, $"Project{projectId}.pptx");
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] ファイルが見つかりません。作成します: {projectFilePath}");
                }
                
                // テンプレートファイルのパスを検索
                string templatesFolder = Path.Combine(basePath, "Templates", $"Tab{groupId}");
                string templatePath = null;
                string fileExtension = ".pptx";
                
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
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] テンプレートファイルを見つけました: {templatePath}");
                            break;
                        }
                    }
                }
                
                // 固定パターンで検索
                if (string.IsNullOrEmpty(templatePath))
                {
                    string[] patterns = {
                        Path.Combine(basePath, "Templates", $"Tab{groupId}", $"project{projectId}{fileExtension}"),
                        Path.Combine(basePath, "Templates", $"project{projectId}{fileExtension}"),
                        Path.Combine(basePath, "Templates", $"Tab{groupId}", $"Tab{groupId}_project{projectId}{fileExtension}")
                    };
                    
                    foreach (var pattern in patterns)
                    {
                        if (File.Exists(pattern))
                        {
                            templatePath = pattern;
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] テンプレートファイルを見つけました（固定パターン）: {templatePath}");
                            break;
                        }
                    }
                }
                
                if (string.IsNullOrEmpty(templatePath))
                {
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] テンプレートファイルが見つかりません: {templatesFolder}");
                    throw new FileNotFoundException($"テンプレートファイルが見つかりません: {templatesFolder}");
                }
                
                // テンプレートファイルを読み取り専用で保護
                FileInfo templateFileInfo = new FileInfo(templatePath);
                if (!templateFileInfo.IsReadOnly)
                {
                    templateFileInfo.IsReadOnly = true;
                }
                
                // PowerPointアプリケーションが開いている場合はプレゼンテーションを閉じる
                CloseAllPowerPointPresentations();
                
                // テンプレートファイルをプロジェクトファイルにコピー
                // コピー先のファイルが読み取り専用の場合は属性を変更
                if (File.Exists(projectFilePath))
                {
                    FileInfo projectFileInfo = new FileInfo(projectFilePath);
                    if (projectFileInfo.IsReadOnly)
                    {
                        projectFileInfo.IsReadOnly = false;
                    }
                }
                
                File.Copy(templatePath, projectFilePath, overwrite: true);
                System.Diagnostics.Debug.WriteLine($"[ResetProject] テンプレートファイルをコピーしました: {templatePath} → {projectFilePath}");
                
                // Initialフォルダにもコピー
                string initialFolderPath = Path.Combine(tabFolder, "Initial");
                string initialFilePath = Path.Combine(initialFolderPath, $"project{projectId}{fileExtension}");
                
                if (!Directory.Exists(initialFolderPath))
                {
                    Directory.CreateDirectory(initialFolderPath);
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] Initialフォルダを作成しました: {initialFolderPath}");
                }
                
                // Initialフォルダのファイルが既に存在する場合、読み取り専用属性を解除
                if (File.Exists(initialFilePath))
                {
                    try
                    {
                        FileInfo initialFileInfo = new FileInfo(initialFilePath);
                        if (initialFileInfo.IsReadOnly)
                        {
                            initialFileInfo.IsReadOnly = false;
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Initialファイルの読み取り専用属性を解除しました: {initialFilePath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Initialファイルの属性変更エラー（続行）: {ex.Message}");
                    }
                }
                
                // リトライロジックでコピー（ファイルロックされている場合に備える）
                int retryCount = 0;
                const int maxRetries = 5;
                bool copied = false;
                while (!copied && retryCount < maxRetries)
                {
                    try
                    {
                        File.Copy(projectFilePath, initialFilePath, overwrite: true);
                        copied = true;
                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Initialフォルダにコピーしました: {initialFilePath}");
                    }
                    catch (IOException ex) when (retryCount < maxRetries - 1)
                    {
                        // ファイルがロックされている場合は少し待ってリトライ
                        retryCount++;
                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Initialファイルコピーリトライ {retryCount}/{maxRetries}: {ex.Message}");
                        Thread.Sleep(200); // 200ms待機
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        // アクセス権限エラーの場合もリトライ
                        retryCount++;
                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Initialファイルコピーリトライ {retryCount}/{maxRetries} (UnauthorizedAccess): {ex.Message}");
                        Thread.Sleep(200); // 200ms待機
                    }
                }
                
                if (!copied)
                {
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] 警告: Initialフォルダへのコピーに失敗しましたが、リセットは完了しています: {initialFilePath}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResetProject] Error: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 開いているすべてのPowerPointプレゼンテーションを保存する（閉じない）。
        /// </summary>
        private void SaveAllPowerPointPresentations()
        {
            try
            {
                PowerPointApp pptApp = null;
                try
                {
                    pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[SaveAllPowerPointPresentations] PowerPointアプリケーションが見つかりません");
                    return;
                }
                for (int i = 1; i <= pptApp.Presentations.Count; i++)
                {
                    try
                    {
                        var pres = pptApp.Presentations[i];
                        pres.Save();
                        System.Diagnostics.Debug.WriteLine($"[SaveAllPowerPointPresentations] 保存しました: {pres.Name}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[SaveAllPowerPointPresentations] 保存エラー: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SaveAllPowerPointPresentations] Error: {ex.Message}");
            }
        }

        private void CloseAllPowerPointPresentations()
        {
            try
            {
                PowerPointApp pptApp = null;
                try
                {
                    pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                }
                catch
                {
                    // PowerPointが起動していない場合は何もしない
                    System.Diagnostics.Debug.WriteLine("[CloseAllPowerPointPresentations] PowerPointアプリケーションが見つかりません");
                    return;
                }
                
                // すべてのプレゼンテーションを閉じる
                while (pptApp.Presentations.Count > 0)
                {
                    PowerPointPresentation openPres = null;
                    try
                    {
                        openPres = pptApp.Presentations[1]; // 1-based index
                        try { openPres.Save(); System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] プレゼンテーションを保存しました: {openPres.Name}"); } catch { }
                        openPres.Close();
                        System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] プレゼンテーションを閉じました: {openPres.Name}");
                    }
                    catch (COMException comEx) when (comEx.HResult == unchecked((int)0x80010108)) // RPC_E_DISCONNECTED
                    {
                        // 既に切断されている場合は無視して続行
                        System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] プレゼンテーションは既に切断されています（無視）: {comEx.Message}");
                        // ループから抜けるために、Presentations.Countを確認する前にbreak
                        break;
                    }
                    catch (Exception closeEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] プレゼンテーションを閉じる際のエラー: {closeEx.Message}");
                        // エラーが発生しても次のプレゼンテーションを試すため、breakしない
                        // ただし、無限ループを避けるために、Presentations.Countが変わらない場合はbreak
                        if (pptApp.Presentations.Count > 0)
                        {
                            try
                            {
                                // 次のプレゼンテーションを取得して再試行
                                var nextPres = pptApp.Presentations[1];
                                if (nextPres == openPres)
                                {
                                    // 同じプレゼンテーションが返された場合はループから抜ける
                                    break;
                                }
                            }
                            catch
                            {
                                break;
                            }
                        }
                    }
                    finally
                    {
                        try
                        {
                            if (openPres != null)
                            {
                                Marshal.ReleaseComObject(openPres);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] Error: {ex.Message}");
            }
        }
        
        protected override void OnClosed(EventArgs e)
        {
            _timer?.Stop();
            _projectTimer?.Stop();
            base.OnClosed(e);
        }
    }
    
    public class ProjectData
    {
        public List<ProjectInfo> Projects { get; set; }
    }
    
    public class ProjectInfo
    {
        public int ProjectId { get; set; }
        public List<TaskInfo> Tasks { get; set; }
    }
    
    public class TaskInfo
    {
        public int TaskId { get; set; }
        public string Description { get; set; }
    }
}


