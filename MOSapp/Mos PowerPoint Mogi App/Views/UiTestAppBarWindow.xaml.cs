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
using System.Threading.Tasks;
using System.Reflection;
using System.Windows.Interop;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using PowerPointPresentation = Microsoft.Office.Interop.PowerPoint.Presentation;
using Microsoft.Office.Interop.PowerPoint;
using System.Configuration;
using Libraries.Group1;
using Libraries;

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
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();

        private const int SW_RESTORE = 9;
        
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

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int nIndex);

        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        // アプリバーの高さは 1920×1080 基準の設計値（物理ピクセル）。Excel と同値。
        private const int APP_BAR_HEIGHT_BASE = 258;
        private const int DESIGN_SCREEN_HEIGHT = 1080;

        /// <summary>物理ピクセル単位の画面幅。</summary>
        private static int PhysicalScreenWidth => GetSystemMetrics(SM_CXSCREEN);

        /// <summary>物理ピクセル単位の画面高さ。</summary>
        private static int PhysicalScreenHeight => GetSystemMetrics(SM_CYSCREEN);

        /// <summary>
        /// 実際の画面高さに合わせてスケールしたアプリバー高さ（物理ピクセル）。
        /// Excel と同じ画面占有比率を維持する。
        /// </summary>
        private static int AppBarHeightPhysical =>
            (int)Math.Round(PhysicalScreenHeight * (double)APP_BAR_HEIGHT_BASE / DESIGN_SCREEN_HEIGHT);

        /// <summary>PowerPoint ウィンドウの高さ = 画面高さ − アプリバー高さ（物理ピクセル）。</summary>
        private static int PowerPointHeightPhysical => PhysicalScreenHeight - AppBarHeightPhysical;
        
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
        /// <summary>最終プロジェクト完了時に、タスク情報を消す前へ確定した破壊判定。</summary>
        private bool _lastTaskDestructiveBaselineConfirmed;
        bool _resetCloseUsedProcessKill;
        bool _resetReopenAfterQuit;
        private int _currentTaskId = 1;
        private bool _isMovingToNextProject;
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
        private List<System.Windows.Controls.Button> _dynamicTaskButtons = new List<System.Windows.Controls.Button>(); // 動的に生成されたタスクボタン（9番目以降）
        private readonly Action _onScoreClick; // 採点ボタン押下時（プロジェクト一覧の採点と同じ処理を実行）
        private bool _fromResultWindow = false; // 結果画面からタスクに飛んできたかどうか
        private ResultWindow _resultWindow = null; // 結果画面への参照
        private readonly HashSet<string> _initialWrongTaskKeys = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _retryTaskKeys = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _preparedRetryTaskKeys = new HashSet<string>(StringComparer.Ordinal);
        private bool _isScoring;
        private bool _isScoreResultOpen;
        private bool _nextProjectPendingAfterScoreResult;
        private bool _examResultPendingAfterScoreResult;
        private Window _instantScoringOverlay;
        private Timer _instantScoringNoticeTimer;
        private bool _instantScoringFinished;

        public UiTestAppBarWindow(int projectId = 1, int groupId = 1, bool showScoreButton = false, bool showPauseButton = false, Action onScoreClick = null)
        {
            InitializeComponent();
            _currentProjectId = projectId;
            _groupId = groupId; // グループIDを保存
            _onScoreClick = onScoreClick;
            var scoreBtn = FindName("ScoreButton") as System.Windows.Controls.Button;
            if (scoreBtn != null)
                scoreBtn.Visibility = showScoreButton ? Visibility.Visible : Visibility.Collapsed;
            var pauseBtn = FindName("PauseButton") as System.Windows.Controls.Button;
            if (pauseBtn != null)
                pauseBtn.Visibility = showPauseButton ? Visibility.Visible : Visibility.Collapsed;
            InitializeTimer();
            InitializeProjectTimer();
            PowerPointChecker1_1.ResetTask4SlideDeletionState();
            LoadClipboardTargets(); // クリップボード対象を先に読み込む
            LoadTasks();
            UpdateTaskDisplay();
            WriteCurrentTaskFile();
            SetWindowPosition();
            // 注意: PowerPointプレゼンテーションはMainViewModelのExecuteOpenProjectで既に開かれている
            // ここでは開かない（PositionPowerPointWindowはSetWindowPositionで呼ばれる）
            this.Activated += (s, e) =>
            {
                try { ScoreResultWindow.TryBringOpenToFront(); }
                catch { }
            };
            this.PreviewMouseDown += (s, e) =>
            {
                try { ScoreResultWindow.TryBringOpenToFront(); }
                catch { }
            };
        }
        
        private void SetWindowPosition()
        {
            // PowerPointウィンドウを配置
            PositionPowerPointWindow();
            ScoreResultWindow.TryBringOpenToFront();

            int screenW = PhysicalScreenWidth;
            int screenH = PhysicalScreenHeight;
            int barH = AppBarHeightPhysical;
            int barTop = screenH - barH;
            
            // ウィンドウハンドルを取得
            IntPtr hWnd = new WindowInteropHelper(this).Handle;
            if (hWnd == IntPtr.Zero)
            {
                // ハンドルが取得できない場合は WPF プロパティで近似配置
                this.Width = screenW;
                this.Height = barH;
                this.Left = 0;
                this.Top = barTop;
                this.Topmost = true;
                return;
            }

            // 現在のウィンドウサイズを取得して境界線のサイズを計算
            GetWindowRect(hWnd, out RECT windowRect);
            GetClientRect(hWnd, out RECT clientRect);

            int borderWidth = (windowRect.right - windowRect.left) - clientRect.right;
            int borderHeight = (windowRect.bottom - windowRect.top) - clientRect.bottom;

            // アプリバーを画面下部に配置（Excel と同じ解像度連動）
            int x = -borderWidth / 2;
            int y = barTop - borderHeight / 2;
            int width = screenW + borderWidth;
            int height = barH + borderHeight;

            MoveWindow(hWnd, x, y, width, height, true);
            
            // ウィンドウを最前面に表示
            this.Topmost = true;
        }

        private void AdjustScreenButton_Click(object sender, RoutedEventArgs e)
        {
            ScoreResultWindow.TryBringOpenToFront();
            Dispatcher.BeginInvoke(new Action(SetWindowPosition), DispatcherPriority.Background);
        }
        
        private void PositionPowerPointWindow()
        {
            // リトライ＋Sleep を UI スレッドで行うとフリーズするため、バックグラウンドで実行する
            int screenW = PhysicalScreenWidth;
            int pptHeight = PowerPointHeightPhysical;

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var pptProcesses = Process.GetProcessesByName("POWERPNT");
                    if (pptProcesses.Length == 0)
                    {
                        System.Diagnostics.Debug.WriteLine("[AppBarWindow] PowerPoint process not found");
                        return;
                    }

                    IntPtr pptHwnd = IntPtr.Zero;
                    uint processId = 0;
                    using (Process pptProcess = pptProcesses[0])
                    {
                        processId = (uint)pptProcess.Id;
                        int retryCount = 0;
                        const int maxRetries = 20;

                        while (pptHwnd == IntPtr.Zero && retryCount < maxRetries)
                        {
                            EnumWindows((windowHandle, lParam) =>
                            {
                                GetWindowThreadProcessId(windowHandle, out uint windowProcessId);
                                if (windowProcessId == processId)
                                {
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
                    }

                    if (pptHwnd != IntPtr.Zero)
                    {
                        try { ShowWindow(pptHwnd, SW_RESTORE); } catch { }

                        GetWindowRect(pptHwnd, out RECT pptWindowRect);
                        GetClientRect(pptHwnd, out RECT pptClientRect);

                        int pptBorderWidth = (pptWindowRect.right - pptWindowRect.left) - pptClientRect.right;
                        int pptBorderHeight = (pptWindowRect.bottom - pptWindowRect.top) - pptClientRect.bottom;

                        int pptX = -pptBorderWidth / 2;
                        int pptY = -pptBorderHeight / 2;
                        int pptWidth = screenW + pptBorderWidth;
                        int pptHeightWithBorder = pptHeight + pptBorderHeight;

                        MoveWindow(pptHwnd, pptX, pptY, pptWidth, pptHeightWithBorder, true);
                        System.Diagnostics.Debug.WriteLine($"[AppBarWindow] PowerPoint window positioned: {pptWidth}x{pptHeightWithBorder} at ({pptX}, {pptY})");
                        ScoreResultWindow.TryBringOpenToFront();
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
            });
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

        private void BeginScoringSession()
        {
            _isScoring = true;
            _timer?.Stop();
            _projectTimer?.Stop();
            System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] BeginScoringSession");
        }

        private void EndScoringSession(bool restartTimers = true)
        {
            _isScoring = false;
            try
            {
                WriteCurrentTaskFile(force: true);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] EndScoringSession WriteCurrentTaskFile: " + ex.Message);
            }

            if (restartTimers && !MainWindow.IsTimerDisabled && !_isPaused)
            {
                _timer?.Start();
                _projectTimer?.Start();
            }

            System.Diagnostics.Debug.WriteLine($"[UiTestAppBarWindow] EndScoringSession restartTimers={restartTimers}");
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
                if (_isScoreResultOpen)
                {
                    _examResultPendingAfterScoreResult = true;
                    return;
                }
                // 試験終了処理：結果画面を表示
                ShowResultWindowAsync();
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
                reviewWindow.OnShowResultRequested = (selectedIds, rangeLabel) =>
                {
                    ShowResultWindowAsync(selectedIds, rangeLabel);
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
        /// 結果画面を表示する（方式A: 表示前に全プロジェクトを採点してから結果を渡す）
        /// </summary>
        private async void ShowResultWindowAsync(IReadOnlyCollection<int> projectIds = null, string rangeLabel = null)
        {
            if (_isScoring)
            {
                System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] Duplicate scoring request ignored.");
                return;
            }

        ScoringProgressOverlay overlay = null;
        DispatcherTimer overlayKeepOnTopTimer = null;
        try
        {
                System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] Showing result window (scoring all projects first)");

                BeginScoringSession();
                
                overlay = CreateScoringProgressOverlay();
                overlay.Window.Show();
                overlay.Window.Activate();

                // PowerPoint 側が前面化しても、採点中表示が背面に回らないように保護する。
                overlayKeepOnTopTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
                overlayKeepOnTopTimer.Tick += (_, __) =>
                {
                    if (overlay?.Window == null || !overlay.Window.IsVisible) return;
                    if (!overlay.Window.IsActive)
                    {
                        overlay.Window.Topmost = false;
                        overlay.Window.Topmost = true;
                        overlay.Window.Activate();
                    }
                };
                overlayKeepOnTopTimer.Start();
                await Task.Yield();
                
                // 全プロジェクトを採点（バックグラウンドで実行）
                Dictionary<int, List<bool>> allResults = null;
                HashSet<int> scoringIds = projectIds != null && projectIds.Count > 0
                    ? new HashSet<int>(projectIds)
                    : null;
                await Task.Run(() =>
                {
                    try
                    {
                        allResults = ScoreAllProjects(
                            scoringIds,
                            (message, completed, total) => Dispatcher.BeginInvoke(new Action(() =>
                                overlay?.Update(message, completed, total))));
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ScoreAllProjects] Error: {ex.Message}");
                    }
                    finally
                    {
                        try { ClosePowerPointAfterBatchScoring(); }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ScoreAllProjects] Close presentations: {ex.Message}");
                        }
                    }
                });
                
                if (overlay != null)
                {
                    overlayKeepOnTopTimer?.Stop();
                    try { overlay.Window.Close(); } catch { }
                    overlay = null;
                }
                
                if (allResults == null)
                    allResults = new Dictionary<int, List<bool>>();

                if (!string.IsNullOrWhiteSpace(rangeLabel) && allResults.Count > 0)
                {
                    try
                    {
                        MosPracticeClient.ScoringLogStore.Append(
                            MosPracticeClient.ScoringLogStore.SubjectPowerPoint,
                            MosPracticeClient.ScoringLogEntry.Create(rangeLabel, _groupId, allResults));
                    }
                    catch (Exception logEx)
                    {
                        System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] Scoring log save: " + logEx.Message);
                    }
                }
                
                // 結果画面を作成（採点結果を渡す）
                var resultWindow = new ResultWindow(_projectTaskCompletedStates, _projectTaskFlaggedStates, _projectTaskViewedStates, _groupId, allResults);
                resultWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                resultWindow.Topmost = true;
                
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
                
                resultWindow.OnNavigateToTask = (projectId, taskId) =>
                {
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
                this.Hide();
            }
            catch (Exception ex)
            {
                if (overlay != null)
                {
                    overlayKeepOnTopTimer?.Stop();
                    try { overlay.Window.Close(); } catch { }
                }
                System.Diagnostics.Debug.WriteLine($"[UiTestAppBarWindow] Error showing result window: {ex.Message}");
                MessageBox.Show($"結果画面の表示中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                ShowReviewPageWindow();
            }
            finally
            {
                EndScoringSession(restartTimers: false);
            }
        }

        private sealed class ScoringProgressOverlay
        {
            public Window Window { get; }
            private readonly TextBlock _message;
            private readonly ProgressBar _progress;

            public ScoringProgressOverlay(Window window, TextBlock message, ProgressBar progress)
            {
                Window = window;
                _message = message;
                _progress = progress;
            }

            public void Update(string message, int completed, int total)
            {
                if (_message != null)
                    _message.Text = total > 0 ? $"{message}（{completed}/{total}）" : message;
                if (_progress != null)
                {
                    _progress.Maximum = Math.Max(total, 1);
                    _progress.Value = Math.Max(0, Math.Min(completed, total));
                }
            }
        }

        private static ScoringProgressOverlay CreateScoringProgressOverlay()
        {
            var message = new TextBlock
            {
                Text = "採点の準備をしています...",
                FontSize = 14,
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Margin = new Thickness(0, 0, 0, 12)
            };
            var progress = new ProgressBar
            {
                Height = 14,
                IsIndeterminate = false,
                Minimum = 0,
                Maximum = 1,
                Value = 0
            };
            var window = new Window
            {
                Title = "採点中",
                Width = 360,
                Height = 150,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                Topmost = true,
                Content = new StackPanel
                {
                    Margin = new Thickness(16, 14, 16, 14),
                    VerticalAlignment = VerticalAlignment.Center,
                    Children = { message, progress }
                }
            };
            return new ScoringProgressOverlay(window, message, progress);
        }

        /// <summary>プレゼンを開き直す処理が 300ms 以上かかる場合のみ「準備中」オーバーレイを表示する（速い遷移ではチラつきを抑える）。</summary>
        private const int PrepareProjectOverlayDelayMs = 300;
        private const int VstoHeartbeatWaitMs = 8000;
        private const int ActiveVstoHeartbeatMaxAgeSeconds = 15;

        /// <summary>レジストリ排他のうえ、VSTO 心拍が新鮮になるまで待つ。失敗時は false。</summary>
        private static bool EnsureVstoReady(int timeoutMs = VstoHeartbeatWaitMs)
        {
            string issue;
            Libraries.VSTOInstallerHelper.EnsureAddInReadyForExam(out issue);
            if (!string.IsNullOrEmpty(issue))
                System.Diagnostics.Debug.WriteLine("[EnsureVstoReady] " + issue);

            if (PPLogReader.IsVstoHeartbeatFresh(ActiveVstoHeartbeatMaxAgeSeconds))
                return true;
            return PPLogReader.WaitForVstoHeartbeat(timeoutMs, ActiveVstoHeartbeatMaxAgeSeconds);
        }

        private static Window CreatePrepareProjectOverlayWindow()
        {
            return new Window
            {
                Title = "プロジェクト準備中",
                SizeToContent = SizeToContent.WidthAndHeight,
                MinWidth = 260,
                MaxWidth = 420,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                Topmost = true,
                Content = new StackPanel
                {
                    Margin = new Thickness(16, 14, 16, 14),
                    MaxWidth = 388,
                    VerticalAlignment = VerticalAlignment.Center,
                    Children =
                    {
                        new TextBlock
                        {
                            Text = "プロジェクトを開いています。\nしばらくお待ちください...",
                            FontSize = 14,
                            TextWrapping = TextWrapping.Wrap,
                            TextAlignment = TextAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Stretch,
                            MaxWidth = 356
                        }
                    }
                }
            };
        }

        private static DispatcherTimer StartPrepareOverlayKeepOnTopTimer(Window overlay)
        {
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            timer.Tick += (_, __) =>
            {
                if (overlay == null || !overlay.IsVisible) return;
                if (!overlay.IsActive)
                {
                    overlay.Topmost = false;
                    overlay.Topmost = true;
                    overlay.Activate();
                }
            };
            timer.Start();
            return timer;
        }

        /// <summary>
        /// 遅延後に準備中オーバーレイを表示しつつ遷移処理を実行する。
        /// 処理の合間で Dispatcher を yield し、オーバーレイの表示キューが処理されるようにする。
        /// </summary>
        private async Task RunWithDelayedPrepareOverlayAsync(Func<Task> transitionAsync)
        {
            bool transitionCompleted = false;
            Window overlay = null;
            DispatcherTimer overlayKeepOnTopTimer = null;
            System.Threading.Timer showOverlayTimer = null;

            void ShowOverlayIfNeeded()
            {
                if (transitionCompleted) return;
                if (overlay != null) return;
                try
                {
                    overlay = CreatePrepareProjectOverlayWindow();
                    overlay.Show();
                    overlay.Activate();
                    overlayKeepOnTopTimer = StartPrepareOverlayKeepOnTopTimer(overlay);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[RunWithDelayedPrepareOverlay] Show overlay: {ex.Message}");
                }
            }

            showOverlayTimer = new System.Threading.Timer(
                _ => Dispatcher.BeginInvoke(new Action(ShowOverlayIfNeeded)),
                null,
                PrepareProjectOverlayDelayMs,
                Timeout.Infinite);

            try
            {
                await Dispatcher.Yield(DispatcherPriority.ApplicationIdle);
                await transitionAsync();
            }
            finally
            {
                transitionCompleted = true;
                showOverlayTimer?.Dispose();
                showOverlayTimer = null;
                overlayKeepOnTopTimer?.Stop();
                overlayKeepOnTopTimer = null;
                if (overlay != null)
                {
                    try { overlay.Close(); } catch { }
                    overlay = null;
                }
            }
        }
        
        /// <summary>
        /// 全プロジェクトを採点し、projectId → タスクごとの正否リスト を返す（方式A用）
        /// </summary>
        private Dictionary<int, List<bool>> ScoreAllProjects(
            ISet<int> projectIds = null,
            Action<string, int, int> progress = null)
        {
            var results = new Dictionary<int, List<bool>>();
            PPGradingPerf.BeginSession("PowerPoint batch scoring");
            var scoringTimer = Stopwatch.StartNew();
            try
            {
            if (_projectData?.Projects == null || _projectData.Projects.Count == 0)
                return results;

            var projectsToScore = _projectData.Projects
                .Where(project => project?.Tasks != null && project.Tasks.Count > 0)
                .Where(project => projectIds == null || projectIds.Count == 0 || projectIds.Contains(project.ProjectId))
                .OrderBy(project => project.ProjectId)
                .ToList();
            int completedProjects = 0;
            progress?.Invoke("採点の準備をしています...", 0, projectsToScore.Count);

            if (!EnsureVstoReady())
            {
                System.Diagnostics.Debug.WriteLine("[ScoreAllProjects] VSTO heartbeat not ready; scoring aborted to avoid all-X false fails");
                MessageBox.Show(
                    "PowerPoint 用 VSTO アドインが応答していないため採点できません。\nアドインを有効にしてから再度実行してください。",
                    "VSTO 未準備",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return results;
            }
            
            PowerPointGrader grader = null;
            Libraries.PPLogReader.BeginBatchScoringFastPoll();
            bool allowSnapshotSkip = false;
            try
            {
                grader = new PowerPointGrader();
                var baselineTimer = Stopwatch.StartNew();
                bool taskFileExists = Libraries.PPLogReader.TryReadCurrentTaskFile(
                    out int taskFileProjectId, out _, out _, out _, out _);
                if (taskFileExists)
                    allowSnapshotSkip = grader.Connect() && grader.LogOpenTaskBaselineDiffOnce(taskFileProjectId);
                else
                    allowSnapshotSkip = _lastTaskDestructiveBaselineConfirmed;
                PPGradingPerf.Log(
                    "ScoreAllProjects.LastTaskBaseline",
                    baselineTimer.ElapsedMilliseconds,
                    $"P{(taskFileExists ? taskFileProjectId : _currentProjectId)} confirmed={allowSnapshotSkip} taskFile={taskFileExists}");
                foreach (var project in projectsToScore)
                {
                    progress?.Invoke($"プロジェクト {project.ProjectId} を採点中", completedProjects, projectsToScore.Count);
                    var projectTimer = Stopwatch.StartNew();
                    var list = new List<bool>();
                    try
                    {
                        var openTimer = Stopwatch.StartNew();
                        bool opened = OpenProjectDocument(project.ProjectId, _groupId, forBatchScoring: true);
                        PPGradingPerf.Log("ScoreAllProjects.OpenProjectDocument", openTimer.ElapsedMilliseconds, $"P{project.ProjectId} ready={opened}");
                        // OpenProjectDocument は CloseAll 後に別ファイルを開くため、採点前に ActivePresentation を必ず取り直す
                        var connectTimer = Stopwatch.StartNew();
                        bool connected = opened && grader.Connect();
                        PPGradingPerf.Log("ScoreAllProjects.Connect", connectTimer.ElapsedMilliseconds, $"P{project.ProjectId} connected={connected}");
                        if (!connected)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ScoreAllProjects] PowerPoint に接続できないかプレゼンがありません (Project {project.ProjectId})");
                            for (int i = 0; i < project.Tasks.Count; i++)
                                list.Add(false);
                            results[project.ProjectId] = list;
                            continue;
                        }
                        foreach (var task in project.Tasks.OrderBy(t => t.TaskId))
                        {
                            bool passed = false;
                            try
                            {
                                // スナップショット ID Mismatch を防ぐため、採点前に current_task を更新し、
                                // VSTO 側の snapshot 更新を短時間待ってから GradeTask を呼ぶ。
                                int attemptNo = Libraries.PPTaskAttemptRegistry.GetAttempt(project.ProjectId, task.TaskId);
                                bool snapshotReady = grader.StartTaskAndWaitForSnapshot(project.ProjectId, task.TaskId, attemptNo, 2000, 50);
                                passed = grader.GradeTask(
                                    project.ProjectId,
                                    task.TaskId,
                                    attemptNo,
                                    batchScoring: true,
                                    skipFreshSnapshotCompare: allowSnapshotSkip && snapshotReady);
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ScoreAllProjects] GradeTask {project.ProjectId}-{task.TaskId}: {ex.Message}");
                            }
                            list.Add(passed);
                        }
                        results[project.ProjectId] = list;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ScoreAllProjects] Project {project.ProjectId}: {ex.Message}");
                        for (int i = 0; i < project.Tasks.Count; i++)
                            list.Add(false);
                        if (list.Count > 0)
                            results[project.ProjectId] = list;
                    }
                    finally
                    {
                        completedProjects++;
                        PPGradingPerf.Log("ScoreAllProjects.Project", projectTimer.ElapsedMilliseconds, $"P{project.ProjectId}");
                        progress?.Invoke($"プロジェクト {project.ProjectId} の採点が完了しました", completedProjects, projectsToScore.Count);
                    }
                }
            }
            finally
            {
                Libraries.PPLogReader.EndBatchScoringFastPoll();
                try { grader?.Dispose(); } catch { }
            }
            System.Diagnostics.Debug.WriteLine($"[ScoreAllProjects] Done: {results.Count} projects");
            return results;
            }
            finally
            {
                PPGradingPerf.Log("ScoreAllProjects.Total", scoringTimer.ElapsedMilliseconds, $"projects={results.Count}");
                PPGradingPerf.EndSession();
            }
        }
        
        /// <summary>
        /// タスクに移動する（結果画面から呼び出される）
        /// </summary>
        public async void NavigateToTask(int projectId, int taskId)
        {
            if (_isScoring)
            {
                System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] Navigation ignored while scoring.");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"NavigateToTask called: ProjectId={projectId}, TaskId={taskId}");
            ScoreResultWindow.TryBringOpenToFront();

            try
            {
                // プロジェクト切り替えやジャンプ前に一旦タスク情報をクリアし、アドイン側の誤検知を防ぐ
                Libraries.PPLogReader.ClearCurrentTaskFile();

                await RunWithDelayedPrepareOverlayAsync(async () =>
                {
                    // レビューページから戻ったときは常に該当プロジェクトのプレゼンテーションを開く
                    await Dispatcher.Yield(DispatcherPriority.ApplicationIdle);
                    OpenProjectDocument(projectId, _groupId);
                    await Dispatcher.Yield(DispatcherPriority.Background);

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
                    WriteCurrentTaskFile();

                    // メインウィンドウを表示
                    this.Show();
                    this.WindowState = WindowState.Normal;
                    this.Activate();
                    this.Focus();

                    System.Diagnostics.Debug.WriteLine($"プロジェクト{projectId}のタスク{taskId}に移動しました");
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ナビゲーションエラー: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 採点結果を結果画面用の解答済み状態に反映する（方式B: 採点で合格したタスクを〇で表示するため）
        /// </summary>
        /// <param name="projectId">対象プロジェクトID</param>
        /// <param name="taskResults">採点結果（TaskNumber は 1 始まり、IsPassed で合格/不合格）</param>
        public void ApplyScoreResults(int projectId, IEnumerable<MOS_PowerPoint_app.TaskResult> taskResults)
        {
            if (taskResults == null) return;
            var list = taskResults.ToList();
            if (list.Count == 0) return;
            int arraySize = list.Max(t => t.TaskNumber);
            if (arraySize < 1) return;
            if (!_projectTaskCompletedStates.ContainsKey(projectId) || _projectTaskCompletedStates[projectId].Length < arraySize)
            {
                var newArray = new bool[arraySize];
                if (_projectTaskCompletedStates.ContainsKey(projectId))
                {
                    var old = _projectTaskCompletedStates[projectId];
                    Array.Copy(old, newArray, Math.Min(old.Length, arraySize));
                }
                _projectTaskCompletedStates[projectId] = newArray;
            }
            bool[] completedStates = _projectTaskCompletedStates[projectId];
            foreach (var task in list)
            {
                int index = task.TaskNumber - 1;
                if (index >= 0 && index < completedStates.Length)
                {
                    completedStates[index] = task.IsPassed;
                }
            }
            System.Diagnostics.Debug.WriteLine($"[ApplyScoreResults] プロジェクト{projectId}: {list.Count(r => r.IsPassed)}/{list.Count} を解答済みに反映しました");
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
                if (_isScoreResultOpen)
                {
                    _nextProjectPendingAfterScoreResult = true;
                    return;
                }
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
        
        private async void ScoreButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isScoring)
                return;

            var scoreButton = sender as System.Windows.Controls.Button;
            if (scoreButton != null)
                scoreButton.IsEnabled = false;

            try
            {
                SyncProjectToMainViewModel();
                System.Diagnostics.Debug.WriteLine($"[ScoreButton] 採点を開始: プロジェクト{_currentProjectId}, グループ{_groupId} (PowerPoint)");

                BeginScoringSession();
                _instantScoringFinished = false;
                if (_onScoreClick != null)
                {
                    const int scoringNoticeDelayMs = 300;
                    _instantScoringNoticeTimer = new Timer(_ => Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (_instantScoringFinished || _instantScoringOverlay != null)
                            return;
                        _instantScoringOverlay = CreateInstantScoringOverlay();
                        _instantScoringOverlay.Show();
                    })), null, scoringNoticeDelayMs, Timeout.Infinite);

                    // 採点本体は画面スレッドの外で行い、バーの更新を止めない。結果表示は ExecuteScore 内で画面側に戻す。
                    await Task.Run(() => _onScoreClick());
                }
                else
                    MessageBox.Show("採点機能は利用できません。プロジェクト一覧からプロジェクトを開いて採点してください。", "採点", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ScoreButton] Error: {ex.Message}");
                MessageBox.Show($"採点中にエラーが発生しました:\n{ex.Message}", "採点エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                EndScoringSession(restartTimers: true);
                CloseInstantScoringOverlay();
                if (scoreButton != null)
                    scoreButton.IsEnabled = true;
            }
        }

        /// <summary>結果ダイアログの直前に呼ぶ。採点中表示が結果を覆わないようにする。</summary>
        public void CloseInstantScoringOverlay()
        {
            _instantScoringFinished = true;
            if (_instantScoringNoticeTimer != null)
            {
                _instantScoringNoticeTimer.Dispose();
                _instantScoringNoticeTimer = null;
            }

            var overlay = _instantScoringOverlay;
            _instantScoringOverlay = null;
            if (overlay == null)
                return;
            try { overlay.Close(); } catch { }
        }

        private static Window CreateInstantScoringOverlay()
        {
            var overlay = new Window
            {
                Title = "採点中",
                Width = 320,
                Height = 140,
                WindowStyle = WindowStyle.None,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ShowInTaskbar = false,
                ResizeMode = ResizeMode.NoResize,
                Topmost = true,
                Background = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(30, 64, 175)),
                BorderThickness = new Thickness(2)
            };
            var stack = new StackPanel
            {
                Margin = new Thickness(24),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            stack.Children.Add(new TextBlock
            {
                Text = "採点中です",
                FontSize = 18,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 12),
                Foreground = new SolidColorBrush(Color.FromRgb(30, 64, 175))
            });
            stack.Children.Add(new ProgressBar
            {
                IsIndeterminate = true,
                Height = 20,
                Width = 260
            });
            overlay.Content = stack;
            return overlay;
        }

        public void ShowScoreResult(IEnumerable<MOS_PowerPoint_app.TaskResult> scoreList)
        {
            bool reviewEnabled = ReviewPageButton == null || ReviewPageButton.IsEnabled;
            bool scoreEnabled = ScoreButton == null || ScoreButton.IsEnabled;
            bool nextEnabled = NextProjectButton == null || NextProjectButton.IsEnabled;
            bool closeEnabled = CloseExamButton == null || CloseExamButton.IsEnabled;
            SetScoreResultActionsEnabled(false);
            _isScoreResultOpen = true;
            try
            {
                ScoreResultWindow.ShowResults(this, scoreList);
            }
            finally
            {
                _isScoreResultOpen = false;
                bool continueAfterResult = _examResultPendingAfterScoreResult || _nextProjectPendingAfterScoreResult;
                if (!continueAfterResult)
                    RestorePowerPointInputAfterScoreResult();
                if (ReviewPageButton != null)
                    ReviewPageButton.IsEnabled = reviewEnabled;
                if (ScoreButton != null)
                    ScoreButton.IsEnabled = scoreEnabled;
                if (NextProjectButton != null)
                    NextProjectButton.IsEnabled = nextEnabled;
                if (CloseExamButton != null)
                    CloseExamButton.IsEnabled = closeEnabled;
                if (_examResultPendingAfterScoreResult)
                {
                    _examResultPendingAfterScoreResult = false;
                    _nextProjectPendingAfterScoreResult = false;
                    ShowResultWindowAsync();
                }
                else if (_nextProjectPendingAfterScoreResult)
                {
                    _nextProjectPendingAfterScoreResult = false;
                    MoveToNextProjectWithMessage();
                }
            }
        }

        private void SetScoreResultActionsEnabled(bool enabled)
        {
            if (ReviewPageButton != null)
                ReviewPageButton.IsEnabled = enabled;
            if (ScoreButton != null)
                ScoreButton.IsEnabled = enabled;
            if (NextProjectButton != null)
                NextProjectButton.IsEnabled = enabled;
            if (CloseExamButton != null)
                CloseExamButton.IsEnabled = enabled;
        }

        /// <summary>
        /// 結果ダイアログが前面を取ったあと、PowerPoint が無効なら戻して一度だけ前面へ出す。
        /// </summary>
        private void RestorePowerPointInputAfterScoreResult()
        {
            try
            {
                IntPtr hwnd = TryGetPowerPointMainWindowHandle();
                if (hwnd == IntPtr.Zero)
                    return;
                if (!IsWindowEnabled(hwnd))
                    EnableWindow(hwnd, true);
                TryForceForeground(hwnd);
            }
            catch
            {
                /* ignore */
            }
        }

        private static IntPtr TryGetPowerPointMainWindowHandle()
        {
            Process[] pptProcesses = Process.GetProcessesByName("POWERPNT");
            try
            {
                if (pptProcesses.Length == 0)
                    return IntPtr.Zero;

                uint processId = (uint)pptProcesses[0].Id;
                IntPtr found = IntPtr.Zero;
                EnumWindows((windowHandle, lParam) =>
                {
                    GetWindowThreadProcessId(windowHandle, out uint windowProcessId);
                    if (windowProcessId != processId)
                        return true;

                    var className = new StringBuilder(256);
                    GetClassName(windowHandle, className, className.Capacity);
                    if (!className.ToString().Contains("PPTFrameClass"))
                        return true;

                    found = windowHandle;
                    return false;
                }, IntPtr.Zero);
                return found;
            }
            finally
            {
                foreach (Process process in pptProcesses)
                    process.Dispose();
            }
        }

        private static void TryForceForeground(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero)
                return;

            uint currentTid = 0;
            uint targetTid = 0;
            bool attached = false;
            try
            {
                currentTid = GetCurrentThreadId();
                targetTid = GetWindowThreadProcessId(hWnd, out _);
                if (currentTid != 0 && targetTid != 0 && currentTid != targetTid)
                    attached = AttachThreadInput(currentTid, targetTid, true);

                ShowWindow(hWnd, SW_RESTORE);
                BringWindowToTop(hWnd);
                SetForegroundWindow(hWnd);
            }
            catch
            {
                /* ignore */
            }
            finally
            {
                if (attached && currentTid != 0 && targetTid != 0)
                {
                    try { AttachThreadInput(currentTid, targetTid, false); } catch { }
                }
            }
        }
        
        /// <summary>
        /// 閉じるボタン（Excelに合わせて結果画面は表示せず、試験終了してメインに戻る）
        /// </summary>
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isScoreResultOpen)
                return;
            var result = MessageBox.Show("アプリ自体を終了します。本当にいいですか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
                return;
            PowerPointStartupInputGate.End();
            _timer?.Stop();
            _projectTimer?.Stop();
            SaveAllPowerPointPresentations();
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
                System.Diagnostics.Debug.WriteLine($"[CloseButton] PowerPoint プロセス終了エラー: {ex.Message}");
            }
            this.Close();
            var main = System.Windows.Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
            if (main != null)
            {
                main.Show();
                main.Activate();
            }
        }
        
        private void ReviewPageButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isScoreResultOpen)
                return;
            ShowReviewPageWindow();
        }

        /// <summary>
        /// 結果画面から来たフラグを設定し、ボタン表示を切り替える
        /// </summary>
        public void SetFromResultWindow(bool fromResultWindow)
        {
            _fromResultWindow = fromResultWindow;
            UpdateReviewPageButtonVisibility();
        }

        /// <summary>
        /// 結果画面への参照を保持する
        /// </summary>
        public void SetResultWindow(ResultWindow resultWindow)
        {
            _resultWindow = resultWindow;
        }

        public void SetInitialWrongTaskKeys(IEnumerable<string> keys)
        {
            _initialWrongTaskKeys.Clear();
            if (keys == null) return;
            foreach (var key in keys)
            {
                if (!string.IsNullOrWhiteSpace(key))
                    _initialWrongTaskKeys.Add(key);
            }
        }

        public bool TryPrepareRetryAttemptFromResult(int projectId, int taskId)
        {
            string key = GetTaskKey(projectId, taskId);
            if (!_initialWrongTaskKeys.Contains(key))
                return false;
            return EnsureRetryAttemptPrepared(projectId, taskId);
        }

        private void UpdateReviewPageButtonVisibility()
        {
            var reviewPageButton = FindName("ReviewPageButton") as System.Windows.Controls.Button;
            var returnToResultButton = FindName("ReturnToResultButton") as System.Windows.Controls.Button;

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
                if (_fromResultWindow)
                {
                    TryRescorePendingRetryTasks();
                }

                // 結果画面を表示
                if (_resultWindow != null && !_resultWindow.IsVisible)
                {
                    _resultWindow.Show();
                    _resultWindow.Activate();
                    _resultWindow.Focus();
                }
                else
                {
                    // 既存の結果画面を探す
                    var resultWindow = System.Windows.Application.Current.Windows.OfType<ResultWindow>().FirstOrDefault();
                    if (resultWindow != null)
                    {
                        resultWindow.Show();
                        resultWindow.Activate();
                        resultWindow.Focus();
                    }
                }

                // AppBarWindowを非表示にする
                this.Hide();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReturnToResultButton] Error: {ex.Message}");
                MessageBox.Show("結果画面の表示に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        
        
        private void LoadTasks()
        {
            try
            {
                // References\JSON を正規配置とし、実行ファイル直下を互換フォールバックにする。
                string jsonPath = PowerPointDataPathHelper.ResolveJsonPath(
                    "MOS模擬アプリ問題文一覧_PowerPoint.json");
                
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
                
                // コピー対象問題JSON（入力・追加・変更・挿入の問題）を優先して読み込む
                string copyTargetJsonName = "MOS模擬アプリ_入力追加変更挿入問題_PowerPoint.json";
                string jsonPath = PowerPointDataPathHelper.ResolveJsonPath(copyTargetJsonName);
                if (File.Exists(jsonPath))
                {
                    try
                    {
                        string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
                        var copyTargetData = JsonConvert.DeserializeObject<ProjectData>(jsonContent);
                        if (copyTargetData?.Projects != null)
                        {
                            _clipboardTargets.Clear();
                            foreach (var project in copyTargetData.Projects)
                            {
                                if (project.Tasks == null) continue;
                                _clipboardTargets[project.ProjectId] = new Dictionary<int, string>();
                                foreach (var task in project.Tasks)
                                {
                                    if (!string.IsNullOrEmpty(task.Description))
                                        _clipboardTargets[project.ProjectId][task.TaskId] = task.Description;
                                }
                            }
                            System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] JSONからクリップボード対象を{_clipboardTargets.Sum(p => p.Value.Count)}件読み込みました: {jsonPath}");
                            return;
                        }
                    }
                    catch (Exception exJson)
                    {
                        System.Diagnostics.Debug.WriteLine($"[LoadClipboardTargets] JSON読み込み失敗、CSVにフォールバック: {exJson.Message}");
                    }
                }
                
                // CSVファイル名（JSONが無い場合のフォールバック）
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
            WriteCurrentTaskFile();
            SyncProjectToMainViewModel();
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
            
            // コピーできる部分を""で囲んでから表示（「…」と入力 の 「」内をコピー対象に）
            string descriptionForDisplay = clipboardTargetDescription ?? currentTask.Description;
            descriptionForDisplay = WrapCopyablePartInQuotes(descriptionForDisplay);
            if (clipboardTargetDescription != null)
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] SetTextWithUnderlineを呼び出します（クリップボード対象）");
            else
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] SetTextWithUnderlineを呼び出します（通常の問題文）");
            SetTextWithUnderline(taskDescriptionTextBlock, descriptionForDisplay);
            
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
        /// 問題文からコピーすべき部分（「…」と入力 の 「」内など）を抽出し、半角ダブルクォートで囲んだ文字列を返します。
        /// 既に"が含まれる問題文はそのまま返します。
        /// </summary>
        private static string WrapCopyablePartInQuotes(string description)
        {
            if (string.IsNullOrEmpty(description)) return description;
            if (description.IndexOf('"') >= 0) return description;
            // 「…」と入力 の 「」内を "…" に変換（コピーできる部分として表示）
            return Regex.Replace(description, @"「([^」]*)」と入力", "\"$1\"と入力");
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

            // XAMLで定義されているタスクボタンの数（8つ）まで処理
            int maxStaticButtonCount = 8; // XAMLで定義されているボタンの数（最大8想定）
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
            
            // 9つ目以降のタスクボタンを動的に生成（タスク数が8より多い場合の保険）
            var container = FindName("TaskButtonsContainer") as System.Windows.Controls.StackPanel;
            if (container != null && maxTaskCount > maxStaticButtonCount)
            {
                // 既存の動的ボタンを削除
                foreach (var btn in _dynamicTaskButtons)
                {
                    container.Children.Remove(btn);
                }
                _dynamicTaskButtons.Clear();
                
                // TaskButton8のインデックスを取得
                var taskButton8 = FindName("TaskButton8") as System.Windows.Controls.Button;
                int insertIndex = taskButton8 != null ? container.Children.IndexOf(taskButton8) + 1 : container.Children.Count - 1;
                
                // 9つ目以降のボタンを生成
                for (int i = maxStaticButtonCount + 1; i <= maxTaskCount; i++)
                {
                    var button = CreateTaskButton(i);
                    _dynamicTaskButtons.Add(button);
                    
                    // TaskButton8の後に挿入
                    container.Children.Insert(insertIndex, button);
                    insertIndex++; // 次の挿入位置を更新
                }
            }
            else if (container != null && maxTaskCount <= maxStaticButtonCount)
            {
                // タスク数が8以下になった場合は動的ボタンを削除
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
            ScoreResultWindow.TryBringOpenToFront();
            if (_currentTaskId > 1)
            {
                _currentTaskId--;
                UpdateTaskDisplay();
                WriteCurrentTaskFile();
            }
        }
        
        private void NextTask_Click(object sender, RoutedEventArgs e)
        {
            ScoreResultWindow.TryBringOpenToFront();
            if (_tasks != null && _currentTaskId < _tasks.Count)
            {
                _currentTaskId++;
                UpdateTaskDisplay();
                WriteCurrentTaskFile();
            }
        }
        
        private void TaskButton_Click(object sender, RoutedEventArgs e)
        {
            ScoreResultWindow.TryBringOpenToFront();
            if (_isScoring)
            {
                System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] Task button ignored while scoring.");
                return;
            }

            if (sender is System.Windows.Controls.Button button && button.Tag != null)
            {
                int taskId = int.Parse(button.Tag.ToString());
                if (taskId >= 1 && taskId <= _tasks.Count)
                {
                    _currentTaskId = taskId;
                    UpdateTaskDisplay();
                    WriteCurrentTaskFile();
                }
            }
        }
        
        /// <summary>
        /// VSTO アドインが現在タスクを参照するため、共有ファイルに ProjectId,TaskId を書き出す。
        /// </summary>
        private void WriteCurrentTaskFile(bool force = false)
        {
            if (_isScoring && !force)
            {
                System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] WriteCurrentTaskFile skipped (scoring in progress)");
                return;
            }

            _lastTaskDestructiveBaselineConfirmed = false;
            try
            {
                var flags = Libraries.PPTaskValidationConfig.GetExemptFlags(_currentProjectId, _currentTaskId);
                if (_fromResultWindow)
                {
                    EnsureRetryAttemptPrepared(_currentProjectId, _currentTaskId);
                }
                int attemptNo = GetCurrentTaskAttempt(_currentProjectId, _currentTaskId);
                Libraries.PPLogReader.WriteCurrentTaskFile(_currentProjectId, _currentTaskId, (int)flags, attemptNo, 0);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[UiTestAppBarWindow] WriteCurrentTaskFile: " + ex.Message);
            }
        }
        
        private async void NextProject_Click(object sender, RoutedEventArgs e)
        {
            if (_isScoreResultOpen)
                return;
            ScoreResultWindow.TryBringOpenToFront();
            await MoveToNextProjectAsync();
        }
        
        private async Task MoveToNextProjectAsync()
        {
            if (_isMovingToNextProject)
                return;

            _isMovingToNextProject = true;
            try
            {
                int maxProjectId = _projectData?.Projects?.Max(p => p.ProjectId) ?? 1;
                // 最終プロジェクトで「次」は「すべて完了」になるため、プレゼン準備オーバーレイは出さない
                if (_currentProjectId >= maxProjectId)
                    await MoveToNextProjectCoreAsync();
                else
                    await RunWithDelayedPrepareOverlayAsync(MoveToNextProjectCoreAsync);
            }
            finally
            {
                _isMovingToNextProject = false;
            }
        }

        private async Task MoveToNextProjectCoreAsync()
        {
            // 破壊的操作の記録は VSTO タスク境界と Grader ゲートに一本化（Word 方式）。
            // プロジェクト遷移時は destructive_errors.log へ追記しない。

            // current_task を消す前に、印刷系タスク（5-1/11-7）は同期再評価して証跡を確定する。
            // 最終タスク 11-7 での高速遷移時の取りこぼしを防ぐため、アプリ側でも境界補完する。
            try
            {
                TryFinalizePrintEvidenceBeforeProjectTransition();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MoveToNextProject] Print evidence finalize error: {ex.Message}");
            }

            await Dispatcher.Yield(DispatcherPriority.Background);

            int maxProjectId = _projectData?.Projects?.Max(p => p.ProjectId) ?? 1;
            bool finishingLastProject = _currentProjectId >= maxProjectId;
            if (finishingLastProject)
                _lastTaskDestructiveBaselineConfirmed = ConfirmLastOpenTaskBaseline(_currentProjectId);

            // プロジェクト切り替え前にタスク情報をクリアし、アドイン側の破壊的操作チェックをスキップさせる
            Libraries.PPLogReader.ClearCurrentTaskFile();

            await Dispatcher.Yield(DispatcherPriority.Background);

            bool enableProjectBackup = false;
            bool.TryParse(ConfigurationManager.AppSettings["EnableProjectBackup"], out enableProjectBackup);
            if (enableProjectBackup)
            {
                // 現在開いているプレゼンテーションを日付・時間付きバックアップフォルダに保存（MMdd_HHmm）
                string basePath = PowerPointDataPathHelper.GetDataRoot();
                string backupSubdir = DateTime.Now.ToString("MMdd_HHmm");
                string backupFolder = Path.Combine(basePath, $"Tab{_groupId}", "backup", backupSubdir);
                try
                {
                    PowerPointApp pptApp = null;
                    try
                    {
                        pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch
                    {
                        // PowerPointが起動していない場合はスキップ
                    }
                    if (pptApp != null && pptApp.Presentations.Count > 0)
                    {
                        if (!Directory.Exists(backupFolder))
                            Directory.CreateDirectory(backupFolder);
                        string backupFilePath = Path.Combine(backupFolder, $"Project{_currentProjectId}.pptx");
                        try
                        {
                            PowerPointPresentation pres = pptApp.Presentations[1];
                            pres.SaveCopyAs(backupFilePath);
                            pres.Save();
                            System.Diagnostics.Debug.WriteLine($"[MoveToNextProject] バックアップ保存: {backupFilePath}");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[MoveToNextProject] バックアップ保存エラー: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[MoveToNextProject] バックアップ処理エラー: {ex.Message}");
                }
            }

            await Dispatcher.Yield(DispatcherPriority.Background);

            if (finishingLastProject)
            {
                System.Diagnostics.Debug.WriteLine($"プロジェクト{maxProjectId}が完了しました");
                MessageBox.Show("すべてのプロジェクトが完了しました。", "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                ShowReviewPageWindow();
                return;
            }

            // 次の実プロジェクトへ移動する。最大番号の次は存在しない。
            _currentProjectId++;

            await Dispatcher.Yield(DispatcherPriority.Background);

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

        /// <summary>
        /// 最終プロジェクトのタスク情報を消す前に、開いている最終タスクの破壊判定を確定する。
        /// </summary>
        private static bool ConfirmLastOpenTaskBaseline(int projectId)
        {
            try
            {
                using (var grader = new PowerPointGrader())
                {
                    return grader.Connect() && grader.LogOpenTaskBaselineDiffOnce(projectId);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ConfirmLastOpenTaskBaseline] " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// プロジェクト遷移直前に 5-1 の印刷設定を同期評価し、条件一致なら証跡ログを明示追記する。
        /// </summary>
        private void TryFinalizePrintEvidenceBeforeProjectTransition()
        {
            bool isTask5_1 = _currentProjectId == 5 && _currentTaskId == 1;
            if (!isTask5_1)
                return;

            int attemptNo = GetCurrentTaskAttempt(_currentProjectId, _currentTaskId);
            string marker = "[Task5-1] Print";
            if (Libraries.PPLogReader.HasMarkerWithinTask(_currentProjectId, _currentTaskId, attemptNo, marker))
                return;

            PowerPointApp pptApp = null;
            PowerPointPresentation pres = null;
            PrintOptions po = null;
            try
            {
                try
                {
                    pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                }
                catch
                {
                    return;
                }
                if (pptApp == null || pptApp.Presentations == null || pptApp.Presentations.Count < 1)
                    return;

                pres = pptApp.ActivePresentation;
                if (pres == null)
                    return;

                po = pres.PrintOptions;
                if (po == null)
                    return;

                bool matched = false;
                const int maxRetries = 20; // up to ~2s (100ms * 20)
                const int retryIntervalMs = 100;
                for (int retry = 0; retry < maxRetries && !matched; retry++)
                {
                    int outputType = (int)po.OutputType;
                    int copies = po.NumberOfCopies;
                    bool collate = Convert.ToInt32(po.Collate) == (int)Microsoft.Office.Core.MsoTriState.msoTrue;

                    if (isTask5_1)
                    {
                        matched = outputType == (int)PpPrintOutputType.ppPrintOutputThreeSlideHandouts
                                  && copies == 4
                                  && collate;
                    }

                    if (!matched && retry + 1 < maxRetries)
                    {
                        Thread.Sleep(retryIntervalMs);
                    }
                }

                if (!matched)
                    return;

                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string line = $"[{timestamp}] [Task {_currentProjectId}-{_currentTaskId}-{attemptNo}] {marker}";
                string evidencePath = Libraries.PPLogReader.GetTaskEvidenceLogPath();
                string mainLogPath = Libraries.PPLogReader.GetLogFilePath();
                File.AppendAllText(evidencePath, line + Environment.NewLine, new UTF8Encoding(false));
                File.AppendAllText(mainLogPath, line + Environment.NewLine, new UTF8Encoding(false));
                System.Diagnostics.Debug.WriteLine($"[MoveToNextProject] Print evidence finalized: {line}");
            }
            finally
            {
                if (po != null) { try { Marshal.ReleaseComObject(po); } catch { } }
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
                if (pptApp != null) { try { Marshal.ReleaseComObject(pptApp); } catch { } }
            }
        }
        
        private bool OpenProjectDocument(int projectId, int groupId)
        {
            return OpenProjectDocument(projectId, groupId, forBatchScoring: false);
        }

        private bool OpenProjectDocument(int projectId, int groupId, bool forBatchScoring, bool logResetPerf = false, bool deferFailureUi = false)
        {
            var reopenSw = Stopwatch.StartNew();
            bool reopenLogged = false;
            string reopenPath = _resetReopenAfterQuit ? "fallback" : "main";
            try
            {
                string tabFolder = PowerPointDataPathHelper.GetTabFolder(groupId);
                string filePath = PowerPointDataPathHelper.GetWorkingProjectPath(groupId, projectId);
                if (!File.Exists(filePath))
                {
                    string legacyPptPath = Path.Combine(tabFolder, $"Project{projectId}.ppt");
                    filePath = File.Exists(legacyPptPath) ? legacyPptPath : null;
                }
                
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    System.Diagnostics.Debug.WriteLine($"プロジェクト{projectId}のファイルが見つかりません: {tabFolder}");
                    if (logResetPerf && !deferFailureUi)
                    {
                        ResetPerfLog.Write("powerpoint", projectId, "reopen", 0, "main", "result=missing");
                        ResetPerfLog.Write("powerpoint", projectId, "ready", 0, "main", "result=missing");
                        reopenLogged = true;
                    }
                    return false;
                }

                // 採点スレッドと競合しないよう、閉じる〜開くまでを COM ロックで直列化する
                lock (PowerPointCheckerCommon.PowerPointComInteropSync)
                {
                    // 既に開いているプレゼンテーションがある場合は閉じる（頑健な方法を使用）
                    var closeTimer = Stopwatch.StartNew();
                    CloseAllPowerPointPresentations();
                    bool presentationsClosed = WaitUntilPresentationsClosed();
                    PPGradingPerf.Log("OpenProjectDocument.ClosePresentations", closeTimer.ElapsedMilliseconds, $"P{projectId} closed={presentationsClosed}");

                    // PowerPointアプリケーションを取得または作成（CloseAll...でプロセスが終了した可能性があるため、必要に応じて再取得）
                    PowerPointApp pptApp = null;
                    try
                    {
                        pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch
                    {
                        PPLogReader.ClearVstoHeartbeat();
                        reopenPath = "fallback";
                        pptApp = new PowerPointApp();
                        pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    }

                    // 新しいプレゼンテーションを開く
                    PowerPointPresentation presentation = null;
                    int retryCount = 0;
                    while (retryCount < 3)
                    {
                        try
                        {
                            presentation = pptApp.Presentations.Open(filePath, WithWindow: Microsoft.Office.Core.MsoTriState.msoTrue);
                            Libraries.PowerPointViewHelper.HideNotesPane(pptApp);
                            System.Diagnostics.Debug.WriteLine($"プレゼンテーションを開きました: {filePath}");
                            break;
                        }
                        catch (Exception ex)
                        {
                            retryCount++;
                            System.Diagnostics.Debug.WriteLine($"プレゼンテーションを開く際のエラー (試行 {retryCount}/3): {ex.Message}");
                            if (retryCount >= 3)
                            {
                                if (!deferFailureUi)
                                {
                                    if (logResetPerf && !reopenLogged)
                                    {
                                        ResetPerfLog.Write("powerpoint", projectId, "reopen", reopenSw.ElapsedMilliseconds, reopenPath, "result=fail");
                                        ResetPerfLog.Write("powerpoint", projectId, "ready", 0, reopenPath, "result=fail");
                                        reopenLogged = true;
                                    }
                                    MessageBox.Show($"プロジェクト{projectId}のファイルを開けませんでした。\nPowerPointを一度終了してから再度お試しください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                                return false;
                            }

                            // 1秒待機してから再試行
                            Thread.Sleep(1000);

                            // PowerPointアプリケーションの状態を確認・再取得
                            try
                            {
                                pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                            }
                            catch
                            {
                                try
                                {
                                    reopenPath = "fallback";
                                    pptApp = new PowerPointApp();
                                }
                                catch { }
                            }

                            if (pptApp != null)
                                pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                        }
                    }

                    if (presentation == null)
                    {
                        if (logResetPerf && !reopenLogged && !deferFailureUi)
                        {
                            ResetPerfLog.Write("powerpoint", projectId, "reopen", reopenSw.ElapsedMilliseconds, reopenPath, "result=fail");
                            ResetPerfLog.Write("powerpoint", projectId, "ready", 0, reopenPath, "result=fail");
                            reopenLogged = true;
                        }
                        return false;
                    }

                    // 起動経路は Presentations.Open のまま。VSTO 心拍が来るまで待ってから続行する。
                    long reopenMs = reopenSw.ElapsedMilliseconds;
                    var readyTimer = Stopwatch.StartNew();
                    bool presentationReady = WaitUntilPresentationReady(pptApp, filePath, ensureVsto: !forBatchScoring);
                    PPGradingPerf.Log("OpenProjectDocument.PresentationReady", readyTimer.ElapsedMilliseconds, $"P{projectId} ready={presentationReady}");
                    if (!presentationReady && deferFailureUi)
                        return false;

                    if (logResetPerf && !reopenLogged)
                    {
                        ResetPerfLog.Write(
                            "powerpoint",
                            projectId,
                            "reopen",
                            reopenMs,
                            reopenPath,
                            _resetReopenAfterQuit ? "quit-process" : (reopenPath == "fallback" ? "cold-start" : null));
                        reopenLogged = true;
                    }
                    if (logResetPerf)
                    {
                        ResetPerfLog.Write(
                            "powerpoint",
                            projectId,
                            "ready",
                            readyTimer.ElapsedMilliseconds,
                            reopenPath,
                            presentationReady ? "signal=vsto-heartbeat result=ok" : "signal=vsto-heartbeat result=fail");
                    }
                    if (!presentationReady)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OpenProjectDocument] Presentation not ready: Project {projectId}");
                        MessageBox.Show(
                            $"プロジェクト{projectId}の準備が完了しませんでした。\nPowerPointを一度終了してから再度お試しください。",
                            "準備未完了",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                        return false;
                    }

                    try
                    {
                        // COMオブジェクトの参照を解放
                        Marshal.ReleaseComObject(presentation);
                    }
                    catch { }

                    // PowerPointウィンドウを画面の上部2/3に配置
                    PositionPowerPointWindow();

                    System.Diagnostics.Debug.WriteLine($"プロジェクト{projectId}のプレゼンテーションを開きました: {filePath}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"プロジェクトプレゼンテーションを開く際のエラー: {ex.Message}");
                if (logResetPerf && !reopenLogged && !deferFailureUi)
                {
                    ResetPerfLog.Write("powerpoint", projectId, "reopen", reopenSw.ElapsedMilliseconds, reopenPath, "result=fail");
                    ResetPerfLog.Write("powerpoint", projectId, "ready", 0, reopenPath, "result=fail");
                }
                return false;
            }
        }

        /// <summary>Close後の固定500msは使わず、プレゼン数が0になるまで短い間隔で確認する。</summary>
        private static bool WaitUntilPresentationsClosed()
        {
            const int timeoutMs = 500;
            var wait = Stopwatch.StartNew();
            while (wait.ElapsedMilliseconds < timeoutMs)
            {
                PowerPointApp app = null;
                try
                {
                    app = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                    if (GetPresentationCountSafe(app) == 0)
                        return true;
                }
                catch (COMException)
                {
                    return true;
                }
                finally
                {
                    if (app != null)
                    {
                        try { Marshal.ReleaseComObject(app); } catch { }
                    }
                }

                Thread.Sleep(50);
            }
            return false;
        }

        /// <summary>Open後の固定800msは使わず、対象プレゼン・スライド取得・VSTO心拍で準備完了を確認する。</summary>
        private static bool WaitUntilPresentationReady(PowerPointApp pptApp, string filePath, bool ensureVsto)
        {
            const int timeoutMs = 3000;
            var wait = Stopwatch.StartNew();
            string expectedPath = NormalizeProjectPath(filePath);
            while (wait.ElapsedMilliseconds < timeoutMs)
            {
                if (IsPresentationReady(pptApp, expectedPath) && (!ensureVsto || EnsureVstoReady(250)))
                    return true;
                Thread.Sleep(50);
            }
            return IsPresentationReady(pptApp, expectedPath) && (!ensureVsto || EnsureVstoReady(250));
        }

        private static bool IsPresentationReady(PowerPointApp pptApp, string expectedPath)
        {
            PowerPointPresentation presentation = null;
            try
            {
                presentation = pptApp?.ActivePresentation;
                if (presentation == null)
                    return false;
                if (!ProjectPathsEqual(presentation.FullName, expectedPath))
                    return false;
                return presentation.Slides.Count >= 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (presentation != null)
                {
                    try { Marshal.ReleaseComObject(presentation); } catch { }
                }
            }
        }

        private static string NormalizeProjectPath(string filePath)
        {
            try { return Path.GetFullPath(filePath); }
            catch { return filePath ?? string.Empty; }
        }

        private static bool ProjectPathsEqual(string left, string right)
        {
            return string.Equals(NormalizeProjectPath(left), NormalizeProjectPath(right), StringComparison.OrdinalIgnoreCase);
        }
        
        private async void MoveToNextProjectWithMessage()
        {
            // メッセージを表示
            MessageBox.Show("5分経ったので次のプロジェクトに移動します", "時間切れ", 
                          MessageBoxButton.OK, MessageBoxImage.Information);
            
            // 次のプロジェクトに移動
            await MoveToNextProjectAsync();
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
                    // リセット前にスナップショットをクリアし、アドイン側に再取得を促す
                    Libraries.PPLogReader.ClearSnapshot();
                    ResetProject(_groupId, _currentProjectId);
                    PowerPointChecker1_1.ResetTask4SlideDeletionState();
                    MessageBox.Show("プロジェクトをリセットしました。", "リセット完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    
                    // リセット後、PowerPointプレゼンテーションを再読み込み。失敗時は一度だけ終了して開き直す。
                    if (!OpenProjectDocument(_currentProjectId, _groupId, forBatchScoring: false, logResetPerf: true, deferFailureUi: true))
                    {
                        QuitPowerPointForResetRetry();
                        _resetReopenAfterQuit = true;
                        try
                        {
                            OpenProjectDocument(_currentProjectId, _groupId, forBatchScoring: false, logResetPerf: true);
                        }
                        finally
                        {
                            _resetReopenAfterQuit = false;
                        }
                    }
                    WriteCurrentTaskFile();
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
            ResetPerfLog.Begin("powerpoint", projectId);
            _resetCloseUsedProcessKill = false;
            var closeSw = Stopwatch.StartNew();
            CloseAllPowerPointPresentations();
            ResetPerfLog.Write(
                "powerpoint",
                projectId,
                "close",
                closeSw.ElapsedMilliseconds,
                _resetCloseUsedProcessKill ? "fallback" : "main",
                _resetCloseUsedProcessKill ? "process-kill" : null);

            var copySw = Stopwatch.StartNew();
            try
            {
                MOS_PowerPoint_app.PowerPointProjectResetHelper.ResetProject(groupId, projectId);
                ResetPerfLog.Write("powerpoint", projectId, "copy", copySw.ElapsedMilliseconds, "main");
            }
            catch
            {
                ResetPerfLog.Write("powerpoint", projectId, "copy", copySw.ElapsedMilliseconds, "main", "result=fail");
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
                lock (PowerPointCheckerCommon.PowerPointComInteropSync)
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
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SaveAllPowerPointPresentations] Error: {ex.Message}");
            }
        }

        /// <summary>Presentations.Count 取得時の COM 例外を握りつぶす（切断直後など）。</summary>
        private static int GetPresentationCountSafe(PowerPointApp pptApp)
        {
            if (pptApp == null) return 0;
            try
            {
                return pptApp.Presentations.Count;
            }
            catch (COMException)
            {
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>リセットの開き直し失敗時だけ、保存せず PowerPoint を終了する。</summary>
        void QuitPowerPointForResetRetry()
        {
            PowerPointApp pptApp = null;
            try
            {
                lock (PowerPointCheckerCommon.PowerPointComInteropSync)
                {
                    try
                    {
                        pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch (COMException)
                    {
                        pptApp = null;
                    }

                    if (pptApp != null)
                    {
                        try { pptApp.DisplayAlerts = PpAlertLevel.ppAlertsNone; } catch { }
                        CloseAllPowerPointPresentations();
                        try { pptApp.Quit(); } catch { }
                        try { Marshal.ReleaseComObject(pptApp); } catch { }
                        pptApp = null;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[QuitPowerPointForResetRetry] " + ex.Message);
            }

            var wait = Stopwatch.StartNew();
            while (wait.ElapsedMilliseconds < 5000)
            {
                if (Process.GetProcessesByName("POWERPNT").Length == 0)
                    return;
                Thread.Sleep(200);
            }

            foreach (Process process in Process.GetProcessesByName("POWERPNT"))
            {
                try
                {
                    if (!process.HasExited)
                        process.Kill();
                }
                catch { }
                finally
                {
                    try { process.Dispose(); } catch { }
                }
            }

            PPLogReader.ClearVstoHeartbeat();
        }

        /// <summary>
        /// 一括採点後に試験ファイルを保存して閉じる。
        /// PowerPointが最後のファイルを閉じた直後に作る未保存の白紙は保存せず閉じ、ウィンドウも残さない。
        /// </summary>
        private void ClosePowerPointAfterBatchScoring()
        {
            CloseAllPowerPointPresentations();
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                PowerPointApp pptApp = null;
                try
                {
                    pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                }
                catch
                {
                    return;
                }

                try { pptApp.DisplayAlerts = PpAlertLevel.ppAlertsNone; } catch { }
                for (int pass = 0; pass < 3; pass++)
                {
                    CloseUnsavedBlankPresentations(pptApp);
                    if (GetPresentationCountSafe(pptApp) == 0)
                        break;
                    Thread.Sleep(150);
                }

                try { pptApp.Quit(); }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ClosePowerPointAfterBatchScoring] Quit: " + ex.Message);
                }
                finally
                {
                    try { Marshal.ReleaseComObject(pptApp); } catch { }
                }
            }
        }

        private static void CloseUnsavedBlankPresentations(PowerPointApp pptApp)
        {
            int count = GetPresentationCountSafe(pptApp);
            for (int i = count; i >= 1; i--)
            {
                PowerPointPresentation pres = null;
                try
                {
                    pres = pptApp.Presentations[i];
                    if (!IsUnsavedBlankPresentation(pres))
                        continue;
                    try { pres.Saved = Microsoft.Office.Core.MsoTriState.msoTrue; } catch { }
                    pres.Close();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[CloseUnsavedBlankPresentations] " + ex.Message);
                }
                finally
                {
                    if (pres != null)
                    {
                        try { Marshal.ReleaseComObject(pres); } catch { }
                    }
                }
            }
        }

        private static bool IsUnsavedBlankPresentation(PowerPointPresentation pres)
        {
            string fullName = null;
            try { fullName = pres.FullName; } catch { }
            if (string.IsNullOrWhiteSpace(fullName))
                return true;
            return fullName.IndexOf("\\", StringComparison.Ordinal) < 0
                && fullName.IndexOf("/", StringComparison.Ordinal) < 0;
        }

        private void CloseAllPowerPointPresentations()
        {
            try
            {
                lock (PowerPointCheckerCommon.PowerPointComInteropSync)
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
                
                // すべてのプレゼンテーションを閉じる（無限ループ防止: 最大試行回数と Count が減らない場合の打ち切り）
                // 注意: Close() 後は openPres は無効になるため、名前は必ず閉じる前に取得すること。
                const int maxAttempts = 25;
                int prevCount = GetPresentationCountSafe(pptApp);
                for (int attempt = 0; attempt < maxAttempts && GetPresentationCountSafe(pptApp) > 0; attempt++)
                {
                    PowerPointPresentation openPres = null;
                    try
                    {
                        openPres = pptApp.Presentations[1]; // 1-based index
                        string presName = "(不明)";
                        try { presName = openPres.Name; } catch { }

                        try
                        {
                            openPres.Save();
                            System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] プレゼンテーションを保存しました: {presName}");
                        }
                        catch { }

                        try
                        {
                            openPres.Close();
                        }
                        catch (COMException comClose)
                        {
                            System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] Close で COM: {comClose.Message}");
                        }

                        System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] プレゼンテーションを閉じました: {presName}");

                        int newCount = GetPresentationCountSafe(pptApp);
                        if (newCount >= prevCount)
                        {
                            System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] Countが減らないため、強制的にプロセスを終了します。 (prev={prevCount}, new={newCount})");
                            _resetCloseUsedProcessKill = true;
                            try
                            {
                                var pptProcesses = System.Diagnostics.Process.GetProcessesByName("POWERPNT");
                                foreach (var proc in pptProcesses) 
                                { 
                                    try { proc.Kill(); proc.WaitForExit(2000); } catch { } 
                                }
                            }
                            catch { }
                            break;
                        }
                        prevCount = newCount;
                    }
                    catch (COMException comEx) when (comEx.HResult == unchecked((int)0x80010108)) // RPC_E_DISCONNECTED
                    {
                        System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] プレゼンテーションは既に切断されています（無視）: {comEx.Message}");
                        break;
                    }
                    catch (COMException comEx)
                    {
                        // オブジェクトは存在しません 等、Close 前後の切断
                        System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] COM: {comEx.Message} (0x{comEx.HResult:X8})");
                        if (GetPresentationCountSafe(pptApp) <= 0)
                            break;
                    }
                    catch (Exception closeEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CloseAllPowerPointPresentations] プレゼンテーションを閉じる際のエラー: {closeEx.Message}");
                        if (GetPresentationCountSafe(pptApp) > 0)
                        {
                            try
                            {
                                var nextPres = pptApp.Presentations[1];
                                if (nextPres == openPres)
                                    break;
                                try { Marshal.ReleaseComObject(nextPres); } catch { }
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
            // ウィンドウを閉じる際にタスク情報をクリアし、次回起動時に古い情報でチェックが走るのを防ぐ
            Libraries.PPLogReader.ClearCurrentTaskFile();
            base.OnClosed(e);
        }
        /// <summary>
        /// アプリバー側のプロジェクト変更を MainViewModel 側に同期させます。
        /// これにより、採点時に正しいプロジェクトのタスク一覧が使用されるようになります。
        /// </summary>
        private void SyncProjectToMainViewModel()
        {
            try
            {
                var mainWin = System.Windows.Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                if (mainWin != null && mainWin.DataContext is MainViewModel vm)
                {
                    // 現在の GroupId と ProjectId に一致するプロジェクトを検索
                    var project = vm.ProjectGroups
                        .FirstOrDefault(g => g.GroupId == _groupId)?
                        .Projects.FirstOrDefault(p => p.ProjectId == _currentProjectId);

                    if (project != null)
                    {
                        vm.CurrentProject = project;
                        System.Diagnostics.Debug.WriteLine($"[Sync] MainViewModel のプロジェクトを更新しました: {project.Name}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Sync] 同期エラー: {ex.Message}");
            }
        }

        private static string GetTaskKey(int projectId, int taskId)
        {
            return $"{projectId}-{taskId}";
        }

        private static int GetCurrentTaskAttempt(int projectId, int taskId)
        {
            return Libraries.PPTaskAttemptRegistry.GetAttempt(projectId, taskId);
        }

        private bool EnsureRetryAttemptPrepared(int projectId, int taskId)
        {
            string key = GetTaskKey(projectId, taskId);
            if (!_initialWrongTaskKeys.Contains(key))
                return false;
            if (_preparedRetryTaskKeys.Contains(key))
                return true;

            int currentAttempt = GetCurrentTaskAttempt(projectId, taskId);
            Libraries.PPTaskAttemptRegistry.SetAttempt(projectId, taskId, currentAttempt + 1);
            _retryTaskKeys.Add(key);
            _preparedRetryTaskKeys.Add(key);
            System.Diagnostics.Debug.WriteLine($"[RetryAttempt] Started for {key}, attempt={GetCurrentTaskAttempt(projectId, taskId)}");
            return true;
        }

        private void TryRescorePendingRetryTasks()
        {
            if (_retryTaskKeys.Count == 0)
                return;

            var keysToScore = _retryTaskKeys.ToList();
            var resultWindow = _resultWindow ?? System.Windows.Application.Current.Windows.OfType<ResultWindow>().FirstOrDefault();
            if (resultWindow == null)
                return;

            BeginScoringSession();
            try
            {
                if (!EnsureVstoReady())
                {
                    MessageBox.Show(
                        "PowerPoint 用 VSTO アドインが応答していないため再採点できません。\nアドインを有効にしてから再度実行してください。",
                        "VSTO 未準備",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                foreach (var key in keysToScore)
                {
                    if (!TryParseTaskKey(key, out int projectId, out int taskId))
                        continue;

                    int attemptNo = GetCurrentTaskAttempt(projectId, taskId);
                    bool? passed = null;
                    bool isError = false;
                    try
                    {
                        using (var grader = new PowerPointGrader())
                        {
                            if (!grader.Connect())
                            {
                                isError = true;
                            }
                            else
                            {
                                grader.StartTaskAndWaitForSnapshot(projectId, taskId, attemptNo, 2000, 50);
                                passed = grader.GradeTask(projectId, taskId, attemptNo);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        isError = true;
                        System.Diagnostics.Debug.WriteLine($"[RetryAttempt] Rescore error for {key}: {ex.Message}");
                    }

                    resultWindow.ApplyRetryScoreResult(projectId, taskId, passed, isError);
                    if (!isError)
                        _retryTaskKeys.Remove(key);
                }
            }
            finally
            {
                EndScoringSession(restartTimers: true);
            }
        }

        private static bool TryParseTaskKey(string key, out int projectId, out int taskId)
        {
            projectId = -1;
            taskId = -1;
            if (string.IsNullOrWhiteSpace(key))
                return false;
            var parts = key.Split('-');
            if (parts.Length != 2)
                return false;
            return int.TryParse(parts[0], out projectId) && int.TryParse(parts[1], out taskId);
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


