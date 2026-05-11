using System;
using System.Windows;
using System.Windows.Threading;
using Newtonsoft.Json;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using Ui.ViewModels;
using MOSExcelMogiApp.Views;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Interop;
using System.Text;
using System.Threading;
using Libraries;

namespace MOSExcelMogiApp
{
    public partial class AppBarWindow : Window
    {
        private MainViewModel _viewModel;
        private DispatcherTimer _timer;
        private TimeSpan _remainingTime;
        private int _currentProjectId = 1;
        private int _currentTaskId = 1;
        // 結果画面など外部から「このタスクを表示したい」と指定されたときに使用する。
        // OnCurrentProjectChanged 内のデフォルトリセット（taskId=1）より優先する。
        private int? _pendingTaskId = null;
        private List<TaskInfo> _tasks;
        private ProjectData _projectData;
        private Dictionary<int, bool[]> _projectTaskCompletedStates = new Dictionary<int, bool[]>();
        private Dictionary<int, bool[]> _projectTaskFlaggedStates = new Dictionary<int, bool[]>();
        private bool _isPaused = false;
        private bool _fromResultWindow = false; // 結果画面から来たかどうか
        private ResultWindow _resultWindow = null; // 結果画面への参照
        private DispatcherTimer _ratioRestoreTimer; // Office サイズ変更を検知して初期比率に戻す用
        private bool _isNavigatingToTask = false; // 連続クリックで多重起動しないためのガード

        // Win32 API
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        // 解像度 1920×1080 前提の配置定数
        private const int SCREEN_WIDTH = 1920;
        private const int SCREEN_HEIGHT = 1080;
        private const int APP_BAR_HEIGHT = 258;
        private static readonly int EXCEL_HEIGHT = SCREEN_HEIGHT - APP_BAR_HEIGHT; // 822
        private static readonly int APP_BAR_TOP = EXCEL_HEIGHT; // 822

        public AppBarWindow(MainViewModel viewModel)
        {
            System.Diagnostics.Debug.WriteLine("[AppBarWindow] Constructor called");
            InitializeComponent();
            System.Diagnostics.Debug.WriteLine("[AppBarWindow] InitializeComponent completed");
            
            _viewModel = viewModel;
            DataContext = _viewModel;
            System.Diagnostics.Debug.WriteLine("[AppBarWindow] ViewModel set");

            // ViewModelから現在のプロジェクトIDを取得
            if (_viewModel.CurrentProject != null)
            {
                _currentProjectId = _viewModel.CurrentProject.ProjectNumber;
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Current project ID: {_currentProjectId}");
            }

            InitializeTimer();
            LoadTasks();
            UpdateTaskDisplay();
            WriteCurrentTaskFile();
            
            // 試験終了時にアプリバーを閉じるイベントを購読
            _viewModel.ExamEnded += OnExamEnded;
            
            // プロジェクト変更時に問題文を更新するイベントを購読
            _viewModel.CurrentProjectChanged += OnCurrentProjectChanged;
            
            // レビューページ表示要求イベントを購読
            _viewModel.OpenReviewPageRequested += OnOpenReviewPageRequested;

            // シェル起動後の共有 Excel 接続完了時に Excel ウィンドウを再配置（起動直後のずれを解消）
            _viewModel.SharedExcelApplicationAttached += OnSharedExcelApplicationAttached;

            System.Diagnostics.Debug.WriteLine("[AppBarWindow] Constructor completed");
        }

        private void SetWindowPosition()
        {
            // Excel の共有参照が確立した後のみ配置を行う（初期化中の新規起動・競合を避ける）
            try
            {
                var sharedExcel = _viewModel?.TryGetSharedExcelApplication();
                if (sharedExcel != null)
                    PositionExcelWindow();
            }
            catch
            {
                // ignore
            }

            // ウィンドウハンドルを取得
            IntPtr hWnd = new WindowInteropHelper(this).Handle;
            if (hWnd == IntPtr.Zero)
            {
                // ハンドルが取得できない場合はWPFプロパティで設定（1920×1080前提）
                this.Width = SCREEN_WIDTH;
                this.Height = APP_BAR_HEIGHT;
                this.Left = 0;
                this.Top = APP_BAR_TOP;
                this.Topmost = true;
                return;
            }

            // 現在のウィンドウサイズを取得して境界線のサイズを計算
            GetWindowRect(hWnd, out RECT windowRect);
            GetClientRect(hWnd, out RECT clientRect);

            int borderWidth = (windowRect.right - windowRect.left) - clientRect.right;
            int borderHeight = (windowRect.bottom - windowRect.top) - clientRect.bottom;

            // アプリバーを 1920×1080 前提で Excel の下に配置
            int x = -borderWidth / 2;
            int y = APP_BAR_TOP - borderHeight / 2;
            int width = SCREEN_WIDTH + borderWidth;
            int height = APP_BAR_HEIGHT + borderHeight;

            MoveWindow(hWnd, x, y, width, height, true);
            
            // ウィンドウを最前面に表示
            this.Topmost = true;
        }

        /// <summary>
        /// 四角（□）ボタンクリック: Excel とアプリバーを 1920×1080 前提の定数位置に再配置する。
        /// 遅延実行でウィンドウ操作が確実に適用されるようにする。
        /// </summary>
        private void PositionWindowButton_Click(object sender, RoutedEventArgs e)
        {
            var delayTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(150) };
            delayTimer.Tick += (s, args) =>
            {
                delayTimer.Stop();
                SetWindowPosition();
            };
            delayTimer.Start();
        }

        /// <param name="bringToForeground">true のときのみ Excel を前面に出す。タイマーから呼ぶ場合は false にし、ダイアログ入力中のフォーカスを奪わない。</param>
        private void PositionExcelWindow(bool bringToForeground = true)
        {
            try
            {
                // まずアプリが保持しているExcelインスタンスの Hwnd を使う（ここで新規起動はしない）
                IntPtr excelHwnd = IntPtr.Zero;
                uint processId = 0;
                try
                {
                    var sharedExcel = _viewModel?.TryGetSharedExcelApplication();
                    if (sharedExcel != null)
                    {
                        int hwnd = 0;
                        try { hwnd = sharedExcel.Hwnd; } catch { hwnd = 0; }
                        if (hwnd != 0)
                        {
                            excelHwnd = new IntPtr(hwnd);
                            try { GetWindowThreadProcessId(excelHwnd, out processId); } catch { processId = 0; }
                        }
                    }
                }
                catch { }

                // フォールバック: 実行中のExcelプロセスを取得（複数ある場合は最も直前に起動したプロセスを対象にする）
                var excelProcesses = Process.GetProcessesByName("EXCEL");
                // #region agent log
                try
                {
                    var logPath = @"c:\Users\kouza\source\repos\MOS PowerPoint app\.cursor\debug.log";
                    File.AppendAllText(logPath, JsonConvert.SerializeObject(new { timestamp = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds, location = "AppBarWindow.PositionExcelWindow", message = "entry", data = new { excelCount = excelProcesses?.Length ?? 0 }, sessionId = "debug-session", hypothesisId = "H2" }) + "\n");
                }
                catch { }
                // #endregion
                if (excelHwnd == IntPtr.Zero && excelProcesses.Length == 0) return;

                Process excelProcess = null;
                if (excelHwnd == IntPtr.Zero)
                {
                    excelProcess = excelProcesses
                        .OrderByDescending(p => { try { return p.StartTime; } catch { return DateTime.MinValue; } })
                        .FirstOrDefault();
                    if (excelProcess == null) return;
                    processId = (uint)excelProcess.Id;
                }
                // #region agent log
                try
                {
                    var logPath = @"c:\Users\kouza\source\repos\MOS PowerPoint app\.cursor\debug.log";
                    DateTime st = DateTime.MinValue; try { if (excelProcess != null) st = excelProcess.StartTime; } catch { }
                    File.AppendAllText(logPath, JsonConvert.SerializeObject(new { timestamp = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds, location = "AppBarWindow.PositionExcelWindow", message = "selected process", data = new { processId, startTime = st.ToString("o"), usedSharedHwnd = excelHwnd != IntPtr.Zero }, sessionId = "debug-session", hypothesisId = "H2" }) + "\n");
                }
                catch { }
                // #endregion
                // Excelのメインウィンドウハンドルを取得（リトライロジック）
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
                // #region agent log
                try
                {
                    var logPath = @"c:\Users\kouza\source\repos\MOS PowerPoint app\.cursor\debug.log";
                    File.AppendAllText(logPath, JsonConvert.SerializeObject(new { timestamp = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds, location = "AppBarWindow.PositionExcelWindow", message = "XLMAIN result", data = new { found = excelHwnd != IntPtr.Zero, retryCount }, sessionId = "debug-session", hypothesisId = "H3" }) + "\n");
                }
                catch { }
                // #endregion
                // ウィンドウハンドルが見つかった場合、必要時のみ前面に持ってきてからリサイズ（プロジェクト1-1と同じ高さで統一）
                if (excelHwnd != IntPtr.Zero)
                {
                    if (bringToForeground)
                    {
                        // 前面が Excel のダイアログ（子ウィンドウ）のときは SetForegroundWindow を呼ばない（ダイアログの選択を解除しない）
                        IntPtr fg = GetForegroundWindow();
                        if (fg != IntPtr.Zero)
                        {
                            GetWindowThreadProcessId(fg, out uint fgPid);
                            bool foregroundIsExcelDialog = (fgPid == processId && fg != excelHwnd);
                            if (!foregroundIsExcelDialog)
                                SetForegroundWindow(excelHwnd);
                        }
                        else
                            SetForegroundWindow(excelHwnd);
                    }

                    // Excelのウィンドウの境界線サイズを取得
                    GetWindowRect(excelHwnd, out RECT excelWindowRect);
                    GetClientRect(excelHwnd, out RECT excelClientRect);
                    
                    int excelBorderWidth = (excelWindowRect.right - excelWindowRect.left) - excelClientRect.right;
                    int excelBorderHeight = (excelWindowRect.bottom - excelWindowRect.top) - excelClientRect.bottom;
                    
                    // Excelのウィンドウを 1920×1080 前提で左上 (0,0)、サイズ 1920×822 にリサイズ（1-1と同じ高さ）
                    int excelX = -excelBorderWidth / 2;
                    int excelY = -excelBorderHeight / 2;
                    int excelWidth = SCREEN_WIDTH + excelBorderWidth;
                    int excelHeight = EXCEL_HEIGHT + excelBorderHeight;
                    // #region agent log
                    try
                    {
                        var logPath = @"c:\Users\kouza\source\repos\MOS PowerPoint app\.cursor\debug.log";
                        File.AppendAllText(logPath, JsonConvert.SerializeObject(new { timestamp = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds, location = "AppBarWindow.PositionExcelWindow", message = "MoveWindow", data = new { excelWidth, excelHeight, excelX, excelY }, sessionId = "debug-session", hypothesisId = "H4" }) + "\n");
                    }
                    catch { }
                    // #endregion
                    // 最大化/最小化状態だと MoveWindow が効かず比率が崩れることがあるため、必ず復元してから移動/リサイズする
                    try { ShowWindow(excelHwnd, SW_RESTORE); } catch { }
                    MoveWindow(excelHwnd, excelX, excelY, excelWidth, excelHeight, true);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Error positioning Excel window: {ex.Message}");
            }
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            System.Diagnostics.Debug.WriteLine("[AppBarWindow] OnContentRendered called");
            
            // プロジェクトリセットボタンが存在するか確認
            if (ProjectResetButton != null)
            {
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] ProjectResetButton found. IsVisible: {ProjectResetButton.IsVisible}, IsEnabled: {ProjectResetButton.IsEnabled}");
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] ProjectResetButton Content: {ProjectResetButton.Content}");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[AppBarWindow] ProjectResetButton is NULL!");
            }
            
            // ウィンドウハンドルが利用可能になるまで少し待機してから配置
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SetWindowPosition();
                StartRatioRestoreTimer();
            }), DispatcherPriority.Loaded);
        }

        /// <summary>Office のウィンドウサイズ変更を監視し、変更されていたら初期の画面比率に戻すタイマーを開始する。</summary>
        private void StartRatioRestoreTimer()
        {
            if (_ratioRestoreTimer != null) return;
            _ratioRestoreTimer = new DispatcherTimer();
            _ratioRestoreTimer.Interval = TimeSpan.FromSeconds(2);
            _ratioRestoreTimer.Tick += RatioRestoreTimer_Tick;
            _ratioRestoreTimer.Start();
        }

        private void RatioRestoreTimer_Tick(object sender, EventArgs e)
        {
            // 計画: タイマーからは配置処理を一切呼ばない。ダイアログのフォーカスを奪わないため。
            // ユーザーは四角ボタンで明示的に再配置できる。
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
            bool timerDisabled = MainWindow.IsTimerDisabled;
            
            System.Diagnostics.Debug.WriteLine($"AppBarWindow: InitializeTimer called, IsTimerDisabled = {timerDisabled}");
            
            if (timerDisabled)
            {
                // タイマーは開始しない
                _timer.Stop();
                System.Diagnostics.Debug.WriteLine("AppBarWindow: Timer disabled, not starting");
                
                // 「一時停止」ボタンをグレーアウト（無効化）する
                UpdatePauseButtonState(true);
            }
            else
            {
                _timer.Start();
                System.Diagnostics.Debug.WriteLine("AppBarWindow: Timer enabled, starting");
                
                // 「一時停止」ボタンを有効化する
                UpdatePauseButtonState(false);
            }
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
                // 試験終了処理
                _viewModel.EndExamCommand.Execute(null);
            }
        }

        private void UpdateTimerDisplay()
        {
            // タイマーの表示を更新
            var timerTextBlock = FindName("TimerTextBlock") as TextBlock;
            if (timerTextBlock != null)
            {
                timerTextBlock.Text = _remainingTime.ToString(@"hh\:mm\:ss");
            }
        }

        private void LoadTasks()
        {
            try
            {
                // ★ Group番号から適切なファイルを選択
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
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"タスク読み込みエラー: {ex.Message}");
                _tasks = new List<TaskInfo>();
            }
        }

        private void LoadTasksFromCsv(int groupId)
        {
            try
            {
                // CSVファイルのパス
                string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "CSV", "解答手順あり模擬試験①問題文.csv");
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Loading tasks from CSV: {csvPath} (GroupId: {groupId})");
                
                if (!File.Exists(csvPath))
                {
                    System.Diagnostics.Debug.WriteLine($"CSVファイルが見つかりません: {csvPath}");
                    _tasks = new List<TaskInfo>();
                    return;
                }
                
                // CSVファイルを読み込む
                var projects = new Dictionary<int, List<TaskInfo>>();
                string[] lines = File.ReadAllLines(csvPath, Encoding.UTF8);
                
                System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv] CSVファイルの行数: {lines.Length}");
                
                // ヘッダー行をスキップ（1行目）
                for (int i = 1; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (string.IsNullOrEmpty(line))
                        continue;
                    
                    // CSVのパース（カンマ区切り、ただし引用符内のカンマは考慮）
                    string[] fields = ParseCsvLine(line);
                    
                    System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv] 行{i}: フィールド数={fields.Length}");
                    if (fields.Length >= 3)
                    {
                        System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv]   グループ={fields[0]}, プロジェクト={fields[1]}, 問題文={fields[2].Substring(0, Math.Min(50, fields[2].Length))}...");
                    }
                    
                    if (fields.Length < 3)
                    {
                        System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv]   フィールド数が不足しています");
                        continue;
                    }
                    
                    // グループ,プロジェクト,問題文,解答操作
                    if (int.TryParse(fields[0].Trim(), out int csvGroupId) && 
                        int.TryParse(fields[1].Trim(), out int projectId) &&
                        csvGroupId == groupId)
                    {
                        string description = fields[2].Trim();
                        if (string.IsNullOrEmpty(description))
                        {
                            System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv]   問題文が空です");
                            continue;
                        }
                        
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
                        
                        System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv]   プロジェクト{projectId}にタスク{taskId}を追加: {description.Substring(0, Math.Min(50, description.Length))}...");
                    }
                }
                
                // デバッグ出力：各プロジェクトのタスク数を確認
                foreach (var kvp in projects)
                {
                    System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv] プロジェクト{kvp.Key}: {kvp.Value.Count}個のタスク");
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
                System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv] 現在のプロジェクトID: {_currentProjectId}");
                System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv] 読み込まれたプロジェクト数: {_projectData.Projects.Count}");
                foreach (var p in _projectData.Projects)
                {
                    System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv]   プロジェクト{p.ProjectId}: {p.Tasks?.Count ?? 0}個のタスク");
                }
                
                var currentProject = _projectData.Projects.Find(p => p.ProjectId == _currentProjectId);
                if (currentProject != null)
                {
                    _tasks = currentProject.Tasks;
                    
                    System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv] プロジェクト {_currentProjectId} のタスク数: {_tasks.Count}");
                    for (int i = 0; i < _tasks.Count; i++)
                    {
                        System.Diagnostics.Debug.WriteLine($"[LoadTasksFromCsv]   タスク{i + 1}: {_tasks[i].Description.Substring(0, Math.Min(50, _tasks[i].Description.Length))}...");
                    }
                    
                    // プロジェクトの状態を初期化（まだ存在しない場合のみ）
                    // TaskIdは1始まりだが、配列インデックスは0始まりなので maxTaskCount のみ
                    int maxTaskCount = Math.Max(_tasks.Count, 8);
                    if (!_projectTaskCompletedStates.ContainsKey(_currentProjectId))
                    {
                        _projectTaskCompletedStates[_currentProjectId] = new bool[maxTaskCount];
                    }
                    if (!_projectTaskFlaggedStates.ContainsKey(_currentProjectId))
                    {
                        _projectTaskFlaggedStates[_currentProjectId] = new bool[maxTaskCount];
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"プロジェクト {_currentProjectId} のタスクをCSVから読み込みました（タスク数: {_tasks.Count}）");
                }
                else
                {
                    _tasks = new List<TaskInfo>();
                    System.Diagnostics.Debug.WriteLine($"プロジェクト {_currentProjectId} が見つかりませんでした");
                    System.Diagnostics.Debug.WriteLine($"利用可能なプロジェクト: {string.Join(", ", _projectData.Projects.Select(p => p.ProjectId))}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CSV読み込みエラー: {ex.Message}\n{ex.StackTrace}");
                _tasks = new List<TaskInfo>();
            }
        }

        private void LoadTasksFromJson(int groupId)
        {
            try
            {
                string jsonFileName = groupId switch
                {
                    1 => "MOS演習問題文一覧.json",        // GroupId=1 → 演習タブ
                    2 => "MOS模擬試験①問題文一覧.json",  // GroupId=2 → 模試①タブ（通常はCSVから読み込むが、フォールバック用）
                    3 => "MOS模擬試験②問題文一覧.json",  // GroupId=3 → 模試②タブ
                    _ => "MOS模擬アプリ問題文一覧.json"
                };
                
                string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", jsonFileName);
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Loading tasks from: {jsonFileName} (GroupId: {groupId})");
                
                string jsonContent = File.ReadAllText(jsonPath);
                
                // 毎回デシリアライズ（キャッシュされていても最新の状態で読み込む）
                _projectData = JsonConvert.DeserializeObject<ProjectData>(jsonContent);

                // 現在のプロジェクトのタスクを取得
                var currentProject = _projectData.Projects.Find(p => p.ProjectId == _currentProjectId);
                if (currentProject != null)
                {
                    _tasks = currentProject.Tasks;
                    
                    // プロジェクトの状態を初期化（まだ存在しない場合のみ）
                    // TaskIdは1始まりだが、配列インデックスは0始まりなので maxTaskCount のみ
                    int maxTaskCount = Math.Max(_tasks.Count, 8);
                    if (!_projectTaskCompletedStates.ContainsKey(_currentProjectId))
                    {
                        _projectTaskCompletedStates[_currentProjectId] = new bool[maxTaskCount];
                    }
                    if (!_projectTaskFlaggedStates.ContainsKey(_currentProjectId))
                    {
                        _projectTaskFlaggedStates[_currentProjectId] = new bool[maxTaskCount];
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"プロジェクト {_currentProjectId} のタスクを読み込みました（タスク数: {_tasks.Count}）");
                }
                else
                {
                    _tasks = new List<TaskInfo>();
                    System.Diagnostics.Debug.WriteLine($"プロジェクト {_currentProjectId} が見つかりませんでした");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"JSON読み込みエラー: {ex.Message}\n{ex.StackTrace}");
                _tasks = new List<TaskInfo>();
            }
        }

        /// <summary>
        /// CSV行をパースします（カンマ区切り、引用符内のカンマを考慮）
        /// </summary>
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

        private void UpdateTaskDisplay()
        {
            // 現在のプロジェクト番号/総プロジェクト数を表示
            int totalProjects = _projectData?.Projects?.Count ?? 0;
            var projectInfoTextBlock = FindName("ProjectInfoTextBlock") as TextBlock;
            if (projectInfoTextBlock != null)
                projectInfoTextBlock.Text = $"{_currentProjectId}/{totalProjects}";

            // タスク説明の表示を更新
            var taskDescriptionTextBlock = FindName("TaskDescriptionTextBlock") as TextBlock;
            if (taskDescriptionTextBlock != null && _tasks != null && _currentTaskId <= _tasks.Count)
            {
                var currentTask = _tasks.Find(t => t.TaskId == _currentTaskId);
                if (currentTask != null)
                {
                    // デバッグ出力
                    System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] Project {_currentProjectId}, Task {_currentTaskId}");
                    System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] Description: {currentTask.Description}");
                    System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] Contains '\"': {currentTask.Description.Contains("\"")}");
                    
                    // "で囲まれた部分に下線を付けて表示
                    SetTextWithUnderline(taskDescriptionTextBlock, currentTask.Description);
                }
            }

            // タスクボタンの状態を更新
            UpdateTaskButtons();

            // ボタンのテキストを更新
            UpdateButtonTexts();

            // VSTO 連携用に現在タスクを共有ファイルへ書き出す
            WriteCurrentTaskFile();
        }

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

            System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] Input text: {text}");
            System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] Text length: {text.Length}");
            
            // 各文字を確認（最初の200文字まで）
            for (int i = 0; i < Math.Min(text.Length, 200); i++)
            {
                char c = text[i];
                if (c == '"' || c == '\u201C' || c == '\u201D') // 半角ダブルクォート、全角左引用符、全角右引用符
                {
                    System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] Found quote character '{c}' (U+{(int)c:X4}) at position {i}");
                }
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
                System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] Found opening \" at position {quoteStart}");
                
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
                
                System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] Found closing \" at position {quoteEnd}");
                
                // "で囲まれた部分のテキスト（"を除く）に下線を付けて追加
                string quotedText = text.Substring(quoteStart + 1, quoteEnd - quoteStart - 1);
                System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] Quoted text: '{quotedText}'");
                
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
                System.Diagnostics.Debug.WriteLine("[SetTextWithUnderline] No quotes found, setting as plain text");
                textBlock.Inlines.Clear();
                textBlock.Text = text;
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] Inlines count: {textBlock.Inlines.Count}");
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
                
                // ユーザーにフィードバックを提供（オプション）
                // ツールチップやトーストメッセージを表示することもできます
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OnUnderlinedTextClick] Error copying to clipboard: {ex.Message}");
                MessageBox.Show($"クリップボードへのコピーに失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateTaskButtons()
        {
            // 現在のプロジェクトの状態を取得
            bool[] completedStates = _projectTaskCompletedStates.ContainsKey(_currentProjectId) ?
                _projectTaskCompletedStates[_currentProjectId] : new bool[8];
            bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(_currentProjectId) ?
                _projectTaskFlaggedStates[_currentProjectId] : new bool[8];

            // すべてのタスクボタンの状態を更新
            for (int i = 1; i <= 8; i++)
            {
                var button = FindName($"TaskButton{i}") as Button;
                if (button != null)
                {
                    // 現在のプロジェクトのタスク数に応じて表示/非表示を制御
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
                            checkTextBlock.Visibility = (arrayIndex >= 0 && arrayIndex < completedStates.Length && completedStates[arrayIndex]) ?
                                Visibility.Visible : Visibility.Collapsed;
                        }

                        // 旗マークの表示制御
                        var flagTextBlock = FindName($"Flag{i}") as TextBlock;
                        if (flagTextBlock != null)
                        {
                            flagTextBlock.Visibility = (arrayIndex >= 0 && arrayIndex < flaggedStates.Length && flaggedStates[arrayIndex]) ?
                                Visibility.Visible : Visibility.Collapsed;
                        }

                        if (i == _currentTaskId)
                        {
                            // 現在のタスクボタンは選択状態
                            button.Background = System.Windows.Media.Brushes.White;
                            button.BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Gray);
                            button.BorderThickness = new Thickness(2);
                        }
                        else
                        {
                            // 他のタスクボタンは非選択状態
                            button.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightGray);
                            button.BorderThickness = new Thickness(0);
                        }
                    }
                }
            }
        }

        private void UpdateButtonTexts()
        {
            // 現在のプロジェクトの状態を取得
            bool[] completedStates = _projectTaskCompletedStates.ContainsKey(_currentProjectId) ?
                _projectTaskCompletedStates[_currentProjectId] : new bool[8];
            bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(_currentProjectId) ?
                _projectTaskFlaggedStates[_currentProjectId] : new bool[8];

            // 解答済みボタンのテキストを更新
            var completeButton = FindName("CompleteButton") as Button;
            if (completeButton != null)
            {
                // TaskIdは1始まり、配列は0始まりなので -1
                if (_currentTaskId > 0 && _currentTaskId <= completedStates.Length && completedStates[_currentTaskId - 1])
                {
                    completeButton.Content = "✓ 解答済み";
                }
                else
                {
                    completeButton.Content = "解答済みにする";
                }
            }

            // フラグボタンのテキストを更新
            var flagButton = FindName("FlagButton") as Button;
            if (flagButton != null)
            {
                // TaskIdは1始まり、配列は0始まりなので -1
                if (_currentTaskId > 0 && _currentTaskId <= flaggedStates.Length && flaggedStates[_currentTaskId - 1])
                {
                    flagButton.Content = "フラグを外す";
                }
                else
                {
                    flagButton.Content = "あとで見直す";
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
            if (sender is Button button && button.Tag != null)
            {
                int taskId = int.Parse(button.Tag.ToString());
                if (taskId >= 1 && taskId <= _tasks.Count)
                {
                    _currentTaskId = taskId;
                    UpdateTaskDisplay();
                }
            }
        }

        private void PauseButton_Click(object sender, RoutedEventArgs e)
        {
            // タイマーが無効化されている場合は何もしない
            if (MainWindow.IsTimerDisabled)
            {
                return;
            }
            
            if (_isPaused)
            {
                // タイマーを再開
                _timer?.Start();
                _isPaused = false;
                
                // ボタンのテキストを「一時停止」に変更
                if (sender is Button button)
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
                if (sender is Button button)
                {
                    button.Content = "再開";
                }
            }
        }
        
        private void UpdatePauseButtonState(bool isDisabled)
        {
            var pauseButton = FindName("PauseButton") as Button;
            if (pauseButton != null)
            {
                if (isDisabled)
                {
                    // ボタンを無効化（グレーアウト）
                    pauseButton.IsEnabled = false;
                    pauseButton.Opacity = 0.5;
                    System.Diagnostics.Debug.WriteLine("AppBarWindow: PauseButton disabled (grayed out)");
                }
                else
                {
                    // ボタンを有効化
                    pauseButton.IsEnabled = true;
                    pauseButton.Opacity = 1.0;
                    System.Diagnostics.Debug.WriteLine("AppBarWindow: PauseButton enabled");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("AppBarWindow: PauseButton is null!");
            }
        }

        private void ProjectResetButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("========================================");
            System.Diagnostics.Debug.WriteLine("[AppBarWindow] ProjectResetButton_Click called");
            System.Diagnostics.Debug.WriteLine($"Sender: {sender?.GetType().Name}");
            System.Diagnostics.Debug.WriteLine($"EventArgs: {e?.GetType().Name}");
            System.Diagnostics.Debug.WriteLine("========================================");
            try
            {
                if (_viewModel?.CurrentProject != null)
                {
                    var currentProject = _viewModel.CurrentProject;
                    int groupId = int.Parse(currentProject.Group.Replace("Group ", ""));
                    int projectId = currentProject.ProjectNumber;
                    
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Current project: Group{groupId}, Project{projectId}");
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Current project file path: {currentProject.FilePath}");
                    
                    var result = MessageBox.Show(
                        $"プロジェクト {groupId}-{projectId} をリセットしますか？\n（編集内容は失われます）", 
                        "確認", 
                        MessageBoxButton.YesNo, 
                        MessageBoxImage.Question);
                    
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] User response: {result}");
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Starting reset process...");
                        
                        // MainWindowのResetProjectメソッドを呼び出す
                        MainWindow mainWindow = Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                        if (mainWindow != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[AppBarWindow] MainWindow found, calling ResetProject");
                            mainWindow.ResetProject(groupId, projectId);
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[AppBarWindow] MainWindow not found");
                            MessageBox.Show("メインウィンドウが見つかりませんでした。", "エラー", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Reset cancelled by user");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[AppBarWindow] CurrentProject is null");
                    MessageBox.Show("リセットするプロジェクトが選択されていません。", 
                        "情報", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Error in ProjectResetButton_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"プロジェクトリセット中にエラーが発生しました: {ex.Message}", 
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void ReviewPageButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // レビューページを開く前に現在のExcelプロジェクトを自動保存
                // closeWorkbook: false にして、ワークブックは開いたままにする
                if (_viewModel != null)
                {
                    System.Diagnostics.Debug.WriteLine("[AppBarWindow] Saving and closing current project before opening review page");
                    _viewModel.SaveCurrentExcelProject(closeWorkbook: true);
                    _viewModel.CloseExcelApplication();
                }

                // メインのバーウィンドウを非表示にする
                this.Hide();
                
                // Group番号を取得
                int groupId = 1; // デフォルト
                if (_viewModel?.CurrentProject != null)
                {
                    string groupStr = _viewModel.CurrentProject.Group.Replace("Group ", "");
                    int.TryParse(groupStr, out groupId);
                }
                
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Opening ReviewPage with GroupId: {groupId}");
                
                // 現在のタイマー残り時間と状態情報を渡す（groupIdも追加）
                var reviewWindow = new ReviewPageWindow(_remainingTime, _projectTaskCompletedStates, _projectTaskFlaggedStates, groupId);
                reviewWindow.OnNavigateToTask = (g, p, t) => NavigateToTask(p, t, g);
                reviewWindow.Closed += (s, args) => 
                {
                    // レビューページが閉じられたらメインウィンドウを再表示
                    this.Show();
                };
                
                // ShowDialog()ではなくShow()を使用
                reviewWindow.Show();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"レビューページ表示エラー: {ex.Message}");
                MessageBox.Show("レビューページの表示に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                // エラーが発生した場合はメインウィンドウを再表示
                this.Show();
            }
        }

        /// <param name="groupIdOverride">結果画面・レビューから遷移するときのグループ。null のときは <see cref="Ui.ViewModels.MainViewModel.CurrentProject"/> から推定。</param>
        public void NavigateToTask(int projectId, int taskId, int? groupIdOverride = null)
        {
            System.Diagnostics.Debug.WriteLine($"[AppBarWindow] NavigateToTask called: ProjectId={projectId}, TaskId={taskId}");

            if (_isNavigatingToTask)
            {
                System.Diagnostics.Debug.WriteLine("[AppBarWindow] NavigateToTask ignored (already navigating)");
                return;
            }
            _isNavigatingToTask = true;
            _pendingTaskId = taskId;

            if (_viewModel != null && !_viewModel.WaitForExcelShutdownToCompleteBeforeOpeningProject())
            {
                _isNavigatingToTask = false;
                _pendingTaskId = null;
                return;
            }
            
            // AppBarWindowを確実に表示（既に表示されている場合は何もしない）
            try
            {
                // ウィンドウが閉じられている場合は何もしない
                if (!this.IsLoaded)
                {
                    System.Diagnostics.Debug.WriteLine("[AppBarWindow] Window is not loaded, skipping Show/Activate");
                }
                else
                {
                    bool wasHidden = !this.IsVisible;
                    if (wasHidden)
                    {
                        this.Show();
                    }
                    // 結果画面から戻る際に常に Activate/Focus すると、Excel の関数入力・ダイアログのフォーカスを奪うため、
                    // ウィンドウを表示した場合のみ前面に出す（既に表示中ならフォーカスは移さない）
                    if (wasHidden)
                    {
                        this.Activate();
                        this.Focus();
                    }
                }
            }
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Error showing window (may be closing): {ex.Message}");
                _isNavigatingToTask = false;
                _pendingTaskId = null;
                return; // NavigateToTaskを中断
            }
            
            ExcelApp excelApp = null;
            ExcelWorkbook targetWorkbook = null;
            
            try
            {
                // プロジェクトIDとタスクIDを設定
                _currentProjectId = projectId;
                _currentTaskId = taskId;
                
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Set current project: {_currentProjectId}, current task: {_currentTaskId}");
                
                // Excelファイルを開く処理を追加
                int groupId = 1;
                if (groupIdOverride.HasValue && groupIdOverride.Value > 0)
                    groupId = groupIdOverride.Value;
                else if (_viewModel?.CurrentProject != null)
                {
                    string groupStr = _viewModel.CurrentProject.Group.Replace("Group ", "");
                    int.TryParse(groupStr, out groupId);
                }

                // プロジェクトファイルパスを取得
                string filePath = _viewModel.GetProjectFilePath(groupId, projectId);
                
                if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                {
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Opening Excel file: {filePath}");

                    // Excel を COM で取得（起動中ならそれを使う／無ければ新規起動→最後にシェル起動＋ROT 接続）
                    try
                    {
                        excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (COMException)
                    {
                        if (_viewModel == null)
                            throw;

                        // 先に対象ブックをシェルで開き短時間で ROT 接続（スタート画面の空 Excel 起動より優先）
                        excelApp = _viewModel.TryOpenWorkbookByShellAndAttachRunningExcel(
                            filePath,
                            delayMs: 600,
                            attachTimeoutMs: 8000);

                        if (excelApp == null)
                        {
                            try
                            {
                                excelApp = ExcelApplicationManager.GetOrCreateExcelApplication(makeVisible: true, timeoutMs: 15000);
                            }
                            catch (Exception ex2)
                            {
                                excelApp = _viewModel.TryOpenWorkbookByShellAndAttachRunningExcel(
                                    filePath,
                                    delayMs: 1000,
                                    attachTimeoutMs: 12000);
                                if (excelApp == null)
                                    throw new InvalidOperationException(
                                        "Excel に接続できませんでした。しばらくしてから再度お試しください。",
                                        ex2);
                            }
                        }
                    }

                    // 既に開いているなら Activate、無ければ Open（Process.Start による多重起動を避ける）
                    string targetFullPathLower = Path.GetFullPath(filePath).ToLowerInvariant();
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            string wbFullLower = (wb.FullName != null ? Path.GetFullPath(wb.FullName) : wb.Name).ToLowerInvariant();
                            if (wbFullLower == targetFullPathLower)
                            {
                                targetWorkbook = wb;
                                break;
                            }
                        }
                        catch
                        {
                            // ignore and continue searching
                        }
                    }

                    if (targetWorkbook == null)
                    {
                        targetWorkbook = excelApp.Workbooks.Open(filePath, ReadOnly: false);
                    }

                    try { targetWorkbook.Activate(); } catch { }

                    // 表示は必須（復習時に見えるように）
                    try { excelApp.Visible = true; } catch { }
                    
                    // プロジェクト情報を更新（イベントハンドラーを一時的に解除してタスクIDのリセットを防ぐ）
                    if (_viewModel != null)
                    {
                        // イベントハンドラーを一時的に解除
                        _viewModel.CurrentProjectChanged -= OnCurrentProjectChanged;
                        
                        _viewModel.CurrentProject = new Ui.ViewModels.ProjectInfo
                        {
                            Name = $"プロジェクト {projectId}",
                            FilePath = filePath,
                            Group = $"Group {groupId}",
                            ProjectNumber = projectId
                        };
                        
                        // イベントハンドラーを再登録
                        _viewModel.CurrentProjectChanged += OnCurrentProjectChanged;

                        // 共有参照へ載せ替え（ここで Release しない。ViewModel が保持し SharedExcelApplicationAttached で再配置）
                        _viewModel.PublishSharedExcelApplication(excelApp);
                        excelApp = null;
                    }

                    // ハンドル確定の遅れ対策: プロジェクト切替と同程度の遅延でもう一度配置
                    var delayTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(1200) };
                    delayTimer.Tick += (s, e) =>
                    {
                        delayTimer.Stop();
                        try { SetWindowPosition(); } catch { }
                    };
                    delayTimer.Start();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Excel file not found: {filePath}");
                    MessageBox.Show($"Excelファイルが見つかりませんでした: {filePath}", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                
                // タスクを再読み込み
                LoadTasks();
                
                // タスクIDが実在するかで確認（TaskId とインデックスがズレるケースがあるため Count では判定しない）
                bool taskExists = _tasks != null && _tasks.Any(t => t.TaskId == _currentTaskId);
                if (!taskExists && _pendingTaskId.HasValue)
                {
                    // 結果画面など外部指定のタスクがあるなら再適用（LoadTasks/CurrentProject変更で上書きされることがある）
                    int desired = _pendingTaskId.Value;
                    if (_tasks != null && _tasks.Any(t => t.TaskId == desired))
                    {
                        _currentTaskId = desired;
                        taskExists = true;
                        System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Re-applied pending TaskId={desired} after LoadTasks");
                    }
                    _pendingTaskId = null;
                }

                if (!taskExists)
                {
                    int firstTaskId = _tasks != null && _tasks.Count > 0 ? _tasks[0].TaskId : 1;
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Warning: TaskId {_currentTaskId} not found (tasks count: {_tasks?.Count ?? 0}). Fallback to TaskId={firstTaskId}");
                    _currentTaskId = firstTaskId;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] TaskId {_currentTaskId} exists, updating display");
                }
                
                // タスク表示を更新（問題文とボタンの選択状態を含む）
                UpdateTaskDisplay();
                
                System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Task display updated: Project={_currentProjectId}, Task={_currentTaskId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"タスクナビゲーションエラー: {ex.Message}");
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // リソースのクリーンアップ
                _isNavigatingToTask = false;
                if (targetWorkbook != null)
                {
                    // Workbooks コレクションの所有は excelApp 側なので Release はしない（参照だけ外す）
                    targetWorkbook = null;
                }
                if (excelApp != null)
                {
                    try { Marshal.ReleaseComObject(excelApp); } catch { }
                }
            }
        }

        private void FlagButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"[FlagButton_Click] Current TaskId={_currentTaskId}, ProjectId={_currentProjectId}");
            
            // 現在のプロジェクトの状態を取得または初期化
            if (!_projectTaskFlaggedStates.ContainsKey(_currentProjectId))
            {
                _projectTaskFlaggedStates[_currentProjectId] = new bool[8];
            }

            bool[] flaggedStates = _projectTaskFlaggedStates[_currentProjectId];

            // TaskIdは1始まり、配列は0始まりなので -1
            int index = _currentTaskId - 1;
            System.Diagnostics.Debug.WriteLine($"[FlagButton_Click] Calculated array index={index} (TaskId {_currentTaskId} - 1)");
            
            if (index >= 0 && index < flaggedStates.Length)
            {
                if (!flaggedStates[index])
                {
                    // フラグを設定
                    flaggedStates[index] = true;
                    System.Diagnostics.Debug.WriteLine($"[FlagButton_Click] Set flaggedStates[{index}] = true for TaskId {_currentTaskId}");
                }
                else
                {
                    // フラグを解除
                    flaggedStates[index] = false;
                    System.Diagnostics.Debug.WriteLine($"[FlagButton_Click] Set flaggedStates[{index}] = false for TaskId {_currentTaskId}");
                }
                
                // デバッグ: 現在の配列状態を表示
                for (int i = 0; i < flaggedStates.Length; i++)
                {
                    if (flaggedStates[i])
                    {
                        System.Diagnostics.Debug.WriteLine($"[FlagButton_Click] Array state: flaggedStates[{i}] = true (corresponds to TaskId {i + 1})");
                    }
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[FlagButton_Click] ERROR: Index {index} is out of range (0-{flaggedStates.Length - 1})");
            }

            // UIを更新
            UpdateTaskDisplay();
        }

        private void CompleteButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"[CompleteButton_Click] Current TaskId={_currentTaskId}, ProjectId={_currentProjectId}");
            
            // 現在のプロジェクトの状態を取得または初期化
            if (!_projectTaskCompletedStates.ContainsKey(_currentProjectId))
            {
                _projectTaskCompletedStates[_currentProjectId] = new bool[8];
            }

            bool[] completedStates = _projectTaskCompletedStates[_currentProjectId];

            // TaskIdは1始まり、配列は0始まりなので -1
            int index = _currentTaskId - 1;
            System.Diagnostics.Debug.WriteLine($"[CompleteButton_Click] Calculated array index={index} (TaskId {_currentTaskId} - 1)");
            
            if (index >= 0 && index < completedStates.Length)
            {
                if (!completedStates[index])
                {
                    // チェックマークを設定
                    completedStates[index] = true;
                    System.Diagnostics.Debug.WriteLine($"[CompleteButton_Click] Set completedStates[{index}] = true for TaskId {_currentTaskId}");
                }
                else
                {
                    // チェックマークを解除
                    completedStates[index] = false;
                    System.Diagnostics.Debug.WriteLine($"[CompleteButton_Click] Set completedStates[{index}] = false for TaskId {_currentTaskId}");
                }
                
                // デバッグ: 現在の配列状態を表示
                for (int i = 0; i < completedStates.Length; i++)
                {
                    if (completedStates[i])
                    {
                        System.Diagnostics.Debug.WriteLine($"[CompleteButton_Click] Array state: completedStates[{i}] = true (corresponds to TaskId {i + 1})");
                    }
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[CompleteButton_Click] ERROR: Index {index} is out of range (0-{completedStates.Length - 1})");
            }

            // UIを更新
            UpdateTaskDisplay();
        }

        private void EndButton_Click(object sender, RoutedEventArgs e)
        {
            // モーダル確認中も DispatcherTimer は進むため、先に止めないと Timer_Tick から試験終了が走り Excel が先に閉じることがある
            bool timerWasEnabled = _timer != null && _timer.IsEnabled;
            if (timerWasEnabled)
                _timer.Stop();

            var result = MessageBox.Show("アプリ自体を終了します。本当にいいですか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                _viewModel.EndExamCommand.Execute(null);
                return;
            }

            if (timerWasEnabled && !MainWindow.IsTimerDisabled)
                _timer?.Start();
        }

        private void OnExamEnded(object sender, EventArgs e)
        {
            // アプリバーを閉じる
            this.Close();
        }

        private void OnCurrentProjectChanged(object sender, EventArgs e)
        {
            // ViewModelのCurrentProjectが変更されたら、問題文を更新
            if (_viewModel.CurrentProject != null)
            {
                int newProjectId = _viewModel.CurrentProject.ProjectNumber;
                // #region agent log
                try
                {
                    var logPath = @"c:\Users\kouza\source\repos\MOS PowerPoint app\.cursor\debug.log";
                    File.AppendAllText(logPath, JsonConvert.SerializeObject(new { timestamp = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds, location = "AppBarWindow.OnCurrentProjectChanged", message = "entry", data = new { newProjectId, _currentProjectId }, sessionId = "debug-session", hypothesisId = "H1" }) + "\n");
                }
                catch { }
                // #endregion
                System.Diagnostics.Debug.WriteLine($"[OnCurrentProjectChanged] Project changed: {_currentProjectId} -> {newProjectId}");
                
                // プロジェクトが変更された場合のみ、タスクIDをリセット
                if (newProjectId != _currentProjectId)
                {
                    _currentProjectId = newProjectId;
                    System.Diagnostics.Debug.WriteLine($"[OnCurrentProjectChanged] Loading tasks for project {_currentProjectId}");
                    LoadTasks();
                    // 結果画面などからの遷移でタスク指定がある場合はそれを優先し、無ければ 1 にリセット
                    if (_pendingTaskId.HasValue)
                    {
                        _currentTaskId = _pendingTaskId.Value;
                        _pendingTaskId = null;
                    }
                    else
                    {
                        _currentTaskId = 1; // 最初のタスクにリセット
                    }
                    UpdateTaskDisplay();

                    // 次のプロジェクトに移動した際も Excel とアプリバーをプロジェクト1-1と同じ高さ・位置（1920×1080前提）に再配置
                    // 新しい Excel ウィンドウが完全に表示されるまで遅延してから実行（シートタブが見えるように）
                    // #region agent log
                    try
                    {
                        var logPath = @"c:\Users\kouza\source\repos\MOS PowerPoint app\.cursor\debug.log";
                        File.AppendAllText(logPath, JsonConvert.SerializeObject(new { timestamp = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds, location = "AppBarWindow.OnCurrentProjectChanged", message = "delay 1200ms started", data = new { newProjectId }, sessionId = "debug-session", hypothesisId = "H1" }) + "\n");
                    }
                    catch { }
                    // #endregion
                    var delayTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(1200) };
                    delayTimer.Tick += (s, args) =>
                    {
                        delayTimer.Stop();
                        // #region agent log
                        try
                        {
                            var logPath = @"c:\Users\kouza\source\repos\MOS PowerPoint app\.cursor\debug.log";
                            File.AppendAllText(logPath, JsonConvert.SerializeObject(new { timestamp = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds, location = "AppBarWindow.OnCurrentProjectChanged", message = "SetWindowPosition from delay", sessionId = "debug-session", hypothesisId = "H1" }) + "\n");
                        }
                        catch { }
                        // #endregion
                        SetWindowPosition();
                    };
                    delayTimer.Start();
                }
                else
                {
                    // 同じプロジェクトの場合は、タスクIDをリセットしない
                    // タスクが読み込まれていない場合は読み込む
                    if (_tasks == null || _tasks.Count == 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OnCurrentProjectChanged] Tasks not loaded, loading now");
                        LoadTasks();
                    }
                    // 同一プロジェクト内のタスク指定がある場合はそれを優先
                    if (_pendingTaskId.HasValue)
                    {
                        _currentTaskId = _pendingTaskId.Value;
                        _pendingTaskId = null;
                    }
                    UpdateTaskDisplay();
                }
            }
        }

        private void OnOpenReviewPageRequested(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[AppBarWindow] OnOpenReviewPageRequested called");
            // レビューページボタンと同じ処理を実行
            ReviewPageButton_Click(sender, new RoutedEventArgs());
        }

        private void OnSharedExcelApplicationAttached(object sender, EventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    SetWindowPosition();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] OnSharedExcelApplicationAttached: {ex.Message}");
                }
            }), DispatcherPriority.Background);
        }

        protected override void OnClosed(EventArgs e)
        {
            ClearCurrentTaskFile();
            _ratioRestoreTimer?.Stop();
            _ratioRestoreTimer = null;
            _timer?.Stop();
            if (_viewModel != null)
            {
                _viewModel.ExamEnded -= OnExamEnded;
                _viewModel.CurrentProjectChanged -= OnCurrentProjectChanged;
                _viewModel.OpenReviewPageRequested -= OnOpenReviewPageRequested;
                _viewModel.SharedExcelApplicationAttached -= OnSharedExcelApplicationAttached;
            }
            base.OnClosed(e);
        }

        private void WriteCurrentTaskFile()
        {
            try
            {
                if (_currentProjectId <= 0 || _currentTaskId <= 0) return;
                string content = $"{_currentProjectId},{_currentTaskId},1";
                File.WriteAllText(ExcelLogReader.GetCurrentTaskFilePath(), content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[AppBarWindow] WriteCurrentTaskFile: " + ex.Message);
            }
        }

        private void ClearCurrentTaskFile()
        {
            try
            {
                ExcelLogReader.ClearCurrentTaskFile();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[AppBarWindow] ClearCurrentTaskFile: " + ex.Message);
            }
        }

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

        private async void ReturnToResultButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Quit 後もプロセスが残ると VSTO が再ロードされずログタブが消えるため、
                // PID を記録して EnsureExcelProcessExited まで完了してから結果画面へ遷移する。
                await Task.Run(() => CloseExcelForReturnToResult());

                _viewModel?.ClearSharedExcelApplication();

                if (_resultWindow != null && !_resultWindow.IsVisible)
                {
                    _resultWindow.Show();
                    _resultWindow.Activate();
                    _resultWindow.Focus();
                }
                else
                {
                    var resultWindow = Application.Current.Windows.OfType<ResultWindow>().FirstOrDefault();
                    if (resultWindow != null)
                    {
                        resultWindow.Show();
                        resultWindow.Activate();
                        resultWindow.Focus();
                    }
                }

                this.Hide();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"結果画面表示エラー: {ex.Message}");
                MessageBox.Show("結果画面の表示に失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 結果画面に戻る前に Excel を保存して終了し、プロセスが残る場合は PID 単位で確実に終了させる。
        /// </summary>
        private static void CloseExcelForReturnToResult()
        {
            ExcelApp excelApp = null;
            int excelPid = -1;

            try
            {
                try
                {
                    excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    System.Diagnostics.Debug.WriteLine("[ReturnToResult] No Excel application");
                    return;
                }

                excelPid = ExcelApplicationManager.TryGetExcelProcessId(excelApp);

                bool originalDisplayAlerts = true;
                try
                {
                    originalDisplayAlerts = excelApp.DisplayAlerts;
                }
                catch
                {
                    /* ignore */
                }

                try
                {
                    excelApp.DisplayAlerts = false;

                    var workbooks = excelApp.Workbooks;
                    foreach (ExcelWorkbook workbook in workbooks)
                    {
                        try
                        {
                            if (workbook.Path != "")
                            {
                                workbook.Save();
                            }
                            else
                            {
                                workbook.Saved = true;
                            }

                            Marshal.ReleaseComObject(workbook);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReturnToResult] workbook: {ex.Message}");
                        }
                    }

                    Marshal.ReleaseComObject(workbooks);
                }
                finally
                {
                    try
                    {
                        excelApp.DisplayAlerts = originalDisplayAlerts;
                    }
                    catch
                    {
                        /* ignore */
                    }
                }

                try
                {
                    excelApp.Quit();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ReturnToResult] Quit: {ex.Message}");
                }

                try
                {
                    Marshal.ReleaseComObject(excelApp);
                }
                catch
                {
                    /* ignore */
                }

                excelApp = null;

                const int quitWaitMs = 10000;
                if (excelPid > 0)
                {
                    ExcelApplicationManager.EnsureExcelProcessExited(
                        excelPid,
                        quitWaitMs,
                        5000,
                        "[ReturnToResult]");
                }
                else
                {
                    ExcelApplicationManager.WaitForAllExcelProcessesGone(quitWaitMs);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReturnToResult] {ex.Message}");
                if (excelPid > 0)
                {
                    ExcelApplicationManager.EnsureExcelProcessExited(
                        excelPid,
                        10000,
                        5000,
                        "[ReturnToResult]");
                }
            }
        }
    }

    // JSON デシリアライズ用のクラス
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
