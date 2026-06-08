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
using System.Windows.Interop;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WordDoc = Microsoft.Office.Interop.Word.Document;
using Libraries;

namespace MOS_Word_app.Views
{
    /// <summary>
    /// AppBarWindow.xaml の相互作用ロジック（Wordアプリ用）
    /// </summary>
    public partial class AppBarWindow : System.Windows.Window
    {
        // Windows API用の定義
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        
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
        private DateTime _projectStartTime; // プロジェクト開始時刻
        private DispatcherTimer _projectTimer; // プロジェクト用タイマー（5分制限）
        private Dictionary<int, Dictionary<int, string>> _clipboardTargets = new Dictionary<int, Dictionary<int, string>>(); // クリップボード対象（プロジェクトID → タスクID → 問題文）
        private bool _isPaused = false; // 一時停止状態
        private DateTime _pauseStartTime; // 一時停止開始時刻（プロジェクトタイマー用）
        private List<System.Windows.Controls.Button> _dynamicTaskButtons = new List<System.Windows.Controls.Button>(); // 動的に生成されたタスクボタン（8番目以降）
        
        // プロジェクト閲覧状況の追跡（50分制限対応）
        private DateTime _examStartTime; // 試験開始時刻
        private Dictionary<int, DateTime> _projectFirstViewTime = new Dictionary<int, DateTime>(); // プロジェクトの初回閲覧時刻

        // 結果画面から戻ってきたときに「結果に戻る」ボタンとして振る舞うための状態
        private bool _isReturnToResultMode = false;
        private Views.ResultWindow _lastResultWindow;
        private readonly HashSet<string> _initialWrongTaskKeys = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _retryTaskKeys = new HashSet<string>(StringComparer.Ordinal);
        private readonly HashSet<string> _preparedRetryTaskKeys = new HashSet<string>(StringComparer.Ordinal);
        private string _lastTaskStartKey;
        private const string ReviewPageButtonLabel = "レビューページ";
        private const string ReturnToResultButtonLabel = "結果に戻る";
        
        public AppBarWindow(int projectId = 1, int groupId = 1)
        {
            InitializeComponent();
            _currentProjectId = projectId;
            _groupId = groupId; // グループIDを保存
            _examStartTime = DateTime.Now; // 試験開始時刻を記録
            InitializeTimer();
            InitializeProjectTimer();
            LoadClipboardTargets(); // クリップボード対象を先に読み込む
            LoadTasks();
            UpdateTaskDisplay();
            SetWindowPosition();
            // 戻ってきたときにタスクボタンの状態を最新に保つ
            this.Activated += (s, e) =>
            {
                try
                {
                    UpdateTaskButtons();
                    UpdateButtonTexts();
                }
                catch { }
            };
            // 注意: WordドキュメントはMainViewModelのExecuteOpenProjectで既に開かれている
            // ここでは開かない（PositionWordWindowはSetWindowPositionで呼ばれる）
            
            // 現在のプロジェクトの初回閲覧時刻を記録
            if (!_projectFirstViewTime.ContainsKey(_currentProjectId))
            {
                _projectFirstViewTime[_currentProjectId] = DateTime.Now;
            }
        }
        
        /// <summary>試験用レイアウト（Word ＋ アプリバー）を適用する。再表示・再開時に呼ぶ。</summary>
        public void ApplyExamWindowLayout()
        {
            SetWindowPosition();
        }

        private void SetWindowPosition()
        {
            var screenWidth = (int)SystemParameters.PrimaryScreenWidth;
            var screenHeight = (int)SystemParameters.PrimaryScreenHeight;
            const double barHeight = WordWindowLayoutHelper.DefaultAppBarHeight;
            int wordTop = (int)(screenHeight - barHeight);

            WordWindowLayoutHelper.PositionWordForExamMode(barHeight, screenWidth, screenHeight);

            IntPtr hWnd = new WindowInteropHelper(this).Handle;
            if (hWnd == IntPtr.Zero)
            {
                this.Width = screenWidth;
                this.Height = barHeight;
                this.Left = 0;
                this.Top = wordTop;
                this.Topmost = true;
                return;
            }

            GetWindowRect(hWnd, out RECT windowRect);
            GetClientRect(hWnd, out RECT clientRect);

            int borderWidth = (windowRect.right - windowRect.left) - clientRect.right;
            int borderHeight = (windowRect.bottom - windowRect.top) - clientRect.bottom;

            int x = -borderWidth / 2;
            int y = wordTop - borderHeight / 2;
            int width = screenWidth + borderWidth;
            int height = (int)barHeight + borderHeight;

            MoveWindow(hWnd, x, y, width, height, true);
            this.Topmost = true;
        }

        private void PositionWordWindow()
        {
            var screenWidth = (int)SystemParameters.PrimaryScreenWidth;
            var screenHeight = (int)SystemParameters.PrimaryScreenHeight;
            WordWindowLayoutHelper.PositionWordForExamMode(
                WordWindowLayoutHelper.DefaultAppBarHeight, screenWidth, screenHeight);
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            // 実測高さが確定した後に再配置して、上方向にずれる問題を防ぐ
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
            if (!MainWindow.IsTimerDisabled)
                _timer.Start();
        }
        
        private void InitializeProjectTimer()
        {
            // プロジェクト用タイマーを初期化（5分制限）
            _projectTimer = new DispatcherTimer();
            _projectTimer.Interval = TimeSpan.FromSeconds(1);
            _projectTimer.Tick += ProjectTimer_Tick;
            
            // 最初のプロジェクト開始時刻を設定
            _projectStartTime = DateTime.Now;
            if (!MainWindow.IsTimerDisabled)
                _projectTimer.Start();
        }
        
        private void Timer_Tick(object sender, EventArgs e)
        {
            if (MainWindow.IsTimerDisabled) return;
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
        
        private void ProjectTimer_Tick(object sender, EventArgs e)
        {
            if (MainWindow.IsTimerDisabled) return;
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
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            var owner = Application.Current.MainWindow;
            var result = owner != null
                ? MessageBox.Show(owner, "アプリ自体を終了します。本当にいいですか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question)
                : MessageBox.Show("アプリ自体を終了します。本当にいいですか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
                return;
            _timer?.Stop();
            this.Close();
        }
        
        private async void ReviewPageButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 「結果に戻る」モードの場合は、隠れている ResultWindow を再表示する
                if (_isReturnToResultMode && _lastResultWindow != null && !_lastResultWindow.IsVisible)
                {
                    await TryRescorePendingRetryTasksAsync();
                    CloseWordDocumentsBeforeReturnToResult();
                    this.Hide();
                    _lastResultWindow.Show();
                    _lastResultWindow.Activate();
                    ClearReturnToResultMode();
                    return;
                }

                // 通常モード: レビューページを開く
                this.Hide();

                // 現在のタイマー残り時間と状態情報を渡す（AppBarWindowのインスタンスとgroupIdも渡す）
                var reviewWindow = new ReviewPageWindow(_remainingTime, _projectTaskCompletedStates, _projectTaskFlaggedStates, null, this, _groupId);
                reviewWindow.OnNavigateToTask = NavigateToTask;
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

        /// <summary>
        /// 結果画面からタスクに戻ってきたときに、「結果に戻る」モードに切り替える。
        /// </summary>
        public void SetReturnToResultMode(Views.ResultWindow resultWindow)
        {
            _lastResultWindow = resultWindow;
            _isReturnToResultMode = true;
            UpdateReviewPageButtonLabel(ReturnToResultButtonLabel);
        }

        /// <summary>
        /// 結果画面を閉じたときなどに、戻りモードを解除する。
        /// </summary>
        public void ClearReturnToResultMode()
        {
            _isReturnToResultMode = false;
            _lastResultWindow = null;
            _retryTaskKeys.Clear();
            UpdateReviewPageButtonLabel(ReviewPageButtonLabel);
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

        public bool TryEnqueueRetryTask(int projectId, int taskId)
        {
            string key = $"{projectId}-{taskId}";
            if (!_initialWrongTaskKeys.Contains(key))
                return false;
            if (IsTaskFlaggedForRetry(projectId, taskId))
                return false;
            _retryTaskKeys.Add(key);
            return true;
        }

        private void SyncTaskTracking()
        {
            if (_isReturnToResultMode)
                EnsureRetryAttemptPrepared(_currentProjectId, _currentTaskId);
            WriteCurrentTaskFile();
            LogTaskStartIfNeeded();
        }

        private void WriteCurrentTaskFile()
        {
            try
            {
                int attemptNo = WordTaskAttemptRegistry.GetAttempt(_currentProjectId, _currentTaskId);
                var flags = WordTaskValidationConfig.GetExemptFlags(_currentProjectId, _currentTaskId);
                string content = $"{_currentProjectId},{_currentTaskId},{(int)flags},{attemptNo}";
                File.WriteAllText(LogReader.GetCurrentTaskFilePath(), content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[AppBarWindow] WriteCurrentTaskFile: " + ex.Message);
            }
        }

        private void LogTaskStartIfNeeded()
        {
            int attemptNo = WordTaskAttemptRegistry.GetAttempt(_currentProjectId, _currentTaskId);
            string key = $"{_currentProjectId}-{_currentTaskId}-{attemptNo}";
            if (string.Equals(_lastTaskStartKey, key, StringComparison.Ordinal))
                return;
            _lastTaskStartKey = key;
            LogReader.LogTaskStart(_currentProjectId, _currentTaskId, attemptNo);
        }

        private bool EnsureRetryAttemptPrepared(int projectId, int taskId)
        {
            string key = $"{projectId}-{taskId}";
            if (!_initialWrongTaskKeys.Contains(key))
                return false;
            if (_preparedRetryTaskKeys.Contains(key))
                return true;
            int cur = WordTaskAttemptRegistry.GetAttempt(projectId, taskId);
            WordTaskAttemptRegistry.SetAttempt(projectId, taskId, Math.Max(1, cur + 1));
            _preparedRetryTaskKeys.Add(key);
            return true;
        }

        /// <summary>
        /// 結果画面からの復習中に表示したタスクを再採点キューへ登録する（PP の WriteCurrentTaskFile 相当）。
        /// </summary>
        private void RegisterCurrentTaskForRetryIfFromResult()
        {
            if (!_isReturnToResultMode)
                return;
            if (TryEnqueueRetryTask(_currentProjectId, _currentTaskId))
                System.Diagnostics.Debug.WriteLine($"[Retry] Enqueued P{_currentProjectId} T{_currentTaskId}");
        }

        private bool IsTaskFlaggedForRetry(int projectId, int taskId)
        {
            if (!_projectTaskFlaggedStates.ContainsKey(projectId))
                return false;
            var flags = _projectTaskFlaggedStates[projectId];
            int idx = taskId - 1;
            return idx >= 0 && idx < flags.Length && flags[idx];
        }

        private async Task TryRescorePendingRetryTasksAsync()
        {
            if (_retryTaskKeys.Count == 0 || _lastResultWindow == null)
                return;

            SaveAllWordDocuments();

            var keysToScore = _retryTaskKeys.ToList();
            await Task.Run(() =>
            {
                foreach (var key in keysToScore)
                {
                    if (!TryParseRetryTaskKey(key, out int projectId, out int taskId))
                        continue;
                    bool? passed = WordBatchScoring.ScoreSingleTask(_groupId, projectId, taskId);
                    if (passed.HasValue)
                        _retryTaskKeys.Remove(key);
                }
            });

            try
            {
                await _lastResultWindow.RefreshResultsAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TryRescorePendingRetryTasks] Error: {ex.Message}");
            }
        }

        private static bool TryParseRetryTaskKey(string key, out int projectId, out int taskId)
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

        private void UpdateReviewPageButtonLabel(string label)
        {
            var btn = this.FindName("ReviewPageButton") as System.Windows.Controls.Button;
            if (btn != null)
                btn.Content = label;
        }
        
        private void NavigateToTask(int projectId, int taskId)
        {
            System.Diagnostics.Debug.WriteLine($"NavigateToTask called in main window: ProjectId={projectId}, TaskId={taskId}");
            
            try
            {
                // レビューページから戻ったときは該当する Word ドキュメントを起動する
                OpenProjectDocument(projectId, _groupId);
                
                // プロジェクトを変更
                if (projectId != _currentProjectId)
                {
                    System.Diagnostics.Debug.WriteLine($"プロジェクト変更: {_currentProjectId} -> {projectId}");
                    _currentProjectId = projectId;
                    LoadCurrentProjectTasks();
                    
                    // プロジェクトの初回閲覧時刻を記録
                    if (!_projectFirstViewTime.ContainsKey(_currentProjectId))
                    {
                        _projectFirstViewTime[_currentProjectId] = DateTime.Now;
                    }
                }
                
                // タスクを変更
                if (taskId != _currentTaskId && taskId >= 1 && taskId <= _tasks.Count)
                {
                    System.Diagnostics.Debug.WriteLine($"タスク変更: {_currentTaskId} -> {taskId}");
                    _currentTaskId = taskId;
                }
                
                // UIを更新
                UpdateTaskDisplay();
                
                // メインウィンドウを表示（レビューページから戻る時）
                this.Show();
                this.WindowState = WindowState.Normal;
                this.Activate();
                this.Focus();
                ApplyExamWindowLayout();
                
                System.Diagnostics.Debug.WriteLine($"プロジェクト{projectId}のタスク{taskId}に移動しました");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ナビゲーションエラー: {ex.Message}");
            }
        }
        
        private void LoadTasks()
        {
            try
            {
                // 試験バーの問題文は Word 用 JSON（PowerPoint と独立）
                string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", "MOS模擬アプリ問題文一覧_Word.json");
                if (!File.Exists(jsonPath))
                    jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOS模擬アプリ問題文一覧_Word.json");
                
                if (!File.Exists(jsonPath))
                {
                    System.Diagnostics.Debug.WriteLine($"[AppBarWindow] Word用問題文JSONが見つかりません: References\\JSON および出力ルートを確認");
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
                string copyTargetJsonName = "MOS模擬アプリ_入力追加変更挿入問題_Word.json";
                string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, copyTargetJsonName);
                if (!File.Exists(jsonPath))
                    jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", copyTargetJsonName);
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
                    int arraySize = Math.Max(taskCount, 100); // タスク数、最小100（配列は0始まりなので+1は不要）
                    
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
                int totalProjects = _projectData.Projects.Max(p => p.ProjectId);
                var projectInfoTextBlock = FindName("ProjectInfoTextBlock") as System.Windows.Controls.TextBlock;
                if (projectInfoTextBlock != null)
                {
                    projectInfoTextBlock.Text = $"プロジェクト {_currentProjectId}/{totalProjects}";
                }
            }
        }
        
        /// <summary>
        /// 現在のプロジェクト・タスクに対応する問題文をWord文書から取得する。
        /// 1-3: 最後の段落に「以降、「CSR活動の進め方」については「前田ゼミ研究ノートVOL.2」を参照。」が含まれる場合、その段落のテキストを返す。
        /// </summary>
        private string GetProblemTextFromWordDocument(int projectId, int taskId)
        {
            if (projectId != 1 || taskId != 3)
                return null;
            const string searchText = "以降、「CSR活動の進め方」については「前田ゼミ研究ノートVOL.2」を参照。";
            WordApp wordApp = null;
            WordDoc doc = null;
            try
            {
                wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                if (wordApp?.ActiveDocument == null) return null;
                doc = wordApp.ActiveDocument;
                int count = doc.Paragraphs.Count;
                if (count < 1) return null;
                var para = doc.Paragraphs[count];
                Microsoft.Office.Interop.Word.Range range = null;
                try
                {
                    range = para.Range;
                    string text = range?.Text ?? "";
                    if (text.Contains(searchText))
                    {
                        return text.Trim().TrimEnd('\r', '\n', '\a');
                    }
                }
                finally
                {
                    if (range != null) Marshal.ReleaseComObject(range);
                }
            }
            catch { }
            finally
            {
                if (doc != null) Marshal.ReleaseComObject(doc);
            }
            return null;
        }

        private void UpdateTaskDisplay()
        {
            System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] 開始: プロジェクト{_currentProjectId}, タスク{_currentTaskId}");
            
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
            
            // Word文書から問題文を取得（1-3は最後の段落の指定文字列を表示）
            string wordDescription = GetProblemTextFromWordDocument(_currentProjectId, _currentTaskId);
            if (wordDescription != null)
            {
                System.Diagnostics.Debug.WriteLine($"[UpdateTaskDisplay] Word文書から問題文を取得して表示");
                SetTextWithUnderline(taskDescriptionTextBlock, wordDescription);
                UpdateTaskButtons();
                UpdateButtonTexts();
                RegisterCurrentTaskForRetryIfFromResult();
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
            SyncTaskTracking();
            RegisterCurrentTaskForRetryIfFromResult();
        }
        
        /// <summary>
        /// テキスト内の"で囲まれた部分と「～を入力してください」パターンに下線を付けてTextBlockに設定します
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
            
            // 「～を入力してください」パターンを検出する正規表現
            // 例: 「テキストを入力してください」「～を入力してください」など
            Regex inputPattern = new Regex(@"「([^」]+)を入力してください」", RegexOptions.Compiled);
            
            // まず「～を入力してください」パターンを処理
            var inputMatches = inputPattern.Matches(text);
            if (inputMatches.Count > 0)
            {
                int lastIndex = 0;
                foreach (Match match in inputMatches)
                {
                    // マッチ前のテキストを追加
                    if (match.Index > lastIndex)
                    {
                        string beforeText = text.Substring(lastIndex, match.Index - lastIndex);
                        if (!string.IsNullOrEmpty(beforeText))
                        {
                            textBlock.Inlines.Add(new Run(beforeText));
                        }
                    }
                    
                    // 「～を入力してください」の「～」部分を下線付きで追加
                    string inputText = match.Groups[1].Value; // 「」内のテキスト（「を入力してください」を除く）
                    var run = new Run(inputText);
                    run.TextDecorations = TextDecorations.Underline;
                    run.Cursor = Cursors.Hand;
                    run.Foreground = new SolidColorBrush(Colors.Blue);
                    run.MouseDown += (sender, e) => OnUnderlinedTextClick(inputText, e);
                    textBlock.Inlines.Add(run);
                    
                    // 「を入力してください」の部分を通常テキストとして追加
                    textBlock.Inlines.Add(new Run("を入力してください"));
                    
                    lastIndex = match.Index + match.Length;
                }
                
                // 残りのテキストを追加
                if (lastIndex < text.Length)
                {
                    string remainingText = text.Substring(lastIndex);
                    if (!string.IsNullOrEmpty(remainingText))
                    {
                        textBlock.Inlines.Add(new Run(remainingText));
                    }
                }
                
                System.Diagnostics.Debug.WriteLine($"[SetTextWithUnderline] 「～を入力してください」パターンを処理しました");
                return;
            }
            
            // 「～を入力してください」パターンが見つからない場合は、"で囲まれた部分を処理
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
                _projectTaskCompletedStates[_currentProjectId] : new bool[Math.Max(maxTaskCount, 100)];
            bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(_currentProjectId) ? 
                _projectTaskFlaggedStates[_currentProjectId] : new bool[Math.Max(maxTaskCount, 100)];

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
                button.BorderThickness = new System.Windows.Thickness(2.0);
                button.Foreground = System.Windows.Media.Brushes.Black;
            }
            else
            {
                // 他のタスクボタンは非選択状態
                button.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.LightGray);
                button.BorderThickness = new System.Windows.Thickness(0.0);
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
                BorderThickness = new System.Windows.Thickness(0.0),
                Margin = new System.Windows.Thickness(0.0, 0.0, 5.0, 0.0),
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
                Name = $"TaskText{taskId}",
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
                Name = $"Check{taskId}",
                Text = "✓",
                FontSize = 16,
                Foreground = System.Windows.Media.Brushes.Green,
                FontWeight = System.Windows.FontWeights.Bold,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new System.Windows.Thickness(0.0, 0.0, 15.0, 0.0),
                Visibility = System.Windows.Visibility.Collapsed
            };
            grid.Children.Add(checkTextBlock);
            
            // 旗マークのTextBlock
            var flagTextBlock = new System.Windows.Controls.TextBlock
            {
                Name = $"Flag{taskId}",
                Text = "🚩",
                FontSize = 16,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Margin = new System.Windows.Thickness(15.0, 0.0, 0.0, 0.0),
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
                _projectTaskCompletedStates[_currentProjectId] : new bool[Math.Max(taskCount, 100)];
            bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(_currentProjectId) ? 
                _projectTaskFlaggedStates[_currentProjectId] : new bool[Math.Max(taskCount, 100)];
            
            // 解答済みボタンのテキストを更新
            var completeButton = FindName("CompleteButton") as System.Windows.Controls.Button;
            
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
            
            // フラグボタンのテキストを更新
            var flagButton = FindName("FlagButton") as System.Windows.Controls.Button;
            
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
            var button = sender as System.Windows.Controls.Button;
            if (button == null || button.Tag == null) return;
            int taskId = int.Parse(button.Tag.ToString());
            if (taskId >= 1 && taskId <= _tasks.Count)
            {
                _currentTaskId = taskId;
                UpdateTaskDisplay();
            }
        }
        
        private void NextProject_Click(object sender, RoutedEventArgs e)
        {
            if (TryShowObjectSelectedWarningIfWordObjectSelected())
                return;
            MoveToNextProject();
        }

        /// <summary>
        /// Wordでグラフ・図・SmartArt・テーブル・プレースホルダーなどが選択されている場合に警告ダイアログを表示する。
        /// </summary>
        private bool TryShowObjectSelectedWarningIfWordObjectSelected()
        {
            try
            {
                WordApp wordApp = null;
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch (COMException)
                {
                    return false;
                }
                if (wordApp?.Selection == null) return false;

                bool hasObject = false;
                try
                {
                    var sel = wordApp.Selection;
                    if (sel.Tables.Count > 0) hasObject = true;
                    else if (sel.InlineShapes.Count > 0) hasObject = true;
                    else
                    {
                        try
                        {
                            if (sel.ShapeRange.Count > 0) hasObject = true;
                        }
                        catch { }
                    }
                }
                catch { }
                if (!hasObject) return false;

                var w = new ObjectSelectedWarningWindow();
                w.ShowDialog();
                return true;
            }
            catch
            {
                return false;
            }
        }
        
        private void MoveToNextProject()
        {
            // 次のプロジェクトに移る前に現在のプロジェクト（Wordドキュメント）を保存する
            SaveAllWordDocuments();
            // 前プロジェクトの文書を閉じ、ActiveDocument の取り違えを防ぐ（プロセス kill は行わない）
            if (!CloseAllWordDocuments())
            {
                TryQuitWord();
                Thread.Sleep(500);
            }
            else
            {
                Thread.Sleep(200);
            }

            // プロジェクトの最大数をチェック（JSONファイルの最大プロジェクトID）
            int maxProjectId = _projectData?.Projects?.Max(p => p.ProjectId) ?? 1;
            
            // 次のプロジェクトに移動
            _currentProjectId++;
            
            if (_currentProjectId > maxProjectId)
            {
                // 最後のプロジェクトを超えた場合はレビューページに移動
                System.Diagnostics.Debug.WriteLine($"プロジェクト{maxProjectId}を超えたため、レビューページに移動します");
                ReviewPageButton_Click(null, null);
                return;
            }
            
            // 新しいプロジェクトのWordドキュメントを開く
            OpenProjectDocument(_currentProjectId, _groupId);
            
            // プロジェクト変更時は状態をリセットしない（Dictionaryで管理）
            
            // 新しいプロジェクトのタスクを読み込み
            LoadCurrentProjectTasks();
            UpdateTaskDisplay();
            
            // プロジェクトタイマーをリセット
            ResetProjectTimer();
            
            // プロジェクトの初回閲覧時刻を記録
            if (!_projectFirstViewTime.ContainsKey(_currentProjectId))
            {
                _projectFirstViewTime[_currentProjectId] = DateTime.Now;
            }
            
            System.Diagnostics.Debug.WriteLine($"プロジェクト{_currentProjectId}に移動しました");
        }
        
        private void OpenProjectDocument(int projectId, int groupId)
        {
            string filePath = null;
            try
            {
                string basePath = @"C:\MOSTest\Word365";
                string workingFolder = Path.Combine(basePath, $"Tab{groupId}");
                string initialFolder = Path.Combine(basePath, $"Tab{groupId}", "Initial");
                string initialInitialFolder = Path.Combine(basePath, $"Tab{groupId}", "Initial", "Initial");
                string[] possibleNames = (groupId == 1 && projectId == 7)
                    ? new[] { $"Project{projectId}.doc", $"project{projectId}.doc" }
                    : new[] { $"Project{projectId}.docx", $"Project{projectId}.doc", $"project{projectId}.docx", $"project{projectId}.doc" };
                string workingFileName = (groupId == 1 && projectId == 7) ? "Project7.doc" : $"Project{projectId}.docx";
                string workingFilePath = Path.Combine(workingFolder, workingFileName);

                // 保存先は Tab\ 直下のみ。まず作業フォルダを参照
                foreach (var fileName in possibleNames)
                {
                    string fullPath = Path.Combine(workingFolder, fileName);
                    if (File.Exists(fullPath))
                    {
                        filePath = fullPath;
                        break;
                    }
                }
                // 作業フォルダに無ければ Initial からコピーしてから開く
                if (string.IsNullOrEmpty(filePath))
                {
                    string sourcePath = null;
                    foreach (var fileName in possibleNames)
                    {
                        string fullPath = Path.Combine(initialFolder, fileName);
                        if (File.Exists(fullPath)) { sourcePath = fullPath; break; }
                    }
                    if (string.IsNullOrEmpty(sourcePath) && Directory.Exists(initialInitialFolder))
                    {
                        foreach (var fileName in possibleNames)
                        {
                            string fullPath = Path.Combine(initialInitialFolder, fileName);
                            if (File.Exists(fullPath)) { sourcePath = fullPath; break; }
                        }
                    }
                    if (!string.IsNullOrEmpty(sourcePath))
                    {
                        try
                        {
                            if (!Directory.Exists(workingFolder))
                                Directory.CreateDirectory(workingFolder);
                            File.Copy(sourcePath, workingFilePath, overwrite: false);
                            filePath = workingFilePath;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[OpenProjectDocument] Initialからコピーエラー: {ex.Message}");
                        }
                    }
                }
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                {
                    System.Diagnostics.Debug.WriteLine($"プロジェクト{projectId}のファイルが見つかりません: {workingFolder} または {initialFolder}");
                    PositionWordWindow();
                    return;
                }
                
                WordApp wordApp = null;
                bool wordWasNotRunning = false;
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch (COMException)
                {
                    wordApp = new WordApp();
                    wordApp.Visible = true;
                    wordWasNotRunning = true;
                    // Word を起動した直後は COM が準備できるまで少し待つ
                    System.Threading.Thread.Sleep(500);
                }
                
                if (wordApp == null) return;
                
                // 同じパスで既に開いているドキュメントがあれば保存せずに閉じる（初期化状態でも常にフォルダから開く）
                string pathLower = System.IO.Path.GetFullPath(filePath).ToLowerInvariant();
                try
                {
                    for (int i = wordApp.Documents.Count; i >= 1; i--)
                    {
                        WordDoc doc = wordApp.Documents[i];
                        try
                        {
                            string fullName = doc.FullName?.ToLowerInvariant() ?? "";
                            string docFullPath = fullName;
                            try { docFullPath = System.IO.Path.GetFullPath(fullName).ToLowerInvariant(); } catch { }
                            if (fullName == pathLower || docFullPath == pathLower)
                            {
                                doc.Close(SaveChanges: false);
                                break;
                            }
                        }
                        finally
                        {
                            if (doc != null) Marshal.ReleaseComObject(doc);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[OpenProjectDocument] 既存ドキュメント閉じる際のエラー: {ex.Message}");
                }

                try
                {
                    wordApp.Documents.Open(filePath, ReadOnly: false, Visible: true); // 常にフォルダから開く
                    if (wordWasNotRunning)
                        System.Threading.Thread.Sleep(300);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Documents.Open エラー: {ex.Message}");
                }
                
                ApplyExamWindowLayout();
                System.Diagnostics.Debug.WriteLine($"プロジェクト{projectId}のドキュメントを開きました: {filePath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"プロジェクトドキュメントを開く際のエラー: {ex.Message}");
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
                int arraySize = Math.Max(taskCount, 100); // 配列は0始まりなので+1は不要
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
                int arraySize = Math.Max(taskCount, 100); // 配列は0始まりなので+1は不要
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
        
        /// <summary>
        /// 50分制限内に目を通せていないプロジェクトのタスク数を取得
        /// </summary>
        public int GetUnviewedTasksCount()
        {
            int unviewedCount = 0;
            var elapsed = DateTime.Now - _examStartTime;
            
            if (_projectData?.Projects != null)
            {
                foreach (var project in _projectData.Projects)
                {
                    // プロジェクトの初回閲覧時刻を取得
                    if (!_projectFirstViewTime.ContainsKey(project.ProjectId))
                    {
                        // 初回閲覧時刻が記録されていない場合は、50分経過していないかチェック
                        if (elapsed.TotalMinutes < 50.0)
                        {
                            // 50分経過していない場合、このプロジェクトのタスク数をカウント
                            unviewedCount += project.Tasks?.Count ?? 0;
                        }
                    }
                    else
                    {
                        // 初回閲覧時刻が記録されている場合、50分以内に閲覧されたかチェック
                        var projectViewTime = _projectFirstViewTime[project.ProjectId];
                        var timeFromStart = projectViewTime - _examStartTime;
                        if (timeFromStart.TotalMinutes > 50.0)
                        {
                            // 50分を超えてから閲覧された場合、このプロジェクトのタスク数をカウント
                            unviewedCount += project.Tasks?.Count ?? 0;
                        }
                    }
                }
            }
            
            return unviewedCount;
        }
        
        /// <summary>
        /// 「あとで見直す」状態の問題数を取得
        /// </summary>
        public int GetFlaggedTasksCount()
        {
            int flaggedCount = 0;
            
            foreach (var projectFlaggedStates in _projectTaskFlaggedStates)
            {
                foreach (bool isFlagged in projectFlaggedStates.Value)
                {
                    if (isFlagged)
                    {
                        flaggedCount++;
                    }
                }
            }
            
            return flaggedCount;
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
                    // リセット前に開いているWordドキュメントを閉じる。閉じられない場合はWordを終了してからリセット
                    if (!CloseAllWordDocuments())
                        TryQuitWord();
                    Thread.Sleep(500); // Word がファイルハンドルを解放するまで待つ
                    ResetProject(_groupId, _currentProjectId);
                    
                    // リセット後、Wordドキュメントを再読み込み
                    OpenProjectDocument(_currentProjectId, _groupId);

                    // Word が立ち上がった後に、アプリバー最前面で完了メッセージを表示する
                    bool originalTopmost = this.Topmost;
                    try
                    {
                        this.Topmost = true;
                        this.Activate();
                        MessageBox.Show(this, "プロジェクトをリセットしました。", "リセット完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    finally
                    {
                        this.Topmost = originalTopmost;
                    }

                    ApplyExamWindowLayout();

                    // 該当プロジェクトの解答済み・フラグ状態をクリア（Initial\Initial の内容に合わせて初期表示）
                    int taskCount = _tasks != null ? _tasks.Count : 0;
                    int arraySize = Math.Max(taskCount, 100);
                    _projectTaskCompletedStates[_currentProjectId] = new bool[arraySize];
                    _projectTaskFlaggedStates[_currentProjectId] = new bool[arraySize];
                    UpdateTaskDisplay();
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
            WordProjectResetHelper.ResetProject(groupId, projectId);
            _lastTaskStartKey = null;
        }
        
        /// <summary>
        /// 全プロジェクトの解答済み・フラグ状態をクリアし、表示を更新する。
        /// すべてリセット実行後に MainWindow から呼ばれる。
        /// </summary>
        public void ClearAllProjectStates()
        {
            _projectTaskCompletedStates.Clear();
            _projectTaskFlaggedStates.Clear();
            UpdateTaskDisplay();
        }
        
        /// <summary>
        /// 開いているすべてのWord文書を保存する（Wordは終了しない）。
        /// 次のプロジェクトに移る前に呼ぶ。
        /// </summary>
        private void SaveAllWordDocuments()
        {
            try
            {
                WordApp wordApp = null;
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch (COMException)
                {
                    return;
                }
                if (wordApp == null) return;
                try
                {
                    wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                    for (int i = wordApp.Documents.Count; i >= 1; i--)
                    {
                        try
                        {
                            WordDoc doc = wordApp.Documents[i];
                            if (doc.Saved == false)
                            {
                                doc.Save();
                            }
                            Marshal.ReleaseComObject(doc);
                        }
                        catch (COMException comEx) when (comEx.HResult == unchecked((int)0x80010108))
                        {
                            break;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[SaveAllWordDocuments] 保存エラー: {ex.Message}");
                        }
                    }
                }
                finally
                {
                    try { if (wordApp != null) Marshal.ReleaseComObject(wordApp); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SaveAllWordDocuments] Error: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 結果画面に戻る前に、開いている Word 文書を保存して閉じる。
        /// </summary>
        private void CloseWordDocumentsBeforeReturnToResult()
        {
            SaveAllWordDocuments();
            if (!CloseAllWordDocuments())
            {
                TryQuitWord();
                Thread.Sleep(500);
            }
            else
            {
                Thread.Sleep(200);
            }
        }

        /// <summary>
        /// 開いているすべてのWord文書を閉じる（保存しない）。
        /// リセット実行前に呼び、上書き後に開き直すため。
        /// </summary>
        /// <returns>閉じた、またはWordが起動していない場合は true。COM切断などで閉じられなかった場合は false。</returns>
        private bool CloseAllWordDocuments()
        {
            try
            {
                WordApp wordApp = null;
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch (COMException)
                {
                    return true; // Word が起動していない場合はリセット可能
                }
                if (wordApp == null) return true;
                while (wordApp.Documents.Count > 0)
                {
                    WordDoc openDoc = null;
                    try
                    {
                        openDoc = wordApp.Documents[1];
                        openDoc.Close(SaveChanges: false);
                    }
                    catch (COMException comEx) when (comEx.HResult == unchecked((int)0x80010108))
                    {
                        // 既に切断されている場合は閉じられない
                        System.Diagnostics.Debug.WriteLine($"[CloseAllWordDocuments] ドキュメントは既に切断されています: {comEx.Message}");
                        return false;
                    }
                    catch (Exception closeEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CloseAllWordDocuments] 閉じる際のエラー: {closeEx.Message}");
                    }
                    finally
                    {
                        try { if (openDoc != null) Marshal.ReleaseComObject(openDoc); } catch { }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CloseAllWordDocuments] Error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Word アプリケーションを終了する（保存しない）。ドキュメントを閉じられない場合のリセット用。
        /// </summary>
        private void TryQuitWord()
        {
            try
            {
                WordApp wordApp = null;
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch (COMException)
                {
                    return;
                }
                if (wordApp == null) return;
                try
                {
                    wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                    wordApp.Quit();
                }
                finally
                {
                    try { if (wordApp != null) Marshal.ReleaseComObject(wordApp); } catch { }
                }
                Thread.Sleep(1000); // プロセスが終了してファイルを解放するまで待つ
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TryQuitWord] Error: {ex.Message}");
            }
        }
        
        protected override void OnClosed(EventArgs e)
        {
            _timer?.Stop();
            _projectTimer?.Stop();
            base.OnClosed(e);
        }
    }
}
