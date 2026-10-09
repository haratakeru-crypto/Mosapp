using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using System.Windows.Media;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.InteropServices;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using MOSExcelMogiApp;
using Newtonsoft.Json.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Threading;
using System.Text;
using Libraries;
using MOSExcelMogiApp.Infrastructure;

namespace MOSExcelMogiApp.Views
{
    /// <summary>
    /// ReviewPageWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ReviewPageWindow : Window
    {

        public ICommand NavigateToTaskCommand { get; private set; }
        /// <summary>グループID、プロジェクトID、タスクIDの順。</summary>
        public Action<int, int, int> OnNavigateToTask { get; set; }
        private DispatcherTimer _timer;
        private TimeSpan _remainingTime;
        private int _groupId = 1; // Group番号（1=模擬①, 2=模擬②, 3=演習）
        private bool _isScoring;

        /// <summary>採点ワークフロー（STAスレッド）内で取得した Excel。Task.Run(MTA) からの COM 呼び出し失敗を避けるため共有する。</summary>
        private ExcelApp _scoringExcelApp;
        private bool _scoringExcelCreatedBySession;
        private int _scoringExcelCreatedPid = -1;
        private readonly Dictionary<string, ExcelWorkbook> _scoringWorkbooks =
            new Dictionary<string, ExcelWorkbook>(StringComparer.OrdinalIgnoreCase);

        /// <summary>バックグラウンドで実行中のExcel終了タスク。採点開始前に完了を待機するために使用します。</summary>
        private static Task _pendingExcelCloseTask;

        public static void SetPendingExcelCloseTask(Task task)
        {
            Interlocked.Exchange(ref _pendingExcelCloseTask, task);
        }

        private static async Task WaitForPendingExcelCloseAsync(string logPrefix)
        {
            var task = Interlocked.Exchange(ref _pendingExcelCloseTask, null);
            if (task == null)
                return;

            System.Diagnostics.Debug.WriteLine(logPrefix + " waiting...");
            try { await task; } catch { }
            System.Diagnostics.Debug.WriteLine(logPrefix + " finished.");
        }

        private Dictionary<int, bool[]> _projectTaskCompletedStates;
        private Dictionary<int, bool[]> _projectTaskFlaggedStates;
        
        public ReviewPageWindow(TimeSpan remainingTime, Dictionary<int, bool[]> completedStates, Dictionary<int, bool[]> flaggedStates, int groupId = 1)
        {
            System.Diagnostics.Debug.WriteLine($"ReviewPageWindow constructor called with remainingTime: {remainingTime}, groupId: {groupId}");
            
            InitializeComponent();
            _remainingTime = remainingTime;
            _projectTaskCompletedStates = completedStates;
            _projectTaskFlaggedStates = flaggedStates;
            _groupId = groupId;
            NavigateToTaskCommand = new RelayCommand<ReviewTaskInfo>(NavigateToTask);
            LoadAllProjects();
            InitializeTimer();
            
            System.Diagnostics.Debug.WriteLine("ReviewPageWindow initialization completed");
        }
        
        private void InitializeTimer()
        {
            UpdateTimerDisplay();
            
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(1);
            _timer.Tick += Timer_Tick;
            // 「タイマーを使用」にチェックが入っているときだけカウントダウン開始
            if (!MainWindow.IsTimerDisabled)
                _timer.Start();
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
        
        private void UpdateTimerDisplay()
        {
            var timerTextBlock = FindName("TimerTextBlock") as System.Windows.Controls.TextBlock;
            if (timerTextBlock != null)
            {
                timerTextBlock.Text = _remainingTime.ToString(@"hh\:mm\:ss");
            }
        }
        
        private void LoadAllProjects()
        {
            try
            {
                ProjectData projectData = null;
                
                // 模試①（GroupId=2）の場合はCSVファイルから読み込む
                if (_groupId == 2)
                {
                    projectData = LoadProjectsFromCsv(_groupId);
                }
                else
                {
                    // その他の場合はJSONファイルから読み込む
                    projectData = LoadProjectsFromJson(_groupId);
                }
                
                if (projectData?.Projects != null)
                {
                    var reviewProjects = new List<ReviewProjectInfo>();
                    
                    foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
                    {
                        var reviewProject = new ReviewProjectInfo
                        {
                            ProjectTitle = $"プロジェクト {project.ProjectId}",
                            Tasks = project.Tasks?.Select(task => new ReviewTaskInfo
                            {
                                TaskTitle = $"タスク {task.TaskId}",
                                Description = RemoveQuotes(task.Description),
                                GroupId = _groupId,
                                ProjectId = project.ProjectId,
                                TaskId = task.TaskId
                            }).ToList() ?? new List<ReviewTaskInfo>()
                        };
                        
                        reviewProjects.Add(reviewProject);
                    }
                    
                    ProjectsItemsControl.ItemsSource = reviewProjects;
                    
                    // UIが完全に読み込まれた後に状態表示を更新
                    this.Dispatcher.BeginInvoke(new Action(() => {
                        UpdateTaskStates();
                    }), DispatcherPriority.Loaded);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"レビューページ読み込みエラー: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show("問題文の読み込みに失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private ProjectData LoadProjectsFromCsv(int groupId)
        {
            try
            {
                // CSVファイルのパス
                string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "CSV", "解答手順あり模擬試験①問題文.csv");
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Loading from CSV: {csvPath} (GroupId: {groupId})");
                
                if (!File.Exists(csvPath))
                {
                    System.Diagnostics.Debug.WriteLine($"CSVファイルが見つかりません: {csvPath}");
                    return new ProjectData { Projects = new List<ProjectInfo>() };
                }
                
                // CSVファイルを読み込む
                var projects = new Dictionary<int, List<TaskInfo>>();
                string[] lines = File.ReadAllLines(csvPath, Encoding.UTF8);
                
                // ヘッダー行をスキップ（1行目）
                for (int i = 1; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (string.IsNullOrEmpty(line))
                        continue;
                    
                    // CSVのパース（カンマ区切り、ただし引用符内のカンマは考慮）
                    string[] fields = ParseCsvLine(line);
                    if (fields.Length < 3)
                        continue;
                    
                    // グループ,プロジェクト,問題文,解答操作
                    if (int.TryParse(fields[0], out int csvGroupId) && 
                        int.TryParse(fields[1], out int projectId) &&
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
                return new ProjectData
                {
                    Projects = projects.Select(kvp => new ProjectInfo
                    {
                        ProjectId = kvp.Key,
                        Tasks = kvp.Value
                    }).ToList()
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CSV読み込みエラー: {ex.Message}\n{ex.StackTrace}");
                return new ProjectData { Projects = new List<ProjectInfo>() };
            }
        }

        private ProjectData LoadProjectsFromJson(int groupId)
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
                
                string jsonPath = DataPathHelper.ResolveJsonPath(jsonFileName);
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Loading from: {jsonFileName} (GroupId: {groupId})");
                
                string jsonContent = File.ReadAllText(jsonPath);
                return JsonConvert.DeserializeObject<ProjectData>(jsonContent);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"JSON読み込みエラー: {ex.Message}\n{ex.StackTrace}");
                return new ProjectData { Projects = new List<ProjectInfo>() };
            }
        }

        /// <summary>
        /// CSV行をパースします（カンマ区切り、引用符内のカンマを考慮）
        /// </summary>
        private string[] ParseCsvLine(string line)
        {
            var fields = new List<string>();
            bool inQuotes = false;
            System.Text.StringBuilder currentField = new System.Text.StringBuilder();
            
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
        
        private void UpdateTaskStates()
        {
            try
            {
                // 各プロジェクトの各タスクの状態を更新
                foreach (var project in ProjectsItemsControl.ItemsSource as List<ReviewProjectInfo>)
                {
                    foreach (var task in project.Tasks)
                    {
                        // 解答済み状態を確認
                        int arrayIndex = task.TaskId - 1;
                        bool isCompleted = false;
                        bool isFlagged = false;
                        
                        if (_projectTaskCompletedStates?.ContainsKey(task.ProjectId) == true)
                        {
                            var completedArray = _projectTaskCompletedStates[task.ProjectId];
                            System.Diagnostics.Debug.WriteLine($"[ReviewPage] Project {task.ProjectId}, Task {task.TaskId}: Checking completedStates[{arrayIndex}], Array length={completedArray.Length}");
                            
                            if (arrayIndex >= 0 && arrayIndex < completedArray.Length)
                            {
                                isCompleted = completedArray[arrayIndex];
                                System.Diagnostics.Debug.WriteLine($"[ReviewPage] completedStates[{arrayIndex}] = {isCompleted}");
                                
                                // デバッグ: 配列の全状態を表示
                                for (int i = 0; i < completedArray.Length; i++)
                                {
                                    if (completedArray[i])
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[ReviewPage] completedStates[{i}] = true (corresponds to TaskId {i + 1})");
                                    }
                                }
                            }
                        }
                        
                        // あとで見直す状態を確認
                        if (_projectTaskFlaggedStates?.ContainsKey(task.ProjectId) == true)
                        {
                            var flaggedArray = _projectTaskFlaggedStates[task.ProjectId];
                            System.Diagnostics.Debug.WriteLine($"[ReviewPage] Project {task.ProjectId}, Task {task.TaskId}: Checking flaggedStates[{arrayIndex}], Array length={flaggedArray.Length}");
                            
                            if (arrayIndex >= 0 && arrayIndex < flaggedArray.Length)
                            {
                                isFlagged = flaggedArray[arrayIndex];
                                System.Diagnostics.Debug.WriteLine($"[ReviewPage] flaggedStates[{arrayIndex}] = {isFlagged}");
                                
                                // デバッグ: 配列の全状態を表示
                                for (int i = 0; i < flaggedArray.Length; i++)
                                {
                                    if (flaggedArray[i])
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[ReviewPage] flaggedStates[{i}] = true (corresponds to TaskId {i + 1})");
                                    }
                                }
                            }
                        }
                        
                        System.Diagnostics.Debug.WriteLine($"[ReviewPage] プロジェクト{task.ProjectId}タスク{task.TaskId}: 解答済み={isCompleted}, フラグ={isFlagged}");
                        
                        // UI要素を更新
                        UpdateTaskUIStates(task.ProjectId, task.TaskId, isCompleted, isFlagged);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"状態更新エラー: {ex.Message}");
            }
        }
        
        private void UpdateTaskUIStates(int projectId, int taskId, bool isCompleted, bool isFlagged)
        {
            try
            {
                // プロジェクト内のタスクを検索
                foreach (var project in ProjectsItemsControl.ItemsSource as List<ReviewProjectInfo>)
                {
                    if (project.ProjectTitle == $"プロジェクト {projectId}")
                    {
                        var targetTask = project.Tasks.FirstOrDefault(t => t.TaskId == taskId);
                        if (targetTask != null)
                        {
                            // タスクのUI要素を検索して更新
                            var taskContainer = FindTaskContainer(projectId, taskId);
                            if (taskContainer != null)
                            {
                                UpdateTaskMarks(taskContainer, isCompleted, isFlagged);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UI状態更新エラー: {ex.Message}");
            }
        }
        
        private FrameworkElement FindTaskContainer(int projectId, int taskId)
        {
            try
            {
                // ItemsControl内のタスクコンテナを検索
                var itemsControl = ProjectsItemsControl;
                if (itemsControl != null)
                {
                    // プロジェクトのコンテナを検索
                    for (int i = 0; i < itemsControl.Items.Count; i++)
                    {
                        var projectContainer = itemsControl.ItemContainerGenerator.ContainerFromIndex(i) as FrameworkElement;
                        if (projectContainer != null)
                        {
                            // プロジェクト内のタスクコンテナを検索
                            var taskItemsControl = FindVisualChild<ItemsControl>(projectContainer);
                            if (taskItemsControl != null)
                            {
                                for (int j = 0; j < taskItemsControl.Items.Count; j++)
                                {
                                    var taskContainer = taskItemsControl.ItemContainerGenerator.ContainerFromIndex(j) as FrameworkElement;
                                    if (taskContainer != null)
                                    {
                                        var taskInfo = taskContainer.DataContext as ReviewTaskInfo;
                                        if (taskInfo != null && taskInfo.ProjectId == projectId && taskInfo.TaskId == taskId)
                                        {
                                            return taskContainer;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"タスクコンテナ検索エラー: {ex.Message}");
            }
            return null;
        }
        
        private void UpdateTaskMarks(FrameworkElement taskContainer, bool isCompleted, bool isFlagged)
        {
            try
            {
                // 解答済みマークを更新
                var completedMark = FindVisualChild<TextBlock>(taskContainer, "CompletedMark");
                if (completedMark != null)
                {
                    completedMark.Visibility = isCompleted ? Visibility.Visible : Visibility.Collapsed;
                }
                
                // あとで見直すマークを更新
                var flaggedMark = FindVisualChild<TextBlock>(taskContainer, "FlaggedMark");
                if (flaggedMark != null)
                {
                    flaggedMark.Visibility = isFlagged ? Visibility.Visible : Visibility.Collapsed;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"マーク更新エラー: {ex.Message}");
            }
        }
        
        private T FindVisualChild<T>(DependencyObject parent, string name = null) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T target)
                {
                    if (string.IsNullOrEmpty(name) || (child is FrameworkElement fe && fe.Name == name))
                    {
                        return target;
                    }
                }
                
                var childOfChild = FindVisualChild<T>(child, name);
                if (childOfChild != null)
                {
                    return childOfChild;
                }
            }
            return null;
        }
        
        private async void NavigateToTask(ReviewTaskInfo taskInfo)
        {
            if (_isScoring)
            {
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Navigation ignored while scoring.");
                return;
            }

            System.Diagnostics.Debug.WriteLine($"NavigateToTask called: ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");

            if (OnNavigateToTask == null || taskInfo.ProjectId <= 0 || taskInfo.TaskId <= 0)
            {
                System.Diagnostics.Debug.WriteLine($"ナビゲーション条件不一致: OnNavigateToTask={OnNavigateToTask != null}, ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");
                return;
            }

            // クリック直後にオーバーレイを表示（Excel 終了待機より前）
            var openingOverlay = new Window
            {
                Title = "タスクを開いています",
                Width = 320,
                Height = 110,
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
                Text = $"タスクを開いています...\nプロジェクト {taskInfo.ProjectId} - タスク {taskInfo.TaskId}",
                FontSize = 13,
                TextAlignment = System.Windows.TextAlignment.Center,
                Foreground = System.Windows.Media.Brushes.SteelBlue
            });
            openingOverlay.Content = stack;
            openingOverlay.Show();
            // 描画を確定させてからバックグラウンド処理へ
            await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);

            try
            {
                System.Diagnostics.Debug.WriteLine("ナビゲーション実行開始");

                // Excel のバックグラウンド終了処理が走っている場合は完了を待つ（競合によるクラッシュを防止）
                await WaitForPendingExcelCloseAsync("[NavigateToTask]");

                // タイマーを停止
                _timer?.Stop();

                // ナビゲーションを実行
                OnNavigateToTask(taskInfo.GroupId > 0 ? taskInfo.GroupId : _groupId, taskInfo.ProjectId, taskInfo.TaskId);

                openingOverlay.Close();
                System.Diagnostics.Debug.WriteLine("ナビゲーション実行完了");

                // レビューページを閉じる
                this.Close();
            }
            catch (Exception ex)
            {
                openingOverlay.Close();
                System.Diagnostics.Debug.WriteLine($"ナビゲーション実行エラー: {ex.Message}");
            }
        }
        
        private void TaskButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("TaskButton_Click called");

            if (_isScoring)
            {
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Task button ignored while scoring.");
                return;
            }
            
            if (sender is Button button && button.DataContext is ReviewTaskInfo taskInfo)
            {
                System.Diagnostics.Debug.WriteLine($"TaskButton clicked: ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");
                NavigateToTask(taskInfo);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("TaskButton_Click: Invalid sender or DataContext");
            }
        }
        
        private void TaskButton_MouseEnter(object sender, System.Windows.Input.MouseEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Mouse entered task button");
        }
        
        private void TaskButton_MouseLeave(object sender, System.Windows.Input.MouseEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Mouse left task button");
        }
        
        private List<int> GetAvailableProjectIds()
        {
            var projects = ProjectsItemsControl.ItemsSource as IEnumerable<ReviewProjectInfo>;
            if (projects == null)
                return new List<int>();

            return projects
                .SelectMany(p => p.Tasks ?? Enumerable.Empty<ReviewTaskInfo>())
                .Select(t => t.ProjectId)
                .Where(id => id > 0)
                .Distinct()
                .OrderBy(id => id)
                .ToList();
        }

        private async void EndExamButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isScoring)
            {
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Duplicate scoring request ignored.");
                return;
            }

            HashSet<int> scoringProjectIds = null;
            string scoringRangeLabel = "すべてのプロジェクト";
            var availableIds = GetAvailableProjectIds();
            if (availableIds.Count > 0)
            {
                if (!MosPracticeClient.ScoringRangeDialog.TrySelect(this, availableIds, out var selectedIds, out var rangeLabel))
                    return;
                scoringProjectIds = new HashSet<int>(selectedIds);
                if (!string.IsNullOrWhiteSpace(rangeLabel))
                    scoringRangeLabel = rangeLabel;
            }

            _isScoring = true;
            Window scoringOverlay = null;
            DispatcherTimer overlayTopmostTimer = null;
            try
            {
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] EndExamButton_Click called");
                
                // ボタンを無効化して再クリックを防止
                if (sender is Button button)
                {
                    button.IsEnabled = false;
                    button.Content = "処理中...";
                }
                
                // タイマーを停止
                _timer?.Stop();

                // 採点中にタスク番号を操作できないよう、採点開始時点でレビュー画面を隠す
                this.Hide();
                
                // UI更新の機会を与える
                await Task.Delay(100);

                // 「採点の準備をしています...」オーバーレイを表示（採点開始前の接続待ち）
                await Dispatcher.InvokeAsync(() =>
                {
                    scoringOverlay = new Window
                    {
                        Title = "採点中",
                        Width = 320,
                        Height = 140,
                        WindowStyle = WindowStyle.None,
                        WindowStartupLocation = WindowStartupLocation.CenterScreen,
                        ShowInTaskbar = false,
                        ResizeMode = ResizeMode.NoResize,
                        Topmost = true,
                        ShowActivated = true,
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
                    var text = new TextBlock
                    {
                        Text = "採点の準備をしています...",
                        FontSize = 18,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(0, 0, 0, 12),
                        Foreground = new SolidColorBrush(Color.FromRgb(30, 64, 175))
                    };
                    var progress = new System.Windows.Controls.ProgressBar
                    {
                        IsIndeterminate = false,
                        Minimum = 0,
                        Maximum = 1,
                        Value = 0,
                        Height = 20,
                        Width = 260
                    };
                    stack.Children.Add(text);
                    stack.Children.Add(progress);
                    scoringOverlay.Content = stack;
                    scoringOverlay.Tag = new ScoringProgressUi { StatusText = text, ProgressBar = progress };
                    scoringOverlay.Show();
                    try
                    {
                        scoringOverlay.Activate();
                        scoringOverlay.Focus();
                    }
                    catch { }
                });

                // Excel などに前面を奪われることがあるため、短時間だけ最前面を維持する
                overlayTopmostTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(120) };
                int overlayRetryCount = 0;
                overlayTopmostTimer.Tick += (s, args) =>
                {
                    overlayRetryCount++;
                    if (scoringOverlay == null || !scoringOverlay.IsVisible || overlayRetryCount > 20)
                    {
                        overlayTopmostTimer.Stop();
                        return;
                    }
                    try
                    {
                        scoringOverlay.Topmost = true;
                        scoringOverlay.Activate();
                    }
                    catch { }
                };
                overlayTopmostTimer.Start();

                // UIスレッドで描画させるために少し待機
                await Task.Delay(50);

                // バックグラウンドでExcelの終了処理が走っている場合は、完了を待つ
                await WaitForPendingExcelCloseAsync("[ReviewPageWindow]");
                
                try
                {
                    // AppBarWindow（下の問題領域）を閉じる（UIスレッドで実行）
                    await Dispatcher.InvokeAsync(() => CloseAppBarWindows(), DispatcherPriority.Background);
                    
                    // 以降、重い採点処理へ
                    
                    
                    // Excel COM は STA 上で呼ぶ（Task.Run のスレッドプールは MTA になり、GetActiveObject / Workbooks が失敗しうる）
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Starting to score all projects (STA)...");
                    ScoringProgressUi progressUi = scoringOverlay?.Tag as ScoringProgressUi;
                    await RunStaAsync(() =>
                    {
                        var totalSw = Stopwatch.StartNew();
                        ExcelGradingPerf.BeginSession("ScoreAllProjects");
                        try
                        {
                            ScoreAllProjects(scoringProjectIds, (message, completed, total) =>
                            {
                                Dispatcher.BeginInvoke(new Action(() =>
                                {
                                    if (progressUi == null)
                                        return;
                                    progressUi.StatusText.Text = message;
                                    progressUi.ProgressBar.IsIndeterminate = false;
                                    progressUi.ProgressBar.Maximum = Math.Max(total, 1);
                                    progressUi.ProgressBar.Value = Math.Max(0, Math.Min(completed, total));
                                }));
                            });
                            var closeSw = Stopwatch.StartNew();
                            CloseExcelApplication();
                            ExcelGradingPerf.Log("ScoreAllProjects.CloseExcel", closeSw.ElapsedMilliseconds);
                        }
                        finally
                        {
                            ExcelGradingPerf.Log("ScoreAllProjects.Total", totalSw.ElapsedMilliseconds, $"group={_groupId}");
                            ExcelGradingPerf.EndSession();
                        }
                    });
                    
                    // 「採点中です」オーバーレイを閉じる
                    await Dispatcher.InvokeAsync(() =>
                    {
                        try
                        {
                            // 念のためタイマー停止
                            try { overlayTopmostTimer?.Stop(); } catch { }
                            scoringOverlay?.Close();
                        }
                        catch { }
                    });
                    await Task.Delay(100);
                    
                    // 保存されている採点結果を取得
                    var allResults = Models.ExamResultStorage.GetAllResults();
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] All results retrieved: {allResults?.Count ?? 0} projects");
                    if (allResults != null && allResults.Count > 0)
                    {
                        foreach (var result in allResults)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Project {result.Key}: {result.Value.Count} tasks");
                        }
                    }
                    
                    // 採点直後のスナップショットを保存（復習前の正答率用）
                    MOSExcelMogiApp.Models.ExamResultStorage.SaveInitialResultsIfEmpty(allResults);

                    try
                    {
                        if (allResults != null && allResults.Count > 0)
                        {
                            MosPracticeClient.ScoringLogStore.Append(
                                MosPracticeClient.ScoringLogStore.SubjectExcel,
                                MosPracticeClient.ScoringLogEntry.Create(scoringRangeLabel, _groupId, allResults));
                        }
                    }
                    catch (Exception logEx)
                    {
                        System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Scoring log save: " + logEx.Message);
                    }
                    
                    // UI更新の機会を与える
                    await Task.Delay(50);
                    
                    // 結果画面ウィンドウを表示（UIスレッドで実行、非同期で表示）
                    ResultWindow resultWindow = null;
                    var initialPresentationTcs = new TaskCompletionSource<bool>();
                    EventHandler onInitialPresentationCompleted = null;

                    await Dispatcher.InvokeAsync(() =>
                    {
                        resultWindow = new ResultWindow(allResults, _groupId);
                        onInitialPresentationCompleted = (_, __) =>
                        {
                            try
                            {
                                resultWindow.InitialPresentationCompleted -= onInitialPresentationCompleted;
                            }
                            catch { }

                            initialPresentationTcs.TrySetResult(true);
                        };
                        resultWindow.InitialPresentationCompleted += onInitialPresentationCompleted;

                        resultWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                        resultWindow.Topmost = true;
                        
                        // OnNavigateToTaskを設定
                        if (OnNavigateToTask != null)
                        {
                            // ReviewPageWindowのOnNavigateToTaskを再利用
                            resultWindow.OnNavigateToTask = (projectId, taskId) =>
                            {
                                // AppBarWindowを確実に表示
                                var appBarWindow = Application.Current.Windows.OfType<AppBarWindow>().FirstOrDefault();
                                if (appBarWindow != null)
                                {
                                    // 結果画面から来たことを記録
                                    appBarWindow.SetFromResultWindow(true);
                                    appBarWindow.SetResultWindow(resultWindow);
                                    resultWindow.PrepareAppBarForResultRetry(appBarWindow, projectId, taskId);
                                    if (!appBarWindow.IsVisible)
                                    {
                                        appBarWindow.Show();
                                    }
                                    appBarWindow.Activate();

                                    // 重要: 表示中のAppBarWindowに直接ナビゲートする（デリゲートの参照先が古い可能性があるため）
                                    appBarWindow.NavigateToTask(projectId, taskId, _groupId);
                                    return;
                                }

                                // フォールバック: 既存の経路（AppBarWindowが見つからない場合）
                                OnNavigateToTask(_groupId, projectId, taskId);
                            };
                        }
                        else
                        {
                            // AppBarWindowからNavigateToTaskを取得
                            var appBarWindow = Application.Current.Windows.OfType<AppBarWindow>().FirstOrDefault();
                            if (appBarWindow != null)
                            {
                                resultWindow.OnNavigateToTask = (projectId, taskId) =>
                                {
                                    // MainWindowを表示
                                    var mainWindow = Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                                    if (mainWindow == null)
                                    {
                                        mainWindow = new MainWindow();
                                        Application.Current.MainWindow = mainWindow;
                                    }
                                    mainWindow.Show();
                                    mainWindow.WindowState = WindowState.Normal;
                                    mainWindow.Activate();
                                    mainWindow.Focus();
                                    
                                    // AppBarWindowを表示
                                    // 結果画面から来たことを記録
                                    appBarWindow.SetFromResultWindow(true);
                                    appBarWindow.SetResultWindow(resultWindow);
                                    resultWindow.PrepareAppBarForResultRetry(appBarWindow, projectId, taskId);
                                    if (!appBarWindow.IsVisible)
                                    {
                                        appBarWindow.Show();
                                    }
                                    appBarWindow.Activate();
                                    
                                    // NavigateToTaskを呼び出す
                                    appBarWindow.NavigateToTask(projectId, taskId, _groupId);
                                };
                            }
                        }
                        
                        // 結果ウィンドウを表示
                        resultWindow.Show();
                        resultWindow.Activate();
                    }, DispatcherPriority.Normal);

                    // ResultWindow の初回データ表示完了まで待つ（固定 500ms+1000ms より短く終わることが多い）
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Waiting for ResultWindow initial presentation...");
                    const int resultWindowInitialTimeoutMs = 8000;
                    Task timeoutTask = Task.Delay(resultWindowInitialTimeoutMs);
                    Task completed = await Task.WhenAny(initialPresentationTcs.Task, timeoutTask).ConfigureAwait(true);
                    if (completed != initialPresentationTcs.Task)
                    {
                        System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] ResultWindow initial presentation timed out; proceeding.");
                        await Dispatcher.InvokeAsync(() =>
                        {
                            try
                            {
                                if (resultWindow != null && onInitialPresentationCompleted != null)
                                    resultWindow.InitialPresentationCompleted -= onInitialPresentationCompleted;
                            }
                            catch { }
                        });
                    }

                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] ResultWindow initial presentation gate passed");
                    
                    // Application.Current.MainWindowをResultWindowへ切り替える
                    await Dispatcher.InvokeAsync(() =>
                    {
                        var resultWindow = Application.Current.Windows.OfType<ResultWindow>().FirstOrDefault();
                        if (resultWindow != null)
                        {
                            Application.Current.MainWindow = resultWindow;
                            System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Set Application.Current.MainWindow to ResultWindow");
                        }
                    }, DispatcherPriority.Normal);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in EndExamButton_Click: {ex.Message}\n{ex.StackTrace}");
                    await Dispatcher.InvokeAsync(() =>
                    {
                        try { overlayTopmostTimer?.Stop(); } catch { }
                        try { scoringOverlay?.Close(); } catch { }
                        RestoreAfterScoringFailure(sender);
                        MessageBox.Show($"試験終了処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in EndExamButton_Click: {ex.Message}\n{ex.StackTrace}");
                await Dispatcher.InvokeAsync(() =>
                {
                    try { overlayTopmostTimer?.Stop(); } catch { }
                    try { scoringOverlay?.Close(); } catch { }
                    RestoreAfterScoringFailure(sender);
                    MessageBox.Show($"試験終了処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private void RestoreAfterScoringFailure(object sender)
        {
            _isScoring = false;
            if (sender is Button button)
            {
                button.IsEnabled = true;
                button.Content = "結果の表示";
            }

            this.Show();
            this.Activate();
            if (!MainWindow.IsTimerDisabled)
                _timer?.Start();
        }

        /// <summary>Excel COM 用に専用 STA スレッドで処理を実行する（MTA からの呼び出しは不安定）。</summary>
        private static Task RunStaAsync(Action action)
        {
            var tcs = new TaskCompletionSource<bool>();
            var thread = new Thread(() =>
            {
                try
                {
                    action();
                    tcs.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            })
            {
                IsBackground = true,
                Name = "MOS_ExcelScoring_STA"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return tcs.Task;
        }

        private static string NormalizeExcelPath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return string.Empty;

            try
            {
                return Path.GetFullPath(filePath).ToLowerInvariant();
            }
            catch
            {
                return filePath.Trim().ToLowerInvariant();
            }
        }

        private bool IsWorkbookOpenInScoringSession(string filePath)
        {
            var excelApp = _scoringExcelApp;
            if (excelApp == null || string.IsNullOrEmpty(filePath))
                return false;

            var normalizedTargetPath = NormalizeExcelPath(filePath);
            try
            {
                if (excelApp.Workbooks == null)
                    return false;

                foreach (ExcelWorkbook wb in excelApp.Workbooks)
                {
                    try
                    {
                        if (NormalizeExcelPath(wb.FullName) == normalizedTargetPath)
                            return true;
                    }
                    catch { }
                    finally
                    {
                        try { Marshal.ReleaseComObject(wb); } catch { }
                    }
                }
            }
            catch { }

            return false;
        }

        /// <summary>
        /// 採点セッションの ActiveWorkbook が期待したパスと一致しているかを軽量に確認する。
        /// 一致している場合は再アクティブ化を省略できる。
        /// </summary>
        private bool IsExpectedWorkbookAlreadyActive(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            var excelApp = _scoringExcelApp;
            if (excelApp == null)
                return false;

            try
            {
                var active = excelApp.ActiveWorkbook;
                if (active == null || string.IsNullOrEmpty(active.FullName))
                    return false;

                return NormalizeExcelPath(active.FullName) == NormalizeExcelPath(filePath);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>採点中に保持している Application を返す（採点中にROTへ戻らない）。</summary>
        private ExcelApp GetExcelApplicationForScoringOptional(out bool releaseWhenDone)
        {
            releaseWhenDone = false;
            return _scoringExcelApp;
        }

        /// <summary>
        /// 採点専用 Excel.Application を取得する（採点中にROT再取得しない）。
        /// </summary>
        private ExcelApp EnsureScoringExcelApplication(bool makeVisible, int timeoutMs, string caller)
        {
            if (_scoringExcelApp != null)
            {
                try
                {
                    // stale COM proxy（Excel再起動後の死んだ参照）検知
                    _ = _scoringExcelApp.Hwnd;
                    try { _scoringExcelApp.Visible = makeVisible; } catch { }
                    return _scoringExcelApp;
                }
                catch (COMException)
                {
                    try { Marshal.ReleaseComObject(_scoringExcelApp); } catch { }
                    _scoringExcelApp = null;
                    _scoringExcelCreatedBySession = false;
                    _scoringExcelCreatedPid = -1;
                }
                catch (Exception)
                {
                    try { Marshal.ReleaseComObject(_scoringExcelApp); } catch { }
                    _scoringExcelApp = null;
                    _scoringExcelCreatedBySession = false;
                    _scoringExcelCreatedPid = -1;
                }
            }

            try
            {
                _scoringExcelApp = new ExcelApp();
                _scoringExcelCreatedBySession = true;
                _scoringExcelCreatedPid = ExcelApplicationManager.TryGetExcelProcessId(_scoringExcelApp);
                try { _scoringExcelApp.DisplayAlerts = false; } catch { }
                try { ApplyScoringWindowLayout(_scoringExcelApp); } catch { }
                try { _scoringExcelApp.Visible = makeVisible; } catch { }
                return _scoringExcelApp;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[{caller}] Dedicated Excel creation failed: {ex.Message}");
                return null;
            }
        }

        private void TryReplaceWithWorkbookBackedExcel(bool makeVisible, int timeoutMs, string caller, string logMessage)
        {
            if (_scoringExcelApp == null) return;

            int currentWbCount = 0;
            try
            {
                currentWbCount = _scoringExcelApp.Workbooks?.Count ?? 0;
            }
            catch
            {
                currentWbCount = 0;
            }
            if (currentWbCount > 0) return;

            try
            {
                var reattached = ExcelApplicationManager.TryAttachRunningExcelApplication(
                    makeVisible: makeVisible,
                    timeoutMs: Math.Min(timeoutMs, 8000));
                if (reattached == null) return;

                int reattachWbCount = 0;
                int reattachHwnd = 0;
                try
                {
                    reattachWbCount = reattached.Workbooks?.Count ?? 0;
                    reattachHwnd = reattached.Hwnd;
                }
                catch
                {
                    reattachWbCount = 0;
                    reattachHwnd = 0;
                }


                if (reattachWbCount <= 0)
                {
                    try { Marshal.ReleaseComObject(reattached); } catch { }
                    return;
                }

                if (!object.ReferenceEquals(_scoringExcelApp, reattached))
                {
                    try { Marshal.ReleaseComObject(_scoringExcelApp); } catch { }
                }
                _scoringExcelApp = reattached;
                _scoringExcelCreatedBySession = false;
                _scoringExcelCreatedPid = -1;
            }
            catch
            {
                // keep current _scoringExcelApp
            }
        }
        
        private sealed class ScoringProgressUi
        {
            public TextBlock StatusText;
            public System.Windows.Controls.ProgressBar ProgressBar;
        }

        private void ScoreAllProjects(ISet<int> projectIds = null, Action<string, int, int> progress = null)
        {
            try
            {
                // ExamResultStorageをクリア
                Models.ExamResultStorage.Clear();
                _scoringExcelApp = null;
                _scoringExcelCreatedBySession = false;
                _scoringExcelCreatedPid = -1;

                // 採点時間を短縮するため、ログファイルを一括読み込みしてキャッシュする
                string logPath = ExcelLogReader.GetLogFilePath();
                if (System.IO.File.Exists(logPath))
                {
                    try
                    {
                        ExcelLogReader.SetLogLinesCache(System.IO.File.ReadAllLines(logPath));
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ScoreAllProjects] Failed to cache log: {ex.Message}");
                    }
                }
                
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

                if (projectIds != null && projectIds.Count > 0)
                    projectList = projectList.Where(p => projectIds.Contains(p.projectId)).ToList();
                
                // ステップ1: 既に開いているファイルを確認し、必要に応じて開く
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Step 1: Checking Excel files...");
                var connectSw = Stopwatch.StartNew();
                EnsureScoringExcelApplication(
                    makeVisible: true,
                    timeoutMs: 15000,
                    caller: "ReviewPageWindow.ScoreAllProjects");
                ExcelGradingPerf.Log(
                    "ScoreAllProjects.Connect",
                    connectSw.ElapsedMilliseconds,
                    _scoringExcelApp == null ? "missing" : "ok");
                
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Scoring one workbook at a time");
                for (int idx = 0; idx < projectList.Count; idx++)
                {
                    var project = projectList[idx];
                    progress?.Invoke(
                        $"プロジェクト{project.projectId}を採点中（{idx + 1}/{projectList.Count}）",
                        idx,
                        projectList.Count);

                    if (string.IsNullOrEmpty(project.filePath) || !File.Exists(project.filePath) || _scoringExcelApp == null)
                    {
                        SaveAllFalse(project.projectId, project.taskCount);
                        progress?.Invoke(
                            $"プロジェクト{project.projectId}を採点しました（{idx + 1}/{projectList.Count}）",
                            idx + 1,
                            projectList.Count);
                        continue;
                    }

                    bool opened = false;
                    try
                    {
                        var projectSw = Stopwatch.StartNew();
                        opened = OpenProjectWorkbook(project.filePath);
                        if (!opened)
                        {
                            SaveAllFalse(project.projectId, project.taskCount);
                            continue;
                        }

                        var results = ExecuteScoringForProject(
                            project.libraryName,
                            project.taskCount,
                            project.filePath,
                            project.projectId,
                            out List<string> failReasons);
                        Models.ExamResultStorage.SaveProjectResult(project.projectId, results, failReasons);
                        ExcelGradingPerf.Log(
                            "ScoreAllProjects.Project",
                            projectSw.ElapsedMilliseconds,
                            $"P{project.projectId}");
                        Application.Current.Dispatcher.BeginInvoke(new Action(() => { }), DispatcherPriority.Background);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error scoring project {project.projectId}: {ex.Message}");
                        SaveAllFalse(project.projectId, project.taskCount);
                    }
                    finally
                    {
                        if (opened)
                            CloseProjectWorkbook(project.filePath);
                        progress?.Invoke(
                            $"プロジェクト{project.projectId}を採点しました（{idx + 1}/{projectList.Count}）",
                            idx + 1,
                            projectList.Count);
                    }
                }
                
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] All projects scored successfully");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in ScoreAllProjects: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
            }
            finally
            {
                ReleaseScoringWorkbookIndex();
                ExcelLogReader.ClearLogLinesCache();
            }
        }

        private static void SaveAllFalse(int projectId, int taskCount)
        {
            var falseResults = new List<bool>();
            var failReasons = new List<string>();
            for (int i = 0; i < taskCount; i++)
            {
                falseResults.Add(false);
                failReasons.Add(ExcelScoreExplanation.UnavailableText);
            }
            Models.ExamResultStorage.SaveProjectResult(projectId, falseResults, failReasons);
        }

        private bool OpenProjectWorkbook(string filePath)
        {
            CloseOtherScoringWorkbooks(filePath);
            if (_scoringExcelApp == null || string.IsNullOrEmpty(filePath))
                return false;

            bool alreadyOpen = IsWorkbookReady(filePath);
            var openSw = Stopwatch.StartNew();
            if (!alreadyOpen)
            {
                try
                {
                    _scoringExcelApp.Workbooks.Open(
                        filePath,
                        UpdateLinks: false,
                        ReadOnly: false,
                        Format: Type.Missing,
                        Password: Type.Missing,
                        WriteResPassword: Type.Missing,
                        IgnoreReadOnlyRecommended: true,
                        Origin: Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                        Delimiter: Type.Missing,
                        Editable: true,
                        Notify: false,
                        Converter: Type.Missing,
                        AddToMru: false,
                        Local: false,
                        CorruptLoad: Microsoft.Office.Interop.Excel.XlCorruptLoad.xlNormalLoad);
                }
                catch (Exception ex)
                {
                    ExcelGradingPerf.Log("OpenProjectWorkbook.Open", openSw.ElapsedMilliseconds, "failed");
                    System.Diagnostics.Debug.WriteLine("[OpenProjectWorkbook] " + ex.Message);
                    return false;
                }
            }

            ExcelGradingPerf.Log(
                "OpenProjectWorkbook.Open",
                alreadyOpen ? 0 : openSw.ElapsedMilliseconds,
                alreadyOpen ? "already-open" : Path.GetFileName(filePath));

            var readySw = Stopwatch.StartNew();
            bool ready = WaitUntilWorkbookReady(filePath, 3000);
            RebuildScoringWorkbookIndex();
            bool active = ready && TryActivateProjectWorkbook(filePath);
            ExcelGradingPerf.Log(
                "OpenProjectWorkbook.Ready",
                readySw.ElapsedMilliseconds,
                active ? "ready" : "timeout");
            return active;
        }

        private void CloseOtherScoringWorkbooks(string keepPath)
        {
            if (_scoringExcelApp?.Workbooks == null)
                return;

            string keep = NormalizeExcelPath(keepPath);
            var closing = new List<ExcelWorkbook>();
            foreach (ExcelWorkbook wb in _scoringExcelApp.Workbooks)
            {
                try
                {
                    if (NormalizeExcelPath(wb.FullName) != keep)
                        closing.Add(wb);
                }
                catch { }
            }

            foreach (ExcelWorkbook wb in closing)
                CloseWorkbookInstance(wb);
            if (closing.Count > 0)
                RebuildScoringWorkbookIndex();
        }

        private void CloseProjectWorkbook(string filePath)
        {
            var sw = Stopwatch.StartNew();
            bool closed = false;
            try
            {
                string key = NormalizeExcelPath(filePath);
                if (!_scoringWorkbooks.TryGetValue(key, out ExcelWorkbook workbook) || workbook == null)
                    workbook = FindOpenWorkbook(filePath);
                if (workbook != null)
                {
                    CloseWorkbookInstance(workbook);
                    closed = true;
                }
                _scoringWorkbooks.Remove(key);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[CloseProjectWorkbook] " + ex.Message);
            }
            finally
            {
                ExcelGradingPerf.Log(
                    "CloseProjectWorkbook.Close",
                    sw.ElapsedMilliseconds,
                    closed ? "closed" : "missing");
            }
        }

        private ExcelWorkbook FindOpenWorkbook(string filePath)
        {
            if (_scoringExcelApp?.Workbooks == null || string.IsNullOrEmpty(filePath))
                return null;

            string target = NormalizeExcelPath(filePath);
            foreach (ExcelWorkbook wb in _scoringExcelApp.Workbooks)
            {
                try
                {
                    if (NormalizeExcelPath(wb.FullName) == target)
                        return wb;
                }
                catch { }
            }
            return null;
        }

        private static void CloseWorkbookInstance(ExcelWorkbook workbook)
        {
            if (workbook == null)
                return;
            try { workbook.Saved = true; } catch { }
            try { workbook.Close(SaveChanges: false); } catch { }
            try { Marshal.ReleaseComObject(workbook); } catch { }
        }

        private bool EnsureProjectWorkbooksReady(List<string> filesToOpen)
        {
            if (filesToOpen == null || filesToOpen.Count == 0) return true;

            const int maxAttempts = 3;
            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                var pending = filesToOpen.Where(path => !IsWorkbookReady(path)).ToList();
                if (pending.Count == 0)
                    return true;

                OpenExcelFilesInBackground(pending);

                bool allReady = true;
                foreach (var filePath in filesToOpen)
                {
                    if (!IsWorkbookReady(filePath))
                    {
                        allReady = false;
                        break;
                    }
                }

                if (allReady) return true;
            }

            return false;
        }

        private bool IsWorkbookReady(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            string target = NormalizeExcelPath(filePath);
            if (_scoringWorkbooks.TryGetValue(target, out ExcelWorkbook cached) && cached != null)
            {
                try
                {
                    return NormalizeExcelPath(cached.FullName) == target
                        && cached.Worksheets != null
                        && cached.Worksheets.Count >= 0;
                }
                catch
                {
                    return false;
                }
            }

            if (_scoringExcelApp?.Workbooks == null)
                return false;

            try
            {
                foreach (ExcelWorkbook wb in _scoringExcelApp.Workbooks)
                {
                    bool matched = false;
                    try
                    {
                        matched = NormalizeExcelPath(wb.FullName) == target
                            && wb.Worksheets != null
                            && wb.Worksheets.Count >= 0;
                    }
                    catch { }
                    finally
                    {
                        if (_scoringWorkbooks.Count == 0)
                        {
                            try { Marshal.ReleaseComObject(wb); } catch { }
                        }
                    }
                    if (matched)
                        return true;
                }
            }
            catch { }

            return false;
        }

        private bool WaitUntilWorkbookReady(string filePath, int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (IsWorkbookReady(filePath))
                    return true;
                Thread.Sleep(30);
            }

            return IsWorkbookReady(filePath);
        }

        private void RebuildScoringWorkbookIndex()
        {
            ReleaseScoringWorkbookIndex();
            if (_scoringExcelApp?.Workbooks == null)
                return;

            foreach (ExcelWorkbook wb in _scoringExcelApp.Workbooks)
            {
                try
                {
                    string fullName = NormalizeExcelPath(wb.FullName);
                    if (string.IsNullOrEmpty(fullName) || wb.Worksheets == null || wb.Worksheets.Count < 0)
                        continue;
                    _scoringWorkbooks[fullName] = wb;
                }
                catch { }
            }
        }

        private void ReleaseScoringWorkbookIndex()
        {
            _scoringWorkbooks.Clear();
        }

        private bool TryActivateProjectWorkbook(string filePath)
        {
            var sw = Stopwatch.StartNew();
            bool activated = false;
            try
            {
                activated = TryActivateProjectWorkbookCore(filePath);
                return activated;
            }
            finally
            {
                ExcelGradingPerf.Log(
                    "ScoreAllProjects.Activate",
                    sw.ElapsedMilliseconds,
                    activated ? "ok" : "failed");
            }
        }

        private bool TryActivateProjectWorkbookCore(string filePath)
        {
            if (TryActivateFromSession(filePath))
                return true;

            try
            {
                EnsureScoringExcelApplication(
                    makeVisible: true,
                    timeoutMs: 15000,
                    caller: "ReviewPageWindow.TryActivateProjectWorkbook");
                RebuildScoringWorkbookIndex();
            }
            catch { }

            if (!IsWorkbookReady(filePath))
                OpenExcelFilesInBackground(new List<string> { filePath });

            RebuildScoringWorkbookIndex();
            return TryActivateFromSession(filePath);
        }

        private bool TryActivateFromSession(string filePath)
        {
            if (IsExpectedWorkbookAlreadyActive(filePath))
                return true;

            string key = NormalizeExcelPath(filePath);
            if (!string.IsNullOrEmpty(key)
                && _scoringWorkbooks.TryGetValue(key, out ExcelWorkbook workbook)
                && workbook != null
                && _scoringExcelApp != null)
            {
                try
                {
                    TryActivateWorkbookForScoring(_scoringExcelApp, workbook);
                    if (WaitUntilActiveWorkbook(filePath, 1500))
                        return true;
                }
                catch (COMException)
                {
                    return ReconnectAndActivate(filePath);
                }
                catch { }
            }

            return ActivateExcelFile(filePath);
        }

        private bool ReconnectAndActivate(string filePath)
        {
            try { if (_scoringExcelApp != null) Marshal.ReleaseComObject(_scoringExcelApp); } catch { }
            _scoringExcelApp = null;
            ReleaseScoringWorkbookIndex();
            try
            {
                EnsureScoringExcelApplication(
                    makeVisible: true,
                    timeoutMs: 15000,
                    caller: "ReviewPageWindow.ReconnectAndActivate");
            }
            catch
            {
                return false;
            }

            OpenExcelFilesInBackground(new List<string> { filePath });
            RebuildScoringWorkbookIndex();
            return ActivateExcelFile(filePath);
        }

        private bool WaitUntilActiveWorkbook(string filePath, int timeoutMs)
        {
            string target = NormalizeExcelPath(filePath);
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                if (IsExpectedWorkbookAlreadyActive(filePath))
                    return true;
                Thread.Sleep(30);
            }

            try
            {
                var active = _scoringExcelApp?.ActiveWorkbook;
                if (active != null && string.Equals(active.Name, Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { }

            return NormalizeExcelPath(filePath) == target && IsExpectedWorkbookAlreadyActive(filePath);
        }

        /// <summary>1タスク分の CheckTask / CheckTask_Impl の解決結果。タスクループ内の GetMethod 繰り返しを避ける。</summary>
        private sealed class CheckTaskMethodBinding
        {
            public MethodInfo ImplMethod;
            /// <summary>private bool CheckTask_* (string filePath) など、パス引き版。</summary>
            public MethodInfo FilePathMethod;
            public MethodInfo PublicMethod;
            public string ResolvedMethodName;

            public static string[] GetMethodNameCandidates(string groupId, string projectId, int taskIndex)
            {
                if (groupId == "1")
                {
                    return new[]
                    {
                        $"CheckTask_1_{projectId}_{taskIndex:D2}",
                        $"CheckTask_1_{projectId}_0{taskIndex}",
                        $"CheckTask_{groupId}_{projectId}_{taskIndex:D2}",
                        $"CheckTask_{groupId}_{projectId}_0{taskIndex}"
                    };
                }

                return new[]
                {
                    $"CheckTask_{groupId}_{projectId}_{taskIndex:D2}",
                    $"CheckTask_{groupId}_{projectId}_0{taskIndex}",
                    $"CheckTask_1_{projectId}_{taskIndex:D2}",
                    $"CheckTask_1_{projectId}_0{taskIndex}"
                };
            }

            public static CheckTaskMethodBinding TryResolve(Type checkerType, string groupId, string projectId, int taskIndex)
            {
                string[] methodNames = GetMethodNameCandidates(groupId, projectId, taskIndex);

                foreach (string methodName in methodNames)
                {
                    MethodInfo implMethod = checkerType.GetMethod(
                        methodName + "_Impl",
                        BindingFlags.Instance | BindingFlags.NonPublic);
                    MethodInfo filePathMethod = checkerType.GetMethod(
                        methodName,
                        BindingFlags.Instance | BindingFlags.NonPublic,
                        null,
                        new[] { typeof(string) },
                        null);
                    MethodInfo method = checkerType.GetMethod(methodName);
                    if (implMethod == null && filePathMethod == null && method == null)
                        continue;

                    return new CheckTaskMethodBinding
                    {
                        ImplMethod = implMethod,
                        FilePathMethod = filePathMethod,
                        PublicMethod = method,
                        ResolvedMethodName = methodName
                    };
                }

                return null;
            }

            /// <summary>既存ロジックと同じ分岐でチェッカーを呼び出す。パスがあれば ActiveWorkbook ではなく指定ファイルを採点する。</summary>
            public bool TryInvoke(object checkerInstance, string expectedFilePath, out bool result)
            {
                result = false;
                if (ImplMethod != null && !string.IsNullOrEmpty(expectedFilePath))
                {
                    result = (bool)ImplMethod.Invoke(checkerInstance, new object[] { expectedFilePath });
                    return true;
                }

                if (FilePathMethod != null && !string.IsNullOrEmpty(expectedFilePath))
                {
                    result = (bool)FilePathMethod.Invoke(checkerInstance, new object[] { expectedFilePath });
                    return true;
                }

                if (PublicMethod != null)
                {
                    result = (bool)PublicMethod.Invoke(checkerInstance, null);
                    return true;
                }

                return false;
            }
        }

        private List<bool> ExecuteScoringForProject(
            string libraryName,
            int taskCount,
            string expectedFilePath,
            int slotProjectId,
            out List<string> failReasons)
        {
            var results = new List<bool>();
            failReasons = new List<string>();
            
            try
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ReviewPageWindow] ExecuteScoringForProject called with libraryName: {libraryName}, taskCount: {taskCount}, slotProjectId: {slotProjectId}");
                
                // Extract group and project numbers from library name
                var parts = libraryName.Replace("ExcelChecker", "").Split('_');
                if (parts.Length >= 2)
                {
                    string groupId = parts[0];
                    string projectId = parts[1];
                    
                    // Create namespace and type name
                    string namespaceName = $"Libraries.Group{groupId}";
                    string fullTypeName = $"{namespaceName}.{libraryName}";
                    
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Looking for type: {fullTypeName}");
                    
                    // すべてのアセンブリから型を検索
                    Type checkerType = null;
                    
                    // 最初にMOSExcelMogiAppアセンブリを優先的にチェック
                    var mainAssembly = AppDomain.CurrentDomain.GetAssemblies()
                        .FirstOrDefault(a => a.GetName().Name == "MOSExcelMogiApp");
                    
                    if (mainAssembly != null)
                    {
                        try
                        {
                            checkerType = mainAssembly.GetType(fullTypeName);
                            
                            // 見つからない場合は、名前で検索
                            if (checkerType == null)
                            {
                                var types = mainAssembly.GetTypes();
                                foreach (var type in types)
                                {
                                    if (type.Name == libraryName)
                                    {
                                        checkerType = type;
                                        break;
                                    }
                                }
                            }
                        }
                        catch (ReflectionTypeLoadException ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] ReflectionTypeLoadException: {ex.Message}");
                            if (ex.LoaderExceptions != null)
                            {
                                foreach (var loaderEx in ex.LoaderExceptions)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Loader exception: {loaderEx?.Message}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error getting type from main assembly: {ex.Message}");
                        }
                    }
                    
                    // まだ見つからない場合は、他のアセンブリもチェック
                    if (checkerType == null)
                    {
                        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                        {
                            if (assembly == mainAssembly) continue;
                            
                            try
                            {
                                checkerType = assembly.GetType(fullTypeName);
                                if (checkerType != null) break;
                                
                                // 名前で検索
                                var types = assembly.GetTypes();
                                foreach (var type in types)
                                {
                                    if (type.Name == libraryName)
                                    {
                                        checkerType = type;
                                        break;
                                    }
                                }
                                if (checkerType != null) break;
                            }
                            catch { }
                        }
                    }
                    
                    // DLLから読み込む（MainViewModelのロジックを追加）
                    if (checkerType == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Type not found in loaded assemblies, trying to load DLL");
                        
                        string[] dllPaths = {
                            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "Debug", "Libraries", $"Group{groupId}", $"{libraryName}.dll"),
                            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", $"Group{groupId}", $"{libraryName}.dll"),
                            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", "bin", "Debug", "net48", $"{libraryName}.dll")
                        };
                        
                        string dllPath = null;
                        foreach (string path in dllPaths)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Trying DLL path: {path}");
                            if (File.Exists(path))
                            {
                                dllPath = path;
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Found DLL at: {dllPath}");
                                break;
                            }
                        }
                        
                        if (dllPath != null)
                        {
                            try
                            {
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Loading DLL from: {dllPath}");
                                Assembly assembly = Assembly.LoadFrom(dllPath);
                                checkerType = assembly.GetTypes().FirstOrDefault(t => t.Name == libraryName);
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Found type from DLL: {checkerType?.Name ?? "null"}");
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error loading DLL: {ex.Message}");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] DLL not found in any of the searched paths");
                        }
                    }
                    
                    if (checkerType != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Type found: {checkerType.Name}");

#if DEBUG
                        // Release では GetMethods 全列挙を避ける（採点時間への影響が大きい）
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Available methods in {checkerType.Name}:");
                        foreach (var method in checkerType.GetMethods())
                        {
                            if (method.Name.StartsWith("CheckTask"))
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] - {method.Name}");
                        }
#endif

                        object checkerInstance = Activator.CreateInstance(checkerType);

                        var taskBindings = new CheckTaskMethodBinding[taskCount];
                        for (int ti = 1; ti <= taskCount; ti++)
                            taskBindings[ti - 1] = CheckTaskMethodBinding.TryResolve(checkerType, groupId, projectId, ti);

                        for (int i = 1; i <= taskCount; i++)
                        {
                            // タスク実行ごとに対象ブックを確認し、別ブックへ戻っていたときだけ索引から戻す。
                            if (!string.IsNullOrEmpty(expectedFilePath)
                                && !IsExpectedWorkbookAlreadyActive(expectedFilePath))
                            {
                                TryActivateFromSession(expectedFilePath);
                            }

                            CheckTaskMethodBinding binding = taskBindings[i - 1];
                            string[] triedNames = CheckTaskMethodBinding.GetMethodNameCandidates(groupId, projectId, i);
                            if (binding == null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] No method found for task {i}, returning false");
                                results.Add(false);
                                failReasons.Add(ExcelScoreExplanation.UnavailableText);
                                continue;
                            }

                            string resolvedName = binding.ResolvedMethodName;
                            string invokeMode =
                                binding.ImplMethod != null && !string.IsNullOrEmpty(expectedFilePath)
                                    ? "impl_with_file_path"
                                    : binding.FilePathMethod != null && !string.IsNullOrEmpty(expectedFilePath)
                                        ? "private_filepath"
                                        : "public_no_args";

                            try
                            {
                                var taskSw = Stopwatch.StartNew();
                                ExcelScoreExplanation.ClearCheckerReason();
                                if (!binding.TryInvoke(checkerInstance, expectedFilePath, out bool invokeResult))
                                {
                                    ExcelGradingPerf.Log(
                                        "GradeTask.CheckerInvoke",
                                        taskSw.ElapsedMilliseconds,
                                        $"P{slotProjectId} T{i} invoke-failed");
                                    ExcelScoreExplanation.ClearCheckerReason();
                                    results.Add(false);
                                    failReasons.Add(ExcelScoreExplanation.UnavailableText);
                                    continue;
                                }

                                string checkerReason = ExcelScoreExplanation.TakeCheckerReason();
                                ExcelGradingPerf.Log(
                                    "GradeTask.CheckerInvoke",
                                    taskSw.ElapsedMilliseconds,
                                    $"P{slotProjectId} T{i} passed={invokeResult}");
                                var gateSw = Stopwatch.StartNew();
                                // ログ・免除設定は画面上のスロット番号（VSTO の [Task N-...]）に合わせる
                                invokeResult = ExcelScoreExplanation.Apply(slotProjectId, i, invokeResult, checkerReason, out string reason);
                                ExcelGradingPerf.Log(
                                    "GradeTask.DestructiveValidation",
                                    gateSw.ElapsedMilliseconds,
                                    $"P{slotProjectId} T{i} passed={invokeResult}");
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Method {resolvedName} result: {invokeResult}");
                                results.Add(invokeResult);
                                failReasons.Add(invokeResult ? "" : (reason ?? ExcelScoreExplanation.RequirementMissText));
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error invoking method {resolvedName}: {ex.Message}");
                                System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
                                results.Add(false);
                                failReasons.Add(ExcelScoreExplanation.UnavailableText);
                            }
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Checker type not found: {fullTypeName}");
                        // 型が見つからない場合はすべてfalse
                        for (int i = 0; i < taskCount; i++)
                        {
                            results.Add(false);
                            failReasons.Add(ExcelScoreExplanation.UnavailableText);
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Invalid library name format: {libraryName}");
                    // 無効な形式の場合はすべてfalse
                    for (int i = 0; i < taskCount; i++)
                    {
                        results.Add(false);
                        failReasons.Add(ExcelScoreExplanation.UnavailableText);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in ExecuteScoringForProject: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
                
                // エラー時はすべてfalse
                while (results.Count < taskCount)
                {
                    results.Add(false);
                    failReasons.Add(ExcelScoreExplanation.UnavailableText);
                }
            }
            
            return results;
        }
        
        private void CloseAppBarWindows()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Closing AppBarWindows");
                
                // すべてのウィンドウを走査してAppBarWindowとUiTestAppBarWindowを閉じる
                foreach (Window window in Application.Current.Windows)
                {
                    if (window is AppBarWindow || window is UiTestAppBarWindow)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Closing window: {window.GetType().Name}");
                        window.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error closing AppBarWindows: {ex.Message}");
            }
        }

        private void CloseExcelApplication()
        {
            ReleaseScoringWorkbookIndex();
            ExcelApp excelApp = _scoringExcelApp;
            bool createdBySession = _scoringExcelCreatedBySession;
            int createdPid = _scoringExcelCreatedPid;
            _scoringExcelApp = null;
            _scoringExcelCreatedBySession = false;
            _scoringExcelCreatedPid = -1;

            if (excelApp == null)
            {
                ExcelGradingPerf.Log("CloseExcel.Quit", 0, "no-app");
                ExcelGradingPerf.Log("CloseExcel.WaitProcessExit", 0, "skipped");
                ExcelGradingPerf.Log("CloseExcel.ForceKill", 0, "skipped");
                return;
            }

            int currentPid = ExcelApplicationManager.TryGetExcelProcessId(excelApp);
            bool quitCalled = false;
            var quitSw = Stopwatch.StartNew();
            try
            {
                try { excelApp.DisplayAlerts = false; } catch { }
                if (excelApp.Workbooks != null && excelApp.Workbooks.Count > 0)
                {
                    var workbooksToClose = new List<ExcelWorkbook>();
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                        workbooksToClose.Add(wb);
                    foreach (ExcelWorkbook wb in workbooksToClose)
                        CloseWorkbookInstance(wb);
                }

                if (createdBySession)
                {
                    excelApp.Quit();
                    quitCalled = true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[CloseExcelApplication] " + ex.Message);
            }
            finally
            {
                if (quitCalled)
                {
                    // Workbooks / Application の同期 Release は Quit 後にブロックし得る。
                    ExcelApplicationManager.AbandonComObjectAfterQuit(excelApp, "[CloseExcelApplication]");
                }
                else
                {
                    try
                    {
                        if (excelApp.Workbooks != null)
                            Marshal.ReleaseComObject(excelApp.Workbooks);
                    }
                    catch { }
                    try { Marshal.ReleaseComObject(excelApp); } catch { }
                }

                ExcelGradingPerf.Log("CloseExcel.Quit", quitSw.ElapsedMilliseconds, quitCalled ? "quit" : "released");
            }

            const int normalExitWaitMs = 1500;
            var waitSw = Stopwatch.StartNew();
            bool exited = currentPid <= 0 || !quitCalled
                || ExcelApplicationManager.WaitForExcelProcessExit(currentPid, normalExitWaitMs);
            ExcelGradingPerf.Log(
                "CloseExcel.WaitProcessExit",
                waitSw.ElapsedMilliseconds,
                exited ? "exited" : "running");

            bool mayForceKill = createdBySession
                && quitCalled
                && createdPid > 0
                && currentPid == createdPid
                && !exited;
            var killSw = Stopwatch.StartNew();
            if (mayForceKill)
            {
                ExcelApplicationManager.EnsureExcelProcessExited(
                    currentPid,
                    0,
                    800,
                    "[CloseExcelApplication]");
            }
            ExcelGradingPerf.Log(
                "CloseExcel.ForceKill",
                killSw.ElapsedMilliseconds,
                mayForceKill ? "killed" : "skipped");
        }
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        
        // プロジェクトファイルのパスを取得するメソッド
        private string GetProjectFilePath(int groupId, int projectId, JToken projectConfig)
        {
            try
            {
                string excelFile = projectConfig["excelFile"]?.ToString();
                string initialDataFile = projectConfig["initialDataFile"]?.ToString();
                string configuredFallback = !string.IsNullOrWhiteSpace(excelFile)
                    ? excelFile
                    : initialDataFile;
                return DataPathHelper.ResolveWorkingFilePath(groupId, projectId, configuredFallback);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetProjectFilePath] Error: {ex.Message}");
                return DataPathHelper.GetWorkingFilePath(groupId, projectId);
            }
        }
        
        // 現在Excelで開いているファイルの中から、指定されたプロジェクトのファイルを取得
        private string GetOpenExcelFileForProject(int groupId, int projectId)
        {
            ExcelApp excelApp = null;
            try
            {
                try
                {
                    excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    return null;
                }
                
                if (excelApp != null && excelApp.Workbooks != null)
                {
                    string expectedFileName = $"project{projectId}.xlsx";
                    
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            string wbName = wb.Name;
                            string wbFullName = wb.FullName;
                            
                            // ファイル名が一致し、かつTab{groupId}フォルダ内のファイルであることを確認
                            if (wbName.Equals(expectedFileName, StringComparison.OrdinalIgnoreCase) &&
                                wbFullName.Contains($"Tab{groupId}"))
                            {
                                System.Diagnostics.Debug.WriteLine($"[GetOpenExcelFileForProject] Found matching open file: {wbFullName}");
                                return wbFullName;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[GetOpenExcelFileForProject] Error checking workbook: {ex.Message}");
                        }
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetOpenExcelFileForProject] Error: {ex.Message}");
                return null;
            }
            finally
            {
                if (excelApp != null)
                {
                    try { Marshal.ReleaseComObject(excelApp); } catch { }
                }
            }
        }
        
        // Excelファイルが既に開いているかどうかを確認するメソッド
        private bool IsExcelFileOpen(string filePath)
        {
            ExcelApp excelApp = null;
            try
            {
                // 採点中は専用Excelのみを参照し、ROTへ戻らない
                excelApp = _scoringExcelApp;
                if (excelApp != null && excelApp.Workbooks != null)
                {
                    string normalizedTargetPath = NormalizeExcelPath(filePath);
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            // 同名ファイル（project1.xlsx など）を別フォルダから誤認しないよう、フルパス一致のみで判定する。
                            string wbFullName = wb.FullName;
                            if (!string.IsNullOrEmpty(wbFullName))
                            {
                                string normalizedWorkbookPath = NormalizeExcelPath(wbFullName);
                                if (normalizedWorkbookPath == normalizedTargetPath)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Excel file is already open: {wb.Name} ({wbFullName})");
                                    return true;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error checking workbook: {ex.Message}");
                        }
                        finally
                        {
                            try { Marshal.ReleaseComObject(wb); } catch { }
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in IsExcelFileOpen: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// 指定したExcelファイルをCOMで非表示のまま開く（採点用）。
        /// 既にExcelが起動している場合はそのインスタンスを利用し Visible=false にする。
        /// </summary>
        private void OpenExcelFilesInBackground(List<string> filesToOpen)
        {
            if (filesToOpen == null || filesToOpen.Count == 0)
            {
                // すべて既に開いている場合でも、採点と同じ Application 参照を STA 上で掴んでおく
                EnsureScoringExcelApplication(
                    makeVisible: true,
                    timeoutMs: 15000,
                    caller: "ReviewPageWindow.OpenExcelFilesInBackground");
                return;
            }
            
            ExcelApp excelApp = null;
            try
            {
                excelApp = EnsureScoringExcelApplication(
                    makeVisible: true,
                    timeoutMs: 30000,
                    caller: "ReviewPageWindow.OpenExcelFilesInBackground");
                if (excelApp == null)
                {
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] OpenExcelFilesInBackground: failed to acquire Excel application");
                    return;
                }
                
                foreach (string filePath in filesToOpen)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) continue;
                        if (IsWorkbookOpenInScoringSession(filePath))
                        {
                            continue;
                        }
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Opening Excel file (background): {filePath}");
                        var openSw = Stopwatch.StartNew();
                        excelApp.Workbooks.Open(filePath,
                            UpdateLinks: false,
                            ReadOnly: false,
                            Format: Type.Missing,
                            Password: Type.Missing,
                            WriteResPassword: Type.Missing,
                            IgnoreReadOnlyRecommended: true,
                            Origin: Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                            Delimiter: Type.Missing,
                            Editable: true,
                            Notify: false,
                            Converter: Type.Missing,
                            AddToMru: false,
                            Local: false,
                            CorruptLoad: Microsoft.Office.Interop.Excel.XlCorruptLoad.xlNormalLoad);
                        ExcelGradingPerf.Log("OpenProjectWorkbook.Open", openSw.ElapsedMilliseconds, Path.GetFileName(filePath));
                        var readySw = Stopwatch.StartNew();
                        bool ready = WaitUntilWorkbookReady(filePath, 3000);
                        ExcelGradingPerf.Log(
                            "OpenProjectWorkbook.Ready",
                            readySw.ElapsedMilliseconds,
                            ready ? "ready" : "timeout");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error opening Excel file: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in OpenExcelFilesInBackground: {ex.Message}");
            }
            // excelAppは解放しない（ScoreAllProjectsでGetActiveObjectにより再利用する）
        }

        /// <summary>
        /// 採点用: Excel を表示したうえでブック／ウィンドウをアクティブにする（非表示のままだと ActiveWorkbook が付かないことがある）。
        /// </summary>
        private static void TryActivateWorkbookForScoring(ExcelApp excelApp, ExcelWorkbook wb)
        {
            if (excelApp == null || wb == null) return;
            try
            {
                excelApp.Visible = true;
                // Excel ウィンドウを最前面へ（ユーザーの要望：レビューページの後ろに隠れないようにする）
                IntPtr hwnd = new IntPtr(excelApp.Hwnd);
                if (hwnd != IntPtr.Zero)
                {
                    SetForegroundWindow(hwnd);
                }
            }
            catch { }
            try { wb.Activate(); } catch { }
            Microsoft.Office.Interop.Excel.Windows wins = null;
            Microsoft.Office.Interop.Excel.Window win = null;
            try
            {
                wins = wb.Windows;
                if (wins != null && wins.Count >= 1)
                {
                    win = wins[1];
                    try { win.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal; } catch { }
                    win.Activate();
                }
            }
            catch { }
            finally
            {
                try { if (win != null) Marshal.ReleaseComObject(win); } catch { }
                try { if (wins != null) Marshal.ReleaseComObject(wins); } catch { }
            }
        }

        private static void ApplyScoringWindowLayout(ExcelApp excelApp)
        {
            if (excelApp == null) return;
            try { excelApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal; } catch { }
            try { excelApp.Left = 150; } catch { }
            try { excelApp.Top = 100; } catch { }
            try { excelApp.Width = 900; } catch { }
            try { excelApp.Height = 620; } catch { }
        }
        
        // Excelファイルをアクティブにするメソッド
        private bool ActivateExcelFile(string filePath)
        {
            return ActivateExcelFileInternal(filePath, allowRetryOnComException: true);
        }

        private bool ActivateExcelFileInternal(string filePath, bool allowRetryOnComException)
        {
            ExcelApp excelApp = null;
            try
            {
                excelApp = EnsureScoringExcelApplication(
                    makeVisible: true,
                    timeoutMs: 30000,
                    caller: "ReviewPageWindow.ActivateExcelFile");
                if (excelApp == null)
                    return false;
                
                if (excelApp != null && excelApp.Workbooks != null)
                {
                    // 正規化したパスで比較（大文字小文字を統一）
                    string normalizedFilePath = NormalizeExcelPath(filePath);

                    // 既に ActiveWorkbook が期待パスなら、再アクティブ化と長いポーリングをスキップする。
                    ExcelWorkbook activeEarly = null;
                    try
                    {
                        activeEarly = excelApp.ActiveWorkbook;
                        if (activeEarly != null)
                        {
                            try
                            {
                                string activeFull = activeEarly.FullName;
                                if (!string.IsNullOrEmpty(activeFull))
                                {
                                    string normalizedActiveEarly = NormalizeExcelPath(activeFull);
                                    if (!string.IsNullOrEmpty(normalizedFilePath) &&
                                        normalizedActiveEarly == normalizedFilePath)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] Short-circuit: already active {activeFull}");
                                        return true;
                                    }
                                }
                            }
                            finally
                            {
                                try { Marshal.ReleaseComObject(activeEarly); } catch { }
                                activeEarly = null;
                            }
                        }
                    }
                    catch
                    {
                        if (activeEarly != null)
                        {
                            try { Marshal.ReleaseComObject(activeEarly); } catch { }
                        }
                    }

                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            string wbFullName = wb.FullName;
                            string normalizedWbPath = NormalizeExcelPath(wbFullName);
                            
                            System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] Comparing:");
                            System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile]   Target: {normalizedFilePath}");
                            System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile]   Workbook: {normalizedWbPath}");
                            
                            // フルパスで完全一致する場合のみアクティブにする
                            if (normalizedWbPath == normalizedFilePath)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] Match found! Activating workbook: {wb.Name}");
                                TryActivateWorkbookForScoring(excelApp, wb);

                                // 採点側（ExcelChecker）は ActiveWorkbook を参照して filePath を取得するものがあるため、
                                // "ActiveWorkbookが期待したブックになった" ことを確認してから true を返す。
                                // 待機は短めにし、失敗時は上位リトライ経路へ委譲する。
                                const int timeoutMs = 900;
                                const int pollIntervalMs = 50;
                                var sw = System.Diagnostics.Stopwatch.StartNew();

                                while (sw.ElapsedMilliseconds < timeoutMs)
                                {
                                    try
                                    {
                                        var active = excelApp.ActiveWorkbook;
                                        if (active != null && !string.IsNullOrEmpty(active.FullName))
                                        {
                                            var normalizedActivePath = NormalizeExcelPath(active.FullName);
                                            if (normalizedActivePath == normalizedFilePath)
                                            {
                                                System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] ActiveWorkbook switched successfully: {active.FullName}");
                                                Marshal.ReleaseComObject(wb);
                                                return true;
                                            }
                                        }
                                    }
                                    catch { }

                                    // 300ms 経過しても切り替わらない場合は、ウィンドウ単位のアクティブ化を試みる
                                    if (sw.ElapsedMilliseconds > 300)
                                    {
                                        TryActivateWorkbookForScoring(excelApp, wb);
                                    }

                                    Thread.Sleep(pollIntervalMs);
                                }

                                // タイムアウトしたが、wb.Activate() は成功しているはずなので、
                                // 名前だけでも一致すれば許容する（OneDrive等のパス不一致対策）。
                                try
                                {
                                    var activeFinal = excelApp.ActiveWorkbook;
                                    if (activeFinal != null && string.Equals(activeFinal.Name, Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase))
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] Warning: FullPath mismatch but Name matches. Proceeding.");
                                        Marshal.ReleaseComObject(wb);
                                        return true;
                                    }

                                }
                                catch { }

                                Marshal.ReleaseComObject(wb);
                                return false;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] Error checking workbook: {ex.Message}");
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] No matching workbook found for: {filePath}");
                    if (TryOpenWorkbookInCurrentExcel(excelApp, filePath))
                    {
                        return true;
                    }
                }
                
                return false;
            }
            catch (COMException)
            {
                // COM不整合時は採点専用インスタンスを作り直して1回だけ再試行
                try { if (_scoringExcelApp != null) Marshal.ReleaseComObject(_scoringExcelApp); } catch { }
                _scoringExcelApp = null;
                if (allowRetryOnComException)
                {
                    EnsureScoringExcelApplication(
                        makeVisible: true,
                        timeoutMs: 30000,
                        caller: "ReviewPageWindow.ActivateExcelFile");
                    return ActivateExcelFileInternal(filePath, allowRetryOnComException: false);
                }
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] Error: {ex.Message}");
                return false;
            }
        }

        private bool TryOpenWorkbookInCurrentExcel(ExcelApp excelApp, string filePath)
        {
            if (excelApp == null || string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return false;

            try
            {
                var wb = excelApp.Workbooks.Open(filePath,
                    UpdateLinks: false,
                    ReadOnly: false,
                    Format: Type.Missing,
                    Password: Type.Missing,
                    WriteResPassword: Type.Missing,
                    IgnoreReadOnlyRecommended: true,
                    Origin: Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,
                    Delimiter: Type.Missing,
                    Editable: true,
                    Notify: false,
                    Converter: Type.Missing,
                    AddToMru: false,
                    Local: false,
                    CorruptLoad: Microsoft.Office.Interop.Excel.XlCorruptLoad.xlNormalLoad);
                if (wb == null) return false;

                try
                {
                    TryActivateWorkbookForScoring(excelApp, wb);
                    var active = excelApp.ActiveWorkbook;
                    if (active == null || string.IsNullOrEmpty(active.FullName))
                        return false;
                    var normalizedActivePath = NormalizeExcelPath(active.FullName);
                    var normalizedTargetPath = NormalizeExcelPath(filePath);
                    return normalizedActivePath == normalizedTargetPath;
                }
                finally
                {
                    try { Marshal.ReleaseComObject(wb); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] TryOpenWorkbookInCurrentExcel failed: {ex.Message}");
                return false;
            }
        }
        
        // Windows APIのSetForegroundWindow関数をインポート
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        
        // 特定のプロジェクトのExcelファイルを閉じるメソッド
        private void CloseProjectExcelFile(string filePath)
        {
            try
            {
                ExcelApp excelApp = null;
                try
                {
                    excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Excel application not found");
                    return;
                }
                
                if (excelApp != null && excelApp.Workbooks != null)
                {
                    string fileName = Path.GetFileName(filePath);
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                                wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                            {
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Closing workbook: {wb.Name}");
                                wb.Close(SaveChanges: false);
                                Marshal.ReleaseComObject(wb);
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error accessing workbook: {ex.Message}");
                        }
                    }
                    Marshal.ReleaseComObject(excelApp.Workbooks);
                }
                
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error closing Excel file: {ex.Message}");
            }
        }

        /// <summary>
        /// テキスト内の"で囲まれた部分の"を削除します
        /// 例：「"最新の商品情報"」→「最新の商品情報」
        /// </summary>
        private string RemoveQuotes(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            // "で囲まれた部分の"を削除
            string result = text;
            int startIndex = 0;
            
            while (startIndex < result.Length)
            {
                // "の開始位置を検索
                int quoteStart = result.IndexOf('"', startIndex);
                if (quoteStart == -1)
                {
                    break;
                }
                
                // "の終了位置を検索
                int quoteEnd = result.IndexOf('"', quoteStart + 1);
                if (quoteEnd == -1)
                {
                    break;
                }
                
                // "を削除（開始と終了の両方）
                result = result.Remove(quoteEnd, 1); // 終了の"を先に削除
                result = result.Remove(quoteStart, 1); // 開始の"を削除
                
                startIndex = quoteStart; // 次の検索開始位置を更新
            }
            
            return result;
        }
    }
    
    public class ReviewProjectInfo
    {
        public string ProjectTitle { get; set; }
        public List<ReviewTaskInfo> Tasks { get; set; }
    }
    
    public class ReviewTaskInfo
    {
        public string TaskTitle { get; set; }
        public string Description { get; set; }
        public int GroupId { get; set; }
        public int ProjectId { get; set; }
        public int TaskId { get; set; }
    }
    
    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Func<T, bool> _canExecute;
        
        public RelayCommand(Action<T> execute, Func<T, bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }
        
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
        
        public bool CanExecute(object parameter)
        {
            System.Diagnostics.Debug.WriteLine($"RelayCommand CanExecute called for parameter: {parameter?.GetType().Name}");
            return _canExecute?.Invoke((T)parameter) ?? true;
        }
        
        public void Execute(object parameter)
        {
            System.Diagnostics.Debug.WriteLine($"RelayCommand Execute called for parameter: {parameter?.GetType().Name}");
            _execute((T)parameter);
        }
    }
}
