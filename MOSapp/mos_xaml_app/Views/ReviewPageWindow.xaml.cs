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
using System.Threading.Tasks;
using System.Threading;
using System.Text;
using Libraries;

namespace MOSExcelMogiApp.Views
{
    /// <summary>
    /// ReviewPageWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ReviewPageWindow : Window
    {
        // #region agent log
        private static readonly string _agentDebugLogPath = @"C:\Users\kouza\source\repos\MOSapp\debug-f11e0d.log";
        private static void AgentLog(string location, string message, object data, string runId, string hypothesisId)
        {
            try
            {
                var payload = new
                {
                    sessionId = "f11e0d",
                    runId,
                    hypothesisId,
                    location,
                    message,
                    data,
                    timestamp = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
                };
                var line = Newtonsoft.Json.JsonConvert.SerializeObject(payload);
                System.IO.File.AppendAllText(_agentDebugLogPath, line + "\n");
            }
            catch
            {
                // ignore logging errors
            }
        }
        // #endregion

        public ICommand NavigateToTaskCommand { get; private set; }
        public Action<int, int> OnNavigateToTask { get; set; } // ProjectId, TaskId
        private DispatcherTimer _timer;
        private TimeSpan _remainingTime;
        private int _groupId = 1; // Group番号（1=模擬①, 2=模擬②, 3=演習）
        
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
                
                string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", jsonFileName);
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
        
        private void NavigateToTask(ReviewTaskInfo taskInfo)
        {
            System.Diagnostics.Debug.WriteLine($"NavigateToTask called: ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");
            
            if (OnNavigateToTask != null && taskInfo.ProjectId > 0 && taskInfo.TaskId > 0)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("ナビゲーション実行開始");
                    
                    // タイマーを停止
                    _timer?.Stop();
                    
                    // ナビゲーションを実行
                    OnNavigateToTask(taskInfo.ProjectId, taskInfo.TaskId);
                    
                    System.Diagnostics.Debug.WriteLine("ナビゲーション実行完了");
                    
                    // レビューページを閉じる
                    this.Close();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ナビゲーション実行エラー: {ex.Message}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"ナビゲーション条件不一致: OnNavigateToTask={OnNavigateToTask != null}, ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");
            }
        }
        
        private void TaskButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("TaskButton_Click called");
            
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
        
        private async void EndExamButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] EndExamButton_Click called");
                // #region agent log
                AgentLog(
                    location: "ReviewPageWindow.EndExamButton_Click",
                    message: "entry",
                    data: new { groupId = _groupId, hasOnNavigateToTask = OnNavigateToTask != null },
                    runId: "pre-fix",
                    hypothesisId: "A");
                // #endregion
                
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
                
                try
                {
                    // AppBarWindow（下の問題領域）を閉じる（UIスレッドで実行）
                    await Dispatcher.InvokeAsync(() => CloseAppBarWindows(), DispatcherPriority.Background);
                    
                    // UI更新の機会を与える
                    await Task.Delay(50);
                    
                    // 「採点中です」オーバーレイを表示（表ではローディング、裏でExcel採点）
                    Window scoringOverlay = null;
                    await Dispatcher.InvokeAsync(() =>
                    {
                        scoringOverlay = new Window
                        {
                            Title = "採点中",
                            Width = 320,
                            Height = 140,
                            WindowStyle = WindowStyle.None,
                            WindowStartupLocation = WindowStartupLocation.CenterScreen,
                            Owner = this,
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
                            Text = "採点中です",
                            FontSize = 18,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(0, 0, 0, 12),
                            Foreground = new SolidColorBrush(Color.FromRgb(30, 64, 175))
                        };
                        var progress = new System.Windows.Controls.ProgressBar
                        {
                            IsIndeterminate = true,
                            Height = 20,
                            Width = 260
                        };
                        stack.Children.Add(text);
                        stack.Children.Add(progress);
                        scoringOverlay.Content = stack;
                        scoringOverlay.Show();
                        try
                        {
                            scoringOverlay.Activate();
                            scoringOverlay.Focus();
                        }
                        catch { }
                    });
                    await Task.Delay(80);
                    
                    // 採点開始直後に既存の Excel を非表示にする（試験中に開いたExcelが前面に出ないように）
                    await Dispatcher.InvokeAsync(() =>
                    {
                        try
                        {
                            var excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                            if (excelApp != null)
                            {
                                excelApp.Visible = false;
                                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Set existing Excel to Visible=false for scoring");
                                Marshal.ReleaseComObject(excelApp);
                            }
                        }
                        catch
                        {
                            // Excel が起動していない場合は無視
                        }
                    });

                    // Excel などに前面を奪われることがあるため、短時間だけ最前面を維持する
                    var overlayTopmostTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(120) };
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
                    
                    // すべてのプロジェクトの採点を実行（バックグラウンド・Excelは非表示）
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Starting to score all projects...");
                    await Task.Run(() => ScoreAllProjects());
                    
                    // 採点が完了したらExcelアプリケーションを閉じる（バックグラウンドで実行）
                    await Task.Run(() => CloseExcelApplication());
                    
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
                    
                    // UI更新の機会を与える
                    await Task.Delay(50);
                    
                    // 結果画面ウィンドウを表示（UIスレッドで実行、非同期で表示）
                    ResultWindow resultWindow = null;
                    
                    await Dispatcher.InvokeAsync(() =>
                    {
                        resultWindow = new ResultWindow(allResults, _groupId);
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
                                    if (!appBarWindow.IsVisible)
                                    {
                                        appBarWindow.Show();
                                    }
                                    appBarWindow.Activate();

                                    // 重要: 表示中のAppBarWindowに直接ナビゲートする（デリゲートの参照先が古い可能性があるため）
                                    appBarWindow.NavigateToTask(projectId, taskId);
                                    return;
                                }

                                // フォールバック: 既存の経路（AppBarWindowが見つからない場合）
                                OnNavigateToTask(projectId, taskId);
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
                                    if (!appBarWindow.IsVisible)
                                    {
                                        appBarWindow.Show();
                                    }
                                    appBarWindow.Activate();
                                    
                                    // NavigateToTaskを呼び出す
                                    appBarWindow.NavigateToTask(projectId, taskId);
                                };
                            }
                        }
                        
                        // 結果ウィンドウを表示
                        resultWindow.Show();
                        resultWindow.Activate();
                    }, DispatcherPriority.Normal);
                    
                    // ResultWindowが完全に読み込まれるまで待つ（Loadedイベントの非同期処理を含む）
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Waiting for ResultWindow to fully load...");
                    await Task.Delay(500); // Loadedイベントが発火するまで待つ
                    
                    // さらにデータ読み込みが完了するまで待つ
                    await Task.Delay(1000); // データ読み込みが完了するまで待つ
                    
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] ResultWindow should be loaded");
                    
                    // ReviewPageWindowを非表示にする（閉じるとResultWindowに影響する可能性があるため）
                    await Dispatcher.InvokeAsync(() =>
                    {
                        System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Hiding ReviewPageWindow");
                        
                        // Application.Current.MainWindowをResultWindowに設定（ReviewPageWindowを非表示にする前に）
                        var resultWindow = Application.Current.Windows.OfType<ResultWindow>().FirstOrDefault();
                        if (resultWindow != null)
                        {
                            Application.Current.MainWindow = resultWindow;
                            System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Set Application.Current.MainWindow to ResultWindow");
                        }
                        
                        // Closeの代わりにHideを使用（ResultWindowの終了ボタンで完全に閉じる）
                        this.Hide();
                    }, DispatcherPriority.Normal);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in EndExamButton_Click: {ex.Message}\n{ex.StackTrace}");
                    await Dispatcher.InvokeAsync(() =>
                    {
                        MessageBox.Show($"試験終了処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                        
                        // ボタンを再有効化
                        if (sender is Button btn)
                        {
                            btn.IsEnabled = true;
                            btn.Content = "試験終了";
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in EndExamButton_Click: {ex.Message}\n{ex.StackTrace}");
                await Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show($"試験終了処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }
        
        private void ScoreAllProjects()
        {
            try
            {
                // ExamResultStorageをクリア
                Models.ExamResultStorage.Clear();
                // #region agent log
                AgentLog(
                    location: "ReviewPageWindow.ScoreAllProjects",
                    message: "start",
                    data: new { groupId = _groupId },
                    runId: "pre-fix",
                    hypothesisId: "A");
                // #endregion
                
                // config.jsonを読み込む
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (!File.Exists(configPath))
                {
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] config.json not found");
                    // #region agent log
                    AgentLog(
                        location: "ReviewPageWindow.ScoreAllProjects",
                        message: "config_missing",
                        data: new { configPath },
                        runId: "pre-fix",
                        hypothesisId: "D");
                    // #endregion
                    return;
                }
                
                string jsonContent = File.ReadAllText(configPath);
                JObject config = JObject.Parse(jsonContent);
                
                // 現在のグループのすべてのプロジェクトを取得
                var groupProjects = config["tabs"]?[_groupId.ToString()]?["projects"];
                if (groupProjects == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] No projects found for group {_groupId}");
                    // #region agent log
                    AgentLog(
                        location: "ReviewPageWindow.ScoreAllProjects",
                        message: "groupProjects_null",
                        data: new { groupId = _groupId },
                        runId: "pre-fix",
                        hypothesisId: "D");
                    // #endregion
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
                // #region agent log
                AgentLog(
                    location: "ReviewPageWindow.ScoreAllProjects",
                    message: "project_list_built",
                    data: new { groupId = _groupId, projectCount = projectList.Count, projects = projectList.Select(p => new { p.projectId, p.taskCount, p.libraryName, p.filePath }).ToList() },
                    runId: "pre-fix",
                    hypothesisId: "A");
                // #endregion
                
                // ステップ1: 既に開いているファイルを確認し、必要に応じて開く
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Step 1: Checking Excel files...");
                
                // すべてのファイルが既に開いているかを確認
                List<string> filesToOpen = new List<string>();
                foreach (var project in projectList)
                {
                    if (string.IsNullOrEmpty(project.filePath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] No file path for project {project.projectId}");
                        continue;
                    }
                    
                    // ファイルが既に開いているかチェック
                    bool isAlreadyOpen = IsExcelFileOpen(project.filePath);
                    
                    if (isAlreadyOpen)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Project {project.projectId} is already open: {project.filePath}");
                    }
                    else
                    {
                        if (File.Exists(project.filePath))
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Project {project.projectId} needs to be opened: {project.filePath}");
                            filesToOpen.Add(project.filePath);
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] File not found for project {project.projectId}: {project.filePath}");
                        }
                    }
                }
                
                // 必要なファイルをCOMで非表示のまま開く
                OpenExcelFilesInBackground(filesToOpen);
                // #region agent log
                AgentLog(
                    location: "ReviewPageWindow.ScoreAllProjects",
                    message: "files_to_open",
                    data: new { filesToOpenCount = filesToOpen.Count, filesToOpen },
                    runId: "pre-fix",
                    hypothesisId: "A");
                // #endregion
                
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Total files to open: {filesToOpen.Count}");
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] All files ready for scoring");
                
                // ステップ2: すべてのプロジェクトを順番に採点（ファイルは開いたまま）
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Step 2: Scoring all projects...");
                for (int idx = 0; idx < projectList.Count; idx++)
                {
                    var project = projectList[idx];
                    // #region agent log
                    AgentLog(
                        location: "ReviewPageWindow.ScoreAllProjects",
                        message: "project_loop_entry",
                        data: new { idx, projectId = project.projectId, project.taskCount, project.libraryName, project.filePath },
                        runId: "pre-fix",
                        hypothesisId: "A");
                    // #endregion
                    
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Scoring project {project.projectId} ({idx + 1}/{projectList.Count}): library={project.libraryName}, taskCount={project.taskCount}");
                    
                    if (string.IsNullOrEmpty(project.filePath) || !File.Exists(project.filePath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] File not found for project {project.projectId}: {project.filePath}");
                        // ファイルが見つからない場合はすべてfalse
                        var falseResults = new List<bool>();
                        for (int i = 0; i < project.taskCount; i++)
                        {
                            falseResults.Add(false);
                        }
                        Models.ExamResultStorage.SaveProjectResult(project.projectId, falseResults);
                        // #region agent log
                        AgentLog(
                            location: "ReviewPageWindow.ScoreAllProjects",
                            message: "file_missing_saved_all_false",
                            data: new { projectId = project.projectId, project.filePath, taskCount = project.taskCount },
                            runId: "pre-fix",
                            hypothesisId: "D");
                        // #endregion
                        continue;
                    }
                    
                    try
                    {
                        // ファイルを明示的にアクティブにする（重要！）
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Activating Excel file for project {project.projectId}: {project.filePath}");
                        bool activated = ActivateExcelFile(project.filePath);
                        // #region agent log
                        string activeWorkbook = null;
                        string activeWorkbookPath = null;
                        int openWorkbookCount = -1;
                        int excelHwnd = 0;
                        try
                        {
                            var excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                            openWorkbookCount = excelApp?.Workbooks?.Count ?? -1;
                            activeWorkbook = excelApp?.ActiveWorkbook?.Name;
                            activeWorkbookPath = excelApp?.ActiveWorkbook?.FullName;
                            excelHwnd = excelApp?.Hwnd ?? 0;
                            if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                        }
                        catch { }
                        AgentLog(
                            location: "ReviewPageWindow.ScoreAllProjects",
                            message: "after_activate",
                            data: new { projectId = project.projectId, activated, expectedPath = project.filePath, openWorkbookCount, excelHwnd, activeWorkbook, activeWorkbookPath },
                            runId: "pre-fix",
                            hypothesisId: "B");
                        // #endregion
                        if (!activated)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] ERROR: Could not activate Excel file for project {project.projectId}");
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] File path: {project.filePath}");
                            
                            // アクティブ化に失敗した場合もfalseを保存
                            var falseResults = new List<bool>();
                            for (int i = 0; i < project.taskCount; i++)
                            {
                                falseResults.Add(false);
                            }
                            Models.ExamResultStorage.SaveProjectResult(project.projectId, falseResults);
                            // #region agent log
                            AgentLog(
                                location: "ReviewPageWindow.ScoreAllProjects",
                                message: "activate_failed_saved_all_false",
                                data: new { projectId = project.projectId, expectedPath = project.filePath, taskCount = project.taskCount },
                                runId: "pre-fix",
                                hypothesisId: "B");
                            // #endregion
                            continue;
                        }
                        
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Successfully activated file for project {project.projectId}");
                        
                        // ファイルがアクティブになるまで十分に待つ
                        System.Threading.Thread.Sleep(1500);
                        
                        // 採点を実行
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Starting scoring for project {project.projectId}");
                        var results = ExecuteScoringForProject(project.libraryName, project.taskCount);
                        
                        // 採点結果を保存
                        Models.ExamResultStorage.SaveProjectResult(project.projectId, results);
                        // #region agent log
                        AgentLog(
                            location: "ReviewPageWindow.ScoreAllProjects",
                            message: "project_scored",
                            data: new { projectId = project.projectId, expectedPath = project.filePath, project.libraryName, taskCount = project.taskCount, resultCount = results?.Count ?? -1, correctCount = results?.Count(r => r) ?? -1 },
                            runId: "pre-fix",
                            hypothesisId: "A");
                        // #endregion
                        
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Project {project.projectId} scored: {results.Count(r => r)}/{results.Count} correct");
                        
                        // UI更新の機会を与える
                        Application.Current.Dispatcher.BeginInvoke(new Action(() => { }), DispatcherPriority.Background);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error scoring project {project.projectId}: {ex.Message}");
                        System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
                        // #region agent log
                        AgentLog(
                            location: "ReviewPageWindow.ScoreAllProjects",
                            message: "exception_scoring_project",
                            data: new { projectId = project.projectId, project.libraryName, expectedPath = project.filePath, ex = ex.Message },
                            runId: "pre-fix",
                            hypothesisId: "E");
                        // #endregion
                        // エラー時はすべてfalse
                        var falseResults = new List<bool>();
                        for (int i = 0; i < project.taskCount; i++)
                        {
                            falseResults.Add(false);
                        }
                        Models.ExamResultStorage.SaveProjectResult(project.projectId, falseResults);
                    }
                }
                
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] All projects scored successfully");
                // #region agent log
                AgentLog(
                    location: "ReviewPageWindow.ScoreAllProjects",
                    message: "end",
                    data: new { groupId = _groupId, projectCount = projectList.Count },
                    runId: "pre-fix",
                    hypothesisId: "A");
                // #endregion
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in ScoreAllProjects: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
                // #region agent log
                AgentLog(
                    location: "ReviewPageWindow.ScoreAllProjects",
                    message: "outer_exception",
                    data: new { ex = ex.Message },
                    runId: "pre-fix",
                    hypothesisId: "E");
                // #endregion
            }
        }
        
        private List<bool> ExecuteScoringForProject(string libraryName, int taskCount)
        {
            var results = new List<bool>();
            
            try
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] ExecuteScoringForProject called with libraryName: {libraryName}, taskCount: {taskCount}");
                // #region agent log
                string activeWbName = null;
                string activeWbFullName = null;
                int excelHwnd = 0;
                try
                {
                    var excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                    activeWbName = excelApp?.ActiveWorkbook?.Name;
                    activeWbFullName = excelApp?.ActiveWorkbook?.FullName;
                    excelHwnd = excelApp?.Hwnd ?? 0;
                    if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                }
                catch { }
                AgentLog(
                    location: "ReviewPageWindow.ExecuteScoringForProject",
                    message: "entry",
                    data: new { libraryName, taskCount, excelHwnd, activeWbName, activeWbFullName },
                    runId: "pre-fix",
                    hypothesisId: "B");
                // #endregion
                
                // Extract group and project numbers from library name
                var parts = libraryName.Replace("ExcelChecker", "").Split('_');
                if (parts.Length >= 2)
                {
                    string groupId = parts[0];
                    string projectId = parts[1];
                    int parsedProjectId = int.TryParse(projectId, out int pid) ? pid : -1;
                    
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
                        // #region agent log
                        AgentLog(
                            location: "ReviewPageWindow.ExecuteScoringForProject",
                            message: "checker_type_found",
                            data: new { libraryName, checkerType = checkerType.FullName },
                            runId: "pre-fix",
                            hypothesisId: "C");
                        // #endregion
                        
                        // デバッグ: 利用可能なメソッドをすべて表示
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Available methods in {checkerType.Name}:");
                        foreach (var method in checkerType.GetMethods())
                        {
                            if (method.Name.StartsWith("CheckTask"))
                            {
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] - {method.Name}");
                            }
                        }
                        
                        object checkerInstance = Activator.CreateInstance(checkerType);
                        
                        for (int i = 1; i <= taskCount; i++)
                        {
                            // MainViewModelと同じロジックで複数のメソッド名形式を試す
                            string[] methodNames;
                            if (groupId == "1")
                            {
                                methodNames = new string[] {
                                    $"CheckTask_1_{projectId}_{i:D2}",          // CheckTask_1_1_01 (優先)
                                    $"CheckTask_1_{projectId}_0{i}",            // CheckTask_1_1_01
                                    $"CheckTask_{groupId}_{projectId}_{i:D2}",  // CheckTask_1_1_01 (fallback)
                                    $"CheckTask_{groupId}_{projectId}_0{i}"     // CheckTask_1_1_01 (fallback)
                                };
                            }
                            else
                            {
                                methodNames = new string[] {
                                    $"CheckTask_{groupId}_{projectId}_{i:D2}",  // CheckTask_2_1_01
                                    $"CheckTask_{groupId}_{projectId}_0{i}",    // CheckTask_2_1_01
                                    $"CheckTask_1_{projectId}_{i:D2}",          // CheckTask_1_1_01 (fallback)
                                    $"CheckTask_1_{projectId}_0{i}"             // CheckTask_1_1_01 (fallback)
                                };
                            }
                            
                            bool methodFound = false;
                            foreach (string methodName in methodNames)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Looking for method: {methodName}");
                                MethodInfo method = checkerType.GetMethod(methodName);
                                
                                if (method != null)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Method found: {methodName}");
                                    // #region agent log
                                    AgentLog(
                                        location: "ReviewPageWindow.ExecuteScoringForProject",
                                        message: "method_found",
                                        data: new { libraryName, taskIndex = i, methodName },
                                        runId: "pre-fix",
                                        hypothesisId: "C");
                                    // #endregion
                                    try
                                    {
                                        bool result = (bool)method.Invoke(checkerInstance, null);
                                        result = ApplyDestructiveValidation(parsedProjectId, i, result);
                                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Method {methodName} result: {result}");
                                        results.Add(result);
                                        // #region agent log
                                        AgentLog(
                                            location: "ReviewPageWindow.ExecuteScoringForProject",
                                            message: "method_result",
                                            data: new { libraryName, taskIndex = i, methodName, result },
                                            runId: "pre-fix",
                                            hypothesisId: "C");
                                        // #endregion
                                        methodFound = true;
                                        break;
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error invoking method {methodName}: {ex.Message}");
                                        System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
                                        // エラー時はfalseとして扱う
                                        results.Add(false);
                                        // #region agent log
                                        AgentLog(
                                            location: "ReviewPageWindow.ExecuteScoringForProject",
                                            message: "method_invoke_exception",
                                            data: new { libraryName, taskIndex = i, methodName, ex = ex.Message },
                                            runId: "pre-fix",
                                            hypothesisId: "E");
                                        // #endregion
                                        methodFound = true;
                                        break;
                                    }
                                }
                                else
                                {
                                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Method not found: {methodName}");
                                }
                            }
                            
                            if (!methodFound)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] No method found for task {i}, returning false");
                                results.Add(false);
                                // #region agent log
                                AgentLog(
                                    location: "ReviewPageWindow.ExecuteScoringForProject",
                                    message: "no_method_for_task_saved_false",
                                    data: new { libraryName, taskIndex = i, tried = methodNames },
                                    runId: "pre-fix",
                                    hypothesisId: "C");
                                // #endregion
                            }
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Checker type not found: {fullTypeName}");
                        // #region agent log
                        AgentLog(
                            location: "ReviewPageWindow.ExecuteScoringForProject",
                            message: "checker_type_not_found",
                            data: new { libraryName, fullTypeName },
                            runId: "pre-fix",
                            hypothesisId: "C");
                        // #endregion
                        // 型が見つからない場合はすべてfalse
                        for (int i = 0; i < taskCount; i++)
                        {
                            results.Add(false);
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Invalid library name format: {libraryName}");
                    // #region agent log
                    AgentLog(
                        location: "ReviewPageWindow.ExecuteScoringForProject",
                        message: "invalid_library_name_format",
                        data: new { libraryName },
                        runId: "pre-fix",
                        hypothesisId: "C");
                    // #endregion
                    // 無効な形式の場合はすべてfalse
                    for (int i = 0; i < taskCount; i++)
                    {
                        results.Add(false);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in ExecuteScoringForProject: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
                // #region agent log
                AgentLog(
                    location: "ReviewPageWindow.ExecuteScoringForProject",
                    message: "outer_exception",
                    data: new { libraryName, taskCount, ex = ex.Message },
                    runId: "pre-fix",
                    hypothesisId: "E");
                // #endregion
                
                // エラー時はすべてfalse
                while (results.Count < taskCount)
                {
                    results.Add(false);
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

        /// <summary>
        /// 破壊的操作検知。入力は常に <c>mos_excel_log.txt</c> の <c>[Op]</c>（<see cref="ExcelLogReader"/>）。
        /// <list type="bullet">
        /// <item><b>全プロジェクト（方式A）</b>: <see cref="ExcelLogReader.TryGetFirstNonExemptViolation"/> — 免除に含まれない操作は違反。範囲は <see cref="ExcelTaskValidationConfig.GetAllowedRanges"/>。</item>
        /// </list>
        /// ルールは <see cref="ExcelTaskValidationConfig"/>（免除 / 許可範囲）。
        /// </summary>
        private bool ApplyDestructiveValidation(int projectId, int taskId, bool checkerResult)
        {
            if (!checkerResult) return false;
            if (projectId <= 0 || taskId <= 0) return checkerResult;

            try
            {
                ExcelValidationExemptFlags exemptFlags = ExcelTaskValidationConfig.GetExemptFlags(projectId, taskId);

                // 方式A: 免除以外の操作はすべて違反。許可範囲があるタスクは TryGetFirstNonExemptViolation 内で範囲判定する。
                if (ExcelLogReader.TryGetFirstNonExemptViolation(
                        projectId,
                        taskId,
                        1,
                        exemptFlags,
                        out string violationMsgA))
                {
                    string line = $"P{projectId}-T{taskId} {violationMsgA}";
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Destructive validation failed (mode A): {line}");
                    ExcelLogReader.AppendDestructiveError(projectId, taskId, 1, line);
                    return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] ApplyDestructiveValidation error: {ex.Message}");
                ExcelLogReader.AppendDestructiveError(projectId, taskId, 1, $"P{projectId}-T{taskId} 例外: {ex.Message}");
                return false;
            }

            return true;
        }
        
        private void CloseExcelApplication()
        {
            ExcelApp excelApp = null;
            int excelPid = -1;

            try
            {
                System.Diagnostics.Debug.WriteLine("[CloseExcelApplication] Starting Excel closure process");

                try
                {
                    excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    System.Diagnostics.Debug.WriteLine("[CloseExcelApplication] No Excel application is running");
                    return;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[CloseExcelApplication] Error getting Excel application: {ex.Message}");
                    return;
                }

                if (excelApp != null)
                {
                    excelPid = ExcelApplicationManager.TryGetExcelProcessId(excelApp);

                    try
                    {
                        // アラートを無効化（自動回復ダイアログを防ぐ）
                        excelApp.DisplayAlerts = false;
                        System.Diagnostics.Debug.WriteLine("[CloseExcelApplication] DisplayAlerts set to false");

                        // すべてのWorkbookを閉じる
                        if (excelApp.Workbooks != null && excelApp.Workbooks.Count > 0)
                        {
                            System.Diagnostics.Debug.WriteLine($"[CloseExcelApplication] Closing {excelApp.Workbooks.Count} workbook(s)");

                            var workbooksToClose = new List<ExcelWorkbook>();
                            foreach (ExcelWorkbook wb in excelApp.Workbooks)
                            {
                                workbooksToClose.Add(wb);
                            }

                            foreach (ExcelWorkbook wb in workbooksToClose)
                            {
                                try
                                {
                                    System.Diagnostics.Debug.WriteLine($"[CloseExcelApplication] Closing workbook: {wb.Name}");

                                    wb.Saved = true;
                                    wb.Close(SaveChanges: false);

                                    Marshal.ReleaseComObject(wb);
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[CloseExcelApplication] Error closing workbook: {ex.Message}");
                                }
                            }
                        }

                        System.Diagnostics.Debug.WriteLine("[CloseExcelApplication] Quitting Excel application");
                        excelApp.Quit();

                        if (excelApp.Workbooks != null)
                        {
                            Marshal.ReleaseComObject(excelApp.Workbooks);
                        }
                        Marshal.ReleaseComObject(excelApp);
                        excelApp = null;

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();

                        System.Diagnostics.Debug.WriteLine("[CloseExcelApplication] Excel closed successfully");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[CloseExcelApplication] Error during Excel closure: {ex.Message}");
                    }
                    finally
                    {
                        // Quit 後も同一 PID が残ると VSTO が再ロードされずログタブが消える（採点→結果→タスク選択経路）
                        if (excelPid > 0)
                        {
                            ExcelApplicationManager.EnsureExcelProcessExited(
                                excelPid,
                                10000,
                                5000,
                                "[CloseExcelApplication]");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CloseExcelApplication] Outer error: {ex.Message}");
            }
            finally
            {
                if (excelApp != null)
                {
                    try { Marshal.ReleaseComObject(excelApp); } catch { }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
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
                // 優先順位1: 現在Excelで開いているファイルをチェック（ユーザーが作業中のファイル）
                string openFilePath = GetOpenExcelFileForProject(groupId, projectId);
                if (!string.IsNullOrEmpty(openFilePath))
                {
                    System.Diagnostics.Debug.WriteLine($"[GetProjectFilePath] Found open Excel file: {openFilePath}");
                    return openFilePath;
                }
                
                // 優先順位2: Initialフォルダをチェック（ユーザーが作業しているファイル）
                string initialPath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\Initial\\project{projectId}.xlsx";
                if (File.Exists(initialPath))
                {
                    System.Diagnostics.Debug.WriteLine($"[GetProjectFilePath] Found Initial folder file: {initialPath}");
                    return initialPath;
                }
                
                // 優先順位3: config.jsonのexcelFileをチェック
                string excelFile = projectConfig["excelFile"]?.ToString();
                if (!string.IsNullOrEmpty(excelFile) && File.Exists(excelFile))
                {
                    System.Diagnostics.Debug.WriteLine($"[GetProjectFilePath] Found excelFile in config: {excelFile}");
                    return excelFile;
                }
                
                // 優先順位4: config.jsonのinitialDataFileをチェック
                string initialDataFile = projectConfig["initialDataFile"]?.ToString();
                if (!string.IsNullOrEmpty(initialDataFile) && File.Exists(initialDataFile))
                {
                    System.Diagnostics.Debug.WriteLine($"[GetProjectFilePath] Found initialDataFile in config: {initialDataFile}");
                    return initialDataFile;
                }
                
                // 優先順位5: Tabフォルダのパスを自動生成
                string generatedPath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\project{projectId}.xlsx";
                System.Diagnostics.Debug.WriteLine($"[GetProjectFilePath] Using generated path: {generatedPath}");
                return generatedPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GetProjectFilePath] Error: {ex.Message}");
                // フォールバック: Initialフォルダのパス
                string fallbackPath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\Initial\\project{projectId}.xlsx";
                System.Diagnostics.Debug.WriteLine($"[GetProjectFilePath] Using fallback path: {fallbackPath}");
                return fallbackPath;
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
                try
                {
                    excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    // Excelが開いていない場合はfalse
                    return false;
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
                                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Excel file is already open: {wb.Name}");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error checking workbook: {ex.Message}");
                        }
                    }
                    Marshal.ReleaseComObject(excelApp.Workbooks);
                }
                
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
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
            if (filesToOpen == null || filesToOpen.Count == 0) return;
            
            ExcelApp excelApp = null;
            try
            {
                try
                {
                    excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                    excelApp.Visible = false;
                }
                catch
                {
                    // COMオートメーション起動を避け、通常起動→接続に寄せる
                    excelApp = ExcelApplicationManager.GetOrCreateExcelApplication(makeVisible: false, timeoutMs: 30000);
                }
                
                foreach (string filePath in filesToOpen)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) continue;
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Opening Excel file (background): {filePath}");
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
                        System.Threading.Thread.Sleep(500);
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
        
        // Excelファイルをアクティブにするメソッド
        private bool ActivateExcelFile(string filePath)
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
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Excel application not found");
                    return false;
                }
                
                if (excelApp != null && excelApp.Workbooks != null)
                {
                    // 正規化したパスで比較（大文字小文字を統一）
                    string normalizedFilePath = Path.GetFullPath(filePath).ToLowerInvariant();
                    
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            string wbFullName = wb.FullName;
                            string normalizedWbPath = Path.GetFullPath(wbFullName).ToLowerInvariant();
                            
                            System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] Comparing:");
                            System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile]   Target: {normalizedFilePath}");
                            System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile]   Workbook: {normalizedWbPath}");
                            
                            // フルパスで完全一致する場合のみアクティブにする
                            if (normalizedWbPath == normalizedFilePath)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] Match found! Activating workbook: {wb.Name}");
                                wb.Activate();

                                // 採点側（ExcelChecker）は ActiveWorkbook を参照して filePath を取得するものがあるため、
                                // "ActiveWorkbookが期待したブックになった" ことを確認してから true を返す。
                                // Active切替が遅延することがあるためポーリングする。
                                const int timeoutMs = 5000;
                                const int pollIntervalMs = 100;
                                var sw = System.Diagnostics.Stopwatch.StartNew();

                                while (sw.ElapsedMilliseconds < timeoutMs)
                                {
                                    try
                                    {
                                        var active = excelApp.ActiveWorkbook;
                                        if (active != null && !string.IsNullOrEmpty(active.FullName))
                                        {
                                            var normalizedActivePath = Path.GetFullPath(active.FullName).ToLowerInvariant();
                                            if (normalizedActivePath == normalizedFilePath)
                                            {
                                                // #region agent log
                                                AgentLog(
                                                    location: "ReviewPageWindow.ActivateExcelFile",
                                                    message: "active_match_found",
                                                    data: new
                                                    {
                                                        expected = filePath,
                                                        excelHwnd = excelApp?.Hwnd ?? 0,
                                                        activeFullName = active.FullName,
                                                        elapsedMs = sw.ElapsedMilliseconds
                                                    },
                                                    runId: "pre-fix",
                                                    hypothesisId: "A");
                                                // #endregion
                                                System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] ActiveWorkbook switched successfully: {active.FullName}");
                                                Marshal.ReleaseComObject(wb);
                                                return true;
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        // ignore and keep polling
                                    }
                                    Thread.Sleep(pollIntervalMs);
                                }

                                // ここまで来たら、wb.Activate() は呼べたが ActiveWorkbook が切り替わっていない
                                try
                                {
                                    var active = excelApp.ActiveWorkbook;
                                    // #region agent log
                                    AgentLog(
                                        location: "ReviewPageWindow.ActivateExcelFile",
                                        message: "active_timeout_or_mismatch",
                                        data: new
                                        {
                                            expected = filePath,
                                            excelHwnd = excelApp?.Hwnd ?? 0,
                                            activeFullName = active?.FullName,
                                            elapsedMs = sw.ElapsedMilliseconds
                                        },
                                        runId: "pre-fix",
                                        hypothesisId: "A");
                                    // #endregion
                                    System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] ActiveWorkbook did not match after timeout. Active={(active?.FullName ?? "null")}, Expected={filePath}");
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
                }
                
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ActivateExcelFile] Error: {ex.Message}");
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
