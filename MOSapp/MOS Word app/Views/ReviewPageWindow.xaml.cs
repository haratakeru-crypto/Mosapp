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
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;
using WordApp = Microsoft.Office.Interop.Word.Application;
using Libraries;

namespace MOS_Word_app.Views
{
    /// <summary>
    /// ReviewPageWindow.xaml の相互作用ロジック（Wordアプリ用）
    /// </summary>
    public partial class ReviewPageWindow : System.Windows.Window
    {
        public ICommand NavigateToTaskCommand { get; private set; }
        public Action<int, int> OnNavigateToTask { get; set; } // ProjectId, TaskId
        private DispatcherTimer _timer;
        private TimeSpan _remainingTime;
        
        private Dictionary<int, bool[]> _projectTaskCompletedStates;
        private Dictionary<int, bool[]> _projectTaskFlaggedStates;
        private Dictionary<int, bool[]> _projectTaskViewedStates;
        private AppBarWindow _appBarWindow;
        private int _groupId = 1;
        private bool _isWindowClosed;
        
        public ReviewPageWindow(TimeSpan remainingTime, Dictionary<int, bool[]> completedStates, Dictionary<int, bool[]> flaggedStates, Dictionary<int, bool[]> viewedStates = null, AppBarWindow appBarWindow = null, int groupId = 1)
        {
            System.Diagnostics.Debug.WriteLine($"ReviewPageWindow constructor called with remainingTime: {remainingTime}");
            
            InitializeComponent();
            this.Closed += (s, args) => { _isWindowClosed = true; };
            _remainingTime = remainingTime;
            _projectTaskCompletedStates = completedStates;
            _projectTaskFlaggedStates = flaggedStates;
            _projectTaskViewedStates = viewedStates ?? new Dictionary<int, bool[]>();
            _appBarWindow = appBarWindow;
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
            if (!MOS_Word_app.MainWindow.IsTimerDisabled)
                _timer.Start();
        }
        
        private void Timer_Tick(object sender, EventArgs e)
        {
            if (MOS_Word_app.MainWindow.IsTimerDisabled) return;
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
                // レビュー画面の問題文は Word 用 JSON（PowerPoint と独立）
                string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", "MOS模擬アプリ問題文一覧_Word.json");
                if (!File.Exists(jsonPath))
                    jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOS模擬アプリ問題文一覧_Word.json");
                
                if (!File.Exists(jsonPath))
                {
                    string path1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", "MOS模擬アプリ問題文一覧_Word.json");
                    string path2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOS模擬アプリ問題文一覧_Word.json");
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] 正誤判定表JSONが見つかりません。BaseDirectory={AppDomain.CurrentDomain.BaseDirectory}");
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] 試したパス1: {path1}");
                    System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] 試したパス2: {path2}");
                    MessageBox.Show($"問題文JSONファイルが見つかりません: {jsonPath}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] 問題文JSONを読み込みました: {jsonPath}");
                string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
                var projectData = JsonConvert.DeserializeObject<ReviewPageProjectData>(jsonContent);
                
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
                                Description = task.Description,
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
                System.Diagnostics.Debug.WriteLine($"レビューページ読み込みエラー: {ex.Message}");
                MessageBox.Show("問題文の読み込みに失敗しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
                        bool isCompleted = _projectTaskCompletedStates?.ContainsKey(task.ProjectId) == true &&
                                         task.TaskId <= _projectTaskCompletedStates[task.ProjectId].Length &&
                                         _projectTaskCompletedStates[task.ProjectId][task.TaskId - 1];
                        
                        // あとで見直す状態を確認
                        bool isFlagged = _projectTaskFlaggedStates?.ContainsKey(task.ProjectId) == true &&
                                       task.TaskId <= _projectTaskFlaggedStates[task.ProjectId].Length &&
                                       _projectTaskFlaggedStates[task.ProjectId][task.TaskId - 1];
                        
                        System.Diagnostics.Debug.WriteLine($"プロジェクト{task.ProjectId}タスク{task.TaskId}: 解答済み={isCompleted}, フラグ={isFlagged}");
                        
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
                if (child is T)
                {
                    var target = (T)child;
                    if (string.IsNullOrEmpty(name))
                    {
                        return target;
                    }
                    if (child is FrameworkElement)
                    {
                        var fe = (FrameworkElement)child;
                        if (fe.Name == name)
                        {
                            return target;
                        }
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
            System.Diagnostics.Debug.WriteLine($"NavigateToTask called: ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");

            if (OnNavigateToTask == null || taskInfo.ProjectId <= 0 || taskInfo.TaskId <= 0)
            {
                System.Diagnostics.Debug.WriteLine($"ナビゲーション条件不一致: OnNavigateToTask={OnNavigateToTask != null}, ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");
                return;
            }

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
            var stack = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(16)
            };
            stack.Children.Add(new TextBlock
            {
                Text = "タスクを開いています...",
                FontSize = 14,
                TextAlignment = TextAlignment.Center,
                Foreground = System.Windows.Media.Brushes.SteelBlue
            });
            openingOverlay.Content = stack;
            openingOverlay.Show();

            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);

            try
            {
                System.Diagnostics.Debug.WriteLine("ナビゲーション実行開始");
                _timer?.Stop();
                OnNavigateToTask(taskInfo.ProjectId, taskInfo.TaskId);
                openingOverlay.Close();
                System.Diagnostics.Debug.WriteLine("ナビゲーション実行完了");
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
        
        private void TaskButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("TaskButton_Click called");
            
            var button = sender as Button;
            if (button != null && button.DataContext is ReviewTaskInfo)
            {
                var taskInfo = (ReviewTaskInfo)button.DataContext;
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
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        
        private static Window CreateScoringOverlayWindow()
        {
            return new Window
            {
                Title = "採点中",
                Width = 320,
                Height = 140,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize,
                ShowInTaskbar = false,
                Topmost = true,
                Content = new StackPanel
                {
                    Margin = new Thickness(16, 14, 16, 14),
                    VerticalAlignment = VerticalAlignment.Center,
                    Children =
                    {
                        new TextBlock
                        {
                            Text = "採点中です。しばらくお待ちください...",
                            FontSize = 14,
                            TextAlignment = TextAlignment.Center,
                            HorizontalAlignment = HorizontalAlignment.Stretch,
                            Margin = new Thickness(0, 0, 0, 12)
                        },
                        new ProgressBar
                        {
                            Height = 14,
                            IsIndeterminate = true,
                            Minimum = 0,
                            Maximum = 100
                        }
                    }
                }
            };
        }

        private static DispatcherTimer StartScoringOverlayKeepOnTopTimer(Window overlay)
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

        private async void EndExamButton_Click(object sender, RoutedEventArgs e)
        {
            Window scoringOverlay = null;
            DispatcherTimer overlayKeepOnTopTimer = null;
            try
            {
                // ボタンを無効化して再クリックを防止
                var button = sender as Button;
                if (button != null)
                {
                    button.IsEnabled = false;
                    button.Content = "処理中...";
                }
                
                // タイマーを停止
                _timer?.Stop();
                
                // UI更新の機会を与える
                await System.Threading.Tasks.Task.Delay(100);
                
                // UiTestAppBarWindow（Word用バー）を非表示にする
                await Dispatcher.InvokeAsync(() => {
                    var appBarWindows = System.Windows.Application.Current.Windows.OfType<UiTestAppBarWindow>().ToList();
                    foreach (var appBar in appBarWindows)
                    {
                        appBar.Hide();
                    }
                }, DispatcherPriority.Background);
                
                // 採点中オーバーレイ（PowerPoint版と同様・最前面維持）
                scoringOverlay = CreateScoringOverlayWindow();
                scoringOverlay.Show();
                scoringOverlay.Activate();
                overlayKeepOnTopTimer = StartScoringOverlayKeepOnTopTimer(scoringOverlay);
                await System.Threading.Tasks.Task.Yield();
                
                // 全プロジェクト一括採点（結果画面の 〇/✖ 表示用）
                await System.Threading.Tasks.Task.Run(() => WordBatchScoring.ScoreAllProjects(_groupId));
                ScoreResultStore.SnapshotGroup(_groupId);

                // 一括採点用の小窓レイアウトから試験用へ戻す（Word を閉じる前）
                await System.Threading.Tasks.Task.Run(() =>
                    WordWindowLayoutHelper.PositionWordForExamMode());
                
                overlayKeepOnTopTimer?.Stop();
                overlayKeepOnTopTimer = null;
                if (scoringOverlay != null)
                {
                    try { scoringOverlay.Close(); } catch { }
                    scoringOverlay = null;
                }
                
                // Wordアプリケーションを閉じる
                await System.Threading.Tasks.Task.Run(() => CloseWordApplication());
                
                // 結果画面ウィンドウを表示（flaggedStates, viewedStates, groupId を渡して時間切れ・CSV対応）
                ResultWindow resultWindow = null;
                await Dispatcher.InvokeAsync(() =>
                {
                    resultWindow = new ResultWindow(_projectTaskFlaggedStates, _projectTaskViewedStates, _groupId);
                    resultWindow.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    resultWindow.Topmost = true;
                    
                    resultWindow.OnNavigateToTask = (projectId, taskId) =>
                    {
                        var uiTestAppBar = System.Windows.Application.Current.Windows.OfType<UiTestAppBarWindow>().FirstOrDefault();
                        if (uiTestAppBar != null)
                        {
                            uiTestAppBar.Show();
                            uiTestAppBar.Activate();
                            uiTestAppBar.SetReturnToResultMode(resultWindow);
                            uiTestAppBar.ApplyExamWindowLayout();
                        }
                        else
                        {
                            var appBarWindow = System.Windows.Application.Current.Windows.OfType<AppBarWindow>().FirstOrDefault();
                            if (appBarWindow != null)
                            {
                                appBarWindow.Show();
                                appBarWindow.Activate();
                                appBarWindow.SetReturnToResultMode(resultWindow);
                                appBarWindow.ApplyExamWindowLayout();
                            }
                        }
                        OnNavigateToTask(projectId, taskId);
                    };
                    
                    resultWindow.Show();
                    resultWindow.Activate();
                }, DispatcherPriority.Normal);
                
                // ReviewPageWindowを非表示にする（閉じた後は Hide しない）
                await Dispatcher.InvokeAsync(() =>
                {
                    if (!_isWindowClosed)
                        this.Hide();
                }, DispatcherPriority.Normal);
            }
            catch (Exception ex)
            {
                if (scoringOverlay != null)
                {
                    overlayKeepOnTopTimer?.Stop();
                    try { scoringOverlay.Close(); } catch { }
                }
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in EndExamButton_Click: {ex.Message}");
                await Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show($"試験終了処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }
        
        private void CloseWordApplication()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Closing Word application...");
                WordApp wordApp = null;
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    System.Diagnostics.Debug.WriteLine("[ReviewPageWindow] Word application not found");
                    return;
                }
                if (wordApp != null)
                {
                    try
                    {
                        while (wordApp.Documents.Count > 0)
                        {
                            try
                            {
                                wordApp.Documents[1].Close(SaveChanges: false);
                            }
                            catch { break; }
                        }
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error quitting Word: {ex.Message}");
                    }
                }
                try
                {
                    foreach (var process in Process.GetProcessesByName("WINWORD"))
                    {
                        try { process.Kill(); } catch { }
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] Error in CloseWordApplication: {ex.Message}");
            }
        }
    }
    
    /// <summary>
    /// MOS模擬アプリ問題文一覧_Word.json 用のデシリアライズモデル
    /// </summary>
    internal class ReviewPageProjectData
    {
        [JsonProperty("projects")]
        public List<ReviewPageProjectInfo> Projects { get; set; }
    }

    internal class ReviewPageProjectInfo
    {
        [JsonProperty("projectId")]
        public int ProjectId { get; set; }
        [JsonProperty("tasks")]
        public List<ReviewPageTaskInfo> Tasks { get; set; }
    }

    internal class ReviewPageTaskInfo
    {
        [JsonProperty("taskId")]
        public int TaskId { get; set; }
        [JsonProperty("description")]
        public string Description { get; set; }
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




