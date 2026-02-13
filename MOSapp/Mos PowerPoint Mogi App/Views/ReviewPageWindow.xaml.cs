using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Newtonsoft.Json;
using System.IO;
using System.Text;

namespace MOS_PowerPoint_app.Views
{
    /// <summary>
    /// ReviewPageWindow.xaml の相互作用ロジック（PowerPoint用）
    /// </summary>
    public partial class ReviewPageWindow : Window
    {
        private DispatcherTimer _timer;
        private TimeSpan _remainingTime;
        private Dictionary<int, bool[]> _projectTaskCompletedStates;
        private Dictionary<int, bool[]> _projectTaskFlaggedStates;
        private Dictionary<int, bool[]> _projectTaskViewedStates;
        private int _groupId;

        public Action<int, int> OnNavigateToTask { get; set; }
        public Action OnShowResultRequested { get; set; }
        public Action OnBackRequested { get; set; }

        public ReviewPageWindow(TimeSpan remainingTime,
            Dictionary<int, bool[]> completedStates,
            Dictionary<int, bool[]> flaggedStates,
            Dictionary<int, bool[]> viewedStates,
            int groupId = 1)
        {
            InitializeComponent();
            _remainingTime = remainingTime;
            _projectTaskCompletedStates = completedStates ?? new Dictionary<int, bool[]>();
            _projectTaskFlaggedStates = flaggedStates ?? new Dictionary<int, bool[]>();
            _projectTaskViewedStates = viewedStates ?? new Dictionary<int, bool[]>();
            _groupId = groupId;

            LoadAllProjects();
            InitializeTimer();
            Dispatcher.BeginInvoke(new Action(UpdateTaskStates), DispatcherPriority.Loaded);
        }

        private void InitializeTimer()
        {
            UpdateTimerDisplay();
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromSeconds(1);
            _timer.Tick += Timer_Tick;
            // 「タイマーを使用」にチェックが入っているときだけカウントダウン開始
            if (!MOS_PowerPoint_app.MainWindow.IsTimerDisabled)
                _timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (MOS_PowerPoint_app.MainWindow.IsTimerDisabled) return;
            if (_remainingTime.TotalSeconds > 0)
            {
                _remainingTime = _remainingTime.Subtract(TimeSpan.FromSeconds(1));
                UpdateTimerDisplay();
            }
            else
            {
                _timer?.Stop();
                UpdateTimerDisplay();
            }
        }

        private void UpdateTimerDisplay()
        {
            if (TimerTextBlock != null)
            {
                TimerTextBlock.Text = _remainingTime.ToString(@"hh\:mm\:ss");
            }
        }

        private void LoadAllProjects()
        {
            try
            {
                string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOS模擬アプリ問題文一覧_PowerPoint.json");
                if (!File.Exists(jsonPath))
                {
                    jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", "MOS模擬アプリ問題文一覧_PowerPoint.json");
                }
                if (!File.Exists(jsonPath))
                {
                    ProjectsItemsControl.ItemsSource = new List<ReviewProjectInfo>();
                    return;
                }

                string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
                var projectData = JsonConvert.DeserializeObject<ProjectData>(jsonContent);
                if (projectData?.Projects == null)
                {
                    ProjectsItemsControl.ItemsSource = new List<ReviewProjectInfo>();
                    return;
                }

                var reviewProjects = new List<ReviewProjectInfo>();
                foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
                {
                    reviewProjects.Add(new ReviewProjectInfo
                    {
                        ProjectTitle = $"プロジェクト {project.ProjectId}",
                        Tasks = project.Tasks?.Select(t => new ReviewTaskInfo
                        {
                            TaskTitle = $"タスク {t.TaskId}",
                            Description = RemoveQuotes(t.Description),
                            ProjectId = project.ProjectId,
                            TaskId = t.TaskId
                        }).ToList() ?? new List<ReviewTaskInfo>()
                    });
                }
                ProjectsItemsControl.ItemsSource = reviewProjects;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] LoadAllProjects error: {ex.Message}");
                ProjectsItemsControl.ItemsSource = new List<ReviewProjectInfo>();
            }
        }

        private static string RemoveQuotes(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            text = text.Trim();
            if ((text.StartsWith("\"") && text.EndsWith("\"")) ||
                (text.StartsWith("'") && text.EndsWith("'")))
            {
                return text.Substring(1, text.Length - 2);
            }
            return text;
        }

        private void UpdateTaskStates()
        {
            try
            {
                var projects = ProjectsItemsControl.ItemsSource as List<ReviewProjectInfo>;
                if (projects == null) return;

                foreach (var project in projects)
                {
                    foreach (var task in project.Tasks)
                    {
                        int arrayIndex = task.TaskId - 1;
                        bool isCompleted = false;
                        bool isFlagged = false;

                        if (_projectTaskCompletedStates?.ContainsKey(task.ProjectId) == true)
                        {
                            var arr = _projectTaskCompletedStates[task.ProjectId];
                            if (arrayIndex >= 0 && arrayIndex < arr.Length) isCompleted = arr[arrayIndex];
                        }
                        if (_projectTaskFlaggedStates?.ContainsKey(task.ProjectId) == true)
                        {
                            var arr = _projectTaskFlaggedStates[task.ProjectId];
                            if (arrayIndex >= 0 && arrayIndex < arr.Length) isFlagged = arr[arrayIndex];
                        }
                        UpdateTaskUIStates(task.ProjectId, task.TaskId, isCompleted, isFlagged);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] UpdateTaskStates error: {ex.Message}");
            }
        }

        private void UpdateTaskUIStates(int projectId, int taskId, bool isCompleted, bool isFlagged)
        {
            try
            {
                var container = FindTaskContainer(projectId, taskId);
                if (container != null)
                {
                    UpdateTaskMarks(container, isCompleted, isFlagged);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] UpdateTaskUIStates error: {ex.Message}");
            }
        }

        private FrameworkElement FindTaskContainer(int projectId, int taskId)
        {
            try
            {
                var itemsControl = ProjectsItemsControl;
                if (itemsControl == null) return null;

                for (int i = 0; i < itemsControl.Items.Count; i++)
                {
                    var projectContainer = itemsControl.ItemContainerGenerator.ContainerFromIndex(i) as FrameworkElement;
                    if (projectContainer == null) continue;

                    var taskItemsControl = FindVisualChild<ItemsControl>(projectContainer);
                    if (taskItemsControl == null) continue;

                    for (int j = 0; j < taskItemsControl.Items.Count; j++)
                    {
                        var taskContainer = taskItemsControl.ItemContainerGenerator.ContainerFromIndex(j) as FrameworkElement;
                        if (taskContainer?.DataContext is ReviewTaskInfo info && info.ProjectId == projectId && info.TaskId == taskId)
                        {
                            return taskContainer;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] FindTaskContainer error: {ex.Message}");
            }
            return null;
        }

        private void UpdateTaskMarks(FrameworkElement container, bool isCompleted, bool isFlagged)
        {
            try
            {
                var completedMark = FindVisualChild<TextBlock>(container, "CompletedMark");
                if (completedMark != null)
                    completedMark.Visibility = isCompleted ? Visibility.Visible : Visibility.Collapsed;

                var flaggedMark = FindVisualChild<TextBlock>(container, "FlaggedMark");
                if (flaggedMark != null)
                    flaggedMark.Visibility = isFlagged ? Visibility.Visible : Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPageWindow] UpdateTaskMarks error: {ex.Message}");
            }
        }

        private static T FindVisualChild<T>(DependencyObject parent, string name = null) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T target)
                {
                    if (string.IsNullOrEmpty(name) || (child is FrameworkElement fe && fe.Name == name))
                        return target;
                }
                var found = FindVisualChild<T>(child, name);
                if (found != null) return found;
            }
            return null;
        }

        private void TaskButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is ReviewTaskInfo taskInfo)
            {
                if (OnNavigateToTask != null && taskInfo.ProjectId > 0 && taskInfo.TaskId > 0)
                {
                    OnNavigateToTask(taskInfo.ProjectId, taskInfo.TaskId);
                    this.Close();
                }
            }
        }

        private void ShowResultButton_Click(object sender, RoutedEventArgs e)
        {
            _timer?.Stop();
            OnShowResultRequested?.Invoke();
            this.Close();
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            _timer?.Stop();
            OnBackRequested?.Invoke();
            this.Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            _timer?.Stop();
            base.OnClosed(e);
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
}
