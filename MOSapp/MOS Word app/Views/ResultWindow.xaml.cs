using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Newtonsoft.Json;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Libraries;

namespace MOS_Word_app.Views
{
    /// <summary>
    /// ResultWindow.xaml の相互作用ロジック（Wordアプリ用・後で見直す・時間切れ）
    /// </summary>
    public partial class ResultWindow : System.Windows.Window
    {
        private Dictionary<int, bool[]> _projectTaskFlaggedStates;
        private Dictionary<int, bool[]> _projectTaskViewedStates;
        private int _groupId;
        private bool _isWindowClosed;

        public Action<int, int> OnNavigateToTask { get; set; }

        public ResultWindow(Dictionary<int, bool[]> projectTaskFlaggedStates = null, Dictionary<int, bool[]> projectTaskViewedStates = null, int groupId = 1)
        {
            InitializeComponent();
            this.Closed += (s, args) => { _isWindowClosed = true; };
            _projectTaskFlaggedStates = projectTaskFlaggedStates ?? new Dictionary<int, bool[]>();
            _projectTaskViewedStates = projectTaskViewedStates ?? new Dictionary<int, bool[]>();
            _groupId = groupId;
            this.Loaded += ResultWindow_Loaded;
            // 結果画面を閉じたときは、アプリバー側の「結果に戻る」モードも解除する
            this.Closed += (s, args) =>
            {
                foreach (var appBar in System.Windows.Application.Current.Windows.OfType<UiTestAppBarWindow>())
                    appBar.ClearReturnToResultMode();
                foreach (var appBar in System.Windows.Application.Current.Windows.OfType<AppBarWindow>())
                    appBar.ClearReturnToResultMode();
            };
        }

        private async void ResultWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await System.Threading.Tasks.Task.Delay(50);
            await LoadResultsAsync();
        }

        /// <summary>
        /// 復習後の再採点結果を反映して結果画面を再描画する。
        /// </summary>
        public async Task RefreshResultsAsync()
        {
            await LoadResultsAsync();
        }

        private async System.Threading.Tasks.Task LoadResultsAsync()
        {
            try
            {
                string jsonPath = WordDataPathHelper.FindProblemJson("MOS模擬アプリ問題文一覧_Word.json");

                if (!File.Exists(jsonPath))
                {
                    System.Diagnostics.Debug.WriteLine($"[ResultWindow] 正誤判定表JSONが見つかりません。BaseDirectory={AppDomain.CurrentDomain.BaseDirectory}");
                }

                var projectData = await System.Threading.Tasks.Task.Run(() =>
                {
                    if (!File.Exists(jsonPath)) return null;
                    string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
                    return JsonConvert.DeserializeObject<ProjectData>(jsonContent);
                });

                if (projectData == null || projectData.Projects == null)
                {
                    System.Diagnostics.Debug.WriteLine("[ResultWindow] 問題文データを読み込めませんでした（ファイル未検出またはデシリアライズ失敗）");
                    await Dispatcher.InvokeAsync(() =>
                    {
                        WrongCountTextBlock.Text = "0";
                        if (SummaryTextBlock != null) SummaryTextBlock.Text = "問題文データを読み込めませんでした。";
                    });
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"[ResultWindow] 問題文JSONを読み込みました: {jsonPath}");
                await UpdateSummaryAsync(projectData);

                var resultProjects = await System.Threading.Tasks.Task.Run(() => ProcessProjectDataRaw(projectData));

                await Dispatcher.InvokeAsync(() =>
                {
                    foreach (var project in resultProjects)
                    {
                        foreach (var task in project.Tasks ?? Enumerable.Empty<ResultTaskInfo>())
                        {
                            if (task.ResultMark == "✖") task.ResultColor = Brushes.Red;
                            else if (task.ResultMark == "〇") task.ResultColor = Brushes.Green;
                            else if (task.ResultMark == "時間切れ") task.ResultColor = new SolidColorBrush(Color.FromRgb(0xB4, 0x53, 0x09));
                            else task.ResultColor = Brushes.Transparent;
                        }
                    }
                    ProjectsItemsControl.ItemsSource = resultProjects;
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error loading results: {ex.Message}");
            }
        }

        private async Task UpdateSummaryAsync(ProjectData projectData)
        {
            if (projectData?.Projects == null) return;

            int totalTasks = projectData.Projects.Sum(p => p.Tasks?.Count ?? 0);
            int initialWrong = CountWrongTasks(projectData, useInitialSnapshot: true);
            int latestWrong = CountWrongTasks(projectData, useInitialSnapshot: false);
            double initialAccuracy = totalTasks > 0 ? (double)(totalTasks - initialWrong) / totalTasks * 100.0 : 0.0;
            double latestAccuracy = totalTasks > 0 ? (double)(totalTasks - latestWrong) / totalTasks * 100.0 : 0.0;

            await Dispatcher.InvokeAsync(() =>
            {
                WrongCountTextBlock.Text = latestWrong.ToString();
                if (AccuracyTextBlock != null)
                    AccuracyTextBlock.Text = $"正答率: 初回 {initialAccuracy:F1}% / 修正後 {latestAccuracy:F1}%";
                if (SummaryTextBlock != null)
                    SummaryTextBlock.Text = $"「あとで見直す」と未閲覧（時間切れ）の合計: {latestWrong}問";
            });
        }

        private int CountWrongTasks(ProjectData projectData, bool useInitialSnapshot)
        {
            int totalWrong = 0;
            GetFirstUnviewedWithoutScoringTask(projectData, out int firstProjectId, out int firstTaskId);
            foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
            {
                if (project.Tasks == null) continue;
                bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(project.ProjectId)
                    ? _projectTaskFlaggedStates[project.ProjectId] : new bool[0];
                bool[] viewedStates = _projectTaskViewedStates.ContainsKey(project.ProjectId)
                    ? _projectTaskViewedStates[project.ProjectId] : new bool[0];
                foreach (var task in project.Tasks)
                {
                    int arrayIndex = task.TaskId - 1;
                    bool isFlagged = arrayIndex >= 0 && arrayIndex < flaggedStates.Length && flaggedStates[arrayIndex];
                    bool isUnread = arrayIndex >= viewedStates.Length || (arrayIndex >= 0 && !viewedStates[arrayIndex]);
                    if (CountsAsWrong(project.ProjectId, task.TaskId, isFlagged, isUnread,
                            firstProjectId, firstTaskId, useInitialSnapshot))
                        totalWrong++;
                }
            }
            return totalWrong;
        }

        private bool CountsAsWrong(int projectId, int taskId, bool isFlagged, bool isUnread,
            int firstUnviewedProjectId, int firstUnviewedTaskId, bool useInitialSnapshot)
        {
            ComputeResultMark(projectId, taskId, isFlagged, isUnread,
                firstUnviewedProjectId, firstUnviewedTaskId, out bool countsAsWrong, useInitialSnapshot);
            return countsAsWrong;
        }

        /// <summary>
        /// 採点結果がなく未閲覧の先頭タスク（時間切れ表示用）。
        /// </summary>
        private void GetFirstUnviewedWithoutScoringTask(ProjectData projectData, out int firstProjectId, out int firstTaskId)
        {
            firstProjectId = 0;
            firstTaskId = 0;
            if (projectData?.Projects == null) return;
            foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
            {
                if (project.Tasks == null) continue;
                bool[] viewedStates = _projectTaskViewedStates.ContainsKey(project.ProjectId) ? _projectTaskViewedStates[project.ProjectId] : new bool[0];
                foreach (var task in project.Tasks.OrderBy(t => t.TaskId))
                {
                    int arrayIndex = task.TaskId - 1;
                    bool isUnread = arrayIndex >= viewedStates.Length || (arrayIndex >= 0 && !viewedStates[arrayIndex]);
                    if (!isUnread) continue;
                    if (ScoreResultStore.IsScored(_groupId, project.ProjectId, task.TaskId)) continue;
                    firstProjectId = project.ProjectId;
                    firstTaskId = task.TaskId;
                    return;
                }
            }
        }

        /// <summary>
        /// PowerPoint版と同様: あとで見直す → 採点結果 → 未採点の未閲覧（時間切れ）。
        /// </summary>
        private string ComputeResultMark(int projectId, int taskId, bool isFlagged, bool isUnread,
            int firstUnviewedProjectId, int firstUnviewedTaskId, out bool countsAsWrong, bool useInitialSnapshot = false)
        {
            countsAsWrong = false;
            if (isFlagged)
            {
                countsAsWrong = true;
                return "✖";
            }
            bool isPassed;
            bool hasResult = useInitialSnapshot
                ? ScoreResultStore.TryGetInitialResult(_groupId, projectId, taskId, out isPassed)
                : ScoreResultStore.TryGetResult(_groupId, projectId, taskId, out isPassed);
            if (hasResult)
            {
                if (!isPassed) countsAsWrong = true;
                return isPassed ? "〇" : "✖";
            }
            if (isUnread)
            {
                if (projectId == firstUnviewedProjectId && taskId == firstUnviewedTaskId)
                {
                    countsAsWrong = true;
                    return "時間切れ";
                }
                return "";
            }
            if (ScoreResultStore.IsIncorrect(_groupId, projectId, taskId))
            {
                countsAsWrong = true;
                return "✖";
            }
            return "";
        }

        private List<ResultProjectInfo> ProcessProjectDataRaw(ProjectData projectData)
        {
            var resultProjects = new List<ResultProjectInfo>();
            GetFirstUnviewedWithoutScoringTask(projectData, out int firstProjectId, out int firstTaskId);

            if (projectData?.Projects == null) return resultProjects;
            foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
            {
                bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(project.ProjectId) ? _projectTaskFlaggedStates[project.ProjectId] : new bool[0];
                bool[] viewedStates = _projectTaskViewedStates.ContainsKey(project.ProjectId) ? _projectTaskViewedStates[project.ProjectId] : new bool[0];

                var resultProject = new ResultProjectInfo
                {
                    ProjectTitle = $"プロジェクト {project.ProjectId}",
                    Tasks = project.Tasks?.Select(task =>
                    {
                        int arrayIndex = task.TaskId - 1;
                        bool isFlagged = arrayIndex >= 0 && arrayIndex < flaggedStates.Length && flaggedStates[arrayIndex];
                        bool isUnread = arrayIndex >= viewedStates.Length || (arrayIndex >= 0 && !viewedStates[arrayIndex]);
                        string resultMark = ComputeResultMark(project.ProjectId, task.TaskId, isFlagged, isUnread,
                            firstProjectId, firstTaskId, out _);
                        return new ResultTaskInfo
                        {
                            TaskTitle = $"タスク {task.TaskId}",
                            Description = RemoveQuotes(task.Description),
                            ProjectId = project.ProjectId,
                            TaskId = task.TaskId,
                            ResultMark = resultMark,
                            ResultColor = null
                        };
                    }).ToList() ?? new List<ResultTaskInfo>()
                };
                resultProjects.Add(resultProject);
            }
            return resultProjects;
        }

        private static string RemoveQuotes(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            text = text.Trim();
            if ((text.StartsWith("\"") && text.EndsWith("\"")) || (text.StartsWith("'") && text.EndsWith("'")))
                text = text.Substring(1, text.Length - 2);
            return text;
        }

        private async void TaskRow_MouseDown(object sender, RoutedEventArgs e)
        {
            var element = sender as FrameworkElement;
            var taskInfo = element?.DataContext as ResultTaskInfo;
            if (taskInfo == null || OnNavigateToTask == null || taskInfo.ProjectId <= 0 || taskInfo.TaskId <= 0) return;
            try
            {
                var initialWrongKeys = ScoreResultStore.GetInitialWrongKeys(_groupId).ToList();
                var uiTestAppBar = System.Windows.Application.Current.Windows.OfType<UiTestAppBarWindow>().FirstOrDefault();
                if (uiTestAppBar != null)
                {
                    uiTestAppBar.Show();
                    uiTestAppBar.Activate();
                    uiTestAppBar.SetReturnToResultMode(this);
                    uiTestAppBar.SetInitialWrongTaskKeys(initialWrongKeys);
                    uiTestAppBar.TryEnqueueRetryTask(taskInfo.ProjectId, taskInfo.TaskId);
                    uiTestAppBar.ApplyExamWindowLayout();
                }
                else
                {
                    var appBarWindow = System.Windows.Application.Current.Windows.OfType<AppBarWindow>().FirstOrDefault();
                    if (appBarWindow != null)
                    {
                        appBarWindow.Show();
                        appBarWindow.Activate();
                        appBarWindow.SetReturnToResultMode(this);
                        appBarWindow.SetInitialWrongTaskKeys(initialWrongKeys);
                        appBarWindow.TryEnqueueRetryTask(taskInfo.ProjectId, taskInfo.TaskId);
                        appBarWindow.ApplyExamWindowLayout();
                    }
                }
                await System.Threading.Tasks.Task.Delay(50);
                OnNavigateToTask(taskInfo.ProjectId, taskInfo.TaskId);
                if (!_isWindowClosed)
                    this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            // プロジェクト一覧画面（MainWindow）を表示してから結果画面を閉じる
            var mainWindow = System.Windows.Application.Current.Windows.OfType<MOS_Word_app.MainWindow>().FirstOrDefault();
            if (mainWindow != null)
            {
                mainWindow.Show();
                mainWindow.Activate();
            }
            this.Close();
        }
    }

    public class ResultProjectInfo
    {
        public string ProjectTitle { get; set; }
        public List<ResultTaskInfo> Tasks { get; set; }
    }

    public class ResultTaskInfo
    {
        public string TaskTitle { get; set; }
        public string Description { get; set; }
        public int ProjectId { get; set; }
        public int TaskId { get; set; }
        public string ResultMark { get; set; }
        public Brush ResultColor { get; set; }
    }
}
