using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private Dictionary<int, bool[]> _projectTaskCompletedStates;
        private Dictionary<int, bool[]> _projectTaskFlaggedStates;
        private Dictionary<int, bool[]> _projectTaskViewedStates;
        private int _groupId;
        private bool _isWindowClosed;
        private List<ResultProjectInfo> _allProjects;
        private bool _showingWrongOnly;
        private bool _seatChartSubmitStarted;
        private HashSet<int> _includedProjectIds;
        private readonly bool _fromScoringLog;
        private readonly bool _isVariantMode;
        private Dictionary<int, List<bool>> _snapshotResults;

        public Action<int, int> OnNavigateToTask { get; set; }

        public bool ShowDispositionColumns => _isVariantMode;

        private int ResultVariantSetNo =>
            _isVariantMode ? WordVariantScoring.DefaultVariantSetNo : 0;

        public ResultWindow(
            Dictionary<int, bool[]> projectTaskFlaggedStates = null,
            Dictionary<int, bool[]> projectTaskViewedStates = null,
            int groupId = 1,
            bool isVariantMode = false,
            Dictionary<int, bool[]> completedStates = null)
            : this(projectTaskFlaggedStates, projectTaskViewedStates, groupId, null, false, isVariantMode, completedStates)
        {
        }

        public ResultWindow(
            Dictionary<int, List<bool>> snapshotResults,
            int groupId,
            bool fromScoringLog,
            bool isVariantMode = false,
            Dictionary<int, bool[]> completedStates = null,
            Dictionary<int, bool[]> flaggedStates = null)
            : this(flaggedStates, null, groupId, snapshotResults, fromScoringLog, isVariantMode, completedStates)
        {
        }

        public ResultWindow(
            Dictionary<int, bool[]> projectTaskFlaggedStates,
            Dictionary<int, bool[]> projectTaskViewedStates,
            int groupId,
            Dictionary<int, List<bool>> snapshotResults,
            bool fromScoringLog,
            bool isVariantMode = false,
            Dictionary<int, bool[]> completedStates = null)
        {
            InitializeComponent();
            this.Closed += (s, args) => { _isWindowClosed = true; };
            _projectTaskCompletedStates = WordTaskReviewState.Clone(completedStates);
            _projectTaskFlaggedStates = WordTaskReviewState.Clone(projectTaskFlaggedStates);
            _projectTaskViewedStates = projectTaskViewedStates ?? new Dictionary<int, bool[]>();
            _groupId = groupId;
            _snapshotResults = snapshotResults;
            _fromScoringLog = fromScoringLog;
            _isVariantMode = isVariantMode;
            this.Loaded += ResultWindow_Loaded;
            if (!_fromScoringLog)
            {
                this.Closed += (s, args) =>
                {
                    foreach (var appBar in System.Windows.Application.Current.Windows.OfType<UiTestAppBarWindow>())
                        appBar.ClearReturnToResultMode();
                    foreach (var appBar in System.Windows.Application.Current.Windows.OfType<AppBarWindow>())
                        appBar.ClearReturnToResultMode();
                };
            }
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
            // 2回目以降の採点も座席表へ送る
            _seatChartSubmitStarted = false;
            await LoadResultsAsync();
        }

        private async System.Threading.Tasks.Task LoadResultsAsync()
        {
            try
            {
                var projectData = await System.Threading.Tasks.Task.Run(LoadProjectData);

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

                RefreshIncludedProjectIds(projectData);
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
                            else if (task.ResultMark == "－") task.ResultColor = new SolidColorBrush(Color.FromRgb(0x6B, 0x72, 0x80));
                            else task.ResultColor = Brushes.Transparent;

                            ApplyPriorityBadgeColors(task);
                        }
                    }
                    _allProjects = resultProjects;
                    ProjectsItemsControl.ItemsSource = _allProjects;
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

            int totalTasks = projectData.Projects
                .Where(p => IsIncludedProject(p.ProjectId))
                .Sum(p => p.Tasks?.Count ?? 0);
            int unimplemented = CountUnimplementedTasks(projectData);
            int countedTasks = Math.Max(0, totalTasks - unimplemented);
            int initialWrong = CountWrongTasks(projectData, useInitialSnapshot: true);
            int latestWrong = CountWrongTasks(projectData, useInitialSnapshot: false);
            double initialAccuracy = countedTasks > 0 ? (double)(countedTasks - initialWrong) / countedTasks * 100.0 : 0.0;
            double latestAccuracy = countedTasks > 0 ? (double)(countedTasks - latestWrong) / countedTasks * 100.0 : 0.0;

            await Dispatcher.InvokeAsync(() =>
            {
                WrongCountTextBlock.Text = latestWrong.ToString();
                if (AccuracyTextBlock != null)
                    AccuracyTextBlock.Text = $"正答率: 初回 {initialAccuracy:F1}% / 修正後 {latestAccuracy:F1}%";
                if (SummaryTextBlock != null)
                    SummaryTextBlock.Text = $"「あとで見直す」と未閲覧（時間切れ）の合計: {latestWrong}問";
                UpdateTotalScoreDisplay();
            });

            await StartSeatChartSubmitAsync(latestWrong);
        }

        private void UpdateTotalScoreDisplay()
        {
            if (TotalScorePanel == null || TotalScoreTextBlock == null)
                return;

            if (PassingRangeBadge != null)
                PassingRangeBadge.Visibility = Visibility.Collapsed;

            if (!_isVariantMode)
            {
                TotalScorePanel.Visibility = Visibility.Collapsed;
                if (LostPointsTextBlock != null)
                    LostPointsTextBlock.Visibility = Visibility.Collapsed;
                return;
            }

            Dictionary<int, List<bool>> results = _fromScoringLog
                ? _snapshotResults
                : ScoreResultStore.GetProjectResults(_groupId, ResultVariantSetNo);

            TotalScorePanel.Visibility = Visibility.Visible;
            bool showPracticeTotal = HasAllVariantProjects(results);
            if (TotalScoreLabel != null)
                TotalScoreLabel.Visibility = showPracticeTotal ? Visibility.Visible : Visibility.Collapsed;
            TotalScoreTextBlock.Visibility = showPracticeTotal ? Visibility.Visible : Visibility.Collapsed;

            if (showPracticeTotal)
            {
                int total = WordPracticeScore.Compute(WordPracticeScore.FlattenResults(results));
                TotalScoreTextBlock.Text = $"{total}点/1000点";
                if (PassingRangeBadge != null && total >= WordPracticeScore.PassingScoreThreshold)
                    PassingRangeBadge.Visibility = Visibility.Visible;
            }
            else
            {
                TotalScoreTextBlock.Text = string.Empty;
            }

            UpdateLostPointsDisplay(results);
        }

        /// <summary>全プロジェクト採点時だけ 1000 点満点の総合点を出す（一部採点は未採点を不正解扱いするため）。</summary>
        private static bool HasAllVariantProjects(Dictionary<int, List<bool>> results)
        {
            var expected = VariantProblemWorkbook.GetProjectIds();
            if (expected == null || expected.Count == 0 || results == null)
                return false;

            foreach (int projectId in expected)
            {
                if (!results.ContainsKey(projectId))
                    return false;
            }

            return true;
        }

        private void UpdateLostPointsDisplay(Dictionary<int, List<bool>> results)
        {
            if (LostPointsTextBlock == null)
                return;

            if (!WordTaskReviewState.ShowLostPoints(_isVariantMode))
            {
                LostPointsTextBlock.Visibility = Visibility.Collapsed;
                return;
            }

            int completedWrongCount = WordTaskReviewState.CountCompletedWrong(
                results,
                _projectTaskCompletedStates,
                IsUnimplementedTask);
            string message = WordTaskReviewState.FormatLostPointsMessage(completedWrongCount);
            if (string.IsNullOrEmpty(message))
            {
                LostPointsTextBlock.Visibility = Visibility.Collapsed;
                return;
            }

            LostPointsTextBlock.Text = message;
            LostPointsTextBlock.Visibility = Visibility.Visible;
        }

        private async Task StartSeatChartSubmitAsync(int wrongTasks)
        {
            if (_fromScoringLog || _seatChartSubmitStarted) return;
            _seatChartSubmitStarted = true;
            try
            {
                await MosPracticeClient.ResultSubmitBinder.BindAsync(
                    SeatChartSubmitStatusText, SeatChartQrImage, wrongTasks, "Word");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ResultWindow] seat chart submit: " + ex.Message);
            }
        }

        private int CountWrongTasks(ProjectData projectData, bool useInitialSnapshot)
        {
            int totalWrong = 0;
            GetFirstUnviewedWithoutScoringTask(projectData, out int firstProjectId, out int firstTaskId);
            foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
            {
                if (!IsIncludedProject(project.ProjectId)) continue;
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

        private int CountUnimplementedTasks(ProjectData projectData)
        {
            if (!_isVariantMode || projectData?.Projects == null)
                return 0;

            int count = 0;
            foreach (var project in projectData.Projects)
            {
                if (!IsIncludedProject(project.ProjectId) || project.Tasks == null)
                    continue;
                foreach (var task in project.Tasks)
                {
                    if (IsUnimplementedTask(project.ProjectId, task.TaskId))
                        count++;
                }
            }
            return count;
        }

        private bool IsUnimplementedTask(int projectId, int taskId)
        {
            if (!_isVariantMode)
                return false;

            if (!_fromScoringLog)
            {
                string reason = ScoreResultStore.GetFailReason(_groupId, projectId, taskId, ResultVariantSetNo);
                if (string.IsNullOrEmpty(reason))
                    reason = ScoreResultStore.GetInitialFailReason(_groupId, projectId, taskId, ResultVariantSetNo);
                if (WordVariantScoring.IsUnimplementedReason(reason))
                    return true;
            }

            return !WordVariantScoring.IsTaskImplemented(_groupId, projectId, taskId);
        }

        private ProjectData LoadProjectData()
        {
            if (_isVariantMode)
            {
                return new ProjectData
                {
                    Projects = VariantProblemWorkbook.GetProjectIds()
                        .Select(projectId => new ProjectInfo
                        {
                            ProjectId = projectId,
                            Tasks = VariantProblemWorkbook.GetTasks(projectId)
                                .Select(task => new TaskInfo
                                {
                                    TaskId = task.TaskId,
                                    Description = task.Description
                                })
                                .ToList()
                        })
                        .ToList()
                };
            }

            string jsonPath = WordDataPathHelper.FindProblemJson("MOS模擬アプリ問題文一覧_Word.json");
            if (!File.Exists(jsonPath))
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] 正誤判定表JSONが見つかりません。BaseDirectory={AppDomain.CurrentDomain.BaseDirectory}");
                return null;
            }

            string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
            System.Diagnostics.Debug.WriteLine($"[ResultWindow] 問題文JSONを読み込みました: {jsonPath}");
            return JsonConvert.DeserializeObject<ProjectData>(jsonContent);
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
                if (!IsIncludedProject(project.ProjectId)) continue;
                if (project.Tasks == null) continue;
                bool[] viewedStates = _projectTaskViewedStates.ContainsKey(project.ProjectId) ? _projectTaskViewedStates[project.ProjectId] : new bool[0];
                foreach (var task in project.Tasks.OrderBy(t => t.TaskId))
                {
                    int arrayIndex = task.TaskId - 1;
                    bool isUnread = arrayIndex >= viewedStates.Length || (arrayIndex >= 0 && !viewedStates[arrayIndex]);
                    if (!isUnread) continue;
                    if (ScoreResultStore.IsScored(_groupId, project.ProjectId, task.TaskId, ResultVariantSetNo)) continue;
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
            if (IsUnimplementedTask(projectId, taskId))
                return "－";
            if (_fromScoringLog)
            {
                bool? passed = TryGetSnapshotPassed(projectId, taskId);
                if (!passed.HasValue)
                    return "";
                countsAsWrong = !passed.Value;
                return passed.Value ? "〇" : "✖";
            }
            if (isFlagged && !_isVariantMode)
            {
                countsAsWrong = true;
                return "✖";
            }
            bool isPassed;
            bool hasResult = useInitialSnapshot
                ? ScoreResultStore.TryGetInitialResult(_groupId, projectId, taskId, out isPassed, ResultVariantSetNo)
                : ScoreResultStore.TryGetResult(_groupId, projectId, taskId, out isPassed, ResultVariantSetNo);
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
            if (ScoreResultStore.IsIncorrect(_groupId, projectId, taskId, ResultVariantSetNo))
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
                if (!IsIncludedProject(project.ProjectId)) continue;
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
                        string failReason = ResolveDisplayedFailReason(project.ProjectId, task.TaskId, resultMark);
                        bool isCompleted = WordTaskReviewState.GetFlag(_projectTaskCompletedStates, project.ProjectId, task.TaskId);
                        var resultTask = new ResultTaskInfo
                        {
                            TaskTitle = $"タスク {task.TaskId}",
                            Description = RemoveQuotes(task.Description),
                            ProjectId = project.ProjectId,
                            TaskId = task.TaskId,
                            ResultMark = resultMark,
                            ResultColor = null,
                            FailReason = failReason,
                            ShowDispositionColumns = ShowDispositionColumns,
                            IsCompleted = isCompleted,
                            IsFlagged = isFlagged,
                            IsCompletedAndWrong = isCompleted && resultMark == "✖",
                            PriorityBadgeText = PriorityBadgeFor(project.ProjectId, task.TaskId)
                        };
                        return resultTask;
                    }).ToList() ?? new List<ResultTaskInfo>()
                };
                resultProjects.Add(resultProject);
            }
            return resultProjects;
        }

        private string PriorityBadgeFor(int projectId, int taskId)
        {
            if (!_isVariantMode)
                return string.Empty;
            return WordVariantPriorityLookup.GetBadgeText(
                WordVariantScoring.DefaultVariantSetNo, projectId, taskId);
        }

        private static void ApplyPriorityBadgeColors(ResultTaskInfo task)
        {
            if (task == null || !task.HasPriorityBadge)
                return;

            if (task.PriorityBadgeText == "優先")
            {
                task.PriorityBadgeBackground = new SolidColorBrush(Color.FromRgb(0x1E, 0x40, 0xAF));
                task.PriorityBadgeForeground = Brushes.White;
                return;
            }

            if (task.PriorityBadgeText == "次点")
            {
                task.PriorityBadgeBackground = new SolidColorBrush(Color.FromRgb(0xD9, 0x77, 0x06));
                task.PriorityBadgeForeground = Brushes.White;
            }
        }

        private void RefreshIncludedProjectIds(ProjectData projectData)
        {
            if (_fromScoringLog)
            {
                _includedProjectIds = _snapshotResults != null
                    ? new HashSet<int>(_snapshotResults.Keys)
                    : new HashSet<int>();
                return;
            }

            var scored = ScoreResultStore.GetScoredProjectIds(_groupId, ResultVariantSetNo);
            if (scored.Count > 0)
            {
                _includedProjectIds = scored;
                return;
            }

            _includedProjectIds = new HashSet<int>();
            if (projectData?.Projects != null)
            {
                foreach (var project in projectData.Projects)
                    _includedProjectIds.Add(project.ProjectId);
            }
        }

        private bool IsIncludedProject(int projectId)
        {
            if (_includedProjectIds == null || _includedProjectIds.Count == 0)
                return true;
            return _includedProjectIds.Contains(projectId);
        }

        private bool? TryGetSnapshotPassed(int projectId, int taskId)
        {
            if (_snapshotResults == null || !_snapshotResults.TryGetValue(projectId, out var list) || list == null)
                return null;
            int index = taskId - 1;
            if (index < 0 || index >= list.Count)
                return null;
            return list[index];
        }

        private static string RemoveQuotes(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            text = text.Trim();
            if ((text.StartsWith("\"") && text.EndsWith("\"")) || (text.StartsWith("'") && text.EndsWith("'")))
                text = text.Substring(1, text.Length - 2);
            return text;
        }

        private void ShowWrongOnlyButton_Click(object sender, RoutedEventArgs e)
        {
            _showingWrongOnly = !_showingWrongOnly;
            if (_showingWrongOnly)
            {
                ProjectsItemsControl.ItemsSource = (_allProjects ?? new List<ResultProjectInfo>())
                    .Select(project => new ResultProjectInfo
                    {
                        ProjectTitle = project.ProjectTitle,
                        Tasks = (project.Tasks ?? new List<ResultTaskInfo>())
                            .Where(task => task.ResultMark == "✖" || task.ResultMark == "時間切れ")
                            .ToList()
                    })
                    .Where(project => project.Tasks.Count > 0)
                    .ToList();
                ShowWrongOnlyButton.Content = "全て表示";
                ShowWrongOnlyButton.Background = new SolidColorBrush(Color.FromRgb(30, 64, 175));
            }
            else
            {
                ProjectsItemsControl.ItemsSource = _allProjects;
                ShowWrongOnlyButton.Content = "間違えた問題のみ表示";
                ShowWrongOnlyButton.Background = new SolidColorBrush(Color.FromRgb(220, 38, 38));
            }
        }

        private string ResolveDisplayedFailReason(int projectId, int taskId, string resultMark)
        {
            if (resultMark != "✖")
                return "";
            if (_fromScoringLog)
                return "";
            string reason = ScoreResultStore.GetFailReason(_groupId, projectId, taskId, ResultVariantSetNo);
            if (!string.IsNullOrEmpty(reason))
                return reason;
            return ScoreResultStore.GetInitialFailReason(_groupId, projectId, taskId, ResultVariantSetNo);
        }

        private void FailReason_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            if (sender is FrameworkElement element
                && element.DataContext is ResultTaskInfo taskInfo
                && taskInfo.HasFailReason)
            {
                ScoreReasonWindow.Show(this, taskInfo.TaskId, taskInfo.FailReason);
            }
        }

        private static bool IsFailReasonClick(RoutedEventArgs e)
        {
            DependencyObject current = e.OriginalSource as DependencyObject;
            while (current != null)
            {
                if (current is FrameworkElement element && Equals(element.Tag, "FailReason"))
                    return true;
                current = VisualTreeHelper.GetParent(current);
            }
            return false;
        }

        private async void TaskRow_MouseDown(object sender, RoutedEventArgs e)
        {
            if (IsFailReasonClick(e))
                return;
            var element = sender as FrameworkElement;
            var taskInfo = element?.DataContext as ResultTaskInfo;
            if (taskInfo == null || OnNavigateToTask == null || taskInfo.ProjectId <= 0 || taskInfo.TaskId <= 0) return;
            try
            {
                bool canRetry = taskInfo.ResultMark != "－";
                var initialWrongKeys = ScoreResultStore.GetInitialWrongKeys(_groupId, ResultVariantSetNo).ToList();
                var uiTestAppBar = System.Windows.Application.Current.Windows.OfType<UiTestAppBarWindow>().FirstOrDefault();
                if (uiTestAppBar != null)
                {
                    uiTestAppBar.Show();
                    uiTestAppBar.Activate();
                    uiTestAppBar.SetReturnToResultMode(this);
                    uiTestAppBar.SetInitialWrongTaskKeys(initialWrongKeys);
                    if (canRetry)
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
                        if (canRetry)
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

        private static void RestoreProjectSelectionAsMainWindow(MOS_Word_app.MainWindow mainWindow)
        {
            if (mainWindow == null)
                return;
            (mainWindow.DataContext as MOS_Word_app.MainViewModel)?.EnableProjectSelection();
            Application.Current.MainWindow = mainWindow;
            mainWindow.Show();
            mainWindow.Activate();
        }

        private void HideLeftoverExamWindows()
        {
            try
            {
                foreach (Window window in Application.Current.Windows)
                {
                    if (window == null || ReferenceEquals(window, this))
                        continue;
                    if (window is UiTestAppBarWindow || window is AppBarWindow || window is ReviewPageWindow)
                    {
                        try { window.Hide(); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ResultWindow] Error hiding leftover windows: " + ex.Message);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            var mainWindow = Application.Current.Windows.OfType<MOS_Word_app.MainWindow>().FirstOrDefault();
            if (mainWindow == null)
                mainWindow = new MOS_Word_app.MainWindow();
            RestoreProjectSelectionAsMainWindow(mainWindow);
            HideLeftoverExamWindows();
            Close();
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
        /// <summary>×のときの学生向け理由。〇・時間切れは空。</summary>
        public string FailReason { get; set; }
        /// <summary>Word の詳細誤答理由は未完成のため、学生向け「詳細」は出さない。</summary>
        public bool HasFailReason => false;
        public bool ShowDispositionColumns { get; set; }
        public bool IsCompleted { get; set; }
        public bool IsFlagged { get; set; }
        public bool IsCompletedAndWrong { get; set; }
        public string PriorityBadgeText { get; set; }
        public bool HasPriorityBadge => !string.IsNullOrEmpty(PriorityBadgeText);
        public Brush PriorityBadgeBackground { get; set; }
        public Brush PriorityBadgeForeground { get; set; }
    }
}
