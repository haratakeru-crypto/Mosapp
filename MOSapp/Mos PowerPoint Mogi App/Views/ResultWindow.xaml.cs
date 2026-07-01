using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;
using Newtonsoft.Json;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Windows.Input;

namespace MOS_PowerPoint_app.Views
{
    /// <summary>
    /// ResultWindow.xaml の相互作用ロジック（PowerPoint用）
    /// </summary>
    public partial class ResultWindow : Window
    {
        private Dictionary<int, bool[]> _projectTaskCompletedStates; // 解答済み状態を保持
        private Dictionary<int, bool[]> _projectTaskFlaggedStates; // 「あとで見直す」フラグ状態
        private Dictionary<int, bool[]> _projectTaskViewedStates; // 閲覧状態（未読問題の追跡用）
        private Dictionary<int, List<bool>> _allProjectScoringResults; // 方式A: 結果表示時に実行した全プロジェクトの採点結果（projectId → タスクごと正否）
        private Dictionary<int, List<bool>> _initialProjectScoringResults; // 初回採点結果の保持
        private HashSet<string> _initialWrongTaskKeys = new HashSet<string>(StringComparer.Ordinal);
        private int _groupId = 1;
        private List<ResultProjectInfo> _allProjects; // すべてのプロジェクトを保持
        private bool _showingWrongOnly = false; // フィルター状態
        private bool _csvExported = false; // CSV出力を1回だけ行うためのフラグ
        public Action<int, int> OnNavigateToTask { get; set; } // ProjectId, TaskId
        public Action OnEndRequested { get; set; } // 終了ボタン押下時（PowerPoint終了・メイン画面に戻る）

        public ResultWindow(Dictionary<int, bool[]> projectTaskCompletedStates = null,
                           Dictionary<int, bool[]> projectTaskFlaggedStates = null, 
                           Dictionary<int, bool[]> projectTaskViewedStates = null, 
                           int groupId = 1,
                           Dictionary<int, List<bool>> allProjectScoringResults = null)
        {
            InitializeComponent();
            _projectTaskCompletedStates = projectTaskCompletedStates ?? new Dictionary<int, bool[]>();
            _projectTaskFlaggedStates = projectTaskFlaggedStates ?? new Dictionary<int, bool[]>();
            _projectTaskViewedStates = projectTaskViewedStates ?? new Dictionary<int, bool[]>();
            _allProjectScoringResults = allProjectScoringResults ?? new Dictionary<int, List<bool>>();
            _initialProjectScoringResults = CloneScoringResults(_allProjectScoringResults);
            _groupId = groupId;
            System.Diagnostics.Debug.WriteLine($"[ResultWindow] Constructor called with {_projectTaskFlaggedStates?.Count ?? 0} projects, scoring results: {_allProjectScoringResults?.Count ?? 0}, groupId: {_groupId}");
            
            // ウィンドウが読み込まれた後にデータを読み込む（非同期）
            this.Loaded += ResultWindow_Loaded;
        }

        private async void ResultWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // UI更新の機会を与える
            await Task.Delay(50);
            
            // 非同期でデータを読み込む
            await LoadResultsAsync();
        }

        private async Task LoadResultsAsync()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] LoadResultsAsync called");
                
                // JSONファイルから全プロジェクトのデータを読み込む（全63問を考慮するため）
                ProjectData projectData = await Task.Run(() =>
                {
                    // JSONファイルのパス（プロジェクトルートまたはReferences/JSON）
                    string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOS模擬アプリ問題文一覧_PowerPoint.json");
                    
                    if (!File.Exists(jsonPath))
                    {
                        jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", "MOS模擬アプリ問題文一覧_PowerPoint.json");
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[ResultWindow] Loading from: {jsonPath}");
                    
                    if (!File.Exists(jsonPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] JSON file not found: {jsonPath}");
                        return null;
                    }
                    
                    // JSONファイルを読み込む（バックグラウンドスレッド）
                    string jsonContent = File.ReadAllText(jsonPath, System.Text.Encoding.UTF8);
                    return JsonConvert.DeserializeObject<ProjectData>(jsonContent);
                });
                
                if (projectData == null || projectData.Projects == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ResultWindow] Project data is null");
                    int fallbackTotal = 63;
                    await Dispatcher.InvokeAsync(() =>
                    {
                        WrongCountTextBlock.Text = "0";
                        TotalCountTextBlock.Text = $"/ {fallbackTotal}";
                        CorrectCountTextBlock.Text = $"{fallbackTotal}";
                        CorrectTotalCountTextBlock.Text = $"/ {fallbackTotal}";
                        AccuracyTextBlock.Text = "初回 100.0% / 修正後 100.0%";
                    });
                    await LoadProjectDataAsync();
                    return;
                }

                int latestWrongTasks = UpdateSummary(projectData);

                // 結果を表示（非同期で読み込む）
                await LoadProjectDataAsync();

                // 採点表CSVをデスクトップに1回だけ出力
                if (!_csvExported && _allProjects != null)
                {
                    try
                    {
                        await Task.Run(() => ExportScoringCsvToDesktop(latestWrongTasks, _allProjects));
                        _csvExported = true;
                    }
                    catch (Exception csvEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] CSV export error: {csvEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error loading results: {ex.Message}\n{ex.StackTrace}");
                await Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show($"結果の読み込み中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        private async Task LoadProjectDataAsync()
        {
            try
            {
                // UIスレッドで読み込み中の表示
                await Dispatcher.InvokeAsync(() =>
                {
                    ProjectsItemsControl.ItemsSource = null;
                });
                
                // バックグラウンドでJSONファイルを読み込む
                ProjectData projectData = await Task.Run(() =>
                {
                    // JSONファイルのパス（プロジェクトルートまたはReferences/JSON）
                    string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOS模擬アプリ問題文一覧_PowerPoint.json");
                    
                    if (!File.Exists(jsonPath))
                    {
                        jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", "MOS模擬アプリ問題文一覧_PowerPoint.json");
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[ResultWindow] Loading from: {jsonPath}");
                    
                    if (!File.Exists(jsonPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] JSON file not found: {jsonPath}");
                        return null;
                    }
                    
                    // JSONファイルを読み込む（バックグラウンドスレッド）
                    string jsonContent = File.ReadAllText(jsonPath, System.Text.Encoding.UTF8);
                    return JsonConvert.DeserializeObject<ProjectData>(jsonContent);
                });
                
                // UIスレッドでデータを処理して表示
                if (projectData == null)
                {
                    await Dispatcher.InvokeAsync(() =>
                    {
                        ProjectsItemsControl.ItemsSource = new List<ResultProjectInfo>();
                    });
                    return;
                }
                
                // データ処理をバックグラウンドで実行（Brushes以外）
                var resultProjects = await Task.Run(() =>
                {
                    return ProcessProjectDataRaw(projectData);
                });
                
                // UIスレッドでBrushesを設定して表示
                await Dispatcher.InvokeAsync(() =>
                {
                    // Brushesを設定
                    foreach (var project in resultProjects)
                    {
                        foreach (var task in project.Tasks)
                        {
                            if (task.ResultMark == "〇")
                            {
                                task.ResultColor = Brushes.Green; // 〇は緑色に設定
                            }
                            else if (task.ResultMark == "✖")
                            {
                                task.ResultColor = Brushes.Red;
                            }
                            else if (task.ResultMark == "時間切れ")
                            {
                                task.ResultColor = new SolidColorBrush(Color.FromRgb(0xB4, 0x53, 0x09)); // オレンジ系
                            }
                            else
                            {
                                task.ResultColor = Brushes.Transparent; // 空欄の場合は透明
                            }
                        }
                    }
                    
                    ProjectsItemsControl.ItemsSource = resultProjects;
                    _allProjects = resultProjects; // すべてのプロジェクトを保存
                }, System.Windows.Threading.DispatcherPriority.Normal);
                
                // UI更新の機会を与える
                await Task.Delay(50);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error loading project data: {ex.Message}\n{ex.StackTrace}");
                await Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show($"問題文の読み込み中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                });
            }
        }

        /// <summary>
        /// 通し順（1-1, 1-2, … 10-8）で最初に未閲覧のタスクの (ProjectId, TaskId) を返す。いなければ (0, 0)。
        /// </summary>
        private void GetFirstUnviewedTask(ProjectData projectData, out int firstProjectId, out int firstTaskId)
        {
            firstProjectId = 0;
            firstTaskId = 0;
            if (projectData?.Projects == null) return;
            foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
            {
                if (project.Tasks == null) continue;
                bool[] viewedStates = _projectTaskViewedStates.ContainsKey(project.ProjectId)
                    ? _projectTaskViewedStates[project.ProjectId]
                    : new bool[0];
                foreach (var task in project.Tasks.OrderBy(t => t.TaskId))
                {
                    int arrayIndex = task.TaskId - 1;
                    bool isUnread = arrayIndex >= viewedStates.Length || (arrayIndex >= 0 && !viewedStates[arrayIndex]);
                    if (isUnread)
                    {
                        firstProjectId = project.ProjectId;
                        firstTaskId = task.TaskId;
                        return;
                    }
                }
            }
        }

        // バックグラウンドで実行するバージョン（Brushesを使わない）
        private List<ResultProjectInfo> ProcessProjectDataRaw(ProjectData projectData)
        {
            var resultProjects = new List<ResultProjectInfo>();
            GetFirstUnviewedTask(projectData, out int firstProjectId, out int firstTaskId);

            try
            {
                if (projectData?.Projects != null)
                {
                    foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
                    {
                        // 各状態を取得
                        bool[] completedStates = _projectTaskCompletedStates.ContainsKey(project.ProjectId) 
                            ? _projectTaskCompletedStates[project.ProjectId] 
                            : new bool[0];
                        bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(project.ProjectId) 
                            ? _projectTaskFlaggedStates[project.ProjectId] 
                            : new bool[0];
                        bool[] viewedStates = _projectTaskViewedStates.ContainsKey(project.ProjectId) 
                            ? _projectTaskViewedStates[project.ProjectId] 
                            : new bool[0];
                        
                        var resultProject = new ResultProjectInfo
                        {
                            ProjectTitle = $"プロジェクト {project.ProjectId}",
                            Tasks = project.Tasks?.Select((task, index) =>
                            {
                                // タスクIDは1始まり、配列は0始まりなので -1
                                int arrayIndex = task.TaskId - 1;
                                
                                bool isCompleted = arrayIndex >= 0 && arrayIndex < completedStates.Length && completedStates[arrayIndex];
                                bool isFlagged = arrayIndex >= 0 && arrayIndex < flaggedStates.Length && flaggedStates[arrayIndex];
                                bool isUnread = arrayIndex >= viewedStates.Length || (arrayIndex >= 0 && !viewedStates[arrayIndex]);
                                
                                string resultMark;
                                // 方式A: 全プロジェクト採点結果がある場合はそれを優先（あとで見直すは従来どおり✖）
                                bool? scoringPassed = null;
                                if (_allProjectScoringResults != null && _allProjectScoringResults.ContainsKey(project.ProjectId))
                                {
                                    var scoreList = _allProjectScoringResults[project.ProjectId];
                                    if (arrayIndex >= 0 && arrayIndex < scoreList.Count)
                                        scoringPassed = scoreList[arrayIndex];
                                }
                                if (scoringPassed.HasValue)
                                {
                                    resultMark = isFlagged ? "✖" : (scoringPassed.Value ? "〇" : "✖");
                                }
                                else if (isUnread)
                                {
                                    // 未閲覧: 最初の未閲覧のみ「時間切れ」、以降は空白
                                    resultMark = (project.ProjectId == firstProjectId && task.TaskId == firstTaskId) ? "時間切れ" : "";
                                }
                                else
                                {
                                    // 閲覧済み: フラグありなら✖、完了かつフラグなしなら〇
                                    resultMark = (isCompleted && !isFlagged) ? "〇" : "✖";
                                }
                                
                                return new ResultTaskInfo
                                {
                                    TaskTitle = $"タスク {task.TaskId}",
                                    Description = RemoveQuotes(task.Description),
                                    ProjectId = project.ProjectId,
                                    TaskId = task.TaskId,
                                    ResultMark = resultMark,
                                    ResultColor = null, // 後でUIスレッドで設定
                                    IsInitiallyWrong = _initialWrongTaskKeys.Contains(GetTaskKey(project.ProjectId, task.TaskId))
                                };
                            }).ToList() ?? new List<ResultTaskInfo>()
                        };
                        
                        resultProjects.Add(resultProject);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error in ProcessProjectDataRaw: {ex.Message}\n{ex.StackTrace}");
            }
            
            return resultProjects;
        }

        /// <summary>
        /// 採点結果を教材用採点表CSVに記入し、デスクトップに書き出す。
        /// </summary>
        private void ExportScoringCsvToDesktop(int totalWrongTasks, List<ResultProjectInfo> resultProjects)
        {
            if (resultProjects == null) return;
            var taskToValue = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var project in resultProjects)
            {
                foreach (var task in project.Tasks ?? Enumerable.Empty<ResultTaskInfo>())
                {
                    string key = $"{task.ProjectId}-{task.TaskId}";
                    string value = task.ResultMark == "✖" ? "×" : (task.ResultMark == "時間切れ" ? "時間切れ" : "");
                    taskToValue[key] = value;
                }
            }

            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOSPP教材用採点表.csv");
            var lines = new List<string>();
            if (File.Exists(templatePath))
            {
                lines.AddRange(File.ReadAllLines(templatePath, Encoding.UTF8));
            }
            else
            {
                lines.Add("教材用プロジェクト,,採点１回目,採点２回目");
                int[] taskCounts = { 8, 8, 7, 8, 7, 7, 5, 5, 5, 8 }; // プロジェクト1～10のタスク数
                for (int p = 1; p <= 10; p++)
                {
                    int taskCount = taskCounts[p - 1];
                    for (int t = 1; t <= taskCount; t++)
                        lines.Add($",{p}-{t},,");
                }
                lines.Add(",×の数,,");
                lines.Add(",▲の数,,");
                lines.Add(",,,");
                lines.Add(",,,");
                lines.Add(",56問,,");
            }

            const int scoreColumnIndex = 2;
            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i];
                string[] parts = line.Split(',');
                if (parts.Length <= scoreColumnIndex) continue;
                string col1 = parts[1].Trim();
                if (taskToValue.TryGetValue(col1, out string value))
                    parts[scoreColumnIndex] = value;
                else if (col1 == "×の数")
                    parts[scoreColumnIndex] = totalWrongTasks.ToString();
                lines[i] = string.Join(",", parts);
            }

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = $"MOSPP教材用採点表_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string outPath = Path.Combine(desktop, fileName);
            var utf8Bom = new UTF8Encoding(true);
            File.WriteAllLines(outPath, lines, utf8Bom);
            System.Diagnostics.Debug.WriteLine($"[ResultWindow] CSV exported to {outPath}");

            ExportWrongAnswersCsvToDesktop(resultProjects);
        }

        private void ExportWrongAnswersCsvToDesktop(List<ResultProjectInfo> resultProjects)
        {
            if (resultProjects == null) return;
            var csvLines = new List<string> { "プロジェクトID,タスク番号,問題文,正誤" };
            foreach (var project in resultProjects)
            {
                foreach (var task in project.Tasks ?? Enumerable.Empty<ResultTaskInfo>())
                {
                    if (string.IsNullOrEmpty(task.ResultMark)) continue;
                    string desc = (task.Description ?? "").Replace("\"", "\"\"");
                    if (desc.Contains(",") || desc.Contains("\n")) desc = "\"" + desc + "\"";
                    csvLines.Add($"{task.ProjectId},{task.TaskId},{desc},{task.ResultMark}");
                }
            }
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fileName = $"MOSPP_間違えた問題_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
            string outPath = Path.Combine(desktop, fileName);
            File.WriteAllLines(outPath, csvLines, new UTF8Encoding(true));
            System.Diagnostics.Debug.WriteLine($"[ResultWindow] Wrong answers CSV exported to {outPath}");
        }

        private string RemoveQuotes(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            
            // 先頭と末尾の引用符を削除
            text = text.Trim();
            if ((text.StartsWith("\"") && text.EndsWith("\"")) ||
                (text.StartsWith("'") && text.EndsWith("'")))
            {
                text = text.Substring(1, text.Length - 2);
            }
            return text;
        }

        private void EndButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[ResultWindow] EndButton_Click called");
            OnEndRequested?.Invoke();
            this.Close();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[ResultWindow] CloseButton_Click called");
            this.Close();
        }

        private void Header_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void ShowWrongOnlyButton_Click(object sender, RoutedEventArgs e)
        {
            _showingWrongOnly = !_showingWrongOnly;
            
            if (_showingWrongOnly)
            {
                // ✖の問題のみ表示
                var wrongOnlyProjects = FilterWrongTasks(_allProjects);
                ProjectsItemsControl.ItemsSource = wrongOnlyProjects;
                ShowWrongOnlyButton.Content = "全て表示";
                ShowWrongOnlyButton.Background = new SolidColorBrush(Color.FromRgb(30, 64, 175)); // #1E40AF
            }
            else
            {
                // 全て表示
                ProjectsItemsControl.ItemsSource = _allProjects;
                ShowWrongOnlyButton.Content = "✖の問題のみ表示";
                ShowWrongOnlyButton.Background = new SolidColorBrush(Color.FromRgb(220, 38, 38)); // #DC2626
            }
        }

        private List<ResultProjectInfo> FilterWrongTasks(List<ResultProjectInfo> projects)
        {
            if (projects == null) return new List<ResultProjectInfo>();
            
            var filteredProjects = new List<ResultProjectInfo>();
            
            foreach (var project in projects)
            {
                var wrongTasks = project.Tasks?.Where(t => t.ResultMark == "✖" || t.ResultMark == "時間切れ").ToList();
                
                // ✖のタスクがある場合のみプロジェクトを追加
                if (wrongTasks != null && wrongTasks.Count > 0)
                {
                    filteredProjects.Add(new ResultProjectInfo
                    {
                        ProjectTitle = project.ProjectTitle,
                        Tasks = wrongTasks
                    });
                }
            }
            
            return filteredProjects;
        }

        private async void TaskRow_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is ResultTaskInfo taskInfo)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Task clicked: ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");
                
                if (OnNavigateToTask != null && taskInfo.ProjectId > 0 && taskInfo.TaskId > 0)
                {
                    try
                    {
                        // AppBarWindowに結果画面からの遷移フラグと参照をセット
                        var appBarWindow = System.Windows.Application.Current.Windows.OfType<UiTestAppBarWindow>().FirstOrDefault();
                        if (appBarWindow != null)
                        {
                            appBarWindow.SetFromResultWindow(true);
                            appBarWindow.SetResultWindow(this);
                            appBarWindow.SetInitialWrongTaskKeys(_initialWrongTaskKeys);
                            appBarWindow.TryPrepareRetryAttemptFromResult(taskInfo.ProjectId, taskInfo.TaskId);
                        }

                        // UI更新の機会を与える
                        await Task.Delay(50);
                        
                        OnNavigateToTask(taskInfo.ProjectId, taskInfo.TaskId);
                        
                        // 結果画面を非表示にする（閉じずに保持）
                        this.Hide();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error navigating to task: {ex.Message}");
                        MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }

        public void ApplyRetryScoreResult(int projectId, int taskId, bool? passed, bool isError)
        {
            if (isError || !passed.HasValue)
                return;

            if (!_allProjectScoringResults.ContainsKey(projectId))
                _allProjectScoringResults[projectId] = new List<bool>();
            var scoreList = _allProjectScoringResults[projectId];
            int index = taskId - 1;
            while (scoreList.Count <= index)
                scoreList.Add(false);
            scoreList[index] = passed.Value;

            _ = Dispatcher.InvokeAsync(async () =>
            {
                try
                {
                    ProjectData projectData = await LoadProjectDataModelAsync();
                    if (projectData != null)
                    {
                        UpdateSummary(projectData);
                        await LoadProjectDataAsync();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ResultWindow] ApplyRetryScoreResult error: {ex.Message}");
                }
            });
        }

        private async Task<ProjectData> LoadProjectDataModelAsync()
        {
            return await Task.Run(() =>
            {
                string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MOS模擬アプリ問題文一覧_PowerPoint.json");
                if (!File.Exists(jsonPath))
                {
                    jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "References", "JSON", "MOS模擬アプリ問題文一覧_PowerPoint.json");
                }
                if (!File.Exists(jsonPath))
                    return null;
                string jsonContent = File.ReadAllText(jsonPath, System.Text.Encoding.UTF8);
                return JsonConvert.DeserializeObject<ProjectData>(jsonContent);
            });
        }

        private int UpdateSummary(ProjectData projectData)
        {
            int totalTasks = projectData.Projects.Sum(p => p.Tasks?.Count ?? 0);
            int initialWrong = CalculateWrongCount(projectData, _initialProjectScoringResults, true);
            int latestWrong = CalculateWrongCount(projectData, _allProjectScoringResults, false);
            CaptureInitialWrongTasks(projectData);

            int latestCorrect = totalTasks - latestWrong;
            double initialAccuracy = totalTasks > 0 ? (double)(totalTasks - initialWrong) / totalTasks * 100.0 : 0.0;
            double latestAccuracy = totalTasks > 0 ? (double)latestCorrect / totalTasks * 100.0 : 0.0;

            Dispatcher.Invoke(() =>
            {
                WrongCountTextBlock.Text = $"{latestWrong}";
                TotalCountTextBlock.Text = $"/ {totalTasks}";
                CorrectCountTextBlock.Text = $"{latestCorrect}";
                CorrectTotalCountTextBlock.Text = $"/ {totalTasks}";
                AccuracyTextBlock.Text = $"初回 {initialAccuracy:F1}% / 修正後 {latestAccuracy:F1}%";
            });
            return latestWrong;
        }

        private int CalculateWrongCount(ProjectData projectData, Dictionary<int, List<bool>> scoringResults, bool collectInitialWrong)
        {
            int totalWrongTasks = 0;
            bool useScoringResults = scoringResults != null && scoringResults.Count > 0;
            foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
            {
                if (project.Tasks == null) continue;
                bool[] completedStates = _projectTaskCompletedStates.ContainsKey(project.ProjectId) ? _projectTaskCompletedStates[project.ProjectId] : new bool[0];
                bool[] flaggedStates = _projectTaskFlaggedStates.ContainsKey(project.ProjectId) ? _projectTaskFlaggedStates[project.ProjectId] : new bool[0];
                bool[] viewedStates = _projectTaskViewedStates.ContainsKey(project.ProjectId) ? _projectTaskViewedStates[project.ProjectId] : new bool[0];
                List<bool> scoreList = useScoringResults && scoringResults.ContainsKey(project.ProjectId) ? scoringResults[project.ProjectId] : null;
                foreach (var task in project.Tasks)
                {
                    int arrayIndex = task.TaskId - 1;
                    bool isCompleted = arrayIndex >= 0 && arrayIndex < completedStates.Length && completedStates[arrayIndex];
                    bool isFlagged = arrayIndex >= 0 && arrayIndex < flaggedStates.Length && flaggedStates[arrayIndex];
                    bool isUnread = arrayIndex >= viewedStates.Length || (arrayIndex >= 0 && !viewedStates[arrayIndex]);
                    bool isWrong;
                    if (scoreList != null && arrayIndex >= 0 && arrayIndex < scoreList.Count)
                        isWrong = isFlagged || !scoreList[arrayIndex];
                    else
                        isWrong = isFlagged || isUnread || !isCompleted;

                    if (isWrong)
                    {
                        totalWrongTasks++;
                        if (collectInitialWrong)
                            _initialWrongTaskKeys.Add(GetTaskKey(project.ProjectId, task.TaskId));
                    }
                }
            }
            return totalWrongTasks;
        }

        private void CaptureInitialWrongTasks(ProjectData projectData)
        {
            if (_initialWrongTaskKeys.Count > 0) return;
            CalculateWrongCount(projectData, _initialProjectScoringResults, true);
        }

        private static Dictionary<int, List<bool>> CloneScoringResults(Dictionary<int, List<bool>> source)
        {
            var clone = new Dictionary<int, List<bool>>();
            if (source == null) return clone;
            foreach (var kvp in source)
                clone[kvp.Key] = kvp.Value != null ? new List<bool>(kvp.Value) : new List<bool>();
            return clone;
        }

        private static string GetTaskKey(int projectId, int taskId)
        {
            return $"{projectId}-{taskId}";
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
        public bool IsInitiallyWrong { get; set; }
    }
}
