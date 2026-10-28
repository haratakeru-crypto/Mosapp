using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Newtonsoft.Json;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Threading;
using Libraries;
using MOSExcelMogiApp;
using MOSExcelMogiApp.Infrastructure;

namespace MOSExcelMogiApp.Views
{
    /// <summary>
    /// ResultWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class ResultWindow : Window
    {
        private Dictionary<int, List<bool>> _allProjectResults;
        private int _groupId = 1;
        private List<ResultProjectInfo> _allProjects; // すべてのプロジェクトを保持
        private bool _showingWrongOnly = false; // フィルター状態
        public Action<int, int> OnNavigateToTask { get; set; } // ProjectId, TaskId
        /// <summary>初回の <see cref="LoadResultsAsync"/>（問題文・一覧の読み込み含む）が完了したときに1回だけ発火する。</summary>
        public event EventHandler InitialPresentationCompleted;
        // 採点直後のスナップショット（復習前の正答率計算用）
        private Dictionary<int, List<bool>> _initialProjectResults;
        private bool _seatChartSubmitStarted;
        private readonly bool _fromScoringLog;
        private readonly bool _isVariantMode;
        private readonly int _variantSetNo;
        private readonly MOSExcelMogiApp.Models.ExamResultScope _resultScope;

        public ResultWindow(
            Dictionary<int, List<bool>> allProjectResults = null,
            int groupId = 1,
            bool fromScoringLog = false,
            bool isVariantMode = false,
            int variantSetNo = 1)
        {
            InitializeComponent();
            _allProjectResults = allProjectResults ?? new Dictionary<int, List<bool>>();
            _groupId = groupId;
            _fromScoringLog = fromScoringLog;
            _isVariantMode = isVariantMode;
            _variantSetNo = variantSetNo < 1 ? 1 : variantSetNo;
            _resultScope = _isVariantMode
                ? MOSExcelMogiApp.Models.ExamResultScope.Variant(_variantSetNo)
                : MOSExcelMogiApp.Models.ExamResultScope.Textbook;
            _initialProjectResults = fromScoringLog
                ? CloneResults(_allProjectResults)
                : MOSExcelMogiApp.Models.ExamResultStorage.GetInitialResults(_resultScope);
            System.Diagnostics.Debug.WriteLine($"[ResultWindow] Constructor called with {_allProjectResults?.Count ?? 0} projects, groupId: {_groupId}, fromLog: {_fromScoringLog}");
            
            this.Loaded += ResultWindow_Loaded;

            if (!_fromScoringLog)
                MOSExcelMogiApp.Models.ExamResultStorage.ResultsChanged += OnResultsChanged;
        }

        private static Dictionary<int, List<bool>> CloneResults(Dictionary<int, List<bool>> source)
        {
            var clone = new Dictionary<int, List<bool>>();
            if (source == null) return clone;
            foreach (var kv in source)
                clone[kv.Key] = kv.Value != null ? new List<bool>(kv.Value) : new List<bool>();
            return clone;
        }

        private async void OnResultsChanged(int projectId)
        {
            try
            {
                // 2回目以降の採点も座席表へ送る
                _seatChartSubmitStarted = false;
                // UIスレッドで再計算・再描画（連続発火に備えて少しだけ遅延）
                await Dispatcher.InvokeAsync(async () =>
                {
                    await Task.Delay(20);

                    var currentResults = MOSExcelMogiApp.Models.ExamResultStorage.GetAllResults(_resultScope);
                    if (currentResults != null && currentResults.Count > 0)
                    {
                        _allProjectResults = currentResults;
                    }
                });

                await LoadResultsAsync();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] OnResultsChanged error: {ex.Message}");
            }
        }

        private async void ResultWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                // UI更新の機会を与える
                await Task.Delay(50);

                // 非同期でデータを読み込む
                await LoadResultsAsync();
            }
            finally
            {
                try
                {
                    InitialPresentationCompleted?.Invoke(this, EventArgs.Empty);
                }
                catch
                {
                    // ignore subscriber errors
                }
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                if (!_fromScoringLog)
                    MOSExcelMogiApp.Models.ExamResultStorage.ResultsChanged -= OnResultsChanged;
            }
            catch { }
            base.OnClosed(e);
        }

        private async Task LoadResultsAsync()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] LoadResultsAsync called. Project count: {_allProjectResults?.Count ?? 0}");
                
                // 採点結果がない場合は全問正解と仮定
                if (_allProjectResults == null || _allProjectResults.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[ResultWindow] No results found, displaying all correct");
                    await DisplayAllCorrectAsync();
                    return;
                }

                // 採点直後（初回）の正答率を計算
                var initialResults = (_initialProjectResults != null && _initialProjectResults.Count > 0)
                    ? _initialProjectResults
                    : _allProjectResults;
                double initialCorrectRate;
                int initialWrongTasks;
                CalculateRate(initialResults, out initialCorrectRate, out _, out initialWrongTasks);

                // 現在（復習後）の正答率を計算（保存されている最新結果を使用）
                var currentResults = _fromScoringLog
                    ? _allProjectResults
                    : MOSExcelMogiApp.Models.ExamResultStorage.GetAllResults(_resultScope);
                if (currentResults == null || currentResults.Count == 0)
                {
                    currentResults = _allProjectResults;
                }
                double currentCorrectRate;
                int currentTotalTasks;
                int currentWrongTasks;
                CalculateRate(currentResults, out currentCorrectRate, out currentTotalTasks, out currentWrongTasks);

                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Initial accuracy: {initialCorrectRate:F1}%, Current accuracy: {currentCorrectRate:F1}%");

                // UIスレッドで更新（正答率と現在の間違い数）
                await Dispatcher.InvokeAsync(() =>
                {
                    CorrectRateTextBlock.Text = $"正答率：{initialCorrectRate:F1}％　復習後：{currentCorrectRate:F1}％";
                    WrongCountTextBlock.Text = $"{currentWrongTasks}問";
                    UpdateTotalScoreDisplay(currentResults);
                });

                await StartSeatChartSubmitAsync(currentWrongTasks);

                // 結果を表示（非同期で読み込む）
                await LoadProjectDataAsync();
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

        private async Task DisplayAllCorrectAsync()
        {
            await Dispatcher.InvokeAsync(() =>
            {
                // 採点直後も復習後も 100% として表示
                CorrectRateTextBlock.Text = "正答率：100.0％　復習後：100.0％";
                WrongCountTextBlock.Text = "0問";
                UpdateTotalScoreDisplay(_allProjectResults);
            });

            await StartSeatChartSubmitAsync(0);
            
            // 問題文データを読み込んで表示（全問正解として）
            await LoadProjectDataAsync();
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
                    if (_isVariantMode && _variantSetNo == 1)
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
                                        }).ToList()
                                }).ToList()
                        };
                    }

                    // Group番号に応じたJSONファイルを選択
                    string jsonFileName = _groupId switch
                    {
                        1 => "MOS演習問題文一覧.json",
                        2 => "MOS模擬試験①問題文一覧.json",
                        3 => "MOS模擬試験②問題文一覧.json",
                        _ => "MOS模擬アプリ問題文一覧.json"
                    };
                    
                    string jsonPath = DataPathHelper.ResolveJsonPath(jsonFileName);
                    System.Diagnostics.Debug.WriteLine($"[ResultWindow] Loading from: {jsonFileName} (GroupId: {_groupId})");
                    
                    if (!File.Exists(jsonPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] JSON file not found: {jsonPath}");
                        return null;
                    }
                    
                    // JSONファイルを読み込む（バックグラウンドスレッド）
                    string jsonContent = File.ReadAllText(jsonPath);
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
                                task.ResultColor = Brushes.Green;
                            }
                            else if (task.ResultMark == "×")
                            {
                                task.ResultColor = Brushes.Red;
                            }
                            else
                            {
                                task.ResultColor = Brushes.Gray;
                            }

                            ApplyPriorityBadgeColors(task);
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

        private async Task StartSeatChartSubmitAsync(int wrongTasks)
        {
            if (_fromScoringLog || _seatChartSubmitStarted) return;
            _seatChartSubmitStarted = true;
            try
            {
                await MosPracticeClient.ResultSubmitBinder.BindAsync(
                    SeatChartSubmitStatusText, SeatChartQrImage, wrongTasks, "Excel");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ResultWindow] seat chart submit: " + ex.Message);
            }
        }

        /// <summary>
        /// 指定された結果ディクショナリから正答率と合計問題数・誤答数を計算するヘルパー。
        /// </summary>
        private static void CalculateRate(
            Dictionary<int, List<bool>> projectResults,
            out double correctRate,
            out int totalTasks,
            out int wrongTasks)
        {
            totalTasks = 0;
            int correctTasks = 0;
            wrongTasks = 0;

            if (projectResults != null)
            {
                foreach (var projectResult in projectResults)
                {
                    foreach (var result in projectResult.Value)
                    {
                        totalTasks++;
                        if (result)
                        {
                            correctTasks++;
                        }
                        else
                        {
                            wrongTasks++;
                        }
                    }
                }
            }

            correctRate = totalTasks > 0 ? (double)correctTasks / totalTasks * 100.0 : 100.0;
        }

        private bool IsIncludedProject(int projectId)
        {
            if (_allProjectResults == null || _allProjectResults.Count == 0)
                return true;
            return _allProjectResults.ContainsKey(projectId);
        }

        // バックグラウンドで実行するバージョン（Brushesを使わない）
        private List<ResultProjectInfo> ProcessProjectDataRaw(ProjectData projectData)
        {
            var resultProjects = new List<ResultProjectInfo>();
            
            try
            {
                if (projectData?.Projects != null)
                {
                    foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
                    {
                        if (!IsIncludedProject(project.ProjectId))
                            continue;

                        // 採点結果を取得
                        var projectResults = _allProjectResults.ContainsKey(project.ProjectId) 
                            ? _allProjectResults[project.ProjectId] 
                            : null;
                        
                        var resultProject = new ResultProjectInfo
                        {
                            ProjectTitle = $"プロジェクト {project.ProjectId}",
                            Tasks = project.Tasks?.Select((task, index) =>
                            {
                                bool? isCorrect = null;
                                if (projectResults != null && index < projectResults.Count)
                                {
                                    isCorrect = projectResults[index];
                                }
                                
                                return new ResultTaskInfo
                                {
                                    TaskTitle = $"タスク {task.TaskId}",
                                    Description = RemoveQuotes(task.Description),
                                    GroupId = _groupId,
                                    ProjectId = project.ProjectId,
                                    TaskId = task.TaskId,
                                    ResultMark = isCorrect.HasValue ? (isCorrect.Value ? "〇" : "×") : "",
                                    FailReason = isCorrect == false ? FailReasonFor(project.ProjectId, task.TaskId) : "",
                                    PriorityBadgeText = PriorityBadgeFor(project.ProjectId, task.TaskId),
                                    ResultColor = null // 後でUIスレッドで設定
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
        
        private List<ResultProjectInfo> ProcessProjectData(ProjectData projectData)
        {
            var resultProjects = new List<ResultProjectInfo>();
            
            try
            {
                if (projectData?.Projects != null)
                {
                    foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
                    {
                        if (!IsIncludedProject(project.ProjectId))
                            continue;

                        // 採点結果を取得
                        var projectResults = _allProjectResults.ContainsKey(project.ProjectId) 
                            ? _allProjectResults[project.ProjectId] 
                            : null;
                        
                        var resultProject = new ResultProjectInfo
                        {
                            ProjectTitle = $"プロジェクト {project.ProjectId}",
                            Tasks = project.Tasks?.Select((task, index) =>
                            {
                                bool? isCorrect = null;
                                if (projectResults != null && index < projectResults.Count)
                                {
                                    isCorrect = projectResults[index];
                                }
                                
                                return new ResultTaskInfo
                                {
                                    TaskTitle = $"タスク {task.TaskId}",
                                    Description = RemoveQuotes(task.Description),
                                    GroupId = _groupId,
                                    ProjectId = project.ProjectId,
                                    TaskId = task.TaskId,
                                    ResultMark = isCorrect.HasValue ? (isCorrect.Value ? "〇" : "×") : "",
                                    FailReason = isCorrect == false ? FailReasonFor(project.ProjectId, task.TaskId) : "",
                                    PriorityBadgeText = PriorityBadgeFor(project.ProjectId, task.TaskId),
                                    ResultColor = isCorrect.HasValue 
                                        ? (isCorrect.Value ? Brushes.Green : Brushes.Red)
                                        : Brushes.Gray
                                };
                            }).ToList() ?? new List<ResultTaskInfo>()
                        };

                        if (resultProject.Tasks != null)
                        {
                            foreach (var task in resultProject.Tasks)
                                ApplyPriorityBadgeColors(task);
                        }
                        
                        resultProjects.Add(resultProject);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error in ProcessProjectData: {ex.Message}\n{ex.StackTrace}");
            }
            
            return resultProjects;
        }
        
        private void LoadProjectDataUI(ProjectData projectData)
        {
            try
            {
                if (projectData?.Projects != null)
                {
                    var resultProjects = ProcessProjectData(projectData);
                    ProjectsItemsControl.ItemsSource = resultProjects;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error in LoadProjectDataUI: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"データの表示中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

        private static void RestoreProjectSelectionAsMainWindow(MOSExcelMogiApp.MainWindow mainWindow)
        {
            if (mainWindow == null)
                return;
            Application.Current.MainWindow = mainWindow;
            mainWindow.Show();
            mainWindow.WindowState = WindowState.Normal;
            mainWindow.Activate();
            mainWindow.Focus();
        }

        private void CloseLeftoverExamWindows()
        {
            try
            {
                var windowsToClose = new List<Window>();
                foreach (Window window in Application.Current.Windows)
                {
                    if (window == null || ReferenceEquals(window, this))
                        continue;
                    string name = window.GetType().Name;
                    if (name == "AppBarWindow" || name == "UiTestAppBarWindow" || name == "ReviewPageWindow")
                        windowsToClose.Add(window);
                }

                foreach (var window in windowsToClose)
                {
                    System.Diagnostics.Debug.WriteLine($"[ResultWindow] Closing leftover window: {window.GetType().Name}");
                    try { window.Close(); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error closing leftover windows: {ex.Message}");
            }
        }

        private async void EndButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[ResultWindow] EndButton_Click called");
            
            // 待機メッセージウィンドウを作成
            Window waitWindow = null;
            
            try
            {
                // 待機メッセージを表示
                waitWindow = new Window
                {
                    Title = "画面切替中",
                    Width = 300,
                    Height = 120,
                    WindowStyle = WindowStyle.None,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    Owner = this,
                    ShowInTaskbar = false,
                    ResizeMode = ResizeMode.NoResize,
                    Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.White),
                    BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 64, 175)),
                    BorderThickness = new Thickness(2)
                };
                
                var stackPanel = new System.Windows.Controls.StackPanel
                {
                    Margin = new Thickness(20),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                
                var textBlock = new System.Windows.Controls.TextBlock
                {
                    Text = "画面を切り替えています...",
                    FontSize = 16,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 10),
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 64, 175))
                };
                
                var progressBar = new System.Windows.Controls.ProgressBar
                {
                    IsIndeterminate = true,
                    Height = 20,
                    Width = 250
                };
                
                stackPanel.Children.Add(textBlock);
                stackPanel.Children.Add(progressBar);
                waitWindow.Content = stackPanel;
                waitWindow.Show();
                
                // UI更新の機会を与える
                await Task.Delay(100);
                
                // MainWindowを先に表示
                await Dispatcher.InvokeAsync(() =>
                {
                    MOSExcelMogiApp.MainWindow mainWindow = null;
                    try
                    {
                        // まず、既存のMainWindowを探す
                        mainWindow = Application.Current.Windows.OfType<MOSExcelMogiApp.MainWindow>().FirstOrDefault();
                        
                        // 見つからない場合は、新しく作成
                        if (mainWindow == null)
                        {
                            System.Diagnostics.Debug.WriteLine("[ResultWindow] Creating new MainWindow");
                            mainWindow = new MOSExcelMogiApp.MainWindow();
                        }
                        else
                        {
                            // 既存のMainWindowが見つかった場合、ViewModelの状態をリセット
                            System.Diagnostics.Debug.WriteLine("[ResultWindow] Resetting MainViewModel state");
                            var viewModel = mainWindow.DataContext as Ui.ViewModels.MainViewModel;
                            if (viewModel != null)
                            {
                                viewModel.IsExcelOverlayVisible = false;
                                viewModel.CurrentProject = null;
                                viewModel.ResultMessage = "";
                                viewModel.EnableProjectSelection();
                            }
                        }
                        
                        // MainWindowを確実に表示
                        System.Diagnostics.Debug.WriteLine("[ResultWindow] Showing MainWindow");
                        RestoreProjectSelectionAsMainWindow(mainWindow);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error showing MainWindow: {ex.Message}");
                        // エラーが発生した場合でも、新しいMainWindowを作成
                        try
                        {
                            mainWindow = new MOSExcelMogiApp.MainWindow();
                            RestoreProjectSelectionAsMainWindow(mainWindow);
                        }
                        catch (Exception ex2)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error creating new MainWindow: {ex2.Message}");
                        }
                    }
                }, DispatcherPriority.Normal);
                
                // MainWindowが表示されるまで待つ
                await Task.Delay(300);
                
                // 待機メッセージウィンドウを閉じる
                if (waitWindow != null)
                {
                    waitWindow.Close();
                    waitWindow = null;
                }
                
                await Dispatcher.InvokeAsync(() => CloseLeftoverExamWindows(), DispatcherPriority.Normal);
                
                // 結果画面を閉じる
                System.Diagnostics.Debug.WriteLine("[ResultWindow] Closing ResultWindow");
                this.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error in EndButton_Click: {ex.Message}\n{ex.StackTrace}");
                
                // 待機メッセージウィンドウを閉じる
                if (waitWindow != null)
                {
                    try { waitWindow.Close(); } catch { }
                }
                
                // エラーが発生した場合でも、ResultWindowを閉じる
                try
                {
                    this.Close();
                }
                catch { }
            }
        }

        private async void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[ResultWindow] CloseButton_Click called");
            
            // 待機メッセージウィンドウを作成
            Window waitWindow = null;
            
            try
            {
                // 待機メッセージを表示
                waitWindow = new Window
                {
                    Title = "画面切替中",
                    Width = 300,
                    Height = 120,
                    WindowStyle = WindowStyle.None,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    Owner = this,
                    ShowInTaskbar = false,
                    ResizeMode = ResizeMode.NoResize,
                    Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.White),
                    BorderBrush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 64, 175)),
                    BorderThickness = new Thickness(2)
                };
                
                var stackPanel = new System.Windows.Controls.StackPanel
                {
                    Margin = new Thickness(20),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                
                var textBlock = new System.Windows.Controls.TextBlock
                {
                    Text = "画面を切り替えています...",
                    FontSize = 16,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 10),
                    Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 64, 175))
                };
                
                var progressBar = new System.Windows.Controls.ProgressBar
                {
                    IsIndeterminate = true,
                    Height = 20,
                    Width = 250
                };
                
                stackPanel.Children.Add(textBlock);
                stackPanel.Children.Add(progressBar);
                waitWindow.Content = stackPanel;
                waitWindow.Show();
                
                // UI更新の機会を与える
                await Task.Delay(100);
                
                // MainWindowを先に表示
                await Dispatcher.InvokeAsync(() =>
                {
                    MOSExcelMogiApp.MainWindow mainWindow = null;
                    try
                    {
                        // まず、既存のMainWindowを探す
                        mainWindow = Application.Current.Windows.OfType<MOSExcelMogiApp.MainWindow>().FirstOrDefault();
                        
                        // 見つからない場合は、新しく作成
                        if (mainWindow == null)
                        {
                            System.Diagnostics.Debug.WriteLine("[ResultWindow] Creating new MainWindow");
                            mainWindow = new MOSExcelMogiApp.MainWindow();
                        }
                        else
                        {
                            // 既存のMainWindowが見つかった場合、ViewModelの状態をリセット
                            System.Diagnostics.Debug.WriteLine("[ResultWindow] Resetting MainViewModel state");
                            var viewModel = mainWindow.DataContext as Ui.ViewModels.MainViewModel;
                            if (viewModel != null)
                            {
                                viewModel.IsExcelOverlayVisible = false;
                                viewModel.CurrentProject = null;
                                viewModel.ResultMessage = "";
                                viewModel.EnableProjectSelection();
                            }
                        }
                        
                        // MainWindowを確実に表示し、本体窓をプロジェクト選択へ戻す
                        System.Diagnostics.Debug.WriteLine("[ResultWindow] Showing MainWindow");
                        RestoreProjectSelectionAsMainWindow(mainWindow);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error showing MainWindow: {ex.Message}");
                        // エラーが発生した場合でも、新しいMainWindowを作成
                        try
                        {
                            mainWindow = new MOSExcelMogiApp.MainWindow();
                            RestoreProjectSelectionAsMainWindow(mainWindow);
                        }
                        catch (Exception ex2)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error creating new MainWindow: {ex2.Message}");
                        }
                    }
                }, DispatcherPriority.Normal);
                
                // MainWindowが表示されるまで待つ
                await Task.Delay(300);
                
                // 待機メッセージウィンドウを閉じる
                if (waitWindow != null)
                {
                    waitWindow.Close();
                    waitWindow = null;
                }
                
                await Dispatcher.InvokeAsync(() => CloseLeftoverExamWindows(), DispatcherPriority.Normal);
                
                // 結果画面を閉じる
                System.Diagnostics.Debug.WriteLine("[ResultWindow] Closing ResultWindow");
                this.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error in CloseButton_Click: {ex.Message}\n{ex.StackTrace}");
                
                // 待機メッセージウィンドウを閉じる
                if (waitWindow != null)
                {
                    try { waitWindow.Close(); } catch { }
                }
                
                // エラーが発生した場合でも、ResultWindowを閉じる
                try
                {
                    this.Close();
                }
                catch { }
            }
        }

        private void Header_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == System.Windows.Input.MouseButton.Left)
            {
                this.DragMove();
            }
        }

        private void ShowWrongOnlyButton_Click(object sender, RoutedEventArgs e)
        {
            _showingWrongOnly = !_showingWrongOnly;
            
            if (_showingWrongOnly)
            {
                // ×の問題のみ表示
                var wrongOnlyProjects = FilterWrongTasks(_allProjects);
                ProjectsItemsControl.ItemsSource = wrongOnlyProjects;
                ShowWrongOnlyButton.Content = "全て表示";
                ShowWrongOnlyButton.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(30, 64, 175)); // #1E40AF
            }
            else
            {
                // 全て表示
                ProjectsItemsControl.ItemsSource = _allProjects;
                ShowWrongOnlyButton.Content = "間違えた問題のみ表示";
                ShowWrongOnlyButton.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(220, 38, 38)); // #DC2626
            }
        }

        private List<ResultProjectInfo> FilterWrongTasks(List<ResultProjectInfo> projects)
        {
            if (projects == null) return new List<ResultProjectInfo>();
            
            var filteredProjects = new List<ResultProjectInfo>();
            
            foreach (var project in projects)
            {
                var wrongTasks = project.Tasks?.Where(t => t.ResultMark == "×").ToList();
                
                // ×のタスクがある場合のみプロジェクトを追加
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

        private string PriorityBadgeFor(int projectId, int taskId)
        {
            if (!_isVariantMode)
                return string.Empty;
            return VariantPriorityLookup.GetBadgeText(_variantSetNo, projectId, taskId);
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

        private void UpdateTotalScoreDisplay(Dictionary<int, List<bool>> results)
        {
            if (TotalScorePanel == null || TotalScoreTextBlock == null)
                return;

            if (!_isVariantMode)
            {
                TotalScorePanel.Visibility = Visibility.Collapsed;
                return;
            }

            TotalScorePanel.Visibility = Visibility.Visible;
            if (PracticeScoreClient.TryCompute(VariantExam.FlattenResults(results), out int total, out string error))
                TotalScoreTextBlock.Text = total.ToString();
            else
                TotalScoreTextBlock.Text = error ?? "総合点を表示できません。";
        }

        private string FailReasonFor(int projectId, int taskId)
        {
            if (_isVariantMode)
                return "";
            return MOSExcelMogiApp.Models.ExamResultStorage.GetFailReason(_resultScope, projectId, taskId);
        }

        private void FailReason_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            if (sender is FrameworkElement element && element.DataContext is ResultTaskInfo taskInfo && taskInfo.HasFailReason)
                ScoreReasonWindow.Show(this, taskInfo.TaskId, taskInfo.FailReason);
        }

        private static bool IsFailReasonClick(MouseButtonEventArgs e)
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

        private async void TaskRow_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (IsFailReasonClick(e))
                return;

            if (sender is FrameworkElement element && element.DataContext is ResultTaskInfo taskInfo)
            {
                if (taskInfo.ProjectId > 0 && taskInfo.TaskId > 0)
                {
                    try
                    {
                        // AppBarWindowを確実に表示してからNavigateToTaskを呼び出す
                        var appBarWindow = Application.Current.Windows.OfType<AppBarWindow>().FirstOrDefault();
                        
                        // AppBarWindowが存在しない、または閉じられている場合は新しく作成
                        if (appBarWindow == null || !appBarWindow.IsLoaded)
                        {
                            
                            // MainWindowからViewModelを取得
                            var mainWindow = Application.Current.Windows.OfType<MainWindow>().FirstOrDefault();
                            if (mainWindow == null)
                            {
                                mainWindow = new MainWindow();
                                Application.Current.MainWindow = mainWindow;
                            }
                            
                            var viewModel = mainWindow.DataContext as Ui.ViewModels.MainViewModel;
                            if (viewModel == null)
                            {
                                MessageBox.Show("ViewModelが見つかりませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                            
                            // 新しいAppBarWindowを作成
                            appBarWindow = new AppBarWindow(viewModel);
                        }

                        // 結果画面からの復習時は採点ボタンを強制表示
                        if (appBarWindow.DataContext is Ui.ViewModels.MainViewModel vm)
                        {
                            vm.ShowScoreButton = true;
                        }
                        
                        // 結果画面から来たことを記録
                        appBarWindow.SetFromResultWindow(true);
                        appBarWindow.SetResultWindow(this);
                        appBarWindow.PrepareResultRetry(taskInfo.ProjectId, taskInfo.TaskId, taskInfo.ResultMark == "×");
                        
                        // AppBarWindowを表示
                        if (!appBarWindow.IsVisible)
                        {
                            appBarWindow.Show();
                        }
                        appBarWindow.Activate();
                        
                        // UI更新の機会を与える
                        await Task.Delay(50);
                        // 重要: 表示しているAppBarWindowインスタンスに対して直接ナビゲートする
                        // （OnNavigateToTask は古いAppBarWindowのデリゲートを保持している可能性があるため）
                        var gid = taskInfo.GroupId > 0 ? taskInfo.GroupId : _groupId;
                        appBarWindow.NavigateToTask(taskInfo.ProjectId, taskInfo.TaskId, gid);
                        
                        // NavigateToTaskの処理が完了するまで少し待つ
                        await Task.Delay(100);
                        
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

        /// <summary>結果一覧のコールバックから、×のタスクへ戻るときだけ再挑戦番号を予約する。</summary>
        public void PrepareAppBarForResultRetry(AppBarWindow appBar, int projectId, int taskId)
        {
            if (appBar == null)
                return;
            appBar.PrepareResultRetry(projectId, taskId, IsTaskCurrentlyWrong(projectId, taskId));
        }

        private bool IsTaskCurrentlyWrong(int projectId, int taskId)
        {
            Dictionary<int, List<bool>> results = _allProjectResults;
            if (!_fromScoringLog)
            {
                var stored = MOSExcelMogiApp.Models.ExamResultStorage.GetAllResults(_resultScope);
                if (stored != null && stored.Count > 0)
                    results = stored;
            }
            if (results == null || !results.TryGetValue(projectId, out var list) || list == null)
                return false;
            int index = taskId - 1;
            if (index < 0 || index >= list.Count)
                return false;
            return !list[index];
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
        public int GroupId { get; set; }
        public int ProjectId { get; set; }
        public int TaskId { get; set; }
        public string ResultMark { get; set; }
        public string FailReason { get; set; }
        public bool HasFailReason => !string.IsNullOrEmpty(FailReason);
        public Brush ResultColor { get; set; }
        public string PriorityBadgeText { get; set; }
        public bool HasPriorityBadge => !string.IsNullOrEmpty(PriorityBadgeText);
        public Brush PriorityBadgeBackground { get; set; }
        public Brush PriorityBadgeForeground { get; set; }
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
