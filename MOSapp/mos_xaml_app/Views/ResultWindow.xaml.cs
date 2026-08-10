using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using Newtonsoft.Json;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Threading;
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

        public ResultWindow(Dictionary<int, List<bool>> allProjectResults = null, int groupId = 1)
        {
            InitializeComponent();
            _allProjectResults = allProjectResults ?? new Dictionary<int, List<bool>>();
            _groupId = groupId;
            // 採点直後のスナップショットを取得（存在しない場合は null または空のディクショナリ）
            _initialProjectResults = MOSExcelMogiApp.Models.ExamResultStorage.GetInitialResults();
            System.Diagnostics.Debug.WriteLine($"[ResultWindow] Constructor called with {_allProjectResults?.Count ?? 0} projects, groupId: {_groupId}");
            
            // ウィンドウが読み込まれた後にデータを読み込む（非同期）
            this.Loaded += ResultWindow_Loaded;

            // 復習採点で結果が更新されたら即時反映
            MOSExcelMogiApp.Models.ExamResultStorage.ResultsChanged += OnResultsChanged;
        }

        private async void OnResultsChanged(int projectId)
        {
            try
            {
                // UIスレッドで再計算・再描画（連続発火に備えて少しだけ遅延）
                await Dispatcher.InvokeAsync(async () =>
                {
                    await Task.Delay(20);

                    var currentResults = MOSExcelMogiApp.Models.ExamResultStorage.GetAllResults();
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
                int _ = 0;
                int __ = 0;
                CalculateRate(initialResults, out initialCorrectRate, out _, out __);

                // 現在（復習後）の正答率を計算（保存されている最新結果を使用）
                var currentResults = MOSExcelMogiApp.Models.ExamResultStorage.GetAllResults();
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
                });

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
            });
            
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
                                    ResultColor = isCorrect.HasValue 
                                        ? (isCorrect.Value ? Brushes.Green : Brushes.Red)
                                        : Brushes.Gray
                                };
                            }).ToList() ?? new List<ResultTaskInfo>()
                        };
                        
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
                            Application.Current.MainWindow = mainWindow;
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
                            }
                        }
                        
                        // MainWindowを確実に表示
                        System.Diagnostics.Debug.WriteLine("[ResultWindow] Showing MainWindow");
                        mainWindow.Show();
                        mainWindow.WindowState = WindowState.Normal;
                        mainWindow.Activate();
                        mainWindow.Focus();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error showing MainWindow: {ex.Message}");
                        // エラーが発生した場合でも、新しいMainWindowを作成
                        try
                        {
                            mainWindow = new MOSExcelMogiApp.MainWindow();
                            Application.Current.MainWindow = mainWindow;
                            mainWindow.Show();
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
                
                // AppBarWindowを閉じる
                await Dispatcher.InvokeAsync(() =>
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine("[ResultWindow] Closing AppBarWindows");
                        var windowsToClose = new List<Window>();
                        foreach (Window window in Application.Current.Windows)
                        {
                            if (window.GetType().Name == "AppBarWindow" || window.GetType().Name == "UiTestAppBarWindow")
                            {
                                windowsToClose.Add(window);
                            }
                        }
                        
                        foreach (var window in windowsToClose)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ResultWindow] Closing window: {window.GetType().Name}");
                            window.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error closing AppBarWindows: {ex.Message}");
                    }
                }, DispatcherPriority.Normal);
                
                // ReviewPageWindowを閉じる処理は削除
                // （非表示になっているため、ユーザーには見えず、アプリケーション終了時に自動的に閉じられる）
                
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
                            Application.Current.MainWindow = mainWindow;
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
                            }
                        }
                        
                        // MainWindowを確実に表示
                        System.Diagnostics.Debug.WriteLine("[ResultWindow] Showing MainWindow");
                        mainWindow.Show();
                        mainWindow.WindowState = WindowState.Normal;
                        mainWindow.Activate();
                        mainWindow.Focus();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error showing MainWindow: {ex.Message}");
                        // エラーが発生した場合でも、新しいMainWindowを作成
                        try
                        {
                            mainWindow = new MOSExcelMogiApp.MainWindow();
                            Application.Current.MainWindow = mainWindow;
                            mainWindow.Show();
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
                
                // AppBarWindowを閉じる
                await Dispatcher.InvokeAsync(() =>
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine("[ResultWindow] Closing AppBarWindows");
                        var windowsToClose = new List<Window>();
                        foreach (Window window in Application.Current.Windows)
                        {
                            if (window.GetType().Name == "AppBarWindow" || window.GetType().Name == "UiTestAppBarWindow")
                            {
                                windowsToClose.Add(window);
                            }
                        }
                        
                        foreach (var window in windowsToClose)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ResultWindow] Closing window: {window.GetType().Name}");
                            window.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResultWindow] Error closing AppBarWindows: {ex.Message}");
                    }
                }, DispatcherPriority.Normal);
                
                // ReviewPageWindowを閉じる処理は削除
                // （非表示になっているため、ユーザーには見えず、アプリケーション終了時に自動的に閉じられる）
                
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

        private async void TaskRow_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
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
        public Brush ResultColor { get; set; }
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
