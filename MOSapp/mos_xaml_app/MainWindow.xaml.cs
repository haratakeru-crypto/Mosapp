using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Ui.ViewModels;
using Core.Adapters;
using Infrastructure;
using System.IO;
using System.Runtime.InteropServices;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using Libraries;
using MOSExcelMogiApp.Infrastructure;
using MOSExcelMogiApp.Models;
using MOSExcelMogiApp.Views;

namespace MOSExcelMogiApp
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;
        private AppBarWindow _appBarWindow;
        private bool _isExiting;
        public static bool IsTimerDisabled { get; private set; } = true; // タイマー無効化フラグ（静的プロパティ）。デフォルトは一時停止。
        
        public MainWindow()
        {
            // アプリ起動時にアドインの診断ログをクリアして肥大化を防ぐ
            ExcelLogReader.ClearDiagnosticLog();
            ExcelLogReader.EnsureLogSchemaV2();

            InitializeComponent();
            
            // Dependency Injection setup
            var repository = new ExcelCheckerRepository();
            var service = new ExcelCheckerService(repository);
            _viewModel = new MainViewModel(service);
            
            DataContext = _viewModel;
            
            // ViewModelのイベントを購読
            _viewModel.ShowAppBarRequested += OnShowAppBarRequested;
            _viewModel.HideMainWindowRequested += OnHideMainWindowRequested;
            _viewModel.ShowMainWindowRequested += OnShowMainWindowRequested;
            _viewModel.ExamEnded += OnExamEnded;

            ApplyExamineeTabState();
            MosPracticeClient.ExamineeStore.EnsureSessionWatch();
            Closing += MainWindow_Closing;
        }

        private void ApplyExamineeTabState()
        {
            _viewModel.SelectedTabIndex = MosPracticeClient.ExamineeStore.IsRegistered ? 1 : 0;
        }

        private void UniversityRegistration_Registered(object sender, EventArgs e)
        {
            _viewModel.SelectedTabIndex = 1;
        }

        private void UniversityRegistration_Deleted(object sender, EventArgs e)
        {
            ScoringLogList?.Reload();
            _viewModel.SelectedTabIndex = 0;
        }

        private void VocabularyTab_StartRequested(object sender, MosPracticeClient.VocabularyStartEventArgs e)
        {
            if (e.Mode != MosPracticeClient.VocabularyStartMode.KeywordOnly)
            {
                MessageBox.Show("「問題文から」モードは準備中です。", "単語帳", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            _viewModel.StartVocabularySession(e.Category);
        }

        private void VocabularyTab_SettingsRequested(object sender, EventArgs e)
        {
            _viewModel.StartVocabularySettings();
        }

        private void ScoringLog_EntryClicked(object sender, MosPracticeClient.ScoringLogEntry entry)
        {
            if (entry == null) return;
            var window = new Views.ResultWindow(
                entry.ProjectResults,
                entry.GroupId,
                fromScoringLog: true,
                isVariantMode: entry.IsVariantMode,
                variantSetNo: entry.VariantSetNo,
                completedStates: entry.GetCompletedStates(),
                flaggedStates: entry.GetFlaggedStates());
            window.Owner = this;
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.Show();
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_isExiting)
                return;

            var result = MessageBox.Show("アプリ自体を終了します。本当にいいですか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
            {
                e.Cancel = true;
                return;
            }

            _isExiting = true;
            try
            {
                // 終了処理中にカーソルを待機状態にする
                this.Cursor = Cursors.Wait;

                // AppBar / 隠れた Review が残るとプロセスが残ることがある
                CloseAppBarForExit();
                CloseLeftoverReviewPages();

                // 試験終了スレッドが動いていれば、Kill 完了まで短時間待つ（即アプリ落とし対策）
                try
                {
                    _viewModel?.WaitForExcelShutdownToCompleteBeforeOpeningProject();
                }
                catch (Exception waitEx)
                {
                    System.Diagnostics.Debug.WriteLine($"[MainWindow] Wait shutdown: {waitEx.Message}");
                }

                // Excel を確実に閉じる（同期実行して完了を待つことでゾンビプロセスを防止）
                _viewModel?.CloseExcelApplication();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Error during closing cleanup: {ex.Message}");
            }
            finally
            {
                this.Cursor = Cursors.Arrow;
            }
        }

        /// <summary>終了時に AppBar を閉じる。</summary>
        private void CloseAppBarForExit()
        {
            var bar = _appBarWindow;
            _appBarWindow = null;
            if (bar == null)
                return;
            try { bar.Closed -= OnAppBarWindowClosed; } catch { /* ignore */ }
            try { bar.Close(); } catch { /* ignore */ }
        }

        private static void CloseLeftoverReviewPages()
        {
            foreach (var review in Application.Current.Windows.OfType<Views.ReviewPageWindow>().ToList())
            {
                try { review.Close(); } catch { /* ignore */ }
            }
        }
        
        private void OnShowAppBarRequested(object sender, EventArgs e)
        {
            if (_isExiting)
                return;

            if (_appBarWindow == null || !_appBarWindow.IsLoaded)
            {
                _appBarWindow = new AppBarWindow(_viewModel);
                _appBarWindow.Closed += OnAppBarWindowClosed;
                // Excel の前面化を優先するため、表示時にフォーカスを奪わない。
                _appBarWindow.ShowActivated = false;
            }

            _viewModel.VocabularyUiAnchorProvider = delegate (string which, out IntPtr hwnd, out Rect rect)
            {
                if (_appBarWindow != null && _appBarWindow.TryGetVocabularyAnchor(which, out hwnd, out rect))
                    return true;
                hwnd = IntPtr.Zero;
                rect = Rect.Empty;
                return false;
            };

            if (!_appBarWindow.IsVisible)
            {
                _appBarWindow.Show();
            }
        }
        
        private void OnAppBarWindowClosed(object sender, EventArgs e)
        {
            if (!ReferenceEquals(sender, _appBarWindow))
                return;
            _appBarWindow = null;
            if (_isExiting)
                return;
            _viewModel?.EnableProjectSelection();
            this.Show();
            this.Activate();
        }

        private void OnHideMainWindowRequested(object sender, EventArgs e)
        {
            this.Hide();
        }
        
        private void OnShowMainWindowRequested(object sender, EventArgs e)
        {
            _viewModel?.EnableProjectSelection();
            this.Show();
            this.Activate();
        }
        
        private void OnExamEnded(object sender, EventArgs e)
        {
            // アプリバーの参照をクリアして、次回新しいインスタンスを作成できるようにする
            _appBarWindow = null;
        }
        
        private void ProjectResetButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[ResetProject] ProjectResetButton_Click called");
            try
            {
                if (_viewModel != null && _viewModel.IsBulkResetInProgress)
                    return;

                var dialog = new BulkResetScopeDialog
                {
                    Owner = this
                };
                if (dialog.ShowDialog() != true)
                {
                    System.Diagnostics.Debug.WriteLine("[ResetProject] Reset cancelled by user");
                    return;
                }

                ExcelBulkResetScope scope = dialog.SelectedScope;
                var targets = _viewModel.GetBulkResetTargets(scope);
                if (targets.Count == 0)
                {
                    MessageBox.Show("リセット対象のプロジェクトが見つかりませんでした。",
                        "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                RunBulkReset(scope, targets);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResetProject] Error in ProjectResetButton_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"プロジェクトリセット中にエラーが発生しました: {ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RunBulkReset(ExcelBulkResetScope scope, List<ExcelResetTarget> targets)
        {
            var waitWindow = new Window
            {
                Title = "リセット中",
                Width = 320,
                Height = 160,
                WindowStyle = WindowStyle.ToolWindow,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ShowInTaskbar = false,
                ResizeMode = ResizeMode.NoResize
            };

            var stackPanel = new StackPanel
            {
                Margin = new Thickness(20),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            var textBlock = new TextBlock
            {
                Text = "リセット中です...",
                FontSize = 16,
                HorizontalAlignment = HorizontalAlignment.Center,
                TextAlignment = TextAlignment.Center,
                Margin = new Thickness(0, 0, 0, 10)
            };
            var progressBar = new ProgressBar
            {
                IsIndeterminate = true,
                Height = 20,
                Width = 250
            };
            stackPanel.Children.Add(textBlock);
            stackPanel.Children.Add(progressBar);
            waitWindow.Content = stackPanel;

            _viewModel.SetBulkResetInProgress(true);
            if (ProjectResetButton != null)
                ProjectResetButton.IsEnabled = false;
            waitWindow.Show();
            try
            {
                ExcelResetTarget currentOpen = _viewModel.TryGetCurrentOpenResetTarget(targets);
                string reopenPath = null;
                int reopenProjectId = 0;
                if (_viewModel.TryGetCurrentGroupProjectId(out int currentGroupId, out int currentProjectId))
                {
                    reopenPath = _viewModel.GetActiveProjectFilePath(currentGroupId, currentProjectId);
                    reopenProjectId = currentProjectId;
                }

                string closeProbePath = currentOpen?.WorkingPath ?? reopenPath ?? targets[0].WorkingPath;
                var closeSw = new System.Diagnostics.Stopwatch();
                bool keepExcel = CloseExcelForReset(closeProbePath, closeSw, out _, out _);

                ExcelBulkResetResult result = ExcelBulkResetCoordinator.CopyTargets(
                    targets,
                    currentOpen,
                    progress =>
                    {
                        textBlock.Text = "リセット中です...\n" + progress;
                        Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);
                    });

                if (scope == ExcelBulkResetScope.Textbook || scope == ExcelBulkResetScope.All)
                    TaskReviewStateStore.Clear(ExamResultScope.Textbook);
                if (scope == ExcelBulkResetScope.Variant || scope == ExcelBulkResetScope.All)
                    TaskReviewStateStore.Clear(ExamResultScope.Variant(ExcelBulkResetCoordinator.DefaultVariantSetNo));

                if (result.ReopenTarget != null)
                {
                    reopenPath = result.ReopenTarget.WorkingPath;
                    reopenProjectId = result.ReopenTarget.ProjectId;
                }
                else if (currentOpen != null)
                {
                    reopenPath = null;
                }

                if (!string.IsNullOrEmpty(reopenPath) && File.Exists(reopenPath))
                {
                    bool reopenLogged = false;
                    ReopenExcelAfterReset(
                        reopenPath,
                        reopenProjectId,
                        keepExcel,
                        logResetPerf: false,
                        detailSuffix: " bulk",
                        ref reopenLogged);
                }

                waitWindow.Close();
                ShowBulkResetResult(scope, result);
            }
            catch (Exception ex)
            {
                try { waitWindow.Close(); } catch { }
                System.Diagnostics.Debug.WriteLine($"[ResetProject] Error in bulk reset: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"プロジェクトリセット中にエラーが発生しました: {ex.Message}",
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _viewModel.SetBulkResetInProgress(false);
                if (ProjectResetButton != null)
                    ProjectResetButton.IsEnabled = true;
            }
        }

        private static void ShowBulkResetResult(ExcelBulkResetScope scope, ExcelBulkResetResult result)
        {
            if (result.TotalFail == 0)
            {
                MessageBox.Show("すべてリセットしました", "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var message = new StringBuilder();
            message.AppendLine("リセットが完了しました。");
            message.AppendLine();
            if (scope == ExcelBulkResetScope.Textbook || scope == ExcelBulkResetScope.All)
            {
                message.AppendLine($"演習 成功: {result.TextbookSuccess} / 失敗: {result.TextbookFail}");
            }
            if (scope == ExcelBulkResetScope.Variant || scope == ExcelBulkResetScope.All)
            {
                message.AppendLine($"類題 成功: {result.VariantSuccess} / 失敗: {result.VariantFail}");
            }
            if (result.FailedProjects.Count > 0)
            {
                message.AppendLine();
                message.Append("失敗したプロジェクト: ");
                message.Append(string.Join(", ", result.FailedProjects));
            }

            MessageBox.Show(message.ToString(), "完了", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        
        public void ResetProject(int groupId, int projectId, bool showMessage = true, bool logResetPerf = false)
        {
            if (logResetPerf)
                ResetPerfLog.Begin("excel", projectId);
            bool closeLogged = false;
            bool copyLogged = false;
            bool reopenLogged = false;
            var closeSw = new System.Diagnostics.Stopwatch();
            var copySw = new System.Diagnostics.Stopwatch();
            try
            {
                // 正規作業ファイルを最優先し、旧 config/Initial 配置は移行元に限定する。
                string projectFilePath = _viewModel?.GetProjectFilePath(groupId, projectId)
                    ?? DataPathHelper.ResolveWorkingFilePath(groupId, projectId);
                
                if (string.IsNullOrEmpty(projectFilePath))
                {
                    string errorMsg = $"プロジェクトファイルのパスを取得できませんでした。\nプロジェクト: Group{groupId}, Project{projectId}";
                    if (showMessage)
                    {
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    throw new Exception(errorMsg);
                }
                
                System.Diagnostics.Debug.WriteLine($"[ResetProject] Resetting project file: {projectFilePath}");
                
                string templatesFolder = System.IO.Path.GetDirectoryName(
                    DataPathHelper.GetTemplateFilePath(groupId, projectId));
                string templatePath = DataPathHelper.ResolveTemplateFilePath(groupId, projectId);
                
                if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
                {
                    string errorMsg = $"テンプレートファイルが見つかりません。\n\nプロジェクト: Group{groupId}, Project{projectId}\n検索フォルダ: {templatesFolder}";
                    
                    if (Directory.Exists(templatesFolder))
                    {
                        var files = Directory.GetFiles(templatesFolder, "*.xlsx");
                        errorMsg += $"\n\nフォルダ内のファイル:\n";
                        foreach (var file in files)
                        {
                            errorMsg += $"  - {System.IO.Path.GetFileName(file)}\n";
                        }
                    }
                    
                    if (showMessage)
                    {
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    throw new Exception(errorMsg);
                }
                
                System.Diagnostics.Debug.WriteLine($"Template file found: {templatePath}");

                System.Diagnostics.Debug.WriteLine($"[ResetProject] Closing workbooks without quitting Excel");
                bool keepExcel = CloseExcelForReset(projectFilePath, closeSw, out string closePath, out string closeDetail);
                if (logResetPerf)
                {
                    ResetPerfLog.Write("excel", projectId, "close", closeSw.ElapsedMilliseconds, closePath, closeDetail);
                    closeLogged = true;
                }

                try
                {
                    copySw.Start();
                    string initialFilePath = DataPathHelper.GetInitialFilePath(groupId, projectId);
                    int retryCount = ExcelProjectResetHelper.RestoreFromTemplate(templatePath, projectFilePath, initialFilePath);
                    copySw.Stop();
                    ExcelProjectResetHelper.ClearLogsAfterSuccessfulCopy(projectId, 0);
                    TaskReviewStateStore.ClearProject(ExamResultScope.Textbook, projectId);
                    if (logResetPerf)
                    {
                        ResetPerfLog.Write("excel", projectId, "copy", copySw.ElapsedMilliseconds, "main", "lockRetries=" + retryCount);
                        copyLogged = true;
                    }

                    if (_viewModel?.CurrentProject != null)
                    {
                        var currentProject = _viewModel.CurrentProject;
                        int currentGroupId = int.Parse(currentProject.Group.Replace("Group ", ""));
                        int currentProjectId = currentProject.ProjectNumber;

                        if (currentGroupId == groupId && currentProjectId == projectId && !_viewModel.IsVariantMode)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Processing Excel file after reset: {projectFilePath}");
                            ReopenExcelAfterReset(projectFilePath, projectId, keepExcel, logResetPerf, null, ref reopenLogged);
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Reopened workbook: {projectFilePath}");
                        }
                    }
                    if (logResetPerf && !reopenLogged)
                    {
                        ResetPerfLog.Write("excel", projectId, "reopen", 0, "main", "skipped");
                        ResetPerfLog.Write("excel", projectId, "ready", 0, "main", "skipped");
                        reopenLogged = true;
                    }
                    
                    if (showMessage)
                    {
                        MessageBox.Show($"プロジェクト (Group{groupId}, Project{projectId}) をリセットしました。", 
                            "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    string errorMsg = $"ファイルへのアクセスが拒否されました。\nファイルが他のプログラムで開かれている可能性があります。\n\nプロジェクト: Group{groupId}, Project{projectId}\n{ex.Message}";
                    if (showMessage)
                    {
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    System.Diagnostics.Debug.WriteLine($"UnauthorizedAccessException in ResetProject: {ex.Message}");
                    throw new Exception(errorMsg, ex);
                }
                catch (IOException ex)
                {
                    string errorMsg = $"ファイルのコピー中にエラーが発生しました。\nファイルがロックされている可能性があります。\n\nプロジェクト: Group{groupId}, Project{projectId}\n{ex.Message}";
                    if (showMessage)
                    {
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    System.Diagnostics.Debug.WriteLine($"IOException in ResetProject: {ex.Message}");
                    throw new Exception(errorMsg, ex);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"プロジェクトリセット中にエラーが発生しました。\nプロジェクト: Group{groupId}, Project{projectId}\n{ex.Message}";
                if (showMessage)
                {
                    MessageBox.Show($"{errorMsg}\n\n{ex.StackTrace}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                System.Diagnostics.Debug.WriteLine($"Error in ResetProject: {ex.Message}\n{ex.StackTrace}");
                if (logResetPerf)
                    LogExcelResetFailure(projectId, closeLogged, copyLogged, reopenLogged, closeSw, copySw);
                throw;
            }
        }

        bool CloseExcelForReset(
            string projectFilePath,
            System.Diagnostics.Stopwatch closeSw,
            out string closePath,
            out string closeDetail)
        {
            closeSw.Start();
            bool keepExcel = _viewModel != null && _viewModel.TryCloseWorkbooksKeepingExcel();
            if (keepExcel)
            {
                for (int attempt = 0; attempt < 3 && IsFileLocked(projectFilePath); attempt++)
                    System.Threading.Thread.Sleep(200);

                if (IsFileLocked(projectFilePath))
                {
                    _viewModel.QuitExcelForProjectReset();
                    keepExcel = false;
                    closePath = "fallback";
                    closeDetail = "quit-process";
                }
                else
                {
                    closePath = "main";
                    closeDetail = "keep-process";
                }
            }
            else
            {
                closePath = "main";
                closeDetail = "no-process";
            }

            closeSw.Stop();
            return keepExcel;
        }

        void ReopenExcelAfterReset(
            string projectFilePath,
            int projectId,
            bool keepExcel,
            bool logResetPerf,
            string detailSuffix,
            ref bool reopenLogged)
        {
            var reopenSw = System.Diagnostics.Stopwatch.StartNew();
            bool opened = keepExcel && _viewModel.TryOpenWorkbookInRunningExcel(projectFilePath);
            long reopenMs = reopenSw.ElapsedMilliseconds;
            long readyMs = 0;
            bool ready = false;
            if (opened)
            {
                var readySw = System.Diagnostics.Stopwatch.StartNew();
                ready = _viewModel.WaitUntilResetWorkbookReady(projectFilePath, 2000);
                readyMs = readySw.ElapsedMilliseconds;
            }

            if (opened && ready)
            {
                if (logResetPerf)
                {
                    ResetPerfLog.Write("excel", projectId, "reopen", reopenMs, "main", "com" + detailSuffix);
                    ResetPerfLog.Write(
                        "excel",
                        projectId,
                        "ready",
                        readyMs,
                        "main",
                        "signal=vsto-startup result=ok" + detailSuffix);
                    reopenLogged = true;
                }
                return;
            }

            if (keepExcel)
                _viewModel.QuitExcelForProjectReset();

            _viewModel.OpenExcelWorkbookAfterResetByShell(projectFilePath);
            if (logResetPerf)
            {
                string path = keepExcel ? "fallback" : "main";
                string detail = (keepExcel ? "quit-process shell" : "shell") + detailSuffix;
                ResetPerfLog.Write("excel", projectId, "reopen", reopenSw.ElapsedMilliseconds, path, detail);
                reopenLogged = true;
                ExcelResetReadyWatch.Start(projectId, projectFilePath, path);
            }
        }

        static void LogExcelResetFailure(
            int projectId,
            bool closeLogged,
            bool copyLogged,
            bool reopenLogged,
            System.Diagnostics.Stopwatch closeSw,
            System.Diagnostics.Stopwatch copySw)
        {
            if (!closeLogged)
                ResetPerfLog.Write("excel", projectId, "close", closeSw.ElapsedMilliseconds, "main", "result=fail");
            else if (!copyLogged)
                ResetPerfLog.Write("excel", projectId, "copy", copySw.ElapsedMilliseconds, "main", "result=fail");
            else if (!reopenLogged)
                ResetPerfLog.Write("excel", projectId, "reopen", 0, "main", "result=fail");
        }

        /// <summary>
        /// 類題モード用リセット: PracticeVariant Templates → 作業用 xlsx に復元し Excel を開き直す（教材 Initial は更新しない）。
        /// </summary>
        public void ResetVariantProject(int groupId, int projectId, int variantSetNo, bool showMessage = true, bool logResetPerf = false)
        {
            if (logResetPerf)
                ResetPerfLog.Begin("excel", projectId);
            bool closeLogged = false;
            bool copyLogged = false;
            bool reopenLogged = false;
            var closeSw = new System.Diagnostics.Stopwatch();
            var copySw = new System.Diagnostics.Stopwatch();
            try
            {
                string projectFilePath = _viewModel?.GetVariantWorkingFilePath(groupId, projectId, variantSetNo);
                if (string.IsNullOrEmpty(projectFilePath) && _viewModel?.CurrentProject != null)
                    projectFilePath = _viewModel.CurrentProject.FilePath;

                if (string.IsNullOrEmpty(projectFilePath))
                {
                    string errorMsg = $"類題ファイルのパスを取得できませんでした。\n類題{variantSetNo} プロジェクト: Group{groupId}, Project{projectId}";
                    if (showMessage)
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    throw new Exception(errorMsg);
                }

                string templatePath = _viewModel.GetVariantTemplateFilePath(groupId, projectId, variantSetNo);
                if (!File.Exists(templatePath))
                {
                    string errorMsg = $"類題テンプレートファイルが見つかりません。\n\n類題{variantSetNo} プロジェクト: Group{groupId}, Project{projectId}\n{templatePath}";
                    if (showMessage)
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    throw new Exception(errorMsg);
                }

                System.Diagnostics.Debug.WriteLine($"[ResetVariantProject] Template: {templatePath}");
                System.Diagnostics.Debug.WriteLine($"[ResetVariantProject] Working file: {projectFilePath}");

                System.Diagnostics.Debug.WriteLine($"[ResetVariantProject] Closing workbooks without quitting Excel");
                bool keepExcel = CloseExcelForReset(projectFilePath, closeSw, out string closePath, out string closeDetail);
                if (logResetPerf)
                {
                    ResetPerfLog.Write(
                        "excel",
                        projectId,
                        "close",
                        closeSw.ElapsedMilliseconds,
                        closePath,
                        closeDetail + " variant=" + variantSetNo);
                    closeLogged = true;
                }

                try
                {
                    copySw.Start();
                    int retryCount = ExcelProjectResetHelper.RestoreFromTemplate(templatePath, projectFilePath);
                    copySw.Stop();
                    ExcelProjectResetHelper.ClearLogsAfterSuccessfulCopy(projectId, variantSetNo);
                    TaskReviewStateStore.ClearProject(ExamResultScope.Variant(variantSetNo), projectId);
                    if (logResetPerf)
                    {
                        ResetPerfLog.Write("excel", projectId, "copy", copySw.ElapsedMilliseconds, "main", "variant=" + variantSetNo + " lockRetries=" + retryCount);
                        copyLogged = true;
                    }

                    if (_viewModel?.CurrentProject != null)
                    {
                        var currentProject = _viewModel.CurrentProject;
                        int currentGroupId = int.Parse(currentProject.Group.Replace("Group ", ""));
                        int currentProjectId = currentProject.ProjectNumber;

                        if (currentGroupId == groupId && currentProjectId == projectId && _viewModel.IsVariantMode)
                        {
                            ReopenExcelAfterReset(
                                projectFilePath,
                                projectId,
                                keepExcel,
                                logResetPerf,
                                " variant=" + variantSetNo,
                                ref reopenLogged);
                            System.Diagnostics.Debug.WriteLine($"[ResetVariantProject] Reopened workbook: {projectFilePath}");
                        }
                    }
                    if (logResetPerf && !reopenLogged)
                    {
                        ResetPerfLog.Write("excel", projectId, "reopen", 0, "main", "skipped");
                        ResetPerfLog.Write("excel", projectId, "ready", 0, "main", "skipped");
                        reopenLogged = true;
                    }

                    if (showMessage)
                    {
                        MessageBox.Show($"類題{variantSetNo}（プロジェクト {groupId}-{projectId}）をリセットしました。",
                            "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    string errorMsg = $"ファイルへのアクセスが拒否されました。\nファイルが他のプログラムで開かれている可能性があります。\n\n類題{variantSetNo} プロジェクト: Group{groupId}, Project{projectId}\n{ex.Message}";
                    if (showMessage)
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    System.Diagnostics.Debug.WriteLine($"UnauthorizedAccessException in ResetVariantProject: {ex.Message}");
                    throw new Exception(errorMsg, ex);
                }
                catch (IOException ex)
                {
                    string errorMsg = $"ファイルのコピー中にエラーが発生しました。\nファイルがロックされている可能性があります。\n\n類題{variantSetNo} プロジェクト: Group{groupId}, Project{projectId}\n{ex.Message}";
                    if (showMessage)
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    System.Diagnostics.Debug.WriteLine($"IOException in ResetVariantProject: {ex.Message}");
                    throw new Exception(errorMsg, ex);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"類題リセット中にエラーが発生しました。\n類題{variantSetNo} プロジェクト: Group{groupId}, Project{projectId}\n{ex.Message}";
                if (showMessage)
                    MessageBox.Show($"{errorMsg}\n\n{ex.StackTrace}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"Error in ResetVariantProject: {ex.Message}\n{ex.StackTrace}");
                if (logResetPerf)
                    LogExcelResetFailure(projectId, closeLogged, copyLogged, reopenLogged, closeSw, copySw);
                throw;
            }
        }
        
        private static void RemoveZoneIdentifier(string filePath)
        {
            ExcelProjectResetHelper.RemoveZoneIdentifier(filePath);
        }

        private bool IsFileLocked(string filePath)
        {
            return ExcelProjectResetHelper.IsFileLocked(filePath);
        }
        
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
                    // Excelが開いていない場合は無視
                    System.Diagnostics.Debug.WriteLine("[MainWindow] Excel application not found");
                    return;
                }
                
                if (excelApp != null && excelApp.Workbooks != null)
                {
                    string fileName = System.IO.Path.GetFileName(filePath);
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                                wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                            {
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Closing workbook: {wb.Name}");
                                wb.Close(SaveChanges: false);
                                Marshal.ReleaseComObject(wb);
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[MainWindow] Error accessing workbook: {ex.Message}");
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
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Error closing Excel file: {ex.Message}");
            }
        }
        
        protected override void OnClosed(EventArgs e)
        {
            // イベント購読を解除
            if (_viewModel != null)
            {
                _viewModel.ShowAppBarRequested -= OnShowAppBarRequested;
                _viewModel.HideMainWindowRequested -= OnHideMainWindowRequested;
                _viewModel.ShowMainWindowRequested -= OnShowMainWindowRequested;
                _viewModel.ExamEnded -= OnExamEnded;
            }
            
            // アプリバーウィンドウを閉じる（Closing で済んでいれば no-op）
            if (!_isExiting)
                CloseAppBarForExit();
            else
                _appBarWindow = null;
            
            base.OnClosed(e);
        }
    }
}
