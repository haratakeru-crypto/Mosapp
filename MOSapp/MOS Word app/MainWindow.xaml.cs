using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
using Libraries;
using MOS_Word_app.Views;

namespace MOS_Word_app
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;
        private Views.UiTestAppBarWindow _appBarWindow;
        private bool _isExiting;

        /// <summary>タイマー無効化フラグ。デフォルトは一時停止（常時停止）。</summary>
        public static bool IsTimerDisabled { get; private set; } = true;

        public MainWindow()
        {
            InitializeComponent();

            _viewModel = new MainViewModel();
            DataContext = _viewModel;

            _viewModel.ShowAppBarRequested += OnShowAppBarRequested;
            _viewModel.HideMainWindowRequested += OnHideMainWindowRequested;
            _viewModel.ShowMainWindowRequested += OnShowMainWindowRequested;
            _viewModel.ExamEnded += OnExamEnded;

            ApplyExamineeTabState();
            MosPracticeClient.ExamineeStore.EnsureSessionWatch();
            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;
        }

        private void ApplyExamineeTabState()
        {
            ApplyIdConfirmTabHeader();
            _viewModel.SelectedTabIndex = MosPracticeClient.ExamineeStore.IsRegistered
                ? MosPracticeClient.ExamineeStore.PracticeTabIndex
                : 0;
        }

        private void ApplyIdConfirmTabHeader()
        {
            if (IdConfirmationTab == null) return;
            IdConfirmationTab.Header = MosPracticeClient.ExamineeStore.IsCertiportVerified
                ? "ID確認 済"
                : "ID確認";
        }

        private void UniversityRegistration_Registered(object sender, EventArgs e)
        {
            _viewModel.SelectedTabIndex = MosPracticeClient.ExamineeStore.PracticeTabIndex;
        }

        private void UniversityRegistration_Deleted(object sender, EventArgs e)
        {
            WordTaskReviewStateStore.Discard();
            ScoringLogList?.Reload();
            IdConfirmation?.Reload();
            ApplyIdConfirmTabHeader();
            _viewModel.SelectedTabIndex = 0;
        }

        private void IdConfirmation_Verified(object sender, EventArgs e)
        {
            ApplyIdConfirmTabHeader();
        }

        private void VariantPlaceholderButton_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ResultMessage = "類題は後から接続します";
        }

        private void ScoringLog_EntryClicked(object sender, MosPracticeClient.ScoringLogEntry entry)
        {
            if (entry == null) return;
            var window = new Views.ResultWindow(
                entry.ProjectResults,
                entry.GroupId,
                true,
                entry.IsVariantMode,
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
                Cursor = Cursors.Wait;
                CloseAppBarForExit();
                CloseLeftoverReviewPages();
                WordApplicationManager.CloseWordApplicationForAppExit();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[MainWindow] Exit cleanup: " + ex.Message);
            }
            finally
            {
                Cursor = Cursors.Arrow;
            }
        }

        /// <summary>終了時に AppBar を閉じる（Closed で MainWindow を再表示しない）。</summary>
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
            var project = _viewModel.CurrentProject;
            if (project == null)
                return;

            bool needRecreate =
                _appBarWindow == null ||
                !_appBarWindow.IsLoaded ||
                _appBarWindow.CurrentProjectId != project.ProjectId ||
                _appBarWindow.CurrentGroupId != project.GroupId ||
                _appBarWindow.IsVariantMode != _viewModel.IsVariantMode;

            if (needRecreate)
            {
                var oldAppBar = _appBarWindow;
                _appBarWindow = new Views.UiTestAppBarWindow(project.ProjectId, project.GroupId, _viewModel.ShowScoreButton, false, _viewModel.IsVariantMode, _viewModel);
                _appBarWindow.Closed += OnAppBarWindowClosed;
                oldAppBar?.Close();
            }

            if (_appBarWindow != null)
            {
                _appBarWindow.Show();
                _appBarWindow.ApplyExamWindowLayout();
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
            this.Show();
            this.Activate();
        }

        private void OnExamEnded(object sender, EventArgs e)
        {
            _appBarWindow = null;
        }

        private void ProjectResetButton_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel != null && _viewModel.IsBulkResetInProgress)
                return;

            var dialog = new BulkResetScopeDialog { Owner = this };
            if (dialog.ShowDialog() != true)
                return;

            WordBulkResetScope scope = dialog.SelectedScope;
            List<WordResetTarget> targets = _viewModel.GetBulkResetTargets(scope);
            if (targets.Count == 0)
            {
                MessageBox.Show(
                    "リセット対象のプロジェクトが見つかりませんでした。",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            RunBulkReset(scope, targets);
        }

        private void RunBulkReset(WordBulkResetScope scope, List<WordResetTarget> targets)
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
                int groupId = _viewModel.CurrentGroupIdForReset;
                string closeProbePath = targets[0].WorkingPath;

                MainViewModel.CloseAllOpenWordDocuments();
                if (!string.IsNullOrEmpty(closeProbePath) && MainViewModel.IsWordFileLocked(closeProbePath))
                {
                    WordApplicationManager.CloseWordApplicationForAppExit();
                    Thread.Sleep(500);
                }
                else
                {
                    Thread.Sleep(300);
                }

                WordProjectResetHelper.UnblockResetReferenceFolders(
                    groupId,
                    WordBulkResetCoordinator.DefaultVariantSetNo,
                    scope == WordBulkResetScope.Textbook || scope == WordBulkResetScope.All,
                    scope == WordBulkResetScope.Variant || scope == WordBulkResetScope.All);

                WordBulkResetResult result = WordBulkResetCoordinator.CopyTargets(
                    targets,
                    progress =>
                    {
                        textBlock.Text = "リセット中です...\n" + progress;
                        Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Background);
                    });

                var uiTestBar = Application.Current.Windows.OfType<UiTestAppBarWindow>().FirstOrDefault();
                if (uiTestBar != null)
                {
                    uiTestBar.ClearAllProjectStates(
                        clearTextbook: scope == WordBulkResetScope.Textbook || scope == WordBulkResetScope.All,
                        clearVariant: scope == WordBulkResetScope.Variant || scope == WordBulkResetScope.All);
                }
                var appBar = Application.Current.Windows.OfType<AppBarWindow>().FirstOrDefault();
                if (appBar != null)
                    appBar.ClearAllProjectStates();

                waitWindow.Close();
                ShowBulkResetResult(scope, result);
                _viewModel.ResultMessage = result.TotalFail == 0
                    ? "すべてリセットしました。"
                    : "リセットが完了しました（一部失敗あり）。";
            }
            catch (Exception ex)
            {
                try { waitWindow.Close(); } catch { }
                System.Diagnostics.Debug.WriteLine("[RunBulkReset] " + ex.Message);
                MessageBox.Show(
                    $"プロジェクトリセット中にエラーが発生しました: {ex.Message}",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                _viewModel.ResultMessage = $"リセット中にエラー: {ex.Message}";
            }
            finally
            {
                _viewModel.SetBulkResetInProgress(false);
                if (ProjectResetButton != null)
                    ProjectResetButton.IsEnabled = true;
            }
        }

        private static void ShowBulkResetResult(WordBulkResetScope scope, WordBulkResetResult result)
        {
            if (result.TotalFail == 0)
            {
                MessageBox.Show("すべてリセットしました", "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var message = new StringBuilder();
            message.AppendLine("リセットが完了しました。");
            message.AppendLine();
            if (scope == WordBulkResetScope.Textbook || scope == WordBulkResetScope.All)
                message.AppendLine($"演習 成功: {result.TextbookSuccess} / 失敗: {result.TextbookFail}");
            if (scope == WordBulkResetScope.Variant || scope == WordBulkResetScope.All)
                message.AppendLine($"類題 成功: {result.VariantSuccess} / 失敗: {result.VariantFail}");
            if (result.FailedProjects.Count > 0)
            {
                message.AppendLine();
                message.Append("失敗したプロジェクト: ");
                message.Append(string.Join(", ", result.FailedProjects));
            }

            MessageBox.Show(message.ToString(), "完了", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        protected override void OnClosed(EventArgs e)
        {
            if (_viewModel != null)
            {
                _viewModel.ShowAppBarRequested -= OnShowAppBarRequested;
                _viewModel.HideMainWindowRequested -= OnHideMainWindowRequested;
                _viewModel.ShowMainWindowRequested -= OnShowMainWindowRequested;
                _viewModel.ExamEnded -= OnExamEnded;
            }
            if (!_isExiting)
                CloseAppBarForExit();
            base.OnClosed(e);
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 起動スプラッシュ（Topmost）閉鎖後に確認を出す
            Dispatcher.BeginInvoke(new Action(() =>
            {
                TryPromptStartupReset();
                TryAutoOpenProject();
            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }

        private void TryPromptStartupReset()
        {
            if (!MosPracticeClient.StartupResetPrompt.Ask(this))
                return;

            if (_viewModel != null && _viewModel.IsBulkResetInProgress)
                return;

            List<WordResetTarget> targets = _viewModel.GetBulkResetTargets(WordBulkResetScope.All);
            if (targets.Count == 0)
            {
                MessageBox.Show(
                    "リセット対象のプロジェクトが見つかりませんでした。",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            RunBulkReset(WordBulkResetScope.All, targets);
        }

        private void TryAutoOpenProject()
        {
            if (!App.AutoOpenGroupId.HasValue || !App.AutoOpenProjectId.HasValue)
                return;

            int groupId = App.AutoOpenGroupId.Value;
            int projectId = App.AutoOpenProjectId.Value;
            App.ClearAutoOpen();

            var group = _viewModel.ProjectGroups?.FirstOrDefault(g => g.GroupId == groupId);
            var project = group?.Projects?.FirstOrDefault(p => p.ProjectId == projectId);
            if (project == null)
                return;

            if (_viewModel.OpenProjectCommand.CanExecute(project))
                _viewModel.OpenProjectCommand.Execute(project);
        }
    }
}
