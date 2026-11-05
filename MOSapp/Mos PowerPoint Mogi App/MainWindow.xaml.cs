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
using Libraries;

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;
        private Views.UiTestAppBarWindow _appBarWindow;
        private bool _isExiting;

        // タイマー無効化フラグ（静的プロパティ）。デフォルトは一時停止。
        public static bool IsTimerDisabled { get; private set; } = true;

        public MainWindow()
        {
            try
            {
                InitializeComponent();

                _viewModel = new MainViewModel();
                DataContext = _viewModel;

                // ViewModelのイベントを購読
                _viewModel.ShowAppBarRequested += OnShowAppBarRequested;
                _viewModel.HideMainWindowRequested += OnHideMainWindowRequested;
                _viewModel.ShowMainWindowRequested += OnShowMainWindowRequested;
                _viewModel.ExamEnded += OnExamEnded;
                _viewModel.ScoreCompleted += OnScoreCompleted;

                ApplyExamineeTabState();
                MosPracticeClient.ExamineeStore.EnsureSessionWatch();
                Loaded += MainWindow_Loaded;
                Closing += MainWindow_Closing;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"MainWindowの初期化中にエラーが発生しました:\n\n{ex.Message}\n\nスタックトレース:\n{ex.StackTrace}",
                    "初期化エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                System.Diagnostics.Debug.WriteLine($"MainWindow初期化エラー: {ex}");
                throw;
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= MainWindow_Loaded;
            // 起動スプラッシュ（Topmost）閉鎖後に確認を出す
            Dispatcher.BeginInvoke(
                new Action(TryPromptStartupReset),
                System.Windows.Threading.DispatcherPriority.ApplicationIdle);
        }

        private void TryPromptStartupReset()
        {
            if (!MosPracticeClient.StartupResetPrompt.Ask(this))
                return;

            RunBulkReset(Libraries.PowerPointBulkResetScope.All);
        }

        private void ProjectResetButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Views.BulkResetScopeDialog { Owner = this };
            if (dialog.ShowDialog() != true)
                return;

            RunBulkReset(dialog.SelectedScope);
        }

        private void RunBulkReset(Libraries.PowerPointBulkResetScope scope)
        {
            if (_viewModel == null)
                return;

            var targets = _viewModel.GetBulkResetTargets(scope);
            if (targets.Count == 0)
            {
                MessageBox.Show(
                    "リセット対象のプロジェクトが見つかりませんでした。",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            if (ProjectResetButton != null)
                ProjectResetButton.IsEnabled = false;
            try
            {
                Libraries.PowerPointBulkResetResult result = _viewModel.RunBulkReset(scope);
                ShowBulkResetResult(scope, result);
                _viewModel.ResultMessage = result.TotalFail == 0
                    ? "すべてリセットしました。"
                    : "リセットが完了しました（一部失敗あり）。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"プロジェクトリセット中にエラーが発生しました: {ex.Message}",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                if (ProjectResetButton != null)
                    ProjectResetButton.IsEnabled = true;
            }
        }

        private static void ShowBulkResetResult(
            Libraries.PowerPointBulkResetScope scope,
            Libraries.PowerPointBulkResetResult result)
        {
            if (result.TotalFail == 0)
            {
                MessageBox.Show("すべてリセットしました", "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var message = new StringBuilder();
            message.AppendLine("リセットが完了しました。");
            message.AppendLine();
            if (scope == Libraries.PowerPointBulkResetScope.Textbook || scope == Libraries.PowerPointBulkResetScope.All)
                message.AppendLine($"演習 成功: {result.TextbookSuccess} / 失敗: {result.TextbookFail}");
            if (scope == Libraries.PowerPointBulkResetScope.Variant || scope == Libraries.PowerPointBulkResetScope.All)
                message.AppendLine($"類題 成功: {result.VariantSuccess} / 失敗: {result.VariantFail}");
            if (result.FailedProjects.Count > 0)
            {
                message.AppendLine();
                message.Append("失敗したプロジェクト: ");
                message.Append(string.Join(", ", result.FailedProjects));
            }

            MessageBox.Show(message.ToString(), "完了", MessageBoxButton.OK, MessageBoxImage.Warning);
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

        private void VariantPlaceholderButton_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.ResultMessage = "類題は後から接続します";
        }

        private void ScoringLog_EntryClicked(object sender, MosPracticeClient.ScoringLogEntry entry)
        {
            if (entry == null) return;
            var window = new Views.ResultWindow(
                null,
                null,
                null,
                entry.GroupId,
                entry.ProjectResults ?? new Dictionary<int, List<bool>>(),
                fromScoringLog: true);
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
                PowerPointApplicationManager.ClosePowerPointApplicationForAppExit();
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

        private void OnShowAppBarRequested(object sender, EventArgs e)
        {
            if (_isExiting)
                return;

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
                _appBarWindow = new Views.UiTestAppBarWindow(project.ProjectId, project.GroupId, _viewModel.ShowScoreButton, false, () => _viewModel.ScoreCommand.Execute(null), _viewModel.IsVariantMode, _viewModel);
                _appBarWindow.Closed += OnAppBarWindowClosed;
                oldAppBar?.Close();
            }
            if (_appBarWindow != null)
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
            this.Show();
            this.Activate();
        }

        private void OnExamEnded(object sender, EventArgs e)
        {
            _appBarWindow = null;
        }

        private void OnScoreCompleted(object sender, EventArgs e)
        {
            var results = _viewModel?.TaskResults;
            if (results == null || results.Count == 0)
                return;
            // 採点結果をアプリバーの解答済み状態に反映（結果画面で〇が正しく表示されるようにする）
            if (_appBarWindow != null && _viewModel?.CurrentProject != null)
            {
                try
                {
                    _appBarWindow.CloseInstantScoringOverlay();
                    _appBarWindow.ApplyScoreResults(_viewModel.CurrentProject.ProjectId, results);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[OnScoreCompleted] ApplyScoreResults error: {ex.Message}");
                }
            }
            if (_appBarWindow != null)
                _appBarWindow.ShowScoreResult(results);
            else
                Views.ScoreResultWindow.ShowResults(this, results);
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
                _viewModel.ScoreCompleted -= OnScoreCompleted;
            }

            if (!_isExiting)
                CloseAppBarForExit();

            base.OnClosed(e);
        }
    }
}
