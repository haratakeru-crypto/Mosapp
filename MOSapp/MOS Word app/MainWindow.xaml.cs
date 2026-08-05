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

namespace MOS_Word_app
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;
        private Views.UiTestAppBarWindow _appBarWindow;

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
            _viewModel.ScoreCompleted += OnScoreCompleted;

            Loaded += MainWindow_Loaded;
            Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var result = MessageBox.Show("アプリ自体を終了します。本当にいいですか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
                e.Cancel = true;
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
                _appBarWindow.CurrentGroupId != project.GroupId;

            if (needRecreate)
            {
                var oldAppBar = _appBarWindow;
                _appBarWindow = new Views.UiTestAppBarWindow(project.ProjectId, project.GroupId, _viewModel.ShowScoreButton, _viewModel.ShowPauseButton);
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
            this.Show();
            this.Activate();
            _appBarWindow = null;
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
            Views.ScoreResultWindow.ShowResults(this, results);
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
            _appBarWindow?.Close();
            base.OnClosed(e);
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            if (!App.AutoOpenGroupId.HasValue || !App.AutoOpenProjectId.HasValue)
            {
                return;
            }

            int groupId = App.AutoOpenGroupId.Value;
            int projectId = App.AutoOpenProjectId.Value;
            App.ClearAutoOpen();

            var group = _viewModel.ProjectGroups?.FirstOrDefault(g => g.GroupId == groupId);
            var project = group?.Projects?.FirstOrDefault(p => p.ProjectId == projectId);

            if (project == null)
            {
                return;
            }

            if (_viewModel.OpenProjectCommand.CanExecute(project))
            {
                _viewModel.OpenProjectCommand.Execute(project);
            }
        }
    }
}
