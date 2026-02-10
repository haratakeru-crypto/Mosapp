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

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;
        private Views.UiTestAppBarWindow _appBarWindow;
        
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

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var result = MessageBox.Show("アプリ自体を終了します。本当にいいですか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
                e.Cancel = true;
        }

        private void OnShowAppBarRequested(object sender, EventArgs e)
        {
            if (_appBarWindow == null || !_appBarWindow.IsLoaded)
            {
                var project = _viewModel.CurrentProject;
                if (project != null)
                {
                    _appBarWindow = new Views.UiTestAppBarWindow(project.ProjectId, project.GroupId, _viewModel.ShowScoreButton, _viewModel.ShowPauseButton);
                    _appBarWindow.Closed += (s, args) =>
                    {
                        // バーウィンドウが閉じられたらメインウィンドウを再表示
                        this.Show();
                        this.Activate();
                        _appBarWindow = null;
                    };
                }
            }
            if (_appBarWindow != null)
            {
                _appBarWindow.Show();
            }
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

        private void TimerCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            IsTimerDisabled = false; // チェックが入っている = タイマー有効
            System.Diagnostics.Debug.WriteLine($"MainWindow: TimerCheckBox checked, IsTimerDisabled = {IsTimerDisabled}");
        }
        
        private void TimerCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            IsTimerDisabled = true; // チェックが外れている = タイマー無効
            System.Diagnostics.Debug.WriteLine($"MainWindow: TimerCheckBox unchecked, IsTimerDisabled = {IsTimerDisabled}");
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
            
            // アプリバーウィンドウを閉じる
            _appBarWindow?.Close();
            
            base.OnClosed(e);
        }
    }
}
