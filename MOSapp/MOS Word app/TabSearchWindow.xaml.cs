using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace MOS_Word_app
{
    public partial class TabSearchWindow : Window
    {
        private MainViewModel _viewModel;
        private TabChecker _tabChecker;
        private int _currentTaskIndex = 0;

        public TabSearchWindow(MainViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
            _tabChecker = new TabChecker();
            
            // 最初のタスクを表示
            if (_viewModel.TabTasks.Count > 0)
            {
                _viewModel.CurrentTabTask = _viewModel.TabTasks[0];
            }
        }

        private void CheckTabButton_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel.CurrentTabTask == null)
            {
                MessageBox.Show("タスクが選択されていません", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // 解答操作からタブ名を抽出（例: "挿入タブを選択" → "挿入"）
                string answer = _viewModel.CurrentTabTask.Answer;
                string tabName = answer.Replace("タブを選択", "").Replace("タブ", "").Trim();

                // タブをチェック
                bool isPassed = _tabChecker.CheckTab(tabName);
                _viewModel.CurrentTabTask.IsPassed = isPassed;

                // 結果を表示（メッセージボックスではなく、すぐに次の問題に進む）
                if (isPassed)
                {
                    // 正解の場合、自動的に次の問題に進む
                    if (_currentTaskIndex < _viewModel.TabTasks.Count - 1)
                    {
                        _currentTaskIndex++;
                        _viewModel.CurrentTabTask = _viewModel.TabTasks[_currentTaskIndex];
                    }
                    else
                    {
                        // すべてのタスクが完了
                        int passedCount = _viewModel.TabTasks.Count(t => t.IsPassed);
                        MessageBox.Show($"すべてのタスクが完了しました！\n合格: {passedCount}/{_viewModel.TabTasks.Count}", 
                            "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                // 不正解の場合は、同じ問題を続ける（メッセージボックスは表示しない）
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void NextTaskButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTaskIndex < _viewModel.TabTasks.Count - 1)
            {
                _currentTaskIndex++;
                _viewModel.CurrentTabTask = _viewModel.TabTasks[_currentTaskIndex];
            }
            else
            {
                MessageBox.Show("これが最後の問題です", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void TaskItem_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var border = sender as System.Windows.Controls.Border;
            if (border == null) return;
            if (!(border.DataContext is TabTaskInfo)) return;
            var task = (TabTaskInfo)border.DataContext;
            int index = _viewModel.TabTasks.IndexOf(task);
            if (index >= 0)
            {
                _currentTaskIndex = index;
                _viewModel.CurrentTabTask = task;
            }
        }

        private void PauseButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("一時停止機能は準備中です", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void EndExamButton_Click(object sender, RoutedEventArgs e)
        {
            int passedCount = _viewModel.TabTasks.Count(t => t.IsPassed);
            int totalCount = _viewModel.TabTasks.Count;
            MessageBox.Show($"試験を終了します。\n合格: {passedCount}/{totalCount}", 
                "試験終了", MessageBoxButton.OK, MessageBoxImage.Information);
            this.Close();
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("すべての回答をリセットしますか？", "確認", 
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
            {
                foreach (var task in _viewModel.TabTasks)
                {
                    task.IsPassed = false;
                }
                _currentTaskIndex = 0;
                _viewModel.CurrentTabTask = _viewModel.TabTasks[0];
            }
        }

        private void NextProjectButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("次のプロジェクト機能は準備中です", "情報", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}

