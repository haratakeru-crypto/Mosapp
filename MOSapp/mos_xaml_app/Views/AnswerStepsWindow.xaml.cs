using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace MOSExcelMogiApp.Views
{
    public partial class AnswerStepsWindow : Window
    {
        public sealed class TaskStepsItem
        {
            public int TaskId { get; set; }
            public string AnswerSteps { get; set; }
        }

        private readonly List<TaskStepsItem> _tasks;
        private readonly string _titlePrefix;
        private readonly Func<int> _getAppBarTaskId;
        private int _currentIndex;

        /// <summary>採点ダイアログなど、単一タスク表示用。</summary>
        public AnswerStepsWindow(string title, string answerSteps)
            : this(
                titlePrefix: title ?? "解答手順",
                tasks: new List<TaskStepsItem>
                {
                    new TaskStepsItem { TaskId = 1, AnswerSteps = answerSteps }
                },
                initialTaskId: 1,
                getAppBarTaskId: null)
        {
        }

        public AnswerStepsWindow(
            string titlePrefix,
            IEnumerable<TaskStepsItem> tasks,
            int initialTaskId,
            Func<int> getAppBarTaskId)
        {
            InitializeComponent();
            _titlePrefix = string.IsNullOrWhiteSpace(titlePrefix) ? "解答手順" : titlePrefix.Trim();
            _tasks = (tasks ?? Enumerable.Empty<TaskStepsItem>())
                .Where(t => t != null && t.TaskId > 0)
                .GroupBy(t => t.TaskId)
                .Select(g => g.First())
                .OrderBy(t => t.TaskId)
                .ToList();
            _getAppBarTaskId = getAppBarTaskId;

            if (_tasks.Count == 0)
            {
                _tasks.Add(new TaskStepsItem { TaskId = Math.Max(1, initialTaskId), AnswerSteps = string.Empty });
            }

            _currentIndex = IndexOfTaskId(initialTaskId);
            if (_currentIndex < 0)
                _currentIndex = 0;

            ShowCurrentTask();
        }

        public void ApplyInitialPlacement(Rect preferredBoundsDip)
        {
            // 類題手順はテキストのみのため既定サイズ(560x480)を維持し、右上位置だけ合わせる
            if (preferredBoundsDip.Width > 0 && preferredBoundsDip.Height > 0)
            {
                Left = preferredBoundsDip.Right - Width;
                Top = preferredBoundsDip.Top;
            }

            ClampToScreen();
        }

        public void ClampToScreen()
        {
            var workArea = SystemParameters.WorkArea;
            if (Width > workArea.Width)
                Width = Math.Max(MinWidth, workArea.Width - 20);
            if (Height > workArea.Height)
                Height = Math.Max(MinHeight, workArea.Height - 20);

            if (Left < workArea.Left)
                Left = workArea.Left;
            if (Top < workArea.Top)
                Top = workArea.Top;
            if (Left + Width > workArea.Right)
                Left = Math.Max(workArea.Left, workArea.Right - Width);
            if (Top + Height > workArea.Bottom)
                Top = Math.Max(workArea.Top, workArea.Bottom - Height);
        }

        public void BringToFrontPassive()
        {
            try
            {
                if (!IsVisible)
                    return;
                Topmost = false;
                Topmost = true;
            }
            catch
            {
                /* ignore */
            }
        }

        /// <summary>既に開いているウィンドウで指定タスクの手順を表示する。</summary>
        public void ShowTask(int taskId, bool confirmIfMismatch)
        {
            int index = IndexOfTaskId(taskId);
            if (index < 0)
                return;
            if (confirmIfMismatch && !ConfirmOpenIfMismatched(_tasks[index].TaskId))
                return;

            _currentIndex = index;
            ShowCurrentTask();
        }

        private int IndexOfTaskId(int taskId)
        {
            return _tasks.FindIndex(t => t.TaskId == taskId);
        }

        private void ShowCurrentTask()
        {
            TaskStepsItem item = _tasks[_currentIndex];
            HeaderTextBlock.Text = $"{_titlePrefix} タスク {item.TaskId} 解答手順";
            AnswerStepsTextBlock.Text = string.IsNullOrWhiteSpace(item.AnswerSteps)
                ? "解答手順が登録されていません。"
                : item.AnswerSteps;
            PrevButton.IsEnabled = _currentIndex > 0;
            NextButton.IsEnabled = _currentIndex < _tasks.Count - 1;
        }

        private bool ConfirmOpenIfMismatched(int guideTaskId)
        {
            int appBarTaskId = 0;
            try
            {
                appBarTaskId = _getAppBarTaskId?.Invoke() ?? 0;
            }
            catch
            {
                appBarTaskId = 0;
            }

            if (appBarTaskId <= 0 || guideTaskId <= 0 || appBarTaskId == guideTaskId)
                return true;

            MessageBoxResult result = MessageBox.Show(
                this,
                $"アプリバーはタスク{appBarTaskId}です。タスク{guideTaskId}の手順を表示しますか？",
                "解答手順",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            return result == MessageBoxResult.Yes;
        }

        private void PrevButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex <= 0)
                return;
            int nextIndex = _currentIndex - 1;
            if (!ConfirmOpenIfMismatched(_tasks[nextIndex].TaskId))
                return;
            _currentIndex = nextIndex;
            ShowCurrentTask();
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentIndex >= _tasks.Count - 1)
                return;
            int nextIndex = _currentIndex + 1;
            if (!ConfirmOpenIfMismatched(_tasks[nextIndex].TaskId))
                return;
            _currentIndex = nextIndex;
            ShowCurrentTask();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
