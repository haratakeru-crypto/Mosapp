using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using MOSExcelMogiApp.Guide;

namespace MOSExcelMogiApp.Views
{
    /// <summary>
    /// モードレス解答手順ガイド（紙芝居）。Excel 操作を妨げない。
    /// </summary>
    public partial class AnswerGuideWindow : Window
    {
        private readonly int _groupId;
        private readonly int _projectId;
        private readonly List<AnswerGuideTask> _tasks;
        private AnswerGuideTask _currentTask;
        private int _stepIndex;
        private bool _fitToWidth = true;
        private bool _hasStoredBounds;
        private double _storedLeft;
        private double _storedTop;
        private double _storedWidth;
        private double _storedHeight;

        public int GuideTaskId => _currentTask?.TaskId ?? 0;

        public bool IsShowingSteps => StepPanel.Visibility == Visibility.Visible && _currentTask != null;

        public AnswerGuideWindow(int groupId, int projectId, IEnumerable<AnswerGuideTask> tasks)
        {
            InitializeComponent();
            _groupId = groupId;
            _projectId = projectId;
            _tasks = (tasks ?? Enumerable.Empty<AnswerGuideTask>()).ToList();

            Closed += AnswerGuideWindow_Closed;
            LocationChanged += AnswerGuideWindow_BoundsChanged;
            SizeChanged += AnswerGuideWindow_BoundsChanged;

            ShowTaskList();
        }

        public void ApplyInitialPlacement(Rect preferredBoundsDip)
        {
            if (_hasStoredBounds)
            {
                Left = _storedLeft;
                Top = _storedTop;
                Width = _storedWidth;
                Height = _storedHeight;
                ClampToScreen();
                return;
            }

            if (preferredBoundsDip.Width > 0 && preferredBoundsDip.Height > 0)
            {
                Width = Math.Max(MinWidth, Math.Min(preferredBoundsDip.Width, 520));
                Height = Math.Max(MinHeight, Math.Min(preferredBoundsDip.Height, 620));
                Left = preferredBoundsDip.Left;
                Top = preferredBoundsDip.Top;
            }

            ClampToScreen();
            RememberBounds();
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

        /// <summary>
        /// AppBar のタスクが変わったときの同期。戻り値はガイド側が切り替わったか。
        /// </summary>
        public void OnAppBarTaskChanged(int appBarTaskId)
        {
            if (!IsVisible || !IsShowingSteps)
                return;

            if (appBarTaskId == GuideTaskId)
                return;

            var result = MessageBox.Show(
                this,
                $"アプリバーはタスク{appBarTaskId}です。手順もタスク{appBarTaskId}に切り替えますか？",
                "解答手順",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
                return;

            AnswerGuideTask target = _tasks.FirstOrDefault(t => t.TaskId == appBarTaskId);
            if (target == null)
            {
                MessageBox.Show(
                    this,
                    "このタスクの解答手順はまだありません",
                    "解答手順",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            ShowTaskSteps(target);
        }

        private void ShowTaskList()
        {
            _currentTask = null;
            _stepIndex = 0;
            HeaderTextBlock.Text = "解答手順 — タスクを選ぶ";
            TaskListPanel.Children.Clear();
            EmptyMessageTextBlock.Visibility = Visibility.Collapsed;
            StepPanel.Visibility = Visibility.Collapsed;
            TaskListScroll.Visibility = Visibility.Visible;
            PrevButton.IsEnabled = false;
            NextButton.IsEnabled = false;
            TaskListButton.IsEnabled = false;

            if (_tasks.Count == 0)
            {
                TaskListScroll.Visibility = Visibility.Collapsed;
                EmptyMessageTextBlock.Text = "このプロジェクトの手順はまだありません";
                EmptyMessageTextBlock.Visibility = Visibility.Visible;
                return;
            }

            foreach (AnswerGuideTask task in _tasks)
            {
                var button = new Button
                {
                    Content = $"タスク {task.TaskId}（{task.Steps.Count}ステップ）",
                    Margin = new Thickness(0, 0, 0, 8),
                    Padding = new Thickness(12, 10, 12, 10),
                    HorizontalContentAlignment = HorizontalAlignment.Left,
                    Tag = task,
                    Background = Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(148, 163, 184)),
                    BorderThickness = new Thickness(1),
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                button.Click += TaskButton_Click;
                TaskListPanel.Children.Add(button);
            }
        }

        private void TaskButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is AnswerGuideTask task)
                ShowTaskSteps(task);
        }

        private void ShowTaskSteps(AnswerGuideTask task)
        {
            _currentTask = task;
            _stepIndex = 0;
            TaskListScroll.Visibility = Visibility.Collapsed;
            EmptyMessageTextBlock.Visibility = Visibility.Collapsed;
            StepPanel.Visibility = Visibility.Visible;
            TaskListButton.IsEnabled = true;
            UpdateStepView();
        }

        private void UpdateStepView()
        {
            if (_currentTask == null || _currentTask.Steps.Count == 0)
            {
                ShowTaskList();
                return;
            }

            if (_stepIndex < 0)
                _stepIndex = 0;
            if (_stepIndex >= _currentTask.Steps.Count)
                _stepIndex = _currentTask.Steps.Count - 1;

            AnswerGuideLoadedStep step = _currentTask.Steps[_stepIndex];
            HeaderTextBlock.Text = $"タスク{_currentTask.TaskId}の手順（{_stepIndex + 1}/{_currentTask.Steps.Count}）";
            CaptionTextBlock.Text = string.IsNullOrWhiteSpace(step.Caption)
                ? "（説明なし）"
                : step.Caption;
            ProgressTextBlock.Text = $"{_stepIndex + 1} / {_currentTask.Steps.Count}";
            PrevButton.IsEnabled = _stepIndex > 0;
            NextButton.IsEnabled = _stepIndex < _currentTask.Steps.Count - 1;
            LoadStepImage(step.ImagePath);
        }

        private void LoadStepImage(string imagePath)
        {
            try
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.UriSource = new Uri(imagePath, UriKind.Absolute);
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                StepImage.Source = bitmap;
                ApplyStretchMode();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AnswerGuideWindow] image load failed: {ex.Message}");
                StepImage.Source = null;
                CaptionTextBlock.Text = "画像を読み込めませんでした。";
            }
        }

        private void ApplyStretchMode()
        {
            if (_fitToWidth)
            {
                StepImage.Stretch = Stretch.Uniform;
                StepImage.Width = double.NaN;
                StepImage.Height = double.NaN;
                StepImage.MaxWidth = Math.Max(120, ActualWidth - 48);
            }
            else if (StepImage.Source is BitmapSource bmp)
            {
                StepImage.Stretch = Stretch.None;
                StepImage.Width = bmp.PixelWidth;
                StepImage.Height = bmp.PixelHeight;
                StepImage.MaxWidth = double.PositiveInfinity;
            }
        }

        private void PrevButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTask == null || _stepIndex <= 0)
                return;
            _stepIndex--;
            UpdateStepView();
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTask == null || _stepIndex >= _currentTask.Steps.Count - 1)
                return;
            _stepIndex++;
            UpdateStepView();
        }

        private void TaskListButton_Click(object sender, RoutedEventArgs e)
        {
            ShowTaskList();
        }

        private void FitWidthButton_Click(object sender, RoutedEventArgs e)
        {
            _fitToWidth = true;
            ApplyStretchMode();
        }

        private void ActualSizeButton_Click(object sender, RoutedEventArgs e)
        {
            _fitToWidth = false;
            ApplyStretchMode();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void AnswerGuideWindow_BoundsChanged(object sender, EventArgs e)
        {
            if (IsLoaded && IsVisible)
                RememberBounds();
        }

        private void RememberBounds()
        {
            if (WindowState != WindowState.Normal)
                return;
            _storedLeft = Left;
            _storedTop = Top;
            _storedWidth = Width;
            _storedHeight = Height;
            _hasStoredBounds = true;
        }

        private void AnswerGuideWindow_Closed(object sender, EventArgs e)
        {
            RememberBounds();
        }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);
            if (_fitToWidth && StepPanel.Visibility == Visibility.Visible)
                ApplyStretchMode();
        }
    }
}
