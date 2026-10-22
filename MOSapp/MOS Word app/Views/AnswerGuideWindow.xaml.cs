using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using MOS_Word_app.Guide;

namespace MOS_Word_app.Views
{
    /// <summary>
    /// モードレス解答手順ガイド（紙芝居）。Word 操作を妨げない。
    /// </summary>
    public partial class AnswerGuideWindow : Window
    {
        private const int MinZoomPercent = 25;
        private const int MaxZoomPercent = 400;

        private static readonly int[] ZoomPresets = { 50, 75, 100, 125, 150, 200 };

        private readonly int _groupId;
        private readonly int _projectId;
        private readonly List<AnswerGuideTask> _tasks;
        private readonly Func<int> _getAppBarTaskId;
        private AnswerGuideTask _currentTask;
        private int _stepIndex;
        private int _zoomPercent = 100;
        private bool _updatingZoomUi;
        private bool _hasStoredBounds;
        private double _storedLeft;
        private double _storedTop;
        private double _storedWidth;
        private double _storedHeight;

        public int GuideTaskId => _currentTask?.TaskId ?? 0;

        public bool IsShowingSteps => StepPanel.Visibility == Visibility.Visible && _currentTask != null;

        public AnswerGuideWindow(
            int groupId,
            int projectId,
            IEnumerable<AnswerGuideTask> tasks,
            Func<int> getAppBarTaskId = null)
        {
            InitializeComponent();
            _groupId = groupId;
            _projectId = projectId;
            _tasks = (tasks ?? Enumerable.Empty<AnswerGuideTask>()).ToList();
            _getAppBarTaskId = getAppBarTaskId ?? (() => 0);

            Closed += AnswerGuideWindow_Closed;
            LocationChanged += AnswerGuideWindow_BoundsChanged;
            SizeChanged += AnswerGuideWindow_BoundsChanged;

            InitZoomComboBox();
            ShowTaskList();
        }

        /// <summary>
        /// 開いた直後に AppBar の現在タスク手順へ合わせる（一致時は確認なし）。
        /// </summary>
        public void OpenForAppBarTask(int appBarTaskId)
        {
            AnswerGuideTask match = _tasks.FirstOrDefault(t => t.TaskId == appBarTaskId);
            if (match != null)
                ShowTaskSteps(match, 0, confirmIfMismatch: false);
            else
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
                Width = Math.Max(MinWidth, Math.Min(preferredBoundsDip.Width, 800));
                Height = Math.Max(MinHeight, Math.Min(preferredBoundsDip.Height, 720));
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

        private int GetAppBarTaskId()
        {
            try
            {
                return _getAppBarTaskId?.Invoke() ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private int IndexOfCurrentTask()
        {
            if (_currentTask == null)
                return -1;
            return _tasks.FindIndex(t => t.TaskId == _currentTask.TaskId);
        }

        private bool HasPreviousTask()
        {
            int index = IndexOfCurrentTask();
            return index > 0;
        }

        private bool HasNextTask()
        {
            int index = IndexOfCurrentTask();
            return index >= 0 && index < _tasks.Count - 1;
        }

        /// <summary>
        /// タスク一覧からの選択時など、AppBar と異なる手順を開く前の確認。
        /// </summary>
        private bool ConfirmOpenIfMismatched(int guideTaskId)
        {
            int appBarTaskId = GetAppBarTaskId();
            if (appBarTaskId <= 0 || guideTaskId <= 0 || appBarTaskId == guideTaskId)
                return true;

            var result = MessageBox.Show(
                this,
                $"アプリバーはタスク{appBarTaskId}です。タスク{guideTaskId}の手順を表示しますか？",
                "解答手順",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            return result == MessageBoxResult.Yes;
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
                ShowTaskSteps(task, 0, confirmIfMismatch: true);
        }

        private void ShowTaskSteps(AnswerGuideTask task, int stepIndex, bool confirmIfMismatch)
        {
            if (task == null || task.Steps.Count == 0)
            {
                ShowTaskList();
                return;
            }

            if (confirmIfMismatch && !ConfirmOpenIfMismatched(task.TaskId))
                return;

            _currentTask = task;
            _stepIndex = Math.Max(0, Math.Min(stepIndex, task.Steps.Count - 1));
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

            bool atFirstStep = _stepIndex <= 0;
            bool atLastStep = _stepIndex >= _currentTask.Steps.Count - 1;
            PrevButton.IsEnabled = !atFirstStep || HasPreviousTask();
            NextButton.IsEnabled = !atLastStep || HasNextTask();

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
                ApplyZoom();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AnswerGuideWindow] image load failed: {ex.Message}");
                StepImage.Source = null;
                CaptionTextBlock.Text = "画像を読み込めませんでした。";
            }
        }

        private void InitZoomComboBox()
        {
            _updatingZoomUi = true;
            try
            {
                ZoomComboBox.Items.Clear();
                foreach (int preset in ZoomPresets)
                    ZoomComboBox.Items.Add(FormatZoomPercent(preset));
                SyncZoomComboText();
            }
            finally
            {
                _updatingZoomUi = false;
            }
        }

        private void ApplyZoom()
        {
            if (!(StepImage.Source is BitmapSource bmp))
                return;

            double dpiX = bmp.DpiX > 1 ? bmp.DpiX : 96.0;
            double dpiY = bmp.DpiY > 1 ? bmp.DpiY : 96.0;
            double dipWidth = bmp.PixelWidth * 96.0 / dpiX;
            double dipHeight = bmp.PixelHeight * 96.0 / dpiY;
            double factor = _zoomPercent / 100.0;

            StepImage.Stretch = Stretch.Fill;
            StepImage.MaxWidth = double.PositiveInfinity;
            StepImage.MaxHeight = double.PositiveInfinity;
            StepImage.Width = dipWidth * factor;
            StepImage.Height = dipHeight * factor;
            SyncZoomComboText();
        }

        private void SetZoomPercent(int percent)
        {
            _zoomPercent = ClampZoomPercent(percent);
            ApplyZoom();
        }

        private static int ClampZoomPercent(int percent)
        {
            if (percent < MinZoomPercent)
                return MinZoomPercent;
            if (percent > MaxZoomPercent)
                return MaxZoomPercent;
            return percent;
        }

        private static string FormatZoomPercent(int percent) => $"{percent}%";

        private void SyncZoomComboText()
        {
            if (ZoomComboBox == null)
                return;

            _updatingZoomUi = true;
            try
            {
                string text = FormatZoomPercent(_zoomPercent);
                if (ZoomComboBox.Text != text)
                    ZoomComboBox.Text = text;
            }
            finally
            {
                _updatingZoomUi = false;
            }
        }

        private bool TryParseZoomInput(string raw, out int percent)
        {
            percent = _zoomPercent;
            if (string.IsNullOrWhiteSpace(raw))
                return false;

            string text = raw.Trim();
            if (text.EndsWith("%", StringComparison.Ordinal))
                text = text.Substring(0, text.Length - 1).Trim();

            if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) &&
                !int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out value))
            {
                return false;
            }

            percent = ClampZoomPercent(value);
            return true;
        }

        private void CommitZoomInput()
        {
            if (_updatingZoomUi)
                return;

            if (!TryParseZoomInput(ZoomComboBox.Text, out int percent))
            {
                SyncZoomComboText();
                return;
            }

            SetZoomPercent(percent);
        }

        private void ZoomComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_updatingZoomUi || ZoomComboBox.SelectedItem == null)
                return;

            if (!TryParseZoomInput(ZoomComboBox.SelectedItem.ToString(), out int percent))
                return;

            SetZoomPercent(percent);
        }

        private void ZoomComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
                return;
            CommitZoomInput();
            e.Handled = true;
        }

        private void ZoomComboBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            CommitZoomInput();
        }

        private void PrevButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTask == null)
                return;

            if (_stepIndex > 0)
            {
                _stepIndex--;
                UpdateStepView();
                return;
            }

            int index = IndexOfCurrentTask();
            if (index <= 0)
                return;

            AnswerGuideTask prev = _tasks[index - 1];
            ShowTaskSteps(prev, prev.Steps.Count - 1, confirmIfMismatch: true);
        }

        private void NextButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTask == null)
                return;

            if (_stepIndex < _currentTask.Steps.Count - 1)
            {
                _stepIndex++;
                UpdateStepView();
                return;
            }

            int index = IndexOfCurrentTask();
            if (index < 0 || index >= _tasks.Count - 1)
                return;

            ShowTaskSteps(_tasks[index + 1], 0, confirmIfMismatch: true);
        }

        private void TaskListButton_Click(object sender, RoutedEventArgs e)
        {
            ShowTaskList();
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
    }
}
