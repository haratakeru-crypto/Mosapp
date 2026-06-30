using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace MOSExcelMogiApp.Views
{
    public class ResultItem
    {
        public string Text { get; set; }
        public Brush Color { get; set; }
        public int TaskId { get; set; }
        public bool IsClickable => Text == "X";
    }

    public partial class ScoringResultDialog : Window
    {
        private int _groupId = 1;
        private int _projectId = 1;
        private DispatcherTimer _keepOnTopTimer;

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public ScoringResultDialog(int taskCount)
        {
            InitializeComponent();
            DisplayEmptyResults(taskCount);
            HookForegroundBehavior();
        }

        public ScoringResultDialog(int taskCount, List<bool> results, int groupId = 1, int projectId = 1)
        {
            InitializeComponent();
            _groupId = groupId;
            _projectId = projectId;
            DisplayResults(results);
            HookForegroundBehavior();
        }

        public ScoringResultDialog(List<bool> results, int groupId = 1, int projectId = 1)
        {
            InitializeComponent();
            _groupId = groupId;
            _projectId = projectId;
            DisplayResults(results);
            HookForegroundBehavior();
        }

        /// <summary>
        /// 採点結果を最前面のモーダルで表示する（Excel が前面に出ても維持）。
        /// </summary>
        public static void ShowResults(Window owner, int taskCount, List<bool> results, int groupId, int projectId)
        {
            var w = new ScoringResultDialog(taskCount, results, groupId, projectId)
            {
                Owner = owner,
                Topmost = true,
                ShowInTaskbar = true
            };

            if (owner != null)
            {
                owner.Topmost = true;
                owner.Activate();
            }

            w.ShowDialog();

            if (owner != null)
                owner.Topmost = true;
        }

        private void HookForegroundBehavior()
        {
            Loaded += ScoringResultDialog_Loaded;
            Closed += ScoringResultDialog_Closed;
        }

        private void ScoringResultDialog_Loaded(object sender, RoutedEventArgs e)
        {
            BringToForeground();
            _keepOnTopTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _keepOnTopTimer.Tick += KeepOnTopTimer_Tick;
            _keepOnTopTimer.Start();
        }

        private void ScoringResultDialog_Closed(object sender, EventArgs e)
        {
            if (_keepOnTopTimer == null)
                return;

            _keepOnTopTimer.Stop();
            _keepOnTopTimer.Tick -= KeepOnTopTimer_Tick;
            _keepOnTopTimer = null;
        }

        private void KeepOnTopTimer_Tick(object sender, EventArgs e)
        {
            if (!IsVisible)
                return;

            if (!IsActive)
            {
                Topmost = false;
                Topmost = true;
                BringToForeground();
            }
        }

        private void BringToForeground()
        {
            Topmost = true;
            Activate();
            try
            {
                var helper = new WindowInteropHelper(this);
                if (helper.Handle != IntPtr.Zero)
                    SetForegroundWindow(helper.Handle);
            }
            catch { }
        }

        private void DisplayEmptyResults(int taskCount)
        {
            var taskNumbers = Enumerable.Range(1, taskCount).ToList();
            TaskNumbersControl.ItemsSource = taskNumbers;

            var resultItems = Enumerable.Range(1, taskCount)
                .Select(i => new ResultItem { Text = "-", Color = Brushes.Gray, TaskId = i })
                .ToList();
            ResultsControl.ItemsSource = resultItems;
        }

        private void DisplayResults(List<bool> results)
        {
            System.Diagnostics.Debug.WriteLine($"[DisplayResults] Called with {results.Count} results, groupId: {_groupId}, projectId: {_projectId}");

            var taskNumbers = Enumerable.Range(1, results.Count).ToList();
            TaskNumbersControl.ItemsSource = taskNumbers;

            var resultItems = results.Select((r, index) => {
                var item = new ResultItem
                {
                    Text = r ? "O" : "X",
                    Color = r ? Brushes.Green : Brushes.Red,
                    TaskId = index + 1
                };
                System.Diagnostics.Debug.WriteLine($"[DisplayResults] Created ResultItem - TaskId: {item.TaskId}, Text: {item.Text}, IsClickable: {item.IsClickable}");
                return item;
            }).ToList();
            ResultsControl.ItemsSource = resultItems;
        }

        private void ResultItem_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[ResultItem_MouseDown] Event fired");

            if (sender is FrameworkElement element && element.DataContext is ResultItem resultItem)
            {
                System.Diagnostics.Debug.WriteLine($"[ResultItem_MouseDown] DataContext found - Text: {resultItem.Text}, TaskId: {resultItem.TaskId}, IsClickable: {resultItem.IsClickable}");

                if (resultItem.IsClickable && resultItem.Text == "X")
                {
                    System.Diagnostics.Debug.WriteLine($"[ResultItem_MouseDown] Showing image for TaskId: {resultItem.TaskId}");
                    ShowImage(resultItem.TaskId);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[ResultItem_MouseDown] Item is not clickable or not X - Text: {resultItem.Text}, IsClickable: {resultItem.IsClickable}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[ResultItem_MouseDown] DataContext is null or wrong type");
            }
        }

        private void ShowImage(int taskId)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[ShowImage] Called with taskId: {taskId}, groupId: {_groupId}, projectId: {_projectId}");

                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                System.Diagnostics.Debug.WriteLine($"[ShowImage] Base directory: {baseDir}");

                string imagePath = Path.Combine(
                    baseDir,
                    "References",
                    "Answers",
                    $"Group{_groupId}",
                    $"Project{_projectId}",
                    $"Task{taskId}.png"
                );

                System.Diagnostics.Debug.WriteLine($"[ShowImage] Constructed image path: {imagePath}");
                System.Diagnostics.Debug.WriteLine($"[ShowImage] File exists: {File.Exists(imagePath)}");

                if (File.Exists(imagePath))
                {
                    System.Diagnostics.Debug.WriteLine("[ShowImage] Opening ImageWindow");
                    var imageWindow = new ImageWindow(imagePath);
                    imageWindow.ShowDialog();
                }
                else
                {
                    string projectRootPath = Path.Combine(
                        baseDir,
                        "..",
                        "..",
                        "References",
                        "Answers",
                        $"Group{_groupId}",
                        $"Project{_projectId}",
                        $"Task{taskId}.png"
                    );
                    projectRootPath = Path.GetFullPath(projectRootPath);

                    System.Diagnostics.Debug.WriteLine($"[ShowImage] Trying alternative path: {projectRootPath}");
                    System.Diagnostics.Debug.WriteLine($"[ShowImage] Alternative path exists: {File.Exists(projectRootPath)}");

                    if (File.Exists(projectRootPath))
                    {
                        System.Diagnostics.Debug.WriteLine("[ShowImage] Opening ImageWindow with alternative path");
                        var imageWindow = new ImageWindow(projectRootPath);
                        imageWindow.ShowDialog();
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[ShowImage] Image file not found in both paths!");
                        MessageBox.Show(
                            $"画像ファイルが見つかりませんでした。\nパス1: {imagePath}\nパス2: {projectRootPath}",
                            "画像が見つかりません",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information
                        );
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ShowImage] Exception: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show(
                    $"画像の表示中にエラーが発生しました: {ex.Message}",
                    "エラー",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
