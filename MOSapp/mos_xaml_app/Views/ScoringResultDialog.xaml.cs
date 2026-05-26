using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.IO;
using System;

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

        public ScoringResultDialog(int taskCount)
        {
            InitializeComponent();
            DisplayEmptyResults(taskCount);
        }

        public ScoringResultDialog(int taskCount, List<bool> results, int groupId = 1, int projectId = 1)
        {
            InitializeComponent();
            _groupId = groupId;
            _projectId = projectId;
            DisplayResults(results);
        }

        public ScoringResultDialog(List<bool> results, int groupId = 1, int projectId = 1)
        {
            InitializeComponent();
            _groupId = groupId;
            _projectId = projectId;
            DisplayResults(results);
        }

        private void DisplayEmptyResults(int taskCount)
        {
            // タスク番号を生成 (1, 2, 3, ...)
            var taskNumbers = Enumerable.Range(1, taskCount).ToList();
            TaskNumbersControl.ItemsSource = taskNumbers;

            // 空の結果を表示
            var resultItems = Enumerable.Range(1, taskCount)
                .Select(i => new ResultItem { Text = "-", Color = Brushes.Gray, TaskId = i })
                .ToList();
            ResultsControl.ItemsSource = resultItems;
        }

        private void DisplayResults(List<bool> results)
        {
            System.Diagnostics.Debug.WriteLine($"[DisplayResults] Called with {results.Count} results, groupId: {_groupId}, projectId: {_projectId}");
            
            // タスク番号を生成 (1, 2, 3, ...)
            var taskNumbers = Enumerable.Range(1, results.Count).ToList();
            TaskNumbersControl.ItemsSource = taskNumbers;

            // 結果をO/Xに変換して色付き
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
                
                // 画像パスを構築: References/Answers/Group{groupId}/Project{projectId}/Task{taskId}.png
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
                    // 代替パスを試す（プロジェクトルートから）
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
            this.Close();
        }
    }
}