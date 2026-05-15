using System;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using Newtonsoft.Json;
using System.IO;

namespace MOSExcelMogiApp.Maui.Pages;

public partial class ReviewPage : ContentPage
{
    private System.Timers.Timer? _timer;
    private TimeSpan _remainingTime;
    private int _groupId = 1; // Group番号（1=模擬①, 2=模擬②, 3=演習）
    
    private Dictionary<int, bool[]> _projectTaskCompletedStates;
    private Dictionary<int, bool[]> _projectTaskFlaggedStates;
    
    public Action<int, int>? OnNavigateToTask { get; set; } // ProjectId, TaskId
    
    public ReviewPage(TimeSpan remainingTime, Dictionary<int, bool[]> completedStates, Dictionary<int, bool[]> flaggedStates, int groupId = 1)
    {
        InitializeComponent();
        _remainingTime = remainingTime;
        _projectTaskCompletedStates = completedStates ?? new Dictionary<int, bool[]>();
        _projectTaskFlaggedStates = flaggedStates ?? new Dictionary<int, bool[]>();
        _groupId = groupId;
        
        LoadAllProjects();
        InitializeTimer();
    }
    
    private void InitializeTimer()
    {
        UpdateTimerDisplay();
        
        // タイマーを1秒間隔で更新
        _timer = new System.Timers.Timer(1000);
        _timer.Elapsed += Timer_Elapsed;
        _timer.AutoReset = true;
        _timer.Start();
    }
    
    private void Timer_Elapsed(object? sender, ElapsedEventArgs e)
    {
        MainThread.BeginInvokeOnMainThread(() =>
        {
            if (_remainingTime.TotalSeconds > 0)
            {
                _remainingTime = _remainingTime.Subtract(TimeSpan.FromSeconds(1));
                UpdateTimerDisplay();
            }
            else
            {
                _timer?.Stop();
                UpdateTimerDisplay();
            }
        });
    }
    
    private void UpdateTimerDisplay()
    {
        if (TimerTextBlock != null)
        {
            TimerTextBlock.Text = _remainingTime.ToString(@"hh\:mm\:ss");
        }
    }
    
    private void LoadAllProjects()
    {
        try
        {
            // Group番号に応じたJSONファイルを選択
            string jsonFileName = _groupId switch
            {
                1 => "MOS演習問題文一覧.json",        // GroupId=1 → 演習タブ
                2 => "MOS模擬試験①問題文一覧.json",  // GroupId=2 → 模試①タブ
                3 => "MOS模擬試験②問題文一覧.json",  // GroupId=3 → 模試②タブ
                _ => "MOS模擬アプリ問題文一覧.json"
            };
            
            // JSONファイルのパスを取得（実行ファイルのディレクトリから相対パスで取得）
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string jsonPath = Path.Combine(baseDirectory, "References", "JSON", jsonFileName);
            
            // もし見つからない場合は、親ディレクトリを探す
            if (!File.Exists(jsonPath))
            {
                string parentPath = Path.Combine(baseDirectory, "..", "..", "..", "..", "References", "JSON", jsonFileName);
                if (File.Exists(parentPath))
                {
                    jsonPath = parentPath;
                }
            }
            
            System.Diagnostics.Debug.WriteLine($"[ReviewPage] Loading from: {jsonFileName} (GroupId: {_groupId})");
            System.Diagnostics.Debug.WriteLine($"[ReviewPage] JSON Path: {jsonPath}");
            
            if (!File.Exists(jsonPath))
            {
                System.Diagnostics.Debug.WriteLine($"[ReviewPage] JSON file not found: {jsonPath}");
                // ダミーデータを表示
                LoadDummyData();
                return;
            }
            
            string jsonContent = File.ReadAllText(jsonPath);
            var projectData = JsonConvert.DeserializeObject<ProjectData>(jsonContent);
            
            if (projectData?.Projects != null)
            {
                var reviewProjects = new List<ReviewProjectInfo>();
                
                foreach (var project in projectData.Projects.OrderBy(p => p.ProjectId))
                {
                    var reviewProject = new ReviewProjectInfo
                    {
                        ProjectTitle = $"プロジェクト {project.ProjectId}",
                        Tasks = project.Tasks?.Select(task => new ReviewTaskInfo
                        {
                            TaskTitle = $"タスク {task.TaskId}",
                            Description = task.Description,
                            ProjectId = project.ProjectId,
                            TaskId = task.TaskId
                        }).ToList() ?? new List<ReviewTaskInfo>()
                    };
                    
                    reviewProjects.Add(reviewProject);
                }
                
                ProjectsCollectionView.ItemsSource = reviewProjects;
                
                // 状態表示を更新
                UpdateTaskStates();
            }
            else
            {
                LoadDummyData();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"レビューページ読み込みエラー: {ex.Message}");
            System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
            LoadDummyData();
        }
    }
    
    private void LoadDummyData()
    {
        // ダミーデータを表示（テスト用）
        var reviewProjects = new List<ReviewProjectInfo>();
        
        for (int projectId = 1; projectId <= 10; projectId++)
        {
            var tasks = new List<ReviewTaskInfo>();
            for (int taskId = 1; taskId <= 7; taskId++)
            {
                tasks.Add(new ReviewTaskInfo
                {
                    TaskTitle = $"タスク {taskId}",
                    Description = $"プロジェクト{projectId}のタスク{taskId}の説明",
                    ProjectId = projectId,
                    TaskId = taskId
                });
            }
            
            reviewProjects.Add(new ReviewProjectInfo
            {
                ProjectTitle = $"プロジェクト {projectId}",
                Tasks = tasks
            });
        }
        
        ProjectsCollectionView.ItemsSource = reviewProjects;
    }
    
    private void UpdateTaskStates()
    {
        // TODO: タスクの状態（解答済み、あとで見直す）を更新
        // 現在は簡易実装のため、後で実装
    }
    
    private void NavigateToTask(ReviewTaskInfo taskInfo)
    {
        System.Diagnostics.Debug.WriteLine($"NavigateToTask called: ProjectId={taskInfo.ProjectId}, TaskId={taskInfo.TaskId}");
        
        if (OnNavigateToTask != null && taskInfo.ProjectId > 0 && taskInfo.TaskId > 0)
        {
            try
 {
                _timer?.Stop();
                OnNavigateToTask(taskInfo.ProjectId, taskInfo.TaskId);
                
                // レビューページを閉じる
                if (Navigation.NavigationStack.Count > 1)
                {
                    Navigation.PopAsync();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ナビゲーション実行エラー: {ex.Message}");
            }
        }
    }
    
    private void TaskButton_Tapped(object? sender, TappedEventArgs e)
    {
        if (sender is Border border)
        {
            // BindingContextを取得
            if (border.BindingContext is ReviewTaskInfo taskInfo)
            {
                NavigateToTask(taskInfo);
            }
            else
            {
                // 親要素から取得を試みる
                var parent = border.Parent;
                while (parent != null)
                {
                    if (parent is BindableObject bindable && bindable.BindingContext is ReviewTaskInfo info)
                    {
                        NavigateToTask(info);
                        break;
                    }
                    parent = parent is Element element ? element.Parent : null;
                }
            }
        }
    }
    
    private void TaskButton_Clicked(object? sender, EventArgs e)
    {
        if (sender is Button button && button.BindingContext is ReviewTaskInfo taskInfo)
        {
            NavigateToTask(taskInfo);
        }
    }
    
    private void EndExamButton_Clicked(object? sender, EventArgs e)
    {
        // TODO: 試験終了処理を実装
        DisplayAlert("情報", "試験終了機能は準備中です。", "OK");
    }
    
    private void CloseButton_Clicked(object? sender, EventArgs e)
    {
        _timer?.Stop();
        if (Navigation.NavigationStack.Count > 1)
        {
            Navigation.PopAsync();
        }
    }
    
    protected override void OnDisappearing()
    {
        _timer?.Stop();
        base.OnDisappearing();
    }
}

// データモデル
public class ReviewProjectInfo
{
    public string ProjectTitle { get; set; } = string.Empty;
    public List<ReviewTaskInfo> Tasks { get; set; } = new();
}

public class ReviewTaskInfo
{
    public string TaskTitle { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public int ProjectId { get; set; }
    public int TaskId { get; set; }
}

// JSONデータモデル
public class ProjectData
{
    public List<ProjectInfo>? Projects { get; set; }
}

public class ProjectInfo
{
    public int ProjectId { get; set; }
    public List<TaskInfo>? Tasks { get; set; }
}

public class TaskInfo
{
    public int TaskId { get; set; }
    public string Description { get; set; } = string.Empty;
}

