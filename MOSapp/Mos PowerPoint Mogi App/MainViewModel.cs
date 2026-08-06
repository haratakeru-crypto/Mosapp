using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using PowerPointPresentation = Microsoft.Office.Interop.PowerPoint.Presentation;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace MOS_PowerPoint_app
{
    public class ProjectInfo
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
        public string Group { get; set; }
        public int ProjectNumber { get; set; }
    }

    public class MainViewModel : INotifyPropertyChanged
    {
        private int _selectedTabIndex;
        private string _resultMessage;
        private bool _showScoreButton;
        private ProjectViewModel _currentProject;
        private ObservableCollection<TaskResult> _taskResults;
        private int _totalScore;
        private int _maxScore;

        public MainViewModel()
        {
            LoadProjects();
            OpenProjectCommand = new RelayCommand(ExecuteOpenProject);
            ScoreCommand = new RelayCommand(ExecuteScore, CanExecuteScore);
            ResetAllProjectsCommand = new RelayCommand(ExecuteResetAllProjects);
            TaskResults = new ObservableCollection<TaskResult>();
        }

        public ObservableCollection<ProjectGroupViewModel> ProjectGroups { get; set; } = new ObservableCollection<ProjectGroupViewModel>();

        public int SelectedTabIndex
        {
            get => _selectedTabIndex;
            set
            {
                _selectedTabIndex = value;
                OnPropertyChanged();
            }
        }

        public string ResultMessage
        {
            get => _resultMessage;
            set
            {
                _resultMessage = value;
                OnPropertyChanged();
            }
        }

        public ICommand OpenProjectCommand { get; }
        public ICommand ScoreCommand { get; }
        public ICommand ResetAllProjectsCommand { get; }

        public ProjectViewModel CurrentProject
        {
            get => _currentProject;
            set
            {
                _currentProject = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CurrentProjectName));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        /// <summary>採点ボタンをアプリバーに表示するか。デフォルトは非表示。</summary>
        public bool ShowScoreButton
        {
            get => _showScoreButton;
            set { _showScoreButton = value; OnPropertyChanged(nameof(ShowScoreButton)); }
        }

        public string CurrentProjectName => CurrentProject?.Name ?? "";

        public event EventHandler ShowAppBarRequested;
        public event EventHandler HideMainWindowRequested;
        /// <summary>採点完了時に発火。採点結果ダイアログの表示に使用する。</summary>
        public event EventHandler ScoreCompleted;
#pragma warning disable 67 // イベントは外部で使用されるため警告を抑制
        public event EventHandler ShowMainWindowRequested;
        public event EventHandler ExamEnded;
#pragma warning restore 67

        public ObservableCollection<TaskResult> TaskResults
        {
            get => _taskResults;
            set
            {
                _taskResults = value;
                OnPropertyChanged();
            }
        }

        public int TotalScore
        {
            get => _totalScore;
            set
            {
                _totalScore = value;
                OnPropertyChanged();
            }
        }

        public int MaxScore
        {
            get => _maxScore;
            set
            {
                _maxScore = value;
                OnPropertyChanged();
            }
        }

        private void LoadProjects()
        {
            try
            {
                // Tab1（演習）のみ読み込む
                foreach (int groupId in new[] { 1 })
                {
                    string tabFolder = PowerPointDataPathHelper.GetTabFolder(groupId);
                    var group = new ProjectGroupViewModel { GroupId = groupId, GroupName = $"Group {groupId}" };
                    
                    if (Directory.Exists(tabFolder))
                    {
                        // Project1.pptxからProject10.pptxを検索
                        for (int projectId = 1; projectId <= 10; projectId++)
                        {
                            string filePath = PowerPointDataPathHelper.GetWorkingProjectPath(groupId, projectId);
                            if (!File.Exists(filePath))
                            {
                                // .ppt は拡張子を変えてコピーせず、旧形式のまま互換利用する。
                                string legacyPptPath = Path.Combine(tabFolder, $"Project{projectId}.ppt");
                                filePath = File.Exists(legacyPptPath) ? legacyPptPath : null;
                            }
                            
                            group.Projects.Add(new ProjectViewModel
                            {
                                GroupId = groupId,
                                ProjectId = projectId,
                                Name = $"プロジェクト{groupId}-{projectId}",
                                FilePath = filePath
                            });
                        }
                    }
                    else
                    {
                        // フォルダが存在しない場合でも、空のプロジェクトリストを作成
                        for (int projectId = 1; projectId <= 10; projectId++)
                        {
                            group.Projects.Add(new ProjectViewModel
                            {
                                GroupId = groupId,
                                ProjectId = projectId,
                                Name = $"プロジェクト{groupId}-{projectId}",
                                FilePath = null
                            });
                        }
                    }
                    
                    ProjectGroups.Add(group);
                }
            }
            catch (Exception ex)
            {
                // 例外をログに記録するが、アプリを継続させる
                System.Diagnostics.Debug.WriteLine($"LoadProjectsエラー: {ex.Message}");
                // 空のプロジェクトグループを作成してアプリを継続（演習のみ）
                foreach (int groupId in new[] { 1 })
                {
                    var group = new ProjectGroupViewModel { GroupId = groupId, GroupName = $"Group {groupId}" };
                    for (int projectId = 1; projectId <= 10; projectId++)
                    {
                        group.Projects.Add(new ProjectViewModel
                        {
                            GroupId = groupId,
                            ProjectId = projectId,
                            Name = $"プロジェクト{groupId}-{projectId}",
                            FilePath = null
                        });
                    }
                    ProjectGroups.Add(group);
                }
            }
        }

        private void ExecuteOpenProject(object parameter)
        {
            if (parameter is ProjectViewModel project)
            {
                if (string.IsNullOrEmpty(project.FilePath) || !File.Exists(project.FilePath))
                {
                    ResultMessage = $"エラー: ファイルが見つかりません: {project.FilePath ?? "パスが設定されていません"}";
                    return;
                }

                try
                {
                    // プロジェクト起動前にタスク情報をクリアし、アドイン側の古いスナップショットとの比較を防止
                    Libraries.PPLogReader.ClearCurrentTaskFile();

                    // PowerPointアプリケーションを取得または作成
                    PowerPointApp pptApp = null;
                    try
                    {
                        pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch
                    {
                        pptApp = new PowerPointApp();
                        pptApp.Visible = MsoTriState.msoTrue;
                    }

                    try
                    {
                        pptApp.Presentations.Open(project.FilePath, WithWindow: MsoTriState.msoTrue);
                        Libraries.PowerPointViewHelper.HideNotesPane(pptApp);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"プレゼンテーションを開く際のエラー（既に開いている可能性があります）: {ex.Message}");
                    }

                    CurrentProject = project;
                    HideMainWindowRequested?.Invoke(this, EventArgs.Empty);
                    ShowAppBarRequested?.Invoke(this, EventArgs.Empty);
                    ResultMessage = $"PowerPointファイルを開きました: {Path.GetFileName(project.FilePath)}";
                }
                catch (Exception ex)
                {
                    ResultMessage = $"エラー: ファイルを開けませんでした: {ex.Message}";
                    System.Diagnostics.Debug.WriteLine($"エラー詳細: {ex.StackTrace}");
                }
            }
        }

        private bool CanExecuteScore(object parameter)
        {
            return CurrentProject != null;
        }

        private void ExecuteScore(object parameter)
        {
            if (CurrentProject == null)
            {
                ResultMessage = "プロジェクトを開いてから実行してください。";
                return;
            }

            TaskResults.Clear();
            ResultMessage = "採点中...";

            string jsonPath = PowerPointDataPathHelper.ResolveJsonPath(
                "MOS模擬アプリ問題文一覧_PowerPoint.json");
            if (!File.Exists(jsonPath))
            {
                ResultMessage = "該当プロジェクトのタスクが見つかりません（問題文JSONがありません）。";
                return;
            }

            MOS_PowerPoint_app.Views.ProjectData projectData;
            try
            {
                string jsonContent = File.ReadAllText(jsonPath, Encoding.UTF8);
                projectData = JsonConvert.DeserializeObject<MOS_PowerPoint_app.Views.ProjectData>(jsonContent);
            }
            catch (Exception ex)
            {
                ResultMessage = $"問題文の読み込みに失敗しました: {ex.Message}";
                return;
            }

            var project = projectData?.Projects?.FirstOrDefault(p => p.ProjectId == CurrentProject.ProjectId);
            if (project?.Tasks == null || project.Tasks.Count == 0)
            {
                ResultMessage = "該当プロジェクトのタスクが見つかりません。";
                return;
            }

            PowerPointGrader grader = null;
            try
            {
                grader = new PowerPointGrader();
                if (!grader.Connect())
                {
                    ResultMessage = "PowerPoint を起動し、対象のファイルを開いた状態で実行してください。";
                    return;
                }

                int passedCount = 0;
                foreach (var task in project.Tasks.OrderBy(t => t.TaskId))
                {
                    bool passed = false;
                    try
                    {
                        // 1タスクごとに current_task を更新し、VSTO 側の snapshot が追いつくのを短時間待つ。
                        int attemptNo = Libraries.PPTaskAttemptRegistry.GetAttempt(CurrentProject.ProjectId, task.TaskId);
                        grader.StartTaskAndWaitForSnapshot(CurrentProject.ProjectId, task.TaskId, attemptNo, 2000, 50);
                        passed = grader.GradeTask(CurrentProject.ProjectId, task.TaskId, attemptNo);
                    }
                    catch
                    {
                        passed = false;
                    }
                    if (passed) passedCount++;
                    TaskResults.Add(new TaskResult
                    {
                        TaskNumber = task.TaskId,
                        IsPassed = passed,
                        TaskName = string.IsNullOrEmpty(task.Description) ? $"タスク{task.TaskId}" : task.Description
                    });
                }
                ResultMessage = $"採点: {passedCount}/{project.Tasks.Count} タスク合格";
                ScoreCompleted?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                ResultMessage = $"採点中にエラーが発生しました: {ex.Message}";
            }
            finally
            {
                grader?.Dispose();
            }
        }

        private void ExecuteResetAllProjects(object parameter)
        {
            var result = MessageBox.Show(
                "すべてのPowerPointプロジェクトをテンプレートからリセットします。\n現在の変更内容は失われます。実行しますか？",
                "すべてをリセットする",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
                return;

            try { Libraries.PPTaskAttemptRegistry.ClearAll(); } catch { }

            // 全プロジェクトリセット時のみ採点用証跡も消す（単体リセットでは残す）
            try { Libraries.PPLogReader.ClearTaskEvidence(); } catch { }

            var errors = new System.Collections.Generic.List<string>();
            int done = 0;
            foreach (int groupId in new[] { 1 }) // Tab1のみリセットし、Tab3（応用編）は除外
            {
                for (int projectId = 1; projectId <= 10; projectId++)
                {
                    try
                    {
                        PowerPointProjectResetHelper.ResetProject(groupId, projectId);
                        done++;
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Tab{groupId} Project{projectId}: {ex.Message}");
                    }
                }
            }

            if (errors.Count == 0)
                ResultMessage = $"すべてのプロジェクトをリセットしました（{done}件）。";
            else
                ResultMessage = $"{done}件リセットしました。失敗: {errors.Count}件";
            MessageBox.Show(
                errors.Count == 0
                    ? $"すべてのプロジェクトをリセットしました（{done}件）。"
                    : $"{done}件リセットしました。\n失敗 {errors.Count}件:\n" + string.Join("\n", errors.Take(10)) + (errors.Count > 10 ? "\n…他" : ""),
                "リセット結果",
                MessageBoxButton.OK,
                errors.Count == 0 ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ProjectGroupViewModel
    {
        public int GroupId { get; set; }
        public string GroupName { get; set; }
        public ObservableCollection<ProjectViewModel> Projects { get; set; } = new ObservableCollection<ProjectViewModel>();
    }

    public class ProjectViewModel
    {
        public int GroupId { get; set; }
        public int ProjectId { get; set; }
        public string Name { get; set; }
        public string FilePath { get; set; }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool> _canExecute;

        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute?.Invoke(parameter) ?? true;
        }

        public void Execute(object parameter)
        {
            _execute(parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }

    public class TaskResult
    {
        public int TaskNumber { get; set; }
        public bool IsPassed { get; set; }
        public string TaskName { get; set; }
    }
}
