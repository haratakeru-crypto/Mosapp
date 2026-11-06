using System;
using System.Collections.Generic;
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
        private bool _showScoreButton = true;
        private bool _showAnswerGuide = true;
        private bool _isAnswerGuideOpen;
        private ProjectViewModel _currentProject;
        private ObservableCollection<TaskResult> _taskResults;
        private int _totalScore;
        private int _maxScore;

        public MainViewModel()
        {
            LoadProjects();
            LoadVariantLaunchProjects();
            OpenProjectCommand = new RelayCommand(p => ExecuteOpenProject(p, openAsVariant: false), _ => !_isOpeningProject);
            OpenVariantProjectCommand = new RelayCommand(p => ExecuteOpenProject(p, openAsVariant: true), CanOpenVariantProject);
            ScoreCommand = new RelayCommand(ExecuteScore, CanExecuteScore);
            TaskResults = new ObservableCollection<TaskResult>();
        }

        public ObservableCollection<ProjectGroupViewModel> ProjectGroups { get; set; } = new ObservableCollection<ProjectGroupViewModel>();

        public IReadOnlyList<ProjectViewModel> VariantLaunchProjects { get; private set; } = Array.Empty<ProjectViewModel>();

        public ICommand OpenVariantProjectCommand { get; }

        public bool IsVariantMode { get; private set; }

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

        /// <summary>採点ボタンをアプリバーに表示するか。デフォルトは表示。</summary>
        public bool ShowScoreButton
        {
            get => _showScoreButton;
            set { _showScoreButton = value; OnPropertyChanged(nameof(ShowScoreButton)); }
        }

        /// <summary>解答手順ガイドの表示。デフォルトは表示。</summary>
        public bool ShowAnswerGuide
        {
            get => _showAnswerGuide;
            set
            {
                if (_showAnswerGuide == value)
                    return;
                _showAnswerGuide = value;
                OnPropertyChanged(nameof(ShowAnswerGuide));
                NotifyAnswerGuideStateChanged();
                if (!value)
                    RequestCloseAnswerGuide();
            }
        }

        /// <summary>通常教材は紙芝居手順、類題はテキスト手順を出す。どちらもプロジェクト選択のオン/オフに従う。</summary>
        public bool ShowAnswerGuideEffective
        {
            get
            {
                if (!ShowAnswerGuide)
                    return false;
                if (IsVariantMode)
                    return true;
                return CurrentProject == null || CurrentProject.GroupId == 1;
            }
        }

        /// <summary>解答手順ガイド表示中。</summary>
        public bool IsAnswerGuideOpen => _isAnswerGuideOpen;

        /// <summary>リセット・レビュー・次プロ・終了を抑止する統合ロック。</summary>
        public bool IsNavigationLocked => _isAnswerGuideOpen;

        public event EventHandler CloseAnswerGuideRequested;

        public void SetAnswerGuideOpen(bool open)
        {
            if (_isAnswerGuideOpen == open)
                return;
            _isAnswerGuideOpen = open;
            OnPropertyChanged(nameof(IsAnswerGuideOpen));
            OnPropertyChanged(nameof(IsNavigationLocked));
            CommandManager.InvalidateRequerySuggested();
        }

        public void RequestCloseAnswerGuide()
        {
            CloseAnswerGuideRequested?.Invoke(this, EventArgs.Empty);
        }

        private void NotifyAnswerGuideStateChanged()
        {
            OnPropertyChanged(nameof(ShowAnswerGuideEffective));
            OnPropertyChanged(nameof(IsNavigationLocked));
            CommandManager.InvalidateRequerySuggested();
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

        private void LoadVariantLaunchProjects()
        {
            var projects = new List<ProjectViewModel>();
            var ids = VariantProblemWorkbook.GetProjectIds().ToList();
            if (ids.Count == 0)
            {
                for (int projectId = 1; projectId <= 6; projectId++)
                    ids.Add(projectId);
            }

            foreach (int projectId in ids)
            {
                string filePath = PowerPointDataPathHelper.FindVariantWorkingFile(1, projectId, 1);
                projects.Add(new ProjectViewModel
                {
                    GroupId = 1,
                    ProjectId = projectId,
                    Name = $"プロジェクト1-{projectId}",
                    FilePath = filePath
                });
            }

            VariantLaunchProjects = projects;
        }

        private bool CanOpenVariantProject(object parameter)
        {
            if (_isOpeningProject)
                return false;
            if (!(parameter is ProjectViewModel project))
                return true;
            return PowerPointDataPathHelper.IsVariantFileAvailable(project.GroupId, project.ProjectId, 1);
        }

        private const int ExistingPptHeartbeatWaitMs = 400;
        private const int VstoHeartbeatWaitAfterLaunchMs = 8000;
        private const int ActiveVstoHeartbeatMaxAgeSeconds = 15;

        private bool _isOpeningProject;

        public void EnableProjectSelection()
        {
            var dispatcher = System.Windows.Application.Current?.Dispatcher;
            if (dispatcher != null && !dispatcher.CheckAccess())
            {
                dispatcher.BeginInvoke(new Action(EnableProjectSelection));
                return;
            }

            if (!_isOpeningProject)
                return;
            _isOpeningProject = false;
            CommandManager.InvalidateRequerySuggested();
        }

        private void DisableProjectSelection()
        {
            _isOpeningProject = true;
            CommandManager.InvalidateRequerySuggested();
        }

        private void ExecuteOpenProject(object parameter)
        {
            ExecuteOpenProject(parameter, openAsVariant: false);
        }

        private void ExecuteOpenProject(object parameter, bool openAsVariant)
        {
            if (!MosPracticeClient.ExamStartGuard.EnsureRegistered())
                return;

            if (parameter is ProjectViewModel project)
            {
                if (openAsVariant)
                {
                    string variantPath = PowerPointDataPathHelper.FindVariantWorkingFile(project.GroupId, project.ProjectId, 1);
                    if (string.IsNullOrEmpty(variantPath) || !File.Exists(variantPath))
                    {
                        ResultMessage = $"エラー: 類題ファイルが見つかりません: {variantPath ?? "パスが設定されていません"}";
                        return;
                    }
                    project.FilePath = variantPath;
                }
                else if (string.IsNullOrEmpty(project.FilePath) || !File.Exists(project.FilePath))
                {
                    ResultMessage = $"エラー: ファイルが見つかりません: {project.FilePath ?? "パスが設定されていません"}";
                    return;
                }

                DisableProjectSelection();
                Views.PowerPointStartupInputGate.Begin();

                string filePath = project.FilePath;
                var dispatcher = System.Windows.Application.Current?.Dispatcher;

                // 開く処理とアドイン待ちは画面スレッドの外で行い、準備中ダイアログを残す。
                System.Threading.Tasks.Task.Run(() =>
                {
                    string errorMessage = null;
                    bool vstoFailed = false;
                    bool opened = false;
                    try
                    {
                        string vstoIssue;
                        if (!Libraries.VSTOInstallerHelper.EnsureAddInReadyForExam(out vstoIssue))
                        {
                            errorMessage = "エラー: PowerPoint 用 VSTO アドインを準備できません。"
                                + (string.IsNullOrEmpty(vstoIssue) ? "" : ("\n" + vstoIssue));
                            vstoFailed = true;
                        }
                        else
                        {
                            Libraries.PPLogReader.ClearCurrentTaskFile();

                            PowerPointApp pptApp = null;
                            bool launchedNew = false;
                            try
                            {
                                pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                            }
                            catch
                            {
                                Libraries.PPLogReader.ClearVstoHeartbeat();
                                pptApp = new PowerPointApp();
                                pptApp.Visible = MsoTriState.msoTrue;
                                launchedNew = true;
                            }

                            try
                            {
                                pptApp.Presentations.Open(filePath, WithWindow: MsoTriState.msoTrue);
                                Libraries.PowerPointViewHelper.HideNotesPane(pptApp);
                                Views.PowerPointStartupInputGate.DisablePowerPointWindows();
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"プレゼンテーションを開く際のエラー（既に開いている可能性があります）: {ex.Message}");
                            }

                            int waitMs = launchedNew ? VstoHeartbeatWaitAfterLaunchMs : ExistingPptHeartbeatWaitMs;
                            if (!Libraries.PPLogReader.IsVstoHeartbeatFresh(ActiveVstoHeartbeatMaxAgeSeconds)
                                && !Libraries.PPLogReader.WaitForVstoHeartbeat(waitMs, ActiveVstoHeartbeatMaxAgeSeconds)
                                && (launchedNew
                                    || !Libraries.PPLogReader.WaitForVstoHeartbeat(VstoHeartbeatWaitAfterLaunchMs, ActiveVstoHeartbeatMaxAgeSeconds)))
                            {
                                errorMessage = "エラー: PowerPoint 用 VSTO アドインが応答していません。アドインを有効にしてから再度開いてください。";
                                vstoFailed = true;
                            }
                            else
                            {
                                opened = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errorMessage = $"エラー: ファイルを開けませんでした: {ex.Message}";
                        System.Diagnostics.Debug.WriteLine($"エラー詳細: {ex.StackTrace}");
                    }

                    if (dispatcher == null)
                    {
                        EnableProjectSelection();
                        return;
                    }

                    dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (!opened)
                        {
                            Views.PowerPointStartupInputGate.End();
                            EnableProjectSelection();
                            ResultMessage = errorMessage ?? "エラー: PowerPointファイルを開けませんでした。";
                            if (vstoFailed)
                            {
                                MessageBox.Show(
                                    ResultMessage,
                                    "VSTO 未準備",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Warning);
                            }
                            return;
                        }

                        IsVariantMode = openAsVariant;
                        CurrentProject = project;
                        NotifyAnswerGuideStateChanged();
                        if (openAsVariant)
                            RequestCloseAnswerGuide();
                        HideMainWindowRequested?.Invoke(this, EventArgs.Empty);
                        ShowAppBarRequested?.Invoke(this, EventArgs.Empty);
                        ResultMessage = $"PowerPointファイルを開きました: {Path.GetFileName(filePath)}";
                    }));
                });
            }
        }

        private bool CanExecuteScore(object parameter)
        {
            return CurrentProject != null;
        }

        private void ExecuteVariantScore()
        {
            if (CurrentProject == null)
            {
                RunOnUi(() => ResultMessage = "プロジェクトを開いてから実行してください。");
                return;
            }

            RunOnUi(() =>
            {
                TaskResults.Clear();
                ResultMessage = "採点中...";
            });

            try
            {
                var results = PowerPointVariantScoring.ScoreProject(
                    CurrentProject.GroupId, CurrentProject.ProjectId);
                if (results.Count == 0)
                {
                    RunOnUi(() => ResultMessage = "該当プロジェクトのタスクが見つかりません。");
                    return;
                }

                int passedCount = results.Count(r => r.IsPassed);
                RunOnUi(() =>
                {
                    foreach (var r in results)
                        TaskResults.Add(r);
                    ResultMessage = $"採点: {passedCount}/{results.Count} タスク合格";
                    ScoreCompleted?.Invoke(this, EventArgs.Empty);
                });
            }
            catch (Exception ex)
            {
                RunOnUi(() => ResultMessage = $"採点中にエラーが発生しました: {ex.Message}");
            }
        }

        private void ExecuteScore(object parameter)
        {
            if (CurrentProject == null)
            {
                RunOnUi(() => ResultMessage = "プロジェクトを開いてから実行してください。");
                return;
            }

            if (IsVariantMode)
            {
                ExecuteVariantScore();
                return;
            }

            if (!Libraries.PPLogReader.IsVstoHeartbeatFresh(ActiveVstoHeartbeatMaxAgeSeconds)
                && !Libraries.PPLogReader.WaitForVstoHeartbeat(VstoHeartbeatWaitAfterLaunchMs, ActiveVstoHeartbeatMaxAgeSeconds))
            {
                RunOnUi(() =>
                {
                    ResultMessage = "PowerPoint 用 VSTO アドインが応答していないため採点できません。アドインを有効にしてから再度実行してください。";
                    MessageBox.Show(ResultMessage, "VSTO 未準備", MessageBoxButton.OK, MessageBoxImage.Warning);
                });
                return;
            }

            RunOnUi(() =>
            {
                TaskResults.Clear();
                ResultMessage = "採点中...";
            });

            string jsonPath = PowerPointDataPathHelper.ResolveJsonPath(
                "MOS模擬アプリ問題文一覧_PowerPoint.json");
            if (!File.Exists(jsonPath))
            {
                RunOnUi(() => ResultMessage = "該当プロジェクトのタスクが見つかりません（問題文JSONがありません）。");
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
                RunOnUi(() => ResultMessage = $"問題文の読み込みに失敗しました: {ex.Message}");
                return;
            }

            var project = projectData?.Projects?.FirstOrDefault(p => p.ProjectId == CurrentProject.ProjectId);
            if (project?.Tasks == null || project.Tasks.Count == 0)
            {
                RunOnUi(() => ResultMessage = "該当プロジェクトのタスクが見つかりません。");
                return;
            }

            PowerPointGrader grader = null;
            try
            {
                grader = new PowerPointGrader();
                if (!grader.Connect())
                {
                    RunOnUi(() => ResultMessage = "PowerPoint を起動し、対象のファイルを開いた状態で実行してください。");
                    return;
                }

                grader.LogOpenTaskBaselineDiffOnce(CurrentProject.ProjectId);
                Libraries.PPLogReader.ClearSnapshot();

                // 開いたまま採点するとき、離脱境界を通っていない現在タスクの証跡を先に確定する。
                Libraries.PPLogReader.RequestOpenTaskEvidenceFlush(CurrentProject.ProjectId);

                int passedCount = 0;
                foreach (var task in project.Tasks.OrderBy(t => t.TaskId))
                {
                    bool passed = false;
                    try
                    {
                        int attemptNo = Libraries.PPTaskAttemptRegistry.GetAttempt(CurrentProject.ProjectId, task.TaskId);
                        passed = grader.GradeTask(CurrentProject.ProjectId, task.TaskId, attemptNo);
                    }
                    catch
                    {
                        passed = false;
                    }
                    if (passed) passedCount++;
                    var taskResult = new TaskResult
                    {
                        TaskNumber = task.TaskId,
                        IsPassed = passed,
                        TaskName = string.IsNullOrEmpty(task.Description) ? $"タスク{task.TaskId}" : task.Description
                    };
                    RunOnUi(() => TaskResults.Add(taskResult));
                }
                int taskCount = project.Tasks.Count;
                RunOnUi(() =>
                {
                    ResultMessage = $"採点: {passedCount}/{taskCount} タスク合格";
                    ScoreCompleted?.Invoke(this, EventArgs.Empty);
                });
            }
            catch (Exception ex)
            {
                RunOnUi(() => ResultMessage = $"採点中にエラーが発生しました: {ex.Message}");
            }
            finally
            {
                grader?.Dispose();
            }
        }

        private static void RunOnUi(Action action)
        {
            var dispatcher = System.Windows.Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
                action();
            else
                dispatcher.Invoke(action);
        }

        public List<Libraries.PowerPointResetTarget> GetBulkResetTargets(Libraries.PowerPointBulkResetScope scope)
        {
            var targets = new List<Libraries.PowerPointResetTarget>();

            if (scope == Libraries.PowerPointBulkResetScope.Textbook || scope == Libraries.PowerPointBulkResetScope.All)
            {
                // Tab1のみ。Tab3（応用編）は除外
                ProjectGroupViewModel group = ProjectGroups?.FirstOrDefault(g => g.GroupId == 1);
                if (group?.Projects != null)
                {
                    foreach (ProjectViewModel project in group.Projects)
                    {
                        targets.Add(new Libraries.PowerPointResetTarget
                        {
                            GroupId = group.GroupId,
                            ProjectId = project.ProjectId,
                            VariantSetNo = 0,
                            Label = "演習 プロジェクト" + project.ProjectId
                        });
                    }
                }
            }

            if (scope == Libraries.PowerPointBulkResetScope.Variant || scope == Libraries.PowerPointBulkResetScope.All)
            {
                int variantSetNo = Libraries.PowerPointBulkResetCoordinator.DefaultVariantSetNo;
                if (VariantLaunchProjects != null)
                {
                    foreach (ProjectViewModel project in VariantLaunchProjects)
                    {
                        string templatePath = PowerPointDataPathHelper.FindVariantTemplateFile(
                            project.GroupId, project.ProjectId, variantSetNo);
                        if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
                            continue;

                        targets.Add(new Libraries.PowerPointResetTarget
                        {
                            GroupId = project.GroupId,
                            ProjectId = project.ProjectId,
                            VariantSetNo = variantSetNo,
                            Label = "類題 プロジェクト" + project.ProjectId
                        });
                    }
                }
            }

            return targets;
        }

        /// <summary>選択範囲の作業ファイルをテンプレートから戻す。確認ダイアログは出さない。</summary>
        public Libraries.PowerPointBulkResetResult RunBulkReset(Libraries.PowerPointBulkResetScope scope)
        {
            var targets = GetBulkResetTargets(scope);
            if (targets.Count == 0)
                return new Libraries.PowerPointBulkResetResult();

            try { Libraries.PPTaskAttemptRegistry.ClearAll(); } catch { }

            // 全プロジェクトリセット時のみ採点用証跡も消す（単体リセットでは残す）
            if (scope == Libraries.PowerPointBulkResetScope.Textbook || scope == Libraries.PowerPointBulkResetScope.All)
            {
                try { Libraries.PPLogReader.ClearTaskEvidence(); } catch { }
            }

            return Libraries.PowerPointBulkResetCoordinator.CopyTargets(targets, null);
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
