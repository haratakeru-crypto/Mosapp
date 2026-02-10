using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Runtime.InteropServices;
using System.Configuration;
using Newtonsoft.Json.Linq;
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
        private bool _showPauseButton;
        private ProjectViewModel _currentProject;
        private ObservableCollection<TaskResult> _taskResults;
        private int _totalScore;
        private int _maxScore;

        public MainViewModel()
        {
            LoadProjects();
            OpenProjectCommand = new RelayCommand(ExecuteOpenProject);
            UiTestCommand = new RelayCommand(ExecuteUiTest);
            ScoreCommand = new RelayCommand(ExecuteScore, CanExecuteScore);
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
        public ICommand UiTestCommand { get; }
        public ICommand ScoreCommand { get; }

        public ProjectViewModel CurrentProject
        {
            get => _currentProject;
            set
            {
                _currentProject = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CurrentProjectName));
            }
        }

        /// <summary>採点ボタンをアプリバーに表示するか。デフォルトは非表示。</summary>
        public bool ShowScoreButton
        {
            get => _showScoreButton;
            set { _showScoreButton = value; OnPropertyChanged(nameof(ShowScoreButton)); }
        }

        /// <summary>一時停止ボタンをアプリバーに表示するか。デフォルトは非表示。</summary>
        public bool ShowPauseButton
        {
            get => _showPauseButton;
            set { _showPauseButton = value; OnPropertyChanged(nameof(ShowPauseButton)); }
        }

        public string CurrentProjectName => CurrentProject?.Name ?? "";

        public event EventHandler ShowAppBarRequested;
        public event EventHandler HideMainWindowRequested;
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
                // 優先順位1: Assets\config.json（mos_xaml_app と同様に、ClickOnceでも同梱しやすい）
                // 優先順位2: App.config の appSettings
                // 優先順位3: 既定値
                string basePath = null;

                try
                {
                    string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                    if (File.Exists(configPath))
                    {
                        string jsonContent = File.ReadAllText(configPath);
                        JObject config = JObject.Parse(jsonContent);
                        basePath = config["powerPointDataPath"]?.ToString();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Assets\\config.json 読み込みエラー: {ex.Message}");
                }

                if (string.IsNullOrWhiteSpace(basePath))
                {
                    basePath = ConfigurationManager.AppSettings["PowerPointDataPath"];
                }

                if (string.IsNullOrWhiteSpace(basePath))
                {
                    basePath = @"C:\MOSTest\PowerPoint365";
                }
                
                // Tab1, Tab3のフォルダからプロジェクトを読み込む（模試①=Tab2は非表示のためスキップ）
                foreach (int groupId in new[] { 1, 3 })
                {
                    string tabFolder = Path.Combine(basePath, $"Tab{groupId}");
                    var group = new ProjectGroupViewModel { GroupId = groupId, GroupName = groupId == 3 ? "応用編" : $"Group {groupId}" };
                    
                    if (Directory.Exists(tabFolder))
                    {
                        // Project1.pptxからProject11.pptxを検索
                        for (int projectId = 1; projectId <= 11; projectId++)
                        {
                            // Project1.pptx, Project2.pptx, ... を検索
                            string[] possibleNames = { $"Project{projectId}.pptx", $"Project{projectId}.ppt" };
                            string filePath = null;
                            
                            foreach (var fileName in possibleNames)
                            {
                                string fullPath = Path.Combine(tabFolder, fileName);
                                if (File.Exists(fullPath))
                                {
                                    filePath = fullPath;
                                    break;
                                }
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
                        for (int projectId = 1; projectId <= 11; projectId++)
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
                // 空のプロジェクトグループを作成してアプリを継続（模試①=Group2はスキップ）
                foreach (int groupId in new[] { 1, 3 })
                {
                    var group = new ProjectGroupViewModel { GroupId = groupId, GroupName = groupId == 3 ? "応用編" : $"Group {groupId}" };
                    for (int projectId = 1; projectId <= 11; projectId++)
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

                    // PowerPointプレゼンテーションを開く
                    PowerPointPresentation presentation = null;
                    try
                    {
                        presentation = pptApp.Presentations.Open(project.FilePath, WithWindow: MsoTriState.msoTrue);
                    }
                    catch (Exception ex)
                    {
                        // ファイルが既に開いている場合は無視
                        System.Diagnostics.Debug.WriteLine($"プレゼンテーションを開く際のエラー（既に開いている可能性があります）: {ex.Message}");
                    }
                    
                    CurrentProject = project;
                    
                    // イベントを発火してアプリバーを表示
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

        private void ExecuteUiTest(object parameter)
        {
            ResultMessage = "UIテスト機能は準備中です";
        }

        private bool CanExecuteScore(object parameter)
        {
            return CurrentProject != null && !string.IsNullOrEmpty(CurrentProject.FilePath);
        }

        private void ExecuteScore(object parameter)
        {
            if (CurrentProject == null || string.IsNullOrEmpty(CurrentProject.FilePath))
            {
                ResultMessage = "エラー: プロジェクトが選択されていません";
                return;
            }

            try
            {
                ResultMessage = "採点中...";
                TaskResults.Clear();

                // 採点機能は後で実装
                ResultMessage = "採点機能は準備中です";
            }
            catch (Exception ex)
            {
                ResultMessage = $"エラー: 採点中にエラーが発生しました: {ex.Message}";
            }
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
