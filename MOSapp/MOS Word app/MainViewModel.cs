using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Reflection;
using WordApp = Microsoft.Office.Interop.Word.Application;
using WordDoc = Microsoft.Office.Interop.Word.Document;
using WordWindow = Microsoft.Office.Interop.Word.Window;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Libraries;

namespace MOS_Word_app
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
        #region Win32 (docx 最前面表示用)
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        #endregion

        private int _selectedTabIndex;
        private string _resultMessage;
        private bool _showScoreButton;
        private bool _showPauseButton;
        private TabTaskInfo _currentTabTask;
        private ProjectViewModel _currentProject;
        private ObservableCollection<TaskResult> _taskResults;
        private int _totalScore;
        private int _maxScore;

        public MainViewModel()
        {
            LoadProjects();
            OpenProjectCommand = new RelayCommand(ExecuteOpenProject);
            UiTestCommand = new RelayCommand(ExecuteUiTest);
            TabSearchCommand = new RelayCommand(ExecuteTabSearch);
            ScoreCommand = new RelayCommand(ExecuteScore, CanExecuteScore);
            ResetAllInGroupCommand = new RelayCommand(ExecuteResetAllInGroup);
            TabTasks = new ObservableCollection<TabTaskInfo>();
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
                OnPropertyChanged(nameof(CurrentGroupIdForReset));
            }
        }

        /// <summary>現在選択中のタブに対応するグループID（ヘッダーの「すべてリセット」用）。演習=1, 応用=3。</summary>
        public int CurrentGroupIdForReset => SelectedTabIndex == 0 ? 1 : 3;

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
        public ICommand TabSearchCommand { get; }
        public ICommand ScoreCommand { get; }
        public ICommand ResetAllInGroupCommand { get; }

        public ObservableCollection<TabTaskInfo> TabTasks { get; set; }

        public TabTaskInfo CurrentTabTask
        {
            get => _currentTabTask;
            set
            {
                _currentTabTask = value;
                OnPropertyChanged();
            }
        }

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

        public string CurrentProjectName => CurrentProject?.Name ?? "";

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

        public event EventHandler ShowAppBarRequested;
        public event EventHandler HideMainWindowRequested;
#pragma warning disable 67 // イベントは MainWindow で購読されるため警告を抑制
        public event EventHandler ShowMainWindowRequested;
        public event EventHandler ExamEnded;
        /// <summary>採点完了時に発火。採点結果ウィンドウの表示に使用する。</summary>
        public event EventHandler ScoreCompleted;
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
            string basePath = @"C:\MOSTest\Word365";
            // #region agent log
            try
            {
                var line = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:LoadProjects\",\"message\":\"LoadProjects entry\",\"data\":{\"basePath\":\"" + (basePath ?? "").Replace("\\", "\\\\") + "\"},\"sessionId\":\"debug-session\",\"hypothesisId\":\"H1\"}\n";
                File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", line);
            }
            catch { }
            // #endregion

            // 保存先は Tab{groupId}\ 直下。一覧は Initial にファイルがあれば表示し、FilePath は作業フォルダ（Tab\）のパスにする
            for (int groupId = 1; groupId <= 3; groupId++)
            {
                string workingFolder = Path.Combine(basePath, $"Tab{groupId}");
                string initialFolder = Path.Combine(basePath, $"Tab{groupId}", "Initial");
                var group = new ProjectGroupViewModel { GroupId = groupId, GroupName = $"Group {groupId}" };
                // #region agent log
                try
                {
                    bool dirExists = Directory.Exists(initialFolder);
                    var line = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:LoadProjects\",\"message\":\"tabFolder check\",\"data\":{\"groupId\":" + groupId + ",\"tabFolder\":\"" + (initialFolder ?? "").Replace("\\", "\\\\") + "\",\"dirExists\":" + dirExists.ToString().ToLowerInvariant() + "},\"sessionId\":\"debug-session\",\"hypothesisId\":\"H1\"}\n";
                    File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", line);
                }
                catch { }
                // #endregion

                for (int projectId = 1; projectId <= 10; projectId++)
                {
                    // 保存先（作業フォルダ）のパス。保存はここにのみ反映する
                    string workingFileName = (groupId == 1 && projectId == 7) ? "Project7.doc" : $"Project{projectId}.docx";
                    string workingFilePath = Path.Combine(workingFolder, workingFileName);
                    string[] possibleNames = (groupId == 1 && projectId == 7)
                        ? new[] { "Project7.doc", "project7.doc" }
                        : new[] { $"Project{projectId}.docx", $"Project{projectId}.doc", $"project{projectId}.docx", $"project{projectId}.doc" };

                    bool existsInWorking = File.Exists(workingFilePath);
                    bool existsInInitial = false;
                    if (Directory.Exists(initialFolder))
                    {
                        foreach (var fileName in possibleNames)
                        {
                            if (File.Exists(Path.Combine(initialFolder, fileName)))
                            {
                                existsInInitial = true;
                                break;
                            }
                        }
                    }
                    string filePath = (existsInWorking || existsInInitial) ? workingFilePath : null;

                    if (groupId == 1 && projectId == 1)
                    {
                        // #region agent log
                        try
                        {
                            var pathEsc = (filePath ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"");
                            var line = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:LoadProjects\",\"message\":\"project 1-1 filePath\",\"data\":{\"filePath\":\"" + pathEsc + "\",\"hasPath\":" + (!string.IsNullOrEmpty(filePath)).ToString().ToLowerInvariant() + "},\"sessionId\":\"debug-session\",\"hypothesisId\":\"H1\"}\n";
                            File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", line);
                        }
                        catch { }
                        // #endregion
                    }
                    group.Projects.Add(new ProjectViewModel
                    {
                        GroupId = groupId,
                        ProjectId = projectId,
                        Name = $"プロジェクト{groupId}-{projectId}",
                        FilePath = filePath
                    });
                }
                
                ProjectGroups.Add(group);
            }
        }

        private void ExecuteOpenProject(object parameter)
        {
            // #region agent log
            try
            {
                string paramType = parameter?.GetType()?.FullName ?? "null";
                bool isPvm = parameter is ProjectViewModel;
                var line = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"ExecuteOpenProject called\",\"data\":{\"parameterNull\":" + (parameter == null).ToString().ToLowerInvariant() + ",\"parameterType\":\"" + (paramType ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\",\"isProjectViewModel\":" + isPvm.ToString().ToLowerInvariant() + "},\"sessionId\":\"debug-session\",\"hypothesisId\":\"param\"}\n";
                var logPath = @"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log";
                try { File.AppendAllText(logPath, line); } catch { File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + "debug.log", line); }
            }
            catch { }
            // #endregion

            if (!(parameter is ProjectViewModel))
            {
                // #region agent log
                try
                {
                    string paramType = parameter?.GetType()?.FullName ?? "null";
                    var line = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"early return parameter not ProjectViewModel\",\"data\":{\"parameterType\":\"" + (paramType ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"},\"sessionId\":\"debug-session\",\"hypothesisId\":\"param\"}\n";
                    File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", line);
                }
                catch { }
                // #endregion
                return;
            }
            var project = (ProjectViewModel)parameter;
            {
                // #region agent log
                try
                {
                    var pathEsc = (project?.FilePath ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"");
                    var line = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"ExecuteOpenProject entry\",\"data\":{\"filePath\":\"" + pathEsc + "\",\"groupId\":" + (project?.GroupId ?? 0) + ",\"projectId\":" + (project?.ProjectId ?? 0) + "},\"sessionId\":\"debug-session\",\"hypothesisId\":\"E\"}\n";
                    var logPath = @"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log";
                    try { File.AppendAllText(logPath, line); } catch { File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + "debug.log", line); }
                }
                catch { }
                // #endregion
                if (string.IsNullOrEmpty(project.FilePath))
                {
                    // #region agent log
                    try
                    {
                        var pathEsc = (project?.FilePath ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"");
                        var line = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"early return file not found\",\"data\":{\"filePath\":\"" + pathEsc + "\",\"pathEmpty\":" + string.IsNullOrEmpty(project?.FilePath).ToString().ToLowerInvariant() + "},\"sessionId\":\"debug-session\",\"hypothesisId\":\"H3\"}\n";
                        File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", line);
                    }
                    catch { }
                    // #endregion
                    ResultMessage = $"エラー: ファイルが見つかりません: {project.FilePath ?? "パスが設定されていません"}";
                    return;
                }
                // 作業フォルダ（Tab\）にファイルが無い場合は Initial からコピーしてから開く（保存は常に Tab\ にのみ反映）
                if (!File.Exists(project.FilePath))
                {
                    string basePath = @"C:\MOSTest\Word365";
                    int groupId = project.GroupId;
                    int projectId = project.ProjectId;
                    string initialFolder = Path.Combine(basePath, $"Tab{groupId}", "Initial");
                    string initialInitialFolder = Path.Combine(basePath, $"Tab{groupId}", "Initial", "Initial");
                    string[] possibleNames = (groupId == 1 && projectId == 7)
                        ? new[] { "Project7.doc", "project7.doc" }
                        : new[] { $"Project{projectId}.docx", $"Project{projectId}.doc", $"project{projectId}.docx", $"project{projectId}.doc" };
                    string sourcePath = null;
                    foreach (var fileName in possibleNames)
                    {
                        string fullPath = Path.Combine(initialFolder, fileName);
                        if (File.Exists(fullPath)) { sourcePath = fullPath; break; }
                    }
                    if (string.IsNullOrEmpty(sourcePath) && Directory.Exists(initialInitialFolder))
                    {
                        foreach (var fileName in possibleNames)
                        {
                            string fullPath = Path.Combine(initialInitialFolder, fileName);
                            if (File.Exists(fullPath)) { sourcePath = fullPath; break; }
                        }
                    }
                    if (string.IsNullOrEmpty(sourcePath))
                    {
                        ResultMessage = $"エラー: 参照元ファイルが見つかりません: {initialFolder}";
                        return;
                    }
                    try
                    {
                        string workingFolder = Path.Combine(basePath, $"Tab{groupId}");
                        if (!Directory.Exists(workingFolder))
                            Directory.CreateDirectory(workingFolder);
                        File.Copy(sourcePath, project.FilePath, overwrite: false);
                    }
                    catch (Exception exCopy)
                    {
                        ResultMessage = $"エラー: ファイルをコピーできませんでした: {exCopy.Message}";
                        return;
                    }
                }

                // VSTOアドインのインストール状態をチェック
                var vstoStatus = Libraries.VSTOInstallerHelper.GetInstallStatus();
                if (!vstoStatus.IsInstalled)
                {
                    string message = vstoStatus.GetInstallationMessage();
                    MessageBoxResult result = MessageBox.Show(
                        message + "\n\nこのままWordを起動しますか？\n（VSTOが必要なタスクの採点が正しく行われない可能性があります）",
                        "VSTOアドイン未インストール",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning
                    );

                    if (result == MessageBoxResult.No)
                    {
                        // #region agent log
                        try { File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"early return VSTO No\",\"data\":{},\"sessionId\":\"debug-session\",\"hypothesisId\":\"vsto\"}\n"); } catch { }
                        // #endregion
                        ResultMessage = "Wordの起動をキャンセルしました。";
                        return;
                    }
                }

                try
                {
                    // Wordアプリケーションを取得または作成
                    WordApp wordApp = null;
                    try
                    {
                        wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                        wordApp.Visible = true;
                        // #region agent log
                        try { File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"Word app GetActiveObject ok\",\"data\":{},\"sessionId\":\"debug-session\",\"hypothesisId\":\"H4\"}\n"); } catch { }
                        // #endregion
                    }
                    catch (Exception exWord)
                    {
                        // #region agent log
                        try { File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"Word app create\",\"data\":{\"error\":\"" + (exWord?.Message ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"},\"sessionId\":\"debug-session\",\"hypothesisId\":\"H4\"}\n"); } catch { }
                        // #endregion
                        wordApp = new WordApp();
                        wordApp.Visible = true;
                    }

                    // 同じパスで既に開いているドキュメントがあれば保存せずに閉じる（メモリではなくフォルダから開き直す）
                    string pathLower = System.IO.Path.GetFullPath(project.FilePath).ToLowerInvariant();
                    try
                    {
                        for (int i = wordApp.Documents.Count; i >= 1; i--)
                        {
                            WordDoc openDoc = wordApp.Documents[i];
                            try
                            {
                                string fullName = openDoc.FullName?.ToLowerInvariant() ?? "";
                                string docFullPath = fullName;
                                try { docFullPath = System.IO.Path.GetFullPath(fullName).ToLowerInvariant(); } catch { }
                                if (fullName == pathLower || docFullPath == pathLower)
                                {
                                    openDoc.Close(SaveChanges: false);
                                    break;
                                }
                            }
                            finally
                            {
                                if (openDoc != null) Marshal.ReleaseComObject(openDoc);
                            }
                        }
                    }
                    catch (Exception exClose)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ExecuteOpenProject] 既存ドキュメント閉じる際のエラー: {exClose.Message}");
                    }

                    // Wordドキュメントを開く（編集可能で開く・常にフォルダから）
                    WordDoc doc = null;
                    try
                    {
                        doc = wordApp.Documents.Open(project.FilePath, ReadOnly: false, Visible: true);
                        // #region agent log
                        try { File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"Documents.Open ok\",\"data\":{},\"sessionId\":\"debug-session\",\"hypothesisId\":\"H5\"}\n"); } catch { }
                        // #endregion
                    }
                    catch (Exception ex)
                    {
                        // #region agent log
                        try { File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"Documents.Open error\",\"data\":{\"error\":\"" + (ex?.Message ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"},\"sessionId\":\"debug-session\",\"hypothesisId\":\"H5\"}\n"); } catch { }
                        // #endregion
                        // ファイルが既に開いている場合は無視
                        System.Diagnostics.Debug.WriteLine($"ドキュメントを開く際のエラー（既に開いている可能性があります）: {ex.Message}");
                    }

                    CurrentProject = project;
                    HideMainWindowRequested?.Invoke(this, EventArgs.Empty);
                    ShowAppBarRequested?.Invoke(this, EventArgs.Empty);
                    BringWordWindowToForeground();
                    ResultMessage = $"Wordファイルを開きました: {Path.GetFileName(project.FilePath)}";
                }
                catch (Exception ex)
                {
                    // #region agent log
                    try { File.AppendAllText(@"c:\Users\kouza\source\repos\MOS Word app\.cursor\debug.log", "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"MainViewModel.cs:ExecuteOpenProject\",\"message\":\"ExecuteOpenProject catch\",\"data\":{\"error\":\"" + (ex?.Message ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"},\"sessionId\":\"debug-session\",\"hypothesisId\":\"H4\"}\n"); } catch { }
                    // #endregion
                    ResultMessage = $"エラー: ファイルを開けませんでした: {ex.Message}";
                    System.Diagnostics.Debug.WriteLine($"エラー詳細: {ex.StackTrace}");
                }
            }
        }

        private void ExecuteResetAllInGroup(object parameter)
        {
            int groupId = 1;
            if (parameter != null)
            {
                if (parameter is int g)
                    groupId = g;
                else if (parameter is string s && int.TryParse(s, out int parsed))
                    groupId = parsed;
            }
            var result = MessageBox.Show(
                $"このタブの全プロジェクトをリセットしますか？\n現在の変更内容は失われます。",
                "すべてリセット確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
                return;
            try
            {
                LogReader.ClearLog();
                for (int projectId = 1; projectId <= 10; projectId++)
                {
                    try
                    {
                        WordProjectResetHelper.ResetProject(groupId, projectId);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ExecuteResetAllInGroup] Project{projectId} リセットエラー: {ex.Message}");
                        MessageBox.Show($"プロジェクト{projectId}のリセット中にエラーが発生しました: {ex.Message}", "リセットエラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                // すべてリセット後、開いているWord文書をすべて閉じる（次に開くときに初期化されたファイルが開く）
                CloseAllWordDocuments();
                // 開いている試験バーがあれば全プロジェクトの解答済み・フラグ状態をクリア
                var uiTestBar = System.Windows.Application.Current.Windows.OfType<Views.UiTestAppBarWindow>().FirstOrDefault();
                if (uiTestBar != null)
                    uiTestBar.ClearAllProjectStates();
                var appBar = System.Windows.Application.Current.Windows.OfType<Views.AppBarWindow>().FirstOrDefault();
                if (appBar != null)
                    appBar.ClearAllProjectStates();
                ResultMessage = "すべてリセットしました。";
                MessageBox.Show("このタブの全プロジェクトをリセットしました。", "リセット完了", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                ResultMessage = $"リセット中にエラーが発生しました: {ex.Message}";
                MessageBox.Show($"リセット中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 開いているすべてのWord文書を閉じる（保存しない）。
        /// すべてリセット実行後に呼び、次に開くときに初期化されたファイルが開くようにする。
        /// </summary>
        private static void CloseAllWordDocuments()
        {
            try
            {
                WordApp wordApp = null;
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch (COMException)
                {
                    return;
                }
                if (wordApp == null) return;
                try
                {
                    while (wordApp.Documents.Count > 0)
                    {
                        WordDoc openDoc = null;
                        try
                        {
                            openDoc = wordApp.Documents[1];
                            openDoc.Close(SaveChanges: false);
                        }
                        catch (COMException comEx) when (comEx.HResult == unchecked((int)0x80010108))
                        {
                            break;
                        }
                        catch (Exception closeEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"[CloseAllWordDocuments] 閉じる際のエラー: {closeEx.Message}");
                        }
                        finally
                        {
                            try { if (openDoc != null) Marshal.ReleaseComObject(openDoc); } catch { }
                        }
                    }
                }
                finally
                {
                    try { if (wordApp != null) Marshal.ReleaseComObject(wordApp); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CloseAllWordDocuments] Error: {ex.Message}");
            }
        }

        /// <summary>
        /// WINWORD プロセスのメインウィンドウ（OpusApp）を検索し、SetForegroundWindow で最前面にする。
        /// </summary>
        private static void BringWordWindowToForeground()
        {
            IntPtr wordHwnd = IntPtr.Zero;
            var wordProcesses = Process.GetProcessesByName("WINWORD");
            if (wordProcesses.Length == 0) return;
            uint targetPid = (uint)wordProcesses[0].Id;
            try
            {
                EnumWindows((hWnd, lParam) =>
                {
                    GetWindowThreadProcessId(hWnd, out uint pid);
                    if (pid != targetPid) return true;
                    var sb = new StringBuilder(256);
                    GetClassName(hWnd, sb, sb.Capacity);
                    if (sb.ToString() != "OpusApp") return true;
                    wordHwnd = hWnd;
                    return false;
                }, IntPtr.Zero);
                if (wordHwnd != IntPtr.Zero)
                    SetForegroundWindow(wordHwnd);
            }
            finally
            {
                foreach (var p in wordProcesses) p.Dispose();
            }
        }

        private void ExecuteUiTest(object parameter)
        {
            ResultMessage = "UIテスト機能は準備中です";
        }

        private void ExecuteTabSearch(object parameter)
        {
            try
            {
                // CSVファイルを読み込む（起動はしない）
                // 複数のパスを試す
                string csvPath = null;
                string[] possiblePaths = new string[]
                {
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tabapp", "CSV", "MOSタブアプリ正誤判定表251121.csv"),
                    Path.Combine(Directory.GetCurrentDirectory(), "tabapp", "CSV", "MOSタブアプリ正誤判定表251121.csv"),
                    Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "tabapp", "CSV", "MOSタブアプリ正誤判定表251121.csv"),
                    Path.Combine(Environment.CurrentDirectory, "tabapp", "CSV", "MOSタブアプリ正誤判定表251121.csv"),
                    @"C:\Users\kouza\source\repos\MOS Word app\tabapp\CSV\MOSタブアプリ正誤判定表251121.csv" // 絶対パス（開発時用）
                };

                foreach (string path in possiblePaths)
                {
                    if (File.Exists(path))
                    {
                        csvPath = path;
                        break;
                    }
                }

                // CSVを読み込んでタスクリストを作成
                TabTasks.Clear();
                if (!string.IsNullOrEmpty(csvPath))
                {
                    try
                    {
                        string[] lines = File.ReadAllLines(csvPath, Encoding.UTF8);
                        for (int i = 1; i < lines.Length; i++) // ヘッダー行をスキップ
                        {
                            if (string.IsNullOrWhiteSpace(lines[i])) continue;
                            string[] parts = lines[i].Split(',');
                            if (parts.Length >= 3)
                            {
                                TabTasks.Add(new TabTaskInfo
                                {
                                    TaskNumber = i,
                                    Question = parts[1].Trim(),
                                    Answer = parts[2].Trim(),
                                    IsPassed = false
                                });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        ResultMessage = $"CSVファイルの読み込みエラー: {ex.Message}";
                    }
                }
                else
                {
                    // CSVファイルが見つからない場合でも、デフォルトのタスクリストを作成
                    ResultMessage = "警告: CSVファイルが見つかりません。デフォルトのタスクリストを使用します。";
                    
                    // デフォルトのタスクリスト
                    TabTasks.Add(new TabTaskInfo { TaskNumber = 1, Question = "挿入タブを探してください", Answer = "挿入タブを選択", IsPassed = false });
                    TabTasks.Add(new TabTaskInfo { TaskNumber = 2, Question = "デザインタブを探してください", Answer = "デザインタブを選択", IsPassed = false });
                    TabTasks.Add(new TabTaskInfo { TaskNumber = 3, Question = "レイアウトタブを探してください", Answer = "レイアウトタブを選択", IsPassed = false });
                    TabTasks.Add(new TabTaskInfo { TaskNumber = 4, Question = "参考資料タブを探してください", Answer = "参考資料タブを選択", IsPassed = false });
                    TabTasks.Add(new TabTaskInfo { TaskNumber = 5, Question = "校閲タブを探してください", Answer = "校閲タブを選択", IsPassed = false });
                    TabTasks.Add(new TabTaskInfo { TaskNumber = 6, Question = "ホームタブを探してください", Answer = "ホームタブを選択", IsPassed = false });
                }

                // Wordを起動（CSVファイルが見つからなくても起動する）
                WordApp wordApp = null;
                try
                {
                    wordApp = (WordApp)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wordApp = new WordApp();
                    wordApp.Visible = true;
                }

                // 新しいドキュメントを作成して問題文を表示
                WordDoc doc = wordApp.Documents.Add();
                Range range = doc.Content;
                range.Text = "タブ探し練習\n\n";
                
                foreach (var task in TabTasks)
                {
                    range.InsertAfter($"タスク{task.TaskNumber}: {task.Question}\n");
                }
                
                range.InsertAfter("\n解答操作のタブをクリックしてください。\n");

                // タブ探しウィンドウを表示
                var tabSearchWindow = new TabSearchWindow(this);
                tabSearchWindow.Show();

                if (string.IsNullOrEmpty(csvPath))
                {
                    ResultMessage = "タブ探しモードを開始しました（CSVファイルが見つかりませんでしたが、デフォルトのタスクリストを使用しています）";
                }
                else
                {
                    ResultMessage = "タブ探しモードを開始しました";
                }
            }
            catch (Exception ex)
            {
                ResultMessage = $"エラー: {ex.Message}";
            }
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

                int projectNumber = CurrentProject.ProjectId;
                // exe と同じ bin\Debug または bin\Release 直下の Dlls サブフォルダのみから読み込む（obj や products は参照しない）
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string dllPath = Path.Combine(baseDir, "Dlls", $"WordChecker1_{projectNumber}.dll");

                if (!File.Exists(dllPath))
                {
                    ResultMessage = $"エラー: チェッカーファイルが見つかりません: WordChecker1_{projectNumber}.dll (bin\\Debug\\Dlls または bin\\Release\\Dlls を確認してください)";
                    return;
                }

                // DLLを読み込む
                Assembly assembly = Assembly.LoadFrom(dllPath);
                string className = $"Libraries.Group1.WordChecker1_{projectNumber}";
                Type checkerType = assembly.GetType(className);

                if (checkerType == null)
                {
                    ResultMessage = $"エラー: クラス '{className}' が見つかりません";
                    return;
                }

                object checkerInstance = Activator.CreateInstance(checkerType);
                int passedCount = 0;
                int totalTasks = 0;

                // 各タスクをチェック（プロジェクトごとにタスク数が異なる）
                int[] taskCounts = { 5, 5, 6, 7, 8, 7, 5, 7, 6, 5 }; // プロジェクト1-10のタスク数
                int maxTasks = projectNumber <= taskCounts.Length ? taskCounts[projectNumber - 1] : 5;

                for (int taskNum = 1; taskNum <= maxTasks; taskNum++)
                {
                    string methodName = $"CheckTask_1_{projectNumber}_{taskNum:D2}";
                    MethodInfo method = checkerType.GetMethod(methodName);

                    if (method != null)
                    {
                        totalTasks++;
                        try
                        {
                            bool result = (bool)method.Invoke(checkerInstance, null);
                            // 採点結果を共有ストアに記録（グループ1固定）
                            ScoreResultStore.RecordResult(1, projectNumber, taskNum, result);
                            
                            TaskResults.Add(new TaskResult
                            {
                                TaskNumber = taskNum,
                                IsPassed = result,
                                TaskName = $"タスク{taskNum}"
                            });

                            if (result)
                            {
                                passedCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            TaskResults.Add(new TaskResult
                            {
                                TaskNumber = taskNum,
                                IsPassed = false,
                                TaskName = $"タスク{taskNum} (エラー: {ex.Message})"
                            });
                        }
                    }
                }

                TotalScore = passedCount;
                MaxScore = totalTasks;

                StringBuilder sb = new StringBuilder();
                sb.AppendLine($"採点完了: {passedCount}/{totalTasks} タスク合格");
                sb.AppendLine($"得点: {passedCount}点 / {totalTasks}点");
                sb.AppendLine();
                sb.AppendLine("詳細:");
                foreach (var task in TaskResults)
                {
                    sb.AppendLine($"  {task.TaskName}: {(task.IsPassed ? "✓ 合格" : "✗ 不合格")}");
                }

                ResultMessage = sb.ToString();
                ScoreCompleted?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExecuteScore] 採点エラー: {ex.Message}\r\n{ex.StackTrace}");
                var msg = new System.Text.StringBuilder();
                msg.AppendLine($"エラー: 採点中にエラーが発生しました: {ex.Message}");
                if (ex.InnerException != null)
                    msg.AppendLine($"内部エラー: {ex.InnerException.Message}");
                msg.AppendLine();
                msg.AppendLine("※VSTOアドインが有効でログが記録されていないと不正解になります。");
                msg.AppendLine($"ログファイル: {Path.Combine(Path.GetTempPath(), "mos_word_log.txt")}");
                ResultMessage = msg.ToString();
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

