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
using System.Threading;
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
        private TabTaskInfo _currentTabTask;
        private ProjectViewModel _currentProject;

        public MainViewModel()
        {
            LoadProjects();
            OpenProjectCommand = new RelayCommand(ExecuteOpenProject);
            TabSearchCommand = new RelayCommand(ExecuteTabSearch);
            ResetAllInGroupCommand = new RelayCommand(ExecuteResetAllInGroup);
            TabTasks = new ObservableCollection<TabTaskInfo>();
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

        /// <summary>ヘッダーの「すべてリセット」用。演習(Group1)固定。</summary>
        public int CurrentGroupIdForReset => 1;

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
        public ICommand TabSearchCommand { get; }
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

        public event EventHandler ShowAppBarRequested;
        public event EventHandler HideMainWindowRequested;
#pragma warning disable 67 // イベントは MainWindow で購読されるため警告を抑制
        public event EventHandler ShowMainWindowRequested;
        public event EventHandler ExamEnded;
#pragma warning restore 67

        private void LoadProjects()
        {
            string basePath = @"C:\MOSTest\Word365";

            // Group1（演習）のみ。保存先は Tab{groupId}\ 直下。一覧は Initial にファイルがあれば表示し、FilePath は作業フォルダ（Tab\）のパスにする
            for (int groupId = 1; groupId <= 1; groupId++)
            {
                string workingFolder = Path.Combine(basePath, $"Tab{groupId}");
                string initialFolder = Path.Combine(basePath, $"Tab{groupId}", "Initial");
                var group = new ProjectGroupViewModel { GroupId = groupId, GroupName = $"Group {groupId}" };

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
            if (!(parameter is ProjectViewModel))
            {
                return;
            }
            var project = (ProjectViewModel)parameter;
            {
                if (string.IsNullOrEmpty(project.FilePath))
                {
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
                    var owner = System.Windows.Application.Current?.MainWindow;
                    string vstoBody = message + "\n\nこのままWordを起動しますか？\n（VSTOが必要なタスクの採点が正しく行われない可能性があります）";
                    MessageBoxResult result = owner != null
                        ? MessageBox.Show(owner, vstoBody, "VSTOアドイン未インストール", MessageBoxButton.YesNo, MessageBoxImage.Warning)
                        : MessageBox.Show(vstoBody, "VSTOアドイン未インストール", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                    if (result == MessageBoxResult.No)
                    {
                        ResultMessage = "Wordの起動をキャンセルしました。";
                        return;
                    }
                }

                try
                {
                    string targetPath = NormalizeDocumentPath(project.FilePath);
                    bool switchingProject = CurrentProject != null
                        && !string.Equals(NormalizeDocumentPath(CurrentProject.FilePath), targetPath, StringComparison.OrdinalIgnoreCase);

                    // 別プロジェクトへ切り替えるときだけ全ドキュメントを保存・閉じる
                    if (switchingProject)
                        SaveAndCloseAllWordDocuments();

                    WordApp wordApp = WordApplicationManager.AcquireWordApplicationForExam(true);

                    if (!TryActivateOpenDocument(wordApp, targetPath))
                    {
                        // 同じパスで既に開いているドキュメントがあれば保存してから閉じ、フォルダから開き直す
                        try
                        {
                            for (int i = wordApp.Documents.Count; i >= 1; i--)
                            {
                                WordDoc openDoc = wordApp.Documents[i];
                                try
                                {
                                    if (DocumentPathsEqual(openDoc.FullName, targetPath))
                                    {
                                        if (!openDoc.Saved)
                                            openDoc.Save();
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

                        try
                        {
                            wordApp.Documents.Open(project.FilePath, ReadOnly: false, Visible: true);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"ドキュメントを開く際のエラー（既に開いている可能性があります）: {ex.Message}");
                        }
                    }

                    CurrentProject = project;
                    HideMainWindowRequested?.Invoke(this, EventArgs.Empty);
                    ShowAppBarRequested?.Invoke(this, EventArgs.Empty);
                    ApplyExamWindowLayoutFromOpenProject();
                    BringWordWindowToForeground();
                    ResultMessage = $"Wordファイルを開きました: {Path.GetFileName(project.FilePath)}";
                }
                catch (Exception ex)
                {
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
                LogReader.ClearTaskEvidence(); // 旧 mos_word_task_evidence.txt が残っていれば削除のみ
                LogReader.ClearDestructiveLog();
                LogReader.ClearSnapshot();
                LogReader.ClearCurrentTaskFile();
                WordTaskAttemptRegistry.ClearAll();
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
        /// 開いているすべての Word 文書を保存する。
        /// </summary>
        private static void SaveAllWordDocuments()
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
                    wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    for (int i = wordApp.Documents.Count; i >= 1; i--)
                    {
                        WordDoc doc = null;
                        try
                        {
                            doc = wordApp.Documents[i];
                            if (!doc.Saved)
                                doc.Save();
                        }
                        catch (COMException comEx) when (comEx.HResult == unchecked((int)0x80010108))
                        {
                            break;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[SaveAllWordDocuments] 保存エラー: {ex.Message}");
                        }
                        finally
                        {
                            try { if (doc != null) Marshal.ReleaseComObject(doc); } catch { }
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
                System.Diagnostics.Debug.WriteLine($"[SaveAllWordDocuments] Error: {ex.Message}");
            }
        }

        /// <summary>
        /// プロジェクト切替前に作業内容をディスクに残して Word 文書を閉じる。
        /// </summary>
        private static void SaveAndCloseAllWordDocuments()
        {
            SaveAllWordDocuments();
            CloseAllWordDocuments();
        }

        private static string NormalizeDocumentPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;
            try
            {
                return Path.GetFullPath(path).ToLowerInvariant();
            }
            catch
            {
                return path.ToLowerInvariant();
            }
        }

        private static bool DocumentPathsEqual(string left, string rightNormalized)
        {
            if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(rightNormalized))
                return false;
            return string.Equals(NormalizeDocumentPath(left), rightNormalized, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>対象ファイルが既に開いていればアクティブ化して true を返す。</summary>
        private static bool TryActivateOpenDocument(WordApp wordApp, string targetPathNormalized)
        {
            if (wordApp == null || string.IsNullOrEmpty(targetPathNormalized))
                return false;

            try
            {
                for (int i = wordApp.Documents.Count; i >= 1; i--)
                {
                    WordDoc doc = wordApp.Documents[i];
                    try
                    {
                        if (!DocumentPathsEqual(doc.FullName, targetPathNormalized))
                            continue;

                        doc.Activate();
                        try { doc.ActiveWindow?.Activate(); } catch { }
                        return true;
                    }
                    finally
                    {
                        try { if (doc != null) Marshal.ReleaseComObject(doc); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TryActivateOpenDocument] {ex.Message}");
            }

            return false;
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

        /// <summary>プロジェクト再開時に試験用レイアウトを適用する（採点用サイズのまま残るのを防ぐ）。</summary>
        private static void ApplyExamWindowLayoutFromOpenProject()
        {
            var uiTestBar = System.Windows.Application.Current.Windows
                .OfType<Views.UiTestAppBarWindow>().FirstOrDefault();
            if (uiTestBar != null)
            {
                uiTestBar.ApplyExamWindowLayout();
                return;
            }

            var appBar = System.Windows.Application.Current.Windows
                .OfType<Views.AppBarWindow>().FirstOrDefault();
            appBar?.ApplyExamWindowLayout();
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

