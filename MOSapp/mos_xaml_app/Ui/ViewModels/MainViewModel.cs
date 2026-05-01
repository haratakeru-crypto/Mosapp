using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Core.Ports.Primary;
using MOSExcelMogiApp.Views;
using Newtonsoft.Json.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;

namespace Ui.ViewModels
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
        [ComImport]
        [Guid("00000016-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleMessageFilter
        {
            [PreserveSig]
            int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);

            [PreserveSig]
            int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);

            [PreserveSig]
            int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
        }

        /// <summary>
        /// Office COM 呼び出し中の RPC_E_CALL_REJECTED を自動リトライする OLE message filter。
        /// この ViewModel 内に閉じた実装にして、csproj 取り込み漏れの影響を受けないようにする。
        /// </summary>
        private sealed class OleMessageFilterScope : IOleMessageFilter, IDisposable
        {
            private IOleMessageFilter _oldFilter;
            private bool _disposed;

            private OleMessageFilterScope() { }

            public static OleMessageFilterScope Enter()
            {
                var scope = new OleMessageFilterScope();
                CoRegisterMessageFilter(scope, out scope._oldFilter);
                return scope;
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                CoRegisterMessageFilter(_oldFilter, out _);
                _oldFilter = null;
            }

            int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) => 0;
            int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) => dwRejectType == 2 ? 100 : -1;
            int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) => 2;

            [DllImport("Ole32.dll")]
            private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
        }

        private readonly IExcelCheckerService _excelCheckerService;
        private int _selectedTabIndex;
        private string _selectedFilePath;
        private string _resultMessage;
        private bool _isExcelOverlayVisible;
        private bool _isShutdownWaitOverlayVisible;
        private bool _showScoreButton;
        private bool _showPauseButton;
        private ProjectInfo _currentProject;
        private ExcelApp _sharedExcelApp;

        /// <summary>試験終了処理の二重起動防止（タイマー経路と終了ボタン確認の競合など）。新規プロジェクト開始時に 0 に戻す。</summary>
        private int _endExamShutdownStarted;

        /// <summary>試験終了スレッドの保存・Quit が終わるまでシグナル。初期は完了済み。</summary>
        private readonly ManualResetEventSlim _excelShutdownFinished = new ManualResetEventSlim(true);

        /// <summary>
        /// アプリが利用する Excel インスタンスを取得（無ければ作成）。
        /// Process.Start による別インスタンス起動や GetActiveObject の取り違えを避けるため、1インスタンスに固定して使い回す。
        /// </summary>
        public ExcelApp GetOrCreateExcelApplication()
        {
            if (_sharedExcelApp != null)
            {
                try
                {
                    _ = _sharedExcelApp.Visible;
                    return _sharedExcelApp;
                }
                catch
                {
                    _sharedExcelApp = null;
                }
            }

            // 失敗時の孤児プロセス掃除は ExcelApplicationManager 側の finally で行う。
            // COM が間に合わない一過性の失敗向けに、短い待機のあと 1 回だけ再試行（計 2 回）。
            const int maxAttempts = 2;
            const int retryDelayMs = 800;
            Exception lastException = null;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    _sharedExcelApp = Libraries.ExcelApplicationManager.GetOrCreateExcelApplication(
                        makeVisible: true,
                        timeoutMs: 30000);
                    return _sharedExcelApp;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    _sharedExcelApp = null;
                    if (attempt < maxAttempts)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[GetOrCreateExcelApplication] attempt {attempt} failed: {ex.Message}; retry after {retryDelayMs}ms");
                        Thread.Sleep(retryDelayMs);
                    }
                }
            }

            throw lastException ?? new InvalidOperationException("Excel を起動できませんでした。");
        }

        /// <summary>
        /// 既存の共有 Excel 参照を返す（無効なら null）。新規起動はしない。
        /// </summary>
        public ExcelApp TryGetSharedExcelApplication()
        {
            if (_sharedExcelApp == null) return null;
            try
            {
                _ = _sharedExcelApp.Hwnd;
                return _sharedExcelApp;
            }
            catch
            {
                _sharedExcelApp = null;
                return null;
            }
        }

        /// <summary>
        /// 結果に戻る等で Excel プロセスを外部から終了したあと、無効な COM 参照を捨てる。
        /// </summary>
        public void ClearSharedExcelApplication()
        {
            _sharedExcelApp = null;
        }
        
        public event EventHandler ExamEnded;
        public event EventHandler ShowAppBarRequested;
        public event EventHandler HideMainWindowRequested;
        public event EventHandler ShowMainWindowRequested;
        public event EventHandler UiTestRequested;
        public event EventHandler CurrentProjectChanged;
        public event EventHandler OpenReviewPageRequested;

        /// <summary>
        /// シェル起動後に共有 Excel への接続試行が UI スレッドで終わったときに発火する。アプリバーが Excel を再配置するために使う。
        /// </summary>
        public event EventHandler SharedExcelApplicationAttached;

        public MainViewModel(IExcelCheckerService excelCheckerService)
        {
            _excelCheckerService = excelCheckerService;
            LoadProjects();
            CheckCommand = new RelayCommand(ExecuteCheck);
            OpenProjectCommand = new RelayCommand(ExecuteOpenProject);
            ScoreCommand = new RelayCommand(p => ExecuteScoreAsync(p));
            EndExamCommand = new RelayCommand(ExecuteEndExam);
            PauseExamCommand = new RelayCommand(ExecutePauseExam);
            ResetExamCommand = new RelayCommand(ExecuteResetExam);
            NextProjectCommand = new RelayCommand(ExecuteNextProject);
            UiTestCommand = new RelayCommand(ExecuteUiTest);
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

        public string SelectedFilePath
        {
            get => _selectedFilePath;
            set
            {
                _selectedFilePath = value;
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

        public ICommand CheckCommand { get; }
        public ICommand OpenProjectCommand { get; }
        public ICommand ScoreCommand { get; }
        public ICommand EndExamCommand { get; }
        public ICommand PauseExamCommand { get; }
        public ICommand ResetExamCommand { get; }
        public ICommand NextProjectCommand { get; }
        public ICommand UiTestCommand { get; }

    public bool IsExcelOverlayVisible
        {
            get => _isExcelOverlayVisible;
            set
            {
                _isExcelOverlayVisible = value;
                OnPropertyChanged();
            }
        }

        /// <summary>前試験の Excel 終了待ちの全画面オーバーレイ（プロジェクト選択で連打したときのフィードバック）</summary>
        public bool IsShutdownWaitOverlayVisible
        {
            get => _isShutdownWaitOverlayVisible;
            set
            {
                if (_isShutdownWaitOverlayVisible == value) return;
                _isShutdownWaitOverlayVisible = value;
                OnPropertyChanged();
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const int SW_RESTORE = 9;

        private static void TryForceForeground(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return;

            uint currentTid = 0;
            uint targetTid = 0;
            bool attached = false;

            try
            {
                currentTid = GetCurrentThreadId();
                targetTid = GetWindowThreadProcessId(hWnd, out _);

                if (currentTid != 0 && targetTid != 0 && currentTid != targetTid)
                {
                    attached = AttachThreadInput(currentTid, targetTid, true);
                }

                ShowWindow(hWnd, SW_RESTORE);
                BringWindowToTop(hWnd);
                SetForegroundWindow(hWnd);
            }
            catch
            {
                /* ignore */
            }
            finally
            {
                if (attached && currentTid != 0 && targetTid != 0)
                {
                    try { AttachThreadInput(currentTid, targetTid, false); } catch { /* ignore */ }
                }
            }
        }

        private static IntPtr TryGetMainWindowByProcessId(int pid, int timeoutMs = 5000)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                IntPtr found = IntPtr.Zero;
                EnumWindows((hWnd, lParam) =>
                {
                    if (!IsWindowVisible(hWnd)) return true;
                    GetWindowThreadProcessId(hWnd, out uint windowPid);
                    if (windowPid == (uint)pid)
                    {
                        found = hWnd;
                        return false;
                    }
                    return true;
                }, IntPtr.Zero);
                if (found != IntPtr.Zero) return found;
                Thread.Sleep(50);
            }
            return IntPtr.Zero;
        }

        private void LaunchAndPositionNotepad()
        {
            int widthPx = 1920;
            int heightPx = 667;

            var psi = new ProcessStartInfo
            {
                FileName = "excel.exe",
                UseShellExecute = true
            };
            var proc = Process.Start(psi);
            if (proc == null) return;

            try { proc.WaitForInputIdle(3000); } catch { }

            IntPtr hWnd = IntPtr.Zero;
            for (int i = 0; i < 20; i++)
            {
                proc.Refresh();
                hWnd = proc.MainWindowHandle;
                if (hWnd != IntPtr.Zero) break;
                Thread.Sleep(50);
            }
            if (hWnd == IntPtr.Zero)
            {
                hWnd = TryGetMainWindowByProcessId(proc.Id, 5000);
                if (hWnd == IntPtr.Zero) return;
            }

            ShowWindow(hWnd, SW_RESTORE);
            bool ok = SetWindowPos(hWnd, IntPtr.Zero, 0, 0, widthPx, heightPx, SWP_NOZORDER | SWP_SHOWWINDOW);
            if (!ok)
            {
                MoveWindow(hWnd, 0, 0, widthPx, heightPx, true);
            }
        }

        private static IntPtr FindTopWindowByProcessName(string processName, int timeoutMs = 6000)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                IntPtr found = IntPtr.Zero;
                EnumWindows((hWnd, lParam) =>
                {
                    if (!IsWindowVisible(hWnd)) return true;
                    GetWindowThreadProcessId(hWnd, out uint windowPid);
                    try
                    {
                        var p = Process.GetProcessById((int)windowPid);
                        if (string.Equals(p.ProcessName, processName, StringComparison.OrdinalIgnoreCase))
                        {
                            found = hWnd;
                            return false;
                        }
                    }
                    catch { }
                    return true;
                }, IntPtr.Zero);
                if (found != IntPtr.Zero) return found;
                Thread.Sleep(50);
            }
            return IntPtr.Zero;
        }

        private void LaunchAndPositionExcel(int appBarHeight)
        {
            int widthPx = 1920;
            // 画面の高さからアプリバーの高さを引いた値に合わせる
            int screenHeight = (int)SystemParameters.PrimaryScreenHeight;
            int heightPx = Math.Max(100, screenHeight - appBarHeight);

            // Excel 実行パス候補
            string[] candidates = new[]
            {
                "excel.exe",
                @"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
                @"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            };

            Process proc = null;
            foreach (var path in candidates)
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = path,
                        Arguments = "/e",
                        UseShellExecute = true
                    };
                    proc = Process.Start(psi);
                    if (proc != null) break;
                }
                catch { }
            }
            if (proc == null) return;

            // Excel のトップレベル HWND を探索
            try { proc.WaitForInputIdle(5000); } catch { }

            IntPtr hWnd = IntPtr.Zero;
            // まずは起動した PID から
            hWnd = TryGetMainWindowByProcessId(proc.Id, 4000);
            if (hWnd == IntPtr.Zero)
            {
                // 見つからなければプロセス名で総当たり（EXCEL）
                hWnd = FindTopWindowByProcessName("EXCEL", 6000);
                if (hWnd == IntPtr.Zero) return;
            }

            TryForceForeground(hWnd);
            bool ok = SetWindowPos(hWnd, IntPtr.Zero, 0, 0, widthPx, heightPx, SWP_NOZORDER | SWP_SHOWWINDOW);
            if (!ok)
            {
                MoveWindow(hWnd, 0, 0, widthPx, heightPx, true);
            }
        }

        /// <summary>
        /// Excel 起動直後の前面化遅延を抑えるため、短時間リトライで可視化/前面化を確定させる。
        /// </summary>
        private void EnsureExcelWindowVisibleAndForeground(ExcelApp excelApp, ExcelWorkbook targetWorkbook)
        {
            if (excelApp == null) return;

            for (int i = 0; i < 3; i++)
            {
                try { excelApp.Visible = true; } catch { }
                try { excelApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal; } catch { }
                try { targetWorkbook?.Activate(); } catch { }
                try { excelApp.ActiveWindow?.Activate(); } catch { }

                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = new IntPtr(excelApp.Hwnd); } catch { hwnd = IntPtr.Zero; }

                if (hwnd != IntPtr.Zero)
                {
                    TryForceForeground(hwnd);
                }

                if (i < 2)
                {
                    Thread.Sleep(100);
                }
            }

            // UI スレッドの入力サイクル後にもう一度だけ前面化を試す
            try
            {
                Application.Current?.Dispatcher?.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        IntPtr hwnd = new IntPtr(excelApp.Hwnd);
                        if (hwnd != IntPtr.Zero)
                        {
                            TryForceForeground(hwnd);
                        }
                    }
                    catch
                    {
                        /* ignore */
                    }
                }), DispatcherPriority.Background);
            }
            catch
            {
                /* ignore */
            }
        }

        public ProjectInfo CurrentProject
        {
            get => _currentProject;
            set
            {
                _currentProject = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CurrentProjectName));
                OnPropertyChanged(nameof(IsNextProjectVisible));
                CurrentProjectChanged?.Invoke(this, EventArgs.Empty);
            }
        }
        
        public string CurrentProjectName => CurrentProject?.Name ?? "プロジェクトが選択されていません";

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
        
        public bool IsNextProjectVisible => CurrentProject != null && CurrentProject.ProjectNumber <= 10;

        private void LoadProjects()
        {
            var allProjects = _excelCheckerService.GetAllProjects();
            
            // 模試①=Group2は非表示のためスキップ。Group1（演習）とGroup3（応用編）のみ追加
            foreach (int groupId in new[] { 1, 3 })
            {
                var group = new ProjectGroupViewModel { GroupId = groupId, GroupName = groupId == 3 ? "応用編" : $"Group {groupId}" };
                
                for (int projectId = 1; projectId <= 10; projectId++)
                {
                    group.Projects.Add(new ProjectViewModel
                    {
                        GroupId = groupId,
                        ProjectId = projectId,
                        Name = $"プロジェクト{groupId}-{projectId}"
                    });
                }
                
                ProjectGroups.Add(group);
            }
        }

        private void ExecuteCheck(object parameter)
        {
            if (parameter is ProjectViewModel project && !string.IsNullOrEmpty(SelectedFilePath))
            {
                bool result = _excelCheckerService.CheckExcel(project.GroupId, project.ProjectId, SelectedFilePath);
                ResultMessage = $"Project{project.GroupId}-{project.ProjectId}: {(result ? "Success" : "Failed")}";
            }
        }

        private void ExecuteOpenProject(object parameter)
        {
            // Excel重複起動チェックを無効化（修正）
            // if (IsExcelRunning())
            // {
            //     ResultMessage = "警告: Excelが既に開いています。先にExcelを閉じてください。";
            //     return;
            // }

            string projectId;
            if (parameter is ProjectViewModel projectViewModel)
            {
                // Handle ProjectViewModel object
                projectId = $"project{projectViewModel.GroupId}-{projectViewModel.ProjectId}";
                System.Diagnostics.Debug.WriteLine($"Converted ProjectViewModel to projectId: {projectId}");
            }
            else
            {
                // Handle string parameter
                projectId = parameter?.ToString();
            }
            
            if (string.IsNullOrEmpty(projectId))
            {
                ResultMessage = "エラー: プロジェクトIDが指定されていません。";
                return;
            }

            if (!WaitForExcelShutdownToCompleteBeforeOpeningProject())
                return;

            string filePath = GetProjectFilePath(projectId);
            if (string.IsNullOrEmpty(filePath))
            {
                ResultMessage = $"エラー: プロジェクトID '{projectId}' からファイルパスを取得できませんでした。";
                System.Diagnostics.Debug.WriteLine($"Failed to get file path for projectId: {projectId}");
                return;
            }
            
            if (!File.Exists(filePath))
            {
                ResultMessage = $"エラー: ファイルが見つかりません: {filePath}\n\nファイルが存在するか確認してください。";
                System.Diagnostics.Debug.WriteLine($"File not found: {filePath}");
                return;
            }

            // Initialフォルダのファイルの場合、該当するInitialフォルダ内のすべてのファイルの読み取り専用属性を解除
            if (filePath.Contains("Initial"))
            {
                try
                {
                    // Initialフォルダのパスを取得
                    string initialFolderPath = System.IO.Path.GetDirectoryName(filePath);
                    
                    if (Directory.Exists(initialFolderPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[ExecuteOpenProject] Removing read-only attributes from all files in Initial folder: {initialFolderPath}");
                        
                        // Initialフォルダ内のすべてのExcelファイルの読み取り専用属性を解除
                        string[] excelFiles = Directory.GetFiles(initialFolderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
                        int removedCount = 0;
                        
                        foreach (string excelFile in excelFiles)
                        {
                            try
                            {
                                FileInfo fileInfo = new FileInfo(excelFile);
                                if (fileInfo.IsReadOnly)
                                {
                                    fileInfo.IsReadOnly = false;
                                    removedCount++;
                                    System.Diagnostics.Debug.WriteLine($"[ExecuteOpenProject] Removed read-only attribute from: {System.IO.Path.GetFileName(excelFile)}");
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[ExecuteOpenProject] Error removing read-only attribute from {excelFile}: {ex.Message}");
                            }
                        }
                        
                        System.Diagnostics.Debug.WriteLine($"[ExecuteOpenProject] Removed read-only attributes from {removedCount} file(s) in Initial folder");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[ExecuteOpenProject] Initial folder does not exist: {initialFolderPath}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExecuteOpenProject] Error processing Initial folder files: {ex.Message}");
                    // 読み取り専用属性の解除に失敗しても、ファイルを開く処理は続行
                }
            }

            try
            {
                // 主経路: 既定の関連付けで開く（余計な excel.exe 起動による Book1 を避ける）。失敗時のみ excel.exe にパスを渡す。
                bool opened = false;
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    });
                    opened = true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExecuteOpenProject] shell-open failed: {ex.Message}");
                }

                if (!opened)
                    opened = StartExcelWithFile(filePath);

                if (!opened)
                {
                    ResultMessage = "エラー: Excel を起動できませんでした。";
                    return;
                }

                Interlocked.Exchange(ref _endExamShutdownStarted, 0);

                if (parameter is ProjectViewModel pvm)
                {
                    CurrentProject = new ProjectInfo
                    {
                        Name = pvm.Name,
                        FilePath = filePath,
                        Group = $"Group {pvm.GroupId}",
                        ProjectNumber = pvm.ProjectId
                    };
                }
                else
                {
                    var parts = projectId.Split('-');
                    string groupName = parts.Length > 0 ? $"Group {parts[0].Replace("tab", "").Replace("project", "")}" : "Unknown";
                    string projectName = parts.Length > 1 ? $"Project {parts[1]}" : "Unknown";
                    CurrentProject = new ProjectInfo
                    {
                        Name = $"{groupName} - {projectName}",
                        FilePath = filePath,
                        Group = groupName,
                        ProjectNumber = parts.Length > 1 && int.TryParse(parts[1], out int num) ? num : 0
                    };
                }

                IsExcelOverlayVisible = true;
                ResultMessage = $"Excelファイルを開きました: {Path.GetFileName(filePath)}";
                TryAttachSharedExcelApplicationAfterShellOpen();

                Application.Current?.Dispatcher?.BeginInvoke(new Action(ShowAppBar), DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExecuteOpenProject] failed error={ex.GetType().Name}:{ex.Message}");
                ResultMessage = $"エラー: ファイルを開けませんでした: {ex.Message}";
            }
        }

        private bool IsExcelRunning()
        {
            Process[] excelProcesses = Process.GetProcessesByName("EXCEL");
            return excelProcesses.Length > 0;
        }

        private static bool StartExcelWithFile(string filePath)
        {
            string[] candidates = new[]
            {
                "excel.exe",
                @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
            };

            foreach (var excelPath in candidates)
            {
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = excelPath,
                        Arguments = $"\"{filePath}\"",
                        UseShellExecute = true
                    };
                    if (Process.Start(psi) != null)
                        return true;
                }
                catch
                {
                    // try next
                }
            }

            return false;
        }

        /// <summary>
        /// シェルでブックを開いたあと、UI をブロックせず ROT へ接続して <see cref="_sharedExcelApp"/> を設定する。
        /// 失敗してもユーザー操作は完了済みのため例外は握りつぶす。
        /// </summary>
        private void TryAttachSharedExcelApplicationAfterShellOpen()
        {
            Task.Run(() =>
            {
                Thread.Sleep(1500);
                var disp = Application.Current?.Dispatcher;
                if (disp == null)
                    return;

                disp.BeginInvoke(DispatcherPriority.Background, new Action(() =>
                {
                    try
                    {
                        if (_sharedExcelApp != null)
                        {
                            try
                            {
                                _ = _sharedExcelApp.Visible;
                                return;
                            }
                            catch
                            {
                                _sharedExcelApp = null;
                            }
                        }

                        using (OleMessageFilterScope.Enter())
                        {
                            var attached = Libraries.ExcelApplicationManager.TryAttachRunningExcelApplication(
                                makeVisible: true,
                                timeoutMs: 15000);
                            if (attached != null)
                                _sharedExcelApp = attached;
                            else
                                System.Diagnostics.Debug.WriteLine(
                                    "[TryAttachSharedExcelApplicationAfterShellOpen] no running Excel in ROT (skipped new launch)");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[TryAttachSharedExcelApplicationAfterShellOpen] {ex.Message}");
                    }
                    finally
                    {
                        SharedExcelApplicationAttached?.Invoke(this, EventArgs.Empty);
                    }
                }));
            });
        }

        /// <summary>
        /// リセット完了後など、空の Excel を先に COM 起動せずブックを開き直す（<see cref="ExecuteOpenProject"/> と同様）。
        /// </summary>
        public void OpenExcelWorkbookAfterResetByShell(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                System.Diagnostics.Debug.WriteLine($"[OpenExcelWorkbookAfterResetByShell] skip: invalid path {filePath}");
                return;
            }

            bool opened = false;
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
                opened = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OpenExcelWorkbookAfterResetByShell] shell-open failed: {ex.Message}");
            }

            if (!opened)
                opened = StartExcelWithFile(filePath);

            if (!opened)
            {
                System.Diagnostics.Debug.WriteLine("[OpenExcelWorkbookAfterResetByShell] failed to start Excel");
                return;
            }

            if (CurrentProject != null)
            {
                CurrentProject = new ProjectInfo
                {
                    Name = CurrentProject.Name,
                    FilePath = filePath,
                    Group = CurrentProject.Group,
                    ProjectNumber = CurrentProject.ProjectNumber
                };
            }

            TryAttachSharedExcelApplicationAfterShellOpen();
        }

        public string GetProjectFilePath(int groupId, int projectId)
        {
            // プロジェクトを開く場合は、必ずInitialフォルダから開く
            string initialPath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\Initial\\project{projectId}.xlsx";
            
            // Initialフォルダにファイルが存在する場合はそれを使用
            if (File.Exists(initialPath))
            {
                System.Diagnostics.Debug.WriteLine($"Using Initial folder file: {initialPath}");
                return initialPath;
            }
            
            // Initialフォルダにファイルが存在しない場合のフォールバック処理
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    string jsonContent = File.ReadAllText(configPath);
                    JObject config = JObject.Parse(jsonContent);
                    
                    var projectConfig = config["tabs"]?[groupId.ToString()]?["projects"]?[projectId.ToString()];
                    if (projectConfig != null)
                    {
                        // initialDataFileを試す
                        string initialDataFile = projectConfig["initialDataFile"]?.ToString();
                        if (!string.IsNullOrEmpty(initialDataFile) && File.Exists(initialDataFile))
                        {
                            System.Diagnostics.Debug.WriteLine($"Using initialDataFile from config: {initialDataFile}");
                            return initialDataFile;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error reading config.json: {ex.Message}");
            }
            
            // それでも見つからない場合は、Initialフォルダのパスを返す（ファイルが後で作成される可能性がある）
            System.Diagnostics.Debug.WriteLine($"Using Initial folder path (file may not exist): {initialPath}");
            return initialPath;
        }

        private string GetProjectFilePath(string projectId)
        {
            // Parse projectId like "tab1-project1-1" or "project1-2"
            System.Diagnostics.Debug.WriteLine($"GetProjectFilePath called with projectId: {projectId}");
            var parts = projectId.Split('-');
            
            if (parts.Length >= 2)
            {
                int groupId = 0;
                int projId = 0;
                
                // Handle format like "project1-2" (groupId-projectId)
                if (parts[0].StartsWith("project"))
                {
                    string groupPart = parts[0].Replace("project", "");
                    string projectPart = parts[1];
                    
                    System.Diagnostics.Debug.WriteLine($"Parsing project format: groupPart={groupPart}, projectPart={projectPart}");
                    
                    if (int.TryParse(groupPart, out groupId) && int.TryParse(projectPart, out projId))
                    {
                        return GetProjectFilePath(groupId, projId);
                    }
                }
                // Handle format like "tab1-project1-1"
                else if (parts[0].StartsWith("tab"))
                {
                    string tabPart = parts[0].Replace("tab", "");
                    string projectPart = parts[1].Replace("project", "");
                    
                    System.Diagnostics.Debug.WriteLine($"Parsing tab format: tabPart={tabPart}, projectPart={projectPart}");
                    
                    if (int.TryParse(tabPart, out groupId) && int.TryParse(projectPart, out projId))
                    {
                        return GetProjectFilePath(groupId, projId);
                    }
                }
            }
            System.Diagnostics.Debug.WriteLine($"Failed to parse projectId: {projectId}");
            return string.Empty;
        }

        private async void ExecuteScoreAsync(object parameter)
        {
            if (CurrentProject == null)
            {
                MessageBox.Show("プロジェクトが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int groupId = int.Parse(CurrentProject.Group.Replace("Group ", ""));
            int projectId = CurrentProject.ProjectNumber;
            var config = LoadConfig();
            var projectConfig = GetProjectConfig(config, groupId, projectId);
            if (projectConfig == null)
            {
                MessageBox.Show("プロジェクト設定が見つかりませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            int taskCount = projectConfig["taskCount"].Value<int>();
            string libraryName = $"ExcelChecker{groupId}_{projectId}";

            // 「採点中です」オーバーレイを表示
            Window scoringOverlay = null;
            Application.Current.Dispatcher.Invoke(() =>
            {
                var owner = Application.Current.MainWindow;
                scoringOverlay = new Window
                {
                    Title = "採点中",
                    Width = 320,
                    Height = 140,
                    WindowStyle = WindowStyle.None,
                    WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen,
                    Owner = owner,
                    ShowInTaskbar = false,
                    ResizeMode = ResizeMode.NoResize,
                    Background = new SolidColorBrush(Color.FromRgb(255, 255, 255)),
                    BorderBrush = new SolidColorBrush(Color.FromRgb(30, 64, 175)),
                    BorderThickness = new Thickness(2)
                };
                var stack = new StackPanel
                {
                    Margin = new Thickness(24),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                var text = new TextBlock
                {
                    Text = "採点中です",
                    FontSize = 18,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 12),
                    Foreground = new SolidColorBrush(Color.FromRgb(30, 64, 175))
                };
                var progress = new ProgressBar
                {
                    IsIndeterminate = true,
                    Height = 20,
                    Width = 260
                };
                stack.Children.Add(text);
                stack.Children.Add(progress);
                scoringOverlay.Content = stack;
                scoringOverlay.Show();
            });
            await Task.Delay(80);

            List<bool> results = null;
            try
            {
                results = await Task.Run(() => ExecuteScoringDirect(libraryName, taskCount));
            }
            catch (Exception ex)
            {
                ResultMessage = $"エラー: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"Error: {ex}");
                Application.Current.Dispatcher.Invoke(() =>
                {
                    try { scoringOverlay?.Close(); } catch { }
                    MessageBox.Show($"採点中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                });
                return;
            }

            // オーバーレイを閉じてから結果ダイアログを表示
            Application.Current.Dispatcher.Invoke(() =>
            {
                try { scoringOverlay?.Close(); } catch { }
            });
            await Task.Delay(80);

            Application.Current.Dispatcher.Invoke(() =>
            {
                try
                {
                    MOSExcelMogiApp.Models.ExamResultStorage.SaveProjectResult(projectId, results);
                    var dialog = new ScoringResultDialog(taskCount, results, groupId, projectId);
                    dialog.ShowDialog();
                    ResultMessage = $"採点完了: {taskCount}問のタスクを採点しました";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"結果の表示中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }

        private JObject LoadConfig()
        {
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
            string jsonContent = File.ReadAllText(configPath);
            return JObject.Parse(jsonContent);
        }
        
        private JToken GetProjectConfig(JObject config, int groupId, int projectId)
        {
            return config["tabs"]?[groupId.ToString()]?["projects"]?[projectId.ToString()];
        }
        
        private List<bool> ExecuteScoring(string libraryName, int taskCount)
        {
            var results = new List<bool>();
            
            try
            {
                // Load the DLL
                Console.WriteLine($"[DEBUG] ExecuteScoring called with libraryName: {libraryName}, taskCount: {taskCount}");
                
                // Try multiple DLL paths in order of preference
                string[] dllPaths = {
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "Debug", "Libraries", $"Group{libraryName.Last()}", $"{libraryName}.dll"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", $"Group{libraryName.Last()}", $"{libraryName}.dll"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", "bin", "Debug", "net48", $"{libraryName}.dll")
                };
                
                string dllPath = null;
                foreach (string path in dllPaths)
                {
                    Console.WriteLine($"[DEBUG] Trying DLL path: {path}");
                    if (File.Exists(path))
                    {
                        dllPath = path;
                        Console.WriteLine($"[DEBUG] Found DLL at: {dllPath}");
                        break;
                    }
                }
                
                if (dllPath == null)
                {
                    Console.WriteLine($"[DEBUG] DLL not found, falling back to source: {dllPath}");
                    // If DLL doesn't exist, try to use the compiled class directly
                    return ExecuteScoringFromSource(libraryName, taskCount);
                }
                
                Console.WriteLine($"[DEBUG] Loading DLL from: {dllPath}");
                Assembly assembly = Assembly.LoadFrom(dllPath);
                Type checkerType = assembly.GetTypes().FirstOrDefault(t => t.Name == libraryName);
                Console.WriteLine($"[DEBUG] Found type: {checkerType?.Name ?? "null"}");
                
                if (checkerType != null)
                {
                    object checkerInstance = Activator.CreateInstance(checkerType);
                    
                    for (int i = 1; i <= taskCount; i++)
                    {
                        string methodName = $"CheckTask_{libraryName.Replace("ExcelChecker", "")}_0{i}";
                        MethodInfo method = checkerType.GetMethod(methodName);
                        
                        if (method != null)
                        {
                            bool result = (bool)method.Invoke(checkerInstance, null);
                            results.Add(result);
                        }
                        else
                        {
                            results.Add(false); // Method not found
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DLL loading error: {ex.Message}");
                // Fallback to source-based execution
                return ExecuteScoringFromSource(libraryName, taskCount);
            }
            
            return results;
        }
        
        private List<bool> ExecuteScoringFromSource(string libraryName, int taskCount)
        {
            var results = new List<bool>();
            
            try
            {
                // Use reflection to load and execute CheckTask methods from DLL
                string dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", GetGroupFolder(libraryName), $"{libraryName}.dll");
                
                if (File.Exists(dllPath))
                {
                    Assembly assembly = Assembly.LoadFrom(dllPath);
                    Type checkerType = assembly.GetType(libraryName);
                    
                    if (checkerType != null)
                    {
                        object checkerInstance = Activator.CreateInstance(checkerType);
                        
                        for (int i = 1; i <= taskCount; i++)
                        {
                            string methodName = GetCheckTaskMethodName(libraryName, i);
                            MethodInfo method = checkerType.GetMethod(methodName);
                            
                            if (method != null)
                            {
                                bool result = (bool)method.Invoke(checkerInstance, null);
                                results.Add(result);
                            }
                            else
                            {
                                results.Add(false);
                            }
                        }
                    }
                }
                else
                {
                    // DLL not found, try to find in loaded assemblies
                    string namespaceName = $"Libraries.Group{libraryName.Last()}";
                    string fullTypeName = $"{namespaceName}.{libraryName}";
                    
                    Type checkerType = Type.GetType(fullTypeName);
                    if (checkerType == null)
                    {
                        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                        {
                            checkerType = assembly.GetType(fullTypeName);
                            if (checkerType != null) break;
                        }
                    }
                    
                    if (checkerType != null)
                    {
                        object checkerInstance = Activator.CreateInstance(checkerType);
                        
                        for (int i = 1; i <= taskCount; i++)
                        {
                            string methodName = $"CheckTask_{libraryName.Replace("ExcelChecker", "")}_0{i}";
                            MethodInfo method = checkerType.GetMethod(methodName);
                            
                            if (method != null)
                            {
                                bool result = (bool)method.Invoke(checkerInstance, null);
                                results.Add(result);
                            }
                            else
                            {
                                results.Add(false);
                            }
                        }
                    }
                    else
                    {
                        // If type not found, return all false
                        for (int i = 0; i < taskCount; i++)
                        {
                            results.Add(false);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Source execution error: {ex.Message}");
                MessageBox.Show($"採点中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                // Return all false on error
                for (int i = 0; i < taskCount; i++)
                {
                    results.Add(false);
                }
            }
            
            return results;
        }
        
        private string GetGroupFolder(string libraryName)
        {
            if (libraryName.StartsWith("ExcelChecker1_")) return "Group1";
            if (libraryName.StartsWith("ExcelChecker2_")) return "Group2";
            if (libraryName.StartsWith("ExcelChecker3_")) return "Group3";
            return "Group1";
        }
        
        private string GetCheckTaskMethodName(string libraryName, int taskNumber)
        {
            // Extract group and project numbers from library name
            var parts = libraryName.Replace("ExcelChecker", "").Split('_');
            if (parts.Length == 2)
            {
                string groupId = parts[0];
                string projectId = parts[1];
                return $"CheckTask_{groupId}_{projectId}_{taskNumber:D2}";
            }
            return $"CheckTask_{taskNumber:D2}";
        }
        
        private List<bool> ExecuteScoringDirect(string libraryName, int taskCount)
        {
            var results = new List<bool>();
            
            try
            {
                Console.WriteLine($"[DEBUG] ExecuteScoringDirect called with libraryName: {libraryName}, taskCount: {taskCount}");
                
                // Extract group and project numbers from library name
                var parts = libraryName.Replace("ExcelChecker", "").Split('_');
                if (parts.Length >= 2)
                {
                    string groupId = parts[0];
                    string projectId = parts[1];
                    
                    // Create namespace and type name
                    string namespaceName = $"Libraries.Group{groupId}";
                    string fullTypeName = $"{namespaceName}.{libraryName}";
                    
                    Console.WriteLine($"[DEBUG] Looking for type: {fullTypeName}");
                    
                    // すべてのアセンブリから型を検索（改善版）
                    Type checkerType = null;
                    
                    // 最初にMOSExcelMogiAppアセンブリを優先的にチェック
                    var mainAssembly = AppDomain.CurrentDomain.GetAssemblies()
                        .FirstOrDefault(a => a.GetName().Name == "MOSExcelMogiApp");
                    
                    if (mainAssembly != null)
                    {
                        Console.WriteLine($"[DEBUG] Checking main assembly: MOSExcelMogiApp");
                        
                        try
                        {
                            checkerType = mainAssembly.GetType(fullTypeName);
                            if (checkerType != null)
                            {
                                Console.WriteLine($"[DEBUG] Type found in main assembly!");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[DEBUG] Error getting type from main assembly: {ex.Message}");
                        }
                        
                        // 見つからない場合は、名前で検索
                        if (checkerType == null)
                        {
                            try
                            {
                                var types = mainAssembly.GetTypes();
                                Console.WriteLine($"[DEBUG] Main assembly has {types.Length} types");
                                
                                foreach (var type in types)
                                {
                                    if (type.Name == libraryName)
                                    {
                                        Console.WriteLine($"[DEBUG] Found type by name: {type.FullName}");
                                        checkerType = type;
                                        break;
                                    }
                                }
                            }
                            catch (ReflectionTypeLoadException ex)
                            {
                                Console.WriteLine($"[DEBUG] ReflectionTypeLoadException: {ex.Message}");
                                // LoaderExceptionsを確認
                                if (ex.LoaderExceptions != null)
                                {
                                    foreach (var loaderEx in ex.LoaderExceptions)
                                    {
                                        Console.WriteLine($"[DEBUG] Loader exception: {loaderEx?.Message}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[DEBUG] Error loading types: {ex.Message}");
                            }
                        }
                    }
                    
                    // まだ見つからない場合は、他のアセンブリもチェック
                    if (checkerType == null)
                    {
                        Console.WriteLine($"[DEBUG] Type not found in main assembly, checking all assemblies...");
                        
                        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                        {
                            if (assembly == mainAssembly) continue; // 既にチェック済み
                            
                            try
                            {
                                Console.WriteLine($"[DEBUG] Checking assembly: {assembly.GetName().Name}");
                                
                                checkerType = assembly.GetType(fullTypeName);
                                if (checkerType != null)
                                {
                                    Console.WriteLine($"[DEBUG] Type found in assembly: {assembly.GetName().Name}");
                                    break;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[DEBUG] Error checking assembly {assembly.GetName().Name}: {ex.Message}");
                            }
                        }
                    }
                    
                    // If type not found, try to load from DLL
                    if (checkerType == null)
                    {
                        Console.WriteLine($"[DEBUG] Type not found in loaded assemblies, trying to load DLL");
                        
                        // Try multiple DLL paths in order of preference
                        string[] dllPaths = {
                            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bin", "Debug", "Libraries", $"Group{groupId}", $"{libraryName}.dll"),
                            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", $"Group{groupId}", $"{libraryName}.dll"),
                            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Libraries", "bin", "Debug", "net48", $"{libraryName}.dll")
                        };
                        
                        string dllPath = null;
                        foreach (string path in dllPaths)
                        {
                            Console.WriteLine($"[DEBUG] Trying DLL path: {path}");
                            if (File.Exists(path))
                            {
                                dllPath = path;
                                Console.WriteLine($"[DEBUG] Found DLL at: {dllPath}");
                                break;
                            }
                        }
                        
                        if (dllPath != null)
                        {
                            try
                            {
                                Console.WriteLine($"[DEBUG] Loading DLL from: {dllPath}");
                                Assembly assembly = Assembly.LoadFrom(dllPath);
                                checkerType = assembly.GetTypes().FirstOrDefault(t => t.Name == libraryName);
                                Console.WriteLine($"[DEBUG] Found type from DLL: {checkerType?.Name ?? "null"}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[DEBUG] Error loading DLL: {ex.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"[DEBUG] DLL not found in any of the searched paths");
                        }
                    }
                    
                    if (checkerType != null)
                    {
                        Console.WriteLine($"[DEBUG] Type found: {checkerType.Name}");
                        
                        // デバッグ: 利用可能なメソッドをすべて表示
                        Console.WriteLine($"[DEBUG] Available methods in {checkerType.Name}:");
                        foreach (var method in checkerType.GetMethods())
                        {
                            if (method.Name.StartsWith("CheckTask"))
                            {
                                Console.WriteLine($"[DEBUG] - {method.Name}");
                            }
                        }
                        
                        object checkerInstance = Activator.CreateInstance(checkerType);
                        
                        for (int i = 1; i <= taskCount; i++)
                        {
                            // 正しいメソッド名の形式を試す
                            // Group1の場合は CheckTask_1_{projectId}_{taskNumber:D2} の形式を優先
                            string[] methodNames;
                            if (groupId == "1")
                            {
                                methodNames = new string[] {
                                    $"CheckTask_1_{projectId}_{i:D2}",          // CheckTask_1_1_01 (優先)
                                    $"CheckTask_1_{projectId}_0{i}",            // CheckTask_1_1_01 (優先)
                                    $"CheckTask_{groupId}_{projectId}_{i:D2}",  // CheckTask_1_1_01 (fallback)
                                    $"CheckTask_{groupId}_{projectId}_0{i}"     // CheckTask_1_1_01 (fallback)
                                };
                            }
                            else
                            {
                                methodNames = new string[] {
                                    $"CheckTask_{groupId}_{projectId}_{i:D2}",  // CheckTask_2_1_01
                                    $"CheckTask_{groupId}_{projectId}_0{i}",    // CheckTask_2_1_01
                                    $"CheckTask_1_{projectId}_{i:D2}",          // CheckTask_1_1_01 (fallback)
                                    $"CheckTask_1_{projectId}_0{i}"             // CheckTask_1_1_01 (fallback)
                                };
                            }
                            
                            bool methodFound = false;
                            foreach (string methodName in methodNames)
                            {
                                Console.WriteLine($"[DEBUG] Looking for method: {methodName}");
                                MethodInfo method = checkerType.GetMethod(methodName);
                                
                                if (method != null)
                                {
                                    Console.WriteLine($"[DEBUG] Method found: {methodName}");
                                    bool result = (bool)method.Invoke(checkerInstance, null);
                                    Console.WriteLine($"[DEBUG] Method {methodName} result: {result}");
                                    results.Add(result);
                                    methodFound = true;
                                    break;
                                }
                                else
                                {
                                    Console.WriteLine($"[DEBUG] Method not found: {methodName}");
                                }
                            }
                            
                            if (!methodFound)
                            {
                                Console.WriteLine($"[DEBUG] No method found for task {i}, returning false");
                                results.Add(false);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[DEBUG] Type not found: {fullTypeName}");
                        // Fill with false results if type not found
                        for (int i = 0; i < taskCount; i++)
                        {
                            results.Add(false);
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"[DEBUG] Invalid library name format: {libraryName}");
                    // Fill with false results if invalid format
                    for (int i = 0; i < taskCount; i++)
                    {
                        results.Add(false);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Direct execution error: {ex.Message}");
                Console.WriteLine($"[DEBUG] Stack trace: {ex.StackTrace}");
                System.Diagnostics.Debug.WriteLine($"Direct execution error: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                MessageBox.Show($"採点中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                // Fill with false results in case of error
                for (int i = 0; i < taskCount; i++)
                {
                    results.Add(false);
                }
            }
            
            return results;
        }

        private void ShowAppBar()
        {
            // メインウィンドウを非表示にしてアプリバーを表示
            HideMainWindowRequested?.Invoke(this, EventArgs.Empty);
            ShowAppBarRequested?.Invoke(this, EventArgs.Empty);
        }
        
        private void CloseExcelApplication()
        {
            ExcelApp excelApp = null;
            ExcelWorkbook activeWorkbook = null;
            int excelPid = -1;

            try
            {
                // Excel COMオブジェクトを使用して保存してから閉じる
                try
                {
                    // アプリが保持しているインスタンスを優先して閉じる
                    excelApp = _sharedExcelApp ?? (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                    if (excelApp != null)
                    {
                        excelPid = Libraries.ExcelApplicationManager.TryGetExcelProcessId(excelApp);
                        activeWorkbook = excelApp.ActiveWorkbook;
                        if (activeWorkbook != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"[CloseExcelApplication] Saving workbook: {activeWorkbook.Name}");
                            // ワークブックを保存
                            activeWorkbook.Save();
                            System.Diagnostics.Debug.WriteLine("[CloseExcelApplication] Workbook saved successfully");
                            
                            // ワークブックを閉じる
                            activeWorkbook.Close(SaveChanges: false);
                            Marshal.ReleaseComObject(activeWorkbook);
                            activeWorkbook = null;
                        }
                        // 結果へ戻る経路と同様に、Excel インスタンス自体を終了する。
                        try { excelApp.Quit(); } catch { }
                        Marshal.ReleaseComObject(excelApp);
                        excelApp = null;
                    }
                }
                catch (COMException)
                {
                    // Excelが開いていない場合は無視
                    System.Diagnostics.Debug.WriteLine("[CloseExcelApplication] No Excel application is running");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[CloseExcelApplication] Error saving/closing Excel: {ex.Message}");
                }

                if (excelPid > 0)
                {
                    Libraries.ExcelApplicationManager.EnsureExcelProcessExited(
                        excelPid,
                        8000,
                        3000,
                        "[CloseExcelApplication]");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Excel終了エラー: {ex.Message}");
            }
            finally
            {
                // リソースのクリーンアップ
                if (activeWorkbook != null)
                {
                    try { Marshal.ReleaseComObject(activeWorkbook); } catch { }
                }
                if (excelApp != null)
                {
                    try { Marshal.ReleaseComObject(excelApp); } catch { }
                }
                _sharedExcelApp = null;
            }
        }

        /// <summary>
        /// プロジェクトリセット前に全ワークブックを保存せずに閉じ、Excel を終了し、共有 COM 参照をクリアする。
        /// 読み取り専用二重オープンやファイルロック残りを防ぐため終了ボタン経路に近いクリーンアップを行う。
        /// </summary>
        public void QuitExcelForProjectReset()
        {
            ExcelApp excelApp = null;
            try
            {
                try
                {
                    excelApp = _sharedExcelApp;
                    if (excelApp != null)
                    {
                        try
                        {
                            _ = excelApp.Visible;
                        }
                        catch
                        {
                            excelApp = null;
                            _sharedExcelApp = null;
                        }
                    }

                    if (excelApp == null)
                    {
                        excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                    }
                }
                catch (COMException)
                {
                    System.Diagnostics.Debug.WriteLine("[QuitExcelForProjectReset] No Excel application");
                    _sharedExcelApp = null;
                    return;
                }

                // Quit 後もプロセスが残ると VSTO が再ロードされずログタブが消えるため、PID を記録して確実に終了させる。
                int excelPid = Libraries.ExcelApplicationManager.TryGetExcelProcessId(excelApp);

                bool originalDisplayAlerts = true;
                try
                {
                    originalDisplayAlerts = excelApp.DisplayAlerts;
                }
                catch
                {
                    /* ignore */
                }

                try
                {
                    excelApp.DisplayAlerts = false;

                    for (int pass = 0; pass < 5 && excelApp.Workbooks.Count > 0; pass++)
                    {
                        int remaining = excelApp.Workbooks.Count;
                        while (excelApp.Workbooks.Count > 0)
                        {
                            ExcelWorkbook wb = null;
                            try
                            {
                                wb = excelApp.Workbooks[1];
                                wb.Close(SaveChanges: false);
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[QuitExcelForProjectReset] Close workbook: {ex.Message}");
                                break;
                            }
                            finally
                            {
                                if (wb != null)
                                {
                                    try
                                    {
                                        Marshal.ReleaseComObject(wb);
                                    }
                                    catch
                                    {
                                        /* ignore */
                                    }
                                }
                            }
                        }

                        if (excelApp.Workbooks.Count > 0 && excelApp.Workbooks.Count == remaining)
                        {
                            Thread.Sleep(300);
                        }
                    }
                }
                finally
                {
                    try
                    {
                        excelApp.DisplayAlerts = originalDisplayAlerts;
                    }
                    catch
                    {
                        /* ignore */
                    }
                }

                try
                {
                    excelApp.Quit();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[QuitExcelForProjectReset] Quit: {ex.Message}");
                }

                _sharedExcelApp = null;
                try
                {
                    Marshal.ReleaseComObject(excelApp);
                }
                catch
                {
                    /* ignore */
                }

                excelApp = null;

                const int quitWaitMs = 10000;
                if (excelPid > 0)
                {
                    Libraries.ExcelApplicationManager.EnsureExcelProcessExited(
                        excelPid,
                        quitWaitMs,
                        5000,
                        "[QuitExcelForProjectReset]");
                }
                else
                {
                    Libraries.ExcelApplicationManager.WaitForAllExcelProcessesGone(quitWaitMs);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[QuitExcelForProjectReset] {ex.Message}");
                _sharedExcelApp = null;
            }
        }

        private void ExecutePauseExam(object parameter) 
        {
        }

        private void ExecuteResetExam(object parameter)
        {
        }
        
        /// <summary>
        /// 現在のExcelプロジェクトを自動保存する共通メソッド
        /// </summary>
        /// <param name="closeWorkbook">保存後にワークブックを閉じるかどうか</param>
        public void SaveCurrentExcelProject(bool closeWorkbook = true)
        {
            if (CurrentProject == null)
            {
                return;
            }

            int groupId = int.Parse(CurrentProject.Group.Replace("Group ", ""));
            int currentProjectNumber = CurrentProject.ProjectNumber;

            System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Saving current project (Group{groupId}, Project{currentProjectNumber})");
            
            // COM Interopを使用して現在のExcelファイルを上書き保存
            ExcelApp excelApp = null;
            ExcelWorkbook workbook = null;
            
            try
            {
                // Excelアプリケーションを取得
                try
                {
                    excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                    System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Got existing Excel application");
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Excel application not found, skipping save");
                    excelApp = null;
                }
                
                if (excelApp != null)
                {
                    // 現在開いているワークブックを検索
                    workbook = null;
                    string currentFileName = System.IO.Path.GetFileName(CurrentProject.FilePath);
                    
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        if (wb.FullName.Equals(CurrentProject.FilePath, StringComparison.OrdinalIgnoreCase) ||
                            wb.Name.Equals(currentFileName, StringComparison.OrdinalIgnoreCase))
                        {
                            workbook = wb;
                            System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Found current workbook: {workbook.Name}");
                            break;
                        }
                    }
                    
                    // ワークブックが見つかった場合は保存
                    if (workbook != null)
                    {
                        // 現在開いているファイルのパスを取得
                        string currentFilePath = workbook.FullName;
                        if (string.IsNullOrEmpty(currentFilePath))
                        {
                            currentFilePath = CurrentProject.FilePath;
                        }
                        
                        System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Current file path: {currentFilePath}");
                        
                        // 既存ファイルの読み取り専用属性を解除
                        if (File.Exists(currentFilePath))
                        {
                            FileInfo fileInfo = new FileInfo(currentFilePath);
                            if (fileInfo.IsReadOnly)
                            {
                                fileInfo.IsReadOnly = false;
                                System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Removed read-only attribute from existing file");
                            }
                        }
                        
                        // ワークブックを上書き保存
                        // 既存ファイルの上書き確認ダイアログを自動で「はい」にするため、DisplayAlertsを無効化
                        bool originalDisplayAlerts = excelApp.DisplayAlerts;
                        try
                        {
                            excelApp.DisplayAlerts = false;
                            System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Disabled Excel display alerts for automatic overwrite");
                            
                            // 現在開いているファイルを上書き保存
                            workbook.Save();
                            
                            System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Saved current project to: {currentFilePath}");
                        }
                        finally
                        {
                            // DisplayAlertsを元の状態に戻す
                            excelApp.DisplayAlerts = originalDisplayAlerts;
                            System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Restored Excel display alerts to original state");
                        }
                        
                        // 保存したファイルの読み取り専用属性を解除
                        FileInfo savedFileInfo = new FileInfo(currentFilePath);
                        if (savedFileInfo.IsReadOnly)
                        {
                            savedFileInfo.IsReadOnly = false;
                            System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Removed read-only attribute from saved file");
                        }
                        
                        // ワークブックを閉じる（closeWorkbookがtrueの場合のみ）
                        if (closeWorkbook)
                        {
                            workbook.Close(SaveChanges: false);
                            workbook = null;
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Current workbook not found, skipping save");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SaveCurrentExcelProject] Error saving current project: {ex.Message}\n{ex.StackTrace}");
            }
            finally
            {
                // COMオブジェクトの解放
                if (workbook != null && closeWorkbook)
                {
                    try
                    {
                        Marshal.ReleaseComObject(workbook);
                    }
                    catch { }
                }
            }
        }

        private void ExecuteNextProject(object parameter)
        {
            if (CurrentProject == null)
            {
                return;
            }

            // オブジェクト選択時は警告を表示して移動しない
            if (TryShowObjectSelectedWarningIfExcelObjectSelected())
                return;

            int groupId = int.Parse(CurrentProject.Group.Replace("Group ", ""));
            int currentProjectNumber = CurrentProject.ProjectNumber;

            // プロジェクト10の場合はレビューページを開く
            if (currentProjectNumber == 10)
            {
                System.Diagnostics.Debug.WriteLine($"[ExecuteNextProject] Opening review page for Project 10");
                OpenReviewPageRequested?.Invoke(this, EventArgs.Empty);
                return;
            }

            // プロジェクト10以降は何もしない
            if (currentProjectNumber >= 10)
            {
                return;
            }

            // 現在のプロジェクトを保存
            SaveCurrentExcelProject(closeWorkbook: true);
            
            // 次のプロジェクトを取得
            int nextProjectNumber = CurrentProject.ProjectNumber + 1;
            
            // 次のプロジェクトのファイルパスを取得（Initialフォルダから）
            string nextFilePath = GetProjectFilePath(groupId, nextProjectNumber);
            
            if (string.IsNullOrEmpty(nextFilePath) || !File.Exists(nextFilePath))
            {
                ResultMessage = $"エラー: 次のプロジェクトファイルが見つかりません: {nextFilePath}";
                return;
            }
            
            try
            {
                // 次のプロジェクトのExcelファイルを開く（Excelプロセスは使い回し、ブックだけ切替）
                ExcelApp excelApp = null;
                ExcelWorkbook targetWorkbook = null;
                try
                {
                    excelApp = GetOrCreateExcelApplication();

                    string targetFullPathLower = Path.GetFullPath(nextFilePath).ToLowerInvariant();

                    // 既に開いているか確認
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            string wbFullPathLower = Path.GetFullPath(wb.FullName).ToLowerInvariant();
                            if (wbFullPathLower == targetFullPathLower)
                            {
                                targetWorkbook = wb;
                                break;
                            }
                        }
                        catch { }
                    }

                    if (targetWorkbook == null)
                    {
                        targetWorkbook = excelApp.Workbooks.Open(nextFilePath, ReadOnly: false);
                    }

                    try { targetWorkbook.Activate(); } catch { }
                    try { excelApp.Visible = true; } catch { }
                }
                finally
                {
                    // 共有インスタンスを使い回すため Release しない
                }
                // プロジェクト情報を更新
                CurrentProject = new ProjectInfo
                {
                    Name = $"プロジェクト{groupId}-{nextProjectNumber}",
                    FilePath = nextFilePath,
                    Group = $"Group {groupId}",
                    ProjectNumber = nextProjectNumber
                };
                OnPropertyChanged(nameof(IsNextProjectVisible));
                ResultMessage = $"次のプロジェクトに移動しました: {Path.GetFileName(nextFilePath)}";
            }
            catch (Exception ex)
            {
                ResultMessage = $"エラー: 次のプロジェクトファイルを開けませんでした: {ex.Message}";
            }
        }

        /// <summary>
        /// Excelでグラフが選択されている場合にのみ警告ダイアログを表示する。
        /// セル・テーブル・その他を選択しているときは警告しない。
        /// 警告を表示した場合は true を返す。
        /// </summary>
        private bool TryShowObjectSelectedWarningIfExcelObjectSelected()
        {
            try
            {
                ExcelApp excelApp = null;
                try
                {
                    excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    return false;
                }
                if (excelApp?.ActiveWorkbook?.ActiveSheet == null)
                    return false;

                object selection = null;
                try
                {
                    selection = excelApp.Selection;
                }
                catch
                {
                    return false;
                }
                if (selection == null) return false;

                string typeName = selection.GetType().Name;
                // グラフ選択時のみ警告する（ChartObject または Chart）
                bool isChartSelected = typeName == "ChartObject" || typeName == "Chart";

                if (!isChartSelected) return false;

                Application.Current?.Dispatcher.Invoke(() =>
                {
                    var w = new ObjectSelectedWarningWindow();
                    w.ShowDialog();
                });
                return true;
            }
            catch
            {
                return false;
            }
        }
        
        private void ExecuteEndExam(object parameter)
        {
            if (Interlocked.CompareExchange(ref _endExamShutdownStarted, 1, 0) != 0)
            {
                System.Diagnostics.Debug.WriteLine("[ExecuteEndExam] duplicate call ignored");
                return;
            }

            IsExcelOverlayVisible = false;
            CurrentProject = null;
            ResultMessage = "試験を終了しました。";

            _excelShutdownFinished.Reset();

            // Office COM は STA 上で扱う（スレッドプール MTA の Task.Run は不安定になり得る）
            var shutdownThread = new Thread(() =>
            {
                try
                {
                    SaveAllExcelWorkbooks();
                    CloseExcelApplication();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ExecuteEndExam] Excel shutdown: {ex.Message}");
                }
                finally
                {
                    _excelShutdownFinished.Set();
                }
            })
            {
                IsBackground = true
            };
            shutdownThread.SetApartmentState(ApartmentState.STA);
            shutdownThread.Start();

            // ExamEndedイベントを発火してアプリバーを閉じ、メインウィンドウを再表示
            ExamEnded?.Invoke(this, EventArgs.Empty);
            ShowMainWindowRequested?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// 前試験の Excel 終了スレッドが完了するまで待つ。完了しない場合は ResultMessage を設定して false。
        /// </summary>
        private bool WaitForExcelShutdownToCompleteBeforeOpeningProject()
        {
            if (_excelShutdownFinished.Wait(0))
                return true;

            System.Diagnostics.Debug.WriteLine("[MainViewModel] 前試験の Excel 終了処理の完了を待機しています…");
            const int maxWaitMs = 12_000;
            const int overlayDelayMs = 400;
            var disp = Application.Current?.Dispatcher;
            var sw = Stopwatch.StartNew();
            var overlayShown = false;
            try
            {
                while (sw.ElapsedMilliseconds < maxWaitMs)
                {
                    if (_excelShutdownFinished.Wait(50))
                        return true;

                    if (!overlayShown && sw.ElapsedMilliseconds >= overlayDelayMs)
                    {
                        IsShutdownWaitOverlayVisible = true;
                        overlayShown = true;
                        disp?.Invoke(() => { }, DispatcherPriority.Loaded);
                    }
                    else
                    {
                        disp?.Invoke(() => { }, DispatcherPriority.Background);
                    }
                }

                ResultMessage =
                    "終了処理の完了に時間がかかっています。少し待ってから、もう一度プロジェクトを開いてください。";
                System.Diagnostics.Debug.WriteLine("[MainViewModel] Excel 終了待機がタイムアウトしました");
                return false;
            }
            finally
            {
                IsShutdownWaitOverlayVisible = false;
            }
        }

        /// <summary>
        /// 保存用に Excel.Application を取得。共有参照を優先し、失敗時は ROT 登録まで短時間リトライする。
        /// </summary>
        private static bool TryAcquireExcelApplicationForSave(ExcelApp sharedRef, int maxWaitMs, out ExcelApp app, out bool releaseComObjectWhenDone)
        {
            app = null;
            releaseComObjectWhenDone = false;

            if (sharedRef != null)
            {
                try
                {
                    _ = sharedRef.Visible;
                    app = sharedRef;
                    releaseComObjectWhenDone = false;
                    return true;
                }
                catch
                {
                    /* GetActiveObject へ */
                }
            }

            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < maxWaitMs)
            {
                try
                {
                    app = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                    releaseComObjectWhenDone = true;
                    return true;
                }
                catch
                {
                    Thread.Sleep(100);
                }
            }

            return false;
        }

        /// <summary>
        /// 開いているすべてのExcelワークブックを保存する（閉じない）。
        /// </summary>
        private void SaveAllExcelWorkbooks()
        {
            const int acquireWaitMs = 3000;
            if (!TryAcquireExcelApplicationForSave(_sharedExcelApp, acquireWaitMs, out var excelApp, out var releaseApp))
            {
                System.Diagnostics.Debug.WriteLine(
                    "[SaveAllExcelWorkbooks] Excel を取得できませんでした（既に終了済み、または ROT 未登録の可能性）");
                return;
            }

            try
            {
                foreach (ExcelWorkbook wb in excelApp.Workbooks)
                {
                    try
                    {
                        wb.Save();
                        System.Diagnostics.Debug.WriteLine($"[SaveAllExcelWorkbooks] 保存しました: {wb.Name}");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[SaveAllExcelWorkbooks] 保存エラー {wb?.Name}: {ex.Message}");
                    }
                }
            }
            finally
            {
                if (releaseApp && excelApp != null)
                {
                    try { Marshal.ReleaseComObject(excelApp); } catch { }
                }
            }
        }
        
        private void ExecuteUiTest(object parameter)
        {
            // UIテスト用のアプリバーウィンドウを表示
            var uiTestAppBar = new MOSExcelMogiApp.Views.UiTestAppBarWindow();
            uiTestAppBar.Show();
            
            // アプリバーの実高さを取得してから Excel を配置
            uiTestAppBar.ContentRendered += (s, e) =>
            {
                int barHeight = (int)Math.Round(uiTestAppBar.ActualHeight > 0 ? uiTestAppBar.ActualHeight : uiTestAppBar.Height);
                LaunchAndPositionExcel(barHeight);
            };

            UiTestRequested?.Invoke(this, EventArgs.Empty);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
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
}