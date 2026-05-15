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
using System.Windows.Input;
using Core.Ports.Primary;
using MOSExcelMogiApp.Views;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;

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
        private readonly IExcelCheckerService _excelCheckerService;
        private int _selectedTabIndex;
        private string _selectedFilePath;
        private string _resultMessage;
        private bool _isExcelOverlayVisible;
        private ProjectInfo _currentProject;
        
        public event EventHandler ExamEnded;
        public event EventHandler ShowAppBarRequested;
        public event EventHandler HideMainWindowRequested;
        public event EventHandler ShowMainWindowRequested;
        public event EventHandler UiTestRequested;

        public MainViewModel(IExcelCheckerService excelCheckerService)
        {
            _excelCheckerService = excelCheckerService;
            LoadProjects();
            CheckCommand = new RelayCommand(ExecuteCheck);
            OpenProjectCommand = new RelayCommand(ExecuteOpenProject);
            ScoreCommand = new RelayCommand(ExecuteScore);
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

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const int SW_RESTORE = 9;

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

            ShowWindow(hWnd, SW_RESTORE);
            SetForegroundWindow(hWnd);
            bool ok = SetWindowPos(hWnd, IntPtr.Zero, 0, 0, widthPx, heightPx, SWP_NOZORDER | SWP_SHOWWINDOW);
            if (!ok)
            {
                MoveWindow(hWnd, 0, 0, widthPx, heightPx, true);
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
            }
        }
        
        public string CurrentProjectName => CurrentProject?.Name ?? "プロジェクトが選択されていません";
        
        public bool IsNextProjectVisible => CurrentProject != null && CurrentProject.ProjectNumber < 10;

        private void LoadProjects()
        {
            var allProjects = _excelCheckerService.GetAllProjects();
            
            for (int groupId = 1; groupId <= 3; groupId++)
            {
                var group = new ProjectGroupViewModel { GroupId = groupId, GroupName = $"Group {groupId}" };
                
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
            if (parameter is ProjectViewModel)
            {
                var project = (ProjectViewModel)parameter;
                if (!string.IsNullOrEmpty(SelectedFilePath))
                {
                    bool result = _excelCheckerService.CheckExcel(project.GroupId, project.ProjectId, SelectedFilePath);
                    ResultMessage = $"Project{project.GroupId}-{project.ProjectId}: {(result ? "Success" : "Failed")}";
                }
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
            if (parameter is ProjectViewModel)
            {
                var projectViewModel = (ProjectViewModel)parameter;
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

            string filePath = GetProjectFilePath(projectId);
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                ResultMessage = $"エラー: ファイルが見つかりません: {filePath}";
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
                
                // プロジェクト情報を設定
                if (parameter is ProjectViewModel)
                {
                    var pvm = (ProjectViewModel)parameter;
                    // Use ProjectViewModel data directly
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
                    // Parse from string projectId
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
                
                // メインウィンドウを非表示にしてアプリバーを表示
                ShowAppBar();
            }
            catch (Exception ex)
            {
                ResultMessage = $"エラー: ファイルを開けませんでした: {ex.Message}";
            }
        }

        private bool IsExcelRunning()
        {
            Process[] excelProcesses = Process.GetProcessesByName("EXCEL");
            return excelProcesses.Length > 0;
        }

        private string GetProjectFilePath(int groupId, int projectId)
        {
            return $"C:\\MOSTest\\Excel365\\Tab{groupId}\\project{projectId}.xlsx";
        }

        private string GetProjectFilePath(string projectId)
        {
            // Parse projectId like "tab1-project1-1" or "project1-2"
            System.Diagnostics.Debug.WriteLine($"GetProjectFilePath called with projectId: {projectId}");
            var parts = projectId.Split('-');
            
            if (parts.Length >= 2)
            {
                // Handle format like "project1-2" (groupId-projectId)
                if (parts[0].StartsWith("project"))
                {
                    string groupPart = parts[0].Replace("project", "");
                    string projectPart = parts[1];
                    
                    System.Diagnostics.Debug.WriteLine($"Parsing project format: groupPart={groupPart}, projectPart={projectPart}");
                    
                    if (int.TryParse(groupPart, out int groupId) && int.TryParse(projectPart, out int projId))
                    {
                        string filePath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\project{projId}.xlsx";
                        System.Diagnostics.Debug.WriteLine($"Generated file path: {filePath}");
                        return filePath;
                    }
                }
                // Handle format like "tab1-project1-1"
                else if (parts[0].StartsWith("tab"))
                {
                    string tabPart = parts[0].Replace("tab", "");
                    string projectPart = parts[1].Replace("project", "");
                    
                    System.Diagnostics.Debug.WriteLine($"Parsing tab format: tabPart={tabPart}, projectPart={projectPart}");
                    
                    if (int.TryParse(tabPart, out int groupId) && int.TryParse(projectPart, out int projId))
                    {
                        string filePath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\project{projId}.xlsx";
                        System.Diagnostics.Debug.WriteLine($"Generated file path: {filePath}");
                        return filePath;
                    }
                }
            }
            System.Diagnostics.Debug.WriteLine($"Failed to parse projectId: {projectId}");
            return string.Empty;
        }

        private void ExecuteScore(object parameter)
        {
            if (CurrentProject != null)
            {
                try
                {
                    // Extract group and project IDs from the project info
                    int groupId = int.Parse(CurrentProject.Group.Replace("Group ", ""));
                    int projectId = CurrentProject.ProjectNumber;
                    
                    // Load config.json to get task count
                    var config = LoadConfig();
                    var projectConfig = GetProjectConfig(config, groupId, projectId);
                    
                    if (projectConfig != null)
                    {
                        int taskCount = projectConfig["taskCount"].Value<int>();
                        
                        // Execute scoring using direct method calls
                        string libraryName = $"ExcelChecker{groupId}_{projectId}";
                        var results = ExecuteScoringDirect(libraryName, taskCount);
                        var dialog = new ScoringResultDialog(taskCount, results);
                        dialog.ShowDialog();
                        
                        ResultMessage = $"採点完了: {taskCount}問のタスクを採点しました";
                    }
                    else
                    {
                        MessageBox.Show("プロジェクト設定が見つかりませんでした。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    ResultMessage = $"エラー: {ex.Message}";
                    System.Diagnostics.Debug.WriteLine($"Error: {ex}");
                    MessageBox.Show($"採点中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("プロジェクトが選択されていません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
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
                    
                    // Try to get the type from loaded assemblies
                    Type checkerType = Type.GetType(fullTypeName);
                    if (checkerType == null)
                    {
                        foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
                        {
                            checkerType = assembly.GetType(fullTypeName);
                            if (checkerType != null) break;
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
                            string[] methodNames = {
                                $"CheckTask_{groupId}_{projectId}_{i:D2}",  // CheckTask_1_2_01
                                $"CheckTask_{groupId}_{projectId}_0{i}",    // CheckTask_1_2_01
                                $"CheckTask_1_{projectId}_{i:D2}",          // CheckTask_1_2_01 (fallback)
                                $"CheckTask_1_{projectId}_0{i}"             // CheckTask_1_2_01 (fallback)
                            };
                            
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
                System.Diagnostics.Debug.WriteLine($"Direct execution error: {ex.Message}");
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
            try
            {
                // Excelプロセスを取得して終了
                var excelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (var process in excelProcesses)
                {
                    process.CloseMainWindow();
                    if (!process.WaitForExit(5000)) // 5秒待機
                    {
                        process.Kill(); // 強制終了
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Excel終了エラー: {ex.Message}");
            }
        }

        private void ExecutePauseExam(object parameter) 
        {
        }

        private void ExecuteResetExam(object parameter)
        {
        }
        
        private void ExecuteNextProject(object parameter)
        {
            if (CurrentProject == null || CurrentProject.ProjectNumber >= 10)
            {
                return;
            }
            
            // Excelアプリケーションを閉じる
            CloseExcelApplication();
            
            // 次のプロジェクトを取得
            int nextProjectNumber = CurrentProject.ProjectNumber + 1;
            int groupId = int.Parse(CurrentProject.Group.Replace("Group ", ""));
            
            // 次のプロジェクトのファイルパスを取得
            string nextFilePath = GetProjectFilePath(groupId, nextProjectNumber);
            
            if (string.IsNullOrEmpty(nextFilePath) || !File.Exists(nextFilePath))
            {
                ResultMessage = $"エラー: 次のプロジェクトファイルが見つかりません: {nextFilePath}";
                return;
            }
            
            try
            {
                // 次のプロジェクトのExcelファイルを開く
                Process.Start(new ProcessStartInfo
                {
                    FileName = nextFilePath,
                    UseShellExecute = true
                });
                
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
        
        private void ExecuteEndExam(object parameter)
        {
            IsExcelOverlayVisible = false;
            CurrentProject = null;
            ResultMessage = "試験を終了しました。";
            
            // Excelアプリケーションを閉じる
            CloseExcelApplication();
            
            // ExamEndedイベントを発火してアプリバーを閉じ、メインウィンドウを再表示
            ExamEnded?.Invoke(this, EventArgs.Empty);
            ShowMainWindowRequested?.Invoke(this, EventArgs.Empty);
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