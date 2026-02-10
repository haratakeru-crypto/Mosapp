using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Newtonsoft.Json.Linq;
using MOS_PowerPoint_app;

namespace MOS_PowerPoint_app.Views
{
    /// <summary>
    /// 科目選択（表紙）画面。パスワード認証後に表示され、Excel / Word / PowerPoint のいずれかを選択する。
    /// </summary>
    public partial class SubjectSelectionWindow : Window
    {
        private MainWindow _powerPointMainWindow;

        public SubjectSelectionWindow()
        {
            InitializeComponent();
        }

        private void ExcelButton_Click(object sender, RoutedEventArgs e)
        {
            string excelExePath = GetExcelAppExePath();
            if (string.IsNullOrEmpty(excelExePath) || !File.Exists(excelExePath))
            {
                MessageBox.Show("Excel アプリが見つかりません。\n" + (excelExePath ?? ""), "科目選択", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string excelInitialPath = GetExcelInitialPath();
            if (!string.IsNullOrEmpty(excelInitialPath) && !Directory.Exists(excelInitialPath))
            {
                excelInitialPath = null;
            }
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = excelExePath,
                    UseShellExecute = false,
                    WorkingDirectory = string.IsNullOrEmpty(excelInitialPath) ? Path.GetDirectoryName(excelExePath) : excelInitialPath
                };
                var process = Process.Start(startInfo);
                if (process != null)
                {
                    process.EnableRaisingEvents = true;
                    process.Exited += OnLaunchedProcessExited;
                }
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excel アプリの起動に失敗しました: {ex.Message}", "科目選択", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Excel の初期フォルダの絶対パスを取得する（Assets\config.json の excelInitialPath。未設定時は C:\MOSTest\Excel365\Tab1\Initial）。
        /// </summary>
        private static string GetExcelInitialPath()
        {
            const string defaultPath = @"C:\MOSTest\Excel365\Tab1\Initial";
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (!File.Exists(configPath)) return Path.GetFullPath(defaultPath);
                string json = File.ReadAllText(configPath);
                var config = JObject.Parse(json);
                var path = config["excelInitialPath"]?.ToString();
                if (string.IsNullOrWhiteSpace(path)) return Path.GetFullPath(defaultPath);
                path = path.Trim();
                return Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path));
            }
            catch
            {
                return Path.GetFullPath(defaultPath);
            }
        }

        /// <summary>
        /// mos_xaml_app（MOS Excel 模擬アプリ）の exe 絶対パスを取得する。mos_xaml_app\bin\Release / Debug を最優先で探す。
        /// </summary>
        private static string GetExcelAppExePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var dir = Path.GetDirectoryName(baseDir);
            if (!string.IsNullOrEmpty(dir))
            {
                var projectDir = Path.GetDirectoryName(dir);
                if (!string.IsNullOrEmpty(projectDir))
                {
                    // 最優先: mos_xaml_app\bin\Release
                    string releasePath = Path.GetFullPath(Path.Combine(projectDir, "mos_xaml_app", "bin", "Release", "MOSExcelMogiApp.exe"));
                    if (File.Exists(releasePath)) return releasePath;
                    // 次: mos_xaml_app\bin\Debug
                    string debugPath = Path.GetFullPath(Path.Combine(projectDir, "mos_xaml_app", "bin", "Debug", "MOSExcelMogiApp.exe"));
                    if (File.Exists(debugPath)) return debugPath;
                }
            }
            try
            {
                string configPath = Path.Combine(baseDir, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var config = JObject.Parse(json);
                    var path = config["excelExePath"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        path = path.Trim();
                        string absolutePath = Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(baseDir, path));
                        if (File.Exists(absolutePath)) return absolutePath;
                    }
                }
            }
            catch { }
            string sameDirExe = Path.GetFullPath(Path.Combine(baseDir, "MOSExcelMogiApp.exe"));
            if (File.Exists(sameDirExe)) return sameDirExe;
            if (string.IsNullOrEmpty(dir)) return null;
            var projectDirFallback = Path.GetDirectoryName(dir);
            if (string.IsNullOrEmpty(projectDirFallback)) return null;
            string rootExe = Path.GetFullPath(Path.Combine(projectDirFallback, "mos_xaml_app", "MOSExcelMogiApp.exe"));
            if (File.Exists(rootExe)) return rootExe;
            return rootExe;
        }

        /// <summary>
        /// 表紙の Word ボタン: MOS Word アプリを起動し、プロジェクト一覧を表示する（試験バーは開かず、ユーザーがプロジェクトを選ぶまで待つ）。
        /// </summary>
        private void WordButton_Click(object sender, RoutedEventArgs e)
        {
            string wordExePath = GetWordAppExePath();
            if (string.IsNullOrEmpty(wordExePath) || !File.Exists(wordExePath))
            {
                MessageBox.Show("Word アプリが見つかりません。\n" + (wordExePath ?? ""), "科目選択", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string wordInitialPath = GetWordInitialPath();
            if (!string.IsNullOrEmpty(wordInitialPath) && !Directory.Exists(wordInitialPath))
            {
                MessageBox.Show($"Word の初期フォルダが見つかりません。\n{wordInitialPath}", "科目選択", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = wordExePath,
                    UseShellExecute = false,
                    WorkingDirectory = string.IsNullOrEmpty(wordInitialPath) ? Path.GetDirectoryName(wordExePath) : wordInitialPath
                };
                // #region agent log
                try
                {
                    var logPath = @"c:\Users\kouza\source\repos\MOS PowerPoint app\.cursor\debug.log";
                    var line = "{\"timestamp\":" + DateTimeOffset.UtcNow.ToUnixTimeMilliseconds() + ",\"location\":\"SubjectSelectionWindow.xaml.cs:WordButton_Click\",\"message\":\"Starting Word process\",\"data\":{\"fileName\":\"" + (wordExePath ?? "").Replace("\\", "\\\\") + "\",\"arguments\":\"" + (startInfo.Arguments ?? "").Replace("\\", "\\\\") + "\",\"workingDir\":\"" + (startInfo.WorkingDirectory ?? "").Replace("\\", "\\\\") + "\"},\"sessionId\":\"debug-session\",\"hypothesisId\":\"A\"}\n";
                    File.AppendAllText(logPath, line);
                }
                catch { }
                // #endregion
                var process = Process.Start(startInfo);
                if (process != null)
                {
                    process.EnableRaisingEvents = true;
                    process.Exited += OnLaunchedProcessExited;
                }
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Word アプリの起動に失敗しました: {ex.Message}", "科目選択", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Word の初期フォルダの絶対パスを取得する（Assets\config.json の wordInitialPath。未設定時は C:\MOSTest\Word365\Tab1\Initial）。
        /// </summary>
        private static string GetWordInitialPath()
        {
            const string defaultPath = @"C:\MOSTest\Word365\Tab1\Initial";
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (!File.Exists(configPath)) return Path.GetFullPath(defaultPath);
                string json = File.ReadAllText(configPath);
                var config = JObject.Parse(json);
                var path = config["wordInitialPath"]?.ToString();
                if (string.IsNullOrWhiteSpace(path)) return Path.GetFullPath(defaultPath);
                path = path.Trim();
                return Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path));
            }
            catch
            {
                return Path.GetFullPath(defaultPath);
            }
        }

        /// <summary>
        /// MOS Word アプリの exe 絶対パスを取得する。表紙プロジェクト内の MOS Word app\bin\Release を最優先し、次に config.json の wordExePath、未設定時は同階層の MOS Word app\bin\... を探索。
        /// </summary>
        private static string GetWordAppExePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            // 最優先: 表紙アプリのプロジェクトフォルダ（baseDir の2階層上 = ...\MOS PowerPoint app）内の MOS Word app\bin\Release
            var dir = Path.GetDirectoryName(baseDir);
            if (!string.IsNullOrEmpty(dir))
            {
                var coverProjectRoot = Path.GetDirectoryName(dir);
                if (!string.IsNullOrEmpty(coverProjectRoot))
                {
                    string inProjectRelease = Path.GetFullPath(Path.Combine(coverProjectRoot, "MOS Word app", "bin", "Release", "MOS Word app.exe"));
                    if (File.Exists(inProjectRelease)) return inProjectRelease;
                }
            }
            try
            {
                string configPath = Path.Combine(baseDir, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var config = JObject.Parse(json);
                    var path = config["wordExePath"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        path = path.Trim();
                        string absolutePath = Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(baseDir, path));
                        if (File.Exists(absolutePath)) return absolutePath;
                    }
                }
            }
            catch { }
            // 表紙アプリ（MOS PowerPoint app）と同階層の「MOS Word app」を探す（fallback）
            if (string.IsNullOrEmpty(dir)) return null;
            var binDir = Path.GetDirectoryName(dir);
            if (string.IsNullOrEmpty(binDir)) return null;
            var projectDir = Path.GetDirectoryName(binDir);
            if (string.IsNullOrEmpty(projectDir)) return null;
            string debugPath = Path.GetFullPath(Path.Combine(projectDir, "MOS Word app", "bin", "Debug", "MOS Word app.exe"));
            if (File.Exists(debugPath)) return debugPath;
            string releasePath = Path.GetFullPath(Path.Combine(projectDir, "MOS Word app", "bin", "Release", "MOS Word app.exe"));
            if (File.Exists(releasePath)) return releasePath;
            return debugPath;
        }

        private void PowerPointButton_Click(object sender, RoutedEventArgs e)
        {
            if (_powerPointMainWindow == null || !_powerPointMainWindow.IsLoaded)
            {
                _powerPointMainWindow = new MainWindow();
                _powerPointMainWindow.Closed += PowerPointMainWindow_Closed;
            }
            _powerPointMainWindow.Show();
            _powerPointMainWindow.Activate();
            this.Hide();
        }

        private void PowerPointMainWindow_Closed(object sender, EventArgs e)
        {
            if (sender is MainWindow wnd)
            {
                wnd.Closed -= PowerPointMainWindow_Closed;
            }
            _powerPointMainWindow = null;
            Application.Current.MainWindow = this;
            this.Show();
            this.Activate();
        }

        /// <summary>
        /// Excel / Word など別プロセスで起動した科目アプリが終了したときに、表紙を再表示する。
        /// </summary>
        private void OnLaunchedProcessExited(object sender, EventArgs e)
        {
            try
            {
                if (sender is Process p)
                    p.Exited -= OnLaunchedProcessExited;
            }
            catch { }
            try
            {
                Dispatcher.Invoke(() =>
                {
                    if (!IsLoaded) return;
                    this.Show();
                    this.Activate();
                });
            }
            catch { }
        }
    }
}
