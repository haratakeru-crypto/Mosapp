using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;

namespace MOSapp.Views
{
    /// <summary>
    /// 科目選択（表紙）画面。パスワード認証後に表示され、Excel / Word / PowerPoint のいずれかを選択する。
    /// </summary>
    public partial class SubjectSelectionWindow : Window
    {
        public SubjectSelectionWindow()
        {
            InitializeComponent();
            Closing += SubjectSelectionWindow_Closing;
        }

        private void SubjectSelectionWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var result = MessageBox.Show("アプリ自体を終了します。本当にいいですか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
                e.Cancel = true;
        }

        /// <summary>
        /// ソリューションルート（.slnx と mos_xaml_app, MOS Word app があるフォルダ）を取得する。
        /// baseDir は ...\MOSapp\MOSapp\bin\Debug なので、3階層上でソリューションルート。
        /// </summary>
        private static string GetSolutionRoot()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string dir = Path.GetDirectoryName(baseDir);   // ...\MOSapp\MOSapp\bin
            if (string.IsNullOrEmpty(dir)) return null;
            dir = Path.GetDirectoryName(dir);             // ...\MOSapp\MOSapp (プロジェクトフォルダ)
            if (string.IsNullOrEmpty(dir)) return null;
            return Path.GetDirectoryName(dir);            // ...\MOSapp (ソリューションルート = .slnx と同階層)
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
                excelInitialPath = null;
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

        private static string GetExcelInitialPath()
        {
            const string defaultPath = @"C:\MOSTest\Excel365\Tab1\Initial";
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (!File.Exists(configPath)) return Path.GetFullPath(defaultPath);
                string json = File.ReadAllText(configPath);
                var path = GetJsonStringValue(json, "excelInitialPath");
                if (string.IsNullOrWhiteSpace(path)) return Path.GetFullPath(defaultPath);
                path = path.Trim();
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                return Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(baseDir, path));
            }
            catch
            {
                return Path.GetFullPath(defaultPath);
            }
        }

        /// <summary>
        /// Excel 模擬アプリの exe パス。excelMogiApp を最優先、次に mos_xaml_app をフォールバック。config.json は任意。
        /// </summary>
        private static string GetExcelAppExePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string solutionRoot = GetSolutionRoot();
            if (!string.IsNullOrEmpty(solutionRoot))
            {
                string excelMogiRelease = Path.Combine(solutionRoot, "excelMogiApp", "bin", "Release", "excelMogiApp.exe");
                if (File.Exists(excelMogiRelease)) return Path.GetFullPath(excelMogiRelease);
                string excelMogiDebug = Path.Combine(solutionRoot, "excelMogiApp", "bin", "Debug", "excelMogiApp.exe");
                if (File.Exists(excelMogiDebug)) return Path.GetFullPath(excelMogiDebug);
                string mosXamlRelease = Path.Combine(solutionRoot, "mos_xaml_app", "bin", "Release", "MOSExcelMogiApp.exe");
                if (File.Exists(mosXamlRelease)) return Path.GetFullPath(mosXamlRelease);
                string mosXamlDebug = Path.Combine(solutionRoot, "mos_xaml_app", "bin", "Debug", "MOSExcelMogiApp.exe");
                if (File.Exists(mosXamlDebug)) return Path.GetFullPath(mosXamlDebug);
            }
            try
            {
                string configPath = Path.Combine(baseDir, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var path = GetJsonStringValue(json, "excelExePath");
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        path = path.Trim();
                        string absolutePath = Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(baseDir, path));
                        if (File.Exists(absolutePath)) return absolutePath;
                    }
                }
            }
            catch { }
            string sameDirExe = Path.Combine(baseDir, "MOSExcelMogiApp.exe");
            if (File.Exists(sameDirExe)) return Path.GetFullPath(sameDirExe);
            return null;
        }

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

        private static string GetWordInitialPath()
        {
            const string defaultPath = @"C:\MOSTest\Word365\Tab1\Initial";
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (!File.Exists(configPath)) return Path.GetFullPath(defaultPath);
                string json = File.ReadAllText(configPath);
                var path = GetJsonStringValue(json, "wordInitialPath");
                if (string.IsNullOrWhiteSpace(path)) return Path.GetFullPath(defaultPath);
                path = path.Trim();
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                return Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(baseDir, path));
            }
            catch
            {
                return Path.GetFullPath(defaultPath);
            }
        }

        /// <summary>
        /// Word 模擬アプリの exe パス。表紙が Release で動いていれば Word の Release を優先、Debug なら Debug を優先。
        /// </summary>
        private static string GetWordAppExePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            bool preferRelease = baseDir.IndexOf("\\Release\\", StringComparison.OrdinalIgnoreCase) >= 0;
            string solutionRoot = GetSolutionRoot();
            if (!string.IsNullOrEmpty(solutionRoot))
            {
                string wordDir = Path.Combine(solutionRoot, "MOS Word app", "bin");
                string releasePath = Path.Combine(wordDir, "Release", "MOS Word app.exe");
                string debugPath = Path.Combine(wordDir, "Debug", "MOS Word app.exe");
                if (preferRelease)
                {
                    if (File.Exists(releasePath)) return Path.GetFullPath(releasePath);
                    if (File.Exists(debugPath)) return Path.GetFullPath(debugPath);
                }
                else
                {
                    if (File.Exists(debugPath)) return Path.GetFullPath(debugPath);
                    if (File.Exists(releasePath)) return Path.GetFullPath(releasePath);
                }
            }
            try
            {
                string configPath = Path.Combine(baseDir, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var path = GetJsonStringValue(json, "wordExePath");
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        path = path.Trim();
                        string absolutePath = Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(baseDir, path));
                        if (File.Exists(absolutePath)) return absolutePath;
                    }
                }
            }
            catch { }
            return null;
        }

        private void PowerPointButton_Click(object sender, RoutedEventArgs e)
        {
            string pptExePath = GetPowerPointAppExePath();
            if (string.IsNullOrEmpty(pptExePath) || !File.Exists(pptExePath))
            {
                MessageBox.Show("PowerPoint アプリが見つかりません。\n" + (pptExePath ?? ""), "科目選択", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = pptExePath,
                    Arguments = "--direct",
                    UseShellExecute = false,
                    WorkingDirectory = Path.GetDirectoryName(pptExePath)
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
                MessageBox.Show($"PowerPoint アプリの起動に失敗しました: {ex.Message}", "科目選択", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// PowerPoint 模擬アプリの exe パス。Mos PowerPoint Mogi App を優先、表紙が Release なら Release を優先。
        /// </summary>
        private static string GetPowerPointAppExePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            bool preferRelease = baseDir.IndexOf("\\Release\\", StringComparison.OrdinalIgnoreCase) >= 0;
            string solutionRoot = GetSolutionRoot();
            const string exeName = "Mos PowerPoint Mogi App.exe";
            if (!string.IsNullOrEmpty(solutionRoot))
            {
                string mogiDir = Path.Combine(solutionRoot, "Mos PowerPoint Mogi App", "bin");
                string releasePath = Path.Combine(mogiDir, "Release", exeName);
                string debugPath = Path.Combine(mogiDir, "Debug", exeName);
                if (preferRelease)
                {
                    if (File.Exists(releasePath)) return Path.GetFullPath(releasePath);
                    if (File.Exists(debugPath)) return Path.GetFullPath(debugPath);
                }
                else
                {
                    if (File.Exists(debugPath)) return Path.GetFullPath(debugPath);
                    if (File.Exists(releasePath)) return Path.GetFullPath(releasePath);
                }
                // フォールバック: フォルダ未リネーム時は従来のパスを参照
                string legacyDir = Path.Combine(solutionRoot, "MOS PowerPoint app", "bin");
                string legacyRelease = Path.Combine(legacyDir, "Release", "MOS PowerPoint app.exe");
                string legacyDebug = Path.Combine(legacyDir, "Debug", "MOS PowerPoint app.exe");
                if (preferRelease)
                {
                    if (File.Exists(legacyRelease)) return Path.GetFullPath(legacyRelease);
                    if (File.Exists(legacyDebug)) return Path.GetFullPath(legacyDebug);
                }
                else
                {
                    if (File.Exists(legacyDebug)) return Path.GetFullPath(legacyDebug);
                    if (File.Exists(legacyRelease)) return Path.GetFullPath(legacyRelease);
                }
            }
            try
            {
                string configPath = Path.Combine(baseDir, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    var path = GetJsonStringValue(json, "powerPointExePath");
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        path = path.Trim();
                        string absolutePath = Path.IsPathRooted(path) ? Path.GetFullPath(path) : Path.GetFullPath(Path.Combine(baseDir, path));
                        if (File.Exists(absolutePath)) return absolutePath;
                    }
                }
            }
            catch { }
            return null;
        }

        private static string GetJsonStringValue(string json, string key)
        {
            if (string.IsNullOrEmpty(json) || string.IsNullOrEmpty(key)) return null;
            var match = Regex.Match(json, "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"([^\"]*)\"");
            return match.Success ? match.Groups[1].Value : null;
        }

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
