using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Libraries
{
    /// <summary>
    /// VSTOアドインのインストール状態を確認・管理するヘルパークラス
    /// </summary>
    public static class VSTOInstallerHelper
    {
        private const string AddInName = "New_MOSWordVSTOAddIn";
        /// <summary>.vsto 直接インストール時のレジストリ名（マニフェスト keyName と一致）。</summary>
        private const string RegistryAddInKeyNameFromManifest = "New_MOSWordVSTOAddIn";
        /// <summary>wordvstosetup.vdproj（MSI）が登録するレジストリ名。</summary>
        private const string RegistryAddInKeyNameFromMsi = "WordMosVsto";

        private static readonly string[] RegistryAddInKeyNames =
        {
            RegistryAddInKeyNameFromManifest,
            RegistryAddInKeyNameFromMsi,
        };

        private const string RegistryAddInsBasePath = @"Software\Microsoft\Office\Word\Addins";
        private const string LoadBehaviorValueName = "LoadBehavior";
        private const string UninstallBasePath = @"Software\Microsoft\Windows\CurrentVersion\Uninstall";

        private static string NormalizeFullPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;
            try
            {
                return Path.GetFullPath(path.Trim());
            }
            catch
            {
                return path.Trim();
            }
        }

        private static bool PathsEqual(string a, string b)
        {
            string na = NormalizeFullPath(a);
            string nb = NormalizeFullPath(b);
            return !string.IsNullOrEmpty(na)
                && !string.IsNullOrEmpty(nb)
                && string.Equals(na, nb, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Manifest の値から .vsto のローカルパスを取り出す（file:///… と |vstolocal 等に対応）。
        /// </summary>
        private static string NormalizeManifestToVstoPath(string manifestRaw)
        {
            if (string.IsNullOrWhiteSpace(manifestRaw))
                return null;

            string s = manifestRaw.Trim();
            int pipe = s.IndexOf('|');
            if (pipe >= 0)
                s = s.Substring(0, pipe);

            if (s.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var uri = new Uri(s);
                    return uri.LocalPath;
                }
                catch
                {
                    s = s.Substring("file:///".Length).Replace('/', Path.DirectorySeparatorChar);
                }
            }
            else if (s.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var uri = new Uri(s);
                    return uri.LocalPath;
                }
                catch { }
            }

            return s.EndsWith(".vsto", StringComparison.OrdinalIgnoreCase) ? s : null;
        }

        private static bool IsAddInKeyEnabled(RegistryKey addInKey)
        {
            if (addInKey == null)
                return false;

            object loadBehavior = addInKey.GetValue(LoadBehaviorValueName);
            if (loadBehavior == null)
                return false;

            int loadBehaviorValue = Convert.ToInt32(loadBehavior);
            // LoadBehavior: 0=切断, 1=接続, 2=登録済みだが起動時は読み込まない, 3=起動時に読み込む, 8/9=初回利用時読み込み
            // Word がエラー後に 3→2 に下げることがある。Manifest があれば「インストール済み」とみなす。
            if (loadBehaviorValue == 0)
                return false;

            object manifest = addInKey.GetValue("Manifest");
            return manifest != null && !string.IsNullOrWhiteSpace(manifest.ToString());
        }

        private static Task<(bool success, string issue)> _backgroundPrepTask;

        /// <summary>
        /// 起動直後に UI をブロックせず VSTO 準備を開始する。プロジェクト開始時は <see cref="EnsureAddInReadyForExam"/> で完了待ち。
        /// </summary>
        public static void StartBackgroundPrepForExam()
        {
            if (_backgroundPrepTask != null)
                return;

            _backgroundPrepTask = Task.Run(() =>
            {
                string issue;
                bool ok = EnsureAddInReadyForExamCore(out issue);
                return (ok, issue);
            });
        }

        /// <summary>
        /// 試験開始前に Release 版 VSTO を有効化する（Debug 登録の上書き、LoadBehavior=3）。
        /// バックグラウンド準備が走っていれば完了を待つ。
        /// </summary>
        public static bool EnsureAddInReadyForExam(out string issue)
        {
            issue = null;
            if (_backgroundPrepTask != null)
            {
                try
                {
                    var result = _backgroundPrepTask.GetAwaiter().GetResult();
                    issue = result.issue;
                    return result.success;
                }
                catch (Exception ex)
                {
                    issue = ex.InnerException?.Message ?? ex.Message;
                    return false;
                }
            }

            return EnsureAddInReadyForExamCore(out issue);
        }

        private static bool EnsureAddInReadyForExamCore(out string issue)
        {
            issue = null;
            try
            {
                string releaseVsto = GetReleaseBuildOutputPath();
                string installed = GetInstallPath();
                bool pointsToDebug = !string.IsNullOrEmpty(installed)
                    && installed.IndexOf("\\bin\\Debug\\", StringComparison.OrdinalIgnoreCase) >= 0;
                bool pointsToWrongBuild = !string.IsNullOrEmpty(installed)
                    && !string.IsNullOrEmpty(releaseVsto)
                    && !PathsEqual(installed, releaseVsto);

                if ((!IsInstalled() || pointsToDebug || pointsToWrongBuild)
                    && !string.IsNullOrEmpty(releaseVsto)
                    && File.Exists(releaseVsto))
                {
                    if (!TrySilentInstall(releaseVsto, out issue))
                        return false;
                }

                if (!IsInstalled())
                {
                    issue = "VSTO add-in is not registered. Run Rebuild-And-Install-WordVSTO.ps1.";
                    return false;
                }

                SetLoadBehaviorForAllKeys(3);
                return true;
            }
            catch (Exception ex)
            {
                issue = ex.Message;
                System.Diagnostics.Debug.WriteLine($"[VSTOInstallerHelper] EnsureAddInReadyForExam: {ex.Message}");
                return false;
            }
        }

        private static void SetLoadBehaviorForAllKeys(int loadBehavior)
        {
            foreach (string leaf in RegistryAddInKeyNames)
            {
                string path = Path.Combine(RegistryAddInsBasePath, leaf).Replace('/', '\\');
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path, writable: true))
                {
                    if (key != null)
                        key.SetValue(LoadBehaviorValueName, loadBehavior, RegistryValueKind.DWord);
                }
            }

            string primary = Path.Combine(RegistryAddInsBasePath, RegistryAddInKeyNameFromManifest).Replace('/', '\\');
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(primary, writable: true))
            {
                if (key != null)
                    key.SetValue(LoadBehaviorValueName, loadBehavior, RegistryValueKind.DWord);
            }
        }

        private static bool TrySilentInstall(string vstoPath, out string issue)
        {
            issue = null;
            if (string.IsNullOrWhiteSpace(vstoPath) || !File.Exists(vstoPath))
            {
                issue = ".vsto not found: " + (vstoPath ?? "(null)");
                return false;
            }

            string installer = GetVstoInstallerPath();
            if (string.IsNullOrEmpty(installer))
            {
                issue = "VSTOInstaller.exe not found.";
                return false;
            }

            try
            {
                if (!TryUninstallAllRegisteredManifests(installer, out issue))
                    return false;

                if (!RunVstoInstaller(installer, "/Install \"" + vstoPath + "\" /Silent", out issue))
                    return false;

                SetLoadBehaviorForAllKeys(3);
                return true;
            }
            catch (Exception ex)
            {
                issue = ex.Message;
                return false;
            }
        }

        private static string GetVstoInstallerPath()
        {
            string installer = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles),
                @"Microsoft Shared\VSTO\10.0\VSTOInstaller.exe");
            return File.Exists(installer) ? installer : null;
        }

        private static IEnumerable<string> EnumerateRegisteredManifestPaths()
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            string installed = GetInstallPath();
            if (!string.IsNullOrEmpty(installed))
                seen.Add(NormalizeFullPath(installed));

            foreach (string manifest in EnumerateUninstallEntryManifests())
            {
                string normalized = NormalizeFullPath(manifest);
                if (!string.IsNullOrEmpty(normalized))
                    seen.Add(normalized);
            }

            foreach (string path in seen)
                yield return path;
        }

        private static IEnumerable<string> EnumerateUninstallEntryManifests()
        {
            using (RegistryKey uninstallRoot = Registry.CurrentUser.OpenSubKey(UninstallBasePath))
            {
                if (uninstallRoot == null)
                    yield break;

                foreach (string subKeyName in uninstallRoot.GetSubKeyNames())
                {
                    using (RegistryKey subKey = uninstallRoot.OpenSubKey(subKeyName))
                    {
                        if (subKey == null)
                            continue;

                        string displayName = subKey.GetValue("DisplayName") as string;
                        if (!string.Equals(displayName, AddInName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        string uninstallString = subKey.GetValue("UninstallString") as string;
                        string manifest = ParseManifestFromUninstallString(uninstallString);
                        if (!string.IsNullOrEmpty(manifest))
                            yield return manifest;

                        string urlUpdate = subKey.GetValue("UrlUpdateInfo") as string;
                        string updateManifest = NormalizeManifestToVstoPath(urlUpdate);
                        if (!string.IsNullOrEmpty(updateManifest))
                            yield return updateManifest;
                    }
                }
            }
        }

        private static string ParseManifestFromUninstallString(string uninstallString)
        {
            if (string.IsNullOrWhiteSpace(uninstallString))
                return null;

            const string marker = "/Uninstall ";
            int idx = uninstallString.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
                return null;

            string tail = uninstallString.Substring(idx + marker.Length).Trim();
            int space = tail.IndexOf(' ');
            if (space > 0)
                tail = tail.Substring(0, space);

            return NormalizeManifestToVstoPath(tail.Trim('"'));
        }

        private static bool TryUninstallAllRegisteredManifests(string installer, out string issue)
        {
            issue = null;
            bool any = false;
            foreach (string manifestPath in EnumerateRegisteredManifestPaths())
            {
                if (string.IsNullOrEmpty(manifestPath))
                    continue;

                any = true;
                if (!RunVstoInstaller(installer, "/Uninstall \"" + manifestPath + "\" /Silent", out issue))
                    return false;
            }

            if (!any)
                return true;

            System.Threading.Thread.Sleep(1500);
            return true;
        }

        private static bool RunVstoInstaller(string installer, string arguments, out string issue)
        {
            issue = null;
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = installer,
                Arguments = arguments,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using (var proc = System.Diagnostics.Process.Start(psi))
            {
                if (proc == null)
                {
                    issue = "Failed to start VSTOInstaller.";
                    return false;
                }

                proc.WaitForExit(120000);
                if (proc.ExitCode != 0)
                {
                    issue = "VSTOInstaller failed (" + arguments + "): exit " + proc.ExitCode;
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// VSTOアドインがインストールされているかチェック
        /// </summary>
        /// <returns>インストールされている場合true</returns>
        public static bool IsInstalled()
        {
            try
            {
                foreach (string leaf in RegistryAddInKeyNames)
                {
                    string path = Path.Combine(RegistryAddInsBasePath, leaf).Replace('/', '\\');
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path))
                    {
                        if (IsAddInKeyEnabled(key))
                            return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VSTOInstallerHelper] Error checking registry: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// VSTOアドインのインストールパスを取得
        /// </summary>
        /// <returns>インストールパス（見つからない場合はnull）</returns>
        public static string GetInstallPath()
        {
            try
            {
                foreach (string leaf in RegistryAddInKeyNames)
                {
                    string path = Path.Combine(RegistryAddInsBasePath, leaf).Replace('/', '\\');
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path))
                    {
                        if (key == null)
                            continue;

                        object manifestPath = key.GetValue("Manifest");
                        if (manifestPath == null)
                            continue;

                        string vsto = NormalizeManifestToVstoPath(manifestPath.ToString());
                        if (!string.IsNullOrEmpty(vsto))
                            return vsto;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VSTOInstallerHelper] Error getting install path: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// インストーラー（wordvstosetup.vdproj）による既定の配置先。
        /// Manufacturer=Rabbit, ProductName=wordvstosetup, DefaultLocation=[ProgramFiles64Folder][Manufacturer]\[ProductName]
        /// </summary>
        private const string InstallerManufacturer = "Rabbit";
        private const string InstallerProductName = "wordvstosetup";

        /// <summary>試験用の Release 版 .vsto のみを返す（Debug は登録対象外）。</summary>
        public static string GetReleaseBuildOutputPath()
        {
            string vstoFileName = AddInName + ".vsto";
            foreach (string installedPath in EnumerateInstallerDeployedPaths(vstoFileName))
            {
                if (File.Exists(installedPath))
                    return installedPath;
            }

            return FindBuildOutputPath(new[]
            {
                Path.Combine("New_MOSWordVSTOAddIn", "New_MOSWordVSTOAddIn", "bin", "Release", vstoFileName)
            });
        }

        /// <summary>
        /// VSTOアドインのビルド出力パス（または配布配置パス）を取得する。
        /// </summary>
        public static string GetBuildOutputPath()
        {
            return GetReleaseBuildOutputPath()
                ?? FindBuildOutputPath(new[]
                {
                    Path.Combine("New_MOSWordVSTOAddIn", "New_MOSWordVSTOAddIn", "bin", "Release", AddInName + ".vsto"),
                    Path.Combine("New_MOSWordVSTOAddIn", "New_MOSWordVSTOAddIn", "bin", "Debug", AddInName + ".vsto"),
                });
        }

        private static string FindBuildOutputPath(string[] relativeSuffixes)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            foreach (string suffix in relativeSuffixes)
            {
                string vstoPath = Path.Combine(baseDir, suffix);
                if (File.Exists(vstoPath))
                    return vstoPath;
            }

            string searchDir = Path.GetFullPath(baseDir);
            for (int i = 0; i < 8; i++)
            {
                string parent = Path.GetDirectoryName(searchDir);
                if (string.IsNullOrEmpty(parent) || parent == searchDir)
                    break;
                searchDir = parent;
                foreach (string suffix in relativeSuffixes)
                {
                    string vstoPath = Path.Combine(searchDir, suffix);
                    if (File.Exists(vstoPath))
                        return vstoPath;
                }

                if (Directory.Exists(Path.Combine(searchDir, "New_MOSWordVSTOAddIn")))
                {
                    foreach (string suffix in relativeSuffixes)
                    {
                        string vstoPath = Path.Combine(searchDir, suffix);
                        if (File.Exists(vstoPath))
                            return vstoPath;
                    }
                }
            }

            string asmDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            if (!string.IsNullOrEmpty(asmDir))
            {
                searchDir = Path.GetFullPath(asmDir);
                for (int i = 0; i < 8; i++)
                {
                    foreach (string suffix in relativeSuffixes)
                    {
                        string vstoPath = Path.Combine(searchDir, suffix);
                        if (File.Exists(vstoPath))
                            return vstoPath;
                    }

                    string parent = Path.GetDirectoryName(searchDir);
                    if (string.IsNullOrEmpty(parent) || parent == searchDir)
                        break;
                    searchDir = parent;
                }
            }

            string currentDir = Directory.GetCurrentDirectory();
            foreach (string suffix in relativeSuffixes)
            {
                string vstoPath = Path.Combine(currentDir, suffix);
                if (File.Exists(vstoPath))
                    return vstoPath;
            }

            return null;
        }

        /// <summary>
        /// インストーラー（wordvstosetup）が配置する可能性のある .vsto パス候補を列挙する。
        /// 64bit / 32bit Program Files 双方を対象とする。
        /// </summary>
        private static System.Collections.Generic.IEnumerable<string> EnumerateInstallerDeployedPaths(string vstoFileName)
        {
            string relative = Path.Combine(InstallerManufacturer, InstallerProductName, vstoFileName);

            // 64bit Program Files (ProgramW6432 を優先し、なければ SpecialFolder)
            string pf64 = Environment.GetEnvironmentVariable("ProgramW6432");
            if (string.IsNullOrEmpty(pf64))
            {
                pf64 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            }
            if (!string.IsNullOrEmpty(pf64))
            {
                yield return Path.Combine(pf64, relative);
            }

            // 32bit Program Files (x86) も念のため
            string pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            if (!string.IsNullOrEmpty(pf86) && !string.Equals(pf86, pf64, StringComparison.OrdinalIgnoreCase))
            {
                yield return Path.Combine(pf86, relative);
            }
        }

        /// <summary>
        /// VSTOアドインのビルド出力ファイルが存在するかチェック
        /// </summary>
        /// <returns>ファイルが存在する場合true</returns>
        public static bool BuildOutputExists()
        {
            string vstoPath = GetBuildOutputPath();
            if (string.IsNullOrEmpty(vstoPath))
                return false;

            // .vstoファイルと.dllファイルの両方が存在するか確認
            string dllPath = Path.Combine(Path.GetDirectoryName(vstoPath), $"{AddInName}.dll");
            return File.Exists(vstoPath) && File.Exists(dllPath);
        }

        /// <summary>
        /// VSTOアドインのインストール状態を詳細にチェック
        /// </summary>
        /// <returns>インストール状態の詳細情報</returns>
        public static VSTOInstallStatus GetInstallStatus()
        {
            var status = new VSTOInstallStatus
            {
                IsInstalled = IsInstalled(),
                InstallPath = GetInstallPath(),
                BuildOutputPath = GetBuildOutputPath(),
                BuildOutputExists = BuildOutputExists()
            };

            return status;
        }

        /// <summary>
        /// VSTOアドインのインストール状態を表すクラス
        /// </summary>
        public class VSTOInstallStatus
        {
            public bool IsInstalled { get; set; }
            public string InstallPath { get; set; }
            public string BuildOutputPath { get; set; }
            public bool BuildOutputExists { get; set; }

            public string GetInstallationMessage()
            {
                if (IsInstalled)
                {
                    return "VSTOアドインはインストールされています。";
                }

                string expectedInstallerPath = System.IO.Path.Combine(
                    System.Environment.GetEnvironmentVariable("ProgramW6432")
                        ?? System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles),
                    "Rabbit", "wordvstosetup", "New_MOSWordVSTOAddIn.vsto");

                if (!BuildOutputExists)
                {
                    return $"VSTOアドインの .vsto ファイルが見つかりません。\n\n" +
                           $"想定される配置先:\n" +
                           $"   {expectedInstallerPath}\n\n" +
                           $"wordvstosetup インストーラー（MSI）を実行して配置してください。";
                }

                return $"VSTOアドインがインストールされていません。\n\n" +
                       $"インストール方法:\n" +
                       $"1. Wordを終了してください\n" +
                       $"2. 以下のファイルをダブルクリックしてインストールしてください:\n" +
                       $"   {BuildOutputPath ?? expectedInstallerPath}";
            }
        }
    }
}






