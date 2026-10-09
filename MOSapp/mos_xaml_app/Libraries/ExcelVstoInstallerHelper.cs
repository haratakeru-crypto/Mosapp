using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Libraries
{
    /// <summary>
    /// Excel VSTO（ExcelAddIn1）の登録・LoadBehavior 管理。
    /// 試験開始前に LoadBehavior=3 を保証し、未登録なら .vsto をサイレントインストールする。
    /// </summary>
    public static class ExcelVstoInstallerHelper
    {
        private const string AddInName = "ExcelAddIn1";
        private const string RegistryAddInKeyName = "ExcelAddIn1";
        private const string RegistryAddInsBasePath = @"Software\Microsoft\Office\Excel\Addins";
        private const string LoadBehaviorValueName = "LoadBehavior";
        private const string InstallerManufacturer = "Rabbit";
        private const string InstallerProductName = "excelvstosetup";

        private static Task<(bool success, string issue)> _backgroundPrepTask;

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
        /// 試験開始前に Excel VSTO を有効化する（未登録ならインストール、LoadBehavior=3）。
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

#if DEBUG
                if (!IsInstalled())
                {
                    string dev = GetBuildOutputPath();
                    if (!string.IsNullOrEmpty(dev) && File.Exists(dev))
                    {
                        if (!TrySilentInstall(dev, out issue))
                            return false;
                    }
                }
#else
                if ((!IsInstalled() || pointsToDebug || pointsToWrongBuild)
                    && !string.IsNullOrEmpty(releaseVsto)
                    && File.Exists(releaseVsto))
                {
                    if (!TrySilentInstall(releaseVsto, out issue))
                        return false;
                }
#endif

                if (!IsInstalled())
                {
                    issue = "Excel VSTO add-in is not registered. Install ExcelAddIn1.vsto or excelvstosetup.";
                    return false;
                }

                SetLoadBehavior(RegistryAddInKeyName, 3);
                ExcelVstoReadiness.RecordHostEvent("vsto-ensure ok loadBehavior=3");
                return true;
            }
            catch (Exception ex)
            {
                issue = ex.Message;
                System.Diagnostics.Debug.WriteLine("[ExcelVstoInstallerHelper] EnsureAddInReadyForExam: " + ex.Message);
                ExcelVstoReadiness.RecordHostEvent("vsto-ensure failed: " + ex.Message);
                return false;
            }
        }

        public static bool IsInstalled()
        {
            try
            {
                return KeyHasManifest(RegistryAddInKeyName);
            }
            catch
            {
                return false;
            }
        }

        public static string GetInstallPath()
        {
            try
            {
                string path = Path.Combine(RegistryAddInsBasePath, RegistryAddInKeyName).Replace('/', '\\');
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path))
                {
                    if (key == null)
                        return null;
                    object manifestPath = key.GetValue("Manifest");
                    if (manifestPath == null)
                        return null;
                    return NormalizeManifestToVstoPath(manifestPath.ToString());
                }
            }
            catch
            {
                return null;
            }
        }

        public static string GetReleaseBuildOutputPath()
        {
            string vstoFileName = AddInName + ".vsto";
#if DEBUG
            string devRelease = FindBuildOutputPath(new[]
            {
                Path.Combine("ExcelAddIn1", "bin", "Release", vstoFileName)
            });
            if (!string.IsNullOrEmpty(devRelease) && File.Exists(devRelease))
                return devRelease;
#endif
            foreach (string installedPath in EnumerateInstallerDeployedPaths(vstoFileName))
            {
                if (File.Exists(installedPath))
                    return installedPath;
            }

            return FindBuildOutputPath(new[]
            {
                Path.Combine("ExcelAddIn1", "bin", "Release", vstoFileName)
            });
        }

        public static string GetBuildOutputPath()
        {
            return GetReleaseBuildOutputPath()
                ?? FindBuildOutputPath(new[]
                {
                    Path.Combine("ExcelAddIn1", "bin", "Release", AddInName + ".vsto"),
                    Path.Combine("ExcelAddIn1", "bin", "Debug", AddInName + ".vsto"),
                });
        }

        private static IEnumerable<string> EnumerateInstallerDeployedPaths(string vstoFileName)
        {
            string relative = Path.Combine(InstallerManufacturer, InstallerProductName, vstoFileName);
            string pf64 = Environment.GetEnvironmentVariable("ProgramW6432");
            if (string.IsNullOrEmpty(pf64))
                pf64 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            if (!string.IsNullOrEmpty(pf64))
                yield return Path.Combine(pf64, relative);

            string pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            if (!string.IsNullOrEmpty(pf86) && !string.Equals(pf86, pf64, StringComparison.OrdinalIgnoreCase))
                yield return Path.Combine(pf86, relative);
        }

        private static bool KeyHasManifest(string leafName)
        {
            try
            {
                string path = Path.Combine(RegistryAddInsBasePath, leafName).Replace('/', '\\');
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path))
                {
                    if (key == null)
                        return false;
                    object manifest = key.GetValue("Manifest");
                    return manifest != null && !string.IsNullOrWhiteSpace(manifest.ToString());
                }
            }
            catch
            {
                return false;
            }
        }

        private static void SetLoadBehavior(string leafName, int loadBehavior)
        {
            try
            {
                string path = Path.Combine(RegistryAddInsBasePath, leafName).Replace('/', '\\');
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path, writable: true))
                {
                    if (key != null)
                        key.SetValue(LoadBehaviorValueName, loadBehavior, RegistryValueKind.DWord);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[ExcelVstoInstallerHelper] SetLoadBehavior " + leafName + ": " + ex.Message);
            }
        }

        private static string NormalizeFullPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return null;
            try { return Path.GetFullPath(path.Trim()); }
            catch { return path.Trim(); }
        }

        private static bool PathsEqual(string a, string b)
        {
            string na = NormalizeFullPath(a);
            string nb = NormalizeFullPath(b);
            return !string.IsNullOrEmpty(na)
                && !string.IsNullOrEmpty(nb)
                && string.Equals(na, nb, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeManifestToVstoPath(string manifestRaw)
        {
            if (string.IsNullOrWhiteSpace(manifestRaw))
                return null;

            string s = manifestRaw.Trim();
            int pipe = s.IndexOf('|');
            if (pipe >= 0)
                s = s.Substring(0, pipe);

            if (s.StartsWith("file:///", StringComparison.OrdinalIgnoreCase)
                || s.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var uri = new Uri(s);
                    return uri.LocalPath;
                }
                catch
                {
                    if (s.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                        s = s.Substring("file:///".Length).Replace('/', Path.DirectorySeparatorChar);
                }
            }

            return s.EndsWith(".vsto", StringComparison.OrdinalIgnoreCase) ? s : null;
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
                if (Directory.Exists(Path.Combine(searchDir, "ExcelAddIn1")))
                {
                    foreach (string suffix in relativeSuffixes)
                    {
                        string vstoPath = Path.Combine(searchDir, suffix);
                        if (File.Exists(vstoPath))
                            return vstoPath;
                    }
                }
            }

            return null;
        }

        private static string GetVstoInstallerPath()
        {
            string installer = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles),
                @"Microsoft Shared\VSTO\10.0\VSTOInstaller.exe");
            return File.Exists(installer) ? installer : null;
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
                string existing = GetInstallPath();
                if (!string.IsNullOrEmpty(existing) && File.Exists(existing))
                {
                    if (!RunVstoInstaller(installer, "/Uninstall \"" + existing + "\" /Silent", out issue))
                        return false;
                    System.Threading.Thread.Sleep(1500);
                }

                if (!RunVstoInstaller(installer, "/Install \"" + vstoPath + "\" /Silent", out issue))
                    return false;
                SetLoadBehavior(RegistryAddInKeyName, 3);
                return true;
            }
            catch (Exception ex)
            {
                issue = ex.Message;
                return false;
            }
        }
    }
}
