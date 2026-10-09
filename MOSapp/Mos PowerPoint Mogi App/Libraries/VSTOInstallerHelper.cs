using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Libraries
{
    /// <summary>
    /// PowerPoint VSTO の登録・LoadBehavior 管理。
    /// 開発用 PowerPointAddIn1 と MSI 用 PowerPointMosVsto を同時に 3 にしない（二重読み込み防止）。
    /// </summary>
    public static class VSTOInstallerHelper
    {
        private const string AddInName = "PowerPointAddIn1";
        /// <summary>.vsto 直接インストール時のレジストリ名（マニフェスト keyName と一致）。</summary>
        private const string RegistryAddInKeyNameFromManifest = "PowerPointAddIn1";
        /// <summary>powerpointvstosetup.vdproj（MSI）が登録するレジストリ名。</summary>
        private const string RegistryAddInKeyNameFromMsi = "PowerPointMosVsto";

        private static readonly string[] RegistryAddInKeyNames =
        {
            RegistryAddInKeyNameFromManifest,
            RegistryAddInKeyNameFromMsi,
        };

        private const string RegistryAddInsBasePath = @"Software\Microsoft\Office\PowerPoint\Addins";
        private const string LoadBehaviorValueName = "LoadBehavior";
        private const string UninstallBasePath = @"Software\Microsoft\Windows\CurrentVersion\Uninstall";

        private const string InstallerManufacturer = "Rabbit";
        private const string InstallerProductName = "powerpointvstosetup";

        private static Task<(bool success, string issue)> _backgroundPrepTask;

        /// <summary>
        /// 開発ビルドは PowerPointAddIn1、製品（Release + MSI）は PowerPointMosVsto を優先する。
        /// </summary>
        private static string ResolvePreferredAddInKeyName()
        {
#if DEBUG
            return RegistryAddInKeyNameFromManifest;
#else
            if (KeyHasManifest(RegistryAddInKeyNameFromMsi) || MsiDeployedVstoExists())
                return RegistryAddInKeyNameFromMsi;
            return RegistryAddInKeyNameFromManifest;
#endif
        }

        private static bool MsiDeployedVstoExists()
        {
            string vstoFileName = AddInName + ".vsto";
            foreach (string path in EnumerateInstallerDeployedPaths(vstoFileName))
            {
                if (File.Exists(path))
                    return true;
            }
            return false;
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
                    "[PptVSTOInstallerHelper] SetLoadBehavior " + leafName + ": " + ex.Message);
            }
        }

        /// <summary>優先キーだけ 3、もう一方は 0。優先キーが無い場合は存在する方にフォールバック。</summary>
        public static string ApplyExclusiveAddInLoadBehavior()
        {
            string active = ResolvePreferredAddInKeyName();
            if (!KeyHasManifest(active))
            {
                string fallback = string.Equals(active, RegistryAddInKeyNameFromManifest, StringComparison.OrdinalIgnoreCase)
                    ? RegistryAddInKeyNameFromMsi
                    : RegistryAddInKeyNameFromManifest;
                if (KeyHasManifest(fallback))
                    active = fallback;
            }

            foreach (string leaf in RegistryAddInKeyNames)
            {
                int behavior = string.Equals(leaf, active, StringComparison.OrdinalIgnoreCase) ? 3 : 0;
                SetLoadBehavior(leaf, behavior);
            }

            System.Diagnostics.Debug.WriteLine("[PptVSTOInstallerHelper] exclusive add-in active=" + active);
            return active;
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
                // 開発中は Debug/Release のどちらでも排他だけ確実に行う（勝手な再インストールはしない）
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
                    issue = "PowerPoint VSTO add-in is not registered. Run Deploy-PowerPointVsto.ps1 or install powerpointvstosetup.";
                    return false;
                }

                ApplyExclusiveAddInLoadBehavior();
                return true;
            }
            catch (Exception ex)
            {
                issue = ex.Message;
                System.Diagnostics.Debug.WriteLine("[PptVSTOInstallerHelper] EnsureAddInReadyForExam: " + ex.Message);
                return false;
            }
        }

        public static bool IsInstalled()
        {
            try
            {
                foreach (string leaf in RegistryAddInKeyNames)
                {
                    if (KeyHasManifest(leaf))
                        return true;
                }
                return false;
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
                string preferred = ResolvePreferredAddInKeyName();
                string[] order = preferred == RegistryAddInKeyNameFromMsi
                    ? new[] { RegistryAddInKeyNameFromMsi, RegistryAddInKeyNameFromManifest }
                    : new[] { RegistryAddInKeyNameFromManifest, RegistryAddInKeyNameFromMsi };

                foreach (string leaf in order)
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
                Path.Combine("PowerPointAddIn1", "bin", "Release", vstoFileName)
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
                Path.Combine("PowerPointAddIn1", "bin", "Release", vstoFileName)
            });
        }

        public static string GetBuildOutputPath()
        {
            return GetReleaseBuildOutputPath()
                ?? FindBuildOutputPath(new[]
                {
                    Path.Combine("PowerPointAddIn1", "bin", "Release", AddInName + ".vsto"),
                    Path.Combine("PowerPointAddIn1", "bin", "Debug", AddInName + ".vsto"),
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
                if (Directory.Exists(Path.Combine(searchDir, "PowerPointAddIn1")))
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

        private static IEnumerable<string> EnumerateRegisteredManifestPaths()
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string installed = GetInstallPath();
            if (!string.IsNullOrEmpty(installed))
                seen.Add(NormalizeFullPath(installed));

            foreach (string leaf in RegistryAddInKeyNames)
            {
                string path = Path.Combine(RegistryAddInsBasePath, leaf).Replace('/', '\\');
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path))
                {
                    if (key == null)
                        continue;
                    object manifest = key.GetValue("Manifest");
                    if (manifest == null)
                        continue;
                    string vsto = NormalizeManifestToVstoPath(manifest.ToString());
                    string n = NormalizeFullPath(vsto);
                    if (!string.IsNullOrEmpty(n))
                        seen.Add(n);
                }
            }

            foreach (string path in seen)
                yield return path;
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
            if (any)
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
                ApplyExclusiveAddInLoadBehavior();
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
