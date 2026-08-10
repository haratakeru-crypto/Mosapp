using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using Newtonsoft.Json.Linq;

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// PowerPoint教材のデータルートと正規ファイル配置を一元管理する。
    /// </summary>
    public static class PowerPointDataPathHelper
    {
        public const string DefaultDataRoot = @"C:\MOSTest\PowerPoint365";

        public static string GetDataRoot()
        {
            string dataRoot = null;
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    JObject config = JObject.Parse(File.ReadAllText(configPath));
                    dataRoot = config["powerPointDataPath"]?.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[PowerPointDataPathHelper] Assets\\config.json 読み込みエラー: {ex.Message}");
            }

            // 旧配置との互換性のためだけに App.config を参照する。
            if (string.IsNullOrWhiteSpace(dataRoot))
                dataRoot = ConfigurationManager.AppSettings["PowerPointDataPath"];

            if (string.IsNullOrWhiteSpace(dataRoot))
                dataRoot = DefaultDataRoot;

            return Environment.ExpandEnvironmentVariables(dataRoot.Trim());
        }

        public static string GetTabFolder(int groupId)
        {
            return Path.Combine(GetDataRoot(), $"Tab{groupId}");
        }

        public static string GetWorkingProjectPath(int groupId, int projectId)
        {
            string tabFolder = GetTabFolder(groupId);
            string canonicalPath = Path.Combine(tabFolder, $"Project{projectId}.pptx");
            return EnsureCanonicalFile(
                canonicalPath,
                new[]
                {
                    Path.Combine(tabFolder, $"project{projectId}.pptx")
                });
        }

        public static string GetTemplateProjectPath(int groupId, int projectId)
        {
            string dataRoot = GetDataRoot();
            string canonicalFolder = Path.Combine(dataRoot, "Templates", $"Tab{groupId}");
            string canonicalPath = Path.Combine(canonicalFolder, $"Project{projectId}.pptx");
            string resolvedPath = EnsureCanonicalFile(
                canonicalPath,
                new[]
                {
                    Path.Combine(canonicalFolder, $"project{projectId}.pptx"),
                    Path.Combine(dataRoot, "Templates", $"Project{projectId}.pptx"),
                    Path.Combine(dataRoot, "Templates", $"project{projectId}.pptx"),
                    Path.Combine(canonicalFolder, $"Tab{groupId}_Project{projectId}.pptx"),
                    Path.Combine(canonicalFolder, $"Tab{groupId}_project{projectId}.pptx")
                });

            ClearReadOnly(resolvedPath, "テンプレート");
            return resolvedPath;
        }

        public static string GetInitialProjectPath(int groupId, int projectId)
        {
            string initialFolder = Path.Combine(GetTabFolder(groupId), "Initial");
            string canonicalPath = Path.Combine(initialFolder, $"Project{projectId}.pptx");
            return EnsureCanonicalFile(
                canonicalPath,
                new[]
                {
                    Path.Combine(initialFolder, $"project{projectId}.pptx")
                });
        }

        public static string ResolveJsonPath(string fileName)
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string referencesPath = Path.Combine(baseDirectory, "References", "JSON", fileName);
            if (File.Exists(referencesPath))
                return referencesPath;

            return Path.Combine(baseDirectory, fileName);
        }

        public static void ClearReadOnly(string path, string fileKind)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            try
            {
                var fileInfo = new FileInfo(path);
                if (fileInfo.IsReadOnly)
                    fileInfo.IsReadOnly = false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[PowerPointDataPathHelper] 読み取り専用解除（{fileKind}）: {ex.Message}");
            }
        }

        private static string EnsureCanonicalFile(string canonicalPath, IEnumerable<string> legacyCandidates)
        {
            if (File.Exists(canonicalPath))
                return canonicalPath;

            foreach (string candidate in legacyCandidates)
            {
                if (string.IsNullOrWhiteSpace(candidate) ||
                    string.Equals(candidate, canonicalPath, StringComparison.Ordinal) ||
                    !File.Exists(candidate))
                {
                    continue;
                }

                Directory.CreateDirectory(Path.GetDirectoryName(canonicalPath));
                try
                {
                    File.Copy(candidate, canonicalPath, overwrite: false);
                    ClearReadOnly(canonicalPath, "正規配置");
                    return canonicalPath;
                }
                catch (IOException)
                {
                    // 別処理が先に正規ファイルを作成した場合は、そのファイルを使用する。
                    if (File.Exists(canonicalPath))
                        return canonicalPath;
                    throw;
                }
            }

            return canonicalPath;
        }
    }
}
