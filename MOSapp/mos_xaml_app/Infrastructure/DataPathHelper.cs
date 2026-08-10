using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;

namespace MOSExcelMogiApp.Infrastructure
{
    /// <summary>
    /// Excel 教材データの正規パスと旧配置からの移行を一元管理する。
    /// </summary>
    public static class DataPathHelper
    {
        private const string DefaultExcelDataPath = @"C:\MOSTest\Excel365";

        public static string ConfigPath =>
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");

        public static string ExcelDataRoot
        {
            get
            {
                try
                {
                    if (File.Exists(ConfigPath))
                    {
                        var config = JObject.Parse(File.ReadAllText(ConfigPath));
                        string configuredRoot = config["excelDataPath"]?.ToString();
                        if (!string.IsNullOrWhiteSpace(configuredRoot))
                            return Path.GetFullPath(Environment.ExpandEnvironmentVariables(configuredRoot));
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DataPathHelper] excelDataPath read failed: {ex.Message}");
                }

                return DefaultExcelDataPath;
            }
        }

        public static string GetWorkingFilePath(int groupId, int projectId) =>
            Path.Combine(ExcelDataRoot, $"Tab{groupId}", $"Project{projectId}.xlsx");

        public static string GetTemplateFilePath(int groupId, int projectId) =>
            Path.Combine(ExcelDataRoot, "Templates", $"Tab{groupId}", $"Project{projectId}.xlsx");

        public static string GetInitialFilePath(int groupId, int projectId) =>
            Path.Combine(ExcelDataRoot, $"Tab{groupId}", "Initial", $"Project{projectId}.xlsx");

        public static string ResolveWorkingFilePath(int groupId, int projectId, string configuredFallback = null)
        {
            string canonical = GetWorkingFilePath(groupId, projectId);
            return EnsureCanonicalFile(canonical, new[]
            {
                Path.Combine(ExcelDataRoot, $"Tab{groupId}", $"project{projectId}.xlsx"),
                // 旧版では Initial を作業ファイルとして開いていた。
                Path.Combine(ExcelDataRoot, $"Tab{groupId}", "Initial", $"Project{projectId}.xlsx"),
                Path.Combine(ExcelDataRoot, $"Tab{groupId}", "Initial", $"project{projectId}.xlsx"),
                configuredFallback
            });
        }

        public static string ResolveTemplateFilePath(int groupId, int projectId)
        {
            string canonical = GetTemplateFilePath(groupId, projectId);
            return EnsureCanonicalFile(canonical, new[]
            {
                Path.Combine(ExcelDataRoot, "Templates", $"Tab{groupId}", $"project{projectId}.xlsx"),
                Path.Combine(ExcelDataRoot, "Templates", $"Project{projectId}.xlsx"),
                Path.Combine(ExcelDataRoot, "Templates", $"project{projectId}.xlsx"),
                Path.Combine(ExcelDataRoot, "Templates", $"Tab{groupId}", $"Tab{groupId}_project{projectId}.xlsx")
            });
        }

        public static string ResolveInitialFilePath(int groupId, int projectId, string configuredFallback = null)
        {
            string canonical = GetInitialFilePath(groupId, projectId);
            return EnsureCanonicalFile(canonical, new[]
            {
                Path.Combine(ExcelDataRoot, $"Tab{groupId}", "Initial", $"project{projectId}.xlsx"),
                configuredFallback
            });
        }

        public static string ResolveVariantWorkingFilePath(
            int groupId, int projectId, int variantSetNo, string configuredFallback = null)
        {
            string variantDirectory = Path.Combine(
                ExcelDataRoot, $"Tab{groupId}", $"PracticeVariant{variantSetNo}");
            string canonical = Path.Combine(variantDirectory, $"Project{projectId}.xlsx");
            return EnsureCanonicalFile(canonical, new[]
            {
                Path.Combine(variantDirectory, $"project{projectId}.xlsx"),
                configuredFallback
            });
        }

        public static string ResolveVariantTemplateFilePath(int groupId, int projectId, int variantSetNo)
        {
            string templateDirectory = Path.Combine(
                ExcelDataRoot, $"Tab{groupId}", $"PracticeVariant{variantSetNo}", "Templates");
            string canonical = Path.Combine(templateDirectory, $"Project{projectId}.xlsx");
            return EnsureCanonicalFile(canonical, new[]
            {
                Path.Combine(templateDirectory, $"project{projectId}.xlsx")
            });
        }

        public static string ResolveJsonPath(string fileName)
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string referencesPath = Path.Combine(baseDirectory, "References", "JSON", fileName);
            return File.Exists(referencesPath)
                ? referencesPath
                : Path.Combine(baseDirectory, fileName);
        }

        private static string EnsureCanonicalFile(string canonicalPath, IEnumerable<string> fallbackPaths)
        {
            if (File.Exists(canonicalPath))
                return canonicalPath;

            foreach (string fallbackPath in fallbackPaths)
            {
                if (string.IsNullOrWhiteSpace(fallbackPath) ||
                    !Path.IsPathRooted(fallbackPath) ||
                    !File.Exists(fallbackPath))
                {
                    continue;
                }

                try
                {
                    string destinationDirectory = Path.GetDirectoryName(canonicalPath);
                    if (!Directory.Exists(destinationDirectory))
                        Directory.CreateDirectory(destinationDirectory);

                    File.Copy(fallbackPath, canonicalPath, overwrite: false);
                    ClearReadOnly(canonicalPath);
                    System.Diagnostics.Debug.WriteLine(
                        $"[DataPathHelper] Migrated legacy file without deleting source: {fallbackPath} -> {canonicalPath}");
                    return canonicalPath;
                }
                catch (IOException)
                {
                    if (File.Exists(canonicalPath))
                        return canonicalPath;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DataPathHelper] Legacy migration failed: {ex.Message}");
                }
            }

            return canonicalPath;
        }

        public static void ClearReadOnly(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return;

            var fileInfo = new FileInfo(filePath);
            if (fileInfo.IsReadOnly)
                fileInfo.IsReadOnly = false;
        }
    }
}
