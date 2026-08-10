using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Newtonsoft.Json;

namespace MOS_Word_app
{
    /// <summary>Word 教材データと問題 JSON の配置規約を一元管理する。</summary>
    public static class WordDataPathHelper
    {
        private const string DefaultWordDataPath = @"C:\MOSTest\Word365";

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool DeleteFileW(string lpFileName);

        private sealed class WordAppConfig
        {
            [JsonProperty("wordDataPath")]
            public string WordDataPath { get; set; }
        }

        public static string RootPath { get; } = LoadRootPath();

        public static string GetWorkingFolder(int groupId)
        {
            return Path.Combine(RootPath, $"Tab{groupId}");
        }

        public static string GetWorkingFilePath(int groupId, int projectId)
        {
            return Path.Combine(GetWorkingFolder(groupId), GetCanonicalFileName(projectId));
        }

        public static string GetWorkingFilePathForSource(int groupId, int projectId, string sourcePath)
        {
            string extension = Path.GetExtension(sourcePath);
            if (!extension.Equals(".doc", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
                extension = projectId == 7 ? ".doc" : ".docx";
            return Path.Combine(GetWorkingFolder(groupId), $"Project{projectId}{extension}");
        }

        public static string GetTemplateFolder(int groupId)
        {
            return Path.Combine(RootPath, "Templates", $"Tab{groupId}");
        }

        public static string GetTemplateFilePath(int groupId, int projectId)
        {
            return Path.Combine(GetTemplateFolder(groupId), GetCanonicalFileName(projectId));
        }

        public static string GetInitialFolder(int groupId)
        {
            return Path.Combine(GetWorkingFolder(groupId), "Initial");
        }

        public static string GetInitialFilePath(int groupId, int projectId)
        {
            return Path.Combine(GetInitialFolder(groupId), GetCanonicalFileName(projectId));
        }

        public static string GetInitialFilePathForSource(int groupId, int projectId, string sourcePath)
        {
            string extension = Path.GetExtension(sourcePath);
            if (!extension.Equals(".doc", StringComparison.OrdinalIgnoreCase) &&
                !extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
                extension = projectId == 7 ? ".doc" : ".docx";
            return Path.Combine(GetInitialFolder(groupId), $"Project{projectId}{extension}");
        }

        public static string FindExistingWorkingFile(int groupId, int projectId)
        {
            return FindFirstExisting(GetWorkingFolder(groupId), GetCompatibleFileNames(projectId));
        }

        /// <summary>
        /// 正規テンプレートを返す。未配置時は旧リセット元の Initial、Initial\Initial
        /// 既存データを安全にコピーして正規配置を作る。
        /// </summary>
        public static string EnsureCanonicalTemplate(int groupId, int projectId)
        {
            string templatePath = FindFirstExisting(
                GetTemplateFolder(groupId), GetCompatibleFileNames(projectId));
            if (!string.IsNullOrEmpty(templatePath))
            {
                MakeWritable(templatePath);
                RemoveZoneIdentifier(templatePath);
                return templatePath;
            }

            string sourcePath = FindFirstExisting(
                GetInitialFolder(groupId), GetCompatibleFileNames(projectId));
            if (string.IsNullOrEmpty(sourcePath))
                sourcePath = FindFirstExisting(
                    Path.Combine(GetInitialFolder(groupId), "Initial"),
                    GetCompatibleFileNames(projectId));
            if (string.IsNullOrEmpty(sourcePath))
                throw new FileNotFoundException(
                    $"Project{projectId} のテンプレートまたは互換リセット元が見つかりません。",
                    GetTemplateFilePath(groupId, projectId));

            templatePath = Path.Combine(
                GetTemplateFolder(groupId),
                $"Project{projectId}{Path.GetExtension(sourcePath).ToLowerInvariant()}");
            Directory.CreateDirectory(Path.GetDirectoryName(templatePath));
            MakeWritable(sourcePath);
            RemoveZoneIdentifier(sourcePath);
            File.Copy(sourcePath, templatePath, overwrite: false);
            MakeWritable(templatePath);
            RemoveZoneIdentifier(templatePath);
            return templatePath;
        }

        /// <summary>作業ファイルが無ければ正規テンプレートから生成して返す。</summary>
        public static string EnsureWorkingFile(int groupId, int projectId)
        {
            string existing = FindExistingWorkingFile(groupId, projectId);
            if (!string.IsNullOrEmpty(existing))
                return existing;

            string sourcePath = EnsureCanonicalTemplate(groupId, projectId);
            string destinationPath = GetWorkingFilePathForSource(groupId, projectId, sourcePath);
            Directory.CreateDirectory(Path.GetDirectoryName(destinationPath));
            File.Copy(sourcePath, destinationPath, overwrite: false);
            MakeWritable(destinationPath);
            RemoveZoneIdentifier(destinationPath);
            return destinationPath;
        }

        public static string FindProblemJson(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return null;

            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string referencesPath = Path.Combine(baseDirectory, "References", "JSON", fileName);
            if (File.Exists(referencesPath))
                return referencesPath;

            string rootPath = Path.Combine(baseDirectory, fileName);
            return File.Exists(rootPath) ? rootPath : null;
        }

        public static void MakeWritable(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return;

            var fileInfo = new FileInfo(filePath);
            if (fileInfo.IsReadOnly)
                fileInfo.IsReadOnly = false;
        }

        public static void RemoveZoneIdentifier(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return;

            try
            {
                MakeWritable(filePath);
                DeleteFileW(filePath + ":Zone.Identifier");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[WordDataPathHelper] Zone.Identifier 解除スキップ ({filePath}): {ex.Message}");
            }
        }

        private static string LoadRootPath()
        {
            try
            {
                string configPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    var config = JsonConvert.DeserializeObject<WordAppConfig>(
                        File.ReadAllText(configPath));
                    if (!string.IsNullOrWhiteSpace(config?.WordDataPath))
                        return Path.GetFullPath(
                            Environment.ExpandEnvironmentVariables(config.WordDataPath.Trim()));
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[WordDataPathHelper] config.json 読み込み失敗: {ex.Message}");
            }

            return DefaultWordDataPath;
        }

        private static string GetCanonicalFileName(int projectId)
        {
            return $"Project{projectId}{(projectId == 7 ? ".doc" : ".docx")}";
        }

        private static IEnumerable<string> GetCompatibleFileNames(int projectId)
        {
            if (projectId == 7)
            {
                yield return $"Project{projectId}.doc";
                yield return $"project{projectId}.doc";
                yield break;
            }

            yield return $"Project{projectId}.docx";
            yield return $"Project{projectId}.doc";
            yield return $"project{projectId}.docx";
            yield return $"project{projectId}.doc";
        }

        private static string FindFirstExisting(string folder, IEnumerable<string> fileNames)
        {
            if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
                return null;

            foreach (string fileName in fileNames)
            {
                string path = Path.Combine(folder, fileName);
                if (File.Exists(path))
                    return path;
            }

            return null;
        }
    }
}
