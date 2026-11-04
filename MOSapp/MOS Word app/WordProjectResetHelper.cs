using System;
using System.IO;
using System.Threading;
using Libraries;

namespace MOS_Word_app
{
    /// <summary>
    /// Word プロジェクトのリセット処理を共通化。演習（Tab1\ProjectN）と類題（PracticeVariant\similar）を分ける。
    /// </summary>
    public static class WordProjectResetHelper
    {
        public const int DefaultVariantSetNo = 1;

        private static readonly string[] WordFileExtensions =
        {
            ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf", ".txt"
        };

        /// <summary>後方互換。演習プロジェクトのリセット。</summary>
        public static void ResetProject(int groupId, int projectId)
        {
            ResetTextbookProject(groupId, projectId);
        }

        public static void ResetTextbookProject(int groupId, int projectId)
        {
            ClearLogsAfterSuccessfulCopy(projectId, 0);

            string workingFolder = WordDataPathHelper.GetWorkingFolder(groupId);
            string templateFolder = WordDataPathHelper.GetTemplateFolder(groupId);
            string initialFolder = WordDataPathHelper.GetInitialFolder(groupId);
            string initialInitialFolder = Path.Combine(initialFolder, "Initial");

            UnblockWordFilesInFolder(workingFolder);
            UnblockWordFilesInFolder(templateFolder);
            UnblockWordFilesInFolder(initialFolder);
            UnblockWordFilesInFolder(initialInitialFolder);

            string sourceFilePath = WordDataPathHelper.EnsureCanonicalTemplate(groupId, projectId);
            WordDataPathHelper.MakeWritable(sourceFilePath);
            WordDataPathHelper.RemoveZoneIdentifier(sourceFilePath);

            string projectFilePath = WordDataPathHelper.GetWorkingFilePathForSource(
                groupId, projectId, sourceFilePath);
            if (!Directory.Exists(workingFolder))
                Directory.CreateDirectory(workingFolder);

            string initialFilePath = WordDataPathHelper.GetInitialFilePathForSource(
                groupId, projectId, sourceFilePath);

            RestoreFromTemplate(sourceFilePath, projectFilePath, initialFilePath);

            if (groupId == 1 && projectId == 7)
                TryDeleteP7DerivativeOutputs(workingFolder);
        }

        public static void ResetVariantProject(int groupId, int projectId, int variantSetNo = DefaultVariantSetNo)
        {
            if (variantSetNo < 1)
                variantSetNo = DefaultVariantSetNo;

            ClearLogsAfterSuccessfulCopy(projectId, variantSetNo);

            string templatePath = WordDataPathHelper.FindVariantTemplateFile(groupId, projectId, variantSetNo);
            if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
            {
                throw new FileNotFoundException(
                    $"類題{variantSetNo} プロジェクト{projectId} のテンプレートが見つかりません。",
                    templatePath ?? string.Empty);
            }

            string variantFolder = WordDataPathHelper.GetVariantFolder(groupId, variantSetNo);
            string templatesFolder = Path.Combine(variantFolder, "Templates");
            UnblockWordFilesInFolder(variantFolder);
            UnblockWordFilesInFolder(templatesFolder);

            WordDataPathHelper.MakeWritable(templatePath);
            WordDataPathHelper.RemoveZoneIdentifier(templatePath);

            string workingPath = WordDataPathHelper.ResolveVariantWorkingFilePath(
                groupId, projectId, variantSetNo, templatePath);
            string workingDirectory = Path.GetDirectoryName(workingPath);
            if (!string.IsNullOrEmpty(workingDirectory) && !Directory.Exists(workingDirectory))
                Directory.CreateDirectory(workingDirectory);

            RestoreFromTemplate(templatePath, workingPath, initialPath: null);
        }

        public static void ClearLogsAfterSuccessfulCopy(int projectId, int variantSetNo)
        {
            LogReader.ClearLogsForProjectScope(projectId, variantSetNo);
        }

        /// <summary>演習作業ファイルと Initial をテンプレートから復元（ログクリアは呼び出し側）。</summary>
        public static void RestoreTextbookProjectFiles(int groupId, int projectId)
        {
            string workingFolder = WordDataPathHelper.GetWorkingFolder(groupId);
            string sourceFilePath = WordDataPathHelper.EnsureCanonicalTemplate(groupId, projectId);
            WordDataPathHelper.MakeWritable(sourceFilePath);
            WordDataPathHelper.RemoveZoneIdentifier(sourceFilePath);

            string projectFilePath = WordDataPathHelper.GetWorkingFilePathForSource(
                groupId, projectId, sourceFilePath);
            if (!Directory.Exists(workingFolder))
                Directory.CreateDirectory(workingFolder);

            string initialFilePath = WordDataPathHelper.GetInitialFilePathForSource(
                groupId, projectId, sourceFilePath);

            RestoreFromTemplate(sourceFilePath, projectFilePath, initialFilePath);

            if (groupId == 1 && projectId == 7)
                TryDeleteP7DerivativeOutputs(workingFolder);
        }

        /// <summary>類題作業ファイルを Templates から復元（Initial は更新しない）。</summary>
        public static void RestoreVariantProjectFiles(int groupId, int projectId, int variantSetNo)
        {
            if (variantSetNo < 1)
                variantSetNo = DefaultVariantSetNo;

            string templatePath = WordDataPathHelper.FindVariantTemplateFile(groupId, projectId, variantSetNo);
            if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
            {
                throw new FileNotFoundException(
                    $"類題{variantSetNo} プロジェクト{projectId} のテンプレートが見つかりません。",
                    templatePath ?? string.Empty);
            }

            WordDataPathHelper.MakeWritable(templatePath);
            WordDataPathHelper.RemoveZoneIdentifier(templatePath);

            string workingPath = WordDataPathHelper.ResolveVariantWorkingFilePath(
                groupId, projectId, variantSetNo, templatePath);
            string workingDirectory = Path.GetDirectoryName(workingPath);
            if (!string.IsNullOrEmpty(workingDirectory) && !Directory.Exists(workingDirectory))
                Directory.CreateDirectory(workingDirectory);

            RestoreFromTemplate(templatePath, workingPath, initialPath: null);
        }

        public static void UnblockResetReferenceFolders(int groupId, int variantSetNo, bool includeTextbook, bool includeVariant)
        {
            if (includeTextbook)
            {
                UnblockWordFilesInFolder(WordDataPathHelper.GetWorkingFolder(groupId));
                UnblockWordFilesInFolder(WordDataPathHelper.GetTemplateFolder(groupId));
                UnblockWordFilesInFolder(WordDataPathHelper.GetInitialFolder(groupId));
                UnblockWordFilesInFolder(Path.Combine(WordDataPathHelper.GetInitialFolder(groupId), "Initial"));
            }

            if (includeVariant)
            {
                string variantFolder = WordDataPathHelper.GetVariantFolder(groupId, variantSetNo);
                UnblockWordFilesInFolder(variantFolder);
                UnblockWordFilesInFolder(Path.Combine(variantFolder, "Templates"));
            }
        }

        /// <summary>テンプレートから作業ファイルへコピー。initialPath 指定時は演習 Initial も更新する。</summary>
        public static int RestoreFromTemplate(string templatePath, string workingPath, string initialPath = null)
        {
            if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
                throw new FileNotFoundException("テンプレートファイルが見つかりません。", templatePath);
            if (string.IsNullOrEmpty(workingPath))
                throw new ArgumentException("作業ファイルのパスが空です。", nameof(workingPath));

            WordDataPathHelper.MakeWritable(templatePath);
            WordDataPathHelper.RemoveZoneIdentifier(templatePath);

            if (File.Exists(workingPath))
            {
                try
                {
                    WordDataPathHelper.MakeWritable(workingPath);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordProjectResetHelper] コピー先の読み取り専用解除エラー: {ex.Message}");
                }
            }

            int retryCount = 0;
            const int maxRetries = 5;
            while (retryCount < maxRetries)
            {
                try
                {
                    File.Copy(templatePath, workingPath, overwrite: true);
                    WordDataPathHelper.RemoveZoneIdentifier(workingPath);
                    try
                    {
                        WordDataPathHelper.MakeWritable(workingPath);
                    }
                    catch { }

                    if (!string.IsNullOrEmpty(initialPath))
                    {
                        string initialDirectory = Path.GetDirectoryName(initialPath);
                        if (!string.IsNullOrEmpty(initialDirectory) && !Directory.Exists(initialDirectory))
                            Directory.CreateDirectory(initialDirectory);
                        WordDataPathHelper.MakeWritable(initialPath);
                        File.Copy(workingPath, initialPath, overwrite: true);
                        WordDataPathHelper.MakeWritable(initialPath);
                        WordDataPathHelper.RemoveZoneIdentifier(initialPath);
                    }

                    return retryCount;
                }
                catch (IOException) when (retryCount < maxRetries - 1)
                {
                    retryCount++;
                    Thread.Sleep(200);
                }
                catch (UnauthorizedAccessException) when (retryCount < maxRetries - 1)
                {
                    retryCount++;
                    Thread.Sleep(200);
                }
            }

            throw new IOException($"ファイルをリセットできませんでした: {workingPath}");
        }

        private static void RemoveZoneIdentifier(string filePath)
        {
            WordDataPathHelper.RemoveZoneIdentifier(filePath);
        }

        private static void UnblockWordFilesInFolder(string folderPath)
        {
            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
                return;

            try
            {
                foreach (string file in Directory.EnumerateFiles(folderPath, "*.*", SearchOption.AllDirectories))
                {
                    string ext = Path.GetExtension(file);
                    if (string.IsNullOrEmpty(ext))
                        continue;

                    foreach (string allowed in WordFileExtensions)
                    {
                        if (ext.Equals(allowed, StringComparison.OrdinalIgnoreCase))
                        {
                            RemoveZoneIdentifier(file);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WordProjectResetHelper] フォルダ一括解除スキップ ({folderPath}): {ex.Message}");
            }
        }

        private static void TryDeleteP7DerivativeOutputs(string workingFolder)
        {
            if (string.IsNullOrEmpty(workingFolder) || !Directory.Exists(workingFolder))
                return;

            foreach (string name in new[] { "朗読会.txt", "朗読会.docm" })
            {
                string path = Path.Combine(workingFolder, name);
                if (!File.Exists(path)) continue;
                try { File.Delete(path); }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordProjectResetHelper] 派生ファイル削除スキップ ({name}): {ex.Message}");
                }
            }
        }
    }
}
