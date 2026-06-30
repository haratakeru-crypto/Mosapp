using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Libraries;

namespace MOS_Word_app
{
    /// <summary>
    /// Word プロジェクトのリセット処理を共通化。MainWindow（すべてリセット）と AppBarWindow（単一リセット）の両方から使用する。
    /// </summary>
    public static class WordProjectResetHelper
    {
        private const string BasePath = @"C:\MOSTest\Word365";

        private static readonly string[] WordFileExtensions =
        {
            ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf", ".txt"
        };

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool DeleteFileW(string lpFileName);

        public static void ResetProject(int groupId, int projectId)
        {
            // 個別リセット時は対象プロジェクトの採点ログ行のみ削除（他プロジェクトのログを保持）
            LogReader.ClearTaskEvidenceForProject(projectId);
            LogReader.ClearDestructiveLogForProject(projectId);
            LogReader.ClearSnapshot();
            LogReader.ClearCurrentTaskFile();
            WordTaskAttemptRegistry.ClearProject(projectId);

            // 保存先（作業フォルダ）: Tab{groupId}\ 直下のみ。参照元: Tab{groupId}\Initial（Templates は使わない）
            string workingFolder = Path.Combine(BasePath, $"Tab{groupId}");
            string initialFolder = Path.Combine(BasePath, $"Tab{groupId}", "Initial");
            string initialInitialFolder = Path.Combine(BasePath, $"Tab{groupId}", "Initial", "Initial");

            // リセット参照フォルダ内の Word データから Zone.Identifier を削除（保護ビュー防止）
            UnblockWordFilesInFolder(workingFolder);
            UnblockWordFilesInFolder(initialFolder);
            UnblockWordFilesInFolder(initialInitialFolder);

            string[] possibleNames = (groupId == 1 && projectId == 7)
                ? new[] { $"Project{projectId}.doc", $"project{projectId}.doc" }
                : new[] { $"Project{projectId}.docx", $"Project{projectId}.doc", $"project{projectId}.docx", $"project{projectId}.doc" };

            // 参照元: Initial 直下、なければ Initial\Initial
            string sourceFilePath = null;
            foreach (var fileName in possibleNames)
            {
                string fullPath = Path.Combine(initialFolder, fileName);
                if (File.Exists(fullPath))
                {
                    sourceFilePath = fullPath;
                    break;
                }
            }
            if (string.IsNullOrEmpty(sourceFilePath) && Directory.Exists(initialInitialFolder))
            {
                foreach (var fileName in possibleNames)
                {
                    string fullPath = Path.Combine(initialInitialFolder, fileName);
                    if (File.Exists(fullPath))
                    {
                        sourceFilePath = fullPath;
                        break;
                    }
                }
            }
            if (string.IsNullOrEmpty(sourceFilePath))
                throw new FileNotFoundException($"リセット用ファイルが見つかりません: {initialFolder} に Project{projectId}.docx 等を配置してください。");

            RemoveZoneIdentifier(sourceFilePath);

            string fileExtension = Path.GetExtension(sourceFilePath);
            string projectFilePath = Path.Combine(workingFolder, $"Project{projectId}{fileExtension}");
            if (!Directory.Exists(workingFolder))
                Directory.CreateDirectory(workingFolder);

            if (File.Exists(projectFilePath))
            {
                try
                {
                    var projectFileInfo = new FileInfo(projectFilePath);
                    if (projectFileInfo.IsReadOnly)
                        projectFileInfo.IsReadOnly = false;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[WordProjectResetHelper] コピー先の読み取り専用解除エラー: {ex.Message}");
                }
            }

            int retryCount = 0;
            const int maxRetries = 5;
            while (retryCount < maxRetries)
            {
                try
                {
                    File.Copy(sourceFilePath, projectFilePath, overwrite: true);
                    // File.Copy は Zone.Identifier ADS も引き継ぐため、コピー直後に削除して保護ビューを防ぐ
                    RemoveZoneIdentifier(projectFilePath);
                    try
                    {
                        var destInfo = new FileInfo(projectFilePath);
                        if (destInfo.IsReadOnly)
                            destInfo.IsReadOnly = false;
                    }
                    catch { }

                    if (groupId == 1 && projectId == 7)
                        TryDeleteP7DerivativeOutputs(workingFolder);

                    return;
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
        }

        /// <summary>
        /// ファイルの Zone.Identifier ADS（NTFS 代替データストリーム）を削除する。
        /// インターネット由来マークを外し、Word の保護ビューで開かれるのを防ぐ。
        /// </summary>
        private static void RemoveZoneIdentifier(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return;

            try
            {
                var fi = new FileInfo(filePath);
                bool wasReadOnly = fi.IsReadOnly;
                if (wasReadOnly)
                {
                    try { fi.IsReadOnly = false; } catch { }
                }

                DeleteFileW(filePath + ":Zone.Identifier");

                if (wasReadOnly)
                {
                    try { new FileInfo(filePath).IsReadOnly = true; } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordProjectResetHelper] Zone.Identifier 削除スキップ ({filePath}): {ex.Message}");
            }
        }

        /// <summary>リセット参照フォルダ内の Word 関連ファイルから Zone.Identifier を一括削除する。</summary>
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
                System.Diagnostics.Debug.WriteLine($"[WordProjectResetHelper] フォルダ一括解除スキップ ({folderPath}): {ex.Message}");
            }
        }

        /// <summary>7-4/7-5 の派生ファイルが残ると 7-2 の状態のみで誤判定し得るため、P7 個別リセット時に削除する。</summary>
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
                    System.Diagnostics.Debug.WriteLine($"[WordProjectResetHelper] 派生ファイル削除スキップ ({name}): {ex.Message}");
                }
            }
        }
    }
}
