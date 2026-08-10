using System;
using System.IO;
using System.Threading;
using Libraries;

namespace MOS_Word_app
{
    /// <summary>
    /// Word プロジェクトのリセット処理を共通化。MainWindow（すべてリセット）と AppBarWindow（単一リセット）の両方から使用する。
    /// </summary>
    public static class WordProjectResetHelper
    {
        private static readonly string[] WordFileExtensions =
        {
            ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf", ".txt"
        };

        public static void ResetProject(int groupId, int projectId)
        {
            // 個別リセット時は対象プロジェクトの採点ログ行のみ削除（他プロジェクトのログを保持）
            LogReader.ClearTaskEvidenceForProject(projectId);
            LogReader.ClearDestructiveLogForProject(projectId);
            LogReader.ClearSnapshot();
            LogReader.ClearCurrentTaskFile();
            WordTaskAttemptRegistry.ClearProject(projectId);

            string workingFolder = WordDataPathHelper.GetWorkingFolder(groupId);
            string templateFolder = WordDataPathHelper.GetTemplateFolder(groupId);
            string initialFolder = WordDataPathHelper.GetInitialFolder(groupId);
            string initialInitialFolder = Path.Combine(initialFolder, "Initial");

            // リセット参照フォルダ内の Word データから Zone.Identifier を削除（保護ビュー防止）
            UnblockWordFilesInFolder(workingFolder);
            UnblockWordFilesInFolder(templateFolder);
            UnblockWordFilesInFolder(initialFolder);
            UnblockWordFilesInFolder(initialInitialFolder);

            // 正規テンプレートを最優先。未配置時は旧 Initial 配置から
            // 既存ファイルをコピーし、正規テンプレートを作成する。
            string sourceFilePath = WordDataPathHelper.EnsureCanonicalTemplate(groupId, projectId);
            WordDataPathHelper.MakeWritable(sourceFilePath);
            WordDataPathHelper.RemoveZoneIdentifier(sourceFilePath);

            string projectFilePath = WordDataPathHelper.GetWorkingFilePathForSource(
                groupId, projectId, sourceFilePath);
            if (!Directory.Exists(workingFolder))
                Directory.CreateDirectory(workingFolder);

            if (File.Exists(projectFilePath))
            {
                try
                {
                    WordDataPathHelper.MakeWritable(projectFilePath);
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
                    WordDataPathHelper.RemoveZoneIdentifier(projectFilePath);
                    try
                    {
                        WordDataPathHelper.MakeWritable(projectFilePath);
                    }
                    catch { }

                    // 3アプリ共通規約: リセット後の作業ファイルを採点用 Initial に反映する。
                    string initialFilePath = WordDataPathHelper.GetInitialFilePathForSource(
                        groupId, projectId, sourceFilePath);
                    string initialDirectory = Path.GetDirectoryName(initialFilePath);
                    if (!Directory.Exists(initialDirectory))
                        Directory.CreateDirectory(initialDirectory);
                    WordDataPathHelper.MakeWritable(initialFilePath);
                    File.Copy(projectFilePath, initialFilePath, overwrite: true);
                    WordDataPathHelper.MakeWritable(initialFilePath);
                    WordDataPathHelper.RemoveZoneIdentifier(initialFilePath);

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
            WordDataPathHelper.RemoveZoneIdentifier(filePath);
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
