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
        private const string BasePath = @"C:\MOSTest\Word365";

        public static void ResetProject(int groupId, int projectId)
        {
            // 個別リセット時は対象プロジェクトの採点ログ行のみ削除（他プロジェクトのログを保持）
            LogReader.ClearTaskEvidenceForProject(projectId);

            // 保存先（作業フォルダ）: Tab{groupId}\ 直下のみ。参照元: Tab{groupId}\Initial（Templates は使わない）
            string workingFolder = Path.Combine(BasePath, $"Tab{groupId}");
            string initialFolder = Path.Combine(BasePath, $"Tab{groupId}", "Initial");
            string initialInitialFolder = Path.Combine(BasePath, $"Tab{groupId}", "Initial", "Initial");

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
                    try
                    {
                        var destInfo = new FileInfo(projectFilePath);
                        if (destInfo.IsReadOnly)
                            destInfo.IsReadOnly = false;
                    }
                    catch { }
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
    }
}
