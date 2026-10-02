using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Libraries;

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// プロジェクトリセット処理を共通化する静的ヘルパー。
    /// アプリバーからの単体リセットと、プロジェクト一覧からの「すべてをリセットする」の両方で使用する。
    /// コピー先の読み取り専用は明示的に解除する。
    /// </summary>
    public static class PowerPointProjectResetHelper
    {
        public static void ResetProject(int groupId, int projectId)
        {
            PPLogReader.ClearLog();
            // プロジェクト1/5/10/11をやり直すときは該当タスクの証跡のみ削除。他プロジェクトの証跡は残す。
            PPLogReader.ClearTaskEvidenceForProject(projectId);
            PPLogReader.ClearCurrentTaskFile();
            PPLogReader.ClearDestructiveLog();
            PPLogReader.ClearSnapshot();
            PPTaskAttemptRegistry.ClearProject(projectId);

            string projectFilePath = PowerPointDataPathHelper.GetWorkingProjectPath(groupId, projectId);
            string templatePath = PowerPointDataPathHelper.GetTemplateProjectPath(groupId, projectId);
            string initialFilePath = PowerPointDataPathHelper.GetInitialProjectPath(groupId, projectId);

            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"テンプレートファイルが見つかりません: {templatePath}");

            string projectFolder = Path.GetDirectoryName(projectFilePath);
            if (!Directory.Exists(projectFolder))
                Directory.CreateDirectory(projectFolder);

            // テンプレートは編集可能な状態を正とし、既存の読み取り専用属性も明示解除する。
            PowerPointDataPathHelper.ClearReadOnly(templatePath, "テンプレート");
            RemoveZoneIdentifier(templatePath);

            if (File.Exists(projectFilePath))
            {
                var projectFileInfo = new FileInfo(projectFilePath);
                if (projectFileInfo.IsReadOnly)
                    projectFileInfo.IsReadOnly = false;
            }

            File.Copy(templatePath, projectFilePath, overwrite: true);
            // File.Copy は Zone.Identifier も引き継ぐため、コピー先だけ外して保護ビューを防ぐ
            RemoveZoneIdentifier(projectFilePath);
            // コピー直後に読み取り専用を解除（テンプレート属性の引き継ぎを防ぐ）
            try
            {
                var projectFileInfo = new FileInfo(projectFilePath);
                if (projectFileInfo.IsReadOnly)
                    projectFileInfo.IsReadOnly = false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PowerPointProjectResetHelper] 読み取り専用解除（プロジェクト）: {ex.Message}");
            }

            string initialFolderPath = Path.GetDirectoryName(initialFilePath);

            if (!Directory.Exists(initialFolderPath))
                Directory.CreateDirectory(initialFolderPath);

            if (File.Exists(initialFilePath))
            {
                try
                {
                    var initialFileInfo = new FileInfo(initialFilePath);
                    if (initialFileInfo.IsReadOnly)
                        initialFileInfo.IsReadOnly = false;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[PowerPointProjectResetHelper] Initial既存ファイルの読み取り専用解除: {ex.Message}");
                }
            }

            int retryCount = 0;
            const int maxRetries = 5;
            bool copied = false;
            while (!copied && retryCount < maxRetries)
            {
                try
                {
                    File.Copy(projectFilePath, initialFilePath, overwrite: true);
                    RemoveZoneIdentifier(initialFilePath);
                    copied = true;
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

            if (copied)
            {
                try
                {
                    var initialFileInfo = new FileInfo(initialFilePath);
                    if (initialFileInfo.IsReadOnly)
                        initialFileInfo.IsReadOnly = false;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[PowerPointProjectResetHelper] 読み取り専用解除（Initial）: {ex.Message}");
                }
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        static extern bool DeleteFileW(string lpFileName);

        /// <summary>このファイルだけから Zone.Identifier を外す。フォルダ全体はたどらない。</summary>
        static void RemoveZoneIdentifier(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return;

            try
            {
                var fileInfo = new FileInfo(filePath);
                if (fileInfo.IsReadOnly)
                {
                    try { fileInfo.IsReadOnly = false; } catch { }
                }

                DeleteFileW(filePath + ":Zone.Identifier");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[PowerPointProjectResetHelper] Zone.Identifier 削除スキップ: {ex.Message}");
            }
        }
    }
}
