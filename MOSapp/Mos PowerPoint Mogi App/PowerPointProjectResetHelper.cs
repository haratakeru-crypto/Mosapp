using System;
using System.Configuration;
using System.IO;
using System.Threading;

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
            string basePath = ConfigurationManager.AppSettings["PowerPointDataPath"]
                ?? @"C:\MOSTest\PowerPoint365";
            string tabFolder = Path.Combine(basePath, $"Tab{groupId}");

            string[] possibleNames = { $"Project{projectId}.pptx", $"Project{projectId}.ppt" };
            string projectFilePath = null;

            foreach (var fileName in possibleNames)
            {
                string fullPath = Path.Combine(tabFolder, fileName);
                if (File.Exists(fullPath))
                {
                    projectFilePath = fullPath;
                    break;
                }
            }

            if (string.IsNullOrEmpty(projectFilePath))
            {
                projectFilePath = Path.Combine(tabFolder, $"Project{projectId}.pptx");
                System.Diagnostics.Debug.WriteLine($"[PowerPointProjectResetHelper] ファイルが見つかりません。作成します: {projectFilePath}");
            }

            string templatesFolder = Path.Combine(basePath, "Templates", $"Tab{groupId}");
            string templatePath = null;
            string fileExtension = ".pptx";

            if (Directory.Exists(templatesFolder))
            {
                var files = Directory.GetFiles(templatesFolder, $"*{fileExtension}", SearchOption.TopDirectoryOnly);
                string searchPattern = $"project{projectId}".ToLower();
                foreach (var file in files)
                {
                    string fn = Path.GetFileNameWithoutExtension(file).ToLower();
                    if (fn.Contains(searchPattern) || fn == searchPattern)
                    {
                        templatePath = file;
                        break;
                    }
                }
            }

            if (string.IsNullOrEmpty(templatePath))
            {
                string[] patterns = {
                    Path.Combine(basePath, "Templates", $"Tab{groupId}", $"project{projectId}{fileExtension}"),
                    Path.Combine(basePath, "Templates", $"project{projectId}{fileExtension}"),
                    Path.Combine(basePath, "Templates", $"Tab{groupId}", $"Tab{groupId}_project{projectId}{fileExtension}")
                };
                foreach (var pattern in patterns)
                {
                    if (File.Exists(pattern))
                    {
                        templatePath = pattern;
                        break;
                    }
                }
            }

            if (string.IsNullOrEmpty(templatePath))
                throw new FileNotFoundException($"テンプレートファイルが見つかりません: {templatesFolder}");

            var templateFileInfo = new FileInfo(templatePath);
            if (!templateFileInfo.IsReadOnly)
                templateFileInfo.IsReadOnly = true;

            if (File.Exists(projectFilePath))
            {
                var projectFileInfo = new FileInfo(projectFilePath);
                if (projectFileInfo.IsReadOnly)
                    projectFileInfo.IsReadOnly = false;
            }

            File.Copy(templatePath, projectFilePath, overwrite: true);
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

            string initialFolderPath = Path.Combine(tabFolder, "Initial");
            string initialFilePath = Path.Combine(initialFolderPath, $"project{projectId}{fileExtension}");

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
    }
}
