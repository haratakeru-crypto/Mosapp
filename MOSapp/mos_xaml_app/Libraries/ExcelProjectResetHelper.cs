using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using MOSExcelMogiApp.Infrastructure;

namespace Libraries
{
    /// <summary>
    /// 教材 / 類題の単体リセットで共有するファイル復元とログクリア。
    /// </summary>
    public static class ExcelProjectResetHelper
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool DeleteFileW(string lpFileName);

        public static void RemoveZoneIdentifier(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return;
            try
            {
                FileInfo fi = new FileInfo(filePath);
                if (fi.IsReadOnly)
                {
                    try { fi.IsReadOnly = false; } catch { }
                }

                DeleteFileW(filePath + ":Zone.Identifier");
            }
            catch { }
        }

        public static bool IsFileLocked(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static int WaitForUnlock(string filePath, int maxRetries = 10, int delayMs = 500)
        {
            int retryCount = 0;
            while (retryCount < maxRetries && IsFileLocked(filePath))
            {
                Thread.Sleep(delayMs);
                retryCount++;
            }
            return retryCount;
        }

        public static int RestoreFromTemplate(string templatePath, string workingPath, string initialPath = null)
        {
            if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
                throw new FileNotFoundException("テンプレートファイルが見つかりません。", templatePath);
            if (string.IsNullOrEmpty(workingPath))
                throw new ArgumentException("作業ファイルのパスが空です。", nameof(workingPath));

            DataPathHelper.ClearReadOnly(templatePath);
            RemoveZoneIdentifier(templatePath);
            int retryCount = WaitForUnlock(workingPath);
            CopyOverwrite(templatePath, workingPath);
            if (!string.IsNullOrEmpty(initialPath))
                CopyOverwrite(workingPath, initialPath);
            return retryCount;
        }

        public static void ClearLogsAfterSuccessfulCopy(int projectId, int variantSetNo)
        {
            ExcelLogReader.ClearLogsForProjectScope(projectId, variantSetNo);
        }

        private static void CopyOverwrite(string sourcePath, string destinationPath)
        {
            string directory = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            if (File.Exists(destinationPath))
            {
                FileInfo destInfo = new FileInfo(destinationPath);
                if (destInfo.IsReadOnly)
                    destInfo.IsReadOnly = false;
            }

            File.Copy(sourcePath, destinationPath, overwrite: true);
            RemoveZoneIdentifier(destinationPath);

            FileInfo copied = new FileInfo(destinationPath);
            if (copied.IsReadOnly)
                copied.IsReadOnly = false;

            Debug.WriteLine("[ExcelProjectResetHelper] Copied " + sourcePath + " -> " + destinationPath);
        }
    }
}
