using System;
using System.IO;
using System.Text;

namespace New_MOSWordVSTOAddIn
{
    /// <summary>
    /// Word 操作ログ。採点用は <see cref="LogTaskEvidence"/> の Project/Task 付き行のみ。
    /// <see cref="LogCommand"/> はデバッグ用（旧2括弧形式）。採点ロジックでは無視される。
    /// </summary>
    public static class Logger
    {
        private static readonly object _lockObject = new object();
        private static readonly string _logFileName = "mos_word_log.txt";
        private static string _logFilePath;

        private static string LogFilePath
        {
            get
            {
                if (string.IsNullOrEmpty(_logFilePath))
                    _logFilePath = Path.Combine(Path.GetTempPath(), _logFileName);
                return _logFilePath;
            }
        }

        /// <summary>デバッグ用。形式: [timestamp] [commandId] Executed（採点では使用しない）</summary>
        public static void LogCommand(string commandId)
        {
            if (string.IsNullOrWhiteSpace(commandId))
                return;

            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry = $"[{timestamp}] [{commandId}] Executed";

                    using (var fileStream = new FileStream(
                        LogFilePath,
                        FileMode.Append,
                        FileAccess.Write,
                        FileShare.ReadWrite | FileShare.Delete))
                    using (var writer = new StreamWriter(fileStream, Encoding.UTF8))
                    {
                        writer.WriteLine(logEntry);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        public static void ClearLog()
        {
            try
            {
                lock (_lockObject)
                {
                    if (File.Exists(LogFilePath))
                        File.Delete(LogFilePath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error clearing log: {ex.Message}");
            }
        }

        public static string GetLogFilePath()
        {
            return LogFilePath;
        }

        /// <summary>
        /// 採点用ログ行を <c>mos_word_log.txt</c> に追記する。
        /// 形式: [timestamp] [ProjectN] [TaskN-M] [commandId] Executed（Task の N は Project と一致すること）
        /// </summary>
        public static void LogTaskEvidence(int projectId, int taskId, string commandId)
        {
            if (string.IsNullOrWhiteSpace(commandId))
                return;

            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry =
                        $"[{timestamp}] [Project{projectId}] [Task{projectId}-{taskId}] [{commandId}] Executed";

                    using (var fileStream = new FileStream(
                        LogFilePath,
                        FileMode.Append,
                        FileAccess.Write,
                        FileShare.ReadWrite | FileShare.Delete))
                    using (var writer = new StreamWriter(fileStream, Encoding.UTF8))
                    {
                        writer.WriteLine(logEntry);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing scoring log: {ex.Message}");
            }
        }
    }
}
