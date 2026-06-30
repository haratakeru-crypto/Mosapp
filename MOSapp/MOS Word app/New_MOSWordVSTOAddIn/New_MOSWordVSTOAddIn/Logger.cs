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
        private static int _currentProjectId = -1;
        private static int _currentTaskId = -1;
        private static int _currentAttemptNo;

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

        /// <summary>Word 内で VSTO が稼働中であることを %TEMP% に記録する。</summary>
        public static void WriteVstoHeartbeat()
        {
            try
            {
                lock (_lockObject)
                {
                    string path = Path.Combine(Path.GetTempPath(), "mos_word_vsto_heartbeat.txt");
                    File.WriteAllText(path, DateTimeOffset.UtcNow.ToUnixTimeMilliseconds().ToString(), Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] WriteVstoHeartbeat: {ex.Message}");
            }
        }

        public static void SetCurrentTaskContext(int projectId, int taskId, int attemptNo)
        {
            lock (_lockObject)
            {
                _currentProjectId = projectId;
                _currentTaskId = taskId;
                _currentAttemptNo = attemptNo < 0 ? 0 : attemptNo;
            }
        }

        /// <summary>
        /// 汎用操作ログ。[Task P-T-A] [Op] Type Detail（TaskStart は試験アプリが記録）。
        /// </summary>
        public static void LogOperation(string operationType, string detail)
        {
            if (string.IsNullOrWhiteSpace(operationType))
                return;
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string taskPrefix = (_currentProjectId > 0 && _currentTaskId > 0)
                        ? $"[Task {_currentProjectId}-{_currentTaskId}-{_currentAttemptNo}] "
                        : "";
                    string logEntry = $"[{timestamp}] {taskPrefix}[Op] {operationType} {detail}".TrimEnd();
                    AppendLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] LogOperation: {ex.Message}");
            }
        }

        private static void AppendLine(string logEntry)
        {
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
