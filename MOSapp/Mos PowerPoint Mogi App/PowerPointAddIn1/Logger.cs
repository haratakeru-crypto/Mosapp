using System;
using System.IO;
using System.Text;

namespace PowerPointAddIn1
{
    /// <summary>
    /// PowerPoint 操作ログを記録するクラス。
    /// 10-4 グレースケール等の VSTO 検証用ログを mos_ppt_log.txt に記録する。
    /// </summary>
    public static class Logger
    {
        private static readonly object _lockObject = new object();
        private static readonly string _logFileName = "mos_ppt_log.txt";
        private static string _logFilePath;

        private static string LogFilePath
        {
            get
            {
                if (string.IsNullOrEmpty(_logFilePath))
                {
                    _logFilePath = Path.Combine(Path.GetTempPath(), _logFileName);
                }
                return _logFilePath;
            }
        }

        /// <summary>
        /// コマンド実行をログに記録
        /// </summary>
        public static void LogCommand(string commandId)
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry = $"[{timestamp}] [{commandId}] Executed";
                    AppendLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        /// <summary>
        /// 10-4 グレースケール操作をログに記録（採点で参照する形式）
        /// </summary>
        public static void LogTask10_4Grayscale()
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry = $"[{timestamp}] [Task10-4] Grayscale";
                    AppendLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
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
        /// ログファイルをクリア
        /// </summary>
        public static void ClearLog()
        {
            try
            {
                lock (_lockObject)
                {
                    if (File.Exists(LogFilePath))
                    {
                        File.Delete(LogFilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error clearing log: {ex.Message}");
            }
        }

        /// <summary>
        /// ログファイルのパスを取得（採点側・デバックタブ用）
        /// </summary>
        public static string GetLogFilePath()
        {
            return LogFilePath;
        }
    }
}
