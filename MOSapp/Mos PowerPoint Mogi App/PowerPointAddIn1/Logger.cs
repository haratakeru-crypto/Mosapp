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

        /// <summary>5-1: 配布資料3スライド・部単位4部で印刷したことを記録。</summary>
        public static void LogTask5_1Print()
        {
            LogTaskTag("Task5-1", "Print");
        }

        /// <summary>11-7: ノート・3部・部単位で印刷したことを記録。</summary>
        public static void LogTask11_7Print()
        {
            LogTaskTag("Task11-7", "Print");
        }

        /// <summary>8-4: オーディオ再生設定（フェードイン4秒等）を記録。</summary>
        public static void LogTask8_4Audio()
        {
            LogTaskTag("Task8-4", "Audio");
        }

        /// <summary>7-2: スライドの再利用を記録。</summary>
        public static void LogTask7_2ReuseSlides()
        {
            LogTaskTag("Task7-2", "ReuseSlides");
        }

        /// <summary>7-3: アウトラインから挿入を記録。</summary>
        public static void LogTask7_3InsertFromOutline()
        {
            LogTaskTag("Task7-3", "InsertFromOutline");
        }

        /// <summary>10-1: ドキュメント検査実行を記録。</summary>
        public static void LogTask10_1DocumentInspector()
        {
            LogTaskTag("Task10-1", "DocumentInspector");
        }

        /// <summary>10-7: レイアウト複製を記録。</summary>
        public static void LogTask10_7LayoutDuplicate()
        {
            LogTaskTag("Task10-7", "LayoutDuplicate");
        }

        /// <summary>
        /// 現在タスク開始を記録（試験アプリがタスクを切り替えたとき）。形式: [timestamp] [TaskStart] P-T
        /// </summary>
        public static void LogTaskStart(int projectId, int taskId)
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry = $"[{timestamp}] [TaskStart] {projectId}-{taskId}";
                    AppendLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        /// <summary>
        /// 汎用操作を記録。現在タスクが指定されていれば [Task P-T] をプレフィックスする。形式: [timestamp] [Task P-T] [Op] Type Detail
        /// </summary>
        public static void LogOperation(string operationType, string detail, int? projectId = null, int? taskId = null)
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string taskPrefix = (projectId.HasValue && taskId.HasValue) ? $"[Task {projectId.Value}-{taskId.Value}] " : "";
                    string logEntry = $"[{timestamp}] {taskPrefix}[Op] {operationType} {detail}";
                    AppendLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Logger] Error writing log: {ex.Message}");
            }
        }

        private static void LogTaskTag(string taskTag, string identifier)
        {
            try
            {
                lock (_lockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string logEntry = $"[{timestamp}] [{taskTag}] {identifier}";
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
