using System;
using System.IO;
using System.Text;

namespace ExcelAddIn1
{
    /// <summary>
    /// Excel 操作ログを %TEMP%\mos_excel_log.txt に記録する。
    /// </summary>
    public static class Logger
    {
        private static readonly object LockObject = new object();
        private static readonly string LogFilePath = Path.Combine(Path.GetTempPath(), "mos_excel_log.txt");
        private static int _currentProjectId = -1;
        private static int _currentTaskId = -1;
        private static int _currentAttemptNo = 1;

        public static void LogCommand(string commandId)
        {
            try
            {
                lock (LockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    AppendToFile($"[{timestamp}] [{commandId}] Executed");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelLogger] Error writing log: " + ex.Message);
            }
        }

        public static void SetCurrentTaskContext(int projectId, int taskId, int attemptNo)
        {
            lock (LockObject)
            {
                _currentProjectId = projectId;
                _currentTaskId = taskId;
                _currentAttemptNo = attemptNo < 1 ? 1 : attemptNo;
            }
        }

        /// <summary>現在のタスク文脈（<see cref="LogOperation"/> の接頭辞に使われる値）。</summary>
        public static void GetCurrentTaskContext(out int projectId, out int taskId, out int attemptNo)
        {
            lock (LockObject)
            {
                projectId = _currentProjectId;
                taskId = _currentTaskId;
                attemptNo = _currentAttemptNo;
            }
        }

        /// <summary>
        /// 一時的にタスク文脈を差し替えて <paramref name="action"/> を実行する（同一スレッド・再入可能）。
        /// Backstage 開閉など、ログ記録時だけ過去のタスク文脈を使う用途。
        /// </summary>
        public static void RunWithTaskContext(int projectId, int taskId, int attemptNo, Action action)
        {
            if (action == null) return;
            lock (LockObject)
            {
                int prevP = _currentProjectId;
                int prevT = _currentTaskId;
                int prevA = _currentAttemptNo;
                try
                {
                    if (projectId > 0 && taskId > 0)
                    {
                        _currentProjectId = projectId;
                        _currentTaskId = taskId;
                        _currentAttemptNo = attemptNo < 1 ? 1 : attemptNo;
                    }
                    action();
                }
                finally
                {
                    _currentProjectId = prevP;
                    _currentTaskId = prevT;
                    _currentAttemptNo = prevA;
                }
            }
        }

        public static void LogTaskStart(int projectId, int taskId, int attemptNo = 1)
        {
            try
            {
                lock (LockObject)
                {
                    if (attemptNo < 1) attemptNo = 1;
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    AppendToFile($"[{timestamp}] [TaskStart] {projectId}-{taskId}-{attemptNo}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelLogger] Error writing task start: " + ex.Message);
            }
        }

        public static void LogOperation(string operationType, string detail)
        {
            try
            {
                lock (LockObject)
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    string taskPrefix = (_currentProjectId > 0 && _currentTaskId > 0)
                        ? $"[Task {_currentProjectId}-{_currentTaskId}-{_currentAttemptNo}] "
                        : "";
                    AppendToFile($"[{timestamp}] {taskPrefix}[Op] {operationType} {detail}".TrimEnd());
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelLogger] Error writing operation log: " + ex.Message);
            }
        }

        public static void ClearLog()
        {
            try
            {
                lock (LockObject)
                {
                    if (File.Exists(LogFilePath))
                    {
                        File.Delete(LogFilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelLogger] Error clearing log: " + ex.Message);
            }
        }

        public static string GetLogFilePath()
        {
            return LogFilePath;
        }

        private static void AppendToFile(string line)
        {
            using (var stream = new FileStream(
                LogFilePath,
                FileMode.Append,
                FileAccess.Write,
                FileShare.ReadWrite | FileShare.Delete))
            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                writer.WriteLine(line);
            }
        }
    }
}
